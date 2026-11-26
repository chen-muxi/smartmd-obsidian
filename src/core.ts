import { localizeMessage } from "./localization";
import { App, Notice, Platform, Plugin, requestUrl } from "obsidian";
import { MicrosoftDeviceCodeSession, MicrosoftDeviceToken } from "./microsoft-device-flow";
import { getAndroidVaultRoot } from "./onedrive-android-vault";

const MICROSOFT_CLIENT_ID = "0ad3b293-4c9e-4d7a-94a3-7043c8bb6248";
const MICROSOFT_REDIRECT_URI = "obsidian://smartkeep-connect-onedrive";
const MICROSOFT_SCOPE = "offline_access User.Read Files.ReadWrite.AppFolder";
import { DEFAULT_SETTINGS, SharedSmartMDState, SmartMDSettings, hasUnfinishedOneDriveSetup, mergeSettings, newId, sharedStateForSync } from "./types";

const MICROSOFT_SECRET = "obsidian-suite-microsoft-token";
const AI_SECRET = "obsidian-suite-ai-key";
const AI_HTTP_SECRET = "obsidian-suite-ai-http-key";
const LEGACY_ONEDRIVE_RESTART_MESSAGE = "正常重启已保留账号、知识库和历史文件身份；旧同步执行许可失效。必须重新只读核验当前范围并明确批准，才可进行任何上传、下载或删除。";

export function getDevicePlatformName(): string {
    if (Platform.isIosApp) return "iOS";
    if (Platform.isAndroidApp) return "Android";
    if (Platform.isWin) return "Windows";
    if (Platform.isMacOS) return "macOS";
    if (Platform.isLinux) return "Linux";
    return "Obsidian device";
}

function httpModelOrigin(address: string): string | null {
    if (!/^http:\/\//i.test(address.trim())) return null;
    try {
        const endpoint = new URL(address);
        return endpoint.protocol === "http:" && !endpoint.username && !endpoint.password ? endpoint.origin : null;
    } catch {
        return null;
    }
}

interface MicrosoftToken {
    accessToken: string;
    refreshToken: string;
    expiresAt: number;
}

export interface MicrosoftDeviceCodeChallenge {
    verificationUri: string;
    userCode: string;
    expiresIn: number;
}

export class SmartMDStore {
    settings: SmartMDSettings = mergeSettings(undefined);
    private listeners = new Set<() => void>();
    private secretCache = new Map<string, string>();
    private compatibilitySecrets = new Set<string>();
    private compatibilityNoticeShown = false;

    private approvedVaultRoot = "";
    private scopedInstallationKey = "";
    private loadedInThisProcess = false;
    private preserveInactiveOneDrive = false;
    private persistenceQueue: Promise<void> = Promise.resolve();

    constructor(private app: App, private plugin: Plugin) {}

    private currentPhysicalVaultRoot(): string {
        const adapter = this.app.vault.adapter;
        if (Platform?.isAndroidApp) return getAndroidVaultRoot(adapter);
        const desktop = adapter as Partial<{ getBasePath(): string; basePath: string; fullPath: string }>;
        const raw = typeof desktop?.getBasePath === "function"
            ? desktop.getBasePath()
            : typeof desktop?.basePath === "string"
                ? desktop.basePath
                : typeof desktop?.fullPath === "string" ? desktop.fullPath : "";
        const normalized = raw.replace(/\\/g, "/").replace(/\/+$/g, "");
        // Never fall back to the user-controlled Vault name. The physical
        // location is part of the authorization scope on Windows.
        return /^(?:[a-zA-Z]:\/|\/\/[^/]+\/[^/]+(?:\/|$))/.test(normalized)
            ? normalized : "";
    }

    async load(): Promise<void> {
        // An in-process data.json reread is not a new plugin lifecycle.
        // It must neither revoke an active approved ticket nor resurrect a
        // persisted grant that was already revoked in this JS process.
        const previousLocal = this.settings.local;
        const previousOneDrive = previousLocal.oneDrive;
        const activeLiveGrant = this.loadedInThisProcess
            && previousOneDrive.runAllowed === true
            && previousOneDrive.confirmed === true;
        const previousScope = [
            previousLocal.microsoftAccountId, previousOneDrive.vaultId,
            previousOneDrive.deviceId, previousOneDrive.installationProofId,
            previousOneDrive.confirmationVaultId, previousOneDrive.confirmationAccountId,
        ];
        let raw = (await this.plugin.loadData()) || {};
        if (!raw.suiteV2) {
            try {
                const adapter = this.app.vault.adapter;
                const legacyPath = ".obsidian/plugins/smartkeep-connect/data.json";
                if (await adapter.exists(legacyPath)) {
                    const legacyText = await adapter.read(legacyPath);
                    const legacyJson = JSON.parse(legacyText);
                    if (legacyJson?.suiteV2) {
                        raw = legacyJson;
                    }
                }
            } catch {}
        }
        this.settings = mergeSettings(raw.suiteV2);
        this.preserveInactiveOneDrive = !Platform.isDesktopApp && !this.currentPhysicalVaultRoot();
        if (this.preserveInactiveOneDrive) {
            this.approvedVaultRoot = "";
            this.scopedInstallationKey = "";
            this.settings.local.oneDrive.runAllowed = false;
            this.settings.local.oneDrive.confirmed = false;
            this.settings.local.oneDrive.enabled = false;
            this.loadedInThisProcess = true;
            return;
        }
        const oneDrive = this.settings.local.oneDrive;
        const hasExistingScope = Boolean(oneDrive.vaultId && oneDrive.deviceId);
        if (!oneDrive.vaultId) {
            oneDrive.vaultId = newId("vault");
        }
        if (!oneDrive.deviceId) {
            oneDrive.deviceId = newId("device");
        }
        if (!oneDrive.deviceName) {
            oneDrive.deviceName = getDevicePlatformName();
        }
        // An Obsidian installation may open multiple Vaults on the SAME PC.
        // Scope this local-only approval to a verified Windows Vault root;
        // the old v1 global installation key must never be treated as consent.
        this.approvedVaultRoot = "";
        this.scopedInstallationKey = "";
        let localProof = "";
        try {
            const physicalRoot = this.currentPhysicalVaultRoot();
            if (physicalRoot && globalThis.crypto?.subtle) {
                const digest = await globalThis.crypto.subtle.digest(
                    "SHA-256", new TextEncoder().encode(physicalRoot),
                );
                const keyHash = Array.from(new Uint8Array(digest))
                    .map((part) => part.toString(16).padStart(2, "0")).join("");
                const key = "smartmd:onedrive:installation-id:v2:" + keyHash;
                let stored = window.localStorage.getItem(key) || "";
                if (!stored) {
                    const proposed = newId("install");
                    window.localStorage.setItem(key, proposed);
                    stored = window.localStorage.getItem(key) || "";
                    if (stored !== proposed) stored = "";
                }
                if (stored) {
                    this.approvedVaultRoot = physicalRoot;
                    this.scopedInstallationKey = key;
                    localProof = stored;
                }
            }
        } catch {
            localProof = "";
            this.approvedVaultRoot = "";
            this.scopedInstallationKey = "";
        }
        if (!localProof || oneDrive.installationProofId !== localProof) {
            // Forensic local-only receipt, never an execution grant. It permits an explicit
            // NEW cloud copy to retain completed historical metadata until final binding.
            const isolatedPrior = oneDrive.installationProofId && oneDrive.firstSyncCompleted === true
                ? JSON.parse(JSON.stringify(oneDrive)) : null;
            oneDrive.confirmed = false;
            oneDrive.runAllowed = false;
            oneDrive.enabled = false;
            oneDrive.firstSyncCompleted = false;
            // Do not erase incident evidence when localStorage proof is lost,
            // a Vault is copied, or a process crashes before browser storage flush.
            // Existing file states and pending approvals are forensic/read-only:
            // confirmed/runAllowed remain false and cannot authorize any writes.
            oneDrive.status = "paused";
            oneDrive.lastError = "检测到知识库位置、设备身份或安装授权变化，执行权限已禁用；历史基线与未决事项仅供人工核查，不得自动恢复同步。";
            oneDrive.deviceId = newId("device");
            if (isolatedPrior && localProof) (this.settings.local as any).oneDriveScopeIsolation = {
                accountId: String(this.settings.local.microsoftAccountId || ""),
                physicalRoot: this.approvedVaultRoot, proofId: localProof, deviceId: oneDrive.deviceId,
                prior: isolatedPrior, isolatedAt: Date.now(),
            };
        } else {
            const currentScope = [
                this.settings.local.microsoftAccountId, oneDrive.vaultId,
                oneDrive.deviceId, oneDrive.installationProofId,
                oneDrive.confirmationVaultId, oneDrive.confirmationAccountId,
            ];
            const identicalLiveScope = previousScope.every(
                (value, i) => value === currentScope[i],
            );
            const accountId = this.settings.local.microsoftAccountId;
            const completedBinding = hasExistingScope && oneDrive.firstSyncCompleted && oneDrive.confirmed && accountId &&
                oneDrive.confirmationAccountId === accountId && oneDrive.confirmationVaultId === oneDrive.vaultId;
            const uploadRecords = oneDrive.pendingUploadOperations;
            const unresolvedWrite = hasUnfinishedOneDriveSetup(oneDrive, accountId, oneDrive.vaultId, oneDrive.deviceId) ||
                oneDrive.androidFirstDownload ||
                (uploadRecords && (typeof uploadRecords !== "object" || Array.isArray(uploadRecords)));
            const keepConnection = this.loadedInThisProcess
                ? activeLiveGrant && identicalLiveScope
                : completedBinding && oneDrive.status !== "stopping" && !unresolvedWrite;
            if (!keepConnection) {
                oneDrive.runAllowed = false;
                oneDrive.enabled = false;
            } else if (!this.loadedInThisProcess && !oneDrive.runAllowed && !oneDrive.enabled &&
                oneDrive.status === "paused" &&
                oneDrive.lastError === LEGACY_ONEDRIVE_RESTART_MESSAGE) {
                // The previous version revoked the binding on every restart.
                // Restore manual sync only: that version cleared the auto toggle.
                oneDrive.runAllowed = true;
                oneDrive.status = "ready";
                oneDrive.lastError = "";
            }
            if (!this.loadedInThisProcess && (oneDrive.status === "syncing" || oneDrive.status === "stopping")) {
                // No active operation survives a process restart. A new round
                // still checks any unresolved upload journal before writing.
                oneDrive.status = oneDrive.runAllowed ? "ready" : "paused";
            }
            if (!this.loadedInThisProcess && oneDrive.runAllowed) {
                // Preserve automatic transactions for reconciliation only.
                // New writes always come from the next round's fresh plan.
                oneDrive.pendingDeletions = Object.fromEntries(Object.entries(oneDrive.pendingDeletions || {})
                    .filter(([path, record]) => record?.automatic === true &&
                        record.accountId === accountId && record.vaultId === oneDrive.vaultId &&
                        record.deviceId === oneDrive.deviceId &&
                        (record.direction === "cloud" || record.direction === "vault") &&
                        record.remoteId && oneDrive.fileStates?.[path]?.remoteId === record.remoteId &&
                        oneDrive.fileStates?.[path]?.remoteEtag === record.remoteEtag &&
                        !path.includes("\\") && path.split("/").every(part => part && part !== "." && part !== "..")));
            }
            if (!oneDrive.runAllowed) oneDrive.enabled = false;
            if (!keepConnection && oneDrive.firstSyncCompleted && !oneDrive.lastError) {
                oneDrive.status = "paused";
            }
        }
        oneDrive.installationProofId = localProof;
        this.loadedInThisProcess = true;
        await this.save();
    }

    installationProofIsValid(): boolean {
        try {
            const root = this.currentPhysicalVaultRoot();
            if (!root || root !== this.approvedVaultRoot || !this.scopedInstallationKey) {
                return false;
            }
            const proof = window.localStorage.getItem(this.scopedInstallationKey) || "";
            return Boolean(proof && this.settings.local.oneDrive.installationProofId === proof);
        } catch {
            return false;
        }
    }

    async save(): Promise<void> {
        await this.persistSettings();
    }

    async saveWithIdentityGuard(beforePersist: () => void): Promise<void> {
        await this.persistSettings(beforePersist);
    }

    private async persistSettings(beforePersist?: () => void): Promise<void> {
        // All modules share data.json. An older background save must not finish
        // after a confirmed journal/baseline save and resurrect an earlier state.
        const operation = this.persistenceQueue.then(async () => {
            beforePersist?.();
            const raw = (await this.plugin.loadData()) || {};
            beforePersist?.();
            // Freeze the same JSON representation the plugin persists, so an
            // in-flight writer never observes later in-memory phase changes.
            const payload = JSON.parse(JSON.stringify({ ...raw, suiteV2: this.settings }));
            if (this.preserveInactiveOneDrive) {
                const previous = raw.suiteV2?.local?.oneDrive;
                if (previous === undefined) delete payload.suiteV2.local.oneDrive;
                else payload.suiteV2.local.oneDrive = previous;
            }
            await this.plugin.saveData(payload);
            for (const listener of this.listeners) listener();
        });
        // A failed write is still returned to its own caller, but does not poison
        // the queue or silently skip later safety checkpoints.
        this.persistenceQueue = operation.catch(() => {});
        await operation;
    }

    /** Read persisted registration facts only; never expose credentials or renew permission. */
    async readPersistedOneDriveRegistrationState(beforeRead?: () => void): Promise<any> {
        const operation = this.persistenceQueue.then(async () => {
            beforeRead?.();
            const raw = await this.plugin.loadData();
            beforeRead?.();
            const local = raw?.suiteV2?.local;
            const sync = local?.oneDrive;
            if (!sync || typeof sync !== "object") return null;
            const selected: Record<string, any> = {};
            for (const key of ["vaultId", "localVaultId", "cloudVaultId", "deviceId", "installationProofId",
                "confirmed", "firstSyncCompleted", "confirmationVaultId", "confirmationAccountId",
                "runAllowed", "enabled", "fileStates", "identicalRegistrationJournal", "firstConflictPreservationJournal", "firstCommitJournal", "firstFileJournal"])
                selected[key] = sync[key];
            return JSON.parse(JSON.stringify({ accountId: String(local.microsoftAccountId || ""), oneDrive: selected, oneDriveBindingHistory: local.oneDriveBindingHistory || [] }));
        });
        this.persistenceQueue = operation.then(() => {}, () => {});
        return await operation;
    }

    subscribe(listener: () => void): () => void {
        this.listeners.add(listener);
        return () => this.listeners.delete(listener);
    }

    async updateShared(shared: Partial<SharedSmartMDState>): Promise<void> {
        this.settings.shared = {
            ...this.settings.shared,
            ...shared,
            updatedAt: Date.now(),
        };
        await this.save();
    }

    exportShared(): SharedSmartMDState {
        return sharedStateForSync(this.settings.shared);
    }

    async importShared(incoming: SharedSmartMDState): Promise<void> {
        if (!incoming || incoming.updatedAt <= this.settings.shared.updatedAt) return;
        // Never replace local publishing/AI-private data with remote fields.
        const approved = sharedStateForSync(incoming);
        this.settings.shared = {
            ...this.settings.shared,
            ...approved,
        };
        await this.save();
    }

    private secretStorage(): any {
        try {
            return (this.app as any).secretStorage;
        } catch {
            return null;
        }
    }

    private localSecretKey(name: string): string {
        const adapter = this.app.vault.adapter as any;
        const vaultLocation = String(adapter?.basePath || adapter?.fullPath || this.app.vault.getName());
        return `smartmd:device-secret:${vaultLocation}:${name}`;
    }

    private legacyLocalSecretKey(name: string): string {
        const adapter = this.app.vault.adapter as any;
        const vaultLocation = String(adapter?.basePath || adapter?.fullPath || this.app.vault.getName());
        return `smartkeep-connect:device-secret:${vaultLocation}:${name}`;
    }

    async setSecret(name: string, value: string): Promise<void> {
        if (!value) {
            // Clear both current and legacy plaintext remnants, even if no
            // secure credential can currently be read.
            return await this.clearSecret(name);
        }
        const storage = this.secretStorage();
        if (!storage || typeof storage.setSecret !== "function" ||
            typeof storage.getSecret !== "function") {
            this.secretCache.delete(name);
            throw new Error("当前设备安全存储不可用，已拒绝保存授权或密钥。");
        }
        try {
            await Promise.resolve(storage.setSecret(name, value));
            // Do not accept a fake-successful write or an unreadable host as
            // durable secret storage. In particular, never fall back to plaintext.
            const stored = await Promise.resolve(storage.getSecret(name));
            if (stored !== value) throw new Error("Secure readback failed");
        } catch {
            this.secretCache.delete(name);
            throw new Error("当前设备安全存储写入或读回失败，已拒绝保存授权或密钥。");
        }
        this.secretCache.set(name, value);
        this.compatibilitySecrets.delete(name);
        // Never leave a superseded v1 plaintext key behind when cleanup works.
        // Reads are never allowed to fall back to either key.
        try {
            window.localStorage.removeItem(this.localSecretKey(name));
            window.localStorage.removeItem(this.legacyLocalSecretKey(name));
        } catch {}
    }

    private async clearSecret(name: string): Promise<void> {
        const storage = this.secretStorage();
        let secureCleared = true;
        if (storage) {
            try {
                if (typeof storage.getSecret !== "function") {
                    throw new Error("Secure readback unavailable");
                }
                if (typeof storage.deleteSecret === "function") {
                    await Promise.resolve(storage.deleteSecret(name));
                } else if (typeof storage.removeSecret === "function") {
                    await Promise.resolve(storage.removeSecret(name));
                } else if (typeof storage.setSecret === "function") {
                    await Promise.resolve(storage.setSecret(name, ""));
                } else {
                    throw new Error("Secure deletion unavailable");
                }
                secureCleared = !(await Promise.resolve(storage.getSecret(name)));
            } catch {
                secureCleared = false;
            }
        }
        // Always clear BOTH plaintext namespaces. Never migrate a legacy
        // credential back to secure storage or resurrect it from the cache.
        let localCleared = false;
        try {
            const keys = [this.localSecretKey(name), this.legacyLocalSecretKey(name)];
            for (const key of keys) window.localStorage.removeItem(key);
            localCleared = keys.every((key) => window.localStorage.getItem(key) === null);
        } catch {}
        this.secretCache.delete(name);
        this.compatibilitySecrets.delete(name);
        if (!secureCleared || !localCleared) {
            throw new Error("本机敏感信息未能完整清除，登录状态保持不变，请检查存储权限后重试。");
        }
    }

    getSecret(name: string): string {
        // Only a readable host SecretStorage value can authorize credentials.
        // Never read legacy plaintext or a stale in-memory cache as a fallback.
        const storage = this.secretStorage();
        try {
            const value = typeof storage?.getSecret === "function"
                ? storage.getSecret(name) : "";
            if (typeof value === "string" && value) {
                this.secretCache.set(name, value);
                return value;
            }
        } catch {}
        this.secretCache.delete(name);
        return "";
    }

    storageBoundaryText(): string {
        return typeof this.secretStorage()?.getSecret === "function"
            ? "当前设备使用 Obsidian 安全存储；授权和密钥不会参与同步。"
            : "当前设备安全存储不可用；不会降级使用明文兼容存储，请重新配置安全存储。";
    }

    async setAIKey(value: string, address = this.settings.shared.aiApiUrl): Promise<void> {
        const origin = httpModelOrigin(address);
        if (origin) {
            if (value) await this.setSecret(AI_HTTP_SECRET, JSON.stringify({ origin, key: value }));
            else if (this.getAIKey(address)) await this.setSecret(AI_HTTP_SECRET, "");
            return;
        }
        if (/^http:/i.test(address.trim())) throw new Error("请先填写有效的 HTTP 模型服务地址");
        await this.setSecret(AI_SECRET, value);
    }

    getAIKey(address = this.settings.shared.aiApiUrl): string {
        const origin = httpModelOrigin(address);
        if (!origin) return /^http:/i.test(address.trim()) ? "" : this.getSecret(AI_SECRET);
        try {
            const saved = JSON.parse(this.getSecret(AI_HTTP_SECRET) || "null");
            return saved?.origin === origin && typeof saved.key === "string" ? saved.key : "";
        } catch {
            return "";
        }
    }

    async setMicrosoftToken(value: MicrosoftToken | null): Promise<void> {
        await this.setSecret(MICROSOFT_SECRET, value ? JSON.stringify(value) : "");
    }

    getMicrosoftToken(): MicrosoftToken | null {
        const raw = this.getSecret(MICROSOFT_SECRET);
        if (!raw) return null;
        try {
            return JSON.parse(raw) as MicrosoftToken;
        } catch {
            return null;
        }
    }
}

export async function sha256Text(value: string): Promise<string> {
    const bytes = new TextEncoder().encode(value);
    return sha256Bytes(bytes.buffer);
}

export async function sha256Bytes(value: ArrayBuffer): Promise<string> {
    const digest = await crypto.subtle.digest("SHA-256", value);
    return Array.from(new Uint8Array(digest))
        .map((byte) => byte.toString(16).padStart(2, "0"))
        .join("");
}

export function normalizeVaultPath(path: string): string {
    return path.replace(/\\/g, "/").replace(/^\/+|\/+$/g, "").replace(/\/{2,}/g, "/");
}

export function isRemoteImage(source: string): boolean {
    return /^(https?:)?\/\//i.test(source) || /^data:/i.test(source);
}

export function findLocalImages(markdown: string): string[] {
    const result = new Set<string>();
    const wiki = /!\[\[([^\]|#]+)(?:[|#][^\]]*)?\]\]/g;
    const standard = /!\[[^\]]*\]\(([^)\s]+)(?:\s+["'][^"']*["'])?\)/g;
    let match: RegExpExecArray | null;
    while ((match = wiki.exec(markdown))) result.add(match[1].trim());
    while ((match = standard.exec(markdown))) {
        const source = match[1].replace(/^<|>$/g, "").trim();
        if (!isRemoteImage(source)) result.add(source);
    }
    return Array.from(result);
}

export function pathMatches(path: string, patterns: string[]): boolean {
    return patterns.some((pattern) => {
        try {
            return new RegExp(pattern, "i").test(path);
        } catch {
            return false;
        }
    });
}

export function formatTime(value: number): string {
    return value ? new Date(value).toLocaleString() : "从未";
}

export class MicrosoftAuth {
    private signingIn = false;
    private pendingVerifier = "";
    private pendingState = "";
    private pendingResolve: (() => void) | null = null;
    private pendingReject: ((error: Error) => void) | null = null;
    private pendingTimer: number | null = null;
    private deviceSession: MicrosoftDeviceCodeSession | null = null;
    // Invalidate in-flight refreshes synchronously when the Microsoft session
    // is revoked. Serialize writes so a late secure-storage write cannot
    // complete after the logout clear.
    private sessionEpoch = 0;
    private tokenMutationTail: Promise<void> = Promise.resolve();
    // A grant is never a usable credential while its secret/account commit is in flight.
    private credentialCommitInProgress = false;
    private credentialSafetyBlocked = false;
    private sessionRevocationsPending = 0;

    private async serializeTokenMutation(work: () => Promise<void>): Promise<void> {
        const previous = this.tokenMutationTail;
        let release!: () => void;
        this.tokenMutationTail = new Promise<void>((resolve) => { release = resolve; });
        await previous;
        try { await work(); } finally { release(); }
    }

    private assertMicrosoftSession(epoch: number, accountId: string): void {
        if (this.sessionEpoch !== epoch || this.sessionRevocationsPending > 0 ||
            this.credentialCommitInProgress || this.credentialSafetyBlocked ||
            String(this.store.settings.local.microsoftAccountId || "") !== accountId) {
            throw new Error("微软账号会话已变化，原异步操作已停止。");
        }
    }

    // OneDrive's direct conditional PUT/DELETE transport uses fetch rather
    // than requestGraph. Give it the same post-token revocation fence.
    createDispatchGuard(): () => void {
        const epoch = this.sessionEpoch;
        const accountId = String(this.store.settings.local.microsoftAccountId || "");
        return () => this.assertMicrosoftSession(epoch, accountId);
    }

    constructor(private store: SmartMDStore) {}

    isSignedIn(): boolean {
        return !this.credentialCommitInProgress && !this.credentialSafetyBlocked &&
            this.sessionRevocationsPending === 0 &&
            Boolean(this.store.settings.local.microsoftAccountId && this.store.getMicrosoftToken());
    }

    isSignInPending(): boolean {
        return this.signingIn || this.deviceSession !== null;
    }

    cancelSignIn(): void {
        // Cancellation invalidates even a callback already inside an awaited OAuth request.
        if (this.signingIn || this.deviceSession !== null) this.sessionEpoch++;
        if (this.signingIn) this.finishSignIn(new Error("微软登录已取消。"));
        const device = this.deviceSession;
        this.deviceSession = null;
        device?.cancel();
    }

    async signOut(): Promise<void> {
        // Revoke synchronously, then queue the clear behind any staged login.
        // Do not report logout complete while a previous secret write can still finish.
        this.sessionRevocationsPending++;
        this.sessionEpoch++;
        this.cancelSignIn();
        try {
            await this.serializeTokenMutation(async () => {
                await this.store.setMicrosoftToken(null);
                this.store.settings.local.microsoftAccountId = "";
                this.store.settings.local.microsoftAccountName = "";
                await this.store.save();
            });
            this.credentialSafetyBlocked = false;
        } catch (error) {
            // A failed secure clear cannot be mistaken for a usable signed-in state.
            this.credentialSafetyBlocked = true;
            throw error;
        } finally {
            this.sessionRevocationsPending--;
        }
    }

    async signIn(openUrl: (url: string) => void | Promise<void>): Promise<void> {
        if (this.isSignInPending() || this.sessionRevocationsPending > 0 ||
            this.credentialSafetyBlocked) throw new Error("微软登录正忙或凭据尚未安全清除。");
        const clientId = MICROSOFT_CLIENT_ID;
        if (!clientId) throw new Error("微软登录暂不可用。");
        // Each new browser grant owns a unique epoch; a late opener/callback
        // cannot settle a later request, even if the OAuth state is reused.
        const epoch = ++this.sessionEpoch;
        this.signingIn = true;
        try {
            this.pendingVerifier = this.randomValue(64);
            this.pendingState = this.randomValue(24);
            const challenge = await this.createChallenge(this.pendingVerifier);
            if (this.sessionEpoch !== epoch || !this.signingIn) {
                throw new Error("微软登录已取消。");
            }
            const query = new URLSearchParams({
                client_id: clientId,
                response_type: "code",
                redirect_uri: MICROSOFT_REDIRECT_URI,
                response_mode: "query",
                scope: MICROSOFT_SCOPE,
                state: this.pendingState,
                code_challenge: challenge,
                code_challenge_method: "S256",
                prompt: "select_account",
            });
            const result = new Promise<void>((resolve, reject) => {
                this.pendingResolve = resolve;
                this.pendingReject = reject;
                this.pendingTimer = window.setTimeout(() => {
                    if (this.sessionEpoch === epoch && this.signingIn) {
                        this.finishSignIn(new Error("微软登录已超时，请重新开始。"));
                    }
                }, 10 * 60 * 1000);
            });
            void result.catch(() => {});
            await openUrl("https://login.microsoftonline.com/consumers/oauth2/v2.0/authorize?" + query.toString());
            if (this.sessionEpoch !== epoch) throw new Error("微软登录已取消。");
            await result;
        } catch (error) {
            if (this.sessionEpoch === epoch && this.signingIn) this.finishSignIn(error);
            throw error;
        }
    }

    async completeSignIn(params: Record<string, string>): Promise<void> {
        if (!this.signingIn || !this.pendingVerifier || !this.pendingState) return;
        const epoch = this.sessionEpoch;
        const verifier = this.pendingVerifier;
        const state = this.pendingState;
        const active = () => {
            if (this.sessionEpoch !== epoch || this.sessionRevocationsPending > 0 ||
                !this.signingIn || this.pendingState !== state || this.pendingVerifier !== verifier) {
                throw new Error("微软网页登录请求已撤销，迟到回调不得恢复授权。");
            }
        };
        try {
            active();
            // An unrelated or old redirect must not cancel a newer OAuth
            // request simply by delivering the previous state afterward.
            if (params.state !== state) return;
            if (params.error) throw new Error(params.error_description || params.error);
            if (!params.code) throw new Error("微软没有返回有效的登录结果。");
            const tokenResponse = await requestUrl({
                url: "https://login.microsoftonline.com/consumers/oauth2/v2.0/token",
                method: "POST",
                contentType: "application/x-www-form-urlencoded",
                throw: false,
                body: new URLSearchParams({
                    client_id: MICROSOFT_CLIENT_ID,
                    grant_type: "authorization_code",
                    code: params.code,
                    redirect_uri: MICROSOFT_REDIRECT_URI,
                    code_verifier: verifier,
                    scope: MICROSOFT_SCOPE,
                }).toString(),
            });
            active();
            if (tokenResponse.status >= 400 || !tokenResponse.json?.access_token ||
                !tokenResponse.json?.refresh_token) {
                throw new Error(tokenResponse.json?.error_description || "微软登录失败。");
            }
            const result = tokenResponse.json;
            const token: MicrosoftToken = {
                accessToken: result.access_token,
                refreshToken: result.refresh_token,
                expiresAt: Date.now() + Number(result.expires_in || 3600) * 1000,
            };
            // Verify identity with the transient grant; never stage B's secret
            // under A's account while the /me request is still outstanding.
            const me = await this.deviceGraphCheck(
                "/me?$select=id,displayName,userPrincipalName", token.accessToken,
            );
            active();
            await this.persistVerifiedMicrosoftLogin(token, me, active);
            active();
            this.finishSignIn();
        } catch (error) {
            // A stale callback must not clear or settle a newer browser/device
            // session. Credential rollback is owned by the serialized commit.
            if (this.sessionEpoch === epoch && this.signingIn &&
                this.pendingState === state && this.pendingVerifier === verifier) {
                this.finishSignIn(error);
            }
        }
    }

    /**
     * Preferred cross-platform OneDrive sign-in. This grant never needs the
     * Windows obsidian:// protocol and never exposes a client secret.
     */
    async signInWithDeviceCode(
        present: (challenge: MicrosoftDeviceCodeChallenge) => void | Promise<void>,
    ): Promise<void> {
        if (this.isSignInPending() || this.sessionRevocationsPending > 0 ||
            this.credentialSafetyBlocked) throw new Error("微软登录正忙或凭据尚未安全清除。");
        // Invalidate operations started under the old account when switching.
        this.sessionEpoch++;
        const session = new MicrosoftDeviceCodeSession(MICROSOFT_CLIENT_ID, MICROSOFT_SCOPE);
        this.deviceSession = session;
        try {
            const token = await session.authorize(present);
            if (this.deviceSession !== session) throw new Error("微软设备码登录已取消。");
            await this.verifyAndStoreDeviceToken(token, session);
        } finally {
            if (this.deviceSession === session) this.deviceSession = null;
        }
    }

    private async deviceGraphCheck(path: string, accessToken: string): Promise<any> {
        let response: any;
        try {
            response = await requestUrl({
                url: "https://graph.microsoft.com/v1.0" + path,
                method: "GET",
                headers: { Authorization: "Bearer " + accessToken },
                throw: false,
            });
        } catch {
            throw new Error("无法连接 Microsoft Graph，登录未保存，请检查网络。");
        }
        if (response.status !== 200 || !response.json?.id) {
            throw new Error("无法验证微软账号身份（HTTP " +
                Number(response.status || 0) + "），登录未保存。");
        }
        return response.json;
    }

    // Browser OAuth and device code share one serial, reversible credential
    // transaction. Logout always queues behind it and therefore wins last.
    private async persistVerifiedMicrosoftLogin(
        token: MicrosoftToken, me: any, assertActive: () => void,
    ): Promise<void> {
        await this.serializeTokenMutation(async () => {
            assertActive();
            const local = this.store.settings.local;
            const previous = this.store.getMicrosoftToken();
            const previousId = local.microsoftAccountId;
            const previousName = local.microsoftAccountName;
            let beganCommit = false;
            this.credentialCommitInProgress = true;
            try {
                assertActive();
                beganCommit = true;
                await this.store.setMicrosoftToken(token);
                assertActive();
                local.microsoftAccountId = String(me.id);
                local.microsoftAccountName = String(
                    me.displayName || me.userPrincipalName || "Microsoft user",
                );
                await this.store.save();
                assertActive();
            } catch (error) {
                if (beganCommit) {
                    try {
                        // Roll back while holding the same mutex. A queued
                        // signOut cannot finish and then be overwritten.
                        await this.store.setMicrosoftToken(previous);
                        local.microsoftAccountId = previousId;
                        local.microsoftAccountName = previousName;
                        await this.store.save();
                    } catch {
                        // The state is unknown; block every Graph dispatch
                        // until an explicit successful logout clears it.
                        this.credentialSafetyBlocked = true;
                        this.sessionEpoch++;
                        local.microsoftAccountId = "";
                        local.microsoftAccountName = "";
                        throw new Error("授权保存失败且本机凭据回滚不完整，请停止同步并检查登录状态。");
                    }
                }
                throw error;
            } finally {
                this.credentialCommitInProgress = false;
            }
        });
    }

    private async verifyAndStoreDeviceToken(
        token: MicrosoftDeviceToken,
        session: MicrosoftDeviceCodeSession,
    ): Promise<void> {
        const epoch = this.sessionEpoch;
        const active = () => {
            session.assertActive();
            if (this.deviceSession !== session || this.sessionEpoch !== epoch ||
                this.sessionRevocationsPending > 0) {
                throw new Error("微软设备码登录已取消。");
            }
        };
        active();
        const me = await this.deviceGraphCheck(
            "/me?$select=id,displayName,userPrincipalName", token.accessToken,
        );
        active();
        try {
            await this.persistVerifiedMicrosoftLogin(token, me, active);
        } catch (error) {
            if (this.credentialSafetyBlocked) throw error;
            throw new Error("微软设备码登录被取消或保存失败，原有授权状态已保留。");
        }
    }

    private randomValue(byteLength: number): string {
        const bytes = new Uint8Array(byteLength);
        crypto.getRandomValues(bytes);
        return this.base64Url(bytes);
    }

    private async createChallenge(verifier: string): Promise<string> {
        const digest = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(verifier));
        return this.base64Url(new Uint8Array(digest));
    }

    private base64Url(bytes: Uint8Array): string {
        let binary = "";
        bytes.forEach((value) => binary += String.fromCharCode(value));
        return window.btoa(binary).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/g, "");
    }

    private finishSignIn(error?: unknown): void {
        if (this.pendingTimer !== null) window.clearTimeout(this.pendingTimer);
        const resolve = this.pendingResolve;
        const reject = this.pendingReject;
        this.pendingTimer = null;
        this.pendingResolve = null;
        this.pendingReject = null;
        this.pendingVerifier = "";
        this.pendingState = "";
        this.signingIn = false;
        if (error) reject?.(error instanceof Error ? error : new Error(String(error)));
        else resolve?.();
    }

    private async invalidateMicrosoftSession(): Promise<void> {
        // Use the same epoch and serialized clear as an explicit logout.
        // The calling refresh/401 path already checked ownership.
        await this.signOut();
    }

    private isGraphAuthFailure(response: any): boolean {
        return response?.status === 401;
    }

    private isRefreshCredentialInvalid(response: any): boolean {
        if (![400, 401].includes(Number(response?.status || 0))) return false;
        const detail = [
            response?.json?.error,
            response?.json?.error_description,
            response?.text,
        ].filter(Boolean).join(" ");
        return /invalid_grant|refresh token.*(?:invalid|expired|revoked)|revok|AADSTS70008|AADSTS700082/i.test(detail);
    }

    async getAccessToken(forceRefresh = false, readGuard?: () => void): Promise<string> {
        const epoch = this.sessionEpoch;
        const accountId = String(this.store.settings.local.microsoftAccountId || "");
        const token = this.store.getMicrosoftToken();
        if (!token) throw new Error("请先登录微软账号。");
        const requireOriginal = () => {
            readGuard?.();
            this.assertMicrosoftSession(epoch, accountId);
            const live = this.store.getMicrosoftToken();
            if (!live || live.refreshToken !== token.refreshToken ||
                live.accessToken !== token.accessToken) {
                throw new Error("微软令牌已由其他操作更新，旧刷新请求不得落地。");
            }
        };
        requireOriginal();
        if (!forceRefresh && token.expiresAt > Date.now() + 120_000) {
            return token.accessToken;
        }
        const clientId = MICROSOFT_CLIENT_ID;
        const response = await requestUrl({
            url: "https://login.microsoftonline.com/consumers/oauth2/v2.0/token",
            method: "POST",
            contentType: "application/x-www-form-urlencoded",
            throw: false,
            body:
                "client_id=" + encodeURIComponent(clientId) +
                "&grant_type=refresh_token" +
                "&refresh_token=" + encodeURIComponent(token.refreshToken) +
                "&scope=" + encodeURIComponent(MICROSOFT_SCOPE),
        });
        // Even a failed stale refresh must not sign out a newer account.
        requireOriginal();
        if (response.status >= 400 || !response.json?.access_token) {
            if (this.isRefreshCredentialInvalid(response)) {
                await this.invalidateMicrosoftSession();
                throw new Error("微软授权已失效，请在 OneDrive 页面重新登录。");
            }
            throw new Error("暂时无法续约微软授权，登录仍保留；网络恢复后可直接重试。");
        }
        const updated: MicrosoftToken = {
            accessToken: response.json.access_token,
            refreshToken: response.json.refresh_token || token.refreshToken,
            expiresAt: Date.now() + Number(response.json.expires_in || 3600) * 1000,
        };
        await this.serializeTokenMutation(async () => {
            requireOriginal();
            await this.store.setMicrosoftToken(updated);
            // Logout can revoke while SecretStorage is awaiting physical I/O.
            // The queued logout clear must win; never return that old token.
            this.assertMicrosoftSession(epoch, accountId);
        });
        return updated.accessToken;
    }

    async requestGraph(pathOrUrl: string, options: { method?: string; body?: any; contentType?: string; headers?: Record<string, string>; beforeDispatch?: () => void } = {}): Promise<any> {
        const url = pathOrUrl.startsWith("http")
            ? pathOrUrl
            : "https://graph.microsoft.com/v1.0" + pathOrUrl;
        const epoch = this.sessionEpoch;
        const accountId = String(this.store.settings.local.microsoftAccountId || "");
        // A canceled read must not renew or invalidate the session on a late 401.
        const readGuard = !options.method || options.method === "GET" ? options.beforeDispatch : undefined;
        const send = async (token: string): Promise<any> => {
            // Final synchronous gate: no await between this check, the
            // caller's operation ticket and the physical Graph dispatch.
            this.assertMicrosoftSession(epoch, accountId);
            options.beforeDispatch?.();
            return requestUrl({
                url,
                method: options.method || "GET",
                headers: { Authorization: "Bearer " + token, ...(options.headers || {}) },
                contentType: options.contentType || (options.body === undefined ? undefined : "application/json"),
                body:
                    options.body === undefined
                        ? undefined
                        : options.contentType && options.contentType !== "application/json"
                          ? options.body
                          : JSON.stringify(options.body),
                throw: false,
            });
        };
        let response = await send(await this.getAccessToken(false, readGuard));
        readGuard?.();
        if (this.isGraphAuthFailure(response)) {
            this.assertMicrosoftSession(epoch, accountId);
            response = await send(await this.getAccessToken(true, readGuard));
            readGuard?.();
            if (this.isGraphAuthFailure(response)) {
                this.assertMicrosoftSession(epoch, accountId);
                await this.invalidateMicrosoftSession();
            }
        }
        return response;
    }

    async graph(pathOrUrl: string, options: { method?: string; body?: any; contentType?: string; headers?: Record<string, string>; beforeDispatch?: () => void } = {}): Promise<any> {
        const response = await this.requestGraph(pathOrUrl, options);
        if (response.status >= 400) {
            if (this.isGraphAuthFailure(response)) {
                const error: any = new Error("微软授权已失效，请重新登录。");
                error.status = response.status;
                throw error;
            }
            const errObj = response.json?.error || {};
            const code = String(errObj.code || response.json?.error || `HTTP_${response.status}`);
            const innerMsg = errObj.innerError?.message;
            const message = String(innerMsg || errObj.message || response.json?.error_description || response.text || "OneDrive 请求失败。");
            const requestId = errObj.innerError?.["request-id"] || response.headers?.["request-id"];
            const error = new GraphApiError({
                status: response.status,
                code,
                message,
                method: options.method || "GET",
                url: pathOrUrl,
                requestId,
                retryAfterMs: graphRetryAfterMs(response.headers?.["retry-after"]),
            });
            throw error;
        }
        return response.json ?? response.arrayBuffer;
    }
}

export function graphRetryAfterMs(header: string | null | undefined): number {
    if (!header) return 0;
    const seconds = Number(header);
    const delay = Number.isFinite(seconds) ? seconds * 1000 : Date.parse(header) - Date.now();
    return Number.isFinite(delay) ? Math.max(0, delay) : 0;
}

export class GraphApiError extends Error {
    status: number;
    code: string;
    method: string;
    url: string;
    requestId?: string;
    retryAfterMs?: number;

    constructor(options: { status: number; code?: string; message: string; method?: string; url?: string; requestId?: string; retryAfterMs?: number }) {
        const shortUrl = (options.url || "").replace("https://graph.microsoft.com/v1.0", "");
        const reqDetail = options.requestId ? ` (request-id: ${options.requestId})` : "";
        super(`[Graph API HTTP ${options.status} ${options.code || "Error"}] ${options.method || "GET"} ${shortUrl} -> ${options.message}${reqDetail}`);
        this.name = "GraphApiError";
        this.status = options.status;
        this.code = options.code || `HTTP_${options.status}`;
        this.method = options.method || "GET";
        this.url = options.url || "";
        this.requestId = options.requestId;
        this.retryAfterMs = options.retryAfterMs;
    }
}

export function reportError(prefix: string, error: unknown): void {
    console.error(prefix, error);
    new Notice(localizeMessage(prefix + "：" + (error instanceof Error ? error.message : String(error))), 8000);
}
