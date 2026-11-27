import { localize, localizeMessage } from "./localization";
import {
    App,
    Modal,
    Notice,
    normalizePath,
    Platform,
    requestUrl,
    TFile,
} from "obsidian";
import { graphRetryAfterMs, MicrosoftAuth, SmartMDStore } from "./core";
import {
    SMARTKEEP_PUBLISH_URL,
    hasUnfinishedOneDriveSetup,
    newId,
    sharedStateForSync,
    SharedSmartMDState,
    SyncFailureDiagnostic,
    type SyncFileState,
    SyncSummary,
    type OneDriveSyncSettings,
    SyncFailureCategory,
    ActionableRemediation,
    type AndroidFirstDownload,
} from "./types";
import { OperationAuthority, RunTicket } from "./operation-authority";
import { inspectAndroidReceiverCandidateInventory, inspectReceiverCandidateInventory } from "./onedrive-receiver-audit";
import { getAndroidAppStorageRoot } from "./onedrive-android-vault";
import { placeAndroidStagedFile, replaceAndroidBinaryIfUnchanged, replaceAndroidMarkdownIfUnchanged } from "./onedrive-android-write";
import { redactOneDriveText, safeOneDriveError } from "./onedrive-error-redaction";


type RemoteFile = {
    id: string;
    path: string;
    name: string;
    size: number;
    etag: string;
    downloadUrl?: string;
    quickXorHash?: string;
    sha1Hash?: string;
    mtime?: number;
};

type LocalFile = {
    path: string;
    file: TFile;
    size: number;
    mtime: number;
};

type RemoteVault = {
    id: string;
    name: string;
    updatedAt: number;
    registrationOpId?: string;
    registerIdentical?: boolean;
    mergeAndroid?: boolean;
};

type SyncPreviewItem = {
    source: string;
    local?: { size: number; mtime: number; sha256?: string };
    remote?: { id: string; etag: string; size: number };
    skipReasons: string[];
};

type SyncPreview = {
    localCount: number;
    remoteCount: number;
    uploadCount: number;
    downloadCount: number;
    skippedCount: number;
    uploadPaths: string[];
    downloadPaths: string[];
    samePaths: string[];
    skippedPaths: string[];
    unknownPaths: string[];
    unexpectedNewTarget: boolean;
    identity?: { accountId: string; localVaultId: string; cloudVaultId: string };
    androidCloudVaultName?: string;
    creatingCloudVault?: boolean;
    items?: Record<string, SyncPreviewItem>;
    priorHistoryCount?: number;
    resumedFirstUpload?: boolean;
    finishFirstUploadOnly?: boolean;
};

function encodePath(path: string): string {
    return path.split("/").filter(Boolean).map(encodeURIComponent).join("/");
}

function requireWindowsHost(): void {
    if (!Platform.isDesktopApp || !Platform.isWin) {
        throw new Error("该高级 OneDrive 操作需要 Windows 桌面版。");
    }
}

function ignoreMatcher(pattern: string): RegExp {
    if (typeof pattern !== "string" || pattern.length > 128 ||
        /\([^)]*\)[+*{]|\(\?|\\[1-9]|[+*?]\s*[+*{]/.test(pattern)) {
        throw new Error("OneDrive 同步排除规则不安全，已停止执行");
    }
    try {
        return new RegExp(pattern, "i");
    } catch {
        const escaped = pattern.replace(/[.+^${}()|[\]\\]/g, "\\$&").replace(/\*/g, ".*").replace(/\?/g, ".");
        return new RegExp(`^${escaped}$`, "i");
    }
}

function sameBytes(left: ArrayBuffer, right: ArrayBuffer): boolean {
    if (left.byteLength !== right.byteLength) return false;
    const a = new Uint8Array(left);
    const b = new Uint8Array(right);
    for (let index = 0; index < a.length; index++) {
        if (a[index] !== b[index]) return false;
    }
    return true;
}

async function contentHash(value: string): Promise<string> {
    const digest = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(value));
    return Array.from(new Uint8Array(digest)).map((part) => part.toString(16).padStart(2, "0")).join("");
}

interface PreparedSync {
    local: LocalFile[];
    allRemote: RemoteFile[];
    androidMerge?: Map<string, { action: "upload" | "download" | "same"; localHash?: string; remoteHash?: string }>;
}

type PendingFirstChoice = {
    id: string;
    name: string;
    mode: "new" | "existing";
    mergeAndroid?: boolean;
    emptyAppRoot?: string;
};

type FirstConnectionSnapshot = {
    accountId: string;
    localVaultId: string;
    deviceId: string;
    cloudVaultId: string;
    mode: "new" | "existing";
    registryHash: string;
    rulesHash: string;
    localVersions: string[];
    remoteVersions: string[];
    localStructure?: string[];
    remoteFolders?: string[];
};

type FirstCommitProposal = {
    accountId: string;
    localVaultId: string;
    deviceId: string;
    cloudVaultId: string;
    cloudVaultName: string;
    mode: "create_new";
    registryEtag: string;
    registryVaults: RemoteVault[];
    snapshotDigest: string;
    uploads: string[];
    downloads: string[];
    conflicts: string[];
    unknown: string[];
    expiresAt: number;
    priorBindingArchive?: any;
};

type FirstCommitTicket = Readonly<{ ticketId: string; digest: string; expiresAt: number }>;
type FirstCommitDecision = { status: "BLOCKED" | "REGISTERED" | "RECONCILE_REQUIRED"; reason?: string };
type FileCopyDecision = { status: "BLOCKED" | "FILES_COPIED" | "RECONCILE_REQUIRED"; reason?: string };
type FirstFileEntry = { path: string; size: number; mtime: number; sha256: string };
type FirstFileJournal = {
    ticketId: string;
    digest: string;
    registrationOpId: string;
    accountId: string;
    localVaultId: string;
    cloudVaultId: string;
    deviceId: string;
    registryDigest: string;
    ruleDigest: string;
    files: FirstFileEntry[];
    activeUploadPath?: string | null;
    sourceMtimeByPath?: Record<string, number>;
    completed: Record<string, { id: string; etag: string }>;
    stage: "TRANSFERRING" | "RECONCILE_REQUIRED" | "FILES_COPIED";
    updatedAt: number;
};
type FirstCommitJournal = {
    priorBindingArchive?: any;
    ticketId: string;
    digest: string;
    identity: { accountId: string; localVaultId: string; deviceId: string; cloudVaultId: string };
    stage: "REGISTERING" | "REGISTERED" | "RECONCILE_REQUIRED";
    registryEtag: string;
    cloudVaultName: string;
    registryBefore: RemoteVault[];
    updatedAt: number;
};

type FirstUploadReadScope = {
    signal: AbortSignal;
    assertCurrent: () => void;
    noteActivity: () => void;
};

export class SyncPreviewModal extends Modal {
    private submitted = false;
    private settled = false;
    constructor(
        app: App,
        private preview: SyncPreview,
        private proceed: (reportProgress: (message: string) => void) => Promise<boolean>,
        private finished: () => void,
        private commitEnabled = false,
    ) {
        super(app);
    }

    onOpen(): void {
        this.titleEl.setText(localize(this.preview.creatingCloudVault ? "新建 OneDrive 知识库预览" : "首次同步预览"));
        const androidPreview = this.preview.androidCloudVaultName !== undefined;
        const tipText = this.preview.finishFirstUploadOnly
            ? localize("已核验首次上传的全部文件。确认只完成绑定，本次不传输文件。之后的本地改动会留给下一次同步。")
            : this.preview.resumedFirstUpload
            ? localize("已核验云端已上传文件。请检查待上传清单；确认后只上传尚未提交的文件。")
            : androidPreview
            ? localize("核对上传和下载清单，确认后开始同步。首次同步不会删除文件。")
            : this.commitEnabled
            ? localize("请核对文件范围。点击【确认并开始同步】将正式执行首次双向无损同步；取消或关闭窗口不会对本地或云端做任何修改。")
            : localize("当前属于只读检查阶段，仅供核对文件范围。正式首次同步尚未开放，无法创建、上传、下载或删除任何文件。");
        this.contentEl.createEl("p", { text: localizeMessage(tipText) });
        if (androidPreview) {
            this.contentEl.createEl("p", { text: localize("此手机：{0}", this.app.vault.getName()) });
            this.contentEl.createEl("p", { text: localize("OneDrive：{0}", this.preview.androidCloudVaultName) });
        } else if (this.preview.identity) {
            const identity = this.preview.identity;
            this.contentEl.createEl("p", { text:
                localize("微软账号：") + redactOneDriveText(identity.accountId) });
            this.contentEl.createEl("p", { text:
                localize("本地 Vault ID：") + redactOneDriveText(identity.localVaultId) +
                localize("；云端知识库 ID：") + redactOneDriveText(identity.cloudVaultId) });
        }
        if (this.preview.priorHistoryCount && !androidPreview) this.contentEl.createEl("p", { text:
            localize("本次明确选择保留本地并创建全新独立云副本；旧 ") + this.preview.priorHistoryCount +
            localize(" 项文件基线及完成历史会保留在本机归档。新副本完整核验前旧基线不清空；旧云库不修改，旧许可不能继承。") });
        if (androidPreview) {
            this.renderAndroidFiles();
        } else {
            this.renderDetailedFiles();
        }
        const unresolved =
            this.preview.unknownPaths.length > 0 || this.preview.unexpectedNewTarget;
        if (unresolved) this.contentEl.createEl("p", {
            text: androidPreview
                ? localize("暂时无法同步。请返回后重试。")
                : this.preview.unexpectedNewTarget
                ? localize("拟新建的云端目标已有未登记内容，不能按全新空库提交；请重新检查云端身份。")
                : localize("有文件的内容或版本无法核实。请重新检查后再同步。"),
        });
        const actions = this.contentEl.createDiv({ cls: "smartmd-actions" });
        actions.createEl("button", { text: localize("取消") }).onclick = () => this.close();
        const start = actions.createEl("button", { text: this.preview.finishFirstUploadOnly ? localize("确认并完成绑定")
            : this.commitEnabled ? localize("确认并开始同步") : localize("尚未开放首次同步"), cls: "mod-cta" });
        start.disabled = !this.commitEnabled || unresolved;
        start.onclick = async () => {
            if (this.submitted || !this.commitEnabled || unresolved) return;
            this.submitted = true;
            this.close();
            new Notice(this.preview.finishFirstUploadOnly ? localize("SmartMD: 正在复核已上传文件。")
                : androidPreview ? localize("正在检查文件，请稍等…")
                : localize("SmartMD: 正在复核首次同步计划，复核通过后才会开始传输。"), 5000);
            try {
                const success = await this.proceed((message) => {
                    const plugin = (this.app as any).plugins?.plugins?.smartmd;
                    if (plugin?.updateStatusBar) {
                        plugin.updateStatusBar(message);
                    }
                });
                if (success) {
                    new Notice(this.preview.finishFirstUploadOnly ? localize("SmartMD: 已完成绑定，可以开始下一次同步。")
                        : androidPreview ? localize("首次同步完成")
                        : localize("SmartMD: 首次同步完成！所有笔记已就绪。"), 6000);
                }
            } catch (error: any) {
                new Notice(localize("SmartMD: 同步异常：") + localizeMessage(safeOneDriveError(error)), 8000);
            } finally {
                this.settle();
            }
        };
    }

    private renderAndroidFiles(): void {
        const groups: Array<[string, string[]]> = [
            ["上传到 OneDrive", this.preview.uploadPaths],
            ["下载到这台手机", this.preview.downloadPaths],
            ["这次跳过", this.preview.skippedPaths],
            ["暂时无法检查", this.preview.unknownPaths],
        ];
        for (const [label, paths] of groups) {
            if (!paths.length) continue;
            const section = this.contentEl.createEl("details", { cls: "smartmd-onedrive-preview-file" });
            section.open = paths.length <= 5;
            section.createEl("summary", { text: localize("{0}：{1} 项", localize(label), paths.length) });
            const files = section.createEl("ul");
            for (const path of paths) files.createEl("li", { text: redactOneDriveText(path) });
        }
        if (groups.every(([, paths]) => paths.length === 0) && !this.preview.unexpectedNewTarget) {
            this.contentEl.createEl("p", { text: localize("两边文件一致，可以直接连接。") });
        }
    }

    private renderDetailedFiles(): void {
        const list = this.contentEl.createEl("ul");
        if (!this.preview.finishFirstUploadOnly) list.createEl("li", { text: localize("本地文件：{0}", this.preview.localCount) });
        list.createEl("li", { text: localize("云端文件：{0}", this.preview.remoteCount) });
        list.createEl("li", { text: localize("预计上传：{0}", this.preview.uploadCount) });
        list.createEl("li", { text: localize("预计下载：{0}", this.preview.downloadCount) });
        if (!this.preview.finishFirstUploadOnly) {
            list.createEl("li", { text: localize("同路径内容相同：{0}", this.preview.samePaths.length) });
            list.createEl("li", { text: localize("因大小限制或排除规则跳过：{0}", this.preview.skippedCount) });
            list.createEl("li", { text: localize("内容无法核实：{0}", this.preview.unknownPaths.length) });
        }
        const groups: Array<[string, string[], string]> = this.preview.finishFirstUploadOnly ? [] : [
            [localize("将上传"), this.preview.uploadPaths, localize("本机 → 云端")],
            [localize("将下载"), this.preview.downloadPaths, localize("云端 → 本机")],
            [localize("内容相同"), this.preview.samePaths, localize("不传输（字节相同）")],
            [localize("跳过"), this.preview.skippedPaths, localize("不执行上传、下载或删除")],
            [localize("无法确认（禁止执行）"), this.preview.unknownPaths, localize("暂停，禁止执行")],
        ];
        for (const [label, paths, direction] of groups) {
            this.contentEl.createEl("p", { text: localize("{0}：{1} 项", localizeMessage(label), paths.length) });
            const files = this.contentEl.createEl("ul");
            for (const path of paths) {
                const item = this.preview.items?.[path];
                if (!item) {
                    files.createEl("li", { text: redactOneDriveText(path) });
                    continue;
                }
                const row = files.createEl("li").createEl("details");
                row.addClass("smartmd-onedrive-preview-file");
                row.createEl("summary", { text: redactOneDriveText(path) +
                    localize("；来源：") + item.source + "；" + direction +
                    (item.skipReasons.length ? "；" + item.skipReasons.join("；") : "") });
                if (this.preview.identity) row.createEl("p", { text:
                    localize("目标范围：本地库 ") + redactOneDriveText(this.preview.identity.localVaultId) +
                    localize("；云库 ") + redactOneDriveText(this.preview.identity.cloudVaultId) });
                row.createEl("p", { text: item.local
                    ? localize("本机版本：{0} 字节；mtime {1}；SHA-256：{2}", item.local.size, item.local.mtime, item.local.sha256 || localize("未读取（跳过）"))
                    : localize("本机目标：预览时不存在，未经最终批准不会创建") });
                row.createEl("p", { text: item.remote
                    ? localize("云端 Item ID：") + redactOneDriveText(item.remote.id) +
                        "；ETag：" + redactOneDriveText(item.remote.etag) +
                        localize("；{0} 字节", item.remote.size)
                    : localize("云端目标：预览时不存在，未经最终批准不会创建") });
            }
        }
    }


    onClose(): void {
        this.contentEl.empty();
        if (!this.submitted) this.settle();
    }

    private settle(): void {
        if (this.settled) return;
        this.settled = true;
        this.finished();
    }
}

function androidVaultDisplayName(vault: Pick<RemoteVault, "name" | "id">, vaults: RemoteVault[]): string {
    const label = vault.name || localize("未命名知识库");
    const hasSameName = vaults.some(other => other.id !== vault.id &&
        (other.name || localize("未命名知识库")) === label);
    return hasSameName ? label + " (" + vault.id + ")" : label;
}

class VaultSelectionModal extends Modal {
    private settled = false;

    constructor(
        app: App, private vaults: RemoteVault[],
        private choose: (choice: RemoteVault | "new" | null) => void,
        private readonly eligibleReceiver = false,
        private readonly allowDesktopRegistration = false,
        private readonly allowAndroidMerge = false,
    ) {
        super(app);
    }

    onOpen(): void {
        if (Platform.isAndroidApp) {
            this.renderAndroidChoices();
            return;
        }
        this.titleEl.setText(localize("选择云端知识库"));
        this.contentEl.createEl("p", {
            text: localize("本地与云端知识库按各自 ID 区分，同名不等于同一份。默认暂不同步；登录和选择均不授权传输、创建或删除。"),
        });
        this.contentEl.createEl("button", { text: localize("暂不同步（默认）") }).onclick = () => this.close();
        this.contentEl.createEl("button", {
            text: localize("保留本地，创建新的云端副本"),
            cls: "mod-cta",
        }).onclick = () => this.finish("new");
        for (const vault of this.vaults) {
            const item = this.contentEl.createDiv({ cls: "smartmd-card" });
            item.createEl("h3", { text: vault.name || localize("未命名知识库") });
            item.createEl("p", { text: localize("云端 ID：") + vault.id });
            const receive = item.createEl("button", { text: localize("使用云端，下载成另一份本地知识库") });
            // P1-DEC-02-R1: allow a verified independent, history-free Windows
            // Vault. Selection is still READ-ONLY, not execution consent.
            receive.disabled = !this.eligibleReceiver;
            if (this.eligibleReceiver) receive.onclick = () => this.finish(vault);
            if (this.allowDesktopRegistration) {
                item.createEl("button", { text: localize("高级：仅登记完全一致的内容") }).onclick =
                    () => this.finish({ ...vault, registerIdentical: true });
                item.createEl("p", { text: localize("高级同内容登记另行核对双端全文，仅在全范围一致并最终确认后记录本机基线，不传正文、不删除、不自动同步。") });
            }
            item.createEl("p", {
                text: this.eligibleReceiver
                    ? localize("当前本地知识库已核实独立且无旧文件/同步历史。请先查看下载预览，再明确确认；不会新建另一份云端知识库。")
                    : localize("当前本地知识库未通过完整的独立空库与历史核验；为保护已有资料，暂不开放此操作。"),
            });
        }
        this.contentEl.createEl("button", {
            text: localize("两份都保留，不合并"),
        }).onclick = () => this.close();
    }

    private renderAndroidChoices(): void {
        const canConnect = this.eligibleReceiver || this.allowAndroidMerge;
        this.titleEl.setText(localize("选择要同步的 OneDrive 知识库"));
        if (this.vaults.length > 0) {
            this.contentEl.createEl("p", { text: canConnect
                ? localize("选中电脑正在同步的知识库。下一步会先显示「{0}」的同步预览。", this.app.vault.getName())
                : localize("这个手机知识库暂时无法连接。请先检查它原来的同步设置。") });
        } else {
            this.contentEl.createEl("p", { text: localize("这个微软账号下还没有 SmartMD 云端知识库。") });
        }
        const list = this.contentEl.createDiv({ cls: "smartmd-android-vault-choices" });
        for (const vault of this.vaults) {
            const button = list.createEl("button", { cls: "smartmd-android-vault-choice" });
            button.disabled = !canConnect;
            button.createSpan({ cls: "smartmd-android-vault-name", text: androidVaultDisplayName(vault, this.vaults) });
            if (canConnect) {
                button.createSpan({ cls: "smartmd-android-vault-detail", text: this.eligibleReceiver
                    ? localize("预览下载到当前知识库")
                    : localize("预览双向同步") });
                button.onclick = () => this.finish({ ...vault, mergeAndroid: this.allowAndroidMerge });
            }
        }
        const actions = this.contentEl.createDiv({ cls: "modal-button-container" });
        actions.createEl("button", { text: localize("取消") }).onclick = () => this.close();
        actions.createEl("button", {
            text: localize("在 OneDrive 新建知识库"),
            cls: this.vaults.length === 0 ? "mod-cta" : undefined,
        }).onclick = () => this.finish("new");
    }

    onClose(): void {
        this.contentEl.empty();
        if (!this.settled) this.choose(null);
    }

    private finish(choice: RemoteVault | "new"): void {
        if (this.settled) return;
        this.settled = true;
        this.choose(choice);
        this.close();
    }
}

type ResumedAction = "UPLOAD" | "DOWNLOAD"
    | "DELETE_CLOUD" | "DELETE_VAULT" | "UNCHANGED" | "HOLD";
type ResumedEntry = {
    path: string; action: ResumedAction; reason: string;
    localHash: string; localMtime: number; localSize: number;
    remoteId: string; remoteEtag: string; remoteSize: number; remoteHash: string;
    baselineHash: string; baselineRemoteId: string; baselineRemoteEtag: string;
    requiresSeparateVerification?: boolean;
};
type ResumedSnapshot = {
    accountId: string; vaultId: string; deviceId: string;
    proof: string; confirmationAccountId: string; confirmationVaultId: string;
    localVaultId: string; cloudVaultId: string;
    ruleHash: string; baselineHash: string; pendingHash: string;
    sharedHash: string; entries: ResumedEntry[];
};
export type ResumedOneDriveReview = {
    ticket: object; accountId: string; vaultId: string; deviceId: string;
    localVaultId: string; cloudVaultId: string;
    expiresAt: number; entries: ResumedEntry[];
};

export class OneDriveSyncService {
    private running = false;
    private preparing = false;
    private saveTimer = 0;
    private operationJob: Promise<boolean> | null = null;
    private followUpQueued = false;
    // Invalidates queued automatic work when a user changes to manual/pauses.
    // A running, already authorized round may finish; no new automatic round may start.
    private autoFollowUpGeneration = 0;
    private followUpTimer = 0;
    private retryTimer = 0;
    private retryAt = 0;
    private retryCount = 0;
    private syncableChangeVersion = 0;
    private activeRoundChangeVersion = 0;
    private observedSharedState = "";
    private statusListener: (() => void) | null = null;
    private firstUploadInspection: { controller: AbortController; label: string } | null = null;
    private firstConnectionError = "";
    private firstConnectionProgress = "";
    private firstConnectionNotice: Notice | null = null;
    private sharedSettingsError = "";
    private readonly firstUploadReadTimeoutMs = 30_000;
    private deleteGuard: ((path: string, readOnly?: boolean) => Promise<boolean>) | null = null;
    private ensuredRemoteFolders = new Set<string>();
    private ensuringRemoteFolders = new Map<string, Promise<void>>();
    private remoteFolderIds = new Map<string, string>();
    private approotId: string | null = null;
    private vaultRegistryCache: RemoteVault[] | null = null;
    private registryReadIsTrusted = false;
    // Ephemeral first-connection choice: never write identity settings until final approval.
    private pendingFirstChoice: PendingFirstChoice | null = null;
    private previewVaultId: string | null = null;
    private firstCommitStaged: { ticket: FirstCommitTicket; proposal: FirstCommitProposal; generation: number } | null = null;
    private firstCommitConsents = new WeakSet<object>();
    private firstCommitInFlight: Promise<FirstCommitDecision> | null = null;
    private activeRegistrationTicket: string | null = null;
    private approvedFirstCommitTicketId: string | null = null;
    private activeFirstSyncTicket: string | null = null;
    private activeFirstReceiverResume = false;
    private firstFilePrepared: { ticket: FirstCommitTicket; journal: FirstFileJournal; generation: number } | null = null;
    private firstFileConsents = new WeakSet<object>();
    private firstFileJob: Promise<FileCopyDecision> | null = null;
    private remoteDeltaCache: { vaultId: string; rootId: string; deltaLink: string; nodes: Record<string, any> } | null = null;
    private pendingSharedImport: {
        accountId: string; vaultId: string; deviceId: string;
        localHash: string; remoteHash: string; fullPayloadHash: string;
        baselineHash: string; fields: SharedSmartMDState; expiresAt: number;
    } | null = null;
    private authority = new OperationAuthority();
    private resumedReviewTickets = new WeakMap<object, {
        snapshot: string; expiresAt: number; targetPath?: string;
    }>();
    // Only one durable completed-upload ledger settlement may run at a time.
    // No new preview/approval can observe an uncommitted deletion.
    private uploadLedgerReconciliationInFlight = false;
    private uploadLedgerSettlementInFlight = false;
    private uploadLedgerSettlementGeneration = 0;
    // Single-use, ephemeral approval for settling a fully verified UNKNOWN
    // upload. No old broad run grant or persisted blanket authorization.
    private unknownSettlementInFlight = false;
    private unknownSettlementTickets = new WeakMap<object, {
        path: string; scope: string; ledger: string; entry: string; baseline: string; expiresAt: number;
    }>();
    private interruptedCleanupTickets = new WeakMap<object, {
        path: string; scope: string; record: string; baseline: string;
        evidence: string; expiresAt: number;
    }>();
    private resumedInFlight = false;
    private resumedApprovalGeneration = 0;
    // Never persisted. Exact one-operation capability is separate from
    // historical settings.runAllowed; it cannot start auto or manual sync.
    private resumedOperation: RunTicket | null = null;
    private resumedTransientState: { confirmed: boolean; runAllowed: boolean } | null = null;
    private resumedScope = "";
    private resumedSettings: any | null = null;
    private resumedAuthorizedPath = "";
    private resumedAuthorizedAction: ResumedAction | null = null;

    static quickXorHash(buffer: ArrayBuffer): string {
        const bytes = new Uint8Array(buffer);
        const cellBits = [BigInt(64), BigInt(64), BigInt(32)];
        const cells = [BigInt(0), BigInt(0), BigInt(0)];

        let cellIdx = 0;
        let bitOffset = BigInt(0);
        const eight = BigInt(8);
        const eleven = BigInt(11);
        const mask64 = (BigInt(1) << BigInt(64)) - BigInt(1);
        const mask32 = (BigInt(1) << BigInt(32)) - BigInt(1);
        const mask8 = BigInt(255);

        for (let i = 0; i < bytes.length; i++) {
            const b = BigInt(bytes[i]);
            const bitsInCell = cellBits[cellIdx];
            if (bitOffset <= bitsInCell - eight) {
                cells[cellIdx] ^= (b << bitOffset);
            } else {
                const low = bitsInCell - bitOffset;
                const nextCell = cellIdx === 2 ? 0 : cellIdx + 1;
                cells[cellIdx] ^= (b << bitOffset);
                cells[nextCell] ^= (b >> low);
            }

            bitOffset += eleven;
            while (bitOffset >= cellBits[cellIdx]) {
                bitOffset -= cellBits[cellIdx];
                cellIdx = cellIdx === 2 ? 0 : cellIdx + 1;
            }
        }

        const rgb = new Uint8Array(20);
        const view = new DataView(rgb.buffer);
        view.setBigUint64(0, cells[0] & mask64, true);
        view.setBigUint64(8, cells[1] & mask64, true);
        view.setUint32(16, Number(cells[2] & mask32), true);

        let length = BigInt(bytes.length);
        for (let i = 0; i < 8; i++) {
            rgb[12 + i] ^= Number(length & mask8);
            length >>= eight;
        }

        let binary = "";
        for (let i = 0; i < 20; i++) binary += String.fromCharCode(rgb[i]);
        return typeof btoa === "function" ? btoa(binary) : Buffer.from(binary, "binary").toString("base64");
    }

    constructor(private app: App, private store: SmartMDStore, private auth: MicrosoftAuth) {}

    setDeleteGuard(guard: (path: string, readOnly?: boolean) => Promise<boolean>): void {
        this.deleteGuard = guard;
    }

    setStatusListener(listener: () => void): void {
        this.statusListener = listener;
    }

    setActiveFirstSyncTicket(ticket: string | null): void {
        this.activeFirstSyncTicket = ticket;
    }

    private notifyStatusChange(): void {
        this.statusListener?.();
    }

    settings(): any {
        return (this.store as any).settings.local.oneDrive;
    }

    private async captureResumedOneDriveReview(targetPath?: string): Promise<ResumedSnapshot> {
        if (this.uploadLedgerReconciliationInFlight || this.uploadLedgerSettlementInFlight ||
            this.unknownSettlementInFlight) {
            throw new Error("旧上传日志正在持久结算，完成前不得预览或批准新上传");
        }
        const s = this.settings();
        const localSettings: any = (this.store as any).settings.local;
        const accountId = String(localSettings.microsoftAccountId || "");
        if (this.running || this.preparing || this.operationJob ||
            !accountId || !s.confirmed || !s.firstSyncCompleted || s.runAllowed ||
            s.confirmationAccountId !== accountId ||
            s.confirmationVaultId !== s.vaultId ||
            (this.store as any).installationProofIsValid?.() !== true) {
            throw new Error("旧许可已失效：须为当前账号、云库及物理Vault另作只读检查；未确证的身份不可签发新操作");
        }
        for (const rule of s.ignorePatterns || []) ignoreMatcher(rule);
        const existingLedger = s.pendingUploadOperations;
        if (existingLedger && (typeof existingLedger !== "object" || Array.isArray(existingLedger))) {
            throw new Error("历史上传操作记录格式无效，只读恢复不得解除旧锁");
        }
        // A paused connection may still have a completed upload receipt.
        // Use the same release guard as sync. Unresolved uploads remain held.
        if (Object.values(existingLedger || {}).some((entry: any) => entry?.phase === "CLEANUP_VERIFIED")) {
            // Set the guard before the first await. A second settings panel or
            // a pre-issued ticket cannot race the baseline or release save.
            this.uploadLedgerReconciliationInFlight = true;
            try {
                await (this.store as any).save();
                if (this.settings() !== s ||
                    String((this.store as any).settings.local.microsoftAccountId || "") !== accountId ||
                    !this.resumedHistoricalIdentityTrusted()) {
                    throw new Error("上传完成日志核查时身份或物理安装证明变化，拒绝解锁");
                }
                await this.clearVerifiedUploadLedger(s.fileStates || {});
            } finally {
                this.uploadLedgerReconciliationInFlight = false;
            }
        }
        const settlementGeneration = this.uploadLedgerSettlementGeneration;
        const scope = JSON.stringify([
            accountId, s.vaultId, s.deviceId, s.installationProofId,
            s.confirmationAccountId, s.confirmationVaultId,
            s.localVaultId || "", s.cloudVaultId || "",
        ]);
        const locals = this.app.vault.getFiles();
        const remotes = await this.remoteFiles(true, true);
        if (!targetPath && locals.length + remotes.length > 4000) {
            throw new Error("当前清单超出限定单次安全核查容量，不得自动恢复旧同步计划");
        }
        const localMap = new Map(locals.map((file) => [file.path, file]));
        const remoteMap = new Map(remotes.map((item) => [item.path, item]));
        const previous = s.fileStates || {};
        const uploadLedger = s.pendingUploadOperations || {};
        if (!uploadLedger || typeof uploadLedger !== "object" || Array.isArray(uploadLedger)) {
            throw new Error("历史上传原操作记录不可信，停止生成新执行票据");
        }
        const orphanStaging = remotes.filter((item) =>
            item.path.startsWith("_suite_staging/"));
        const matchesOrphan = (mother: string): boolean => {
            const safeName = mother.replace(/[\\/:*?"<>|]/g, "_");
            return orphanStaging.some((item) =>
                item.path.split("/").slice(2).join("/") === safeName);
        };
        const paths = [...new Set([
            ...localMap.keys(), ...remoteMap.keys(), ...Object.keys(previous),
            ...Object.keys(s.pendingDeletions || {}),
            ...Object.keys(uploadLedger),
        ])].sort();
        const entries: ResumedEntry[] = [];
        // After a cold restart, verify changed remote bytes before approval.
        // Android can verify one configured file in a separate preview when
        // the combined size of changed files exhausts the full preview budget.
        const permittedBytes = this.maxBytes();
        let inspectionBudget = Platform.isAndroidApp ? permittedBytes : 16_000_000;
        for (const path of targetPath ? paths.filter(item => item === targetPath) : paths) {
            if (!path || path.startsWith("/") || path.includes("\\") ||
                path.split("/").some((part) => !part || part === "." || part === "..")) {
                throw new Error("同步文件相对路径不可信，停止批准新范围");
            }
            const localFile = localMap.get(path);
            const remote = remoteMap.get(path);
            const old: any = previous[path];
            const small = (!localFile || localFile.stat.size <= permittedBytes)
                && (!remote || remote.size <= permittedBytes);
            const unknownPriorWrite = Boolean(uploadLedger[path]) || matchesOrphan(path);
            const blocked = this.isIgnored(path) || /^_suite_staging\//i.test(path)
                || /^\.suite-(?:staging|recovery)\//i.test(path) || !small
                || unknownPriorWrite
                || Boolean(s.pendingDeletions?.[path]);
            const localHash = localFile && small
                ? await this.sha256Bytes(await this.app.vault.readBinary(localFile)) : "";
            let action: ResumedAction = "HOLD";
            let reason = "需核对同步范围与文件版本，本轮暂不写入";
            if (blocked) reason = unknownPriorWrite
                ? "原上传操作/暂存副本未按操作ID、阶段、源SHA与目标版本完成只读对账，母文件不得重发"
                : "忽略/超限/删除保护：不得继承执行许可";
            else if (localFile && remote && old?.localHash && old.remoteId === remote.id) {
                const changedLocal = old.localHash !== localHash;
                const changedRemote = old.remoteEtag !== remote.etag;
                if (!changedLocal && !changedRemote) { action = "UNCHANGED"; reason = "可信基线未变化"; }
                else if (changedLocal && changedRemote) {
                    action = await this.oneDriveVersionWins(path, localFile.stat.mtime, remote)
                        ? "DOWNLOAD" : "UPLOAD";
                    reason = "采用修改时间较新的版本，时间相同以 OneDrive 为准";
                }
                else { action = changedLocal ? "UPLOAD" : "DOWNLOAD"; reason = "仅一端变化，允许本条独立预览"; }
            } else if (!localFile && remote && !old) {
                action = "DOWNLOAD"; reason = "新云端文件，用户按当前ItemID/ETag逐项批准下载";
            } else if (localFile && !remote && !old) {
                action = "UPLOAD"; reason = "新本地文件，用户按当前实际SHA逐项批准上传";
            } else if (localFile && remote && !old) {
                action = await this.oneDriveVersionWins(path, localFile.stat.mtime, remote)
                    ? "DOWNLOAD" : "UPLOAD";
                reason = "采用修改时间较新的版本，时间相同以 OneDrive 为准";
            }
            const deletion = s.pendingDeletions?.[path];
            const safeProtectedPath = !this.isIgnored(path) && small
                && !/^_suite_staging\//i.test(path)
                && !/^\.suite-(?:staging|recovery)\//i.test(path);
            if (safeProtectedPath && deletion && old
                && deletion.accountId === accountId &&
                deletion.vaultId === s.vaultId && deletion.deviceId === s.deviceId) {
                if (deletion.direction === "cloud" && !localFile && remote
                    && remote.id === deletion.remoteId
                    && remote.etag === deletion.remoteEtag
                    && old.remoteId === remote.id && old.remoteEtag === remote.etag) {
                    action = "DELETE_CLOUD";
                    reason = "本地原稿缺失但云端存在；只允许明确批准此云版本删除，必须先物理冷备";
                } else if (deletion.direction === "vault" && localFile && !remote
                    && localHash && localHash === deletion.localHash
                    && (!Platform.isAndroidApp || old.localHash === localHash)
                    && localFile.stat.mtime === deletion.localMtime
                    && localFile.stat.size === deletion.localSize
                    && old.remoteId === deletion.remoteId) {
                    // Re-approval preview must never mutate an external binding
                    // or create a recovery archive before explicit consent.
                    const externalProtected = this.deleteGuard
                        ? await this.deleteGuard(path, true) : false;
                    if (!externalProtected) {
                        action = "DELETE_VAULT";
                        reason = "云端文件已缺失，需核对本地版后确认删除";
                    }
                }
            }
            let remoteHash = "";
            let requiresSeparateVerification: boolean | undefined;
            if ((action === "DOWNLOAD" || action === "DELETE_CLOUD"
                || (action === "UPLOAD" && remote)) && remote) {
                if (remote.size > inspectionBudget) {
                    action = "HOLD"; reason = "只读下载校验预算不足，禁止未验版本写入";
                    if (Platform.isAndroidApp && !targetPath) requiresSeparateVerification = true;
                } else {
                    const data = await this.remoteBytes(remote);
                    if (data.byteLength !== remote.size) {
                        throw new Error("云端真实正文与清单大小不一致，不得批准");
                    }
                    remoteHash = await this.sha256Bytes(data);
                    inspectionBudget -= data.byteLength;
                    if (localFile && remoteHash === localHash) {
                        action = "UNCHANGED";
                        reason = "双端内容一致，无需传输";
                    } else if (localFile && old?.localHash === remoteHash && old.localHash !== localHash) {
                        action = "UPLOAD";
                        reason = "仅本机正文有变化";
                    }
                }
            }
            entries.push({
                path, action, reason, localHash, localMtime: localFile?.stat.mtime || 0,
                localSize: localFile?.stat.size || 0, remoteId: remote?.id || "",
                remoteEtag: remote?.etag || "", remoteSize: remote?.size || 0,
                remoteHash, baselineHash: old?.localHash || "",
                baselineRemoteId: old?.remoteId || "",
                baselineRemoteEtag: old?.remoteEtag || "",
                requiresSeparateVerification,
            });
        }
        if (scope !== JSON.stringify([
            String((this.store as any).settings.local.microsoftAccountId || ""),
            this.settings().vaultId, this.settings().deviceId, this.settings().installationProofId,
            this.settings().confirmationAccountId, this.settings().confirmationVaultId,
            this.settings().localVaultId || "", this.settings().cloudVaultId || "",
        ]) || this.settings() !== s ||
            (this.store as any).installationProofIsValid?.() !== true || s.runAllowed) {
            throw new Error("预览期间账号、知识库、设备或物理安装证明已变化，禁止旧范围");
        }
        if (this.uploadLedgerReconciliationInFlight || this.uploadLedgerSettlementInFlight ||
            this.uploadLedgerSettlementGeneration !== settlementGeneration) {
            throw new Error("预览期间旧上传日志结算变化，必须重新只读检查");
        }
        return {
            accountId, vaultId: s.vaultId, deviceId: s.deviceId,
            proof: String(s.installationProofId || ""),
            confirmationAccountId: s.confirmationAccountId,
            confirmationVaultId: s.confirmationVaultId,
            localVaultId: String(s.localVaultId || ""), cloudVaultId: String(s.cloudVaultId || ""),
            ruleHash: JSON.stringify([s.ignorePatterns || [], s.maxFileSizeMb]),
            baselineHash: JSON.stringify(s.fileStates || {}),
            pendingHash: JSON.stringify([s.pendingDeletions || {},
                s.pendingUploadOperations || {}, s.firstCommitJournal || null,
                s.firstFileJournal || null]),
            sharedHash: this.sharedStateFingerprint(), entries,
        };
    }

    async prepareResumedOneDriveReview(targetPath?: string): Promise<ResumedOneDriveReview> {
        const generation = this.resumedApprovalGeneration;
        if (targetPath && !Platform.isAndroidApp) throw new Error("当前平台不支持指定单条重新核验");
        await this.verifyCompletedIdenticalRegistration();
        if (this.resumedInFlight || this.uploadLedgerReconciliationInFlight ||
            this.uploadLedgerSettlementInFlight || this.unknownSettlementInFlight) {
            throw new Error("上一轮操作仍在执行或上传日志正在持久结算，禁止签发新预览");
        }
        const snapshot = await this.captureResumedOneDriveReview(targetPath);
        if (generation !== this.resumedApprovalGeneration || (targetPath && snapshot.entries.length !== 1)) {
            throw new Error("重新核验的执行范围已变化，请重新打开预览");
        }
        const ticket = Object.freeze({});
        const expiresAt = Date.now() + 5 * 60_000;
        this.resumedReviewTickets.set(ticket, { snapshot: JSON.stringify(snapshot), expiresAt, targetPath });
        return {
            ticket, accountId: snapshot.accountId, vaultId: snapshot.vaultId,
            localVaultId: snapshot.localVaultId, cloudVaultId: snapshot.cloudVaultId,
            deviceId: snapshot.deviceId, expiresAt,
            entries: snapshot.entries.map((entry) => ({ ...entry })),
        };
    }

    /**
     * Read-only, fail-closed evidence for a previously UNKNOWN upload.
     *
     * The staged object can be the only surviving proof of the original
     * operation. Matching target bytes are NOT permission to remove the
     * staging copy, release the ledger, or promote fileStates: those require
     * a separately specified and tested recovery transaction.
     */
    canAuditUnknownUpload(record: { phase: string; sourceMtime?: number }): boolean {
        return record.phase === "RECONCILE_REQUIRED" ||
            (record.phase === "FINAL_PUT_ISSUED" && Number.isFinite(record.sourceMtime));
    }

    async auditPendingUnknownUploadsReadOnly(): Promise<{
        accountId: string; vaultId: string; readOnly: true; canResume: false;
        entries: Array<{
            path: string; status: "VERIFIED_BUT_LOCKED" | "UNVERIFIED_HOLD";
            reason: string; remoteId: string; remoteEtag: string;
            stagingId: string; stagingEtag: string; sourceSha256: string;
        }>;
    }> {
        if (this.running || this.preparing || this.operationJob || this.resumedInFlight ||
            this.uploadLedgerReconciliationInFlight || this.uploadLedgerSettlementInFlight ||
            !this.resumedHistoricalIdentityTrusted()) {
            throw new Error("旧上传只读对账要求当前物理Vault身份可信、全部旧执行许可撤销且没有并发操作");
        }
        const settings = this.settings();
        const accountId = String((this.store as any).settings.local.microsoftAccountId || "");
        const scope = this.resumedContextKey();
        const ledger = settings.pendingUploadOperations || {};
        if (typeof ledger !== "object" || Array.isArray(ledger)) {
            throw new Error("上传历史记录格式不可信，不能推断原操作完成");
        }
        const frozen = JSON.stringify(ledger);
        const keys = Object.keys(ledger);
        if (keys.length > 100) throw new Error("未决上传数量超出独立只读审核容量");
        const remotes = await this.remoteFiles(true, true);
        if (remotes.length > 4000) throw new Error("云端清单超出完整只读审查范围");
        const entries: Array<{
            path: string; status: "VERIFIED_BUT_LOCKED" | "UNVERIFIED_HOLD";
            reason: string; remoteId: string; remoteEtag: string;
            stagingId: string; stagingEtag: string; sourceSha256: string;
        }> = [];
        let remainingBytes = Platform.isAndroidApp ? 3 * this.maxBytes() : 16_000_000;
        for (const path of keys.sort()) {
            const item: any = ledger[path];
            const result = {
                path, status: "UNVERIFIED_HOLD" as "VERIFIED_BUT_LOCKED" | "UNVERIFIED_HOLD",
                reason: "原上传操作不能完整核实，仍需保留原始锁与所有云端暂存副本",
                remoteId: "", remoteEtag: "", stagingId: "", stagingEtag: "",
                sourceSha256: "",
            };
            entries.push(result);
            const legalPath = Boolean(path && !path.startsWith("/") && !path.includes("\\")
                && path.split("/").every((part) => part && part !== "." && part !== "..")
                && !path.startsWith("_suite_staging/") && !this.isIgnored(path));
            const expectedStage = item && typeof item.operationId === "string"
                ? "_suite_staging/" + item.operationId.replace(/[\\/:*?"<>|.]/g, "_")
                    + "/" + path.replace(/[\\/:*?"<>|]/g, "_")
                : "";
            const sourceSha = String(item?.sourceSha256 || "").toLowerCase();
            const bytes = Number(item?.sourceBytes);
            if (!legalPath || !item || typeof item !== "object" ||
                !this.canAuditUnknownUpload(item) ||
                item.direction !== "upload" || item.motherPath !== path ||
                item.accountId !== accountId || item.vaultId !== settings.vaultId ||
                item.deviceId !== settings.deviceId ||
                !item.operationId || !expectedStage ||
                item.stagingPath !== expectedStage ||
                !item.stageId || !item.stageEtag ||
                !/^[0-9a-f]{64}$/.test(sourceSha) ||
                !Number.isSafeInteger(bytes) || bytes < 0 ||
                bytes > this.maxResumedFileBytes() ||
                3 * bytes > remainingBytes) continue;
            const local = this.app.vault.getAbstractFileByPath(path);
            const mother = remotes.filter((remote) => remote.path === path);
            const staging = remotes.filter((remote) => remote.path === expectedStage);
            if (!(local instanceof TFile) || mother.length !== 1 || staging.length !== 1 ||
                local.stat.size !== bytes ||
                mother[0].size !== bytes || staging[0].size !== bytes ||
                !mother[0].id || !mother[0].etag ||
                staging[0].id !== item.stageId ||
                staging[0].etag !== item.stageEtag ||
                (item.finalId && mother[0].id !== item.finalId) ||
                (item.finalEtag && mother[0].etag !== item.finalEtag) ||
                (item.expectedRemoteId && mother[0].id !== item.expectedRemoteId) ||
                (item.expectedRemoteEtag && mother[0].etag === item.expectedRemoteEtag)) continue;
            if (item.sourceMtime !== undefined && mother[0].mtime !== item.sourceMtime) {
                result.reason = localize("OneDrive 尚未确认文件原修改时间：{0}", redactOneDriveText(path));
                continue;
            }
            // Unknown remote writes must never be treated as a successful
            // known baseline merely because the filename happens to exist.
            // Charge the full potential read budget before any awaited I/O,
            // even if one of the three reads subsequently fails.
            remainingBytes -= bytes * 3;
            try {
                const localBytes = await this.app.vault.readBinary(local);
                if (localBytes.byteLength !== bytes ||
                    await this.sha256Bytes(localBytes) !== sourceSha) continue;
                const verify = async (file: RemoteFile) => {
                    const data = await this.remoteBytes(file);
                    return data.byteLength === bytes && await this.sha256Bytes(data) === sourceSha;
                };
                if (!await verify(staging[0]) || !await verify(mother[0])) continue;
                result.status = "VERIFIED_BUT_LOCKED";
                result.reason = "原操作暂存和目标正文均逐字节核实；没有删除暂存、重试上传或解除历史锁的授权";
                result.sourceSha256 = sourceSha;
                result.remoteId = mother[0].id;
                result.remoteEtag = mother[0].etag;
                result.stagingId = staging[0].id;
                result.stagingEtag = staging[0].etag;
            } catch {
                // Timeout/401/429/content mismatch is unverified, not absent.
                result.reason = "云端正文无法完整只读核对，保留UNKNOWN历史与全部原稿";
            }
        }
        if (this.settings() !== settings || this.resumedContextKey() !== scope ||
            JSON.stringify(settings.pendingUploadOperations || {}) !== frozen ||
            this.running || this.preparing || this.operationJob || this.resumedInFlight ||
            this.uploadLedgerReconciliationInFlight || this.uploadLedgerSettlementInFlight ||
            !this.resumedHistoricalIdentityTrusted()) {
            throw new Error("只读对账期间身份、上传历史或设备状态变化，结果无效");
        }
        return { accountId, vaultId: settings.vaultId, readOnly: true,
            canResume: false, entries };
    }

    /**
     * Single-file permission: only a fully matched original UNKNOWN operation
     * may enter settlement. The ticket is in memory and expires; showing the
     * audit alone never grants this capability.
     */
    async prepareUnknownUploadSettlement(path: string): Promise<{
        ticket: object; path: string; accountId: string; vaultId: string;
        remoteId: string; remoteEtag: string; stagingId: string;
        stagingEtag: string; sourceSha256: string; expiresAt: number;
    }> {
        if (this.unknownSettlementInFlight || this.running || this.preparing ||
            this.operationJob || this.resumedInFlight ||
            this.uploadLedgerReconciliationInFlight || this.uploadLedgerSettlementInFlight) {
            throw new Error("上传恢复与其他操作并发，须重新只读审查");
        }
        const audit = await this.auditPendingUnknownUploadsReadOnly();
        const entry = audit.entries.find((item) => item.path === path);
        const settings = this.settings();
        const record = settings.pendingUploadOperations?.[path];
        const baseline = JSON.stringify(settings.fileStates?.[path] || null);
        if (!entry || entry.status !== "VERIFIED_BUT_LOCKED" || !record ||
            !this.canAuditUnknownUpload(record) ||
            (record.sourceBaseline !== undefined ? record.sourceBaseline !== baseline : baseline !== "null")) {
            throw new Error("原操作的两处云端副本及本机SHA尚未全部核实，不允许结算");
        }
        const ticket = Object.freeze({});
        const expiresAt = Date.now() + 5 * 60_000;
        this.unknownSettlementTickets.set(ticket, {
            path, scope: this.resumedContextKey(),
            ledger: JSON.stringify(record), entry: JSON.stringify(entry),
            baseline, expiresAt,
        });
        return { ticket, path, accountId: audit.accountId, vaultId: audit.vaultId,
            remoteId: entry.remoteId, remoteEtag: entry.remoteEtag,
            stagingId: entry.stagingId, stagingEtag: entry.stagingEtag,
            sourceSha256: entry.sourceSha256, expiresAt };
    }

    /**
     * One-use explicit approval for a single previously UNKNOWN upload.
     * No upload is ever retried. Persist the verified target baseline AND the
     * historical ledger together, with the ledger still blocking new writes.
     * A conditional stage-only cleanup then uses the existing durable
     * CLEANUP_VERIFIED release mechanism. Any failure leaves a HOLD.
     */
    async confirmUnknownUploadSettlement(ticket: object, approved = false): Promise<boolean> {
        const issued = ticket && typeof ticket === "object"
            ? this.unknownSettlementTickets.get(ticket) : undefined;
        if (ticket && typeof ticket === "object") this.unknownSettlementTickets.delete(ticket);
        if (!approved) return false;
        if (!issued || issued.expiresAt <= Date.now() ||
            this.unknownSettlementInFlight || this.running || this.preparing ||
            this.operationJob || this.resumedInFlight ||
            this.uploadLedgerReconciliationInFlight || this.uploadLedgerSettlementInFlight) {
            throw new Error("UNKNOWN结算确认已经失效或有并发操作，须重新获取只读票据");
        }
        this.unknownSettlementInFlight = true;
        let durableTransitionStarted = false;
        try {
            const s = this.settings();
            const records = s.pendingUploadOperations;
            const record: any = records?.[issued.path];
            if (!this.resumedHistoricalIdentityTrusted() ||
                issued.scope !== this.resumedContextKey() ||
                !record || JSON.stringify(record) !== issued.ledger ||
                JSON.stringify(s.fileStates?.[issued.path] || null) !== issued.baseline) {
                throw new Error("旧上传原操作或当前物理身份变化，不能继续确认");
            }
            const audit = await this.auditPendingUnknownUploadsReadOnly();
            const verified = audit.entries.find((item) => item.path === issued.path);
            if (!verified || verified.status !== "VERIFIED_BUT_LOCKED" ||
                JSON.stringify(verified) !== issued.entry ||
                issued.expiresAt <= Date.now() ||
                issued.scope !== this.resumedContextKey() ||
                JSON.stringify(records[issued.path]) !== issued.ledger ||
                JSON.stringify(s.fileStates?.[issued.path] || null) !== issued.baseline) {
                throw new Error("确认时本地/远端正文或版本已变化，继续保留UNKNOWN锁");
            }
            const file = this.app.vault.getAbstractFileByPath(issued.path);
            if (!(file instanceof TFile) ||
                await this.sha256Bytes(await this.app.vault.readBinary(file)) !== verified.sourceSha256 ||
                this.settings() !== s || !this.resumedHistoricalIdentityTrusted() ||
                JSON.stringify(s.fileStates?.[issued.path] || null) !== issued.baseline) {
                throw new Error("原始本机文件或设备证明在结算前变化，保持UNKNOWN锁");
            }

            // This fence is set before the first durable mutation: no ordinary
            // preview or upload can interleave with an unfinished settlement.
            this.uploadLedgerReconciliationInFlight = true;
            this.resumedApprovalGeneration++;
            const beforeStates = s.fileStates || {};
            const priorRecord = { ...record };
            const remoteState = {
                localMtime: file.stat.mtime, localSize: file.stat.size,
                localHash: verified.sourceSha256,
                remoteId: verified.remoteId, remoteEtag: verified.remoteEtag,
            };
            // A crash after this atomic save leaves the verified baseline AND
            // the ledger visible. Until stage cleanup, ordinary sync is HOLD.
            s.fileStates = { ...beforeStates, [issued.path]: remoteState };
            records[issued.path] = {
                ...record, phase: "FINAL_ACKNOWLEDGED",
                finalId: verified.remoteId, finalEtag: verified.remoteEtag,
                updatedAt: Date.now(),
            };
            s.status = "paused";
            s.runAllowed = false;
            s.enabled = false;
            s.lastError = "UNKNOWN原稿与暂存已核实，可信基线已保存；暂存清理完成前仍锁定本文件";
            try {
                await (this.store as any).save();
            } catch (error) {
                if (this.settings() === s && s.pendingUploadOperations === records) {
                    s.fileStates = beforeStates;
                    records[issued.path] = priorRecord;
                }
                throw error;
            }
            durableTransitionStarted = true;
            await this.setUploadPhase(issued.path, "CLEANUP_ISSUED");

            // Revalidate both independent cloud bytes, exact IDs and ETags
            // after the durable checkpoint, before the only cloud mutation.
            const remotes = await this.remoteFiles(true, true);
            const original = remotes.find((item) => item.path === issued.path);
            const stage = remotes.find((item) => item.id === verified.stagingId);
            if (!original || !stage ||
                original.id !== verified.remoteId ||
                original.etag !== verified.remoteEtag ||
                (record.sourceMtime !== undefined && original.mtime !== record.sourceMtime) ||
                stage.etag !== verified.stagingEtag ||
                stage.path !== record.stagingPath ||
                await this.sha256Bytes(await this.remoteBytes(original)) !== verified.sourceSha256 ||
                await this.sha256Bytes(await this.remoteBytes(stage)) !== verified.sourceSha256 ||
                await this.sha256Bytes(await this.app.vault.readBinary(file)) !== verified.sourceSha256) {
                throw new Error("暂存清理前远端目标/暂存或本机字节发生变化；保留持久UNKNOWN操作锁");
            }
            // If-Match on exactly the original staging Item ID. Do not delete
            // the final target, previous UNKNOWN, or any unrelated test file.
            await this.deleteRemote(stage, undefined, () => {
                if (!this.unknownSettlementInFlight ||
                    !this.uploadLedgerReconciliationInFlight ||
                    this.settings() !== s || records[issued.path]?.phase !== "CLEANUP_ISSUED" ||
                    records[issued.path]?.operationId !== priorRecord.operationId ||
                    !this.resumedHistoricalIdentityTrusted() ||
                    s.fileStates?.[issued.path]?.remoteId !== verified.remoteId ||
                    s.fileStates?.[issued.path]?.remoteEtag !== verified.remoteEtag) {
                    throw new Error("条件清理派发前原操作身份或可信版本失效，拒绝云端DELETE");
                }
            });
            await this.setUploadPhase(issued.path, "CLEANUP_VERIFIED");
            await this.clearVerifiedUploadLedger(s.fileStates);
            if (records[issued.path]) {
                throw new Error("暂存清理已完成但本地上传日志尚未安全结算，继续保持暂停");
            }
            s.lastError = "已按当前文件明确授权核实并结算旧上传；其他文件仍须独立批准";
            s.status = "paused";
            s.runAllowed = false;
            s.enabled = false;
            this.notifyStatusChange();
            return true;
        } catch (error) {
            // Never turn a post-commit exception into another upload.
            // If a durable checkpoint exists, leave the ledger and verified
            // baseline in place for subsequent read-only crash recovery.
            if (durableTransitionStarted) {
                const s = this.settings();
                s.status = "paused";
                s.runAllowed = false;
                s.enabled = false;
            }
            throw error;
        } finally {
            this.uploadLedgerReconciliationInFlight = false;
            this.unknownSettlementInFlight = false;
            this.notifyStatusChange();
        }
    }

    // After a verified final target was durably saved, the staging DELETE
    // acknowledgement may be lost. This is a separate recovery state from
    // RECONCILE_REQUIRED; neither a 404 nor a filename is sufficient alone.
    private async inspectInterruptedUploadCleanup(path: string, operation?: RunTicket): Promise<{
        path: string; sourceSha256: string; remoteId: string; remoteEtag: string;
        stageId: string; stageEtag: string; stagePath: string;
        stageState: "PRESENT" | "ABSENT_CONFIRMED_BY_GET";
    }> {
        const s = this.settings();
        const record: any = s.pendingUploadOperations?.[path];
        const baseline: any = s.fileStates?.[path];
        const account = String((this.store as any).settings.local.microsoftAccountId || "");
        if (!(operation ? this.isCurrent(operation) : this.resumedHistoricalIdentityTrusted()) ||
            !record || !baseline ||
            !["FINAL_ACKNOWLEDGED", "CLEANUP_ISSUED", "CLEANUP_VERIFIED"].includes(record.phase) ||
            record.motherPath !== path || record.direction !== "upload" ||
            record.accountId !== account || record.vaultId !== s.vaultId ||
            record.deviceId !== s.deviceId ||
            typeof record.operationId !== "string" || !record.operationId ||
            !record.stageId || !record.stageEtag || !record.finalId || !record.finalEtag ||
            !/^[0-9a-f]{64}$/i.test(String(record.sourceSha256 || "")) ||
            baseline.remoteId !== record.finalId ||
            baseline.remoteEtag !== record.finalEtag ||
            baseline.localHash !== record.sourceSha256 ||
            !path || path.startsWith("/") || path.includes("\\") ||
            path.split("/").some((part) => !part || part === "." || part === "..") ||
            this.isIgnored(path)) {
            throw new Error("结算中断的原操作、最终可信版本或物理身份不完整，继续锁定");
        }
        const expectedStage = "_suite_staging/" +
            record.operationId.replace(/[\\/:*?"<>|.]/g, "_") + "/" +
            path.replace(/[\\/:*?"<>|]/g, "_");
        const length = Number(record.sourceBytes);
        if (record.stagingPath !== expectedStage ||
            !Number.isSafeInteger(length) || length < 0 ||
            length > (operation ? this.maxBytes() : this.maxResumedFileBytes())) {
            throw new Error("原操作暂存路径或源字节大小不可信，拒绝恢复");
        }
        const local = this.app.vault.getAbstractFileByPath(path);
        if (!operation && (!(local instanceof TFile) || local.stat.size !== length ||
            await this.sha256Bytes(await this.app.vault.readBinary(local)) !==
                String(record.sourceSha256).toLowerCase())) {
            throw new Error("结算中断后本机原文已改变，必须保留历史锁");
        }
        const remoteList = operation
            ? (await Promise.all([this.readRemoteFileMetadata(path), this.readRemoteFileMetadata(expectedStage)]))
                .filter((item): item is RemoteFile => item !== null)
            : await this.remoteFiles(true, true);
        if (!operation && remoteList.length > 4000) throw new Error("远端清单超限，拒绝恢复");
        const mother = remoteList.filter((item) => item.path === path);
        const stage = remoteList.filter((item) => item.path === expectedStage);
        if (mother.length !== 1 || stage.length > 1 ||
            mother[0].id !== record.finalId ||
            mother[0].etag !== record.finalEtag ||
            mother[0].size !== length ||
            (record.sourceMtime !== undefined && mother[0].mtime !== (record.verifiedRemoteMtime ?? record.sourceMtime)) ||
            await this.sha256Bytes(await this.remoteBytes(mother[0])) !==
                String(record.sourceSha256).toLowerCase()) {
            throw new Error("结算中断后云端最终原稿身份或字节不匹配");
        }
        let stageState: "PRESENT" | "ABSENT_CONFIRMED_BY_GET";
        if (stage.length === 1) {
            if (record.phase === "CLEANUP_VERIFIED" ||
                stage[0].id !== record.stageId ||
                stage[0].id === mother[0].id ||
                stage[0].etag !== record.stageEtag ||
                stage[0].size !== length ||
                await this.sha256Bytes(await this.remoteBytes(stage[0])) !==
                    String(record.sourceSha256).toLowerCase()) {
                throw new Error("原暂存版本改变或完成日志不可信，禁止重复DELETE");
            }
            stageState = "PRESENT";
        } else {
            // Absence in a listing is not a success acknowledgement. Require
            // an independent direct GET against the historical staging Item ID.
            const response = await this.fetchGraph(
                "/drive/items/" + encodeURIComponent(String(record.stageId)),
                { method: "GET" }, operation,
            );
            if (response.status !== 404) {
                throw Object.assign(new Error("原暂存Item ID尚未由微软GET确认为不存在，保持历史锁"),
                    { status: response.status, retryAfterMs: graphRetryAfterMs(response.headers.get("retry-after")) });
            }
            stageState = "ABSENT_CONFIRMED_BY_GET";
        }
        return {
            path, sourceSha256: String(record.sourceSha256).toLowerCase(),
            remoteId: mother[0].id, remoteEtag: mother[0].etag,
            stageId: String(record.stageId), stageEtag: String(record.stageEtag),
            stagePath: expectedStage, stageState,
        };
    }

    async prepareInterruptedUploadCleanup(path: string): Promise<{
        ticket: object; path: string; stageState: "PRESENT" | "ABSENT_CONFIRMED_BY_GET";
        accountId: string; vaultId: string; remoteId: string;
        remoteEtag: string; stageId: string; stageEtag: string;
        sourceSha256: string;
    }> {
        if (this.running || this.preparing || this.operationJob ||
            this.resumedInFlight || this.unknownSettlementInFlight ||
            this.uploadLedgerReconciliationInFlight || this.uploadLedgerSettlementInFlight) {
            throw new Error("结算中断复核禁止与其他操作并发");
        }
        const s = this.settings();
        const scope = this.resumedContextKey();
        const record = JSON.stringify(s.pendingUploadOperations?.[path] || null);
        const baseline = JSON.stringify(s.fileStates?.[path] || null);
        const evidence = await this.inspectInterruptedUploadCleanup(path);
        if (s !== this.settings() || scope !== this.resumedContextKey() ||
            record !== JSON.stringify(s.pendingUploadOperations?.[path] || null) ||
            baseline !== JSON.stringify(s.fileStates?.[path] || null)) {
            throw new Error("只读审查期间旧操作或物理范围改变，不能签发结算票据");
        }
        const ticket = Object.freeze({});
        this.interruptedCleanupTickets.set(ticket, {
            path, scope, record, baseline,
            evidence: JSON.stringify(evidence), expiresAt: Date.now() + 300_000,
        });
        return { ticket, accountId: String((this.store as any).settings.local.microsoftAccountId || ""),
            vaultId: s.vaultId, ...evidence };
    }

    async confirmInterruptedUploadCleanup(ticket: object, approved = false): Promise<boolean> {
        const issued = ticket && typeof ticket === "object"
            ? this.interruptedCleanupTickets.get(ticket) : undefined;
        if (ticket && typeof ticket === "object") this.interruptedCleanupTickets.delete(ticket);
        if (!approved) return false;
        if (!issued || issued.expiresAt <= Date.now() ||
            this.running || this.preparing || this.operationJob ||
            this.resumedInFlight || this.unknownSettlementInFlight ||
            this.uploadLedgerReconciliationInFlight || this.uploadLedgerSettlementInFlight) {
            throw new Error("结算中断票据不存在、失效或发生并发操作");
        }
        this.unknownSettlementInFlight = true;
        this.uploadLedgerReconciliationInFlight = true;
        try {
            const s = this.settings();
            const ledger = s.pendingUploadOperations;
            const record: any = ledger?.[issued.path];
            if (issued.scope !== this.resumedContextKey() ||
                issued.record !== JSON.stringify(record || null) ||
                issued.baseline !== JSON.stringify(s.fileStates?.[issued.path] || null)) {
                throw new Error("结算中断记录或可信版本与票据不符，停止恢复");
            }
            const evidence = await this.inspectInterruptedUploadCleanup(issued.path);
            if (JSON.stringify(evidence) !== issued.evidence ||
                issued.expiresAt <= Date.now() ||
                issued.scope !== this.resumedContextKey() ||
                issued.record !== JSON.stringify(ledger[issued.path])) {
                throw new Error("重新只读核对时版本或暂存状态改变，必须重新确认");
            }
            if (evidence.stageState === "PRESENT") {
                if (record.phase !== "CLEANUP_ISSUED") {
                    await this.setUploadPhase(issued.path, "CLEANUP_ISSUED");
                }
                await this.deleteRemote({
                    path: evidence.stagePath, name: issued.path,
                    id: evidence.stageId, etag: evidence.stageEtag,
                    size: Number(record.sourceBytes),
                }, undefined, () => {
                    if (!this.unknownSettlementInFlight ||
                        !this.uploadLedgerReconciliationInFlight ||
                        this.settings() !== s ||
                        ledger[issued.path]?.phase !== "CLEANUP_ISSUED" ||
                        ledger[issued.path]?.operationId !== record.operationId ||
                        !this.resumedHistoricalIdentityTrusted()) {
                        throw new Error("云端条件清理派发前旧操作身份变化");
                    }
                });
            }
            // If stage is absent, this does not prove who deleted it. With
            // explicit consent + independently checked GET404 + exact target
            // bytes, the operator can acknowledge the absent staging object.
            await this.setUploadPhase(issued.path, "CLEANUP_VERIFIED", {
                cleanupReconciledBy: evidence.stageState === "PRESENT"
                    ? "CONDITIONAL_DELETE_ACK" : "EXPLICIT_ABSENT_STAGE_GET404",
            });
            await this.clearVerifiedUploadLedger(s.fileStates || {});
            if (ledger[issued.path]) {
                throw new Error("已核验结算仍有未决日志，不得对其他范围恢复授权");
            }
            s.status = "paused";
            s.runAllowed = false;
            s.enabled = false;
            this.notifyStatusChange();
            return true;
        } finally {
            this.uploadLedgerReconciliationInFlight = false;
            this.unknownSettlementInFlight = false;
            this.notifyStatusChange();
        }
    }

    private resumedContextKey(): string {
        const s = this.settings();
        const local: any = (this.store as any).settings.local;
        return JSON.stringify([
            String(local.microsoftAccountId || ""), s.vaultId, s.deviceId,
            s.installationProofId, s.confirmationAccountId, s.confirmationVaultId,
            s.localVaultId || "", s.cloudVaultId || "",
            s.confirmed, s.firstSyncCompleted,
            this.sharedStateFingerprint(), JSON.stringify(s.ignorePatterns || []),
            s.maxFileSizeMb,
        ]);
    }

    private resumedHistoricalIdentityTrusted(): boolean {
        if (this.identicalRegistrationBlocked()) return false;
        const s = this.settings();
        const local: any = (this.store as any).settings.local;
        const accountId = String(local.microsoftAccountId || "");
        return Boolean(accountId && s.confirmed === true
            && (s.firstSyncCompleted === true || this.androidFirstMergeIsTrusted(accountId, s.vaultId))
            && s.runAllowed === false
            && s.enabled === false
            && s.confirmationAccountId === accountId
            && s.confirmationVaultId === s.vaultId
            && (this.store as any).installationProofIsValid?.() === true);
    }

    private resumedProtectedActionAllowed(
        operation: RunTicket | undefined, path: string, action: ResumedAction,
    ): boolean {
        return Boolean(operation && this.resumedOperation === operation
            && this.resumedAuthorizedPath === path
            && this.resumedAuthorizedAction === action
            && this.resumedOperationIsCurrent(operation));
    }

    private resumedOperationIsCurrent(operation: RunTicket): boolean {
        return this.resumedOperation === operation
            && this.resumedInFlight && this.resumedTransientState?.runAllowed === true
            && this.resumedSettings === this.settings()
            && this.authority.isCurrent(operation)
            && this.resumedScope === this.resumedContextKey()
            && this.settings().confirmed === true
            && this.settings().runAllowed === false
            && this.settings().enabled === false
            && (this.store as any).installationProofIsValid?.() === true;
    }

    async confirmResumedOneDriveReview(
        ticket: object, targetPath: string, approved = false,
    ): Promise<boolean> {
        const issued = ticket && typeof ticket === "object"
            ? this.resumedReviewTickets.get(ticket) : undefined;
        // Every click, cancel and failure consumes this ticket. Never persist it.
        if (ticket && typeof ticket === "object") this.resumedReviewTickets.delete(ticket);
        if (!approved) return false;
        if (!issued || (issued.targetPath && issued.targetPath !== targetPath) ||
            issued.expiresAt <= Date.now() || this.resumedInFlight ||
            this.uploadLedgerReconciliationInFlight || this.uploadLedgerSettlementInFlight ||
            this.unknownSettlementInFlight || this.running || this.preparing || this.operationJob) {
            throw new Error("重启后的独立预览票据不存在、过期或正在执行，必须重新只读检查");
        }
        this.resumedInFlight = true;
        const startedGeneration = this.resumedApprovalGeneration;
        let operation: RunTicket | null = null;
        try {
            const snapshot = await this.captureResumedOneDriveReview(issued.targetPath);
            if (this.resumedApprovalGeneration !== startedGeneration ||
                JSON.stringify(snapshot) !== issued.snapshot ||
                Date.now() >= issued.expiresAt) {
                throw new Error("本机字节SHA、云端ItemID/ETag、完整清单或身份已经变化，旧预览不可执行");
            }
            const chosen = snapshot.entries.find((row) => row.path === targetPath);
            if (!chosen || !["UPLOAD", "DOWNLOAD", "DELETE_CLOUD", "DELETE_VAULT"].includes(chosen.action) ||
                chosen.path.startsWith("/") || chosen.path.includes("\\") ||
                chosen.path.split("/").includes("..")) {
                throw new Error("本条不是已完成只读审核的精确上传/下载范围，不允许自动选赢家或删除");
            }
            const s = this.settings();
            const permittedBytes = this.maxResumedFileBytes();
            if (chosen.localSize > permittedBytes || chosen.remoteSize > permittedBytes) {
                throw new Error("本次文件超过当前生效大小上限，不能沿用旧确认");
            }
            const scopeKey = this.resumedContextKey();
            const ephemeral = { confirmed: true, runAllowed: true };
            operation = this.authority.start(ephemeral, this.contextKey());
            this.resumedTransientState = ephemeral;
            this.resumedOperation = operation;
            this.resumedSettings = s;
            this.resumedScope = scopeKey;
            this.resumedAuthorizedPath = targetPath;
            this.resumedAuthorizedAction = chosen.action;
            // No mutation of s.runAllowed/confirmed/auto settings is permitted.
            // Even a crash during physical write cannot persist a broad grant.
            const beforeBaseline: Record<string, any> = { ...(s.fileStates || {}) };
            if (chosen.action === "DELETE_CLOUD" || chosen.action === "DELETE_VAULT") {
                await this.applyOneDriveDeletion(targetPath, operation);
                this.assertCurrent(operation);
                this.notifyStatusChange();
                return true;
            }
            let verifiedRemote: RemoteFile | null = null;
            let finalLocal: TFile | null = null;
            this.assertCurrent(operation);
            if (chosen.action === "UPLOAD") {
                const current = this.app.vault.getAbstractFileByPath(targetPath);
                if (!(current instanceof TFile)) throw new Error("已确认上传的本地原稿不存在");
                const data = await this.app.vault.readBinary(current);
                if (data.byteLength !== chosen.localSize ||
                    await this.sha256Bytes(data) !== chosen.localHash) {
                    throw new Error("最终操作前源文件字节已改变");
                }
                const remoteBefore = (await this.remoteFiles(true, true))
                    .find((item) => item.path === targetPath);
                if (chosen.remoteId
                    ? remoteBefore?.id !== chosen.remoteId || remoteBefore?.etag !== chosen.remoteEtag
                    : Boolean(remoteBefore)) {
                    throw new Error("上传前远端文件ID/ETag或同名占用变化，停止操作");
                }
                const uploaded = await this.uploadStaged(
                    operation, targetPath, data, false, remoteBefore || null,
                );
                this.assertCurrent(operation);
                verifiedRemote = (await this.remoteFiles(true, true))
                    .find((item) => item.path === targetPath) || null;
                if (!verifiedRemote?.id || !verifiedRemote.etag ||
                    verifiedRemote.id !== String(uploaded?.id || "") ||
                    verifiedRemote.etag !== String(uploaded?.eTag || uploaded?.etag || "") ||
                    !sameBytes(await this.remoteBytes(verifiedRemote), data)) {
                    throw new Error("云端已发提交后读回不完整，保留原基线并要求人工只读对账");
                }
                const currentAgain = this.app.vault.getAbstractFileByPath(targetPath);
                if (!(currentAgain instanceof TFile) ||
                    await this.sha256Bytes(await this.app.vault.readBinary(currentAgain)) !== chosen.localHash) {
                    throw new Error("云端提交后本地原稿变化，不能提升旧同步基线");
                }
                finalLocal = currentAgain;
            } else {
                const remoteBefore = (await this.remoteFiles(true, true))
                    .find((item) => item.path === targetPath);
                if (!remoteBefore || remoteBefore.id !== chosen.remoteId ||
                    remoteBefore.etag !== chosen.remoteEtag ||
                    await this.sha256Bytes(await this.remoteBytes(remoteBefore)) !== chosen.remoteHash) {
                    throw new Error("下载前云端版本或原文字节变化，旧确认不可执行");
                }
                const localFile = this.app.vault.getAbstractFileByPath(targetPath);
                const oldLocal = localFile instanceof TFile
                    ? { path: targetPath, file: localFile,
                        size: localFile.stat.size, mtime: localFile.stat.mtime } : undefined;
                if (Boolean(oldLocal) !== Boolean(chosen.localHash) ||
                    (oldLocal && (oldLocal.mtime !== chosen.localMtime ||
                        await this.sha256Bytes(await this.app.vault.readBinary(oldLocal.file)) !== chosen.localHash))) {
                    throw new Error("下载前本机目标改写，拒绝覆盖");
                }
                await this.download(operation, remoteBefore, targetPath, oldLocal, false,
                    chosen.localHash || undefined, chosen.remoteHash);
                this.assertCurrent(operation);
                const current = this.app.vault.getAbstractFileByPath(targetPath);
                if (!(current instanceof TFile) ||
                    await this.sha256Bytes(await this.app.vault.readBinary(current)) !== chosen.remoteHash) {
                    throw new Error("下载后目标物理读回SHA不匹配，不能提升旧基线");
                }
                finalLocal = current;
                verifiedRemote = remoteBefore;
            }
            this.assertCurrent(operation);
            if (!finalLocal || !verifiedRemote) throw new Error("双端验证未完成");
            const localHash = await this.sha256Bytes(await this.app.vault.readBinary(finalLocal));
            s.fileStates = {
                ...beforeBaseline,
                [targetPath]: {
                    localMtime: finalLocal.stat.mtime, localSize: finalLocal.stat.size,
                    localHash, remoteId: verifiedRemote.id,
                    remoteEtag: verifiedRemote.etag,
                },
            };
            s.status = "paused";
            s.enabled = false;
            s.runAllowed = false;
            s.lastError = "已单独批准并核验本条文件；其他范围仍须逐项重新预览和明确确认。";
            try {
                await (this.store as any).save();
                await this.clearVerifiedUploadLedger(s.fileStates);
            } catch (error) {
                s.fileStates = beforeBaseline;
                throw error;
            }
            this.notifyStatusChange();
            return true;
        } catch (error) {
            const s = this.settings();
            s.runAllowed = false;
            s.enabled = false;
            s.status = "paused";
            s.lastError = "本次单条重启后许可已失效或提交状态未知：" + safeOneDriveError(error);
            await (this.store as any).save();
            throw error;
        } finally {
            if (operation) this.authority.finishFinal(operation);
            if (this.resumedTransientState) this.resumedTransientState.runAllowed = false;
            this.resumedOperation = null;
            this.resumedTransientState = null;
            this.resumedSettings = null;
            this.resumedScope = "";
            this.resumedAuthorizedPath = "";
            this.resumedAuthorizedAction = null;
            this.resumedInFlight = false;
        }
    }

    private deletionApprovalRunning = false;

    // The presence of a historical deletion marker never overrides the
    // CURRENT file inclusion policy or a restored local original. Keep that
    // marker and baseline for review, but do not expose an executable action.
    private pendingDeletionPathInCurrentScope(path: string, direction: "cloud" | "vault"): boolean {
        if (this.isIgnored(path)) return false;
        const raw = this.app.vault.getAbstractFileByPath(path);
        const limit = this.maxBytes();
        const priorSize = Number(this.settings().fileStates?.[path]?.localSize || 0);
        if (!Number.isFinite(priorSize) || priorSize > limit) return false;
        if (direction === "cloud") return !raw;
        if (direction !== "vault" || !(raw instanceof TFile) || raw.stat.size > limit) return false;
        if (!Platform.isAndroidApp) return true;
        const settings = this.settings();
        const baseline = settings.fileStates?.[path];
        const pending = settings.pendingDeletions?.[path];
        return Boolean(baseline?.localHash && pending?.localHash === baseline.localHash);
    }

    private deletionStormState(
        local: LocalFile[], remote: RemoteFile[], previous: Record<string, any>,
        skippedPaths: Set<string>,
    ): { frozen: boolean; reason: string } {
        const localMap = new Map(local.map((item) => [item.path, item]));
        const remoteMap = new Map(remote.map((item) => [item.path, item]));
        const paths = [...new Set([
            ...localMap.keys(), ...remoteMap.keys(), ...Object.keys(previous),
        ])].filter((path) => !skippedPaths.has(path) && !this.isIgnored(path));
        const missingLocal = paths.filter((path) => !localMap.has(path) &&
            remoteMap.has(path) && Boolean(previous[path]?.localMtime));
        const missingRemote = paths.filter((path) => localMap.has(path) &&
            !remoteMap.has(path) && Boolean(previous[path]?.remoteId) &&
            previous[path]?.localMtime === localMap.get(path)?.mtime &&
            previous[path]?.localSize === localMap.get(path)?.size);
        const total = missingLocal.length + missingRemote.length;
        const files = Math.max(local.length, remote.length, Object.keys(previous).length);
        const threshold = Math.max(3, Math.ceil(files * 0.05));
        const remoteUnexpectedlyEmpty = remote.length === 0 &&
            Object.entries(previous).some(([path, item]) => !skippedPaths.has(path) && !this.isIgnored(path) && Boolean(item?.remoteId));
        const frozen = total > threshold || remoteUnexpectedlyEmpty;
        return { frozen, reason: remoteUnexpectedlyEmpty
            ? "DELETION_FROZEN：历史非空的云端清单变为空，必须重新核实，禁止按旧删除确认移除本机原件。"
            : frozen
                ? `DELETION_FROZEN：待删除 ${total} 项超过安全阈值 ${threshold} 项。请先核对缺失来源、恢复误删文件，再重新检查；最终确认不能绕过风暴保护。`
                : "" };
    }

    private async applyOneDriveDeletion(
        targetPath: string, operation: RunTicket,
        automaticInventory?: () => Promise<{ rootId: string; files: RemoteFile[] }>,
    ): Promise<void> {
        const automatic = !!automaticInventory;
        const resumedApproved = Boolean(!automatic
            && this.resumedProtectedActionAllowed(
                operation, targetPath, this.resumedAuthorizedAction || "HOLD",
            ) && (this.resumedAuthorizedAction === "DELETE_CLOUD"
                || this.resumedAuthorizedAction === "DELETE_VAULT"));
        if (!(resumedApproved || (automatic && this.running && this.settings().firstSyncCompleted && this.canRun())) ||
            this.deletionApprovalRunning) {
            throw new Error("运行未授权、仍在同步中或删除清单不合法");
        }
        this.assertCurrent(operation);
        if (this.hasPendingUploadAt(targetPath)) throw new Error("该文件上次上传仍在核对，暂时无法删除。");
        const settings = this.settings();
        const requireNoDeletionStorm = (allRemote: RemoteFile[]) => {
            if (automatic) return;
            const skipped = this.skippedLocalPaths();
            const remote = allRemote.filter((item) => {
                if (this.isIgnored(item.path) || item.size > this.maxBytes()) {
                    skipped.add(item.path);
                    return false;
                }
                return true;
            });
            const storm = this.deletionStormState(
                this.localFiles(), remote, settings.fileStates || {}, skipped,
            );
            if (storm.frozen) {
                settings.status = "paused";
                settings.lastError = storm.reason;
                const error: any = new Error(storm.reason);
                error.isDeletionFrozen = true;
                throw error;
            }
        };
        // A pending marker pins an old version, not the current filter policy.
        // Bind this approval to the current policy and recheck its source
        // after every asynchronous read/backup, immediately before deletion.
        const ruleScope = JSON.stringify([settings.ignorePatterns || [], this.maxBytes()]);
        const requireCurrentDeletionScope = (path: string, direction: "cloud" | "vault") => {
            if (this.settings() !== settings ||
                JSON.stringify([settings.ignorePatterns || [], this.maxBytes()]) !== ruleScope ||
                !this.pendingDeletionPathInCurrentScope(path, direction)) {
                throw new Error("删除候选的排除规则、大小范围或本地原稿已变化，旧确认不能执行");
            }
        };
        const consented = settings.pendingDeletions || {};
        const path = targetPath;
        const candidate = consented[path];
        const currentAccount = String((this.store as any).settings.local.microsoftAccountId || "");
        if (!candidate || (automatic && candidate.automatic !== true) || !currentAccount ||
            candidate.accountId !== currentAccount ||
            candidate.vaultId !== settings.vaultId ||
            candidate.deviceId !== settings.deviceId ||
            !path || path.includes("\\") || path.split("/").some(part => !part || part === "." || part === "..") ||
            (resumedApproved && candidate.direction !==
                (this.resumedAuthorizedAction === "DELETE_CLOUD" ? "cloud" : "vault"))) {
            throw new Error("删除范围或版本已失效");
        }
        requireCurrentDeletionScope(path, candidate.direction);

        const record = { ...candidate };
        this.deletionApprovalRunning = true;
        try {
            this.assertCurrent(operation);
            requireCurrentDeletionScope(path, record.direction);
            const current = settings.pendingDeletions?.[path];
            if (!current || JSON.stringify(current) !== JSON.stringify(record)) {
                throw new Error("删除确认的目标版本已变化，拒绝旧许可");
            }
            const verifiedInventory = await automaticInventory?.();
            const allRemote = verifiedInventory?.files ?? await this.remoteFiles(true, true);
            if (automatic) await this.store.save();
            this.assertCurrent(operation);
            requireCurrentDeletionScope(path, record.direction);
            requireNoDeletionStorm(allRemote);
            const remote = allRemote.find((item) => item.path === path);
            const local = this.localFiles().find((item) => item.path === path);
            const old = settings.fileStates?.[path];
            if (!old) throw new Error("原始双端基线已消失，拒绝删除");
            if (record.direction === "cloud") {
                if (local || !remote || !remote.id || !remote.etag ||
                    !Number.isFinite(remote.size) || remote.size > this.maxBytes() ||
                    remote.id !== record.remoteId ||
                    remote.etag !== record.remoteEtag ||
                    old.remoteId !== record.remoteId || old.remoteEtag !== record.remoteEtag ||
                    await this.app.vault.adapter.exists(path)) {
                    throw new Error("云端文件版本或本地删除状态发生变化，拒绝危险操作");
                }
                if (!automatic) await this.createRecoveryFromRemote(operation, path, remote);
                const newest = await this.readRemoteFileMetadata(path);
                if (verifiedInventory && await this.readDeletionRoot(operation) !== verifiedInventory.rootId) {
                    throw new Error(localize("OneDrive 同步目录已变化，已暂缓删除。"));
                }
                const afterBackupRemote = automatic ? allRemote : await this.remoteFiles(true, true);
                this.assertCurrent(operation);
                requireCurrentDeletionScope(path, record.direction);
                requireNoDeletionStorm(afterBackupRemote);
                if (JSON.stringify(settings.pendingDeletions?.[path]) !== JSON.stringify(record)) {
                    throw new Error("冷备期间删除候选已变化，拒绝旧许可");
                }
                if (!newest || newest.id !== remote.id || newest.etag !== remote.etag ||
                    (Number.isFinite(newest.size) && newest.size > this.maxBytes()) ||
                    !afterBackupRemote.some((item) => item.path === path &&
                        item.id === remote.id && item.etag === remote.etag) ||
                    await this.app.vault.adapter.exists(path)) {
                    throw new Error("云端原稿在冷备后变更或已超出当前范围，禁止删除");
                }
                this.authority.beginFinal(operation);
                // The token and optional 401-refresh await occur inside
                // fetchGraph. A final pre-send guard must execute AFTER
                // each token resolves, before each physical DELETE fetch.
                const verifyBeforeDeleteDispatch = () => {
                    this.assertCurrent(operation);
                    requireCurrentDeletionScope(path, record.direction);
                    requireNoDeletionStorm(afterBackupRemote);
                    if (JSON.stringify(settings.pendingDeletions?.[path]) !== JSON.stringify(record) ||
                        settings.fileStates?.[path]?.remoteId !== remote.id ||
                        settings.fileStates?.[path]?.remoteEtag !== remote.etag) {
                        throw new Error("取得凭据期间删除候选或原版本已变化，禁止发出云端删除");
                    }
                };
                try {
                    await this.deleteRemote(remote, operation, verifyBeforeDeleteDispatch, async () => {
                        if (await this.app.vault.adapter.exists(path)) {
                            throw new Error("本地文件已恢复，取消删除");
                        }
                    });
                } catch (error: any) {
                    if (automatic && (this.isTemporarySyncFailure(error) || [404, 412].includes(Number(error?.status)))) {
                        error.recheckDeletion = true;
                    }
                    throw error;
                } finally {
                    this.clearRemoteDeltaCache();
                }
            } else if (record.direction === "vault") {
                if (remote || !local ||
                    local.mtime !== record.localMtime ||
                    local.size !== record.localSize) {
                    throw new Error("知识库原稿已变化或云端版本重新出现，拒绝删除");
                }
                if (verifiedInventory && (old.remoteId !== record.remoteId || old.localHash !== record.localHash)) {
                    throw new Error(localize("OneDrive 尚未确认此文件已删除。"));
                }
                const approvedBytes = await this.app.vault.readBinary(local.file);
                if (await this.sha256Bytes(approvedBytes) !== record.localHash) {
                    throw new Error("知识库原稿已变化，拒绝删除");
                }
                if (this.deleteGuard && await this.deleteGuard(path)) {
                    throw new Error("文件受外部文件夹同步保护，本次 OneDrive 删除许可不得跨域生效");
                }
                const finalRemote = automatic ? allRemote : await this.remoteFiles(true, true);
                this.assertCurrent(operation);
                requireCurrentDeletionScope(path, record.direction);
                requireNoDeletionStorm(finalRemote);
                const reappeared = finalRemote.find((item) => item.path === path);
                if (reappeared || (this.deleteGuard && await this.deleteGuard(path)) ||
                    (verifiedInventory && !await this.confirmRemoteDeletion(operation, path, old.remoteId, verifiedInventory.rootId))) {
                    throw new Error("删除前云端原稿或跨模块保护发生变化，禁止本地删除");
                }
                this.assertCurrent(operation);
                requireCurrentDeletionScope(path, record.direction);
                requireNoDeletionStorm(finalRemote);
                const currentLocal = this.app.vault.getAbstractFileByPath(path);
                if (!(currentLocal instanceof TFile) || currentLocal !== local.file ||
                    local.file.stat.mtime !== record.localMtime ||
                    local.file.stat.size !== record.localSize ||
                    JSON.stringify(settings.pendingDeletions?.[path]) !== JSON.stringify(record)) {
                    throw new Error("删除前本机文件、版本或删除候选变化，旧确认失效");
                }
                if (!Platform.isAndroidApp && typeof this.app.vault.trash !== "function") {
                    throw new Error("Obsidian原生回收站不可用，不执行危险删除");
                }
                // Full-list and cross-module checks can yield after the
                // first content read. Re-read actual bytes after those
                // awaits, including same-mtime/same-size local edits.
                const finalBytes = await this.app.vault.readBinary(currentLocal);
                const finalHash = await this.sha256Bytes(finalBytes);
                this.assertCurrent(operation);
                requireCurrentDeletionScope(path, record.direction);
                requireNoDeletionStorm(finalRemote);
                if (finalHash !== record.localHash ||
                    this.app.vault.getAbstractFileByPath(path) !== currentLocal ||
                    currentLocal.stat.mtime !== record.localMtime ||
                    currentLocal.stat.size !== record.localSize ||
                    JSON.stringify(settings.pendingDeletions?.[path]) !== JSON.stringify(record)) {
                    throw new Error("最终清单核对后本机正文或删除候选已变化，保留原文");
                }
                this.authority.beginFinal(operation);
                if (Platform.isAndroidApp) await this.app.vault.delete(currentLocal);
                else await this.app.vault.trash(currentLocal, true);
                if (await this.app.vault.adapter.exists(path)) throw new Error("本地原路径仍然存在，未完成同步删除");
            } else {
                throw new Error("未知删除方向");
            }
            if (!automatic) {
                settings.status = "paused";
                settings.lastError = "本批确认的删除已执行，需再同步核对剩余文件";
            }
            await this.finishOneDriveDeletion(path);
        } catch (error: any) {
            if (error?.isDeletionFrozen) await (this.store as any).save();
            throw error;
        } finally {
            this.deletionApprovalRunning = false;
        }
    }

    private async finishOneDriveDeletion(path: string): Promise<void> {
        const settings = this.settings();
        const previous = settings.fileStates[path];
        const records = settings.pendingDeletions!;
        const record = records[path];
        delete settings.fileStates[path];
        delete records[path];
        try {
            await this.store.save();
        } catch (error) {
            if (this.settings() === settings && settings.pendingDeletions === records) {
                if (previous) settings.fileStates[path] ??= previous;
                if (record) records[path] ??= record;
            }
            throw error;
        }
    }

    private async readDeletionRoot(operation: RunTicket): Promise<string> {
        return this.readSyncRoot(this.settings().vaultId, () => this.assertCurrent(operation));
    }

    private async readSyncRoot(vaultId: string, assertScope: () => void): Promise<string> {
        const path = encodePath(`vaults/${vaultId}/files`);
        const root = await this.graph("GET", `/drive/special/approot:/${path}`, undefined, undefined, assertScope);
        if (!root?.id || !root.folder) throw new Error(localize("OneDrive 同步目录已变化，已暂缓删除。"));
        return String(root.id);
    }

    private async verifyFirstSyncRoot(
        vaultId: string, baselines: Record<string, SyncFileState>, assertScope: () => void,
    ): Promise<string | undefined> {
        assertScope();
        if (!Object.values(baselines).some(file => file.remoteId)) return undefined;
        const rootId = await this.readSyncRoot(vaultId, assertScope);
        const files = await this.remoteFiles(true, true, { vaultId }, rootId);
        if (!files.some(file => baselines[file.path]?.remoteId === file.id) ||
            await this.readSyncRoot(vaultId, assertScope) !== rootId) {
            throw new Error(localize("OneDrive 同步目录已变化，已暂缓删除。"));
        }
        return rootId;
    }

    private async confirmRemoteDeletion(operation: RunTicket, path: string, previousId: string, rootId: string): Promise<boolean> {
        if (!previousId) return false;
        const guard = () => this.assertCurrent(operation);
        if (await this.readDeletionRoot(operation) !== rootId) return false;
        try {
            await this.graph("GET",
                "/drive/items/" + encodeURIComponent(previousId) + "?$select=id,name,eTag,size,fileSystemInfo",
                undefined, undefined, guard);
            return false;
        } catch (error: any) {
            guard();
            if (error?.status !== 404) throw error;
            return !await this.readRemoteFileMetadata(path);
        }
    }

    private remoteIdentityConflict: {
        accountId: string; vaultId: string; path: string;
        items: { id: string; etag: string; path: string }[];
    } | null = null;

    getOneDriveRemoteIdentityConflict(): { path: string; items: { id: string; etag: string; path: string }[] } | null {
        const conflict = this.remoteIdentityConflict;
        if (!conflict || conflict.accountId !== String((this.store as any).settings.local.microsoftAccountId || "") ||
            conflict.vaultId !== this.basePath()) return null;
        return { path: conflict.path, items: conflict.items.map((item) => ({ ...item })) };
    }

    async recheckOneDriveRemoteIdentities(): Promise<void> {
        if (this.running || this.preparing || this.operationJob) throw new Error("当前操作尚未结束，稍后再只读核对");
        await this.remoteFiles(true, true);
    }

    private identicalRegistrationRequested = false;
    private identicalRegistrationInFlight = false;
    private identicalRegistrationVerifiedId = "";
    private identicalRegistrationTickets = new WeakMap<object, { snapshot: string; expiresAt: number }>();

    private identicalRegistrationBlocked(): boolean {
        const record = this.settings().identicalRegistrationJournal;
        return this.identicalRegistrationInFlight || Boolean(record &&
            (record.phase !== "COMPLETED" || this.identicalRegistrationVerifiedId !== record.operationId));
    }

    private async verifyCompletedIdenticalRegistration(): Promise<void> {
        const record = this.settings().identicalRegistrationJournal;
        if (!record || this.identicalRegistrationVerifiedId === record.operationId) return;
        if (record.phase !== "COMPLETED" || this.identicalRegistrationInFlight ||
            !this.identicalRegistrationIdentityMatches(record.snapshot)) {
            throw new Error("原同内容登记仍待核实，必须先核验原操作");
        }
        // A completed baseline is history. Current edits belong in the fresh daily preview.
        this.assertIdenticalSnapshotPhysical(record.snapshot, false);
        if (!await this.persistedIdenticalBindingMatches(record, true)) throw new Error("登记历史回执与实际持久绑定不一致");
        if (!this.identicalRegistrationIdentityMatches(record.snapshot)) throw new Error("登记身份已变化");
        this.assertIdenticalSnapshotPhysical(record.snapshot, false);
        this.identicalRegistrationVerifiedId = record.operationId;
    }

    private identicalRegistrationIdentityMatches(snapshot: any): boolean {
        const settings = this.settings();
        const root = String((this.app.vault.adapter as any)?.getBasePath?.() || "")
            .replace(/\\/g, "/").replace(/\/+$/, "").toLowerCase();
        return Boolean(snapshot && (this.store as any).installationProofIsValid?.() === true &&
            root === snapshot.physicalRoot &&
            String((this.store as any).settings.local.microsoftAccountId || "") === snapshot.accountId &&
            settings.deviceId === snapshot.deviceId && settings.installationProofId === snapshot.proofId &&
            (settings.localVaultId || settings.vaultId) === snapshot.localVaultId &&
            (!settings.cloudVaultId || settings.cloudVaultId === snapshot.cloudVaultId));
    }

    private assertIdenticalSnapshotPhysical(snapshot: any, files = true): void {
        requireWindowsHost();
        if (!this.identicalRegistrationIdentityMatches(snapshot)) throw new Error("登记作用域已变化，禁止写本机状态");
        const fs = require("fs") as typeof import("fs");
        const paths = require("path").win32 as typeof import("path").win32;
        const identity = (value: import("fs").Stats) => [String(value.dev), String(value.ino), String(value.birthtimeMs)].join(":");
        const root = paths.normalize(snapshot.physicalRoot);
        const rootStat = fs.lstatSync(root);
        if (!rootStat.isDirectory() || rootStat.isSymbolicLink() || identity(rootStat) !== snapshot.rootPhysicalIdentity ||
            paths.normalize(fs.realpathSync(root)).toLowerCase() !== root.toLowerCase()) throw new Error("登记物理根已变化");
        if (files) for (const file of snapshot.files) {
            const absolute = paths.resolve(root, file.path);
            const stat = fs.lstatSync(absolute);
            if (!absolute.toLowerCase().startsWith(root.toLowerCase() + "\\") || !stat.isFile() || stat.isSymbolicLink() ||
                identity(stat) !== file.physicalIdentity || stat.size !== file.localSize ||
                paths.normalize(fs.realpathSync(absolute)).toLowerCase() !== absolute.toLowerCase())
                throw new Error("登记文件物理身份已变化");
        }
    }

    private async saveIdenticalRegistration(record: any, checkFiles = true): Promise<void> {
        this.assertIdenticalSnapshotPhysical(record.snapshot, checkFiles);
        await (this.store as any).saveWithIdentityGuard(() => {
            if (this.settings().identicalRegistrationJournal !== record) throw new Error("登记原操作已变化");
            this.assertIdenticalSnapshotPhysical(record.snapshot, checkFiles);
        });
    }

    private async captureIdenticalRegistration(cloudVaultId: string, ownRecord?: any): Promise<any> {
        requireWindowsHost();
        const settings = this.settings();
        if (!(this.store as any).installationProofIsValid?.() || settings.runAllowed || settings.enabled ||
            (!ownRecord && (settings.confirmed || settings.firstSyncCompleted ||
                Object.keys(settings.fileStates || {}).length || settings.confirmationAccountId || settings.confirmationVaultId)) ||
            (ownRecord && settings.identicalRegistrationJournal !== ownRecord) ||
            Object.keys(settings.pendingUploadOperations || {}).length ||
            Object.keys(settings.pendingDeletions || {}).length || settings.pendingSharedUploadOperation ||
            settings.firstCommitJournal || settings.firstFileJournal ||
            ((this.store as any).settings.local.recoveryEntries || []).length) {
            throw new Error("同内容登记需要可核实的新作用域；旧绑定、恢复或未决操作必须先处理");
        }
        const accountId = String((this.store as any).settings.local.microsoftAccountId || "");
        const localVaultId = ownRecord?.snapshot.localVaultId || settings.vaultId;
        if (!accountId || !localVaultId || !cloudVaultId || localVaultId === cloudVaultId)
            throw new Error("账号或独立本地/云端库身份不可核实");
        const rawRoot = String((this.app.vault.adapter as any)?.getBasePath?.() || "");
        const paths = require("path").win32 as typeof import("path").win32;
        const root = paths.normalize(rawRoot);
        if (!/^[A-Za-z]:\\/.test(root) || root.length <= 3) throw new Error("无法核实Windows物理根");
        const nativeFs = require("fs") as typeof import("fs");
        const fs = nativeFs.promises;
        const physical = paths.normalize(await fs.realpath(root));
        const stat = await fs.lstat(root);
        if (physical.toLowerCase() !== root.toLowerCase() || !stat.isDirectory() || stat.isSymbolicLink())
            throw new Error("物理根重定向或不可核实");
        const rootAudit = await inspectReceiverCandidateInventory(this.app, null, undefined, {
            accountId, vaultId: settings.vaultId, deviceId: settings.deviceId,
            installationProofId: settings.installationProofId,
        });
        if (!rootAudit.independentPhysicalRootVerified) throw new Error("本地根不是可核实的独立Vault：" + rootAudit.reason);
        const identity = (value: import("fs").Stats) => [String(value.dev), String(value.ino), String(value.birthtimeMs)].join(":");
        const rootPhysicalIdentity = identity(stat);
        const assertPhysical = (absolute?: string, expectedFile?: import("fs").Stats) => {
            const now = nativeFs.lstatSync(root);
            if (!now.isDirectory() || now.isSymbolicLink() || identity(now) !== rootPhysicalIdentity ||
                paths.normalize(nativeFs.realpathSync(root)).toLowerCase() !== root.toLowerCase())
                throw new Error("登记读取期间物理根变化，停止访问");
            if (absolute && expectedFile) {
                const file = nativeFs.lstatSync(absolute);
                if (!file.isFile() || file.isSymbolicLink() || identity(file) !== identity(expectedFile) ||
                    file.size !== expectedFile.size || paths.normalize(nativeFs.realpathSync(absolute)).toLowerCase() !== absolute.toLowerCase())
                    throw new Error("登记读取期间文件或父路径变化，禁止读取外部正文");
            }
        };
        const readProofs: Array<{ absolute: string; stat: import("fs").Stats }> = [];
        const physicalRoot = rawRoot.replace(/\\/g, "/").replace(/\/+$/, "").toLowerCase();
        const before = this.previewVaultId;
        this.previewVaultId = cloudVaultId;
        try {
            await this.assertFirstCommitGraphAccount(accountId);
            this.vaultRegistryCache = null;
            const registry = await this.readVaultRegistry();
            if (!this.registryReadIsTrusted ||
                registry.filter((v: any) => v?.id === cloudVaultId).length !== 1)
                throw new Error("目标登记身份不完整，拒绝同名猜测或重建注册表");
            const local = this.localFiles().sort((a, b) => a.path.localeCompare(b.path));
            const remote = (await this.remoteFiles(true, true)).filter((item) => !this.isIgnored(item.path));
            if (!local.length || local.length !== remote.length ||
                local.some((f) => f.size > this.maxBytes()) || remote.some((f) => f.size > this.maxBytes()))
                throw new Error("只支持完整适用范围全部同内容；空库、单侧文件或超限项目不在本次登记范围");
            const remoteMap = new Map(remote.map((item) => [item.path, item]));
            const localKeys = new Set<string>();
            const files: any[] = [];
            for (const item of local) {
                const key = item.path.normalize("NFC").toLowerCase();
                if (localKeys.has(key)) throw new Error("本机存在路径碰撞，停止登记");
                localKeys.add(key);
                const peer = remoteMap.get(item.path);
                if (!peer?.id || !peer.etag || peer.size !== item.size) throw new Error("存在单侧文件、大小差异或未知云端版本");
                const absolute = paths.resolve(root, item.path);
                if (!absolute.toLowerCase().startsWith(root.toLowerCase() + "\\")) throw new Error("本机路径不在当前Vault内");
                const fileStat = await fs.lstat(absolute);
                const resolvedFile = paths.normalize(await fs.realpath(absolute));
                if (!fileStat.isFile() || fileStat.isSymbolicLink() ||
                    resolvedFile.toLowerCase() !== absolute.toLowerCase() || fileStat.size !== item.size)
                    throw new Error("本机文件重定向、链接或版本不可核实，停止登记");
                if (this.deleteGuard && await this.deleteGuard(item.path, true)) throw new Error("路径受其它边界保护，不能登记");
                // Recheck AFTER the async external guard, before opening or reading a body.
                assertPhysical(absolute, fileStat);
                const handle = await fs.open(absolute, "r");
                let a: ArrayBuffer;
                try {
                    const opened = await handle.stat();
                    if (identity(opened) !== identity(fileStat) || !opened.isFile())
                        throw new Error("读取句柄不是已批准的本机文件，拒绝正文读取");
                    assertPhysical(absolute, fileStat);
                    a = Uint8Array.from(await handle.readFile()).buffer;
                    assertPhysical(absolute, fileStat);
                } finally { await handle.close(); }
                const b = await this.remoteBytes(peer);
                assertPhysical(absolute, fileStat);
                readProofs.push({ absolute, stat: fileStat });
                if (a.byteLength !== item.size || b.byteLength !== peer.size || !sameBytes(a, b))
                    throw new Error("双端内容不同或正文不可核实，禁止登记及覆盖");
                files.push({ path: item.path, localMtime: item.mtime, localSize: item.size,
                    localHash: await this.sha256Bytes(a), remoteHash: await this.sha256Bytes(b),
                    physicalIdentity: identity(fileStat),
                    remoteId: peer.id, remoteEtag: peer.etag });
            }
            const result = { accountId, localVaultId, cloudVaultId, deviceId: settings.deviceId,
                proofId: settings.installationProofId, physicalRoot, rootPhysicalIdentity,
                registryHash: await contentHash(JSON.stringify(registry)),
                rulesHash: await contentHash(JSON.stringify([this.maxBytes(), settings.ignorePatterns || []])), files };
            assertPhysical();
            for (const proof of readProofs) assertPhysical(proof.absolute, proof.stat);
            if (!this.identicalRegistrationIdentityMatches(result)) throw new Error("读取期间物理根或账号/设备发生变化");
            return result;
        } finally { this.previewVaultId = before; }
    }

    async prepareIdenticalOneDriveRegistration(cloudVaultId: string): Promise<any> {
        if (this.running || this.identicalRegistrationInFlight || this.settings().identicalRegistrationJournal ||
            !this.identicalRegistrationRequested || this.pendingFirstChoice?.id !== cloudVaultId)
            throw new Error("必须从当前云库的独立高级同内容入口重新检查");
        this.identicalRegistrationInFlight = true;
        try {
            const snapshot = await this.captureIdenticalRegistration(cloudVaultId);
            const ticket = Object.freeze({});
            this.identicalRegistrationTickets.set(ticket, { snapshot: JSON.stringify(snapshot), expiresAt: Date.now() + 300000 });
            return { ticket, ...snapshot, proofId: undefined };
        } finally { this.identicalRegistrationInFlight = false; }
    }

    private async persistedIdenticalBindingMatches(record: any, historical = false, exactStage?: "BOUND_UNVERIFIED" | "COMPLETED"): Promise<boolean> {
        const persisted = await (this.store as any).readPersistedOneDriveRegistrationState(() => this.assertIdenticalSnapshotPhysical(record.snapshot, false));
        const state = persisted?.oneDrive;
        return Boolean(state && persisted.accountId === record.snapshot.accountId &&
            state.vaultId === record.snapshot.cloudVaultId && state.localVaultId === record.snapshot.localVaultId &&
            state.cloudVaultId === record.snapshot.cloudVaultId && state.deviceId === record.snapshot.deviceId &&
            state.installationProofId === record.snapshot.proofId && state.runAllowed === false && state.enabled === false &&
            state.confirmationAccountId === record.snapshot.accountId && state.confirmationVaultId === record.snapshot.cloudVaultId &&
            state.identicalRegistrationJournal?.operationId === record.operationId &&
            (historical ? state.firstSyncCompleted === true && state.confirmed === true &&
                state.identicalRegistrationJournal?.phase === "COMPLETED" &&
                JSON.stringify(state.identicalRegistrationJournal?.snapshot) === JSON.stringify(record.snapshot) &&
                JSON.stringify(state.identicalRegistrationJournal?.fileStates) === JSON.stringify(record.fileStates)
                : JSON.stringify(state.fileStates) === JSON.stringify(record.fileStates) &&
                    (!exactStage || (state.identicalRegistrationJournal?.phase === exactStage &&
                        state.firstSyncCompleted === (exactStage === "COMPLETED") && state.confirmed === true))));
    }

    async confirmIdenticalOneDriveRegistration(ticket: object, approved = false): Promise<boolean> {
        const issued = this.identicalRegistrationTickets.get(ticket);
        this.identicalRegistrationTickets.delete(ticket);
        if (!approved) return false;
        if (!issued || issued.expiresAt <= Date.now() || this.identicalRegistrationInFlight || this.running)
            throw new Error("同内容登记批准已失效，请重新预览");
        this.identicalRegistrationInFlight = true;
        const settings = this.settings();
        let record: any;
        try {
            const expected = JSON.parse(issued.snapshot);
            const latest = await this.captureIdenticalRegistration(expected.cloudVaultId);
            if (JSON.stringify(latest) !== issued.snapshot) throw new Error("同内容预览已过期，未登记");
            const fileStates: Record<string, any> = {};
            for (const item of latest.files) fileStates[item.path] = { localMtime: item.localMtime,
                localSize: item.localSize, localHash: item.localHash, remoteId: item.remoteId, remoteEtag: item.remoteEtag };
            this.assertIdenticalSnapshotPhysical(latest);
            record = { operationId: "identical-" + crypto.randomUUID(), phase: "PREPARED",
                snapshot: latest, fileStates, before: {
                    vaultId: settings.vaultId, localVaultId: settings.localVaultId, cloudVaultId: settings.cloudVaultId,
                    confirmed: settings.confirmed, firstSyncCompleted: settings.firstSyncCompleted,
                    confirmationAccountId: settings.confirmationAccountId, confirmationVaultId: settings.confirmationVaultId,
                    fileStates: settings.fileStates, lastSuccessAt: settings.lastSuccessAt,
                }, createdAt: Date.now() };
            settings.identicalRegistrationJournal = record;
            await this.saveIdenticalRegistration(record);
            if (JSON.stringify(await this.captureIdenticalRegistration(latest.cloudVaultId, record)) !== issued.snapshot)
                throw new Error("登记保存期间内容或范围变化，保留原操作待核对");
            return await this.completeIdenticalRegistration(record);
        } catch (error) {
            if (record && settings === this.settings() && settings.identicalRegistrationJournal === record) {
                record.phase = "RECONCILE_REQUIRED";
                settings.firstSyncCompleted = false; settings.runAllowed = false; settings.enabled = false;
                settings.lastSuccessAt = record.before.lastSuccessAt;
                settings.status = "paused";
                settings.lastError = "同内容登记未核实，保留原操作；须先核验再继续。";
                try { await this.saveIdenticalRegistration(record, false); } catch { /* never persist into a changed root/account */ }
            }
            throw error;
        } finally { this.identicalRegistrationInFlight = false; }
    }

    private async completeIdenticalRegistration(record: any): Promise<boolean> {
        const settings = this.settings();
        const expected = JSON.stringify(record.snapshot);
        if (JSON.stringify(await this.captureIdenticalRegistration(record.snapshot.cloudVaultId, record)) !== expected)
            throw new Error("原登记范围已变化，不能沿旧许可继续");
        Object.assign(settings, { vaultId: record.snapshot.cloudVaultId, localVaultId: record.snapshot.localVaultId,
            cloudVaultId: record.snapshot.cloudVaultId, confirmationAccountId: record.snapshot.accountId,
            confirmationVaultId: record.snapshot.cloudVaultId, confirmed: true, firstSyncCompleted: false,
            fileStates: record.fileStates, runAllowed: false, enabled: false, status: "paused" });
        record.phase = "BOUND_UNVERIFIED";
        await this.saveIdenticalRegistration(record);
        if (!await this.persistedIdenticalBindingMatches(record, false, "BOUND_UNVERIFIED")) throw new Error("登记实际持久读回不一致");
        if (JSON.stringify(await this.captureIdenticalRegistration(record.snapshot.cloudVaultId, record)) !== expected)
            throw new Error("持久登记后当前范围已变化，保持暂停");
        settings.firstSyncCompleted = true;
        settings.lastSuccessAt = Date.now();
        settings.lastError = "双端内容已核实并登记；尚无传输许可，请重新预览后独立批准日常同步。";
        record.phase = "COMPLETED";
        record.verifiedAt = Date.now();
        await this.saveIdenticalRegistration(record);
        if (!await this.persistedIdenticalBindingMatches(record, false, "COMPLETED") ||
            JSON.stringify(await this.captureIdenticalRegistration(record.snapshot.cloudVaultId, record)) !== expected)
            throw new Error("最终登记保存期间范围变化，原操作待核验");
        this.identicalRegistrationVerifiedId = record.operationId;
        return true;
    }

    async reconcileIdenticalOneDriveRegistration(): Promise<boolean> {
        if (this.identicalRegistrationInFlight || this.running) throw new Error("原登记仍在执行");
        const settings = this.settings();
        const record = settings.identicalRegistrationJournal;
        if (!record) return true;
        if (record.phase === "COMPLETED") { await this.verifyCompletedIdenticalRegistration(); return true; }
        if (!this.identicalRegistrationIdentityMatches(record.snapshot)) throw new Error("原登记身份不匹配，不能重新绑定");
        this.identicalRegistrationInFlight = true;
        try {
            if (!await this.persistedIdenticalBindingMatches(record)) {
                const persisted = await (this.store as any).readPersistedOneDriveRegistrationState(() => this.assertIdenticalSnapshotPhysical(record.snapshot, false));
                const old = record.before;
                const state = persisted?.oneDrive;
                if (!state || persisted.accountId !== record.snapshot.accountId || state.vaultId !== old.vaultId ||
                    state.deviceId !== record.snapshot.deviceId || state.installationProofId !== record.snapshot.proofId ||
                    state.confirmed !== old.confirmed || state.firstSyncCompleted !== old.firstSyncCompleted ||
                    JSON.stringify(state.fileStates) !== JSON.stringify(old.fileStates))
                    throw new Error("原登记结果与预期新旧绑定均不一致，保留记录待人工核查");
                // The old binding is still durable: invalidate the consumed approval, never register again.
                const priorReceipt = settings.lastIdenticalRegistration;
                settings.lastIdenticalRegistration = { ...record, phase: "NOT_COMMITTED" };
                Object.assign(settings, record.before);
                settings.identicalRegistrationJournal = null;
                settings.runAllowed = false; settings.enabled = false;
                try { await (this.store as any).saveWithIdentityGuard(() => this.assertIdenticalSnapshotPhysical(record.snapshot, false)); }
                catch (error) {
                    settings.identicalRegistrationJournal = record;
                    if (priorReceipt === undefined) delete settings.lastIdenticalRegistration;
                    else settings.lastIdenticalRegistration = priorReceipt;
                    throw error;
                }
                return false;
            }
            // This operation only wrote metadata. Exact persisted target facts prove the original
            // registration; current edits are later changes for a NEW daily preview, not authority.
            Object.assign(settings, { vaultId: record.snapshot.cloudVaultId, localVaultId: record.snapshot.localVaultId,
                cloudVaultId: record.snapshot.cloudVaultId, confirmationAccountId: record.snapshot.accountId,
                confirmationVaultId: record.snapshot.cloudVaultId, fileStates: record.fileStates });
            settings.firstSyncCompleted = true; settings.confirmed = true;
            settings.runAllowed = false; settings.enabled = false; settings.status = "paused";
            settings.lastSuccessAt = record.createdAt;
            settings.lastError = "原同内容登记已核实；当前后续变化须另行预览和独立批准日常同步。";
            record.phase = "COMPLETED"; record.verifiedAt = Date.now();
            await this.saveIdenticalRegistration(record, false);
            if (!await this.persistedIdenticalBindingMatches(record, true) ||
                !this.identicalRegistrationIdentityMatches(record.snapshot)) throw new Error("原登记结算尚未核实");
            this.identicalRegistrationVerifiedId = record.operationId;
            return true;
        } catch (error) {
            settings.runAllowed = false; settings.enabled = false; settings.firstSyncCompleted = false;
            record.phase = "RECONCILE_REQUIRED";
            try { await this.saveIdenticalRegistration(record, false); } catch { /* no writes into a changed physical root */ }
            throw error;
        } finally { this.identicalRegistrationInFlight = false; }
    }

    private async performIdenticalRegistrationPreview(cloudVaultId: string): Promise<boolean> {
        const preview = await this.prepareIdenticalOneDriveRegistration(cloudVaultId);
        return await new Promise<boolean>((resolve) => {
            const modal = new Modal(this.app); let decided = false;
            modal.onOpen = () => {
                modal.titleEl.setText(localize("高级：仅登记完全一致的内容"));
                modal.contentEl.createEl("p", { text: localize("账号：") + preview.accountId + localize("；本地库：") + preview.localVaultId +
                    localize("；云库：") + preview.cloudVaultId + localize("；设备：") + preview.deviceId });
                modal.contentEl.createEl("p", { text: localize("以下全范围正文已一致；只保存本机绑定/基线，不上传、落地下载、覆盖或删除正文，不开启同步许可。") });
                modal.contentEl.createEl("p", { text: localize("本次同内容：") + preview.files.length +
                    localize(" 项；正文上传：0；落地下载：0；删除：0；所有项目仍须本次独立确认后才登记。") });
                for (const file of preview.files) modal.contentEl.createEl("p", { text: file.path +
                    localize("；大小：") + file.localSize + localize("字节；本地/云端SHA：") + file.localHash +
                    localize("；远端ID：") + file.remoteId + "；ETag：" + file.remoteEtag });
                modal.contentEl.createEl("button", { text: localize("取消，不登记") }).onclick = () => modal.close();
                const confirm = modal.contentEl.createEl("button", { text: localize("独立确认，仅登记上述同内容基线") });
                confirm.onclick = async () => { if (decided) return; decided = true; confirm.disabled = true;
                    try { resolve(await this.confirmIdenticalOneDriveRegistration(preview.ticket, true)); }
                    catch (error) { new Notice(localizeMessage(safeOneDriveError(error)), 8000); resolve(false); }
                    finally { modal.close(); }
                };
            };
            modal.onClose = () => { modal.contentEl.empty(); if (!decided) {
                decided = true; void this.confirmIdenticalOneDriveRegistration(preview.ticket, false); resolve(false);
            } };
            modal.open();
        });
    }

    private sharedSchema = "smartmd-shared-dec07-v1";
    private sharedReconciliationInFlight = false;

    getOneDriveSettingsSharingPreview(): {
        schema: string; scope: string; accountId: string; vaultId: string;
        localVaultId: string; deviceId: string; fields: SharedSmartMDState;
    } {
        const settings = this.settings();
        const accountId = String((this.store as any).settings.local.microsoftAccountId || "");
        const localVaultId = String(settings.localVaultId || settings.vaultId || "");
        const fields = sharedStateForSync((this.store as any).settings.shared);
        return {
            schema: this.sharedSchema, accountId, vaultId: settings.vaultId,
            localVaultId, deviceId: settings.deviceId, fields,
            scope: JSON.stringify([this.sharedSchema, accountId, settings.vaultId,
                localVaultId, settings.deviceId, fields]),
        };
    }

    hasOneDriveSettingsSharingApproval(): boolean {
        const settings = this.settings();
        return settings.sharedSettingsApproved === true && settings.sharedUploadApproved === true &&
            !this.sharedReconciliationInFlight && !settings.pendingSharedUploadOperation && settings.sharedApprovalSchema === this.sharedSchema &&
            settings.sharedApprovalScope === this.getOneDriveSettingsSharingPreview().scope;
    }

    async approveOneDriveSettingsSharing(approved = false, previewScope?: string): Promise<void> {
        if (!approved || !this.canRun() || this.running) {
            throw new Error("必须在当前设备、账号和库身份确认后单独批准四字段共享");
        }
        const settings = this.settings();
        if (this.sharedReconciliationInFlight || (settings.pendingSharedUploadOperation &&
            settings.pendingSharedUploadOperation.phase !== "FAILED")) {
            throw new Error("原共享操作结果尚未只读核实，不得以新批准重发");
        }
        const preview = this.getOneDriveSettingsSharingPreview();
        if (previewScope !== undefined && previewScope !== preview.scope) {
            throw new Error("共享字段或身份已变化，必须重新预览确认");
        }
        if (settings.pendingSharedUploadOperation?.phase === "FAILED") {
            settings.lastSharedUploadOperation = { ...settings.pendingSharedUploadOperation };
            settings.pendingSharedUploadOperation = null;
        }
        settings.sharedSettingsApproved = true;
        settings.sharedUploadApproved = true;
        settings.sharedApprovalAccount = preview.accountId;
        settings.sharedApprovalVault = preview.vaultId;
        settings.sharedApprovalDevice = preview.deviceId;
        settings.sharedApprovalSchema = preview.schema;
        settings.sharedApprovalScope = preview.scope;
        settings.sharedApprovalId = "shared-" + crypto.randomUUID();
        try { await (this.store as any).save(); }
        catch (error) { settings.sharedUploadApproved = false; throw error; }
        if (preview.scope !== this.getOneDriveSettingsSharingPreview().scope) {
            settings.sharedUploadApproved = false;
            throw new Error("共享批准保存期间字段或身份变化，必须重新预览");
        }
    }

    /** Read only: never retries an upload or deletes a retained staging copy. */
    async reconcileOneDriveSharedUpload(): Promise<boolean> {
        if (this.sharedReconciliationInFlight) throw new Error("原共享操作正在只读核验");
        this.sharedReconciliationInFlight = true;
        try { return await this.reconcileSharedUploadRecord(); }
        finally { this.sharedReconciliationInFlight = false; }
    }

    private async reconcileSharedUploadRecord(): Promise<boolean> {
        const settings = this.settings();
        const record = settings.pendingSharedUploadOperation;
        if (!record) return true;
        const sameIdentity = () => {
            const now = this.getOneDriveSettingsSharingPreview();
            return settings === this.settings() && settings.pendingSharedUploadOperation === record &&
                record.accountId === now.accountId && record.vaultId === now.vaultId &&
                record.localVaultId === now.localVaultId && record.deviceId === now.deviceId &&
                record.schema === now.schema;
        };
        if (!sameIdentity()) return false;
        const meta = await this.readRemoteFileMetadata(record.path, true);
        if (!meta?.id || !meta.etag || !sameIdentity() ||
            (record.finalId && (meta.id !== record.finalId || meta.etag !== record.finalEtag))) return false;
        const body = await this.remoteBytes(meta);
        if (!sameIdentity() || await this.sha256Bytes(body) !== record.bodyHash) return false;
        const after = await this.readRemoteFileMetadata(record.path, true);
        if (!sameIdentity() || after?.id !== meta.id || after?.etag !== meta.etag) return false;
        // Matching remote bytes settle the original operation, never apply cloud settings locally.
        const previousHash = settings.sharedStateHash;
        const previousReceipt = settings.lastSharedUploadOperation;
        settings.sharedStateHash = record.fieldsHash;
        settings.lastSharedUploadOperation = { ...record, phase: "VERIFIED", finalId: meta.id,
            finalEtag: meta.etag, stagingRetained: record.phase !== "CLEANED", verifiedAt: Date.now() };
        settings.pendingSharedUploadOperation = null;
        try { await (this.store as any).save(); }
        catch (error) {
            settings.sharedStateHash = previousHash;
            settings.pendingSharedUploadOperation = record;
            if (previousReceipt === undefined) delete settings.lastSharedUploadOperation;
            else settings.lastSharedUploadOperation = previousReceipt;
            throw error;
        }
        return true;
    }

    getPendingOneDriveIncomingSharedSettings(): {
        accountId: string; vaultId: string; fields: SharedSmartMDState; expiresAt: number;
    } | null {
        const pending = this.pendingSharedImport;
        if (!pending || pending.expiresAt <= Date.now()) return null;
        const localAccount = String((this.store as any).settings.local.microsoftAccountId || "");
        const oneDrive = this.settings();
        if (pending.accountId !== localAccount ||
            pending.vaultId !== oneDrive.vaultId ||
            pending.deviceId !== oneDrive.deviceId ||
            oneDrive.sharedSettingsApproved !== true ||
            oneDrive.sharedApprovalAccount !== localAccount ||
            oneDrive.sharedApprovalVault !== oneDrive.vaultId ||
            oneDrive.sharedApprovalDevice !== oneDrive.deviceId) return null;
        return {
            accountId: pending.accountId, vaultId: pending.vaultId,
            fields: { ...pending.fields }, expiresAt: pending.expiresAt,
        };
    }

    async confirmPendingOneDriveIncomingSharedSettings(approved = false): Promise<boolean> {
        const pending = this.pendingSharedImport;
        // Single-use decision: a stale or failed approval is never replayable.
        this.pendingSharedImport = null;
        if (!approved) return false;
        const oneDrive = this.settings();
        const accountId = String((this.store as any).settings.local.microsoftAccountId || "");
        if (!pending || pending.expiresAt <= Date.now() || this.running ||
            !this.canRun() || oneDrive.sharedSettingsApproved !== true ||
            accountId !== pending.accountId || oneDrive.vaultId !== pending.vaultId ||
            oneDrive.deviceId !== pending.deviceId ||
            oneDrive.sharedApprovalAccount !== accountId ||
            oneDrive.sharedApprovalVault !== oneDrive.vaultId ||
            oneDrive.sharedApprovalDevice !== oneDrive.deviceId ||
            String(oneDrive.sharedStateHash || "") !== pending.baselineHash) {
            throw new Error("共享设置接收确认已失效，必须重新预览");
        }
        const settings = (this.store as any).settings;
        const currentHash = await contentHash(JSON.stringify(sharedStateForSync(settings.shared)));
        if (currentHash !== pending.localHash) throw new Error("本机共享设置已变化，拒绝旧确认");
        const remote = await this.readSharedState("vaults/" + pending.vaultId + "/plugin-state.json");
        if (!remote || await contentHash(JSON.stringify(remote)) !== pending.fullPayloadHash ||
            await contentHash(JSON.stringify(sharedStateForSync(remote.shared))) !== pending.remoteHash ||
            String((this.store as any).settings.local.microsoftAccountId || "") !== pending.accountId ||
            oneDrive.vaultId !== pending.vaultId || oneDrive.deviceId !== pending.deviceId ||
            oneDrive.sharedSettingsApproved !== true ||
            oneDrive.sharedApprovalAccount !== pending.accountId ||
            oneDrive.sharedApprovalVault !== pending.vaultId ||
            oneDrive.sharedApprovalDevice !== pending.deviceId ||
            String(oneDrive.sharedStateHash || "") !== pending.baselineHash ||
            await contentHash(JSON.stringify(sharedStateForSync(settings.shared))) !== pending.localHash ||
            !this.canRun()) {
            throw new Error("远端共享设置或当前身份已变化，必须重新预览");
        }
        const before = settings.shared;
        const beforeHash = oneDrive.sharedStateHash;
        settings.shared = { ...before, ...sharedStateForSync(remote.shared) };
        oneDrive.sharedStateHash = pending.remoteHash;
        try {
            await (this.store as any).save();
        } catch (error) {
            settings.shared = before;
            oneDrive.sharedStateHash = beforeHash;
            throw error;
        }
        return true;
    }

    statusText(): string {
        const settings = this.settings();
        if (settings.status === "stopping") return "正在停止，等待当前提交完成并核对状态";
        if (this.running || settings.status === "syncing") return "同步中";
        if (this.firstUploadInspection) return this.firstUploadInspection.label;
        if (this.preparing) return this.firstConnectionProgress || (Platform.isAndroidApp
            ? "正在查找 OneDrive 知识库…" : "正在连接 OneDrive 并检查同步范围");
        if (this.firstConnectionError) return this.firstConnectionError;
        if (this.retryAt > Date.now() && settings.enabled && this.canRun()) {
            return localize("同步将在 {0} 自动继续。", new Date(this.retryAt).toLocaleTimeString());
        }
        if (settings.status === "partial_success") return localizeMessage(settings.lastError) || localize("部分同步完成（存在未同步文件）");
        if (settings.status === "failed") return localize("同步失败：{0}", localizeMessage(settings.lastError) || localize("未知原因"));
        if (settings.status === "paused") return settings.lastError ? localize("同步未完成：{0}", localizeMessage(settings.lastError)) : "已暂停";
        if (settings.lastSuccessAt) return `上次成功：${new Date(settings.lastSuccessAt).toLocaleString()}`;
        return "尚未完成首次同步";
    }


    isBusy(): boolean {
        return Boolean(this.operationJob) || this.preparing || this.running || this.settings().status === "stopping";
    }

    canCancelFirstUploadInspection(): boolean {
        return this.firstUploadInspection !== null && !this.firstUploadInspection.controller.signal.aborted;
    }

    cancelFirstUploadInspection(): void {
        this.firstUploadInspection?.controller.abort(new Error(localize("已取消 OneDrive 文件检查，可以重试。")));
    }

    actionLabel(): string {
        const status = this.settings().status;
        if (status === "stopping") return "正在停止…";
        if (status === "syncing" || this.running) return "同步中…";
        if (this.preparing) return "检查中…";
        if (Platform.isAndroidApp && !this.settings().firstSyncCompleted) return "开始同步";
        return "立即同步";
    }

    private statusCallback: ((statusText: string) => void) | null = null;

    setStatusCallback(callback: (statusText: string) => void): void {
        this.statusCallback = callback;
    }

    autoSyncWarning(): string {
        const settings = this.settings();
        return settings.enabled
            && settings.firstSyncCompleted
            && !settings.onStartup
            && !settings.afterSave
            && Number(settings.intervalMinutes || 0) === 0
            ? "自动同步已开启，但没有启用任何触发项；当前不会自动同步。"
            : "";
    }

    canRun(): boolean {
        if (this.identicalRegistrationBlocked()) return false;
        const settings = this.settings();
        const local = (this.store as any).settings.local;
        return this.authority.canRun(settings)
            && typeof (this.store as any).installationProofIsValid === "function"
            && (this.store as any).installationProofIsValid() === true
            && settings.confirmationVaultId === settings.vaultId
            && settings.confirmationAccountId === String(local.microsoftAccountId || "");
    }

    canStartManualSync(): boolean {
        return !this.settings().firstSyncCompleted || this.canRun() || this.canResumeManualSync();
    }

    private canResumeManualSync(): boolean {
        const settings = this.settings();
        const accountId = this.store.settings.local.microsoftAccountId;
        const uploads = settings.pendingUploadOperations;
        return this.resumedHistoricalIdentityTrusted() &&
            settings.status !== "stopping" &&
            !this.deletionApprovalRunning && !this.resumedInFlight &&
            !this.uploadLedgerReconciliationInFlight && !this.uploadLedgerSettlementInFlight &&
            !this.unknownSettlementInFlight && !settings.androidFirstDownload &&
            (!uploads || (typeof uploads === "object" && !Array.isArray(uploads))) &&
            !hasUnfinishedOneDriveSetup(settings, accountId, settings.vaultId, settings.deviceId);
    }

    private async resumeManualSync(): Promise<RemoteFile[] | null> {
        if (!this.canResumeManualSync()) return null;
        const settings = this.settings();
        const scope = this.resumedContextKey();
        const generation = this.resumedApprovalGeneration;
        const unchanged = () => this.settings() === settings && this.resumedContextKey() === scope &&
            this.resumedApprovalGeneration === generation;
        await this.assertFirstCommitGraphAccount(settings.confirmationAccountId!);
        const remote = await this.remoteFiles(true, true);
        if (!unchanged() || !this.canResumeManualSync()) return null;
        this.authority.confirm(settings);
        try {
            await this.store.saveWithIdentityGuard(() => {
                if (!unchanged() || !this.canRun()) throw new Error("知识库连接已变化，请再试一次。");
            });
        } catch (error) {
            if (unchanged()) {
                await this.authority.stop(settings);
                settings.status = "paused";
            }
            throw error;
        }
        if (!unchanged() || !this.canRun()) return null;
        this.resumedApprovalGeneration++;
        this.resumedReviewTickets = new WeakMap();
        return remote;
    }

    async manualSync(): Promise<boolean> {
        return await this.runUnifiedOperation(() => this.performManualSync());
    }

    private async performManualSync(): Promise<boolean> {
        const settings = this.settings();
        if (settings.identicalRegistrationJournal && !settings.firstSyncCompleted &&
            !((this.store as any).settings.local.oneDriveScopeIsolation && settings.identicalRegistrationJournal.phase === "COMPLETED")) {
            new Notice(localize("原同内容登记待核实；请先核验原操作，不会新建另一登记或传输正文。"), 8000);
            return false;
        }
        if (settings.firstSyncCompleted) {
            if (this.canRun()) return await this.sync(false);
            try {
                await this.verifyCompletedIdenticalRegistration();
                const remote = await this.resumeManualSync();
                return remote !== null && await this.sync(false, { local: this.localFiles(), allRemote: remote });
            } catch (error: any) {
                if (this.settings() === settings) await this.fail(error);
                return false;
            }
        }
        this.preparing = true;
        this.firstConnectionError = "";
        this.notifyStatusChange();
        const initialAccount = String((this.store as any).settings.local.microsoftAccountId || "");
        const initialGeneration = this.resumedApprovalGeneration;
        const assertFirstGrantCurrent = () => {
            if (initialGeneration !== this.resumedApprovalGeneration || this.settings() !== settings ||
                this.store.settings.local.microsoftAccountId !== initialAccount || !this.store.installationProofIsValid()) {
                throw new Error("首次预览的身份或许可已变化，必须重新预览");
            }
        };
        try {
            if (!initialAccount) throw new Error("请先完成微软账号登录。");
            // First use can open the OneDrive app folder. Creating a cloud
            // vault and transferring files still require the preview action.
            this.vaultRegistryCache = null;
            this.registryReadIsTrusted = false;
            if (settings.firstFileJournal && !settings.firstCommitJournal?.priorBindingArchive) {
                return await this.resumeFirstFileCopy(assertFirstGrantCurrent, initialGeneration);
            }
            if (settings.firstCommitJournal && !settings.firstCommitJournal.priorBindingArchive) {
                return await this.resumeRegisteredFirstCopy(assertFirstGrantCurrent, initialGeneration);
            }
            if (!(await this.prepareVaultIdentity())) return false;
            const choice = this.pendingFirstChoice;
            if (!choice) return false;
            this.previewVaultId = choice.id;
            if (this.identicalRegistrationRequested) return await this.performIdenticalRegistrationPreview(choice.id);
            const mergeAndroid = choice.mergeAndroid === true;
            if (mergeAndroid) {
                if (!this.androidFirstMergeIsTrusted(initialAccount, choice.id) && !this.canStartAndroidFirstMerge(initialAccount)) {
                    throw new Error(localize("当前知识库已有同步历史或身份已变化，请重新核对。"));
                }
            } else await this.assertFirstReceiverIdentity(choice, initialAccount);
            const prepared = await this.captureFirstConnectionSnapshot();
            if (String((this.store as any).settings.local.microsoftAccountId || "") !== initialAccount) {
                throw new Error("微软账号已切换，请重新检查。");
            }
            const local = prepared.local;
            const allRemote = prepared.remote;
            const resumingReceiver = choice.mode === "existing" && !mergeAndroid &&
                this.incompleteFirstReceiverScope(initialAccount, choice.id);
            if (resumingReceiver) {
                await this.assertIncompleteFirstReceiverFiles(
                    initialAccount, choice.id, local, allRemote);
            } else if (choice.mode === "existing" && !mergeAndroid &&
                (local.length > 0 || Object.keys(settings.fileStates || {}).length > 0 ||
                 ((this.store as any).settings.local.recoveryEntries || []).length > 0)) {
                throw new Error("云端是另一份独立知识库。请在独立且严格验空的 Vault 中连接；已接收过一部分的库必须有可信原绑定与历史记录，禁止未经核对直接覆盖。");
            }
            const remote = allRemote.filter((item) => !this.isIgnored(item.path) && item.size <= this.maxBytes());
            const localPaths = new Set(local.map((item) => item.path));
            const remoteMap = new Map(remote.map((item) => [item.path, item]));
            const uploadPaths = local.filter((item) => !remoteMap.has(item.path)).map((item) => item.path).sort();
            const downloadPaths = remote.filter((item) => !localPaths.has(item.path)).map((item) => item.path).sort();
            const localFingerprints = new Map<string, string>(
                prepared.snapshot.localVersions.map((value) => {
                    const [path, , , sha256] = JSON.parse(value) as [string, number, number, string];
                    return [path, sha256];
                }),
            );
            const samePaths: string[] = [];
            const unknownPaths: string[] = [];
            const plannedRemoteHashes = new Map<string, string>();
            for (const item of local) {
                const counterpart = remoteMap.get(item.path);
                if (!counterpart) continue;
                try {
                    const localContent = mergeAndroid
                        ? await this.app.vault.adapter.readBinary(item.path)
                        : await this.app.vault.readBinary(item.file);
                    if (mergeAndroid && await this.sha256Bytes(localContent) !== localFingerprints.get(item.path)) {
                        throw new Error(localize("本机文件在同步时变化，请重新同步。"));
                    }
                    const remoteContent = await this.remoteBytes(counterpart);
                    if (remoteContent.byteLength !== counterpart.size) throw new Error("OneDrive file size changed");
                    if (mergeAndroid) plannedRemoteHashes.set(item.path, await this.sha256Bytes(remoteContent));
                    if (sameBytes(localContent, remoteContent)) samePaths.push(item.path);
                    else if (await this.oneDriveVersionWins(item.path, item.mtime, counterpart)) downloadPaths.push(item.path);
                    else uploadPaths.push(item.path);
                } catch {
                    // A failed content read is UNKNOWN, never "same" or a harmless zero-difference.
                    unknownPaths.push(item.path);
                }
            }
            const rawLocal = this.app.vault.getFiles();
            const skippedPaths = [...new Set([
                ...rawLocal.filter((file) => this.isIgnored(file.path) ||
                    file.stat.size > this.maxBytes()).map((file) => file.path),
                ...allRemote.filter((item) => this.isIgnored(item.path) || item.size > this.maxBytes())
                    .map((item) => item.path),
            ])].sort();
            const rawLocalMap = new Map(rawLocal.map((file) => [file.path, file]));
            const allRemoteMap = new Map(allRemote.map((item) => [item.path, item]));
            // Reuse the already byte-verified snapshot: do not read excluded
            // configuration, recovery content or oversized bodies for display.
            const androidMerge: PreparedSync["androidMerge"] = mergeAndroid ? new Map() : undefined;
            if (androidMerge) {
                const skipped = new Set(skippedPaths);
                const exactPaths = new Map<string, { path: string; isFile: boolean; included: boolean }>();
                const inventory = [
                    ...(prepared.snapshot.localStructure || []).map(entry => {
                        const [path, isFile] = JSON.parse(entry) as [string, boolean];
                        return { path, isFile, included: isFile && localPaths.has(path) && !skipped.has(path) };
                    }),
                    ...allRemote.map(file => ({
                        path: file.path, isFile: true,
                        included: remoteMap.has(file.path) && !skipped.has(file.path),
                    })),
                    ...(prepared.snapshot.remoteFolders || []).map(path => ({ path, isFile: false, included: false })),
                ];
                // Register participating paths first. Excluded entries still
                // describe actual directories and occupied names.
                for (const entry of inventory.filter(item => item.included).concat(inventory.filter(item => !item.included))) {
                    const parts = entry.path.split("/");
                    let prefix = "";
                    for (const [index, part] of parts.entries()) {
                        prefix = prefix ? prefix + "/" + part : part;
                        const isFile = entry.isFile && index === parts.length - 1;
                        const key = prefix.normalize("NFC").toLocaleLowerCase("en-US");
                        const existing = exactPaths.get(key);
                        if (existing && (existing.included || entry.included) &&
                            (existing.path !== prefix || existing.isFile !== isFile)) {
                            throw new Error(localize("两端文件或文件夹名称不兼容，请先核对：{0}", prefix));
                        }
                        if (!existing) exactPaths.set(key, { path: prefix, isFile, included: entry.included });
                    }
                }
                // The first round adds files in either direction; it cannot infer a
                // deletion for a missing entry in a previously verified baseline.
                for (const path of Object.keys(settings.fileStates || {})) {
                    if (!skipped.has(path) && (!localPaths.has(path) || !remoteMap.has(path))) {
                        throw new Error(localize("首次连接中已同步的文件一端缺失，请核对后重试。"));
                    }
                }
                for (const [action, paths] of [["upload", uploadPaths], ["download", downloadPaths], ["same", samePaths]] as const) {
                    for (const path of paths) {
                        if (skipped.has(path)) continue;
                        let remoteHash = plannedRemoteHashes.get(path);
                        if (action === "download" && !remoteHash) {
                            try {
                                const peer = remoteMap.get(path)!;
                                const content = await this.remoteBytes(peer);
                                if (content.byteLength !== peer.size) throw new Error("OneDrive file size changed");
                                remoteHash = await this.sha256Bytes(content);
                            } catch {
                                unknownPaths.push(path);
                                continue;
                            }
                        }
                        androidMerge.set(path, { action, localHash: localFingerprints.get(path), remoteHash });
                    }
                    // A path excluded on either side should not appear in an approved transfer.
                    for (let i = paths.length - 1; i >= 0; i--) {
                        if (skipped.has(paths[i])) paths.splice(i, 1);
                    }
                }
            }
            const items: Record<string, SyncPreviewItem> = {};
            for (const path of new Set([...localPaths, ...allRemoteMap.keys(), ...skippedPaths])) {
                const raw = rawLocalMap.get(path);
                const captured = local.find((entry) => entry.path === path);
                const cloud = allRemoteMap.get(path);
                const reasons: string[] = [];
                if (this.isIgnored(path)) {
                    const directory = path.split("/").find((part) =>
                        [".obsidian", ".trash", ".git", ".suite-recovery", ".suite-staging",
                         "node_modules", "_suite_staging"].includes(part));
                    reasons.push(directory ? "排除目录：" + directory : "命中当前排除规则或恢复副本保护");
                }
                if (raw && raw.stat.size > this.maxBytes()) reasons.push(
                    `本机文件超过单文件上限：${raw.stat.size} > ${this.maxBytes()} 字节`);
                if (cloud && cloud.size > this.maxBytes()) reasons.push(
                    `云端文件超过单文件上限：${cloud.size} > ${this.maxBytes()} 字节`);
                items[path] = {
                    source: raw && cloud ? "本机与云端" : raw ? "本机" : "云端",
                    local: raw ? { size: captured?.size ?? raw.stat.size,
                        mtime: captured?.mtime ?? raw.stat.mtime,
                        sha256: localFingerprints.get(path) } : undefined,
                    remote: cloud ? { id: cloud.id, etag: cloud.etag, size: cloud.size } : undefined,
                    skipReasons: reasons,
                };
            }
            const preview: SyncPreview = {
                localCount: local.length,
                remoteCount: remote.length,
                uploadCount: uploadPaths.length,
                downloadCount: downloadPaths.length,
                skippedCount: skippedPaths.length,
                uploadPaths, downloadPaths, samePaths: samePaths.sort(),
                skippedPaths,
                unknownPaths: unknownPaths.sort(),
                unexpectedNewTarget: choice.mode === "new" && allRemote.length > 0,
                identity: { accountId: initialAccount,
                    localVaultId: prepared.snapshot.localVaultId, cloudVaultId: choice.id },
                androidCloudVaultName: Platform.isAndroidApp
                    ? androidVaultDisplayName(choice, this.vaultRegistryCache ?? []) : undefined,
                creatingCloudVault: choice.mode === "new",
                items,
                priorHistoryCount: choice.mode === "new" ? Object.keys(settings.fileStates || {}).length : 0,
            };
            // User confirmation advances read-only preview to initial sync execution
            this.firstConnectionNotice?.hide();
            await new Promise<void>((resolve) => {
                new SyncPreviewModal(
                    this.app, preview,
                    async (reportProgress) => {
                        assertFirstGrantCurrent();
                        await this.assertFirstConnectionSnapshotFresh(prepared.snapshot);
                        if (mergeAndroid) {
                            if (!this.androidFirstMergeIsTrusted(initialAccount, choice.id) && !this.canStartAndroidFirstMerge(initialAccount)) {
                                throw new Error(localize("当前知识库已有同步历史或身份已变化，请重新核对。"));
                            }
                        } else if (choice.mode === "existing") {
                            await this.assertFirstReceiverIdentity(choice, initialAccount);
                            if (resumingReceiver) {
                                if (preview.uploadCount ||
                                    preview.unknownPaths.length) {
                                    throw new Error("首次接收续传还有未经批准的上传或尚未核实的文件");
                                }
                                await this.assertIncompleteFirstReceiverFiles(
                                    initialAccount, choice.id, local, allRemote, true);
                                if (Platform.isAndroidApp) {
                                    await this.settleAndroidFirstDownload(local, allRemote, assertFirstGrantCurrent);
                                }
                            }
                        }
                        if (choice.mode === "new" && allRemote.length > 0) {
                            throw new Error("拟新建云端目录已有内容，禁止按全新副本提交；请重新检查。");
                        }
                        if (choice.mode === "new") {
                            let approvedSnapshot = prepared.snapshot;
                            if (choice.emptyAppRoot) {
                                const registryHash = await this.bootstrapEmptyAppFolderRegistry(
                                    initialAccount, choice.emptyAppRoot, assertFirstGrantCurrent);
                                choice.emptyAppRoot = undefined;
                                await this.readVaultRegistry();
                                approvedSnapshot = { ...prepared.snapshot, registryHash };
                                await this.assertFirstConnectionSnapshotFresh(approvedSnapshot);
                            }
                            // Four stage registration -> version-pinned upload ->
                            // per-byte verification -> durable binding. Never use
                            // the generic sync(true) shortcut before registry.
                            const meta = await this.readRemoteFileMetadata("vaults/registry.json", true);
                            const registry = await this.readSharedState("vaults/registry.json");
                            if (!meta?.etag || !Array.isArray(registry?.vaults)) {
                                throw new Error("云端注册表版本不可信，禁止创建新库");
                            }
                            const ticket = await this.stageFirstConnectionCommit({
                                accountId: initialAccount,
                                localVaultId: settings.vaultId,
                                deviceId: settings.deviceId,
                                cloudVaultId: choice.id,
                                cloudVaultName: choice.name,
                                mode: "create_new",
                                registryEtag: meta.etag,
                                registryVaults: registry.vaults,
                                snapshotDigest: await contentHash(JSON.stringify(approvedSnapshot)),
                                uploads: local.map((entry) => entry.path).sort(),
                                downloads: [], conflicts: [], unknown: [],
                                expiresAt: Date.now() + 5 * 60_000,
                            }, initialGeneration);
                            assertFirstGrantCurrent();
                            const consent = this.approveFirstConnection(
                                ticket.ticketId, ticket.digest, true);
                            const registered = await this.commitFirstConnection(ticket, consent);
                            if (registered.status !== "REGISTERED") {
                                throw new Error(registered.reason || "云端独立库登记尚未确认");
                            }
                            assertFirstGrantCurrent();
                            reportProgress(localize("云端独立库已登记，正在校验每篇上传计划"));
                            const filesTicket = await this.prepareFirstFileTransfer(prepared.snapshot.localVersions, initialGeneration);
                            assertFirstGrantCurrent();
                            const filesConsent = this.approveFirstFileTransfer(
                                filesTicket.ticketId, filesTicket.digest, true);
                            const copied = await this.executeFirstFileTransfer(
                                filesTicket, filesConsent);
                            if (copied.status !== "FILES_COPIED") {
                                throw new Error(copied.reason || "文件事务需要只读对账");
                            }
                            assertFirstGrantCurrent();
                            // saveVerifiedFirstFileBinding has already
                            // verified the exact copied files and stored the
                            // local/cloud distinct identity pair.
                            await this.completeFirstFileCopy(choice.id, choice.name, initialAccount, assertFirstGrantCurrent);
                            return true;
                        }
                        if (choice.mode !== "existing" ||
                            (!mergeAndroid && !resumingReceiver && (this.localFiles().length ||
                                Object.keys(settings.fileStates || {}).length))) {
                            throw new Error("不具备可信独立空库或已核验的首次接收续传条件");
                        }
                        if (Platform.isAndroidApp) assertFirstGrantCurrent();
                        const beforeBinding = {
                            confirmed: settings.confirmed, runAllowed: settings.runAllowed,
                            vaultId: settings.vaultId, confirmationVaultId: settings.confirmationVaultId,
                            confirmationAccountId: settings.confirmationAccountId, vaultName: settings.vaultName,
                            androidFirstMerge: settings.androidFirstMerge,
                        };
                        this.authority.confirm(settings);
                        settings.confirmed = true;
                        settings.runAllowed = true;
                        settings.vaultId = choice.id;
                        settings.confirmationVaultId = choice.id;
                        settings.confirmationAccountId = initialAccount;
                        settings.vaultName = choice.name;
                        if (mergeAndroid) settings.androidFirstMerge = {
                            accountId: initialAccount, cloudVaultId: choice.id,
                            proofId: settings.installationProofId!, deviceId: settings.deviceId,
                        };
                        try { await this.store.saveWithIdentityGuard(assertFirstGrantCurrent); }
                        catch (error) {
                            if (this.settings() === settings && this.resumedApprovalGeneration === initialGeneration) {
                                Object.assign(settings, beforeBinding);
                                settings.runAllowed = false;
                            }
                            throw error;
                        }
                        this.activeFirstReceiverResume = resumingReceiver;
                        this.setActiveFirstSyncTicket("manual-first-sync-" + Date.now());
                        return await this.sync(true, {
                            local, allRemote, androidMerge,
                        }, reportProgress);
                    },
                    resolve,
                    true,
                ).open();
            });
            return Boolean(settings.firstSyncCompleted);
        } catch (error: any) {
            this.firstConnectionNotice?.hide();
            this.firstConnectionError = localize("OneDrive 首次连接已停止：") + localizeMessage(safeOneDriveError(error));
            new Notice(this.firstConnectionError, 8000);
            return false;
        } finally {
            this.activeFirstReceiverResume = false;
            this.pendingFirstChoice = null;
            this.previewVaultId = null;
            this.firstConnectionProgress = "";
            this.firstConnectionNotice?.hide();
            this.firstConnectionNotice = null;
            this.preparing = false;
            this.notifyStatusChange();
        }
    }

    scheduleAfterSave(): void {
        this.markSyncableContentChanged();
        const settings = this.settings();
        if (!this.canRun() || !settings.enabled || !settings.afterSave || !settings.firstSyncCompleted) return;
        window.clearTimeout(this.saveTimer);
        const delay = settings.autoMode === "smart" ? 10000 : 2500;
        this.saveTimer = window.setTimeout(() => void this.requestSync(), delay);
    }

    async requestSync(): Promise<boolean> {
        const s = this.settings();
        // Scheduled, save and follow-up triggers must not run after the user
        // switches back to manual, even if a timer was already queued.
        if (!this.canRun() || !s.enabled || s.autoMode === "manual" ||
            !s.firstSyncCompleted || Date.now() < this.retryAt) return false;
        return await this.runUnifiedOperation(() => this.sync(false, null, null, true));
    }

    private isTemporarySyncFailure(error: any): boolean {
        return [408, 429, 500, 502, 503, 504].includes(Number(error?.status)) ||
            /\b(?:ERR_INTERNET_DISCONNECTED|ERR_NETWORK_CHANGED|ERR_CONNECTION_(?:RESET|CLOSED|REFUSED)|ERR_(?:NAME_NOT_RESOLVED|TIMED_OUT)|ETIMEDOUT|ECONNRESET|Failed to fetch)\b/i
                .test(String(error?.cause?.message || error?.message || error));
    }

    private scheduleBackgroundRetry(retryAfterMs = 0): void {
        const settings = this.settings();
        if (!settings.enabled || settings.autoMode === "manual" || !settings.firstSyncCompleted || !this.canRun()) return;
        const generation = this.autoFollowUpGeneration;
        const scope = this.contextKey();
        window.clearTimeout(this.retryTimer);
        const backoff = Math.min(15_000 * 2 ** this.retryCount, 300_000);
        const delay = Math.min(Math.max(backoff, retryAfterMs), 2_147_483_647);
        this.retryCount = Math.min(this.retryCount + 1, 5);
        this.retryAt = Date.now() + delay;
        this.retryTimer = window.setTimeout(() => {
            this.retryAt = 0;
            if (generation !== this.autoFollowUpGeneration || settings !== this.settings() ||
                scope !== this.contextKey()) return;
            void this.requestSync().catch(console.error);
        }, delay);
    }

    private clearBackgroundRetry(): void {
        window.clearTimeout(this.retryTimer);
        this.retryAt = 0;
        this.retryCount = 0;
    }

    private sharedStateFingerprint(): string {
        const shared = sharedStateForSync((this.store as any).settings.shared) as any;
        delete shared.updatedAt;
        return JSON.stringify(shared);
    }

    private observeSharedStateChange(): void {
        const next = this.sharedStateFingerprint();
        if (!this.observedSharedState) {
            this.observedSharedState = next;
            return;
        }
        if (next !== this.observedSharedState) {
            this.observedSharedState = next;
            this.syncableChangeVersion++;
        }
    }

    private markSyncableContentChanged(): void {
        this.syncableChangeVersion++;
        const s = this.settings();
        if (this.operationJob && s.enabled && s.autoMode !== "manual" &&
            this.syncableChangeVersion > this.activeRoundChangeVersion) {
            this.followUpQueued = true;
        }
    }

    // Vault events also fire for the sync engine's own create/modify/restore writes.
    // A queued event is not proof of a new user edit. Compare actual bytes with
    // the committed per-file baseline, so same-size/same-mtime user edits still
    // trigger a follow-up while our own completed writes do not.
    private async hasUnappliedSyncChange(sharedAtStart: string): Promise<boolean> {
        if (this.sharedStateFingerprint() !== sharedAtStart) return true;
        try {
            const settings = this.settings();
            // An unresolved conflict already retained both versions. Vault writes from
            // that preservation must not start another copy/overwrite cycle.
            if (settings.status !== "ready") return true; // Preserve other failure/recovery behavior.
            const states = settings.fileStates || {};
            const local = this.localFiles();
            const paths = new Set(local.map((entry) => entry.path));
            for (const entry of local) {
                const state = states[entry.path];
                if (!state || !state.localHash) return true; // No trusted byte baseline.
                const current = await this.sha256Bytes(await this.app.vault.readBinary(entry.file));
                if (current !== state.localHash) return true;
            }
            for (const [path, state] of Object.entries(states) as [string, any][]) {
                if (!paths.has(path) && !this.isIgnored(path)
                    && Boolean(state?.localMtime)) return true;
            }
            return false;
        } catch {
            // Failure to verify is not evidence that a real user change can be discarded.
            return true;
        }
    }

    private async runUnifiedOperation(primary: () => Promise<boolean>): Promise<boolean> {
        this.observeSharedStateChange();
        const canStartAutomaticFollowUp = (generation: number): boolean => {
            const s = this.settings();
            return generation === this.autoFollowUpGeneration
                && s.enabled && s.autoMode !== "manual"
                && s.firstSyncCompleted && Date.now() >= this.retryAt && this.canRun();
        };
        if (this.operationJob) {
            if (canStartAutomaticFollowUp(this.autoFollowUpGeneration)
                && this.syncableChangeVersion > this.activeRoundChangeVersion) {
                this.followUpQueued = true;
            }
            return await this.operationJob;
        }
        const initialGeneration = this.autoFollowUpGeneration;
        this.activeRoundChangeVersion = this.syncableChangeVersion;
        const sharedAtStart = this.sharedStateFingerprint();
        const job = (async () => {
            let result = await primary();
            if (this.followUpQueued && canStartAutomaticFollowUp(initialGeneration)) {
                this.followUpQueued = false;
                // A switch to manual may happen WHILE the SHA check awaits disk I/O.
                if (await this.hasUnappliedSyncChange(sharedAtStart)
                    && canStartAutomaticFollowUp(initialGeneration)) {
                    this.activeRoundChangeVersion = this.syncableChangeVersion;
                    result = await this.sync(false, null, null, true);
                }
            }
            this.observedSharedState = this.sharedStateFingerprint();
            return result;
        })();
        this.operationJob = job;
        this.notifyStatusChange();
        try {
            return await job;
        } finally {
            if (this.operationJob === job) this.operationJob = null;
            if (this.followUpQueued && canStartAutomaticFollowUp(initialGeneration)) {
                this.followUpQueued = false;
                window.clearTimeout(this.followUpTimer);
                this.followUpTimer = window.setTimeout(() => {
                    void (async () => {
                        if (canStartAutomaticFollowUp(initialGeneration)
                            && await this.hasUnappliedSyncChange(this.observedSharedState)
                            && canStartAutomaticFollowUp(initialGeneration)) {
                            await this.requestSync();
                        }
                    })();
                }, 500);
            }
            this.notifyStatusChange();
        }
    }

    async pause(): Promise<void> {
        const settings = this.settings();
        // Invalidate pending inbound settings immediately, before any await.
        const stopped = this.authority.stop(settings);
        const resumedStopped = this.resumedTransientState
            ? this.authority.stop(this.resumedTransientState) : Promise.resolve();
        this.resumedApprovalGeneration++;
        this.autoFollowUpGeneration++;
        this.followUpQueued = false;
        this.clearBackgroundRetry();
        this.resumedReviewTickets = new WeakMap();
        this.pendingSharedImport = null;
        settings.status = "stopping";
        settings.lastError = "";
        if (typeof window !== "undefined") window.clearTimeout(this.saveTimer);
        if (typeof window !== "undefined") window.clearTimeout(this.followUpTimer);
        await (this.store as any).save();
        this.notifyStatusChange();
        await Promise.all([stopped, resumedStopped]);
        settings.status = "paused";
        await (this.store as any).save();
        this.notifyStatusChange();
    }

    async selectManualSyncMode(): Promise<void> {
        const settings = this.settings();
        // Invalidate both an in-memory follow-up and any asynchronous
        // hasUnappliedSyncChange() still awaiting a byte check.
        this.autoFollowUpGeneration++;
        this.followUpQueued = false;
        this.clearBackgroundRetry();
        // Switching from smart/custom to manual is NOT a global revoke.
        // Keep this session's newly approved manual permission, but cancel
        // queued automatic timers and disable every automatic trigger.
        settings.autoMode = "manual";
        settings.enabled = false;
        settings.onStartup = false;
        settings.afterSave = false;
        settings.intervalMinutes = 0;
        if (typeof window !== "undefined") {
            window.clearTimeout(this.saveTimer);
            window.clearTimeout(this.followUpTimer);
        }
        await (this.store as any).save();
        this.notifyStatusChange();
    }

    async enableAutoSync(): Promise<void> {
        const settings = this.settings();
        if (!settings.firstSyncCompleted) throw new Error("首次手动同步成功后才能启用自动同步。");
        if (this.isBusy()) {
            throw new Error("请等当前同步结束后再切换同步方式。");
        }
        if (!this.canRun()) {
            throw new Error("本轮同步执行许可已失效。请先重新只读核对账号、云库、版本和文件范围并明确批准；切换自动同步方式不能恢复旧许可。");
        }
        const local = (this.store as any).settings.local;
        this.authority.confirm(settings);
        settings.confirmationVaultId = settings.vaultId;
        settings.confirmationAccountId = String(local.microsoftAccountId || "");
        settings.status = "ready";
        settings.lastError = "";
        await (this.store as any).save();
        this.notifyStatusChange();
    }

    dispose(): void {
        this.resumedApprovalGeneration++;
        this.resumedReviewTickets = new WeakMap();
        this.clearBackgroundRetry();
        if (this.resumedTransientState) this.resumedTransientState.runAllowed = false;
        void this.authority.shutdown();
        if (typeof window !== "undefined") window.clearTimeout(this.saveTimer);
        if (typeof window !== "undefined") window.clearTimeout(this.followUpTimer);
    }

    async sync(
        firstSync = false,
        prepared: PreparedSync | null = null,
        reportProgress: ((message: string) => void) | null = null,
        silentSuccess = false,
    ): Promise<boolean> {
        if (this.running || this.deletionApprovalRunning || this.resumedInFlight || this.unknownSettlementInFlight ||
            this.uploadLedgerReconciliationInFlight || this.uploadLedgerSettlementInFlight || !this.canRun()) return false;
        if (firstSync) {
            if (!this.activeFirstSyncTicket) return false;
            this.activeFirstSyncTicket = null;
        }
        const firstReceiverResume = firstSync && this.activeFirstReceiverResume;
        if (!firstSync && !this.settings().firstSyncCompleted) return false;
        const existing = this.settings().pendingUploadOperations;
        if (existing && (typeof existing !== "object" || Array.isArray(existing))) {
            throw new Error("历史上传操作记录格式无效，禁止继续同步");
        }
        if (firstSync && Object.keys(existing || {}).length) return false;
        // A prior round may have durably saved its SHA/ItemID/ETag baseline
        // and confirmed staging cleanup, then stopped before releasing the
        // local journal. Persist the baseline first; release only matching
        // CLEANUP_VERIFIED records, never an UNKNOWN or incomplete operation.
        if (Object.values(existing || {}).some((entry: any) => entry?.phase === "CLEANUP_VERIFIED")) {
            this.uploadLedgerReconciliationInFlight = true;
            try {
                await (this.store as any).save();
                await this.clearVerifiedUploadLedger(this.settings().fileStates || {});
            } finally {
                this.uploadLedgerReconciliationInFlight = false;
            }
        }
        const operation = this.authority.start(this.settings(), this.contextKey());
        this.running = true;
        this.notifyStatusChange();
        const settings = this.settings();
        for (const [path, record] of Object.entries(settings.pendingUploadOperations || {}) as [string, any][]) {
            if (record && typeof record === "object" &&
                settings.lastError === "OneDrive 未确认文件原修改时间 (处理文件: " + redactOneDriveText(path) + ")") {
                record.legacyTimestampFailure = true;
            }
        }
        settings.status = "syncing";
        settings.lastError = "";
        settings.pendingDeletions = settings.pendingDeletions || {};
        for (const pattern of settings.ignorePatterns || []) ignoreMatcher(pattern);
        let currentProcessingPath = "";
        let temporaryFileFailure = false;
        try {
            const rootScope = {
                accountId: settings.confirmationAccountId || "",
                vaultId: settings.vaultId,
                deviceId: settings.deviceId,
            };
            const storedRoot = settings.deletionRoot;
            const storedRootMatches = storedRoot?.accountId === rootScope.accountId &&
                storedRoot.vaultId === rootScope.vaultId && storedRoot.deviceId === rootScope.deviceId;
            let knownDeletionRoot = storedRootMatches ? storedRoot.id : "";
            await (this.store as any).save();
            reportProgress?.(localize("正在读取本地与云端状态…"));
            const local = prepared?.local || this.localFiles();
            let previous: Record<string, SyncFileState> = settings.fileStates || {};
            let allRemote: RemoteFile[];
            if (!firstSync && !knownDeletionRoot && Object.values(previous).some(file => file.remoteId)) {
                // Migrate an existing binding using its file identities, never
                // a mutable listing cache or new files imported on a later run.
                const rootId = await this.readDeletionRoot(operation);
                allRemote = await this.remoteFiles(true, true, undefined, rootId);
                if (!allRemote.some(file => previous[file.path]?.remoteId === file.id) ||
                    await this.readDeletionRoot(operation) !== rootId) {
                    throw new Error(localize("OneDrive 同步目录已变化，已暂缓删除。"));
                }
                knownDeletionRoot = rootId;
                settings.deletionRoot = { ...rootScope, id: rootId };
                await this.store.save();
            } else {
                allRemote = prepared?.allRemote || await this.remoteFiles();
            }
            if (!firstSync && !knownDeletionRoot && !Object.keys(previous).length &&
                !Object.keys(settings.pendingUploadOperations || {}).length) {
                knownDeletionRoot = await this.readDeletionRoot(operation);
                settings.deletionRoot = { ...rootScope, id: knownDeletionRoot };
                await this.store.saveWithIdentityGuard(() => this.assertCurrent(operation));
            }
            if (!firstSync && !prepared && !allRemote.length &&
                Object.values(previous).some(item => Boolean((item as SyncFileState)?.remoteId))) {
                // Confirm an empty delta against a full listing of the bound
                // vault. An empty list never permits deleting a local copy.
                allRemote = await this.remoteFiles(true, true);
            }
            if (!firstSync && Object.keys(settings.pendingUploadOperations || {}).length) {
                const recovery = await this.recoverVerifiedUploads(operation);
                temporaryFileFailure = recovery.retryNeeded;
                if (recovery.refreshed) {
                    const rootId = settings.deletionRoot?.id;
                    allRemote = await this.remoteFiles(true, true, undefined, rootId);
                    if (rootId && await this.readDeletionRoot(operation) !== rootId) {
                        throw new Error(localize("OneDrive 同步目录已变化，已暂缓删除。"));
                    }
                    previous = settings.fileStates || {};
                }
            }
            const skippedPaths = this.skippedLocalPaths();
            for (const item of allRemote) {
                if (this.isIgnored(item.path) || item.size > this.maxBytes()) skippedPaths.add(item.path);
            }
            const remote = allRemote.filter((item) => !skippedPaths.has(item.path) && !this.isIgnored(item.path));

            const localMap = new Map(local.map((item) => [item.path, item]));
            const remoteMap = new Map(remote.map((item) => [item.path, item]));
            const allPaths = new Set([...localMap.keys(), ...remoteMap.keys(), ...Object.keys(previous).filter((path) => !skippedPaths.has(path) && !this.isIgnored(path)), ]);
            // Skipping an ignored or oversized path is not evidence that its
            // historical local/remote file was deleted. Preserve the last
            // trusted baseline so re-enabling it cannot manufacture a fresh
            // untracked download or bypass deletion review.
            const next: Record<string, any> = Object.fromEntries(
                Object.entries(previous).filter(([path]) =>
                    skippedPaths.has(path) || this.isIgnored(path)),
            );

            const firstSyncMatches = new Map<string, boolean>();
            if (firstSync && !prepared?.androidMerge) {
                const candidates = [...allPaths].filter((path) => localMap.has(path) && remoteMap.has(path) && !previous[path]);
                let cursor = 0;
                let completed = 0;
                const isMobile = typeof document !== "undefined" && document.body?.classList.contains("is-mobile");
                const concurrency = isMobile ? 3 : 5;
                const compareNext = async () => {
                    while (cursor < candidates.length) {
                        const path = candidates[cursor++];
                        const localItem = localMap.get(path)!;
                        const remoteItem = remoteMap.get(path)!;

                        // 层次一：大小剪枝（0ms）
                        if (localItem.size !== remoteItem.size) {
                            firstSyncMatches.set(path, false);
                            completed++;
                            reportProgress?.(localize("正在确认同路径文件 {0}/{1}…", completed, candidates.length));
                            continue;
                        }
                        // 层次二：空文件剪枝（0ms）
                        if (localItem.size === 0 && remoteItem.size === 0) {
                            firstSyncMatches.set(path, true);
                            completed++;
                            reportProgress?.(localize("正在确认同路径文件 {0}/{1}…", completed, candidates.length));
                            continue;
                        }

                        const localData = await this.app.vault.readBinary(localItem.file);

                        // 层次三：QuickXorHash 剪枝（0 网络请求）
                        if (remoteItem.quickXorHash) {
                            const localHash = OneDriveSyncService.quickXorHash(localData);
                            if (localHash === remoteItem.quickXorHash) {
                                firstSyncMatches.set(path, true);
                                completed++;
                                reportProgress?.(localize("正在确认同路径文件 {0}/{1}…", completed, candidates.length));
                                continue;
                            }
                        }

                        // 层次四：SHA-1 剪枝（0 网络请求）
                        if (remoteItem.sha1Hash && globalThis.crypto?.subtle) {
                            try {
                                const digest = await globalThis.crypto.subtle.digest("SHA-1", localData);
                                const hex = Array.from(new Uint8Array(digest)).map((b) => b.toString(16).padStart(2, "0")).join("").toUpperCase();
                                if (hex === remoteItem.sha1Hash.toUpperCase()) {
                                    firstSyncMatches.set(path, true);
                                    completed++;
                                    reportProgress?.(localize("正在确认同路径文件 {0}/{1}…", completed, candidates.length));
                                    continue;
                                }
                            } catch {
                                // subtle digest fallback to remoteBytes
                            }
                        }

                        // 退化兜底：网络下载字节比对
                        const remoteData = await this.remoteBytes(remoteItem);
                        firstSyncMatches.set(path, sameBytes(localData, remoteData));
                        completed++;
                        reportProgress?.(localize("正在确认同路径文件 {0}/{1}…", completed, candidates.length));
                    }
                };
                await Promise.all(Array.from(
                    { length: Math.min(concurrency, Math.max(1, candidates.length)) },
                    () => compareNext(),
                ));
            }

            let processingIndex = 0;
            let successCount = 0;
            let unchangedCount = 0;
            let pendingUploadCount = 0;
            let pendingRechecks = 0;
            const startTime = Date.now();
            const failedDiagnostics: SyncFailureDiagnostic[] = [];
            let deletionInventory: Promise<{ rootId: string; files: RemoteFile[] }> | undefined;
            const loadDeletionInventory = () => deletionInventory ??= (async () => {
                const rootId = await this.readDeletionRoot(operation);
                if (knownDeletionRoot && rootId !== knownDeletionRoot) {
                    throw new Error(localize("OneDrive 同步目录已变化，已暂缓删除。"));
                }
                const files = await this.remoteFiles(true, true, undefined, rootId);
                if (await this.readDeletionRoot(operation) !== rootId ||
                    (!knownDeletionRoot && !files.some(file => previous[file.path]?.remoteId === file.id))) {
                    throw new Error(localize("OneDrive 同步目录已变化，已暂缓删除。"));
                }
                settings.deletionRoot = { ...rootScope, id: rootId };
                return { rootId, files };
            })();

            let lastFlushTime = Date.now();
            let itemsSinceLastFlush = 0;
            const flushIfNeeded = async (force = false) => {
                itemsSinceLastFlush++;
                if (!firstSync || Platform.isAndroidApp || force ||
                    itemsSinceLastFlush >= 20 || Date.now() - lastFlushTime >= 3000) {
                    await this.persistFinalState(operation, next);
                    itemsSinceLastFlush = 0;
                    lastFlushTime = Date.now();
                }
            };

            for (const path of [...allPaths].sort()) {
                currentProcessingPath = path;
                this.assertCurrent(operation);
                processingIndex++;
                const progressText = `正在执行同步计划 ${processingIndex}/${allPaths.size}…`;
                reportProgress?.(progressText);
                this.statusCallback?.(`SmartMD: 同步中 (${processingIndex}/${allPaths.size})`);
                const localItem = localMap.get(path);
                const remoteItem = remoteMap.get(path);
                const old = previous[path];
                if (this.hasPendingUploadAt(path)) {
                    if (old) next[path] = old;
                    delete settings.pendingDeletions[path];
                    skippedPaths.add(path);
                    pendingUploadCount++;
                    continue;
                }
                if (localItem && remoteItem) delete settings.pendingDeletions[path];

                try {
                    if (!localItem && !remoteItem) {
                        const interrupted = settings.pendingDeletions[path];
                        if (!firstSync && interrupted?.automatic && old?.remoteId === interrupted.remoteId &&
                            interrupted.accountId === settings.confirmationAccountId && interrupted.vaultId === settings.vaultId &&
                            interrupted.deviceId === settings.deviceId && !await this.app.vault.adapter.exists(path) &&
                            await this.confirmRemoteDeletion(operation, path, old.remoteId, (await loadDeletionInventory()).rootId)) {
                            await this.finishOneDriveDeletion(path);
                            successCount++;
                        } else if (old) next[path] = old;
                        continue;
                    }
                    let attempts = 0;
                    let lastAttemptError: any = null;
                    while (attempts < 2) {
                        attempts++;
                        try {
                            if (prepared?.androidMerge) {
                                const plan = prepared.androidMerge.get(path);
                                if (!plan) break;
                                const localBytes = localItem ? await this.app.vault.adapter.readBinary(path) : null;
                                if (Boolean(localBytes) !== Boolean(plan.localHash) ||
                                    (localBytes && await this.sha256Bytes(localBytes) !== plan.localHash)) {
                                    throw new Error(localize("本机文件在同步时变化，请重新同步。"));
                                }
                                await this.assertFresh(operation, path, localItem, remoteItem);
                                let id = remoteItem?.id, etag = remoteItem?.etag;
                                let localHash = plan.localHash;
                                if (plan.action === "upload" && localBytes) {
                                    if (remoteItem && (!plan.remoteHash ||
                                        await this.sha256Bytes(await this.remoteBytes(remoteItem)) !== plan.remoteHash)) {
                                        throw new Error(localize("云端文件在同步时变化，请重新同步。"));
                                    }
                                    const uploaded = await this.uploadStaged(operation, path, localBytes, false, remoteItem || null);
                                    id = uploaded.id;
                                    etag = uploaded.eTag || uploaded.etag;
                                } else if (plan.action === "download" && remoteItem) {
                                    localHash = await this.download(operation, remoteItem, path, localItem, false,
                                        plan.localHash, plan.remoteHash);
                                }
                                if (plan.action === "same" && (!remoteItem || !plan.remoteHash || plan.localHash !== plan.remoteHash ||
                                    await this.sha256Bytes(await this.remoteBytes(remoteItem)) !== plan.remoteHash)) {
                                    throw new Error(localize("云端文件在同步时变化，请重新同步。"));
                                }
                                const current = this.app.vault.getAbstractFileByPath(path);
                                if (!(current instanceof TFile) || !id || !etag || !localHash ||
                                    await this.sha256Bytes(await this.app.vault.adapter.readBinary(path)) !== localHash) {
                                    throw new Error(localize("本机文件在同步时变化，请重新同步。"));
                                }
                                next[path] = {
                                    localMtime: current.stat.mtime, localSize: current.stat.size, localHash,
                                    remoteId: id, remoteEtag: etag,
                                };
                                // Persist verified unchanged baselines before declaring the binding complete.
                                if (plan.action === "same") unchangedCount++;
                                else successCount++;
                                await flushIfNeeded();
                                break;
                            }
                            if (firstReceiverResume) {
                                if (localItem && remoteItem && old) {
                                    // Resume is DOWNLOAD-ONLY. Do not reinterpret a
                                    // changed timestamp as permission to reupload A.
                                    try {
                                        await this.assertFresh(operation, path, localItem, remoteItem);
                                        const actual = await this.app.vault.readBinary(localItem.file);
                                        if (!old.localHash ||
                                            (await this.sha256Bytes(actual)) !== old.localHash ||
                                            old.remoteId !== remoteItem.id ||
                                            old.remoteEtag !== remoteItem.etag ||
                                            old.localMtime !== localItem.mtime ||
                                            old.localSize !== localItem.size ||
                                            !sameBytes(actual, await this.remoteBytes(remoteItem))) {
                                            throw new Error("FIRST_RECEIVER_VERIFIED_BASELINE_CHANGED");
                                        }
                                    } catch {
                                        const unsafe: any = new Error(
                                            "首次接收续传的已完成文件或云端版本已变化，禁止继续下载或上传");
                                        unsafe.isFatalAuth = true;
                                        throw unsafe;
                                    }
                                    next[path] = { ...old };
                                    unchangedCount++;
                                    lastAttemptError = null;
                                    break;
                                }
                                // Only a previously unreceived cloud file may be
                                // downloaded; any other case is an unsafe upload,
                                // deletion or missing historical baseline.
                                if (localItem || old || !remoteItem) {
                                    const unsafe: any = new Error(
                                        "首次接收续传出现未授权的新增本地文件或历史缺失，拒绝写入");
                                    unsafe.isFatalAuth = true;
                                    throw unsafe;
                                }
                            }
                            if (localItem && remoteItem) {
                                const localBytes = await this.app.vault.readBinary(localItem.file);
                                const currentLocalHash = await this.sha256Bytes(localBytes);
                                const remoteChanged = !old || old.remoteEtag !== remoteItem.etag;
                                // Legacy baselines predate SHA-256. Establish a
                                // real byte-equivalent reference once, rather
                                // than guessing from mtime/size or overwriting.
                                let legacySame = false;
                                if (old && !old.localHash &&
                                    old.localMtime === localItem.mtime &&
                                    old.localSize === localItem.size && !remoteChanged) {
                                    legacySame = sameBytes(localBytes,
                                        await this.remoteBytes(remoteItem));
                                }
                                const localChanged = !old ||
                                    (old.localHash ? old.localHash !== currentLocalHash : !legacySame);
                                if (firstSync && !old && firstSyncMatches.get(path) === true) {
                                    next[path] = {
                                        localHash: currentLocalHash,
                                        localMtime: localItem.mtime,
                                        localSize: localItem.size,
                                        remoteEtag: remoteItem.etag,
                                        remoteId: remoteItem.id,
                                    };
                                    await flushIfNeeded();
                                    unchangedCount++;
                                    lastAttemptError = null;
                                    break;
                                }
                                if (firstSync && old && !localChanged && !remoteChanged) {
                                    next[path] = old;
                                    await flushIfNeeded();
                                    unchangedCount++;
                                    lastAttemptError = null;
                                    break;
                                }
                                if (!localChanged && !remoteChanged) {
                                    next[path] = old;
                                    unchangedCount++;
                                    lastAttemptError = null;
                                    break;
                                }
                                let resultingEtag = remoteItem.etag;
                                let resultingId = remoteItem.id;
                                let resultingHash = currentLocalHash;
                                let takeCloud = remoteChanged && !localChanged;
                                if (localChanged && remoteChanged) {
                                    const peer = await this.remoteBytes(remoteItem);
                                    if (sameBytes(localBytes, peer)) {
                                        await this.assertFresh(operation, path, localItem, remoteItem);
                                        if (!sameBytes(await this.app.vault.readBinary(localItem.file), localBytes)) {
                                            throw new Error(localize("本机文件在同步时变化，请重新同步。"));
                                        }
                                        next[path] = {
                                            localMtime: localItem.mtime, localSize: localItem.size,
                                            remoteEtag: remoteItem.etag, remoteId: remoteItem.id,
                                            localHash: currentLocalHash,
                                        };
                                        await flushIfNeeded();
                                        unchangedCount++;
                                        lastAttemptError = null;
                                        break;
                                    }
                                    takeCloud = (!old?.localHash || await this.sha256Bytes(peer) !== old.localHash)
                                        && await this.oneDriveVersionWins(path, localItem.mtime, remoteItem);
                                }
                                if (takeCloud) {
                                    resultingHash = await this.download(operation, remoteItem, path, localItem, false, currentLocalHash);
                                } else if (localChanged || firstSync) {
                                    const data = await this.app.vault.readBinary(localItem.file);
                                    if (!sameBytes(data, localBytes)) throw new Error(localize("本机文件在同步时变化，请重新同步。"));
                                    await this.assertFresh(operation, path, localItem, remoteItem);
                                    const uploaded = await this.uploadStaged(operation, path, data, false, remoteItem);
                                    resultingEtag = uploaded.eTag || uploaded.cTag || resultingEtag;
                                    resultingId = uploaded.id || resultingId;
                                    resultingHash = await this.sha256Bytes(data);
                                }
                                const current = this.app.vault.getAbstractFileByPath(path) as TFile | null;
                                next[path] = {
                                    localMtime: takeCloud ? current?.stat.mtime || localItem.mtime : localItem.mtime,
                                    localSize: takeCloud ? current?.stat.size || localItem.size : localItem.size,
                                    remoteEtag: resultingEtag,
                                    remoteId: resultingId,
                                    localHash: resultingHash,
                                };
                                await flushIfNeeded();
                                successCount++;
                                lastAttemptError = null;
                                break;
                            }

                            if (localItem && !remoteItem) {
                                if (!firstSync && old?.remoteId) {
                                    if (!old.localHash || !await this.confirmRemoteDeletion(
                                        operation, path, old.remoteId, (await loadDeletionInventory()).rootId)) {
                                        throw new Error(localize("OneDrive 尚未确认此文件已删除。"));
                                    }
                                    const localHash = await this.sha256Bytes(await this.app.vault.readBinary(localItem.file));
                                    if (localHash === old.localHash) {
                                        settings.pendingDeletions[path] = {
                                            accountId: settings.confirmationAccountId,
                                            vaultId: settings.vaultId, deviceId: settings.deviceId,
                                            direction: "vault", remoteId: old.remoteId, remoteEtag: old.remoteEtag,
                                            localMtime: localItem.mtime, localSize: localItem.size,
                                            localHash, requestedAt: Date.now(), automatic: true,
                                        };
                                        await this.applyOneDriveDeletion(path, operation, loadDeletionInventory);
                                        successCount++;
                                        lastAttemptError = null;
                                        break;
                                    }
                                    // An edit made on this device supersedes the deletion.
                                    delete settings.pendingDeletions[path];
                                }
                                const readMtime = localItem.mtime;
                                const data = await this.app.vault.readBinary(localItem.file);
                                let uploaded: any;
                                if (firstSync) {
                                    uploaded = await this.upload(path, data, false, null, operation, undefined, undefined, readMtime);
                                    uploaded = await this.setUploadedModificationTime(uploaded, readMtime, data, operation);
                                } else {
                                    await this.assertFresh(operation, path, localItem, undefined);
                                    uploaded = await this.uploadStaged(operation, path, data, false, null);
                                }
                                next[path] = {
                                    localMtime: readMtime,
                                    localSize: data.byteLength,
                                    remoteEtag: uploaded?.eTag || uploaded?.etag || "",
                                    remoteId: uploaded?.id || "",
                                    localHash: await this.sha256Bytes(data),
                                };
                                await flushIfNeeded();
                                successCount++;
                                lastAttemptError = null;
                                break;
                            }

                            if (!localItem && remoteItem) {
                                if (!firstSync && old?.localMtime && old.remoteId &&
                                    old.remoteId === remoteItem.id && old.remoteEtag && old.remoteEtag === remoteItem.etag) {
                                    settings.pendingDeletions[path] = {
                                        accountId: settings.confirmationAccountId,
                                        vaultId: settings.vaultId, deviceId: settings.deviceId,
                                        direction: "cloud", remoteId: remoteItem.id,
                                        remoteEtag: remoteItem.etag,
                                        requestedAt: Date.now(),
                                        automatic: true,
                                    };
                                    await this.applyOneDriveDeletion(path, operation, loadDeletionInventory);
                                    successCount++;
                                    lastAttemptError = null;
                                    break;
                                }
                                // A changed cloud version takes precedence over local absence.
                                delete settings.pendingDeletions[path];
                                const firstAndroidReceive = firstSync && Platform.isAndroidApp;
                                const localHash = await this.download(operation, remoteItem, path, undefined, firstAndroidReceive);
                                const downloaded = this.app.vault.getAbstractFileByPath(path) as TFile | null;
                                next[path] = {
                                    localMtime: downloaded?.stat.mtime || Date.now(),
                                    localSize: downloaded?.stat.size || remoteItem.size,
                                    remoteEtag: remoteItem.etag,
                                    remoteId: remoteItem.id,
                                    localHash,
                                };
                                if (firstAndroidReceive) {
                                    const intent = settings.androidFirstDownload;
                                    delete settings.androidFirstDownload;
                                    try { await flushIfNeeded(true); }
                                    catch (error) {
                                        settings.androidFirstDownload = intent;
                                        delete settings.fileStates[path];
                                        throw error;
                                    }
                                } else {
                                    await flushIfNeeded();
                                }
                                successCount++;
                                lastAttemptError = null;
                                break;
                            }
                        } catch (err: any) {
                            lastAttemptError = err;
                            if (!err?.isUnknownCommit && !err?.recheckDeletion && !err?.retryAfterMs && err?.status === 503 && attempts < 2) {
                                await new Promise((r) => setTimeout(r, 200));
                                continue;
                            }
                            throw err;
                        }
                    }
                    if (lastAttemptError) throw lastAttemptError;
                    if (next[path] && next[path] !== old && !next[path].localHash) {
                        const current = this.app.vault.getAbstractFileByPath(path);
                        if (current instanceof TFile) {
                            next[path].localHash = await this.sha256Bytes(
                                await this.app.vault.readBinary(current));
                        }
                    }
                } catch (fileError: any) {
                    // These errors can invalidate the cloud Vault scope for
                    // the whole round (permission revoked, root disappeared,
                    // or drive out of quota). Do not try the next session or
                    // treat a second file's success as recovery of the scope.
                    const status = Number(fileError?.status || 0);
                    if (status === 429 || (fileError?.retryAfterMs && this.isTemporarySyncFailure(fileError))) {
                        await flushIfNeeded(true);
                        throw fileError;
                    }
                    const invalidCloudScope = status === 401 || status === 403 || status === 507 ||
                        (status === 404 && !fileError?.isMissingFile && !fileError?.recheckDeletion);
                    if (!this.isCurrent(operation) || settings.androidFirstDownload || fileError?.isFatalAuth ||
                        (fileError?.isUnknownCommit && (firstSync || !this.hasPendingUploadAt(path))) || invalidCloudScope) {
                        throw fileError;
                    }
                    const pendingUpload = this.hasPendingUploadAt(path);
                    temporaryFileFailure ||= pendingUpload || Boolean(fileError?.recheckDeletion || fileError?.recheckFile) ||
                        this.isTemporarySyncFailure(fileError);
                    if (pendingUpload) delete settings.pendingDeletions[path];
                    if (old) {
                        next[path] = old;
                    } else {
                        delete next[path];
                    }
                    if (pendingUpload) {
                        skippedPaths.add(path);
                        pendingUploadCount++;
                    } else if (!firstSync && fileError?.recheckFile) {
                        pendingRechecks++;
                    } else {
                        failedDiagnostics.push(this.parseFileErrorToDiagnostic(path, fileError, localItem ? "UPLOAD" : "DOWNLOAD"));
                    }
                }
            }

            await flushIfNeeded(true);
            this.assertCurrent(operation);

            const durationMs = Date.now() - startTime;
            const failedCount = failedDiagnostics.length;
            const pendingUploads = Object.keys(settings.pendingUploadOperations || {}).length;
            temporaryFileFailure ||= pendingUploads > 0 || pendingRechecks > 0;
            const summary: SyncSummary = {
                total: allPaths.size,
                processed: processingIndex,
                success: successCount,
                failed: failedCount,
                unchanged: unchangedCount,
                pendingUploads,
                pendingRechecks,
                outcome: "incomplete",
                conflicts: 0,
                durationMs,
                finishedAt: Date.now(),
            };
            settings.lastSyncSummary = summary;
            settings.failedDiagnostics = failedDiagnostics.slice(0, 50);

            if (failedCount > 0) {
                if (successCount > 0) {
                    summary.outcome = "partial";
                    settings.fileStates = this.mergePendingUploadBaselines(next);
                    settings.status = "partial_success";
                    settings.lastError = `部分同步完成：${successCount} 项成功，${failedCount} 项失败（已生成排查指引）`;
                    settings.lastSkippedCount = skippedPaths.size;
                    await (this.store as any).save();
                    if (temporaryFileFailure && !firstSync) this.scheduleBackgroundRetry();
                    this.statusCallback?.(`SmartMD: 部分同步 (${failedCount} 失败)`);
                    if (!silentSuccess || !this.retryAt) new Notice(localize("OneDrive 同步部分完成：{0} 成功，{1} 失败。可在设置查看详情。", successCount, failedCount), 8000);
                    return false;
                } else {
                    settings.fileStates = this.mergePendingUploadBaselines(next);
                    settings.status = "failed";
                    settings.lastError = `同步失败：全部 ${failedCount} 项传输均未成功。`;
                    settings.lastSkippedCount = skippedPaths.size;
                    await (this.store as any).save();
                    if (temporaryFileFailure && !firstSync) this.scheduleBackgroundRetry();
                    this.statusCallback?.("SmartMD: 同步失败");
                    if (!silentSuccess || !this.retryAt) new Notice(localize("OneDrive 同步失败：{0}", localizeMessage(settings.lastError)), 8000);
                    return false;
                }
            }

            const wasFirstCompleted = settings.firstSyncCompleted === true;
            const lastConfirmedSuccessAt = settings.lastSuccessAt;
            const firstMerge = settings.androidFirstMerge;
            const previousMode = {
                autoMode: settings.autoMode, enabled: settings.enabled, onStartup: settings.onStartup,
                afterSave: settings.afterSave, intervalMinutes: settings.intervalMinutes,
            };
            settings.fileStates = this.mergePendingUploadBaselines(next);
            settings.lastSkippedCount = skippedPaths.size;
            settings.failedDiagnostics = [];
            try {
                reportProgress?.(localize("正在确认同步及恢复记录…"));
                try {
                    await this.syncSharedState(operation);
                    this.sharedSettingsError = "";
                } catch (error) {
                    this.assertCurrent(operation);
                    this.sharedSettingsError = safeOneDriveError(error);
                }
                await this.maintainOneDriveRecoveries(operation);
                this.assertCurrent(operation);
                await (this.store as any).save();
                if (firstSync && !wasFirstCompleted) {
                    const rootId = await this.verifyFirstSyncRoot(settings.vaultId, next, () => this.assertCurrent(operation));
                    this.assertCurrent(operation);
                    settings.deletionRoot = rootId ? { ...rootScope, id: rootId } : undefined;
                    Object.assign(settings, this.newBindingSyncMode());
                }
                settings.firstSyncCompleted = true;
                delete settings.androidFirstMerge;
                settings.status = "ready";
                settings.lastSuccessAt = Date.now();
                settings.lastError = "";
                summary.outcome = "complete";
                reportProgress?.(localize("正在保存同步结果…"));
                if (firstSync) await this.store.saveWithIdentityGuard(() => this.assertCurrent(operation));
                else await this.store.save();
            } catch (error) {
                summary.outcome = "incomplete";
                settings.firstSyncCompleted = wasFirstCompleted;
                if (firstSync && !wasFirstCompleted) {
                    settings.deletionRoot = storedRoot;
                    Object.assign(settings, previousMode);
                }
                if (firstMerge) settings.androidFirstMerge = firstMerge;
                settings.lastSuccessAt = lastConfirmedSuccessAt;
                throw error;
            }
            if (temporaryFileFailure && !firstSync) this.scheduleBackgroundRetry();
            else this.clearBackgroundRetry();
            const checkingCount = pendingUploads + pendingRechecks;
            this.statusCallback?.(checkingCount ? localize("SmartMD: 正在核对文件") : "SmartMD: 已同步");
            if (!reportProgress && !silentSuccess) {
                const otherSkipped = skippedPaths.size - pendingUploadCount;
                const message = checkingCount
                    ? localize("OneDrive 正在后台核对 {0} 个文件。", checkingCount)
                    : localize("OneDrive 同步完成");
                new Notice(otherSkipped ? message + " " + localize("另有 {0} 个文件已跳过。", otherSkipped) : message);
            }
            return true;

        } catch (error: any) {
            if (!this.isCurrent(operation)) return false;
            const scopedUpload = !firstSync && currentProcessingPath && this.hasPendingUploadAt(currentProcessingPath);
            if (error?.isUnknownCommit && !scopedUpload) {
                // A stage cleanup or final PUT has an unresolved outcome.
                // Even when the trusted mother baseline is already durable,
                // do not carry the round's old broad permission forward.
                settings.runAllowed = false;
                settings.enabled = false;
            }
            if (!firstSync && (!error?.isUnknownCommit || scopedUpload) && this.isTemporarySyncFailure(error)) {
                this.scheduleBackgroundRetry(Number(error?.retryAfterMs) || 0);
            }
            await this.fail(error, currentProcessingPath, silentSuccess && Boolean(this.retryAt));
            return false;
        } finally {
            if (firstSync) {
                this.activeFirstSyncTicket = null;
            }
            this.running = false;
            this.authority.finishFinal(operation);
            this.notifyStatusChange();
        }
    }

    private contextKey(): string {
        const settings = this.settings();
        const local = (this.store as any).settings.local;
        return `${settings.vaultId}|${String(local.microsoftAccountId || "")}`;
    }

    /**
     * The normal upload path can checkpoint a verified mother baseline while
     * sync() still holds an older per-round "next". That snapshot must never
     * replace a committed baseline protected by a pending cleanup journal,
     * including on ordinary/partial/paused round exits.
     */
    private mergePendingUploadBaselines(next: Record<string, any>): Record<string, any> {
        const settings = this.settings();
        const ledger = settings.pendingUploadOperations || {};
        const current = settings.fileStates || {};
        const accountId = String((this.store as any).settings.local.microsoftAccountId || "");
        const merged: Record<string, any> = { ...next };
        for (const [path, record] of Object.entries(ledger) as [string, any][]) {
            if (!["FINAL_ACKNOWLEDGED", "CLEANUP_ISSUED", "CLEANUP_VERIFIED"]
                .includes(String(record?.phase || ""))) continue;
            const baseline = current[path];
            if (!record || record.motherPath !== path || record.direction !== "upload" ||
                record.accountId !== accountId || record.vaultId !== settings.vaultId ||
                record.deviceId !== settings.deviceId ||
                !record.finalId || !record.finalEtag ||
                !/^[0-9a-f]{64}$/i.test(String(record.sourceSha256 || "")) ||
                !baseline || baseline.remoteId !== record.finalId ||
                baseline.remoteEtag !== record.finalEtag ||
                baseline.localHash !== record.sourceSha256) {
                if (baseline) merged[path] = baseline;
                else delete merged[path];
                continue;
            }
            merged[path] = baseline;
        }
        return merged;
    }

    private async persistFinalState(operation: RunTicket, next: Record<string, any>): Promise<void> {
        if (!this.authority.isFinalizing(operation)) return;
        this.settings().fileStates = this.mergePendingUploadBaselines({ ...(this.settings().fileStates || {}), ...next });
        await (this.store as any).save();
        // Only after the trusted SHA/ItemID/ETag baseline has persisted may an
        // earlier successful upload's durable stage lock be released.
        await this.clearVerifiedUploadLedger(this.settings().fileStates);
    }

    private uploadLedger(): Record<string, any> {
        const s = this.settings();
        if (!s.pendingUploadOperations || typeof s.pendingUploadOperations !== "object"
            || Array.isArray(s.pendingUploadOperations)) {
            s.pendingUploadOperations = {};
        }
        return s.pendingUploadOperations;
    }

    private hasPendingUploadAt(path: string): boolean {
        const key = path.normalize("NFC").toLowerCase();
        return Object.keys(this.settings().pendingUploadOperations || {}).some(pending => {
            const other = pending.normalize("NFC").toLowerCase();
            return key === other || key.startsWith(other + "/") || other.startsWith(key + "/");
        });
    }

    private async recoverVerifiedUploads(operation: RunTicket): Promise<{ refreshed: boolean; retryNeeded: boolean }> {
        const settings = this.settings();
        const ledger = this.uploadLedger();
        let refreshed = false;
        let retryNeeded = false;
        const deferred = new Set<string>();
        for (const [path, record] of Object.entries(ledger)) {
            if (!record ||
                record.motherPath !== path || record.direction !== "upload" ||
                record.accountId !== settings.confirmationAccountId || record.vaultId !== settings.vaultId ||
                record.deviceId !== settings.deviceId || typeof record.operationId !== "string" ||
                !record.operationId ||
                !/^[a-f0-9]{64}$/.test(record.sourceSha256 || "") ||
                !Number.isFinite(record.sourceMtime) || record.sourceMtime <= 0 ||
                !Number.isSafeInteger(record.sourceBytes) || record.sourceBytes < 0 || record.sourceBytes > this.maxBytes() ||
                !path || path.includes("\\") || path.split("/").some(part => !part || part === "." || part === "..") ||
                this.isIgnored(path)) continue;
            const stagePath = "_suite_staging/" + record.operationId.replace(/[\\/:*?"<>|.]/g, "_") +
                "/" + path.replace(/[\\/:*?"<>|]/g, "_");
            if (record.stagingPath !== stagePath) continue;
            const rules = JSON.stringify([settings.ignorePatterns, this.maxBytes()]);
            const guard = () => {
                this.assertCurrent(operation);
                if (ledger !== settings.pendingUploadOperations || ledger[path] !== record ||
                    rules !== JSON.stringify([settings.ignorePatterns, this.maxBytes()])) throw new Error("上传记录或同步范围已变化");
            };
            try {
                if (["STAGE_SESSION_ISSUED", "STAGE_PUT_ISSUED"].includes(record.phase) &&
                    !record.stageId && !record.stageEtag && !record.contentAcknowledgedId &&
                    record.sourceBaseline === JSON.stringify(settings.fileStates?.[path] || null)) {
                    // No request for the final path was sent. A late staging
                    // request can only reach the previous operation's own path.
                    if (!(this.app.vault.getAbstractFileByPath(path) instanceof TFile) ||
                        await this.readRemoteFileMetadata(stagePath)) continue;
                    guard();
                    delete ledger[path];
                    try {
                        await this.store.saveWithIdentityGuard(() => {
                            this.assertCurrent(operation);
                            if (ledger !== settings.pendingUploadOperations || ledger[path] ||
                                rules !== JSON.stringify([settings.ignorePatterns, this.maxBytes()])) {
                                throw new Error("上传记录或同步范围已变化");
                            }
                        });
                    } catch (error) {
                        if (ledger === settings.pendingUploadOperations && !ledger[path]) ledger[path] = record;
                        throw error;
                    }
                    refreshed = true;
                    continue;
                }
                if (!record.stageId || !record.stageEtag) continue;
                if (["FINAL_ACKNOWLEDGED", "CLEANUP_ISSUED"].includes(record.phase)) {
                    const evidence = await this.inspectInterruptedUploadCleanup(path, operation);
                    guard();
                    await this.finishVerifiedStageCleanup(operation, path, evidence, guard);
                    refreshed = true;
                    continue;
                }
                if (!["RECONCILE_REQUIRED", "FINAL_PUT_ISSUED"].includes(record.phase) ||
                    record.sourceBaseline !== JSON.stringify(settings.fileStates?.[path] || null)) continue;
                const mother = await this.readRemoteFileMetadata(path);
                const stage = await this.readRemoteFileMetadata(stagePath);
                if (!mother?.id || !mother.etag || !stage || stage.id !== record.stageId || stage.id === mother.id || stage.etag !== record.stageEtag ||
                    stage.size !== record.sourceBytes || mother.size !== record.sourceBytes ||
                    (record.finalId && mother.id !== record.finalId) ||
                    (record.finalEtag && mother.etag !== record.finalEtag) ||
                    (record.expectedRemoteId && mother.id !== record.expectedRemoteId) ||
                    (record.expectedRemoteEtag && mother.etag === record.expectedRemoteEtag)) continue;
                const pinned = record.contentAcknowledgedId === mother.id && record.contentAcknowledgedEtag === mother.etag;
                const acknowledgedAtSourceTime = record.contentAcknowledgedId === mother.id &&
                    record.contentAcknowledgedEtag && mother.mtime === record.sourceMtime;
                const local = this.app.vault.getAbstractFileByPath(path);
                if (!(local instanceof TFile)) continue;
                const originalStillLocal = async () =>
                    local.stat.size === record.sourceBytes &&
                    await this.sha256Bytes(await this.app.vault.readBinary(local)) === record.sourceSha256;
                const localMatches = await originalStillLocal();
                if (!localMatches && !pinned && !acknowledgedAtSourceTime) continue;
                const data = await this.remoteBytes(mother);
                if (await this.sha256Bytes(data) !== record.sourceSha256 ||
                    !sameBytes(data, await this.remoteBytes(stage))) continue;
                guard();
                let verifiedMtime = record.sourceMtime;
                let verifiedEtag = mother.etag;
                if (mother.mtime !== record.sourceMtime) {
                    if (pinned && localMatches && local.stat.mtime === record.sourceMtime) {
                        const timestampGuard = () => {
                            guard();
                            if (this.app.vault.getAbstractFileByPath(path) !== local ||
                                local.stat.mtime !== record.sourceMtime ||
                                record.sourceBaseline !== JSON.stringify(settings.fileStates?.[path] || null)) {
                                throw new Error("上传原文件已变化");
                            }
                        };
                        this.authority.beginFinal(operation);
                        try {
                            const updated = await this.setUploadedModificationTime(
                                { id: mother.id, eTag: mother.etag }, record.sourceMtime, data, operation, timestampGuard);
                            verifiedEtag = updated.eTag || updated.etag;
                        } catch (error: any) {
                            timestampGuard();
                            if (error?.isFatalAuth || error?.retryAfterMs ||
                                [401, 403, 429, 507].includes(Number(error?.status)) || !mother.mtime) {
                                throw error;
                            }
                            // The upload receipt proves this version, even when its
                            // timestamp can't be restored. Verify it again below.
                            verifiedMtime = mother.mtime;
                        } finally {
                            this.authority.finishFinal(operation);
                        }
                    } else if (mother.mtime && (pinned || (record.legacyTimestampFailure && !record.contentAcknowledgedId))) {
                        // Keep the proven upload separate from any later local edit.
                        // Unpinned legacy uploads still require matching local bytes.
                        verifiedMtime = mother.mtime;
                    } else continue;
                }
                refreshed = true;
                const current = await this.readRemoteFileMetadata(path);
                if (!current || current.id !== mother.id || current.etag !== verifiedEtag ||
                    current.mtime !== verifiedMtime || !sameBytes(data, await this.remoteBytes(current)) ||
                    (!pinned && !acknowledgedAtSourceTime && !await originalStillLocal()) ||
                    this.app.vault.getAbstractFileByPath(path) !== local) continue;
                guard();
                const priorStates = settings.fileStates;
                const priorRecord = { ...record };
                settings.fileStates = { ...priorStates, [path]: {
                    localMtime: record.sourceMtime, localSize: record.sourceBytes, localHash: record.sourceSha256,
                    remoteId: current.id, remoteEtag: current.etag,
                } };
                Object.assign(record, { phase: "FINAL_ACKNOWLEDGED", finalId: current.id, finalEtag: current.etag,
                    verifiedRemoteMtime: verifiedMtime });
                try { await this.store.saveWithIdentityGuard(guard); }
                catch (error) {
                    if (this.settings() === settings && ledger[path] === record) {
                        settings.fileStates = priorStates;
                        ledger[path] = priorRecord;
                    }
                    throw error;
                }
                await this.finishVerifiedStageCleanup(operation, path, {
                    remoteId: current.id, remoteEtag: current.etag, stageId: stage.id, stageEtag: stage.etag,
                    stagePath, stageState: "PRESENT",
                }, guard);
            } catch (error: any) {
                if (!this.isCurrent(operation)) throw new Error("同步已停止");
                if (error?.isFatalAuth || [401, 403, 507].includes(Number(error?.status))) throw error;
                if (this.isTemporarySyncFailure(error)) {
                    if (Number(error?.status) === 429 || error?.retryAfterMs) throw error;
                    retryNeeded = true;
                    deferred.add(path);
                }
            }
        }
        for (const [path, record] of Object.entries(ledger)) {
            if (deferred.has(path) || !record || record.direction !== "upload" ||
                record.motherPath !== path || record.accountId !== settings.confirmationAccountId ||
                record.vaultId !== settings.vaultId || record.deviceId !== settings.deviceId ||
                !path || path.includes("\\") || path.split("/").some(part => !part || part === "." || part === "..") ||
                this.isIgnored(path)) continue;
            try {
                const result = await this.replanUnsettledUpload(operation, path, record);
                if (result === "replanned") refreshed = true;
                if (result === "retry") retryNeeded = true;
            } catch (error: any) {
                this.assertCurrent(operation);
                if (error?.isFatalAuth || error?.retryAfterMs || [401, 403, 429, 507].includes(Number(error?.status))) throw error;
                retryNeeded ||= this.isTemporarySyncFailure(error);
            }
        }
        return { refreshed, retryNeeded };
    }

    private async replanUnsettledUpload(
        operation: RunTicket, path: string,
        record: NonNullable<OneDriveSyncSettings["pendingUploadOperations"]>[string],
    ): Promise<"replanned" | "retry" | "hold"> {
        const settings = this.settings();
        const ledger = this.uploadLedger();
        const pinnedRecord = JSON.stringify(record);
        const policy = JSON.stringify([settings.ignorePatterns, this.maxBytes()]);
        const guard = () => {
            this.assertCurrent(operation);
            if (settings.pendingUploadOperations !== ledger ||
                JSON.stringify([settings.ignorePatterns, this.maxBytes()]) !== policy) {
                throw new Error("上传记录或同步范围已变化");
            }
        };
        const priorRoot = settings.deletionRoot;
        const rootId = await this.readDeletionRoot(operation);
        if (priorRoot && (priorRoot.vaultId !== settings.vaultId || priorRoot.accountId !== settings.confirmationAccountId ||
            priorRoot.deviceId !== settings.deviceId || rootId !== priorRoot.id)) return "hold";
        if (!priorRoot) {
            // An older empty vault may have started its first upload before a
            // root was stored. Require its independently identified stage.
            if (Object.keys(settings.fileStates || {}).length || !record.operationId || !record.stageId || !record.stageEtag) return "hold";
            const expectedStage = "_suite_staging/" + record.operationId.replace(/[\\/:*?"<>|.]/g, "_") +
                "/" + path.replace(/[\\/:*?"<>|]/g, "_");
            if (record.stagingPath !== expectedStage) return "hold";
            const stage = await this.readRemoteFileMetadata(expectedStage);
            if (stage?.id !== record.stageId || stage.etag !== record.stageEtag) return "hold";
        }

        const local = this.app.vault.getAbstractFileByPath(path);
        if (local && !(local instanceof TFile)) return "hold";
        if (local && local.stat.size > this.maxBytes()) return "hold";
        const localMtime = local?.stat.mtime;
        const localData = local ? await this.app.vault.readBinary(local) : null;
        const cloud = await this.readRemoteFileMetadata(path);
        if (cloud && (!cloud.id || !cloud.etag || cloud.size > this.maxBytes())) return "hold";
        const finalAttempt = record.phase === "FINAL_PUT_ISSUED" || record.phase === "RECONCILE_REQUIRED" ||
            record.phase === "FINAL_ACKNOWLEDGED" || record.phase === "CLEANUP_ISSUED" || record.phase === "CLEANUP_VERIFIED";
        const acknowledged = !!cloud && ((record.contentAcknowledgedId === cloud.id && record.contentAcknowledgedEtag === cloud.etag) ||
            (record.finalId === cloud.id && record.finalEtag === cloud.etag));
        const cloudData = cloud ? await this.remoteBytes(cloud) : null;
        const identical = !!localData && !!cloudData && sameBytes(localData, cloudData);
        const matchesOriginal = finalAttempt && !!cloudData && cloudData.byteLength === record.sourceBytes &&
            /^[a-f0-9]{64}$/.test(record.sourceSha256) && await this.sha256Bytes(cloudData) === record.sourceSha256;
        const replacedExpectedVersion = !!cloud && record.expectedRemoteId === cloud.id &&
            !!record.expectedRemoteEtag && record.expectedRemoteEtag !== cloud.etag;
        const retiredSmallPrecondition = !!cloud && !!record.expectedRemoteId && !!record.expectedRemoteEtag &&
            (record.expectedRemoteId !== cloud.id || record.expectedRemoteEtag !== cloud.etag) &&
            Number.isSafeInteger(record.sourceBytes) && record.sourceBytes >= 0 && record.sourceBytes < 4_000_000;
        const uncertainLargeOverwrite = finalAttempt && !!record.expectedRemoteId && record.sourceBytes >= 4_000_000 && !acknowledged;
        const safeNewFile = !cloud && !!localData && record.expectedRemoteId === "" &&
            record.expectedRemoteEtag === "" && /^[a-f0-9]{64}$/.test(record.sourceSha256) &&
            localData.byteLength === record.sourceBytes && await this.sha256Bytes(localData) === record.sourceSha256;
        const beforeFinal = record.phase === "STAGE_SESSION_ISSUED" || record.phase === "STAGE_PUT_ISSUED" ||
            record.phase === "STAGE_ACKNOWLEDGED";
        if (uncertainLargeOverwrite || (!beforeFinal && !(finalAttempt && (acknowledged || retiredSmallPrecondition || safeNewFile ||
            (matchesOriginal && (identical || replacedExpectedVersion)))))) return "hold";
        const hasLaterLocalEdit = !identical && matchesOriginal &&
            Number.isFinite(record.sourceMtime) && Number(localMtime) > Number(record.sourceMtime);
        const olderUploadedVersion = !!localData && !identical && matchesOriginal && acknowledged &&
            Number.isFinite(record.sourceMtime);
        if (hasLaterLocalEdit && !acknowledged && localData) {
            await this.writeVerifiedRecovery(operation, path, localData);
        }
        if (await this.readDeletionRoot(operation) !== rootId) return "hold";
        const currentCloud = await this.readRemoteFileMetadata(path);
        if (cloud?.id !== currentCloud?.id || cloud?.etag !== currentCloud?.etag) return "retry";
        if (local ? !sameBytes(localData!, await this.app.vault.readBinary(local)) :
            await this.app.vault.adapter.exists(path)) return "retry";
        if (await this.readDeletionRoot(operation) !== rootId) return "hold";
        const checkLocal = () => {
            guard();
            if ((this.app.vault.getAbstractFileByPath(path) || null) !== (local || null) ||
                local?.stat.mtime !== localMtime) throw new Error("本机文件已变化");
        };
        checkLocal();
        if (JSON.stringify(ledger[path]) !== pinnedRecord) return "hold";

        const prior = settings.fileStates;
        const next = { ...prior };
        // The uncertain upload cannot establish a deletion. When one side is
        // absent, let ordinary sync copy the surviving version.
        if (!local || !cloud || (hasLaterLocalEdit && !acknowledged)) delete next[path];
        else if (identical && localData) {
            next[path] = { localMtime: local.stat.mtime, localSize: localData.byteLength,
                localHash: await this.sha256Bytes(localData), remoteId: cloud.id, remoteEtag: cloud.etag };
        } else if (olderUploadedVersion) {
            // OneDrive may report the time it received the old contents.
            // The bytes link it to the earlier local version, so the later
            // local edit should be processed as the only subsequent change.
            next[path] = { localMtime: record.sourceMtime!, localSize: record.sourceBytes,
                localHash: record.sourceSha256, remoteId: cloud.id, remoteEtag: cloud.etag };
        }
        checkLocal();
        if (JSON.stringify(ledger[path]) !== pinnedRecord || settings.fileStates !== prior) return "hold";
        await this.store.saveWithIdentityGuard(() => {
            checkLocal();
            if (JSON.stringify(ledger[path]) !== pinnedRecord || settings.fileStates !== prior ||
                settings.deletionRoot !== priorRoot) throw new Error("上传记录已变化");
        }, () => {
            settings.fileStates = next;
            if (!priorRoot) settings.deletionRoot = { accountId: settings.confirmationAccountId!, vaultId: settings.vaultId,
                deviceId: settings.deviceId, id: rootId };
            const newRoot = settings.deletionRoot;
            delete ledger[path];
            return () => {
                if (settings.pendingUploadOperations === ledger && !ledger[path]) ledger[path] = record;
                if (settings.fileStates === next) settings.fileStates = prior;
                if (settings.deletionRoot === newRoot) settings.deletionRoot = priorRoot;
            };
        });
        // The old staging object may hold the only previous version.
        // This path retires the local lock and leaves that object intact.
        return "replanned";
    }

    private async finishVerifiedStageCleanup(operation: RunTicket, path: string, evidence: {
        remoteId: string; remoteEtag: string; stageId: string; stageEtag: string; stagePath: string;
        stageState: "PRESENT" | "ABSENT_CONFIRMED_BY_GET";
    }, guard: () => void): Promise<void> {
        guard();
        this.authority.beginFinal(operation);
        try {
            if (evidence.stageState === "PRESENT") await this.setUploadPhase(path, "CLEANUP_ISSUED");
            const latest = await this.readRemoteFileMetadata(path);
            if (!latest || latest.id !== evidence.remoteId || latest.etag !== evidence.remoteEtag) return;
            guard();
            if (evidence.stageState === "PRESENT") {
                await this.deleteRemote({ path: evidence.stagePath, id: evidence.stageId,
                    etag: evidence.stageEtag, name: "", size: 0 }, operation, guard);
            }
            guard();
            await this.setUploadPhase(path, "CLEANUP_VERIFIED");
            await this.clearVerifiedUploadLedger(this.settings().fileStates);
        } finally {
            this.authority.finishFinal(operation);
        }
    }

    private async setUploadPhase(
        motherPath: string, phase: string, extra: Record<string, unknown> = {},
    ): Promise<void> {
        const ledger = this.uploadLedger();
        if (!ledger[motherPath]) throw new Error("待核上传操作身份丢失，停止云端写入");
        const previous = String(ledger[motherPath].phase || "");
        // A confirmed final version / issued cleanup must never be rewritten
        // as an earlier unknown PUT. That would erase the only recoverable
        // checkpoint after a successfully dispatched staging DELETE.
        const normalOrder = [
            "STAGE_SESSION_ISSUED", "STAGE_ACKNOWLEDGED",
            "FINAL_PUT_ISSUED", "FINAL_ACKNOWLEDGED",
            "CLEANUP_ISSUED", "CLEANUP_VERIFIED",
        ];
        const previousRank = normalOrder.indexOf(previous);
        const nextRank = normalOrder.indexOf(phase);
        if ((previousRank >= 0 && nextRank >= 0 && nextRank < previousRank) ||
            (phase === "RECONCILE_REQUIRED" &&
                ["FINAL_ACKNOWLEDGED", "CLEANUP_ISSUED", "CLEANUP_VERIFIED"].includes(previous))) {
            throw new Error("原上传结算阶段不得倒退，保留已核实的最终文件与暂存清理日志");
        }
        Object.assign(ledger[motherPath], extra, { phase, updatedAt: Date.now() });
        await (this.store as any).save();
    }

    private async clearVerifiedUploadLedger(next: Record<string, any>): Promise<void> {
        if (this.uploadLedgerSettlementInFlight) {
            throw new Error("另一上传日志结算尚未持久完成，禁止并发释放");
        }
        this.uploadLedgerSettlementInFlight = true;
        try {
            const ledger = this.uploadLedger();
            const settings = this.settings();
            const accountId = String((this.store as any).settings.local.microsoftAccountId || "");
            const released: Array<[string, any]> = [];
            for (const [path, record] of Object.entries(ledger) as [string, any][]) {
                const state = next[path];
                if (record?.phase !== "CLEANUP_VERIFIED" || !state ||
                    record.motherPath !== path || record.direction !== "upload" ||
                    record.accountId !== accountId ||
                    record.vaultId !== settings.vaultId || record.deviceId !== settings.deviceId ||
                    !record.stageId || !record.stageEtag || !record.stagingPath ||
                    !record.finalId || !record.finalEtag || !record.sourceSha256 ||
                    state.remoteId !== record.finalId ||
                    state.remoteEtag !== record.finalEtag ||
                    state.localHash !== record.sourceSha256) continue;
                released.push([path, record]);
                delete ledger[path];
            }
            if (released.length) {
                // Invalidate previews even if persistence fails and the journal is
                // restored. Do not let a stale snapshot authorize a new operation.
                this.uploadLedgerSettlementGeneration++;
                try {
                    await (this.store as any).save();
                } catch (error) {
                    // A newer UNKNOWN operation takes precedence over a stale
                    // completed record; rollback must never overwrite its identity.
                    if (this.settings() === settings &&
                        this.settings().pendingUploadOperations === ledger) {
                        for (const [path, record] of released) {
                            if (!Object.prototype.hasOwnProperty.call(ledger, path)) {
                                ledger[path] = record;
                            }
                        }
                    }
                    throw error;
                }
            }
        } finally {
            this.uploadLedgerSettlementInFlight = false;
        }
    }

    private async sha256Bytes(data: ArrayBuffer): Promise<string> {
        const value = await crypto.subtle.digest("SHA-256", data);
        return [...new Uint8Array(value)]
            .map((byte) => byte.toString(16).padStart(2, "0")).join("");
    }

    private localFiles(): LocalFile[] {
        return this.app.vault.getFiles()
            .filter((file) => file.stat.size <= this.maxBytes())
            .filter((file) => !this.isIgnored(file.path))
            .map((file) => ({ path: file.path, file, size: file.stat.size, mtime: file.stat.mtime }));
    }

    private isCurrent(operation: RunTicket): boolean {
        return this.resumedOperation === operation
            ? this.resumedOperationIsCurrent(operation)
            : this.authority.isCurrent(operation) && this.canRun();
    }

    private assertCurrent(operation: RunTicket): void {
        this.authority.require(operation);
        if (!this.isCurrent(operation) || operation.contextKey !== this.contextKey()) {
            throw new Error("同步身份或范围已变化，正式落地已停止");
        }
    }

    private async assertFresh(
        operation: RunTicket,
        path: string,
        expectedLocal: LocalFile | undefined,
        expectedRemote: RemoteFile | undefined,
    ): Promise<void> {
        this.assertCurrent(operation);
        let remote: RemoteFile | null | undefined = undefined;
        if (typeof (this.auth as any)?.graph === "function") {
            remote = await this.readRemoteFileMetadata(path, false);
        } else {
            remote = (await this.remoteFiles(false)).find((item) => item.path === path);
        }
        this.assertCurrent(operation);
        const local = this.localFiles().find((item) => item.path === path);
        const localMatches = expectedLocal
            ? local?.mtime === expectedLocal.mtime && local?.size === expectedLocal.size
            : !local;
        const remoteMatches = expectedRemote
            ? remote?.id === expectedRemote.id && remote?.etag === expectedRemote.etag
            : !remote;
        if (!localMatches || !remoteMatches) {
            throw Object.assign(new Error(`状态复核后再次变化：${path}`), { recheckFile: true });
        }
    }

    private maxBytes(): number {
        return Math.min(250, Math.max(1, Number(this.settings().maxFileSizeMb || 250))) * 1024 * 1024;
    }

    private maxResumedFileBytes(): number {
        return Platform.isAndroidApp ? this.maxBytes() : Math.min(4_000_000, this.maxBytes());
    }

    private isIgnored(path: string): boolean {
        // Keep historical recovery files out of ordinary sync.
        if (path.startsWith("OneDrive恢复的笔记/") ||
            /\.conflict-[^/]+-\d{4}-\d{2}-\d{2}T[\d-]+Z(?:\.[^/]+)?$/.test(path)) {
            return true;
        }
        return (this.settings().ignorePatterns || []).map(ignoreMatcher).some((pattern: RegExp) => pattern.test(path));
    }

    private skippedLocalPaths(): Set<string> {
        return new Set(this.app.vault.getFiles()
            .filter((file) => !this.isIgnored(file.path) && file.stat.size > this.maxBytes())
            .map((file) => file.path));
    }

    /**
     * An interrupted FIRST RECEIVE retains its previously approved account,
     * cloud Vault and installed physical root. After a cold restart these are
     * identity evidence only: runAllowed is revoked and a new preview + click
     * is mandatory. This is never a fresh empty-Vault authorization.
     */
    private incompleteFirstReceiverScope(accountId: string, cloudVaultId: string): boolean {
        const s = this.settings();
        const local = (this.store as any).settings.local;
        const previous = s.fileStates || {};
        const noEntries = (value: any): boolean => !value ||
            (typeof value === "object" && !Array.isArray(value) &&
                Object.keys(value).length === 0);
        return Boolean(accountId && cloudVaultId &&
            s.firstSyncCompleted !== true && s.confirmed === true &&
            s.vaultId === cloudVaultId && s.confirmationVaultId === cloudVaultId &&
            s.confirmationAccountId === accountId &&
            String(local.microsoftAccountId || "") === accountId &&
            s.installationProofId && s.deviceId &&
            Number(s.lastSuccessAt || 0) === 0 &&
            typeof (this.store as any).installationProofIsValid === "function" &&
            (this.store as any).installationProofIsValid() === true &&
            (Platform.isAndroidApp || Object.keys(previous).length > 0) &&
            noEntries(s.pendingUploadOperations) && noEntries(s.pendingDeletions) &&
            noEntries(s.pendingConflicts) &&
            (!Array.isArray(local.recoveryEntries) || local.recoveryEntries.length === 0) &&
            !s.firstCommitJournal && !s.firstFileJournal);
    }

    /**
     * Only files with a durable first-download baseline may be present. Do not
     * turn an interrupted receiver into an implicit UPLOAD or deletion round.
     * The final check reads real bytes again before issuing a new ticket.
     */
    private async assertIncompleteFirstReceiverFiles(
        accountId: string, cloudVaultId: string,
        local: LocalFile[], remote: RemoteFile[], final = false,
    ): Promise<void> {
        if (!this.incompleteFirstReceiverScope(accountId, cloudVaultId)) {
            throw new Error("首次接收历史身份或未决事务不可信，必须保持暂停");
        }
        const previous = this.settings().fileStates || {};
        const state = this.settings();
        const scope = {
            accountId, vaultId: cloudVaultId,
            deviceId: String(state.deviceId || ""),
            installationProofId: String(state.installationProofId || ""),
        };
        const intent = Platform.isAndroidApp ? this.androidFirstDownloadIntent() : null;
        const history = { cloudVaultId, trackedPaths: Object.keys(previous),
            android: Platform.isAndroidApp ? { intent } : undefined };
        const inventory = Platform?.isAndroidApp
            ? await inspectAndroidReceiverCandidateInventory(this.app, scope, history)
            : await inspectReceiverCandidateInventory(this.app, null, undefined, scope, history);
        if (inventory.status !== "ELIGIBLE_INCOMPLETE_RECEIVER") {
            throw new Error("首次接收本地存在未核实的隐藏文件、旧历史或物理范围，拒绝续传");
        }
        const byPath = new Map(remote.map((item) => [item.path, item]));
        const localNames = new Set(local.map((item) => item.path));
        const pendingLocal = intent && local.find(item => item.path === intent.path);
        if (local.length !== Object.keys(previous).length + Number(Boolean(pendingLocal)) ||
            localNames.size !== local.length) {
            throw new Error("首次接收的本地文件数量或历史基线不一致，禁止续传");
        }
        for (const [path, old] of Object.entries(previous) as [string, any][]) {
            const current = local.find((item) => item.path === path);
            const cloud = byPath.get(path);
            if (!current || !cloud || !old?.remoteId || !old?.remoteEtag ||
                !/^[0-9a-f]{64}$/i.test(String(old.localHash || "")) ||
                old.remoteId !== cloud.id || !old.localMtime ||
                !Number.isFinite(Number(old.localSize)) ||
                (final && (old.remoteEtag !== cloud.etag ||
                    old.localMtime !== current.mtime ||
                    old.localSize !== current.size))) {
                throw new Error("首次接收的已下载文件或云端版本发生变化，保留两端内容并停止续传");
            }
            if (final) {
                const data = await this.app.vault.readBinary(current.file);
                if ((await this.sha256Bytes(data)) !== old.localHash ||
                    !sameBytes(data, await this.remoteBytes(cloud))) {
                    throw new Error("首次接收原文与可信云端版本不同，禁止当作已同步文件");
                }
            }
        }
        if (intent) await this.inspectAndroidFirstDownload(intent, pendingLocal || undefined, byPath.get(intent.path));
    }

    private async freshFirstReceiverEligible(accountId: string): Promise<boolean> {
        const s = this.settings();
        const local = (this.store as any).settings.local;
        if (!accountId || accountId !== String(local.microsoftAccountId || "") ||
            s.confirmed || s.runAllowed || s.enabled || s.firstSyncCompleted ||
            s.confirmationAccountId || s.confirmationVaultId || s.cloudVaultId ||
            s.localVaultId || Number(s.lastSuccessAt || 0) !== 0 ||
            Object.keys(s.fileStates || {}).length ||
            Object.keys(s.pendingUploadOperations || {}).length ||
            Object.keys(s.pendingDeletions || {}).length ||
            s.androidFirstDownload ||
            (local.recoveryEntries || []).length ||
            typeof (this.store as any).installationProofIsValid !== "function" ||
            (this.store as any).installationProofIsValid() !== true) return false;
        const evidence = {
            accountId, vaultId: String(s.vaultId || ""),
            deviceId: String(s.deviceId || ""),
            installationProofId: String(s.installationProofId || ""),
        };
        if (!evidence.vaultId || !evidence.deviceId || !evidence.installationProofId) return false;
        const inventory = Platform?.isAndroidApp
            ? await inspectAndroidReceiverCandidateInventory(this.app, evidence)
            : await inspectReceiverCandidateInventory(this.app, null, undefined, evidence);
        return inventory.status === "ELIGIBLE_EMPTY_RECEIVER";
    }

    private androidFirstDownloadIntent(): AndroidFirstDownload | null {
        const settings = this.settings();
        const intent = settings.androidFirstDownload as AndroidFirstDownload | undefined;
        if (!intent) return null;
        if (intent.accountId !== this.store.settings.local.microsoftAccountId ||
            intent.vaultId !== settings.vaultId || intent.deviceId !== settings.deviceId ||
            intent.proofId !== settings.installationProofId || !this.store.installationProofIsValid() ||
            !intent.remoteId || !intent.remoteEtag || !/^[a-f0-9]{64}$/.test(intent.sha256) ||
            !Number.isSafeInteger(intent.size) || intent.size < 0 || intent.size > this.maxBytes() ||
            this.isIgnored(intent.path)) {
            throw new Error("安卓首次下载的操作记录或身份不可信");
        }
        return intent;
    }

    private async inspectAndroidFirstDownload(
        intent: AndroidFirstDownload, local: LocalFile | undefined, remote: RemoteFile | undefined,
    ): Promise<boolean> {
        const adapter = this.app.vault.adapter;
        const staging = await adapter.exists(intent.stagingPath);
        const target = await adapter.exists(intent.path);
        const currentRemote = await this.readRemoteFileMetadata(intent.path);
        const validRemote = (file: RemoteFile | null | undefined) =>
            file?.id === intent.remoteId && file.etag === intent.remoteEtag && file.size === intent.size;
        if (!validRemote(remote) || !validRemote(currentRemote) ||
            target !== Boolean(local) || (staging && target && getAndroidAppStorageRoot(adapter))) {
            throw new Error("安卓首次下载的待核对文件或云端版本已变化");
        }
        if (target || staging) {
            const cloudBytes = await this.remoteBytes(remote!);
            for (const path of [target ? intent.path : "", staging ? intent.stagingPath : ""].filter(Boolean)) {
                const bytes = await adapter.readBinary(path);
                if (bytes.byteLength !== intent.size || await this.sha256Bytes(bytes) !== intent.sha256 ||
                    !sameBytes(bytes, cloudBytes)) {
                    throw new Error("安卓首次下载的原件或暂存文件字节不符，保留现场");
                }
            }
        }
        if (this.androidFirstDownloadIntent() !== intent) throw new Error("安卓首次下载操作身份已变化");
        return staging;
    }

    private async settleAndroidFirstDownload(
        local: LocalFile[], remote: RemoteFile[], assertGrantCurrent: () => void,
    ): Promise<void> {
        const intent = this.androidFirstDownloadIntent();
        if (!intent) return;
        const settings = this.settings();
        const diskFile = local.find(file => file.path === intent.path);
        const staged = await this.inspectAndroidFirstDownload(intent, diskFile, remote.find(file => file.path === intent.path));
        assertGrantCurrent();
        if (staged) await this.removeVerifiedStaging(intent.stagingPath);
        assertGrantCurrent();
        const previous = settings.fileStates;
        if (diskFile) settings.fileStates = { ...previous, [intent.path]: {
            localMtime: diskFile.mtime, localSize: diskFile.size, localHash: intent.sha256,
            remoteId: intent.remoteId, remoteEtag: intent.remoteEtag,
        } };
        delete settings.androidFirstDownload;
        try { await this.store.saveWithIdentityGuard(assertGrantCurrent); }
        catch (error) {
            settings.fileStates = previous;
            settings.androidFirstDownload = intent;
            throw error;
        }
    }

    private async assertFirstReceiverIdentity(choice: PendingFirstChoice, accountId: string): Promise<void> {
        if (choice.mode !== "existing") return;
        const s = this.settings();
        const previouslyApproved = s.confirmed === true && s.runAllowed === true
            && s.vaultId === choice.id
            && s.confirmationVaultId === choice.id
            && s.confirmationAccountId === accountId
            && typeof (this.store as any).installationProofIsValid === "function"
            && (this.store as any).installationProofIsValid() === true;
        if (!previouslyApproved &&
            !this.incompleteFirstReceiverScope(accountId, choice.id) &&
            !(await this.freshFirstReceiverEligible(accountId))) {
            throw new Error("当前知识库未通过独立空库、安装证明与历史记录核验，已停止首次接收；原云端知识库保持不变。");
        }
    }

    // An absent registry is NEVER sufficient evidence of a new account.
    // This inventory runs only after the user consented to a special-folder
    // GET that Microsoft may implement as implicit AppFolder creation.
    private async readEmptyAppFolderRootId(): Promise<string> {
        const root: any = await this.graph("GET", "/drive/special/approot");
        // The special URL itself selects the application folder. Some
        // personal OneDrive Graph responses omit the specialFolder facet
        // even when the returned item is the genuine approot directory.
        // A PRESENT but contradictory facet is still an identity failure;
        // do not reject a proven folder merely for an omitted optional facet.
        if (!root?.folder || !root?.id ||
            (root.specialFolder &&
                String(root.specialFolder.name || "").toLowerCase() !== "approot")) {
            throw new Error("微软AppFolder根目录或专用身份不可信，不能初始化空白注册表");
        }
        const id = String(root.id);
        let next: string | null = `/drive/items/${encodeURIComponent(id)}/children?$top=200`;
        const visited = new Set<string>();
        while (next) {
            this.validateGraphEndpoint(next);
            if (visited.has(next) || visited.size >= 50) {
                throw new Error("AppFolder清单分页异常，无法证明目录为空");
            }
            visited.add(next);
            const page: any = await this.graph("GET", next);
            if (!Array.isArray(page?.value) || page.value.length > 0) {
                throw new Error("AppFolder内已有内容或清单不可信，禁止以空白状态初始化");
            }
            const link: unknown = page["@odata.nextLink"];
            if (link !== undefined && link !== null && link !== "" && typeof link !== "string") {
                throw new Error("AppFolder分页身份异常，停止初始化");
            }
            next = typeof link === "string" && link ? link : null;
        }
        return id;
    }

    // A previous attempt may have committed the registry while the reply or
    // local completion save was lost. Only the original durable operation ID
    // may release this hold; never blindly repeat the create request.
    private async reconcileEmptyAppFolderBootstrap(): Promise<void> {
        const s = this.settings();
        const journal: any = s.emptyAppFolderBootstrap;
        if (!journal) return;
        const account = String((this.store as any).settings.local.microsoftAccountId || "");
        if (!journal.opId || journal.accountId !== account ||
            journal.localVaultId !== s.vaultId || journal.deviceId !== s.deviceId ||
            (this.store as any).installationProofIsValid?.() !== true) {
            throw new Error("空白AppFolder历史初始化的设备或账号身份不可信，拒绝恢复");
        }
        const payload = await this.readSharedState("vaults/registry.json");
        if (!Array.isArray(payload?.vaults) || payload.vaults.length ||
            payload.bootstrapOpId !== journal.opId) {
            throw new Error("空白AppFolder初始化结果未知，必须保留历史操作并人工对账");
        }
        await this.assertFirstCommitGraphAccount(account);
        const meta = await this.readRemoteFileMetadata("vaults/registry.json", true);
        if (!meta?.id || !meta.etag) throw new Error("空白注册表版本不可信");
        delete s.emptyAppFolderBootstrap;
        await (this.store as any).save();
        this.vaultRegistryCache = null;
        this.registryReadIsTrusted = false;
    }

    private async inspectEmptyAppFolderForSetup(accountId: string): Promise<string> {
        const settings = this.settings();
        const local = (this.store as any).settings.local;
        const hasEntries = (value: any): boolean => Boolean(value &&
            typeof value === "object" && Object.keys(value).length);
        if (settings.emptyAppFolderBootstrap || settings.firstSyncCompleted ||
            settings.confirmed || settings.cloudVaultId ||
            settings.firstCommitJournal || settings.firstFileJournal ||
            hasEntries(settings.fileStates) || hasEntries(settings.pendingUploadOperations) ||
            hasEntries(settings.pendingConflicts) || hasEntries(settings.pendingDeletions) ||
            hasEntries(local.recoveryEntries) ||
            (this.store as any).installationProofIsValid?.() !== true) {
            throw new Error("本机尚有旧同步、恢复或初始化历史，不能建立空白注册表");
        }
        // The missing registry may be a deletion or corruption, not a new user.
        // Only a *fully enumerated empty app root* permits a new bootstrap.
        await this.assertFirstCommitGraphAccount(accountId);
        const firstRoot = await this.readEmptyAppFolderRootId();
        if (await this.readSharedState("vaults/registry.json") !== null) {
            throw new Error("空白AppFolder检查时注册表已出现，必须重新读取");
        }
        return firstRoot;
    }

    private async bootstrapEmptyAppFolderRegistry(
        accountId: string, firstRoot: string, assertGrantCurrent: () => void,
    ): Promise<string> {
        const settings = this.settings();
        assertGrantCurrent();
        const sameScope = () => this.settings() === settings &&
            String((this.store as any).settings.local.microsoftAccountId || "") === accountId &&
            settings.firstSyncCompleted !== true && settings.runAllowed !== true &&
            (this.store as any).installationProofIsValid?.() === true;
        if (!sameScope()) throw new Error("首次初始化账号、设备或权限已变化");
        if (firstRoot !== await this.inspectEmptyAppFolderForSetup(accountId)) {
            throw new Error("空白目录二次复核已变化，停止创建");
        }
        assertGrantCurrent();
        const journal: any = {
            opId: typeof crypto.randomUUID === "function" ? crypto.randomUUID() :
                Date.now().toString(36) + "-" + Math.random().toString(36).slice(2),
            accountId, localVaultId: settings.vaultId, deviceId: settings.deviceId,
            rootId: firstRoot, stage: "PREPARED", updatedAt: Date.now(),
        };
        settings.emptyAppFolderBootstrap = journal;
        try { await (this.store as any).save(); } catch (error) {
            delete settings.emptyAppFolderBootstrap;
            throw error;
        }
        const guard = () => {
            assertGrantCurrent();
            if (!sameScope() || settings.emptyAppFolderBootstrap !== journal ||
                journal.stage !== "PREPARED") {
                throw new Error("空白AppFolder创建许可已失效");
            }
        };
        try {
            guard();
            await this.ensureFolder("vaults", undefined, guard);
            guard();
            const folder: any = await this.graph("GET", "/drive/special/approot:/vaults");
            if (!folder?.folder || !folder?.id) throw new Error("空白vaults目录身份不可信");
            const children: any = await this.graph("GET",
                `/drive/items/${encodeURIComponent(String(folder.id))}/children?$top=200`);
            if (!Array.isArray(children?.value) || children.value.length || children["@odata.nextLink"]) {
                throw new Error("初始化期间vaults目录已出现其他内容，拒绝写注册表");
            }
            guard();
            await this.assertFirstCommitGraphAccount(accountId);
            guard();
            const body = new TextEncoder().encode(JSON.stringify({
                vaults: [], bootstrapOpId: journal.opId, updatedAt: Date.now(),
            })).buffer;
            await this.upload("vaults/registry.json", body, true, null,
                undefined, undefined, guard);
            guard();
            const current = await this.readSharedState("vaults/registry.json");
            const meta = await this.readRemoteFileMetadata("vaults/registry.json", true);
            if (!Array.isArray(current?.vaults) || current.vaults.length ||
                current.bootstrapOpId !== journal.opId || !meta?.id || !meta.etag) {
                throw new Error("空注册表已发出但回读身份不完整，保留事务待对账");
            }
            journal.stage = "VERIFIED";
            await (this.store as any).save();
            delete settings.emptyAppFolderBootstrap;
            await (this.store as any).save();
            this.vaultRegistryCache = null;
            return await contentHash(JSON.stringify(current));
        } catch (error) {
            // A failed save may leave the remote registry committed; never
            // drop the durable identity of the original bootstrap operation.
            settings.emptyAppFolderBootstrap = journal;
            journal.stage = "RECONCILE_REQUIRED";
            journal.updatedAt = Date.now();
            try { await (this.store as any).save(); } catch {}
            throw new Error("空白AppFolder初始化未完成，已保留原操作待只读对账");
        }
    }

    private canStartAndroidFirstMerge(accountId: string): boolean {
        const settings = this.settings();
        const local = this.store.settings.local;
        return Boolean(Platform.isAndroidApp && accountId && local.microsoftAccountId === accountId &&
            this.store.installationProofIsValid() && settings.installationProofId && settings.deviceId &&
            !settings.firstSyncCompleted && !settings.confirmed && !settings.runAllowed && !settings.enabled &&
            !settings.confirmationAccountId && !settings.confirmationVaultId && !settings.cloudVaultId &&
            !settings.androidFirstMerge && !settings.androidFirstDownload &&
            !settings.firstCommitJournal && !settings.firstFileJournal && !settings.identicalRegistrationJournal &&
            !(settings as any).firstConflictPreservationJournal &&
            !Object.keys(settings.fileStates || {}).length && !Object.keys(settings.pendingUploadOperations || {}).length &&
            !Object.keys(settings.pendingDeletions || {}).length && !Object.keys(settings.pendingConflicts || {}).length &&
            !local.recoveryEntries?.length);
    }

    private androidFirstMergeIsTrusted(accountId: string, cloudVaultId: string): boolean {
        const settings = this.settings();
        const record = settings.androidFirstMerge;
        return Boolean(Platform.isAndroidApp && accountId && cloudVaultId && !settings.firstSyncCompleted &&
            settings.confirmed && record && this.store.settings.local.microsoftAccountId === accountId &&
            settings.confirmationAccountId === accountId && settings.confirmationVaultId === cloudVaultId &&
            settings.vaultId === cloudVaultId && record.accountId === accountId && record.cloudVaultId === cloudVaultId &&
            record.proofId === settings.installationProofId && record.deviceId === settings.deviceId &&
            this.store.installationProofIsValid() && !settings.firstCommitJournal &&
            !settings.firstFileJournal && !settings.androidFirstDownload);
    }

    private async prepareVaultIdentity(): Promise<boolean> {
        // Discovery is deliberately read-only; in-memory selections are not durable binding.
        this.pendingFirstChoice = null;
        this.identicalRegistrationRequested = false;
        const settings = this.settings();
        if (settings.firstSyncCompleted) return this.canRun();
        const accountId = String((this.store as any).settings.local.microsoftAccountId || "");
        if (!accountId) return false;
        await this.reconcileEmptyAppFolderBootstrap();
        const vaults = await this.readVaultRegistry();
        let emptyAppRoot: string | undefined;
        if (!this.registryReadIsTrusted) {
            // Existing vault directories without their registry are not a
            // new installation: never overwrite or silently reconstruct them.
            if (vaults.length) throw new Error("云端知识库列表不完整。请稍后重试。");
            emptyAppRoot = await this.inspectEmptyAppFolderForSetup(accountId);
        }
        const current = vaults.find((vault) => vault.id === settings.vaultId);
        if (settings.androidFirstMerge) {
            if (!current || !this.androidFirstMergeIsTrusted(accountId, current.id)) {
                throw new Error(localize("当前知识库已有同步历史或身份已变化，请重新核对。"));
            }
            this.pendingFirstChoice = { id: current.id, name: current.name, mode: "existing", mergeAndroid: true };
            return true;
        }
        if (current && this.incompleteFirstReceiverScope(accountId, current.id)) {
            // Do not inherit runAllowed from the old process. Reopen read-only
            // preview of the SAME registered cloud Vault; final click is new.
            this.pendingFirstChoice = { id: current.id, name: current.name, mode: "existing" };
            return true;
        }
        if (current) {
            // 已完成首次同步且凭据匹配的稳态运行
            const isCompletedTrusted = settings.firstSyncCompleted === true
                && settings.confirmed === true && settings.runAllowed === true
                && settings.confirmationVaultId === settings.vaultId
                && settings.confirmationAccountId === accountId;
            if (isCompletedTrusted) {
                this.pendingFirstChoice = { id: current.id, name: current.name, mode: "existing" };
                return true;
            }
            // 首次同步中途中断的合法断点续连 (BOUND_INCOMPLETE_CONFIRMED)
            const isResumptionTrusted = settings.confirmed === true && settings.runAllowed === true
                && settings.confirmationVaultId === settings.vaultId
                && settings.confirmationAccountId === accountId;
            if (isResumptionTrusted) {
                this.pendingFirstChoice = { id: current.id, name: current.name, mode: "existing" };
                return true;
            }
            // 若凭据不匹配（不可信配置/导入配置），绝不静默 return false 造成死锁，降级打开模态窗重新确认
        }

        const eligibleReceiver = vaults.length > 0 && await this.freshFirstReceiverEligible(accountId);
        const choice = vaults.length === 0 ? "new" : await new Promise<RemoteVault | "new" | null>((resolve) => {
            new VaultSelectionModal(this.app, vaults, resolve, eligibleReceiver,
                Platform.isDesktopApp && Platform.isWin, !eligibleReceiver && this.canStartAndroidFirstMerge(accountId)).open();
        });
        if (!choice) return false;
        if (choice === "new") {
            const newId = "vault-" + (typeof crypto.randomUUID === "function"
                ? crypto.randomUUID()
                : Date.now().toString(36) + "-" + Math.random().toString(36).slice(2));
            this.pendingFirstChoice = { id: newId, name: this.app.vault.getName(), mode: "new", emptyAppRoot };
        } else {
            const match = vaults.find((vault) => vault.id === choice.id);
            if (!match) throw new Error("云端知识库选择已过期，需重新检查。");
            this.pendingFirstChoice = { id: match.id, name: match.name, mode: "existing", mergeAndroid: choice.mergeAndroid };
            this.identicalRegistrationRequested = choice.registerIdentical === true;
            if (Platform.isAndroidApp) {
                this.firstConnectionProgress = localize("正在准备「{0}」的同步预览…", androidVaultDisplayName(match, vaults));
                this.firstConnectionNotice = new Notice(this.firstConnectionProgress);
                this.notifyStatusChange();
            }
        }
        return true;
    }

    private assertIndependentCopyScope(archive: any): void {
        requireWindowsHost();
        const settings = this.settings();
        const local: any = (this.store as any).settings.local;
        const fs = require("fs") as typeof import("fs");
        const path = require("path").win32 as typeof import("path").win32;
        const root = path.normalize(String((this.app.vault.adapter as any).getBasePath?.() || ""));
        if (!archive || local.microsoftAccountId !== archive.accountId ||
            settings.deviceId !== archive.deviceId || settings.installationProofId !== archive.proofId ||
            (this.store as any).installationProofIsValid?.() !== true ||
            root.toLowerCase() !== archive.physicalRoot.toLowerCase() ||
            path.normalize(fs.realpathSync(root)).toLowerCase() !== root.toLowerCase())
            throw new Error("独立副本的当前账号、物理根或新安装证明已变化");
        const stat = fs.lstatSync(root);
        if (!stat.isDirectory() || stat.isSymbolicLink() ||
            [String(stat.dev), String(stat.ino), String(stat.birthtimeMs)].join(":") !== archive.rootPhysicalIdentity)
            throw new Error("独立副本的物理根身份已变化");
        const hash = (value: any) => (require("node:crypto") as typeof import("node:crypto"))
            .createHash("sha256").update(JSON.stringify(value)).digest("hex");
        const hasPending = (x: any) => Object.keys(x.pendingUploadOperations || {}).length ||
            Object.keys(x.pendingDeletions || {}).length || Object.keys(x.pendingConflicts || {}).length ||
            x.pendingSharedUploadOperation || x.emptyAppFolderBootstrap || x.firstConflictPreservationJournal ||
            (x.identicalRegistrationJournal && x.identicalRegistrationJournal.phase !== "COMPLETED");
        if (settings.runAllowed || settings.enabled || hasPending(settings) || (local.recoveryEntries || []).length ||
            hash(archive.priorOneDrive) !== archive.priorDigest ||
            JSON.stringify(archive.isolation) !== JSON.stringify(local.oneDriveScopeIsolation) ||
            JSON.stringify(archive.priorOneDrive) !== JSON.stringify(archive.isolation.prior))
            throw new Error("独立副本旧完整历史、未决事项或当前许可已变化");
        const registration = settings.firstCommitJournal as FirstCommitJournal | undefined;
        if (registration?.priorBindingArchive) {
            if (JSON.stringify(registration.priorBindingArchive) !== JSON.stringify(archive) ||
                registration.identity.accountId !== archive.accountId || registration.identity.deviceId !== archive.deviceId ||
                registration.identity.cloudVaultId !== archive.newCloudVaultId ||
                (settings.firstFileJournal && settings.firstFileJournal.registrationOpId !== registration.ticketId))
                throw new Error("独立副本新原操作与旧归档不一致");
        } else if (JSON.stringify(settings.firstCommitJournal) !== JSON.stringify(archive.priorOneDrive.firstCommitJournal) ||
            JSON.stringify(settings.firstFileJournal) !== JSON.stringify(archive.priorOneDrive.firstFileJournal))
            throw new Error("旧完成事务记录已变化");
        const original = JSON.stringify(archive.priorOneDrive.fileStates || {});
        const current = JSON.stringify(settings.fileStates || {});
        const accepted = (local.oneDriveBindingHistory || []).find((h: any) =>
            h.cloudVaultId === archive.newCloudVaultId && h.priorDigest === archive.priorDigest);
        if (current !== original && (!accepted || JSON.stringify(accepted.newFileStates) !== current))
            throw new Error("独立副本提交前旧基线已变化");
    }

    private async captureIndependentCopyArchive(newCloudVaultId: string): Promise<any | null> {
        const settings = this.settings();
        if (!Object.keys(settings.fileStates || {}).length && !settings.firstCommitJournal && !settings.firstFileJournal)
            return null;
        requireWindowsHost();
        const local: any = (this.store as any).settings.local;
        const receipt = local.oneDriveScopeIsolation;
        const prior = receipt?.prior;
        const pending = (x: any) => Boolean(x && (Object.keys(x.pendingUploadOperations || {}).length ||
            Object.keys(x.pendingDeletions || {}).length || Object.keys(x.pendingConflicts || {}).length ||
            x.pendingSharedUploadOperation || x.emptyAppFolderBootstrap || x.firstConflictPreservationJournal ||
            (x.identicalRegistrationJournal && x.identicalRegistrationJournal.phase !== "COMPLETED")));
        if (!prior || prior.firstSyncCompleted !== true || prior.confirmed !== true ||
            settings.firstSyncCompleted || settings.confirmed || settings.runAllowed || settings.enabled ||
            receipt.accountId !== local.microsoftAccountId || prior.confirmationAccountId !== local.microsoftAccountId ||
            receipt.deviceId !== settings.deviceId || receipt.proofId !== settings.installationProofId ||
            pending(prior) || pending(settings) || (local.recoveryEntries || []).length ||
            JSON.stringify(settings.fileStates || {}) !== JSON.stringify(prior.fileStates || {}) ||
            JSON.stringify(settings.firstCommitJournal) !== JSON.stringify(prior.firstCommitJournal) ||
            JSON.stringify(settings.firstFileJournal) !== JSON.stringify(prior.firstFileJournal) ||
            newCloudVaultId === prior.cloudVaultId || newCloudVaultId === prior.vaultId ||
            !newCloudVaultId.startsWith("vault-"))
            throw new Error("旧完成历史、新根隔离事实或全新目标不可核实；未决事项不得归档绕锁");
        const oldCommit = prior.firstCommitJournal, oldFiles = prior.firstFileJournal;
        if (oldCommit || oldFiles) {
            if (!oldCommit || !oldFiles || oldCommit.stage !== "REGISTERED" || oldFiles.stage !== "FILES_COPIED" ||
                oldFiles.registrationOpId !== oldCommit.ticketId || oldCommit.identity.accountId !== receipt.accountId ||
                oldCommit.identity.cloudVaultId !== prior.cloudVaultId || oldFiles.cloudVaultId !== prior.cloudVaultId ||
                oldFiles.accountId !== receipt.accountId || oldFiles.localVaultId !== oldCommit.identity.localVaultId ||
                oldFiles.deviceId !== oldCommit.identity.deviceId ||
                JSON.stringify(Object.keys(oldFiles.completed || {}).sort()) !== JSON.stringify(Object.keys(prior.fileStates || {}).sort()) ||
                Object.entries(prior.fileStates || {}).some(([p, value]: [string, any]) =>
                    !value.localHash || !value.remoteId || !value.remoteEtag || oldFiles.completed[p]?.id !== value.remoteId ||
                    oldFiles.completed[p]?.etag !== value.remoteEtag || !oldFiles.files.some((f: any) => f.path === p && f.sha256 === value.localHash)))
                throw new Error("旧首事务不是相互一致的已完成历史，不能开始新独立副本");
        }
        const fs = require("fs") as typeof import("fs");
        const path = require("path").win32 as typeof import("path").win32;
        const root = path.normalize(receipt.physicalRoot);
        const audit = await inspectReceiverCandidateInventory(this.app, null, undefined, {
            accountId: receipt.accountId, vaultId: settings.vaultId, deviceId: settings.deviceId,
            installationProofId: settings.installationProofId,
        });
        if (!audit.independentPhysicalRootVerified) throw new Error("复制目标不是物理独立Vault");
        const stat = fs.lstatSync(root);
        const archive = {
            accountId: receipt.accountId, deviceId: settings.deviceId, proofId: settings.installationProofId,
            physicalRoot: root, rootPhysicalIdentity: [String(stat.dev), String(stat.ino), String(stat.birthtimeMs)].join(":"),
            newCloudVaultId, priorOneDrive: JSON.parse(JSON.stringify(prior)),
            priorDigest: await contentHash(JSON.stringify(prior)), isolation: JSON.parse(JSON.stringify(receipt)),
        };
        this.assertIndependentCopyScope(archive);
        return archive;
    }

    // D02-B3: staging does not authorize any remote or local write.
    private async firstConnectionSnapshotDigest(): Promise<string> {
        return contentHash(JSON.stringify((await this.captureFirstConnectionSnapshot()).snapshot));
    }

    private async stageFirstConnectionCommit(
        input: FirstCommitProposal, generation = this.resumedApprovalGeneration,
    ): Promise<FirstCommitTicket> {
        const settings = this.settings();
        const accountId = String((this.store as any).settings.local.microsoftAccountId || "");
        const priorBindingArchive = await this.captureIndependentCopyArchive(input?.cloudVaultId || "");
        if (settings.firstSyncCompleted || ((settings as any).firstCommitJournal && !priorBindingArchive)) {
            throw new Error("首次同步状态或历史提交尚未对账，不能签发新票据。");
        }
        if (!input || input.mode !== "create_new" || !input.registryEtag || !input.snapshotDigest
            || !input.cloudVaultId.startsWith("vault-") || !input.cloudVaultName
            || !this.pendingFirstChoice || this.pendingFirstChoice.mode !== "new"
            || this.pendingFirstChoice.id !== input.cloudVaultId
            || this.previewVaultId !== input.cloudVaultId
            || input.localVaultId !== settings.vaultId || !input.localVaultId
            || input.cloudVaultId === input.localVaultId
            || input.deviceId !== settings.deviceId || !input.deviceId
            || input.accountId !== accountId || !accountId) {
            throw new Error("确认前身份不完整或云端知识库不匹配，禁止签发事务票据。");
        }
        if (input.expiresAt <= Date.now() || input.expiresAt > Date.now() + 15 * 60_000) {
            throw new Error("预览有效期无效，必须重新检查。");
        }
        if (!Array.isArray(input.uploads) || !Array.isArray(input.downloads)
            || !Array.isArray(input.conflicts) || !Array.isArray(input.unknown)
            || input.downloads.length || input.conflicts.length || input.unknown.length) {
            throw new Error("存在冲突、未知或未经批准的下载范围，不能提交当前首次连接。");
        }
        const validPaths = input.uploads.every((name) => typeof name === "string"
            && name && !name.startsWith("/") && !name.includes("\\")
            && !name.split("/").some((part) => !part || part === "." || part === ".."));
        const uniquePaths = new Set(input.uploads);
        const currentPaths = this.localFiles().map((item) => item.path).sort();
        if (!validPaths || uniquePaths.size !== input.uploads.length
            || JSON.stringify([...uniquePaths].sort()) !== JSON.stringify(currentPaths)) {
            throw new Error("上传范围不完整、重复或路径不安全，禁止签发票据。");
        }
        if (!Array.isArray(input.registryVaults)
            || input.registryVaults.some((item) => !item?.id || item.id === input.cloudVaultId)
            || new Set(input.registryVaults.map((item) => item.id)).size !== input.registryVaults.length) {
            throw new Error("云端知识库注册身份异常，不能直接建立独立副本。");
        }
        const currentRegistry = await this.readSharedState("vaults/registry.json");
        const registryMeta = await this.readRemoteFileMetadata("vaults/registry.json", true);
        if (!Array.isArray(currentRegistry?.vaults) || !registryMeta
            || registryMeta.etag !== input.registryEtag
            || JSON.stringify(currentRegistry.vaults) !== JSON.stringify(input.registryVaults)) {
            throw new Error("云端注册表已变化，不能签发旧预览票据。");
        }
        if ((await this.firstConnectionSnapshotDigest()) !== input.snapshotDigest) {
            throw new Error("本地或云端预览版本已过期，不能签发票据。");
        }
        const proposal = JSON.parse(JSON.stringify(input)) as FirstCommitProposal;
        if (priorBindingArchive) proposal.priorBindingArchive = priorBindingArchive;
        const ticket = Object.freeze({
            ticketId: typeof crypto.randomUUID === "function" ? crypto.randomUUID()
                : Date.now().toString(36) + "-" + Math.random().toString(36).slice(2),
            digest: await contentHash(JSON.stringify(proposal)),
            expiresAt: proposal.expiresAt,
        });
        if (generation !== this.resumedApprovalGeneration) throw new Error("首次注册审批已撤销");
        this.firstCommitStaged = { ticket, proposal, generation };
        return ticket;
    }

    // Only the future D03 final-click handler may call this, never selection/preview.
    private approveFirstConnection(ticketId: string, digest: string, confirmed: boolean): object {
        const staged = this.firstCommitStaged;
        if (!confirmed || !staged || staged.ticket.ticketId !== ticketId
            || staged.generation !== this.resumedApprovalGeneration
            || staged.ticket.digest !== digest || Date.now() >= staged.ticket.expiresAt) {
            throw new Error("未取得对应预览的最终人工确认，不得授权首次提交。");
        }
        const consent = Object.freeze({ ticketId, digest, issuedAt: Date.now() });
        this.firstCommitConsents.add(consent);
        return consent;
    }

    private async commitFirstConnection(ticket: FirstCommitTicket, consent: object | null): Promise<FirstCommitDecision> {
        const staged = this.firstCommitStaged;
        const settings = this.settings();
        const blocked = (reason: string): FirstCommitDecision => ({ status: "BLOCKED", reason });
        if (!ticket || !consent || typeof consent !== "object"
            || !this.firstCommitConsents.has(consent) || !staged
            || staged.generation !== this.resumedApprovalGeneration
            || ticket.ticketId !== staged.ticket.ticketId
            || ticket.digest !== staged.ticket.digest || ticket.expiresAt !== staged.ticket.expiresAt
            || Date.now() >= ticket.expiresAt || settings.firstSyncCompleted
            || ((settings as any).firstCommitJournal && !staged?.proposal.priorBindingArchive) || this.firstCommitInFlight) {
            return blocked("最终确认票据失效、已使用或存在未对账事务。");
        }
        const proposal = staged.proposal;
        if (proposal.priorBindingArchive && JSON.stringify(await this.captureIndependentCopyArchive(proposal.cloudVaultId)) !==
            JSON.stringify(proposal.priorBindingArchive)) return blocked("旧完成历史或新独立作用域已变化");
        const local = (this.store as any).settings.local;
        if (proposal.accountId !== String(local.microsoftAccountId || "")
            || proposal.localVaultId !== settings.vaultId
            || proposal.deviceId !== settings.deviceId
            || this.pendingFirstChoice?.id !== proposal.cloudVaultId
            || this.previewVaultId !== proposal.cloudVaultId) {
            return blocked("微软账号、设备或知识库身份已变化，禁止执行旧计划。");
        }
        // Only one caller may consume the in-memory capability and submit remotely.
        this.firstCommitConsents.delete(consent);
        this.approvedFirstCommitTicketId = ticket.ticketId;
        const job = this.performFirstRegistryCommit(ticket, proposal);
        this.firstCommitInFlight = job;
        try {
            return await job;
        } finally {
            if (this.firstCommitInFlight === job) this.firstCommitInFlight = null;
            if (this.approvedFirstCommitTicketId === ticket.ticketId) this.approvedFirstCommitTicketId = null;
            this.firstCommitStaged = null;
        }
    }

    private async performFirstRegistryCommit(
        ticket: FirstCommitTicket, proposal: FirstCommitProposal,
    ): Promise<FirstCommitDecision> {
        const blocked = (reason: string): FirstCommitDecision => ({ status: "BLOCKED", reason });
        const settings = this.settings();
        if (!this.approvedFirstCommitTicketId
            || this.approvedFirstCommitTicketId !== ticket?.ticketId
            || this.firstCommitStaged?.ticket.ticketId !== ticket?.ticketId
            || this.firstCommitStaged?.ticket.digest !== ticket?.digest
            || ((settings as any).firstCommitJournal && !proposal.priorBindingArchive)) {
            return blocked("事务执行入口缺少有效的最终用户确认能力票据。");
        }
        try {
            this.assertFirstCommitIdentity(proposal);
            await this.assertFirstCommitGraphAccount(proposal.accountId);
            if ((await this.firstConnectionSnapshotDigest()) !== proposal.snapshotDigest) {
                return blocked("预览版本失效，需重新检查。");
            }
            const registryMeta = await this.readRemoteFileMetadata("vaults/registry.json", true);
            const registry = await this.readSharedState("vaults/registry.json");
            if (this.firstCommitStaged?.generation !== this.resumedApprovalGeneration ||
                registryMeta?.etag !== proposal.registryEtag
                || JSON.stringify(registry?.vaults) !== JSON.stringify(proposal.registryVaults)) {
                return blocked("远端注册表 ETag 或身份发生变化。");
            }
        } catch {
            return blocked("最终提交前的只读版本核验失败，未写入。");
        }
        const journal: FirstCommitJournal = {
            ...(proposal.priorBindingArchive ? { priorBindingArchive: proposal.priorBindingArchive } : {}),
            ticketId: ticket.ticketId,
            digest: ticket.digest,
            identity: {
                accountId: proposal.accountId, localVaultId: proposal.localVaultId,
                deviceId: proposal.deviceId, cloudVaultId: proposal.cloudVaultId,
            },
            registryEtag: proposal.registryEtag,
            cloudVaultName: proposal.cloudVaultName,
            registryBefore: JSON.parse(JSON.stringify(proposal.registryVaults)),
            stage: "REGISTERING",
            updatedAt: Date.now(),
        };
        (settings as any).firstCommitJournal = journal;
        if (proposal.priorBindingArchive) {
            // Original completed journals remain recoverable inside this new durable operation.
            delete settings.firstFileJournal;
            delete settings.identicalRegistrationJournal;
        }
        try {
            if (proposal.priorBindingArchive) {
                await (this.store as any).saveWithIdentityGuard(() => this.assertIndependentCopyScope(proposal.priorBindingArchive));
                const persisted = await (this.store as any).readPersistedOneDriveRegistrationState(() =>
                    this.assertIndependentCopyScope(proposal.priorBindingArchive));
                if (persisted?.accountId !== proposal.accountId ||
                    JSON.stringify(persisted.oneDrive?.firstCommitJournal) !== JSON.stringify(journal) ||
                    JSON.stringify(persisted.oneDrive.fileStates) !== JSON.stringify(proposal.priorBindingArchive.priorOneDrive.fileStates))
                    throw new Error("独立副本原操作与旧归档未实际持久核实");
            } else await (this.store as any).save(); // Durable before the first remote write.
        } catch {
            if (proposal.priorBindingArchive) {
                journal.stage = "RECONCILE_REQUIRED";
                try { await (this.store as any).saveWithIdentityGuard(() => this.assertIndependentCopyScope(proposal.priorBindingArchive)); } catch {}
                return blocked("新副本原操作持久结果待核实，旧历史和基线保留，未执行云端写入。");
            }
            delete (settings as any).firstCommitJournal;
            return blocked("无法持久保存提交票据，未执行任何云端写入。");
        }
        this.activeRegistrationTicket = ticket.ticketId;
        try {
            this.assertFirstCommitIdentity(proposal);
            if (this.firstCommitStaged?.generation !== this.resumedApprovalGeneration) {
                throw new Error("首次注册审批已撤销");
            }
            await this.registerVault(
                { id: proposal.cloudVaultId, name: proposal.cloudVaultName, updatedAt: Date.now(), registrationOpId: ticket.ticketId },
                { etag: proposal.registryEtag, vaults: proposal.registryVaults },
            );
            journal.stage = "REGISTERED";
            journal.updatedAt = Date.now();
            await (this.store as any).save();
            // Metadata registered only; no note transfer, local binding, or automatic sync.
            return { status: "REGISTERED" };
        } catch {
            journal.stage = "RECONCILE_REQUIRED";
            journal.updatedAt = Date.now();
            try { await (this.store as any).save(); } catch {}
            // Never blindly retry a PUT: the server may already have committed it.
            return { status: "RECONCILE_REQUIRED", reason: "注册提交结果未知，需要只读对账，禁止自动重试。" };
        } finally {
            this.activeRegistrationTicket = null;
        }
    }

    private async inspectFirstCommitJournal(): Promise<{
        stage: string; cloudVaultId: string;
    } | null> {
        const j = (this.settings() as any).firstCommitJournal as FirstCommitJournal | undefined;
        if (!j) return null;
        return { stage: j.stage, cloudVaultId: j.identity.cloudVaultId };
    }

    // After a crash / network ambiguity, only GET and a verified local journal update are allowed.
    // This method never resubmits the previous PUT or starts file transfers.
    private async reconcileFirstCommitJournal(): Promise<FirstCommitDecision> {
        const journal = (this.settings() as any).firstCommitJournal as FirstCommitJournal | undefined;
        if (!journal) return { status: "BLOCKED", reason: "本设备无待对账事务。" };
        const local = (this.store as any).settings.local;
        if (String(local.microsoftAccountId || "") !== journal.identity.accountId
            || this.settings().vaultId !== journal.identity.localVaultId
            || this.settings().deviceId !== journal.identity.deviceId) {
            return { status: "BLOCKED", reason: "设备或账号已变，禁止继承历史提交许可。" };
        }
        try {
            const registry = await this.readSharedState("vaults/registry.json");
            if (!Array.isArray(registry?.vaults)) {
                return { status: "RECONCILE_REQUIRED", reason: "远端注册状态异常。" };
            }
            const matches = registry.vaults.filter((item: any) =>
                item?.id === journal.identity.cloudVaultId);
            const earlierVaults = registry.vaults.filter((item: any) =>
                item?.id !== journal.identity.cloudVaultId);
            if (matches.length !== 1
                || matches[0].name !== journal.cloudVaultName
                || matches[0].registrationOpId !== journal.ticketId
                || JSON.stringify(earlierVaults) !== JSON.stringify(journal.registryBefore)) {
                return { status: "RECONCILE_REQUIRED", reason: "无法确认本事务云端登记，请人工对账。" };
            }
            if (journal.stage !== "REGISTERED") {
                journal.stage = "REGISTERED";
                journal.updatedAt = Date.now();
                await (this.store as any).save();
            }
            return { status: "REGISTERED" };
        } catch {
            return { status: "RECONCILE_REQUIRED", reason: "只读对账未完成，禁止自动重试。" };
        }
    }

    private async assertFirstCommitGraphAccount(accountId: string, scope?: FirstUploadReadScope): Promise<void> {
        const identity: any = await this.graph("GET", "/me", undefined, undefined, scope?.assertCurrent);
        if (!accountId || !identity?.id || String(identity.id) !== accountId) {
            throw new Error("Microsoft Graph 当前账户与已批准账户不一致，停止云端提交。");
        }
    }

    private assertFirstCommitIdentity(proposal: FirstCommitProposal): void {
        if (this.identicalRegistrationBlocked() && !proposal.priorBindingArchive) throw new Error("原同内容登记未核实，禁止注册新云库");
        if (proposal.priorBindingArchive) {
            if (proposal.cloudVaultId !== proposal.priorBindingArchive.newCloudVaultId) throw new Error("旧cloud不能作为新独立目标");
            this.assertIndependentCopyScope(proposal.priorBindingArchive);
        }
        const current = (this.store as any).settings.local;
        const settings = this.settings();
        if ((Platform.isAndroidApp && !this.store.installationProofIsValid()) ||
            String(current.microsoftAccountId || "") !== proposal.accountId
            || settings.vaultId !== proposal.localVaultId
            || settings.deviceId !== proposal.deviceId
            || settings.firstSyncCompleted === true
            || this.previewVaultId !== proposal.cloudVaultId) {
            throw new Error("最终提交的微软账号或知识库身份已变化，禁止写入。");
        }
    }

    private async firstFileHash(data: ArrayBuffer): Promise<string> {
        const digest = await crypto.subtle.digest("SHA-256", data);
        return Array.from(new Uint8Array(digest)).map((byte) => byte.toString(16).padStart(2, "0")).join("");
    }

    private async currentFirstFileEntries(scope?: FirstUploadReadScope, onFile?: (done: number, total: number) => void): Promise<FirstFileEntry[]> {
        const local = this.localFiles().sort((a, b) => a.path.localeCompare(b.path));
        const items: FirstFileEntry[] = [];
        const archivedBinding = Boolean(this.settings().firstCommitJournal?.priorBindingArchive);
        if (archivedBinding && this.skippedLocalPaths().size) {
            throw new Error("存在未批准的大文件跳过项，不能宣称独立云端副本完整。");
        }
        for (const item of local) {
            scope?.assertCurrent();
            if (!item.path || item.path.startsWith("/") || item.path.includes("\\")
                || item.path !== normalizePath(item.path)
                || item.path.split("/").some((part) => !part || part === "." || part === "..")
                || (archivedBinding && item.size >= 4_000_000)) {
                throw new Error("文件路径或大小尚不支持条件原子提交，已停止首次文件传输。");
            }
            const data = await this.app.vault.readBinary(item.file);
            if (data.byteLength !== item.size) throw new Error("本地文件扫描与字节数据不一致，已停止。");
            items.push({ path: item.path, size: item.size, mtime: item.mtime, sha256: await this.firstFileHash(data) });
            onFile?.(items.length, local.length);
        }
        if (new Set(items.map((item) => item.path)).size !== items.length) {
            throw new Error("本地文件路径重复，已停止首次传输。");
        }
        return items;
    }

    private assertFirstFileIdentity(journal: FirstFileJournal): void {
        if (this.identicalRegistrationBlocked()) throw new Error("原同内容登记未核实，禁止首轮文件传输");
        const settings = this.settings();
        const account = (this.store as any).settings.local;
        const registered = settings.firstCommitJournal as FirstCommitJournal | undefined;
        if (registered?.priorBindingArchive) {
            if (registered.priorBindingArchive.newCloudVaultId !== journal.cloudVaultId) throw new Error("新独立副本目标与原批准不同");
            this.assertIndependentCopyScope(registered.priorBindingArchive);
        }
        if ((Platform.isAndroidApp && !this.store.installationProofIsValid()) ||
            settings.firstSyncCompleted || !registered || registered.stage !== "REGISTERED"
            || registered.ticketId !== journal.registrationOpId
            || registered.identity.cloudVaultId !== journal.cloudVaultId
            || registered.identity.localVaultId !== journal.localVaultId
            || registered.identity.accountId !== journal.accountId
            || registered.identity.deviceId !== journal.deviceId
            || String(account.microsoftAccountId || "") !== journal.accountId
            || settings.vaultId !== journal.localVaultId || settings.deviceId !== journal.deviceId) {
            throw new Error("账号、本地知识库或注册事务身份不一致，文件传输已阻断。");
        }
    }

    private async verifyFirstFileRegistration(journal: FirstFileJournal, scope?: FirstUploadReadScope): Promise<void> {
        this.assertFirstFileIdentity(journal);
        await this.assertFirstCommitGraphAccount(journal.accountId, scope);
        const registry: any = await this.readSharedState("vaults/registry.json", scope);
        const registered = this.settings().firstCommitJournal as FirstCommitJournal;
        if (!Array.isArray(registry?.vaults) || !Array.isArray(registered?.registryBefore)) {
            throw new Error("云端注册表不完整，禁止文件传输。");
        }
        const targets = registry.vaults.filter((item: any) => item?.id === journal.cloudVaultId);
        const remaining = registry.vaults.filter((item: any) => item?.id !== journal.cloudVaultId);
        if (targets.length !== 1 || targets[0].registrationOpId !== journal.registrationOpId
            || targets[0].name !== registered.cloudVaultName
            || JSON.stringify(remaining) !== JSON.stringify(registered.registryBefore)
            || await contentHash(JSON.stringify(registry)) !== journal.registryDigest) {
            throw new Error("云端注册版本、身份或其他知识库内容已变化，禁止传输。");
        }
    }

    private async readFirstCopyRemoteFiles(journal: FirstFileJournal, scope?: FirstUploadReadScope, onPage?: (page: number) => void): Promise<RemoteFile[]> {
        const items = await this.remoteFiles(true, true, { vaultId: journal.cloudVaultId, scope, onPage });
        if (!Array.isArray(items) || items.some((item) => !item.id || !item.path || !item.etag
            || !Number.isFinite(item.size))
            || new Set(items.map((item) => item.path)).size !== items.length) {
            throw new Error("云端文件列表缺失身份或不完整，禁止传输。");
        }
        return items;
    }

    private async verifyFirstFileContent(
        journal: FirstFileJournal, requireAll: boolean, checkRemoteBytes = true,
    ): Promise<void> {
        this.assertFirstFileIdentity(journal);
        const current = await this.currentFirstFileEntries();
        if (JSON.stringify(current) !== JSON.stringify(journal.files)) {
            throw new Error("本地文件路径、内容 SHA256 或版本发生变化，旧上传票据已失效。");
        }
        const rules = await contentHash(JSON.stringify([this.maxBytes(), this.settings().ignorePatterns || []]));
        if (rules !== journal.ruleDigest) throw new Error("同步排除规则已变化，旧票据失效。");
        await this.verifyFirstFileRegistration(journal);
        const remote = await this.readFirstCopyRemoteFiles(journal);
        const uploaded = journal.completed || {};
        if (remote.length !== Object.keys(uploaded).length
            || (requireAll && Object.keys(uploaded).length !== journal.files.length)) {
            throw new Error("云端文件数与已批准提交日志不一致，必须只读对账。");
        }
        for (const item of remote) {
            const approved = journal.files.find((file) => file.path === item.path);
            const completed = uploaded[item.path];
            if (!approved || !completed || completed.id !== item.id || completed.etag !== item.etag
                || approved.size !== item.size
                || (checkRemoteBytes && await this.firstFileHash(await this.remoteBytes(item)) !== approved.sha256)) {
                throw new Error("云端存在未经批准、被替换或内容不同的文件，不能覆盖或提交绑定。");
            }
        }
    }

    // This stages only a NEW, empty, registered cloud copy. It never writes.
    private async prepareFirstFileTransfer(approvedVersions: string[], generation: number): Promise<FirstCommitTicket> {
        const settings = this.settings();
        const registration = settings.firstCommitJournal as FirstCommitJournal | undefined;
        if (!registration || registration.stage !== "REGISTERED" || settings.firstFileJournal
            || this.firstFileJob || this.firstFilePrepared || this.running) {
            throw new Error("云端注册事务尚未安全完成或存在未对账的首次文件事务。");
        }
        const journal: FirstFileJournal = {
            ticketId: "", digest: "", registrationOpId: registration.ticketId,
            accountId: registration.identity.accountId, localVaultId: registration.identity.localVaultId,
            cloudVaultId: registration.identity.cloudVaultId, deviceId: registration.identity.deviceId,
            registryDigest: "", ruleDigest: "", files: [], activeUploadPath: null, completed: {},
            stage: "TRANSFERRING", updatedAt: Date.now(),
        };
        this.assertFirstFileIdentity(journal);
        await this.assertFirstCommitGraphAccount(journal.accountId);
        const registry = await this.readSharedState("vaults/registry.json");
        journal.registryDigest = await contentHash(JSON.stringify(registry));
        await this.verifyFirstFileRegistration(journal);
        if ((await this.readFirstCopyRemoteFiles(journal)).length !== 0) {
            throw new Error("新建云端目标已存在文件或内容冲突，禁止作为空副本上传。");
        }
        journal.files = await this.currentFirstFileEntries();
        const currentVersions = journal.files
            .map(file => JSON.stringify([file.path, file.size, file.mtime, file.sha256])).sort();
        if (generation !== this.resumedApprovalGeneration ||
            JSON.stringify(currentVersions) !== JSON.stringify(approvedVersions)) {
            throw new Error("注册后本机预览版本或许可发生变化，禁止发出文件上传");
        }
        journal.ruleDigest = await contentHash(JSON.stringify([this.maxBytes(), settings.ignorePatterns || []]));
        const ticket: FirstCommitTicket = Object.freeze({
            ticketId: typeof crypto.randomUUID === "function" ? crypto.randomUUID()
                : Date.now().toString(36) + "-" + Math.random().toString(36).slice(2),
            digest: await contentHash(JSON.stringify(journal)), expiresAt: Date.now() + 5 * 60_000,
        });
        journal.ticketId = ticket.ticketId;
        journal.digest = ticket.digest;
        this.firstFilePrepared = { ticket, journal, generation };
        return ticket;
    }

    private approveFirstFileTransfer(ticketId: string, digest: string, confirmed: boolean): object {
        const staged = this.firstFilePrepared;
        if (!confirmed || !staged || staged.ticket.ticketId !== ticketId || staged.ticket.digest !== digest
            || staged.generation !== this.resumedApprovalGeneration
            || Date.now() >= staged.ticket.expiresAt || this.firstFileJob) {
            throw new Error("尚未获得与首次文件预览一致的最终人工授权。");
        }
        const consent = Object.freeze({ ticketId, digest, confirmedAt: Date.now() });
        this.firstFileConsents.add(consent);
        return consent;
    }

    private async executeFirstFileTransfer(
        ticket: FirstCommitTicket | null, consent: object | null,
    ): Promise<FileCopyDecision> {
        const blocked = (reason: string): FileCopyDecision => ({ status: "BLOCKED", reason });
        const staged = this.firstFilePrepared;
        if (!ticket || !consent || typeof consent !== "object" || !this.firstFileConsents.has(consent)
            || !staged || !staged.ticket
            || staged.generation !== this.resumedApprovalGeneration
            || ticket.ticketId !== staged.ticket.ticketId || ticket.digest !== staged.ticket.digest
            || ticket.expiresAt !== staged.ticket.expiresAt || Date.now() >= ticket.expiresAt
            || this.settings().firstFileJournal || this.firstFileJob) {
            return blocked("缺少已确认且未使用的完整文件级同步票据。");
        }
        // Consume the capability before the first await. Duplicate clicks cannot enter.
        this.firstFileConsents.delete(consent);
        const work = this.performApprovedFirstFileTransfer(staged.journal, ticket, staged.generation);
        this.firstFileJob = work;
        try {
            return await work;
        } finally {
            if (this.firstFileJob === work) this.firstFileJob = null;
            this.firstFilePrepared = null;
        }
    }

    private async performApprovedFirstFileTransfer(
        journal: FirstFileJournal, ticket: FirstCommitTicket, generation: number,
        assertPreviousCurrent?: () => void,
    ): Promise<FileCopyDecision> {
        const blocked = (reason: string): FileCopyDecision => ({ status: "BLOCKED", reason });
        try {
            if (Date.now() >= ticket.expiresAt) return blocked("最终批准已过期，必须重新预览。");
            await this.verifyFirstFileContent(journal, false);
            assertPreviousCurrent?.();
            if (generation !== this.resumedApprovalGeneration) return blocked("首次文件上传许可已撤销");
        } catch {
            return blocked("提交前本地/云端版本或账号发生变化，未执行写入。");
        }
        const settings = this.settings();
        // The plan is durable before any remote folder creation or content PUT.
        const previous = settings.firstFileJournal;
        settings.firstFileJournal = JSON.parse(JSON.stringify(journal)) as FirstFileJournal;
        const durable = settings.firstFileJournal as FirstFileJournal;
        try {
            await (this.store as any).save();
        } catch {
            if (previous) settings.firstFileJournal = previous;
            else delete settings.firstFileJournal;
            return blocked("本地文件事务日志无法持久保存，云端保持不变。");
        }
        try {
            const activePath = durable.activeUploadPath;
            const active = activePath ? durable.files.find(file => file.path === activePath) : undefined;
            const files = active ? [active, ...durable.files.filter(file => file !== active)] : durable.files;
            for (const file of files) {
                if (Object.prototype.hasOwnProperty.call(durable.completed, file.path)) continue;
                // Read and validate current state before each individual write.
                // Recheck all approved identities, metadata and local hashes before every write.
                // Download remote bytes once in final verification to avoid N² content fetches.
                await this.verifyFirstFileContent(durable, false, false);
                const local = this.app.vault.getAbstractFileByPath(file.path);
                if (!(local instanceof TFile)) throw new Error("本地源文件不再存在。");
                const bytes = await this.app.vault.readBinary(local);
                if (bytes.byteLength !== file.size || await this.firstFileHash(bytes) !== file.sha256) {
                    throw new Error("本地文件内容已变化，旧票据不可继续。");
                }
                const previous = await this.readFirstCopyRemoteFiles(durable);
                if (previous.some((remote) => remote.path === file.path)) {
                    throw new Error("云端出现同路径文件，拒绝覆盖。");
                }
                await this.assertFirstCommitGraphAccount(durable.accountId);
                this.assertFirstFileIdentity(durable);
                // Persist the only possibly dispatched file before attempting its write.
                durable.activeUploadPath = file.path;
                durable.sourceMtimeByPath = { ...durable.sourceMtimeByPath, [file.path]: file.mtime };
                await (this.store as SmartMDStore).save();
                const before = this.previewVaultId;
                this.previewVaultId = durable.cloudVaultId;
                let result: any;
                try {
                    // A null destination selects an exclusive upload session for a new file.
                    const archive = (settings.firstCommitJournal as FirstCommitJournal | undefined)?.priorBindingArchive;
                    const guard = () => {
                        if (generation !== this.resumedApprovalGeneration || settings.firstFileJournal !== durable ||
                            durable.stage !== "TRANSFERRING") throw new Error("独立副本文件操作已暂停或变化");
                        this.assertFirstFileIdentity(durable);
                        if (!archive) return;
                        const fs = require("fs") as typeof import("fs");
                        const path = require("path").win32 as typeof import("path").win32;
                        const target = path.resolve(archive.physicalRoot, file.path);
                        const stat = fs.lstatSync(target);
                        if (!stat.isFile() || stat.isSymbolicLink() || path.normalize(fs.realpathSync(target)).toLowerCase() !== target.toLowerCase())
                            throw new Error("独立副本源文件物理身份变化");
                        const fd = fs.openSync(target, "r");
                        try {
                            const opened = fs.fstatSync(fd);
                            if (opened.dev !== stat.dev || opened.ino !== stat.ino || opened.birthtimeMs !== stat.birthtimeMs ||
                                (require("node:crypto") as typeof import("node:crypto")).createHash("sha256").update(fs.readFileSync(fd)).digest("hex") !== file.sha256)
                                throw new Error("取得凭据期间本机源字节变化");
                        } finally { fs.closeSync(fd); }
                    };
                    result = await this.upload(file.path, bytes, false, null, undefined, undefined, guard, file.mtime);
                    result = await this.setUploadedModificationTime(result, file.mtime, bytes, undefined, guard);
                } finally {
                    this.previewVaultId = before;
                }
                if (!result?.id || !(result.eTag || result.etag)) {
                    throw new Error("云端上传结果缺少 ID 或 ETag，必须只读对账。");
                }
                Object.defineProperty(durable.completed, file.path, {
                    value: { id: String(result.id), etag: String(result.eTag || result.etag) }, enumerable: true,
                });
                durable.activeUploadPath = null;
                durable.updatedAt = Date.now();
                await (this.store as any).save();
            }
            await this.verifyFirstFileContent(durable, true);
            if (generation !== this.resumedApprovalGeneration) throw new Error("首次文件上传许可已撤销");
            await this.saveVerifiedFirstFileBinding(durable);
            return { status: "FILES_COPIED" };
        } catch {
            durable.stage = "RECONCILE_REQUIRED";
            durable.updatedAt = Date.now();
            try { await (this.store as any).save(); } catch {}
            return { status: "RECONCILE_REQUIRED", reason: "首次云端副本结果不明确，禁止盲目重试，需只读对账。" };
        }
    }

    private async resumeFirstFileCopy(assertGrantCurrent: () => void, generation: number): Promise<boolean> {
        const settings = this.settings();
        const prior = settings.firstFileJournal as FirstFileJournal;
        const registered = settings.firstCommitJournal as FirstCommitJournal;
        if (!["TRANSFERRING", "RECONCILE_REQUIRED", "FILES_COPIED"].includes(prior.stage)) {
            throw new Error("首次上传事务阶段不可信，保持暂停");
        }
        const originalJournal = JSON.stringify(prior);
        const finishOnly = prior.stage === "FILES_COPIED";
        const assertReadCurrent = () => {
            assertGrantCurrent();
            if (settings.firstFileJournal !== prior) throw new Error("原文件操作已变化");
        };
        const assertPreviousCurrent = () => {
            assertReadCurrent();
            if (JSON.stringify(prior) !== originalJournal) {
                throw new Error("原文件操作已变化");
            }
        };
        // Legacy transfers were sequential and persisted each completion before
        // starting the next file. Only the first unconfirmed path could be in flight.
        const activePath = prior.activeUploadPath === undefined
            ? prior.files.find(file => !Object.prototype.hasOwnProperty.call(prior.completed, file.path))?.path
            : prior.activeUploadPath;
        const inspect = () => this.runFirstUploadInspection(assertReadCurrent, async (scope, read) => {
            assertPreviousCurrent();
            const files = finishOnly ? [] : await read(localize("正在核对本地同步文件"), progress =>
                this.currentFirstFileEntries(scope, (done, total) => progress(localize("正在核对本地同步文件：{0}/{1}", done, total))));
            const ruleDigest = finishOnly ? prior.ruleDigest
                : await contentHash(JSON.stringify([this.maxBytes(), settings.ignorePatterns || []]));
            await read(localize("正在核对 OneDrive 账号和原知识库"), () => this.verifyFirstFileRegistration(prior, scope));
            const observed: FirstFileJournal["completed"] = {};
            const remote = await read(localize("正在读取 OneDrive 文件清单"), progress =>
                this.readFirstCopyRemoteFiles(prior, scope, page => progress(localize("正在读取 OneDrive 文件清单：第 {0} 页", page))));
            const approved = new Map(prior.files.map(file => [file.path, file]));
            const current = new Map(files.map(file => [file.path, file]));
            const ordered = [...remote].sort((a, b) => a.path.localeCompare(b.path));
            let verified = 0;
            await read(localize("正在核对已上传文件：{0}/{1}", verified, ordered.length), async (refreshDeadline) => {
                for (const item of ordered) {
                    const file = approved.get(item.path);
                    const previous = Object.prototype.hasOwnProperty.call(prior.completed, item.path) ? prior.completed[item.path] : null;
                    if (!file || item.size !== file.size || (previous && (previous.id !== item.id || previous.etag !== item.etag)) ||
                        await this.firstFileHash(await this.remoteBytes(item, scope)) !== file.sha256) {
                        throw new Error("远端已有文件无法核实，保留两端并禁止重新上传");
                    }
                    const expectedMtime = Object.prototype.hasOwnProperty.call(prior.sourceMtimeByPath || {}, item.path)
                        ? prior.sourceMtimeByPath?.[item.path] : undefined;
                    if (expectedMtime !== undefined && item.mtime !== expectedMtime) {
                        throw new Error(localize("OneDrive 尚未确认文件原修改时间：{0}", redactOneDriveText(item.path)));
                    }
                    if (!finishOnly && current.get(item.path)?.sha256 !== file.sha256) {
                        throw new Error(localize("已上传文件的本地版本已改动、移走或被排除：{0}。首次同步已暂停，两端文件均已保留。", redactOneDriveText(item.path)));
                    }
                    Object.defineProperty(observed, item.path, { value: { id: item.id, etag: item.etag }, enumerable: true });
                    refreshDeadline(localize("正在核对已上传文件：{0}/{1}", ++verified, ordered.length));
                }
            });
            if (Object.keys(prior.completed).some(path => !Object.prototype.hasOwnProperty.call(observed, path))) {
                throw new Error("已上传的云端文件丢失，首次上传保持暂停");
            }
            const unresolvedPath = !finishOnly && activePath && !Object.prototype.hasOwnProperty.call(observed, activePath) ? activePath : null;
            if (unresolvedPath) {
                const original = approved.get(unresolvedPath);
                if (!original || current.get(unresolvedPath)?.sha256 !== original.sha256) {
                    throw new Error(localize("上次上传结果尚未核实，本地版本已改变：{0}。两端文件均已保留。", redactOneDriveText(unresolvedPath)));
                }
            }
            if (finishOnly) this.assertFinalizedFirstFileBaseline(prior);
            assertPreviousCurrent();
            return { files, ruleDigest, observed, unresolvedPath };
        });
        const reviewed = await inspect();
        const missing = finishOnly ? [] : reviewed.files.map(file => file.path)
            .filter(path => !Object.prototype.hasOwnProperty.call(reviewed.observed, path));
        const preview: SyncPreview = {
            localCount: reviewed.files.length, remoteCount: Object.keys(reviewed.observed).length,
            uploadCount: missing.length, downloadCount: 0, skippedCount: 0,
            uploadPaths: missing, downloadPaths: [], samePaths: finishOnly ? [] : Object.keys(reviewed.observed).sort(),
            skippedPaths: [], unknownPaths: [], unexpectedNewTarget: false,
            identity: { accountId: prior.accountId, localVaultId: prior.localVaultId, cloudVaultId: prior.cloudVaultId },
            resumedFirstUpload: true,
            finishFirstUploadOnly: finishOnly,
        };
        const expiresAt = Date.now() + 5 * 60_000;
        await new Promise<void>(resolve => new SyncPreviewModal(this.app, preview, async (progress) => {
            const latest = await inspect();
            if (Date.now() >= expiresAt || JSON.stringify(latest) !== JSON.stringify(reviewed)) {
                throw new Error("首次上传预览已过期，或本机、云端版本或同步范围已变化");
            }
            const journal: FirstFileJournal = {
                ...prior, files: latest.files, ruleDigest: latest.ruleDigest, completed: latest.observed,
                activeUploadPath: latest.unresolvedPath,
            };
            if (!finishOnly) {
                journal.stage = "TRANSFERRING";
                journal.ticketId = newId("first-resume");
                journal.digest = await contentHash(JSON.stringify({ ...journal, digest: "" }));
                const ticket = { ticketId: journal.ticketId, digest: journal.digest, expiresAt };
                progress("正在续传尚未提交的文件");
                const result = await this.performApprovedFirstFileTransfer(journal, ticket, generation, assertPreviousCurrent);
                if (result.status !== "FILES_COPIED") throw new Error(result.reason || "首次上传仍需对账");
            } else assertPreviousCurrent();
            assertGrantCurrent();
            await this.completeFirstFileCopy(prior.cloudVaultId, registered.cloudVaultName, prior.accountId, assertGrantCurrent);
            return true;
        }, resolve, true).open());
        return settings.firstSyncCompleted === true;
    }

    private async runFirstUploadInspection<T>(
        assertGrantCurrent: () => void,
        inspect: (scope: FirstUploadReadScope,
            read: <R>(label: string, task: (progress: (label: string) => void) => Promise<R>) => Promise<R>) => Promise<T>,
    ): Promise<T> {
        const controller = new AbortController();
        const inspection = { controller, label: "" };
        this.firstUploadInspection = inspection;
        let timer = 0;
        let lastPaint = 0;
        let lastActivity = 0;
        const scope: FirstUploadReadScope = {
            signal: controller.signal,
            assertCurrent: () => {
                if (controller.signal.aborted) throw controller.signal.reason;
                assertGrantCurrent();
            },
            noteActivity: () => {
                scope.assertCurrent();
                if (Date.now() - lastActivity >= 250) resetDeadline();
            },
        };
        const resetDeadline = () => {
            window.clearTimeout(timer);
            lastActivity = Date.now();
            timer = window.setTimeout(() => controller.abort(new Error(localize(
                "OneDrive 检查超时（{0}）。已保留同步记录，可以重试。", inspection.label,
            ))), this.firstUploadReadTimeoutMs);
        };
        const progress = (label: string) => {
            scope.assertCurrent();
            inspection.label = label;
            resetDeadline();
            if (Date.now() - lastPaint >= 250) {
                lastPaint = Date.now();
                this.notifyStatusChange();
            }
        };
        let onAbort: () => void = () => {};
        const aborted = new Promise<never>((_, reject) => {
            onAbort = () => reject(controller.signal.reason);
            controller.signal.addEventListener("abort", onAbort, { once: true });
        });
        try {
            const result = await Promise.race([aborted, inspect(scope, async (label, task) => {
                progress(label);
                return await task(progress);
            })]);
            scope.assertCurrent();
            return result;
        } catch (error) {
            if (controller.signal.aborted) this.firstConnectionError = safeOneDriveError(error);
            throw error;
        } finally {
            window.clearTimeout(timer);
            controller.signal.removeEventListener("abort", onAbort);
            // Invalidate any read that settled after cancellation.
            controller.abort();
            if (this.firstUploadInspection === inspection) this.firstUploadInspection = null;
            this.notifyStatusChange();
        }
    }

    private assertFinalizedFirstFileBaseline(journal: FirstFileJournal): void {
        const settings = this.settings();
        const states = settings.fileStates || {};
        if (settings.localVaultId !== journal.localVaultId || settings.cloudVaultId !== journal.cloudVaultId ||
            Object.keys(journal.completed).length !== journal.files.length || Object.keys(states).length !== journal.files.length ||
            journal.files.some(file => {
                const remote = journal.completed[file.path];
                const baseline = states[file.path];
                return !remote || !baseline || baseline.localHash !== file.sha256 ||
                    baseline.localMtime !== file.mtime || baseline.localSize !== file.size ||
                    baseline.remoteId !== remote.id || baseline.remoteEtag !== remote.etag;
            })) {
            throw new Error(localize("首次上传记录与同步基线不一致，已保留两端文件。"));
        }
    }

    private async resumeRegisteredFirstCopy(assertGrantCurrent: () => void, generation: number): Promise<boolean> {
        const settings = this.settings();
        const existing = settings.firstCommitJournal as FirstCommitJournal;
        await this.assertFirstCommitGraphAccount(existing.identity.accountId);
        const registration = await this.reconcileFirstCommitJournal();
        assertGrantCurrent();
        if (registration.status !== "REGISTERED") {
            throw new Error(registration.reason || "旧云端知识库登记尚未核实");
        }
        const current = await this.currentFirstFileEntries();
        const approvedVersions = current.map(file => JSON.stringify([file.path, file.size, file.mtime, file.sha256])).sort();
        const ticket = await this.prepareFirstFileTransfer(approvedVersions, generation);
        const identity = existing.identity;
        const preview: SyncPreview = {
            localCount: current.length, remoteCount: 0, uploadCount: current.length,
            downloadCount: 0, skippedCount: 0,
            uploadPaths: current.map(file => file.path), downloadPaths: [], samePaths: [],
            skippedPaths: [], unknownPaths: [], unexpectedNewTarget: false, identity,
        };
        try {
            assertGrantCurrent();
            await new Promise<void>(resolve => new SyncPreviewModal(this.app, preview, async () => {
                assertGrantCurrent();
                const consent = this.approveFirstFileTransfer(ticket.ticketId, ticket.digest, true);
                const outcome = await this.executeFirstFileTransfer(ticket, consent);
                if (outcome.status !== "FILES_COPIED") throw new Error(outcome.reason || "首次上传尚未完成");
                await this.completeFirstFileCopy(identity.cloudVaultId, existing.cloudVaultName,
                    identity.accountId, assertGrantCurrent);
                return true;
            }, resolve, true).open());
            return settings.firstSyncCompleted === true;
        } finally {
            if (this.firstFilePrepared?.ticket === ticket) this.firstFilePrepared = null;
        }
    }

    private async completeFirstFileCopy(
        cloudVaultId: string, name: string, accountId: string, assertGrantCurrent: () => void,
    ): Promise<void> {
        assertGrantCurrent();
        const settings = this.settings();
        const previous = { ...settings };
        const rootId = await this.verifyFirstSyncRoot(cloudVaultId, settings.fileStates || {}, assertGrantCurrent);
        Object.assign(settings, { vaultId: cloudVaultId, confirmationVaultId: cloudVaultId,
            confirmationAccountId: accountId, vaultName: name, confirmed: true, runAllowed: true,
            firstSyncCompleted: true, status: "ready", lastSuccessAt: Date.now(),
            deletionRoot: rootId ? { accountId, vaultId: cloudVaultId, deviceId: settings.deviceId, id: rootId } : undefined,
            ...this.newBindingSyncMode() });
        try { await this.store.saveWithIdentityGuard(assertGrantCurrent); }
        catch (error) {
            if (this.settings() === settings) Object.assign(settings, previous, { runAllowed: false, enabled: false });
            throw error;
        }
    }

    private newBindingSyncMode() {
        return { autoMode: "smart", enabled: true, onStartup: true, afterSave: true,
            intervalMinutes: Platform.isMobile ? 5 : 1 } as const;
    }

    private async saveVerifiedFirstFileBinding(journal: FirstFileJournal): Promise<void> {
        const settings = this.settings();
        // A successful first upload must also persist exact per-file baselines.
        // Without these, a later incremental sync cannot distinguish an
        // existing note from an untracked local/cloud change.
        const registration = settings.firstCommitJournal as FirstCommitJournal | undefined;
        const archive = registration?.priorBindingArchive;
        if (archive) this.assertIndependentCopyScope(archive);
        if (Object.keys(settings.fileStates || {}).length && !archive) {
            throw new Error("首次独立云端副本已有文件基线，禁止覆盖历史记录。");
        }
        const current = await this.currentFirstFileEntries();
        if (JSON.stringify(current) !== JSON.stringify(journal.files)) {
            throw new Error("首次绑定前本地文件版本已变化，禁止保存旧基线。");
        }
        const completed = journal.completed || {};
        const approvedPaths = current.map((file) => file.path).sort();
        if (JSON.stringify(Object.keys(completed).sort()) !== JSON.stringify(approvedPaths)) {
            throw new Error("首次绑定缺少已核验的云端文件记录，不能标记完成。");
        }
        const baselines: Record<string, {
            localMtime: number; localSize: number; localHash: string;
            remoteId: string; remoteEtag: string;
        }> = {};
        for (const file of current) {
            const remote = completed[file.path];
            if (!remote?.id || !remote?.etag || !file.sha256) {
                throw new Error("首次绑定文件缺少哈希、云端ID或ETag，禁止保存。");
            }
            baselines[file.path] = {
                localMtime: file.mtime, localSize: file.size, localHash: file.sha256,
                remoteId: remote.id, remoteEtag: remote.etag,
            };
        }
        const prior = {
            localVaultId: settings.localVaultId,
            cloudVaultId: settings.cloudVaultId,
            fileStates: settings.fileStates,
            stage: journal.stage,
            updatedAt: journal.updatedAt,
        };
        journal.stage = "FILES_COPIED";
        journal.updatedAt = Date.now();
        settings.localVaultId = journal.localVaultId;
        settings.cloudVaultId = journal.cloudVaultId;
        settings.fileStates = baselines;
        settings.firstSyncCompleted = false;
        settings.enabled = false;
        settings.runAllowed = false;
        settings.confirmed = false;
        if (archive) {
            const local: any = (this.store as any).settings.local;
            const histories = local.oneDriveBindingHistory || [];
            const history = { operationId: registration!.ticketId, cloudVaultId: journal.cloudVaultId,
                priorDigest: archive.priorDigest, archive, newFileStates: baselines };
            const existing = histories.find((h: any) => h.operationId === history.operationId);
            if (existing && JSON.stringify(existing) !== JSON.stringify(history)) throw new Error("新绑定归档原操作不一致");
            if (!existing) local.oneDriveBindingHistory = [...histories, history];
        }
        try {
            if (archive) {
                await (this.store as any).saveWithIdentityGuard(() => this.assertIndependentCopyScope(archive));
                const persisted = await (this.store as any).readPersistedOneDriveRegistrationState(() => this.assertIndependentCopyScope(archive));
                if (JSON.stringify(persisted?.oneDrive?.fileStates) !== JSON.stringify(baselines) ||
                    persisted.oneDrive.localVaultId !== journal.localVaultId || persisted.oneDrive.cloudVaultId !== journal.cloudVaultId ||
                    persisted.oneDrive.firstFileJournal?.stage !== "FILES_COPIED" ||
                    persisted.accountId !== archive.accountId ||
                    JSON.stringify(persisted.oneDrive.firstCommitJournal) !== JSON.stringify(registration) ||
                    !persisted.oneDriveBindingHistory?.some((h: any) => h.operationId === registration!.ticketId &&
                        h.cloudVaultId === journal.cloudVaultId && h.priorDigest === archive.priorDigest &&
                        JSON.stringify(h.archive) === JSON.stringify(archive) && JSON.stringify(h.newFileStates) === JSON.stringify(baselines)))
                    throw new Error("新独立绑定与旧历史归档尚未同时实际持久核实");
            } else await (this.store as any).save();
        } catch (error) {
            if (archive) {
                // Preserve both the new original operation and prior history on ambiguous ACK/readback.
                journal.stage = "RECONCILE_REQUIRED";
                settings.runAllowed = false; settings.enabled = false; settings.firstSyncCompleted = false;
                try { await (this.store as any).saveWithIdentityGuard(() => this.assertIndependentCopyScope(archive)); } catch {}
                throw error;
            }
            // No phantom binding in memory when the durable write was rejected.
            if (prior.localVaultId === undefined) delete settings.localVaultId;
            else settings.localVaultId = prior.localVaultId;
            if (prior.cloudVaultId === undefined) delete settings.cloudVaultId;
            else settings.cloudVaultId = prior.cloudVaultId;
            settings.fileStates = prior.fileStates;
            journal.stage = prior.stage;
            journal.updatedAt = prior.updatedAt;
            throw error;
        }
    }

    private async reconcileFirstFileTransfer(): Promise<FileCopyDecision> {
        const journal = this.settings().firstFileJournal as FirstFileJournal | undefined;
        if (!journal) return { status: "BLOCKED", reason: "没有待对账的文件事务。" };
        const ruleDigest = await contentHash(JSON.stringify([
            this.maxBytes(), this.settings().ignorePatterns || [],
        ]));
        if (ruleDigest !== journal.ruleDigest) {
            return { status: "RECONCILE_REQUIRED", reason: "同步范围/排除规则已变化，旧文件事务不能自动接受。" };
        }
        if (journal.stage === "FILES_COPIED") {
            try {
                this.assertFirstFileIdentity(journal);
                await this.assertFirstCommitGraphAccount(journal.accountId);
                return { status: "FILES_COPIED" };
            } catch {
                return { status: "BLOCKED", reason: "账号或知识库身份与已完成的事务不匹配。" };
            }
        }
        try {
            // This call only performs remote GETs, content downloads and local settings update.
            // It does not replay a single PUT, POST, DELETE or download into the local Vault.
            await this.verifyFirstFileRegistration(journal);
            const local = await this.currentFirstFileEntries();
            if (JSON.stringify(local) !== JSON.stringify(journal.files)) {
                return { status: "RECONCILE_REQUIRED", reason: "本地原始内容已经变化，不能自动接受旧结果。" };
            }
            const remote = await this.readFirstCopyRemoteFiles(journal);
            if (remote.length !== journal.files.length) {
                return { status: "RECONCILE_REQUIRED", reason: "远端副本文件数量未对齐，等待人工检查。" };
            }
            const observed: Record<string, { id: string; etag: string }> = {};
            for (const item of remote) {
                const approved = journal.files.find((file) => file.path === item.path);
                if (!approved || item.size !== approved.size
                    || await this.firstFileHash(await this.remoteBytes(item)) !== approved.sha256) {
                    return { status: "RECONCILE_REQUIRED", reason: "云端副本存在未经授权内容或文件差异。" };
                }
                observed[item.path] = { id: item.id, etag: item.etag };
            }
            journal.completed = observed;
            await this.saveVerifiedFirstFileBinding(journal);
            return { status: "FILES_COPIED" };
        } catch {
            return { status: "RECONCILE_REQUIRED", reason: "只读对账失败，禁止修改已有两端内容。" };
        }
    }

    private async captureFirstConnectionSnapshot(): Promise<{
        snapshot: FirstConnectionSnapshot;
        local: LocalFile[];
        remote: RemoteFile[];
    }> {
        const choice = this.pendingFirstChoice;
        if (!choice || this.previewVaultId !== choice.id || (!this.registryReadIsTrusted && !choice.emptyAppRoot)) {
            throw new Error("首次连接身份尚未通过只读核验。");
        }
        const settings = this.settings();
        const accountId = String((this.store as any).settings.local.microsoftAccountId || "");
        if (!accountId || settings.firstSyncCompleted) throw new Error("首次连接账号或同步阶段已变化。");
        const local = this.localFiles();
        const localVersions: string[] = [];
        for (const item of local) {
            const data = choice.mergeAndroid
                ? await this.app.vault.adapter.readBinary(item.path)
                : await this.app.vault.readBinary(item.file);
            if (data.byteLength !== item.size) throw new Error("预览期间本机文件版本发生变化");
            localVersions.push(JSON.stringify([item.path, item.size, item.mtime, await this.sha256Bytes(data)]));
        }
        const folderPaths = choice.mergeAndroid ? new Set<string>() : undefined;
        const remote = await this.remoteFiles(true, true, folderPaths ? { vaultId: choice.id, folderPaths } : undefined);
        for (const item of remote) {
            if (!item.id || !item.path || !item.etag) {
                throw new Error("云端文件身份或 ETag 不完整，已停止预览。");
            }
        }
        const registry = choice.emptyAppRoot
            ? { vaults: [], emptyAppRoot: await this.inspectEmptyAppFolderForSetup(accountId) }
            : await this.readSharedState("vaults/registry.json");
        if (!Array.isArray(registry?.vaults)) throw new Error("云端注册表已变化，需重新检查。");
        const remoteIdentityExists = registry.vaults.some((item: any) => item?.id === choice.id);
        if (remoteIdentityExists !== (choice.mode === "existing")) {
            throw new Error("云端知识库注册身份已变化，需重新选择。");
        }
        return {
            snapshot: {
                accountId,
                localVaultId: String(settings.vaultId || ""),
                deviceId: String(settings.deviceId || ""),
                cloudVaultId: choice.id,
                mode: choice.mode,
                registryHash: await contentHash(JSON.stringify(registry)),
                rulesHash: await contentHash(JSON.stringify([
                    this.maxBytes(), settings.ignorePatterns || [], settings.firstSyncCompleted
                ])),
                localVersions: localVersions.sort(),
                localStructure: choice.mergeAndroid ? this.app.vault.getAllLoadedFiles()
                    .filter(file => file.path).map(file => JSON.stringify([file.path, file instanceof TFile])).sort() : undefined,
                remoteFolders: folderPaths ? [...folderPaths].sort() : undefined,
                remoteVersions: remote.map((item) =>
                    JSON.stringify([item.path, item.id, item.etag, item.size, item.mtime])
                ).sort(),
            },
            local, remote,
        };
    }

    private async assertFirstConnectionSnapshotFresh(expected: FirstConnectionSnapshot): Promise<void> {
        const updated = (await this.captureFirstConnectionSnapshot()).snapshot;
        if (JSON.stringify(updated) !== JSON.stringify(expected)) {
            throw new Error("同步预览已过期：账号、知识库身份或文件版本发生变化，请重新检查。");
        }
    }

    private async readVaultFoldersWithoutWriting(): Promise<RemoteVault[]> {
        const existing = await this.graph("GET", "/drive/special/approot:/vaults").catch((error: any) => {
            if (error?.status === 404) return null;
            throw error;
        });
        if (!existing) return [];
        if (!existing.folder) throw new Error("云端 vaults 路径不是文件夹，已停止连接。");
        let next: string | null = "/drive/special/approot:/vaults:/children?$top=200";
        const seenPages = new Set<string>();
        const folders: RemoteVault[] = [];
        const ids = new Set<string>();
        while (next) {
            if (seenPages.has(next)) throw new Error("云端知识库列表分页出现循环，已停止连接。");
            seenPages.add(next);
            const page: any = await this.graph("GET", next);
            if (!Array.isArray(page?.value)) throw new Error("云端知识库列表不完整，已停止连接。");
            for (const item of page.value) {
                if (!item?.folder || !String(item.name || "").startsWith("vault-")) continue;
                const id = String(item.name);
                if (ids.has(id)) throw new Error("云端知识库文件夹身份重复，已停止连接。");
                ids.add(id);
                folders.push({ id, name: id, updatedAt: 0 });
            }
            const link = page["@odata.nextLink"];
            if (link == null || link === "") {
                next = null;
            } else if (typeof link === "string" && link.startsWith("https://graph.microsoft.com/v1.0/")) {
                next = link;
            } else {
                throw new Error("云端知识库列表分页地址不可信，已停止连接。");
            }
        }
        return folders;
    }

    private async readVaultRegistry(): Promise<RemoteVault[]> {
        this.ensureRemoteIdentityFresh();
        if (this.vaultRegistryCache) return [...this.vaultRegistryCache];
        const payload = await this.readSharedState("vaults/registry.json");
        if (payload !== null) {
            if (!Array.isArray(payload?.vaults)) {
                throw new Error("云端注册表结构异常，已停止首次连接；不会把未知知识库当作空库。");
            }
            const vaults: RemoteVault[] = [];
            const seenIds = new Set<string>();
            for (const item of payload.vaults) {
                const id = typeof item?.id === "string" ? item.id.trim() : "";
                if (!id || seenIds.has(id)) {
                    throw new Error("云端知识库注册身份无效或重复，已停止首次连接。");
                }
                seenIds.add(id);
                vaults.push({
                    id,
                    name: typeof item.name === "string" ? item.name : id,
                    updatedAt: Number.isFinite(item.updatedAt) ? item.updatedAt : 0,
                });
            }
            this.registryReadIsTrusted = true;
            this.vaultRegistryCache = vaults;
            return [...vaults];
        }
        // Registry absent: GET-only inventory is diagnostic, never trusted binding evidence.
        this.registryReadIsTrusted = false;
        this.vaultRegistryCache = await this.readVaultFoldersWithoutWriting();
        return [...this.vaultRegistryCache];
    }

    private async registerVault(
        vault: RemoteVault,
        expected?: { etag: string; vaults: RemoteVault[] },
    ): Promise<void> {
        const journal = (this.settings() as any).firstCommitJournal as FirstCommitJournal | undefined;
        if (!this.activeRegistrationTicket || !expected?.etag || !this.registryReadIsTrusted
            || this.approvedFirstCommitTicketId !== this.activeRegistrationTicket
            || this.firstCommitStaged?.ticket.ticketId !== this.activeRegistrationTicket
            || journal?.ticketId !== this.activeRegistrationTicket
            || journal?.stage !== "REGISTERING") {
            throw new Error("缺少已持久化的最终确认事务票据，禁止写云端知识库注册表。");
        }
        // Re-read the authoritative registry and its version, never trust an earlier cache.
        const current = await this.readSharedState("vaults/registry.json");
        if (!Array.isArray(current?.vaults)
            || JSON.stringify(current.vaults) !== JSON.stringify(expected.vaults)
            || current.vaults.some((item: any) => item?.id === vault.id)) {
            throw new Error("云端注册表内容或知识库身份已变化，禁止覆盖并需要重新对账。");
        }
        const metadata = await this.readRemoteFileMetadata("vaults/registry.json", true);
        if (!metadata?.id || metadata.etag !== expected.etag) {
            throw new Error("云端注册表 ETag 变化，条件注册已停止。");
        }
        if (this.firstCommitStaged?.proposal) {
            this.assertFirstCommitIdentity(this.firstCommitStaged.proposal);
            await this.assertFirstCommitGraphAccount(this.firstCommitStaged.proposal.accountId);
        }
        const merged = [...expected.vaults, vault];
        const payload = JSON.stringify({ updatedAt: Date.now(), vaults: merged }, null, 2);
        const encoded = new TextEncoder().encode(payload);
        if (encoded.byteLength >= 4_000_000) throw new Error("注册表超出条件 PUT 支持的大小，已停止写入。");
        const registrationGuard = () => {
            const staged = this.firstCommitStaged;
            const ticketId = this.activeRegistrationTicket;
            const liveJournal = (this.settings() as any).firstCommitJournal as FirstCommitJournal | undefined;
            if (!ticketId || this.approvedFirstCommitTicketId !== ticketId ||
                staged?.generation !== this.resumedApprovalGeneration ||
                staged?.ticket.ticketId !== ticketId ||
                liveJournal?.ticketId !== ticketId ||
                liveJournal?.stage !== "REGISTERING" || !staged?.proposal) {
                throw new Error("首次注册票据已失效，禁止异步令牌等待后继续写入。");
            }
            this.assertFirstCommitIdentity(staged.proposal);
        };
        await this.upload(
            "vaults/registry.json",
            encoded.buffer,
            true,
            metadata, // upload emits If-Match expected ETag, never unconditional overwrite
            undefined, undefined, registrationGuard,
        );
        registrationGuard();
        this.vaultRegistryCache = merged;
    }

    private async oneDriveVersionWins(path: string, localMtime: number, remote: RemoteFile): Promise<boolean> {
        const verified = await this.readRemoteFileMetadata(path);
        if (!verified || verified.id !== remote.id || verified.etag !== remote.etag ||
            !verified.mtime || !Number.isFinite(localMtime) || localMtime <= 0) {
            throw new Error(localize("无法核对文件的修改时间，请重新同步。"));
        }
        return verified.mtime >= localMtime;
    }

    private async readRemoteFileMetadata(path: string, appRelative = false): Promise<RemoteFile | null> {
        const normalizedPath = normalizePath(path);
        const remotePath = appRelative
            ? normalizedPath
            : normalizePath(`${this.basePath()}/${normalizedPath}`);
        const item = await this.graph(
            "GET",
            `/drive/special/approot:/${encodePath(remotePath)}?$select=id,name,size,eTag,file,folder,fileSystemInfo,lastModifiedDateTime`,
        ).catch((error: any) => {
            if (error?.status === 404) return null;
            throw error;
        });
        if (!item) return null;
        if (!item.file) {
            throw new Error("云端存在同名文件夹，已安全停止同步，请先在 OneDrive 中确认该项目。");
        }
        return {
            id: String(item.id),
            path: normalizedPath,
            name: String(item.name || normalizedPath.split("/").pop() || normalizedPath),
            size: Number(item.size || 0),
            etag: String(item.eTag || item.etag || ""),
            mtime: this.remoteModificationTime(item),
        };
    }

    private remoteModificationTime(item: { fileSystemInfo?: { lastModifiedDateTime?: string }; lastModifiedDateTime?: string }): number | undefined {
        for (const value of [item.fileSystemInfo?.lastModifiedDateTime, item.lastModifiedDateTime]) {
            const time = value ? Date.parse(value) : NaN;
            if (Number.isFinite(time) && time > 0) return time;
        }
        return undefined;
    }

    private basePath(): string {
        // Preview uses a short-lived target; the confirmed persistent identity stays untouched.
        return `vaults/${this.previewVaultId || this.settings().vaultId}/files`;
    }

    private async token(forceRefresh = false): Promise<string> {
        const token = await (this.auth as any).getAccessToken(forceRefresh);
        if (!token) throw new Error("请先登录微软账号并授权 OneDrive");
        return token;
    }

    private async graph(
        method: string, url: string, body?: any,
        headers?: Record<string, string>, beforeDispatch?: () => void,
    ): Promise<any> {
        this.validateGraphEndpoint(url);
        // MicrosoftAuth.requestGraph calls this synchronously after every
        // access-token/401-refresh await, immediately before requestUrl.
        beforeDispatch?.();
        const result = await (this.auth as any).graph(url, {
            method, body, headers, beforeDispatch,
        });
        beforeDispatch?.();
        return result;
    }

    private validateGraphEndpoint(url: string): void {
        if (url.startsWith("/")) {
            if (url.startsWith("//")) throw new Error("不可信 Graph 请求路径");
            return;
        }
        let parsed: URL;
        try { parsed = new URL(url); } catch { throw new Error("Graph 分页链接格式不可信"); }
        if (parsed.protocol !== "https:" ||
            parsed.hostname.toLowerCase() !== "graph.microsoft.com" ||
            parsed.username || parsed.password || parsed.port ||
            !parsed.pathname.startsWith("/v1.0/")) {
            throw new Error("Graph 分页链接不属于微软受信任 API，禁止携带授权访问");
        }
    }

    private async getApprootId(): Promise<string> {
        if (this.approotId) return this.approotId;
        const root = await this.graph("GET", "/drive/special/approot?$select=id,name");
        if (!root?.id) throw new Error("无法获取 OneDrive AppRoot 物理节点标识");
        this.approotId = String(root.id);
        return this.approotId;
    }

    private async ensureFolder(
        path: string, operation?: RunTicket, beforeGraphDispatch?: () => void,
    ): Promise<void> {
        if (operation) this.assertCurrent(operation);
        this.ensureRemoteIdentityFresh();
        const target = normalizePath(path);
        if (!target || this.ensuredRemoteFolders.has(target)) return;
        let pending = this.ensuringRemoteFolders.get(target);
        if (!pending) {
            pending = (async () => {
                const endpoint = `/drive/special/approot:/${encodePath(target)}`;
                const lookup = await this.graph("GET", endpoint).catch((error: any) => {
                    if (error?.status === 404) return null;
                    throw error;
                });
                if (operation) this.assertCurrent(operation);
                if (lookup) {
                    if (!lookup.folder) {
                        throw new Error("云端存在同名文件，已安全停止同步，请先在 OneDrive 中确认该项目。");
                    }
                    if (lookup.id) this.remoteFolderIds.set(target, String(lookup.id));
                    this.ensuredRemoteFolders.add(target);
                    return;
                }

                const separator = target.lastIndexOf("/");
                const parentPath = separator >= 0 ? target.slice(0, separator) : "";
                const folderName = separator >= 0 ? target.slice(separator + 1) : target;
                if (parentPath) await this.ensureFolder(parentPath, operation, beforeGraphDispatch);
                if (operation) this.assertCurrent(operation);
                let parentId = "";
                if (parentPath) {
                    parentId = this.remoteFolderIds.get(parentPath) || "";
                    if (!parentId) {
                        const parentMeta = await this.graph("GET", `/drive/special/approot:/${encodePath(parentPath)}?$select=id`).catch(() => null);
                        if (parentMeta?.id) {
                            parentId = String(parentMeta.id);
                            this.remoteFolderIds.set(parentPath, parentId);
                        }
                    }
                    if (!parentId) {
                        throw new Error(`无法获取云端父目录 [${parentPath}] 的 ID，已安全停止创建`);
                    }
                } else {
                    parentId = await this.getApprootId();
                }
                const childrenEndpoint = `/drive/items/${parentId}/children`;
                try {
                    if (operation) this.assertCurrent(operation);
                    const created = await this.graph("POST", childrenEndpoint, {
                        name: folderName,
                        folder: {},
                    }, undefined, () => {
                        if (operation) this.assertCurrent(operation);
                        beforeGraphDispatch?.();
                    });
                    if (operation) this.assertCurrent(operation);
                    if (created?.id) this.remoteFolderIds.set(target, String(created.id));
                } catch (error: any) {

                    const alreadyExists =
                        error?.status === 409 ||
                        String(error?.message || "").toLowerCase().includes("already exists");
                    if (!alreadyExists) throw error;
                    const existing = await this.graph("GET", endpoint);
                    if (operation) this.assertCurrent(operation);
                    if (!existing?.folder) {
                        throw new Error("云端存在同名文件，已安全停止同步，请先在 OneDrive 中确认该项目。");
                    }
                    if (existing?.id) this.remoteFolderIds.set(target, String(existing.id));
                }
                this.ensuredRemoteFolders.add(target);
            })();
            this.ensuringRemoteFolders.set(target, pending);
        }
        try {
            await pending;
            if (operation) this.assertCurrent(operation);
        } finally {
            this.ensuringRemoteFolders.delete(target);
        }
    }

    private remoteIdentityKey = "";

    private ensureRemoteIdentityFresh(): void {
        // Pure protocol tests may instantiate the transport without a settings
        // store; production always supplies one through SmartMDStore.
        if (!(this.store as any)?.settings?.local?.oneDrive) return;
        const key = this.contextKey() + "|" + String(this.settings().deviceId || "");
        if (this.remoteIdentityKey === key) return;
        this.remoteIdentityKey = key;
        this.remoteDeltaCache = null;
        this.approotId = null;
        this.ensuredRemoteFolders.clear();
        this.ensuringRemoteFolders.clear();
        this.remoteFolderIds.clear();
        this.vaultRegistryCache = null;
        this.registryReadIsTrusted = false;
    }

    private remoteDeltaCacheKey(): string {
        const local = (this.store as any).settings.local;
        // The previous cursor did not request modification metadata.
        return "smartmd:onedrive-delta:v3:" +
            encodeURIComponent(String(local.microsoftAccountId || "")) + ":" +
            encodeURIComponent(String(this.settings().deviceId || "")) + ":" +
            encodeURIComponent(String(this.settings().vaultId || "")) +
            (Platform.isAndroidApp ? ":" + encodeURIComponent(String(this.settings().installationProofId || "")) : "");
    }

    private loadRemoteDeltaCache(): { vaultId: string; rootId: string; deltaLink: string; nodes: Record<string, any> } {
        const vaultId = String(this.settings().vaultId || "");
        if (this.remoteDeltaCache?.vaultId === vaultId) return this.remoteDeltaCache;
        try {
            const parsed = JSON.parse(window.localStorage.getItem(this.remoteDeltaCacheKey()) || "null");
            if (parsed?.vaultId === vaultId && parsed.nodes && typeof parsed.nodes === "object") {
                this.remoteDeltaCache = parsed;
                return parsed;
            }
        } catch {}
        this.remoteDeltaCache = { vaultId, rootId: "", deltaLink: "", nodes: {} };
        return this.remoteDeltaCache;
    }

    private saveRemoteDeltaCache(cache: { vaultId: string; rootId: string; deltaLink: string; nodes: Record<string, any> }): void {
        this.remoteDeltaCache = cache;
        try {
            window.localStorage.setItem(this.remoteDeltaCacheKey(), JSON.stringify(cache));
        } catch {}
    }

    private clearRemoteDeltaCache(): void {
        try {
            window.localStorage.removeItem(this.remoteDeltaCacheKey());
        } catch {}
        this.remoteDeltaCache = null;
    }

    private pathFromDeltaNode(node: any, nodes: Record<string, any>, rootId: string): string {
        if (!node?.name) return "";
        const parts = [String(node.name)];
        const visited = new Set<string>([String(node.id || "")]);
        let parentId = String(node.parentId || "");
        while (parentId && parentId !== rootId) {
            if (visited.has(parentId)) return "";
            visited.add(parentId);
            const parent = nodes[parentId];
            if (!parent?.name) return "";
            parts.unshift(String(parent.name));
            parentId = String(parent.parentId || "");
        }
        return parentId === rootId ? normalizePath(parts.join("/")) : "";
    }

    private async remoteFiles(
        forceFull = false, readOnly = false,
        firstCopy?: { vaultId: string; scope?: FirstUploadReadScope; onPage?: (page: number) => void; folderPaths?: Set<string> },
        expectedRootId?: string,
    ): Promise<RemoteFile[]> {
        firstCopy?.scope?.assertCurrent();
        this.ensureRemoteIdentityFresh();
        const vaultId = firstCopy?.vaultId || this.previewVaultId || this.settings().vaultId;
        const basePath = `vaults/${vaultId}/files`;
        const guard = firstCopy?.scope?.assertCurrent;
        // Listing a sync scope must never create it. A missing remote folder
        // is an identity/preflight failure, not permission to POST a new folder.
        // Folder creation remains in separately authorized write paths.
        const loaded = readOnly
            ? { vaultId, rootId: "", deltaLink: "", nodes: {} as Record<string, any> }
            : this.loadRemoteDeltaCache();
        const cache = { ...loaded, nodes: { ...(loaded.nodes || {}) } };
        if (forceFull || !cache.deltaLink || !cache.rootId) {
            const root = await this.graph("GET", `/drive/special/approot:/${encodePath(basePath)}`, undefined, undefined, guard).catch((error: any) => {
                const confirmedNew = this.settings().firstCommitJournal as FirstCommitJournal | undefined;
                if (readOnly && error?.status === 404
                    && (this.pendingFirstChoice?.mode === "new"
                        || (confirmedNew?.stage === "REGISTERED"
                            && confirmedNew.identity.cloudVaultId === (firstCopy?.vaultId || this.previewVaultId)))) return null;
                throw error;
            });
            if (readOnly && !root) return []; // New vault not created: a legitimate empty read-only target.
            if (!root?.folder || !root?.id) throw new Error("云端知识库目录身份未知，无法完成只读检查。");
            if (expectedRootId && String(root.id) !== expectedRootId) {
                throw new Error(localize("OneDrive 同步目录已变化，已暂缓删除。"));
            }
            cache.rootId = String(root.id);
            cache.deltaLink = "";
            cache.nodes = {};
        }
        let next: string | null = cache.deltaLink || `/drive/items/${cache.rootId}/delta?$select=id,name,size,eTag,cTag,parentReference,file,folder,deleted,fileSystemInfo,lastModifiedDateTime`;
        const visitedPages = new Set<string>();
        try {
            while (next) {
                this.validateGraphEndpoint(next);
                if (visitedPages.has(next) || visitedPages.size >= 500) {
                    throw new Error("云端增量清单分页循环或超出安全页数，已保留旧基线并停止");
                }
                visitedPages.add(next);
                const page: any = await this.graph("GET", next, undefined, undefined, guard);
                if (!Array.isArray(page.value)) throw new Error("云端清单不完整，已暂停删除并等待重新比对");
                for (const item of page.value || []) {
                    const id = String(item.id || "");
                    if (!id || id === cache.rootId) continue;
                    if (item.deleted) {
                        delete cache.nodes[id];
                        continue;
                    }
                    const previous = cache.nodes[id] || {};
                    cache.nodes[id] = {
                        ...previous,
                        id,
                        name: String(item.name || previous.name || ""),
                        parentId: String(item.parentReference?.id || previous.parentId || ""),
                        folder: Boolean(item.folder),
                        size: Number(item.size ?? previous.size ?? 0),
                        etag: String(item.eTag || item.cTag || previous.etag || ""),
                        quickXorHash: String(item.file?.hashes?.quickXorHash || previous.quickXorHash || ""),
                        sha1Hash: String(item.file?.hashes?.sha1Hash || previous.sha1Hash || ""),
                        mtime: this.remoteModificationTime(item),
                    };
                }
                const nextLink: unknown = page["@odata.nextLink"];
                const deltaLink: unknown = page["@odata.deltaLink"];
                if (nextLink !== undefined && nextLink !== null && nextLink !== "") {
                    if (typeof nextLink !== "string") throw new Error("云端分页地址类型错误，已停止同步");
                    this.validateGraphEndpoint(nextLink);
                }
                if (deltaLink !== undefined && deltaLink !== null && deltaLink !== "") {
                    if (typeof deltaLink !== "string") throw new Error("云端增量游标格式错误，已停止同步");
                    this.validateGraphEndpoint(deltaLink);
                }
                if (nextLink && deltaLink) throw new Error("云端同时返回分页与结束标识，清单不可信");
                next = typeof nextLink === "string" && nextLink ? nextLink : null;
                if (typeof deltaLink === "string" && deltaLink) cache.deltaLink = deltaLink;
                if (!next && !page["@odata.deltaLink"]) {
                    throw new Error("云端清单没有完整结束证明，已保留上一轮关系并暂停删除");
                }
                firstCopy?.onPage?.(visitedPages.size);
            }
        } catch (error: any) {
            if (!readOnly && !forceFull && (error?.status === 400 || error?.status === 404 || error?.status === 410)) {
                // Do not erase the last trusted persisted snapshot before the
                // full rescan succeeds. A second 429/timeout must preserve it.
                return this.remoteFiles(true);
            }
            throw error;
        }
        guard?.();
        {
            const seenPaths = new Map<string, { id: string; etag: string; path: string }>();
            this.remoteIdentityConflict = null;
            for (const node of Object.values(cache.nodes) as any[]) {
                if (node.folder) {
                    if (firstCopy?.folderPaths) {
                        const relative = this.pathFromDeltaNode(node, cache.nodes, cache.rootId);
                        if (!relative) throw new Error("云端文件夹路径无法核对，请重新同步");
                        firstCopy.folderPaths.add(relative);
                    }
                    continue;
                }
                const relative = this.pathFromDeltaNode(node, cache.nodes, cache.rootId);
                const windowsKey = relative.normalize("NFC").toLocaleLowerCase("en-US");
                if (!relative || !node.id || !node.etag) {
                    throw new Error("云端文件清单包含未知路径或缺少版本标识，已停止预览。");
                }
                const item = { id: String(node.id), etag: String(node.etag), path: relative };
                const previous = seenPaths.get(windowsKey);
                if (previous) {
                    this.remoteIdentityConflict = {
                        accountId: String((this.store as any).settings.local.microsoftAccountId || ""),
                        vaultId: basePath, path: relative, items: [previous, item],
                    };
                    throw new Error("云端重复身份：" + redactOneDriveText(relative) + "；Item ID：" +
                        redactOneDriveText(previous.id) + " / " + redactOneDriveText(item.id) +
                        "。已停止同步，保留双方；请重新只读核对云端身份。");
                }
                seenPaths.set(windowsKey, item);
            }
        }
        if (!readOnly) this.saveRemoteDeltaCache(cache);
        return Object.values(cache.nodes)
            .filter((node: any) => !node.folder)
            .map((node: any) => ({
                id: node.id,
                path: this.pathFromDeltaNode(node, cache.nodes, cache.rootId),
                name: node.name,
                size: node.size || 0,
                etag: node.etag || "",
                quickXorHash: node.quickXorHash || undefined,
                sha1Hash: node.sha1Hash || undefined,
                mtime: node.mtime,
            }))
            .filter((item: RemoteFile) => Boolean(item.path));
    }

    private async upload(
        path: string,
        data: ArrayBuffer,
        appRelative = false,
        expectedRemote?: RemoteFile | null,
        operation?: RunTicket,
        journalMotherPath?: string,
        beforeGraphDispatch?: () => void,
        sourceMtime?: number,
    ): Promise<any> {
        // Capture the Microsoft session before any directory/session/journal
        // awaits; a new account cannot inherit this upload URL.
        const sessionGuard: (() => void) | undefined =
            (this.auth as any).createDispatchGuard?.();
        sessionGuard?.();
        if (operation) this.assertCurrent(operation);
        if (appRelative && expectedRemote === undefined) {
            expectedRemote = await this.readRemoteFileMetadata(path, true);
            if (operation) this.assertCurrent(operation);
        }
        const remotePath = appRelative ? path : `${this.basePath()}/${path}`;
        const slash = remotePath.lastIndexOf("/");
        if (slash >= 0) await this.ensureFolder(remotePath.slice(0, slash), operation, beforeGraphDispatch);
        // Directory inspection can be asynchronous: recheck before any PUT/POST.
        if (operation) this.assertCurrent(operation);
        const pathTarget = `/drive/special/approot:/${encodePath(remotePath)}`;
        const itemTarget = expectedRemote ? `/drive/items/${expectedRemote.id}` : "";
        const versionHeaders: Record<string, string> = expectedRemote === undefined
            ? {}
            : expectedRemote
                ? { "If-Match": expectedRemote.etag.startsWith('"') || expectedRemote.etag.startsWith('W/"') ? expectedRemote.etag : `"${expectedRemote.etag}"` }
                : {}; // 个人版 OneDrive 网关对 PUT ...:/content 不支持 If-None-Match: *（会报 400 invalidRequest），新文件置空以兼容个人版网关
        if (expectedRemote && !expectedRemote.etag) {
            throw new Error("OneDrive 旧文件版本缺少 ETag，禁止无条件覆盖");
        }
        // A small PUT without an exclusive server-side condition overwrites
        // another device's new file. Use Graph upload-session conflict=fail
        // when creating even small files, including app-relative metadata.
        // Empty new files have no byte range to commit in an upload session.
        // Graph's URL conflictBehavior=fail keeps zero-byte creation exclusive;
        // verified against the dedicated personal account (run308:201 then409).
        if ((data.byteLength < 4_000_000 && expectedRemote) || (data.byteLength === 0 && expectedRemote === null)) {
            const endpoint = expectedRemote
                ? `${itemTarget}/content`
                : `${pathTarget}:/content?@microsoft.graph.conflictBehavior=fail`;
            let response: Response;
            try {
                response = await this.fetchGraph(endpoint, {
                    method: "PUT",
                    body: data,
                    headers: { "Content-Type": "application/octet-stream", ...versionHeaders },
                }, operation, beforeGraphDispatch);
            } catch (error) {
                if (operation && !this.isCurrent(operation)) throw error;
                const uncertain: any = new Error("条件PUT请求结果未知，不得盲目重试，须核ItemID/ETag和实际远端字节");
                uncertain.isUnknownCommit = true;
                uncertain.cause = error;
                throw uncertain;
            }
            if (!response.ok) {
                const uncertain = [408, 429, 500, 502, 503, 504].includes(response.status);
                const message = uncertain
                    ? `条件PUT已发出但结果未知 (HTTP ${response.status})，需只读对账`
                    : await this.fetchError(response, `上传失败 (${response.status})`);
                const error: any = new Error(message);
                error.status = response.status;
                error.retryAfterMs = graphRetryAfterMs(response.headers.get("retry-after"));
                error.isUnknownCommit = uncertain;
                throw error;
            }
            try {
                const payload = await response.json();
                if (!payload?.id || !(payload?.eTag || payload?.etag)) {
                    throw new Error("条件PUT成功响应缺少可核ItemID/ETag");
                }
                return payload;
            } catch {
                const unknown: any = new Error("已发条件PUT后响应解码或版本字段无效；原操作必须先只读对账");
                unknown.isUnknownCommit = true;
                throw unknown;
            }
        }

        const sessionEndpoint = expectedRemote
            ? `${itemTarget}/createUploadSession`
            : `${pathTarget}:/createUploadSession`;
        if (operation) this.assertCurrent(operation);
        // The URL already identifies the target name or Item ID. Sending
        // item.name again causes HTTP 400 invalidRequest on personal OneDrive;
        // a conflict=fail-only item was verified against the real test tenant.
        const session = await this.graph("POST", sessionEndpoint, {
            item: {
                "@microsoft.graph.conflictBehavior": "fail",
                ...(sourceMtime ? { fileSystemInfo: { lastModifiedDateTime: new Date(sourceMtime).toISOString() } } : {}),
            },
        }, versionHeaders, () => {
            if (operation) this.assertCurrent(operation);
            beforeGraphDispatch?.();
        });
        if (operation) this.assertCurrent(operation);
        if (!session?.uploadUrl) throw new Error("OneDrive 无法建立带防覆盖约束的上传会话，已安全停止");
        const chunkSize = 327_680 * 20;
        let finalItem: any = null;
        for (let start = 0; start < data.byteLength; start += chunkSize) {
            const end = Math.min(start + chunkSize, data.byteLength);
            const chunk = data.slice(start, end);
            // A PUT may have committed even when its 429/503/timeout response
            // is lost. Never blindly resend the same range/session without
            // first reconciling server nextExpectedRanges and exact file SHA.
            if (operation) this.assertCurrent(operation);
            if (journalMotherPath && path !== journalMotherPath) {
                await this.setUploadPhase(journalMotherPath, "STAGE_PUT_ISSUED", {
                    // Never store a credential-bearing uploadUrl in data.json.
                    sessionFingerprint: await this.sha256Bytes(
                        new TextEncoder().encode(session.uploadUrl).buffer,
                    ),
                });
            }
            // All staging SHA/journal awaits are now complete. Gate the exact
            // upload-session PUT like Graph POST/DELETE; no await before fetch.
            // A revoked ticket never turns a not-yet-sent request into UNKNOWN.
            sessionGuard?.();
            if (operation) this.assertCurrent(operation);
            beforeGraphDispatch?.();
            let response: Response;
            try {
                response = await fetch(session.uploadUrl, {
                    method: "PUT",
                    body: chunk,
                    headers: {
                        "Content-Length": String(chunk.byteLength),
                        "Content-Range": `bytes ${start}-${end - 1}/${data.byteLength}`,
                        "Content-Type": "application/octet-stream",
                    },
                });
            } catch (error) {
                const failure: any = new Error("上传分块已发出但服务器结果未知；保留暂存与旧SHA，必须只读对账");
                failure.isUnknownCommit = true;
                failure.cause = error;
                throw failure;
            }
            if (!response.ok) {
                const unknown = [408, 429, 500, 502, 503, 504].includes(response.status);
                const failure: any = new Error(unknown
                    ? "上传分块已发出但服务端结果未知（HTTP " + response.status + "）；停止重发，要求只读对账"
                    : await this.fetchError(response, `上传分块失败 (${response.status})`));
                failure.status = response.status;
                failure.retryAfterMs = graphRetryAfterMs(response.headers.get("retry-after"));
                failure.isUnknownCommit = unknown;
                throw failure;
            }
            let payload: any;
            try {
                payload = await response.json();
            } catch {
                const unknown: any = new Error("已发上传分块后成功回包解析失败，必须停止剩余计划并按原会话对账");
                unknown.isUnknownCommit = true;
                throw unknown;
            }
            if (response.status === 200 || response.status === 201) {
                if (!payload?.id || !(payload?.eTag || payload?.etag)) {
                    const unknown: any = new Error("已发上传分块后成功回包缺ItemID/ETag，提交状态未知");
                    unknown.isUnknownCommit = true;
                    throw unknown;
                }
                finalItem = payload;
            }
        }
        if (!finalItem?.id || !(finalItem.eTag || finalItem.etag)) {
            const unknown: any = new Error("已发分块但未获得可核实的最终ItemID/ETag，必须只读对账");
            unknown.isUnknownCommit = true;
            throw unknown;
        }
        return finalItem;
    }

    private async setUploadedModificationTime(
        uploaded: any, mtime: number, data: ArrayBuffer, operation?: RunTicket, beforeGraphDispatch?: () => void,
        acceptAcknowledgedTimestamp = false,
    ): Promise<any> {
        const version = uploaded?.eTag || uploaded?.etag;
        if (!uploaded?.id || !version || !Number.isFinite(mtime) || mtime <= 0) {
            throw new Error(localize("无法核对文件的修改时间，请重新同步。"));
        }
        if (this.remoteModificationTime(uploaded) === mtime &&
            uploaded.fileSystemInfo?.lastModifiedDateTime) return uploaded;
        const itemPath = "/drive/items/" + encodeURIComponent(uploaded.id);
        const checkCurrent = () => {
            if (operation) this.assertCurrent(operation);
            beforeGraphDispatch?.();
        };
        try {
            const updated = await this.graph("PATCH", itemPath,
                { fileSystemInfo: { lastModifiedDateTime: new Date(mtime).toISOString() } },
                { "If-Match": version }, checkCurrent);
            if (updated?.id === uploaded.id && (updated.eTag || updated.etag) &&
                (acceptAcknowledgedTimestamp ||
                    (updated.fileSystemInfo?.lastModifiedDateTime && this.remoteModificationTime(updated) === mtime))) {
                return { ...uploaded, ...updated, eTag: updated.eTag || updated.etag };
            }
        } catch (error: any) {
            checkCurrent();
            if (error?.isFatalAuth || error?.retryAfterMs || [401, 403, 429, 507].includes(Number(error?.status))) throw error;
        }
        try {
            // OneDrive can save the time while omitting it from the PATCH
            // response. Check the exact item and bytes before accepting it.
            const select = "?$select=id,name,eTag,size,fileSystemInfo";
            const current = await this.graph("GET", itemPath + select, undefined, undefined, checkCurrent);
            const etag = current?.eTag || current?.etag;
            if (current?.id === uploaded.id && etag && current.size === data.byteLength &&
                this.remoteModificationTime(current) === mtime &&
                sameBytes(await this.remoteBytes({ id: current.id, etag, path: String(current.name || ""),
                    name: String(current.name || ""), size: current.size }), data)) {
                const after = await this.graph("GET", itemPath + select, undefined, undefined, checkCurrent);
                if (after?.id === current.id && (after.eTag || after.etag) === etag &&
                    this.remoteModificationTime(after) === mtime) return { ...uploaded, ...current, eTag: etag };
            }
        } catch (error: any) {
            checkCurrent();
            if (error?.isFatalAuth || error?.retryAfterMs || [401, 403, 429, 507].includes(Number(error?.status))) throw error;
        }
        throw Object.assign(new Error("OneDrive 未确认文件原修改时间"), { isUnknownCommit: true });
    }

    private async uploadStaged(
        operation: RunTicket,
        path: string,
        data: ArrayBuffer,
        appRelative = false,
        expectedRemote?: RemoteFile | null,
        beforeGraphDispatch?: () => void,
        onFinalAcknowledged?: (item: any) => Promise<void>,
    ): Promise<any> {
        // A previously issued ticket must not enter cloud writes while the
        // durable ledger-release transaction is unresolved.
        if (this.uploadLedgerReconciliationInFlight || this.uploadLedgerSettlementInFlight) {
            throw new Error("历史上传日志仍在持久结算，不得开始新暂存或母文件写入");
        }
        this.assertCurrent(operation);
        const expectedLocal = appRelative ? undefined : this.localFiles().find((item) => item.path === path);
        const safeOpId = operation.id.replace(/[\\/:*?"<>|.]/g, "_");
        const stagingPath = `_suite_staging/${safeOpId}/${path.replace(/[\\/:*?"<>|]/g, "_")}`;
        if (!appRelative) {
            const ledger = this.uploadLedger();
            if (ledger[path]) throw new Error("原上传操作及暂存副本尚未只读对账，不得新建第二个暂存会话");
            if (Object.values(ledger).some(record => record?.stagingPath === stagingPath)) {
                throw new Error("另一个文件仍在等待上传核对，请稍后重试此文件。");
            }
            const local = (this.store as any).settings.local;
            ledger[path] = {
                operationId: operation.id,
                accountId: String(local.microsoftAccountId || ""),
                vaultId: this.settings().vaultId,
                deviceId: this.settings().deviceId,
                motherPath: path, direction: "upload",
                sourceSha256: await this.sha256Bytes(data),
                sourceBytes: data.byteLength,
                sourceMtime: expectedLocal?.mtime,
                sourceBaseline: JSON.stringify(this.settings().fileStates?.[path] || null),
                expectedRemoteId: expectedRemote?.id || "",
                expectedRemoteEtag: expectedRemote?.etag || "",
                stagingPath, phase: "STAGE_SESSION_ISSUED",
                updatedAt: Date.now(),
            };
            // This must durably complete BEFORE the first session POST/PUT.
            await (this.store as any).save();
        }
        beforeGraphDispatch?.();
        const staged = await this.upload(stagingPath, data, appRelative, null, operation,
            appRelative ? undefined : path, beforeGraphDispatch);
        if (!appRelative) await this.setUploadPhase(path, "STAGE_ACKNOWLEDGED", {
            stageId: String(staged?.id || ""),
            stageEtag: String(staged?.eTag || staged?.etag || ""),
        });
        this.assertCurrent(operation);
        if (!appRelative) await this.assertFresh(operation, path, expectedLocal, expectedRemote || undefined);
        this.assertCurrent(operation);
        if (!appRelative && expectedLocal) {
            const current = this.app.vault.getAbstractFileByPath(path);
            if (!(current instanceof TFile) || !sameBytes(await this.app.vault.readBinary(current), data)) {
                throw new Error(`源文件在最终提交前再次变化，已保留当前内容与云端暂存：${path}`);
            }
        }
        this.authority.beginFinal(operation);
        let verificationStep: "timestamp" | "readback" | undefined;
        try {
            if (!appRelative) await this.setUploadPhase(path, "FINAL_PUT_ISSUED");
            let finalUploaded = await this.upload(
                path, data, appRelative, expectedRemote, operation,
                appRelative ? undefined : path, beforeGraphDispatch, expectedLocal?.mtime,
            );
            if (!appRelative && finalUploaded?.id && (finalUploaded.eTag || finalUploaded.etag)) {
                await this.setUploadPhase(path, "FINAL_PUT_ISSUED", {
                    contentAcknowledgedId: finalUploaded.id,
                    contentAcknowledgedEtag: finalUploaded.eTag || finalUploaded.etag,
                });
            }
            if (expectedLocal) {
                verificationStep = "timestamp";
                try {
                    const contentEtag = finalUploaded?.eTag || finalUploaded?.etag;
                    finalUploaded = await this.setUploadedModificationTime(
                        finalUploaded, expectedLocal.mtime, data, operation, beforeGraphDispatch, !appRelative,
                    );
                    const timestampEtag = finalUploaded?.eTag || finalUploaded?.etag;
                    if (!appRelative && timestampEtag !== contentEtag) {
                        await this.setUploadPhase(path, "FINAL_PUT_ISSUED", {
                            contentAcknowledgedId: finalUploaded.id,
                            contentAcknowledgedEtag: timestampEtag,
                        });
                    }
                } catch (error: any) {
                    if (appRelative || !error?.isUnknownCommit) throw error;
                    // OneDrive may store a different timestamp. Confirm the exact
                    // acknowledged file and bytes before accepting that metadata.
                }
            }
            if (!appRelative) {
                // IMPORTANT: the normal upload path must durably save the
                // verified mother baseline BEFORE dispatching stage DELETE.
                // The resumed UNKNOWN settlement path already does this; the
                // ordinary transfer must offer the same crash recovery.
                const finalId = String(finalUploaded?.id || "");
                const finalEtag = String(finalUploaded?.eTag || finalUploaded?.etag || "");
                if (!finalId || !finalEtag) {
                    throw new Error("普通上传目标已发写入但没有可信ItemID/ETag，保留原操作等待只读对账");
                }
                const localFile = this.app.vault.getAbstractFileByPath(path);
                if (!(localFile instanceof TFile)) {
                    throw new Error("正式写入后本机原稿改变，停止暂存清理并保护已上传原文");
                }
                verificationStep = "readback";
                const remote = await this.readRemoteFileMetadata(path);
                if (!remote || remote.id !== finalId || remote.etag !== finalEtag ||
                    remote.size !== data.byteLength || (expectedLocal && remote.mtime === undefined) ||
                    !sameBytes(await this.remoteBytes(remote), data)) {
                    throw new Error("普通上传母文件尚未完成真实版本和SHA读回，禁止删除最后的暂存副本");
                }
                const verified = await this.readRemoteFileMetadata(path);
                if (!verified || verified.id !== remote.id || verified.etag !== remote.etag ||
                    verified.mtime !== remote.mtime) {
                    throw new Error("核对上传结果时云端版本再次变化");
                }
                this.assertCurrent(operation);
                const settings = this.settings();
                const record = this.uploadLedger()[path];
                if (!record || record.phase !== "FINAL_PUT_ISSUED" ||
                    record.operationId !== operation.id ||
                    record.motherPath !== path ||
                    record.sourceSha256 !== await this.sha256Bytes(data)) {
                    throw new Error("上传原操作身份或阶段漂移，不能登记母文件基线");
                }
                if (this.app.vault.getAbstractFileByPath(path) !== localFile ||
                    !sameBytes(await this.app.vault.readBinary(localFile), data)) {
                    throw new Error("核对上传结果时本机文件再次变化");
                }
                this.assertCurrent(operation);
                const priorStates = settings.fileStates || {};
                const priorRecord = { ...record };
                settings.fileStates = {
                    ...priorStates, [path]: {
                        localMtime: record.sourceMtime,
                        localSize: record.sourceBytes,
                        localHash: record.sourceSha256,
                        remoteId: finalId, remoteEtag: finalEtag,
                    },
                };
                Object.assign(record, {
                    phase: "FINAL_ACKNOWLEDGED",
                    finalId, finalEtag, updatedAt: Date.now(),
                });
                try {
                    // Atomic from the perspective of the persisted plugin
                    // settings: the baseline and the blocking ledger are
                    // committed in one save, ahead of conditional DELETE.
                    await (this.store as any).save();
                } catch (error) {
                    if (this.settings() === settings &&
                        this.settings().pendingUploadOperations?.[path] === record) {
                        settings.fileStates = priorStates;
                        this.uploadLedger()[path] = priorRecord;
                    }
                    throw error;
                }
            }
            if (appRelative) {
                // Persist an acknowledged mother before any readback or cleanup can fail.
                await onFinalAcknowledged?.(finalUploaded);
                beforeGraphDispatch?.();
                const meta = await this.readRemoteFileMetadata(path, true);
                const finalEtag = String(finalUploaded?.eTag || finalUploaded?.etag || "");
                if (!meta?.id || !meta.etag || meta.id !== finalUploaded?.id || meta.etag !== finalEtag ||
                    !sameBytes(await this.remoteBytes(meta), data)) {
                    throw new Error("共享母文件未读回匹配版本及正文，保留暂存副本");
                }
                beforeGraphDispatch?.();
            }
            if (!staged?.id || !(staged?.eTag || staged?.etag)) {
                throw new Error("正式文件已提交，但云端暂存副本无可信删除版本，需人工只读对账");
            }
            if (!appRelative) await this.setUploadPhase(path, "CLEANUP_ISSUED");
            await this.deleteRemote({
                ...staged,
                etag: staged.eTag || staged.etag,
            }, operation, beforeGraphDispatch);
            if (!appRelative) await this.setUploadPhase(path, "CLEANUP_VERIFIED");
            return finalUploaded;
        } catch (error: any) {
            const phase = !appRelative ? this.uploadLedger()[path]?.phase : "";
            // Once final content has been acknowledged, a failed cleanup
            // response is an UNKNOWN cleanup, not a retryable next-file error.
            const durableFinal = phase === "FINAL_ACKNOWLEDGED" ||
                phase === "CLEANUP_ISSUED" || phase === "CLEANUP_VERIFIED";
            if (durableFinal) {
                // DELETE may have succeeded while its acknowledgement was
                // lost. Preserve the trusted mother baseline and monotonic
                // cleanup phase for a later single-file GET404 recovery.
                error.isUnknownCommit = true;
            } else if (!appRelative) {
                try {
                    await this.setUploadPhase(path, "RECONCILE_REQUIRED", { verificationStep });
                } catch {
                    // Never suppress the original failure or assume a failed
                    // journal update proves that the remote request failed.
                }
            }
            if (error?.status === 409 || error?.status === 412 || /\b(409|412)\b/.test(String(error?.message || error))) {
                const wrapped: any = new Error(`OneDrive 目标版本已变化，条件提交被拒绝并保留双方内容：${path}`);
                wrapped.status = Number(error?.status || 0) || undefined;
                // An already-persisted mother and CLEANUP_ISSUED must remain
                // classified as unresolved completion across error wrapping.
                wrapped.isUnknownCommit = Boolean(error?.isUnknownCommit || durableFinal);
                throw wrapped;
            }
            throw error;
        }
    }

    private graphUrl(pathOrUrl: string): string {
        this.validateGraphEndpoint(pathOrUrl);
        return pathOrUrl.startsWith("https:")
            ? pathOrUrl : "https://graph.microsoft.com/v1.0" + pathOrUrl;
    }

    private async fetchGraph(
        pathOrUrl: string, options: any, operation?: RunTicket,
        beforeDispatch?: () => void, checkPhysicalPath?: () => Promise<void>,
    ): Promise<Response> {
        const sessionGuard: (() => void) | undefined =
            (this.auth as any).createDispatchGuard?.();
        const send = async (token: string) => {
            // Both the OneDrive operation and Microsoft account session must
            // remain valid after awaiting a token or 401 refresh.
            sessionGuard?.();
            if (operation) this.assertCurrent(operation);
            if (checkPhysicalPath) {
                await checkPhysicalPath();
                sessionGuard?.();
                if (operation) this.assertCurrent(operation);
            }
            // Synchronous final check after each credential await. No await
            // between this check and the physical network dispatch.
            beforeDispatch?.();
            return fetch(this.graphUrl(pathOrUrl), {
                ...options,
                headers: { ...(options.headers || {}), Authorization: `Bearer ${token}` },
            });
        };
        let response = await send(await (this.auth as any).getAccessToken());
        if (response.status === 401) response = await send(await (this.auth as any).getAccessToken(true));
        return response;
    }

    private async fetchError(response: Response, fallback: string): Promise<string> {
        try {
            const payload: any = await response.clone().json();
            const inner = payload?.error?.innerError;
            const deepMsg = inner?.message || payload?.error?.message || payload?.error_description || fallback;
            const reqId = inner?.["request-id"] || payload?.error?.["request-id"] || "";
            const code = payload?.error?.code || "";
            const status = response.status;
            let result = deepMsg;
            if (code && !deepMsg.includes(code)) result = `[${code}] ${result}`;
            if (status && !result.includes(`HTTP ${status}`)) result = `HTTP ${status}: ${result}`;
            if (reqId) result = `${result} (request-id: ${reqId})`;
            return result;
        } catch {
            try {
                return (await response.clone().text()) || fallback;
            } catch {
                return fallback;
            }
        }
    }

    private async remoteBytes(remote: RemoteFile, scope?: FirstUploadReadScope): Promise<ArrayBuffer> {
        const response = await this.downloadResponse(`/drive/items/${remote.id}`, scope);
        scope?.assertCurrent();
        if (response.status < 200 || response.status >= 300) {
            const error: any = new Error(`下载失败：${remote.path} (HTTP ${response.status})`);
            error.status = response.status;
            error.retryAfterMs = graphRetryAfterMs(response.headers?.["retry-after"]);
            error.isMissingFile = response.status === 404;
            throw error;
        }
        return response.arrayBuffer;
    }

    private async downloadResponse(metadataPath: string, scope?: FirstUploadReadScope): Promise<any> {
        scope?.assertCurrent();
        let normalizedEndpoint = metadataPath.trim();
        if (normalizedEndpoint.includes("approot:") && normalizedEndpoint.endsWith(":")) {
            normalizedEndpoint = normalizedEndpoint.slice(0, -1);
        }
        const separator = normalizedEndpoint.includes("?") ? "&" : "?";
        const metadata = await (this.auth as any).requestGraph(
            `${normalizedEndpoint}${separator}$select=@microsoft.graph.downloadUrl`,
            scope ? { beforeDispatch: scope.assertCurrent } : undefined,
        );
        scope?.assertCurrent();
        if (metadata.status < 200 || metadata.status >= 300) return metadata;
        const downloadUrl = metadata.json?.["@microsoft.graph.downloadUrl"];
        if (!downloadUrl) throw new Error("OneDrive 未提供有效的文件下载地址");
        try {
            const response = await fetch(downloadUrl, { cache: "no-store", signal: scope?.signal });
            if (!response.ok && (response.status === 429 || response.headers.has("retry-after"))) {
                return { status: response.status, headers: { "retry-after": response.headers.get("retry-after") } };
            }
            if (!response.ok) throw new Error(`download failed (${response.status})`);
            const meter = scope && response.body && typeof TransformStream === "function"
                ? new TransformStream<Uint8Array, Uint8Array>({
                    transform(chunk, output) {
                        scope.noteActivity();
                        output.enqueue(chunk);
                    },
                })
                : null;
            const arrayBuffer = meter && response.body
                ? await new Response(response.body.pipeThrough(meter)).arrayBuffer()
                : await response.arrayBuffer();
            scope?.assertCurrent();
            return {
                status: response.status,
                arrayBuffer,
                text: new TextDecoder().decode(arrayBuffer),
                json: null,
            };
        } catch {
            scope?.assertCurrent();
            return requestUrl({
                url: downloadUrl,
                method: "GET",
                headers: { "Cache-Control": "no-cache" },
                throw: false,
            });
        }
    }

    private async removeVerifiedStaging(path: string): Promise<void> {
        if (!path.startsWith(".suite-staging/onedrive/") ||
            path.includes("..") || !path.includes("/")) {
            throw new Error("暂存路径身份不可信，拒绝删除其他用户文件");
        }
        const indexed = this.app.vault.getAbstractFileByPath(path);
        const adapter: any = this.app.vault.adapter;
        if (indexed instanceof TFile) {
            await this.app.fileManager.trashFile(indexed);
        } else if (await adapter.exists(path)) {
            if (typeof adapter.remove !== "function") {
                throw new Error("隐藏暂存文件存在但缺乏物理清理能力，禁止报告成功");
            }
            await adapter.remove(path);
        }
        if (await adapter.exists(path)) throw new Error("暂存磁盘文件未能彻底清理");
    }

    private async download(
        operation: RunTicket,
        remote: RemoteFile,
        path: string,
        expectedLocal: LocalFile | undefined,
        firstAndroidReceive = false,
        expectedLocalHash?: string,
        expectedRemoteHash?: string,
    ): Promise<string> {
        this.assertCurrent(operation);
        const vault = this.app.vault;
        const adapter = vault.adapter;
        const sharedAndroid = Platform.isAndroidApp && !getAndroidAppStorageRoot(adapter);
        const before = expectedLocal ? await adapter.readBinary(path) : null;
        if (expectedLocalHash && (!before || await this.sha256Bytes(before) !== expectedLocalHash)) {
            throw new Error("本机笔记已在最终预览后修改，保留原件");
        }
        this.assertCurrent(operation);
        const data = await this.remoteBytes(remote);
        if (expectedRemoteHash && await this.sha256Bytes(data) !== expectedRemoteHash) {
            throw new Error("云端笔记的正文已在最终预览后改变，保留本机原件");
        }
        // A late response cannot create a local staging file after revocation.
        this.assertCurrent(operation);
        const safePath = Platform.isAndroidApp ? "payload.bin" : path.replace(/[\\/:*?"<>|]/g, "_");
        const stagingPath = normalizePath(".suite-staging/onedrive/" +
            (Platform.isAndroidApp ? crypto.randomUUID() : operation.id) + "/" + safePath);
        const localHash = await this.sha256Bytes(data);
        if (firstAndroidReceive) {
            const settings = this.settings();
            const local = (this.store as SmartMDStore).settings.local;
            if (settings.androidFirstDownload || !remote.id || !remote.etag || data.byteLength !== remote.size) {
                throw new Error("安卓首次接收存在未决操作或下载版本不可信");
            }
            this.assertCurrent(operation);
            settings.androidFirstDownload = {
                path, stagingPath, accountId: local.microsoftAccountId, vaultId: settings.vaultId,
                deviceId: settings.deviceId, proofId: settings.installationProofId,
                remoteId: remote.id, remoteEtag: remote.etag, size: data.byteLength, sha256: localHash,
            };
            await this.store.save();
            this.assertCurrent(operation);
        }
        await this.ensureLocalFolder(stagingPath);
        if (Platform.isAndroidApp && await this.app.vault.adapter.exists(stagingPath)) {
            throw new Error("安卓端下载暂存位置已占用");
        }
        this.assertCurrent(operation);
        try {
            const writeOptions = remote.mtime ? { mtime: remote.mtime } : undefined;
            await this.app.vault.createBinary(stagingPath, data, writeOptions);
            if (!sameBytes(await this.readPhysicalBinary(stagingPath), data)) {
                throw new Error("OneDrive 下载暂存内容读回失败，禁止覆盖目标");
            }
            await this.assertFresh(operation, path, expectedLocal, remote);
            this.assertCurrent(operation);
            this.authority.beginFinal(operation);
            await this.ensureLocalFolder(path);
            const current = this.app.vault.getAbstractFileByPath(path);
            if (current instanceof TFile) {
                if (!expectedLocal || current !== expectedLocal.file || !before ||
                    !sameBytes(await adapter.readBinary(path), before) ||
                    vault.getAbstractFileByPath(path) !== current || vault.adapter !== adapter) {
                    throw new Error(localize("本机文件在同步时变化，请重新同步。"));
                }
                this.assertCurrent(operation);
                try {
                    if (Platform.isAndroidApp && current.extension.toLowerCase() === "md") {
                        await replaceAndroidMarkdownIfUnchanged(
                            this.app.vault, current, path, before, data, () => this.assertCurrent(operation), writeOptions,
                        );
                    } else if (Platform.isAndroidApp && !sharedAndroid) {
                        await replaceAndroidBinaryIfUnchanged(
                            this.app.vault, current, path, stagingPath, before, data, () => this.assertCurrent(operation),
                        );
                    } else {
                        await this.app.vault.modifyBinary(current, data, writeOptions);
                    }
                } catch {
                    throw new Error("Obsidian 文件更新失败，请重新同步核对");
                }
            } else {
                if (expectedLocal || await adapter.exists(path) ||
                    vault.getAbstractFileByPath(path) || vault.adapter !== adapter) {
                    throw new Error("未被索引的原文件已存在，拒绝覆盖");
                }
                this.assertCurrent(operation);
                if (Platform.isAndroidApp && !sharedAndroid) {
                    await placeAndroidStagedFile(this.app.vault, stagingPath, path, () => this.assertCurrent(operation));
                } else {
                    await this.app.vault.createBinary(path, data, writeOptions);
                }
            }
            const actual = await adapter.readBinary(path);
            if (!sameBytes(actual, data)) {
                throw new Error("知识库落地读回校验失败，请重新同步核对");
            }
            return localHash;
        } finally {
            // Covers successful downloads, precommit errors and partial
            // temporary-file creation. Index visibility is not proof of
            // on-disk deletion: adapter.exists is checked afterward.
            try {
                await this.removeVerifiedStaging(stagingPath);
            } catch (cleanupError) {
                // The caller persists this error before revoking the one-file
                // grant. Keep the exact operation and owned relative staging
                // path for manual reconciliation; never report a clean abort
                // or silently reuse the unverified temporary file.
                throw new Error(
                    "OneDrive 下载暂存清理失败，须核对未决操作 " + operation.id +
                    " 的暂存路径 " + stagingPath + "；" +
                    safeOneDriveError(cleanupError),
                );
            }
        }
    }

    private async readPhysicalBinary(path: string): Promise<ArrayBuffer> {
        const item = this.app.vault.getAbstractFileByPath(path);
        if (item instanceof TFile) return await this.app.vault.readBinary(item);
        const adapter: any = this.app.vault.adapter;
        if (typeof adapter.readBinary === "function" &&
            await adapter.exists(path)) return await adapter.readBinary(path);
        throw new Error("OneDrive 隐藏恢复副本无法物理读回，已停止删除");
    }

    private async writeVerifiedRecovery(
        operation: RunTicket, path: string, data: ArrayBuffer,
        source: "vault" | "external" = "vault",
        sourceTrashPath?: string,
    ): Promise<void> {
        this.assertCurrent(operation);
        const local: any = (this.store as any).settings.local;
        const items = Array.isArray(local.recoveryEntries) ? local.recoveryEntries : [];
        const bytesUsed = items.reduce((sum: number, r: any) => sum + Number(r.contentBytes || 0), 0);
        if (items.length >= 5000 || bytesUsed + data.byteLength > 512 * 1024 * 1024) {
            throw new Error("OneDrive 恢复档案达到安全容量上限，停止危险写入");
        }
        const id = "onedrive_" + Date.now().toString(36) + "_" +
            Math.random().toString(36).slice(2, 10);
        const target = normalizePath(".suite-recovery/onedrive/" + id + ".bin");
        await this.ensureLocalFolder(target);
        if (Platform.isAndroidApp) {
            await this.createAndroidLocalCopy(target, data, () => this.assertCurrent(operation));
        } else {
            await this.app.vault.createBinary(target, data);
        }
        const reread = await this.readPhysicalBinary(target);
        if (!sameBytes(data, reread)) {
            throw new Error("OneDrive 冷备份物理读回哈希不符，禁止删除或覆盖原稿");
        }
        const digest = await this.sha256Bytes(data);
        const now = Date.now();
        items.push({
            id, bindingId: "onedrive", relativePath: path, source,
            sourceTrashPath,
            recoveryPath: target, deletedAt: now,
            expiresAt: now + 7 * 24 * 60 * 60 * 1000,
            verifiedAt: now, contentHash: digest,
            contentBytes: data.byteLength, resolvedAt: 0,
        });
        local.recoveryEntries = items;
        await (this.store as any).save();
        this.assertCurrent(operation);
    }

    private async createRecoveryFromRemote(
        operation: RunTicket, path: string, remote: RemoteFile,
    ): Promise<void> {
        const data = await this.remoteBytes(remote);
        await this.writeVerifiedRecovery(operation, path, data, "external");
    }

    listOneDriveRecoveries(): Array<any> {
        const local: any = (this.store as any).settings.local;
        return (local.recoveryEntries || []).filter((entry: any) =>
            entry.bindingId === "onedrive");
    }

    private androidRecoverySourcePath(record: any): string {
        const path = String(record.sourceTrashPath || "");
        const relative = String(record.relativePath || "");
        if (!path.startsWith(".trash/smartmd-onedrive/") ||
            !/^\.trash\/smartmd-onedrive\/[a-f0-9-]{36}\//.test(path) ||
            path.split("/").slice(3).join("/") !== relative ||
            path.split("/").some(part => !part || part === "." || part === ".." || /[\\:\u0000-\u001f]/.test(part))) {
            throw new Error("安卓回收原件路径不可信");
        }
        return path;
    }

    async exportAndroidTrashedOriginal(recoveryId: string): Promise<string> {
        if (!Platform.isAndroidApp || !this.store.installationProofIsValid()) throw new Error("安卓回收库身份不可信");
        const record = this.listOneDriveRecoveries().find(item => item.id === recoveryId);
        if (!record) throw new Error("找不到安卓回收记录");
        const path = this.androidRecoverySourcePath(record);
        const bytes = await this.readPhysicalBinary(path);
        const hash = await this.sha256Bytes(bytes);
        const extension = path.split("/").pop() || "original.md";
        const destination = "OneDrive恢复的笔记/" + crypto.randomUUID() + "-" + extension;
        await this.createAndroidLocalCopy(destination, bytes, () => {
            if (!this.store.installationProofIsValid()) throw new Error("安卓恢复库身份已变化");
        });
        record.trashExportPath = destination;
        record.trashExportHash = hash;
        await this.store.save();
        return destination;
    }

    async exportOneDriveRecoveryToNewNote(recoveryId: string): Promise<string> {
        const record = this.listOneDriveRecoveries().find((item: any) =>
            item.id === recoveryId);
        if (!record || !String(record.recoveryPath).startsWith(".suite-recovery/onedrive/") ||
            String(record.recoveryPath).includes("..") || !record.contentHash) {
            throw new Error("恢复路径或内容校验凭据不可信");
        }
        const data = await this.readPhysicalBinary(record.recoveryPath);
        if ((await this.sha256Bytes(data)) !== record.contentHash) {
            throw new Error("恢复副本内容发生变化，停止另存");
        }
        const base = String(record.relativePath || "untitled").split("/").pop()
            ?.replace(/[^a-zA-Z0-9_.-]/g, "_") || "untitled.md";
        // A previous export may have failed byte verification but left an output
        // behind. Never overwrite or remove it: use a bounded new filename.
        for (let attempt = 0; attempt < 4; attempt++) {
            const suffix = attempt === 0 ? "" : attempt === 3
                ? "-copy-" + crypto.randomUUID() : "-retry-" + attempt;
            const destination = normalizePath(
                "OneDrive恢复的笔记/" + record.id + suffix + "-" + base);
            if (await this.app.vault.adapter.exists(destination) ||
                this.app.vault.getAbstractFileByPath(destination)) {
                continue;
            }
            await this.ensureLocalFolder(destination);
            // createBinary rejects existing paths. If another writer wins after
            // the existence check, fail closed rather than risking overwrite.
            if (Platform.isAndroidApp) {
                await this.createAndroidLocalCopy(destination, data);
            } else {
                await this.app.vault.createBinary(destination, data);
            }
            if ((await this.sha256Bytes(await this.readPhysicalBinary(destination))) !== record.contentHash) {
                throw new Error("恢复另存读回核验失败，原副本保持锁定");
            }
            const previousExportedAt = record.exportedAt;
            const previousExportPath = record.exportPath;
            const previousResolvedAt = record.resolvedAt;
            record.exportedAt = Date.now();
            record.exportPath = destination;
            record.resolvedAt = record.sourceTrashPath ? 0 : record.exportedAt;
            try {
                await (this.store as any).save();
            } catch (error) {
                // A storage error may occur before or after disk commit.
                // Fail closed in memory rather than authorizing resolution
                // with an export marker that we cannot trust.
                record.exportedAt = previousExportedAt;
                record.exportPath = previousExportPath;
                record.resolvedAt = previousResolvedAt;
                throw error;
            }
            return destination;
        }
        throw new Error("安全恢复目标名称均已占用，原恢复副本保持锁定");
    }

    private async verifyOneDriveRecoveryExport(record: any): Promise<void> {
        if (!record || !record.verifiedAt || !record.exportedAt) {
            throw new Error("恢复副本尚未校验并另存，不能解除保留锁");
        }
        // The plugin may have restarted after a save error. A persisted
        // exportedAt marker alone is not proof that the restored file is
        // still present and byte-identical. Reverify both protected copies.
        const base = String(record.relativePath || "untitled").split("/").pop()
            ?.replace(/[^a-zA-Z0-9_.-]/g, "_") || "untitled.md";
        const approvedExports = [0, 1, 2].map((attempt) => normalizePath(
            "OneDrive恢复的笔记/" + record.id +
            (attempt === 0 ? "" : "-retry-" + attempt) + "-" + base));
        const archivePath = String(record.recoveryPath || "");
        const exportPath = String(record.exportPath || "");
        const copyPrefix = "OneDrive恢复的笔记/" + record.id + "-copy-";
        const uuid = exportPath.startsWith(copyPrefix) && exportPath.endsWith("-" + base)
            ? exportPath.slice(copyPrefix.length, -base.length - 1) : "";
        const trustedCopy = /^[a-f0-9]{8}(?:-[a-f0-9]{4}){3}-[a-f0-9]{12}$/.test(uuid);
        if ((!approvedExports.includes(exportPath) && !trustedCopy) ||
            !/^onedrive_[a-z0-9]+_[a-z0-9]+$/.test(String(record.id)) ||
            archivePath !== ".suite-recovery/onedrive/" + record.id + ".bin" ||
            !record.contentHash || !Number.isSafeInteger(record.contentBytes) ||
            record.contentBytes < 0) {
            throw new Error("恢复记录的目标或来源身份不可信，原副本保持锁定");
        }
        const archive = await this.readPhysicalBinary(archivePath);
        const exported = await this.readPhysicalBinary(exportPath);
        if (exported.byteLength !== record.contentBytes ||
            archive.byteLength !== record.contentBytes ||
            (await this.sha256Bytes(exported)) !== record.contentHash ||
            (await this.sha256Bytes(archive)) !== record.contentHash) {
            throw new Error("恢复文件或原副本已变化，拒绝解除恢复保护");
        }
        if (record.sourceTrashPath) {
            const source = this.androidRecoverySourcePath(record);
            if (await this.app.vault.adapter.exists(source)) {
                const originalHash = await this.sha256Bytes(await this.readPhysicalBinary(source));
                if (originalHash !== record.contentHash) {
                    const exportedOriginal = String(record.trashExportPath || "");
                    if (!exportedOriginal.startsWith("OneDrive恢复的笔记/") ||
                        exportedOriginal.includes("..") || exportedOriginal.includes("\\") ||
                        record.trashExportHash !== originalHash ||
                        await this.sha256Bytes(await this.readPhysicalBinary(exportedOriginal)) !== originalHash) {
                        throw new Error("回收目录有较晚版本，请先另存并核对回收原件");
                    }
                }
            }
        }
    }

    private async deleteRemote(
        item: RemoteFile, operation?: RunTicket, beforeDispatch?: () => void,
        checkPhysicalPath?: () => Promise<void>,
    ): Promise<void> {
        if (operation) this.assertCurrent(operation);
        if (!item.id || !item.etag) {
            throw new Error("OneDrive 删除缺少可信 Item ID 或 ETag，拒绝无条件删除");
        }
        const response = await this.fetchGraph("/drive/items/" + encodeURIComponent(item.id), {
            method: "DELETE",
            headers: { "If-Match": item.etag },
        }, operation, beforeDispatch, checkPhysicalPath);
        // 404 after conditional DELETE is not proof of our own successful commit.
        // Keep the old baseline and pending record until independent reconciliation.
        if (!response.ok) {
            const error = new Error(await this.fetchError(response,
                "OneDrive删除未获确认（含404），保留旧基线等待只读对账"));
            if (checkPhysicalPath || [401, 403, 408, 429, 500, 502, 503, 504, 507].includes(response.status)) {
                Object.assign(error, { status: response.status, retryAfterMs: graphRetryAfterMs(response.headers.get("retry-after")) });
            }
            throw error;
        }
    }

    private async ensureLocalFolder(filePath: string): Promise<void> {
        const slash = normalizePath(filePath).lastIndexOf("/");
        if (slash < 0) return;
        const parts = normalizePath(filePath.slice(0, slash)).split("/").filter(Boolean);
        let current = "";
        for (const part of parts) {
            current = current ? `${current}/${part}` : part;
            if (!(await this.app.vault.adapter.exists(current))) await this.app.vault.createFolder(current);
        }
    }

    private async createAndroidLocalCopy(
        path: string, bytes: ArrayBuffer, assertCurrent: () => void = () => {},
    ): Promise<void> {
        const adapter = this.app.vault.adapter;
        if (!getAndroidAppStorageRoot(adapter)) {
            await this.ensureLocalFolder(path);
            if (await adapter.exists(path) || this.app.vault.getAbstractFileByPath(path) ||
                this.app.vault.adapter !== adapter) throw new Error("安卓端恢复位置已占用，保留现有文件");
            assertCurrent();
            if (path.startsWith(".")) await adapter.writeBinary(path, bytes);
            else await this.app.vault.createBinary(path, bytes);
            if (!sameBytes(await adapter.readBinary(path), bytes)) throw new Error("安卓端副本落地读回失败：" + path);
            return;
        }
        const stage = ".suite-staging/onedrive/" + crypto.randomUUID() + "/copy.bin";
        await this.ensureLocalFolder(path);
        await this.ensureLocalFolder(stage);
        if (await this.app.vault.adapter.exists(stage)) throw new Error("安卓端暂存位置已占用");
        assertCurrent();
        await this.app.vault.createBinary(stage, bytes);
        if (!sameBytes(await this.readPhysicalBinary(stage), bytes)) {
            throw new Error("安卓端副本暂存读回失败：" + stage);
        }
        try {
            await placeAndroidStagedFile(this.app.vault, stage, path, assertCurrent);
        } catch (error) {
            throw new Error("安卓端副本未能落地，保留暂存 " + stage + "；" + safeOneDriveError(error));
        }
        if (!sameBytes(await this.readPhysicalBinary(path), bytes)) {
            throw new Error("安卓端副本落地读回失败：" + path);
        }
    }

    private async syncSharedState(operation: RunTicket): Promise<void> {
        this.assertCurrent(operation);
        const oneDrive = this.settings();
        if (oneDrive.pendingSharedUploadOperation) {
            if (!await this.reconcileOneDriveSharedUpload()) {
                throw new Error("原共享操作仍待只读对账，禁止新上传或自动重试");
            }
            return;
        }
        const localAccount = String((this.store as any).settings.local.microsoftAccountId || "");
        // Four-field sharing is opt-in, separately from granting note sync.
        if (oneDrive.sharedSettingsApproved !== true ||
            !this.canRun() || !localAccount ||
            oneDrive.sharedApprovalAccount !== localAccount ||
            oneDrive.sharedApprovalVault !== oneDrive.vaultId ||
            oneDrive.sharedApprovalDevice !== oneDrive.deviceId) return;
        const settings = (this.store as any).settings;
        const path = "vaults/" + oneDrive.vaultId + "/plugin-state.json";
        const localShared = sharedStateForSync(settings.shared);
        const localHash = await contentHash(JSON.stringify(localShared));
        const previously = String(oneDrive.sharedStateHash || "");
        this.assertCurrent(operation);
        const remotePayload = await this.readSharedState(path);
        // The remote result is data, never renewed authorization.
        this.assertCurrent(operation);
        if (remotePayload === null) {
            await this.sendApprovedSharedState(operation, path, localShared, localHash, null);
            return;
        }
        const safeRemote = sharedStateForSync(remotePayload.shared);
        const remoteHash = await contentHash(JSON.stringify(safeRemote));
        // A revocation during asynchronous hashing invalidates this result.
        this.assertCurrent(operation);
        if (remoteHash === localHash) {
            this.pendingSharedImport = null;
            oneDrive.sharedStateHash = localHash;
            return;
        }
        if (!previously || localHash === previously) {
            // A remote change is not consent to modify this device. Keep only
            // the four-field closed subset in an ephemeral, single-use preview.
            const fullPayloadHash = await contentHash(JSON.stringify(remotePayload));
            this.assertCurrent(operation);
            this.pendingSharedImport = {
                accountId: localAccount, vaultId: oneDrive.vaultId,
                deviceId: oneDrive.deviceId, localHash, remoteHash,
                baselineHash: previously,
                fullPayloadHash,
                fields: { ...safeRemote }, expiresAt: Date.now() + 5 * 60_000,
            };
            return;
        }
        this.pendingSharedImport = null;
        if (remoteHash !== previously && localHash !== previously) {
            throw new Error("双方共享字段均有更改，已暂停等待人工选择");
        }
        const meta = await this.readRemoteFileMetadata(path, true);
        if (!meta?.etag) throw new Error("无法验证云端插件共享设置版本，拒绝无条件覆盖");
        await this.sendApprovedSharedState(operation, path, localShared, localHash, meta);
    }

    getSharedSettingsError(): string {
        return this.sharedSettingsError;
    }

    private async sendApprovedSharedState(
        operation: RunTicket, path: string, fields: SharedSmartMDState,
        fieldsHash: string, expectedRemote: RemoteFile | null,
    ): Promise<void> {
        const settings = this.settings();
        if (!this.hasOneDriveSettingsSharingApproval()) {
            throw new Error("本次共享批准已失效或已使用，必须重新预览并独立确认");
        }
        const preview = this.getOneDriveSettingsSharingPreview();
        if (JSON.stringify(fields) !== JSON.stringify(preview.fields)) {
            throw new Error("共享字段在预览后变化，拒绝旧批准");
        }
        const body = new TextEncoder().encode(JSON.stringify({ updatedAt: Date.now(), shared: fields })).buffer;
        const bodyHash = await this.sha256Bytes(body);
        this.assertCurrent(operation);
        if (!this.hasOneDriveSettingsSharingApproval() ||
            preview.scope !== this.getOneDriveSettingsSharingPreview().scope) {
            throw new Error("共享批准在准备期间变化，必须重新确认");
        }
        const record = {
            operationId: settings.sharedApprovalId, runOperationId: operation.id,
            accountId: preview.accountId, vaultId: preview.vaultId, localVaultId: preview.localVaultId,
            deviceId: preview.deviceId, schema: preview.schema, scope: preview.scope,
            fieldsHash, bodyHash, path, phase: "ISSUED", createdAt: Date.now(),
            stagingPath: `_suite_staging/${operation.id.replace(/[\\/:*?"<>|.]/g, "_")}/${path.replace(/[\\/:*?"<>|]/g, "_")}`,
            finalId: "", finalEtag: "",
        };
        settings.sharedUploadApproved = false;
        settings.pendingSharedUploadOperation = record;
        // Single-use consumption and the original operation must be durable before any POST/PUT.
        await (this.store as any).save();
        const guard = () => {
            this.assertCurrent(operation);
            if (settings !== this.settings() || settings.pendingSharedUploadOperation !== record ||
                settings.sharedSettingsApproved !== true || settings.sharedApprovalId !== record.operationId ||
                settings.sharedApprovalSchema !== this.sharedSchema ||
                record.scope !== this.getOneDriveSettingsSharingPreview().scope) {
                throw new Error("共享字段、身份或许可已变化，停止原操作写入并保留对账记录");
            }
        };
        guard();
        try {
            const final = await this.uploadStaged(operation, path, body, true, expectedRemote, guard,
                async (item) => {
                    guard();
                    record.finalId = String(item?.id || "");
                    record.finalEtag = String(item?.eTag || item?.etag || "");
                    record.phase = "FINAL_ACKNOWLEDGED";
                    await (this.store as any).save();
                });
            guard();
            record.finalId = String(final?.id || "");
            record.finalEtag = String(final?.eTag || final?.etag || "");
            record.phase = "CLEANED";
            await (this.store as any).save();
            if (!await this.reconcileOneDriveSharedUpload()) {
                throw new Error("共享写入尚未读回可信版本及正文，不得报告成功");
            }
        } catch (error) {
            const failure: any = error;
            const status = Number(failure?.status || 0);
            const definiteFailure = !record.finalId && failure?.isUnknownCommit !== true &&
                [400, 401, 403, 404, 409, 412, 413, 507].includes(status);
            record.phase = definiteFailure ? "FAILED" : record.finalId
                ? "FINAL_ACKNOWLEDGED_RECONCILE_REQUIRED" : "RECONCILE_REQUIRED";
            Object.assign(record, { failureStatus: status || undefined,
                failureCategory: definiteFailure ? "CONFIRMED_FAILURE" : "UNRESOLVED", updatedAt: Date.now() });
            // A save failure never reopens the consumed approval; the durable ISSUED record also blocks retries.
            try { await (this.store as any).save(); } catch { /* preserve the original error */ }
            throw error;
        }
    }

    private async readSharedState(path: string, scope?: FirstUploadReadScope): Promise<any | null> {
        const response = await this.downloadResponse(`/drive/special/approot:/${encodePath(path)}`, scope);
        scope?.assertCurrent();
        if (response.status === 404) return null;
        if (response.status < 200 || response.status >= 300) throw new Error(response.json?.error?.message || response.text || `无法读取插件状态 [${path}]`);
        try {
            return JSON.parse(response.text);
        } catch {
            throw new Error("云端插件状态异常，已暂停同步且未覆盖本地设置");
        }
    }

    private mergeSharedState(local: any, remote: any): any {
        local = local && typeof local === "object" ? local : {};
        remote = sharedStateForSync(remote);
        const publishSpaces = new Map<string, any>();
        const locallyConfirmed = new Map<string, any>();
        const publishTombstones: Record<string, number> = {};
        for (const source of [remote.publishTombstones || {}, local.publishTombstones || {}]) {
            for (const [key, value] of Object.entries(source)) {
                publishTombstones[key] = Math.max(Number(publishTombstones[key] || 0), Number(value || 0));
            }
        }
        for (const space of local.publishSpaces || []) {
            const key = space.localId || space.remoteId || space.id || space.code;
            if (key && space.confirmed === true) locallyConfirmed.set(key, space);
        }
        for (const space of [...(remote.publishSpaces || []), ...(local.publishSpaces || [])]) {
            const identities = [space.id, space.localId, space.remoteId, space.code].filter(Boolean);
            const key = identities[0];
            if (!key || identities.some((identity) => publishTombstones[identity])) continue;
            const existing = publishSpaces.get(key);
            if (!existing || Number(space.lastPublishedAt || space.updatedAt || 0) >= Number(existing.lastPublishedAt || existing.updatedAt || 0)) publishSpaces.set(key, space);
        }
        const localAi = local.ai || {};
        const remoteAi = remote.ai || {};
        return {
            ...remote,
            ...local,
            publishServerUrl: SMARTKEEP_PUBLISH_URL,
            publishTombstones,
            publishSpaces: [...publishSpaces.values()].map((space: any) => {
                const key = space.localId || space.remoteId || space.id || space.code;
                const confirmation = locallyConfirmed.get(key);
                const scopeKey = `${space.scope?.type || space.scopeType || ""}:${space.scope?.path || space.scopePath || ""}`;
                const accountId = String((this.store as any).settings.local.smartkeepAccountId || "");
                const confirmed = confirmation?.confirmed === true
                    && confirmation?.confirmationScope === scopeKey
                    && confirmation?.confirmationAccount === accountId;
                return {
                    ...space,
                    confirmed,
                    runAllowed: confirmed && confirmation?.runAllowed === true,
                    confirmationScope: confirmed ? confirmation.confirmationScope : "",
                    confirmationAccount: confirmed ? confirmation.confirmationAccount : "",
                };
            }),
            aiApiUrl: local.aiApiUrl || localAi.endpoint || remote.aiApiUrl || remoteAi.endpoint || "",
            aiModel: local.aiModel || localAi.model || remote.aiModel || remoteAi.model || "",
            aiSystemPrompt:
                local.aiSystemPrompt || localAi.systemPrompt || remote.aiSystemPrompt || remoteAi.systemPrompt || "",
        };
    }

    private async fail(error: any, currentPath = "", silent = false): Promise<void> {
        const settings = this.settings();
        settings.status = "failed";

        const detail = currentPath ? ` (处理文件: ${redactOneDriveText(currentPath)})` : "";
        settings.lastError = `${safeOneDriveError(error)}${detail}`;
        console.error("[SmartMD OneDrive Sync Failure Detail]:", settings.lastError);
        await (this.store as any).save();
        if (!silent) new Notice(localize("OneDrive 同步已暂停：{0}", localizeMessage(settings.lastError)), 8000);
    }

    private async maintainOneDriveRecoveries(operation: RunTicket): Promise<void> {
        const local = this.store.settings.local;
        const now = Date.now();
        const graceMs = 7 * 24 * 60 * 60 * 1000;
        const resolved = new Map<any, string>();
        const removed = new Set<any>();
        for (const entry of this.listOneDriveRecoveries()) {
            const { deletedAt, verifiedAt, exportedAt, resolvedAt, expiresAt } = entry;
            if (![deletedAt, verifiedAt, exportedAt, expiresAt].every(Number.isSafeInteger) ||
                deletedAt <= 0 || verifiedAt < deletedAt || exportedAt < verifiedAt ||
                expiresAt < verifiedAt + graceMs || now < exportedAt) continue;
            const canClean = !entry.sourceTrashPath && Number.isSafeInteger(resolvedAt) &&
                resolvedAt >= exportedAt && now >= expiresAt && now - resolvedAt >= graceMs;
            if (resolvedAt && !canClean) continue;
            const snapshot = JSON.stringify(entry);
            try {
                await this.verifyOneDriveRecoveryExport(entry);
                this.assertCurrent(operation);
                if (!local.recoveryEntries?.includes(entry) || JSON.stringify(entry) !== snapshot) continue;
                if (!resolvedAt) {
                    resolved.set(entry, snapshot);
                    continue;
                }
                const adapter = this.app.vault.adapter;
                if (!await adapter.trashSystem(entry.recoveryPath)) continue;
                if (!await adapter.exists(entry.recoveryPath)) removed.add(entry);
            } catch {
                // Retain an unverifiable or inaccessible backup for another sync.
                this.assertCurrent(operation);
            }
        }
        if (!resolved.size && !removed.size) return;
        await this.store.saveWithIdentityGuard(() => this.assertCurrent(operation), () => {
            const previous = local.recoveryEntries;
            const changed = [...resolved].filter(([entry, snapshot]) =>
                previous?.includes(entry) && JSON.stringify(entry) === snapshot);
            const before = changed.map(([entry]) => [entry, entry.resolvedAt] as const);
            for (const [entry] of changed) entry.resolvedAt = now;
            const retained = previous?.filter(entry => !removed.has(entry));
            local.recoveryEntries = retained;
            return () => {
                for (const [entry, old] of before) if (entry.resolvedAt === now) entry.resolvedAt = old;
                if (local.recoveryEntries === retained) local.recoveryEntries = previous;
            };
        });
    }

    async isolateUnconfirmed(): Promise<void> {
        const settings = this.settings();
        if (settings.confirmed !== true) this.authority.isolate(settings);
        await (this.store as any).save();
    }

    private parseFileErrorToDiagnostic(path: string, fileError: any, phase: "UPLOAD" | "DOWNLOAD"): SyncFailureDiagnostic {
        const safeMsg = safeOneDriveError(fileError);

        let status = Number(fileError?.status || 0);
        if (!status) {
            const statusMatch = safeMsg.match(/HTTP\s+(\d{3})/i);
            if (statusMatch) status = Number(statusMatch[1]);
        }

        let code = String(fileError?.code || "");
        if (!code) {
            const codeMatch = safeMsg.match(/\[Graph API HTTP \d+ ([^\]]+)\]/i);
            if (codeMatch) code = codeMatch[1];
            else code = status === 400 ? "invalidRequest" : "unknownError";
        }

        // A Graph error code is untrusted server text, not a safe log field.
        // Preserve only a bounded set of diagnostic identifiers; otherwise
        // use a neutral value instead of persisting arbitrary credential text.
        const trustedGraphCode = /^(?:invalidRequest|nameAlreadyExists|itemNotFound|accessDenied|quotaExceeded|activityLimitReached|invalidAuthenticationToken|unauthenticated|generalException|resourceModified|preconditionFailed|notAllowed|unknownError|HTTP_[0-9]{3})$/i.test(code)
            ? code : "unknownError";
        let requestId = String(fileError?.requestId || "");
        if (!requestId) {
            const reqMatch = safeMsg.match(/request-id:\s*([a-zA-Z0-9-_]+)/i);
            if (reqMatch) requestId = reqMatch[1];
        }

        // Never derive suggested filenames from a raw absolute private path:
        // replacing slashes with '_' would otherwise leak each directory name.
        const safePath = redactOneDriveText(path);
        const fileName = safePath.split("/").pop() || safePath;
        const hasInvalidChars = /[:*?"<>|\\]/.test(fileName) || /[\s.]$/.test(fileName) || safeMsg.toLowerCase().includes("invalid characters") || safeMsg.toLowerCase().includes("namecontainsinvalidcharacters");

        let category: SyncFailureCategory = "UNKNOWN_ERROR";
        let humanReason = status ? `云端传输遇到异常 (HTTP ${status})` : "文件同步尚未完成";
        let humanExplanation = status
            ? `OneDrive 返回 HTTP ${status}。`
            : "这次同步没有记录到 OneDrive 的 HTTP 响应。";
        let remediation: ActionableRemediation = {
            actionType: "RETRY_LATER",
            badgeText: "稍后重试",
            actionTitle: "重新尝试同步",
            stepInstructions: [
                "1. 检查本地网络连接；",
                "2. 重新点击设置中的【立即同步】；",
                "3. 若仍无法同步，请复制下方技术凭据反馈排查。"
            ],
        };

        if (hasInvalidChars || (status === 400 && /[:*?"<>|]/.test(path))) {
            category = "PATH_INVALID_CHARS";
            humanReason = "文件名或路径包含微软 OneDrive 不支持的非法字符";
            humanExplanation = "OneDrive 个人版禁止文件或目录名中包含 : * ? \" < > | 等特殊符号，或以空格/句点结尾。";
            const sanitizedName = fileName.replace(/[:*?"<>|\\]/g, "_").replace(/[\s.]+$/, "");
            const parentDir = path.includes("/") ? path.slice(0, path.lastIndexOf("/") + 1) : "";
            const suggestedNewPath = `${parentDir}${sanitizedName}`;
            remediation = {
                actionType: "RENAME_FILE",
                badgeText: "建议重命名",
                actionTitle: "修改文件名中的特殊符号",
                stepInstructions: [
                    `1. 在 Obsidian 知识库中找到文件：${path}；`,
                    `2. 将文件名中的特殊字符（如冒号、问号、星号）修改为普通字符或下划线；`,
                    `3. 修改后无需手动上传，再次点击【立即同步】即可自动上传。`
                ],
                suggestedNewPath,
                canAutoFix: true,
            };
        } else if (status === 429 || status === 503 || safeMsg.includes("429") || safeMsg.includes("activityLimitReached")) {
            category = "RATE_LIMITED";
            humanReason = "触发微软 OneDrive 突发请求频次流控限制 (HTTP 429)";
            humanExplanation = "同步请求速度超过了微软个人版网关的突发限流阈值，云端暂时拒绝响应。";
            remediation = {
                actionType: "RETRY_LATER",
                badgeText: "稍后重试",
                actionTitle: "等待云端流控恢复",
                stepInstructions: [
                    "1. 插件内置的指数退避重试已自动保护，已成功文件已安全落盘；",
                    "2. 请静待 1~2 分钟让微软网关额度恢复；",
                    "3. 再次点击【立即同步】，插件将精准续传剩余失败文件。"
                ],
            };
        } else if (status === 507 || safeMsg.includes("507") || safeMsg.includes("quotaExceeded")) {
            category = "QUOTA_EXCEEDED";
            humanReason = "微软 OneDrive 个人云端存储容量已满 (HTTP 507)";
            humanExplanation = "您的 OneDrive 存储配额已耗尽，无法再容纳新上传的内容。";
            remediation = {
                actionType: "FREE_SPACE",
                badgeText: "空间不足",
                actionTitle: "清理云盘空间或升级配额",
                stepInstructions: [
                    "1. 登录 onedrive.live.com 查看云盘容量；",
                    "2. 清理回收站或删除不常用的大文件；",
                    "3. 腾出空间后回到 Obsidian 重新点击【立即同步】。"
                ],
            };
        } else if (status === 401 || safeMsg.includes("401") || safeMsg.includes("unauthenticated") || safeMsg.includes("tokenExpired")) {
            category = "AUTH_EXPIRED";
            humanReason = "OneDrive 登录令牌失效或授权过期 (HTTP 401)";
            humanExplanation = "由于长时间未使用或云端安全策略，当前设备的微软登录授权已失效。";
            remediation = {
                actionType: "RELOGIN",
                badgeText: "需要重登",
                actionTitle: "重新登录微软账号",
                stepInstructions: [
                    "1. 在 SmartMD 设置面板点击【重新登录】；",
                    "2. 按照弹窗提示完成浏览器设备码授权；",
                    "3. 登录后重新检查同步范围，再确认继续未完成的文件同步。"
                ],
            };
        } else if (status === 504 || safeMsg.includes("504") || safeMsg.includes("timeout") || safeMsg.includes("fetch failed") || safeMsg.includes("ECONNRESET")) {
            category = "NETWORK_TIMEOUT";
            humanReason = "连接微软 Graph API 服务器超时或网络中断";
            humanExplanation = "本地到微软服务器的网络连接不稳定，请求在规定时间内未获得响应。";
            remediation = {
                actionType: "CHECK_NETWORK",
                badgeText: "网络超时",
                actionTitle: "检查网络与代理设置",
                stepInstructions: [
                    "1. 检查系统网络连接与梯子/代理软件是否通畅；",
                    "2. 确保可以正常访问 graph.microsoft.com；",
                    "3. 恢复后点击【立即同步】，断点续传将直接接续。"
                ],
            };
        } else if (status === 409 || safeMsg.includes("409") || safeMsg.includes("nameAlreadyExists")) {
            category = "NAME_CONFLICT";
            humanReason = "云端存在同名或大小写不敏感冲突 (HTTP 409)";
            humanExplanation = "OneDrive 个人版是大小写不敏感的，若存在类似 Readme.md 与 readme.md 将发生冲突。";
            remediation = {
                actionType: "RENAME_FILE",
                badgeText: "名称冲突",
                actionTitle: "重命名冲突文件",
                stepInstructions: [
                    "1. 检查是否存在同名但大小写不同的文件；",
                    "2. 将其中一个文件更名；",
                    "3. 再次点击【立即同步】。"
                ],
            };
        }

        return {
            id: `diag_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 8)}`,
            relativePath: redactOneDriveText(path),
            phase,
            category,
            humanReason: redactOneDriveText(humanReason),
            humanExplanation: redactOneDriveText(humanExplanation),
            evidence: {
                httpStatus: status,
                gatewayErrorCode: redactOneDriveText(trustedGraphCode),
                requestId: redactOneDriveText(requestId),
                occurredAt: Date.now(),
                httpMethod: phase === "UPLOAD" ? "PUT" : "GET",
                sanitizedEndpoint: `/drive/items/...:${redactOneDriveText(path)}`,
                rawErrorMessage: safeMsg,
            },
            remediation: {
                ...remediation,
                suggestedNewPath: remediation.suggestedNewPath
                    ? redactOneDriveText(remediation.suggestedNewPath) : undefined,
                stepInstructions: remediation.stepInstructions
                    .map((step) => redactOneDriveText(step)),
            },
        };
    }
}
