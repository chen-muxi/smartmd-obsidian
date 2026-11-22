import { App } from "obsidian";

/** Inventory only. ELIGIBLE is a preview prerequisite, NEVER execution consent. */
export type ReceiverCandidateAudit = Readonly<{
    status: "ELIGIBLE_EMPTY_RECEIVER" | "ELIGIBLE_INCOMPLETE_RECEIVER" | "EMPTY_INVENTORY_UNPROVEN" | "BLOCKED";
    reason: string;
    inspectedEntries: number;
    rootBirthtimeMs: number | null;
    bindingAuthorized: false;
    /** Root-only fact; never implies empty inventory or binding authority. */
    independentPhysicalRootVerified: boolean;
}>;

/** Scope proven separately by SmartMDStore's per-physical-root installation key. */
export type ReceiverCurrentInstallation = Readonly<{
    accountId: string; vaultId: string; deviceId: string; installationProofId: string;
}>;

/** Only verified former first-download paths; NOT authority to write. */
export type ReceiverIncompleteHistory = Readonly<{
    cloudVaultId: string;
    trackedPaths: readonly string[];
}>;

type FsReadOnly = {
    lstat(path: string): Promise<{ isDirectory(): boolean; isSymbolicLink(): boolean; birthtimeMs?: number }>;
    realpath(path: string): Promise<string>;
    readdir(path: string): Promise<Array<string | { name: string }>>;
    readFile?(path: string, encoding: "utf8"): Promise<string>;
};

const allowedConfigFiles = new Set([
    "app.json", "appearance.json", "core-plugins.json", "core-plugins-migration.json",
    "hotkeys.json", "graph.json", "workspace.json", "workspace-mobile.json",
    "community-plugins.json",
]);
const allowedPluginFiles = new Set(["main.js", "manifest.json", "styles.css"]);

/**
 * A saved SmartMD initialization file may be present before the user selects
 * OneDrive: SmartMDStore.load() saves it on first launch. Check the actual
 * on-disk object, NOT a filename allowlist or a copied settings assertion.
 */
function freshLocalSmartMDState(json: string, scope: ReceiverCurrentInstallation): boolean {
    try {
        const payload = JSON.parse(json);
        const local = payload?.suiteV2?.local;
        const sync = local?.oneDrive;
        if (!local || !sync || !scope.accountId || !scope.vaultId ||
            !scope.deviceId || !scope.installationProofId ||
            local.microsoftAccountId !== scope.accountId ||
            sync.vaultId !== scope.vaultId || sync.deviceId !== scope.deviceId ||
            sync.installationProofId !== scope.installationProofId) return false;
        const noEntries = (value: unknown) => !value ||
            (Array.isArray(value) ? value.length === 0 :
                typeof value === "object" && Object.keys(value).length === 0);
        if (!noEntries(local.recoveryEntries) || !noEntries(local.externalBindings) ||
            !noEntries(local.externalRecoveryArchives) || !noEntries(local.externalConflicts) ||
            !noEntries(sync.fileStates) || !noEntries(sync.pendingUploadOperations) ||
            !noEntries(sync.pendingDeletions) || !noEntries(sync.pendingConflicts) ||
            !noEntries(sync.failedDiagnostics) || !noEntries(sync.failedItems)) return false;
        if (sync.confirmed || sync.runAllowed || sync.enabled || sync.firstSyncCompleted ||
            Number(sync.lastSuccessAt || 0) !== 0 || sync.confirmationVaultId ||
            sync.confirmationAccountId || sync.cloudVaultId || sync.localVaultId ||
            sync.firstCommitJournal || sync.firstFileJournal || sync.lastSyncSummary ||
            sync.sharedStateHash || sync.sharedSettingsApproved ||
            sync.sharedApprovalVault || sync.pendingSharedImport ||
            sync.remoteDeltaCache || sync.deltaLink) return false;
        return true;
    } catch {
        return false;
    }
}

/**
 * A cold-restarted, partially downloaded receiver must have the same
 * installation, account, registered cloud Vault and EXACT tracked history.
 * A data.json copy or a fabricated empty note list is insufficient.
 */
function incompleteLocalSmartMDState(
    json: string, scope: ReceiverCurrentInstallation, history: ReceiverIncompleteHistory,
): boolean {
    try {
        const local = JSON.parse(json)?.suiteV2?.local;
        const sync = local?.oneDrive;
        if (!sync || !scope.accountId || !scope.installationProofId ||
            local.microsoftAccountId !== scope.accountId ||
            sync.vaultId !== scope.vaultId || sync.deviceId !== scope.deviceId ||
            sync.installationProofId !== scope.installationProofId ||
            sync.vaultId !== history.cloudVaultId ||
            sync.confirmationVaultId !== history.cloudVaultId ||
            sync.confirmationAccountId !== scope.accountId ||
            sync.confirmed !== true || sync.firstSyncCompleted === true ||
            sync.enabled === true || Number(sync.lastSuccessAt || 0) !== 0) return false;
        const empty = (value: unknown): boolean => !value ||
            (Array.isArray(value) ? value.length === 0 :
                typeof value === "object" && Object.keys(value).length === 0);
        if (!empty(local.recoveryEntries) || !empty(local.externalBindings) ||
            !empty(local.externalRecoveryArchives) || !empty(local.externalConflicts) ||
            !empty(sync.pendingUploadOperations) || !empty(sync.pendingConflicts) ||
            !empty(sync.pendingDeletions) || sync.firstCommitJournal ||
            sync.firstFileJournal) return false;
        const existing = Object.keys(sync.fileStates || {}).sort();
        const allowed = [...history.trackedPaths].sort();
        return allowed.length > 0 && JSON.stringify(existing) === JSON.stringify(allowed) &&
            existing.every((key) => {
                const item = sync.fileStates[key];
                return !!item && !!item.remoteId && !!item.remoteEtag &&
                    /^[0-9a-f]{64}$/i.test(String(item.localHash || ""));
            });
    } catch {
        return false;
    }
}

/** Inventory alone never grants binding or permission; the final consent does. */
export async function inspectReceiverCandidateInventory(
    app: App, sourcePhysicalRoot: string | null, readOnlyFs?: FsReadOnly,
    currentInstallation?: ReceiverCurrentInstallation,
    incomplete?: ReceiverIncompleteHistory,
): Promise<ReceiverCandidateAudit> {
    let inspectedEntries = 0;
    let rootBirthtimeMs: number | null = null;
    let independentPhysicalRootVerified = false;
    let verifiedCurrentSmartMDData = false;
    const verifiedNotes = new Set<string>();
    const tracked = incomplete ? new Set(incomplete.trackedPaths) : null;
    const result = (status: ReceiverCandidateAudit["status"], reason: string): ReceiverCandidateAudit =>
        ({ status, reason, inspectedEntries, rootBirthtimeMs, bindingAuthorized: false, independentPhysicalRootVerified });
    try {
        if (incomplete && (!currentInstallation || !incomplete.cloudVaultId ||
            !tracked?.size || [...tracked].some((item) =>
                !item || item.includes("\\") || item.includes(":") ||
                item.length > 1000 || item.split("/").some((part) =>
                    !part || part === "." || part === ".." || part.startsWith("."))))) {
            return result("BLOCKED", "INVALID_INCOMPLETE_RECEIVER_HISTORY");
        }
        const adapter: any = app.vault.adapter;
        if (!adapter || typeof adapter.getBasePath !== "function") return result("BLOCKED", "NO_DESKTOP_ROOT");
        const paths = require("path").win32 as typeof import("path").win32;
        const root = paths.normalize(String(adapter.getBasePath() || ""));
        const isWindowsRoot = /^[A-Za-z]:\\/.test(root) && root.length > 3;
        if (!isWindowsRoot) return result("BLOCKED", "INVALID_WINDOWS_ROOT");
        const fs = readOnlyFs || (require("fs").promises as FsReadOnly);
        const physical = await fs.realpath(root);
        if (paths.normalize(physical).toLocaleLowerCase("en-US") !== root.toLocaleLowerCase("en-US"))
            return result("BLOCKED", "ROOT_REDIRECT_OR_JUNCTION");
        if (sourcePhysicalRoot && paths.normalize(sourcePhysicalRoot).toLocaleLowerCase("en-US") === root.toLocaleLowerCase("en-US"))
            return result("BLOCKED", "SAME_PHYSICAL_ROOT");
        const rootStat = await fs.lstat(root);
        if (!rootStat.isDirectory() || rootStat.isSymbolicLink()) return result("BLOCKED", "INVALID_ROOT_TYPE");
        if (Number.isFinite(rootStat.birthtimeMs) && Number(rootStat.birthtimeMs) > 0)
            rootBirthtimeMs = Number(rootStat.birthtimeMs);
        if (currentInstallation) {
            // A nested folder is not an independent Obsidian Vault, even if empty.
            let parent = paths.dirname(root);
            while (parent !== paths.dirname(parent)) {
                try {
                    await fs.lstat(paths.join(parent, ".obsidian"));
                    return result("BLOCKED", "NESTED_INSIDE_ANOTHER_VAULT");
                } catch (error: any) {
                    if (error?.code !== "ENOENT") return result("BLOCKED", "PARENT_ROOT_UNVERIFIABLE");
                }
                parent = paths.dirname(parent);
            }
        }
        independentPhysicalRootVerified = Boolean(currentInstallation);
        const pending = [{ absolute: root, relative: "", depth: 0 }];
        const maxEntries = incomplete ? 4096 : 256;
        while (pending.length) {
            const current = pending.pop()!;
            if (current.depth > (incomplete ? 16 : 4)) return result("BLOCKED", "CONFIG_DEPTH_LIMIT");
            const names = await fs.readdir(current.absolute);
            for (const entry of names) {
                inspectedEntries++;
                if (inspectedEntries > maxEntries) return result("BLOCKED", "INVENTORY_LIMIT");
                const name = typeof entry === "string" ? entry : String(entry.name || "");
                if (!name || name === "." || name === ".." || name.includes("/") || name.includes("\\"))
                    return result("BLOCKED", "INVALID_ENTRY_NAME");
                const rel = current.relative ? current.relative + "/" + name : name;
                const target = paths.join(current.absolute, name);
                const stat = await fs.lstat(target);
                if (stat.isSymbolicLink()) return result("BLOCKED", "LINK_OR_REPARSE_POINT");
                if (stat.isDirectory()) {
                    const obConfig = rel === ".obsidian" || rel === ".obsidian/plugins" ||
                        rel === ".obsidian/plugins/smartmd";
                    const knownNoteFolder = Boolean(tracked && [...tracked].some(
                        (file) => file.startsWith(rel + "/")));
                    // Download staging folders may remain physically EMPTY
                    // after the verified staged file is removed. No file
                    // inside those folders is approved historical evidence.
                    const emptyDownloadStage = Boolean(incomplete &&
                        (rel === ".suite-staging" || rel === ".suite-staging/onedrive" ||
                         rel.startsWith(".suite-staging/onedrive/")));
                    if (obConfig || knownNoteFolder || emptyDownloadStage) {
                        pending.push({ absolute: target, relative: rel, depth: current.depth + 1 });
                    } else return result("BLOCKED", "UNEXPECTED_DIRECTORY_OR_HISTORY");
                } else {
                    const config = current.relative === ".obsidian" && allowedConfigFiles.has(name);
                    const plugin = current.relative === ".obsidian/plugins/smartmd" && allowedPluginFiles.has(name);
                    if (rel === ".obsidian/plugins/smartmd/data.json" && currentInstallation) {
                        const text = fs.readFile && (stat as any).size > 0 &&
                            (stat as any).size <= 512 * 1024
                            ? await fs.readFile(target, "utf8") : "";
                        const authorizedHistory = incomplete
                            ? incompleteLocalSmartMDState(text, currentInstallation, incomplete)
                            : freshLocalSmartMDState(text, currentInstallation);
                        if (!authorizedHistory) return result("BLOCKED", "UNTRUSTED_SMARTMD_HISTORY");
                        verifiedCurrentSmartMDData = true;
                    } else if (tracked?.has(rel)) {
                        verifiedNotes.add(rel);
                    } else if (!config && !plugin) {
                        return result("BLOCKED", "UNEXPECTED_FILE_OR_HISTORY");
                    }
                }
            }
        }
        if (currentInstallation) {
            if (incomplete) {
                return verifiedCurrentSmartMDData &&
                    verifiedNotes.size === tracked?.size
                    ? result("ELIGIBLE_INCOMPLETE_RECEIVER", "VERIFIED_ONLY_PRIOR_DOWNLOADED_FILES")
                    : result("BLOCKED", "INCOMPLETE_RECEIVER_HISTORY_NOT_PROVEN");
            }
            return verifiedCurrentSmartMDData
                ? result("ELIGIBLE_EMPTY_RECEIVER", "CLEAN_INDEPENDENT_VAULT_WITH_CURRENT_INSTALLATION")
                : result("BLOCKED", "CURRENT_INSTALLATION_DATA_MISSING");
        }
        return result("EMPTY_INVENTORY_UNPROVEN", "TRUSTED_CREATION_ISSUER_MISSING");
    } catch {
        return result("BLOCKED", "INVENTORY_INCOMPLETE_OR_UNREADABLE");
    }
}
