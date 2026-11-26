export type PublishScopeType = "file" | "folder" | "vault";
export type PublishVisibility = "public" | "password";
export type FeedbackStatus = "open" | "resolved" | "rejected";
export type BookmarkSyncStatus = "pending" | "synced" | "detached" | "user_removed" | "not_applicable" | "failed";
export type PublishOperationStatus = "preparing" | "submitting" | "reconciling" | "success" | "needs_confirmation" | "pending_approval" | "failed";
export type PublishReconciliationStatus = "matched" | "pending_confirmation" | "needs_restore" | "remote_withdrawn";

export interface PublishPendingOperation {
    action: "update" | "withdraw" | "access";
    requestId: string;
    expectedVersion: number;
    syncId?: string;
    includeNew?: boolean;
    includeDeletes?: boolean;
    desiredAccess?: PublishVisibility;
}

export interface PublishScope {
    type: "document" | "folder" | "vault";
    path: string;
    title: string;
}

export interface PublishedDocument {
    id: string;
    localPath: string;
    publicPath: string;
    modifiedAt: number;
    url?: string;
}

export interface PublishedSpace {
    id: string;
    code: string;
    title: string;
    url: string;
    access: PublishVisibility;
    scope: PublishScope;
    documents: PublishedDocument[];
    pendingDocuments: string[];
    pendingDeletes: string[];
    feedbackOpen: number;
    bookmarkStatus: BookmarkSyncStatus;
    bookmarkError: string;
    lastPublishedAt: number;
    cloudVersion?: number;
    reconciliationStatus?: PublishReconciliationStatus;
    operationStatus?: PublishOperationStatus;
    statusMessage?: string;
    scopeKnown?: boolean;
    pendingOperation?: PublishPendingOperation;
}

export interface PublishedDocumentState {
    id: string;
    path: string;
    approved: boolean;
    lastHash: string;
    updatedAt: number;
}

export interface PublishedSpaceState {
    localId: string;
    remoteId: string;
    code: string;
    title: string;
    scopeType: PublishScopeType;
    scopePath: string;
    visibility: PublishVisibility;
    autoUpdate: boolean;
    documents: Record<string, PublishedDocumentState>;
    pendingPaths: string[];
    pendingDeletes: string[];
    lastSuccessAt: number;
}

export interface SharedSmartMDState {
    updatedAt: number;
    publishSpaces: PublishedSpaceState[];
    publishTombstones: Record<string, number>;
    aiApiUrl: string;
    aiModel: string;
    aiSystemPrompt: string;
    publishServerUrl: string;
}

export type SyncFailureCategory =
    | "PATH_INVALID_CHARS"
    | "PATH_TOO_LONG"
    | "FILE_TOO_LARGE"
    | "NAME_CONFLICT"
    | "AUTH_EXPIRED"
    | "QUOTA_EXCEEDED"
    | "RATE_LIMITED"
    | "NETWORK_TIMEOUT"
    | "FILE_LOCKED_OR_CHANGED"
    | "UNKNOWN_ERROR";

export interface TechnicalEvidence {
    httpStatus: number;
    gatewayErrorCode: string;
    requestId: string;
    occurredAt: number;
    httpMethod: string;
    sanitizedEndpoint?: string;
    rawErrorMessage: string;
}

export interface ActionableRemediation {
    actionType: "RENAME_FILE" | "SPLIT_OR_IGNORE" | "RELOGIN" | "RETRY_LATER" | "FREE_SPACE" | "CHECK_NETWORK";
    badgeText: string;
    actionTitle: string;
    stepInstructions: string[];
    suggestedNewPath?: string;
    canAutoFix?: boolean;
}

export interface SyncFailureDiagnostic {
    id: string;
    relativePath: string;
    fileSize?: number;
    phase: "SCAN" | "UPLOAD" | "DOWNLOAD" | "DELETE" | "STAGING" | "RECONCILE";
    evidence: TechnicalEvidence;
    category: SyncFailureCategory;
    humanReason: string;
    humanExplanation: string;
    remediation: ActionableRemediation;
}

export interface SyncSummary {
    total: number;
    processed: number;
    success: number;
    failed: number;
    unchanged: number;
    /** Saved conflict copies are unresolved, not successful synchronized files. */
    conflicts?: number;
    durationMs: number;
    finishedAt: number;
}

export interface SyncFileState {
    localMtime: number;
    localSize: number;
    remoteEtag: string;
    remoteId: string;
    localHash?: string;
}

export interface OneDriveSyncSettings {
    vaultId: string;
    vaultName: string;
    deviceId: string;
    deviceName: string;
    autoMode: "manual" | "smart" | "custom";
    enabled: boolean;
    onStartup: boolean;
    afterSave: boolean;
    firstSyncCompleted: boolean;
    intervalMinutes: number;
    maxFileSizeMb: number;
    ignorePatterns: string[];
    lastSuccessAt: number;
    status: string;
    lastError: string;
    lastSkippedCount: number;
    sharedStateHash: string;
    fileStates: Record<string, SyncFileState>;
    androidFirstDownload?: AndroidFirstDownload;
    androidFirstMerge?: { accountId: string; cloudVaultId: string; proofId: string; deviceId: string };
    runAllowed: boolean;
    confirmed: boolean;
    /** Locally persisted account/cloud Vault binding; identity is not execution consent. */
    confirmationVaultId?: string;
    confirmationAccountId?: string;
    /** Per-installation proof exists outside the copied Vault and is checked at load. */
    installationProofId?: string;
    /** A deletion request is a preview, never a permission to mutate another side. */
    pendingDeletions?: Record<string, {
        /** Current-source scoped human review only; missing legacy fields deny approval. */
        accountId?: string;
        vaultId?: string;
        deviceId?: string;
        direction: "cloud" | "vault";
        remoteId?: string;
        remoteEtag?: string;
        localMtime?: number;
        localSize?: number;
        localHash?: string;
        requestedAt: number;
    }>;
    /** Durable local-only UNKNOWN/in-flight upload ledger.
     * Links the mother path to the stage, original operation, local source SHA,
     * target version and phase. A stored record DENIES fresh permission and
     * is never treated as proof of successful commit. Do not store uploadUrl,
     * access tokens or other credential-bearing Graph session URLs. */
    pendingUploadOperations?: Record<string, {
        operationId: string;
        accountId: string; vaultId: string; deviceId: string;
        motherPath: string; direction: "upload";
        sourceSha256: string; sourceBytes: number; sourceMtime?: number; sourceBaseline?: string;
        expectedRemoteId: string; expectedRemoteEtag: string;
        stagingPath: string;
        phase: "STAGE_SESSION_ISSUED" | "STAGE_PUT_ISSUED" |
            "STAGE_ACKNOWLEDGED" | "FINAL_PUT_ISSUED" |
            "FINAL_ACKNOWLEDGED" | "CLEANUP_ISSUED" |
            "CLEANUP_VERIFIED" | "RECONCILE_REQUIRED";
        stageId?: string; stageEtag?: string;
        finalId?: string; finalEtag?: string;
        sessionFingerprint?: string;
        updatedAt: number;
    }>;
    /** Historical evidence from releases that created separate copies. */
    pendingConflicts?: Record<string, {
        accountId: string;
        vaultId: string;
        deviceId: string;
        remoteId: string;
        remoteEtag: string;
        localCopyPath: string;
        remoteCopyId: string;
        remoteCopyEtag: string;
        localHash: string;
        createdAt: number;
    }>;
    lastSyncSummary?: SyncSummary;
    failedDiagnostics?: SyncFailureDiagnostic[];
    /** Local-only; do not include in shared OneDrive state. */
    firstCommitJournal?: {

        ticketId: string;
        digest: string;
        identity: { accountId: string; localVaultId: string; deviceId: string; cloudVaultId: string };
        stage: "REGISTERING" | "REGISTERED" | "RECONCILE_REQUIRED";
        registryEtag: string;
        cloudVaultName: string;
        registryBefore: Array<{ id: string; name: string; updatedAt: number; registrationOpId?: string }>;
        updatedAt: number;
    };
    /** Verified independently; legacy vaultId remains unchanged until full sync approval. */
    localVaultId?: string;
    cloudVaultId?: string;
    /** Durable pending copy; retained on interruption, never transmitted as shared settings. */
    firstFileJournal?: {
        ticketId: string;
        digest: string;
        registrationOpId: string;
        accountId: string;
        localVaultId: string;
        cloudVaultId: string;
        deviceId: string;
        registryDigest: string;
        ruleDigest: string;
        files: Array<{ path: string; size: number; mtime: number; sha256: string }>;
        completed: Record<string, { id: string; etag: string }>;
        stage: "TRANSFERRING" | "RECONCILE_REQUIRED" | "FILES_COPIED";
        updatedAt: number;
    };
}

export function hasUnfinishedOneDriveSetup(
    settings: OneDriveSyncSettings, accountId: string, cloudVaultId: string, deviceId: string,
): boolean {
    const registration = settings.firstCommitJournal;
    const firstFile = settings.firstFileJournal;
    const matchesBinding = (receipt: { accountId: string; cloudVaultId: string; deviceId: string } | undefined) =>
        settings.firstSyncCompleted && receipt?.accountId === accountId &&
        receipt.cloudVaultId === cloudVaultId && receipt.deviceId === deviceId;
    return Boolean(
        (registration && (registration.stage !== "REGISTERED" || !matchesBinding(registration.identity))) ||
        (firstFile && (firstFile.stage !== "FILES_COPIED" || !matchesBinding(firstFile))),
    );
}

export interface AndroidFirstDownload {
    path: string;
    stagingPath: string;
    accountId: string;
    vaultId: string;
    deviceId: string;
    proofId: string;
    remoteId: string;
    remoteEtag: string;
    size: number;
    sha256: string;
}

export interface ExternalFileSnapshot {
    externalHash: string;
    vaultHash: string;
    externalExists: boolean;
    vaultExists: boolean;
}

export interface ExternalConflict {
    id: string;
    bindingId: string;
    relativePath: string;
    externalContent: string;
    vaultContent: string;
    createdAt: number;
}

export interface RecoveryEntry {
    id: string;
    bindingId: string;
    relativePath: string;
    source: "external" | "vault";
    recoveryPath: string;
    deletedAt: number;
    expiresAt: number;
}

export interface ExternalFolderBinding {
    id: string;
    externalPath: string;
    targetPath: string;
    nickname: string;
    paused?: boolean;
    status: "pending" | "syncing" | "normal" | "active" | "paused" | "offline" | "conflict" | "failed" | string;
    lastSuccessAt?: number;
    lastSyncTime?: number;
    lastError?: string;
    snapshot?: Record<string, ExternalFileSnapshot>;
    fileStates?: Record<string, any>;
    conflicts?: any[];
    recoveryItems?: any[];
    runAllowed?: boolean;
    confirmed?: boolean;
    confirmationScope?: string;
    filterConfig?: ExternalFilterConfig;
    appliedFilterHash?: string;
    [key: string]: any;
}

export interface LocalSmartMDState {
    interfaceLanguage: "en" | "zh-CN";
    smartkeepAccountName: string;
    smartkeepAccountId: string;
    microsoftAccountName: string;
    microsoftAccountId: string;
    oneDrive: OneDriveSyncSettings;
    externalBindings: ExternalFolderBinding[];
    externalRecoveryArchives: any[];
    externalConflicts: ExternalConflict[];
    recoveryEntries: RecoveryEntry[];
    legacyMigrationChecked: boolean;
    rolloutEnabled: boolean;
}

export interface SmartMDSettings {
    schemaVersion: number;
    shared: SharedSmartMDState;
    local: LocalSmartMDState;
}

export interface PublishDocumentPayload {
    id: string;
    path: string;
    title: string;
    markdown: string;
    contentHash: string;
}

export interface PublishPreview {
    documents: PublishDocumentPayload[];
    blockedImages: Array<{ path: string; image: string }>;
    pendingCount: number;
}

export interface FeedbackItem {
    id: number;
    documentId: string;
    documentPath: string;
    quote: string;
    reporterName: string;
    body: string;
    status: FeedbackStatus;
    resolution: string;
    createdAt: number;
    resolvedAt: number;
}

export interface SyncPreview {
    localFiles: number;
    remoteFiles: number;
    localOnly: number;
    remoteOnly: number;
    possibleConflicts: number;
    skippedLargeFiles: number;
}

export const SMARTKEEP_PUBLISH_URL = "https://md.smartkeep.cn";

export const DEFAULT_SETTINGS: SmartMDSettings = {
    schemaVersion: 2,
    shared: {
        updatedAt: 0,
        publishSpaces: [],
        publishTombstones: {},
        aiApiUrl: "https://api.openai.com/v1/chat/completions",
        aiModel: "gpt-4o",
        aiSystemPrompt: "You are a careful Markdown editor. Preserve the author's meaning and formatting unless explicitly asked to change them.",
        publishServerUrl: SMARTKEEP_PUBLISH_URL,
    },
    local: {
        interfaceLanguage: "en",
        smartkeepAccountName: "",
        smartkeepAccountId: "",
        microsoftAccountName: "",
        microsoftAccountId: "",
        oneDrive: {
            vaultId: "",
            vaultName: "",
            deviceId: "",
            deviceName: "",
            autoMode: "manual",
            enabled: false,
            onStartup: false,
            afterSave: false,
            firstSyncCompleted: false,
            intervalMinutes: 0,
            maxFileSizeMb: 250,
            ignorePatterns: [
                "^\\.obsidian/",
                "^\\.trash/",
                "^\\.git/",
                "^\\.suite-recovery/",
                "^\\.suite-staging/",
                "(^|/)_suite_staging/",
                "(^|/)node_modules/",
            ],
            lastSuccessAt: 0,
            status: "idle",
            lastError: "",
            lastSkippedCount: 0,
            sharedStateHash: "",
            fileStates: {},
            runAllowed: false,
            confirmed: false,
        },
        externalBindings: [],
        externalRecoveryArchives: [],
        externalConflicts: [],
        recoveryEntries: [],
        legacyMigrationChecked: false,
        rolloutEnabled: true,
    },
};

export function newId(prefix: string): string {
    const random = typeof crypto !== "undefined" && "randomUUID" in crypto
        ? crypto.randomUUID()
        : Math.random().toString(36).slice(2) + Date.now().toString(36);
    return prefix + "-" + random;
}

export function mergeSettings(raw: Partial<SmartMDSettings> | undefined): SmartMDSettings {
    const source = raw || {};
    const shared = source.shared || ({} as SharedSmartMDState);
    const local = source.local || ({} as LocalSmartMDState);
    const oneDrive = local.oneDrive || ({} as OneDriveSyncSettings);
    return {
        schemaVersion: 2,
        shared: {
            ...DEFAULT_SETTINGS.shared,
            ...shared,
            publishSpaces: Array.isArray(shared.publishSpaces)
                ? shared.publishSpaces.map((space: any) => ({
                    ...space,
                    documents: Array.isArray(space.documents) ? space.documents : [],
                    pendingDocuments: Array.isArray(space.pendingDocuments) ? space.pendingDocuments : [],
                    pendingDeletes: Array.isArray(space.pendingDeletes) ? space.pendingDeletes : [],
                    cloudVersion: Number(space.cloudVersion || 0),
                    reconciliationStatus: space.reconciliationStatus || "pending_confirmation",
                    operationStatus: space.operationStatus || "needs_confirmation",
                    statusMessage: String(space.statusMessage || ""),
                    scopeKnown: space.scopeKnown !== false,
                }))
                : [],
            publishTombstones: (shared as any).publishTombstones && typeof (shared as any).publishTombstones === "object"
                ? { ...(shared as any).publishTombstones }
                : {},
            publishServerUrl: SMARTKEEP_PUBLISH_URL,
        },
        local: {
            ...DEFAULT_SETTINGS.local,
            ...local,
            interfaceLanguage: local.interfaceLanguage === "zh-CN" ? "zh-CN" : "en",
            smartkeepAccountName: String((local as any).smartkeepAccountName || ""),
            smartkeepAccountId: String((local as any).smartkeepAccountId || ""),
            oneDrive: {
                ...DEFAULT_SETTINGS.local.oneDrive,
                ...oneDrive,
                vaultName: String((oneDrive as any).vaultName || ""),
                autoMode: ["manual", "smart", "custom"].includes(String((oneDrive as any).autoMode))
                    ? (oneDrive as any).autoMode
                    : Boolean((oneDrive as any).enabled)
                        && Boolean((oneDrive as any).onStartup ?? (oneDrive as any).syncOnStartup)
                        && Boolean((oneDrive as any).afterSave ?? (oneDrive as any).syncOnSave)
                        && [1, 5].includes(Number((oneDrive as any).intervalMinutes || 0))
                        ? "smart"
                        : Boolean((oneDrive as any).enabled) ? "custom" : "manual",
                enabled: Boolean((oneDrive as any).enabled),
                onStartup: Boolean((oneDrive as any).onStartup ?? (oneDrive as any).syncOnStartup),
                afterSave: Boolean((oneDrive as any).afterSave ?? (oneDrive as any).syncOnSave),
                firstSyncCompleted: Boolean((oneDrive as any).firstSyncCompleted ?? (oneDrive as any).initialized),
                maxFileSizeMb: Number((oneDrive as any).maxFileSizeMb || 250) === 10 ? 250 : Number((oneDrive as any).maxFileSizeMb || 250),
                ignorePatterns: Array.from(new Set([
                    ...(Array.isArray(oneDrive.ignorePatterns)
                        ? oneDrive.ignorePatterns
                        : DEFAULT_SETTINGS.local.oneDrive.ignorePatterns),
                    "^\\.suite-recovery/",
                    "^\\.suite-staging/",
                    "(^|/)_suite_staging/",
                ])),
                status: String((oneDrive as any).status || (oneDrive as any).lastStatus || "idle"),
                lastError: String((oneDrive as any).lastError || ""),
                lastSkippedCount: Number((oneDrive as any).lastSkippedCount || 0),
                sharedStateHash: String((oneDrive as any).sharedStateHash || ""),
                fileStates: (oneDrive as any).fileStates || (oneDrive as any).fileState || {},
                confirmed: (oneDrive as any).confirmed === true,
                runAllowed: (oneDrive as any).confirmed === true && (oneDrive as any).runAllowed === true,
            },
            externalBindings: Array.isArray(local.externalBindings)
                ? local.externalBindings.map((binding: any) => {
                    // Persisted legacy binding state is never a new write permission.
                    const scope = `${binding.externalPath}|${binding.targetPath}`;
                    const trusted = binding.confirmed === true
                        && !!binding.externalPath && !!binding.targetPath
                        && binding.confirmationScope === scope
                        && typeof binding.canonicalExternalPath === "string"
                        && !!binding.canonicalExternalPath
                        && Number.isFinite(binding.rootDev)
                        && Number.isFinite(binding.rootIno);
                    const runAllowed = trusted && binding.runAllowed === true;
                    return {
                        ...binding,
                        confirmed: trusted,
                        runAllowed,
                        status: runAllowed ? binding.status
                            : (trusted && ["offline", "failed"].includes(binding.status) ? binding.status : "paused"),
                    };
                })
                : [],
            externalRecoveryArchives: Array.isArray((local as any).externalRecoveryArchives)
                ? (local as any).externalRecoveryArchives
                : [],
            externalConflicts: Array.isArray(local.externalConflicts) ? local.externalConflicts : [],
            recoveryEntries: Array.isArray(local.recoveryEntries) ? local.recoveryEntries : [],
        },
    };
}

export function sharedStateForSync(input: any): SharedSmartMDState {
    // OWNER-POLICY-P1-DEC-07: a strict four-key subset. Omit absent/invalid
    // fields; never silently replace private local settings with empty values.
    const source = input && typeof input === "object" && !Array.isArray(input)
        ? input : {};
    const result: Record<string, unknown> = {};
    if (typeof source.updatedAt === "number" && Number.isSafeInteger(source.updatedAt)
        && source.updatedAt >= 0) {
        result.updatedAt = source.updatedAt;
    }
    const approvedUrl = "https://api.openai.com/v1/chat/completions";
    if (source.aiApiUrl === approvedUrl && source.aiModel === "gpt-4o") {
        result.aiApiUrl = approvedUrl;
        result.aiModel = "gpt-4o";
    }
    if (source.publishServerUrl === SMARTKEEP_PUBLISH_URL) {
        result.publishServerUrl = SMARTKEEP_PUBLISH_URL;
    }
    // Existing callers use SharedSmartMDState; actual OneDrive wire data is
    // intentionally a sparse subset and never contains a fifth field.
    return result as unknown as SharedSmartMDState;
}


export type ExternalFilterMode = "exclude" | "include";
export type FilterSyntax = "glob" | "exact" | "regex";
export type FilterTarget = "folder" | "file" | "all";

export interface FilterRuleItem {
    id: string;
    pattern: string;
    target: FilterTarget;
    syntax: FilterSyntax;
    enabled: boolean;
}

export interface ExternalFilterConfig {
    mode: ExternalFilterMode;
    builtinPresets: {
        ignoreBuildArtifacts: boolean;
        ignoreIDEConfig: boolean;
    };
    userRules: FilterRuleItem[];
    filterHash: string;
}
