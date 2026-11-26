import { localize, localizeMessage, setInterfaceLanguage, type InterfaceLanguage } from "./localization";
import {
    App,
    type Command,
    type Editor,
    type MarkdownFileInfo,
    MarkdownView,
    Menu,
    Modal,
    Notice,
    Platform,
    Plugin,
    PluginSettingTab,
    Setting,
    setTooltip,
    TFile,
    TFolder,
    WorkspaceLeaf,
} from "obsidian";
import { AIEditorView, AI_VIEW_TYPE, type ChatQuote } from "./ai-editor";
import { AIChatPreview } from "./ai-chat-preview";
import { modelSettings, testModelConnection } from "./ai-model";
import { getDevicePlatformName, MicrosoftAuth, SmartMDStore } from "./core";
import { MicrosoftDeviceCodeModal } from "./microsoft-device-modal";
import { ExternalSyncService } from "./external-sync";
import { OneDriveSyncService } from "./onedrive";
import { safeOneDriveError } from "./onedrive-error-redaction";
import { getAndroidVaultRoot } from "./onedrive-android-vault";
import { PublishService } from "./publish";
import { SmartKeepAuth } from "./smartkeep-auth";
import { SyncFailureDiagnostic, SMARTKEEP_PUBLISH_URL } from "./types";


export class SmartMDRuntime {
    private readonly aiPreview = new AIChatPreview();
    readonly app: App;
    plugin!: Plugin;
    store!: SmartMDStore;
    auth!: MicrosoftAuth;
    smartkeepAuth!: SmartKeepAuth;
    publish!: PublishService;
    oneDrive!: OneDriveSyncService;
    external: ExternalSyncService | null = null;
    statusBarItem: HTMLElement | null = null;
    private microsoftAccountJob: Promise<void> | null = null;
    private smartkeepAccountJob: Promise<void> | null = null;
    private readonly commandIds = new Set<string>();
    private readonly ribbonLabels = new Map<HTMLElement, string>();
    private statusBarText = "";

    settingsTab: SmartMDSettingsTab | null = null;
    private feedbackDots = new Set<HTMLElement>();
    private feedbackIndicatorsActive = false;
    private feedbackObserver: MutationObserver | null = null;
    private feedbackDotTimer: number | null = null;

    get supportsOneDrive(): boolean {
        if (Platform.isAndroidApp) {
            return Boolean(getAndroidVaultRoot(this.app.vault.adapter));
        }
        return Platform.isDesktopApp && Platform.isWin;
    }

    get supportsExternalFolderSync(): boolean {
        return Platform.isDesktopApp && Platform.isWin;
    }

    get isDesktop(): boolean {
        return Platform.isDesktopApp;
    }

    constructor(pluginOrApp: Plugin | App, plugin?: Plugin, statusBarItem?: HTMLElement | null) {
        this.plugin = plugin || pluginOrApp as Plugin;
        this.app = plugin ? pluginOrApp as App : this.plugin.app;
        this.statusBarItem = statusBarItem || null;
    }

    updateStatusBar(text: string): void {
        this.statusBarText = text;
        this.updateStatusBarLanguage();
    }

    private updateStatusBarLanguage(): void {
        if (!this.statusBarItem) return;
        this.statusBarItem.setText(this.statusBarText ? localizeMessage(this.statusBarText) : localize("SmartMD: 就绪"));
        this.statusBarItem.title = localize("点击打开 SmartMD 设置");
    }

    async changeInterfaceLanguage(language: InterfaceLanguage): Promise<void> {
        const previous = this.store.settings.local.interfaceLanguage;
        if (language === previous) return;
        this.store.settings.local.interfaceLanguage = language;
        try {
            await this.store.save();
        } catch (error) {
            this.store.settings.local.interfaceLanguage = previous;
            throw error;
        }
        setInterfaceLanguage(language);
        this.updateStatusBarLanguage();
        this.registerCommands();
        for (const leaf of this.app.workspace.getLeavesOfType(AI_VIEW_TYPE)) {
            if (leaf.view instanceof AIEditorView) leaf.view.refreshLanguage();
        }
        for (const [icon, label] of this.ribbonLabels) setTooltip(icon, localize(label), { placement: "right" });
    }

    setSettingsTab(tab: SmartMDSettingsTab): void {
        this.settingsTab = tab;
    }

    async syncOneDriveFromShortcut(): Promise<void> {
        if (!this.oneDrive.settings().firstSyncCompleted) {
            this.openSettingsTab("onedrive");
            return;
        }
        await this.oneDrive.manualSync();
    }

    openSettingsTab(page?: "home" | "publish" | "onedrive" | "external" | "ai" | "about"): void {
        const doOpen = () => {
            try {
                const setting = (this.app as any).setting;
                if (setting) {
                    setting.open();
                    if (setting.openTabById) {
                        setting.openTabById(this.plugin.manifest.id);
                    }
                }
                if (page && this.settingsTab) {
                    this.settingsTab.setActivePage(page);
                }
            } catch (e) {
                console.error("Failed to open settings tab", e);
            }
        };
        doOpen();
        window.setTimeout(doOpen, 100);
    }

    private async prepareDeviceIdentity(): Promise<void> {
        if (typeof window === "undefined") return;
        const settings = (this.store as any).settings.local.oneDrive;
        const markerKey = "smartkeep-connect:onedrive-device-id-v2";
        let deviceId = "";
        try {
            deviceId = window.localStorage.getItem(markerKey) || "";
            if (!deviceId) {
                const random = typeof crypto !== "undefined" && "randomUUID" in crypto
                    ? crypto.randomUUID()
                    : Math.random().toString(36).slice(2) + Date.now().toString(36);
                deviceId = `device-${random}`;
                window.localStorage.setItem(markerKey, deviceId);
            }
        } catch {
            deviceId = String(settings.deviceId || "");
        }
        if (!deviceId) return;

        const inheritedHistory = Boolean(settings.deviceId && settings.deviceId !== deviceId);
        const pollutedHistory = Object.keys(settings.fileStates || {}).some((path) => /^\.suite-staging\//i.test(path));
        settings.deviceId = deviceId;
        if (!settings.deviceName) {
            settings.deviceName = getDevicePlatformName();
        }
        if (inheritedHistory || pollutedHistory) {
            // Keep old sync baselines and pending actions as read-only evidence.
            // Never promote a copied/inherited device's history into authority:
            // an explicit new approval must precede any remote/local write.
            settings.confirmed = false;
            settings.runAllowed = false;
            settings.enabled = false;
            settings.firstSyncCompleted = false;
            settings.lastError = "设备身份或历史记录发生变化；执行许可已撤销，旧基线和未决项目仅供人工重新核对。";
            settings.status = "paused";
            try {
                window.localStorage.removeItem(`smartkeep-connect:onedrive-delta:${settings.vaultId}`);
            } catch {}
        }
        await (this.store as any).save();
    }

    async load(): Promise<void> {
        // 双插件运行时互斥哨兵：防止 smartkeep-connect 与 smartmd 同时激活引发脑裂
        const activePlugins = (this.app as any).plugins?.plugins;
        if (activePlugins && activePlugins["smartkeep-connect"] && this.plugin?.manifest?.id === "smartmd") {
            new Notice(localize("【安全警告】检测到旧版 smartkeep-connect 插件仍处于激活状态！为防止数据脑裂冲突，SmartMD 已暂停加载。请在第三方插件列表中禁用旧插件。"), 0);
            console.error("FATAL: Both smartmd and smartkeep-connect are active. Split-brain prevention triggered.");
            return;
        }

        const Store: any = SmartMDStore;
        this.store = Store.length >= 2 ? new Store(this.app, this.plugin) : new Store(this.plugin);
        await (this.store as any).load();
        setInterfaceLanguage(this.store.settings.local.interfaceLanguage);
        this.updateStatusBarLanguage();
        if (this.supportsOneDrive) await this.prepareDeviceIdentity();
        const Auth: any = MicrosoftAuth;
        this.auth = new Auth(this.store);
        // Keep the existing Microsoft Entra redirect; also register the new SmartMD alias.
        // Do not switch the outbound OAuth URI without updating the Entra app registration.
        const onOneDriveAuth = (params: Record<string, string>) => {
            void this.auth.completeSignIn(params).then(() => {
                this.openSettingsTab("onedrive");
            });
        };
        if (this.supportsOneDrive) this.plugin.registerObsidianProtocolHandler("smartkeep-connect-onedrive", onOneDriveAuth);
        if (this.supportsOneDrive) this.plugin.registerObsidianProtocolHandler("smartmd-onedrive", onOneDriveAuth);
        this.smartkeepAuth = new SmartKeepAuth(this.store);
        // SmartKeep's identity provider has not yet allowlisted the SmartMD callback.
        const onPublishingAuth = (params: Record<string, string>) => {
            void this.smartkeepAuth.completeSignIn(params).then(() => {
                this.openSettingsTab("publish");
            });
        };
        this.plugin.registerObsidianProtocolHandler("smartkeep-connect-auth", onPublishingAuth);
        this.plugin.registerObsidianProtocolHandler("smartmd-auth", onPublishingAuth);
        this.publish = new PublishService(this.app, this.store, this.smartkeepAuth);
        this.smartkeepAuth.setPublishLifecycle(
            () => this.publish.stopForSignOut(),
            () => this.publish.resumeAfterSignIn(),
        );
        this.oneDrive = new OneDriveSyncService(this.app, this.store, this.auth);
        this.oneDrive.setStatusCallback((text: string) => this.updateStatusBar(text));
        if (this.supportsExternalFolderSync) {
            const external = new ExternalSyncService(this.app, this.store);
            this.external = external;
            if (this.supportsOneDrive) {
                this.oneDrive.setDeleteGuard((path, readOnly) =>
                    external.protectRemoteDelete(path, readOnly));
            }
        }

        await this.applySafeMigration();
        this.plugin.registerEditorExtension(this.aiPreview.extension);
        this.plugin.registerView(AI_VIEW_TYPE, (leaf: WorkspaceLeaf) => new AIEditorView(leaf, this.store, this, this.aiPreview));
        this.registerCommands();
        this.registerRibbonIcons();
        this.registerMenus();
        this.registerVaultFlows();
        this.registerSchedulers();
        const publishStartup = (this.smartkeepAuth.isSignedIn() ? this.publish.resumeAfterSignIn() : Promise.resolve()).catch(() => {
            console.warn("[SmartMD] 发布服务启动对账未完成");
        });
        this.app.workspace.onLayoutReady(() => {
            void publishStartup.then(() => this.publish.checkContinuousSync());
        });
        this.plugin.registerInterval(window.setInterval(() => this.publish.checkContinuousSync(), 30000));
        this.registerFeedbackIndicators();
        this.plugin.registerInterval(window.setInterval(() => void this.publish.refreshFeedbackSummary(), 30000));
    }

    unload(): void {
        this.auth?.cancelSignIn();
        this.feedbackIndicatorsActive = false;
        this.feedbackObserver?.disconnect();
        if (this.feedbackDotTimer !== null) window.clearTimeout(this.feedbackDotTimer);
        this.publish.setFeedbackChangeCallback(null);
        for (const dot of this.feedbackDots) dot.remove();
        this.feedbackDots.clear();
        document.querySelectorAll(".smartmd-feedback-dot").forEach(dot => dot.remove());
        this.publish.dispose();
        this.oneDrive.dispose();
        this.external?.dispose();
    }

    private registerCommands(): void {
        for (const id of this.commandIds) this.plugin.removeCommand(id);
        this.commandIds.clear();
        const addCommand = (command: Command) => {
            this.commandIds.add(command.id);
            this.plugin.addCommand(command);
        };
        addCommand({
            id: "suite-open-ai-editor",
            name: localize("打开 AI 编辑器"),
            callback: () => void this.openAI(),
        });
        addCommand({
            id: "suite-quote-selection-in-ai-chat",
            name: localize("用AI编辑选中的文字"),
            editorCallback: (editor, info) => {
                const view = this.findMarkdownEditor(editor, info);
                void this.quoteInAI(view ? AIEditorView.captureQuote(view) : null);
            },
        });
        addCommand({
            id: "suite-publish-current-document",
            name: localize("发布当前文档"),
            checkCallback: (checking: boolean) => {
                const file = this.app.workspace.getActiveFile();
                if (!file || file.extension.toLowerCase() !== "md") return false;
                if (!checking) void this.publish.publishFile(file);
                return true;
            },
        });
        addCommand({
            id: "suite-open-publish-manager",
            name: localize("打开云端发布管理"),
            callback: () => this.publish.openManager(),
        });
        if (this.supportsOneDrive) {
            addCommand({
                id: "suite-onedrive-sync-now",
                name: localize("立即同步到 OneDrive"),
                callback: () => void this.syncOneDriveFromShortcut(),
            });
        }
        if (this.supportsExternalFolderSync) {
            addCommand({
                id: "suite-open-external-sync",
                name: localize("管理外部文件夹同步"),
                callback: () => this.external?.openManager(),
            });
        }
    }

    private registerRibbonIcons(): void {
        const addIcon = (icon: string, label: string, action: () => void) => {
            const element = this.plugin.addRibbonIcon(icon, localize(label), action);
            this.ribbonLabels.set(element, label);
        };
        addIcon("cloud", "SmartMD 发布管理", () => this.publish.openManager());
        if (this.supportsOneDrive) addIcon("refresh-cw", "SmartMD OneDrive 同步", () => void this.syncOneDriveFromShortcut());
        addIcon("wand-sparkles", "SmartMD AI 编辑器", () => void this.openAI());
    }

    private registerMenus(): void {
        this.plugin.registerEvent(this.app.workspace.on("editor-menu", (menu, editor, info) => {
            const view = this.findMarkdownEditor(editor, info);
            if (!view || view.getMode() !== "source") return;
            if (editor.somethingSelected()) {
                const quote = AIEditorView.captureQuote(view);
                menu.addItem(item => item.setTitle(localize("用AI编辑选中的文字")).setIcon("text-cursor")
                    .onClick(() => void this.quoteInAI(quote)));
            }
            const document = AIEditorView.captureQuote(view, "document");
            if (document) menu.addItem(item => item.setTitle(localize("用AI编辑此文档")).setIcon("file-text")
                .onClick(() => void this.quoteInAI(document)));
        }));
        this.plugin.registerEvent(this.app.workspace.on("file-menu", (menu: Menu, file) => {
            if ((file instanceof TFile || file instanceof TFolder) && this.publish.publishedLink(file.path)) {
                this.suppressDuplicatePublishItems(menu);
                menu.addItem((item) => item
                    .setTitle(localize("查看发布链接"))
                    .setIcon("external-link")
                    .onClick(() => this.publish.openPublishedLink(file)));
                if (file instanceof TFile && file.extension.toLowerCase() === "md" && this.publish.feedbackForPath(file.path)) {
                    menu.addItem(item => item.setTitle(localize("处理问题反馈")).setIcon("message-circle")
                        .onClick(() => void this.publish.openDocumentFeedback(file)));
                }
                return;
            }
            if (file instanceof TFile && file.extension.toLowerCase() === "md") {
                menu.addItem((item) => item
                    .setTitle(localize("发布到 SmartMD"))
                    .setIcon("cloud-upload")
                    .onClick(() => void this.publish.publishFile(file)));
            }
            if (file instanceof TFolder) {
                menu.addItem((item) => item
                    .setTitle(localize("发布此文件夹"))
                    .setIcon("cloud-upload")
                    .onClick(() => void this.publish.publishFolder(file)));
            }
        }));
    }

    /** Scope compatibility to this one published object's menu, never global providers. */
    private suppressDuplicatePublishItems(menu: Menu): void {
        // Obsidian exposes no public removeItem API. Guard the actual 1.15 menu shape.
        const target = menu as any;
        if (!Array.isArray(target.items)) return;
        const nativeSource = (this.app as any).internalPlugins?.plugins?.publish?.instance?.onEnable?.toString() || "";
        const prune = () => {
            for (let index = target.items.length - 1; index >= 0; index--) {
                const item = target.items[index];
                const callback = typeof item.callback === "function" ? item.callback.toString() : "";
                const nativePublish = item.section === "action"
                    && Boolean(item.iconEl?.querySelector("svg.lucide-send"))
                    && /\.(?:uploadFile|openPublishChanges)\s*\(/.test(callback)
                    && nativeSource.includes(callback);
                const legacySmartMD = Boolean(item.iconEl?.querySelector("svg.lucide-cloud-upload"))
                    && /\.publish\.(?:publishFile|publishFolder)\s*\(/.test(callback);
                if (!nativePublish && !legacySmartMD) continue;
                target.items.splice(index, 1);
                item.dom?.remove();
            }
        };
        prune();
        // Later providers run synchronously too; filter before native/DOM menu presentation.
        const addItem = target.addItem;
        target.addItem = function (builder: any) {
            const result = addItem.call(this, builder);
            prune();
            return result;
        };
    }

    private renderFeedbackDots(): void {
        const rows = new Map<HTMLElement, string>();
        // Obsidian 1.15 keeps native paths on FileExplorer items, not DOM attributes.
        for (const leaf of this.app.workspace.getLeavesOfType("file-explorer")) {
            const items = (leaf.view as any).fileItems || {};
            for (const item of Object.values(items) as any[]) {
                if (item.file instanceof TFile && item.selfEl?.matches?.(".nav-file-title")) rows.set(item.selfEl, item.file.path);
            }
        }
        // Also support rows exposed with data-path by other explorer versions.
        for (const row of Array.from(document.querySelectorAll<HTMLElement>(".nav-file-title[data-path]"))) {
            if (!rows.has(row)) rows.set(row, row.getAttribute("data-path") || "");
        }
        const activeDots = new Set<HTMLElement>();
        for (const [row, path] of rows) {
            const file = this.app.vault.getAbstractFileByPath(path);
            const pending = file instanceof TFile && file.extension.toLowerCase() === "md" && Number(this.publish.feedbackForPath(path)?.open || 0) > 0;
            const existing = row.querySelector<HTMLElement>(".smartmd-feedback-dot");
            if (!pending) { existing?.remove(); continue; }
            if (existing) { activeDots.add(existing); continue; }
            const title = row.querySelector(".nav-file-title-content");
            if (!title) continue;
            const dot = row.ownerDocument.createElement("span");
            dot.className = "smartmd-feedback-dot";
            dot.setAttribute("role", "img");
            dot.setAttribute("aria-label", "有待处理的读者反馈");
            dot.title = localize("有待处理的读者反馈");
            title.parentElement?.insertBefore(dot, title);
            activeDots.add(dot);
        }
        // Includes detached rows in collapsed folders and secondary explorer leaves.
        for (const dot of this.feedbackDots) if (!activeDots.has(dot)) dot.remove();
        this.feedbackDots = activeDots;
    }

    private registerFeedbackIndicators(): void {
        this.feedbackIndicatorsActive = true;
        const schedule = () => {
            if (!this.feedbackIndicatorsActive) return;
            if (this.feedbackDotTimer !== null) window.clearTimeout(this.feedbackDotTimer);
            this.feedbackDotTimer = window.setTimeout(() => { this.feedbackDotTimer = null; if (this.feedbackIndicatorsActive) this.renderFeedbackDots(); }, 30);
        };
        this.publish.setFeedbackChangeCallback(() => this.renderFeedbackDots());
        this.plugin.registerEvent(this.app.workspace.on("layout-change", schedule));
        this.app.workspace.onLayoutReady(() => {
            if (!this.feedbackIndicatorsActive) return;
            this.feedbackObserver = new MutationObserver(records => {
                if (records.some(record => record.type === "attributes" || [...Array.from(record.addedNodes), ...Array.from(record.removedNodes)].some(node =>
                    node instanceof Element && (node.matches(".nav-file-title,.nav-files-container") || Boolean(node.querySelector(".nav-file-title")))))) schedule();
            });
            this.feedbackObserver.observe(document.body, { subtree: true, childList: true, attributes: true, attributeFilter: ["data-path"] });
            this.renderFeedbackDots();
        });
    }

    private registerVaultFlows(): void {
        this.plugin.registerEvent(this.app.vault.on("modify", (file) => {
            this.publish.onVaultChange(file);
            if (this.supportsOneDrive) this.oneDrive.scheduleAfterSave();
        }));
        this.plugin.registerEvent(this.app.vault.on("create", (file) => {
            this.publish.onVaultChange(file);
            if (this.supportsOneDrive) this.oneDrive.scheduleAfterSave();
        }));
        this.plugin.registerEvent(this.app.vault.on("rename", (file, oldPath) => {
            this.publish.onVaultRename(file, oldPath);
            if (this.supportsOneDrive) this.oneDrive.scheduleAfterSave();
        }));
        this.plugin.registerEvent(this.app.vault.on("delete", (file) => {
            this.publish.onVaultDelete(file);
            if (this.supportsOneDrive) this.oneDrive.scheduleAfterSave();
        }));
    }

    private registerSchedulers(): void {
        if (this.external) {
            this.plugin.registerInterval(window.setInterval(() => void this.external?.syncAll(), 3000));
        }
        if (!this.supportsOneDrive) return;
        this.plugin.registerInterval(window.setInterval(() => {
            const settings = this.oneDrive.settings();
            const intervalMinutes = Number(settings.intervalMinutes);
            if (!this.oneDrive.canRun() || !settings.enabled || !Number.isFinite(intervalMinutes) ||
                intervalMinutes <= 0 || !settings.firstSyncCompleted) return;
            const due = Date.now() - Number(settings.lastSuccessAt || 0) >= intervalMinutes * 60 * 1000;
            if (due) void this.oneDrive.requestSync();
        }, 10000));

        const settings = this.oneDrive.settings();
        if (this.oneDrive.canRun() && settings.enabled && settings.onStartup && settings.firstSyncCompleted) {
            window.setTimeout(() => void this.oneDrive.requestSync(), 2500);
        }
    }

    private async applySafeMigration(): Promise<void> {
        if (this.supportsOneDrive) await this.oneDrive.isolateUnconfirmed();
        await this.publish.pauseLegacySpaces();
        if (this.external) await this.external.markLegacyBindingsPaused();
        await (this.store as any).save();
    }

    async openAI(): Promise<void> {
        const existing = this.app.workspace.getLeavesOfType(AI_VIEW_TYPE)[0];
        const leaf = existing || this.app.workspace.getRightLeaf(false);
        if (!leaf) return;
        await leaf.setViewState({ type: AI_VIEW_TYPE, active: true });
        this.app.workspace.revealLeaf(leaf);
    }

    private findMarkdownEditor(editor: Editor, info: MarkdownFileInfo): MarkdownView | null {
        // editor-menu may supply MarkdownFileInfo instead of the MarkdownView itself.
        if (info instanceof MarkdownView && info.editor === editor) return info;
        for (const { view } of this.app.workspace.getLeavesOfType("markdown")) {
            if (view instanceof MarkdownView && view.editor === editor && view.file === info.file) {
                return view;
            }
        }
        return null;
    }

    private async quoteInAI(quote: ChatQuote | null): Promise<void> {
        if (!quote) {
            new Notice(localize("请选择同一段正文或同一组列表中的文字"));
            return;
        }
        await this.openAI();
        const chat = this.app.workspace.getLeavesOfType(AI_VIEW_TYPE)[0]?.view;
        if (chat instanceof AIEditorView) chat.useQuote(quote);
    }

    isMicrosoftSignedIn(): boolean {
        const auth: any = this.auth;
        return Boolean(auth?.isSignedIn?.());
    }

    isMicrosoftSignInPending(): boolean {
        return this.auth.isSignInPending();
    }

    cancelMicrosoftSignIn(): void {
        this.auth.cancelSignIn();
    }

    private async openMicrosoftLoginUrl(url: string): Promise<void> {
        const destination = new URL(url);
        if (destination.protocol !== "https:" || destination.hostname !== "login.microsoftonline.com") {
            throw new Error("微软登录地址无效，已拒绝打开浏览器。");
        }
        if (Platform.isDesktopApp && typeof (window as any).require === "function") {
            const shell = (window as any).require("electron")?.shell;
            if (typeof shell?.openExternal === "function") {
                await shell.openExternal(url);
                return;
            }
        }
        // Mobile hosts may not provide Electron's native browser launcher.
        window.open(url, "_blank");
    }

    async signIn(): Promise<void> {
        if (this.microsoftAccountJob) return await this.microsoftAccountJob;
        const job = (async () => {
            const auth: any = this.auth;
            const action = auth.signIn || auth.login;
            if (!action) throw new Error("当前版本无法启动微软登录");
            await action.call(auth, (url: string) => this.openMicrosoftLoginUrl(url));
            new Notice(localize("微软账号登录成功"));
        })();
        this.microsoftAccountJob = job;
        try {
            await job;
        } finally {
            if (this.microsoftAccountJob === job) this.microsoftAccountJob = null;
        }
    }

    async signInWithDeviceCode(): Promise<void> {
        if (this.microsoftAccountJob) return await this.microsoftAccountJob;
        const modal = new MicrosoftDeviceCodeModal(this.app, () => this.cancelMicrosoftSignIn());
        modal.open();
        const job = (async () => {
            try {
                await this.auth.signInWithDeviceCode((challenge) => modal.showChallenge(challenge));
                new Notice(localize("微软账号登录成功。首次 OneDrive 同步仍需单独确认。"));
            } finally {
                modal.dismiss();
            }
        })();
        this.microsoftAccountJob = job;
        try {
            await job;
        } finally {
            if (this.microsoftAccountJob === job) this.microsoftAccountJob = null;
        }
    }

    async signOut(): Promise<void> {
        if (this.microsoftAccountJob) return await this.microsoftAccountJob;
        const job = (async () => {
            await this.oneDrive.pause();
            const auth: any = this.auth;
            const action = auth.signOut || auth.logout;
            if (action) await action.call(auth);
            new Notice(localize("已停止同步并退出微软账号，本地文件未改变"));
        })();
        this.microsoftAccountJob = job;
        try {
            await job;
        } finally {
            if (this.microsoftAccountJob === job) this.microsoftAccountJob = null;
        }
    }

    isSmartKeepSignedIn(): boolean {
        return this.smartkeepAuth.isSignedIn();
    }

    async signInSmartKeep(): Promise<void> {
        if (this.smartkeepAccountJob) return await this.smartkeepAccountJob;
        const job = (async () => {
            await this.smartkeepAuth.signIn((url) => window.open(url, "_blank"));
            new Notice(localize("SmartKeep 账号登录成功"));
        })();
        this.smartkeepAccountJob = job;
        try {
            await job;
        } finally {
            if (this.smartkeepAccountJob === job) this.smartkeepAccountJob = null;
        }
    }

    async signOutSmartKeep(): Promise<void> {
        if (this.smartkeepAccountJob) return await this.smartkeepAccountJob;
        const job = (async () => {
            await this.smartkeepAuth.signOut();
            new Notice(localize("已退出 SmartKeep MD，本地文件和发布内容未改变"));
        })();
        this.smartkeepAccountJob = job;
        try {
            await job;
        } finally {
            if (this.smartkeepAccountJob === job) this.smartkeepAccountJob = null;
        }
    }
}

export class SmartMDSettingsTab extends PluginSettingTab {
    private activePage: "home" | "publish" | "onedrive" | "external" | "ai" | "about" = "home";
    private refreshFrame = 0;
    private stateSignature = "";
    private aiKeySave: Promise<{ address: string; saved: boolean } | null> = Promise.resolve(null);
    private aiRevision = 0;
    private aiTestRunning = false;
    private aiTestController?: AbortController;
    private aiTestTarget?: { aiApiUrl: string; aiModel: string };
    private aiTestMessage = "";
    private refreshAiTest?: () => void;
    private aiAddressShown?: string;
    private aiModelShown?: string;

    constructor(app: App, plugin: Plugin, private runtime: SmartMDRuntime) {
        super(app, plugin);
        this.runtime.oneDrive.setStatusListener(() => this.scheduleDisplay());
        const unsubscribe = this.runtime.store.subscribe(() => {
            const current = this.runtime.store.settings.shared;
            const target = this.aiTestTarget;
            const refreshAi = this.aiFormNeedsRefresh();
            if (refreshAi || (target && (target.aiApiUrl !== current.aiApiUrl || target.aiModel !== current.aiModel))) {
                this.aiSettingsChanged();
            }
            const next = this.currentStateSignature();
            if (next !== this.stateSignature || refreshAi) this.scheduleDisplay();
        });
        (plugin as any).register?.(unsubscribe);
    }

    setActivePage(page: "home" | "publish" | "onedrive" | "external" | "ai" | "about"): void {
        this.activePage = page;
        if (this.containerEl.isConnected) {
            this.display();
        }
    }

    display(): void {
        const { containerEl } = this;
        this.refreshAiTest = undefined;
        containerEl.empty();
        containerEl.addClass("smartmd-settings");
        this.renderHeader(containerEl);
        this.renderNavigation(containerEl);

        const content = containerEl.createDiv({ cls: "smartmd-settings-content" });
        if (this.activePage === "home") this.homeSection(content);
        else if (this.activePage === "publish") this.publishSection(content);
        else if (this.activePage === "onedrive" && this.runtime.supportsOneDrive) this.oneDriveSection(content);
        else if (this.activePage === "external" && this.runtime.supportsExternalFolderSync) this.externalSection(content);
        else if (this.activePage === "ai") this.aiSection(content);
        else if (this.activePage === "about") this.aboutSection(content);
        else {
            this.activePage = "home";
            this.homeSection(content);
        }
        this.stateSignature = this.currentStateSignature();
    }

    private currentStateSignature(): string {
        const local = (this.runtime.store as any).settings.local;
        return JSON.stringify({
            smartkeep: [local.smartkeepAccountId, this.runtime.isSmartKeepSignedIn()],
            microsoft: [local.microsoftAccountId, this.runtime.isMicrosoftSignedIn(), this.runtime.isMicrosoftSignInPending()],
            oneDrive: [
                local.oneDrive.status,
                local.oneDrive.lastSuccessAt,
                local.oneDrive.lastError,
                local.oneDrive.firstSyncCompleted,
                local.oneDrive.enabled,
                local.oneDrive.onStartup,
                local.oneDrive.afterSave,
                local.oneDrive.intervalMinutes,
                this.runtime.oneDrive.getPendingOneDriveIncomingSharedSettings()?.expiresAt || 0,
                this.runtime.oneDrive.isBusy(),
            ],
            publish: this.runtime.publish.spaces().map((space) => [
                space.id,
                space.operationStatus,
                space.reconciliationStatus,
                space.statusMessage,
                space.cloudVersion,
            ]),
            external: [
                ...(this.runtime.external?.bindings() || []).map((binding: any) => [
                    binding.id,
                    binding.status,
                    binding.pendingCount,
                    binding.lastError,
                    (binding.conflicts || []).length,
                    (binding.recoveryItems || []).length,
                ]),
                ...(this.runtime.external?.recoveryArchives() || []).map((archive: any) => [
                    archive.id,
                    (archive.recoveryItems || []).length,
                    archive.lastError,
                ]),
            ],
        });
    }

    private scheduleDisplay(): void {
        if (!this.containerEl.isConnected || this.refreshFrame || (this.activePage === "ai" && !this.aiFormNeedsRefresh())) return;
        this.refreshFrame = window.requestAnimationFrame(() => {
            this.refreshFrame = 0;
            if (!this.containerEl.isConnected) return;
            if (this.activePage === "ai" && !this.aiFormNeedsRefresh()) return;
            this.display();
        });
    }

    private aiFormNeedsRefresh(): boolean {
        const current = this.runtime.store.settings.shared;
        return this.activePage === "ai" && (this.aiAddressShown !== current.aiApiUrl || this.aiModelShown !== current.aiModel);
    }

    private renderHeader(container: HTMLElement): void {
        const header = container.createDiv({ cls: "smartmd-settings-header" });
        header.createEl("p", { text: localize("本地知识库的发布、同步与 AI 编辑中心。每项能力独立运行，互不触发删除或回滚。") });
    }

    private renderNavigation(container: HTMLElement): void {
        const tabs: Array<{ id: typeof this.activePage; label: string }> = [
            { id: "home", label: localize("首页") },
            { id: "publish", label: localize("云端发布") },
        ];
        if (this.runtime.supportsOneDrive) {
            tabs.push({ id: "onedrive", label: "OneDrive" });
        }
        if (this.runtime.supportsExternalFolderSync) {
            tabs.push({ id: "external", label: localize("外部文件夹") });
        }
        tabs.push({ id: "ai", label: localize("AI 编辑") });
        tabs.push({ id: "about", label: localize("关于") });

        const navigation = container.createDiv({ cls: "smartmd-settings-tabs" });
        navigation.setAttr("role", "tablist");
        for (const tab of tabs) {
            const button = navigation.createEl("button", {
                cls: "smartmd-settings-tab",
                text: tab.label,
                attr: { type: "button", role: "tab", "aria-selected": String(this.activePage === tab.id) },
            });
            button.classList.toggle("is-active", this.activePage === tab.id);
            button.addEventListener("click", () => {
                this.activePage = tab.id;
                this.display();
                if (tab.id === "publish" && this.runtime.isSmartKeepSignedIn()) {
                    void this.runtime.publish.refreshSpaces().then(() => {
                        if (this.activePage === "publish" && this.containerEl.isConnected) this.display();
                    });
                }
            });
        }
    }

    private homeSection(container: HTMLElement): void {
        new Setting(container)
            .setName(localize("语言"))
            .addDropdown(dropdown => dropdown
                .addOption("en", "English")
                .addOption("zh-CN", "简体中文")
                .setValue(this.runtime.store.settings.local.interfaceLanguage)
                .onChange(async value => {
                    if (value !== "en" && value !== "zh-CN") return;
                    dropdown.setDisabled(true);
                    try {
                        await this.runtime.changeInterfaceLanguage(value);
                        this.display();
                    } catch {
                        dropdown.setValue(this.runtime.store.settings.local.interfaceLanguage);
                        dropdown.setDisabled(false);
                        new Notice(localize("无法保存语言设置。"));
                    }
                }));
        const shared = (this.runtime.store as any).settings.shared;
        const grid = container.createDiv({ cls: "smartmd-feature-grid" });
        const cards: Array<{ page: typeof this.activePage; title: string; description: string; status: string }> = [
            {
                page: "publish",
                title: localize("云端发布"),
                description: localize("发布 Markdown、文件夹或知识库，并管理公开访问与问题反馈。"),
                status: this.runtime.publish.accountSpaceCount() === null ? localize("账号空间待查询") : localize("账号空间 {0}/3（所有知识库共用）", this.runtime.publish.accountSpaceCount()),
            },
            {
                page: "ai",
                title: localize("基础 AI 编辑器"),
                description: localize("使用自备模型生成修改建议，确认差异后再写入文档。"),
                status: shared.aiApiUrl && shared.aiModel ? localize("模型信息已配置") : localize("等待配置模型"),
            },
        ];
        if (this.runtime.supportsOneDrive) {
            cards.splice(cards.length - 1, 0, {
                page: "onedrive",
                title: localize("OneDrive 多设备同步"),
                description: localize("在插件独立空间同步知识库文件和非敏感插件状态。"),
                status: localizeMessage(this.runtime.oneDrive.statusText()),
            });
        }
        if (this.runtime.supportsExternalFolderSync) {
            cards.splice(cards.length - 1, 0, {
                page: "external",
                title: localize("外部文件夹同步"),
                description: localize("扫描外部文件夹及子目录中的全部 Markdown，保留原目录结构并在知识库内建立副本；任一端修改都会实时同步。"),
                status: localize("{0} 个同步项", this.runtime.external?.bindings().length || 0),
            });
        }
        cards.push({
            page: "about",
            title: localize("关于 SmartMD"),
            description: localize("了解项目、SmartKeep 主站、联系方式和未来更新计划。"),
            status: localize("公益开源方向"),
        });

        for (const card of cards) {
            const button = grid.createEl("button", {
                cls: "smartmd-feature-card",
                attr: { type: "button" },
            });
            button.createDiv({ cls: "smartmd-feature-status", text: card.status });
            button.createDiv({ cls: "smartmd-feature-title", text: card.title });
            button.createEl("p", { text: card.description });
            button.createDiv({ cls: "smartmd-home-link", text: localize("进入设置") });
            button.addEventListener("click", () => {
                this.activePage = card.page;
                this.display();
            });
        }
    }

    private microsoftAccountSetting(container: HTMLElement): void {
        const settings = (this.runtime.store as any).settings.local;
        const signedIn = this.runtime.isMicrosoftSignedIn();
        const pending = this.runtime.isMicrosoftSignInPending();
        const needsLocalLogin = Boolean(settings.microsoftAccountId && !signedIn);
        new Setting(container)
            .setName(signedIn ? localize("已登录：{0}", settings.microsoftAccountName) : pending ? localize("正在等待微软授权") : needsLocalLogin ? localize("此设备需要重新登录") : localize("尚未登录"))
            .setDesc(
                signedIn
                    ? localize("OneDrive 已获得当前设备授权。该授权只用于 OneDrive 多设备同步。")
                    : pending
                        ? localize("请在微软官网完成授权；如果无法完成，可取消登录后重新尝试。")
                    : needsLocalLogin
                        ? localize("账号信息来自其他设备，授权不会随知识库复制。请在当前设备重新登录。")
                        : localize("登录后才能使用 OneDrive。云端发布使用独立的 SmartKeep 账号。"),
            )
            .addButton((button) => button
                .setButtonText(signedIn ? localize("退出") : pending ? localize("取消登录") : localize("设备码登录"))
                .setCta()
                .onClick(async () => {
                    const wasSignedIn = signedIn;
                    const wasPending = pending;
                    button.setDisabled(true).setButtonText(wasSignedIn ? localize("正在停止…") : wasPending ? localize("正在取消…") : localize("登录中…"));
                    try {
                        if (wasSignedIn) await this.runtime.signOut();
                        else if (wasPending) this.runtime.cancelMicrosoftSignIn();
                        else {
                            // Expose a live cancel action while the browser awaits OAuth.
                            const attempt = this.runtime.signInWithDeviceCode();
                            this.display();
                            await attempt;
                        }
                    } catch (error: any) {
                        new Notice(localize("{0}失败：{1}", wasSignedIn ? localize("退出") : localize("登录"), localizeMessage(String(error.message || error))), 8000);
                    }
                    try {
                        this.display();
                    } catch (error) {
                        console.error("SmartMD settings refresh failed", error);
                        new Notice(localize("账号状态已更新，请重新打开插件设置。"), 8000);
                    }
                }));
    }

    private browserLoginFallbackSetting(container: HTMLElement): void {
        new Setting(container)
            .setName(localize("浏览器授权（备用）"))
            .setDesc(localize("仅在设备码登录不可用时使用，需要浏览器返回 Obsidian。"))
            .addButton((button) => button
                .setButtonText(localize("浏览器登录"))
                .onClick(async () => {
                    const attempt = this.runtime.signIn();
                    this.display();
                    try {
                        await attempt;
                    } catch (error: any) {
                        if (!String(error?.message || error).includes("登录已取消")) {
                            new Notice(localize("微软浏览器授权失败：") + localizeMessage(safeOneDriveError(error)), 8000);
                        }
                    } finally {
                        this.display();
                    }
                }));
    }

    private publishSection(container: HTMLElement): void {
        const local = (this.runtime.store as any).settings.local;
        const shared = (this.runtime.store as any).settings.shared;
        const signedIn = this.runtime.isSmartKeepSignedIn();
        new Setting(container).setClass("smartmd-settings-section-title").setName(localize("云端发布")).setHeading();
        container.createEl("p", { text: localize("SmartKeep MD 只保留最新版 Markdown 和问题反馈，不保存本地图片或正文历史版本。公开链接会自动保存到 SmartKeep，密码内容不会生成书签。") });
        new Setting(container)
            .setName(signedIn ? localize("发布账号：{0}", local.smartkeepAccountName) : localize("发布账号：尚未登录"))
            .setDesc(signedIn
                ? localize("该 SmartKeep 账号用于确认发布内容所有者、管理问题反馈和保存公开发布链接。")
                : localize("登录 SmartKeep 后才能发布；OneDrive 和自备模型 AI 不受影响。"))
            .addButton((button) => button
                .setButtonText(signedIn ? localize("退出") : localize("登录 SmartKeep"))
                .setCta()
                .onClick(async () => {
                    try {
                        if (signedIn) {
                            button.setButtonText(localize("正在停止")).setDisabled(true);
                            await this.runtime.signOutSmartKeep();
                        }
                        else await this.runtime.signInSmartKeep();
                        this.display();
                    } catch (error: any) {
                        new Notice(localize("{0}失败：{1}", signedIn ? localize("退出") : localize("登录"), localizeMessage(String(error.message || error))), 8000);
                    }
                }));
        new Setting(container)
            .setName(localize("发布服务地址"))
            .setDesc(localize("正式发布入口固定为 {0}，普通设置不可修改。", SMARTKEEP_PUBLISH_URL));
        const spaceSetting = new Setting(container)
            .setName(localize("发布空间"))
            .setDesc(this.runtime.publish.accountSpaceCount() === null
                ? localize("账号空间：待查询。所有知识库共用最多 3 个名额，可点管理查询。")
                : localize("账号空间：{0}/3（所有知识库共用）。新文档默认等待确认。", this.runtime.publish.accountSpaceCount()));
        const actions = spaceSetting.infoEl.createDiv({ cls: "smartmd-setting-actions" });
        const pubCurrentBtn = actions.createEl("button", {
            text: localize("发布当前文档"),
            attr: { type: "button" },
        });
        pubCurrentBtn.disabled = this.runtime.publish.hasBlockingFailure();
        pubCurrentBtn.addEventListener("click", () => {
            const file = this.app.workspace.getActiveFile();
            if (file?.extension.toLowerCase() === "md") void this.runtime.publish.publishFile(file);
            else new Notice(localize("请先打开一篇 Markdown 文档"));
        });
        const pubVaultBtn = actions.createEl("button", {
            text: localize("发布整个知识库"),
            attr: { type: "button" },
        });
        pubVaultBtn.disabled = this.runtime.publish.hasBlockingFailure();
        pubVaultBtn.addEventListener("click", () => void this.runtime.publish.publishVault());
        const reconcileBtn = actions.createEl("button", {
            text: localize("重新对账"),
            attr: { type: "button" },
        });
        reconcileBtn.addEventListener("click", async () => {
            reconcileBtn.disabled = true;
            reconcileBtn.setText(localize("正在对账..."));
            await this.runtime.publish.refreshSpaces();
            this.display();
        });
        spaceSetting.addButton((button) => button.setButtonText(localize("管理")).setCta().onClick(() => this.runtime.publish.openManager()));
    }

    private oneDriveSection(container: HTMLElement): void {
        const settings = this.runtime.oneDrive.settings();
        const autoMode = settings.enabled ? settings.autoMode || "custom" : "manual";
        const intervalMinutes = Number(settings.intervalMinutes || 0);
        const nextCheckAt = settings.enabled && intervalMinutes > 0 && settings.lastSuccessAt
            ? Number(settings.lastSuccessAt) + intervalMinutes * 60 * 1000
            : 0;
        let statusDescription = nextCheckAt
            ? localize("{0}；下次检查：{1}", localizeMessage(this.runtime.oneDrive.statusText()), new Date(Math.max(Date.now(), nextCheckAt)).toLocaleString())
            : localizeMessage(this.runtime.oneDrive.statusText());
        if (!settings.firstSyncCompleted && !this.runtime.oneDrive.isBusy()) {
            statusDescription += " " + localize("第一次使用时，OneDrive 可能会创建 SmartMD 文件夹。");
        }
        new Setting(container).setClass("smartmd-settings-section-title").setName(localize("OneDrive 多设备同步")).setHeading();
        container.createEl("p", { text: Platform.isAndroidApp
            ? localize("用 OneDrive 在手机和电脑之间同步笔记和附件。")
            : localize("使用插件独立空间同步知识库文件与非敏感插件状态。授权、模型密钥、外部路径和设备历史不会跨设备同步。{0}", localizeMessage(this.runtime.store.storageBoundaryText())) });
        const buildInfo = container.createEl("div", { cls: "smartmd-build-info" });
        buildInfo.addClass("smartmd-build-info");
        buildInfo.setText(localize("运行实例：v{0}", this.runtime.plugin.manifest.version));
        new Setting(container).setClass("smartmd-settings-subheading").setName(localize("微软账号")).setHeading();
        this.microsoftAccountSetting(container);
        if (!this.runtime.isMicrosoftSignedIn() && !this.runtime.isMicrosoftSignInPending()) {
            this.browserLoginFallbackSetting(container);
        }
        const oneDriveStatus = new Setting(container)
            .setName(localize("同步状态"))
            .setDesc(statusDescription)
            .addButton((button) => {
                const busy = this.runtime.oneDrive.isBusy();
                button
                    .setButtonText(localizeMessage(this.runtime.oneDrive.actionLabel()))
                    .setDisabled(busy || !this.runtime.oneDrive.canStartManualSync())
                    .setCta()
                    .onClick(async () => {
                        if (this.runtime.oneDrive.isBusy()) return;
                        if (!Platform.isAndroidApp) {
                            new Notice(localize("[SmartMD v{0}] 正在检查并启动同步…", this.runtime.plugin.manifest.version), 3000);
                        }
                        button.setDisabled(true).setButtonText(localize("正在检查…"));
                        try {
                            await this.runtime.oneDrive.manualSync();
                        } finally {
                            this.display();
                        }
                    });
            });
        if (this.runtime.oneDrive.canCancelFirstUploadInspection()) {
            oneDriveStatus.addButton(button => button
                .setButtonText(localize("取消检查"))
                .onClick(() => this.runtime.oneDrive.cancelFirstUploadInspection()));
        }

        new Setting(container)
            .setName(localize("上次成功同步"))
            .setDesc(settings.lastSuccessAt ? new Date(Number(settings.lastSuccessAt)).toLocaleString() : localize("尚无成功记录"));

        if (settings.identicalRegistrationJournal &&
            (settings.identicalRegistrationJournal.phase !== "COMPLETED" || !settings.firstSyncCompleted)) {
            const registration = settings.identicalRegistrationJournal;
            const scope = registration.snapshot || {};
            new Setting(container).setName(localize("同内容登记原操作待核实"))
                .setDesc(localize("原操作：") + String(registration.operationId || "未知") +
                    localize("；账号：") + String(scope.accountId || "未知") + localize("；本地库：") + String(scope.localVaultId || "未知") +
                    localize("；云库：") + String(scope.cloudVaultId || "未知") +
                    localize("。仅核实实际持久绑定和原历史基线，不表示当前字节仍一致，不授予同步权；当前变化必须另行预览。"))
                .addButton((button) => button.setButtonText(localize("核验原绑定并结算登记")).onClick(async () => {
                    try { await this.runtime.oneDrive.reconcileIdenticalOneDriveRegistration(); }
                    catch (error) { new Notice(localizeMessage(safeOneDriveError(error)), 8000); }
                    this.display();
                }));
            const original = container.createEl("details");
            original.createEl("summary", { text: localize("查看原批准范围与历史版本") });
            for (const file of (Array.isArray(scope.files) ? scope.files : [])) original.createEl("p", { text:
                String(file.path || "未知路径") + localize("；原SHA：") + String(file.localHash || "未知") +
                localize("；远端ID：") + String(file.remoteId || "未知") + "；ETag：" + String(file.remoteEtag || "未知") });
        }

        const identityConflict = this.runtime.oneDrive.getOneDriveRemoteIdentityConflict();
        if (identityConflict) {
            const section = container.createDiv({ cls: "smartmd-onedrive-identity-conflict" });
            new Setting(section).setClass("smartmd-settings-subheading").setName(localize("云端重复身份，已停止同步")).setHeading();
            section.createEl("p", { text: localize("冲突路径：") + identityConflict.path });
            for (const item of identityConflict.items) section.createEl("p", { text:
                localize("路径：") + item.path + "；Item ID：" + item.id + "；ETag：" + item.etag });
            section.createEl("p", { text: localize("双方原文保留；只读检查不批准覆盖、删除或继续传输。") });
            const recheck = section.createEl("button", { text: localize("重新只读核对云端身份") });
            recheck.onclick = async () => {
                recheck.disabled = true;
                try { await this.runtime.oneDrive.recheckOneDriveRemoteIdentities();
                    new Notice(localize("只读核对完成。尚未批准传输。"), 7000);
                } catch (error: any) { new Notice(localizeMessage(safeOneDriveError(error)), 8000); }
                finally { this.display(); }
            };
        }

        if (settings.firstSyncCompleted && settings.confirmed && !this.runtime.oneDrive.canStartManualSync()) {
            const section = container.createDiv({ cls: "smartmd-onedrive-resume-review" });
            new Setting(section).setClass("smartmd-settings-subheading").setName(localize("同步已暂停")).setHeading();
            const reviewButton = section.createEl("button", { text: localize("重新只读核对本次文件范围") });
            const openFileReview = async (path?: string) => {
                reviewButton.disabled = true;
                try {
                    const preview = await this.runtime.oneDrive.prepareResumedOneDriveReview(path);
                    const modal = new Modal(this.runtime.app);
                    let chosen = false;
                    modal.onOpen = () => {
                        modal.titleEl.setText(localize("检查要同步的文件"));
                        modal.contentEl.createEl("p", { text:
                            localize("账户：") + preview.accountId + localize("；本地库：") + preview.localVaultId +
                            localize("；云库：") + preview.cloudVaultId + localize("；设备：") + preview.deviceId });
                        modal.contentEl.createEl("p", { text:
                            localize("所有文件、旧基线和来源版本已重新读取。此窗口不会启动后台同步；每次只能批准一项，其他文件必须重新预览。内容不同时采用较新的版本，修改时间相同以 OneDrive 为准。删除需核对方向和当前版本。") });
                        const actions = modal.contentEl.createDiv({ cls: "smartmd-actions" });
                        actions.createEl("button", { text: localize("取消，继续保持暂停") }).onclick = () => modal.close();
                        for (const entry of preview.entries) {
                            const row = modal.contentEl.createDiv({ cls: "smartmd-card" });
                            row.createEl("strong", { text: entry.path });
                            row.createEl("p", { text: localize("处理：") + entry.action + " · " + localizeMessage(entry.reason) });
                            row.createEl("p", { text: localize("本地大小：") + entry.localSize +
                                localize("字节；云端大小：") + entry.remoteSize + localize("字节") });
                            if (entry.localHash) row.createEl("p", { text: localize("本地 SHA-256：") + entry.localHash });
                            if (entry.remoteId) row.createEl("p", { text:
                                localize("远端 Item ID：") + entry.remoteId + "；ETag：" + entry.remoteEtag });
                            if (entry.remoteHash) row.createEl("p", { text:
                                localize("远端当前正文 SHA-256：") + entry.remoteHash });
                            if (entry.requiresSeparateVerification) {
                                row.createEl("button", { text: localize("单独校验并预览此文件") }).onclick = () => {
                                    modal.close();
                                    void openFileReview(entry.path);
                                };
                            }
                            const actionNames: Record<string, string> = {
                                UPLOAD: "上传此文件",
                                DOWNLOAD: "下载此文件",
                                DELETE_CLOUD: "删除云端此版本（先冷备）",
                                DELETE_VAULT: "删除本机文件",
                            };
                            const actionLabel = localizeMessage(actionNames[entry.action]);
                            if (!actionLabel) continue;
                            if (entry.action.startsWith("DELETE_")) row.createEl("p", { text:
                                localize("危险操作：必须再次确认本次删除方向与版本。取消即保留原文。") });
                            const approve = row.createEl("button", {
                                text: localize("明确确认本条：") + actionLabel + localize("（仅本次/仅此路径）"),
                            });
                            approve.onclick = async () => {
                                if (chosen) return;
                                chosen = true;
                                approve.disabled = true;
                                try {
                                    const committed = await this.runtime.oneDrive.confirmResumedOneDriveReview(
                                        preview.ticket, entry.path, true,
                                    );
                                    if (committed) new Notice(localize("本条已按当前版本单独核对；其余范围仍暂停。"), 6000);
                                } catch (error: any) {
                                    new Notice(localize("本轮授权未生效：") + localizeMessage(safeOneDriveError(error)), 8000);
                                } finally {
                                    modal.close();
                                    this.display();
                                }
                            };
                        }
                    };
                    modal.onClose = () => {
                        if (!chosen) void this.runtime.oneDrive.confirmResumedOneDriveReview(
                            preview.ticket, "", false,
                        );
                        modal.contentEl.empty();
                    };
                    modal.open();
                } catch (error: any) {
                    new Notice(localize("只读重新核对失败：") + localizeMessage(safeOneDriveError(error)), 8000);
                } finally {
                    reviewButton.disabled = false;
                }
            };
            reviewButton.onclick = () => void openFileReview();
        }

        // UNKNOWN uploads are not failed uploads that may be retried.
        // A read-only audit can explain matched original/staging versions,
        // but deliberately has no "release", "retry" or "delete" button.
        const pendingUnknown = Object.entries(settings.pendingUploadOperations || {})
            .filter(([, record]: [string, any]) => record && this.runtime.oneDrive.canAuditUnknownUpload(record));
        if (Object.keys(settings.pendingUploadOperations || {}).length && this.runtime.oneDrive.canStartManualSync()) {
            const entries = Object.keys(settings.pendingUploadOperations || {});
            const info = new Setting(container)
                .setName(localize("有 {0} 个文件等待核对上传结果", entries.length))
                .setDesc(localize("其他文件照常同步。每次同步会自动检查这些文件。"));
            const details = info.descEl.createEl("details");
            details.createEl("summary", { text: localize("查看文件") });
            for (const path of entries) details.createEl("p", { text: path });
        } else if (pendingUnknown.length) {
            const section = container.createDiv({ cls: "smartmd-onedrive-unknown-review" });
            new Setting(section).setClass("smartmd-settings-subheading").setName(localize("OneDrive 上次上传结果未明：") + pendingUnknown.length + localize(" 项")).setHeading();
            section.createEl("p", { text:
                localize("历史上传默认锁定且绝不重传。先独立只读核原操作、云端目标、暂存副本及本机SHA；只有完全匹配时才能对该文件单独二次确认结算，其他范围仍暂停。") });
            const button = section.createEl("button", { text: localize("只读检查上次未知上传（不修改云端）") });
            button.onclick = async () => {
                button.disabled = true;
                try {
                    const audit = await this.runtime.oneDrive.auditPendingUnknownUploadsReadOnly();
                    const modal = new Modal(this.runtime.app);
                    modal.onOpen = () => {
                        modal.titleEl.setText(localize("上次未知上传：只读证据与逐文件结算"));
                        modal.contentEl.createEl("p", { text:
                            localize("账号：") + audit.accountId + localize("；云库：") + audit.vaultId +
                            localize("。只读检查不修改任何文件；完全一致时可另行确认该项结算，不会重传原上传。") });
                        for (const entry of audit.entries) {
                            const row = modal.contentEl.createDiv({ cls: "smartmd-card" });
                            row.createEl("strong", { text: entry.path });
                            row.createEl("p", { text: (entry.status === "VERIFIED_BUT_LOCKED"
                                ? localize("三处字节一致：只读检查仍锁定，可选择单文件确认结算。") : localize("未通过完整核验，继续锁定。")) +
                                " " + entry.reason });
                            if (entry.remoteId && entry.stagingId) {
                                row.createEl("p", { text:
                                    localize("远端原稿ID：") + entry.remoteId + localize("；版本：") + entry.remoteEtag });
                                row.createEl("p", { text:
                                    localize("独立暂存ID：") + entry.stagingId + localize("；版本：") + entry.stagingEtag });
                                row.createEl("p", { text: localize("本机/远端源SHA-256：") + entry.sourceSha256 });
                            }
                            if (entry.status === "VERIFIED_BUT_LOCKED") {
                                const choose = row.createEl("button", {
                                    text: localize("重新核对并准备单文件结算（不重传）"),
                                });
                                choose.onclick = async () => {
                                    choose.disabled = true;
                                    try {
                                        const proof = await this.runtime.oneDrive.prepareUnknownUploadSettlement(entry.path);
                                        const confirm = new Modal(this.runtime.app);
                                        let consumed = false;
                                        confirm.onOpen = () => {
                                            confirm.titleEl.setText(localize("确认结算这一次 UNKNOWN 上传"));
                                            confirm.contentEl.createEl("p", { text:
                                                localize("仅此文件：") + proof.path + localize("；微软账号：") + proof.accountId +
                                                localize("；云库：") + proof.vaultId });
                                            confirm.contentEl.createEl("p", { text:
                                                localize("目标ID：") + proof.remoteId + "；ETag：" + proof.remoteEtag });
                                            confirm.contentEl.createEl("p", { text:
                                                localize("原操作暂存ID：") + proof.stagingId + "；ETag：" + proof.stagingEtag });
                                            confirm.contentEl.createEl("p", { text:
                                                localize("三个版本已核实的SHA-256：") + proof.sourceSha256 });
                                            confirm.contentEl.createEl("p", { text:
                                                localize("确认后先持久记录可信基线，再按精确暂存Item ID/ETag清理该测试文件的暂存副本；不会重传原稿或开启自动同步。失败时保持旧操作锁。") });
                                            const cancel = confirm.contentEl.createEl("button", {
                                                text: localize("取消，保持暂停"),
                                            });
                                            cancel.onclick = () => confirm.close();
                                            const approve = confirm.contentEl.createEl("button", {
                                                text: localize("确认：只结算这一文件"),
                                            });
                                            approve.onclick = async () => {
                                                if (consumed) return;
                                                consumed = true;
                                                approve.disabled = true;
                                                try {
                                                    const settled = await this.runtime.oneDrive.confirmUnknownUploadSettlement(proof.ticket, true);
                                                    if (!settled) throw new Error("本次结算没有生效");
                                                    new Notice(localize("该文件UNKNOWN已核对结算；其他同步仍暂停。"), 7000);
                                                } catch (error: any) {
                                                    new Notice(localize("结算未全部确认，保留待处理记录：") +
                                                        localizeMessage(safeOneDriveError(error)), 8000);
                                                } finally {
                                                    confirm.close();
                                                    modal.close();
                                                    this.display();
                                                }
                                            };
                                        };
                                        confirm.onClose = () => {
                                            if (!consumed) {
                                                consumed = true;
                                                void this.runtime.oneDrive.confirmUnknownUploadSettlement(proof.ticket, false);
                                            }
                                        };
                                        confirm.open();
                                    } catch (error: any) {
                                        new Notice(localize("结算前核验失败，继续锁定：") +
                                            localizeMessage(safeOneDriveError(error)), 8000);
                                    } finally {
                                        choose.disabled = false;
                                    }
                                };
                            }
                        }
                        modal.contentEl.createEl("button", { text: localize("关闭，保持暂停") }).onclick = () => modal.close();
                    };
                    modal.open();
                } catch (error: any) {
                    new Notice(localize("未知上传只读核验未完成，继续保留历史锁：") +
                        localizeMessage(safeOneDriveError(error)), 8000);
                } finally {
                    button.disabled = false;
                }
            };
        }

        const interruptedCleanup = Object.entries(settings.pendingUploadOperations || {})
            .filter(([, record]: [string, any]) =>
                ["FINAL_ACKNOWLEDGED", "CLEANUP_ISSUED", "CLEANUP_VERIFIED"].includes(
                    String(record?.phase || "")));
        if (interruptedCleanup.length) {
            const section = container.createDiv({ cls: "smartmd-onedrive-interrupted-cleanup" });
            new Setting(section).setClass("smartmd-settings-subheading").setName(localize("OneDrive 待结算的上传清理：") + interruptedCleanup.length + localize(" 项")).setHeading();
            section.createEl("p", { text:
                localize("历史可信基线已保存，但原暂存清理的结果可能未知。只有独立核对目标字节、原操作身份及暂存现状后，才能单文件确认恢复；不重传母文件。") });
            for (const [path] of interruptedCleanup) {
                const row = section.createDiv({ cls: "smartmd-card" });
                row.createEl("strong", { text: path });
                const prepareButton = row.createEl("button", {
                    text: localize("只读核对并准备恢复这一文件"),
                });
                prepareButton.onclick = async () => {
                    prepareButton.disabled = true;
                    try {
                        const proof = await this.runtime.oneDrive.prepareInterruptedUploadCleanup(path);
                        const modal = new Modal(this.runtime.app);
                        let consumed = false;
                        modal.onOpen = () => {
                            modal.titleEl.setText(localize("核对原操作后恢复清理"));
                            modal.contentEl.createEl("p", { text:
                                localize("仅此文件：") + proof.path + localize("；账号：") + proof.accountId +
                                localize("；云库：") + proof.vaultId });
                            modal.contentEl.createEl("p", { text:
                                localize("母文件ID：") + proof.remoteId + localize("；版本：") + proof.remoteEtag });
                            modal.contentEl.createEl("p", { text:
                                localize("原暂存ID：") + proof.stageId + localize("；版本：") + proof.stageEtag });
                            modal.contentEl.createEl("p", { text:
                                localize("已核验源SHA：") + proof.sourceSha256 + localize("。暂存状态：") +
                                (proof.stageState === "PRESENT"
                                    ? localize("仍存在（只条件清理原暂存）")
                                    : localize("对原Item ID的GET已确认404（本次仅认可不存在并释放旧锁）")) });
                            modal.contentEl.createEl("p", { text:
                                localize("这不是重新上传，也不允许删除最终母文件；确认只结算本次已核验的历史操作。") });
                            modal.contentEl.createEl("button", {
                                text: localize("取消，继续锁定"),
                            }).onclick = () => modal.close();
                            const approve = modal.contentEl.createEl("button", {
                                text: localize("确认恢复这一次清理"),
                            });
                            approve.onclick = async () => {
                                if (consumed) return;
                                consumed = true;
                                approve.disabled = true;
                                try {
                                    const done = await this.runtime.oneDrive.confirmInterruptedUploadCleanup(proof.ticket, true);
                                    if (!done) throw new Error("原操作仍未结算");
                                    new Notice(localize("这次历史暂存状态已核实结算；其他文件仍暂停。"), 7000);
                                } catch (error: any) {
                                    new Notice(localize("中断清理未成功结算，继续保留原记录：") +
                                        localizeMessage(safeOneDriveError(error)), 8000);
                                } finally {
                                    modal.close();
                                    this.display();
                                }
                            };
                        };
                        modal.onClose = () => {
                            if (!consumed) {
                                consumed = true;
                                void this.runtime.oneDrive.confirmInterruptedUploadCleanup(proof.ticket, false);
                            }
                        };
                        modal.open();
                    } catch (error: any) {
                        new Notice(localize("清理恢复只读核验失败，继续锁定：") +
                            localizeMessage(safeOneDriveError(error)), 8000);
                    } finally {
                        prepareButton.disabled = false;
                    }
                };
            }
        }

        const retained = this.runtime.oneDrive.listAndroidRetainedDeletionPaths();
        if (retained.length) {
            const area = container.createDiv({ cls: "smartmd-onedrive-delete-review" });
            new Setting(area).setClass("smartmd-settings-subheading").setName(localize("云端缺失，本机笔记已保留")).setHeading();
            area.createEl("p", { text: localize("下面的安卓笔记有未同步的编辑，或缺少可核验的同步版本。请比较两个版本并保留需要的内容。") });
            const paths = area.createEl("ul");
            for (const path of retained) paths.createEl("li", { text: path });
        }
        const pending = this.runtime.oneDrive.listPendingOneDriveDeletions();
        if (pending.length) {
            const area = container.createDiv({ cls: "smartmd-onedrive-delete-review" });
            new Setting(area).setClass("smartmd-settings-subheading").setName(localize("待人工确认的单端删除（") + pending.length + localize("项）")).setHeading();
            area.createEl("p", { text:
                localize("勾选要在另一端删除的文件。") });
            const chosen = new Set<string>();
            const displayed = new Map(pending.map(item => [item.path, item.reviewKey]));
            for (const item of pending) {
                const row = area.createDiv({ cls: "smartmd-card" });
                const checkbox = row.createEl("input", { type: "checkbox" });
                checkbox.onchange = () => {
                    if (checkbox.checked) chosen.add(item.path);
                    else chosen.delete(item.path);
                };
                row.createEl("strong", { text: item.path });
                row.createEl("p", { text: item.direction === "cloud"
                    ? localize("授权后会将云端这份移入 OneDrive 回收站")
                    : Platform.isAndroidApp ? localize("授权后直接删除安卓上的本地文件")
                    : localize("授权后会将此本地笔记移入 Windows 系统回收站") });
            }
            area.createEl("button", { text: localize("确认勾选文件的删除"), cls: "mod-warning" })
                .onclick = async () => {
                    if (!chosen.size) { new Notice(localize("请先勾选目标文件")); return; }
                    try {
                        await this.runtime.oneDrive.confirmPendingOneDriveDeletions(
                            [...chosen], true, undefined, displayed);
                        new Notice(localize("已删除选中的文件"));
                    } catch (error: any) {
                        new Notice(localize("本次删除未完成：") + localizeMessage(safeOneDriveError(error)), 8000);
                    }
                    this.display();
                };
        }
        const recoveries = this.runtime.oneDrive.listOneDriveRecoveries();
        if (recoveries.length) {
            const area = container.createDiv({ cls: "smartmd-onedrive-recoveries" });
            new Setting(area).setClass("smartmd-settings-subheading").setName(localize("OneDrive 独立恢复档案：") + recoveries.length + localize("项")).setHeading();
            area.createEl("p", { text:
                localize("未解决、无法验证或唯一副本不会自动销毁。需要先另存旧版并核验，再明确确认问题解决。") });
            for (const item of recoveries) {
                const row = area.createDiv({ cls: "smartmd-card" });
                row.createEl("strong", { text: String(item.relativePath || "未知文件") });
                row.createEl("p", { text: localize("恢复 ID：") + String(item.id || "") });
                row.createEl("p", { text: item.resolvedAt ? localize("已解决（保留缓冲）")
                    : item.exportedAt ? localize("已单独导出，待确认解决")
                        : localize("尚未解决，保留锁生效") });
                if (item.cleanupError) {
                    row.createEl("p", { cls: "smartmd-warning", text:
                        localize("清理失败，恢复档案仍保留：") + localizeMessage(safeOneDriveError(item.cleanupError)) });
                }
                if (Platform.isAndroidApp && item.sourceTrashPath) {
                    row.createEl("button", { text: localize("另存安卓回收原件") }).onclick = async () => {
                        try {
                            const path = await this.runtime.oneDrive.exportAndroidTrashedOriginal(item.id);
                            new Notice(localize("安卓回收原件已另存至：") + path);
                        } catch (error: any) { new Notice(localizeMessage(safeOneDriveError(error)), 8000); }
                    };
                }
                if (!item.exportedAt) {
                    row.createEl("button", { text: localize("另存恢复副本，不覆盖原稿") }).onclick = async () => {
                        try {
                            const path = await this.runtime.oneDrive.exportOneDriveRecoveryToNewNote(item.id);
                            new Notice(localize("旧版本已核验并另存至：") + path);
                        } catch (error: any) {
                            new Notice(localize("导出失败：") + localizeMessage(safeOneDriveError(error)), 8000);
                        }
                        this.display();
                    };
                } else if (!item.resolvedAt) {
                    row.createEl("button", { text: localize("确认恢复问题已解决") }).onclick = async () => {
                        if (!window.confirm(localize("确认已核验独立导出的恢复笔记？恢复档案将继续受7天缓冲保护。"))) return;
                        try { await this.runtime.oneDrive.confirmOneDriveRecoveryResolved(item.id); }
                        catch (error: any) { new Notice(localizeMessage(safeOneDriveError(error)), 8000); }
                        this.display();
                    };
                }
            }
        }
        const sharingApproved = this.runtime.oneDrive.hasOneDriveSettingsSharingApproval();
        const sharingError = this.runtime.oneDrive.getSharedSettingsError();
        new Setting(container)
            .setName(localize("OneDrive 插件设置共享（独立授权）"))
            .setDesc(sharingError ? localizeMessage(sharingError) : sharingApproved
                ? localize("已单独许可本次字段版本的一次共享；发送后需重新预览批准")
                : localize("默认关闭；只允许更新时间、批准的模型端点/型号及固定发布服务地址"))
            .addButton((button) => button.setButtonText(
                sharingApproved ? localize("本次已批准") : localize("预览并批准四字段"))
                .setDisabled(sharingApproved)
                .onClick(async () => {
                    const preview = this.runtime.oneDrive.getOneDriveSettingsSharingPreview();
                    const fields = Object.entries(preview.fields).map(([key, value]) => key + "：" + String(value)).join("\n");
                    if (!window.confirm(localize("微软账号：{0}；云库：{1}；本地库：{2}；设备：{3}\n本次仅共享以下已配置字段（一次有效）：\n{4}\n不共享凭据、提示词、发布空间和设备执行许可。确定独立授权吗？", preview.accountId, preview.vaultId, preview.localVaultId, preview.deviceId, fields))) return;
                    try { await this.runtime.oneDrive.approveOneDriveSettingsSharing(true, preview.scope); }
                    catch (error: any) { new Notice(localizeMessage(safeOneDriveError(error)), 8000); }
                    this.display();
                }));

        if (settings.pendingSharedUploadOperation) {
            new Setting(container)
                .setName(localize("共享设置原操作待核验"))
                .setDesc(settings.pendingSharedUploadOperation.phase === "FAILED"
                    ? localize("原请求已明确失败；不会自动重试，下一次发送必须重新预览并独立批准。")
                    : localize("原请求结果尚未确认，不会自动重发；只读核验不会清理旧云设置或应用云端字段。"))
                .addButton((button) => button.setButtonText(localize("只读核验原共享操作")).onClick(async () => {
                    try {
                        const verified = await this.runtime.oneDrive.reconcileOneDriveSharedUpload();
                        new Notice(verified ? localize("已核实原共享内容，后续发送需重新独立批准。") :
                            localize("尚不能核实原操作，继续保留记录并禁止重发。"), 8000);
                    } catch (error: any) { new Notice(localizeMessage(safeOneDriveError(error)), 8000); }
                    this.display();
                }));
        }

        const pendingShared = this.runtime.oneDrive.getPendingOneDriveIncomingSharedSettings();
        if (pendingShared) {
            const previewLines = Object.entries(pendingShared.fields).map(([key, value]) =>
                key + "：" + String(value)).join("\n");
            new Setting(container)
                .setName(localize("收到云端共享设置变更（等待本机确认）"))
                .setDesc(localize("微软账号：") + pendingShared.accountId + localize("；云端库：") +
                    pendingShared.vaultId + localize("。仅允许四字段白名单，不会自动改写提示词、发布空间或执行任务。"))
                .addButton((button) => button.setButtonText(localize("预览并批准接收"))
                    .onClick(async () => {
                        if (!window.confirm(localize("即将在本机应用以下获准的共享设置：\n{0}\n本操作与上传共享、笔记同步授权相互独立。确认接收？", previewLines))) return;
                        try {
                            await this.runtime.oneDrive.confirmPendingOneDriveIncomingSharedSettings(true);
                        } catch (error: any) {
                            new Notice(localize("接收未生效：") + localizeMessage(safeOneDriveError(error)), 8000);
                        }
                        this.display();
                    }))
                .addButton((button) => button.setButtonText(localize("拒绝此远端变更"))
                    .onClick(async () => {
                        await this.runtime.oneDrive.confirmPendingOneDriveIncomingSharedSettings(false);
                        this.display();
                    }));
        }

        const summary = settings.lastSyncSummary;
        const diagnostics: SyncFailureDiagnostic[] = settings.failedDiagnostics || [];

        if (summary) {
            const summaryDiv = container.createEl("div", { cls: "smartmd-sync-summary" });
            summaryDiv.addClass("smartmd-sync-summary");

            const summaryTitle = summaryDiv.createEl("div");
            summaryTitle.addClass("smartmd-sync-summary-title");

            const statusBadge = settings.status === "ready"
                ? "🟢 全部同步成功"
                : settings.status === "partial_success"
                    ? "🟡 部分同步完成 (已隔离失败项)"
                    : "🔴 同步未完成";
            summaryTitle.createSpan({ text: localize("上次同步汇报：{0}", localizeMessage(statusBadge)) });
            if (summary.durationMs) {
                summaryTitle.createSpan({
                    text: localize("耗时: {0}s", localizeMessage((summary.durationMs / 1000).toFixed(1))),
                    cls: "smartmd-text-muted",
                });
            }

            const statsRow = summaryDiv.createEl("div");
            statsRow.addClass("smartmd-sync-summary-counts");
            const counts: Array<[string, number, string?]> = [
                ["总计", summary.total],
                ["成功", summary.success, "smartmd-sync-count-success"],
                ["失败", summary.failed, summary.failed > 0 ? "smartmd-sync-count-failed" : undefined],
                ["未变更", summary.unchanged],
            ];
            for (const [label, count, cls] of counts) {
                const item = statsRow.createSpan({ text: localize(label) + ": ", cls });
                item.createEl("b", { text: String(count) });
            }
        }

        if (diagnostics.length > 0) {
            const diagSection = container.createEl("div", { cls: "smartmd-diagnostic-center" });

            const diagHeader = diagSection.createEl("div");
            diagHeader.addClass("smartmd-diagnostic-header");

            const headerLeft = diagHeader.createEl("div");
            new Setting(headerLeft).setClass("smartmd-diag-title").setName(localize("⚠️ 失败文件智能诊断中心 ({0} 项需要调整)", diagnostics.length)).setHeading();
            headerLeft.addClass("smartmd-diagnostic-header-title");

            const retryBtn = diagHeader.createEl("button", { text: localize("仅重试失败项"), cls: "mod-cta" });
            retryBtn.onclick = async () => {
                retryBtn.disabled = true;
                retryBtn.setText(localize("正在重试…"));
                try {
                    await this.runtime.syncOneDriveFromShortcut();
                } finally {
                    this.display();
                }
            };

            for (const diag of diagnostics) {
                const card = diagSection.createEl("div", { cls: "smartmd-diag-card" });
                card.addClass("smartmd-diagnostic-card");

                // 标题行
                const cardTitle = card.createEl("div");
                cardTitle.addClass("smartmd-diagnostic-card-title");
                cardTitle.createSpan({ text: `📄 ${diag.relativePath}` });
                const badge = cardTitle.createSpan({ text: `[${localizeMessage(diag.remediation?.badgeText || "失败")}]` });
                badge.addClass("smartmd-diagnostic-badge");

                // 凭什么失败（技术证据）
                const evDiv = card.createEl("div");
                evDiv.addClass("smartmd-diagnostic-evidence");
                const ev = diag.evidence || {} as any;
                evDiv.createEl("b", { text: localize("【凭什么失败】 ") });
                evDiv.createSpan({ text: `HTTP ${ev.httpStatus || 0} (${ev.gatewayErrorCode || localize("未知代码")})` });
                if (ev.requestId) {
                    evDiv.createSpan({ text: " | Request-ID: " });
                    evDiv.createEl("code", { text: ev.requestId });
                }

                // 为什么这几个会失败（通俗说明）
                const reasonDiv = card.createEl("div");
                reasonDiv.addClass("smartmd-diagnostic-reason");
                reasonDiv.createEl("b", { text: localize("【为什么失败】 ") });
                reasonDiv.createSpan({ text: [localizeMessage(diag.humanReason), localizeMessage(diag.humanExplanation)].join(" · ") });

                // 要怎么调整才能不失败（行动指南）
                const remDiv = card.createEl("div");
                remDiv.addClass("smartmd-diagnostic-action");
                remDiv.createEl("b", { text: localize("【怎么调整】 ") });
                remDiv.createEl("span", { text: localizeMessage(diag.remediation?.actionTitle) || localize("操作指引") });
                const instructList = remDiv.createEl("div");
                instructList.addClass("smartmd-diagnostic-instructions");
                for (const line of (diag.remediation?.stepInstructions || [])) {
                    instructList.createEl("div", { text: localizeMessage(line) });
                }
                if (diag.remediation?.suggestedNewPath) {
                    const suggestP = remDiv.createEl("div");
                    suggestP.addClass("smartmd-diagnostic-suggestion");
                    suggestP.createSpan({ text: localize("建议更名为：") });
                    suggestP.createEl("code", { text: diag.remediation.suggestedNewPath });
                }
            }
        }

        new Setting(container)
            .setName(localize("云端知识库"))
            .setDesc(settings.vaultName
                ? localize("已连接：{0}", settings.vaultName)
                : Platform.isAndroidApp ? localize("尚未连接 OneDrive 知识库")
                : localize("首次同步时选择已有云端知识库，或创建新的独立空间。"));
        new Setting(container)
            .setName(localize("同步方式"))
            .setDesc(settings.firstSyncCompleted
                ? localize("推荐使用智能自动同步；移动端只在 Obsidian 打开期间运行。")
                : localize("首次手动同步成功后才能启用自动同步。"))
            .addDropdown((dropdown) => dropdown
                .addOption("smart", localize("智能自动同步（推荐）"))
                .addOption("manual", localize("仅手动同步"))
                .addOption("custom", localize("自定义"))
                .setValue(autoMode)
                .setDisabled(!settings.firstSyncCompleted)
                .onChange(async (value) => {
                    if (value !== "manual" && !this.runtime.oneDrive.canRun()) {
                        new Notice(localize("请先恢复同步，再开启自动同步。"), 8000);
                        this.display();
                        return;
                    }
                    try {
                        if (value === "manual") {
                            await this.runtime.oneDrive.selectManualSyncMode();
                        } else {
                            settings.autoMode = value;
                            settings.enabled = true;
                            if (value === "smart") {
                                settings.onStartup = true;
                                settings.afterSave = true;
                                settings.intervalMinutes = Platform.isMobile ? 5 : 1;
                            }
                            await this.runtime.oneDrive.enableAutoSync();
                        }
                        await (this.runtime.store as any).save();
                    } catch (error: any) {
                        // A failed mode transition must not leave a misleading
                        // enabled toggle or a latent automatic trigger.
                        settings.enabled = false;
                        settings.autoMode = "manual";
                        settings.onStartup = false;
                        settings.afterSave = false;
                        settings.intervalMinutes = 0;
                        await (this.runtime.store as any).save();
                        new Notice(localize("切换同步模式未成功：") + localizeMessage(safeOneDriveError(error)), 8000);
                    } finally {
                        this.display();
                    }
                }));
        if (autoMode === "smart") {
            new Setting(container)
                .setName(localize("智能自动同步"))
                .setDesc(Platform.isMobile
                    ? localize("打开 Obsidian 后同步；停止编辑 10 秒后同步；每 5 分钟检查一次云端变化。")
                    : localize("打开 Obsidian 后同步；停止编辑 10 秒后同步；每 1 分钟检查一次云端变化。"));
        }
        if (autoMode === "custom") {
            const warning = this.runtime.oneDrive.autoSyncWarning();
            if (warning) container.createEl("p", { cls: "smartmd-warning", text: localizeMessage(warning) });
            new Setting(container)
                .setName(localize("打开 Obsidian 后同步"))
                .addToggle((toggle) => toggle.setValue(Boolean(settings.onStartup)).onChange(async (value) => {
                    settings.onStartup = value;
                    await (this.runtime.store as any).save();
                    this.display();
                }));
            new Setting(container)
                .setName(localize("停止编辑后同步"))
                .setDesc(localize("连续修改会合并处理，停止编辑约 2.5 秒后执行一次。"))
                .addToggle((toggle) => toggle.setValue(Boolean(settings.afterSave)).onChange(async (value) => {
                    settings.afterSave = value;
                    await (this.runtime.store as any).save();
                    this.display();
                }));
            new Setting(container)
                .setName(localize("定时检查云端"))
                .setDesc(localize("用于发现其他设备上传的新变化。"))
                .addDropdown((dropdown) => dropdown
                    .addOption("0", localize("关闭"))
                    .addOption("1", localize("每 1 分钟"))
                    .addOption("5", localize("每 5 分钟"))
                    .addOption("10", localize("每 10 分钟"))
                    .addOption("30", localize("每 30 分钟"))
                    .setValue(String(settings.intervalMinutes || 0))
                    .onChange(async (value) => {
                        settings.intervalMinutes = Number(value);
                        await (this.runtime.store as any).save();
                        this.display();
                    }));
        }
        new Setting(container)
            .setName(localize("设备名称"))
            .setDesc(localize("此名称用于识别当前同步设备。"))
            .addText((text) => text.setValue(settings.deviceName || "").onChange(async (value) => { settings.deviceName = value.trim(); await (this.runtime.store as any).save(); }));
    }

    private externalSection(container: HTMLElement): void {
        new Setting(container).setClass("smartmd-settings-section-title").setName(localize("外部文件夹同步")).setHeading();
        container.createEl("p", {
            text: localize("选择知识库外部的文件夹后，插件会扫描该文件夹及子目录中的全部 Markdown 文档，保留原有目录结构，并在知识库的“外部同步文件夹”中建立副本。之后无论修改知识库内的副本，还是修改外部文件夹中的原文档，另一端都会实时同步更新。图片、代码和附件不会被同步或改动。"),
        });
        new Setting(container)
            .setName(localize("添加外部文件夹"))
            .setDesc(localize("当前有 {0} 个独立同步项。", this.runtime.external?.bindings().length || 0))
            .addButton((button) => button
                .setButtonText(localize("添加同步文件夹"))
                .setCta()
                .onClick(() => this.runtime.external?.openAddBinding(() => this.display())));

        const bindings = this.runtime.external?.bindings() || [];
        const list = container.createDiv({ cls: "smartmd-external-inline-list" });
        if (!bindings.length) {
            list.createEl("p", { cls: "smartmd-external-empty", text: localize("尚未添加外部同步文件夹。") });
        }

        for (const binding of bindings) {
            const item = list.createDiv({ cls: "smartmd-external-inline-item" });
            new Setting(item).setClass("smartmd-settings-subheading").setName(binding.nickname).setHeading();
            const pathRow = item.createDiv({ cls: "smartmd-external-inline-path-row" });
            pathRow.createDiv({ cls: "smartmd-external-inline-path", text: binding.externalPath });
            const actions = pathRow.createDiv({ cls: "smartmd-external-inline-actions" });
            actions.createEl("button", { text: localize("解除绑定") }).onclick = async () => {
                const confirmed = window.confirm(
                    localize("确定解除“{0}”的绑定吗？\n\n只会停止这个文件夹的同步，其他同步文件夹会继续运行。外部原文件夹和知识库内现有副本都会保留。", binding.nickname),
                );
                if (!confirmed) return;
                await this.runtime.external?.remove(binding.id);
                this.display();
            };
            actions.createEl("button", { text: localize("立即同步") }).onclick = async () => {
                await this.runtime.external?.syncBinding(binding);
                this.display();
            };
            actions.createEl("button", { text: localize("⚙️ 筛选规则") }).onclick = () => {
                this.runtime.external?.openFilterModal(binding, () => this.display());
            };
            actions.createEl("button", { text: localize("打开恢复区") }).onclick = () => this.runtime.external?.openRecovery(binding.id);
            const isActive = binding.status === "active";
            actions.createEl("button", { text: isActive ? localize("暂停") : localize("恢复同步") }).onclick = async () => {
                if (isActive) await this.runtime.external?.pause(binding.id);
                else await this.runtime.external?.resume(binding.id);
                this.display();
            };
            const filterSummary = binding.filterConfig?.mode === "include"
                ? `白名单 (${(binding.filterConfig.userRules || []).length}条规则)`
                : `排除模式 (${(binding.filterConfig?.userRules || []).length}条规则)`;
            item.createEl("p", {
                cls: "setting-item-description smartmd-external-inline-status",
                text: localize("{0} · 筛选: {1}", localizeMessage(this.runtime.external?.statusLabel(binding.status)), localizeMessage(filterSummary)),
            });
            if (binding.lastSuccessAt) item.createEl("p", { cls: "setting-item-description",
                text: localize("上次成功：{0}", new Date(binding.lastSuccessAt).toLocaleString()) });
            if (binding.lastError) item.createEl("p", { cls: "setting-item-description", text: localizeMessage(binding.lastError) });
        }
        const archives = this.runtime.external?.recoveryArchives() || [];
        if (archives.length) {
            new Setting(container).setClass("smartmd-settings-subheading").setName(localize("已解绑项目的 7 天恢复区")).setHeading();
            for (const archive of archives) {
                new Setting(container).setName(archive.nickname).addButton(button => button
                    .setButtonText(localize("打开恢复区")).onClick(() => this.runtime.external?.openRecovery(archive.id)));
            }
        }
    }

    private aiSettingsChanged(): void {
        this.aiRevision++;
        this.aiTestController?.abort();
        if (!this.aiTestRunning) {
            this.aiTestMessage = "";
            this.refreshAiTest?.();
        }
    }

    private async testAiConnection(): Promise<void> {
        if (this.aiTestRunning) return;
        if (this.aiFormNeedsRefresh()) {
            this.aiSettingsChanged();
            this.aiTestMessage = "模型设置已变化，请重新测试。";
            this.display();
            return;
        }
        const revision = this.aiRevision;
        const controller = new AbortController();
        const { aiApiUrl, aiModel } = this.runtime.store.settings.shared;
        this.aiTestController = controller;
        this.aiTestTarget = { aiApiUrl, aiModel };
        this.aiTestRunning = true;
        this.aiTestMessage = "正在测试模型连接…";
        this.refreshAiTest?.();
        let timer: number | undefined;
        try {
            const test = async () => {
                const stored = await this.aiKeySave;
                if (controller.signal.aborted || revision !== this.aiRevision) return false;
                const settings = modelSettings(this.runtime.store);
                if (settings.aiApiUrl !== aiApiUrl || settings.aiModel !== aiModel) return false;
                if (stored?.address === settings.aiApiUrl && !stored.saved) throw new Error("模型密钥未保存，请重新填写");
                const canContinue = () => {
                    const current = modelSettings(this.runtime.store);
                    return !controller.signal.aborted && settings.aiApiUrl === current.aiApiUrl
                        && settings.aiModel === current.aiModel && settings.secret === current.secret;
                };
                await testModelConnection(settings, canContinue);
                return canContinue();
            };
            const stopped = new Promise<never>((_, reject) => {
                controller.signal.addEventListener("abort", () => {
                    reject(new Error("等待模型响应超过 30 秒，请检查服务状态后重试。"));
                }, { once: true });
                timer = window.setTimeout(() => controller.abort(), 30_000);
            });
            const matches = await Promise.race([test(), stopped]);
            this.aiTestMessage = matches ? "测试通过：模型已返回可应用的改写和润色结果。" : "模型设置已变化，请重新测试。";
        } catch (error) {
            const detail = error instanceof Error ? error.message : "无法连接模型";
            this.aiTestMessage = `测试失败：${detail}`;
        } finally {
            if (timer !== undefined) window.clearTimeout(timer);
            this.aiTestController = undefined;
            this.aiTestTarget = undefined;
            if (revision !== this.aiRevision) this.aiTestMessage = "模型设置已变化，请重新测试。";
            this.aiTestRunning = false;
            this.refreshAiTest?.();
        }
    }

    private aiSection(container: HTMLElement): void {
        const shared = this.runtime.store.settings.shared;
        this.aiAddressShown = shared.aiApiUrl;
        this.aiModelShown = shared.aiModel;
        let keyInput: HTMLInputElement | undefined;
        new Setting(container).setClass("smartmd-settings-section-title").setName(localize("基础 AI 编辑器")).setHeading();
        container.createEl("p", { text: localize("支持兼容 OpenAI Chat Completions（/chat/completions）的接口。按服务文档填写地址、模型 ID 和所需的 API Key。") });
        container.createEl("p", { text: localize("AI 编辑里的「联网搜索」使用 Responses（/responses）。请选用服务商开放搜索的模型。下方的连接测试检查普通编辑。") });
        container.createEl("p", { text: localizeMessage(this.runtime.store.storageBoundaryText()) });
        const addExample = (input: HTMLInputElement, id: string, example: string) => {
            const field = createDiv({ cls: "smartmd-ai-model-field" });
            input.before(field);
            field.appendChild(input);
            field.createDiv({ cls: "smartmd-ai-model-example", attr: { id } }).createEl("code", { text: example });
            input.setAttribute("aria-describedby", id);
        };
        new Setting(container).setClass("smartmd-ai-model-setting")
            .setName(localize("模型服务地址"))
            .setDesc(localize("支持 HTTP 和 HTTPS。填写 API 基础地址（保留 /v1 等路径）或完整接口地址。公网服务请使用 HTTPS。"))
            .addText((text) => {
                text.setPlaceholder(localize("API 基础地址")).setValue(shared.aiApiUrl || "").onChange(async (value) => {
                    const current = this.runtime.store.settings.shared;
                    current.aiApiUrl = value.trim();
                    this.aiAddressShown = current.aiApiUrl;
                    if (keyInput) keyInput.value = this.runtime.store.getAIKey(current.aiApiUrl);
                    this.aiSettingsChanged();
                    await this.runtime.store.save();
                });
                addExample(text.inputEl, "smartmd-ai-url-example", "本机：http://127.0.0.1:11434/v1　云端：https://api.openai.com/v1");
            });
        new Setting(container).setClass("smartmd-ai-model-setting")
            .setName(localize("模型名称"))
            .setDesc(localize("填写服务商标注的模型 ID，按原样复制。"))
            .addText((text) => {
                text.setPlaceholder(localize("模型 ID")).setValue(shared.aiModel || "").onChange(async (value) => {
                    const current = this.runtime.store.settings.shared;
                    current.aiModel = value.trim();
                    this.aiModelShown = current.aiModel;
                    this.aiSettingsChanged();
                    await this.runtime.store.save();
                });
                addExample(text.inputEl, "smartmd-ai-model-example", "示例：gpt-4o");
            });
        const keySetting = new Setting(container).setClass("smartmd-ai-model-setting")
            .setName(localize("模型密钥"))
            .setDesc(localize("直接覆盖即可更换，删空即可移除。HTTP 服务无需认证时可留空。"));
        const keyError = keySetting.descEl.createDiv({
            cls: "smartmd-ai-key-error",
            attr: { role: "alert" },
        });
        keySetting.addText((text) => {
            keyInput = text.inputEl;
            keyInput.type = "password";
            keyInput.autocomplete = "off";
            text.setPlaceholder(localize("未填写密钥")).setValue(this.runtime.store.getAIKey(shared.aiApiUrl)).onChange((value) => {
                const address = this.runtime.store.settings.shared.aiApiUrl;
                if (this.aiAddressShown !== address) {
                    this.scheduleDisplay();
                    return;
                }
                this.aiSettingsChanged();
                keyError.setText("");
                this.aiKeySave = this.aiKeySave.then(() => this.runtime.store.setAIKey(value.trim(), address)).then(() => ({ address, saved: true }), () => {
                    if (keyInput?.value === value && this.runtime.store.settings.shared.aiApiUrl === address) {
                        keyInput.value = this.runtime.store.getAIKey(address);
                        keyError.setText(localize("模型密钥未能保存至设备安全存储"));
                    }
                    return { address, saved: false };
                });
            });
            addExample(text.inputEl, "smartmd-ai-key-example", "示例：sk-…（服务要求认证时填写）");
        });
        const connection = new Setting(container)
            .setName(localize("模型连接"))
            .setDesc(localize("用内置示例测试选区改写和后续润色。两次请求均只使用示例文字。"));
        const result = connection.descEl.createDiv({ attr: { role: "status", "aria-atomic": "true" } });
        connection.addButton(button => {
            this.refreshAiTest = () => {
                result.setText(this.aiTestMessage);
                button.setDisabled(this.aiTestRunning).setButtonText(this.aiTestRunning ? localize("正在测试…") : localize("测试连接"));
            };
            this.refreshAiTest();
            button.onClick(() => this.testAiConnection());
        });
        new Setting(container)
            .setName(localize("AI 编辑器"))
            .setDesc(localize("在笔记编辑区右键，选择用 AI 编辑选中文字或整篇文档。"))
            .addButton((button) => button.setButtonText(localize("打开 AI 编辑器")).setCta().onClick(() => void this.runtime.openAI()));
    }

    private aboutSection(container: HTMLElement): void {
        new Setting(container).setClass("smartmd-settings-section-title").setName(localize("关于")).setHeading();
        const project = container.createDiv({ cls: "smartmd-about-section" });
        new Setting(project).setClass("smartmd-settings-subheading").setName(localize("项目介绍")).setHeading();
        project.createEl("p", { text: localize("SmartMD 是一款面向 Obsidian 的本地知识库增强插件，提供 Markdown 发布、多设备同步、桌面外部文件夹同步、AI 辅助编辑能力。各项功能相互隔离，单项功能失败不会触发其他功能删除或回滚。") });
        const principles = project.createEl("ul");
        for (const text of [
            "本地知识库始终由用户掌控。",
            "发布内容以本地 Markdown 为唯一正文来源。",
            "AI 修改必须经过用户确认。",
            "OneDrive 同步保留最新文件，同一修改时间以云端为准。",
            "项目将以公益和开源方向持续维护。",
        ]) principles.createEl("li", { text: localizeMessage(text) });

        const mainSite = container.createDiv({ cls: "smartmd-about-section" });
        new Setting(mainSite).setClass("smartmd-settings-subheading").setName(localize("SmartKeep 主站")).setHeading();
        mainSite.createEl("p", { text: localize("SmartKeep 是一个智能书签管理工具，可以帮助用户分类保存网页、整理资料、管理公开分享链接，并逐步形成个人网络知识入口。SmartKeep 负责书签和账号授权，SmartKeep MD 负责 Markdown 发布，两个站点不会自动读取彼此的私人内容。") });
        mainSite.createEl("button", { text: localize("访问 SmartKeep"), cls: "mod-cta" }).onclick = () => window.open("https://smartkeep.me", "_blank");

        const contact = container.createDiv({ cls: "smartmd-about-section" });
        new Setting(contact).setClass("smartmd-settings-subheading").setName(localize("联系管理员")).setHeading();
        contact.createEl("p", { text: localize("管理员邮箱：support@smartkeep.me。反馈问题时建议附上插件版本、设备类型和问题现象；涉及资料安全时请在标题中注明“数据安全问题”。") });
        contact.createEl("button", { text: localize("发送邮件") }).onclick = () => window.open("mailto:support@smartkeep.me?subject=SmartKeep%20Connect%20问题反馈", "_blank");

        const roadmap = container.createDiv({ cls: "smartmd-about-section" });
        new Setting(roadmap).setClass("smartmd-settings-subheading").setName(localize("未来更新计划")).setHeading();
        roadmap.createEl("p", { text: localize("以下方向仍处于规划阶段，不承诺具体上线时间：") });
        const list = roadmap.createEl("ul");
        for (const text of [
            "根据用户反馈考虑轻量会员服务。",
            "为会员提供限量图片托管。",
            "为会员提供限量官方 AI 编辑额度。",
            "优化大型知识库发布速度。",
            "改进知识图谱和阅读体验。",
            "增加更多兼容模型服务。",
            "根据服务器压力调整免费发布范围。",
        ]) list.createEl("li", { text: localizeMessage(text) });
        roadmap.createEl("p", { cls: "setting-item-description", text: localize("以上功能将根据用户数量、服务器压力和实际反馈决定是否开发。当前版本不会提前收费，也不会限制用户自备模型。") });
    }
}
