import { localize, localizeMessage } from "./localization";
import {
    App,
    Modal,
    Notice,
    requestUrl,
    Setting,
    TAbstractFile,
    TFile,
    TFolder,
} from "obsidian";
import { SmartMDStore } from "./core";
import { SmartKeepAuth } from "./smartkeep-auth";
import {
    PublishedDocument,
    PublishedSpace,
    PublishPendingOperation,
    PublishScope,
    type PublishRegion,
    SMARTMD_PUBLISH_SITES,
} from "./types";
import { OperationAuthority, RunState, RunTicket } from "./operation-authority";

type PublishPreview = {
    scope: PublishScope;
    files: TFile[];
    blocked: Array<{ path: string; images: string[] }>;
    conflicts: string[];
};

type PublishApiResult = {
    id?: string;
    code?: string;
    url?: string;
    documents?: PublishedDocument[];
};

const LOCAL_IMAGE = /!\[[^\]]*\]\((?!https?:\/\/)([^)]+)\)|!\[\[[^\]]+\]\]/gi;

function id(prefix: string): string {
    return `${prefix}_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 10)}`;
}

function normalize(path: string): string {
    return path.replace(/\\/g, "/").replace(/^\/+|\/+$/g, "");
}

function localImages(markdown: string): string[] {
    const found: string[] = [];
    let match: RegExpExecArray | null;
    LOCAL_IMAGE.lastIndex = 0;
    while ((match = LOCAL_IMAGE.exec(markdown)) !== null) found.push(match[1] || match[0]);
    return found;
}

function scopeContains(scope: PublishScope, path: string): boolean {
    const target = normalize(scope.path);
    const current = normalize(path);
    if (scope.type === "vault") return true;
    if (scope.type === "document") return current === target;
    return current === target || current.startsWith(`${target}/`);
}

function relativeToScope(scope: PublishScope, file: TFile): string {
    if (scope.type === "vault") return normalize(file.path);
    if (scope.type === "document") return file.name;
    const root = normalize(scope.path);
    return normalize(file.path).slice(root.length + 1);
}

async function confirm(app: App, title: string, body: string, confirmText: string): Promise<boolean> {
    return await new Promise((resolve) => {
        const modal = new ConfirmModal(app, title, body, confirmText, resolve);
        modal.open();
    });
}

export class ConfirmModal extends Modal {
    private settled = false;

    constructor(
        app: App,
        private titleText: string,
        private bodyText: string,
        private confirmText: string,
        private finish: (value: boolean) => void,
    ) {
        super(app);
    }

    onOpen(): void {
        this.modalEl.addClass("smartmd-modal-dialog");
        this.titleEl.setText(localizeMessage(this.titleText));
        this.contentEl.createEl("p", { text: localizeMessage(this.bodyText), cls: "smartkeep-modal-subdesc" });
        const actions = this.contentEl.createDiv({ cls: "smartmd-modal-actions" });
        const cancel = actions.createEl("button", { text: localize("取消"), attr: { type: "button" } });
        const ok = actions.createEl("button", { text: localizeMessage(this.confirmText), cls: "mod-cta", attr: { type: "button" } });
        cancel.onclick = () => { this.finishOnce(false); this.close(); };
        ok.onclick = () => { this.finishOnce(true); this.close(); };
    }

    onClose(): void {
        this.finishOnce(false);
        this.contentEl.empty();
    }

    private finishOnce(value: boolean): void {
        if (this.settled) return;
        this.settled = true;
        this.finish(value);
    }
}

export class PublishSetupModal extends Modal {
    private access: "public" | "password" = "public";
    private submitted = false;
    private closed = false;
    private released = false;

    constructor(
        app: App,
        private preview: PublishPreview,
        private submit: (value: { title: string; access: "public" | "password"; password: string }) => Promise<void>,
        private release: () => void,
    ) {
        super(app);
    }

    onOpen(): void {
        this.modalEl.addClass("smartmd-modal-dialog");
        this.titleEl.setText(localize("发布预览"));
        this.contentEl.addClass("smartmd-publish-modal");

        const summary = this.contentEl.createDiv({ cls: "smartmd-publish-summary" });
        const summaryHeader = summary.createDiv({ cls: "smartmd-publish-summary-header" });
        const countEl = summaryHeader.createDiv({ cls: "smartmd-publish-count" });
        countEl.createSpan({ text: localize("将发布 ") });
        countEl.createEl("strong", { text: `${this.preview.files.length}` });
        countEl.createSpan({ text: localize(" 篇 Markdown") });
        summaryHeader.createDiv({ cls: "smartmd-publish-badge", text: localize("确认前无云端变化") });
        summary.createEl("p", {
            cls: "smartmd-publish-summary-hint",
            text: localize("发布后，此范围内 Markdown 的新增、修改、重命名和删除会持续同步到原空间；子目录包含在内。公开链接会自动保存至 SmartKeep，密码内容不会生成公开书签。"),
        });

        if (this.preview.blocked.length) {
            const warning = this.contentEl.createDiv({ cls: "smartmd-warning" });
            warning.createEl("strong", { text: localize("有 {0} 篇包含本地图片，必须先替换为公开图片链接：", this.preview.blocked.length) });
            for (const item of this.preview.blocked.slice(0, 8)) {
                warning.createDiv({ text: item.path, cls: "smartmd-warning-item" });
            }
        }
        if (this.preview.conflicts.length) {
            this.contentEl.createEl("p", {
                cls: "smartmd-warning",
                text: localize("发现 {0} 个同路径内容。发布时只更新已确认范围，不会自动删除。", this.preview.conflicts.length),
            });
        }

        let titleInput: HTMLInputElement | null = null;
        new Setting(this.contentEl)
            .setName(localize("发布空间名称"))
            .setDesc(localize("在 SmartMD 与书签中展示的名称"))
            .addText((text) => {
                text.setValue(this.preview.scope.title).setPlaceholder(localize("例如：我的技术笔记"));
                titleInput = text.inputEl;
            });

        let passwordSetting: Setting | null = null;
        let passwordInput: HTMLInputElement | null = null;

        new Setting(this.contentEl)
            .setName(localize("访问权限"))
            .setDesc(localize("选择公开访问或密码保护"))
            .addDropdown((dropdown) => {
                dropdown.addOption("public", localize("公开访问（可被搜索引擎收录）"));
                dropdown.addOption("password", localize("密码访问（不允许搜索引擎收录）"));
                dropdown.setValue(this.access);
                dropdown.onChange((value) => {
                    this.access = value as "public" | "password";
                    if (passwordSetting) {
                        passwordSetting.settingEl.toggleClass("smartmd-publish-password-hidden", this.access !== "password");
                    }
                });
            });

        passwordSetting = new Setting(this.contentEl)
            .setName(localize("访问密码"))
            .setDesc(localize("至少需 6 个字符"))
            .addText((text) => {
                text.inputEl.type = "password";
                text.setPlaceholder(localize("输入访问密码"));
                passwordInput = text.inputEl;
            });
        passwordSetting.setClass("smartmd-publish-password-hidden");

        const actions = this.contentEl.createDiv({ cls: "smartmd-modal-actions" });
        const status = this.contentEl.createEl("p", { cls: "smartmd-modal-status" });

        const cancel = actions.createEl("button", { text: localize("取消"), attr: { type: "button" } });
        cancel.onclick = () => this.close();

        const publish = actions.createEl("button", { text: localize("确认发布"), cls: "mod-cta", attr: { type: "button" } });
        publish.disabled = this.preview.blocked.length > 0;

        publish.onclick = async () => {
            if (this.submitted) return;
            const titleValue = (titleInput?.value || "").trim();
            if (!titleValue) return void new Notice(localize("请填写发布空间名称"));
            const pwdValue = passwordInput?.value || "";
            if (this.access === "password" && pwdValue.length < 6) {
                return void new Notice(localize("访问密码至少需要 6 个字符"));
            }
            this.submitted = true;
            publish.disabled = true;
            cancel.disabled = true;
            publish.setText(localize("正在发布..."));
            status.setText(localize("正在创建发布空间并上传 Markdown，请稍候。"));
            try {
                await this.submit({ title: titleValue, access: this.access, password: pwdValue });
                this.close();
            } catch (error: any) {
                status.setText(localize("发布失败：{0}", localizeMessage(String(error?.message || error))));
                status.addClass("smartmd-warning");
                cancel.disabled = false;
                publish.disabled = false;
                publish.setText(localize("确认发布"));
            } finally {
                this.submitted = false;
                if (this.closed) this.releaseOnce();
            }
        };
    }

    onClose(): void {
        this.closed = true;
        if (!this.submitted) this.releaseOnce();
        this.contentEl.empty();
    }

    private releaseOnce(): void {
        if (this.released) return;
        this.released = true;
        this.release();
    }
}

export class PublishManagerModal extends Modal {
    private unsubscribe: (() => void) | null = null;
    private loading = true;

    constructor(app: App, private service: PublishService) {
        super(app);
    }

    async onOpen(): Promise<void> {
        this.titleEl.setText(localize("账号云端空间（所有知识库）"));
        this.contentEl.createEl("p", { text: localize("正在加载发布状态…") });
        this.unsubscribe = this.service.subscribeStatus(() => void this.render());
        void this.service.refreshSpaces().finally(() => {
            this.loading = false;
            void this.render();
        });
    }

    onClose(): void {
        this.unsubscribe?.();
        this.unsubscribe = null;
        this.contentEl.empty();
    }

    private async render(): Promise<void> {
        this.contentEl.empty();
        if (this.loading) {
            this.contentEl.createEl("p", { text: localize("正在加载发布状态…") });
            return;
        }
        const refreshActions = this.contentEl.createDiv({ cls: "smartmd-actions" });
        const refresh = refreshActions.createEl("button", { text: localize("重新对账"), cls: "mod-cta" });
        refresh.onclick = async () => {
            refresh.disabled = true;
            refresh.setText(localize("正在对账..."));
            await this.service.refreshSpaces();
            await this.render();
        };
        if (this.service.hasPendingCreates()) {
            this.contentEl.createEl("p", { text: localize("有首次发布尚未完成。请按原名称、访问方式和范围重新发起发布并确认，将复用原业务凭证；重新对账不会自动上传。"), cls: "smartmd-warning" });
        }
        this.contentEl.createEl("p", { text: localize("空间属于 SmartKeep 账号。所有知识库共用最多 3 个名额；撤回不需要本地原稿，更新正文仍需关联当前知识库。") });
        const count = this.service.accountSpaceCount();
        this.contentEl.createEl("p", { text: count === null ? localize("账号空间：待查询") : localize("账号空间：{0}/3（所有知识库共用）", count) });
        const isCurrent = this.service.sessionGuard();
        const spaces = this.service.accountSpaces();
        if (!spaces.length) {
            this.contentEl.createEl("p", { text: localize("暂无发布空间。请在文档或文件夹菜单中选择发布。") });
            return;
        }
        for (const space of spaces) {
            const canAct = () => isCurrent() && this.service.accountSpaces().includes(space);
            const card = this.contentEl.createDiv({ cls: "smartmd-card" });
            card.createEl("h3", { text: space.title });
            card.createEl("p", {
                text: localize("{0} · {1} · {2} 篇已发布 · {3} 篇待同步 · {4} 条待处理", this.service.scopeLabel(space), space.access === "public" ? localize("公开") : localize("密码访问"), space.documents.length, space.pendingDocuments.length, space.feedbackOpen || 0),
            });
            card.createEl("p", { text: localize("最后更新：{0}", space.lastPublishedAt ? new Date(space.lastPublishedAt).toLocaleString() : localize("尚未上传正文")) });
            card.createEl("p", {
                text: space.scopeKnown === false && space.operationStatus === "needs_confirmation"
                    ? localize("发布状态：云端空间可管理；更新正文需关联本地文件")
                    : localize("发布状态：{0}{1}", localizeMessage(this.service.statusLabel(space)), localizeMessage(space.statusMessage ? ` · ${space.statusMessage}` : "")),
                cls: space.operationStatus === "failed" || (space.scopeKnown !== false && space.operationStatus === "needs_confirmation") ? "smartmd-warning" : "",
            });
            card.createEl("p", { text: localize("SmartKeep 书签：{0}", this.service.bookmarkLabel(space)) });
            card.createEl("p", {
                text: this.service.canRun(space)
                    ? localize("本地关联：已确认当前知识库，可明确更新正文或撤回云端空间。")
                    : localize("本地关联：尚未关联当前知识库。可查看反馈或撤回；关联并确认本地文件前不能更新正文。"),
            });
            if (space.url) {
                const link = card.createEl("a", { text: localize("打开发布页面"), href: space.url });
                link.target = "_blank";
            }
            const actions = card.createDiv({ cls: "smartmd-actions" });
            if (this.service.isBusy(space)) {
                actions.createEl("button", { text: localizeMessage(this.service.statusLabel(space)) }).disabled = true;
                continue;
            }
            if (space.operationStatus === "failed" && space.pendingOperation && (space.pendingOperation.action === "withdraw" ? this.service.canWithdraw(space) : this.service.canRun(space))) {
                const label = space.pendingOperation.action === "update"
                    ? "确认用当前本地内容重试更新"
                    : space.pendingOperation.action === "withdraw" ? "确认重试撤回" : "确认重试访问设置";
                actions.createEl("button", { text: localizeMessage(label), cls: "mod-cta" }).onclick = async () => {
                    if (!canAct()) return;
                    await this.service.retryOperation(space.id);
                    await this.render();
                };
            }
            if (this.service.canWithdraw(space) && !space.pendingOperation) {
                actions.createEl("button", { text: localize("撤回空间") }).onclick = async () => {
                    if (!canAct() || !this.service.reserveConfirmation(space.id, "withdraw")) return;
                    const approved = await confirm(this.app, localize("撤回发布空间"), localize("确认撤回账号空间“{0}”？\n{1}\n其中 {2} 篇云端正文和留言将不再可访问，并释放 1 个账号名额；任何知识库中的本地文件都不会改变。", space.title, space.url, space.documents.length), localize("确认撤回"));
                    if (!canAct()) return;
                    if (approved) await this.service.withdrawSpace(space.id, false, true);
                    else await this.service.cancelReservedConfirmation(space.id);
                    await this.render();
                };
            }
            if (this.service.hasBlockingFailure()) continue;
            if (space.operationStatus === "failed") {
                continue;
            }
            if (!this.service.canRun(space)) {
                if (this.service.canConfirm(space)) {
                    actions.createEl("button", { text: localize("确认账号与范围并恢复"), cls: "mod-cta" }).onclick = async () => {
                        if (!canAct()) return;
                        await this.service.confirmLegacySpace(space.id);
                        await this.render();
                    };
                }
                continue;
            }
            actions.createEl("button", { text: localize("复制链接") }).onclick = async () => {
                if (!space.url) return;
                await navigator.clipboard.writeText(space.url);
                new Notice(localize("发布链接已复制"));
            };
            actions.createEl("button", { text: localize("立即同步到云端") }).onclick = async () => {
                if (!canAct()) return;
                await this.service.updateSpace(space.id, true, true);
                await this.render();
            };
            if (space.pendingDocuments.length) {
                actions.createEl("button", { text: localize("立即同步新增文档"), cls: "mod-cta" }).onclick = async () => {
                    if (!canAct()) return;
                    await this.service.approvePending(space.id);
                    await this.render();
                };
            }
            if (space.pendingDeletes.length) {
                actions.createEl("button", { text: localize("同步删除 {0} 篇", space.pendingDeletes.length) }).onclick = async () => {
                    if (!canAct()) return;
                    await this.service.confirmPendingDeletes(space.id);
                    await this.render();
                };
            }
            actions.createEl("button", { text: space.access === "public" ? localize("改为密码访问") : localize("改为公开") }).onclick = async () => {
                if (!canAct()) return;
                await this.service.toggleAccess(space.id);
                await this.render();
            };
            if (["detached", "user_removed"].includes(space.bookmarkStatus)) {
                actions.createEl("button", { text: localize("重新生成书签") }).onclick = async () => {
                    if (!canAct()) return;
                    await this.service.retryBookmark(space.id);
                    await this.render();
                };
            }

        }
    }
}

class FeedbackModal extends Modal {
    private deleting = new Set<string>();
    constructor(app: App, private service: PublishService, private space: PublishedSpace, private feedback: any[], private isCurrent: () => boolean, private documentTitle?: string) {
        super(app);
    }

    onOpen(): void {
        this.titleEl.setText(localize("{0} · 问题反馈", this.documentTitle || this.space.title));
        this.render();
    }

    private render(): void {
        this.contentEl.empty();
        const sorted = [...this.feedback].sort((a, b) => (a.status === "open" ? -1 : 1) - (b.status === "open" ? -1 : 1));
        if (!sorted.length) return void this.contentEl.createEl("p", { text: localize("暂无问题反馈。") });
        for (const item of sorted) {
            const card = this.contentEl.createDiv({ cls: "smartmd-card" });
            card.createEl("strong", { text: localizeMessage(`${item.authorName} · ${item.status === "open" ? "待处理" : item.status === "resolved" ? "已解决" : "不采纳"}`) });
            card.createEl("p", { text: localize("文档：{0}", item.documentPath || localize("文档路径未知")) });
            card.createEl("blockquote", { text: item.quote });
            card.createEl("p", { text: item.body });
            if (item.resolution) card.createEl("p", { text: localize("所有者结论：{0}", item.resolution) });
            const actions = card.createDiv({ cls: "smartmd-actions" });
            const remove = actions.createEl("button", { text: localize("删除意见"), cls: "mod-warning" });
            remove.disabled = this.deleting.has(item.id);
            remove.onclick = async () => {
                if (this.deleting.has(item.id)) return;
                this.deleting.add(item.id);
                remove.disabled = true;
                try {
                    const approved = await confirm(this.app, localize("删除意见"), localize("确认删除这条修改意见？删除后无法恢复，文章正文和其他意见不受影响。"), localize("确认删除"));
                    if (!approved || !this.isCurrent()) return;
                    await this.service.deleteFeedback(this.space.id, item.id);
                    this.feedback = this.feedback.filter(feedback => feedback.id !== item.id);
                    this.render();
                    new Notice(localize("意见已删除"));
                } catch {
                    new Notice(localize("删除未确认成功，意见仍保留在列表中，请重试。"), 7000);
                } finally {
                    this.deleting.delete(item.id);
                    this.render();
                }
            };
            if (item.status !== "open") continue;
            actions.createEl("button", { text: localize("标记已解决"), cls: "mod-cta" }).onclick = async () => {
                if (!this.isCurrent()) return;
                await this.service.setFeedbackStatus(this.space.id, item.id, "resolved", "已在本地处理并同步最新版");
                item.status = "resolved";
                item.resolution = "已在本地处理并同步最新版";
                this.render();
            };
            actions.createEl("button", { text: localize("不采纳") }).onclick = async () => {
                const reason = window.prompt(localize("请填写不采纳的简短理由"));
                if (!reason?.trim()) return;
                if (!this.isCurrent()) return;
                await this.service.setFeedbackStatus(this.space.id, item.id, "rejected", reason.trim());
                item.status = "rejected";
                item.resolution = reason.trim();
                this.render();
            };
        }
    }
}

export class PublishService {
    private updateTimers = new Map<string, number>();
    private deferredUpdates = new Set<PublishedSpace>();
    private activePreviews = new Set<string>();
    private updateJobs = new Map<string, Promise<void>>();
    private creating = false;
    private accountSpaceIds = new Set<string>();
    private accountSpaceAccount = "";
    private authority = new OperationAuthority();
    private reconciliationJob: Promise<void> | null = null;
    private reconciliationEpoch = 0;
    private resuming = false;
    private signingOut = false;
    private disposed = false;
    private pendingAiApprovalPaths = new Map<string, symbol | null>();
    private feedbackSummary = new Map<string, { total: number; open: number }>();
    private feedbackAccount = "";
    private feedbackEpoch = 0;
    private feedbackJob: Promise<void> | null = null;
    private feedbackChanged: (() => void) | null = null;
    private feedbackWrites: Promise<void> = Promise.resolve();

    constructor(private app: App, private store: SmartMDStore, private auth: SmartKeepAuth) {
        const saved = (this.store as any).settings?.local?.pendingAiApprovalPaths;
        if (Array.isArray(saved)) {
            for (const p of saved) {
                if (typeof p === "string" && p) this.pendingAiApprovalPaths.set(normalize(p), null);
            }
        }
    }

    setFeedbackChangeCallback(callback: (() => void) | null): void { this.feedbackChanged = callback; }

    private clearFeedbackSummary(): void {
        this.feedbackEpoch++;
        this.feedbackJob = null;
        this.feedbackAccount = "";
        this.feedbackSummary.clear();
        this.feedbackChanged?.();
    }

    private applyFeedbackSummary(account: string, rows: any[], spaceId?: string): void {
        if (account !== String((this.store as any).settings.local.smartkeepAccountId || "") || !this.auth.isSignedIn()) return;
        if (spaceId) this.feedbackEpoch++; // Invalidate any poll started before the mutation response.
        if (!spaceId || this.feedbackAccount !== account) this.feedbackSummary.clear();
        if (spaceId) for (const key of this.feedbackSummary.keys()) if (key.startsWith(`${spaceId}:`)) this.feedbackSummary.delete(key);
        this.feedbackAccount = account;
        for (const row of rows) this.feedbackSummary.set(`${row.spaceId}:${row.documentId}`, { total: Number(row.total || 0), open: Number(row.open || 0) });
        this.feedbackChanged?.();
    }

    async refreshFeedbackSummary(): Promise<void> {
        if (this.disposed || this.signingOut) return;
        const account = String((this.store as any).settings.local.smartkeepAccountId || "");
        if (!account || !this.auth.isSignedIn()) { this.clearFeedbackSummary(); return; }
        if (this.feedbackAccount && this.feedbackAccount !== account) this.clearFeedbackSummary();
        if (this.feedbackJob) return await this.feedbackJob;
        const epoch = this.feedbackEpoch;
        const job = (async () => {
            try {
                const result = await this.api<any>("GET", "/api/v1/feedback/summary", undefined, [], undefined, true);
                if (epoch === this.feedbackEpoch) this.applyFeedbackSummary(account, result.documents || []);
            } catch { /* Keep the last confirmed snapshot for the same account; never treat failure as zero. */ }
        })();
        this.feedbackJob = job;
        try { await job; } finally { if (this.feedbackJob === job) this.feedbackJob = null; }
    }

    feedbackForPath(path: string): { spaceId: string; documentId: string; total: number; open: number } | null {
        if (!this.auth.isSignedIn() || this.feedbackAccount !== String((this.store as any).settings.local.smartkeepAccountId || "")) return null;
        const target = this.publishedLink(path);
        if (!target?.documentPublished) return null;
        const doc = this.spaces().find(space => space.id === target.spaceId)?.documents.find(item => normalize(item.localPath) === normalize(path));
        if (!doc) return null;
        const counts = this.feedbackSummary.get(`${target.spaceId}:${doc.id}`);
        return counts && counts.total > 0 ? { spaceId: target.spaceId, documentId: doc.id, ...counts } : null;
    }

    async openDocumentFeedback(file: TFile): Promise<void> {
        const target = this.feedbackForPath(file.path);
        if (!target) return void new Notice(localize("此文档暂无问题反馈。"));
        await this.openFeedback(target.spaceId, target.documentId);
    }

    private async persistAiPending(): Promise<void> {
        const local = (this.store as any).settings?.local;
        if (local) {
            local.pendingAiApprovalPaths = this.getAiPendingPaths();
            await (this.store as any).save?.();
        }
    }

    isAiPending(filePath: string): boolean {
        return this.pendingAiApprovalPaths.has(normalize(filePath));
    }

    markAiPending(filePath: string): { revision: symbol; persisted: Promise<void> } {
        const norm = normalize(filePath);
        const revision = Symbol();
        this.pendingAiApprovalPaths.set(norm, revision);
        for (const space of this.spaces()) {
            if (scopeContains(space.scope, norm)) {
                const timer = this.updateTimers.get(space.id);
                if (timer) {
                    window.clearTimeout(timer);
                    this.updateTimers.delete(space.id);
                }
                space.operationStatus = "pending_approval";
                space.statusMessage = "文档包含未核准的 AI 改动，自动公网发布已暂停。";
            }
        }
        return { revision, persisted: this.persistAiPending() };
    }

    async clearAiPending(filePath: string, revision: symbol): Promise<boolean> {
        const norm = normalize(filePath);
        if (this.pendingAiApprovalPaths.get(norm) !== revision) return false;
        this.pendingAiApprovalPaths.delete(norm);
        try {
            await this.persistAiPending();
            return !this.pendingAiApprovalPaths.has(norm);
        } catch {
            if (!this.pendingAiApprovalPaths.has(norm)) this.pendingAiApprovalPaths.set(norm, revision);
            const local = (this.store as any).settings?.local;
            if (local) local.pendingAiApprovalPaths = this.getAiPendingPaths();
            return false;
        }
    }

    getAiPendingPaths(): string[] {
        return Array.from(this.pendingAiApprovalPaths.keys());
    }

    private publishTombstones(): Record<string, number> {
        const shared = (this.store as any).settings.shared;
        shared.publishTombstones = shared.publishTombstones || {};
        return shared.publishTombstones;
    }

    sessionGuard(): () => boolean {
        const epoch = this.reconciliationEpoch;
        const region = this.auth.activeRegion();
        const account = this.store.settings.local.smartkeepAccountId;
        return () => !this.disposed && !this.signingOut && this.auth.isSignedIn() &&
            epoch === this.reconciliationEpoch && region === this.auth.activeRegion() &&
            account === this.store.settings.local.smartkeepAccountId;
    }

    subscribeStatus(listener: () => void): () => void {
        return (this.store as any).subscribe?.(listener) || (() => {});
    }

    private publishIdentities(space: any): string[] {
        return Array.from(new Set([
            space?.id,
            space?.localId,
            space?.remoteId,
            space?.code,
        ].map((value) => String(value || "")).filter(Boolean)
            .map(value => this.auth.activeRegion() === "me" ? "me:" + value : value)));
    }

    private markWithdrawn(space: any): void {
        const timestamp = Date.now();
        for (const identity of this.publishIdentities(space)) {
            this.publishTombstones()[identity] = timestamp;
        }
    }

    private clearWithdrawn(space: any): void {
        for (const identity of this.publishIdentities(space)) {
            delete this.publishTombstones()[identity];
        }
    }

    canRun(space: PublishedSpace): boolean {
        return !this.publishWritesPaused() && this.hasRunPermission(space);
    }

    private publishWritesPaused(): boolean {
        return this.disposed || this.signingOut || this.resuming;
    }

    private hasRunPermission(space: PublishedSpace): boolean {
        const local = (this.store as any).settings.local;
        return (space.publishRegion ?? "cn") === this.auth.activeRegion()
            && this.authority.canRun(space as PublishedSpace & RunState)
            && (space as any).confirmationScope === `${space.scope.type}:${space.scope.path}`
            && (space as any).confirmationAccount === String(local.smartkeepAccountId || "");
    }

    accountSpaces(): PublishedSpace[] {
        const account = String((this.store as any).settings.local.smartkeepAccountId || "");
        if (!this.auth.isSignedIn() || !account || account !== this.accountSpaceAccount) return [];
        return this.spaces().filter((space) => this.accountSpaceIds.has(space.id));
    }

    accountSpaceCount(): number | null {
        const account = String((this.store as any).settings.local.smartkeepAccountId || "");
        return this.auth.isSignedIn() && account && account === this.accountSpaceAccount
            ? this.accountSpaceIds.size : null;
    }

    canWithdraw(space: PublishedSpace): boolean {
        return !this.publishWritesPaused() && this.accountSpaces().includes(space);
    }

    isBusy(space: PublishedSpace): boolean {
        return ["preparing", "submitting", "reconciling"].includes(space.operationStatus || "");
    }

    hasBlockingFailure(): boolean {
        return this.spaces().some((space) => space.operationStatus === "failed");
    }

    reserveConfirmation(spaceId: string, action: "withdraw"): boolean {
        const space = this.spaces().find((item) => item.id === spaceId);
        if (!space || !this.canWithdraw(space) || this.isBusy(space) || space.pendingOperation) return false;
        space.operationStatus = "preparing";
        space.statusMessage = "正在等待用户确认";
        space.pendingOperation = {
            action,
            requestId: id(action),
            expectedVersion: Number(space.cloudVersion || 0),
        };
        void (this.store as any).save();
        return true;
    }

    async cancelReservedConfirmation(spaceId: string): Promise<void> {
        const space = this.spaces().find((item) => item.id === spaceId);
        if (!space || space.operationStatus !== "preparing") return;
        space.pendingOperation = undefined;
        space.operationStatus = "success";
        space.statusMessage = "已取消，本次未产生云端变化";
        await (this.store as any).save();
    }

    canConfirm(space: PublishedSpace): boolean {
        return !this.publishWritesPaused() && space.scopeKnown !== false && space.reconciliationStatus !== "remote_withdrawn";
    }

    statusLabel(space: PublishedSpace): string {
        const labels: Record<string, string> = {
            pending_approval: "待发布确认",
            preparing: "准备中",
            submitting: "正在提交",
            reconciling: "正在对账",
            success: "成功",
            needs_confirmation: "需确认",
            failed: "失败",
        };
        return labels[space.operationStatus || "needs_confirmation"] || "需确认";
    }

    async pauseLegacySpaces(): Promise<void> {
        for (const space of this.spaces()) {
            if ((space as any).confirmed !== true || !Number(space.cloudVersion || 0)) {
                await this.authority.isolate(space as PublishedSpace & RunState);
                space.operationStatus = "needs_confirmation";
                space.reconciliationStatus = "pending_confirmation";
                space.statusMessage = "旧发布记录需要先与云端确认";
            }
        }
        for (const timer of this.updateTimers.values()) window.clearTimeout(timer);
        this.updateTimers.clear();
        await (this.store as any).save();
    }

    async confirmLegacySpace(spaceId: string): Promise<void> {
        const space = this.spaces().find((item) => item.id === spaceId);
        if (!space || this.canRun(space)) return;
        if (!this.canConfirm(space)) throw new Error("该记录缺少可信的本地发布范围，只能保持只读待确认");
        if (!this.auth.isSignedIn()) throw new Error("请先登录发布空间所属的 SmartKeep 账号");
        const files = this.app.vault.getMarkdownFiles().filter((file) => scopeContains(space.scope, file.path));
        if (!files.length) throw new Error("原发布范围当前不存在，无法解除升级保护");
        const scope = space.scope.type === "vault" ? "整个知识库" : space.scope.path;
        const isCurrent = this.sessionGuard();
        const approved = await confirm(
            this.app,
            localize("确认旧发布空间"),
            localize("请确认当前账号拥有“{0}”，并确认继续使用原范围：{1}。确认后恢复此范围 Markdown 新增、修改和删除的持续同步。", space.title, scope),
            localize("确认并恢复"),
        );
        if (!approved || !isCurrent() || !this.spaces().includes(space)) return;
        this.authority.confirm(space as PublishedSpace & RunState);
        (space as any).confirmationScope = `${space.scope.type}:${space.scope.path}`;
        (space as any).confirmationAccount = (this.store as any).settings.local.smartkeepAccountId || "";
        space.operationStatus = "success";
        space.reconciliationStatus = "matched";
        space.statusMessage = "已确认云端状态，持续同步已恢复";
        await (this.store as any).save();
        this.checkContinuousSync();
    }

    dispose(): void {
        this.disposed = true;
        this.deferredUpdates.clear();
        this.clearFeedbackSummary();
        void this.authority.shutdown();
        this.creating = false;
        for (const timer of this.updateTimers.values()) window.clearTimeout(timer);
        this.updateTimers.clear();
    }

    async stopForSignOut(): Promise<void> {
        this.signingOut = true;
        this.resuming = false;
        this.deferredUpdates.clear();
        this.reconciliationEpoch++;
        this.reconciliationJob = null;
        this.accountSpaceIds.clear();
        this.accountSpaceAccount = "";
        this.clearFeedbackSummary();
        for (const timer of this.updateTimers.values()) window.clearTimeout(timer);
        this.updateTimers.clear();
        const stopped = this.authority.suspend();
        for (const space of this.spaces()) {
            (space as any).runAllowed = false;
            (space as any).confirmed = false;
        }
        await (this.store as any).save();
        await stopped;
        await (this.store as any).save();
    }

    async resumeAfterSignIn(): Promise<void> {
        if (this.disposed) return;
        this.signingOut = false;
        this.resuming = true;
        const epoch = this.reconciliationEpoch;
        this.authority.resume();
        try {
            await this.reconcileSpaces();
            if (epoch === this.reconciliationEpoch) await this.refreshFeedbackSummary();
        } finally {
            if (epoch === this.reconciliationEpoch) this.resuming = false;
        }
        if (this.disposed || epoch !== this.reconciliationEpoch) return;
        for (const space of this.deferredUpdates) {
            if (this.spaces().includes(space)) this.scheduleContinuousSync(space);
        }
        this.deferredUpdates.clear();
    }

    hasPendingCreates(): boolean {
        const local = this.store.settings.local;
        const region = this.auth.activeRegion();
        return Object.values((local as any).pendingPublishCreates || {}).some((item: any) =>
            (item.publishRegion ?? "cn") === region && (!item.account || item.account === local.smartkeepAccountId),
        );
    }

    spaces(): PublishedSpace[] {
        const region = this.auth.activeRegion();
        return this.allSpaces().filter(space => (space.publishRegion ?? "cn") === region);
    }

    private allSpaces(): PublishedSpace[] {
        return ((this.store as any).settings?.shared?.publishSpaces || []) as PublishedSpace[];
    }

    private rememberSpace(space: PublishedSpace): void {
        const spaces = this.allSpaces();
        const index = spaces.findIndex(item => item.id === space.id &&
            (item.publishRegion ?? "cn") === (space.publishRegion ?? "cn"));
        if (index >= 0) spaces.splice(index, 1, space);
        else spaces.push(space);
    }

    private forgetSpace(space: PublishedSpace): void {
        const spaces = this.allSpaces();
        const index = spaces.indexOf(space);
        if (index >= 0) spaces.splice(index, 1);
    }

    openManager(): void {
        new PublishManagerModal(this.app, this).open();
    }

    scopeLabel(space: PublishedSpace): string {
        return localize(space.scope.type === "vault" ? "整个知识库" : space.scope.type === "folder" ? "文件夹" : "单篇文档");
    }

    bookmarkLabel(space: PublishedSpace): string {
        const labels: Record<string, string> = {
            pending: "等待同步",
            synced: "已自动保存",
            detached: "已由用户移动",
            user_removed: "已由用户移除",
            not_applicable: "密码内容不生成",
            failed: "同步失败，等待重试",
        };
        const label = labels[space.bookmarkStatus] || "等待同步";
        return space.bookmarkError ? localize("{0}（{1}）", localizeMessage(label), localizeMessage(space.bookmarkError)) : localizeMessage(label);
    }

    async refreshSpaces(): Promise<void> {
        if (this.resuming) {
            await this.reconciliationJob;
            return;
        }
        await this.reconcileSpaces();
    }

    private async reconcileSpaces(): Promise<void> {
        if (this.disposed || this.signingOut || !this.auth.isSignedIn()) return;
        if (this.reconciliationJob) return await this.reconciliationJob;
        const account = String((this.store as any).settings.local.smartkeepAccountId || "");
        const epoch = this.reconciliationEpoch;
        const isCurrent = () => !this.disposed && !this.signingOut && epoch === this.reconciliationEpoch
            && account === String((this.store as any).settings.local.smartkeepAccountId || "");
        const job = (async () => {
          try {
            await this.reconcilePendingCreates(isCurrent);
            if (!isCurrent()) return;
            for (const space of [...this.spaces()]) {
                const credentialState = await this.resolvePendingOperation(space, isCurrent);
                if (!isCurrent()) return;
                if (
                    credentialState === "not_committed"
                    && space.pendingOperation
                    && space.operationStatus === "failed"
                ) {
                    if (space.pendingOperation.action === "access") {
                        space.pendingOperation = undefined;
                        space.operationStatus = "success";
                        space.statusMessage = "云端确认原访问方式变更未成立，可以重新操作";
                    } else {
                        space.statusMessage = "云端确认原操作未提交；请明确确认后重试，对账不会自动发布当前本地内容。";
                    }
                }
            }
            const result = await this.api<any>("GET", "/api/v1/spaces");
            if (!isCurrent()) return;
            if (!account || !this.auth.isSignedIn()) {
                throw new Error("发布账号已变化，请重新查询账号空间");
            }
            const remotes = Array.isArray(result.spaces) ? result.spaces : [];
            const remoteById = new Map(remotes.map((item: any) => [String(item.id), item]));
            for (const local of [...this.spaces()]) {
                const remote: any = remoteById.get(local.id);
                if (!remote) {
                    if (!this.spaces().includes(local)) continue;
                    await this.authority.isolate(local as PublishedSpace & RunState);
                    if (!isCurrent()) return;
                    local.reconciliationStatus = "remote_withdrawn";
                    local.operationStatus = "needs_confirmation";
                    local.statusMessage = "云端已撤回，本地仅保留只读记录";
                    this.markWithdrawn(local);
                    continue;
                }
                remoteById.delete(local.id);
                this.clearWithdrawn({ ...local, ...remote });
                const previousVersion = Number(local.cloudVersion || 0);
                const remoteVersion = Number(remote.version ?? remote.updatedAt ?? 0);
                const previousDocuments = [...local.documents];
                const knownById = new Map(previousDocuments.map((doc) => [doc.id, doc]));
                const knownByPath = new Map(previousDocuments.map((doc) => [normalize(doc.publicPath), doc]));
                let hasUnknownDocument = false;
                local.documents = (remote.documents || []).map((doc: any) => {
                    const known = knownById.get(String(doc.id)) || knownByPath.get(normalize(String(doc.path || doc.publicPath || "")));
                    if (!known) hasUnknownDocument = true;
                    return {
                        id: String(doc.id),
                        localPath: known?.localPath || this.localPathForRemote(local, String(doc.path || doc.publicPath || "")),
                        publicPath: String(doc.path || doc.publicPath || ""),
                        modifiedAt: Number(doc.modifiedAt || doc.modified_at || 0),
                        url: doc.url,
                    };
                });
                local.title = remote.title || local.title;
                local.code = remote.code || local.code;
                local.url = remote.url || local.url;
                local.access = remote.access || local.access;
                local.feedbackOpen = Number(remote.feedbackOpen || 0);
                local.bookmarkStatus = remote.bookmarkStatus || local.bookmarkStatus || "pending";
                local.bookmarkError = remote.bookmarkError || "";
                local.lastPublishedAt = Number(remote.updatedAt || local.lastPublishedAt || 0);
                local.cloudVersion = remoteVersion;
                const versionDrift = previousVersion > 0 && remoteVersion > 0 && previousVersion !== remoteVersion;
                if (versionDrift || hasUnknownDocument) {
                    await this.authority.isolate(local as PublishedSpace & RunState);
                    if (!isCurrent()) return;
                    local.reconciliationStatus = "pending_confirmation";
                    local.operationStatus = "needs_confirmation";
                    local.statusMessage = versionDrift
                        ? "云端已有其他设备提交的新版本，已刷新并停止旧状态操作"
                        : "云端包含本地未识别内容，等待人工确认";
                } else if (!local.documents.length && previousDocuments.length) {
                    await this.authority.isolate(local as PublishedSpace & RunState);
                    if (!isCurrent()) return;
                    local.reconciliationStatus = "needs_restore";
                    local.operationStatus = "needs_confirmation";
                    local.statusMessage = "云端当前没有正文，如需恢复必须人工确认";
                } else if (!this.isBusy(local) && !(local.operationStatus === "failed" && local.pendingOperation)) {
                    local.reconciliationStatus = "matched";
                    const canRun = this.hasRunPermission(local);
                    local.operationStatus = canRun ? "success" : "needs_confirmation";
                    local.statusMessage = canRun ? "本地记录与云端公开状态一致" : "云端状态已确认，等待确认本地范围";
                }
            }
            for (const remote of remoteById.values() as any) {
                this.clearWithdrawn(remote);
                this.rememberSpace({
                    publishRegion: this.auth.activeRegion(),
                    id: String(remote.id),
                    code: String(remote.code || ""),
                    title: String(remote.title || "未命名发布空间"),
                    url: String(remote.url || ""),
                    access: remote.access === "password" ? "password" : "public",
                    scope: {
                        type: remote.scopeType === "vault" ? "vault" : remote.scopeType === "folder" ? "folder" : "document",
                        path: "",
                        title: String(remote.title || "未识别发布范围"),
                    },
                    documents: (remote.documents || []).map((doc: any) => ({
                        id: String(doc.id),
                        localPath: String(doc.path || ""),
                        publicPath: String(doc.path || ""),
                        modifiedAt: Number(doc.modifiedAt || doc.modified_at || 0),
                        url: doc.url,
                    })),
                    pendingDocuments: [],
                    pendingDeletes: [],
                    feedbackOpen: Number(remote.feedbackOpen || 0),
                    bookmarkStatus: remote.bookmarkStatus || "pending",
                    bookmarkError: remote.bookmarkError || "",
                    lastPublishedAt: Number(remote.updatedAt || 0),
                    cloudVersion: Number(remote.version ?? remote.updatedAt ?? 0),
                    reconciliationStatus: "pending_confirmation",
                    operationStatus: "needs_confirmation",
                    statusMessage: "云端存在本地未识别记录，已保持只读",
                    scopeKnown: false,
                    confirmed: false,
                    runAllowed: false,
                    confirmationScope: "",
                    confirmationAccount: "",
                } as PublishedSpace);
            }
            this.accountSpaceIds = new Set<string>(remotes.map((remote: any) => String(remote.id)));
            this.accountSpaceAccount = account;
            await (this.store as any).save();
          } catch (error: any) {
            if (!isCurrent()) return;
            for (const space of this.spaces()) {
                if (this.isBusy(space)) continue;
                space.operationStatus = "failed";
                space.statusMessage = `对账失败，上一正常版本保持不变：${String(error?.message || error)}`;
            }
            await (this.store as any).save();
          }
        })();
        this.reconciliationJob = job;
        try {
            await job;
        } finally {
            if (this.reconciliationJob === job) this.reconciliationJob = null;
        }
    }

    publishedLink(path: string): { url: string; documentPublished: boolean; spaceId: string } | null {
        const account = String((this.store as any).settings.local.smartkeepAccountId || "");
        const target = normalize(path);
        if (!account || !this.auth.isSignedIn()) return null;
        const base = this.auth.serverBase();
        const candidates = this.spaces().filter((space) =>
            space.scopeKnown !== false && space.reconciliationStatus !== "remote_withdrawn"
            && Boolean(space.code) && space.documents.length > 0
            && (space as any).confirmationAccount === account
            && (space as any).confirmationScope === `${space.scope.type}:${space.scope.path}`
            && scopeContains(space.scope, target),
        );
        // A committed document has a real article URL. Folders/new entries open
        // the existing space and its directory, never a fabricated folder route.
        for (const space of candidates) {
            const document = space.documents.find((doc) => normalize(doc.localPath) === target);
            if (document) {
                const encoded = normalize(document.publicPath).split("/").map(encodeURIComponent).join("/");
                return { url: `${base}/s/${encodeURIComponent(space.code)}/${encoded}`, documentPublished: true, spaceId: space.id };
            }
        }
        const space = candidates.sort((a, b) => normalize(b.scope.path).length - normalize(a.scope.path).length)[0];
        return space ? { url: `${base}/s/${encodeURIComponent(space.code)}/`, documentPublished: false, spaceId: space.id } : null;
    }

    openPublishedLink(file: TAbstractFile): void {
        const target = this.publishedLink(file.path);
        if (!target) return void new Notice(localize("此范围没有可用发布链接，请刷新发布管理后重试。"));
        if (file instanceof TFile && !target.documentPublished) {
            new Notice(localize("此文件尚未单独发布，已打开所属发布空间。新增正文正在等待自动同步。"));
        }
        window.open(target.url, "_blank");
    }

    async publishFile(file: TFile): Promise<void> {
        await this.startPublish({ type: "document", path: file.path, title: file.basename });
    }

    async publishFolder(folder: TFolder): Promise<void> {
        await this.startPublish({ type: "folder", path: folder.path, title: folder.name });
    }

    async publishVault(): Promise<void> {
        await this.startPublish({ type: "vault", path: "", title: this.app.vault.getName() });
    }

    private async startPublish(scope: PublishScope): Promise<void> {
        if (!this.auth.isSignedIn()) return void new Notice(localize("请先登录 SmartKeep 账号，再开始发布"));
        if (this.publishWritesPaused()) return void new Notice(localize("发布服务正在初始化或退出，请稍后再试"));
        const existing = this.publishedLink(scope.path);
        if (existing) {
            new Notice(localize("此范围已有发布空间，范围内的新增、修改和删除将自动同步。"));
            window.open(existing.url, "_blank");
            return;
        }
        if (this.hasBlockingFailure()) return void new Notice(localize("发布状态尚未恢复，请先使用“重新对账”"));
        const isCurrent = this.sessionGuard();
        const previewKey = `${this.auth.activeRegion()}:${scope.type}:${normalize(scope.path)}`;
        if (this.activePreviews.has(previewKey)) return void new Notice(localize("该发布范围正在准备中，请勿重复操作"));
        this.activePreviews.add(previewKey);
        try {
            const preview = await this.preview(scope);
            if (!isCurrent()) {
                this.activePreviews.delete(previewKey);
                return;
            }
            if (!preview.files.length) {
                this.activePreviews.delete(previewKey);
                return void new Notice(localize("所选范围内没有 Markdown 文档"));
            }
            new PublishSetupModal(this.app, preview, async (options) => {
                if (!isCurrent()) throw new Error("发布登录已变化，请重新开始操作");
                await this.createSpace(preview, options);
            }, () => this.activePreviews.delete(previewKey)).open();
        } catch (error: any) {
            this.activePreviews.delete(previewKey);
            new Notice(localize("发布预览失败：{0}", localizeMessage(String(error?.message || error))), 7000);
        }
    }

    private async preview(scope: PublishScope): Promise<PublishPreview> {
        const files = this.app.vault.getMarkdownFiles().filter((file) => scopeContains(scope, file.path));
        const blocked: PublishPreview["blocked"] = [];
        for (const file of files) {
            const images = localImages(await this.app.vault.cachedRead(file));
            if (images.length) blocked.push({ path: file.path, images });
        }
        return { scope, files, blocked, conflicts: [] };
    }

    private async createSpace(
        preview: PublishPreview,
        options: { title: string; access: "public" | "password"; password: string },
    ): Promise<void> {
        if (this.publishWritesPaused()) throw new Error("发布服务正在初始化或退出，请稍后再试");
        if (this.publishedLink(preview.scope.path)) throw new Error("此范围已属于已有发布空间，请在管理中确认更新，不要重复创建。");
        if (this.hasBlockingFailure()) throw new Error("发布状态尚未恢复，请先重新对账");
        if (this.creating) throw new Error("已有发布任务正在执行，请稍后再试");
        const operationState: RunState = { confirmed: true, runAllowed: true };
        const operation = this.authority.start(operationState);
        this.creating = true;
        let remoteId = "";
        const local = (this.store as any).settings.local;
        local.pendingPublishCreates = local.pendingPublishCreates || {};
        const region = this.auth.activeRegion();
        const existingPending = Object.values(local.pendingPublishCreates).find((item: any) =>
            (item.publishRegion ?? "cn") === region
            && (!item.account || item.account === local.smartkeepAccountId)
            && item.title === options.title
            && item.access === options.access
            && item.scope?.type === preview.scope.type
            && normalize(item.scope?.path || "") === normalize(preview.scope.path),
        ) as any;
        const requestId = String(existingPending?.requestId || id("create"));
        const syncId = requestId;
        local.pendingPublishCreates[requestId] = {
            publishRegion: region,
            account: local.smartkeepAccountId,
            requestId,
            syncId,
            title: options.title,
            access: options.access,
            scope: preview.scope,
            createdAt: Date.now(),
        };
        try {
            // Keep the password only in this explicit submission. Read-only recovery
            // never needs it; a later confirmed retry supplies it again.
            await (this.store as any).save();
            let response: PublishApiResult = {};
            if (existingPending) {
                const recovered = await this.api<any>(
                    "GET",
                    `/api/v1/spaces/requests/${requestId}`,
                    undefined,
                    [404],
                );
                response = recovered.space || recovered || {};
            }
            if (!response.id) {
                this.authority.beginFinal(operation);
                response = await this.api<PublishApiResult>("POST", "/api/v1/spaces", {
                    title: options.title,
                    access: options.access,
                    password: options.password || undefined,
                    scopeType: preview.scope.type,
                    requestId,
                });
                this.authority.finishFinal(operation);
            }
            if (!this.authority.isCurrent(operation)) {
                throw new Error("创建请求已完成，当前操作已停止；下次启动将按唯一凭证找回并暂停展示");
            }
            remoteId = response.id || "";
            let resolved = response;
            if (!remoteId) {
                const listed = await this.api<any>("GET", `/api/v1/spaces/requests/${requestId}`);
                this.assertActive(operation);
                resolved = listed.space || listed || {};
                remoteId = resolved.id || "";
            }
            if (!remoteId) throw new Error("发布已被服务接受，但未能确认新空间身份；已停止后续上传");
            const space: PublishedSpace = {
                publishRegion: region,
                id: remoteId,
                code: resolved.code || "",
                title: options.title,
                url: resolved.url || "",
                access: options.access,
                scope: preview.scope,
                documents: [],
                pendingDocuments: [],
                pendingDeletes: [],
                feedbackOpen: 0,
                bookmarkStatus: options.access === "public" ? "pending" : "not_applicable",
                bookmarkError: "",
                lastPublishedAt: 0,
                cloudVersion: Number((resolved as any).version ?? (resolved as any).updatedAt ?? 0),
                reconciliationStatus: "matched",
                operationStatus: "submitting",
                statusMessage: "正在提交完整首个版本",
                scopeKnown: true,
                confirmed: true,
                runAllowed: true,
                confirmationScope: `${preview.scope.type}:${preview.scope.path}`,
                confirmationAccount: String(local.smartkeepAccountId || ""),
            } as PublishedSpace;
            const previousResult = await this.api<any>(
                "GET",
                `/api/v1/spaces/${remoteId}/operations/${requestId}`,
                undefined,
                [404],
            );
            if (previousResult?.committed) {
                this.applyRecoveredCreate(space, previousResult.result || {});
            } else {
            await this.uploadFiles(
                space,
                preview.files,
                true,
                operation,
                Number((resolved as any).version ?? (resolved as any).updatedAt ?? 0),
                syncId,
                [],
            );
            }
            this.assertActive(operation);
            space.operationStatus = "success";
            space.reconciliationStatus = "matched";
            space.statusMessage = "云端已确认完整首个版本";
            this.rememberSpace(space);
            if (this.accountSpaceAccount === String(local.smartkeepAccountId || "")) this.accountSpaceIds.add(space.id);
            delete local.pendingPublishCreates[requestId];
            await (this.store as any).save();
            new Notice(localize("发布完成：{0} 篇文档", preview.files.length));
        } catch (error) {
            const status = Number((error as any)?.status || 0);
            if (!remoteId && status >= 400 && status < 500 && status !== 408) {
                delete local.pendingPublishCreates[requestId];
                await (this.store as any).save();
            }
            throw error;
        } finally {
            this.authority.finishFinal(operation);
            this.creating = false;
        }
    }

    private async uploadFiles(
        space: PublishedSpace,
        files: TFile[],
        includeNew: boolean,
        operation?: RunTicket,
        boundVersion?: number,
        suppliedSyncId?: string,
        deletePaths: string[] = [],
    ): Promise<void> {
        const ticket = operation || this.authority.start(
            space as PublishedSpace & RunState,
            `${space.scope.type}:${space.scope.path}|${String((this.store as any).settings.local.smartkeepAccountId || "")}`,
        );
        const payload: any[] = [];
        const requirePublish = () => {
            this.assertActive(ticket);
            if (this.app.vault.getMarkdownFiles().some(file => scopeContains(space.scope, file.path) && this.isAiPending(file.path))) {
                throw new Error("文档包含未核准的 AI 改动，已停止自动公网发布");
            }
            for (const item of payload) {
                const file = this.app.vault.getAbstractFileByPath(item.localPath);
                if (!(file instanceof TFile) || !scopeContains(space.scope, file.path) || file.stat.mtime !== item.modifiedAt || file.stat.size !== item.sourceSize) {
                    throw new Error("派发前发布范围或内容已变化，已停止提交");
                }
            }
        };
        requirePublish();
        const expectedVersion = (boundVersion ?? Number(space.cloudVersion || 0)) || await this.currentSpaceVersion(space.id);
        const permitExpiresAt = Date.now() + 30 * 60 * 1000;
        const existing = new Map(space.documents.map((doc) => [normalize(doc.localPath), doc]));
        for (const file of files) {
            if (this.isAiPending(file.path)) continue;
            const readPath = file.path;
            const readMtime = file.stat.mtime;
            const readSize = file.stat.size;
            const markdown = await this.app.vault.cachedRead(file);
            if (file.path !== readPath || !scopeContains(space.scope, readPath) || file.stat.mtime !== readMtime || file.stat.size !== readSize) {
                throw new Error("读取时发布范围或内容已变化，已停止提交");
            }
            requirePublish();
            const images = localImages(markdown);
            if (images.length) throw new Error(`${file.path} 包含本地图片，已停止发布`);
            const known = existing.get(normalize(file.path));
            if (!known && !includeNew) continue;
            payload.push({
                id: known?.id || id("doc"),
                path: relativeToScope(space.scope, file),
                localPath: file.path,
                title: file.basename,
                markdown,
                modifiedAt: file.stat.mtime,
                sourceSize: file.stat.size,
            });
        }
        const deleteDocuments = deletePaths
            .map((path) => space.documents.find((item) => normalize(item.localPath) === normalize(path)))
            .filter(Boolean) as PublishedDocument[];
        if (!payload.length && !deleteDocuments.length) return;
        const batches: any[][] = [];
        let batch: any[] = [];
        let batchBytes = 0;
        for (const item of payload) {
            const itemBytes = new TextEncoder().encode(JSON.stringify(item)).length + 64;
            if (batch.length && batchBytes + itemBytes > 2_500_000) {
                batches.push(batch);
                batch = [];
                batchBytes = 0;
            }
            batch.push(item);
            batchBytes += itemBytes;
        }
        if (batch.length) batches.push(batch);
        if (!batches.length) batches.push([]);
        const syncId = suppliedSyncId || id("publish");
        let result: PublishApiResult & { bookmarkStatus?: any; bookmarkError?: string } = {};
        for (let index = 0; index < batches.length; index += 1) {
            this.assertActive(ticket);
            for (const item of batches[index]) {
                const current = this.app.vault.getAbstractFileByPath(item.localPath);
                if (
                    !(current instanceof TFile)
                    || current.stat.mtime !== item.modifiedAt
                    || current.stat.size !== item.sourceSize
                    || await this.app.vault.cachedRead(current) !== item.markdown
                ) {
                    throw new Error(`发布前内容再次变化，已停止并保留当前版本：${item.localPath}`);
                }
            }
            result = await this.api("PUT", `/api/v1/spaces/${space.id}/documents`, {
                syncId,
                reset: index === 0,
                finalize: false,
                deleteDocumentIds: deleteDocuments.map((doc) => doc.id),
                documents: batches[index].map(({ sourceSize, ...document }) => document),
            }, [], requirePublish);
            this.assertActive(ticket);
        }
        for (const item of payload) {
            const current = this.app.vault.getAbstractFileByPath(item.localPath);
            if (
                !(current instanceof TFile)
                || current.stat.mtime !== item.modifiedAt
                || current.stat.size !== item.sourceSize
                || await this.app.vault.cachedRead(current) !== item.markdown
            ) throw new Error(`正式发布前内容再次变化，暂存未提交：${item.localPath}`);
        }
        this.assertActive(ticket);
        this.authority.beginFinal(ticket);
        try {
            result = await this.api("POST", `/api/v1/spaces/${space.id}/documents/${syncId}/commit`, {
                expectedVersion,
                permitExpiresAt,
                requestId: syncId,
            }, [], requirePublish);
            const returned = result.documents || payload.map((doc) => ({
                id: doc.id,
                localPath: doc.localPath,
                publicPath: doc.path,
                modifiedAt: doc.modifiedAt,
            }));
            for (const doc of returned as any[]) {
                const localPath = doc.localPath || payload.find((item) => item.id === doc.id)?.localPath;
                const index = space.documents.findIndex((item) => item.id === doc.id || normalize(item.localPath) === normalize(localPath));
                const next = { ...doc, localPath, publicPath: doc.publicPath || doc.path } as PublishedDocument;
                if (index >= 0) space.documents[index] = next;
                else space.documents.push(next);
            }
            space.lastPublishedAt = Date.now();
            space.cloudVersion = Number((result as any).version || expectedVersion);
            space.bookmarkStatus = result.bookmarkStatus || space.bookmarkStatus || "pending";
            space.bookmarkError = result.bookmarkError || "";
            space.pendingDocuments = space.pendingDocuments.filter((path) => !payload.some((item) => item.localPath === path));
            space.pendingDeletes = space.pendingDeletes.filter((path) => !payload.some((item) => normalize(item.localPath) === normalize(path)));
            if (deleteDocuments.length) {
                const deletedIds = new Set(deleteDocuments.map((doc) => doc.id));
                space.documents = space.documents.filter((doc) => !deletedIds.has(doc.id));
                space.pendingDeletes = space.pendingDeletes.filter((path) => !deletePaths.some((item) => normalize(item) === normalize(path)));
            }
            await (this.store as any).save();
        } finally {
            this.authority.finishFinal(ticket);
        }
    }

    async updateSpace(spaceId: string, includeNew: boolean, includeDeletes = false, retry = false): Promise<void> {
        const space = this.spaces().find((item) => item.id === spaceId);
        if (!space || !this.canRun(space)) throw new Error("发布空间尚未获得运行许可或人工确认");
        const running = this.updateJobs.get(spaceId);
        if (running) return await running;
        if (this.hasBlockingFailure() && !retry) throw new Error("发布状态尚未恢复，请先重新对账");
        if (this.isBusy(space)) throw new Error("发布操作正在进行中，请等待本次结果");
        if (space.operationStatus === "failed" && !retry) throw new Error("上次操作失败，请使用唯一的重试入口");
        const job = this.performUpdateSpace(spaceId, includeNew, includeDeletes, retry);
        this.updateJobs.set(spaceId, job);
        try {
            await job;
        } finally {
            this.updateJobs.delete(spaceId);
        }
    }

    private async performUpdateSpace(spaceId: string, includeNew: boolean, includeDeletes: boolean, retry: boolean): Promise<void> {
        const space = this.spaces().find((item) => item.id === spaceId);
        if (!space) return;
        const reusable = retry && space.pendingOperation?.action === "update" ? space.pendingOperation : null;
        const credential = reusable?.requestId || id("update");
        const pending: PublishPendingOperation = reusable || {
            action: "update",
            requestId: credential,
            syncId: credential,
            expectedVersion: Number(space.cloudVersion || 0) || await this.currentSpaceVersion(space.id),
            includeNew,
            includeDeletes,
        };
        space.pendingOperation = pending;
        space.operationStatus = "preparing";
        space.statusMessage = "正在准备完整发布版本";
        await (this.store as any).save();
        try {
            const files = this.app.vault.getMarkdownFiles().filter((file) => scopeContains(space.scope, file.path) && !this.isAiPending(file.path));
            space.operationStatus = "submitting";
            space.statusMessage = "正在一次性提交正文、目录和关联关系";
            await (this.store as any).save();
            await this.uploadFiles(
                space,
                files,
                Boolean(pending.includeNew),
                undefined,
                pending.expectedVersion,
                pending.syncId,
                pending.includeDeletes ? [...space.pendingDeletes] : [],
            );
            space.pendingOperation = undefined;
            space.operationStatus = "reconciling";
            space.statusMessage = "提交完成，正在核对云端结果";
            await (this.store as any).save();
            space.operationStatus = "success";
            space.reconciliationStatus = "matched";
            space.statusMessage = "云端已切换为完整新版本";
            await (this.store as any).save();
            new Notice(localize("云端已更新为本地最新版"));
        } catch (error: any) {
            await this.handleOperationFailure(space, error);
            throw error;
        }
    }

    async approvePending(spaceId: string): Promise<void> {
        await this.updateSpace(spaceId, true, true);
    }

    async confirmPendingDeletes(spaceId: string): Promise<void> {
        const space = this.spaces().find((item) => item.id === spaceId);
        if (!space || !this.canRun(space) || !space.pendingDeletes.length) return;
        await this.updateSpace(spaceId, false, true);
        new Notice(localize("已从云端移除确认删除的文档，本地回收规则不受影响"));
    }

    async toggleAccess(spaceId: string, retry = false): Promise<void> {
        const space = this.spaces().find((item) => item.id === spaceId);
        if (!space || !this.canRun(space)) return;
        if (this.hasBlockingFailure() && !retry) throw new Error("发布状态尚未恢复，请先重新对账");
        if (this.isBusy(space)) throw new Error("发布操作正在进行中，请等待本次结果");
        if (space.operationStatus === "failed" && !retry) throw new Error("上次操作失败，请使用唯一的重试入口");
        const isCurrent = this.sessionGuard();
        const reusable = retry && space.pendingOperation?.action === "access" ? space.pendingOperation : null;
        const access = reusable?.desiredAccess || (space.access === "public" ? "password" : "public");
        if (!reusable) {
            space.operationStatus = "preparing";
            space.statusMessage = "正在等待用户确认访问方式";
            space.pendingOperation = {
                action: "access",
                requestId: id("access"),
                expectedVersion: Number(space.cloudVersion || 0),
                desiredAccess: access,
            };
            await (this.store as any).save();
        }
        let password = "";
        if (access === "password") {
            password = window.prompt(localize("请输入至少 6 个字符的访问密码")) || "";
            if (password.length < 6) {
                if (isCurrent()) await this.cancelReservedConfirmation(spaceId);
                return void new Notice(localize("访问密码至少需要 6 个字符"));
            }
        } else {
            const proceed = await confirm(this.app, localize("改为公开访问"), localize("公开后页面可被搜索引擎收录，发布链接也会自动保存到 SmartKeep。"), localize("确认公开"));
            if (!isCurrent()) return;
            if (!proceed) return await this.cancelReservedConfirmation(spaceId);
        }
        if (!isCurrent() || !this.canRun(space)) return;
        const ticket = this.authority.start(space as PublishedSpace & RunState);
        const pending: PublishPendingOperation = space.pendingOperation!;
        if (!pending.expectedVersion) pending.expectedVersion = await this.currentSpaceVersion(spaceId);
        space.pendingOperation = pending;
        space.operationStatus = "submitting";
        space.statusMessage = "正在提交新的访问方式";
        await (this.store as any).save();
        this.assertActive(ticket);
        this.authority.beginFinal(ticket);
        try {
            const result = await this.api<any>("PATCH", `/api/v1/spaces/${spaceId}`, {
                access,
                password,
                expectedVersion: pending.expectedVersion,
                requestId: pending.requestId,
                permitExpiresAt: Date.now() + 30 * 60 * 1000,
            });
            space.access = access;
            space.cloudVersion = Number(result.version || pending.expectedVersion);
            space.bookmarkStatus = result.bookmarkStatus || (access === "public" ? "pending" : "not_applicable");
            space.bookmarkError = result.bookmarkError || "";
            space.pendingOperation = undefined;
            space.operationStatus = "success";
            space.reconciliationStatus = "matched";
            space.statusMessage = "访问方式已生效，旧访问资格已失效";
            await (this.store as any).save();
        } catch (error: any) {
            await this.handleOperationFailure(space, error);
            throw error;
        } finally {
            this.authority.finishFinal(ticket);
        }
    }

    async retryBookmark(spaceId: string): Promise<void> {
        const space = this.spaces().find((item) => item.id === spaceId);
        if (!space || !this.canRun(space) || space.access !== "public") return;
        const force = ["detached", "user_removed"].includes(space.bookmarkStatus);
        const result = await this.api<any>("POST", `/api/v1/spaces/${spaceId}/bookmark/retry`, { force });
        space.bookmarkStatus = result.bookmarkStatus || "pending";
        space.bookmarkError = result.bookmarkError || "";
        await (this.store as any).save();
    }

    /** Coalesce vault events; retain changes arriving during an in-flight commit. */
    private scheduleContinuousSync(space: PublishedSpace): void {
        window.clearTimeout(this.updateTimers.get(space.id));
        if (this.resuming && !this.disposed) {
            this.deferredUpdates.add(space);
            return;
        }
        if (!this.canRun(space) || space.operationStatus === "failed") return;
        this.updateTimers.set(space.id, window.setTimeout(() => {
            this.updateTimers.delete(space.id);
            if (!this.canRun(space) || space.operationStatus === "failed") return;
            if (this.isBusy(space) || this.updateJobs.has(space.id)) {
                this.scheduleContinuousSync(space);
                return;
            }
            if (this.app.vault.getMarkdownFiles().some(file => scopeContains(space.scope, file.path) && this.isAiPending(file.path))) {
                space.operationStatus = "pending_approval";
                space.statusMessage = "文档包含未核准的 AI 改动，自动公网发布已暂停。";
                void (this.store as any).save();
                return;
            }
            // Rebuild deletion intent from the current inventory, including changes while offline.
            space.pendingDeletes = space.documents.filter(doc => {
                const file = this.app.vault.getAbstractFileByPath(doc.localPath);
                return !(file instanceof TFile) || file.extension.toLowerCase() !== "md" || !scopeContains(space.scope, file.path);
            }).map(doc => doc.localPath);
            void this.updateSpace(space.id, true, true).catch(() => {
                // performUpdateSpace retains the failure and original operation for explicit recovery.
            });
        }, 1500));
    }

    /** Called after layout is ready; do not couple writes to the read-only refresh action. */
    checkContinuousSync(): void {
        if (this.publishWritesPaused()) return;
        for (const space of this.spaces()) {
            if (!this.canRun(space) || space.operationStatus === "failed" || space.pendingOperation || this.isBusy(space)) continue;
            // Missing publication root on startup may be an unavailable mount, not an intentional delete.
            if (space.scope.type !== "vault" && !this.app.vault.getAbstractFileByPath(space.scope.path)) continue;
            const files = this.app.vault.getMarkdownFiles().filter(file => scopeContains(space.scope, file.path));
            const known = new Map(space.documents.map(doc => [normalize(doc.localPath), doc]));
            if (space.pendingDocuments.length || space.pendingDeletes.length || files.length !== space.documents.length
                || files.some(file => !known.has(normalize(file.path)) || Number(known.get(normalize(file.path))!.modifiedAt) !== file.stat.mtime)) {
                this.scheduleContinuousSync(space);
            }
        }
    }

    onVaultChange(file: TAbstractFile): void {
        if (!(file instanceof TFile) || file.extension.toLowerCase() !== "md") return;
        for (const space of this.spaces()) {
            if (!this.hasRunPermission(space) || !scopeContains(space.scope, file.path)) continue;
            if (!space.documents.some(doc => normalize(doc.localPath) === normalize(file.path)) && !space.pendingDocuments.includes(file.path)) {
                space.pendingDocuments.push(file.path);
            }
            space.pendingDeletes = space.pendingDeletes.filter(path => normalize(path) !== normalize(file.path));
            void (this.store as any).save();
            this.scheduleContinuousSync(space);
        }
    }

    onVaultRename(file: TAbstractFile, oldPath: string): void {
        const oldRoot = normalize(oldPath);
        const newRoot = normalize(file.path);
        let aiTaintChanged = false;
        for (const pending of this.getAiPendingPaths()) {
            if (pending === oldRoot) {
                this.pendingAiApprovalPaths.delete(pending);
                this.pendingAiApprovalPaths.set(newRoot, null);
                aiTaintChanged = true;
            } else if (pending.startsWith(`${oldRoot}/`)) {
                this.pendingAiApprovalPaths.delete(pending);
                this.pendingAiApprovalPaths.set(`${newRoot}${pending.slice(oldRoot.length)}`, null);
                aiTaintChanged = true;
            }
        }
        if (aiTaintChanged) {
            void this.persistAiPending().catch(() => new Notice(localize("公开审批状态尚未保存，请检查插件数据存储")));
        }
        this.applyVaultRename(file, oldRoot, newRoot);
    }

    private applyVaultRename(file: TAbstractFile, oldRoot: string, newRoot: string): void {
        if (this.updateJobs.size) {
            void Promise.allSettled([...this.updateJobs.values()]).then(() => this.applyVaultRename(file, oldRoot, newRoot));
            return;
        }
        for (const space of this.spaces()) {
            if (!this.hasRunPermission(space)) continue;
            if (space.operationStatus === "failed") continue;
            let changed = false;
            // Follow a rename of the publication root itself, retaining its trusted binding.
            if (normalize(space.scope.path) === oldRoot || normalize(space.scope.path).startsWith(`${oldRoot}/`)) {
                space.scope.path = `${newRoot}${normalize(space.scope.path).slice(oldRoot.length)}`;
                (space as any).confirmationScope = `${space.scope.type}:${space.scope.path}`;
                changed = true;
            }
            for (const doc of space.documents) {
                const current = normalize(doc.localPath);
                if (current !== oldRoot && !current.startsWith(`${oldRoot}/`)) continue;
                const next = `${newRoot}${current.slice(oldRoot.length)}`;
                if (scopeContains(space.scope, next)) {
                    doc.localPath = next;
                    space.pendingDeletes = space.pendingDeletes.filter((path) => normalize(path) !== current);
                } else if (!space.pendingDeletes.some((path) => normalize(path) === current)) {
                    space.pendingDeletes.push(doc.localPath);
                }
                changed = true;
            }
            const renamedPending = space.pendingDocuments.filter((path) => {
                const current = normalize(path);
                return current === oldRoot || current.startsWith(`${oldRoot}/`);
            });
            if (renamedPending.length) {
                space.pendingDocuments = space.pendingDocuments.filter((path) => !renamedPending.includes(path));
                for (const pending of renamedPending) {
                    const current = normalize(pending);
                    const next = `${newRoot}${current.slice(oldRoot.length)}`;
                    if (scopeContains(space.scope, next) && !space.pendingDocuments.includes(next)) space.pendingDocuments.push(next);
                }
                changed = true;
            }
            if (!changed && file instanceof TFile && file.extension.toLowerCase() === "md" && scopeContains(space.scope, newRoot)) {
                if (!space.pendingDocuments.includes(newRoot)) space.pendingDocuments.push(newRoot);
                changed = true;
            }
            if (!changed && this.app.vault.getMarkdownFiles().some(item => scopeContains(space.scope, item.path) && (normalize(item.path) === newRoot || normalize(item.path).startsWith(`${newRoot}/`)))) changed = true;
            if (!changed) continue;
            void (this.store as any).save();
            this.scheduleContinuousSync(space);
        }
        this.feedbackChanged?.();
    }

    onVaultDelete(file: TAbstractFile): void {
        const deletedRoot = normalize(file.path);
        for (const space of this.spaces()) {
            if (!this.hasRunPermission(space)) continue;
            const previouslyPending = space.pendingDocuments.length;
            space.pendingDocuments = space.pendingDocuments.filter((path) => {
                const current = normalize(path);
                return current !== deletedRoot && !current.startsWith(`${deletedRoot}/`);
            });
            const deletedDocuments = space.documents.filter((doc) => {
                const current = normalize(doc.localPath);
                return current === deletedRoot || current.startsWith(`${deletedRoot}/`);
            });
            for (const doc of deletedDocuments) {
                if (!space.pendingDeletes.some((path) => normalize(path) === normalize(doc.localPath))) space.pendingDeletes.push(doc.localPath);
            }
            if (!deletedDocuments.length && previouslyPending === space.pendingDocuments.length) continue;
            void (this.store as any).save();
            this.scheduleContinuousSync(space);
        }
    }

    async openFeedback(spaceId: string, documentId?: string): Promise<void> {
        const space = this.spaces().find((item) => item.id === spaceId);
        if (!space) return;
        const account = String((this.store as any).settings.local.smartkeepAccountId || "");
        const epoch = this.feedbackEpoch;
        const result = await this.api<any>("GET", `/api/v1/spaces/${spaceId}/feedback${documentId ? `?documentId=${encodeURIComponent(documentId)}` : ""}`);
        if (epoch !== this.feedbackEpoch || account !== String((this.store as any).settings.local.smartkeepAccountId || "") || !this.auth.isSignedIn()) return;
        const feedback = (result.feedback || []).filter((item: any) => !documentId || String(item.document_id || item.documentId) === documentId).map((item: any) => ({
            ...item,
            authorName: item.authorName || item.author_name || "匿名访问者",
            documentPath: item.documentPath || item.document_path || "",
            resolution: item.resolution || "",
        }));
        const title = documentId ? space.documents.find(doc => doc.id === documentId)?.localPath.split("/").pop() : undefined;
        new FeedbackModal(this.app, this, space, feedback, this.sessionGuard(), title).open();
    }

    private async mutateFeedback(spaceId: string, action: (beforeSend: () => void) => Promise<any>): Promise<any> {
        const space = this.spaces().find(item => item.id === spaceId);
        const account = String((this.store as any).settings.local.smartkeepAccountId || "");
        if (!space || !account || !this.canWithdraw(space)) throw new Error("请先确认当前账号拥有此发布空间");
        const ticket = this.authority.start({ confirmed: true, runAllowed: true }, account);
        const previous = this.feedbackWrites;
        let release!: () => void;
        this.feedbackWrites = new Promise<void>(resolve => { release = resolve; });
        const requireOwner = () => {
            this.authority.require(ticket);
            if (!this.auth.isSignedIn() || account !== String((this.store as any).settings.local.smartkeepAccountId || "") || !this.canWithdraw(space)) {
                throw new Error("发布账号已变化，处理意见已停止");
            }
        };
        try {
            await previous;
            requireOwner();
            this.feedbackEpoch++;
            this.authority.beginFinal(ticket);
            const result = await action(requireOwner);
            requireOwner();
            if (Array.isArray(result.documentFeedback)) this.applyFeedbackSummary(account, result.documentFeedback, spaceId);
            if (typeof result.feedbackOpen === "number") space.feedbackOpen = result.feedbackOpen;
            await (this.store as any).save();
            return result;
        } finally {
            this.authority.finishFinal(ticket);
            release();
        }
    }

    async deleteFeedback(spaceId: string, feedbackId: string): Promise<void> {
        await this.mutateFeedback(spaceId, async beforeSend => {
            const result = await this.api<any>("DELETE", `/api/v1/spaces/${spaceId}/feedback/${encodeURIComponent(feedbackId)}`, undefined, [], beforeSend);
            if (result.deleted !== feedbackId) throw new Error("删除结果未确认");
            return result;
        });
    }

    async setFeedbackStatus(spaceId: string, feedbackId: string, status: "resolved" | "rejected", resolution: string): Promise<void> {
        await this.mutateFeedback(spaceId, async beforeSend => {
            const result = await this.api<any>("PATCH", `/api/v1/spaces/${spaceId}/feedback/${encodeURIComponent(feedbackId)}`, { status, resolution }, [], beforeSend);
            if (result.status !== status) throw new Error("处理结果未确认");
            return result;
        });
    }

    async withdrawSpace(spaceId: string, retry = false, reserved = false): Promise<void> {
        const existing = this.spaces().find((item) => item.id === spaceId);
        if (!existing) return;
        if (!this.canWithdraw(existing)) throw new Error("请先查询并确认当前账号拥有此云端空间");
        if (existing.pendingOperation && existing.pendingOperation.action !== "withdraw") throw new Error("此空间仍有结果未确认的操作，请先重新对账");
        if (this.isBusy(existing) && !reserved) throw new Error("发布操作正在进行中，请等待本次结果");
        if (existing.operationStatus === "failed" && !retry) throw new Error("上次操作失败，请使用唯一的重试入口");
        const reusable = (retry || reserved) && existing.pendingOperation?.action === "withdraw" ? existing.pendingOperation : null;
        const account = String((this.store as any).settings.local.smartkeepAccountId || "");
        // The cloud-only ticket is fixed before awaits and does not authorize local publishing.
        const ticket = this.authority.start({ confirmed: true, runAllowed: true }, account);
        const requireWithdrawal = () => {
            this.assertActive(ticket);
            if (account !== String((this.store as any).settings.local.smartkeepAccountId || "") || !this.canWithdraw(existing)) {
                throw new Error("发布账号已变化，撤回已停止，请重新确认");
            }
        };
        try {
            const pending: PublishPendingOperation = reusable || {
                action: "withdraw",
                requestId: id("withdraw"),
                expectedVersion: Number(existing.cloudVersion || 0) || await this.currentSpaceVersion(spaceId),
            };
            existing.pendingOperation = pending;
            existing.operationStatus = "submitting";
            existing.statusMessage = "正在请求云端撤回";
            await (this.store as any).save();
            requireWithdrawal();
            this.authority.beginFinal(ticket);
            await this.api("DELETE", `/api/v1/spaces/${spaceId}`, {
                expectedVersion: pending.expectedVersion,
                requestId: pending.requestId,
                permitExpiresAt: Date.now() + 30 * 60 * 1000,
            }, [], requireWithdrawal);
            this.markWithdrawn(existing);
            this.accountSpaceIds.delete(spaceId);
            this.forgetSpace(existing);
            this.feedbackChanged?.();
            await (this.store as any).save();
        } catch (error: any) {
            await this.handleOperationFailure(existing, error);
            throw error;
        } finally {
            this.authority.finishFinal(ticket);
        }
        new Notice(localize("已停止发布，本地内容未改变"));
    }

    private async currentSpaceVersion(spaceId: string): Promise<number> {
        const result = await this.api<any>("GET", `/api/v1/spaces/${spaceId}/version`);
        const version = Number(result.version ?? result.updatedAt);
        if (!Number.isFinite(version) || version < 0) {
            throw new Error("发布空间当前版本无法确认，已停止正式提交");
        }
        return version;
    }

    private async api<T = any>(method: string, path: string, body?: unknown, acceptedStatuses: number[] = [], beforeSend?: () => void, quiet = false): Promise<T> {
        const epoch = this.reconciliationEpoch;
        const { base, token } = await this.auth.getPublishSession();
        try {
            if (epoch !== this.reconciliationEpoch || this.disposed || this.signingOut) {
                throw new Error("发布登录已变化，请重新开始操作");
            }
            beforeSend?.();
            const response = await requestUrl({
                url: `${base}${path}`,
                method,
                headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
                body: body === undefined ? undefined : JSON.stringify(body),
                throw: false,
            });
            let json: any = null;
            try {
                json = response.json;
            } catch {
                try {
                    json = JSON.parse(response.text || "");
                } catch {
                    json = null;
                }
            }
            if ((response.status < 200 || response.status >= 300) && !acceptedStatuses.includes(response.status)) {
                const failure: any = new Error(
                    response.status === 401
                        ? "SmartKeep 登录已失效，请重新登录"
                        : json?.message || (response.text && !response.text.startsWith("<") ? response.text : "") || `请求失败 (${response.status})`,
                );
                failure.status = response.status;
                throw failure;
            }
            if (response.status === 204) return {} as T;
            if (json && typeof json === "object") return json as T;
            return {} as T;
        } catch (error: any) {
            if (!quiet && !this.disposed && !this.signingOut && epoch === this.reconciliationEpoch) {
                new Notice(localize("发布服务不可用：{0}", localizeMessage(String(error.message || error))), 7000);
            }
            throw error;
        }
    }

    private assertActive(operation: RunTicket): void {
        this.authority.require(operation);
        if ("scope" in operation.state && !this.canRun(operation.state as PublishedSpace)) {
            throw new Error("发布身份或范围已变化，正式落地已停止");
        }
    }

    private async reconcilePendingCreates(isCurrent: () => boolean): Promise<void> {
        const local = (this.store as any).settings.local;
        const pending = local.pendingPublishCreates || {};
        const region = this.auth.activeRegion();
        let changed = false;
        for (const requestId of Object.keys(pending)) {
            if (!isCurrent()) return;
            const item = pending[requestId];
            if ((item.publishRegion ?? "cn") !== region || (item.account && item.account !== local.smartkeepAccountId)) continue;
            let result: any;
            try {
                result = await this.api<any>("GET", `/api/v1/spaces/requests/${requestId}`, undefined, [404]);
            } catch {
                continue;
            }
            if (!isCurrent()) return;
            let remote = result.space || result;
            if (!remote?.id) {
                item.status = "awaiting_confirmation";
                changed = true;
                continue;
            }
            const operation = await this.api<any>(
                "GET", `/api/v1/spaces/${remote.id}/operations/${requestId}`, undefined, [404],
            );
            if (!isCurrent()) return;
            if (!operation?.committed) {
                item.status = "awaiting_confirmation";
                changed = true;
                continue;
            }
            let space = this.spaces().find((entry) => entry.id === remote.id);
            if (!space) space = {
                publishRegion: region,
                id: remote.id,
                code: remote.code || "",
                title: remote.title || item.title,
                url: remote.url || "",
                access: remote.access || item.access,
                scope: item.scope,
                documents: [],
                pendingDocuments: [],
                pendingDeletes: [],
                feedbackOpen: 0,
                bookmarkStatus: "pending",
                bookmarkError: "",
                lastPublishedAt: 0,
                cloudVersion: Number(remote.version ?? remote.updatedAt ?? 0),
                reconciliationStatus: "needs_restore",
                operationStatus: "preparing",
                statusMessage: "正在查询首次发布的原业务凭证",
                scopeKnown: true,
                confirmed: true,
                runAllowed: true,
                confirmationScope: `${item.scope.type}:${item.scope.path}`,
                confirmationAccount: String(local.smartkeepAccountId || ""),
            } as PublishedSpace;
            if (!this.spaces().includes(space)) this.rememberSpace(space);
            try {
                this.applyRecoveredCreate(space, operation.result || {});
                space.operationStatus = "success";
                space.reconciliationStatus = "matched";
                space.statusMessage = "已通过原业务凭证找回首次发布结果";
                delete pending[requestId];
                changed = true;
            } catch (error: any) {
                space.operationStatus = "failed";
                space.statusMessage = `首次发布尚未确认，保留原业务凭证等待重新对账：${String(error?.message || error)}`;
                changed = true;
            }
        }
        if (changed) await (this.store as any).save();
    }

    async retryOperation(spaceId: string, credentialChecked = false): Promise<void> {
        if (this.publishWritesPaused()) return;
        const space = this.spaces().find((item) => item.id === spaceId);
        const pending = space?.pendingOperation;
        if (!space || !pending || space.operationStatus !== "failed") return;
        if (!credentialChecked) {
            const epoch = this.reconciliationEpoch;
            const isCurrent = () => !this.publishWritesPaused() && this.auth.isSignedIn() && epoch === this.reconciliationEpoch;
            const state = await this.resolvePendingOperation(space, isCurrent);
            if (!isCurrent() || state === "committed") return;
        }
        if (pending.action === "update") {
            await this.updateSpace(spaceId, Boolean(pending.includeNew), Boolean(pending.includeDeletes), true);
        } else if (pending.action === "withdraw") {
            await this.withdrawSpace(spaceId, true);
        } else {
            await this.toggleAccess(spaceId, true);
        }
    }

    private async handleOperationFailure(space: PublishedSpace, error: any): Promise<void> {
        if (Number(error?.status) === 409) {
            space.pendingOperation = undefined;
            space.operationStatus = "needs_confirmation";
            space.reconciliationStatus = "pending_confirmation";
            space.statusMessage = "云端版本已变化，旧设备操作被拒绝，正在刷新真实状态";
            await (this.store as any).save();
            await this.refreshSpaces();
            return;
        }
        space.operationStatus = "failed";
        space.statusMessage = `本次操作失败，上一正常版本保持不变：${String(error?.message || error)}`;
        await (this.store as any).save();
    }

    private async resolvePendingOperation(space: PublishedSpace, isCurrent: () => boolean): Promise<"none" | "committed" | "not_committed"> {
        const pending = space.pendingOperation;
        if (!pending) return "none";
        const result = await this.api<any>(
            "GET",
            `/api/v1/spaces/${space.id}/operations/${encodeURIComponent(pending.requestId)}`,
            undefined,
            [404],
        );
        if (!isCurrent()) return "none";
        if (!result?.committed) return "not_committed";
        if (pending.action === "withdraw") {
            this.accountSpaceIds.delete(space.id);
            this.markWithdrawn(space);
            this.forgetSpace(space);
            return "committed";
        }
        space.cloudVersion = Number(result.result?.version || space.cloudVersion || 0);
        if (pending.action === "access" && result.result?.access) space.access = result.result.access;
        space.pendingOperation = undefined;
        space.operationStatus = "success";
        space.statusMessage = "已找回同一业务凭证的云端结果";
        return "committed";
    }

    private applyRecoveredCreate(space: PublishedSpace, result: any): void {
        space.documents = Array.isArray(result.documents)
            ? result.documents.map((doc: any) => ({
                id: String(doc.id),
                localPath: String(doc.localPath || doc.publicPath || ""),
                publicPath: String(doc.publicPath || doc.path || ""),
                modifiedAt: Number(doc.modifiedAt || 0),
                url: doc.url,
            }))
            : space.documents;
        space.cloudVersion = Number(result.version || space.cloudVersion || 0);
        space.lastPublishedAt = Number(result.version || Date.now());
        space.bookmarkStatus = result.bookmarkStatus || space.bookmarkStatus;
        space.bookmarkError = result.bookmarkError || "";
    }

    private localPathForRemote(space: PublishedSpace, publicPath: string): string {
        const path = normalize(publicPath);
        if (space.scope.type === "vault") return path;
        if (space.scope.type === "document") return space.scope.path || path;
        const root = normalize(space.scope.path);
        return root ? `${root}/${path}` : path;
    }
}
