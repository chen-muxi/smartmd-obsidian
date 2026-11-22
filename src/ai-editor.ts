import {
    ItemView,
    MarkdownView,
    Modal,
    Notice,
    TFile,
    WorkspaceLeaf,
} from "obsidian";
import { sha256Text, SmartMDStore } from "./core";
import type { ParagraphRange } from "./ai-paragraph";
import { AIChatPreview } from "./ai-chat-preview";
import { paragraphSelection, trackedOldIn, type TextRange } from "./ai-chat-tracking";
import { modelSettings, requestConversation, webSources, type ModelMessage, type WebSource } from "./ai-model";
import { applyReviewChanges, documentReviewChanges, reverseReviewChanges, type ReviewChange } from "./ai-review-blocks";

export const AI_VIEW_TYPE = "obsidian-suite-ai-editor";

type ChatMessage = { role: "user" | "assistant"; content: string; quote?: { filePath: string }; sources?: WebSource[] };
type SavedDraft = {
    filePath: string;
    base: string;
    from: number;
    to: number;
    kind: "document" | "selection" | "cleanup";
    replacement: string;
    history: ModelMessage[];
};
type ChatSession = {
    id: string;
    messages: ChatMessage[];
    context?: { filePath: string | null; messages: ModelMessage[] };
    draft?: SavedDraft;
};
const isChatMessage = (message: ChatMessage) => (message?.role === "user" || message?.role === "assistant") && typeof message.content === "string";
const sessionTitle = (session: ChatSession) =>
    session.messages.find(message => message.role === "user")?.content.trim().replace(/\s+/g, " ").slice(0, 34) || "新会话";
const atTranscriptEnd = (transcript: HTMLElement) =>
    transcript.scrollHeight - transcript.clientHeight - transcript.scrollTop <= 2;

function readSavedDraft(draft: SavedDraft | undefined): SavedDraft | undefined {
    if (!draft || typeof draft.filePath !== "string" || typeof draft.base !== "string" || !/^[a-f\d]{64}$/.test(draft.base)
        || !Number.isSafeInteger(draft.from) || !Number.isSafeInteger(draft.to) || draft.from < 0 || draft.to <= draft.from
        || !["document", "selection", "cleanup"].includes(draft.kind) || typeof draft.replacement !== "string") return;
    return { ...draft, history: Array.isArray(draft.history) ? draft.history.filter(isChatMessage).slice(-8) : [] };
}
type DocumentTarget = { view: MarkdownView; file: TFile; filePath: string; note: string };
export type ChatQuote = DocumentTarget & TextRange & ({ scope: "selection"; paragraph: ParagraphRange } | { scope: "document" });
type ConversationChange = DocumentTarget & TextRange & {
    kind: "edit" | "cleanup";
    replacement: string;
    scope?: ChatQuote["scope"];
    changes: ReviewChange[];
    diskBase: string;
};
type UndoRecord = {
    sessionId: string;
    file: TFile;
    filePath: string;
    before: string;
    after: string;
    pendingBefore: boolean;
    publishRevision?: symbol;
    reverse: ReviewChange[];
};

export class AIEditorView extends ItemView {
    private suggestion: ConversationChange | null = null;
    private shownReview: ConversationChange | null = null;
    private undoRecord: UndoRecord | null = null;
    private busy = false;
    private searchNext = false;
    private committing = false;
    private restoring = false;
    private closed = false;
    private sessionGeneration = 0;
    private activeRequestContext: { file?: TFile; note?: string; matches: () => boolean; invalidated: boolean } | null = null;
    private diskObservation: { file: TFile; done: Promise<void> } | null = null;
    private messages: ChatMessage[] = [];
    private sessions: ChatSession[] = [];
    private sessionId = "";
    private historyKey = "";
    private historyOpen = false;
    private displayedTranscript?: { sessionId: string; messageCount: number; scrollTop: number; followLatest: boolean };
    private transcriptResize?: ResizeObserver;
    private sessionDrafts = new Map<string, { quote: ChatQuote | null;
        thread: AIEditorView["modelThread"]; input: string }>();
    private modelThread: { quote: ChatQuote | null; file: TFile | null; messages: ModelMessage[] } | null = null;
    private chatInput = "";
    private quote: ChatQuote | null = null;

    constructor(leaf: WorkspaceLeaf, private store: SmartMDStore, private runtime?: any, private preview?: AIChatPreview) {
        super(leaf);
    }

    getViewType(): string {
        return AI_VIEW_TYPE;
    }

    getDisplayText(): string {
        return "AI 编辑";
    }

    getIcon(): string {
        return "wand-sparkles";
    }

    async onOpen(): Promise<void> {
        this.closed = false;
        const adapter = this.app.vault.adapter as typeof this.app.vault.adapter & { getBasePath?: () => string };
        const location = adapter.getBasePath?.() || this.app.vault.getName();
        this.historyKey = "smartmd-ai-history:" + location;
        try {
            const saved = JSON.parse(window.localStorage.getItem(this.historyKey) || "null");
            if (Array.isArray(saved?.sessions)) {
                this.sessions = saved.sessions.filter((entry: ChatSession) => typeof entry?.id === "string" && Array.isArray(entry.messages))
                    .map((entry: ChatSession) => ({
                        id: entry.id,
                        messages: entry.messages.filter(isChatMessage),
                        draft: readSavedDraft(entry.draft),
                        context: entry.context && Array.isArray(entry.context.messages)
                            ? { filePath: entry.context.filePath, messages: entry.context.messages.filter(isChatMessage) }
                            : undefined,
                    }));
                const selected = this.sessions.find(entry => entry.id === saved.active) || this.sessions[0];
                if (selected) {
                    this.sessionId = selected.id;
                    this.messages = selected.messages;
                }
            }
        } catch {}
        if (!this.sessionId) this.createSession();
        const workspace = this.app.workspace as any;
        const editorChange = workspace.on?.("editor-change", () => this.observeContextChange());
        const activeLeafChange = workspace.on?.("active-leaf-change", () => this.observeContextChange());
        const fileOpen = workspace.on?.("file-open", () => this.observeContextChange());
        if (editorChange) this.registerEvent(editorChange);
        if (activeLeafChange) this.registerEvent(activeLeafChange);
        if (fileOpen) this.registerEvent(fileOpen);
        this.registerEvent(this.app.vault.on("modify", file => this.observeVaultModification(file)));
        this.registerEvent(this.app.vault.on("rename", file => this.invalidateFile(file)));
        this.registerEvent(this.app.vault.on("delete", file => this.invalidateFile(file)));
        if (typeof document !== "undefined") {
            this.registerDomEvent(document, "selectionchange", () => this.observeContextChange());
        }
        this.render();
        this.app.workspace.onLayoutReady(() => {
            if (!this.closed) void this.restoreDraft(false);
        });
    }

    async onClose(): Promise<void> {
        this.closed = true;
        this.transcriptResize?.disconnect();
        this.transcriptResize = undefined;
        this.displayedTranscript = undefined;
        this.sessionGeneration++;
        this.suggestion = null;
        this.quote = null;
        this.modelThread = null;
        this.preview?.clear();
        this.busy = false;
        this.sessionDrafts.clear();
    }

    private saveHistory(sessions = this.sessions, active = this.sessionId): boolean {
        try {
            window.localStorage.setItem(this.historyKey, JSON.stringify({
                active,
                sessions: sessions.map(session => ({
                    id: session.id,
                    context: session.context,
                    draft: session.draft,
                    messages: session.messages.map(({ role, content, quote, sources }) => ({
                        role, content, quote: quote && { filePath: quote.filePath }, sources,
                    })),
                })),
            }));
            return true;
        } catch {
            return false;
        }
    }

    private currentSession(): ChatSession | undefined {
        return this.sessions.find(session => session.id === this.sessionId);
    }

    private forgetDraft(): boolean {
        const session = this.currentSession();
        if (!session?.draft) return true;
        const draft = session.draft;
        delete session.draft;
        if (this.saveHistory()) return true;
        session.draft = draft;
        new Notice("本机存储写入失败，草稿仍然保留，请重试。", 7000);
        return false;
    }

    private addMessage(message: ChatMessage): void {
        this.messages.push(message);
        this.saveHistory();
    }

    private createSession(): void {
        const session = { id: crypto.randomUUID(), messages: [] };
        this.sessions.unshift(session);
        this.selectSession(session);
    }

    private selectSession(session: ChatSession, persistSelection = true): void {
        if (this.committing || this.busy || this.restoring) return;
        if (session.id === this.sessionId) {
            this.render(true);
            void this.restoreDraft();
            return;
        }
        if (this.currentSession()) {
            this.sessionDrafts.set(this.sessionId, { quote: this.quote, thread: this.modelThread, input: this.chatInput });
        }
        this.sessionGeneration++;
        this.preview?.clear();
        this.sessionId = session.id;
        this.messages = session.messages;
        const draft = this.sessionDrafts.get(session.id);
        this.sessionDrafts.delete(session.id);
        this.quote = !session.draft && draft?.quote && this.documentMatches(draft.quote) ? draft.quote : null;
        this.suggestion = null;
        this.shownReview = null;
        this.searchNext = false;
        this.modelThread = draft?.thread ?? null;
        this.chatInput = draft?.input ?? "";
        if (persistSelection) this.saveHistory();
        this.render();
        if (session.draft) void this.restoreDraft();
    }

    private confirmDeleteSession(session: ChatSession): void {
        const modal = new Modal(this.app);
        modal.titleEl.setText("删除会话");
        modal.contentEl.createEl("p", { text: sessionTitle(session) });
        modal.contentEl.createEl("p", { text: "删除后，此会话的聊天记录和未应用草稿将从本机移除。" });
        const actions = modal.contentEl.createDiv({ cls: "modal-button-container" });
        const cancel = actions.createEl("button", { text: "取消" });
        cancel.onclick = () => modal.close();
        const confirm = actions.createEl("button", { text: "确认删除", cls: "mod-warning" });
        confirm.onclick = () => {
            const index = this.sessions.indexOf(session);
            if (index < 0 || this.closed || this.busy || this.committing || this.restoring) {
                modal.close();
                return;
            }
            const remaining = this.sessions.filter(entry => entry !== session);
            if (!remaining.length) remaining.push({ id: crypto.randomUUID(), messages: [] });
            const deletingCurrent = session.id === this.sessionId;
            const next = deletingCurrent ? remaining[Math.min(index, remaining.length - 1)] : null;
            if (!this.saveHistory(remaining, next?.id ?? this.sessionId)) {
                new Notice("未能删除会话，请稍后重试。");
                return;
            }
            this.sessions = remaining;
            this.sessionDrafts.delete(session.id);
            if (this.undoRecord?.sessionId === session.id) this.undoRecord = null;
            modal.close();
            if (next) {
                this.selectSession(next, false);
            } else {
                const rows = this.containerEl.querySelectorAll<HTMLElement>(".smartmd-ai-history-row");
                Array.from(rows).find(row => row.dataset.sessionId === session.id)?.remove();
            }
            this.containerEl.querySelector<HTMLElement>(".smartmd-ai-history > summary")?.focus({ preventScroll: true });
        };
        modal.open();
    }

    private async restoreDraft(openNote = true): Promise<void> {
        const saved = this.currentSession()?.draft;
        if (!saved || this.closed || this.busy || this.committing || this.restoring || this.suggestion) return;
        const file = this.app.vault.getAbstractFileByPath(saved.filePath);
        const initialView = this.activeEditor();
        if (!openNote && (initialView?.file !== file || initialView?.getMode() !== "source")) return;
        if (!(file instanceof TFile)) {
            new Notice("草稿对应的原笔记已找不到。");
            return;
        }
        this.restoring = true;
        this.render();
        const generation = this.sessionGeneration;
        const stillCurrent = () => !this.closed && generation === this.sessionGeneration && this.currentSession()?.draft === saved
            && this.app.vault.getAbstractFileByPath(saved.filePath) === file;
        try {
            const disk = await this.app.vault.read(file);
            const diskBase = await sha256Text(disk.replace(/\r\n?/g, "\n"));
            if (!stillCurrent()) return;
            if (diskBase !== saved.base) throw new Error("原笔记已变化，暂时无法恢复草稿。");
            let view = this.activeEditor();
            if (view?.file !== file || view.getMode() !== "source") {
                if (!openNote) return;
                const existing = this.app.workspace.getLeavesOfType("markdown").find(leaf =>
                    leaf.view instanceof MarkdownView && leaf.view.file === file && leaf.view.getMode() === "source");
                const leaf = existing ?? this.app.workspace.getLeaf("tab");
                if (!existing) {
                    await leaf.openFile(file, { state: { mode: "source" } });
                }
                if (!stillCurrent()) return;
                this.app.workspace.setActiveLeaf(leaf);
                view = this.activeEditor();
            }
            if (!view || view.file !== file || view.getMode() !== "source") throw new Error("请打开原笔记的编辑视图后重试。");
            const note = view.editor.getValue();
            const target = { view, file, filePath: saved.filePath, note };
            if (await sha256Text(note) !== saved.base || saved.to > note.length) throw new Error("原笔记已变化，暂时无法恢复草稿。");
            let quote: ChatQuote | null = null;
            let change: ConversationChange;
            if (saved.kind === "cleanup") {
                const entry = (await trackedOldIn(note)).find(item => item.from === saved.from && item.to === saved.to);
                if (!entry) throw new Error("原笔记已变化，暂时无法恢复草稿。");
                change = { ...target, ...entry, kind: "cleanup", replacement: "", diskBase: disk,
                    changes: [{ from: entry.from, to: entry.to, insert: "" }] };
            } else {
                const selection = saved.kind === "selection" ? paragraphSelection(note, saved.from, saved.to) : null;
                if (selection && selection.from === saved.from && selection.to === saved.to) {
                    quote = { ...target, ...selection, scope: "selection" };
                } else if (saved.kind === "document" && saved.from === 0 && saved.to === note.length) {
                    quote = { ...target, scope: "document", from: 0, to: note.length, text: note };
                }
                if (!quote) throw new Error("无法恢复草稿对应的引用。");
                change = this.editSuggestion(quote, saved.replacement, disk);
            }
            if (await this.app.vault.read(file) !== disk) throw new Error("原笔记已变化，暂时无法恢复草稿。");
            if (!stillCurrent() || !this.documentMatches(target)) return;
            if (!this.showReview(change)) {
                new Notice("编辑视图尚未准备好，请稍后重新选择此会话。");
                return;
            }
            this.quote = quote;
            this.suggestion = change;
            this.modelThread = { quote, file, messages: saved.history.slice() };
        } catch (error) {
            if (stillCurrent()) new Notice(error instanceof Error ? error.message : "无法恢复草稿，请重试。");
        } finally {
            this.restoring = false;
            if (!this.closed) this.render();
        }
    }

    private editSuggestion(quote: ChatQuote, replacement: string, diskBase: string): ConversationChange {
        const changes = documentReviewChanges(quote.text, replacement).map(item => ({
            ...item, from: item.from + quote.from, to: item.to + quote.from,
        }));
        return { ...quote, kind: "edit", replacement, changes, diskBase };
    }

    private showReview(change: ConversationChange): boolean {
        this.shownReview = null;
        try {
            const shown = this.preview?.show(change.view.editor, change.note, change.changes) ?? false;
            if (shown) this.shownReview = change;
            return shown;
        } catch {
            this.preview?.clear();
            return false;
        }
    }

    private cancelReview(): void {
        if (this.committing || this.busy || this.restoring) return;
        if (!this.forgetDraft()) return;
        this.suggestion = null;
        this.preview?.clear();
        this.addMessage({ role: "assistant", content: "已取消本次修改建议。" });
        this.render();
    }

    static captureQuote(view: MarkdownView, scope: ChatQuote["scope"] = "selection"): ChatQuote | null {
        if (!view.file || view.getMode() !== "source") return null;
        const editor = view.editor;
        const note = editor.getValue();
        const target = { note, view, file: view.file, filePath: view.file.path };
        if (scope === "document") return note.trim() ? { ...target, scope, from: 0, to: note.length, text: note } : null;
        if (editor.listSelections().length !== 1) return null;
        const selection = paragraphSelection(note, editor.posToOffset(editor.getCursor("from")), editor.posToOffset(editor.getCursor("to")));
        return selection && { ...target, ...selection, scope };
    }

    useQuote(quote: ChatQuote): void {
        if (this.committing || this.restoring || !this.documentMatches(quote)) {
            new Notice("选区已变化，请重新引用");
            return;
        }
        if (!this.forgetDraft()) return;
        this.sessionGeneration++;
        this.suggestion = null;
        this.preview?.clear();
        this.quote = quote;
        this.render();
        this.containerEl.querySelector<HTMLTextAreaElement>(".smartmd-ai-chat-input")?.focus();
    }

    private renderConversation(root: HTMLElement): HTMLElement {
        const transcript = root.createDiv({ cls: "smartmd-ai-transcript", attr: { role: "log", "aria-live": "polite" } });
        if (!this.messages.length) transcript.createEl("p", {
            text: "在笔记编辑区右键，选择编辑选中文字或整篇文档。",
            cls: "smartmd-ai-chat-hint"
        });
        for (const message of this.messages) {
            const bubble = transcript.createDiv({ cls: ["smartmd-ai-bubble", "smartmd-ai-" + message.role] });
            bubble.createDiv({ cls: "smartmd-ai-sender", text: message.role === "user" ? "你" : "AI" });
            if (message.quote) {
                bubble.createEl("small", { cls: "smartmd-ai-sent-quote", text: message.quote.filePath });
            }
            bubble.createDiv({ text: message.content });
            const sources = webSources(message.sources);
            if (message.sources) {
                const list = bubble.createDiv({ cls: "smartmd-ai-web-sources" });
                list.createEl("strong", { text: sources.length ? "搜索来源" : "已联网搜索，服务未提供来源链接" });
                for (const source of sources) {
                    list.createEl("a", { text: source.title, cls: "external-link",
                        attr: { href: source.url, target: "_blank", rel: "noopener noreferrer" } });
                }
            }
        }
        if (this.busy) transcript.createEl("p", { text: "正在回复…", cls: "smartmd-ai-chat-hint" });

        const composer = root.createDiv({ cls: "smartmd-ai-composer" });
        const saved = this.currentSession()?.draft;
        if (saved && !this.suggestion) {
            const tag = composer.createDiv({ cls: "smartmd-ai-quote" });
            tag.createEl("strong", { text: saved.filePath + " · 未应用草稿" });
            const discard = tag.createEl("button", { text: "取消草稿" });
            discard.disabled = this.busy || this.committing || this.restoring;
            discard.onclick = () => this.cancelReview();
        }
        if (this.quote) {
            const tag = composer.createDiv({ cls: "smartmd-ai-quote" });
            const editor = this.quote.view.editor;
            const line = editor.offsetToPos(this.quote.from).line + 1;
            tag.createEl("strong", { text: this.quote.filePath + (this.quote.scope === "document" ? " · 整篇文档" : " · 第 " + line + " 行") });
            const remove = tag.createEl("button", { text: "移除引用" });
            remove.disabled = this.committing || this.busy;
            remove.onclick = () => {
                if (!this.forgetDraft()) return;
                this.quote = null;
                this.suggestion = null;
                this.preview?.clear();
                this.render();
            };
        }
        if (this.suggestion) {
            const change = this.suggestion;
            const review = composer.createDiv({ cls: "smartmd-ai-batch-review", attr: { role: "group", "aria-label": "本轮修改建议" } });
            review.createEl("strong", {
                text: change.kind === "edit" ? `本轮建议 · ${change.changes.length} 处修改` : "待清理的划线旧文",
            });
            const controls = review.createDiv({ cls: "smartmd-ai-batch-actions" });
            const cancel = controls.createEl("button", { text: "取消本轮" });
            const apply = controls.createEl("button", {
                text: change.kind === "edit" ? "应用全部" : "确认删除旧文", cls: "mod-cta",
            });
            cancel.disabled = apply.disabled = this.busy || this.committing || this.restoring;
            cancel.onclick = () => { if (this.suggestion === change) this.cancelReview(); };
            apply.onclick = () => { if (this.suggestion === change) void this.accept(); };
        }
        const inputBox = composer.createDiv({ cls: "smartmd-ai-input-box" });
        const input = inputBox.createEl("textarea", {
            cls: "smartmd-ai-chat-input",
            attr: { "aria-label": "给 AI 的消息", placeholder: this.quote ? "告诉 AI 怎样修改引用内容" : "给 AI 发消息" }
        });
        input.value = this.chatInput;
        input.rows = 3;
        input.oninput = () => { this.chatInput = input.value; };
        input.onkeydown = event => {
            if (event.key === "Enter" && !event.shiftKey && !event.isComposing) {
                event.preventDefault();
                void this.sendConversation();
            }
        };
        const actions = inputBox.createDiv({ cls: "smartmd-ai-composer-actions" });
        const searchLabel = actions.createEl("label", {
            cls: "smartmd-ai-web-search",
            attr: { title: "本条消息使用模型服务的联网搜索" },
        });
        const search = searchLabel.createEl("input", { type: "checkbox" });
        search.checked = this.searchNext;
        search.disabled = this.committing || this.busy || this.restoring;
        search.onchange = () => { this.searchNext = search.checked; };
        searchLabel.createSpan({ text: "联网搜索" });
        const send = actions.createEl("button", { text: "发送", cls: "mod-cta" });
        send.disabled = this.committing || this.busy || this.restoring;
        send.onclick = () => void this.sendConversation();
        return transcript;
    }

    private documentMatches(target: DocumentTarget): boolean {
        const view = this.activeEditor();
        return view === target.view && view.file === target.file && view.getMode() === "source"
            && target.file.path === target.filePath && this.app.vault.getAbstractFileByPath(target.filePath) === target.file
            && view.editor.getValue() === target.note;
    }

    private async sendConversation(): Promise<void> {
        const input = this.chatInput.trim();
        if (!input || this.busy || this.committing || this.restoring) return;
        const command = input.replace(/[。！!]+$/, "");
        const session = this.currentSession();
        if (((this.suggestion || session?.draft) && /^(?:应用|应用到笔记|应用全部|确认应用|取消|取消修改|取消本轮)$/.test(command))
            || (this.undoRecord?.sessionId === this.sessionId && /^(?:撤销|撤销本次 AI 修改)$/.test(command))) {
            this.addMessage({ role: "user", content: input });
            this.chatInput = "";
            this.searchNext = false;
            if (command.startsWith("撤销")) await this.undo();
            else if (command.startsWith("取消")) this.cancelReview();
            else if (this.suggestion) await this.accept();
            else {
                await this.restoreDraft();
                if (this.suggestion) new Notice("请先查看笔记里的修改建议，再应用。");
            }
            this.render();
            return;
        }
        const quote = this.quote;
        const listEdit = quote?.scope === "selection" && Boolean(quote.paragraph.listLines || quote.paragraph.listPrefix);
        if (quote && !this.documentMatches(quote)) return void this.invalidate();
        const view = this.activeEditor();
        const target = view?.file && view.getMode() === "source"
            ? { view, file: view.file, filePath: view.file.path, note: view.editor.getValue() } : null;
        const generation = this.sessionGeneration;
        const requestContext = target && { file: target.file, note: target.note, matches: () => this.documentMatches(target), invalidated: false };
        const draft = this.suggestion?.kind === "edit" ? this.suggestion.replacement : undefined;
        const file = target?.file ?? null;
        if (!this.modelThread || this.modelThread.quote !== quote || this.modelThread.file !== file) {
            const prior = session?.context;
            const messages = !quote && prior?.filePath === (file?.path ?? null) ? prior.messages.slice() : [];
            this.modelThread = { quote, file, messages };
        }
        const thread = this.modelThread;
        const history = thread.messages.slice(-8);
        const search = this.searchNext;
        this.searchNext = false;
        this.addMessage({ role: "user", content: input, quote: quote ? { filePath: quote.filePath } : undefined });
        this.chatInput = "";
        this.busy = true;
        this.activeRequestContext = requestContext;
        this.render();
        try {
            if (target) await this.awaitDiskObservation(target.file);
            if (generation !== this.sessionGeneration || quote && !this.documentMatches(quote)) return;
            const diskBase = target ? await this.app.vault.read(target.file) : "";
            const entries = !quote && target ? await trackedOldIn(target.note) : [];
            if (generation !== this.sessionGeneration || quote && !this.documentMatches(quote)) return;
            const reply = await requestConversation(modelSettings(this.store), history, {
                instruction: input, reference: quote?.text,
                scope: quote?.scope === "document" ? "document" : listEdit ? "list" : undefined,
                proposedDraft: quote ? draft : undefined,
                trackedOldIds: entries.map(entry => entry.id),
            }, search);
            if (target) await this.awaitDiskObservation(target.file);
            if (generation !== this.sessionGeneration) return;
            thread.messages.push({ role: "user", content: input });
            if (reply.action === "say") {
                thread.messages.push({ role: "assistant", content: reply.text });
                if (session && !quote && !requestContext?.invalidated) {
                    session.context = { filePath: file?.path ?? null, messages: thread.messages.slice(-8) };
                }
                if (session?.draft && this.suggestion && !requestContext?.invalidated) {
                    session.draft.history = thread.messages.slice(-8);
                }
                this.addMessage({ role: "assistant", content: reply.text, sources: search ? reply.sources : undefined });
                return;
            }
            let change: ConversationChange;
            if (reply.action === "edit") {
                if (!quote || this.quote !== quote || !this.documentMatches(quote)) throw new Error("引用已变化，请重新选择要编辑的文字");
                change = this.editSuggestion(quote, reply.text, diskBase);
            } else {
                const entry = entries.find(item => item.id === reply.id);
                if (!target || !entry || requestContext?.invalidated || !this.documentMatches(target)) {
                    throw new Error("找不到可安全清理的旧文，请确认当前笔记");
                }
                change = { ...target, ...entry, kind: "cleanup", replacement: "", diskBase,
                    changes: [{ from: entry.from, to: entry.to, insert: "" }] };
            }
            const base = change.changes.length ? await sha256Text(change.note) : "";
            await this.awaitDiskObservation(change.file);
            if (generation !== this.sessionGeneration || requestContext?.invalidated || !this.documentMatches(change)) {
                throw new Error("笔记已变化，请重新发起修改");
            }
            if (!change.changes.length) {
                if (!this.forgetDraft()) return;
                this.suggestion = null;
                this.preview?.clear();
                this.addMessage({ role: "assistant", content: "这次建议与原文相同。", sources: search ? reply.sources : undefined });
                return;
            }
            this.suggestion = change;
            if (session) {
                session.draft = { filePath: change.filePath, base, from: change.from, to: change.to,
                    kind: change.scope ?? "cleanup", replacement: change.replacement, history: thread.messages.slice(-8) };
                if (!this.saveHistory()) new Notice("本机存储写入失败，此草稿可能无法在重启后恢复。", 7000);
            }
            const shown = this.showReview(change);
            this.addMessage({ role: "assistant", content: shown
                ? "建议已标在笔记里。可以继续告诉我要改哪里，确认后用输入框上方的操作区处理本轮建议。" : "建议已生成，请打开原笔记的编辑视图查看。",
                sources: search ? reply.sources : undefined });
        } catch (error) {
            if (generation === this.sessionGeneration) {
                this.addMessage({ role: "assistant", content: error instanceof Error ? error.message : "这次请求无法完成，请重试" });
            }
        } finally {
            if (this.activeRequestContext === requestContext) this.activeRequestContext = null;
            this.busy = false;
            this.render();
        }
    }

    private render(jumpToLatest = false): void {
        if (this.closed) return;
        const root = this.containerEl.children[1] as HTMLElement;
        const summary = root.querySelector(".smartmd-ai-history > summary");
        const keepHistoryFocus = summary === root.ownerDocument.activeElement;
        const previous = root.querySelector<HTMLElement>(".smartmd-ai-transcript");
        const displayed = this.displayedTranscript;
        const changed = displayed?.sessionId !== this.sessionId || displayed.messageCount !== this.messages.length;
        const state = {
            sessionId: this.sessionId,
            messageCount: this.messages.length,
            scrollTop: previous?.clientHeight ? previous.scrollTop : displayed?.scrollTop ?? 0,
            followLatest: jumpToLatest || changed
                || (previous?.clientHeight ? atTranscriptEnd(previous) : displayed?.followLatest ?? true),
        };
        this.displayedTranscript = state;
        this.transcriptResize?.disconnect();
        root.empty();
        root.addClass("suite-ai-view", "suite-ai-conversation");
        root.createEl("h2", { text: "AI 编辑" });
        const newSummary = this.renderHistory(root.createDiv({ cls: "smartmd-ai-history-panel" }));
        const transcript = this.renderConversation(root);
        if (keepHistoryFocus) newSummary.focus({ preventScroll: true });
        let positioned = false;
        const positionTranscript = () => {
            if (!transcript.clientHeight) return;
            if (state.followLatest || !positioned) {
                transcript.scrollTop = state.followLatest ? transcript.scrollHeight : state.scrollTop;
                state.scrollTop = transcript.scrollTop;
            }
            positioned = true;
        };
        transcript.onscroll = () => {
            if (!positioned || !transcript.clientHeight) return;
            state.scrollTop = transcript.scrollTop;
            state.followLatest = atTranscriptEnd(transcript);
        };
        positionTranscript();
        this.transcriptResize = new ResizeObserver(positionTranscript);
        this.transcriptResize.observe(transcript);
    }

    private renderHistory(panel: HTMLElement): HTMLElement {
        const busy = this.busy || this.committing || this.restoring;
        const history = panel.createEl("details", { cls: "smartmd-ai-history" });
        history.open = this.historyOpen;
        history.ontoggle = () => { if (history.isConnected) this.historyOpen = history.open; };
        const summary = history.createEl("summary", { attr: { "aria-label": "会话历史" } });
        summary.createSpan({ text: "会话历史" });
        const current = this.currentSession();
        if (current) {
            const title = current.draft && !this.suggestion
                ? sessionTitle(current) + (this.restoring ? " · 恢复中" : " · 待恢复")
                : sessionTitle(current);
            summary.createSpan({ cls: "smartmd-ai-current-session", text: title, attr: { title } });
        }
        const sessions = history.createEl("ul", { cls: "smartmd-ai-history-list" });
        for (const session of this.sessions) {
            const title = sessionTitle(session);
            const row = sessions.createEl("li", { cls: "smartmd-ai-history-row", attr: { "data-session-id": session.id } });
            const open = row.createEl("button", { text: title, cls: "smartmd-ai-session-open", attr: { title } });
            if (session === current) open.setAttribute("aria-current", "true");
            open.disabled = busy;
            open.onclick = () => {
                this.historyOpen = history.open = false;
                this.selectSession(session);
            };
            const remove = row.createEl("button", { text: "删除", cls: "smartmd-ai-session-delete",
                attr: { "aria-label": "删除会话 " + title } });
            remove.disabled = busy;
            remove.onclick = () => this.confirmDeleteSession(session);
        }
        const newChat = panel.createEl("button", { text: "新建会话" });
        newChat.disabled = busy;
        newChat.onclick = () => {
            this.historyOpen = false;
            this.createSession();
        };
        return summary;
    }
    private activeEditor(): MarkdownView | null {
        const workspace = this.app.workspace;
        const active = workspace.getActiveViewOfType(MarkdownView);
        if (active) return active;
        if (workspace.getActiveViewOfType(AIEditorView) !== this) return null;
        const view = workspace.getMostRecentLeaf(workspace.rootSplit)?.view;
        return view instanceof MarkdownView ? view : null;
    }

    private invalidateFile(file: unknown): void {
        const request = this.activeRequestContext;
        if (request && request.file === file) request.invalidated = true;
        for (const draft of this.sessionDrafts.values()) {
            if (draft.quote?.file === file) {
                draft.quote = null;
                draft.thread = null;
            }
        }
        if (this.quote?.file === file || this.suggestion?.file === file) this.invalidate();
    }

    private observeVaultModification(file: unknown): void {
        if (!(file instanceof TFile)) return;
        const quote = this.quote?.file === file ? this.quote : null;
        const change = this.suggestion?.file === file ? this.suggestion : null;
        const request = this.activeRequestContext?.file === file ? this.activeRequestContext : null;
        const note = quote?.note ?? change?.note ?? request?.note;
        const dormant = [...this.sessionDrafts.values()].filter(draft => draft.quote?.file === file);
        if (note === undefined && !dormant.length) return;
        const observation = { file, done: Promise.resolve() };
        const invalidateDifferentDisk = (disk?: string) => {
            let restoredDraftChanged = false;
            for (const draft of dormant) {
                const original = draft.quote?.note;
                if (disk === original) continue;
                if (draft.quote && this.quote === draft.quote) {
                    restoredDraftChanged = true;
                }
                draft.quote = null;
                draft.thread = null;
            }
            if (disk !== note && request && this.activeRequestContext === request) request.invalidated = true;
            if (restoredDraftChanged || (disk !== note && (quote && this.quote === quote || change && this.suggestion === change))) {
                this.invalidate();
            }
        };
        observation.done = this.app.vault.read(file).then(invalidateDifferentDisk, () => invalidateDifferentDisk()).finally(() => {
            if (this.diskObservation === observation) this.diskObservation = null;
        });
        this.diskObservation = observation;
    }

    private async awaitDiskObservation(file: TFile): Promise<void> {
        while (this.diskObservation?.file === file) await this.diskObservation.done;
    }

    private async accept(): Promise<void> {
        if (this.committing || this.busy) return;
        const conversation = this.suggestion;
        if (!conversation || (conversation.kind === "edit" && conversation.replacement === conversation.text)) return;
        if (!this.documentMatches(conversation)) return void this.invalidate();
        if (this.shownReview !== conversation) {
            if (this.showReview(conversation)) new Notice("请先查看笔记里的修改建议，再应用。");
            else new Notice("请打开原笔记的编辑视图查看建议。");
            return;
        }
        const { view, file, filePath, note: before } = conversation;
        const editor = view.editor;
        const publish = this.runtime?.publish;
        const pendingBefore = publish?.isAiPending?.(filePath) === true;
        this.committing = true;
        this.render();
        try {
            // Freeze publication in memory and on disk before the editor changes.
            const approval = publish?.markAiPending?.(filePath);
            await approval?.persisted;
            await this.awaitDiskObservation(file);
            const diskNow = await this.app.vault.read(file);
            if (this.suggestion !== conversation || !this.documentMatches(conversation)
                || (diskNow !== conversation.diskBase && diskNow !== before)) {
                if (!pendingBefore && approval) await publish.clearAiPending(filePath, approval.revision);
                if (this.suggestion === conversation) this.invalidate();
                return;
            }
            this.suggestion = null;
            this.quote = null;
            this.preview?.clear();
            if (!this.preview?.replaceChanges(editor, before, conversation.changes)) {
                if (!pendingBefore && approval && editor.getValue() === before) await publish.clearAiPending(filePath, approval.revision);
                new Notice("编辑器未能按预览应用修改，请重新打开笔记后重试。", 7000);
                return;
            }
            this.undoRecord = {
                sessionId: this.sessionId, file, filePath, before, after: applyReviewChanges(before, conversation.changes), pendingBefore,
                publishRevision: approval?.revision, reverse: reverseReviewChanges(before, conversation.changes),
            };
            try {
                await view.save();
                const stored = await this.app.vault.read(file);
                if (this.app.vault.getAbstractFileByPath(filePath) !== file || stored !== this.undoRecord.after) {
                    throw new Error("Saved file does not match the applied change");
                }
                this.forgetDraft();
                new Notice("已应用并保存修改");
                this.addMessage({ role: "assistant", content: conversation.kind === "edit"
                    ? "已应用并保存修改。可以在这里说「撤销」。" : "已删除选定的划线旧文，保留了新文。" });
            } catch {
                new Notice("修改已应用到编辑器，尚未确认保存，请检查当前笔记", 7000);
                this.addMessage({ role: "assistant", content: "改动已应用到编辑器，保存尚未确认。" });
            }
        } catch {
            new Notice("无法完成 AI 修改，请检查插件数据保存状态", 7000);
        } finally {
            this.committing = false;
            this.render();
        }
    }

    private invalidate(): void {
        if (this.quote || this.suggestion) this.sessionGeneration++;
        this.quote = null;
        this.suggestion = null;
        this.preview?.clear();
        this.render();
        if (this.currentSession()?.draft) new Notice("当前编辑位置已变化，选择原会话可恢复草稿。", 5500);
    }

    private observeContextChange(): void {
        const request = this.activeRequestContext;
        if (request && !request.matches()) request.invalidated = true;
        if (this.quote && !this.documentMatches(this.quote)
            || this.suggestion && !this.documentMatches(this.suggestion)) return void this.invalidate();
    }

    private async undo(): Promise<void> {
        if (this.committing || this.busy) return;
        const record = this.undoRecord;
        const view = this.activeEditor();
        if (!record || record.sessionId !== this.sessionId || view?.file !== record.file || view.file.path !== record.filePath
            || this.app.vault.getAbstractFileByPath(record.filePath) !== record.file || view.editor.getValue() !== record.after
            || view.getMode() !== "source") {
            return void new Notice("文档之后又发生了变化，无法安全撤销本次 AI 修改");
        }
        this.committing = true;
        this.suggestion = null;
        this.quote = null;
        this.preview?.clear();
        this.render();
        let restored = false;
        try {
            const saved = await this.app.vault.read(record.file);
            if (saved !== record.after || this.activeEditor() !== view || view.file !== record.file
                || this.app.vault.getAbstractFileByPath(record.filePath) !== record.file
                || view.getMode() !== "source" || view.editor.getValue() !== record.after) {
                new Notice("文档之后又发生了变化，无法安全撤销本次 AI 修改");
                return;
            }
            if (!this.preview?.replaceChanges(view.editor, record.after, record.reverse)) throw new Error("Could not restore the reviewed edit");
            this.undoRecord = null;
            await view.save();
            const stored = await this.app.vault.read(record.file);
            restored = this.app.vault.getAbstractFileByPath(record.filePath) === record.file && stored === record.before;
            if (restored && view.file === record.file && view.editor.getValue() === record.before
                && !record.pendingBefore && record.publishRevision) {
                await this.runtime?.publish?.clearAiPending?.(record.filePath, record.publishRevision);
            }
        } catch {
            // Retain publication approval if restoration was not verified on disk.
        } finally {
            this.committing = false;
        }
        if (restored) this.addMessage({ role: "assistant", content: "已撤销本次 AI 修改。" });
        this.render();
        new Notice(restored ? "已撤销本次 AI 修改" : "未能确认撤销已保存，请检查当前笔记", restored ? undefined : 7000);
    }
}
