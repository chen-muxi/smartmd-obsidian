import { EditorSelection, EditorState, StateEffect, StateField, type Extension } from "@codemirror/state";
import { Decoration, type DecorationSet, EditorView, ViewPlugin, WidgetType } from "@codemirror/view";
import { editorInfoField, sanitizeHTMLToDom, type Editor } from "obsidian";
import { marked } from "marked";
import { applyReviewChanges, type ReviewChange } from "./ai-review-blocks";

const showPreview = StateEffect.define<DecorationSet>();
const ignoreListNumbers = (note: string) => note.replace(/^[ \t]*(?:(?:[-+*]|\d{1,9}[.)])(?:[ \t]+|$))+/gm,
    prefix => prefix.replace(/\d+(?=[.)])/g, "0"));
const previewField = StateField.define<DecorationSet>({
    create: () => Decoration.none,
    update(value, transaction) {
        if (transaction.docChanged) value = Decoration.none;
        for (const effect of transaction.effects) if (effect.is(showPreview)) value = effect.value;
        return value;
    },
    provide: field => EditorView.decorations.from(field),
});

const previewTags = new Set("p h1 h2 h3 h4 h5 h6 blockquote ul ol li table thead tbody tr th td pre code strong em del br hr".split(" "));

function renderDraft(root: HTMLElement, text: string): void {
    const preview = sanitizeHTMLToDom(marked.parse(text, { gfm: true, async: false }));
    const append = (container: Node, nodes: NodeListOf<ChildNode>) => {
        for (const node of Array.from(nodes)) {
            if (node.nodeType === 3) {
                container.appendChild(root.ownerDocument.createTextNode(node.textContent ?? ""));
            } else if (node.nodeType === 1) {
                const tag = (node as HTMLElement).localName;
                if (previewTags.has(tag)) {
                    const child = root.ownerDocument.createElement(tag);
                    append(child, node.childNodes);
                    container.appendChild(child);
                } else if (tag !== "script" && tag !== "style") {
                    append(container, node.childNodes);
                }
            }
        }
    };
    append(root, preview.childNodes);
}

class NewTextWidget extends WidgetType {
    constructor(private content: string) { super(); }
    eq(other: NewTextWidget): boolean { return this.content === other.content; }
    toDOM(view: EditorView): HTMLElement {
        const block = view.dom.ownerDocument.createElement("div");
        block.className = "smartmd-ai-block-new";
        block.setAttribute("aria-label", "AI 修改建议");
        const body = block.createDiv({ cls: "smartmd-ai-block-body" });
        if (this.content) {
            if (this.content.trim()) renderDraft(body, this.content);
            else body.setText("建议修改空行或空格");
            const source = block.createEl("details", { cls: "smartmd-ai-block-source" });
            source.open = !this.content.trim() || (!body.textContent?.trim() && !body.querySelector("hr"));
            source.createEl("summary", { text: "查看 Markdown" });
            source.createEl("pre", { text: this.content });
        } else {
            body.setText("建议删除此处");
        }
        return block;
    }
}

export class AIChatPreview {
    private views = new Set<EditorView>();
    private active: EditorView | null = null;
    readonly extension: Extension = [previewField, ViewPlugin.define(view => {
        this.views.add(view);
        return { destroy: () => {
            this.views.delete(view);
            if (this.active === view) this.active = null;
        } };
    })];

    show(editor: Editor, note: string, changes: ReviewChange[]): boolean {
        this.clear();
        const view = Array.from(this.views).find(candidate => candidate.state.field(editorInfoField, false)?.editor === editor);
        if (!view || view.state.doc.toString() !== note) return false;
        if (!changes.length) return false;
        const ranges = changes.map(change => {
            const length = note.slice(change.from, change.to).trimEnd().length;
            const at = length ? change.from + length - 1 : Math.max(0, change.from - 1);
            const atStart = change.from === 0 && change.to === 0;
            const anchor = atStart ? 0 : view.state.doc.lineAt(at).to;
            return Decoration.widget({ block: true, side: atStart ? -1 : 1,
                widget: new NewTextWidget(change.insert) }).range(anchor);
        });
        // The widget follows the last line of the original block, which can extend beyond the viewport.
        const firstPreview = EditorSelection.cursor(ranges[0].from, -1);
        view.dispatch({
            effects: [showPreview.of(Decoration.set(ranges, true)), EditorView.scrollIntoView(firstPreview, { y: "start", yMargin: 32 })],
        });
        this.active = view;
        return true;
    }

    replaceChanges(editor: Editor, note: string, changes: ReviewChange[]): boolean {
        const view = Array.from(this.views).find(candidate => candidate.state.field(editorInfoField, false)?.editor === editor);
        if (!view || view.state.readOnly || !view.contentDOM.isContentEditable || view.state.doc.toString() !== note) return false;
        const expected = applyReviewChanges(note, changes);
        let transaction = view.state.update({ changes, userEvent: "smartmd.edit" });
        const filtered = transaction.newDoc.toString();
        if (!transaction.docChanged) return false;
        if (filtered !== expected) {
            // Obsidian's list input filter may renumber even items outside a programmatic edit.
            if (ignoreListNumbers(filtered) !== ignoreListNumbers(expected)) return false;
            transaction = view.state.update({ changes, userEvent: "smartmd.edit", filter: false });
            // Preserve range-level change vetoes even when the exact edit cannot use transaction rewrites.
            for (const filter of view.state.facet(EditorState.changeFilter)) {
                const result = filter(transaction);
                if (result !== true && (result === false || result.length > 0)) return false;
            }
        }
        if (transaction.newDoc.toString() !== expected) return false;
        view.dispatch(transaction);
        return editor.getValue() === expected;
    }

    clear(): void {
        const view = this.active;
        this.active = null;
        // Obsidian can deliver editor-change during a CodeMirror update.
        if (view) queueMicrotask(() => {
            if (view !== this.active && this.views.has(view) && view.state.field(previewField, false)?.size) {
                view.dispatch({ effects: showPreview.of(Decoration.none) });
            }
        });
    }
}
