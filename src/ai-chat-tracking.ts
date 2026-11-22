import { lexer, type Tokens } from "marked";
import { sha256Text } from "./core";
import { isParagraphReplacement, paragraphAt, type ParagraphRange } from "./ai-paragraph";

export interface TextRange {
    from: number;
    to: number;
    text: string;
}

export type ChatReply =
    | { action: "say"; text: string }
    | { action: "edit"; text: string }
    | { action: "cleanup"; id: string };

export interface TrackedOld extends TextRange {
    id: string;
    time: number;
    old: string;
}

const trackerPattern = /<!--smartmd-ai-old:([a-f0-9-]{36}):([a-z0-9]{1,12}):([1-9]\d{0,5}):([a-f0-9]{64})-->/;
const trackerMarker = new RegExp(trackerPattern.source + " ", "g");
const isTrackerComment = (raw: string) => raw.match(trackerPattern)?.[0] === raw;

export function parseChatReply(content: string): ChatReply {
    const fenced = /^```(?:json)?[ \t]*\n([\s\S]+)\n```$/i.exec(content);
    try {
        const payload = JSON.parse(fenced?.[1] ?? content);
        if (typeof payload?.text === "string" && (payload.action === "edit" || (payload.action === "say" && payload.text.trim()))) {
            const text = payload.action === "edit" ? payload.text : payload.text.trim();
            return { action: payload.action, text: text.replace(/\r\n?/g, "\n") };
        }
        if (payload?.action === "cleanup" && typeof payload.id === "string") return { action: "cleanup", id: payload.id };
    } catch {
        // A response outside the action contract never creates a write.
    }
    throw new Error("模型没有返回可审阅的聊天结果，请重试");
}

export function paragraphSelection(note: string, from: number, to: number): (TextRange & { paragraph: ParagraphRange }) | null {
    const selected = note.slice(from, to);
    from += selected.length - selected.trimStart().length;
    to -= selected.length - selected.trimEnd().length;
    if (to <= from) return null;
    const paragraph = paragraphAt(note, from, isTrackerComment);
    if (!paragraph || to > paragraph.to) return null;
    // Leave the first line's syntax in place when binding the editor selection.
    if (paragraph.listLines) {
        const selectedLines = paragraph.listLines.filter(line => !line.protected && Math.max(from, line.body) < Math.min(to, line.to));
        const firstLine = selectedLines[0];
        const lastLine = selectedLines[selectedLines.length - 1];
        if (!firstLine || !lastLine) return null;
        from = Math.max(from, firstLine.body);
        to = from + note.slice(from, Math.min(to, lastLine.to)).trimEnd().length;
    } else if (paragraph.listPrefix) {
        from = Math.max(from, paragraph.from);
    }
    if (from < paragraph.from || to <= from) return null;
    for (const match of paragraph.text.matchAll(trackerMarker)) {
        const markerStart = paragraph.from + match.index!;
        if (paragraph.listLines?.some(line => line.protected && line.from <= markerStart && markerStart < line.to)) continue;
        if (from < markerStart + match[0].length && to > markerStart - Number(match[3])) return null;
    }
    const text = note.slice(from, to);
    const hasStrikethrough = paragraph.listLines ? paragraph.listLines.some(line => {
        const start = Math.max(from, line.body);
        const end = Math.min(to, line.to);
        return !line.protected && start < end && note.slice(start, end).includes("~~");
    }) : text.includes("~~");
    if (hasStrikethrough) return null;
    return { from, to, text, paragraph };
}

function hasTrackedMarkAt(text: string, paragraph: Tokens.Paragraph, at: number, struck: string, marker: string): boolean {
    let offset = 0;
    for (let index = 0; index < paragraph.tokens.length; index++) {
        const inline = paragraph.tokens[index];
        if (!text.startsWith(inline.raw, offset) || offset > at) return false;
        if (offset === at) {
            const next = paragraph.tokens[index + 1];
            return inline.type === "del" && inline.raw === struck && next?.type === "html" && next.raw === marker;
        }
        offset += inline.raw.length;
    }
    return false;
}

export function validatedReplacement(
    target: TextRange & { paragraph: ParagraphRange }, note: string, proposed: string
): string {
    const paragraph = note.slice(target.paragraph.from, target.from) + proposed + note.slice(target.to, target.paragraph.to);
    if (proposed === target.text) throw new Error("新文与引用原文相同");
    if (!isParagraphReplacement(note, target.paragraph, paragraph, isTrackerComment)) {
        throw new Error("建议会改变段落、列表结构或受保护的 Markdown 内容，请重新指示");
    }
    return proposed;
}

export async function trackedOldIn(note: string): Promise<TrackedOld[]> {
    // Recognize the previous build's saved marks for explicit cleanup. New edits never create these marks.
    if (!note.includes("<!--smartmd-ai-old:")) return [];
    const found: TrackedOld[] = [];
    let offset = 0;
    for (const token of lexer(note, { gfm: true })) {
        if (!note.startsWith(token.raw, offset)) break;
        if (token.type === "paragraph") {
            for (const match of token.raw.matchAll(trackerMarker)) {
                const start = match.index! - Number(match[3]);
                if (start < 0) continue;
                const struck = token.raw.slice(start, match.index);
                if (!struck.startsWith("~~") || !struck.endsWith("~~")
                    || !hasTrackedMarkAt(token.raw, token as Tokens.Paragraph, start, struck, match[0].trimEnd())
                    || !paragraphAt(note, offset + start, isTrackerComment)
                    || await sha256Text(match[1] + ":" + struck) !== match[4]) continue;
                const end = match.index! + match[0].length;
                found.push({ id: match[1], time: parseInt(match[2], 36), old: struck.slice(2, -2),
                    from: offset + start, to: offset + end, text: token.raw.slice(start, end) });
            }
        }
        offset += token.raw.length;
    }
    const count = new Map<string, number>();
    for (const entry of found) count.set(entry.id, (count.get(entry.id) ?? 0) + 1);
    return found.filter(entry => count.get(entry.id) === 1).sort((a, b) => b.time - a.time);
}
