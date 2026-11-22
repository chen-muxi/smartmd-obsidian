import { lexer, walkTokens, type Tokens } from "marked";

export interface ParagraphRange {
    from: number;
    to: number;
    text: string;
    protectedMarkup: string[];
    listPrefix?: string;
    listLines?: ListLine[];
}

interface ListLine {
    from: number;
    body: number;
    to: number;
    syntax: string;
    protected?: boolean;
}

const definitionLine = /^\[[^\]\n]+\]:/;
const listLinePrefix = /^[ \t]*(?:(?:(?:[-+*]|\d{1,9}[.)])(?:[ \t]+|$))+(?:\[[^\]\n]\](?:[ \t]+|$))?)?/u;

function insideDelimitedBlock(note: string, from: number, to: number, marker: string): boolean {
    let open = false;
    for (let index = note.indexOf(marker); index >= 0 && index < to; index = note.indexOf(marker, index + marker.length)) {
        if (index + marker.length > from || note[index - 1] === "\\") return true;
        open = !open;
    }
    return open;
}

function paragraphMarkup(tokens: Tokens.Paragraph["tokens"], text: string, allowHtml?: (raw: string) => boolean): string[] | null {
    const protectedMarkup = [
        ...(text.match(/\\*!?\[\[[^\n]*?\]\]|\\*\[\^[^\]\n]+\]|\\*\^\[[^\]\n]*\]|\\*\$(?!\$)(?:\\.|[^$\\\n])+\$/g) ?? []),
    ];
    const trailingId = /(?:^|[ \t])(\\*\^[\w-]+)[ \t]*$/.exec(text);
    if (trailingId) protectedMarkup.push(trailingId[1]);
    const pending = [...tokens];
    while (pending.length) {
        const inline = pending.pop()!;
        if (inline.type === "html") {
            if (!allowHtml?.(inline.raw)) return null;
            protectedMarkup.push(inline.raw);
        } else if (inline.type === "link" || inline.type === "image" || inline.type === "codespan") {
            protectedMarkup.push(inline.raw);
        } else if ("tokens" in inline && Array.isArray(inline.tokens)) {
            pending.push(...inline.tokens);
        }
    }
    return protectedMarkup.sort();
}

function hasAmbiguousDelimiter(token: Tokens.Generic): boolean {
    let ambiguous = false;
    walkTokens([token], child => {
        if ((child.type === "code" || child.type === "codespan" || child.type === "html") && /%%|\$\$/.test(child.raw)) {
            ambiguous = true;
        }
    });
    return ambiguous;
}

type ListBlockMarkup = string | [string, string[]] | ListMarkup;
type ListMarkup = [boolean, boolean, ListBlockMarkup[][]];
type ProtectedListBlock = { type: "code" | "blockquote"; raw: string; maskDelimiters: boolean };
const quoteDelimiter = /%%|\$\$/;

function quoteDelimitersAreClosed(quote: Tokens.Blockquote): boolean {
    if (!quoteDelimiter.test(quote.raw) || hasAmbiguousDelimiter(quote)) return false;
    if (quote.raw.match(/\[\[[^\n]*?\]\]/g)?.some(link => quoteDelimiter.test(link))) return false;
    if (/(?:^|[^$])\$(?!\$)[^$\n]*%%[^$\n]*\$(?!\$)/m.test(quote.raw)) return false;
    let nestedDelimiter = false;
    walkTokens([quote], token => {
        if ((token.type === "link" || token.type === "image" || token.type === "list") && quoteDelimiter.test(token.raw)) {
            nestedDelimiter = true;
        }
    });
    if (nestedDelimiter) return false;

    let open: { marker: string; depth: number; line: number; standalone: boolean } | undefined;
    for (const [index, row] of quote.raw.trimEnd().split("\n").entries()) {
        const prefix = /^(?:[ \t]*>[ \t]?)+/.exec(row)?.[0] ?? "";
        const depth = [...prefix].filter(character => character === ">").length;
        if (open && depth !== open.depth) return false;
        const body = row.slice(prefix.length);
        for (const match of body.matchAll(/%%|\$\$/g)) {
            const marker = match[0];
            const at = match.index!;
            if (!depth || body[at - 1] === "\\" || body[at - 1] === marker[0] || body[at + marker.length] === marker[0]) return false;
            const standalone = body.trim() === marker;
            if (open) {
                if (open.marker !== marker || (marker === "$$" && open.line !== index && (!open.standalone || !standalone))) return false;
                open = undefined;
            } else {
                open = { marker, depth, line: index, standalone };
            }
        }
    }
    return !open;
}

function hasUnresolvedQuoteDelimiters(token: Tokens.Generic): boolean {
    let unresolved = false;
    walkTokens([token], child => {
        if (child.type === "blockquote" && quoteDelimiter.test(child.raw) && !quoteDelimitersAreClosed(child as Tokens.Blockquote)) {
            unresolved = true;
        }
    });
    return unresolved;
}

function listMarkup(list: Tokens.List, protectedBlocks: ProtectedListBlock[], allowHtml?: (raw: string) => boolean): ListMarkup | null {
    const structure: ListBlockMarkup[][] = [];
    for (const item of list.items) {
        const blocks: ListBlockMarkup[] = [listLinePrefix.exec(item.raw)![0]];
        for (const [index, block] of item.tokens.entries()) {
            if (block.type === "list") {
                const nested = listMarkup(block as Tokens.List, protectedBlocks, allowHtml);
                if (nested === null) return null;
                blocks.push(nested);
            } else if (block.type === "paragraph" || block.type === "text") {
                if (!block.tokens) return null;
                // Marked leaves a checkbox in loose-list paragraph.raw but removes it in tight lists.
                const inlineCheckbox = block.tokens[0];
                const marker = inlineCheckbox?.type === "checkbox" ? inlineCheckbox.raw
                    : !item.task && index === 0 ? /^\[[^\]\n]\][ \t]+/u.exec(block.raw)?.[0] : undefined;
                const taskText = marker ? block.raw.slice(marker.length)
                    : item.task && index === 1 && item.tokens[0].type === "checkbox" ? block.raw : undefined;
                if (taskText !== undefined && lexer(taskText, { gfm: true })[0]?.type !== "paragraph") return null;
                const markup = paragraphMarkup(block.tokens, block.raw, allowHtml);
                if (!markup) return null;
                blocks.push([block.type, markup]);
            } else if (block.type === "space" || block.type === "checkbox") {
                blocks.push(block.type);
            } else if (block.type === "code" || block.type === "blockquote") {
                protectedBlocks.push({
                    type: block.type,
                    raw: block.raw,
                    maskDelimiters: block.type === "code" || quoteDelimitersAreClosed(block as Tokens.Blockquote),
                });
                blocks.push([block.type, [block.raw]]);
            } else {
                return null;
            }
        }
        structure.push(blocks);
    }
    return [list.ordered, list.loose, structure];
}

function protectListSource(note: string, lines: ListLine[], blocks: ProtectedListBlock[]): Set<number> | null {
    const maskedLines = new Set<number>();
    if (!blocks.length) return maskedLines;

    const positions = new Map<string, number[]>();
    const source = lines.map((line, index) => {
        const content = note.slice(line.from, line.to);
        const start = content.trimStart();
        const body = note.slice(line.body, line.to).trimStart();
        for (const key of new Set([start, body])) {
            const found = positions.get(key);
            if (found) found.push(index);
            else positions.set(key, [index]);
        }
        return content;
    });
    const normalized = source.map(line => line.trimStart());
    const groups = new Map<string, { type: ProtectedListBlock["type"]; rows: string[]; count: number; maskDelimiters: boolean }>();
    for (const block of blocks) {
        const rows = block.raw.replace(/\n+$/, "").split("\n").map(line => line.trimStart());
        const key = JSON.stringify(rows);
        const group = groups.get(key);
        if (group) {
            if (group.type !== block.type) return null;
            group.count++;
            group.maskDelimiters &&= block.maskDelimiters;
        } else {
            groups.set(key, { type: block.type, rows, count: 1, maskDelimiters: block.maskDelimiters });
        }
    }
    // Match outer fences and longer blocks first, so their examples or tail lines cannot match another block.
    const orderedGroups = Array.from(groups.values()).sort((left, right) =>
        Number(right.type === "code") - Number(left.type === "code") || right.rows.length - left.rows.length);
    for (const { type, rows, count, maskDelimiters } of orderedGroups) {
        const matches = (positions.get(rows[0]) ?? []).filter(start =>
            rows.length <= lines.length - start && rows.every((row, index) =>
                !lines[start + index].protected && (index === 0 || row === normalized[start + index])));
        if (matches.length !== count) return null;
        for (const start of matches) {
            for (let index = start; index < start + rows.length; index++) {
                const line = lines[index];
                if (line.protected) return null;
                line.protected = true;
                // The lexer removes container indentation. Preserve the actual protected source line too.
                line.syntax = JSON.stringify([line.syntax, type, source[index]]);
                // Ignore delimiters only when code or a self-contained quote owns these exact lines.
                if (maskDelimiters) maskedLines.add(index);
            }
        }
    }
    return maskedLines;
}

function listBlockAt(note: string, list: Tokens.List, from: number, cursor: number, allowHtml?: (raw: string) => boolean): ParagraphRange | null {
    const text = list.raw.replace(/\n+$/, "");
    const to = from + text.length;
    if (cursor > to) return null;
    const protectedBlocks: ProtectedListBlock[] = [];
    const markup = listMarkup(list, protectedBlocks, allowHtml);
    if (markup === null) return null;
    let offset = from;
    const rows = text.split("\n");
    const listLines: ListLine[] = rows.map(line => {
        const prefix = listLinePrefix.exec(line)![0];
        const body = offset + prefix.length;
        const end = offset + line.length;
        // Preserve indentation, markers, hard line breaks and block IDs at their source lines.
        const suffix = /(?:(?:^|[ \t])\\*\^[\w-]+[ \t]*| {2,}|\\)$/.exec(line)?.[0] ?? "";
        const syntax = JSON.stringify([prefix, suffix, line.trim() === ""]);
        const range = { from: offset, body, to: end, syntax };
        offset = end + 1;
        return range;
    });
    const maskedLines = protectListSource(note, listLines, protectedBlocks);
    if (!maskedLines) return null;
    const delimiterSource = maskedLines.size
        ? note.slice(0, from) + rows.map((line, index) => maskedLines.has(index) ? " ".repeat(line.length) : line).join("\n")
        : note;
    if (insideDelimitedBlock(delimiterSource, from, to, "%%") || insideDelimitedBlock(delimiterSource, from, to, "$$")) return null;
    // Marked can merge a definition-looking line into adjacent prose, so protect it at the source level.
    if (listLines.some(line => !line.protected && definitionLine.test(note.slice(line.body, line.to)))) return null;
    return { from, to, text, protectedMarkup: [JSON.stringify(markup)], listLines };
}

function listItemAt(note: string, list: Tokens.List, from: number, cursor: number, allowHtml?: (raw: string) => boolean): ParagraphRange | null {
    for (const item of list.items) {
        // Only edit when the item can be mapped back to the editor without lexer normalization.
        if (!note.startsWith(item.raw, from)) return null;
        const itemEnd = from + item.raw.length;
        const line = item.raw.replace(/\n+$/, "");
        const to = from + line.length;
        if (cursor <= to) {
            const prefix = /^ {0,3}(?:[-+*]|\d{1,9}[.)])[ \t]+(?:\[[ xX]\][ \t]+)?/.exec(line)?.[0];
            if (!prefix || line.includes("\n")) return null;
            const content = item.tokens[item.task ? 1 : 0];
            const text = line.slice(prefix.length);
            // Obsidian recognizes custom task states that Marked treats as plain text.
            if (!item.task && /^\[[^\]\n]\](?:[ \t]|$)/u.test(text)) return null;
            if (!text.trim() || definitionLine.test(text) || item.tokens.length !== (item.task ? 2 : 1) || !content
                || (content.type !== "paragraph" && content.type !== "text") || !content.tokens || content.raw !== text
                || insideDelimitedBlock(note, from, to, "%%") || insideDelimitedBlock(note, from, to, "$$")) return null;
            // Marked treats everything after a checkbox as text, including block markers.
            if (item.task && lexer(text, { gfm: true })[0]?.type !== "paragraph") return null;
            const protectedMarkup = paragraphMarkup(content.tokens, text, allowHtml);
            return protectedMarkup && { from: from + prefix.length, to, text, protectedMarkup, listPrefix: prefix };
        }
        if (cursor < itemEnd || hasAmbiguousDelimiter(item)) return null;
        from = itemEnd;
    }
    return null;
}

export function paragraphAt(note: string, cursor: number, allowHtml?: (raw: string) => boolean): ParagraphRange | null {
    // Obsidian's editor positions use LF. Keep lexer offsets tied to the exact editor text.
    if (note.includes("\r")) return null;
    let offset = 0;
    if (/^\uFEFF?---[ \t]*\n/.test(note)) {
        const openingEnd = note.indexOf("\n") + 1;
        const closing = /^(?:---|\.\.\.)[ \t]*(?:\n|$)/m.exec(note.slice(openingEnd));
        if (!closing) return null;
        offset = openingEnd + closing.index + closing[0].length;
    }
    if (cursor < offset) return null;
    for (const token of lexer(note.slice(offset), { gfm: true })) {
        // A lexer that skipped or normalized any prefix cannot authorize a source offset.
        if (!note.startsWith(token.raw, offset)) return null;
        const from = offset;
        offset += token.raw.length;
        if (cursor >= offset && !(cursor === note.length && offset === note.length)) {
            // Delimiters inside code or HTML make the raw Obsidian comment/math boundaries ambiguous.
            if (hasAmbiguousDelimiter(token) || hasUnresolvedQuoteDelimiters(token)) return null;
            continue;
        }
        if (token.type === "list") {
            const list = token as Tokens.List;
            const block = listBlockAt(note, list, from, cursor, allowHtml);
            if (block || hasUnresolvedQuoteDelimiters(list)) return block;
            return listItemAt(note, list, from, cursor, allowHtml);
        }
        if (token.type !== "paragraph") return null;
        const text = token.raw.replace(/\n+$/, "");
        const to = from + text.length;
        if (cursor > to || !text.trim()
            || /^(?: {2,}|\t)|^\s*(?:\[\^[^\]]+\]:|\^[\w-]+\s*$)/m.test(text)) return null;
        if (insideDelimitedBlock(note, from, to, "%%") || insideDelimitedBlock(note, from, to, "$$")) return null;
        const protectedMarkup = paragraphMarkup((token as Tokens.Paragraph).tokens, text, allowHtml);
        return protectedMarkup && { from, to, text, protectedMarkup };
    }
    return null;
}

export function isParagraphReplacement(
    note: string, target: ParagraphRange, replacement: string, allowHtml?: (raw: string) => boolean
): boolean {
    const updated = note.slice(0, target.from) + replacement + note.slice(target.to);
    const parsed = paragraphAt(updated, target.from, allowHtml);
    return !!parsed && parsed.from === target.from && parsed.to === target.from + replacement.length
        && parsed.listPrefix === target.listPrefix
        && (parsed.listLines?.length ?? 0) === (target.listLines?.length ?? 0)
        && (!target.listLines || target.listLines.every((line, index) => line.syntax === parsed.listLines?.[index]?.syntax))
        && parsed.protectedMarkup.length === target.protectedMarkup.length
        && parsed.protectedMarkup.every((value, index) => value === target.protectedMarkup[index]);
}
