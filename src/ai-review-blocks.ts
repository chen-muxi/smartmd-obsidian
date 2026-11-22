import { lexer } from "marked";

export interface ReviewChange {
    from: number;
    to: number;
    insert: string;
}

interface Block {
    from: number;
    to: number;
    text: string;
}

function blocks(text: string): Block[] {
    if (!text) return [];
    const results: Block[] = [];
    let offset = 0;
    const yaml = /^\uFEFF?---[ \t]*\n[\s\S]*?\n(?:---|\.\.\.)[ \t]*(?:\n|$)/.exec(text)?.[0];
    if (yaml) {
        offset = yaml.length;
        results.push({ from: 0, to: offset, text: yaml });
    }
    for (const token of lexer(text.slice(offset), { gfm: true })) {
        if (!text.startsWith(token.raw, offset)) break;
        const from = offset;
        offset += token.raw.length;
        if (token.type === "space" && results.length) {
            const preceding = results[results.length - 1];
            preceding.to = offset;
            preceding.text = text.slice(preceding.from, offset);
        } else {
            results.push({ from, to: offset, text: token.raw });
        }
    }
    if (offset < text.length) results.push({ from: offset, to: text.length, text: text.slice(offset) });
    return results;
}

// Unique blocks give large documents stable anchors without a quadratic diff matrix.
function uniqueAnchors(before: Block[], after: Block[], oldStart: number, oldEnd: number, newStart: number, newEnd: number) {
    const positions = (items: Block[], from: number, to: number) => {
        const found = new Map<string, number>();
        for (let i = from; i < to; i++) found.set(items[i].text, found.has(items[i].text) ? -1 : i);
        return found;
    };
    const oldPositions = positions(before, oldStart, oldEnd);
    const newPositions = positions(after, newStart, newEnd);
    type Anchor = { old: number; next: number; previous?: Anchor };
    const ends: Anchor[] = [];
    for (const [text, old] of oldPositions) {
        const next = newPositions.get(text) ?? -1;
        if (old < 0 || next < 0) continue;
        let from = 0;
        let to = ends.length;
        while (from < to) {
            const middle = Math.floor((from + to) / 2);
            if (ends[middle].next < next) from = middle + 1;
            else to = middle;
        }
        ends[from] = { old, next, previous: ends[from - 1] };
    }
    const result: Anchor[] = [];
    for (let anchor: Anchor | undefined = ends[ends.length - 1]; anchor; anchor = anchor.previous) result.push(anchor);
    return result.reverse();
}

function repeatedAnchors(before: Block[], after: Block[], oldStart: number, oldEnd: number, newStart: number, newEnd: number) {
    const match = (left: Block[], right: Block[], start: number, end: number, lower: number, upper: number) => {
        const positions = new Map<string, number[]>();
        for (let index = lower; index < upper; index++) {
            const matches = positions.get(right[index].text);
            if (matches) matches.push(index);
            else positions.set(right[index].text, [index]);
        }
        const anchors: [number, number][] = [];
        for (let index = start; index < end; index++) {
            const matches = positions.get(left[index].text);
            if (!matches) continue;
            let from = 0;
            let to = matches.length;
            while (from < to) {
                const middle = Math.floor((from + to) / 2);
                if (matches[middle] < lower) from = middle + 1;
                else to = middle;
            }
            if (from < matches.length) {
                anchors.push([index, matches[from]]);
                lower = matches[from] + 1;
            }
        }
        return anchors;
    };
    const forward = match(before, after, oldStart, oldEnd, newStart, newEnd);
    const backward = match(after, before, newStart, newEnd, oldStart, oldEnd);
    return forward.length >= backward.length ? forward : backward.map(([next, old]) => [old, next]);
}

export function documentReviewChanges(original: string, revised: string): ReviewChange[] {
    const before = blocks(original);
    const after = blocks(revised);
    const changes: ReviewChange[] = [];
    let oldIndex = 0;
    let newIndex = 0;

    const replaceGap = (oldEnd: number, newEnd: number) => {
        while (oldIndex < oldEnd || newIndex < newEnd) {
            const old = oldIndex < oldEnd ? before[oldIndex++] : undefined;
            const replacement = newIndex < newEnd ? after[newIndex++].text : "";
            const from = old?.from ?? before[oldEnd]?.from ?? original.length;
            if ((old?.text ?? "") !== replacement) changes.push({ from, to: old?.to ?? from, insert: replacement });
        }
    };

    const oldCount = before.length;
    const newCount = after.length;
    while (oldIndex < oldCount && newIndex < newCount && before[oldIndex].text === after[newIndex].text) {
        oldIndex++;
        newIndex++;
    }
    let oldEnd = oldCount;
    let newEnd = newCount;
    while (oldEnd > oldIndex && newEnd > newIndex && before[oldEnd - 1].text === after[newEnd - 1].text) {
        oldEnd--;
        newEnd--;
    }
    const width = newEnd - newIndex + 1;
    const height = oldEnd - oldIndex + 1;
    if (width * height <= 4_000_000) {
        const baseOld = oldIndex;
        const baseNew = newIndex;
        const lengths = new Uint32Array(width * height);
        for (let i = height - 2; i >= 0; i--) {
            for (let j = width - 2; j >= 0; j--) {
                const index = i * width + j;
                lengths[index] = before[baseOld + i].text === after[baseNew + j].text
                    ? lengths[index + width + 1] + 1
                    : Math.max(lengths[index + width], lengths[index + 1]);
            }
        }
        let i = 0;
        let j = 0;
        while (i < height - 1 && j < width - 1) {
            if (before[baseOld + i].text === after[baseNew + j].text) {
                replaceGap(baseOld + i, baseNew + j);
                oldIndex++;
                newIndex++;
                i++;
                j++;
            } else if (lengths[(i + 1) * width + j] >= lengths[i * width + j + 1]) {
                i++;
            } else {
                j++;
            }
        }
    } else {
        const anchors = uniqueAnchors(before, after, oldIndex, oldEnd, newIndex, newEnd);
        for (const anchor of [...anchors, { old: oldEnd, next: newEnd }]) {
            for (const [old, next] of repeatedAnchors(before, after, oldIndex, anchor.old, newIndex, anchor.next)) {
                replaceGap(old, next);
                oldIndex++;
                newIndex++;
            }
            replaceGap(anchor.old, anchor.next);
            if (anchor.old < oldEnd) {
                oldIndex++;
                newIndex++;
            }
        }
    }
    replaceGap(oldEnd, newEnd);
    return changes;
}

export function applyReviewChanges(note: string, changes: ReviewChange[]): string {
    let from = 0;
    let result = "";
    for (const change of changes) {
        result += note.slice(from, change.from) + change.insert;
        from = change.to;
    }
    return result + note.slice(from);
}

export function reverseReviewChanges(note: string, changes: ReviewChange[]): ReviewChange[] {
    let displacement = 0;
    return changes.map(change => {
        const from = change.from + displacement;
        displacement += change.insert.length - (change.to - change.from);
        return { from, to: from + change.insert.length, insert: note.slice(change.from, change.to) };
    });
}
