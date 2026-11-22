/**
 * OneDrive user-visible diagnostics must never echo raw external Error, URL,
 * stack, credentials or private absolute file paths. No real tokens are needed
 * to verify this boundary; tests use obviously synthetic placeholders only.
 */
export function redactOneDriveText(value: unknown): string {
    let text = String(value ?? "");
    // Quoted JSON credentials need a separate pass: the legacy key=value
    // filter does not cross the closing quote in {"token":"value"}.
    // Accept either ordinary JSON quotes or escaped JSON within an error
    // message, and conservatively mask the entire value through its closing
    // quote. Never parse or log a real token to produce a diagnostic.
    const quotedKey = /((?:\\?["'])(?:access_token|refresh_token|id_token|client_secret|api[_-]?key|authorization|password|passphrase|token|secret|code|signature|sig)(?:\\?["'])\s*:\s*)(\\?["'])?/gi;
    let searchFrom = 0;
    for (;;) {
        quotedKey.lastIndex = searchFrom;
        const key = quotedKey.exec(text);
        if (!key) break;
        const quote = key[2]?.endsWith("'") ? "'" : '"';
        const escapedQuote = key[2]?.startsWith("\\") || false;
        const begin = quotedKey.lastIndex;
        let end = begin, closeLength = 0;
        if (!key[2]) {
            // Numeric/boolean/unquoted JSON secrets must not bypass redaction.
            while (end < text.length && !/[,}\]\s]/.test(text[end])) end++;
        } else {
            while (end < text.length) {
                const matchEnd = escapedQuote
                    ? text[end] === "\\" && text[end + 1] === quote
                    : text[end] === quote;
                if (matchEnd) {
                    // A quote preceded by an odd number of unescaped backslashes
                    // is part of the JSON value, NOT the terminal quote.
                    let backslashes = 0;
                    let cursor = end - 1;
                    while (cursor >= begin && text[cursor] === "\\") {
                        backslashes++;
                        cursor--;
                    }
                    const isEscapedContent = !escapedQuote && backslashes % 2 === 1;
                    // In a once-stringified JSON message, a closing escaped
                    // quote has 1 (mod 4) consecutive backslashes, including
                    // this slash. An escaped quote *inside* the secret has
                    // 3 (mod 4); parity alone incorrectly exposes its suffix.
                    const slashIsEscaped = escapedQuote && backslashes % 4 !== 0;
                    if (!isEscapedContent && !slashIsEscaped) {
                        const len = escapedQuote ? 2 : 1;
                        const tail = text.slice(end + len);
                        if (/^\s*(?:[,}\]]|$)/.test(tail)) {
                            closeLength = len;
                            break;
                        }
                    }
                }
                end++;
            }
        }
        if (!closeLength) {
            // Malformed JSON: fail closed for the remaining field, avoiding a
            // partial replacement that could expose a private credential.
            end = text.length;
        }
        text = text.slice(0, begin) + "***" + text.slice(end);
        searchFrom = begin + 3 + closeLength;
    }
    // Signed/credential URL query values and plain key=value.
    text = text.replace(/\b(Bearer|Basic)\s+[^\s"'<>;,]+/gi, "$1 ***");
    text = text.replace(
        /((?:\b(?:access_token|refresh_token|id_token|client_secret|api[_-]?key|authorization|password|passphrase|token|secret|code|signature|sig)\b)\s*[:=]\s*["']?)[^\s&"'<>;,}]+/gi,
        "$1***",
    );
    text = text.replace(/(https?:\/\/)[^\s/@]+:[^\s/@]+@/gi, "$1***@");
    // Remove absolute Windows drive and UNC paths, even when embedded in logs.
    text = text.replace(/\b[A-Za-z]:[\\/]+(?:[^\s\\/"'<>|()]+[\\/]+)*[^\s\\/"'<>|()]*/g, "[private path]");
    text = text.replace(/\\\\[^\s\\/]+[\\/]+[^\s"'<>|()]+/g, "[private path]");
    text = text.replace(/\/(?:Users|home|private|mnt|var|tmp)\/(?:[^\s"'<>|()]+\/?)+/gi, "[private path]");
    return text.slice(0, 2400);
}

export function safeOneDriveError(error: unknown): string {
    const messages: string[] = [];
    const seen = new Set<unknown>();
    let current: unknown = error;
    for (let depth = 0; depth < 3 && current != null && !seen.has(current); depth++) {
        seen.add(current);
        if (typeof current === "string") {
            messages.push(current);
            break;
        }
        if (typeof current !== "object") {
            messages.push(String(current));
            break;
        }
        const value = current as { message?: unknown; cause?: unknown; error?: unknown };
        if (typeof value.message === "string" && value.message) messages.push(value.message);
        // Stack traces are deliberately not output: they can include private
        // paths or credentials not recognizable as a conventional key=value.
        current = value.cause ?? value.error;
    }
    return redactOneDriveText(messages.join(" | ") || "未知错误");
}
