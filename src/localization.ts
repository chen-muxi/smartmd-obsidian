import { englishMessages } from "./locale-en";

export type InterfaceLanguage = "en" | "zh-CN";

let language: InterfaceLanguage = "en";
const parameters = /\{(\d+)\}/g;

export function setInterfaceLanguage(value: InterfaceLanguage): void {
    language = value;
}

function format(message: string, args: readonly unknown[]): string {
    return message.replace(parameters, (placeholder, index) =>
        Number(index) < args.length ? String(args[Number(index)]) : placeholder);
}

/** Use placeholders for dynamic values so names, paths and document text remain untouched. */
export function localize(message: string, ...args: unknown[]): string {
    const translation = language === "en" && Object.prototype.hasOwnProperty.call(englishMessages, message)
        ? englishMessages[message]
        : message;
    return format(translation, args);
}

interface MessagePattern {
    expression: RegExp;
    slots: number[];
    translation: string;
}

function escapeRegex(value: string): string {
    return value.replace(/[.*+?^\x24{}()|[\]\\]/g, "\\$&");
}

// Display old persisted status messages and internal errors without changing the stored
// text or any caller that classifies an error before displaying it.
const knownMessagePatterns: MessagePattern[] = Object.entries(englishMessages)
    .filter(([message]) => /\{\d+\}/.test(message))
    .sort(([left], [right]) => right.replace(parameters, "").length - left.replace(parameters, "").length)
    .map(([message, translation]) => {
        const slots = [...message.matchAll(parameters)].map(match => Number(match[1]));
        return {
            expression: new RegExp("^" + message.split(parameters).map((part, index) =>
                index % 2 === 0 ? escapeRegex(part) : "([\\s\\S]*?)").join("") + "$"),
            slots,
            translation,
        };
    });

function translateMessage(message: string): string {
    if (Object.prototype.hasOwnProperty.call(englishMessages, message)) return englishMessages[message];
    if (message.length > 4096 || !/\p{Script=Han}/u.test(message)) return message;
    for (const { expression, slots, translation } of knownMessagePatterns) {
        const match = expression.exec(message);
        if (!match) continue;
        const args: string[] = [];
        slots.forEach((slot, index) => {
            args[slot] = match[index + 1];
        });
        return format(translation, args);
    }
    return message;
}

export function localizeMessage(value: string): string;
export function localizeMessage<T>(value: T): T;
export function localizeMessage(value: unknown): unknown {
    return language === "en" && typeof value === "string" ? translateMessage(value) : value;
}
