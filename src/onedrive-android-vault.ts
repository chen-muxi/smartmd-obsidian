import type { DataAdapter } from "obsidian";

type AndroidPathAdapter = Pick<DataAdapter, "getResourcePath"> & {
    getFullPath(path: string): string;
};

/** An App storage Vault has a private, stable directory tied to this Obsidian install. */
export function getAndroidAppStorageRoot(adapter: DataAdapter): string {
    const native = adapter as Partial<AndroidPathAdapter>;
    if (typeof native.getFullPath !== "function") return "";
    try {
        const relativeRoot = native.getFullPath("").replace(/\/+$/, "");
        if (!relativeRoot || relativeRoot === "." || relativeRoot === ".." ||
            /[/\\\u0000-\u001f]/.test(relativeRoot)) return "";
        const resource = new URL(adapter.getResourcePath(""));
        if (resource.origin !== "http://localhost" || resource.username || resource.password ||
            resource.search || resource.hash) return "";
        const physicalPath = decodeURIComponent(resource.pathname).replace(/\/+$/, "");
        const prefix = /^\/_capacitor_file_\/storage\/emulated\/[0-9]+\/Android\/data\/md\.obsidian\/files\//;
        if (!prefix.test(physicalPath) || physicalPath.replace(prefix, "") !== relativeRoot) return "";
        return physicalPath.slice("/_capacitor_file_".length);
    } catch {
        return "";
    }
}
