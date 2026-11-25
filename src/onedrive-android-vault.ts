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

/** Verify Android's shared-storage adapter and WebView agree on the Vault root. */
export function getAndroidVaultRoot(adapter: DataAdapter): string {
    const appRoot = getAndroidAppStorageRoot(adapter);
    if (appRoot) return appRoot;
    try {
        const native = adapter as Partial<AndroidPathAdapter>;
        if (typeof native.getFullPath !== "function") return "";
        const resource = new URL(adapter.getResourcePath(""));
        if (resource.origin !== "http://localhost" || resource.username || resource.password ||
            resource.search || resource.hash || !resource.pathname.startsWith("/_capacitor_file_/")) return "";
        const root = decodeURIComponent(resource.pathname.slice("/_capacitor_file_".length)).replace(/\/+$/, "");
        const relative = /^\/storage\/(?:emulated\/[0-9]+|[a-f0-9]{4}-[a-f0-9]{4})\/(.+)$/i.exec(root)?.[1];
        if (!relative || /^(?:Android|\.)(?:\/|$)/i.test(relative) ||
            relative.split("/").some(part => !part || part === "." || part === ".." || /[\\:\u0000-\u001f]/.test(part))) return "";
        return native.getFullPath("").replace(/\/+$/, "") === root ? root : "";
    } catch {
        return "";
    }
}
