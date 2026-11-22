import type { TFile, Vault } from "obsidian";
import { getAndroidAppStorageRoot } from "./onedrive-android-vault";

function safeRelative(path: string): boolean {
    return Boolean(path) && path.split("/").every(
        part => part && part !== "." && part !== ".." && !/[\\:\u0000-\u001f]/.test(part),
    );
}

/** Compare inside Obsidian's guarded read/modify operation. The caller owns recovery. */
export async function replaceAndroidMarkdownIfUnchanged(
    vault: Pick<Vault, "process" | "getAbstractFileByPath">,
    file: TFile,
    path: string,
    expected: ArrayBuffer,
    replacement: ArrayBuffer,
    assertCurrent: () => void,
): Promise<void> {
    if (file.extension.toLowerCase() !== "md") {
        throw new Error("安卓端只允许使用经过验证的 Markdown 原位更新。");
    }
    const decoder = new TextDecoder("utf-8", { fatal: true, ignoreBOM: true });
    const originalText = decoder.decode(expected);
    const newText = decoder.decode(replacement);
    await vault.process(file, (currentText) => {
        assertCurrent();
        if (file.path !== path || vault.getAbstractFileByPath(path) !== file) {
            throw new Error("安卓端写入目标已移动或更换，保留本机内容。");
        }
        if (currentText !== originalText) {
            throw new Error("安卓端写入前原文已变化，保留本机内容。");
        }
        return newText;
    });
}

/** On Android App storage, publish a verified temporary file with the host's no-replace move. */
export async function placeAndroidStagedFile(
    vault: Pick<Vault, "adapter">,
    stage: string,
    destination: string,
    assertCurrent: () => void,
): Promise<void> {
    if (!stage.startsWith(".suite-staging/onedrive/") || destination.startsWith(".suite-staging/")) {
        throw new Error("安卓端暂存文件或目标位置不可信。");
    }
    await moveAndroidFile(vault, stage, destination, assertCurrent);
}

async function moveAndroidFile(
    vault: Pick<Vault, "adapter">,
    source: string,
    destination: string,
    assertCurrent: () => void,
): Promise<void> {
    const adapter = vault.adapter;
    const root = getAndroidAppStorageRoot(adapter);
    if (!root || !safeRelative(source) || !safeRelative(destination)) {
        throw new Error("安卓端移动来源或目标位置不可信。");
    }
    if (await adapter.exists(destination)) {
        throw new Error("安卓端目标已存在，保留两份原件。");
    }
    assertCurrent();
    if (getAndroidAppStorageRoot(adapter) !== root) {
        throw new Error("安卓 Vault 物理范围已变化，拒绝落盘。");
    }
    // The Android adapter is verified on-device to reject an existing target.
    // An error stays an error and never falls back to a write that could replace it.
    await adapter.rename(source, destination);
}

export async function trashAndroidFile(
    vault: Pick<Vault, "adapter" | "getAbstractFileByPath">,
    file: TFile,
    path: string,
    trashPath: string,
    expected: ArrayBuffer,
    assertCurrent: () => void,
    assertSource: () => void,
): Promise<void> {
    const adapter = vault.adapter;
    const root = getAndroidAppStorageRoot(adapter);
    if (!root || !safeRelative(path) || !safeRelative(trashPath) ||
        !trashPath.startsWith(".trash/smartmd-onedrive/") || path.startsWith(".")) {
        throw new Error("安卓回收站路径或源文件不可信。");
    }
    const assertRoot = () => {
        if (vault.adapter !== adapter || getAndroidAppStorageRoot(adapter) !== root) {
            throw new Error("安卓 Vault 物理范围已变化。");
        }
    };
    const matches = (actual: ArrayBuffer) => {
        const approved = new Uint8Array(expected);
        return actual.byteLength === approved.length &&
            new Uint8Array(actual).every((byte, index) => byte === approved[index]);
    };
    if (await adapter.exists(trashPath)) throw new Error("安卓回收位置已占用。");
    const before = await adapter.readBinary(path);
    assertRoot();
    assertCurrent();
    assertSource();
    if (!matches(before) || file.path !== path || vault.getAbstractFileByPath(path) !== file) {
        throw new Error("安卓端待删除的笔记已修改、移动或更换。");
    }
    try {
        await adapter.rename(path, trashPath);
        const retained = await adapter.readBinary(trashPath);
        if (!matches(retained) || await adapter.exists(path)) throw new Error("安卓原件在移动期间发生变化。");
        assertRoot();
        assertCurrent();
    } catch (error) {
        // A late edit or revocation may arrive between the final read and
        // the native move. Return those exact bytes without replacing a
        // newly created source. The persisted recovery record pins trashPath.
        try {
            assertRoot();
            if (await adapter.exists(trashPath)) {
                if (await adapter.exists(path)) throw new Error("原路径已重新出现");
                assertRoot();
                await adapter.rename(trashPath, path);
            }
        } catch {
            throw new Error("安卓删除结果须核对，原稿或新编辑保留在 " + trashPath);
        }
        throw error;
    }
}
