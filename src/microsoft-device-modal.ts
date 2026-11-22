import { App, Modal, Notice, Platform } from "obsidian";
import { MicrosoftDeviceCodeChallenge } from "./microsoft-device-flow";

/** Ephemeral device-authorization dialog. Nothing is saved to plugin data. */
export class MicrosoftDeviceCodeModal extends Modal {
    private finished = false;
    private explicitlyCancelled = false;

    constructor(app: App, private cancelAuthorization: () => void) {
        super(app);
    }

    onOpen(): void {
        this.titleEl.setText("登录微软账号 · OneDrive");
        this.contentEl.empty();
        this.contentEl.createEl("p", {
            text: "正在向微软请求设备码，请稍候……",
        });
    }

    showChallenge(challenge: MicrosoftDeviceCodeChallenge): void {
        const url = new URL(challenge.verificationUri);
        if (url.protocol !== "https:" ||
            !["www.microsoft.com", "microsoft.com"].includes(url.hostname)) {
            throw new Error("非微软官方验证地址已被阻止。");
        }
        const content = this.contentEl;
        content.empty();
        content.createEl("p", {
            text: "在微软官方网站完成验证。只输入此窗口生成的代码，不要使用别人发来的设备码。",
        });
        content.createEl("p", { text: "第一步：打开微软验证页面" });
        const link = content.createEl("a", {
            text: challenge.verificationUri,
            href: challenge.verificationUri,
        });
        link.setAttribute("target", "_blank");
        link.setAttribute("rel", "noopener noreferrer");
        link.addEventListener("click", (event) => {
            event.preventDefault();
            void this.openVerificationPage(challenge.verificationUri);
        });
        const openButton = content.createEl("button", { text: "打开微软验证页面" });
        openButton.addEventListener("click", () => {
            void this.openVerificationPage(challenge.verificationUri);
        });
        content.createEl("p", { text: "第二步：输入以下设备码并授权" });
        const code = content.createEl("p", { cls: "smartkeep-microsoft-device-user-code" });
        code.createEl("strong", { text: challenge.userCode });
        const copyButton = content.createEl("button", { text: "复制设备码" });
        copyButton.addEventListener("click", () => void this.copyCode(challenge.userCode));
        const minutes = Math.ceil(challenge.expiresIn / 60);
        content.createEl("p", {
            text: "设备码约 " + minutes +
                " 分钟内有效。完成微软网站授权后，Obsidian 会自动继续；关闭此窗口不会中断登录，可在设置面板随时取消。",
        });
        content.createEl("p", {
            text: "权限为 User.Read、Files.ReadWrite.AppFolder，以及用于保持登录的 offline_access。成功后不会自动同步知识库。",
        });
        const cancelButton = content.createEl("button", { text: "取消登录" });
        cancelButton.addEventListener("click", () => this.explicitCancel());
    }

    explicitCancel(): void {
        this.explicitlyCancelled = true;
        this.cancelAuthorization();
        this.close();
    }

    private async openVerificationPage(uri: string): Promise<void> {
        try {
            if (Platform.isDesktopApp && typeof (window as any).require === "function") {
                const shell = (window as any).require("electron")?.shell;
                if (typeof shell?.openExternal === "function") {
                    await shell.openExternal(uri);
                    return;
                }
            }
            window.open(uri, "_blank");
        } catch {
            new Notice("无法自动打开微软网页，请复制上方地址到浏览器。");
        }
    }

    private async copyCode(code: string): Promise<void> {
        try {
            if (navigator.clipboard?.writeText) {
                await navigator.clipboard.writeText(code);
            } else if (Platform.isDesktopApp &&
                typeof (window as any).require === "function") {
                (window as any).require("electron").clipboard.writeText(code);
            } else {
                throw new Error("clipboard unavailable");
            }
            new Notice("设备码已复制。只在微软官方验证页面使用。");
        } catch {
            new Notice("无法复制设备码，请手动输入页面显示的代码。");
        }
    }

    dismiss(): void {
        this.finished = true;
        this.close();
    }

    onClose(): void {
        this.contentEl.empty();
    }
}
