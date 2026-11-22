import { Plugin } from "obsidian";
import { SmartMDRuntime, SmartMDSettingsTab, SMARTMD_BUILD_VERSION, SMARTMD_BUILD_FINGERPRINT } from "./src/runtime";

export default class SmartMDPlugin extends Plugin {
    private runtime!: SmartMDRuntime;
    private statusBarItem: HTMLElement | null = null;

    async onload(): Promise<void> {
        console.log(`%c[SmartMD] 插件已加载 | 版本: ${SMARTMD_BUILD_VERSION} | 编译指纹: ${SMARTMD_BUILD_FINGERPRINT}`, "color: #10b981; font-weight: bold; font-size: 13px;");
        try {
            this.statusBarItem = this.addStatusBarItem();
            this.statusBarItem.setText("SmartMD: 就绪");
            this.statusBarItem.setCssStyles({ cursor: "pointer" });
            this.statusBarItem.title = "点击打开 SmartMD 设置";
            this.statusBarItem.onclick = () => {
                this.runtime?.openSettingsTab("onedrive");
            };
        } catch {}
        this.runtime = new SmartMDRuntime(this.app, this, this.statusBarItem);
        await this.runtime.load();
        const settingsTab = new SmartMDSettingsTab(this.app, this, this.runtime);
        this.runtime.setSettingsTab(settingsTab);
        this.addSettingTab(settingsTab);
    }

    updateStatusBar(text: string): void {
        this.runtime?.updateStatusBar(text);
    }

    onunload(): void {
        this.runtime?.unload();
    }
}
