import { localize } from "./src/localization";
import { Plugin } from "obsidian";
import { SmartMDRuntime, SmartMDSettingsTab } from "./src/runtime";

export default class SmartMDPlugin extends Plugin {
    private runtime!: SmartMDRuntime;
    private statusBarItem: HTMLElement | null = null;

    async onload(): Promise<void> {
        try {
            this.statusBarItem = this.addStatusBarItem();
            this.statusBarItem.setText(localize("SmartMD: 就绪"));
            this.statusBarItem.addClass("smartmd-status-bar-item");
            this.statusBarItem.title = localize("点击打开 SmartMD 设置");
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
