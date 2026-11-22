import {
    App,
    Modal,
    Notice,
    normalizePath,
    Platform,
    TFile,
    TFolder,
} from "obsidian";
import { SmartMDStore } from "./core";
import { OperationAuthority, RunTicket } from "./operation-authority";
import { ExternalFilterConfig, FilterRuleItem } from "./types";
import { FilterMatcher } from "./filter-matcher";

declare const require: any;

type ScannedFile = {
    relativePath: string;
    absolutePath: string;
    content: string;
    hash: string;
    mtimeMs?: number;
    size?: number;
};

type BindingPreview = {
    markdownCount: number;
    targetPath: string;
    conflicts: string[];
    filterSummary?: string;
};

function uid(prefix: string): string {
    return `${prefix}_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 9)}`;
}

async function digest(content: string): Promise<string> {
    const bytes = new TextEncoder().encode(content);
    const result = await crypto.subtle.digest("SHA-256", bytes);
    return Array.from(new Uint8Array(result)).map((value) => value.toString(16).padStart(2, "0")).join("");
}

function safeName(value: string): string {
    return value.replace(/[\\/:*?"<>|]/g, "-").trim() || "外部文件夹";
}

function inside(parent: string, child: string, pathApi: any): boolean {
    const relative = pathApi.relative(pathApi.resolve(parent), pathApi.resolve(child));
    return relative === "" || (!relative.startsWith("..") && !pathApi.isAbsolute(relative));
}

class ExternalBindingModal extends Modal {
    private folderPath = "";
    private nickname = "";

    constructor(app: App, private service: ExternalSyncService, private onAdded?: () => void) {
        super(app);
    }

    onOpen(): void {
        this.modalEl.addClass("smartmd-external-modal");
        this.contentEl.addClass("smartmd-external-add-modal");
        this.titleEl.setText("添加外部同步文件夹");
        this.contentEl.createEl("p", {
            text: "插件会扫描所选文件夹及子目录中的全部 Markdown，保留原目录结构并在知识库内建立副本。知识库副本与外部原文档的修改会实时双向同步；图片、代码和附件不会被同步或改动。",
        });

        const pathField = this.contentEl.createDiv({ cls: "smartmd-external-field" });
        pathField.createEl("label", { text: "外部文件夹路径" });
        const pathRow = pathField.createDiv({ cls: "smartmd-external-path-row" });
        const input = pathRow.createEl("input", { type: "text", cls: "smartmd-external-path-input" });
        input.placeholder = "外部文件夹完整路径";
        input.oninput = () => this.folderPath = input.value.trim();
        pathRow.createEl("button", { text: "浏览…" }).onclick = async () => {
            const selected = await this.service.chooseFolder();
            if (selected) { this.folderPath = selected; input.value = selected; }
        };

        const nameField = this.contentEl.createDiv({ cls: "smartmd-external-field" });
        nameField.createEl("label", { text: "知识库内名称（可选）" });
        const nickname = nameField.createEl("input", { type: "text", cls: "smartmd-external-name-input" });
        nickname.placeholder = "默认使用外部文件夹名称";
        nickname.oninput = () => this.nickname = nickname.value.trim();
        nameField.createEl("p", {
            cls: "setting-item-description smartmd-external-help",
            text: "只用于知识库内显示和区分同名文件夹，不会修改外部文件夹名称。",
        });

        // 筛选规则设置折叠区
        const filterDetails = this.contentEl.createEl("details", { cls: "smartmd-external-filter-details" });
        filterDetails.style.marginTop = "16px";
        filterDetails.style.marginBottom = "16px";
        filterDetails.style.border = "1px solid var(--background-modifier-border)";
        filterDetails.style.borderRadius = "var(--radius-m)";
        filterDetails.style.padding = "10px 14px";
        filterDetails.style.backgroundColor = "var(--background-secondary)";

        const filterSummary = filterDetails.createEl("summary", { text: "⚙️ 筛选规则设置（默认已自动排除 node_modules 等构建依赖）" });
        filterSummary.style.cursor = "pointer";
        filterSummary.style.fontWeight = "bold";

        const filterContent = filterDetails.createDiv();
        filterContent.style.marginTop = "12px";

        // 模式选择
        const modeRow = filterContent.createDiv({ cls: "smartmd-external-field" });
        modeRow.createEl("label", { text: "筛选模式" });
        const modeSelect = modeRow.createEl("select");
        modeSelect.createEl("option", { text: "排除模式 (黑名单，推荐) - 排除指定目录与文件", value: "exclude" });
        modeSelect.createEl("option", { text: "仅包含模式 (白名单) - 仅同步指定目录与文件", value: "include" });

        // 开箱即用推荐项
        const presetBuildRow = filterContent.createDiv();
        presetBuildRow.style.margin = "6px 0";
        const presetBuildCheck = presetBuildRow.createEl("input", { type: "checkbox" });
        presetBuildCheck.checked = true;
        presetBuildCheck.id = "preset-build-new";
        const presetBuildLabel = presetBuildRow.createEl("label", { text: " 自动排除常用构建缓存 (node_modules, dist, build, target 等)" });
        presetBuildLabel.htmlFor = "preset-build-new";

        const presetIDERow = filterContent.createDiv();
        presetIDERow.style.margin = "6px 0";
        const presetIDECheck = presetIDERow.createEl("input", { type: "checkbox" });
        presetIDECheck.checked = true;
        presetIDECheck.id = "preset-ide-new";
        const presetIDELabel = presetIDERow.createEl("label", { text: " 自动排除 IDE 与系统隐藏配置 (.vscode, .idea, .git, .obsidian 等)" });
        presetIDELabel.htmlFor = "preset-ide-new";

        // 自定义规则文本框
        const rulesField = filterContent.createDiv({ cls: "smartmd-external-field" });
        rulesField.style.marginTop = "10px";
        rulesField.createEl("label", { text: "自定义规则（每行一条通配符或精确名称，例如 archive/** 或 *draft*）" });
        const rulesTextarea = rulesField.createEl("textarea");
        rulesTextarea.rows = 3;
        rulesTextarea.placeholder = "例如：\narchive/**\n*draft*\nREADME.md";
        rulesTextarea.style.width = "100%";
        rulesTextarea.style.boxSizing = "border-box";

        const actions = this.contentEl.createDiv({ cls: "smartmd-external-actions" });
        actions.createEl("button", { text: "取消" }).onclick = () => this.close();
        const preview = actions.createEl("button", { text: "预览并添加", cls: "mod-cta" });
        preview.onclick = async () => {
            preview.disabled = true;
            try {
                const userRules: FilterRuleItem[] = [];
                const lines = rulesTextarea.value.split("\n").map((l) => l.trim()).filter(Boolean);
                for (const line of lines) {
                    const isWildcard = line.includes("*") || line.includes("?");
                    const isFolder = line.endsWith("/") || line.includes("/**");
                    userRules.push({
                        id: uid("rule"),
                        pattern: line,
                        target: isFolder ? "folder" : "all",
                        syntax: isWildcard ? "glob" : "exact",
                        enabled: true,
                    });
                }
                const filterConfig: ExternalFilterConfig = {
                    mode: modeSelect.value as any,
                    builtinPresets: {
                        ignoreBuildArtifacts: presetBuildCheck.checked,
                        ignoreIDEConfig: presetIDECheck.checked,
                    },
                    userRules,
                    filterHash: "",
                };
                filterConfig.filterHash = FilterMatcher.computeFilterHash(filterConfig);

                await this.service.previewAndAdd(this.folderPath, this.nickname, filterConfig);
                this.close();
                this.onAdded?.();
            } catch (error: any) {
                new Notice(`无法添加：${String(error.message || error)}`, 7000);
            } finally {
                preview.disabled = false;
            }
        };
    }

    onClose(): void {
        this.contentEl.empty();
    }
}

class BindingPreviewModal extends Modal {
    constructor(
        app: App,
        private preview: BindingPreview,
        private confirmAdd: () => Promise<void>,
    ) {
        super(app);
    }

    onOpen(): void {
        this.titleEl.setText("首次绑定预览");
        this.contentEl.createEl("p", { text: `发现 ${this.preview.markdownCount} 篇 Markdown。` });
        if (this.preview.filterSummary) {
            this.contentEl.createEl("p", { text: `筛选设定：${this.preview.filterSummary}`, cls: "setting-item-description" });
        }
        this.contentEl.createEl("p", { text: `知识库位置：${this.preview.targetPath}` });
        if (this.preview.conflicts.length) {
            this.contentEl.createEl("p", { text: `有 ${this.preview.conflicts.length} 个潜在冲突。绑定后会暂停这些文档，不会自动覆盖。` });
            for (const path of this.preview.conflicts.slice(0, 10)) this.contentEl.createEl("div", { text: path });
        }
        this.contentEl.createEl("p", { text: "首次绑定只从外部导入，不会删除或覆盖任何已有内容。" });
        const actions = this.contentEl.createDiv({ cls: "smartmd-actions" });
        actions.createEl("button", { text: "取消" }).onclick = () => this.close();
        const confirm = actions.createEl("button", { text: "确认绑定", cls: "mod-cta" });
        confirm.onclick = async () => {
            confirm.disabled = true;
            try { await this.confirmAdd(); this.close(); } finally { confirm.disabled = false; }
        };
    }

    onClose(): void {
        this.contentEl.empty();
    }
}

export class ExternalFilterModal extends Modal {
    private filterConfig: ExternalFilterConfig;

    constructor(
        app: App,
        private service: ExternalSyncService,
        private binding: any,
        private onSaved?: () => void,
    ) {
        super(app);
        const existing = binding.filterConfig || {
            mode: "exclude",
            builtinPresets: { ignoreBuildArtifacts: true, ignoreIDEConfig: true },
            userRules: [],
            filterHash: "",
        };
        this.filterConfig = JSON.parse(JSON.stringify(existing));
        if (!this.filterConfig.builtinPresets) {
            this.filterConfig.builtinPresets = { ignoreBuildArtifacts: true, ignoreIDEConfig: true };
        }
        if (!Array.isArray(this.filterConfig.userRules)) {
            this.filterConfig.userRules = [];
        }
    }

    onOpen(): void {
        this.modalEl.addClass("smartmd-external-modal");
        this.render();
    }

    onClose(): void {
        this.contentEl.empty();
    }

    private render(): void {
        this.contentEl.empty();
        this.titleEl.setText(`【${this.binding.nickname}】筛选与排除规则`);

        const intro = this.contentEl.createEl("p", {
            text: "配置此文件夹同步时的过滤规则。修改规则后系统会自动重校准同步基线；受排除的已同步镜像会移入 Obsidian 回收站；外部源目录绝不删除，未同步创作必须先冷备成功。",
        });
        intro.style.color = "var(--text-muted)";
        intro.style.fontSize = "var(--font-ui-small)";

        // 1. 模式选择
        const modeField = this.contentEl.createDiv({ cls: "smartmd-external-field" });
        modeField.createEl("label", { text: "筛选模式" });
        const modeSelect = modeField.createEl("select");
        modeSelect.createEl("option", { text: "排除模式 (黑名单，推荐) - 排除指定目录与文件", value: "exclude" });
        modeSelect.createEl("option", { text: "仅包含模式 (白名单) - 仅同步指定目录与文件", value: "include" });
        modeSelect.value = this.filterConfig.mode || "exclude";
        modeSelect.onchange = () => {
            this.filterConfig.mode = modeSelect.value as any;
        };

        // 2. 开箱即用推荐预设
        const presetsBox = this.contentEl.createDiv({ cls: "smartmd-card" });
        presetsBox.createEl("h4", { text: "开箱即用推荐预设", attr: { style: "margin: 0 0 8px;" } });

        const presetBuildRow = presetsBox.createDiv();
        presetBuildRow.style.margin = "6px 0";
        const buildCheck = presetBuildRow.createEl("input", { type: "checkbox" });
        buildCheck.id = "modal-preset-build";
        buildCheck.checked = this.filterConfig.builtinPresets.ignoreBuildArtifacts !== false;
        buildCheck.onchange = () => {
            this.filterConfig.builtinPresets.ignoreBuildArtifacts = buildCheck.checked;
        };
        const buildLabel = presetBuildRow.createEl("label", { text: " 自动排除常用构建依赖与缓存 (node_modules, dist, build, target 等)" });
        buildLabel.htmlFor = "modal-preset-build";

        const presetIDERow = presetsBox.createDiv();
        presetIDERow.style.margin = "6px 0";
        const ideCheck = presetIDERow.createEl("input", { type: "checkbox" });
        ideCheck.id = "modal-preset-ide";
        ideCheck.checked = this.filterConfig.builtinPresets.ignoreIDEConfig !== false;
        ideCheck.onchange = () => {
            this.filterConfig.builtinPresets.ignoreIDEConfig = ideCheck.checked;
        };
        const ideLabel = presetIDERow.createEl("label", { text: " 自动排除 IDE 与系统隐藏配置 (.vscode, .idea, .git, .obsidian 等)" });
        ideLabel.htmlFor = "modal-preset-ide";

        // 3. 用户自定义规则列表
        const rulesHeader = this.contentEl.createDiv();
        rulesHeader.style.display = "flex";
        rulesHeader.style.justifyContent = "space-between";
        rulesHeader.style.alignItems = "center";
        rulesHeader.style.marginTop = "16px";
        rulesHeader.style.marginBottom = "8px";
        rulesHeader.createEl("label", { text: `自定义规则 (${this.filterConfig.userRules.length} 条)`, attr: { style: "font-weight: bold;" } });

        const ruleListContainer = this.contentEl.createDiv();
        ruleListContainer.style.maxHeight = "220px";
        ruleListContainer.style.overflowY = "auto";
        ruleListContainer.style.border = "1px solid var(--background-modifier-border)";
        ruleListContainer.style.borderRadius = "var(--radius-m)";
        ruleListContainer.style.padding = "8px";
        ruleListContainer.style.backgroundColor = "var(--background-secondary)";

        if (this.filterConfig.userRules.length === 0) {
            ruleListContainer.createEl("p", {
                text: "暂无自定义规则。下方可快速添加通配符或精确匹配规则。",
                attr: { style: "margin: 6px; color: var(--text-muted); font-size: 13px;" },
            });
        } else {
            for (let i = 0; i < this.filterConfig.userRules.length; i++) {
                const rule = this.filterConfig.userRules[i];
                const row = ruleListContainer.createDiv();
                row.style.display = "flex";
                row.style.alignItems = "center";
                row.style.gap = "8px";
                row.style.padding = "4px 8px";
                row.style.borderBottom = i < this.filterConfig.userRules.length - 1 ? "1px solid var(--background-modifier-border)" : "none";

                // 类型徽章
                const syntaxBadge = row.createEl("span", {
                    text: rule.syntax === "exact" ? "精确" : rule.syntax === "regex" ? "正则" : "通配符",
                });
                syntaxBadge.style.fontSize = "11px";
                syntaxBadge.style.padding = "2px 6px";
                syntaxBadge.style.borderRadius = "4px";
                syntaxBadge.style.backgroundColor = rule.syntax === "exact" ? "var(--interactive-accent)" : "var(--background-modifier-border)";
                syntaxBadge.style.color = rule.syntax === "exact" ? "var(--text-on-accent)" : "var(--text-normal)";

                // 维度徽章
                const targetBadge = row.createEl("span", {
                    text: rule.target === "folder" ? "文件夹" : rule.target === "file" ? "文件" : "全部",
                });
                targetBadge.style.fontSize = "11px";
                targetBadge.style.padding = "2px 6px";
                targetBadge.style.borderRadius = "4px";
                targetBadge.style.backgroundColor = "var(--background-primary)";
                targetBadge.style.color = "var(--text-muted)";

                // 模式内容
                const patText = row.createEl("span", { text: rule.pattern });
                patText.style.flex = "1";
                patText.style.fontFamily = "var(--font-monospace)";
                patText.style.fontSize = "13px";

                // 删除按钮
                const delBtn = row.createEl("button", { text: "删除" });
                delBtn.style.padding = "2px 8px";
                delBtn.onclick = () => {
                    this.filterConfig.userRules.splice(i, 1);
                    this.render();
                };
            }
        }

        // 4. 添加新规则输入栏
        const addCard = this.contentEl.createDiv({ cls: "smartmd-card" });
        addCard.style.marginTop = "12px";
        addCard.createEl("h4", { text: "添加规则", attr: { style: "margin: 0 0 8px;" } });

        const addRow = addCard.createDiv();
        addRow.style.display = "grid";
        addRow.style.gridTemplateColumns = "1fr auto auto auto";
        addRow.style.gap = "8px";
        addRow.style.alignItems = "center";

        const patternInput = addRow.createEl("input", { type: "text" });
        patternInput.placeholder = "例如 archive/** 或 *draft* 或 README.md";

        const targetSelect = addRow.createEl("select");
        targetSelect.createEl("option", { text: "全维度匹配", value: "all" });
        targetSelect.createEl("option", { text: "仅文件夹", value: "folder" });
        targetSelect.createEl("option", { text: "仅文件", value: "file" });

        const syntaxSelect = addRow.createEl("select");
        syntaxSelect.createEl("option", { text: "通配符 (* / ** / ?)", value: "glob" });
        syntaxSelect.createEl("option", { text: "精确匹配 (完全一致)", value: "exact" });
        syntaxSelect.createEl("option", { text: "正则表达式 (高级)", value: "regex" });

        patternInput.oninput = () => {
            const val = patternInput.value.trim();
            if (val.includes("*") || val.includes("?")) {
                syntaxSelect.value = "glob";
            } else if (!val.includes("/") && val.length > 0) {
                syntaxSelect.value = "exact";
            }
        };

        const addBtn = addRow.createEl("button", { text: "添加", cls: "mod-cta" });
        addBtn.onclick = () => {
            const pat = patternInput.value.trim();
            if (!pat) {
                new Notice("请输入规则表达式");
                return;
            }
            this.filterConfig.userRules.push({
                id: uid("rule"),
                pattern: pat,
                target: targetSelect.value as any,
                syntax: syntaxSelect.value as any,
                enabled: true,
            });
            this.render();
        };

        // 5. 提示说明
        const tips = this.contentEl.createEl("div");
        tips.style.marginTop = "10px";
        tips.style.fontSize = "12px";
        tips.style.color = "var(--text-muted)";
        tips.createEl("p", {
            text: "💡 提示：支持零门槛通配符（* 匹配任意单层字符，** 匹配多级目录，? 匹配单个字符）与精确完全匹配。修改规则会将已同步的受排除知识库镜像移至回收站，不会删除外部源。",
            attr: { style: "margin: 4px 0;" },
        });

        // 6. 底部操作栏
        const actions = this.contentEl.createDiv({ cls: "smartmd-external-actions" });
        actions.createEl("button", { text: "取消" }).onclick = () => this.close();
        const saveBtn = actions.createEl("button", { text: "保存并立即生效", cls: "mod-cta" });
        saveBtn.onclick = async () => {
            saveBtn.disabled = true;
            saveBtn.setText("正在保存并重校准…");
            try {
                await this.service.updateFilterConfig(this.binding.id, this.filterConfig);
                new Notice(`已更新“${this.binding.nickname}”的筛选规则并完成同步基线重校准！`, 5000);
                this.close();
                this.onSaved?.();
            } catch (err: any) {
                new Notice(`保存失败：${String(err.message || err)}`, 8000);
            } finally {
                saveBtn.disabled = false;
                saveBtn.setText("保存并立即生效");
            }
        };
    }
}

export class ExternalSyncManagerModal extends Modal {
    private unsubscribe: (() => void) | null = null;

    constructor(app: App, private service: ExternalSyncService) {
        super(app);
    }

    onOpen(): void {
        this.titleEl.setText("外部文件夹同步");
        this.render();
        this.unsubscribe = this.service.subscribeStatus(() => this.render());
    }

    onClose(): void {
        this.unsubscribe?.();
        this.unsubscribe = null;
        this.contentEl.empty();
    }

    private render(): void {
        this.contentEl.empty();
        this.contentEl.createEl("p", {
            text: "每个外部文件夹都是独立同步项，可以分别暂停、立即同步或解除绑定。对其中一个文件夹执行操作，不会影响其他同步文件夹。",
        });
        this.contentEl.createEl("button", { text: "添加同步文件夹", cls: "mod-cta" }).onclick = () => {
            new ExternalBindingModal(this.app, this.service).open();
        };
        const bindings = this.service.bindings();
        if (!bindings.length) this.contentEl.createEl("p", { text: "暂无外部同步文件夹。" });
        for (const binding of bindings) {
            const card = this.contentEl.createDiv({ cls: "smartmd-card" });
            card.createEl("h3", { text: binding.nickname });
            card.createEl("p", { text: binding.externalPath });
            const filterSummary = binding.filterConfig?.mode === "include"
                ? `白名单模式 (${(binding.filterConfig.userRules || []).length} 条规则)`
                : `排除模式 (${(binding.filterConfig?.userRules || []).length} 条规则)`;
            card.createEl("p", {
                text: `${this.service.statusLabel(binding.status)} · 筛选: ${filterSummary} · 待处理 ${binding.pendingCount || 0} · 冲突 ${(binding.conflicts || []).length} · 可恢复 ${(binding.recoveryItems || []).length}`,
            });
            if (binding.lastSuccessAt) card.createEl("p", { text: `上次成功：${new Date(binding.lastSuccessAt).toLocaleString()}` });
            if (binding.lastError) card.createEl("p", { text: binding.lastError, cls: "smartmd-warning" });
            const actions = card.createDiv({ cls: "smartmd-actions" });
            if (binding.status === "active") {
                actions.createEl("button", { text: "暂停" }).onclick = async () => { await this.service.pause(binding.id); this.render(); };
            } else {
                actions.createEl("button", { text: "恢复并重新比对" }).onclick = async () => { await this.service.resume(binding.id); this.render(); };
            }
            actions.createEl("button", { text: "立即同步" }).onclick = async () => { await this.service.syncBinding(binding); this.render(); };
            actions.createEl("button", { text: "⚙️ 筛选规则" }).onclick = () => {
                new ExternalFilterModal(this.app, this.service, binding, () => this.render()).open();
            };
            if ((binding.conflicts || []).length) {
                actions.createEl("button", { text: "查看冲突" }).onclick = () => this.openConflicts(binding);
            }
            const pendingDeletes = this.service.listPendingDeletions(binding.id);
            if (pendingDeletes.length) {
                actions.createEl("button", { text: "待人工确认的删除 " + pendingDeletes.length + " 篇" }).onclick =
                    () => this.openPendingDeletions(binding);
            }
            const pendingExports = this.service.listPendingExports(binding.id);
            if (pendingExports.length) {
                actions.createEl("button", { text: "待确认外写 " + pendingExports.length + " 篇" }).onclick =
                    () => this.openPendingExports(binding);
            }
            if ((binding.recoveryItems || []).length) {
                actions.createEl("button", { text: "查看恢复区" }).onclick = () => this.openRecovery(binding);
            }
            actions.createEl("button", { text: "解除此文件夹绑定" }).onclick = async () => {
                const confirmed = window.confirm(
                    `确定解除“${binding.nickname}”的绑定吗？\n\n只会停止这个文件夹的同步，其他同步文件夹会继续运行。外部原文件夹和知识库内现有副本都会保留。`,
                );
                if (confirmed) {
                    await this.service.remove(binding.id);
                    this.render();
                }
            };
        }
        const detached = this.service.recoveryArchives().filter((item: any) =>
            item.detached === true && Array.isArray(item.recoveryItems) && item.recoveryItems.length > 0,
        );
        if (detached.length) this.contentEl.createEl("h3", { text: "已解绑的恢复档案" });
        for (const archive of detached) {
            const card = this.contentEl.createDiv({ cls: "smartmd-card" });
            card.createEl("strong", { text: archive.nickname || "已解绑文件夹" });
            card.createEl("p", {
                text: `保留 ${archive.recoveryItems.length} 条恢复记录；仅能独立另存，不会重新授权外部同步。`,
            });
            card.createEl("button", { text: "查看解绑恢复档案" }).onclick =
                () => this.openRecovery(archive);
        }
    }

    private openConflicts(binding: any): void {
        this.contentEl.empty();
        this.contentEl.createEl("h3", { text: `${binding.nickname} · 冲突中心` });
        for (const conflict of binding.conflicts || []) {
            const card = this.contentEl.createDiv({ cls: "smartmd-card" });
            card.createEl("strong", { text: conflict.relativePath });
            card.createEl("p", { text: "两端内容均已完整保留，解决前不会互相覆盖。" });
            if (conflict.externalExists === false) card.createEl("p", { text: "外部版本：已删除", cls: "smartmd-warning" });
            if (conflict.vaultExists === false) card.createEl("p", { text: "知识库版本：已删除", cls: "smartmd-warning" });
            const actions = card.createDiv({ cls: "smartmd-actions" });
            actions.createEl("button", { text: "保留外部" }).onclick = async () => {
                try { await this.service.resolveConflict(binding.id, conflict.id, "external"); }
                catch (e: any) { new Notice(String(e.message || e), 8000); }
                this.render();
            };
            actions.createEl("button", { text: "保留知识库" }).onclick = async () => {
                try { await this.service.resolveConflict(binding.id, conflict.id, "vault"); }
                catch (e: any) { new Notice(String(e.message || e), 8000); }
                this.render();
            };
            actions.createEl("button", { text: "保留两份", cls: "mod-cta" }).onclick = async () => {
                try { await this.service.resolveConflict(binding.id, conflict.id, "both"); }
                catch (e: any) { new Notice(String(e.message || e), 8000); }
                this.render();
            };
        }
        this.contentEl.createEl("button", { text: "返回" }).onclick = () => this.render();
    }

    private openPendingDeletions(binding: any): void {
        this.contentEl.empty();
        this.contentEl.createEl("h3", { text: binding.nickname + " · 待确认删除" });
        this.contentEl.createEl("p", { text: "单端文件缺失不等于允许删除另一端。只有本批次用户明确选择的当前版本才会执行；外部原件移入隐藏恢复备份，知识库副本移入回收站，并保留独立恢复档案。" });
        const entries = this.service.listPendingDeletions(binding.id);
        const selected = new Set<string>();
        for (const entry of entries) {
            const row = this.contentEl.createDiv({ cls: "smartmd-card" });
            const input = row.createEl("input", { type: "checkbox" });
            input.onchange = () => {
                if (input.checked) selected.add(entry.relativePath);
                else selected.delete(entry.relativePath);
            };
            row.createEl("strong", { text: entry.relativePath });
            row.createEl("p", { text: entry.direction === "vault" ?
                "将把外部原文件移到可恢复档案" : "将把知识库映射笔记移至Obsidian回收站" });
            row.createEl("p", { text: "当前版本SHA256前缀：" + entry.hash.slice(0, 16) });
        }
        this.contentEl.createEl("button", { text: "确认本批删除范围", cls: "mod-cta" }).onclick = async () => {
            const chosen = [...selected];
            if (!chosen.length) { new Notice("请先勾选本批目标"); return; }
            if (!window.confirm("将按勾选清单移走另一端的真实文件：" +
                chosen.join("、") + "。操作前会建立恢复档案；确定本次版本与范围吗？")) return;
            try {
                await this.service.confirmPendingDeletions(binding.id, chosen, true);
                new Notice("已按确认清单完成处理，恢复记录已保留");
            } catch (error: any) {
                new Notice("删除处理已停止：" + String(error.message || error), 8000);
            }
            this.render();
        };
        this.contentEl.createEl("button", { text: "取消返回" }).onclick = () => this.render();
    }

    private openPendingExports(binding: any): void {
        this.contentEl.empty();
        this.contentEl.createEl("h3", { text: binding.nickname + " · 审核待写出新稿" });
        this.contentEl.createEl("p", { text: "只有本次勾选并确认的版本才会写往外部源目录。外部已有文件、内容变化或磁盘身份改变一律拒绝覆盖。" });
        const items = this.service.listPendingExports(binding.id);
        const selected = new Set<string>();
        for (const item of items) {
            const row = this.contentEl.createDiv({ cls: "smartmd-card" });
            const toggle = row.createEl("input", { type: "checkbox" });
            toggle.onchange = () => { if (toggle.checked) selected.add(item.relativePath); else selected.delete(item.relativePath); };
            row.createEl("strong", { text: "本地：" + item.vaultPath });
            row.createEl("p", { text: "外写位置：" + item.externalPath });
            row.createEl("p", { text: "本地版本摘要：" + item.hash.slice(0, 16) });
        }
        this.contentEl.createEl("button", { text: "确认所选文件外写", cls: "mod-cta" }).onclick = async () => {
            const chosen = [...selected];
            if (!chosen.length) { new Notice("请先勾选文件"); return; }
            if (!window.confirm("仅本次批准向外部目录新建下列 " + chosen.length +
                " 篇 Markdown（不会覆盖已有文件）：\n" + chosen.join("\n") + "\n确认本次写出？")) return;
            try {
                await this.service.confirmPendingExports(binding.id, chosen, true);
                new Notice("本批外写已执行，已核对写回内容");
            } catch (error: any) {
                new Notice("写出停止：" + String(error.message || error), 8000);
            }
            this.render();
        };
        this.contentEl.createEl("button", { text: "取消并返回" }).onclick = () => this.render();
    }

    private openRecovery(binding: any): void {
        this.contentEl.empty();
        this.contentEl.createEl("h3", { text: `${binding.nickname} · 7 天恢复区` });
        if (binding.detached === true) {
            this.contentEl.createEl("p", {
                text: "已解绑：只允许将恢复记录另存为独立笔记，不会恢复原外部路径的写入或同步权限。",
                cls: "smartmd-warning",
            });
        }
        for (const item of binding.recoveryItems || []) {
            const card = this.contentEl.createDiv({ cls: "smartmd-card" });
            card.createEl("strong", { text: item.relativePath });
            card.createEl("p", { text: `${item.source === "external" ? "外部来源" : "知识库来源"} · ${new Date(item.deletedAt).toLocaleString()}` });
            if (!item.restoredAt && !item.exportedAt) {
                card.createEl("button", { text: "另存恢复副本（不覆盖）" }).onclick = async () => {
                    try {
                        const location = await this.service.exportRecoveryToNewNote(binding.id, item.id);
                        new Notice("已安全另存到：" + location);
                        this.render();
                    } catch (e: any) { new Notice(String(e.message || e), 8000); }
                };
            }
            if (!item.restoredAt && !item.exportedAt && binding.detached !== true) {
                card.createEl("button", { text: "恢复到两端", cls: "mod-cta" }).onclick = async () => {
                    try { await this.service.restore(binding.id, item.id); this.render(); }
                    catch (e: any) { new Notice(String(e.message || e), 8000); }
                };
            } else if (item.restoredAt || item.exportedAt) {
                card.createEl("p", { text: item.resolvedAt ? "已确认解决；缓冲保留 7 天" :
                    (item.exportedAt ? "已安全导出独立笔记；待确认解决" : "已恢复并核验；待确认解决") });
                if (!item.resolvedAt) {
                    card.createEl("button", { text: "确认恢复问题已解决" }).onclick = async () => {
                        if (!window.confirm("已逐项核验外部和知识库两端内容？确认后仍保留 7 天安全缓冲。")) return;
                        try { await this.service.confirmRecoveryResolved(binding.id, item.id); this.render(); }
                        catch (e: any) { new Notice(String(e.message || e), 8000); }
                    };
                }
            }
        }
        this.contentEl.createEl("button", { text: "返回" }).onclick = () => this.render();
    }
}

export class ExternalSyncService {
    private fs: any;
    private path: any;
    private running = new Set<string>();
    private syncJobs = new Map<string, Promise<void>>();
    private authority = new OperationAuthority();

    constructor(private app: App, private store: SmartMDStore) {
        if (!Platform.isDesktopApp) return;
        this.fs = require("fs");
        this.path = require("path");
    }

    bindings(): any[] {
        return (this.store as any).settings.local.externalBindings || [];
    }

    recoveryArchives(): any[] {
        const local = (this.store as any).settings.local;
        local.externalRecoveryArchives = Array.isArray(local.externalRecoveryArchives)
            ? local.externalRecoveryArchives
            : [];
        return local.externalRecoveryArchives;
    }

    subscribeStatus(listener: () => void): () => void {
        return (this.store as any).subscribe?.(listener) || (() => {});
    }

    private knownBindings(): any[] {
        return [...this.bindings(), ...this.recoveryArchives()];
    }

    openManager(): void {
        if (!Platform.isDesktopApp) return;
        new ExternalSyncManagerModal(this.app, this).open();
    }

    openAddBinding(onAdded?: () => void): void {
        if (!Platform.isDesktopApp) return;
        new ExternalBindingModal(this.app, this, onAdded).open();
    }

    openFilterModal(binding: any, onSaved?: () => void): void {
        if (!Platform.isDesktopApp) return;
        new ExternalFilterModal(this.app, this, binding, onSaved).open();
    }

    async updateFilterConfig(id: string, filterConfig: ExternalFilterConfig): Promise<void> {
        const binding = this.bindings().find((item) => item.id === id);
        if (!binding) return;
        const issues = FilterMatcher.validateConfig(filterConfig);
        if (issues.length) throw new Error("筛选规则不安全：" + issues.join("；"));
        // Wait for the old job before publishing a new configuration.
        // 并发保护：如果当前正在运行旧任务，等待旧任务完全结束
        while (this.syncJobs.has(id) || this.running.has(id)) {
            const existingJob = this.syncJobs.get(id);
            if (existingJob) {
                try { await existingJob; } catch {}
                if (this.syncJobs.get(id) === existingJob) {
                    this.syncJobs.delete(id);
                }
            } else {
                await new Promise((r) => setTimeout(r, 10));
            }
        }

        filterConfig.filterHash = FilterMatcher.computeFilterHash(filterConfig);
        binding.filterConfig = JSON.parse(JSON.stringify(filterConfig));
        await (this.store as any).save();
        await this.syncBinding(binding);
        if (binding.status === "offline" || binding.status === "failed") {
            throw new Error("规则已保存但尚未安全执行，请检查源目录与恢复情况");
        }
    }

    listPendingDeletions(bindingId: string): Array<{
        relativePath: string; direction: "vault" | "external"; hash: string;
    }> {
        const binding = this.bindings().find((b: any) => b.id === bindingId);
        if (!binding) return [];
        return Object.entries(binding.fileStates || {})
            .filter(([, entry]) => {
                const state: any = entry;
                return state?.pendingDeletion === "vault" || state?.protectedDelete === "external";
            })
            .map(([relativePath, entry]) => {
                const state: any = entry;
                return {
                    relativePath,
                    direction: state.pendingDeletion === "vault" ? "vault" : "external",
                    hash: String(state.pendingDeletion === "vault" ? state.externalHash : state.vaultHash),
                };
            });
    }

    async confirmPendingDeletions(bindingId: string, relativePaths: string[], approved = false): Promise<void> {
        if (!approved) throw new Error("缺少针对本次目标文件与版本的明确删除确认");
        const binding = this.bindings().find((b: any) => b.id === bindingId);
        if (!binding || !this.canRun(binding)) throw new Error("绑定未获有效的删除确认许可");
        const chosen = [...new Set(relativePaths)];
        if (!chosen.length || chosen.length > 100) throw new Error("请选择1至100篇逐批处理");
        const existing = this.syncJobs.get(binding.id);
        if (existing) await existing;
        const job = this.performPendingDeletions(binding, chosen);
        this.syncJobs.set(binding.id, job);
        try {
            await job;
        } finally {
            if (this.syncJobs.get(binding.id) === job) this.syncJobs.delete(binding.id);
        }
    }

    private async performPendingDeletions(binding: any, relativePaths: string[]): Promise<void> {
        await this.assertBindingRoot(binding);
        const ticket = this.authority.start(binding, binding.externalPath + "|" + binding.targetPath);
        try {
            for (const relativePath of relativePaths) {
                this.authority.require(ticket);
                await this.assertSafeExternalRelativePath(binding, relativePath);
                const item = binding.fileStates?.[relativePath];
                if (!item || (!item.pendingDeletion && !item.protectedDelete)) {
                    throw new Error("文件不再属于本次删除清单，请重新预览：" + relativePath);
                }
                const current = await this.currentPair(binding, relativePath);
                if (item.pendingDeletion === "vault") {
                    if (!current.external || current.vault ||
                        current.external.hash !== item.externalHash) {
                        throw new Error("外部原稿身份、版本或删除意图已变化：" + relativePath);
                    }
                    const backup = await this.createRecovery(binding, relativePath, current.external.content, "external");
                    const archived = await this.stageExternalDelete(binding, relativePath,
                        current.external.absolutePath,
                        () => this.isCurrent(binding, ticket),
                        () => this.authority.beginFinal(ticket),
                        current.external.hash,
                    );
                    if (await digest(await this.fs.promises.readFile(archived, "utf8")) !== current.external.hash) {
                        throw new Error("外部原稿归档后读回不匹配，原备份已保留待人工核查");
                    }
                    backup.externalBackupPath = archived;
                    delete binding.fileStates[relativePath];
                } else if (item.protectedDelete === "external") {
                    if (current.external || !current.vault || current.vault.hash !== item.vaultHash) {
                        throw new Error("知识库原稿在删除确认前变化：" + relativePath);
                    }
                    await this.createRecovery(binding, relativePath, current.vault.content, "vault");
                    if (typeof (this.app.vault as any).trash !== "function") {
                        throw new Error("Obsidian回收站不可用，拒绝确认删除");
                    }
                    this.authority.beginFinal(ticket);
                    const recheck = await this.app.vault.cachedRead(current.vault.file);
                    if (await digest(recheck) !== current.vault.hash) {
                        throw new Error("知识库正文再次变化，保留原稿");
                    }
                    await (this.app.vault as any).trash(current.vault.file, false);
                    delete binding.fileStates[relativePath];
                } else {
                    throw new Error("未知的待确认删除方向");
                }
                await (this.store as any).save();
            }
        } finally {
            this.authority.finishFinal(ticket);
        }
    }

    listPendingExports(bindingId: string): Array<{ relativePath: string; vaultPath: string; externalPath: string; hash: string }> {
        const binding = this.bindings().find((b: any) => b.id === bindingId);
        if (!binding) return [];
        return Object.entries(binding.fileStates || {})
            .filter(([, state]) => (state as any)?.pendingExport === true)
            .map(([relativePath, state]) => ({
                relativePath,
                vaultPath: normalizePath(binding.targetPath + "/" + relativePath),
                externalPath: this.path.join(binding.externalPath, ...relativePath.split("/")),
                hash: String((state as any).vaultHash || ""),
            }));
    }

    async confirmPendingExports(bindingId: string, relativePaths: string[], approved = false): Promise<void> {
        if (!approved) throw new Error("必须经用户审阅并明确确认本次文件清单");
        const binding = this.bindings().find((b: any) => b.id === bindingId);
        if (!binding || !this.canRun(binding)) throw new Error("同步绑定尚未获得写入许可");
        const chosen = [...new Set(relativePaths)];
        if (!chosen.length || chosen.length > 100) throw new Error("请选择 1 至 100 篇文件逐批确认");
        const active = this.syncJobs.get(binding.id);
        if (active) await active;
        // Serialize an approval transaction with the poller and manual sync.
        const job = this.performPendingExports(binding, chosen);
        this.syncJobs.set(binding.id, job);
        try {
            await job;
        } finally {
            if (this.syncJobs.get(binding.id) === job) this.syncJobs.delete(binding.id);
        }
    }

    private async performPendingExports(binding: any, paths: string[]): Promise<void> {
        await this.assertBindingRoot(binding);
        const ticket = this.authority.start(binding, binding.externalPath + "|" + binding.targetPath);
        try {
            for (const relativePath of paths) {
                const expected = binding.fileStates?.[relativePath];
                if (!expected?.pendingExport || !FilterMatcher.isFileAccepted(relativePath, binding.filterConfig)) {
                    throw new Error("文件不在本次待写出清单中：" + relativePath);
                }
                await this.assertSafeExternalRelativePath(binding, relativePath);
                const fresh = await this.currentPair(binding, relativePath);
                if (fresh.external || !fresh.vault || fresh.vault.hash !== expected.vaultHash) {
                    throw new Error("目标已存在或本地稿变化，请重新预览：" + relativePath);
                }
                const dest = this.path.join(binding.externalPath, ...relativePath.split("/"));
                await this.atomicWrite(
                    dest,
                    fresh.vault.content,
                    () => this.isCurrent(binding, ticket),
                    async () => {
                        const current = await this.currentPair(binding, relativePath);
                        return !current.external && current.vault?.hash === expected.vaultHash;
                    },
                    () => this.authority.beginFinal(ticket),
                );
                const readBack = await this.fs.promises.readFile(dest, "utf8");
                if (await digest(readBack) !== expected.vaultHash) {
                    throw new Error("外写读回校验失败，需人工核查：" + relativePath);
                }
                const stat = await this.fs.promises.stat(dest);
                binding.fileStates[relativePath] = {
                    externalHash: expected.vaultHash,
                    vaultHash: expected.vaultHash,
                    externalMtimeMs: Math.floor(stat.mtimeMs),
                    externalSize: stat.size,
                    vaultMtime: fresh.vault.file.stat?.mtime,
                    vaultSize: fresh.vault.file.stat?.size,
                };
                await (this.store as any).save();
            }
        } finally {
            this.authority.finishFinal(ticket);
        }
    }

    statusLabel(status: string): string {
        return ({
            pending: "待确认",
            scanning: "首次扫描",
            active: "同步正常",
            paused: "已暂停",
            offline: "外部文件夹离线",
            conflict: "存在冲突",
            failed: "同步失败",
            stopping: "正在停止",
        } as Record<string, string>)[status] || status;
    }

    async chooseFolder(): Promise<string> {
        // Level 1: 原生对话框探测
        try {
            const electron = require("electron");
            const dialog = electron.remote?.dialog || electron.dialog;
            if (dialog?.showOpenDialog) {
                const result = await dialog.showOpenDialog({ properties: ["openDirectory"] });
                if (!result?.canceled && result?.filePaths?.[0]) return result.filePaths[0];
            }
        } catch {}

        // Level 2: HTML5 WebkitDirectory 隐式选择器
        try {
            return await new Promise<string>((resolve) => {
                const input = document.createElement("input");
                input.type = "file";
                (input as any).webkitdirectory = true;
                input.style.display = "none";
                document.body.appendChild(input);
                input.onchange = () => {
                    const files = input.files;
                    document.body.removeChild(input);
                    if (!files || files.length === 0) return resolve("");
                    const firstFile: any = files[0];
                    const fullPath: string = firstFile.path || "";
                    const relativePath: string = firstFile.webkitRelativePath || "";
                    if (fullPath && relativePath) {
                        const parts = relativePath.split(/[\/\\]/).filter(Boolean);
                        let root = this.path.resolve(fullPath);
                        for (let i = 1; i < parts.length; i++) root = this.path.dirname(root);
                        // The directory picker must never guess the root using
                        // an ancestor folder with an identical name.
                        const expected = this.path.join(root, ...parts.slice(1));
                        const same = (a: string, b: string) =>
                            this.pathKey(a) === this.pathKey(b);
                        if (parts.length >= 2 &&
                            same(this.path.basename(root), parts[0]) &&
                            same(expected, fullPath)) {
                            return resolve(root);
                        }
                    }
                    // An unreliable selection requires a manual absolute path.
                    resolve("");
                };
                input.oncancel = () => {
                    document.body.removeChild(input);
                    resolve("");
                };
                input.click();
            });
        } catch {
            return "";
        }
    }

    async addBindingConfirmed(externalPath: string, nickname?: string, filterConfig?: ExternalFilterConfig): Promise<any> {
        if (!externalPath) throw new Error("请选择外部文件夹");
        const absolute = await this.validatePath(externalPath);
        const files = await this.scanExternal(absolute, filterConfig);
        const display = safeName(nickname || this.path.basename(absolute));
        const target = await this.uniqueTarget(`外部同步文件夹/${display}`);
        const conflicts: string[] = [];
        for (const file of files) {
            const targetFile = this.app.vault.getAbstractFileByPath(normalizePath(`${target}/${file.relativePath}`));
            if (targetFile instanceof TFile && await this.app.vault.cachedRead(targetFile) !== file.content) conflicts.push(file.relativePath);
        }
        await this.addBinding(absolute, display, target, files, conflicts, filterConfig);
        return this.bindings().find((b: any) => b.externalPath === absolute);
    }

    async previewAndAdd(folderPath: string, nickname: string, filterConfig?: ExternalFilterConfig): Promise<void> {
        if (!folderPath) throw new Error("请选择外部文件夹");
        const absolute = await this.validatePath(folderPath);
        const files = await this.scanExternal(absolute, filterConfig);
        const display = safeName(nickname || this.path.basename(absolute));
        const target = await this.uniqueTarget(`外部同步文件夹/${display}`);
        const conflicts: string[] = [];
        for (const file of files) {
            const targetFile = this.app.vault.getAbstractFileByPath(normalizePath(`${target}/${file.relativePath}`));
            if (targetFile instanceof TFile && await this.app.vault.cachedRead(targetFile) !== file.content) conflicts.push(file.relativePath);
        }
        const filterSummary = filterConfig?.mode === "include"
            ? `白名单模式 (规则数: ${(filterConfig.userRules || []).length})`
            : `排除模式 (规则数: ${(filterConfig?.userRules || []).length}，已自动排除构建与IDE项)`;
        const preview = { markdownCount: files.length, targetPath: target, conflicts, filterSummary };
        new BindingPreviewModal(this.app, preview, async () => {
            await this.addBinding(absolute, display, target, files, conflicts, filterConfig);
        }).open();
    }

    private async addBinding(
        externalPath: string,
        nickname: string,
        targetPath: string,
        files: ScannedFile[],
        conflicts: string[],
        filterConfig?: ExternalFilterConfig,
    ): Promise<void> {
        const canonical = await this.validatePath(externalPath);
        if (this.pathKey(canonical) !== this.pathKey(externalPath)) {
            throw new Error("首次预览后源目录身份变化，请重新预览");
        }
        const freshFiles = await this.scanExternal(canonical, filterConfig);
        const oldVersions = new Map(files.map((file) => [file.relativePath, file.hash]));
        if (freshFiles.length !== files.length ||
            freshFiles.some((file) => oldVersions.get(file.relativePath) !== file.hash)) {
            throw new Error("首次预览后源文件已变化，请重新预览");
        }
        const rootIdentity = await this.fs.promises.stat(canonical);
        const initialFilterConfig: ExternalFilterConfig = filterConfig || {
            mode: "exclude",
            builtinPresets: { ignoreBuildArtifacts: true, ignoreIDEConfig: true },
            userRules: [],
            filterHash: "default_filter_hash",
        };
        initialFilterConfig.filterHash = FilterMatcher.computeFilterHash(initialFilterConfig);

        const binding: any = {
            id: uid("external"),
            nickname,
            externalPath,
            canonicalExternalPath: externalPath,
            rootDev: rootIdentity.dev,
            rootIno: rootIdentity.ino,
            targetPath,
            status: conflicts.length ? "conflict" : "active",
            lastSuccessAt: 0,
            lastError: "",
            pendingCount: 0,
            fileStates: {},
            conflicts: [],
            recoveryItems: [],
            legacy: false,
            confirmed: true,
            runAllowed: true,
            confirmationScope: `${externalPath}|${targetPath}`,
            filterConfig: initialFilterConfig,
            appliedFilterHash: initialFilterConfig.filterHash,
        };
        this.bindings().push(binding);
        await (this.store as any).save();
        const ticket = this.authority.start(binding);
        try {
            for (const item of files) {
                // Per-file CAS: the source may change after the batch preview.
                const freshContent = await this.fs.promises.readFile(item.absolutePath, "utf8");
                if (await digest(freshContent) !== item.hash) {
                    throw new Error("首次导入源文件在确认后变化，请重新预览：" + item.relativePath);
                }
                const vaultPath = normalizePath(`${targetPath}/${item.relativePath}`);
                const current = this.app.vault.getAbstractFileByPath(vaultPath);
                if (current instanceof TFile) {
                    const vaultContent = await this.app.vault.cachedRead(current);
                    if (vaultContent !== item.content) {
                        binding.conflicts.push(await this.makeConflict(binding, item.relativePath, item.content, vaultContent, ticket));
                        continue;
                    }
                } else {
                    await this.stageVaultWrite(binding, item.relativePath, item.content, ticket,
                        async () => {
                            const currentText = await this.fs.promises.readFile(item.absolutePath, "utf8");
                            return await digest(currentText) === item.hash;
                        });
                }
                binding.fileStates[item.relativePath] = {
                    externalHash: item.hash,
                    vaultHash: item.hash,
                    externalMtimeMs: item.mtimeMs,
                    externalSize: item.size,
                    vaultMtime: (this.app.vault.getAbstractFileByPath(vaultPath) as TFile)?.stat?.mtime,
                    vaultSize: (this.app.vault.getAbstractFileByPath(vaultPath) as TFile)?.stat?.size,
                    lastDeepHashAt: Date.now(),
                };
            }
        } catch (error) {
            const stopped = this.authority.stop(binding);
            binding.status = "stopping";
            await (this.store as any).save();
            this.authority.finishFinal(ticket);
            await stopped;
            binding.status = "paused";
            await (this.store as any).save();
            throw error;
        }
        binding.lastSuccessAt = Date.now();
        await (this.store as any).save();
        this.authority.finishFinal(ticket);
        new Notice(`已绑定 ${nickname}，首次导入未覆盖或删除任何内容`);
    }

    private pathKey(value: string): string {
        const normalized = this.path.normalize(this.path.resolve(value)).replace(/[\\/]+$/, "");
        return process.platform === "win32" ? normalized.toLocaleLowerCase() : normalized;
    }

    private async canonicalPath(value: string): Promise<string> {
        return this.path.normalize(await this.fs.promises.realpath(this.path.resolve(value)));
    }

    private async validatePath(externalPath: string): Promise<string> {
        const canonical = await this.canonicalPath(externalPath).catch(() => "");
        const stat = canonical ? await this.fs.promises.stat(canonical).catch(() => null) : null;
        if (!stat?.isDirectory()) throw new Error("路径不存在、不可访问或不是文件夹");
        const rawVaultRoot = this.path.resolve((this.app.vault.adapter as any).basePath || "");
        const vaultRoot = await this.canonicalPath(rawVaultRoot).catch(() => rawVaultRoot);
        if (inside(vaultRoot, canonical, this.path) || inside(canonical, vaultRoot, this.path)) throw new Error("外部文件夹不能位于知识库内，也不能包含知识库");
        for (const binding of this.bindings()) {
            const existing = await this.canonicalPath(binding.externalPath).catch(() =>
                binding.canonicalExternalPath || this.path.resolve(binding.externalPath));
            if (
                this.pathKey(existing) === this.pathKey(canonical)
                || inside(existing, canonical, this.path)
                || inside(canonical, existing, this.path)
            ) {
                throw new Error(`该范围与“${binding.nickname}”重复或上下级重叠`);
            }
        }
        return canonical;
    }

    private async uniqueTarget(base: string): Promise<string> {
        let candidate = normalizePath(base);
        let number = 2;
        while (this.bindings().some((item) => item.targetPath === candidate)) candidate = `${normalizePath(base)} (${number++})`;
        return candidate;
    }

    private async assertBindingRoot(binding: any): Promise<void> {
        const actual = await this.canonicalPath(binding.externalPath);
        const stat = await this.fs.promises.stat(actual);
        if (!stat.isDirectory() ||
            this.pathKey(actual) !== this.pathKey(binding.canonicalExternalPath || binding.externalPath) ||
            (binding.rootDev !== undefined && binding.rootDev !== stat.dev) ||
            (binding.rootIno !== undefined && binding.rootIno !== stat.ino)) {
            throw new Error("外部目录身份或挂载设备变化，停止同步与清理");
        }
    }

    private async scanExternal(root: string, filterConfig?: ExternalFilterConfig, metadataOnly = false): Promise<ScannedFile[]> {
        const output: ScannedFile[] = [];
        const walk = async (current: string): Promise<void> => {
            const entries = await this.fs.promises.readdir(current, { withFileTypes: true });
            for (const entry of entries) {
                const absolutePath = this.path.join(current, entry.name);
                if (entry.isSymbolicLink()) continue;
                if (entry.isDirectory()) {
                    const dirRel = normalizePath(this.path.relative(root, absolutePath));
                    // 目录级早期剪枝（Directory Pruning）：命中系统不可逾越黑名单或排除规则直接短路 O(1) 跳过子树
                    if (FilterMatcher.shouldPruneDirectory(dirRel, entry.name, filterConfig)) {
                        continue;
                    }
                    await walk(absolutePath);
                } else if (entry.isFile() && entry.name.toLowerCase().endsWith(".md")) {
                    const relPath = normalizePath(this.path.relative(root, absolutePath));
                    // 文件级复合判定（含全祖先目录校验）
                    if (!FilterMatcher.isFileAccepted(relPath, filterConfig)) {
                        continue;
                    }
                    const stat = await this.fs.promises.stat(absolutePath);
                    if (!stat.isFile()) throw new Error("外部文件类型变化，取消本轮扫描");
                    if (stat.size > 16 * 1024 * 1024) {
                        throw new Error("单篇 Markdown 超过 16MiB 安全扫描上限：" + relPath);
                    }
                    if (output.length >= 20000) throw new Error("外部 Markdown 数量超过 20000，暂停扩大扫描");
                    const content = metadataOnly ? "" : await this.fs.promises.readFile(absolutePath, "utf8");
                    output.push({
                        relativePath: relPath,
                        absolutePath,
                        content,
                        hash: metadataOnly ? "" : await digest(content),
                        mtimeMs: Math.floor(stat.mtimeMs),
                        size: stat.size,
                    });
                }
            }
        };
        await walk(root);
        return output;
    }

    private async scanVault(binding: any, metadataOnly = false): Promise<Map<string, { file: TFile; content: string; hash: string; mtime?: number; size?: number }>> {
        const result = new Map<string, { file: TFile; content: string; hash: string; mtime?: number; size?: number }>();
        const prefix = `${normalizePath(binding.targetPath)}/`;
        for (const file of this.app.vault.getMarkdownFiles()) {
            if (!file.path.startsWith(prefix)) continue;
            const relPath = file.path.slice(prefix.length);
            // 双向对称性网关拦截：受排除路径下的本地文件不参与双向对账
            if (!FilterMatcher.isFileAccepted(relPath, binding.filterConfig)) {
                continue;
            }
            const content = metadataOnly ? "" : await this.app.vault.cachedRead(file);
            result.set(relPath, {
                file,
                content,
                hash: metadataOnly ? "" : await digest(content),
                mtime: file.stat?.mtime,
                size: file.stat?.size,
            });
        }
        return result;
    }

    async syncAll(): Promise<void> {
        if (!Platform.isDesktopApp) return;
        for (const binding of this.bindings()) {
            if (this.canRun(binding)) await this.syncBinding(binding);
        }
        await this.purgeExpiredRecovery();
    }

    async syncBinding(binding: any): Promise<void> {
        const existing = this.syncJobs.get(binding.id);
        if (existing) return await existing;
        const job = this.performSyncBinding(binding);
        this.syncJobs.set(binding.id, job);
        try {
            await job;
        } finally {
            if (this.syncJobs.get(binding.id) === job) this.syncJobs.delete(binding.id);
        }
    }

    private async performSyncBinding(binding: any): Promise<void> {
        if (
            !Platform.isDesktopApp
            || this.running.has(binding.id)
            || !this.canRun(binding)
        ) return;
        const ticket = this.authority.start(binding, `${binding.externalPath}|${binding.targetPath}`);
        const partial: Record<string, any> = {};
        this.running.add(binding.id);
        try {
            binding.status = "syncing";
            binding.lastError = "";
            await (this.store as any).save();

            // Revalidate persisted rules before any pruning or file reads.
            // An unsafe legacy regex must not silently become "no match" and
            // broaden the set of files eligible for synchronization.
            if (binding.filterConfig) {
                const filterIssues = FilterMatcher.validateConfig(binding.filterConfig);
                if (filterIssues.length) {
                    throw new Error("筛选规则不安全，已停止同步：" + filterIssues.join("；"));
                }
            }
            // Validate the original device before ANY filter pruning.
            await this.assertBindingRoot(binding);
            // 规则漂移感知与基线安全脱钩（Filter Reconciliation Pipeline）
            const currentFilterHash = binding.filterConfig?.filterHash || (binding.filterConfig ? FilterMatcher.computeFilterHash(binding.filterConfig) : "");
            if (binding.filterConfig) {
                binding.filterConfig.filterHash = currentFilterHash;
            }
            if ((binding.appliedFilterHash || "") !== currentFilterHash) {
                await this.reconcileBaseline(binding);
            }

            const stat = await this.fs.promises.stat(binding.externalPath).catch(() => null);
            if (!stat?.isDirectory()) {
                await this.authority.stop(binding);
                binding.status = "offline";
                binding.lastError = "外部文件夹不可访问。已暂停同步，不会将离线解释为删除。";
                await (this.store as any).save();
                return;
            }
            let externalFiles: ScannedFile[];
            try {
                externalFiles = await this.scanExternal(binding.externalPath, binding.filterConfig, true);
                const completedRoot = await this.fs.promises.stat(binding.externalPath);
                if (!completedRoot.isDirectory()) throw new Error("外部同步根目录不可用");
            } catch (error) {
                await this.authority.stop(binding);
                binding.status = "offline";
                binding.lastError = error instanceof Error
                    ? error.message
                    : "外部同步根目录不可用。已暂停同步，不会将离线解释为删除。";
                await (this.store as any).save();
                return;
            }
            const external = new Map(externalFiles.map((item) => [item.relativePath, item]));
            const vault = await this.scanVault(binding, true);
            const previous = binding.fileStates || {};
            const allPaths = new Set([...external.keys(), ...vault.keys(), ...Object.keys(previous)]);
            const next: Record<string, any> = partial;
            binding.pendingCount = allPaths.size;
            let protectedDeletes = 0;

            // 防误删风暴小样本与全量抹除熔断机制 (Circuit Breaker)
            const baselinePaths = Object.keys(previous).filter((p) => previous[p]?.externalHash);
            const totalBaseline = baselinePaths.length;
            const missingCount = baselinePaths.filter((p) => !external.has(p)).length;
            const isStorm = missingCount > 5 || (totalBaseline >= 2 && missingCount === totalBaseline) || (totalBaseline >= 5 && (missingCount / totalBaseline) > 0.2);
            if (isStorm) {
                await this.authority.stop(binding);
                binding.status = "paused";
                binding.lastError = `检测到外部 ${missingCount} 篇文档批量缺失（疑似整目录移动或批量误操作），已触发防误删风暴熔断保护，同步已暂停。请在设置面板确认后再恢复。`;
                await (this.store as any).save();
                new Notice(`[SmartMD] ${binding.nickname} 触发防误删风暴熔断保护：外部 ${missingCount} 篇文档缺失，已紧急暂停`, 9000);
                return;
            }

            for (const relativePath of [...allPaths].sort()) {
                if (!this.isCurrent(binding, ticket)) return;
                const waitingConflict = (binding.conflicts || []).find((item: any) => item.relativePath === relativePath);
                if (waitingConflict) {
                    const latest = await this.currentPair(binding, relativePath);
                    if (!this.conflictMatches(waitingConflict, latest)) {
                        if (!latest.external && !latest.vault) {
                            binding.conflicts = binding.conflicts.filter((item: any) => item.id !== waitingConflict.id);
                            delete previous[relativePath];
                            await this.cleanupConflictCopies(waitingConflict);
                        } else {
                            const replacement = await this.makeConflict(
                                binding,
                                relativePath,
                                latest.external?.content ?? null,
                                latest.vault?.content ?? null,
                                ticket,
                            );
                            await this.cleanupConflictCopies(waitingConflict);
                            Object.assign(waitingConflict, replacement);
                        }
                    }
                    binding.status = binding.conflicts.length ? "conflict" : "syncing";
                    continue;
                }
                const ext = external.get(relativePath);
                const local = vault.get(relativePath);
                const old = previous[relativePath];

                // Quick stat-based skip MUST happen before any content read or hashing.
                if (ext && local && old && old.externalHash && old.vaultHash) {
                    const externalStable = old.externalMtimeMs !== undefined &&
                        old.externalMtimeMs === ext.mtimeMs && old.externalSize === ext.size;
                    const localStable = old.vaultMtime !== undefined &&
                        old.vaultMtime === local.mtime && old.vaultSize === local.size;
                    // Stat-only validation is unsafe for writes within a filesystem
                    // timestamp tick or when a file was changed without updating mtime.
                    // Require quiescent metadata and periodically rehash all files.
                    const now = Date.now();
                    const aged = Number.isFinite(ext.mtimeMs) && Number.isFinite(local.mtime)
                        && now - Number(ext.mtimeMs) > 4000
                        && now - Number(local.mtime) > 4000;
                    const recentDeepCheck = Number(old.lastDeepHashAt || 0) > 0
                        && now - Number(old.lastDeepHashAt) < 30000;
                    if (externalStable && localStable && aged && recentDeepCheck) {
                        next[relativePath] = old;
                        binding.pendingCount = Math.max(0, binding.pendingCount - 1);
                        continue;
                    }
                }
                // A source changed after the metadata snapshot: never adopt
                // its new bytes as though the original scan had observed them.
                if (ext) {
                    const beforeRead = await this.fs.promises.stat(ext.absolutePath);
                    if (Math.floor(beforeRead.mtimeMs) !== ext.mtimeMs || beforeRead.size !== ext.size) {
                        const latest = await this.currentPair(binding, relativePath);
                        if (latest.external && latest.vault &&
                            latest.external.hash !== latest.vault.hash) {
                            binding.conflicts.push(await this.makeConflict(
                                binding, relativePath, latest.external.content, latest.vault.content, ticket,
                            ));
                            binding.status = "conflict";
                        }
                        if (old) next[relativePath] = old;
                        binding.pendingCount = Math.max(0, binding.pendingCount - 1);
                        continue;
                    }
                    ext.content = await this.fs.promises.readFile(ext.absolutePath, "utf8");
                    ext.hash = await digest(ext.content);
                }
                if (local) {
                    local.content = await this.app.vault.cachedRead(local.file);
                    local.hash = await digest(local.content);
                }
                const fresh = await this.currentPair(binding, relativePath);
                if (!this.isCurrent(binding, ticket)) return;
                if (!this.samePair(ext, local, fresh)) {
                    if (
                        fresh.external
                        && fresh.vault
                        && fresh.external.hash !== fresh.vault.hash
                        && !(binding.conflicts || []).some((item: any) => item.relativePath === relativePath)
                    ) {
                        binding.conflicts.push(await this.makeConflict(
                            binding,
                            relativePath,
                            fresh.external.content,
                            fresh.vault.content,
                            ticket,
                        ));
                        binding.status = "conflict";
                    }
                    if (old) next[relativePath] = old;
                    binding.pendingCount = Math.max(0, binding.pendingCount - 1);
                    continue;
                }
                if (ext && local) {
                    if (ext.hash === local.hash) {
                        next[relativePath] = {
                            externalHash: ext.hash,
                            vaultHash: local.hash,
                            externalMtimeMs: ext.mtimeMs,
                            externalSize: ext.size,
                            vaultMtime: local.mtime,
                            vaultSize: local.size,
                        };
                    } else {
                        const externalChanged = !old || old.externalHash !== ext.hash;
                        const vaultChanged = !old || old.vaultHash !== local.hash;
                        if (!externalChanged && !vaultChanged) {
                            next[relativePath] = {
                                externalHash: ext.hash,
                                vaultHash: local.hash,
                                externalMtimeMs: ext.mtimeMs,
                                externalSize: ext.size,
                                vaultMtime: local.mtime,
                                vaultSize: local.size,
                            };
                        } else if (externalChanged && vaultChanged) {
                            binding.conflicts.push(await this.makeConflict(binding, relativePath, ext.content, local.content, ticket));
                            binding.status = "conflict";
                            continue;
                        } else if (externalChanged) {
                            await this.stageVaultWrite(
                                binding,
                                relativePath,
                                ext.content,
                                ticket,
                                async () => this.samePair(ext, local, await this.currentPair(binding, relativePath)),
                            );
                            next[relativePath] = {
                                externalHash: ext.hash,
                                vaultHash: ext.hash,
                                externalMtimeMs: ext.mtimeMs,
                                externalSize: ext.size,
                                vaultMtime: local.file.stat?.mtime,
                                vaultSize: local.file.stat?.size,
                            };
                        } else {
                            await this.atomicWrite(
                                ext.absolutePath,
                                local.content,
                                () => this.isCurrent(binding, ticket),
                                async () => this.samePair(ext, local, await this.currentPair(binding, relativePath)),
                                () => this.authority.beginFinal(ticket),
                                { binding, relativePath },
                            );
                            next[relativePath] = {
                                externalHash: local.hash,
                                vaultHash: local.hash,
                                externalMtimeMs: ext.mtimeMs,
                                externalSize: ext.size,
                                vaultMtime: local.mtime,
                                vaultSize: local.size,
                            };
                        }
                    }
                } else if (ext && !local) {
                    if (old?.pendingDeletion === "vault") {
                        next[relativePath] = { ...old, externalHash: ext.hash };
                        protectedDeletes++;
                    } else if (
                        old?.externalExists === true
                        && old?.vaultExists === false
                        && old.externalHash === ext.hash
                    ) {
                        next[relativePath] = old;
                    } else if (old?.vaultHash) {
                        // Missing a mirror never authorizes moving the external source.
                        next[relativePath] = {
                            ...old,
                            externalHash: ext.hash,
                            vaultHash: "",
                            pendingDeletion: "vault",
                        };
                        protectedDeletes++;
                    } else {
                        await this.stageVaultWrite(
                            binding,
                            relativePath,
                            ext.content,
                            ticket,
                            async () => this.samePair(ext, local, await this.currentPair(binding, relativePath)),
                        );
                        next[relativePath] = {
                            externalHash: ext.hash,
                            vaultHash: ext.hash,
                            externalMtimeMs: ext.mtimeMs,
                            externalSize: ext.size,
                        };
                    }
                } else if (!ext && local) {
                    if (
                        old?.externalExists === false
                        && old?.vaultExists === true
                        && old.vaultHash === local.hash
                    ) {
                        next[relativePath] = old;
                    } else if (old?.protectedDelete === "external") {
                        next[relativePath] = { ...old, vaultHash: local.hash };
                        protectedDeletes++;
                    } else if (old?.externalHash) {
                        await this.createRecovery(binding, relativePath, local.content, "vault");
                        next[relativePath] = {
                            externalHash: "",
                            vaultHash: local.hash,
                            protectedDelete: "external",
                        };
                        protectedDeletes++;
                    } else {
                        // 落实 P1-DEC-05: 知识库新建无基线 Markdown 必须挂起为 pendingExport，严禁静默外写
                        // 但若该文件已被规则排除（!FilterMatcher.isFileAccepted），则是 UNTRACKED_LOCAL，严禁标记 pendingExport！
                        if (FilterMatcher.isFileAccepted(relativePath, binding.filterConfig)) {
                            next[relativePath] = {
                                externalHash: "",
                                vaultHash: local.hash,
                                pendingExport: true,
                                vaultMtime: local.mtime,
                                vaultSize: local.size,
                            };
                        }
                    }
                }
                if (next[relativePath] && next[relativePath] !== old) {
                    next[relativePath].lastDeepHashAt = Date.now();
                } else if (next[relativePath] === old && ext && local) {
                    next[relativePath].lastDeepHashAt = Date.now();
                }
                binding.pendingCount = Math.max(0, binding.pendingCount - 1);
            }
            if (!this.isCurrent(binding, ticket)) return;
            binding.fileStates = next;
            binding.lastSuccessAt = Date.now();
            binding.lastError = protectedDeletes
                ? `${protectedDeletes} 项外部删除已停在恢复保护边界，知识库与 OneDrive 内容均未删除。`
                : "";
            binding.status = (binding.conflicts || []).length ? "conflict" : "active";
            await (this.store as any).save();
        } catch (error: any) {
            if (!this.isCurrent(binding, ticket)) return;
            // Any failed source identity / read / sync pass must revoke the
            // current run generation. In particular, a disconnected/replaced
            // source must not auto-sync when the same path reappears.
            // Do not await stop here: an in-flight final commit is released
            // only by this method's finally block.
            void this.authority.stop(binding);
            const stat = await this.fs.promises.stat(binding.externalPath).catch(() => null);
            binding.status = stat?.isDirectory() ? "failed" : "offline";
            binding.lastError = stat?.isDirectory()
                ? `同步失败，原版本已保留：${String(error.message || error)}`
                : "外部文件夹不可访问。已暂停同步，不会将离线解释为删除。";
            await (this.store as any).save();
            new Notice(`${binding.nickname} 已暂停：${binding.lastError}`, 8000);
        } finally {
            try {
                if (this.authority.isFinalizing(ticket) && this.bindings().includes(binding)) {
                    binding.fileStates = { ...(binding.fileStates || {}), ...partial };
                    await (this.store as any).save();
                }
            } finally {
                this.running.delete(binding.id);
                this.authority.finishFinal(ticket);
            }
        }
    }

    private async reconcileBaseline(binding: any): Promise<void> {
        // Never prune a Vault mirror while the source directory identity is
        // unknown, even if a filter rule changed.
        await this.assertBindingRoot(binding);
        const config = binding.filterConfig;
        const currentFilterHash = config?.filterHash || (config ? FilterMatcher.computeFilterHash(config) : "");
        const targetPrefix = `${normalizePath(binding.targetPath)}/`;
        let prunedCount = 0;
        const affectedFolderPaths = new Set<string>();

        // 1. 扫描 Vault 中该 binding 目标目录下的所有 Markdown 文件
        const vaultFiles = typeof this.app.vault.getMarkdownFiles === "function"
            ? this.app.vault.getMarkdownFiles()
            : [];

        // A partial Vault index is not evidence that a conflict, deletion
        // request or unapproved export has vanished. Fail closed before
        // applying any destructive filter updates.
        if ((binding.conflicts || []).some((c: any) =>
            c.relativePath && !FilterMatcher.isFileAccepted(c.relativePath, config))) {
            throw new Error("筛选变更涉及尚未解决的冲突，保留双方版本等待处理");
        }
        for (const [relPath, state] of Object.entries(binding.fileStates || {})) {
            if (FilterMatcher.isFileAccepted(relPath, config)) continue;
            const tracked: any = state;
            if (tracked?.pendingDeletion || tracked?.protectedDelete || tracked?.pendingExport) {
                throw new Error("筛选变更涉及待确认的文件操作，保留旧基线等待处理");
            }
        }

        for (const file of vaultFiles) {
            if (!file.path.startsWith(targetPrefix)) continue;
            const relPath = file.path.slice(targetPrefix.length);

            // 判定是否不符合新筛选规则
            if (!FilterMatcher.isFileAccepted(relPath, config)) {
                // A path under the target is not proof it belongs to this
                // binding. Only a trusted two-sided baseline is prunable.
                const tracked = binding.fileStates?.[relPath];
                if (!tracked?.externalHash) continue;
                if (!tracked.vaultHash) {
                    throw new Error("缺少可信双端同步基线，筛选收敛已暂停并保留原稿");
                }
                // An unresolved conflict owns both versions; filter changes
                // cannot silently erase the conflict or its recovery copies.
                if (binding.conflicts?.some((c: any) => c.relativePath === relPath)) {
                    throw new Error("冲突尚未解决，筛选收敛已暂停并保留双方版本");
                }
                // 收集其所有父目录路径，供后续自底向上清理空目录
                const segments = relPath.split("/");
                let curDir = normalizePath(binding.targetPath);
                for (let i = 0; i < segments.length - 1; i++) {
                    curDir = `${curDir}/${segments[i]}`;
                    affectedFolderPaths.add(curDir);
                }

                // 2. 比对是否为本地未同步脏创作（Dirty Check）
                let isDirty = false;
                // Read actual current bytes, not a potentially stale cached
                // snapshot. Failure to read either side is NOT delete consent.
                const localContent = await this.app.vault.read(file);
                let extContent: string;
                try {
                    const extFilePath = this.path.join(binding.externalPath, ...relPath.split("/"));
                    extContent = await this.fs.promises.readFile(extFilePath, "utf8");
                } catch {
                    throw new Error("外部原文无法可信读取，筛选收敛已停止并保留知识库镜像");
                }
                isDirty = localContent !== extContent;

                if (isDirty) {
                    await this.createRecovery(binding, relPath, localContent, "vault");
                }
                if (typeof (this.app.vault as any).trash !== "function") {
                    throw new Error("宿主回收站不可用，已停止规则收敛");
                }
                await this.assertBindingRoot(binding);
                // The external file may change after the initial dirty check.
                // Verify both live sides again before disposing of a mirror.
                const latestExternal = await this.fs.promises.readFile(
                    this.path.join(binding.externalPath, ...relPath.split("/")), "utf8");
                if (latestExternal !== extContent) {
                    throw new Error("外部原文在筛选收敛前发生变化，保留知识库镜像");
                }
                if (binding.conflicts?.some((c: any) => c.relativePath === relPath) ||
                    await this.app.vault.read(file) !== localContent) {
                    throw new Error("镜像版本或冲突状态在移除前发生变化，安全停止");
                }
                await (this.app.vault as any).trash(file, false);
                prunedCount++;

                // 4. 从基线解绑
                if (binding.fileStates && binding.fileStates[relPath]) {
                    delete binding.fileStates[relPath];
                }

                // A filter update never destroys unresolved conflict records
                // or their independently recoverable snapshots.
            }
        }

        // 6. 自底向上清理空目录（Bottom-up Empty Folder Pruning）
        const sortedDirs = Array.from(affectedFolderPaths).sort((a, b) => b.length - a.length);
        const normalizedTarget = normalizePath(binding.targetPath);

        for (const dirPath of sortedDirs) {
            if (dirPath === normalizedTarget) continue; // 坚决不删除 targetPath 自身

            const abstractFolder = this.app.vault.getAbstractFileByPath(dirPath);
            if (!abstractFolder) continue;

            // 仅当目录彻底变空时才修剪（若含有非 md 资产则坚决保留）
            let isEmpty = false;
            if (abstractFolder instanceof TFolder) {
                isEmpty = abstractFolder.children.length === 0;
            } else {
                const allFiles = typeof this.app.vault.getFiles === "function" ? this.app.vault.getFiles() : [];
                isEmpty = !allFiles.some((f: any) => f.path.startsWith(`${dirPath}/`));
            }

            if (isEmpty) {
                if (typeof (this.app.vault as any).trash === "function") {
                    await (this.app.vault as any).trash(abstractFolder, false);
                } else if (typeof (this.app.vault as any).delete === "function") {
                    await (this.app.vault as any).delete(abstractFolder);
                }
            }
        }

        // 7. 处理之前不在 vault 中但记录在 fileStates 中的死项
        if (binding.fileStates) {
            for (const relativePath of Object.keys(binding.fileStates)) {
                if (!FilterMatcher.isFileAccepted(relativePath, config)) {
                    delete binding.fileStates[relativePath];
                }
            }
        }

        // 8. 更新已生效指纹并落盘
        binding.appliedFilterHash = currentFilterHash;
        await (this.store as any).save();

        // 9. 明确的用户反馈提示
        if (prunedCount > 0 && typeof Notice !== "undefined") {
            new Notice(`[SmartMD] 筛选规则已生效，已从知识库移出 ${prunedCount} 篇受排除文档`);
        }
    }

    async pause(id: string): Promise<void> {
        const binding = this.bindings().find((item) => item.id === id);
        if (!binding) return;
        const stopped = this.authority.stop(binding);
        binding.status = "stopping";
        await (this.store as any).save();
        await stopped;
        binding.status = "paused";
        await (this.store as any).save();
    }

    async resume(id: string): Promise<void> {
        const binding = this.bindings().find((item) => item.id === id);
        if (!binding) return;
        // A resume button is not an authorization grant for an unconfirmed
        // legacy binding or a source/target scope that changed.
        const approvedScope = `${binding.externalPath}|${binding.targetPath}`;
        if (binding.confirmed !== true || binding.confirmationScope !== approvedScope) {
            await this.authority.isolate(binding);
            binding.status = "paused";
            binding.lastError = "绑定未获可信授权或同步范围已改变；保持暂停。请重新核查并明确确认绑定范围，不能复用旧许可。";
            await (this.store as any).save();
            return;
        }
        try {
            await this.assertBindingRoot(binding);
            await this.scanExternal(binding.externalPath);
            const completedRoot = await this.fs.promises.stat(binding.externalPath);
            if (!completedRoot.isDirectory()) throw new Error("外部同步根目录不可用");
        } catch (error) {
            await this.authority.stop(binding);
            binding.status = "offline";
            binding.lastError = error instanceof Error
                ? error.message
                : "外部同步根目录不可用。完整重扫前不会恢复。";
            await (this.store as any).save();
            return;
        }
        this.authority.confirm(binding);
        binding.confirmationScope = `${binding.externalPath}|${binding.targetPath}`;
        binding.status = "active";
        binding.lastError = "";
        await (this.store as any).save();
        await this.syncBinding(binding);
    }

    async remove(id: string): Promise<void> {
        const binding = this.bindings().find((item) => item.id === id);
        if (!binding) return;
        const stopped = this.authority.stop(binding);
        binding.status = "stopping";
        await (this.store as any).save();
        await stopped;
        if ((binding.recoveryItems || []).length) {
            this.recoveryArchives().push({
                ...binding,
                id: uid("detached"),
                bindingId: binding.id,
                detached: true,
                detachedAt: Date.now(),
                status: "detached",
                confirmed: false,
                runAllowed: false,
                recoveryItems: [...binding.recoveryItems],
            });
        }
        const index = this.bindings().findIndex((item) => item.id === id);
        if (index >= 0) this.bindings().splice(index, 1);
        await (this.store as any).save();
        new Notice(`已解除“${binding.nickname}”的绑定。其他同步项未受影响，外部原文件夹和知识库副本均已保留`);
    }

    private validatedHiddenMarkdown(pathValue: string): string {
        const path = normalizePath(pathValue);
        if ((!path.startsWith(".suite-recovery/") && !path.startsWith(".suite-staging/")) ||
            !path.endsWith(".md") || path.split("/").some((part) =>
                !part || part === "." || part === ".." || part.includes("\\\\"))) {
            throw new Error("拒绝访问非同步范围的隐藏备份路径");
        }
        return path;
    }

    private async readHiddenMarkdown(pathValue: string): Promise<string> {
        const path = this.validatedHiddenMarkdown(pathValue);
        const indexed = this.app.vault.getAbstractFileByPath(path);
        if (indexed instanceof TFile) return this.app.vault.read(indexed);
        const adapter: any = this.app.vault.adapter;
        if (typeof adapter.read !== "function" || !(await adapter.exists(path))) {
            throw new Error("隐藏恢复正文缺失或不可读取");
        }
        return adapter.read(path);
    }

    private async createHiddenMarkdown(pathValue: string, content: string): Promise<void> {
        const path = this.validatedHiddenMarkdown(pathValue);
        await this.ensureVaultFolder(path);
        if (await this.app.vault.adapter.exists(path)) throw new Error("隐藏恢复目标已存在，不覆盖");
        await this.app.vault.create(path, content);
        if (await digest(await this.readHiddenMarkdown(path)) !== await digest(content)) {
            throw new Error("隐藏恢复内容物理读回失败，禁止正式覆盖");
        }
    }

    private async removeHiddenMarkdown(pathValue: string): Promise<void> {
        const path = this.validatedHiddenMarkdown(pathValue);
        const indexed = this.app.vault.getAbstractFileByPath(path);
        if (indexed instanceof TFile) {
            if (path.startsWith(".suite-recovery/deletions/") && this.app.fileManager?.trashFile) {
                await this.app.fileManager.trashFile(indexed);
            } else {
                await this.app.vault.delete(indexed, true);
            }
        } else if (await this.app.vault.adapter.exists(path)) {
            await this.app.vault.adapter.remove(path);
        }
        if (await this.app.vault.adapter.exists(path)) throw new Error("隐藏文件磁盘清理失败");
    }

    private async makeConflict(
        binding: any,
        relativePath: string,
        externalContent: string | null,
        vaultContent: string | null,
        ticket: RunTicket,
    ): Promise<any> {
        const conflictId = uid("conflict");
        const stagingRoot = normalizePath(`.suite-staging/conflicts/${binding.id}/${ticket.id}/${conflictId}`);
        const root = normalizePath(`.suite-recovery/conflicts/${binding.id}/${conflictId}`);
        const stagedExternal = stagingRoot + "/external.md";
        const stagedVault = stagingRoot + "/vault.md";
        await this.createHiddenMarkdown(stagedExternal, externalContent || "");
        await this.createHiddenMarkdown(stagedVault, vaultContent || "");
        if (!this.isCurrent(binding, ticket)) {
            throw new Error("同步已停止，冲突内容仅保留在隔离暂存区");
        }
        const fresh = await this.currentPair(binding, relativePath);
        if ((fresh.external?.content ?? null) !== externalContent || (fresh.vault?.content ?? null) !== vaultContent) {
            throw new Error("冲突内容已变化，需要重新确认");
        }
        this.authority.beginFinal(ticket);
        const created: string[] = [];
        try {
            const first = root + "/external.md";
            created.push(first);
            await this.createHiddenMarkdown(first, externalContent || "");
            const second = root + "/vault.md";
            created.push(second);
            await this.createHiddenMarkdown(second, vaultContent || "");
            if (!this.isCurrent(binding, ticket)) {
                throw new Error("同步已停止，未生成新的冲突副本");
            }
            await this.removeHiddenMarkdown(stagedExternal);
            await this.removeHiddenMarkdown(stagedVault);
        } catch (error) {
            for (const path of created) await this.removeHiddenMarkdown(path).catch(() => {});
            throw error;
        }
        return {
            id: conflictId,
            relativePath,
            externalCopy: `${root}/external.md`,
            vaultCopy: `${root}/vault.md`,
            externalExists: externalContent !== null,
            vaultExists: vaultContent !== null,
            externalHash: externalContent === null ? "" : await digest(externalContent),
            vaultHash: vaultContent === null ? "" : await digest(vaultContent),
            createdAt: Date.now(),
        };
    }

    private conflictMatches(conflict: any, pair: {
        external: ScannedFile | null;
        vault: { file: TFile; content: string; hash: string } | null;
    }): boolean {
        const expectedExternalExists = conflict.externalExists !== false;
        const expectedVaultExists = conflict.vaultExists !== false;
        return Boolean(pair.external) === expectedExternalExists
            && Boolean(pair.vault) === expectedVaultExists
            && (!pair.external || pair.external.hash === conflict.externalHash)
            && (!pair.vault || pair.vault.hash === conflict.vaultHash);
    }

    private async cleanupConflictCopies(conflict: any): Promise<void> {
        for (const path of [conflict.externalCopy, conflict.vaultCopy]) {
            await this.removeHiddenMarkdown(path);
        }
    }

    private siblingPath(relativePath: string, label: string): string {
        const slash = relativePath.lastIndexOf("/");
        const dot = relativePath.lastIndexOf(".");
        const hasExtension = dot > slash;
        const stem = hasExtension ? relativePath.slice(0, dot) : relativePath;
        const extension = hasExtension ? relativePath.slice(dot) : ".md";
        return `${stem} (${label})${extension}`;
    }

    private async keepBoth(
        binding: any,
        conflict: any,
        externalContent: string,
        vaultContent: string,
        ticket: RunTicket,
    ): Promise<void> {
        const stagingRoot = normalizePath(
            `.suite-staging/conflicts/${binding.id}/${ticket.id}/${conflict.id}/keep-both`,
        );
        const stagedForVault = stagingRoot + "/to-vault.md";
        const stagedForExternal = stagingRoot + "/to-external.md";
        // Preflight failures (name collisions, stale decisions) and later
        // partial commits must not leave untracked hidden staging copies.
        try {
        await this.createHiddenMarkdown(stagedForVault, externalContent);
        await this.createHiddenMarkdown(stagedForExternal, vaultContent);
        const vaultRelativePath = this.siblingPath(conflict.relativePath, "外部版本");
        const externalRelativePath = this.siblingPath(conflict.relativePath, "知识库版本");
        const vaultVisiblePath = normalizePath(`${binding.targetPath}/${vaultRelativePath}`);
        const externalVisiblePath = this.path.join(
            binding.externalPath,
            ...externalRelativePath.split("/"),
        );
        const stillBound = async (): Promise<boolean> => {
            if (!this.isCurrent(binding, ticket)) return false;
            const pair = await this.currentPair(binding, conflict.relativePath);
            return pair.external?.hash === conflict.externalHash
                && pair.vault?.hash === conflict.vaultHash;
        };
        if (!await stillBound()) {
            throw new Error("冲突内容或运行许可已变化，保留双方未执行");
        }
        if (
            this.app.vault.getAbstractFileByPath(vaultVisiblePath)
            || await this.fs.promises.stat(externalVisiblePath).then(() => true).catch(() => false)
        ) {
            throw new Error("保留双方的目标副本已存在，未覆盖任何内容");
        }

        binding.status = "syncing";
        binding.lastError = "正在完成“保留双方”";
        await (this.store as any).save();
        this.authority.beginFinal(ticket);
        let createdVault: TFile | null = null;
        let createdExternal = false;
        const previousConflicts = [...(binding.conflicts || [])];
        const previousFileState = binding.fileStates?.[conflict.relativePath];
        try {
            if (!await stillBound()) throw new Error("最终落地前状态已变化");
            await this.ensureVaultFolder(vaultVisiblePath);
            createdVault = await this.app.vault.create(
                vaultVisiblePath,
                await this.readHiddenMarkdown(stagedForVault),
            );
            if (!await stillBound()) throw new Error("第一份副本完成后状态已变化");
            await this.fs.promises.mkdir(this.path.dirname(externalVisiblePath), { recursive: true });
            await this.fs.promises.writeFile(
                externalVisiblePath,
                await this.readHiddenMarkdown(stagedForExternal),
                { encoding: "utf8", flag: "wx" },
            );
            createdExternal = true;
            if (!await stillBound()) throw new Error("第二份副本完成后状态已变化");

            binding.conflicts = previousConflicts.filter((item: any) => item.id !== conflict.id);
            binding.fileStates = binding.fileStates || {};
            binding.fileStates[conflict.relativePath] = {
                externalHash: conflict.externalHash,
                vaultHash: conflict.vaultHash,
            };
            binding.status = binding.conflicts.length ? "conflict" : "active";
            binding.lastError = "";
            await (this.store as any).save();

            await this.removeHiddenMarkdown(stagedForVault);
            await this.removeHiddenMarkdown(stagedForExternal);
            await this.cleanupConflictCopies(conflict);
        } catch (error) {
            if (createdExternal) {
                await this.fs.promises.unlink(externalVisiblePath).catch(() => {});
            }
            if (createdVault) {
                await this.app.fileManager.trashFile(createdVault).catch(() => {});
            }
            binding.conflicts = previousConflicts;
            binding.fileStates = binding.fileStates || {};
            if (previousFileState) {
                binding.fileStates[conflict.relativePath] = previousFileState;
            } else {
                delete binding.fileStates[conflict.relativePath];
            }
            if (this.isCurrent(binding, ticket)) {
                binding.status = "conflict";
                binding.lastError = error instanceof Error ? error.message : "保留双方未完成";
                await (this.store as any).save();
            }
            throw error;
        }
        } finally {
            let stagingCleanupError: unknown = null;
            for (const stagePath of [stagedForVault, stagedForExternal]) {
                try {
                    await this.removeHiddenMarkdown(stagePath);
                } catch (error) {
                    if (!stagingCleanupError) stagingCleanupError = error;
                }
            }
            if (stagingCleanupError) {
                throw new Error("保留双方的隐藏暂存未能彻底清理：" + String(stagingCleanupError));
            }
        }
    }

    async resolveConflict(bindingId: string, conflictId: string, choice: "external" | "vault" | "both"): Promise<void> {
        const existing = this.syncJobs.get(bindingId);
        if (existing) return await existing;
        const job = this.performResolveConflict(bindingId, conflictId, choice);
        this.syncJobs.set(bindingId, job);
        try {
            await job;
        } finally {
            if (this.syncJobs.get(bindingId) === job) this.syncJobs.delete(bindingId);
        }
    }

    private async performResolveConflict(bindingId: string, conflictId: string, choice: "external" | "vault" | "both"): Promise<void> {
        const binding = this.bindings().find((item) => item.id === bindingId);
        const conflict = binding?.conflicts?.find((item: any) => item.id === conflictId);
        if (!binding || !conflict || !this.canRun(binding)) return;
        const ticket = this.authority.start(binding, `${binding.externalPath}|${binding.targetPath}`);
        try {
        const externalContent = await this.readHiddenMarkdown(conflict.externalCopy);
        const vaultContent = await this.readHiddenMarkdown(conflict.vaultCopy);
        const fresh = await this.currentPair(binding, conflict.relativePath);
        this.authority.require(ticket);
        if (!this.conflictMatches(conflict, fresh)) {
            const replacement = await this.makeConflict(
                binding,
                conflict.relativePath,
                fresh.external?.content ?? null,
                fresh.vault?.content ?? null,
                ticket,
            );
            await this.cleanupConflictCopies(conflict);
            Object.assign(conflict, replacement);
            binding.status = "conflict";
            await (this.store as any).save();
            new Notice("冲突出现后内容再次变化，原选择已失效，请重新确认。", 8000);
            return;
        }
        const externalExists = conflict.externalExists !== false;
        const vaultExists = conflict.vaultExists !== false;
        if (!externalExists || !vaultExists) {
            if (!externalExists && !vaultExists) {
                delete binding.fileStates[conflict.relativePath];
            } else if (choice === "both") {
                binding.fileStates[conflict.relativePath] = {
                    externalExists,
                    vaultExists,
                    externalHash: conflict.externalHash,
                    vaultHash: conflict.vaultHash,
                };
            } else if (choice === "external" && !externalExists) {
                if (fresh.vault) {
                    await this.createRecovery(binding, conflict.relativePath, fresh.vault.content, "vault");
                    this.authority.require(ticket);
                    if (!this.conflictMatches(conflict, await this.currentPair(binding, conflict.relativePath))) {
                        throw new Error("删除版本在最终落地前再次变化，旧选择已失效");
                    }
                    this.authority.beginFinal(ticket);
                    await this.app.fileManager.trashFile(fresh.vault.file);
                }
                delete binding.fileStates[conflict.relativePath];
            } else if (choice === "vault" && !vaultExists) {
                if (fresh.external) {
                    const recovery = await this.createRecovery(binding, conflict.relativePath, fresh.external.content, "external");
                    recovery.externalBackupPath = await this.stageExternalDelete(
                        binding,
                        conflict.relativePath,
                        fresh.external.absolutePath,
                        () => this.isCurrent(binding, ticket),
                        () => this.authority.beginFinal(ticket),
                    );
                }
                delete binding.fileStates[conflict.relativePath];
            } else {
                const selectedContent = choice === "external"
                    ? fresh.external?.content
                    : fresh.vault?.content;
                if (selectedContent === undefined) throw new Error("所选版本已删除，请重新选择");
                if (!fresh.vault) await this.stageVaultWrite(binding, conflict.relativePath, selectedContent, ticket);
                if (!fresh.external) {
                    await this.atomicWrite(
                        this.path.join(binding.externalPath, ...conflict.relativePath.split("/")),
                        selectedContent,
                        () => this.isCurrent(binding, ticket),
                        async () => this.conflictMatches(conflict, await this.currentPair(binding, conflict.relativePath)),
                        () => this.authority.beginFinal(ticket),
                    );
                }
                const selectedHash = await digest(selectedContent);
                binding.fileStates[conflict.relativePath] = {
                    externalHash: selectedHash,
                    vaultHash: selectedHash,
                };
            }
            binding.conflicts = binding.conflicts.filter((item: any) => item.id !== conflict.id);
            binding.status = binding.conflicts.length ? "conflict" : "active";
            binding.lastError = "";
            await this.cleanupConflictCopies(conflict);
            await (this.store as any).save();
            return;
        }
        if (choice === "both") {
            await this.keepBoth(
                binding,
                conflict,
                externalContent,
                vaultContent,
                ticket,
            );
            return;
        }
        const vaultPath = normalizePath(`${binding.targetPath}/${conflict.relativePath}`);
        const externalPath = this.path.join(binding.externalPath, ...conflict.relativePath.split("/"));
        if (choice === "external") {
            await this.stageVaultWrite(
                binding,
                conflict.relativePath,
                externalContent,
                ticket,
                async () => {
                    const current = await this.currentPair(binding, conflict.relativePath);
                    return current.external?.hash === conflict.externalHash && current.vault?.hash === conflict.vaultHash;
                },
            );
        } else if (choice === "vault") {
            await this.atomicWrite(
                externalPath,
                vaultContent,
                () => this.isCurrent(binding, ticket),
                async () => {
                    const current = await this.currentPair(binding, conflict.relativePath);
                    return current.external?.hash === conflict.externalHash && current.vault?.hash === conflict.vaultHash;
                },
                () => this.authority.beginFinal(ticket),
                { binding, relativePath: conflict.relativePath },
            );
        } else {
            const copyPath = vaultPath.replace(/\.md$/i, `.external-${Date.now()}.md`);
            await this.ensureVaultFolder(copyPath);
            this.authority.require(ticket);
            await this.app.vault.create(copyPath, externalContent);
            await this.atomicWrite(
                externalPath.replace(/\.md$/i, `.vault-${Date.now()}.md`),
                vaultContent,
                () => this.isCurrent(binding, ticket),
                async () => true,
                () => this.authority.beginFinal(ticket),
            );
        }
        if (!this.authority.isFinalizing(ticket)) this.authority.require(ticket);
        binding.conflicts = binding.conflicts.filter((item: any) => item.relativePath !== conflict.relativePath);
        const resolvedHash = await digest(choice === "external" ? externalContent : vaultContent);
        binding.fileStates = binding.fileStates || {};
        binding.fileStates[conflict.relativePath] = { externalHash: resolvedHash, vaultHash: resolvedHash };
        binding.status = binding.conflicts.length ? "conflict" : "active";
        binding.lastError = "";
        await (this.store as any).save();
        } finally {
            this.authority.finishFinal(ticket);
        }
    }

    private async createRecovery(binding: any, relativePath: string, content: string, source: "external" | "vault"): Promise<any> {
        const recoveryItems = (binding.recoveryItems || []) as any[];
        const backupBytes = recoveryItems.reduce((total: number, item: any) =>
            total + Number(item.contentBytes || 0), 0);
        const contentBytes = new TextEncoder().encode(content).byteLength;
        if (recoveryItems.length >= 5000 || backupBytes + contentBytes > 512 * 1024 * 1024) {
            throw new Error("恢复档案已达容量上限，请核验导出后再继续写入");
        }
        const recoveryId = uid("recovery");
        const storagePath = normalizePath(`.suite-recovery/deletions/${binding.id}/${recoveryId}.md`);
        const deletedAt = Date.now();
        await this.createHiddenMarkdown(storagePath, content);
        // Retention begins only after the hidden recovery bytes are verified.
        const verifiedAt = Date.now();
        const item = {
            id: recoveryId,
            relativePath,
            source,
            deletedAt,
            expiresAt: verifiedAt + 7 * 24 * 60 * 60 * 1000,
            storagePath,
            verifiedAt,
            contentHash: await digest(content),
            contentBytes,
            resolvedAt: 0,
        };
        binding.recoveryItems.push(item);
        await (this.store as any).save();
        return item;
    }

    async restore(bindingId: string, recoveryId: string): Promise<void> {
        const binding = this.knownBindings().find((item) =>
            item.id === bindingId || item.bindingId === bindingId);
        const item = binding?.recoveryItems?.find((entry: any) => entry.id === recoveryId);
        if (!binding || !item) return;
        // Approval for one recovery file must never renew an unconfirmed,
        // paused, detached or scope-changed binding's ongoing sync license.
        if (!this.canRun(binding)) {
            binding.lastError = "绑定没有有效的当前同步授权，已保留恢复档案；请重新确认绑定范围。";
            await (this.store as any).save();
            throw new Error(binding.lastError);
        }
        const canonical = await this.canonicalPath(binding.externalPath).catch(() => "");
        const expectedCanonical = binding.canonicalExternalPath || binding.externalPath;
        if (!canonical || this.pathKey(canonical) !== this.pathKey(expectedCanonical)) {
            binding.lastError = "原外部路径当前无法确认，恢复记录已保留待重试。";
            await (this.store as any).save();
            throw new Error(binding.lastError);
        }
        const ownerId = String(binding.bindingId || binding.id || "");
        const itemId = String(item.id || "");
        const expectedPath = ".suite-recovery/deletions/" + ownerId + "/" + itemId + ".md";
        // Keep old records without a verifiable hash; do not restore their bytes.
        if (!item.contentHash || !/^[a-f0-9]{64}$/i.test(String(item.contentHash))) {
            throw new Error("旧恢复记录缺少可信SHA-256，保留原档案并拒绝危险恢复");
        }
        if (!/^[a-zA-Z0-9_-]+$/.test(ownerId) ||
            !/^[a-zA-Z0-9_-]+$/.test(itemId) ||
            item.storagePath !== expectedPath) throw new Error("恢复路径身份不可信");
        const content = await this.readHiddenMarkdown(item.storagePath);
        const storedHash = await digest(content);
        if (item.contentHash !== storedHash) throw new Error("恢复档案哈希校验失败");
        const current = await this.currentPair(binding, item.relativePath);
        if (
            (current.external && current.external.hash !== storedHash)
            || (current.vault && current.vault.hash !== storedHash)
        ) {
            binding.lastError = "恢复目标已有其他版本，未覆盖并保留待恢复记录。";
            await (this.store as any).save();
            throw new Error(binding.lastError);
        }
        const approved = typeof window !== "undefined" && typeof window.confirm === "function"
            && window.confirm(`确认将“${item.relativePath}”恢复到外部文件夹和知识库副本吗？`);
        if (!approved) return;
        // The above confirmation applies only to this recovery request.
        // It never creates or extends the binding's persistent run license.
        if (!this.canRun(binding)) throw new Error("恢复操作授权已失效，未写入任何目标");
        // Freeze BOTH targets across the confirmation dialog. In particular,
        // an external file created while the user reviews this dialog must
        // never be overwritten by an old recovery preview.
        const approvedExternal = current.external || undefined;
        let approvedVault = current.vault || undefined;
        const validateRecoverySnapshot = async (): Promise<boolean> => {
            const latest = await this.currentPair(binding, item.relativePath);
            if (!this.samePair(approvedExternal, approvedVault, latest)) return false;
            // currentPair's Vault cache can be stale: verify actual bytes too.
            if (latest.vault &&
                await digest(await this.app.vault.read(latest.vault.file)) !== approvedVault?.hash) return false;
            return true;
        };
        if (!(await validateRecoverySnapshot())) {
            throw new Error("恢复目标在用户确认期间发生变化，已拒绝覆盖并保留恢复档案");
        }
        const ticket = this.authority.start(binding);
        try {
            if (!current.vault) {
                await this.stageVaultWrite(binding, item.relativePath, content, ticket, validateRecoverySnapshot);
                const afterVaultWrite = await this.currentPair(binding, item.relativePath);
                if (!afterVaultWrite.vault || afterVaultWrite.vault.hash !== storedHash ||
                    (approvedExternal ? afterVaultWrite.external?.hash !== approvedExternal.hash
                        : afterVaultWrite.external !== null) ||
                    await digest(await this.app.vault.read(afterVaultWrite.vault.file)) !== storedHash) {
                    throw new Error("恢复中目标版本发生变化，已保留恢复档案");
                }
                approvedVault = afterVaultWrite.vault;
            }
            if (!current.external) {
                await this.atomicWrite(
                    this.path.join(binding.externalPath, ...item.relativePath.split("/")),
                    content,
                    () => this.isCurrent(binding, ticket),
                    validateRecoverySnapshot,
                    () => this.authority.beginFinal(ticket),
                );
            }
            binding.fileStates = binding.fileStates || {};
            binding.fileStates[item.relativePath] = { externalHash: storedHash, vaultHash: storedHash };
            const verified = await this.currentPair(binding, item.relativePath);
            if (!verified.vault || !verified.external ||
                verified.vault.hash !== storedHash || verified.external.hash !== storedHash) {
                throw new Error("恢复后读回不一致；已保留恢复记录");
            }
            item.restoredAt = Date.now();
            binding.lastError = "";
            await (this.store as any).save();
        } finally {
            this.authority.finishFinal(ticket);
        }
    }

    async exportRecoveryToNewNote(bindingId: string, recoveryId: string): Promise<string> {
        const binding = this.knownBindings().find((b: any) => b.id === bindingId || b.bindingId === bindingId);
        const item = binding?.recoveryItems?.find((r: any) => r.id === recoveryId);
        if (!item) throw new Error("恢复记录不存在");
        const ownerId = String(binding.bindingId || binding.id || "");
        const itemId = String(item.id || "");
        const expectedPath = ".suite-recovery/deletions/" + ownerId + "/" + itemId + ".md";
        if (!/^[a-zA-Z0-9_-]+$/.test(ownerId) ||
            !/^[a-zA-Z0-9_-]+$/.test(itemId) ||
            item.storagePath !== expectedPath) throw new Error("恢复路径身份不可信");
        const content = await this.readHiddenMarkdown(item.storagePath);
        if (!item.contentHash || await digest(content) !== item.contentHash) {
            throw new Error("恢复文件哈希不符，拒绝导出");
        }
        const folder = normalizePath("恢复的外部笔记");
        const safePart = safeName(this.path.basename(item.relativePath)).replace(/\.md$/i, "");
        const output = normalizePath(folder + "/" + safePart + "-" + item.id + ".md");
        if (this.app.vault.getAbstractFileByPath(output)) {
            throw new Error("恢复目标已存在，不覆盖当前内容");
        }
        await this.ensureVaultFolder(output);
        const created = await this.app.vault.create(output, content);
        if (await digest(await this.app.vault.read(created)) !== item.contentHash) {
            throw new Error("导出文件验证失败，恢复档案继续保留");
        }
        item.exportedAt = Date.now();
        item.exportPath = output;
        await (this.store as any).save();
        return output;
    }

    async confirmRecoveryResolved(bindingId: string, recoveryId: string): Promise<void> {
        const binding = this.knownBindings().find((entry: any) =>
            entry.id === bindingId || entry.bindingId === bindingId);
        const item = binding?.recoveryItems?.find((entry: any) => entry.id === recoveryId);
        if (!item || !(item.restoredAt || item.exportedAt) || !item.verifiedAt) {
            throw new Error("恢复或独立导出尚未完成并验证，不能标记已解决");
        }
        item.resolvedAt = Date.now();
        await (this.store as any).save();
    }

    async protectRemoteDelete(vaultPath: string, readOnly = false): Promise<boolean> {
        for (const binding of this.bindings()) {
            const prefix = `${normalizePath(binding.targetPath)}/`;
            if (!normalizePath(vaultPath).startsWith(prefix)) continue;
            const relPath = vaultPath.slice(prefix.length);
            const file = this.app.vault.getAbstractFileByPath(vaultPath);
            if (!(file instanceof TFile) || file.extension.toLowerCase() !== "md") continue;
            // A corrupt or unsafe persisted include rule must not turn off
            // external-folder protection for an in-scope Vault document.
            if (binding.filterConfig && FilterMatcher.validateConfig(binding.filterConfig).length) {
                if (readOnly) return true;
                await this.authority.pause(binding);
                binding.status = "paused";
                binding.lastError = "筛选规则不安全，已暂停并阻止跨域删除；请重新确认规则。";
                await (this.store as any).save();
                return true;
            }
            if (!FilterMatcher.isFileAccepted(relPath, binding.filterConfig)) continue;
            if (readOnly) return true;
            // Revoke external execution before any fallible recovery write.
            // A failed backup must not leave the binding actively syncing.
            await this.authority.pause(binding);
            binding.status = "paused";
            binding.lastError = "收到其他设备的删除请求，已暂停并等待恢复副本核验。";
            await (this.store as any).save();
            try {
                await this.createRecovery(binding, relPath,
                    await this.app.vault.cachedRead(file), "vault");
            } catch (error) {
                binding.lastError = "跨模块删除保护已暂停，但恢复副本保存失败；请核查后重新确认。";
                try { await (this.store as any).save(); } catch {}
                throw error;
            }
            binding.lastError = "收到其他设备的删除请求，已停在恢复区等待确认。";
            await (this.store as any).save();
            return true;
        }
        return false;
    }

    markLegacyBindingsPaused(): Promise<void> {
        for (const binding of this.bindings()) {
            if (binding.confirmed !== true) {
                this.authority.isolate(binding);
                binding.status = "paused";
            }
        }
        return (this.store as any).save();
    }

    dispose(): void {
        void this.authority.shutdown();
    }

    private isCurrent(binding: any, ticket: RunTicket): boolean {
        return this.authority.isCurrent(ticket)
            && ticket.state === binding
            && this.canRun(binding)
            && this.knownBindings().includes(binding);
    }

    private canRun(binding: any): boolean {
        return this.authority.canRun(binding)
            && binding.confirmationScope === `${binding.externalPath}|${binding.targetPath}`;
    }

    private async assertSafeExternalRelativePath(binding: any, relativePath: string): Promise<void> {
        const parts = relativePath.replace(/\\/g, "/").split("/");
        if (!parts.length || parts.some((part) =>
            !part || part === "." || part === ".." || part.includes(":")) ||
            !relativePath.toLowerCase().endsWith(".md")) {
            throw new Error("外部相对路径不安全，已拒绝读写");
        }
        await this.assertBindingRoot(binding);
        const resolved = this.path.resolve(binding.externalPath, ...parts);
        if (!inside(binding.externalPath, resolved, this.path) ||
            this.pathKey(resolved) === this.pathKey(binding.externalPath)) {
            throw new Error("外部路径越过授权目录边界");
        }
        let current = binding.externalPath;
        for (let i = 0; i < parts.length; i++) {
            current = this.path.join(current, parts[i]);
            const stat = await this.fs.promises.lstat(current).catch((err: any) => {
                if (err?.code === "ENOENT") return null;
                throw err;
            });
            if (stat?.isSymbolicLink() || (i < parts.length - 1 && stat && !stat.isDirectory())) {
                throw new Error("外部路径包含链接或意外文件类型，拒绝跨目录操作");
            }
        }
    }

    private async currentPair(binding: any, relativePath: string): Promise<{
        external: ScannedFile | null;
        vault: { file: TFile; content: string; hash: string } | null;
    }> {
        const root = await this.fs.promises.stat(binding.externalPath).catch(() => null);
        if (!root?.isDirectory()) throw new Error("外部同步根目录不可访问");
        await this.assertSafeExternalRelativePath(binding, relativePath);
        const absolutePath = this.path.join(binding.externalPath, ...relativePath.split("/"));
        const externalContent = await this.fs.promises.readFile(absolutePath, "utf8").catch((error: any) => {
            if (error?.code === "ENOENT") return null;
            throw error;
        });
        const external = externalContent === null ? null : {
            relativePath,
            absolutePath,
            content: externalContent,
            hash: await digest(externalContent),
        };
        const vaultPath = normalizePath(`${binding.targetPath}/${relativePath}`);
        const file = this.app.vault.getAbstractFileByPath(vaultPath);
        const vault = file instanceof TFile
            ? await this.app.vault.cachedRead(file).then(async (content: string) => ({
                file,
                content,
                hash: await digest(content),
            }))
            : null;
        return { external, vault };
    }

    private samePair(
        expectedExternal: ScannedFile | undefined,
        expectedVault: { file: TFile; content: string; hash: string } | undefined,
        current: {
            external: ScannedFile | null;
            vault: { file: TFile; content: string; hash: string } | null;
        },
    ): boolean {
        return (expectedExternal ? current.external?.hash === expectedExternal.hash : current.external === null)
            && (expectedVault ? current.vault?.hash === expectedVault.hash : current.vault === null);
    }

    private async purgeExpiredRecovery(): Promise<void> {
        for (const binding of this.knownBindings()) {
            const retained: any[] = [];
            for (const item of binding.recoveryItems || []) {
                // Old records may have an expiry earlier than verified readback t0.
                const verifiedStart = Number(item.verifiedAt || item.deletedAt || 0);
                const expiresAt = Math.max(
                    Number(item.expiresAt || 0), verifiedStart + 7 * 24 * 60 * 60 * 1000,
                );
                const now = Date.now();
                const sixDays = 6 * 24 * 60 * 60 * 1000;
                const month = 30 * 24 * 60 * 60 * 1000;
                if (Number(item.deletedAt || 0) > 0 &&
                    now >= Number(item.deletedAt) + sixDays &&
                    (!item.lastReminderAt ||
                        (now >= expiresAt && now - Number(item.lastReminderAt) >= month))) {
                    item.lastReminderAt = now;
                    new Notice("外部同步恢复档案待处理：" + String(item.relativePath || "未知文件") +
                        "。未确认解决前不会自动销毁，请在恢复区核查。", 9000);
                }
                const resolvedAt = Number(item.resolvedAt || 0);
                const bufferUntil = resolvedAt + 7 * 24 * 60 * 60 * 1000;
                if (!resolvedAt || !item.verifiedAt || !(item.restoredAt || item.exportedAt) ||
                    expiresAt > Date.now() || bufferUntil > Date.now() ||
                    Date.now() < Number(item.deletedAt || 0)) {
                    retained.push(item);
                    continue;
                }
                try {
                    const ownerId = String(binding.bindingId || binding.id || "");
                    const recoveryId = String(item.id || "");
                    const safeVaultPrefix = ".suite-recovery/deletions/" + ownerId + "/";
                    if (!/^[a-zA-Z0-9_-]+$/.test(ownerId) ||
                        !/^[a-zA-Z0-9_-]+$/.test(recoveryId) ||
                        item.storagePath !== safeVaultPrefix + recoveryId + ".md" ||
                        !item.contentHash || !/^[a-f0-9]{64}$/i.test(String(item.contentHash))) {
                        throw new Error("恢复记录路径或可信哈希缺失，保留档案禁止自动清理");
                    }
                    // Verify actual hidden bytes before any physical backup removal.
                    const stored = await this.readHiddenMarkdown(item.storagePath);
                    if (await digest(stored) !== item.contentHash) {
                        throw new Error("隐藏恢复原文哈希不符，保留档案");
                    }
                    // Old "exported" or "restored" flags are not evidence that
                    // a current, readable replacement copy still exists.
                    let safeCopyVerified = false;
                    const exportPath = String(item.exportPath || "");
                    if (item.exportedAt && exportPath.startsWith("恢复的外部笔记/") &&
                        !exportPath.includes("\\") &&
                        exportPath.split("/").every((part: string) => part && part !== "." && part !== "..")) {
                        const exported = this.app.vault.getAbstractFileByPath(exportPath);
                        if (exported instanceof TFile &&
                            await digest(await this.app.vault.read(exported)) === item.contentHash) {
                            safeCopyVerified = true;
                        }
                    }
                    if (!safeCopyVerified && item.restoredAt) {
                        const pair = await this.currentPair(binding, item.relativePath);
                        safeCopyVerified = pair.external?.hash === item.contentHash &&
                            pair.vault?.hash === item.contentHash;
                    }
                    if (!safeCopyVerified) {
                        throw new Error("已恢复/导出的有效副本无法重新读回，保留恢复档案");
                    }
                    if (item.externalBackupPath) {
                        const backupPath = this.path.resolve(item.externalBackupPath);
                        const parent = this.path.dirname(backupPath);
                        if (this.path.basename(parent).toLowerCase() !== ".smartkeep-recovery" ||
                            !inside(binding.externalPath, backupPath, this.path) ||
                            this.pathKey(backupPath) === this.pathKey(binding.externalPath)) {
                            throw new Error("外部备份目录身份不可信，拒绝清理");
                        }
                        const backupContent = await this.fs.promises.readFile(backupPath, "utf8");
                        if (await digest(backupContent) !== item.contentHash) {
                            throw new Error("外部恢复备份字节不符，保留档案");
                        }
                        await this.fs.promises.unlink(backupPath);
                    }
                    const indexedCopy = this.app.vault.getAbstractFileByPath(item.storagePath);
                    if (!(indexedCopy instanceof TFile) &&
                        !(await this.app.vault.adapter.exists(item.storagePath))) {
                        throw new Error("恢复原文件已缺失，保留恢复条目待核验");
                    }
                    await this.removeHiddenMarkdown(item.storagePath);
                } catch (error: any) {
                    item.cleanupError = `到期恢复内容清理失败，将自动重试：${String(error.message || error)}`;
                    binding.lastError = item.cleanupError;
                    retained.push(item);
                }
            }
            binding.recoveryItems = retained;
        }
        const local = (this.store as any).settings.local;
        local.externalRecoveryArchives = this.recoveryArchives().filter((item: any) => (item.recoveryItems || []).length);
        await (this.store as any).save();
    }

    private async atomicWrite(
        destination: string,
        content: string,
        permit: () => boolean = () => true,
        validate: () => Promise<boolean> = async () => true,
        enterFinal: () => void = () => {},
        backupContext?: { binding: any; relativePath: string },
    ): Promise<void> {
        await this.fs.promises.mkdir(this.path.dirname(destination), { recursive: true });
        const readOriginal = async (): Promise<string | null> => {
            try {
                return await this.fs.promises.readFile(destination, "utf8");
            } catch (error: any) {
                if (error?.code === "ENOENT") return null;
                throw error; // permissions failure is never "missing file"
            }
        };
        const original = await readOriginal();
        const temporary = destination + ".smartkeep-stage-" + uid("commit") + ".tmp";
        let temporaryMayBeOurs = false;
        try {
            // Exclusive create means EEXIST refers to someone else's file.
            // Other write errors may leave partial bytes that must be cleaned.
            try {
                await this.fs.promises.writeFile(temporary, content, { encoding: "utf8", flag: "wx" });
                temporaryMayBeOurs = true;
            } catch (error: any) {
                if (error?.code !== "EEXIST") temporaryMayBeOurs = true;
                throw error;
            }
            if (!permit() || !(await validate())) {
                throw new Error("最终复核未通过，暂存内容未进入正式文件");
            }
            if (await readOriginal() !== original) {
                throw new Error("目标文件在准备写入期间变化，拒绝覆盖");
            }
            if (original !== null) {
                const recoveryRoot = this.path.join(this.path.dirname(destination), ".smartkeep-recovery");
                await this.fs.promises.mkdir(recoveryRoot, { recursive: true });
                const backup = this.path.join(recoveryRoot, uid("original") + "-" +
                    this.path.basename(destination) + ".bak");
                await this.fs.promises.copyFile(
                    destination, backup, this.fs.constants.COPYFILE_EXCL,
                );
                const verified = await this.fs.promises.readFile(backup, "utf8");
                if (await digest(verified) !== await digest(original)) {
                    throw new Error("外部原文件备份校验失败，不执行正式覆盖");
                }
                if (backupContext) {
                    const record = await this.createRecovery(
                        backupContext.binding, backupContext.relativePath, original, "external",
                    );
                    record.externalBackupPath = backup;
                    await (this.store as any).save();
                }
            }
            if (!permit() || !(await validate()) || await readOriginal() !== original) {
                throw new Error("正式提交前原文件又发生变化，保留原件并撤销写入");
            }
            enterFinal();
            if (original !== null) {
                // Atomic replace on the same volume: original path always remains
                // visible as either the old complete bytes or the new complete bytes.
                await this.fs.promises.rename(temporary, destination);
            } else {
                // Atomic exclusive creation; no race with a new external file.
                // Unsupported filesystems fail closed instead of using unsafe copy.
                await this.fs.promises.link(temporary, destination);
            }
            if (await this.fs.promises.readFile(destination, "utf8") !== content) {
                throw new Error("提交后读回校验失败，保留备份等待人工核查");
            }
        } finally {
            if (temporaryMayBeOurs) {
                try {
                    await this.fs.promises.unlink(temporary);
                } catch (error: any) {
                    if (error?.code !== "ENOENT") {
                        throw new Error("外部临时文件清理失败：" + String(error?.message || error));
                    }
                }
            }
        }
    }

    private async stageVaultWrite(
        binding: any,
        relativePath: string,
        content: string,
        ticket: RunTicket,
        validate: () => Promise<boolean> = async () => true,
    ): Promise<void> {
        const stagingPath = normalizePath(`.suite-staging/external/${binding.id}/${ticket.id}/${uid("write")}.md`);
        await this.ensureVaultFolder(stagingPath);
        try {
            // Even a staging create that partially writes and then throws
            // must enter the physical cleanup finally block.
            await this.app.vault.create(stagingPath, content);
            this.authority.require(ticket);
            if (!this.knownBindings().includes(binding)) throw new Error("同步项已解除绑定，暂存内容未落地");
            if (!(await validate())) throw new Error("最终复核发现双方再次变化，暂存内容未落地");
            if (!this.isCurrent(binding, ticket)) throw new Error("同步身份或范围已变化，暂存内容未落地");
            this.authority.beginFinal(ticket);
            const vaultPath = normalizePath(`${binding.targetPath}/${relativePath}`);
            const current = this.app.vault.getAbstractFileByPath(vaultPath);
            if (current instanceof TFile) {
                // Every overwritten Vault version enters the verified recovery registry.
                const currentContent = await this.app.vault.read(current);
                await this.createRecovery(binding, relativePath, currentContent, "vault");
                if (await this.app.vault.read(current) !== currentContent) {
                    throw new Error("知识库原稿在备份之后发生变化，拒绝覆盖");
                }
                await this.app.vault.modify(current, content);
            } else {
                await this.ensureVaultFolder(vaultPath);
                await this.app.vault.create(vaultPath, content);
            }
            const landed = this.app.vault.getAbstractFileByPath(vaultPath);
            if (!(landed instanceof TFile) || await this.app.vault.read(landed) !== content) {
                throw new Error("知识库提交后读回版本不符，已保留恢复档案");
            }
        } finally {
            // 落实 BL-EXT-009 与终审要求：在 finally 块中强制底层物理销毁暂存文件
            await this.removeCommittedExternalStage(stagingPath);
        }
    }

    private async removeCommittedExternalStage(stagingPath: string): Promise<void> {
        const parts = stagingPath.split("/");
        if (
            parts.length !== 5
            || parts[0] !== ".suite-staging"
            || parts[1] !== "external"
            || !parts[4].endsWith(".md")
            || parts.some((part) => !part || part === "." || part === ".." || part.includes("\\"))
        ) {
            throw new Error("拒绝清理非本次外部同步暂存路径");
        }

        const adapter = this.app.vault.adapter;
        try {
            const indexed = this.app.vault.getAbstractFileByPath(stagingPath);
            if (indexed instanceof TFile) {
                await this.app.vault.delete(indexed, true);
            } else if (await adapter.exists(stagingPath)) {
                // DataAdapter can address hidden files which getAbstractFileByPath
                // cannot resolve in a real Obsidian vault.
                await adapter.remove(stagingPath);
            }
            if (await adapter.exists(stagingPath)) {
                throw new Error("存储适配器未实际移除暂存文件");
            }
        } catch (error: any) {
            // The official Markdown may already have been committed. Make the
            // partial outcome explicit and stop the binding from claiming
            // successful synchronization instead of silently leaking copies.
            throw new Error(`暂存文件清理失败，目标正文可能已写入，请核对版本：${String(error.message || error)}`);
        }
    }

    private async stageExternalDelete(
        binding: any,
        relativePath: string,
        source: string,
        permit: () => boolean,
        enterFinal: () => void = () => {},
        expectedHash?: string,
    ): Promise<string> {
        const recoveryRoot = this.path.join(binding.externalPath, ".smartkeep-recovery");
        await this.fs.promises.mkdir(recoveryRoot, { recursive: true });
        const target = this.path.join(
            recoveryRoot,
            `${Date.now()}-${relativePath.replace(/[\\/:*?"<>|]/g, "_")}.bak`,
        );
        if (!permit()) throw new Error("运行许可已撤销，删除未落地");
        await this.assertSafeExternalRelativePath(binding, relativePath);
        if (expectedHash && await digest(await this.fs.promises.readFile(source, "utf8")) !== expectedHash) {
            throw new Error("外部文件在删除确认后变化，拒绝按旧版本移动");
        }
        enterFinal();
        await this.fs.promises.rename(source, target);
        return target;
    }

    private async ensureVaultFolder(filePath: string): Promise<void> {
        const slash = normalizePath(filePath).lastIndexOf("/");
        if (slash < 0) return;
        const folder = filePath.slice(0, slash);
        const parts = normalizePath(folder).split("/");
        let current = "";
        for (const part of parts) {
            current = current ? `${current}/${part}` : part;
            if (!(await this.app.vault.adapter.exists(current))) await this.app.vault.createFolder(current);
        }
    }
}
