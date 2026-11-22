import { ExternalFilterConfig, FilterRuleItem } from "./types";

const IMMUTABLE_DIR_BLACKLIST = new Set([
    ".git",
    ".svn",
    ".hg",
    ".obsidian",
    ".smartkeep-recovery",
    ".suite-staging",
    ".suite-recovery",
]);

const IMMUTABLE_FILE_BLACKLIST = new Set([
    ".ds_store",
    "thumbs.db",
    "desktop.ini",
]);

const DEFAULT_BUILD_ARTIFACTS = new Set([
    "node_modules",
    "dist",
    "build",
    "target",
    "vendor",
    ".venv",
    "venv",
    ".cache",
    "bin",
    "obj",
    "out",
]);

const DEFAULT_IDE_CONFIGS = new Set([
    ".vscode",
    ".idea",
]);

/** 静态检查 ReDoS 危险回溯嵌套量词，如 (a+)+, (a*)*, (a+)* */
const DANGEROUS_REGEX_PATTERN = /(\([^)]*[+*]\)[+*])|(\[[^\]]*\][+*][+*])/;

/**
 * Fail-closed structural analysis for user regex filters (BL-EXT-015).
 * Only allow a bounded grammar: literal atoms, character classes, one
 * unquantified literal-alternative group, and controlled token quantifiers.
 * Backtracking repeats must be disjoint or separated by a proven delimiter.
 * This avoids guessing safety from regex text fragments alone.
 */
function hasSafeRegexStructure(pattern: string): boolean {
    type Atom = { source: string; literal?: string; group?: boolean; min: number; max: number };
    const atoms: Atom[] = [];
    let i = 0;
    let groups = 0;
    while (i < pattern.length) {
        const ch = pattern[i];
        if (ch === "^" && i === 0) { i++; continue; }
        if (ch === "$" && i === pattern.length - 1) { i++; continue; }
        let atom: Atom;
        if (ch === "(") {
            if (++groups > 1) return false;
            const close = pattern.indexOf(")", i + 1);
            if (close < 0) return false;
            const body = pattern.slice(i + 1, close);
            const branches = body.split("|");
            if (!branches.length || branches.length > 4 ||
                branches.some((part) => !/^[a-zA-Z0-9_-]{1,20}$/.test(part))) return false;
            atom = { source: pattern.slice(i, close + 1), group: true, min: 1, max: 1 };
            i = close + 1;
        } else if (ch === "[") {
            let j = i + 1;
            let escaped = false;
            for (; j < pattern.length; j++) {
                if (!escaped && pattern[j] === "]") break;
                if (!escaped && pattern[j] === "\\") escaped = true;
                else escaped = false;
            }
            if (j >= pattern.length || j === i + 1) return false;
            atom = { source: pattern.slice(i, j + 1), min: 1, max: 1 };
            i = j + 1;
        } else if (ch === "\\") {
            const next = pattern[i + 1];
            if (!next) return false;
            if (!"dDsSwWnrt\\/.^$*+?()[]{}|-".includes(next)) return false;
            const literal = next === "n" ? "\n" : next === "r" ? "\r" : next === "t" ? "\t"
                : "dDsSwW".includes(next) ? undefined : next;
            atom = { source: pattern.slice(i, i + 2), literal, min: 1, max: 1 };
            i += 2;
        } else if (ch === ".") {
            atom = { source: ".", min: 1, max: 1 };
            i++;
        } else {
            if ("^$*+?{}()[]|".includes(ch)) return false;
            atom = { source: ch, literal: ch, min: 1, max: 1 };
            i++;
        }
        const next = pattern[i];
        if (next === "*" || next === "+" || next === "?" || next === "{") {
            if (atom.group) return false;
            if (next === "*") { atom.min = 0; atom.max = Infinity; i++; }
            else if (next === "+") { atom.min = 1; atom.max = Infinity; i++; }
            else if (next === "?") { atom.min = 0; atom.max = 1; i++; }
            else {
                const count = /^\{(\d{1,3})(?:,(\d{0,3}))?\}/.exec(pattern.slice(i));
                if (!count) return false;
                atom.min = Number(count[1]);
                atom.max = count[2] === undefined ? atom.min
                    : count[2] === "" ? Infinity : Number(count[2]);
                if (atom.min > 64 || atom.max < atom.min || atom.max > 64) return false;
                i += count[0].length;
            }
        }
        atoms.push(atom);
    }
    const variable = atoms.map((atom, index) => ({ atom, index }))
        .filter(({ atom }) => atom.max > atom.min);
    // Unanchored variable-width repetitions are quadratic under repeated search.
    if (variable.length && !pattern.startsWith("^")) return false;
    const matches = (atom: Atom, literal: string): boolean => {
        try { return new RegExp("^(?:" + atom.source + ")$", "i").test(literal); }
        catch { return true; }
    };
    for (let a = 0; a < variable.length; a++) {
        for (let b = a + 1; b < variable.length; b++) {
            const left = variable[a], right = variable[b];
            if (left.atom.literal !== undefined && !matches(right.atom, left.atom.literal)) continue;
            if (right.atom.literal !== undefined && !matches(left.atom, right.atom.literal)) continue;
            let separated = false;
            for (let k = left.index + 1; k < right.index; k++) {
                const middle = atoms[k];
                if (middle.min >= 1 && middle.literal !== undefined &&
                    !matches(left.atom, middle.literal) && !matches(right.atom, middle.literal)) {
                    separated = true;
                    break;
                }
            }
            if (!separated) return false;
        }
    }
    return true;
}

export class FilterMatcher {
    /** 编译通配符 Glob 为安全、确定性的正则表达式 */
    static globToRegex(glob: string): RegExp {
        const clean = glob.trim().replace(/\\/g, "/");
        // 去除首尾斜杠以适配相对路径
        const trimmed = clean.replace(/^\/+|\/+$/g, "");
        if (!trimmed) return /^$/;

        // 如果是 dir/** 模式，支持匹配 dir 本身及其所有深层子孙
        if (trimmed.endsWith("/**")) {
            const prefix = trimmed.slice(0, -3).replace(/\/+$/, "");
            if (prefix) {
                const subRegex = this.globToRegex(prefix);
                const subStr = subRegex.source.slice(1, -1);
                return new RegExp(`^${subStr}(?:/.*)?$`, "i");
            }
        }

        let regexStr = "";
        let i = 0;
        while (i < trimmed.length) {
            const char = trimmed[i];
            if (char === "*") {
                if (trimmed[i + 1] === "*") {
                    // ** 匹配多级目录
                    regexStr += ".*";
                    i += 2;
                    if (trimmed[i] === "/") i++; // 跳过紧跟的斜杠
                    continue;
                } else {
                    // * 匹配单层任意字符（不包含 /）
                    regexStr += "[^/]*";
                    i++;
                    continue;
                }
            } else if (char === "?") {
                regexStr += "[^/]";
                i++;
                continue;
            } else if (["[", "]", "(", ")", "{", "}", "+", ".", "^", "$", "|"].includes(char)) {
                regexStr += `\\${char}`;
            } else {
                regexStr += char;
            }
            i++;
        }
        return new RegExp(`^${regexStr}$`, "i");
    }

    /** 编译并校验用户正则表达式，防御 ReDoS 与语法错误 */
    static compileSafeRegex(pattern: string): RegExp | null {
        if (!pattern || pattern.length > 128) return null;
        // Parse a deliberately narrow grammar instead of just searching for
        // obvious "(a+)+" text. Unsupported syntax fails closed.
        if (!hasSafeRegexStructure(pattern)) return null;
        // Additional defense for legacy variants of quantified groups.
        if (/\([^)]*\)[+*{]|\(\?|\\[1-9]|[+*?]\s*[+*{]/.test(pattern)) return null;
        if (DANGEROUS_REGEX_PATTERN.test(pattern)) {
            console.warn(`[SmartMD Filter] 拒绝高危 ReDoS 正则: ${pattern}`);
            return null;
        }
        try {
            return new RegExp(pattern, "i");
        } catch {
            return null;
        }
    }

    static validateConfig(config: ExternalFilterConfig): string[] {
        const errors: string[] = [];
        if (config.mode !== "include" && config.mode !== "exclude") errors.push("未知筛选模式");
        for (const rule of config.userRules || []) {
            if (rule.enabled === false) continue;
            if (!rule.pattern || !rule.pattern.trim()) {
                errors.push("筛选表达式不能为空");
            } else if (rule.syntax === "regex" && !this.compileSafeRegex(rule.pattern)) {
                errors.push("正则表达式不安全或格式错误：" + rule.pattern.slice(0, 40));
            } else if (rule.syntax === "glob" && rule.pattern.length > 256) {
                errors.push("通配符过长");
            } else if (!["regex", "glob", "exact"].includes(rule.syntax) ||
                       !["folder", "file", "all"].includes(rule.target)) {
                errors.push("筛选规则维度或语法不正确");
            }
        }
        return errors;
    }

    private static matchesDir(path: string, name: string, rule: FilterRuleItem): boolean {
        if (rule.target === "file") return false;
        const pattern = rule.pattern.trim().replace(/\\/g, "/").replace(/^\/+|\/+$/g, "");
        if (!pattern) return false;
        if (rule.syntax === "exact") {
            return pattern.includes("/")
                ? path.toLowerCase() === pattern.toLowerCase()
                : name.toLowerCase() === pattern.toLowerCase();
        }
        const re = rule.syntax === "regex"
            ? this.compileSafeRegex(rule.pattern)
            : this.globToRegex(pattern);
        return !!re && (re.test(path) || re.test(name));
    }

    private static matchesFile(path: string, name: string, rule: FilterRuleItem): boolean {
        if (rule.target === "folder") return false;
        const pattern = rule.pattern.trim().replace(/\\/g, "/").replace(/^\/+|\/+$/g, "");
        if (!pattern) return false;
        if (rule.syntax === "exact") {
            return pattern.includes("/")
                ? path.toLowerCase() === pattern.toLowerCase()
                : name.toLowerCase() === pattern.toLowerCase();
        }
        const re = rule.syntax === "regex"
            ? this.compileSafeRegex(rule.pattern)
            : this.globToRegex(pattern);
        return !!re && (re.test(path) || re.test(name));
    }

    /** 计算配置的确定性规则指纹，用于漂移感知 */
    static computeFilterHash(config?: ExternalFilterConfig): string {
        if (!config) return "default_filter_hash";
        const parts: string[] = [
            config.mode || "exclude",
            String(config.builtinPresets?.ignoreBuildArtifacts !== false),
            String(config.builtinPresets?.ignoreIDEConfig !== false),
        ];
        const activeRules = (config.userRules || [])
            .filter((r) => r.enabled !== false && r.pattern?.trim())
            .map((r) => `${r.target || "all"}:${r.syntax || "glob"}:${r.pattern.trim()}`)
            .sort();
        parts.push(activeRules.join(";"));

        let hash = 0;
        const str = parts.join("|");
        for (let i = 0; i < str.length; i++) {
            hash = (hash << 5) - hash + str.charCodeAt(i);
            hash |= 0;
        }
        return `fh_${Math.abs(hash).toString(36)}`;
    }

    /**
     * 目录级早期剪枝（Directory Pruning）
     * 在 readdir 遍历到子目录时执行，命中则直接 O(1) 短路跳过整棵子树递归
     */
    static shouldPruneDirectory(dirRelativePath: string, dirName: string, config?: ExternalFilterConfig): boolean {
        const rel = dirRelativePath.replace(/\\/g, "/").replace(/^\/+|\/+$/g, "");
        const lower = dirName.toLowerCase();
        if (IMMUTABLE_DIR_BLACKLIST.has(lower) ||
            rel.toLowerCase().split("/").some((x) => IMMUTABLE_DIR_BLACKLIST.has(x))) return true;
        if (config?.builtinPresets?.ignoreBuildArtifacts !== false && DEFAULT_BUILD_ARTIFACTS.has(lower)) return true;
        if (config?.builtinPresets?.ignoreIDEConfig !== false && DEFAULT_IDE_CONFIGS.has(lower)) return true;
        const rules = (config?.userRules || []).filter((r) => r.enabled !== false && !!r.pattern?.trim());
        if ((config?.mode || "exclude") === "exclude") {
            return rules.some((r) => this.matchesDir(rel, dirName, r));
        }
        if (!rules.length) return true;
        // A file rule or a basename rule may match arbitrarily deep files.
        // Never prune until the scanner can PROVE that no descendant can match.
        return !rules.some((r) => {
            const pattern = r.pattern.trim().replace(/\\/g, "/").replace(/^\/+|\/+$/g, "");
            if (r.target !== "folder" || r.syntax !== "exact" || !pattern.includes("/")) return true;
            const p = pattern.toLowerCase(), d = rel.toLowerCase();
            return p === d || p.startsWith(d + "/") || d.startsWith(p + "/");
        });
    }

    /**
     * 文件级复合判定
     * 内部递归校验全部祖先目录，彻底解决 scanVault 扁平遍历下的穿透漏洞
     */
    static isFileAccepted(relativePath: string, config?: ExternalFilterConfig): boolean {
        const rel = relativePath.replace(/\\/g, "/").replace(/^\/+/, "");
        if (!rel.toLowerCase().endsWith(".md") || rel.includes("../") || rel.startsWith("/")) return false;
        const parts = rel.split("/");
        if (parts.some((p) => !p || p === "." || p === "..")) return false;
        const fileName = parts[parts.length - 1];
        if (IMMUTABLE_FILE_BLACKLIST.has(fileName.toLowerCase()) || fileName.startsWith("~$")) return false;
        const rules = (config?.userRules || []).filter((r) => r.enabled !== false && !!r.pattern?.trim());
        let dirMatched = false;
        let dir = "";
        for (let i = 0; i < parts.length - 1; i++) {
            dir = dir ? dir + "/" + parts[i] : parts[i];
            if (this.shouldPruneDirectory(dir, parts[i], {
                ...config, mode: "exclude",
                userRules: [], builtinPresets: config?.builtinPresets
            } as ExternalFilterConfig)) return false;
            if (rules.some((r) => this.matchesDir(dir, parts[i], r))) dirMatched = true;
        }
        const fileMatched = rules.some((r) => this.matchesFile(rel, fileName, r));
        if ((config?.mode || "exclude") === "exclude") return !(dirMatched || fileMatched);
        return dirMatched || fileMatched;
    }

}
