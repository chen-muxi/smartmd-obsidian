import { Notice, requestUrl } from "obsidian";
import { SmartMDStore } from "./core";
import { SMARTKEEP_PUBLISH_URL } from "./types";

const SMARTKEEP_LOGIN = "https://account.smartkeep.cn/login";
const SMARTKEEP_REDIRECT = "obsidian://smartkeep-connect-auth";
const SMARTKEEP_SECRET = "smartkeep-connect-md-session";

type SmartKeepSession = {
    accessToken: string;
    expiresAt: number;
};

export class SmartKeepAuth {
    private signingIn = false;
    private completing = false;
    private pendingState = "";
    private pendingResolve: (() => void) | null = null;
    private pendingReject: ((error: Error) => void) | null = null;
    private pendingTimer: number | null = null;
    private beforeSignOut: () => Promise<void> = async () => {};
    private afterSignIn: () => Promise<void> = async () => {};

    constructor(private store: SmartMDStore) {}

    setPublishLifecycle(beforeSignOut: () => Promise<void>, afterSignIn: () => Promise<void>): void {
        this.beforeSignOut = beforeSignOut;
        this.afterSignIn = afterSignIn;
    }

    isSignedIn(): boolean {
        const session = this.session();
        return Boolean(
            this.store.settings.local.smartkeepAccountId
            && session
            && session.expiresAt > Date.now(),
        );
    }

    accountName(): string {
        return this.store.settings.local.smartkeepAccountName || "SmartKeep 用户";
    }

    async signIn(openUrl: (url: string) => void): Promise<void> {
        if (this.signingIn) throw new Error("SmartKeep 登录正在进行中");
        this.signingIn = true;
        this.pendingState = this.randomValue();
        const query = new URLSearchParams({
            client: "smartkeep-md",
            redirect_uri: SMARTKEEP_REDIRECT,
            state: this.pendingState,
        });
        const result = new Promise<void>((resolve, reject) => {
            this.pendingResolve = resolve;
            this.pendingReject = reject;
            this.pendingTimer = window.setTimeout(
                () => this.finish(new Error("SmartKeep 登录已超时，请重新开始")),
                10 * 60 * 1000,
            );
        });
        openUrl(`${SMARTKEEP_LOGIN}?${query.toString()}`);
        await result;
    }

    async completeSignIn(params: Record<string, string>): Promise<void> {
        if (!this.signingIn || !this.pendingState || this.completing) return;
        this.completing = true;
        try {
            if (params.error) throw new Error(params.error_description || params.error);
            if (!params.exchange_code) throw new Error("SmartKeep 没有返回有效的登录结果");
            if (params.state !== this.pendingState) throw new Error("SmartKeep 登录校验失败，请重新开始");
            const base = this.serverBase();
            const response = await requestUrl({
                url: `${base}/api/v1/auth/exchange`,
                method: "POST",
                contentType: "application/json",
                body: JSON.stringify({ exchangeCode: params.exchange_code }),
                throw: false,
            });
            let json: any = null;
            try {
                json = response.json;
            } catch {
                try {
                    json = JSON.parse(response.text || "");
                } catch {
                    json = null;
                }
            }
            if (response.status < 200 || response.status >= 300 || !json?.accessToken) {
                throw new Error(json?.message || (response.text && !response.text.startsWith("<") ? response.text : "") || "SmartKeep MD 登录失败");
            }
            const account = json.account || {};
            await this.store.setSecret(SMARTKEEP_SECRET, JSON.stringify({
                accessToken: json.accessToken,
                expiresAt: Number(json.expiresAt || Date.now() + 30 * 24 * 60 * 60 * 1000),
            } as SmartKeepSession));
            this.store.settings.local.smartkeepAccountId = String(account.id || "");
            this.store.settings.local.smartkeepAccountName = String(account.displayName || "SmartKeep 用户");
            await this.store.save();
            try {
                await this.afterSignIn();
            } catch {}
            this.finish();
        } catch (error) {
            await this.clearLocal();
            this.finish(error);
        } finally {
            this.completing = false;
        }
    }

    async signOut(): Promise<void> {
        await this.beforeSignOut();
        const session = this.session();
        if (session) {
            try {
                await requestUrl({
                    url: `${this.serverBase()}/api/v1/auth/logout`,
                    method: "POST",
                    headers: { Authorization: `Bearer ${session.accessToken}` },
                    throw: false,
                });
            } catch {}
        }
        await this.clearLocal();
    }

    async getAccessToken(): Promise<string> {
        const session = this.session();
        if (!session || session.expiresAt <= Date.now()) {
            await this.clearLocal();
            throw new Error("请先登录 SmartKeep 账号");
        }
        return session.accessToken;
    }

    private serverBase(): string {
        return SMARTKEEP_PUBLISH_URL;
    }

    private session(): SmartKeepSession | null {
        const raw = this.store.getSecret(SMARTKEEP_SECRET);
        if (!raw) return null;
        try {
            const value = JSON.parse(raw) as SmartKeepSession;
            return value.accessToken ? value : null;
        } catch {
            return null;
        }
    }

    private async clearLocal(): Promise<void> {
        await this.store.setSecret(SMARTKEEP_SECRET, "");
        this.store.settings.local.smartkeepAccountId = "";
        this.store.settings.local.smartkeepAccountName = "";
        await this.store.save();
    }

    private randomValue(): string {
        const bytes = new Uint8Array(24);
        crypto.getRandomValues(bytes);
        return Array.from(bytes).map((value) => value.toString(16).padStart(2, "0")).join("");
    }

    private finish(error?: unknown): void {
        if (this.pendingTimer !== null) window.clearTimeout(this.pendingTimer);
        const resolve = this.pendingResolve;
        const reject = this.pendingReject;
        this.pendingTimer = null;
        this.pendingResolve = null;
        this.pendingReject = null;
        this.pendingState = "";
        this.signingIn = false;
        if (error) {
            const failure = error instanceof Error ? error : new Error(String(error));
            reject?.(failure);
            new Notice(`SmartKeep 登录失败：${failure.message}`, 8000);
        } else {
            resolve?.();
        }
    }
}
