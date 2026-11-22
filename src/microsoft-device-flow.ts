import { requestUrl } from "obsidian";

/** Contains only what a user must see in the Microsoft sign-in dialog. */
export interface MicrosoftDeviceCodeChallenge {
    verificationUri: string;
    userCode: string;
    expiresIn: number;
}

export interface MicrosoftDeviceToken {
    accessToken: string;
    refreshToken: string;
    expiresAt: number;
}

/**
 * Public-client OAuth 2.0 Device Authorization Grant (RFC 8628).
 * Device codes, user codes and tokens remain in memory and are never logged.
 */
export class MicrosoftDeviceCodeSession {
    private cancelled = false;
    private wakePoll: (() => void) | null = null;

    constructor(private clientId: string, private scope: string) {}

    cancel(): void {
        this.cancelled = true;
        this.wakePoll?.();
        this.wakePoll = null;
    }

    assertActive(): void {
        if (this.cancelled) throw new Error("微软设备码登录已取消。");
    }

    private async send(url: string, body: URLSearchParams): Promise<any> {
        try {
            return await requestUrl({
                url,
                method: "POST",
                contentType: "application/x-www-form-urlencoded",
                body: body.toString(),
                throw: false,
            });
        } catch {
            throw new Error("无法连接微软身份验证服务，请检查网络后重试。");
        }
    }

    private async waitForNextPoll(ms: number): Promise<void> {
        this.assertActive();
        await new Promise<void>((resolve) => {
            const timer = window.setTimeout(() => {
                this.wakePoll = null;
                resolve();
            }, ms);
            this.wakePoll = () => {
                window.clearTimeout(timer);
                this.wakePoll = null;
                resolve();
            };
        });
        this.assertActive();
    }

    async authorize(
        present: (challenge: MicrosoftDeviceCodeChallenge) => void | Promise<void>,
    ): Promise<MicrosoftDeviceToken> {
        this.assertActive();
        const identity = "https://login.microsoftonline.com/consumers/oauth2/v2.0";
        const issued = await this.send(identity + "/devicecode", new URLSearchParams({
            client_id: this.clientId,
            scope: this.scope,
        }));
        this.assertActive();
        if (issued.status !== 200) {
            const code = String(issued.json?.error || "");
            if (code === "invalid_client") {
                throw new Error("微软应用未启用公共客户端设备码登录，请检查应用注册设置。");
            }
            throw new Error("无法获取微软设备码，请检查应用权限和账号类型。");
        }

        const payload = issued.json || {};
        const deviceCode = String(payload.device_code || "");
        const userCode = String(payload.user_code || "");
        let verification: URL;
        try {
            verification = new URL(String(payload.verification_uri || ""));
        } catch {
            throw new Error("微软返回的设备码验证地址不合法。");
        }
        if (
            verification.protocol !== "https:" ||
            !["microsoft.com", "www.microsoft.com"].includes(verification.hostname) ||
            !deviceCode || !userCode
        ) {
            throw new Error("微软返回的设备码授权信息不完整或不可信。");
        }
        const lifetime = Number(payload.expires_in);
        if (!Number.isFinite(lifetime) || lifetime <= 0) {
            throw new Error("微软没有返回有效的设备码有效期。");
        }
        const expiresIn = Math.min(1800, lifetime);
        let interval = Math.max(5, Math.min(1800, Number(payload.interval) || 5));
        const deadline = Date.now() + expiresIn * 1000;
        await present({
            verificationUri: verification.toString(),
            userCode,
            expiresIn,
        });
        this.assertActive();

        while (Date.now() < deadline) {
            this.assertActive();
            await this.waitForNextPoll(Math.min(interval * 1000, Math.max(1, deadline - Date.now())));
            this.assertActive();
            if (Date.now() >= deadline) break;
            const response = await this.send(identity + "/token", new URLSearchParams({
                client_id: this.clientId,
                grant_type: "urn:ietf:params:oauth:grant-type:device_code",
                device_code: deviceCode,
            }));
            this.assertActive();

            if (response.status === 200) {
                const body = response.json || {};
                if (!body.access_token || !body.refresh_token) {
                    throw new Error("微软没有返回完整的授权令牌，登录未完成。");
                }
                return {
                    accessToken: String(body.access_token),
                    refreshToken: String(body.refresh_token),
                    expiresAt: Date.now() + Math.max(1, Number(body.expires_in) || 3600) * 1000,
                };
            }

            const reason = String(response.json?.error || "");
            if (reason === "authorization_pending") continue;
            if (reason === "slow_down" || response.status === 429 ||
                reason === "temporarily_unavailable" || response.status === 503) {
                interval = Math.min(1800, interval + 5);
                continue;
            }
            if (reason === "authorization_declined" || reason === "access_denied") {
                throw new Error("微软授权已被拒绝，未连接 OneDrive。");
            }
            if (reason === "expired_token" || reason === "bad_verification_code") {
                throw new Error("微软设备码已过期，请重新开始登录。");
            }
            if (reason === "invalid_client") {
                throw new Error("微软应用不允许设备码登录，请检查公共客户端设置。");
            }
            throw new Error("微软设备码授权未通过，请重新开始登录。");
        }
        throw new Error("微软设备码已超时，请重新开始登录。");
    }
}
