import { PublicClientApplication, type AccountInfo, type ICachePlugin } from '@azure/msal-node';
import { AUTHORITY, CLIENT_ID, SCOPES } from './config';

export interface DeviceCodeInfo {
	userCode: string;
	verificationUri: string;
	message: string;
	expiresIn: number;
}

export class MsalDeviceCodeAuth {
	private readonly pca: PublicClientApplication;
	private account: AccountInfo | null = null;

	constructor(cachePlugin: ICachePlugin) {
		this.pca = new PublicClientApplication({
			auth: { clientId: CLIENT_ID, authority: AUTHORITY },
			cache: { cachePlugin },
		});
	}

	async signIn(onDeviceCode: (info: DeviceCodeInfo) => void): Promise<string> {
		const result = await this.pca.acquireTokenByDeviceCode({
			scopes: SCOPES,
			deviceCodeCallback: (response) => {
				onDeviceCode({
					userCode: response.userCode,
					verificationUri: response.verificationUri,
					message: response.message,
					expiresIn: response.expiresIn,
				});
			},
		});
		if (!result) throw new Error('Device code sign-in returned no result');
		this.account = result.account;
		return result.accessToken;
	}

	async getAccessTokenSilent(): Promise<string | null> {
		const accounts = await this.pca.getTokenCache().getAllAccounts();
		const account = this.account ?? accounts[0];
		if (!account) return null;
		this.account = account;
		try {
			const result = await this.pca.acquireTokenSilent({ scopes: SCOPES, account });
			return result?.accessToken ?? null;
		} catch {
			return null;
		}
	}

	async signOut(): Promise<void> {
		const cache = this.pca.getTokenCache();
		const accounts = await cache.getAllAccounts();
		for (const account of accounts) {
			await cache.removeAccount(account);
		}
		this.account = null;
	}
}
