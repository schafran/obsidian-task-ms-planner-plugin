import {
	PublicClientApplication,
	CryptoProvider,
	type AccountInfo,
	type ICachePlugin,
} from '@azure/msal-node';
import { AUTHORITY, CLIENT_ID, REDIRECT_URI, SCOPES } from './config';
import { ObsidianNetworkClient } from './obsidian-network-client';
import { waitForAuthCode } from './loopback-server';
import { shell } from 'electron';

const LOOPBACK_PORT = 8080;
const SIGN_IN_TIMEOUT_MS = 5 * 60 * 1000;

export class MsalDeviceCodeAuth {
	private readonly pca: PublicClientApplication;
	private readonly crypto = new CryptoProvider();
	private account: AccountInfo | null = null;

	constructor(cachePlugin: ICachePlugin) {
		this.pca = new PublicClientApplication({
			auth: { clientId: CLIENT_ID, authority: AUTHORITY },
			cache: { cachePlugin },
			system: { networkClient: new ObsidianNetworkClient() },
		});
	}

	async signIn(onWaitingForBrowser: () => void): Promise<string> {
		const { verifier, challenge } = await this.crypto.generatePkceCodes();
		const state = this.crypto.createNewGuid();

		const authCodeUrl = await this.pca.getAuthCodeUrl({
			scopes: SCOPES,
			redirectUri: REDIRECT_URI,
			codeChallenge: challenge,
			codeChallengeMethod: 'S256',
			state,
		});

		const authCodePromise = waitForAuthCode(LOOPBACK_PORT, SIGN_IN_TIMEOUT_MS);
		onWaitingForBrowser();
		await shell.openExternal(authCodeUrl);

		const { code, state: returnedState } = await authCodePromise;
		if (returnedState !== state) {
			throw new Error('Sign-in state mismatch; possible interference, please try again');
		}

		const result = await this.pca.acquireTokenByCode({
			scopes: SCOPES,
			redirectUri: REDIRECT_URI,
			code,
			codeVerifier: verifier,
		});
		if (!result) throw new Error('Sign-in returned no result');
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
