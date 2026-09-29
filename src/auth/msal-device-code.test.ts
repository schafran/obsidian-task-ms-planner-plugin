import { beforeEach, describe, expect, it, vi } from 'vitest';

const {
	getAuthCodeUrl,
	acquireTokenByCode,
	acquireTokenSilent,
	getAllAccounts,
	removeAccount,
	generatePkceCodes,
	createNewGuid,
	waitForAuthCode,
} = vi.hoisted(() => ({
	getAuthCodeUrl: vi.fn(),
	acquireTokenByCode: vi.fn(),
	acquireTokenSilent: vi.fn(),
	getAllAccounts: vi.fn(),
	removeAccount: vi.fn(),
	generatePkceCodes: vi.fn(),
	createNewGuid: vi.fn(),
	waitForAuthCode: vi.fn(),
}));

import { shell } from 'electron';
// eslint-disable-next-line @typescript-eslint/unbound-method -- test-only mock reference, never called unbound
const openExternal = shell.openExternal as ReturnType<typeof vi.fn>;

vi.mock('@azure/msal-node', () => ({
	PublicClientApplication: vi.fn().mockImplementation(() => ({
		getAuthCodeUrl,
		acquireTokenByCode,
		acquireTokenSilent,
		getTokenCache: () => ({ getAllAccounts, removeAccount }),
	})),
	CryptoProvider: vi.fn().mockImplementation(() => ({
		generatePkceCodes,
		createNewGuid,
	})),
}));

vi.mock('./loopback-server', () => ({
	waitForAuthCode,
}));

import { MsalDeviceCodeAuth } from './msal-device-code';

describe('MsalDeviceCodeAuth', () => {
	beforeEach(() => {
		getAuthCodeUrl.mockReset();
		acquireTokenByCode.mockReset();
		acquireTokenSilent.mockReset();
		getAllAccounts.mockReset();
		removeAccount.mockReset();
		generatePkceCodes.mockReset();
		createNewGuid.mockReset();
		waitForAuthCode.mockReset();
		openExternal.mockReset();

		generatePkceCodes.mockResolvedValue({ verifier: 'verifier-1', challenge: 'challenge-1' });
		createNewGuid.mockReturnValue('state-1');
		getAuthCodeUrl.mockResolvedValue('https://login.microsoftonline.com/authorize?...');
		openExternal.mockResolvedValue(undefined);
	});

	it('signIn opens the browser and exchanges the returned code for a token', async () => {
		waitForAuthCode.mockResolvedValue({ code: 'auth-code-1', state: 'state-1' });
		acquireTokenByCode.mockResolvedValue({ accessToken: 'token-1', account: { homeAccountId: 'acc-1' } });

		const auth = new MsalDeviceCodeAuth({} as never);
		const onWaitingForBrowser = vi.fn();
		const token = await auth.signIn(onWaitingForBrowser);

		expect(onWaitingForBrowser).toHaveBeenCalled();
		expect(openExternal).toHaveBeenCalledWith('https://login.microsoftonline.com/authorize?...');
		expect(acquireTokenByCode).toHaveBeenCalledWith(
			expect.objectContaining({ code: 'auth-code-1', codeVerifier: 'verifier-1' }),
		);
		expect(token).toBe('token-1');
	});

	it('signIn throws on state mismatch', async () => {
		waitForAuthCode.mockResolvedValue({ code: 'auth-code-1', state: 'wrong-state' });
		const auth = new MsalDeviceCodeAuth({} as never);
		await expect(auth.signIn(vi.fn())).rejects.toThrow(/state mismatch/);
	});

	it('signIn throws when MSAL returns no result', async () => {
		waitForAuthCode.mockResolvedValue({ code: 'auth-code-1', state: 'state-1' });
		acquireTokenByCode.mockResolvedValue(null);
		const auth = new MsalDeviceCodeAuth({} as never);
		await expect(auth.signIn(vi.fn())).rejects.toThrow(/no result/);
	});

	it('getAccessTokenSilent returns null when there is no cached account', async () => {
		getAllAccounts.mockResolvedValue([]);
		const auth = new MsalDeviceCodeAuth({} as never);
		expect(await auth.getAccessTokenSilent()).toBeNull();
	});

	it('getAccessTokenSilent uses the cached account to acquire a token silently', async () => {
		const account = { homeAccountId: 'acc-1' };
		getAllAccounts.mockResolvedValue([account]);
		acquireTokenSilent.mockResolvedValue({ accessToken: 'token-2' });
		const auth = new MsalDeviceCodeAuth({} as never);
		const token = await auth.getAccessTokenSilent();
		expect(token).toBe('token-2');
		expect(acquireTokenSilent).toHaveBeenCalledWith(
			expect.objectContaining({ account }),
		);
	});

	it('getAccessTokenSilent returns null if silent acquisition fails', async () => {
		getAllAccounts.mockResolvedValue([{ homeAccountId: 'acc-1' }]);
		acquireTokenSilent.mockRejectedValue(new Error('interaction_required'));
		const auth = new MsalDeviceCodeAuth({} as never);
		expect(await auth.getAccessTokenSilent()).toBeNull();
	});

	it('signOut removes all cached accounts', async () => {
		const accounts = [{ homeAccountId: 'acc-1' }, { homeAccountId: 'acc-2' }];
		getAllAccounts.mockResolvedValue(accounts);
		const auth = new MsalDeviceCodeAuth({} as never);
		await auth.signOut();
		expect(removeAccount).toHaveBeenCalledTimes(2);
		expect(removeAccount).toHaveBeenCalledWith(accounts[0]);
		expect(removeAccount).toHaveBeenCalledWith(accounts[1]);
	});
});
