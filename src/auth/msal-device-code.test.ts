import { beforeEach, describe, expect, it, vi } from 'vitest';
import type { DeviceCodeRequest } from '@azure/msal-node';

type DeviceCodeResponse = Parameters<DeviceCodeRequest['deviceCodeCallback']>[0];

const acquireTokenByDeviceCode = vi.fn();
const acquireTokenSilent = vi.fn();
const getAllAccounts = vi.fn();
const removeAccount = vi.fn();

vi.mock('@azure/msal-node', () => ({
	PublicClientApplication: vi.fn().mockImplementation(() => ({
		acquireTokenByDeviceCode,
		acquireTokenSilent,
		getTokenCache: () => ({ getAllAccounts, removeAccount }),
	})),
}));

import { MsalDeviceCodeAuth } from './msal-device-code';

describe('MsalDeviceCodeAuth', () => {
	beforeEach(() => {
		acquireTokenByDeviceCode.mockReset();
		acquireTokenSilent.mockReset();
		getAllAccounts.mockReset();
		removeAccount.mockReset();
	});

	it('signIn forwards the device code details to the callback and returns the access token', async () => {
		acquireTokenByDeviceCode.mockImplementation(
			async ({
				deviceCodeCallback,
			}: {
				deviceCodeCallback: (response: DeviceCodeResponse) => void;
			}) => {
				deviceCodeCallback({
					userCode: 'ABC123',
					deviceCode: 'device-code-1',
					verificationUri: 'https://microsoft.com/devicelogin',
					message: 'Go there and enter ABC123',
					expiresIn: 900,
					interval: 5,
				});
				return { accessToken: 'token-1', account: { homeAccountId: 'acc-1' } };
			},
		);

		const auth = new MsalDeviceCodeAuth({} as never);
		const onDeviceCode = vi.fn();
		const token = await auth.signIn(onDeviceCode);

		expect(onDeviceCode).toHaveBeenCalledWith({
			userCode: 'ABC123',
			verificationUri: 'https://microsoft.com/devicelogin',
			message: 'Go there and enter ABC123',
			expiresIn: 900,
		});
		expect(token).toBe('token-1');
	});

	it('signIn throws when MSAL returns no result', async () => {
		acquireTokenByDeviceCode.mockResolvedValue(null);
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
