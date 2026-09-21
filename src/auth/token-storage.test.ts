import { describe, expect, it, vi } from 'vitest';
import type { SecretStorage } from 'obsidian';
import type { TokenCacheContext } from '@azure/msal-node';
import { createCachePlugin, hasStoredSession, MSAL_CACHE_SECRET_ID } from './token-storage';

function fakeTokenCacheContext(
	tokenCache: { deserialize: (cache: string) => void; serialize: () => string },
	cacheHasChanged: boolean,
): TokenCacheContext {
	return { tokenCache, cacheHasChanged } as unknown as TokenCacheContext;
}

function fakeSecretStorage(initial: Record<string, string> = {}) {
	const store = { ...initial };
	const setSecret = vi.fn((id: string, secret: string) => {
		store[id] = secret;
	});
	const getSecret = vi.fn((id: string) => store[id] ?? null);
	const listSecrets = vi.fn(() => Object.keys(store));
	const storage: SecretStorage = { setSecret, getSecret, listSecrets };
	return { storage, setSecret, getSecret, listSecrets };
}

describe('hasStoredSession', () => {
	it('is false when no cache secret has been stored yet', () => {
		expect(hasStoredSession(fakeSecretStorage().storage)).toBe(false);
	});

	it('is true once a cache secret exists', () => {
		expect(
			hasStoredSession(fakeSecretStorage({ [MSAL_CACHE_SECRET_ID]: '{}' }).storage),
		).toBe(true);
	});
});

describe('createCachePlugin', () => {
	it('loads the cache from secret storage on beforeCacheAccess', async () => {
		const { storage } = fakeSecretStorage({ [MSAL_CACHE_SECRET_ID]: '{"cached":true}' });
		const plugin = createCachePlugin(storage);
		const deserialize = vi.fn();
		await plugin.beforeCacheAccess(
			fakeTokenCacheContext({ deserialize, serialize: vi.fn() }, false),
		);
		expect(deserialize).toHaveBeenCalledWith('{"cached":true}');
	});

	it('does nothing on beforeCacheAccess when no cache is stored', async () => {
		const { storage } = fakeSecretStorage();
		const plugin = createCachePlugin(storage);
		const deserialize = vi.fn();
		await plugin.beforeCacheAccess(
			fakeTokenCacheContext({ deserialize, serialize: vi.fn() }, false),
		);
		expect(deserialize).not.toHaveBeenCalled();
	});

	it('persists the cache on afterCacheAccess only if it changed', async () => {
		const { storage, setSecret } = fakeSecretStorage();
		const plugin = createCachePlugin(storage);
		const serialize = vi.fn(() => '{"cached":true}');
		await plugin.afterCacheAccess(
			fakeTokenCacheContext({ deserialize: vi.fn(), serialize }, true),
		);
		expect(setSecret).toHaveBeenCalledWith(MSAL_CACHE_SECRET_ID, '{"cached":true}');
	});

	it('does not write when the cache has not changed', async () => {
		const { storage, setSecret } = fakeSecretStorage();
		const plugin = createCachePlugin(storage);
		await plugin.afterCacheAccess(
			fakeTokenCacheContext({ deserialize: vi.fn(), serialize: vi.fn() }, false),
		);
		expect(setSecret).not.toHaveBeenCalled();
	});
});
