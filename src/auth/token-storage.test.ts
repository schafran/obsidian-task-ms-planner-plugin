import { describe, expect, it, vi } from 'vitest';
import type { SecretStorage } from 'obsidian';
import { createCachePlugin, hasStoredSession, MSAL_CACHE_SECRET_ID } from './token-storage';

function fakeSecretStorage(initial: Record<string, string> = {}): SecretStorage {
	const store = { ...initial };
	return {
		setSecret: vi.fn((id: string, secret: string) => {
			store[id] = secret;
		}),
		getSecret: vi.fn((id: string) => store[id] ?? null),
		listSecrets: vi.fn(() => Object.keys(store)),
	};
}

describe('hasStoredSession', () => {
	it('is false when no cache secret has been stored yet', () => {
		expect(hasStoredSession(fakeSecretStorage())).toBe(false);
	});

	it('is true once a cache secret exists', () => {
		expect(
			hasStoredSession(fakeSecretStorage({ [MSAL_CACHE_SECRET_ID]: '{}' })),
		).toBe(true);
	});
});

describe('createCachePlugin', () => {
	it('loads the cache from secret storage on beforeCacheAccess', async () => {
		const secretStorage = fakeSecretStorage({ [MSAL_CACHE_SECRET_ID]: '{"cached":true}' });
		const plugin = createCachePlugin(secretStorage);
		const deserialize = vi.fn();
		await plugin.beforeCacheAccess!({
			tokenCache: { deserialize, serialize: vi.fn() },
			cacheHasChanged: false,
		} as never);
		expect(deserialize).toHaveBeenCalledWith('{"cached":true}');
	});

	it('does nothing on beforeCacheAccess when no cache is stored', async () => {
		const secretStorage = fakeSecretStorage();
		const plugin = createCachePlugin(secretStorage);
		const deserialize = vi.fn();
		await plugin.beforeCacheAccess!({
			tokenCache: { deserialize, serialize: vi.fn() },
			cacheHasChanged: false,
		} as never);
		expect(deserialize).not.toHaveBeenCalled();
	});

	it('persists the cache on afterCacheAccess only if it changed', async () => {
		const secretStorage = fakeSecretStorage();
		const plugin = createCachePlugin(secretStorage);
		const serialize = vi.fn(() => '{"cached":true}');
		await plugin.afterCacheAccess!({
			tokenCache: { deserialize: vi.fn(), serialize },
			cacheHasChanged: true,
		} as never);
		expect(secretStorage.setSecret).toHaveBeenCalledWith(
			MSAL_CACHE_SECRET_ID,
			'{"cached":true}',
		);
	});

	it('does not write when the cache has not changed', async () => {
		const secretStorage = fakeSecretStorage();
		const plugin = createCachePlugin(secretStorage);
		await plugin.afterCacheAccess!({
			tokenCache: { deserialize: vi.fn(), serialize: vi.fn() },
			cacheHasChanged: false,
		} as never);
		expect(secretStorage.setSecret).not.toHaveBeenCalled();
	});
});
