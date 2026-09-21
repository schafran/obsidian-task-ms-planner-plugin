import type { SecretStorage } from 'obsidian';
import type { ICachePlugin } from '@azure/msal-node';

export const MSAL_CACHE_SECRET_ID = 'obsidian-todo-sync-msal-cache';

export function hasStoredSession(secretStorage: SecretStorage): boolean {
	return secretStorage.getSecret(MSAL_CACHE_SECRET_ID) !== null;
}

export function createCachePlugin(secretStorage: SecretStorage): ICachePlugin {
	return {
		beforeCacheAccess: async (context) => {
			const cached = secretStorage.getSecret(MSAL_CACHE_SECRET_ID);
			if (cached) context.tokenCache.deserialize(cached);
		},
		afterCacheAccess: async (context) => {
			if (context.cacheHasChanged) {
				secretStorage.setSecret(MSAL_CACHE_SECRET_ID, context.tokenCache.serialize());
			}
		},
	};
}
