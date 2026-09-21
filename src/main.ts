import { Notice, Plugin } from 'obsidian';
import { DEFAULT_SETTINGS, TodoSyncSettings, TodoSyncSettingTab } from './settings';
import { createVaultAdapter } from './obsidian-tasks/vault-adapter';
import { createCachePlugin, hasStoredSession } from './auth/token-storage';
import { MsalDeviceCodeAuth } from './auth/msal-device-code';
import { TodoClient } from './todo-api/client';
import { SyncEngine } from './sync/engine';
import { loadSyncData, serializeSyncData, type SyncData } from './sync/state-store';
import { DeviceCodeModal } from './ui/device-code-modal';
import { RetryAfterError } from './todo-api/types';

interface PersistedData {
	settings?: Partial<TodoSyncSettings>;
	sync?: unknown;
}

export default class TodoSyncPlugin extends Plugin {
	settings!: TodoSyncSettings;
	private syncData!: SyncData;
	private auth!: MsalDeviceCodeAuth;
	private todoClient!: TodoClient;
	private engine!: SyncEngine;
	private pollIntervalId: number | null = null;

	async onload() {
		await this.loadPersistedData();

		this.auth = new MsalDeviceCodeAuth(createCachePlugin(this.app.secretStorage));
		this.todoClient = new TodoClient(() => this.getAccessToken());
		this.engine = new SyncEngine({
			todo: this.todoClient,
			vault: createVaultAdapter(this.app),
			now: () => new Date(),
			reminderTime: this.settings.reminderTime,
			listName: this.settings.listName,
			log: (message) => console.log(`[todo-sync] ${message}`),
		});

		this.addSettingTab(new TodoSyncSettingTab(this.app, this));

		this.addCommand({
			id: 'sign-in-to-microsoft',
			name: 'Sign in to Microsoft',
			callback: () => this.signIn(),
		});
		this.addCommand({
			id: 'sign-out-of-microsoft',
			name: 'Sign out',
			callback: () => this.signOut(),
		});
		this.addCommand({
			id: 'sync-now',
			name: 'Sync now',
			callback: () => this.runSyncCycle(),
		});

		if (hasStoredSession(this.app.secretStorage)) {
			this.startPolling();
		}
	}

	onunload() {
		this.stopPolling();
	}

	async loadPersistedData(): Promise<void> {
		const raw = ((await this.loadData()) ?? {}) as PersistedData;
		this.settings = { ...DEFAULT_SETTINGS, ...raw.settings };
		this.syncData = loadSyncData(raw.sync);
	}

	async saveSettings(): Promise<void> {
		await this.persist();
	}

	private async persist(): Promise<void> {
		const data: PersistedData = {
			settings: this.settings,
			sync: serializeSyncData(this.syncData),
		};
		await this.saveData(data);
	}

	private async getAccessToken(): Promise<string> {
		const token = await this.auth.getAccessTokenSilent();
		if (!token) {
			this.stopPolling();
			new Notice('Microsoft sign-in expired. Run "Sign in to Microsoft" to reconnect.');
			throw new Error('No valid access token; re-authentication required');
		}
		return token;
	}

	async signIn(): Promise<void> {
		const modal = new DeviceCodeModal(this.app);
		modal.open();
		try {
			await this.auth.signIn((info) => modal.showDeviceCode(info));
			this.settings.signedInAccountLabel = 'Microsoft account';
			await this.saveSettings();
			modal.showSuccess(this.settings.signedInAccountLabel);
			this.startPolling();
			await this.runSyncCycle();
		} catch (error) {
			modal.showError(error instanceof Error ? error.message : String(error));
		}
	}

	async signOut(): Promise<void> {
		this.stopPolling();
		await this.auth.signOut();
		this.settings.signedInAccountLabel = null;
		await this.saveSettings();
		new Notice('Signed out of Microsoft To Do sync.');
	}

	async handleSignInOutFromSettings(): Promise<void> {
		if (this.settings.signedInAccountLabel) {
			await this.signOut();
		} else {
			await this.signIn();
		}
		this.app.workspace.trigger('layout-change');
	}

	startPolling(): void {
		this.stopPolling();
		const intervalMs = this.settings.pollIntervalMinutes * 60 * 1000;
		this.pollIntervalId = window.setInterval(() => this.runSyncCycle(), intervalMs);
		this.registerInterval(this.pollIntervalId);
	}

	restartPolling(): void {
		if (this.pollIntervalId !== null) this.startPolling();
	}

	private stopPolling(): void {
		if (this.pollIntervalId !== null) {
			window.clearInterval(this.pollIntervalId);
			this.pollIntervalId = null;
		}
	}

	async runSyncCycle(): Promise<void> {
		try {
			this.syncData = await this.engine.runPollCycle(this.syncData);
			await this.persist();
		} catch (error) {
			if (error instanceof RetryAfterError) {
				console.warn(`[todo-sync] throttled, backing off ${error.retryAfterSeconds}s`);
				return;
			}
			console.error('[todo-sync] sync cycle failed', error);
		}
	}
}
