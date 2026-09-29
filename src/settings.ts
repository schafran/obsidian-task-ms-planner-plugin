import { App, PluginSettingTab, Setting } from 'obsidian';
import type TodoSyncPlugin from './main';

export interface TodoSyncSettings {
	listName: string;
	pollIntervalMinutes: number;
	reminderTime: string; // HH:mm, local
	signedInAccountLabel: string | null;
}

export const DEFAULT_SETTINGS: TodoSyncSettings = {
	listName: 'Obsidian',
	pollIntervalMinutes: 10,
	reminderTime: '08:00',
	signedInAccountLabel: null,
};

export class TodoSyncSettingTab extends PluginSettingTab {
	constructor(
		app: App,
		private readonly plugin: TodoSyncPlugin,
	) {
		super(app, plugin);
	}

	display(): void {
		const { containerEl } = this;
		containerEl.empty();

		new Setting(containerEl)
			.setName('Account')
			.setDesc(
				this.plugin.settings.signedInAccountLabel
					? `Signed in as ${this.plugin.settings.signedInAccountLabel}`
					: 'Not signed in',
			)
			.addButton((button) => {
				button
					.setButtonText(
						this.plugin.settings.signedInAccountLabel ? 'Sign out' : 'Sign in',
					)
					.onClick(async () => {
						button.setDisabled(true);
						try {
							await this.plugin.handleSignInOutFromSettings();
						} finally {
							this.display();
						}
					});
			});

		new Setting(containerEl)
			.setName('To Do list name')
			.setDesc('Created automatically on first sign-in if it does not already exist.')
			.addText((text) =>
				text.setValue(this.plugin.settings.listName).onChange(async (value) => {
					this.plugin.settings.listName = value.trim() || DEFAULT_SETTINGS.listName;
					await this.plugin.saveSettings();
				}),
			);

		new Setting(containerEl)
			.setName('Poll interval (minutes)')
			.addText((text) =>
				text
					.setValue(String(this.plugin.settings.pollIntervalMinutes))
					.onChange(async (value) => {
						const minutes = Number(value);
						if (Number.isFinite(minutes) && minutes > 0) {
							this.plugin.settings.pollIntervalMinutes = minutes;
							await this.plugin.saveSettings();
							this.plugin.restartPolling();
						}
					}),
			);

		new Setting(containerEl)
			.setName('Default reminder time')
			.setDesc('Local time used for the reminder on tasks pushed from Obsidian to To Do.')
			.addText((text) =>
				text.setValue(this.plugin.settings.reminderTime).onChange(async (value) => {
					if (/^\d{2}:\d{2}$/.test(value)) {
						this.plugin.settings.reminderTime = value;
						await this.plugin.saveSettings();
					}
				}),
			);
	}
}
