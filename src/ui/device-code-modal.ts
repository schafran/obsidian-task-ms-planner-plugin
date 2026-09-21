import { App, Modal } from 'obsidian';
import type { DeviceCodeInfo } from '../auth/msal-device-code';

export class DeviceCodeModal extends Modal {
	constructor(app: App) {
		super(app);
	}

	showDeviceCode(info: DeviceCodeInfo): void {
		const { contentEl } = this;
		contentEl.empty();
		contentEl.createEl('h2', { text: 'Sign in to Microsoft' });
		contentEl.createEl('p', {
			text: `Go to ${info.verificationUri} and enter this code:`,
		});
		contentEl.createEl('p', { text: info.userCode, cls: 'todo-sync-device-code' });
		contentEl.createEl('p', {
			text: 'Waiting for you to complete sign-in in your browser...',
		});
	}

	showSuccess(accountLabel: string): void {
		const { contentEl } = this;
		contentEl.empty();
		contentEl.createEl('p', { text: `Signed in as ${accountLabel}.` });
	}

	showError(message: string): void {
		const { contentEl } = this;
		contentEl.empty();
		contentEl.createEl('p', { text: `Sign-in failed: ${message}` });
	}

	onClose(): void {
		this.contentEl.empty();
	}
}
