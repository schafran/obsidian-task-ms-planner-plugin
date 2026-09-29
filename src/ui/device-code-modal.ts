import { App, Modal } from 'obsidian';

export class DeviceCodeModal extends Modal {
	constructor(app: App) {
		super(app);
	}

	showWaitingForBrowser(): void {
		const { contentEl } = this;
		contentEl.empty();
		contentEl.createEl('h2', { text: 'Sign in to Microsoft' });
		contentEl.createEl('p', {
			text: 'Opening your browser to sign in. Waiting for you to complete sign-in...',
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
