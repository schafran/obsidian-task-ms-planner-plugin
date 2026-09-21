import { vi } from 'vitest';

export const requestUrl = vi.fn();

export class Notice {
	constructor(_message: string) {}
}

export class Component {
	load() {}
	unload() {}
	registerEvent(..._args: unknown[]) {}
	registerDomEvent(..._args: unknown[]) {}
	registerInterval(id: number) {
		return id;
	}
}

export class Plugin extends Component {
	app: unknown;
	manifest: unknown;
	constructor(app: unknown, manifest: unknown) {
		super();
		this.app = app;
		this.manifest = manifest;
	}
	addCommand(..._args: unknown[]) {}
	addSettingTab(..._args: unknown[]) {}
	addRibbonIcon(..._args: unknown[]) {}
	addStatusBarItem() {
		return { setText(_t: string) {} };
	}
	async loadData(): Promise<unknown> {
		return {};
	}
	async saveData(_data: unknown): Promise<void> {}
}

export class Modal {
	app: unknown;
	contentEl: HTMLElement;
	constructor(app: unknown) {
		this.app = app;
		this.contentEl = document.createElement('div');
	}
	open() {
		this.onOpen();
	}
	close() {
		this.onClose();
	}
	onOpen() {}
	onClose() {}
}

export class PluginSettingTab {
	app: unknown;
	plugin: unknown;
	containerEl: HTMLElement;
	constructor(app: unknown, plugin: unknown) {
		this.app = app;
		this.plugin = plugin;
		this.containerEl = document.createElement('div');
	}
	display() {}
}

export class Setting {
	settingEl: HTMLElement;
	constructor(containerEl: HTMLElement) {
		this.settingEl = document.createElement('div');
		containerEl.appendChild(this.settingEl);
	}
	setName(_n: string) {
		return this;
	}
	setDesc(_d: string) {
		return this;
	}
	addText(cb: (t: unknown) => void) {
		cb({ setValue: () => ({}), onChange: () => ({}), setPlaceholder: () => ({}) });
		return this;
	}
	addButton(cb: (b: unknown) => void) {
		cb({ setButtonText: () => ({ onClick: () => ({}) }), onClick: () => ({}) });
		return this;
	}
	addToggle(cb: (t: unknown) => void) {
		cb({ setValue: () => ({}), onChange: () => ({}) });
		return this;
	}
}

export class TFile {
	path: string;
	stat: { mtime: number };
	constructor(path: string, mtime = 0) {
		this.path = path;
		this.stat = { mtime };
	}
}

export function normalizePath(path: string): string {
	return path.replace(/\\/g, '/');
}

export interface SecretStorage {
	setSecret(id: string, secret: string): void;
	getSecret(id: string): string | null;
	listSecrets(): string[];
}
