import type { App, TFile } from 'obsidian';

export interface VaultFile {
	path: string;
	mtimeMs: number;
}

export interface VaultAdapter {
	listMarkdownFiles(): VaultFile[];
	read(file: VaultFile): Promise<string>;
	update(file: VaultFile, mutate: (content: string) => string): Promise<void>;
	getFile(path: string): VaultFile | null;
}

function toVaultFile(file: TFile): VaultFile {
	return { path: file.path, mtimeMs: file.stat.mtime };
}

export function createVaultAdapter(app: App): VaultAdapter {
	return {
		listMarkdownFiles() {
			return app.vault.getMarkdownFiles().map(toVaultFile);
		},
		async read(file) {
			const tFile = app.vault.getFileByPath(file.path);
			if (!tFile) throw new Error(`File not found: ${file.path}`);
			return app.vault.read(tFile);
		},
		async update(file, mutate) {
			const tFile = app.vault.getFileByPath(file.path);
			if (!tFile) throw new Error(`File not found: ${file.path}`);
			await app.vault.process(tFile, mutate);
		},
		getFile(path) {
			const tFile = app.vault.getFileByPath(path);
			return tFile ? toVaultFile(tFile) : null;
		},
	};
}
