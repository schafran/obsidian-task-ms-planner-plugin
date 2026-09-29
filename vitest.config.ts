import { defineConfig } from 'vitest/config';
import path from 'node:path';

export default defineConfig({
	test: {
		environment: 'jsdom',
		alias: {
			obsidian: path.resolve(__dirname, 'tests/mocks/obsidian.ts'),
			electron: path.resolve(__dirname, 'tests/mocks/electron.ts'),
		},
	},
});
