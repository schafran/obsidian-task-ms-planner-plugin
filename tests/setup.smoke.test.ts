import { describe, expect, it } from 'vitest';
import { Notice } from 'obsidian';

describe('test harness', () => {
	it('resolves the obsidian mock instead of the real (typings-only) package', () => {
		expect(() => new Notice('hello')).not.toThrow();
	});
});
