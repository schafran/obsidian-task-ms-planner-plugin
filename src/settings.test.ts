import { describe, expect, it } from 'vitest';
import { DEFAULT_SETTINGS } from './settings';

describe('DEFAULT_SETTINGS', () => {
	it('defaults to a 10 minute poll interval and 08:00 reminder time', () => {
		expect(DEFAULT_SETTINGS.pollIntervalMinutes).toBe(10);
		expect(DEFAULT_SETTINGS.reminderTime).toBe('08:00');
		expect(DEFAULT_SETTINGS.listName).toBe('Obsidian');
	});
});
