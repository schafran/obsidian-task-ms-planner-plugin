import { describe, expect, it } from 'vitest';
import { getIsoWeekNotePath, insertActionItem } from './weekly-note';

describe('getIsoWeekNotePath', () => {
	it('computes the ISO week for a Thursday (unambiguous case)', () => {
		// 2026-01-01 is a Thursday -> ISO week 1 of 2026.
		expect(getIsoWeekNotePath(new Date(Date.UTC(2026, 0, 1)))).toBe(
			'Calendar/Weekly/2026-W01.md',
		);
	});

	it('rolls a Sunday into the same ISO week as its Monday', () => {
		// 2026-09-21 is a Monday (ISO week 39); 2026-09-27 is the following Sunday.
		expect(getIsoWeekNotePath(new Date(Date.UTC(2026, 8, 21)))).toBe(
			'Calendar/Weekly/2026-W39.md',
		);
		expect(getIsoWeekNotePath(new Date(Date.UTC(2026, 8, 27)))).toBe(
			'Calendar/Weekly/2026-W39.md',
		);
	});

	it('assigns late-December dates to week 1 of the next year when applicable', () => {
		// 2025-12-31 is a Wednesday in the ISO week containing 2026-01-01.
		expect(getIsoWeekNotePath(new Date(Date.UTC(2025, 11, 31)))).toBe(
			'Calendar/Weekly/2026-W01.md',
		);
	});
});

describe('insertActionItem', () => {
	const NEW_LINE = '- [ ] Buy milk 📅 2026-10-01 %%todo:xyz%%';

	it('replaces the empty placeholder line under the heading', () => {
		const content = [
			'# Week 39',
			'',
			'### Other Action Items',
			'- [ ] ✅',
			'',
			'### Notes',
		].join('\n');
		const result = insertActionItem(content, NEW_LINE);
		expect(result).toBe(
			['# Week 39', '', '### Other Action Items', NEW_LINE, '', '### Notes'].join(
				'\n',
			),
		);
	});

	it('appends after existing items when there is no placeholder', () => {
		const content = [
			'### Other Action Items',
			'- [ ] Existing item 📅 2026-09-25',
			'### Notes',
		].join('\n');
		const result = insertActionItem(content, NEW_LINE);
		expect(result).toBe(
			[
				'### Other Action Items',
				'- [ ] Existing item 📅 2026-09-25',
				NEW_LINE,
				'### Notes',
			].join('\n'),
		);
	});

	it('appends directly under the heading when the section is empty and heading is last line', () => {
		const content = ['# Week 39', '', '### Other Action Items'].join('\n');
		const result = insertActionItem(content, NEW_LINE);
		expect(result).toBe(
			['# Week 39', '', '### Other Action Items', NEW_LINE].join('\n'),
		);
	});

	it('throws if the heading is missing (caller must have checked the file exists and matches)', () => {
		expect(() => insertActionItem('# Week 39\nno such section', NEW_LINE)).toThrow(
			/Other Action Items/,
		);
	});
});
