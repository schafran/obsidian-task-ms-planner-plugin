import { describe, expect, it } from 'vitest';
import { findMarkerRanges } from './marker-ranges';

describe('findMarkerRanges', () => {
	it('finds a marker mid-line with offset applied', () => {
		const line = '- [ ] Buy milk %%todo:abc%% 📅 2026-10-01';
		const start = line.indexOf('%%');
		expect(findMarkerRanges(line, 100)).toEqual([
			{ from: 100 + start, to: 100 + start + '%%todo:abc%%'.length },
		]);
	});

	it('handles ids with =, - and _', () => {
		const marker = '%%todo:AAMk_g-tAAA=%%';
		const line = `- [x] Title ${marker} 📅 2026-10-01`;
		const [range] = findMarkerRanges(line, 0);
		expect(line.slice(range!.from, range!.to)).toBe(marker);
	});

	it('returns nothing without a marker', () => {
		expect(findMarkerRanges('- [ ] Buy milk 📅 2026-10-01', 0)).toEqual([]);
	});

	it('ignores non-task lines', () => {
		expect(findMarkerRanges('Some text %%todo:abc%%', 0)).toEqual([]);
	});

	it('supports indented tasks', () => {
		expect(findMarkerRanges('    - [ ] Nested %%todo:x%% 📅 2026-10-01', 0)).toHaveLength(1);
	});
});
