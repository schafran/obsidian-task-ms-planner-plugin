import { describe, expect, it } from 'vitest';
import { parseTaskLine, renderTaskLine } from './parser';

describe('parseTaskLine', () => {
	it('returns null for a non-checkbox line', () => {
		expect(parseTaskLine('Just some text')).toBeNull();
	});

	it('returns null for a checkbox line with no due date', () => {
		expect(parseTaskLine('- [ ] Undated task')).toBeNull();
	});

	it('parses an open task with a due date', () => {
		expect(parseTaskLine('- [ ] Renew passport 📅 2026-10-01')).toEqual({
			indent: '',
			checked: false,
			title: 'Renew passport',
			dueDate: '2026-10-01',
			doneDate: null,
			recurring: false,
			todoId: null,
		});
	});

	it('parses a completed task with a done date', () => {
		expect(
			parseTaskLine('- [x] Renew passport 📅 2026-10-01 ✅ 2026-09-30'),
		).toEqual({
			indent: '',
			checked: true,
			title: 'Renew passport',
			dueDate: '2026-10-01',
			doneDate: '2026-09-30',
			recurring: false,
			todoId: null,
		});
	});

	it('parses the sync marker', () => {
		const parsed = parseTaskLine(
			'- [ ] Renew passport %%todo:AAMkAGI1%% 📅 2026-10-01',
		);
		expect(parsed?.todoId).toBe('AAMkAGI1');
		expect(parsed?.title).toBe('Renew passport');
	});

	it('flags recurrence and does not treat 🔁 as part of the title', () => {
		const parsed = parseTaskLine('- [ ] Water plants 📅 2026-10-01 🔁 every week');
		expect(parsed?.recurring).toBe(true);
	});

	it('preserves tags and wikilinks as opaque title text', () => {
		const parsed = parseTaskLine(
			'- [ ] Review [[Design doc]] #urgent 📅 2026-10-01',
		);
		expect(parsed?.title).toBe('Review [[Design doc]] #urgent');
	});

	it('is case-insensitive on the checkbox mark', () => {
		expect(parseTaskLine('- [X] Done 📅 2026-10-01')?.checked).toBe(true);
	});

	it('captures leading whitespace for nested tasks and preserves it on render', () => {
		const line = '    - [ ] Nested 📅 2026-10-01';
		const parsed = parseTaskLine(line);
		expect(parsed?.indent).toBe('    ');
		expect(renderTaskLine(parsed!)).toBe(line);
	});
});

describe('renderTaskLine', () => {
	it('renders an open task with a due date', () => {
		expect(
			renderTaskLine({
				indent: '',
				checked: false,
				title: 'Renew passport',
				dueDate: '2026-10-01',
				doneDate: null,
				recurring: false,
				todoId: null,
			}),
		).toBe('- [ ] Renew passport 📅 2026-10-01');
	});

	it('renders a completed task with the marker', () => {
		expect(
			renderTaskLine({
				indent: '',
				checked: true,
				title: 'Renew passport',
				dueDate: '2026-10-01',
				doneDate: '2026-09-30',
				recurring: false,
				todoId: 'AAMkAGI1',
			}),
		).toBe('- [x] Renew passport %%todo:AAMkAGI1%% 📅 2026-10-01 ✅ 2026-09-30');
	});

	it('parses a legacy line with the marker after the dates and re-renders it marker-first', () => {
		const parsed = parseTaskLine('- [x] Renew passport 📅 2026-10-01 ✅ 2026-09-30 %%todo:abc%%');
		expect(parsed).toMatchObject({ title: 'Renew passport', todoId: 'abc', dueDate: '2026-10-01' });
		expect(renderTaskLine(parsed!)).toBe(
			'- [x] Renew passport %%todo:abc%% 📅 2026-10-01 ✅ 2026-09-30',
		);
	});

	it('round-trips parse -> render for a marker-bearing line', () => {
		const line = '- [ ] Review [[Design doc]] #urgent %%todo:abc%% 📅 2026-10-01';
		const parsed = parseTaskLine(line);
		expect(parsed).not.toBeNull();
		expect(renderTaskLine(parsed!)).toBe(line);
	});
});

describe('tasks without a due date', () => {
	it('parses a marker line with no due date', () => {
		const parsed = parseTaskLine('- [ ] Buy milk %%todo:abc%% NO DUE DATE');
		expect(parsed).toMatchObject({ title: 'Buy milk', dueDate: null, todoId: 'abc' });
	});

	it('round-trips a no-due-date line', () => {
		const line = '- [x] Buy milk %%todo:abc%% NO DUE DATE ✅ 2026-09-21';
		const parsed = parseTaskLine(line);
		expect(parsed).toMatchObject({ dueDate: null, doneDate: '2026-09-21', checked: true });
		expect(renderTaskLine(parsed!)).toBe(line);
	});

	it('heals a legacy line with a dangling empty 📅', () => {
		const parsed = parseTaskLine('- [ ] Buy milk %%todo:abc%% 📅 ');
		expect(parsed).toMatchObject({ title: 'Buy milk', dueDate: null });
		expect(renderTaskLine(parsed!)).toBe('- [ ] Buy milk %%todo:abc%% NO DUE DATE');
	});

	it('still ignores a plain checkbox without marker or due date', () => {
		expect(parseTaskLine('- [ ] Buy milk')).toBeNull();
	});
});
