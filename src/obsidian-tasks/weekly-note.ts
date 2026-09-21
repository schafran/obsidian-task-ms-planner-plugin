const HEADING = '### Other Action Items';
const PLACEHOLDER = '- [ ] ✅';

export function getIsoWeekNotePath(date: Date): string {
	const d = new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate()));
	const dayNum = (d.getUTCDay() + 6) % 7; // Monday = 0
	d.setUTCDate(d.getUTCDate() - dayNum + 3); // move to this ISO week's Thursday
	const isoYear = d.getUTCFullYear();
	const firstThursday = new Date(Date.UTC(isoYear, 0, 4));
	const firstDayNum = (firstThursday.getUTCDay() + 6) % 7;
	firstThursday.setUTCDate(firstThursday.getUTCDate() - firstDayNum + 3);
	const week =
		1 + Math.round((d.getTime() - firstThursday.getTime()) / (7 * 24 * 3600 * 1000));
	return `Calendar/Weekly/${isoYear}-W${String(week).padStart(2, '0')}.md`;
}

export function insertActionItem(content: string, newLine: string): string {
	const lines = content.split('\n');
	const headingIndex = lines.findIndex((l) => l.trim() === HEADING);
	if (headingIndex === -1) {
		throw new Error(`"${HEADING}" heading not found in note`);
	}

	let sectionEnd = lines.length;
	for (let i = headingIndex + 1; i < lines.length; i++) {
		if (lines[i]!.trim().startsWith('#')) {
			sectionEnd = i;
			break;
		}
	}

	const placeholderIndex = lines
		.slice(headingIndex + 1, sectionEnd)
		.findIndex((l) => l.trim() === PLACEHOLDER);

	if (placeholderIndex !== -1) {
		lines[headingIndex + 1 + placeholderIndex] = newLine;
		return lines.join('\n');
	}

	let lastContentIndex = headingIndex;
	for (let i = headingIndex + 1; i < sectionEnd; i++) {
		if (lines[i]!.trim() !== '') lastContentIndex = i;
	}
	lines.splice(lastContentIndex + 1, 0, newLine);
	return lines.join('\n');
}
