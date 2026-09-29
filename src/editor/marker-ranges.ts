import { MARKER_RE } from '../obsidian-tasks/parser';

const TASK_LINE_RE = /^\s*-\s\[[ xX]\]\s/;

export interface MarkerRange {
	from: number;
	to: number;
}

/** Document ranges of `%%todo:id%%` markers on a task line; empty for non-task lines. */
export function findMarkerRanges(lineText: string, lineFrom: number): MarkerRange[] {
	if (!TASK_LINE_RE.test(lineText)) return [];
	const re = new RegExp(MARKER_RE.source, 'g');
	const ranges: MarkerRange[] = [];
	for (let m = re.exec(lineText); m; m = re.exec(lineText)) {
		ranges.push({ from: lineFrom + m.index, to: lineFrom + m.index + m[0].length });
	}
	return ranges;
}
