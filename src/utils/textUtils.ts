export function getTrimmedText(text: string, trimTo: number): string {
    if (text && text.length > trimTo) {
        return `${text.substring(0, trimTo)}...`;
    }

    return text;
}


export function getTextLinesCount(text: string): number {
    return text.split("\n").length;
}

export interface ITextPositionRange {
    startRow: number;
    startColumn: number;
    endRow: number;
    endChar: number;
}

export function getTextPositionRange(text: string, pattern: string): ITextPositionRange {
    const firstIndex = text.indexOf(pattern);
    const before = text.substring(0, firstIndex - 1);
    const beforeUntilCurrentRow = before.substring(0, before.lastIndexOf("\n"));
    const range = { startRow: 0, startColumn: 0, endRow: 0, endChar: 0 };
    range.startRow = getTextLinesCount(beforeUntilCurrentRow)+1;
    range.startColumn = firstIndex - before.length;
    const matchLinesCount = getTextLinesCount(pattern);
    range.endRow = range.startRow + matchLinesCount;
    // range.endChar = range.
    return range;
}
