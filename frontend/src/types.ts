export type LabelMode = 'tape24mm' | 'cable';

export interface ParsedLabel {
  fr: string;
  idx: string;
  to: string;
}

export interface ParsedTape24mmLabel {
  left: string;
  right: string;
  idx: string;
}

export type ParsedItem = ParsedLabel | ParsedTape24mmLabel;

export interface GeneratedBlock {
  idx: string;
  blockText: string;
  parsed?: ParsedItem;
}

export interface GeneratorOptions {
  mode: LabelMode;
  // Cable options
  colSeparator: '12spaces' | 'tab' | '8spaces';
  alignBrackets: boolean;
  minPadding: number;
  // 24mm tape options
  tapeSpaces: number;
  tapeMiddleBlankLines: number;
  // Shared options
  autoSort: boolean;
}
