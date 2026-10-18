export interface ParsedLabel {
  fr: string;
  idx: string;
  to: string;
}

export interface GeneratedBlock {
  idx: string;
  blockText: string;
  parsed: ParsedLabel;
}

export interface GeneratorOptions {
  colSeparator: '12spaces' | 'tab' | '8spaces';
  alignBrackets: boolean;
  autoSort: boolean;
  minPadding: number;
}
