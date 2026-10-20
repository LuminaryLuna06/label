import * as XLSX from 'xlsx';
import type {
  ParsedLabel,
  ParsedTape24mmLabel,
  ParsedItem,
  GeneratedBlock,
  GeneratorOptions,
  LabelMode,
} from '../types';

export const SAMPLE_DATA_CABLE = `Fr: S4_U13-14_SRV số 1_Slot 4_Port 1    [5]\tTo: S4_U41_ODF_port 1->N2_U13_ODF_port 1    [5]
Fr: S4_U13-14_SRV số 1_Slot 4_Port 2    [2]\tTo: S4_U41_ODF_port 2->N2_U13_ODF_port 2    [2]
Fr: S4_U13-14_SRV số 1_iDrac    [10]\tTo: S4_U40_PP_Port 5->N1_U3_PP_Port 5    [10]
Fr: N2_U13_ODF_port 1    [1]\tTo: N3_U1-13_CoreSw01_Slot 10_port 1    [1]
Fr: N2_U13_ODF_port 2    [7]\tTo: N4_U1-13_CoreSw02_Slot 10_port 1    [7]
Fr: N1_U3_PP_Port 5    [3]\tTo: N1_U20_SWMgmt_port 18    [3]`;

// Dữ liệu mẫu chuẩn cho loại nhãn 2 cột khổ băng 24mm
export const SAMPLE_DATA_24MM = `E01        E01
E02        E02
E03        E03
E04        E04
E05        E05
E10        E10`;

export const SAMPLE_DATA = SAMPLE_DATA_CABLE;

/**
 * Phân tích dữ liệu nhãn cáp 2 chiều (Fr/To) từ text/Excel.
 */
export function parseCableLines(rawText: string, autoSort = true): ParsedLabel[] {
  const normalized = rawText
    .normalize('NFC')
    .replace(/\u00a0/g, ' ')
    .replace(/\u200b/g, '')
    .replace(/\u3000/g, ' ');

  const lines = normalized.split('\n');
  const pattern = /Fr:\s*(.+?)\s*\[([^\]]+)\]\s+To:\s*(.+?)(?:\s*\[([^\]]+)\])?$/i;

  const parsed: ParsedLabel[] = [];

  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed) continue;

    const match = pattern.exec(trimmed);
    if (match) {
      const fr = match[1].trim();
      const idx = match[2].trim();
      const to = match[3].trim();
      parsed.push({ fr, idx, to });
    }
  }

  if (autoSort) {
    parsed.sort((a, b) =>
      a.idx.localeCompare(b.idx, undefined, { numeric: true, sensitivity: 'base' })
    );
  }

  return parsed;
}

export const parseRawLines = parseCableLines;

/**
 * Phân tích dữ liệu 2 cột cho khổ băng in 24mm (ví dụ: E01        E01).
 * Hỗ trợ phân cách bằng Tab, 2+ khoảng trắng, dấu |, hoặc phẩy.
 * Nếu một dòng chỉ chứa 1 giá trị, tự động nhân đôi sang cột 2 để in 2 bên.
 */
export function parseTape24mmLines(rawText: string, autoSort = true): ParsedTape24mmLabel[] {
  const normalized = rawText
    .normalize('NFC')
    .replace(/\u00a0/g, ' ')
    .replace(/\u200b/g, '')
    .replace(/\u3000/g, ' ');

  const lines = normalized.split('\n');
  const parsed: ParsedTape24mmLabel[] = [];

  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed) continue;

    let left = '';
    let right = '';

    if (trimmed.includes('\t')) {
      const parts = trimmed.split('\t').map((s) => s.trim()).filter(Boolean);
      left = parts[0] || '';
      right = parts.length > 1 ? parts[1] : left;
    } else if (/\s{2,}/.test(trimmed)) {
      const parts = trimmed.split(/\s{2,}/).map((s) => s.trim()).filter(Boolean);
      left = parts[0] || '';
      right = parts.length > 1 ? parts[1] : left;
    } else if (trimmed.includes('|')) {
      const parts = trimmed.split('|').map((s) => s.trim()).filter(Boolean);
      left = parts[0] || '';
      right = parts.length > 1 ? parts[1] : left;
    } else if (trimmed.includes(';')) {
      const parts = trimmed.split(';').map((s) => s.trim()).filter(Boolean);
      left = parts[0] || '';
      right = parts.length > 1 ? parts[1] : left;
    } else if (trimmed.includes(',')) {
      const parts = trimmed.split(',').map((s) => s.trim()).filter(Boolean);
      left = parts[0] || '';
      right = parts.length > 1 ? parts[1] : left;
    } else {
      const parts = trimmed.split(/\s+/).map((s) => s.trim()).filter(Boolean);
      if (parts.length >= 2) {
        left = parts[0];
        right = parts[1];
      } else {
        left = trimmed;
        right = trimmed;
      }
    }

    if (!left && !right) continue;
    if (!right) right = left;

    const idx = left === right ? left : `${left} - ${right}`;
    parsed.push({ left, right, idx });
  }

  if (autoSort) {
    parsed.sort((a, b) =>
      a.idx.localeCompare(b.idx, undefined, { numeric: true, sensitivity: 'base' })
    );
  }

  return parsed;
}

/**
 * Sinh chuỗi nhãn cáp 2 chiều (Fr/To) 4 dòng với căn lề Monospace.
 */
export function generateCableBlocks(
  parsedList: ParsedLabel[],
  options: GeneratorOptions
): GeneratedBlock[] {
  let colSep = ' '.repeat(12);
  if (options.colSeparator === 'tab') {
    colSep = '\t';
  } else if (options.colSeparator === '8spaces') {
    colSep = ' '.repeat(8);
  }

  return parsedList.map((item) => {
    const f1 = `Fr: ${item.fr}`;
    const t1 = `To: ${item.to}`;
    const f2 = `Fr: ${item.to}`;
    const t2 = `To: ${item.fr}`;

    const maxLen = Math.max(f1.length, t1.length, f2.length, t2.length);

    const formatRow = (leftText: string) => {
      let padding = ' '.repeat(4);
      if (options.alignBrackets) {
        padding = ' '.repeat(maxLen - leftText.length + options.minPadding);
      }
      const singleCol = `${leftText}${padding}[${item.idx}]`;
      return `${singleCol}${colSep}${singleCol}`;
    };

    const blockText = `${formatRow(f1)}\n${formatRow(t1)}\n\n${formatRow(f2)}\n${formatRow(t2)}`;

    return {
      idx: item.idx,
      blockText,
      parsed: item,
    };
  });
}

/**
 * Sinh chuỗi nhãn 2 cột khổ băng 24mm:
 * - Điều chỉnh chính xác số khoảng trắng (mặc định chuẩn 32 khoảng trắng) giữa 2 cột.
 * - Giữ 1 dòng trống ở giữa để phân bổ đều chiều dọc trên khổ băng 24mm.
 *
 * Cấu trúc đầu ra:
 * E01                                E01
 * 
 * E01                                E01
 */
export function generateTape24mmBlocks(
  parsedList: ParsedTape24mmLabel[],
  options: GeneratorOptions
): GeneratedBlock[] {
  const spaces = options.tapeSpaces ?? 32;
  const middleBlanks = options.tapeMiddleBlankLines ?? 1;

  const spacer = ' '.repeat(Math.max(0, spaces));
  const middle = '\n'.repeat(Math.max(0, middleBlanks) + 1);

  return parsedList.map((item) => {
    const row = `${item.left}${spacer}${item.right}`;
    const blockText = `${row}${middle}${row}`;

    return {
      idx: item.idx,
      blockText,
      parsed: item,
    };
  });
}

/**
 * Sinh chuỗi nhãn theo chế độ đang chọn (cable hoặc tape24mm).
 */
export function generateBlocks(
  parsedList: ParsedItem[],
  options: GeneratorOptions
): GeneratedBlock[] {
  if (options.mode === 'tape24mm') {
    return generateTape24mmBlocks(parsedList as ParsedTape24mmLabel[], options);
  }
  return generateCableBlocks(parsedList as ParsedLabel[], options);
}

/**
 * Sao chép đa định dạng ép cứng font Monospace (chuẩn ClipboardItem).
 */
export async function copyWithMonospace(formattedText: string): Promise<boolean> {
  const htmlContent = `<pre style="font-family: 'SF Mono', Consolas, 'Courier New', monospace; font-size: 13px; margin: 0;">${formattedText}</pre>`;

  try {
    const blobText = new Blob([formattedText], { type: 'text/plain' });
    const blobHtml = new Blob([htmlContent], { type: 'text/html' });

    const clipboardItem = new ClipboardItem({
      'text/plain': blobText,
      'text/html': blobHtml,
    });

    await navigator.clipboard.write([clipboardItem]);
    return true;
  } catch {
    try {
      const textArea = document.createElement('textarea');
      textArea.value = formattedText;
      textArea.style.position = 'fixed';
      textArea.style.left = '-999999px';
      document.body.appendChild(textArea);
      textArea.focus();
      textArea.select();
      document.execCommand('copy');
      textArea.remove();
      return true;
    } catch {
      return false;
    }
  }
}

/**
 * Xuất danh sách nhãn cáp ra tệp tin Excel (.xlsx) chuẩn cấu trúc từng cột.
 */
export function exportCableToExcel(parsedList: ParsedLabel[], filename = 'nhan_cap_chuan.xlsx') {
  const wb = XLSX.utils.book_new();
  const rows: (string | number)[][] = [];

  for (const item of parsedList) {
    const f1 = `Fr: ${item.fr}`;
    const t1 = `To: ${item.to}`;
    const f2 = `Fr: ${item.to}`;
    const t2 = `To: ${item.fr}`;
    const idVal = `[${item.idx}]`;

    // Khối Đầu 1
    rows.push([f1, idVal, '', f1, idVal]);
    rows.push([t1, idVal, '', t1, idVal]);
    rows.push([]); // dòng trống

    // Khối Đầu 2 (Đảo)
    rows.push([f2, idVal, '', f2, idVal]);
    rows.push([t2, idVal, '', t2, idVal]);
    rows.push([], []); // khoảng cách giữa các nhãn
  }

  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws['!cols'] = [
    { wch: 48 }, // Cột A: Tên cổng 1
    { wch: 10 }, // Cột B: [Số]
    { wch: 4 },  // Cột C: Khoảng trống
    { wch: 48 }, // Cột D: Tên cổng 2
    { wch: 10 }, // Cột E: [Số]
  ];

  XLSX.utils.book_append_sheet(wb, ws, 'Nhan_Cap');
  XLSX.writeFile(wb, filename);
}

/**
 * Xuất danh sách nhãn 2 cột khổ 24mm ra tệp tin Excel (.xlsx).
 */
export function exportTape24mmToExcel(
  parsedList: ParsedTape24mmLabel[],
  filename = 'nhan_24mm_chuan.xlsx'
) {
  const wb = XLSX.utils.book_new();
  const rows: (string | number)[][] = [];

  for (const item of parsedList) {
    // Dòng 1: Cột 1 và Cột 2
    rows.push([item.left, '', item.right]);
    // Dòng 2: Dòng trống ở giữa để phân bổ đều trên khổ 24mm
    rows.push([]);
    // Dòng 3: Cột 1 và Cột 2 lặp lại
    rows.push([item.left, '', item.right]);
    // Khoảng trống cách biệt giữa các nhãn
    rows.push([], []);
  }

  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws['!cols'] = [
    { wch: 28 }, // Cột A: Cột 1
    { wch: 8 },  // Cột B: Khoảng trống
    { wch: 28 }, // Cột C: Cột 2
  ];

  XLSX.utils.book_append_sheet(wb, ws, 'Nhan_24mm');
  XLSX.writeFile(wb, filename);
}

/**
 * Hàm điều phối xuất Excel hỗ trợ cả 2 chế độ nhãn.
 */
export function exportToExcel(
  parsedList: ParsedItem[],
  mode: LabelMode = 'cable',
  filename?: string
) {
  if (mode === 'tape24mm') {
    exportTape24mmToExcel(
      parsedList as ParsedTape24mmLabel[],
      filename || 'nhan_24mm_chuan.xlsx'
    );
  } else {
    exportCableToExcel(
      parsedList as ParsedLabel[],
      filename || 'nhan_cap_chuan.xlsx'
    );
  }
}
