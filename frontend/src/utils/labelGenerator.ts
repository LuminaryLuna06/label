import * as XLSX from 'xlsx';
import type { ParsedLabel, GeneratedBlock, GeneratorOptions } from '../types';

export const SAMPLE_DATA = `Fr: S4_U13-14_SRV số 1_Slot 4_Port 1    [5]\tTo: S4_U41_ODF_port 1->N2_U13_ODF_port 1    [5]
Fr: S4_U13-14_SRV số 1_Slot 4_Port 2    [2]\tTo: S4_U41_ODF_port 2->N2_U13_ODF_port 2    [2]
Fr: S4_U13-14_SRV số 1_iDrac    [10]\tTo: S4_U40_PP_Port 5->N1_U3_PP_Port 5    [10]
Fr: N2_U13_ODF_port 1    [1]\tTo: N3_U1-13_CoreSw01_Slot 10_port 1    [1]
Fr: N2_U13_ODF_port 2    [7]\tTo: N4_U1-13_CoreSw02_Slot 10_port 1    [7]
Fr: N1_U3_PP_Port 5    [3]\tTo: N1_U20_SWMgmt_port 18    [3]`;

/**
 * Phân tích dữ liệu thô từ Excel thành danh sách đối tượng ParsedLabel.
 * Tự động chuẩn hóa Unicode NFC và hỗ trợ tự động sắp xếp theo thứ tự số tăng dần (1, 2, 5, 10...).
 */
export function parseRawLines(rawText: string, autoSort = true): ParsedLabel[] {
  // Chuẩn hóa Unicode NFC để tránh lệch độ dài ký tự tiếng Việt
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

  // Tự động sắp xếp nhãn tăng dần tự nhiên (Natural Sort: 1, 2, 10 thay vì 1, 10, 2)
  if (autoSort) {
    parsed.sort((a, b) =>
      a.idx.localeCompare(b.idx, undefined, { numeric: true, sensitivity: 'base' })
    );
  }

  return parsed;
}

/**
 * Sinh chuỗi 4 dòng định dạng nhãn cáp 2 chiều với căn lề Monospace chuẩn xác.
 */
export function generateBlocks(
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
 * Sao chép đa định dạng ép cứng font Monospace (chuẩn ClipboardItem theo đúng yêu cầu).
 */
export async function copyWithMonospace(formattedText: string): Promise<boolean> {
  // Bọc chuỗi kết quả trong thẻ <pre> và ép cứng font monospace
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
    // Fallback nếu trình duyệt chưa cấp quyền ClipboardItem đầy đủ
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
 * Xuất danh sách nhãn ra tệp tin Excel (.xlsx) chuẩn cấu trúc từng cột, tránh hoàn toàn lỗi lệch font.
 */
export function exportToExcel(parsedList: ParsedLabel[], filename = 'nhan_cap_chuan.xlsx') {
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

  // Đặt độ rộng các cột tự động đẹp mắt
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
