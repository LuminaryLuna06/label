import { useState, useMemo } from 'react';
import {
  Tag,
  Copy,
  Check,
  Download,
  FileSpreadsheet,
  Trash2,
  Sparkles,
  ClipboardPaste,
  ArrowDownUp,
  Settings2,
} from 'lucide-react';
import {
  SAMPLE_DATA,
  parseRawLines,
  generateBlocks,
  copyWithMonospace,
  exportToExcel,
} from './utils/labelGenerator';
import type { GeneratorOptions } from './types';
import './App.css';

export default function App() {
  const [inputText, setInputText] = useState<string>(SAMPLE_DATA);
  const [options, setOptions] = useState<GeneratorOptions>({
    autoSort: true,
    colSeparator: '12spaces',
    alignBrackets: true,
    minPadding: 2,
  });

  const [copiedId, setCopiedId] = useState<string | null>(null);
  const [isCopiedAll, setIsCopiedAll] = useState<boolean>(false);
  const [toastMessage, setToastMessage] = useState<string | null>(null);

  const showToast = (msg: string) => {
    setToastMessage(msg);
    setTimeout(() => {
      setToastMessage(null);
    }, 2200);
  };

  // Tự động phân tích và sinh danh sách nhãn (tính toán phản ứng nhanh)
  const parsedItems = useMemo(() => {
    return parseRawLines(inputText, options.autoSort);
  }, [inputText, options.autoSort]);

  const generatedBlocks = useMemo(() => {
    return generateBlocks(parsedItems, options);
  }, [parsedItems, options]);

  // Sao chép từng nhãn đơn lẻ
  const handleCopySingle = async (idx: string, text: string) => {
    const success = await copyWithMonospace(text);
    if (success) {
      setCopiedId(idx);
      showToast(`✔ Đã copy Nhãn [${idx}] (Ép font Monospace chống lệch)!`);
      setTimeout(() => {
        setCopiedId((prev) => (prev === idx ? null : prev));
      }, 2000);
    }
  };

  // Sao chép tất cả nhãn
  const handleCopyAll = async () => {
    if (!generatedBlocks.length) return;
    const allText = generatedBlocks.map((b) => b.blockText).join('\n\n\n');
    const success = await copyWithMonospace(allText);
    if (success) {
      setIsCopiedAll(true);
      showToast(`✔ Đã copy toàn bộ ${generatedBlocks.length} nhãn vào Clipboard!`);
      setTimeout(() => {
        setIsCopiedAll(false);
      }, 2000);
    }
  };

  // Tải file .txt
  const handleDownloadTxt = () => {
    if (!generatedBlocks.length) return;
    const allText = generatedBlocks.map((b) => b.blockText).join('\n\n\n') + '\n\n\n';
    const blob = new Blob([allText], { type: 'text/plain;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'nhan_cap.txt';
    a.click();
    URL.revokeObjectURL(url);
    showToast('💾 Đã tải về tệp tin nhan_cap.txt');
  };

  // Xuất file Excel .xlsx
  const handleExportExcel = () => {
    if (!parsedItems.length) return;
    exportToExcel(parsedItems, 'nhan_cap_chuan.xlsx');
    showToast('📊 Đã xuất tệp tin Excel nhan_cap_chuan.xlsx');
  };

  // Dán từ clipboard
  const handlePasteClipboard = async () => {
    try {
      const text = await navigator.clipboard.readText();
      setInputText(text);
      showToast('📋 Đã dán dữ liệu từ Clipboard');
    } catch {
      showToast('⚠️ Không thể đọc Clipboard (Vui lòng bấm Ctrl+V)');
    }
  };

  return (
    <div className="app-container">
      {/* HEADER */}
      <header className="app-header">
        <div className="header-brand">
          <div className="brand-icon">
            <Tag size={24} />
          </div>
          <div>
            <div className="header-title-row">
              <h1>Tạo Nhãn Cáp Mạng</h1>
              <span className="tech-badge">React + TypeScript</span>
              <span className="offline-badge">Client-side Pure</span>
            </div>
            <p className="header-subtitle">
              Tự động sắp xếp nhãn tăng dần • Ép cứng font Monospace khi Copy chống lệch format trong
              Word / Docs / Excel
            </p>
          </div>
        </div>
      </header>

      {/* MAIN TWO-COLUMN LAYOUT */}
      <main className="app-main">
        {/* KHỐI TRÁI: INPUT & OPTIONS */}
        <section className="pane pane-left">
          <div className="pane-header">
            <span className="pane-title">📥 Dữ liệu nguồn (Input)</span>
            <div className="btn-toolbar">
              <button
                type="button"
                className="btn-tool btn-primary-tool"
                onClick={handlePasteClipboard}
                title="Dán từ Clipboard"
              >
                <ClipboardPaste size={14} /> Dán
              </button>
              <button
                type="button"
                className="btn-tool"
                onClick={() => setInputText(SAMPLE_DATA)}
                title="Nạp dữ liệu mẫu"
              >
                <Sparkles size={14} /> Mẫu thử
              </button>
              <button
                type="button"
                className="btn-tool btn-danger-tool"
                onClick={() => setInputText('')}
                title="Xóa trắng"
              >
                <Trash2 size={14} /> Xóa
              </button>
            </div>
          </div>

          <textarea
            className="input-textarea"
            value={inputText}
            onChange={(e) => setInputText(e.target.value)}
            placeholder="Dán dữ liệu từ file Excel (định dạng Fr: ... [số] To: ... [số])..."
            spellCheck={false}
          />

          {/* CẤU HÌNH & TÙY CHỌN */}
          <div className="options-card">
            <div className="options-header">
              <Settings2 size={16} />
              <span>Cấu hình & Tự động sắp xếp</span>
            </div>

            {/* Checkbox tự động sắp xếp */}
            <label className="checkbox-row">
              <input
                type="checkbox"
                checked={options.autoSort}
                onChange={(e) => setOptions({ ...options, autoSort: e.target.checked })}
              />
              <span className="checkbox-label">
                <ArrowDownUp size={15} className="option-icon" />
                <strong>Tự động sắp xếp nhãn tăng dần</strong> (1 → 2 → 5 → 10...)
              </span>
            </label>

            {/* Chọn khoảng cách cột */}
            <div className="radio-group-container">
              <span className="radio-group-title">Khoảng cách 2 cột:</span>
              <div className="radio-options">
                <label className="radio-label">
                  <input
                    type="radio"
                    name="colSep"
                    value="12spaces"
                    checked={options.colSeparator === '12spaces'}
                    onChange={() => setOptions({ ...options, colSeparator: '12spaces' })}
                  />
                  <span>12 Spaces (In ấn/Txt)</span>
                </label>
                <label className="radio-label">
                  <input
                    type="radio"
                    name="colSep"
                    value="tab"
                    checked={options.colSeparator === 'tab'}
                    onChange={() => setOptions({ ...options, colSeparator: 'tab' })}
                  />
                  <span>Tab '\t' (Excel 2 cột)</span>
                </label>
                <label className="radio-label">
                  <input
                    type="radio"
                    name="colSep"
                    value="8spaces"
                    checked={options.colSeparator === '8spaces'}
                    onChange={() => setOptions({ ...options, colSeparator: '8spaces' })}
                  />
                  <span>8 Spaces</span>
                </label>
              </div>
            </div>

            {/* Chọn căn lề [số] */}
            <div className="radio-group-container">
              <span className="radio-group-title">Căn lề [số]:</span>
              <div className="radio-options">
                <label className="radio-label">
                  <input
                    type="radio"
                    name="alignBrackets"
                    checked={options.alignBrackets}
                    onChange={() => setOptions({ ...options, alignBrackets: true })}
                  />
                  <span>Căn thẳng cột (main.py)</span>
                </label>
                <label className="radio-label">
                  <input
                    type="radio"
                    name="alignBrackets"
                    checked={!options.alignBrackets}
                    onChange={() => setOptions({ ...options, alignBrackets: false })}
                  />
                  <span>Cố định 4 spaces</span>
                </label>
              </div>
            </div>
          </div>
        </section>

        {/* KHỐI PHẢI: OUTPUT & ACTIONS */}
        <section className="pane pane-right">
          <div className="pane-header pane-header-output">
            <div className="output-status">
              <span className="pane-title">📤 Danh sách nhãn</span>
              <span className={`count-pill ${generatedBlocks.length ? 'count-active' : ''}`}>
                {generatedBlocks.length} nhãn
              </span>
              {options.autoSort && generatedBlocks.length > 0 && (
                <span className="sort-pill">Đã sắp xếp 1 → n</span>
              )}
            </div>

            {generatedBlocks.length > 0 && (
              <div className="btn-toolbar btn-toolbar-actions">
                <button
                  type="button"
                  className={`btn-action ${isCopiedAll ? 'btn-copied' : 'btn-primary-action'}`}
                  onClick={handleCopyAll}
                >
                  {isCopiedAll ? <Check size={14} /> : <Copy size={14} />}
                  {isCopiedAll ? '✔ Đã copy tất cả' : 'Copy tất cả'}
                </button>
                <button type="button" className="btn-action" onClick={handleDownloadTxt}>
                  <Download size={14} /> .txt
                </button>
                <button
                  type="button"
                  className="btn-action btn-excel-action"
                  onClick={handleExportExcel}
                >
                  <FileSpreadsheet size={14} /> Excel (.xlsx)
                </button>
              </div>
            )}
          </div>

          <div className="labels-scroll-container">
            {generatedBlocks.length === 0 ? (
              <div className="empty-placeholder">
                <div className="empty-icon">🏷️</div>
                <h3>Chưa có nhãn nào được tạo</h3>
                <p>
                  Hãy dán dữ liệu vào ô bên trái hoặc bấm <strong>"Mẫu thử"</strong> để bắt đầu.
                </p>
              </div>
            ) : (
              generatedBlocks.map((item) => {
                const isCopied = copiedId === item.idx;
                return (
                  <article key={item.idx} className="label-card">
                    <header className="card-top">
                      <div className="card-idx">
                        <Tag size={14} className="tag-icon" />
                        <span>Nhãn [{item.idx}]</span>
                      </div>
                      <button
                        type="button"
                        className={`card-copy-btn ${isCopied ? 'copied' : ''}`}
                        onClick={() => handleCopySingle(item.idx, item.blockText)}
                      >
                        {isCopied ? <Check size={13} /> : <Copy size={13} />}
                        <span>{isCopied ? '✔ Đã copy' : 'Copy'}</span>
                      </button>
                    </header>
                    <pre className="code-preview">{item.blockText}</pre>
                  </article>
                );
              })
            )}
          </div>
        </section>
      </main>

      {/* TOAST NOTIFICATION */}
      {toastMessage && <div className="toast-float">{toastMessage}</div>}
    </div>
  );
}
