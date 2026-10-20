import { useState, useMemo, useEffect } from 'react';
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
  SlidersHorizontal,
} from 'lucide-react';
import {
  SAMPLE_DATA_CABLE,
  SAMPLE_DATA_24MM,
  parseCableLines,
  parseTape24mmLines,
  generateBlocks,
  copyWithMonospace,
  exportToExcel,
} from './utils/labelGenerator';
import type { GeneratorOptions, LabelMode } from './types';
import './App.css';

const STORAGE_KEY_INPUT_24MM = 'cable_label_input_24mm';
const STORAGE_KEY_INPUT_CABLE = 'cable_label_input_cable';
const STORAGE_KEY_OPTIONS = 'cable_label_options';

export default function App() {
  const [options, setOptions] = useState<GeneratorOptions>(() => {
    try {
      const savedOpts = localStorage.getItem(STORAGE_KEY_OPTIONS);
      if (savedOpts) {
        const parsed = JSON.parse(savedOpts);
        return {
          mode: parsed.mode || 'tape24mm',
          colSeparator: parsed.colSeparator || '12spaces',
          alignBrackets: parsed.alignBrackets ?? true,
          minPadding: parsed.minPadding ?? 2,
          tapeSpaces: parsed.tapeSpaces ?? 32,
          tapeMiddleBlankLines: parsed.tapeMiddleBlankLines ?? 1,
          autoSort: parsed.autoSort ?? true,
        };
      }
    } catch (e) {
      console.warn('Lỗi đọc options từ localStorage', e);
    }
    return {
      mode: 'tape24mm',
      colSeparator: '12spaces',
      alignBrackets: true,
      minPadding: 2,
      tapeSpaces: 32,
      tapeMiddleBlankLines: 1,
      autoSort: true,
    };
  });

  const [input24mm, setInput24mm] = useState<string>(() => {
    try {
      const saved = localStorage.getItem(STORAGE_KEY_INPUT_24MM);
      if (saved !== null) return saved;
      return SAMPLE_DATA_24MM;
    } catch {
      return SAMPLE_DATA_24MM;
    }
  });

  const [inputCable, setInputCable] = useState<string>(() => {
    try {
      const saved = localStorage.getItem(STORAGE_KEY_INPUT_CABLE);
      if (saved !== null) return saved;
      const legacy = localStorage.getItem('cable_label_source_input');
      if (legacy !== null && legacy.includes('Fr:')) return legacy;
      return SAMPLE_DATA_CABLE;
    } catch {
      return SAMPLE_DATA_CABLE;
    }
  });

  // Tự động lưu input 24mm vào localStorage
  useEffect(() => {
    try {
      localStorage.setItem(STORAGE_KEY_INPUT_24MM, input24mm);
    } catch (e) {
      console.error('Không thể lưu input 24mm', e);
    }
  }, [input24mm]);

  // Tự động lưu input Cable vào localStorage
  useEffect(() => {
    try {
      localStorage.setItem(STORAGE_KEY_INPUT_CABLE, inputCable);
    } catch (e) {
      console.error('Không thể lưu input cable', e);
    }
  }, [inputCable]);

  // Tự động đồng bộ tùy chọn cấu hình vào LocalStorage
  useEffect(() => {
    try {
      localStorage.setItem(STORAGE_KEY_OPTIONS, JSON.stringify(options));
    } catch (e) {
      console.error('Không thể lưu tùy chọn vào localStorage', e);
    }
  }, [options]);

  const [copiedId, setCopiedId] = useState<string | null>(null);
  const [isCopiedAll, setIsCopiedAll] = useState<boolean>(false);
  const [toastMessage, setToastMessage] = useState<string | null>(null);

  const showToast = (msg: string) => {
    setToastMessage(msg);
    setTimeout(() => {
      setToastMessage(null);
    }, 2200);
  };

  const currentInput = options.mode === 'tape24mm' ? input24mm : inputCable;
  const setCurrentInput = options.mode === 'tape24mm' ? setInput24mm : setInputCable;

  // Nạp dữ liệu mẫu tương ứng với chế độ đang chọn
  const handleLoadSample = () => {
    if (options.mode === 'tape24mm') {
      setInput24mm(SAMPLE_DATA_24MM);
    } else {
      setInputCable(SAMPLE_DATA_CABLE);
    }
    showToast('✨ Đã nạp dữ liệu mẫu thử');
  };

  // Xóa trắng dữ liệu nguồn và xóa khỏi LocalStorage
  const handleClearInput = () => {
    setCurrentInput('');
    try {
      if (options.mode === 'tape24mm') {
        localStorage.removeItem(STORAGE_KEY_INPUT_24MM);
      } else {
        localStorage.removeItem(STORAGE_KEY_INPUT_CABLE);
      }
    } catch (e) {
      console.warn('Lỗi xóa input từ localStorage', e);
    }
    showToast('🗑️ Đã xóa trắng dữ liệu');
  };

  // Chuyển đổi qua lại giữa các loại nhãn
  const handleSwitchMode = (newMode: LabelMode) => {
    setOptions((prev) => ({ ...prev, mode: newMode }));
    showToast(
      newMode === 'tape24mm'
        ? '🏷️ Chuyển sang Nhãn 2 cột khổ 24mm (32 khoảng trắng)'
        : '🔀 Chuyển sang Nhãn cáp 2 chiều (Fr / To)'
    );
  };

  // Tự động phân tích và sinh danh sách nhãn
  const parsedItems = useMemo(() => {
    if (options.mode === 'tape24mm') {
      return parseTape24mmLines(currentInput, options.autoSort);
    }
    return parseCableLines(currentInput, options.autoSort);
  }, [currentInput, options.mode, options.autoSort]);

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
    a.download = options.mode === 'tape24mm' ? 'nhan_24mm.txt' : 'nhan_cap.txt';
    a.click();
    URL.revokeObjectURL(url);
    showToast(`💾 Đã tải về tệp tin ${a.download}`);
  };

  // Xuất file Excel .xlsx
  const handleExportExcel = () => {
    if (!parsedItems.length) return;
    const filename =
      options.mode === 'tape24mm' ? 'nhan_24mm_chuan.xlsx' : 'nhan_cap_chuan.xlsx';
    exportToExcel(parsedItems, options.mode, filename);
    showToast(`📊 Đã xuất tệp tin Excel ${filename}`);
  };

  // Dán từ clipboard
  const handlePasteClipboard = async () => {
    try {
      const text = await navigator.clipboard.readText();
      setCurrentInput(text);
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
              <h1>Hệ Thống Tạo Nhãn Cáp & Thiết Bị</h1>
              <span className="tech-badge">React 19 + TypeScript</span>
              <span className="offline-badge">Client-side Pure</span>
            </div>
            <p className="header-subtitle">
              Hỗ trợ nhãn 2 cột khổ băng 24mm (32 khoảng trắng) & nhãn cáp 2 chiều (Fr/To) • Ép font Monospace chống lệch định dạng
            </p>
          </div>
        </div>
      </header>

      {/* CHỌN LOẠI NHÃN (MODE TABS) */}
      <nav className="mode-nav">
        <button
          type="button"
          className={`mode-tab-btn ${options.mode === 'tape24mm' ? 'active' : ''}`}
          onClick={() => handleSwitchMode('tape24mm')}
        >
          <SlidersHorizontal size={16} />
          <span>Nhãn 2 Cột Khổ Băng 24mm (32 Khoảng Trắng)</span>
          <span className="mode-badge-new">Mới</span>
        </button>
        <button
          type="button"
          className={`mode-tab-btn ${options.mode === 'cable' ? 'active' : ''}`}
          onClick={() => handleSwitchMode('cable')}
        >
          <Tag size={16} />
          <span>Nhãn Cáp 2 Chiều (Fr / To - 4 Dòng)</span>
        </button>
      </nav>

      {/* MAIN TWO-COLUMN LAYOUT */}
      <main className="app-main">
        {/* KHỐI TRÁI: INPUT & OPTIONS */}
        <section className="pane pane-left">
          <div className="pane-header">
            <span className="pane-title">
              📥 Dữ liệu nguồn ({options.mode === 'tape24mm' ? 'Khổ 24mm' : 'Fr / To'})
            </span>
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
                onClick={handleLoadSample}
                title="Nạp dữ liệu mẫu"
              >
                <Sparkles size={14} /> Mẫu thử
              </button>
              <button
                type="button"
                className="btn-tool btn-danger-tool"
                onClick={handleClearInput}
                title="Xóa trắng"
              >
                <Trash2 size={14} /> Xóa
              </button>
            </div>
          </div>

          <textarea
            className="input-textarea"
            value={currentInput}
            onChange={(e) => setCurrentInput(e.target.value)}
            placeholder={
              options.mode === 'tape24mm'
                ? 'Dán dữ liệu nhãn 2 cột (Ví dụ: E01        E01 hoặc 2 cột từ Excel)...'
                : 'Dán dữ liệu từ file Excel (định dạng Fr: ... [số] To: ... [số])...'
            }
            spellCheck={false}
          />

          {/* CẤU HÌNH & TÙY CHỌN */}
          <div className="options-card">
            <div className="options-header">
              <Settings2 size={16} />
              <span>
                Cấu hình tùy chọn •{' '}
                {options.mode === 'tape24mm' ? 'Khổ băng 24mm' : 'Nhãn cáp 2 chiều'}
              </span>
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
                <strong>Tự động sắp xếp nhãn tăng dần</strong>{' '}
                {options.mode === 'tape24mm' ? '(E01 → E02 → E10...)' : '(1 → 2 → 5 → 10...)'}
              </span>
            </label>

            {options.mode === 'tape24mm' ? (
              /* TÙY CHỌN RIÊNG CHO KHỔ BĂNG 24MM */
              <>
                {/* Chọn khoảng cách cột (mặc định chuẩn 32 khoảng trắng) */}
                <div className="radio-group-container">
                  <span className="radio-group-title">Khoảng cách giữa 2 cột:</span>
                  <div className="radio-options">
                    <label className="radio-label">
                      <input
                        type="radio"
                        name="tapeSpaces"
                        checked={options.tapeSpaces === 32}
                        onChange={() => setOptions({ ...options, tapeSpaces: 32 })}
                      />
                      <strong>32 Spaces (Chuẩn 24mm)</strong>
                    </label>
                    <label className="radio-label">
                      <input
                        type="radio"
                        name="tapeSpaces"
                        checked={options.tapeSpaces === 24}
                        onChange={() => setOptions({ ...options, tapeSpaces: 24 })}
                      />
                      <span>24 Spaces</span>
                    </label>
                    <label className="radio-label">
                      <input
                        type="radio"
                        name="tapeSpaces"
                        checked={options.tapeSpaces === 40}
                        onChange={() => setOptions({ ...options, tapeSpaces: 40 })}
                      />
                      <span>40 Spaces</span>
                    </label>
                  </div>

                  <div className="custom-spaces-row">
                    <span className="field-label">Tùy chỉnh:</span>
                    <input
                      type="number"
                      min={0}
                      max={120}
                      className="input-number-mini"
                      value={options.tapeSpaces}
                      onChange={(e) => {
                        const val = parseInt(e.target.value, 10);
                        if (!isNaN(val) && val >= 0) {
                          setOptions({ ...options, tapeSpaces: val });
                        }
                      }}
                    />
                    <span className="field-unit">khoảng trắng</span>
                  </div>
                </div>

                {/* Phân bổ chiều dọc: 1 dòng trống ở giữa */}
                <div className="radio-group-container">
                  <span className="radio-group-title">Phân bổ chiều dọc khổ băng:</span>
                  <div className="radio-options">
                    <label className="radio-label">
                      <input
                        type="radio"
                        name="tapeMiddleBlankLines"
                        checked={options.tapeMiddleBlankLines === 1}
                        onChange={() => setOptions({ ...options, tapeMiddleBlankLines: 1 })}
                      />
                      <strong>1 dòng trống ở giữa (Chuẩn phân bổ 24mm)</strong>
                    </label>
                    <label className="radio-label">
                      <input
                        type="radio"
                        name="tapeMiddleBlankLines"
                        checked={options.tapeMiddleBlankLines === 0}
                        onChange={() => setOptions({ ...options, tapeMiddleBlankLines: 0 })}
                      />
                      <span>Không có dòng trống</span>
                    </label>
                    <label className="radio-label">
                      <input
                        type="radio"
                        name="tapeMiddleBlankLines"
                        checked={options.tapeMiddleBlankLines === 2}
                        onChange={() => setOptions({ ...options, tapeMiddleBlankLines: 2 })}
                      />
                      <span>2 dòng trống</span>
                    </label>
                  </div>
                </div>
              </>
            ) : (
              /* TÙY CHỌN CHO NHÃN CÁP 2 CHIỀU (FR/TO) */
              <>
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
              </>
            )}
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
              {options.mode === 'tape24mm' && (
                <span className="format-pill">Khổ 24mm • {options.tapeSpaces} spaces</span>
              )}
              {options.autoSort && generatedBlocks.length > 0 && (
                <span className="sort-pill">Đã sắp xếp tự nhiên</span>
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
