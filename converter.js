/* ===== 仕訳帳コンバーター UI制御 ===== */

// --- 状態管理 ---
let currentStep = 1;
let rawHeaders = [];
let rawData = [];
let columnMapping = {};
let conversionRules = {
    dateFormat: 'auto',
    accountMapping: {},
    taxMapping: {},
    fixedValues: {},
};
let convertedData = [];
let validationErrors = [];
let displayedRows = 0;
const ROWS_PER_PAGE = 100;
let selectedCompany = ''; // 現在選択中の会社名
let selectedIndustry = ''; // 現在選択中の業種（会社なしでも設定可能）
let defaultKashiKamoku = ''; // デフォルト貸方勘定科目（空=自動）

// --- HTMLエスケープ ---
function escapeHtml(str) {
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
}

// --- 初期化 ---
document.addEventListener('DOMContentLoaded', function () {
    initConverter();
});

function initConverter() {
    setupDropZone();
    setupKeyboardShortcuts();
    setupPageDrop();
    setupPasteAutoDetect();
    renderDateFormatOptions();
    renderFixedValueUI();
    renderCompanySelect();
    renderIndustrySelect();
    initApiKeyPanel();
}

// ===== キーボードショートカット =====

// ===== Gemini APIキー設定 =====

function initApiKeyPanel() {
    const input = document.getElementById('geminiApiKeyInput');
    const status = document.getElementById('apiKeyStatus');
    const saved = getGeminiApiKey();
    if (saved) {
        input.value = saved;
        status.textContent = '✓ 設定済み';
        status.style.color = 'var(--success)';
    }
}

function toggleApiKeyPanel() {
    const panel = document.getElementById('apiKeyPanel');
    panel.style.display = panel.style.display === 'none' ? '' : 'none';
}

function saveGeminiApiKey() {
    const input = document.getElementById('geminiApiKeyInput');
    const status = document.getElementById('apiKeyStatus');
    const key = input.value.trim();

    setGeminiApiKey(key);

    if (key) {
        status.textContent = '✓ 保存しました';
        status.style.color = 'var(--success)';
        showToast('Gemini APIキーを保存しました（PDF読み込み時にGemini APIを使用します）');
    } else {
        status.textContent = '';
        showToast('APIキーを削除しました（Tesseract.js OCRを使用します）');
    }
}

function toggleApiKeyVisibility() {
    const input = document.getElementById('geminiApiKeyInput');
    input.type = input.type === 'password' ? 'text' : 'password';
}

function setupKeyboardShortcuts() {
    document.addEventListener('keydown', function (e) {
        // 入力中はショートカットを無効化
        const tag = e.target.tagName.toLowerCase();
        if (tag === 'input' || tag === 'textarea' || tag === 'select') return;

        // Ctrl+Enter: 次のステップへ
        if (e.ctrlKey && e.key === 'Enter') {
            e.preventDefault();
            if (currentStep < 4) goToStep(currentStep + 1);
        }

        // Escape: 前のステップへ
        if (e.key === 'Escape') {
            e.preventDefault();
            if (currentStep > 1) goToStep(currentStep - 1);
        }

        // Ctrl+S: CSVダウンロード（ステップ4の場合）
        if (e.ctrlKey && e.key === 's') {
            e.preventDefault();
            if (currentStep === 4 && convertedData.length > 0) {
                exportCSV();
            }
        }

        // Ctrl+C: クリップボードにコピー（ステップ4で何も選択していない場合）
        if (e.ctrlKey && e.key === 'c' && !window.getSelection().toString()) {
            if (currentStep === 4 && convertedData.length > 0) {
                e.preventDefault();
                copyToClipboard();
            }
        }
    });
}

// ===== ステップ制御 =====

function goToStep(stepNum) {
    // 領収書OCRの場合はrawDataなしでStep4に直接遷移可能
    const isReceiptMode = convertedData.length > 0 && convertedData.some(r => r._isReceipt);

    // ステップ1 → 2への遷移はデータが必要（領収書モードは除外）
    if (stepNum >= 2 && rawData.length === 0 && !isReceiptMode) {
        showToast('先にデータを読み込んでください');
        return;
    }

    // ステップ遷移前の処理
    if (stepNum === 2 && currentStep < 2) {
        renderMappingUI();
    }
    if (stepNum === 3) {
        collectMappingFromUI();
        renderRulesUI();
    }
    if (stepNum === 4 && !isReceiptMode) {
        collectRulesFromUI();
        resultFilter = 'all';
        runConversion();
    }

    currentStep = stepNum;

    // 全ステップを非表示にして、現在のステップだけ表示
    document.querySelectorAll('.converter-step').forEach(el => el.classList.remove('active'));
    document.getElementById('step' + stepNum).classList.add('active');

    // インジケーター更新
    updateStepIndicator();

    // ページトップにスクロール
    window.scrollTo(0, 0);
}

function updateStepIndicator() {
    const items = document.querySelectorAll('.step-item');
    items.forEach((item, idx) => {
        const stepNum = idx + 1;
        item.classList.remove('active', 'completed');
        if (stepNum === currentStep) {
            item.classList.add('active');
        } else if (stepNum < currentStep) {
            item.classList.add('completed');
        }
    });
}

// ===== ページ全体ドラッグ&ドロップ =====

function setupPageDrop() {
    let dragCounter = 0;

    document.addEventListener('dragenter', function (e) {
        e.preventDefault();
        dragCounter++;
        if (dragCounter === 1) {
            document.body.classList.add('page-dragover');
        }
    });

    document.addEventListener('dragleave', function (e) {
        e.preventDefault();
        dragCounter--;
        if (dragCounter === 0) {
            document.body.classList.remove('page-dragover');
        }
    });

    document.addEventListener('dragover', function (e) {
        e.preventDefault();
    });

    document.addEventListener('drop', function (e) {
        e.preventDefault();
        dragCounter = 0;
        document.body.classList.remove('page-dragover');

        const file = e.dataTransfer.files[0];
        if (file) {
            // ステップ1以外にいる場合も受け付ける（新しいデータとして処理）
            if (currentStep !== 1) {
                // 確認なしで新しいデータを読み込む
                rawHeaders = [];
                rawData = [];
                columnMapping = {};
                conversionRules = { dateFormat: 'auto', accountMapping: {}, taxMapping: {}, fixedValues: {} };
                convertedData = [];
                validationErrors = [];
                displayedRows = 0;
            }
            processFile(file);
        }
    });
}

// ===== ステップ1: データ入力 =====

// 入力タブ切替
function switchInputTab(tab) {
    document.querySelectorAll('.input-tab').forEach(el => el.classList.remove('active'));
    document.querySelectorAll('.input-panel').forEach(el => el.classList.remove('active'));

    if (tab === 'file') {
        document.querySelector('.input-tab:nth-child(1)').classList.add('active');
        document.getElementById('filePanel').classList.add('active');
    } else {
        document.querySelector('.input-tab:nth-child(2)').classList.add('active');
        document.getElementById('pastePanel').classList.add('active');
    }
}

// ドロップゾーンの設定
function setupDropZone() {
    const dropZone = document.getElementById('dropZone');

    dropZone.addEventListener('dragover', function (e) {
        e.preventDefault();
        dropZone.classList.add('dragover');
    });
    dropZone.addEventListener('dragleave', function () {
        dropZone.classList.remove('dragover');
    });
    dropZone.addEventListener('drop', function (e) {
        e.preventDefault();
        dropZone.classList.remove('dragover');
        const files = e.dataTransfer.files;
        if (files.length > 1) {
            // 複数ファイル → PDF一括処理
            processMultiplePDFs(files);
        } else if (files.length === 1) {
            processFile(files[0]);
        }
    });
}

// ファイル選択ハンドラ
function handleFileSelect(input) {
    const files = input.files;
    if (files.length > 1) {
        processMultiplePDFs(files);
    } else if (files.length === 1) {
        processFile(files[0]);
    }
}

// ファイル読み込み処理
function processFile(file) {
    const ext = file.name.split('.').pop().toLowerCase();

    // Excelファイル（.xlsx, .xls）の検出
    if (ext === 'xlsx' || ext === 'xls') {
        showExcelWarning(file.name);
        return;
    }

    // PDFファイル
    if (ext === 'pdf') {
        processPDFFile(file);
        return;
    }

    // CSV/テキストファイル
    const encodingSelect = document.getElementById('encodingSelect');
    const encoding = encodingSelect.value;

    const reader = new FileReader();

    reader.onload = function (e) {
        const text = e.target.result;

        // UTF-8で読んで文字化けしていたらShift_JISでリトライ
        if (encoding === 'UTF-8' && hasGarbledChars(text)) {
            const retryReader = new FileReader();
            retryReader.onload = function (e2) {
                const sjisText = e2.target.result;
                // Shift_JISで読み直して改善されたか確認
                if (!hasGarbledChars(sjisText) || countGarbledChars(sjisText) < countGarbledChars(text)) {
                    encodingSelect.value = 'Shift_JIS';
                    parseAndPreview(sjisText);
                    showFileInfo(file, 'Shift_JIS（自動検出）');
                    showToast('Shift_JISエンコーディングを自動検出しました');
                } else {
                    parseAndPreview(text);
                    showFileInfo(file);
                }
            };
            retryReader.readAsText(file, 'Shift_JIS');
            return;
        }

        parseAndPreview(text);
        showFileInfo(file);
    };

    reader.onerror = function () {
        showToast('ファイルの読み込みに失敗しました');
    };

    reader.readAsText(file, encoding);
}

// PDFファイル読み込み処理（テキストレイヤー優先 → なければOCR）
async function processPDFFile(file) {
    if (typeof pdfjsLib === 'undefined') {
        showToast('PDF.jsライブラリが読み込まれていません。ネット接続を確認してください');
        return;
    }

    showToast('PDF解析中...');
    showFileInfo(file, 'PDF');

    try {
        const arrayBuffer = await file.arrayBuffer();

        // まずテキストレイヤーから抽出を試みる（コピーを渡す：pdf.jsがバッファをdetachするため）
        const pdfRows = await extractTextFromPDF(arrayBuffer.slice(0));
        console.log('[PDF] テキストレイヤー抽出行数:', pdfRows.length);

        if (pdfRows.length > 0) {
            // テキストレイヤーあり → そのまま変換
            const tsvText = pdfRowsToTSV(pdfRows);
            if (tsvText && tsvText.trim().length > 0) {
                document.getElementById('delimiterSelect').value = '\t';
                parseAndPreview(tsvText);
                showFileInfo(file, 'PDF → テキスト変換');
                showToast(`PDFから ${pdfRows.length} 行を抽出しました`);
                return;
            }
        }

        // テキストレイヤーなし → OCRにフォールバック
        console.log('[PDF] テキストレイヤーなし → OCR開始');
        await processPDFWithOCR(arrayBuffer, file);

    } catch (e) {
        hideOcrProgress();
        showToast('PDF読み込みエラー: ' + e.message);
        console.error('[PDF] エラー:', e);
    }
}

// OCRでスキャンPDFを処理（Gemini API優先、Tesseract.jsフォールバック）
async function processPDFWithOCR(arrayBuffer, file) {
    const geminiKey = getGeminiApiKey();

    // Gemini APIキーがあればGeminiで処理
    if (geminiKey) {
        await processPDFWithGemini(arrayBuffer, file);
        return;
    }

    // Tesseract.jsフォールバック
    await processPDFWithTesseract(arrayBuffer, file);
}

// Gemini APIでPDFを処理（汎用書類対応）
async function processPDFWithGemini(arrayBuffer, file) {
    showOcrProgress();

    try {
        const geminiResult = await ocrWithGemini(arrayBuffer, (info) => {
            updateOcrProgress(info.status, info.progress);
        });

        console.log('[Gemini] 解析結果:', geminiResult);

        if (geminiResult.confidence < 0.3 || geminiResult.entries.length === 0) {
            hideOcrProgress();
            showToast('書類を認識できませんでした（confidence: ' + geminiResult.confidence + '）');
            return;
        }

        // Gemini結果からMF行を生成
        const patterns = selectedCompany ? getJournalPatterns(selectedCompany) : [];
        const rules = selectedCompany ? getCorrectionRules(selectedCompany) : [];
        const mfRows = buildMFRowsFromGeminiResult(geminiResult, patterns, rules, selectedIndustry, defaultKashiKamoku);

        // 取引No採番して追加
        mfRows.forEach(row => {
            row.torihikiNo = convertedData.length + 1;
            convertedData.push(row);
        });

        // 勘定科目が未入力の行をGeminiで推測
        if (mfRows.some(r => !r.kariKamoku || !r.kashiKamoku)) {
            updateOcrProgress('勘定科目を推測中...', 0.9);
            await estimateAccountsWithGemini(convertedData, selectedIndustry);
        }

        hideOcrProgress();

        // バリデーション更新
        validationErrors = [];
        convertedData.forEach((row, idx) => {
            const rowErrors = validateMFRow(row, idx);
            validationErrors.push(...rowErrors);
        });

        showReceiptResultStep4();

        const docLabel = geminiResult.documentType ?
            (mfRows[0]?._documentTypeLabel || geminiResult.documentType) : 'PDF';
        showFileInfo(file, `PDF → ${docLabel}`);
        showToast(`${docLabel}から ${mfRows.length} 件の取引を読み取りました`);

    } catch (e) {
        hideOcrProgress();
        console.error('[Gemini] エラー:', e);

        // Gemini失敗時にTesseract.jsにフォールバック
        if (typeof Tesseract !== 'undefined') {
            showToast('Gemini APIエラー。Tesseract.jsで再試行します...');
            await processPDFWithTesseract(arrayBuffer, file);
        } else {
            showToast('OCRエラー: ' + e.message);
        }
    }
}

// Tesseract.jsでPDFを処理（従来のOCR）
async function processPDFWithTesseract(arrayBuffer, file) {
    if (typeof Tesseract === 'undefined') {
        showToast('OCRライブラリが読み込まれていません。Gemini APIキーを設定するか、ネット接続を確認してください');
        return;
    }

    showOcrProgress();

    try {
        const ocrText = await ocrPDFPages(arrayBuffer, (info) => {
            updateOcrProgress(info.status, info.progress);
        });

        hideOcrProgress();

        if (!ocrText || ocrText.trim().length === 0) {
            showToast('OCRでテキストを認識できませんでした');
            return;
        }

        console.log('[OCR] 認識テキスト（先頭200文字）:', ocrText.substring(0, 200));

        // 領収書として解析を試みる
        const receiptData = parseReceiptText(ocrText);

        if (receiptData.confidence >= 0.3) {
            processReceiptOCR(receiptData, file);
        } else {
            document.getElementById('delimiterSelect').value = 'auto';
            parseAndPreview(ocrText);
            showFileInfo(file, 'PDF → OCR変換');
            showToast('OCRでPDFを読み取りました（認識結果を確認してください）');
        }

    } catch (e) {
        hideOcrProgress();
        showToast('OCR処理エラー: ' + e.message);
        console.error('[OCR] エラー:', e);
    }
}

// --- 領収書OCR → 仕訳データ自動生成 ---
function processReceiptOCR(receiptData, file) {
    const patterns = selectedCompany ? getJournalPatterns(selectedCompany) : [];
    const rules = selectedCompany ? getCorrectionRules(selectedCompany) : [];
    const mfRow = buildMFRowFromReceipt(receiptData, patterns, rules, selectedIndustry, defaultKashiKamoku);

    // 取引No採番
    mfRow.torihikiNo = convertedData.length + 1;

    // 既存データに追加（複数PDF対応）
    convertedData.push(mfRow);

    // バリデーション更新
    validationErrors = [];
    convertedData.forEach((row, idx) => {
        const rowErrors = validateMFRow(row, idx);
        validationErrors.push(...rowErrors);
    });

    showReceiptResultStep4();
    showFileInfo(file, 'PDF → 領収書OCR');
    showToast(`領収書を読み取りました: ${receiptData.storeName || '(店名不明)'} ${receiptData.totalAmount ? '¥' + Number(receiptData.totalAmount).toLocaleString() : ''}`);
}

// --- Step4に直接遷移して仕訳行を表示 ---
function showReceiptResultStep4() {
    displayedRows = 0;
    resultFilter = 'all';

    // ステータスバー更新
    const statsBar = document.getElementById('statsBar');
    if (statsBar) {
        const errorCount = validationErrors.length;
        const patternCount = convertedData.filter(r => r._matchedPattern).length;
        const geminiCount = convertedData.filter(r => r._geminiAccount).length;
        const ocrCount = convertedData.filter(r => r._isReceipt).length;

        // 書類種別ごとの件数
        const docTypeCounts = {};
        convertedData.forEach(r => {
            const label = r._documentTypeLabel || (r._isReceipt ? '領収書' : '');
            if (label) docTypeCounts[label] = (docTypeCounts[label] || 0) + 1;
        });
        const docTypeSummary = Object.entries(docTypeCounts).map(([k, v]) => `${k} ${v}件`).join('、');

        statsBar.innerHTML = `
            <span>📊 ${convertedData.length} 行</span>
            ${docTypeSummary ? `<span>📄 ${docTypeSummary}</span>` : ''}
            ${patternCount > 0 ? `<span>🔄 パターン適用 ${patternCount} 件</span>` : ''}
            ${geminiCount > 0 ? `<span style="color: #d2a8ff;">🤖 Gemini科目 ${geminiCount} 件</span>` : ''}
            ${errorCount > 0 ? `<span style="color: var(--error);">⚠ ${errorCount} 件のエラー</span>` : '<span style="color: var(--success);">✓ エラーなし</span>'}
        `;
        statsBar.style.display = 'flex';
    }

    goToStep(4);
    renderConvertedPreview();

    // 仕訳パターンカードも表示
    if (typeof renderJournalPatternsCard === 'function') {
        renderJournalPatternsCard();
    }
}

// --- 複数PDFファイルの一括処理 ---
async function processMultiplePDFs(files) {
    // 最初の処理で既存データをクリア
    convertedData = [];
    validationErrors = [];

    const pdfFiles = Array.from(files).filter(f => f.name.toLowerCase().endsWith('.pdf'));
    const nonPdfFiles = Array.from(files).filter(f => !f.name.toLowerCase().endsWith('.pdf'));

    if (pdfFiles.length === 0) {
        if (files.length > 0) processFile(files[0]);
        return;
    }

    showOcrProgress();
    let processedCount = 0;

    const geminiKey = getGeminiApiKey();
    const patterns = selectedCompany ? getJournalPatterns(selectedCompany) : [];
    const rules = selectedCompany ? getCorrectionRules(selectedCompany) : [];

    if (geminiKey) {
        // --- Geminiバッチ処理（1回のAPI呼び出しで全PDF処理） ---
        try {
            // 1. 全PDFを画像に変換
            const imageDataArray = [];
            for (let i = 0; i < pdfFiles.length; i++) {
                updateOcrProgress(`画像変換中: ${i + 1} / ${pdfFiles.length}`, i / pdfFiles.length * 0.2);
                const ab = await pdfFiles[i].arrayBuffer();
                const base64Images = await pdfToBase64Images(ab);
                imageDataArray.push({ fileName: pdfFiles[i].name, base64Images });
            }

            // 2. 一括でGemini APIに送信
            const batchResults = await ocrBatchWithGemini(imageDataArray, (info) => {
                updateOcrProgress(info.status, 0.2 + info.progress * 0.6);
            });

            // 3. 結果からMF行を生成
            updateOcrProgress('仕訳データを生成中...', 0.85);
            for (const result of batchResults) {
                const fileIdx = result.fileIndex - 1;
                const fileName = (fileIdx >= 0 && fileIdx < pdfFiles.length) ? pdfFiles[fileIdx].name : '';

                if (result.confidence < 0.3 || result.entries.length === 0) continue;

                const mfRows = buildMFRowsFromGeminiResult(result, patterns, rules, selectedIndustry, defaultKashiKamoku);
                mfRows.forEach(row => {
                    row.torihikiNo = convertedData.length + 1;
                    row.memo = fileName;
                    convertedData.push(row);
                    processedCount++;
                });
            }

            // 4. 科目未入力行をGemini推測
            if (convertedData.some(r => !r.kariKamoku)) {
                updateOcrProgress('勘定科目を推測中...', 0.92);
                await estimateAccountsWithGemini(convertedData, selectedIndustry);
            }

        } catch (e) {
            console.error('[Gemini Batch] エラー:', e);
            showToast('Geminiバッチ処理エラー: ' + e.message + '（個別処理にフォールバック）');

            // バッチ失敗時は個別Tesseractフォールバック
            for (let i = 0; i < pdfFiles.length; i++) {
                try {
                    updateOcrProgress(`Tesseract: ${i + 1} / ${pdfFiles.length}`, i / pdfFiles.length);
                    const ab = await pdfFiles[i].arrayBuffer();
                    const pdfRows = await extractTextFromPDF(ab.slice(0));
                    let ocrText = '';
                    if (pdfRows.length > 0) {
                        ocrText = pdfRows.map(row => row.map(item => item.text).join(' ')).join('\n');
                    } else if (typeof Tesseract !== 'undefined') {
                        ocrText = await ocrPDFPages(ab, (info) => {
                            updateOcrProgress(`${i + 1}/${pdfFiles.length}: ${info.status}`, (i + info.progress) / pdfFiles.length);
                        });
                    }
                    if (ocrText && ocrText.trim()) {
                        const receiptData = parseReceiptText(ocrText);
                        if (receiptData.confidence >= 0.3) {
                            const mfRow = buildMFRowFromReceipt(receiptData, patterns, rules, selectedIndustry, defaultKashiKamoku);
                            mfRow.torihikiNo = convertedData.length + 1;
                            mfRow.memo = pdfFiles[i].name;
                            convertedData.push(mfRow);
                            processedCount++;
                        }
                    }
                } catch (err) {
                    console.error(`[PDF] ${pdfFiles[i].name}:`, err);
                }
            }
        }
    } else {
        // --- Tesseract.jsフロー（APIキーなし） ---
        for (let i = 0; i < pdfFiles.length; i++) {
            updateOcrProgress(`${i + 1} / ${pdfFiles.length} 件目: ${pdfFiles[i].name}`, i / pdfFiles.length);
            try {
                const ab = await pdfFiles[i].arrayBuffer();
                const pdfRows = await extractTextFromPDF(ab.slice(0));
                let ocrText = '';
                if (pdfRows.length > 0) {
                    ocrText = pdfRows.map(row => row.map(item => item.text).join(' ')).join('\n');
                } else if (typeof Tesseract !== 'undefined') {
                    ocrText = await ocrPDFPages(ab, (info) => {
                        updateOcrProgress(`${i + 1}/${pdfFiles.length}: ${info.status}`, (i + info.progress) / pdfFiles.length);
                    });
                }
                if (ocrText && ocrText.trim()) {
                    const receiptData = parseReceiptText(ocrText);
                    if (receiptData.confidence >= 0.3) {
                        const mfRow = buildMFRowFromReceipt(receiptData, patterns, rules, selectedIndustry, defaultKashiKamoku);
                        mfRow.torihikiNo = convertedData.length + 1;
                        mfRow.memo = pdfFiles[i].name;
                        convertedData.push(mfRow);
                        processedCount++;
                    }
                }
            } catch (e) {
                console.error(`[PDF] ${pdfFiles[i].name}:`, e);
            }
        }
    }

    hideOcrProgress();

    if (processedCount === 0) {
        showToast('書類を認識できませんでした');
        return;
    }

    // バリデーション
    validationErrors = [];
    convertedData.forEach((row, idx) => {
        const rowErrors = validateMFRow(row, idx);
        validationErrors.push(...rowErrors);
    });

    showReceiptResultStep4();
    showToast(`${processedCount} 件の取引を読み取りました`);
}

// OCRプログレスモーダルの表示/非表示/更新
function showOcrProgress() {
    document.getElementById('ocrProgressModal').classList.add('show');
    const msg = getGeminiApiKey() ? 'Gemini APIで解析中...' : '日本語OCRデータを読み込み中...';
    updateOcrProgress(msg, 0);
}

function hideOcrProgress() {
    document.getElementById('ocrProgressModal').classList.remove('show');
}

function updateOcrProgress(status, progress) {
    const pct = Math.round((progress || 0) * 100);
    document.getElementById('ocrProgressStatus').textContent = status || '処理中...';
    document.getElementById('ocrProgressBar').style.width = pct + '%';
    document.getElementById('ocrProgressPercent').textContent = pct + '%';
}

// 文字化け検出（UTF-8で読んだ時にShift_JISバイト列が化けるパターン）
function hasGarbledChars(text) {
    return countGarbledChars(text) > text.length * 0.02;
}

function countGarbledChars(text) {
    // 置換文字（U+FFFD）、またはShift_JIS→UTF-8文字化けの典型パターン
    let count = 0;
    for (let i = 0; i < text.length; i++) {
        const code = text.charCodeAt(i);
        // U+FFFD（置換文字）
        if (code === 0xFFFD) { count++; continue; }
        // 制御文字（タブ・改行・CR以外の0x00-0x1F）
        if (code < 0x20 && code !== 0x09 && code !== 0x0A && code !== 0x0D) { count++; continue; }
    }
    return count;
}

// ファイル情報を表示
function showFileInfo(file, encodingNote) {
    const fileInfo = document.getElementById('fileInfo');
    const sizeKB = (file.size / 1024).toFixed(1);
    const encInfo = encodingNote ? `<span class="file-info-encoding">${escapeHtml(encodingNote)}</span>` : '';
    fileInfo.innerHTML = `<div class="file-info">
        <span>📄</span>
        <span class="file-info-name">${escapeHtml(file.name)}</span>
        <span class="file-info-size">(${sizeKB} KB)</span>
        ${encInfo}
    </div>`;
    fileInfo.style.display = 'block';
}

// Excelファイルの警告表示
function showExcelWarning(fileName) {
    const fileInfo = document.getElementById('fileInfo');
    fileInfo.innerHTML = `<div class="excel-warning">
        <div class="excel-warning-title">⚠ Excelファイルは直接読み込めません</div>
        <div class="excel-warning-file">選択されたファイル: ${escapeHtml(fileName)}</div>
        <div class="excel-warning-guide">
            <strong>以下のいずれかの方法をお試しください:</strong>
            <ol>
                <li><strong>CSVとして保存</strong> → Excelで「名前を付けて保存」→ 形式を「CSV UTF-8」または「CSV」に変更して保存 → そのCSVをアップロード</li>
                <li><strong>コピー＆ペースト</strong> → Excelでデータ範囲を選択してコピー（Ctrl+C）→ 上の「テキスト貼り付け」タブに貼り付け</li>
            </ol>
        </div>
    </div>`;
    fileInfo.style.display = 'block';
    showToast('Excelファイルは直接読み込めません。CSVに変換するかコピペしてください');
}

// ペーストエリアの自動検出設定
function setupPasteAutoDetect() {
    const pasteArea = document.getElementById('pasteArea');

    // Ctrl+V でペーストした時に自動読み込み
    pasteArea.addEventListener('paste', function () {
        // ペースト完了後にデータを処理（少し遅延が必要）
        setTimeout(() => {
            const text = pasteArea.value.trim();
            if (text && text.split('\n').length >= 2) {
                // 2行以上あれば自動的に読み込む
                parseAndPreview(text);
                showToast('貼り付けデータを自動検出しました');
            }
        }, 100);
    });
}

// テキスト貼り付け処理
function handlePasteData() {
    const text = document.getElementById('pasteArea').value.trim();
    if (!text) {
        showToast('データを入力してください');
        return;
    }
    parseAndPreview(text);
}

// パース + プレビュー表示
function parseAndPreview(text) {
    const delimiterSelect = document.getElementById('delimiterSelect').value;
    const delimiter = delimiterSelect === 'auto' ? detectDelimiter(text) : delimiterSelect;
    const hasHeader = document.getElementById('hasHeader').checked;

    const parsed = parseCSV(text, delimiter, hasHeader);
    rawHeaders = parsed.headers;
    rawData = parsed.rows;

    if (rawData.length === 0) {
        showToast('データが見つかりませんでした');
        return;
    }

    // ヘッダー行自動検出で先頭行をスキップした場合に通知
    if (hasHeader && parsed.headerRowIdx > 0) {
        showToast(`${parsed.headerRowIdx} 行のタイトル行をスキップし、${parsed.headerRowIdx + 1} 行目をヘッダーとして検出しました`);
    }

    renderDataPreview();
    document.getElementById('step1Next').disabled = false;

    // 自動変換を試みる
    tryAutoConvert();
}

// データ読み込み後に自動マッピング→変換→ステップ4まで一気に進む
function tryAutoConvert() {
    let detectionMethod = '';

    // 1. ヘッダー名ベースの自動マッピング
    columnMapping = autoDetectMapping(rawHeaders);

    // 2. プリセットで補強（ヘッダーからデータ種別を推定）
    const bestPreset = detectBestPreset();
    if (bestPreset) {
        applyPresetSilent(bestPreset);
        detectionMethod = `「${bestPreset.label}」として`;
    }

    // 3. 必須フィールドが揃っているかチェック
    let mappedFields = Object.values(columnMapping);
    let hasDate = mappedFields.includes('torihikiDate');
    let hasAmount = mappedFields.includes('kariKingaku') || mappedFields.includes('kashiKingaku');

    // 4. ヘッダー名で足りなかったら、データ内容ベースで補完
    if (!hasDate || !hasAmount) {
        const contentMapping = autoDetectMappingByContent(rawHeaders, rawData);

        // 既存マッピングに足りない部分だけ補完
        for (const [idx, field] of Object.entries(contentMapping)) {
            if (!columnMapping[idx] && !Object.values(columnMapping).includes(field)) {
                columnMapping[idx] = field;
            }
        }

        mappedFields = Object.values(columnMapping);
        hasDate = mappedFields.includes('torihikiDate');
        hasAmount = mappedFields.includes('kariKingaku') || mappedFields.includes('kashiKingaku');

        if (!detectionMethod && (hasDate || hasAmount)) {
            detectionMethod = 'データ内容から自動で';
        }
    }

    // 5. それでも足りなかったらデータ内容ベースで全面的に判定
    if (!hasDate || !hasAmount) {
        columnMapping = autoDetectMappingByContent(rawHeaders, rawData);

        mappedFields = Object.values(columnMapping);
        hasDate = mappedFields.includes('torihikiDate');
        hasAmount = mappedFields.includes('kariKingaku') || mappedFields.includes('kashiKingaku');
        detectionMethod = 'データ内容から自動で';
    }

    // 6. 最終チェック：最低限のフィールドがない場合はステップ2に案内
    if (!hasDate && !hasAmount) {
        showToast(`${rawData.length} 行を読み込みました。カラムを割り当ててください`);
        return;
    }

    // 7. 日付形式を自動検出
    const dateColIdx = Object.entries(columnMapping).find(([_, v]) => v === 'torihikiDate');
    if (dateColIdx) {
        const samples = rawData.slice(0, 10).map(row => row[parseInt(dateColIdx[0])] || '');
        const detected = detectDateFormat(samples);
        if (detected !== 'auto') {
            conversionRules.dateFormat = detected;
        }
    }

    // 8. 変換実行 → ステップ4へ直行
    collectRulesFromUI();
    runConversion();

    currentStep = 4;
    document.querySelectorAll('.converter-step').forEach(el => el.classList.remove('active'));
    document.getElementById('step4').classList.add('active');
    updateStepIndicator();
    window.scrollTo(0, 0);

    if (!detectionMethod) detectionMethod = '自動で';
    const errorInfo = validationErrors.length > 0 ? `（エラー ${validationErrors.length} 件）` : '';
    showToast(`${rawData.length} 行を${detectionMethod}変換しました ${errorInfo}`);
}

// ヘッダーからベストなプリセットを推定
function detectBestPreset() {
    const normalizedHeaders = rawHeaders.map(h => h.trim().toLowerCase());

    let bestPreset = null;
    let bestScore = 0;

    MAPPING_PRESETS.forEach(preset => {
        let score = 0;
        preset.mapping.forEach(rule => {
            const matched = normalizedHeaders.some(h =>
                rule.match.some(k => h.includes(k.toLowerCase()))
            );
            if (matched) score++;
        });

        // マッチ率が50%以上で最高スコアなら採用
        if (score >= preset.mapping.length * 0.5 && score > bestScore) {
            bestScore = score;
            bestPreset = preset;
        }
    });

    return bestPreset;
}

// プリセットをトースト無しで適用（自動変換用）
function applyPresetSilent(preset) {
    // マッピングをリセットしてプリセットで割当
    columnMapping = {};

    rawHeaders.forEach((header, idx) => {
        const normalized = header.trim().toLowerCase();
        for (const rule of preset.mapping) {
            if (!rule.field) continue;
            if (rule.match.some(k => normalized.includes(k.toLowerCase()))) {
                if (!Object.values(columnMapping).includes(rule.field)) {
                    columnMapping[idx] = rule.field;
                    break;
                }
            }
        }
    });

    // 自動推定でフォールバック
    const autoMapping = autoDetectMapping(rawHeaders);
    for (const [idx, field] of Object.entries(autoMapping)) {
        if (!columnMapping[idx] && !Object.values(columnMapping).includes(field)) {
            columnMapping[idx] = field;
        }
    }

    // 固定値を適用
    if (preset.fixedValues) {
        conversionRules.fixedValues = { ...conversionRules.fixedValues, ...preset.fixedValues };
    }
}

// データプレビュー描画
function renderDataPreview() {
    const container = document.getElementById('dataPreview');
    container.style.display = 'block';

    const count = document.getElementById('previewCount');
    count.textContent = `${rawData.length} 行`;

    const PREVIEW_LIMIT = 20;
    const previewRows = rawData.slice(0, PREVIEW_LIMIT);
    let html = '<table class="table"><thead><tr>';
    rawHeaders.forEach(h => {
        html += `<th>${escapeHtml(h)}</th>`;
    });
    html += '</tr></thead><tbody>';

    previewRows.forEach(row => {
        html += '<tr>';
        rawHeaders.forEach((_, idx) => {
            html += `<td>${escapeHtml(row[idx] || '')}</td>`;
        });
        html += '</tr>';
    });
    html += '</tbody></table>';

    document.getElementById('previewTable').innerHTML = html;

    const info = document.getElementById('previewInfo');
    if (rawData.length > PREVIEW_LIMIT) {
        info.textContent = `最初の${PREVIEW_LIMIT}行を表示中（全${rawData.length}行）`;
    } else {
        info.textContent = `全${rawData.length}行を表示`;
    }
}

// ===== ステップ2: カラムマッピング =====

function renderMappingUI() {
    // 前回のマッピングがなければ自動推定
    if (Object.keys(columnMapping).length === 0) {
        columnMapping = autoDetectMapping(rawHeaders);
    }

    renderPresetButtons();
    renderMappingTable();
    renderMappingPreview();
    updateMappingStatus();
}

// プリセットボタン描画
function renderPresetButtons() {
    const container = document.getElementById('presetButtons');
    let html = '';
    MAPPING_PRESETS.forEach(preset => {
        html += `<button class="preset-btn" onclick="applyPreset('${preset.id}')" title="${escapeHtml(preset.desc)}">
            ${escapeHtml(preset.label)}
            <span class="preset-desc">${escapeHtml(preset.desc)}</span>
        </button>`;
    });
    container.innerHTML = html;
}

// プリセット適用
function applyPreset(presetId) {
    const preset = MAPPING_PRESETS.find(p => p.id === presetId);
    if (!preset) return;

    // マッピングをリセット
    columnMapping = {};

    // ヘッダー名でマッチングして自動割当
    rawHeaders.forEach((header, idx) => {
        const normalized = header.trim().toLowerCase();
        for (const rule of preset.mapping) {
            if (!rule.field) continue; // null = 無視するカラム
            if (rule.match.some(k => normalized.includes(k.toLowerCase()))) {
                // 重複チェック
                if (!Object.values(columnMapping).includes(rule.field)) {
                    columnMapping[idx] = rule.field;
                    break;
                }
            }
        }
    });

    // マッチしなかったカラムは通常の自動推定にフォールバック
    const autoMapping = autoDetectMapping(rawHeaders);
    for (const [idx, field] of Object.entries(autoMapping)) {
        if (!columnMapping[idx] && !Object.values(columnMapping).includes(field)) {
            columnMapping[idx] = field;
        }
    }

    // 固定値をステップ3に反映するために保存
    if (preset.fixedValues) {
        conversionRules.fixedValues = { ...conversionRules.fixedValues, ...preset.fixedValues };
    }

    renderMappingTable();
    renderMappingPreview();
    updateMappingStatus();

    const fixedInfo = preset.fixedValues
        ? '（固定値: ' + Object.entries(preset.fixedValues).map(([k, v]) => {
            const label = MF_COLUMNS.find(c => c.key === k)?.label || k;
            return `${label}=${v}`;
        }).join(', ') + '）'
        : '';
    showToast(`「${preset.label}」プリセットを適用しました ${fixedInfo}`);
}

// マッピングテーブル描画
function renderMappingTable() {
    const container = document.getElementById('mappingUI');

    let html = '<table class="mapping-table"><thead><tr>' +
        '<th class="mt-col-source">入力カラム</th>' +
        '<th class="mt-col-samples">サンプル値</th>' +
        '<th class="mt-col-arrow"></th>' +
        '<th class="mt-col-target">MFフィールド</th>' +
        '</tr></thead><tbody>';

    rawHeaders.forEach((header, idx) => {
        const samples = rawData.slice(0, 3).map(row => row[idx] || '').filter(v => v);
        const sampleText = samples.length > 0 ? samples.join(' / ') : '(空)';
        const currentValue = columnMapping[idx] || '';
        const isMapped = currentValue !== '';

        html += `<tr class="mapping-table-row${isMapped ? ' mapped' : ''}">
            <td class="mt-source">${escapeHtml(header)}</td>
            <td class="mt-samples" title="${escapeHtml(sampleText)}">${escapeHtml(sampleText)}</td>
            <td class="mt-arrow">${isMapped ? '→' : ''}</td>
            <td class="mt-target">
                <select onchange="updateMapping(${idx}, this.value)" id="mapping_${idx}" class="mapping-select${isMapped ? ' selected' : ''}">
                    <option value="">-- 未割当 --</option>`;

        MAPPABLE_COLUMNS.forEach(col => {
            const selected = currentValue === col.key ? ' selected' : '';
            // 他のカラムで使用中のフィールドにマークを付ける
            const usedBy = Object.entries(columnMapping).find(([k, v]) => v === col.key && parseInt(k) !== idx);
            const usedLabel = usedBy ? ' (使用中)' : '';
            html += `<option value="${col.key}"${selected}${usedBy ? ' class="option-used"' : ''}>${col.label}${usedLabel}</option>`;
        });

        html += `</select></td></tr>`;
    });

    html += '</tbody></table>';
    container.innerHTML = html;
}

// マッピング状況サマリー
function updateMappingStatus() {
    const status = document.getElementById('mappingStatus');
    const mapped = Object.keys(columnMapping).length;
    const total = rawHeaders.length;
    const required = ['torihikiDate', 'kariKingaku'];
    const missingRequired = required.filter(f => !Object.values(columnMapping).includes(f));

    let html = `<span class="mapping-count">${mapped}/${total} カラム割当済</span>`;
    if (missingRequired.length > 0) {
        const labels = missingRequired.map(f => MF_COLUMNS.find(c => c.key === f)?.label || f);
        html += `<span class="mapping-warn">必須: ${labels.join(', ')} が未割当</span>`;
    } else {
        html += `<span class="mapping-ok">必須フィールド割当済</span>`;
    }
    status.innerHTML = html;
}

function autoDetectAndApply() {
    columnMapping = autoDetectMapping(rawHeaders);
    renderMappingTable();
    renderMappingPreview();
    updateMappingStatus();
    showToast('カラムを自動推定しました');
}

function updateMapping(sourceIdx, targetField) {
    if (targetField) {
        // 重複チェック：同じMFフィールドが他のカラムに割り当てられていたら解除
        for (const [key, val] of Object.entries(columnMapping)) {
            if (val === targetField && parseInt(key) !== sourceIdx) {
                delete columnMapping[key];
                const otherSelect = document.getElementById('mapping_' + key);
                if (otherSelect) {
                    otherSelect.value = '';
                    otherSelect.classList.remove('selected');
                }
            }
        }
        columnMapping[sourceIdx] = targetField;
    } else {
        delete columnMapping[sourceIdx];
    }

    // 行の見た目を更新（テーブル全体を再描画せずに済ませる）
    const row = document.getElementById('mapping_' + sourceIdx)?.closest('tr');
    if (row) {
        const arrow = row.querySelector('.mt-arrow');
        const select = document.getElementById('mapping_' + sourceIdx);
        if (targetField) {
            row.classList.add('mapped');
            arrow.textContent = '→';
            select.classList.add('selected');
        } else {
            row.classList.remove('mapped');
            arrow.textContent = '';
            select.classList.remove('selected');
        }
    }

    renderMappingPreview();
    updateMappingStatus();
}

function collectMappingFromUI() {
    columnMapping = {};
    rawHeaders.forEach((_, idx) => {
        const select = document.getElementById('mapping_' + idx);
        if (select && select.value) {
            columnMapping[idx] = select.value;
        }
    });
}

function clearAllMapping() {
    columnMapping = {};
    renderMappingTable();
    renderMappingPreview();
    updateMappingStatus();
    showToast('マッピングをクリアしました');
}

function renderMappingPreview() {
    const container = document.getElementById('mappingPreview');
    const tableContainer = document.getElementById('mappingPreviewTable');

    const mappedFields = Object.values(columnMapping);
    if (mappedFields.length === 0) {
        container.style.display = 'none';
        return;
    }
    container.style.display = 'block';

    // MFカラムのうちマッピングされたものだけ表示
    const activeCols = MAPPABLE_COLUMNS.filter(c => mappedFields.includes(c.key));
    const previewRows = rawData.slice(0, 3);

    let html = '<table class="table"><thead><tr>';
    activeCols.forEach(col => {
        html += `<th>${col.label}</th>`;
    });
    html += '</tr></thead><tbody>';

    previewRows.forEach(row => {
        html += '<tr>';
        activeCols.forEach(col => {
            const sourceIdx = Object.entries(columnMapping).find(([_, v]) => v === col.key);
            const value = sourceIdx ? (row[parseInt(sourceIdx[0])] || '') : '';
            html += `<td>${escapeHtml(value)}</td>`;
        });
        html += '</tr>';
    });
    html += '</tbody></table>';

    tableContainer.innerHTML = html;
}

// ===== ステップ3: 変換ルール =====

function renderDateFormatOptions() {
    const select = document.getElementById('dateFormatSelect');
    select.innerHTML = '';
    DATE_FORMATS.forEach(fmt => {
        const option = document.createElement('option');
        option.value = fmt.id;
        option.textContent = fmt.label;
        select.appendChild(option);
    });
}

function renderRulesUI() {
    // 日付形式の自動検出
    const dateColIdx = Object.entries(columnMapping).find(([_, v]) => v === 'torihikiDate');
    if (dateColIdx) {
        const samples = rawData.slice(0, 10).map(row => row[parseInt(dateColIdx[0])] || '');
        const detected = detectDateFormat(samples);
        const info = document.getElementById('dateDetectInfo');
        if (detected !== 'auto') {
            info.textContent = `検出された形式: ${DATE_FORMATS.find(f => f.id === detected)?.label || detected}`;
            document.getElementById('dateFormatSelect').value = detected;
            conversionRules.dateFormat = detected;
        } else {
            info.textContent = '日付の形式を自動検出できませんでした。手動で選択してください。';
        }
    }

    // 勘定科目マッピング
    renderAccountMappingUI();

    // 固定値UI
    renderFixedValueUI();
}

function renderAccountMappingUI() {
    const container = document.getElementById('accountMappingList');

    // 入力データから使われている科目名を抽出
    const kamokuIdx = Object.entries(columnMapping).find(([_, v]) => v === 'kariKamoku');
    const uniqueNames = new Set();
    if (kamokuIdx) {
        rawData.forEach(row => {
            const val = (row[parseInt(kamokuIdx[0])] || '').trim();
            if (val) uniqueNames.add(val);
        });
    }

    let html = '';

    // プリセットで変換対象になるものを表示
    uniqueNames.forEach(name => {
        const alias = ACCOUNT_ALIASES[name];
        if (alias) {
            html += `<div class="account-mapping-row" data-preset="true">
                <input type="text" value="${escapeHtml(name)}" readonly style="background: var(--bg-secondary);">
                <span class="mapping-arrow">→</span>
                <input type="text" value="${escapeHtml(alias)}" readonly style="background: var(--bg-secondary);">
                <span style="font-size: 11px; color: var(--text-muted);">プリセット</span>
            </div>`;
        }
    });

    // カスタムマッピング行
    const customMappings = conversionRules.accountMapping || {};
    for (const [from, to] of Object.entries(customMappings)) {
        html += createAccountMappingRowHtml(from, to);
    }

    container.innerHTML = html;
}

function createAccountMappingRowHtml(from, to) {
    return `<div class="account-mapping-row" data-custom="true">
        <input type="text" class="acmap-from" value="${escapeHtml(from || '')}" placeholder="変換前の科目名">
        <span class="mapping-arrow">→</span>
        <input type="text" class="acmap-to" value="${escapeHtml(to || '')}" placeholder="MFの勘定科目名" list="accountSuggestions">
        <button class="btn-icon" onclick="this.parentElement.remove()" title="削除">✕</button>
    </div>`;
}

function addAccountMappingRow() {
    const container = document.getElementById('accountMappingList');
    const div = document.createElement('div');
    div.innerHTML = createAccountMappingRowHtml('', '');
    container.appendChild(div.firstElementChild);
}

function renderFixedValueUI() {
    const container = document.getElementById('fixedValueList');
    // 固定値設定可能なフィールド
    const fixableFields = MAPPABLE_COLUMNS.filter(c =>
        !['torihikiDate'].includes(c.key)
    );

    let html = '<datalist id="accountSuggestions">';
    COMMON_ACCOUNTS.forEach(a => {
        html += `<option value="${escapeHtml(a)}">`;
    });
    html += '</datalist>';

    fixableFields.forEach(field => {
        const currentValue = (conversionRules.fixedValues && conversionRules.fixedValues[field.key]) || '';
        const isAccountField = field.key.includes('Kamoku');
        const isTaxField = field.key.includes('Zeiku');

        html += `<div class="fixed-value-row">
            <label>${field.label}</label>`;

        if (isTaxField) {
            html += `<select class="fixed-value-input" data-field="${field.key}">
                <option value="">（設定なし）</option>`;
            TAX_CATEGORIES.forEach(cat => {
                const selected = currentValue === cat ? ' selected' : '';
                html += `<option value="${escapeHtml(cat)}"${selected}>${escapeHtml(cat)}</option>`;
            });
            html += '</select>';
        } else if (isAccountField) {
            html += `<input type="text" class="fixed-value-input" data-field="${field.key}"
                value="${escapeHtml(currentValue)}" placeholder="（設定なし）" list="accountSuggestions">`;
        } else {
            html += `<input type="text" class="fixed-value-input" data-field="${field.key}"
                value="${escapeHtml(currentValue)}" placeholder="（設定なし）">`;
        }

        html += '</div>';
    });

    container.innerHTML = html;
}

function collectRulesFromUI() {
    // 日付形式
    conversionRules.dateFormat = document.getElementById('dateFormatSelect').value;

    // カスタム勘定科目マッピング
    conversionRules.accountMapping = {};
    document.querySelectorAll('.account-mapping-row[data-custom="true"]').forEach(row => {
        const from = row.querySelector('.acmap-from')?.value.trim();
        const to = row.querySelector('.acmap-to')?.value.trim();
        if (from && to) {
            conversionRules.accountMapping[from] = to;
        }
    });

    // 固定値
    conversionRules.fixedValues = {};
    document.querySelectorAll('.fixed-value-input').forEach(input => {
        const field = input.dataset.field;
        const value = input.value.trim();
        if (value) {
            conversionRules.fixedValues[field] = value;
        }
    });
}

// ===== ステップ4: プレビュー & エクスポート =====

// 編集中のセル情報
let pendingCorrection = null; // { rowIdx, field, oldValue, newValue }

let lastSkippedRows = 0;

function runConversion() {
    // 選択中の会社の訂正ルールと仕訳パターンと業種を変換ルールに組み込む
    conversionRules.correctionRules = getCorrectionRules(selectedCompany);
    conversionRules.journalPatterns = getJournalPatterns(selectedCompany);
    conversionRules.industry = selectedIndustry;
    conversionRules.defaultKashiKamoku = defaultKashiKamoku;

    const result = convertToMFFormat(rawData, rawHeaders, columnMapping, conversionRules);
    convertedData = result.rows;
    validationErrors = result.errors;
    lastSkippedRows = result.skippedRows || 0;

    renderConversionResult();
    displayedRows = 0;
    renderConvertedPreview();
    renderCorrectionRulesCard();
    renderJournalPatternsCard();
}

function renderConversionResult() {
    const container = document.getElementById('conversionResult');
    const errorRows = new Set(validationErrors.map(e => e.row));

    // 合計金額を計算
    const stats = calcConversionStats(convertedData);

    let html = '';

    if (validationErrors.length === 0) {
        html += `<div class="success-summary">
            <div class="success-summary-title">✓ ${convertedData.length} 件の仕訳を正常に変換しました</div>
        </div>`;
    } else {
        html += '<div class="error-summary">';
        html += `<div class="error-summary-title">⚠ ${errorRows.size} 行にエラーがあります（全${convertedData.length}行中）</div>`;
        html += '<div class="error-summary-list">';
        validationErrors.slice(0, 20).forEach(err => {
            const fieldLabel = MF_COLUMNS.find(c => c.key === err.field)?.label || err.field;
            html += `<div>行 ${err.row + 1}: [${escapeHtml(fieldLabel)}] ${escapeHtml(err.message)}</div>`;
        });
        if (validationErrors.length > 20) {
            html += `<div>... 他 ${validationErrors.length - 20} 件のエラー</div>`;
        }
        html += '</div></div>';
    }

    // サマリーカード
    html += '<div class="stats-bar">';
    html += `<div class="stat-item"><span class="stat-label">借方合計</span><span class="stat-value">${formatAmount(stats.totalDebit)}</span></div>`;
    html += `<div class="stat-item"><span class="stat-label">貸方合計</span><span class="stat-value">${formatAmount(stats.totalCredit)}</span></div>`;
    if (stats.totalDebit === stats.totalCredit) {
        html += `<div class="stat-item stat-ok"><span class="stat-label">貸借</span><span class="stat-value">一致</span></div>`;
    } else {
        const diff = Math.abs(stats.totalDebit - stats.totalCredit);
        html += `<div class="stat-item stat-warn"><span class="stat-label">差額</span><span class="stat-value">${formatAmount(diff)}</span></div>`;
    }
    if (stats.dateRange.from && stats.dateRange.to) {
        html += `<div class="stat-item"><span class="stat-label">期間</span><span class="stat-value">${stats.dateRange.from} 〜 ${stats.dateRange.to}</span></div>`;
    }
    if (lastSkippedRows > 0) {
        html += `<div class="stat-item"><span class="stat-label">スキップ</span><span class="stat-value">${lastSkippedRows} 行</span></div>`;
    }
    html += '</div>';

    container.innerHTML = html;
    document.getElementById('resultCount').textContent = `${convertedData.length} 行`;
}

// 変換結果の統計情報を計算
function calcConversionStats(rows) {
    let totalDebit = 0;
    let totalCredit = 0;
    let minDate = null;
    let maxDate = null;

    rows.forEach(row => {
        const debit = parseInt(row.kariKingaku) || 0;
        const credit = parseInt(row.kashiKingaku) || 0;
        totalDebit += debit;
        totalCredit += credit;

        if (row.torihikiDate) {
            if (!minDate || row.torihikiDate < minDate) minDate = row.torihikiDate;
            if (!maxDate || row.torihikiDate > maxDate) maxDate = row.torihikiDate;
        }
    });

    return {
        totalDebit,
        totalCredit,
        dateRange: { from: minDate, to: maxDate },
    };
}

// 金額フォーマット（3桁カンマ区切り）
function formatAmount(num) {
    return num.toLocaleString('ja-JP') + ' 円';
}

// 表示中のアクティブカラム（編集時に参照）
let activeColsCache = [];
let resultFilter = 'all'; // 'all' | 'errors'

function setResultFilter(filter) {
    resultFilter = filter;
    displayedRows = 0;

    // ボタンの見た目を更新
    document.querySelectorAll('.filter-btn').forEach(btn => btn.classList.remove('active'));
    if (filter === 'all') {
        document.querySelector('.filter-btn:nth-child(1)').classList.add('active');
    } else {
        document.getElementById('filterErrorBtn').classList.add('active');
    }

    renderConvertedPreview();
}

function renderConvertedPreview() {
    const container = document.getElementById('resultTable');
    const showMore = document.getElementById('showMoreBtn');
    const filterBar = document.getElementById('filterBar');

    const errorRowSet = new Set(validationErrors.map(e => e.row));

    // パターン適用行を収集
    const patternRowSet = new Set();
    convertedData.forEach((row, i) => {
        if (row._matchedPattern) patternRowSet.add(i);
    });

    // フィルターバー表示（エラーまたはパターンがある場合）
    if (validationErrors.length > 0 || patternRowSet.size > 0) {
        filterBar.style.display = 'flex';
        document.getElementById('filterErrorBtn').textContent = `エラーのみ (${errorRowSet.size})`;
        // パターンフィルターボタンを動的に追加/更新
        let patternBtn = document.getElementById('filterPatternBtn');
        if (!patternBtn && patternRowSet.size > 0) {
            patternBtn = document.createElement('button');
            patternBtn.className = 'filter-btn';
            patternBtn.id = 'filterPatternBtn';
            patternBtn.onclick = function () { setResultFilter('patterns'); };
            filterBar.appendChild(patternBtn);
        }
        if (patternBtn) {
            patternBtn.textContent = `パターン適用 (${patternRowSet.size})`;
            patternBtn.style.display = patternRowSet.size > 0 ? '' : 'none';
            patternBtn.classList.toggle('active', resultFilter === 'patterns');
        }
    } else {
        filterBar.style.display = 'none';
    }

    // フィルタリング：表示する行のインデックス配列を作る
    let visibleIndices = [];
    for (let i = 0; i < convertedData.length; i++) {
        if (resultFilter === 'errors' && !errorRowSet.has(i)) continue;
        if (resultFilter === 'patterns' && !patternRowSet.has(i)) continue;
        visibleIndices.push(i);
    }

    const endIdx = Math.min(displayedRows + ROWS_PER_PAGE, visibleIndices.length);

    // 表示するカラム（値があるカラムだけ表示）
    activeColsCache = MF_COLUMNS.filter(col => {
        return convertedData.some(row => row[col.key] !== undefined && row[col.key] !== '');
    });

    let html = '<table class="table"><thead><tr>';
    activeColsCache.forEach(col => {
        html += `<th>${col.label}</th>`;
    });
    html += '</tr></thead><tbody>';

    for (let vi = 0; vi < endIdx; vi++) {
        const i = visibleIndices[vi];
        const row = convertedData[i];
        const isError = errorRowSet.has(i);
        const isPattern = patternRowSet.has(i);
        const isReceipt = row._isReceipt;
        const isDefault = row._defaultAccount;
        const isGeminiAccount = row._geminiAccount;
        const rowClasses = [];
        if (isError) rowClasses.push('error-row');
        if (isReceipt) rowClasses.push('row-receipt');
        if (isPattern) rowClasses.push('row-pattern-applied');
        if (isGeminiAccount && !isPattern && !isDefault) rowClasses.push('row-gemini-account');
        if (isDefault && !isPattern) rowClasses.push('row-default-account');
        let rowTitle = '';
        if (isReceipt) rowTitle = row._documentTypeLabel || '領収書OCR';
        if (isPattern) rowTitle += (rowTitle ? ' / ' : '') + `パターン: ${row._matchedKeyword || ''}`;
        if (isDefault) rowTitle += (rowTitle ? ' / ' : '') + 'デフォルト科目適用';
        if (isGeminiAccount) rowTitle += (rowTitle ? ' / ' : '') + 'Gemini科目推測';
        html += `<tr${rowClasses.length ? ` class="${rowClasses.join(' ')}"` : ''} title="${escapeHtml(rowTitle)}">`;
        activeColsCache.forEach(col => {
            const val = String(row[col.key] || '');
            // 取引Noは編集不可
            if (col.key === 'torihikiNo') {
                html += `<td>${escapeHtml(val)}</td>`;
            } else {
                html += `<td class="editable-cell" data-row="${i}" data-field="${col.key}" onclick="startCellEdit(this)">${escapeHtml(val)}</td>`;
            }
        });
        html += '</tr>';
    }
    html += '</tbody></table>';

    container.innerHTML = html;
    displayedRows = endIdx;

    if (displayedRows < visibleIndices.length) {
        showMore.style.display = 'inline-flex';
        showMore.textContent = `さらに表示（残り ${visibleIndices.length - displayedRows} 行）`;
    } else {
        showMore.style.display = 'none';
    }
}

function showMoreResults() {
    renderConvertedPreview();
}

// --- セル編集 ---

function startCellEdit(td) {
    // 既に編集中なら何もしない
    if (td.classList.contains('editing')) return;

    const rowIdx = parseInt(td.dataset.row);
    const field = td.dataset.field;
    const currentValue = String(convertedData[rowIdx][field] || '');

    td.classList.add('editing');
    const input = document.createElement('input');
    input.type = 'text';
    input.className = 'cell-input';
    input.value = currentValue;
    td.textContent = '';
    td.appendChild(input);
    input.focus();
    input.select();

    // Enterで確定、Escでキャンセル
    input.addEventListener('keydown', function (e) {
        if (e.key === 'Enter') {
            finishCellEdit(td, rowIdx, field, currentValue, input.value);
        } else if (e.key === 'Escape') {
            cancelCellEdit(td, currentValue);
        }
    });

    // フォーカスを失ったら確定
    input.addEventListener('blur', function () {
        // ポップアップ表示中は無視（ポップアップのボタンクリックでblurが発火するため）
        setTimeout(() => {
            if (td.classList.contains('editing')) {
                finishCellEdit(td, rowIdx, field, currentValue, input.value);
            }
        }, 150);
    });
}

function cancelCellEdit(td, originalValue) {
    td.classList.remove('editing');
    td.textContent = originalValue;
}

function finishCellEdit(td, rowIdx, field, oldValue, newValue) {
    td.classList.remove('editing');
    newValue = newValue.trim();

    // 変更がなければそのまま戻す
    if (newValue === oldValue) {
        td.textContent = oldValue;
        return;
    }

    // データを更新
    convertedData[rowIdx][field] = newValue;
    td.textContent = newValue;
    td.classList.add('cell-corrected');

    // ルール保存ポップアップを表示
    if (oldValue !== '') {
        showRulePopup(field, oldValue, newValue);
    }
}

// --- ルール保存ポップアップ ---

function showRulePopup(field, oldValue, newValue) {
    const fieldLabel = MF_COLUMNS.find(c => c.key === field)?.label || field;
    pendingCorrection = { field, from: oldValue, to: newValue };

    const detail = document.getElementById('rulePopupDetail');
    const companyNote = selectedCompany
        ? `<span style="font-size: 12px;">「${escapeHtml(selectedCompany)}」のルールとして保存されます</span>`
        : `<span style="font-size: 12px; color: var(--warning);">会社を選択するとルール保存できます</span>`;

    detail.innerHTML = `[${escapeHtml(fieldLabel)}] の値<br>` +
        `<strong>${escapeHtml(oldValue)}</strong> → <strong>${escapeHtml(newValue)}</strong><br>` +
        companyNote;

    document.getElementById('rulePopup').classList.add('show');

    // 5秒後に自動で閉じる
    clearTimeout(showRulePopup._timer);
    showRulePopup._timer = setTimeout(dismissRulePopup, 8000);
}

function dismissRulePopup() {
    clearTimeout(showRulePopup._timer);
    document.getElementById('rulePopup').classList.remove('show');
    pendingCorrection = null;
}

function saveRuleFromPopup() {
    if (!pendingCorrection) return;

    if (!selectedCompany) {
        showToast('ルールを保存するには会社を選択してください');
        dismissRulePopup();
        return;
    }

    addCorrectionRule(selectedCompany, {
        field: pendingCorrection.field,
        from: pendingCorrection.from,
        to: pendingCorrection.to,
    });

    const fieldLabel = MF_COLUMNS.find(c => c.key === pendingCorrection.field)?.label || pendingCorrection.field;
    showToast(`[${selectedCompany}] 訂正ルールを保存: [${fieldLabel}] ${pendingCorrection.from} → ${pendingCorrection.to}`);

    dismissRulePopup();
    renderCorrectionRulesCard();
    updateCompanyRuleCount();
}

// --- 訂正ルール一覧カード ---

function renderCorrectionRulesCard() {
    const rules = getCorrectionRules(selectedCompany);
    const card = document.getElementById('correctionRulesCard');
    const list = document.getElementById('correctionRulesList');

    if (rules.length === 0) {
        card.style.display = 'none';
        return;
    }

    card.style.display = 'block';

    let html = '';
    rules.forEach(rule => {
        const fieldLabel = MF_COLUMNS.find(c => c.key === rule.field)?.label || rule.field;
        html += `<div class="correction-rule-item">
            <span class="correction-rule-field">${escapeHtml(fieldLabel)}</span>
            <span class="correction-rule-from">${escapeHtml(rule.from)}</span>
            <span class="mapping-arrow">→</span>
            <span class="correction-rule-to">${escapeHtml(rule.to)}</span>
            <button class="btn-icon" onclick="deleteCorrectionRuleAndRefresh('${rule.id}')" title="削除">✕</button>
        </div>`;
    });

    list.innerHTML = html;
}

function deleteCorrectionRuleAndRefresh(id) {
    deleteCorrectionRule(selectedCompany, id);
    renderCorrectionRulesCard();
    updateCompanyRuleCount();
    showToast('訂正ルールを削除しました');
}

function clearAllCorrectionRules() {
    const label = selectedCompany || 'すべて';
    if (!confirm(`「${label}」の訂正ルールをすべて削除しますか？`)) return;
    clearCorrectionRules(selectedCompany);
    renderCorrectionRulesCard();
    updateCompanyRuleCount();
    showToast('訂正ルールを削除しました');
}

function applyCorrectionsAndRerun() {
    runConversion();
    showToast('訂正ルールを適用して再変換しました');
}

// ===== エクスポート =====

function exportCSV() {
    if (convertedData.length === 0) {
        showToast('変換データがありません');
        return;
    }

    const csvText = generateMFCSV(convertedData);

    // UTF-8 BOM付きでダウンロード
    const bom = '\uFEFF';
    const blob = new Blob([bom + csvText], { type: 'text/csv;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    const dateStr = new Date().toISOString().slice(0, 10);
    const companyPrefix = selectedCompany ? `${selectedCompany}_` : '';
    a.download = `MF仕訳帳_${companyPrefix}${dateStr}.csv`;
    a.click();
    URL.revokeObjectURL(url);

    showToast('CSVをダウンロードしました');
}

// クリップボードにコピー（タブ区切りでExcelに直接貼り付け可能）
function copyToClipboard() {
    if (convertedData.length === 0) {
        showToast('変換データがありません');
        return;
    }

    const headers = MF_COLUMNS.map(c => c.label);
    const lines = [headers.join('\t')];

    convertedData.forEach(row => {
        const values = MF_COLUMNS.map(col => String(row[col.key] || ''));
        lines.push(values.join('\t'));
    });

    const text = lines.join('\n');

    navigator.clipboard.writeText(text).then(() => {
        showToast(`${convertedData.length} 行をクリップボードにコピーしました（Excelに貼り付け可能）`);
    }).catch(() => {
        // フォールバック：textarea経由でコピー
        const ta = document.createElement('textarea');
        ta.value = text;
        ta.style.position = 'fixed';
        ta.style.opacity = '0';
        document.body.appendChild(ta);
        ta.select();
        document.execCommand('copy');
        document.body.removeChild(ta);
        showToast(`${convertedData.length} 行をクリップボードにコピーしました`);
    });
}

// ===== テンプレート =====

function openTemplateSaveModal() {
    document.getElementById('templateNameInput').value = '';
    document.getElementById('templateSaveModal').classList.add('show');
}

function openTemplateLoadModal() {
    renderTemplateList();
    document.getElementById('templateLoadModal').classList.add('show');
}

function closeModal(id) {
    document.getElementById(id).classList.remove('show');
}

function saveCurrentTemplate() {
    const name = document.getElementById('templateNameInput').value.trim();
    if (!name) {
        showToast('テンプレート名を入力してください');
        return;
    }

    // 現在のマッピング・ルールを収集
    if (currentStep >= 2) collectMappingFromUI();
    if (currentStep >= 3) collectRulesFromUI();

    // マッピングの要約を生成
    const mappedFieldLabels = Object.values(columnMapping)
        .map(key => MF_COLUMNS.find(c => c.key === key)?.label || key)
        .filter((v, i, a) => a.indexOf(v) === i);

    const template = {
        name: name,
        company: selectedCompany || '',
        mapping: { ...columnMapping },
        rules: JSON.parse(JSON.stringify(conversionRules)),
        sourceHeaders: [...rawHeaders],
        mappingSummary: mappedFieldLabels.join('・'),
    };

    saveConverterTemplate(template);
    closeModal('templateSaveModal');
    showToast(`テンプレート「${name}」を保存しました`);
}

function renderTemplateList() {
    const templates = getConverterTemplates();
    const list = document.getElementById('templateList');
    const empty = document.getElementById('noTemplates');

    if (templates.length === 0) {
        list.style.display = 'none';
        empty.style.display = 'block';
        return;
    }

    list.style.display = 'block';
    empty.style.display = 'none';

    let html = '';
    templates.forEach(tpl => {
        const date = new Date(tpl.createdAt).toLocaleDateString('ja-JP');
        const companyLabel = tpl.company ? `[${escapeHtml(tpl.company)}] ` : '';
        const summary = tpl.mappingSummary || (tpl.sourceHeaders || []).join(', ').slice(0, 50);
        const fixedVals = tpl.rules?.fixedValues
            ? Object.entries(tpl.rules.fixedValues)
                .map(([k, v]) => `${MF_COLUMNS.find(c => c.key === k)?.label || k}=${v}`)
                .join(', ')
            : '';
        const fixedInfo = fixedVals ? ` | 固定値: ${escapeHtml(fixedVals)}` : '';

        html += `<div class="template-item" onclick="loadTemplate('${tpl.id}')">
            <div>
                <div class="template-item-name">${companyLabel}${escapeHtml(tpl.name)}</div>
                <div class="template-item-date">${date} | ${escapeHtml(summary)}${fixedInfo}</div>
            </div>
            <button class="btn-icon" onclick="event.stopPropagation(); deleteTemplateAndRefresh('${tpl.id}')" title="削除">🗑</button>
        </div>`;
    });

    list.innerHTML = html;
}

function loadTemplate(id) {
    const tpl = getConverterTemplate(id);
    if (!tpl) {
        showToast('テンプレートが見つかりません');
        return;
    }

    columnMapping = tpl.mapping || {};
    conversionRules = tpl.rules || { dateFormat: 'auto', accountMapping: {}, taxMapping: {}, fixedValues: {} };

    // 会社名も復元
    if (tpl.company) {
        const select = document.getElementById('companySelect');
        // 該当する会社がリストにあれば選択
        for (const opt of select.options) {
            if (opt.value === tpl.company) {
                select.value = tpl.company;
                selectedCompany = tpl.company;
                updateCompanyRuleCount();
                updateDeleteCompanyBtn();
                break;
            }
        }
    }

    closeModal('templateLoadModal');
    showToast(`テンプレート「${tpl.name}」を読み込みました`);
}

function deleteTemplateAndRefresh(id) {
    if (!confirm('このテンプレートを削除しますか？')) return;
    deleteConverterTemplate(id);
    renderTemplateList();
    showToast('テンプレートを削除しました');
}

// ===== リセット =====

function resetConverter() {
    if (!confirm('すべての入力内容をリセットしますか？')) return;

    rawHeaders = [];
    rawData = [];
    columnMapping = {};
    conversionRules = { dateFormat: 'auto', accountMapping: {}, taxMapping: {}, fixedValues: {} };
    convertedData = [];
    validationErrors = [];
    displayedRows = 0;

    // UI をリセット
    document.getElementById('dataPreview').style.display = 'none';
    document.getElementById('fileInfo').style.display = 'none';
    document.getElementById('fileInput').value = '';
    document.getElementById('pasteArea').value = '';
    document.getElementById('step1Next').disabled = true;

    goToStep(1);
    showToast('リセットしました');
}

// ===== 会社選択 =====

function renderCompanySelect() {
    const select = document.getElementById('companySelect');
    const companies = getCompanies();

    // 現在の選択を保持
    const current = select.value;

    select.innerHTML = '<option value="">（会社を選択）</option>';
    companies.forEach(name => {
        const option = document.createElement('option');
        option.value = name;
        option.textContent = name;
        if (name === current || name === selectedCompany) option.selected = true;
        select.appendChild(option);
    });

    selectedCompany = select.value;
    updateCompanyRuleCount();
    updateDeleteCompanyBtn();
    updateIndustrySelect();
    updateDefaultKashiSelect();
}

function onCompanyChange() {
    const select = document.getElementById('companySelect');
    selectedCompany = select.value;
    updateCompanyRuleCount();
    updateDeleteCompanyBtn();
    updateIndustrySelect();
    updateDefaultKashiSelect();

    // ステップ4にいる場合はカードを更新
    if (currentStep === 4) {
        renderCorrectionRulesCard();
        renderJournalPatternsCard();
    }
}

function updateCompanyRuleCount() {
    const countEl = document.getElementById('companyRuleCount');
    if (!selectedCompany) {
        countEl.textContent = '';
        return;
    }
    const rules = getCorrectionRules(selectedCompany);
    const patterns = getJournalPatterns(selectedCompany);
    const parts = [];
    if (rules.length > 0) parts.push(`訂正: ${rules.length}件`);
    if (patterns.length > 0) parts.push(`パターン: ${patterns.length}件`);
    countEl.textContent = parts.join(' / ');
}

function updateDeleteCompanyBtn() {
    const btn = document.getElementById('deleteCompanyBtn');
    btn.style.display = selectedCompany ? 'inline-flex' : 'none';
}

// --- 業種セレクトの初期構築 ---
function renderIndustrySelect() {
    const select = document.getElementById('industrySelect');
    select.innerHTML = '<option value="">（未設定）</option>';
    INDUSTRY_LIST.forEach(ind => {
        const opt = document.createElement('option');
        opt.value = ind;
        opt.textContent = ind;
        select.appendChild(opt);
    });
}

// --- 会社変更時に業種を同期 ---
function updateIndustrySelect() {
    const select = document.getElementById('industrySelect');
    if (selectedCompany) {
        // 会社に保存済みの業種をセット
        const saved = getCompanyIndustry(selectedCompany);
        select.value = saved || '';
        selectedIndustry = saved || '';
    }
}

function onIndustryChange() {
    const select = document.getElementById('industrySelect');
    selectedIndustry = select.value;

    // 会社が選択されている場合は会社に紐づけて保存
    if (selectedCompany) {
        setCompanyIndustry(selectedCompany, selectedIndustry);
    }

    showToast(selectedIndustry ? `業種を「${selectedIndustry}」に設定しました` : '業種設定を解除しました');

    // 変換済みデータがある場合は再変換
    if (convertedData.length > 0 && rawData.length > 0) {
        runConversion();
    }
}

function onDefaultKashiChange() {
    const select = document.getElementById('defaultKashiSelect');
    defaultKashiKamoku = select.value;

    // 会社が選択されている場合は保存
    if (selectedCompany) {
        const key = 'mf_converter_default_kashi_' + selectedCompany;
        if (defaultKashiKamoku) {
            localStorage.setItem(key, defaultKashiKamoku);
        } else {
            localStorage.removeItem(key);
        }
    }

    showToast(defaultKashiKamoku ? `貸方デフォルトを「${defaultKashiKamoku}」に設定しました` : '貸方デフォルトを自動に戻しました');

    // 変換済みデータがある場合は再変換
    if (convertedData.length > 0 && rawData.length > 0) {
        runConversion();
    }
}

// 会社変更時に貸方デフォルトを同期
function updateDefaultKashiSelect() {
    const select = document.getElementById('defaultKashiSelect');
    if (selectedCompany) {
        const saved = localStorage.getItem('mf_converter_default_kashi_' + selectedCompany);
        select.value = saved || '';
        defaultKashiKamoku = saved || '';
    }
}

function openAddCompanyPrompt() {
    const name = prompt('会社名を入力してください:');
    if (!name || !name.trim()) return;

    const trimmed = name.trim();
    addCompany(trimmed);
    renderCompanySelect();

    // 新しく追加した会社を選択
    document.getElementById('companySelect').value = trimmed;
    selectedCompany = trimmed;
    updateCompanyRuleCount();
    updateDeleteCompanyBtn();
    updateIndustrySelect();
    showToast(`「${trimmed}」を追加しました`);
}

function openDeleteCompanyPrompt() {
    if (!selectedCompany) return;
    const rules = getCorrectionRules(selectedCompany);
    const patterns = getJournalPatterns(selectedCompany);
    const details = [];
    if (rules.length > 0) details.push(`訂正ルール ${rules.length} 件`);
    if (patterns.length > 0) details.push(`仕訳パターン ${patterns.length} 件`);
    const msg = details.length > 0
        ? `「${selectedCompany}」を削除しますか？\n（${details.join('、')}も削除されます）`
        : `「${selectedCompany}」を削除しますか？`;

    if (!confirm(msg)) return;

    deleteCompany(selectedCompany);
    selectedCompany = '';
    renderCompanySelect();
    updateCompanyRuleCount();
    updateDeleteCompanyBtn();
    updateIndustrySelect();

    if (currentStep === 4) {
        renderCorrectionRulesCard();
        renderJournalPatternsCard();
    }
    showToast('会社を削除しました');
}

// ===== 仕訳パターン管理 =====

function renderJournalPatternsCard() {
    const card = document.getElementById('journalPatternsCard');
    const list = document.getElementById('journalPatternsList');
    const badge = document.getElementById('patternCountBadge');
    const info = document.getElementById('patternInfo');

    const patterns = getJournalPatterns(selectedCompany);

    // 会社未選択でもカードは表示（インポート案内のため）
    card.style.display = 'block';

    if (!selectedCompany) {
        badge.textContent = '';
        info.innerHTML = '<span style="color: var(--warning);">会社を選択すると仕訳パターンを管理できます</span>';
        list.innerHTML = '';
        return;
    }

    if (patterns.length === 0) {
        badge.textContent = '';
        info.innerHTML = '<span>過去月のMF仕訳帳CSVをインポートすると、摘要→科目のパターンを自動学習します</span>';
        list.innerHTML = '';
        return;
    }

    badge.textContent = patterns.length;
    const appliedCount = convertedData.filter(r => r._matchedPattern).length;
    info.innerHTML = appliedCount > 0
        ? `<span style="color: var(--accent);">✓ ${appliedCount} 行にパターンを適用しました</span>`
        : `<span>${patterns.length} 件のパターンが登録されています</span>`;

    // パターン一覧（キーワード順にソート）
    const sorted = [...patterns].sort((a, b) => a.keyword.localeCompare(b.keyword, 'ja'));
    let html = '';
    sorted.forEach(p => {
        const accounts = [];
        if (p.kariKamoku) accounts.push(`<span class="pattern-account-label">借方:</span>${escapeHtml(p.kariKamoku)}`);
        if (p.kashiKamoku) accounts.push(`<span class="pattern-account-label">貸方:</span>${escapeHtml(p.kashiKamoku)}`);
        if (p.kariZeiku) accounts.push(`<span class="pattern-account-label">税:</span>${escapeHtml(p.kariZeiku)}`);

        html += `<div class="journal-pattern-item">
            <span class="pattern-keyword">${escapeHtml(p.keyword)}</span>
            <span class="mapping-arrow">→</span>
            <span class="pattern-account">${accounts.join('　')}</span>
            <span class="pattern-count">×${p.count || 1}</span>
            <button class="pattern-delete" onclick="deletePatternAndRefresh('${p.id}')" title="削除">✕</button>
        </div>`;
    });
    list.innerHTML = html;
}

function importJournalPatterns() {
    if (!selectedCompany) {
        showToast('先に会社を選択してください');
        return;
    }
    document.getElementById('patternFileInput').click();
}

// 後方互換（旧関数名からの呼び出し対応）
function importJournalPatternsCSV() { importJournalPatterns(); }

function handlePatternFileSelect(input) {
    const file = input.files[0];
    if (!file) return;

    const isPDF = file.name.toLowerCase().endsWith('.pdf');

    if (isPDF) {
        // PDFファイル: ArrayBufferとして読み込み
        const reader = new FileReader();
        reader.onload = async function (e) {
            showToast('PDF解析中...', 'info');
            const result = await importPatternsFromPDF(selectedCompany, e.target.result);
            handlePatternImportResult(result);
        };
        reader.readAsArrayBuffer(file);
    } else {
        // CSV/テキストファイル: テキストとして読み込み
        const reader = new FileReader();
        reader.onload = function (e) {
            let text = e.target.result;

            // 文字化けチェック → Shift_JISリトライ
            if (hasGarbledChars(text)) {
                const retryReader = new FileReader();
                retryReader.onload = function (e2) {
                    processPatternCSV(e2.target.result);
                };
                retryReader.readAsText(file, 'Shift_JIS');
                return;
            }

            processPatternCSV(text);
        };
        reader.readAsText(file, 'UTF-8');
    }

    // input をリセット（同じファイルを再選択できるように）
    input.value = '';
}

function processPatternCSV(csvText) {
    const result = importPatternsFromMFCSV(selectedCompany, csvText);
    handlePatternImportResult(result);
}

function handlePatternImportResult(result) {
    if (result.error) {
        showToast(result.error);
        return;
    }

    showToast(`${result.count} 件の仕訳パターンを登録しました`);
    renderJournalPatternsCard();
    updateCompanyRuleCount();

    // パターンが追加されたら再変換
    if (result.count > 0 && convertedData.length > 0) {
        runConversion();
    }
}

function deletePatternAndRefresh(id) {
    deleteJournalPattern(selectedCompany, id);
    renderJournalPatternsCard();
    updateCompanyRuleCount();
}

function clearAllPatterns() {
    if (!selectedCompany) return;
    const patterns = getJournalPatterns(selectedCompany);
    if (patterns.length === 0) return;
    if (!confirm(`「${selectedCompany}」の仕訳パターン ${patterns.length} 件を全て削除しますか？`)) return;

    clearJournalPatterns(selectedCompany);
    renderJournalPatternsCard();
    updateCompanyRuleCount();
    showToast('仕訳パターンを全て削除しました');
}

function applyPatternsAndRerun() {
    if (convertedData.length === 0) {
        showToast('先にデータを変換してください');
        return;
    }
    runConversion();
    showToast('仕訳パターンを再適用しました');
}
