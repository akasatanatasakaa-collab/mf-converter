/* ===== 仕訳帳コンバーター 変換エンジン ===== */

// --- 区切り文字の自動検出 ---
function detectDelimiter(text) {
    const lines = text.split('\n').filter(l => l.trim()).slice(0, 10);
    if (lines.length === 0) return ',';

    const candidates = ['\t', ',', ';'];
    let bestDelimiter = ',';
    let bestScore = -1;

    candidates.forEach(d => {
        const counts = lines.map(line => {
            // ダブルクォート内の区切り文字はカウントしない
            let count = 0;
            let inQuote = false;
            for (let i = 0; i < line.length; i++) {
                if (line[i] === '"') inQuote = !inQuote;
                if (!inQuote && line[i] === d) count++;
            }
            return count;
        });

        if (counts[0] === 0) return;

        // 全行で同じカウントなら一貫性が高い
        const allSame = counts.every(c => c === counts[0]);
        const avgCount = counts.reduce((a, b) => a + b, 0) / counts.length;
        const score = allSame ? avgCount * 3 : avgCount;

        if (score > bestScore) {
            bestScore = score;
            bestDelimiter = d;
        }
    });

    return bestDelimiter;
}

// --- ヘッダー行の自動検出 ---
// タイトル行やセクション行をスキップし、実際のカラムヘッダー行を見つける
function detectHeaderRow(rows, maxScan) {
    const scanLimit = Math.min(maxScan || 10, rows.length);
    let bestIdx = 0;
    let bestScore = -1;

    // ヘッダーらしいキーワード
    const headerKeywords = [
        '日付', '年月日', '金額', '科目', '摘要', '内容', '取引', '備考', '税',
        '氏名', '名前', '担当', '部門', '残高', 'date', 'amount', '支出', '入金',
        '出金', '品名', '明細', 'No', '番号', '区分',
    ];

    for (let i = 0; i < scanLimit; i++) {
        const row = rows[i];
        if (!row) continue;

        let score = 0;
        const nonEmpty = row.filter(c => c.trim() !== '').length;

        // 非空セルが少なすぎる行はタイトル行の可能性が高い
        if (nonEmpty < 3) continue;

        // 非空セル数が多いほどスコアアップ
        score += nonEmpty * 2;

        // ヘッダーキーワードとの一致
        row.forEach(cell => {
            const c = cell.trim().toLowerCase();
            if (c === '') return;
            headerKeywords.forEach(kw => {
                if (c.includes(kw.toLowerCase())) score += 5;
            });
        });

        // セルが全部数値 or 全部空に近い行は非ヘッダー
        const numericCells = row.filter(c => /^[\d¥￥\\,.\-\s]+$/.test(c.trim()) && c.trim() !== '').length;
        if (nonEmpty > 0 && numericCells >= nonEmpty * 0.7) score -= 10;

        // 「前月残高」「合計」「例」などはヘッダーではない
        const joined = row.join('');
        if (/前月残高|合計|小計|例/.test(joined)) score -= 20;

        if (score > bestScore) {
            bestScore = score;
            bestIdx = i;
        }
    }

    return bestIdx;
}

// --- CSVパーサー（RFC 4180準拠） ---
function parseCSV(text, delimiter, hasHeader) {
    const rows = [];
    let currentRow = [];
    let currentField = '';
    let inQuote = false;
    let i = 0;

    while (i < text.length) {
        const ch = text[i];

        if (inQuote) {
            if (ch === '"') {
                // 次の文字もダブルクォートならエスケープ
                if (i + 1 < text.length && text[i + 1] === '"') {
                    currentField += '"';
                    i += 2;
                } else {
                    inQuote = false;
                    i++;
                }
            } else {
                currentField += ch;
                i++;
            }
        } else {
            if (ch === '"') {
                inQuote = true;
                i++;
            } else if (ch === delimiter) {
                currentRow.push(currentField);
                currentField = '';
                i++;
            } else if (ch === '\r') {
                // \r\n または \r を行末として扱う
                currentRow.push(currentField);
                currentField = '';
                if (currentRow.some(f => f.trim() !== '')) {
                    rows.push(currentRow);
                }
                currentRow = [];
                i++;
                if (i < text.length && text[i] === '\n') i++;
            } else if (ch === '\n') {
                currentRow.push(currentField);
                currentField = '';
                if (currentRow.some(f => f.trim() !== '')) {
                    rows.push(currentRow);
                }
                currentRow = [];
                i++;
            } else {
                currentField += ch;
                i++;
            }
        }
    }

    // 最後のフィールド・行を処理
    if (currentField || currentRow.length > 0) {
        currentRow.push(currentField);
        if (currentRow.some(f => f.trim() !== '')) {
            rows.push(currentRow);
        }
    }

    // ヘッダーとデータ行に分離
    let headers = [];
    let dataRows = rows;
    let headerRowIdx = 0;

    if (hasHeader && rows.length > 0) {
        // ヘッダー行を自動検出（タイトル行をスキップ）
        headerRowIdx = detectHeaderRow(rows, 10);
        headers = rows[headerRowIdx].map(h => h.trim());
        dataRows = rows.slice(headerRowIdx + 1);
    } else {
        // ヘッダーなしの場合、列番号をヘッダーにする
        const maxCols = Math.max(...rows.map(r => r.length), 0);
        headers = Array.from({ length: maxCols }, (_, i) => `列${i + 1}`);
    }

    return { headers, rows: dataRows, headerRowIdx };
}

// --- 日付形式の自動検出 ---
function detectDateFormat(samples) {
    const patterns = [
        { id: 'yyyy/MM/dd', regex: /^\d{4}\/\d{1,2}\/\d{1,2}$/ },
        { id: 'yyyy-MM-dd', regex: /^\d{4}-\d{1,2}-\d{1,2}$/ },
        { id: 'yyyyMMdd',   regex: /^\d{8}$/ },
        { id: 'MM/dd/yyyy', regex: /^\d{1,2}\/\d{1,2}\/\d{4}$/ },
        { id: 'yyyy年M月d日', regex: /^\d{4}年\d{1,2}月\d{1,2}日$/ },
        { id: 'wareki',     regex: /^(令和|R)\d{1,2}年\d{1,2}月\d{1,2}日?$/ },
        { id: 'M/d',        regex: /^\d{1,2}\/\d{1,2}$/ },
        { id: 'M-d',        regex: /^\d{1,2}-\d{1,2}$/ },
        { id: 'M月d日',      regex: /^\d{1,2}月\d{1,2}日$/ },
    ];

    // 空でないサンプルでマッチを試みる
    const validSamples = samples.filter(s => s && String(s).trim());
    if (validSamples.length === 0) return 'auto';

    for (const pattern of patterns) {
        const matchCount = validSamples.filter(s => pattern.regex.test(String(s).trim())).length;
        if (matchCount >= validSamples.length * 0.7) {
            return pattern.id;
        }
    }

    return 'auto';
}

// --- 日付変換 ---
function convertDate(value, sourceFormat) {
    if (!value) return '';
    const v = String(value).trim();
    if (!v) return '';

    let year, month, day;

    // 自動検出の場合、順番に試す
    if (sourceFormat === 'auto') {
        const formats = ['yyyy/MM/dd', 'yyyy/M/d', 'yyyy-MM-dd', 'yyyyMMdd', 'yyyy年M月d日', 'wareki', 'MM/dd/yyyy', 'M/d', 'M-d', 'M月d日'];
        for (const fmt of formats) {
            const result = convertDate(v, fmt);
            if (result) return result;
        }
        return '';
    }

    try {
        switch (sourceFormat) {
            case 'yyyy/MM/dd':
            case 'yyyy/M/d': {
                const parts = v.split('/');
                if (parts.length !== 3) return '';
                year = parseInt(parts[0]);
                month = parseInt(parts[1]);
                day = parseInt(parts[2]);
                break;
            }
            case 'yyyy-MM-dd': {
                const parts = v.split('-');
                if (parts.length !== 3) return '';
                year = parseInt(parts[0]);
                month = parseInt(parts[1]);
                day = parseInt(parts[2]);
                break;
            }
            case 'yyyyMMdd': {
                if (v.length !== 8 || !/^\d{8}$/.test(v)) return '';
                year = parseInt(v.slice(0, 4));
                month = parseInt(v.slice(4, 6));
                day = parseInt(v.slice(6, 8));
                break;
            }
            case 'MM/dd/yyyy': {
                const parts = v.split('/');
                if (parts.length !== 3) return '';
                month = parseInt(parts[0]);
                day = parseInt(parts[1]);
                year = parseInt(parts[2]);
                break;
            }
            case 'yyyy年M月d日': {
                const match = v.match(/(\d{4})年(\d{1,2})月(\d{1,2})日/);
                if (!match) return '';
                year = parseInt(match[1]);
                month = parseInt(match[2]);
                day = parseInt(match[3]);
                break;
            }
            case 'wareki': {
                const match = v.match(/(?:令和|R)(\d{1,2})年(\d{1,2})月(\d{1,2})日?/);
                if (!match) return '';
                year = 2018 + parseInt(match[1]);
                month = parseInt(match[2]);
                day = parseInt(match[3]);
                break;
            }
            case 'M/d': {
                // 年なし日付（例: 2/1, 12/25）→ 今年の日付として扱う
                const parts = v.split('/');
                if (parts.length !== 2) return '';
                month = parseInt(parts[0]);
                day = parseInt(parts[1]);
                year = new Date().getFullYear();
                break;
            }
            case 'M-d': {
                const parts = v.split('-');
                if (parts.length !== 2) return '';
                month = parseInt(parts[0]);
                day = parseInt(parts[1]);
                year = new Date().getFullYear();
                break;
            }
            case 'M月d日': {
                const match = v.match(/(\d{1,2})月(\d{1,2})日/);
                if (!match) return '';
                month = parseInt(match[1]);
                day = parseInt(match[2]);
                year = new Date().getFullYear();
                break;
            }
            default:
                return '';
        }
    } catch (e) {
        return '';
    }

    // バリデーション
    if (!year || !month || !day) return '';
    if (year < 1900 || year > 2100) return '';
    if (month < 1 || month > 12) return '';
    if (day < 1 || day > 31) return '';

    return `${year}/${String(month).padStart(2, '0')}/${String(day).padStart(2, '0')}`;
}

// --- 金額正規化 ---
function normalizeAmount(value) {
    if (value === null || value === undefined || value === '') return '';
    let str = String(value).trim();
    if (!str) return '';

    // 全角数字→半角
    str = str.replace(/[０-９]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFF10 + 0x30));
    // 全角マイナス→半角
    str = str.replace(/ー/g, '-');
    // 円記号・バックスラッシュ除去
    str = str.replace(/[¥￥\\]/g, '');
    // 3桁カンマ除去
    str = str.replace(/,/g, '');
    // スペース除去
    str = str.trim();

    const num = parseFloat(str);
    if (isNaN(num)) return '';

    return Math.round(num);
}

// --- 勘定科目変換 ---
function mapAccountName(inputName, customMapping) {
    if (!inputName) return '';
    const trimmed = String(inputName).trim();
    if (!trimmed) return '';

    // カスタムマッピングを優先
    if (customMapping && customMapping[trimmed]) return customMapping[trimmed];

    // プリセットの別名テーブル
    if (ACCOUNT_ALIASES[trimmed]) return ACCOUNT_ALIASES[trimmed];

    // 該当なしはそのまま返す
    return trimmed;
}

// --- 税区分変換 ---
function mapTaxCategory(inputValue, customMapping) {
    if (!inputValue) return '';
    const trimmed = String(inputValue).trim();
    if (!trimmed) return '';

    if (customMapping && customMapping[trimmed]) return customMapping[trimmed];
    return trimmed;
}

// --- メイン変換処理 ---
function convertToMFFormat(rawRows, headers, mapping, rules) {
    const errors = [];
    const convertedRows = [];
    let skippedRows = 0;
    let lastValidDate = ''; // 日付引き継ぎ用
    let skipRemainder = false; // 例データセクション検出後は全スキップ

    rawRows.forEach((row, rowIdx) => {
        // 例データマーカー以降は全てスキップ
        if (skipRemainder) {
            skippedRows++;
            return;
        }

        const mfRow = {};

        // マッピングに従ってフィールドを割り当て
        for (const [sourceIdx, mfField] of Object.entries(mapping)) {
            const idx = parseInt(sourceIdx);
            const value = idx < row.length ? row[idx] : '';
            mfRow[mfField] = value;
        }

        // 「例」マーカー行を検出したら以降全てスキップ
        if (row.some(cell => /^↓?例$/.test(String(cell).trim()))) {
            skipRemainder = true;
            skippedRows++;
            return;
        }

        // 不要行スキップ：合計行、前月残高行、セクション見出しなどを除外
        const rowJoined = row.join('');
        if (/^(合計|小計|総計|前月残高|【.*】)/.test(rowJoined.trim()) ||
            /^[\s,]*$/.test(rowJoined)) {
            skippedRows++;
            return;
        }
        if (row.some(cell => /^前月残高$/.test(String(cell).trim()))) {
            skippedRows++;
            return;
        }

        // ヘッダー行の再出現をスキップ（例データのサブヘッダー等）
        const mappedHeaderCount = Object.entries(mapping).filter(([idx, _]) => {
            const val = (row[parseInt(idx)] || '').trim();
            return val && headers.includes(val);
        }).length;
        if (mappedHeaderCount >= 2) {
            skippedRows++;
            return;
        }

        // 空行スキップ：日付も金額もない行は除外
        const hasDate = mfRow.torihikiDate && String(mfRow.torihikiDate).trim();
        const hasKari = mfRow.kariKingaku && String(mfRow.kariKingaku).trim();
        const hasKashi = mfRow.kashiKingaku && String(mfRow.kashiKingaku).trim();
        if (!hasDate && !hasKari && !hasKashi) {
            skippedRows++;
            return; // この行をスキップ
        }

        // 日付も摘要もない行は合計行の可能性が高いのでスキップ
        // （日付引き継ぎは摘要や内容がある場合のみ適用）
        const hasText = (mfRow.tekiyou && String(mfRow.tekiyou).trim()) ||
                        (mfRow.torihikisaki && String(mfRow.torihikisaki).trim()) ||
                        (mfRow.kariKamoku && String(mfRow.kariKamoku).trim()) ||
                        (mfRow.kashiKamoku && String(mfRow.kashiKamoku).trim()) ||
                        (mfRow.memo && String(mfRow.memo).trim());
        if (!hasDate && !hasText && (hasKari || hasKashi)) {
            skippedRows++;
            return; // 合計行・集計行として除外
        }

        // 日付引き継ぎ：日付がないが金額+摘要がある場合、前の行の日付を使う
        if (!hasDate && (hasKari || hasKashi) && lastValidDate) {
            mfRow.torihikiDate = lastValidDate;
        }

        // 固定値を適用
        if (rules.fixedValues) {
            for (const [field, value] of Object.entries(rules.fixedValues)) {
                if (value) mfRow[field] = value;
            }
        }

        // 日付変換
        if (mfRow.torihikiDate) {
            const originalDateStr = String(mfRow.torihikiDate).trim(); // 変換前の値を保存
            const converted = convertDate(mfRow.torihikiDate, rules.dateFormat || 'auto');
            if (!converted) {
                errors.push({ row: rowIdx, field: 'torihikiDate', message: `日付変換エラー: "${mfRow.torihikiDate}"` });
            }
            mfRow.torihikiDate = converted;
            // 日付引き継ぎ用に変換前の値を保存
            if (converted) lastValidDate = originalDateStr;
        }

        // 金額変換
        if (mfRow.kariKingaku !== undefined && mfRow.kariKingaku !== '') {
            const amount = normalizeAmount(mfRow.kariKingaku);
            if (amount === '' && mfRow.kariKingaku) {
                errors.push({ row: rowIdx, field: 'kariKingaku', message: `金額変換エラー: "${mfRow.kariKingaku}"` });
            }
            mfRow.kariKingaku = amount;
        }
        if (mfRow.kashiKingaku !== undefined && mfRow.kashiKingaku !== '') {
            const amount = normalizeAmount(mfRow.kashiKingaku);
            if (amount === '' && mfRow.kashiKingaku) {
                errors.push({ row: rowIdx, field: 'kashiKingaku', message: `金額変換エラー: "${mfRow.kashiKingaku}"` });
            }
            mfRow.kashiKingaku = amount;
        }

        // マイナス金額の振分：借方がマイナスなら貸方に移動（逆も同様）
        if (mfRow.kariKingaku && mfRow.kariKingaku < 0 && !mfRow.kashiKingaku) {
            mfRow.kashiKingaku = Math.abs(mfRow.kariKingaku);
            mfRow.kariKingaku = '';
        }
        if (mfRow.kashiKingaku && mfRow.kashiKingaku < 0 && !mfRow.kariKingaku) {
            mfRow.kariKingaku = Math.abs(mfRow.kashiKingaku);
            mfRow.kashiKingaku = '';
        }

        // 借方金額のみの場合、貸方金額を同額にする
        if (mfRow.kariKingaku && !mfRow.kashiKingaku) {
            mfRow.kashiKingaku = mfRow.kariKingaku;
        }
        if (mfRow.kashiKingaku && !mfRow.kariKingaku) {
            mfRow.kariKingaku = mfRow.kashiKingaku;
        }

        // 勘定科目変換
        if (mfRow.kariKamoku) {
            mfRow.kariKamoku = mapAccountName(mfRow.kariKamoku, rules.accountMapping || {});
        }
        if (mfRow.kashiKamoku) {
            mfRow.kashiKamoku = mapAccountName(mfRow.kashiKamoku, rules.accountMapping || {});
        }

        // 税区分変換
        if (mfRow.kariZeiku) {
            mfRow.kariZeiku = mapTaxCategory(mfRow.kariZeiku, rules.taxMapping || {});
        }
        if (mfRow.kashiZeiku) {
            mfRow.kashiZeiku = mapTaxCategory(mfRow.kashiZeiku, rules.taxMapping || {});
        }

        // 仕訳パターンを適用（空のフィールドのみ埋める）
        if (rules.journalPatterns && rules.journalPatterns.length > 0) {
            const matched = applyJournalPatterns(mfRow, rules.journalPatterns);
            if (matched) {
                mfRow._matchedPattern = matched.id;
                mfRow._matchedKeyword = matched.keyword;
            }
        }

        // 業種デフォルト勘定科目を適用（仕訳パターンで埋まらなかった場合のフォールバック）
        // 業種未設定でも「共通」項目は適用される
        if (!mfRow.kariKamoku && mfRow.tekiyou) {
            const defaultKamoku = matchDefaultAccount(mfRow.tekiyou, rules.industry || '');
            if (defaultKamoku) {
                mfRow.kariKamoku = defaultKamoku;
                mfRow._defaultAccount = true;
            }
        }

        // デフォルト貸方勘定科目を適用（仕訳パターンで埋まらなかった場合）
        if (!mfRow.kashiKamoku && rules.defaultKashiKamoku) {
            mfRow.kashiKamoku = rules.defaultKashiKamoku;
        }

        // 訂正ルールを適用
        applyCorrectionRules(mfRow, rules.correctionRules || []);

        // 取引No採番
        mfRow.torihikiNo = rowIdx + 1;

        convertedRows.push(mfRow);
    });

    // バリデーション
    convertedRows.forEach((row, rowIdx) => {
        const rowErrors = validateMFRow(row, rowIdx);
        errors.push(...rowErrors);
    });

    return { rows: convertedRows, errors, skippedRows };
}

// --- バリデーション ---
function validateMFRow(row, rowIndex) {
    const errors = [];

    // 取引日は必須
    if (!row.torihikiDate) {
        errors.push({ row: rowIndex, field: 'torihikiDate', message: '取引日が未入力です' });
    }

    // 借方金額 or 貸方金額のどちらかは必要
    if (!row.kariKingaku && !row.kashiKingaku) {
        errors.push({ row: rowIndex, field: 'kariKingaku', message: '金額が未入力です' });
    }

    // 借方勘定科目は必須（金額がある場合）
    if (row.kariKingaku && !row.kariKamoku) {
        errors.push({ row: rowIndex, field: 'kariKamoku', message: '借方勘定科目が未入力です' });
    }

    // 貸方勘定科目は必須（金額がある場合）
    if (row.kashiKingaku && !row.kashiKamoku) {
        errors.push({ row: rowIndex, field: 'kashiKamoku', message: '貸方勘定科目が未入力です' });
    }

    return errors;
}

// --- MF仕訳帳CSV生成 ---
function generateMFCSV(rows) {
    const headers = MF_COLUMNS.map(c => c.label);
    const csvLines = [headers.join(',')];

    rows.forEach(row => {
        const values = MF_COLUMNS.map(col => {
            const val = String(row[col.key] || '');
            // カンマ、ダブルクォート、改行を含む場合はエスケープ
            if (val.includes(',') || val.includes('"') || val.includes('\n') || val.includes('\r')) {
                return '"' + val.replace(/"/g, '""') + '"';
            }
            return val;
        });
        csvLines.push(values.join(','));
    });

    return csvLines.join('\r\n');
}

// --- 仕訳パターン適用 ---
// 摘要のキーワード部分一致で過去仕訳パターンを検索し、空フィールドを埋める
function applyJournalPatterns(mfRow, patterns) {
    if (!patterns || patterns.length === 0) return null;

    const tekiyou = String(mfRow.tekiyou || '').trim();
    if (!tekiyou) return null;

    // 最長キーワード一致を優先（長い順にソート）
    const sorted = [...patterns].sort((a, b) => b.keyword.length - a.keyword.length);

    for (const pattern of sorted) {
        if (tekiyou.includes(pattern.keyword)) {
            // 空のフィールドのみ埋める（既存値は上書きしない）
            const fields = ['kariKamoku', 'kashiKamoku', 'kariZeiku', 'kashiZeiku', 'kariHojo', 'kashiHojo'];
            fields.forEach(field => {
                if (!mfRow[field] && pattern[field]) {
                    mfRow[field] = pattern[field];
                }
            });
            return pattern; // マッチしたパターンを返す
        }
    }

    return null;
}

// --- 訂正ルール適用 ---
function applyCorrectionRules(mfRow, correctionRules) {
    if (!correctionRules || correctionRules.length === 0) return;

    correctionRules.forEach(rule => {
        const currentVal = String(mfRow[rule.field] || '').trim();
        if (currentVal === rule.from) {
            mfRow[rule.field] = rule.to;
        }
    });
}

// --- カラムの自動マッピング推定（ヘッダー名ベース） ---
function autoDetectMapping(headers) {
    const mapping = {};

    headers.forEach((header, idx) => {
        const normalized = header.trim().toLowerCase();
        for (const [mfField, keywords] of Object.entries(MAPPING_HINTS)) {
            if (keywords.some(k => normalized.includes(k.toLowerCase()))) {
                // 重複チェック
                if (!Object.values(mapping).includes(mfField)) {
                    mapping[idx] = mfField;
                    break;
                }
            }
        }
    });

    return mapping;
}

// --- データ内容ベースのカラム自動判定 ---
// ヘッダー名が不明でも、値のパターンから日付・金額・テキストを判定
function autoDetectMappingByContent(headers, rows) {
    const sampleRows = rows.slice(0, 20);
    if (sampleRows.length === 0) return {};

    const colCount = headers.length;
    const colInfo = [];

    // 各カラムの内容を分析
    for (let col = 0; col < colCount; col++) {
        const values = sampleRows.map(r => (r[col] || '').trim()).filter(v => v !== '');
        colInfo.push({
            idx: col,
            header: headers[col],
            values: values,
            type: classifyColumn(values),
        });
    }

    const mapping = {};
    const used = new Set();

    // 1. 日付カラムを探す（最初に見つかったもの）
    const dateCol = colInfo.find(c => c.type === 'date' && !used.has(c.idx));
    if (dateCol) {
        mapping[dateCol.idx] = 'torihikiDate';
        used.add(dateCol.idx);
    }

    // 2. 金額カラムを探す
    const amountCols = colInfo.filter(c => c.type === 'amount' && !used.has(c.idx));

    if (amountCols.length === 1) {
        // 金額カラムが1つ → 借方金額（貸方は自動複写される）
        mapping[amountCols[0].idx] = 'kariKingaku';
        used.add(amountCols[0].idx);
    } else if (amountCols.length >= 2) {
        // 金額カラムが2つ以上 → 出金/入金 と推定
        // ヘッダーに「出」「支」「引」があれば借方、「入」「収」「預」があれば貸方
        let kariAssigned = false;
        let kashiAssigned = false;

        amountCols.forEach(c => {
            const h = c.header.toLowerCase();
            if (!kariAssigned && (h.includes('出') || h.includes('支') || h.includes('引') || h.includes('利用'))) {
                mapping[c.idx] = 'kariKingaku';
                used.add(c.idx);
                kariAssigned = true;
            } else if (!kashiAssigned && (h.includes('入') || h.includes('収') || h.includes('預'))) {
                mapping[c.idx] = 'kashiKingaku';
                used.add(c.idx);
                kashiAssigned = true;
            }
        });

        // ヘッダーでは判別できなかった場合、位置で割当
        if (!kariAssigned && !kashiAssigned) {
            mapping[amountCols[0].idx] = 'kariKingaku';
            used.add(amountCols[0].idx);
            if (amountCols.length >= 2) {
                mapping[amountCols[1].idx] = 'kashiKingaku';
                used.add(amountCols[1].idx);
            }
        } else if (!kariAssigned) {
            const remaining = amountCols.find(c => !used.has(c.idx));
            if (remaining) {
                mapping[remaining.idx] = 'kariKingaku';
                used.add(remaining.idx);
            }
        } else if (!kashiAssigned) {
            const remaining = amountCols.find(c => !used.has(c.idx));
            if (remaining) {
                mapping[remaining.idx] = 'kashiKingaku';
                used.add(remaining.idx);
            }
        }
    }

    // 3. テキストカラムから摘要・取引先を割当
    const textCols = colInfo.filter(c => c.type === 'text' && !used.has(c.idx));
    if (textCols.length >= 1) {
        // 最も値が多様なカラムを摘要にする
        const sorted = textCols.sort((a, b) => {
            const uniqueA = new Set(a.values).size;
            const uniqueB = new Set(b.values).size;
            return uniqueB - uniqueA;
        });
        mapping[sorted[0].idx] = 'tekiyou';
        used.add(sorted[0].idx);

        if (sorted.length >= 2) {
            mapping[sorted[1].idx] = 'torihikisaki';
            used.add(sorted[1].idx);
        }
    }

    // 4. 勘定科目っぽいカラムを探す（よく使う勘定科目との一致率で判定）
    const remainingText = colInfo.filter(c => c.type === 'text' && !used.has(c.idx));
    for (const col of remainingText) {
        const matchCount = col.values.filter(v =>
            COMMON_ACCOUNTS.includes(v) || ACCOUNT_ALIASES[v]
        ).length;
        if (matchCount >= col.values.length * 0.3 && col.values.length >= 2) {
            mapping[col.idx] = 'kariKamoku';
            used.add(col.idx);
            break;
        }
    }

    return mapping;
}

// カラムの値から型を判定
function classifyColumn(values) {
    if (values.length === 0) return 'empty';

    let dateCount = 0;
    let amountCount = 0;

    values.forEach(v => {
        // 日付判定
        if (isDateLike(v)) {
            dateCount++;
        }
        // 金額判定
        if (isAmountLike(v)) {
            amountCount++;
        }
    });

    const threshold = values.length * 0.6;
    if (dateCount >= threshold) return 'date';
    if (amountCount >= threshold) return 'amount';
    return 'text';
}

// 日付っぽい値かどうか
function isDateLike(value) {
    const v = value.trim();
    // yyyy/MM/dd, yyyy-MM-dd, yyyyMMdd, yyyy年M月d日, 令和x年, MM/dd/yyyy
    if (/^\d{4}[\/\-]\d{1,2}[\/\-]\d{1,2}$/.test(v)) return true;
    if (/^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4}$/.test(v)) return true;
    if (/^\d{8}$/.test(v)) return true;
    if (/^\d{4}年\d{1,2}月\d{1,2}日$/.test(v)) return true;
    if (/^(令和|R)\d{1,2}年\d{1,2}月\d{1,2}日?$/.test(v)) return true;
    // 年なし日付: M/d, M-d, M月d日（月1-12, 日1-31の範囲チェック）
    const mdSlash = v.match(/^(\d{1,2})\/(\d{1,2})$/);
    if (mdSlash) {
        const m = parseInt(mdSlash[1]), d = parseInt(mdSlash[2]);
        if (m >= 1 && m <= 12 && d >= 1 && d <= 31) return true;
    }
    const mdHyphen = v.match(/^(\d{1,2})-(\d{1,2})$/);
    if (mdHyphen) {
        const m = parseInt(mdHyphen[1]), d = parseInt(mdHyphen[2]);
        if (m >= 1 && m <= 12 && d >= 1 && d <= 31) return true;
    }
    if (/^\d{1,2}月\d{1,2}日$/.test(v)) return true;
    return false;
}

// 金額っぽい値かどうか
function isAmountLike(value) {
    let v = value.trim();
    // 全角数字→半角
    v = v.replace(/[０-９]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFF10 + 0x30));
    // 円記号・カンマ・スペース除去
    v = v.replace(/[¥￥\\,、 　]/g, '');
    // マイナス記号の正規化
    v = v.replace(/^[ー−‐–—]/g, '-');
    // 残りが数値か
    return /^-?\d+(\.\d+)?$/.test(v) && v.length >= 1;
}

// ===== 領収書OCR テキスト解析 =====

// --- 領収書テキストから構造化データを抽出 ---
function parseReceiptText(ocrText) {
    if (!ocrText || !ocrText.trim()) {
        return { confidence: 0 };
    }

    const lines = ocrText.split('\n').map(l => l.trim()).filter(l => l);
    const fullText = lines.join(' ');

    const date = extractReceiptDate(lines, fullText);
    const storeName = extractStoreName(lines);
    const { total, totalLine } = extractReceiptTotal(lines, fullText);
    const tax = extractReceiptTax(lines, fullText);
    const taxRate = detectTaxRate(total, tax);
    const items = extractReceiptItems(lines, totalLine);

    // 信頼度スコア（抽出できた項目数に基づく）
    let confidence = 0;
    if (date) confidence += 0.3;
    if (storeName) confidence += 0.2;
    if (total) confidence += 0.3;
    if (items.length > 0) confidence += 0.1;
    if (tax) confidence += 0.1;

    console.log('[領収書解析]', { date, storeName, total, tax, taxRate, items, confidence });

    return {
        date,
        storeName,
        totalAmount: total,
        tax,
        taxRate,
        items,
        confidence,
        rawText: ocrText
    };
}

// --- 日付抽出 ---
function extractReceiptDate(lines, fullText) {
    const patterns = [
        // 西暦: 2026/2/15, 2026-02-15, 2026年2月15日, 2026.2.15
        /(\d{4})\s*[年\/\-\.]\s*(\d{1,2})\s*[月\/\-\.]\s*(\d{1,2})\s*日?/,
        // 令和: 令和8年2月15日
        /令和\s*(\d{1,2})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日?/,
        // R8.2.15, R8/2/15
        /R\.?\s*(\d{1,2})\s*[\.\/]\s*(\d{1,2})\s*[\.\/]\s*(\d{1,2})/,
    ];

    for (const line of lines) {
        // 西暦パターン
        const m1 = line.match(patterns[0]);
        if (m1) {
            return `${m1[1]}/${String(m1[2]).padStart(2, '0')}/${String(m1[3]).padStart(2, '0')}`;
        }

        // 令和パターン
        const m2 = line.match(patterns[1]);
        if (m2) {
            const year = 2018 + parseInt(m2[1]);
            return `${year}/${String(m2[2]).padStart(2, '0')}/${String(m2[3]).padStart(2, '0')}`;
        }

        // R略記パターン
        const m3 = line.match(patterns[2]);
        if (m3) {
            const year = 2018 + parseInt(m3[1]);
            return `${year}/${String(m3[2]).padStart(2, '0')}/${String(m3[3]).padStart(2, '0')}`;
        }
    }

    return '';
}

// --- 合計金額抽出 ---
function extractReceiptTotal(lines, fullText) {
    // 合計キーワードを含む行から金額を探す
    const totalKeywords = /(?:合\s*計|お買い?上げ?|お会計|総合計|ご請求|税込|お預り|お支払|支払\s*額|請求\s*額)/;
    const amountPattern = /[¥￥]?\s*([0-9０-９,，]+)\s*円?/;

    let bestAmount = 0;
    let bestLine = -1;

    lines.forEach((line, idx) => {
        if (totalKeywords.test(line)) {
            const m = line.match(amountPattern);
            if (m) {
                const amount = parseReceiptAmount(m[1]);
                // 最大金額を合計とみなす（小計 < 合計 < お預り の場合）
                if (amount > bestAmount) {
                    bestAmount = amount;
                    bestLine = idx;
                }
            }
        }
    });

    // キーワードで見つからない場合、最大金額を合計とみなす
    if (bestAmount === 0) {
        lines.forEach((line, idx) => {
            const m = line.match(amountPattern);
            if (m) {
                const amount = parseReceiptAmount(m[1]);
                if (amount > bestAmount) {
                    bestAmount = amount;
                    bestLine = idx;
                }
            }
        });
    }

    return { total: bestAmount || '', totalLine: bestLine };
}

// --- 税額抽出 ---
function extractReceiptTax(lines, fullText) {
    const taxKeywords = /(?:消費税|内\s*税|うち\s*税|税\s*額|外\s*税)/;
    const amountPattern = /[¥￥]?\s*([0-9０-９,，]+)\s*円?/;

    for (const line of lines) {
        if (taxKeywords.test(line)) {
            const m = line.match(amountPattern);
            if (m) {
                return parseReceiptAmount(m[1]);
            }
        }
    }
    return '';
}

// --- 税率判定 ---
function detectTaxRate(total, tax) {
    if (!total || !tax) return '';
    const rate = tax / (total - tax);
    if (Math.abs(rate - 0.10) < 0.02) return '10%';
    if (Math.abs(rate - 0.08) < 0.02) return '8%';
    // 内税計算: tax / total
    const innerRate = tax / total;
    if (Math.abs(innerRate - 10 / 110) < 0.02) return '10%';
    if (Math.abs(innerRate - 8 / 108) < 0.02) return '8%';
    return '';
}

// --- 店名抽出 ---
function extractStoreName(lines) {
    // 先頭5行から店名らしき行を探す
    const storeKeywords = /(?:株式会社|有限会社|合同会社|（株）|\(株\)|店|屋|商店|薬局|スーパー|マート|ストア|センター|サービス)/;
    const excludePatterns = /(?:TEL|tel|電話|〒|\d{3}-\d{4}|http|www|レジ|No|担当|領収)/;

    const searchLines = lines.slice(0, Math.min(lines.length, 8));

    for (const line of searchLines) {
        if (excludePatterns.test(line)) continue;
        if (line.length < 2 || line.length > 30) continue;

        if (storeKeywords.test(line)) {
            // 「株式会社」等の前後の不要部分を除去
            return line.replace(/^\s*[　\s]*/, '').replace(/[　\s]*$/, '');
        }
    }

    // キーワードなしでも先頭の短い行（2〜20文字）を店名候補とする
    for (const line of searchLines) {
        if (excludePatterns.test(line)) continue;
        if (line.length >= 2 && line.length <= 20 && !/^\d+$/.test(line)) {
            return line;
        }
    }

    return '';
}

// --- 品目抽出 ---
function extractReceiptItems(lines, totalLineIdx) {
    const items = [];
    const amountPattern = /[¥￥]?\s*([0-9０-９,，]+)\s*円?$/;
    const excludeKeywords = /(?:合\s*計|小\s*計|消費税|内\s*税|お釣|お預|お買い?上|お会計|税込|税抜|支払|請求|領収|ポイント|外\s*税)/;

    // 合計行より前の行から品目を探す
    const endLine = totalLineIdx > 0 ? totalLineIdx : lines.length;

    for (let i = 0; i < endLine; i++) {
        const line = lines[i];
        if (excludeKeywords.test(line)) continue;

        const m = line.match(amountPattern);
        if (m) {
            const amount = parseReceiptAmount(m[1]);
            const name = line.replace(amountPattern, '').trim();
            if (name && amount > 0 && amount < 1000000) {
                items.push({ name, amount });
            }
        }
    }

    return items;
}

// --- 金額文字列をパース ---
function parseReceiptAmount(str) {
    if (!str) return 0;
    // 全角数字→半角
    let s = String(str).replace(/[０-９]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFF10 + 0x30));
    // カンマ・スペース除去
    s = s.replace(/[,，\s]/g, '');
    const num = parseInt(s, 10);
    return isNaN(num) ? 0 : num;
}

// --- 領収書データからMF仕訳行を生成 ---
function buildMFRowFromReceipt(receiptData, journalPatterns, correctionRules, industry, defaultKashi) {
    const mfRow = {};

    // 取引日
    mfRow.torihikiDate = receiptData.date || '';

    // 金額（借方=貸方で同額）
    mfRow.kariKingaku = receiptData.totalAmount || '';
    mfRow.kashiKingaku = receiptData.totalAmount || '';

    // 取引先
    mfRow.torihikisaki = receiptData.storeName || '';

    // 摘要を組み立て
    const parts = [];
    if (receiptData.storeName) parts.push(receiptData.storeName);
    if (receiptData.items && receiptData.items.length > 0) {
        const itemNames = receiptData.items.slice(0, 3).map(i => i.name);
        parts.push(itemNames.join('、'));
        if (receiptData.items.length > 3) parts.push('他');
    }
    mfRow.tekiyou = parts.join(' ') || '';

    // 税率から税区分を推定
    if (receiptData.taxRate === '10%') {
        mfRow.kariZeiku = '課対仕入10%';
    } else if (receiptData.taxRate === '8%') {
        mfRow.kariZeiku = '課対仕入8%';
    }

    // 仕訳パターンを適用（科目・税区分の自動入力）
    const matched = applyJournalPatterns(mfRow, journalPatterns || []);
    if (matched) {
        mfRow._matchedPattern = matched.id;
        mfRow._matchedKeyword = matched.keyword;
    }

    // 業種デフォルト勘定科目（仕訳パターンで埋まらなかった場合のフォールバック）
    // 業種未設定でも「共通」項目は適用される
    if (!mfRow.kariKamoku && mfRow.tekiyou) {
        const defaultKamoku = matchDefaultAccount(mfRow.tekiyou, industry || '');
        if (defaultKamoku) {
            mfRow.kariKamoku = defaultKamoku;
            mfRow._defaultAccount = true;
        }
    }

    // デフォルト貸方勘定科目（仕訳パターンで埋まらなかった場合）
    if (!mfRow.kashiKamoku && defaultKashi) {
        mfRow.kashiKamoku = defaultKashi;
    }

    // 訂正ルールを適用
    applyCorrectionRules(mfRow, correctionRules || []);

    // OCR領収書であることを示すフラグ
    mfRow._isReceipt = true;
    mfRow._receiptConfidence = receiptData.confidence;

    return mfRow;
}

// --- Gemini汎用フォーマットからMF仕訳行を生成 ---
// ocrWithGemini()の結果（entries配列）を受け取り、MF行の配列を返す
function buildMFRowsFromGeminiResult(geminiResult, journalPatterns, correctionRules, industry, defaultKashi) {
    const docType = geminiResult.documentType || 'other';
    const entries = geminiResult.entries || [];
    const rows = [];

    // 書類種別の日本語ラベル
    const docTypeLabels = {
        receipt: '領収書', invoice: '請求書', credit_card: 'クレカ明細',
        bank_statement: '通帳', petty_cash: '小口現金出納帳',
        expense_report: '経費精算書', sales_data: '売上データ', other: 'その他'
    };

    for (const entry of entries) {
        const mfRow = {};

        // 取引日
        mfRow.torihikiDate = entry.date || '';

        // 金額
        const amount = entry.amount || 0;
        if (entry.isIncome) {
            // 入金：借方=現金等、貸方=売上等
            mfRow.kashiKingaku = amount;
            mfRow.kariKingaku = amount;
        } else {
            // 支払：借方=経費等、貸方=現金等
            mfRow.kariKingaku = amount;
            mfRow.kashiKingaku = amount;
        }

        // 取引先
        mfRow.torihikisaki = entry.counterparty || '';

        // 摘要を組み立て
        const parts = [];
        if (entry.counterparty) parts.push(entry.counterparty);
        if (entry.description) parts.push(entry.description);
        if (parts.length === 0 && entry.items && entry.items.length > 0) {
            parts.push(entry.items.slice(0, 3).map(i => i.name).join('、'));
        }
        mfRow.tekiyou = parts.join(' ') || '';

        // 税率から税区分を推定
        if (entry.taxRate === '10%') {
            mfRow.kariZeiku = entry.isIncome ? '課税売上10%' : '課対仕入10%';
        } else if (entry.taxRate === '8%') {
            mfRow.kariZeiku = entry.isIncome ? '課税売上8%' : '課対仕入8%';
        }

        // ① 固定科目（通帳・小口現金 = 口座が確定している書類）
        //    これらは仕訳パターンで上書きしない
        if (!entry.isIncome) {
            if (docType === 'bank_statement') mfRow.kashiKamoku = '普通預金';
            if (docType === 'petty_cash') mfRow.kashiKamoku = '小口現金';
        } else {
            if (docType === 'bank_statement') mfRow.kariKamoku = '普通預金';
            if (docType === 'petty_cash') mfRow.kariKamoku = '小口現金';
        }

        // ② 仕訳パターンを適用（空フィールドを埋める + 補助科目含む）
        const matched = applyJournalPatterns(mfRow, journalPatterns || []);
        if (matched) {
            mfRow._matchedPattern = matched.id;
            mfRow._matchedKeyword = matched.keyword;
        }

        // ③ 業種デフォルト勘定科目（借方 or 貸方の空き側）
        const targetField = entry.isIncome ? 'kashiKamoku' : 'kariKamoku';
        if (!mfRow[targetField] && mfRow.tekiyou) {
            const defaultKamoku = matchDefaultAccount(mfRow.tekiyou, industry || '');
            if (defaultKamoku) {
                mfRow[targetField] = defaultKamoku;
                mfRow._defaultAccount = true;
            }
        }

        // ④ 書類種別フォールバック（パターンで埋まらなかった場合の貸方デフォルト）
        if (!mfRow.kashiKamoku) {
            if (docType === 'invoice') mfRow.kashiKamoku = '買掛金';
            else if (docType === 'credit_card') mfRow.kashiKamoku = '未払金';
            else if (docType === 'expense_report') mfRow.kashiKamoku = '未払金';
        }

        // ⑤ ユーザー設定のデフォルト貸方（それでも埋まらなかった場合）
        if (!mfRow.kashiKamoku && defaultKashi) {
            mfRow.kashiKamoku = defaultKashi;
        }

        // 訂正ルールを適用
        applyCorrectionRules(mfRow, correctionRules || []);

        // メタ情報
        mfRow._isReceipt = true;
        mfRow._documentType = docType;
        mfRow._documentTypeLabel = docTypeLabels[docType] || docType;
        mfRow._receiptConfidence = geminiResult.confidence;

        rows.push(mfRow);
    }

    return rows;
}
