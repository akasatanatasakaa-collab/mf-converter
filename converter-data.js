/* ===== MF仕訳帳コンバーター データ管理 ===== */

// --- 共通ユーティリティ（旧 app.js から移植）---

// ユニークID生成
function generateId(prefix) {
    return prefix + '_' + Date.now() + '_' + Math.random().toString(36).substr(2, 5);
}

// トースト通知
function showToast(message) {
    // 既存のトーストを削除
    const existing = document.querySelector('.toast');
    if (existing) existing.remove();

    const toast = document.createElement('div');
    toast.className = 'toast';
    toast.textContent = message;
    document.body.appendChild(toast);

    // 表示アニメーション
    requestAnimationFrame(() => {
        toast.classList.add('show');
    });

    // 3秒後に非表示
    setTimeout(() => {
        toast.classList.remove('show');
        setTimeout(() => toast.remove(), 300);
    }, 3000);
}

// --- 業種別デフォルト勘定科目 [業種, 項目キーワード, 勘定科目] ---
const ACCOUNT_DEFAULTS = [
    ["共通", "ガソリン代", "車両費"],
    ["共通", "ＭＦ等のクラウドサービス", "通信費"],
    ["共通", "車検", "車両費"],
    ["共通", "洗車代", "車両費"],
    ["共通", "お菓子", "会議費"],
    ["共通", "固定資産税", "租税公課"],
    ["共通", "自動車税", "租税公課"],
    ["共通", "収入印紙", "租税公課"],
    ["共通", "仕入", "仕入高"],
    ["共通", "段ボール", "荷造運賃"],
    ["共通", "宅配便", "荷造運賃"],
    ["共通", "給与", "給料賃金"],
    ["共通", "賞与", "給料賃金"],
    ["共通", "健康保険", "法定福利費"],
    ["共通", "厚生年金", "法定福利費"],
    ["共通", "雇用保険", "法定福利費"],
    ["共通", "社員旅行", "福利厚生費"],
    ["共通", "慶弔", "福利厚生費"],
    ["共通", "忘年会", "福利厚生費"],
    ["共通", "業務委託", "業務委託費"],
    ["共通", "会議室", "会議費"],
    ["共通", "お茶", "会議費"],
    ["共通", "会食", "接待交際費"],
    ["共通", "お中元", "接待交際費"],
    ["共通", "お歳暮", "接待交際費"],
    ["共通", "セミナー", "研修採用費"],
    ["共通", "求人", "研修採用費"],
    ["共通", "広告", "広告宣伝費"],
    ["共通", "パンフレット", "広告宣伝費"],
    ["共通", "税理士", "支払報酬"],
    ["共通", "弁護士", "支払報酬"],
    ["共通", "顧問料", "支払報酬"],
    ["共通", "火災保険", "保険料"],
    ["共通", "自動車保険", "保険料"],
    ["共通", "家賃", "地代家賃"],
    ["共通", "駐車場", "地代家賃"],
    ["共通", "電気代", "水道光熱費"],
    ["共通", "電気料金", "水道光熱費"],
    ["共通", "水道代", "水道光熱費"],
    ["共通", "ガス料金", "水道光熱費"],
    ["共通", "ガス代", "水道光熱費"],
    ["共通", "リース", "リース料"],
    ["共通", "電球", "消耗品費"],
    ["共通", "洗剤", "消耗品費"],
    ["共通", "ティッシュ", "消耗品費"],
    ["共通", "文房具", "消耗品費"],
    ["共通", "新幹線", "旅費交通費"],
    ["共通", "航空券", "旅費交通費"],
    ["共通", "電車", "旅費交通費"],
    ["共通", "ホテル", "旅費交通費"],
    ["共通", "宿泊", "旅費交通費"],
    ["共通", "電話料金", "通信費"],
    ["共通", "携帯電話", "通信費"],
    ["共通", "インターネット", "通信費"],
    ["共通", "プロバイダ", "通信費"],
    ["共通", "切手", "通信費"],
    ["共通", "レターパック", "通信費"],
    ["共通", "NTT", "通信費"],
    ["共通", "USEN", "通信費"],
    ["共通", "振込手数料", "支払手数料"],
    ["共通", "修理", "修繕費"],
    ["共通", "商工会議所", "諸会費"],
    ["共通", "年会費", "諸会費"],
    ["共通", "新聞", "新聞図書費"],
    ["共通", "書籍", "新聞図書費"],
    ["共通", "ゴミ処理", "雑費"],
    ["共通", "ゴミ回収", "雑費"],
    ["共通", "清掃", "雑費"],
    ["共通", "ガソリン", "車両費"],
    ["共通", "高速道路", "車両費"],
    ["共通", "ETC", "車両費"],
    ["共通", "印紙税", "租税公課"],
    ["共通", "事業所税", "租税公課"],
    ["法人共通", "役員報酬", "役員報酬"],
    ["法人共通", "法人税", "未払法人税等"],
    ["個人事業主共通", "事業主", "事業主勘定"],
    ["飲食業", "肉", "仕入"],
    ["飲食業", "魚", "仕入"],
    ["飲食業", "野菜", "仕入"],
    ["飲食業", "食材", "仕入"],
    ["飲食業", "酒", "仕入"],
    ["飲食業", "ビール", "仕入"],
    ["飲食業", "ワイン", "仕入"],
    ["飲食業", "割り箸", "仕入"],
    ["飲食業", "弁当容器", "仕入"],
    ["医療業", "医薬品", "医薬品費"],
    ["医療業", "注射器", "診療材料費"],
    ["医療業", "ガーゼ", "診療材料費"],
    ["医療業", "包帯", "診療材料費"],
    ["医療業", "MRI", "医療機器"],
    ["医療業", "レントゲン", "医療機器"],
    ["医療業", "内視鏡", "医療機器"],
    ["医療業", "診療報酬", "保険診療収入"],
    ["医療業", "人間ドック", "自由診療収入"],
    ["医療業", "予防接種", "自由診療収入"],
    ["医療業", "医師会", "租税公課"],
    ["歯医者", "技工", "外注技工費"],
    ["建設業", "木材", "材料費"],
    ["建設業", "鉄骨", "材料費"],
    ["建設業", "セメント", "材料費"],
    ["建設業", "外注", "外注費"],
    ["建設業", "重機", "減価償却費"],
    ["IT・ソフトウェア", "AWS", "通信費"],
    ["IT・ソフトウェア", "GCP", "通信費"],
    ["IT・ソフトウェア", "SaaS", "通信費"],
    ["IT・ソフトウェア", "サーバー", "通信費"],
    ["製造業", "原材料", "原材料"],
    ["製造業", "鉄板", "原材料"],
    ["製造業", "プラスチック", "原材料"],
    ["美容・理容業", "シャンプー", "材料費"],
    ["美容・理容業", "カラー剤", "材料費"],
    ["美容・理容業", "パーマ液", "材料費"],
    ["不動産賃貸業", "クロス張替", "修繕費"],
    ["不動産賃貸業", "ハウスクリーニング", "修繕費"],
    ["農業", "バナナ", "仕入"],
];

// 選択可能な業種リスト
const INDUSTRY_LIST = [
    "IT・ソフトウェア", "コンサルティング", "不動産仲介業", "不動産売買業",
    "不動産管理業", "不動産賃貸業", "不動産開発業", "人材派遣・紹介",
    "医療業", "学習塾", "宿泊業", "広告代理店", "建設業", "歯医者",
    "美容・理容業", "製造業", "農業", "飲食業"
];

// --- 業種に応じたデフォルト勘定科目を検索 ---
function matchDefaultAccount(text, industry) {
    if (!text) return null;
    const t = String(text).trim();
    if (!t) return null;

    // 共通系は業種未設定でも常に適用
    const commonTypes = ['共通', '法人共通', '個人事業主共通'];
    let bestMatch = null;
    let bestLen = 0;

    for (const entry of ACCOUNT_DEFAULTS) {
        const [ind, keyword, kamoku] = entry;
        // 業種フィルタ: 共通系は常にマッチ、それ以外は選択業種のみ
        if (!commonTypes.includes(ind) && ind !== industry) continue;

        if (t.includes(keyword) && keyword.length > bestLen) {
            bestMatch = kamoku;
            bestLen = keyword.length;
        }
    }

    return bestMatch;
}

// --- MF仕訳帳カラム定義 ---
const MF_COLUMNS = [
    { key: 'torihikiNo',    label: '取引No' },
    { key: 'torihikiDate',  label: '取引日' },
    { key: 'kariKamoku',    label: '借方勘定科目' },
    { key: 'kariHojo',      label: '借方補助科目' },
    { key: 'kariBumon',     label: '借方部門' },
    { key: 'kariZeiku',     label: '借方税区分' },
    { key: 'kariKingaku',   label: '借方金額' },
    { key: 'kashiKamoku',   label: '貸方勘定科目' },
    { key: 'kashiHojo',     label: '貸方補助科目' },
    { key: 'kashiBumon',    label: '貸方部門' },
    { key: 'kashiZeiku',    label: '貸方税区分' },
    { key: 'kashiKingaku',  label: '貸方金額' },
    { key: 'tekiyou',       label: '摘要' },
    { key: 'memo',          label: '仕訳メモ' },
    { key: 'tag',           label: 'タグ' },
    { key: 'torihikisaki',  label: '取引先' },
];

// マッピング可能なカラム（取引Noは自動採番なので除外）
const MAPPABLE_COLUMNS = MF_COLUMNS.filter(c => c.key !== 'torihikiNo');

// --- カラム名→MFフィールドの自動推定ヒント ---
const MAPPING_HINTS = {
    'torihikiDate':  ['日付', '取引日', 'date', '年月日', '発生日', '支払日', '入金日', '利用日', '計上日', '起票日', '実行日', '振替日', '処理日', '決済日'],
    'kariKamoku':    ['借方勘定科目', '借方科目', '勘定科目', '科目', '費目', '経費科目', '勘定', '仕訳科目'],
    'kariHojo':      ['借方補助科目', '補助科目', '補助'],
    'kariBumon':     ['借方部門', '部門', 'department'],
    'kariZeiku':     ['借方税区分', '税区分', '消費税区分', '税率'],
    'kariKingaku':   ['借方金額', '金額', 'amount', '支出', '出金', '支払金額', '利用金額', '引落金額', '引落額', '出金額', '支出額', '請求金額', '請求額'],
    'kashiKamoku':   ['貸方勘定科目', '貸方科目', '相手科目', '入金科目', '相手勘定'],
    'kashiHojo':     ['貸方補助科目'],
    'kashiBumon':    ['貸方部門'],
    'kashiZeiku':    ['貸方税区分'],
    'kashiKingaku':  ['貸方金額', '入金', '入金金額', '収入', '入金額', '収入額', '売上金額', '受取金額'],
    'tekiyou':       ['摘要', '内容', '取引内容', '品名', '備考', '明細', '利用先', '支払先名', '概要', '説明', '用途', '品目', '項目', 'description', 'memo'],
    'memo':          ['仕訳メモ', 'メモ', 'note', 'notes', '注記', '注釈', 'コメント'],
    'tag':           ['タグ', 'tag', 'tags', 'ラベル', 'label'],
    'torihikisaki':  ['取引先', '相手先', '支払先', '仕入先', '店名', '顧客名', '得意先', '請求先', '会社名', 'vendor', 'customer'],
};

// --- よく使う勘定科目 ---
const COMMON_ACCOUNTS = [
    '現金', '普通預金', '当座預金', '定期預金',
    '売掛金', '受取手形', '買掛金', '支払手形', '未払金', '前受金', '預り金',
    '売上高', '仕入高',
    '給料手当', '法定福利費', '福利厚生費',
    '旅費交通費', '通信費', '消耗品費', '事務用品費',
    '水道光熱費', '支払手数料', '保険料', '租税公課',
    '減価償却費', '接待交際費', '広告宣伝費',
    '地代家賃', '修繕費', '荷造運賃', '車両費',
    '外注費', '会議費', '新聞図書費', '諸会費', '研修費',
    '支払利息', '受取利息', '雑費', '雑収入', '雑損失',
    '事業主貸', '事業主借',
];

// --- 勘定科目の別名マッピング（プリセット） ---
const ACCOUNT_ALIASES = {
    // 旅費交通費
    '交通費': '旅費交通費',
    'タクシー代': '旅費交通費',
    '電車代': '旅費交通費',
    'バス代': '旅費交通費',
    '高速代': '旅費交通費',
    'ETC': '旅費交通費',
    '駐車代': '旅費交通費',
    '出張費': '旅費交通費',
    '宿泊費': '旅費交通費',
    // 通信費
    '電話代': '通信費',
    '携帯代': '通信費',
    'インターネット代': '通信費',
    '切手代': '通信費',
    '郵便代': '通信費',
    'プロバイダ料': '通信費',
    // 地代家賃
    '家賃': '地代家賃',
    '駐車場代': '地代家賃',
    '事務所家賃': '地代家賃',
    '月極駐車場': '地代家賃',
    // 事務用品費
    '文房具': '事務用品費',
    'コピー用紙': '事務用品費',
    '事務用品': '事務用品費',
    // 車両費
    'ガソリン代': '車両費',
    '車検代': '車両費',
    '車両整備': '車両費',
    // 接待交際費
    '飲食代': '接待交際費',
    '接待費': '接待交際費',
    '贈答品': '接待交際費',
    'お歳暮': '接待交際費',
    'お中元': '接待交際費',
    // 支払手数料
    '振込手数料': '支払手数料',
    '手数料': '支払手数料',
    '銀行手数料': '支払手数料',
    'ATM手数料': '支払手数料',
    // 水道光熱費
    '電気代': '水道光熱費',
    'ガス代': '水道光熱費',
    '水道代': '水道光熱費',
    '電力': '水道光熱費',
    // 租税公課
    '印紙代': '租税公課',
    '収入印紙': '租税公課',
    '固定資産税': '租税公課',
    '自動車税': '租税公課',
    '住民税': '租税公課',
    '事業税': '租税公課',
    '消費税': '租税公課',
    // 会議費
    '会議室代': '会議費',
    '打ち合わせ': '会議費',
    // 新聞図書費
    '書籍': '新聞図書費',
    '新聞': '新聞図書費',
    '雑誌': '新聞図書費',
    'サブスク': '新聞図書費',
    // 広告宣伝費
    '広告費': '広告宣伝費',
    'WEB広告': '広告宣伝費',
    'チラシ': '広告宣伝費',
};

// --- 税区分マスタ ---
const TAX_CATEGORIES = [
    '対象外',
    '非課税',
    '不課税',
    '課税売上10%',
    '課税売上8%（軽減）',
    '課税仕入10%',
    '課税仕入8%（軽減）',
    '免税売上',
    '共通対応仕入',
    '非課税売上対応仕入',
    '課税売上対応仕入',
];

// --- 日付形式の選択肢 ---
const DATE_FORMATS = [
    { id: 'auto',           label: '自動検出' },
    { id: 'yyyy/MM/dd',     label: 'yyyy/MM/dd （例: 2024/01/15）' },
    { id: 'yyyy-MM-dd',     label: 'yyyy-MM-dd （例: 2024-01-15）' },
    { id: 'yyyy/M/d',       label: 'yyyy/M/d （例: 2024/1/5）' },
    { id: 'MM/dd/yyyy',     label: 'MM/dd/yyyy （例: 01/15/2024）' },
    { id: 'yyyyMMdd',       label: 'yyyyMMdd （例: 20240115）' },
    { id: 'yyyy年M月d日',    label: 'yyyy年M月d日 （例: 2024年1月15日）' },
    { id: 'wareki',         label: '令和x年M月d日 （例: 令和6年1月15日）' },
    { id: 'M/d',            label: 'M/d ※年なし （例: 2/1 → 今年の日付）' },
    { id: 'M-d',            label: 'M-d ※年なし （例: 2-1 → 今年の日付）' },
    { id: 'M月d日',          label: 'M月d日 ※年なし （例: 2月1日 → 今年の日付）' },
];

// --- マッピングプリセット ---
// ヘッダー名のキーワードで自動マッチする。matchはヘッダー名に含まれるキーワード配列
const MAPPING_PRESETS = [
    {
        id: 'bank',
        label: '銀行通帳',
        desc: '日付・摘要・出金・入金',
        mapping: [
            { match: ['日付', '年月日', 'date'], field: 'torihikiDate' },
            { match: ['摘要', '内容', '取引内容', '明細'], field: 'tekiyou' },
            { match: ['出金', '支出', '引落', '支払'], field: 'kariKingaku' },
            { match: ['入金', '収入', '預入'], field: 'kashiKingaku' },
            { match: ['残高'], field: null },
        ],
        fixedValues: { kashiKamoku: '普通預金' },
    },
    {
        id: 'creditcard',
        label: 'クレカ明細',
        desc: '利用日・利用先・金額',
        mapping: [
            { match: ['日付', '利用日', 'date'], field: 'torihikiDate' },
            { match: ['利用先', '支払先', '店名', '加盟店'], field: 'tekiyou' },
            { match: ['金額', '利用金額', '支払金額', 'amount'], field: 'kariKingaku' },
        ],
        fixedValues: { kashiKamoku: '未払金' },
    },
    {
        id: 'expense',
        label: '経費精算',
        desc: '日付・科目・金額・摘要',
        mapping: [
            { match: ['日付', 'date'], field: 'torihikiDate' },
            { match: ['科目', '勘定科目', '費目', '経費科目'], field: 'kariKamoku' },
            { match: ['金額', 'amount'], field: 'kariKingaku' },
            { match: ['摘要', '内容', '備考', '明細'], field: 'tekiyou' },
            { match: ['取引先', '支払先', '店名'], field: 'torihikisaki' },
        ],
        fixedValues: { kashiKamoku: '現金' },
    },
    {
        id: 'sales',
        label: '売上データ',
        desc: '日付・取引先・金額',
        mapping: [
            { match: ['日付', 'date'], field: 'torihikiDate' },
            { match: ['取引先', '顧客', '得意先', '請求先'], field: 'torihikisaki' },
            { match: ['金額', '売上', '請求金額', 'amount'], field: 'kashiKingaku' },
            { match: ['摘要', '品名', '内容', '明細'], field: 'tekiyou' },
        ],
        fixedValues: { kariKamoku: '売掛金', kashiKamoku: '売上高' },
    },
];

// --- テンプレート管理（localStorage） ---
const TEMPLATE_STORAGE_KEY = 'kichou_converter_templates';

function getConverterTemplates() {
    return JSON.parse(localStorage.getItem(TEMPLATE_STORAGE_KEY) || '[]');
}

function saveConverterTemplate(template) {
    const templates = getConverterTemplates();
    template.id = generateId('tpl');
    template.createdAt = new Date().toISOString();
    templates.push(template);
    localStorage.setItem(TEMPLATE_STORAGE_KEY, JSON.stringify(templates));
    return template;
}

function updateConverterTemplate(id, data) {
    const templates = getConverterTemplates();
    const idx = templates.findIndex(t => t.id === id);
    if (idx !== -1) {
        Object.assign(templates[idx], data);
        templates[idx].updatedAt = new Date().toISOString();
        localStorage.setItem(TEMPLATE_STORAGE_KEY, JSON.stringify(templates));
    }
}

function deleteConverterTemplate(id) {
    const templates = getConverterTemplates().filter(t => t.id !== id);
    localStorage.setItem(TEMPLATE_STORAGE_KEY, JSON.stringify(templates));
}

function getConverterTemplate(id) {
    return getConverterTemplates().find(t => t.id === id) || null;
}

// --- 会社プロファイル管理（localStorage） ---
// 訂正ルール・テンプレートを会社ごとに管理するための仕組み
const COMPANIES_KEY = 'kichou_converter_companies';
const CORRECTION_RULES_KEY = 'kichou_correction_rules'; // { 会社名: [ルール配列] }

function getCompanies() {
    return JSON.parse(localStorage.getItem(COMPANIES_KEY) || '[]');
}

function addCompany(name) {
    const companies = getCompanies();
    if (!companies.includes(name)) {
        companies.push(name);
        companies.sort();
        localStorage.setItem(COMPANIES_KEY, JSON.stringify(companies));
    }
}

// --- 会社の業種設定 ---
const COMPANY_INDUSTRY_KEY = 'kichou_company_industry'; // { 会社名: 業種名 }

function getCompanyIndustry(company) {
    const data = JSON.parse(localStorage.getItem(COMPANY_INDUSTRY_KEY) || '{}');
    return data[company] || '';
}

function setCompanyIndustry(company, industry) {
    const data = JSON.parse(localStorage.getItem(COMPANY_INDUSTRY_KEY) || '{}');
    data[company] = industry;
    localStorage.setItem(COMPANY_INDUSTRY_KEY, JSON.stringify(data));
}

function deleteCompany(name) {
    // 会社を削除し、関連する訂正ルール・仕訳パターンも削除
    const companies = getCompanies().filter(c => c !== name);
    localStorage.setItem(COMPANIES_KEY, JSON.stringify(companies));

    const allRules = getAllCorrectionRules();
    delete allRules[name];
    localStorage.setItem(CORRECTION_RULES_KEY, JSON.stringify(allRules));

    const allPatterns = getAllJournalPatterns();
    delete allPatterns[name];
    localStorage.setItem(JOURNAL_PATTERNS_KEY, JSON.stringify(allPatterns));

    const industryData = JSON.parse(localStorage.getItem(COMPANY_INDUSTRY_KEY) || '{}');
    delete industryData[name];
    localStorage.setItem(COMPANY_INDUSTRY_KEY, JSON.stringify(industryData));
}

// --- 訂正ルール管理（会社別・localStorage） ---
// 構造: { "A社": [{id, field, from, to, createdAt}], "B社": [...] }

function getAllCorrectionRules() {
    return JSON.parse(localStorage.getItem(CORRECTION_RULES_KEY) || '{}');
}

function getCorrectionRules(company) {
    if (!company) return [];
    const all = getAllCorrectionRules();
    return all[company] || [];
}

function addCorrectionRule(company, rule) {
    if (!company) return rule;
    const all = getAllCorrectionRules();
    if (!all[company]) all[company] = [];

    // 同じfield+fromの既存ルールがあれば上書き
    const existIdx = all[company].findIndex(r => r.field === rule.field && r.from === rule.from);
    if (existIdx !== -1) {
        all[company][existIdx].to = rule.to;
        all[company][existIdx].updatedAt = new Date().toISOString();
    } else {
        rule.id = generateId('crule');
        rule.createdAt = new Date().toISOString();
        all[company].push(rule);
    }

    localStorage.setItem(CORRECTION_RULES_KEY, JSON.stringify(all));
    // 会社名も登録
    addCompany(company);
    return rule;
}

function deleteCorrectionRule(company, id) {
    if (!company) return;
    const all = getAllCorrectionRules();
    if (!all[company]) return;
    all[company] = all[company].filter(r => r.id !== id);
    localStorage.setItem(CORRECTION_RULES_KEY, JSON.stringify(all));
}

function clearCorrectionRules(company) {
    if (!company) return;
    const all = getAllCorrectionRules();
    all[company] = [];
    localStorage.setItem(CORRECTION_RULES_KEY, JSON.stringify(all));
}

// --- 仕訳パターン管理（会社別・localStorage） ---
// 過去のMF仕訳帳CSVからパターンを抽出し、新データの変換時に自動適用する
// 構造: { "会社名": [{id, keyword, kariKamoku, kashiKamoku, kariZeiku, kashiZeiku, kariHojo, kashiHojo, count, createdAt}] }
const JOURNAL_PATTERNS_KEY = 'kichou_journal_patterns';

function getAllJournalPatterns() {
    return JSON.parse(localStorage.getItem(JOURNAL_PATTERNS_KEY) || '{}');
}

function getJournalPatterns(company) {
    if (!company) return [];
    const all = getAllJournalPatterns();
    return all[company] || [];
}

function addJournalPattern(company, pattern) {
    if (!company) return pattern;
    const all = getAllJournalPatterns();
    if (!all[company]) all[company] = [];

    // 同じキーワードの既存パターンがあれば上書き
    const existIdx = all[company].findIndex(p => p.keyword === pattern.keyword);
    if (existIdx !== -1) {
        const existing = all[company][existIdx];
        existing.kariKamoku = pattern.kariKamoku || existing.kariKamoku;
        existing.kashiKamoku = pattern.kashiKamoku || existing.kashiKamoku;
        existing.kariZeiku = pattern.kariZeiku || existing.kariZeiku;
        existing.kashiZeiku = pattern.kashiZeiku || existing.kashiZeiku;
        existing.kariHojo = pattern.kariHojo || existing.kariHojo;
        existing.kashiHojo = pattern.kashiHojo || existing.kashiHojo;
        existing.count = (existing.count || 1) + (pattern.count || 1);
        existing.updatedAt = new Date().toISOString();
    } else {
        pattern.id = generateId('jp');
        pattern.count = pattern.count || 1;
        pattern.createdAt = new Date().toISOString();
        all[company].push(pattern);
    }

    localStorage.setItem(JOURNAL_PATTERNS_KEY, JSON.stringify(all));
    addCompany(company);
    return pattern;
}

function deleteJournalPattern(company, id) {
    if (!company) return;
    const all = getAllJournalPatterns();
    if (!all[company]) return;
    all[company] = all[company].filter(p => p.id !== id);
    localStorage.setItem(JOURNAL_PATTERNS_KEY, JSON.stringify(all));
}

function clearJournalPatterns(company) {
    if (!company) return;
    const all = getAllJournalPatterns();
    all[company] = [];
    localStorage.setItem(JOURNAL_PATTERNS_KEY, JSON.stringify(all));
}

// --- MF仕訳帳CSVからパターンを抽出 ---
// 摘要全体をキーワードにする（同じ店でも費目が違えば別パターン）
// 例: 「仕入れ費（リウボウ」と「消耗品費（リウボウ」は別キーワード
function extractKeywordFromTekiyou(tekiyou) {
    if (!tekiyou) return '';
    const trimmed = String(tekiyou).trim();
    if (!trimmed || trimmed.length < 2) return '';
    return trimmed;
}

// --- PDFをページごとに画像（Canvas）に変換 ---
async function pdfToImages(arrayBuffer, scale) {
    if (typeof pdfjsLib === 'undefined') {
        throw new Error('PDF.jsライブラリが読み込まれていません');
    }
    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
    const images = [];
    const renderScale = scale || 2; // 高解像度でレンダリング（OCR精度向上）

    for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
        const page = await pdf.getPage(pageNum);
        const viewport = page.getViewport({ scale: renderScale });
        const canvas = document.createElement('canvas');
        canvas.width = viewport.width;
        canvas.height = viewport.height;
        const ctx = canvas.getContext('2d');

        await page.render({ canvasContext: ctx, viewport: viewport }).promise;
        images.push(canvas);
    }

    return images;
}

// --- OCRでスキャンPDFからテキストを抽出 ---
async function ocrPDFPages(arrayBuffer, onProgress) {
    const images = await pdfToImages(arrayBuffer);
    const totalPages = images.length;

    if (totalPages === 0) {
        throw new Error('PDFページを画像に変換できませんでした');
    }

    // Tesseract.jsワーカーを作成（日本語+英語）
    const worker = await Tesseract.createWorker('jpn+eng', 1, {
        logger: (m) => {
            if (m.status === 'recognizing text' && onProgress) {
                // 各ページ内の進捗をトータル進捗に換算
                const pageProgress = m.progress || 0;
                const currentPage = parseInt(m.userJobId) || 0;
                const total = (currentPage + pageProgress) / totalPages;
                onProgress({
                    status: `ページ ${currentPage + 1} / ${totalPages} を読み取り中...`,
                    progress: total
                });
            }
        }
    });

    let allText = '';

    for (let i = 0; i < totalPages; i++) {
        if (onProgress) {
            onProgress({
                status: `ページ ${i + 1} / ${totalPages} を読み取り中...`,
                progress: i / totalPages
            });
        }

        const result = await worker.recognize(images[i], {}, { userJobId: String(i) });
        allText += result.data.text + '\n';
    }

    await worker.terminate();
    return allText;
}

// --- PDFからテキストを抽出（テキストレイヤー優先） ---
async function extractTextFromPDF(arrayBuffer) {
    if (typeof pdfjsLib === 'undefined') {
        throw new Error('PDF.jsライブラリが読み込まれていません');
    }

    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
    console.log('[PDF] ページ数:', pdf.numPages);
    const allRows = [];

    for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
        const page = await pdf.getPage(pageNum);
        const textContent = await page.getTextContent();
        console.log(`[PDF] ページ${pageNum}: テキストアイテム数=${textContent.items.length}`);

        // テキストアイテムをY座標でグループ化（同じ行のテキストをまとめる）
        const lineMap = {};
        textContent.items.forEach(item => {
            // 空テキストはスキップ
            if (!item.str || !item.str.trim()) return;
            // Y座標を丸めて同一行を判定（PDF座標は下が0）
            const y = Math.round(item.transform[5]);
            if (!lineMap[y]) lineMap[y] = [];
            lineMap[y].push({
                x: item.transform[4],
                text: item.str
            });
        });

        // Y座標を降順ソート（上から下へ）
        const sortedYs = Object.keys(lineMap).map(Number).sort((a, b) => b - a);

        sortedYs.forEach(y => {
            // 同一行内はX座標でソート（左から右へ）
            const items = lineMap[y].sort((a, b) => a.x - b.x);
            allRows.push(items);
        });
    }

    return allRows;
}

// --- PDFのテキスト行をCSVテキストに変換 ---
function pdfRowsToCSV(pdfRows) {
    if (pdfRows.length === 0) return '';

    // MF仕訳帳のヘッダーキーワードでヘッダー行を検出
    const mfHeaders = MF_COLUMNS.map(c => c.label);
    let headerRowIdx = -1;

    for (let i = 0; i < Math.min(pdfRows.length, 15); i++) {
        const lineText = pdfRows[i].map(item => item.text).join(' ');
        const matchCount = mfHeaders.filter(h => lineText.includes(h)).length;
        // 3つ以上のMFヘッダーが含まれていればヘッダー行
        if (matchCount >= 3) {
            headerRowIdx = i;
            break;
        }
    }

    if (headerRowIdx === -1) {
        // ヘッダーが見つからない場合はテキスト全体をタブ区切りで返す
        return pdfRows.map(row =>
            row.map(item => item.text).join('\t')
        ).join('\n');
    }

    // ヘッダー行のX座標からカラム境界を推定
    const headerItems = pdfRows[headerRowIdx];
    const columns = headerItems.map(item => ({
        label: item.text.trim(),
        x: item.x
    })).filter(col => col.label);

    // 各データ行をカラムに振り分け
    const csvLines = [];
    // ヘッダー行
    csvLines.push(columns.map(col => col.label).join('\t'));

    // データ行（ヘッダーの次の行から）
    for (let i = headerRowIdx + 1; i < pdfRows.length; i++) {
        const row = pdfRows[i];
        const cells = new Array(columns.length).fill('');

        row.forEach(item => {
            // 最も近いカラムを探す
            let bestCol = 0;
            let bestDist = Infinity;
            columns.forEach((col, colIdx) => {
                const dist = Math.abs(item.x - col.x);
                if (dist < bestDist) {
                    bestDist = dist;
                    bestCol = colIdx;
                }
            });
            // 同じセルに既にテキストがあればスペースで連結
            if (cells[bestCol]) {
                cells[bestCol] += ' ' + item.text.trim();
            } else {
                cells[bestCol] = item.text.trim();
            }
        });

        csvLines.push(cells.join('\t'));
    }

    return csvLines.join('\n');
}

// --- PDFの行データを汎用タブ区切りテキストに変換（Step1入力用） ---
function pdfRowsToTSV(pdfRows) {
    if (pdfRows.length === 0) return '';

    // カラム境界をX座標のクラスタリングで推定
    // 全行のX座標を収集
    const allXPositions = [];
    pdfRows.forEach(row => {
        row.forEach(item => {
            if (item.text.trim()) allXPositions.push(item.x);
        });
    });

    if (allXPositions.length === 0) return '';

    // X座標をソートしてクラスタリング（近い座標をグループ化）
    allXPositions.sort((a, b) => a - b);
    const clusters = [];
    const THRESHOLD = 15; // 15ポイント以内は同じカラムとみなす

    allXPositions.forEach(x => {
        const existing = clusters.find(c => Math.abs(c.center - x) < THRESHOLD);
        if (existing) {
            existing.positions.push(x);
            existing.center = existing.positions.reduce((a, b) => a + b, 0) / existing.positions.length;
        } else {
            clusters.push({ center: x, positions: [x] });
        }
    });

    // 出現頻度が低すぎるクラスタを除外（全行の10%未満）
    const minCount = Math.max(2, pdfRows.length * 0.1);
    const columnPositions = clusters
        .filter(c => c.positions.length >= minCount)
        .map(c => c.center)
        .sort((a, b) => a - b);

    if (columnPositions.length === 0) {
        // クラスタリング失敗時はスペース連結で返す
        return pdfRows.map(row =>
            row.map(item => item.text).join('\t')
        ).join('\n');
    }

    // 各行をカラムに振り分け
    const tsvLines = [];
    pdfRows.forEach(row => {
        const cells = new Array(columnPositions.length).fill('');

        row.forEach(item => {
            if (!item.text.trim()) return;
            // 最も近いカラム位置を探す
            let bestCol = 0;
            let bestDist = Infinity;
            columnPositions.forEach((pos, idx) => {
                const dist = Math.abs(item.x - pos);
                if (dist < bestDist) {
                    bestDist = dist;
                    bestCol = idx;
                }
            });
            if (cells[bestCol]) {
                cells[bestCol] += ' ' + item.text.trim();
            } else {
                cells[bestCol] = item.text.trim();
            }
        });

        tsvLines.push(cells.join('\t'));
    });

    return tsvLines.join('\n');
}

// --- PDFからパターンをインポート ---
async function importPatternsFromPDF(company, arrayBuffer) {
    try {
        const pdfRows = await extractTextFromPDF(arrayBuffer);
        if (pdfRows.length === 0) {
            return { count: 0, error: 'PDFからテキストを抽出できませんでした' };
        }

        const csvText = pdfRowsToCSV(pdfRows);
        if (!csvText) {
            return { count: 0, error: 'PDFの内容をパースできませんでした' };
        }

        return importPatternsFromMFCSV(company, csvText);
    } catch (e) {
        return { count: 0, error: 'PDF読み込みエラー: ' + e.message };
    }
}

function importPatternsFromMFCSV(company, csvText) {
    if (!company) return { count: 0, error: '会社を選択してください' };

    // MF仕訳帳CSVをパース
    const delimiter = detectDelimiter(csvText);
    const parsed = parseCSV(csvText, delimiter, true);

    if (parsed.rows.length === 0) {
        return { count: 0, error: 'データが見つかりませんでした' };
    }

    // MFカラムのインデックスを特定（ヘッダー名で検索）
    const colIndex = {};
    MF_COLUMNS.forEach(col => {
        const idx = parsed.headers.findIndex(h =>
            h.trim() === col.label || h.trim() === col.key
        );
        if (idx !== -1) colIndex[col.key] = idx;
    });

    // 摘要列が見つからない場合
    if (colIndex.tekiyou === undefined) {
        return { count: 0, error: '摘要列が見つかりませんでした' };
    }

    // 各行からパターンを抽出
    const patternMap = {}; // keyword → パターンデータ
    parsed.rows.forEach(row => {
        const tekiyou = (row[colIndex.tekiyou] || '').trim();
        const keyword = extractKeywordFromTekiyou(tekiyou);
        if (!keyword) return;

        const pattern = {
            keyword: keyword,
            kariKamoku: colIndex.kariKamoku !== undefined ? (row[colIndex.kariKamoku] || '').trim() : '',
            kashiKamoku: colIndex.kashiKamoku !== undefined ? (row[colIndex.kashiKamoku] || '').trim() : '',
            kariZeiku: colIndex.kariZeiku !== undefined ? (row[colIndex.kariZeiku] || '').trim() : '',
            kashiZeiku: colIndex.kashiZeiku !== undefined ? (row[colIndex.kashiZeiku] || '').trim() : '',
            kariHojo: colIndex.kariHojo !== undefined ? (row[colIndex.kariHojo] || '').trim() : '',
            kashiHojo: colIndex.kashiHojo !== undefined ? (row[colIndex.kashiHojo] || '').trim() : '',
        };

        // 科目が1つも入っていないパターンはスキップ
        if (!pattern.kariKamoku && !pattern.kashiKamoku) return;

        if (patternMap[keyword]) {
            patternMap[keyword].count++;
        } else {
            pattern.count = 1;
            patternMap[keyword] = pattern;
        }
    });

    // パターンを一括登録
    let addedCount = 0;
    for (const pattern of Object.values(patternMap)) {
        addJournalPattern(company, pattern);
        addedCount++;
    }

    return { count: addedCount, error: null };
}

// ===== Gemini API OCR =====

const GEMINI_API_KEY_KEY = 'mf_converter_gemini_api_key';

function getGeminiApiKey() {
    return localStorage.getItem(GEMINI_API_KEY_KEY) || '';
}

function setGeminiApiKey(key) {
    if (key) {
        localStorage.setItem(GEMINI_API_KEY_KEY, key);
    } else {
        localStorage.removeItem(GEMINI_API_KEY_KEY);
    }
}

// PDFの各ページをbase64画像に変換
async function pdfToBase64Images(arrayBuffer, scale) {
    const images = await pdfToImages(arrayBuffer, scale || 1.5);
    return images.map(canvas => {
        // data:image/png;base64,... からプレフィックスを除去
        const dataUrl = canvas.toDataURL('image/png');
        return dataUrl.replace(/^data:image\/png;base64,/, '');
    });
}

// ファイルをBase64文字列に変換（プレフィックスなし）
function fileToBase64(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => {
            const base64 = reader.result.split(',')[1];
            resolve(base64);
        };
        reader.onerror = reject;
        reader.readAsDataURL(file);
    });
}

// Gemini APIをストリーミングで呼び出し、受信バイト数でプログレスを更新
async function fetchGeminiStreaming(apiKey, requestBody, onProgress) {
    const response = await fetch(
        `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:streamGenerateContent?alt=sse&key=${apiKey}`,
        {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(requestBody)
        }
    );

    if (!response.ok) {
        const errorData = await response.json().catch(() => ({}));
        const msg = errorData?.error?.message || response.statusText;
        throw new Error(`Gemini API エラー (${response.status}): ${msg}`);
    }

    const reader = response.body.getReader();
    const decoder = new TextDecoder();
    let fullText = '';
    let receivedBytes = 0;

    while (true) {
        const { done, value } = await reader.read();
        if (done) break;

        receivedBytes += value.length;
        const chunk = decoder.decode(value, { stream: true });

        // SSEイベントからテキストを抽出
        for (const line of chunk.split('\n')) {
            if (!line.startsWith('data: ')) continue;
            try {
                const data = JSON.parse(line.slice(6));
                const text = data?.candidates?.[0]?.content?.parts?.[0]?.text || '';
                fullText += text;
            } catch (e) { /* 部分JSONは無視 */ }
        }

        if (onProgress) onProgress(receivedBytes, fullText.length, fullText);
    }

    console.log('[Gemini] ストリーミング応答:', fullText);
    return fullText;
}

// ストリーミングで受信したテキストからJSONを解析
function parseGeminiTextResponse(text) {
    let jsonStr = text.trim();
    const jsonMatch = jsonStr.match(/```(?:json)?\s*([\s\S]*?)```/);
    if (jsonMatch) jsonStr = jsonMatch[1].trim();

    try {
        return JSON.parse(jsonStr);
    } catch (e) {
        console.error('[Gemini] JSONパースエラー:', e, 'テキスト:', jsonStr);
        throw new Error('Gemini APIの応答をパースできませんでした');
    }
}

// Gemini APIレスポンスからJSONを解析する共通関数（非ストリーミング用）
function parseGeminiJsonResponse(responseData) {
    const text = responseData?.candidates?.[0]?.content?.parts?.[0]?.text || '';
    console.log('[Gemini] 応答:', text);

    let jsonStr = text.trim();
    const jsonMatch = jsonStr.match(/```(?:json)?\s*([\s\S]*?)```/);
    if (jsonMatch) jsonStr = jsonMatch[1].trim();

    try {
        return JSON.parse(jsonStr);
    } catch (e) {
        console.error('[Gemini] JSONパースエラー:', e, 'テキスト:', jsonStr);
        throw new Error('Gemini APIの応答をパースできませんでした');
    }
}

// Gemini OCR結果のentryを正規化する共通関数
function normalizeGeminiEntry(entry) {
    return {
        date: entry.date || '',
        counterparty: entry.counterparty || '',
        amount: Number(entry.amount) || 0,
        tax: Number(entry.tax) || 0,
        taxRate: entry.taxRate || '',
        description: entry.description || '',
        items: Array.isArray(entry.items) ? entry.items.map(i => ({
            name: i.name || '', amount: Number(i.amount) || 0, taxRate: i.taxRate || ''
        })) : [],
        isIncome: !!entry.isIncome
    };
}

// --- Gemini OCR プロンプト定義 ---
const GEMINI_DOC_TYPES = `対応する書類種別:
- receipt（領収書・レシート）
- invoice（請求書）
- credit_card（クレジットカード明細）
- bank_statement（銀行通帳・銀行の入出金明細。銀行名や口座番号が記載されている）
- petty_cash（小口現金出納帳・現金出納帳。「小口現金」「現金出納帳」等の文字がある、または銀行名がなく現金の入出金を記録した帳簿）
- expense_report（経費精算書）
- sales_data（売上データ・売上一覧）
- other（その他）`;

const GEMINI_ENTRY_FORMAT = `{
      "date": "YYYY/MM/DD",
      "counterparty": "取引先・店名",
      "amount": 金額（数値、税込合計）,
      "tax": 消費税額（数値、不明なら0）,
      "taxRate": "10%" または "8%"（全品目同じ税率の場合。混在なら""）,
      "description": "摘要・内容の要約",
      "items": [{"name": "品名", "amount": 金額数値, "taxRate": "10%"または"8%"}],
      "isIncome": false
    }`;

const GEMINI_RULES = `ルール:
- 金額はカンマなしの数値で返す
- 日付が和暦の場合は西暦に変換する（令和6年=2024年、令和7年=2025年、令和8年=2026年）
- 領収書・請求書は通常1件のentry。クレカ明細・通帳・小口現金出納帳・経費精算書・売上データは複数entriesを返す
- isIncomeは入金・売上の場合true、支払・経費の場合false
- 通帳の場合: 入金行はisIncome=true、出金行はisIncome=false
- 小口現金出納帳の判定: 書類に「小口現金」「現金出納帳」等の文字があるか、銀行名・口座番号がなく現金の収支を記録した帳簿はpetty_cash。入金欄はisIncome=true、出金・支出欄はisIncome=false。bank_statementにしないこと
- bank_statementとpetty_cashの区別: 銀行名・支店名・口座番号が記載されていればbank_statement、それ以外はpetty_cash
- クレカ明細の場合: すべてisIncome=false（支払）
- 品名や明細が読み取れない場合はitemsを空配列にする
- 読み取れない書類の場合はconfidenceを0.1以下にする
- レシートに10%税率と8%軽減税率の品目が混在する場合、itemsに各品目のtaxRateを必ず記載すること
- itemsのamountは税込金額で記載すること`;

const GEMINI_OCR_PROMPT = `この画像は日本語の会計・経理関連の書類です。
まず書類の種類を判定し、取引データをJSON形式で抽出してください。

${GEMINI_DOC_TYPES}

必ず以下のJSON形式のみを返してください（説明文や\`\`\`は不要）:
{
  "documentType": "書類種別",
  "confidence": 0.0〜1.0の読み取り確信度,
  "entries": [
    ${GEMINI_ENTRY_FORMAT}
  ]
}

${GEMINI_RULES}`;

// Gemini APIで書類を解析（単一PDF用・ストリーミング対応）
async function ocrWithGemini(arrayBuffer, onProgress) {
    const apiKey = getGeminiApiKey();
    if (!apiKey) throw new Error('Gemini APIキーが設定されていません');

    if (onProgress) onProgress({ status: 'PDF画像変換中...', progress: 0.1 });

    const base64Images = await pdfToBase64Images(arrayBuffer);

    if (onProgress) onProgress({ status: 'Gemini APIに送信中...', progress: 0.2 });

    const imageParts = base64Images.map(b64 => ({
        inline_data: { mime_type: 'image/png', data: b64 }
    }));

    const text = await fetchGeminiStreaming(apiKey, {
        contents: [{ parts: [...imageParts, { text: GEMINI_OCR_PROMPT }] }]
    }, (bytes, textLen, streamText) => {
        if (onProgress) {
            const p = Math.min(0.2 + (textLen / 800) * 0.6, 0.85);
            onProgress({ status: `Gemini応答受信中... ${textLen}文字`, progress: p, streamText });
        }
    });

    if (onProgress) onProgress({ status: '応答を解析中...', progress: 0.9 });

    const result = parseGeminiTextResponse(text);
    result.confidence = Number(result.confidence) || 0.5;
    result.documentType = result.documentType || 'other';
    if (!Array.isArray(result.entries)) result.entries = [];
    result.entries = result.entries.map(normalizeGeminiEntry);
    result._sourceImages = base64Images.map(b64 => `data:image/png;base64,${b64}`);
    return result;
}

// Gemini APIで画像（写真）を解析（複数レシート対応・ストリーミング）
async function ocrImageWithGemini(base64Data, mimeType, onProgress) {
    const apiKey = getGeminiApiKey();
    if (!apiKey) throw new Error('Gemini APIキーが設定されていません');

    if (onProgress) onProgress({ status: 'Gemini APIに送信中...', progress: 0.2 });

    const prompt = `この画像は日本語の会計・経理関連の書類（レシート、領収書等）の写真です。
1枚の画像に複数のレシートや領収書が写っている場合があります。
写っている全てのレシート・領収書を個別に識別し、それぞれの取引データを抽出してください。

${GEMINI_DOC_TYPES}

必ず以下のJSON形式のみを返してください（説明文や\`\`\`は不要）:
{
  "documentType": "書類種別",
  "confidence": 0.0〜1.0の読み取り確信度,
  "entries": [
    ${GEMINI_ENTRY_FORMAT}
  ]
}

${GEMINI_RULES}
- 1枚の写真に複数のレシートが写っている場合、各レシートを個別のentryとして返すこと
- 台紙に貼られたレシートやホッチキスで束になったレシートも個別に識別すること`;

    const text = await fetchGeminiStreaming(apiKey, {
        contents: [{
            parts: [
                { inline_data: { mime_type: mimeType, data: base64Data } },
                { text: prompt }
            ]
        }]
    }, (bytes, textLen, streamText) => {
        if (onProgress) {
            const p = Math.min(0.2 + (textLen / 600) * 0.6, 0.85);
            onProgress({ status: `Gemini応答受信中... ${textLen}文字`, progress: p, streamText });
        }
    });

    if (onProgress) onProgress({ status: '応答を解析中...', progress: 0.9 });

    const result = parseGeminiTextResponse(text);
    result.confidence = Number(result.confidence) || 0.5;
    result.documentType = result.documentType || 'receipt';
    if (!Array.isArray(result.entries)) result.entries = [];
    result.entries = result.entries.map(normalizeGeminiEntry);
    return result;
}

// Gemini APIで複数PDFを一括解析（バッチ処理）
async function ocrBatchWithGemini(imageDataArray, onProgress) {
    const apiKey = getGeminiApiKey();
    if (!apiKey) throw new Error('Gemini APIキーが設定されていません');

    if (onProgress) onProgress({ status: 'Gemini APIに一括送信中...', progress: 0.3 });

    // 全画像をpartsに並べる（各画像の前にラベルを付ける）
    const parts = [];
    imageDataArray.forEach((item, idx) => {
        parts.push({ text: `--- 書類 ${idx + 1}: ${item.fileName} ---` });
        item.base64Images.forEach(b64 => {
            parts.push({
                inline_data: { mime_type: 'image/png', data: b64 }
            });
        });
    });

    const prompt = `以下は${imageDataArray.length}件の日本語の会計・経理関連書類の画像です。
各書類を「--- 書類 N ---」のラベルで区切っています。
それぞれの書類について種類を判定し、取引データを抽出してください。

${GEMINI_DOC_TYPES}

必ず以下のJSON配列形式のみを返してください（説明文や\`\`\`は不要）:
[
  {
    "fileIndex": 1,
    "documentType": "書類種別",
    "confidence": 0.0〜1.0,
    "entries": [
      ${GEMINI_ENTRY_FORMAT}
    ]
  }
]

${GEMINI_RULES}
- fileIndexは1始まりで、書類の番号に対応させること
- 必ず${imageDataArray.length}件分の結果を返すこと`;

    parts.push({ text: prompt });

    const text = await fetchGeminiStreaming(apiKey, {
        contents: [{ parts }]
    }, (bytes, textLen, streamText) => {
        if (onProgress) {
            const p = Math.min(0.3 + (textLen / 1200) * 0.5, 0.85);
            onProgress({ status: `Gemini応答受信中... ${textLen}文字`, progress: p, streamText });
        }
    });

    if (onProgress) onProgress({ status: '応答を解析中...', progress: 0.9 });

    const results = parseGeminiTextResponse(text);
    if (!Array.isArray(results)) throw new Error('Gemini APIの応答が配列ではありません');

    return results.map(r => ({
        fileIndex: Number(r.fileIndex) || 0,
        documentType: r.documentType || 'other',
        confidence: Number(r.confidence) || 0.5,
        entries: (Array.isArray(r.entries) ? r.entries : []).map(normalizeGeminiEntry)
    }));
}

// 動画ファイルを直接Geminiに送信してOCR（フレーム抽出より高精度）
async function ocrVideoWithGemini(file, onProgress) {
    const apiKey = getGeminiApiKey();
    if (!apiKey) throw new Error('Gemini APIキーが設定されていません');

    const MAX_INLINE_SIZE = 15 * 1024 * 1024; // 15MB（base64化で約33%増加するため）

    if (file.size <= MAX_INLINE_SIZE) {
        return await ocrVideoInlineWithGemini(file, apiKey, onProgress);
    } else {
        return await ocrVideoWithFileApi(file, apiKey, onProgress);
    }
}

// 小さい動画: inline_dataで直接送信
async function ocrVideoInlineWithGemini(file, apiKey, onProgress) {
    if (onProgress) onProgress({ status: '動画を準備中...', progress: 0.1 });

    const base64 = await fileToBase64(file);
    const mimeType = file.type || 'video/mp4';

    if (onProgress) onProgress({ status: 'Gemini APIで動画を解析中...', progress: 0.3 });

    const prompt = buildVideoOcrPrompt();
    const text = await fetchGeminiStreaming(apiKey, {
        contents: [{
            parts: [
                { inline_data: { mime_type: mimeType, data: base64 } },
                { text: prompt }
            ]
        }]
    }, (bytes, textLen, streamText) => {
        if (onProgress) {
            const p = Math.min(0.3 + (textLen / 1000) * 0.5, 0.85);
            onProgress({ status: `Gemini応答受信中... ${textLen}文字`, progress: p, streamText });
        }
    });

    if (onProgress) onProgress({ status: '応答を解析中...', progress: 0.9 });

    return parseVideoOcrTextResponse(text);
}

// 大きい動画: File APIでアップロードしてから解析
async function ocrVideoWithFileApi(file, apiKey, onProgress) {
    if (onProgress) onProgress({ status: '動画をアップロード中...', progress: 0.1 });

    const mimeType = file.type || 'video/mp4';

    // 1. アップロードセッション開始
    const startResponse = await fetch(
        `https://generativelanguage.googleapis.com/upload/v1beta/files?key=${apiKey}`,
        {
            method: 'POST',
            headers: {
                'X-Goog-Upload-Protocol': 'resumable',
                'X-Goog-Upload-Command': 'start',
                'X-Goog-Upload-Header-Content-Length': String(file.size),
                'X-Goog-Upload-Header-Content-Type': mimeType,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                file: { displayName: file.name }
            })
        }
    );

    if (!startResponse.ok) {
        throw new Error(`ファイルアップロード開始エラー (${startResponse.status})`);
    }

    const uploadUrl = startResponse.headers.get('X-Goog-Upload-URL');
    if (!uploadUrl) throw new Error('アップロードURLが取得できませんでした');

    // 2. ファイルデータをアップロード
    if (onProgress) onProgress({ status: '動画をアップロード中...', progress: 0.3 });

    const uploadResponse = await fetch(uploadUrl, {
        method: 'PUT',
        headers: {
            'X-Goog-Upload-Offset': '0',
            'X-Goog-Upload-Command': 'upload, finalize',
            'Content-Length': String(file.size)
        },
        body: file
    });

    if (!uploadResponse.ok) {
        throw new Error(`ファイルアップロードエラー (${uploadResponse.status})`);
    }

    const uploadResult = await uploadResponse.json();
    const fileName = uploadResult.file?.name;
    const fileUri = uploadResult.file?.uri;

    if (!fileName || !fileUri) throw new Error('アップロード結果からファイル情報を取得できませんでした');

    // 3. ファイルがACTIVEになるまで待機
    if (onProgress) onProgress({ status: '動画の処理を待機中...', progress: 0.45 });

    await waitForGeminiFileActive(fileName, apiKey, onProgress);

    // 4. ストリーミングで解析
    if (onProgress) onProgress({ status: 'Gemini APIで動画を解析中...', progress: 0.6 });

    const prompt = buildVideoOcrPrompt();
    const text = await fetchGeminiStreaming(apiKey, {
        contents: [{
            parts: [
                { file_data: { file_uri: fileUri, mime_type: mimeType } },
                { text: prompt }
            ]
        }]
    }, (bytes, textLen, streamText) => {
        if (onProgress) {
            const p = Math.min(0.6 + (textLen / 1000) * 0.25, 0.88);
            onProgress({ status: `Gemini応答受信中... ${textLen}文字`, progress: p, streamText });
        }
    });

    if (onProgress) onProgress({ status: '応答を解析中...', progress: 0.9 });

    // 5. アップロードしたファイルを削除（クリーンアップ）
    try {
        await fetch(
            `https://generativelanguage.googleapis.com/v1beta/${fileName}?key=${apiKey}`,
            { method: 'DELETE' }
        );
    } catch (e) {
        console.warn('[Gemini] ファイル削除失敗（無視）:', e);
    }

    return parseVideoOcrTextResponse(text);
}

// Gemini File APIでファイルがACTIVEになるまでポーリング
async function waitForGeminiFileActive(fileName, apiKey, onProgress) {
    const MAX_WAIT = 120000; // 最大2分
    const POLL_INTERVAL = 3000; // 3秒間隔
    const start = Date.now();

    while (Date.now() - start < MAX_WAIT) {
        const res = await fetch(
            `https://generativelanguage.googleapis.com/v1beta/${fileName}?key=${apiKey}`
        );
        if (!res.ok) throw new Error(`ファイル状態確認エラー (${res.status})`);

        const fileInfo = await res.json();

        if (fileInfo.state === 'ACTIVE') return;
        if (fileInfo.state === 'FAILED') throw new Error('動画ファイルの処理に失敗しました');

        if (onProgress) {
            const elapsed = Math.min((Date.now() - start) / MAX_WAIT, 0.95);
            onProgress({ status: '動画の処理を待機中...', progress: 0.45 + elapsed * 0.1 });
        }

        await new Promise(resolve => setTimeout(resolve, POLL_INTERVAL));
    }

    throw new Error('動画ファイルの処理がタイムアウトしました');
}

// 動画OCR用のプロンプトを構築
function buildVideoOcrPrompt() {
    return `この動画にはレシート・領収書・請求書などの会計書類が映っています。
動画をめくりながら撮影しているため、複数の書類が順番に映ります。

全ての書類を個別に識別し、それぞれの取引データをJSON形式で抽出してください。
同じ書類が複数回映っている場合は、最も鮮明なフレームから1回だけ抽出してください（重複排除）。

${GEMINI_DOC_TYPES}

必ず以下のJSON配列形式のみを返してください（説明文や\`\`\`は不要）:
[
  {
    "timestamp": 動画内でこの書類が最も鮮明に映っているフレームの秒数（数値）,
    "documentType": "書類種別",
    "confidence": 0.0〜1.0の読み取り確信度,
    "entries": [
      ${GEMINI_ENTRY_FORMAT}
    ]
  }
]

${GEMINI_RULES}
- 動画内で映る全ての書類を個別のオブジェクトとして返すこと
- timestampは動画の先頭からの秒数（例: 3.5）で、その書類が最も読みやすい瞬間を指定すること
- ぼやけて読めない書類はconfidenceを0.1以下にして含めること
- 書類が映っていないフレームは無視すること`;
}

// 動画OCR結果の正規化（共通処理）
function normalizeVideoOcrResults(results) {
    if (Array.isArray(results)) {
        return results.map(r => ({
            timestamp: Number(r.timestamp) || 0,
            documentType: r.documentType || 'receipt',
            confidence: Number(r.confidence) || 0.5,
            entries: (Array.isArray(r.entries) ? r.entries : []).map(normalizeGeminiEntry)
        }));
    }

    if (results && typeof results === 'object') {
        return [{
            timestamp: Number(results.timestamp) || 0,
            documentType: results.documentType || 'receipt',
            confidence: Number(results.confidence) || 0.5,
            entries: (Array.isArray(results.entries) ? results.entries : []).map(normalizeGeminiEntry)
        }];
    }

    throw new Error('Gemini APIの応答を解析できませんでした');
}

// 動画OCR結果のパース（非ストリーミング用）
function parseVideoOcrResponse(data) {
    return normalizeVideoOcrResults(parseGeminiJsonResponse(data));
}

// 動画OCR結果のパース（ストリーミングテキスト版）
function parseVideoOcrTextResponse(text) {
    return normalizeVideoOcrResults(parseGeminiTextResponse(text));
}

// Gemini APIで勘定科目・税区分を推測（空フィールドを補完）
async function estimateAccountsWithGemini(rows, industry) {
    const apiKey = getGeminiApiKey();
    if (!apiKey) return 0;

    // 補完が必要な行を抽出（借方科目・貸方科目・税区分のいずれかが空）
    const targets = [];
    rows.forEach((row, idx) => {
        if (!row.tekiyou && !row.torihikisaki) return;
        const missing = [];
        if (!row.kariKamoku) missing.push('借方勘定科目');
        if (!row.kashiKamoku) missing.push('貸方勘定科目');
        if (!row.kariZeiku) missing.push('借方税区分');
        if (missing.length === 0) return;
        targets.push({
            idx,
            tekiyou: row.tekiyou || '',
            torihikisaki: row.torihikisaki || '',
            kariKamoku: row.kariKamoku || '',
            kashiKamoku: row.kashiKamoku || '',
            kariKingaku: row.kariKingaku || '',
            kashiKingaku: row.kashiKingaku || '',
            missing
        });
    });

    if (targets.length === 0) return 0;

    // 使用可能な勘定科目リストを構築
    const accountSet = new Set();
    ACCOUNT_DEFAULTS.forEach(entry => accountSet.add(entry[2]));
    COMMON_ACCOUNTS.forEach(a => accountSet.add(a));
    const accountList = [...accountSet].sort();

    // 税区分リスト
    const taxList = TAX_CATEGORIES.join('、');

    const prompt = `あなたは日本の記帳代行の専門家です。
以下の取引について、空欄のフィールドを推測してください。

使用可能な勘定科目（この中から選ぶこと）:
${accountList.join('、')}

使用可能な税区分:
${taxList}

${industry ? `業種: ${industry}` : ''}

取引一覧:
${targets.map((t, i) => {
    let info = `${i + 1}. 摘要「${t.tekiyou}」`;
    if (t.torihikisaki) info += ` 取引先「${t.torihikisaki}」`;
    if (t.kariKamoku) info += ` 借方科目「${t.kariKamoku}」`;
    if (t.kashiKamoku) info += ` 貸方科目「${t.kashiKamoku}」`;
    info += ` → 空欄: ${t.missing.join('・')}`;
    return info;
}).join('\n')}

必ず以下のJSON配列のみを返してください（説明文や\`\`\`は不要）:
[
  {
    "index": 1,
    "kariKamoku": "借方勘定科目（空欄の場合のみ推測、それ以外は空文字）",
    "kashiKamoku": "貸方勘定科目（空欄の場合のみ推測、それ以外は空文字）",
    "kariZeiku": "借方税区分（空欄の場合のみ推測、それ以外は空文字）",
    "confidence": 0.0〜1.0
  }
]

ルール:
- 空欄でないフィールドは空文字を返す（上書きしない）
- 不明な場合はconfidenceを0.3以下にする
- 税区分は取引内容と勘定科目から推測する（経費の仕入は通常「課税仕入10%」、軽減税率対象は「課税仕入8%（軽減）」等）`;

    try {
        const response = await fetch(
            `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`,
            {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    contents: [{ parts: [{ text: prompt }] }]
                })
            }
        );

        if (!response.ok) return 0;

        const data = await response.json();
        const text = data?.candidates?.[0]?.content?.parts?.[0]?.text || '';
        console.log('[Gemini] 科目推測応答:', text);

        let jsonStr = text.trim();
        const jsonMatch = jsonStr.match(/```(?:json)?\s*([\s\S]*?)```/);
        if (jsonMatch) jsonStr = jsonMatch[1].trim();

        const suggestions = JSON.parse(jsonStr);
        if (!Array.isArray(suggestions)) return 0;

        // 推測結果を適用（confidence 0.3以上のみ）
        let appliedCount = 0;
        suggestions.forEach(s => {
            const idx = Number(s.index) - 1;
            if (idx < 0 || idx >= targets.length) return;
            if (Number(s.confidence) < 0.3) return;

            const rowIdx = targets[idx].idx;
            const row = rows[rowIdx];
            let applied = false;

            // 借方勘定科目
            if (!row.kariKamoku && s.kariKamoku && accountList.includes(s.kariKamoku)) {
                row.kariKamoku = s.kariKamoku;
                applied = true;
            }
            // 貸方勘定科目
            if (!row.kashiKamoku && s.kashiKamoku && accountList.includes(s.kashiKamoku)) {
                row.kashiKamoku = s.kashiKamoku;
                applied = true;
            }
            // 借方税区分
            if (!row.kariZeiku && s.kariZeiku && TAX_CATEGORIES.includes(s.kariZeiku)) {
                row.kariZeiku = s.kariZeiku;
                applied = true;
            }

            if (applied) {
                row._geminiAccount = true;
                row._geminiConfidence = Number(s.confidence);
                appliedCount++;
            }
        });

        console.log(`[Gemini] 科目推測: ${appliedCount}/${targets.length} 件適用`);
        return appliedCount;
    } catch (e) {
        console.error('[Gemini] 科目推測エラー:', e);
        return 0;
    }
}
