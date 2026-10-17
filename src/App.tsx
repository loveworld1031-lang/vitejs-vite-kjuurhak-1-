import React, { useState, useEffect, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { supabase } from './supabaseClient';

// ==========================================
// 1. 進階規格標準化：內建水電「公稱直徑尺寸對照字典」
// ==========================================
function normalizeSpec(text: string): {
  cleanText: string;
  sizeTags: string[];
} {
  if (!text) return { cleanText: '', sizeTags: [] };

  // 🚀 先行將 Unicode 上標 ² 轉為 2，並移除多餘空白
  let clean = text
    .toLowerCase()
    .replace(/²/g, '2')
    .replace(/[\s\t\n\r]/g, '');
  let sizeTags: string[] = [];

  // 【A. 電線電纜截面積標準化】(mm2, mm², m㎡, SQ, 平方 ➡️ sq)
  const sqRegex = /(\d+(?:\.\d+)?)\s*(?:mm2|m㎡|sq|平方(?:公厘)?)/i;
  const sqMatch = clean.match(sqRegex);
  if (sqMatch) {
    sizeTags.push(`${sqMatch[1]}sq`);
  }

  // 【B & C. 支援異徑/多維度尺寸擷取】
  const inchRegex = /(\d+(?:\.\d+)?)\s*(?:"|吋|inch|φ|ф)/gi;
  const matches = [...clean.matchAll(inchRegex)];

  if (matches.length > 0) {
    matches.forEach((m) => sizeTags.push(`${m[1]}inch`));
  } else {
    // 【常規分數與水電代碼備援】
    let fractionText = clean;
    if (fractionText.includes('1/2'))
      fractionText = fractionText.replace(/1[-_]1\/2/g, '1.5');
    if (fractionText.includes('3/4'))
      fractionText = fractionText.replace(/1[-_]3\/4/g, '1.75');
    if (fractionText.includes('1/4'))
      fractionText = fractionText.replace(/1[-_]1\/4/g, '1.25');

    const fractionInchRegex = /(\d+)\s*\/\s*(\d+)\s*(?:"|吋|inch)?/i;
    const fractionMatch = fractionText.match(fractionInchRegex);
    if (fractionMatch) {
      const decimal =
        parseInt(fractionMatch[1], 10) / parseInt(fractionMatch[2], 10);
      sizeTags.push(`${decimal}inch`);
    } else {
      const fenRegex = /(\d+)分(?!之)/;
      const fenMatch = fractionText.match(fenRegex);
      if (fenMatch) {
        const decimal = parseInt(fenMatch[1], 10) / 8;
        sizeTags.push(`${decimal}inch`);
      } else if (/(?:^|\D)(?:300mm|300a|12吋|12")(?:$|\D)/i.test(clean))
        sizeTags.push('12inch');
      else if (/(?:^|\D)(?:250mm|250a|10吋|10")(?:$|\D)/i.test(clean))
        sizeTags.push('10inch');
      else if (/(?:^|\D)(?:200mm|200a|a200|8吋|8")(?:$|\D)/i.test(clean))
        sizeTags.push('8inch');
      else if (/(?:^|\D)(?:150mm|150a|w150|a150|6吋|6")(?:$|\D)/i.test(clean))
        sizeTags.push('6inch');
      else if (/(?:^|\D)(?:125mm|125a|w125|a125|5吋|5")(?:$|\D)/i.test(clean))
        sizeTags.push('5inch');
      else if (
        /(?:^|\D)(?:100mm|100a|e100|w100|a100|b100|4吋|4")(?:$|\D)/i.test(clean)
      )
        sizeTags.push('4inch');
      else if (
        /(?:^|\D)(?:75mm|77mm|80mm|80a|e80|w80|a80|b80|3吋|3")(?:$|\D)/i.test(
          clean
        )
      )
        sizeTags.push('3inch');
      else if (
        /(?:^|\D)(?:65mm|65a|e65|w65|a65|b65|2-1\/2|2\.5)(?:$|\D)/i.test(
          clean
        ) ||
        clean.includes('2-1/2')
      )
        sizeTags.push('2.5inch');
      else if (
        /(?:^|\D)(?:50mm|52mm|50a|e52|w52|w50|a50|b50|2吋|2")(?:$|\D)/i.test(
          clean
        )
      )
        sizeTags.push('2inch');
      else if (
        /(?:^|\D)(?:40mm|41mm|42mm|40a|e41|w41|w40|1-1\/2|1\.5|吋半)(?:$|\D)/i.test(
          clean
        ) ||
        clean.includes('1-1/2') ||
        clean.includes('吋半')
      )
        sizeTags.push('1.5inch');
      else if (
        /(?:^|\D)(?:32mm|35mm|32a|e35|w35|w32|1-1\/4|1\.25)(?:$|\D)/i.test(
          clean
        ) ||
        clean.includes('1-1/4')
      )
        sizeTags.push('1.25inch');
      else if (
        /(?:^|\D)(?:25mm|28mm|25a|e28|w28|w25|1吋|1")(?:$|\D)/i.test(clean)
      )
        sizeTags.push('1inch');
      else if (
        /(?:^|\D)(?:20mm|22mm|20a|e20|w20|3\/4)(?:$|\D)/i.test(clean) ||
        clean.includes('3/4')
      )
        sizeTags.push('0.75inch');
      else if (
        /(?:^|\D)(?:15mm|16mm|15a|16a|e16|w16|1\/2)(?:$|\D)/i.test(clean) ||
        clean.includes('1/2')
      )
        sizeTags.push('0.5inch');
      else if (
        /(?:^|\D)(?:13mm|13a|w13|3\/8)(?:$|\D)/i.test(clean) ||
        clean.includes('3/8')
      )
        sizeTags.push('0.375inch');
    }
  }

  // 過濾重複尺寸
  sizeTags = Array.from(new Set(sizeTags));

  return { cleanText: clean, sizeTags };
}

// ==========================================
// 1j. 系列產品模式泛化 (Pattern Skeleton Extractor)
// ==========================================
function extractPatternSkeleton(text: string): string {
  if (!text) return '';
  let clean = text
    .toLowerCase()
    .replace(/²/g, '2')
    .replace(/[\s\t\n\r]/g, '');
  // 移除電線電纜截面積 (如 5.5mm2, 14sq, 250mm2, 5.5平方)
  clean = clean.replace(/(\d+(?:\.\d+)?)\s*(?:mm2|m㎡|sq|平方(?:公厘)?)/gi, '');
  // 移除吋/分數/公稱直徑尺寸 (如 1/2", 3/4", 2", 1-1/2")
  clean = clean.replace(/(\d+(?:\.\d+)?)\s*(?:"|吋|inch|φ|ф)/gi, '');
  clean = clean.replace(/\d+\s*\/\s*\d+\s*(?:"|吋|inch)?/gi, '');
  clean = clean.replace(/\d+分(?!之)/gi, '');
  clean = clean.replace(
    /\b(65mm|65a|42mm|40a|35mm|32a|16mm|16a|15a|22mm|20a|28mm|25a|52mm|50a|77mm|80a|100mm|100a)\b/gi,
    ''
  );
  // 移除多餘分隔符號
  clean = clean.replace(/[，,_\-/#*()（）]/g, '');
  return clean;
}

// ==========================================
// 1b. 備援尺寸提取
// ==========================================
function resolveDbSizeTags(dbRow: DBItem): string[] {
  if (dbRow.size_spec) {
    const fromSpec = normalizeSpec(dbRow.size_spec).sizeTags;
    if (fromSpec.length > 0) return fromSpec;
  }
  const fromDescription = normalizeSpec(dbRow.description).sizeTags;
  if (fromDescription.length > 0) return fromDescription;
  const kwItems = dbRow.keywords.split(/[，,]/);
  for (const kw of kwItems) {
    const tags = normalizeSpec(kw).sizeTags;
    if (tags.length > 0) return tags;
  }
  return [];
}

// ==========================================
// 1c. 芯數比對
// ==========================================
function extractCoreCountTag(text: string): string {
  if (!text) return '';
  const t = text.toLowerCase().replace(/[\s\t\n\r]/g, '');
  const m = t.match(/(\d+)\s*(?:\/c|\/?c\b|芯|心)/i);
  return m ? `${m[1]}c` : '';
}

function resolveDbCoreCountTag(dbRow: DBItem): string {
  const fromDescription = extractCoreCountTag(dbRow.description);
  if (fromDescription) return fromDescription;
  const kwItems = dbRow.keywords.split(/[，,]/);
  for (const kw of kwItems) {
    const tag = extractCoreCountTag(kw);
    if (tag) return tag;
  }
  return '';
}

// ==========================================
// 1e. 材質比對
// ==========================================
const MATERIAL_KEYWORDS = [
  '鑄鐵',
  '白鐵',
  '不鏽鋼',
  'sus304',
  'sus316',
  '青銅',
  '砲金銅',
  '黃銅',
  '紅銅',
  '碳鋼',
  '鍍鋅',
  '塑鋼',
  'pvc',
  'pe',
  'upvc',
  'cpvc',
  'xlpe',
  '耐火',
  '耐燃',
  '耐熱',
  'fr',
];

function extractMaterialTag(text: string): string {
  if (!text) return '';
  const t = text.toLowerCase();
  if (/(xlpe|交連聚乙烯)/.test(t)) return 'XLPE';
  if (/(耐火|耐燃|fr\.c|\bfr\b)/.test(t)) return '耐燃';
  if (/耐熱/.test(t)) return '耐熱';
  if (/(pvc|聚氯乙烯)/.test(t)) return 'PVC';
  if (/(砲金銅|青銅)/.test(t)) return '砲金銅';
  if (/鑄鐵/.test(t)) return '鑄鐵';
  if (/(鍍鋅|gip|stpg)/.test(t)) return '鍍鋅鋼';
  if (/(白鐵|不鏽鋼|sus304|sus316)/.test(t)) return '不鏽鋼';
  for (const m of MATERIAL_KEYWORDS) {
    if (t.includes(m)) return m;
  }
  return '';
}

function resolveDbMaterialTag(dbRow: DBItem): string {
  return (
    dbRow.material ||
    extractMaterialTag(dbRow.item_tag) ||
    extractMaterialTag(dbRow.description) ||
    extractMaterialTag(dbRow.keywords)
  );
}

// ==========================================
// 1f. 產品型態/閥類比對
// ==========================================
function extractProductSubTypeTag(text: string): string {
  if (!text) return '';
  const t = text.toLowerCase();
  // ⚡ E管 / 導線管 / PVC電管
  if (
    /(?:e管|電管|導線管|灰色塑膠管|灰色管|pvc電管|pvc灰色|(?:^|[^a-z0-9])e[-_]?(?:16|20|28|35|41|52|65|80|100)(?:$|[^a-z0-9]))/i.test(
      t
    )
  )
    return 'E';
  // 💧 W管 / 給水管 / 自來水管
  if (
    /(?:w管|自來水|給水管|自來水管|pvc給水|pvc自來水|(?:^|[^a-z0-9])w[-_]?(?:13|16|20|25|28|35|40|41|50|52|65|80|100|125|150)(?:$|[^a-z0-9]))/i.test(
      t
    )
  )
    return 'W';
  // 🟠 A管 / 排水管 / 橘色管 / 透氣管 / 污水管 / 廢水管
  if (
    /(?:a管|排水|透氣|污水|廢水|橘管|橘色管|橘色|pvc排水|pvc橘色|(?:^|[^a-z0-9])a[-_]?(?:50|65|80|100|125|150|200)(?:$|[^a-z0-9]))/i.test(
      t
    )
  )
    return 'A';
  // 🔵 B管
  if (
    /(?:b管|一般用|(?:^|[^a-z0-9])b[-_]?(?:50|65|80|100)(?:$|[^a-z0-9]))/i.test(
      t
    )
  )
    return 'B';
  if (/發泡|dwv/.test(t)) return 'DWV';
  if (/vvr|控制電纜/.test(t)) return 'VVR';
  if (/耐火|耐燃|fr\.c|\bfr\b/.test(t)) return '耐火';
  if (/耐熱/.test(t)) return '耐熱';
  if (/xlpe|交連/.test(t)) return 'XLPE';
  if (/emt/.test(t)) return 'EMT';
  if (/sch40|stpg/.test(t)) return 'STPG';
  if (/gip|cns6445/.test(t)) return 'GIP';
  if (/閘閥|閘門閥/.test(t)) return '閘閥';
  if (/球塞閥|球塞/.test(t)) return '球塞閥';
  if (/y型過濾器|過濾器/.test(t)) return 'Y型過濾器';
  if (/逆止閥|逆止/.test(t)) return '逆止閥';
  if (/底閥/.test(t)) return '底閥';
  return '';
}

// ==========================================
// 1f2. 工程大類情境偵測 (電氣大類 vs 給排水大類)
// ==========================================
function detectCategoryFromText(
  text: string
): 'ELECTRICAL' | 'PLUMBING' | null {
  if (!text) return null;
  const t = text.toLowerCase();
  if (
    /(電氣|強電|弱電|電力|照明|配電|動力|電信|弱電設備|電氣工程|電氣設備|高低壓|導線|電纜|受配電)/i.test(
      t
    )
  ) {
    return 'ELECTRICAL';
  }
  if (
    /(給排水|給水|自來水|冷水|熱水|排水|衛生|污水|雨水|給水工程|排水工程|衛生工程|配管工程|衛生設備|給水設備)/i.test(
      t
    )
  ) {
    return 'PLUMBING';
  }
  return null;
}

// ==========================================
// 1g. 通用工程規格詞判斷
// ==========================================
function isGenericRatingCode(text: string): boolean {
  return /^\d+(\.\d+)?[a-z]{1,3}$/i.test(text);
}

// ==========================================
// 1h. Sørensen-Dice 雙字元相似度演算法
// ==========================================
function getBigrams(s: string): string[] {
  const grams: string[] = [];
  for (let i = 0; i < s.length - 1; i++) grams.push(s.slice(i, i + 2));
  return grams;
}

function calculateDiceSimilarity(str1: string, str2: string): number {
  if (!str1 || !str2) return 0;

  const set1 = new Set(getBigrams(str1));
  const set2 = new Set(getBigrams(str2));

  // 防呆：字串太短切不出 bigram，退回全等比對
  if (set1.size === 0 || set2.size === 0) {
    return str1 === str2 ? 1 : 0;
  }

  let intersectionCount = 0;
  for (const gram of set1) {
    if (set2.has(gram)) intersectionCount++;
  }
  return (2 * intersectionCount) / (set1.size + set2.size);
}

// ==========================================
// 1i. 差異關鍵字推導
// ==========================================
interface DifferentiatorEntry {
  rowId: number;
  ownKeywords: string[];
}
type DifferentiatorMap = Map<string, DifferentiatorEntry[]>;

function buildDifferentiatorMap(productDB: DBItem[]): DifferentiatorMap {
  const groups = new Map<string, DBItem[]>();
  for (const row of productDB) {
    const key = `${row.item_tag}|||${row.description}`;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key)!.push(row);
  }

  const result: DifferentiatorMap = new Map();
  for (const [key, rows] of groups.entries()) {
    if (rows.length < 2) continue;

    const kwSetsPerRow = rows.map(
      (r) =>
        new Set(
          r.keywords
            .split(/[，,]/)
            .map((k) => k.trim().toLowerCase())
            .filter(Boolean)
        )
    );

    const commonKw = kwSetsPerRow.reduce(
      (acc, set) => new Set([...acc].filter((x) => set.has(x))),
      new Set<string>(kwSetsPerRow[0] ?? [])
    );

    const entries: DifferentiatorEntry[] = rows.map((r, idx) => ({
      rowId: r.id,
      ownKeywords: [...kwSetsPerRow[idx]].filter(
        (k) => !commonKw.has(k) && !isGenericRatingCode(k)
      ),
    }));
    result.set(key, entries);
  }
  return result;
}

// ==========================================
// 2. 資料結構定義
// ==========================================
interface DBItem {
  id: number;
  item_tag: string;
  description: string;
  list_price: number;
  keywords: string;
  brand?: string;
  material?: string;
  product_type?: string;
  size_spec?: string;
  unit?: string;
}

interface CloudRawItem {
  id: number;
  item_tag: string;
  description: string;
  list_price: any;
  keywords: string;
  brand?: string;
  material?: string;
  product_type?: string;
  size_spec?: string;
  unit?: string;
}

interface MemoryPayload {
  description: string;
  item_tag: string;
}

interface PatternMemoryPayload {
  patternKey: string;
  seriesPrefix: string;
  itemTag: string;
  productType?: string;
  brand?: string;
}

interface MatchResult {
  rowNumber: number;
  clientName: string;
  matchedId: number | string;
  matchedDescription: string;
  listPrice: number;
  itemTag: string;
  confidenceScore: number;
  status: 'GREEN' | 'YELLOW' | 'RED' | 'MEMORY' | 'AMBIGUOUS';
  candidates: DBItem[];
  manualSearchQuery: string;
  isComposing: boolean;
  aiSuggested?: boolean;
  patternMatched?: boolean;
  inferredContextTag?: string;
}

interface OllamaMatchResponse {
  selectedIndex: number | null;
  confidence: number;
  reason: string;
}

interface DetectedInfo {
  sheetName: string;
  allSheetNames: string[];
  headerRow: number;
  nameCol: number;
  priceCol: number;
  qtyCol: number;
  unitCol: number;
  rawRows: any[][];
  previewRows: any[][];
  colCount: number;
}

const BLACKLIST_KEYWORDS =
  /(工資|運費|測試|安裝|配合費|另件|配件|小計|合計|總計|以下空白|備註)/i;

// ==========================================
// 3. AI 複核
// ==========================================
const AI_CONCURRENCY = 4;
const AI_CONFIDENCE_THRESHOLD = 60;

function normalizeOllamaEndpoint(rawUrl: string): {
  endpoint: string;
  isOpenAICompatible: boolean;
} {
  const trimmed = rawUrl.trim().replace(/\/+$/, '');
  if (/\/v1\/chat\/completions$/i.test(trimmed))
    return { endpoint: trimmed, isOpenAICompatible: true };
  if (/\/v1$/i.test(trimmed))
    return {
      endpoint: `${trimmed}/chat/completions`,
      isOpenAICompatible: true,
    };
  return { endpoint: trimmed, isOpenAICompatible: false };
}

function extractJsonFromModelOutput(raw: string): any {
  let text = (raw ?? '').trim();
  const fenceMatch = text.match(/```(?:json)?\s*([\s\S]*?)```/i);
  if (fenceMatch) text = fenceMatch[1].trim();
  if (!text.startsWith('{')) {
    const braceMatch = text.match(/\{[\s\S]*\}/);
    if (braceMatch) text = braceMatch[0];
  }
  return JSON.parse(text || '{}');
}

async function askOllamaForBestMatch(
  clientName: string,
  candidates: DBItem[],
  model: string,
  ollamaUrl: string,
  apiKey: string
): Promise<OllamaMatchResponse | null> {
  if (candidates.length === 0) return null;
  const candidateList = candidates
    .map(
      (c, idx) =>
        `${idx + 1}. [${c.item_tag}] ${c.description}（$${c.list_price}）`
    )
    .join('\n');
  const prompt = `你是水電材料規格判斷專家。請判斷下列「客戶標單品名」最符合哪一個候選規格。
客戶標單品名：「${clientName}」
候選清單：\n${candidateList}\n
規則：
- 若候選中有明確符合規格、尺寸、用途的項目，選出該項目編號
- 若全部候選都不符合（尺寸不對、材質不對、用途不對），selectedIndex 請回傳 null
- 只回傳純 JSON 物件本身，不要加 \`\`\` 或任何 Markdown 標記，也不要有其他文字，格式為：{"selectedIndex": number | null, "confidence": 0-100, "reason": "簡短理由"}`;

  try {
    const { endpoint, isOpenAICompatible } = normalizeOllamaEndpoint(ollamaUrl);
    const requestBody = isOpenAICompatible
      ? {
          model,
          messages: [{ role: 'user', content: prompt }],
          response_format: { type: 'json_object' },
          temperature: 0.1,
        }
      : {
          model,
          messages: [{ role: 'user', content: prompt }],
          format: 'json',
          stream: false,
          options: { temperature: 0.1 },
        };
    const res = await fetch(endpoint, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        ...(apiKey ? { Authorization: `Bearer ${apiKey}` } : {}),
      },
      body: JSON.stringify(requestBody),
    });
    if (!res.ok) throw new Error(`Ollama HTTP ${res.status}`);
    const data = await res.json();
    const rawContent = isOpenAICompatible
      ? data?.choices?.[0]?.message?.content
      : data?.message?.content;
    const parsed = extractJsonFromModelOutput(rawContent);
    const rawIndex = parsed.selectedIndex;
    const selectedIndex =
      typeof rawIndex === 'number' &&
      rawIndex >= 1 &&
      rawIndex <= candidates.length
        ? rawIndex - 1
        : null;
    return {
      selectedIndex,
      confidence: typeof parsed.confidence === 'number' ? parsed.confidence : 0,
      reason: typeof parsed.reason === 'string' ? parsed.reason : '',
    };
  } catch (err) {
    return null;
  }
}

function detectSheetStructure(rawRows: any[][]) {
  let headerRow = 1,
    nameCol = 2,
    priceCol = 5,
    qtyCol = 4,
    unitCol = 3,
    maxCols = 6;
  let found = false;
  for (let r = 0; r < Math.min(rawRows.length, 40); r++) {
    const row = rawRows[r] || [];
    if (row.length > maxCols) maxCols = row.length;
    let tName = -1,
      tPrice = -1,
      tQty = -1,
      tUnit = -1;
    for (let c = 0; c < row.length; c++) {
      const txt = String(row[c] || '')
        .trim()
        .replace(/\s+/g, '');
      if (/(品名|項目|材料|規格|說明)/.test(txt) && tPrice === -1)
        tName = c + 1;
      if (/(單價|牌價)/.test(txt)) tPrice = c + 1;
      if (/(數量|Qty)/.test(txt)) tQty = c + 1;
      if (/(單位|Unit)/.test(txt)) tUnit = c + 1;
    }
    if (tName !== -1 && tPrice !== -1) {
      headerRow = r + 1;
      nameCol = tName;
      priceCol = tPrice;
      qtyCol = tQty !== -1 ? tQty : tPrice - 1;
      unitCol = tUnit !== -1 ? tUnit : tName + 1;
      found = true;
      break;
    }
  }
  return { headerRow, nameCol, priceCol, qtyCol, unitCol, maxCols, found };
}

export default function App() {
  const [clientFile, setClientFile] = useState<File | null>(null);
  const [allSheetNames, setAllSheetNames] = useState<string[]>([]);
  const [selectedSheets, setSelectedSheets] = useState<string[]>([]);
  const [sheetStructures, setSheetStructures] = useState<{
    [sheetName: string]: DetectedInfo;
  }>({});
  const [activeSheetTab, setActiveSheetTab] = useState<string>('');
  const [allSheetResults, setAllSheetResults] = useState<{
    [sheetName: string]: MatchResult[];
  }>({});
  const workbookRef = useRef<any>(null);
  const [productDB, setProductDB] = useState<DBItem[]>([]);
  const differentiatorMap = useMemo(
    () => buildDifferentiatorMap(productDB),
    [productDB]
  );
  const [discounts, setDiscounts] = useState<{ [key: string]: number }>(() => {
    try {
      const saved = localStorage.getItem('mep_category_discounts');
      if (saved) return JSON.parse(saved);
    } catch (e) {}
    return { otherProductDiscount: 1.5 };
  });
  const [batchDiscountVal, setBatchDiscountVal] = useState<string>('1.5');
  const [statusMessage, setStatusMessage] = useState<string>(
    '📡正在連線至 Supabase 雲端大腦庫...'
  );
  const [isProcessing, setIsProcessing] = useState<boolean>(false);
  const [minScoreThreshold, setMinScoreThreshold] = useState<number>(45);

  const handleApplyBatchDiscount = (val: number) => {
    if (isNaN(val) || val < 0) return;
    setDiscounts((prev) => {
      const next = { ...prev };
      Object.keys(next).forEach((k) => {
        next[k] = val;
      });
      try {
        localStorage.setItem('mep_category_discounts', JSON.stringify(next));
      } catch (e) {}
      return next;
    });
  };

  const handleSingleDiscountChange = (key: string, val: number) => {
    const validVal = isNaN(val) ? 0 : val;
    setDiscounts((prev) => {
      const next = { ...prev, [key]: validVal };
      try {
        localStorage.setItem('mep_category_discounts', JSON.stringify(next));
      } catch (e) {}
      return next;
    });
  };

  const handleResetDiscounts = () => {
    setDiscounts((prev) => {
      const next = { ...prev };
      Object.keys(next).forEach((k) => {
        next[k] = 1.5;
      });
      try {
        localStorage.setItem('mep_category_discounts', JSON.stringify(next));
      } catch (e) {}
      return next;
    });
  };

  const [ollamaEnabled, setOllamaEnabled] = useState<boolean>(true);
  const [ollamaModel, setOllamaModel] = useState<string>('qwen2.5:7b');
  const [ollamaUrl, setOllamaUrl] = useState<string>(
    'http://localhost:11434/api/chat'
  );
  const [ollamaApiKey, setOllamaApiKey] = useState<string>('');
  const [aiProgress, setAiProgress] = useState<{
    done: number;
    total: number;
  } | null>(null);
  const [ollamaTestStatus, setOllamaTestStatus] = useState<string>('');
  const [isTestingOllama, setIsTestingOllama] = useState<boolean>(false);

  const transformRawData = (rawData: CloudRawItem[]): DBItem[] => {
    return rawData.map((item: CloudRawItem) => {
      const parsedPrice = parseFloat(
        String(item.list_price).replace(/[,*]/g, '')
      );
      return {
        id: Number(item.id),
        item_tag: String(item.item_tag || '').trim(),
        description: String(item.description || '').trim(),
        list_price: isNaN(parsedPrice) ? 0 : parsedPrice,
        keywords: String(item.keywords || '').trim(),
        brand: item.brand ? String(item.brand).trim() : undefined,
        material: item.material ? String(item.material).trim() : undefined,
        product_type: item.product_type
          ? String(item.product_type).trim()
          : undefined,
        size_spec: item.size_spec ? String(item.size_spec).trim() : undefined,
        unit: item.unit ? String(item.unit).trim() : undefined,
      };
    });
  };

  useEffect(() => {
    async function loadCloudDB() {
      try {
        const { data, error } = await supabase
          .from('price_db')
          .select(
            'id, item_tag, description, list_price, keywords, brand, material, product_type, size_spec, unit'
          );
        if (error) throw error;
        if (data && data.length > 0) {
          const standardizedDB = transformRawData(
            data as unknown as CloudRawItem[]
          );
          const categories = new Set<string>();
          standardizedDB.forEach((item: DBItem) =>
            categories.add(item.item_tag)
          );

          let savedDiscounts: { [key: string]: number } = {};
          try {
            const raw = localStorage.getItem('mep_category_discounts');
            if (raw) savedDiscounts = JSON.parse(raw);
          } catch (e) {}

          const initialDiscounts: { [key: string]: number } = {
            otherProductDiscount: savedDiscounts['otherProductDiscount'] ?? 1.5,
          };
          categories.forEach((cat: string) => {
            initialDiscounts[cat] = savedDiscounts[cat] ?? 1.5;
          });

          setDiscounts(initialDiscounts);
          try {
            localStorage.setItem(
              'mep_category_discounts',
              JSON.stringify(initialDiscounts)
            );
          } catch (e) {}

          setProductDB(standardizedDB);
          setStatusMessage(
            `💡 雲端連線成功！已自動載入最新牌價庫（共 ${standardizedDB.length} 筆）與上次設定的分類折數記憶。請上傳客戶標單！`
          );
        } else {
          setStatusMessage(
            '⚠️ 雲端 price_db 資料表為空，請先至 Supabase 後台匯入 CSV。'
          );
        }
      } catch (err: any) {
        setStatusMessage(
          `❌ 雲端大腦連線失敗: ${
            err?.message || '請確認 supabaseClient.ts 是否正常設定'
          }。`
        );
      }
    }
    loadCloudDB();
  }, []);

  const handleClientFileChange = async (
    e: React.ChangeEvent<HTMLInputElement>
  ) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setClientFile(file);
    setAllSheetResults({});
    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: 'array' });
      workbookRef.current = workbook;

      const structures: { [sheetName: string]: DetectedInfo } = {};
      const validSheets: string[] = [];

      for (const sName of workbook.SheetNames) {
        const sheet = workbook.Sheets[sName];
        const rawRows = XLSX.utils.sheet_to_json<any[]>(sheet, {
          header: 1,
          defval: '',
        });
        const structure = detectSheetStructure(rawRows);
        const previewRows = rawRows
          .slice(0, 12)
          .map((row: any[]) =>
            Array.from(
              { length: Math.max(structure.maxCols, row.length) },
              (_, c: number) => String(row[c] || '')
            )
          );

        structures[sName] = {
          sheetName: sName,
          allSheetNames: workbook.SheetNames,
          headerRow: structure.headerRow,
          nameCol: structure.nameCol,
          priceCol: structure.priceCol,
          qtyCol: structure.qtyCol,
          unitCol: structure.unitCol,
          rawRows,
          previewRows,
          colCount: structure.maxCols,
        };

        const isSummarySheet =
          /(總表|總彙總|彙總|目錄|封面|標單首頁|投標須知|投標清單|注意事項|計價標準)/i.test(
            sName
          );
        if (structure.found && !isSummarySheet) {
          validSheets.push(sName);
        }
      }

      const defaultSelected =
        validSheets.length > 0
          ? validSheets
          : workbook.SheetNames.length > 0
          ? [workbook.SheetNames[0]]
          : [];
      const initialActive = defaultSelected[0] || workbook.SheetNames[0] || '';

      setAllSheetNames(workbook.SheetNames);
      setSheetStructures(structures);
      setSelectedSheets(defaultSelected);
      setActiveSheetTab(initialActive);

      setStatusMessage(
        `✅ 標單載入成功！活頁簿內共偵測到 ${workbook.SheetNames.length} 個工作表（已自動為您勾選 ${defaultSelected.length} 個明細表）。請設定折數後點擊「⚡ 一鍵全活頁簿智慧安全比對」！`
      );
    } catch (err: any) {
      setStatusMessage(`❌ 標單載入失敗: ${err.message}`);
    }
  };

  const handleToggleSelectSheet = (sheetName: string) => {
    setSelectedSheets((prev) =>
      prev.includes(sheetName)
        ? prev.filter((s) => s !== sheetName)
        : [...prev, sheetName]
    );
  };

  const handleSelectAllDetectedSheets = () => {
    const detectedNames = allSheetNames.filter((s) => {
      const isSummary =
        /(總表|總彙總|彙總|目錄|封面|標單首頁|投標須知|投標清單|注意事項)/i.test(
          s
        );
      return !isSummary;
    });
    setSelectedSheets(detectedNames.length > 0 ? detectedNames : allSheetNames);
  };

  const handleDeselectAllSheets = () => {
    setSelectedSheets([]);
  };

  const handleApplyColumnsToAllSheets = (sourceSheetName: string) => {
    const src = sheetStructures[sourceSheetName];
    if (!src) return;
    setSheetStructures((prev) => {
      const next = { ...prev };
      selectedSheets.forEach((sName) => {
        if (next[sName]) {
          next[sName] = {
            ...next[sName],
            nameCol: src.nameCol,
            priceCol: src.priceCol,
            qtyCol: src.qtyCol,
            unitCol: src.unitCol,
          };
        }
      });
      return next;
    });
    setStatusMessage(
      `📋 已將「${sourceSheetName}」的欄位設定（品名: ${getColLetter(
        src.nameCol
      )}欄, 單價: ${getColLetter(src.priceCol)}欄）套用至所有已選取工作表！`
    );
  };

  // 🚀 單一工作表智慧比對核心演算法
  const matchSheetData = async (
    targetSheetName: string,
    struct: DetectedInfo,
    cloudMemoryLib: { [key: string]: string },
    localMemory: { [key: string]: MemoryPayload },
    patternMemory: { [key: string]: PatternMemoryPayload }
  ): Promise<MatchResult[]> => {
    const results: MatchResult[] = [];
    const startRowIdx = struct.headerRow;
    const totalRows = struct.rawRows.length;

    let currentMajorCategory: 'ELECTRICAL' | 'PLUMBING' | 'UNKNOWN' =
      detectCategoryFromText(targetSheetName) || 'UNKNOWN';

    let i = startRowIdx;
    let loopCounter = 0;

    while (i < totalRows) {
      if (++loopCounter % 30 === 0) {
        await new Promise((resolve) => setTimeout(resolve, 0));
      }

      const row = struct.rawRows[i];
      if (!row) {
        i++;
        continue;
      }

      const fullRowText = row
        .slice(0, Math.min(row.length, 8))
        .map((c: any) => String(c || ''))
        .join(' ');
      const rowCategory = detectCategoryFromText(fullRowText);
      if (rowCategory) {
        currentMajorCategory = rowCategory;
      }

      let clientName = String(row[struct.nameCol - 1] || '').trim();
      const clientQty = String(row[struct.qtyCol - 1] || '').trim();
      const clientUnit = String(row[struct.unitCol - 1] || '').trim();
      const parsedQty = parseFloat(clientQty.replace(/,/g, ''));

      if (
        !clientName ||
        !clientUnit ||
        clientQty === '' ||
        isNaN(parsedQty) ||
        parsedQty === 0
      ) {
        i++;
        continue;
      }

      let nextRowIdx = i + 1;
      while (nextRowIdx < totalRows) {
        const nextRow = struct.rawRows[nextRowIdx];
        if (!nextRow) break;
        const nextName = String(nextRow[struct.nameCol - 1] || '').trim();
        const nextQty = String(nextRow[struct.qtyCol - 1] || '').trim();
        const nextUnit = String(nextRow[struct.unitCol - 1] || '').trim();

        if (
          nextName !== '' &&
          (nextQty === '' || parseFloat(nextQty.replace(/,/g, '')) === 0) &&
          nextUnit === ''
        ) {
          clientName += ' ' + nextName;
          nextRowIdx++;
        } else break;
      }

      const clientNorm = normalizeSpec(clientName);
      const lookupKey = clientNorm.cleanText;
      const clientCoreTag = extractCoreCountTag(clientNorm.cleanText);
      const clientMaterialTag = extractMaterialTag(clientNorm.cleanText);

      const isPvcPipe =
        /(pvc|塑膠|聚氯乙烯|硬質塑膠|塑膠管|pvc管)/i.test(clientName) &&
        /(管|管路|導線管|電管|給水管|排水管)/i.test(clientName);
      let clientSubType = extractProductSubTypeTag(clientName);
      let inferredContextTag = '';

      if (!clientSubType && isPvcPipe) {
        if (currentMajorCategory === 'ELECTRICAL') {
          clientSubType = 'E';
          inferredContextTag = '⚡ 電氣大類 (判定E管)';
        } else if (currentMajorCategory === 'PLUMBING') {
          if (/(橘色|橘管|污水|排水|透氣|廢水|落水|a管)/i.test(clientName)) {
            clientSubType = 'A';
            inferredContextTag = '🟠 給排水 (橘管判定A管)';
          } else {
            clientSubType = 'W';
            inferredContextTag = '💧 給排水 (預設判定W管)';
          }
        }
      }

      if (BLACKLIST_KEYWORDS.test(clientName)) {
        i = nextRowIdx;
        continue;
      }

      let targetDesc = '';
      let forcedTag = '';
      if (localMemory[lookupKey]) {
        targetDesc = localMemory[lookupKey].description;
        forcedTag = localMemory[lookupKey].item_tag;
      } else if (cloudMemoryLib[lookupKey]) {
        targetDesc = cloudMemoryLib[lookupKey];
      }

      if (targetDesc) {
        if (targetDesc === 'REMOVE_MAPPING') {
          results.push({
            rowNumber: i + 1,
            clientName,
            matchedId: 'REMOVE_MAPPING',
            matchedDescription: '不套用牌價（手動排除項）',
            listPrice: 0,
            itemTag: 'otherProductDiscount',
            confidenceScore: 100,
            status: 'RED',
            candidates: [],
            manualSearchQuery: '',
            isComposing: false,
          });
          i = nextRowIdx;
          continue;
        }
        const dbFound = productDB.find((d: DBItem) =>
          forcedTag
            ? d.description === targetDesc && d.item_tag === forcedTag
            : d.description === targetDesc
        );
        if (dbFound) {
          results.push({
            rowNumber: i + 1,
            clientName,
            matchedId: dbFound.id,
            matchedDescription: dbFound.description,
            listPrice: dbFound.list_price,
            itemTag: dbFound.item_tag,
            confidenceScore: 100,
            status: 'MEMORY',
            candidates: [dbFound],
            manualSearchQuery: '',
            isComposing: false,
            patternMatched: false,
          });
          i = nextRowIdx;
          continue;
        }
      }

      const clientSkeleton = extractPatternSkeleton(clientName);
      if (
        !targetDesc &&
        clientSkeleton &&
        clientSkeleton.length >= 3 &&
        patternMemory[clientSkeleton]
      ) {
        const pRule = patternMemory[clientSkeleton];
        const candidateInSeries = productDB.find((d: DBItem) => {
          const dbSizeTags = resolveDbSizeTags(d);
          const sizeMatch =
            clientNorm.sizeTags.length > 0 &&
            dbSizeTags.length > 0 &&
            clientNorm.sizeTags.every((t) => dbSizeTags.includes(t));
          if (!sizeMatch) return false;

          const dbCoreTag = resolveDbCoreCountTag(d);
          if (clientCoreTag && dbCoreTag && clientCoreTag !== dbCoreTag)
            return false;

          const tagMatch = !pRule.itemTag || d.item_tag === pRule.itemTag;
          const typeMatch =
            pRule.productType && d.product_type
              ? d.product_type === pRule.productType
              : true;
          const descSkeleton = extractPatternSkeleton(d.description);
          const seriesMatch = pRule.seriesPrefix
            ? descSkeleton.includes(pRule.seriesPrefix) ||
              pRule.seriesPrefix.includes(descSkeleton)
            : true;

          return tagMatch && (typeMatch || seriesMatch);
        });

        if (candidateInSeries) {
          results.push({
            rowNumber: i + 1,
            clientName,
            matchedId: candidateInSeries.id,
            matchedDescription: candidateInSeries.description,
            listPrice: candidateInSeries.list_price,
            itemTag: candidateInSeries.item_tag,
            confidenceScore: 98,
            status: 'MEMORY',
            candidates: [candidateInSeries],
            manualSearchQuery: '',
            isComposing: false,
            aiSuggested: false,
            patternMatched: true,
          });
          i = nextRowIdx;
          continue;
        }
      }

      const scoredItems: {
        dbRow: DBItem;
        score: number;
        differentiatorAmbiguous: boolean;
        hasTextEvidence: boolean;
      }[] = [];

      for (const dbRow of productDB) {
        const dbFullText = (
          dbRow.item_tag +
          ' ' +
          dbRow.description +
          ' ' +
          dbRow.keywords
        ).toLowerCase();

        const isClientPipe =
          /(管|管路|導線管|電管|配管)/i.test(clientName) &&
          !/(線|電線|電纜)/i.test(clientName);
        const isDbWire =
          /(電線|電纜|花線|控制線|絕緣線|裸銅線|平分線|絞線)/i.test(dbFullText);
        if (isClientPipe && isDbWire) continue;

        const isClientWire =
          /(電線|電纜|線|絞線|平分線)/i.test(clientName) &&
          !/(管|導線管)/i.test(clientName);
        const isDbPipe =
          /(管|管路|導線管|電管|給水管|排水管)/i.test(dbFullText) &&
          !/(電線|電纜)/i.test(dbRow.item_tag);
        if (isClientWire && isDbPipe) continue;

        const dbNorm = normalizeSpec(dbRow.description);
        const dbSizeTags = resolveDbSizeTags(dbRow);

        let sizeScoreBonus = 0;
        if (clientNorm.sizeTags.length > 0 && dbSizeTags.length > 0) {
          const intersection = clientNorm.sizeTags.filter((tag) =>
            dbSizeTags.includes(tag)
          );
          if (intersection.length === clientNorm.sizeTags.length) {
            sizeScoreBonus = 45;
          } else if (intersection.length > 0) {
            sizeScoreBonus = 10;
          } else {
            sizeScoreBonus = -70;
          }
        }

        let coreCountBonus = 0;
        const dbCoreTag = resolveDbCoreCountTag(dbRow);
        if (clientCoreTag && dbCoreTag) {
          coreCountBonus = clientCoreTag === dbCoreTag ? 35 : -45;
        }

        const dbMaterialTag = resolveDbMaterialTag(dbRow);
        if (
          clientMaterialTag &&
          dbMaterialTag &&
          clientMaterialTag !== dbMaterialTag
        )
          continue;
        const materialBonus = clientMaterialTag && dbMaterialTag ? 45 : 0;

        const dbSubType =
          dbRow.product_type ||
          extractProductSubTypeTag(dbRow.item_tag) ||
          extractProductSubTypeTag(dbRow.description) ||
          extractProductSubTypeTag(dbRow.keywords);
        let subTypeBonus = 0;
        if (clientSubType && dbSubType) {
          subTypeBonus = clientSubType === dbSubType ? 50 : -80;
        }

        let brandBonus = 0;
        if (
          dbRow.brand &&
          clientNorm.cleanText.includes(dbRow.brand.toLowerCase())
        ) {
          brandBonus = 15;
        }

        const ambiguityGroup = differentiatorMap.get(
          `${dbRow.item_tag}|||${dbRow.description}`
        );
        let differentiatorBonus = 0;
        let differentiatorAmbiguous = false;

        if (ambiguityGroup) {
          const thisEntry = ambiguityGroup.find((e) => e.rowId === dbRow.id);
          const thisOwnKw = thisEntry?.ownKeywords ?? [];
          const clientHasOwnKw = thisOwnKw.some(
            (kw) => kw && clientNorm.cleanText.includes(kw)
          );
          const clientHasOtherKw = ambiguityGroup.some(
            (e) =>
              e.rowId !== dbRow.id &&
              e.ownKeywords.some(
                (kw) => kw && clientNorm.cleanText.includes(kw)
              )
          );
          if (clientHasOwnKw) differentiatorBonus = 60;
          else if (clientHasOtherKw) continue;
          else differentiatorAmbiguous = true;
        }

        const kwArray = dbRow.keywords
          .split(/[，,]/)
          .map((k: string) => k.trim().toLowerCase())
          .filter((k: string) => k);
        let matchCount = 0;
        let meaningfulMatchCount = 0;
        let coreFeatureBonus = 0;

        kwArray.forEach((kw: string) => {
          const kwNorm = normalizeSpec(kw);
          const normKw = kwNorm.cleanText;
          if (normKw && clientNorm.cleanText.indexOf(normKw) !== -1) {
            matchCount++;
            if (kwNorm.sizeTags.length === 0 && !isGenericRatingCode(normKw))
              meaningfulMatchCount++;
            if (
              /(橘色|排水|廢水|給水|灰色|e管|b管|絕緣|電力電纜|耐熱|閘閥|球塞|過濾器|逆止|底閥)/.test(
                normKw
              )
            )
              coreFeatureBonus += 20;
          }
        });

        let keywordBaseScore = 0;
        if (kwArray.length > 0) {
          keywordBaseScore = Math.min(40, meaningfulMatchCount * 15);
        }

        const combinedDbText =
          normalizeSpec(dbRow.item_tag).cleanText + dbNorm.cleanText;
        const similarity = calculateDiceSimilarity(
          clientNorm.cleanText,
          combinedDbText
        );
        let textOverlapBonus = Math.round(similarity * 35);

        const hasTextEvidence =
          meaningfulMatchCount > 0 ||
          differentiatorBonus > 0 ||
          coreFeatureBonus > 0 ||
          subTypeBonus > 0 ||
          similarity > 0.15;
        const totalScore =
          keywordBaseScore +
          sizeScoreBonus +
          coreCountBonus +
          materialBonus +
          subTypeBonus +
          brandBonus +
          differentiatorBonus +
          coreFeatureBonus +
          textOverlapBonus;

        scoredItems.push({
          dbRow,
          score: Math.max(0, Math.min(100, totalScore)),
          differentiatorAmbiguous,
          hasTextEvidence,
        });
      }

      scoredItems.sort((a, b) => b.score - a.score);
      const topCandidates = scoredItems.slice(0, 4).map((x) => x.dbRow);
      const bestMatch = scoredItems[0];

      let status: 'GREEN' | 'YELLOW' | 'RED' | 'AMBIGUOUS' = 'RED';
      let finalDescription = `未找到相符規格（匹配得分低於 ${minScoreThreshold} 分已自動過濾）`;
      let finalPrice = 0,
        finalTag = 'otherProductDiscount';
      let finalId: number | string = 'REMOVE_MAPPING';

      if (bestMatch && bestMatch.score >= minScoreThreshold) {
        if (!bestMatch.hasTextEvidence) {
          status = 'RED';
          finalId = 'REMOVE_MAPPING';
          finalPrice = 0;
          finalTag = 'otherProductDiscount';
          finalDescription = `⚠️ 資料庫無相符品項（僅尺寸巧合，缺乏文字佐證，原始分數:${bestMatch.score}）`;
        } else {
          finalId = bestMatch.dbRow.id;
          finalDescription = bestMatch.dbRow.description;
          finalPrice = bestMatch.dbRow.list_price;
          finalTag = bestMatch.dbRow.item_tag;
          if (bestMatch.differentiatorAmbiguous) {
            status = 'AMBIGUOUS';
          } else {
            const secondBest = scoredItems[1];
            const scoreGap = secondBest
              ? bestMatch.score - secondBest.score
              : 999;
            const isTooClose =
              secondBest &&
              secondBest.dbRow.id !== bestMatch.dbRow.id &&
              scoreGap < 15;
            status = bestMatch.score >= 80 && !isTooClose ? 'GREEN' : 'YELLOW';
          }
        }
      }
      results.push({
        rowNumber: i + 1,
        clientName,
        matchedId: finalId,
        matchedDescription: finalDescription,
        listPrice: finalPrice,
        itemTag: finalTag,
        confidenceScore:
          bestMatch && bestMatch.hasTextEvidence ? bestMatch.score : 0,
        status,
        candidates: topCandidates,
        manualSearchQuery: '',
        isComposing: false,
        inferredContextTag,
      });
      i = nextRowIdx;
    }
    return results;
  };

  const handleStartMatching = async () => {
    if (selectedSheets.length === 0 || productDB.length === 0) return;
    setIsProcessing(true);
    setStatusMessage(
      '⚡ 正在同步調取雲端與地端記憶，執行全活頁簿多工作表智慧比對...'
    );

    try {
      const { data: cloudMemory, error: memError } = await supabase
        .from('learning_memory')
        .select('client_name, matched_description');
      const cloudMemoryLib: { [key: string]: string } = {};
      if (!memError && cloudMemory) {
        cloudMemory.forEach((row: any) => {
          const cleanKey = normalizeSpec(row.client_name).cleanText;
          if (cleanKey) cloudMemoryLib[cleanKey] = row.matched_description;
        });
      }

      let localMemory: { [key: string]: MemoryPayload } = {};
      try {
        const stored = localStorage.getItem('optimized_learning_memory');
        if (stored) localMemory = JSON.parse(stored);
      } catch (e) {}

      let patternMemory: { [key: string]: PatternMemoryPayload } = {};
      try {
        const storedP = localStorage.getItem(
          'optimized_pattern_learning_memory'
        );
        if (storedP) patternMemory = JSON.parse(storedP);
      } catch (e) {}

      const newResultsMap: { [sheetName: string]: MatchResult[] } = {};

      for (let sIdx = 0; sIdx < selectedSheets.length; sIdx++) {
        const sName = selectedSheets[sIdx];
        const struct = sheetStructures[sName];
        if (!struct) continue;

        setStatusMessage(
          `⚡ 正在智慧比對工作表 (${sIdx + 1}/${
            selectedSheets.length
          })：「${sName}」...`
        );
        const sheetResults = await matchSheetData(
          sName,
          struct,
          cloudMemoryLib,
          localMemory,
          patternMemory
        );
        newResultsMap[sName] = sheetResults;
      }

      setAllSheetResults(newResultsMap);
      if (
        !selectedSheets.includes(activeSheetTab) &&
        selectedSheets.length > 0
      ) {
        setActiveSheetTab(selectedSheets[0]);
      }

      let grandTotal = 0,
        grandGreen = 0,
        grandMem = 0,
        grandYellow = 0,
        grandRed = 0,
        grandAmbiguous = 0;
      Object.values(newResultsMap).forEach((list) => {
        grandTotal += list.length;
        list.forEach((r) => {
          if (r.status === 'GREEN') grandGreen++;
          else if (r.status === 'MEMORY') grandMem++;
          else if (r.status === 'YELLOW') grandYellow++;
          else if (r.status === 'AMBIGUOUS') grandAmbiguous++;
          else if (r.status === 'RED') grandRed++;
        });
      });

      setStatusMessage(
        `🎉 全活頁簿比對完成！共 ${selectedSheets.length} 個工作表、${grandTotal} 筆品項（🟢 高信心: ${grandGreen} 筆，🧠 大腦記憶: ${grandMem} 筆，🟡 待確認: ${grandYellow} 筆，⚠️ 資訊不足: ${grandAmbiguous} 筆，🔴 排除: ${grandRed} 筆）。請於下方頁籤切換覆核！`
      );

      if (ollamaEnabled) {
        await runAiBatchRecheck(newResultsMap);
      }
    } catch (e: any) {
      setStatusMessage(`❌ 比對失敗: ${e.message}`);
    } finally {
      setIsProcessing(false);
    }
  };

  const runAiBatchRecheck = async (resultsMap: {
    [sheetName: string]: MatchResult[];
  }) => {
    const allTargets: { sheetName: string; r: MatchResult; idx: number }[] = [];
    Object.entries(resultsMap).forEach(([sName, list]) => {
      list.forEach((r, idx) => {
        if (
          (r.status === 'YELLOW' || r.status === 'RED') &&
          r.candidates.length > 0
        ) {
          allTargets.push({ sheetName: sName, r, idx });
        }
      });
    });

    if (allTargets.length === 0) return;

    setAiProgress({ done: 0, total: allTargets.length });
    setStatusMessage(
      `🤖 正在啟動地端 AI（${ollamaModel}）跨表語意複核... (0/${allTargets.length})`
    );
    let doneCount = 0,
      failCount = 0,
      cursor = 0;

    const worker = async () => {
      while (cursor < allTargets.length) {
        const myTurn = cursor;
        cursor++;
        const { sheetName, r, idx } = allTargets[myTurn];
        const originalMatchedId = r.matchedId;
        const aiResult = await askOllamaForBestMatch(
          r.clientName,
          r.candidates,
          ollamaModel,
          ollamaUrl,
          ollamaApiKey
        );
        if (!aiResult) failCount++;

        doneCount++;
        setAiProgress({ done: doneCount, total: allTargets.length });
        setStatusMessage(
          `🤖 地端 AI 跨表語意複核中... (${doneCount}/${allTargets.length})`
        );

        if (
          aiResult &&
          aiResult.selectedIndex !== null &&
          aiResult.confidence >= AI_CONFIDENCE_THRESHOLD
        ) {
          const picked = r.candidates[aiResult.selectedIndex];
          if (picked) {
            setAllSheetResults((prev) => {
              const currentList = prev[sheetName];
              if (!currentList) return prev;
              const nextList = [...currentList];
              const target = nextList[idx];
              if (!target || target.matchedId !== originalMatchedId)
                return prev;
              nextList[idx] = {
                ...target,
                matchedId: picked.id,
                matchedDescription: picked.description,
                listPrice: picked.list_price,
                itemTag: picked.item_tag,
                confidenceScore: aiResult.confidence,
                status: 'YELLOW',
                aiSuggested: true,
              };
              return { ...prev, [sheetName]: nextList };
            });
          }
        }
      }
    };
    await Promise.all(
      Array.from({ length: Math.min(AI_CONCURRENCY, allTargets.length) }, () =>
        worker()
      )
    );
    setAiProgress(null);
    setStatusMessage(
      failCount > 0
        ? `⚠️ 地端 AI 複核完成，但 ${failCount}/${allTargets.length} 筆呼叫失敗。`
        : `🎉 地端 AI 跨表複核完成！請於下方人工覆核面板各頁籤中確認 🤖 標籤的建議。`
    );
  };

  const handleCandidateChange = async (
    index: number,
    selectedIdStr: string
  ) => {
    const currentList = allSheetResults[activeSheetTab];
    if (!currentList) return;
    const nextResults = [...currentList];
    const target = nextResults[index];
    let localMemory: { [key: string]: MemoryPayload } = {};
    try {
      const stored = localStorage.getItem('optimized_learning_memory');
      if (stored) localMemory = JSON.parse(stored);
    } catch (e) {}
    const cleanLookupKey = normalizeSpec(target.clientName).cleanText;
    if (!cleanLookupKey) return;

    if (selectedIdStr === 'REMOVE_MAPPING') {
      target.matchedId = 'REMOVE_MAPPING';
      target.matchedDescription = '不套用牌價（手動排除）';
      target.listPrice = 0;
      target.itemTag = 'otherProductDiscount';
      target.confidenceScore = 0;
      target.status = 'RED';
      target.aiSuggested = false;
      target.patternMatched = false;
      setAllSheetResults((prev) => ({
        ...prev,
        [activeSheetTab]: nextResults,
      }));
      localMemory[cleanLookupKey] = {
        description: 'REMOVE_MAPPING',
        item_tag: 'otherProductDiscount',
      };
      localStorage.setItem(
        'optimized_learning_memory',
        JSON.stringify(localMemory)
      );
      await supabase
        .from('learning_memory')
        .upsert({
          client_name: target.clientName.trim(),
          matched_description: 'REMOVE_MAPPING',
        });
    } else {
      const targetId = Number(selectedIdStr);
      const found = productDB.find((c: DBItem) => c.id === targetId);
      if (found) {
        target.matchedId = found.id;
        target.matchedDescription = found.description;
        target.listPrice = found.list_price;
        target.itemTag = found.item_tag;
        target.confidenceScore = 100;
        target.status = 'MEMORY';
        target.aiSuggested = false;
        target.patternMatched = false;
        setAllSheetResults((prev) => ({
          ...prev,
          [activeSheetTab]: nextResults,
        }));

        localMemory[cleanLookupKey] = {
          description: found.description,
          item_tag: found.item_tag,
        };
        localStorage.setItem(
          'optimized_learning_memory',
          JSON.stringify(localMemory)
        );
        await supabase
          .from('learning_memory')
          .upsert({
            client_name: target.clientName.trim(),
            matched_description: found.description,
          });

        const skeletonKey = extractPatternSkeleton(target.clientName);
        if (skeletonKey && skeletonKey.length >= 3) {
          let patternMemory: { [key: string]: PatternMemoryPayload } = {};
          try {
            const storedP = localStorage.getItem(
              'optimized_pattern_learning_memory'
            );
            if (storedP) patternMemory = JSON.parse(storedP);
          } catch (e) {}

          const dbSkeleton = extractPatternSkeleton(found.description);
          patternMemory[skeletonKey] = {
            patternKey: skeletonKey,
            seriesPrefix: dbSkeleton,
            itemTag: found.item_tag,
            productType: found.product_type,
            brand: found.brand,
          };
          localStorage.setItem(
            'optimized_pattern_learning_memory',
            JSON.stringify(patternMemory)
          );
        }
      }
    }
  };

  const handleSearchTextChange = (index: number, val: string) => {
    setAllSheetResults((prev) => {
      const list = prev[activeSheetTab];
      if (!list) return prev;
      const next = [...list];
      next[index].manualSearchQuery = val;
      return { ...prev, [activeSheetTab]: next };
    });
  };

  const handleCompositionStart = (index: number) => {
    setAllSheetResults((prev) => {
      const list = prev[activeSheetTab];
      if (!list) return prev;
      const next = [...list];
      next[index].isComposing = true;
      return { ...prev, [activeSheetTab]: next };
    });
  };

  const handleCompositionEnd = (index: number, val: string) => {
    setAllSheetResults((prev) => {
      const list = prev[activeSheetTab];
      if (!list) return prev;
      const next = [...list];
      next[index].isComposing = false;
      next[index].manualSearchQuery = val;
      return { ...prev, [activeSheetTab]: next };
    });
  };

  const handleExportFile = async () => {
    if (!clientFile || Object.keys(allSheetResults).length === 0) return;
    setIsProcessing(true);
    setStatusMessage(
      '⏳ 正在複製原標單樣式、批次回填所有工作表之折後單價，並自動同步大腦學習記憶...'
    );
    try {
      const buffer = await clientFile.arrayBuffer();
      const workbook = new ExcelJS.Workbook();
      const isOldXls = clientFile.name.endsWith('.xls');

      let useExcelJS = true;
      try {
        if (isOldXls) {
          const tempXlsxWorkbook = XLSX.read(buffer, { type: 'array' });
          const xlsxWriteBuffer = XLSX.write(tempXlsxWorkbook, {
            bookType: 'xlsx',
            type: 'array',
          });
          await workbook.xlsx.load(xlsxWriteBuffer);
        } else {
          try {
            await workbook.xlsx.load(buffer);
          } catch (loadErr: any) {
            console.warn(
              'ExcelJS 原生載入失敗（可能含有特殊控制字元），啟動 SheetJS XML 淨化器重載:',
              loadErr
            );
            const tempXlsxWorkbook = XLSX.read(buffer, { type: 'array' });
            const sanitizedBuffer = XLSX.write(tempXlsxWorkbook, {
              bookType: 'xlsx',
              type: 'array',
            });
            await workbook.xlsx.load(sanitizedBuffer);
          }
        }
      } catch (excelJsErr: any) {
        console.warn(
          'ExcelJS 無法載入此活頁簿，切換至 SheetJS 高相容回填模式:',
          excelJsErr
        );
        useExcelJS = false;
      }

      let localMemory: { [key: string]: MemoryPayload } = {};
      try {
        const stored = localStorage.getItem('optimized_learning_memory');
        if (stored) localMemory = JSON.parse(stored);
      } catch (e) {}

      let patternMemory: { [key: string]: PatternMemoryPayload } = {};
      try {
        const storedP = localStorage.getItem(
          'optimized_pattern_learning_memory'
        );
        if (storedP) patternMemory = JSON.parse(storedP);
      } catch (e) {}

      const cloudUpsertPayloads: {
        client_name: string;
        matched_description: string;
      }[] = [];
      let finalBlob: Blob;
      let totalExportedItems = 0;
      let totalExportedSheets = 0;

      if (useExcelJS) {
        Object.entries(allSheetResults).forEach(([sName, items]) => {
          if (items.length === 0) return;
          const worksheet = workbook.getWorksheet(sName);
          const struct = sheetStructures[sName];
          if (!worksheet || !struct) return;
          totalExportedSheets++;

          items.forEach((item: MatchResult) => {
            if (item.listPrice === 0 || item.matchedId === 'REMOVE_MAPPING') {
              if (item.matchedId === 'REMOVE_MAPPING') {
                const cleanKey = normalizeSpec(item.clientName).cleanText;
                if (cleanKey) {
                  localMemory[cleanKey] = {
                    description: 'REMOVE_MAPPING',
                    item_tag: 'otherProductDiscount',
                  };
                  cloudUpsertPayloads.push({
                    client_name: item.clientName.trim(),
                    matched_description: 'REMOVE_MAPPING',
                  });
                }
              }
              return;
            }

            totalExportedItems++;
            const multiplier =
              discounts[item.itemTag] !== undefined
                ? discounts[item.itemTag]
                : discounts['otherProductDiscount'];
            const validMultiplier = isNaN(multiplier) ? 1.0 : multiplier;
            const finalPrice = Math.round(item.listPrice * validMultiplier);

            const row = worksheet.getRow(item.rowNumber);
            const cell = row.getCell(struct.priceCol);
            cell.value = Number(finalPrice);
            cell.numFmt = '#,##0';
            cell.font = { color: { argb: 'FFFF0000' }, bold: true };

            const cleanKey = normalizeSpec(item.clientName).cleanText;
            if (cleanKey && item.matchedDescription) {
              localMemory[cleanKey] = {
                description: item.matchedDescription,
                item_tag: item.itemTag,
              };
              cloudUpsertPayloads.push({
                client_name: item.clientName.trim(),
                matched_description: item.matchedDescription,
              });

              const skeletonKey = extractPatternSkeleton(item.clientName);
              if (skeletonKey && skeletonKey.length >= 3) {
                const dbSkeleton = extractPatternSkeleton(
                  item.matchedDescription
                );
                const foundDbItem = productDB.find(
                  (c: DBItem) => c.id === Number(item.matchedId)
                );
                patternMemory[skeletonKey] = {
                  patternKey: skeletonKey,
                  seriesPrefix: dbSkeleton,
                  itemTag: item.itemTag,
                  productType: foundDbItem?.product_type,
                  brand: foundDbItem?.brand,
                };
              }
            }
          });
        });

        workbook.calcProperties.fullCalcOnLoad = true;
        const writeBuffer = await workbook.xlsx.writeBuffer();
        finalBlob = new Blob([writeBuffer], {
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        });
      } else {
        const sheetJsWb = XLSX.read(buffer, {
          type: 'array',
          cellStyles: true,
        });

        Object.entries(allSheetResults).forEach(([sName, items]) => {
          if (items.length === 0) return;
          const sheet = sheetJsWb.Sheets[sName];
          const struct = sheetStructures[sName];
          if (!sheet || !struct) return;
          totalExportedSheets++;

          items.forEach((item: MatchResult) => {
            if (item.listPrice === 0 || item.matchedId === 'REMOVE_MAPPING') {
              if (item.matchedId === 'REMOVE_MAPPING') {
                const cleanKey = normalizeSpec(item.clientName).cleanText;
                if (cleanKey) {
                  localMemory[cleanKey] = {
                    description: 'REMOVE_MAPPING',
                    item_tag: 'otherProductDiscount',
                  };
                  cloudUpsertPayloads.push({
                    client_name: item.clientName.trim(),
                    matched_description: 'REMOVE_MAPPING',
                  });
                }
              }
              return;
            }

            totalExportedItems++;
            const multiplier =
              discounts[item.itemTag] !== undefined
                ? discounts[item.itemTag]
                : discounts['otherProductDiscount'];
            const validMultiplier = isNaN(multiplier) ? 1.0 : multiplier;
            const finalPrice = Math.round(item.listPrice * validMultiplier);

            const cellRef = XLSX.utils.encode_cell({
              r: item.rowNumber - 1,
              c: struct.priceCol - 1,
            });
            sheet[cellRef] = { t: 'n', v: finalPrice };

            const cleanKey = normalizeSpec(item.clientName).cleanText;
            if (cleanKey && item.matchedDescription) {
              localMemory[cleanKey] = {
                description: item.matchedDescription,
                item_tag: item.itemTag,
              };
              cloudUpsertPayloads.push({
                client_name: item.clientName.trim(),
                matched_description: item.matchedDescription,
              });

              const skeletonKey = extractPatternSkeleton(item.clientName);
              if (skeletonKey && skeletonKey.length >= 3) {
                const dbSkeleton = extractPatternSkeleton(
                  item.matchedDescription
                );
                const foundDbItem = productDB.find(
                  (c: DBItem) => c.id === Number(item.matchedId)
                );
                patternMemory[skeletonKey] = {
                  patternKey: skeletonKey,
                  seriesPrefix: dbSkeleton,
                  itemTag: item.itemTag,
                  productType: foundDbItem?.product_type,
                  brand: foundDbItem?.brand,
                };
              }
            }
          });
        });

        const writeBuffer = XLSX.write(sheetJsWb, {
          bookType: 'xlsx',
          type: 'array',
        });
        finalBlob = new Blob([writeBuffer], {
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        });
      }

      localStorage.setItem(
        'optimized_learning_memory',
        JSON.stringify(localMemory)
      );
      localStorage.setItem(
        'optimized_pattern_learning_memory',
        JSON.stringify(patternMemory)
      );

      if (cloudUpsertPayloads.length > 0) {
        const uniquePayloads = Array.from(
          new Map(cloudUpsertPayloads.map((p) => [p.client_name, p])).values()
        );
        try {
          await supabase.from('learning_memory').upsert(uniquePayloads);
        } catch (dbErr) {
          console.warn('雲端記憶同步略過:', dbErr);
        }
      }

      // 將所有工作表的項目狀態標記為 MEMORY
      setAllSheetResults((prev) => {
        const next: { [sName: string]: MatchResult[] } = {};
        Object.entries(prev).forEach(([sName, list]) => {
          next[sName] = list.map((r) =>
            r.listPrice > 0 && r.matchedId !== 'REMOVE_MAPPING'
              ? { ...r, status: 'MEMORY', confidenceScore: 100 }
              : r
          );
        });
        return next;
      });

      const url = URL.createObjectURL(finalBlob);
      const a = document.createElement('a');
      const newFileName =
        clientFile.name.replace(/\.xls[x]?$/, '') + '_全活頁簿智慧回填.xlsx';
      a.href = url;
      a.download = newFileName;
      a.click();
      URL.revokeObjectURL(url);
      setStatusMessage(
        `🏆 整本標單匯出成功！已完整回填 ${totalExportedSheets} 個工作表（共 ${totalExportedItems} 筆品項單價），並自動將本次所有認可規格存入大腦記憶與系列泛化庫。新檔名：[${newFileName}]`
      );
    } catch (err: any) {
      setStatusMessage(`❌ 匯出失敗: ${err.message}`);
    } finally {
      setIsProcessing(false);
    }
  };

  const handleClearMemoryCache = () => {
    if (
      window.confirm(
        '⚠️ 確定要清空本機瀏覽器的智慧學習紀錄與折數記憶嗎？(包含精確記憶、系列泛化記憶與分類折數設定，雲端資料庫不會受影響)'
      )
    ) {
      localStorage.removeItem('optimized_learning_memory');
      localStorage.removeItem('optimized_pattern_learning_memory');
      localStorage.removeItem('mep_category_discounts');
      alert('🧼 本地大腦記憶與分類折數設定已重置清空。');
      if (Object.keys(allSheetResults).length > 0) handleStartMatching();
    }
  };

  const handleTestOllamaConnection = async () => {
    setIsTestingOllama(true);
    setOllamaTestStatus('⏳ 測試連線中...');
    try {
      const testCandidates: DBItem[] = [
        {
          id: -1,
          item_tag: '測試類別',
          description: '1/2吋 PVC 給水管',
          list_price: 100,
          keywords: '給水,pvc,1/2吋',
        },
      ];
      const result = await askOllamaForBestMatch(
        '測試品名 1/2吋給水管',
        testCandidates,
        ollamaModel,
        ollamaUrl,
        ollamaApiKey
      );
      if (result)
        setOllamaTestStatus(`✅ 連線成功！模型「${ollamaModel}」回應正常`);
      else setOllamaTestStatus('❌ 連線失敗或回應解析失敗，請查看 Console。');
    } finally {
      setIsTestingOllama(false);
    }
  };

  const getColLetter = (index: number) => String.fromCharCode(64 + index);
  const discountKeys = useMemo(() => Object.keys(discounts), [discounts]);

  const activeStruct = sheetStructures[activeSheetTab];
  const activeResults = allSheetResults[activeSheetTab] || [];
  const hasAnyResults = Object.keys(allSheetResults).length > 0;

  return (
    <div className="min-h-screen bg-slate-100 p-6 text-slate-800 font-sans">
      <div className="max-w-7xl mx-auto bg-white rounded-2xl shadow-xl p-6 space-y-6">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
          <div>
            <h1 className="text-2xl font-black text-emerald-600">
              🏢 昇億公司專用系統{' '}
              <span className="text-xs bg-indigo-600 text-white px-2 py-0.5 rounded ml-2">
                全活頁簿多頁籤智配版
              </span>
            </h1>
            <p className="text-xs text-slate-400 mt-1">
              支援多工作表自動辨識、批次跨表智配、多頁籤覆核面板與一鍵整本 Excel
              活頁簿回填匯出。
            </p>
          </div>
          <button
            onClick={handleClearMemoryCache}
            className="text-[11px] self-start sm:self-center bg-slate-200 hover:bg-red-100 text-slate-500 hover:text-red-700 px-3 py-1.5 rounded-lg font-bold transition"
          >
            🧼 清空本地大腦快取
          </button>
        </div>

        {/* 步驟 1：上傳客戶標單 */}
        <div className="border rounded-xl p-4 bg-slate-50 border-dashed border-slate-300 text-center">
          <h3 className="font-bold text-sm text-slate-700 mb-2">
            📋 步驟 1：請點擊或拖曳上傳客戶標單 (.xls / .xlsx)
          </h3>
          <input
            type="file"
            accept=".xlsx,.xls"
            onChange={handleClientFileChange}
            className="text-sm cursor-pointer mx-auto"
          />
          {clientFile && (
            <p className="mt-2 text-xs text-emerald-600 font-bold">
              🎯 當前處理標單：{clientFile.name}（共 {allSheetNames.length}{' '}
              個工作表）
            </p>
          )}
        </div>

        {/* 步驟 2：多工作表選取與欄位校對管理 */}
        {allSheetNames.length > 0 && (
          <div className="border border-amber-200 bg-amber-50/40 rounded-xl p-4 space-y-4">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 border-b border-amber-200 pb-2">
              <div>
                <h3 className="font-extrabold text-sm text-amber-900">
                  📑 步驟 2：多工作表管理與欄位校對面板
                </h3>
                <p className="text-xs text-amber-700">
                  請勾選本次需要回填單價的工作表（已自動為您勾選明細表，排除彙總總表）：
                </p>
              </div>
              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={handleSelectAllDetectedSheets}
                  className="text-[11px] bg-white hover:bg-amber-100 text-amber-900 font-bold px-2.5 py-1 rounded border border-amber-300 transition"
                >
                  ☑️ 全選明細表
                </button>
                <button
                  type="button"
                  onClick={handleDeselectAllSheets}
                  className="text-[11px] bg-white hover:bg-slate-100 text-slate-600 font-bold px-2.5 py-1 rounded border border-slate-300 transition"
                >
                  ⬜ 全部取消
                </button>
              </div>
            </div>

            {/* 工作表勾選標籤清單 */}
            <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-2.5">
              {allSheetNames.map((sName) => {
                const struct = sheetStructures[sName];
                const isSelected = selectedSheets.includes(sName);
                const isActive = activeSheetTab === sName;
                const isSummary =
                  /(總表|總彙總|彙總|目錄|封面|標單首頁|投標須知|投標清單|注意事項)/i.test(
                    sName
                  );

                return (
                  <div
                    key={`sheet-card-${sName}`}
                    onClick={() => setActiveSheetTab(sName)}
                    className={`p-2.5 rounded-lg border text-xs cursor-pointer transition flex items-center justify-between ${
                      isActive
                        ? 'ring-2 ring-blue-500 bg-white shadow-sm'
                        : 'bg-white/80 hover:bg-white'
                    } ${
                      isSelected
                        ? 'border-amber-400'
                        : 'border-slate-200 opacity-70'
                    }`}
                  >
                    <label
                      className="flex items-center gap-2 cursor-pointer flex-1 min-w-0"
                      onClick={(e) => e.stopPropagation()}
                    >
                      <input
                        type="checkbox"
                        checked={isSelected}
                        onChange={() => handleToggleSelectSheet(sName)}
                        className="accent-blue-600 w-4 h-4 rounded"
                      />
                      <span
                        className="font-bold text-slate-800 truncate"
                        title={sName}
                      >
                        {sName}
                      </span>
                    </label>
                    <div className="flex items-center gap-1 ml-2 flex-shrink-0">
                      {isSummary ? (
                        <span className="text-[9px] bg-slate-100 text-slate-500 px-1.5 py-0.5 rounded font-medium">
                          彙總表
                        </span>
                      ) : struct?.found ? (
                        <span className="text-[9px] bg-emerald-100 text-emerald-800 px-1.5 py-0.5 rounded font-bold">
                          品名:{getColLetter(struct.nameCol)} 單價:
                          {getColLetter(struct.priceCol)}
                        </span>
                      ) : (
                        <span className="text-[9px] bg-orange-100 text-orange-800 px-1.5 py-0.5 rounded font-bold">
                          待確認
                        </span>
                      )}
                    </div>
                  </div>
                );
              })}
            </div>

            {/* 當前選定工作表細部欄位設定 */}
            {activeStruct && (
              <div className="bg-white p-3 rounded-lg border border-amber-300 space-y-2">
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 border-b pb-2">
                  <span className="text-xs font-black text-blue-900">
                    🔧 設定「{activeSheetTab}」欄位對應（第{' '}
                    {activeStruct.headerRow} 列起）：
                  </span>
                  <button
                    type="button"
                    onClick={() =>
                      handleApplyColumnsToAllSheets(activeSheetTab)
                    }
                    className="text-[11px] bg-blue-50 hover:bg-blue-100 text-blue-800 font-bold px-3 py-1 rounded border border-blue-300 transition"
                  >
                    📋 將此工作表欄位套用至所有已選取工作表
                  </button>
                </div>
                <div className="grid grid-cols-2 sm:grid-cols-4 gap-4 text-xs pt-1">
                  <div>
                    <label className="block font-bold text-blue-700 mb-0.5">
                      📦 材料品名欄：
                    </label>
                    <select
                      value={activeStruct.nameCol}
                      onChange={(e) =>
                        setSheetStructures({
                          ...sheetStructures,
                          [activeSheetTab]: {
                            ...activeStruct,
                            nameCol: parseInt(e.target.value),
                          },
                        })
                      }
                      className="w-full border rounded p-1 bg-white font-bold"
                    >
                      {Array.from({ length: activeStruct.colCount }).map(
                        (_, i) => (
                          <option key={`nameCol-${i}`} value={i + 1}>
                            {getColLetter(i + 1)} 欄
                          </option>
                        )
                      )}
                    </select>
                  </div>
                  <div>
                    <label className="block font-bold text-green-700 mb-0.5">
                      💰 單價回填欄：
                    </label>
                    <select
                      value={activeStruct.priceCol}
                      onChange={(e) =>
                        setSheetStructures({
                          ...sheetStructures,
                          [activeSheetTab]: {
                            ...activeStruct,
                            priceCol: parseInt(e.target.value),
                          },
                        })
                      }
                      className="w-full border rounded p-1 bg-white font-bold"
                    >
                      {Array.from({ length: activeStruct.colCount }).map(
                        (_, i) => (
                          <option key={`priceCol-${i}`} value={i + 1}>
                            {getColLetter(i + 1)} 欄
                          </option>
                        )
                      )}
                    </select>
                  </div>
                  <div>
                    <label className="block font-bold text-slate-700 mb-0.5">
                      🔢 數量欄：
                    </label>
                    <select
                      value={activeStruct.qtyCol}
                      onChange={(e) =>
                        setSheetStructures({
                          ...sheetStructures,
                          [activeSheetTab]: {
                            ...activeStruct,
                            qtyCol: parseInt(e.target.value),
                          },
                        })
                      }
                      className="w-full border rounded p-1 bg-white"
                    >
                      {Array.from({ length: activeStruct.colCount }).map(
                        (_, i) => (
                          <option key={`qtyCol-${i}`} value={i + 1}>
                            {getColLetter(i + 1)} 欄
                          </option>
                        )
                      )}
                    </select>
                  </div>
                  <div>
                    <label className="block font-bold text-purple-700 mb-0.5">
                      📏 單位欄（過濾用）：
                    </label>
                    <select
                      value={activeStruct.unitCol}
                      onChange={(e) =>
                        setSheetStructures({
                          ...sheetStructures,
                          [activeSheetTab]: {
                            ...activeStruct,
                            unitCol: parseInt(e.target.value),
                          },
                        })
                      }
                      className="w-full border rounded p-1 bg-white"
                    >
                      {Array.from({ length: activeStruct.colCount }).map(
                        (_, i) => (
                          <option key={`unitCol-${i}`} value={i + 1}>
                            {getColLetter(i + 1)} 欄
                          </option>
                        )
                      )}
                    </select>
                  </div>
                </div>
              </div>
            )}
          </div>
        )}

        {/* 步驟 3：產品分類計價折數 (倍率) 設定面板 */}
        {productDB.length > 0 && (
          <div className="border border-blue-200 rounded-xl p-4 bg-blue-50/40 space-y-3">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
              <div className="flex items-center gap-2">
                <h3 className="font-extrabold text-sm text-blue-900">
                  💰 步驟 3：產品分類計價折數 (倍率) 面板：
                </h3>
                <span className="text-[10px] bg-blue-600 text-white font-bold px-2 py-0.5 rounded shadow-sm">
                  🧠 自動記憶中
                </span>
              </div>
              <span className="text-xs text-blue-700 font-medium">
                💡
                提示：在此設定折數，所有工作表單價與匯出將即時聯動計算，且會自動為您記住！
              </span>
            </div>

            {/* 快速批次套用工具列 */}
            <div className="flex flex-wrap items-center gap-2 bg-white/80 p-2.5 rounded-lg border border-blue-200 text-xs">
              <span className="font-bold text-blue-900">⚡ 快速批次套用：</span>
              <button
                type="button"
                onClick={() => handleApplyBatchDiscount(1.0)}
                className="bg-slate-50 hover:bg-blue-50 text-blue-800 font-bold px-2.5 py-1 rounded border border-blue-300 transition text-[11px] active:scale-95"
              >
                1.0 (原價)
              </button>
              <button
                type="button"
                onClick={() => handleApplyBatchDiscount(0.85)}
                className="bg-slate-50 hover:bg-blue-50 text-blue-800 font-bold px-2.5 py-1 rounded border border-blue-300 transition text-[11px] active:scale-95"
              >
                0.85 (85折)
              </button>
              <button
                type="button"
                onClick={() => handleApplyBatchDiscount(0.9)}
                className="bg-slate-50 hover:bg-blue-50 text-blue-800 font-bold px-2.5 py-1 rounded border border-blue-300 transition text-[11px] active:scale-95"
              >
                0.9 (9折)
              </button>
              <button
                type="button"
                onClick={() => handleApplyBatchDiscount(1.15)}
                className="bg-slate-50 hover:bg-blue-50 text-blue-800 font-bold px-2.5 py-1 rounded border border-blue-300 transition text-[11px] active:scale-95"
              >
                1.15 (溢價15%)
              </button>
              <button
                type="button"
                onClick={() => handleApplyBatchDiscount(1.5)}
                className="bg-slate-50 hover:bg-blue-50 text-blue-800 font-bold px-2.5 py-1 rounded border border-blue-300 transition text-[11px] active:scale-95"
              >
                1.5 (預設1.5倍)
              </button>
              <button
                type="button"
                onClick={handleResetDiscounts}
                className="bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold px-2.5 py-1 rounded border border-slate-300 transition text-[11px] active:scale-95"
              >
                🔄 重設全部為 1.5
              </button>
              <div className="flex items-center gap-1.5 ml-auto">
                <span className="text-slate-500 font-medium text-[11px]">
                  自訂全部：
                </span>
                <input
                  type="number"
                  step="0.01"
                  min="0"
                  max="5"
                  value={batchDiscountVal}
                  onChange={(e) => setBatchDiscountVal(e.target.value)}
                  placeholder="倍率"
                  className="w-16 border border-slate-300 rounded p-1 text-center bg-white text-xs font-bold text-blue-600 focus:outline-none focus:border-blue-500"
                />
                <button
                  type="button"
                  onClick={() =>
                    handleApplyBatchDiscount(parseFloat(batchDiscountVal))
                  }
                  className="bg-blue-600 hover:bg-blue-700 text-white px-3 py-1 rounded font-bold text-[11px] shadow transition active:scale-95"
                >
                  套用全部
                </button>
              </div>
            </div>

            <div className="grid grid-cols-2 sm:grid-cols-4 gap-4">
              {discountKeys.map((key: string) => (
                <div
                  key={`discount-box-${key}`}
                  className="bg-white p-2.5 rounded-lg border shadow-sm flex flex-col justify-between"
                >
                  <label className="block text-[10px] font-bold text-slate-400 truncate mb-1">
                    {key === 'otherProductDiscount' ? '📦 其他未分類產品' : key}
                  </label>
                  <div className="flex items-center border-b border-slate-200 focus-within:border-blue-500">
                    <input
                      type="number"
                      step="0.01"
                      min="0"
                      max="5"
                      value={discounts[key]}
                      onChange={(e) =>
                        handleSingleDiscountChange(
                          key,
                          parseFloat(e.target.value)
                        )
                      }
                      className="w-full text-sm font-black text-blue-600 border-none p-0 focus:ring-0"
                    />
                    <span className="text-xs font-bold text-slate-400 ml-1">
                      倍
                    </span>
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}

        {/* 步驟 4：執行智慧比對與參數設定 */}
        <div className="p-4 rounded-xl border bg-emerald-50 text-emerald-900 border-emerald-200 space-y-4">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
            <div>
              <h3 className="font-bold text-xs text-emerald-800 mb-1">
                ⚡ 步驟 4：全活頁簿智慧安全比對控制台：
              </h3>
              <span className="text-sm font-bold block">{statusMessage}</span>
            </div>
            {selectedSheets.length > 0 && productDB.length > 0 && (
              <button
                onClick={handleStartMatching}
                disabled={isProcessing}
                className="w-full sm:w-auto bg-emerald-600 hover:bg-emerald-700 disabled:opacity-50 text-white font-black py-2.5 px-8 rounded-lg shadow-md transition transform active:scale-95"
              >
                {isProcessing
                  ? '⏳ 正在全表批次比對中...'
                  : `⚡ 開始全活頁簿智慧比對（共 ${selectedSheets.length} 個工作表）`}
              </button>
            )}
          </div>

          {selectedSheets.length > 0 && productDB.length > 0 && (
            <div className="bg-white p-3 rounded-xl border border-emerald-100 flex flex-col sm:flex-row items-stretch sm:items-center justify-between gap-4">
              <div className="flex-1">
                <div className="flex justify-between text-xs font-bold text-slate-500 mb-1">
                  <span>🎯 模糊比對信心度及格門檻設定</span>
                  <span className="text-emerald-600 text-sm font-black">
                    {minScoreThreshold} 分
                  </span>
                </div>
                <input
                  type="range"
                  min="0"
                  max="100"
                  value={minScoreThreshold}
                  onChange={(e) =>
                    setMinScoreThreshold(parseInt(e.target.value))
                  }
                  className="w-full accent-emerald-600 cursor-pointer h-2 bg-slate-200 rounded-lg"
                />
              </div>
              <div className="flex items-center gap-2">
                <span className="text-xs text-slate-400 font-medium">
                  手動輸入：
                </span>
                <input
                  type="number"
                  min="0"
                  max="100"
                  value={minScoreThreshold}
                  onChange={(e) =>
                    setMinScoreThreshold(
                      Math.max(0, Math.min(100, parseInt(e.target.value) || 0))
                    )
                  }
                  className="w-16 text-center text-xs font-bold border border-slate-300 rounded p-1 text-emerald-600"
                />
                <span className="text-xs text-slate-400">分</span>
              </div>
            </div>
          )}

          {selectedSheets.length > 0 && productDB.length > 0 && (
            <div className="bg-white p-3 rounded-xl border border-purple-100 space-y-2">
              <div className="flex items-center justify-between gap-3 flex-wrap">
                <label className="flex items-center gap-2 text-xs font-bold text-purple-700 cursor-pointer">
                  <input
                    type="checkbox"
                    checked={ollamaEnabled}
                    onChange={(e) => setOllamaEnabled(e.target.checked)}
                    className="accent-purple-600 w-4 h-4"
                  />
                  🤖 啟用地端 AI（Ollama）跨工作表語意複核
                </label>
                {aiProgress && (
                  <span className="text-[11px] font-bold text-purple-600">
                    複核進度：{aiProgress.done}/{aiProgress.total}
                  </span>
                )}
              </div>
              {ollamaEnabled && (
                <div className="flex flex-col sm:flex-row gap-3 text-[11px]">
                  <div className="flex-1">
                    <label className="block text-slate-400 font-bold mb-0.5">
                      模型名稱
                    </label>
                    <input
                      type="text"
                      value={ollamaModel}
                      onChange={(e) => setOllamaModel(e.target.value)}
                      className="w-full border rounded p-1 bg-slate-50"
                    />
                  </div>
                  <div className="flex-1">
                    <label className="block text-slate-400 font-bold mb-0.5">
                      API 位址
                    </label>
                    <input
                      type="text"
                      value={ollamaUrl}
                      onChange={(e) => setOllamaUrl(e.target.value)}
                      className="w-full border rounded p-1 bg-slate-50"
                    />
                  </div>
                  <div className="flex-1">
                    <label className="block text-slate-400 font-bold mb-0.5">
                      API Key
                    </label>
                    <input
                      type="password"
                      value={ollamaApiKey}
                      onChange={(e) => setOllamaApiKey(e.target.value)}
                      className="w-full border rounded p-1 bg-slate-50"
                    />
                  </div>
                </div>
              )}
              {ollamaEnabled && (
                <div className="flex items-center gap-3 pt-1">
                  <button
                    type="button"
                    onClick={handleTestOllamaConnection}
                    disabled={isTestingOllama}
                    className="text-[11px] font-bold bg-purple-600 hover:bg-purple-700 disabled:opacity-50 text-white px-3 py-1.5 rounded-lg transition"
                  >
                    {isTestingOllama ? '⏳ 測試中...' : '🔌 測試 Ollama 連線'}
                  </button>
                  {ollamaTestStatus && (
                    <span className="text-[11px] font-medium text-slate-600">
                      {ollamaTestStatus}
                    </span>
                  )}
                </div>
              )}
            </div>
          )}
        </div>

        {/* 步驟 5：多工作表頁籤人工覆核面板 */}
        {hasAnyResults && (
          <div className="border rounded-xl p-4 space-y-4">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b pb-3">
              <div>
                <h3 className="font-bold text-sm text-slate-700">
                  🔍 步驟 5：多工作表人工覆核面板：
                </h3>
                <p className="text-xs text-slate-400">
                  點擊下方工作表頁籤切換檢視與微調，最後按右側按鈕一鍵匯出整本活頁簿！
                </p>
              </div>
              <button
                onClick={handleExportFile}
                disabled={isProcessing}
                className="bg-blue-600 hover:bg-blue-700 disabled:opacity-50 text-white font-extrabold py-2 px-6 rounded-lg shadow text-sm transition transform active:scale-95"
              >
                {isProcessing
                  ? '⏳ 正在製表匯出中...'
                  : '💾 匯出全活頁簿智慧回填標單.xlsx'}
              </button>
            </div>

            {/* 📑 工作表頁籤導覽列 */}
            <div className="flex items-center gap-2 overflow-x-auto pb-1 border-b border-slate-200">
              {allSheetNames.map((sName) => {
                const results = allSheetResults[sName];
                if (!results) return null;
                const total = results.length;
                const yellowCount = results.filter(
                  (r) => r.status === 'YELLOW' || r.status === 'AMBIGUOUS'
                ).length;
                const isCurrent = activeSheetTab === sName;

                return (
                  <button
                    key={`review-tab-${sName}`}
                    type="button"
                    onClick={() => setActiveSheetTab(sName)}
                    className={`px-3.5 py-2 rounded-t-lg text-xs font-bold transition flex items-center gap-1.5 flex-shrink-0 ${
                      isCurrent
                        ? 'bg-blue-600 text-white shadow-sm'
                        : 'bg-slate-100 hover:bg-slate-200 text-slate-700'
                    }`}
                  >
                    <span>{sName}</span>
                    <span
                      className={`text-[10px] px-1.5 py-0.2 rounded-full font-bold ${
                        isCurrent
                          ? 'bg-white/20 text-white'
                          : 'bg-slate-200 text-slate-600'
                      }`}
                    >
                      {total}筆
                    </span>
                    {yellowCount > 0 ? (
                      <span
                        className={`text-[9px] px-1.5 py-0.2 rounded font-black ${
                          isCurrent
                            ? 'bg-amber-400 text-amber-950'
                            : 'bg-amber-100 text-amber-800'
                        }`}
                      >
                        🟡 {yellowCount} 待查
                      </span>
                    ) : (
                      <span
                        className={`text-[9px] px-1.5 py-0.2 rounded font-black ${
                          isCurrent
                            ? 'bg-emerald-400 text-emerald-950'
                            : 'bg-emerald-100 text-emerald-800'
                        }`}
                      >
                        🟢 準確
                      </span>
                    )}
                  </button>
                );
              })}
            </div>

            {/* 當前分頁表格 */}
            {activeResults.length > 0 ? (
              <div className="overflow-x-auto rounded-lg border">
                <table className="w-full text-xs text-left border-collapse">
                  <thead>
                    <tr className="bg-slate-100 border-b text-slate-600 font-bold">
                      <th className="p-2.5 w-12 text-center">列號</th>
                      <th className="p-2.5 w-1/3">客戶標單品名</th>
                      <th className="p-2.5 w-5/12">💡 匹配雲端牌價規格</th>
                      <th className="p-2.5 w-1/5">🔍 強行搜尋全表</th>
                      <th className="p-2 text-right">原始牌價</th>
                      <th className="p-2 text-right text-emerald-600 font-bold">
                        加折後單價
                      </th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-200">
                    {activeResults.map((res: MatchResult, idx: number) => {
                      const multiplier =
                        discounts[res.itemTag] !== undefined
                          ? discounts[res.itemTag]
                          : discounts['otherProductDiscount'];
                      const finalCalculatedPrice = Math.round(
                        res.listPrice * multiplier
                      );
                      const searchTxt = res.manualSearchQuery
                        .trim()
                        .toLowerCase();
                      const showGlobalSearch =
                        searchTxt.length > 0 && !res.isComposing;
                      const filteredDropdownItems = showGlobalSearch
                        ? productDB
                            .filter(
                              (db: DBItem) =>
                                db.description
                                  .toLowerCase()
                                  .includes(searchTxt) ||
                                db.keywords.toLowerCase().includes(searchTxt) ||
                                db.item_tag.toLowerCase().includes(searchTxt)
                            )
                            .slice(0, 45)
                        : res.candidates;

                      return (
                        <tr
                          key={`match-row-${activeSheetTab}-${res.rowNumber}-${idx}`}
                          className={`hover:bg-slate-50 ${
                            res.status === 'MEMORY' ? 'bg-blue-50/20' : ''
                          }`}
                        >
                          <td className="p-2.5 text-center font-mono text-slate-400 bg-slate-50">
                            {res.rowNumber}
                          </td>
                          <td className="p-2.5 font-medium text-slate-900">
                            {res.clientName}
                          </td>
                          <td className="p-2.5 text-slate-600">
                            {res.aiSuggested && (
                              <span className="inline-block mr-1 rounded px-1.5 py-0.5 text-[9px] font-bold bg-purple-100 text-purple-700">
                                🤖 AI建議
                              </span>
                            )}
                            {res.inferredContextTag && (
                              <span className="inline-block mr-1 rounded px-1.5 py-0.5 text-[9px] font-bold bg-teal-50 text-teal-700 border border-teal-200">
                                {res.inferredContextTag}
                              </span>
                            )}
                            <span
                              className={`inline-block mr-2 rounded px-1.5 py-0.5 text-[9px] font-bold ${
                                res.patternMatched
                                  ? 'bg-indigo-600 text-white shadow-sm'
                                  : res.status === 'MEMORY'
                                  ? 'bg-blue-600 text-white shadow-sm'
                                  : res.status === 'GREEN'
                                  ? 'bg-green-100 text-green-800'
                                  : res.status === 'AMBIGUOUS'
                                  ? 'bg-orange-100 text-orange-800'
                                  : res.status === 'YELLOW'
                                  ? 'bg-yellow-100 text-yellow-800'
                                  : 'bg-red-100 text-red-800'
                              }`}
                            >
                              {res.patternMatched
                                ? '🧠⚡ 系列泛化命中'
                                : res.status === 'MEMORY'
                                ? '🧠 大腦記憶命中'
                                : res.status === 'GREEN'
                                ? '🟢 準確'
                                : res.status === 'AMBIGUOUS'
                                ? '⚠️ 資訊不足'
                                : res.status === 'YELLOW'
                                ? '🟡 待查'
                                : '🔴 排除'}
                            </span>
                            <select
                              value={res.matchedId}
                              onChange={(e) =>
                                handleCandidateChange(idx, e.target.value)
                              }
                              className="text-xs bg-white border rounded p-0.5 max-w-[340px] font-medium text-slate-700"
                            >
                              {res.matchedId !== 'REMOVE_MAPPING' && (
                                <option value={res.matchedId}>
                                  🏆 ［{res.itemTag}］ {res.matchedDescription}{' '}
                                  (${res.listPrice}元)
                                </option>
                              )}
                              {showGlobalSearch && (
                                <option value="REMOVE_MAPPING">
                                  --- 🔍 全表強行過濾結果 (
                                  {filteredDropdownItems.length}筆) ---
                                </option>
                              )}
                              {filteredDropdownItems.map(
                                (cand: DBItem, cIdx: number) => {
                                  if (cand.id === res.matchedId) return null;
                                  return (
                                    <option
                                      key={`cand-${res.rowNumber}-${cIdx}`}
                                      value={cand.id}
                                    >
                                      {showGlobalSearch ? '🔍' : '🥈'} ［
                                      {cand.item_tag}］ {cand.description} ($
                                      {cand.list_price}元)
                                    </option>
                                  );
                                }
                              )}
                              <option value="REMOVE_MAPPING">
                                ❌ 此品項不套用牌價 /
                                屬於工資或排除項（保持空白）
                              </option>
                            </select>
                          </td>
                          <td className="p-2">
                            <input
                              type="text"
                              placeholder="選單沒有？在此打字強搜全表關鍵字..."
                              value={res.manualSearchQuery}
                              onCompositionStart={() =>
                                handleCompositionStart(idx)
                              }
                              onCompositionEnd={(e) =>
                                handleCompositionEnd(
                                  idx,
                                  (e.target as HTMLInputElement).value
                                )
                              }
                              onChange={(e) =>
                                handleSearchTextChange(idx, e.target.value)
                              }
                              className="w-full text-[11px] border border-slate-300 rounded px-2 py-0.5 focus:border-blue-500 focus:outline-none bg-slate-50 focus:bg-white text-indigo-700 font-medium"
                            />
                          </td>
                          <td className="p-2 text-right font-mono text-slate-400">
                            ${res.listPrice}
                          </td>
                          <td className="p-2 text-right font-mono font-black text-emerald-600 bg-emerald-50/10">
                            {res.listPrice > 0
                              ? `$${finalCalculatedPrice.toLocaleString()}`
                              : '-'}
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            ) : (
              <div className="p-8 text-center bg-slate-50 rounded-lg border border-dashed text-slate-400 text-xs">
                此工作表「{activeSheetTab}
                」尚未進行比對或無相符項目。請確認上方步驟 2
                已勾選此表並點擊「開始全活頁簿智慧比對」。
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
}
