import React, { useState, useEffect, useMemo } from 'react';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { supabase } from './supabaseClient';

// ==========================================
// 1. 進階規格標準化：內建水電「公稱直徑尺寸對照字典」
// ==========================================
function normalizeSpec(text: string): { cleanText: string; sizeTag: string } {
  if (!text) return { cleanText: '', sizeTag: '' };

  let clean = text.toLowerCase().replace(/[\s\t\n\r]/g, '');
  let sizeTag = '';

  // 【A. 電線電纜截面積標準化】(mm2, m㎡, SQ, 平方 ➡️ sq)
  const sqRegex = /(\d+(?:\.\d+)?)\s*(?:mm2|m㎡|sq|平方(?:公厘)?)/i;
  const sqMatch = clean.match(sqRegex);
  if (sqMatch) {
    return { cleanText: clean, sizeTag: `${sqMatch[1]}sq` };
  }

  // 【B. 水電專用：公稱管徑 mm / A 轉 Inch 智慧對照字典】
  if (/\b(16mm|16a|15a|1\/2)\b/.test(clean) || clean.includes('1/2'))
    sizeTag = '0.5inch';
  else if (/\b(22mm|20a|3\/4)\b/.test(clean) || clean.includes('3/4'))
    sizeTag = '0.75inch';
  else if (/\b(28mm|25a|1吋|1")\b/.test(clean)) sizeTag = '1inch';
  else if (/\b(35mm|32a|1-1\/4|1\.25)\b/.test(clean)) sizeTag = '1.25inch';
  else if (/\b(42mm|40a|1-1\/2|1\.5)\b/.test(clean) || clean.includes('吋半'))
    sizeTag = '1.5inch';
  else if (/\b(52mm|50a|2吋|2")\b/.test(clean)) sizeTag = '2inch';
  else if (/\b(65mm|65a|2-1\/2|2\.5)\b/.test(clean)) sizeTag = '2.5inch';
  else if (/\b(77mm|80a|3吋|3")\b/.test(clean)) sizeTag = '3inch';
  else if (/\b(100mm|100a|4吋|4")\b/.test(clean)) sizeTag = '4inch';

  // 【C. 常規分數/吋標籤 Regex 補漏機制 (由大到小排序，防範子字串誤觸)】
  if (!sizeTag) {
    let fractionText = clean;
    if (fractionText.includes('1/2'))
      fractionText = fractionText.replace(/1[-_]1\/2/g, '1.5');
    if (fractionText.includes('3/4'))
      fractionText = fractionText.replace(/1[-_]3\/4/g, '1.75');
    if (fractionText.includes('1/4'))
      fractionText = fractionText.replace(/1[-_]1\/4/g, '1.25');

    const inchRegex = /(\d+(?:\.\d+)?)\s*(?:"|吋|φ|ф|inch)/i;
    const inchMatch = fractionText.match(inchRegex);
    if (inchMatch) {
      sizeTag = `${inchMatch[1]}inch`;
    }
  }

  return { cleanText: clean, sizeTag };
}

// ==========================================
// 1b. 補強：description 欄位缺單位時，從 keywords 備援抓取尺寸標籤
//     （實測發現電纜/電線類 description 只填裸數字如「5.5」，
//      單位資訊實際上藏在 keywords 裡，例如「5.5mm2」）
// ==========================================
function resolveDbSizeTag(dbRow: DBItem): string {
  const fromDescription = normalizeSpec(dbRow.description).sizeTag;
  if (fromDescription) return fromDescription;
  // 備援：把 keywords 逐項拆開找看看有沒有帶單位的規格字樣
  const kwItems = dbRow.keywords.split(/[，,]/);
  for (const kw of kwItems) {
    const tag = normalizeSpec(kw).sizeTag;
    if (tag) return tag;
  }
  return '';
}

// ==========================================
// 1c. 新增規格維度：電纜/電線的「芯數」(2C / 3C / 4C 等)
//     現有邏輯完全沒有處理這個維度，只能靠鬆散關鍵字子字串比對，
//     容易因為客戶寫法不同（3/C、3C、3芯、3心）而抓錯芯數、比到錯誤單價
// ==========================================
function extractCoreCountTag(text: string): string {
  if (!text) return '';
  const t = text.toLowerCase().replace(/[\s\t\n\r]/g, '');
  const m = t.match(/(\d+)\s*(?:\/?c\b|芯|心)/i);
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
// 2. 資料結構與介面定義（標準英文欄位）
// ==========================================
interface DBItem {
  id: number;
  item_tag: string;
  description: string;
  list_price: number;
  keywords: string;
}

interface CloudRawItem {
  id: number;
  item_tag: string;
  description: string;
  list_price: any;
  keywords: string;
}

interface MemoryPayload {
  description: string;
  item_tag: string;
}

interface MatchResult {
  rowNumber: number;
  clientName: string;
  matchedId: number | string;
  matchedDescription: string;
  listPrice: number;
  itemTag: string;
  confidenceScore: number;
  status: 'GREEN' | 'YELLOW' | 'RED' | 'MEMORY';
  candidates: DBItem[];
  manualSearchQuery: string;
  isComposing: boolean; // 💡 引入組字鎖定狀態，完美阻絕注音輸入法衝突
  aiSuggested?: boolean; // 🤖 標記此結果是否由地端 AI（Ollama）語意複核給出
}

interface OllamaMatchResponse {
  selectedIndex: number | null;
  confidence: number;
  reason: string;
}

interface DetectedInfo {
  sheetName: string;
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
// 3. 地端 AI（Ollama）語意複核模組
// ==========================================
const AI_CONCURRENCY = 4; // 同時發送給 Ollama 的最大並發請求數
const AI_CONFIDENCE_THRESHOLD = 60; // AI 判斷信心低於此分數則不採用

// 🔀 自動判斷端點格式：路徑含 /v1 視為 OpenAI 相容 API（LiteLLM / vLLM / Ollama OpenAI 相容模式），
// 否則視為 Ollama 原生 /api/chat 格式
function normalizeOllamaEndpoint(rawUrl: string): {
  endpoint: string;
  isOpenAICompatible: boolean;
} {
  const trimmed = rawUrl.trim().replace(/\/+$/, '');
  if (/\/v1\/chat\/completions$/i.test(trimmed)) {
    return { endpoint: trimmed, isOpenAICompatible: true };
  }
  if (/\/v1$/i.test(trimmed)) {
    return { endpoint: `${trimmed}/chat/completions`, isOpenAICompatible: true };
  }
  return { endpoint: trimmed, isOpenAICompatible: false };
}

// 🧹 容錯 JSON 解析：有些模型不管有沒有要求 json_object，仍習慣把回應包在
// ```json ... ``` code fence 裡，這裡自動剝除後再解析，剝不出來則退而求其次抓第一個 {...} 區塊
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

// 🤖 呼叫地端 Ollama，讓 AI 從既有候選清單中挑出最符合客戶品名的一筆
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

候選清單：
${candidateList}

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

    if (!res.ok) throw new Error(`Ollama HTTP ${res.status}（端點：${endpoint}）`);

    const data = await res.json();
    const rawContent = isOpenAICompatible
      ? data?.choices?.[0]?.message?.content
      : data?.message?.content;
    const parsed = extractJsonFromModelOutput(rawContent);

    const rawIndex = parsed.selectedIndex;
    const selectedIndex =
      typeof rawIndex === 'number' && rawIndex >= 1 && rawIndex <= candidates.length
        ? rawIndex - 1
        : null;

    return {
      selectedIndex,
      confidence: typeof parsed.confidence === 'number' ? parsed.confidence : 0,
      reason: typeof parsed.reason === 'string' ? parsed.reason : '',
    };
  } catch (err) {
    console.error('🤖 Ollama 地端 AI 呼叫失敗', err);
    return null; // 失敗就直接 fallback，不影響原本規則比對結果
  }
}

export default function App() {
  const [clientFile, setClientFile] = useState<File | null>(null);
  const [detected, setDetected] = useState<DetectedInfo | null>(null);
  const [productDB, setProductDB] = useState<DBItem[]>([]);
  const [discounts, setDiscounts] = useState<{ [key: string]: number }>({
    otherProductDiscount: 1.0,
  });
  const [matchResults, setMatchResults] = useState<MatchResult[]>([]);
  const [statusMessage, setStatusMessage] = useState<string>(
    '📡正在連線至 Supabase 雲端大腦庫...'
  );
  const [isProcessing, setIsProcessing] = useState<boolean>(false);
  const [minScoreThreshold, setMinScoreThreshold] = useState<number>(45);

  // 🤖 地端 AI（Ollama）語意複核相關設定
  const [ollamaEnabled, setOllamaEnabled] = useState<boolean>(true);
  const [ollamaModel, setOllamaModel] = useState<string>('qwen2.5:7b');
  const [ollamaUrl, setOllamaUrl] = useState<string>(
    'http://localhost:11434/api/chat'
  );
  const [ollamaApiKey, setOllamaApiKey] = useState<string>('');
  const [aiProgress, setAiProgress] = useState<{ done: number; total: number } | null>(
    null
  );
  const [ollamaTestStatus, setOllamaTestStatus] = useState<string>('');
  const [isTestingOllama, setIsTestingOllama] = useState<boolean>(false);

  // 防禦性型別轉換轉換器
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
      };
    });
  };

  // 🔄 自動從 Supabase 同步英文欄位牌價表
  useEffect(() => {
    async function loadCloudDB() {
      try {
        const { data, error } = await supabase
          .from('price_db')
          .select('id, item_tag, description, list_price, keywords');

        if (error) throw error;

        if (data && data.length > 0) {
          const standardizedDB = transformRawData(
            data as unknown as CloudRawItem[]
          );

          const categories = new Set<string>();
          standardizedDB.forEach((item: DBItem) =>
            categories.add(item.item_tag)
          );

          const initialDiscounts: { [key: string]: number } = {
            otherProductDiscount: 1.0,
          };
          categories.forEach((cat: string) => {
            initialDiscounts[cat] = 0.85;
          });

          setDiscounts(initialDiscounts);
          setProductDB(standardizedDB);
          setStatusMessage(
            `💡 雲端連線成功！已自動載入最新牌價庫（共 ${standardizedDB.length} 筆）。請上傳客戶標單！`
          );
        } else {
          setStatusMessage(
            '⚠️ 雲端 price_db 資料表為空，請先至 Supabase 後台匯入 CSV。'
          );
        }
      } catch (err: any) {
        setStatusMessage(`❌ 雲端大腦連線失敗: ${err.message}。`);
      }
    }
    loadCloudDB();
  }, []);

  // 讀取客戶標單 (.xls / .xlsx 雙通殺機制)
  const handleClientFileChange = async (
    e: React.ChangeEvent<HTMLInputElement>
  ) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setClientFile(file);
    setMatchResults([]);
    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: 'array' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rawRows = XLSX.utils.sheet_to_json<any[]>(sheet, {
        header: 1,
        defval: '',
      });

      let headerRow = 1,
        nameCol = 2,
        priceCol = 5,
        qtyCol = 4,
        unitCol = 3,
        maxCols = 6;

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
          break;
        }
      }

      const previewRows = rawRows
        .slice(0, 12)
        .map((row: any[]) =>
          Array.from(
            { length: Math.max(maxCols, row.length) },
            (_, c: number) => String(row[c] || '')
          )
        );

      setDetected({
        sheetName: workbook.SheetNames[0],
        headerRow,
        nameCol,
        priceCol,
        qtyCol,
        unitCol,
        rawRows,
        previewRows,
        colCount: maxCols,
      });
      setStatusMessage(
        '✅ 標單載入成功！常駐校對面板已就緒，請點擊「開始智慧安全比對」。'
      );
    } catch (err: any) {
      setStatusMessage(`❌ 標單載入失敗: ${err.message}`);
    }
  };

  // 執行智慧演算法比對（加強版 Keywords 特徵陣列加權與動態紅利加分）
  const handleStartMatching = async () => {
    if (!detected || productDB.length === 0) return;
    setStatusMessage(
      '⚡ 正在同步調取雲端與地端記憶，執行去骨架化與 Keywords 深度陣列特徵加權比對...'
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
      } catch (e) {
        console.error('讀取本地記憶庫失敗', e);
      }

      const results: MatchResult[] = [];
      const startRowIdx = detected.headerRow;
      const totalRows = detected.rawRows.length;

      let i = startRowIdx;
      while (i < totalRows) {
        const row = detected.rawRows[i];
        if (!row) {
          i++;
          continue;
        }

        let clientName = String(row[detected.nameCol - 1] || '').trim();
        const clientQty = String(row[detected.qtyCol - 1] || '').trim();
        const clientUnit = String(row[detected.unitCol - 1] || '').trim();
        const parsedQty = parseFloat(clientQty.replace(/,/g, ''));

        // 👮 硬性過濾護城河：只有前面數量欄有數字、單位不為空的格子，才判定為主行
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

        // 多行品名智慧延伸拼接邏輯（Look-ahead 機制）
        let nextRowIdx = i + 1;
        while (nextRowIdx < totalRows) {
          const nextRow = detected.rawRows[nextRowIdx];
          if (!nextRow) break;
          const nextName = String(nextRow[detected.nameCol - 1] || '').trim();
          const nextQty = String(nextRow[detected.qtyCol - 1] || '').trim();
          const nextUnit = String(nextRow[detected.unitCol - 1] || '').trim();

          if (
            nextName !== '' &&
            (nextQty === '' || parseFloat(nextQty.replace(/,/g, '')) === 0) &&
            nextUnit === ''
          ) {
            clientName += ' ' + nextName;
            nextRowIdx++;
          } else {
            break;
          }
        }

        const clientNorm = normalizeSpec(clientName);
        const lookupKey = clientNorm.cleanText;
        const clientCoreTag = extractCoreCountTag(clientNorm.cleanText);

        // 負向排除黑名單：工資或施工項目直接強制歸零
        if (BLACKLIST_KEYWORDS.test(clientName)) {
          i = nextRowIdx;
          continue;
        }

        // 🧠 歷史記憶大腦攔截
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

          const dbFound = productDB.find((d: DBItem) => {
            if (forcedTag) {
              return d.description === targetDesc && d.item_tag === forcedTag;
            }
            return d.description === targetDesc;
          });

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
            });
            i = nextRowIdx;
            continue;
          }
        }

        // 📊 💡 優化演算法：D 欄 Keywords 特徵拆解交叉計分法與核心權重補償
        const scoredItems: { dbRow: DBItem; score: number }[] = [];

        for (const dbRow of productDB) {
          const dbNorm = normalizeSpec(dbRow.description);
          const dbSizeTag = dbNorm.sizeTag || resolveDbSizeTag(dbRow);

          // 📐 尺寸防禦網優化：改為動態加扣分制，取代極端一刀切淘汰
          let sizeScoreBonus = 0;
          if (clientNorm.sizeTag && dbSizeTag) {
            if (clientNorm.sizeTag === dbSizeTag) {
              sizeScoreBonus = 40; // 🎯 尺寸完全吻合額外大加 40 分
            } else {
              sizeScoreBonus = -50; // ❌ 尺寸完全不符重扣 50 分
            }
          }

          // 🧵 芯數比對（2C/3C/4C 等）：電纜/電線類的第二個獨立規格維度，
          // 同一 mm2 底下不同芯數是完全不同的品項與單價
          let coreCountBonus = 0;
          const dbCoreTag = resolveDbCoreCountTag(dbRow);
          if (clientCoreTag && dbCoreTag) {
            if (clientCoreTag === dbCoreTag) {
              coreCountBonus = 35; // 🎯 芯數完全吻合加分
            } else {
              coreCountBonus = -45; // ❌ 芯數不符重扣（避免比到錯誤芯數但同尺寸的品項）
            }
          }

          // 將 說明欄(keywords) 徹底拆解成特徵陣列
          const kwArray = dbRow.keywords
            .split(/[，,]/)
            .map((k: string) => k.trim().toLowerCase())
            .filter((k: string) => k);
          let matchCount = 0;
          let coreFeatureBonus = 0; // 🌟 黃金核心特徵紅利字眼

          kwArray.forEach((kw: string) => {
            const normKw = normalizeSpec(kw).cleanText;
            if (normKw && clientNorm.cleanText.indexOf(normKw) !== -1) {
              matchCount++;
              // 針對管材顏色及用途關鍵字給予強力分流紅利 (+20)
              if (
                /(橘色|排水|廢水|給水|灰色|e管|b管|絕緣|電力電纜|耐熱)/.test(
                  normKw
                )
              ) {
                coreFeatureBonus += 20;
              }
            }
          });

          // 核心交叉加權計分
          let keywordBaseScore = 0;
          if (kwArray.length > 0) {
            keywordBaseScore = Math.round((matchCount / kwArray.length) * 40); // 基礎關鍵字比例佔40分
          }

          // 中文品名描述重疊度獎勵
          let textOverlapBonus = 0;
          if (
            clientNorm.cleanText.includes(dbNorm.cleanText) ||
            dbNorm.cleanText.includes(clientNorm.cleanText)
          ) {
            textOverlapBonus = 20;
          }

          // 總體評分模型合成
          const totalScore =
            keywordBaseScore +
            sizeScoreBonus +
            coreCountBonus +
            coreFeatureBonus +
            textOverlapBonus;

          scoredItems.push({
            dbRow,
            score: Math.max(0, Math.min(100, totalScore)),
          });
        }

        scoredItems.sort((a, b) => b.score - a.score);
        const topCandidates = scoredItems
          .slice(0, 4)
          .map((x: { dbRow: DBItem; score: number }) => x.dbRow);
        const bestMatch = scoredItems[0];

        let status: 'GREEN' | 'YELLOW' | 'RED' = 'RED';
        let finalDescription = `未找到相符規格（匹配得分低於 ${minScoreThreshold} 分已自動過濾）`;
        let finalPrice = 0;
        let finalTag = 'otherProductDiscount';
        let finalId: number | string = 'REMOVE_MAPPING';

        if (bestMatch && bestMatch.score >= minScoreThreshold) {
          finalId = bestMatch.dbRow.id;
          finalDescription = bestMatch.dbRow.description;
          finalPrice = bestMatch.dbRow.list_price;
          finalTag = bestMatch.dbRow.item_tag;
          status = bestMatch.score >= 80 ? 'GREEN' : 'YELLOW';
        }

        results.push({
          rowNumber: i + 1,
          clientName,
          matchedId: finalId,
          matchedDescription: finalDescription,
          listPrice: finalPrice,
          itemTag: finalTag,
          confidenceScore: bestMatch ? bestMatch.score : 0,
          status,
          candidates: topCandidates,
          manualSearchQuery: '',
          isComposing: false,
        });

        i = nextRowIdx;
      }

      setMatchResults(results);
      const greenCount = results.filter(
        (r: MatchResult) => r.status === 'GREEN' || r.status === 'MEMORY'
      ).length;
      setStatusMessage(
        `🎉 比對成功！本次門檻：${minScoreThreshold} 分。(🟢 高信心：${greenCount} 筆，🧠 智慧記憶/待確認：${
          results.length - greenCount
        } 筆)`
      );

      // 🤖 對信心不足（YELLOW / RED）的品項，呼叫地端 Ollama 進行語意複核
      if (ollamaEnabled) {
        await runAiRecheck(results);
      }
    } catch (e: any) {
      setStatusMessage(`❌ 比對失敗: ${e.message}`);
    }
  };

  // 🤖 地端 AI 批次複核：只處理低信心（YELLOW / RED）且有候選清單的品項
  const runAiRecheck = async (currentResults: MatchResult[]) => {
    const targets = currentResults
      .map((r, idx) => ({ r, idx }))
      .filter(
        ({ r }) =>
          (r.status === 'YELLOW' || r.status === 'RED') &&
          r.candidates.length > 0
      );

    if (targets.length === 0) {
      setStatusMessage(
        `ℹ️ 本次比對沒有 🟡待查 / 🔴排除 的品項需要 AI 複核（可能全部命中 🟢/🧠，或候選清單為空）。`
      );
      return;
    }

    console.log(
      `🤖 [AI複核] 開始處理 ${targets.length} 筆，模型：${ollamaModel}，位址：${ollamaUrl}`
    );

    setAiProgress({ done: 0, total: targets.length });
    setStatusMessage(
      `🤖 正在啟動地端 AI（${ollamaModel}）語意複核... (0/${targets.length})`
    );

    let doneCount = 0;
    let failCount = 0;
    let cursor = 0;

    const worker = async () => {
      while (cursor < targets.length) {
        const myTurn = cursor;
        cursor++;
        const { r, idx } = targets[myTurn];

        const originalMatchedId = r.matchedId; // 用來偵測使用者是否已手動覆核過
        const aiResult = await askOllamaForBestMatch(
          r.clientName,
          r.candidates,
          ollamaModel,
          ollamaUrl,
          ollamaApiKey
        );

        if (!aiResult) failCount++;
        console.log(`🤖 [AI複核] 「${r.clientName}」→`, aiResult);

        doneCount++;
        setAiProgress({ done: doneCount, total: targets.length });
        setStatusMessage(
          `🤖 地端 AI 語意複核中... (${doneCount}/${targets.length})${
            failCount > 0 ? `，⚠️ 已有 ${failCount} 筆呼叫失敗` : ''
          }`
        );

        if (
          aiResult &&
          aiResult.selectedIndex !== null &&
          aiResult.confidence >= AI_CONFIDENCE_THRESHOLD
        ) {
          const picked = r.candidates[aiResult.selectedIndex];
          if (picked) {
            setMatchResults((prev) => {
              const next = [...prev];
              const target = next[idx];
              // 若使用者在 AI 處理期間已手動覆核過此列，則不覆蓋使用者的選擇
              if (!target || target.matchedId !== originalMatchedId) return next;
              next[idx] = {
                ...target,
                matchedId: picked.id,
                matchedDescription: picked.description,
                listPrice: picked.list_price,
                itemTag: picked.item_tag,
                confidenceScore: aiResult.confidence,
                status: 'YELLOW', // 維持 YELLOW，保留人工最終確認
                aiSuggested: true,
              };
              return next;
            });
          }
        }
      }
    };

    await Promise.all(
      Array.from({ length: Math.min(AI_CONCURRENCY, targets.length) }, () =>
        worker()
      )
    );

    setAiProgress(null);
    setStatusMessage(
      failCount > 0
        ? `⚠️ 地端 AI 複核完成，但 ${failCount}/${targets.length} 筆呼叫失敗（請開瀏覽器 Console 查看錯誤，常見原因為 CORS 未開啟或 Ollama 連不上）。`
        : `🎉 地端 AI 複核完成！共處理 ${targets.length} 筆待確認品項，請於下方人工覆核面板確認 🤖 標籤的建議。`
    );
  };

  // 下拉式選單手動覆核切換
  const handleCandidateChange = async (
    index: number,
    selectedIdStr: string
  ) => {
    const nextResults = [...matchResults];
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
      setMatchResults(nextResults);

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
        setMatchResults(nextResults);

        // 🔒 大腦記憶隔離：同時記錄材料標準描述與大類別標籤
        localMemory[cleanLookupKey] = {
          description: found.description,
          item_tag: found.item_tag,
        };
        localStorage.setItem(
          'optimized_learning_memory',
          JSON.stringify(localMemory)
        );

        await supabase.from('learning_memory').upsert({
          client_name: target.clientName.trim(),
          matched_description: found.description,
        });
      }
    }
  };

  // 💡 注音輸入法防禦核心：監聽鍵盤選字狀態
  const handleSearchTextChange = (index: number, val: string) => {
    setMatchResults((prev: MatchResult[]) => {
      const next = [...prev];
      next[index].manualSearchQuery = val;
      return next;
    });
  };

  const handleCompositionStart = (index: number) => {
    setMatchResults((prev: MatchResult[]) => {
      const next = [...prev];
      next[index].isComposing = true;
      return next;
    });
  };

  const handleCompositionEnd = (index: number, val: string) => {
    setMatchResults((prev: MatchResult[]) => {
      const next = [...prev];
      next[index].isComposing = false;
      next[index].manualSearchQuery = val;
      return next;
    });
  };

  // 最終標單數據寫回並下載
  const handleExportFile = async () => {
    if (!clientFile || matchResults.length === 0 || !detected) return;
    setIsProcessing(true);
    setStatusMessage('⏳ 正在複製原標單樣式並精密填入折後單價...');

    try {
      const buffer = await clientFile.arrayBuffer();
      const workbook = new ExcelJS.Workbook();
      const isOldXls = clientFile.name.endsWith('.xls');

      if (isOldXls) {
        const tempXlsxWorkbook = XLSX.read(buffer, { type: 'array' });
        const xlsxWriteBuffer = XLSX.write(tempXlsxWorkbook, {
          bookType: 'xlsx',
          type: 'array',
        });
        await workbook.xlsx.load(xlsxWriteBuffer);
      } else {
        await workbook.xlsx.load(buffer);
      }

      const worksheet =
        workbook.getWorksheet(detected.sheetName) || workbook.worksheets[0];

      matchResults.forEach((item: MatchResult) => {
        if (item.listPrice === 0 || item.matchedId === 'REMOVE_MAPPING') return;

        const multiplier =
          discounts[item.itemTag] !== undefined
            ? discounts[item.itemTag]
            : discounts['otherProductDiscount'];

        // 浮點數誤差防禦：強制四捨五入整數
        const validMultiplier = isNaN(multiplier) ? 1.0 : multiplier;
        const finalPrice = Math.round(item.listPrice * validMultiplier);

        const row = worksheet.getRow(item.rowNumber);
        const cell = row.getCell(detected.priceCol);

        cell.value = Number(finalPrice);
        cell.numFmt = '#,##0';
        cell.font = { color: { argb: 'FFFF0000' }, bold: true };
      });

      const writeBuffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([writeBuffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');

      const newFileName =
        clientFile.name.replace(/\.xls[x]?$/, '') + '_精準智慧回填.xlsx';
      a.href = url;
      a.download = newFileName;
      a.click();
      URL.revokeObjectURL(url);

      setStatusMessage(
        `🏆 標單回填成功！智慧大腦記憶已全面就緒。新檔名：[${newFileName}]`
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
        '⚠️ 確定要清空本機瀏覽器的智慧學習紀錄嗎？(雲端資料庫不會受影響)'
      )
    ) {
      localStorage.removeItem('optimized_learning_memory');
      alert('🧼 本地骨架快取記憶已重置清空。');
      if (matchResults.length > 0) handleStartMatching();
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
      if (result) {
        setOllamaTestStatus(
          `✅ 連線成功！模型「${ollamaModel}」回應正常（selectedIndex=${result.selectedIndex}, confidence=${result.confidence}）`
        );
      } else {
        setOllamaTestStatus(
          '❌ 連線失敗或回應解析失敗，請打開瀏覽器 Console（F12）查看詳細錯誤訊息。'
        );
      }
    } finally {
      setIsTestingOllama(false);
    }
  };

  const getColLetter = (index: number) => String.fromCharCode(64 + index);
  const discountKeys = useMemo(() => Object.keys(discounts), [discounts]);

  return (
    <div className="min-h-screen bg-slate-100 p-6 text-slate-800 font-sans">
      <div className="max-w-7xl mx-auto bg-white rounded-2xl shadow-xl p-6 space-y-6">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
          <div>
            <h1 className="text-2xl font-black text-emerald-600">
              🏢 水電標單雲端共用系統{' '}
              <span className="text-xs bg-indigo-600 text-white px-2 py-0.5 rounded ml-2">
                🧠 終極智配優化版
              </span>
            </h1>
            <p className="text-xs text-slate-400 mt-1">
              D 欄 Keywords
              特徵獨特性陣列比對加權，支援全表解鎖數據池與注音輸入法不卡字防護鎖。
            </p>
          </div>
          <button
            onClick={handleClearMemoryCache}
            className="text-[11px] self-start sm:self-center bg-slate-200 hover:bg-red-100 text-slate-500 hover:text-red-700 px-3 py-1.5 rounded-lg font-bold transition"
          >
            🧼 清空本地大腦快取
          </button>
        </div>

        {/* 標單上傳 */}
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
              🎯 當前處理標單：{clientFile.name}
            </p>
          )}
        </div>

        {/* 常駐欄位校對網格 */}
        {detected && (
          <div className="border border-amber-200 bg-amber-50/40 rounded-xl p-4 space-y-3">
            <h3 className="font-bold text-xs text-amber-900">
              ⚙️ 步驟 2：常駐欄位校對網格：
            </h3>
            <div className="grid grid-cols-2 sm:grid-cols-4 gap-4 text-xs">
              <div>
                <label className="block font-bold text-blue-700 mb-0.5">
                  📦 材料品名欄：
                </label>
                <select
                  value={detected.nameCol}
                  onChange={(e) =>
                    setDetected({
                      ...detected,
                      nameCol: parseInt(e.target.value),
                    })
                  }
                  className="w-full border rounded p-1 bg-white"
                >
                  {Array.from({ length: detected.colCount }).map(
                    (_, i: number) => (
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
                  value={detected.priceCol}
                  onChange={(e) =>
                    setDetected({
                      ...detected,
                      priceCol: parseInt(e.target.value),
                    })
                  }
                  className="w-full border rounded p-1 bg-white"
                >
                  {Array.from({ length: detected.colCount }).map((_, i) => (
                    <option key={`priceCol-${i}`} value={i + 1}>
                      {getColLetter(i + 1)} 欄
                    </option>
                  ))}
                </select>
              </div>
              <div>
                <label className="block font-bold text-slate-700 mb-0.5">
                  🔢 數量欄：
                </label>
                <select
                  value={detected.qtyCol}
                  onChange={(e) =>
                    setDetected({
                      ...detected,
                      qtyCol: parseInt(e.target.value),
                    })
                  }
                  className="w-full border rounded p-1 bg-white"
                >
                  {Array.from({ length: detected.colCount }).map((_, i) => (
                    <option key={`qtyCol-${i}`} value={i + 1}>
                      {getColLetter(i + 1)} 欄
                    </option>
                  ))}
                </select>
              </div>
              <div>
                <label className="block font-bold text-purple-700 mb-0.5">
                  📏 單位欄（過濾用）：
                </label>
                <select
                  value={detected.unitCol}
                  onChange={(e) =>
                    setDetected({
                      ...detected,
                      unitCol: parseInt(e.target.value),
                    })
                  }
                  className="w-full border rounded p-1 bg-white"
                >
                  {Array.from({ length: detected.colCount }).map((_, i) => (
                    <option key={`unitCol-${i}`} value={i + 1}>
                      {getColLetter(i + 1)} 欄
                    </option>
                  ))}
                </select>
              </div>
            </div>
          </div>
        )}

        {/* 執行區 */}
        <div className="p-4 rounded-xl border bg-emerald-50 text-emerald-900 border-emerald-200 space-y-4">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
            <span className="text-sm font-bold block">{statusMessage}</span>
            {detected && productDB.length > 0 && (
              <button
                onClick={handleStartMatching}
                className="w-full sm:w-auto bg-emerald-600 hover:bg-emerald-700 text-white font-black py-2 px-8 rounded-lg shadow-md transition transform active:scale-95"
              >
                ⚡ 開始智慧安全比對
              </button>
            )}
          </div>

          {detected && productDB.length > 0 && (
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

          {detected && productDB.length > 0 && (
            <div className="bg-white p-3 rounded-xl border border-purple-100 space-y-2">
              <div className="flex items-center justify-between gap-3 flex-wrap">
                <label className="flex items-center gap-2 text-xs font-bold text-purple-700 cursor-pointer">
                  <input
                    type="checkbox"
                    checked={ollamaEnabled}
                    onChange={(e) => setOllamaEnabled(e.target.checked)}
                    className="accent-purple-600 w-4 h-4"
                  />
                  🤖 啟用地端 AI（Ollama）語意複核（僅處理 🟡待查 / 🔴排除 品項）
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
                      placeholder="qwen2.5:7b"
                    />
                  </div>
                  <div className="flex-1">
                    <label className="block text-slate-400 font-bold mb-0.5">
                      API 位址（支援 Ollama 原生 /api/chat，或 OpenAI 相容 /v1 端點）
                    </label>
                    <input
                      type="text"
                      value={ollamaUrl}
                      onChange={(e) => setOllamaUrl(e.target.value)}
                      className="w-full border rounded p-1 bg-slate-50"
                      placeholder="http://localhost:11434/api/chat 或 https://xxx/v1"
                    />
                  </div>
                  <div className="flex-1">
                    <label className="block text-slate-400 font-bold mb-0.5">
                      API Key（無驗證則留空）
                    </label>
                    <input
                      type="password"
                      value={ollamaApiKey}
                      onChange={(e) => setOllamaApiKey(e.target.value)}
                      className="w-full border rounded p-1 bg-slate-50"
                      placeholder="例如 123"
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

        {/* 方案 A 折數面板 */}
        {matchResults.length > 0 && (
          <div className="border border-blue-200 rounded-xl p-4 bg-blue-50/30 space-y-3">
            <div className="flex justify-between items-center">
              <h3 className="font-extrabold text-sm text-blue-900">
                💰 步驟 3：產品分類計價折數 (倍率) 面板：
              </h3>
              <span className="text-xs text-blue-700 font-semibold">
                💡 提示：打 85 折輸入 0.85；溢價銷售請直接輸入 1.15
              </span>
            </div>
            <div className="grid grid-cols-2 sm:grid-cols-4 gap-4">
              {discountKeys.map((key: string) => (
                <div
                  key={`discount-box-${key}`}
                  className="bg-white p-2 rounded-lg border shadow-sm flex flex-col justify-between"
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
                        setDiscounts({
                          ...discounts,
                          [key]: parseFloat(e.target.value) || 0,
                        })
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

        {/* 📉 人工覆核清單面板 */}
        {matchResults.length > 0 && (
          <div className="border rounded-xl p-4 space-y-4">
            <div className="flex justify-between items-center border-b pb-3">
              <h3 className="font-bold text-sm text-slate-700">
                🔍 步驟 4：人工覆核面板：
              </h3>
              <button
                onClick={handleExportFile}
                disabled={isProcessing}
                className="bg-blue-600 hover:bg-blue-700 text-white font-extrabold py-2 px-6 rounded-lg shadow text-sm"
              >
                {isProcessing ? '⏳ 正在製表...' : '💾 匯出智慧回填標單.xlsx'}
              </button>
            </div>

            <div className="overflow-x-auto rounded-lg border">
              <table className="w-full text-xs text-left border-collapse">
                <thead>
                  <tr className="bg-slate-100 border-b text-slate-600 font-bold">
                    <th className="p-2.5 w-12 text-center">列號</th>
                    <th className="p-2.5 w-1/3">客戶標單品名</th>
                    <th className="p-2.5 w-5/12">
                      💡 匹配雲端牌價規格（已內嵌中文類別提示）
                    </th>
                    <th className="p-2.5 w-1/5">
                      🔍 找不到？輸入關鍵字強行搜尋全表（💡
                      已安裝注音不卡字防護鎖）
                    </th>
                    <th className="p-2 text-right">原始牌價</th>
                    <th className="p-2 text-right text-emerald-600 font-bold">
                      加折後單價
                    </th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200">
                  {matchResults.map((res: MatchResult, idx: number) => {
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

                    // 🧠 解除數據池封印！只要強行搜尋框「有字且非組字中」，立刻接管切換為全表解鎖搜尋
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
                        key={`match-row-${res.rowNumber}-${idx}`}
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
                          <span
                            className={`inline-block mr-2 rounded px-1.5 py-0.5 text-[9px] font-bold ${
                              res.status === 'MEMORY'
                                ? 'bg-blue-600 text-white'
                                : res.status === 'GREEN'
                                ? 'bg-green-100 text-green-800'
                                : res.status === 'YELLOW'
                                ? 'bg-yellow-100 text-yellow-800'
                                : 'bg-red-100 text-red-800'
                            }`}
                          >
                            {res.status === 'MEMORY'
                              ? '🧠 大腦記憶命中'
                              : res.status === 'GREEN'
                              ? '🟢 準確'
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
                                🏆 ［{res.itemTag}］ {res.matchedDescription} ($
                                {res.listPrice}元)
                              </option>
                            )}

                            {showGlobalSearch && (
                              <option value="REMOVE_MAPPING">
                                --- 🔍 物理 D 欄全表強行過濾結果 (
                                {filteredDropdownItems.length}筆) ---
                              </option>
                            )}

                            {/* 在候選清單選項中，硬性插入精準的 ［cand.item_tag］ 繁體中文提示字 */}
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
                              ❌ 此品項不套用牌價 / 屬於工資或排除項（保持空白）
                            </option>
                          </select>
                        </td>

                        {/* 注音打字不卡死防護鎖 */}
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
          </div>
        )}
      </div>
    </div>
  );
}
