import { useState } from 'react';
import type { ChangeEvent, DragEvent } from 'react';
import * as XLSX from 'xlsx';
// 引入 Lucide 圖示以提供高品質的 UI 視覺引導與極佳的視覺美感
import { 
  Upload, 
  Download, 
  AlertTriangle, 
  AlertCircle, 
  FileSpreadsheet, 
  ArrowRight, 
  Table, 
  Sparkles, 
  BookOpen, 
  ChevronDown, 
  ChevronUp, 
  RefreshCw, 
  FileText, 
  CheckCircle2,
  Info,
  Layers,
  Copy,
} from 'lucide-react';

type CellValue = string | number;
type RawRow = CellValue[];
type SubtotalType = '' | '小計' | '合計' | '總計';
type DataTab = 'versionB' | 'versionA' | 'raw';
type FeynmanTab = 'concept' | 'scenario' | 'lab';
type DownloadType = 'A' | 'B';

interface ProcessedRow {
  originalRowIndex: number;
  cells: CellValue[];
  filledDownFlags: boolean[];
  isSubtotal: boolean;
  subtotalType: SubtotalType;
  isGroupHeader: boolean;
}

interface Diagnostics {
  headerTrashRows: number;
  subtotalRowsCount: number;
  filledDownCellsCount: number;
  formattedDatesCount: number;
  originalGrandTotal: number;
  calculatedDetailTotal: number;
  isReconciled: boolean;
  diffAmount: number;
}

interface Kpis {
  originalRowsCount: number;
  cleanedRowsCount: number;
  uniqueCustomersCount: number;
  uniqueProductsCount: number;
  uniqueSalesCount: number;
  totalSales: number;
  totalQty: number;
  totalProfit: number;
  maxSalesVal: number;
  salesColLetter: string;
  dateColLetter: string;
  custColLetter: string;
  prodIdColLetter: string;
  qtyColLetter: string;
  priceColLetter: string;
  amountColLetter: string;
  profitColLetter: string;
  regionColLetter: string;
}

// ==========================================
// 內建 ERP 壞報表範例 CSV 數據 (供使用者一鍵測試)
// ==========================================
const SAMPLE_CSV = `集中資訊股份有限公司,,,,,,,,,
銷售日報,,,,,,,,,
報表期間,2026/04/01 - 2026/06/04,,,列印人員,系統批次,,,,
,,,,,,,,,
業務,日期,客戶,產品代號,產品名稱,區域,數量,單價,銷售金額,毛利
業務：A001 林業務,,,,,,,,,
A001 林業務,2026/04/15,客戶01,P048,產品048,北區,9,800,"7,200",1185
,2026/04/12,,P028,產品028,中區,2,600,"1,200",166
,2026/06/04,,P002,產品002,東區,18,800,"14,400",3790
,2026/04/29,,P038,產品038,中區,9,600,"5,400",1468
,2026/05/25,客戶11,P018,產品018,南區,5,800,"4,000",1246
,2026/04/14,,P025,產品025,南區,4,1200,"4,800",1390
,2026/04/06,,P030,產品030,中區,5,3200,"16,000",4484
客戶11 小計,,,,,,,14,000,
A001 林業務 小計,,,,,,,61,400,
業務：B002 王業務,,,,,,,,,
B002 王業務,2026/05/19,客戶20,P029,產品029,南區,10,2200,"22,000",6996
,2026/04/08,客戶20,P048,產品048,南區,4,800,"3,200",784
,2026/04/11,客戶20,P016,產品016,北區,6,2200,"13,200",1782
,2026/05/23,客戶20,P045,產品045,中區,20,1500,"30,000",5348
,2026/05/07,客戶20,P019,產品019,中區,23,1500,"34,500",4631
,2026/05/04,客戶26,P041,產品041,中區,19,3200,"60,800",17071
,2026/05/25,,P035,產品035,南區,8,3200,"25,600",3835
,,,,,,,,,
,2026/04/19,客戶26,P004,產品004,南區,6,1200,"7,200",1721
客戶20 小計,,,,,,,99,700,
客戶26 小計,,,,,,,93,600,
B002 王業務 小計,,,,,,,193,300,
總計,,,,,,,254,700,`;

export default function App() {
  // ==========================================
  // 狀態管理 (State)
  // ==========================================
  const [isDragging, setIsDragging] = useState(false);
  const [fileName, setFileName] = useState('');
  
  // 原始與清洗後的資料集
  const [headers, setHeaders] = useState<string[]>([]);
  const [rawPreview, setRawPreview] = useState<RawRow[]>([]);
  const [versionA, setVersionA] = useState<ProcessedRow[]>([]); // 含小計對帳表
  const [versionB, setVersionB] = useState<ProcessedRow[]>([]); // 純明細分析表
  
  // 分析指標與診斷數據
  const [diagnostics, setDiagnostics] = useState<Diagnostics | null>(null);
  const [kpis, setKpis] = useState<Kpis | null>(null);
  
  // 介面切換
  const [activeTab, setActiveTab] = useState<DataTab>('versionB');
  const [activeFeynmanTab, setActiveFeynmanTab] = useState<FeynmanTab>('concept');
  const [isDebugOpen, setIsDebugOpen] = useState(false);
  const [showCopiedText, setShowCopiedText] = useState('');

  // ==========================================
  // 工具函數 (Utilities)
  // ==========================================
  
  // 將數字索引轉換為 Excel 的欄位英文字母 (A, B, C... Z, AA...)
  const getColLetter = (index: number): string => {
    if (index === -1 || index === undefined) return "?";
    let temp = index;
    let letter = "";
    while (temp >= 0) {
      letter = String.fromCharCode((temp % 26) + 65) + letter;
      temp = Math.floor(temp / 26) - 1;
    }
    return letter;
  };

  // 1. 日期轉換器：安全處理 Excel 序號與各種日期字串
  const formatExcelDate = (val: CellValue): string => {
    if (!val) return '';
    if (typeof val === 'number') {
      const date = new Date((val - 25569) * 86400 * 1000);
      return formatDateObject(date);
    }
    const str = String(val).trim();
    if (/^\d{5}$/.test(str)) {
      const date = new Date((parseInt(str, 10) - 25569) * 86400 * 1000);
      return formatDateObject(date);
    }
    const parsed = Date.parse(str.replace(/\./g, '/'));
    if (!isNaN(parsed)) {
      return formatDateObject(new Date(parsed));
    }
    return str; 
  };

  const formatDateObject = (date: Date): string => {
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, '0');
    const d = String(date.getDate()).padStart(2, '0');
    return `${y}/${m}/${d}`;
  };

  // 2. 數值解析器：拔除千分位、貨幣符號、前後空白，強制轉 float
  const parseNumeric = (val: CellValue | null | undefined): number => {
    if (val === null || val === undefined) return 0;
    if (typeof val === 'number') return val;
    const cleaned = String(val).replace(/[,$\s%]/g, '');
    const num = parseFloat(cleaned);
    return isNaN(num) ? 0 : num;
  };

  // 複製文字提示
  const copyToClipboard = (text: string, label: string): void => {
    const textArea = document.createElement('textarea');
    textArea.value = text;
    document.body.appendChild(textArea);
    textArea.select();
    try {
      document.execCommand('copy');
      setShowCopiedText(label);
      setTimeout(() => setShowCopiedText(''), 2000);
    } catch (err) {
      console.error('無法複製', err);
    }
    document.body.removeChild(textArea);
  };

  // ==========================================
  // 核心清洗引擎 (Data Cleaning Engine)
  // ==========================================
  const processDataRows = (rawRows: RawRow[], name: string): void => {
    setFileName(name);
    setRawPreview(rawRows.slice(0, 25)); 

    try {
      // 步驟 1：【去頁首尾】智慧欄位定位器
      const KEYWORDS = ['日期', '客戶', '數量', '單價', '金額', '銷售', '業務', '區域', '毛利'];
      let headerIdx = -1;
      let maxMatches = 0;

      for (let i = 0; i < Math.min(rawRows.length, 15); i++) {
        const row = rawRows[i];
        if (!row || !Array.isArray(row)) continue;
        let matches = 0;
        row.forEach(cell => {
          const s = String(cell || '').trim();
          if (KEYWORDS.some(k => s.includes(k))) matches++;
        });
        if (matches > maxMatches && matches >= 3) {
          maxMatches = matches;
          headerIdx = i;
        }
      }

      if (headerIdx === -1) headerIdx = 4;

      const originalHeaders = rawRows[headerIdx].map(h => String(h || '').trim());
      setHeaders(originalHeaders);

      // 定位關鍵欄位之索引位置
      const salesColIdx = originalHeaders.findIndex(h => h.includes('業務'));
      const dateColIdx = originalHeaders.findIndex(h => h.includes('日期'));
      const custColIdx = originalHeaders.findIndex(h => h.includes('客戶'));
      const prodIdColIdx = originalHeaders.findIndex(h => h.includes('產品代號') || h.includes('品項') || h.includes('產品'));
      const qtyColIdx = originalHeaders.findIndex(h => h.includes('數量'));
      const priceColIdx = originalHeaders.findIndex(h => h.includes('單價'));
      const amountColIdx = originalHeaders.findIndex(h => h.includes('金額') || h.includes('銷售金額'));
      const profitColIdx = originalHeaders.findIndex(h => h.includes('毛利'));
      const regionColIdx = originalHeaders.findIndex(h => h.includes('區域'));

      // 定義需要【向下填滿 (Fill Down)】的維度欄位索引
      const fillDownCols: number[] = [];
      if (salesColIdx !== -1) fillDownCols.push(salesColIdx);
      if (dateColIdx !== -1) fillDownCols.push(dateColIdx);
      if (custColIdx !== -1) fillDownCols.push(custColIdx);
      if (regionColIdx !== -1) fillDownCols.push(regionColIdx);

      // 定義需要【資料型態轉換】的數值欄位索引
      const numericCols: number[] = [];
      if (qtyColIdx !== -1) numericCols.push(qtyColIdx);
      if (priceColIdx !== -1) numericCols.push(priceColIdx);
      if (amountColIdx !== -1) numericCols.push(amountColIdx);
      if (profitColIdx !== -1) numericCols.push(profitColIdx);

      // 狀態追蹤器
      const lastNonEmpty: Record<number, CellValue> = {};
      const versionARows: ProcessedRow[] = [];
      const versionBRows: ProcessedRow[] = [];
      
      let missingValueCount = 0;
      let subtotalRowCount = 0;
      let formattedDateCount = 0;
      let originalGrandTotal = 0;

      // 統計用變數 (擴充的統計分析)
      let totalQty = 0;
      let totalProfit = 0;
      let maxSalesVal = 0;
      const uniqueCusts = new Set<CellValue>();
      const uniqueProds = new Set<CellValue>();
      const uniqueSales = new Set<CellValue>();

      // 遍歷欄位名稱列之後的每一行數據
      for (let i = headerIdx + 1; i < rawRows.length; i++) {
        const row = rawRows[i];
        if (!row || row.every(cell => String(cell || '').trim() === '')) {
          continue; 
        }

        // 偵測並解析「課別/業務分群列」
        const nonBgCells = row.map((c, idx) => ({val: String(c || '').trim(), idx})).filter(o => o.val !== '');
        const isGroupHeader = nonBgCells.length === 1 && (nonBgCells[0].val.includes('：') || nonBgCells[0].val.includes(':'));

        if (isGroupHeader) {
          const val = nonBgCells[0].val;
          const cleanVal = val.split(/：|:/)[1]?.trim() || val;
          if (salesColIdx !== -1) {
            lastNonEmpty[salesColIdx] = cleanVal;
          }
          continue; 
        }

        // 偵測是否為小計、合計、或總計列
        let isSubtotal = false;
        let subtotalType: SubtotalType = '';
        row.forEach(cell => {
          const s = String(cell || '').trim();
          if (s.includes('小計') || s.includes('合計') || s.includes('總計') || s.toLowerCase().includes('total') || s.toLowerCase().includes('subtotal')) {
            isSubtotal = true;
            if (s.includes('小計')) subtotalType = '小計';
            else if (s.includes('總計') || s.toLowerCase().includes('grand total')) subtotalType = '總計';
            else subtotalType = '合計';
          }
        });

        // 雙版本分流處理
        if (isSubtotal) {
          subtotalRowCount++;
          const subtotalCells = row.map((cell, colIdx) => {
            const strVal = String(cell || '').trim();
            if (numericCols.includes(colIdx) && strVal !== '') {
              return parseNumeric(strVal);
            }
            return strVal;
          });

          const detectedSubtotalType = subtotalType as SubtotalType;
          if (detectedSubtotalType === '總計' && amountColIdx !== -1) {
            const totalAmtStr = String(row[amountColIdx] || '').trim();
            originalGrandTotal = parseNumeric(totalAmtStr);
          }

          versionARows.push({
            originalRowIndex: i,
            cells: subtotalCells,
            filledDownFlags: Array(row.length).fill(false),
            isSubtotal: true,
            subtotalType: detectedSubtotalType,
            isGroupHeader: false
          });
          continue; 
        }

        // 排除無效雜訊行
        const hasTransactionData = numericCols.some(idx => {
          const val = String(row[idx] || '').trim();
          return val !== '' && !isNaN(parseFloat(val.replace(/[,$\s%]/g, '')));
        });

        if (!hasTransactionData && row.filter(c => String(c || '').trim() !== '').length < 3) {
          continue; 
        }

        // 執行資料清洗與「向下填滿」
        const cleanedCells: CellValue[] = [];
        const filledDownFlags = Array(row.length).fill(false);

        for (let colIdx = 0; colIdx < row.length; colIdx++) {
          let cellVal: CellValue = String(row[colIdx] || '').trim();

          // 1. 向下填滿邏輯
          if (fillDownCols.includes(colIdx)) {
            if (cellVal === '') {
              if (lastNonEmpty[colIdx] !== undefined) {
                cellVal = lastNonEmpty[colIdx];
                filledDownFlags[colIdx] = true;
                missingValueCount++;
              }
            } else {
              lastNonEmpty[colIdx] = cellVal;
            }
          }

          // 2. 文字前後去空格
          if (!numericCols.includes(colIdx) && colIdx !== dateColIdx) {
            cellVal = String(cellVal).replace(/^\s+|\s+$/g, '');
          }

          // 3. 日期格式化為 YYYY/MM/DD
          if (colIdx === dateColIdx && cellVal !== '') {
            const formatted = formatExcelDate(cellVal);
            if (formatted !== cellVal) {
              cellVal = formatted;
              formattedDateCount++;
            }
          }

          // 4. 強制數值轉碼型態
          if (numericCols.includes(colIdx)) {
            cellVal = parseNumeric(cellVal);
          }

          cleanedCells.push(cellVal);
        }

        const rowObj: ProcessedRow = {
          originalRowIndex: i,
          cells: cleanedCells,
          filledDownFlags,
          isSubtotal: false,
          subtotalType: '',
          isGroupHeader: false
        };

        // 進行新指標的加總統計 (版本 B 明細專屬)
        if (qtyColIdx !== -1) totalQty += Number(cleanedCells[qtyColIdx] || 0);
        if (profitColIdx !== -1) totalProfit += Number(cleanedCells[profitColIdx] || 0);
        if (amountColIdx !== -1) {
          const amt = Number(cleanedCells[amountColIdx] || 0);
          if (amt > maxSalesVal) maxSalesVal = amt;
        }
        if (custColIdx !== -1 && cleanedCells[custColIdx]) uniqueCusts.add(cleanedCells[custColIdx]);
        if (prodIdColIdx !== -1 && cleanedCells[prodIdColIdx]) uniqueProds.add(cleanedCells[prodIdColIdx]);
        if (salesColIdx !== -1 && cleanedCells[salesColIdx]) uniqueSales.add(cleanedCells[salesColIdx]);

        versionARows.push(rowObj);
        versionBRows.push(rowObj);
      }

      // 計算清洗後版本 B 的加總金額
      let totalSalesB = 0;
      versionBRows.forEach(r => {
        if (amountColIdx !== -1) {
          totalSalesB += Number(r.cells[amountColIdx] || 0);
        }
      });

      // 備份原始總計
      if (originalGrandTotal === 0) {
        const lastSub = [...versionARows].reverse().find(r => r.isSubtotal);
        if (lastSub && amountColIdx !== -1) {
          originalGrandTotal = Number(lastSub.cells[amountColIdx] || 0);
        }
      }

      const reconciled = Math.round(originalGrandTotal) === Math.round(totalSalesB);

      // 設定輸出狀態
      setVersionA(versionARows);
      setVersionB(versionBRows);
      
      setDiagnostics({
        headerTrashRows: headerIdx,
        subtotalRowsCount: subtotalRowCount,
        filledDownCellsCount: missingValueCount,
        formattedDatesCount: formattedDateCount,
        originalGrandTotal,
        calculatedDetailTotal: totalSalesB,
        isReconciled: reconciled,
        diffAmount: Math.abs(originalGrandTotal - totalSalesB)
      });

      // 整合豐富的 KPIs 提供 Excel 完美比對
      setKpis({
        originalRowsCount: rawRows.length,
        cleanedRowsCount: versionBRows.length,
        uniqueCustomersCount: uniqueCusts.size,
        uniqueProductsCount: uniqueProds.size || 1, // 防止除以 0
        uniqueSalesCount: uniqueSales.size,
        totalSales: totalSalesB,
        totalQty,
        totalProfit,
        maxSalesVal,
        // 保存對應 Excel 的欄位字母
        salesColLetter: getColLetter(salesColIdx),
        dateColLetter: getColLetter(dateColIdx),
        custColLetter: getColLetter(custColIdx),
        prodIdColLetter: getColLetter(prodIdColIdx),
        qtyColLetter: getColLetter(qtyColIdx),
        priceColLetter: getColLetter(priceColIdx),
        amountColLetter: getColLetter(amountColIdx),
        profitColLetter: getColLetter(profitColIdx),
        regionColLetter: getColLetter(regionColIdx)
      });

    } catch (error: unknown) {
      console.error("清洗出錯：", error);
      alert("❌ 報表解析失敗，請確認檔案格式是否正確。");
    }
  };

  // ==========================================
  // 檔案上傳事件監聽
  // ==========================================
  const handleFileChange = (event: ChangeEvent<HTMLInputElement>): void => {
    const file = event.target.files?.[0];
    if (file) processExcelFile(file);
  };

  const processExcelFile = (file: File): void => {
    const reader = new FileReader();
    reader.onload = (): void => {
      try {
        if (!(reader.result instanceof ArrayBuffer)) {
          throw new Error('無法取得 Excel 的 ArrayBuffer 資料。');
        }
        const workbook = XLSX.read(reader.result, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        if (!firstSheetName) throw new Error('活頁簿中沒有工作表。');
        const worksheet = workbook.Sheets[firstSheetName];
        if (!worksheet) throw new Error(`找不到工作表：${firstSheetName}`);
        const rawRows = XLSX.utils.sheet_to_json<CellValue[]>(worksheet, {
          header: 1,
          defval: '',
          raw: true,
        });
        processDataRows(rawRows, file.name);
      } catch (error: unknown) {
        console.error('Excel 解析失敗：', error);
        alert('❌ Excel 檔案解析失敗，請確認檔案格式與內容。');
      }
    };
    reader.onerror = (): void => {
      console.error('FileReader 讀取失敗：', reader.error);
      alert('❌ 檔案讀取失敗，請重新選擇檔案。');
    };
    reader.readAsArrayBuffer(file);
  };

  const handleDragOver = (event: DragEvent<HTMLDivElement>): void => {
    event.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = (): void => {
    setIsDragging(false);
  };

  const handleDrop = (event: DragEvent<HTMLDivElement>): void => {
    event.preventDefault();
    setIsDragging(false);
    const file = event.dataTransfer.files[0];
    if (file) processExcelFile(file);
  };

  // 載入內建範例數據
  const loadSampleData = (): void => {
    const rawRows = SAMPLE_CSV.split('\n').map(line => {
      const result: CellValue[] = [];
      let current = '';
      let inQuotes = false;
      for (let i = 0; i < line.length; i++) {
        const char = line[i];
        if (char === '"') {
          inQuotes = !inQuotes;
        } else if (char === ',' && !inQuotes) {
          result.push(current);
          current = '';
        } else {
          current += char;
        }
      }
      result.push(current);
      return result;
    });
    processDataRows(rawRows, "內建_ERP銷售日報_範例資料.xlsx");
  };

  // ==========================================
  // 下載輸出模組 (多工作表: Data + KPI)
  // ==========================================
  const handleDownload = (type: DownloadType): void => {
    if (!kpis) {
      alert('⚠️ 請先上傳並完成資料清洗，再下載結果。');
      return;
    }
    const rowsToExport = type === 'B' ? versionB : versionA;
    const suffix = type === 'B' ? '版本B_純明細分析表' : '版本A_含小計對帳表';
    
    // 生成精確的 _yyyymmdd_hhnnss 時間戳記 (VB 的 nn 代表分鐘)
    const now = new Date();
    const yyyy = now.getFullYear();
    const mm = String(now.getMonth() + 1).padStart(2, '0');
    const dd = String(now.getDate()).padStart(2, '0');
    const hh = String(now.getHours()).padStart(2, '0');
    const nn = String(now.getMinutes()).padStart(2, '0'); 
    const ss = String(now.getSeconds()).padStart(2, '0');
    const timestamp = `_${yyyy}${mm}${dd}_${hh}${nn}${ss}`;

    // 建立新活頁簿
    const wb = XLSX.utils.book_new();

    // 1. 寫入 "Data" 乾淨明細工作表 (名稱一律強制為 Data)
    const dataSheetData: RawRow[] = [headers];
    rowsToExport.forEach(r => {
      dataSheetData.push(r.cells);
    });
    const wsData = XLSX.utils.aoa_to_sheet(dataSheetData);
    XLSX.utils.book_append_sheet(wb, wsData, "Data");

    // 2. 建立並寫入 "KPI" 數據驗證工作表 (用於對帳比對與公式驗證)
    const kpiSheetData = [
      ["GAI 銷售日報二次分析 - 數據核對驗證指標表 (Gigi 智慧提供)"],
      [`清洗原始檔案名稱: ${fileName}`],
      [`匯出執行時間: ${yyyy}/${mm}/${dd} ${hh}:${nn}:${ss}`],
      [],
      ["指標分類", "核對指標名稱", "洗淨後數值", "建議 Excel 驗證公式 (請直接貼在 Excel 任意儲存格)"],
      ["整體銷售", "總銷售金額 (Sales)", kpis.totalSales, `=SUM(Data!${kpis.amountColLetter}:${kpis.amountColLetter})`],
      ["整體銷售", "總交易筆數 (Transactions)", kpis.cleanedRowsCount, `=COUNTA(Data!${kpis.amountColLetter}:${kpis.amountColLetter})-1`],
      ["整體銷售", "總出貨數量 (Qty)", kpis.totalQty, `=SUM(Data!${kpis.qtyColLetter}:${kpis.qtyColLetter})`],
      ["整體銷售", "總利潤額 (Gross Profit)", kpis.totalProfit, `=SUM(Data!${kpis.profitColLetter}:${kpis.profitColLetter})`],
      ["整體獲利", "整體毛利率 (Margin)", kpis.totalSales > 0 ? (kpis.totalProfit / kpis.totalSales) : 0, `=SUM(Data!${kpis.profitColLetter}:${kpis.profitColLetter})/SUM(Data!${kpis.amountColLetter}:${kpis.amountColLetter})`],
      ["單筆分析", "最大單筆銷售額 (Max Sale)", kpis.maxSalesVal, `=MAX(Data!${kpis.amountColLetter}:${kpis.amountColLetter})`],
      ["客戶維度", "不重複客戶數", kpis.uniqueCustomersCount, `Office365適用: =ROWS(UNIQUE(FILTER(Data!${kpis.custColLetter}2:${kpis.custColLetter}${kpis.cleanedRowsCount + 1}, Data!${kpis.custColLetter}2:${kpis.custColLetter}${kpis.cleanedRowsCount + 1}<>"")))`],
      ["產品維度", "不重複產品數", kpis.uniqueProductsCount, `Office365適用: =ROWS(UNIQUE(FILTER(Data!${kpis.prodIdColLetter}2:${kpis.prodIdColLetter}${kpis.cleanedRowsCount + 1}, Data!${kpis.prodIdColLetter}2:${kpis.prodIdColLetter}${kpis.cleanedRowsCount + 1}<>"")))`],
      ["業務維度", "不重複業務數", kpis.uniqueSalesCount, `Office365適用: =ROWS(UNIQUE(FILTER(Data!${kpis.salesColLetter}2:${kpis.salesColLetter}${kpis.cleanedRowsCount + 1}, Data!${kpis.salesColLetter}2:${kpis.salesColLetter}${kpis.cleanedRowsCount + 1}<>"")))`],
      [],
      ["💡 Gigi 助教的 Excel 核對對帳指引："],
      ["1. 本活頁簿包含兩個工作表：【Data】與【KPI】。"],
      ["2. 當您將【Data】工作表拿去建立樞紐分析表時，請務必比對樞紐的總計金額是否與上述指標完全一致。"],
      ["3. 所有『合計/小計列』皆已排除，利用上述公式可 100% 準確核對，防止重複計算。"]
    ];
    const wsKPI = XLSX.utils.aoa_to_sheet(kpiSheetData);
    XLSX.utils.book_append_sheet(wb, wsKPI, "KPI");

    // 格式化輸出檔名：原始檔名 + _yyyymmdd_hhnnss + 檔案後綴字
    const dotIdx = fileName.lastIndexOf('.');
    const baseName = dotIdx !== -1 ? fileName.substring(0, dotIdx) : fileName;
    const ext = dotIdx !== -1 ? fileName.substring(dotIdx) : '.xlsx';
    const outName = `${baseName}${timestamp}_${suffix}${ext}`;

    XLSX.writeFile(wb, outName);
  };

  return (
    <div className="min-h-screen bg-slate-50 text-slate-800 font-sans">
      {/* 頂部美化導航列 */}
      <header className="bg-gradient-to-r from-slate-900 via-indigo-950 to-slate-900 text-white shadow-md py-4 px-6">
        <div className="max-w-7xl mx-auto flex flex-col md:flex-row justify-between items-center gap-4">
          <div className="flex items-center gap-3">
            <div className="p-2 bg-indigo-600 rounded-lg shadow-inner">
              <Sparkles className="w-6 h-6 text-emerald-400 animate-pulse" />
            </div>
            <div>
              <h1 className="text-xl font-bold tracking-tight">GAI 銷售日報智慧清洗與 KPI 對帳核對儀表板</h1>
              <p className="text-xs text-slate-300">將 ERP 紊亂報表自動轉換為 Data + KPI 標準雙工作表活頁簿</p>
            </div>
          </div>
          <div className="flex items-center gap-2 bg-indigo-900/50 px-4 py-2 rounded-full border border-indigo-500/30">
            <span className="w-2.5 h-2.5 bg-emerald-400 rounded-full animate-ping"></span>
            <span className="text-xs font-medium text-indigo-100">Gigi 助教正在線上協助您 💖</span>
          </div>
        </div>
      </header>

      {/* 歡迎橫幅（助教語氣） */}
      <section className="bg-indigo-50 border-b border-indigo-100 py-6 px-6">
        <div className="max-w-7xl mx-auto flex items-start gap-4">
          <div className="w-12 h-12 rounded-full bg-indigo-600 flex items-center justify-center text-white font-bold text-lg flex-shrink-0 shadow-md">
            Gi
          </div>
          <div>
            <h4 className="font-semibold text-indigo-900">嗨！我是你的 GAI 輔助學習助教 Gigi 💖</h4>
            <p className="text-sm text-slate-600 mt-1 leading-relaxed">
              永遠記得我們的最高指導原則：<strong className="text-indigo-700">「先診斷，再自動化。」</strong><br />
              依您的要求，我們進行了全面升級！現在洗乾淨的明細工作表會一律命名為 <strong className="text-emerald-700">`Data`</strong>；並且新增了獨立的 <strong className="text-indigo-700">`KPI` 核對工作表</strong>。
              同時，我們也支援在匯出檔案時自動在檔名後綴加入 <strong className="text-indigo-700">`_yyyymmdd_hhnnss`</strong> 時間戳記，讓您的每次產出都能被完美追蹤！
            </p>
          </div>
        </div>
      </section>

      <main className="max-w-7xl mx-auto px-4 py-8 space-y-8">
        
        {/* 第一部分：上傳區與載入範例 */}
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
          <div className="lg:col-span-2">
            <div 
              onDragOver={handleDragOver}
              onDragLeave={handleDragLeave}
              onDrop={handleDrop}
              className={`border-2 border-dashed rounded-xl p-8 flex flex-col items-center justify-center transition-all duration-200 h-64 relative overflow-hidden ${
                isDragging 
                  ? 'border-indigo-600 bg-indigo-50/50 shadow-inner' 
                  : 'border-slate-300 bg-white hover:border-indigo-400 hover:bg-slate-50/30'
              }`}
            >
              <input 
                type="file" 
                id="excelFile" 
                accept=".xlsx, .xls, .csv" 
                onChange={handleFileChange} 
                className="absolute inset-0 opacity-0 cursor-pointer"
              />
              <div className="text-center space-y-4">
                <div className="mx-auto w-16 h-16 bg-indigo-100 rounded-full flex items-center justify-center text-indigo-600 shadow-sm">
                  <Upload className="w-8 h-8" />
                </div>
                <div>
                  <p className="font-semibold text-slate-700 text-lg">
                    拖曳檔案到此處，或 <span className="text-indigo-600 underline">點擊瀏覽檔案</span>
                  </p>
                  <p className="text-xs text-slate-400 mt-1">支援 .xlsx, .xls, .csv 檔案</p>
                </div>
              </div>
            </div>
          </div>

          <div className="bg-white rounded-xl border border-slate-200 p-6 flex flex-col justify-between shadow-sm">
            <div className="space-y-4">
              <div className="flex items-center gap-2 text-indigo-900 font-semibold">
                <FileSpreadsheet className="w-5 h-5 text-indigo-600" />
                <span>快速體驗與測試</span>
              </div>
              <p className="text-sm text-slate-500 leading-relaxed">
                如果您手邊沒有 ERP 導出的 Excel 檔案，我們已經為您嵌入了包含「合併儲存格、多層次小計、首尾干擾與格式錯亂」的典型壞報表。
              </p>
            </div>
            <button
              onClick={loadSampleData}
              className="mt-6 w-full py-3 bg-gradient-to-r from-indigo-600 to-indigo-700 text-white rounded-lg font-medium hover:from-indigo-700 hover:to-indigo-800 transition-all flex items-center justify-center gap-2 shadow"
            >
              <RefreshCw className="w-4 h-4" />
              載入內建壞報表範例
            </button>
          </div>
        </div>

        {/* 如果資料已解析，顯示 KPI 看板與 Gigi 的診斷表 */}
        {diagnostics && kpis && (
          <div className="space-y-8 animate-fadeIn">
            
            {/* KPI 看板與自動核對 */}
            <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
              
              {/* 卡片 1：列數與筆數對齊 */}
              <div className="bg-white rounded-xl border border-slate-200 p-6 shadow-sm flex flex-col justify-between">
                <div>
                  <p className="text-xs font-semibold text-slate-400 tracking-wider uppercase">原始總列數 VS 乾淨交易筆數</p>
                  <div className="flex items-baseline gap-2 mt-2">
                    <span className="text-3xl font-bold text-slate-800">{kpis.originalRowsCount}</span>
                    <span className="text-slate-400 text-xs">列原檔</span>
                    <ArrowRight className="w-4 h-4 text-slate-400" />
                    <span className="text-3xl font-bold text-indigo-600">{kpis.cleanedRowsCount}</span>
                    <span className="text-slate-400 text-xs">列純明細</span>
                  </div>
                </div>
                <div className="text-xs text-slate-500 mt-4 border-t pt-2 border-slate-100 flex justify-between">
                  <span>不重複產品：{kpis.uniqueProductsCount} 支</span>
                  <span>不重複業務：{kpis.uniqueSalesCount} 位</span>
                </div>
              </div>

              {/* 卡片 2：核心對帳檢核 */}
              <div className={`rounded-xl border p-6 shadow-sm flex flex-col justify-between ${
                diagnostics.isReconciled 
                  ? 'bg-emerald-50/50 border-emerald-200 text-emerald-900' 
                  : 'bg-amber-50/50 border-amber-200 text-amber-900'
              }`}>
                <div>
                  <p className="text-xs font-semibold tracking-wider uppercase text-slate-500">核心金額自動對帳</p>
                  <div className="flex items-center gap-2 mt-2">
                    {diagnostics.isReconciled ? (
                      <CheckCircle2 className="w-8 h-8 text-emerald-600 flex-shrink-0" />
                    ) : (
                      <AlertTriangle className="w-8 h-8 text-amber-600 flex-shrink-0" />
                    )}
                    <div>
                      <h3 className="text-xl font-bold">
                        {diagnostics.isReconciled ? "對帳成功 🟢" : "發現金額落差 ⚠️"}
                      </h3>
                      <p className="text-xs opacity-80 mt-1">
                        原檔總計：${diagnostics.originalGrandTotal.toLocaleString()} <br />
                        清洗明細：${diagnostics.calculatedDetailTotal.toLocaleString()}
                      </p>
                    </div>
                  </div>
                </div>
                <div className="text-xs mt-4 border-t pt-2 border-slate-200/50">
                  {diagnostics.isReconciled ? (
                    <span>對帳 100% 成功！清洗後金額完全相符。</span>
                  ) : (
                    <span>
                      ⚠️ 差額 ${diagnostics.diffAmount.toLocaleString()}。這代表原 ERP 自帶的小計公式有漏計（如王業務小計），Gigi 幫您精準抓實了！
                    </span>
                  )}
                </div>
              </div>

              {/* 卡片 3：商業基本指標 */}
              <div className="bg-white rounded-xl border border-slate-200 p-6 shadow-sm flex flex-col justify-between">
                <div>
                  <p className="text-xs font-semibold text-slate-400 tracking-wider uppercase">核心利潤與營運摘要</p>
                  <div className="grid grid-cols-2 gap-2 mt-2">
                    <div>
                      <span className="text-xl font-bold text-slate-800">${kpis.totalProfit.toLocaleString()}</span>
                      <p className="text-xs text-slate-400">總毛利額</p>
                    </div>
                    <div>
                      <span className="text-xl font-bold text-indigo-600">
                        {kpis.totalSales > 0 ? ((kpis.totalProfit / kpis.totalSales) * 100).toFixed(1) : 0}%
                      </span>
                      <p className="text-xs text-slate-400">整體毛利率</p>
                    </div>
                  </div>
                </div>
                <div className="text-xs text-slate-500 mt-4 border-t pt-2 border-slate-100 flex justify-between">
                  <span>總銷售量：{kpis.totalQty.toLocaleString()}</span>
                  <span>最大單筆：${kpis.maxSalesVal.toLocaleString()}</span>
                </div>
              </div>

            </div>

            {/* 新增功能：Excel 二次分析對帳驗證大面板 */}
            <div className="bg-gradient-to-r from-indigo-50 to-sky-50 rounded-xl border border-indigo-100 p-6 shadow-sm">
              <div className="flex items-center gap-2 mb-4 text-indigo-950 font-bold text-base">
                <Layers className="w-5 h-5 text-indigo-600" />
                <h3>Gigi 助教的「Excel 二次驗證公式速查」</h3>
              </div>
              <p className="text-xs text-slate-600 mb-4">
                當您下載活頁簿後，明細資料已完美收納至 <span className="bg-emerald-100 text-emerald-800 px-1 py-0.5 rounded font-mono font-bold">Data</span> 工作表，核對指標則位於 <span className="bg-indigo-100 text-indigo-800 px-1 py-0.5 rounded font-mono font-bold">KPI</span> 工作表。您可以使用下方公式快速驗證！
              </p>
              
              <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                
                <div className="bg-white p-4 rounded-lg border border-indigo-100 flex flex-col justify-between">
                  <div>
                    <div className="flex justify-between items-center text-xs text-slate-400 font-semibold">
                      <span>總銷售金額 (Sales)</span>
                      <span className="text-indigo-600">欄位 {kpis.amountColLetter}</span>
                    </div>
                    <div className="text-lg font-bold text-slate-800 mt-1">${kpis.totalSales.toLocaleString()}</div>
                    <div className="bg-slate-50 p-2 rounded mt-2 text-xs font-mono text-slate-500 select-all border border-slate-100 relative">
                      =SUM(Data!{kpis.amountColLetter}:{kpis.amountColLetter})
                    </div>
                  </div>
                  <button 
                    onClick={() => copyToClipboard(`=SUM(Data!${kpis.amountColLetter}:${kpis.amountColLetter})`, 'sales')}
                    className="mt-2 text-left text-xs text-indigo-600 font-semibold flex items-center gap-1 hover:text-indigo-800"
                  >
                    <Copy className="w-3 h-3" />
                    {showCopiedText === 'sales' ? '已複製！' : '複製 Excel 公式'}
                  </button>
                </div>

                <div className="bg-white p-4 rounded-lg border border-indigo-100 flex flex-col justify-between">
                  <div>
                    <div className="flex justify-between items-center text-xs text-slate-400 font-semibold">
                      <span>總利潤金額 (Profit)</span>
                      <span className="text-indigo-600">欄位 {kpis.profitColLetter}</span>
                    </div>
                    <div className="text-lg font-bold text-slate-800 mt-1">${kpis.totalProfit.toLocaleString()}</div>
                    <div className="bg-slate-50 p-2 rounded mt-2 text-xs font-mono text-slate-500 select-all border border-slate-100 relative">
                      =SUM(Data!{kpis.profitColLetter}:{kpis.profitColLetter})
                    </div>
                  </div>
                  <button 
                    onClick={() => copyToClipboard(`=SUM(Data!${kpis.profitColLetter}:${kpis.profitColLetter})`, 'profit')}
                    className="mt-2 text-left text-xs text-indigo-600 font-semibold flex items-center gap-1 hover:text-indigo-800"
                  >
                    <Copy className="w-3 h-3" />
                    {showCopiedText === 'profit' ? '已複製！' : '複製 Excel 公式'}
                  </button>
                </div>

                <div className="bg-white p-4 rounded-lg border border-indigo-100 flex flex-col justify-between">
                  <div>
                    <div className="flex justify-between items-center text-xs text-slate-400 font-semibold">
                      <span>總交易筆數 (Count)</span>
                      <span className="text-indigo-600">欄位 {kpis.amountColLetter}</span>
                    </div>
                    <div className="text-lg font-bold text-slate-800 mt-1">{kpis.cleanedRowsCount} 筆</div>
                    <div className="bg-slate-50 p-2 rounded mt-2 text-xs font-mono text-slate-500 select-all border border-slate-100 relative">
                      =COUNTA(Data!{kpis.amountColLetter}:{kpis.amountColLetter})-1
                    </div>
                  </div>
                  <button 
                    onClick={() => copyToClipboard(`=COUNTA(Data!${kpis.amountColLetter}:${kpis.amountColLetter})-1`, 'count')}
                    className="mt-2 text-left text-xs text-indigo-600 font-semibold flex items-center gap-1 hover:text-indigo-800"
                  >
                    <Copy className="w-3 h-3" />
                    {showCopiedText === 'count' ? '已複製！' : '複製 Excel 公式'}
                  </button>
                </div>

              </div>
            </div>

            {/* Gigi 的 報表問題診斷表 (MANDATORY) */}
            <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
              <div className="bg-slate-900 text-white py-4 px-6 flex justify-between items-center">
                <div className="flex items-center gap-2">
                  <Sparkles className="w-5 h-5 text-emerald-400" />
                  <h3 className="font-semibold text-lg">Gigi 助教的「報表問題診斷與修復表」</h3>
                </div>
                <span className="text-xs bg-indigo-900 text-indigo-200 py-1 px-2.5 rounded font-mono">先診斷，再自動化</span>
              </div>
              <div className="overflow-x-auto">
                <table className="w-full text-left border-collapse text-sm">
                  <thead>
                    <tr className="bg-slate-100 text-slate-600 border-b border-slate-200">
                      <th className="py-3 px-4 font-semibold">編號</th>
                      <th className="py-3 px-4 font-semibold">問題類型</th>
                      <th className="py-3 px-4 font-semibold">問題描述 (ERP壞排版)</th>
                      <th className="py-3 px-4 font-semibold">對分析之影響</th>
                      <th className="py-3 px-4 font-semibold">Gigi 智慧修復方式</th>
                      <th className="py-3 px-4 font-semibold">使用工具</th>
                      <th className="py-3 px-4 font-semibold">檢核狀態</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100 text-slate-600">
                    <tr>
                      <td className="py-3.5 px-4 font-mono font-bold text-indigo-600">01</td>
                      <td className="py-3.5 px-4 font-semibold text-slate-800">頁首非數據雜訊</td>
                      <td className="py-3.5 px-4">前 {diagnostics.headerTrashRows} 列包含公司名、列印日期等干擾抬頭。</td>
                      <td className="py-3.5 px-4">導致欄位錯位、公式載入失敗。</td>
                      <td className="py-3.5 px-4">自動定位最多特徵關鍵字的第 {diagnostics.headerTrashRows + 1} 列作為 Header 欄位名稱。</td>
                      <td className="py-3.5 px-4 font-mono text-xs bg-slate-50 text-indigo-600 py-0.5 px-1.5 rounded">智慧網格定位器</td>
                      <td className="py-3.5 px-4"><span className="inline-flex items-center px-2 py-0.5 rounded text-xs font-medium bg-emerald-100 text-emerald-800">100% 已切除</span></td>
                    </tr>
                    <tr>
                      <td className="py-3.5 px-4 font-mono font-bold text-indigo-600">02</td>
                      <td className="py-3.5 px-4 font-semibold text-slate-800">群組標題列獨立</td>
                      <td className="py-3.5 px-4">`業務：A001 林業務` 獨立成列，明細無分類屬性。</td>
                      <td className="py-3.5 px-4">無法對業務人員進行篩選與加總統計。</td>
                      <td className="py-3.5 px-4">解析冒號前後內容，自動填入「業務」欄位進行向下填滿。</td>
                      <td className="py-3.5 px-4 font-mono text-xs bg-slate-50 text-indigo-600 py-0.5 px-1.5 rounded">群組對齊狀態機</td>
                      <td className="py-3.5 px-4"><span className="inline-flex items-center px-2 py-0.5 rounded text-xs font-medium bg-emerald-100 text-emerald-800">對齊成功</span></td>
                    </tr>
                    <tr>
                      <td className="py-3.5 px-4 font-mono font-bold text-indigo-600">03</td>
                      <td className="py-3.5 px-4 font-semibold text-slate-800">空白儲存格/未填滿</td>
                      <td className="py-3.5 px-4">「日期」、「客戶」等欄位因排版留空。</td>
                      <td className="py-3.5 px-4">樞紐分析會出現大量不名所以的 `(空白)` 分類。</td>
                      <td className="py-3.5 px-4">自動在「日期」與「客戶」等關鍵欄位執行「向下填滿」共 {diagnostics.filledDownCellsCount} 格。</td>
                      <td className="py-3.5 px-4 font-mono text-xs bg-slate-50 text-indigo-600 py-0.5 px-1.5 rounded">向上層記憶填充器</td>
                      <td className="py-3.5 px-4"><span className="inline-flex items-center px-2 py-0.5 rounded text-xs font-medium bg-emerald-100 text-emerald-800">{diagnostics.filledDownCellsCount} 格已補滿</span></td>
                    </tr>
                    <tr>
                      <td className="py-3.5 px-4 font-mono font-bold text-indigo-600">04</td>
                      <td className="py-3.5 px-4 font-semibold text-slate-800">文字空格與雜訊</td>
                      <td className="py-3.5 px-4">客戶名稱、產品名稱中含有全形或半形空格。</td>
                      <td className="py-3.5 px-4">`\"客戶A\"` 與 `\"客戶A \"` 會被判定成兩個群組，造成數值失真。</td>
                      <td className="py-3.5 px-4">強制針對所有非數值之文字欄位進行兩端去空格處理。</td>
                      <td className="py-3.5 px-4 font-mono text-xs bg-slate-50 text-indigo-600 py-0.5 px-1.5 rounded">Regex 文字洗淨機</td>
                      <td className="py-3.5 px-4"><span className="inline-flex items-center px-2 py-0.5 rounded text-xs font-medium bg-emerald-100 text-emerald-800">空格清除</span></td>
                    </tr>
                    <tr>
                      <td className="py-3.5 px-4 font-mono font-bold text-indigo-600">05</td>
                      <td className="py-3.5 px-4 font-semibold text-slate-800">數值與日期型態錯亂</td>
                      <td className="py-3.5 px-4">金額與數量包含千分位、日期格式不一致或為 Excel 序號。</td>
                      <td className="py-3.5 px-4">無法直接在 Excel 進行 SUM 加總，也無法展開年/月/日。</td>
                      <td className="py-3.5 px-4">將日期標準化為 YYYY/MM/DD（共修正 {diagnostics.formattedDatesCount} 筆）；將數值欄位強制轉換為 Float。</td>
                      <td className="py-3.5 px-4 font-mono text-xs bg-slate-50 text-indigo-600 py-0.5 px-1.5 rounded">型態強制轉碼器</td>
                      <td className="py-3.5 px-4"><span className="inline-flex items-center px-2 py-0.5 rounded text-xs font-medium bg-emerald-100 text-emerald-800">型態一致化</span></td>
                    </tr>
                    <tr>
                      <td className="py-3.5 px-4 font-mono font-bold text-indigo-600">06</td>
                      <td className="py-3.5 px-4 font-semibold text-slate-800">小計列重複干擾</td>
                      <td className="py-3.5 px-4">報表中夾雜著多層次的「客戶小計」、「業務小計」、「總計」。</td>
                      <td className="py-3.5 px-4">直接對整直欄加總會導致重複計算，銷售總額瞬間倍增。</td>
                      <td className="py-3.5 px-4">智慧偵測含有小計或總計字樣之行，隔離為對帳表 (版本 A) 與無小計明細表 (版本 B)。</td>
                      <td className="py-3.5 px-4 font-mono text-xs bg-slate-50 text-indigo-600 py-0.5 px-1.5 rounded">雙版本資料流分碼</td>
                      <td className="py-3.5 px-4"><span className="inline-flex items-center px-2 py-0.5 rounded text-xs font-medium bg-emerald-100 text-emerald-800">兩版本分流成功</span></td>
                    </tr>
                  </tbody>
                </table>
              </div>
            </div>

            {/* 中間：雙版本切換與資料預覽 */}
            <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
              <div className="border-b border-slate-200 bg-slate-50 px-6 pt-4 flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
                
                {/* Tabs 頁籤 */}
                <div className="flex gap-2">
                  <button
                    onClick={() => setActiveTab('versionB')}
                    className={`py-3 px-4 text-sm font-semibold border-b-2 transition-all flex items-center gap-2 ${
                      activeTab === 'versionB'
                        ? 'border-indigo-600 text-indigo-600'
                        : 'border-transparent text-slate-500 hover:text-slate-800'
                    }`}
                  >
                    <Table className="w-4 h-4" />
                    版本 B：純明細分析表 (匯出至 Data 工作表 🌟)
                  </button>
                  <button
                    onClick={() => setActiveTab('versionA')}
                    className={`py-3 px-4 text-sm font-semibold border-b-2 transition-all flex items-center gap-2 ${
                      activeTab === 'versionA'
                        ? 'border-amber-600 text-amber-600'
                        : 'border-transparent text-slate-500 hover:text-slate-800'
                    }`}
                  >
                    <FileText className="w-4 h-4" />
                    版本 A：含小計對帳表 (僅供原廠核帳)
                  </button>
                  <button
                    onClick={() => setActiveTab('raw')}
                    className={`py-3 px-4 text-sm font-semibold border-b-2 transition-all flex items-center gap-2 ${
                      activeTab === 'raw'
                        ? 'border-slate-800 text-slate-800'
                        : 'border-transparent text-slate-500 hover:text-slate-800'
                    }`}
                  >
                    <Info className="w-4 h-4" />
                    原始前 25 筆數據 (對照原檔)
                  </button>
                </div>

                <div className="text-xs text-slate-400 pb-3 md:pb-0">
                  顯示前 15 筆，確認向下填滿與去小計之邏輯
                </div>
              </div>

              {/* 預覽網格 */}
              <div className="overflow-x-auto">
                {activeTab === 'raw' ? (
                  <table className="w-full text-left border-collapse text-sm">
                    <thead>
                      <tr className="bg-slate-100 text-slate-500 border-b border-slate-200">
                        {rawPreview[4] && rawPreview[4].map((h, i) => (
                          <th key={i} className="py-2.5 px-4 font-semibold">{String(h || `欄位${i}`)}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100 text-slate-500 font-mono">
                      {rawPreview.slice(0, 15).map((row, rIdx) => (
                        <tr key={rIdx} className="hover:bg-slate-50/50">
                          {row.map((cell, cIdx) => (
                            <td key={cIdx} className="py-2 px-4 whitespace-nowrap max-w-xs overflow-hidden text-ellipsis">
                              {String(cell || '')}
                            </td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                ) : (
                  <table className="w-full text-left border-collapse text-sm">
                    <thead>
                      <tr className="bg-slate-100 text-slate-600 border-b border-slate-200">
                        <th className="py-3 px-4 font-semibold text-slate-400 w-16 text-center">原列號</th>
                        {headers.map((h, i) => (
                          <th key={i} className="py-3 px-4 font-semibold">{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {(activeTab === 'versionB' ? versionB : versionA).slice(0, 15).map((row, rIdx) => {
                        const isSub = row.isSubtotal;
                        return (
                          <tr 
                            key={rIdx} 
                            className={`hover:bg-slate-50/50 transition-colors ${
                              isSub 
                                ? row.subtotalType === '總計' 
                                  ? 'bg-red-50 text-red-900 font-bold' 
                                  : 'bg-amber-50/60 text-amber-900 font-semibold'
                                : ''
                            }`}
                          >
                            <td className="py-2.5 px-4 text-center font-mono text-slate-400 text-xs border-r border-slate-100">
                              {row.originalRowIndex + 1}
                            </td>
                            {row.cells.map((cell, cIdx) => {
                              const isFilled = row.filledDownFlags[cIdx];
                              return (
                                <td key={cIdx} className="py-2.5 px-4 whitespace-nowrap">
                                  {isFilled ? (
                                    <span className="inline-flex items-center gap-1.5 bg-blue-50 text-blue-800 text-xs px-2 py-0.5 rounded border border-blue-100 font-medium">
                                      <span className="text-blue-500 font-bold">↓</span> {String(cell)}
                                    </span>
                                  ) : (
                                    <span>
                                      {typeof cell === 'number' 
                                        ? cell.toLocaleString() 
                                        : String(cell)}
                                    </span>
                                  )}
                                </td>
                              );
                            })}
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                )}
              </div>
            </div>

            {/* 下方下載與匯出按鈕區 */}
            <div className="flex flex-col sm:flex-row justify-center items-center gap-4 py-4">
              <button
                onClick={() => handleDownload('B')}
                className="w-full sm:w-auto px-8 py-4 bg-emerald-600 hover:bg-emerald-700 text-white font-semibold rounded-xl shadow-md hover:shadow-lg transition-all duration-150 flex items-center justify-center gap-3 text-lg"
              >
                <Download className="w-5 h-5" />
                下載 版本 B（含 KPI 驗證與 Data 明細）
                <span className="bg-emerald-800 text-emerald-100 text-xs px-2 py-0.5 rounded">推薦 🌟</span>
              </button>

              <button
                onClick={() => handleDownload('A')}
                className="w-full sm:w-auto px-8 py-4 bg-slate-800 hover:bg-slate-900 text-slate-100 font-semibold rounded-xl border border-slate-700 hover:border-slate-800 shadow transition-all duration-150 flex items-center justify-center gap-3 text-lg"
              >
                <Download className="w-5 h-5" />
                下載 版本 A（含小計對帳表）
                <span className="bg-slate-700 text-slate-200 text-xs px-2 py-0.5 rounded">對帳用</span>
              </button>
            </div>

          </div>
        )}

        {/* 費曼學習教室：Gigi 助教的職場課 */}
        <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
          <div className="bg-gradient-to-r from-indigo-900 to-indigo-950 text-white py-4 px-6 flex items-center gap-2">
            <BookOpen className="w-5 h-5 text-indigo-400" />
            <h3 className="font-semibold text-lg">Gigi 的「費曼學習教室」：一秒搞懂報表二次分析</h3>
          </div>
          
          <div className="border-b border-slate-200 bg-slate-50 flex">
            <button 
              onClick={() => setActiveFeynmanTab('concept')}
              className={`py-3 px-6 text-sm font-medium border-b-2 transition-all ${
                activeFeynmanTab === 'concept'
                  ? 'border-indigo-600 text-indigo-600 bg-white'
                  : 'text-slate-500 hover:text-slate-800'
              }`}
            >
              費曼學習：為什麼樞紐討厭小計與合併單格？
            </button>
            <button 
              onClick={() => setActiveFeynmanTab('scenario')}
              className={`py-3 px-6 text-sm font-medium border-b-2 transition-all ${
                activeFeynmanTab === 'scenario'
                  ? 'border-indigo-600 text-indigo-600 bg-white'
                  : 'text-slate-500 hover:text-slate-800'
              }`}
            >
              職場實務：這能幫我省下多少時間？
            </button>
            <button 
              onClick={() => setActiveFeynmanTab('lab')}
              className={`py-3 px-6 text-sm font-medium border-b-2 transition-all ${
                activeFeynmanTab === 'lab'
                  ? 'border-indigo-600 text-indigo-600 bg-white'
                  : 'text-slate-500 hover:text-slate-800'
              }`}
            >
              實作練習 Lab：自己動手做做看
            </button>
          </div>

          <div className="p-6 space-y-6">
            {activeFeynmanTab === 'concept' && (
              <div className="space-y-4 leading-relaxed">
                <div className="bg-indigo-50 border-l-4 border-indigo-600 p-4 rounded-r-lg">
                  <h4 className="font-bold text-indigo-900">1. 概念核心解構：為什麼要清洗？</h4>
                  <p className="text-sm text-slate-700 mt-1">
                    給人看的報表就像是<strong className="text-indigo-800">「擺盤精緻的便當」</strong>，飯、菜、肉都放在格子裡；但給電腦（樞紐分析）分析的數據必須是<strong className="text-indigo-800">「火鍋配料盆」</strong>——必須是一列列結構完全一樣、沒有空位、可以隨時自由排列組合的「純明細數據 (Raw Data)」。
                  </p>
                </div>
                
                <h4 className="font-bold text-slate-800 mt-4">2. 實物生活類比</h4>
                <p className="text-sm text-slate-600">
                  想像你去大賣場買東西，明細單上如果寫著：
                  <span className="block bg-slate-50 p-2 font-mono text-xs text-slate-500 rounded my-2 border border-slate-100">
                    2026/04/15 台北店 <br />
                    - 蘋果 10 元 <br />
                    - 香蕉 15 元 <br />
                    - 水果小計：25 元
                  </span>
                  如果你把「水果小計」也放進去加總，賣場最後統計總銷售時就會重複加上那 25 元，導致帳目完全翻倍！這就是為什麼二次分析時，我們必須把所有「小計列、合計列、空白間隔」通通挖掉的原因。
                </p>

                <h4 className="font-bold text-slate-800 mt-4">3. 常見三大清洗地雷</h4>
                <ul className="list-disc pl-5 text-sm text-slate-600 space-y-2">
                  <li>
                    <strong className="text-red-600">地雷一：過度填滿 (Over-filling)。</strong>
                    隨意將「品項代號」或「單價」也向下填滿，導致下一列空欄位被填入錯誤的單價。
                  </li>
                  <li>
                    <strong className="text-red-600">地雷二：漏除空格。</strong>
                    有些 ERP 在人名或客戶名中會留有半形空白，沒做 Trim() 會讓 `"林業務"` 與 `"林業務 "` 在樞紐中變成兩個人。
                  </li>
                  <li>
                    <strong className="text-red-600">地雷三：字串加總。</strong>
                    金額欄位在 Excel 中如果維持文字型態（可能帶著 `$` 符號），樞紐分析將只會「計算個數」而無法進行「SUM 數值求和」。
                  </li>
                </ul>
              </div>
            )}

            {activeFeynmanTab === 'scenario' && (
              <div className="space-y-4 text-sm text-slate-600 leading-relaxed">
                <h4 className="font-bold text-slate-800 text-base">日常場景：苦命助理與主管的交叉分析要求</h4>
                <p>
                  每週五，主管 Gino 都會跟你要一份「各客戶在各區域的銷售貢獻與毛利趨勢分析」。
                  如果你的 ERP 報表是「壞報表」，你過去可能必須：
                </p>
                <ol className="list-decimal pl-5 space-y-1 text-slate-600">
                  <li>手動刪除每一頁導出時殘留的抬頭非數據列。</li>
                  <li>手動複製「客戶名稱」，在幾百列空白儲存格裡按 `Ctrl+D`（向下填滿）。</li>
                  <li>一列一列手動刪除「小計」跟「合計」列。</li>
                  <li>好不容易做完樞紐，卻發現加總出來的金額跟原廠報表對不上，不知道是手動複製錯了，還是哪裡多刪了明細。</li>
                </ol>
                <div className="bg-emerald-50 border-l-4 border-emerald-500 p-4 rounded-r-lg mt-4">
                  <h5 className="font-bold text-emerald-900">💡 使用 Gigi 的清洗工具：</h5>
                  <p className="text-emerald-800 mt-1">
                    只需要一秒鐘！我們同時提供「對帳版本」與「樞紐專用明細」，透過對帳狀態機自動核實原廠總金額。原本要耗費 3 小時的手工勞動，現在僅需 3 秒鐘即可完美匯入 Excel 樞紐分析。
                  </p>
                </div>
              </div>
            )}

            {activeFeynmanTab === 'lab' && (
              <div className="space-y-4 text-sm text-slate-600 leading-relaxed">
                <h4 className="font-bold text-slate-800 text-base">Lab 實作演練：驗證 ERP 報表是否藏有計算 Bug</h4>
                <p>
                  <strong>情境：</strong>
                  在我們提供的「內建銷售日報範例」中，業務 B002 王業務底下有一列客戶 20 的明細（Row 17），金額為 $3,200。
                </p>
                <div className="bg-slate-50 p-4 rounded-lg border border-slate-200">
                  <h5 className="font-bold text-slate-800">🔬 步驟引導：</h5>
                  <ul className="list-decimal pl-5 space-y-2 mt-2">
                    <li>點選上方「載入內建壞報表範例」按鈕。</li>
                    <li>觀察對帳看板：此時會顯示 <strong className="text-amber-700">「發現金額落差」</strong>，原檔總計為 $254,700，但明細加總卻為 $249,500，中間存在 $5,200 差額。</li>
                    <li>切換到「版本 A：含小計對帳表」，對比「客戶20 小計」金額：ERP 自帶的小計為 $99,700。</li>
                    <li>手動加總客戶 20 的所有明細：$22,000 + $3,200 + $13,200 + $30,000 + $34,500 = $102,900。</li>
                    <li><strong>發現結論：</strong>原 ERP 系統的小計竟然少算了一筆 $3,200 的交易！這證明了原廠系統寫入與導出時常有漏洞，利用清洗器建立無小計的標準表，才是最精確的資料檢核之道。</li>
                  </ul>
                </div>
              </div>
            )}
          </div>
        </div>

        {/* 偵錯支援表單 (Debug Support) */}
        <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
          <button 
            onClick={() => setIsDebugOpen(!isDebugOpen)}
            className="w-full bg-slate-100 py-3.5 px-6 flex justify-between items-center text-slate-700 hover:bg-slate-200 transition-colors"
          >
            <div className="flex items-center gap-2 font-semibold">
              <AlertCircle className="w-5 h-5 text-indigo-600" />
              <span>遇到報表「出錯、跑不出來、格式跑掉」？點此查看 Gigi 偵錯指南</span>
            </div>
            {isDebugOpen ? <ChevronUp className="w-5 h-5" /> : <ChevronDown className="w-5 h-5" />}
          </button>
          
          {isDebugOpen && (
            <div className="p-6 space-y-4 text-sm animate-slideDown">
              <h4 className="font-bold text-slate-800">故障排查與可能原因排序：</h4>
              <div className="overflow-x-auto">
                <table className="w-full border-collapse text-left text-xs md:text-sm">
                  <thead>
                    <tr className="bg-slate-50 text-slate-600 border-b border-slate-200">
                      <th className="py-2.5 px-4 font-semibold">優先順序</th>
                      <th className="py-2.5 px-4 font-semibold">可能原因</th>
                      <th className="py-2.5 px-4 font-semibold">判斷依據</th>
                      <th className="py-2.5 px-4 font-semibold">Gigi 建議修正方式</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100 text-slate-600">
                    <tr>
                      <td className="py-3 px-4 font-bold text-indigo-600">1</td>
                      <td className="py-3 px-4 font-semibold text-slate-800">日期欄位變成亂碼或空白</td>
                      <td className="py-3 px-4">原始 Excel 儲存格被設定為特殊自訂格式，或存成非 M365 相容日期。</td>
                      <td className="py-3 px-4">清洗器已內建 Excel 序號轉換演算法，若仍出現空白，請確保在 Excel 中將此直欄設定為「標準簡短日期」再行上傳。</td>
                    </tr>
                    <tr>
                      <td className="py-3 px-4 font-bold text-indigo-600">2</td>
                      <td className="py-3 px-4 font-semibold text-slate-800">小計列沒有被完全清除</td>
                      <td className="py-3 px-4">ERP 輸出的小計不叫「小計」，叫「SubTotal」或「客戶結算」。</td>
                      <td className="py-3 px-4">清洗器的篩選引擎已對「小計、合計、總計、total」進行全不分大小寫匹配。若仍有殘留，請點擊聯絡 Gigi 擴充特徵字庫。</td>
                    </tr>
                    <tr>
                      <td className="py-3 px-4 font-bold text-indigo-600">3</td>
                      <td className="py-3 px-4 font-semibold text-slate-800">向下填滿填錯欄位</td>
                      <td className="py-3 px-4">產品名稱或數量也莫名其妙被向下複製了。</td>
                      <td className="py-3 px-4">清洗器嚴格限定「業務、日期、客戶、區域」等四個分類維度進行 Down-filling，不會汙染任何交易數字與金額欄位。</td>
                    </tr>
                  </tbody>
                </table>
              </div>
            </div>
          )}
        </div>

      </main>

      {/* 頁腳與 AI 風險提示 */}
      <footer className="bg-slate-900 text-slate-400 py-8 px-6 mt-12 border-t border-slate-800">
        <div className="max-w-7xl mx-auto flex flex-col md:flex-row justify-between items-center gap-4 text-xs">
          <p>© 2026 Gigi 銷售日報清洗助理. All rights reserved.</p>
          <div className="bg-slate-800 text-amber-200 border border-amber-800/50 p-3 rounded-lg max-w-lg leading-normal">
            ⚠️ <strong className="text-amber-100">Gigi 助教提醒</strong>：AI 產生的公式與程式碼請務必人工抽樣測試 (比對筆數、金額)。請確認未上傳真實且未經匿名化的機密個資或商業數據。
          </div>
        </div>
      </footer>
    </div>
  );
}
