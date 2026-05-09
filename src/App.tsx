/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useCallback, useMemo } from 'react';
import { 
  FileText, 
  FileSpreadsheet, 
  Upload, 
  CheckCircle2, 
  AlertCircle, 
  RefreshCcw, 
  ArrowRightLeft,
  Search,
  Eye,
  FileSearch,
  Info,
  History,
  ChevronRight,
  Clock,
  Package,
  FolderOpen,
  FileArchive
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import * as mammoth from 'mammoth';
import * as XLSX from 'xlsx';
import JSZip from 'jszip';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';

// --- Utilities ---
function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

// --- Types ---
type FileType = 'audit_report' | 'shareholder_info' | 'net_assets' | 'profit' | 'balance' | 'trial_balance';

interface FileState {
  file: File | null;
  status: 'idle' | 'parsing' | 'ready' | 'error';
  data: any;
  error?: string;
}

interface ComparisonResult {
  field: string;
  wordValue: string | number;
  excelValue: string | number;
  isMatch: boolean;
  sourceFile: string;
  confidence: number;
}

const REQUIRED_EXCELS: { key: FileType; label: string }[] = [
  { key: 'net_assets', label: '净资产变动表' },
  { key: 'profit', label: '利润表' },
  { key: 'balance', label: '资产负债表' },
  { key: 'trial_balance', label: '余额表' },
];

interface BatchProduct {
  id: string;
  name: string;
  files: Record<FileType, FileState>;
  isReady: boolean;
  history: HistoryItem[];
}

interface HistoryItem {
  id: string;
  productName: string;
  compareTime: string;
  mainFile: string;
  excelFiles: string[];
  status: '正常' | '有差异';
  diffCount: number;
}

const MOCK_HISTORY = [
  {
    id: 'H001',
    productName: '某量化策略1号全周期进取型资产管理计划',
    compareTime: '2026-05-07 10:22',
    mainFile: '2025年度审计报告_最终版.docx',
    excelFiles: ['余额表.xlsx', '净资产变动表.xlsx', '利润表.xlsx'],
    status: '有差异',
    diffCount: 3
  },
  {
    id: 'H002',
    productName: '睿智增长集合资产管理计划01期',
    compareTime: '2026-05-06 15:45',
    mainFile: '2025年度审计报告.docx',
    excelFiles: ['余额表.xlsx', '利润表.xlsx'],
    status: '正常',
    diffCount: 0
  },
  {
    id: 'H003',
    productName: '稳健型固收+增强私募基金',
    compareTime: '2026-05-05 09:12',
    mainFile: '审计报告（初稿）.docx',
    excelFiles: ['余额表.xlsx'],
    status: '有差异',
    diffCount: 12
  }
];

// --- Static Demo Data ---
const demoNetAssetsExcel = [
  ['项目', '本年实收资本', '本年其他综合收益', '本年未分配利润', '本年净资产合计'],
  ['一、上年年末余额', '34,251,624.09', '', '34,251,624.09', '68,503,248.18'],
  ['加：会计政策变更', '', '', '', ''],
  ['前期差错更正', '', '', '', ''],
  ['其他', '', '', '', ''],
  ['二、本年年初余额', '34,251,624.09', '', '34,251,624.09', '68,503,248.18'],
  ['三、本期增减变动额', '', '', '-17311759.23', '-17311759.23'],
  ['（一）综合收益总额', '', '', '-34,238,021.49', '-34,238,021.49'],
  ['（二）产品持有人申购和赎回', '', '', '51177886.35', '51177886.35'],
  ['其中：产品申购', '34,251,624.09', '', '34253425.18', '68505049.27'],
  ['产品赎回', '-34,251,624.09', '', '16,924,461.17', '-17,327,162.92'],
  ['（三）利润分配', '', '', '-34,251,624.09', '-34,251,624.09'],
  ['四、本期期末余额', '34,251,624.09', '', '16939864.86', '51191488.95']
];

export default function App() {
  const [files, setFiles] = useState<Record<FileType, FileState>>({
    audit_report: { file: null, status: 'idle', data: null },
    shareholder_info: { file: null, status: 'idle', data: null },
    net_assets: { file: null, status: 'idle', data: null },
    profit: { file: null, status: 'idle', data: null },
    balance: { file: null, status: 'idle', data: null },
    trial_balance: { file: null, status: 'idle', data: null },
  });

  const [batchProducts, setBatchProducts] = useState<BatchProduct[]>([]);
  const [uploadMode, setUploadMode] = useState<'batch' | 'manual'>('batch');
  const [globalShareholder, setGlobalShareholder] = useState<FileState>({ file: null, status: 'idle', data: null });
  const [isProcessing, setIsProcessing] = useState(false);
  const [selectedProductId, setSelectedProductId] = useState<string | null>(null);
  const [previewContent, setPreviewContent] = useState<{ type: 'word' | 'excel'; title: string; data: any } | null>(null);
  const [selectedItems, setSelectedItems] = useState<string[]>([]);

  const [view, setView] = useState<'upload' | 'comparison' | 'history'>('upload');
  const [activeTab, setActiveTab] = useState<string>('net_assets');
  const [wordHtml, setWordHtml] = useState<string>('');

  // --- Demo Data Support ---
  const loadDemoData = () => {
    // Audit Report Data (Image 1)
    const demoWord = `
      <h3>七、财务报表主要项目附注</h3>
      <h4>（一）银行存款</h4>
      <table>
        <tr><th>项目</th><th>2025年12月31日</th><th>2024年12月31日</th></tr>
        <tr><td>活期存款</td><td>3,114,074.61</td><td>7,969,396.39</td></tr>
        <tr><td>应计利息</td><td>1,423.59</td><td>47,580.14</td></tr>
        <tr style="font-weight:bold;"><td>合计</td><td>3,115,498.20</td><td>8,016,976.53</td></tr>
      </table>
      <h4>（二）交易性金融资产</h4>
      <table>
        <tr><th rowspan="2">项目</th><th colspan="4">2024年12月31日</th></tr>
        <tr><th>成本</th><th>应计利息</th><th>公允价值</th><th>公允价值变动</th></tr>
        <tr><td>债券</td><td>50,000,000.00</td><td>266,561.65</td><td>52,410,000.00</td><td>2,410,000.00</td></tr>
        <tr style="font-weight:bold;"><td>合计</td><td>50,000,000.00</td><td>266,561.65</td><td>52,410,000.00</td><td>2,410,000.00</td></tr>
      </table>
    `;

    // Net Asset Change Demo Table (Matching the new screenshot)
    const demoNetAssetsWord = JSON.parse(JSON.stringify(demoNetAssetsExcel));
    // Simulate some differences in Word data
    demoNetAssetsWord[6][3] = '-85803205.9'; // Explicit mismatch
    demoNetAssetsWord[8][3] = '-17313560.32'; // Explicit mismatch
    demoNetAssetsWord[12][3] = '-51551581.81'; // Explicit mismatch

    setWordHtml(demoWord);
    setFiles(prev => ({
      ...prev,
      audit_report: { file: new File([], "演示报表.docx"), status: 'ready', data: demoNetAssetsWord },
      trial_balance: { file: new File([], "科目余额表.xlsx"), status: 'ready', data: [] },
      shareholder_info: { file: new File([], "demo.xlsx"), status: 'ready', data: [] },
      net_assets: { file: new File([], "净资产变动表.xlsx"), status: 'ready', data: demoNetAssetsExcel },
      profit: { file: new File([], "利润表.xlsx"), status: 'ready', data: [] },
      balance: { file: new File([], "资产负债表.xlsx"), status: 'ready', data: [] },
    }));
    setView('comparison');
  };

  // --- Export Logic ---
  const exportToCSV = (productName: string, data: any[][]) => {
    const csvContent = data.map(row => row.join(',')).join('\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `${productName}_比对结果.csv`;
    link.click();
  };

  const batchExport = () => {
    if (selectedItems.length === 0) {
      alert('请先选择要导出的记录');
      return;
    }
    // In a real app, this might create a ZIP or multiple files
    alert(`批量导出成功！已处理 ${selectedItems.length} 个产品。`);
  };

  const toggleSelection = (id: string) => {
    setSelectedItems(prev => 
      prev.includes(id) ? prev.filter(i => i !== id) : [...prev, id]
    );
  };

  const handleManualFileUpload = async (type: FileType, file: File) => {
    try {
      const data = await parseFileContent(type, file);
      setFiles(prev => ({
        ...prev,
        [type]: { file, status: 'ready', data }
      }));
    } catch (err) {
      setFiles(prev => ({
        ...prev,
        [type]: { file, status: 'error', data: null, error: '解析失败' }
      }));
    }
  };

  const startManualComparison = () => {
    const manualProduct: BatchProduct = {
      id: 'manual_prod_' + Date.now(),
      name: files.audit_report.file?.name.replace('.docx', '') || '单项上传产品',
      files: { ...files },
      isReady: files.audit_report.status === 'ready' && 
                (files.net_assets.status === 'ready' || files.balance.status === 'ready'),
      history: []
    };
    setBatchProducts([manualProduct]);
    setSelectedProductId(manualProduct.id);
    setView('comparison');
  };

  // --- Handlers ---
  const handleGlobalShareholderUpload = async (file: File) => {
    setGlobalShareholder({ file, status: 'parsing', data: null });
    try {
      const arrayBuffer = await file.arrayBuffer();
      const workbook = XLSX.read(arrayBuffer, { type: 'array' });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
      setGlobalShareholder({ file, status: 'ready', data: jsonData });
    } catch (error) {
      setGlobalShareholder({ file, status: 'error', data: null, error: '解析失败' });
    }
  };

  const parseFileContent = async (type: FileType, file: File): Promise<any> => {
    const arrayBuffer = await file.arrayBuffer();
    if (type === 'audit_report') {
      const result = await mammoth.convertToHtml({ arrayBuffer });
      // In a real app we'd extract specific tables here
      return result.value; 
    } else {
      const workbook = XLSX.read(arrayBuffer, { type: 'array' });
      return XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 1 });
    }
  };

  const handleBatchUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const uploadedFiles = e.target.files;
    if (!uploadedFiles || uploadedFiles.length === 0) return;

    setIsProcessing(true);
    const newProducts: Record<string, BatchProduct> = {};

    const processSingleFile = async (path: string, file: File) => {
      const parts = path.split('/');
      // If it's something like "ProductA/Audit.docx" or "root/ProductA/Audit.docx"
      const productName = parts.length > 1 ? parts[parts.length - 2] : '未命名产品';
      const fileName = file.name.toLowerCase();

      if (!newProducts[productName]) {
        newProducts[productName] = {
          id: Math.random().toString(36).substr(2, 9),
          name: productName,
          files: {
            audit_report: { file: null, status: 'idle', data: null },
            shareholder_info: { file: null, status: 'idle', data: null },
            net_assets: { file: null, status: 'idle', data: null },
            profit: { file: null, status: 'idle', data: null },
            balance: { file: null, status: 'idle', data: null },
            trial_balance: { file: null, status: 'idle', data: null },
          },
          isReady: false,
          history: []
        };
      }

      let type: FileType | null = null;
      if (fileName.includes('审计报告') || fileName.endsWith('.docx')) type = 'audit_report';
      else if (fileName.includes('净资产') || fileName.includes('净值变动')) type = 'net_assets';
      else if (fileName.includes('利润')) type = 'profit';
      else if (fileName.includes('资产负债')) type = 'balance';
      else if (fileName.includes('余额表')) type = 'trial_balance';

      if (type) {
        try {
          const data = await parseFileContent(type, file);
          newProducts[productName].files[type] = { file, status: 'ready', data };
        } catch (err) {
          newProducts[productName].files[type] = { file, status: 'error', data: null };
        }
      }
    };

    const firstFile = uploadedFiles[0];
    if (firstFile.name.endsWith('.zip')) {
      try {
        const zip = await JSZip.loadAsync(firstFile);
        const entries = Object.values(zip.files).filter(f => !f.dir);
        for (const entry of entries) {
          const blob = await entry.async('blob');
          const file = new File([blob], entry.name.split('/').pop() || 'file');
          await processSingleFile(entry.name, file);
        }
      } catch (err) {
        console.error("ZIP Error", err);
      }
    } else {
      for (let i = 0; i < uploadedFiles.length; i++) {
        const file = uploadedFiles[i];
        // WebkitRelativePath preserves folder structure if using "directory" attribute
        await processSingleFile((file as any).webkitRelativePath || file.name, file);
      }
    }

    setBatchProducts(Object.values(newProducts).map(p => {
      // Check if critical files exist
      const hasCritical = p.files.audit_report.status === 'ready' && 
                          (p.files.net_assets.status === 'ready' || p.files.balance.status === 'ready');
      return { ...p, isReady: hasCritical };
    }));
    setIsProcessing(false);
  };

  const isAllFilesReady = useMemo(() => {
    return batchProducts.length > 0 || globalShareholder.status === 'ready';
  }, [batchProducts, globalShareholder]);

  // --- Comparison Logic for side-by-side ---
  const currentProduct = useMemo(() => {
    return batchProducts.find(p => p.id === selectedProductId) || batchProducts[0];
  }, [batchProducts, selectedProductId]);

  const tableData = useMemo(() => {
    let excelData: any[][] = [[]];
    let wordData: any[][] = [[]];
    const currentFiles = currentProduct?.files || files;

    if (activeTab === 'net_assets') {
      excelData = currentFiles.net_assets.data as any[][] || demoNetAssetsExcel;
      wordData = currentFiles.audit_report.data as any[][] || demoNetAssetsExcel;
    } else if (activeTab === 'shareholder_info') {
      excelData = globalShareholder.data as any[][] || [
        ['内部分码', '持有人', '2025年12月31日', '2024年12月31日'],
        ['A01', '某1账户', '992,459,596.35', '992,459,596.35'],
        ['A02', '某2账户', '396,983,838.53', '396,983,838.53'],
        ['', '合计', '1,389,443,434.88', '1,389,443,434.88']
      ];
      wordData = [
        ['持有人', '2025年12月31日', '2024年12月31日'],
        ['某1账户', '992,459,596.35', '992,459,596.35'],
        ['某2账户', '396,983,838.53', '396,983,838.53'],
        ['合计', '1,389,443,434.88', '1,389,443,434.88']
      ];
    } else {
      const customData = currentFiles[activeTab as FileType]?.data;
      excelData = Array.isArray(customData) ? customData : [['暂无匹配系统报表数据', '']];
      wordData = [['暂未从报告中提取到对应表格', '']];
    }

    return { excel: excelData, word: wordData };
  }, [activeTab, currentProduct, files, globalShareholder]);

  // Mock Annotation logic based on image
  const renderAuditDocument = () => {
    if (activeTab === 'shareholder_info') {
      return (
        <div className="space-y-8 animate-in fade-in slide-in-from-bottom-4 duration-500">
          <div className="bg-white p-8 border border-slate-200 shadow-sm relative">
            <p className="text-sm font-medium text-slate-800 mb-6">
              2. 托管人、产品管理人主要股东及其控制的机构所持有的本产品份额
            </p>
            <p className="text-sm text-slate-600 leading-relaxed mb-6">
              于 12 月 31 日，产品的产品托管人未持有本产品份额。产品的产品管理人主要股东及其控制的机构持有本产品份额列示如下：
            </p>
            
            <table className="w-full text-xs border-collapse border border-slate-300">
              <thead>
                <tr className="bg-slate-50">
                  <th className="border border-slate-300 p-2 text-center bg-[#FDE9D9]">持有人</th>
                  <th className="border border-slate-300 p-2 text-center bg-[#FDE9D9]">2025 年 12 月 31 日</th>
                  <th className="border border-slate-300 p-2 text-center bg-[#FDE9D9]">2024 年 12 月 31 日</th>
                </tr>
              </thead>
              <tbody>
                {tableData.word.slice(1).map((row, i) => (
                  <tr key={i}>
                    {row.map((cell: any, ci: number) => (
                      <td key={ci} className="border border-slate-300 p-2 text-right">{cell}</td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>

            {/* Annotation Badge */}
            <div className="absolute -right-4 top-1/2 translate-x-full flex items-start gap-2 max-w-[150px]">
              <div className="w-8 h-8 rounded bg-red-500 flex items-center justify-center shrink-0">
                <FileSearch className="text-white w-5 h-5" />
              </div>
              <div className="pt-1">
                <div className="text-xs font-bold text-red-500">Administrator</div>
                <div className="text-[10px] text-slate-500">与 TA 数据核对</div>
              </div>
            </div>
          </div>
        </div>
      );
    }

    return (
      <div className="bg-white p-10 border border-slate-200 shadow-sm relative min-h-full">
        {wordHtml && activeTab === 'net_assets' ? (
          <div className="prose prose-sm max-w-none prose-slate" dangerouslySetInnerHTML={{ __html: wordHtml }} />
        ) : (
          <div className="space-y-6">
             <div className="flex items-center gap-2 mb-4">
               <span className="bg-red-50 border border-red-100 text-red-600 px-2 py-0.5 text-xs font-bold rounded">第 24 页</span>
               <span className="bg-[#FDE9D9] px-2 py-0.5 text-xs font-bold">(二) 产品管理人、托管行报酬</span>
             </div>
             <p className="text-sm font-bold">1. 管理人报酬</p>
             <p className="text-sm leading-7">
               管理费每日按该产品前一日资产净值的 0.15% 的年费率计提。本产品在本年度共计提管理人报酬人民币 
               <span className="bg-yellow-200 px-1 font-bold">3,100,338.78</span> 元（上年度：人民币 3,131,381.61 元），于本年末应付管理人报酬为人民币 
               <span className="bg-yellow-200 px-1 font-bold">794,223.58</span> 元（上年末：人民币 803,589.86 元）。
             </p>
             <p className="text-sm font-bold mt-6">2. 托管费</p>
             <p className="text-sm leading-7">
               托管费每日按该产品前一日资产净值的 0.008% 的年费率计提。本产品在本年度共计提托管费人民币 
               <span className="bg-yellow-200 px-1 font-bold">165,351.41</span> 元（上年度：人民币 167,007.00 元），于本年末应付托管费为人民币 41,903.65 元（上年末：人民币 42,397.30 元）。
             </p>

            {/* Mock Annotation */}
            <div className="absolute -right-4 top-20 translate-x-full flex items-start gap-2 max-w-[150px]">
              <div className="w-8 h-8 rounded bg-red-500 flex items-center justify-center shrink-0">
                <FileSearch className="text-white w-5 h-5" />
              </div>
              <div className="pt-1">
                <div className="text-xs font-bold text-red-500">Administrator</div>
                <div className="text-[10px] text-slate-500">与余额表核对</div>
              </div>
            </div>
          </div>
        )}
      </div>
    );
  };

  return (
    <div className="min-h-screen bg-slate-100 flex flex-col">
      {/* Preview Modal */}
      <AnimatePresence>
        {previewContent && (
          <motion.div 
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 z-[100] flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm"
            onClick={() => setPreviewContent(null)}
          >
            <motion.div 
              initial={{ scale: 0.95, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.95, opacity: 0 }}
              className="bg-white rounded-2xl shadow-2xl w-full max-w-4xl h-[80vh] flex flex-col overflow-hidden"
              onClick={(e) => e.stopPropagation()}
            >
              <div className="p-4 border-b border-slate-100 flex items-center justify-between">
                <h3 className="font-bold text-slate-800">{previewContent.title}</h3>
                <button onClick={() => setPreviewContent(null)} className="p-2 hover:bg-slate-100 rounded-lg">
                  <AlertCircle className="w-5 h-5 text-slate-400 rotate-45" />
                </button>
              </div>
              <div className="flex-1 overflow-auto p-6 bg-slate-50">
                {previewContent.type === 'word' ? (
                  <div className="prose max-w-none prose-slate" dangerouslySetInnerHTML={{ __html: previewContent.data }} />
                ) : (
                  <div className="bg-white border rounded-lg overflow-hidden shadow-sm">
                    <table className="w-full text-xs text-left">
                      <thead className="bg-slate-100 font-bold border-b text-slate-700">
                        <tr>
                          <th className="p-2.5">文件名</th>
                          <th className="p-2.5 text-right">操作</th>
                        </tr>
                      </thead>
                      <tbody>
                        {previewContent.data.slice(1).map((row: any[], i: number) => (
                          <tr key={i} className="border-b last:border-0 hover:bg-slate-50 transition-colors">
                            <td className="p-2.5 text-slate-600 font-medium">{row[0]}</td>
                            <td className="p-2.5 text-right">
                              <button className="text-primary hover:underline font-bold text-xs">查看内容</button>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>
            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>

      {/* Top Navbar */}
      <nav className="bg-slate-800 text-white px-6 py-3 flex items-center justify-between shadow-md">
        <div className="flex items-center gap-2">
          <div className="w-8 h-8 bg-yellow-400 rounded-full flex items-center justify-center">
            <Search className="w-5 h-5 text-slate-800" />
          </div>
          <span className="font-bold tracking-tight">审计数据对比工具</span>
        </div>
        <div className="flex items-center gap-4 text-sm opacity-80">
          <div className="flex items-center gap-1">
            <span className="w-6 h-6 bg-slate-700 rounded-full flex items-center justify-center text-[10px]">admin</span>
            <span>admin</span>
          </div>
        </div>
      </nav>

      {/* Main Container */}
      <div className="flex-1 flex overflow-hidden">
        {/* Sidebar */}
        <aside className="w-64 bg-white border-r border-slate-200 hidden md:flex flex-col">
          <div className="p-4 border-b border-slate-100 font-bold text-slate-400 text-xs uppercase tracking-widest">
            功能菜单
          </div>
          <nav className="p-2 space-y-1">
            <button 
              onClick={() => setView('upload')}
              className={cn(
                "w-full flex items-center gap-3 px-4 py-3 text-sm font-bold rounded-lg transition-colors",
                (view === 'upload' || view === 'comparison') ? "text-primary bg-red-50" : "text-slate-600 hover:bg-slate-50"
              )}
            >
              <FileSearch className="w-5 h-5" />
              产品财务审计
            </button>
            <button 
              onClick={() => setView('history')}
              className={cn(
                "w-full flex items-center gap-3 px-4 py-3 text-sm font-bold rounded-lg transition-colors",
                view === 'history' ? "text-primary bg-red-50" : "text-slate-600 hover:bg-slate-50"
              )}
            >
              <History className="w-5 h-5" />
              产品对比历史
            </button>
          </nav>
        </aside>

        {/* Content Area */}
        <main className="flex-1 overflow-auto bg-slate-50 relative flex flex-col">
          {/* Main Content Tabs */}
          {(view === 'upload' || view === 'comparison') && (
            <div className="bg-white border-b border-slate-200 px-6 flex items-center gap-2">
              <button 
                onClick={() => setView('upload')}
                className={cn(
                  "px-6 py-3 text-sm font-bold border-b-2 transition-all",
                  view === 'upload' ? "border-primary text-primary" : "border-transparent text-slate-400 hover:text-slate-600"
                )}
              >
                数据采集上传
              </button>
              <AnimatePresence>
                {(view === 'comparison' || isAllFilesReady) && (
                  <motion.button 
                    initial={{ opacity: 0, x: -10 }}
                    animate={{ opacity: 1, x: 0 }}
                    onClick={() => setView('comparison')}
                    className={cn(
                      "px-6 py-3 text-sm font-bold border-b-2 transition-all flex items-center gap-2",
                      view === 'comparison' ? "border-primary text-primary" : "border-transparent text-slate-400 hover:text-slate-600"
                    )}
                  >
                    审计比对报告
                    {view === 'comparison' && <span className="w-2 h-2 bg-primary rounded-full animate-pulse" />}
                  </motion.button>
                )}
              </AnimatePresence>
            </div>
          )}

          <div className="flex-1 overflow-auto p-8">
            {view === 'history' ? (
              <div className="max-w-6xl mx-auto space-y-6">
                <header className="mb-8 flex items-end justify-between">
                  <div>
                    <h1 className="text-3xl font-extrabold text-slate-800 tracking-tight">产品对比历史</h1>
                    <p className="text-slate-500">查看并管理已完成的所有审计比对任务记录。</p>
                  </div>
                  <div className="flex gap-3">
                    <button 
                      onClick={batchExport}
                      disabled={selectedItems.length === 0}
                      className={cn(
                        "px-6 py-2.5 rounded-xl font-bold text-sm shadow-sm transition-all border",
                        selectedItems.length > 0
                          ? "bg-primary text-white border-primary hover:shadow-lg"
                          : "bg-slate-100 text-slate-400 border-slate-200 cursor-not-allowed"
                      )}
                    >
                      批量导出当前选择 ({selectedItems.length})
                    </button>
                    <button className="px-6 py-2.5 bg-white text-slate-600 border border-slate-200 rounded-xl font-bold text-sm shadow-sm hover:bg-slate-50">
                      清除历史记录
                    </button>
                  </div>
                </header>

                <div className="bg-white rounded-2xl border border-slate-200 shadow-xl overflow-hidden">
                  <table className="w-full text-left border-collapse">
                    <thead>
                      <tr className="bg-slate-50 border-b border-slate-100">
                        <th className="px-6 py-4 w-10">
                          <input 
                            type="checkbox" 
                            className="rounded accent-primary" 
                            checked={selectedItems.length === (MOCK_HISTORY.length + batchProducts.length) && selectedItems.length > 0}
                            onChange={(e) => {
                              if (e.target.checked) {
                                setSelectedItems([...MOCK_HISTORY.map(h => h.id), ...batchProducts.map(p => p.id)]);
                              } else {
                                setSelectedItems([]);
                              }
                            }}
                          />
                        </th>
                        <th className="px-6 py-4 text-xs font-bold text-slate-400 uppercase tracking-widest">产品名称 / 比对时间</th>
                        <th className="px-6 py-4 text-xs font-bold text-slate-400 uppercase tracking-widest text-center">审计报告</th>
                        <th className="px-6 py-4 text-xs font-bold text-slate-400 uppercase tracking-widest text-center">核对附件数</th>
                        <th className="px-6 py-4 text-xs font-bold text-slate-400 uppercase tracking-widest">比对状态</th>
                        <th className="px-6 py-4 text-xs font-bold text-slate-400 uppercase tracking-widest text-right">操作</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-50">
                      {[...batchProducts.map(p => ({
                        id: p.id,
                        productName: p.name,
                        compareTime: '刚刚',
                        mainFile: p.files.audit_report.file?.name || '未上传报告',
                        excelFiles: Object.keys(p.files).filter(k => k !== 'audit_report' && p.files[k as FileType].status === 'ready' && p.files[k as FileType].file !== null),
                        status: p.isReady ? '正常' : '有差异', // Simple mock status for session products
                        diffCount: 2,
                        source: 'session'
                      })), ...MOCK_HISTORY].sort((a, b) => a.status === '有差异' ? -1 : 1).map((item) => (
                        <tr key={item.id} className={cn(
                          "hover:bg-slate-50/80 transition-colors group",
                          selectedItems.includes(item.id) && "bg-blue-50/30"
                        )}>
                          <td className="px-6 py-4">
                            <input 
                              type="checkbox" 
                              className="rounded accent-primary"
                              checked={selectedItems.includes(item.id)}
                              onChange={() => toggleSelection(item.id)}
                            />
                          </td>
                          <td className="px-6 py-4">
                            <div className="font-bold text-slate-700 mb-1">{item.productName}</div>
                            <div className="text-xs text-slate-400 flex items-center gap-1">
                              <Clock className="w-3 h-3" />
                              {item.compareTime}
                            </div>
                          </td>
                          <td className="px-6 py-4 text-center">
                            <button 
                              onClick={() => setPreviewContent({ type: 'word', title: item.mainFile, data: wordHtml || '正在解析内容...' })}
                              className="inline-flex items-center gap-2 text-sm text-slate-500 bg-slate-100 px-3 py-1 rounded-md hover:bg-red-50 hover:text-primary transition-colors border border-transparent hover:border-red-100"
                            >
                              <FileText className="w-4 h-4 text-primary" />
                              {item.mainFile.length > 15 ? item.mainFile.substring(0, 15) + '...' : item.mainFile}
                            </button>
                          </td>
                          <td className="px-6 py-4 text-center">
                            <button 
                              onClick={() => setPreviewContent({ type: 'excel', title: '附件列表', data: [['文件名', '状态'], ...item.excelFiles.map(f => [f, '就绪']), ['份额表.xlsx', '引用全局']] })}
                              className="text-sm font-bold text-slate-600 bg-slate-100 w-10 h-10 inline-flex items-center justify-center rounded-full hover:bg-blue-100 hover:text-blue-600 transition-colors"
                            >
                              {item.excelFiles.length + 1}
                            </button>
                          </td>
                          <td className="px-6 py-4">
                            <div 
                              onClick={() => {
                                setSelectedProductId(item.id);
                                setView('comparison');
                              }}
                              className={cn(
                                "inline-flex items-center gap-1.5 px-3 py-1 rounded-full text-xs font-bold cursor-pointer hover:shadow-sm border transition-all",
                                item.status === '正常' 
                                  ? "bg-green-100 text-green-700 border-green-200 hover:bg-green-200" 
                                  : "bg-red-100 text-red-700 border-red-200 hover:bg-red-200"
                              )}
                            >
                              {item.status === '正常' ? <CheckCircle2 className="w-3.5 h-3.5" /> : <AlertCircle className="w-3.5 h-3.5" />}
                              {item.status === '正常' ? '核对一致' : `异常 (${item.diffCount})`}
                            </div>
                          </td>
                          <td className="px-6 py-4 text-right">
                            <div className="flex items-center justify-end gap-3">
                              <button 
                                onClick={() => {
                                  setSelectedProductId(item.id);
                                  setView('comparison');
                                }}
                                className="inline-flex items-center gap-1 text-sm font-bold text-primary hover:underline hover:text-primary/80"
                              >
                                详情
                              </button>
                              <div className="w-px h-3 bg-slate-200" />
                              <button 
                                onClick={() => exportToCSV(item.productName, tableData.excel)}
                                className="text-sm font-bold text-blue-600 hover:underline"
                              >
                                导出
                              </button>
                              <div className="w-px h-3 bg-slate-200" />
                              <button 
                                onClick={() => {
                                  setSelectedProductId(item.id);
                                  setView('upload');
                                }}
                                className="text-sm font-bold text-slate-400 hover:text-slate-600"
                              >
                                重新上传
                              </button>
                            </div>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                  <div className="p-4 bg-white border-t border-slate-100 text-center">
                    <p className="text-xs text-slate-400 font-medium">显示最近 20 条比对记录</p>
                  </div>
                </div>
              </div>
            ) : view === 'upload' ? (
              <div className="max-w-5xl mx-auto space-y-10">
                <header className="text-center md:text-left mb-12 flex flex-col md:flex-row md:items-end justify-between gap-6">
                  <div>
                    <h1 className="text-4xl font-extrabold text-slate-900 tracking-tight mb-2">审计报告核对工具</h1>
                    <p className="text-slate-500 text-lg font-medium">智能比对财务报表与审计说明的一致性，支持批量自动化处理。</p>
                  </div>
                  <div className="flex bg-slate-200 p-1 rounded-2xl self-center md:self-auto">
                    <button 
                      onClick={() => setUploadMode('batch')}
                      className={cn(
                        "px-6 py-2 rounded-xl text-sm font-black transition-all",
                        uploadMode === 'batch' ? "bg-white text-primary shadow-sm" : "text-slate-500 hover:text-slate-700"
                      )}
                    >
                      批量/ZIP上传
                    </button>
                    <button 
                      onClick={() => setUploadMode('manual')}
                      className={cn(
                        "px-6 py-2 rounded-xl text-sm font-black transition-all",
                        uploadMode === 'manual' ? "bg-white text-primary shadow-sm" : "text-slate-500 hover:text-slate-700"
                      )}
                    >
                      单产品手动上传
                    </button>
                  </div>
                </header>

                {selectedProductId && batchProducts.find(p => p.id === selectedProductId) && (
                  <motion.div 
                    initial={{ opacity: 0, y: -20 }}
                    animate={{ opacity: 1, y: 0 }}
                    className="bg-primary/5 border border-primary/20 rounded-2xl p-6 mb-8 flex items-center justify-between"
                  >
                    <div className="flex items-center gap-4">
                      <div className="p-3 bg-primary text-white rounded-xl shadow-lg">
                        <FolderOpen className="w-6 h-6" />
                      </div>
                      <div>
                        <h2 className="text-lg font-bold text-slate-800">正在修改：{batchProducts.find(p => p.id === selectedProductId)?.name}</h2>
                        <p className="text-sm text-slate-500 font-medium">您可以单独选择文件进行覆盖上传</p>
                      </div>
                    </div>
                    <button 
                      onClick={() => setSelectedProductId(null)}
                      className="px-4 py-2 text-sm font-bold text-slate-400 hover:text-slate-600"
                    >
                      取消修改模式
                    </button>
                  </motion.div>
                )}

                <div className="grid grid-cols-1 md:grid-cols-3 gap-8">
                  {uploadMode === 'manual' ? (
                    <section className="md:col-span-2 bg-white p-8 rounded-3xl border border-slate-200 shadow-xl shadow-slate-200/50 flex flex-col">
                      <div className="mb-6">
                        <h2 className="text-xl font-bold flex items-center gap-3 text-slate-800">
                          <div className="p-2 bg-primary/10 rounded-lg"><FileText className="w-6 h-6 text-primary" /></div>
                          单产品文件选取
                        </h2>
                        <p className="mt-2 text-sm text-slate-400 font-medium ml-11">
                          请手动上传该产品的审计报告及其对应的一系列财务报表 Excel 文件。
                        </p>
                      </div>
                      
                      <div className="space-y-6">
                        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                          <div className={cn(
                            "relative border-2 border-dashed rounded-2xl p-6 transition-all",
                            files.audit_report.status === 'ready' ? "bg-green-50 border-green-200" : "bg-slate-50 border-slate-200 hover:bg-slate-100"
                          )}>
                            <input 
                              type="file" accept=".docx" 
                              className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10"
                              onChange={(e) => e.target.files?.[0] && handleManualFileUpload('audit_report', e.target.files[0])}
                            />
                            <div className="flex items-center gap-4">
                              {files.audit_report.status === 'ready' ? <CheckCircle2 className="text-green-500 w-8 h-8" /> : <Upload className="text-slate-400 w-8 h-8" />}
                              <div className="overflow-hidden">
                                <div className="font-bold text-sm text-slate-700">审计报告 (Word)</div>
                                <div className="text-xs text-slate-400 truncate">{files.audit_report.file?.name || '点击上传 .docx'}</div>
                              </div>
                            </div>
                          </div>
                          {REQUIRED_EXCELS.slice(0, 5).map(item => (
                            <div key={item.key} className={cn(
                              "relative border-2 border-dashed rounded-2xl p-4 transition-all",
                              files[item.key].status === 'ready' ? "bg-green-50 border-green-200" : "bg-slate-50 border-slate-200 hover:bg-slate-100"
                            )}>
                              <input 
                                type="file" accept=".xlsx,.xls" 
                                className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10"
                                onChange={(e) => e.target.files?.[0] && handleManualFileUpload(item.key, e.target.files[0])}
                              />
                              <div className="flex items-center gap-3">
                                {files[item.key].status === 'ready' ? <CheckCircle2 className="text-green-500 w-5 h-5" /> : <FileSpreadsheet className="text-slate-300 w-5 h-5" />}
                                <div className="overflow-hidden">
                                  <div className="font-bold text-xs text-slate-700">{item.label}</div>
                                  <div className="text-[10px] text-slate-400 truncate">{files[item.key].file?.name || '点击上传 Excel'}</div>
                                </div>
                              </div>
                            </div>
                          ))}
                        </div>
                      </div>
                    </section>
                  ) : (
                    <>
                      {/* Left: Batch Upload */}
                      <section className="md:col-span-2 bg-white p-8 rounded-3xl border border-slate-200 shadow-xl shadow-slate-200/50 flex flex-col">
                    <div className="mb-6 flex items-start justify-between">
                      <div className="space-y-1">
                        <h2 className="text-xl font-bold flex items-center gap-3 text-slate-800">
                          <div className="p-2 bg-blue-100 rounded-lg"><Package className="w-6 h-6 text-blue-600" /></div>
                          批量目录/压缩包上传
                        </h2>
                        <p className="text-sm text-slate-400 font-medium ml-11">
                          支持将包含多个产品的根目录直接拖入，或上传打包好的 ZIP 压缩包。系统将自动根据子目录名称识别产品及对应的报表文件。
                        </p>
                      </div>
                      {batchProducts.length > 0 && (
                        <span className="text-xs bg-blue-50 text-blue-600 px-3 py-1.5 rounded-full font-bold">已识别 {batchProducts.length} 个产品</span>
                      )}
                    </div>
                    
                    <div className="relative flex-1 min-h-[300px]">
                      <input 
                        type="file" 
                        multiple
                        // @ts-ignore
                        webkitdirectory="true"
                        className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10"
                        onChange={handleBatchUpload}
                      />
                      <input 
                        type="file" 
                        accept=".zip"
                        className="hidden" 
                        id="zip-upload"
                        onChange={handleBatchUpload}
                      />
                      
                      <div className={cn(
                        "h-full border-2 border-dashed rounded-2xl p-8 flex flex-col items-center justify-center transition-all duration-300",
                        batchProducts.length > 0 
                          ? "border-blue-200 bg-blue-50" 
                          : "border-slate-200 hover:border-primary/40 hover:bg-slate-50"
                      )}>
                        {isProcessing ? (
                          <div className="flex flex-col items-center gap-4">
                            <RefreshCcw className="w-12 h-12 text-blue-500 animate-spin" />
                            <span className="text-lg font-bold text-slate-600">正在解析目录结构...</span>
                          </div>
                        ) : batchProducts.length > 0 ? (
                          <div className="w-full h-full overflow-auto space-y-2">
                            {batchProducts.map(p => (
                              <div key={p.id} className="flex items-center justify-between p-3 bg-white rounded-xl border border-blue-100 shadow-sm">
                                <div className="flex items-center gap-3">
                                  <FolderOpen className="w-5 h-5 text-blue-400" />
                                  <span className="font-bold text-slate-700 text-sm">{p.name}</span>
                                </div>
                                <div className="flex items-center gap-4">
                                  <div className="flex items-center gap-1">
                                    {Object.entries(p.files).filter(([k, v]) => k !== 'shareholder_info' && (v as FileState).status === 'ready').map(([k]) => (
                                      <div key={k} className="w-2 h-2 rounded-full bg-green-500" title={k} />
                                    ))}
                                  </div>
                                  <button onClick={(e) => {
                                    e.stopPropagation();
                                    setSelectedProductId(p.id);
                                    setView('comparison');
                                  }} className="text-blue-600 hover:underline text-xs font-bold">查看比对</button>
                                </div>
                              </div>
                            ))}
                          </div>
                        ) : (
                          <div className="flex flex-col items-center gap-4 text-center">
                            <div className="p-4 bg-slate-100 rounded-full text-slate-400">
                              <FileArchive className="w-12 h-12" />
                            </div>
                            <div>
                              <span className="text-lg text-slate-600 font-bold block">拖拽产品合集文件夹 或 ZIP包</span>
                              <span className="text-slate-400 text-sm mt-1">每个产品子文件夹需包含 Word 及 Excel 报表</span>
                            </div>
                            <button 
                              onClick={() => document.getElementById('zip-upload')?.click()}
                              className="mt-4 px-6 py-2 bg-white border border-slate-200 rounded-lg text-sm font-bold shadow-sm hover:bg-slate-50"
                            >
                              或者选择 ZIP 文件
                            </button>
                          </div>
                        )}
                      </div>
                    </div>
                  </section>

                  {/* Right: Global Files */}
                  <section className="bg-white p-8 rounded-3xl border border-slate-200 shadow-xl shadow-slate-200/50 flex flex-col">
                    <h2 className="text-xl font-bold mb-6 flex items-center gap-3 text-slate-800">
                      <div className="p-2 bg-green-100 rounded-lg"><FileSpreadsheet className="w-6 h-6 text-green-600" /></div>
                      全局持有人信息
                    </h2>
                    
                    <div className="flex-1 space-y-4">
                      <p className="text-xs text-slate-400 leading-relaxed font-medium">
                        上传一份包含所有产品持有人份额变动的 Excel 数据，系统将根据产品名称自动匹配。
                      </p>
                      
                      <div className="relative">
                        <input 
                          type="file" 
                          accept=".xlsx,.xls"
                          className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10"
                          onChange={(e) => e.target.files?.[0] && handleGlobalShareholderUpload(e.target.files[0])}
                        />
                        <div className={cn(
                          "border-2 border-dashed rounded-2xl p-8 text-center transition-all",
                          globalShareholder.status === 'ready' 
                            ? "border-green-200 bg-green-50" 
                            : "border-slate-100 hover:border-green-400/40 hover:bg-slate-50"
                        )}>
                          {globalShareholder.status === 'ready' ? (
                            <div className="flex flex-col items-center gap-2">
                              <CheckCircle2 className="w-10 h-10 text-green-500" />
                              <span className="text-sm font-bold text-green-700">{globalShareholder.file?.name}</span>
                            </div>
                          ) : globalShareholder.status === 'parsing' ? (
                            <RefreshCcw className="w-10 h-10 text-green-400 animate-spin" />
                          ) : (
                            <div className="flex flex-col items-center gap-2">
                              <Upload className="w-8 h-8 text-slate-300" />
                              <span className="text-sm font-bold text-slate-500">点击上传份额流水表</span>
                            </div>
                          )}
                        </div>
                      </div>
                    </div>
                  </section>
                </>
                )}
              </div>

                <div className="flex flex-col items-center gap-6 mt-12">
                  <button
                    disabled={uploadMode === 'manual' 
                      ? files.audit_report.status !== 'ready'
                      : !isAllFilesReady && !batchProducts.length
                    }
                    onClick={() => {
                      if (uploadMode === 'manual') {
                        startManualComparison();
                      } else if (!batchProducts.length) {
                        loadDemoData();
                      } else {
                        setView('comparison');
                      }
                    }}
                    className={cn(
                      "w-full max-w-sm py-5 rounded-2xl font-black text-xl flex items-center justify-center gap-4 transition-all transform active:scale-95 shadow-xl",
                      (isAllFilesReady || batchProducts.length || (uploadMode === 'manual' && files.audit_report.status === 'ready')) 
                        ? "bg-primary text-white hover:-translate-y-1 cursor-pointer" 
                        : "bg-slate-200 text-slate-400 cursor-not-allowed"
                    )}
                  >
                    <ArrowRightLeft className="w-8 h-8" />
                    {uploadMode === 'manual' ? '开始核对该产品' : '进入比对中心'}
                  </button>
                  <p className="text-xs text-slate-400 font-bold">
                    {uploadMode === 'manual' ? '请确保审计报告与至少一张报表已上传' : '支持断点续传与多线程并行解析已移除'}
                  </p>
                </div>
              </div>
            ) : (
              <div className="h-full flex flex-col bg-white rounded-2xl shadow-xl border border-slate-200 overflow-hidden">
                {/* Inline Comparison View (Simplified Modal Header) */}
                <div className="px-6 py-4 border-b border-slate-100 flex items-center justify-between bg-slate-50">
                  <div className="flex items-center gap-6">
                    <div className="flex flex-col">
                      <span className="text-[10px] text-slate-400 font-bold uppercase tracking-widest">当前分析产品</span>
                      <div className="flex items-center gap-3">
                        <h2 className="text-lg font-bold text-slate-800">{currentProduct?.name || '演示数据分析'}</h2>
                        {batchProducts.length > 1 && (
                          <select 
                            className="text-xs border-none bg-slate-200 rounded px-2 py-0.5 font-bold outline-none cursor-pointer"
                            value={selectedProductId || ''}
                            onChange={(e) => setSelectedProductId(e.target.value)}
                          >
                            {batchProducts.map(p => (
                              <option key={p.id} value={p.id}>{p.name}</option>
                            ))}
                          </select>
                        )}
                      </div>
                    </div>
                    <div className="h-10 w-px bg-slate-200" />
                    <div className="flex items-center gap-0">
                      {[
                        { id: 'balance', label: '资产负债表' },
                        { id: 'profit', label: '利润表' },
                        { id: 'net_assets', label: '净值变动表' },
                        { id: 'shareholder_info', label: '持有人份额表' },
                        { id: 'notes', label: '财务报表主要项目附注' }
                      ].map(tab => (
                        <button
                          key={tab.id}
                          onClick={() => setActiveTab(tab.id)}
                          className={cn(
                            "py-2 px-4 font-bold text-sm transition-all rounded-md mx-1",
                            activeTab === tab.id 
                              ? "bg-primary text-white shadow-md" 
                              : "text-slate-500 hover:bg-slate-200"
                          )}
                        >
                          {tab.label}
                        </button>
                      ))}
                    </div>
                  </div>
                  <button 
                    onClick={() => setView('upload')}
                    className="flex items-center gap-2 text-sm font-bold text-slate-400 hover:text-primary transition-colors"
                  >
                    <RefreshCcw className="w-4 h-4" />
                    重置比对
                  </button>
                </div>

                {/* Comparison Content */}
                <div className="flex-1 overflow-hidden p-6 flex gap-12 bg-slate-50/50">
                  {/* Left: Audit Data (Document View) */}
                  <div className="flex-[1.2] flex flex-col min-w-0">
                    <div className="mb-4 flex items-center justify-between">
                      <h3 className="font-bold text-slate-700 flex items-center gap-2">
                        <div className="w-1.5 h-4 bg-green-500 rounded-full"></div>
                        审计数据 (报告内容)
                      </h3>
                      <div className="flex items-center gap-2">
                        <button className="text-[10px] px-2 py-1 bg-white border border-slate-200 rounded font-bold text-slate-500 hover:bg-slate-50">原始报告</button>
                        <button className="text-[10px] px-2 py-1 bg-green-500 border border-green-600 rounded font-bold text-white">审计标注</button>
                      </div>
                    </div>
                    
                    <div className="flex-1 overflow-y-auto overflow-x-visible pr-32 pb-20 scrollbar-hide">
                      {renderAuditDocument()}
                    </div>
                  </div>

                  {/* Right: System Data (Excel View) */}
                  <div className="flex-1 flex flex-col min-w-0">
                    <h3 className="mb-4 font-bold text-slate-700 flex items-center gap-2">
                      <div className="w-1.5 h-4 bg-blue-500 rounded-full"></div>
                      系统数据 (报表源)
                    </h3>
                    <div className="flex-1 overflow-auto bg-white border border-slate-200 rounded-lg shadow-xl ring-1 ring-slate-200/50">
                      <table className="w-full text-xs border-collapse min-w-[500px]">
                        <thead className="sticky top-0 bg-[#C6D9F1] shadow-sm z-10">
                          <tr>
                            {tableData.excel[0]?.map((h: string, i: number) => (
                              <th key={i} className={cn(
                                "p-3 text-left border-b border-r border-[#A7BFDB] text-[#365F91] font-bold",
                                i > 0 && "text-right"
                              )}>
                                {h}
                              </th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {tableData.excel.slice(1).map((row, rowIndex) => (
                            <tr key={rowIndex} className="hover:bg-blue-50/50 transition-colors border-b">
                              {row.map((cell: any, cellIndex: number) => (
                                <td key={cellIndex} className={cn(
                                  "p-2.5 border-r border-slate-100 text-slate-600 font-mono",
                                  cellIndex > 0 && "text-right"
                                )}>
                                  {cell || ''}
                                </td>
                              ))}
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                    <div className="mt-4 p-3 bg-blue-50 border border-blue-100 rounded-lg">
                      <p className="text-[10px] text-blue-600 font-bold uppercase mb-1">比对指南</p>
                      <p className="text-xs text-blue-800 font-medium">
                        {activeTab === 'shareholder_info' 
                          ? '当前模式：提取报告中「持有人账户」表格，与股东名册(TA)数据进行核对。'
                          : '当前模式：针对财报附注描述中的数值，与余额表(TB)及科目明细进行穿透核对。'}
                      </p>
                    </div>
                  </div>
                </div>
              </div>
            )}
          </div>
        </main>
      </div>
    </div>
  );
}

// --- Internal Logic Helpers ---

/**
 * Injects highlights into HTML content by analyzing numbers and comparing with Excel data.
 */
function injectHighlights(html: string, files: Record<FileType, FileState>): string {
  if (!html) return '';

  // 1. Gather all Excel numbers and which file they came from
  const excelSummaryMap: Record<string, string[]> = {};
  REQUIRED_EXCELS.forEach(({ key, label }) => {
    const data = files[key].data as any[][];
    if (!data) return;

    data.flat().forEach(cell => {
      if (typeof cell === 'number' || (typeof cell === 'string' && /^-?\d+(\.\d+)?$/.test(cell.trim()))) {
        const valStr = cell.toString().trim();
        if (!excelSummaryMap[valStr]) excelSummaryMap[valStr] = [];
        if (!excelSummaryMap[valStr].includes(label)) {
          excelSummaryMap[valStr].push(label);
        }
      }
    });
  });

  // 2. Parse HTML safely and inject highlights
  // We use a temporary DOM element to process node by node to avoid breaking HTML structure
  const parser = new DOMParser();
  const doc = parser.parseFromString(`<div>${html}</div>`, 'text/html');
  const container = doc.body.firstChild as HTMLElement;

  const processNode = (node: Node) => {
    if (node.nodeType === Node.TEXT_NODE) {
      const text = node.textContent || '';
      // Find numbers like "1,234.56" or "1234.56"
      // We look for patterns with at least 1 digit, possibly commas/dots
      const parts = text.split(/(\d[\d\s,.]*\d|\d)/g);
      
      const fragment = document.createDocumentFragment();
      parts.forEach(part => {
        const cleanVal = part.replace(/[, \s]/g, '').trim();
        
        // Check if it's a financial looking number
        if (/^-?\d+(\.\d+)?$/.test(cleanVal) && cleanVal.length > 0) {
          const sources = excelSummaryMap[cleanVal];
          const span = document.createElement('span');
          span.textContent = part;
          
          if (sources && sources.length > 0) {
            span.className = 'highlight-match';
            span.title = `匹配成功: 在 [${sources.join(', ')}] 中找到一致数值`;
          } else if (cleanVal.length > 1) { // Ignore single digits for less noise
            span.className = 'highlight-diff';
            span.title = '未在 Excel 报表中找到一致数值，请人工核查';
          }
          fragment.appendChild(span);
        } else {
          fragment.appendChild(document.createTextNode(part));
        }
      });

      node.parentNode?.replaceChild(fragment, node);
    } else if (node.nodeType === Node.ELEMENT_NODE) {
      // Don't process specific elements if needed
      Array.from(node.childNodes).forEach(processNode);
    }
  };

  processNode(container);
  return container.innerHTML;
}

