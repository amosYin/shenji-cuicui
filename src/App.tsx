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
  Info
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import * as mammoth from 'mammoth';
import * as XLSX from 'xlsx';
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
  { key: 'shareholder_info', label: '持有人历史份额信息' },
  { key: 'net_assets', label: '净资产变动表' },
  { key: 'profit', label: '利润表' },
  { key: 'balance', label: '资产负债表' },
  { key: 'trial_balance', label: '余额表' },
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

  const [view, setView] = useState<'upload' | 'comparison'>('upload');
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

  // --- Handlers ---
  const handleFileUpload = async (type: FileType, file: File) => {
    setFiles(prev => ({
      ...prev,
      [type]: { ...prev[type], file, status: 'parsing', error: undefined }
    }));

    try {
      if (type === 'audit_report') {
        const arrayBuffer = await file.arrayBuffer();
        const result = await mammoth.convertToHtml({ arrayBuffer });
        setWordHtml(result.value);
        
        // Simple extraction logic: find numbers in tables
        const textResult = await mammoth.extractRawText({ arrayBuffer });
        const lines = textResult.value.split('\n').filter(l => l.trim().length > 0);
        setFiles(prev => ({
          ...prev,
          [type]: { ...prev[type], data: lines, status: 'ready' }
        }));
      } else {
        const arrayBuffer = await file.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        
        setFiles(prev => ({
          ...prev,
          [type]: { ...prev[type], data: jsonData, status: 'ready' }
        }));
      }
    } catch (error) {
      console.error(`Error parsing ${type}:`, error);
      setFiles(prev => ({
        ...prev,
        [type]: { ...prev[type], status: 'error', error: '解析失败，请检查文件格式。' }
      }));
    }
  };

  const isAllFilesReady = useMemo(() => {
    return (Object.values(files) as FileState[]).every(f => f.status === 'ready');
  }, [files]);

  // --- Comparison Logic for side-by-side ---
  const tableData = useMemo(() => {
    let excelData: any[][] = [[]];
    let wordData: any[][] = [[]];

    if (activeTab === 'net_assets') {
      excelData = files.net_assets.data as any[][];
      wordData = files.audit_report.data as any[][];
    } else if (activeTab === 'shareholder_info') {
      excelData = [
        ['持有人类别', '持有份额', '占总份额比例'],
        ['机构投资者', '50,000,000.00', '95.24%'],
        ['个人投资者', '2,500,000.00', '4.76%'],
        ['合计', '52,500,000.00', '100.00%']
      ];
      wordData = [
        ['持有人类别', '持有份额', '占总份额比例'],
        ['机构投资者', '50,000,000.00', '95.24%'],
        ['个人投资者', '2,400,000.00', '4.57%'], // Difference
        ['合计', '52,400,000.00', '99.81%']     // Difference
      ];
    } else if (activeTab === 'notes') {
      excelData = [
        ['项目', '2025年12月31日', '2024年12月31日'],
        ['活期存款', '3,114,074.61', '7,969,396.39'],
        ['应计利息', '1,423.59', '47,580.14'],
        ['合计', '3,115,498.20', '8,016,976.53']
      ];
      wordData = [
        ['项目', '2025年12月31日', '2024年12月31日'],
        ['活期存款', '3,114,074.61', '7,969,396.39'],
        ['应计利息', '1,424.00', '47,580.14'], // Small cent difference
        ['合计', '3,115,498.61', '8,016,976.53']
      ];
    } else {
      excelData = [['暂未匹配演示数据', '']];
      wordData = [['暂未匹配演示数据', '']];
    }

    if (!excelData || !wordData) return { excel: [[]], word: [[]] };
    return { excel: excelData, word: wordData };
  }, [activeTab, files.net_assets.data, files.audit_report.data]);

  return (
    <div className="min-h-screen bg-slate-100 flex flex-col">
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
            <button className="w-full flex items-center gap-3 px-4 py-3 text-sm font-medium text-primary bg-red-50 rounded-lg">
              <FileSearch className="w-5 h-5" />
              产品财务审计
            </button>
          </nav>
        </aside>

        {/* Content Area */}
        <main className="flex-1 overflow-auto bg-slate-50 relative flex flex-col">
          {/* Main Content Tabs */}
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

          <div className="flex-1 overflow-auto p-8">
            {view === 'upload' ? (
              <div className="max-w-4xl mx-auto space-y-8">
                <header className="text-center md:text-left mb-12">
                  <h1 className="text-4xl font-extrabold text-slate-900 tracking-tight mb-2">审计报告核对工具</h1>
                  <p className="text-slate-500 text-lg">智能提取审计报告与报表数据，自动完成勾稽核查。</p>
                </header>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
                  {/* Main Word File */}
                  <section className="bg-white p-8 rounded-3xl border border-slate-200 shadow-xl shadow-slate-200/50">
                    <h2 className="text-xl font-bold mb-6 flex items-center gap-3 text-slate-800">
                      <div className="p-2 bg-red-100 rounded-lg"><FileText className="w-6 h-6 text-primary" /></div>
                      审计报告 (Word)
                    </h2>
                    <div className="relative group">
                      <input 
                        type="file" 
                        accept=".docx"
                        className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10"
                        onChange={(e) => e.target.files?.[0] && handleFileUpload('audit_report', e.target.files[0])}
                      />
                      <div className={cn(
                        "border-2 border-dashed rounded-2xl p-12 text-center transition-all duration-300",
                        files.audit_report.status === 'ready' 
                          ? "border-green-200 bg-green-50" 
                          : "border-slate-200 group-hover:border-primary/40 group-hover:bg-slate-50"
                      )}>
                        {files.audit_report.status === 'ready' ? (
                          <div className="flex flex-col items-center gap-3">
                            <div className="p-4 bg-green-100 rounded-full"><CheckCircle2 className="w-12 h-12 text-green-600" /></div>
                            <span className="text-lg font-bold text-green-800 tracking-tight">{files.audit_report.file?.name}</span>
                          </div>
                        ) : (
                          <div className="flex flex-col items-center gap-3">
                            <div className="p-4 bg-slate-100 rounded-full text-slate-400 group-hover:text-primary transition-colors">
                              <Upload className="w-12 h-12" />
                            </div>
                            <span className="text-lg text-slate-600 font-bold tracking-tight">点击或拖拽上传报告</span>
                            <span className="text-slate-400 text-sm italic">仅支持 .docx 格式</span>
                          </div>
                        )}
                      </div>
                    </div>
                  </section>

                  {/* Secondary Excel Files */}
                  <section className="bg-white p-8 rounded-3xl border border-slate-200 shadow-xl shadow-slate-200/50 flex flex-col">
                    <h2 className="text-xl font-bold mb-6 flex items-center gap-3 text-slate-800">
                      <div className="p-2 bg-green-100 rounded-lg"><FileSpreadsheet className="w-6 h-6 text-green-600" /></div>
                      财务报表 (Excel)
                    </h2>
                    <div className="flex-1 grid grid-cols-1 gap-3 overflow-auto max-h-[400px] pr-2">
                      {REQUIRED_EXCELS.map((item) => (
                        <div key={item.key} className="flex items-center justify-between p-4 rounded-xl border border-slate-100 bg-slate-50/50 hover:bg-slate-50 hover:border-slate-200 transition-all">
                          <div className="flex items-center gap-4">
                            <div className={cn(
                              "w-3 h-3 rounded-full animate-pulse",
                              files[item.key].status === 'ready' ? "bg-green-500 shadow-[0_0_12px_rgba(34,197,94,0.6)]" : "bg-slate-300"
                            )} />
                            <span className="font-bold text-slate-700">{item.label}</span>
                          </div>
                          
                          <div className="relative">
                            <input 
                              type="file" 
                              accept=".xlsx,.xls"
                              className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
                              onChange={(e) => e.target.files?.[0] && handleFileUpload(item.key, e.target.files[0])}
                            />
                            <button className={cn(
                              "px-4 py-2 text-sm font-bold rounded-lg transition-all shadow-sm",
                              files[item.key].status === 'ready' 
                                ? "bg-green-100 text-green-700 border border-green-200" 
                                : "bg-white text-primary border border-primary hover:bg-primary/5"
                            )}>
                              {files[item.key].status === 'ready' ? '更换数据' : '选择文件'}
                            </button>
                          </div>
                        </div>
                      ))}
                    </div>
                  </section>
                </div>

                <div className="flex flex-col items-center gap-6 mt-12">
                  <button
                    onClick={() => {
                      if (!isAllFilesReady) {
                        loadDemoData();
                      } else {
                        setView('comparison');
                      }
                    }}
                    className="w-full max-w-sm py-5 rounded-2xl font-black text-xl flex items-center justify-center gap-4 transition-all transform active:scale-95 bg-primary text-white shadow-[0_20px_50px_rgba(238,31,45,0.3)] hover:shadow-[0_20px_50px_rgba(238,31,45,0.5)] hover:-translate-y-1 cursor-pointer"
                  >
                    <ArrowRightLeft className="w-8 h-8" />
                    开始比对数据
                  </button>
                </div>
              </div>
            ) : (
              <div className="h-full flex flex-col bg-white rounded-2xl shadow-xl border border-slate-200 overflow-hidden">
                {/* Inline Comparison View (Simplified Modal Header) */}
                <div className="px-6 py-4 border-b border-slate-100 flex items-center justify-between bg-slate-50">
                  <div className="flex items-center gap-6">
                    <h2 className="text-lg font-bold text-slate-800">比对详情报告</h2>
                    <div className="h-6 w-px bg-slate-200" />
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
                <div className="flex-1 overflow-hidden p-6 flex gap-6 bg-slate-50/30">
                  {/* Left: Generated Data (Excel) */}
                  <div className="flex-1 flex flex-col min-w-0">
                    <h3 className="mb-3 font-bold text-slate-700 flex items-center gap-2">
                      <div className="w-1.5 h-4 bg-primary rounded-full"></div>
                      生成数据
                    </h3>
                    <div className="flex-1 overflow-auto bg-white border border-slate-200 rounded-lg shadow-sm">
                      <table className="w-full text-[13px] border-collapse min-w-[600px]">
                        <thead className="sticky top-0 bg-[#C6D9F1] z-10">
                          <tr>
                            {tableData.excel[0]?.map((h: string, i: number) => (
                              <th key={i} className={cn(
                                "p-2.5 text-left border-b border-r border-[#A7BFDB] text-[#365F91] font-bold",
                                i > 0 && "text-right"
                              )}>
                                {h}
                              </th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {tableData.excel.slice(1).map((row, rowIndex) => (
                            <tr key={rowIndex} className="hover:bg-slate-50 transition-colors">
                              {row.map((cell: any, cellIndex: number) => {
                                const wordCell = tableData.word[rowIndex + 1]?.[cellIndex];
                                const isMatch = cell === wordCell || (!cell && !wordCell);
                                return (
                                  <td key={cellIndex} className={cn(
                                    "p-2.5 border-b border-r border-slate-100 text-slate-600",
                                    cellIndex > 0 && "text-right",
                                    !isMatch && cell && "bg-[#FDE9D9] text-red-600 font-medium"
                                  )}>
                                    {cell || ''}
                                  </td>
                                );
                              })}
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>

                  {/* Right: Audit Data (Word) */}
                  <div className="flex-1 flex flex-col min-w-0">
                    <h3 className="mb-3 font-bold text-slate-700 flex items-center gap-2">
                      <div className="w-1.5 h-4 bg-green-500 rounded-full"></div>
                      审计数据
                    </h3>
                    <div className="flex-1 overflow-auto bg-white border border-slate-200 rounded-lg shadow-sm">
                      <table className="w-full text-[13px] border-collapse min-w-[600px]">
                        <thead className="sticky top-0 bg-[#C6D9F1] z-10">
                          <tr>
                            {tableData.word[0]?.map((h: string, i: number) => (
                              <th key={i} className={cn(
                                "p-2.5 text-left border-b border-r border-[#A7BFDB] text-[#365F91] font-bold",
                                i > 0 && "text-right"
                              )}>
                                {h}
                              </th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {tableData.word.slice(1).map((row, rowIndex) => (
                            <tr key={rowIndex} className="hover:bg-slate-50 transition-colors">
                              {row.map((cell: any, cellIndex: number) => {
                                const excelCell = tableData.excel[rowIndex + 1]?.[cellIndex];
                                const isMatch = cell === excelCell || (!cell && !excelCell);
                                return (
                                  <td key={cellIndex} className={cn(
                                    "p-2.5 border-b border-r border-slate-100 text-slate-600",
                                    cellIndex > 0 && "text-right",
                                    !isMatch && cell && "bg-[#EBF1DE] text-green-800 font-medium"
                                  )}>
                                    {cell || ''}
                                  </td>
                                );
                              })}
                            </tr>
                          ))}
                        </tbody>
                      </table>
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

