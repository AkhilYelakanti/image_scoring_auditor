/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useMemo, useCallback } from 'react';
import * as XLSX from 'xlsx';
import { 
  FileUp, 
  ArrowRight, 
  Download, 
  CheckCircle2, 
  AlertCircle, 
  Filter, 
  Search,
  Image as ImageIcon,
  ChevronDown,
  ChevronUp,
  Columns,
  RefreshCw,
  Info,
  Database
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';

// --- Utilities ---
function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

// --- Types ---
interface ImageItem {
  TABLE_ID: string | number;
  CSOR_IMAGE_NAME: string;
  UPC: string | number;
  CUSTOMER_ITEM_CODE: string;
  GL: string;
  DERIVED_DEPARTMENT: string;
  DESTINATION: string;
  DOWNLOAD_SOURCE: string;
  SOURCE_IMAGE_TYPE: string;
  IMAGE_STATUS: string;
  PKG_NAME: string;
  PKG_STATUS: string;
  CSOR_ITEM_IMAGE_MASTER_RAW_ID: string | number;
  SOURCE_IMAGE_URL: string;
  [key: string]: any;
}

interface ComparisonResult {
  upc: string;
  previous: ImageItem | null;
  current: ImageItem | null;
  hasChanged: boolean;
  changes: string[];
}

// --- Components ---

const ComparisonRow = ({ result }: { result: ComparisonResult; key?: string }) => {
  const isChanged = result.hasChanged;
  
  return (
    <motion.div 
      layout
      initial={{ opacity: 0 }}
      animate={{ opacity: 1 }}
      className={cn(
        "grid grid-cols-[160px_1fr_1fr_120px] items-stretch border-b border-slate-100 transition-colors",
        isChanged ? "bg-blue-50/5" : "hover:bg-slate-50/50"
      )}
    >
      {/* UPC / Meta */}
      <div className="px-6 py-4 flex flex-col justify-center border-r border-slate-100">
        <span className="text-sm font-bold text-slate-900 font-mono tracking-tight">{result.upc}</span>
        <span className="text-[10px] text-slate-400 font-mono mt-1">ID: {result.current?.TABLE_ID || result.previous?.TABLE_ID}</span>
        {(result.current?.GL || result.previous?.GL) && (
          <span className="mt-2 px-1.5 py-0.5 bg-slate-100 text-[9px] font-bold rounded text-slate-500 w-fit uppercase tracking-tighter">
            {result.current?.GL || result.previous?.GL}
          </span>
        )}
      </div>

      {/* Previous Version */}
      <div className="px-6 py-4 border-r border-slate-100 flex items-center gap-4">
        <div className="w-16 h-16 bg-slate-50 rounded border border-slate-200 flex items-center justify-center shrink-0 relative overflow-hidden group">
          <div className="absolute inset-0 bg-gradient-to-tr from-slate-200 to-transparent opacity-20"></div>
          {result.previous ? (
            <img 
              src={result.previous.SOURCE_IMAGE_URL} 
              alt="Previous"
              referrerPolicy="no-referrer"
              className="w-full h-full object-contain mix-blend-multiply relative z-10"
              onError={(e) => {
                (e.target as HTMLImageElement).src = 'https://placehold.co/400?text=No+Image';
              }}
            />
          ) : (
            <span className="text-[9px] text-slate-300 font-bold uppercase z-10">Empty</span>
          )}
        </div>
        <div className="flex flex-col min-w-0">
          <span className="text-[11px] font-medium text-slate-600 truncate">
            {result.previous?.CSOR_IMAGE_NAME || '---'}
          </span>
          <span className="text-[9px] text-slate-400 truncate mt-0.5">
            {result.previous?.DOWNLOAD_SOURCE || 'No Source'}
          </span>
        </div>
      </div>

      {/* Current Scored Version */}
      <div className={cn(
        "px-6 py-4 border-r border-slate-100 flex items-center gap-4 transition-colors",
        isChanged ? "bg-blue-50/20" : "bg-slate-50/20"
      )}>
        <div className={cn(
          "w-16 h-16 rounded border flex items-center justify-center shrink-0 shadow-sm relative overflow-hidden",
          isChanged ? "bg-white border-blue-200 ring-2 ring-blue-500/5" : "bg-white border-slate-200"
        )}>
          {result.current ? (
            <img 
              src={result.current.SOURCE_IMAGE_URL} 
              alt="Current"
              referrerPolicy="no-referrer"
              className="w-full h-full object-contain mix-blend-multiply"
              onError={(e) => {
                (e.target as HTMLImageElement).src = 'https://placehold.co/400?text=No+Image';
              }}
            />
          ) : (
            <span className="text-[9px] text-slate-300 font-bold uppercase">Removed</span>
          )}
        </div>
        <div className="flex flex-col min-w-0">
          <span className={cn(
            "text-[11px] font-bold truncate",
            isChanged ? "text-blue-700" : "text-slate-700"
          )}>
            {result.current?.CSOR_IMAGE_NAME || '---'}
          </span>
          <span className="text-[9px] text-blue-500 truncate mt-0.5 font-medium">
            {result.current?.DOWNLOAD_SOURCE || 'Scoring Active'}
          </span>
          {isChanged && result.changes.length > 0 && (
            <span className="text-[8px] text-amber-600 font-bold mt-1 uppercase tracking-tighter">
              Δ {result.changes.join(', ')}
            </span>
          )}
        </div>
      </div>

      {/* Status */}
      <div className="px-4 py-4 flex flex-col items-center justify-center">
        {isChanged ? (
          <div className="px-2 py-1 bg-blue-100 text-blue-700 text-[9px] font-bold rounded border border-blue-200 uppercase tracking-tight">
            Updated
          </div>
        ) : (
          <div className="px-2 py-1 bg-slate-50 text-slate-400 text-[9px] font-bold rounded border border-slate-200 uppercase tracking-tight">
            Matched
          </div>
        )}
      </div>
    </motion.div>
  );
};


export default function App() {
  const [prevData, setPrevData] = useState<ImageItem[] | null>(null);
  const [newData, setNewData] = useState<ImageItem[] | null>(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [filter, setFilter] = useState<'all' | 'changed'>('changed');
  const [searchQuery, setSearchQuery] = useState('');

  const handleFileUpload = (type: 'prev' | 'new') => async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsProcessing(true);
    try {
      const reader = new FileReader();
      reader.onload = (evt) => {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const data = XLSX.utils.sheet_to_json(ws) as ImageItem[];
        
        if (type === 'prev') setPrevData(data);
        else setNewData(data);
        setIsProcessing(false);
      };
      reader.readAsBinaryString(file);
    } catch (error) {
      console.error('Error reading excel:', error);
      setIsProcessing(false);
    }
  };

  const comparisonResults = useMemo(() => {
    if (!prevData && !newData) return [];

    const allUpcs = new Set([
      ...(prevData?.map(d => String(d.UPC)) || []),
      ...(newData?.map(d => String(d.UPC)) || [])
    ]);

    const prevMap = new Map<string, ImageItem>(prevData?.map(d => [String(d.UPC), d]) || []);
    const newMap = new Map<string, ImageItem>(newData?.map(d => [String(d.UPC), d]) || []);

    const results: ComparisonResult[] = Array.from(allUpcs).map(upc => {
      const prev = prevMap.get(upc) || null;
      const current = newMap.get(upc) || null;
      
      const changes: string[] = [];
      if (prev && current) {
        if (prev.CSOR_IMAGE_NAME !== current.CSOR_IMAGE_NAME) changes.push('IMAGE_NAME');
        if (prev.SOURCE_IMAGE_URL !== current.SOURCE_IMAGE_URL) changes.push('URL');
        if (prev.IMAGE_STATUS !== current.IMAGE_STATUS) changes.push('STATUS');
        if (prev.PKG_NAME !== current.PKG_NAME) changes.push('PACKAGE');
      } else if (prev || current) {
        changes.push(prev ? 'DELETED' : 'CREATED');
      }

      return {
        upc,
        previous: prev,
        current,
        hasChanged: changes.length > 0,
        changes
      };
    });

    return results;
  }, [prevData, newData]);

  const filteredResults = useMemo(() => {
    let results = comparisonResults;
    if (filter === 'changed') {
      results = results.filter(r => r.hasChanged);
    }
    if (searchQuery) {
      results = results.filter(r => r.upc.includes(searchQuery));
    }
    return results;
  }, [comparisonResults, filter, searchQuery]);

  const stats = useMemo(() => {
    return {
      total: comparisonResults.length,
      changed: comparisonResults.filter(r => r.hasChanged).length,
      same: comparisonResults.filter(r => !r.hasChanged).length
    };
  }, [comparisonResults]);

  const exportData = () => {
    const dataToExport = filteredResults.map(r => ({
      UPC: r.upc,
      Status: r.hasChanged ? 'Changed' : 'Same',
      Change_Types: r.changes.join(', '),
      Prev_Image_Name: r.previous?.CSOR_IMAGE_NAME || 'N/A',
      New_Image_Name: r.current?.CSOR_IMAGE_NAME || 'N/A',
      Prev_URL: r.previous?.SOURCE_IMAGE_URL || 'N/A',
      New_URL: r.current?.SOURCE_IMAGE_URL || 'N/A',
    }));

    const ws = XLSX.utils.json_to_sheet(dataToExport);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Comparison");
    XLSX.writeFile(wb, `Image_Comparison_Result_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  return (
    <div className="h-screen flex flex-col bg-[#F8FAFC] text-slate-900 font-sans overflow-hidden">
      {/* Header */}
      <header className="bg-[#0F172A] text-white px-6 py-4 flex justify-between items-center shrink-0 shadow-lg relative z-50">
        <div className="flex items-center gap-3">
          <div className="w-8 h-8 bg-blue-600 rounded flex items-center justify-center">
            <Database className="w-5 h-5 text-white" />
          </div>
          <h1 className="text-xl font-semibold tracking-tight">Image Scoring Auditor</h1>
        </div>

        <div className="flex items-center gap-4 text-xs">
          {prevData && (
            <span className="px-3 py-1.5 bg-slate-700/50 rounded-full border border-slate-600/50">
              Baseline: <strong className="ml-1 text-slate-100">{prevData.length} records</strong>
            </span>
          )}
          {newData && (
            <span className="px-3 py-1.5 bg-slate-700/50 rounded-full border border-slate-600/50">
              Target: <strong className="ml-1 text-blue-300">{newData.length} records</strong>
            </span>
          )}
          {comparisonResults.length > 0 && (
            <button 
              onClick={exportData}
              className="bg-blue-600 hover:bg-blue-500 px-5 py-2 rounded-md font-bold transition-all flex items-center gap-2 shadow-sm active:scale-95"
            >
              <Download className="w-3.5 h-3.5" /> Export Results
            </button>
          )}
        </div>
      </header>

      {/* Nav / Filters */}
      <nav className="bg-white border-b border-slate-200 px-6 py-3 flex justify-between items-center shrink-0 shadow-sm z-40">
        <div className="flex gap-1 p-1 bg-slate-100 rounded-lg border border-slate-200">
          <button 
            onClick={() => setFilter('all')}
            className={cn(
              "px-4 py-1.5 rounded-md text-xs font-bold transition-all",
              filter === 'all' ? "bg-white shadow-sm text-slate-800" : "text-slate-500 hover:text-slate-700"
            )}
          >
            All Items ({stats.total})
          </button>
          <button 
            onClick={() => setFilter('changed')}
            className={cn(
              "px-4 py-1.5 rounded-md text-xs font-bold transition-all",
              filter === 'changed' ? "bg-white shadow-sm text-blue-600" : "text-slate-500 hover:text-slate-700"
            )}
          >
            Changed Only ({stats.changed})
          </button>
        </div>

        <div className="flex items-center gap-6">
          <div className="text-[10px] text-slate-500 flex gap-4 uppercase font-bold tracking-widest">
            <span>Matches: <span className="text-emerald-600 font-mono">{( (stats.same / (stats.total || 1)) * 100).toFixed(1)}%</span></span>
            <span>Delta: <span className="text-blue-600 font-mono">{( (stats.changed / (stats.total || 1)) * 100).toFixed(1)}%</span></span>
          </div>
          <div className="relative">
            <input 
              type="text" 
              placeholder="Filter by UPC..." 
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              className="w-64 pl-10 pr-4 py-2 bg-slate-50 border border-slate-200 rounded-lg text-xs focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500/50 transition-all font-mono"
            />
            <Search className="w-4 h-4 text-slate-400 absolute left-3 top-2.5" />
          </div>
        </div>
      </nav>

      <main className="flex-1 flex flex-col min-h-0 bg-white relative">
        {/* Table Header */}
        <div className="grid grid-cols-[160px_1fr_1fr_120px] border-b border-slate-200 bg-slate-50/80 text-[10px] font-bold uppercase tracking-wider text-slate-500 z-30">
          <div className="px-6 py-3">UPC / Identifier</div>
          <div className="px-6 py-3 border-l border-slate-200">Baseline Assignment</div>
          <div className="px-6 py-3 border-l border-slate-200">Scored Assignment (Current)</div>
          <div className="px-6 py-3 border-l border-slate-200 text-center">Audit Status</div>
        </div>

        {/* List Content */}
        <div className="flex-1 overflow-y-auto no-scrollbar">
          {!prevData && !newData ? (
            <div className="h-full flex flex-col items-center justify-center p-20">
              <div className="grid grid-cols-1 md:grid-cols-2 gap-8 w-full max-w-4xl">
                <div className="relative group">
                  <input type="file" accept=".xlsx,.xls,.csv" onChange={handleFileUpload('prev')} className="absolute inset-0 opacity-0 cursor-pointer z-10" />
                  <div className="h-64 rounded-3xl border-2 border-dashed border-slate-200 bg-white flex flex-col items-center justify-center gap-4 transition-all group-hover:border-blue-500/50 group-hover:bg-blue-50/10">
                    <div className="w-16 h-16 bg-slate-50 rounded-2xl flex items-center justify-center text-slate-400 group-hover:text-blue-500 transition-colors">
                      <FileUp className="w-8 h-8" />
                    </div>
                    <div className="text-center">
                      <p className="font-bold text-slate-700">Baseline Data</p>
                      <p className="text-[10px] text-slate-400 uppercase tracking-widest mt-1">Upload Previous Excel</p>
                    </div>
                  </div>
                </div>
                <div className="relative group">
                  <input type="file" accept=".xlsx,.xls,.csv" onChange={handleFileUpload('new')} className="absolute inset-0 opacity-0 cursor-pointer z-10" />
                  <div className="h-64 rounded-3xl border-2 border-dashed border-slate-200 bg-white flex flex-col items-center justify-center gap-4 transition-all group-hover:border-blue-500/50 group-hover:bg-blue-50/10">
                    <div className="w-16 h-16 bg-slate-50 rounded-2xl flex items-center justify-center text-slate-400 group-hover:text-blue-500 transition-colors">
                      <FileUp className="w-8 h-8" />
                    </div>
                    <div className="text-center">
                      <p className="font-bold text-slate-700">Audit Target</p>
                      <p className="text-[10px] text-slate-400 uppercase tracking-widest mt-1">Upload Scored Excel</p>
                    </div>
                  </div>
                </div>
              </div>
              <div className="mt-16 flex items-center gap-12 text-slate-400">
                <div className="flex flex-col items-center gap-2">
                  <span className="text-[10px] font-bold">STEP 01</span>
                  <span className="text-xs">Baseline Data</span>
                </div>
                <ArrowRight className="w-4 h-4 opacity-20" />
                <div className="flex flex-col items-center gap-2">
                  <span className="text-[10px] font-bold">STEP 02</span>
                  <span className="text-xs">New Scoring</span>
                </div>
                <ArrowRight className="w-4 h-4 opacity-20" />
                <div className="flex flex-col items-center gap-2">
                  <span className="text-[10px] font-bold">STEP 03</span>
                  <span className="text-xs">Visual Audit</span>
                </div>
              </div>
            </div>
          ) : (
            <div className="divide-y divide-slate-100">
              <AnimatePresence mode="popLayout">
                {filteredResults.length > 0 ? (
                  filteredResults.map((result) => (
                    <ComparisonRow key={result.upc} result={result} />
                  ))
                ) : (
                  <motion.div initial={{ opacity: 0 }} animate={{ opacity: 1 }} className="py-20 flex flex-col items-center text-slate-400 font-medium">
                    <Search className="w-12 h-12 mb-4 opacity-10" />
                    <p className="text-sm">No results match your current filters</p>
                  </motion.div>
                )}
              </AnimatePresence>
            </div>
          )}
        </div>
      </main>

      {/* Footer */}
      <footer className="bg-slate-50 border-t border-slate-200 px-6 py-2.5 flex justify-between items-center text-[10px] text-slate-500 shrink-0 font-bold uppercase tracking-wider">
        <div className="flex items-center gap-6">
          <div className="flex items-center gap-2">
            <div className="w-1.5 h-1.5 rounded-full bg-emerald-500"></div> Engine Synchronized
          </div>
          <div className="flex items-center gap-2">
            <div className="w-1.5 h-1.5 rounded-full bg-blue-500"></div> {filteredResults.length} Items Displayed
          </div>
        </div>
        <div className="flex items-center gap-4">
          <span>{new Date().toLocaleDateString(undefined, { month: 'short', day: 'numeric', year: 'numeric' })}</span>
          <span className="w-px h-2 bg-slate-300"></span>
          <span>Auditor V2.4.0</span>
        </div>
      </footer>

      {/* Processing Loader */}
      {isProcessing && (
        <div className="fixed inset-0 bg-slate-900/40 backdrop-blur-sm z-[100] flex items-center justify-center">
          <div className="bg-white p-8 rounded-3xl shadow-2xl flex flex-col items-center gap-4">
            <div className="w-12 h-12 border-4 border-blue-600 border-t-transparent rounded-full animate-spin"></div>
            <p className="font-bold text-slate-700 tracking-tight">Processing Analytics...</p>
          </div>
        </div>
      )}
    </div>
  );
}

