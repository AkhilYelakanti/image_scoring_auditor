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
  Database,
  CheckSquare,
  Square,
  FileText,
  Eye,
  EyeOff,
  Trash2,
  MoreVertical,
  Camera,
  Layout,
  Maximize2
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';
import { jsPDF } from 'jspdf';
import html2canvas from 'html2canvas';
import { List } from 'react-window';

/// --- Utilities ---
function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

/**
 * Sanitizes image URLs to ensure they use HTTPS to avoid mixed-content blocks
 * and handles common image hosting patterns.
 */
function formatImageUrl(url: string | undefined): string {
  if (!url) return 'https://placehold.co/400?text=No+URL';
  
  let sanitized = String(url).trim();
  
  // Handle protocol-relative URLs
  if (sanitized.startsWith('//')) {
    sanitized = 'https:' + sanitized;
  }
  
  // Force HTTPS for common providers
  if (sanitized.startsWith('http://')) {
    sanitized = sanitized.replace('http://', 'https://');
  }
  
  return sanitized;
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
  auditStatus: 'pending' | 'reviewed' | 'ignored';
}

type FilterStatus = 'all' | 'changed' | 'reviewed' | 'ignored';
type SortKey = 'upc' | 'prev_name' | 'curr_name' | 'status' | 'auditStatus';
type SortOrder = 'asc' | 'desc';

// --- Components ---

interface ComparisonRowProps {
  result: ComparisonResult;
  onPreview: (res: ComparisonResult) => void;
  isSelected: boolean;
  onSelect: (upc: string) => void;
  style?: React.CSSProperties;
  density?: 'compact' | 'standard' | 'large';
}

const ComparisonRow = React.memo(({ 
  result, 
  onPreview,
  isSelected,
  onSelect,
  style,
  density = 'standard'
}: ComparisonRowProps) => {
  const isChanged = result.hasChanged;
  
  return (
    <div 
      style={{
        ...style,
        height: style?.height || 'auto'
      }}
      className={cn(
        "grid grid-cols-[48px_160px_1fr_1fr_120px] items-stretch border-b border-slate-100 transition-colors overflow-hidden relative",
        isSelected ? "bg-blue-50/40" : isChanged ? "bg-blue-50/10" : "hover:bg-slate-50/50"
      )}
      id={`row-${result.upc}`}
    >
      {/* Change Highlight Bar */}
      {isChanged && (
        <div className="absolute left-0 top-0 bottom-0 w-1 bg-amber-400 z-20 shadow-[0_0_8px_rgba(251,191,36,0.5)]" />
      )}
      {/* Selection */}
      <div className="flex items-center justify-center border-r border-slate-100">
        <button 
          onClick={() => onSelect(result.upc)}
          className={cn(
            "w-5 h-5 rounded flex items-center justify-center transition-all",
            isSelected ? "bg-blue-600 text-white" : "border border-slate-300 text-transparent hover:border-blue-400"
          )}
        >
          {isSelected && <CheckSquare className="w-3.5 h-3.5" />}
        </button>
      </div>

      {/* UPC / Meta */}
      <div className="px-6 py-2 flex flex-col justify-center border-r border-slate-100 italic">
        <span className={cn(
          "font-bold text-slate-900 font-mono tracking-tight",
          density === 'compact' ? "text-xs" : "text-sm"
        )}>{result.upc}</span>
        {density !== 'compact' && (
          <div className="flex items-center gap-1.5 mt-1">
            <span className="text-[10px] text-slate-400 font-mono">ID: {result.current?.TABLE_ID || result.previous?.TABLE_ID}</span>
            {result.auditStatus !== 'pending' && (
              <span className={cn(
                "text-[8px] px-1 rounded font-bold uppercase",
                result.auditStatus === 'reviewed' ? "bg-emerald-100 text-emerald-700" : "bg-slate-200 text-slate-600"
              )}>
                {result.auditStatus}
              </span>
            )}
          </div>
        )}
      </div>

      {/* Previous Version */}
      <div className="px-6 py-4 border-r border-slate-100 flex items-center gap-4 group cursor-zoom-in" onClick={() => onPreview(result)}>
        <div className={cn(
          "bg-slate-50 rounded border border-slate-200 flex items-center justify-center shrink-0 relative overflow-hidden transition-all",
          density === 'compact' ? "w-10 h-10" : density === 'standard' ? "w-16 h-16" : "w-32 h-32"
        )}>
          <div className="absolute inset-0 bg-gradient-to-tr from-slate-200 to-transparent opacity-20"></div>
          {result.previous ? (
            <img 
              src={formatImageUrl(result.previous.SOURCE_IMAGE_URL)} 
              alt="Previous"
              referrerPolicy="no-referrer"
              className="w-full h-full object-contain mix-blend-multiply relative z-10 p-1"
              loading="lazy"
              onError={(e) => {
                (e.target as HTMLImageElement).src = 'https://placehold.co/400?text=No+Image';
              }}
            />
          ) : (
            <span className={cn(
              "text-slate-300 font-bold uppercase z-10",
              density === 'large' ? "text-xs" : "text-[9px]"
            )}>Empty</span>
          )}
        </div>
        <div className="flex flex-col min-w-0 max-w-[300px]">
          <span className={cn(
            "font-bold text-slate-700 truncate block",
            density === 'large' ? "text-sm" : "text-[11px]"
          )}>
            {result.previous?.CSOR_IMAGE_NAME || '---'}
          </span>
          <span className="text-[9px] text-slate-400 truncate mt-1 uppercase font-mono tracking-widest">
            {result.previous ? 'V1.0 Baseline' : 'N/A'}
          </span>
        </div>
      </div>

      {/* Current Scored Version */}
      <div className={cn(
        "px-6 py-4 border-r border-slate-100 flex items-center gap-4 transition-colors group cursor-zoom-in",
        isChanged ? "bg-blue-50/20" : "bg-slate-50/20"
      )} onClick={() => onPreview(result)}>
        <div className={cn(
          "rounded border flex items-center justify-center shrink-0 shadow-sm relative overflow-hidden transition-all",
          isChanged ? "bg-white border-blue-200 ring-2 ring-blue-500/5" : "bg-white border-slate-200",
          density === 'compact' ? "w-10 h-10" : density === 'standard' ? "w-16 h-16" : "w-32 h-32"
        )}>
          {result.current ? (
            <img 
              src={formatImageUrl(result.current.SOURCE_IMAGE_URL)} 
              alt="Current"
              referrerPolicy="no-referrer"
              className="w-full h-full object-contain mix-blend-multiply p-1"
              loading="lazy"
              onError={(e) => {
                (e.target as HTMLImageElement).src = 'https://placehold.co/400?text=No+Image';
              }}
            />
          ) : (
            <span className={cn(
              "text-slate-300 font-bold uppercase",
              density === 'large' ? "text-xs" : "text-[9px]"
            )}>Removed</span>
          )}
        </div>
        <div className="flex flex-col min-w-0 max-w-[300px]">
          <span className={cn(
            "font-bold truncate",
            isChanged ? "text-blue-700 underline decoration-blue-200 underline-offset-2" : "text-slate-700",
            density === 'large' ? "text-sm" : "text-[11px]"
          )}>
            {result.current?.CSOR_IMAGE_NAME || '---'}
          </span>
          <span className={cn(
            "text-[9px] truncate mt-1 font-bold uppercase tracking-tighter",
            isChanged ? "text-blue-500" : "text-slate-400"
          )}>
            {isChanged ? 'New Scoring Engine' : 'No Change'}
          </span>
        </div>
      </div>

      {/* Status */}
      <div className="px-4 py-4 flex flex-col items-center justify-center min-w-[100px]">
        {isChanged ? (
          <div className="flex flex-col items-center gap-1.5">
            <div className="px-2 py-1 bg-amber-100 text-amber-700 text-[9px] font-bold rounded border border-amber-200 uppercase tracking-tight shadow-sm flex items-center gap-1">
              <RefreshCw className="w-2.5 h-2.5" />
              Updated
            </div>
            <div className="flex flex-wrap gap-1 justify-center max-w-[90px]">
              {result.changes.map(change => (
                <span key={change} className="px-1.5 py-0.5 bg-blue-50 text-blue-600 text-[7px] font-extrabold rounded border border-blue-100 uppercase tracking-tighter">
                  {change}
                </span>
              ))}
            </div>
          </div>
        ) : (
          <div className="px-2 py-1 bg-emerald-50 text-emerald-600 text-[9px] font-bold rounded border border-emerald-200 uppercase tracking-tight flex items-center gap-1">
            <CheckCircle2 className="w-2.5 h-2.5" />
            Matched
          </div>
        )}
      </div>
    </div>
  );
});

export default function App() {
  const [prevData, setPrevData] = useState<ImageItem[] | null>(null);
  const [newData, setNewData] = useState<ImageItem[] | null>(null);
  const [resultsFromWorker, setResultsFromWorker] = useState<ComparisonResult[]>([]);
  const [workerStats, setWorkerStats] = useState({ 
    total: 0,
    changed: 0,
    same: 0,
    reviewed: 0,
    ignored: 0,
    ignoredNoImages: 0,
    totalPrev: 0, 
    totalNew: 0, 
    totalMerged: 0 
  });
  const [isProcessing, setIsProcessing] = useState(false);
  const [processingProgress, setProcessingProgress] = useState('');
  const [isAuditActive, setIsAuditActive] = useState(false);
  const [filter, setFilter] = useState<FilterStatus>('changed');
  const [searchQuery, setSearchQuery] = useState('');
  const [sortConfig, setSortConfig] = useState<{ key: SortKey; order: SortOrder }>({ key: 'upc', order: 'asc' });
  const [previewItem, setPreviewItem] = useState<ComparisonResult | null>(null);
  const [userStatus, setUserStatus] = useState<Record<string, 'pending' | 'reviewed' | 'ignored'>>({});
  const [selectedUpcs, setSelectedUpcs] = useState<Set<string>>(new Set());
  const [listHeight, setListHeight] = useState(window.innerHeight - 250);
  const [rowDensity, setRowDensity] = useState<'compact' | 'standard' | 'large'>('standard');

  const rowHeights = {
    compact: 64,
    standard: 110,
    large: 200
  };

  const clearAll = () => {
    setPrevData(null);
    setNewData(null);
    setIsAuditActive(false);
    setUserStatus({});
    setSelectedUpcs(new Set());
    setSearchQuery('');
  };

  const listContainerRef = React.useRef<HTMLDivElement>(null);

  // Handle dynamic height for virtualization
  React.useEffect(() => {
    const updateHeight = () => {
      if (listContainerRef.current) {
        setListHeight(listContainerRef.current.clientHeight);
      } else {
        setListHeight(window.innerHeight - 320); // Fallback
      }
    };

    const observer = new ResizeObserver(updateHeight);
    if (listContainerRef.current) observer.observe(listContainerRef.current);
    
    updateHeight();
    window.addEventListener('resize', updateHeight);
    
    return () => {
      observer.disconnect();
      window.removeEventListener('resize', updateHeight);
    };
  }, [isAuditActive]);

  const [debouncedSearch, setDebouncedSearch] = useState('');
  React.useEffect(() => {
    const timer = setTimeout(() => setDebouncedSearch(searchQuery), 300);
    return () => clearTimeout(timer);
  }, [searchQuery]);

  const handleFileUpload = (type: 'prev' | 'new') => async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsProcessing(true);
    // Allow UI to update before heavy processing
    setTimeout(() => {
      const reader = new FileReader();
      reader.onload = (evt) => {
        try {
          const data = new Uint8Array(evt.target?.result as ArrayBuffer);
          const wb = XLSX.read(data, { type: 'array', cellDates: true, cellNF: false, cellText: false });
          const wsname = wb.SheetNames[0];
          const ws = wb.Sheets[wsname];
          const jsonData = XLSX.utils.sheet_to_json(ws, { defval: "" }) as ImageItem[];
          
          if (type === 'prev') setPrevData(jsonData);
          else setNewData(jsonData);
        } catch (err) {
          console.error('Processing error:', err);
          alert('Error processing file. Please ensure it is a valid Excel/CSV.');
        } finally {
          setIsProcessing(false);
        }
      };
      reader.readAsArrayBuffer(file);
    }, 100);
  };

  const getUpcValue = (item: any): string => {
    if (!item) return '';
    // Look for any key that contains "UPC", "GTIN", "IDENTIFIER", "EAN", "ITEM", or "CODE" case-insensitive
    const keys = Object.keys(item);
    const upcKey = keys.find(k => {
      const lower = k.toLowerCase();
      return lower.includes('upc') || 
             lower.includes('gtin') || 
             lower.includes('identifier') || 
             lower.includes('ean') ||
             (lower.includes('item') && lower.includes('id')) ||
             (lower.includes('item') && lower.includes('code'));
    });
    
    const val = upcKey ? item[upcKey] : (item.UPC || item.GTIN || item.ID);
    return String(val || '').trim();
  };

  // Background Worker Setup
  const workerRef = React.useRef<Worker | null>(null);

  React.useEffect(() => {
    workerRef.current = new Worker(new URL('./comparisonWorker.js', import.meta.url));
    workerRef.current.onmessage = (e) => {
      const { results, stats } = e.data;
      setResultsFromWorker(results);
      setWorkerStats(stats);
      setIsProcessing(false);
      setIsAuditActive(true);
      setProcessingProgress('');
    };
    return () => workerRef.current?.terminate();
  }, []);

  React.useEffect(() => {
    if (prevData && newData) {
      setIsProcessing(true);
      setProcessingProgress(debouncedSearch ? `Searching "${debouncedSearch}"...` : 'Analyzing records...');
      workerRef.current?.postMessage({ 
        prevData, 
        newData, 
        config: {
          filter,
          searchQuery: debouncedSearch,
          sortConfig,
          userStatus
        }
      });
    }
  }, [prevData, newData, filter, debouncedSearch, sortConfig, userStatus]);

  const filteredResults = resultsFromWorker;

  const stats = useMemo(() => {
    return workerStats;
  }, [workerStats]);

  const handleSelect = (upc: string) => {
    setSelectedUpcs(prev => {
      const next = new Set(prev);
      if (next.has(upc)) next.delete(upc);
      else next.add(upc);
      return next;
    });
  };

  const handleBulkAction = (action: 'reviewed' | 'ignored' | 'pending') => {
    setUserStatus(prev => {
      const next = { ...prev };
      selectedUpcs.forEach(upc => {
        next[upc] = action;
      });
      return next;
    });
    setSelectedUpcs(new Set());
  };

  const toggleSort = (key: SortKey) => {
    setSortConfig(prev => ({
      key,
      order: prev.key === key && prev.order === 'asc' ? 'desc' : 'asc'
    }));
  };

  const exportData = (mode: 'all' | 'changed' | 'selection') => {
    const filterFn = mode === 'changed' 
      ? (r: ComparisonResult) => r.hasChanged 
      : mode === 'selection' 
        ? (r: ComparisonResult) => selectedUpcs.has(r.upc)
        : () => true;

    const dataToExport = resultsFromWorker.filter(filterFn).map(r => ({
      UPC: r.upc,
      Audit_Status: r.auditStatus.toUpperCase(),
      Change_Status: r.hasChanged ? 'Changed' : 'Same',
      Change_Types: r.changes.join(', '),
      Baseline_Image_Name: r.previous?.CSOR_IMAGE_NAME || '---',
      Scored_Image_Name: r.current?.CSOR_IMAGE_NAME || '---',
      Baseline_URL: r.previous?.SOURCE_IMAGE_URL || '---',
      Scored_URL: r.current?.SOURCE_IMAGE_URL || '---',
      TABLE_ID: r.current?.TABLE_ID || r.previous?.TABLE_ID || '---'
    }));

    const ws = XLSX.utils.json_to_sheet(dataToExport);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Comparison_Audit");
    XLSX.writeFile(wb, `Image_Audit_${mode.toUpperCase()}_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  const getBase64Image = async (url: string): Promise<string | null> => {
    try {
      const response = await fetch(url);
      const blob = await response.blob();
      return new Promise((resolve) => {
        const reader = new FileReader();
        reader.onloadend = () => resolve(reader.result as string);
        reader.readAsDataURL(blob);
      });
    } catch {
      return null;
    }
  };

  const exportPDF = async () => {
    setIsProcessing(true);
    try {
      const doc = new jsPDF('p', 'mm', 'a4');
      const resultsToExport = filteredResults;
      
      const primaryColor = [15, 23, 42];
      doc.setFillColor(primaryColor[0], primaryColor[1], primaryColor[2]);
      doc.rect(0, 0, 210, 40, 'F');
      
      doc.setTextColor(255, 255, 255);
      doc.setFontSize(22);
      doc.setFont("helvetica", "bold");
      doc.text('IMAGE SCORING AUDIT', 15, 20);
      
      doc.setFontSize(10);
      doc.setFont("helvetica", "normal");
      doc.text(`Generated: ${new Date().toLocaleString()}`, 15, 30);
      doc.text(`View: ${filter.toUpperCase()} | Total Records: ${resultsToExport.length}`, 195, 30, { align: 'right' });
      
      let y = 50;
      
      for (const item of resultsToExport) {
        if (y > 220) {
          doc.addPage();
          y = 20;
        }

        doc.setFillColor(248, 250, 252);
        doc.rect(15, y, 180, 55, 'F'); // Increased height for image area
        doc.setDrawColor(226, 232, 240);
        doc.rect(15, y, 180, 55, 'S');

        doc.setTextColor(primaryColor[0], primaryColor[1], primaryColor[2]);
        doc.setFontSize(10);
        doc.setFont("helvetica", "bold");
        doc.text(`UPC: ${item.upc}`, 20, y + 8);
        
        doc.setFontSize(7);
        doc.text(`AUDIT STATE: ${item.auditStatus.toUpperCase()}`, 20, y + 14);
        
        if (item.hasChanged) {
          doc.setTextColor(220, 38, 38);
          doc.text(`CHANGES: ${item.changes.join(', ')}`, 20, y + 18);
        }

        // Images in PDF
        const prevImg = await getBase64Image(formatImageUrl(item.previous?.SOURCE_IMAGE_URL));
        const currImg = await getBase64Image(formatImageUrl(item.current?.SOURCE_IMAGE_URL));

        if (prevImg) {
          try { doc.addImage(prevImg, 'JPEG', 20, y + 25, 25, 25); } catch(e) {}
        }
        if (currImg) {
          try { doc.addImage(currImg, 'JPEG', 105, y + 25, 25, 25); } catch(e) {}
        }

        doc.setTextColor(100, 116, 139);
        doc.setFontSize(7);
        doc.text('BASELINE (V1)', 50, y + 35);
        doc.text('SCORED DATA', 135, y + 35);
        
        doc.setTextColor(primaryColor[0], primaryColor[1], primaryColor[2]);
        doc.setFontSize(6);
        doc.text(String(item.previous?.CSOR_IMAGE_NAME || 'EMPTY').substring(0, 45), 50, y + 40);
        doc.text(String(item.current?.CSOR_IMAGE_NAME || 'REMOVED').substring(0, 45), 135, y + 40);

        if (item.hasChanged) {
          doc.setDrawColor(234, 179, 8);
          doc.setLineWidth(0.5);
          doc.line(15, y, 15, y + 55);
          doc.setLineWidth(0.2);
        }

        y += 62;
      }
      
      doc.save(`Audit_Export_${new Date().toISOString().split('T')[0]}.pdf`);
    } catch (err) {
      console.error('PDF Export failed:', err);
    } finally {
      setIsProcessing(false);
    }
  };

  return (
    <div className="h-screen flex flex-col bg-[#F8FAFC] text-slate-900 font-sans overflow-hidden">
      {/* Header */}
      <header className="bg-[#0F172A] text-white px-6 py-4 flex justify-between items-center shrink-0 shadow-lg relative z-50">
        <div className="flex items-center gap-3">
          <div className="w-8 h-8 bg-blue-600 rounded flex items-center justify-center">
            <Database className="w-5 h-5 text-white" />
          </div>
          <h1 className="text-xl font-semibold tracking-tight uppercase tracking-widest text-[14px]">Visual Scored Auditor</h1>
        </div>

        <div className="flex items-center gap-4 text-xs">
          {prevData && (
            <span className="px-3 py-1.5 bg-slate-700/50 rounded-full border border-slate-600/50">
              Baseline: <strong className="ml-1 text-slate-100">{workerStats.totalPrev || prevData.length}</strong>
            </span>
          )}
          {newData && (
            <span className="px-3 py-1.5 bg-slate-700/50 rounded-full border border-slate-600/50">
              Target: <strong className="ml-1 text-blue-300">{workerStats.totalNew || newData.length}</strong>
            </span>
          )}
          {isAuditActive && (
            <span className="px-3 py-1.5 bg-blue-500/20 rounded-full border border-blue-500/30 text-blue-300">
              Unique Merged: <strong className="ml-1 text-white">{workerStats.totalMerged}</strong>
            </span>
          )}
          {workerStats.ignoredNoImages > 0 && (
            <span className="px-3 py-1.5 bg-red-950/30 rounded-full border border-red-900/30 text-red-300" title="Records with no images in both versions were automatically ignored">
              Ignored (No Images): <strong className="ml-1 text-white">{workerStats.ignoredNoImages}</strong>
            </span>
          )}
          {resultsFromWorker.length > 0 && isAuditActive && (
            <div className="flex items-center gap-2">
              <button 
                onClick={clearAll}
                className="px-4 py-2 text-slate-400 hover:text-white transition-colors flex items-center gap-2 text-[10px] font-bold uppercase"
              >
                <RefreshCw className="w-3.5 h-3.5" /> Reset All
              </button>
              <div className="w-px h-4 bg-slate-700 mx-2"></div>
              <div className="relative group">
                <button className="bg-slate-700 hover:bg-slate-600 px-4 py-2 rounded-md font-bold transition-all flex items-center gap-2">
                  <Download className="w-3.5 h-3.5" /> Export Excel
                </button>
                <div className="absolute right-0 top-full mt-2 w-48 bg-white rounded-xl shadow-2xl border border-slate-200 hidden group-hover:block overflow-hidden z-[100]">
                  <button onClick={() => exportData('all')} className="w-full px-4 py-3 text-left text-xs font-bold text-slate-700 hover:bg-slate-50 border-b border-slate-100 flex items-center gap-2 italic">
                    <FileText className="w-3.5 h-3.5 text-blue-500" /> All Items Table
                  </button>
                  <button onClick={() => exportData('changed')} className="w-full px-4 py-3 text-left text-xs font-bold text-slate-700 hover:bg-slate-50 flex items-center gap-2 italic">
                    <AlertCircle className="w-3.5 h-3.5 text-amber-500" /> Only Changed Records
                  </button>
                </div>
              </div>
              <button 
                onClick={exportPDF}
                className="bg-red-600 hover:bg-red-500 px-4 py-2 rounded-md font-bold transition-all flex items-center gap-2 shadow-sm"
              >
                <Camera className="w-3.5 h-3.5" /> PDF Report
              </button>
            </div>
          )}
        </div>
      </header>

      {/* Nav / Filters */}
      <nav className="bg-white border-b border-slate-200 px-6 py-3 flex justify-between items-center shrink-0 shadow-sm z-40">
        <div className="flex gap-1 p-1 bg-slate-100 rounded-lg border border-slate-200">
          {[
            { id: 'all', label: 'All', count: stats.total, color: 'slate' },
            { id: 'changed', label: 'Δ Changed', count: stats.changed, color: 'blue' },
            { id: 'reviewed', label: '✓ Reviewed', count: stats.reviewed, color: 'emerald' },
            { id: 'ignored', label: '✕ Ignored', count: stats.ignored, color: 'orange' }
          ].map(f => (
            <button 
              key={f.id}
              onClick={() => setFilter(f.id as FilterStatus)}
              className={cn(
                "px-3 py-1.5 rounded-md text-[10px] font-bold transition-all uppercase tracking-tight",
                filter === f.id 
                  ? `bg-white shadow-sm text-${f.color === 'slate' ? 'slate-800' : f.color + '-600'}` 
                  : "text-slate-500 hover:text-slate-700 hover:bg-slate-200/50"
              )}
            >
              {f.label} ({f.count})
            </button>
          ))}
        </div>

        <div className="flex items-center gap-6">
          {isAuditActive && (
            <div className="flex items-center gap-2 p-1 bg-slate-100 rounded-lg border border-slate-200">
              {(['compact', 'standard', 'large'] as const).map(d => (
                <button 
                  key={d}
                  onClick={() => setRowDensity(d)}
                  className={cn(
                    "p-1.5 rounded transition-all",
                    rowDensity === d ? "bg-white shadow-sm text-blue-600" : "text-slate-400 hover:text-slate-600"
                  )}
                  title={`Density: ${d}`}
                >
                  {d === 'compact' && <Layout className="w-3.5 h-3.5 rotate-90" />}
                  {d === 'standard' && <Columns className="w-3.5 h-3.5" />}
                  {d === 'large' && <Maximize2 className="w-3.5 h-3.5" />}
                </button>
              ))}
            </div>
          )}

          <div className="text-[10px] text-slate-500 flex gap-4 uppercase font-bold tracking-widest">
            <span>Score Match: <span className="text-emerald-600 font-mono text-[12px]">{( (stats.same / (stats.total || 1)) * 100).toFixed(1)}%</span></span>
            <span>Audit Delta: <span className="text-blue-600 font-mono text-[12px]">{( (stats.changed / (stats.total || 1)) * 100).toFixed(1)}%</span></span>
          </div>
          <div className="relative">
            <input 
              type="text" 
              placeholder="Quick UPC find..." 
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
        <div className="grid grid-cols-[48px_160px_1fr_1fr_120px] border-b border-slate-200 bg-slate-50/80 text-[10px] font-bold uppercase tracking-wider text-slate-500 z-30 select-none">
          <div className="flex items-center justify-center border-r border-slate-200">
            <button 
              onClick={() => {
                if (selectedUpcs.size === filteredResults.length) setSelectedUpcs(new Set());
                else setSelectedUpcs(new Set(filteredResults.map(r => r.upc)));
              }}
              className={cn(
                "w-5 h-5 rounded flex items-center justify-center transition-all",
                selectedUpcs.size === filteredResults.length && filteredResults.length > 0 ? "bg-blue-600 text-white" : "border border-slate-300"
              )}
            >
              {selectedUpcs.size === filteredResults.length && filteredResults.length > 0 && <CheckSquare className="w-3.5 h-3.5" />}
            </button>
          </div>
          <div className="px-6 py-3 cursor-pointer hover:bg-slate-100 flex items-center justify-between group" onClick={() => toggleSort('upc')}>
            <span>UPC / Identifier</span>
            {sortConfig.key === 'upc' ? (
              sortConfig.order === 'asc' ? <ChevronUp className="w-3 h-3 text-blue-500" /> : <ChevronDown className="w-3 h-3 text-blue-500" />
            ) : <RefreshCw className="w-3 h-3 opacity-0 group-hover:opacity-30" />}
          </div>
          <div className="px-6 py-3 border-l border-slate-200 cursor-pointer hover:bg-slate-100 flex items-center justify-between group" onClick={() => toggleSort('prev_name')}>
            <span>Baseline Assignment (V1)</span>
            {sortConfig.key === 'prev_name' ? (
              sortConfig.order === 'asc' ? <ChevronUp className="w-3 h-3 text-blue-500" /> : <ChevronDown className="w-3 h-3 text-blue-500" />
            ) : <RefreshCw className="w-3 h-3 opacity-0 group-hover:opacity-30" />}
          </div>
          <div className="px-6 py-3 border-l border-slate-200 cursor-pointer hover:bg-slate-100 flex items-center justify-between group" onClick={() => toggleSort('curr_name')}>
            <span>New Scoring Target</span>
            {sortConfig.key === 'curr_name' ? (
              sortConfig.order === 'asc' ? <ChevronUp className="w-3 h-3 text-blue-500" /> : <ChevronDown className="w-3 h-3 text-blue-500" />
            ) : <RefreshCw className="w-3 h-3 opacity-0 group-hover:opacity-30" />}
          </div>
          <div className="px-6 py-3 border-l border-slate-200 cursor-pointer hover:bg-slate-100 flex items-center justify-between group" onClick={() => toggleSort('auditStatus')}>
            <span>Audit State</span>
            {sortConfig.key === 'auditStatus' ? (
              sortConfig.order === 'asc' ? <ChevronUp className="w-3 h-3 text-blue-500" /> : <ChevronDown className="w-3 h-3 text-blue-500" />
            ) : <RefreshCw className="w-3 h-3 opacity-0 group-hover:opacity-30" />}
          </div>
        </div>

        {/* List Content */}
        <div className="flex-1 relative" id="audit-table-body">
          {!isAuditActive ? (
            <div className="h-full overflow-y-auto no-scrollbar bg-slate-50">
              <div className="flex flex-col items-center justify-center p-20 min-h-full">
                <div className="text-center mb-12">
                  <h2 className="text-3xl font-bold tracking-tight text-slate-800">Ready for Audit</h2>
                  <p className="text-slate-500 mt-2">Upload both datasets to begin cross-reference analysis</p>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-8 w-full max-w-4xl">
                  {/* Baseline Column */}
                  <div className="relative group">
                    <input type="file" accept=".xlsx,.xls,.csv" onChange={handleFileUpload('prev')} className="absolute inset-0 opacity-0 cursor-pointer z-10" />
                    <div className={cn(
                      "h-64 rounded-3xl border-2 border-dashed flex flex-col items-center justify-center gap-4 transition-all shadow-sm",
                      prevData ? "border-emerald-500 bg-emerald-50/20" : "border-slate-200 bg-white group-hover:border-blue-500/50 group-hover:bg-blue-50/10"
                    )}>
                      <div className={cn(
                        "w-16 h-16 rounded-2xl flex items-center justify-center transition-colors",
                        prevData ? "bg-emerald-100 text-emerald-600" : "bg-slate-50 text-slate-400 group-hover:text-blue-500"
                      )}>
                        {prevData ? <CheckCircle2 className="w-8 h-8" /> : <FileUp className="w-8 h-8" />}
                      </div>
                      <div className="text-center px-6">
                        <p className={cn("font-bold", prevData ? "text-emerald-700" : "text-slate-700")}>
                          {prevData ? "Baseline Verified" : "Baseline Data (V1)"}
                        </p>
                        <p className="text-[10px] text-slate-400 uppercase tracking-widest mt-1">
                          {prevData ? `${prevData.length.toLocaleString()} Records Active` : "Upload Previous Reference"}
                        </p>
                      </div>
                    </div>
                  </div>

                  {/* Scored Data Column */}
                  <div className="relative group">
                    <input type="file" accept=".xlsx,.xls,.csv" onChange={handleFileUpload('new')} className="absolute inset-0 opacity-0 cursor-pointer z-10" />
                    <div className={cn(
                      "h-64 rounded-3xl border-2 border-dashed flex flex-col items-center justify-center gap-4 transition-all shadow-sm",
                      newData ? "border-blue-500 bg-blue-50/20" : "border-slate-200 bg-white group-hover:border-blue-500/50 group-hover:bg-blue-50/10"
                    )}>
                      <div className={cn(
                        "w-16 h-16 rounded-2xl flex items-center justify-center transition-colors",
                        newData ? "bg-blue-100 text-blue-600" : "bg-slate-50 text-slate-400 group-hover:text-blue-500"
                      )}>
                        {newData ? <CheckCircle2 className="w-8 h-8" /> : <FileUp className="w-8 h-8" />}
                      </div>
                      <div className="text-center px-6">
                        <p className={cn("font-bold", newData ? "text-blue-700" : "text-slate-700")}>
                          {newData ? "Scored Data Synced" : "SCORED DATA"}
                        </p>
                        <p className="text-[10px] text-slate-400 uppercase tracking-widest mt-1">
                          {newData ? `${newData.length.toLocaleString()} Records Cached` : "Upload New Scored Data"}
                        </p>
                      </div>
                    </div>
                  </div>
                </div>

                <div className="mt-12 flex flex-col items-center gap-6">
                  {prevData && newData ? (
                    <motion.button
                      initial={{ scale: 0.9, opacity: 0 }}
                      animate={{ scale: 1, opacity: 1 }}
                      onClick={() => setIsAuditActive(true)}
                      className="px-12 py-4 bg-blue-600 hover:bg-blue-500 text-white rounded-2xl font-bold text-lg shadow-xl shadow-blue-500/20 flex items-center gap-3 active:scale-95 transition-all"
                    >
                      <Columns className="w-6 h-6" />
                      Sync & Start Audit
                    </motion.button>
                  ) : (
                    <div className="flex items-center gap-12 text-slate-400">
                      <div className={cn("flex flex-col items-center gap-2 opacity-50", prevData && "opacity-100 text-emerald-600")}>
                        <span className="text-[10px] font-bold">TASK 01</span>
                        <span className="text-xs font-bold uppercase italic">V1 Data</span>
                      </div>
                      <ArrowRight className="w-4 h-4 opacity-20" />
                      <div className={cn("flex flex-col items-center gap-2 opacity-50", newData && "opacity-100 text-blue-600")}>
                        <span className="text-[10px] font-bold">TASK 02</span>
                        <span className="text-xs font-bold uppercase italic">Scored Data</span>
                      </div>
                    </div>
                  )}
                </div>
              </div>
            </div>
          ) : (
            <div className="h-full">
              {isProcessing && (
                <div className="absolute inset-0 bg-slate-900/40 backdrop-blur-md z-[500] flex flex-col items-center justify-center text-white">
                  <div className="relative">
                    <div className="w-20 h-20 border-4 border-blue-500/20 border-t-blue-500 rounded-full animate-spin"></div>
                    <div className="absolute inset-0 flex items-center justify-center">
                      <RefreshCw className="w-8 h-8 text-blue-400 animate-pulse" />
                    </div>
                  </div>
                  <p className="mt-8 text-sm font-black uppercase tracking-[0.3em] animate-pulse drop-shadow-lg">
                    {processingProgress || 'Processing Dataset...'}
                  </p>
                  <p className="mt-2 text-[10px] text-slate-300 font-mono">Comparing large datasets. Please wait.</p>
                </div>
              )}

              {/* Bulk Action Bar */}
              <AnimatePresence>
                {selectedUpcs.size > 0 && (
                  <motion.div 
                    initial={{ y: 50, opacity: 0 }}
                    animate={{ y: 0, opacity: 1 }}
                    exit={{ y: 50, opacity: 0 }}
                    className="absolute bottom-10 left-1/2 -translate-x-1/2 bg-[#0F172A] text-white px-6 py-3 rounded-2xl shadow-2xl z-[100] flex items-center gap-8 border border-white/10"
                  >
                    <div className="flex items-center gap-3 pr-8 border-r border-white/20 font-bold italic">
                      <span className="text-blue-400 font-mono text-lg">{selectedUpcs.size}</span>
                      <span className="text-xs uppercase tracking-widest text-slate-400">Rows Active</span>
                    </div>
                    <div className="flex items-center gap-3">
                      <button 
                        onClick={() => handleBulkAction('reviewed')}
                        className="px-4 py-2 bg-emerald-600 hover:bg-emerald-500 rounded-lg text-[10px] font-bold uppercase tracking-widest flex items-center gap-2 transition-all"
                      >
                        <CheckCircle2 className="w-3.5 h-3.5" /> Reviewed
                      </button>
                      <button 
                        onClick={() => handleBulkAction('ignored')}
                        className="px-4 py-2 bg-slate-700 hover:bg-slate-600 rounded-lg text-[10px] font-bold uppercase tracking-widest flex items-center gap-2 transition-all"
                      >
                        <EyeOff className="w-3.5 h-3.5" /> Ignore
                      </button>
                      <button 
                        onClick={() => handleBulkAction('pending')}
                        className="px-4 py-2 text-red-400 hover:text-red-300 rounded-lg text-[10px] font-bold uppercase tracking-widest flex items-center gap-2 transition-all"
                      >
                        <Trash2 className="w-3.5 h-3.5" /> Reset
                      </button>
                    </div>
                  </motion.div>
                )}
              </AnimatePresence>

              <div ref={listContainerRef} className="flex-1 relative overflow-hidden bg-white">
                {filteredResults.length > 0 ? (
                  <div className="h-full">
                    {/* Virtualized List */}
                    <List
                      style={{ height: listHeight, width: '100%' }} 
                      rowCount={filteredResults.length}
                      rowHeight={rowHeights[rowDensity]} 
                      className="scrollbar-custom"
                      rowProps={{}}
                      rowComponent={({ index, style }: any) => (
                        <ComparisonRow 
                          key={filteredResults[index].upc}
                          result={filteredResults[index]} 
                          onPreview={(res) => setPreviewItem(res)}
                          isSelected={selectedUpcs.has(filteredResults[index].upc)}
                          onSelect={handleSelect}
                          style={style}
                          density={rowDensity}
                        />
                      )}
                    />
                  </div>
                ) : (
                  <motion.div initial={{ opacity: 0 }} animate={{ opacity: 1 }} className="py-20 flex flex-col items-center text-slate-400 font-medium h-full justify-center">
                    <Search className="w-12 h-12 mb-4 opacity-10" />
                    <p className="text-sm">No results match your active filters or sorting</p>
                  </motion.div>
                )}
              </div>
          </div>
        )}
      </div>
    </main>

      {/* Footer */}
      <footer className="bg-white border-t border-slate-200 px-6 py-2.5 flex justify-between items-center text-[10px] text-slate-500 shrink-0 font-bold uppercase tracking-wider relative z-50 shadow-[0_-4px_6px_-1px_rgba(0,0,0,0.05)]">
        <div className="flex items-center gap-6">
          <div className="flex items-center gap-2">
            <div className="w-1.5 h-1.5 rounded-full bg-emerald-500"></div> System Verified
          </div>
          <div className="flex items-center gap-2">
            <div className="w-1.5 h-1.5 rounded-full bg-blue-500"></div> Showing {filteredResults.length} / {workerStats.totalMerged || resultsFromWorker.length}
          </div>
        </div>
        <div className="flex items-center gap-4">
          <span className="text-slate-400">Server Time: {new Date().toLocaleTimeString()}</span>
          <span className="w-px h-2 bg-slate-300"></span>
          <span className="text-blue-600">Image Scoring Auditor V2.5</span>
        </div>
      </footer>

      {/* Preview Modal */}
      <AnimatePresence>
        {previewItem && (
          <motion.div 
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 z-[200] bg-slate-900/90 backdrop-blur-md flex items-center justify-center p-8"
            onClick={() => setPreviewItem(null)}
          >
            <motion.div 
              initial={{ scale: 0.9, opacity: 0, y: 20 }}
              animate={{ scale: 1, opacity: 1, y: 0 }}
              exit={{ scale: 0.9, opacity: 0, y: 20 }}
              className="bg-white rounded-3xl w-full max-w-6xl overflow-hidden shadow-2xl flex flex-col max-h-[90vh]"
              onClick={(e) => e.stopPropagation()}
            >
              <div className="px-8 py-6 border-b border-slate-100 flex items-center justify-between bg-slate-50/50">
                <div className="flex items-center gap-4">
                  <div className="w-10 h-10 bg-blue-600 text-white rounded-xl flex items-center justify-center">
                    <ImageIcon className="w-5 h-5" />
                  </div>
                  <div>
                    <h2 className="text-lg font-bold tracking-tight">Audit Deep Dive: {previewItem.upc}</h2>
                    <p className="text-xs text-slate-400 font-bold uppercase tracking-widest mt-0.5">Asset Comparison Inspection</p>
                  </div>
                </div>
                <button 
                  onClick={() => setPreviewItem(null)}
                  className="w-10 h-10 rounded-full hover:bg-slate-200 transition-colors flex items-center justify-center text-slate-400"
                >
                  <RefreshCw className="w-5 h-5 hover:rotate-180 transition-transform duration-500" />
                </button>
              </div>

              <div className="flex-1 overflow-y-auto grid grid-cols-2 divide-x divide-slate-100 p-8 gap-8">
                {/* Left Side */}
                <div className="flex flex-col gap-6">
                  <div className="flex items-center justify-between">
                    <span className="px-3 py-1 bg-slate-100 text-slate-600 text-[10px] font-bold rounded-full uppercase">Baseline Version</span>
                    <span className="text-xs font-mono text-slate-400">ID: {previewItem.previous?.TABLE_ID || '---'}</span>
                  </div>
                  <div className="aspect-square bg-slate-50 rounded-2xl border border-slate-200 flex items-center justify-center overflow-hidden p-8 shadow-inner">
                    <img 
                      src={formatImageUrl(previewItem.previous?.SOURCE_IMAGE_URL)} 
                      alt="Baseline"
                      className="max-w-full max-h-full object-contain mix-blend-multiply"
                      referrerPolicy="no-referrer"
                    />
                  </div>
                  <div className="bg-slate-50 rounded-xl p-4 border border-slate-100">
                    <p className="text-[10px] font-bold text-slate-400 uppercase mb-2">Baseline Filename</p>
                    <p className="text-sm font-mono break-all font-bold text-slate-600">{previewItem.previous?.CSOR_IMAGE_NAME || 'Not Assigned'}</p>
                  </div>
                </div>

                {/* Right Side */}
                <div className="flex flex-col gap-6 pl-8">
                  <div className="flex items-center justify-between">
                    <span className="px-3 py-1 bg-blue-100 text-blue-600 text-[10px] font-bold rounded-full uppercase">New Scored Version</span>
                    <span className="text-xs font-mono text-slate-400">ID: {previewItem.current?.TABLE_ID || '---'}</span>
                  </div>
                  <div className={cn(
                    "aspect-square rounded-2xl border flex items-center justify-center overflow-hidden p-8 shadow-inner transition-colors",
                    previewItem.hasChanged ? "bg-blue-50/30 border-blue-200 ring-8 ring-blue-500/5" : "bg-slate-50 border-slate-200"
                  )}>
                    <img 
                      src={formatImageUrl(previewItem.current?.SOURCE_IMAGE_URL)} 
                      alt="Target"
                      className="max-w-full max-h-full object-contain mix-blend-multiply"
                      referrerPolicy="no-referrer"
                    />
                  </div>
                  <div className={cn(
                    "rounded-xl p-4 border transition-colors",
                    previewItem.hasChanged ? "bg-blue-50 border-blue-100" : "bg-slate-50 border-slate-100"
                  )}>
                    <p className="text-[10px] font-bold text-blue-400 uppercase mb-2">New Target Filename</p>
                    <p className={cn(
                      "text-sm font-mono break-all font-bold",
                      previewItem.hasChanged ? "text-blue-700" : "text-slate-600"
                    )}>{previewItem.current?.CSOR_IMAGE_NAME || 'Removed'}</p>
                  </div>
                </div>
              </div>

              {previewItem.hasChanged && (
                <div className="px-8 py-4 bg-amber-50 border-t border-amber-100 flex items-center gap-3">
                  <AlertCircle className="w-4 h-4 text-amber-500" />
                  <span className="text-xs font-bold text-amber-700 uppercase tracking-tight">
                    Change Detection: {previewItem.changes.join(' • ')}
                  </span>
                </div>
              )}
            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>

      {/* Processing Loader */}
      {isProcessing && (
        <div className="fixed inset-0 bg-slate-900/40 backdrop-blur-sm z-[300] flex items-center justify-center">
          <div className="bg-white p-10 rounded-[32px] shadow-2xl flex flex-col items-center gap-6 border border-slate-100">
            <div className="relative">
              <div className="w-16 h-16 border-4 border-blue-100 rounded-full"></div>
              <div className="w-16 h-16 border-4 border-blue-600 border-t-transparent rounded-full animate-spin absolute inset-0"></div>
            </div>
            <div className="text-center">
              <p className="font-bold text-slate-800 text-lg tracking-tight">Syncing Repositories</p>
              <p className="text-xs text-slate-400 font-bold uppercase tracking-widest mt-1">Analyzing cross-referenced data</p>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

