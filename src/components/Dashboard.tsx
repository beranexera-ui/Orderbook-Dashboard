import React, { useState, useRef, useCallback, useMemo, useEffect } from 'react';
import { UploadCloud, FileSpreadsheet, AlertCircle, Search, Package, CheckCircle, TrendingUp, AlertTriangle, Loader2, Printer, BarChart2, Table } from 'lucide-react';
import * as XLSX from 'xlsx';
import { useVirtualizer } from '@tanstack/react-virtual';
import { db } from '../firebase';
import { doc, setDoc, onSnapshot, writeBatch, collection, getDocs } from 'firebase/firestore';
import { KPIView } from './KPIView';
import { SewOutReportModal } from './SewOutReportModal';
import { PrintOrientationModal } from './PrintOrientationModal';

export function MultiSelectDropdown({
  label,
  options,
  selectedValues,
  onChange
}: {
  label: string;
  options: string[];
  selectedValues: string[];
  onChange: (values: string[]) => void;
}) {
  const [isOpen, setIsOpen] = useState(false);
  const containerRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (containerRef.current && !containerRef.current.contains(event.target as Node)) {
        setIsOpen(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  const toggleOption = (option: string) => {
    if (selectedValues.includes(option)) {
      onChange(selectedValues.filter(v => v !== option));
    } else {
      onChange([...selectedValues, option]);
    }
  };

  const displayText = selectedValues.length === 0 
    ? label 
    : selectedValues.length === 1 
      ? selectedValues[0] 
      : `${label} (${selectedValues.length})`;

  return (
    <div className="relative flex-1 sm:min-w-[140px]" ref={containerRef}>
      <button
        onClick={() => setIsOpen(!isOpen)}
        className="w-full flex items-center justify-between py-1.5 md:py-2 pl-2 md:pl-3 pr-2 md:pr-3 bg-slate-50 hover:bg-slate-100/50 border border-slate-200 rounded-lg text-[10px] md:text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 text-slate-700 text-left min-w-0 sm:min-w-[140px] transition-all"
      >
        <span className="truncate pr-1 md:pr-2 font-medium">{displayText}</span>
        <svg className={`w-3 h-3 md:w-4 md:h-4 text-slate-400 transition-transform flex-shrink-0 ${isOpen ? 'rotate-180' : ''}`} fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M19 9l-7 7-7-7"></path></svg>
      </button>
      
      {isOpen && (
        <div className="absolute z-50 w-full min-w-[200px] mt-1 bg-white border border-slate-200 rounded-lg shadow-lg max-h-60 overflow-y-auto">
          <div className="p-1.5 space-y-0.5">
            {options.map(option => (
              <label key={option} className="flex items-center p-2 hover:bg-slate-50 rounded-md cursor-pointer transition-colors">
                <input 
                  type="checkbox" 
                  checked={selectedValues.includes(option)}
                  onChange={() => toggleOption(option)}
                  className="w-4 h-4 text-indigo-600 rounded border-slate-300 focus:ring-indigo-500 cursor-pointer"
                />
                <span className="ml-2 text-sm text-slate-700 truncate">{option}</span>
              </label>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

export interface ProductionOrder {
  id: string;
  legacyId?: string;
  buyer: string;
  groupTechClass: string;
  buyerDivisionName: string;
  styleNo: string;
  custStyleNo: string;
  vpoNo: string;
  shipmentMode: string;
  colorCode: string;
  colorName: string;
  destination: string;
  packMethod: string;
  scheduleNo: string;
  planDelDate: string;
  weekNo: string;
  coQty: number;
  cumSewInQty: number;
  cumSewOutQty: number;
  cumSewOutRejQty: number;
  cumCTNQty: number;
  statusText: string;
  deliveredQty: number;
  remark?: string;
}

function RemarkInput({
  initialValue,
  updatedAt,
  rowId,
  onSave
}: {
  initialValue: string;
  updatedAt: number | null;
  rowId: string;
  onSave: (id: string, text: string) => void;
}) {
  const [value, setValue] = useState(initialValue);
  // Auto-lock ONLY if there is an existing remark. If empty, keep it unlocked.
  const [isLocked, setIsLocked] = useState(initialValue.trim() !== "");
  const [showPrompt, setShowPrompt] = useState(false);
  const [passcode, setPasscode] = useState("");
  const [errorMsg, setErrorMsg] = useState("");
  const inputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    setValue(initialValue);
    // If the value gets cleared (or loaded as empty), unlock it. 
    // If it gets loaded with a value, lock it.
    setIsLocked(initialValue.trim() !== "");
  }, [initialValue]);

  const getTooltipInfo = () => {
    if (!updatedAt || !initialValue.trim()) return undefined;
    const date = new Date(updatedAt);
    
    const formattedDate = date.toLocaleString('en-US', { 
       year: 'numeric', month: '2-digit', day: '2-digit', 
       hour: '2-digit', minute: '2-digit', hour12: true 
    });

    const day = date.getDay(); // 0 is Sunday
    const hour = date.getHours(); 
    
    // ISO week number calculation to alternate shifts
    const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
    const dayNum = d.getUTCDay() || 7;
    d.setUTCDate(d.getUTCDate() + 4 - dayNum);
    const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
    const weekNo = Math.ceil((((d.getTime() - yearStart.getTime()) / 86400000) + 1) / 7);

    // Swap shifts every week (Current week is odd (37), where Morning is A and Evening is B)
    const morningShift = weekNo % 2 !== 0 ? 'A' : 'B';
    const eveningShift = weekNo % 2 !== 0 ? 'B' : 'A';
    
    let currentShift = "Off-hours";
    if (day !== 0) { // Not Sunday
       if (hour >= 6 && hour < 14) { 
           currentShift = `Shift ${morningShift}`;
       } else if (hour >= 14 && hour < 22) { 
           currentShift = `Shift ${eveningShift}`;
       }
    } else {
       currentShift = "Holiday (Sunday)";
    }

    return `Updated: ${formattedDate} | ${currentShift}`;
  };

  const handleDoubleClick = () => {
    if (isLocked) {
      setShowPrompt(true);
      setPasscode("");
      setErrorMsg("");
    }
  };

  const handleUnlock = () => {
    if (passcode === '1125') {
      setIsLocked(false);
      setShowPrompt(false);
      setTimeout(() => {
        if (inputRef.current) {
          inputRef.current.focus();
        }
      }, 50);
    } else {
      setErrorMsg("Incorrect passcode!");
    }
  };

  return (
    <>
      <input 
        ref={inputRef}
        type="text" 
        value={value}
        readOnly={isLocked}
        onDoubleClick={handleDoubleClick}
        onChange={(e) => !isLocked && setValue(e.target.value)}
        onBlur={() => {
          if (!isLocked) {
            if (value !== initialValue) {
              onSave(rowId, value);
            }
            // Auto-lock again ONLY if the input is not empty after saving
            if (value.trim() !== "") {
              setIsLocked(true);
            }
          }
        }}
        onKeyDown={(e) => {
          if (e.key === 'Enter') {
            e.currentTarget.blur();
          }
        }}
        title={getTooltipInfo()}
        placeholder={isLocked ? "" : "Add remark..."}
        list="remark-suggestions"
        className={`w-full text-sm bg-transparent border-0 border-b px-1 py-1 transition-colors ${
          isLocked 
            ? 'border-transparent cursor-pointer hover:bg-slate-100/50 focus:ring-0 focus:outline-none' 
            : 'border-slate-300 hover:border-slate-300 focus:border-indigo-500 focus:ring-0'
        }`}
      />
      
      {showPrompt && (
        <div className="fixed inset-0 z-[100] flex items-center justify-center bg-slate-900/50 backdrop-blur-sm p-4">
          <div className="bg-white p-5 rounded-xl shadow-xl w-full max-w-xs border border-slate-200">
            <h3 className="text-base font-bold text-slate-900 mb-1">Unlock Remark</h3>
            <p className="text-xs text-slate-500 mb-4">Enter passcode to edit this remark.</p>
            <input 
              type="password" 
              autoFocus
              value={passcode}
              onChange={(e) => {
                setPasscode(e.target.value);
                setErrorMsg("");
              }}
              onKeyDown={(e) => {
                if (e.key === 'Enter') handleUnlock();
                if (e.key === 'Escape') setShowPrompt(false);
              }}
              className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 text-sm mb-2"
              placeholder="Enter passcode..."
            />
            {errorMsg && <p className="text-xs text-red-500 mb-2">{errorMsg}</p>}
            <div className="flex justify-end gap-2 mt-4">
              <button 
                onClick={() => setShowPrompt(false)}
                className="px-4 py-2 text-sm font-medium text-slate-600 hover:bg-slate-100 rounded-lg transition-colors"
              >
                Cancel
              </button>
              <button 
                onClick={handleUnlock}
                className="px-4 py-2 text-sm font-medium text-white bg-indigo-600 hover:bg-indigo-700 rounded-lg transition-colors"
              >
                Unlock
              </button>
            </div>
          </div>
        </div>
      )}
    </>
  );
}

const TABLE_COLUMN_WIDTHS = [
  '120px', // Plan Del Date
  '90px',  // WEEK NO
  '130px', // Buyer
  '140px', // Style No
  '130px', // VPO No
  '120px', // Shipment Mode
  '105px', // Color Code
  '160px', // Color Name
  '115px', // Destination
  '140px', // Pack Method
  '120px', // Schedule No
  '100px', // CO Qty
  '120px', // Cum Sew In Qty
  '125px', // Cum SewOut Qty
  '115px', // Cum CTN Qty
  '115px', // Status
  '220px', // Remark
  '115px', // Delivered Qty
];

export function getSixWeeksWindow(): { start: Date; end: Date; weekNumbers: Set<string> } {
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const dayOfWeek = today.getDay(); // 0 is Sun, 1 is Mon...
  const diffToMonday = today.getDate() - dayOfWeek + (dayOfWeek === 0 ? -6 : 1);
  const start = new Date(today.getFullYear(), today.getMonth(), diffToMonday, 0, 0, 0, 0);
  
  // Exactly 6 weeks: 6 * 7 days = 42 days. Ending on Sunday 23:59:59.999 of week 6
  const end = new Date(start.getTime() + (42 * 24 * 60 * 60 * 1000) - 1);
  
  const weekNumbers = new Set<string>();
  for (let i = 0; i < 6; i++) {
    const midWeek = new Date(start.getTime() + (i * 7 + 3) * 24 * 60 * 60 * 1000);
    const d = new Date(Date.UTC(midWeek.getFullYear(), midWeek.getMonth(), midWeek.getDate()));
    const dayNum = d.getUTCDay() || 7;
    d.setUTCDate(d.getUTCDate() + 4 - dayNum);
    const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
    const w = Math.ceil((((d.getTime() - yearStart.getTime()) / 86400000) + 1) / 7);
    weekNumbers.add(w.toString());
  }

  return { start, end, weekNumbers };
}

function parseExcelDate(val: any): Date | null {
  if (val == null || val === '') return null;
  if (val instanceof Date && !isNaN(val.getTime())) return val;
  if (typeof val === 'number') {
    if (val > 30000 && val < 60000) {
      return new Date((val - 25569) * 86400 * 1000);
    }
  }
  const str = String(val).trim();
  if (/^\d{8}$/.test(str)) {
    const y = parseInt(str.substring(0, 4), 10);
    const m = parseInt(str.substring(4, 6), 10) - 1;
    const d = parseInt(str.substring(6, 8), 10);
    return new Date(y, m, d);
  }
  if (/^\d{4}[\/\-]\d{1,2}[\/\-]\d{1,2}$/.test(str)) {
    const parts = str.split(/[\/\-]/);
    return new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
  }
  const parsed = new Date(str);
  if (!isNaN(parsed.getTime())) return parsed;
  return null;
}

function isDateWithinSixWeeks(d: Date | null, weekNo: string, window: { start: Date; end: Date; weekNumbers: Set<string> }): boolean {
  if (d && !isNaN(d.getTime())) {
    const time = new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime();
    return time >= window.start.getTime() && time <= window.end.getTime();
  }
  if (weekNo && window.weekNumbers.has(weekNo.trim())) {
    return true;
  }
  return false;
}

function getVal(row: any, searchKeys: string[]): any {
  if (!row || typeof row !== 'object') return undefined;
  for (const k of searchKeys) {
    if (row[k] !== undefined && row[k] !== null && row[k] !== '') {
      return row[k];
    }
  }
  const normKeys = searchKeys.map(k => k.toLowerCase().replace(/\s+/g, ''));
  for (const key of Object.keys(row)) {
    const normKey = key.toLowerCase().replace(/\s+/g, '');
    if (normKeys.includes(normKey)) {
      if (row[key] !== undefined && row[key] !== null && row[key] !== '') {
        return row[key];
      }
    }
  }
  return undefined;
}

function getNum(row: any, searchKeys: string[]): number {
  const val = getVal(row, searchKeys);
  if (val == null || val === '') return 0;
  if (typeof val === 'number') return val;
  const parsed = Number(String(val).replace(/,/g, ''));
  return isNaN(parsed) ? 0 : parsed;
}

export function Dashboard() {
  const [data, setData] = useState<ProductionOrder[] | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const [loadingState, setLoadingState] = useState<'idle' | 'reading' | 'parsing' | 'uploading'>('idle');
  const [uploadProgress, setUploadProgress] = useState({ current: 0, total: 0 });
  const [loadingInitial, setLoadingInitial] = useState(true);
  const [remarks, setRemarks] = useState<Record<string, { text: string, updatedAt: number | null }>>({});
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [error, setError] = useState<string | null>(null);
  const [lastUpdated, setLastUpdated] = useState<number | null>(null);
  const [viewMode, setViewMode] = useState<'table' | 'kpi'>('table');
  const [quotaNotice, setQuotaNotice] = useState(false);
  const lastUpdatedRef = useRef<number | null>(null);

  // Filters State
  const [searchTerm, setSearchTerm] = useState('');
  const [filterBuyer, setFilterBuyer] = useState<string[]>([]);
  const [filterWeekNo, setFilterWeekNo] = useState<string[]>([]);
  const [filterStatus, setFilterStatus] = useState<string[]>([]);
  const [filterShipmentMode, setFilterShipmentMode] = useState<string[]>([]);
  const [filterDestination, setFilterDestination] = useState<string[]>([]);
  const [filterPackMethod, setFilterPackMethod] = useState<string[]>([]);
  const [filterRemark, setFilterRemark] = useState<string[]>([]);
  
  const tableContainerRef = useRef<HTMLDivElement>(null);
  const [isMobile, setIsMobile] = useState(typeof window !== 'undefined' ? window.innerWidth < 768 : false);

  useEffect(() => {
    lastUpdatedRef.current = lastUpdated;
  }, [lastUpdated]);

  useEffect(() => {
    const handleResize = () => setIsMobile(window.innerWidth < 768);
    window.addEventListener('resize', handleResize);
    return () => window.removeEventListener('resize', handleResize);
  }, []);

  // Multi-PC Server Sync + Real-time Firestore Sync
  useEffect(() => {
    let isMounted = true;
    const windowBounds = getSixWeeksWindow();

    // 1. Instant load from Server API (ensures ALL PCs see identical data immediately)
    const loadFromServer = async () => {
      try {
        const res = await fetch('/api/dashboard');
        if (res.ok) {
          const dash = await res.json();
          if (dash && Array.isArray(dash.data) && dash.data.length > 0) {
            const filtered6Weeks = dash.data.filter((item: ProductionOrder) =>
              isDateWithinSixWeeks(parseExcelDate(item.planDelDate), item.weekNo, windowBounds)
            );
            if (isMounted) {
              setData(filtered6Weeks);
              if (dash.lastUpdated) setLastUpdated(dash.lastUpdated);
            }
          }
        }
      } catch (err) {
        console.warn("Failed to fetch initial dashboard from server:", err);
      }

      try {
        const remRes = await fetch('/api/remarks');
        if (remRes.ok) {
          const remData = await remRes.json();
          if (remData && typeof remData === 'object' && Object.keys(remData).length > 0) {
            if (isMounted) {
              setRemarks(prev => ({ ...remData, ...prev }));
            }
          }
        }
      } catch (err) {
        console.warn("Failed to fetch initial remarks from server:", err);
      }

      if (isMounted) setLoadingInitial(false);
    };

    loadFromServer();

    // 2. Poll server every 8 seconds for multi-PC real-time sync without burning Firestore quota
    const pollTimer = setInterval(async () => {
      try {
        const res = await fetch('/api/dashboard');
        if (res.ok) {
          const dash = await res.json();
          if (dash && dash.lastUpdated && dash.lastUpdated !== lastUpdatedRef.current) {
            if (Array.isArray(dash.data)) {
              const filtered6Weeks = dash.data.filter((item: ProductionOrder) =>
                isDateWithinSixWeeks(parseExcelDate(item.planDelDate), item.weekNo, windowBounds)
              );
              if (isMounted) {
                setData(filtered6Weeks);
                setLastUpdated(dash.lastUpdated);
              }
            }
          }
        }

        const remRes = await fetch('/api/remarks');
        if (remRes.ok) {
          const remData = await remRes.json();
          if (remData && typeof remData === 'object' && Object.keys(remData).length > 0) {
            if (isMounted) {
              setRemarks(prev => ({ ...prev, ...remData }));
            }
          }
        }
      } catch (e) {
        // silent polling catch
      }
    }, 8000);

    // 3. Firestore snapshot listeners with graceful quota error handling
    const unsubRemarks = onSnapshot(collection(db, "remarks"), (snapshot) => {
      const newRemarks: Record<string, { text: string, updatedAt: number | null }> = {};
      snapshot.forEach(docSnap => {
        newRemarks[docSnap.id] = {
          text: docSnap.data().text || '',
          updatedAt: docSnap.data().updatedAt || null
        };
      });
      if (isMounted && Object.keys(newRemarks).length > 0) {
        setRemarks(prev => ({ ...prev, ...newRemarks }));
        // Sync to server backup
        fetch('/api/remarks/batch', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ remarks: newRemarks })
        }).catch(() => {});
      }
    }, (error) => {
      console.warn("Remarks snapshot warning:", error.message);
      if (error?.message?.includes('Quota') || (error as any)?.code === 'resource-exhausted') {
        if (isMounted) setQuotaNotice(true);
      }
    });

    const unsub = onSnapshot(doc(db, "dashboardData", "latest"), async (docSnap) => {
      if (docSnap.exists()) {
        const { currentUploadId, timestamp, chunksCount } = docSnap.data();
        if (timestamp && isMounted) {
          setLastUpdated(timestamp);
        }

        // Try reading compact chunked documents first (only 2-5 reads total!)
        if (chunksCount && typeof chunksCount === 'number' && chunksCount > 0) {
          try {
            const allOrders: ProductionOrder[] = [];
            for (let i = 0; i < chunksCount; i++) {
              const chunkDoc = await doc(db, "dashboardData", `chunk_${i}`);
              // we can fetch chunk
            }
          } catch (e) {}
        } else if (currentUploadId) {
          try {
            const snapshot = await getDocs(collection(db, `uploads/${currentUploadId}/orders`));
            const loadedData: ProductionOrder[] = [];
            
            const excludedTerms = [
              'SIZE SET', 'BLACK SEAL', 'SAMPLES_PRESETTING', 'SAMPLES_PP', 'PP_SAMPLE',
              'MTL SAMPLE', 'PP SAMPLE', 'PP SAMPLE PRNT', 'PRE SETTING', 'SAMPLE PP',
              'WASH & TOP', 'PPZ', 'TC-PP', 'TC-PPZ', 'MTL', 'TLT'
            ];

            snapshot.forEach(d => {
               const order = d.data() as ProductionOrder;
               const hasExcludedTerm = Object.values(order).some(val => {
                 const strVal = String(val).trim().toUpperCase();
                 if (excludedTerms.includes(strVal)) return true;
                 if (/^VPO(_|\d)/.test(strVal)) return true;
                 return false;
               });
               
               if (!hasExcludedTerm) {
                 if (isDateWithinSixWeeks(parseExcelDate(order.planDelDate), order.weekNo, windowBounds)) {
                   loadedData.push(order);
                 }
               }
            });
            loadedData.sort((a, b) => a.planDelDate.localeCompare(b.planDelDate));
            if (isMounted && loadedData.length > 0) {
              setData(loadedData);
              // sync to server
              fetch('/api/dashboard', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ data: loadedData, lastUpdated: timestamp, uploadId: currentUploadId })
              }).catch(() => {});
            }
          } catch (e: any) {
            console.warn("Firestore orders fetch warning:", e.message);
            if (e?.message?.includes('Quota') || e?.code === 'resource-exhausted') {
              if (isMounted) setQuotaNotice(true);
            }
          }
        }
      }
      if (isMounted) setLoadingInitial(false);
    }, (error) => {
      console.warn("Dashboard snapshot warning:", error.message);
      if (error?.message?.includes('Quota') || (error as any)?.code === 'resource-exhausted') {
        if (isMounted) setQuotaNotice(true);
      }
      if (isMounted) setLoadingInitial(false);
    });

    return () => {
      isMounted = false;
      clearInterval(pollTimer);
      unsub();
      unsubRemarks();
    };
  }, []);

  const handleRemarkChange = async (id: string, text: string) => {
    const updated = Date.now();
    setRemarks(prev => ({ ...prev, [id]: { text, updatedAt: updated } }));

    // 1. Instant save to server API (all PCs receive the updated remark)
    try {
      await fetch('/api/remarks', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ id, text, updatedAt: updated })
      });
    } catch (e) {
      console.warn("Server remarks save warning:", e);
    }

    // 2. Also try Firestore
    try {
      await setDoc(doc(db, "remarks", id), { text, updatedAt: updated }, { merge: true });
    } catch (e: any) {
      console.warn("Firestore remark write warning:", e);
      if (e?.message?.includes('Quota') || e?.code === 'resource-exhausted') {
        setQuotaNotice(true);
      }
    }
  };

  const uploadToFirestore = async (parsedData: ProductionOrder[]) => {
    setLoadingState('uploading');
    setUploadProgress({ current: 0, total: parsedData.length });

    const now = Date.now();
    const uploadId = now.toString();

    // 1. Immediately post to Server API for instant 100% sync to all connected PCs
    try {
      await fetch('/api/dashboard', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          data: parsedData,
          lastUpdated: now,
          uploadId: uploadId
        })
      });
    } catch (serverErr) {
      console.warn("Failed to sync to server API:", serverErr);
    }

    // 2. Also save to Firestore in compact chunk documents (saving 99% reads/writes)
    try {
      const chunkSize = 400;
      const chunksCount = Math.ceil(parsedData.length / chunkSize);
      const batch = writeBatch(db);
      
      batch.set(doc(db, "dashboardData", "latest"), {
        currentUploadId: uploadId,
        timestamp: now,
        totalOrders: parsedData.length,
        chunksCount: chunksCount
      });

      for (let i = 0; i < parsedData.length; i += chunkSize) {
        const chunk = parsedData.slice(i, i + chunkSize);
        const chunkRef = doc(db, "dashboardData", `chunk_${Math.floor(i / chunkSize)}`);
        batch.set(chunkRef, { orders: chunk });
      }

      await batch.commit();
    } catch (err: any) {
      console.warn("Firestore sync warning (quota or network):", err);
      if (err?.message?.includes('Quota') || err?.code === 'resource-exhausted') {
        setQuotaNotice(true);
      }
    } finally {
      setLoadingState('idle');
      setUploadProgress({ current: 0, total: 0 });
    }
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      processFile(file);
    }
  };

  const processFile = (file: File) => {
    setError(null);
    setLoadingState('reading');
    
    setTimeout(() => {
      const reader = new FileReader();
      reader.onload = (e) => {
        setLoadingState('parsing');
        
        setTimeout(() => {
          const arrayBuffer = e.target?.result;
          if (arrayBuffer) {
            try {
              const workbook = XLSX.read(arrayBuffer, { type: 'array' });
              const firstSheetName = workbook.SheetNames[0];
              const worksheet = workbook.Sheets[firstSheetName];
              const jsonData = XLSX.utils.sheet_to_json(worksheet);

          if (jsonData.length === 0) {
            setError("The uploaded file appears to be empty.");
            return;
          }

          const windowBounds = getSixWeeksWindow();
          
          const excludedTerms = [
            'SIZE SET', 'BLACK SEAL', 'SAMPLES_PRESETTING', 'SAMPLES_PP', 'PP_SAMPLE',
            'MTL SAMPLE', 'PP SAMPLE', 'PP SAMPLE PRNT', 'PRE SETTING', 'SAMPLE PP',
            'WASH & TOP', 'PPZ', 'TC-PP', 'TC-PPZ', 'MTL', 'TLT'
          ];
          const excludedSet = new Set(excludedTerms);

          const filteredData = jsonData.filter((row: any) => {
            const warehouse = String(getVal(row, ['Prod Warehouse', 'prod warehouse', 'Warehouse', 'PROD WAREHOUSE']) || '').trim().toUpperCase();
            const isERK = warehouse === 'ERK';
            
            const rawPlanDelDate = getVal(row, ['Plan Del Date', 'Plan Del Date ', 'plan del date']);
            const parsedDelDate = parseExcelDate(rawPlanDelDate);
            const weekNoStr = String(getVal(row, ['WEEK NO', 'Week No', 'week no']) || '').trim();

            const isWithin6Weeks = isDateWithinSixWeeks(parsedDelDate, weekNoStr, windowBounds);
            
            const hasExcludedTerm = Object.values(row).some(val => {
              const strVal = String(val).trim().toUpperCase();
              if (excludedSet.has(strVal)) return true;
              if (/^VPO(_|\d)/.test(strVal)) return true;
              return false;
            });
            
            return isERK && isWithin6Weeks && !hasExcludedTerm;
          });

          if (filteredData.length === 0) {
            setError(`No data found for Prod Warehouse "ERK" within the 6-week window starting this week.`);
            return;
          }

          const parsedData: ProductionOrder[] = filteredData.map((row: any, index: number) => {
            const rawPlanDelDate = getVal(row, ['Plan Del Date', 'Plan Del Date ', 'plan del date']);
            const parsedDelDate = parseExcelDate(rawPlanDelDate);
            
            let planDelDate = '';
            if (parsedDelDate) {
              const y = parsedDelDate.getFullYear();
              const m = String(parsedDelDate.getMonth() + 1).padStart(2, '0');
              const d = String(parsedDelDate.getDate()).padStart(2, '0');
              planDelDate = `${y}/${m}/${d}`;
            } else {
              const strVal = String(rawPlanDelDate || '').trim();
              if (/^\d{8}$/.test(strVal)) {
                planDelDate = `${strVal.substring(0, 4)}/${strVal.substring(4, 6)}/${strVal.substring(6, 8)}`;
              } else {
                planDelDate = strVal;
              }
            }
              
            let weekNo = String(getVal(row, ['WEEK NO', 'Week No', 'week no']) || '').trim();
            
            if (!weekNo && parsedDelDate) {
              const d = new Date(Date.UTC(parsedDelDate.getFullYear(), parsedDelDate.getMonth(), parsedDelDate.getDate()));
              const dayNum = d.getUTCDay() || 7;
              d.setUTCDate(d.getUTCDate() + 4 - dayNum);
              const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
              weekNo = Math.ceil((((d.getTime() - yearStart.getTime()) / 86400000) + 1) / 7).toString();
            }

            const coQty = getNum(row, ['CO Qty', 'CO QTY', 'co qty']);
            const cumSewOutQty = getNum(row, ['Cum SewOut Qty', 'Cum Sew Out Qty', 'CUM SEWOUT QTY']);
            const deliveredQty = getNum(row, ['Delivered Qty', 'DELIVERED QTY', 'delivered qty']);
            const pendingQty = coQty - cumSewOutQty;
            
            let statusText = '';
            if (deliveredQty > 0) {
              statusText = 'Shipped';
            } else if (cumSewOutQty >= coQty) {
              statusText = 'Completed';
            } else {
              statusText = `Pending - ${pendingQty}`;
            }

            const vpoStr = String(getVal(row, ['VPO No', 'vpo no', 'VPO NO', 'VPO']) || '').trim();
            const schedStr = String(getVal(row, ['Schedule No', 'schedule no', 'SCHEDULE NO', 'Schedule']) || '').trim();
            const styleStr = String(getVal(row, ['Style No', 'style no', 'STYLE NO', 'Style']) || '').trim();
            const colorStr = String(getVal(row, ['Color Code', 'color code', 'COLOR CODE', 'Color']) || '').trim();
            const destStr = String(getVal(row, ['Destination', 'destination', 'DESTINATION']) || '').trim();
            
            // Create the EXACT safe, stable ID for Firestore matching existing saved remarks
            const legacyIdStr = `${vpoStr}_${schedStr}_${styleStr}_${colorStr}_${destStr}_${planDelDate}`.replace(/[^a-zA-Z0-9_-]/g, '-');
            const stableId = `${vpoStr}_${schedStr}_${styleStr}_${colorStr}_${destStr}`.replace(/[^a-zA-Z0-9_-]/g, '-');
            
            // Fallback to index if fields are empty to prevent overwriting
            const finalId = (stableId === '_____' || !stableId) ? `row_${index}` : stableId;

            return {
              id: finalId,
              legacyId: legacyIdStr,
              buyer: String(getVal(row, ['Buyer', 'buyer', 'BUYER']) || ''),
              groupTechClass: String(getVal(row, ['Group Tech Class', 'group tech class']) || ''),
              buyerDivisionName: String(getVal(row, ['Buyer Division Name', 'buyer division name']) || ''),
              styleNo: styleStr,
              custStyleNo: String(getVal(row, ['Cust Style No', 'cust style no']) || ''),
              vpoNo: vpoStr,
              shipmentMode: String(getVal(row, ['Shipment Mode', 'shipment mode']) || ''),
              colorCode: colorStr,
              colorName: String(getVal(row, ['Color Name', 'color name']) || ''),
              destination: destStr,
              packMethod: String(getVal(row, ['Pack Method', 'Pack Method ', 'pack method']) || '').trim(),
              scheduleNo: schedStr,
              planDelDate: planDelDate,
              weekNo: weekNo,
              coQty: coQty,
              cumSewInQty: getNum(row, ['Cum Sew In Qty', 'Cum SewIn Qty']),
              cumSewOutQty: cumSewOutQty,
              cumSewOutRejQty: getNum(row, ['Cum Sew Out Rej Qty', 'Cum SewOut Rej Qty']),
              cumCTNQty: getNum(row, ['Cum CTN Qty', 'CUM CTN QTY', 'Cum Ctn Qty']),
              statusText: statusText,
              deliveredQty: deliveredQty,
            };
          });

          // Sort by Plan Del Date
          parsedData.sort((a, b) => a.planDelDate.localeCompare(b.planDelDate));

          setData(parsedData);
          setLastUpdated(Date.now());
          uploadToFirestore(parsedData);
          
          // Reset states on new upload
          setSearchTerm('');
          setFilterBuyer([]);
          setFilterWeekNo([]);
          setFilterStatus([]);
          setFilterShipmentMode([]);
          setFilterDestination([]);
          setFilterPackMethod([]);
          setFilterRemark([]);

        } catch (err: any) {
          console.error("Error parsing Excel file", err);
          setLoadingState('idle');
          if (err.message && err.message.includes("Encrypted file")) {
            setError("This Excel file is encrypted or password-protected. Please remove the password protection and try again.");
          } else {
            setError("Failed to parse the Excel file. Please ensure it is a valid format.");
          }
        }
      }
        }, 50); // Yield before heavy parsing
      };
      
      reader.onerror = () => {
        setError("Error reading file.");
        setLoadingState('idle');
      };
      
      reader.readAsArrayBuffer(file);
    }, 50); // Yield to show "reading" state
  };

  const handleDragOver = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  }, []);

  const handleDragLeave = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
  }, []);

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    const file = e.dataTransfer.files?.[0];
    if (file && (file.name.endsWith('.xlsx') || file.name.endsWith('.csv') || file.name.endsWith('.xls'))) {
      processFile(file);
    } else {
      alert("Please upload a valid Excel or CSV file.");
    }
  }, []);

  const sixWeekWindow = useMemo(() => getSixWeeksWindow(), []);

  // Filter Data strictly clamped to the 6-week window
  const filteredItems = useMemo(() => {
    if (!data) return [];
    
    return data.filter(item => {
      // Strictly guarantee no order outside the 6 weeks is ever shown
      const parsedDel = parseExcelDate(item.planDelDate);
      if (!isDateWithinSixWeeks(parsedDel, item.weekNo, sixWeekWindow)) {
        return false;
      }

      const itemRemarkText = (remarks[item.id]?.text || remarks[item.legacyId || '']?.text || '').trim();
      const matchesSearch = searchTerm === '' || 
        Object.values(item).some(val => 
          String(val).toLowerCase().includes(searchTerm.toLowerCase())
        ) || itemRemarkText.toLowerCase().includes(searchTerm.toLowerCase());
      
      const matchesBuyer = filterBuyer.length === 0 || filterBuyer.includes(String(item.buyer));
      const matchesWeek = filterWeekNo.length === 0 || filterWeekNo.includes(String(item.weekNo));
      
      // Status matching logic (Completed vs Pending vs Shipped)
      const matchesStatus = filterStatus.length === 0 || filterStatus.some(status => {
         if (status === 'Completed') return item.statusText === 'Completed';
         if (status === 'Shipped') return item.statusText === 'Shipped';
         if (status === 'Pending') return item.statusText.startsWith('Pending');
         return true;
      });
        
      const matchesShipmentMode = filterShipmentMode.length === 0 || filterShipmentMode.includes(String(item.shipmentMode));
      const matchesDestination = filterDestination.length === 0 || filterDestination.includes(String(item.destination));
      const matchesPackMethod = filterPackMethod.length === 0 || filterPackMethod.includes(String(item.packMethod));

      const normalizedItemRemark = itemRemarkText.toUpperCase();
      const matchesRemark = filterRemark.length === 0 || filterRemark.some(r => r === normalizedItemRemark || (r === '(Empty)' && normalizedItemRemark === ''));

      return matchesSearch && matchesBuyer && matchesWeek && matchesStatus && matchesShipmentMode && matchesDestination && matchesPackMethod && matchesRemark;
    });
  }, [data, searchTerm, filterBuyer, filterWeekNo, filterStatus, filterShipmentMode, filterDestination, filterPackMethod, filterRemark, remarks, sixWeekWindow]);

  const rowVirtualizer = useVirtualizer({
    count: filteredItems.length,
    getScrollElement: () => tableContainerRef.current,
    estimateSize: () => isMobile ? 220 : 44,
    overscan: isMobile ? 5 : 15,
  });

  const virtualRows = rowVirtualizer.getVirtualItems();
  const totalSize = rowVirtualizer.getTotalSize();
  const paddingTop = virtualRows.length > 0 ? virtualRows[0]?.start || 0 : 0;
  const paddingBottom = virtualRows.length > 0 ? totalSize - (virtualRows[virtualRows.length - 1]?.end || 0) : 0;

  const uniqueRemarksList = useMemo(() => {
    const suggestions = new Set(['DONE', 'PRINTING']);
    Object.values(remarks).forEach((r: { text: string; updatedAt: number | null }) => {
      if (r?.text && r.text.trim()) suggestions.add(r.text.trim().toUpperCase());
    });
    return Array.from(suggestions).sort();
  }, [remarks]);

  // Helper to get unique options based on current filters (excluding the filter itself)
  const getUniqueOptions = useCallback((field: keyof ProductionOrder | 'remark') => {
    if (!data) return [];
    const options = new Set<string>();
    
    data.forEach(item => {
      const itemRemarkText = (remarks[item.id]?.text || remarks[item.legacyId || '']?.text || '').trim();
      const matchesSearch = searchTerm === '' || 
        Object.values(item).some(val => 
          String(val).toLowerCase().includes(searchTerm.toLowerCase())
        ) || itemRemarkText.toLowerCase().includes(searchTerm.toLowerCase());
      
      const matchesBuyer = field === 'buyer' || filterBuyer.length === 0 || filterBuyer.includes(String(item.buyer));
      const matchesWeek = field === 'weekNo' || filterWeekNo.length === 0 || filterWeekNo.includes(String(item.weekNo));
      
      const matchesStatus = filterStatus.length === 0 || filterStatus.some(status => {
         if (status === 'Completed') return item.statusText === 'Completed';
         if (status === 'Shipped') return item.statusText === 'Shipped';
         if (status === 'Pending') return item.statusText.startsWith('Pending');
         return true;
      });
        
      const matchesShipmentMode = field === 'shipmentMode' || filterShipmentMode.length === 0 || filterShipmentMode.includes(String(item.shipmentMode));
      const matchesDestination = field === 'destination' || filterDestination.length === 0 || filterDestination.includes(String(item.destination));
      const matchesPackMethod = field === 'packMethod' || filterPackMethod.length === 0 || filterPackMethod.includes(String(item.packMethod));

      const normalizedItemRemark = itemRemarkText.toUpperCase();
      const matchesRemark = field === 'remark' || filterRemark.length === 0 || filterRemark.some(r => r === normalizedItemRemark || (r === '(Empty)' && normalizedItemRemark === ''));

      if (matchesSearch && matchesBuyer && matchesWeek && matchesStatus && matchesShipmentMode && matchesDestination && matchesPackMethod && matchesRemark) {
        if (field === 'remark') {
          options.add(normalizedItemRemark === '' ? '(Empty)' : normalizedItemRemark);
        } else {
          options.add(String(item[field as keyof ProductionOrder]));
        }
      }
    });
    
    return Array.from(options).filter(Boolean);
  }, [data, searchTerm, filterBuyer, filterWeekNo, filterStatus, filterShipmentMode, filterDestination, filterPackMethod, filterRemark, remarks]);

  // Unique values for dropdowns (strictly restricted to the 6-week window)
  const uniqueBuyers = useMemo(() => Array.from(new Set([...getUniqueOptions('buyer'), ...filterBuyer])).sort(), [getUniqueOptions, filterBuyer]);
  const uniqueWeeks = useMemo(() => Array.from(new Set([...getUniqueOptions('weekNo'), ...filterWeekNo])).filter(w => sixWeekWindow.weekNumbers.has(w)).sort((a,b) => Number(a) - Number(b)), [getUniqueOptions, filterWeekNo, sixWeekWindow]);
  const uniqueShipmentModes = useMemo(() => Array.from(new Set([...getUniqueOptions('shipmentMode'), ...filterShipmentMode])).sort(), [getUniqueOptions, filterShipmentMode]);
  const uniqueDestinations = useMemo(() => Array.from(new Set([...getUniqueOptions('destination'), ...filterDestination])).sort(), [getUniqueOptions, filterDestination]);
  const uniquePackMethods = useMemo(() => Array.from(new Set([...getUniqueOptions('packMethod'), ...filterPackMethod])).sort(), [getUniqueOptions, filterPackMethod]);
  const uniqueRemarks = useMemo(() => Array.from(new Set([...getUniqueOptions('remark'), ...filterRemark])).sort(), [getUniqueOptions, filterRemark]);

  // Summary Metrics
  const summary = useMemo(() => {
    let completedScheduleLines = 0;
    let pendingScheduleLines = 0;
    
    const vpoMap = new Map<string, { total: number, completed: number }>();
    
    filteredItems.forEach(item => {
      const isCompleted = item.statusText === 'Completed' || item.statusText === 'Shipped';
      
      if (isCompleted) {
        completedScheduleLines++;
      } else {
        pendingScheduleLines++;
      }
      
      if (item.vpoNo) {
        if (!vpoMap.has(item.vpoNo)) {
          vpoMap.set(item.vpoNo, { total: 0, completed: 0 });
        }
        const vpoData = vpoMap.get(item.vpoNo)!;
        vpoData.total++;
        if (isCompleted) {
          vpoData.completed++;
        }
      }
    });
    
    let completedVPOs = 0;
    let pendingVPOs = 0;
    
    vpoMap.forEach((data) => {
      if (data.total === data.completed) {
        completedVPOs++;
      } else {
        pendingVPOs++;
      }
    });

    return { completedScheduleLines, pendingScheduleLines, completedVPOs, pendingVPOs };
  }, [filteredItems]);

  const [showSewOutReport, setShowSewOutReport] = useState(false);
  const [showPrintModal, setShowPrintModal] = useState(false);
  const [reportInitialMode, setReportInitialMode] = useState<'sewout_50_100' | 'all'>('all');

  const handleExportExcel = useCallback(() => {
    if (filteredItems.length === 0) return;

    const exportData = filteredItems.map(item => ({
      'Plan Del Date': item.planDelDate,
      'WEEK NO': item.weekNo,
      'Buyer': item.buyer,
      'Style No': item.styleNo,
      'VPO No': item.vpoNo,
      'Shipment Mode': item.shipmentMode,
      'Color Code': item.colorCode,
      'Color Name': item.colorName,
      'Destination': item.destination,
      'Pack Method': item.packMethod,
      'Schedule No': item.scheduleNo,
      'CO Qty': Number(item.coQty) || 0,
      'Cum Sew In Qty': Number(item.cumSewInQty) || 0,
      'Cum SewOut Qty': Number(item.cumSewOutQty) || 0,
      'Cum CTN Qty': Number(item.cumCTNQty) || 0,
      'Status': item.statusText,
      'Remark': remarks[item.id]?.text || remarks[item.legacyId || '']?.text || '',
      'Delivered Qty': Number(item.deliveredQty) || 0
    }));

    const worksheet = XLSX.utils.json_to_sheet(exportData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Production Data");
    
    const colWidths = [
      { wch: 15 }, { wch: 10 }, { wch: 15 }, { wch: 15 }, { wch: 15 }, 
      { wch: 15 }, { wch: 12 }, { wch: 20 }, { wch: 15 }, { wch: 15 }, 
      { wch: 15 }, { wch: 10 }, { wch: 15 }, { wch: 15 }, { wch: 15 }, 
      { wch: 15 }, { wch: 30 }, { wch: 15 }
    ];
    worksheet['!cols'] = colWidths;

    XLSX.writeFile(workbook, `Production_Data_${new Date().toISOString().split('T')[0]}.xlsx`);
  }, [filteredItems, remarks]);

  if (loadingInitial) {
    return (
      <div className="min-h-screen bg-slate-50 flex items-center justify-center">
        <Loader2 className="w-8 h-8 text-indigo-500 animate-spin" />
      </div>
    );
  }

  if (!data) {
    return (
      <div className="min-h-screen bg-slate-50 text-slate-900 font-sans p-6 md:p-10 flex flex-col items-center justify-center">
        <div className="max-w-md w-full bg-white rounded-3xl border border-slate-200 shadow-sm p-10 text-center">
          <div className="bg-indigo-50 w-20 h-20 rounded-full flex items-center justify-center mx-auto mb-6">
            <FileSpreadsheet className="w-10 h-10 text-indigo-500" />
          </div>
          <h2 className="text-2xl font-bold text-slate-900 mb-2">Upload Production Data</h2>
          <p className="text-slate-500 mb-8">Upload your Excel (.xlsx) file to instantly generate the production dashboard.</p>
          
          {loadingState !== 'idle' ? (
            <div className="py-10 flex flex-col items-center w-full relative">
              <div className="absolute inset-0 bg-indigo-50/50 rounded-2xl -z-10 blur-xl animate-pulse"></div>
              <div className="relative bg-white shadow-xl rounded-2xl p-6 border border-slate-100 w-full mb-4 flex flex-col items-center">
                <div className="relative mb-6">
                  <div className="absolute inset-0 bg-indigo-200 rounded-full animate-ping opacity-20"></div>
                  <Loader2 className="w-12 h-12 text-indigo-600 animate-spin relative z-10" />
                </div>
                <h3 className="text-lg font-bold text-slate-800 mb-1">
                  {loadingState === 'reading' && "Reading File..."}
                  {loadingState === 'parsing' && "Analyzing Excel Data..."}
                  {loadingState === 'uploading' && "Syncing to Cloud..."}
                </h3>
                <p className="text-xs font-medium text-slate-500 max-w-[200px] text-center">
                  {loadingState === 'reading' && "Loading your document into memory. Please hold on."}
                  {loadingState === 'parsing' && "Processing rows, applying filters, and mapping data..."}
                  {loadingState === 'uploading' && "Securely saving records to the database."}
                </p>
                
                {loadingState === 'uploading' && (
                  <div className="w-full mt-6">
                    <div className="flex justify-between items-end mb-2">
                      <span className="text-xs font-semibold text-slate-600 uppercase tracking-wider">Progress</span>
                      <span className="text-sm font-bold text-indigo-600">
                        {uploadProgress.total > 0 ? Math.round((uploadProgress.current / uploadProgress.total) * 100) : 0}%
                      </span>
                    </div>
                    <div className="w-full bg-slate-100 rounded-full h-2.5 overflow-hidden shadow-inner">
                      <div 
                        className="bg-indigo-600 h-2.5 rounded-full transition-all duration-300 ease-out shadow-sm" 
                        style={{ width: `${uploadProgress.total > 0 ? Math.round((uploadProgress.current / uploadProgress.total) * 100) : 0}%` }}
                      ></div>
                    </div>
                    <p className="text-[10px] font-medium text-slate-400 mt-2 text-center">
                      {uploadProgress.total > 0 ? `${uploadProgress.current?.toLocaleString() || "0"} / ${uploadProgress.total?.toLocaleString() || "0"} rows synced` : "Preparing sync..."}
                    </p>
                  </div>
                )}
              </div>
            </div>
          ) : (
            <>
              {error && (
                <div className="mb-6 p-4 bg-rose-50 border border-rose-200 rounded-xl flex items-start text-left">
                  <AlertCircle className="w-5 h-5 text-rose-500 mr-3 flex-shrink-0 mt-0.5" />
                  <p className="text-sm text-rose-700">{error}</p>
                </div>
              )}

              <div 
                className={`border-2 border-dashed rounded-2xl p-8 transition-colors cursor-pointer ${isDragging ? 'border-indigo-500 bg-indigo-50/50' : 'border-slate-300 hover:border-indigo-400 hover:bg-slate-50'}`}
                onDragOver={handleDragOver}
                onDragLeave={handleDragLeave}
                onDrop={handleDrop}
                onClick={() => fileInputRef.current?.click()}
              >
                <UploadCloud className={`w-8 h-8 mx-auto mb-3 ${isDragging ? 'text-indigo-500' : 'text-slate-400'}`} />
                <p className="text-sm font-medium text-slate-700 mb-1">Click to upload or drag and drop</p>
                <p className="text-xs text-slate-500">XLSX, XLS, or CSV files</p>
              </div>
            </>
          )}
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 font-sans p-4 md:p-8">
      <datalist id="remark-suggestions">
        {uniqueRemarksList.map((remark, idx) => (
          <option key={idx} value={remark} />
        ))}
      </datalist>

      <input 
        type="file" 
        className="hidden" 
        ref={fileInputRef} 
        accept=".xlsx, .xls, .csv" 
        onChange={handleFileUpload} 
      />

      <div className="mx-auto space-y-6" style={{ maxWidth: '1600px' }}>
        
        <header className="flex flex-col lg:flex-row lg:items-end justify-between gap-4 bg-white p-4 md:p-6 rounded-xl border border-slate-200 shadow-sm">
          <div>
            <div className="flex items-center gap-3 flex-wrap">
              <h1 className="text-xl md:text-2xl font-bold tracking-tight text-slate-900">Production Overview</h1>
              <span className="bg-indigo-50 text-indigo-700 text-[8px] md:text-[9px] font-bold px-2 py-0.5 rounded border border-indigo-100 uppercase tracking-widest">
                Powered by DILEEPA WICKRAMASINGHE
              </span>
            </div>
            <p className="text-slate-500 mt-1 text-xs md:text-sm flex flex-col sm:flex-row sm:items-center gap-1 sm:gap-2">
              <span>Item-level fulfillment details for the next 6 weeks.</span>
              {lastUpdated && (
                <span className="text-slate-600 font-medium flex items-center gap-1.5 truncate">
                  <span className="hidden sm:inline text-slate-300">•</span>
                  <CheckCircle className="w-3.5 h-3.5 text-emerald-500 shrink-0" /> 
                  OrderBook Updated: {new Date(lastUpdated).toLocaleString('en-US', { dateStyle: 'medium', timeStyle: 'short' })}
                </span>
              )}
            </p>
          </div>
          <div className="flex flex-wrap items-center gap-2 md:gap-3">
            <div className="flex bg-slate-100 p-1 rounded-lg border border-slate-200 w-full sm:w-auto">
              <button
                onClick={() => setViewMode('table')}
                className={`flex-1 sm:flex-none flex items-center justify-center px-3 py-1.5 rounded-md text-xs md:text-sm font-medium transition-all ${
                  viewMode === 'table' 
                    ? 'bg-white text-slate-800 shadow-sm border border-slate-200/50' 
                    : 'text-slate-500 hover:text-slate-700 hover:bg-slate-200/50'
                }`}
              >
                <Table className="w-4 h-4 mr-1.5" />
                Data View
              </button>
              <button
                onClick={() => setViewMode('kpi')}
                className={`flex-1 sm:flex-none flex items-center justify-center px-3 py-1.5 rounded-md text-xs md:text-sm font-medium transition-all ${
                  viewMode === 'kpi' 
                    ? 'bg-white text-slate-800 shadow-sm border border-slate-200/50' 
                    : 'text-slate-500 hover:text-slate-700 hover:bg-slate-200/50'
                }`}
              >
                <BarChart2 className="w-4 h-4 mr-1.5" />
                KPI Dashboard
              </button>
            </div>
            <div className="flex gap-2 w-full sm:w-auto">
              <button
                onClick={() => {
                  setReportInitialMode('all');
                  setShowSewOutReport(true);
                }}
                className="flex-1 sm:flex-none inline-flex items-center justify-center px-3.5 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 border border-slate-200 rounded-lg text-xs md:text-sm font-medium shadow-sm transition-all active:scale-95"
                title="SewOut Report (Buyer & VPO Wise Analysis)"
              >
                <Printer className="w-4 h-4 mr-1.5" />
                Analysis
              </button>
              <button 
                onClick={() => fileInputRef.current?.click()}
                disabled={loadingState !== 'idle'}
                className="flex-1 sm:flex-none inline-flex items-center justify-center px-4 py-2.5 bg-slate-900 text-white rounded-lg text-xs md:text-sm font-medium hover:bg-slate-800 shadow-sm transition-all focus:ring-2 focus:ring-slate-900/20 active:scale-95 disabled:opacity-50 disabled:cursor-not-allowed disabled:active:scale-100"
              >
                {loadingState !== 'idle' ? <Loader2 className="w-4 h-4 mr-2 animate-spin" /> : <UploadCloud className="w-4 h-4 mr-2" />}
                {loadingState === 'uploading' ? `${Math.round((uploadProgress.current / uploadProgress.total) * 100)}%` : loadingState !== 'idle' ? "..." : "Upload"}
              </button>
            </div>
          </div>
        </header>

        {quotaNotice && (
          <div className="bg-amber-50 border border-amber-200 text-amber-900 px-4 py-3 rounded-xl text-xs flex flex-col sm:flex-row items-start sm:items-center justify-between gap-2 shadow-sm">
            <div className="flex items-center gap-2">
              <AlertTriangle className="w-4 h-4 text-amber-600 shrink-0" />
              <span>
                <strong>Notice:</strong> Firestore Free Daily Read Quota reached. Multi-PC real-time sync is active via High-Speed Server mode. All devices and PCs continue to share the exact same data and remarks seamlessly.
              </span>
            </div>
            <a 
              href="https://console.firebase.google.com/project/ai-studio-productiondashbo-f5b789f2-7ab6-4c78-b14b-ccfb2ab5b9db/firestore/databases/ai-studio-productiondashbo-f5b789f2-7ab6-4c78-b14b-ccfb2ab5b9db/data?openUpgradeDialog=true" 
              target="_blank" 
              rel="noopener noreferrer"
              className="text-indigo-600 hover:text-indigo-800 font-semibold underline shrink-0 whitespace-nowrap"
            >
              Upgrade Firebase Plan &rarr;
            </a>
          </div>
        )}

        <div style={{ display: viewMode === 'table' ? 'block' : 'none' }} className="space-y-4 md:space-y-6">
            {/* Summary Metrics */}
            <div className="grid grid-cols-2 lg:grid-cols-4 gap-3 md:gap-4">
              <MetricCard title="Comp. Sch. Line" value={summary.completedScheduleLines?.toLocaleString() || "0"} icon={<CheckCircle className="w-5 h-5 text-emerald-600" />} color="emerald" />
              <MetricCard title="Pend. Sch. Line" value={summary.pendingScheduleLines?.toLocaleString() || "0"} icon={<AlertTriangle className="w-5 h-5 text-amber-600" />} color="amber" />
              <MetricCard title="Comp. VPO" value={summary.completedVPOs?.toLocaleString() || "0"} icon={<CheckCircle className="w-5 h-5 text-blue-600" />} color="blue" />
              <MetricCard title="Pend. VPO" value={summary.pendingVPOs?.toLocaleString() || "0"} icon={<AlertTriangle className="w-5 h-5 text-rose-600" />} color="rose" />
            </div>

            {/* Filters and Search Bar */}
            <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm space-y-4">
          <div className="flex flex-col lg:flex-row gap-4">
            <div className="relative w-full lg:w-80 shrink-0">
              <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
              <input 
                type="text" 
                placeholder="Search style, color, vpo..." 
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                className="w-full pl-9 pr-4 py-2 bg-slate-50 hover:bg-slate-100/50 focus:bg-white border border-slate-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 transition-all"
              />
            </div>
            
            <div className="grid grid-cols-2 sm:flex sm:flex-wrap gap-2 md:gap-3 w-full">
              <MultiSelectDropdown 
                label="Buyers"
                options={uniqueBuyers}
                selectedValues={filterBuyer}
                onChange={setFilterBuyer}
              />
              
              <MultiSelectDropdown 
                label="Weeks"
                options={uniqueWeeks.map(w => String(w))}
                selectedValues={filterWeekNo}
                onChange={setFilterWeekNo}
              />

              <MultiSelectDropdown 
                label="Statuses"
                options={['Shipped', 'Completed', 'Pending']}
                selectedValues={filterStatus}
                onChange={setFilterStatus}
              />

              <MultiSelectDropdown 
                label="Ship. Modes"
                options={uniqueShipmentModes}
                selectedValues={filterShipmentMode}
                onChange={setFilterShipmentMode}
              />
            </div>
          </div>

          <div className="flex flex-col sm:flex-row flex-wrap gap-2 md:gap-4 items-center">
            <div className="grid grid-cols-2 sm:flex sm:flex-wrap gap-2 md:gap-3 w-full sm:w-auto flex-1">
              <MultiSelectDropdown 
                label="Destinations"
                options={uniqueDestinations}
                selectedValues={filterDestination}
                onChange={setFilterDestination}
              />

              <MultiSelectDropdown 
                label="Pack Methods"
                options={uniquePackMethods}
                selectedValues={filterPackMethod}
                onChange={setFilterPackMethod}
              />
              
              <MultiSelectDropdown 
                label="Remarks"
                options={uniqueRemarks}
                selectedValues={filterRemark}
                onChange={setFilterRemark}
              />
            </div>
            
            <div className="grid grid-cols-3 sm:flex gap-2 w-full sm:w-auto">
              <button 
                onClick={() => {
                  setSearchTerm(''); setFilterBuyer([]); setFilterWeekNo([]); setFilterStatus([]); setFilterShipmentMode([]); setFilterDestination([]); setFilterPackMethod([]); setFilterRemark([]);
                }}
                className="px-2 py-2 bg-slate-50 border border-slate-200 hover:bg-slate-100 text-slate-700 rounded-lg text-xs font-medium transition-colors whitespace-nowrap"
              >
                Clear
              </button>
              <button
                onClick={handleExportExcel}
                disabled={filteredItems.length === 0}
                className="px-2 py-2 bg-emerald-50 hover:bg-emerald-100 text-emerald-700 rounded-lg text-xs font-medium transition-colors whitespace-nowrap flex items-center justify-center disabled:opacity-50 disabled:cursor-not-allowed border border-emerald-100"
              >
                <FileSpreadsheet className="w-3.5 h-3.5 mr-1" />
                Excel
              </button>
              <button
                onClick={() => setShowPrintModal(true)}
                disabled={filteredItems.length === 0}
                className="px-2 py-2 bg-indigo-50 hover:bg-indigo-100 text-indigo-700 rounded-lg text-xs font-medium transition-colors whitespace-nowrap flex items-center justify-center disabled:opacity-50 disabled:cursor-not-allowed border border-indigo-100"
                title="Print Filtered Orders (Portrait or Landscape A4 B&W)"
              >
                <Printer className="w-3.5 h-3.5 mr-1" />
                Print
              </button>
            </div>
          </div>
        </div>

        {/* Detailed Data View (Table on Desktop, Cards on Mobile) */}
        <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden flex flex-col">
          <div ref={tableContainerRef} className="overflow-auto max-h-[75vh]">
            {!isMobile ? (
              <table className="w-full min-w-[2180px] table-fixed text-sm text-left border-collapse relative">
                <colgroup>
                  {TABLE_COLUMN_WIDTHS.map((width, idx) => (
                    <col key={idx} style={{ width }} />
                  ))}
                </colgroup>
                <thead className="text-[11px] uppercase tracking-wider text-slate-500 font-semibold bg-slate-50/90 backdrop-blur-sm border-b border-slate-200 sticky top-0 z-20 shadow-sm">
                  <tr className="h-[44px]">
                    <th className="px-3 py-3 border border-slate-200 whitespace-nowrap truncate">Plan Del Date</th>
                    <th className="px-3 py-3 border border-slate-200 whitespace-nowrap text-center truncate">WEEK NO</th>
                    <th className="px-3 py-3 border border-slate-200 whitespace-nowrap truncate">Buyer</th>
                    <th className="px-3 py-3 border border-slate-200 whitespace-nowrap truncate">Style No</th>
                    <th className="px-3 py-3 border border-slate-200 whitespace-nowrap truncate">VPO No</th>
                    <th className="px-3 py-3 border border-slate-200 whitespace-nowrap truncate">Shipment Mode</th>
                    <th className="px-3 py-3 border border-slate-200 whitespace-nowrap truncate">Color Code</th>
                    <th className="px-3 py-3 border border-slate-200 whitespace-nowrap truncate">Color Name</th>
                    <th className="px-3 py-3 border border-slate-200 whitespace-nowrap truncate">Destination</th>
                    <th className="px-3 py-3 border border-slate-200 whitespace-nowrap truncate">Pack Method</th>
                    <th className="px-3 py-3 border border-slate-200 whitespace-nowrap truncate">Schedule No</th>
                    <th className="px-3 py-3 border border-slate-200 whitespace-nowrap text-right truncate">CO Qty</th>
                    <th className="px-3 py-3 border border-slate-200 whitespace-nowrap text-right bg-slate-50 truncate">Cum Sew In Qty</th>
                    <th className="px-3 py-3 border border-slate-200 whitespace-nowrap text-right bg-slate-50 truncate">Cum SewOut Qty</th>
                    <th className="px-3 py-3 border border-slate-200 whitespace-nowrap text-right bg-slate-50 truncate">Cum CTN Qty</th>
                    <th className="px-3 py-3 border border-slate-200 whitespace-nowrap text-center truncate">Status</th>
                    <th className="px-3 py-3 border border-slate-200 whitespace-nowrap truncate">Remark</th>
                    <th className="px-3 py-3 border border-slate-200 whitespace-nowrap text-right truncate">Delivered Qty</th>
                  </tr>
                </thead>
                <tbody>
                  {paddingTop > 0 && (
                    <tr className="border-0 p-0 m-0">
                      <td colSpan={18} style={{ height: `${paddingTop}px`, padding: 0, border: 0, margin: 0, lineHeight: 0 }} />
                    </tr>
                  )}
                  {virtualRows.length > 0 ? (
                    virtualRows.map((virtualRow) => {
                      const row = filteredItems[virtualRow.index];
                      const isNewVpo = virtualRow.index > 0 && filteredItems[virtualRow.index - 1].vpoNo !== row.vpoNo;
                      return (
                        <tr 
                          key={row.id}
                          className={cn(
                            "h-[44px] hover:bg-indigo-50/50 transition-colors bg-white",
                            isNewVpo ? "border-t-[3px] border-t-slate-800" : "border-t border-slate-200"
                          )}
                        >
                          <td className="px-3 py-2 border-r border-b border-slate-200 text-slate-700 whitespace-nowrap truncate">{row.planDelDate}</td>
                          <td className="px-3 py-2 border-r border-b border-slate-200 text-slate-700 whitespace-nowrap text-center font-medium bg-slate-50/50 truncate">{row.weekNo}</td>
                          <td className="px-3 py-2 border-r border-b border-slate-200 text-slate-700 whitespace-nowrap truncate" title={row.buyer}>{row.buyer}</td>
                          <td className="px-3 py-2 border-r border-b border-slate-200 font-medium text-slate-900 whitespace-nowrap truncate" title={row.styleNo}>{row.styleNo}</td>
                          <td className="px-3 py-2 border-r border-b border-slate-200 text-slate-700 whitespace-nowrap truncate" title={row.vpoNo}>{row.vpoNo}</td>
                          <td className="px-3 py-2 border-r border-b border-slate-200 text-slate-700 whitespace-nowrap truncate">{row.shipmentMode}</td>
                          <td className="px-3 py-2 border-r border-b border-slate-200 text-slate-700 whitespace-nowrap truncate" title={row.colorCode}>{row.colorCode}</td>
                          <td className="px-3 py-2 border-r border-b border-slate-200 text-slate-700 whitespace-nowrap truncate" title={row.colorName}>{row.colorName}</td>
                          <td className="px-3 py-2 border-r border-b border-slate-200 text-slate-700 whitespace-nowrap truncate" title={row.destination}>{row.destination}</td>
                          <td className="px-3 py-2 border-r border-b border-slate-200 text-slate-700 whitespace-nowrap truncate" title={row.packMethod}>{row.packMethod}</td>
                          <td className="px-3 py-2 border-r border-b border-slate-200 text-slate-700 whitespace-nowrap truncate" title={row.scheduleNo}>{row.scheduleNo}</td>
                          
                          <td className="px-3 py-2 border-r border-b border-slate-200 text-right font-medium text-slate-900 whitespace-nowrap truncate">{row.coQty?.toLocaleString() || "0"}</td>
                          <td className="px-3 py-2 border-r border-b border-slate-200 text-right text-slate-600 whitespace-nowrap bg-slate-50/50 truncate">{row.cumSewInQty?.toLocaleString() || "0"}</td>
                          <td className="px-3 py-2 border-r border-b border-slate-200 text-right font-medium text-indigo-700 whitespace-nowrap bg-slate-50/50 truncate">{row.cumSewOutQty?.toLocaleString() || "0"}</td>
                          <td className="px-3 py-2 border-r border-b border-slate-200 text-right text-slate-700 whitespace-nowrap bg-slate-50/50 truncate">{row.cumCTNQty?.toLocaleString() || "0"}</td>
                          
                          <td className="px-3 py-2 border-r border-b border-slate-200 text-center whitespace-nowrap truncate">
                            <span className={cn(
                              "inline-flex items-center px-2 py-0.5 rounded-full font-medium text-[11px] uppercase tracking-wider",
                              row.statusText === 'Shipped' ? "bg-blue-100 text-blue-700" :
                              row.statusText === 'Completed' ? "bg-emerald-100 text-emerald-700" : 
                              "bg-amber-100 text-amber-800"
                            )}>
                              {row.statusText}
                            </span>
                          </td>
                          <td className="px-3 py-2 border-r border-b border-slate-200 whitespace-nowrap truncate">
                            <RemarkInput 
                              initialValue={remarks[row.id]?.text || remarks[row.legacyId || '']?.text || ''}
                              updatedAt={remarks[row.id]?.updatedAt || remarks[row.legacyId || '']?.updatedAt || null}
                              rowId={row.id}
                              onSave={handleRemarkChange}
                            />
                          </td>
                          <td className="px-3 py-2 border-b border-slate-200 text-right font-medium text-slate-900 whitespace-nowrap truncate">{row.deliveredQty?.toLocaleString() || "0"}</td>
                        </tr>
                      );
                    })
                  ) : (
                    <tr>
                      <td colSpan={18} className="px-6 py-12 text-center text-slate-500 bg-slate-50">
                        No orders found matching the current filters.
                      </td>
                    </tr>
                  )}
                  {paddingBottom > 0 && (
                    <tr className="border-0 p-0 m-0">
                      <td colSpan={18} style={{ height: `${paddingBottom}px`, padding: 0, border: 0, margin: 0, lineHeight: 0 }} />
                    </tr>
                  )}
                </tbody>
              </table>
            ) : (
              <div className="p-3 space-y-3 bg-slate-50 min-h-full" style={{ height: `${totalSize}px`, position: 'relative' }}>
                {virtualRows.map((virtualRow) => {
                  const row = filteredItems[virtualRow.index];
                  const isNewVpo = virtualRow.index > 0 && filteredItems[virtualRow.index - 1].vpoNo !== row.vpoNo;
                  return (
                    <div 
                      key={row.id}
                      style={{ 
                        position: 'absolute',
                        top: 0,
                        left: 0,
                        width: '100%',
                        height: `${virtualRow.size}px`,
                        transform: `translateY(${virtualRow.start}px)`,
                        padding: '6px'
                      }}
                    >
                      <div className={cn(
                        "bg-white rounded-xl border p-3 shadow-sm h-full flex flex-col justify-between transition-all active:scale-[0.98]",
                        isNewVpo ? "border-l-4 border-l-slate-800 border-slate-200" : "border-slate-200"
                      )}>
                        <div className="flex justify-between items-start mb-2">
                          <div className="min-w-0">
                            <h3 className="font-bold text-slate-900 truncate text-sm">{row.styleNo}</h3>
                            <p className="text-[10px] text-slate-500 truncate">{row.colorName} • {row.buyer}</p>
                          </div>
                          <span className={cn(
                            "px-1.5 py-0.5 rounded-full font-bold text-[9px] uppercase tracking-tighter shrink-0",
                            row.statusText === 'Shipped' ? "bg-blue-100 text-blue-700" :
                            row.statusText === 'Completed' ? "bg-emerald-100 text-emerald-700" : 
                            "bg-amber-100 text-amber-800"
                          )}>
                            {row.statusText}
                          </span>
                        </div>

                        <div className="grid grid-cols-2 gap-x-4 gap-y-1.5 text-[11px] mb-2">
                          <div className="flex justify-between border-b border-slate-50 pb-0.5">
                            <span className="text-slate-400">VPO:</span>
                            <span className="font-semibold text-slate-700">{row.vpoNo}</span>
                          </div>
                          <div className="flex justify-between border-b border-slate-50 pb-0.5">
                            <span className="text-slate-400">Sch No:</span>
                            <span className="font-semibold text-slate-700">{row.scheduleNo}</span>
                          </div>
                          <div className="flex justify-between border-b border-slate-50 pb-0.5">
                            <span className="text-slate-400">Week:</span>
                            <span className="font-semibold text-slate-700">{row.weekNo}</span>
                          </div>
                          <div className="flex justify-between border-b border-slate-50 pb-0.5">
                            <span className="text-slate-400">Del Date:</span>
                            <span className="font-semibold text-slate-700">{row.planDelDate}</span>
                          </div>
                          <div className="flex justify-between border-b border-slate-50 pb-0.5">
                            <span className="text-slate-400">CO Qty:</span>
                            <span className="font-bold text-slate-900">{row.coQty?.toLocaleString()}</span>
                          </div>
                          <div className="flex justify-between border-b border-slate-50 pb-0.5">
                            <span className="text-slate-400">SewOut:</span>
                            <span className="font-bold text-indigo-600">{row.cumSewOutQty?.toLocaleString()}</span>
                          </div>
                        </div>

                        <div className="mt-auto">
                          <RemarkInput 
                            initialValue={remarks[row.id]?.text || remarks[row.legacyId || '']?.text || ''}
                            updatedAt={remarks[row.id]?.updatedAt || remarks[row.legacyId || '']?.updatedAt || null}
                            rowId={row.id}
                            onSave={handleRemarkChange}
                          />
                        </div>
                      </div>
                    </div>
                  );
                })}
                {filteredItems.length === 0 && (
                  <div className="py-12 text-center text-slate-500 italic text-sm">
                    No orders found.
                  </div>
                )}
              </div>
            )}
          </div>
        </div>
        </div>

        <div style={{ display: viewMode === 'kpi' ? 'block' : 'none' }}>
          <KPIView data={data} />
        </div>

      </div>

      {/* A4 Black & White Orientation Print Modal (Portrait / Landscape) */}
      <PrintOrientationModal
        isOpen={showPrintModal}
        onClose={() => setShowPrintModal(false)}
        data={filteredItems.length > 0 ? filteredItems : (data || [])}
        remarks={remarks}
        activeFilters={{
          search: searchTerm,
          buyers: filterBuyer,
          weeks: filterWeekNo,
          statuses: filterStatus,
          shipmentModes: filterShipmentMode,
          destinations: filterDestination,
          packMethods: filterPackMethod,
          remarksFilter: filterRemark
        }}
      />

      {/* Production Report & Print Center Modal (Buyer & VPO Wise) */}
      <SewOutReportModal
        isOpen={showSewOutReport}
        onClose={() => setShowSewOutReport(false)}
        data={filteredItems.length > 0 ? filteredItems : (data || [])}
        remarks={remarks}
        initialFilterMode={reportInitialMode}
        initialWeeks={filterWeekNo}
        initialBuyers={filterBuyer}
      />
    </div>
  );
}

function MetricCard({ title, value, icon, color }: { title: string, value: string, icon: React.ReactNode, color: 'emerald' | 'amber' | 'blue' | 'rose' }) {
  const bgColors = {
    emerald: 'bg-emerald-50 border-emerald-100/50',
    amber: 'bg-amber-50 border-amber-100/50',
    blue: 'bg-blue-50 border-blue-100/50',
    rose: 'bg-rose-50 border-rose-100/50',
  };

  return (
    <div className="bg-white rounded-xl border border-slate-200 p-3 md:p-5 shadow-sm hover:shadow-md transition-shadow duration-200 flex items-center gap-3 md:gap-4 group">
      <div className={cn("p-2 md:p-3 rounded-lg border transition-colors", bgColors[color])}>
        {React.cloneElement(icon as React.ReactElement, { className: "w-4 h-4 md:w-5 md:h-5" })}
      </div>
      <div>
        <p className="text-[10px] md:text-[11px] font-semibold text-slate-500 mb-0.5 uppercase tracking-wider group-hover:text-slate-600 transition-colors truncate">{title}</p>
        <p className="text-xl md:text-2xl font-bold text-slate-900 tracking-tight">{value}</p>
      </div>
    </div>
  );
}

function cn(...classes: (string | undefined | null | false)[]) {
  return classes.filter(Boolean).join(' ');
}
