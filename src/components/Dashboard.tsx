import React, { useState, useRef, useCallback, useMemo, useEffect } from 'react';
import { UploadCloud, FileSpreadsheet, AlertCircle, Search, Package, CheckCircle, TrendingUp, AlertTriangle, Loader2, Printer, BarChart2, Table, RefreshCw } from 'lucide-react';
import * as XLSX from 'xlsx';
import { useVirtualizer } from '@tanstack/react-virtual';
import { db } from '../firebase';
import { doc, getDoc, setDoc, onSnapshot, getDocFromServer, serverTimestamp } from 'firebase/firestore';
import { KPIView } from './KPIView';
import { SewOutReportModal } from './SewOutReportModal';
import { PrintOrientationModal } from './PrintOrientationModal';

enum OperationType {
  CREATE = 'create',
  UPDATE = 'update',
  DELETE = 'delete',
  LIST = 'list',
  GET = 'get',
  WRITE = 'write',
}

interface FirestoreErrorInfo {
  error: string;
  operationType: OperationType;
  path: string | null;
  authInfo: {
    userId?: string | null;
    email?: string | null;
    emailVerified?: boolean | null;
  }
}

function handleFirestoreError(error: unknown, operationType: OperationType, path: string | null) {
  const errInfo: FirestoreErrorInfo = {
    error: error instanceof Error ? error.message : String(error),
    authInfo: {
      userId: null,
      email: null,
      emailVerified: null,
    },
    operationType,
    path
  };
  console.error('Firestore Error: ', JSON.stringify(errInfo));
  throw new Error(JSON.stringify(errInfo));
}

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
  ped?: string;
  weekNo: string;
  coQty: number;
  cumSewInQty: number;
  cumSewOutQty: number;
  cumSewOutRejQty: number;
  cumCTNQty: number;
  statusText: string;
  deliveredQty: number;
  cifOffering?: string;
  remark?: string;
  productionLine?: string;
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

function CifOfferingInput({
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
  const [isLocked, setIsLocked] = useState(initialValue.trim() !== "");
  const [showPrompt, setShowPrompt] = useState(false);
  const [passcode, setPasscode] = useState("");
  const [errorMsg, setErrorMsg] = useState("");
  const inputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    setValue(initialValue);
    setIsLocked(initialValue.trim() !== "");
  }, [initialValue]);

  const getTooltipInfo = () => {
    if (!updatedAt || !initialValue.trim()) return undefined;
    const date = new Date(updatedAt);
    
    const formattedDate = date.toLocaleString('en-US', { 
       year: 'numeric', month: '2-digit', day: '2-digit', 
       hour: '2-digit', minute: '2-digit', hour12: true 
    });

    const day = date.getDay();
    const hour = date.getHours(); 
    
    const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
    const dayNum = d.getUTCDay() || 7;
    d.setUTCDate(d.getUTCDate() + 4 - dayNum);
    const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
    const weekNo = Math.ceil((((d.getTime() - yearStart.getTime()) / 86400000) + 1) / 7);

    const morningShift = weekNo % 2 !== 0 ? 'A' : 'B';
    const eveningShift = weekNo % 2 !== 0 ? 'B' : 'A';
    
    let currentShift = "Off-hours";
    if (day !== 0) {
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
        placeholder={isLocked ? "" : "Add CIF Offering..."}
        list="cif-offering-suggestions"
        className={`w-full text-sm bg-transparent border-0 border-b px-1 py-1 transition-colors ${
          isLocked 
            ? 'border-transparent cursor-pointer hover:bg-slate-100/50 focus:ring-0 focus:outline-none font-medium text-slate-800' 
            : 'border-slate-300 hover:border-slate-300 focus:border-indigo-500 focus:ring-0'
        }`}
      />
      
      {showPrompt && (
        <div className="fixed inset-0 z-[100] flex items-center justify-center bg-slate-900/50 backdrop-blur-sm p-4">
          <div className="bg-white p-5 rounded-xl shadow-xl w-full max-w-xs border border-slate-200">
            <h3 className="text-base font-bold text-slate-900 mb-1">Unlock CIF Offering</h3>
            <p className="text-xs text-slate-500 mb-4">Enter passcode to edit this CIF Offering.</p>
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

function ProductionLineInput({
  initialValue,
  rowId,
  onSave
}: {
  initialValue: string;
  rowId: string;
  onSave: (id: string, text: string) => void;
}) {
  const [value, setValue] = useState(initialValue);
  const [isEditing, setIsEditing] = useState(false);

  useEffect(() => {
    setValue(initialValue);
  }, [initialValue]);

  const handleBlur = () => {
    setIsEditing(false);
    if (value.trim() !== initialValue.trim()) {
      onSave(rowId, value.trim());
    }
  };

  const handleKeyDown = (e: React.KeyboardEvent<HTMLInputElement>) => {
    if (e.key === 'Enter') {
      (e.target as HTMLInputElement).blur();
    } else if (e.key === 'Escape') {
      setValue(initialValue);
      setIsEditing(false);
    }
  };

  return (
    <div className="w-full relative group">
      <input 
        type="text"
        value={value}
        onChange={(e) => setValue(e.target.value)}
        onFocus={() => setIsEditing(true)}
        onBlur={handleBlur}
        onKeyDown={handleKeyDown}
        placeholder={isEditing ? "Type line..." : "-"}
        list="production-line-options"
        className={cn(
          "w-full text-xs font-semibold px-2 py-1 rounded transition-all truncate text-left",
          value.trim() 
            ? "bg-slate-100 hover:bg-indigo-50/80 text-slate-800 hover:text-indigo-900 border border-slate-200 hover:border-indigo-300 focus:bg-white focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500" 
            : "bg-transparent text-slate-400 border border-transparent hover:border-slate-200 hover:bg-slate-50 focus:bg-white focus:border-indigo-400 focus:ring-1 focus:ring-indigo-500"
        )}
        title={value.trim() ? `Production Line: ${value.trim()}` : "Click to assign Production Line"}
      />
    </div>
  );
}

function getRowCifOffering(
  item: ProductionOrder, 
  cifMap: Record<string, { text: string; updatedAt?: number | null }>,
  normalizedMap?: Map<string, { text: string; updatedAt: number | null }>
): { text: string; updatedAt: number | null } {
  if (!cifMap || typeof cifMap !== 'object') return { text: '', updatedAt: null };

  const formatCif = (r: { text: string; updatedAt?: number | null } | undefined): { text: string; updatedAt: number | null } | null => {
    if (!r || !r.text || r.text.trim() === '') return null;
    return { text: r.text.trim(), updatedAt: r.updatedAt ?? null };
  };

  const normalize = (s: string) => s.toLowerCase().replace(/[^a-z0-9]/g, '');

  // 1. Direct match on row.id
  let res = formatCif(cifMap[item.id]);
  if (res) return res;

  // 2. Direct match on row.legacyId
  if (item.legacyId) {
    res = formatCif(cifMap[item.legacyId]);
    if (res) return res;
  }

  // 3. Raw combined keys
  const rawStable = `${item.vpoNo}_${item.scheduleNo}_${item.styleNo}_${item.colorCode}_${item.destination}`;
  res = formatCif(cifMap[rawStable]);
  if (res) return res;

  const rawLegacy = `${item.vpoNo}_${item.scheduleNo}_${item.styleNo}_${item.colorCode}_${item.destination}_${item.planDelDate}`;
  res = formatCif(cifMap[rawLegacy]);
  if (res) return res;

  // 4. Normalized fuzzy search
  const normStable = normalize(rawStable);
  const normLegacy = normalize(rawLegacy);
  
  if (normalizedMap) {
    const found = normalizedMap.get(normStable) || normalizedMap.get(normLegacy);
    if (found) return found;
  }

  if (item.cifOffering) {
    return { text: item.cifOffering, updatedAt: null };
  }

  return { text: '', updatedAt: null };
}

const TABLE_COLUMN_WIDTHS = [
  '120px', // Plan Del Date
  '110px', // PED
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
  '200px', // CIF Offering
  '135px', // Production Lines
  '115px', // Delivered Qty
];

export function getSixWeeksWindow(): { start: Date; end: Date; weekNumbers: Set<string> } {
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const dayOfWeek = today.getDay(); // 0 is Sun, 1 is Mon...
  const diffToMonday = today.getDate() - dayOfWeek + (dayOfWeek === 0 ? -6 : 1);
  const currentMonday = new Date(today.getFullYear(), today.getMonth(), diffToMonday, 0, 0, 0, 0);
  
  // Allow a grace window of 2 weeks prior so active weeks 41-42 with remarks are preserved
  const start = new Date(currentMonday.getTime() - (14 * 24 * 60 * 60 * 1000));
  
  // 6 weeks forward from current Monday (ending on Sunday 23:59:59.999 of week 6)
  const end = new Date(currentMonday.getTime() + (42 * 24 * 60 * 60 * 1000) - 1);
  
  const weekNumbers = new Set<string>();
  // Include from 2 weeks prior to 6 weeks forward (total 8-9 weeks)
  for (let i = -2; i < 7; i++) {
    const midWeek = new Date(currentMonday.getTime() + (i * 7 + 3) * 24 * 60 * 60 * 1000);
    const d = new Date(Date.UTC(midWeek.getFullYear(), midWeek.getMonth(), midWeek.getDate()));
    const dayNum = d.getUTCDay() || 7;
    d.setUTCDate(d.getUTCDate() + 4 - dayNum);
    const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
    const w = Math.ceil((((d.getTime() - yearStart.getTime()) / 86400000) + 1) / 7);
    weekNumbers.add(w.toString());
  }

  // Explicitly ensure weeks 41 through 48 are always present in the active set
  ['41', '42', '43', '44', '45', '46', '47', '48'].forEach(w => weekNumbers.add(w));

  return { start, end, weekNumbers };
}

function parseExcelDate(val: any): Date | null {
  if (val == null || val === '') return null;
  if (val instanceof Date && !isNaN(val.getTime())) return val;
  
  // Handle numeric Excel date serials (e.g. 45604 or "45604")
  const numVal = Number(val);
  if (!isNaN(numVal) && numVal > 30000 && numVal < 60000) {
    return new Date(Math.round((numVal - 25569) * 86400 * 1000));
  }

  const str = String(val).trim();
  if (!str) return null;

  if (/^\d{8}$/.test(str)) {
    const y = parseInt(str.substring(0, 4), 10);
    const m = parseInt(str.substring(4, 6), 10) - 1;
    const d = parseInt(str.substring(6, 8), 10);
    return new Date(y, m, d);
  }
  if (/^\d{4}[\/\-\.]\d{1,2}[\/\-\.]\d{1,2}$/.test(str)) {
    const parts = str.split(/[\/\-\.]/);
    return new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
  }
  if (/^\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{4}$/.test(str)) {
    const parts = str.split(/[\/\-\.]/);
    const d = parseInt(parts[0], 10);
    const m = parseInt(parts[1], 10) - 1;
    const y = parseInt(parts[2], 10);
    return new Date(y, m, d);
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
  if (weekNo) {
    const cleanWeek = weekNo.trim().replace(/^W(eek)?[\s\-_]*/i, '');
    if (window.weekNumbers.has(cleanWeek) || window.weekNumbers.has(weekNo.trim())) {
      return true;
    }
  }
  return false;
}

function getVal(row: any, searchKeys: string[]): any {
  if (!row || typeof row !== 'object') return undefined;
  for (const k of searchKeys) {
    if (row[k] !== undefined && row[k] !== null && String(row[k]).trim() !== '') {
      return row[k];
    }
  }
  const normKeys = searchKeys.map(k => k.toLowerCase().replace(/[^a-z0-9]/g, ''));
  for (const key of Object.keys(row)) {
    const normKey = key.toLowerCase().replace(/[^a-z0-9]/g, '');
    if (normKeys.includes(normKey)) {
      if (row[key] !== undefined && row[key] !== null && String(row[key]).trim() !== '') {
        return row[key];
      }
    }
  }
  return undefined;
}

async function safeJson(res: Response) {
  try {
    const text = await res.text();
    if (!text || text.trim() === '') return null;
    return JSON.parse(text);
  } catch (e) {
    return null;
  }
}

function getNum(row: any, searchKeys: string[]): number {
  const val = getVal(row, searchKeys);
  if (val == null || val === '') return 0;
  if (typeof val === 'number') return val;
  const parsed = Number(String(val).replace(/,/g, ''));
  return isNaN(parsed) ? 0 : parsed;
}

function getRowRemark(
  item: ProductionOrder, 
  remarksMap: Record<string, { text: string; updatedAt?: number | null }>,
  normalizedMap?: Map<string, { text: string; updatedAt: number | null }>
): { text: string; updatedAt: number | null } {
  if (!remarksMap || typeof remarksMap !== 'object') return { text: '', updatedAt: null };

  const formatRemark = (r: { text: string; updatedAt?: number | null } | undefined): { text: string; updatedAt: number | null } | null => {
    if (!r || !r.text || r.text.trim() === '') return null;
    return { text: r.text.trim(), updatedAt: r.updatedAt ?? null };
  };

  const normalize = (s: string) => s.toLowerCase().replace(/[^a-z0-9]/g, '');

  // 1. Direct match on row.id
  let res = formatRemark(remarksMap[item.id]);
  if (res) return res;

  // 2. Direct match on row.legacyId
  if (item.legacyId) {
    res = formatRemark(remarksMap[item.legacyId]);
    if (res) return res;
  }

  // 3. Raw combined keys
  const rawStable = `${item.vpoNo}_${item.scheduleNo}_${item.styleNo}_${item.colorCode}_${item.destination}`;
  res = formatRemark(remarksMap[rawStable]);
  if (res) return res;

  const rawLegacy = `${item.vpoNo}_${item.scheduleNo}_${item.styleNo}_${item.colorCode}_${item.destination}_${item.planDelDate}`;
  res = formatRemark(remarksMap[rawLegacy]);
  if (res) return res;

  // 4. Normalized fuzzy search using pre-calculated Map if available
  const normStable = normalize(rawStable);
  const normLegacy = normalize(rawLegacy);
  
  if (normalizedMap) {
    const found = normalizedMap.get(normStable) || normalizedMap.get(normLegacy);
    if (found) return found;
  } else {
    // Fallback if Map not provided (slow)
    if (normStable || normLegacy) {
      for (const [k, v] of Object.entries(remarksMap)) {
        const normK = normalize(k);
        if (v?.text && (normK === normStable || normK === normLegacy)) {
          const formatted = formatRemark(v);
          if (formatted) return formatted;
        }
      }
    }
  }

  if (item.remark) {
    return { text: item.remark, updatedAt: null };
  }

  return { text: '', updatedAt: null };
}

function getRowProductionLine(
  item: ProductionOrder, 
  linesMap: Record<string, { text: string; updatedAt?: number | null }>
): { text: string; updatedAt: number | null } {
  if (!linesMap || typeof linesMap !== 'object') {
    return { text: '', updatedAt: null };
  }

  const formatLine = (r: { text: string; updatedAt?: number | null } | undefined): { text: string; updatedAt: number | null } | null => {
    if (!r || !r.text || r.text.trim() === '') return null;
    return { text: r.text.trim(), updatedAt: r.updatedAt ?? null };
  };

  // 1. Direct match on row.id
  let res = formatLine(linesMap[item.id]);
  if (res) return res;

  // 2. Direct match on row.legacyId
  if (item.legacyId) {
    res = formatLine(linesMap[item.legacyId]);
    if (res) return res;
  }

  // 3. Raw combined keys
  const rawStable = `${item.vpoNo}_${item.scheduleNo}_${item.styleNo}_${item.colorCode}_${item.destination}`;
  res = formatLine(linesMap[rawStable]);
  if (res) return res;

  return { text: '', updatedAt: null };
}

function isFirestoreQuotaError(msg: string): boolean {
  return /quota exceeded|resource-exhausted|free daily read units/i.test(msg);
}

async function compressOrdersToBase64(orders: ProductionOrder[]): Promise<string | null> {
  if (typeof CompressionStream === 'undefined') return null;
  try {
    const tuples = orders.map(item => [
      item.id || '', item.legacyId || '', item.buyer || '', item.groupTechClass || '', item.buyerDivisionName || '',
      item.styleNo || '', item.custStyleNo || '', item.vpoNo || '', item.shipmentMode || '', item.colorCode || '',
      item.colorName || '', item.destination || '', item.packMethod || '', item.scheduleNo || '', item.planDelDate || '',
      item.ped || '', item.weekNo || '', Number(item.coQty) || 0, Number(item.cumSewInQty) || 0, Number(item.cumSewOutQty) || 0,
      Number(item.cumSewOutRejQty) || 0, Number(item.cumCTNQty) || 0, item.statusText || '', Number(item.deliveredQty) || 0,
      item.cifOffering || '', item.remark || ''
    ]);
    const jsonBytes = new TextEncoder().encode(JSON.stringify(tuples));
    const stream = new Blob([jsonBytes]).stream().pipeThrough(new CompressionStream('gzip'));
    const compressedBuf = await new Response(stream).arrayBuffer();
    const bytes = new Uint8Array(compressedBuf);
    let binary = '';
    const CHUNK = 0x4000;
    for (let i = 0; i < bytes.length; i += CHUNK) {
      binary += String.fromCharCode(...bytes.subarray(i, i + CHUNK));
    }
    return btoa(binary);
  } catch (e) {
    console.warn('Gzip compression fallback:', e);
    return null;
  }
}

async function decompressOrdersFromBase64(base64: string): Promise<ProductionOrder[]> {
  if (typeof DecompressionStream === 'undefined') {
    throw new Error('Browser does not support DecompressionStream');
  }
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream('gzip'));
  const decompressedText = await new Response(stream).text();
  const parsed = JSON.parse(decompressedText);
  if (!Array.isArray(parsed)) return [];
  return parsed.map((row: any) => {
    if (!Array.isArray(row)) return row as ProductionOrder;
    return {
      id: String(row[0] ?? ''),
      legacyId: String(row[1] ?? ''),
      buyer: String(row[2] ?? ''),
      groupTechClass: String(row[3] ?? ''),
      buyerDivisionName: String(row[4] ?? ''),
      styleNo: String(row[5] ?? ''),
      custStyleNo: String(row[6] ?? ''),
      vpoNo: String(row[7] ?? ''),
      shipmentMode: String(row[8] ?? ''),
      colorCode: String(row[9] ?? ''),
      colorName: String(row[10] ?? ''),
      destination: String(row[11] ?? ''),
      packMethod: String(row[12] ?? ''),
      scheduleNo: String(row[13] ?? ''),
      planDelDate: String(row[14] ?? ''),
      ped: String(row[15] ?? ''),
      weekNo: String(row[16] ?? ''),
      coQty: Number(row[17]) || 0,
      cumSewInQty: Number(row[18]) || 0,
      cumSewOutQty: Number(row[19]) || 0,
      cumSewOutRejQty: Number(row[20]) || 0,
      cumCTNQty: Number(row[21]) || 0,
      statusText: String(row[22] ?? ''),
      deliveredQty: Number(row[23]) || 0,
      cifOffering: String(row[24] ?? ''),
      remark: String(row[25] ?? ''),
    };
  });
}

export function Dashboard() {
  const [data, setData] = useState<ProductionOrder[] | null>(() => {
    try {
      const dataKeys = ['production_dashboard_data', 'production_data', 'dashboard_data', 'orders_data', 'excel_data'];
      for (const k of dataKeys) {
        const val = localStorage.getItem(k);
        if (val) {
          const parsed = JSON.parse(val);
          if (Array.isArray(parsed) && parsed.length > 0) return parsed;
        }
      }
    } catch (e) {}
    return null;
  });

  const [lastUpdated, setLastUpdated] = useState<number | null>(() => {
    try {
      const val = localStorage.getItem('production_dashboard_last_updated');
      if (val) return Number(val);
    } catch (e) {}
    return null;
  });

  const lastUpdatedRef = useRef<number>(lastUpdated || 0);
  const dataRef = useRef<ProductionOrder[] | null>(data);
  useEffect(() => {
    dataRef.current = data;
  }, [data]);

  const appliedCloudUploadIdRef = useRef<string | null>((() => {
    try {
      if (data && data.length > 0) {
        return localStorage.getItem('production_dashboard_upload_id');
      }
    } catch (e) {}
    return null;
  })());

  const [isDragging, setIsDragging] = useState(false);
  const [loadingState, setLoadingState] = useState<'idle' | 'reading' | 'parsing' | 'uploading'>('idle');
  const [uploadProgress, setUploadProgress] = useState({ current: 0, total: 0 });
  const [syncStatus, setSyncStatus] = useState<'connecting' | 'synced' | 'syncing' | 'error'>('connecting');
  const [syncError, setSyncError] = useState<string | null>(null);
  const [listenerNonce, setListenerNonce] = useState(0);
  const retryAttemptRef = useRef(0);

  const [remarks, setRemarks] = useState<Record<string, { text: string, updatedAt: number | null }>>(() => {
    try {
      const val = localStorage.getItem('production_dashboard_remarks');
      if (val) return JSON.parse(val);
    } catch (e) {}
    return {};
  });

  const [productionLines, setProductionLines] = useState<Record<string, { text: string, updatedAt: number | null }>>(() => {
    try {
      const val = localStorage.getItem('production_dashboard_production_lines');
      if (val) return JSON.parse(val);
    } catch (e) {}
    return {};
  });

  const [cifOfferings, setCifOfferings] = useState<Record<string, { text: string, updatedAt: number | null }>>(() => {
    try {
      const val = localStorage.getItem('production_dashboard_cif_offerings');
      if (val) return JSON.parse(val);
    } catch (e) {}
    return {};
  });

  const saveLocalDashboardCache = useCallback((orders: ProductionOrder[], timestamp: number, uploadId?: string | null) => {
    try {
      // Remove legacy duplicate keys to prevent hitting browser 5MB localStorage quota
      ['production_data', 'dashboard_data', 'orders_data', 'excel_data'].forEach(k => {
        try { localStorage.removeItem(k); } catch {}
      });
      localStorage.setItem('production_dashboard_data', JSON.stringify(orders));
      localStorage.setItem('production_dashboard_last_updated', String(timestamp));
      if (uploadId) {
        localStorage.setItem('production_dashboard_upload_id', uploadId);
      }
    } catch (e) {
      console.warn('Local cache save warning:', e);
    }
  }, []);

  const updateDashboardData = useCallback((
    newOrders: ProductionOrder[],
    timestamp: number,
    source: 'firestore_v3' | 'upload' | 'manual_sync' | 'server_api' | 'sse_init' | 'sse_dashboard',
    uploadId?: string | null
  ) => {
    if (!newOrders || !Array.isArray(newOrders) || newOrders.length === 0) return;

    const isAuthoritative = source === 'firestore_v3' || source === 'upload' || source === 'manual_sync';
    const hasNoData = !dataRef.current || dataRef.current.length === 0;
    const isNewerTimestamp = timestamp > lastUpdatedRef.current;

    if (isAuthoritative || hasNoData || (isNewerTimestamp && (!uploadId || uploadId !== appliedCloudUploadIdRef.current))) {
      lastUpdatedRef.current = timestamp;
      if (uploadId) {
        appliedCloudUploadIdRef.current = uploadId;
      }
      dataRef.current = newOrders;
      setLastUpdated(timestamp);
      setData(newOrders);
      saveLocalDashboardCache(newOrders, timestamp, uploadId);
    }
  }, [saveLocalDashboardCache]);
  
  // Pre-calculate normalized remarks map for fast fuzzy matching
  const normalizedRemarksMap = useMemo(() => {
    const map = new Map<string, { text: string; updatedAt: number | null }>();
    const normalize = (s: string) => s.toLowerCase().replace(/[^a-z0-9]/g, '');
    
    Object.entries(remarks).forEach(([id, data]) => {
      const remarkData = data as { text: string; updatedAt: number | null };
      if (remarkData.text && remarkData.text.trim() !== '') {
        const normKey = normalize(id);
        if (normKey) {
          // Keep the newest remark if there's a collision in normalized keys
          const existing = map.get(normKey);
          if (!existing || (remarkData.updatedAt || 0) > (existing.updatedAt || 0)) {
            map.set(normKey, { text: remarkData.text.trim(), updatedAt: remarkData.updatedAt || null });
          }
        }
      }
    });
    return map;
  }, [remarks]);

  const normalizedCifOfferingsMap = useMemo(() => {
    const map = new Map<string, { text: string; updatedAt: number | null }>();
    const normalize = (s: string) => s.toLowerCase().replace(/[^a-z0-9]/g, '');
    
    Object.entries(cifOfferings).forEach(([id, data]) => {
      const cifData = data as { text: string; updatedAt: number | null };
      if (cifData.text && cifData.text.trim() !== '') {
        const normKey = normalize(id);
        if (normKey) {
          const existing = map.get(normKey);
          if (!existing || (cifData.updatedAt || 0) > (existing.updatedAt || 0)) {
            map.set(normKey, { text: cifData.text.trim(), updatedAt: cifData.updatedAt || null });
          }
        }
      }
    });
    return map;
  }, [cifOfferings]);

  const fileInputRef = useRef<HTMLInputElement>(null);
  const [error, setError] = useState<string | null>(null);
  const [viewMode, setViewMode] = useState<'table' | 'kpi'>('table');

  // Filters State
  const [searchTerm, setSearchTerm] = useState('');
  const [filterBuyer, setFilterBuyer] = useState<string[]>([]);
  const [filterWeekNo, setFilterWeekNo] = useState<string[]>([]);
  const [filterStatus, setFilterStatus] = useState<string[]>([]);
  const [filterShipmentMode, setFilterShipmentMode] = useState<string[]>([]);
  const [filterDestination, setFilterDestination] = useState<string[]>([]);
  const [filterPackMethod, setFilterPackMethod] = useState<string[]>([]);
  const [filterCifOffering, setFilterCifOffering] = useState<string[]>([]);
  const [filterRemark, setFilterRemark] = useState<string[]>([]);
  const [filterPED, setFilterPED] = useState<string[]>([]);
  const [filterProductionLine, setFilterProductionLine] = useState<string[]>([]);
  
  const tableContainerRef = useRef<HTMLDivElement>(null);
  const [isMobile, setIsMobile] = useState(typeof window !== 'undefined' ? window.innerWidth < 768 : false);

  const [showPrintModal, setShowPrintModal] = useState(false);
  const [showSewOutReport, setShowSewOutReport] = useState(false);
  const [reportInitialMode, setReportInitialMode] = useState<'sewout_50_100' | 'all'>('all');
  const [recoveryMessage, setRecoveryMessage] = useState<string | null>(null);
  const [isRecovering, setIsRecovering] = useState(false);

  // Helper to apply a dashboard_v3/main snapshot and fetch all chunk parts reliably
  const applyDashboardSnapshot = useCallback(async (mainData: any, forceServerParts = false) => {
    if (!mainData) return;

    const cloudUploadId = String(mainData.uploadId || mainData.lastUpdated || '');
    const expectedTotal = typeof mainData.totalCount === 'number' ? mainData.totalCount : null;

    // Skip redundant part reads if this exact cloud upload is already loaded cleanly in memory
    if (
      !forceServerParts &&
      dataRef.current &&
      dataRef.current.length > 0 &&
      cloudUploadId &&
      appliedCloudUploadIdRef.current === cloudUploadId &&
      (expectedTotal === null || dataRef.current.length === expectedTotal)
    ) {
      if (mainData.lastUpdated && mainData.lastUpdated !== lastUpdatedRef.current) {
        lastUpdatedRef.current = mainData.lastUpdated;
        setLastUpdated(mainData.lastUpdated);
      }
      setSyncStatus('synced');
      setSyncError(null);
      return;
    }

    let allOrders: ProductionOrder[] = [];

    if (typeof mainData.compressedGzip === 'string' && mainData.compressedGzip.length > 0) {
      allOrders = await decompressOrdersFromBase64(mainData.compressedGzip);
    } else {
      allOrders = Array.isArray(mainData.orders) ? [...mainData.orders] : [];
      const chunkCount = mainData.chunkCount || 1;

      if (chunkCount > 1) {
        setSyncStatus('syncing');
        const fetchPart = async (index: number, attempt = 0): Promise<ProductionOrder[]> => {
          const partRef = doc(db, 'dashboard_v3', `part_${index}`);
          let partSnap;
          try {
            partSnap = (forceServerParts || attempt > 0)
              ? await getDocFromServer(partRef)
              : await getDoc(partRef);
          } catch {
            partSnap = await getDoc(partRef);
          }

          if (!partSnap.exists()) {
            if (attempt < 2) {
              await new Promise(r => setTimeout(r, 500 * (attempt + 1)));
              return fetchPart(index, attempt + 1);
            }
            throw new Error(`Missing orderbook chunk part_${index}`);
          }

          const partData = partSnap.data();
          if (mainData.uploadId && partData?.uploadId && partData.uploadId !== mainData.uploadId && attempt < 2) {
            await new Promise(r => setTimeout(r, 500 * (attempt + 1)));
            return fetchPart(index, attempt + 1);
          }

          return Array.isArray(partData?.orders) ? partData.orders : [];
        };

        const partIndices = Array.from({ length: chunkCount - 1 }, (_, idx) => idx + 1);
        const partResults = await Promise.all(partIndices.map(i => fetchPart(i)));
        for (const chunkOrders of partResults) {
          allOrders = allOrders.concat(chunkOrders);
        }
      }
    }

    if (allOrders.length > 0) {
      const ts = mainData.lastUpdated || Date.now();
      updateDashboardData(
        allOrders,
        ts,
        forceServerParts ? 'manual_sync' : 'firestore_v3',
        cloudUploadId || String(ts)
      );
      setSyncStatus('synced');
      setSyncError(null);
    }
  }, [updateDashboardData]);

  const handleManualCloudSync = useCallback(async () => {
    setSyncStatus('syncing');
    setSyncError(null);
    try {
      const mainSnap = await getDocFromServer(doc(db, 'dashboard_v3', 'main'));
      if (mainSnap.exists()) {
        await applyDashboardSnapshot(mainSnap.data(), true);
      } else {
        setSyncStatus('synced');
      }
    } catch (err: any) {
      const msg = err?.message || 'Unable to reach cloud sync server';
      setSyncStatus('error');
      setSyncError(isFirestoreQuotaError(msg) ? 'Firestore free daily read quota exceeded. It resets in 24h (or upgrade to Blaze).' : msg);
    }
  }, [applyDashboardSnapshot]);

  // 1. Re-connect automatically when network comes back online
  useEffect(() => {
    const handleOnline = () => {
      retryAttemptRef.current = 0;
      setListenerNonce(n => n + 1);
    };
    window.addEventListener('online', handleOnline);
    return () => window.removeEventListener('online', handleOnline);
  }, []);

  // 2. Real-time Sync from Firestore (Single-Doc Compressed Low-Read Mode for Dashboard, Remarks, Lines, and CIF)
  useEffect(() => {
    let isMounted = true;
    let retryTimer: ReturnType<typeof setTimeout> | null = null;
    setSyncStatus('connecting');

    // A. Dashboard Data Sync (v3 main document)
    const unsubDashboardV3 = onSnapshot(doc(db, 'dashboard_v3', 'main'), async (snap) => {
      if (!isMounted) return;
      retryAttemptRef.current = 0;
      if (!snap.exists()) {
        setSyncStatus('synced');
        return;
      }
      // Wait for server confirmation if this is only our own local pending write
      if (snap.metadata.hasPendingWrites) return;

      const mainData = snap.data();
      if (!mainData) return;

      try {
        await applyDashboardSnapshot(mainData, false);
      } catch (err: any) {
        if (!isMounted) return;
        const msg = err?.message || String(err);
        console.warn('Dashboard V3 chunk load error:', msg);
        setSyncStatus('error');
        setSyncError(isFirestoreQuotaError(msg) ? 'Firestore free daily read quota exceeded. Resets in 24h.' : `Cloud sync error: ${msg}`);
      }
    }, (err: any) => {
      if (!isMounted) return;
      const msg = err?.message || String(err);
      console.warn('Dashboard V3 sync error:', msg);
      setSyncStatus('error');
      if (isFirestoreQuotaError(msg)) {
        const quotaMsg = 'Firestore free tier daily quota exceeded. Resets in 24h — upgrade Firebase Blaze or wait for daily reset.';
        setSyncError(quotaMsg);
        setError(quotaMsg);
      } else {
        setSyncError('Cloud listener disconnected. Reconnecting...');
        // Exponential backoff (30s, 60s, 120s, max 300s) to avoid burning read quota on flaky networks
        const delayMs = Math.min(300000, 30000 * Math.pow(2, retryAttemptRef.current));
        retryAttemptRef.current += 1;
        retryTimer = setTimeout(() => {
          if (isMounted && !document.hidden) {
            setListenerNonce(n => n + 1);
          }
        }, delayMs);
      }
    });

    // B. Remarks Sync (remarks_v2/all)
    const unsubRemarks = onSnapshot(doc(db, 'remarks_v2', 'all'), (snap) => {
      if (snap.exists() && isMounted) {
        const newRemarks = snap.data() as Record<string, { text: string, updatedAt: number | null }>;
        setRemarks(prev => {
          const merged = { ...prev };
          let changed = false;
          Object.entries(newRemarks).forEach(([id, data]) => {
            if (data && typeof data === 'object') {
              const prevUpdated = merged[id]?.updatedAt || 0;
              const newUpdated = data.updatedAt || 0;
              if (!merged[id] || newUpdated >= prevUpdated) {
                if (merged[id]?.text !== data.text || merged[id]?.updatedAt !== data.updatedAt) {
                  merged[id] = data;
                  changed = true;
                }
              }
            }
          });
          if (changed) {
            try {
              localStorage.setItem('production_dashboard_remarks', JSON.stringify(merged));
            } catch (e) {}
            return merged;
          }
          return prev;
        });
      }
    }, (err) => {
      console.warn('Remarks sync error:', err);
    });

    // C. Production Lines Sync (production_lines_v2/all)
    const unsubProductionLines = onSnapshot(doc(db, 'production_lines_v2', 'all'), (snap) => {
      if (snap.exists() && isMounted) {
        const newLines = snap.data() as Record<string, { text: string, updatedAt: number | null }>;
        setProductionLines(prev => {
          const merged = { ...prev };
          let changed = false;
          Object.entries(newLines).forEach(([id, data]) => {
            if (data && typeof data === 'object') {
              const prevUpdated = merged[id]?.updatedAt || 0;
              const newUpdated = data.updatedAt || 0;
              if (!merged[id] || newUpdated >= prevUpdated) {
                if (merged[id]?.text !== data.text || merged[id]?.updatedAt !== data.updatedAt) {
                  merged[id] = data;
                  changed = true;
                }
              }
            }
          });
          if (changed) {
            try {
              localStorage.setItem('production_dashboard_production_lines', JSON.stringify(merged));
            } catch (e) {}
            return merged;
          }
          return prev;
        });
      }
    }, (err) => {
      console.warn('Production lines sync error:', err);
    });

    // D. CIF Offerings Sync (cif_offerings_v2/all)
    const unsubCifOfferings = onSnapshot(doc(db, 'cif_offerings_v2', 'all'), (snap) => {
      if (snap.exists() && isMounted) {
        const newCif = snap.data() as Record<string, { text: string, updatedAt: number | null }>;
        setCifOfferings(prev => {
          const merged = { ...prev };
          let changed = false;
          Object.entries(newCif).forEach(([id, data]) => {
            if (data && typeof data === 'object') {
              const prevUpdated = merged[id]?.updatedAt || 0;
              const newUpdated = data.updatedAt || 0;
              if (!merged[id] || newUpdated >= prevUpdated) {
                if (merged[id]?.text !== data.text || merged[id]?.updatedAt !== data.updatedAt) {
                  merged[id] = data;
                  changed = true;
                }
              }
            }
          });
          if (changed) {
            try {
              localStorage.setItem('production_dashboard_cif_offerings', JSON.stringify(merged));
            } catch (e) {}
            return merged;
          }
          return prev;
        });
      }
    }, (err) => {
      console.warn('CIF offerings sync error:', err);
    });

    return () => {
      isMounted = false;
      if (retryTimer) clearTimeout(retryTimer);
      unsubDashboardV3();
      unsubRemarks();
      unsubProductionLines();
      unsubCifOfferings();
    };
  }, [applyDashboardSnapshot, listenerNonce]);

  // 3. Multi-PC Real-Time Sync & Server Backup Integration (SSE + /api/remarks + /api/dashboard)
  useEffect(() => {
    let isMounted = true;

    // Load remarks immediately from server backup
    fetch('/api/remarks')
      .then(res => safeJson(res))
      .then((serverRemarks: Record<string, { text: string; updatedAt: number | null }> | null) => {
        if (!isMounted || !serverRemarks || typeof serverRemarks !== 'object') return;
        if (Object.keys(serverRemarks).length > 0) {
          setRemarks(prev => {
            const merged = { ...serverRemarks, ...prev };
            try {
              localStorage.setItem('production_dashboard_remarks', JSON.stringify(merged));
            } catch (e) {}
            return merged;
          });
        }
      })
      .catch(() => {});

    // Load production lines immediately from server backup
    fetch('/api/production-lines')
      .then(res => safeJson(res))
      .then((serverLines: Record<string, { text: string; updatedAt: number | null }> | null) => {
        if (!isMounted || !serverLines || typeof serverLines !== 'object') return;
        if (Object.keys(serverLines).length > 0) {
          setProductionLines(prev => {
            const merged = { ...serverLines, ...prev };
            try {
              localStorage.setItem('production_dashboard_production_lines', JSON.stringify(merged));
            } catch (e) {}
            return merged;
          });
        }
      })
      .catch(() => {});

    // Load CIF offerings immediately from server backup
    fetch('/api/cif-offerings')
      .then(res => safeJson(res))
      .then((serverCif: Record<string, { text: string; updatedAt: number | null }> | null) => {
        if (!isMounted || !serverCif || typeof serverCif !== 'object') return;
        if (Object.keys(serverCif).length > 0) {
          setCifOfferings(prev => {
            const merged = { ...serverCif, ...prev };
            try {
              localStorage.setItem('production_dashboard_cif_offerings', JSON.stringify(merged));
            } catch (e) {}
            return merged;
          });
        }
      })
      .catch(() => {});

    // Load dashboard order items from server backup ONLY if newer or no data loaded yet
    fetch('/api/dashboard')
      .then(res => safeJson(res))
      .then(serverDashboard => {
        if (!isMounted || !serverDashboard?.data || !Array.isArray(serverDashboard.data)) return;
        if (serverDashboard.data.length > 0 && serverDashboard.lastUpdated) {
          updateDashboardData(serverDashboard.data, serverDashboard.lastUpdated, 'server_api', serverDashboard.uploadId);
        }
      })
      .catch(() => {});

    // Connect to Server-Sent Events (SSE) for instant same-server updates
    let eventSource: EventSource | null = null;
    try {
      eventSource = new EventSource('/api/realtime');
      eventSource.onmessage = (event) => {
        if (!isMounted || !event.data) return;
        try {
          const payload = JSON.parse(event.data);
          if (payload.type === 'init') {
            if (payload.remarks && typeof payload.remarks === 'object') {
              setRemarks(prev => ({ ...payload.remarks, ...prev }));
            }
            if (payload.productionLines && typeof payload.productionLines === 'object') {
              setProductionLines(prev => ({ ...payload.productionLines, ...prev }));
            }
            if (payload.cifOfferings && typeof payload.cifOfferings === 'object') {
              setCifOfferings(prev => ({ ...payload.cifOfferings, ...prev }));
            }
            if (payload.dashboard?.data && Array.isArray(payload.dashboard.data) && payload.dashboard.data.length > 0 && payload.dashboard.lastUpdated) {
              updateDashboardData(payload.dashboard.data, payload.dashboard.lastUpdated, 'sse_init', payload.dashboard.uploadId);
            }
          } else if (payload.type === 'remark' && payload.id) {
            setRemarks(prev => {
              const next = { ...prev, [payload.id]: { text: payload.text, updatedAt: payload.updatedAt } };
              try {
                localStorage.setItem('production_dashboard_remarks', JSON.stringify(next));
              } catch (e) {}
              return next;
            });
          } else if (payload.type === 'remarks_batch' && payload.remarks) {
            setRemarks(prev => {
              const merged = { ...prev, ...payload.remarks };
              try {
                localStorage.setItem('production_dashboard_remarks', JSON.stringify(merged));
              } catch (e) {}
              return merged;
            });
          } else if (payload.type === 'production_line' && payload.id) {
            setProductionLines(prev => {
              const next = { ...prev, [payload.id]: { text: payload.text, updatedAt: payload.updatedAt } };
              try {
                localStorage.setItem('production_dashboard_production_lines', JSON.stringify(next));
              } catch (e) {}
              return next;
            });
          } else if (payload.type === 'production_lines_batch' && payload.lines) {
            setProductionLines(prev => {
              const merged = { ...prev, ...payload.lines };
              try {
                localStorage.setItem('production_dashboard_production_lines', JSON.stringify(merged));
              } catch (e) {}
              return merged;
            });
          } else if (payload.type === 'cif_offering' && payload.id) {
            setCifOfferings(prev => {
              const next = { ...prev, [payload.id]: { text: payload.text, updatedAt: payload.updatedAt } };
              try {
                localStorage.setItem('production_dashboard_cif_offerings', JSON.stringify(next));
              } catch (e) {}
              return next;
            });
          } else if (payload.type === 'cif_offerings_batch' && payload.cifOfferings) {
            setCifOfferings(prev => {
              const merged = { ...prev, ...payload.cifOfferings };
              try {
                localStorage.setItem('production_dashboard_cif_offerings', JSON.stringify(merged));
              } catch (e) {}
              return merged;
            });
          } else if (payload.type === 'dashboard' && payload.data && Array.isArray(payload.data) && payload.lastUpdated) {
            updateDashboardData(payload.data, payload.lastUpdated, 'sse_dashboard', payload.uploadId);
          }
        } catch (e) {}
      };
    } catch (e) {}

    return () => {
      isMounted = false;
      eventSource?.close();
    };
  }, [updateDashboardData]);

  const handleRemarkChange = async (id: string, text: string) => {
    const updated = Date.now();
    const remarkData = { text, updatedAt: updated };
    
    // 1. Update local state immediately for responsiveness
    setRemarks(prev => {
      const next = { ...prev, [id]: remarkData };
      try {
        localStorage.setItem('production_dashboard_remarks', JSON.stringify(next));
      } catch (e) {}
      return next;
    });

    // 2. Persist to Firestore (Primary) - Direct setDoc with merge for maximum reliability & instant cross-PC sync
    try {
      const remarksRef = doc(db, 'remarks_v2', 'all');
      await setDoc(remarksRef, { [id]: remarkData }, { merge: true });
    } catch (err) {
      console.warn('Firestore Save Notice for remark:', err);
    }

    // 3. Sync to Express Server API for redundancy & local SSE broadcast
    fetch('/api/remarks', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ id, text, updatedAt: updated })
    }).catch(() => {});
  };

  const handleCifOfferingChange = async (id: string, text: string) => {
    const updated = Date.now();
    const cifData = { text, updatedAt: updated };
    
    setCifOfferings(prev => {
      const next = { ...prev, [id]: cifData };
      try {
        localStorage.setItem('production_dashboard_cif_offerings', JSON.stringify(next));
      } catch (e) {}
      return next;
    });

    try {
      const cifRef = doc(db, 'cif_offerings_v2', 'all');
      await setDoc(cifRef, { [id]: cifData }, { merge: true });
    } catch (err) {
      console.warn('Firestore Save Notice for CIF offering:', err);
    }

    fetch('/api/cif-offerings', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ id, text, updatedAt: updated })
    }).catch(() => {});
  };

  const handleProductionLineChange = async (id: string, text: string) => {
    const updated = Date.now();
    const lineData = { text, updatedAt: updated };
    
    setProductionLines(prev => {
      const next = { ...prev, [id]: lineData };
      try {
        localStorage.setItem('production_dashboard_production_lines', JSON.stringify(next));
      } catch (e) {}
      return next;
    });

    try {
      const linesRef = doc(db, 'production_lines_v2', 'all');
      await setDoc(linesRef, { [id]: lineData }, { merge: true });
    } catch (err) {
      console.warn('Firestore Save Notice for production line:', err);
    }

    fetch('/api/production-lines', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ id, text, updatedAt: updated })
    }).catch(() => {});
  };

  const persistDashboardData = async (parsedData: ProductionOrder[]) => {
    setLoadingState('uploading');
    setSyncStatus('syncing');
    setSyncError(null);
    setUploadProgress({ current: 0, total: parsedData.length });

    // Sanitize all order fields to ensure valid Firestore types (no undefined/NaN)
    const cleanedData: ProductionOrder[] = parsedData.map(item => ({
      id: String(item.id || ''),
      legacyId: String(item.legacyId || ''),
      buyer: String(item.buyer || ''),
      groupTechClass: String(item.groupTechClass || ''),
      buyerDivisionName: String(item.buyerDivisionName || ''),
      styleNo: String(item.styleNo || ''),
      custStyleNo: String(item.custStyleNo || ''),
      vpoNo: String(item.vpoNo || ''),
      shipmentMode: String(item.shipmentMode || ''),
      colorCode: String(item.colorCode || ''),
      colorName: String(item.colorName || ''),
      destination: String(item.destination || ''),
      packMethod: String(item.packMethod || ''),
      scheduleNo: String(item.scheduleNo || ''),
      planDelDate: String(item.planDelDate || ''),
      ped: String(item.ped || ''),
      weekNo: String(item.weekNo || ''),
      coQty: Number.isFinite(Number(item.coQty)) ? Number(item.coQty) : 0,
      cumSewInQty: Number.isFinite(Number(item.cumSewInQty)) ? Number(item.cumSewInQty) : 0,
      cumSewOutQty: Number.isFinite(Number(item.cumSewOutQty)) ? Number(item.cumSewOutQty) : 0,
      cumSewOutRejQty: Number.isFinite(Number(item.cumSewOutRejQty)) ? Number(item.cumSewOutRejQty) : 0,
      cumCTNQty: Number.isFinite(Number(item.cumCTNQty)) ? Number(item.cumCTNQty) : 0,
      statusText: String(item.statusText || ''),
      deliveredQty: Number.isFinite(Number(item.deliveredQty)) ? Number(item.deliveredQty) : 0,
      cifOffering: String(item.cifOffering || ''),
      remark: String(item.remark || ''),
    }));

    const now = Date.now();
    const uploadId = `${now}_${Math.random().toString(36).slice(2, 7)}`;

    // 1. Immediately update local state, timestamp ref, and localStorage
    updateDashboardData(cleanedData, now, 'upload', uploadId);

    // 2. Send to Express Server API for local SSE broadcast
    fetch('/api/dashboard', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ data: cleanedData, lastUpdated: now, uploadId })
    }).catch(() => {});

    // 3. Persist to Firestore using Gzip single-doc compression (~90% fewer Read/Write units!),
    //    falling back to safe 250-item chunks only if compressed payload exceeds 750KB.
    try {
      const compressedGzip = await compressOrdersToBase64(cleanedData);
      if (compressedGzip && compressedGzip.length < 750000) {
        await setDoc(doc(db, 'dashboard_v3', 'main'), {
          compressedGzip,
          format: 'gzip_v1',
          chunkCount: 1,
          totalCount: cleanedData.length,
          lastUpdated: now,
          uploadId,
          serverUpdatedAt: serverTimestamp(),
        });
      } else {
        const CHUNK_SIZE = 250;
        const chunks: ProductionOrder[][] = [];
        for (let i = 0; i < cleanedData.length; i += CHUNK_SIZE) {
          chunks.push(cleanedData.slice(i, i + CHUNK_SIZE));
        }

        let syncedRows = 0;

        // Write all part_1..part_N documents FIRST in parallel batches of 4
        if (chunks.length > 1) {
          for (let batchStart = 1; batchStart < chunks.length; batchStart += 4) {
            const batchIndices = [];
            for (let j = batchStart; j < Math.min(batchStart + 4, chunks.length); j++) {
              batchIndices.push(j);
            }
            await Promise.all(
              batchIndices.map(async (idx) => {
                await setDoc(doc(db, 'dashboard_v3', `part_${idx}`), {
                  orders: chunks[idx],
                  uploadId
                });
                syncedRows += chunks[idx].length;
              })
            );
            setUploadProgress({ current: syncedRows, total: cleanedData.length });
          }
        }

        // Write main document LAST to atomically notify all connected PCs
        await setDoc(doc(db, 'dashboard_v3', 'main'), {
          orders: chunks[0] || [],
          chunkCount: chunks.length,
          totalCount: cleanedData.length,
          lastUpdated: now,
          uploadId: uploadId,
          serverUpdatedAt: serverTimestamp(),
        });
      }

      setUploadProgress({ current: cleanedData.length, total: cleanedData.length });
      setSyncStatus('synced');
      setSyncError(null);
      setLoadingState('idle');
    } catch (err: any) {
      const msg = err?.message || String(err);
      console.error('Firestore Upload Sync Error:', msg);
      setSyncStatus('error');
      const friendlyMsg = isFirestoreQuotaError(msg)
        ? 'Firestore daily quota exceeded! Other PCs cannot receive this upload until quota resets.'
        : `Cloud sync upload failed (${msg}). Other PCs may not see this update.`;
      setSyncError(friendlyMsg);
      setError(friendlyMsg);
      setLoadingState('idle');
    }
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      processFile(file);
    }
    e.target.value = '';
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
            setLoadingState('idle');
            return;
          }

          const windowBounds = getSixWeeksWindow();
          
          const excludedTerms = [
            'SIZE SET', 'BLACK SEAL', 'SAMPLES_PRESETTING', 'SAMPLES_PP', 'PP_SAMPLE',
            'MTL SAMPLE', 'PP SAMPLE', 'PP SAMPLE PRNT', 'PRE SETTING', 'SAMPLE PP',
            'WASH & TOP', 'PPZ', 'TC-PP', 'TC-PPZ', 'MTL', 'TLT'
          ];
          const excludedSet = new Set(excludedTerms);

          const hasWarehouseColumn = jsonData.some((row: any) =>
            getVal(row, ['Prod Warehouse', 'prod warehouse', 'Warehouse', 'PROD WAREHOUSE', 'WH', 'Prod WH']) !== undefined
          );

          const isRowExcluded = (row: any) => {
            return Object.values(row).some(val => {
              const strVal = String(val).trim().toUpperCase();
              if (excludedSet.has(strVal)) return true;
              if (/^VPO(_|\d)/.test(strVal)) return true;
              return false;
            });
          };

          const isRowERK = (row: any) => {
            if (!hasWarehouseColumn) return true;
            const warehouse = String(getVal(row, ['Prod Warehouse', 'prod warehouse', 'Warehouse', 'PROD WAREHOUSE', 'WH', 'Prod WH']) || '').trim().toUpperCase();
            return !warehouse || warehouse === 'ERK' || warehouse.includes('ERK');
          };

          // Primary filter: 6-week window + ERK warehouse
          let targetRows = jsonData.filter((row: any) => {
            if (isRowExcluded(row)) return false;
            if (!isRowERK(row)) return false;
            const rawPlanDelDate = getVal(row, ['Plan Del Date', 'Plan Del Date ', 'plan del date']);
            const parsedDelDate = parseExcelDate(rawPlanDelDate);
            const weekNoStr = String(getVal(row, ['WEEK NO', 'Week No', 'week no']) || '').trim();
            return isDateWithinSixWeeks(parsedDelDate, weekNoStr, windowBounds);
          });

          // Fallback 1: If 0 items matched 6-week window, keep all non-excluded ERK rows
          if (targetRows.length === 0) {
            targetRows = jsonData.filter((row: any) => !isRowExcluded(row) && isRowERK(row));
          }

          // Fallback 2: Keep all non-excluded rows from file
          if (targetRows.length === 0) {
            targetRows = jsonData.filter((row: any) => !isRowExcluded(row));
          }

          if (targetRows.length === 0) {
            setError("No valid production orders found in the uploaded file.");
            setLoadingState('idle');
            return;
          }

          const excelExtractedRemarks: Record<string, { text: string; updatedAt: number }> = {};
          const excelExtractedCifOfferings: Record<string, { text: string; updatedAt: number }> = {};

          const parsedData: ProductionOrder[] = targetRows.map((row: any, index: number) => {
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
            
            const legacyIdStr = `${vpoStr}_${schedStr}_${styleStr}_${colorStr}_${destStr}_${planDelDate}`.replace(/[^a-zA-Z0-9_-]/g, '-');
            const stableId = `${vpoStr}_${schedStr}_${styleStr}_${colorStr}_${destStr}`.replace(/[^a-zA-Z0-9_-]/g, '-');
            const finalId = (stableId === '_____' || !stableId) ? `row_${index}` : stableId;

            // Expanded Remark column search
            const rowRemark = String(getVal(row, ['Remark', 'remark', 'REMARK', 'Remarks', 'REMARKS', 'Comment', 'Comments', 'Note', 'Notes', 'Production Remark', 'Production Remarks', 'PROD REMARK', 'PROD REMARKS']) || '').trim();
            
            if (rowRemark) {
              excelExtractedRemarks[finalId] = { text: rowRemark, updatedAt: Date.now() };
              if (legacyIdStr) {
                excelExtractedRemarks[legacyIdStr] = { text: rowRemark, updatedAt: Date.now() };
              }
            }

            const rowCifOffering = String(getVal(row, ['CIF Offering', 'cif offering', 'CIF OFFERING', 'CIF', 'Cif Offering', 'CIF OFFERINGS', 'Cif']) || '').trim();
            if (rowCifOffering) {
              excelExtractedCifOfferings[finalId] = { text: rowCifOffering, updatedAt: Date.now() };
              if (legacyIdStr) {
                excelExtractedCifOfferings[legacyIdStr] = { text: rowCifOffering, updatedAt: Date.now() };
              }
            }

            const rawPED = getVal(row, ['PED', 'ped', 'Ped', 'P.E.D.', 'P.E.D', 'PED Date', 'PED DATE', 'PED_DATE', 'PED-DATE', 'Plan Ex Del Date', 'Plan Ex-Del Date', 'PLAN EX DEL DATE', 'Plan Ex-Factory Date', 'Plan Ex Factory Date', 'Ex Del Date', 'Ex-Del Date', 'Plan Ex Del', 'PLAN EX DEL', 'AH']);
            const parsedPEDDate = parseExcelDate(rawPED);
            let pedStr = '';
            if (parsedPEDDate) {
              const y = parsedPEDDate.getFullYear();
              const m = String(parsedPEDDate.getMonth() + 1).padStart(2, '0');
              const d = String(parsedPEDDate.getDate()).padStart(2, '0');
              pedStr = `${y}/${m}/${d}`;
            } else {
              const strVal = String(rawPED || '').trim();
              if (/^\d{8}$/.test(strVal)) {
                pedStr = `${strVal.substring(0, 4)}/${strVal.substring(4, 6)}/${strVal.substring(6, 8)}`;
              } else {
                pedStr = strVal;
              }
            }

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
              ped: pedStr,
              weekNo: weekNo,
              coQty: coQty,
              cumSewInQty: getNum(row, ['Cum Sew In Qty', 'Cum SewIn Qty']),
              cumSewOutQty: cumSewOutQty,
              cumSewOutRejQty: getNum(row, ['Cum Sew Out Rej Qty', 'Cum SewOut Rej Qty']),
              cumCTNQty: getNum(row, ['Cum CTN Qty', 'CUM CTN QTY', 'Cum Ctn Qty']),
              statusText: statusText,
              deliveredQty: deliveredQty,
              cifOffering: rowCifOffering,
              remark: rowRemark // Store it in the object as well
            };
          });

          if (Object.keys(excelExtractedCifOfferings).length > 0) {
            setCifOfferings(prev => {
              const merged = { ...excelExtractedCifOfferings, ...prev };
              try {
                localStorage.setItem('production_dashboard_cif_offerings', JSON.stringify(merged));
              } catch (e) {}
              return merged;
            });

            const syncExtractedCif = async () => {
              try {
                const cifRef = doc(db, 'cif_offerings_v2', 'all');
                await setDoc(cifRef, excelExtractedCifOfferings, { merge: true });
              } catch (err) {
                console.error("Batch CIF offerings sync failed", err);
              }
            };
            syncExtractedCif();
          }

          if (Object.keys(excelExtractedRemarks).length > 0) {
            setRemarks(prev => {
              const merged = { ...excelExtractedRemarks, ...prev };
              try {
                localStorage.setItem('production_dashboard_remarks', JSON.stringify(merged));
              } catch (e) {}
              return merged;
            });

            // Sync extracted remarks to Firestore in background (Single-Doc Optimized)
            const syncExtracted = async () => {
              try {
                const remarksRef = doc(db, 'remarks_v2', 'all');
                await setDoc(remarksRef, excelExtractedRemarks, { merge: true });
              } catch (err) {
                console.error("Batch remark sync failed", err);
              }
            };
            syncExtracted();
          }

          parsedData.sort((a, b) => a.planDelDate.localeCompare(b.planDelDate));

          persistDashboardData(parsedData);
          
          setSearchTerm('');
          setFilterBuyer([]);
          setFilterWeekNo([]);
          setFilterStatus([]);
          setFilterShipmentMode([]);
          setFilterDestination([]);
          setFilterPackMethod([]);
          setFilterRemark([]);
          setFilterPED([]);

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

  const dragCounterRef = useRef(0);

  useEffect(() => {
    const handleWindowDrag = (e: DragEvent) => e.preventDefault();
    window.addEventListener('dragover', handleWindowDrag);
    window.addEventListener('drop', handleWindowDrag);
    return () => {
      window.removeEventListener('dragover', handleWindowDrag);
      window.removeEventListener('drop', handleWindowDrag);
    };
  }, []);

  const handleDragEnter = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    dragCounterRef.current += 1;
    if (e.dataTransfer.items && e.dataTransfer.items.length > 0) {
      setIsDragging(true);
    }
  }, []);

  const handleDragOver = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    e.dataTransfer.dropEffect = 'copy';
    if (!isDragging) setIsDragging(true);
  }, [isDragging]);

  const handleDragLeave = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    dragCounterRef.current -= 1;
    if (dragCounterRef.current <= 0) {
      dragCounterRef.current = 0;
      setIsDragging(false);
    }
  }, []);

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    dragCounterRef.current = 0;
    setIsDragging(false);
    const file = e.dataTransfer.files?.[0];
    if (file) {
      const fileNameLower = file.name.toLowerCase();
      if (
        fileNameLower.endsWith('.xlsx') || 
        fileNameLower.endsWith('.xls') || 
        fileNameLower.endsWith('.csv') ||
        file.type.includes('spreadsheet') ||
        file.type.includes('excel') ||
        file.type.includes('csv')
      ) {
        processFile(file);
      } else {
        setError("Please upload a valid Excel (.xlsx, .xls) or CSV file.");
      }
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

      const itemRemarkText = getRowRemark(item, remarks, normalizedRemarksMap).text.trim();
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

      const itemCifText = getRowCifOffering(item, cifOfferings, normalizedCifOfferingsMap).text.trim();
      const normalizedItemCif = itemCifText.toUpperCase();
      const matchesCifOffering = filterCifOffering.length === 0 || filterCifOffering.some(c => c === normalizedItemCif || (c === '(Empty)' && normalizedItemCif === ''));

      const pedVal = (item.ped && item.ped.trim()) ? item.ped.trim() : '(Empty)';
      const matchesPED = filterPED.length === 0 || filterPED.includes(pedVal);

      return matchesSearch && matchesBuyer && matchesWeek && matchesStatus && matchesShipmentMode && matchesDestination && matchesPackMethod && matchesRemark && matchesCifOffering && matchesPED;
    });
  }, [data, searchTerm, filterBuyer, filterWeekNo, filterStatus, filterShipmentMode, filterDestination, filterPackMethod, filterRemark, filterCifOffering, filterPED, remarks, cifOfferings, sixWeekWindow]);

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

  const uniqueCifOfferingsList = useMemo(() => {
    const suggestions = new Set<string>();
    Object.values(cifOfferings).forEach((r: { text: string; updatedAt: number | null }) => {
      if (r?.text && r.text.trim()) suggestions.add(r.text.trim().toUpperCase());
    });
    return Array.from(suggestions).sort();
  }, [cifOfferings]);

  // Helper to get unique options based on current filters (excluding the filter itself)
  const getUniqueOptions = useCallback((field: keyof ProductionOrder | 'remark' | 'cifOffering' | 'ped') => {
    if (!data) return [];
    const options = new Set<string>();
    
    data.forEach(item => {
      const itemRemarkText = getRowRemark(item, remarks, normalizedRemarksMap).text.trim();
      const itemCifText = getRowCifOffering(item, cifOfferings, normalizedCifOfferingsMap).text.trim();

      const matchesSearch = searchTerm === '' || 
        Object.values(item).some(val => 
          String(val).toLowerCase().includes(searchTerm.toLowerCase())
        ) || itemRemarkText.toLowerCase().includes(searchTerm.toLowerCase())
        || itemCifText.toLowerCase().includes(searchTerm.toLowerCase());
      
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

      const normalizedItemCif = itemCifText.toUpperCase();
      const matchesCifOffering = field === 'cifOffering' || filterCifOffering.length === 0 || filterCifOffering.some(c => c === normalizedItemCif || (c === '(Empty)' && normalizedItemCif === ''));

      const pedVal = (item.ped && item.ped.trim()) ? item.ped.trim() : '(Empty)';
      const matchesPED = field === 'ped' || filterPED.length === 0 || filterPED.includes(pedVal);

      if (matchesSearch && matchesBuyer && matchesWeek && matchesStatus && matchesShipmentMode && matchesDestination && matchesPackMethod && matchesRemark && matchesCifOffering && matchesPED) {
        if (field === 'remark') {
          options.add(normalizedItemRemark === '' ? '(Empty)' : normalizedItemRemark);
        } else if (field === 'cifOffering') {
          options.add(normalizedItemCif === '' ? '(Empty)' : normalizedItemCif);
        } else if (field === 'ped') {
          options.add(pedVal);
        } else {
          options.add(String(item[field as keyof ProductionOrder]));
        }
      }
    });
    
    return Array.from(options).filter(Boolean);
  }, [data, searchTerm, filterBuyer, filterWeekNo, filterStatus, filterShipmentMode, filterDestination, filterPackMethod, filterRemark, filterCifOffering, filterPED, remarks, cifOfferings]);

  // Unique values for dropdowns (strictly restricted to the 6-week window)
  const uniqueBuyers = useMemo(() => Array.from(new Set([...getUniqueOptions('buyer'), ...filterBuyer])).sort(), [getUniqueOptions, filterBuyer]);
  const uniqueWeeks = useMemo(() => Array.from(new Set([...getUniqueOptions('weekNo'), ...filterWeekNo])).filter(w => sixWeekWindow.weekNumbers.has(w)).sort((a,b) => Number(a) - Number(b)), [getUniqueOptions, filterWeekNo, sixWeekWindow]);
  const uniqueShipmentModes = useMemo(() => Array.from(new Set([...getUniqueOptions('shipmentMode'), ...filterShipmentMode])).sort(), [getUniqueOptions, filterShipmentMode]);
  const uniqueDestinations = useMemo(() => Array.from(new Set([...getUniqueOptions('destination'), ...filterDestination])).sort(), [getUniqueOptions, filterDestination]);
  const uniquePackMethods = useMemo(() => Array.from(new Set([...getUniqueOptions('packMethod'), ...filterPackMethod])).sort(), [getUniqueOptions, filterPackMethod]);
  const uniqueCifOfferings = useMemo(() => Array.from(new Set([...getUniqueOptions('cifOffering'), ...filterCifOffering])).sort(), [getUniqueOptions, filterCifOffering]);
  const uniqueRemarks = useMemo(() => Array.from(new Set([...getUniqueOptions('remark'), ...filterRemark])).sort(), [getUniqueOptions, filterRemark]);
  const uniquePEDs = useMemo(() => Array.from(new Set([...getUniqueOptions('ped'), ...filterPED])).sort(), [getUniqueOptions, filterPED]);

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

  const handleExportExcel = useCallback(() => {
    if (filteredItems.length === 0) return;

    const exportData = filteredItems.map(item => ({
      'Plan Del Date': item.planDelDate,
      'PED': item.ped || '',
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
      'Remark': getRowRemark(item, remarks, normalizedRemarksMap).text,
      'CIF Offering': getRowCifOffering(item, cifOfferings, normalizedCifOfferingsMap).text,
      'Production Lines': getRowProductionLine(item, productionLines).text,
      'Delivered Qty': Number(item.deliveredQty) || 0
    }));

    const worksheet = XLSX.utils.json_to_sheet(exportData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Production Data");
    
    const colWidths = [
      { wch: 15 }, { wch: 15 }, { wch: 10 }, { wch: 15 }, { wch: 15 }, { wch: 15 }, 
      { wch: 15 }, { wch: 12 }, { wch: 20 }, { wch: 15 }, { wch: 15 }, 
      { wch: 15 }, { wch: 10 }, { wch: 15 }, { wch: 15 }, { wch: 15 }, 
      { wch: 15 }, { wch: 30 }, { wch: 25 }, { wch: 20 }, { wch: 15 }
    ];
    worksheet['!cols'] = colWidths;

    XLSX.writeFile(workbook, `Production_Data_${new Date().toISOString().split('T')[0]}.xlsx`);
  }, [filteredItems, remarks, cifOfferings]);

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
                className={`block border-2 border-dashed rounded-2xl p-8 transition-colors cursor-pointer select-none ${
                  isDragging ? 'border-indigo-500 bg-indigo-50/70' : 'border-slate-300 hover:border-indigo-400 hover:bg-slate-50'
                }`}
                onClick={() => fileInputRef.current?.click()}
                onDragOver={(e) => handleDragOver(e)}
                onDragEnter={(e) => handleDragEnter(e)}
                onDragLeave={(e) => handleDragLeave(e)}
                onDrop={(e) => handleDrop(e)}
              >
                <UploadCloud className={`w-8 h-8 mx-auto mb-3 pointer-events-none transition-colors ${isDragging ? 'text-indigo-500' : 'text-slate-400'}`} />
                <p className="text-sm font-medium text-slate-700 mb-1 pointer-events-none">Click to upload or drag and drop</p>
                <p className="text-xs text-slate-500 pointer-events-none">XLSX, XLS, or CSV files</p>
              </div>

              <input 
                id="main-file-upload"
                type="file" 
                className="hidden" 
                ref={fileInputRef} 
                accept=".xlsx,.xls,.csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel,text/csv" 
                onChange={handleFileUpload} 
              />
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
        id="dashboard-header-file-upload"
        type="file" 
        className="hidden" 
        ref={fileInputRef} 
        accept=".xlsx,.xls,.csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel,text/csv" 
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
            <div className="text-slate-500 mt-1 text-xs md:text-sm flex flex-col sm:flex-row sm:items-center gap-1 sm:gap-2">
              <span>Item-level fulfillment details for the next 6 weeks.</span>
              {lastUpdated && (
                <span className="text-slate-600 font-medium flex items-center gap-1.5 truncate">
                  <span className="hidden sm:inline text-slate-300">•</span>
                  <CheckCircle className="w-3.5 h-3.5 text-emerald-500 shrink-0" /> 
                  OrderBook Updated: {new Date(lastUpdated).toLocaleString('en-US', { dateStyle: 'medium', timeStyle: 'short' })}
                </span>
              )}
              <span className="font-medium flex items-center gap-1.5 truncate">
                <span className="hidden sm:inline text-slate-300">•</span>
                {syncStatus === 'error' ? (
                  <span className="text-rose-600 flex items-center gap-1.5">
                    <span className="w-2 h-2 bg-rose-500 rounded-full inline-block"></span>
                    Cloud Sync Error
                  </span>
                ) : syncStatus === 'syncing' || syncStatus === 'connecting' ? (
                  <span className="text-indigo-600 flex items-center gap-1.5">
                    <Loader2 className="w-3 h-3 animate-spin inline-block" />
                    {syncStatus === 'connecting' ? 'Connecting Cloud...' : 'Syncing Cloud...'}
                  </span>
                ) : (
                  <span className="text-slate-400 flex items-center gap-1.5">
                    <span className="w-2 h-2 bg-emerald-500 rounded-full animate-pulse inline-block"></span>
                    Cloud Sync Active
                  </span>
                )}
                <button
                  onClick={handleManualCloudSync}
                  disabled={syncStatus === 'syncing'}
                  title="Force refresh latest Orderbook from Cloud"
                  className="ml-1 inline-flex items-center gap-1 px-1.5 py-0.5 text-[11px] font-medium text-slate-600 hover:text-indigo-600 bg-slate-100 hover:bg-indigo-50 rounded border border-slate-200 transition-colors disabled:opacity-50"
                >
                  <RefreshCw className={`w-3 h-3 ${syncStatus === 'syncing' ? 'animate-spin' : ''}`} />
                  Sync
                </button>
              </span>
            </div>
            {(syncError || error) && (
              <div className="mt-2 px-3 py-2 bg-rose-50 border border-rose-200 rounded-lg flex items-center gap-2 text-xs text-rose-700">
                <AlertCircle className="w-4 h-4 text-rose-500 shrink-0" />
                <span>{syncError || error}</span>
              </div>
            )}
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
                label="CIF Offerings"
                options={uniqueCifOfferings}
                selectedValues={filterCifOffering}
                onChange={setFilterCifOffering}
              />

              <MultiSelectDropdown 
                label="Remarks"
                options={uniqueRemarks}
                selectedValues={filterRemark}
                onChange={setFilterRemark}
              />

              <MultiSelectDropdown 
                label="PED Dates"
                options={uniquePEDs}
                selectedValues={filterPED}
                onChange={setFilterPED}
              />
            </div>
            
            <div className="grid grid-cols-3 sm:flex gap-2 w-full sm:w-auto">
              <button 
                onClick={() => {
                  setSearchTerm(''); setFilterBuyer([]); setFilterWeekNo([]); setFilterStatus([]); setFilterShipmentMode([]); setFilterDestination([]); setFilterPackMethod([]); setFilterCifOffering([]); setFilterRemark([]); setFilterPED([]);
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
              <table className="w-full min-w-[2490px] table-fixed text-sm text-left border-collapse relative">
                <colgroup>
                  {TABLE_COLUMN_WIDTHS.map((width, idx) => (
                    <col key={idx} style={{ width }} />
                  ))}
                </colgroup>
                <thead className="text-[11px] uppercase tracking-wider text-slate-500 font-semibold bg-slate-50/90 backdrop-blur-sm border-b border-slate-200 sticky top-0 z-20 shadow-sm">
                  <tr className="h-[44px]">
                    <th className="px-3 py-3 border border-slate-200 whitespace-nowrap truncate">Plan Del Date</th>
                    <th className="px-3 py-3 border border-slate-200 whitespace-nowrap text-center truncate">PED</th>
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
                    <th className="px-3 py-3 border border-slate-200 whitespace-nowrap truncate">CIF Offering</th>
                    <th className="px-3 py-3 border border-slate-200 whitespace-nowrap truncate">Production Lines</th>
                    <th className="px-3 py-3 border border-slate-200 whitespace-nowrap text-right truncate">Delivered Qty</th>
                  </tr>
                </thead>
                <tbody>
                  {paddingTop > 0 && (
                    <tr className="border-0 p-0 m-0">
                      <td colSpan={21} style={{ height: `${paddingTop}px`, padding: 0, border: 0, margin: 0, lineHeight: 0 }} />
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
                          <td className="px-3 py-2 border-r border-b border-slate-200 text-slate-700 whitespace-nowrap text-center font-medium bg-slate-50/50 truncate">{row.ped || '-'}</td>
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
                              initialValue={getRowRemark(row, remarks, normalizedRemarksMap).text}
                              updatedAt={getRowRemark(row, remarks, normalizedRemarksMap).updatedAt}
                              rowId={row.id}
                              onSave={handleRemarkChange}
                            />
                          </td>
                          <td className="px-3 py-2 border-r border-b border-slate-200 whitespace-nowrap truncate">
                            <CifOfferingInput 
                              initialValue={getRowCifOffering(row, cifOfferings, normalizedCifOfferingsMap).text}
                              updatedAt={getRowCifOffering(row, cifOfferings, normalizedCifOfferingsMap).updatedAt}
                              rowId={row.id}
                              onSave={handleCifOfferingChange}
                            />
                          </td>
                          <td className="px-3 py-2 border-r border-b border-slate-200 whitespace-nowrap truncate">
                            <ProductionLineInput 
                              initialValue={getRowProductionLine(row, productionLines).text}
                              rowId={row.id}
                              onSave={handleProductionLineChange}
                            />
                          </td>
                          <td className="px-3 py-2 border-b border-slate-200 text-right font-medium text-slate-900 whitespace-nowrap truncate">{row.deliveredQty?.toLocaleString() || "0"}</td>
                        </tr>
                      );
                    })
                  ) : (
                    <tr>
                      <td colSpan={21} className="px-6 py-12 text-center text-slate-500 bg-slate-50">
                        No orders found matching the current filters.
                      </td>
                    </tr>
                  )}
                  {paddingBottom > 0 && (
                    <tr className="border-0 p-0 m-0">
                      <td colSpan={21} style={{ height: `${paddingBottom}px`, padding: 0, border: 0, margin: 0, lineHeight: 0 }} />
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
                            <span className="text-slate-400">PED:</span>
                            <span className="font-semibold text-slate-700">{row.ped || '-'}</span>
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

                        <div className="mt-auto space-y-1.5">
                          <RemarkInput 
                            initialValue={getRowRemark(row, remarks, normalizedRemarksMap).text}
                            updatedAt={getRowRemark(row, remarks, normalizedRemarksMap).updatedAt}
                            rowId={row.id}
                            onSave={handleRemarkChange}
                          />
                          <CifOfferingInput 
                            initialValue={getRowCifOffering(row, cifOfferings, normalizedCifOfferingsMap).text}
                            updatedAt={getRowCifOffering(row, cifOfferings, normalizedCifOfferingsMap).updatedAt}
                            rowId={row.id}
                            onSave={handleCifOfferingChange}
                          />
                          <ProductionLineInput 
                            initialValue={getRowProductionLine(row, productionLines).text}
                            rowId={row.id}
                            onSave={handleProductionLineChange}
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
        cifOfferings={cifOfferings}
        productionLines={productionLines}
        activeFilters={{
          search: searchTerm,
          buyers: filterBuyer,
          weeks: filterWeekNo,
          statuses: filterStatus,
          shipmentModes: filterShipmentMode,
          destinations: filterDestination,
          packMethods: filterPackMethod,
          cifOfferingsFilter: filterCifOffering,
          remarksFilter: filterRemark,
          peds: filterPED
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

      <div className="mt-8 text-[11px] text-slate-400 text-center pb-4 flex flex-col items-center gap-2">
        <div className="flex items-center gap-2 flex-wrap justify-center font-medium">
          <span className="inline-flex items-center px-2 py-0.5 rounded-full bg-emerald-50 text-emerald-700 border border-emerald-200">
            <span className="w-1.5 h-1.5 rounded-full bg-emerald-500 mr-1.5 animate-pulse"></span>
            {Object.keys(remarks).length} Active Remarks Synced
          </span>
          <span className="text-slate-300">•</span>
          <span>Window: {sixWeekWindow.start.toLocaleDateString()} - {sixWeekWindow.end.toLocaleDateString()}</span>
        </div>
        
        {recoveryMessage && (
          <div className="text-xs font-semibold px-3 py-1 rounded-lg bg-indigo-50 text-indigo-700 border border-indigo-200 animate-fade-in">
            {recoveryMessage}
          </div>
        )}

        <button 
          disabled={isRecovering}
          onClick={async () => {
            setIsRecovering(true);
            setRecoveryMessage("Restoring remarks and orders from server database...");
            try {
              sessionStorage.removeItem('migration_done');
              
              // 1. Fetch remarks from server backup
              const remarksRes = await fetch('/api/remarks');
              const serverRemarks = await safeJson(remarksRes);
              
              if (serverRemarks && typeof serverRemarks === 'object') {
                const count = Object.keys(serverRemarks).length;
                setRemarks(prev => {
                  const merged = { ...serverRemarks, ...prev };
                  try {
                    localStorage.setItem('production_dashboard_remarks', JSON.stringify(merged));
                  } catch (e) {}
                  return merged;
                });

                // 2. Sync to Firestore remarks_v2/all
                const remarksRef = doc(db, 'remarks_v2', 'all');
                await setDoc(remarksRef, serverRemarks, { merge: true });
                setRecoveryMessage(`Sync Complete! Successfully restored ${count} remarks across active weeks.`);
              }

              // 3. Fetch dashboard if needed
              const dashRes = await fetch('/api/dashboard');
              const serverDash = await safeJson(dashRes);
              if (serverDash?.data && Array.isArray(serverDash.data) && serverDash.data.length > 0) {
                updateDashboardData(serverDash.data, serverDash.lastUpdated || Date.now(), 'manual_sync', serverDash.uploadId);
              }

              setTimeout(() => setRecoveryMessage(null), 6000);
            } catch (err: any) {
              setRecoveryMessage(`Recovery warning: ${err.message || 'Check connection'}`);
              setTimeout(() => setRecoveryMessage(null), 5000);
            } finally {
              setIsRecovering(false);
            }
          }}
          className="px-3 py-1 border border-slate-300 hover:border-indigo-400 bg-white hover:bg-slate-50 text-slate-600 rounded-lg text-xs font-medium shadow-sm transition-all disabled:opacity-50"
        >
          {isRecovering ? "Syncing..." : "Force Sync Recovery"}
        </button>
      </div>
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
