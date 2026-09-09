import React, { useState, useRef, useCallback, useMemo, useEffect } from 'react';
import { UploadCloud, FileSpreadsheet, AlertCircle, Search, Package, CheckCircle, TrendingUp, AlertTriangle, Loader2, Printer, BarChart2, Table } from 'lucide-react';
import * as XLSX from 'xlsx';
import { useVirtualizer } from '@tanstack/react-virtual';
import { db } from '../firebase';
import { doc, setDoc, onSnapshot, writeBatch, collection, getDocs } from 'firebase/firestore';
import { KPIView } from './KPIView';

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
        className="w-full flex items-center justify-between py-2 pl-3 pr-3 bg-slate-50 hover:bg-slate-100/50 border border-slate-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 text-slate-700 text-left min-w-[140px] transition-all"
      >
        <span className="truncate pr-2 font-medium">{displayText}</span>
        <svg className={`w-4 h-4 text-slate-400 transition-transform flex-shrink-0 ${isOpen ? 'rotate-180' : ''}`} fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M19 9l-7 7-7-7"></path></svg>
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
  orderToShippedPct: number;
  remark?: string;
}

function RemarkInput({
  initialValue,
  rowId,
  onSave
}: {
  initialValue: string;
  rowId: string;
  onSave: (id: string, text: string) => void;
}) {
  const [value, setValue] = useState(initialValue);

  useEffect(() => {
    setValue(initialValue);
  }, [initialValue]);

  return (
    <input 
      type="text" 
      value={value}
      onChange={(e) => setValue(e.target.value)}
      onBlur={() => {
        if (value !== initialValue) {
          onSave(rowId, value);
        }
      }}
      onKeyDown={(e) => {
        if (e.key === 'Enter') {
          e.currentTarget.blur();
        }
      }}
      placeholder="Add remark..."
      list="remark-suggestions"
      className="w-full text-sm bg-transparent border-0 border-b border-transparent hover:border-slate-300 focus:border-indigo-500 focus:ring-0 px-1 py-1 transition-colors"
    />
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

export function Dashboard() {
  const [data, setData] = useState<ProductionOrder[] | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const [loadingState, setLoadingState] = useState<'idle' | 'reading' | 'parsing' | 'uploading'>('idle');
  const [uploadProgress, setUploadProgress] = useState({ current: 0, total: 0 });
  const [loadingInitial, setLoadingInitial] = useState(true);
  const [remarks, setRemarks] = useState<Record<string, string>>({});
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [error, setError] = useState<string | null>(null);
  const [lastUpdated, setLastUpdated] = useState<number | null>(null);
  const [viewMode, setViewMode] = useState<'table' | 'kpi'>('table');

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

  useEffect(() => {
    const unsubRemarks = onSnapshot(collection(db, "remarks"), (snapshot) => {
      const newRemarks: Record<string, string> = {};
      snapshot.forEach(doc => {
        newRemarks[doc.id] = doc.data().text || '';
      });
      setRemarks(newRemarks);
    }, (error) => {
      console.error("Remarks snapshot error:", error);
    });

    const unsub = onSnapshot(doc(db, "dashboardData", "latest"), async (docSnap) => {
      if (docSnap.exists()) {
        const { currentUploadId, timestamp } = docSnap.data();
        if (timestamp) {
          setLastUpdated(timestamp);
        } else {
          setLastUpdated(null);
        }
        if (currentUploadId) {
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
                 loadedData.push(order);
               }
            });
            loadedData.sort((a, b) => a.planDelDate.localeCompare(b.planDelDate));
            setData(loadedData);
          } catch(e) {
            console.error("Error fetching data", e);
          }
        } else {
           setData(null);
        }
      } else {
        setData(null);
      }
      setLoadingInitial(false);
    }, (error) => {
      console.error("Snapshot error:", error);
      setLoadingInitial(false);
    });
    return () => {
      unsub();
      unsubRemarks();
    };
  }, []);

  const handleRemarkChange = async (id: string, text: string) => {
    try {
      await setDoc(doc(db, "remarks", id), { text, updatedAt: Date.now() }, { merge: true });
    } catch (e) {
      console.error("Failed to save remark", e);
    }
  };

  const uploadToFirestore = async (parsedData: ProductionOrder[]) => {
    setLoadingState('uploading');
    setUploadProgress({ current: 0, total: parsedData.length });
    try {
      const uploadId = Date.now().toString();
      const batchSize = 400; // max 500 operations per batch
      
      let processed = 0;
      for (let i = 0; i < parsedData.length; i += batchSize) {
        const chunk = parsedData.slice(i, i + batchSize);
        const batch = writeBatch(db);
        
        chunk.forEach(order => {
          const orderRef = doc(db, `uploads/${uploadId}/orders`, order.id);
          batch.set(orderRef, order);
        });
        
        await batch.commit();
        processed += chunk.length;
        setUploadProgress({ current: processed, total: parsedData.length });
      }

      await setDoc(doc(db, "dashboardData", "latest"), {
        currentUploadId: uploadId,
        timestamp: Date.now()
      });
      
    } catch (err) {
      console.error(err);
      alert("Failed to upload data to database.");
      setLoadingState('idle');
    } finally {
      setTimeout(() => {
        setLoadingState('idle');
        setUploadProgress({ current: 0, total: 0 });
      }, 1000);
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

          const getVal = (row: any, searchKeys: string[]) => {
            const normKeys = searchKeys.map(k => k.toLowerCase().replace(/\s+/g, ''));
            for (const key of Object.keys(row)) {
              const normKey = key.toLowerCase().replace(/\s+/g, '');
              if (normKeys.includes(normKey)) {
                return row[key];
              }
            }
            return undefined;
          };

          const getNum = (row: any, searchKeys: string[]) => {
            const val = getVal(row, searchKeys);
            if (val == null || val === '') return 0;
            if (typeof val === 'number') return val;
            const parsed = Number(String(val).replace(/,/g, ''));
            return isNaN(parsed) ? 0 : parsed;
          };

          const today = new Date();
          today.setHours(0, 0, 0, 0);
          
          const dayOfWeek = today.getDay();
          const diffToMonday = today.getDate() - dayOfWeek + (dayOfWeek === 0 ? -6 : 1);
          const startOfCurrentWeek = new Date(today);
          startOfCurrentWeek.setDate(diffToMonday);

          const endOf6Weeks = new Date(startOfCurrentWeek);
          endOf6Weeks.setDate(startOfCurrentWeek.getDate() + 42);
          
          const excludedTerms = [
            'SIZE SET', 'BLACK SEAL', 'SAMPLES_PRESETTING', 'SAMPLES_PP', 'PP_SAMPLE',
            'MTL SAMPLE', 'PP SAMPLE', 'PP SAMPLE PRNT', 'PRE SETTING', 'SAMPLE PP',
            'WASH & TOP', 'PPZ', 'TC-PP', 'TC-PPZ', 'MTL', 'TLT'
          ];

          const filteredData = jsonData.filter((row: any) => {
            const warehouse = String(row['Prod Warehouse'] || '').trim().toUpperCase();
            const isERK = warehouse === 'ERK';
            
            const rawPlanDelDate = String(row['Plan Del Date'] || row['Plan Del Date '] || '').trim();
            let isWithin6Weeks = false;
            
            if (rawPlanDelDate.length === 8) {
              const year = parseInt(rawPlanDelDate.substring(0, 4));
              const month = parseInt(rawPlanDelDate.substring(4, 6)) - 1;
              const dayNum = parseInt(rawPlanDelDate.substring(6, 8));
              const rowDate = new Date(year, month, dayNum);
              
              isWithin6Weeks = rowDate >= startOfCurrentWeek && rowDate < endOf6Weeks;
            }
            
            const hasExcludedTerm = Object.values(row).some(val => {
              const strVal = String(val).trim().toUpperCase();
              if (excludedTerms.includes(strVal)) return true;
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
            const rawPlanDelDate = String(row['Plan Del Date'] || row['Plan Del Date '] || '').trim();
            const planDelDate = rawPlanDelDate.length === 8 
              ? `${rawPlanDelDate.substring(0, 4)}/${rawPlanDelDate.substring(4, 6)}/${rawPlanDelDate.substring(6, 8)}`
              : rawPlanDelDate;
              
            let weekNo = String(row['WEEK NO'] || '');
            
            if (!weekNo && rawPlanDelDate.length === 8) {
              const year = parseInt(rawPlanDelDate.substring(0, 4));
              const month = parseInt(rawPlanDelDate.substring(4, 6)) - 1;
              const day = parseInt(rawPlanDelDate.substring(6, 8));
              const d = new Date(Date.UTC(year, month, day));
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

            const vpoStr = String(getVal(row, ['VPO No', 'vpo no']) || '').trim();
            const schedStr = String(getVal(row, ['Schedule No', 'schedule no']) || '').trim();
            const styleStr = String(getVal(row, ['Style No', 'style no']) || '').trim();
            const colorStr = String(getVal(row, ['Color Code', 'color code']) || '').trim();
            const destStr = String(getVal(row, ['Destination', 'destination']) || '').trim();
            
            // Create a safe, stable ID for Firestore
            const stableId = `${vpoStr}_${schedStr}_${styleStr}_${colorStr}_${destStr}_${planDelDate}`.replace(/[^a-zA-Z0-9_-]/g, '-');
            
            // Fallback to index if fields are empty to prevent overwriting
            const finalId = (stableId === '_____' || !stableId) ? `row_${index}` : stableId;

            return {
              id: finalId,
              buyer: getVal(row, ['Buyer']) || '',
              groupTechClass: getVal(row, ['Group Tech Class']) || '',
              buyerDivisionName: getVal(row, ['Buyer Division Name']) || '',
              styleNo: styleStr,
              custStyleNo: getVal(row, ['Cust Style No']) || '',
              vpoNo: vpoStr,
              shipmentMode: getVal(row, ['Shipment Mode']) || '',
              colorCode: colorStr,
              colorName: getVal(row, ['Color Name']) || '',
              destination: destStr,
              packMethod: String(getVal(row, ['Pack Method', 'Pack Method ']) || '').trim(),
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
              orderToShippedPct: getNum(row, ['Order to shipped %']),
            };
          });

          // Sort by Plan Del Date
          parsedData.sort((a, b) => a.planDelDate.localeCompare(b.planDelDate));

          setData(parsedData);
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

  // Filter Data
  const filteredItems = useMemo(() => {
    if (!data) return [];
    
    return data.filter(item => {
      const itemRemarkText = (remarks[item.id] || '').trim();
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
  }, [data, searchTerm, filterBuyer, filterWeekNo, filterStatus, filterShipmentMode, filterDestination, filterPackMethod, filterRemark, remarks]);

  const rowVirtualizer = useVirtualizer({
    count: filteredItems.length,
    getScrollElement: () => tableContainerRef.current,
    estimateSize: () => 44,
    overscan: 15,
  });

  const virtualRows = rowVirtualizer.getVirtualItems();
  const totalSize = rowVirtualizer.getTotalSize();
  const paddingTop = virtualRows.length > 0 ? virtualRows[0]?.start || 0 : 0;
  const paddingBottom = virtualRows.length > 0 ? totalSize - (virtualRows[virtualRows.length - 1]?.end || 0) : 0;

  const uniqueRemarksList = useMemo(() => {
    const suggestions = new Set(['DONE', 'PRINTING']);
    Object.values(remarks).forEach(r => {
      if (r && r.trim()) suggestions.add(r.trim().toUpperCase());
    });
    return Array.from(suggestions).sort();
  }, [remarks]);

  // Helper to get unique options based on current filters (excluding the filter itself)
  const getUniqueOptions = useCallback((field: keyof ProductionOrder | 'remark') => {
    if (!data) return [];
    const options = new Set<string>();
    
    data.forEach(item => {
      const itemRemarkText = (remarks[item.id] || '').trim();
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

  // Unique values for dropdowns
  const uniqueBuyers = useMemo(() => Array.from(new Set([...getUniqueOptions('buyer'), ...filterBuyer])).sort(), [getUniqueOptions, filterBuyer]);
  const uniqueWeeks = useMemo(() => Array.from(new Set([...getUniqueOptions('weekNo'), ...filterWeekNo])).sort((a,b) => Number(a) - Number(b)), [getUniqueOptions, filterWeekNo]);
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

  const [showPrintModal, setShowPrintModal] = useState(false);

  const executePrint = (orientation: 'portrait' | 'landscape') => {
    setShowPrintModal(false);
    const iframe = document.createElement('iframe');
    iframe.style.position = 'absolute';
    iframe.style.width = '0px';
    iframe.style.height = '0px';
    iframe.style.border = 'none';
    document.body.appendChild(iframe);

    const doc = iframe.contentWindow?.document;
    if (!doc) return;

    let tableRowsHTML = '';
    let prevVPO: string | null = null;
    
    filteredItems.forEach((item, index) => {
      if (index > 0 && item.vpoNo !== prevVPO) {
        tableRowsHTML += `<tr style="background-color: #1e293b; height: 4px;"><td colspan="13" style="padding: 0; border: none;"></td></tr>`;
      }
      prevVPO = item.vpoNo;
      tableRowsHTML += `
        <tr>
          <td>${item.planDelDate || ''}</td>
          <td>${item.weekNo || ''}</td>
          <td>${item.buyer || ''}</td>
          <td>${item.styleNo || ''}</td>
          <td>${item.vpoNo || ''}</td>
          <td>${item.shipmentMode || ''}</td>
          <td>${item.colorCode || ''}</td>
          <td>${item.colorName || ''}</td>
          <td>${item.destination || ''}</td>
          <td>${item.packMethod || ''}</td>
          <td>${item.scheduleNo || ''}</td>
          <td>${item.coQty != null ? item.coQty?.toLocaleString() || "0" : ''}</td>
          <td>${item.statusText || ''}</td>
        </tr>
      `;
    });

    doc.open();
    doc.write(`
      <!DOCTYPE html>
      <html>
      <head>
        <title>Production Data Report</title>
        <style>
          @page { size: A4 ${orientation}; margin: 10mm; }
          body { 
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; 
            font-size: 8pt; 
            color: #000; 
            background: #fff;
          }
          table { width: 100%; border-collapse: collapse; margin-top: 10px; }
          th, td { border: 1px solid #ccc; padding: 4px; text-align: left; word-wrap: break-word; }
          th { background-color: #f1f5f9; font-weight: bold; }
          h2 { font-size: 14pt; margin: 0 0 10px 0; }
          .header { margin-bottom: 15px; }
          .print-time { font-size: 8pt; color: #555; }
        </style>
      </head>
      <body>
        <div class="header">
          <h2>Production Orders Report</h2>
          <div class="print-time">Printed on: ${new Date()?.toLocaleString() || "0"} &bull; Total Records: ${filteredItems.length}</div>
        </div>
        <table>
          <thead>
            <tr>
              <th>Plan Del Date</th>
              <th>WEEK NO</th>
              <th>Buyer</th>
              <th>Style No</th>
              <th>VPO No</th>
              <th>Shipment Mode</th>
              <th>Color Code</th>
              <th>Color Name</th>
              <th>Destination</th>
              <th>Pack Method</th>
              <th>Schedule No</th>
              <th>CO Qty</th>
              <th>Status</th>
            </tr>
          </thead>
          <tbody>
            ${tableRowsHTML}
          </tbody>
        </table>
      </body>
      </html>
    `);
    doc.close();

    setTimeout(() => {
      iframe.contentWindow?.focus();
      iframe.contentWindow?.print();
      setTimeout(() => {
        document.body.removeChild(iframe);
      }, 1000);
    }, 250);
  };

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
                <input 
                  type="file" 
                  className="hidden" 
                  ref={fileInputRef} 
                  accept=".xlsx, .xls, .csv" 
                  onChange={handleFileUpload} 
                />
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

      <div className="mx-auto space-y-6" style={{ maxWidth: '1600px' }}>
        
        <header className="flex flex-col sm:flex-row sm:items-end justify-between gap-4 bg-white p-6 rounded-xl border border-slate-200 shadow-sm">
          <div>
            <div className="flex items-center gap-3">
              <h1 className="text-2xl font-bold tracking-tight text-slate-900">Production Overview</h1>
              <span className="bg-indigo-50 text-indigo-700 text-[9px] font-bold px-2 py-0.5 rounded border border-indigo-100 uppercase tracking-widest">
                Powered by DILEEPA WICKRAMASINGHE
              </span>
            </div>
            <p className="text-slate-500 mt-1.5 text-sm flex items-center gap-2">
              <span>Item-level fulfillment details for the next 6 weeks.</span>
              {lastUpdated && (
                <>
                  <span className="text-slate-300">•</span>
                  <span className="text-slate-600 font-medium flex items-center gap-1.5"><CheckCircle className="w-3.5 h-3.5 text-emerald-500" /> OrderBook Updated: {new Date(lastUpdated).toLocaleString('en-US', { dateStyle: 'medium', timeStyle: 'short' })}</span>
                </>
              )}
            </p>
          </div>
          <div className="flex items-center gap-3">
            <div className="flex bg-slate-100 p-1 rounded-lg border border-slate-200">
              <button
                onClick={() => setViewMode('table')}
                className={`flex items-center px-3 py-1.5 rounded-md text-sm font-medium transition-all ${
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
                className={`flex items-center px-3 py-1.5 rounded-md text-sm font-medium transition-all ${
                  viewMode === 'kpi' 
                    ? 'bg-white text-slate-800 shadow-sm border border-slate-200/50' 
                    : 'text-slate-500 hover:text-slate-700 hover:bg-slate-200/50'
                }`}
              >
                <BarChart2 className="w-4 h-4 mr-1.5" />
                KPI Dashboard
              </button>
            </div>
            <button 
              onClick={() => setData(null)}
              disabled={loadingState !== 'idle'}
              className="inline-flex items-center px-4 py-2.5 bg-slate-900 text-white rounded-lg text-sm font-medium hover:bg-slate-800 shadow-sm transition-all focus:ring-2 focus:ring-slate-900/20 active:scale-95 disabled:opacity-50 disabled:cursor-not-allowed disabled:active:scale-100"
            >
              {loadingState !== 'idle' ? <Loader2 className="w-4 h-4 mr-2 animate-spin" /> : <UploadCloud className="w-4 h-4 mr-2" />}
              {loadingState !== 'idle' ? "Processing..." : "Upload New File"}
            </button>
          </div>
        </header>

        <div style={{ display: viewMode === 'table' ? 'block' : 'none' }} className="space-y-6">
            {/* Summary Metrics */}
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
              <MetricCard title="Completed Schedule Line" value={summary.completedScheduleLines?.toLocaleString() || "0"} icon={<CheckCircle className="w-5 h-5 text-emerald-600" />} color="emerald" />
              <MetricCard title="Pending Schedule Line" value={summary.pendingScheduleLines?.toLocaleString() || "0"} icon={<AlertTriangle className="w-5 h-5 text-amber-600" />} color="amber" />
              <MetricCard title="Completed VPO" value={summary.completedVPOs?.toLocaleString() || "0"} icon={<CheckCircle className="w-5 h-5 text-blue-600" />} color="blue" />
              <MetricCard title="Pending VPO" value={summary.pendingVPOs?.toLocaleString() || "0"} icon={<AlertTriangle className="w-5 h-5 text-rose-600" />} color="rose" />
            </div>

            {/* Filters and Search Bar */}
            <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm space-y-4">
          <div className="flex flex-col lg:flex-row gap-4">
            <div className="relative w-full lg:w-80 shrink-0">
              <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
              <input 
                type="text" 
                placeholder="Search anything (style, color, vpo)..." 
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                className="w-full pl-9 pr-4 py-2 bg-slate-50 hover:bg-slate-100/50 focus:bg-white border border-slate-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 transition-all"
              />
            </div>
            
            <div className="flex flex-wrap sm:flex-nowrap gap-3 w-full">
              <MultiSelectDropdown 
                label="All Buyers"
                options={uniqueBuyers}
                selectedValues={filterBuyer}
                onChange={setFilterBuyer}
              />
              
              <MultiSelectDropdown 
                label="All Weeks"
                options={uniqueWeeks.map(w => String(w))}
                selectedValues={filterWeekNo}
                onChange={setFilterWeekNo}
              />

              <MultiSelectDropdown 
                label="All Statuses"
                options={['Shipped', 'Completed', 'Pending']}
                selectedValues={filterStatus}
                onChange={setFilterStatus}
              />
            </div>
          </div>

          <div className="flex flex-col sm:flex-row gap-4 items-center">
            <MultiSelectDropdown 
              label="All Shipment Modes"
              options={uniqueShipmentModes}
              selectedValues={filterShipmentMode}
              onChange={setFilterShipmentMode}
            />
            
            <MultiSelectDropdown 
              label="All Destinations"
              options={uniqueDestinations}
              selectedValues={filterDestination}
              onChange={setFilterDestination}
            />

            <MultiSelectDropdown 
              label="All Pack Methods"
              options={uniquePackMethods}
              selectedValues={filterPackMethod}
              onChange={setFilterPackMethod}
            />
            
            <MultiSelectDropdown 
              label="All Remarks"
              options={uniqueRemarks}
              selectedValues={filterRemark}
              onChange={setFilterRemark}
            />
            
            <button 
              onClick={() => {
                setSearchTerm(''); setFilterBuyer([]); setFilterWeekNo([]); setFilterStatus([]); setFilterShipmentMode([]); setFilterDestination([]); setFilterPackMethod([]); setFilterRemark([]);
              }}
              className="w-full sm:w-auto px-5 py-2 bg-slate-50 border border-slate-200 hover:bg-slate-100 text-slate-700 rounded-lg text-sm font-medium transition-colors whitespace-nowrap"
            >
              Clear Filters
            </button>
            <button
              onClick={() => setShowPrintModal(true)}
              disabled={filteredItems.length === 0}
              className="w-full sm:w-auto px-5 py-2 bg-indigo-50 hover:bg-indigo-100 text-indigo-700 rounded-lg text-sm font-medium transition-colors whitespace-nowrap flex items-center justify-center disabled:opacity-50 disabled:cursor-not-allowed border border-indigo-100"
            >
              <Printer className="w-4 h-4 mr-2" />
              Print
            </button>
          </div>
        </div>

        {/* Detailed Data Table */}
        <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden flex flex-col">
          <div ref={tableContainerRef} className="overflow-auto max-h-[70vh]">
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
                            initialValue={remarks[row.id] || ''}
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
          </div>
        </div>
        </div>

        <div style={{ display: viewMode === 'kpi' ? 'block' : 'none' }}>
          <KPIView data={data} />
        </div>

      </div>

      {/* Print Modal */}
      {showPrintModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/50 backdrop-blur-sm p-4">
          <div className="bg-white p-6 rounded-2xl shadow-xl w-full max-w-sm border border-slate-200">
            <h3 className="text-lg font-bold text-slate-900 mb-2">Print Options</h3>
            <p className="text-sm text-slate-500 mb-6">Choose how you want to print the report.</p>
            <div className="flex flex-col gap-3">
              <button 
                onClick={() => executePrint('portrait')} 
                className="w-full px-4 py-3 bg-indigo-50 hover:bg-indigo-100 text-indigo-700 rounded-xl font-medium transition-colors text-sm flex items-center justify-center"
              >
                Portrait
              </button>
              <button 
                onClick={() => executePrint('landscape')} 
                className="w-full px-4 py-3 bg-indigo-50 hover:bg-indigo-100 text-indigo-700 rounded-xl font-medium transition-colors text-sm flex items-center justify-center"
              >
                Landscape
              </button>
              <button 
                onClick={() => setShowPrintModal(false)} 
                className="w-full px-4 py-3 mt-2 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl font-medium transition-colors text-sm flex items-center justify-center"
              >
                Cancel
              </button>
            </div>
          </div>
        </div>
      )}
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
    <div className="bg-white rounded-xl border border-slate-200 p-5 shadow-sm hover:shadow-md transition-shadow duration-200 flex items-center gap-4 group">
      <div className={cn("p-3 rounded-lg border transition-colors", bgColors[color])}>
        {icon}
      </div>
      <div>
        <p className="text-[11px] font-semibold text-slate-500 mb-0.5 uppercase tracking-wider group-hover:text-slate-600 transition-colors">{title}</p>
        <p className="text-2xl font-bold text-slate-900 tracking-tight">{value}</p>
      </div>
    </div>
  );
}

function cn(...classes: (string | undefined | null | false)[]) {
  return classes.filter(Boolean).join(' ');
}
