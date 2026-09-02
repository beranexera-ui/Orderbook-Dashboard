import React, { useState, useRef, useCallback, useMemo, useEffect } from 'react';
import { UploadCloud, FileSpreadsheet, AlertCircle, Search, Package, CheckCircle, TrendingUp, AlertTriangle, Loader2 } from 'lucide-react';
import * as XLSX from 'xlsx';
import { useVirtualizer } from '@tanstack/react-virtual';
import { db } from '../firebase';
import { doc, setDoc, onSnapshot, writeBatch, collection, getDocs } from 'firebase/firestore';

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
  statusText: string;
  deliveredQty: number;
  orderToShippedPct: number;
  remark?: string;
}

export function Dashboard() {
  const [data, setData] = useState<ProductionOrder[] | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const [loadingState, setLoadingState] = useState<'idle' | 'reading' | 'parsing' | 'uploading'>('idle');
  const [uploadProgress, setUploadProgress] = useState({ current: 0, total: 0 });
  const [loadingInitial, setLoadingInitial] = useState(true);
  const [remarks, setRemarks] = useState<Record<string, string>>({});
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [error, setError] = useState<string | null>(null);

  // Filters State
  const [searchTerm, setSearchTerm] = useState('');
  const [filterBuyer, setFilterBuyer] = useState('');
  const [filterWeekNo, setFilterWeekNo] = useState('');
  const [filterStatus, setFilterStatus] = useState('');
  const [filterShipmentMode, setFilterShipmentMode] = useState('');
  const [filterDestination, setFilterDestination] = useState('');
  const [filterPackMethod, setFilterPackMethod] = useState('');
  
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
        const { currentUploadId } = docSnap.data();
        if (currentUploadId) {
          try {
            const snapshot = await getDocs(collection(db, `uploads/${currentUploadId}/orders`));
            const loadedData: ProductionOrder[] = [];
            snapshot.forEach(d => {
               loadedData.push(d.data() as ProductionOrder);
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
            'WASH & TOP', 'PPZ', 'TC-PP', 'TC-PPZ'
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

            const coQty = Number(row['CO Qty']) || 0;
            const cumSewOutQty = Number(row['Cum SewOut Qty']) || 0;
            const deliveredQty = Number(row['Delivered Qty']) || 0;
            const pendingQty = coQty - cumSewOutQty;
            
            let statusText = '';
            if (deliveredQty > 0) {
              statusText = 'Shipped';
            } else if (cumSewOutQty >= coQty) {
              statusText = 'Completed';
            } else {
              statusText = `Pending - ${pendingQty}`;
            }

            const vpoStr = String(row['VPO No'] || '').trim();
            const schedStr = String(row['Schedule No'] || '').trim();
            const styleStr = String(row['Style No'] || '').trim();
            const colorStr = String(row['Color Code'] || '').trim();
            const destStr = String(row['Destination'] || '').trim();
            
            // Create a safe, stable ID for Firestore
            const stableId = `${vpoStr}_${schedStr}_${styleStr}_${colorStr}_${destStr}_${planDelDate}`.replace(/[^a-zA-Z0-9_-]/g, '-');
            
            // Fallback to index if fields are empty to prevent overwriting
            const finalId = (stableId === '_____' || !stableId) ? `row_${index}` : stableId;

            return {
              id: finalId,
              buyer: row['Buyer'] || '',
              groupTechClass: row['Group Tech Class'] || '',
              buyerDivisionName: row['Buyer Division Name'] || '',
              styleNo: row['Style No'] || '',
              custStyleNo: row['Cust Style No'] || '',
              vpoNo: row['VPO No'] || '',
              shipmentMode: row['Shipment Mode'] || '',
              colorCode: row['Color Code'] || '',
              colorName: row['Color Name'] || '',
              destination: row['Destination'] || '',
              packMethod: row['Pack Method'] || '',
              scheduleNo: row['Schedule No'] || '',
              planDelDate: planDelDate,
              weekNo: weekNo,
              coQty: coQty,
              cumSewInQty: Number(row['Cum Sew In Qty']) || 0,
              cumSewOutQty: cumSewOutQty,
              cumSewOutRejQty: Number(row['Cum Sew Out Rej Qty']) || 0,
              statusText: statusText,
              deliveredQty: Number(row['Delivered Qty']) || 0,
              orderToShippedPct: Number(row['Order to shipped %']) || 0,
            };
          });

          // Sort by Plan Del Date
          parsedData.sort((a, b) => a.planDelDate.localeCompare(b.planDelDate));

          setData(parsedData);
          uploadToFirestore(parsedData);
          
          // Reset states on new upload
          setSearchTerm('');
          setFilterBuyer('');
          setFilterWeekNo('');
          setFilterStatus('');
          setFilterShipmentMode('');
          setFilterDestination('');
          setFilterPackMethod('');

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
      const matchesSearch = searchTerm === '' || 
        Object.values(item).some(val => 
          String(val).toLowerCase().includes(searchTerm.toLowerCase())
        );
      
      const matchesBuyer = filterBuyer === '' || item.buyer === filterBuyer;
      const matchesWeek = filterWeekNo === '' || item.weekNo === filterWeekNo;
      
      // Status matching logic (Completed vs Pending vs Shipped)
      const matchesStatus = filterStatus === '' || 
        (filterStatus === 'Completed' ? item.statusText === 'Completed' : 
         filterStatus === 'Shipped' ? item.statusText === 'Shipped' :
         filterStatus === 'Pending' ? item.statusText.startsWith('Pending') : true);
        
      const matchesShipmentMode = filterShipmentMode === '' || item.shipmentMode === filterShipmentMode;
      const matchesDestination = filterDestination === '' || item.destination === filterDestination;
      const matchesPackMethod = filterPackMethod === '' || item.packMethod === filterPackMethod;

      return matchesSearch && matchesBuyer && matchesWeek && matchesStatus && matchesShipmentMode && matchesDestination && matchesPackMethod;
    });
  }, [data, searchTerm, filterBuyer, filterWeekNo, filterStatus, filterShipmentMode, filterDestination, filterPackMethod]);

  const rowVirtualizer = useVirtualizer({
    count: filteredItems.length,
    getScrollElement: () => tableContainerRef.current,
    estimateSize: () => 45,
    overscan: 10,
  });

  const virtualRows = rowVirtualizer.getVirtualItems();
  const totalSize = rowVirtualizer.getTotalSize();
  const paddingTop = virtualRows.length > 0 ? virtualRows[0]?.start || 0 : 0;
  const paddingBottom = virtualRows.length > 0 ? totalSize - (virtualRows[virtualRows.length - 1]?.end || 0) : 0;

  // Unique values for dropdowns
  const uniqueBuyers = useMemo(() => Array.from(new Set(data?.map(d => d.buyer))).filter(Boolean).sort(), [data]);
  const uniqueWeeks = useMemo(() => Array.from(new Set(data?.map(d => d.weekNo))).filter(Boolean).sort((a,b) => Number(a) - Number(b)), [data]);
  const uniqueShipmentModes = useMemo(() => Array.from(new Set(data?.map(d => d.shipmentMode))).filter(Boolean).sort(), [data]);
  const uniqueDestinations = useMemo(() => Array.from(new Set(data?.map(d => d.destination))).filter(Boolean).sort(), [data]);
  const uniquePackMethods = useMemo(() => Array.from(new Set(data?.map(d => d.packMethod))).filter(Boolean).sort(), [data]);

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
                      {uploadProgress.total > 0 ? `${uploadProgress.current.toLocaleString()} / ${uploadProgress.total.toLocaleString()} rows synced` : "Preparing sync..."}
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
      <div className="mx-auto space-y-6" style={{ maxWidth: '1600px' }}>
        
        <header className="flex flex-col sm:flex-row sm:items-end justify-between gap-4 bg-white p-6 rounded-2xl border border-slate-200 shadow-sm">
          <div>
            <h1 className="text-2xl font-bold tracking-tight text-slate-900">Production Overview</h1>
            <p className="text-slate-500 mt-1">Item-level fulfillment details for the next 6 weeks.</p>
          </div>
          <button 
            onClick={() => setData(null)}
            disabled={loadingState !== 'idle'}
            className="inline-flex items-center px-4 py-2.5 bg-slate-900 text-white rounded-xl text-sm font-medium hover:bg-slate-800 shadow-sm transition-colors cursor-pointer disabled:opacity-50 disabled:cursor-not-allowed"
          >
            {loadingState !== 'idle' ? <Loader2 className="w-4 h-4 mr-2 animate-spin" /> : <UploadCloud className="w-4 h-4 mr-2" />}
            {loadingState !== 'idle' ? "Processing..." : "Upload New File"}
          </button>
        </header>

        {/* Summary Metrics */}
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
          <MetricCard title="Completed Schedule Line" value={summary.completedScheduleLines.toLocaleString()} icon={<CheckCircle className="w-5 h-5 text-emerald-500" />} />
          <MetricCard title="Pending Schedule Line" value={summary.pendingScheduleLines.toLocaleString()} icon={<AlertTriangle className="w-5 h-5 text-amber-500" />} />
          <MetricCard title="Completed VPO" value={summary.completedVPOs.toLocaleString()} icon={<CheckCircle className="w-5 h-5 text-blue-500" />} />
          <MetricCard title="Pending VPO" value={summary.pendingVPOs.toLocaleString()} icon={<AlertTriangle className="w-5 h-5 text-orange-500" />} />
        </div>

        {/* Filters and Search Bar */}
        <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm space-y-4">
          <div className="flex flex-col lg:flex-row gap-4">
            <div className="relative flex-1">
              <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
              <input 
                type="text" 
                placeholder="Search anything (style, color, vpo)..." 
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                className="w-full pl-9 pr-4 py-2 bg-slate-50 border border-slate-200 rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 transition-all"
              />
            </div>
            
            <div className="flex flex-wrap sm:flex-nowrap gap-3">
              <select value={filterBuyer} onChange={(e) => setFilterBuyer(e.target.value)} className="flex-1 sm:w-auto py-2 pl-3 pr-8 bg-slate-50 border border-slate-200 rounded-xl text-sm focus:outline-none focus:border-indigo-500">
                <option value="">All Buyers</option>
                {uniqueBuyers.map(b => <option key={b} value={b}>{b}</option>)}
              </select>
              
              <select value={filterWeekNo} onChange={(e) => setFilterWeekNo(e.target.value)} className="flex-1 sm:w-auto py-2 pl-3 pr-8 bg-slate-50 border border-slate-200 rounded-xl text-sm focus:outline-none focus:border-indigo-500">
                <option value="">All Weeks</option>
                {uniqueWeeks.map(w => <option key={w} value={w}>Week {w}</option>)}
              </select>

              <select value={filterStatus} onChange={(e) => setFilterStatus(e.target.value)} className="flex-1 sm:w-auto py-2 pl-3 pr-8 bg-slate-50 border border-slate-200 rounded-xl text-sm focus:outline-none focus:border-indigo-500">
                <option value="">All Statuses</option>
                <option value="Shipped">Shipped</option>
                <option value="Completed">Completed</option>
                <option value="Pending">Pending</option>
              </select>
            </div>
          </div>

          <div className="flex flex-col sm:flex-row gap-4 items-center">
            <select value={filterShipmentMode} onChange={(e) => setFilterShipmentMode(e.target.value)} className="w-full sm:flex-1 py-2 pl-3 pr-8 bg-slate-50 border border-slate-200 rounded-xl text-sm focus:outline-none focus:border-indigo-500">
              <option value="">All Shipment Modes</option>
              {uniqueShipmentModes.map(s => <option key={s} value={s}>{s}</option>)}
            </select>
            
            <select value={filterDestination} onChange={(e) => setFilterDestination(e.target.value)} className="w-full sm:flex-1 py-2 pl-3 pr-8 bg-slate-50 border border-slate-200 rounded-xl text-sm focus:outline-none focus:border-indigo-500">
              <option value="">All Destinations</option>
              {uniqueDestinations.map(d => <option key={d} value={d}>{d}</option>)}
            </select>

            <select value={filterPackMethod} onChange={(e) => setFilterPackMethod(e.target.value)} className="w-full sm:flex-1 py-2 pl-3 pr-8 bg-slate-50 border border-slate-200 rounded-xl text-sm focus:outline-none focus:border-indigo-500">
              <option value="">All Pack Methods</option>
              {uniquePackMethods.map(p => <option key={p} value={p}>{p}</option>)}
            </select>
            
            <button 
              onClick={() => {
                setSearchTerm(''); setFilterBuyer(''); setFilterWeekNo(''); setFilterStatus(''); setFilterShipmentMode(''); setFilterDestination(''); setFilterPackMethod('');
              }}
              className="w-full sm:w-auto px-5 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl text-sm font-medium transition-colors whitespace-nowrap"
            >
              Clear Filters
            </button>
          </div>
        </div>

        {/* Detailed Data Table */}
        <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden flex flex-col">
          <div ref={tableContainerRef} className="overflow-auto max-h-[70vh]">
            <table className="w-full text-sm text-left border-collapse relative">
              <thead className="text-xs text-slate-600 font-semibold bg-slate-100 border-b border-slate-200 sticky top-0 z-20 shadow-sm">
                <tr>
                  <th className="px-3 py-3 border border-slate-200 whitespace-nowrap">Plan Del Date</th>
                  <th className="px-3 py-3 border border-slate-200 whitespace-nowrap text-center">WEEK NO</th>
                  <th className="px-3 py-3 border border-slate-200 whitespace-nowrap">Buyer</th>
                  <th className="px-3 py-3 border border-slate-200 whitespace-nowrap">Group Tech Class</th>
                  <th className="px-3 py-3 border border-slate-200 whitespace-nowrap">Buyer Division Name</th>
                  <th className="px-3 py-3 border border-slate-200 whitespace-nowrap">Style No</th>
                  <th className="px-3 py-3 border border-slate-200 whitespace-nowrap">Cust Style No</th>
                  <th className="px-3 py-3 border border-slate-200 whitespace-nowrap">VPO No</th>
                  <th className="px-3 py-3 border border-slate-200 whitespace-nowrap">Shipment Mode</th>
                  <th className="px-3 py-3 border border-slate-200 whitespace-nowrap">Color Code</th>
                  <th className="px-3 py-3 border border-slate-200 whitespace-nowrap">Color Name</th>
                  <th className="px-3 py-3 border border-slate-200 whitespace-nowrap">Destination</th>
                  <th className="px-3 py-3 border border-slate-200 whitespace-nowrap">Pack Method</th>
                  <th className="px-3 py-3 border border-slate-200 whitespace-nowrap">Schedule No</th>
                  <th className="px-3 py-3 border border-slate-200 whitespace-nowrap text-right">CO Qty</th>
                  <th className="px-3 py-3 border border-slate-200 whitespace-nowrap text-right bg-slate-50">Cum Sew In Qty</th>
                  <th className="px-3 py-3 border border-slate-200 whitespace-nowrap text-right bg-slate-50">Cum SewOut Qty</th>
                  <th className="px-3 py-3 border border-slate-200 whitespace-nowrap text-right bg-slate-50">Cum Sew Out Rej</th>
                  <th className="px-3 py-3 border border-slate-200 whitespace-nowrap text-center">Status</th>
                  <th className="px-3 py-3 border border-slate-200 whitespace-nowrap">Remark</th>
                  <th className="px-3 py-3 border border-slate-200 whitespace-nowrap text-right">Delivered Qty</th>
                </tr>
              </thead>
              <tbody>
                {paddingTop > 0 && (
                  <tr>
                    <td colSpan={21} style={{ height: `${paddingTop}px` }}></td>
                  </tr>
                )}
                {virtualRows.length > 0 ? (
                  virtualRows.map((virtualRow) => {
                    const row = filteredItems[virtualRow.index];
                    return (
                      <tr key={row.id} className="hover:bg-indigo-50/50 transition-colors bg-white">
                        <td className="px-3 py-2.5 border border-slate-200 text-slate-700 whitespace-nowrap">{row.planDelDate}</td>
                      <td className="px-3 py-2.5 border border-slate-200 text-slate-700 whitespace-nowrap text-center font-medium bg-slate-50/50">{row.weekNo}</td>
                      <td className="px-3 py-2.5 border border-slate-200 text-slate-700 whitespace-nowrap">{row.buyer}</td>
                      <td className="px-3 py-2.5 border border-slate-200 text-slate-700 whitespace-nowrap">{row.groupTechClass}</td>
                      <td className="px-3 py-2.5 border border-slate-200 text-slate-700 whitespace-nowrap truncate max-w-[150px]" title={row.buyerDivisionName}>{row.buyerDivisionName}</td>
                      <td className="px-3 py-2.5 border border-slate-200 font-medium text-slate-900 whitespace-nowrap">{row.styleNo}</td>
                      <td className="px-3 py-2.5 border border-slate-200 text-slate-700 whitespace-nowrap">{row.custStyleNo}</td>
                      <td className="px-3 py-2.5 border border-slate-200 text-slate-700 whitespace-nowrap">{row.vpoNo}</td>
                      <td className="px-3 py-2.5 border border-slate-200 text-slate-700 whitespace-nowrap">{row.shipmentMode}</td>
                      <td className="px-3 py-2.5 border border-slate-200 text-slate-700 whitespace-nowrap">{row.colorCode}</td>
                      <td className="px-3 py-2.5 border border-slate-200 text-slate-700 whitespace-nowrap truncate max-w-[150px]" title={row.colorName}>{row.colorName}</td>
                      <td className="px-3 py-2.5 border border-slate-200 text-slate-700 whitespace-nowrap">{row.destination}</td>
                      <td className="px-3 py-2.5 border border-slate-200 text-slate-700 whitespace-nowrap truncate max-w-[150px]" title={row.packMethod}>{row.packMethod}</td>
                      <td className="px-3 py-2.5 border border-slate-200 text-slate-700 whitespace-nowrap">{row.scheduleNo}</td>
                      
                      <td className="px-3 py-2.5 border border-slate-200 text-right font-medium text-slate-900 whitespace-nowrap">{row.coQty.toLocaleString()}</td>
                      <td className="px-3 py-2.5 border border-slate-200 text-right text-slate-600 whitespace-nowrap bg-slate-50/50">{row.cumSewInQty.toLocaleString()}</td>
                      <td className="px-3 py-2.5 border border-slate-200 text-right font-medium text-indigo-700 whitespace-nowrap bg-slate-50/50">{row.cumSewOutQty.toLocaleString()}</td>
                      <td className="px-3 py-2.5 border border-slate-200 text-right text-rose-600 whitespace-nowrap bg-slate-50/50">{row.cumSewOutRejQty.toLocaleString()}</td>
                      
                      <td className="px-3 py-2.5 border border-slate-200 text-center whitespace-nowrap">
                        <span className={cn(
                          "inline-flex items-center px-2 py-0.5 rounded-full font-medium text-[11px] uppercase tracking-wider",
                          row.statusText === 'Shipped' ? "bg-blue-100 text-blue-700" :
                          row.statusText === 'Completed' ? "bg-emerald-100 text-emerald-700" : 
                          "bg-amber-100 text-amber-800"
                        )}>
                          {row.statusText}
                        </span>
                      </td>
                      <td className="px-3 py-2.5 border border-slate-200 whitespace-nowrap min-w-[200px]">
                        <input 
                          type="text" 
                          value={remarks[row.id] || ''}
                          onChange={(e) => {
                            setRemarks(prev => ({ ...prev, [row.id]: e.target.value }));
                          }}
                          onBlur={(e) => handleRemarkChange(row.id, e.target.value)}
                          placeholder="Add remark..."
                          className="w-full text-sm bg-transparent border-0 border-b border-transparent hover:border-slate-300 focus:border-indigo-500 focus:ring-0 px-1 py-1 transition-colors"
                        />
                      </td>
                      <td className="px-3 py-2.5 border border-slate-200 text-right font-medium text-slate-900 whitespace-nowrap">{row.deliveredQty.toLocaleString()}</td>
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
                  <tr>
                    <td colSpan={21} style={{ height: `${paddingBottom}px` }}></td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
        </div>

      </div>
    </div>
  );
}

function MetricCard({ title, value, icon }: { title: string, value: string, icon: React.ReactNode }) {
  return (
    <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm flex items-center gap-4">
      <div className="p-3 bg-indigo-50 rounded-xl border border-indigo-100">
        {icon}
      </div>
      <div>
        <p className="text-[11px] font-semibold text-slate-500 mb-0.5 uppercase tracking-wider">{title}</p>
        <p className="text-2xl font-bold text-slate-900">{value}</p>
      </div>
    </div>
  );
}

function cn(...classes: (string | undefined | null | false)[]) {
  return classes.filter(Boolean).join(' ');
}
