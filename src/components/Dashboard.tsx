import React, { useState, useRef, useCallback } from 'react';
import { UploadCloud, FileSpreadsheet, AlertCircle } from 'lucide-react';
import * as XLSX from 'xlsx';

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
}

export function Dashboard() {
  const [data, setData] = useState<ProductionOrder[] | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const [error, setError] = useState<string | null>(null);

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      processFile(file);
    }
  };

  const processFile = (file: File) => {
    setError(null);
    const reader = new FileReader();
    reader.onload = (e) => {
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
          
          // Get start of current week (assuming Monday)
          const dayOfWeek = today.getDay();
          const diffToMonday = today.getDate() - dayOfWeek + (dayOfWeek === 0 ? -6 : 1);
          const startOfCurrentWeek = new Date(today);
          startOfCurrentWeek.setDate(diffToMonday);

          const endOf6Weeks = new Date(startOfCurrentWeek);
          endOf6Weeks.setDate(startOfCurrentWeek.getDate() + 42); // 6 weeks from start of this week
          
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
            
            return isERK && isWithin6Weeks;
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
            
            // Auto-generate WEEK NO from YYYYMMDD if missing
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
            const pendingQty = coQty - cumSewOutQty;
            const statusText = cumSewOutQty >= coQty ? 'Completed' : `Pending - ${pendingQty}`;

            return {
              id: index.toString(),
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

          setData(parsedData);
        } catch (err: any) {
          console.error("Error parsing Excel file", err);
          if (err.message && err.message.includes("Encrypted file")) {
            setError("This Excel file is encrypted or password-protected. Please remove the password protection and try again.");
          } else {
            setError("Failed to parse the Excel file. Please ensure it is a valid format.");
          }
        }
      }
    };
    reader.readAsArrayBuffer(file);
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

  if (!data) {
    return (
      <div className="min-h-screen bg-slate-50 text-slate-900 font-sans p-6 md:p-10 flex flex-col items-center justify-center">
        <div className="max-w-md w-full bg-white rounded-3xl border border-slate-200 shadow-sm p-10 text-center">
          <div className="bg-indigo-50 w-20 h-20 rounded-full flex items-center justify-center mx-auto mb-6">
            <FileSpreadsheet className="w-10 h-10 text-indigo-500" />
          </div>
          <h2 className="text-2xl font-bold text-slate-900 mb-2">Upload Production Data</h2>
          <p className="text-slate-500 mb-8">Upload your Excel (.xlsx) file to instantly generate the production dashboard.</p>
          
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
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 font-sans p-6 md:p-10">
      <div className="max-w-7xl mx-auto space-y-8">
        
        <header className="flex flex-col md:flex-row md:items-end justify-between gap-4">
          <div>
            <h1 className="text-3xl font-bold tracking-tight text-slate-900">Production Overview</h1>
            <p className="text-slate-500 mt-2">Visualizing item-level fulfillment details.</p>
          </div>
          <button 
            onClick={() => setData(null)}
            className="inline-flex items-center px-4 py-2 bg-white border border-slate-200 rounded-lg text-sm font-medium text-slate-700 hover:bg-slate-50 shadow-sm transition-colors cursor-pointer"
          >
            <UploadCloud className="w-4 h-4 mr-2" />
            Upload New File
          </button>
        </header>

        {/* Detailed Data Table */}
        <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
          <div className="p-6 border-b border-slate-200">
            <h2 className="text-lg font-semibold text-slate-800">Fulfillment Details</h2>
            <p className="text-sm text-slate-500">Item-level breakdown of the uploaded dataset.</p>
          </div>
          <div className="overflow-x-auto max-h-[70vh]">
            <table className="w-full text-sm text-left relative">
              <thead className="text-xs text-slate-500 uppercase bg-slate-50/90 border-b border-slate-200 sticky top-0 backdrop-blur-sm">
                <tr>
                  <th className="px-4 py-4 font-medium whitespace-nowrap">Buyer</th>
                  <th className="px-4 py-4 font-medium whitespace-nowrap">Group Tech Class</th>
                  <th className="px-4 py-4 font-medium whitespace-nowrap">Buyer Division Name</th>
                  <th className="px-4 py-4 font-medium whitespace-nowrap">Style No</th>
                  <th className="px-4 py-4 font-medium whitespace-nowrap">Cust Style No</th>
                  <th className="px-4 py-4 font-medium whitespace-nowrap">VPO No</th>
                  <th className="px-4 py-4 font-medium whitespace-nowrap">Shipment Mode</th>
                  <th className="px-4 py-4 font-medium whitespace-nowrap">Color Code</th>
                  <th className="px-4 py-4 font-medium whitespace-nowrap">Color Name</th>
                  <th className="px-4 py-4 font-medium whitespace-nowrap">Destination</th>
                  <th className="px-4 py-4 font-medium whitespace-nowrap">Pack Method</th>
                  <th className="px-4 py-4 font-medium whitespace-nowrap">Schedule No</th>
                  <th className="px-4 py-4 font-medium whitespace-nowrap">Plan Del Date</th>
                  <th className="px-4 py-4 font-medium whitespace-nowrap text-center">WEEK NO</th>
                  <th className="px-4 py-4 font-medium whitespace-nowrap text-right">CO Qty</th>
                  <th className="px-4 py-4 font-medium whitespace-nowrap text-right">Cum Sew In Qty</th>
                  <th className="px-4 py-4 font-medium whitespace-nowrap text-right">Cum SewOut Qty</th>
                  <th className="px-4 py-4 font-medium whitespace-nowrap text-right">Status</th>
                  <th className="px-4 py-4 font-medium whitespace-nowrap text-right">Cum Sew Out Rej Qty</th>
                  <th className="px-4 py-4 font-medium whitespace-nowrap text-right">Delivered Qty</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {data.map((row) => (
                  <tr key={row.id} className="hover:bg-slate-50 transition-colors">
                    <td className="px-4 py-4 text-slate-700 whitespace-nowrap">{row.buyer}</td>
                    <td className="px-4 py-4 text-slate-700 whitespace-nowrap">{row.groupTechClass}</td>
                    <td className="px-4 py-4 text-slate-700 whitespace-nowrap">{row.buyerDivisionName}</td>
                    <td className="px-4 py-4 font-medium text-slate-900 whitespace-nowrap">{row.styleNo}</td>
                    <td className="px-4 py-4 text-slate-700 whitespace-nowrap">{row.custStyleNo}</td>
                    <td className="px-4 py-4 text-slate-700 whitespace-nowrap">{row.vpoNo}</td>
                    <td className="px-4 py-4 text-slate-700 whitespace-nowrap">{row.shipmentMode}</td>
                    <td className="px-4 py-4 text-slate-700 whitespace-nowrap">{row.colorCode}</td>
                    <td className="px-4 py-4 text-slate-700 whitespace-nowrap">{row.colorName}</td>
                    <td className="px-4 py-4 text-slate-700 whitespace-nowrap">{row.destination}</td>
                    <td className="px-4 py-4 text-slate-700 whitespace-nowrap truncate max-w-[200px]" title={row.packMethod}>{row.packMethod}</td>
                    <td className="px-4 py-4 text-slate-700 whitespace-nowrap">{row.scheduleNo}</td>
                    <td className="px-4 py-4 text-slate-700 whitespace-nowrap">{row.planDelDate}</td>
                    <td className="px-4 py-4 text-slate-700 whitespace-nowrap text-center">{row.weekNo}</td>
                    <td className="px-4 py-4 text-right font-medium text-slate-900 whitespace-nowrap">{row.coQty.toLocaleString()}</td>
                    <td className="px-4 py-4 text-right text-slate-700 whitespace-nowrap">{row.cumSewInQty.toLocaleString()}</td>
                    <td className="px-4 py-4 text-right text-slate-700 whitespace-nowrap">{row.cumSewOutQty.toLocaleString()}</td>
                    <td className="px-4 py-4 text-right whitespace-nowrap">
                      <span className={cn(
                        "inline-flex items-center px-2.5 py-0.5 rounded-full font-medium text-xs",
                        row.statusText === 'Completed' ? "bg-emerald-100 text-emerald-700" : "bg-amber-100 text-amber-700"
                      )}>
                        {row.statusText}
                      </span>
                    </td>
                    <td className="px-4 py-4 text-right text-slate-700 whitespace-nowrap">{row.cumSewOutRejQty.toLocaleString()}</td>
                    <td className="px-4 py-4 text-right font-medium text-slate-900 whitespace-nowrap">{row.deliveredQty.toLocaleString()}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>

      </div>
    </div>
  );
}

function cn(...classes: (string | undefined | null | false)[]) {
  return classes.filter(Boolean).join(' ');
}
