import React, { useState, useRef, useCallback } from 'react';
import { 
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, 
  PieChart, Pie, Cell 
} from 'recharts';
import { Package, TrendingUp, CheckCircle, AlertCircle, UploadCloud, FileSpreadsheet } from 'lucide-react';
import * as XLSX from 'xlsx';

const COLORS = ['#0ea5e9', '#6366f1', '#10b981', '#f59e0b', '#ef4444'];

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

          const currentYear = new Date().getFullYear().toString();
          
          const filteredData = jsonData.filter((row: any) => {
            const warehouse = String(row['Prod Warehouse'] || '').trim().toUpperCase();
            const isERK = warehouse === 'ERK';
            
            // Checking common date fields in the dataset to match current year (e.g., "2026")
            const placementDate = String(row['Order placement date'] || '').trim();
            const fobDate = String(row['FOB Date'] || '').trim();
            const reqDelDate = String(row['Req Del date'] || '').trim();
            
            const isCurrentYear = placementDate.startsWith(currentYear) || 
                                  fobDate.startsWith(currentYear) || 
                                  reqDelDate.startsWith(currentYear);
            
            return isERK && isCurrentYear;
          });

          if (filteredData.length === 0) {
            setError(`No data found for Prod Warehouse "ERK" in the current year (${currentYear}).`);
            return;
          }

          const parsedData: ProductionOrder[] = filteredData.map((row: any, index: number) => {
            const planDelDate = String(row['Plan Del Date'] || row['Plan Del Date '] || '');
            let weekNo = String(row['WEEK NO'] || '');
            
            // Auto-generate WEEK NO from YYYYMMDD if missing
            if (!weekNo && planDelDate.length === 8) {
              const year = parseInt(planDelDate.substring(0, 4));
              const month = parseInt(planDelDate.substring(4, 6)) - 1;
              const day = parseInt(planDelDate.substring(6, 8));
              const d = new Date(Date.UTC(year, month, day));
              const dayNum = d.getUTCDay() || 7;
              d.setUTCDate(d.getUTCDate() + 4 - dayNum);
              const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
              weekNo = Math.ceil((((d.getTime() - yearStart.getTime()) / 86400000) + 1) / 7).toString();
            }

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
              coQty: Number(row['CO Qty']) || 0,
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

  const totalOrdered = data.reduce((acc, curr) => acc + curr.coQty, 0);
  const totalDelivered = data.reduce((acc, curr) => acc + curr.deliveredQty, 0);
  const overallFulfillment = totalOrdered ? ((totalDelivered / totalOrdered) * 100).toFixed(2) : "0.00";
  
  const shortShippedItems = data.filter(item => item.orderToShippedPct < 100).length;

  const aggregatedByStyle = data.reduce((acc, curr) => {
    const baseStyle = curr.styleNo.split('-')[0];
    if (!acc[baseStyle]) {
      acc[baseStyle] = { styleNo: baseStyle, coQty: 0, deliveredQty: 0 };
    }
    acc[baseStyle].coQty += curr.coQty;
    acc[baseStyle].deliveredQty += curr.deliveredQty;
    return acc;
  }, {} as Record<string, { styleNo: string, coQty: number, deliveredQty: number }>);
  
  const chartDataStyle = Object.values(aggregatedByStyle);

  const aggregatedByBuyer = data.reduce((acc, curr) => {
    if (!acc[curr.buyer]) {
      acc[curr.buyer] = { buyer: curr.buyer, qty: 0 };
    }
    acc[curr.buyer].qty += curr.coQty;
    return acc;
  }, {} as Record<string, { buyer: string, qty: number }>);

  const chartDataBuyer = Object.values(aggregatedByBuyer);

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 font-sans p-6 md:p-10">
      <div className="max-w-7xl mx-auto space-y-8">
        
        <header className="flex flex-col md:flex-row md:items-end justify-between gap-4">
          <div>
            <h1 className="text-3xl font-bold tracking-tight text-slate-900">Production Overview</h1>
            <p className="text-slate-500 mt-2">Visualizing order quantities, delivery status, and fulfillment efficiency.</p>
          </div>
          <button 
            onClick={() => setData(null)}
            className="inline-flex items-center px-4 py-2 bg-white border border-slate-200 rounded-lg text-sm font-medium text-slate-700 hover:bg-slate-50 shadow-sm transition-colors cursor-pointer"
          >
            <UploadCloud className="w-4 h-4 mr-2" />
            Upload New File
          </button>
        </header>

        {/* Metrics Grid */}
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6">
          <MetricCard 
            title="Total Ordered Qty" 
            value={totalOrdered.toLocaleString()} 
            icon={<Package className="w-5 h-5 text-blue-500" />}
          />
          <MetricCard 
            title="Total Delivered Qty" 
            value={totalDelivered.toLocaleString()} 
            icon={<CheckCircle className="w-5 h-5 text-emerald-500" />}
          />
          <MetricCard 
            title="Overall Fulfillment" 
            value={`${overallFulfillment}%`} 
            icon={<TrendingUp className="w-5 h-5 text-indigo-500" />}
          />
          <MetricCard 
            title="Short Shipped Orders" 
            value={shortShippedItems.toString()} 
            icon={<AlertCircle className="w-5 h-5 text-rose-500" />}
          />
        </div>

        {/* Main Charts Area */}
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
          
          <div className="lg:col-span-2 bg-white rounded-2xl border border-slate-200 shadow-sm p-6">
            <div className="mb-6">
              <h2 className="text-lg font-semibold text-slate-800">Order vs. Delivered by Style</h2>
              <p className="text-sm text-slate-500">Comparison of Customer Order (CO) Qty and Delivered Qty</p>
            </div>
            <div className="h-[400px] w-full">
              <ResponsiveContainer width="100%" height="100%">
                <BarChart
                  data={chartDataStyle}
                  margin={{ top: 20, right: 30, left: 20, bottom: 5 }}
                >
                  <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                  <XAxis dataKey="styleNo" axisLine={false} tickLine={false} tick={{ fill: '#64748b' }} dy={10} />
                  <YAxis axisLine={false} tickLine={false} tick={{ fill: '#64748b' }} />
                  <Tooltip 
                    contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 10px 15px -3px rgb(0 0 0 / 0.1)' }}
                    cursor={{ fill: '#f1f5f9' }}
                  />
                  <Legend wrapperStyle={{ paddingTop: '20px' }} />
                  <Bar dataKey="coQty" name="Ordered Qty" fill="#0ea5e9" radius={[4, 4, 0, 0]} />
                  <Bar dataKey="deliveredQty" name="Delivered Qty" fill="#10b981" radius={[4, 4, 0, 0]} />
                </BarChart>
              </ResponsiveContainer>
            </div>
          </div>

          <div className="bg-white rounded-2xl border border-slate-200 shadow-sm p-6 flex flex-col">
            <div className="mb-2">
              <h2 className="text-lg font-semibold text-slate-800">Order Distribution</h2>
              <p className="text-sm text-slate-500">Total Volume by Buyer</p>
            </div>
            <div className="h-[350px] w-full flex-1 mt-4">
              <ResponsiveContainer width="100%" height="100%">
                <PieChart>
                  <Pie
                    data={chartDataBuyer}
                    cx="50%"
                    cy="50%"
                    innerRadius={70}
                    outerRadius={110}
                    paddingAngle={5}
                    dataKey="qty"
                    nameKey="buyer"
                  >
                    {chartDataBuyer.map((entry, index) => (
                      <Cell key={`cell-${index}`} fill={COLORS[index % COLORS.length]} />
                    ))}
                  </Pie>
                  <Tooltip 
                    contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 10px 15px -3px rgb(0 0 0 / 0.1)' }}
                  />
                  <Legend verticalAlign="bottom" height={36} />
                </PieChart>
              </ResponsiveContainer>
            </div>
          </div>
        </div>

        {/* Detailed Data Table */}
        <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
          <div className="p-6 border-b border-slate-200">
            <h2 className="text-lg font-semibold text-slate-800">Fulfillment Details</h2>
            <p className="text-sm text-slate-500">Item-level breakdown of the uploaded dataset.</p>
          </div>
          <div className="overflow-x-auto max-h-[500px]">
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

function MetricCard({ title, value, icon }: { title: string, value: string, icon: React.ReactNode }) {
  return (
    <div className="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm flex items-center justify-between">
      <div>
        <p className="text-sm font-medium text-slate-500 mb-1">{title}</p>
        <p className="text-2xl font-bold text-slate-900">{value}</p>
      </div>
      <div className="p-3 bg-slate-50 rounded-xl">
        {icon}
      </div>
    </div>
  );
}

function cn(...classes: (string | undefined | null | false)[]) {
  return classes.filter(Boolean).join(' ');
}
