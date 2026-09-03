import React, { useMemo, useState, useEffect } from 'react';
import { ProductionOrder } from './Dashboard';
import { 
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip as RechartsTooltip, ResponsiveContainer,
  Cell, Legend, ComposedChart, Line
} from 'recharts';
import { Package, Calendar, Briefcase, Activity, CheckCircle, AlertCircle, AlertTriangle, Search, X } from 'lucide-react';

interface KPIViewProps {
  data: ProductionOrder[];
}

export type ScheduleDetail = {
  scheduleNo: string;
  coQty: number;
  sewOutQty: number;
  pendingQty: number;
};

export type VPODetail = {
  vpo: string;
  buyer: string;
  style: string;
  coQty: number;
  sewOutQty: number;
  pendingQty: number;
  progress: number;
  schedules: ScheduleDetail[];
};

const COLORS = ['#10b981', '#0ea5e9', '#6366f1', '#8b5cf6', '#d946ef', '#f43f5e', '#f59e0b'];

// Helper to get current ISO week
function getCurrentISOWeek() {
  const date = new Date();
  const dayNum = date.getUTCDay() || 7;
  date.setUTCDate(date.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(date.getUTCFullYear(),0,1));
  return Math.ceil((((date.getTime() - yearStart.getTime()) / 86400000) + 1)/7).toString();
}

export function KPIView({ data }: KPIViewProps) {
  const [selectedWeek, setSelectedWeek] = useState<string>('All');
  const [selectedBuyer, setSelectedBuyer] = useState<string>('All');
  const [vpoSearchQuery, setVpoSearchQuery] = useState('');
  const [selectedVPO, setSelectedVPO] = useState<VPODetail | null>(null);
  
  // Extract unique weeks & buyers
  const availableWeeks = useMemo(() => {
    const weeks = new Set<string>();
    data.forEach(d => {
      if (d.weekNo) weeks.add(d.weekNo.toString());
    });
    return Array.from(weeks).sort((a, b) => Number(a) - Number(b));
  }, [data]);

  const availableBuyers = useMemo(() => {
    const buyers = new Set<string>();
    data.forEach(d => {
      if (d.buyer) buyers.add(d.buyer);
    });
    return Array.from(buyers).sort((a, b) => a.localeCompare(b));
  }, [data]);

  useEffect(() => {
    const currentWeekStr = getCurrentISOWeek();
    if (availableWeeks.length > 0 && selectedWeek === 'All') {
      if (availableWeeks.includes(currentWeekStr)) {
        setSelectedWeek(currentWeekStr);
      } else {
        setSelectedWeek(availableWeeks[0]);
      }
    }
  }, [availableWeeks, selectedWeek]);

  const kpiData = useMemo(() => {
    let filteredData = data;
    if (selectedWeek !== 'All') {
      filteredData = filteredData.filter(d => d.weekNo?.toString() === selectedWeek);
    }
    if (selectedBuyer !== 'All') {
      filteredData = filteredData.filter(d => d.buyer === selectedBuyer);
    }
    
    let totalCOQty = 0;
    let totalSewIn = 0;
    let totalSewOut = 0;
    let totalRejects = 0;

    const buyerMap = new Map<string, { name: string; coQty: number; sewOutQty: number; pendingQty: number }>();
    const vpoMap = new Map<string, { vpo: string; buyer: string; style: string; coQty: number; sewOutQty: number; schedules: Map<string, ScheduleDetail> }>();

    filteredData.forEach(item => {
      const co = Number(item.coQty) || 0;
      const sewIn = Number(item.cumSewInQty) || 0;
      const sewOut = Number(item.cumSewOutQty) || 0;
      const rej = Number(item.cumSewOutRejQty) || 0;
      
      totalCOQty += co;
      totalSewIn += sewIn;
      totalSewOut += sewOut;
      totalRejects += rej;

      const buyer = item.buyer || 'Unknown';
      if (!buyerMap.has(buyer)) {
        buyerMap.set(buyer, { name: buyer, coQty: 0, sewOutQty: 0, pendingQty: 0 });
      }
      const b = buyerMap.get(buyer)!;
      b.coQty += co;
      b.sewOutQty += sewOut;
      b.pendingQty = Math.max(0, b.coQty - b.sewOutQty);

      if (item.vpoNo) {
        if (!vpoMap.has(item.vpoNo)) {
          vpoMap.set(item.vpoNo, { vpo: item.vpoNo, buyer: item.buyer || '', style: item.styleNo || '', coQty: 0, sewOutQty: 0, schedules: new Map() });
        }
        const v = vpoMap.get(item.vpoNo)!;
        v.coQty += co;
        v.sewOutQty += sewOut;
        if (item.scheduleNo) {
          if (!v.schedules.has(item.scheduleNo)) {
            v.schedules.set(item.scheduleNo, { scheduleNo: item.scheduleNo, coQty: 0, sewOutQty: 0, pendingQty: 0 });
          }
          const s = v.schedules.get(item.scheduleNo)!;
          s.coQty += co;
          s.sewOutQty += sewOut;
          s.pendingQty = Math.max(0, s.coQty - s.sewOutQty);
        }
      }
    });

    const totalPending = Math.max(0, totalCOQty - totalSewOut);
    const overallProgress = totalCOQty > 0 ? (totalSewOut / totalCOQty) * 100 : 0;
    const rejectionRate = totalSewOut > 0 ? (totalRejects / (totalSewOut + totalRejects)) * 100 : 0;

    const buyerData = Array.from(buyerMap.values());
    
    // Top Buyers by Completed (Sew Out)
    const topCompletedBuyers = [...buyerData]
      .sort((a, b) => b.sewOutQty - a.sewOutQty)
      .slice(0, 7);

    // Top Buyers by Pending (CO - Sew Out)
    const topPendingBuyers = [...buyerData]
      .sort((a, b) => b.pendingQty - a.pendingQty)
      .slice(0, 7);

    const topPendingVPOs: VPODetail[] = Array.from(vpoMap.values())
      .map(v => ({
        ...v,
        pendingQty: Math.max(0, v.coQty - v.sewOutQty),
        progress: v.coQty > 0 ? (v.sewOutQty / v.coQty) * 100 : 0,
        schedules: Array.from(v.schedules.values()).sort((a, b) => b.pendingQty - a.pendingQty)
      }))
      .filter(v => v.pendingQty > 0)
      .sort((a, b) => b.pendingQty - a.pendingQty)
      .slice(0, 50); // Top 50 Pending VPOs

    return {
      totalCOQty,
      totalSewIn,
      totalSewOut,
      totalRejects,
      totalPending,
      overallProgress,
      rejectionRate,
      topCompletedBuyers,
      topPendingBuyers,
      topPendingVPOs
    };
  }, [data, selectedWeek, selectedBuyer]);

  const filteredTopPendingVPOs = useMemo(() => {
    const query = vpoSearchQuery.toLowerCase().trim();
    if (!query) return kpiData.topPendingVPOs;
    return kpiData.topPendingVPOs.filter(vpo => {
      if (String(vpo.vpo).toLowerCase().includes(query)) return true;
      if (vpo.schedules.some(s => String(s.scheduleNo).toLowerCase().includes(query))) return true;
      return false;
    });
  }, [kpiData.topPendingVPOs, vpoSearchQuery]);

  const groupedPendingVPOs = useMemo(() => {
    const groups = new Map<string, VPODetail[]>();
    filteredTopPendingVPOs.forEach(vpo => {
      const buyer = vpo.buyer || 'Unknown Buyer';
      if (!groups.has(buyer)) {
        groups.set(buyer, []);
      }
      groups.get(buyer)!.push(vpo);
    });
    
    // Sort buyers by total pending qty descending
    return Array.from(groups.entries()).sort((a, b) => {
       const aPending = a[1].reduce((sum, v) => sum + v.pendingQty, 0);
       const bPending = b[1].reduce((sum, v) => sum + v.pendingQty, 0);
       return bPending - aPending;
    });
  }, [filteredTopPendingVPOs]);

  return (
    <div className="space-y-6 animate-in fade-in duration-500">
      
      {/* Control Panel */}
      <div className="bg-white rounded-2xl p-5 border border-slate-200 shadow-sm flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div>
          <h2 className="text-xl font-bold text-slate-800 flex items-center gap-2">
            <Activity className="w-5 h-5 text-indigo-500" />
            Production (Sewing) Analytics
          </h2>
          <p className="text-sm text-slate-500 mt-1">Real-time sewing execution and performance tracking</p>
        </div>
        
        <div className="flex flex-col sm:flex-row items-center gap-3">
          <div className="flex items-center bg-slate-50 border border-slate-200 rounded-lg p-1 w-full sm:w-auto">
            <Briefcase className="w-4 h-4 text-slate-400 ml-2" />
            <select 
              value={selectedBuyer} 
              onChange={(e) => setSelectedBuyer(e.target.value)}
              className="bg-transparent border-none text-sm font-semibold focus:outline-none focus:ring-0 text-slate-700 py-1.5 pl-2 pr-8 w-full cursor-pointer"
            >
              <option value="All">All Buyers</option>
              {availableBuyers.map(b => (
                <option key={b} value={b}>{b}</option>
              ))}
            </select>
          </div>
          <div className="flex items-center bg-slate-50 border border-slate-200 rounded-lg p-1 w-full sm:w-auto">
            <Calendar className="w-4 h-4 text-slate-400 ml-2" />
            <select 
              value={selectedWeek} 
              onChange={(e) => setSelectedWeek(e.target.value)}
              className="bg-transparent border-none text-sm font-semibold focus:outline-none focus:ring-0 text-slate-700 py-1.5 pl-2 pr-8 w-full cursor-pointer"
            >
              <option value="All">All Weeks</option>
              {availableWeeks.map(w => (
                <option key={w} value={w}>Week {w}</option>
              ))}
            </select>
          </div>
        </div>
      </div>

      {/* KPI Top Level Cards */}
      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
        
        {/* Order Qty */}
        <div className="bg-white rounded-2xl p-5 border border-slate-200 shadow-sm relative overflow-hidden group">
          <div className="absolute -right-4 -top-4 w-16 h-16 bg-blue-50 rounded-full group-hover:scale-150 transition-transform duration-500"></div>
          <div className="relative z-10">
            <div className="flex items-center justify-between mb-2">
              <p className="text-xs font-bold text-slate-500 uppercase tracking-wider">Total CO Qty</p>
              <Package className="w-4 h-4 text-blue-500" />
            </div>
            <p className="text-3xl font-black text-slate-900">{kpiData.totalCOQty.toLocaleString()}</p>
          </div>
        </div>
        
        {/* Completed (Sew Out) Qty */}
        <div className="bg-white rounded-2xl p-5 border border-slate-200 shadow-sm relative overflow-hidden group">
          <div className="absolute -right-4 -top-4 w-16 h-16 bg-emerald-50 rounded-full group-hover:scale-150 transition-transform duration-500"></div>
          <div className="relative z-10">
            <div className="flex items-center justify-between mb-2">
              <p className="text-xs font-bold text-slate-500 uppercase tracking-wider">Completed (Sew Out)</p>
              <CheckCircle className="w-4 h-4 text-emerald-500" />
            </div>
            <p className="text-3xl font-black text-slate-900">{kpiData.totalSewOut.toLocaleString()}</p>
          </div>
        </div>

        {/* Pending Qty */}
        <div className="bg-white rounded-2xl p-5 border border-slate-200 shadow-sm relative overflow-hidden group">
          <div className="absolute -right-4 -top-4 w-16 h-16 bg-amber-50 rounded-full group-hover:scale-150 transition-transform duration-500"></div>
          <div className="relative z-10">
            <div className="flex items-center justify-between mb-2">
              <p className="text-xs font-bold text-slate-500 uppercase tracking-wider">Pending Production</p>
              <AlertCircle className="w-4 h-4 text-amber-500" />
            </div>
            <p className="text-3xl font-black text-slate-900">{kpiData.totalPending.toLocaleString()}</p>
          </div>
        </div>

        {/* Rejection */}
        <div className="bg-white rounded-2xl p-5 border border-slate-200 shadow-sm relative overflow-hidden group">
          <div className="absolute -right-4 -top-4 w-16 h-16 bg-rose-50 rounded-full group-hover:scale-150 transition-transform duration-500"></div>
          <div className="relative z-10">
            <div className="flex items-center justify-between mb-2">
              <p className="text-xs font-bold text-slate-500 uppercase tracking-wider">Sew Out Rejects</p>
              <AlertTriangle className="w-4 h-4 text-rose-500" />
            </div>
            <p className="text-3xl font-black text-slate-900">{kpiData.totalRejects.toLocaleString()} <span className="text-sm font-semibold text-rose-500 ml-1">({kpiData.rejectionRate.toFixed(1)}%)</span></p>
          </div>
        </div>

      </div>

      {/* Production Progress Bar */}
      <div className="bg-gradient-to-r from-slate-900 to-slate-800 rounded-2xl p-6 shadow-md">
        <div className="flex justify-between items-end mb-3">
          <div>
            <p className="text-xs font-bold text-slate-400 uppercase tracking-wider mb-1">Overall Sewing Completion Progress</p>
            <h3 className="text-2xl font-black text-white">{kpiData.overallProgress.toFixed(1)}% <span className="text-sm font-medium text-slate-400 ml-2 font-normal">Completed</span></h3>
          </div>
          <div className="text-right">
            <p className="text-sm text-slate-300"><span className="font-bold text-white">{kpiData.totalSewIn.toLocaleString()}</span> Items currently Sewn In</p>
          </div>
        </div>
        <div className="w-full bg-slate-700 rounded-full h-3 overflow-hidden">
          <div 
            className="bg-emerald-500 h-3 rounded-full transition-all duration-1000" 
            style={{ width: `${Math.min(100, kpiData.overallProgress)}%` }}
          ></div>
        </div>
      </div>

      {/* Charts Section */}
      <div className="grid grid-cols-1 xl:grid-cols-2 gap-6">
        
        {/* Top Completed Buyers */}
        <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm flex flex-col">
          <div className="mb-6">
            <h3 className="text-base font-bold text-slate-800 flex items-center gap-2">
               Top Completed Buyers (Sew Out)
            </h3>
            <p className="text-xs text-slate-500">Highest volume of successfully sewn out items</p>
          </div>
          <div className="h-[320px] w-full flex-grow">
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={kpiData.topCompletedBuyers} layout="vertical" margin={{ top: 0, right: 30, left: 40, bottom: 0 }}>
                <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke="#f1f5f9" />
                <XAxis type="number" axisLine={false} tickLine={false} tick={{ fontSize: 11, fill: '#64748b' }} tickFormatter={(val) => `${val / 1000}k`} />
                <YAxis dataKey="name" type="category" axisLine={false} tickLine={false} tick={{ fontSize: 11, fontWeight: 600, fill: '#334155' }} width={90} />
                <RechartsTooltip 
                  cursor={{ fill: '#f8fafc' }}
                  contentStyle={{ borderRadius: '12px', border: '1px solid #e2e8f0', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }}
                />
                <Bar dataKey="sewOutQty" name="Completed (Sew Out)" fill="#10b981" radius={[0, 6, 6, 0]} barSize={24}>
                  {kpiData.topCompletedBuyers.map((entry, index) => (
                    <Cell key={`cell-${index}`} fill={COLORS[index % COLORS.length]} />
                  ))}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>

        {/* Top Pending Buyers */}
        <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm flex flex-col">
          <div className="mb-6">
            <h3 className="text-base font-bold text-slate-800 flex items-center gap-2">
               Pending vs Completed (Top Pending Buyers)
            </h3>
            <p className="text-xs text-slate-500">Proportion of completed and remaining production per buyer</p>
          </div>
          <div className="h-[320px] w-full flex-grow">
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={kpiData.topPendingBuyers} layout="vertical" margin={{ top: 0, right: 30, left: 40, bottom: 0 }}>
                <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke="#f1f5f9" />
                <XAxis type="number" axisLine={false} tickLine={false} tick={{ fontSize: 11, fill: '#64748b' }} tickFormatter={(val) => `${val / 1000}k`} />
                <YAxis dataKey="name" type="category" axisLine={false} tickLine={false} tick={{ fontSize: 11, fontWeight: 600, fill: '#334155' }} width={90} />
                <RechartsTooltip 
                  cursor={{ fill: '#f8fafc' }}
                  contentStyle={{ borderRadius: '12px', border: '1px solid #e2e8f0', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }}
                />
                <Legend wrapperStyle={{ fontSize: '12px', paddingTop: '10px' }} />
                <Bar dataKey="sewOutQty" name="Completed (Sew Out)" stackId="a" fill="#10b981" barSize={28} />
                <Bar dataKey="pendingQty" name="Pending Qty" stackId="a" fill="#f59e0b" radius={[0, 4, 4, 0]} barSize={28} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>

      </div>

      {/* Pending VPOs List */}
      <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm">
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 mb-6">
          <div>
            <h3 className="text-base font-bold text-slate-800 flex items-center gap-2">
               <AlertCircle className="w-5 h-5 text-amber-500" /> Action Required: Critical Pending VPOs ({filteredTopPendingVPOs.length})
            </h3>
            <p className="text-xs text-slate-500 mt-1">Individual VPOs with the highest remaining production volume</p>
          </div>
          <div className="relative w-full md:w-64">
            <div className="absolute inset-y-0 left-0 pl-3 flex items-center pointer-events-none">
              <Search className="h-4 w-4 text-slate-400" />
            </div>
            <input
              type="text"
              placeholder="Search VPO or Schedule..."
              value={vpoSearchQuery}
              onChange={(e) => setVpoSearchQuery(e.target.value)}
              className="block w-full pl-9 pr-3 py-2 border border-slate-200 rounded-lg text-sm bg-slate-50 focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 transition-colors placeholder:text-slate-400"
            />
          </div>
        </div>
        
        <div className="space-y-6 max-h-[500px] overflow-y-auto pr-2 custom-scrollbar [&::-webkit-scrollbar]:w-2 [&::-webkit-scrollbar-track]:bg-slate-100 [&::-webkit-scrollbar-thumb]:bg-slate-300 [&::-webkit-scrollbar-thumb]:rounded-full hover:[&::-webkit-scrollbar-thumb]:bg-slate-400">
          {groupedPendingVPOs.map(([buyer, vpos], groupIdx) => (
            <div key={groupIdx} className="space-y-3">
              <h4 className="text-sm font-bold text-slate-700 bg-slate-100 px-3 py-1.5 rounded-lg inline-block">
                {buyer} <span className="text-slate-500 font-normal text-xs ml-1">({vpos.length} VPOs)</span>
              </h4>
              <div className="space-y-3 pl-2 border-l-2 border-slate-100">
                {vpos.map((vpo, idx) => (
                  <div 
                    key={idx} 
                    onClick={() => setSelectedVPO(vpo)}
                    className="flex flex-col sm:flex-row sm:items-center gap-4 p-4 rounded-xl border border-slate-100 bg-slate-50/50 hover:bg-slate-50 transition-colors cursor-pointer"
                  >
                     <div className="w-full sm:w-1/4">
                        <p className="text-sm font-bold text-slate-800">{vpo.vpo}</p>
                        <p className="text-xs text-slate-500">Style: {vpo.style}</p>
                        {vpo.schedules.length > 0 && (
                          <p className="text-[10px] text-slate-400 mt-0.5 truncate">Sch: {vpo.schedules.map(s => s.scheduleNo).join(', ')}</p>
                        )}
                     </div>
                     
                     <div className="w-full sm:w-2/4 flex items-center gap-4">
                        <div className="flex-grow bg-slate-200 rounded-full h-2.5 overflow-hidden">
                           <div 
                             className={`h-2.5 rounded-full ${vpo.progress < 30 ? 'bg-rose-500' : vpo.progress < 70 ? 'bg-amber-500' : 'bg-emerald-500'}`}
                             style={{ width: `${Math.min(100, vpo.progress)}%` }}
                           ></div>
                        </div>
                        <span className="text-xs font-bold text-slate-700 min-w-[3rem] text-right">{vpo.progress.toFixed(0)}%</span>
                     </div>
                     
                     <div className="w-full sm:w-1/4 text-left sm:text-right flex flex-col sm:items-end">
                        <p className="text-sm font-bold text-amber-600">{vpo.pendingQty.toLocaleString()} Pending</p>
                        <p className="text-xs text-slate-400">{vpo.sewOutQty.toLocaleString()} / {vpo.coQty.toLocaleString()} Completed</p>
                     </div>
                  </div>
                ))}
              </div>
            </div>
          ))}
          {groupedPendingVPOs.length === 0 && (
            <div className="text-center py-8 text-slate-500 text-sm">
              No pending VPOs found matching your search.
            </div>
          )}
        </div>
      </div>

      {/* VPO Details Modal */}
      {selectedVPO && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/50 backdrop-blur-sm animate-in fade-in duration-200">
          <div className="bg-white rounded-2xl shadow-xl w-full max-w-2xl overflow-hidden flex flex-col max-h-[90vh] animate-in zoom-in-95 duration-200">
            <div className="flex items-center justify-between p-5 border-b border-slate-100">
              <div>
                <h3 className="text-lg font-bold text-slate-800 flex items-center gap-2">
                  <Package className="w-5 h-5 text-indigo-500" />
                  VPO Details: {selectedVPO.vpo}
                </h3>
                <p className="text-sm text-slate-500 mt-1">{selectedVPO.buyer} • Style {selectedVPO.style}</p>
              </div>
              <button 
                onClick={() => setSelectedVPO(null)}
                className="p-2 text-slate-400 hover:text-slate-600 hover:bg-slate-100 rounded-lg transition-colors"
              >
                <X className="w-5 h-5" />
              </button>
            </div>
            
            <div className="p-5 overflow-y-auto">
              <div className="grid grid-cols-3 gap-4 mb-6">
                <div className="bg-slate-50 p-4 rounded-xl border border-slate-100">
                  <p className="text-xs font-semibold text-slate-500 uppercase tracking-wider mb-1">Total Order</p>
                  <p className="text-2xl font-bold text-slate-800">{selectedVPO.coQty.toLocaleString()}</p>
                </div>
                <div className="bg-emerald-50 p-4 rounded-xl border border-emerald-100">
                  <p className="text-xs font-semibold text-emerald-600 uppercase tracking-wider mb-1">Sew Out</p>
                  <p className="text-2xl font-bold text-emerald-700">{selectedVPO.sewOutQty.toLocaleString()}</p>
                </div>
                <div className="bg-amber-50 p-4 rounded-xl border border-amber-100">
                  <p className="text-xs font-semibold text-amber-600 uppercase tracking-wider mb-1">Pending</p>
                  <p className="text-2xl font-bold text-amber-700">{selectedVPO.pendingQty.toLocaleString()}</p>
                </div>
              </div>

              <h4 className="text-sm font-bold text-slate-800 mb-3 border-b border-slate-100 pb-2">Schedule Breakdown</h4>
              
              {selectedVPO.schedules.length > 0 ? (
                <div className="space-y-3">
                  {selectedVPO.schedules.map((sch, i) => (
                    <div key={i} className="flex flex-col sm:flex-row sm:items-center justify-between p-3 rounded-lg border border-slate-100 bg-white shadow-sm gap-2">
                      <div className="flex items-center gap-3">
                        <div className="bg-indigo-50 text-indigo-600 p-2 rounded-md">
                          <Calendar className="w-4 h-4" />
                        </div>
                        <div>
                          <p className="text-sm font-bold text-slate-700">Sch: {sch.scheduleNo}</p>
                        </div>
                      </div>
                      
                      <div className="flex items-center gap-6 text-sm">
                        <div className="text-right">
                          <p className="text-xs text-slate-400 uppercase">Order</p>
                          <p className="font-semibold text-slate-700">{sch.coQty.toLocaleString()}</p>
                        </div>
                        <div className="text-right">
                          <p className="text-xs text-emerald-500 uppercase">Sew Out</p>
                          <p className="font-semibold text-emerald-600">{sch.sewOutQty.toLocaleString()}</p>
                        </div>
                        <div className="text-right w-16">
                          <p className="text-xs text-amber-500 uppercase">Pending</p>
                          <p className="font-bold text-amber-600">{sch.pendingQty.toLocaleString()}</p>
                        </div>
                      </div>
                    </div>
                  ))}
                </div>
              ) : (
                <p className="text-sm text-slate-500 text-center py-4">No schedule details available for this VPO.</p>
              )}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
