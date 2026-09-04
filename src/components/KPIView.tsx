import React, { useMemo, useState, useEffect, useRef } from 'react';
import { ProductionOrder } from './Dashboard';
import { useVirtualizer } from '@tanstack/react-virtual';
import { MultiSelectDropdown } from './Dashboard';
import * as XLSX from 'xlsx';
import { 
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip as RechartsTooltip, ResponsiveContainer,
  Cell, Legend, ComposedChart, Line
} from 'recharts';
import { Package, Calendar, Briefcase, Activity, CheckCircle, AlertCircle, AlertTriangle, Search, X, ChevronDown, ChevronRight, Download } from 'lucide-react';

interface KPIViewProps {
  data: ProductionOrder[];
}

export type ScheduleDetail = {
  scheduleNo: string;
  coQty: number;
  sewOutQty: number;
  pendingQty: number;
  rejects: number;
};

export type VPODetail = {
  vpo: string;
  buyer: string;
  style: string;
  coQty: number;
  sewOutQty: number;
  pendingQty: number;
  rejects: number;
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
  const [selectedWeeks, setSelectedWeeks] = useState<string[]>([]);
  const [selectedBuyers, setSelectedBuyers] = useState<string[]>([]);
  const [vpoSearchQuery, setVpoSearchQuery] = useState('');
  const [selectedVPO, setSelectedVPO] = useState<VPODetail | null>(null);
  const [expandedBuyer, setExpandedBuyer] = useState<string | null>(null);
  const [fgInModal, setFgInModal] = useState<{
    isOpen: boolean;
    buyer: string;
    title: string;
    type: 'styles' | 'vpos' | 'schedules';
    items: string[];
  }>({ isOpen: false, buyer: '', title: '', type: 'styles', items: [] });
  const [fgInSearch, setFgInSearch] = useState('');

  const fgInFilteredItems = useMemo(() => {
    if (!fgInModal.isOpen) return [];
    const searchLower = fgInSearch.toLowerCase();
    return fgInModal.items.filter(item => String(item).toLowerCase().includes(searchLower));
  }, [fgInModal.isOpen, fgInModal.items, fgInSearch]);

  const fgInListRef = useRef<HTMLDivElement>(null);
  
  const fgInVirtualizer = useVirtualizer({
    count: fgInFilteredItems.length,
    getScrollElement: () => fgInListRef.current,
    estimateSize: () => 45,
    overscan: 10,
  });
  
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

  const kpiData = useMemo(() => {
    let filteredData = data;
    if (selectedWeeks.length > 0) {
      filteredData = filteredData.filter(d => d.weekNo && selectedWeeks.includes(d.weekNo.toString()));
    }
    if (selectedBuyers.length > 0) {
      filteredData = filteredData.filter(d => d.buyer && selectedBuyers.includes(d.buyer));
    }
    
    let totalCOQty = 0;
    let totalSewIn = 0;
    let totalSewOut = 0;
    let totalRejects = 0;

    const buyerMap = new Map<string, { name: string; coQty: number; sewOutQty: number; pendingQty: number }>();
    const vpoMap = new Map<string, { vpo: string; buyer: string; style: string; coQty: number; sewOutQty: number; schedules: Map<string, ScheduleDetail> }>();
    const fgInMap = new Map<string, { buyer: string; styles: Set<string>; vpos: Set<string>; schedules: Set<string> }>();

    filteredData.forEach(item => {
      const co = Number(item.coQty) || 0;
      const sewIn = Number(item.cumSewInQty) || 0;
      const sewOut = Number(item.cumSewOutQty) || 0;
      const rej = Number(item.cumSewOutRejQty) || 0;
      const cumCTN = Number(item.cumCTNQty) || 0;
      const deliv = Number(item.deliveredQty) || 0;
      
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

      if (cumCTN > 0 && deliv === 0) {
        if (!fgInMap.has(buyer)) {
          fgInMap.set(buyer, { buyer, styles: new Set(), vpos: new Set(), schedules: new Set() });
        }
        const f = fgInMap.get(buyer)!;
        if (item.styleNo) f.styles.add(item.styleNo);
        if (item.vpoNo) f.vpos.add(item.vpoNo);
        if (item.scheduleNo) f.schedules.add(item.scheduleNo);
      }

      if (item.vpoNo) {
        if (!vpoMap.has(item.vpoNo)) {
          vpoMap.set(item.vpoNo, { vpo: item.vpoNo, buyer: item.buyer || '', style: item.styleNo || '', coQty: 0, sewOutQty: 0, rejects: 0, schedules: new Map() });
        }
        const v = vpoMap.get(item.vpoNo)!;
        v.coQty += co;
        v.sewOutQty += sewOut;
        v.rejects += rej;
        if (item.scheduleNo) {
          if (!v.schedules.has(item.scheduleNo)) {
            v.schedules.set(item.scheduleNo, { scheduleNo: item.scheduleNo, coQty: 0, sewOutQty: 0, pendingQty: 0, rejects: 0 });
          }
          const s = v.schedules.get(item.scheduleNo)!;
          s.coQty += co;
          s.sewOutQty += sewOut;
          s.rejects += rej;
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

    let vpoCountAll = 0, schCountAll = 0;
    let vpoCountSewOut = 0, schCountSewOut = 0;
    let vpoCountPending = 0, schCountPending = 0;
    let vpoCountRejects = 0, schCountRejects = 0;

    for (const v of vpoMap.values()) {
      vpoCountAll++;
      if (v.sewOutQty > 0) vpoCountSewOut++;
      if (v.coQty > v.sewOutQty) vpoCountPending++;
      if (v.rejects > 0) vpoCountRejects++;
      
      for (const s of v.schedules.values()) {
        schCountAll++;
        if (s.sewOutQty > 0) schCountSewOut++;
        if (s.pendingQty > 0) schCountPending++;
        if (s.rejects > 0) schCountRejects++;
      }
    }

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

    const fgInStats = Array.from(fgInMap.values()).map(f => ({
      buyer: f.buyer,
      styleCount: f.styles.size,
      vpoCount: f.vpos.size,
      scheduleCount: f.schedules.size,
      styles: Array.from(f.styles).sort(),
      vpos: Array.from(f.vpos).sort(),
      schedules: Array.from(f.schedules).sort()
    })).sort((a, b) => b.vpoCount - a.vpoCount);

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
      topPendingVPOs,
      fgInStats,
      counts: {
        all: { vpo: vpoCountAll, sch: schCountAll },
        sewOut: { vpo: vpoCountSewOut, sch: schCountSewOut },
        pending: { vpo: vpoCountPending, sch: schCountPending },
        rejects: { vpo: vpoCountRejects, sch: schCountRejects }
      }
    };
  }, [data, selectedWeeks, selectedBuyers]);

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

  useEffect(() => {
    if (groupedPendingVPOs.length > 0 && (!expandedBuyer || !groupedPendingVPOs.some(g => g[0] === expandedBuyer))) {
      setExpandedBuyer(groupedPendingVPOs[0][0]);
    } else if (groupedPendingVPOs.length === 0) {
      setExpandedBuyer(null);
    }
  }, [groupedPendingVPOs]);

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
        
        <div className="flex flex-wrap sm:flex-nowrap items-center gap-3 w-full md:w-auto">
          <MultiSelectDropdown 
            label="All Buyers"
            options={availableBuyers}
            selectedValues={selectedBuyers}
            onChange={setSelectedBuyers}
          />
          <MultiSelectDropdown 
            label="All Weeks"
            options={availableWeeks}
            selectedValues={selectedWeeks}
            onChange={setSelectedWeeks}
          />
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
            <p className="text-3xl font-black text-slate-900">{kpiData.totalCOQty?.toLocaleString() || "0"}</p>
            <div className="mt-2.5 text-[10px] font-bold text-slate-400 bg-slate-50 px-2.5 py-1.5 rounded-lg border border-slate-100 flex items-center justify-between">
              <span>VPOs: <span className="text-slate-700">{kpiData.counts.all.vpo}</span></span>
              <span className="w-[1px] h-3 bg-slate-200"></span>
              <span>Schedules: <span className="text-slate-700">{kpiData.counts.all.sch}</span></span>
            </div>
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
            <p className="text-3xl font-black text-slate-900">{kpiData.totalSewOut?.toLocaleString() || "0"}</p>
            <div className="mt-2.5 text-[10px] font-bold text-slate-400 bg-slate-50 px-2.5 py-1.5 rounded-lg border border-slate-100 flex items-center justify-between">
              <span>VPOs: <span className="text-slate-700">{kpiData.counts.sewOut.vpo}</span></span>
              <span className="w-[1px] h-3 bg-slate-200"></span>
              <span>Schedules: <span className="text-slate-700">{kpiData.counts.sewOut.sch}</span></span>
            </div>
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
            <p className="text-3xl font-black text-slate-900">{kpiData.totalPending?.toLocaleString() || "0"}</p>
            <div className="mt-2.5 text-[10px] font-bold text-slate-400 bg-slate-50 px-2.5 py-1.5 rounded-lg border border-slate-100 flex items-center justify-between">
              <span>VPOs: <span className="text-slate-700">{kpiData.counts.pending.vpo}</span></span>
              <span className="w-[1px] h-3 bg-slate-200"></span>
              <span>Schedules: <span className="text-slate-700">{kpiData.counts.pending.sch}</span></span>
            </div>
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
            <p className="text-3xl font-black text-slate-900">{kpiData.totalRejects?.toLocaleString() || "0"} <span className="text-sm font-semibold text-rose-500 ml-1">({kpiData.rejectionRate.toFixed(1)}%)</span></p>
            <div className="mt-2.5 text-[10px] font-bold text-slate-400 bg-slate-50 px-2.5 py-1.5 rounded-lg border border-slate-100 flex items-center justify-between">
              <span>VPOs: <span className="text-slate-700">{kpiData.counts.rejects.vpo}</span></span>
              <span className="w-[1px] h-3 bg-slate-200"></span>
              <span>Schedules: <span className="text-slate-700">{kpiData.counts.rejects.sch}</span></span>
            </div>
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
            <p className="text-sm text-slate-300"><span className="font-bold text-white">{kpiData.totalSewIn?.toLocaleString() || "0"}</span> Items currently Sewn In</p>
          </div>
        </div>
        <div className="w-full bg-slate-700 rounded-full h-3 overflow-hidden">
          <div 
            className="bg-emerald-500 h-3 rounded-full transition-all duration-1000" 
            style={{ width: `${Math.min(100, kpiData.overallProgress)}%` }}
          ></div>
        </div>
      </div>

      {/* FG IN Section */}
      <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm flex flex-col">
        <div className="mb-6">
          <h3 className="text-base font-bold text-slate-800 flex items-center gap-2">
            <Package className="w-5 h-5 text-indigo-500" />
            FG IN
          </h3>
          <p className="text-xs text-slate-500 mt-1">Buyer wise Style No, VPO, and Schedule counts (Cum CTN Qty &gt; 0 &amp; Delivered Qty = 0)</p>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse">
            <thead>
              <tr className="border-b border-slate-200 bg-slate-50">
                <th className="py-3 px-4 text-xs font-semibold text-slate-500 uppercase tracking-wider">Buyer</th>
                <th className="py-3 px-4 text-xs font-semibold text-slate-500 uppercase tracking-wider text-right">Style Count</th>
                <th className="py-3 px-4 text-xs font-semibold text-slate-500 uppercase tracking-wider text-right">VPO Count</th>
                <th className="py-3 px-4 text-xs font-semibold text-slate-500 uppercase tracking-wider text-right">Schedule Count</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {kpiData.fgInStats.map((stat, idx) => (
                <tr key={idx} className="hover:bg-slate-50/50 transition-colors">
                  <td className="py-3 px-4 text-sm font-medium text-slate-700">{stat.buyer}</td>
                  <td 
                    className="py-3 px-4 text-sm font-semibold text-indigo-600 text-right cursor-pointer hover:underline"
                    onClick={() => setFgInModal({ isOpen: true, buyer: stat.buyer, title: 'Styles', type: 'styles', items: stat.styles })}
                  >
                    {stat.styleCount?.toLocaleString() || "0"}
                  </td>
                  <td 
                    className="py-3 px-4 text-sm font-semibold text-emerald-600 text-right cursor-pointer hover:underline"
                    onClick={() => setFgInModal({ isOpen: true, buyer: stat.buyer, title: 'VPOs', type: 'vpos', items: stat.vpos })}
                  >
                    {stat.vpoCount?.toLocaleString() || "0"}
                  </td>
                  <td 
                    className="py-3 px-4 text-sm font-semibold text-amber-600 text-right cursor-pointer hover:underline"
                    onClick={() => setFgInModal({ isOpen: true, buyer: stat.buyer, title: 'Schedules', type: 'schedules', items: stat.schedules })}
                  >
                    {stat.scheduleCount?.toLocaleString() || "0"}
                  </td>
                </tr>
              ))}
              {kpiData.fgInStats.length === 0 && (
                <tr>
                  <td colSpan={4} className="py-8 text-center text-sm text-slate-500">
                    No records found matching FG IN criteria.
                  </td>
                </tr>
              )}
            </tbody>
            {kpiData.fgInStats.length > 0 && (
              <tfoot className="bg-slate-50 font-bold border-t-2 border-slate-200">
                <tr>
                  <td className="py-3 px-4 text-sm text-slate-800 uppercase tracking-wider">Total</td>
                  <td className="py-3 px-4 text-sm text-indigo-700 text-right">
                    {kpiData.fgInStats.reduce((sum, stat) => sum + stat.styleCount, 0).toLocaleString()}
                  </td>
                  <td className="py-3 px-4 text-sm text-emerald-700 text-right">
                    {kpiData.fgInStats.reduce((sum, stat) => sum + stat.vpoCount, 0).toLocaleString()}
                  </td>
                  <td className="py-3 px-4 text-sm text-amber-700 text-right">
                    {kpiData.fgInStats.reduce((sum, stat) => sum + stat.scheduleCount, 0).toLocaleString()}
                  </td>
                </tr>
              </tfoot>
            )}
          </table>
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
        
        <div className="space-y-4 max-h-[500px] overflow-y-auto pr-2 custom-scrollbar [&::-webkit-scrollbar]:w-2 [&::-webkit-scrollbar-track]:bg-slate-100 [&::-webkit-scrollbar-thumb]:bg-slate-300 [&::-webkit-scrollbar-thumb]:rounded-full hover:[&::-webkit-scrollbar-thumb]:bg-slate-400">
          {groupedPendingVPOs.map(([buyer, vpos], groupIdx) => {
            const isExpanded = expandedBuyer === buyer;
            return (
              <div key={groupIdx} className="border border-slate-200 rounded-xl overflow-hidden bg-slate-50/30">
                <button
                  onClick={() => setExpandedBuyer(isExpanded ? null : buyer)}
                  className="w-full flex items-center justify-between p-4 bg-white hover:bg-slate-50 transition-colors focus:outline-none"
                >
                  <div className="flex items-center gap-3">
                    <div className={`p-1.5 rounded-lg ${isExpanded ? 'bg-indigo-100 text-indigo-600' : 'bg-slate-100 text-slate-500'} transition-colors`}>
                      {isExpanded ? <ChevronDown className="w-4 h-4" /> : <ChevronRight className="w-4 h-4" />}
                    </div>
                    <h4 className="text-sm font-bold text-slate-700">
                      {buyer}
                    </h4>
                  </div>
                  <div className="flex items-center gap-4 text-xs font-semibold">
                    <span className="text-slate-500 bg-slate-100 px-2 py-1 rounded-md">{vpos.length} VPOs</span>
                    <span className="text-amber-600 bg-amber-50 px-2 py-1 rounded-md">
                      {vpos.reduce((sum, v) => sum + v.pendingQty, 0).toLocaleString()} Pending
                    </span>
                  </div>
                </button>
                
                {isExpanded && (
                  <div className="p-4 pt-0 border-t border-slate-100 bg-white">
                    <div className="space-y-3 mt-4">
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
                              <p className="text-sm font-bold text-amber-600">{vpo.pendingQty?.toLocaleString() || "0"} Pending</p>
                              <p className="text-xs text-slate-400">{vpo.sewOutQty?.toLocaleString() || "0"} / {vpo.coQty?.toLocaleString() || "0"} Completed</p>
                           </div>
                        </div>
                      ))}
                    </div>
                  </div>
                )}
              </div>
            );
          })}
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
                  <p className="text-2xl font-bold text-slate-800">{selectedVPO.coQty?.toLocaleString() || "0"}</p>
                </div>
                <div className="bg-emerald-50 p-4 rounded-xl border border-emerald-100">
                  <p className="text-xs font-semibold text-emerald-600 uppercase tracking-wider mb-1">Sew Out</p>
                  <p className="text-2xl font-bold text-emerald-700">{selectedVPO.sewOutQty?.toLocaleString() || "0"}</p>
                </div>
                <div className="bg-amber-50 p-4 rounded-xl border border-amber-100">
                  <p className="text-xs font-semibold text-amber-600 uppercase tracking-wider mb-1">Pending</p>
                  <p className="text-2xl font-bold text-amber-700">{selectedVPO.pendingQty?.toLocaleString() || "0"}</p>
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
                          <p className="font-semibold text-slate-700">{sch.coQty?.toLocaleString() || "0"}</p>
                        </div>
                        <div className="text-right">
                          <p className="text-xs text-emerald-500 uppercase">Sew Out</p>
                          <p className="font-semibold text-emerald-600">{sch.sewOutQty?.toLocaleString() || "0"}</p>
                        </div>
                        <div className="text-right w-16">
                          <p className="text-xs text-amber-500 uppercase">Pending</p>
                          <p className="font-bold text-amber-600">{sch.pendingQty?.toLocaleString() || "0"}</p>
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

      {/* FG IN Details Modal */}
      {fgInModal.isOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/40 backdrop-blur-sm">
          <div className="bg-white rounded-2xl shadow-xl w-full max-w-md overflow-hidden animate-in fade-in zoom-in-95 duration-200">
            <div className="p-4 border-b border-slate-100 flex items-center justify-between bg-slate-50/50">
              <div>
                <h3 className="text-lg font-bold text-slate-800">{fgInModal.buyer}</h3>
                <p className="text-sm text-slate-500">FG IN {fgInModal.title} ({fgInModal.items.length})</p>
              </div>
              <button 
                onClick={() => { setFgInModal(prev => ({ ...prev, isOpen: false })); setFgInSearch(''); }}
                className="p-2 hover:bg-slate-200 rounded-full text-slate-500 transition-colors"
              >
                <X className="w-5 h-5" />
              </button>
            </div>
            
            <div className="p-4 border-b border-slate-100">
              <div className="relative">
                <div className="absolute inset-y-0 left-0 pl-3 flex items-center pointer-events-none">
                  <Search className="h-4 w-4 text-slate-400" />
                </div>
                <input
                  type="text"
                  placeholder={`Search ${fgInModal.title.toLowerCase()}...`}
                  value={fgInSearch}
                  onChange={(e) => setFgInSearch(e.target.value)}
                  className="block w-full pl-9 pr-3 py-2 border border-slate-200 rounded-lg text-sm bg-slate-50 focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 transition-colors placeholder:text-slate-400"
                />
              </div>
            </div>
            
            <div 
              ref={fgInListRef}
              className="p-2 max-h-[400px] overflow-y-auto custom-scrollbar"
            >
              {fgInFilteredItems.length > 0 ? (
                <div
                  style={{
                    height: `${fgInVirtualizer.getTotalSize()}px`,
                    width: '100%',
                    position: 'relative',
                  }}
                >
                  {fgInVirtualizer.getVirtualItems().map((virtualItem) => {
                    const item = fgInFilteredItems[virtualItem.index];
                    return (
                      <div
                        key={virtualItem.key}
                        style={{
                          position: 'absolute',
                          top: 0,
                          left: 0,
                          width: '100%',
                          height: `${virtualItem.size}px`,
                          transform: `translateY(${virtualItem.start}px)`,
                        }}
                        className="px-4 flex items-center hover:bg-slate-50 rounded-lg text-sm text-slate-700 font-medium border-b border-slate-50 last:border-transparent transition-colors"
                      >
                        {item}
                      </div>
                    );
                  })}
                </div>
              ) : (
                <div className="px-4 py-8 text-center text-sm text-slate-500">
                  No {fgInModal.title.toLowerCase()} found matching your search.
                </div>
              )}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
