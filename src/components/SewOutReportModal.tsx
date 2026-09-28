import React, { useState, useMemo, useCallback, useEffect, useRef } from 'react';
import { ProductionOrder } from './Dashboard';
import { 
  Printer, 
  FileSpreadsheet, 
  X, 
  Search, 
  CheckSquare, 
  Square, 
  ChevronDown, 
  ChevronRight, 
  CheckCircle2, 
  SlidersHorizontal, 
  Layers, 
  Check, 
  ExternalLink, 
  Download, 
  Eye, 
  AlertCircle,
  Calendar
} from 'lucide-react';
import * as XLSX from 'xlsx';

interface SewOutReportModalProps {
  isOpen: boolean;
  onClose: () => void;
  data: ProductionOrder[];
  remarks: Record<string, { text: string; updatedAt: number | null }>;
  initialFilterMode?: 'sewout_50_100' | 'all';
  initialWeeks?: string[];
  initialBuyers?: string[];
}

export interface VPOOrderGroup {
  vpoNo: string;
  styleNo: string;
  custStyleNo?: string;
  items: ProductionOrder[];
  totalCOQty: number;
  totalSewOutQty: number;
  totalPendingQty: number;
  overallPct: number;
}

export interface BuyerGroup {
  buyer: string;
  vpoGroups: VPOOrderGroup[];
  totalOrders: number;
  totalVPOs: number;
  totalCOQty: number;
  totalSewOutQty: number;
  totalPendingQty: number;
  overallPct: number;
}

export function SewOutReportModal({ 
  isOpen, 
  onClose, 
  data, 
  remarks, 
  initialFilterMode = 'all',
  initialWeeks,
  initialBuyers
}: SewOutReportModalProps) {
  // Preset Mode: 'all' (0-100%) or 'sewout_50_100' (50-100%)
  const [reportMode, setReportMode] = useState<'sewout_50_100' | 'all' | 'custom'>(
    initialFilterMode === 'sewout_50_100' ? 'sewout_50_100' : 'all'
  );

  const [minPct, setMinPct] = useState<number>(initialFilterMode === 'sewout_50_100' ? 50 : 0);
  const [maxPct, setMaxPct] = useState<number>(100);

  // Buyer Checkbox State
  const [selectedBuyers, setSelectedBuyers] = useState<string[]>([]);
  const [buyerSearchQuery, setBuyerSearchQuery] = useState<string>('');
  const [showBuyerPanel, setShowBuyerPanel] = useState<boolean>(false);

  // Week Checkbox State
  const [selectedWeeks, setSelectedWeeks] = useState<string[]>([]);
  const [weekSearchQuery, setWeekSearchQuery] = useState<string>('');
  const [showWeekPanel, setShowWeekPanel] = useState<boolean>(false);

  // Search & Filtering inside modal
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [expandedBuyers, setExpandedBuyers] = useState<Record<string, boolean>>({});
  const [expandedVPOs, setExpandedVPOs] = useState<Record<string, boolean>>({});
  
  // Print & View Options
  const [printOrientation, setPrintOrientation] = useState<'landscape' | 'portrait'>('landscape');
  const [viewTab, setViewTab] = useState<'interactive' | 'print_preview'>('interactive');
  const [printPopupBlockedUrl, setPrintPopupBlockedUrl] = useState<string | null>(null);
  const [statusNotification, setStatusNotification] = useState<{ message: string; type: 'success' | 'warning' | 'info' } | null>(null);

  // Synchronize when modal opens or initial props change
  useEffect(() => {
    if (isOpen) {
      if (initialFilterMode === 'sewout_50_100') {
        setReportMode('sewout_50_100');
        setMinPct(50);
        setMaxPct(100);
      } else {
        setReportMode('all');
        setMinPct(0);
        setMaxPct(100);
      }
      setPrintPopupBlockedUrl(null);
      setStatusNotification(null);
      setSearchQuery('');
    }
  }, [isOpen, initialFilterMode]);

  // Handle Preset Mode Changes
  const handleModeChange = (mode: 'sewout_50_100' | 'all' | 'custom') => {
    setReportMode(mode);
    if (mode === 'sewout_50_100') {
      setMinPct(50);
      setMaxPct(100);
    } else if (mode === 'all') {
      setMinPct(0);
      setMaxPct(100);
    }
  };

  // Compute all unique buyers available from the dataset
  const availableBuyers = useMemo(() => {
    if (!data || data.length === 0) return [];
    const buyerMap: Record<string, number> = {};

    data.forEach(item => {
      const co = Number(item.coQty) || 0;
      if (co <= 0) return;
      const sewOut = Number(item.cumSewOutQty) || 0;
      const pct = (sewOut / co) * 100;
      
      // Match current range
      if (pct >= minPct && pct <= maxPct) {
        const b = item.buyer?.trim() || 'Unknown Buyer';
        buyerMap[b] = (buyerMap[b] || 0) + 1;
      }
    });

    return Object.keys(buyerMap).sort().map(name => ({
      name,
      orderCount: buyerMap[name]
    }));
  }, [data, minPct, maxPct]);

  // Auto-select buyers when modal opens or availableBuyers changes
  useEffect(() => {
    if (availableBuyers.length > 0) {
      if (initialBuyers && initialBuyers.length > 0) {
        const matching = initialBuyers.filter(b => availableBuyers.some(ab => ab.name === b));
        if (matching.length > 0) {
          setSelectedBuyers(matching);
          return;
        }
      }
      setSelectedBuyers(availableBuyers.map(b => b.name));
    } else {
      setSelectedBuyers([]);
    }
  }, [availableBuyers, initialBuyers, isOpen]);

  // Compute all unique weeks available from the dataset
  const availableWeeks = useMemo(() => {
    if (!data || data.length === 0) return [];
    const weekMap: Record<string, number> = {};

    data.forEach(item => {
      const co = Number(item.coQty) || 0;
      if (co <= 0) return;
      const sewOut = Number(item.cumSewOutQty) || 0;
      const pct = (sewOut / co) * 100;
      
      // Match current range
      if (pct >= minPct && pct <= maxPct) {
        const w = item.weekNo?.trim() || 'No Week';
        weekMap[w] = (weekMap[w] || 0) + 1;
      }
    });

    return Object.keys(weekMap).sort((a, b) => {
      return a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' });
    }).map(name => ({
      name,
      orderCount: weekMap[name]
    }));
  }, [data, minPct, maxPct]);

  // Auto-select weeks when modal opens or availableWeeks changes
  useEffect(() => {
    if (availableWeeks.length > 0) {
      if (initialWeeks && initialWeeks.length > 0) {
        const matching = initialWeeks.filter(w => availableWeeks.some(aw => aw.name === w));
        if (matching.length > 0) {
          setSelectedWeeks(matching);
          return;
        }
      }
      setSelectedWeeks(availableWeeks.map(w => w.name));
    } else {
      setSelectedWeeks([]);
    }
  }, [availableWeeks, initialWeeks, isOpen]);

  // Filtered buyers list in the checklist panel based on search
  const filteredAvailableBuyers = useMemo(() => {
    if (!buyerSearchQuery.trim()) return availableBuyers;
    const q = buyerSearchQuery.toLowerCase();
    return availableBuyers.filter(b => b.name.toLowerCase().includes(q));
  }, [availableBuyers, buyerSearchQuery]);

  // Filtered weeks list in the checklist panel based on search
  const filteredAvailableWeeks = useMemo(() => {
    if (!weekSearchQuery.trim()) return availableWeeks;
    const q = weekSearchQuery.toLowerCase();
    return availableWeeks.filter(w => w.name.toLowerCase().includes(q));
  }, [availableWeeks, weekSearchQuery]);

  // Toggle individual buyer checkbox (Tick)
  const toggleBuyer = (buyerName: string) => {
    setSelectedBuyers(prev => {
      if (prev.includes(buyerName)) {
        return prev.filter(b => b !== buyerName);
      } else {
        return [...prev, buyerName];
      }
    });
  };

  const selectAllBuyers = () => {
    setSelectedBuyers(availableBuyers.map(b => b.name));
  };

  const clearAllBuyers = () => {
    setSelectedBuyers([]);
  };

  // Toggle individual week checkbox (Tick)
  const toggleWeek = (weekName: string) => {
    setSelectedWeeks(prev => {
      if (prev.includes(weekName)) {
        return prev.filter(w => w !== weekName);
      } else {
        return [...prev, weekName];
      }
    });
  };

  const selectAllWeeks = () => {
    setSelectedWeeks(availableWeeks.map(w => w.name));
  };

  const clearAllWeeks = () => {
    setSelectedWeeks([]);
  };

  // Filter qualifying orders: coQty > 0, sewOut % is between minPct and maxPct, buyer is ticked, AND week is ticked
  const qualifyingOrders = useMemo(() => {
    if (!data || data.length === 0 || selectedBuyers.length === 0 || selectedWeeks.length === 0) return [];

    const query = searchQuery.trim().toLowerCase();

    return data.filter(item => {
      const co = Number(item.coQty) || 0;
      if (co <= 0) return false;

      const sewOut = Number(item.cumSewOutQty) || 0;
      const pct = (sewOut / co) * 100;

      if (pct < minPct || pct > maxPct) return false;

      const buyerName = item.buyer?.trim() || 'Unknown Buyer';
      if (!selectedBuyers.includes(buyerName)) return false;

      const weekName = item.weekNo?.trim() || 'No Week';
      if (!selectedWeeks.includes(weekName)) return false;

      if (query) {
        const matchVPO = (item.vpoNo || '').toLowerCase().includes(query);
        const matchStyle = (item.styleNo || '').toLowerCase().includes(query);
        const matchColor = (item.colorName || '').toLowerCase().includes(query) || (item.colorCode || '').toLowerCase().includes(query);
        const matchDest = (item.destination || '').toLowerCase().includes(query);
        const matchBuyer = buyerName.toLowerCase().includes(query);
        const matchSchedule = (item.scheduleNo || '').toLowerCase().includes(query);
        const matchWeek = weekName.toLowerCase().includes(query);
        return matchVPO || matchStyle || matchColor || matchDest || matchBuyer || matchSchedule || matchWeek;
      }

      return true;
    });
  }, [data, selectedBuyers, selectedWeeks, minPct, maxPct, searchQuery]);

  // Group qualifying orders first by Buyer, then strictly partitioned by VPO No
  const buyerGroups: BuyerGroup[] = useMemo(() => {
    if (qualifyingOrders.length === 0) return [];

    // Map: Buyer -> Map: VPO No -> Items
    const buyerMap: Record<string, Record<string, ProductionOrder[]>> = {};

    qualifyingOrders.forEach(item => {
      const buyer = item.buyer?.trim() || 'Unknown Buyer';
      const vpo = item.vpoNo?.trim() || 'Unassigned VPO';

      if (!buyerMap[buyer]) {
        buyerMap[buyer] = {};
      }
      if (!buyerMap[buyer][vpo]) {
        buyerMap[buyer][vpo] = [];
      }
      buyerMap[buyer][vpo].push(item);
    });

    const result: BuyerGroup[] = [];

    // Sort Buyers alphabetically
    Object.keys(buyerMap).sort().forEach(buyerName => {
      const vpoDict = buyerMap[buyerName];
      const vpoGroups: VPOOrderGroup[] = [];

      let buyerTotalCO = 0;
      let buyerTotalSewOut = 0;
      let buyerTotalOrders = 0;

      // Sort VPOs numerically/alphabetically
      Object.keys(vpoDict).sort().forEach(vpoNo => {
        const items = vpoDict[vpoNo];
        // Sort items by Plan Del Date, Schedule No
        items.sort((a, b) => {
          const dateComp = (a.planDelDate || '').localeCompare(b.planDelDate || '');
          if (dateComp !== 0) return dateComp;
          return (a.scheduleNo || '').localeCompare(b.scheduleNo || '');
        });

        let vpoCO = 0;
        let vpoSewOut = 0;
        let primaryStyle = '';
        let primaryCustStyle = '';

        items.forEach(item => {
          const co = Number(item.coQty) || 0;
          const sewOut = Number(item.cumSewOutQty) || 0;
          vpoCO += co;
          vpoSewOut += sewOut;
          if (!primaryStyle && item.styleNo) primaryStyle = item.styleNo;
          if (!primaryCustStyle && item.custStyleNo) primaryCustStyle = item.custStyleNo;
        });

        const vpoPending = Math.max(0, vpoCO - vpoSewOut);
        const vpoPct = vpoCO > 0 ? (vpoSewOut / vpoCO) * 100 : 0;

        vpoGroups.push({
          vpoNo,
          styleNo: primaryStyle || '-',
          custStyleNo: primaryCustStyle,
          items,
          totalCOQty: vpoCO,
          totalSewOutQty: vpoSewOut,
          totalPendingQty: vpoPending,
          overallPct: vpoPct
        });

        buyerTotalCO += vpoCO;
        buyerTotalSewOut += vpoSewOut;
        buyerTotalOrders += items.length;
      });

      const buyerTotalPending = Math.max(0, buyerTotalCO - buyerTotalSewOut);
      const buyerOverallPct = buyerTotalCO > 0 ? (buyerTotalSewOut / buyerTotalCO) * 100 : 0;

      result.push({
        buyer: buyerName,
        vpoGroups,
        totalOrders: buyerTotalOrders,
        totalVPOs: vpoGroups.length,
        totalCOQty: buyerTotalCO,
        totalSewOutQty: buyerTotalSewOut,
        totalPendingQty: buyerTotalPending,
        overallPct: buyerOverallPct
      });
    });

    return result;
  }, [qualifyingOrders]);

  // Overall Grand Totals across selected buyers
  const grandTotals = useMemo(() => {
    let totalCOQty = 0;
    let totalSewOutQty = 0;
    let totalPendingQty = 0;
    let totalOrders = qualifyingOrders.length;
    let totalVPOs = 0;

    buyerGroups.forEach(g => {
      totalCOQty += g.totalCOQty;
      totalSewOutQty += g.totalSewOutQty;
      totalPendingQty += g.totalPendingQty;
      totalVPOs += g.totalVPOs;
    });

    const overallPct = totalCOQty > 0 ? (totalSewOutQty / totalCOQty) * 100 : 0;

    return {
      totalBuyers: buyerGroups.length,
      totalVPOs,
      totalOrders,
      totalCOQty,
      totalSewOutQty,
      totalPendingQty,
      overallPct
    };
  }, [buyerGroups, qualifyingOrders]);

  // Expand / Collapse controls
  const toggleBuyerExpand = (buyer: string) => {
    setExpandedBuyers(prev => ({
      ...prev,
      [buyer]: prev[buyer] === undefined ? false : !prev[buyer]
    }));
  };

  const isBuyerExpanded = (buyer: string) => {
    return expandedBuyers[buyer] !== false; // Default expanded
  };

  const toggleVPOExpand = (vpoKey: string) => {
    setExpandedVPOs(prev => ({
      ...prev,
      [vpoKey]: prev[vpoKey] === undefined ? false : !prev[vpoKey]
    }));
  };

  const isVPOExpanded = (vpoKey: string) => {
    return expandedVPOs[vpoKey] !== false; // Default expanded
  };

  const expandAll = () => {
    const updatedBuyers: Record<string, boolean> = {};
    const updatedVPOs: Record<string, boolean> = {};
    buyerGroups.forEach(b => {
      updatedBuyers[b.buyer] = true;
      b.vpoGroups.forEach(v => {
        updatedVPOs[`${b.buyer}_${v.vpoNo}`] = true;
      });
    });
    setExpandedBuyers(updatedBuyers);
    setExpandedVPOs(updatedVPOs);
  };

  const collapseAll = () => {
    const updatedBuyers: Record<string, boolean> = {};
    const updatedVPOs: Record<string, boolean> = {};
    buyerGroups.forEach(b => {
      updatedBuyers[b.buyer] = false;
      b.vpoGroups.forEach(v => {
        updatedVPOs[`${b.buyer}_${v.vpoNo}`] = false;
      });
    });
    setExpandedBuyers(updatedBuyers);
    setExpandedVPOs(updatedVPOs);
  };

  // Generate Standalone, Self-Contained HTML Report with clean A4 layout
  const generatePrintHTML = useCallback((includeAutoPrint: boolean = false) => {
    let buyerSectionsHTML = '';

    buyerGroups.forEach((buyerGroup, bIdx) => {
      let vpoSectionsHTML = '';

      buyerGroup.vpoGroups.forEach((vpoGroup, vIdx) => {
        let itemsHTML = '';

        vpoGroup.items.forEach((item, idx) => {
          const co = Number(item.coQty) || 0;
          const sewOut = Number(item.cumSewOutQty) || 0;
          const pending = Math.max(0, co - sewOut);
          const pct = co > 0 ? ((sewOut / co) * 100).toFixed(1) : '0.0';
          const remarkText = remarks[item.id]?.text || remarks[item.legacyId || '']?.text || '';

          itemsHTML += `
            <tr style="${idx % 2 === 1 ? 'background-color: #f8fafc;' : 'background-color: #ffffff;'}">
              <td style="text-align: center; color: #64748b; font-size: 7.5pt;">${idx + 1}</td>
              <td style="font-weight: 600;">${item.scheduleNo || '-'}</td>
              <td>${item.colorCode || ''}${item.colorName ? ` - ${item.colorName}` : ''}</td>
              <td>${item.destination || '-'}</td>
              <td style="text-align: center;">${item.planDelDate || '-'}</td>
              <td style="text-align: center; font-weight: 500;">${item.weekNo || '-'}</td>
              <td style="text-align: right; font-weight: 500;">${co.toLocaleString()}</td>
              <td style="text-align: right; font-weight: 700; color: #047857;">${sewOut.toLocaleString()}</td>
              <td style="text-align: right; font-weight: 500; color: #b45309;">${pending.toLocaleString()}</td>
              <td style="text-align: right; font-weight: 700; color: #1e40af;">${pct}%</td>
              <td style="text-align: center; font-size: 7.5pt;">${item.statusText || '-'}</td>
              <td style="font-size: 7.5pt; color: #475569;">${remarkText}</td>
            </tr>
          `;
        });

        vpoSectionsHTML += `
          <div class="vpo-block" style="margin-bottom: 10px; border: 1px solid #cbd5e1; border-radius: 4px; background-color: #ffffff; page-break-inside: auto; break-inside: auto;">
            <!-- Distinct VPO Header Bar -->
            <div class="vpo-header" style="background-color: #1e293b; color: #ffffff; padding: 5px 8px; display: flex; justify-content: space-between; align-items: center; font-size: 8.5pt; break-after: avoid; page-break-after: avoid;">
              <div>
                <span style="font-size: 9.5pt; font-weight: 800; color: #38bdf8; letter-spacing: 0.5px;">VPO NO: ${vpoGroup.vpoNo}</span>
                <span style="margin-left: 12px; color: #f1f5f9;">Style: <strong>${vpoGroup.styleNo}</strong>${vpoGroup.custStyleNo ? ` (${vpoGroup.custStyleNo})` : ''}</span>
                <span style="margin-left: 8px; font-size: 7.5pt; color: #94a3b8;">(${vpoGroup.items.length} Schedule${vpoGroup.items.length > 1 ? 's' : ''})</span>
              </div>
              <div style="font-size: 8pt;">
                Order Qty: <strong>${vpoGroup.totalCOQty.toLocaleString()}</strong> &nbsp;|&nbsp; 
                Sew Out: <strong style="color: #34d399;">${vpoGroup.totalSewOutQty.toLocaleString()}</strong> &nbsp;|&nbsp; 
                Pending: <strong style="color: #fde047;">${vpoGroup.totalPendingQty.toLocaleString()}</strong> &nbsp;|&nbsp; 
                SewOut: <strong style="color: #67e8f9;">${vpoGroup.overallPct.toFixed(1)}%</strong>
              </div>
            </div>

            <!-- VPO Schedules Table -->
            <table style="width: 100%; border-collapse: collapse; font-size: 7.8pt;">
              <thead>
                <tr style="background-color: #f1f5f9; color: #334155; font-weight: 700; border-bottom: 1px solid #cbd5e1;">
                  <th style="padding: 4px; width: 28px; text-align: center;">#</th>
                  <th style="padding: 4px; text-align: left;">Schedule No</th>
                  <th style="padding: 4px; text-align: left;">Color</th>
                  <th style="padding: 4px; text-align: left;">Destination</th>
                  <th style="padding: 4px; text-align: center;">Plan Del Date</th>
                  <th style="padding: 4px; text-align: center;">Week</th>
                  <th style="padding: 4px; text-align: right;">Order Qty (CO)</th>
                  <th style="padding: 4px; text-align: right;">Cum Sew Out</th>
                  <th style="padding: 4px; text-align: right;">Pending</th>
                  <th style="padding: 4px; text-align: right;">SewOut %</th>
                  <th style="padding: 4px; text-align: center;">Status</th>
                  <th style="padding: 4px; text-align: left;">Remark</th>
                </tr>
              </thead>
              <tbody>
                ${itemsHTML}
              </tbody>
              <tfoot>
                <tr style="background-color: #f8fafc; font-weight: bold; border-top: 1.5px solid #cbd5e1; font-size: 8pt;">
                  <td colspan="6" style="padding: 5px 8px; text-align: right; color: #475569;">
                    Subtotal for VPO ${vpoGroup.vpoNo}:
                  </td>
                  <td style="padding: 5px 4px; text-align: right;">${vpoGroup.totalCOQty.toLocaleString()}</td>
                  <td style="padding: 5px 4px; text-align: right; color: #047857;">${vpoGroup.totalSewOutQty.toLocaleString()}</td>
                  <td style="padding: 5px 4px; text-align: right; color: #b45309;">${vpoGroup.totalPendingQty.toLocaleString()}</td>
                  <td style="padding: 5px 4px; text-align: right; color: #1e40af;">${vpoGroup.overallPct.toFixed(1)}%</td>
                  <td colspan="2"></td>
                </tr>
              </tfoot>
            </table>
          </div>
        `;
      });

      buyerSectionsHTML += `
        <div class="buyer-section" style="margin-bottom: 16px; page-break-inside: auto; break-inside: auto;">
          <!-- Distinct Buyer Header Banner -->
          <div class="buyer-header" style="background-color: #0f172a; color: #ffffff; padding: 6px 10px; border-radius: 4px; display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px; border-left: 5px solid #4f46e5; break-after: avoid; page-break-after: avoid;">
            <div>
              <span style="font-size: 10.5pt; font-weight: 800; letter-spacing: 0.5px;">BUYER: ${buyerGroup.buyer.toUpperCase()}</span>
              <span style="margin-left: 10px; font-size: 8pt; color: #94a3b8; font-weight: normal;">(${buyerGroup.totalVPOs} VPO${buyerGroup.totalVPOs > 1 ? 's' : ''}, ${buyerGroup.totalOrders} Lines)</span>
            </div>
            <div style="font-size: 8pt;">
              Total Order Qty: <strong style="color: #ffffff;">${buyerGroup.totalCOQty.toLocaleString()}</strong> &nbsp;|&nbsp; 
              Sew Out: <strong style="color: #34d399;">${buyerGroup.totalSewOutQty.toLocaleString()}</strong> &nbsp;|&nbsp; 
              Pending: <strong style="color: #fde047;">${buyerGroup.totalPendingQty.toLocaleString()}</strong> &nbsp;|&nbsp; 
              Progress: <strong style="color: #93c5fd;">${buyerGroup.overallPct.toFixed(1)}%</strong>
            </div>
          </div>

          <!-- All VPOs for this Buyer -->
          ${vpoSectionsHTML}

          <!-- Buyer Summary Footer -->
          <div style="background-color: #e2e8f0; border: 1.5px solid #94a3b8; padding: 6px 10px; border-radius: 4px; font-weight: bold; font-size: 8pt; color: #0f172a; display: flex; justify-content: space-between; margin-top: 4px; margin-bottom: 12px; break-inside: avoid; page-break-inside: avoid;">
            <span>TOTAL FOR BUYER: ${buyerGroup.buyer.toUpperCase()} (${buyerGroup.totalVPOs} VPOs)</span>
            <span>
              Order Qty: ${buyerGroup.totalCOQty.toLocaleString()} &nbsp;&bull;&nbsp; 
              Sew Out: ${buyerGroup.totalSewOutQty.toLocaleString()} &nbsp;&bull;&nbsp; 
              Pending: ${buyerGroup.totalPendingQty.toLocaleString()} &nbsp;&bull;&nbsp; 
              Overall Rate: ${buyerGroup.overallPct.toFixed(1)}%
            </span>
          </div>
        </div>
      `;
    });

    const currentDateStr = new Date().toLocaleString('en-US', {
      dateStyle: 'medium',
      timeStyle: 'short'
    });

    const reportTitleText = reportMode === 'sewout_50_100'
      ? 'SEWING OUT 50% – 100% REPORT (BUYER & VPO WISE)'
      : reportMode === 'all'
      ? 'PRODUCTION ORDERS REPORT (BUYER & VPO WISE)'
      : `SEWING OUT ${minPct}% – ${maxPct}% REPORT (BUYER & VPO WISE)`;

    return `<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <title>${reportTitleText}</title>
  <style>
    @page { 
      size: A4 ${printOrientation}; 
      margin: 8mm; 
    }
    * { box-sizing: border-box; }
    html, body {
      margin: 0;
      padding: 0;
      background: #ffffff;
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif; 
      font-size: 8pt; 
      color: #0f172a; 
      -webkit-print-color-adjust: exact !important;
      print-color-adjust: exact !important;
    }
    body { 
      padding: 10px;
    }
    .print-controls-banner {
      background: #f1f5f9;
      border: 1px solid #cbd5e1;
      padding: 8px 12px;
      border-radius: 6px;
      margin-bottom: 10px;
      display: flex;
      justify-content: space-between;
      align-items: center;
      font-size: 9pt;
    }
    .print-btn {
      background: #4f46e5;
      color: #fff;
      border: none;
      padding: 6px 14px;
      border-radius: 4px;
      font-weight: bold;
      cursor: pointer;
    }
    .report-header {
      border-bottom: 2px solid #0f172a;
      padding-bottom: 4px;
      margin-bottom: 8px;
      display: flex;
      justify-content: space-between;
      align-items: flex-end;
    }
    .report-title {
      font-size: 12pt;
      font-weight: 800;
      color: #0f172a;
      margin: 0 0 2px 0;
    }
    .report-subtitle {
      font-size: 8pt;
      color: #475569;
      margin: 0;
    }
    .report-meta {
      text-align: right;
      font-size: 7.5pt;
      color: #475569;
      line-height: 1.3;
    }
    .summary-card {
      background-color: #f8fafc;
      border: 1px solid #cbd5e1;
      border-radius: 6px;
      padding: 6px 10px;
      margin-bottom: 10px;
      display: flex;
      justify-content: space-between;
      align-items: center;
      break-inside: avoid;
      page-break-inside: avoid;
    }
    .summary-stat {
      text-align: center;
      padding: 0 6px;
    }
    .summary-stat-val {
      font-size: 11pt;
      font-weight: 800;
      color: #0f172a;
    }
    .summary-stat-label {
      font-size: 6.8pt;
      font-weight: 600;
      color: #64748b;
      text-transform: uppercase;
      letter-spacing: 0.5px;
    }
    .grand-total-box {
      background-color: #0f172a;
      color: #ffffff;
      padding: 8px 14px;
      border-radius: 4px;
      margin-top: 12px;
      display: flex;
      justify-content: space-between;
      align-items: center;
      font-size: 8.5pt;
      font-weight: bold;
      break-inside: avoid;
      page-break-inside: avoid;
    }
    @media print {
      html, body { 
        padding: 0 !important; 
        margin: 0 !important; 
        background: #ffffff !important;
      }
      .print-controls-banner { 
        display: none !important; 
        height: 0 !important;
        margin: 0 !important;
        padding: 0 !important;
      }
      .buyer-section { 
        break-inside: auto !important; 
        page-break-inside: auto !important; 
        margin-bottom: 12px !important;
      }
      .buyer-header {
        break-after: avoid !important;
        page-break-after: avoid !important;
      }
      .vpo-block { 
        break-inside: auto !important; 
        page-break-inside: auto !important; 
        margin-bottom: 8px !important;
      }
      .vpo-header {
        break-after: avoid !important;
        page-break-after: avoid !important;
      }
      table {
        page-break-inside: auto !important;
        break-inside: auto !important;
      }
      tr {
        break-inside: avoid !important;
        page-break-inside: avoid !important;
      }
      thead {
        display: table-header-group !important;
      }
      tfoot {
        display: table-row-group !important;
        break-inside: avoid !important;
        page-break-inside: avoid !important;
      }
      .grand-total-box {
        break-inside: avoid !important;
        page-break-inside: avoid !important;
      }
    }
  </style>
  ${includeAutoPrint ? `
  <script>
    window.addEventListener('load', function() {
      setTimeout(function() {
        window.focus();
        window.print();
      }, 500);
    });
  </script>
  ` : ''}
</head>
<body>
  <!-- Print Bar (visible only in browser tab, hidden during actual print) -->
  <div class="print-controls-banner">
    <div>
      <strong>Print Ready Document:</strong> Use the button on the right or press <kbd style="background: #e2e8f0; padding: 2px 5px; border-radius: 3px;">Ctrl + P</kbd>
    </div>
    <button class="print-btn" onclick="window.print()">Print Now / Save PDF</button>
  </div>

  <div class="report-header">
    <div>
      <h1 class="report-title">${reportTitleText}</h1>
      <p class="report-subtitle">Separated by VPO No under each Buyer &bull; Progress Range: ${minPct}% – ${maxPct}%</p>
    </div>
    <div class="report-meta">
      <div>Printed: <strong>${currentDateStr}</strong></div>
      <div>Orientation: <strong>${printOrientation.toUpperCase()}</strong></div>
      <div>Buyers: <strong>${selectedBuyers.length === availableBuyers.length ? 'All Buyers' : `${selectedBuyers.length} selected`}</strong></div>
      <div>Weeks: <strong>${selectedWeeks.length === availableWeeks.length ? 'All Weeks' : `${selectedWeeks.length} selected`}</strong></div>
    </div>
  </div>

  <div class="summary-card">
    <div class="summary-stat">
      <div class="summary-stat-val">${grandTotals.totalBuyers}</div>
      <div class="summary-stat-label">Buyers</div>
    </div>
    <div class="summary-stat">
      <div class="summary-stat-val">${grandTotals.totalVPOs}</div>
      <div class="summary-stat-label">Total VPOs</div>
    </div>
    <div class="summary-stat">
      <div class="summary-stat-val">${grandTotals.totalOrders}</div>
      <div class="summary-stat-label">Orders</div>
    </div>
    <div class="summary-stat">
      <div class="summary-stat-val">${grandTotals.totalCOQty.toLocaleString()}</div>
      <div class="summary-stat-label">Total Order Qty</div>
    </div>
    <div class="summary-stat">
      <div class="summary-stat-val" style="color: #047857;">${grandTotals.totalSewOutQty.toLocaleString()}</div>
      <div class="summary-stat-label">Total Sew Out</div>
    </div>
    <div class="summary-stat">
      <div class="summary-stat-val" style="color: #b45309;">${grandTotals.totalPendingQty.toLocaleString()}</div>
      <div class="summary-stat-label">Total Pending</div>
    </div>
    <div class="summary-stat">
      <div class="summary-stat-val" style="color: #1e40af;">${grandTotals.overallPct.toFixed(1)}%</div>
      <div class="summary-stat-label">Overall Progress</div>
    </div>
  </div>

  ${buyerSectionsHTML}

  <div class="grand-total-box">
    <span>GRAND TOTAL (${grandTotals.totalBuyers} Buyers &bull; ${grandTotals.totalVPOs} VPOs &bull; ${grandTotals.totalOrders} Lines)</span>
    <span>
      CO Qty: ${grandTotals.totalCOQty.toLocaleString()} &nbsp;|&nbsp; 
      Sew Out: ${grandTotals.totalSewOutQty.toLocaleString()} &nbsp;|&nbsp; 
      Pending: ${grandTotals.totalPendingQty.toLocaleString()} &nbsp;|&nbsp; 
      Progress: ${grandTotals.overallPct.toFixed(1)}%
    </span>
  </div>
</body>
</html>`;
  }, [buyerGroups, grandTotals, minPct, maxPct, printOrientation, remarks, reportMode]);

  // Robust Print Handler:
  // 1. Creates a Blob URL with the self-contained print page + auto-print script
  // 2. Opens it in a new window/tab (bypassing any iframe sandbox restrictions)
  // 3. If popup is blocked by the browser, provides an instant one-click link + falls back to window.print
  const handleOpenPrintWindow = useCallback(() => {
    if (buyerGroups.length === 0) {
      setStatusNotification({
        message: "No data available to print. Please tick at least one buyer.",
        type: 'warning'
      });
      return;
    }

    try {
      const htmlContent = generatePrintHTML(true);
      const blob = new Blob([htmlContent], { type: 'text/html;charset=utf-8' });
      const url = URL.createObjectURL(blob);

      const win = window.open(url, '_blank');

      if (!win || win.closed || typeof win.closed === 'undefined') {
        // Browser popup blocker triggered. Store URL to show clickable button in UI
        setPrintPopupBlockedUrl(url);
        setStatusNotification({
          message: "Popup blocked by browser. Please click 'Open Print Tab' below or use Download HTML.",
          type: 'warning'
        });
      } else {
        setPrintPopupBlockedUrl(null);
        setStatusNotification({
          message: "Print window opened! If your browser requires, allow print in the new tab.",
          type: 'success'
        });
      }
    } catch (err: any) {
      console.warn("Popup print exception, falling back to window.print", err);
      window.print();
    }
  }, [buyerGroups, generatePrintHTML]);

  // Direct Browser Print (for fullscreen/standalone view)
  const handleDirectPrint = useCallback(() => {
    if (buyerGroups.length === 0) {
      setStatusNotification({
        message: selectedBuyers.length === 0 
          ? "No data to print. Please tick at least one buyer." 
          : "No data to print. Please tick at least one week.",
        type: 'warning'
      });
      return;
    }
    // Switch to Print Preview and trigger print
    setViewTab('print_preview');
    setTimeout(() => {
      window.print();
    }, 300);
  }, [buyerGroups, selectedBuyers]);

  // Download Standalone Printable HTML File
  const handleDownloadHTML = useCallback(() => {
    if (buyerGroups.length === 0) {
      setStatusNotification({
        message: selectedBuyers.length === 0 
          ? "No data to download. Please tick at least one buyer." 
          : "No data to download. Please tick at least one week.",
        type: 'warning'
      });
      return;
    }

    const htmlContent = generatePrintHTML(true);
    const blob = new Blob([htmlContent], { type: 'text/html;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `Production_Report_Buyer_VPO_${new Date().toISOString().slice(0, 10)}.html`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);

    setStatusNotification({
      message: "Printable HTML report downloaded. Open it in Chrome/Edge to print immediately.",
      type: 'success'
    });
  }, [buyerGroups, generatePrintHTML, selectedBuyers]);

  // Export to Excel: Structured by Buyer and Separated by VPO No
  const handleExportExcel = useCallback(() => {
    if (buyerGroups.length === 0) return;

    const rows: any[] = [];

    rows.push([reportMode === 'sewout_50_100' ? 'SEWING OUT 50% - 100% REPORT (BUYER & VPO WISE)' : 'PRODUCTION REPORT (BUYER & VPO WISE)']);
    const weekExportText = selectedWeeks.length === availableWeeks.length ? 'All Weeks' : `${selectedWeeks.length} Weeks (${selectedWeeks.join(', ')})`;
    rows.push([`Generated: ${new Date().toLocaleString()}`, `Range: ${minPct}% - ${maxPct}%`, `Weeks: ${weekExportText}`]);
    rows.push([`Total Buyers: ${grandTotals.totalBuyers}`, `Total VPOs: ${grandTotals.totalVPOs}`, `Total Orders: ${grandTotals.totalOrders}`]);
    rows.push([]);

    buyerGroups.forEach(buyerGroup => {
      // Buyer Title Banner
      rows.push([`BUYER: ${buyerGroup.buyer.toUpperCase()}`, `(${buyerGroup.totalVPOs} VPOs, ${buyerGroup.totalOrders} Orders)`, '', '', '', '', '', '', '', '', '', '', '', '', '']);
      rows.push([]);

      buyerGroup.vpoGroups.forEach(vpoGroup => {
        // VPO Header
        rows.push([
          `>> VPO: ${vpoGroup.vpoNo}`,
          `Style: ${vpoGroup.styleNo}`,
          vpoGroup.custStyleNo ? `(${vpoGroup.custStyleNo})` : '',
          `Order Qty: ${vpoGroup.totalCOQty}`,
          `Sew Out: ${vpoGroup.totalSewOutQty}`,
          `Pending: ${vpoGroup.totalPendingQty}`,
          `Rate: ${vpoGroup.overallPct.toFixed(1)}%`
        ]);

        // Table Column Headers
        rows.push([
          '#',
          'VPO No',
          'Style No',
          'Schedule No',
          'Color Code',
          'Color Name',
          'Destination',
          'Plan Del Date',
          'WEEK NO',
          'Order Qty (CO)',
          'Cum Sew In Qty',
          'Cum SewOut Qty',
          'Pending Qty',
          'SewOut %',
          'Status',
          'Remark'
        ]);

        vpoGroup.items.forEach((item, idx) => {
          const co = Number(item.coQty) || 0;
          const sewOut = Number(item.cumSewOutQty) || 0;
          const pending = Math.max(0, co - sewOut);
          const pct = co > 0 ? Number(((sewOut / co) * 100).toFixed(1)) : 0;
          const remarkText = remarks[item.id]?.text || remarks[item.legacyId || '']?.text || '';

          rows.push([
            idx + 1,
            item.vpoNo,
            item.styleNo,
            item.scheduleNo,
            item.colorCode,
            item.colorName,
            item.destination,
            item.planDelDate,
            item.weekNo,
            co,
            Number(item.cumSewInQty) || 0,
            sewOut,
            pending,
            `${pct}%`,
            item.statusText,
            remarkText
          ]);
        });

        // VPO Subtotal
        rows.push([
          '',
          `Subtotal for VPO ${vpoGroup.vpoNo}`,
          '',
          '',
          '',
          '',
          '',
          '',
          '',
          vpoGroup.totalCOQty,
          '',
          vpoGroup.totalSewOutQty,
          vpoGroup.totalPendingQty,
          `${vpoGroup.overallPct.toFixed(1)}%`,
          '',
          ''
        ]);
        rows.push([]); // Gap between VPOs
      });

      // Buyer Subtotal
      rows.push([
        '',
        `*** TOTAL FOR BUYER: ${buyerGroup.buyer.toUpperCase()} ***`,
        '',
        '',
        '',
        '',
        '',
        '',
        '',
        buyerGroup.totalCOQty,
        '',
        buyerGroup.totalSewOutQty,
        buyerGroup.totalPendingQty,
        `${buyerGroup.overallPct.toFixed(1)}%`,
        '',
        ''
      ]);
      rows.push([]);
      rows.push([]); // Double gap between Buyers
    });

    // Grand Total Row
    rows.push([
      '',
      '=== GRAND TOTAL (ALL SELECTED BUYERS) ===',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      grandTotals.totalCOQty,
      '',
      grandTotals.totalSewOutQty,
      grandTotals.totalPendingQty,
      `${grandTotals.overallPct.toFixed(1)}%`,
      '',
      ''
    ]);

    const worksheet = XLSX.utils.aoa_to_sheet(rows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Buyer_VPO_Report');

    const fileName = `Production_Buyer_VPO_${new Date().toISOString().slice(0, 10)}.xlsx`;
    XLSX.writeFile(workbook, fileName);
  }, [buyerGroups, grandTotals, minPct, maxPct, remarks, reportMode]);

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/70 backdrop-blur-sm p-2 sm:p-4 overflow-y-auto">
      <div className="bg-white w-full max-w-7xl max-h-[96vh] rounded-2xl shadow-2xl border border-slate-200 flex flex-col overflow-hidden animate-in fade-in zoom-in-95 duration-200">
        
        {/* Modal Header */}
        <div className="px-6 py-4 bg-slate-900 text-white flex items-center justify-between border-b border-slate-800">
          <div>
            <div className="flex items-center gap-2.5">
              <span className="p-1.5 bg-indigo-500/20 text-indigo-400 rounded-lg border border-indigo-500/30">
                <Printer className="w-5 h-5" />
              </span>
              <h2 className="text-lg font-bold tracking-tight text-white">
                Production Report & Print Center (Buyer & VPO Wise)
              </h2>
            </div>
            <p className="text-xs text-slate-400 mt-1">
              Select buyers with checkboxes. Orders are separated by <strong>VPO No</strong> under each Buyer.
            </p>
          </div>
          <button
            onClick={onClose}
            className="p-2 text-slate-400 hover:text-white hover:bg-slate-800 rounded-lg transition-colors"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Top Control Bar: Mode Toggle, Buyer Filter Button, View Switcher & Print Actions */}
        <div className="p-3.5 bg-slate-50 border-b border-slate-200 flex flex-wrap items-center justify-between gap-3 text-sm">
          
          <div className="flex flex-wrap items-center gap-2.5">
            {/* Mode Selector */}
            <div className="flex bg-slate-200/80 p-0.5 rounded-lg text-xs font-semibold">
              <button
                type="button"
                onClick={() => handleModeChange('sewout_50_100')}
                className={`px-3 py-1.5 rounded-md transition-all ${
                  reportMode === 'sewout_50_100'
                    ? 'bg-white text-indigo-900 shadow-sm'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                50% – 100% SewOut
              </button>
              <button
                type="button"
                onClick={() => handleModeChange('all')}
                className={`px-3 py-1.5 rounded-md transition-all ${
                  reportMode === 'all'
                    ? 'bg-white text-indigo-900 shadow-sm'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                All Orders (0-100%)
              </button>
              <button
                type="button"
                onClick={() => handleModeChange('custom')}
                className={`px-3 py-1.5 rounded-md transition-all ${
                  reportMode === 'custom'
                    ? 'bg-white text-indigo-900 shadow-sm'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                Custom Range
              </button>
            </div>

            {/* Custom Range Inputs if custom mode is selected */}
            {reportMode === 'custom' && (
              <div className="flex items-center gap-1.5 bg-white px-2 py-1 rounded-lg border border-slate-300 shadow-sm text-xs">
                <span className="font-semibold text-slate-500">Range:</span>
                <input
                  type="number"
                  min="0"
                  max="100"
                  value={minPct}
                  onChange={(e) => setMinPct(Math.max(0, Math.min(100, Number(e.target.value) || 0)))}
                  className="w-12 px-1 py-0.5 border border-slate-200 rounded text-center font-bold text-slate-800"
                />
                <span className="text-slate-400">% to</span>
                <input
                  type="number"
                  min="0"
                  max="100"
                  value={maxPct}
                  onChange={(e) => setMaxPct(Math.max(0, Math.min(100, Number(e.target.value) || 0)))}
                  className="w-12 px-1 py-0.5 border border-slate-200 rounded text-center font-bold text-slate-800"
                />
                <span className="text-slate-400">%</span>
              </div>
            )}

            {/* Buyer Checklist Toggle Button */}
            <button
              type="button"
              onClick={() => {
                setShowBuyerPanel(!showBuyerPanel);
                if (!showBuyerPanel) setShowWeekPanel(false);
              }}
              className={`py-1.5 px-3 rounded-lg text-xs font-bold border flex items-center gap-2 transition-all ${
                selectedBuyers.length > 0 
                  ? 'bg-indigo-50 border-indigo-300 text-indigo-900 hover:bg-indigo-100/70 shadow-sm' 
                  : 'bg-rose-50 border-rose-300 text-rose-700'
              }`}
            >
              <CheckSquare className="w-4 h-4 text-indigo-600" />
              <span>
                {selectedBuyers.length === availableBuyers.length && availableBuyers.length > 0
                  ? `All Buyers Ticked (${availableBuyers.length})`
                  : selectedBuyers.length === 0
                  ? 'No Buyer Ticked!'
                  : `Buyers: ${selectedBuyers.length} of ${availableBuyers.length}`}
              </span>
              <ChevronDown className={`w-3.5 h-3.5 text-slate-500 transition-transform ${showBuyerPanel ? 'rotate-180' : ''}`} />
            </button>

            {/* Week Checklist Toggle Button */}
            <button
              type="button"
              onClick={() => {
                setShowWeekPanel(!showWeekPanel);
                if (!showWeekPanel) setShowBuyerPanel(false);
              }}
              className={`py-1.5 px-3 rounded-lg text-xs font-bold border flex items-center gap-2 transition-all ${
                selectedWeeks.length > 0 
                  ? 'bg-amber-50 border-amber-300 text-amber-950 hover:bg-amber-100/70 shadow-sm' 
                  : 'bg-rose-50 border-rose-300 text-rose-700'
              }`}
            >
              <Calendar className="w-4 h-4 text-amber-600" />
              <span>
                {selectedWeeks.length === availableWeeks.length && availableWeeks.length > 0
                  ? `All Weeks Ticked (${availableWeeks.length})`
                  : selectedWeeks.length === 0
                  ? 'No Week Ticked!'
                  : `Weeks: ${selectedWeeks.length} of ${availableWeeks.length}`}
              </span>
              <ChevronDown className={`w-3.5 h-3.5 text-slate-500 transition-transform ${showWeekPanel ? 'rotate-180' : ''}`} />
            </button>

            {/* Quick Search */}
            <div className="relative">
              <Search className="w-3.5 h-3.5 absolute left-2.5 top-1/2 -translate-y-1/2 text-slate-400" />
              <input
                type="text"
                placeholder="Search VPO, style, color..."
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                className="pl-8 pr-3 py-1.5 bg-white border border-slate-300 rounded-lg text-xs placeholder:text-slate-400 text-slate-700 w-40 sm:w-48 focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500"
              />
              {searchQuery && (
                <button 
                  onClick={() => setSearchQuery('')}
                  className="absolute right-2 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600"
                >
                  <X className="w-3 h-3" />
                </button>
              )}
            </div>
          </div>

          {/* Action Buttons: Orientation, Print Popup, Download HTML, Excel */}
          <div className="flex items-center gap-2">
            {/* View Tab Toggle */}
            <div className="flex bg-slate-200/80 p-0.5 rounded-lg text-xs font-semibold">
              <button
                type="button"
                onClick={() => setViewTab('interactive')}
                className={`px-2.5 py-1 rounded transition-all flex items-center gap-1 ${
                  viewTab === 'interactive' ? 'bg-white text-indigo-900 shadow-sm' : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                <Layers className="w-3.5 h-3.5" />
                View Table
              </button>
              <button
                type="button"
                onClick={() => setViewTab('print_preview')}
                className={`px-2.5 py-1 rounded transition-all flex items-center gap-1 ${
                  viewTab === 'print_preview' ? 'bg-white text-indigo-900 shadow-sm' : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                <Eye className="w-3.5 h-3.5" />
                Print Preview
              </button>
            </div>

            {/* Orientation */}
            <div className="flex bg-white rounded-lg border border-slate-300 p-0.5 text-xs font-medium">
              <button
                onClick={() => setPrintOrientation('landscape')}
                className={`px-2 py-1 rounded transition-all ${
                  printOrientation === 'landscape' ? 'bg-indigo-600 text-white shadow-sm' : 'text-slate-600 hover:text-slate-900'
                }`}
                title="Landscape (Recommended for wide table)"
              >
                Landscape
              </button>
              <button
                onClick={() => setPrintOrientation('portrait')}
                className={`px-2 py-1 rounded transition-all ${
                  printOrientation === 'portrait' ? 'bg-indigo-600 text-white shadow-sm' : 'text-slate-600 hover:text-slate-900'
                }`}
                title="Portrait"
              >
                Portrait
              </button>
            </div>

            {/* Export Excel */}
            <button
              onClick={handleExportExcel}
              disabled={buyerGroups.length === 0}
              className="px-3 py-1.5 bg-emerald-50 hover:bg-emerald-100 text-emerald-700 border border-emerald-200 rounded-lg text-xs font-semibold flex items-center gap-1 transition-colors disabled:opacity-50"
              title="Download structured Excel report"
            >
              <FileSpreadsheet className="w-3.5 h-3.5" />
              Excel
            </button>

            {/* Download HTML Print File */}
            <button
              onClick={handleDownloadHTML}
              disabled={buyerGroups.length === 0}
              className="px-3 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 border border-slate-300 rounded-lg text-xs font-semibold flex items-center gap-1 transition-colors disabled:opacity-50"
              title="Download standalone HTML to print anytime"
            >
              <Download className="w-3.5 h-3.5" />
              Save HTML
            </button>

            {/* PRINT POPUP BUTTON */}
            <button
              id="print-sewout-report-btn"
              onClick={handleOpenPrintWindow}
              disabled={buyerGroups.length === 0}
              className="px-4 py-1.5 bg-indigo-600 hover:bg-indigo-700 text-white rounded-lg text-xs font-bold flex items-center gap-1.5 shadow transition-all hover:shadow-md active:scale-95 disabled:opacity-50"
            >
              <Printer className="w-4 h-4" />
              Print Report Now
            </button>
          </div>
        </div>

        {/* Status Notification Banner (if popup blocked or print ready) */}
        {statusNotification && (
          <div className={`px-6 py-2 flex items-center justify-between text-xs font-medium border-b ${
            statusNotification.type === 'warning' 
              ? 'bg-amber-50 text-amber-900 border-amber-200' 
              : statusNotification.type === 'success'
              ? 'bg-emerald-50 text-emerald-900 border-emerald-200'
              : 'bg-blue-50 text-blue-900 border-blue-200'
          }`}>
            <div className="flex items-center gap-2">
              <AlertCircle className="w-4 h-4 flex-shrink-0" />
              <span>{statusNotification.message}</span>
            </div>
            {printPopupBlockedUrl && (
              <a
                href={printPopupBlockedUrl}
                target="_blank"
                rel="noopener noreferrer"
                className="px-3 py-1 bg-amber-600 hover:bg-amber-700 text-white rounded font-bold flex items-center gap-1 shadow-sm text-xs ml-4"
              >
                <ExternalLink className="w-3 h-3" />
                Open Print Tab
              </a>
            )}
          </div>
        )}

        {/* BUYER TICK CHECKLIST PANEL (Explicitly visible so user can tick buyers) */}
        {showBuyerPanel && (
          <div className="px-6 py-3 bg-indigo-50/60 border-b border-indigo-100 flex flex-col gap-2">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-2">
                <span className="text-xs font-bold uppercase tracking-wider text-indigo-950 flex items-center gap-1.5">
                  <CheckSquare className="w-3.5 h-3.5 text-indigo-600" />
                  Tick Buyers to Include:
                </span>
                <span className="text-xs text-indigo-700 font-medium">
                  ({selectedBuyers.length} selected of {availableBuyers.length})
                </span>
              </div>
              <div className="flex items-center gap-3 text-xs">
                <button
                  type="button"
                  onClick={selectAllBuyers}
                  className="font-bold text-indigo-700 hover:text-indigo-900 hover:underline"
                >
                  Select All
                </button>
                <span className="text-indigo-300">&bull;</span>
                <button
                  type="button"
                  onClick={clearAllBuyers}
                  className="font-bold text-rose-600 hover:text-rose-800 hover:underline"
                >
                  Clear All
                </button>
                <span className="text-indigo-300">&bull;</span>
                <div className="relative">
                  <input
                    type="text"
                    placeholder="Filter buyer names..."
                    value={buyerSearchQuery}
                    onChange={(e) => setBuyerSearchQuery(e.target.value)}
                    className="pl-6 pr-2 py-0.5 text-[11px] bg-white border border-indigo-200 rounded focus:outline-none w-36"
                  />
                  <Search className="w-3 h-3 absolute left-1.5 top-1/2 -translate-y-1/2 text-slate-400" />
                </div>
              </div>
            </div>

            {/* Checkbox Chips Grid */}
            <div className="flex flex-wrap gap-1.5 max-h-32 overflow-y-auto pr-1">
              {filteredAvailableBuyers.length === 0 ? (
                <span className="text-xs text-slate-400 italic py-1">No buyers found matching criteria.</span>
              ) : (
                filteredAvailableBuyers.map(b => {
                  const isTicked = selectedBuyers.includes(b.name);
                  return (
                    <button
                      key={b.name}
                      type="button"
                      onClick={() => toggleBuyer(b.name)}
                      className={`px-2.5 py-1 rounded-lg text-xs font-medium flex items-center gap-1.5 transition-all border ${
                        isTicked
                          ? 'bg-indigo-600 text-white border-indigo-600 shadow-sm'
                          : 'bg-white text-slate-700 border-slate-300 hover:bg-slate-100'
                      }`}
                    >
                      <input
                        type="checkbox"
                        checked={isTicked}
                        onChange={() => {}} // handled by button
                        className="w-3.5 h-3.5 pointer-events-none rounded border-slate-300 text-indigo-600 focus:ring-0 cursor-pointer"
                      />
                      <span className="font-semibold">{b.name}</span>
                      <span className={`text-[10px] px-1 rounded font-mono ${
                        isTicked ? 'bg-indigo-700/60 text-indigo-100' : 'bg-slate-200 text-slate-600'
                      }`}>
                        {b.orderCount}
                      </span>
                    </button>
                  );
                })
              )}
            </div>
          </div>
        )}

        {/* WEEK TICK CHECKLIST PANEL (Explicitly visible so user can select/tick weeks) */}
        {showWeekPanel && (
          <div className="px-6 py-3 bg-amber-50/70 border-b border-amber-200 flex flex-col gap-2">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-2">
                <span className="text-xs font-bold uppercase tracking-wider text-amber-950 flex items-center gap-1.5">
                  <Calendar className="w-3.5 h-3.5 text-amber-600" />
                  Tick Weeks to Include:
                </span>
                <span className="text-xs text-amber-800 font-medium">
                  ({selectedWeeks.length} selected of {availableWeeks.length})
                </span>
              </div>
              <div className="flex items-center gap-3 text-xs">
                <button
                  type="button"
                  onClick={selectAllWeeks}
                  className="font-bold text-amber-800 hover:text-amber-950 hover:underline"
                >
                  Select All
                </button>
                <span className="text-amber-300">&bull;</span>
                <button
                  type="button"
                  onClick={clearAllWeeks}
                  className="font-bold text-rose-600 hover:text-rose-800 hover:underline"
                >
                  Clear All
                </button>
                <span className="text-amber-300">&bull;</span>
                <div className="relative">
                  <input
                    type="text"
                    placeholder="Filter week..."
                    value={weekSearchQuery}
                    onChange={(e) => setWeekSearchQuery(e.target.value)}
                    className="pl-6 pr-2 py-0.5 text-[11px] bg-white border border-amber-300 rounded focus:outline-none w-28"
                  />
                  <Search className="w-3 h-3 absolute left-1.5 top-1/2 -translate-y-1/2 text-slate-400" />
                </div>
              </div>
            </div>

            {/* Checkbox Chips Grid for Weeks */}
            <div className="flex flex-wrap gap-1.5 max-h-32 overflow-y-auto pr-1">
              {filteredAvailableWeeks.length === 0 ? (
                <span className="text-xs text-slate-400 italic py-1">No weeks found matching criteria.</span>
              ) : (
                filteredAvailableWeeks.map(w => {
                  const isTicked = selectedWeeks.includes(w.name);
                  return (
                    <button
                      key={w.name}
                      type="button"
                      onClick={() => toggleWeek(w.name)}
                      className={`px-2.5 py-1 rounded-lg text-xs font-medium flex items-center gap-1.5 transition-all border ${
                        isTicked
                          ? 'bg-amber-600 text-white border-amber-600 shadow-sm'
                          : 'bg-white text-slate-700 border-slate-300 hover:bg-slate-100'
                      }`}
                    >
                      <input
                        type="checkbox"
                        checked={isTicked}
                        onChange={() => {}} // handled by button
                        className="w-3.5 h-3.5 pointer-events-none rounded border-slate-300 text-amber-600 focus:ring-0 cursor-pointer"
                      />
                      <span className="font-semibold">{w.name}</span>
                      <span className={`text-[10px] px-1 rounded font-mono ${
                        isTicked ? 'bg-amber-700/70 text-amber-100' : 'bg-slate-200 text-slate-600'
                      }`}>
                        {w.orderCount}
                      </span>
                    </button>
                  );
                })
              )}
            </div>
          </div>
        )}

        {/* Grand Total Metrics Summary */}
        <div className="px-6 py-2.5 bg-white border-b border-slate-200 grid grid-cols-2 sm:grid-cols-7 gap-2">
          <div className="p-2 bg-slate-50 border border-slate-200 rounded-lg">
            <span className="text-[10px] font-bold text-slate-500 uppercase tracking-wider block">Selected Buyers</span>
            <span className="text-base font-extrabold text-slate-900">{grandTotals.totalBuyers}</span>
          </div>
          <div className="p-2 bg-amber-50/70 border border-amber-200 rounded-lg">
            <span className="text-[10px] font-bold text-amber-800 uppercase tracking-wider block">Selected Weeks</span>
            <span className="text-base font-extrabold text-amber-900">{selectedWeeks.length}</span>
          </div>
          <div className="p-2 bg-slate-50 border border-slate-200 rounded-lg">
            <span className="text-[10px] font-bold text-indigo-700 uppercase tracking-wider block">Total VPOs</span>
            <span className="text-base font-extrabold text-indigo-900">{grandTotals.totalVPOs}</span>
          </div>
          <div className="p-2 bg-slate-50 border border-slate-200 rounded-lg">
            <span className="text-[10px] font-bold text-slate-500 uppercase tracking-wider block">Total Orders</span>
            <span className="text-base font-extrabold text-slate-900">{grandTotals.totalOrders}</span>
          </div>
          <div className="p-2 bg-slate-50 border border-slate-200 rounded-lg">
            <span className="text-[10px] font-bold text-slate-500 uppercase tracking-wider block">Total Order Qty</span>
            <span className="text-base font-extrabold text-slate-900">{grandTotals.totalCOQty.toLocaleString()}</span>
          </div>
          <div className="p-2 bg-emerald-50/70 border border-emerald-200 rounded-lg">
            <span className="text-[10px] font-bold text-emerald-700 uppercase tracking-wider block">Total Sew Out</span>
            <span className="text-base font-extrabold text-emerald-800">{grandTotals.totalSewOutQty.toLocaleString()}</span>
          </div>
          <div className="p-2 bg-indigo-50/70 border border-indigo-200 rounded-lg col-span-2 sm:col-span-1">
            <span className="text-[10px] font-bold text-indigo-700 uppercase tracking-wider block">Overall Sew Out %</span>
            <span className="text-base font-extrabold text-indigo-800">{grandTotals.overallPct.toFixed(1)}%</span>
          </div>
        </div>

        {/* Content Body: Switchable between Interactive Table and Live Print Preview */}
        <div className="flex-1 overflow-y-auto p-4 sm:p-6 bg-slate-100/70">
          
          {viewTab === 'print_preview' ? (
            /* Live On-Screen Print Preview (A4 Paper Representation) */
            <div className="max-w-5xl mx-auto bg-white p-8 rounded-xl shadow-lg border border-slate-300">
              <div className="flex items-center justify-between pb-3 mb-4 border-b-2 border-slate-900">
                <div>
                  <h1 className="text-lg font-extrabold text-slate-900">
                    {reportMode === 'sewout_50_100' ? 'SEWING OUT 50% – 100% REPORT (BUYER & VPO WISE)' : 'PRODUCTION REPORT (BUYER & VPO WISE)'}
                  </h1>
                  <p className="text-xs text-slate-600">Separated by VPO No under each Buyer &bull; {minPct}% – {maxPct}%</p>
                </div>
                <div className="text-right text-xs text-slate-500">
                  <div>Orientation: <strong>{printOrientation.toUpperCase()}</strong></div>
                  <div>Buyers: <strong>{buyerGroups.length}</strong> &bull; VPOs: <strong>{grandTotals.totalVPOs}</strong></div>
                </div>
              </div>

              {/* Render Buyer Groups in Print Preview */}
              <div className="space-y-6">
                {buyerGroups.map(buyerGroup => (
                  <div key={buyerGroup.buyer} className="border border-slate-300 rounded-lg overflow-hidden">
                    <div className="bg-slate-900 text-white px-3 py-2 flex justify-between items-center text-xs font-bold">
                      <span>BUYER: {buyerGroup.buyer.toUpperCase()}</span>
                      <span>Order Qty: {buyerGroup.totalCOQty.toLocaleString()} | Sew Out: {buyerGroup.totalSewOutQty.toLocaleString()} | Rate: {buyerGroup.overallPct.toFixed(1)}%</span>
                    </div>

                    <div className="p-3 space-y-4 bg-slate-50/50">
                      {buyerGroup.vpoGroups.map(vpoGroup => (
                        <div key={vpoGroup.vpoNo} className="border border-slate-300 rounded overflow-hidden bg-white">
                          <div className="bg-slate-800 text-white px-3 py-1.5 flex justify-between items-center text-[11px] font-semibold">
                            <span className="text-sky-300 font-bold">VPO NO: {vpoGroup.vpoNo} &bull; Style: {vpoGroup.styleNo}</span>
                            <span>Order: {vpoGroup.totalCOQty.toLocaleString()} | SewOut: {vpoGroup.totalSewOutQty.toLocaleString()} | Rate: {vpoGroup.overallPct.toFixed(1)}%</span>
                          </div>

                          <table className="w-full text-[11px] text-left border-collapse">
                            <thead>
                              <tr className="bg-slate-100 text-slate-600 font-bold border-b border-slate-200">
                                <th className="p-1 text-center w-6">#</th>
                                <th className="p-1">Schedule</th>
                                <th className="p-1">Color</th>
                                <th className="p-1">Destination</th>
                                <th className="p-1 text-center">Plan Date</th>
                                <th className="p-1 text-center">Week</th>
                                <th className="p-1 text-right">CO Qty</th>
                                <th className="p-1 text-right">Sew Out</th>
                                <th className="p-1 text-right">Pending</th>
                                <th className="p-1 text-right">SewOut %</th>
                                <th className="p-1 text-center">Status</th>
                              </tr>
                            </thead>
                            <tbody>
                              {vpoGroup.items.map((item, idx) => {
                                const co = Number(item.coQty) || 0;
                                const sewOut = Number(item.cumSewOutQty) || 0;
                                const pending = Math.max(0, co - sewOut);
                                const pct = co > 0 ? ((sewOut / co) * 100).toFixed(1) : '0.0';
                                return (
                                  <tr key={idx} className={idx % 2 === 1 ? 'bg-slate-50' : 'bg-white'}>
                                    <td className="p-1 text-center text-slate-400">{idx + 1}</td>
                                    <td className="p-1 font-medium">{item.scheduleNo || '-'}</td>
                                    <td className="p-1">{item.colorCode} - {item.colorName}</td>
                                    <td className="p-1">{item.destination}</td>
                                    <td className="p-1 text-center">{item.planDelDate}</td>
                                    <td className="p-1 text-center">{item.weekNo}</td>
                                    <td className="p-1 text-right font-medium">{co.toLocaleString()}</td>
                                    <td className="p-1 text-right font-bold text-emerald-700">{sewOut.toLocaleString()}</td>
                                    <td className="p-1 text-right font-medium text-amber-700">{pending.toLocaleString()}</td>
                                    <td className="p-1 text-right font-bold text-blue-700">{pct}%</td>
                                    <td className="p-1 text-center text-[10px]">{item.statusText}</td>
                                  </tr>
                                );
                              })}
                            </tbody>
                            <tfoot>
                              <tr className="bg-slate-100 font-bold border-t border-slate-300 text-slate-800">
                                <td colSpan={6} className="p-1 text-right">Subtotal for VPO {vpoGroup.vpoNo}:</td>
                                <td className="p-1 text-right">{vpoGroup.totalCOQty.toLocaleString()}</td>
                                <td className="p-1 text-right text-emerald-700">{vpoGroup.totalSewOutQty.toLocaleString()}</td>
                                <td className="p-1 text-right text-amber-700">{vpoGroup.totalPendingQty.toLocaleString()}</td>
                                <td className="p-1 text-right text-blue-700">{vpoGroup.overallPct.toFixed(1)}%</td>
                                <td></td>
                              </tr>
                            </tfoot>
                          </table>
                        </div>
                      ))}
                    </div>
                  </div>
                ))}
              </div>
            </div>
          ) : (
            /* Interactive Structured View */
            <div className="space-y-6">
              {buyerGroups.length > 0 && (
                <div className="flex items-center justify-between text-xs text-slate-500 px-1">
                  <span>Showing <strong>{buyerGroups.length}</strong> Buyer Group{buyerGroups.length > 1 ? 's' : ''} &bull; <strong>{grandTotals.totalVPOs}</strong> VPO{grandTotals.totalVPOs > 1 ? 's' : ''} &bull; {grandTotals.totalOrders} lines</span>
                  <div className="flex items-center gap-2">
                    <button
                      onClick={expandAll}
                      className="text-indigo-600 hover:text-indigo-800 font-semibold hover:underline"
                    >
                      Expand All
                    </button>
                    <span>&bull;</span>
                    <button
                      onClick={collapseAll}
                      className="text-indigo-600 hover:text-indigo-800 font-semibold hover:underline"
                    >
                      Collapse All
                    </button>
                  </div>
                </div>
              )}

              {buyerGroups.length === 0 ? (
                <div className="bg-white rounded-2xl border border-slate-200 p-12 text-center shadow-sm">
                  <div className="w-12 h-12 bg-slate-100 rounded-full flex items-center justify-center mx-auto text-slate-400 mb-3">
                    <Search className="w-6 h-6" />
                  </div>
                  <h3 className="text-base font-bold text-slate-800 mb-1">No Orders Match the Filter</h3>
                  <p className="text-xs text-slate-500 max-w-md mx-auto mb-4">
                    {selectedBuyers.length === 0
                      ? 'No buyers are selected. Please tick at least one buyer.'
                      : selectedWeeks.length === 0
                      ? 'No weeks are selected. Please tick at least one week.'
                      : `Make sure buyers and weeks are ticked and orders have sewing out within ${minPct}% – ${maxPct}%.`}
                  </p>
                  <div className="flex items-center justify-center gap-2">
                    {selectedBuyers.length === 0 && (
                      <button
                        onClick={selectAllBuyers}
                        className="px-4 py-2 bg-indigo-600 text-white rounded-lg text-xs font-semibold shadow hover:bg-indigo-700"
                      >
                        Select All Buyers
                      </button>
                    )}
                    {selectedWeeks.length === 0 && (
                      <button
                        onClick={selectAllWeeks}
                        className="px-4 py-2 bg-amber-600 text-white rounded-lg text-xs font-semibold shadow hover:bg-amber-700"
                      >
                        Select All Weeks
                      </button>
                    )}
                  </div>
                </div>
              ) : (
                buyerGroups.map((buyerGroup) => {
                  const isExpanded = isBuyerExpanded(buyerGroup.buyer);

                  return (
                    <div
                      key={buyerGroup.buyer}
                      className="bg-white rounded-2xl border border-slate-200/90 shadow-sm overflow-hidden transition-all duration-200 hover:shadow-md"
                    >
                      {/* BUYER HEADER CARD */}
                      <div
                        onClick={() => toggleBuyerExpand(buyerGroup.buyer)}
                        className="px-5 py-3.5 bg-gradient-to-r from-slate-900 to-slate-800 text-white flex items-center justify-between cursor-pointer select-none"
                      >
                        <div className="flex items-center gap-3">
                          <button
                            type="button"
                            className="p-1 rounded bg-slate-800 text-slate-300 hover:text-white"
                          >
                            <ChevronDown className={`w-4 h-4 transition-transform duration-200 ${isExpanded ? '' : '-rotate-90'}`} />
                          </button>
                          <div>
                            <div className="flex items-center gap-2">
                              <span className="font-extrabold text-base tracking-wide text-white uppercase">
                                {buyerGroup.buyer}
                              </span>
                              <span className="text-[11px] px-2 py-0.5 rounded-full bg-indigo-500/20 text-indigo-300 font-semibold border border-indigo-400/30">
                                {buyerGroup.totalVPOs} VPO{buyerGroup.totalVPOs > 1 ? 's' : ''}
                              </span>
                              <span className="text-[11px] px-2 py-0.5 rounded-full bg-slate-700/60 text-slate-300 font-semibold">
                                {buyerGroup.totalOrders} Lines
                              </span>
                            </div>
                          </div>
                        </div>

                        {/* Buyer Right KPIs */}
                        <div className="flex items-center gap-4 text-xs font-medium">
                          <div className="hidden sm:block text-right">
                            <span className="text-slate-400 block text-[10px] uppercase font-bold">Order Qty</span>
                            <span className="font-extrabold text-white text-sm">{buyerGroup.totalCOQty.toLocaleString()}</span>
                          </div>
                          <div className="text-right">
                            <span className="text-emerald-400 block text-[10px] uppercase font-bold">Sew Out</span>
                            <span className="font-extrabold text-emerald-400 text-sm">{buyerGroup.totalSewOutQty.toLocaleString()}</span>
                          </div>
                          <div className="hidden sm:block text-right">
                            <span className="text-amber-400 block text-[10px] uppercase font-bold">Pending</span>
                            <span className="font-extrabold text-amber-300 text-sm">{buyerGroup.totalPendingQty.toLocaleString()}</span>
                          </div>
                          <div className="text-right pl-3 border-l border-slate-700">
                            <span className="text-sky-300 block text-[10px] uppercase font-bold">Progress</span>
                            <span className="font-extrabold text-sky-400 text-sm">{buyerGroup.overallPct.toFixed(1)}%</span>
                          </div>
                        </div>
                      </div>

                      {/* BUYER CONTENT: STRICTLY PARTITIONED BY VPO NO */}
                      {isExpanded && (
                        <div className="p-4 sm:p-5 space-y-4 bg-slate-50/50">
                          {buyerGroup.vpoGroups.map((vpoGroup) => {
                            const vpoKey = `${buyerGroup.buyer}_${vpoGroup.vpoNo}`;
                            const isVPOOpen = isVPOExpanded(vpoKey);

                            return (
                              <div
                                key={vpoGroup.vpoNo}
                                className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden transition-all hover:border-slate-300"
                              >
                                {/* VPO HEADER BAR */}
                                <div
                                  onClick={() => toggleVPOExpand(vpoKey)}
                                  className="px-4 py-2.5 bg-slate-100 hover:bg-slate-200/70 border-b border-slate-200 flex flex-wrap items-center justify-between gap-2 cursor-pointer select-none transition-colors"
                                >
                                  <div className="flex items-center gap-2.5">
                                    <button
                                      type="button"
                                      className="p-0.5 rounded text-slate-500 hover:text-slate-800"
                                    >
                                      <ChevronDown className={`w-3.5 h-3.5 transition-transform duration-200 ${isVPOOpen ? '' : '-rotate-90'}`} />
                                    </button>
                                    <span className="font-extrabold text-sm text-indigo-950 font-mono tracking-tight bg-indigo-100/80 px-2 py-0.5 rounded border border-indigo-200/60">
                                      VPO: {vpoGroup.vpoNo}
                                    </span>
                                    <span className="text-xs text-slate-700 font-semibold">
                                      Style: <strong className="text-slate-900">{vpoGroup.styleNo}</strong>
                                      {vpoGroup.custStyleNo && <span className="text-slate-500 font-normal"> ({vpoGroup.custStyleNo})</span>}
                                    </span>
                                    <span className="text-[11px] text-slate-400 font-medium">
                                      ({vpoGroup.items.length} schedule{vpoGroup.items.length > 1 ? 's' : ''})
                                    </span>
                                  </div>

                                  <div className="flex items-center gap-3 text-xs">
                                    <span className="text-slate-600">
                                      Order: <strong className="text-slate-900">{vpoGroup.totalCOQty.toLocaleString()}</strong>
                                    </span>
                                    <span className="text-emerald-700 font-semibold">
                                      Sew Out: <strong>{vpoGroup.totalSewOutQty.toLocaleString()}</strong>
                                    </span>
                                    <span className="text-amber-700 font-semibold">
                                      Pending: <strong>{vpoGroup.totalPendingQty.toLocaleString()}</strong>
                                    </span>
                                    <span className="px-2 py-0.5 rounded-full bg-indigo-50 text-indigo-700 font-bold border border-indigo-200">
                                      {vpoGroup.overallPct.toFixed(1)}%
                                    </span>
                                  </div>
                                </div>

                                {/* VPO TABLE DATA */}
                                {isVPOOpen && (
                                  <div className="overflow-x-auto">
                                    <table className="w-full text-left text-xs border-collapse">
                                      <thead className="text-[10px] uppercase font-bold tracking-wider text-slate-500 bg-slate-50/80 border-b border-slate-200">
                                        <tr>
                                          <th className="py-2 px-3 text-center w-8">#</th>
                                          <th className="py-2 px-3">Schedule No</th>
                                          <th className="py-2 px-3">Color</th>
                                          <th className="py-2 px-3">Destination</th>
                                          <th className="py-2 px-3 text-center">Plan Del Date</th>
                                          <th className="py-2 px-3 text-center">Week</th>
                                          <th className="py-2 px-3 text-right">Order Qty</th>
                                          <th className="py-2 px-3 text-right">Sew In</th>
                                          <th className="py-2 px-3 text-right">Cum Sew Out</th>
                                          <th className="py-2 px-3 text-right">Pending</th>
                                          <th className="py-2 px-3 text-right">SewOut %</th>
                                          <th className="py-2 px-3 text-center">Status</th>
                                          <th className="py-2 px-3">Remark</th>
                                        </tr>
                                      </thead>
                                      <tbody className="divide-y divide-slate-100">
                                        {vpoGroup.items.map((item, idx) => {
                                          const co = Number(item.coQty) || 0;
                                          const sewIn = Number(item.cumSewInQty) || 0;
                                          const sewOut = Number(item.cumSewOutQty) || 0;
                                          const pending = Math.max(0, co - sewOut);
                                          const pct = co > 0 ? (sewOut / co) * 100 : 0;
                                          const remarkText = remarks[item.id]?.text || remarks[item.legacyId || '']?.text || '';

                                          return (
                                            <tr key={idx} className="hover:bg-slate-50/80 transition-colors">
                                              <td className="py-2 px-3 text-center text-slate-400 font-mono text-[11px]">{idx + 1}</td>
                                              <td className="py-2 px-3 font-semibold text-slate-800">{item.scheduleNo || '-'}</td>
                                              <td className="py-2 px-3 text-slate-700">
                                                <span className="font-mono text-[11px]">{item.colorCode}</span>
                                                {item.colorName && <span className="ml-1 text-slate-500 font-sans">({item.colorName})</span>}
                                              </td>
                                              <td className="py-2 px-3 text-slate-600">{item.destination || '-'}</td>
                                              <td className="py-2 px-3 text-center text-slate-600 font-mono">{item.planDelDate || '-'}</td>
                                              <td className="py-2 px-3 text-center text-slate-700 font-medium">{item.weekNo || '-'}</td>
                                              <td className="py-2 px-3 text-right font-medium text-slate-900">{co.toLocaleString()}</td>
                                              <td className="py-2 px-3 text-right text-slate-600">{sewIn.toLocaleString()}</td>
                                              <td className="py-2 px-3 text-right font-bold text-emerald-600">{sewOut.toLocaleString()}</td>
                                              <td className="py-2 px-3 text-right font-medium text-amber-700">{pending.toLocaleString()}</td>
                                              <td className="py-2 px-3 text-right">
                                                <span className={`inline-block px-1.5 py-0.5 rounded text-[11px] font-bold ${
                                                  pct >= 100 
                                                    ? 'bg-emerald-100 text-emerald-800' 
                                                    : pct >= 80 
                                                    ? 'bg-blue-100 text-blue-800' 
                                                    : 'bg-amber-100 text-amber-800'
                                                }`}>
                                                  {pct.toFixed(1)}%
                                                </span>
                                              </td>
                                              <td className="py-2 px-3 text-center">
                                                <span className="text-[10px] px-1.5 py-0.5 rounded bg-slate-100 text-slate-700 font-medium">
                                                  {item.statusText || '-'}
                                                </span>
                                              </td>
                                              <td className="py-2 px-3 text-slate-500 text-[11px] max-w-xs truncate" title={remarkText}>
                                                {remarkText || '-'}
                                              </td>
                                            </tr>
                                          );
                                        })}
                                      </tbody>
                                      <tfoot>
                                        <tr className="bg-slate-50/90 font-bold text-slate-800 border-t border-slate-200">
                                          <td colSpan={6} className="py-2 px-3 text-right text-slate-500 font-medium text-[11px]">
                                            Subtotal for VPO {vpoGroup.vpoNo}:
                                          </td>
                                          <td className="py-2 px-3 text-right">{vpoGroup.totalCOQty.toLocaleString()}</td>
                                          <td className="py-2 px-3 text-right text-slate-400">-</td>
                                          <td className="py-2 px-3 text-right text-emerald-700 font-extrabold">{vpoGroup.totalSewOutQty.toLocaleString()}</td>
                                          <td className="py-2 px-3 text-right text-amber-700">{vpoGroup.totalPendingQty.toLocaleString()}</td>
                                          <td className="py-2 px-3 text-right text-indigo-700 font-extrabold">{vpoGroup.overallPct.toFixed(1)}%</td>
                                          <td colSpan={2}></td>
                                        </tr>
                                      </tfoot>
                                    </table>
                                  </div>
                                )}
                              </div>
                            );
                          })}

                          {/* BUYER SUMMARY FOOTER */}
                          <div className="bg-slate-200/70 border border-slate-300 rounded-xl px-4 py-2.5 flex flex-wrap items-center justify-between text-xs font-bold text-slate-800">
                            <span>TOTAL FOR BUYER: {buyerGroup.buyer.toUpperCase()} ({buyerGroup.totalVPOs} VPOs)</span>
                            <div className="flex items-center gap-4">
                              <span>Order Qty: <strong className="text-slate-900">{buyerGroup.totalCOQty.toLocaleString()}</strong></span>
                              <span className="text-emerald-700">Sew Out: <strong>{buyerGroup.totalSewOutQty.toLocaleString()}</strong></span>
                              <span className="text-amber-800">Pending: <strong>{buyerGroup.totalPendingQty.toLocaleString()}</strong></span>
                              <span className="text-indigo-800 font-extrabold">Overall: {buyerGroup.overallPct.toFixed(1)}%</span>
                            </div>
                          </div>
                        </div>
                      )}
                    </div>
                  );
                })
              )}
            </div>
          )}

        </div>

        {/* Modal Footer */}
        <div className="px-6 py-3 bg-slate-50 border-t border-slate-200 flex flex-wrap items-center justify-between gap-3 text-xs">
          <div className="flex items-center gap-2 text-slate-500">
            <span className="font-semibold text-slate-700">Tips:</span>
            <span>Tick or untick any buyer to customize. Every VPO prints with dedicated subtotal tables.</span>
          </div>
          <div className="flex items-center gap-2.5">
            <button
              onClick={onClose}
              className="px-4 py-2 bg-slate-200 hover:bg-slate-300 text-slate-700 rounded-lg font-semibold transition-colors"
            >
              Close
            </button>
            <button
              onClick={handleOpenPrintWindow}
              disabled={buyerGroups.length === 0}
              className="px-5 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-lg font-bold flex items-center gap-1.5 shadow transition-all disabled:opacity-50"
            >
              <Printer className="w-4 h-4" />
              Print Report
            </button>
          </div>
        </div>

      </div>
    </div>
  );
}
