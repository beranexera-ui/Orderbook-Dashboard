import React, { useState, useMemo } from 'react';
import { Printer, X, FileText, CheckCircle2, Sliders, LayoutGrid, Check } from 'lucide-react';
import { ProductionOrder } from './Dashboard';

interface PrintOrientationModalProps {
  isOpen: boolean;
  onClose: () => void;
  data: ProductionOrder[];
  remarks: Record<string, { text: string; updatedAt: number | null }>;
  activeFilters?: {
    search?: string;
    buyers?: string[];
    weeks?: string[];
    statuses?: string[];
    shipmentModes?: string[];
    destinations?: string[];
    packMethods?: string[];
    remarksFilter?: string[];
  };
}

export function PrintOrientationModal({
  isOpen,
  onClose,
  data,
  remarks,
  activeFilters
}: PrintOrientationModalProps) {
  const [orientation, setOrientation] = useState<'portrait' | 'landscape'>('landscape');
  const includeSummary = false;

  // Compute summary totals for filtered items
  const totals = useMemo(() => {
    let totalCO = 0;
    let totalSewIn = 0;
    let totalSewOut = 0;
    let totalDelivered = 0;
    const buyersSet = new Set<string>();
    const vpoSet = new Set<string>();

    data.forEach(item => {
      totalCO += Number(item.coQty) || 0;
      if (item.buyer) buyersSet.add(item.buyer);
      if (item.vpoNo) vpoSet.add(item.vpoNo);
    });

    return {
      count: data.length,
      buyersCount: buyersSet.size,
      vpoCount: vpoSet.size,
      totalCO
    };
  }, [data]);

  // Build active filter tags for display and print header
  const filterDescriptions = useMemo(() => {
    if (!activeFilters) return [];
    const list: string[] = [];
    if (activeFilters.search?.trim()) list.push(`Search: "${activeFilters.search.trim()}"`);
    if (activeFilters.buyers && activeFilters.buyers.length > 0) list.push(`Buyers: ${activeFilters.buyers.join(', ')}`);
    if (activeFilters.weeks && activeFilters.weeks.length > 0) list.push(`Weeks: ${activeFilters.weeks.join(', ')}`);
    if (activeFilters.statuses && activeFilters.statuses.length > 0) list.push(`Status: ${activeFilters.statuses.join(', ')}`);
    if (activeFilters.shipmentModes && activeFilters.shipmentModes.length > 0) list.push(`Shipment: ${activeFilters.shipmentModes.join(', ')}`);
    if (activeFilters.destinations && activeFilters.destinations.length > 0) list.push(`Destination: ${activeFilters.destinations.join(', ')}`);
    if (activeFilters.packMethods && activeFilters.packMethods.length > 0) list.push(`Pack Method: ${activeFilters.packMethods.join(', ')}`);
    if (activeFilters.remarksFilter && activeFilters.remarksFilter.length > 0) list.push(`Remark: ${activeFilters.remarksFilter.join(', ')}`);
    return list;
  }, [activeFilters]);

  // Group data by VPO so each VPO is clearly separated with its own header & subtotal
  const vpoGroups = useMemo(() => {
    if (!data || data.length === 0) return [];

    // Map: VPO No -> Items
    const vpoMap = new Map<string, ProductionOrder[]>();

    data.forEach(item => {
      const vpoKey = (item.vpoNo && item.vpoNo.trim()) ? item.vpoNo.trim() : 'UNASSIGNED VPO';
      if (!vpoMap.has(vpoKey)) {
        vpoMap.set(vpoKey, []);
      }
      vpoMap.get(vpoKey)!.push(item);
    });

    const groups: {
      vpoNo: string;
      buyer: string;
      styleNo: string;
      items: ProductionOrder[];
      totalCO: number;
    }[] = [];

    vpoMap.forEach((items, vpoNo) => {
      const totalCO = items.reduce((acc, curr) => acc + (Number(curr.coQty) || 0), 0);
      const buyer = items[0]?.buyer || '-';
      const styleNo = items[0]?.styleNo || '-';
      groups.push({
        vpoNo,
        buyer,
        styleNo,
        items,
        totalCO
      });
    });

    return groups;
  }, [data]);

  // Generate Clean Black & White A4 HTML
  const generateBWPrintHTML = (selectedOrientation: 'portrait' | 'landscape') => {
    const isPortrait = selectedOrientation === 'portrait';

    // Generous, crisp font sizes and comfortable box height
    const cellFontSize = isPortrait ? '8pt' : '9.5pt';
    const headerFontSize = isPortrait ? '8.5pt' : '9.5pt';
    const cellPadding = isPortrait ? '5px 5.5px' : '6px 7.5px';
    const grandTotalFontSize = isPortrait ? '9.5pt' : '11pt';

    const now = new Date();
    const dateFormatted = now.toLocaleDateString('en-US', {
      year: 'numeric',
      month: 'short',
      day: '2-digit'
    });
    const timeFormatted = now.toLocaleTimeString('en-US', {
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit'
    });

    let overallRowIndex = 0;

    // Render continuous clean table rows. Separate different VPOs with a clear subtle boundary & subtotal
    const tableBodiesHTML = vpoGroups.map((group, gIdx) => {
      const groupRowsHTML = group.items.map(item => {
        overallRowIndex++;
        const co = Number(item.coQty) || 0;
        const colorDisplay = item.colorCode && item.colorName
          ? `${item.colorCode}-${item.colorName}`
          : (item.colorCode || item.colorName || '-');
        const remark = remarks[item.id]?.text || remarks[item.legacyId || '']?.text || '';

        return `
          <tr>
            <td style="text-align: center; white-space: nowrap;">${item.planDelDate || '-'}</td>
            <td style="text-align: center; white-space: nowrap;">${item.weekNo || '-'}</td>
            <td style="text-align: center; white-space: nowrap;">${co > 0 ? co.toLocaleString() : '-'}</td>
            <td style="text-align: center; white-space: nowrap;">${item.styleNo || '-'}</td>
            <td style="font-weight: 800; text-align: center; white-space: nowrap;">${item.vpoNo || '-'}</td>
            <td style="text-align: center; white-space: nowrap;">${item.destination || '-'}</td>
            <td style="text-align: center; white-space: nowrap;">${colorDisplay}</td>
            <td style="font-weight: 800; text-align: center; white-space: nowrap;">${item.scheduleNo || '-'}</td>
            <td style="text-align: center; white-space: nowrap;">${item.shipmentMode || '-'}</td>
            <td style="text-align: left; word-break: break-word;">${remark}</td>
          </tr>
        `;
      }).join('');

      return `
        <tbody class="vpo-tbody">
          <!-- Subtle separator space between different VPOs -->
          ${gIdx > 0 ? `
          <tr class="vpo-spacer-row">
            <td colspan="10"></td>
          </tr>
          ` : ''}

          <!-- Data rows for this VPO -->
          ${groupRowsHTML}
        </tbody>
      `;
    }).join('');

    const filterString = filterDescriptions.length > 0
      ? filterDescriptions.join(' | ')
      : 'All Records (No Filters Applied)';

    return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>Production_Report_A4_${selectedOrientation.toUpperCase()}_${data.length}_items</title>
  <style>
    @page {
      size: A4 ${selectedOrientation};
      margin: 8mm 6mm 8mm 6mm;
    }
    * {
      box-sizing: border-box;
      -webkit-print-color-adjust: exact !important;
      print-color-adjust: exact !important;
    }
    html, body {
      margin: 0;
      padding: 0;
      background: #ffffff;
      color: #000000;
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Arial, sans-serif;
    }
    body {
      padding: 8px 12px;
    }
    /* Toolbar visible only in browser view, hidden when printed */
    .no-print {
      display: flex;
      justify-content: space-between;
      align-items: center;
      background: #f1f5f9;
      border: 1px solid #cbd5e1;
      padding: 8px 14px;
      margin-bottom: 12px;
      border-radius: 6px;
      font-size: 12px;
      font-family: sans-serif;
    }
    .print-btn {
      background: #0f172a;
      color: #ffffff;
      border: 1px solid #0f172a;
      padding: 6px 16px;
      font-size: 12px;
      font-weight: bold;
      border-radius: 4px;
      cursor: pointer;
      display: inline-flex;
      align-items: center;
      gap: 6px;
    }
    .print-btn:hover {
      background: #334155;
    }
    @media print {
      .no-print {
        display: none !important;
      }
      body {
        padding: 0 !important;
      }
      table {
        page-break-inside: auto !important;
      }
      tr {
        page-break-inside: avoid !important;
        break-inside: avoid !important;
      }
      thead {
        display: table-header-group !important;
      }
      tfoot {
        display: table-row-group !important;
        page-break-inside: avoid !important;
      }
    }

    /* Black & White Clean Layout */
    .header-table {
      width: 100%;
      border-collapse: collapse;
      margin-bottom: 6px;
      border-bottom: 1.5px solid #000000;
      padding-bottom: 4px;
    }
    .header-table td {
      border: none !important;
      padding: 1px 0;
      vertical-align: bottom;
    }
    .main-title {
      font-size: 12pt;
      font-weight: 900;
      letter-spacing: 0.5px;
      text-transform: uppercase;
      margin: 0 0 2px 0;
      color: #000000;
    }
    .sub-title {
      font-size: 7.5pt;
      color: #333333;
      margin: 0;
    }
    .meta-info {
      text-align: right;
      font-size: 7pt;
      color: #222222;
      line-height: 1.35;
    }

    /* Summary Bar */
    .summary-box {
      width: 100%;
      border: 1px solid #000000;
      background-color: #fafafa;
      margin-bottom: 8px;
      font-size: 7.5pt;
      border-collapse: collapse;
    }
    .summary-box td {
      border: 0.5px solid #cccccc;
      padding: 4px 6px;
      text-align: center;
    }
    .summary-label {
      font-size: 6.5pt;
      text-transform: uppercase;
      font-weight: 600;
      color: #444444;
      display: block;
    }
    .summary-val {
      font-size: 9pt;
      font-weight: bold;
      color: #000000;
    }

    /* Unified Data Table - Black & White with Clean Proportions matching user spec */
    table.data-table {
      width: 100%;
      border-collapse: collapse;
      font-size: ${cellFontSize};
      border: 2px solid #000000;
      margin-bottom: 6px;
    }
    table.data-table th, table.data-table td {
      border: 1px solid #000000;
      padding: ${cellPadding};
      vertical-align: middle;
      line-height: 1.35;
    }
    table.data-table th {
      background-color: #e2e8f0 !important;
      color: #000000;
      font-weight: 800;
      font-size: ${headerFontSize};
      text-align: center;
      white-space: nowrap;
      border: 1px solid #000000;
    }
    table.data-table tr td {
      background-color: #ffffff !important;
    }

    /* Subtle clean separator space between different VPOs - Double border like user image */
    tr.vpo-spacer-row td {
      height: 4px !important;
      padding: 0 !important;
      background-color: #cbd5e1 !important;
      border-left: 1px solid #000000 !important;
      border-right: 1px solid #000000 !important;
      border-top: 2px solid #000000 !important;
      border-bottom: 2px solid #000000 !important;
    }

    /* Grand Total Row directly in table footer */
    tr.grand-total-row td {
      background-color: #e2e8f0 !important;
      font-weight: 800;
      border-top: 2px solid #000000 !important;
      border-bottom: 2px solid #000000 !important;
      font-size: ${grandTotalFontSize};
      padding: ${cellPadding};
    }

    .page-footer {
      margin-top: 8px;
      font-size: 6.5pt;
      color: #555555;
      display: flex;
      justify-content: space-between;
      border-top: 0.5px solid #888888;
      padding-top: 3px;
    }
  </style>
  <script>
    window.addEventListener('load', function() {
      setTimeout(function() {
        window.focus();
        window.print();
      }, 400);
    });
  </script>
</head>
<body>
  <!-- Browser top toolbar -->
  <div class="no-print">
    <div>
      <strong>A4 Black & White Print Preview (${selectedOrientation.toUpperCase()})</strong> &bull; ${data.length} orders across ${totals.vpoCount} VPOs
    </div>
    <div style="display: flex; gap: 8px;">
      <button class="print-btn" onclick="window.print()">
        🖨️ Print Now / Save PDF
      </button>
      <button class="print-btn" style="background: #475569; border-color: #475569;" onclick="window.close()">
        ✕ Close
      </button>
    </div>
  </div>

  <!-- Report Header -->
  <table class="header-table">
    <tr>
      <td>
        <h1 class="main-title">PRODUCTION ORDER REPORT (BY VPO)</h1>
        <p class="sub-title">DILEEPA WICKRAMASINGHE PRODUCTION SYSTEM &bull; FORMAT: A4 ${selectedOrientation.toUpperCase()}</p>
        <p class="sub-title" style="margin-top: 2px; font-weight: 600;">Filters: ${filterString}</p>
      </td>
      <td class="meta-info">
        <div>Printed Date: <strong>${dateFormatted}</strong></div>
        <div>Printed Time: <strong>${timeFormatted}</strong></div>
        <div>Page Orientation: <strong>${selectedOrientation.toUpperCase()}</strong></div>
        <div>Total Rows: <strong>${data.length}</strong> (${totals.vpoCount} VPOs)</div>
      </td>
    </tr>
  </table>

  ${includeSummary ? `
  <!-- Summary Box -->
  <table class="summary-box">
    <tr>
      <td>
        <span class="summary-label">Total Filtered Rows</span>
        <span class="summary-val">${totals.count}</span>
      </td>
      <td>
        <span class="summary-label">Total Buyers</span>
        <span class="summary-val">${totals.buyersCount}</span>
      </td>
      <td>
        <span class="summary-label">Total VPOs</span>
        <span class="summary-val">${totals.vpoCount}</span>
      </td>
      <td>
        <span class="summary-label">Total CO Qty</span>
        <span class="summary-val">${totals.totalCO.toLocaleString()}</span>
      </td>
    </tr>
  </table>
  ` : ''}

  <!-- Single Unified Data Table with clean natural fit and subtle VPO spacing -->
  <table class="data-table">
    <thead>
      <tr>
        <th style="text-align: center; white-space: nowrap;">Plan Del Date</th>
        <th style="text-align: center; white-space: nowrap;">WEEK NO</th>
        <th style="text-align: center; white-space: nowrap;">Order Qty</th>
        <th style="text-align: center; white-space: nowrap;">Style</th>
        <th style="text-align: center; white-space: nowrap;">VPO No.</th>
        <th style="text-align: center; white-space: nowrap;">Destination</th>
        <th style="text-align: center; white-space: nowrap;">Color Name</th>
        <th style="text-align: center; white-space: nowrap;">Schedule</th>
        <th style="text-align: center; white-space: nowrap;">Mode</th>
        <th style="text-align: center; white-space: nowrap;">Remarks</th>
      </tr>
    </thead>
    ${tableBodiesHTML}
    <tfoot>
      <tr class="grand-total-row">
        <td colspan="2" style="text-align: right; text-transform: uppercase;">
          OVERALL TOTAL:
        </td>
        <td style="text-align: center; font-weight: 900; white-space: nowrap;">
          ${totals.totalCO.toLocaleString()}
        </td>
        <td colspan="7" style="text-align: left; font-size: 8.5pt; color: #333;">
          (${totals.count} Lines &bull; ${totals.vpoCount} VPOs &bull; ${totals.buyersCount} Buyers)
        </td>
      </tr>
    </tfoot>
  </table>

  <div class="page-footer">
    <span>Production Management Tracking System &bull; DILEEPA WICKRAMASINGHE</span>
    <span>A4 ${selectedOrientation.toUpperCase()} &bull; Black & White Report &bull; VPO Separated &bull; Total ${data.length} Rows</span>
  </div>
</body>
</html>`;
  };

  // Launch Print in New Tab / Window
  const handlePrint = (selectedOrientation: 'portrait' | 'landscape') => {
    if (data.length === 0) return;

    const htmlContent = generateBWPrintHTML(selectedOrientation);
    const blob = new Blob([htmlContent], { type: 'text/html;charset=utf-8' });
    const blobUrl = URL.createObjectURL(blob);

    // Try window.open first
    const printWindow = window.open(blobUrl, '_blank');
    if (!printWindow || printWindow.closed || typeof printWindow.closed === 'undefined') {
      // Fallback: create invisible iframe to trigger print dialog directly
      const iframe = document.createElement('iframe');
      iframe.style.position = 'fixed';
      iframe.style.right = '0';
      iframe.style.bottom = '0';
      iframe.style.width = '0';
      iframe.style.height = '0';
      iframe.style.border = '0';
      document.body.appendChild(iframe);
      iframe.src = blobUrl;
      iframe.onload = () => {
        try {
          iframe.contentWindow?.focus();
          iframe.contentWindow?.print();
        } catch (e) {
          console.error("Iframe print error:", e);
        }
        setTimeout(() => {
          document.body.removeChild(iframe);
          URL.revokeObjectURL(blobUrl);
        }, 60000);
      };
    } else {
      setTimeout(() => URL.revokeObjectURL(blobUrl), 120000);
    }

    onClose();
  };

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/60 backdrop-blur-sm p-4 overflow-y-auto">
      <div className="bg-white w-full max-w-lg rounded-2xl shadow-2xl border border-slate-200 overflow-hidden animate-in fade-in zoom-in-95 duration-200">
        
        {/* Header */}
        <div className="px-6 py-4 bg-slate-900 text-white flex items-center justify-between border-b border-slate-800">
          <div className="flex items-center gap-2.5">
            <span className="p-2 bg-indigo-500/20 text-indigo-400 rounded-lg border border-indigo-500/30">
              <Printer className="w-5 h-5" />
            </span>
            <div>
              <h2 className="text-base font-bold tracking-tight text-white">
                Print Filtered Orders (A4 Black & White)
              </h2>
              <p className="text-xs text-slate-400">
                Choose Portrait or Landscape orientation for your printout
              </p>
            </div>
          </div>
          <button
            onClick={onClose}
            className="p-1.5 text-slate-400 hover:text-white hover:bg-slate-800 rounded-lg transition-colors"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Modal Body */}
        <div className="p-6 space-y-5">
          
          {/* Active Data Info Banner */}
          <div className="bg-slate-50 border border-slate-200 rounded-xl p-3.5 flex items-center justify-between">
            <div>
              <span className="text-xs font-semibold text-slate-500 uppercase tracking-wider block">
                Ready to Print
              </span>
              <span className="text-lg font-bold text-slate-900">
                {data.length.toLocaleString()} {data.length === 1 ? 'Order' : 'Orders'}
              </span>
              <span className="text-xs text-slate-500 ml-2">
                ({totals.buyersCount} Buyers, {totals.vpoCount} VPOs)
              </span>
            </div>
            <div className="text-right">
              <span className="text-xs font-semibold text-slate-500 uppercase tracking-wider block">
                Total CO Qty
              </span>
              <span className="text-base font-extrabold text-indigo-600">
                {totals.totalCO.toLocaleString()}
              </span>
            </div>
          </div>

          {/* Active Filters Summary */}
          {filterDescriptions.length > 0 && (
            <div className="bg-amber-50/70 border border-amber-200/80 rounded-lg p-2.5 text-xs text-amber-900">
              <span className="font-bold">Active Table Filters:</span>
              <p className="text-[11px] text-amber-800 mt-0.5 line-clamp-2">
                {filterDescriptions.join(' • ')}
              </p>
            </div>
          )}

          {/* Orientation Choice Cards */}
          <div>
            <label className="text-xs font-bold text-slate-700 uppercase tracking-wider block mb-2.5">
              Select Orientation:
            </label>
            <div className="grid grid-cols-2 gap-3.5">
              
              {/* Landscape Card */}
              <div
                onClick={() => setOrientation('landscape')}
                className={`p-4 rounded-xl border-2 cursor-pointer transition-all flex flex-col items-center text-center relative ${
                  orientation === 'landscape'
                    ? 'border-indigo-600 bg-indigo-50/40 shadow-sm'
                    : 'border-slate-200 hover:border-slate-300 bg-white'
                }`}
              >
                {orientation === 'landscape' && (
                  <div className="absolute top-2 right-2 w-5 h-5 bg-indigo-600 text-white rounded-full flex items-center justify-center">
                    <Check className="w-3 h-3 stroke-[3]" />
                  </div>
                )}
                {/* Horizontal Page Graphic */}
                <div className="w-16 h-11 border-2 border-slate-700 rounded bg-white shadow-sm flex flex-col justify-center gap-1 p-1 mb-2.5">
                  <div className="w-full h-1 bg-slate-300 rounded-sm"></div>
                  <div className="w-full h-1 bg-slate-200 rounded-sm"></div>
                  <div className="w-3/4 h-1 bg-slate-200 rounded-sm"></div>
                </div>
                <h4 className="text-sm font-bold text-slate-900">Landscape</h4>
                <span className="text-[10px] font-semibold text-emerald-700 bg-emerald-50 px-1.5 py-0.5 rounded border border-emerald-200 mt-1">
                  Recommended
                </span>
                <p className="text-[11px] text-slate-500 mt-1">
                  Wide table (VPO wise separated layout)
                </p>
              </div>

              {/* Portrait Card */}
              <div
                onClick={() => setOrientation('portrait')}
                className={`p-4 rounded-xl border-2 cursor-pointer transition-all flex flex-col items-center text-center relative ${
                  orientation === 'portrait'
                    ? 'border-indigo-600 bg-indigo-50/40 shadow-sm'
                    : 'border-slate-200 hover:border-slate-300 bg-white'
                }`}
              >
                {orientation === 'portrait' && (
                  <div className="absolute top-2 right-2 w-5 h-5 bg-indigo-600 text-white rounded-full flex items-center justify-center">
                    <Check className="w-3 h-3 stroke-[3]" />
                  </div>
                )}
                {/* Vertical Page Graphic */}
                <div className="w-11 h-16 border-2 border-slate-700 rounded bg-white shadow-sm flex flex-col justify-center gap-1.5 p-1 mb-2.5">
                  <div className="w-full h-1 bg-slate-300 rounded-sm"></div>
                  <div className="w-full h-1 bg-slate-200 rounded-sm"></div>
                  <div className="w-full h-1 bg-slate-200 rounded-sm"></div>
                  <div className="w-2/3 h-1 bg-slate-200 rounded-sm"></div>
                </div>
                <h4 className="text-sm font-bold text-slate-900">Portrait</h4>
                <span className="text-[10px] font-semibold text-slate-600 bg-slate-100 px-1.5 py-0.5 rounded border border-slate-200 mt-1">
                  Standard A4
                </span>
                <p className="text-[11px] text-slate-500 mt-1">
                  Vertical format, compact font size
                </p>
              </div>

            </div>
          </div>

        </div>

        {/* Footer with direct action buttons */}
        <div className="px-6 py-4 bg-slate-50 border-t border-slate-200 flex items-center justify-between gap-3">
          <button
            onClick={onClose}
            className="px-4 py-2 bg-white border border-slate-200 hover:bg-slate-100 text-slate-700 rounded-lg text-xs font-semibold transition-colors"
          >
            Cancel
          </button>

          <div className="flex items-center gap-2">
            <button
              onClick={() => handlePrint('portrait')}
              disabled={data.length === 0}
              className={`px-3.5 py-2 rounded-lg text-xs font-bold transition-all flex items-center gap-1.5 ${
                orientation === 'portrait'
                  ? 'bg-indigo-600 hover:bg-indigo-700 text-white shadow-sm'
                  : 'bg-slate-200 hover:bg-slate-300 text-slate-700'
              }`}
            >
              <Printer className="w-3.5 h-3.5" />
              Print Portrait
            </button>

            <button
              onClick={() => handlePrint('landscape')}
              disabled={data.length === 0}
              className={`px-4 py-2 rounded-lg text-xs font-bold transition-all flex items-center gap-1.5 ${
                orientation === 'landscape'
                  ? 'bg-indigo-600 hover:bg-indigo-700 text-white shadow-sm'
                  : 'bg-slate-200 hover:bg-slate-300 text-slate-700'
              }`}
            >
              <Printer className="w-3.5 h-3.5" />
              Print Landscape
            </button>
          </div>
        </div>

      </div>
    </div>
  );
}
