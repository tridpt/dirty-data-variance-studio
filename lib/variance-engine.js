/**
 * 📉 DataScrub & Variance Studio — Động Cơ Giải Phẫu Biến Động (Price-Volume-Mix Bridge & Key-Driver Attribution)
 *
 * Tính năng cốt lõi:
 * 1. Bóc tách chính xác tuyệt đối (Zero-Residual Exact Identity):
 *    R1 - R0 = ΔVolume (Sản lượng) + ΔMix (Cơ cấu mặt hàng) + ΔPrice (Đơn giá/Chiết khấu) + ΔNew (Mới) + ΔDiscontinued (Ngưng)
 * 2. Truy vết Đa chiều (Key-Driver Attribution): Quét toàn bộ các chiều (Chi nhánh, Nhân viên, Kênh, Sản phẩm...)
 *    chỉ đích danh Top 3 Kéo Xuống (Headwinds) & Top 3 Kéo Lên (Tailwinds)
 * 3. Giải phẫu Chéo 2 Tầng (2-Level Root-Cause Drilldown): Bấm vào bất kỳ Chi nhánh / Nhân viên / Kênh nào ->
 *    chạy ngay PVM Bridge cục bộ + chỉ mặt thủ phạm con bên trong phân khúc đó!
 * 4. Kiểm toán Bẫy Dữ Liệu Bẩn (Dirty-vs-Clean Trap): Đo lường mức sai lệch báo cáo nếu tổng hợp trực tiếp trên file bẩn gốc
 * 5. Tự động biên tập "Bản Giải Trình Gửi Sếp" (Executive Memo tiếng Việt súc tích, có số liệu chứng minh)
 */

const { parseVietnameseNumber, cleanInvisibleAndNFC } = require('./viet-cleaner');

// ============================================================================
// 1. HỖ TRỢ KỲ THỜI GIAN (THÁNG / TUẦN / QUÝ) & ĐỊNH DẠNG TIẾNG VIỆT
// ============================================================================

function isoDateToPeriod(isoStr, granularity = 'month') {
  const s = String(isoStr || '').trim();
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return s || '';
  const year = Number(m[1]);
  const month = Number(m[2]);
  const day = Number(m[3]);

  if (granularity === 'quarter') {
    const q = Math.floor((month - 1) / 3) + 1;
    return `${year}-Q${q}`;
  }
  if (granularity === 'week') {
    const dt = new Date(Date.UTC(year, month - 1, day));
    const dayNum = dt.getUTCDay() || 7;
    dt.setUTCDate(dt.getUTCDate() + 4 - dayNum);
    const yearStart = new Date(Date.UTC(dt.getUTCFullYear(), 0, 1));
    const weekNo = Math.ceil(((dt - yearStart) / 86400000 + 1) / 7);
    return `${dt.getUTCFullYear()}-W${String(weekNo).padStart(2, '0')}`;
  }
  return `${year}-${String(month).padStart(2, '0')}`;
}

function extractAvailablePeriods(rows, dateCol, granularity = 'month') {
  const counts = new Map();
  for (const r of rows) {
    const rawDate = dateCol ? r?.[dateCol] : r?._period;
    const p = isoDateToPeriod(rawDate, granularity);
    if (!p) continue;
    counts.set(p, (counts.get(p) || 0) + 1);
  }
  return Array.from(counts.entries())
    .map(([period, rowCount]) => ({ period, rowCount }))
    .sort((a, b) => a.period.localeCompare(b.period));
}

function formatVnCompact(val) {
  if (val === null || val === undefined || !Number.isFinite(val)) return '0';
  const sign = val < 0 ? '-' : '';
  const abs = Math.abs(val);
  if (abs >= 1_000_000_000) {
    return `${sign}${(abs / 1_000_000_000).toLocaleString('vi-VN', { maximumFractionDigits: 2 })} tỷ`;
  }
  if (abs >= 1_000_000) {
    return `${sign}${(abs / 1_000_000).toLocaleString('vi-VN', { maximumFractionDigits: 2 })} tr`;
  }
  if (abs >= 10_000) {
    return `${sign}${Math.round(abs).toLocaleString('vi-VN')}`;
  }
  return `${sign}${abs.toLocaleString('vi-VN', { maximumFractionDigits: 2 })}`;
}

function formatSignedCompact(val) {
  if (!Number.isFinite(val) || val === 0) return '0';
  return val > 0 ? `+${formatVnCompact(val)}` : formatVnCompact(val);
}

function formatPct(pct) {
  if (pct === null || pct === undefined || !Number.isFinite(pct)) return 'N/A';
  const sign = pct > 0 ? '+' : '';
  return `${sign}${pct.toFixed(1)}%`;
}

// ============================================================================
// 2. THUẬT TOÁN PRICE - VOLUME - MIX BRIDGE CHÍNH XÁC TUYỆT ĐỐI
// ============================================================================

function computeExactPvmBridge(rowsP0, rowsP1, { metricCol, quantityCol, itemCol }) {
  const hasExplicitQty =
    Boolean(quantityCol) &&
    [...rowsP0, ...rowsP1].some((r) => typeof r?.[quantityCol] === 'number' && r[quantityCol] > 0);

  const effectiveGroupCol = itemCol || '_SINGLE_PORTFOLIO_';
  const map0 = new Map();
  const map1 = new Map();

  const accumulate = (targetMap, list) => {
    for (const r of list) {
      const mVal = Number(r?.[metricCol]);
      if (!Number.isFinite(mVal)) continue;
      const qVal = hasExplicitQty ? Number(r?.[quantityCol]) : 1;
      if (!Number.isFinite(qVal) || qVal <= 0) continue;

      const key = itemCol ? String(r?.[itemCol] || 'Chưa phân loại').trim() || 'Chưa phân loại' : 'Toàn danh mục';
      const cur = targetMap.get(key) || { item: key, rev: 0, qty: 0, txCount: 0 };
      cur.rev += mVal;
      cur.qty += qVal;
      cur.txCount += 1;
      targetMap.set(key, cur);
    }
  };

  accumulate(map0, rowsP0);
  accumulate(map1, rowsP1);

  const allKeys = new Set([...map0.keys(), ...map1.keys()]);
  const continuing = [];
  const newItems = [];
  const discontinuedItems = [];

  let totalR0 = 0;
  let totalR1 = 0;
  let totalQ0All = 0;
  let totalQ1All = 0;

  for (const key of allKeys) {
    const d0 = map0.get(key) || { item: key, rev: 0, qty: 0, txCount: 0 };
    const d1 = map1.get(key) || { item: key, rev: 0, qty: 0, txCount: 0 };
    totalR0 += d0.rev;
    totalR1 += d1.rev;
    totalQ0All += d0.qty;
    totalQ1All += d1.qty;

    if (d0.qty > 0 && d1.qty > 0) {
      continuing.push({ key, d0, d1 });
    } else if (d0.qty === 0 && d1.qty > 0) {
      newItems.push({ key, d0, d1 });
    } else if (d0.qty > 0 && d1.qty === 0) {
      discontinuedItems.push({ key, d0, d1 });
    }
  }

  const QC0 = continuing.reduce((s, x) => s + x.d0.qty, 0);
  const QC1 = continuing.reduce((s, x) => s + x.d1.qty, 0);
  const RC0 = continuing.reduce((s, x) => s + x.d0.rev, 0);
  const avgP0 = QC0 > 0 ? RC0 / QC0 : 0;

  let volumeEffect = 0;
  let mixEffect = 0;
  let priceEffect = 0;
  let newEffect = 0;
  let discontinuedEffect = 0;

  const itemBreakdown = [];

  for (const { key, d0, d1 } of continuing) {
    const p0 = d0.rev / d0.qty;
    const p1 = d1.rev / d1.qty;
    const s0 = QC0 > 0 ? d0.qty / QC0 : 0;
    const s1 = QC1 > 0 ? d1.qty / QC1 : 0;

    // Công thức bóc tách chuẩn McKinsey / Big4 FP&A (tổng đại số khớp tuyệt đối RC1 - RC0)
    const itemVol = QC0 > 0 ? ((QC1 - QC0) / QC0) * d0.rev : 0;
    const itemMix = QC1 * (s1 - s0) * (p0 - avgP0);
    const itemPrice = d1.qty * (p1 - p0);

    volumeEffect += itemVol;
    mixEffect += itemMix;
    priceEffect += itemPrice;

    itemBreakdown.push({
      item: key,
      status: 'continuing',
      q0: Math.round(d0.qty * 100) / 100,
      q1: Math.round(d1.qty * 100) / 100,
      p0: Math.round(p0 * 100) / 100,
      p1: Math.round(p1 * 100) / 100,
      share0Pct: Math.round(s0 * 1000) / 10,
      share1Pct: Math.round(s1 * 1000) / 10,
      r0: Math.round(d0.rev * 100) / 100,
      r1: Math.round(d1.rev * 100) / 100,
      deltaTotal: Math.round((d1.rev - d0.rev) * 100) / 100,
      deltaPct: d0.rev !== 0 ? ((d1.rev - d0.rev) / Math.abs(d0.rev)) * 100 : null,
      volumeImpact: Math.round(itemVol * 100) / 100,
      mixImpact: Math.round(itemMix * 100) / 100,
      priceImpact: Math.round(itemPrice * 100) / 100,
      portfolioImpact: 0
    });
  }

  for (const { key, d1 } of newItems) {
    newEffect += d1.rev;
    itemBreakdown.push({
      item: key,
      status: 'new',
      q0: 0,
      q1: Math.round(d1.qty * 100) / 100,
      p0: 0,
      p1: Math.round((d1.rev / d1.qty) * 100) / 100,
      share0Pct: 0,
      share1Pct: 0,
      r0: 0,
      r1: Math.round(d1.rev * 100) / 100,
      deltaTotal: Math.round(d1.rev * 100) / 100,
      deltaPct: null,
      volumeImpact: 0,
      mixImpact: 0,
      priceImpact: 0,
      portfolioImpact: Math.round(d1.rev * 100) / 100
    });
  }

  for (const { key, d0 } of discontinuedItems) {
    discontinuedEffect -= d0.rev;
    itemBreakdown.push({
      item: key,
      status: 'discontinued',
      q0: Math.round(d0.qty * 100) / 100,
      q1: 0,
      p0: Math.round((d0.rev / d0.qty) * 100) / 100,
      p1: 0,
      share0Pct: 0,
      share1Pct: 0,
      r0: Math.round(d0.rev * 100) / 100,
      r1: 0,
      deltaTotal: Math.round(-d0.rev * 100) / 100,
      deltaPct: -100,
      volumeImpact: 0,
      mixImpact: 0,
      priceImpact: 0,
      portfolioImpact: Math.round(-d0.rev * 100) / 100
    });
  }

  itemBreakdown.sort((a, b) => a.deltaTotal - b.deltaTotal);

  const netDelta = totalR1 - totalR0;
  const netDeltaPct = totalR0 !== 0 ? (netDelta / Math.abs(totalR0)) * 100 : 0;

  // Xây dựng các bước biểu đồ thác nước (Waterfall Steps)
  const rawSteps = [
    {
      id: 'volume',
      label: hasExplicitQty ? 'Sản lượng (Volume)' : 'Số giao dịch (Volume)',
      delta: volumeEffect,
      pctOfP0: totalR0 !== 0 ? (volumeEffect / Math.abs(totalR0)) * 100 : 0,
      description: hasExplicitQty
        ? 'Thay đổi quy mô tổng sản lượng bán/vận hành giữ nguyên giá & tỷ trọng gốc'
        : 'Thay đổi tổng số lượng đơn/giao dịch phát sinh giữa 2 kỳ'
    },
    {
      id: 'mix',
      label: 'Cơ cấu mặt hàng (Mix)',
      delta: mixEffect,
      pctOfP0: totalR0 !== 0 ? (mixEffect / Math.abs(totalR0)) * 100 : 0,
      description: 'Dịch chuyển tỷ trọng giữa nhóm mặt hàng đơn giá cao và đơn giá thấp'
    },
    {
      id: 'price',
      label: hasExplicitQty ? 'Đơn giá / Giá bán (Price)' : 'Quy mô đơn TB (Ticket/Rate)',
      delta: priceEffect,
      pctOfP0: totalR0 !== 0 ? (priceEffect / Math.abs(totalR0)) * 100 : 0,
      description: hasExplicitQty
        ? 'Biến động đơn giá thực thu trên từng mặt hàng (khuyến mãi, chiết khấu, điều chỉnh giá)'
        : 'Biến động giá trị trung bình trên mỗi giao dịch trong từng nhóm'
    }
  ];

  if (Math.abs(newEffect) > 0.001) {
    rawSteps.push({
      id: 'new_items',
      label: 'Mới phát sinh (New)',
      delta: newEffect,
      pctOfP0: totalR0 !== 0 ? (newEffect / Math.abs(totalR0)) * 100 : 0,
      description: 'Đóng góp từ các mặt hàng / hạng mục mới xuất hiện ở kỳ so sánh'
    });
  }
  if (Math.abs(discontinuedEffect) > 0.001) {
    rawSteps.push({
      id: 'discontinued',
      label: 'Ngưng phát sinh (Lost)',
      delta: discontinuedEffect,
      pctOfP0: totalR0 !== 0 ? (discontinuedEffect / Math.abs(totalR0)) * 100 : 0,
      description: 'Sụt giảm do các mặt hàng / hạng mục có ở kỳ gốc nhưng không phát sinh ở kỳ này'
    });
  }

  let running = totalR0;
  const waterfall = [
    {
      id: 'p0',
      kind: 'total',
      label: 'Kỳ gốc (P0)',
      start: 0,
      end: totalR0,
      value: totalR0,
      pctOfP0: 100
    }
  ];

  for (const st of rawSteps) {
    // Nếu không có cột mặt hàng riêng thì Mix = 0, vẫn hiển thị 3 trụ chính Volume / Mix / Price cho chuẩn PVM
    const nextVal = running + st.delta;
    waterfall.push({
      id: st.id,
      kind: 'step',
      label: st.label,
      description: st.description,
      start: running,
      end: nextVal,
      value: st.delta,
      pctOfP0: st.pctOfP0
    });
    running = nextVal;
  }

  waterfall.push({
    id: 'p1',
    kind: 'total',
    label: 'Kỳ so sánh (P1)',
    start: 0,
    end: totalR1,
    value: totalR1,
    pctOfP0: totalR0 !== 0 ? (totalR1 / Math.abs(totalR0)) * 100 : 100
  });

  return {
    mode: hasExplicitQty ? 'FULL_PVM' : 'COUNT_RATE_MIX',
    effectiveGroupCol,
    totalR0,
    totalR1,
    netDelta,
    netDeltaPct,
    totalQ0: totalQ0All,
    totalQ1: totalQ1All,
    avgPrice0: totalQ0All > 0 ? totalR0 / totalQ0All : 0,
    avgPrice1: totalQ1All > 0 ? totalR1 / totalQ1All : 0,
    components: {
      volume: volumeEffect,
      mix: mixEffect,
      price: priceEffect,
      newItems: newEffect,
      discontinued: discontinuedEffect,
      residualCheck: totalR1 - (totalR0 + volumeEffect + mixEffect + priceEffect + newEffect + discontinuedEffect)
    },
    steps: rawSteps,
    waterfall,
    itemBreakdown
  };
}

// ============================================================================
// 3. QUÉT KEY-DRIVERS ĐA CHIỀU & GIẢI PHẪU CHÉO 2 TẦNG (CROSS-DRILL)
// ============================================================================

function computeDimensionDrivers(rowsP0, rowsP1, dimCol, metricCol, totalR0, netDelta) {
  const sum0 = new Map();
  const sum1 = new Map();
  const cnt0 = new Map();
  const cnt1 = new Map();

  for (const r of rowsP0) {
    const val = Number(r?.[metricCol]);
    if (!Number.isFinite(val)) continue;
    const key = String(r?.[dimCol] || 'Chưa phân loại').trim() || 'Chưa phân loại';
    sum0.set(key, (sum0.get(key) || 0) + val);
    cnt0.set(key, (cnt0.get(key) || 0) + 1);
  }
  for (const r of rowsP1) {
    const val = Number(r?.[metricCol]);
    if (!Number.isFinite(val)) continue;
    const key = String(r?.[dimCol] || 'Chưa phân loại').trim() || 'Chưa phân loại';
    sum1.set(key, (sum1.get(key) || 0) + val);
    cnt1.set(key, (cnt1.get(key) || 0) + 1);
  }

  const allKeys = new Set([...sum0.keys(), ...sum1.keys()]);
  const members = [];

  for (const key of allKeys) {
    const v0 = sum0.get(key) || 0;
    const v1 = sum1.get(key) || 0;
    const delta = v1 - v0;
    const deltaPct = v0 !== 0 ? (delta / Math.abs(v0)) * 100 : v1 > 0 ? 100 : 0;
    const impactOnTotalPct = totalR0 !== 0 ? (delta / Math.abs(totalR0)) * 100 : 0;
    const shareOfNetMovePct = Math.abs(netDelta) > 0.0001 ? (delta / Math.abs(netDelta)) * 100 : 0;

    members.push({
      name: key,
      p0: Math.round(v0 * 100) / 100,
      p1: Math.round(v1 * 100) / 100,
      tx0: cnt0.get(key) || 0,
      tx1: cnt1.get(key) || 0,
      delta: Math.round(delta * 100) / 100,
      deltaPct: Math.round(deltaPct * 10) / 10,
      impactOnTotalPct: Math.round(impactOnTotalPct * 100) / 100,
      shareOfNetMovePct: Math.round(shareOfNetMovePct * 10) / 10
    });
  }

  const sortedAsc = members.slice().sort((a, b) => a.delta - b.delta);
  const sortedDesc = members.slice().sort((a, b) => b.delta - a.delta);

  return {
    dimension: dimCol,
    memberCount: members.length,
    topNegative: sortedAsc.filter((m) => m.delta < 0).slice(0, 3),
    topPositive: sortedDesc.filter((m) => m.delta > 0).slice(0, 3),
    allMembers: sortedAsc
  };
}

function computeSegmentDeepDive(rowsP0, rowsP1, mapping, filterDim, filterVal) {
  const sub0 = rowsP0.filter((r) => String(r?.[filterDim] || 'Chưa phân loại').trim() === filterVal);
  const sub1 = rowsP1.filter((r) => String(r?.[filterDim] || 'Chưa phân loại').trim() === filterVal);

  const subPvm = computeExactPvmBridge(sub0, sub1, {
    metricCol: mapping.metricCol,
    quantityCol: mapping.quantityCol,
    itemCol: mapping.itemCol
  });

  const otherDims = (mapping.dimensionCols || []).filter((d) => d !== filterDim);
  const subDrivers = otherDims.map((d) =>
    computeDimensionDrivers(sub0, sub1, d, mapping.metricCol, subPvm.totalR0, subPvm.netDelta)
  );

  return {
    filterDim,
    filterVal,
    rowCountP0: sub0.length,
    rowCountP1: sub1.length,
    pvm: subPvm,
    subDrivers
  };
}

// ============================================================================
// 4. ĐO LƯỜNG "BẪY BÁO CÁO BẨN" (DIRTY VS CLEAN DISTORTION AUDIT)
// ============================================================================

function computeDirtyTrapComparison(rawRows, mapping, period0, period1, granularity, cleanBridge) {
  const dateCol = mapping.dateCol;
  const metricCol = mapping.metricCol;
  if (!metricCol) return null;

  // Giả lập cách Excel PivotTable đọc trực tiếp file bẩn chưa qua làm sạch:
  // - Ngày dạng YYYY-MM-DD hoặc DD/MM/YYYY đọc thô (các dòng MM/DD/YYYY như 10/04/2026 bị hiểu nhầm sang tháng 04!)
  // - Không lọc dòng trùng lặp ngầm
  // - Không khử ngoại lai gõ dư số 0
  // - Không gom tên viết sai ("TP.HCM" vs "Sài Gòn")
  const naiveParsePeriod = (rawDate) => {
    const s = cleanInvisibleAndNFC(rawDate);
    if (!s) return '';
    let m = s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
    if (m) {
      return isoDateToPeriod(
        `${m[1]}-${String(Number(m[2])).padStart(2, '0')}-${String(Number(m[3])).padStart(2, '0')}`,
        granularity
      );
    }
    m = s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})/);
    if (m) {
      const d = Number(m[1]);
      const mo = Number(m[2]);
      let y = Number(m[3]);
      if (y < 100) y += 2000;
      if (mo >= 1 && mo <= 12 && d >= 1 && d <= 31) {
        return isoDateToPeriod(`${y}-${String(mo).padStart(2, '0')}-${String(d).padStart(2, '0')}`, granularity);
      }
    }
    return '';
  };

  let rawR0 = 0;
  let rawR1 = 0;
  for (const r of rawRows) {
    const p = dateCol ? naiveParsePeriod(r?.[dateCol]) : '';
    const num = parseVietnameseNumber(r?.[metricCol]).num;
    if (num === null) continue;
    if (p === period0) rawR0 += num;
    else if (p === period1) rawR1 += num;
  }

  const rawDelta = rawR1 - rawR0;
  const rawDeltaPct = rawR0 !== 0 ? (rawDelta / Math.abs(rawR0)) * 100 : 0;

  const p0Distortion = rawR0 - cleanBridge.totalR0;
  const p1Distortion = rawR1 - cleanBridge.totalR1;
  const deltaDistortion = rawDelta - cleanBridge.netDelta;
  const pctPointError = rawDeltaPct - cleanBridge.netDeltaPct;

  // Đếm mức độ phân mảnh danh mục trên chiều đầu tiên (VD: Chi nhánh bị tách làm bao nhiêu tên rác)
  const primaryDim = (mapping.dimensionCols || []).find((d) => d !== mapping.itemCol) || mapping.itemCol;
  const rawDistinctDim = primaryDim
    ? new Set(rawRows.map((r) => String(r?.[primaryDim] ?? '').trim()).filter(Boolean)).size
    : 0;

  return {
    rawP0: Math.round(rawR0 * 100) / 100,
    rawP1: Math.round(rawR1 * 100) / 100,
    rawDelta: Math.round(rawDelta * 100) / 100,
    rawDeltaPct: Math.round(rawDeltaPct * 10) / 10,
    p0Distortion: Math.round(p0Distortion * 100) / 100,
    p1Distortion: Math.round(p1Distortion * 100) / 100,
    deltaDistortion: Math.round(deltaDistortion * 100) / 100,
    pctPointError: Math.round(pctPointError * 10) / 10,
    primaryDim,
    rawDistinctDim
  };
}

// ============================================================================
// 5. SINH BÁO CÁO TỰ ĐỘNG "TRẢ LỜI SẾP TRONG 30 GIÂY" (EXECUTIVE MEMO)
// ============================================================================

function buildExecutiveNarrative({
  datasetTitle,
  period0,
  period1,
  mapping,
  polarity,
  pvm,
  dimensionDrivers,
  primaryDeepDive,
  dirtyTrap,
  cleanStats
}) {
  const metricName = mapping.metricCol || 'Chỉ số chính';
  const isUp = pvm.netDelta >= 0;
  const verb = isUp ? 'TĂNG' : 'GIẢM';
  const isFavorable = polarity === 'lower_is_better' ? !isUp : isUp;
  const statusIcon = isFavorable ? '🟢' : '🔴';

  // Xác định động lực PVM lớn nhất theo chiều biến động chính (hoặc biên độ tuyệt đối lớn nhất)
  const sortedSteps = pvm.steps
    .slice()
    .sort((a, b) => (isUp ? b.delta - a.delta : a.delta - b.delta));
  const mainPvmDriver = sortedSteps[0];

  const lines = [];
  lines.push(
    `📌 BÁO CÁO GIẢI PHẪU BIẾN ĐỘNG: ${(datasetTitle || metricName).toUpperCase()} (${period1} so với ${period0})`
  );
  lines.push(
    `${statusIcon} Tổng quan ${metricName}: Kỳ ${period1} đạt ${formatVnCompact(pvm.totalR1)} (so với ${period0}: ${formatVnCompact(
      pvm.totalR0
    )}) ➔ ${verb} ${formatVnCompact(Math.abs(pvm.netDelta))} (${formatPct(pvm.netDeltaPct)}).`
  );

  // Giải phẫu PVM Bridge
  const pvmParts = pvm.steps.map((s) => `${s.label}: ${formatSignedCompact(s.delta)} (${formatPct(s.pctOfP0)})`);
  lines.push(`🔬 Giải phẫu nguyên nhân gốc (PVM Bridge): ${pvmParts.join(' | ')}.`);
  if (mainPvmDriver) {
    lines.push(
      `   ➔ Trọng tâm lớn nhất đến từ [${mainPvmDriver.label}] làm biến động ${formatSignedCompact(
        mainPvmDriver.delta
      )} (${formatPct(mainPvmDriver.pctOfP0)} so với nền kỳ gốc).`
    );
  }

  // Điểm danh Top 3 Thủ phạm kéo xuống / kéo lên theo từng chiều quản trị
  const targetDirectionDrivers = dimensionDrivers.filter((d) =>
    isUp ? d.topPositive.length > 0 : d.topNegative.length > 0
  );

  for (const dimBlock of targetDirectionDrivers.slice(0, 3)) {
    const list = isUp ? dimBlock.topPositive : dimBlock.topNegative;
    const formattedTop = list
      .map((m, idx) => `#${idx + 1} ${m.name} (${formatSignedCompact(m.delta)} | ${formatPct(m.deltaPct)})`)
      .join('; ');
    lines.push(
      `🎯 Top ${isUp ? 'đóng góp tăng' : 'kéo giảm'} theo [${dimBlock.dimension}]: ${formattedTop}.`
    );
  }

  // Nếu có giải phẫu chéo tầng 2 cho điểm nóng số 1
  if (primaryDeepDive && primaryDeepDive.subDrivers?.length > 0) {
    const subSnippets = [];
    for (const sd of primaryDeepDive.subDrivers.slice(0, 2)) {
      const culprit = (isUp ? sd.topPositive : sd.topNegative)[0];
      if (culprit) {
        subSnippets.push(`${sd.dimension}: ${culprit.name} (${formatSignedCompact(culprit.delta)})`);
      }
    }
    if (subSnippets.length > 0) {
      lines.push(
        `🕵️ Đào sâu điểm nóng số 1 [${primaryDeepDive.filterDim} = ${primaryDeepDive.filterVal}]: Nguyên nhân trực tiếp nằm ở ${subSnippets.join(
          ' kết hợp '
        )}.`
      );
    }
  }

  // Cảnh báo giá trị của việc làm sạch dữ liệu trước khi lên báo cáo
  if (cleanStats && cleanStats.totalIssuesFound > 0) {
    const distortionNote =
      dirtyTrap && (Math.abs(dirtyTrap.p1Distortion) > 1 || Math.abs(dirtyTrap.pctPointError) >= 0.3)
        ? ` — ngăn chặn sai lệch ${formatSignedCompact(dirtyTrap.p1Distortion)} ở kỳ ${period1} (nếu chạy Pivot trên file bẩn gốc sẽ báo nhầm biến động là ${formatPct(
            dirtyTrap.rawDeltaPct
          )} thay vì thực tế ${formatPct(pvm.netDeltaPct)}!)`
        : '';
    lines.push(
      `🛡️ Kiểm toán độ sạch dữ liệu: Đã tự động xử lý ${cleanStats.totalIssuesFound} lỗi dữ liệu bẩn (Điểm sức khỏe dữ liệu tăng từ ${cleanStats.rawHealthScore}/100 ➔ ${cleanStats.cleanHealthScore}/100)${distortionNote}.`
    );
  }

  return lines.join('\n');
}

// ============================================================================
// 6. HÀM ĐIỀU PHỐI PHÂN TÍCH BIẾN ĐỘNG TỔNG THỂ (RUN VARIANCE ANALYSIS)
// ============================================================================

function runVarianceAnalysis({
  datasetTitle = 'Báo cáo Phân tích',
  rawRows = [],
  cleanRows = [],
  cleanStats = null,
  mapping = {},
  period0 = null,
  period1 = null,
  granularity = 'month',
  polarity = null,
  drillDim = null,
  drillVal = null
}) {
  const metricCol = mapping.metricCol;
  const dateCol = mapping.dateCol;
  const effectivePolarity = polarity || mapping.polarity || 'higher_is_better';

  const availablePeriods = extractAvailablePeriods(cleanRows, dateCol, granularity);
  if (!metricCol || availablePeriods.length === 0) {
    return {
      ready: false,
      reason: !metricCol
        ? 'Chưa xác định được cột Chỉ số số liệu (Doanh thu / Chi phí / Giá trị) để tính biến động.'
        : 'Chưa có ít nhất 1 kỳ thời gian hợp lệ từ cột Ngày tháng để đối chiếu.',
      availablePeriods,
      granularity
    };
  }

  let p1 = period1 && availablePeriods.some((x) => x.period === period1) ? period1 : availablePeriods[availablePeriods.length - 1].period;
  let p0 =
    period0 && availablePeriods.some((x) => x.period === period0)
      ? period0
      : availablePeriods.length >= 2
      ? availablePeriods[availablePeriods.length - 2].period
      : availablePeriods[0].period;

  if (p0 === p1 && availablePeriods.length >= 2) {
    p0 = availablePeriods.find((x) => x.period !== p1)?.period || p0;
  }

  const rowsP0 = [];
  const rowsP1 = [];
  for (const r of cleanRows) {
    const p = isoDateToPeriod(dateCol ? r?.[dateCol] : r?._period, granularity);
    if (p === p0) rowsP0.push(r);
    else if (p === p1) rowsP1.push(r);
  }

  const pvm = computeExactPvmBridge(rowsP0, rowsP1, {
    metricCol,
    quantityCol: mapping.quantityCol,
    itemCol: mapping.itemCol
  });

  // Quét toàn bộ các chiều danh mục (ưu tiên các chiều Chi nhánh, Nhân viên, Kênh bán lên đầu trước Mặt hàng)
  const rawDims = Array.isArray(mapping.dimensionCols) ? mapping.dimensionCols : [];
  const orderedDims = [
    ...rawDims.filter((d) => d !== mapping.itemCol),
    ...(mapping.itemCol && rawDims.includes(mapping.itemCol) ? [mapping.itemCol] : [])
  ];

  const dimensionDrivers = orderedDims.map((dim) =>
    computeDimensionDrivers(rowsP0, rowsP1, dim, metricCol, pvm.totalR0, pvm.netDelta)
  );

  // Chọn phân khúc mặc định để giải phẫu chéo tầng 2 (hoặc theo lựa chọn click của người dùng)
  let targetDrillDim = drillDim;
  let targetDrillVal = drillVal;

  if (!targetDrillDim || !targetDrillVal) {
    const firstNonItemBlock = dimensionDrivers[0];
    if (firstNonItemBlock) {
      const candidate =
        pvm.netDelta <= 0
          ? firstNonItemBlock.topNegative[0] || firstNonItemBlock.topPositive[0]
          : firstNonItemBlock.topPositive[0] || firstNonItemBlock.topNegative[0];
      if (candidate) {
        targetDrillDim = firstNonItemBlock.dimension;
        targetDrillVal = candidate.name;
      }
    }
  }

  const deepDive =
    targetDrillDim && targetDrillVal
      ? computeSegmentDeepDive(rowsP0, rowsP1, mapping, targetDrillDim, targetDrillVal)
      : null;

  const dirtyTrap = computeDirtyTrapComparison(rawRows, mapping, p0, p1, granularity, pvm);

  const executiveMemo = buildExecutiveNarrative({
    datasetTitle,
    period0: p0,
    period1: p1,
    mapping,
    polarity: effectivePolarity,
    pvm,
    dimensionDrivers,
    primaryDeepDive: deepDive,
    dirtyTrap,
    cleanStats
  });

  // Xu hướng qua tất cả các kỳ có trong dữ liệu (phục vụ mini trendline)
  const periodTrend = availablePeriods.map(({ period }) => {
    let sum = 0;
    let qty = 0;
    let count = 0;
    for (const r of cleanRows) {
      const p = isoDateToPeriod(dateCol ? r?.[dateCol] : r?._period, granularity);
      if (p === period) {
        const v = Number(r?.[metricCol]);
        if (Number.isFinite(v)) {
          sum += v;
          count++;
        }
        if (mapping.quantityCol && Number.isFinite(Number(r?.[mapping.quantityCol]))) {
          qty += Number(r[mapping.quantityCol]);
        }
      }
    }
    return {
      period,
      total: Math.round(sum * 100) / 100,
      quantity: Math.round(qty * 100) / 100,
      txCount: count
    };
  });

  return {
    ready: true,
    granularity,
    polarity: effectivePolarity,
    availablePeriods,
    period0: p0,
    period1: p1,
    rowCountP0: rowsP0.length,
    rowCountP1: rowsP1.length,
    periodTrend,
    pvm,
    dimensionDrivers,
    deepDive,
    dirtyTrap,
    executiveMemo
  };
}

module.exports = {
  isoDateToPeriod,
  extractAvailablePeriods,
  formatVnCompact,
  formatSignedCompact,
  formatPct,
  computeExactPvmBridge,
  computeDimensionDrivers,
  computeSegmentDeepDive,
  runVarianceAnalysis
};
