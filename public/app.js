/**
 * 🧼📉 DataScrub & Variance Studio — Frontend Interactive Application (app.js)
 */

const studio = {
  datasets: [],
  recipes: [],
  state: null,
  diffFilter: 'all',
  diffSearch: '',
  showInlineDiff: true,
  lastSqlExecuted: ''
};

// ============================================================================
// 1. TIỆN ÍCH CHUNG & ĐỊNH DẠNG SỐ LIỆU TIẾNG VIỆT
// ============================================================================

function escapeHtml(val) {
  return String(val ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function fmtCompact(val) {
  const n = Number(val);
  if (!Number.isFinite(n)) return '--';
  const sign = n < 0 ? '-' : '';
  const abs = Math.abs(n);
  if (abs >= 1_000_000_000) {
    return `${sign}${(abs / 1_000_000_000).toLocaleString('vi-VN', { maximumFractionDigits: 2 })} tỷ`;
  }
  if (abs >= 1_000_000) {
    return `${sign}${(abs / 1_000_000).toLocaleString('vi-VN', { maximumFractionDigits: 2 })} tr`;
  }
  if (abs >= 1000) {
    return `${sign}${Math.round(abs).toLocaleString('vi-VN')}`;
  }
  return `${sign}${abs.toLocaleString('vi-VN', { maximumFractionDigits: 2 })}`;
}

function fmtSigned(val) {
  const n = Number(val);
  if (!Number.isFinite(n) || n === 0) return '0';
  return n > 0 ? `+${fmtCompact(n)}` : fmtCompact(n);
}

function fmtFullNum(val) {
  const n = Number(val);
  if (!Number.isFinite(n)) return escapeHtml(val ?? '');
  if (Number.isInteger(n) || Math.abs(n) >= 100) {
    return Math.round(n).toLocaleString('vi-VN');
  }
  return n.toLocaleString('vi-VN', { maximumFractionDigits: 2 });
}

function fmtPct(val) {
  const n = Number(val);
  if (val === null || val === undefined || !Number.isFinite(n)) return 'N/A';
  return `${n > 0 ? '+' : ''}${n.toFixed(1)}%`;
}

function isFavorableDelta(delta, polarity = 'higher_is_better') {
  if (delta === 0) return true;
  return polarity === 'lower_is_better' ? delta < 0 : delta > 0;
}

function showToast(msg) {
  const wrap = document.getElementById('toastContainer');
  const el = document.createElement('div');
  el.className = 'toast';
  el.textContent = msg;
  wrap.appendChild(el);
  setTimeout(() => {
    el.remove();
  }, 3800);
}

async function apiJson(url, options = {}) {
  const res = await fetch(url, {
    headers: { 'Content-Type': 'application/json' },
    ...options
  });
  const data = await res.json();
  if (!res.ok) {
    throw new Error(data.error || `Lỗi HTTP ${res.status}`);
  }
  return data;
}

// ============================================================================
// 2. KHỞI TẠO & CHUYỂN TAB / GIAO DIỆN
// ============================================================================

function switchTab(tabName) {
  document.querySelectorAll('.tab-btn').forEach((b) => {
    b.classList.toggle('active', b.dataset.tab === tabName);
  });
  document.querySelectorAll('.tab-panel').forEach((p) => {
    p.classList.toggle('active', p.id === `tab-${tabName}`);
  });
  if (tabName === 'sql' && !studio.lastSqlExecuted && studio.state?.sqlMeta?.presets?.[0]) {
    const firstSql = studio.state.sqlMeta.presets[0].sql;
    document.getElementById('sqlEditor').value = firstSql;
    runSqlQuery(firstSql);
  }
}

function populateDatasetSelect() {
  const sel = document.getElementById('datasetSelect');
  const opts = studio.datasets.map(
    (d) => `<option value="${escapeHtml(d.id)}">${escapeHtml(d.name)} (${d.rowCount} dòng)</option>`
  );
  if (studio.state?.datasetId === 'custom_upload') {
    opts.unshift(
      `<option value="custom_upload">📂 [Tệp tải lên] ${escapeHtml(studio.state.fileName)} (${studio.state.stats.rawRowCount} dòng)</option>`
    );
  }
  sel.innerHTML = opts.join('');
  sel.value = studio.state?.datasetId || 'retail_fmcg';
}

function populateRecipeDropdowns() {
  const optionsHtml =
    `<option value="">✨ Tự động nhận diện thông minh (Auto-Smart)</option>` +
    studio.recipes
      .map((r) => `<option value="${escapeHtml(r.id)}">${escapeHtml(r.name)}</option>`)
      .join('');

  document.getElementById('modalRecipeSelect').innerHTML = optionsHtml;
  document.getElementById('selBatchRecipe').innerHTML = studio.recipes
    .map((r) => `<option value="${escapeHtml(r.id)}">${escapeHtml(r.name)}</option>`)
    .join('');

  if (studio.state?.activeRecipeId) {
    document.getElementById('selBatchRecipe').value = studio.state.activeRecipeId;
  }
  document.getElementById('badgeRecipeCount').textContent = String(studio.recipes.length);
}

// ============================================================================
// 3. RENDER 4-CARD KPI RIBBON & CONTEXT BAR
// ============================================================================

function renderTopContextAndKpis() {
  const st = studio.state;
  if (!st) return;

  document.getElementById('activeIndustryTag').textContent = st.industry || 'Đa ngành';
  document.getElementById('activeDatasetTitle').textContent = `${st.datasetTitle} • ${st.stats.rawRowCount} dòng gốc ➔ ${st.stats.cleanRowCount} dòng sạch`;

  const dlSample = document.getElementById('btnDownloadDirtySample');
  if (st.datasetId !== 'custom_upload') {
    dlSample.href = `/api/datasets/download-raw-xlsx?datasetId=${encodeURIComponent(st.datasetId)}`;
    dlSample.classList.remove('hidden');
  } else {
    dlSample.classList.add('hidden');
  }

  // Sheet selector nếu file XLSX nhiều Sheet
  const sheetWrap = document.getElementById('sheetSelectorWrap');
  const sheetSel = document.getElementById('sheetSelect');
  if (st.availableSheets && st.availableSheets.length > 1) {
    sheetWrap.classList.remove('hidden');
    sheetSel.innerHTML = st.availableSheets
      .map((s) => `<option value="${escapeHtml(s)}">${escapeHtml(s)}</option>`)
      .join('');
    sheetSel.value = st.sheetName;
  } else {
    sheetWrap.classList.add('hidden');
  }

  const rcpBadge = document.getElementById('activeRecipeBadge');
  const activeRcp = studio.recipes.find((r) => r.id === st.activeRecipeId);
  if (activeRcp) {
    rcpBadge.textContent = `⚡ Đang áp dụng: ${activeRcp.name}`;
    rcpBadge.classList.remove('hidden');
  } else {
    rcpBadge.classList.add('hidden');
  }

  // KPI 1: Data Health Score
  const { rawHealthScore, cleanHealthScore, totalIssuesFound, droppedRowCount } = st.stats;
  document.getElementById('kpiHealthScore').innerHTML = `
    <span style="color:var(--text-muted);font-size:15px;text-decoration:line-through">${rawHealthScore}</span>
    <span>➔</span>
    <span style="color:var(--good)">${cleanHealthScore}/100</span>
  `;
  document.getElementById('kpiHealthBar').style.width = `${cleanHealthScore}%`;
  document.getElementById('kpiHealthSub').textContent = `Đã tự động sửa ${totalIssuesFound} lỗi bẩn • Loại ${droppedRowCount} dòng trùng/lỗi`;
  document.getElementById('badgeIssuesCount').textContent = `${totalIssuesFound} lỗi`;

  // KPI 2 & 3 & 4: Variance + Dirty Trap
  const vr = st.variance;
  if (!vr || !vr.ready) {
    document.getElementById('kpiDirtyTrap').textContent = 'Chưa đủ kỳ so sánh';
    document.getElementById('kpiDirtyTrapSub').textContent = vr?.reason || '';
    document.getElementById('kpiVarianceDelta').textContent = '--';
    document.getElementById('kpiVarianceSub').textContent = 'Hãy kiểm tra ánh xạ cột Ngày & Chỉ số';
    document.getElementById('kpiTopDriver').textContent = '--';
    document.getElementById('kpiTopDriverSub').textContent = '';
    document.getElementById('badgeVariancePct').textContent = '--';
    return;
  }

  const pvm = vr.pvm;
  const polarity = vr.polarity;
  const fav = isFavorableDelta(pvm.netDelta, polarity);

  // KPI 2: Dirty Trap
  const dt = vr.dirtyTrap;
  if (dt && (Math.abs(dt.p1Distortion) > 1 || Math.abs(dt.pctPointError) >= 0.2)) {
    document.getElementById('kpiDirtyTrap').innerHTML = `
      <span>Lệch ${fmtSigned(dt.p1Distortion)}</span>
      <span class="delta-pill warn">Báo nhầm ${fmtPct(dt.rawDeltaPct)}</span>
    `;
    document.getElementById('kpiDirtyTrapSub').textContent = `Trên file bẩn kỳ ${vr.period1} bị đội/lệch số do ngoại lai & dòng trùng!`;
  } else {
    document.getElementById('kpiDirtyTrap').innerHTML = `<span style="color:var(--good)">0 sai lệch trọng yếu</span>`;
    document.getElementById('kpiDirtyTrapSub').textContent = `Phân mảnh danh mục giảm mạnh nhờ chuẩn hóa tên`;
  }

  // KPI 3: Clean Variance
  document.getElementById('kpiVarianceLabel').textContent = `📉 ${st.mapping.metricCol || 'Chỉ số'} (${vr.period0} ➔ ${vr.period1})`;
  document.getElementById('kpiVarianceDelta').innerHTML = `
    <span>${fmtCompact(pvm.totalR1)}</span>
    <span class="delta-pill ${fav ? 'good' : 'bad'}">${fmtPct(pvm.netDeltaPct)} (${fmtSigned(pvm.netDelta)})</span>
  `;
  document.getElementById('kpiVarianceSub').textContent = `Kỳ gốc ${vr.period0}: ${fmtCompact(pvm.totalR0)} • Kỳ này ${vr.period1}: ${fmtCompact(pvm.totalR1)}`;
  document.getElementById('badgeVariancePct').textContent = fmtPct(pvm.netDeltaPct);

  // KPI 4: Primary Driver
  const isUp = pvm.netDelta >= 0;
  const topStep = pvm.steps.slice().sort((a, b) => (isUp ? b.delta - a.delta : a.delta - b.delta))[0];
  const firstDim = vr.dimensionDrivers[0];
  const topEntity = firstDim ? (isUp ? firstDim.topPositive[0] : firstDim.topNegative[0]) : null;

  document.getElementById('kpiTopDriver').textContent = topStep
    ? `${topStep.label}: ${fmtSigned(topStep.delta)} (${fmtPct(topStep.pctOfP0)})`
    : '--';
  document.getElementById('kpiTopDriverSub').textContent = topEntity
    ? `Điểm nóng #1 [${firstDim.dimension}]: ${topEntity.name} (${fmtSigned(topEntity.delta)})`
    : 'Xem chi tiết cầu nối PVM ở Tab 2';
}

// ============================================================================
// 4. RENDER TAB 1: TRẠM LỌC DỮ LIỆU BẨN VN
// ============================================================================

function syncCleanerControlsFromState() {
  const st = studio.state;
  const cfg = st.config;
  document.getElementById('chkDates').checked = Boolean(cfg.standardizeDates);
  document.getElementById('chkPhones').checked = Boolean(cfg.standardizePhones);
  document.getElementById('chkNumbers').checked = Boolean(cfg.standardizeNumbers);
  document.getElementById('chkFuzzy').checked = Boolean(cfg.enableFuzzyMatching);
  document.getElementById('rngFuzzy').value = cfg.fuzzyThreshold || 0.84;
  document.getElementById('lblFuzzyVal').textContent = `${Math.round((cfg.fuzzyThreshold || 0.84) * 100)}%`;
  document.getElementById('selDedupe').value = cfg.dedupeStrategy || 'remove_hidden';
  document.getElementById('selOutlier').value = cfg.outlierStrategy || 'median_impute';
  document.getElementById('selDatePref').value = cfg.datePreference || 'AUTO_VN';

  // Sync Column Role Mapping selects
  const fillColSelect = (id, selectedVal, allowNone = true) => {
    const el = document.getElementById(id);
    const opts = [];
    if (allowNone) opts.push('<option value="">-- Không dùng --</option>');
    for (const c of st.columns) {
      opts.push(`<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`);
    }
    el.innerHTML = opts.join('');
    el.value = selectedVal && st.columns.includes(selectedVal) ? selectedVal : '';
  };

  fillColSelect('mapDateCol', st.mapping.dateCol);
  fillColSelect('mapMetricCol', st.mapping.metricCol, false);
  fillColSelect('mapQtyCol', st.mapping.quantityCol);
  fillColSelect('mapItemCol', st.mapping.itemCol);
  fillColSelect('mapPhoneCol', st.mapping.phoneCol);
  fillColSelect('mapIdCol', st.mapping.idCol);
}

function renderCleanerTab() {
  const st = studio.state;
  syncCleanerControlsFromState();

  document.getElementById('cleanerTimingLabel').textContent = `Xử lý hoàn tất trong ${st.stats.elapsedMs} ms`;

  // 1. Column Health Cards
  const colGrid = document.getElementById('columnHealthGrid');
  colGrid.innerHTML = st.columnHealth
    .map((c) => {
      const reducedFrag =
        c.detectedType === 'categorical' && c.distinctBefore > c.distinctAfter
          ? `<div style="color:var(--good);font-size:11.5px;font-weight:600">Gom tên: ${c.distinctBefore} biến thể ➔ ${c.distinctAfter} chuẩn</div>`
          : `<div class="muted" style="font-size:11.5px">${c.distinctAfter} giá trị phân biệt</div>`;
      return `
        <div class="col-card">
          <div class="col-card-top">
            <span class="col-card-name" title="${escapeHtml(c.name)}">${escapeHtml(c.name)}</span>
            <span class="type-badge">${escapeHtml(c.detectedType)}</span>
          </div>
          <div class="col-card-stats">
            <span>Đầy đủ: ${c.completenessPct}%</span>
            <span style="color:${c.issuesFixed > 0 ? 'var(--accent)' : 'var(--text-muted)'}">
              ${c.issuesFixed > 0 ? `✨ Sửa ${c.issuesFixed} ô` : '✓ Chuẩn'}
            </span>
          </div>
          ${reducedFrag}
        </div>
      `;
    })
    .join('');

  // 2. Fuzzy Clusters List
  const fzList = document.getElementById('fuzzyClustersList');
  document.getElementById('fuzzyClusterCount').textContent = `${st.fuzzyClusters.length} cụm tên chuẩn hóa`;

  if (st.fuzzyClusters.length === 0) {
    fzList.innerHTML = `<div class="muted">Không phát hiện cụm tên viết sai / viết tắt nào ở ngưỡng hiện tại.</div>`;
  } else {
    fzList.innerHTML = st.fuzzyClusters
      .map(
        (cl, idx) => `
        <div class="fuzzy-item">
          <div class="fuzzy-item-head">
            <span><strong>[${escapeHtml(cl.column)}]</strong> ➔ Chuẩn: <strong style="color:var(--good)">${escapeHtml(
              cl.canonical
            )}</strong></span>
            <span class="muted">Đã gom ${cl.changedRows}/${cl.totalRows} dòng</span>
          </div>
          <div class="fuzzy-variants">
            ${cl.variants
              .map(
                (v) =>
                  `<span class="variant-chip ${v.name !== cl.canonical ? 'dirty' : ''}">${escapeHtml(v.name)} (${v.count})</span>`
              )
              .join('')}
          </div>
          <div class="fuzzy-target-row">
            <input type="text" class="control-input compact" data-fuzzy-idx="${idx}" value="${escapeHtml(
              cl.canonical
            )}" title="Đặt tên chuẩn tùy chỉnh cho cụm này" />
            <button class="btn btn-secondary btn-sm btn-apply-canonical" data-fuzzy-idx="${idx}">Đổi tên chuẩn</button>
          </div>
        </div>
      `
      )
      .join('');
  }

  // 3. Outliers & Hidden Duplicates List
  const anomList = document.getElementById('outliersAndDupesList');
  const dupRows = st.stagedPreview.filter((r) => r.flags.includes('duplicate_hidden') || r.flags.includes('duplicate_exact'));
  const totalAnoms = st.outliers.length + st.stats.issueCounts.duplicate_hidden + st.stats.issueCounts.duplicate_exact;
  document.getElementById('anomalySummaryCount').textContent = `${totalAnoms} phát hiện`;

  const anomItemsHtml = [];
  for (const out of st.outliers) {
    anomItemsHtml.push(`
      <div class="anomaly-item">
        <div><strong>🚨 Ngoại lai Dòng #${out.rowNumber}</strong> • Nhóm: <strong>${escapeHtml(out.group)}</strong></div>
        <div class="muted">${escapeHtml(out.reason)}</div>
        <div style="margin-top:3px">
          Gốc: <span class="cell-diff-old">${fmtFullNum(out.originalValue)}</span>
          <span class="cell-diff-arrow">➔</span>
          <span class="cell-diff-new">${out.resolvedValue !== null ? fmtFullNum(out.resolvedValue) : 'Đã loại bỏ dòng'}</span>
        </div>
      </div>
    `);
  }
  for (const d of dupRows) {
    const isHidden = d.flags.includes('duplicate_hidden');
    anomItemsHtml.push(`
      <div class="anomaly-item dup">
        <div><strong>👯 ${isHidden ? 'Trùng lặp ngầm' : 'Trùng khớp 100%'} ở Dòng #${d.rowNumber}</strong></div>
        <div class="muted">${escapeHtml(d.dropReason || 'Phát hiện trùng giao dịch sau khi chuẩn hóa ngày, tên & SĐT')}</div>
      </div>
    `);
  }
  anomList.innerHTML =
    anomItemsHtml.length > 0
      ? anomItemsHtml.join('')
      : `<div class="muted">Không phát hiện giá trị ngoại lai hoặc dòng trùng lặp nào.</div>`;

  // 4. Before / After Diff Grid
  renderCleanerDiffTable();
}

function renderCleanerDiffTable() {
  const st = studio.state;
  if (!st) return;

  const thead = document.querySelector('#cleanerDiffTable thead');
  const tbody = document.querySelector('#cleanerDiffTable tbody');

  const numericColSet = new Set(
    st.columnHealth.filter((c) => c.detectedType === 'numeric').map((c) => c.name)
  );

  thead.innerHTML = `
    <tr>
      <th>#</th>
      <th>Trạng thái</th>
      ${st.columns.map((c) => `<th>${escapeHtml(c)}</th>`).join('')}
    </tr>
  `;

  const searchLower = studio.diffSearch.trim().toLowerCase();
  const filtered = st.stagedPreview.filter((row) => {
    const changeValues = Object.values(row.changes || {});
    if (studio.diffFilter === 'changed' && changeValues.length === 0 && row.flags.length === 0) return false;
    if (studio.diffFilter === 'date' && !changeValues.some((ch) => String(ch.type).includes('date') || String(ch.type).includes('mm_dd') || String(ch.type).includes('serial') || String(ch.type).includes('iso'))) return false;
    if (studio.diffFilter === 'fuzzy' && !changeValues.some((ch) => ch.type === 'fuzzy_standardized')) return false;
    if (studio.diffFilter === 'phone' && !changeValues.some((ch) => String(ch.type).includes('phone'))) return false;
    if (studio.diffFilter === 'duplicate' && !row.flags.some((f) => f.includes('duplicate'))) return false;
    if (studio.diffFilter === 'outlier' && !row.flags.includes('outlier')) return false;

    if (searchLower) {
      const hay = st.columns
        .map((c) => `${row.raw?.[c] ?? ''} ${row.clean?.[c] ?? ''}`)
        .join(' ')
        .toLowerCase();
      if (!hay.includes(searchLower)) return false;
    }
    return true;
  });

  tbody.innerHTML = filtered
    .map((row) => {
      const changeCount = Object.keys(row.changes || {}).length;
      let badgeHtml = `<span class="status-tag clean">✓ Sạch</span>`;
      if (row.dropped) {
        badgeHtml = `<span class="status-tag dropped" title="${escapeHtml(row.dropReason)}">✕ Đã loại (${escapeHtml(
          row.flags.includes('duplicate_hidden') ? 'Trùng ngầm' : row.flags.includes('duplicate_exact') ? 'Trùng' : 'Outlier'
        )})</span>`;
      } else if (changeCount > 0 || row.flags.length > 0) {
        badgeHtml = `<span class="status-tag fixed">✨ Sửa ${changeCount} ô</span>`;
      }

      const cellsHtml = st.columns
        .map((col) => {
          const isNum = numericColSet.has(col);
          const ch = row.changes?.[col];
          const cleanVal = row.clean?.[col];
          const formattedClean = isNum && typeof cleanVal === 'number' ? fmtFullNum(cleanVal) : escapeHtml(cleanVal ?? '');

          if (ch && studio.showInlineDiff) {
            const oldDisp =
              typeof ch.old === 'number' ? fmtFullNum(ch.old) : escapeHtml(String(ch.old ?? ''));
            return `<td class="${isNum ? 'num-cell' : ''}">
              <span class="cell-diff-old">${oldDisp}</span><span class="cell-diff-arrow">➔</span><span class="cell-diff-new">${formattedClean}</span>
            </td>`;
          }
          return `<td class="${isNum ? 'num-cell' : ''}">${formattedClean}</td>`;
        })
        .join('');

      return `<tr class="${row.dropped ? 'row-dropped' : ''}">
        <td class="num-cell">${row.rowNumber}</td>
        <td>${badgeHtml}</td>
        ${cellsHtml}
      </tr>`;
    })
    .join('');

  document.getElementById('cleanerTableFooter').innerHTML = `
    <span>Đang hiển thị <strong>${filtered.length}</strong> / ${st.totalPreviewCount} dòng kiểm toán</span>
    <span>Mẹo: Bấm các nút lọc phía trên để xem riêng Ngày lệch, Gom tên Fuzzy, Trùng ngầm hoặc Ngoại lai</span>
  `;
}

// ============================================================================
// 5. RENDER TAB 2: GIẢI PHẪU BIẾN ĐỘNG (WATERFALL SVG + KEY DRIVERS + DRILLDOWN)
// ============================================================================

function renderWaterfallSvg(pvm, polarity = 'higher_is_better', p0Label = 'P0', p1Label = 'P1') {
  const steps = pvm.waterfall || [];
  if (!steps.length) return '';

  const width = 760;
  const height = 290;
  const padLeft = 64;
  const padRight = 24;
  const padTop = 34;
  const padBottom = 56;
  const plotW = width - padLeft - padRight;
  const plotH = height - padTop - padBottom;

  let maxVal = 0;
  let minVal = 0;
  for (const s of steps) {
    maxVal = Math.max(maxVal, s.start, s.end);
    minVal = Math.min(minVal, s.start, s.end, 0);
  }
  const span = Math.max(1, maxVal - minVal);
  const yMax = maxVal + span * 0.14;
  const yMin = minVal < 0 ? minVal - span * 0.1 : 0;
  const yRange = Math.max(1, yMax - yMin);

  const yToPx = (v) => padTop + plotH - ((v - yMin) / yRange) * plotH;
  const slotW = plotW / steps.length;
  const barW = Math.min(68, slotW * 0.58);

  // Lưới ngang 4 vạch
  const gridLines = [];
  for (let i = 0; i <= 4; i++) {
    const val = yMin + (yRange * i) / 4;
    const y = yToPx(val);
    gridLines.push(`
      <line x1="${padLeft}" y1="${y}" x2="${width - padRight}" y2="${y}" stroke="var(--border-subtle)" stroke-dasharray="3,3" />
      <text x="${padLeft - 8}" y="${y + 4}" text-anchor="end" fill="var(--text-muted)" font-size="10.5">${escapeHtml(
        fmtCompact(val)
      )}</text>
    `);
  }

  const barsAndConnectors = [];
  steps.forEach((st, idx) => {
    const cx = padLeft + slotW * idx + slotW / 2;
    const x = cx - barW / 2;
    const y1 = yToPx(st.start);
    const y2 = yToPx(st.end);
    const topY = Math.min(y1, y2);
    const barH = Math.max(3, Math.abs(y2 - y1));

    let fill = '#38bdf8';
    if (st.kind === 'total') {
      fill = idx === 0 ? '#64748b' : '#0ea5e9';
    } else {
      const good = isFavorableDelta(st.value, polarity);
      fill = good ? '#10b981' : '#f43f5e';
    }

    // Đường nối bậc thang sang cột kế tiếp
    if (idx < steps.length - 1) {
      const nextCx = padLeft + slotW * (idx + 1) + slotW / 2;
      const connY = yToPx(st.end);
      barsAndConnectors.push(`
        <line x1="${x + barW}" y1="${connY}" x2="${nextCx - barW / 2}" y2="${connY}" stroke="var(--text-muted)" stroke-width="1.2" stroke-dasharray="4,3" />
      `);
    }

    const valText = st.kind === 'total' ? fmtCompact(st.value) : fmtSigned(st.value);
    const subPctText = st.kind === 'total' ? '' : `(${fmtPct(st.pctOfP0)})`;
    const labelY = st.kind !== 'total' && st.value < 0 ? topY + barH + 13 : topY - 7;

    const shortAxisLabel =
      st.id === 'p0'
        ? `Kỳ ${p0Label}`
        : st.id === 'p1'
        ? `Kỳ ${p1Label}`
        : st.label.replace(/\s*\(.*\)/, '');

    barsAndConnectors.push(`
      <g>
        <title>${escapeHtml(st.label)}: ${escapeHtml(valText)} ${escapeHtml(subPctText)}</title>
        <rect x="${x}" y="${topY}" width="${barW}" height="${barH}" rx="5" fill="${fill}" opacity="0.92" />
        <text x="${cx}" y="${labelY}" text-anchor="middle" fill="var(--text-main)" font-size="11" font-weight="700">${escapeHtml(
          valText
        )}</text>
        ${
          subPctText
            ? `<text x="${cx}" y="${
                st.value < 0 ? labelY + 12 : labelY - 12
              }" text-anchor="middle" fill="var(--text-muted)" font-size="10">${escapeHtml(subPctText)}</text>`
            : ''
        }
        <text x="${cx}" y="${height - 26}" text-anchor="middle" fill="var(--text-secondary)" font-size="11" font-weight="600">${escapeHtml(
          shortAxisLabel
        )}</text>
      </g>
    `);
  });

  return `
    <svg viewBox="0 0 ${width} ${height}" role="img" aria-label="Biểu đồ thác nước cầu nối biến động PVM">
      ${gridLines.join('')}
      ${barsAndConnectors.join('')}
    </svg>
  `;
}

function renderVarianceTab() {
  const st = studio.state;
  const vr = st?.variance;
  if (!vr || !vr.ready) {
    document.getElementById('waterfallSvgContainer').innerHTML = `<div class="muted">${escapeHtml(
      vr?.reason || 'Chưa đủ dữ liệu kỳ thời gian để phân tích biến động.'
    )}</div>`;
    return;
  }

  // Controls
  document.getElementById('selGranularity').value = vr.granularity || 'month';
  document.getElementById('selPolarity').value = vr.polarity || 'higher_is_better';

  const pOpts = vr.availablePeriods
    .map((p) => `<option value="${escapeHtml(p.period)}">${escapeHtml(p.period)} (${p.rowCount} dòng)</option>`)
    .join('');
  const selP0 = document.getElementById('selPeriod0');
  const selP1 = document.getElementById('selPeriod1');
  selP0.innerHTML = pOpts;
  selP1.innerHTML = pOpts;
  selP0.value = vr.period0;
  selP1.value = vr.period1;

  // Mini period trend pills
  document.getElementById('periodTrendMini').innerHTML = vr.periodTrend
    .map(
      (pt) => `
      <div class="trend-pill ${pt.period === vr.period0 || pt.period === vr.period1 ? 'active-p' : ''}">
        <div class="muted">${escapeHtml(pt.period)}</div>
        <strong>${fmtCompact(pt.total)}</strong>
      </div>
    `
    )
    .join('');

  // Waterfall Chart & Step Cards
  const pvm = vr.pvm;
  document.getElementById('pvmSubtitle').textContent = `Đối chiếu ${st.mapping.metricCol} từ ${vr.period0} (${fmtCompact(
    pvm.totalR0
  )}) sang ${vr.period1} (${fmtCompact(pvm.totalR1)}) • Chênh lệch: ${fmtSigned(pvm.netDelta)} (${fmtPct(pvm.netDeltaPct)})`;

  const resErr = Math.abs(pvm.components.residualCheck || 0);
  document.getElementById('pvmExactBadge').textContent =
    resErr < 0.01 ? '✓ Khớp toán học 100% (Sai số = 0)' : `Sai số: ${resErr.toFixed(2)}`;

  document.getElementById('waterfallSvgContainer').innerHTML = renderWaterfallSvg(
    pvm,
    vr.polarity,
    vr.period0,
    vr.period1
  );

  document.getElementById('pvmCardsStrip').innerHTML = pvm.steps
    .map((s) => {
      const good = isFavorableDelta(s.delta, vr.polarity);
      return `
        <div class="pvm-mini-card">
          <div class="title">${escapeHtml(s.label)}</div>
          <div class="val" style="color:${good ? 'var(--good)' : 'var(--bad)'}">
            ${fmtSigned(s.delta)} <small>(${fmtPct(s.pctOfP0)})</small>
          </div>
          <div class="muted" style="font-size:11px;margin-top:3px">${escapeHtml(s.description)}</div>
        </div>
      `;
    })
    .join('');

  // Executive Memo + Dirty Trap Callout
  document.getElementById('executiveMemoBox').textContent = vr.executiveMemo || '';
  const dt = vr.dirtyTrap;
  const trapEl = document.getElementById('dirtyTrapCallout');
  if (dt && (Math.abs(dt.p1Distortion) > 1 || Math.abs(dt.pctPointError) >= 0.2)) {
    trapEl.classList.remove('hidden');
    trapEl.innerHTML = `
      <strong>🪤 Kiểm chứng giá trị làm sạch dữ liệu:</strong> Nếu tổng hợp trực tiếp trên file bẩn gốc chưa lọc,
      kỳ ${escapeHtml(vr.period1)} sẽ hiển thị <strong>${fmtCompact(dt.rawP1)}</strong> (biến động ảo <strong>${fmtPct(
        dt.rawDeltaPct
      )}</strong>) thay vì con số chuẩn <strong>${fmtCompact(pvm.totalR1)} (${fmtPct(pvm.netDeltaPct)})</strong>!
    `;
  } else {
    trapEl.classList.add('hidden');
  }

  // Multi-Dimension Key Drivers
  const activeDrillDim = vr.deepDive?.filterDim;
  const activeDrillVal = vr.deepDive?.filterVal;

  document.getElementById('dimensionDriversGrid').innerHTML = vr.dimensionDrivers
    .map((dimBlock) => {
      const maxAbs = Math.max(1, ...dimBlock.allMembers.map((m) => Math.abs(m.delta)));
      // Hiển thị tối đa 6 thành viên đáng chú ý nhất (ưu tiên các đầu kéo giảm mạnh nhất & kéo tăng mạnh nhất)
      const shown =
        dimBlock.allMembers.length <= 6
          ? dimBlock.allMembers
          : [...dimBlock.topNegative, ...dimBlock.topPositive.slice().reverse()];

      return `
        <div class="driver-dim-card">
          <div class="col-card-top">
            <h4>🧭 Theo ${escapeHtml(dimBlock.dimension)}</h4>
            <span class="muted">${dimBlock.memberCount} nhóm</span>
          </div>
          ${shown
            .map((m) => {
              const good = isFavorableDelta(m.delta, vr.polarity);
              const barPct = Math.min(100, Math.round((Math.abs(m.delta) / maxAbs) * 100));
              const isSelected = activeDrillDim === dimBlock.dimension && activeDrillVal === m.name;
              return `
                <div class="driver-row ${isSelected ? 'selected-drill' : ''}" data-drill-dim="${escapeHtml(
                  dimBlock.dimension
                )}" data-drill-val="${escapeHtml(m.name)}">
                  <div class="driver-row-top">
                    <span>${escapeHtml(m.name)}</span>
                    <span style="color:${good ? 'var(--good)' : 'var(--bad)'}">
                      ${fmtSigned(m.delta)} (${fmtPct(m.deltaPct)})
                    </span>
                  </div>
                  <div class="muted" style="font-size:11.3px;display:flex;justify-content:space-between;margin-top:2px">
                    <span>${escapeHtml(vr.period0)}: ${fmtCompact(m.p0)} ➔ ${escapeHtml(vr.period1)}: ${fmtCompact(m.p1)}</span>
                    <span>Tác động nền: ${fmtPct(m.impactOnTotalPct)}</span>
                  </div>
                  <div class="driver-bar-bg">
                    <div class="driver-bar-fill" style="width:${barPct}%;background:${
                      good ? 'var(--good)' : 'var(--bad)'
                    }"></div>
                  </div>
                </div>
              `;
            })
            .join('')}
        </div>
      `;
    })
    .join('');

  // Render 2-Level Deep Dive
  renderDeepDiveSection(vr.deepDive, vr.polarity, vr.period0, vr.period1);

  // Render Item-Level PVM Matrix Table
  const pvmThead = document.querySelector('#pvmBreakdownTable thead');
  const pvmTbody = document.querySelector('#pvmBreakdownTable tbody');

  pvmThead.innerHTML = `
    <tr>
      <th>${escapeHtml(st.mapping.itemCol || 'Mặt hàng / Hạng mục')}</th>
      <th class="num-cell">SL (${escapeHtml(vr.period0)})</th>
      <th class="num-cell">SL (${escapeHtml(vr.period1)})</th>
      <th class="num-cell">Đơn giá (${escapeHtml(vr.period0)})</th>
      <th class="num-cell">Đơn giá (${escapeHtml(vr.period1)})</th>
      <th class="num-cell">Tỷ trọng SL</th>
      <th class="num-cell">Δ Sản lượng (Vol)</th>
      <th class="num-cell">Δ Cơ cấu (Mix)</th>
      <th class="num-cell">Δ Đơn giá (Price)</th>
      <th class="num-cell">Tổng Chênh lệch</th>
    </tr>
  `;

  pvmTbody.innerHTML = pvm.itemBreakdown
    .map((b) => {
      const colorOf = (val) =>
        val === 0 ? 'var(--text-muted)' : isFavorableDelta(val, vr.polarity) ? 'var(--good)' : 'var(--bad)';
      return `
        <tr>
          <td><strong>${escapeHtml(b.item)}</strong></td>
          <td class="num-cell">${fmtFullNum(b.q0)}</td>
          <td class="num-cell">${fmtFullNum(b.q1)}</td>
          <td class="num-cell">${fmtCompact(b.p0)}</td>
          <td class="num-cell">${fmtCompact(b.p1)}</td>
          <td class="num-cell">${b.share0Pct}% ➔ ${b.share1Pct}%</td>
          <td class="num-cell" style="color:${colorOf(b.volumeImpact)}">${fmtSigned(b.volumeImpact)}</td>
          <td class="num-cell" style="color:${colorOf(b.mixImpact)}">${fmtSigned(b.mixImpact)}</td>
          <td class="num-cell" style="color:${colorOf(b.priceImpact)}">${fmtSigned(b.priceImpact)}</td>
          <td class="num-cell" style="font-weight:700;color:${colorOf(b.deltaTotal)}">
            ${fmtSigned(b.deltaTotal)} (${fmtPct(b.deltaPct)})
          </td>
        </tr>
      `;
    })
    .join('');
}

function renderDeepDiveSection(deepDive, polarity, p0, p1) {
  const card = document.getElementById('deepDiveCard');
  if (!deepDive) {
    card.classList.add('hidden');
    return;
  }
  card.classList.remove('hidden');

  document.getElementById('deepDiveTitle').textContent = `🕵️ Giải Phẫu Chéo Tầng 2: [${deepDive.filterDim} = ${deepDive.filterVal}]`;
  document.getElementById('deepDiveSubtitle').textContent = `Biến động riêng của "${deepDive.filterVal}" từ ${p0} (${fmtCompact(
    deepDive.pvm.totalR0
  )}) sang ${p1} (${fmtCompact(deepDive.pvm.totalR1)}): ${fmtSigned(deepDive.pvm.netDelta)} (${fmtPct(
    deepDive.pvm.netDeltaPct
  )})`;

  const subPvmHtml = `
    <div class="schema-table-box">
      <h4 style="margin-bottom:8px">🔬 PVM Cục Bộ Của "${escapeHtml(deepDive.filterVal)}"</h4>
      <div class="pvm-cards-strip" style="margin-top:0">
        ${deepDive.pvm.steps
          .map((s) => {
            const good = isFavorableDelta(s.delta, polarity);
            return `
              <div class="pvm-mini-card">
                <div class="title">${escapeHtml(s.label)}</div>
                <div class="val" style="color:${good ? 'var(--good)' : 'var(--bad)'}">${fmtSigned(s.delta)}</div>
                <div class="muted" style="font-size:11px">${fmtPct(s.pctOfP0)}</div>
              </div>
            `;
          })
          .join('')}
      </div>
    </div>
  `;

  const subDimsHtml = `
    <div class="sub-driver-columns">
      ${deepDive.subDrivers
        .slice(0, 3)
        .map(
          (sd) => `
          <div class="schema-table-box">
            <h4 style="margin-bottom:6px;font-size:12.5px">📌 Chi tiết theo ${escapeHtml(sd.dimension)}</h4>
            ${sd.allMembers
              .slice(0, 4)
              .map((m) => {
                const good = isFavorableDelta(m.delta, polarity);
                return `
                  <div style="display:flex;justify-content:space-between;gap:6px;padding:4px 0;border-bottom:1px solid var(--border-subtle);font-size:12px">
                    <span>${escapeHtml(m.name)}</span>
                    <strong style="color:${good ? 'var(--good)' : 'var(--bad)'}">${fmtSigned(m.delta)}</strong>
                  </div>
                `;
              })
              .join('')}
          </div>
        `
        )
        .join('')}
    </div>
  `;

  document.getElementById('deepDiveContent').innerHTML = subPvmHtml + subDimsHtml;
}

// ============================================================================
// 6. RENDER TAB 3: SQLITE WORKBENCH
// ============================================================================

function renderSqlMeta() {
  const meta = studio.state?.sqlMeta;
  if (!meta) return;

  // Schema Tree
  document.getElementById('sqlSchemaTree').innerHTML = meta.tables
    .map(
      (t) => `
      <div class="schema-table-box">
        <div class="schema-table-title" data-insert-sql="${escapeHtml(t.tableName)}">
          <span>📋 ${escapeHtml(t.tableName)}</span>
          <span class="muted">${t.rowCount} dòng</span>
        </div>
        <div class="schema-cols-list">
          ${t.columns
            .map(
              (c) => `
              <span class="schema-col-chip" data-insert-sql="${escapeHtml(c.name)}" title="${escapeHtml(
                c.originalLabel ? `${c.originalLabel} (${c.type})` : c.type
              )}">${escapeHtml(c.name)} <small style="opacity:0.6">${escapeHtml(c.type)}</small></span>
            `
            )
            .join('')}
        </div>
      </div>
    `
    )
    .join('');

  // Dynamic Presets
  document.getElementById('sqlPresetsBar').innerHTML = meta.presets
    .map(
      (p, idx) => `
      <button class="sql-preset-btn" data-preset-idx="${idx}">
        <strong>${escapeHtml(p.title)}</strong>
        <span>${escapeHtml(p.description)}</span>
      </button>
    `
    )
    .join('');

  if (!document.getElementById('sqlEditor').value && meta.presets[0]) {
    document.getElementById('sqlEditor').value = meta.presets[0].sql;
  }
}

async function runSqlQuery(sqlText) {
  const errBanner = document.getElementById('sqlErrorBanner');
  errBanner.classList.add('hidden');

  try {
    const res = await apiJson('/api/sql/query', {
      method: 'POST',
      body: JSON.stringify({ sql: sqlText })
    });
    studio.lastSqlExecuted = sqlText;

    document.getElementById('sqlTelemetryBadge').textContent = `⚡ ${res.elapsedMs} ms • ${res.totalRowCount} dòng`;

    const thead = document.querySelector('#sqlResultTable thead');
    const tbody = document.querySelector('#sqlResultTable tbody');

    thead.innerHTML = `<tr>${res.columns.map((c) => `<th>${escapeHtml(c)}</th>`).join('')}</tr>`;
    tbody.innerHTML = res.rows
      .map(
        (r) =>
          `<tr>${res.columns
            .map((c) => {
              const v = r[c];
              const isNum = typeof v === 'number' && Number.isFinite(v);
              return `<td class="${isNum ? 'num-cell' : ''}">${isNum ? fmtFullNum(v) : escapeHtml(v ?? 'NULL')}</td>`;
            })
            .join('')}</tr>`
      )
      .join('');
  } catch (err) {
    errBanner.textContent = `❌ Lỗi SQL: ${err.message}`;
    errBanner.classList.remove('hidden');
  }
}

// ============================================================================
// 7. RENDER TAB 4: KHO RECIPE 1-CHẠM
// ============================================================================

function renderRecipesTab() {
  populateRecipeDropdowns();
  const st = studio.state;

  // Gợi ý sẵn tên Recipe theo dataset hiện tại
  const inpName = document.getElementById('inpRecipeName');
  if (!inpName.value) {
    inpName.value = `Recipe Chuẩn — ${st?.datasetTitle || 'Báo cáo tuần'}`;
    document.getElementById('inpRecipeIndustry').value = st?.industry || 'Đa ngành';
  }

  document.getElementById('recipesGallery').innerHTML = studio.recipes
    .map((r) => {
      const isCurrent = st?.activeRecipeId === r.id;
      return `
        <div class="recipe-card ${isCurrent ? 'active-rcp' : ''}">
          <div>
            <div class="col-card-top">
              <strong>${escapeHtml(r.name)}</strong>
              <span class="type-badge">${r.builtin ? 'Chuẩn Hệ Thống' : 'Cá Nhân'}</span>
            </div>
            <p class="muted" style="margin-top:6px">${escapeHtml(r.description)}</p>
            <div class="recipe-meta-chips">
              <span class="variant-chip">Ngành: ${escapeHtml(r.industry)}</span>
              <span class="variant-chip">Fuzzy: ${Math.round((r.config?.fuzzyThreshold || 0.84) * 100)}%</span>
              <span class="variant-chip">Ngoại lai: ${escapeHtml(r.config?.outlierStrategy || 'median_impute')}</span>
              <span class="variant-chip">Đã chạy: ${r.runCount || 1} lần</span>
            </div>
          </div>
          <div class="btn-row" style="justify-content:flex-end;margin-top:4px">
            ${
              !r.builtin
                ? `<button class="btn btn-ghost btn-sm btn-del-recipe" data-rcp-id="${escapeHtml(r.id)}">🗑️ Xóa</button>`
                : ''
            }
            <button class="btn btn-primary btn-sm btn-apply-recipe" data-rcp-id="${escapeHtml(r.id)}">
              ⚡ Áp dụng 1-Chạm
            </button>
          </div>
        </div>
      `;
    })
    .join('');
}

// ============================================================================
// 8. ĐIỀU PHỐI CẬP NHẬT TOÀN BỘ GIAO DIỆN & GỌI API PIPELINE
// ============================================================================

function renderAll() {
  populateDatasetSelect();
  renderTopContextAndKpis();
  renderCleanerTab();
  renderVarianceTab();
  renderSqlMeta();
  renderRecipesTab();
}

async function triggerPipelineUpdate(overrides = {}) {
  const st = studio.state;
  const config = {
    ...st.config,
    standardizeDates: document.getElementById('chkDates').checked,
    standardizePhones: document.getElementById('chkPhones').checked,
    standardizeNumbers: document.getElementById('chkNumbers').checked,
    enableFuzzyMatching: document.getElementById('chkFuzzy').checked,
    fuzzyThreshold: Number(document.getElementById('rngFuzzy').value),
    dedupeStrategy: document.getElementById('selDedupe').value,
    outlierStrategy: document.getElementById('selOutlier').value,
    datePreference: document.getElementById('selDatePref').value,
    ...(overrides.config || {})
  };

  const mapping = {
    ...st.mapping,
    dateCol: document.getElementById('mapDateCol').value || null,
    metricCol: document.getElementById('mapMetricCol').value || null,
    quantityCol: document.getElementById('mapQtyCol').value || null,
    itemCol: document.getElementById('mapItemCol').value || null,
    phoneCol: document.getElementById('mapPhoneCol').value || null,
    idCol: document.getElementById('mapIdCol').value || null,
    ...(overrides.mapping || {})
  };

  const varianceParams = {
    granularity: document.getElementById('selGranularity').value || 'month',
    period0: document.getElementById('selPeriod0').value || null,
    period1: document.getElementById('selPeriod1').value || null,
    polarity: document.getElementById('selPolarity').value || 'higher_is_better',
    ...(overrides.varianceParams || {})
  };

  try {
    const data = await apiJson('/api/pipeline/run', {
      method: 'POST',
      body: JSON.stringify({ config, mapping, varianceParams })
    });
    studio.state = data.state;
    studio.lastSqlExecuted = '';
    renderAll();
  } catch (err) {
    showToast(`Lỗi cập nhật đường ống: ${err.message}`);
  }
}

async function handleFileUploadObject(file, recipeId = '') {
  if (!file) return;
  const arrayBuffer = await file.arrayBuffer();
  const bytes = new Uint8Array(arrayBuffer);
  let binary = '';
  const chunkSize = 0x8000;
  for (let i = 0; i < bytes.length; i += chunkSize) {
    binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunkSize));
  }
  const base64 = btoa(binary);

  const res = await apiJson('/api/upload', {
    method: 'POST',
    body: JSON.stringify({
      fileName: file.name,
      base64,
      recipeId: recipeId || null
    })
  });

  studio.recipes = res.recipes || studio.recipes;
  studio.state = res.state;
  studio.lastSqlExecuted = '';
  document.getElementById('sqlEditor').value = res.state.sqlMeta?.presets?.[0]?.sql || '';
  document.getElementById('uploadModal').classList.add('hidden');
  renderAll();

  const skipNote = res.skippedTitleRows > 0 ? ` (Đã tự bỏ qua ${res.skippedTitleRows} dòng tiêu đề trang trí đầu bảng)` : '';
  showToast(`✅ Đã nạp & làm sạch "${file.name}": ${res.state.stats.cleanRowCount} dòng chuẩn!${skipNote}`);
}

async function simulateIncomingNextPeriodBatch() {
  const baseDsId = studio.state?.datasetId === 'custom_upload' ? 'retail_fmcg' : studio.state?.datasetId || 'retail_fmcg';
  const selectedRecipeId =
    document.getElementById('selBatchRecipe').value || studio.state?.activeRecipeId || null;

  const res = await apiJson('/api/datasets/load-sample', {
    method: 'POST',
    body: JSON.stringify({
      datasetId: baseDsId,
      periodShift: 1,
      recipeId: selectedRecipeId
    })
  });

  studio.recipes = res.recipes || studio.recipes;
  studio.state = res.state;
  studio.lastSqlExecuted = '';
  document.getElementById('sqlEditor').value = res.state.sqlMeta?.presets?.[0]?.sql || '';
  renderAll();
  switchTab('variance');
  showToast(
    `🚀 Đã chạy Recipe 1-Chạm trên File Kỳ Mới (${res.state.variance.period0} ➔ ${res.state.variance.period1}) trong ${res.state.stats.elapsedMs} ms!`
  );
}

// ============================================================================
// 9. ĐĂNG KÝ SỰ KIỆN TƯƠNG TÁC (EVENT LISTENERS)
// ============================================================================

function bindEvents() {
  // Chuyển Tab
  document.querySelectorAll('.tab-btn').forEach((btn) => {
    btn.addEventListener('click', () => switchTab(btn.dataset.tab));
  });

  document.getElementById('btnJumpToVariance').addEventListener('click', () => switchTab('variance'));

  // Đổi giao diện Sáng / Tối
  document.getElementById('btnThemeToggle').addEventListener('click', () => {
    const root = document.documentElement;
    const next = root.getAttribute('data-theme') === 'dark' ? 'light' : 'dark';
    root.setAttribute('data-theme', next);
    document.getElementById('btnThemeToggle').textContent = next === 'dark' ? '☀️' : '🌙';
  });

  // Đổi bộ dữ liệu mẫu 5 ngành
  document.getElementById('datasetSelect').addEventListener('change', async (e) => {
    const id = e.target.value;
    if (id === 'custom_upload') return;
    const res = await apiJson('/api/datasets/load-sample', {
      method: 'POST',
      body: JSON.stringify({ datasetId: id, periodShift: 0 })
    });
    studio.recipes = res.recipes || studio.recipes;
    studio.state = res.state;
    studio.lastSqlExecuted = '';
    document.getElementById('sqlEditor').value = res.state.sqlMeta?.presets?.[0]?.sql || '';
    document.getElementById('inpRecipeName').value = `Recipe Chuẩn — ${res.state.datasetTitle}`;
    document.getElementById('inpRecipeIndustry').value = res.state.industry;
    renderAll();
    showToast(`Đã nạp bộ dữ liệu ngành: ${res.state.industry}`);
  });

  // Đổi Sheet trên file .xlsx đa Sheet
  document.getElementById('sheetSelect').addEventListener('change', async (e) => {
    const res = await apiJson('/api/upload/switch-sheet', {
      method: 'POST',
      body: JSON.stringify({ sheetName: e.target.value })
    });
    studio.state = res.state;
    studio.lastSqlExecuted = '';
    renderAll();
  });

  // Các nút cấu hình làm sạch tự động chạy lại
  ['chkDates', 'chkPhones', 'chkNumbers', 'chkFuzzy', 'selDedupe', 'selOutlier', 'selDatePref'].forEach((id) => {
    document.getElementById(id).addEventListener('change', () => triggerPipelineUpdate());
  });

  const rngFuzzy = document.getElementById('rngFuzzy');
  rngFuzzy.addEventListener('input', () => {
    document.getElementById('lblFuzzyVal').textContent = `${Math.round(Number(rngFuzzy.value) * 100)}%`;
  });
  rngFuzzy.addEventListener('change', () => triggerPipelineUpdate());

  // Ánh xạ cột
  document.getElementById('btnToggleMappingDrawer').addEventListener('click', () => {
    document.getElementById('mappingDrawer').classList.toggle('hidden');
  });
  document.getElementById('btnApplyMapping').addEventListener('click', () => {
    triggerPipelineUpdate({ varianceParams: { period0: null, period1: null } });
    showToast('Đã cập nhật ánh xạ vai trò cột & tính lại PVM Bridge!');
  });

  // Đổi tên chuẩn tùy chỉnh trong cụm Fuzzy
  document.getElementById('fuzzyClustersList').addEventListener('click', (e) => {
    const btn = e.target.closest('.btn-apply-canonical');
    if (!btn) return;
    const idx = Number(btn.dataset.fuzzyIdx);
    const cl = studio.state.fuzzyClusters[idx];
    const inp = document.querySelector(`input[data-fuzzy-idx="${idx}"]`);
    if (!cl || !inp || !inp.value.trim()) return;

    const newTarget = inp.value.trim();
    const curAliases = { ...(studio.state.config.customAliases || {}) };
    const colDict =
      typeof curAliases[cl.column] === 'object' && !Array.isArray(curAliases[cl.column])
        ? { ...curAliases[cl.column] }
        : {};
    for (const v of cl.variants) {
      colDict[v.name] = newTarget;
    }
    curAliases[cl.column] = colDict;
    triggerPipelineUpdate({ config: { customAliases: curAliases } });
    showToast(`Đã chuẩn hóa toàn bộ cụm [${cl.column}] về "${newTarget}"!`);
  });

  // Bộ lọc bảng Diff
  document.getElementById('diffFilterPills').addEventListener('click', (e) => {
    const btn = e.target.closest('.pill-btn');
    if (!btn) return;
    studio.diffFilter = btn.dataset.filter;
    document.querySelectorAll('#diffFilterPills .pill-btn').forEach((b) => b.classList.toggle('active', b === btn));
    renderCleanerDiffTable();
  });

  document.getElementById('gridSearchInput').addEventListener('input', (e) => {
    studio.diffSearch = e.target.value;
    renderCleanerDiffTable();
  });

  document.getElementById('chkShowDiffInline').addEventListener('change', (e) => {
    studio.showInlineDiff = e.target.checked;
    renderCleanerDiffTable();
  });

  // Tab 2: Đổi kỳ P0 / P1 / Granularity / Polarity
  document.getElementById('selGranularity').addEventListener('change', () => {
    triggerPipelineUpdate({ varianceParams: { period0: null, period1: null } });
  });
  ['selPeriod0', 'selPeriod1', 'selPolarity'].forEach((id) => {
    document.getElementById(id).addEventListener('change', () => triggerPipelineUpdate());
  });

  // Click bất kỳ dòng Key-Driver nào để Giải phẫu chéo tầng 2
  document.getElementById('dimensionDriversGrid').addEventListener('click', async (e) => {
    const rowEl = e.target.closest('.driver-row');
    if (!rowEl) return;
    const drillDim = rowEl.dataset.drillDim;
    const drillVal = rowEl.dataset.drillVal;

    const res = await apiJson('/api/variance/drilldown', {
      method: 'POST',
      body: JSON.stringify({ drillDim, drillVal })
    });
    studio.state.variance.deepDive = res.deepDive;
    document.querySelectorAll('.driver-row').forEach((el) => {
      el.classList.toggle(
        'selected-drill',
        el.dataset.drillDim === drillDim && el.dataset.drillVal === drillVal
      );
    });
    renderDeepDiveSection(
      res.deepDive,
      studio.state.variance.polarity,
      studio.state.variance.period0,
      studio.state.variance.period1
    );
    document.getElementById('deepDiveCard').scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  });

  // Copy Executive Memo
  document.getElementById('btnCopyMemo').addEventListener('click', async () => {
    const text = document.getElementById('executiveMemoBox').textContent;
    await navigator.clipboard.writeText(text);
    showToast('📋 Đã sao chép Bản giải trình biến động vào Clipboard!');
  });

  // Tab 3: SQLite Workbench
  document.getElementById('sqlPresetsBar').addEventListener('click', (e) => {
    const btn = e.target.closest('.sql-preset-btn');
    if (!btn) return;
    const preset = studio.state.sqlMeta.presets[Number(btn.dataset.presetIdx)];
    if (!preset) return;
    document.getElementById('sqlEditor').value = preset.sql;
    runSqlQuery(preset.sql);
  });

  document.getElementById('sqlSchemaTree').addEventListener('click', (e) => {
    const chip = e.target.closest('[data-insert-sql]');
    if (!chip) return;
    const token = chip.dataset.insertSql;
    const ed = document.getElementById('sqlEditor');
    const start = ed.selectionStart || ed.value.length;
    const end = ed.selectionEnd || ed.value.length;
    ed.value = ed.value.slice(0, start) + token + ed.value.slice(end);
    ed.focus();
  });

  document.getElementById('btnRunSql').addEventListener('click', () => {
    runSqlQuery(document.getElementById('sqlEditor').value);
  });

  document.getElementById('sqlEditor').addEventListener('keydown', (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
      e.preventDefault();
      runSqlQuery(e.target.value);
    }
  });

  const exportSqlAs = async (format) => {
    const sql = document.getElementById('sqlEditor').value;
    const res = await fetch('/api/sql/export', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ sql, format })
    });
    if (!res.ok) {
      const err = await res.json();
      showToast(`Lỗi xuất SQL: ${err.error}`);
      return;
    }
    const blob = await res.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `ket_qua_sql.${format}`;
    a.click();
    URL.revokeObjectURL(url);
  };

  document.getElementById('btnExportSqlCsv').addEventListener('click', () => exportSqlAs('csv'));
  document.getElementById('btnExportSqlXlsx').addEventListener('click', () => exportSqlAs('xlsx'));

  // Tab 4: Kho Recipe & Giả lập File Kỳ Mới
  document.getElementById('btnOpenSaveRecipe').addEventListener('click', () => {
    switchTab('recipes');
    document.getElementById('inpRecipeName').focus();
  });

  document.getElementById('formSaveRecipe').addEventListener('submit', async (e) => {
    e.preventDefault();
    const name = document.getElementById('inpRecipeName').value.trim();
    const industry = document.getElementById('inpRecipeIndustry').value.trim();
    const description = document.getElementById('inpRecipeDesc').value.trim();

    const res = await apiJson('/api/recipes', {
      method: 'POST',
      body: JSON.stringify({
        name,
        industry,
        description,
        config: studio.state.config,
        mapping: studio.state.mapping
      })
    });
    studio.recipes = res.recipes;
    studio.state.activeRecipeId = res.saved.id;
    renderAll();
    showToast(`💾 Đã lưu Recipe "${res.saved.name}" thành công!`);
  });

  document.getElementById('recipesGallery').addEventListener('click', async (e) => {
    const applyBtn = e.target.closest('.btn-apply-recipe');
    if (applyBtn) {
      const res = await apiJson('/api/recipes/apply', {
        method: 'POST',
        body: JSON.stringify({ recipeId: applyBtn.dataset.rcpId })
      });
      studio.recipes = res.recipes;
      studio.state = res.state;
      renderAll();
      showToast('⚡ Đã áp dụng Recipe 1-Chạm lên bộ dữ liệu đang mở!');
      return;
    }
    const delBtn = e.target.closest('.btn-del-recipe');
    if (delBtn) {
      const res = await apiJson(`/api/recipes/${encodeURIComponent(delBtn.dataset.rcpId)}`, {
        method: 'DELETE'
      });
      studio.recipes = res.recipes;
      renderRecipesTab();
      showToast('Đã xóa Recipe khỏi kho.');
    }
  });

  document.getElementById('btnSimulateNextBatch').addEventListener('click', simulateIncomingNextPeriodBatch);
  document.getElementById('btnSimulateNextMonthTop').addEventListener('click', simulateIncomingNextPeriodBatch);

  document.getElementById('batchFileInput').addEventListener('change', (e) => {
    const file = e.target.files?.[0];
    const rcpId = document.getElementById('selBatchRecipe').value;
    if (file) handleFileUploadObject(file, rcpId);
  });

  // Drag & Drop trên Recipe Dropzone
  const setupDropzone = (el, getRecipeFn) => {
    el.addEventListener('dragover', (e) => {
      e.preventDefault();
      el.classList.add('dragover');
    });
    el.addEventListener('dragleave', () => el.classList.remove('dragover'));
    el.addEventListener('drop', (e) => {
      e.preventDefault();
      el.classList.remove('dragover');
      const file = e.dataTransfer?.files?.[0];
      if (file) handleFileUploadObject(file, getRecipeFn());
    });
  };

  setupDropzone(document.getElementById('quickBatchDropzone'), () => document.getElementById('selBatchRecipe').value);
  setupDropzone(document.getElementById('modalDropArea'), () => document.getElementById('modalRecipeSelect').value);

  // Modal Nhập File / Dán Clipboard
  const modal = document.getElementById('uploadModal');
  document.getElementById('btnOpenUpload').addEventListener('click', () => modal.classList.remove('hidden'));
  document.getElementById('btnCloseUpload').addEventListener('click', () => modal.classList.add('hidden'));
  modal.addEventListener('click', (e) => {
    if (e.target === modal) modal.classList.add('hidden');
  });

  document.getElementById('modalFileInput').addEventListener('change', (e) => {
    const file = e.target.files?.[0];
    const rcpId = document.getElementById('modalRecipeSelect').value;
    if (file) handleFileUploadObject(file, rcpId);
  });

  document.getElementById('btnSubmitPaste').addEventListener('click', async () => {
    const text = document.getElementById('pasteTableArea').value.trim();
    if (!text) {
      showToast('Vui lòng dán bảng dữ liệu từ Excel hoặc CSV vào ô văn bản trước.');
      return;
    }
    try {
      const res = await apiJson('/api/upload', {
        method: 'POST',
        body: JSON.stringify({
          fileName: 'bang_dan_clipboard.tsv',
          text,
          recipeId: document.getElementById('modalRecipeSelect').value || null
        })
      });
      studio.recipes = res.recipes || studio.recipes;
      studio.state = res.state;
      studio.lastSqlExecuted = '';
      modal.classList.add('hidden');
      renderAll();
      showToast(`✅ Đã nạp ${res.state.stats.cleanRowCount} dòng sạch từ bảng dán Clipboard!`);
    } catch (err) {
      showToast(`Lỗi đọc bảng dán: ${err.message}`);
    }
  });
}

// ============================================================================
// 10. KHỞI CHẠY ỨNG DỤNG
// ============================================================================

async function initApp() {
  bindEvents();
  try {
    const boot = await apiJson('/api/bootstrap');
    studio.datasets = boot.datasets;
    studio.recipes = boot.recipes;
    studio.state = boot.state;
    renderAll();
  } catch (err) {
    showToast(`Không thể tải dữ liệu khởi tạo: ${err.message}`);
  }
}

initApp();
