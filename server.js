/**
 * 🧼📉 DataScrub & Variance Studio — Backend Server (Port 4195)
 * Trạm Làm Sạch Dữ Liệu Bẩn Đặc Thù Việt Nam, Tự Động Giải Phẫu Biến Động Doanh Số (PVM Bridge & Key-Driver),
 * SQLite Analytical Workbench & Xuất Báo Cáo Sạch 1-Chạm (.XLSX / .CSV)
 */

// Tự động nạp cờ --experimental-sqlite nếu người dùng chạy trực tiếp `node server.js`
try {
  require('node:sqlite');
} catch (_) {
  const { spawn } = require('node:child_process');
  const child = spawn(
    process.execPath,
    ['--experimental-sqlite', '--no-warnings', __filename, ...process.argv.slice(2)],
    { stdio: 'inherit', env: process.env }
  );
  child.on('exit', (code) => process.exit(code ?? 0));
  return;
}

const http = require('node:http');
const fs = require('node:fs');
const path = require('node:path');
const { URL } = require('node:url');

const {
  parseTabularUpload,
  buildMultiSheetXlsx,
  buildUtf8BomCsv
} = require('./lib/xlsx-codec');
const { DEFAULT_CLEAN_CONFIG, runCleaningPipeline } = require('./lib/viet-cleaner');
const {
  runVarianceAnalysis,
  computeSegmentDeepDive,
  isoDateToPeriod,
  formatVnCompact,
  formatSignedCompact,
  formatPct
} = require('./lib/variance-engine');
const { SqlWorkbench } = require('./lib/sql-workbench');
const {
  listRecipes,
  getRecipeById,
  upsertRecipe,
  incrementRecipeRun,
  deleteRecipe
} = require('./lib/recipe-store');
const { listSampleDatasets, getSampleDataset } = require('./lib/sample-datasets');

const PORT = Number(process.env.PORT || 4195);
const PUBLIC_DIR = path.join(__dirname, 'public');

// ============================================================================
// 1. BỘ ĐIỀU PHỐI TRẠNG THÁI PHIÊN LÀM VIỆC (STUDIO SESSION STATE)
// ============================================================================

const sqlWorkbench = new SqlWorkbench();

const session = {
  datasetId: 'retail_fmcg',
  datasetTitle: '',
  industry: '',
  sourceType: 'sample',
  fileName: 'retail_fmcg_dirty.xlsx',
  availableSheets: ['Data'],
  sheetName: 'Data',
  columns: [],
  rawRows: [],
  activeRecipeId: null,
  cleanResult: null,
  varianceResult: null,
  sqlMeta: null
};

function recomputeSession({
  config = null,
  mapping = null,
  varianceParams = {},
  recipeId = undefined
} = {}) {
  const effectiveConfig = config || session.cleanResult?.config || { ...DEFAULT_CLEAN_CONFIG };
  const cleanRes = runCleaningPipeline(session.columns, session.rawRows, effectiveConfig, mapping);

  const varRes = runVarianceAnalysis({
    datasetTitle: session.datasetTitle,
    rawRows: session.rawRows,
    cleanRows: cleanRes.cleanRows,
    cleanStats: cleanRes.stats,
    mapping: cleanRes.mapping,
    period0: varianceParams.period0 || null,
    period1: varianceParams.period1 || null,
    granularity: varianceParams.granularity || mapping?.granularity || 'month',
    polarity: varianceParams.polarity || cleanRes.mapping.polarity || 'higher_is_better',
    drillDim: varianceParams.drillDim || null,
    drillVal: varianceParams.drillVal || null
  });

  // Đồng bộ polarity vào mapping để lưu Recipe chính xác
  cleanRes.mapping.polarity = varRes.polarity || cleanRes.mapping.polarity || 'higher_is_better';
  cleanRes.mapping.granularity = varRes.granularity || 'month';

  const sqlMeta = sqlWorkbench.hydrate({
    cleanResult: cleanRes,
    rawRows: session.rawRows,
    varianceResult: varRes
  });

  if (recipeId !== undefined) {
    session.activeRecipeId = recipeId;
  }
  session.cleanResult = cleanRes;
  session.varianceResult = varRes;
  session.sqlMeta = sqlMeta;
}

function loadSampleIntoSession(datasetId = 'retail_fmcg', periodShift = 0, explicitRecipeId = null) {
  const ds = getSampleDataset(datasetId, periodShift);
  session.datasetId = ds.id;
  session.datasetTitle =
    periodShift > 0 ? `${ds.name} (File Kỳ Mới Tháng 11)` : ds.name;
  session.industry = ds.industry;
  session.sourceType = periodShift > 0 ? 'simulated_next_batch' : 'sample';
  session.fileName = `${ds.id}${periodShift > 0 ? '_batch_moi' : '_raw_dirty'}.xlsx`;
  session.availableSheets = ['Du_Lieu_Goc'];
  session.sheetName = 'Du_Lieu_Goc';
  session.columns = ds.columns;
  session.rawRows = ds.rows;

  const chosenRecipe = getRecipeById(explicitRecipeId || ds.recommendedRecipeId);
  if (chosenRecipe) {
    incrementRecipeRun(chosenRecipe.id);
    // Chỉ giữ các trường mapping khớp với cột thực tế của dataset
    const filteredMapping = sanitizeMappingForColumns(chosenRecipe.mapping, ds.columns);
    recomputeSession({
      config: chosenRecipe.config,
      mapping: filteredMapping,
      varianceParams: {
        granularity: chosenRecipe.mapping?.granularity || 'month',
        polarity: chosenRecipe.mapping?.polarity || null
      },
      recipeId: chosenRecipe.id
    });
  } else {
    recomputeSession({ recipeId: null });
  }
}

function sanitizeMappingForColumns(rawMapping, columns) {
  if (!rawMapping || typeof rawMapping !== 'object') return null;
  const colSet = new Set(columns || []);
  const out = {};
  for (const key of ['dateCol', 'idCol', 'phoneCol', 'metricCol', 'quantityCol', 'unitPriceCol', 'itemCol']) {
    if (rawMapping[key] && colSet.has(rawMapping[key])) {
      out[key] = rawMapping[key];
    }
  }
  if (Array.isArray(rawMapping.dimensionCols)) {
    const validDims = rawMapping.dimensionCols.filter((c) => colSet.has(c));
    if (validDims.length > 0) out.dimensionCols = validDims;
  }
  if (rawMapping.polarity) out.polarity = rawMapping.polarity;
  if (rawMapping.granularity) out.granularity = rawMapping.granularity;
  return Object.keys(out).length > 0 ? out : null;
}

function serializeActiveState(previewLimit = 250) {
  const cr = session.cleanResult;
  return {
    datasetId: session.datasetId,
    datasetTitle: session.datasetTitle,
    industry: session.industry,
    sourceType: session.sourceType,
    fileName: session.fileName,
    sheetName: session.sheetName,
    availableSheets: session.availableSheets,
    activeRecipeId: session.activeRecipeId,
    columns: session.columns,
    sqlColumns: cr.sqlColumns,
    mapping: cr.mapping,
    config: cr.config,
    stats: cr.stats,
    columnHealth: cr.columnHealth,
    fuzzyClusters: cr.fuzzyClusters,
    outliers: cr.outliers,
    auditLogs: cr.auditLogs.slice(0, 400),
    totalAuditLogCount: cr.auditLogs.length,
    stagedPreview: cr.stagedPreview.slice(0, previewLimit),
    totalPreviewCount: cr.stagedPreview.length,
    variance: session.varianceResult,
    sqlMeta: session.sqlMeta
  };
}

// Khởi tạo dữ liệu mẫu Bán lẻ FMCG ngay khi bật Server
loadSampleIntoSession('retail_fmcg', 0);

// ============================================================================
// 2. BỘ XUẤT WORKBOOK EXCEL .XLSX 4 SHEET CHUẨN GỬI SẾP
// ============================================================================

function buildExecutiveWorkbookBuffer() {
  const cr = session.cleanResult;
  const vr = session.varianceResult;

  // SHEET 1: TỔNG QUAN & GIẢI PHẪU BIẾN ĐỘNG CHO SẾP
  const s1Matrix = [];
  s1Matrix.push([
    { v: `BÁO CÁO LÀM SẠCH DỮ LIỆU & GIẢI PHẪU BIẾN ĐỘNG — ${session.datasetTitle.toUpperCase()}`, s: 1 },
    { v: '', s: 1 },
    { v: '', s: 1 },
    { v: '', s: 1 },
    { v: '', s: 1 }
  ]);
  s1Matrix.push([
    { v: 'Tệp nguồn:', s: 2 },
    { v: session.fileName, s: 0 },
    { v: 'Điểm sức khỏe dữ liệu:', s: 2 },
    { v: `${cr.stats.rawHealthScore}/100 ➔ ${cr.stats.cleanHealthScore}/100`, s: 6 },
    { v: `Đã xử lý ${cr.stats.totalIssuesFound} lỗi bẩn`, s: 0 }
  ]);

  if (vr && vr.ready) {
    const pvm = vr.pvm;
    s1Matrix.push([]);
    s1Matrix.push([
      { v: `1. CẦU NỐI BIẾN ĐỘNG PRICE-VOLUME-MIX (${vr.period0} ➔ ${vr.period1})`, s: 2 },
      { v: 'Giá trị tuyệt đối', s: 2 },
      { v: '% Tác động so với Kỳ gốc', s: 2 },
      { v: 'Diễn giải nghiệp vụ', s: 2 },
      { v: '', s: 2 }
    ]);
    s1Matrix.push([
      { v: `Kỳ gốc (${vr.period0})`, s: 0 },
      { v: Math.round(pvm.totalR0), t: 'n', s: 3 },
      { v: 1, t: 'n', s: 5 },
      { v: `Tổng ${cr.mapping.metricCol} sau làm sạch ở kỳ gốc`, s: 0 },
      { v: '', s: 0 }
    ]);

    for (const step of pvm.steps) {
      s1Matrix.push([
        { v: `Δ ${step.label}`, s: 0 },
        { v: Math.round(step.delta), t: 'n', s: step.delta >= 0 ? 7 : 8 },
        { v: step.pctOfP0 / 100, t: 'n', s: 5 },
        { v: step.description, s: 0 },
        { v: '', s: 0 }
      ]);
    }

    s1Matrix.push([
      { v: `Kỳ so sánh (${vr.period1})`, s: 0 },
      { v: Math.round(pvm.totalR1), t: 'n', s: 3 },
      { v: pvm.totalR0 !== 0 ? pvm.totalR1 / pvm.totalR0 : 1, t: 'n', s: 5 },
      { v: `Chênh lệch ròng: ${formatSignedCompact(pvm.netDelta)} (${formatPct(pvm.netDeltaPct)})`, s: 6 },
      { v: '', s: 0 }
    ]);

    // Thêm bảng Top Drivers theo từng chiều
    s1Matrix.push([]);
    s1Matrix.push([
      { v: '2. CHỈ ĐÍCH DANH ĐƠN VỊ / CÁ NHÂN TÁC ĐỘNG MẠNH NHẤT (KEY DRIVERS)', s: 2 },
      { v: `Kỳ ${vr.period0}`, s: 2 },
      { v: `Kỳ ${vr.period1}`, s: 2 },
      { v: 'Chênh lệch', s: 2 },
      { v: '% Tăng/Giảm', s: 2 }
    ]);

    for (const dimBlock of vr.dimensionDrivers) {
      for (const m of dimBlock.allMembers.slice(0, 6)) {
        s1Matrix.push([
          { v: `[${dimBlock.dimension}] ${m.name}`, s: 0 },
          { v: Math.round(m.p0), t: 'n', s: 3 },
          { v: Math.round(m.p1), t: 'n', s: 3 },
          { v: Math.round(m.delta), t: 'n', s: m.delta >= 0 ? 7 : 8 },
          { v: m.deltaPct / 100, t: 'n', s: 5 }
        ]);
      }
    }

    // Thêm Bản giải trình gửi sếp
    s1Matrix.push([]);
    s1Matrix.push([{ v: '3. BẢN TÓM TẮT ĐIỀU HÀNH GỬI LÃNH ĐẠO (EXECUTIVE MEMO)', s: 2 }]);
    for (const line of String(vr.executiveMemo || '').split('\n')) {
      s1Matrix.push([{ v: line, s: 0 }]);
    }
  }

  // SHEET 2: BẢNG GIẢI PHẪU PVM CHI TIẾT TỪNG MẶT HÀNG
  const s2Columns = [
    { key: 'item', label: cr.mapping.itemCol || 'Mặt hàng / Nhóm' },
    { key: 'status', label: 'Trạng thái' },
    { key: 'q0', label: `Sản lượng (${vr?.period0 || 'P0'})` },
    { key: 'q1', label: `Sản lượng (${vr?.period1 || 'P1'})` },
    { key: 'p0', label: `Đơn giá TB (${vr?.period0 || 'P0'})` },
    { key: 'p1', label: `Đơn giá TB (${vr?.period1 || 'P1'})` },
    { key: 'r0', label: `Giá trị (${vr?.period0 || 'P0'})` },
    { key: 'r1', label: `Giá trị (${vr?.period1 || 'P1'})` },
    { key: 'volumeImpact', label: 'Tác động Sản lượng (Volume)', format: 'delta' },
    { key: 'mixImpact', label: 'Tác động Cơ cấu (Mix)', format: 'delta' },
    { key: 'priceImpact', label: 'Tác động Đơn giá (Price)', format: 'delta' },
    { key: 'portfolioImpact', label: 'Mới / Ngưng phát sinh', format: 'delta' },
    { key: 'deltaTotal', label: 'Tổng Chênh lệch', format: 'delta' }
  ];

  // SHEET 3: DỮ LIỆU ĐÃ LÀM SẠCH CHUẨN HÓA (SẴN SÀNG PIVOT / POWER BI)
  const s3Cols = [...session.columns, ...(cr.cleanRows[0]?._period ? ['_period'] : [])];
  const s3Columns = s3Cols.map((c) => ({
    key: c,
    label: c === '_period' ? 'Kỳ chuẩn (YYYY-MM)' : c
  }));

  // SHEET 4: NHẬT KÝ KIỂM TOÁN DỮ LIỆU BẨN ĐÃ XỬ LÝ
  const s4Columns = [
    { key: 'rowNumber', label: 'Dòng gốc #' },
    { key: 'column', label: 'Cột dữ liệu' },
    { key: 'issueType', label: 'Mã loại lỗi' },
    { key: 'oldValue', label: 'Giá trị bẩn ban đầu' },
    { key: 'newValue', label: 'Giá trị chuẩn sau xử lý' },
    { key: 'detail', label: 'Diễn giải quy tắc xử lý' }
  ];

  return buildMultiSheetXlsx([
    { name: '1_Tong_Quan_Cho_Sep', matrix: s1Matrix },
    { name: '2_Chi_Tiet_PVM_SKU', columns: s2Columns, rows: vr?.pvm?.itemBreakdown || [] },
    { name: '3_Du_Lieu_Sach_Chuan', columns: s3Columns, rows: cr.cleanRows },
    { name: '4_Nhat_Ky_Lam_Sach', columns: s4Columns, rows: cr.auditLogs }
  ]);
}

// ============================================================================
// 3. HTTP HELPERS & STATIC ASSET SERVER
// ============================================================================

const MIME_TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.js': 'application/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.ico': 'image/x-icon'
};

function sendJson(res, statusCode, data) {
  const body = JSON.stringify(data);
  res.writeHead(statusCode, {
    'Content-Type': 'application/json; charset=utf-8',
    'Access-Control-Allow-Origin': '*',
    'Cache-Control': 'no-store'
  });
  res.end(body);
}

function readJsonBody(req, maxBytes = 25 * 1024 * 1024) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    let total = 0;
    req.on('data', (chunk) => {
      total += chunk.length;
      if (total > maxBytes) {
        reject(new Error('Dữ liệu tải lên vượt quá giới hạn 25MB.'));
        req.destroy();
        return;
      }
      chunks.push(chunk);
    });
    req.on('end', () => {
      try {
        const raw = Buffer.concat(chunks).toString('utf8');
        resolve(raw ? JSON.parse(raw) : {});
      } catch (err) {
        reject(new Error('Payload JSON không hợp lệ: ' + err.message));
      }
    });
    req.on('error', reject);
  });
}

function serveStaticFile(reqPath, res) {
  const rel = reqPath === '/' ? 'index.html' : reqPath.replace(/^\/+/, '');
  const resolved = path.resolve(PUBLIC_DIR, rel);
  if (!resolved.startsWith(PUBLIC_DIR)) {
    sendJson(res, 403, { error: 'Forbidden' });
    return;
  }
  if (!fs.existsSync(resolved) || fs.statSync(resolved).isDirectory()) {
    sendJson(res, 404, { error: 'Không tìm thấy trang yêu cầu.' });
    return;
  }
  const ext = path.extname(resolved).toLowerCase();
  res.writeHead(200, {
    'Content-Type': MIME_TYPES[ext] || 'application/octet-stream',
    'Cache-Control': 'no-cache'
  });
  fs.createReadStream(resolved).pipe(res);
}

// ============================================================================
// 4. ĐỊNH TUYẾN API CHÍNH
// ============================================================================

const server = http.createServer(async (req, res) => {
  if (req.method === 'OPTIONS') {
    res.writeHead(204, {
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type'
    });
    return res.end();
  }

  const parsedUrl = new URL(req.url, `http://${req.headers.host || 'localhost'}`);
  const pathname = parsedUrl.pathname;

  try {
    // 1. Health check
    if (req.method === 'GET' && pathname === '/api/health') {
      return sendJson(res, 200, {
        status: 'ok',
        service: 'dirty-data-variance-studio',
        port: PORT,
        sqliteReady: true,
        activeDataset: session.datasetTitle,
        cleanRowCount: session.cleanResult?.stats?.cleanRowCount || 0
      });
    }

    // 2. Bootstrap toàn bộ Studio trong 1 lần gọi
    if (req.method === 'GET' && pathname === '/api/bootstrap') {
      return sendJson(res, 200, {
        datasets: listSampleDatasets(),
        recipes: listRecipes(),
        state: serializeActiveState()
      });
    }

    // 3. Chuyển đổi sang 1 trong 5 bộ dữ liệu mẫu đa ngành (hoặc giả lập file kỳ mới tuần/tháng sau!)
    if (req.method === 'POST' && pathname === '/api/datasets/load-sample') {
      const body = await readJsonBody(req);
      const datasetId = body.datasetId || 'retail_fmcg';
      const periodShift = Number(body.periodShift || 0);
      const recipeId = body.recipeId || null;
      loadSampleIntoSession(datasetId, periodShift, recipeId);
      return sendJson(res, 200, {
        recipes: listRecipes(),
        state: serializeActiveState()
      });
    }

    // 4. Tải file Excel Bẩn Nguyên Bản (.XLSX) của bất kỳ ngành nào về máy để xem hoặc kéo-thả thử
    if (req.method === 'GET' && pathname === '/api/datasets/download-raw-xlsx') {
      const dsId = parsedUrl.searchParams.get('datasetId') || session.datasetId || 'retail_fmcg';
      const shift = Number(parsedUrl.searchParams.get('periodShift') || 0);
      const ds = getSampleDataset(dsId, shift);
      const xlsxBuf = buildMultiSheetXlsx([
        {
          name: 'Du_Lieu_Phong_Ban_Gui',
          columns: ds.columns,
          rows: ds.rows
        }
      ]);
      const fileName = `${ds.id}${shift > 0 ? '_ky_moi' : '_file_ban_goc'}.xlsx`;
      res.writeHead(200, {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': `attachment; filename="${fileName}"`,
        'Content-Length': xlsxBuf.length
      });
      return res.end(xlsxBuf);
    }

    // 5. Upload tệp Excel (.xlsx) / CSV / TSV / JSON hoặc Dán trực tiếp từ Clipboard Excel
    if (req.method === 'POST' && pathname === '/api/upload') {
      const body = await readJsonBody(req);
      const fileName = String(body.fileName || 'du_lieu_tai_len.xlsx').trim();
      const buffer = body.base64 ? Buffer.from(body.base64, 'base64') : undefined;
      const text = body.text !== undefined ? String(body.text) : undefined;

      const parsed = parseTabularUpload({
        fileName,
        buffer,
        text,
        sheetName: body.sheetName || null
      });

      if (!parsed.columns.length || !parsed.rows.length) {
        return sendJson(res, 400, {
          error: 'Không đọc được dòng dữ liệu hợp lệ nào từ tệp tải lên. Vui lòng kiểm tra lại cấu trúc bảng.'
        });
      }

      session.datasetId = 'custom_upload';
      session.datasetTitle = body.title || fileName.replace(/\.[^.]+$/, '');
      session.industry = 'Dữ liệu Tải lên Thực tế';
      session.sourceType = 'upload';
      session.fileName = fileName;
      session.availableSheets = parsed.availableSheets || ['Data'];
      session.sheetName = parsed.sheetName || 'Data';
      session.columns = parsed.columns;
      session.rawRows = parsed.rows;
      // Lưu lại buffer tạm nếu người dùng đổi Sheet trên cùng file .xlsx
      session._lastUploadBuffer = buffer || null;

      const recipe = body.recipeId ? getRecipeById(body.recipeId) : null;
      if (recipe) {
        incrementRecipeRun(recipe.id);
        const mapped = sanitizeMappingForColumns(recipe.mapping, parsed.columns);
        recomputeSession({
          config: recipe.config,
          mapping: mapped,
          varianceParams: {
            granularity: recipe.mapping?.granularity || 'month',
            polarity: recipe.mapping?.polarity || null
          },
          recipeId: recipe.id
        });
      } else {
        recomputeSession({ recipeId: null });
      }

      return sendJson(res, 200, {
        skippedTitleRows: parsed.skippedTitleRows || 0,
        recipes: listRecipes(),
        state: serializeActiveState()
      });
    }

    // 6. Đổi Worksheet trên tệp .XLSX nhiều Sheet vừa tải lên
    if (req.method === 'POST' && pathname === '/api/upload/switch-sheet') {
      const body = await readJsonBody(req);
      if (!session._lastUploadBuffer) {
        return sendJson(res, 400, { error: 'Hiện không có tệp .XLSX đa Sheet nào trong bộ nhớ phiên.' });
      }
      const parsed = parseTabularUpload({
        fileName: session.fileName,
        buffer: session._lastUploadBuffer,
        sheetName: body.sheetName
      });
      session.sheetName = parsed.sheetName;
      session.columns = parsed.columns;
      session.rawRows = parsed.rows;
      recomputeSession({ recipeId: null });
      return sendJson(res, 200, { state: serializeActiveState() });
    }

    // 7. Chạy lại Pipeline Làm sạch + Ánh xạ cột + Giải phẫu Biến động khi DA tinh chỉnh cấu hình
    if (req.method === 'POST' && pathname === '/api/pipeline/run') {
      const body = await readJsonBody(req);
      recomputeSession({
        config: body.config || session.cleanResult?.config,
        mapping: body.mapping || session.cleanResult?.mapping,
        varianceParams: body.varianceParams || {},
        recipeId: body.recipeId !== undefined ? body.recipeId : session.activeRecipeId
      });
      return sendJson(res, 200, { state: serializeActiveState() });
    }

    // 8. Giải phẫu chéo tầng 2 tức thì (Click vào bất kỳ Chi nhánh / Nhân viên / Kênh nào)
    if (req.method === 'POST' && pathname === '/api/variance/drilldown') {
      const body = await readJsonBody(req);
      const vr = session.varianceResult;
      const cr = session.cleanResult;
      if (!vr || !vr.ready) {
        return sendJson(res, 400, { error: 'Chưa sẵn sàng dữ liệu phân tích biến động.' });
      }
      const dateCol = cr.mapping.dateCol;
      const gran = vr.granularity || 'month';
      const rowsP0 = cr.cleanRows.filter((r) => isoDateToPeriod(dateCol ? r[dateCol] : r._period, gran) === vr.period0);
      const rowsP1 = cr.cleanRows.filter((r) => isoDateToPeriod(dateCol ? r[dateCol] : r._period, gran) === vr.period1);

      const deepDive = computeSegmentDeepDive(rowsP0, rowsP1, cr.mapping, body.drillDim, body.drillVal);
      session.varianceResult.deepDive = deepDive;
      return sendJson(res, 200, { deepDive });
    }

    // 9. Thực thi truy vấn SQLite trên Workbench
    if (req.method === 'POST' && pathname === '/api/sql/query') {
      const body = await readJsonBody(req);
      const result = sqlWorkbench.executeQuery(body.sql);
      return sendJson(res, 200, result);
    }

    // 10. Xuất kết quả truy vấn SQL ra .XLSX hoặc .CSV
    if (req.method === 'POST' && pathname === '/api/sql/export') {
      const body = await readJsonBody(req);
      const result = sqlWorkbench.executeQuery(body.sql, 50000);
      const format = body.format === 'csv' ? 'csv' : 'xlsx';

      if (format === 'csv') {
        const csvStr = buildUtf8BomCsv(result.columns, result.rows);
        const buf = Buffer.from(csvStr, 'utf8');
        res.writeHead(200, {
          'Content-Type': 'text/csv; charset=utf-8',
          'Content-Disposition': 'attachment; filename="sql_query_result.csv"',
          'Content-Length': buf.length
        });
        return res.end(buf);
      }

      const xlsxBuf = buildMultiSheetXlsx([
        {
          name: 'SQL_Result',
          columns: result.columns,
          rows: result.rows
        }
      ]);
      res.writeHead(200, {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': 'attachment; filename="sql_query_result.xlsx"',
        'Content-Length': xlsxBuf.length
      });
      return res.end(xlsxBuf);
    }

    // 11. Quản lý Kho Recipe (Danh sách, Lưu mới, Áp dụng 1-Chạm, Xóa)
    if (req.method === 'GET' && pathname === '/api/recipes') {
      return sendJson(res, 200, { recipes: listRecipes() });
    }

    if (req.method === 'POST' && pathname === '/api/recipes') {
      const body = await readJsonBody(req);
      const saved = upsertRecipe({
        id: body.id,
        name: body.name,
        industry: body.industry || session.industry,
        description: body.description,
        expectedColumns: session.columns,
        config: body.config || session.cleanResult?.config,
        mapping: body.mapping || session.cleanResult?.mapping
      });
      session.activeRecipeId = saved.id;
      return sendJson(res, 200, {
        saved,
        recipes: listRecipes()
      });
    }

    if (req.method === 'POST' && pathname === '/api/recipes/apply') {
      const body = await readJsonBody(req);
      const recipe = getRecipeById(body.recipeId);
      if (!recipe) {
        return sendJson(res, 404, { error: 'Không tìm thấy Recipe yêu cầu.' });
      }
      incrementRecipeRun(recipe.id);
      const mapped = sanitizeMappingForColumns(recipe.mapping, session.columns);
      recomputeSession({
        config: recipe.config,
        mapping: mapped || session.cleanResult?.mapping,
        varianceParams: {
          granularity: recipe.mapping?.granularity || 'month',
          polarity: recipe.mapping?.polarity || null
        },
        recipeId: recipe.id
      });
      return sendJson(res, 200, {
        recipes: listRecipes(),
        state: serializeActiveState()
      });
    }

    if (req.method === 'DELETE' && pathname.startsWith('/api/recipes/')) {
      const id = decodeURIComponent(pathname.slice('/api/recipes/'.length));
      deleteRecipe(id);
      if (session.activeRecipeId === id) session.activeRecipeId = null;
      return sendJson(res, 200, { recipes: listRecipes() });
    }

    // 12. Xuất Báo cáo Sạch 4-Sheet (.XLSX) & Bảng Sạch (.CSV UTF-8 BOM)
    if (req.method === 'GET' && pathname === '/api/export/report-xlsx') {
      const xlsxBuf = buildExecutiveWorkbookBuffer();
      const safeSlug = String(session.datasetId || 'report').replace(/[^a-zA-Z0-9_-]/g, '_');
      const outName = `BaoCao_Sach_PVM_${safeSlug}.xlsx`;
      res.writeHead(200, {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': `attachment; filename="${outName}"`,
        'Content-Length': xlsxBuf.length
      });
      return res.end(xlsxBuf);
    }

    if (req.method === 'GET' && pathname === '/api/export/clean-csv') {
      const cr = session.cleanResult;
      const cols = [...session.columns, ...(cr.cleanRows[0]?._period ? ['_period'] : [])];
      const csvStr = buildUtf8BomCsv(cols, cr.cleanRows);
      const buf = Buffer.from(csvStr, 'utf8');
      const safeSlug = String(session.datasetId || 'clean').replace(/[^a-zA-Z0-9_-]/g, '_');
      res.writeHead(200, {
        'Content-Type': 'text/csv; charset=utf-8',
        'Content-Disposition': `attachment; filename="DuLieu_Sach_${safeSlug}.csv"`,
        'Content-Length': buf.length
      });
      return res.end(buf);
    }

    // Phục vụ giao diện Frontend
    if (req.method === 'GET') {
      return serveStaticFile(pathname, res);
    }

    return sendJson(res, 404, { error: 'API Endpoint không tồn tại.' });
  } catch (err) {
    console.error('[Studio Error]', err);
    return sendJson(res, 400, { error: err.message || 'Đã xảy ra lỗi xử lý dữ liệu.' });
  }
});

if (require.main === module) {
  server.listen(PORT, '0.0.0.0', () => {
    console.log(`\n🧼📉 DataScrub & Variance Studio đang chạy tại: http://localhost:${PORT}`);
    console.log(`   • Máy lọc dữ liệu bẩn VN + Price-Volume-Mix Bridge + SQLite Workbench sẵn sàng!\n`);
  });
}

module.exports = { server };
