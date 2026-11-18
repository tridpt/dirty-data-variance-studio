/**
 * 🗄️ DataScrub & Variance Studio — SQLite Analytical Workbench (sql-workbench.js)
 *
 * - Sử dụng trực tiếp động cơ C++ SQLite 3 tích hợp trong Node.js 22 (`node:sqlite`)
 * - Tự động nạp 4 bảng phân tích: `clean_data`, `raw_data`, `audit_log`, `pvm_summary`
 * - Hỗ trợ đầy đủ CTE (`WITH ... AS`), Window Functions (`LAG`, `RANK`, `SUM() OVER`), Subquery, `GROUP BY`, `HAVING`
 * - Tự động sinh 6 Mẫu Truy Vấn SQL Thực Chiến (Dynamic Presets) khớp chính xác với tên cột của bộ dữ liệu đang mở!
 */

const { DatabaseSync } = require('node:sqlite');

function quoteIdent(name) {
  return `"${String(name || '').replace(/"/g, '""')}"`;
}

class SqlWorkbench {
  constructor() {
    this.db = null;
    this.schemaInfo = [];
    this.colMap = new Map(); // originalName -> sqlName
    this.presets = [];
  }

  hydrate({ cleanResult, rawRows, varianceResult }) {
    if (this.db) {
      try {
        this.db.close();
      } catch (_) {}
    }

    const db = new DatabaseSync(':memory:');
    this.db = db;
    this.colMap.clear();

    const sqlCols = cleanResult.sqlColumns || [];
    for (const sc of sqlCols) {
      this.colMap.set(sc.original, sc.sqlName);
    }

    // 1. Tạo bảng clean_data
    const cleanColDefs = [
      'row_num INTEGER PRIMARY KEY',
      'period_month TEXT',
      ...sqlCols.map((sc) => `${quoteIdent(sc.sqlName)} ${sc.type === 'numeric' ? 'REAL' : 'TEXT'}`)
    ];
    db.exec(`CREATE TABLE clean_data (${cleanColDefs.join(', ')});`);

    if (cleanResult.cleanRows.length > 0) {
      const placeholders = ['?', '?', ...sqlCols.map(() => '?')].join(', ');
      const insertClean = db.prepare(
        `INSERT INTO clean_data (row_num, period_month, ${sqlCols.map((sc) => quoteIdent(sc.sqlName)).join(', ')}) VALUES (${placeholders})`
      );
      db.exec('BEGIN TRANSACTION;');
      try {
        cleanResult.cleanRows.forEach((row, idx) => {
          const vals = [
            idx + 1,
            row._period || null,
            ...sqlCols.map((sc) => {
              const v = row[sc.original];
              if (v === '' || v === undefined || v === null) return null;
              if (sc.type === 'numeric') {
                const n = Number(v);
                return Number.isFinite(n) ? n : null;
              }
              return String(v);
            })
          ];
          insertClean.run(...vals);
        });
        db.exec('COMMIT;');
      } catch (err) {
        db.exec('ROLLBACK;');
        throw err;
      }
    }

    // 2. Tạo bảng raw_data (dữ liệu bẩn nguyên bản để đối chiếu SQL)
    const rawColDefs = ['row_num INTEGER PRIMARY KEY', ...sqlCols.map((sc) => `${quoteIdent(sc.sqlName)} TEXT`)];
    db.exec(`CREATE TABLE raw_data (${rawColDefs.join(', ')});`);

    if (Array.isArray(rawRows) && rawRows.length > 0) {
      const placeholders = ['?', ...sqlCols.map(() => '?')].join(', ');
      const insertRaw = db.prepare(
        `INSERT INTO raw_data (row_num, ${sqlCols.map((sc) => quoteIdent(sc.sqlName)).join(', ')}) VALUES (${placeholders})`
      );
      db.exec('BEGIN TRANSACTION;');
      try {
        rawRows.forEach((row, idx) => {
          const vals = [
            idx + 1,
            ...sqlCols.map((sc) => {
              const v = row?.[sc.original];
              return v === undefined || v === null ? null : String(v);
            })
          ];
          insertRaw.run(...vals);
        });
        db.exec('COMMIT;');
      } catch (err) {
        db.exec('ROLLBACK;');
        throw err;
      }
    }

    // 3. Tạo bảng audit_log
    db.exec(`
      CREATE TABLE audit_log (
        log_id INTEGER PRIMARY KEY AUTOINCREMENT,
        row_number INTEGER,
        column_name TEXT,
        issue_type TEXT,
        old_value TEXT,
        new_value TEXT,
        detail TEXT
      );
    `);
    if (Array.isArray(cleanResult.auditLogs) && cleanResult.auditLogs.length > 0) {
      const insertAudit = db.prepare(
        `INSERT INTO audit_log (row_number, column_name, issue_type, old_value, new_value, detail) VALUES (?, ?, ?, ?, ?, ?)`
      );
      db.exec('BEGIN TRANSACTION;');
      try {
        for (const l of cleanResult.auditLogs) {
          insertAudit.run(
            l.rowNumber || 0,
            l.column || '',
            l.issueType || '',
            l.oldValue ?? '',
            l.newValue ?? '',
            l.detail ?? ''
          );
        }
        db.exec('COMMIT;');
      } catch (err) {
        db.exec('ROLLBACK;');
        throw err;
      }
    }

    // 4. Tạo bảng pvm_summary
    db.exec(`
      CREATE TABLE pvm_summary (
        item TEXT,
        status TEXT,
        q0 REAL,
        q1 REAL,
        p0 REAL,
        p1 REAL,
        r0 REAL,
        r1 REAL,
        delta_total REAL,
        delta_pct REAL,
        volume_impact REAL,
        mix_impact REAL,
        price_impact REAL,
        portfolio_impact REAL
      );
    `);
    const breakdown = varianceResult?.pvm?.itemBreakdown || [];
    if (breakdown.length > 0) {
      const insertPvm = db.prepare(
        `INSERT INTO pvm_summary VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
      );
      db.exec('BEGIN TRANSACTION;');
      try {
        for (const b of breakdown) {
          insertPvm.run(
            b.item,
            b.status,
            b.q0,
            b.q1,
            b.p0,
            b.p1,
            b.r0,
            b.r1,
            b.deltaTotal,
            b.deltaPct,
            b.volumeImpact,
            b.mixImpact,
            b.priceImpact,
            b.portfolioImpact
          );
        }
        db.exec('COMMIT;');
      } catch (err) {
        db.exec('ROLLBACK;');
        throw err;
      }
    }

    this.schemaInfo = this._introspectTables(sqlCols);
    this.presets = this._buildDynamicPresets(cleanResult.mapping, varianceResult);
    return {
      tables: this.schemaInfo,
      presets: this.presets
    };
  }

  _introspectTables(sqlCols) {
    const labelLookup = new Map(sqlCols.map((sc) => [sc.sqlName, sc.original]));
    const tables = ['clean_data', 'pvm_summary', 'audit_log', 'raw_data'];
    return tables.map((tbl) => {
      const cols = this.db.prepare(`PRAGMA table_info(${quoteIdent(tbl)})`).all();
      const countRow = this.db.prepare(`SELECT COUNT(*) AS cnt FROM ${quoteIdent(tbl)}`).get();
      return {
        tableName: tbl,
        rowCount: countRow?.cnt || 0,
        columns: cols.map((c) => ({
          name: c.name,
          type: c.type || 'TEXT',
          originalLabel: labelLookup.get(c.name) || null
        }))
      };
    });
  }

  _buildDynamicPresets(mapping = {}, varianceResult = null) {
    const sqlOf = (orig) => (orig && this.colMap.get(orig)) || null;
    const metricSql = sqlOf(mapping.metricCol) || 'row_num';
    const qtySql = sqlOf(mapping.quantityCol);
    const itemSql = sqlOf(mapping.itemCol);
    const nonItemDims = (mapping.dimensionCols || []).filter((d) => d !== mapping.itemCol);
    const dim1Sql = sqlOf(nonItemDims[0]) || itemSql || 'period_month';
    const dim2Sql = sqlOf(nonItemDims[1]) || dim1Sql;

    const p0 = varianceResult?.period0 || '2026-09';
    const p1 = varianceResult?.period1 || '2026-10';

    return [
      {
        id: 'mom_window_lag',
        title: '📈 Tăng trưởng qua các Kỳ (Window LAG)',
        description: 'Tính tổng chỉ số, số dòng và % tăng/giảm liên tiếp bằng hàm cửa sổ LAG() OVER',
        sql: `WITH monthly AS (
  SELECT
    period_month,
    COUNT(*) AS so_giao_dich,
    ${qtySql ? `ROUND(SUM(${qtySql}), 2) AS tong_san_luong,` : ''}
    ROUND(SUM(${metricSql}), 0) AS tong_gia_tri
  FROM clean_data
  WHERE period_month IS NOT NULL AND period_month != ''
  GROUP BY period_month
)
SELECT
  period_month AS ky_bao_cao,
  so_giao_dich,
  ${qtySql ? 'tong_san_luong,' : ''}
  tong_gia_tri,
  ROUND(tong_gia_tri - LAG(tong_gia_tri) OVER (ORDER BY period_month), 0) AS chenh_lech_ky_truoc,
  ROUND(
    (tong_gia_tri - LAG(tong_gia_tri) OVER (ORDER BY period_month)) * 100.0
    / NULLIF(LAG(tong_gia_tri) OVER (ORDER BY period_month), 0),
    2
  ) AS tang_truong_pct
FROM monthly
ORDER BY period_month;`
      },
      {
        id: 'branch_variance_bridge',
        title: `🎯 Xếp hạng Biến động theo ${nonItemDims[0] || 'Phân khúc'} (${p0} vs ${p1})`,
        description: 'Chỉ đích danh đơn vị kéo chỉ số đi xuống nhiều nhất giữa 2 kỳ đã chọn',
        sql: `SELECT
  ${dim1Sql} AS don_vi,
  ROUND(SUM(CASE WHEN period_month = '${p0}' THEN ${metricSql} ELSE 0 END), 0) AS ky_goc_${p0.replace(/[^a-zA-Z0-9]/g, '_')},
  ROUND(SUM(CASE WHEN period_month = '${p1}' THEN ${metricSql} ELSE 0 END), 0) AS ky_nay_${p1.replace(/[^a-zA-Z0-9]/g, '_')},
  ROUND(
    SUM(CASE WHEN period_month = '${p1}' THEN ${metricSql} ELSE 0 END) -
    SUM(CASE WHEN period_month = '${p0}' THEN ${metricSql} ELSE 0 END),
    0
  ) AS bien_dong_tuyet_doi,
  ROUND(
    (SUM(CASE WHEN period_month = '${p1}' THEN ${metricSql} ELSE 0 END) -
     SUM(CASE WHEN period_month = '${p0}' THEN ${metricSql} ELSE 0 END)) * 100.0
    / NULLIF(SUM(CASE WHEN period_month = '${p0}' THEN ${metricSql} ELSE 0 END), 0),
    1
  ) AS bien_dong_pct
FROM clean_data
GROUP BY ${dim1Sql}
ORDER BY bien_dong_tuyet_doi ASC;`
      },
      {
        id: 'pareto_80_20',
        title: `🏆 Phân tích Pareto 80/20 (${nonItemDims[1] || nonItemDims[0] || 'Nhân sự / Kênh'})`,
        description: 'Tính tỷ trọng đóng góp và tỷ trọng cộng dồn lũy kế (Cumulative Share %)',
        sql: `WITH agg AS (
  SELECT
    ${dim2Sql} AS doi_tuong,
    COUNT(*) AS so_don,
    ROUND(SUM(${metricSql}), 0) AS tong_gia_tri
  FROM clean_data
  WHERE period_month = '${p1}'
  GROUP BY ${dim2Sql}
)
SELECT
  doi_tuong,
  so_don,
  tong_gia_tri,
  ROUND(tong_gia_tri * 100.0 / SUM(tong_gia_tri) OVER (), 2) AS ty_trong_pct,
  ROUND(
    SUM(tong_gia_tri) OVER (ORDER BY tong_gia_tri DESC ROWS BETWEEN UNBOUNDED PRECEDING AND CURRENT ROW) * 100.0
    / SUM(tong_gia_tri) OVER (),
    2
  ) AS luy_ke_pareto_pct
FROM agg
ORDER BY tong_gia_tri DESC;`
      },
      {
        id: 'pvm_sku_sql',
        title: '🔬 Bảng Giải phẫu Price-Volume-Mix theo Mặt hàng',
        description: 'Truy vấn trực tiếp bảng pvm_summary để lọc các mặt hàng bị âm Cơ cấu (Mix) hoặc Đơn giá (Price)',
        sql: `SELECT
  item AS mat_hang,
  q0 AS sl_ky_goc,
  q1 AS sl_ky_nay,
  ROUND(p0, 0) AS don_gia_goc,
  ROUND(p1, 0) AS don_gia_nay,
  ROUND(volume_impact, 0) AS tac_dong_san_luong,
  ROUND(mix_impact, 0) AS tac_dong_co_cau_mix,
  ROUND(price_impact, 0) AS tac_dong_gia_ban,
  ROUND(delta_total, 0) AS tong_chenh_lech
FROM pvm_summary
ORDER BY delta_total ASC;`
      },
      {
        id: 'audit_dirty_summary',
        title: '🧼 Thống kê Các loại Lỗi Dữ liệu Bẩn đã xử lý',
        description: 'Tổng hợp Nhật ký Làm sạch (audit_log) theo nhóm bệnh dữ liệu & cột bị ảnh hưởng',
        sql: `SELECT
  issue_type AS loai_loi_du_lieu,
  column_name AS cot_bi_anh_huong,
  COUNT(*) AS so_o_da_sua,
  MIN(old_value || ' ➔ ' || new_value) AS vi_du_minh_hoa
FROM audit_log
GROUP BY issue_type, column_name
ORDER BY so_o_da_sua DESC;`
      },
      {
        id: 'raw_vs_clean_fragmentation',
        title: '🆚 So sánh Phân mảnh Danh mục: File Bẩn vs File Sạch',
        description: 'Chứng minh cách dữ liệu bẩn làm vỡ PivotTable nếu không chuẩn hóa Fuzzy Matching',
        sql: `SELECT
  '1. Trước khi lọc (raw_data)' AS trang_thai,
  COUNT(DISTINCT ${dim1Sql}) AS so_ten_phan_manh,
  COUNT(*) AS tong_so_dong
FROM raw_data
UNION ALL
SELECT
  '2. Sau khi làm sạch (clean_data)' AS trang_thai,
  COUNT(DISTINCT ${dim1Sql}) AS so_ten_phan_manh,
  COUNT(*) AS tong_so_dong
FROM clean_data;`
      }
    ];
  }

  executeQuery(sqlText, maxRows = 1500) {
    if (!this.db) {
      throw new Error('Kho dữ liệu SQLite chưa được khởi tạo. Vui lòng nạp bộ dữ liệu trước.');
    }
    const trimmed = String(sqlText || '').trim();
    if (!trimmed) {
      throw new Error('Câu lệnh SQL đang để trống.');
    }

    // Chặn các câu lệnh phá hủy hoặc truy cập tệp ngoài
    const strippedComments = trimmed
      .replace(/--.*$/gm, '')
      .replace(/\/\*[\s\S]*?\*\//g, '')
      .trim();

    if (!/^(SELECT|WITH|PRAGMA|EXPLAIN)\b/i.test(strippedComments)) {
      throw new Error(
        'Workbench ở chế độ Phân tích An toàn (Read-Only): Chỉ cho phép thực thi SELECT, WITH ... SELECT, PRAGMA hoặc EXPLAIN.'
      );
    }
    if (/\b(ATTACH|DETACH|DROP|ALTER|INSERT|UPDATE|DELETE|REPLACE|VACUUM)\b/i.test(strippedComments)) {
      throw new Error('Câu lệnh chứa từ khóa ghi/xóa bị chặn trong môi trường kiểm toán phân tích.');
    }

    const startedAt = performance.now();
    const stmt = this.db.prepare(trimmed);
    const allRows = stmt.all();
    const elapsedMs = Math.round((performance.now() - startedAt) * 100) / 100;

    // Chuyển null-prototype object từ node:sqlite thành plain object JSON-safe
    const plainRows = allRows.slice(0, maxRows).map((r) => ({ ...r }));
    const columns =
      plainRows.length > 0
        ? Object.keys(plainRows[0])
        : typeof stmt.columns === 'function'
        ? stmt.columns().map((c) => c.name)
        : [];

    return {
      columns,
      rows: plainRows,
      totalRowCount: allRows.length,
      truncated: allRows.length > maxRows,
      elapsedMs
    };
  }
}

module.exports = {
  SqlWorkbench
};
