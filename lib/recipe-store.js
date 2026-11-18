/**
 * ⚡ DataScrub & Variance Studio — Kho Lưu Trữ "Recipe" Quy Trình Làm Sạch 1-Chạm (recipe-store.js)
 *
 * - Lưu trữ cấu hình làm sạch (Date mode, Phone MNP, Fuzzy threshold, Custom Aliases, Dedupe, Outlier strategy)
 *   + ánh xạ cột (Role Mapping) + chiều hướng chỉ số (Polarity) ra tệp `data/recipes.json`
 * - Tuần sau / Tháng sau khi nhận file Excel/CSV mới từ phòng ban, DA chỉ cần chọn Recipe (hoặc để Auto-Match theo tên cột)
 *   -> 3 giây ra ngay Bảng chuẩn + PVM Bridge + Bản giải trình cho sếp!
 */

const fs = require('node:fs');
const path = require('node:path');
const crypto = require('node:crypto');
const { DEFAULT_CLEAN_CONFIG } = require('./viet-cleaner');

const DATA_DIR = path.join(__dirname, '..', 'data');
const RECIPES_FILE = path.join(DATA_DIR, 'recipes.json');

const BUILTIN_RECIPES = [
  {
    id: 'rcp_vn_standard_sales',
    name: '🛒 Chuẩn DA Bán Lẻ & Đa Kênh VN (Auto-PVM)',
    industry: 'Bán lẻ / FMCG / F&B',
    builtin: true,
    description:
      'Tự động cứu ngày Mỹ MM/DD/YYYY lạc dòng, đổi SĐT 11->10 số, gom tên Tỉnh/Thành & Nhân viên (Fuzzy 84%), khử dòng trùng ngầm và thay thế ngoại lai gõ dư số 0 bằng Trung vị mặt hàng.',
    expectedColumns: ['Ngày GD', 'Chi nhánh', 'Nhân viên', 'Kênh bán', 'Sản phẩm', 'Số lượng', 'Doanh thu'],
    config: {
      ...DEFAULT_CLEAN_CONFIG,
      datePreference: 'AUTO_VN',
      fuzzyThreshold: 0.84,
      dedupeStrategy: 'remove_hidden',
      outlierStrategy: 'median_impute',
      customAliases: {
        'Chi nhánh': {
          'Q1 - HCM': 'TP. Hồ Chí Minh',
          'CN Sài Gòn': 'TP. Hồ Chí Minh',
          'CN Cầu Giấy': 'Hà Nội'
        }
      }
    },
    mapping: {
      polarity: 'higher_is_better',
      granularity: 'month'
    },
    createdAt: '2026-10-01T08:00:00.000Z',
    runCount: 24
  },
  {
    id: 'rcp_logistics_freight',
    name: '🚚 Đối Soát Cước Vận Tải & Sản Lượng Tuyến (Logistics)',
    industry: 'Logistics & Chuỗi cung ứng',
    builtin: true,
    description:
      'Chuẩn hóa tên Kho khai thác / Tuyến xe, xử lý cước phí ghi kèm "tr"/"đ", kẹp biên Winsorize các chuyến xe phụ phí đột biến và phân tích PVM theo Tuyến vận chuyển.',
    expectedColumns: ['Ngày vận đơn', 'Trung tâm khai thác', 'Điều phối viên', 'Dịch vụ vận chuyển', 'Số tấn quy đổi', 'Tổng cước thu'],
    config: {
      ...DEFAULT_CLEAN_CONFIG,
      datePreference: 'AUTO_VN',
      fuzzyThreshold: 0.83,
      dedupeStrategy: 'remove_hidden',
      outlierStrategy: 'median_impute',
      customAliases: {}
    },
    mapping: {
      polarity: 'higher_is_better',
      granularity: 'month'
    },
    createdAt: '2026-10-05T09:00:00.000Z',
    runCount: 15
  },
  {
    id: 'rcp_hr_cost_guard',
    name: '👥 Kiểm Toán Biến Động Quỹ Lương & OT (HR / Finance)',
    industry: 'Nhân sự & Tài chính Nội bộ',
    builtin: true,
    description:
      'Chế độ Chi phí (Tăng = Cảnh báo đỏ): Chuẩn hóa tên Khối/Phòng ban & Cấp bậc, bắt bảng công nhập đúp và bóc tách xem Quỹ lương tăng do Tăng giờ công, Đổi cơ cấu nhân sự cấp cao hay Tăng đơn giá OT.',
    expectedColumns: ['Ngày chốt công', 'Văn phòng', 'Quản lý', 'Cấp bậc vị trí', 'Khối phòng ban', 'Số công quy đổi', 'Tổng chi phí lương'],
    config: {
      ...DEFAULT_CLEAN_CONFIG,
      datePreference: 'AUTO_VN',
      fuzzyThreshold: 0.85,
      dedupeStrategy: 'remove_hidden',
      outlierStrategy: 'median_impute',
      customAliases: {}
    },
    mapping: {
      polarity: 'lower_is_better',
      granularity: 'month'
    },
    createdAt: '2026-10-10T10:30:00.000Z',
    runCount: 11
  }
];

function ensureRecipesLoaded() {
  if (!fs.existsSync(DATA_DIR)) {
    fs.mkdirSync(DATA_DIR, { recursive: true });
  }
  if (!fs.existsSync(RECIPES_FILE)) {
    fs.writeFileSync(RECIPES_FILE, JSON.stringify(BUILTIN_RECIPES, null, 2), 'utf8');
    return BUILTIN_RECIPES.slice();
  }
  try {
    const raw = JSON.parse(fs.readFileSync(RECIPES_FILE, 'utf8'));
    if (!Array.isArray(raw)) return BUILTIN_RECIPES.slice();
    // Đảm bảo luôn có các builtin recipes mới nhất
    const existingIds = new Set(raw.map((r) => r.id));
    let merged = raw.slice();
    for (const b of BUILTIN_RECIPES) {
      if (!existingIds.has(b.id)) merged.unshift(b);
    }
    return merged;
  } catch (_) {
    return BUILTIN_RECIPES.slice();
  }
}

function saveAllRecipes(list) {
  if (!fs.existsSync(DATA_DIR)) {
    fs.mkdirSync(DATA_DIR, { recursive: true });
  }
  fs.writeFileSync(RECIPES_FILE, JSON.stringify(list, null, 2), 'utf8');
}

function listRecipes() {
  return ensureRecipesLoaded();
}

function getRecipeById(id) {
  return ensureRecipesLoaded().find((r) => r.id === id) || null;
}

function upsertRecipe(payload) {
  const all = ensureRecipesLoaded();
  const id = payload.id && !payload.id.startsWith('rcp_vn_') && !payload.id.startsWith('rcp_logistics') && !payload.id.startsWith('rcp_hr_')
    ? payload.id
    : `rcp_${crypto.randomBytes(4).toString('hex')}`;

  const existingIdx = all.findIndex((r) => r.id === id && !r.builtin);
  const record = {
    id,
    name: String(payload.name || 'Recipe Làm Sạch Tùy Chỉnh').trim(),
    industry: String(payload.industry || 'Đa ngành').trim(),
    builtin: false,
    description: String(payload.description || 'Quy trình chuẩn hóa dữ liệu & giải phẫu biến động lưu tự động.').trim(),
    expectedColumns: Array.isArray(payload.expectedColumns) ? payload.expectedColumns : [],
    config: {
      ...DEFAULT_CLEAN_CONFIG,
      ...(payload.config || {})
    },
    mapping: payload.mapping || {},
    createdAt: existingIdx >= 0 ? all[existingIdx].createdAt : new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    runCount: existingIdx >= 0 ? (all[existingIdx].runCount || 0) + 1 : 1
  };

  if (existingIdx >= 0) {
    all[existingIdx] = record;
  } else {
    all.unshift(record);
  }
  saveAllRecipes(all);
  return record;
}

function incrementRecipeRun(id) {
  const all = ensureRecipesLoaded();
  const found = all.find((r) => r.id === id);
  if (found) {
    found.runCount = (found.runCount || 0) + 1;
    found.lastRunAt = new Date().toISOString();
    saveAllRecipes(all);
  }
  return found;
}

function deleteRecipe(id) {
  const all = ensureRecipesLoaded();
  const target = all.find((r) => r.id === id);
  if (!target) return false;
  if (target.builtin) {
    throw new Error('Không thể xóa Recipe chuẩn mặc định của hệ thống (bạn có thể nhân bản thành Recipe mới).');
  }
  const filtered = all.filter((r) => r.id !== id);
  saveAllRecipes(filtered);
  return true;
}

module.exports = {
  listRecipes,
  getRecipeById,
  upsertRecipe,
  incrementRecipeRun,
  deleteRecipe
};
