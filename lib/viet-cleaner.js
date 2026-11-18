/**
 * 🧼 DataScrub & Variance Studio — Máy Lọc Dữ Liệu Bẩn Đặc Thù Việt Nam (viet-cleaner.js)
 *
 * Giải quyết trọn bộ 7 căn bệnh kinh điển của file Excel/CSV phòng ban gửi cho DA:
 * 1. Ngày tháng hỗn loạn: dd/mm/yyyy trộn mm/dd/yyyy (tự suy luận ngữ cảnh tháng lân cận), Excel Serial (46308), chữ tiếng Việt
 * 2. Số điện thoại VN: +84, mất số 0 đầu do Excel, dấu chấm/cách, và đổi đầu số 11 số -> 10 số (016x -> 03x, 012x -> 07x/08x)
 * 3. Tiền tệ & số kiểu VN: "12.500.000 đ", "8,5 tr", "1,2 tỷ", "450k", "(1.200.000)", lỗi công thức #DIV/0!, #REF!, #VALUE!
 * 4. Unicode Tổ hợp (NFD macOS/iOS) vs Dựng sẵn (NFC Unikey), ký tự tàng hình Zero-Width & dư khoảng trắng
 * 5. Gom nhóm tên viết sai/viết tắt bằng Từ điển Địa danh/Kinh doanh VN + Thuật toán Fuzzy Clustering có chốt chặn Tên riêng & Size/Mã số
 * 6. Bắt dòng Trùng lặp ngầm (Hidden Semantic Duplicates — khác cách gõ chữ/ngày/SĐT nhưng thực chất là cùng 1 giao dịch)
 * 7. Bắt Giá trị Ngoại lai (Outliers) theo Ngữ cảnh Mặt hàng (kiểm tra kép Implied Unit Rate & MAD/IQR để bắt lỗi gõ dư số 0 mà không phạt nhầm đơn sỉ lớn)
 */

// ============================================================================
// 1. TỪ ĐIỂN CHUẨN HÓA ĐỊA DANH, KÊNH & ĐẦU SỐ VIỄN THÔNG VIỆT NAM
// ============================================================================

const MNP_11_TO_10_PREFIXES = {
  // Viettel
  '0162': '032',
  '0163': '033',
  '0164': '034',
  '0165': '035',
  '0166': '036',
  '0167': '037',
  '0168': '038',
  '0169': '039',
  // MobiFone
  '0120': '070',
  '0121': '079',
  '0122': '077',
  '0126': '076',
  '0128': '078',
  // VinaPhone
  '0123': '083',
  '0124': '084',
  '0125': '085',
  '0127': '081',
  '0129': '082',
  // Vietnamobile & Gmobile
  '0186': '056',
  '0188': '058',
  '0199': '059'
};

const BUILTIN_VN_ALIASES = new Map([
  // TP. Hồ Chí Minh
  ['tphcm', 'TP. Hồ Chí Minh'],
  ['tp hcm', 'TP. Hồ Chí Minh'],
  ['tp ho chi minh', 'TP. Hồ Chí Minh'],
  ['thanh pho ho chi minh', 'TP. Hồ Chí Minh'],
  ['ho chi minh', 'TP. Hồ Chí Minh'],
  ['hcm', 'TP. Hồ Chí Minh'],
  ['sg', 'TP. Hồ Chí Minh'],
  ['sai gon', 'TP. Hồ Chí Minh'],
  ['saigon', 'TP. Hồ Chí Minh'],
  ['tp sai gon', 'TP. Hồ Chí Minh'],
  // Hà Nội
  ['hn', 'Hà Nội'],
  ['ha noi', 'Hà Nội'],
  ['hanoi', 'Hà Nội'],
  ['tp ha noi', 'Hà Nội'],
  ['tp hn', 'Hà Nội'],
  ['thu do ha noi', 'Hà Nội'],
  // Đà Nẵng
  ['dn', 'Đà Nẵng'],
  ['da nang', 'Đà Nẵng'],
  ['danang', 'Đà Nẵng'],
  ['tp da nang', 'Đà Nẵng'],
  ['tp dn', 'Đà Nẵng'],
  // Hải Phòng
  ['hp', 'Hải Phòng'],
  ['hai phong', 'Hải Phòng'],
  ['tp hai phong', 'Hải Phòng'],
  // Cần Thơ
  ['ct', 'Cần Thơ'],
  ['can tho', 'Cần Thơ'],
  ['tp can tho', 'Cần Thơ'],
  // Bình Dương
  ['bd', 'Bình Dương'],
  ['binh duong', 'Bình Dương'],
  ['tinh binh duong', 'Bình Dương'],
  // Bắc Ninh
  ['bn', 'Bắc Ninh'],
  ['bac ninh', 'Bắc Ninh'],
  ['tinh bac ninh', 'Bắc Ninh'],
  // Đồng Nai
  ['dong nai', 'Đồng Nai'],
  ['dnai', 'Đồng Nai'],
  ['bien hoa dong nai', 'Đồng Nai'],
  // Bà Rịa - Vũng Tàu
  ['brvt', 'Bà Rịa - Vũng Tàu'],
  ['br vt', 'Bà Rịa - Vũng Tàu'],
  ['ba ria vung tau', 'Bà Rịa - Vũng Tàu'],
  ['vung tau', 'Bà Rịa - Vũng Tàu'],
  // Khánh Hòa
  ['khanh hoa', 'Khánh Hòa'],
  ['nha trang', 'Khánh Hòa'],
  // Kênh bán / Vận hành phổ biến
  ['b2b dai ly', 'Đại lý B2B'],
  ['dai ly b2b', 'Đại lý B2B'],
  ['kenh dai ly', 'Đại lý B2B'],
  ['dai ly', 'Đại lý B2B'],
  ['gt dai ly', 'Đại lý B2B'],
  ['cua hang truc tiep', 'Cửa hàng Trực tiếp'],
  ['offline store', 'Cửa hàng Trực tiếp'],
  ['pos cua hang', 'Cửa hàng Trực tiếp'],
  ['tai cua hang', 'Cửa hàng Trực tiếp'],
  ['san tmdt', 'Sàn TMĐT'],
  ['tmdt', 'Sàn TMĐT'],
  ['e commerce', 'Sàn TMĐT'],
  ['ecommerce', 'Sàn TMĐT'],
  ['shopee lazada tiktok', 'Sàn TMĐT'],
  ['app mobile', 'Ứng dụng Mobile'],
  ['mobile app', 'Ứng dụng Mobile'],
  ['ung dung di dong', 'Ứng dụng Mobile'],
  ['ung dung mobile', 'Ứng dụng Mobile'],
  ['app', 'Ứng dụng Mobile']
]);

// ============================================================================
// 2. XỬ LÝ VĂN BẢN TIẾNG VIỆT & KHOẢNG CÁCH XÂU (FUZZY SIMILARITY)
// ============================================================================

function cleanInvisibleAndNFC(val) {
  if (val === null || val === undefined) return '';
  return String(val)
    .normalize('NFC')
    .replace(/[\u200B-\u200D\uFEFF\u00A0]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function stripVietAccents(str) {
  return cleanInvisibleAndNFC(str)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/Đ/g, 'D');
}

function toAsciiSkeleton(str) {
  return stripVietAccents(str)
    .toLowerCase()
    .replace(/[._\-/,()[\]]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function toSnakeIdentifier(str, fallbackIdx = 0) {
  const base = stripVietAccents(str)
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '');
  if (!base) return `col_${fallbackIdx + 1}`;
  if (/^\d/.test(base)) return `c_${base}`;
  return base;
}

function hasVietnameseDiacritics(str) {
  return /[àáạảãâầấậẩẫăằắặẳẵèéẹẻẽêềếệểễìíịỉĩòóọỏõôồốộổỗơờớợởỡùúụủũưừứựửữỳýỵỷỹđ]/i.test(String(str || ''));
}

function levenshteinDistance(a, b) {
  if (a === b) return 0;
  if (!a.length) return b.length;
  if (!b.length) return a.length;
  const v0 = new Int32Array(b.length + 1);
  const v1 = new Int32Array(b.length + 1);
  for (let i = 0; i <= b.length; i++) v0[i] = i;
  for (let i = 0; i < a.length; i++) {
    v1[0] = i + 1;
    for (let j = 0; j < b.length; j++) {
      const cost = a[i] === b[j] ? 0 : 1;
      v1[j + 1] = Math.min(v1[j] + 1, v0[j + 1] + 1, v0[j] + cost);
    }
    for (let j = 0; j <= b.length; j++) v0[j] = v1[j];
  }
  return v0[b.length];
}

function jaroWinkler(s1, s2) {
  if (s1 === s2) return 1;
  const len1 = s1.length;
  const len2 = s2.length;
  if (!len1 || !len2) return 0;

  const matchDist = Math.max(0, Math.floor(Math.max(len1, len2) / 2) - 1);
  const s1Matches = new Uint8Array(len1);
  const s2Matches = new Uint8Array(len2);

  let matches = 0;
  for (let i = 0; i < len1; i++) {
    const start = Math.max(0, i - matchDist);
    const end = Math.min(i + matchDist + 1, len2);
    for (let j = start; j < end; j++) {
      if (s2Matches[j] || s1[i] !== s2[j]) continue;
      s1Matches[i] = 1;
      s2Matches[j] = 1;
      matches++;
      break;
    }
  }
  if (!matches) return 0;

  let t = 0;
  let k = 0;
  for (let i = 0; i < len1; i++) {
    if (!s1Matches[i]) continue;
    while (!s2Matches[k]) k++;
    if (s1[i] !== s2[k]) t++;
    k++;
  }
  t /= 2;

  const jaro = (matches / len1 + matches / len2 + (matches - t) / matches) / 3;
  let prefix = 0;
  for (let i = 0; i < Math.min(4, len1, len2); i++) {
    if (s1[i] === s2[i]) prefix++;
    else break;
  }
  return jaro + prefix * 0.1 * (1 - jaro);
}

/**
 * Chốt chặn chống gộp nhầm tên riêng VN ("Nguyễn Văn Hùng" vs "Nguyễn Văn Dũng")
 * và chống gộp nhầm mã số / kích cỡ ("Xe tải 5T" vs "Xe tải 10T", "Size M" vs "Size L")
 */
function computeEntitySimilarity(rawA, rawB, colRole = 'dimension') {
  const skelA = toAsciiSkeleton(rawA);
  const skelB = toAsciiSkeleton(rawB);
  if (!skelA || !skelB) return 0;
  if (skelA === skelB) return 0.995; // Chỉ khác dấu tiếng Việt / hoa thường / dấu câu

  // 1. Chốt chặn chữ số khác nhau (VD: 40ft vs 20ft, SMT-01 vs SMT-02, Quận 1 vs Quận 7)
  const numsA = (skelA.match(/\d+/g) || []).join('|');
  const numsB = (skelB.match(/\d+/g) || []).join('|');
  if (numsA !== numsB && (numsA !== '' || numsB !== '')) {
    return 0.2;
  }

  // 2. Chốt chặn hậu tố Size / Hạng đơn chữ cái ở cuối (VD: Size S vs Size M vs Size L, Hạng A vs Hạng B)
  const tokensA = skelA.split(' ');
  const tokensB = skelB.split(' ');
  const lastA = tokensA[tokensA.length - 1];
  const lastB = tokensB[tokensB.length - 1];
  if (lastA.length === 1 && lastB.length === 1 && lastA !== lastB) {
    return 0.25;
  }

  // 3. Chốt chặn Họ Tên người Việt (2-4 âm tiết): Nếu cùng Họ + Đệm mà khác Tên chính -> KHÔNG gộp!
  const looksLikePersonCol = /nhan_vien|nv|sale|rm|phu_trach|quan_ly|truong_ca|dieu_phoi|staff|rep|employee/i.test(
    colRole
  );
  if (looksLikePersonCol || (tokensA.length >= 3 && tokensB.length >= 3 && tokensA.length <= 4 && tokensB.length <= 4)) {
    const commonVietSurnames = new Set([
      'nguyen', 'tran', 'le', 'pham', 'hoang', 'huynh', 'phan', 'vu', 'vo', 'dang', 'bui', 'do', 'ho', 'ngo', 'duong', 'ly', 'trinh', 'dinh', 'lam', 'mai', 'luong', 'cao', 'thai', 'ta'
    ]);
    if (commonVietSurnames.has(tokensA[0]) && commonVietSurnames.has(tokensB[0])) {
      if (lastA !== lastB) {
        const distLast = levenshteinDistance(lastA, lastB);
        // "hung" vs "dung", "minh" vs "linh", "anh" vs "thanh" -> khác tên thật, cấm gộp!
        if (!(Math.min(lastA.length, lastB.length) >= 5 && distLast === 1)) {
          return 0.35;
        }
      }
      // Cùng Tên cuối nhưng khác Họ đầu ("Nguyễn Minh Khoa" vs "Trần Minh Khoa") -> cấm gộp!
      if (tokensA[0] !== tokensB[0] && levenshteinDistance(tokensA[0], tokensB[0]) > 1) {
        return 0.35;
      }
    }
  }

  // 4. Chốt chặn từ khóa địa lý đối lập (Bắc Ninh vs Bắc Giang, Hà Nội vs Hà Nam)
  const geoPairs = [
    ['ninh', 'giang'],
    ['noi', 'nam'],
    ['tinh', 'nam'],
    ['thuan', 'phuoc'],
    ['duong', 'phuoc'],
    ['cai', 'son']
  ];
  for (const [g1, g2] of geoPairs) {
    if ((lastA === g1 && lastB === g2) || (lastA === g2 && lastB === g1)) {
      return 0.3;
    }
  }

  // 5. Kết hợp Jaro-Winkler + Normalized Levenshtein + Token-Sort
  const maxLen = Math.max(skelA.length, skelB.length);
  const levSim = 1 - levenshteinDistance(skelA, skelB) / maxLen;
  const jwSim = jaroWinkler(skelA, skelB);

  const sortedA = tokensA.slice().sort().join(' ');
  const sortedB = tokensB.slice().sort().join(' ');
  const tokenSortSim = 1 - levenshteinDistance(sortedA, sortedB) / Math.max(sortedA.length, sortedB.length);

  // Kiểm tra trường hợp viết tắt tiền tố như "NV Khoa - Trần Minh Khoa" vs "Trần Minh Khoa"
  let containmentBonus = 0;
  if (tokensA.length !== tokensB.length) {
    const shorter = tokensA.length < tokensB.length ? skelA : skelB;
    const longer = tokensA.length < tokensB.length ? skelB : skelA;
    if (shorter.length >= 8 && longer.endsWith(shorter)) {
      containmentBonus = 0.14;
    }
  }

  return Math.min(0.99, Math.max(levSim * 0.45 + jwSim * 0.35 + tokenSortSim * 0.2, tokenSortSim * 0.85) + containmentBonus);
}

// ============================================================================
// 3. BỘ PHÂN TÍCH NGÀY THÁNG ĐA ĐỊNH DẠNG + SUY LUẬN NGỮ CẢNH THÁNG
// ============================================================================

function excelSerialToIso(serial) {
  const num = Number(serial);
  if (!Number.isFinite(num) || num < 30000 || num > 65000) return null; // Khoảng năm 1982 -> 2077
  const utcDays = Math.floor(num > 59 ? num - 1 : num);
  const d = new Date(Date.UTC(1899, 11, 31) + utcDays * 86400000);
  if (Number.isNaN(d.getTime())) return null;
  return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}-${String(d.getUTCDate()).padStart(2, '0')}`;
}

function isValidYMD(y, m, d) {
  if (y < 1970 || y > 2100 || m < 1 || m > 12 || d < 1 || d > 31) return false;
  const dt = new Date(Date.UTC(y, m - 1, d));
  return dt.getUTCFullYear() === y && dt.getUTCMonth() === m - 1 && dt.getUTCDate() === d;
}

function formatYMD(y, m, d) {
  return `${String(y).padStart(4, '0')}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
}

const EN_MONTHS = {
  jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6,
  jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12
};

/**
 * Phân tích sơ bộ từng ô ngày tháng (trả về unambiguous hoặc ambiguous { a, b, year })
 */
function inspectRawDateToken(raw) {
  const s = cleanInvisibleAndNFC(raw);
  if (!s) return { kind: 'empty' };

  // 1. Số Excel Serial (VD: 46305 hoặc "46305")
  if (/^\d{5}(\.\d+)?$/.test(s)) {
    const iso = excelSerialToIso(Number(s));
    if (iso) return { kind: 'unambiguous', iso, reason: 'excel_serial' };
  }

  // 2. Định dạng ISO chuẩn hoặc YYYY/MM/DD, YYYY.MM.DD
  let m = s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})(?:[T\s].*)?$/);
  if (m) {
    const y = Number(m[1]);
    const p1 = Number(m[2]);
    const p2 = Number(m[3]);
    if (isValidYMD(y, p1, p2)) {
      const normalized = formatYMD(y, p1, p2);
      return {
        kind: 'unambiguous',
        iso: normalized,
        reason: s === normalized ? 'clean_iso' : 'iso_variant'
      };
    }
    if (p1 > 12 && isValidYMD(y, p2, p1)) {
      return { kind: 'unambiguous', iso: formatYMD(y, p2, p1), reason: 'yyyy_dd_mm_swapped' };
    }
  }

  // 3. Tiếng Việt: "Ngày 12 tháng 10 năm 2026" hoặc "12 thg 10, 2026" / "12-T10-2026"
  const lower = stripVietAccents(s).toLowerCase();
  m = lower.match(/(?:ngay\s*)?(\d{1,2})\s*(?:thang|thg|t|-t)\s*(\d{1,2})[\s,/-]*(?:nam\s*)?(\d{2,4})/);
  if (m) {
    const d = Number(m[1]);
    const mo = Number(m[2]);
    let y = Number(m[3]);
    if (y < 100) y += 2000;
    if (isValidYMD(y, mo, d)) {
      return { kind: 'unambiguous', iso: formatYMD(y, mo, d), reason: 'vn_text_date' };
    }
  }

  // 4. Tháng chữ tiếng Anh: "15-Oct-2026" hoặc "Oct 15, 2026"
  m = lower.match(/^(\d{1,2})[\s\-/]+([a-z]{3,9})[\s\-/,]+(\d{2,4})$/);
  if (m && EN_MONTHS[m[2].slice(0, 3)]) {
    const d = Number(m[1]);
    const mo = EN_MONTHS[m[2].slice(0, 3)];
    let y = Number(m[3]);
    if (y < 100) y += 2000;
    if (isValidYMD(y, mo, d)) {
      return { kind: 'unambiguous', iso: formatYMD(y, mo, d), reason: 'en_text_date' };
    }
  }
  m = lower.match(/^([a-z]{3,9})[\s\-/]+(\d{1,2})[\s\-/,]+(\d{2,4})$/);
  if (m && EN_MONTHS[m[1].slice(0, 3)]) {
    const mo = EN_MONTHS[m[1].slice(0, 3)];
    const d = Number(m[2]);
    let y = Number(m[3]);
    if (y < 100) y += 2000;
    if (isValidYMD(y, mo, d)) {
      return { kind: 'unambiguous', iso: formatYMD(y, mo, d), reason: 'en_text_date' };
    }
  }

  // 5. Dạng A/B/YYYY hoặc A-B-YY (ổ bệnh lớn nhất: dd/mm/yyyy trộn mm/dd/yyyy)
  m = s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})(?:\s+.*)?$/);
  if (m) {
    const a = Number(m[1]);
    const b = Number(m[2]);
    let y = Number(m[3]);
    if (y < 100) y += 2000;

    if (a > 12 && b >= 1 && b <= 12 && isValidYMD(y, b, a)) {
      return { kind: 'unambiguous', iso: formatYMD(y, b, a), reason: 'dd_mm_yyyy' };
    }
    if (b > 12 && a >= 1 && a <= 12 && isValidYMD(y, a, b)) {
      return { kind: 'unambiguous', iso: formatYMD(y, a, b), reason: 'us_mm_dd_yyyy' };
    }
    if (a === b && isValidYMD(y, b, a)) {
      return { kind: 'unambiguous', iso: formatYMD(y, b, a), reason: 'symmetric_d_m' };
    }
    if (a >= 1 && a <= 12 && b >= 1 && b <= 12) {
      return {
        kind: 'ambiguous',
        a,
        b,
        y,
        sep: s.includes('-') ? '-' : s.includes('.') ? '.' : '/',
        ddmmIso: isValidYMD(y, b, a) ? formatYMD(y, b, a) : null,
        mmddIso: isValidYMD(y, a, b) ? formatYMD(y, a, b) : null
      };
    }
  }

  return { kind: 'invalid', raw: s };
}

/**
 * Giải quyết toàn cục cả cột ngày tháng bằng Thống kê Tháng + Dòng lân cận
 */
function disambiguateDateColumn(rawValues, datePreference = 'AUTO_VN') {
  const inspected = rawValues.map(inspectRawDateToken);
  const monthFreq = new Map();
  let ddmmCount = 0;
  let mmddCount = 0;

  for (const item of inspected) {
    if (item.kind === 'unambiguous' && item.iso) {
      const ym = item.iso.slice(0, 7);
      monthFreq.set(ym, (monthFreq.get(ym) || 0) + 1);
      if (item.reason === 'dd_mm_yyyy') ddmmCount++;
      if (item.reason === 'us_mm_dd_yyyy') mmddCount++;
    }
  }

  return inspected.map((item, idx) => {
    if (item.kind === 'empty') return { iso: '', changed: false, issueType: 'missing_date' };
    if (item.kind === 'invalid') return { iso: item.raw, changed: false, issueType: 'invalid_date' };
    if (item.kind === 'unambiguous') {
      const rawTrim = cleanInvisibleAndNFC(rawValues[idx]);
      return {
        iso: item.iso,
        changed: rawTrim !== item.iso,
        issueType: item.reason === 'clean_iso' ? null : item.reason
      };
    }

    // Xử lý ô nhập nhằng (cả a <= 12 và b <= 12, a != b)
    if (datePreference === 'US_MMDD' && item.mmddIso) {
      return { iso: item.mmddIso, changed: true, issueType: 'us_mm_dd_yyyy' };
    }
    if (datePreference === 'STRICT_DDMM' && item.ddmmIso) {
      return { iso: item.ddmmIso, changed: true, issueType: 'dd_mm_yyyy' };
    }

    // Chế độ AUTO_VN thông minh:
    // Mặc định ưu tiên chuẩn Việt Nam DD/MM/YYYY, NHƯNG nếu tháng theo DD/MM hoàn toàn không tồn tại
    // trong báo cáo mà tháng theo MM/DD lại khớp với kỳ báo cáo & dòng lân cận -> phát hiện dán nhầm chuẩn Mỹ!
    const ymDdmm = item.ddmmIso ? item.ddmmIso.slice(0, 7) : '';
    const ymMmdd = item.mmddIso ? item.mmddIso.slice(0, 7) : '';
    const freqDdmm = monthFreq.get(ymDdmm) || 0;
    const freqMmdd = monthFreq.get(ymMmdd) || 0;

    let neighborDdmm = 0;
    let neighborMmdd = 0;
    for (let delta = -3; delta <= 3; delta++) {
      if (delta === 0) continue;
      const nb = inspected[idx + delta];
      if (nb && nb.kind === 'unambiguous' && nb.iso) {
        const nbYm = nb.iso.slice(0, 7);
        if (nbYm === ymDdmm) neighborDdmm++;
        if (nbYm === ymMmdd) neighborMmdd++;
      }
    }

    const shouldFlipToUs =
      item.mmddIso &&
      ((freqDdmm === 0 && freqMmdd >= 3 && (mmddCount > 0 || neighborMmdd >= 1 || item.sep === '-')) ||
        (mmddCount > ddmmCount * 2 && freqMmdd > freqDdmm));

    if (shouldFlipToUs) {
      return {
        iso: item.mmddIso,
        changed: true,
        issueType: 'us_mm_dd_context_rescued'
      };
    }

    return {
      iso: item.ddmmIso || item.mmddIso || cleanInvisibleAndNFC(rawValues[idx]),
      changed: true,
      issueType: 'dd_mm_yyyy'
    };
  });
}

// ============================================================================
// 4. CHUẨN HÓA SỐ ĐIỆN THOẠI VIỆT NAM & TIỀN TỆ / SỐ LIỆU
// ============================================================================

function normalizeVietnamPhone(raw) {
  const s = cleanInvisibleAndNFC(raw);
  if (!s) return { value: '', changed: false, issueType: null, valid: true };

  // Bỏ dấu nháy đơn đầu của Excel ('0909...) và các ký tự phân cách
  let digits = s.replace(/^'+/, '').replace(/[^\d+]/g, '');
  let issueType = null;

  if (digits.startsWith('+84')) {
    digits = '0' + digits.slice(3);
    issueType = 'phone_country_code';
  } else if (digits.startsWith('0084')) {
    digits = '0' + digits.slice(4);
    issueType = 'phone_country_code';
  } else if (digits.startsWith('84') && (digits.length === 11 || digits.length === 12)) {
    digits = '0' + digits.slice(2);
    issueType = 'phone_country_code';
  } else if (digits.length === 9 && /^[35789]/.test(digits)) {
    // Mất số 0 ở đầu do định dạng cột General/Number trong Excel
    digits = '0' + digits;
    issueType = 'phone_missing_zero';
  } else if (digits.length === 10 && /^1[2689]/.test(digits)) {
    digits = '0' + digits;
  }

  // Chuyển đổi đầu số 11 số cũ sang 10 số theo quy hoạch Bộ TT&TT
  if (digits.length === 11 && digits.startsWith('01')) {
    const prefix4 = digits.slice(0, 4);
    if (MNP_11_TO_10_PREFIXES[prefix4]) {
      digits = MNP_11_TO_10_PREFIXES[prefix4] + digits.slice(4);
      issueType = 'phone_11_to_10_mnp';
    }
  }

  const validMobileOrLandline =
    (digits.length === 10 && /^0[235789]\d{8}$/.test(digits) && !/^0(\d)\1{8}$/.test(digits)) ||
    (digits.length === 11 && /^02\d{9}$/.test(digits));

  if (!validMobileOrLandline) {
    return {
      value: digits || s,
      changed: digits !== s,
      issueType: 'phone_invalid',
      valid: false
    };
  }

  if (!issueType && digits !== s) {
    issueType = 'phone_format_punct';
  }

  return {
    value: digits,
    changed: digits !== s,
    issueType,
    valid: true
  };
}

function parseVietnameseNumber(raw) {
  if (typeof raw === 'number') {
    if (!Number.isFinite(raw)) return { num: null, changed: false, issueType: 'excel_error' };
    return { num: raw, changed: false, issueType: null };
  }
  const s = cleanInvisibleAndNFC(raw);
  if (!s || /^(-|--|n\/a|na|null|none|trống|chưa có)$/i.test(s)) {
    return { num: null, changed: s !== '', issueType: s ? 'null_placeholder' : 'missing_num' };
  }
  if (/^#(DIV\/0!|REF!|VALUE!|NAME\?|N\/A|NUM!)/i.test(s)) {
    return { num: null, changed: true, issueType: 'excel_formula_error' };
  }

  let working = s.toLowerCase();
  let negative = false;
  if (/^\(.*\)$/.test(working)) {
    negative = true;
    working = working.slice(1, -1).trim();
  }
  if (working.startsWith('-')) {
    negative = true;
    working = working.slice(1).trim();
  }

  // Bỏ ký hiệu tiền tệ VNĐ, đ, VND, USD trước khi xét hậu tố tỷ / triệu / nghìn
  working = working.replace(/(vnđ|vnd|đồng|dong|₫|\$|usd)/gi, ' ').trim();

  let multiplier = 1;
  if (/(?:^|[\d\s.,])(tỷ|ty|tỉ|ti|bn)(?=$|[\s.,đ])/i.test(working)) {
    multiplier = 1_000_000_000;
    working = working.replace(/(tỷ|ty|tỉ|ti|bn)/gi, '');
  } else if (/(?:^|[\d\s.,])(triệu|trieu|tr|m)(?=$|[\s.,đ])/i.test(working)) {
    multiplier = 1_000_000;
    working = working.replace(/(triệu|trieu|tr|m)/gi, '');
  } else if (/(?:^|[\d\s.,])(nghìn|nghin|ngàn|ngan|k)(?=$|[\s.,đ])/i.test(working)) {
    multiplier = 1_000;
    working = working.replace(/(nghìn|nghin|ngàn|ngan|k)/gi, '');
  }

  working = working.replace(/đ/gi, '').replace(/\s+/g, '');

  if (!working || !/^[\d.,]+$/.test(working) || !/\d/.test(working)) {
    return { num: null, changed: false, issueType: 'unparseable_num' };
  }

  const dotCount = (working.match(/\./g) || []).length;
  const commaCount = (working.match(/,/g) || []).length;

  let normalizedDigits = working;
  if (dotCount > 0 && commaCount > 0) {
    // Cả chấm lẫn phẩy: dấu nào đứng cuối cùng là dấu thập phân
    if (working.lastIndexOf(',') > working.lastIndexOf('.')) {
      // VD: 12.500.000,50 -> chuẩn VN
      normalizedDigits = working.replace(/\./g, '').replace(',', '.');
    } else {
      // VD: 12,500,000.50 -> chuẩn US
      normalizedDigits = working.replace(/,/g, '');
    }
  } else if (dotCount > 1) {
    // Nhiều dấu chấm -> chắc chắn là phân cách hàng nghìn VN (12.500.000)
    normalizedDigits = working.replace(/\./g, '');
  } else if (commaCount > 1) {
    // Nhiều dấu phẩy -> phân cách hàng nghìn US (12,500,000)
    normalizedDigits = working.replace(/,/g, '');
  } else if (dotCount === 1 && commaCount === 0) {
    const [left, right] = working.split('.');
    // Nếu có đúng 3 chữ số sau dấu chấm và không có hậu tố triệu/tỷ (VD: "450.000" hoặc "12.500") -> hàng nghìn VN!
    if (multiplier === 1 && right.length === 3 && left.length >= 1 && left !== '0') {
      normalizedDigits = left + right;
    }
  } else if (commaCount === 1 && dotCount === 0) {
    const [left, right] = working.split(',');
    if (multiplier === 1 && right.length === 3 && left.length >= 1 && left !== '0') {
      normalizedDigits = left + right;
    } else {
      normalizedDigits = `${left}.${right}`;
    }
  }

  const parsed = Number(normalizedDigits);
  if (!Number.isFinite(parsed)) {
    return { num: null, changed: false, issueType: 'unparseable_num' };
  }

  let finalNum = parsed * multiplier;
  if (negative) finalNum = -finalNum;
  // Làm tròn tránh sai số dấu phẩy động lẻ .000000001
  finalNum = Math.round(finalNum * 10000) / 10000;

  const isFormattedDirty = !/^-?\d+(\.\d+)?$/.test(s);
  return {
    num: finalNum,
    changed: isFormattedDirty,
    issueType: isFormattedDirty ? 'currency_number_format' : null
  };
}

// ============================================================================
// 5. NHẬN DIỆN VAI TRÒ CỘT TỰ ĐỘNG (SCHEMA & ROLE INFERENCE)
// ============================================================================

function inferColumnSchema(columns, rows) {
  const sampleRows = rows.slice(0, Math.min(rows.length, 250));
  const profiles = [];

  for (let idx = 0; idx < columns.length; idx++) {
    const col = columns[idx];
    const snake = toSnakeIdentifier(col, idx);
    const nonEmptyVals = [];
    let dateHits = 0;
    let phoneHits = 0;
    let numHits = 0;

    for (const r of sampleRows) {
      const raw = r?.[col];
      const str = cleanInvisibleAndNFC(raw);
      if (!str) continue;
      nonEmptyVals.push(raw);

      const dt = inspectRawDateToken(str);
      if (dt.kind === 'unambiguous' || dt.kind === 'ambiguous') dateHits++;

      const digitsOnly = str.replace(/[^\d+]/g, '');
      if (digitsOnly.length >= 9 && digitsOnly.length <= 13 && /^(\+?84|0084|0|[35789])/.test(digitsOnly)) {
        phoneHits++;
      }

      const np = parseVietnameseNumber(raw);
      if (np.num !== null) numHits++;
    }

    const totalNonEmpty = nonEmptyVals.length || 1;
    const distinctSet = new Set(nonEmptyVals.map((v) => cleanInvisibleAndNFC(v).toLowerCase()));
    const distinctRatio = distinctSet.size / totalNonEmpty;

    const headerKey = stripVietAccents(col).toLowerCase();
    let detectedType = 'categorical';

    const isPhoneHeader = /\b(sdt|dien thoai|so dien thoai|phone|mobile|tel|lien he)\b/.test(headerKey);
    const isDateHeader = /\b(ngay|date|thoi gian|created|ky chot|ngay ct|disbursement|shipment)\b/.test(headerKey);
    const isIdHeader = /\b(ma don|ma gd|ma van don|ma ho so|ma lo|ma bang|order id|txn id|id|invoice|so hd|code|ma kh)\b/.test(
      headerKey
    );

    if (isPhoneHeader || (phoneHits / totalNonEmpty >= 0.65 && !isIdHeader)) {
      detectedType = 'phone';
    } else if (isDateHeader || (dateHits / totalNonEmpty >= 0.65 && numHits / totalNonEmpty < 0.85)) {
      detectedType = 'date';
    } else if (!isIdHeader && numHits / totalNonEmpty >= 0.72) {
      detectedType = 'numeric';
    } else if (isIdHeader || (distinctRatio > 0.88 && sampleRows.length >= 20)) {
      detectedType = 'id';
    } else {
      detectedType = 'categorical';
    }

    profiles.push({
      name: col,
      sqlName: snake,
      detectedType,
      nonEmptyCount: nonEmptyVals.length,
      missingCount: rows.length - nonEmptyVals.length,
      distinctCount: distinctSet.size,
      headerKey
    });
  }

  // Đảm bảo sqlName duy nhất
  const usedSql = new Map();
  for (const p of profiles) {
    if (usedSql.has(p.sqlName)) {
      const n = usedSql.get(p.sqlName) + 1;
      usedSql.set(p.sqlName, n);
      p.sqlName = `${p.sqlName}_${n}`;
    } else {
      usedSql.set(p.sqlName, 1);
    }
  }

  // Tự động gán Role phục vụ PVM Bridge & Key-Driver
  const pickCol = (filterFn) => profiles.find(filterFn)?.name || null;

  const dateCol = pickCol((p) => p.detectedType === 'date');
  const phoneCol = pickCol((p) => p.detectedType === 'phone');
  const idCol = pickCol((p) => p.detectedType === 'id');

  const numericCols = profiles.filter((p) => p.detectedType === 'numeric');
  const metricCol =
    numericCols.find((p) =>
      /(doanh thu|thanh tien|tong tien|tong cuoc|thu nhap|gia tri|chi phi|quy luong|doanh so|revenue|amount|total|sales|net)/.test(
        p.headerKey
      )
    )?.name ||
    numericCols[numericCols.length - 1]?.name ||
    null;

  const quantityCol =
    numericCols.find(
      (p) =>
        p.name !== metricCol &&
        /(so luong|san luong|so tan|so chuyen|so khe uoc|so cong|gio quy doi|qty|quantity|volume|units|headcount)/.test(
          p.headerKey
        )
    )?.name || null;

  const unitPriceCol =
    numericCols.find(
      (p) => p.name !== metricCol && p.name !== quantityCol && /(don gia|gia ban|unit price|price|rate|muc cuoc)/.test(p.headerKey)
    )?.name || null;

  const catCols = profiles.filter((p) => p.detectedType === 'categorical');
  const itemCol =
    catCols.find((p) =>
      /(san pham|mat hang|ten hang|sku|product|item|dich vu|tuyen|goi tin dung|linh kien|cap bac|vi tri|nhom hang)/.test(
        p.headerKey
      )
    )?.name ||
    catCols[0]?.name ||
    null;

  const dimensionCols = catCols.map((p) => p.name);

  // Tự động nhận diện chiều hướng tốt/xấu của chỉ số (Chi phí / Lỗi / Hao hụt -> thấp là tốt)
  const metricHeaderKey = profiles.find((p) => p.name === metricCol)?.headerKey || '';
  const defaultPolarity = /(chi phi|quy luong|hao hut|ton that|cost|expense|churn|defect)/.test(metricHeaderKey)
    ? 'lower_is_better'
    : 'higher_is_better';

  return {
    profiles,
    suggestedMapping: {
      dateCol,
      idCol,
      phoneCol,
      metricCol,
      quantityCol,
      unitPriceCol,
      itemCol,
      dimensionCols,
      polarity: defaultPolarity
    }
  };
}

// ============================================================================
// 6. GOM CỤM FUZZY MATCHING CHO CỘT DANH MỤC (CHI NHÁNH / NHÂN VIÊN / SẢN PHẨM)
// ============================================================================

function toSmartDisplayCase(str) {
  const clean = cleanInvisibleAndNFC(str);
  if (!clean) return '';
  // Giữ nguyên các từ viết tắt in hoa hợp lệ như B2B, TMĐT, AI, CNC, BMS, SME, TP.
  return clean
    .split(' ')
    .map((word) => {
      if (/^(TP\.|B2B|B2C|TMĐT|VIP|SME|AI|CNC|BMS|QC|OT|POS|RM|KCN|NM|CN)$/i.test(word)) {
        return word.toUpperCase() === 'TP.' ? 'TP.' : word.toUpperCase();
      }
      if (/[0-9]/.test(word) && word === word.toUpperCase()) return word;
      return word.charAt(0).toUpperCase() + word.slice(1);
    })
    .join(' ');
}

function buildFuzzyMappingForColumn(values, { threshold = 0.84, customAliases = {}, colName = '' } = {}) {
  const freqMap = new Map();
  for (const raw of values) {
    const clean = cleanInvisibleAndNFC(raw);
    if (!clean) continue;
    freqMap.set(clean, (freqMap.get(clean) || 0) + 1);
  }

  // Giai đoạn 1: Áp dụng customAliases của người dùng + Từ điển chuẩn Việt Nam
  const directCanonicalMap = new Map();
  const normalizedCustom = new Map();
  for (const [k, v] of Object.entries(customAliases || {})) {
    if (k && v) normalizedCustom.set(toAsciiSkeleton(k), cleanInvisibleAndNFC(v));
  }

  const distinctVariants = Array.from(freqMap.entries())
    .map(([label, count]) => ({ label, count, skel: toAsciiSkeleton(label) }))
    .sort((a, b) => {
      // Ưu tiên bản có dấu tiếng Việt đầy đủ và tần suất xuất hiện cao làm đại diện chuẩn (Canonical)
      const scoreA = a.count * 2 + (hasVietnameseDiacritics(a.label) ? 8 : 0) + (a.label === toSmartDisplayCase(a.label) ? 2 : 0);
      const scoreB = b.count * 2 + (hasVietnameseDiacritics(b.label) ? 8 : 0) + (b.label === toSmartDisplayCase(b.label) ? 2 : 0);
      return scoreB - scoreA;
    });

  // Gán các từ viết tắt địa danh / kênh bán chuẩn trước
  const clusters = []; // [{ canonical, variants: Map<rawVariant, count>, reason }]
  const findOrCreateCluster = (canonical, reason = 'dictionary') => {
    let found = clusters.find((c) => toAsciiSkeleton(c.canonical) === toAsciiSkeleton(canonical));
    if (!found) {
      found = { canonical, variants: new Map(), reason };
      clusters.push(found);
    }
    return found;
  };

  const remaining = [];
  for (const item of distinctVariants) {
    if (normalizedCustom.has(item.skel)) {
      const target = normalizedCustom.get(item.skel);
      const cl = findOrCreateCluster(target, 'custom_recipe');
      cl.variants.set(item.label, item.count);
      directCanonicalMap.set(item.label, cl.canonical);
    } else if (BUILTIN_VN_ALIASES.has(item.skel)) {
      const target = BUILTIN_VN_ALIASES.get(item.skel);
      const cl = findOrCreateCluster(target, 'vn_alias_dict');
      cl.variants.set(item.label, item.count);
      directCanonicalMap.set(item.label, cl.canonical);
    } else {
      remaining.push(item);
    }
  }

  // Giai đoạn 2: Gom cụm Fuzzy đối với các tên còn lại (Nhân viên, Sản phẩm, Chi nhánh đặc thù...)
  for (const item of remaining) {
    let bestCluster = null;
    let bestSim = 0;

    for (const cl of clusters) {
      const sim = computeEntitySimilarity(item.label, cl.canonical, colName);
      if (sim >= threshold && sim > bestSim) {
        bestSim = sim;
        bestCluster = cl;
      }
    }

    if (bestCluster) {
      bestCluster.variants.set(item.label, item.count);
      if (bestCluster.reason === 'singleton') bestCluster.reason = 'fuzzy_match';
      directCanonicalMap.set(item.label, bestCluster.canonical);
    } else {
      const cleanCanonical = toSmartDisplayCase(item.label);
      const newCl = {
        canonical: cleanCanonical,
        variants: new Map([[item.label, item.count]]),
        reason: 'singleton'
      };
      clusters.push(newCl);
      directCanonicalMap.set(item.label, cleanCanonical);
    }
  }

  const activeClusters = clusters
    .filter((c) => {
      if (c.variants.size > 1) return true;
      const [onlyVar] = c.variants.keys();
      return onlyVar && onlyVar !== c.canonical;
    })
    .map((c) => {
      const variantList = Array.from(c.variants.entries()).map(([name, count]) => ({ name, count }));
      const totalRows = variantList.reduce((acc, v) => acc + v.count, 0);
      const changedRows = variantList.filter((v) => v.name !== c.canonical).reduce((acc, v) => acc + v.count, 0);
      return {
        column: colName,
        canonical: c.canonical,
        reason: c.reason,
        totalRows,
        changedRows,
        variants: variantList
      };
    })
    .filter((c) => c.changedRows > 0)
    .sort((a, b) => b.changedRows - a.changedRows);

  return {
    mapValue: (raw) => {
      const trimmed = cleanInvisibleAndNFC(raw);
      if (!trimmed) return { value: '', changed: raw !== '', issueType: null };
      if (normalizedCustom.has(toAsciiSkeleton(trimmed))) {
        const target = normalizedCustom.get(toAsciiSkeleton(trimmed));
        return {
          value: target,
          changed: String(raw) !== target,
          issueType: trimmed !== target ? 'fuzzy_standardized' : String(raw) !== target ? 'whitespace_unicode' : null
        };
      }
      const mapped = directCanonicalMap.get(trimmed) || trimmed;
      let issueType = null;
      if (trimmed !== mapped) {
        issueType = 'fuzzy_standardized';
      } else if (String(raw) !== mapped) {
        issueType = 'whitespace_unicode';
      }
      return {
        value: mapped,
        changed: String(raw) !== mapped,
        issueType
      };
    },
    clusters: activeClusters
  };
}

// ============================================================================
// 7. PHÁT HIỆN & XỬ LÝ NGOẠI LAI THEO NGỮ CẢNH (CONTEXT-AWARE MAD/IQR OUTLIERS)
// ============================================================================

function computeQuantile(sortedArr, q) {
  if (!sortedArr.length) return 0;
  const pos = (sortedArr.length - 1) * q;
  const base = Math.floor(pos);
  const rest = pos - base;
  if (sortedArr[base + 1] !== undefined) {
    return sortedArr[base] + rest * (sortedArr[base + 1] - sortedArr[base]);
  }
  return sortedArr[base];
}

function buildGroupStats(nums) {
  const valid = nums.filter((n) => typeof n === 'number' && Number.isFinite(n) && n > 0).sort((a, b) => a - b);
  if (valid.length < 4) return null;
  const median = computeQuantile(valid, 0.5);
  const q1 = computeQuantile(valid, 0.25);
  const q3 = computeQuantile(valid, 0.75);
  const iqr = Math.max(q3 - q1, median * 0.15);
  const absDevs = valid.map((v) => Math.abs(v - median)).sort((a, b) => a - b);
  const mad = Math.max(computeQuantile(absDevs, 0.5), median * 0.08);
  return {
    count: valid.length,
    median,
    q1,
    q3,
    iqr,
    mad,
    upperFence: Math.max(q3 + 3.0 * iqr, median + 5.5 * mad, median * 3.4),
    lowerFence: Math.max(0, Math.min(q1 - 3.0 * iqr, median * 0.08))
  };
}

// ============================================================================
// 8. ĐƯỜNG ỐNG LÀM SẠCH TỔNG THỂ (RUN CLEANING PIPELINE)
// ============================================================================

const DEFAULT_CLEAN_CONFIG = {
  normalizeUnicodeWhitespace: true,
  standardizeDates: true,
  datePreference: 'AUTO_VN', // AUTO_VN | STRICT_DDMM | US_MMDD
  standardizePhones: true,
  standardizeNumbers: true,
  enableFuzzyMatching: true,
  fuzzyThreshold: 0.84,
  customAliases: {}, // { [columnName]: { [variant]: canonical } } hoặc phẳng
  dedupeStrategy: 'remove_hidden', // remove_hidden | remove_exact | flag_only | off
  outlierStrategy: 'median_impute', // median_impute | winsorize | drop_rows | flag_only | off
  fillMissingNumeric: false
};

function runCleaningPipeline(columns, rawRows, userConfig = {}, explicitMapping = null) {
  const startedAt = performance.now();
  const config = { ...DEFAULT_CLEAN_CONFIG, ...(userConfig || {}) };
  const inferred = inferColumnSchema(columns, rawRows);
  const mapping = {
    ...inferred.suggestedMapping,
    ...(explicitMapping || {})
  };

  // Đồng bộ lại kiểu cột dựa trên mapping người dùng chọn (nếu có override)
  const colTypeMap = new Map();
  for (const p of inferred.profiles) {
    colTypeMap.set(p.name, p.detectedType);
  }
  if (mapping.dateCol) colTypeMap.set(mapping.dateCol, 'date');
  if (mapping.phoneCol) colTypeMap.set(mapping.phoneCol, 'phone');
  if (mapping.idCol) colTypeMap.set(mapping.idCol, 'id');
  if (mapping.metricCol) colTypeMap.set(mapping.metricCol, 'numeric');
  if (mapping.quantityCol) colTypeMap.set(mapping.quantityCol, 'numeric');
  if (mapping.unitPriceCol) colTypeMap.set(mapping.unitPriceCol, 'numeric');

  // Bước 1: Tiền tính toán cột Ngày tháng & bộ Gom cụm Fuzzy cho các cột Categorical
  const dateResolutions = new Map();
  const fuzzyMatchers = new Map();
  const allFuzzyClusters = [];

  for (const col of columns) {
    const cType = colTypeMap.get(col);
    if (cType === 'date' && config.standardizeDates) {
      dateResolutions.set(
        col,
        disambiguateDateColumn(
          rawRows.map((r) => r?.[col]),
          config.datePreference
        )
      );
    } else if (cType === 'categorical') {
      const colCustomAliases =
        config.customAliases && typeof config.customAliases[col] === 'object' && !Array.isArray(config.customAliases[col])
          ? config.customAliases[col]
          : config.customAliases || {};

      if (config.enableFuzzyMatching) {
        const built = buildFuzzyMappingForColumn(
          rawRows.map((r) => r?.[col]),
          {
            threshold: Number(config.fuzzyThreshold || 0.84),
            customAliases: colCustomAliases,
            colName: col
          }
        );
        fuzzyMatchers.set(col, built);
        allFuzzyClusters.push(...built.clusters);
      }
    }
  }

  // Bước 2: Chạy chuẩn hóa cấp độ từng Ô (Cell-Level Cleaning)
  const auditLogs = [];
  const issueCounts = {
    whitespace_unicode: 0,
    date_standardized: 0,
    us_date_rescued: 0,
    phone_standardized: 0,
    phone_invalid: 0,
    number_formatted: 0,
    fuzzy_standardized: 0,
    duplicate_exact: 0,
    duplicate_hidden: 0,
    outlier_detected: 0,
    missing_or_error: 0
  };

  const stagedRows = [];

  for (let rIdx = 0; rIdx < rawRows.length; rIdx++) {
    const srcRow = rawRows[rIdx];
    const cleanRow = {};
    const rowChanges = {};

    for (const col of columns) {
      const rawVal = srcRow?.[col] ?? '';
      const cType = colTypeMap.get(col);

      if (cType === 'date' && config.standardizeDates && dateResolutions.has(col)) {
        const res = dateResolutions.get(col)[rIdx];
        cleanRow[col] = res.iso;
        if (res.issueType === 'missing_date' || res.issueType === 'invalid_date') {
          issueCounts.missing_or_error++;
          rowChanges[col] = { old: rawVal, new: res.iso, type: res.issueType };
        } else if (res.changed) {
          issueCounts.date_standardized++;
          if (res.issueType && res.issueType.includes('us_mm_dd')) {
            issueCounts.us_date_rescued++;
          }
          rowChanges[col] = { old: rawVal, new: res.iso, type: res.issueType || 'date_standardized' };
          auditLogs.push({
            rowNumber: rIdx + 1,
            column: col,
            issueType: res.issueType?.includes('us_mm_dd') ? 'us_date_rescued' : 'date_standardized',
            oldValue: String(rawVal),
            newValue: String(res.iso),
            detail:
              res.issueType === 'excel_serial'
                ? 'Chuyển mã số Excel Serial Date sang YYYY-MM-DD'
                : res.issueType?.includes('us_mm_dd')
                ? 'Phát hiện ngày chuẩn Mỹ MM/DD/YYYY lệch định dạng -> Đảo chuẩn về YYYY-MM-DD'
                : 'Chuẩn hóa ngày tháng về chuẩn ISO YYYY-MM-DD'
          });
        }
      } else if (cType === 'phone' && config.standardizePhones) {
        const res = normalizeVietnamPhone(rawVal);
        cleanRow[col] = res.value;
        if (!res.valid) {
          issueCounts.phone_invalid++;
          rowChanges[col] = { old: rawVal, new: res.value, type: 'phone_invalid' };
          auditLogs.push({
            rowNumber: rIdx + 1,
            column: col,
            issueType: 'phone_invalid',
            oldValue: String(rawVal),
            newValue: String(res.value),
            detail: 'SĐT không đúng độ dài hoặc đầu số viễn thông Việt Nam'
          });
        } else if (res.changed) {
          issueCounts.phone_standardized++;
          rowChanges[col] = { old: rawVal, new: res.value, type: res.issueType || 'phone_standardized' };
          auditLogs.push({
            rowNumber: rIdx + 1,
            column: col,
            issueType: 'phone_standardized',
            oldValue: String(rawVal),
            newValue: String(res.value),
            detail:
              res.issueType === 'phone_11_to_10_mnp'
                ? 'Chuyển đổi thuê bao 11 số cũ sang đầu 10 số chuẩn mới'
                : res.issueType === 'phone_missing_zero'
                ? 'Khôi phục số 0 ở đầu bị mất do Excel'
                : 'Chuẩn hóa mã vùng +84 và ký tự phân cách SĐT'
          });
        }
      } else if (cType === 'numeric' && config.standardizeNumbers) {
        const res = parseVietnameseNumber(rawVal);
        cleanRow[col] = res.num !== null ? res.num : '';
        if (res.num === null && String(rawVal).trim() !== '') {
          issueCounts.missing_or_error++;
          rowChanges[col] = { old: rawVal, new: '', type: res.issueType };
          auditLogs.push({
            rowNumber: rIdx + 1,
            column: col,
            issueType: 'missing_or_error',
            oldValue: String(rawVal),
            newValue: 'NULL',
            detail: 'Xử lý ô lỗi công thức Excel / ký tự rỗng'
          });
        } else if (res.changed && res.num !== null) {
          issueCounts.number_formatted++;
          rowChanges[col] = { old: rawVal, new: res.num, type: 'number_formatted' };
          auditLogs.push({
            rowNumber: rIdx + 1,
            column: col,
            issueType: 'number_formatted',
            oldValue: String(rawVal),
            newValue: String(res.num),
            detail: 'Chuyển chuỗi tiền tệ/đơn vị tiếng Việt sang số thực chuẩn tính toán'
          });
        }
      } else if (cType === 'id') {
        const cleanedId = cleanInvisibleAndNFC(rawVal).toUpperCase().replace(/\s*-\s*/g, '-').replace(/\s+/g, '');
        cleanRow[col] = cleanedId;
        if (String(rawVal) !== cleanedId && String(rawVal).trim() !== '') {
          issueCounts.whitespace_unicode++;
          rowChanges[col] = { old: rawVal, new: cleanedId, type: 'id_normalized' };
        }
      } else if (cType === 'categorical' && fuzzyMatchers.has(col)) {
        const res = fuzzyMatchers.get(col).mapValue(rawVal);
        cleanRow[col] = res.value;
        if (res.changed) {
          if (res.issueType === 'fuzzy_standardized') {
            issueCounts.fuzzy_standardized++;
            rowChanges[col] = { old: rawVal, new: res.value, type: 'fuzzy_standardized' };
            auditLogs.push({
              rowNumber: rIdx + 1,
              column: col,
              issueType: 'fuzzy_standardized',
              oldValue: String(rawVal),
              newValue: String(res.value),
              detail: `Gom biến thể viết tắt / sai dấu về tên chuẩn "${res.value}"`
            });
          } else {
            issueCounts.whitespace_unicode++;
            rowChanges[col] = { old: rawVal, new: res.value, type: 'whitespace_unicode' };
            auditLogs.push({
              rowNumber: rIdx + 1,
              column: col,
              issueType: 'whitespace_unicode',
              oldValue: JSON.stringify(rawVal),
              newValue: String(res.value),
              detail: 'Chuẩn hóa bảng mã Unicode NFC & cắt khoảng trắng thừa'
            });
          }
        }
      } else if (config.normalizeUnicodeWhitespace) {
        const cleanedStr = typeof rawVal === 'number' ? rawVal : cleanInvisibleAndNFC(rawVal);
        cleanRow[col] = cleanedStr;
        if (typeof rawVal === 'string' && rawVal !== cleanedStr) {
          issueCounts.whitespace_unicode++;
          rowChanges[col] = { old: rawVal, new: cleanedStr, type: 'whitespace_unicode' };
        }
      } else {
        cleanRow[col] = rawVal;
      }
    }

    // Tự động suy ra cột kỳ _period_month (YYYY-MM) nếu có cột ngày hợp lệ
    const isoDate = mapping.dateCol ? String(cleanRow[mapping.dateCol] || '') : '';
    const derivedPeriod = /^\d{4}-\d{2}-\d{2}$/.test(isoDate) ? isoDate.slice(0, 7) : '';

    stagedRows.push({
      _rowNumber: rIdx + 1,
      _period: derivedPeriod,
      _changes: rowChanges,
      _flags: [],
      _dropped: false,
      _dropReason: null,
      raw: srcRow,
      clean: cleanRow
    });
  }

  // Bước 3: Bắt Dòng Trùng Lặp Chính Xác & Trùng Lặp Ngầm (Exact & Hidden Semantic Duplicates)
  if (config.dedupeStrategy !== 'off') {
    const exactSeen = new Map();
    const semanticSeen = new Map();
    const idSeen = new Map();

    // Các cột làm chữ ký ngữ nghĩa (bỏ cột ID tự tăng nếu muốn bắt 2 dòng nhập trùng khác mã, VÀ kiểm tra cả trùng mã ID!)
    const semanticCols = columns.filter((c) => c !== mapping.idCol);

    for (const entry of stagedRows) {
      const rawSig = columns.map((c) => String(entry.raw?.[c] ?? '')).join('||');
      const cleanSemanticSig = semanticCols
        .map((c) => {
          const v = entry.clean[c];
          return typeof v === 'number' ? String(v) : toAsciiSkeleton(String(v ?? ''));
        })
        .join('||');
      const cleanIdVal = mapping.idCol ? String(entry.clean[mapping.idCol] || '') : '';

      let dupType = null;
      let dupOfRow = null;

      if (exactSeen.has(rawSig)) {
        dupType = 'duplicate_exact';
        dupOfRow = exactSeen.get(rawSig);
      } else if (
        config.dedupeStrategy !== 'remove_exact' &&
        ((cleanIdVal && idSeen.has(cleanIdVal)) || (semanticCols.length >= 3 && semanticSeen.has(cleanSemanticSig)))
      ) {
        dupType = 'duplicate_hidden';
        dupOfRow = cleanIdVal && idSeen.has(cleanIdVal) ? idSeen.get(cleanIdVal) : semanticSeen.get(cleanSemanticSig);
      }

      if (dupType) {
        issueCounts[dupType]++;
        entry._flags.push(dupType);
        const shouldRemove =
          config.dedupeStrategy === 'remove_hidden' ||
          (config.dedupeStrategy === 'remove_exact' && dupType === 'duplicate_exact');

        if (shouldRemove) {
          entry._dropped = true;
          entry._dropReason = dupType === 'duplicate_exact' ? `Trùng khớp 100% dòng #${dupOfRow}` : `Trùng lặp ngầm với dòng #${dupOfRow}`;
        }
        auditLogs.push({
          rowNumber: entry._rowNumber,
          column: mapping.idCol || 'Toàn dòng',
          issueType: dupType,
          oldValue: `Dòng #${entry._rowNumber}`,
          newValue: shouldRemove ? `Đã loại bỏ (Trùng dòng #${dupOfRow})` : `Gắn cờ trùng dòng #${dupOfRow}`,
          detail:
            dupType === 'duplicate_hidden'
              ? `Trùng lặp ngầm sau khi chuẩn hóa ngày/tên/SĐT với dòng #${dupOfRow}`
              : `Bản ghi lặp lại y hệt dòng #${dupOfRow}`
        });
      } else {
        exactSeen.set(rawSig, entry._rowNumber);
        if (cleanSemanticSig.replace(/\|/g, '').length > 0) {
          semanticSeen.set(cleanSemanticSig, entry._rowNumber);
        }
        if (cleanIdVal) {
          idSeen.set(cleanIdVal, entry._rowNumber);
        }
      }
    }
  }

  // Bước 4: Phát hiện & Xử lý Giá trị Ngoại lai (Outliers) theo Ngữ cảnh Nhóm Mặt Hàng
  const activeEntries = stagedRows.filter((e) => !e._dropped);
  const targetMetricCol = mapping.metricCol;
  const qtyCol = mapping.quantityCol;
  const groupCol = mapping.itemCol;

  const outlierDetails = [];
  if (config.outlierStrategy !== 'off' && targetMetricCol) {
    // Gom nhóm theo Mặt hàng/Dịch vụ để so sánh đúng mặt bằng giá (nếu có >= 4 dòng/nhóm, ngược lại fallback toàn cục)
    const globalMetricStats = buildGroupStats(activeEntries.map((e) => e.clean[targetMetricCol]));
    const byGroupMetric = new Map();
    const byGroupUnitRate = new Map();

    if (groupCol) {
      const bucketVals = new Map();
      const bucketRates = new Map();
      for (const e of activeEntries) {
        const gKey = String(e.clean[groupCol] || '_ALL_');
        const mVal = e.clean[targetMetricCol];
        const qVal = qtyCol ? e.clean[qtyCol] : null;
        if (typeof mVal === 'number' && Number.isFinite(mVal)) {
          if (!bucketVals.has(gKey)) bucketVals.set(gKey, []);
          bucketVals.get(gKey).push(mVal);
          if (typeof qVal === 'number' && qVal > 0) {
            if (!bucketRates.has(gKey)) bucketRates.set(gKey, []);
            bucketRates.get(gKey).push(mVal / qVal);
          }
        }
      }
      for (const [gKey, arr] of bucketVals.entries()) {
        const st = buildGroupStats(arr);
        if (st) byGroupMetric.set(gKey, st);
      }
      for (const [gKey, arr] of bucketRates.entries()) {
        const st = buildGroupStats(arr);
        if (st) byGroupUnitRate.set(gKey, st);
      }
    }

    for (const entry of activeEntries) {
      const mVal = entry.clean[targetMetricCol];
      if (typeof mVal !== 'number' || !Number.isFinite(mVal)) continue;

      const gKey = groupCol ? String(entry.clean[groupCol] || '_ALL_') : '_ALL_';
      const stats = byGroupMetric.get(gKey) || globalMetricStats;
      if (!stats) continue;

      const qVal = qtyCol && typeof entry.clean[qtyCol] === 'number' && entry.clean[qtyCol] > 0 ? entry.clean[qtyCol] : null;
      const rateStats = byGroupUnitRate.get(gKey);
      const impliedRate = qVal ? mVal / qVal : null;

      let isOutlier = false;
      let expectedValue = stats.median;
      let reason = '';

      if (mVal < 0) {
        isOutlier = true;
        expectedValue = Math.abs(mVal) <= stats.upperFence ? Math.abs(mVal) : stats.median;
        reason = 'Giá trị âm bất thường (nhập nhầm dấu trừ)';
      } else if (rateStats && impliedRate !== null) {
        // Nếu có cột Số lượng: kiểm tra Đơn giá ngầm định (mVal / qVal) so với trung vị nhóm!
        // Nhờ vậy đơn sỉ số lượng lớn hợp lệ KHÔNG bị báo nhầm, còn gõ dư số 0 (đơn giá vọt gấp 4-10 lần) bị bắt trúng 100%!
        if (impliedRate > Math.max(rateStats.upperFence, rateStats.median * 3.2) && mVal > stats.q3 * 1.8) {
          isOutlier = true;
          expectedValue = Math.round(rateStats.median * qVal);
          reason = `Đơn giá ngầm định cao gấp ${(impliedRate / rateStats.median).toFixed(1)}x trung vị mặt hàng "${gKey}" (nghi gõ dư số 0)`;
        }
      } else if (mVal > Math.max(stats.upperFence, stats.median * 4.5)) {
        isOutlier = true;
        expectedValue = Math.round(stats.median);
        reason = `Giá trị cao gấp ${(mVal / stats.median).toFixed(1)}x trung vị nhóm (vượt ngưỡng MAD/IQR)`;
      }

      if (isOutlier) {
        issueCounts.outlier_detected++;
        entry._flags.push('outlier');
        let replacement = mVal;

        if (config.outlierStrategy === 'median_impute') {
          replacement = expectedValue;
          entry.clean[targetMetricCol] = replacement;
          entry._changes[targetMetricCol] = { old: mVal, new: replacement, type: 'outlier_imputed' };
        } else if (config.outlierStrategy === 'winsorize') {
          const cap = rateStats && qVal ? Math.round(rateStats.upperFence * qVal) : Math.round(stats.upperFence);
          replacement = mVal < 0 ? expectedValue : Math.min(mVal, cap);
          entry.clean[targetMetricCol] = replacement;
          entry._changes[targetMetricCol] = { old: mVal, new: replacement, type: 'outlier_winsorized' };
        } else if (config.outlierStrategy === 'drop_rows') {
          entry._dropped = true;
          entry._dropReason = `Loại bỏ dòng ngoại lai (${reason})`;
        }

        outlierDetails.push({
          rowNumber: entry._rowNumber,
          group: gKey,
          column: targetMetricCol,
          originalValue: mVal,
          resolvedValue: entry._dropped ? null : replacement,
          strategy: config.outlierStrategy,
          reason
        });

        auditLogs.push({
          rowNumber: entry._rowNumber,
          column: targetMetricCol,
          issueType: 'outlier_detected',
          oldValue: String(mVal),
          newValue: entry._dropped ? 'Đã loại bỏ dòng' : String(replacement),
          detail: reason
        });
      }
    }
  }

  // Bước 5: Tổng hợp kết quả Bảng Sạch, Điểm Sức Khỏe Dữ Liệu (Health Score) & Bẫy Dữ Liệu Bẩn
  const finalCleanRows = [];
  for (const e of stagedRows) {
    if (e._dropped) continue;
    const rowOut = { ...e.clean };
    if (e._period) rowOut._period = e._period;
    finalCleanRows.push(rowOut);
  }

  const totalCells = Math.max(1, rawRows.length * columns.length);
  const totalIssuesFound = Object.values(issueCounts).reduce((a, b) => a + b, 0);
  const rawHealthScore = Math.max(
    18,
    Math.min(100, Math.round((1 - Math.min(0.82, (totalIssuesFound * 1.6) / totalCells)) * 1000) / 10)
  );

  // Sau khi làm sạch: chỉ còn các lỗi chưa thể tự sửa (như SĐT sai hẳn độ dài hoặc ô rỗng không điền)
  const residualIssues = issueCounts.phone_invalid + issueCounts.missing_or_error;
  const cleanHealthScore = Math.max(
    rawHealthScore,
    Math.min(100, Math.round((1 - Math.min(0.25, residualIssues / totalCells)) * 1000) / 10)
  );

  // Thống kê sức khỏe từng cột sau khi quét
  const columnHealth = inferred.profiles.map((p) => {
    const colIssues = auditLogs.filter((l) => l.column === p.name).length;
    const completeness = rawRows.length > 0 ? Math.round((p.nonEmptyCount / rawRows.length) * 1000) / 10 : 100;
    const distinctClean = new Set(finalCleanRows.map((r) => String(r[p.name] ?? '')).filter(Boolean)).size;
    return {
      ...p,
      detectedType: colTypeMap.get(p.name) || p.detectedType,
      completenessPct: completeness,
      issuesFixed: colIssues,
      distinctBefore: p.distinctCount,
      distinctAfter: distinctClean
    };
  });

  const elapsedMs = Math.round((performance.now() - startedAt) * 10) / 10;

  return {
    columns,
    sqlColumns: inferred.profiles.map((p) => ({
      original: p.name,
      sqlName: p.sqlName,
      type: colTypeMap.get(p.name) || p.detectedType
    })),
    mapping,
    config,
    stats: {
      rawRowCount: rawRows.length,
      cleanRowCount: finalCleanRows.length,
      droppedRowCount: rawRows.length - finalCleanRows.length,
      totalIssuesFound,
      rawHealthScore,
      cleanHealthScore,
      elapsedMs,
      issueCounts
    },
    columnHealth,
    fuzzyClusters: allFuzzyClusters,
    outliers: outlierDetails,
    auditLogs,
    stagedPreview: stagedRows.map((e) => ({
      rowNumber: e._rowNumber,
      period: e._period,
      dropped: e._dropped,
      dropReason: e._dropReason,
      flags: e._flags,
      changes: e._changes,
      raw: e.raw,
      clean: e.clean
    })),
    cleanRows: finalCleanRows
  };
}

module.exports = {
  DEFAULT_CLEAN_CONFIG,
  cleanInvisibleAndNFC,
  stripVietAccents,
  toAsciiSkeleton,
  toSnakeIdentifier,
  disambiguateDateColumn,
  normalizeVietnamPhone,
  parseVietnameseNumber,
  inferColumnSchema,
  buildFuzzyMappingForColumn,
  runCleaningPipeline
};
