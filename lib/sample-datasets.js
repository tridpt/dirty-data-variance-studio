/**
 * 🏭 DataScrub & Variance Studio — 5 Bộ Dữ Liệu Thực Chiến Đa Ngành Nghề Việt Nam (sample-datasets.js)
 *
 * Bao phủ 5 ngành nghề tuyển DA nhiều nhất tại Việt Nam:
 * 1. 🛒 Bán lẻ & FMCG (Doanh số Đa kênh)
 * 2. 🚚 Logistics & Chuỗi cung ứng (Cước vận tải & Sản lượng tuyến)
 * 3. 🏦 Ngân hàng — Tín dụng SME (Giải ngân & Thu nhập thuần hồ sơ vay)
 * 4. 🏭 Sản xuất & Nhà máy (Sản lượng xuất xưởng & Chuyền máy SMT/CNC)
 * 5. 👥 Nhân sự — HR C&B (Giải phẫu nguyên nhân Quỹ lương & OT tăng vọt)
 *
 * Mỗi bộ dữ liệu đều được gài sẵn các lỗi "đặc sản Excel phòng ban Việt Nam":
 * - Ngày dd/mm/yyyy trộn mm/dd/yyyy (cả ngày nhập nhằng 10/04/2026 giữa tháng 10), số Excel Serial 46308, chữ tiếng Việt
 * - SĐT +84, mất số 0 đầu, đầu 11 số cũ (0168/0122...), dấu chấm/gạch ngang
 * - Tên Tỉnh/Thành viết tắt ("TP.HCM", "SG", "Sài Gòn", "HN", "ĐN", "BD") + Unicode tổ hợp NFD + tên Nhân viên không dấu/dư cách
 * - Tiền tệ "18.500.000 đ", "12,4 tr", dòng trùng lặp ngầm (Hidden Duplicates) & gõ dư số 0 (Outliers 10x)
 */

function createSeededRng(seed = 20261018) {
  let s = seed >>> 0;
  return () => {
    s = (s * 1664525 + 1013904223) >>> 0;
    return s / 4294967296;
  };
}

function isoToExcelSerial(iso) {
  const [y, m, d] = iso.split('-').map(Number);
  const utc = Date.UTC(y, m - 1, d);
  const epoch = Date.UTC(1899, 11, 31);
  const days = Math.round((utc - epoch) / 86400000);
  return days >= 60 ? days + 1 : days;
}

function dirtyDateify(iso, idx, rng) {
  const [y, m, d] = iso.split('-');
  const roll = rng();
  // Cố ý gài các mẫu ngày bẩn đa dạng thực tế
  if (idx % 17 === 3) {
    return String(isoToExcelSerial(iso)); // Excel Serial mất format (VD: 46308)
  }
  if (idx % 13 === 2 && Number(d) > 12) {
    return `${m}/${d}/${y}`; // Chuẩn Mỹ MM/DD/YYYY rõ ràng (VD: 10/19/2026)
  }
  if (idx % 19 === 5 && Number(d) <= 9 && m === '10') {
    return `${m}-0${Number(d)}-${y}`; // Chuẩn Mỹ nhập nhằng giữa tháng 10 (VD: 10-04-2026 -> 4 tháng 10!)
  }
  if (idx % 23 === 7) {
    return `Ngày ${d} tháng ${m} năm ${y}`;
  }
  if (roll < 0.42) {
    return `${d}/${m}/${y}`; // Chuẩn VN DD/MM/YYYY
  }
  if (roll < 0.62) {
    return `${Number(d)}/${Number(m)}/${y.slice(2)}`; // D/M/YY
  }
  if (roll < 0.78) {
    return `${y}.${m}.${d}`;
  }
  return iso;
}

function dirtyPhoneify(base10, idx, rng) {
  const clean = base10.replace(/\D/g, '');
  const roll = rng();
  if (idx % 16 === 4) {
    // Đầu 11 số cũ Viettel / MobiFone
    return '0168' + clean.slice(3);
  }
  if (idx % 21 === 8) {
    return '0122' + clean.slice(3);
  }
  if (idx % 15 === 6) {
    // Mất số 0 đầu do cột General trong Excel
    return clean.slice(1);
  }
  if (roll < 0.28) {
    return `+84 ${clean.slice(1, 4)} ${clean.slice(4, 7)} ${clean.slice(7)}`;
  }
  if (roll < 0.52) {
    return `${clean.slice(0, 4)}.${clean.slice(4, 7)}.${clean.slice(7)}`;
  }
  if (roll < 0.68) {
    return `'${clean}`;
  }
  return clean;
}

function dirtyMoneyify(amount, idx, rng) {
  const rounded = Math.round(amount);
  const roll = rng();
  if (idx % 11 === 1) {
    return `${rounded.toLocaleString('de-DE')} đ`; // 12.500.000 đ
  }
  if (idx % 14 === 3 && rounded % 100000 === 0) {
    const tr = (rounded / 1_000_000).toFixed(2).replace(/\.?0+$/, '').replace('.', ',');
    return `${tr} tr`; // 12,5 tr
  }
  if (roll < 0.25) {
    return `${rounded.toLocaleString('en-US')} VNĐ`;
  }
  if (roll < 0.4) {
    return ` ${rounded.toLocaleString('de-DE')} `;
  }
  return rounded;
}

function pickVariant(variants, rng) {
  const r = rng();
  if (r < 0.52) return variants[0]; // Bản chuẩn chiếm đa số
  return variants[1 + Math.floor(rng() * (variants.length - 1))];
}

// ============================================================================
// BỘ 1: BÁN LẺ & FMCG (RETAIL & OMNICHANNEL SALES)
// ============================================================================

function buildRetailDataset(periodShift = 0) {
  const rng = createSeededRng(10102026 + periodShift * 997);
  const months =
    periodShift === 0
      ? ['2026-08', '2026-09', '2026-10']
      : ['2026-09', '2026-10', '2026-11'];

  const branches = {
    hcm: ['TP. Hồ Chí Minh', 'TP.HCM', 'Sài Gòn', 'Tp. Hồ Chí Minh ', 'SG', 'TP HCM'],
    hn: ['Hà Nội', 'HN', 'TP. Hà Nội', 'Ha Noi', 'Hà Nội  ', 'Ha\u0300 Nô\u0323i'],
    dn: ['Đà Nẵng', 'ĐN', 'TP Đà Nẵng', 'Da Nang'],
    ct: ['Cần Thơ', 'CT', 'TP. Cần Thơ']
  };

  const repsByBranch = {
    hcm: [
      ['Trần Minh Khoa', 'Tran Minh Khoa', 'Trần  Minh Khoa ', 'NV Khoa - Trần Minh Khoa'],
      ['Lê Hoàng Nam', 'Le Hoang Nam', 'Lê Hoàng Nam ']
    ],
    hn: [
      ['Nguyễn Thị Lan Anh', 'Nguyen Thi Lan Anh', 'Nguyễn Thị  Lan Anh'],
      ['Phạm Quốc Bảo', 'Pham Quoc Bao', 'Phạm Quốc Bảo ']
    ],
    dn: [['Võ Thanh Tùng', 'Vo Thanh Tung', 'Võ  Thanh Tùng']],
    ct: [['Huỳnh Ngọc Mai', 'Huynh Ngoc Mai', 'Huỳnh Ngọc Mai ']]
  };

  const channels = {
    b2b: ['Đại lý B2B', 'B2B Đại lý', 'Kênh Đại lý', 'GT Đại lý'],
    store: ['Cửa hàng Trực tiếp', 'POS Cửa hàng', 'Tại cửa hàng'],
    ecom: ['Sàn TMĐT', 'TMĐT', 'Ecommerce'],
    app: ['Ứng dụng Mobile', 'App Mobile', 'Mobile App']
  };

  const products = [
    {
      id: 'vip_combo',
      variants: ['Combo Quà Doanh Nghiệp VIP', 'Combo Qua Doanh Nghiep VIP', 'Combo Quà Doanh Nghiệp VIP '],
      basePrice: 1_200_000
    },
    {
      id: 'espresso',
      variants: ['Máy Pha Espresso Mini', 'May Pha Espresso Mini', 'Máy  Pha Espresso Mini'],
      basePrice: 2_400_000
    },
    {
      id: 'nut_box',
      variants: ['Hộp Hạt Macca Thượng Hạng', 'Hop Hat Macca Thuong Hang', 'Hộp Hạt Macca Thượng Hạng '],
      basePrice: 480_000
    },
    {
      id: 'coffee_bag',
      variants: ['Cà Phê Đặc Sản Túi 250g', 'Ca Phe Dac San Tui 250g', 'Cà Phê Đặc Sản  Túi 250g'],
      basePrice: 140_000
    }
  ];

  const phones = [
    '0909234567', '0912345678', '0987654321', '0938112233',
    '0976445566', '0945778899', '0963221144', '0903889977'
  ];

  const rows = [];
  let orderSeq = 1001 + periodShift * 400;

  // Quy hoạch kịch bản thực tế:
  // Kỳ giữa (mIdx=1) cao khỏe; Kỳ cuối (mIdx=2) sụt giảm ~15% chủ yếu ở TP.HCM (Trần Minh Khoa - Đại lý B2B)
  // do tụt mạnh mặt hàng giá cao (Combo VIP & Máy Espresso) + giảm giá chiết khấu 10-12%, trong khi túi cà phê 140k tăng nhẹ -> Âm Mix & Âm Price cực rõ!
  const plan = [
    // [branchKey, repIdx, channelKey, prodIdx, qtyM0, qtyM1, qtyM2, priceFactorM2]
    ['hcm', 0, 'b2b', 0, 26, 32, 11, 0.88],
    ['hcm', 0, 'b2b', 1, 14, 18, 6, 0.90],
    ['hcm', 0, 'app', 3, 45, 50, 95, 0.92],
    ['hcm', 1, 'store', 2, 34, 38, 30, 0.95],
    ['hcm', 1, 'ecom', 3, 55, 60, 82, 0.94],
    ['hn', 0, 'b2b', 0, 20, 22, 19, 0.96],
    ['hn', 0, 'store', 1, 10, 12, 11, 0.97],
    ['hn', 1, 'ecom', 2, 36, 40, 44, 0.98],
    ['hn', 1, 'app', 3, 50, 54, 62, 0.96],
    ['dn', 0, 'b2b', 0, 12, 15, 10, 0.93],
    ['dn', 0, 'store', 2, 24, 28, 25, 0.96],
    ['ct', 0, 'store', 2, 20, 22, 24, 1.0],
    ['ct', 0, 'b2b', 3, 40, 42, 48, 0.98]
  ];

  months.forEach((ym, mIdx) => {
    plan.forEach((cfg, cIdx) => {
      const [bKey, rIdx, chKey, pIdx, q0, q1, q2, pf2] = cfg;
      const totalQty = mIdx === 0 ? q0 : mIdx === 1 ? q1 : q2;
      // Chia thành 3 giao dịch trong tháng để bảng có ~125 dòng chân thực
      const splits = [
        Math.max(1, Math.floor(totalQty * 0.35)),
        Math.max(1, Math.floor(totalQty * 0.35)),
        Math.max(1, totalQty - Math.floor(totalQty * 0.35) * 2)
      ];

      splits.forEach((qty, sIdx) => {
        const day = Math.min(28, 3 + cIdx + sIdx * 7);
        const iso = `${ym}-${String(day).padStart(2, '0')}`;
        const prod = products[pIdx];
        const pFactor = mIdx === 2 ? pf2 : mIdx === 1 ? 1.0 : 0.99;
        const unitPrice = Math.round((prod.basePrice * pFactor) / 1000) * 1000;
        const rev = qty * unitPrice;

        const rowIdx = rows.length;
        const code = `ORD-${orderSeq++}`;
        const phoneBase = phones[(cIdx + sIdx + mIdx) % phones.length];

        rows.push({
          'Mã đơn': rowIdx % 18 === 5 ? ` ${code.toLowerCase()} ` : code,
          'Ngày GD': dirtyDateify(iso, rowIdx, rng),
          'Chi nhánh': pickVariant(branches[bKey], rng),
          'Nhân viên': pickVariant(repsByBranch[bKey][rIdx], rng),
          'Kênh bán': pickVariant(channels[chKey], rng),
          'Sản phẩm': pickVariant(prod.variants, rng),
          'SĐT khách hàng': dirtyPhoneify(phoneBase, rowIdx, rng),
          'Số lượng': qty,
          'Doanh thu': dirtyMoneyify(rev, rowIdx, rng)
        });
      });
    });
  });

  // Gài 2 dòng Ngoại lai "Gõ dư 1 số 0" ở kỳ cuối (làm bẫy báo cáo bẩn tăng ảo hàng trăm triệu!)
  const lastPeriodRows = rows.slice(-25);
  if (lastPeriodRows[3]) {
    lastPeriodRows[3]['Doanh thu'] = '144.000.000 đ'; // Đáng lẽ ~14.4tr -> gõ dư 1 số 0!
  }
  if (lastPeriodRows[11]) {
    lastPeriodRows[11]['Doanh thu'] = 128000000; // Đáng lẽ ~12.8tr -> gõ dư 1 số 0!
  }

  // Gài 2 dòng Trùng lặp 100% (Copy-paste đúp dòng) & 3 dòng Trùng lặp ngầm (nhập lại khác cách viết tắt)
  const sampleExact = { ...rows[rows.length - 8] };
  rows.push(sampleExact);

  const hiddenSrc1 = rows[rows.length - 15];
  rows.push({
    'Mã đơn': hiddenSrc1['Mã đơn'].toLowerCase().replace('-', ' - '),
    'Ngày GD': '2026.10.17',
    'Chi nhánh': 'Sài Gòn',
    'Nhân viên': 'Tran Minh Khoa',
    'Kênh bán': 'B2B Đại lý',
    'Sản phẩm': hiddenSrc1['Sản phẩm'],
    'SĐT khách hàng': '+84 909 234 567',
    'Số lượng': hiddenSrc1['Số lượng'],
    'Doanh thu': hiddenSrc1['Doanh thu']
  });

  const hiddenSrc2 = rows[rows.length - 20];
  rows.push({
    ...hiddenSrc2,
    'Mã đơn': `DUP-${orderSeq++}`, // Khác mã đơn nhưng cùng ngày, chi nhánh, NV, khách, SP, số tiền!
    'Chi nhánh': hiddenSrc2['Chi nhánh'] === 'Hà Nội' ? 'HN' : hiddenSrc2['Chi nhánh'],
    'SĐT khách hàng': String(hiddenSrc2['SĐT khách hàng']).replace(/\D/g, '')
  });

  return {
    id: 'retail_fmcg',
    name: '🛒 Bán Lẻ & Chuỗi FMCG — Doanh Số Đa Kênh',
    industry: 'Bán lẻ / F&B / FMCG',
    badge: 'Giảm -15% do Âm Mix & Chiết khấu B2B tại TP.HCM',
    description:
      'File tổng hợp doanh số từ 4 khu vực gửi về: Lẫn lộn "TP.HCM/Sài Gòn/SG", ngày Mỹ MM/DD/YYYY lạc giữa tháng 10, SĐT 11 số cũ, dòng nhập đúp ngầm và 2 dòng gõ dư số 0 làm sai lệch hoàn toàn kết luận nếu không lọc sạch.',
    recommendedRecipeId: 'rcp_vn_standard_sales',
    columns: ['Mã đơn', 'Ngày GD', 'Chi nhánh', 'Nhân viên', 'Kênh bán', 'Sản phẩm', 'SĐT khách hàng', 'Số lượng', 'Doanh thu'],
    rows
  };
}

// ============================================================================
// BỘ 2: LOGISTICS & VẬN TẢI (FREIGHT & FLEET OPERATIONS)
// ============================================================================

function buildLogisticsDataset(periodShift = 0) {
  const rng = createSeededRng(20202026 + periodShift * 733);
  const months = periodShift === 0 ? ['2026-08', '2026-09', '2026-10'] : ['2026-09', '2026-10', '2026-11'];

  const hubs = {
    bd: ['Bình Dương', 'BD', 'Kho Sóng Thần - Bình Dương' === '' ? '' : 'Tỉnh Bình Dương', 'Binh Duong'],
    hcm: ['TP. Hồ Chí Minh', 'TP.HCM', 'Sài Gòn', 'SG'],
    bn: ['Bắc Ninh', 'BN', 'Bắc Ninh ', 'Bac Ninh'],
    dn: ['Đà Nẵng', 'ĐN', 'Da Nang']
  };

  const coordinators = {
    bd: [['Đặng Quốc Huy', 'Dang Quoc Huy', 'Đặng  Quốc Huy']],
    hcm: [['Ngô Thanh Phong', 'Ngo Thanh Phong', 'Ngô Thanh Phong ']],
    bn: [['Bùi Minh Tuấn', 'Bui Minh Tuan', 'Bùi  Minh Tuấn']],
    dn: [['Lý Hoàng Hải', 'Ly Hoang Hai', 'Lý Hoàng Hải ']]
  };

  const services = [
    { variants: ['Container Lạnh 40ft Bắc-Nam', 'Container Lanh 40ft Bac-Nam'], rate: 3_600_000 },
    { variants: ['Xe Tải Khô 8T Liên Tỉnh', 'Xe Tai Kho 8T Lien Tinh'], rate: 1_450_000 },
    { variants: ['Giao Nhanh Nội Đô 2T', 'Giao Nhanh Noi Do 2T'], rate: 420_000 }
  ];

  const segments = [
    ['Chuỗi Siêu thị Lạnh', 'Chuoi Sieu thi Lanh'],
    ['Nhà máy Xuất khẩu FDI', 'Nha may Xuat khau FDI'],
    ['Thương mại Điện tử', 'Thuong mai Dien tu']
  ];

  const rows = [];
  let seq = 5001 + periodShift * 300;

  const plan = [
    ['bd', 0, 1, 24, 30, 13, 0.93],
    ['bd', 1, 1, 35, 38, 32, 0.96],
    ['bd', 2, 2, 50, 55, 92, 0.95],
    ['hcm', 0, 0, 18, 22, 16, 0.95],
    ['hcm', 2, 2, 60, 65, 88, 0.94],
    ['bn', 1, 1, 40, 44, 42, 0.98],
    ['bn', 0, 1, 16, 18, 17, 0.99],
    ['dn', 0, 0, 15, 19, 9, 0.92],
    ['dn', 1, 2, 28, 30, 27, 0.97]
  ];

  months.forEach((ym, mIdx) => {
    plan.forEach((cfg, cIdx) => {
      const [hKey, sIdx, segIdx, q0, q1, q2, rf2] = cfg;
      const totalTons = mIdx === 0 ? q0 : mIdx === 1 ? q1 : q2;
      const splits = [Math.floor(totalTons * 0.5), totalTons - Math.floor(totalTons * 0.5)];

      splits.forEach((tons, k) => {
        if (tons <= 0) return;
        const day = Math.min(27, 4 + cIdx * 2 + k * 6);
        const iso = `${ym}-${String(day).padStart(2, '0')}`;
        const svc = services[sIdx];
        const unitRate = Math.round((svc.rate * (mIdx === 2 ? rf2 : 1)) / 1000) * 1000;
        const freight = tons * unitRate;
        const idx = rows.length;

        rows.push({
          'Mã vận đơn': `WB-${seq++}`,
          'Ngày vận đơn': dirtyDateify(iso, idx, rng),
          'Trung tâm khai thác': pickVariant(hubs[hKey], rng),
          'Điều phối viên': pickVariant(coordinators[hKey][0], rng),
          'Dịch vụ vận chuyển': pickVariant(svc.variants, rng),
          'Nhóm khách hàng': pickVariant(segments[segIdx], rng),
          'SĐT liên hệ': dirtyPhoneify(`09385566${10 + ((cIdx + k) % 80)}`, idx, rng),
          'Số tấn quy đổi': tons,
          'Tổng cước thu': dirtyMoneyify(freight, idx, rng)
        });
      });
    });
  });

  // Gài ngoại lai gõ dư số 0 và dòng trùng lặp ngầm
  rows[rows.length - 4]['Tổng cước thu'] = '380.000.000 đ';
  rows.push({
    ...rows[rows.length - 9],
    'Trung tâm khai thác': 'BD',
    'Điều phối viên': 'Dang Quoc Huy'
  });

  return {
    id: 'logistics_fleet',
    name: '🚚 Logistics & Chuỗi Cung Ứng — Cước Vận Tải Tuyến',
    industry: 'Logistics & Giao vận',
    badge: 'Tụt doanh thu cước Cont Lạnh 40ft tại Bình Dương & Đà Nẵng',
    description:
      'Bảng kê vận đơn từ 4 trung tâm khai thác: Lẫn tên viết tắt tỉnh ("BD", "BN", "ĐN", "SG"), cước phí ghi kèm chữ "tr"/"đ", trùng vận đơn đối soát và sụt giảm tuyến đường dài biên cao.',
    recommendedRecipeId: 'rcp_logistics_freight',
    columns: [
      'Mã vận đơn',
      'Ngày vận đơn',
      'Trung tâm khai thác',
      'Điều phối viên',
      'Dịch vụ vận chuyển',
      'Nhóm khách hàng',
      'SĐT liên hệ',
      'Số tấn quy đổi',
      'Tổng cước thu'
    ],
    rows
  };
}

// ============================================================================
// BỘ 3: NGÂN HÀNG — TÍN DỤNG DOANH NGHIỆP SME (BANKING SME LENDING)
// ============================================================================

function buildBankingDataset(periodShift = 0) {
  const rng = createSeededRng(30302026 + periodShift * 557);
  const months = periodShift === 0 ? ['2026-08', '2026-09', '2026-10'] : ['2026-09', '2026-10', '2026-11'];

  const branches = {
    hn: ['Hà Nội', 'HN', 'TP. Hà Nội', 'Ha Noi'],
    hcm: ['TP. Hồ Chí Minh', 'TP.HCM', 'Sài Gòn', 'SG'],
    hp: ['Hải Phòng', 'HP', 'Hai Phong']
  };

  const rms = {
    hn: [['Hoàng Gia Khánh', 'Hoang Gia Khanh', 'Hoàng  Gia Khánh']],
    hcm: [['Phan Thanh Thảo', 'Phan Thanh Thao', 'Phan Thanh Thảo ']],
    hp: [['Đỗ Tuấn Kiệt', 'Do Tuan Kiet', 'Đỗ Tuấn Kiệt ']]
  };

  const loanPackages = [
    { variants: ['Tài Trợ Dự Án Trung Hạn', 'Tai Tro Du An Trung Han'], feePerContract: 68_000_000 },
    { variants: ['Tài Trợ Xuất Nhập Khẩu', 'Tai Tro Xuat Nhap Khau'], feePerContract: 34_000_000 },
    { variants: ['Thấu Chi Vốn Lưu Động', 'Thau Chi Von Luu Dong'], feePerContract: 11_500_000 }
  ];

  const segments = [
    ['Doanh nghiệp SME Vừa', 'Doanh nghiep SME Vua'],
    ['Doanh nghiệp Siêu nhỏ', 'Doanh nghiep Sieu nho']
  ];

  const rows = [];
  let seq = 8001 + periodShift * 200;

  const plan = [
    ['hcm', 0, 0, 8, 10, 4, 0.89],
    ['hcm', 2, 1, 14, 15, 26, 0.92],
    ['hn', 0, 0, 9, 11, 6, 0.91],
    ['hn', 1, 0, 12, 14, 13, 0.96],
    ['hn', 2, 1, 12, 14, 21, 0.94],
    ['hp', 1, 0, 10, 11, 11, 0.98],
    ['hp', 2, 1, 8, 9, 10, 0.97]
  ];

  months.forEach((ym, mIdx) => {
    plan.forEach((cfg, cIdx) => {
      const [bKey, pIdx, sIdx, q0, q1, q2, yf2] = cfg;
      const contracts = mIdx === 0 ? q0 : mIdx === 1 ? q1 : q2;
      const splits = [Math.ceil(contracts / 2), Math.max(1, contracts - Math.ceil(contracts / 2))];

      splits.forEach((cnt, k) => {
        const day = Math.min(28, 5 + cIdx * 3 + k * 5);
        const iso = `${ym}-${String(day).padStart(2, '0')}`;
        const pkg = loanPackages[pIdx];
        const unitFee = Math.round((pkg.feePerContract * (mIdx === 2 ? yf2 : 1)) / 100000) * 100000;
        const income = cnt * unitFee;
        const idx = rows.length;

        rows.push({
          'Mã hồ sơ': `LN-${seq++}`,
          'Ngày giải ngân': dirtyDateify(iso, idx, rng),
          'Chi nhánh': pickVariant(branches[bKey], rng),
          'Chuyên viên QHKH': pickVariant(rms[bKey][0], rng),
          'Gói tín dụng': pickVariant(pkg.variants, rng),
          'Phân khúc KH': pickVariant(segments[sIdx], rng),
          'SĐT khách hàng': dirtyPhoneify(`09048877${12 + ((cIdx + k) % 70)}`, idx, rng),
          'Số khế ước': cnt,
          'Thu nhập thuần kỳ đầu': dirtyMoneyify(income, idx, rng)
        });
      });
    });
  });

  rows[rows.length - 3]['Thu nhập thuần kỳ đầu'] = '1.850.000.000 đ'; // Ngoại lai dư số 0
  rows.push({ ...rows[rows.length - 7], 'Chi nhánh': 'SG' });

  return {
    id: 'banking_sme',
    name: '🏦 Ngân Hàng — Giải Ngân & Thu Nhập Tín Dụng SME',
    industry: 'Ngân hàng & Tài chính (BFSI)',
    badge: 'Số khế ước tăng nhưng Thu nhập giảm do dịch chuyển sang Thấu chi lãi mỏng',
    description:
      'Báo cáo giải ngân tín dụng SME: Tổng số khế ước cho vay không giảm nhưng Thu nhập thuần sụt mạnh vì RM đẩy gói Thấu chi ngắn hạn thay vì Tài trợ Dự án + ưu đãi giảm phí.',
    recommendedRecipeId: 'rcp_vn_standard_sales',
    columns: [
      'Mã hồ sơ',
      'Ngày giải ngân',
      'Chi nhánh',
      'Chuyên viên QHKH',
      'Gói tín dụng',
      'Phân khúc KH',
      'SĐT khách hàng',
      'Số khế ước',
      'Thu nhập thuần kỳ đầu'
    ],
    rows
  };
}

// ============================================================================
// BỘ 4: SẢN XUẤT & NHÀ MÁY (MANUFACTURING & FACTORY OUTPUT)
// ============================================================================

function buildManufacturingDataset(periodShift = 0) {
  const rng = createSeededRng(40402026 + periodShift * 409);
  const months = periodShift === 0 ? ['2026-08', '2026-09', '2026-10'] : ['2026-09', '2026-10', '2026-11'];

  const plants = {
    bn: ['Bắc Ninh', 'BN', 'Bac Ninh', 'Tỉnh Bắc Ninh'],
    hp: ['Hải Phòng', 'HP', 'Hai Phong'],
    bd: ['Bình Dương', 'BD', 'Binh Duong']
  };

  const supervisors = {
    bn: [['Trịnh Văn Hòa', 'Trinh Van Hoa', 'Trịnh  Văn Hòa']],
    hp: [['Mai Xuân Trường', 'Mai Xuan Truong', 'Mai Xuân Trường ']],
    bd: [['Lương Thế Vinh', 'Luong The Vinh', 'Lương Thế Vinh ']]
  };

  const components = [
    { variants: ['Module Camera AI Quang Học', 'Module Camera AI Quang Hoc'], stdVal: 950_000 },
    { variants: ['Bo Mạch Nguồn BMS 48V', 'Bo Mach Nguon BMS 48V'], stdVal: 520_000 },
    { variants: ['Cáp Cảm Biến Tiêu Chuẩn', 'Cap Cam Bien Tieu Chuan'], stdVal: 95_000 }
  ];

  const lines = [
    ['Chuyền SMT-01', 'Chuyen SMT-01'],
    ['Chuyền SMT-02', 'Chuyen SMT-02'],
    ['Chuyền Lắp Ráp A1', 'Chuyen Lap Rap A1']
  ];

  const rows = [];
  let seq = 3001 + periodShift * 250;

  const plan = [
    ['bn', 0, 1, 140, 160, 72, 0.98], // Chuyền SMT-02 bảo trì làm sản lượng Module Camera AI tụt mạnh!
    ['bn', 1, 0, 180, 190, 182, 0.99],
    ['bn', 2, 2, 320, 340, 460, 1.0],
    ['hp', 0, 0, 110, 120, 116, 0.99],
    ['hp', 1, 1, 150, 165, 160, 1.0],
    ['bd', 1, 0, 130, 140, 134, 0.98],
    ['bd', 2, 2, 260, 280, 310, 0.99]
  ];

  months.forEach((ym, mIdx) => {
    plan.forEach((cfg, cIdx) => {
      const [pKey, compIdx, lineIdx, q0, q1, q2, vf2] = cfg;
      const qty = mIdx === 0 ? q0 : mIdx === 1 ? q1 : q2;
      const splits = [Math.floor(qty / 2), qty - Math.floor(qty / 2)];

      splits.forEach((units, k) => {
        const day = Math.min(28, 3 + cIdx * 3 + k * 8);
        const iso = `${ym}-${String(day).padStart(2, '0')}`;
        const comp = components[compIdx];
        const val = units * Math.round((comp.stdVal * (mIdx === 2 ? vf2 : 1)) / 1000) * 1000;
        const idx = rows.length;

        rows.push({
          'Mã lô SX': `LOT-${seq++}`,
          'Ngày chốt ca': dirtyDateify(iso, idx, rng),
          'Nhà máy': pickVariant(plants[pKey], rng),
          'Trưởng ca': pickVariant(supervisors[pKey][0], rng),
          'Dòng linh kiện': pickVariant(comp.variants, rng),
          'Chuyền sản xuất': pickVariant(lines[lineIdx], rng),
          'SĐT QC': dirtyPhoneify(`09773344${15 + ((cIdx + k) % 60)}`, idx, rng),
          'Sản lượng đạt chuẩn': units,
          'Giá trị xuất xưởng': dirtyMoneyify(val, idx, rng)
        });
      });
    });
  });

  rows[rows.length - 5]['Giá trị xuất xưởng'] = '680.000.000 đ';
  rows.push({ ...rows[rows.length - 10], 'Nhà máy': 'BN' });

  return {
    id: 'manufacturing_cost',
    name: '🏭 Sản Xuất & Nhà Máy — Giá Trị Xuất Xưởng Chuyền SMT',
    industry: 'Sản xuất Công nghiệp & Điện tử',
    badge: 'Giảm giá trị xuất xưởng do dừng chuyền SMT-02 tại Bắc Ninh',
    description:
      'Nhật ký chốt ca từ 3 nhà máy Bắc Ninh, Hải Phòng, Bình Dương: Tổng số chiếc xuất xưởng tăng (nhờ Cáp cảm biến giá rẻ) nhưng Giá trị xuất xưởng sụt giảm mạnh vì dừng chuyền Module Camera AI.',
    recommendedRecipeId: 'rcp_vn_standard_sales',
    columns: [
      'Mã lô SX',
      'Ngày chốt ca',
      'Nhà máy',
      'Trưởng ca',
      'Dòng linh kiện',
      'Chuyền sản xuất',
      'SĐT QC',
      'Sản lượng đạt chuẩn',
      'Giá trị xuất xưởng'
    ],
    rows
  };
}

// ============================================================================
// BỘ 5: NHÂN SỰ — GIẢI PHẪU QUỸ LƯƠNG & OT TĂNG VỌT (HR PAYROLL COST)
// ============================================================================

function buildHrPayrollDataset(periodShift = 0) {
  const rng = createSeededRng(50502026 + periodShift * 313);
  const months = periodShift === 0 ? ['2026-08', '2026-09', '2026-10'] : ['2026-09', '2026-10', '2026-11'];

  const offices = {
    hcm: ['TP. Hồ Chí Minh', 'TP.HCM', 'Sài Gòn', 'SG'],
    hn: ['Hà Nội', 'HN', 'Ha Noi'],
    dn: ['Đà Nẵng', 'ĐN', 'Da Nang']
  };

  const managers = {
    hcm: [['Vũ Minh Châu', 'Vu Minh Chau', 'Vũ  Minh Châu']],
    hn: [['Đinh Thái Sơn', 'Dinh Thai Son', 'Đinh Thái Sơn ']],
    dn: [['Cao Thùy Linh', 'Cao Thuy Linh', 'Cao Thùy Linh ']]
  };

  const roles = [
    { variants: ['Chuyên Gia AI & Dữ Liệu', 'Chuyen Gia AI & Du Lieu'], dailyCost: 2_600_000 },
    { variants: ['Chuyên Viên Vận Hành Kho', 'Chuyen Vien Van Hanh Kho'], dailyCost: 680_000 },
    { variants: ['Cộng Tác Viên Giao Nhận', 'Cong Tac Vien Giao Nhan'], dailyCost: 420_000 }
  ];

  const depts = [
    ['Khối Công Nghệ & Data', 'Khoi Cong Nghe & Data'],
    ['Khối Vận Hành Chuỗi', 'Khoi Van Hanh Chuoi'],
    ['Khối Kinh Doanh', 'Khoi Kinh Doanh']
  ];

  const rows = [];
  let seq = 9001 + periodShift * 200;

  // Tháng 10: Quỹ lương TĂNG VỌT ~17% so với Tháng 09 do tuyển thêm Chuyên gia AI/Data (Mix tăng) + phụ cấp OT đêm 1.22x ở Vận hành Kho (Rate tăng)!
  const plan = [
    ['hcm', 0, 0, 88, 92, 138, 1.06],
    ['hcm', 1, 1, 210, 220, 245, 1.22],
    ['hcm', 2, 2, 180, 190, 165, 1.02],
    ['hn', 0, 0, 66, 70, 96, 1.05],
    ['hn', 1, 1, 150, 155, 162, 1.14],
    ['dn', 1, 1, 110, 115, 118, 1.04],
    ['dn', 2, 2, 95, 100, 92, 1.0]
  ];

  months.forEach((ym, mIdx) => {
    plan.forEach((cfg, cIdx) => {
      const [oKey, roleIdx, deptIdx, q0, q1, q2, otFactor2] = cfg;
      const workdays = mIdx === 0 ? q0 : mIdx === 1 ? q1 : q2;
      const splits = [Math.floor(workdays / 2), workdays - Math.floor(workdays / 2)];

      splits.forEach((days, k) => {
        const day = Math.min(28, 10 + cIdx * 2 + k * 7);
        const iso = `${ym}-${String(day).padStart(2, '0')}`;
        const role = roles[roleIdx];
        const unitCost = Math.round((role.dailyCost * (mIdx === 2 ? otFactor2 : 1)) / 1000) * 1000;
        const payroll = days * unitCost;
        const idx = rows.length;

        rows.push({
          'Mã bảng công': `PR-${seq++}`,
          'Ngày chốt công': dirtyDateify(iso, idx, rng),
          'Văn phòng': pickVariant(offices[oKey], rng),
          'Quản lý': pickVariant(managers[oKey][0], rng),
          'Cấp bậc vị trí': pickVariant(role.variants, rng),
          'Khối phòng ban': pickVariant(depts[deptIdx], rng),
          'SĐT nhân sự': dirtyPhoneify(`09886655${20 + ((cIdx + k) % 65)}`, idx, rng),
          'Số công quy đổi': days,
          'Tổng chi phí lương': dirtyMoneyify(payroll, idx, rng)
        });
      });
    });
  });

  rows[rows.length - 4]['Tổng chi phí lương'] = '920.000.000 đ';
  rows.push({ ...rows[rows.length - 8], 'Văn phòng': 'TP.HCM' });

  return {
    id: 'hr_payroll',
    name: '👥 Nhân Sự (HR C&B) — Giải Phẫu Quỹ Lương & OT Tăng Vọt',
    industry: 'Nhân sự & Tài chính Nội bộ',
    badge: 'Chi phí Quỹ lương vọt +18% do Đổi cơ cấu Senior & Tăng ca Kho',
    description:
      'Bảng tổng hợp công & chi phí lương từ các văn phòng: Giải đáp câu hỏi của Giám đốc Tài chính vì sao tổng ngày công tăng ít nhưng Quỹ lương tháng 10 lại đội lên hàng trăm triệu đồng.',
    recommendedRecipeId: 'rcp_hr_cost_guard',
    columns: [
      'Mã bảng công',
      'Ngày chốt công',
      'Văn phòng',
      'Quản lý',
      'Cấp bậc vị trí',
      'Khối phòng ban',
      'SĐT nhân sự',
      'Số công quy đổi',
      'Tổng chi phí lương'
    ],
    rows
  };
}

const BUILDERS = {
  retail_fmcg: buildRetailDataset,
  logistics_fleet: buildLogisticsDataset,
  banking_sme: buildBankingDataset,
  manufacturing_cost: buildManufacturingDataset,
  hr_payroll: buildHrPayrollDataset
};

function listSampleDatasets() {
  return Object.keys(BUILDERS).map((id) => {
    const ds = BUILDERS[id](0);
    return {
      id: ds.id,
      name: ds.name,
      industry: ds.industry,
      badge: ds.badge,
      description: ds.description,
      recommendedRecipeId: ds.recommendedRecipeId,
      rowCount: ds.rows.length,
      columnCount: ds.columns.length,
      columns: ds.columns
    };
  });
}

function getSampleDataset(id, periodShift = 0) {
  const fn = BUILDERS[id] || BUILDERS.retail_fmcg;
  return fn(periodShift);
}

module.exports = {
  listSampleDatasets,
  getSampleDataset
};
