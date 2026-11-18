const test = require('node:test');
const assert = require('node:assert/strict');

const {
  disambiguateDateColumn,
  normalizeVietnamPhone,
  parseVietnameseNumber,
  buildFuzzyMappingForColumn,
  runCleaningPipeline
} = require('../lib/viet-cleaner');
const { runVarianceAnalysis } = require('../lib/variance-engine');
const { parseXlsxBuffer, buildMultiSheetXlsx } = require('../lib/xlsx-codec');
const { getSampleDataset, listSampleDatasets } = require('../lib/sample-datasets');
const { server } = require('../server');

test('1. Chuẩn hóa Ngày tháng VN: cứu ngày Mỹ MM/DD lạc dòng, Excel Serial & ngày chữ tiếng Việt', () => {
  const rawDates = [
    '15/10/2026',
    '18/10/2026',
    '10/23/2026', // Chuẩn Mỹ rõ ràng
    '10-04-2026', // Nhập nhằng (04/10 hay 10/04) nằm giữa tập dữ liệu tháng 09 & 10!
    '21/09/2026',
    '46308', // Excel Serial Date
    'Ngày 08 tháng 09 năm 2026'
  ];

  const resolved = disambiguateDateColumn(rawDates, 'AUTO_VN');
  assert.equal(resolved[0].iso, '2026-10-15');
  assert.equal(resolved[2].iso, '2026-10-23');
  assert.equal(resolved[3].iso, '2026-10-04', 'Phải tự suy luận ngữ cảnh tháng 10 thay vì đẩy nhầm sang tháng 04');
  assert.equal(resolved[4].iso, '2026-09-21');
  assert.match(resolved[5].iso, /^2026-10-\d{2}$/);
  assert.equal(resolved[6].iso, '2026-09-08');
});

test('2. Chuẩn hóa SĐT Việt Nam (+84, mất số 0 đầu, chuyển đầu 11 số -> 10 số MNP) & Tiền tệ VN', () => {
  assert.equal(normalizeVietnamPhone('+84 909.234.567').value, '0909234567');
  assert.equal(normalizeVietnamPhone('01689123456').value, '0389123456'); // Viettel 0168 -> 038
  assert.equal(normalizeVietnamPhone('01223456789').value, '0773456789'); // MobiFone 0122 -> 077
  assert.equal(normalizeVietnamPhone('912345678').value, '0912345678'); // Khôi phục số 0 đầu
  assert.equal(normalizeVietnamPhone('12345').valid, false);

  assert.equal(parseVietnameseNumber('12.500.000 đ').num, 12500000);
  assert.equal(parseVietnameseNumber('8,4 tr').num, 8400000);
  assert.equal(parseVietnameseNumber('1,25 tỷ').num, 1250000000);
  assert.equal(parseVietnameseNumber('(1.250.000)').num, -1250000);
  assert.equal(parseVietnameseNumber('#DIV/0!').num, null);
});

test('3. Fuzzy Matching thông minh: Gom đúng địa danh/tên viết sai nhưng KHÔNG gộp nhầm tên người Việt & Size khác nhau', () => {
  const cities = ['TP. Hồ Chí Minh', 'TP.HCM', 'Sài Gòn', 'SG', 'Hà Nội', 'HN', 'Ha\u0300 Nô\u0323i'];
  const cityMatcher = buildFuzzyMappingForColumn(cities, { colName: 'Chi nhánh' });
  assert.equal(cityMatcher.mapValue('Sài Gòn').value, 'TP. Hồ Chí Minh');
  assert.equal(cityMatcher.mapValue('SG').value, 'TP. Hồ Chí Minh');
  assert.equal(cityMatcher.mapValue('Ha\u0300 Nô\u0323i').value, 'Hà Nội');

  // Chốt chặn chống gộp nhầm tên riêng VN cùng Họ + Đệm nhưng khác Tên chính!
  const staff = [
    'Nguyễn Văn Hùng',
    'Nguyễn Văn Hùng',
    'Nguyen Van Hung',
    'Nguyễn Văn Dũng',
    'Nguyễn Văn Dũng'
  ];
  const staffMatcher = buildFuzzyMappingForColumn(staff, { colName: 'Nhân viên', threshold: 0.82 });
  assert.equal(staffMatcher.mapValue('Nguyen Van Hung').value, 'Nguyễn Văn Hùng');
  assert.equal(staffMatcher.mapValue('Nguyễn Văn Dũng').value, 'Nguyễn Văn Dũng');
  assert.notEqual(
    staffMatcher.mapValue('Nguyễn Văn Hùng').value,
    staffMatcher.mapValue('Nguyễn Văn Dũng').value
  );

  // Chốt chặn chống gộp nhầm tải trọng / mã số
  const trucks = ['Xe Tải 5T', 'Xe Tải 5T', 'Xe Tải 10T'];
  const truckMatcher = buildFuzzyMappingForColumn(trucks, { colName: 'Dịch vụ' });
  assert.notEqual(truckMatcher.mapValue('Xe Tải 5T').value, truckMatcher.mapValue('Xe Tải 10T').value);
});

test('4. Bắt dòng trùng lặp ngầm & Kiểm chứng khớp đại số PVM Bridge sai số = 0 trên cả 5 ngành', () => {
  const allMeta = listSampleDatasets();
  assert.equal(allMeta.length, 5);

  for (const meta of allMeta) {
    const ds = getSampleDataset(meta.id, 0);
    const cleanRes = runCleaningPipeline(ds.columns, ds.rows);
    assert.ok(cleanRes.stats.cleanHealthScore > cleanRes.stats.rawHealthScore);
    assert.ok(cleanRes.stats.droppedRowCount >= 1, `Dataset ${meta.id} phải bắt được các dòng trùng lặp gài sẵn`);

    const varRes = runVarianceAnalysis({
      datasetTitle: ds.name,
      rawRows: ds.rows,
      cleanRows: cleanRes.cleanRows,
      cleanStats: cleanRes.stats,
      mapping: cleanRes.mapping
    });

    assert.equal(varRes.ready, true);
    const { totalR0, totalR1, components } = varRes.pvm;
    const sumBridge =
      totalR0 +
      components.volume +
      components.mix +
      components.price +
      components.newItems +
      components.discontinued;

    assert.ok(
      Math.abs(totalR1 - sumBridge) < 1e-5,
      `Cầu nối PVM ngành ${meta.id} phải khớp toán học 100% (R1=${totalR1}, sumBridge=${sumBridge})`
    );
  }
});

test('5. Kiểm thử End-to-End Server HTTP: Bootstrap, SQLite Window Query, Upload .XLSX có dòng tiêu đề thừa & Xuất 4-Sheet .XLSX', async () => {
  await new Promise((resolve) => server.listen(0, '127.0.0.1', resolve));
  const addr = server.address();
  const baseUrl = `http://127.0.0.1:${addr.port}`;

  try {
    // 5.1 Bootstrap
    const bootRes = await fetch(`${baseUrl}/api/bootstrap`);
    assert.equal(bootRes.status, 200);
    const boot = await bootRes.json();
    assert.equal(boot.datasets.length, 5);
    assert.ok(boot.state.variance.ready);

    // 5.2 SQLite Window Function LAG()
    const firstPresetSql = boot.state.sqlMeta.presets[0].sql;
    const sqlRes = await fetch(`${baseUrl}/api/sql/query`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ sql: firstPresetSql })
    });
    assert.equal(sqlRes.status, 200);
    const sqlData = await sqlRes.json();
    assert.ok(sqlData.rows.length >= 2);

    // Chặn câu lệnh phá hủy DROP TABLE
    const blockRes = await fetch(`${baseUrl}/api/sql/query`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ sql: 'DROP TABLE clean_data;' })
    });
    assert.equal(blockRes.status, 400);

    // 5.3 Xuất báo cáo .XLSX 4-Sheet và đọc ngược lại kiểm chứng cấu trúc OpenXML
    const exportRes = await fetch(`${baseUrl}/api/export/report-xlsx`);
    assert.equal(exportRes.status, 200);
    const xlsxBuf = Buffer.from(await exportRes.arrayBuffer());
    const parsedCleanSheet = parseXlsxBuffer(xlsxBuf, '3_Du_Lieu_Sach_Chuan');
    assert.equal(parsedCleanSheet.availableSheets.length, 4);
    assert.equal(parsedCleanSheet.rows.length, boot.state.stats.cleanRowCount);

    // 5.4 Tạo 1 file .XLSX có 2 dòng tiêu đề công ty trang trí ở đầu -> Upload kiểm tra tự nhận diện dòng Header!
    const customWorkbook = buildMultiSheetXlsx([
      {
        name: 'Bao_Cao_Chi_Nhanh',
        matrix: [
          ['CÔNG TY CỔ PHẦN VẬN HÀNH VIỆT NAM'],
          ['BÁO CÁO DOANH SỐ CHI NHÁNH THÁNG 09 & 10/2026'],
          ['Ngày GD', 'Chi nhánh', 'Sản phẩm', 'Số lượng', 'Doanh thu'],
          ['12/09/2026', 'TP.HCM', 'Gói Cao Cấp', 10, '20.000.000 đ'],
          ['18/09/2026', 'HN', 'Gói Tiêu Chuẩn', 15, '7.500.000 đ'],
          ['10/14/2026', 'Sài Gòn', 'Gói Cao Cấp', 4, '7.200.000 đ'],
          ['20/10/2026', 'Hà Nội ', 'Gói Tiêu Chuẩn', 22, '11.000.000 đ']
        ]
      }
    ]);

    const upRes = await fetch(`${baseUrl}/api/upload`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        fileName: 'bao_cao_phong_ban.xlsx',
        base64: customWorkbook.toString('base64')
      })
    });
    assert.equal(upRes.status, 200);
    const upData = await upRes.json();
    assert.equal(upData.skippedTitleRows, 2, 'Phải tự bỏ qua 2 dòng tiêu đề công ty đầu bảng');
    assert.equal(upData.state.stats.cleanRowCount, 4);
    assert.equal(upData.state.variance.ready, true);
  } finally {
    await new Promise((resolve) => server.close(resolve));
  }
});
