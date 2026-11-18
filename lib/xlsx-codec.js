/**
 * 📊 DataScrub & Variance Studio — Native Zero-Dependency .XLSX / .CSV / .TSV / .JSON Codec
 * - Đọc tệp Microsoft Excel (.xlsx) đa Sheet chuẩn ECMA-376 (SharedStrings, InlineStr, Excel Date Serial, Sparse Coords)
 * - Tự động nhận diện dòng Header thật (bỏ qua 1-5 dòng tiêu đề trang trí thường gặp trong file Excel kế toán/kinh doanh VN)
 * - Tự động nhận diện dấu phân cách CSV (, hoặc ; theo chuẩn Excel Regional Việt Nam hoặc Tab từ Ctrl+C Excel)
 * - Xuất tệp .XLSX đa Sheet có định dạng Header, độ rộng cột tự động, số chuẩn (#,##0), phần trăm (0.00%) & CSV UTF-8 BOM
 */

const zlib = require('node:zlib');

function decodeXmlEntities(str) {
  return String(str || '')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, dec) => String.fromCodePoint(Number(dec)))
    .replace(/&#x([0-9a-fA-F]+);/g, (_, hex) => String.fromCodePoint(parseInt(hex, 16)))
    .replace(/&amp;/g, '&');
}

function escapeXml(str) {
  return String(str ?? '')
    .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

// ============================================================================
// 1. CRC32 & ZIP CONTAINER (READ + WRITE)
// ============================================================================

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) {
      c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    }
    table[n] = c >>> 0;
  }
  return table;
})();

function crc32(buf) {
  let crc = 0xffffffff;
  for (let i = 0; i < buf.length; i++) {
    crc = CRC_TABLE[(crc ^ buf[i]) & 0xff] ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function unzipEntries(buf) {
  if (!Buffer.isBuffer(buf) || buf.length < 22) {
    throw new Error('Tệp tải lên không phải định dạng nén .XLSX hợp lệ.');
  }

  const entries = new Map();
  let eocdOffset = -1;
  const minSearch = Math.max(0, buf.length - 65557);
  for (let i = buf.length - 22; i >= minSearch; i--) {
    if (buf.readUInt32LE(i) === 0x06054b50) {
      eocdOffset = i;
      break;
    }
  }

  if (eocdOffset !== -1) {
    const totalEntries = buf.readUInt16LE(eocdOffset + 10);
    const cdOffset = buf.readUInt32LE(eocdOffset + 16);
    let ptr = cdOffset;

    for (let idx = 0; idx < totalEntries && ptr + 46 <= buf.length; idx++) {
      if (buf.readUInt32LE(ptr) !== 0x02014b50) break;
      const compressionMethod = buf.readUInt16LE(ptr + 10);
      const compressedSize = buf.readUInt32LE(ptr + 20);
      const nameLen = buf.readUInt16LE(ptr + 28);
      const extraLen = buf.readUInt16LE(ptr + 30);
      const commentLen = buf.readUInt16LE(ptr + 32);
      const localHeaderOffset = buf.readUInt32LE(ptr + 42);

      const rawName = buf.subarray(ptr + 46, ptr + 46 + nameLen).toString('utf8').replace(/^\/+/, '');
      ptr += 46 + nameLen + extraLen + commentLen;

      if (localHeaderOffset + 30 <= buf.length && buf.readUInt32LE(localHeaderOffset) === 0x04034b50) {
        const localNameLen = buf.readUInt16LE(localHeaderOffset + 26);
        const localExtraLen = buf.readUInt16LE(localHeaderOffset + 28);
        const dataStart = localHeaderOffset + 30 + localNameLen + localExtraLen;
        const dataEnd = dataStart + compressedSize;
        if (dataEnd <= buf.length) {
          const rawSlice = buf.subarray(dataStart, dataEnd);
          const content =
            compressionMethod === 8
              ? zlib.inflateRawSync(rawSlice)
              : compressionMethod === 0
              ? Buffer.from(rawSlice)
              : null;
          if (content) entries.set(rawName, content);
        }
      }
    }
  }

  if (entries.size === 0) {
    let offset = 0;
    while (offset + 30 <= buf.length) {
      if (buf.readUInt32LE(offset) !== 0x04034b50) break;
      const compressionMethod = buf.readUInt16LE(offset + 8);
      const compressedSize = buf.readUInt32LE(offset + 18);
      const nameLen = buf.readUInt16LE(offset + 26);
      const extraLen = buf.readUInt16LE(offset + 28);
      const rawName = buf.subarray(offset + 30, offset + 30 + nameLen).toString('utf8').replace(/^\/+/, '');
      const dataStart = offset + 30 + nameLen + extraLen;
      const dataEnd = dataStart + compressedSize;
      if (dataEnd > buf.length) break;
      const rawSlice = buf.subarray(dataStart, dataEnd);
      if (compressionMethod === 8) {
        entries.set(rawName, zlib.inflateRawSync(rawSlice));
      } else if (compressionMethod === 0) {
        entries.set(rawName, Buffer.from(rawSlice));
      }
      offset = dataEnd;
    }
  }

  return entries;
}

function zipEntries(fileMap) {
  const localChunks = [];
  const centralChunks = [];
  let offset = 0;
  let count = 0;

  for (const [name, contentInput] of Object.entries(fileMap)) {
    const rawBuf = Buffer.isBuffer(contentInput) ? contentInput : Buffer.from(String(contentInput), 'utf8');
    const compressed = zlib.deflateRawSync(rawBuf, { level: 6 });
    const nameBuf = Buffer.from(name, 'utf8');
    const sum = crc32(rawBuf);

    const localHeader = Buffer.alloc(30);
    localHeader.writeUInt32LE(0x04034b50, 0);
    localHeader.writeUInt16LE(20, 4); // version needed
    localHeader.writeUInt16LE(0x0800, 6); // UTF-8 flag
    localHeader.writeUInt16LE(8, 8); // deflate
    localHeader.writeUInt16LE(0, 10);
    localHeader.writeUInt16LE(0x21, 12);
    localHeader.writeUInt32LE(sum, 14);
    localHeader.writeUInt32LE(compressed.length, 18);
    localHeader.writeUInt32LE(rawBuf.length, 22);
    localHeader.writeUInt16LE(nameBuf.length, 26);
    localHeader.writeUInt16LE(0, 28);

    localChunks.push(localHeader, nameBuf, compressed);

    const centralHeader = Buffer.alloc(46);
    centralHeader.writeUInt32LE(0x02014b50, 0);
    centralHeader.writeUInt16LE(20, 4);
    centralHeader.writeUInt16LE(20, 6);
    centralHeader.writeUInt16LE(0x0800, 8);
    centralHeader.writeUInt16LE(8, 10);
    centralHeader.writeUInt16LE(0, 12);
    centralHeader.writeUInt16LE(0x21, 14);
    centralHeader.writeUInt32LE(sum, 16);
    centralHeader.writeUInt32LE(compressed.length, 20);
    centralHeader.writeUInt32LE(rawBuf.length, 24);
    centralHeader.writeUInt16LE(nameBuf.length, 28);
    centralHeader.writeUInt16LE(0, 30);
    centralHeader.writeUInt16LE(0, 32);
    centralHeader.writeUInt16LE(0, 34);
    centralHeader.writeUInt16LE(0, 36);
    centralHeader.writeUInt32LE(0, 38);
    centralHeader.writeUInt32LE(offset, 42);

    centralChunks.push(centralHeader, nameBuf);

    offset += localHeader.length + nameBuf.length + compressed.length;
    count++;
  }

  const centralBuf = Buffer.concat(centralChunks);
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0);
  eocd.writeUInt16LE(0, 4);
  eocd.writeUInt16LE(0, 6);
  eocd.writeUInt16LE(count, 8);
  eocd.writeUInt16LE(count, 10);
  eocd.writeUInt32LE(centralBuf.length, 12);
  eocd.writeUInt32LE(offset, 16);
  eocd.writeUInt16LE(0, 20);

  return Buffer.concat([...localChunks, centralBuf, eocd]);
}

// ============================================================================
// 2. EXCEL COORDINATE & SERIAL DATE HELPERS
// ============================================================================

function colLettersToIndex(letters) {
  let idx = 0;
  const clean = String(letters || '').toUpperCase();
  for (let i = 0; i < clean.length; i++) {
    const code = clean.charCodeAt(i);
    if (code < 65 || code > 90) break;
    idx = idx * 26 + (code - 64);
  }
  return idx - 1;
}

function colIndexToLetters(index) {
  let n = index + 1;
  let str = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    str = String.fromCharCode(65 + rem) + str;
    n = Math.floor((n - 1) / 26);
  }
  return str;
}

function excelSerialToIsoDate(serial) {
  const num = Number(serial);
  if (!Number.isFinite(num) || num < 1 || num > 100000) return null;
  // Lotus 1-2-3 leap year bug adjustment (serial 60 = imaginary 1900-02-29)
  const utcDays = Math.floor(num > 59 ? num - 1 : num);
  const epochMs = Date.UTC(1899, 11, 31) + utcDays * 86400000;
  const d = new Date(epochMs);
  if (Number.isNaN(d.getTime())) return null;
  const yyyy = d.getUTCFullYear();
  const mm = String(d.getUTCMonth() + 1).padStart(2, '0');
  const dd = String(d.getUTCDate()).padStart(2, '0');
  return `${yyyy}-${mm}-${dd}`;
}

// ============================================================================
// 3. PARSE .XLSX WORKBOOK (MULTI-SHEET + SMART HEADER ROW)
// ============================================================================

function parseSharedStrings(xmlText) {
  if (!xmlText) return [];
  const strings = [];
  const siRegex = /<si\b[^>]*>([\s\S]*?)<\/si>/g;
  let match;
  while ((match = siRegex.exec(xmlText)) !== null) {
    const inner = match[1].replace(/<rPh\b[^>]*>[\s\S]*?<\/rPh>/g, '');
    const parts = [];
    const tRegex = /<t(?:\s+[^>]*)?>([\s\S]*?)<\/t>/g;
    let tm;
    while ((tm = tRegex.exec(inner)) !== null) {
      parts.push(decodeXmlEntities(tm[1]));
    }
    strings.push(parts.join(''));
  }
  return strings;
}

function parseDateStyleIndices(stylesXml) {
  const dateStyleSet = new Set();
  if (!stylesXml) return dateStyleSet;

  const customDateFmtIds = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 30, 36, 45, 46, 47, 50, 57]);
  const numFmtRegex = /<numFmt\b([^>]*)\/?>/g;
  let m;
  while ((m = numFmtRegex.exec(stylesXml)) !== null) {
    const attrs = m[1];
    const idMatch = attrs.match(/\bnumFmtId="(\d+)"/);
    const codeMatch = attrs.match(/\bformatCode="([^"]*)"/);
    if (idMatch && codeMatch) {
      const id = Number(idMatch[1]);
      const code = decodeXmlEntities(codeMatch[1])
        .replace(/\[[^\]]*\]/g, '')
        .replace(/"[^"]*"/g, '')
        .toLowerCase();
      if (/[dmy]/.test(code) && !/[#0]/.test(code)) {
        customDateFmtIds.add(id);
      }
    }
  }

  const cellXfsMatch = stylesXml.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/);
  if (cellXfsMatch) {
    const xfRegex = /<xf\b([^>]*)(?:\/>|>[\s\S]*?<\/xf>)/g;
    let idx = 0;
    let xm;
    while ((xm = xfRegex.exec(cellXfsMatch[1])) !== null) {
      const idMatch = xm[1].match(/\bnumFmtId="(\d+)"/);
      if (idMatch && customDateFmtIds.has(Number(idMatch[1]))) {
        dateStyleSet.add(idx);
      }
      idx++;
    }
  }
  return dateStyleSet;
}

function discoverWorkbookSheets(entries) {
  const wbXml = entries.get('xl/workbook.xml')?.toString('utf8') || '';
  const relsXml = entries.get('xl/_rels/workbook.xml.rels')?.toString('utf8') || '';

  const relMap = new Map();
  const relRegex = /<Relationship\b([^>]*)\/?>/g;
  let rm;
  while ((rm = relRegex.exec(relsXml)) !== null) {
    const idM = rm[1].match(/\bId="([^"]+)"/);
    const targetM = rm[1].match(/\bTarget="([^"]+)"/);
    if (idM && targetM) {
      let target = targetM[1].replace(/^\/+/, '');
      if (!target.startsWith('xl/')) target = `xl/${target.replace(/^\.\//, '')}`;
      relMap.set(idM[1], target);
    }
  }

  const sheets = [];
  const sheetRegex = /<sheet\b([^>]*)\/?>/g;
  let sm;
  while ((sm = sheetRegex.exec(wbXml)) !== null) {
    const nameM = sm[1].match(/\bname="([^"]+)"/);
    const rIdM = sm[1].match(/\br:id="([^"]+)"/);
    const stateM = sm[1].match(/\bstate="([^"]+)"/);
    if (stateM && stateM[1] === 'veryHidden') continue;
    const sheetName = nameM ? decodeXmlEntities(nameM[1]) : `Sheet${sheets.length + 1}`;
    const path = (rIdM && relMap.get(rIdM[1])) || `xl/worksheets/sheet${sheets.length + 1}.xml`;
    if (entries.has(path)) {
      sheets.push({ name: sheetName, path });
    }
  }

  if (sheets.length === 0) {
    for (const key of entries.keys()) {
      if (/^xl\/worksheets\/sheet\d+\.xml$/i.test(key)) {
        sheets.push({ name: key.split('/').pop().replace('.xml', ''), path: key });
      }
    }
  }
  return sheets;
}

function parseWorksheetMatrix(sheetXml, sharedStrings, dateStyleSet) {
  const rows = [];
  let maxCols = 0;
  const rowRegex = /<row\b[^>]*>([\s\S]*?)<\/row>/g;
  let rm;

  while ((rm = rowRegex.exec(sheetXml)) !== null) {
    const rowContent = rm[1];
    const rowValues = [];
    let nextColIdx = 0;

    const cellRegex = /<c\b([^>]*)(?:\/>|>([\s\S]*?)<\/c>)/g;
    let cm;
    while ((cm = cellRegex.exec(rowContent)) !== null) {
      const attrs = cm[1] || '';
      const inner = cm[2] || '';
      const refM = attrs.match(/\br="([A-Za-z]+)\d+"/);
      const colIdx = refM ? colLettersToIndex(refM[1]) : nextColIdx;
      nextColIdx = colIdx + 1;

      const typeM = attrs.match(/\bt="([^"]+)"/);
      const styleM = attrs.match(/\bs="(\d+)"/);
      const cellType = typeM ? typeM[1] : '';
      const styleIdx = styleM ? Number(styleM[1]) : -1;

      let value = '';
      if (cellType === 'inlineStr') {
        const parts = [];
        const tRegex = /<t(?:\s+[^>]*)?>([\s\S]*?)<\/t>/g;
        let tm;
        while ((tm = tRegex.exec(inner)) !== null) {
          parts.push(decodeXmlEntities(tm[1]));
        }
        value = parts.join('');
      } else {
        const vMatch = inner.match(/<v(?:\s+[^>]*)?>([\s\S]*?)<\/v>/);
        const rawV = vMatch ? decodeXmlEntities(vMatch[1]) : '';
        if (cellType === 's') {
          const sIdx = Number(rawV);
          value = Number.isInteger(sIdx) && sIdx >= 0 && sIdx < sharedStrings.length ? sharedStrings[sIdx] : rawV;
        } else if (cellType === 'b') {
          value = rawV === '1' ? 'TRUE' : 'FALSE';
        } else if (cellType === 'str' || cellType === 'e') {
          value = rawV;
        } else if (rawV !== '') {
          if (dateStyleSet.has(styleIdx) && !Number.isNaN(Number(rawV))) {
            value = excelSerialToIsoDate(Number(rawV)) || rawV;
          } else {
            const num = Number(rawV);
            value = Number.isFinite(num) ? num : rawV;
          }
        }
      }

      while (rowValues.length < colIdx) rowValues.push('');
      rowValues[colIdx] = value;
    }

    if (rowValues.some((v) => String(v ?? '').trim() !== '')) {
      if (rowValues.length > maxCols) maxCols = rowValues.length;
      rows.push(rowValues);
    }
  }

  return rows.map((r) => {
    const padded = r.slice();
    while (padded.length < maxCols) padded.push('');
    return padded;
  });
}

/**
 * Tự động tìm dòng Header thật trong ma trận (bỏ qua dòng tiêu đề công ty/báo cáo ở đầu file Excel)
 */
function matrixToRecords(matrix) {
  if (!Array.isArray(matrix) || matrix.length === 0) {
    return { columns: [], rows: [], skippedTitleRows: 0 };
  }

  let bestHeaderIdx = 0;
  let bestScore = -1;
  const scanLimit = Math.min(8, matrix.length);

  for (let i = 0; i < scanLimit; i++) {
    const row = matrix[i];
    const nonEmpty = row.filter((c) => String(c ?? '').trim() !== '');
    if (nonEmpty.length < 2) continue;
    const stringLike = nonEmpty.filter((c) => typeof c === 'string' && !/^\d+(\.\d+)?$/.test(c.trim()));
    const uniqueCount = new Set(nonEmpty.map((c) => String(c).trim().toLowerCase())).size;
    // Ưu tiên dòng có nhiều cột không rỗng, hầu hết là chữ không trùng lặp
    const score = nonEmpty.length * 3 + stringLike.length * 2 + uniqueCount * 2 - i * 0.4;
    if (score > bestScore) {
      bestScore = score;
      bestHeaderIdx = i;
    }
  }

  const rawHeaders = matrix[bestHeaderIdx];
  const usedNames = new Map();
  const activeColIndices = [];
  const columns = [];

  for (let c = 0; c < rawHeaders.length; c++) {
    const headerCandidate = String(rawHeaders[c] ?? '').replace(/\s+/g, ' ').trim();
    const hasDataBelow = matrix.slice(bestHeaderIdx + 1).some((r) => String(r[c] ?? '').trim() !== '');
    if (!headerCandidate && !hasDataBelow) continue;

    let colName = headerCandidate || `Cot_${c + 1}`;
    const lower = colName.toLowerCase();
    if (usedNames.has(lower)) {
      const nextCount = usedNames.get(lower) + 1;
      usedNames.set(lower, nextCount);
      colName = `${colName}_${nextCount}`;
    } else {
      usedNames.set(lower, 1);
    }
    activeColIndices.push(c);
    columns.push(colName);
  }

  const rows = [];
  for (let r = bestHeaderIdx + 1; r < matrix.length; r++) {
    const src = matrix[r];
    const obj = {};
    let nonBlank = 0;
    for (let k = 0; k < columns.length; k++) {
      const val = src[activeColIndices[k]] ?? '';
      if (String(val).trim() !== '') nonBlank++;
      obj[columns[k]] = val;
    }
    // Bỏ qua dòng Tổng cộng (Total / Cộng) cuối bảng Excel nếu chỉ có chữ Tổng cộng ở đầu
    const firstCellsText = columns
      .slice(0, 3)
      .map((col) => String(obj[col] ?? '').trim().toLowerCase())
      .join(' ');
    if (
      r >= matrix.length - 3 &&
      /^(tổng cộng|tong cong|grand total|total)\b/.test(firstCellsText.trim()) &&
      nonBlank <= Math.ceil(columns.length * 0.5)
    ) {
      continue;
    }
    if (nonBlank > 0) rows.push(obj);
  }

  return {
    columns,
    rows,
    skippedTitleRows: bestHeaderIdx
  };
}

function parseXlsxBuffer(buf, preferredSheet = null) {
  const entries = unzipEntries(buf);
  const sharedStrings = parseSharedStrings(entries.get('xl/sharedStrings.xml')?.toString('utf8'));
  const dateStyleSet = parseDateStyleIndices(entries.get('xl/styles.xml')?.toString('utf8'));
  const sheets = discoverWorkbookSheets(entries);

  if (sheets.length === 0) {
    throw new Error('Không tìm thấy Worksheet dữ liệu nào bên trong tệp .XLSX.');
  }

  let selected = sheets[0];
  if (preferredSheet !== null && preferredSheet !== undefined && preferredSheet !== '') {
    const byName = sheets.find((s) => s.name.toLowerCase() === String(preferredSheet).toLowerCase());
    if (byName) {
      selected = byName;
    } else if (Number.isInteger(Number(preferredSheet)) && sheets[Number(preferredSheet)]) {
      selected = sheets[Number(preferredSheet)];
    }
  } else if (sheets.length > 1) {
    // Ưu tiên sheet có nhiều dòng dữ liệu nhất nếu sheet 1 chỉ là bìa hướng dẫn
    let bestRowCount = -1;
    for (const s of sheets) {
      const xml = entries.get(s.path)?.toString('utf8') || '';
      const approxRows = (xml.match(/<row\b/g) || []).length;
      if (approxRows > bestRowCount) {
        bestRowCount = approxRows;
        selected = s;
      }
    }
  }

  const sheetXml = entries.get(selected.path)?.toString('utf8') || '';
  const matrix = parseWorksheetMatrix(sheetXml, sharedStrings, dateStyleSet);
  const parsed = matrixToRecords(matrix);

  return {
    format: 'xlsx',
    sheetName: selected.name,
    availableSheets: sheets.map((s) => s.name),
    ...parsed
  };
}

// ============================================================================
// 4. PARSE CSV / TSV / CLIPBOARD / JSON
// ============================================================================

function detectDelimiter(text) {
  const lines = String(text || '')
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter(Boolean)
    .slice(0, 15);
  if (lines.length === 0) return ',';

  const candidates = ['\t', ';', ',', '|'];
  let bestDelim = ',';
  let bestScore = -1;

  for (const delim of candidates) {
    const counts = lines.map((line) => {
      let inQuotes = false;
      let c = 0;
      for (let i = 0; i < line.length; i++) {
        const ch = line[i];
        if (ch === '"') inQuotes = !inQuotes;
        else if (!inQuotes && ch === delim) c++;
      }
      return c;
    });
    const nonZero = counts.filter((c) => c > 0);
    if (nonZero.length === 0) continue;
    const avg = counts.reduce((a, b) => a + b, 0) / counts.length;
    const variance = counts.reduce((acc, c) => acc + Math.pow(c - avg, 2), 0) / counts.length;
    // Độ ổn định cao (variance thấp) và số cột >= 2
    const score = avg * 4 - variance * 2 + (delim === '\t' ? 0.5 : 0);
    if (score > bestScore) {
      bestScore = score;
      bestDelim = delim;
    }
  }
  return bestDelim;
}

function parseDelimitedMatrix(text, explicitDelimiter = null) {
  const cleaned = String(text || '').replace(/^\uFEFF/, '');
  const delimiter = explicitDelimiter || detectDelimiter(cleaned);

  const rows = [];
  let row = [];
  let field = '';
  let inQuotes = false;

  for (let i = 0; i < cleaned.length; i++) {
    const ch = cleaned[i];
    const next = cleaned[i + 1];

    if (inQuotes) {
      if (ch === '"' && next === '"') {
        field += '"';
        i++;
      } else if (ch === '"') {
        inQuotes = false;
      } else {
        field += ch;
      }
    } else if (ch === '"' && field.length === 0) {
      inQuotes = true;
    } else if (ch === delimiter) {
      row.push(field);
      field = '';
    } else if (ch === '\r' && next === '\n') {
      row.push(field);
      if (row.some((v) => v.trim() !== '')) rows.push(row);
      row = [];
      field = '';
      i++;
    } else if (ch === '\n' || ch === '\r') {
      row.push(field);
      if (row.some((v) => v.trim() !== '')) rows.push(row);
      row = [];
      field = '';
    } else {
      field += ch;
    }
  }

  if (field.length > 0 || row.length > 0) {
    row.push(field);
    if (row.some((v) => v.trim() !== '')) rows.push(row);
  }

  const maxCols = rows.reduce((m, r) => Math.max(m, r.length), 0);
  const padded = rows.map((r) => {
    const copy = r.slice();
    while (copy.length < maxCols) copy.push('');
    return copy;
  });

  return { delimiter, matrix: padded };
}

function parseDelimitedText(text, explicitDelimiter = null) {
  const { delimiter, matrix } = parseDelimitedMatrix(text, explicitDelimiter);
  const parsed = matrixToRecords(matrix);
  return {
    format: delimiter === '\t' ? 'tsv' : 'csv',
    delimiter,
    sheetName: 'Data',
    availableSheets: ['Data'],
    ...parsed
  };
}

function parseJsonInput(text) {
  const raw = JSON.parse(String(text || '').replace(/^\uFEFF/, ''));
  const arr = Array.isArray(raw)
    ? raw
    : Array.isArray(raw?.rows)
    ? raw.rows
    : Array.isArray(raw?.data)
    ? raw.data
    : null;
  if (!arr || arr.length === 0) {
    throw new Error('JSON tải lên cần là mảng các đối tượng (Array of Objects) hoặc có trường "rows"/"data".');
  }
  const colSet = new Set();
  for (const item of arr) {
    if (item && typeof item === 'object') {
      Object.keys(item).forEach((k) => colSet.add(k));
    }
  }
  const columns = Array.from(colSet);
  const rows = arr.map((item) => {
    const obj = {};
    for (const c of columns) {
      const val = item?.[c];
      obj[c] = val === null || val === undefined ? '' : typeof val === 'object' ? JSON.stringify(val) : val;
    }
    return obj;
  });
  return {
    format: 'json',
    sheetName: 'JSON_Data',
    availableSheets: ['JSON_Data'],
    columns,
    rows,
    skippedTitleRows: 0
  };
}

function parseTabularUpload({ fileName = 'upload.csv', buffer, text, sheetName = null }) {
  const lower = String(fileName || '').toLowerCase();
  const buf = buffer || (text !== undefined ? Buffer.from(String(text), 'utf8') : Buffer.alloc(0));

  // Kiểm tra chữ ký ZIP PK\x03\x04 cho file .xlsx
  if (lower.endsWith('.xlsx') || (buf.length > 4 && buf.readUInt32LE(0) === 0x04034b50)) {
    return parseXlsxBuffer(buf, sheetName);
  }
  const utf8Text = text !== undefined ? String(text) : buf.toString('utf8');
  if (lower.endsWith('.json') || utf8Text.trim().startsWith('[') || utf8Text.trim().startsWith('{')) {
    try {
      return parseJsonInput(utf8Text);
    } catch (_) {
      // Fallback sang delimited nếu không phải JSON hợp lệ
    }
  }
  return parseDelimitedText(utf8Text, lower.endsWith('.tsv') ? '\t' : null);
}

// ============================================================================
// 5. BUILD MULTI-SHEET STYLED .XLSX WORKBOOK & UTF-8 BOM CSV
// ============================================================================

function sanitizeSheetName(name, idx) {
  const clean = String(name || `Sheet_${idx + 1}`)
    .replace(/[\\/?*[\]:]/g, '_')
    .trim()
    .slice(0, 31);
  return clean || `Sheet_${idx + 1}`;
}

/**
 * Style indices in xl/styles.xml:
 * 0 = Default cell
 * 1 = Dark Navy Header (Bold white text, #0F172A fill)
 * 2 = Section Banner (Bold indigo/cyan text, #E0F2FE fill)
 * 3 = Integer / VND Number (#,##0)
 * 4 = Decimal Number (#,##0.00)
 * 5 = Percentage (0.00%)
 * 6 = Highlight Warning / Alert cell (#FEF3C7 fill)
 * 7 = Positive Gain cell (#DCFCE7 fill, #,##0)
 * 8 = Negative Drop cell (#FEE2E2 fill, #,##0)
 */
function buildStylesXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="3">
    <numFmt numFmtId="164" formatCode="#,##0"/>
    <numFmt numFmtId="165" formatCode="#,##0.00"/>
    <numFmt numFmtId="166" formatCode="+0.0%;-0.0%;0.0%"/>
  </numFmts>
  <fonts count="4">
    <font><sz val="11"/><color rgb="FF0F172A"/><name val="Segoe UI"/></font>
    <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Segoe UI"/></font>
    <font><b/><sz val="11"/><color rgb="FF0C4A6E"/><name val="Segoe UI"/></font>
    <font><b/><sz val="11"/><color rgb="FF1E293B"/><name val="Segoe UI"/></font>
  </fonts>
  <fills count="7">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF0F172A"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFE0F2FE"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFEF3C7"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFDCFCE7"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFEE2E2"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border>
      <left style="thin"><color rgb="FFCBD5E1"/></left>
      <right style="thin"><color rgb="FFCBD5E1"/></right>
      <top style="thin"><color rgb="FFCBD5E1"/></top>
      <bottom style="thin"><color rgb="FFCBD5E1"/></bottom>
      <diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="9">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>
    <xf numFmtId="165" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>
    <xf numFmtId="166" fontId="3" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="4" borderId="1" xfId="0" applyFill="1" applyBorder="1"/>
    <xf numFmtId="164" fontId="3" fillId="5" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="164" fontId="3" fillId="6" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>
  </cellXfs>
</styleSheet>`;
}

function buildWorksheetXml(sheetSpec) {
  // Hỗ trợ truyền thẳng matrix rows với style cell tùy chọn hoặc { columns, rows }
  const matrix = [];
  if (Array.isArray(sheetSpec.matrix)) {
    for (const r of sheetSpec.matrix) matrix.push(r);
  } else {
    const cols = sheetSpec.columns || [];
    matrix.push(cols.map((c) => ({ v: typeof c === 'object' ? c.label : c, s: 1 })));
    const colKeys = cols.map((c) => (typeof c === 'object' ? c.key : c));
    const colFormats = cols.map((c) => (typeof c === 'object' ? c.format : null));

    for (const item of sheetSpec.rows || []) {
      const rowCells = colKeys.map((k, idx) => {
        const val = item?.[k];
        const fmt = colFormats[idx];
        if (typeof val === 'number' && Number.isFinite(val)) {
          let s = Number.isInteger(val) || Math.abs(val) >= 100 ? 3 : 4;
          if (fmt === 'pct') s = 5;
          else if (fmt === 'delta') s = val > 0 ? 7 : val < 0 ? 8 : 3;
          return { v: val, t: 'n', s };
        }
        return { v: val ?? '', t: 'str', s: 0 };
      });
      matrix.push(rowCells);
    }
  }

  const colWidths = [];
  for (const row of matrix) {
    for (let c = 0; c < row.length; c++) {
      const cell = row[c] && typeof row[c] === 'object' && 'v' in row[c] ? row[c] : { v: row[c] };
      const len = String(cell.v ?? '').length;
      colWidths[c] = Math.min(52, Math.max(colWidths[c] || 12, len + 4));
    }
  }

  const colsXml =
    colWidths.length > 0
      ? `<cols>${colWidths
          .map((w, idx) => `<col min="${idx + 1}" max="${idx + 1}" width="${w}" customWidth="1"/>`)
          .join('')}</cols>`
      : '';

  const rowsXml = matrix
    .map((row, rIdx) => {
      const rowNum = rIdx + 1;
      const cellsXml = row
        .map((rawCell, cIdx) => {
          const cell = rawCell && typeof rawCell === 'object' && 'v' in rawCell ? rawCell : { v: rawCell, s: 0 };
          const ref = `${colIndexToLetters(cIdx)}${rowNum}`;
          const sAttr = cell.s !== undefined ? ` s="${cell.s}"` : ' s="0"';
          const val = cell.v;
          if (val === null || val === undefined || val === '') {
            return `<c r="${ref}"${sAttr}/>`;
          }
          if ((cell.t === 'n' || typeof val === 'number') && Number.isFinite(Number(val))) {
            return `<c r="${ref}"${sAttr}><v>${Number(val)}</v></c>`;
          }
          return `<c r="${ref}" t="inlineStr"${sAttr}><is><t xml:space="preserve">${escapeXml(val)}</t></is></c>`;
        })
        .join('');
      const htAttr = rIdx === 0 ? ' ht="24" customHeight="1"' : '';
      return `<row r="${rowNum}"${htAttr}>${cellsXml}</row>`;
    })
    .join('');

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews><sheetView workbookViewId="0"/></sheetViews>
  ${colsXml}
  <sheetData>${rowsXml}</sheetData>
</worksheet>`;
}

function buildMultiSheetXlsx(sheets) {
  const validSheets = Array.isArray(sheets) && sheets.length > 0 ? sheets : [{ name: 'Sheet1', columns: ['Info'], rows: [] }];
  const files = {};

  const contentTypesOverrides = [
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
  ];

  const wbSheetEntries = [];
  const wbRelsEntries = [
    '<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
  ];

  const usedNames = new Set();
  validSheets.forEach((spec, i) => {
    let sheetName = sanitizeSheetName(spec.name, i);
    while (usedNames.has(sheetName.toLowerCase())) {
      sheetName = `${sheetName.slice(0, 28)}_${i + 1}`;
    }
    usedNames.add(sheetName.toLowerCase());

    const sheetFile = `worksheets/sheet${i + 1}.xml`;
    const rId = `rIdSheet${i + 1}`;
    files[`xl/${sheetFile}`] = buildWorksheetXml(spec);
    contentTypesOverrides.push(
      `<Override PartName="/xl/${sheetFile}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`
    );
    wbSheetEntries.push(`<sheet name="${escapeXml(sheetName)}" sheetId="${i + 1}" r:id="${rId}"/>`);
    wbRelsEntries.push(
      `<Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="${sheetFile}"/>`
    );
  });

  files['[Content_Types].xml'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  ${contentTypesOverrides.join('\n  ')}
</Types>`;

  files['_rels/.rels'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;

  files['xl/workbook.xml'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    ${wbSheetEntries.join('\n    ')}
  </sheets>
</workbook>`;

  files['xl/_rels/workbook.xml.rels'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${wbRelsEntries.join('\n  ')}
</Relationships>`;

  files['xl/styles.xml'] = buildStylesXml();

  return zipEntries(files);
}

function buildUtf8BomCsv(columns, rows) {
  const escapeCsvCell = (val) => {
    if (val === null || val === undefined) return '';
    const str = String(val);
    if (/[",\r\n;]/.test(str)) {
      return `"${str.replace(/"/g, '""')}"`;
    }
    return str;
  };
  const headerLine = columns.map(escapeCsvCell).join(',');
  const bodyLines = (rows || []).map((r) => columns.map((c) => escapeCsvCell(r?.[c])).join(','));
  return '\uFEFF' + [headerLine, ...bodyLines].join('\r\n');
}

module.exports = {
  parseTabularUpload,
  parseXlsxBuffer,
  parseDelimitedText,
  parseJsonInput,
  buildMultiSheetXlsx,
  buildUtf8BomCsv,
  excelSerialToIsoDate
};
