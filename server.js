const express    = require('express');
const cors       = require('cors');
const path       = require('path');
const fetch      = (...a) => import('node-fetch').then(({ default: f }) => f(...a));
const ExcelJS    = require('exceljs');
const { google } = require('googleapis');

const app       = express();
const BOT_TOKEN = process.env.BOT_TOKEN || '8653709594:AAGG5Jotx8Oz-NMBnNuIfwGbNqF3o0gV6J8';
const CHAT_ID   = process.env.CHAT_ID   || '64796';
const SHEET_ID  = process.env.SHEET_ID  || '1Fn4d-d-yUare8nOZbhAxt4o-MQWmgV3zPbSSJ9i2vRY';
const PORT      = process.env.PORT      || 3000;

// Google auth via env var (JSON string) or hardcoded for dev
function getAuth() {
  const raw = process.env.GOOGLE_KEY || JSON.stringify({
    type: "service_account",
    project_id: "sendmax-buyer-report",
    private_key_id: "0f7795d7edf567174ad0e5aa14dda834e348e3ac",
    private_key: "-----BEGIN PRIVATE KEY-----\nMIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQCgsIipKdDADp7U\nj9+h51nAhTnpI77z9hNuM1VPGSIKZtsUL34bggiEy2uGb27a4/wpp+dnqgBdo4kL\nurcfcZXAfBLm8pIySNWQRlqoedlbUrdkS4EpxrDXfKI1e0zwA4ZAfNCkl9VCH5Lg\nanPAI168KZqku1Am2qTs60sNZkHTcRnOT91mVEgtEhj1v5dWaKW+T5ChBYHJ+b0v\nyVaNokcjOZfc2dsd2e4DJ/vJSAlIPRft8JN6iFumXOrTdO5p5ZEBWh4Phr9UQ0UT\nq6OWZusXncOiG0IQFVPOQCN1RcttHQ3itP6Fhi1Ad/INTKme/dm+2dY8fpt6DF9P\nfFEUHsb5AgMBAAECggEABAdI9cy5kgcLzlJpDdsJ8fylDgY6ZE1YAVxCGECODOq1\nNVVXeQMSiZm78Vvb0JsEmFrAP2YQqkhOSD//FUe4tGGn66hLJ/2RdRN+dFjRrStu\ns5H9A/oaNtzGCEc2D5JXDEjUxjuUOhsmiH/CtR9t4ZO9b8HOaZrzpYjkkcptfgp+\nqoQ/Hu8/BJkZaOtuVG8fme48z0leueiqKwqBEZt/CCVYPx9SEWzMhBoIBDWVQdmI\nwM+9xvlgPvMiDF/VMalN5O0HehojFZ0ILcvcuVzHxyptTCW57RuXMpFmKhzn0JGW\nc0cNl6UmfXXXDLxpQhUh/5IkE04kqbs4s//r8nWMcQKBgQDNZP8uRyw3sy/HdDqp\n6Z61a9eyCTvPxRpQ2yo5zboakCZDZp6X7Gt4GG+1XFp7/E1y3ZelIUVZLBQ2gxDM\nAvUMyWvAl+bxUJsgfg+f0LirxtDtSawKkmvhy5uOtQ9v8TNHe9jlz05qnyOgCXDJ\nuZ0fMpmKjuj+ZPZi3CqX/EFCfQKBgQDIR9Q/EHPz1LRaOn5YLXVTuNy4fjTkdXbc\nZ9ZonOq5LFYcw5T7F254RQF9FAV5cOVsS1LDO6ItRn7Fh7kelerSFsFFw2jGG1J6\ndsCMskgDLi2GUvowqTV5Jyp3OPTFhLtXBE7Lec15nWHz0IZVnQUZLBfW56se5M1P\nTurlIPsjLQKBgC9kiIUTIx7PF808AnVGVLqj4adQ/h4Y6tORXpLr529Np6aRvTJs\n+PdDHDOK1c4aD3zbqKwPm0LSKu/RGkUYwZZru5M6+azg3oqGy/AiZxory/WUtD/U\nPVRieMha+u7nht+NsiSEyyFLCxtYUlAZpimWsVsJUrWwcjGpf6KGMEC5AoGBAJLQ\nHKUSAX2m2YIeyMkIoa36oEnZshoahUPyFjjfdizBdGpncIlTVtZw9zNh/tVTzcnB\nuWyGBU8dEwbng+sIyHay2YCzStWV57L/5Av/EEzMh/CkWORc1Wb3L+jvZnZ34wBZ\nlw0PP3E7zjb1D9QoYzldmigd8NGXtWzCk+vcblwNAoGAPLXnxv5T8uMLVTJ9CWKg\ncxO9t+zeYcVUY17RxsP+XOGeve7VnX6u82MiPA43IxjIKyMrZ0ubVqXs9TAwlkON\ngYwI13BFq/QTvtKTmaRm5/EcPLjZyTA/B1pg7oM7GfABAoe9wyXiehNcBfxpWNBa\nOxITq0jm1XQ0SHQvM9zFeBE=\n-----END PRIVATE KEY-----\n",
    client_email: "buyer-report@sendmax-buyer-report.iam.gserviceaccount.com",
    client_id: "113366771135886993925",
    auth_uri: "https://accounts.google.com/o/oauth2/auth",
    token_uri: "https://oauth2.googleapis.com/token",
    auth_provider_x509_cert_url: "https://www.googleapis.com/oauth2/v1/certs",
    client_x509_cert_url: "https://www.googleapis.com/robot/v1/metadata/x509/buyer-report%40sendmax-buyer-report.iam.gserviceaccount.com",
    universe_domain: "googleapis.com"
  });
  const creds = typeof raw === 'string' ? JSON.parse(raw) : raw;
  return new google.auth.GoogleAuth({ credentials: creds, scopes: ['https://www.googleapis.com/auth/spreadsheets'] });
}

app.use(cors());
app.use(express.json({ limit: '4mb' }));
app.use(express.static(path.join(__dirname, 'public')));
app.get('/', (_, res) => res.sendFile(path.join(__dirname, 'public', 'index.html')));

// ── Build Excel ──────────────────────────────────────────────────────────────
async function buildExcel(payload) {
  const { buyerName, date, suppliers, globalComment } = payload;
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Buyer Report Bot';

  // ── Sheet 1: Summary ──
  const summary = wb.addWorksheet('Сводка 汇总');
  summary.views = [{ state: 'frozen', ySplit: 3 }];

  // Title row
  summary.mergeCells('A1:L1');
  const titleCell = summary.getCell('A1');
  titleCell.value = `Ежедневный отчёт закупщика / 采购员日报   |   ${buyerName}   |   ${date}`;
  titleCell.font  = { bold: true, size: 13, color: { argb: 'FFFFFFFF' } };
  titleCell.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1A6B3C' } };
  titleCell.alignment = { vertical: 'middle', horizontal: 'center' };
  summary.getRow(1).height = 32;

  // Header row
  const headers = [
    '№', 'Поставщик / 供应商', 'Товар / 商品', 'Категория / 类别',
    'Цена (¥) / 单价', 'MOQ / 最小起订量', 'Срок / 生产周期',
    'Статус / 状态', 'Качество / 质量', 'Надёжность / 可靠性',
    'Гибкость цены / 价格', 'Коммуникация / 沟通'
  ];
  const hRow = summary.getRow(2);
  headers.forEach((h, i) => {
    const cell = hRow.getCell(i + 1);
    cell.value = h;
    cell.font  = { bold: true, size: 10, color: { argb: 'FFFFFFFF' } };
    cell.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2D5A3D' } };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    cell.border = { bottom: { style: 'thin', color: { argb: 'FF1A6B3C' } } };
  });
  summary.getRow(2).height = 36;

  const statusColors = {
    'Договорились': 'FF1A6B3C', 'Заказ размещён': 'FF185FA5',
    'Образцы запрошены': 'FF5B2DA0', 'В процессе': 'FF0A5FA0',
    'Отказали': 'FFB83232', 'Требует уточнения': 'FF8A6200',
  };

  const colWidths = [4, 28, 22, 16, 10, 10, 12, 20, 11, 11, 11, 11];
  colWidths.forEach((w, i) => { summary.getColumn(i + 1).width = w; });

  suppliers.forEach((s, idx) => {
    const row = summary.addRow([
      idx + 1,
      s.name || '—', s.goods || '—', s.category || '—',
      s.price ? Number(s.price) : '—',
      s.moq   ? Number(s.moq)   : '—',
      s.leadTime || '—',
      s.status || '—',
      s.ratings?.[0] ?? '—', s.ratings?.[1] ?? '—',
      s.ratings?.[2] ?? '—', s.ratings?.[3] ?? '—',
    ]);
    row.height = 22;
    const statusKey = (s.status || '').split('/')[0].trim();
    const statusColor = statusColors[statusKey];
    row.eachCell((cell, col) => {
      cell.alignment = { vertical: 'middle', wrapText: col === 2 };
      cell.border = { bottom: { style: 'hair', color: { argb: 'FFD0CFC8' } } };
      if (col === 8 && statusColor) {
        cell.font = { bold: true, color: { argb: statusColor } };
      }
      if (col >= 9 && typeof cell.value === 'number') {
        const stars = '★'.repeat(cell.value) + '☆'.repeat(5 - cell.value);
        cell.value = stars;
        cell.font = { color: { argb: 'FFD4900A' } };
      }
      if (idx % 2 === 1) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF7F6F3' } };
      }
    });
  });

  // Global comment
  summary.addRow([]);
  const gcLabel = summary.addRow(['Общий комментарий / 总体备注:']);
  gcLabel.getCell(1).font = { bold: true, color: { argb: 'FF2D5A3D' } };
  const gcRow = summary.addRow([globalComment || '—']);
  gcRow.getCell(1).alignment = { wrapText: true };
  summary.mergeCells(`A${gcRow.number}:L${gcRow.number}`);
  gcRow.height = 48;

  // ── Sheet 2: Detail per supplier ──
  const detail = wb.addWorksheet('Детали 详情');
  detail.getColumn(1).width = 28;
  detail.getColumn(2).width = 60;

  const accent = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1A6B3C' } };
  const light  = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEAF4EE' } };
  const mid    = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2D5A3D' } };

  let r = 1;
  suppliers.forEach((s, idx) => {
    // Supplier header
    detail.mergeCells(`A${r}:B${r}`);
    const sh = detail.getCell(`A${r}`);
    sh.value = `Поставщик ${idx+1} / 供应商 ${idx+1}: ${s.name || '—'}`;
    sh.font  = { bold: true, size: 12, color: { argb: 'FFFFFFFF' } };
    sh.fill  = accent;
    sh.alignment = { vertical: 'middle' };
    detail.getRow(r).height = 28;
    r++;

    const fields = [
      ['Товар / 商品', s.goods],
      ['Категория / 类别', s.category],
      ['Цена за ед. (¥) / 单价', s.price],
      ['MOQ (шт.) / 最小起订量', s.moq],
      ['Срок производства / 生产周期', s.leadTime],
    ];
    fields.forEach(([label, val]) => {
      const row = detail.addRow([label, val || '—']);
      row.getCell(1).font = { bold: true, size: 10 };
      row.getCell(1).fill = light;
      row.getCell(2).alignment = { wrapText: true };
      row.height = 20;
      r++;
    });

    // Catalogs
    if (s.catalogs?.length) {
      const ch = detail.addRow(['Каталоги / 目录链接', '']);
      ch.getCell(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
      ch.getCell(1).fill = mid;
      detail.mergeCells(`A${r}:B${r}`);
      ch.getCell(1).value = 'Каталоги / 目录链接';
      r++;
      s.catalogs.forEach(c => {
        if (!c.name && !c.url) return;
        const crow = detail.addRow([c.name || '—', c.url || '—']);
        if (c.url) {
          crow.getCell(2).value = { text: c.url, hyperlink: c.url };
          crow.getCell(2).font  = { color: { argb: 'FF185FA5' }, underline: true };
        }
        crow.height = 20;
        r++;
      });
    }

    // Status
    const statusRow = detail.addRow(['Статус / 谈判状态', s.status || '—']);
    statusRow.getCell(1).font = { bold: true }; statusRow.getCell(1).fill = light;
    r++;
    if (s.statusNote) {
      const snRow = detail.addRow(['Комментарий / 备注', s.statusNote]);
      snRow.getCell(1).font = { bold: true }; snRow.getCell(1).fill = light;
      snRow.getCell(2).alignment = { wrapText: true };
      r++;
    }

    // Ratings
    const rLabels = ['Качество / 质量', 'Надёжность / 可靠性', 'Гибкость цены / 价格', 'Коммуникация / 沟通'];
    rLabels.forEach((rl, ri) => {
      const v = s.ratings?.[ri];
      const stars = typeof v === 'number' ? '★'.repeat(v) + '☆'.repeat(5-v) : '—';
      const rrow = detail.addRow([rl, stars]);
      rrow.getCell(1).font = { bold: true }; rrow.getCell(1).fill = light;
      rrow.getCell(2).font = { color: { argb: 'FFD4900A' } };
      r++;
    });

    // Issues
    if (s.issues?.length) {
      const irow = detail.addRow(['Проблемы / 问题', s.issues.join('\n')]);
      irow.getCell(1).font = { bold: true, color: { argb: 'FFB83232' } };
      irow.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFBF0F0' } };
      irow.getCell(2).alignment = { wrapText: true };
      irow.height = Math.max(20, s.issues.length * 16);
      r++;
    }
    if (s.issuesNote) {
      const inrow = detail.addRow(['Описание проблем / 说明', s.issuesNote]);
      inrow.getCell(1).font = { bold: true }; inrow.getCell(1).fill = light;
      inrow.getCell(2).alignment = { wrapText: true }; inrow.height = 36;
      r++;
    }

    // Next steps
    if (s.steps?.length) {
      const steps = s.steps.filter(st => st.text || st.date);
      if (steps.length) {
        const stHead = detail.addRow(['Следующие шаги / 下一步行动', '']);
        stHead.getCell(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
        stHead.getCell(1).fill = mid;
        detail.mergeCells(`A${r}:B${r}`);
        stHead.getCell(1).value = 'Следующие шаги / 下一步行动';
        r++;
        steps.forEach((st, si) => {
          const stRow = detail.addRow([`${si+1}. ${st.text || '—'}`, st.date || '']);
          stRow.height = 20; r++;
        });
      }
    }

    // Spacer
    detail.addRow([]); r++;
  });

  // ── Sheet 3: Supplier DB ──
  const db = wb.addWorksheet('База поставщиков');
  db.views = [{ state: 'frozen', ySplit: 2 }];

  const dbHeaders = [
    'Дата / 日期', 'Закупщик / 采购员', 'Поставщик / 供应商',
    'Товар / 商品', 'Категория / 类别', 'Цена (¥) / 单价',
    'MOQ / 最小起订量', 'Срок / 生产周期', 'Статус / 状态',
    'Качество / 质量', 'Надёжность / 可靠性', 'Гибкость / 价格灵活',
    'Коммуникация / 沟通', 'Проблемы / 问题', 'Каталог 1 / 目录1', 'Каталог 2 / 目录2',
  ];
  db.mergeCells('A1:P1');
  const dbTitle = db.getCell('A1');
  dbTitle.value = 'База поставщиков / 供应商数据库';
  dbTitle.font  = { bold: true, size: 13, color: { argb: 'FFFFFFFF' } };
  dbTitle.fill  = accent;
  dbTitle.alignment = { horizontal: 'center', vertical: 'middle' };
  db.getRow(1).height = 28;

  const dbhRow = db.getRow(2);
  dbHeaders.forEach((h, i) => {
    const cell = dbhRow.getCell(i + 1);
    cell.value = h;
    cell.font  = { bold: true, size: 9, color: { argb: 'FFFFFFFF' } };
    cell.fill  = mid;
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
  });
  db.getRow(2).height = 32;

  const dbWidths = [12,16,24,20,14,9,9,11,18,10,10,10,10,24,28,28];
  dbWidths.forEach((w, i) => { db.getColumn(i + 1).width = w; });

  const ratingLabel = v => typeof v === 'number' ? '★'.repeat(v)+'☆'.repeat(5-v) : '—';

  suppliers.forEach((s, idx) => {
    const cats = (s.catalogs || []).filter(c => c.url);
    const row = db.addRow([
      date, buyerName,
      s.name || '—', s.goods || '—', s.category || '—',
      s.price ? Number(s.price) : '—',
      s.moq   ? Number(s.moq)   : '—',
      s.leadTime || '—',
      s.status || '—',
      ratingLabel(s.ratings?.[0]), ratingLabel(s.ratings?.[1]),
      ratingLabel(s.ratings?.[2]), ratingLabel(s.ratings?.[3]),
      (s.issues || []).join(', ') || '—',
      cats[0]?.url || '—', cats[1]?.url || '—',
    ]);
    row.height = 20;
    if (cats[0]?.url) {
      row.getCell(15).value = { text: cats[0].url, hyperlink: cats[0].url };
      row.getCell(15).font  = { color: { argb: 'FF185FA5' }, underline: true };
    }
    if (cats[1]?.url) {
      row.getCell(16).value = { text: cats[1].url, hyperlink: cats[1].url };
      row.getCell(16).font  = { color: { argb: 'FF185FA5' }, underline: true };
    }
    // Status color
    const statusKey = (s.status || '').split('/')[0].trim();
    const statusColors2 = {
      'Договорились': 'FF1A6B3C', 'Заказ размещён': 'FF185FA5',
      'Образцы запрошены': 'FF5B2DA0', 'В процессе': 'FF0A5FA0',
      'Отказали': 'FFB83232', 'Требует уточнения': 'FF8A6200',
    };
    if (statusColors2[statusKey]) {
      row.getCell(9).font = { bold: true, color: { argb: statusColors2[statusKey] } };
    }
    [10,11,12,13].forEach(col => {
      row.getCell(col).font = { color: { argb: 'FFD4900A' } };
    });
    if (idx % 2 === 1) {
      row.eachCell(cell => {
        if (!cell.fill || cell.fill.fgColor?.argb === 'FF185FA5') return;
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF7F6F3' } };
      });
    }
  });

  db.autoFilter = { from: 'A2', to: `P2` };

  return wb;
}

// ── Send Excel to Telegram ──────────────────────────────────────────────────
async function sendExcelToTelegram(wb, filename, caption) {
  const buf      = await wb.xlsx.writeBuffer();
  const boundary = '----Boundary' + Math.random().toString(36).slice(2);
  let body = '';
  body += `--${boundary}\r\nContent-Disposition: form-data; name="chat_id"\r\n\r\n${CHAT_ID}\r\n`;
  body += `--${boundary}\r\nContent-Disposition: form-data; name="caption"\r\n\r\n${caption}\r\n`;
  body += `--${boundary}\r\nContent-Disposition: form-data; name="parse_mode"\r\n\r\nMarkdown\r\n`;
  const pre        = Buffer.from(body, 'utf8');
  const fileHeader = Buffer.from(`--${boundary}\r\nContent-Disposition: form-data; name="document"; filename="${filename}"\r\nContent-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n`, 'utf8');
  const post       = Buffer.from(`\r\n--${boundary}--\r\n`, 'utf8');
  const multipart  = Buffer.concat([pre, fileHeader, Buffer.from(buf), post]);
  const res = await fetch(`https://api.telegram.org/bot${BOT_TOKEN}/sendDocument`, {
    method: 'POST',
    headers: { 'Content-Type': `multipart/form-data; boundary=${boundary}` },
    body: multipart,
  });
  const json = await res.json();
  if (!json.ok) throw new Error(JSON.stringify(json));
}

// ── Append to Google Sheets DB ──────────────────────────────────────────────
async function appendToSheets(payload) {
  const { buyerName, date, suppliers } = payload;
  const auth    = getAuth();
  const sheets  = google.sheets({ version: 'v4', auth });

  // Ensure sheet "База поставщиков" exists with headers
  const meta = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID });
  const sheetNames = meta.data.sheets.map(s => s.properties.title);

  if (!sheetNames.includes('База поставщиков')) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SHEET_ID,
      requestBody: { requests: [{ addSheet: { properties: { title: 'База поставщиков' } } }] },
    });
    // Add headers
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID, range: 'База поставщиков!A1',
      valueInputOption: 'RAW',
      requestBody: { values: [[
        'Дата', 'Закупщик', 'Поставщик', 'Товар', 'Категория',
        'Цена (¥)', 'MOQ', 'Срок', 'Статус',
        'Качество ★', 'Надёжность ★', 'Гибкость цены ★', 'Коммуникация ★',
        'Проблемы', 'Комментарий', 'Каталог 1', 'Каталог 2',
      ]] },
    });
  }

  // Append rows
  const rows = suppliers.map(s => {
    const cats = (s.catalogs || []).filter(c => c.url);
    return [
      date, buyerName,
      s.name || '', s.goods || '', s.category || '',
      s.price || '', s.moq || '', s.leadTime || '',
      (s.status || '').split('/')[0].trim(),
      s.ratings?.[0] ?? '', s.ratings?.[1] ?? '',
      s.ratings?.[2] ?? '', s.ratings?.[3] ?? '',
      (s.issues || []).join('; ') || '',
      s.statusNote || '',
      cats[0]?.url || '', cats[1]?.url || '',
    ];
  });

  await sheets.spreadsheets.values.append({
    spreadsheetId: SHEET_ID, range: 'База поставщиков!A1',
    valueInputOption: 'RAW', insertDataOption: 'INSERT_ROWS',
    requestBody: { values: rows },
  });
}

// ── Routes ──────────────────────────────────────────────────────────────────
app.post('/send-report', async (req, res) => {
  try {
    const payload = req.body;
    if (!payload?.suppliers?.length) return res.status(400).json({ ok: false, error: 'No data' });

    const { buyerName, date } = payload;
    const filename = `отчёт_${(buyerName||'закупщик').replace(/\s+/g,'_')}_${new Date().toISOString().slice(0,10)}.xlsx`;
    const caption  = `📊 *Отчёт закупщика*\n👤 ${buyerName || '—'}\n📅 ${date || '—'}\n📦 Поставщиков: ${payload.suppliers.length}`;

    const wb = await buildExcel(payload);

    // Run in parallel
    await Promise.all([
      sendExcelToTelegram(wb, filename, caption),
      appendToSheets(payload).catch(e => console.error('Sheets error:', e)),
    ]);

    res.json({ ok: true });
  } catch (e) {
    console.error(e);
    res.status(500).json({ ok: false, error: e.message });
  }
});

app.listen(PORT, () => console.log(`Server on port ${PORT}`));  headers.forEach((h, i) => {
    const cell = hRow.getCell(i + 1);
    cell.value = h;
    cell.font  = { bold: true, size: 10, color: { argb: 'FFFFFFFF' } };
    cell.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2D5A3D' } };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    cell.border = { bottom: { style: 'thin', color: { argb: 'FF1A6B3C' } } };
  });
  summary.getRow(2).height = 36;

  const statusColors = {
    'Договорились': 'FF1A6B3C', 'Заказ размещён': 'FF185FA5',
    'Образцы запрошены': 'FF5B2DA0', 'В процессе': 'FF0A5FA0',
    'Отказали': 'FFB83232', 'Требует уточнения': 'FF8A6200',
  };

  const colWidths = [4, 28, 22, 16, 10, 10, 12, 20, 11, 11, 11, 11];
  colWidths.forEach((w, i) => { summary.getColumn(i + 1).width = w; });

  suppliers.forEach((s, idx) => {
    const row = summary.addRow([
      idx + 1,
      s.name || '—', s.goods || '—', s.category || '—',
      s.price ? Number(s.price) : '—',
      s.moq   ? Number(s.moq)   : '—',
      s.leadTime || '—',
      s.status || '—',
      s.ratings?.[0] ?? '—', s.ratings?.[1] ?? '—',
      s.ratings?.[2] ?? '—', s.ratings?.[3] ?? '—',
    ]);
    row.height = 22;
    const statusKey = (s.status || '').split('/')[0].trim();
    const statusColor = statusColors[statusKey];
    row.eachCell((cell, col) => {
      cell.alignment = { vertical: 'middle', wrapText: col === 2 };
      cell.border = { bottom: { style: 'hair', color: { argb: 'FFD0CFC8' } } };
      if (col === 8 && statusColor) {
        cell.font = { bold: true, color: { argb: statusColor } };
      }
      if (col >= 9 && typeof cell.value === 'number') {
        const stars = '★'.repeat(cell.value) + '☆'.repeat(5 - cell.value);
        cell.value = stars;
        cell.font = { color: { argb: 'FFD4900A' } };
      }
      if (idx % 2 === 1) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF7F6F3' } };
      }
    });
  });

  // Global comment
  summary.addRow([]);
  const gcLabel = summary.addRow(['Общий комментарий / 总体备注:']);
  gcLabel.getCell(1).font = { bold: true, color: { argb: 'FF2D5A3D' } };
  const gcRow = summary.addRow([globalComment || '—']);
  gcRow.getCell(1).alignment = { wrapText: true };
  summary.mergeCells(`A${gcRow.number}:L${gcRow.number}`);
  gcRow.height = 48;

  // ── Sheet 2: Detail per supplier ──
  const detail = wb.addWorksheet('Детали / 详情');
  detail.getColumn(1).width = 28;
  detail.getColumn(2).width = 60;

  const accent = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1A6B3C' } };
  const light  = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEAF4EE' } };
  const mid    = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2D5A3D' } };

  let r = 1;
  suppliers.forEach((s, idx) => {
    // Supplier header
    detail.mergeCells(`A${r}:B${r}`);
    const sh = detail.getCell(`A${r}`);
    sh.value = `Поставщик ${idx+1} / 供应商 ${idx+1}: ${s.name || '—'}`;
    sh.font  = { bold: true, size: 12, color: { argb: 'FFFFFFFF' } };
    sh.fill  = accent;
    sh.alignment = { vertical: 'middle' };
    detail.getRow(r).height = 28;
    r++;

    const fields = [
      ['Товар / 商品', s.goods],
      ['Категория / 类别', s.category],
      ['Цена за ед. (¥) / 单价', s.price],
      ['MOQ (шт.) / 最小起订量', s.moq],
      ['Срок производства / 生产周期', s.leadTime],
    ];
    fields.forEach(([label, val]) => {
      const row = detail.addRow([label, val || '—']);
      row.getCell(1).font = { bold: true, size: 10 };
      row.getCell(1).fill = light;
      row.getCell(2).alignment = { wrapText: true };
      row.height = 20;
      r++;
    });

    // Catalogs
    if (s.catalogs?.length) {
      const ch = detail.addRow(['Каталоги / 目录链接', '']);
      ch.getCell(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
      ch.getCell(1).fill = mid;
      detail.mergeCells(`A${r}:B${r}`);
      ch.getCell(1).value = 'Каталоги / 目录链接';
      r++;
      s.catalogs.forEach(c => {
        if (!c.name && !c.url) return;
        const crow = detail.addRow([c.name || '—', c.url || '—']);
        if (c.url) {
          crow.getCell(2).value = { text: c.url, hyperlink: c.url };
          crow.getCell(2).font  = { color: { argb: 'FF185FA5' }, underline: true };
        }
        crow.height = 20;
        r++;
      });
    }

    // Status
    const statusRow = detail.addRow(['Статус / 谈判状态', s.status || '—']);
    statusRow.getCell(1).font = { bold: true }; statusRow.getCell(1).fill = light;
    r++;
    if (s.statusNote) {
      const snRow = detail.addRow(['Комментарий / 备注', s.statusNote]);
      snRow.getCell(1).font = { bold: true }; snRow.getCell(1).fill = light;
      snRow.getCell(2).alignment = { wrapText: true };
      r++;
    }

    // Ratings
    const rLabels = ['Качество / 质量', 'Надёжность / 可靠性', 'Гибкость цены / 价格', 'Коммуникация / 沟通'];
    rLabels.forEach((rl, ri) => {
      const v = s.ratings?.[ri];
      const stars = typeof v === 'number' ? '★'.repeat(v) + '☆'.repeat(5-v) : '—';
      const rrow = detail.addRow([rl, stars]);
      rrow.getCell(1).font = { bold: true }; rrow.getCell(1).fill = light;
      rrow.getCell(2).font = { color: { argb: 'FFD4900A' } };
      r++;
    });

    // Issues
    if (s.issues?.length) {
      const irow = detail.addRow(['Проблемы / 问题', s.issues.join('\n')]);
      irow.getCell(1).font = { bold: true, color: { argb: 'FFB83232' } };
      irow.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFBF0F0' } };
      irow.getCell(2).alignment = { wrapText: true };
      irow.height = Math.max(20, s.issues.length * 16);
      r++;
    }
    if (s.issuesNote) {
      const inrow = detail.addRow(['Описание проблем / 说明', s.issuesNote]);
      inrow.getCell(1).font = { bold: true }; inrow.getCell(1).fill = light;
      inrow.getCell(2).alignment = { wrapText: true }; inrow.height = 36;
      r++;
    }

    // Next steps
    if (s.steps?.length) {
      const steps = s.steps.filter(st => st.text || st.date);
      if (steps.length) {
        const stHead = detail.addRow(['Следующие шаги / 下一步行动', '']);
        stHead.getCell(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
        stHead.getCell(1).fill = mid;
        detail.mergeCells(`A${r}:B${r}`);
        stHead.getCell(1).value = 'Следующие шаги / 下一步行动';
        r++;
        steps.forEach((st, si) => {
          const stRow = detail.addRow([`${si+1}. ${st.text || '—'}`, st.date || '']);
          stRow.height = 20; r++;
        });
      }
    }

    // Spacer
    detail.addRow([]); r++;
  });

  // ── Sheet 3: Supplier DB ──
  const db = wb.addWorksheet('База поставщиков / 供应商数据库');
  db.views = [{ state: 'frozen', ySplit: 2 }];

  const dbHeaders = [
    'Дата / 日期', 'Закупщик / 采购员', 'Поставщик / 供应商',
    'Товар / 商品', 'Категория / 类别', 'Цена (¥) / 单价',
    'MOQ / 最小起订量', 'Срок / 生产周期', 'Статус / 状态',
    'Качество / 质量', 'Надёжность / 可靠性', 'Гибкость / 价格灵活',
    'Коммуникация / 沟通', 'Проблемы / 问题', 'Каталог 1 / 目录1', 'Каталог 2 / 目录2',
  ];
  db.mergeCells('A1:P1');
  const dbTitle = db.getCell('A1');
  dbTitle.value = 'База поставщиков / 供应商数据库';
  dbTitle.font  = { bold: true, size: 13, color: { argb: 'FFFFFFFF' } };
  dbTitle.fill  = accent;
  dbTitle.alignment = { horizontal: 'center', vertical: 'middle' };
  db.getRow(1).height = 28;

  const dbhRow = db.getRow(2);
  dbHeaders.forEach((h, i) => {
    const cell = dbhRow.getCell(i + 1);
    cell.value = h;
    cell.font  = { bold: true, size: 9, color: { argb: 'FFFFFFFF' } };
    cell.fill  = mid;
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
  });
  db.getRow(2).height = 32;

  const dbWidths = [12,16,24,20,14,9,9,11,18,10,10,10,10,24,28,28];
  dbWidths.forEach((w, i) => { db.getColumn(i + 1).width = w; });

  const ratingLabel = v => typeof v === 'number' ? '★'.repeat(v)+'☆'.repeat(5-v) : '—';

  suppliers.forEach((s, idx) => {
    const cats = (s.catalogs || []).filter(c => c.url);
    const row = db.addRow([
      date, buyerName,
      s.name || '—', s.goods || '—', s.category || '—',
      s.price ? Number(s.price) : '—',
      s.moq   ? Number(s.moq)   : '—',
      s.leadTime || '—',
      s.status || '—',
      ratingLabel(s.ratings?.[0]), ratingLabel(s.ratings?.[1]),
      ratingLabel(s.ratings?.[2]), ratingLabel(s.ratings?.[3]),
      (s.issues || []).join(', ') || '—',
      cats[0]?.url || '—', cats[1]?.url || '—',
    ]);
    row.height = 20;
    if (cats[0]?.url) {
      row.getCell(15).value = { text: cats[0].url, hyperlink: cats[0].url };
      row.getCell(15).font  = { color: { argb: 'FF185FA5' }, underline: true };
    }
    if (cats[1]?.url) {
      row.getCell(16).value = { text: cats[1].url, hyperlink: cats[1].url };
      row.getCell(16).font  = { color: { argb: 'FF185FA5' }, underline: true };
    }
    // Status color
    const statusKey = (s.status || '').split('/')[0].trim();
    const statusColors2 = {
      'Договорились': 'FF1A6B3C', 'Заказ размещён': 'FF185FA5',
      'Образцы запрошены': 'FF5B2DA0', 'В процессе': 'FF0A5FA0',
      'Отказали': 'FFB83232', 'Требует уточнения': 'FF8A6200',
    };
    if (statusColors2[statusKey]) {
      row.getCell(9).font = { bold: true, color: { argb: statusColors2[statusKey] } };
    }
    [10,11,12,13].forEach(col => {
      row.getCell(col).font = { color: { argb: 'FFD4900A' } };
    });
    if (idx % 2 === 1) {
      row.eachCell(cell => {
        if (!cell.fill || cell.fill.fgColor?.argb === 'FF185FA5') return;
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF7F6F3' } };
      });
    }
  });

  db.autoFilter = { from: 'A2', to: `P2` };

  return wb;
}

// ── Send Excel to Telegram ──────────────────────────────────────────────────
async function sendExcelToTelegram(wb, filename, caption) {
  const buf      = await wb.xlsx.writeBuffer();
  const boundary = '----Boundary' + Math.random().toString(36).slice(2);
  let body = '';
  body += `--${boundary}\r\nContent-Disposition: form-data; name="chat_id"\r\n\r\n${CHAT_ID}\r\n`;
  body += `--${boundary}\r\nContent-Disposition: form-data; name="caption"\r\n\r\n${caption}\r\n`;
  body += `--${boundary}\r\nContent-Disposition: form-data; name="parse_mode"\r\n\r\nMarkdown\r\n`;
  const pre        = Buffer.from(body, 'utf8');
  const fileHeader = Buffer.from(`--${boundary}\r\nContent-Disposition: form-data; name="document"; filename="${filename}"\r\nContent-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n`, 'utf8');
  const post       = Buffer.from(`\r\n--${boundary}--\r\n`, 'utf8');
  const multipart  = Buffer.concat([pre, fileHeader, Buffer.from(buf), post]);
  const res = await fetch(`https://api.telegram.org/bot${BOT_TOKEN}/sendDocument`, {
    method: 'POST',
    headers: { 'Content-Type': `multipart/form-data; boundary=${boundary}` },
    body: multipart,
  });
  const json = await res.json();
  if (!json.ok) throw new Error(JSON.stringify(json));
}

// ── Append to Google Sheets DB ──────────────────────────────────────────────
async function appendToSheets(payload) {
  const { buyerName, date, suppliers } = payload;
  const auth    = getAuth();
  const sheets  = google.sheets({ version: 'v4', auth });

  // Ensure sheet "База поставщиков" exists with headers
  const meta = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID });
  const sheetNames = meta.data.sheets.map(s => s.properties.title);

  if (!sheetNames.includes('База поставщиков')) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SHEET_ID,
      requestBody: { requests: [{ addSheet: { properties: { title: 'База поставщиков' } } }] },
    });
    // Add headers
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID, range: 'База поставщиков!A1',
      valueInputOption: 'RAW',
      requestBody: { values: [[
        'Дата', 'Закупщик', 'Поставщик', 'Товар', 'Категория',
        'Цена (¥)', 'MOQ', 'Срок', 'Статус',
        'Качество ★', 'Надёжность ★', 'Гибкость цены ★', 'Коммуникация ★',
        'Проблемы', 'Комментарий', 'Каталог 1', 'Каталог 2',
      ]] },
    });
  }

  // Append rows
  const rows = suppliers.map(s => {
    const cats = (s.catalogs || []).filter(c => c.url);
    return [
      date, buyerName,
      s.name || '', s.goods || '', s.category || '',
      s.price || '', s.moq || '', s.leadTime || '',
      (s.status || '').split('/')[0].trim(),
      s.ratings?.[0] ?? '', s.ratings?.[1] ?? '',
      s.ratings?.[2] ?? '', s.ratings?.[3] ?? '',
      (s.issues || []).join('; ') || '',
      s.statusNote || '',
      cats[0]?.url || '', cats[1]?.url || '',
    ];
  });

  await sheets.spreadsheets.values.append({
    spreadsheetId: SHEET_ID, range: 'База поставщиков!A1',
    valueInputOption: 'RAW', insertDataOption: 'INSERT_ROWS',
    requestBody: { values: rows },
  });
}

// ── Routes ──────────────────────────────────────────────────────────────────
app.post('/send-report', async (req, res) => {
  try {
    const payload = req.body;
    if (!payload?.suppliers?.length) return res.status(400).json({ ok: false, error: 'No data' });

    const { buyerName, date } = payload;
    const filename = `отчёт_${(buyerName||'закупщик').replace(/\s+/g,'_')}_${new Date().toISOString().slice(0,10)}.xlsx`;
    const caption  = `📊 *Отчёт закупщика*\n👤 ${buyerName || '—'}\n📅 ${date || '—'}\n📦 Поставщиков: ${payload.suppliers.length}`;

    const wb = await buildExcel(payload);

    // Run in parallel
    await Promise.all([
      sendExcelToTelegram(wb, filename, caption),
      appendToSheets(payload).catch(e => console.error('Sheets error:', e)),
    ]);

    res.json({ ok: true });
  } catch (e) {
    console.error(e);
    res.status(500).json({ ok: false, error: e.message });
  }
});

app.listen(PORT, () => console.log(`Server on port ${PORT}`));
