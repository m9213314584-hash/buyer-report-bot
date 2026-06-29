const express = require('express');
const cors    = require('cors');
const fetch   = (...a) => import('node-fetch').then(({default:f}) => f(...a));

const app        = express();
const BOT_TOKEN  = process.env.BOT_TOKEN  || '8653709594:AAGG5Jotx8Oz-NMBnNuIfwGbNqF3o0gV6J8';
const CHAT_ID    = process.env.CHAT_ID    || '64796';
const PORT       = process.env.PORT       || 3000;

app.use(cors());
app.use(express.json({ limit: '2mb' }));

app.get('/', (_, res) => res.send('Buyer Report Bot is running ✅'));

app.post('/send-report', async (req, res) => {
  try {
    const { name, date, report } = req.body;
    if (!report) return res.status(400).json({ ok: false, error: 'No report' });

    const caption =
      `📋 *Отчёт закупщика*\n` +
      `👤 ${name || 'Не указано'}\n` +
      `📅 ${date || new Date().toLocaleDateString('ru-RU')}`;

    // Отправляем файл как document
    const filename = `отчёт_${(name||'закупщик').replace(/\s+/g,'_')}_${new Date().toISOString().slice(0,10)}.txt`;
    const boundary = '----FormBoundary' + Math.random().toString(36).slice(2);

    // Encode content with BOM for correct Russian/Chinese display
    const content  = '\uFEFF' + report;
    const fileBytes = Buffer.from(content, 'utf8');

    let body = '';
    body += `--${boundary}\r\nContent-Disposition: form-data; name="chat_id"\r\n\r\n${CHAT_ID}\r\n`;
    body += `--${boundary}\r\nContent-Disposition: form-data; name="caption"\r\n\r\n${caption}\r\n`;
    body += `--${boundary}\r\nContent-Disposition: form-data; name="parse_mode"\r\n\r\nMarkdown\r\n`;

    const pre  = Buffer.from(body, 'utf8');
    const post = Buffer.from(`\r\n--${boundary}--\r\n`, 'utf8');
    const fileHeader = Buffer.from(
      `--${boundary}\r\nContent-Disposition: form-data; name="document"; filename="${filename}"\r\nContent-Type: text/plain; charset=utf-8\r\n\r\n`,
      'utf8'
    );

    const multipart = Buffer.concat([pre, fileHeader, fileBytes, post]);

    const tgRes = await fetch(
      `https://api.telegram.org/bot${BOT_TOKEN}/sendDocument`,
      {
        method:  'POST',
        headers: { 'Content-Type': `multipart/form-data; boundary=${boundary}` },
        body:    multipart,
      }
    );
    const tgJson = await tgRes.json();
    if (!tgJson.ok) throw new Error(JSON.stringify(tgJson));

    res.json({ ok: true });
  } catch (e) {
    console.error(e);
    res.status(500).json({ ok: false, error: e.message });
  }
});

app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
