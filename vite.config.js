import { defineConfig } from 'vite';
import { PDFDocument, StandardFonts, rgb } from 'pdf-lib';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.dirname(fileURLToPath(import.meta.url));
const lucideDir = path.join(root, 'node_modules/lucide/dist/umd');

// Đọc body dạng raw
const readBody = (req) => new Promise((resolve, reject) => {
  const chunks = [];
  req.on('data', (c) => chunks.push(c));
  req.on('end', () => resolve(Buffer.concat(chunks)));
  req.on('error', reject);
});

// Đánh số trang PDF: trang lẻ -> góc phải, trang chẵn -> góc trái
async function numberPdf(req, res) {
  try {
    const q = new URL(req.url, 'http://x').searchParams;
    const start = parseInt(q.get('start')) || 1;
    const skip = parseInt(q.get('skip')) || 0;
    const size = parseInt(q.get('size')) || 11;
    const margin = parseFloat(q.get('margin')) || 36;
    const doc = await PDFDocument.load(await readBody(req));
    const font = await doc.embedFont(StandardFonts.TimesRoman);
    doc.getPages().forEach((p, i) => {
      if (i < skip) return;
      const n = start + i - skip;
      const t = String(n), w = font.widthOfTextAtSize(t, size);
      const x = n % 2 === 1 ? p.getWidth() - margin - w : margin;
      p.drawText(t, { x, y: margin * 0.7, size, font, color: rgb(0, 0, 0) });
    });
    res.setHeader('Content-Type', 'application/pdf');
    res.end(Buffer.from(await doc.save()));
  } catch (e) {
    res.statusCode = 400;
    res.end('Không đọc được PDF: ' + e.message);
  }
}

// Plugin: API /api/number-pdf + phục vụ lucide tại /vendor/lucide
function soanSach() {
  const attach = (server) => {
    server.middlewares.use('/api/number-pdf', (req, res, next) =>
      req.method === 'POST' ? numberPdf(req, res) : next());
    server.middlewares.use('/vendor/lucide', (req, res, next) => {
      const f = path.join(lucideDir, path.basename(req.url.split('?')[0]));
      if (fs.existsSync(f) && fs.statSync(f).isFile()) {
        res.setHeader('Content-Type', 'application/javascript');
        fs.createReadStream(f).pipe(res);
      } else next();
    });
  };
  return {
    name: 'soan-sach',
    configureServer: attach,
    configurePreviewServer: attach,
    closeBundle() {
      const out = path.join(root, 'dist/vendor/lucide');
      fs.mkdirSync(out, { recursive: true });
      fs.copyFileSync(path.join(lucideDir, 'lucide.min.js'), path.join(out, 'lucide.min.js'));
    },
  };
}

export default defineConfig({
  plugins: [soanSach()],
  server: { port: 3000 },
  preview: { port: 3000 },
});
