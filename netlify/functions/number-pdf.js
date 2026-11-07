import { PDFDocument, StandardFonts, rgb } from 'pdf-lib';

// Đánh số trang PDF: trang lẻ -> góc phải, trang chẵn -> góc trái
export default async (req) => {
  if (req.method !== 'POST') return new Response('Method Not Allowed', { status: 405 });
  try {
    const q = new URL(req.url).searchParams;
    const start = parseInt(q.get('start')) || 1;
    const skip = parseInt(q.get('skip')) || 0;
    const size = parseInt(q.get('size')) || 11;
    const margin = parseFloat(q.get('margin')) || 36;
    const doc = await PDFDocument.load(new Uint8Array(await req.arrayBuffer()));
    const font = await doc.embedFont(StandardFonts.TimesRoman);
    doc.getPages().forEach((p, i) => {
      if (i < skip) return;
      const n = start + i - skip;
      const t = String(n), w = font.widthOfTextAtSize(t, size);
      const x = n % 2 === 1 ? p.getWidth() - margin - w : margin;
      p.drawText(t, { x, y: margin * 0.7, size, font, color: rgb(0, 0, 0) });
    });
    return new Response(await doc.save(), { headers: { 'Content-Type': 'application/pdf' } });
  } catch (e) {
    return new Response('Không đọc được PDF: ' + e.message, { status: 400 });
  }
};

export const config = { path: '/api/number-pdf' };
