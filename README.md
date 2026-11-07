# Soạn sách

```
npm install   # chỉ cần chạy 1 lần
npm start     # mở http://localhost:3000
```

- `npm run build` → build ra thư mục `dist/`
- `npm run preview` → chạy thử bản build
- Đánh số trang PDF: `/api/number-pdf` (dev: `vite.config.js`, Netlify: `netlify/functions/number-pdf.js`)

## Deploy Netlify
Đẩy code lên GitHub → Netlify "Import from Git". Cấu hình đã có sẵn trong `netlify.toml`.
