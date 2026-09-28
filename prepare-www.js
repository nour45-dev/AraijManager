const fs = require('fs');
const path = require('path');

const wwwDir = path.join(__dirname, 'www');
if (!fs.existsSync(wwwDir)) {
  fs.mkdirSync(wwwDir, { recursive: true });
}

const assets = [
  'index.html',
  'app.js',
  'data.js',
  'students_db.json',
  'sw.js',
  'manifest.json',
  'tailwind.min.js',
  'lucide.min.js',
  'chart.min.js',
  'html2pdf.bundle.min.js',
  'icon-192.png',
  'icon-512.png',
  'icon.svg',
  'logo.png',
  'logo.jpg',
  'cairo-bold.ttf',
  'cairo-regular.ttf',
  'cairo-semibold.ttf'
];

const androidPublicDir = path.join(__dirname, 'android', 'app', 'src', 'main', 'assets', 'public');
if (!fs.existsSync(androidPublicDir)) {
  try { fs.mkdirSync(androidPublicDir, { recursive: true }); } catch(e) {}
}

assets.forEach(file => {
  const src = path.join(__dirname, file);
  const dest = path.join(wwwDir, file);
  if (fs.existsSync(src)) {
    fs.copyFileSync(src, dest);
  }
  if (fs.existsSync(androidPublicDir)) {
    const androidDest = path.join(androidPublicDir, file);
    try { fs.copyFileSync(src, androidDest); } catch(e) {}
  }
});

console.log('✅ www & android/public directories prepared successfully with web assets for Capacitor.');
