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

assets.forEach(file => {
  const src = path.join(__dirname, file);
  const dest = path.join(wwwDir, file);
  if (fs.existsSync(src)) {
    fs.copyFileSync(src, dest);
  }
});

console.log('✅ www directory prepared successfully with web assets for Capacitor.');
