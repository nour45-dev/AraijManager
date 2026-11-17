const fs = require('fs');
const path = require('path');
const { AraijDatabase } = require('./db.js');

console.log('=== Rebuilding Master Database (Resolving Code Collisions) ===');

const SNAP_DIR = 'C:\\Users\\Nour Eldin\\.gemini\\antigravity\\brain\\3567ee56-185d-4728-a309-9248c2e96721\\scratch\\snap';
const sheetRaw = JSON.parse(fs.readFileSync(path.join(SNAP_DIR, 'sheet_api.json'), 'utf8')).students;
const rwStudents = JSON.parse(fs.readFileSync(path.join(SNAP_DIR, 'railway.json'), 'utf8'));

// Delete old sqlite db + wal/shm if exist
const dbFile = path.join(__dirname, 'araij.db');
for (const ext of ['', '-wal', '-shm']) {
  if (fs.existsSync(dbFile + ext)) fs.unlinkSync(dbFile + ext);
}

const db = new AraijDatabase(dbFile);

function norm(s) {
  return (s || '')
    .replace(/[أإآ]/g, 'ا')
    .replace(/[ة]/g, 'ه')
    .replace(/[ى]/g, 'ي')
    .replace(/[\s\-_]+/g, ' ')
    .trim()
    .toLowerCase();
}

function canonSub(n) {
  n = norm(n);
  if (n.includes('احص') || n.includes('إحص')) return 'احصاء';
  if (n.includes('فيز') || n.includes('فزي')) return 'فزياء';
  if (n.includes('احي') || n.includes('أحي')) return 'احياء';
  if (n.includes('كيم')) return 'كيمياء';
  if (n.includes('رياض')) return 'رياضة';
  if (n.includes('عرب')) return 'عربي';
  if (n.includes('انجل') || n.includes('إنجل')) return 'انجليزي';
  if (n.includes('فرنس')) return 'فرنساوي';
  if (n.includes('تار')) return 'تاريخ';
  if (n.includes('جغراف')) return 'جغرافيا';
  if (n.includes('فلسف')) return 'فلسفة';
  if (n.includes('علوم')) return 'علوم';
  return n;
}

function canonTeacher(t) {
  t = norm(t).replace(/\[.*$/, '').trim();
  t = t.replace(/^(ا|م|أ)\s*\/\s*/, '').trim();
  if (t === 'علي' || t === 'علي نور') return 'ا/ علي نور';
  if (t === 'محمد نور') return 'ا/ محمد نور';
  return 'ا/ ' + t;
}

// Known unique code resolutions for the 8 colliding sheet codes
const COLLISION_CODE_MAP = {
  '1566|انس رضا محمد': '2566',
  '1566|حماده اسماعيل اسماعيل': '2568',
  '1567|مهاب رضا محمود': '2567',
  '1567|منه عماد': '2570',
  '1568|ملك ابراهيم محمد': '2576',
  '1569|ملك ابراهيم حامد': '2577',
  '1570|رويدا سمير عبد الحمديد': '2579',
  '2684|فرحه سيد عواد': '3686',
  '2684|ملك حسن عويس': '3695',
  '2685|علي احمد عبد الرحمن': '3687',
  '2685|ياسمين هشام': '3696',
  '2686|حسن رامي حسن': '3694'
};

// Track distinct people by code + norm(name)
const sheetPeople = new Map();
const usedCodes = new Set();

sheetRaw.forEach(s => {
  const rawCode = String(s.code || '').trim();
  const nName = norm(s.name);
  if (!rawCode || !nName) return;

  const collisionKey = `${rawCode}|${nName}`;
  const personKey = collisionKey;
  let finalCode = COLLISION_CODE_MAP[collisionKey] || rawCode;

  if (!sheetPeople.has(personKey)) {
    // Ensure finalCode isn't already claimed by a different person
    while (usedCodes.has(finalCode)) {
      finalCode = finalCode + '_dup';
    }
    usedCodes.add(finalCode);

    sheetPeople.set(personKey, {
      code: finalCode,
      rawCode,
      name: s.name.trim(),
      area: s.area || '',
      phone: s.phone || '',
      parentPhone: s.parentPhone || '',
      grade: s.grade || '',
      specialization: s.specialization || '',
      regDate: s.regDate || '',
      subjects: new Map()
    });
  }

  const st = sheetPeople.get(personKey);
  if (s.name && s.name.trim().length > st.name.length) st.name = s.name.trim();
  if (s.phone && s.phone.trim()) st.phone = s.phone.trim();
  if (s.parentPhone && s.parentPhone.trim()) st.parentPhone = s.parentPhone.trim();
  if (s.area && s.area.trim()) st.area = s.area.trim();
  if (s.grade && s.grade.trim()) st.grade = s.grade.trim();
  if (s.specialization && s.specialization.trim()) st.specialization = s.specialization.trim();

  const ts = s.teachersSummary || '';
  const parts = ts.split('|').map(x => x.trim()).filter(Boolean);
  parts.forEach(p => {
    const slash = p.indexOf('/');
    if (slash === -1) return;
    const sub = canonSub(p.substring(0, slash).trim());
    const teacher = canonTeacher(p.substring(slash + 1));
    if (!sub || !teacher) return;

    if (!st.subjects.has(sub)) {
      st.subjects.set(sub, {
        teacher,
        hissas: Array(8).fill(''),
        oct: Array(8).fill(''),
        nov: Array(8).fill('')
      });
    }
    const entry = st.subjects.get(sub);
    entry.teacher = teacher;

    const mHissas = p.match(/\[حصص:\s*([^\]]+)\]/);
    const mOct = p.match(/\[شهر 10(?:\s*\(أكتوبر\))?:\s*([^\]]+)\]/);
    const mNov = p.match(/\[شهر 11(?:\s*\(نوفمبر\))?:\s*([^\]]+)\]/);

    if (mHissas) mHissas[1].split(',').forEach((v, i) => { v = v.trim(); if (v && v !== '_' && v !== '-' && i < 8) entry.hissas[i] = v; });
    if (mOct) mOct[1].split(',').forEach((v, i) => { v = v.trim(); if (v && v !== '_' && v !== '-' && i < 8) entry.oct[i] = v; });
    if (mNov) mNov[1].split(',').forEach((v, i) => { v = v.trim(); if (v && v !== '_' && v !== '-' && i < 8) entry.nov[i] = v; });
  });
});

// Map railway students by code AND by norm(name)
const rwByCode = new Map();
const rwByName = new Map();
rwStudents.forEach(s => {
  rwByCode.set(String(s.code).trim(), s);
  if (s.name) rwByName.set(norm(s.name), s);
});

let studentsInserted = 0, enrollmentsInserted = 0, octMarks = 0, novMarks = 0;

for (const st of sheetPeople.values()) {
  db.upsertStudent(st);
  studentsInserted++;

  const rw = rwByCode.get(st.code) || rwByName.get(norm(st.name));

  for (const [subName, subData] of st.subjects.entries()) {
    db.setEnrollment(st.code, subName, subData.teacher);
    enrollmentsInserted++;

    const rwSub = rw && rw.academicSubjects && rw.academicSubjects[subName];
    const rwOct = (rwSub && rwSub.months && rwSub.months['شهر 10 (أكتوبر)']) || Array(8).fill('');
    const rwNov = (rwSub && rwSub.months && rwSub.months['شهر 11 (نوفمبر)']) || Array(8).fill('');

    // October: pristine baseline from rwOct (if not /15) or sheet oct
    for (let i = 0; i < 8; i++) {
      let mark = '';
      if (rwOct[i] && rwOct[i].trim() && !rwOct[i].includes('/15')) {
        mark = rwOct[i].trim();
      } else if (subData.oct[i] && subData.oct[i].trim() && !subData.oct[i].includes('/15')) {
        mark = subData.oct[i].trim();
      }
      if (mark) {
        db.recordAttendance(st.code, subName, 'شهر 10 (أكتوبر)', i + 1, mark);
        octMarks++;
      }
    }

    // November: sheet nov + sheet hissas (only if not identical to oct, or if nov is empty) + rwNov
    const hissasIsOct = subData.hissas.some(Boolean) && subData.oct.some(Boolean) && subData.hissas.every((v, i) => v === subData.oct[i]);
    for (let i = 0; i < 8; i++) {
      const n = subData.nov[i];
      const h = hissasIsOct ? '' : subData.hissas[i];
      const r = rwNov[i];
      const mark = (n || h || r || '').trim();
      if (mark) {
        db.recordAttendance(st.code, subName, 'شهر 11 (نوفمبر)', i + 1, mark);
        novMarks++;
      }
    }
  }
}

console.log(`✅ Students: ${studentsInserted}, Enrollments: ${enrollmentsInserted}, Oct marks: ${octMarks}, Nov marks: ${novMarks}`);

const allStudents = db.getAllStudents('شهر 11 (نوفمبر)');
const byTG = {};
allStudents.forEach(s => {
  Object.entries(s.academicSubjects || {}).forEach(([sub, d]) => {
    const k = `${d.teacher} | ${s.grade}`;
    byTG[k] = (byTG[k] || 0) + 1;
  });
});
console.log('\n--- VERIFIED TEACHER+GRADE COUNTS ---');
Object.entries(byTG).sort().forEach(([k, v]) => console.log(k.padEnd(32), v));

fs.writeFileSync('students_db.json', JSON.stringify(allStudents, null, 2), 'utf8');
fs.writeFileSync('data.js', `window.STUDENTS_DATA = ${JSON.stringify(allStudents, null, 2)};\n`, 'utf8');
console.log('\n✅ Saved araij.db, students_db.json, and data.js!');
