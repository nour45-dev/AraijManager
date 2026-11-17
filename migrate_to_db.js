const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');
const { AraijDatabase } = require('./db.js');

console.log('=== Rebuilding Master Database (Full Oct & Nov Fidelity) ===');

const SNAP_DIR = 'C:\\Users\\Nour Eldin\\.gemini\\antigravity\\brain\\3567ee56-185d-4728-a309-9248c2e96721\\scratch\\snap';
const sheetRaw = JSON.parse(fs.readFileSync(path.join(SNAP_DIR, 'sheet_api.json'), 'utf8')).students;
const rwStudents = JSON.parse(fs.readFileSync(path.join(SNAP_DIR, 'railway.json'), 'utf8'));
const g5c = JSON.parse(execSync('git show 5cd6bf9:students_db.json', { cwd: __dirname, maxBuffer: 300 * 1024 * 1024 }));
const g941 = JSON.parse(execSync('git show 941308b:students_db.json', { cwd: __dirname, maxBuffer: 300 * 1024 * 1024 }));

const dbFile = path.join(__dirname, 'araij.db');
for (const ext of ['', '-wal', '-shm']) {
  if (fs.existsSync(dbFile + ext)) fs.unlinkSync(dbFile + ext);
}
const db = new AraijDatabase(dbFile);

const norm = s => (s || '').replace(/[أإآ]/g, 'ا').replace(/ة/g, 'ه').replace(/ى/g, 'ي').replace(/[\s\-_]+/g, ' ').trim().toLowerCase();

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
  t = norm(t).replace(/\[.*$/, '').trim().replace(/^(ا|م|أ)\s*\/\s*/, '').trim();
  if (t === 'علي' || t === 'علي نور') return 'ا/ علي نور';
  if (t === 'محمد نور') return 'ا/ محمد نور';
  return 'ا/ ' + t;
}

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

const sheetPeople = new Map();
const usedCodes = new Set();

sheetRaw.forEach(s => {
  const rawCode = String(s.code || '').trim();
  const nName = norm(s.name);
  if (!rawCode || !nName) return;

  const collisionKey = `${rawCode}|${nName}`;
  let finalCode = COLLISION_CODE_MAP[collisionKey] || rawCode;

  if (!sheetPeople.has(collisionKey)) {
    while (usedCodes.has(finalCode)) finalCode = finalCode + '_dup';
    usedCodes.add(finalCode);
    sheetPeople.set(collisionKey, {
      code: finalCode, rawCode, name: s.name.trim(),
      area: s.area || '', phone: s.phone || '', parentPhone: s.parentPhone || '',
      grade: s.grade || '', specialization: s.specialization || '', regDate: s.regDate || '',
      subjects: new Map()
    });
  }

  const st = sheetPeople.get(collisionKey);
  if (s.name && s.name.trim().length > st.name.length) st.name = s.name.trim();
  if (s.phone && s.phone.trim()) st.phone = s.phone.trim();
  if (s.parentPhone && s.parentPhone.trim()) st.parentPhone = s.parentPhone.trim();
  if (s.area && s.area.trim()) st.area = s.area.trim();
  if (s.grade && s.grade.trim()) st.grade = s.grade.trim();
  if (s.specialization && s.specialization.trim()) st.specialization = s.specialization.trim();

  (s.teachersSummary || '').split('|').map(x => x.trim()).filter(Boolean).forEach(p => {
    const slash = p.indexOf('/');
    if (slash === -1) return;
    const sub = canonSub(p.substring(0, slash).trim());
    const teacher = canonTeacher(p.substring(slash + 1));
    if (!sub || !teacher) return;

    if (!st.subjects.has(sub)) {
      st.subjects.set(sub, { teacher, hissas: Array(8).fill(''), oct: Array(8).fill(''), nov: Array(8).fill('') });
    }
    const entry = st.subjects.get(sub);
    entry.teacher = teacher;

    const mH = p.match(/\[حصص:\s*([^\]]+)\]/);
    const mO = p.match(/\[شهر 10(?:\s*\(أكتوبر\))?:\s*([^\]]+)\]/);
    const mN = p.match(/\[شهر 11(?:\s*\(نوفمبر\))?:\s*([^\]]+)\]/);
    if (mH) mH[1].split(',').forEach((v, i) => { v = v.trim(); if (v && v !== '_' && v !== '-' && i < 8) entry.hissas[i] = v; });
    if (mO) mO[1].split(',').forEach((v, i) => { v = v.trim(); if (v && v !== '_' && v !== '-' && i < 8) entry.oct[i] = v; });
    if (mN) mN[1].split(',').forEach((v, i) => { v = v.trim(); if (v && v !== '_' && v !== '-' && i < 8) entry.nov[i] = v; });
  });
});

const idxMap = (arr) => {
  const byCode = new Map(), byName = new Map();
  arr.forEach(s => { byCode.set(String(s.code).trim(), s); if (s.name) byName.set(norm(s.name), s); });
  return { byCode, byName };
};
const m5c = idxMap(g5c);
const m941 = idxMap(g941);
const mRw = idxMap(rwStudents);

const findSub = (mapObj, code, name, subName) => {
  const s = mapObj.byCode.get(code) || mapObj.byName.get(norm(name));
  if (!s || !s.academicSubjects) return null;
  if (s.academicSubjects[subName]) return s.academicSubjects[subName];
  for (const [k, v] of Object.entries(s.academicSubjects)) {
    if (canonSub(k) === subName) return v;
  }
  return null;
};

let studentsInserted = 0, enrollmentsInserted = 0, octMarks = 0, novMarks = 0;

for (const st of sheetPeople.values()) {
  db.upsertStudent(st);
  studentsInserted++;

  for (const [subName, subData] of st.subjects.entries()) {
    db.setEnrollment(st.code, subName, subData.teacher);
    enrollmentsInserted++;

    const sub5 = findSub(m5c, st.code, st.name, subName);
    const sub941 = findSub(m941, st.code, st.name, subName);
    const subRw = findSub(mRw, st.code, st.name, subName);

    const o5 = (sub5 && sub5.months && sub5.months['شهر 10 (أكتوبر)']) || Array(8).fill('');
    const rwOct = (subRw && subRw.months && subRw.months['شهر 10 (أكتوبر)']) || Array(8).fill('');
    const n941 = (sub941 && sub941.months && sub941.months['شهر 11 (نوفمبر)']) || Array(8).fill('');
    const rwNov = (subRw && subRw.months && subRw.months['شهر 11 (نوفمبر)']) || Array(8).fill('');

    // Check if sheet [شهر 10] was contaminated by November (e.g. Hamada /15 where 5cd6bf9 had ✓/غ)
    const sheetOctIsContaminated = subData.oct.some((v, i) => v && v.includes('/15') && o5[i] && !o5[i].includes('/15'));

    // 1. OCTOBER: 5cd6bf9 (pure Oct 18 baseline) > sheet [شهر 10] (if not contaminated) > rwOct
    for (let i = 0; i < 8; i++) {
      let mark = (o5[i] || '').trim();
      if (!mark) {
        if (subData.oct[i] && (!sheetOctIsContaminated || !subData.oct[i].includes('/15'))) {
          mark = subData.oct[i].trim();
        } else if (rwOct[i] && !rwOct[i].includes('/15')) {
          mark = rwOct[i].trim();
        }
      }
      if (mark) {
        db.recordAttendance(st.code, subName, 'شهر 10 (أكتوبر)', i + 1, mark);
        octMarks++;
      }
    }

    // 2. NOVEMBER:
    // Use sheet [شهر 11] as primary; fill missing slots from [حصص] ONLY if [حصص] is NOT pure October
    const hissasMatchesTrueOct = subData.hissas.some(Boolean) && subData.hissas.every((v, i) => !v || v === (o5[i] || subData.oct[i] || ''));
    for (let i = 0; i < 8; i++) {
      const n = (subData.nov[i] || '').trim();
      const h = (hissasMatchesTrueOct && !sheetOctIsContaminated) ? '' : (subData.hissas[i] || '').trim();
      const r941 = (n941[i] || '').trim();
      const rRw = (rwNov[i] || '').trim();
      const mark = n || h || rRw || r941 || '';
      if (mark) {
        db.recordAttendance(st.code, subName, 'شهر 11 (نوفمبر)', i + 1, mark);
        novMarks++;
      }
    }
  }
}

console.log(`✅ Students: ${studentsInserted}, Enrollments: ${enrollmentsInserted}, Oct marks: ${octMarks}, Nov marks: ${novMarks}`);

const allStudents = db.getAllStudents('شهر 11 (نوفمبر)');
fs.writeFileSync('students_db.json', JSON.stringify(allStudents, null, 2), 'utf8');
fs.writeFileSync('data.js', `window.STUDENTS_DATA = ${JSON.stringify(allStudents, null, 2)};\n`, 'utf8');
console.log('✅ Saved araij.db, students_db.json, and data.js!');
