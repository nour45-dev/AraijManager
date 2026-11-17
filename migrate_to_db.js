const fs = require('fs');
const path = require('path');
const { AraijDatabase } = require('./db.js');

console.log('=== Starting Master Database Migration ===');

const sheetRaw = JSON.parse(fs.readFileSync('snap_today/sheet_api.json', 'utf8')).students;
const rwStudents = JSON.parse(fs.readFileSync('snap_today/railway_students.json', 'utf8'));

// Delete old sqlite db if exists to start fresh
const dbFile = path.join(__dirname, 'araij.db');
if (fs.existsSync(dbFile)) {
  fs.unlinkSync(dbFile);
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

// 1. Build authentic enrollments and student details from Google Sheet
const sheetStudentsByCode = new Map();

sheetRaw.forEach(s => {
  const code = String(s.code || '').trim();
  if (!code) return;
  if (!sheetStudentsByCode.has(code)) {
    sheetStudentsByCode.set(code, {
      code,
      name: s.name,
      area: s.area || '',
      phone: s.phone || '',
      parentPhone: s.parentPhone || '',
      grade: s.grade || '',
      specialization: s.specialization || '',
      regDate: s.regDate || '',
      subjects: new Map()
    });
  }

  const st = sheetStudentsByCode.get(code);
  if (s.name && s.name.trim() && s.name.length > (st.name || '').length) st.name = s.name.trim();
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
    const rawSub = p.substring(0, slash).trim();
    const sub = canonSub(rawSub);
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

    if (mHissas) {
      mHissas[1].split(',').forEach((v, i) => {
        v = v.trim();
        if (v && v !== '_' && v !== '-' && i < 8) entry.hissas[i] = v;
      });
    }
    if (mOct) {
      mOct[1].split(',').forEach((v, i) => {
        v = v.trim();
        if (v && v !== '_' && v !== '-' && i < 8) entry.oct[i] = v;
      });
    }
    if (mNov) {
      mNov[1].split(',').forEach((v, i) => {
        v = v.trim();
        if (v && v !== '_' && v !== '-' && i < 8) entry.nov[i] = v;
      });
    }
  });
});

// Map railway students for baseline attendance
const rwMap = new Map();
rwStudents.forEach(s => rwMap.set(String(s.code).trim(), s));

let studentsInserted = 0;
let enrollmentsInserted = 0;
let octMarksInserted = 0;
let novMarksInserted = 0;

for (const [code, st] of sheetStudentsByCode.entries()) {
  db.upsertStudent(st);
  studentsInserted++;

  const rw = rwMap.get(code);

  for (const [subName, subData] of st.subjects.entries()) {
    db.setEnrollment(code, subName, subData.teacher);
    enrollmentsInserted++;

    const rwSub = rw && rw.academicSubjects && rw.academicSubjects[subName];
    const rwOct = (rwSub && rwSub.months && rwSub.months['شهر 10 (أكتوبر)']) || Array(8).fill('');
    const rwNov = (rwSub && rwSub.months && rwSub.months['شهر 11 (نوفمبر)']) || Array(8).fill('');

    // 1. October marks:
    for (let i = 0; i < 8; i++) {
      let mark = '';
      if (rwOct[i] && rwOct[i].trim() && !rwOct[i].includes('/15')) {
        mark = rwOct[i].trim();
      } else if (subData.oct[i] && subData.oct[i].trim() && !subData.oct[i].includes('/15')) {
        mark = subData.oct[i].trim();
      }
      if (mark) {
        db.recordAttendance(code, subName, 'شهر 10 (أكتوبر)', i + 1, mark);
        octMarksInserted++;
      }
    }

    // 2. November marks:
    for (let i = 0; i < 8; i++) {
      const h = subData.hissas[i];
      const n = subData.nov[i];
      const r = rwNov[i];
      const mark = (h || n || r || '').trim();
      if (mark) {
        db.recordAttendance(code, subName, 'شهر 11 (نوفمبر)', i + 1, mark);
        novMarksInserted++;
      }
    }
  }
}

console.log(`✅ Inserted ${studentsInserted} students.`);
console.log(`✅ Inserted ${enrollmentsInserted} enrollments.`);
console.log(`✅ Recorded ${octMarksInserted} October marks.`);
console.log(`✅ Recorded ${novMarksInserted} November marks.`);

// Verify via getAllStudents()
const allStudents = db.getAllStudents('شهر 11 (نوفمبر)');
console.log(`\nVerification via db.getAllStudents(): ${allStudents.length} students retrieved.`);

// Compute teacher totals from DB
const teacherCounts = {};
allStudents.forEach(s => {
  for (const [sub, d] of Object.entries(s.academicSubjects)) {
    const t = d.teacher;
    teacherCounts[t] = (teacherCounts[t] || 0) + 1;
  }
});

console.log('\n--- VERIFIED TEACHER COUNTS FROM DB ---');
console.table(Object.entries(teacherCounts).sort((a,b) => b[1] - a[1]).map(([t, count]) => ({ Teacher: t, Students: count })));

// Check sample students
const sampleCodes = ['3222', '3596', '3594', '3595', '3160', '3178'];
console.log('\n--- SAMPLE STUDENTS CHECK ---');
sampleCodes.forEach(code => {
  const s = allStudents.find(x => x.code == code);
  if (s && s.academicSubjects && s.academicSubjects['انجليزي']) {
    console.log(`#${code} ${s.name}:`);
    console.log('  Oct:', s.academicSubjects['انجليزي'].months['شهر 10 (أكتوبر)'].slice(0, 6));
    console.log('  Nov:', s.academicSubjects['انجليزي'].months['شهر 11 (نوفمبر)'].slice(0, 6));
  }
});

// Write to students_db.json and data.js
fs.writeFileSync('students_db.json', JSON.stringify(allStudents, null, 2), 'utf8');
fs.writeFileSync('data.js', `window.STUDENTS_DATA = ${JSON.stringify(allStudents, null, 2)};\n`, 'utf8');
console.log('\n✅ Successfully synced clean DB export to students_db.json and data.js!');
