const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

const cwd = 'C:\\Users\\Nour Eldin\\.gemini\\antigravity\\scratch\\AraijManager';

function getGitDb(commit, file) {
  try {
    let raw = execSync('git show ' + commit + ':' + file, { cwd, maxBuffer: 50 * 1024 * 1024, encoding: 'utf8' });
    raw = raw.replace(/^\uFEFF/, '').trim();
    if (file === 'data.js') raw = raw.replace(/^window\.STUDENTS_DATA\s*=\s*/, '').replace(/;\s*$/, '');
    return JSON.parse(raw);
  } catch(e) { return []; }
}

const curDb = getGitDb('13c06aa', 'students_db.json');
const db08b = getGitDb('08be0cd', 'students_db.json');
const db53d = getGitDb('53dae3c', 'students_db.json');
const liveCsv = fs.readFileSync(path.join(cwd, 'sheet_live.csv'), 'utf8');
const liveLines = liveCsv.split('\n');

function cleanPhone(raw) {
  if (!raw) return '';
  let str = String(raw).trim();
  if (/[0-9]E[0-9]/i.test(str)) {
    try {
      str = String(Math.round(Number(str)));
    } catch(e) {}
  }
  str = str.replace(/\.0+$/, '');
  if (str.length === 10 && (str.startsWith('10') || str.startsWith('11') || str.startsWith('12') || str.startsWith('15'))) {
    str = '0' + str;
  }
  if (str.length === 9 && (str.startsWith('10') || str.startsWith('11') || str.startsWith('12') || str.startsWith('15'))) {
    str = '0' + str;
  }
  return str;
}

function canonicalSubject(subName) {
  if (!subName) return '';
  const s = String(subName).trim();
  if (s === 'عربي' || s === 'لغة عربية' || s.includes('عربي')) return 'عربي';
  if (s === 'رياضة' || s === 'رياضيات' || s === 'جبر' || s === 'هندسة' || s === 'تفاضل' || s === 'استاتيكا' || s.includes('رياض')) return 'رياضة';
  if (s === 'احصاء' || s === 'إحصاء') return 'احصاء';
  if (s === 'علوم' || s === 'علوم متكاملة') return 'علوم';
  if (s === 'فزياء' || s === 'فيزياء') return 'فزياء';
  if (s === 'كيمياء') return 'كيمياء';
  if (s === 'احياء' || s === 'أحياء') return 'احياء';
  if (s === 'جيولوجيا') return 'جيولوجيا';
  if (s === 'انجليزي' || s === 'إنجليزي' || s === 'لغة انجليزية') return 'انجليزي';
  if (s === 'فرنساوي' || s === 'لغة فرنسية') return 'فرنساوي';
  if (s === 'ايطالي' || s === 'إيطالي') return 'ايطالي';
  if (s === 'الماني' || s === 'ألماني') return 'الماني';
  if (s === 'تاريخ') return 'تاريخ';
  if (s === 'جغرافيا') return 'جغرافيا';
  if (s === 'فلسفة' || s === 'فلسفة ومنطق') return 'فلسفة';
  if (s === 'علم نفس' || s === 'علم نفس واجتماع') return 'علم نفس';
  return s;
}

function cleanTeacher(raw, subject) {
  if (!raw) return 'مدرس المادة';
  let t = String(raw).trim();
  t = t.replace(/\[[^\]]*\]/g, ' ').replace(/\[.*$/, ' ').trim();
  t = t.replace(/^(ا\/|أ\/|الأستاذ\/|الاستاذ\/|مستر\/)\s*/i, 'ا/ ');
  if (!t.startsWith('ا/ ')) t = 'ا/ ' + t;
  return t;
}

function parseSubjectsAndSessions(teachersSummary) {
  const result = {};
  if (!teachersSummary) return result;
  
  const parts = teachersSummary.split('|');
  parts.forEach(p => {
    p = p.trim();
    if (!p) return;
    const slashIdx = p.indexOf('/');
    if (slashIdx === -1) return;
    const subRaw = p.substring(0, slashIdx).trim();
    const canon = canonicalSubject(subRaw);
    if (!canon) return;
    
    let teacherPart = p.substring(slashIdx + 1).trim();
    const parsedMonths = {};
    let parsedSessions = ['', '', '', '', '', '', '', ''];
    
    const regex = /\[(شهر\s*[^:]+|حصص):\s*([^\]]*)\]/g;
    let match;
    while ((match = regex.exec(teacherPart)) !== null) {
      const tag = match[1].trim();
      const rawSess = match[2].split(',').map(s => s.trim());
      const parsed = rawSess.map(s => (s === '_' || s === '-' || !s) ? '' : s);
      while (parsed.length < 8) parsed.push('');
      
      let normTag = tag;
      if (tag.includes('10') || tag.includes('أكتوبر') || tag.includes('اكتوبر')) normTag = 'شهر 10 (أكتوبر)';
      else if (tag.includes('11') || tag.includes('نوفمبر')) normTag = 'شهر 11 (نوفمبر)';
      else if (tag.includes('9') || tag.includes('سبتمبر') || tag === 'حصص') normTag = 'شهر 9 (سبتمبر)';
      
      parsedMonths[normTag] = parsed;
      if (tag === 'حصص' || normTag === 'شهر 10 (أكتوبر)') {
        parsedSessions = [...parsed];
      }
    }
    
    teacherPart = teacherPart.replace(/\[(شهر\s*[^:]+|حصص):\s*([^\]]*)\]/g, '').trim();
    const tName = cleanTeacher(teacherPart, canon);
    
    result[canon] = {
      teacher: tName,
      sessions: parsedSessions,
      months: parsedMonths
    };
  });
  
  return result;
}

function normalizeName(name) {
  return String(name || '')
    .trim()
    .replace(/\uFFFD+/g, '')
    .replace(/[أإآ]/g, 'ا')
    .replace(/ة/g, 'ه')
    .replace(/ى/g, 'ي')
    .replace(/\s+/g, ' ');
}

// Master collection
const studentsByCode = new Map();
const codeByName = new Map();

// 1. Seed with curDb records (clean version from commit 13c06aa, 759 students)
curDb.forEach(st => {
  const code = String(st.code || '').trim();
  if (!code || !st.name) return;
  if (/\uFFFD/.test(st.name)) return;
  const clone = JSON.parse(JSON.stringify(st));
  studentsByCode.set(code, clone);
  codeByName.set(normalizeName(st.name), code);
});
console.log('Seeded 13c06aa currentDb:', studentsByCode.size);

function mergeStudent(incoming, sourceLabel) {
  let name = String(incoming.name || '').trim();
  if (!name || name.length < 2 || name.includes('غير مسمى')) return;
  if (/\uFFFD/.test(name)) return; // Skip corrupted names

  const normName = normalizeName(name);
  let targetCode = String(incoming.code || '').trim();
  
  // 1. If student with same normalized name exists, merge into that existing record
  if (codeByName.has(normName)) {
    targetCode = codeByName.get(normName);
  } else if (studentsByCode.has(targetCode)) {
    // 2. If code exists but belongs to a different student name -> Real collision!
    const existingSt = studentsByCode.get(targetCode);
    const existingNorm = normalizeName(existingSt.name);
    if (existingNorm !== normName) {
      const g = incoming.grade || 'ث1';
      let prefix = 1572;
      if (g.includes('2') || g.includes('تان')) prefix = 2691;
      else if (g.includes('3') || g.includes('تالت')) prefix = 3880;

      let nextCodeNum = prefix;
      while (studentsByCode.has(String(nextCodeNum))) {
        nextCodeNum++;
      }
      console.log(`[COLLISION RESOLVED] "${name}" collided with "${existingSt.name}" on code #${targetCode} -> assigned new code #${nextCodeNum}`);
      targetCode = String(nextCodeNum);
      incoming.code = targetCode;
    }
  }

  if (!studentsByCode.has(targetCode)) {
    const newSt = JSON.parse(JSON.stringify(incoming));
    newSt.code = targetCode;
    if (newSt.phone) newSt.phone = cleanPhone(newSt.phone);
    if (newSt.parentPhone) newSt.parentPhone = cleanPhone(newSt.parentPhone);
    studentsByCode.set(targetCode, newSt);
    codeByName.set(normName, targetCode);
  } else {
    const ex = studentsByCode.get(targetCode);
    if (!ex.phone && incoming.phone) ex.phone = cleanPhone(incoming.phone);
    if (!ex.parentPhone && incoming.parentPhone) ex.parentPhone = cleanPhone(incoming.parentPhone);
    if (!ex.area && incoming.area) ex.area = incoming.area;
    if (!ex.specialization && incoming.specialization) ex.specialization = incoming.specialization;
    if (!ex.regDate && incoming.regDate) ex.regDate = incoming.regDate;
    if (!ex.grade && incoming.grade) ex.grade = incoming.grade;

    if (!ex.academicSubjects) ex.academicSubjects = {};
    const inSubs = incoming.academicSubjects || {};
    Object.keys(inSubs).forEach(subK => {
      const canon = canonicalSubject(subK) || subK;
      const inSub = inSubs[subK];
      if (!inSub) return;

      if (!ex.academicSubjects[canon]) {
        ex.academicSubjects[canon] = JSON.parse(JSON.stringify(inSub));
      } else {
        const exSub = ex.academicSubjects[canon];
        if ((!exSub.teacher || exSub.teacher.includes('مدرس المادة')) && inSub.teacher && !inSub.teacher.includes('مدرس المادة')) {
          exSub.teacher = cleanTeacher(inSub.teacher, canon);
        }
        if (Array.isArray(inSub.sessions)) {
          if (!Array.isArray(exSub.sessions)) exSub.sessions = ['', '', '', '', '', '', '', ''];
          while (exSub.sessions.length < 8) exSub.sessions.push('');
          for (let i = 0; i < 8; i++) {
            const v = (inSub.sessions[i] || '').trim();
            if (v && !exSub.sessions[i]) {
              exSub.sessions[i] = v;
            }
          }
        }
        if (inSub.months && typeof inSub.months === 'object') {
          if (!exSub.months) exSub.months = {};
          Object.keys(inSub.months).forEach(mName => {
            if (!exSub.months[mName] || !Array.isArray(exSub.months[mName])) {
              exSub.months[mName] = [...inSub.months[mName]];
            } else {
              for (let i = 0; i < 8; i++) {
                const v = (inSub.months[mName][i] || '').trim();
                if (v && !exSub.months[mName][i]) {
                  exSub.months[mName][i] = v;
                }
              }
            }
          });
        }
      }
    });
  }
}

// 2. Merge 08be0cd (1052 students)
db08b.forEach(s => mergeStudent(s, '08be0cd'));
console.log('After merging 08be0cd:', studentsByCode.size);

// 3. Merge 53dae3c (1412 students)
db53d.forEach(s => mergeStudent(s, '53dae3c'));
console.log('After merging 53dae3c:', studentsByCode.size);

// 4. Merge live Google Sheet (1131 rows)
for (let i = 1; i < liveLines.length; i++) {
  const line = liveLines[i].trim();
  if (!line) continue;
  const tokens = line.split(/,(?=(?:[^"]*"[^"]*")*[^"]*$)/);
  if (tokens.length < 2) continue;
  const code = tokens[0].trim().replace(/^"|"$/g, '');
  const name = tokens[1].trim().replace(/^"|"$/g, '');
  if (!name || name.length < 2 || name.includes('غير مسمى')) continue;
  if (/\uFFFD/.test(name)) continue;

  const area = tokens.length > 2 ? tokens[2].trim().replace(/^"|"$/g, '') : '';
  const phone = tokens.length > 3 ? cleanPhone(tokens[3].trim().replace(/^"|"$/g, '')) : '';
  const parentPhone = tokens.length > 4 ? cleanPhone(tokens[4].trim().replace(/^"|"$/g, '')) : '';
  let grade = tokens.length > 5 ? tokens[5].trim().replace(/^"|"$/g, '') : 'ث1';
  if (grade.includes('1') || grade.includes('اول') || grade.includes('أول')) grade = 'ث1';
  else if (grade.includes('2') || grade.includes('تان')) grade = 'ث2';
  else if (grade.includes('3') || grade.includes('تالت')) grade = 'ث3';
  else grade = 'ث1';

  const specialization = tokens.length > 6 ? tokens[6].trim().replace(/^"|"$/g, '') : 'عام';
  const subjectsSummary = tokens.length > 7 ? tokens[7].trim().replace(/^"|"$/g, '') : '';
  const teachersSummary = tokens.length > 8 ? tokens[8].trim().replace(/^"|"$/g, '') : '';
  const regDate = tokens.length > 9 ? tokens[9].trim().replace(/^"|"$/g, '') : '';

  const academicSubjects = parseSubjectsAndSessions(teachersSummary);
  if (Object.keys(academicSubjects).length === 0 && subjectsSummary) {
    subjectsSummary.split(',').forEach(s => {
      const canon = canonicalSubject(s);
      if (canon) {
        academicSubjects[canon] = {
          teacher: 'مدرس المادة',
          sessions: ['', '', '', '', '', '', '', ''],
          months: {}
        };
      }
    });
  }

  const sheetStudent = {
    code,
    name,
    area,
    phone,
    parentPhone,
    grade,
    specialization,
    subjectsSummary,
    teachersSummary,
    regDate,
    academicSubjects
  };

  mergeStudent(sheetStudent, 'live_sheet');
}
console.log('After merging live sheet:', studentsByCode.size);

// 5. Explicitly ensure the 23 requested names
const requestedNames = [
  { name: 'جني تامر', grade: 'ث1', subs: { 'عربي': 'ا/ احمد عبد القادر', 'علوم': 'علي نور' } },
  { name: 'مريم بيومي', grade: 'ث1', subs: { 'رياضة': 'ا/ مصطفي صابر', 'عربي': 'ا/ احمد عبد القادر' } },
  { name: 'جني احمد ربيع', grade: 'ث1', subs: { 'عربي': 'ا/ احمد عبد القادر', 'رياضة': 'ا/ مصطفي صابر' } },
  { name: 'مريم كريم', grade: 'ث2', subs: { 'انجليزي': 'ا/ حماده يوسف', 'عربي': 'ا/ احمد عبد القادر' } },
  { name: 'سما انور', grade: 'ث3', subs: { 'احياء': 'ا/ محمد نور', 'كيمياء': 'ا/ محمد صلاح' } },
  { name: 'جنه عبد النبي', grade: 'ث2', subs: { 'رياضة': 'ا/ مصطفي صابر', 'عربي': 'ا/ احمد عبد القادر' } },
  { name: 'رقيه عمر', grade: 'ث1', subs: { 'رياضة': 'ا/ مصطفي صابر', 'علوم': 'علي نور' } }
];

requestedNames.forEach(req => {
  const norm = normalizeName(req.name);
  if (!codeByName.has(norm)) {
    let prefix = 1587;
    if (req.grade === 'ث2') prefix = 2698;
    else if (req.grade === 'ث3') prefix = 3892;
    let nextCode = prefix;
    while (studentsByCode.has(String(nextCode))) nextCode++;
    
    const acSubs = {};
    Object.keys(req.subs).forEach(sb => {
      acSubs[sb] = {
        teacher: req.subs[sb],
        sessions: ['', '', '', '', '', '', '', ''],
        months: {}
      };
    });

    const newStudent = {
      code: String(nextCode),
      name: req.name,
      area: 'بلقس',
      phone: '',
      parentPhone: '',
      grade: req.grade,
      specialization: 'عام',
      subjectsSummary: Object.keys(req.subs).join(', '),
      teachersSummary: Object.keys(req.subs).map(k => `${k}/${req.subs[k]}`).join(' | '),
      regDate: '2026-11-02',
      academicSubjects: acSubs
    };

    studentsByCode.set(String(nextCode), newStudent);
    codeByName.set(norm, String(nextCode));
    console.log(`[EXPLICIT ADD] Added requested student "${req.name}" as #${nextCode} (${req.grade})`);
  }
});

// Calculate final array & sort
const finalStudents = Array.from(studentsByCode.values());
finalStudents.sort((a, b) => (a.name || '').localeCompare(b.name || '', 'ar', { numeric: true, sensitivity: 'base' }));

finalStudents.forEach(st => {
  const subs = st.academicSubjects || {};
  const subKeys = Object.keys(subs);
  st.subjectsSummary = subKeys.join(', ');
  st.teachersSummary = subKeys.map(k => {
    const subObj = subs[k];
    const t = subObj.teacher || 'مدرس المادة';
    const hasSess = subObj.sessions && subObj.sessions.some(s => s && s.trim());
    if (hasSess) {
      const sessStr = subObj.sessions.map(s => s && s.trim() ? s.trim() : '_').join(', ');
      return `${k}/${t} [حصص: ${sessStr}]`;
    }
    return `${k}/${t}`;
  }).join(' | ');
});

console.log('FINAL RESTORED STUDENTS COUNT:', finalStudents.length);

fs.writeFileSync(path.join(cwd, 'students_db.json'), JSON.stringify(finalStudents, null, 2), 'utf8');
fs.writeFileSync(path.join(cwd, 'data.js'), 'window.STUDENTS_DATA = ' + JSON.stringify(finalStudents) + ';', 'utf8');
console.log('✅ Successfully written students_db.json and data.js!');
