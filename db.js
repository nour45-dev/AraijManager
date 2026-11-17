const fs = require('fs');
const path = require('path');

let DatabaseSync = null;
try {
  const sqlite = require('node:sqlite');
  DatabaseSync = sqlite.DatabaseSync;
  console.log('[DB] Native node:sqlite engine is available.');
} catch (e) {
  console.warn('[DB] Native node:sqlite not available in this Node runtime, using JSON storage engine.');
}

class AraijDatabase {
  constructor(dbPath) {
    this.dbPath = dbPath || path.join(__dirname, 'araij.db');
    this.jsonFile = path.join(__dirname, 'students_db.json');
    this.dataJsFile = path.join(__dirname, 'data.js');
    this.hasSqlite = false;

    if (DatabaseSync) {
      try {
        const dir = path.dirname(this.dbPath);
        if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
        this.sqlite = new DatabaseSync(this.dbPath);
        this.initSchema();
        this.hasSqlite = true;
        console.log(`[DB] Connected to SQLite database at: ${this.dbPath}`);
      } catch (err) {
        console.error('[DB] SQLite initialization failed, falling back to JSON:', err.message);
        this.hasSqlite = false;
      }
    }
  }

  initSchema() {
    if (!this.sqlite) return;
    this.sqlite.exec(`
      PRAGMA journal_mode = WAL;
      PRAGMA synchronous = NORMAL;

      CREATE TABLE IF NOT EXISTS students (
        code TEXT PRIMARY KEY,
        name TEXT NOT NULL,
        area TEXT,
        phone TEXT,
        parentPhone TEXT,
        grade TEXT,
        specialization TEXT,
        regDate TEXT,
        createdAt TEXT,
        updatedAt TEXT
      );

      CREATE TABLE IF NOT EXISTS enrollments (
        studentCode TEXT NOT NULL,
        subject TEXT NOT NULL,
        teacher TEXT NOT NULL,
        createdAt TEXT,
        PRIMARY KEY (studentCode, subject),
        FOREIGN KEY (studentCode) REFERENCES students(code) ON DELETE CASCADE
      );

      CREATE TABLE IF NOT EXISTS attendance (
        studentCode TEXT NOT NULL,
        subject TEXT NOT NULL,
        month TEXT NOT NULL,
        sessionNum INTEGER NOT NULL,
        mark TEXT NOT NULL,
        updatedAt TEXT,
        PRIMARY KEY (studentCode, subject, month, sessionNum),
        FOREIGN KEY (studentCode) REFERENCES students(code) ON DELETE CASCADE
      );

      CREATE TABLE IF NOT EXISTS audit_logs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        timestamp INTEGER NOT NULL,
        action TEXT NOT NULL,
        studentCode TEXT,
        studentName TEXT,
        details TEXT,
        author TEXT,
        createdAt TEXT
      );

      CREATE INDEX IF NOT EXISTS idx_enrollments_teacher ON enrollments(teacher);
      CREATE INDEX IF NOT EXISTS idx_enrollments_subject ON enrollments(subject);
      CREATE INDEX IF NOT EXISTS idx_attendance_lookup ON attendance(studentCode, subject, month);
    `);
  }

  upsertStudent(s) {
    if (!s || !s.code) return;
    if (this.hasSqlite) {
      try {
        const now = new Date().toISOString();
        const stmt = this.sqlite.prepare(`
          INSERT INTO students (code, name, area, phone, parentPhone, grade, specialization, regDate, createdAt, updatedAt)
          VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
          ON CONFLICT(code) DO UPDATE SET
            name = excluded.name,
            area = excluded.area,
            phone = excluded.phone,
            parentPhone = excluded.parentPhone,
            grade = excluded.grade,
            specialization = excluded.specialization,
            regDate = excluded.regDate,
            updatedAt = excluded.updatedAt
        `);
        stmt.run(
          String(s.code || '').trim(),
          (s.name || '').trim(),
          (s.area || '').trim(),
          (s.phone || '').trim(),
          (s.parentPhone || '').trim(),
          (s.grade || 'ث1').trim(),
          (s.specialization || 'عام').trim(),
          (s.regDate || '').trim(),
          s.createdAt || now,
          now
        );
      } catch (err) {
        console.error('[DB] Error in upsertStudent:', err.message);
      }
    }
  }

  setEnrollment(studentCode, subject, teacher) {
    if (!studentCode || !subject || !teacher) return;
    if (this.hasSqlite) {
      try {
        const now = new Date().toISOString();
        const stmt = this.sqlite.prepare(`
          INSERT INTO enrollments (studentCode, subject, teacher, createdAt)
          VALUES (?, ?, ?, ?)
          ON CONFLICT(studentCode, subject) DO UPDATE SET
            teacher = excluded.teacher
        `);
        stmt.run(String(studentCode).trim(), subject.trim(), teacher.trim(), now);
      } catch (err) {
        console.error('[DB] Error in setEnrollment:', err.message);
      }
    }
  }

  deleteEnrollment(studentCode, subject) {
    if (!studentCode || !subject) return;
    if (this.hasSqlite) {
      try {
        this.sqlite.prepare(`DELETE FROM enrollments WHERE studentCode = ? AND subject = ?`).run(String(studentCode).trim(), subject.trim());
        this.sqlite.prepare(`DELETE FROM attendance WHERE studentCode = ? AND subject = ?`).run(String(studentCode).trim(), subject.trim());
      } catch (err) {
        console.error('[DB] Error in deleteEnrollment:', err.message);
      }
    }
  }

  recordAttendance(studentCode, subject, month, sessionNum, mark) {
    if (!studentCode || !subject || !month || !sessionNum) return;
    if (this.hasSqlite) {
      try {
        const now = new Date().toISOString();
        if (!mark || !mark.trim() || mark.trim() === '_' || mark.trim() === '-') {
          const delStmt = this.sqlite.prepare(`
            DELETE FROM attendance WHERE studentCode = ? AND subject = ? AND month = ? AND sessionNum = ?
          `);
          delStmt.run(String(studentCode).trim(), subject.trim(), month.trim(), Number(sessionNum));
          return;
        }

        const stmt = this.sqlite.prepare(`
          INSERT INTO attendance (studentCode, subject, month, sessionNum, mark, updatedAt)
          VALUES (?, ?, ?, ?, ?, ?)
          ON CONFLICT(studentCode, subject, month, sessionNum) DO UPDATE SET
            mark = excluded.mark,
            updatedAt = excluded.updatedAt
        `);
        stmt.run(String(studentCode).trim(), subject.trim(), month.trim(), Number(sessionNum), mark.trim(), now);
      } catch (err) {
        console.error('[DB] Error in recordAttendance:', err.message);
      }
    }
  }

  deleteStudent(code) {
    if (!code) return;
    if (this.hasSqlite) {
      try {
        this.sqlite.prepare(`DELETE FROM students WHERE code = ?`).run(String(code).trim());
        this.sqlite.prepare(`DELETE FROM enrollments WHERE studentCode = ?`).run(String(code).trim());
        this.sqlite.prepare(`DELETE FROM attendance WHERE studentCode = ?`).run(String(code).trim());
      } catch (err) {
        console.error('[DB] Error in deleteStudent:', err.message);
      }
    }
  }

  syncFullStudent(st) {
    if (!st || !st.code) return;
    this.upsertStudent(st);
    if (st.academicSubjects && typeof st.academicSubjects === 'object') {
      for (const [subName, subData] of Object.entries(st.academicSubjects)) {
        if (!subData || !subData.teacher) continue;
        this.setEnrollment(st.code, subName, subData.teacher);
        if (subData.months && typeof subData.months === 'object') {
          for (const [mName, mSessions] of Object.entries(subData.months)) {
            if (Array.isArray(mSessions)) {
              mSessions.forEach((val, idx) => {
                if (val && String(val).trim()) {
                  this.recordAttendance(st.code, subName, mName, idx + 1, String(val).trim());
                }
              });
            }
          }
        }
      }
    }
  }

  getAllStudents(activeMonth = 'شهر 11 (نوفمبر)') {
    if (this.hasSqlite) {
      try {
        const studentsRows = this.sqlite.prepare(`SELECT * FROM students ORDER BY code ASC`).all();
        if (studentsRows && studentsRows.length > 500) {
          const enrollmentsRows = this.sqlite.prepare(`SELECT * FROM enrollments`).all();
          const attendanceRows = this.sqlite.prepare(`SELECT * FROM attendance`).all();

          const enrollmentsByCode = new Map();
          enrollmentsRows.forEach(e => {
            if (!enrollmentsByCode.has(e.studentCode)) enrollmentsByCode.set(e.studentCode, []);
            enrollmentsByCode.get(e.studentCode).push(e);
          });

          const attMap = new Map();
          attendanceRows.forEach(a => {
            const k = `${a.studentCode}|${a.subject}|${a.month}`;
            if (!attMap.has(k)) attMap.set(k, Array(8).fill(''));
            const arr = attMap.get(k);
            const idx = a.sessionNum - 1;
            if (idx >= 0 && idx < 8) {
              arr[idx] = a.mark;
            }
          });

          const result = [];

          studentsRows.forEach(st => {
            const code = st.code;
            const enrolled = enrollmentsByCode.get(code) || [];
            const academicSubjects = {};
            const subNames = [];
            const teacherParts = [];

            enrolled.forEach(e => {
              const subName = e.subject;
              const teacher = e.teacher;
              subNames.push(subName);

              const mOct = attMap.get(`${code}|${subName}|شهر 10 (أكتوبر)`) || Array(8).fill('');
              const mNov = attMap.get(`${code}|${subName}|شهر 11 (نوفمبر)`) || Array(8).fill('');
              const mSep = attMap.get(`${code}|${subName}|شهر 9 (سبتمبر)`) || [...mOct];
              const m1 = attMap.get(`${code}|${subName}|شهر 1`) || [...mOct];

              const activeSess = activeMonth.includes('11') ? [...mNov] : (activeMonth.includes('10') ? [...mOct] : [...mNov]);

              academicSubjects[subName] = {
                teacher: teacher,
                sessions: activeSess,
                months: {
                  'شهر 9 (سبتمبر)': mSep,
                  'شهر 10 (أكتوبر)': mOct,
                  'شهر 11 (نوفمبر)': mNov,
                  'شهر 1': m1
                }
              };

              const novTags = mNov.map(x => x || '_').join(', ');
              const octTags = mOct.map(x => x || '_').join(', ');
              const sessTags = activeSess.map(x => x || '_').join(', ');
              teacherParts.push(`${subName}/${teacher} [حصص: ${sessTags}] [شهر 10 (أكتوبر): ${octTags}] [شهر 11 (نوفمبر): ${novTags}]`);
            });

            result.push({
              code: st.code,
              name: st.name,
              area: st.area || '',
              phone: st.phone || '',
              parentPhone: st.parentPhone || '',
              grade: st.grade || '',
              specialization: st.specialization || '',
              subjectsSummary: subNames.join('، '),
              teachersSummary: teacherParts.join(' | '),
              regDate: st.regDate || '',
              academicSubjects: academicSubjects
            });
          });

          return result;
        }
      } catch (err) {
        console.error('[DB] Error in getAllStudents SQLite query:', err.message);
      }
    }

    // Fallback: Read from students_db.json
    try {
      if (fs.existsSync(this.jsonFile)) {
        return JSON.parse(fs.readFileSync(this.jsonFile, 'utf8'));
      }
    } catch (err) {
      console.error('[DB] Error reading fallback JSON file:', err.message);
    }
    return [];
  }
}

module.exports = { AraijDatabase };
