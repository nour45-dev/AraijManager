const fs = require('fs');
const path = require('path');

let Pool = null;
try {
  Pool = require('pg').Pool;
} catch (e) {
  console.warn('[DB] pg package not available.');
}

let DatabaseSync = null;
try {
  DatabaseSync = require('node:sqlite').DatabaseSync;
} catch (e) {
  console.warn('[DB] Native node:sqlite not available in this Node runtime.');
}

function formatStudentsFromRows(studentsRows, enrollmentsRows, attendanceRows, activeMonth = 'شهر 11 (نوفمبر)') {
  const enrollmentsByCode = new Map();
  enrollmentsRows.forEach(e => {
    const sc = e.studentCode || e.student_code;
    if (!enrollmentsByCode.has(sc)) enrollmentsByCode.set(sc, []);
    enrollmentsByCode.get(sc).push(e);
  });

  const attMap = new Map();
  attendanceRows.forEach(a => {
    const sc = a.studentCode || a.student_code;
    const sNum = a.sessionNum !== undefined ? a.sessionNum : a.session_num;
    const k = `${sc}|${a.subject}|${a.month}`;
    if (!attMap.has(k)) attMap.set(k, Array(8).fill(''));
    const arr = attMap.get(k);
    const idx = Number(sNum) - 1;
    if (idx >= 0 && idx < 8) arr[idx] = a.mark;
  });

  return studentsRows.map(st => {
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
        teacher,
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

    return {
      code: st.code,
      name: st.name,
      area: st.area || '',
      phone: st.phone || '',
      parentPhone: st.parentPhone !== undefined ? st.parentPhone : (st.parent_phone || ''),
      grade: st.grade || '',
      specialization: st.specialization || '',
      subjectsSummary: subNames.join('، '),
      teachersSummary: teacherParts.join(' | '),
      regDate: st.regDate !== undefined ? st.regDate : (st.reg_date || ''),
      academicSubjects
    };
  });
}

class AraijDatabase {
  constructor(dbPath) {
    this.dbPath = dbPath || path.join(__dirname, 'araij.db');
    this.jsonFile = path.join(__dirname, 'students_db.json');
    this.engine = 'json';
    this.pool = null;
    this.sqlite = null;
    this.hasSqlite = false;
    this.hasPostgres = false;

    const dbUrl = (process.env.DATABASE_URL || '').trim();
    if (dbUrl && Pool) {
      const useSsl = !(dbUrl.includes('.railway.internal') || dbUrl.includes('localhost') || dbUrl.includes('127.0.0.1') || process.env.PGSSL === 'false');
      this.pool = new Pool({
        connectionString: dbUrl,
        ssl: useSsl ? { rejectUnauthorized: false } : false,
        max: 10,
        idleTimeoutMillis: 30000
      });
      this.pool.on('error', (err) => {
        console.error('[DB/PG] Unexpected pool error:', err.message);
      });
    }

    if (DatabaseSync) {
      try {
        const dir = path.dirname(this.dbPath);
        if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
        this.sqlite = new DatabaseSync(this.dbPath);
        this.initSqliteSchema();
        this.hasSqlite = true;
        this.engine = 'sqlite';
      } catch (err) {
        console.error('[DB/SQLite] Init failed:', err.message);
      }
    }
  }

  async init() {
    if (this.pool) {
      try {
        await this.pool.query('SELECT 1');
        await this.initPostgresSchema();
        this.hasPostgres = true;
        this.engine = 'postgres';
        console.log('[DB] Connected to PostgreSQL database successfully.');
        return 'postgres';
      } catch (err) {
        console.error('[DB/PG] Connection/schema init failed, falling back to SQLite:', err.message);
        this.hasPostgres = false;
      }
    }
    if (this.hasSqlite) {
      console.log(`[DB] Using SQLite database at: ${this.dbPath}`);
      return 'sqlite';
    }
    console.log('[DB] Using JSON file fallback.');
    return 'json';
  }

  async initPostgresSchema() {
    await this.pool.query(`
      CREATE TABLE IF NOT EXISTS students (
        code TEXT PRIMARY KEY,
        name TEXT NOT NULL,
        area TEXT DEFAULT '',
        phone TEXT DEFAULT '',
        parent_phone TEXT DEFAULT '',
        grade TEXT DEFAULT 'ث1',
        specialization TEXT DEFAULT 'عام',
        reg_date TEXT DEFAULT '',
        created_at TIMESTAMPTZ DEFAULT NOW(),
        updated_at TIMESTAMPTZ DEFAULT NOW()
      );

      CREATE TABLE IF NOT EXISTS enrollments (
        student_code TEXT NOT NULL REFERENCES students(code) ON DELETE CASCADE,
        subject TEXT NOT NULL,
        teacher TEXT NOT NULL,
        created_at TIMESTAMPTZ DEFAULT NOW(),
        PRIMARY KEY (student_code, subject)
      );

      CREATE TABLE IF NOT EXISTS attendance (
        student_code TEXT NOT NULL REFERENCES students(code) ON DELETE CASCADE,
        subject TEXT NOT NULL,
        month TEXT NOT NULL,
        session_num INTEGER NOT NULL,
        mark TEXT NOT NULL,
        updated_at TIMESTAMPTZ DEFAULT NOW(),
        PRIMARY KEY (student_code, subject, month, session_num)
      );

      CREATE TABLE IF NOT EXISTS audit_logs (
        id TEXT PRIMARY KEY,
        user_name TEXT,
        action_type TEXT,
        details TEXT,
        student_code TEXT,
        student_name TEXT,
        date_str TEXT,
        time_str TEXT,
        full_time TEXT,
        created_at TIMESTAMPTZ DEFAULT NOW()
      );

      CREATE TABLE IF NOT EXISTS deleted_codes (
        code TEXT PRIMARY KEY,
        deleted_at TIMESTAMPTZ DEFAULT NOW()
      );

      CREATE INDEX IF NOT EXISTS idx_pg_enrollments_teacher ON enrollments(teacher);
      CREATE INDEX IF NOT EXISTS idx_pg_enrollments_subject ON enrollments(subject);
      CREATE INDEX IF NOT EXISTS idx_pg_attendance_lookup ON attendance(student_code, subject, month);
    `);
  }

  initSqliteSchema() {
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

  // ---- SQLite sync helpers (used by local migration & fallback) ----
  upsertStudent(s) {
    if (!s || !s.code || !this.hasSqlite) return;
    const now = new Date().toISOString();
    this.sqlite.prepare(`
      INSERT INTO students (code, name, area, phone, parentPhone, grade, specialization, regDate, createdAt, updatedAt)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      ON CONFLICT(code) DO UPDATE SET
        name = excluded.name, area = excluded.area, phone = excluded.phone,
        parentPhone = excluded.parentPhone, grade = excluded.grade,
        specialization = excluded.specialization, regDate = excluded.regDate, updatedAt = excluded.updatedAt
    `).run(
      String(s.code).trim(), (s.name || '').trim(), (s.area || '').trim(),
      (s.phone || '').trim(), (s.parentPhone || '').trim(), (s.grade || 'ث1').trim(),
      (s.specialization || 'عام').trim(), (s.regDate || '').trim(), s.createdAt || now, now
    );
  }

  setEnrollment(studentCode, subject, teacher) {
    if (!studentCode || !subject || !teacher || !this.hasSqlite) return;
    const now = new Date().toISOString();
    this.sqlite.prepare(`
      INSERT INTO enrollments (studentCode, subject, teacher, createdAt)
      VALUES (?, ?, ?, ?)
      ON CONFLICT(studentCode, subject) DO UPDATE SET teacher = excluded.teacher
    `).run(String(studentCode).trim(), subject.trim(), teacher.trim(), now);
  }

  recordAttendance(studentCode, subject, month, sessionNum, mark) {
    if (!studentCode || !subject || !month || !sessionNum || !this.hasSqlite) return;
    const now = new Date().toISOString();
    if (!mark || !mark.trim() || mark.trim() === '_' || mark.trim() === '-') {
      this.sqlite.prepare(`DELETE FROM attendance WHERE studentCode = ? AND subject = ? AND month = ? AND sessionNum = ?`)
        .run(String(studentCode).trim(), subject.trim(), month.trim(), Number(sessionNum));
      return;
    }
    this.sqlite.prepare(`
      INSERT INTO attendance (studentCode, subject, month, sessionNum, mark, updatedAt)
      VALUES (?, ?, ?, ?, ?, ?)
      ON CONFLICT(studentCode, subject, month, sessionNum) DO UPDATE SET mark = excluded.mark, updatedAt = excluded.updatedAt
    `).run(String(studentCode).trim(), subject.trim(), month.trim(), Number(sessionNum), mark.trim(), now);
  }

  syncFullStudent(st) {
    if (!st || !st.code) return;
    if (this.hasSqlite) {
      this.upsertStudent(st);
      if (st.academicSubjects && typeof st.academicSubjects === 'object') {
        for (const [subName, subData] of Object.entries(st.academicSubjects)) {
          if (!subData || !subData.teacher) continue;
          this.setEnrollment(st.code, subName, subData.teacher);
          if (subData.months && typeof subData.months === 'object') {
            for (const [mName, mSessions] of Object.entries(subData.months)) {
              if (Array.isArray(mSessions)) {
                mSessions.forEach((val, idx) => {
                  this.recordAttendance(st.code, subName, mName, idx + 1, val ? String(val).trim() : '');
                });
              }
            }
          }
        }
      }
    }
  }

  deleteStudent(code) {
    if (!code) return;
    const c = String(code).trim();
    if (this.hasSqlite) {
      try {
        this.sqlite.prepare(`DELETE FROM students WHERE code = ?`).run(c);
        this.sqlite.prepare(`DELETE FROM enrollments WHERE studentCode = ?`).run(c);
        this.sqlite.prepare(`DELETE FROM attendance WHERE studentCode = ?`).run(c);
      } catch (err) { console.error('[DB/SQLite] deleteStudent error:', err.message); }
    }
    if (this.hasPostgres) {
      this.pgDeleteStudent(c).catch(e => console.error('[DB/PG] deleteStudent error:', e.message));
    }
  }

  getAllStudents(activeMonth = 'شهر 11 (نوفمبر)') {
    if (this.hasSqlite) {
      try {
        const studentsRows = this.sqlite.prepare(`SELECT * FROM students ORDER BY code ASC`).all();
        if (studentsRows && studentsRows.length > 0) {
          const enrollmentsRows = this.sqlite.prepare(`SELECT * FROM enrollments`).all();
          const attendanceRows = this.sqlite.prepare(`SELECT * FROM attendance`).all();
          return formatStudentsFromRows(studentsRows, enrollmentsRows, attendanceRows, activeMonth);
        }
      } catch (err) { console.error('[DB/SQLite] getAllStudents error:', err.message); }
    }
    try {
      if (fs.existsSync(this.jsonFile)) return JSON.parse(fs.readFileSync(this.jsonFile, 'utf8'));
    } catch (err) { console.error('[DB] JSON fallback error:', err.message); }
    return [];
  }

  // ---- PostgreSQL async methods ----
  async pgGetAllStudents(activeMonth = 'شهر 11 (نوفمبر)') {
    if (!this.hasPostgres) return null;
    const [sRes, eRes, aRes] = await Promise.all([
      this.pool.query('SELECT * FROM students ORDER BY code ASC'),
      this.pool.query('SELECT * FROM enrollments'),
      this.pool.query('SELECT * FROM attendance')
    ]);
    return formatStudentsFromRows(sRes.rows, eRes.rows, aRes.rows, activeMonth);
  }

  async pgGetDeletedCodes() {
    if (!this.hasPostgres) return [];
    const res = await this.pool.query('SELECT code FROM deleted_codes');
    return res.rows.map(r => r.code);
  }

  async pgGetAuditLogs(limit = 500) {
    if (!this.hasPostgres) return [];
    const res = await this.pool.query('SELECT * FROM audit_logs ORDER BY created_at DESC LIMIT $1', [limit]);
    return res.rows.map(r => ({
      id: r.id,
      userName: r.user_name,
      actionType: r.action_type,
      details: r.details,
      studentCode: r.student_code,
      studentName: r.student_name,
      dateStr: r.date_str,
      timeStr: r.time_str,
      fullTime: r.full_time
    }));
  }

  async pgAddAuditLog(log) {
    if (!this.hasPostgres || !log) return;
    await this.pool.query(
      `INSERT INTO audit_logs (id, user_name, action_type, details, student_code, student_name, date_str, time_str, full_time)
       VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9)
       ON CONFLICT (id) DO NOTHING`,
      [String(log.id || ('log_' + Date.now())), log.userName || 'مشرف', log.actionType || '', log.details || '', log.studentCode || '', log.studentName || '', log.dateStr || '', log.timeStr || '', log.fullTime || '']
    );
  }

  async pgDeleteStudent(code) {
    if (!this.hasPostgres || !code) return;
    const c = String(code).trim();
    await this.pool.query('INSERT INTO deleted_codes (code) VALUES ($1) ON CONFLICT (code) DO NOTHING', [c]);
    await this.pool.query('DELETE FROM students WHERE code = $1', [c]);
  }

  async pgUpsertStudentInClient(client, st) {
    const code = String(st.code || '').trim();
    if (!code) return;
    await client.query(
      `INSERT INTO students (code, name, area, phone, parent_phone, grade, specialization, reg_date, updated_at)
       VALUES ($1, $2, $3, $4, $5, $6, $7, $8, NOW())
       ON CONFLICT (code) DO UPDATE SET
         name = EXCLUDED.name, area = EXCLUDED.area, phone = EXCLUDED.phone,
         parent_phone = EXCLUDED.parent_phone, grade = EXCLUDED.grade,
         specialization = EXCLUDED.specialization, reg_date = EXCLUDED.reg_date, updated_at = NOW()`,
      [code, (st.name || '').trim(), (st.area || '').trim(), (st.phone || '').trim(), (st.parentPhone || '').trim(), (st.grade || 'ث1').trim(), (st.specialization || 'عام').trim(), (st.regDate || '').trim()]
    );

    const subjEntries = Object.entries(st.academicSubjects || {}).filter(([, v]) => v && v.teacher);
    const keepSubs = subjEntries.map(([k]) => k);
    if (keepSubs.length > 0) {
      await client.query('DELETE FROM enrollments WHERE student_code = $1 AND NOT (subject = ANY($2::text[]))', [code, keepSubs]);
      await client.query('DELETE FROM attendance WHERE student_code = $1 AND NOT (subject = ANY($2::text[]))', [code, keepSubs]);
    } else {
      await client.query('DELETE FROM enrollments WHERE student_code = $1', [code]);
      await client.query('DELETE FROM attendance WHERE student_code = $1', [code]);
    }

    for (const [subName, subData] of subjEntries) {
      await client.query(
        `INSERT INTO enrollments (student_code, subject, teacher)
         VALUES ($1, $2, $3)
         ON CONFLICT (student_code, subject) DO UPDATE SET teacher = EXCLUDED.teacher`,
        [code, subName, String(subData.teacher).trim()]
      );

      if (subData.months && typeof subData.months === 'object') {
        for (const [mName, mSessions] of Object.entries(subData.months)) {
          if (!Array.isArray(mSessions)) continue;
          for (let i = 0; i < Math.min(8, mSessions.length); i++) {
            const val = mSessions[i] ? String(mSessions[i]).trim() : '';
            if (!val || val === '_' || val === '-') {
              await client.query(
                'DELETE FROM attendance WHERE student_code = $1 AND subject = $2 AND month = $3 AND session_num = $4',
                [code, subName, mName, i + 1]
              );
            } else {
              await client.query(
                `INSERT INTO attendance (student_code, subject, month, session_num, mark, updated_at)
                 VALUES ($1, $2, $3, $4, $5, NOW())
                 ON CONFLICT (student_code, subject, month, session_num) DO UPDATE SET mark = EXCLUDED.mark, updated_at = NOW()`,
                [code, subName, mName, i + 1, val]
              );
            }
          }
        }
      }
    }
  }

  async pgSyncStudents(studentList) {
    if (!this.hasPostgres || !Array.isArray(studentList) || studentList.length === 0) return;
    const client = await this.pool.connect();
    try {
      await client.query('BEGIN');
      for (const st of studentList) {
        await this.pgUpsertStudentInClient(client, st);
      }
      await client.query('COMMIT');
    } catch (err) {
      await client.query('ROLLBACK');
      console.error('[DB/PG] pgSyncStudents transaction error:', err.message);
      throw err;
    } finally {
      client.release();
    }
  }

  async pgSeedInitial(studentList, logsList = [], deletedList = []) {
    if (!this.hasPostgres) return;
    console.log(`[DB/PG] Seeding PostgreSQL with ${studentList.length} students...`);
    await this.pgSyncStudents(studentList);
    for (const c of deletedList) {
      await this.pool.query('INSERT INTO deleted_codes (code) VALUES ($1) ON CONFLICT (code) DO NOTHING', [String(c).trim()]);
    }
    for (const log of logsList.slice(0, 500)) {
      await this.pgAddAuditLog(log);
    }
    console.log('[DB/PG] Initial PostgreSQL seed complete!');
  }
}

module.exports = { AraijDatabase };
