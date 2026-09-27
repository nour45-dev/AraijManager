const http = require('http');
const path = require('path');
const fs = require('fs');
const express = require('express');
const cors = require('cors');
const WebSocket = require('ws');

const app = express();
const server = http.createServer(app);

// ✅ FIX 1: تحديد مسار /ws بشكل صريح لضمان عمل الـ reverse proxy على Railway
const wss = new WebSocket.Server({ server, path: '/ws' });

const PORT = process.env.PORT || 3000;
const GOOGLE_APPS_SCRIPT_URL = process.env.SHEETS_URL || 'https://script.google.com/macros/s/AKfycbymmlVQKxkdXpLlsE0Z7DuhizyHHGlEKHesCAj0FJ4lfKlSBJRLThN0GFfISbi5CabmAg/exec';

const DB_FILE = path.join(__dirname, 'students_db.json');
const AUDIT_FILE = path.join(__dirname, 'audit_log.json');

// Middleware
app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));

// In-Memory State
let students = [];
let auditLogs = [];

// 1. Initialize Students Database
function initDatabase() {
  try {
    if (fs.existsSync(DB_FILE)) {
      const raw = fs.readFileSync(DB_FILE, 'utf8');
      const loaded = JSON.parse(raw);
      students = (Array.isArray(loaded) ? loaded : []).filter(s => s && s.name && s.name.trim() && s.name.trim().length >= 2 && !s.name.includes('غير مسمى'));
      console.log(`[DB] Loaded ${students.length} valid students from students_db.json`);
    } else {
      // Seed from data.js if exists
      const dataJsPath = path.join(__dirname, 'data.js');
      if (fs.existsSync(dataJsPath)) {
        const content = fs.readFileSync(dataJsPath, 'utf8');
        const match = content.match(/window\.STUDENTS_DATA\s*=\s*(\[[\s\S]*?\]);/);
        if (match) {
          const rawParsed = JSON.parse(match[1]);
          students = (Array.isArray(rawParsed) ? rawParsed : []).filter(s => s && s.name && s.name.trim() && s.name.trim().length >= 2 && !s.name.includes('غير مسمى'));
          saveDatabase();
          console.log(`[DB] Seeded ${students.length} valid students from data.js`);
        }
      }
    }
  } catch (err) {
    console.error('[DB] Error initializing database:', err);
    students = [];
  }

  try {
    if (fs.existsSync(AUDIT_FILE)) {
      auditLogs = JSON.parse(fs.readFileSync(AUDIT_FILE, 'utf8'));
    }
  } catch (e) {
    auditLogs = [];
  }
}

function saveDatabase() {
  try {
    fs.writeFileSync(DB_FILE, JSON.stringify(students, null, 2), 'utf8');
  } catch (err) {
    console.error('[DB] Error saving students_db.json:', err);
  }
}

function saveAuditLogs() {
  try {
    if (auditLogs.length > 1000) auditLogs = auditLogs.slice(0, 1000);
    fs.writeFileSync(AUDIT_FILE, JSON.stringify(auditLogs, null, 2), 'utf8');
  } catch (err) {
    console.error('[DB] Error saving audit_log.json:', err);
  }
}

initDatabase();

// ✅ FIX 2: WebSocket Real-time Broadcast مع sound flag
function broadcast(data, excludeWs = null) {
  const payload = JSON.stringify(data);
  wss.clients.forEach(client => {
    if (client !== excludeWs && client.readyState === WebSocket.OPEN) {
      try {
        client.send(payload);
      } catch (e) {
        console.warn('[WS] Failed to send to client:', e.message);
      }
    }
  });
}

// ✅ FIX 3: Server-side Heartbeat — يمنع Railway من قفل الـ connections بعد 30 ثانية
const HEARTBEAT_INTERVAL = 20000; // كل 20 ثانية

function heartbeat() {
  this.isAlive = true;
}

const heartbeatTimer = setInterval(() => {
  wss.clients.forEach(ws => {
    if (ws.isAlive === false) {
      console.log('[WS] Terminating dead connection');
      return ws.terminate();
    }
    ws.isAlive = false;
    try {
      ws.ping();
    } catch (e) {}
  });
}, HEARTBEAT_INTERVAL);

wss.on('close', () => {
  clearInterval(heartbeatTimer);
});

wss.on('connection', (ws, req) => {
  const ip = req.socket.remoteAddress;
  console.log(`[WS] Client connected from ${ip}. Total connected: ${wss.clients.size}`);

  // Mark as alive for heartbeat
  ws.isAlive = true;
  ws.on('pong', heartbeat.bind(ws));

  // Send initial welcome & count
  ws.send(JSON.stringify({
    type: 'CONNECTED',
    message: 'Connected to Araij Manager Live Sync Server',
    studentsCount: students.length,
    timestamp: Date.now()
  }));

  ws.on('message', message => {
    try {
      const msg = JSON.parse(message);
      if (msg.type === 'PING') {
        ws.send(JSON.stringify({ type: 'PONG', timestamp: Date.now() }));
        ws.isAlive = true; // Reset heartbeat on client ping too
      } else if (msg.type === 'SYNC_ACTION' || msg.type === 'DATA_UPDATE') {
        const action = msg.action;
        const payload = msg.payload || msg.data || msg;
        const byUser = msg.byUser || msg.sender || 'مشرف';
        const clientId = msg.clientId || '';
        handleSyncAction(action, payload, byUser, ws, clientId);
      }
    } catch (e) {
      console.warn('[WS] Error processing message:', e);
    }
  });

  ws.on('close', () => {
    console.log(`[WS] Client disconnected. Remaining: ${wss.clients.size}`);
  });

  ws.on('error', (err) => {
    console.warn('[WS] Socket error:', err.message);
  });
});

// Central Handler for Sync Actions (both REST & WS)
function handleSyncAction(action, payload, byUser = 'مشرف', sourceWs = null, clientId = '') {
  let updatedCount = 0;
  const now = new Date();
  const dateStr = now.toLocaleDateString('ar-EG');
  const timeStr = now.toLocaleTimeString('ar-EG', { hour: '2-digit', minute: '2-digit', second: '2-digit' });

  if (action === 'add_student') {
    const newSt = payload.student || payload;
    const code = String(newSt.code || '').trim();
    const existingIdx = students.findIndex(s => String(s.code).trim() === code);
    if (existingIdx >= 0) {
      students[existingIdx] = { ...students[existingIdx], ...newSt };
    } else {
      students.unshift(newSt);
    }
    updatedCount = 1;

    auditLogs.unshift({
      id: 'log_' + Date.now(),
      userName: byUser,
      actionType: 'تسجيل طالب جديد',
      details: `تسجيل الطالب "${newSt.name}" بالكود #${code}`,
      studentCode: code,
      studentName: newSt.name,
      dateStr,
      timeStr,
      fullTime: `${dateStr} - ${timeStr}`
    });

  } else if (action === 'update_student') {
    const updated = payload.student || payload;
    const code = String(updated.code || '').trim();
    const idx = students.findIndex(s => String(s.code).trim() === code);
    if (idx >= 0) {
      students[idx] = { ...students[idx], ...updated };
    } else {
      students.unshift(updated);
    }
    updatedCount = 1;

    auditLogs.unshift({
      id: 'log_' + Date.now(),
      userName: byUser,
      actionType: 'تعديل بيانات طالب',
      details: `تعديل بيانات/حصص الطالب "${updated.name}" (#${code})`,
      studentCode: code,
      studentName: updated.name,
      dateStr,
      timeStr,
      fullTime: `${dateStr} - ${timeStr}`
    });

  } else if (action === 'delete_student') {
    const code = String(payload.code || '').trim();
    const target = students.find(s => String(s.code).trim() === code);
    const targetName = target ? target.name : '';
    students = students.filter(s => String(s.code).trim() !== code);
    updatedCount = 1;

    auditLogs.unshift({
      id: 'log_' + Date.now(),
      userName: byUser,
      actionType: 'حذف طالب',
      details: `حذف الطالب "${targetName}" (#${code}) نهائياً`,
      studentCode: code,
      studentName: targetName,
      dateStr,
      timeStr,
      fullTime: `${dateStr} - ${timeStr}`
    });

  } else if (action === 'batch_update_students' || action === 'bulk_session') {
    const batchList = payload.students || [];
    batchList.forEach(st => {
      const code = String(st.code || '').trim();
      const idx = students.findIndex(s => String(s.code).trim() === code);
      if (idx >= 0) {
        students[idx] = { ...students[idx], ...st };
      } else {
        students.unshift(st);
      }
    });
    updatedCount = batchList.length;

    auditLogs.unshift({
      id: 'log_' + Date.now(),
      userName: byUser,
      actionType: 'رصد جماعي للحضور',
      details: `رصد وتحديث حضور ${batchList.length} طالب دفعة واحدة`,
      studentCode: '',
      studentName: '',
      dateStr,
      timeStr,
      fullTime: `${dateStr} - ${timeStr}`
    });

  } else if (action === 'log_activity') {
    const log = payload.log || payload;
    auditLogs.unshift(log);
  }

  saveDatabase();
  saveAuditLogs();

  // ✅ FIX 4: Broadcast مع sound: true لتشغيل الصوت عند كل المشرفين
  broadcast({
    type: 'DATA_UPDATE',
    action,
    payload,
    data: payload,
    byUser,
    sender: byUser,
    clientId,
    timestamp: Date.now(),
    studentsCount: students.length,
    sound: true   // 🔔 علامة تشغيل الصوت عند المستلمين
  }, sourceWs);

  // Forward to Google Sheets Webhook asynchronously in background
  if (GOOGLE_APPS_SCRIPT_URL) {
    forwardToGoogleSheets(action, payload, byUser);
  }

  return { status: 'success', action, updatedCount, totalStudents: students.length };
}

// Background forwarding to Google Sheets
function forwardToGoogleSheets(action, payload, byUser) {
  try {
    const https = require('https');
    const httpLib = GOOGLE_APPS_SCRIPT_URL.startsWith('https') ? https : http;
    const bodyStr = JSON.stringify({ action, ...payload, byUser });
    const urlObj = new URL(GOOGLE_APPS_SCRIPT_URL);

    const req = httpLib.request(urlObj, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(bodyStr)
      }
    });

    req.on('error', (err) => {
      console.warn('[Google Sheets Relay] Relay error (non-fatal):', err.message);
    });

    req.write(bodyStr);
    req.end();
  } catch (e) {
    // Non-fatal background forward
  }
}

// Canonical helpers for server-side sheet parsing
const GOOGLE_SHEET_CSV_URL = 'https://docs.google.com/spreadsheets/d/13J9pKu0dnL4rcyhYDTh3lQY2upe2yCc7jaV16AhTCpI/export?format=csv&gid=0';

function canonicalSubject(name) {
  if (!name) return '';
  const n = String(name).trim();
  if (n === 'فيزياء' || n === 'فزياء') return 'فزياء';
  if (n === 'أحياء' || n === 'احياء') return 'احياء';
  if (n === 'رياضيات' || n === 'رياضة') return 'رياضة';
  if (n === 'إحصاء' || n === 'احصاء') return 'احصاء';
  if (n === 'لغة عربية' || n === 'عربى' || n === 'عربي') return 'عربي';
  if (n === 'لغة انجليزية' || n === 'إنجليزي' || n === 'انجليزي') return 'انجليزي';
  if (n === 'لغة فرنسية' || n === 'فرنساوى' || n === 'فرنساوي') return 'فرنساوي';
  return n;
}

function cleanTeacher(raw) {
  if (!raw) return 'مدرس المادة';
  return String(raw).trim().replace(/^(أستاذ|الاستاذ|الأستاذ|استاذ)\s+/i, 'ا/ ');
}

function normalizeGrade(g) {
  if (!g) return 'ث1';
  const str = String(g).trim();
  if (str.includes('1') || str.includes('١') || str.includes('أولى')) return 'ث1';
  if (str.includes('2') || str.includes('٢') || str.includes('ثانية')) return 'ث2';
  if (str.includes('3') || str.includes('٣') || str.includes('ثالثة')) return 'ث3';
  return 'ث1';
}

function parseAcademicSubjectsFromSummary(subjectsSummary, teachersSummary) {
  const result = {};
  if (teachersSummary && typeof teachersSummary === 'string') {
    const parts = teachersSummary.split(/[|\n;]/).map(p => p.trim()).filter(Boolean);
    parts.forEach(p => {
      const slash = p.indexOf('/');
      if (slash > -1) {
        const rawSub = p.substring(0, slash).trim();
        const canon = canonicalSubject(rawSub);
        if (!canon) return;
        let tPart = p.substring(slash + 1).trim();
        let sessions = ['', '', '', '', '', '', '', ''];
        const sessMatch = tPart.match(/\[(شهر\s*[^:]+|حصص):\s*([^\]]*)\]/);
        if (sessMatch) {
          const rawSess = sessMatch[2].split(',').map(s => s.trim());
          for (let i = 0; i < 8 && i < rawSess.length; i++) {
            sessions[i] = (rawSess[i] === '_' || rawSess[i] === '-') ? '' : rawSess[i];
          }
          tPart = tPart.replace(/\[(شهر\s*[^:]+|حصص):\s*([^\]]*)\]/, '').trim();
        }
        result[canon] = {
          teacher: cleanTeacher(tPart),
          sessions: sessions,
          months: {
            'شهر 9 (سبتمبر)': [...sessions],
            'شهر 1': [...sessions]
          }
        };
      }
    });
  }

  if (subjectsSummary && typeof subjectsSummary === 'string') {
    const subs = subjectsSummary.split(/[,،|\n;]/).map(s => s.trim()).filter(Boolean);
    subs.forEach(s => {
      const canon = canonicalSubject(s);
      if (canon && !result[canon]) {
        result[canon] = {
          teacher: 'مدرس المادة',
          sessions: ['', '', '', '', '', '', '', ''],
          months: {
            'شهر 9 (سبتمبر)': ['', '', '', '', '', '', '', ''],
            'شهر 1': ['', '', '', '', '', '', '', '']
          }
        };
      }
    });
  }
  return result;
}

function computeStudentMetrics(academicSubjects) {
  let totalScore = 0;
  let totalMax = 0;
  let totalPresent = 0;
  let totalAbsent = 0;
  let totalRecorded = 0;

  const subjects = academicSubjects || {};
  Object.keys(subjects).forEach(subKey => {
    const sessions = subjects[subKey].sessions || [];
    sessions.forEach(val => {
      if (!val || !val.trim()) return;
      const cleanVal = val.trim();
      totalRecorded++;
      if (cleanVal === 'غ' || cleanVal === 'غايب' || cleanVal === 'غياب') {
        totalAbsent++;
      } else if (cleanVal === '✓' || cleanVal === 'حاضر' || cleanVal === 'حضر') {
        totalPresent++;
      } else if (cleanVal.includes('/')) {
        totalPresent++;
        const parts = cleanVal.split('/');
        const earned = parseFloat(parts[0]);
        const max = parseFloat(parts[1]);
        if (!isNaN(earned) && !isNaN(max) && max > 0) {
          totalScore += earned;
          totalMax += max;
        }
      } else {
        const num = parseFloat(cleanVal);
        if (!isNaN(num)) {
          totalPresent++;
          totalScore += num;
          totalMax += 20;
        } else {
          totalPresent++;
        }
      }
    });
  });

  const attendanceRate = totalRecorded > 0 ? Math.round((totalPresent / totalRecorded) * 100) : 100;
  const averageScore = totalMax > 0 ? Math.round((totalScore / totalMax) * 100 * 10) / 10 : 0;

  return { totalScore, totalMax, totalPresent, totalAbsent, totalRecordedSessions: totalRecorded, attendanceRate, averageScore };
}

function rebuildStudentSummary(st) {
  if (!st || !st.academicSubjects) return;
  const subs = Object.keys(st.academicSubjects);
  st.subjectsSummary = subs.join(', ');
  st.teachersSummary = subs.map(k => {
    const subObj = st.academicSubjects[k] || {};
    const tName = subObj.teacher || 'مدرس المادة';
    let p = `${k}/${tName}`;
    if (subObj.sessions && Array.isArray(subObj.sessions) && subObj.sessions.some(s => s && String(s).trim())) {
      const sessStr = subObj.sessions.map(s => (s && String(s).trim()) ? String(s).trim() : '_').join(', ');
      p += ` [حصص: ${sessStr}]`;
    }
    return p;
  }).join(' | ');
}

// Automatic Live Pull from Google Sheets
let _isSyncingSheets = false;
async function syncFromGoogleSheets() {
  if (_isSyncingSheets) return;
  _isSyncingSheets = true;
  try {
    const https = require('https');
    let fetchedStudents = [];

    // Helper to fetch via HTTPS with redirects
    const fetchUrl = (targetUrl) => new Promise((resolve, reject) => {
      const get = (url, depth = 0) => {
        if (depth > 5) return reject(new Error('Too many redirects'));
        https.get(url, (res) => {
          if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
            return get(res.headers.location, depth + 1);
          }
          let data = '';
          res.on('data', chunk => data += chunk);
          res.on('end', () => resolve(data));
        }).on('error', reject);
      };
      get(targetUrl);
    });

    // 1. Try Google Apps Script API first
    try {
      const jsonRaw = await fetchUrl(`${GOOGLE_APPS_SCRIPT_URL}?action=get_all&_t=${Date.now()}`);
      const parsed = JSON.parse(jsonRaw);
      if (parsed && parsed.status === 'success' && Array.isArray(parsed.students) && parsed.students.length > 500) {
        fetchedStudents = parsed.students;
        console.log(`[Google Sheets Sync] Successfully fetched ${fetchedStudents.length} students from Apps Script API`);
      }
    } catch (apiErr) {
      console.warn('[Google Sheets Sync] Apps Script API failed, trying CSV:', apiErr.message);
    }

    // 2. Fallback to direct Google Sheets CSV export
    if (fetchedStudents.length === 0) {
      try {
        const csvRaw = await fetchUrl(GOOGLE_SHEET_CSV_URL);
        const lines = csvRaw.split('\n');
        for (let i = 1; i < lines.length; i++) {
          const line = lines[i].trim();
          if (!line) continue;
          const tokens = line.split(/,(?=(?:[^"]*"[^"]*")*[^"]*$)/);
          if (tokens.length < 2) continue;
          const code = tokens[0].trim().replace(/^"|"$/g, '');
          const name = tokens[1].trim().replace(/^"|"$/g, '');
          if (!name || name.length < 2 || name.includes('غير مسمى')) continue;
          const area = tokens.length > 2 ? tokens[2].trim().replace(/^"|"$/g, '') : '';
          const phone = tokens.length > 3 ? tokens[3].trim().replace(/^"|"$/g, '') : '';
          const parentPhone = tokens.length > 4 ? tokens[4].trim().replace(/^"|"$/g, '') : '';
          const grade = tokens.length > 5 ? tokens[5].trim().replace(/^"|"$/g, '') : 'ث1';
          const specialization = tokens.length > 6 ? tokens[6].trim().replace(/^"|"$/g, '') : 'عام';
          const subjectsSummary = tokens.length > 7 ? tokens[7].trim().replace(/^"|"$/g, '') : '';
          const teachersSummary = tokens.length > 8 ? tokens[8].trim().replace(/^"|"$/g, '') : '';
          const regDate = tokens.length > 9 ? tokens[9].trim().replace(/^"|"$/g, '') : '';
          fetchedStudents.push({ code, name, area, phone, parentPhone, grade, specialization, subjectsSummary, teachersSummary, regDate });
        }
        console.log(`[Google Sheets Sync] Successfully fetched ${fetchedStudents.length} students from CSV`);
      } catch (csvErr) {
        console.warn('[Google Sheets Sync] CSV export also failed:', csvErr.message);
      }
    }

    if (fetchedStudents.length > 0) {
      let updatedCount = 0;
      let newCount = 0;

      fetchedStudents.forEach(st => {
        const code = String(st.code || '').trim();
        if (!code || !st.name) return;
        const normG = normalizeGrade(st.grade);
        const academic = parseAcademicSubjectsFromSummary(st.subjectsSummary, st.teachersSummary);
        const metrics = computeStudentMetrics(academic);

        const existingIdx = students.findIndex(s => String(s.code).trim() === code);
        if (existingIdx >= 0) {
          const existing = students[existingIdx];

          existing.name = st.name;
          existing.grade = normG;
          if (st.phone) existing.phone = st.phone;
          if (st.parentPhone) existing.parentPhone = st.parentPhone;
          if (st.area) existing.area = st.area;
          if (st.specialization) existing.specialization = st.specialization;
          if (st.regDate) existing.regDate = st.regDate;

          // ✅ FIX: دمج غير إتلافي — لا نمسح الحصص المسجلة حياً في السيستم بحصص فارغة من الشيت
          if (academic && Object.keys(academic).length > 0) {
            if (!existing.academicSubjects) existing.academicSubjects = {};
            Object.keys(academic).forEach(subKey => {
              if (!existing.academicSubjects[subKey]) {
                existing.academicSubjects[subKey] = academic[subKey];
              } else {
                const exSub = existing.academicSubjects[subKey];
                const shSub = academic[subKey];
                if (shSub.teacher && (!exSub.teacher || exSub.teacher === 'مدرس المادة')) {
                  exSub.teacher = shSub.teacher;
                }
                if (shSub.sessions && Array.isArray(shSub.sessions)) {
                  if (!exSub.sessions || !Array.isArray(exSub.sessions)) {
                    exSub.sessions = ['', '', '', '', '', '', '', ''];
                  }
                  while (exSub.sessions.length < 8) exSub.sessions.push('');
                  shSub.sessions.forEach((val, sIdx) => {
                    // فقط إذا كانت الحصة في السيرفر فارغة والشيت فيه قيمة نأخذها
                    if (val && val.trim() && (!exSub.sessions[sIdx] || !exSub.sessions[sIdx].trim())) {
                      exSub.sessions[sIdx] = val.trim();
                    }
                  });
                }
                if (!exSub.months) exSub.months = {};
                exSub.months['شهر 9 (سبتمبر)'] = [...exSub.sessions];
                exSub.months['شهر 1'] = [...exSub.sessions];
              }
            });
            existing._metrics = computeStudentMetrics(existing.academicSubjects);
            rebuildStudentSummary(existing);
          } else if (!existing.academicSubjects) {
            existing.academicSubjects = {};
            existing._metrics = computeStudentMetrics(existing.academicSubjects);
          }
          updatedCount++;
        } else {
          const newSt = {
            code,
            name: st.name,
            area: st.area || '',
            phone: st.phone || '',
            parentPhone: st.parentPhone || '',
            grade: normG,
            specialization: st.specialization || 'عام',
            subjectsSummary: st.subjectsSummary || '',
            teachersSummary: st.teachersSummary || '',
            regDate: st.regDate || new Date().toLocaleDateString('ar-EG'),
            academicSubjects: academic,
            _metrics: metrics,
            _searchString: `${code} ${st.name} ${st.phone || ''} ${st.parentPhone || ''} ${st.area || ''} ${normG} ${st.specialization || ''}`.toLowerCase()
          };
          students.push(newSt);
          newCount++;
        }
      });

      saveDatabase();
      console.log(`[Google Sheets Sync] Complete: ${students.length} total (${newCount} new, ${updatedCount} updated).`);

      broadcast({
        type: 'DATA_UPDATE',
        action: 'google_sheets_live_sync',
        timestamp: Date.now(),
        studentsCount: students.length,
        sound: false  // لا نشغل صوت عند مزامنة الـ Google Sheets التلقائية
      });
    }
  } catch (err) {
    console.error('[Google Sheets Sync] Error:', err.message);
  } finally {
    _isSyncingSheets = false;
  }
}

// Background sync on startup and every 60 seconds
setTimeout(syncFromGoogleSheets, 5000);
setInterval(syncFromGoogleSheets, 60000);


// REST Endpoints
app.get('/api/health', (req, res) => {
  res.json({
    status: 'ok',
    app: 'Araij Manager Pro',
    version: '4.1.0',
    clientsCount: wss.clients.size,
    studentsCount: students.length,
    time: new Date().toISOString()
  });
});

app.get('/api/students', (req, res) => {
  res.setHeader('Cache-Control', 'no-store');
  res.json({
    status: 'success',
    count: students.length,
    students: students
  });
});

app.get('/api/force_sync_sheets', async (req, res) => {
  try {
    await syncFromGoogleSheets();
    res.json({
      status: 'success',
      message: 'تمت مزامنة شيت جوجل بنجاح',
      studentsCount: students.length
    });
  } catch (e) {
    res.status(500).json({ status: 'error', message: e.message });
  }
});

app.get('/api/audit', (req, res) => {
  res.setHeader('Cache-Control', 'no-store');
  res.json({
    status: 'success',
    count: auditLogs.length,
    logs: auditLogs.slice(0, 500)
  });
});

app.post('/api/sync', (req, res) => {
  const action = req.body.action;
  const payload = req.body.payload || req.body.data || req.body;
  const byUser = req.body.byUser || req.body.sender || 'مشرف';
  const clientId = req.body.clientId || '';
  if (!action) {
    return res.status(400).json({ status: 'error', message: 'Action required' });
  }
  const result = handleSyncAction(action, payload, byUser, null, clientId);
  res.json(result);
});

// Backward compatibility for old local server sync
app.post('/api/sync_local', (req, res) => {
  const { action, student, code, byUser, clientId } = req.body;
  const payload = { student: student || req.body, code: code || (student && student.code) };
  const result = handleSyncAction(action || 'update_student', payload, byUser || 'مشرف', null, clientId || '');
  res.json(result);
});

app.get('/api/get_students', (req, res) => {
  res.setHeader('Cache-Control', 'no-store');
  res.json(students);
});

// Serve static frontend assets
app.use(express.static(__dirname));

// Single Page Application Fallback
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

// Start Server
server.listen(PORT, '0.0.0.0', () => {
  console.log(`====================================================`);
  console.log(`🚀 Araij Manager Live Server running on port ${PORT}`);
  console.log(`🌐 Web App: http://localhost:${PORT}`);
  console.log(`⚡ WebSocket Server: Ready at ws://localhost:${PORT}/ws`);
  console.log(`💥 Loaded ${students.length} students into memory`);
  console.log(`====================================================`);
});