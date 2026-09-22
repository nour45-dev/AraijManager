const http = require('http');
const path = require('path');
const fs = require('fs');
const express = require('express');
const cors = require('cors');
const WebSocket = require('ws');

const app = express();
const server = http.createServer(app);
const wss = new WebSocket.Server({ server });

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

// WebSocket Real-time Broadcast
function broadcast(data, excludeWs = null) {
  const payload = JSON.stringify(data);
  wss.clients.forEach(client => {
    if (client !== excludeWs && client.readyState === WebSocket.OPEN) {
      client.send(payload);
    }
  });
}

wss.on('connection', (ws, req) => {
  const ip = req.socket.remoteAddress;
  console.log(`[WS] Client connected from ${ip}. Total connected: ${wss.clients.size}`);

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

  // 1. Broadcast instant real-time update to all connected clients
  broadcast({
    type: 'DATA_UPDATE',
    action,
    payload,
    data: payload,
    byUser,
    sender: byUser,
    clientId,
    timestamp: Date.now(),
    studentsCount: students.length
  }, sourceWs);

  // 2. Forward to Google Sheets Webhook asynchronously in background
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

// REST Endpoints
app.get('/api/health', (req, res) => {
  res.json({
    status: 'ok',
    app: 'Araij Manager Pro',
    version: '4.0.0',
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
  console.log(`ًںڑ€ Araij Manager Live Server running on port ${PORT}`);
  console.log(`ًںŒگ Web App: http://localhost:${PORT}`);
  console.log(`âڑ، WebSocket Server: Ready for live cross-device sync`);
  console.log(`ًں‘¥ Loaded ${students.length} students into memory`);
  console.log(`====================================================`);
});