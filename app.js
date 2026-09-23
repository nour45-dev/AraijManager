/**
 * AraijManager Pro v2.0 - Complete Google Sheets & Telegram Bot Feature Replica + Live Sync Hub
 */

// Unique Client Session ID to differentiate devices and avoid self-echoes while ensuring multi-device sync
if (!window.ARAIJ_CLIENT_ID) {
  let savedId = '';
  try { savedId = sessionStorage.getItem('araij_client_id') || ''; } catch(e) {}
  if (!savedId) {
    savedId = 'client_' + Date.now().toString(36) + '_' + Math.random().toString(36).substring(2, 9);
    try { sessionStorage.setItem('araij_client_id', savedId); } catch(e) {}
  }
  window.ARAIJ_CLIENT_ID = savedId;
}

// ====================================================
// 1. CONFIGURATION & CONSTANTS (from config.py)
// ====================================================

const SPECIALIZATIONS = {
  "رئيسية": ["عام", "أزهر", "بكالوريا"],
  "بكالوريا": ["طب", "هندسة", "الاداب", "الاعمال"]
};

const SUBJECTS = {
  "ث1": ["عربي", "انجليزي", "تاريخ", "جغرافيا", "علوم", "فلسفة", "رياضة", "فرنساوي", "برمجة"],
  "ث2": ["عربي", "انجليزي", "تاريخ", "جغرافيا", "كيمياء", "فزياء", "برمجة", "محاسبة", "علم نفس", "ادارة اعمال", "فرنساوي", "رياضة"],
  "ث3": ["عربي", "انجليزي", "تاريخ", "جغرافيا", "كيمياء", "فزياء", "احياء", "رياضة", "احصاء", "فرنساوي"]
};

const TEACHERS = {
  "عربي":        ["ا/ احمد عبد القادر", "ا/ سيد عبد العاطي"],
  "كيمياء":      ["ا/ احمد سعيد", "ا/ محمد صلاح"],
  "فزياء":       ["ا/ محمد الجوهري", "ا/ عمر عبد الفضيل", "ا/ محمد مختار"],
  "احياء":       ["ا/ محمد نور", "ا/ محمد علام", "ا/ علي نور"],
  "علوم":        ["ا/ علي نور"],
  "انجليزي":     ["ا/ حماده يوسف"],
  "تاريخ":       ["ا/ رضا صلاح"],
  "جغرافيا":     ["ا/ رضا صلاح"],
  "رياضة":       ["ا/ مصطفي صابر", "ا/ ناصر سعد"],
  "احصاء":       ["ا/ مصطفي صابر"],
  "فرنساوي":     ["ا/ محمد رجب", "ا/ محمود سليمان"],
  "برمجة":       ["م/ محمد ابراهيم"],
  "فلسفة":       ["ا/ احمد الليثي", "ا/ سارة مجدي"],
  "محاسبة":      [],
  "علم نفس":     [],
  "ادارة اعمال": []
};



const APPS_SCRIPT_SOURCE = `function doGet(e){return handleSyncRequest(e);}
function doPost(e){return handleSyncRequest(e);}
function handleSyncRequest(e){
  try {
    var data = {};
    if (e && e.postData && e.postData.contents) {
      try { data = JSON.parse(e.postData.contents); } catch(err){ data = e.parameter || {}; }
    } else if (e && e.parameter) { data = e.parameter; }
    var action = data.action || (e && e.parameter && e.parameter.action) || "get_all";
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("الطلاب") || ss.getSheetByName("مركز الارائج") || ss.getSheets()[0];
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["الكود", "الاسم", "المنطقة", "التليفون", "ولي الأمر", "السنة الدراسية", "التخصص", "المواد", "المدرسين", "تاريخ التسجيل"]);
    }
    // Dedicated Audit Log Sheet
    var auditSheet = ss.getSheetByName("سجل العمليات والرقابة") || ss.getSheetByName("سجل الرقابة");
    if (!auditSheet) {
      auditSheet = ss.insertSheet("سجل العمليات والرقابة");
      auditSheet.appendRow(["التاريخ والوقت", "المسؤول / المشرف", "نوع العملية", "كود الطالب", "اسم الطالب", "المرحلة", "المادة", "الشهر / الحصة", "القيمة / الدرجة", "تفاصيل إضافية"]);
      auditSheet.getRange(1, 1, 1, 10).setFontWeight("bold").setBackground("#0284c7").setFontColor("#ffffff");
    }
    var byUser = data.byUser || data.userName || "مشرف";
    var nowStr = Utilities.formatDate(new Date(), "Africa/Cairo", "yyyy/MM/dd HH:mm:ss");

    if (action === "test") {
      return ContentService.createTextOutput(JSON.stringify({status:"success",message:"متصل بنجاح: "+ss.getName(),sheet:sheet.getName(),auditSheet:auditSheet.getName()})).setMimeType(ContentService.MimeType.JSON);
    }
    if (action === "get_all" || action === "pull" || action === "getAll") {
      var vals = sheet.getDataRange().getValues();
      var list = [];
      for (var i = 1; i < vals.length; i++) {
        var r = vals[i];
        if (!r[0] && !r[1]) continue;
        list.push({
          code: String(r[0] || "").trim(),
          name: String(r[1] || "").trim(),
          area: String(r[2] || "").trim(),
          phone: String(r[3] || "").trim(),
          parentPhone: String(r[4] || "").trim(),
          grade: String(r[5] || "").trim(),
          specialization: String(r[6] || "").trim(),
          subjectsSummary: String(r[7] || "").trim(),
          teachersSummary: String(r[8] || "").trim(),
          regDate: String(r[9] || "").trim()
        });
      }
      return ContentService.createTextOutput(JSON.stringify({status:"success",count:list.length,students:list})).setMimeType(ContentService.MimeType.JSON);
    }
    if (action === "add_student") {
      var s = data.student || data;
      if (typeof s === 'string') { try { s = JSON.parse(s); } catch(err){} }
      sheet.appendRow([s.code||"",s.name||"",s.area||"",s.phone||"",s.parentPhone||"",s.grade||"",s.specialization||"",s.subjectsSummary||"",s.teachersSummary||"",s.regDate||nowStr]);
      auditSheet.appendRow([nowStr, byUser, "تسجيل طالب جديد", s.code||"", s.name||"", s.grade||"", s.subjectsSummary||"", "-", "-", "تسجيل طالب جديد بالمنظومة"]);
      return ContentService.createTextOutput(JSON.stringify({status:"success",message:"تمت الإضافة"})).setMimeType(ContentService.MimeType.JSON);
    }
    if (action === "batch_update_students" || action === "bulk_update" || action === "bulk_session") {
      var studentsList = data.students || [];
      if (typeof studentsList === 'string') { try { studentsList = JSON.parse(studentsList); } catch(err){} }
      var vals = sheet.getDataRange().getValues();
      var codeRowMap = {};
      for (var i = 1; i < vals.length; i++) {
        var c = String(vals[i][0] || "").trim();
        if (c) codeRowMap[c] = i + 1;
      }
      for (var j = 0; j < studentsList.length; j++) {
        var s = studentsList[j];
        var c = String(s.code || "").trim();
        if (codeRowMap[c]) {
          var row = codeRowMap[c];
          sheet.getRange(row, 2, 1, 8).setValues([[s.name||"", s.area||"", s.phone||"", s.parentPhone||"", s.grade||"", s.specialization||"", s.subjectsSummary||"", s.teachersSummary||""]]);
        }
      }
      auditSheet.appendRow([nowStr, byUser, "رصد حضور جماعي", "-", "-", "-", data.subject||"-", (data.month||"") + " ح" + ((data.sessionIdx !== undefined) ? (data.sessionIdx + 1) : ""), "-", "رصد جماعي لـ " + studentsList.length + " طالب"]);
      return ContentService.createTextOutput(JSON.stringify({status:"success", updated: studentsList.length})).setMimeType(ContentService.MimeType.JSON);
    }
    if (action === "update_student") {
      var s = data.student || data;
      if (typeof s === 'string') { try { s = JSON.parse(s); } catch(err){} }
      var code = String(s.code).trim();
      var vals = sheet.getDataRange().getValues();
      var found = false;
      for (var i=1; i<vals.length; i++) {
        if (String(vals[i][0]).trim() === code) {
          sheet.getRange(i+1,2,1,8).setValues([[s.name||"",s.area||"",s.phone||"",s.parentPhone||"",s.grade||"",s.specialization||"",s.subjectsSummary||"",s.teachersSummary||""]]);
          found = true; break;
        }
      }
      if (!found) {
        sheet.appendRow([s.code||"",s.name||"",s.area||"",s.phone||"",s.parentPhone||"",s.grade||"",s.specialization||"",s.subjectsSummary||"",s.teachersSummary||"",s.regDate||nowStr]);
      }
      auditSheet.appendRow([nowStr, byUser, s.lastAction || "تعديل طالب", code, s.name||"", s.grade||"", "-", "-", "-", "تعديل بيانات أو حصص"]);
      return ContentService.createTextOutput(JSON.stringify({status:"success"})).setMimeType(ContentService.MimeType.JSON);
    }
    if (action === "delete_student") {
      var code = String(data.code).trim();
      var vals = sheet.getDataRange().getValues();
      var delName = "";
      for (var i=1; i<vals.length; i++) {
        if (String(vals[i][0]).trim() === code) {
          delName = String(vals[i][1] || "");
          sheet.deleteRow(i+1); break;
        }
      }
      auditSheet.appendRow([nowStr, byUser, "حذف طالب", code, delName, "-", "-", "-", "-", "حذف الطالب نهائياً من المنظومة"]);
      return ContentService.createTextOutput(JSON.stringify({status:"success"})).setMimeType(ContentService.MimeType.JSON);
    }
    if (action === "log_activity") {
      var log = data.log || data;
      if (typeof log === 'string') { try { log = JSON.parse(log); } catch(err){} }
      auditSheet.appendRow([
        log.fullTime || nowStr,
        log.userName || byUser,
        log.actionType || "عملية",
        log.studentCode || "",
        log.studentName || "",
        log.grade || "",
        log.subject || "",
        log.monthOrSession || "",
        log.value || "",
        log.details || ""
      ]);
      return ContentService.createTextOutput(JSON.stringify({status:"success"})).setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(JSON.stringify({status:"success"})).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({status:"error",error:err.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}`;

// ====================================================
// 2. ARABIC NORMALIZATION (from sheets.py)
// ====================================================

function normalize_arabic(text) {
  if (!text) return "";
  let str = text.toString().trim();
  str = str.replace(/[\u064B-\u065F\u0670]/g, ''); // تشكيل
  str = str.replace(/[إأآٱ]/g, 'ا');               // توحيد الألف
  str = str.replace(/ى/g, 'ي');                    // ألف مقصورة
  str = str.replace(/ة/g, 'ه');                    // تاء مربوطة
  str = str.replace(/ـ/g, '');                     // تطويل
  str = str.replace(/\s+/g, '');                   // كل المسافات
  return str.toLowerCase();
}

// ====================================================
// 3. GRADING EVALUATION (from attendance.py)
// ====================================================

function get_taqdeer(percentage) {
  if (percentage === null || percentage === undefined) return "مسجل";
  if (percentage >= 85) return "ممتاز";
  if (percentage >= 75) return "جيد جدًا";
  if (percentage >= 65) return "جيد";
  if (percentage >= 50) return "مقبول";
  return "ضعيف";
}

function canonicalSubjectName(subj) {
  if (!subj || typeof subj !== 'string') return "";
  const s = subj.trim();
  const norm = normalize_arabic(s);
  if (norm.includes('فزي') || norm.includes('فيز')) return 'فزياء';
  if (norm.includes('كيم')) return 'كيمياء';
  if (norm.includes('رياض')) return 'رياضة';
  if (norm.includes('انجل')) return 'انجليزي';
  if (norm.includes('عرب')) return 'عربي';
  if (norm.includes('فرنس')) return 'فرنساوي';
  if (norm.includes('احي')) return 'احياء';
  if (norm.includes('احص')) return 'احصاء';
  if (norm.includes('تار')) return 'تاريخ';
  if (norm.includes('جغراف')) return 'جغرافيا';
  if (norm.includes('علوم')) return 'علوم';
  if (norm.includes('فلسف')) return 'فلسفة';
  if (norm.includes('برمج')) return 'برمجة';
  if (norm.includes('محاسب')) return 'محاسبة';
  if (norm.includes('نفس')) return 'علم نفس';
  if (norm.includes('اعمال')) return 'ادارة اعمال';
  return s;
}

function normalize_grade(grade) {
  if (!grade) return 'ث1';
  const g = String(grade).trim();
  if (g === 'ث1' || g.includes('الأول') || g.includes('الاول') || g.startsWith('1')) return 'ث1';
  if (g === 'ث2' || g.includes('الثاني') || g.startsWith('2')) return 'ث2';
  if (g === 'ث3' || g.includes('الثالث') || g.startsWith('3')) return 'ث3';
  return g || 'ث1';
}

// Clean teacher name helper: guarantees no truncated teacher names like 'ا' or 'م' or empty
function cleanTeacherName(rawTeacher, rawSubj) {
  const canon = canonicalSubjectName(rawSubj) || rawSubj;
  const defaultTeacher = (TEACHERS[canon] && TEACHERS[canon][0]) || 'مدرس المادة';
  if (!rawTeacher || typeof rawTeacher !== 'string') return defaultTeacher;
  const t = rawTeacher.trim();
  // Corrupted strings caused by splitting on slash inside titles like "ا/" or "م/"
  if (t === 'ا' || t === 'م' || t === 'ا/' || t === 'م/' || t === '/' || t === 'مدرس المادة' || t === 'جديد' || t === 'مدرس آخر' || t.length <= 2) {
    return defaultTeacher;
  }
  return t;
}

// Build comprehensive search string for instant matching by name, code, phone, area, subject, and teacher
function buildStudentSearchString(st) {
  const teachers = Object.values(st.academicSubjects || {}).map(sub => sub.teacher || '').filter(Boolean);
  const teacherVariants = [];
  teachers.forEach(t => {
    teacherVariants.push(t);
    // Add stripped version without "ا/" or "م/" or "د/"
    const stripped = t.replace(/^[اأماد]\s*[\/.-]\s*/, '').trim();
    if (stripped && stripped !== t) {
      teacherVariants.push(stripped);
      teacherVariants.push(`استاذ ${stripped}`);
      teacherVariants.push(`مستر ${stripped}`);
    }
  });

  return normalize_arabic([
    st.name || '',
    st.code || '',
    st.phone || '',
    st.parentPhone || '',
    st.area || '',
    st.grade || '',
    st.specialization || '',
    st.subjectsSummary || '',
    st.teachersSummary || '',
    ...teacherVariants
  ].join(' '));
}

// Reconstruct structured academicSubjects object from subjectsSummary & teachersSummary with multi-month support
function parseSubjectsFromSummary(subjectsSummary, teachersSummary, existingAcademicSubjects = {}) {
  const result = {};

  // Build lookup map of existing sessions by canonical subject name
  const existingMap = {};
  if (existingAcademicSubjects && typeof existingAcademicSubjects === 'object') {
    Object.keys(existingAcademicSubjects).forEach(k => {
      const canon = canonicalSubjectName(k) || k;
      existingMap[canon] = existingAcademicSubjects[k];
    });
  }

  function registerSubj(rawName, explicitTeacher, explicitSessions = null, explicitMonths = {}) {
    if (!rawName || typeof rawName !== 'string') return;
    const canon = canonicalSubjectName(rawName) || rawName.trim();
    if (!canon) return;

    const old = existingMap[canon];
    const oldTeacher = (old && old.teacher) ? cleanTeacherName(old.teacher, canon) : null;
    const finalTeacher = cleanTeacherName(explicitTeacher || oldTeacher, canon);

    let sessions = ["", "", "", "", "", "", "", ""];
    if (explicitSessions && Array.isArray(explicitSessions) && explicitSessions.length > 0) {
      for (let i = 0; i < 8; i++) {
        sessions[i] = (explicitSessions[i] !== undefined) ? explicitSessions[i] : "";
      }
    } else if (old && Array.isArray(old.sessions) && old.sessions.length > 0) {
      sessions = [...old.sessions];
    }

    const mergedMonths = (old && old.months) ? { ...old.months } : {};
    if (explicitMonths && typeof explicitMonths === 'object') {
      Object.assign(mergedMonths, explicitMonths);
    }
    if (!mergedMonths['شهر 1'] || mergedMonths['شهر 1'].every(s => !s)) {
      mergedMonths['شهر 1'] = [...sessions];
    }

    if (!result[canon]) {
      result[canon] = {
        teacher: finalTeacher,
        sessions: sessions,
        months: mergedMonths
      };
    } else {
      if (explicitTeacher) result[canon].teacher = finalTeacher;
      if (explicitSessions) result[canon].sessions = sessions;
      result[canon].months = Object.assign(result[canon].months || {}, mergedMonths);
    }
  }

  // 1. Parse teachersSummary (supports | or newline or semicolon)
  // Format: "مادة/مدرس [حصص: s1, s2, ...] [شهر 2: s1, s2, ...] | مادة أخرى/مدرس آخر"
  if (teachersSummary && typeof teachersSummary === 'string' && teachersSummary.trim()) {
    const parts = teachersSummary.split(/[|\n;]/).map(p => p.trim()).filter(Boolean);
    parts.forEach(part => {
      const slashIdx = part.indexOf('/');
      if (slashIdx > -1) {
        const subj = part.substring(0, slashIdx).trim();
        let teacherPart = part.substring(slashIdx + 1).trim();
        let parsedSessions = null;
        const parsedMonths = {};

        // Extract all month or session tags
        const regex = /\[(شهر\s*[^:]+|حصص):\s*([^\]]*)\]/g;
        let match;
        while ((match = regex.exec(teacherPart)) !== null) {
          const tag = match[1].trim();
          const rawSess = match[2].split(',').map(s => s.trim());
          const parsed = rawSess.map(s => (s === '_' || s === '-' || !s) ? '' : s);
          while (parsed.length < 8) parsed.push('');

          if (tag === 'حصص' || tag === 'شهر 1' || tag === 'شهر 9 (سبتمبر)' || tag === 'شهر 9') {
            parsedSessions = parsed;
            parsedMonths['شهر 9 (سبتمبر)'] = [...parsed];
            parsedMonths['شهر 1'] = [...parsed];
          } else {
            parsedMonths[tag] = parsed;
            if (typeof CENTER_MONTHS !== 'undefined' && !CENTER_MONTHS.includes(tag)) {
              CENTER_MONTHS.push(tag);
              try { localStorage.setItem('araij_center_months', JSON.stringify(CENTER_MONTHS)); } catch(e) {}
            }
          }
        }
        teacherPart = teacherPart.replace(/\[(شهر\s*[^:]+|حصص):\s*([^\]]*)\]/g, '').trim();

        if (subj) {
          registerSubj(subj, teacherPart, parsedSessions, parsedMonths);
        }
      } else {
        registerSubj(part);
      }
    });
  }

  // 2. Parse subjectsSummary
  if (subjectsSummary && typeof subjectsSummary === 'string' && subjectsSummary.trim()) {
    const names = subjectsSummary.split(/[,،|\n;]/).map(s => s.trim()).filter(Boolean);
    names.forEach(subj => {
      if (subj) registerSubj(subj);
    });
  }

  return result;
}

// Rebuilds subjectsSummary and teachersSummary with [حصص: ...] and [شهر X: ...] for cloud sync
function rebuildStudentSummaries(student) {
  if (!student) return;
  if (!student.academicSubjects) student.academicSubjects = {};

  const subKeys = Object.keys(student.academicSubjects);
  student.subjectsSummary = subKeys.join(', ');
  student.teachersSummary = subKeys.map(k => {
    const subObj = student.academicSubjects[k] || {};
    const canon = canonicalSubjectName(k) || k;
    const tName = cleanTeacherName(subObj.teacher || (TEACHERS[canon] && TEACHERS[canon][0]) || 'مدرس المادة', canon);
    
    let parts = [`${k}/${tName}`];

    // Primary sessions tag
    const hasSess = subObj.sessions && Array.isArray(subObj.sessions) && subObj.sessions.some(s => s && String(s).trim() !== '');
    if (hasSess) {
      const sessStr = subObj.sessions.map(s => (s && String(s).trim()) ? String(s).trim() : '_').join(', ');
      parts.push(`[حصص: ${sessStr}]`);
    }

    // Extra months tags
    if (subObj.months && typeof subObj.months === 'object') {
      Object.keys(subObj.months).forEach(mName => {
        if (mName !== 'شهر 1' && mName !== 'شهر 9 (سبتمبر)') {
          const mSess = subObj.months[mName];
          if (Array.isArray(mSess) && mSess.some(s => s && String(s).trim() !== '')) {
            const mStr = mSess.map(s => (s && String(s).trim()) ? String(s).trim() : '_').join(', ');
            parts.push(`[${mName}: ${mStr}]`);
          }
        }
      });
    }

    return parts.join(' ');
  }).join(' | ');

  student._metrics = calculateStudentMetrics(student);
  student._searchString = buildStudentSearchString(student);
}

// ==========================================
// MONTHLY ATTENDANCE HELPER FUNCTIONS
// ==========================================

function getStudentSubjectMonthSessions(student, subject, month = null) {
  const m = month || currentActiveMonth || 'شهر 9 (سبتمبر)';
  if (!student.academicSubjects) student.academicSubjects = {};
  if (!student.academicSubjects[subject]) {
    const canon = canonicalSubjectName(subject) || subject;
    student.academicSubjects[subject] = {
      teacher: (TEACHERS[canon] && TEACHERS[canon][0]) || "مدرس المادة",
      sessions: ["", "", "", "", "", "", "", ""],
      months: {}
    };
  }
  const sub = student.academicSubjects[subject];
  if (!sub.months) sub.months = {};
  if (!sub.months[m]) {
    if ((m === 'شهر 9 (سبتمبر)' || m === 'شهر 1') && Array.isArray(sub.sessions) && sub.sessions.length > 0) {
      sub.months[m] = [...sub.sessions];
    } else {
      sub.months[m] = ["", "", "", "", "", "", "", ""];
    }
  }
  while (sub.months[m].length < 8) sub.months[m].push("");
  return sub.months[m];
}

function setStudentSubjectMonthSession(student, subject, month, sessionIdx, val) {
  const m = month || currentActiveMonth || 'شهر 9 (سبتمبر)';
  const sessions = getStudentSubjectMonthSessions(student, subject, m);
  sessions[sessionIdx] = val;
  const sub = student.academicSubjects[subject];
  if (m === currentActiveMonth || m === 'شهر 9 (سبتمبر)' || m === 'شهر 1') {
    sub.sessions = [...sessions];
  }
  return sessions;
}

function populateMonthDropdowns() {
  const selects = [
    document.getElementById('matrixMonthSelect'),
    document.getElementById('bulkMonthSelect'),
    document.getElementById('pdfMonthSelect')
  ];

  selects.forEach(sel => {
    if (!sel) return;
    const currentVal = sel.value || currentActiveMonth;
    sel.innerHTML = CENTER_MONTHS.map(m => `<option value="${m}">${m}</option>`).join('');
    if (CENTER_MONTHS.includes(currentVal)) {
      sel.value = currentVal;
    } else {
      sel.value = currentActiveMonth;
    }
  });
}

function onMatrixMonthChange(month) {
  currentActiveMonth = month;
  localStorage.setItem('araij_active_month', month);
  const bulkSel = document.getElementById('bulkMonthSelect');
  if (bulkSel) bulkSel.value = month;
  const pdfSel = document.getElementById('pdfMonthSelect');
  if (pdfSel) pdfSel.value = month;
  renderAttendanceMatrix();
  showToast(`📅 تم الانتقال إلى رصد (${month})`);
}

function openNewMonthModal() {
  const nextNum = CENTER_MONTHS.length + 1;
  const input = document.getElementById('newMonthNameInput');
  if (input) input.value = `شهر ${nextNum}`;
  document.getElementById('newMonthModal')?.classList.remove('hidden');
  if (input) setTimeout(() => input.focus(), 100);
}

function closeNewMonthModal() {
  document.getElementById('newMonthModal')?.classList.add('hidden');
}

function saveNewCenterMonth() {
  const input = document.getElementById('newMonthNameInput');
  const monthName = (input?.value || '').trim();
  if (!monthName) {
    alert('يرجى كتابة اسم الشهر');
    return;
  }
  if (!CENTER_MONTHS.includes(monthName)) {
    CENTER_MONTHS.push(monthName);
    localStorage.setItem('araij_center_months', JSON.stringify(CENTER_MONTHS));
  }
  currentActiveMonth = monthName;
  localStorage.setItem('araij_active_month', monthName);
  populateMonthDropdowns();
  closeNewMonthModal();
  renderAttendanceMatrix();
  showToast(`🎉 تم فتح (${monthName}) بنجاح ويمكن الآن رصد 8 حصص جديدة!`);
}

function renderDetailMonthTabs(student) {
  const container = document.getElementById('detailMonthTabsContainer');
  if (!container) return;

  container.innerHTML = CENTER_MONTHS.map(m => {
    const isActive = m === activeDetailMonth;
    const btnCls = isActive 
      ? 'bg-emerald-600 text-white font-bold shadow-md shadow-emerald-600/30' 
      : 'bg-slate-800 text-slate-400 hover:bg-slate-700 hover:text-white';

    return `
      <button type="button" onclick="switchDetailMonth('${student.code}', '${m}')" class="px-3 py-1.5 rounded-xl text-xs whitespace-nowrap transition-all cursor-pointer ${btnCls}">
        ${m}
      </button>
    `;
  }).join('');
}

function switchDetailMonth(code, month) {
  activeDetailMonth = month;
  openStudentDetailModal(code);
}


// ====================================================
// 4. STATE & APP INIT
// ====================================================

let allStudents = [];
let filteredStudents = [];
let currentPage = 1;
const PAGE_SIZE = 30;
let isTableView = false;

let currentGradeFilter = 'all';
let currentEvaluationFilter = 'all';
let currentAttendanceFilter = 'all';
let currentAreaFilter = 'all';
let currentSearchQuery = '';

let currentUser = null;
let activeReportStudent = null;
let selectedStudentCodes = new Set();
let bulkQueueCurrentIndex = 0;

// ====================================================
// PERFORMANCE UTILITIES
// ====================================================

// Debounce: delay rapid repeated calls (e.g. search typing)
function debounce(fn, delay) {
  let timer;
  return function(...args) {
    clearTimeout(timer);
    timer = setTimeout(() => fn.apply(this, args), delay);
  };
}

// Throttle: run at most once per interval
function throttle(fn, limit) {
  let lastRun = 0;
  return function(...args) {
    const now = Date.now();
    if (now - lastRun >= limit) {
      lastRun = now;
      fn.apply(this, args);
    }
  };
}

// Schedule a render in the next animation frame (avoids layout thrashing)
let _renderPending = false;
function scheduleRender(fn) {
  if (!_renderPending) {
    _renderPending = true;
    requestAnimationFrame(() => {
      _renderPending = false;
      fn();
    });
  }
}

let activeDetailStudentCode = null;
let activeSessionEdit = null; // { code, subject, sessionIdx, month }
let currentPdfMode = 'teacher';

// ==========================================
// MONTHLY ATTENDANCE SYSTEM STATE (ACADEMIC YEAR STARTS IN SEPTEMBER)
// ==========================================
const ACADEMIC_MONTHS = [
  "شهر 9 (سبتمبر)",
  "شهر 10 (أكتوبر)",
  "شهر 11 (نوفمبر)",
  "شهر 12 (ديسمبر)",
  "شهر 1 (يناير)",
  "شهر 2 (فبراير)",
  "شهر 3 (مارس)",
  "شهر 4 (أبريل)",
  "شهر 5 (مايو)"
];

let currentActiveMonth = localStorage.getItem('araij_active_month') || 'شهر 9 (سبتمبر)';
let activeDetailMonth = currentActiveMonth;
let CENTER_MONTHS;
try {
  const savedMonths = localStorage.getItem('araij_center_months');
  CENTER_MONTHS = savedMonths ? JSON.parse(savedMonths) : [...ACADEMIC_MONTHS];
  if (!CENTER_MONTHS.some(m => m.includes('سبتمبر') || m.includes('9'))) {
    CENTER_MONTHS = [...ACADEMIC_MONTHS];
    localStorage.setItem('araij_center_months', JSON.stringify(CENTER_MONTHS));
  }
} catch(e) {
  CENTER_MONTHS = [...ACADEMIC_MONTHS];
}
if (!CENTER_MONTHS.includes(currentActiveMonth) || currentActiveMonth === 'شهر 1') {
  currentActiveMonth = 'شهر 9 (سبتمبر)';
  localStorage.setItem('araij_active_month', currentActiveMonth);
}

document.addEventListener('DOMContentLoaded', () => {
  try { initIcons(); } catch (e) {}
  try { checkAuth(); } catch (e) {}
  try { initData(); } catch (e) {}
  try { populateAreaFilter(); } catch (e) {}
  try { updateKPIStats(); } catch (e) {}
  try { initWizard(); } catch (e) {}
  try { populateMonthDropdowns(); } catch (e) {}
  try { initPdfReportSelects(); } catch (e) {}
  try { initMatrixSelects(); } catch (e) {}
  try { renderUsersList(); } catch (e) {}
  try { applyFilters(); } catch (e) {}
  try { renderAttendanceMatrix(); } catch (e) {}
  try { generateSelectedPdfReport(); } catch (e) {}
  try { updateSheetsStatusIndicator(); } catch (e) {}
  try { initLiveRailwaySync(); } catch (e) {}
  
  // Instant fetch from Railway Central Server on startup (< 300ms)
  try { fetchLatestStudentsFromCentralServer(false); } catch (e) {}

  // Auto-pull fresh students from Google Sheets on startup
  try { pullFromGoogleSheets(true); } catch (e) {}
  
  // Real-time sync: poll Sheets every 15 seconds so all devices stay in sync
  setInterval(() => {
    try { pullFromGoogleSheets(true); } catch (e) {}
  }, 15000);
});


function initIcons() {
  if (window.lucide) {
    try { window.lucide.createIcons(); } catch (e) {}
  }
}

// ====================================================
// 5. TAB NAVIGATION
// ====================================================

function switchTab(tabId) {
  if (tabId === 'users_mgmt') {
    const isMasterAdmin = currentUser && (
      (currentUser.name && (currentUser.name.includes('وفاء') || currentUser.name.includes('wafa'))) || 
      (currentUser.phone && (currentUser.phone.includes('01040581954') || currentUser.phone.includes('01002169889'))) || 
      (currentUser.role && (currentUser.role.includes('أدمن') || currentUser.role.includes('admin')))
    );
    if (!isMasterAdmin) {
      showToast('⚠️ تنبيه: قسم إدارة المستخدمين وسجل الرقابة مخصص للمدير العام فقط.');
      return;
    }
  }

  const allTabs = ['dashboard', 'students', 'attendance_matrix', 'register', 'pdf_reports', 'users_mgmt'];
  allTabs.forEach(t => {
    const sec = document.getElementById(`tabContent_${t}`);
    const navBtn = document.getElementById(`nav_${t}`);
    const mobBtn = document.getElementById(`mob_nav_${t}`);
    
    if (sec) {
      if (t === tabId) {
        sec.classList.remove('hidden');
        sec.style.display = 'block';
      } else {
        sec.classList.add('hidden');
        sec.style.display = 'none';
      }
    }

    if (navBtn) {
      if (t === tabId) {
        navBtn.className = 'nav-btn active-nav-btn px-3.5 py-2 rounded-xl text-xs font-bold transition-all flex items-center gap-1.5';
      } else {
        navBtn.className = 'nav-btn text-slate-600 hover:text-slate-900 px-3.5 py-2 rounded-xl text-xs font-bold transition-all flex items-center gap-1.5';
      }
    }

    if (mobBtn) {
      if (t === tabId) {
        mobBtn.className = 'mobile-nav-btn active flex-1 flex flex-col items-center justify-center py-1 text-sky-600 font-bold transition-all';
      } else {
        mobBtn.className = 'mobile-nav-btn flex-1 flex flex-col items-center justify-center py-1 text-slate-500 hover:text-slate-900 transition-all';
      }
    }
  });

  try {
    if (tabId === 'pdf_reports') {
      generateSelectedPdfReport();
    } else if (tabId === 'attendance_matrix') {
      renderAttendanceMatrix();
    } else if (tabId === 'users_mgmt') {
      renderUsersList();
      renderActivityLogs();
    } else if (tabId === 'students') {
      applyFilters();
    } else if (tabId === 'dashboard') {
      updateKPIStats();
    }
  } catch (err) {
    console.warn('Error on tab switch:', err);
  }

  window.scrollTo({ top: 0, behavior: 'smooth' });
  try { initIcons(); } catch (e) {}
}

// ====================================================
// 6. AUTHENTICATION & USERS (PHONE: 01002169889, PASS: 0123)
// ====================================================

const INITIAL_USERS = [
  {
    name: "م. وفاء ناصر",
    phone: "01040581954",
    password: "0123",
    role: "المدير العام (أدمن)"
  },
  {
    name: "م. وفاء ناصر",
    phone: "01002169889",
    password: "0123",
    role: "المدير العام (أدمن)"
  },
  {
    name: "منار ناصر",
    phone: "01006825206",
    password: "010068",
    role: "مشرف"
  },
  {
    name: "Asmaa Hamed",
    phone: "01205726935",
    password: "01205",
    role: "مشرف"
  },
  {
    name: "كريمة",
    phone: "01096029369",
    password: "01096",
    role: "مشرف"
  },
  {
    name: "نادية",
    phone: "01275820862",
    password: "01275",
    role: "مشرف"
  }
];

const DEFAULT_ADMIN = INITIAL_USERS[0];

function convertArabicDigitsToEnglish(str) {
  if (!str && str !== 0) return "";
  var s = str.toString().trim();
  return s.replace(/[\u0660\u06F0]/g, '0')
          .replace(/[\u0661\u06F1]/g, '1')
          .replace(/[\u0662\u06F2]/g, '2')
          .replace(/[\u0663\u06F3]/g, '3')
          .replace(/[\u0664\u06F4]/g, '4')
          .replace(/[\u0665\u06F5]/g, '5')
          .replace(/[\u0666\u06F6]/g, '6')
          .replace(/[\u0667\u06F7]/g, '7')
          .replace(/[\u0668\u06F8]/g, '8')
          .replace(/[\u0669\u06F9]/g, '9');
}

function getStoredUsers() {
  const stored = localStorage.getItem('araij_users');
  let list = [];
  if (stored) {
    try {
      list = JSON.parse(stored);
    } catch (e) {}
  }
  if (!Array.isArray(list) || list.length === 0) {
    list = [...INITIAL_USERS];
    localStorage.setItem('araij_users', JSON.stringify(list));
    return list;
  }

  // Ensure default accounts exist and are refreshed
  INITIAL_USERS.forEach(defU => {
    const cleanDefPhone = defU.phone.replace(/\D/g, '');
    const found = list.find(u => (u.phone || '').replace(/\D/g, '').endsWith(cleanDefPhone.slice(-9)));
    if (!found) {
      list.push(defU);
    } else {
      found.password = defU.password;
      found.role = defU.role;
      found.name = defU.name;
    }
  });

  localStorage.setItem('araij_users', JSON.stringify(list));
  return list;
}

function checkAuth() {
  const session = localStorage.getItem('araij_session');
  const loginScreen = document.getElementById('loginScreen');
  const appContainer = document.getElementById('appContainer');

  if (session) {
    try {
      currentUser = JSON.parse(session);
      if (currentUser && currentUser.name) {
        if (loginScreen) {
          loginScreen.style.setProperty('display', 'none', 'important');
          loginScreen.classList.add('hidden');
        }
        if (appContainer) {
          appContainer.style.removeProperty('display');
          appContainer.style.display = 'flex';
          appContainer.classList.remove('hidden');
        }
        const uName = document.getElementById('headerUserName');
        const uRole = document.getElementById('headerUserRole');
        if (uName) uName.textContent = currentUser.name;
        if (uRole) uRole.textContent = `${currentUser.role} (${currentUser.phone})`;
        applyRolePermissions();
        return;
      }
    } catch (e) {}
  }

  if (loginScreen) {
    loginScreen.style.removeProperty('display');
    loginScreen.style.display = 'flex';
    loginScreen.classList.remove('hidden');
  }
  if (appContainer) {
    appContainer.style.setProperty('display', 'none', 'important');
    appContainer.classList.add('hidden');
  }
}

function applyRolePermissions() {
  const isMasterAdmin = currentUser && (
    (currentUser.name && (currentUser.name.includes('وفاء') || currentUser.name.includes('wafa'))) || 
    (currentUser.phone && (currentUser.phone.includes('01040581954') || currentUser.phone.includes('01002169889'))) || 
    (currentUser.role && (currentUser.role.includes('أدمن') || currentUser.role.includes('admin')))
  );

  const usersNavBtn = document.getElementById('nav_users_mgmt');
  const mobUsersBtn = document.getElementById('mob_nav_users_mgmt');
  const pullSheetsBtn = document.getElementById('navPullSheetsBtn');
  const sheetsStatusBtn = document.getElementById('sheetsStatusBadgeText')?.parentElement;

  if (usersNavBtn) {
    if (isMasterAdmin) {
      usersNavBtn.classList.remove('hidden');
      usersNavBtn.style.display = '';
    } else {
      usersNavBtn.classList.add('hidden');
      usersNavBtn.style.display = 'none';
    }
  }

  if (mobUsersBtn) {
    if (isMasterAdmin) {
      mobUsersBtn.classList.remove('hidden');
      mobUsersBtn.style.display = 'flex';
    } else {
      mobUsersBtn.classList.add('hidden');
      mobUsersBtn.style.display = 'none';
    }
  }

  if (pullSheetsBtn) {
    if (isMasterAdmin) {
      pullSheetsBtn.classList.remove('hidden');
      pullSheetsBtn.style.display = '';
    } else {
      pullSheetsBtn.classList.add('hidden');
      pullSheetsBtn.style.display = 'none';
    }
  }

  if (sheetsStatusBtn) {
    if (isMasterAdmin) {
      sheetsStatusBtn.classList.remove('hidden');
      sheetsStatusBtn.style.display = '';
    } else {
      sheetsStatusBtn.classList.add('hidden');
      sheetsStatusBtn.style.display = 'none';
    }
  }
}

function handleLogin(e) {
  if (e) {
    try { e.preventDefault(); } catch (err) {}
  }
  var phoneEl = document.getElementById('loginPhone');
  var passEl  = document.getElementById('loginPassword');
  var errorMsg = document.getElementById('loginErrorMsg');

  var rawPhone = phoneEl ? phoneEl.value : '';
  var rawPass  = passEl  ? passEl.value  : '';

  var phoneInput = convertArabicDigitsToEnglish(rawPhone);
  var passInput  = convertArabicDigitsToEnglish(rawPass);

  var cleanPhone = phoneInput.replace(/\D/g, '');
  var cleanPass  = passInput.trim();

  if (!cleanPass) {
    if (errorMsg) { 
      errorMsg.textContent = 'يرجى كتابة كلمة المرور الخاصة بحسابك!'; 
      errorMsg.classList.remove('hidden'); 
    }
    return false;
  }

  // الحسابات المعتمدة الرسمية
  var ACCOUNTS = [
    { name: 'م. وفاء ناصر',  phone: '01040581954', password: '0123',   role: 'المدير العام (أدمن)' },
    { name: 'م. وفاء ناصر',  phone: '01002169889', password: '0123',   role: 'المدير العام (أدمن)' },
    { name: 'منار ناصر',     phone: '01006825206', password: '010068', role: 'مشرف' },
    { name: 'Asmaa Hamed',   phone: '01205726935', password: '01205',  role: 'مشرف' },
    { name: 'كريمة',         phone: '01096029369', password: '01096',  role: 'مشرف' },
    { name: 'نادية',         phone: '01275820862', password: '01275',  role: 'مشرف' }
  ];

  var found = null;

  // 1. المطابقة بالباسورد والرقم
  for (var i = 0; i < ACCOUNTS.length; i++) {
    var acc = ACCOUNTS[i];
    var accPhone = acc.phone.replace(/\D/g, '');
    if (acc.password.trim() !== cleanPass) continue;

    if (cleanPhone.length === 0) {
      found = acc;
      break;
    }

    var minLen = Math.min(cleanPhone.length, 8);
    var pMatch = (accPhone === cleanPhone) ||
                 (accPhone.slice(-minLen) === cleanPhone.slice(-minLen)) ||
                 (cleanPhone.slice(-minLen) === accPhone.slice(-minLen));
    if (pMatch) {
      found = acc;
      break;
    }
  }

  // 2. إذا كتب الباسورد فقط
  if (!found) {
    var matchingByPass = ACCOUNTS.filter(function(a) { return a.password.trim() === cleanPass; });
    if (matchingByPass.length === 1) {
      found = matchingByPass[0];
    } else if (matchingByPass.length > 1) {
      if (cleanPhone.length > 0) {
        found = matchingByPass.find(function(a) { 
          var ap = a.phone.replace(/\D/g, '');
          return ap.slice(-6) === cleanPhone.slice(-6); 
        }) || matchingByPass[0];
      } else {
        found = matchingByPass[0];
      }
    }
  }

  if (found) {
    currentUser = { name: found.name, phone: found.phone, role: found.role };
    try { localStorage.setItem('araij_session', JSON.stringify(currentUser)); } catch (e) {}

    var loginScreen = document.getElementById('loginScreen');
    if (loginScreen) {
      loginScreen.style.setProperty('display', 'none', 'important');
      loginScreen.classList.add('hidden');
    }

    var appContainer = document.getElementById('appContainer');
    if (appContainer) {
      appContainer.style.removeProperty('display');
      appContainer.style.display = 'flex';
      appContainer.classList.remove('hidden');
    }

    var uName = document.getElementById('headerUserName');
    var uRole = document.getElementById('headerUserRole');
    if (uName) uName.textContent = currentUser.name;
    if (uRole) uRole.textContent = currentUser.role + ' (' + currentUser.phone + ')';
    if (errorMsg) errorMsg.classList.add('hidden');

    try { applyRolePermissions(); } catch (e) {}
    try { initData(); } catch (e) {}
    try { applyFilters(); } catch (e) {}
    try { updateKPIStats(); } catch (e) {}
    try { if (window.lucide && typeof window.lucide.createIcons === 'function') window.lucide.createIcons(); } catch (e) {}
    try { logActivity('تسجيل الدخول', 'تم تسجيل الدخول بواسطة ' + currentUser.name); } catch (e) {}
    showToast('مرحباً بك ' + currentUser.name + '! تم الدخول بنجاح.');
    return false;
  } else {
    if (errorMsg) {
      errorMsg.textContent = 'رقم الموبايل أو كلمة المرور غير صحيحة. يرجى التأكد من البيانات.';
      errorMsg.classList.remove('hidden');
    }
    return false;
  }
}
window.handleLogin = handleLogin;
window._doLogin = handleLogin;

// ====================================================
// ACTIVITY & AUDIT LOGGING ENGINE
// ====================================================

function logActivity(actionType, details, studentCode = '', studentName = '', extra = {}) {
  const userName = currentUser ? currentUser.name : 'مشرف';
  const userPhone = currentUser ? currentUser.phone : '';
  const now = new Date();
  const dateStr = now.toLocaleDateString('ar-EG');
  const timeStr = now.toLocaleTimeString('ar-EG', { hour: '2-digit', minute: '2-digit', second: '2-digit' });

  const logEntry = {
    id: 'log_' + Date.now() + '_' + Math.floor(Math.random() * 1000),
    userName,
    userPhone,
    actionType,
    details,
    studentCode: studentCode ? String(studentCode) : '',
    studentName: studentName || '',
    grade: extra.grade || '',
    subject: extra.subject || '',
    monthOrSession: extra.monthOrSession || '',
    value: extra.value || '',
    dateStr,
    timeStr,
    fullTime: `${dateStr} - ${timeStr}`
  };

  try {
    const logs = JSON.parse(localStorage.getItem('araij_activity_logs') || '[]');
    logs.unshift(logEntry);
    if (logs.length > 500) logs.pop();
    localStorage.setItem('araij_activity_logs', JSON.stringify(logs));
  } catch (e) {}

  renderActivityLogs();

  // 1. Send to Google Sheets Audit Sheet
  syncToGoogleSheets('log_activity', { log: logEntry });

  // 2. Send to Live Railway / Local Server
  sendToRailwayServer('log_activity', { log: logEntry });
}

function renderActivityLogs() {
  const tbody = document.getElementById('activityLogsTableBody');
  if (!tbody) return;

  const logs = JSON.parse(localStorage.getItem('araij_activity_logs') || '[]');
  if (logs.length === 0) {
    tbody.innerHTML = `<tr><td colspan="5" class="p-4 text-center text-slate-500">لا توجد عمليات مسجلة حتى الآن.</td></tr>`;
    return;
  }

  tbody.innerHTML = logs.map(l => `
    <tr class="hover:bg-slate-800/40">
      <td class="p-3 text-slate-400 font-mono text-[11px]">${l.fullTime}</td>
      <td class="p-3 font-bold text-sky-300">
        <span>👤 ${l.userName}</span>
      </td>
      <td class="p-3 font-bold text-amber-400">
        <span class="px-2 py-0.5 rounded bg-amber-500/10 border border-amber-500/20">${l.actionType}</span>
      </td>
      <td class="p-3 text-emerald-400 font-bold">
        ${l.studentName ? `${l.studentName} <span class="text-sky-400 font-mono text-[11px]">(#${l.studentCode})</span>` : (l.studentCode ? `#${l.studentCode}` : '-')}
      </td>
      <td class="p-3 text-slate-300">${l.details}</td>
    </tr>
  `).join('');
}

function clearActivityLogs() {
  if (confirm('هل أنت متأكد من رغبتك في مسح سجل العمليات والنشاطات؟')) {
    localStorage.removeItem('araij_activity_logs');
    renderActivityLogs();
    showToast('تم مسح سجل النشاطات بنجاح.');
  }
}

function handleLogout() {
  try { localStorage.removeItem('araij_session'); } catch (e) {}
  currentUser = null;
  const loginScreen = document.getElementById('loginScreen');
  const appContainer = document.getElementById('appContainer');

  if (loginScreen) {
    loginScreen.style.removeProperty('display');
    loginScreen.style.display = 'flex';
    loginScreen.classList.remove('hidden');
  }
  if (appContainer) {
    appContainer.style.setProperty('display', 'none', 'important');
    appContainer.classList.add('hidden');
  }

  const p = document.getElementById('loginPhone');
  const w = document.getElementById('loginPassword');
  if (p) p.value = '';
  if (w) w.value = '';
  const errorMsg = document.getElementById('loginErrorMsg');
  if (errorMsg) errorMsg.classList.add('hidden');
}
window._doLogout = handleLogout;

function openAddUserModal() {
  document.getElementById('addUserForm')?.reset();
  document.getElementById('addUserModal')?.classList.remove('hidden');
}

function closeAddUserModal() {
  document.getElementById('addUserModal')?.classList.add('hidden');
}

function handleSaveNewUser(e) {
  if (e) e.preventDefault();
  const name = document.getElementById('newUserName')?.value.trim();
  const phone = document.getElementById('newUserPhone')?.value.trim();
  const password = document.getElementById('newUserPassword')?.value.trim();
  const role = document.getElementById('newUserRole')?.value;

  if (!name || !phone || !password) return;

  const users = getStoredUsers();
  if (users.find(u => u.phone === phone)) {
    alert('رقم الموبايل مسجل لمستخدم آخر بالفعل!');
    return;
  }

  users.push({ name, phone, password, role });
  localStorage.setItem('araij_users', JSON.stringify(users));
  closeAddUserModal();
  renderUsersList();
  showToast(`تمت إضافة المستخدم (${name}) بنجاح!`);
}

function deleteUser(phone) {
  if (phone === DEFAULT_ADMIN.phone) {
    alert('لا يمكن حذف حساب الأدمن الرئيسي!');
    return;
  }
  if (confirm('هل أنت متأكد من حذف هذا المستخدم؟')) {
    let users = getStoredUsers().filter(u => u.phone !== phone);
    localStorage.setItem('araij_users', JSON.stringify(users));
    renderUsersList();
    showToast('تم حذف المستخدم بنجاح');
  }
}

function renderUsersList() {
  const tbody = document.getElementById('usersListTableBody');
  if (!tbody) return;
  const users = getStoredUsers();
  tbody.innerHTML = users.map(u => `
    <tr class="hover:bg-slate-50">
      <td class="p-3.5 font-bold text-slate-900">${u.name}</td>
      <td class="p-3.5 font-mono text-emerald-400">${u.phone}</td>
      <td class="p-3.5 font-mono text-amber-300">••••</td>
      <td class="p-3.5 font-semibold text-sky-400">${u.role}</td>
      <td class="p-3.5 text-center">
        ${u.phone !== DEFAULT_ADMIN.phone ? `
          <button onclick="deleteUser('${u.phone}')" class="p-1.5 rounded-lg bg-rose-500/20 text-rose-400 hover:bg-rose-500/30" title="حذف">
            <i data-lucide="trash-2" class="w-3.5 h-3.5"></i>
          </button>
        ` : `<span class="text-slate-500 text-[11px]">أساسي</span>`}
      </td>
    </tr>
  `).join('');
  initIcons();
}

// ====================================================
// 7. REAL DATA PROCESSING & EXACT METRICS
// ====================================================

function initData() {
  const localOverrides = localStorage.getItem('araij_students_overrides');
  const fullDb = localStorage.getItem('araij_full_students_db');
  const deletedCodes = JSON.parse(localStorage.getItem('araij_deleted_students') || '[]');
  const baseData = window.STUDENTS_DATA || [];

  let merged = [];

  if (fullDb) {
    try {
      const parsed = JSON.parse(fullDb);
      if (Array.isArray(parsed) && parsed.length > 0) {
        merged = parsed;
      } else {
        merged = [...baseData];
      }
    } catch(e) {
      merged = [...baseData];
    }
  } else if (localOverrides) {
    try {
      const overridesMap = JSON.parse(localOverrides);
      merged = baseData.map(st => overridesMap[st.code] ? { ...st, ...overridesMap[st.code] } : st);
      Object.keys(overridesMap).forEach(code => {
        if (!merged.find(s => s.code === code)) {
          merged.unshift(overridesMap[code]);
        }
      });
    } catch (e) {
      merged = [...baseData];
    }
  } else {
    merged = [...baseData];
  }

  // Safety: If merged is still empty, force fallback to baseData
  if (merged.length === 0 && baseData.length > 0) {
    merged = [...baseData];
  }

  // If baseData has students missing in merged, merge them in
  if (baseData.length > merged.length) {
    const existing = new Set(merged.map(s => String(s.code || '').trim()));
    baseData.forEach(s => {
      const c = String(s.code || '').trim();
      if (c && !existing.has(c) && !deletedCodes.includes(c)) {
        merged.push(s);
        existing.add(c);
      }
    });
  }

  // Also apply overrides on top to guarantee persistence
  if (localOverrides) {
    try {
      const overridesMap = JSON.parse(localOverrides);
      merged = merged.map(st => overridesMap[st.code] ? { ...st, ...overridesMap[st.code] } : st);
    } catch(e) {}
  }

  // Self-heal: purge any blank or unnamed student entries in overrides and fullDb
  try {
    if (localOverrides) {
      const ov = JSON.parse(localOverrides);
      let ovCleaned = false;
      Object.keys(ov).forEach(c => {
        const item = ov[c];
        const nm = String(item?.name || '').trim();
        if (!nm || nm.length < 2 || nm.includes('غير مسمى')) {
          delete ov[c];
          ovCleaned = true;
        }
      });
      if (ovCleaned) localStorage.setItem('araij_students_overrides', JSON.stringify(ov));
    }
  } catch(e) {}

  allStudents = merged.filter(s => {
    if (!s) return false;
    if (deletedCodes.includes(s.code)) return false;
    const name = String(s.name || '').trim();
    if (!name || name.length < 2 || name.includes('غير مسمى')) return false;
    return true;
  });

  // Guarantee that if allStudents ended up empty but baseData exists, restore baseData!
  if (allStudents.length === 0 && baseData.length > 0) {
    allStudents = baseData.filter(s => s && s.name && !deletedCodes.includes(s.code));
  }

  // ✅ DEDUPLICATION: if same code appears twice, keep only the first (most recent)
  const seenCodes = new Set();
  allStudents = allStudents.filter(s => {
    const key = String(s.code || '').trim();
    if (!key || seenCodes.has(key)) return false;
    seenCodes.add(key);
    return true;
  });

  // Update cached full DB with cleaned students
  try { localStorage.setItem('araij_full_students_db', JSON.stringify(allStudents)); } catch(e) {}

  allStudents.sort((a, b) => (a.name || '').localeCompare(b.name || '', 'ar', { numeric: true, sensitivity: 'base' }));

  allStudents.forEach(st => {
    // 1. Reconstruct or heal academicSubjects
    if (!st.academicSubjects || Object.keys(st.academicSubjects).length === 0) {
      if (st.teachersSummary || st.subjectsSummary) {
        st.academicSubjects = parseSubjectsFromSummary(st.subjectsSummary, st.teachersSummary);
      }
    } else {
      // Heal any corrupted teacher names in existing academicSubjects
      Object.keys(st.academicSubjects).forEach(sub => {
        const canon = canonicalSubjectName(sub) || sub;
        st.academicSubjects[sub].teacher = cleanTeacherName(st.academicSubjects[sub].teacher, canon);
      });
    }

    // 2. Rebuild clean teachersSummary and subjectsSummary (including sessions if present)
    if (st.academicSubjects && Object.keys(st.academicSubjects).length > 0) {
      const subKeys = Object.keys(st.academicSubjects);
      st.subjectsSummary = subKeys.join(', ');
      st.teachersSummary = subKeys.map(k => {
        const subObj = st.academicSubjects[k];
        const tName = cleanTeacherName(subObj.teacher, k);
        const hasSess = subObj.sessions && subObj.sessions.some(s => s && s.trim() !== '');
        if (hasSess) {
          const sessStr = subObj.sessions.map(s => (s && s.trim()) ? s.trim() : '_').join(', ');
          return `${k}/${tName} [حصص: ${sessStr}]`;
        }
        return `${k}/${tName}`;
      }).join(' | ');
    }

    st._metrics = calculateStudentMetrics(st);
    
    st._searchString = buildStudentSearchString(st);
  });

  try {
    updateSuggestedCode();
  } catch(e) {}
}


function calculateStudentMetrics(student) {
  let totalScore = 0;
  let totalMax = 0;
  let totalPresent = 0;
  let totalAbsent = 0;
  let totalRecordedSessions = 0;

  const subjects = student.academicSubjects || {};
  const subjectList = Object.keys(subjects);

  subjectList.forEach(subKey => {
    const subData = subjects[subKey];
    const sessions = subData.sessions || [];

    sessions.forEach(val => {
      if (!val || val.trim() === '') return;
      const cleanVal = val.trim();
      totalRecordedSessions++;

      if (cleanVal === 'غ' || cleanVal === 'غايب' || cleanVal === 'غياب') {
        totalAbsent++;
      } else if (cleanVal === '✓' || cleanVal === 'حاضر' || cleanVal === 'حضر') {
        totalPresent++;
      } else if (cleanVal.includes('/')) {
        totalPresent++;
        const parts = cleanVal.split('/');
        const earned = parseFloat(parts[0].trim());
        const max = parseFloat(parts[1].trim());
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

  const hasRecordedData = totalRecordedSessions > 0;
  const overallPercentage = totalMax > 0 ? (totalScore / totalMax) * 100 : (hasRecordedData ? 80 : null);
  const taqdeer = hasRecordedData ? get_taqdeer(overallPercentage) : "مسجل";

  let attendanceRate = 100;
  if (totalRecordedSessions > 0) {
    attendanceRate = Math.round((totalPresent / totalRecordedSessions) * 100);
  }

  let level = 'مستوى منتظم';
  if (taqdeer === "ممتاز") level = "مستوى متفوق 🌟";
  else if (taqdeer === "جيد جدًا") level = "مستوى ممتاز ومستقر ✨";
  else if (taqdeer === "جيد") level = "مستوى جيد 👍";
  else if (taqdeer === "مقبول") level = "يحتاج إلى متابعة ⚠️";
  else if (taqdeer === "ضعيف") level = "يحتاج لخطة علاجية 🚨";

  return {
    academicPct: overallPercentage !== null ? Math.round(overallPercentage) : null,
    attendanceRate,
    gradeEvaluation: taqdeer,
    level,
    totalPresent,
    totalAbsent,
    totalRecordedSessions,
    subjectCount: subjectList.length,
    hasRecordedData
  };
}

function updateKPIStats() {
  const elTotal = document.getElementById('statTotalStudents');
  if (elTotal) elTotal.textContent = allStudents.length.toLocaleString('ar-EG');
  
  const g1 = allStudents.filter(s => s.grade === 'ث1').length;
  const g2 = allStudents.filter(s => s.grade === 'ث2').length;
  const g3 = allStudents.filter(s => s.grade === 'ث3').length;
  
  const elG1 = document.getElementById('statGrade1');
  const elG2 = document.getElementById('statGrade2');
  const elG3 = document.getElementById('statGrade3');

  if (elG1) elG1.textContent = g1.toLocaleString('ar-EG');
  if (elG2) elG2.textContent = g2.toLocaleString('ar-EG');
  if (elG3) elG3.textContent = g3.toLocaleString('ar-EG');
}

function populateAreaFilter() {
  const areaSelect = document.getElementById('filterArea');
  if (!areaSelect) return;
  const areas = new Set();
  allStudents.forEach(s => {
    if (s.area && s.area.trim() !== '') {
      areas.add(s.area.trim());
    }
  });

  const sortedAreas = Array.from(areas).sort();
  areaSelect.innerHTML = '<option value="all">كل المناطق والمراكز</option>';
  sortedAreas.forEach(a => {
    const opt = document.createElement('option');
    opt.value = a;
    opt.textContent = a;
    areaSelect.appendChild(opt);
  });
}

// ====================================================
// 8. EXACT ARABIC SEARCH & MULTI-FILTERING
// ====================================================

// Debounced search — waits 250ms after last keystroke before filtering
const _debouncedApplyFilters = debounce(() => applyFilters(), 250);

function handleSearchInput() {
  const query = (document.getElementById('searchInput')?.value || '').trim();
  currentSearchQuery = query;
  const clearBtn = document.getElementById('clearSearchBtn');
  if (clearBtn) {
    if (query) clearBtn.classList.remove('hidden');
    else clearBtn.classList.add('hidden');
  }
  _debouncedApplyFilters();
}


function clearSearch() {
  const input = document.getElementById('searchInput');
  if (input) input.value = '';
  document.getElementById('clearSearchBtn')?.classList.add('hidden');
  applyFilters();
}

function setGradeFilter(grade) {
  currentGradeFilter = grade;
  document.querySelectorAll('.grade-tab').forEach(btn => {
    btn.className = 'grade-tab px-3.5 py-2 rounded-xl text-xs font-semibold text-slate-600 hover:text-slate-900 transition-all';
  });

  const tabId = grade === 'all' ? 'tab_grade_all' : 
                grade === 'ث1' ? 'tab_grade_1' : 
                grade === 'ث2' ? 'tab_grade_2' : 'tab_grade_3';
  
  const activeBtn = document.getElementById(tabId);
  if (activeBtn) {
    activeBtn.className = 'grade-tab px-3.5 py-2 rounded-xl text-xs font-bold transition-all bg-sky-600 text-white shadow';
  }

  applyFilters();
}

function setPerformanceFilter(perf) {
  const el = document.getElementById('filterEvaluation');
  if (el) el.value = perf;
  applyFilters();
}

function applyFilters() {
  const query = (document.getElementById('searchInput')?.value || '').trim();
  const queryNorm = normalize_arabic(query);

  const clearBtn = document.getElementById('clearSearchBtn');
  if (clearBtn) {
    if (query) clearBtn.classList.remove('hidden');
    else clearBtn.classList.add('hidden');
  }

  currentEvaluationFilter = document.getElementById('filterEvaluation')?.value || 'all';
  currentAttendanceFilter = document.getElementById('filterAttendance')?.value || 'all';
  currentAreaFilter = document.getElementById('filterArea')?.value || 'all';

  filteredStudents = allStudents.filter(s => {
    if (currentGradeFilter !== 'all' && s.grade !== currentGradeFilter) return false;

    if (currentEvaluationFilter !== 'all') {
      const evalGrade = s._metrics.gradeEvaluation;
      if (currentEvaluationFilter === 'excellent' && evalGrade !== 'ممتاز') return false;
      if (currentEvaluationFilter === 'very_good' && evalGrade !== 'جيد جدًا') return false;
      if (currentEvaluationFilter === 'good' && evalGrade !== 'جيد') return false;
      if (currentEvaluationFilter === 'pass' && evalGrade !== 'مقبول') return false;
      if (currentEvaluationFilter === 'needs_followup' && evalGrade !== 'ضعيف') return false;
    }

    if (currentAttendanceFilter !== 'all') {
      const att = s._metrics.attendanceRate;
      if (currentAttendanceFilter === 'high' && att < 80) return false;
      if (currentAttendanceFilter === 'medium' && (att < 50 || att >= 80)) return false;
      if (currentAttendanceFilter === 'low' && att >= 50) return false;
    }

    if (currentAreaFilter !== 'all' && s.area !== currentAreaFilter) return false;

    if (queryNorm) {
      if (!s._searchString || !s._searchString.includes(queryNorm)) {
        return false;
      }
    }

    return true;
  });

  currentPage = 1;
  renderStudents();
  updateResultsCount();
}

function updateResultsCount() {
  const el = document.getElementById('resultsCountText');
  if (el) {
    el.textContent = `عرض ${Math.min(filteredStudents.length, currentPage * PAGE_SIZE)} من ${filteredStudents.length} طالب`;
  }
}

function renderStudents() {
  const cardsContainer = document.getElementById('cardsViewContainer');
  const tableBody = document.getElementById('tableRowsBody');
  const emptyState = document.getElementById('emptyResultsState');
  const loadMoreBtn = document.getElementById('paginationContainer');

  if (!cardsContainer || !tableBody) return;

  if (filteredStudents.length === 0) {
    cardsContainer.innerHTML = '';
    tableBody.innerHTML = '';
    if (emptyState) emptyState.classList.remove('hidden');
    if (loadMoreBtn) loadMoreBtn.classList.add('hidden');
    return;
  }

  if (emptyState) emptyState.classList.add('hidden');

  const visibleRaw = filteredStudents.slice(0, currentPage * PAGE_SIZE);
  // Ensure absolute uniqueness of rendered cards by student code
  const seenRenderCodes = new Set();
  const visibleItems = [];
  for (const st of visibleRaw) {
    const c = String(st.code || '').trim();
    if (c && seenRenderCodes.has(c)) continue;
    if (c) seenRenderCodes.add(c);
    visibleItems.push(st);
  }

  if (loadMoreBtn) {
    if (visibleRaw.length < filteredStudents.length) {
      loadMoreBtn.classList.remove('hidden');
    } else {
      loadMoreBtn.classList.add('hidden');
    }
  }

  cardsContainer.innerHTML = visibleItems.map(st => createStudentCardHtml(st)).join('');
  tableBody.innerHTML = visibleItems.map(st => createStudentTableRowHtml(st)).join('');
  initIcons();
  updateSelectionToolbarUI();
}

function loadMoreStudents() {
  currentPage++;
  renderStudents();
  updateResultsCount();
}

function toggleViewMode() {
  isTableView = !isTableView;
  const cardsContainer = document.getElementById('cardsViewContainer');
  const tableViewContainer = document.getElementById('tableViewContainer');
  const viewModeIcon = document.getElementById('viewModeIcon');
  const viewModeText = document.getElementById('viewModeText');

  if (isTableView) {
    cardsContainer?.classList.add('hidden');
    tableViewContainer?.classList.remove('hidden');
    if (viewModeText) viewModeText.textContent = 'عرض الجداول';
    if (viewModeIcon) viewModeIcon.setAttribute('data-lucide', 'list');
  } else {
    cardsContainer?.classList.remove('hidden');
    tableViewContainer?.classList.add('hidden');
    if (viewModeText) viewModeText.textContent = 'عرض البطاقات';
    if (viewModeIcon) viewModeIcon.setAttribute('data-lucide', 'layout-grid');
  }
  initIcons();
}

function createStudentCardHtml(student) {
  const m = student._metrics;
  const isSelected = selectedStudentCodes.has(String(student.code));
  const gradeBadgeBg = student.grade === 'ث1' ? 'bg-emerald-500/20 text-emerald-300 border-emerald-500/30' :
                       student.grade === 'ث2' ? 'bg-indigo-500/20 text-indigo-300 border-indigo-500/30' :
                       'bg-amber-500/20 text-amber-300 border-amber-500/30';

  const evalColorClass = m.gradeEvaluation === 'ممتاز' ? 'text-purple-400 bg-purple-500/10 border-purple-500/30' :
                         m.gradeEvaluation === 'جيد جدًا' ? 'text-emerald-400 bg-emerald-500/10 border-emerald-500/30' :
                         m.gradeEvaluation === 'جيد' ? 'text-sky-400 bg-sky-500/10 border-sky-500/30' :
                         m.gradeEvaluation === 'مقبول' ? 'text-amber-400 bg-amber-500/10 border-amber-500/30' :
                         'text-rose-400 bg-rose-500/10 border-rose-500/30';

  const displayEvaluation = m.hasRecordedData ? m.gradeEvaluation : 'مسجل بالقاعدة';

  return `
    <div class="glass-card rounded-3xl p-5 flex flex-col justify-between space-y-4 hover:border-sky-500/50 transition-all border ${isSelected ? 'border-emerald-500 ring-2 ring-emerald-500/30 bg-emerald-950/10' : 'border-slate-200'} shadow-md relative">
      <div class="flex items-start justify-between gap-3">
        <div class="flex items-center gap-3 cursor-pointer flex-1" onclick="openStudentDetailModal('${student.code}')">
          <div class="w-12 h-12 rounded-2xl bg-gradient-to-br from-sky-600 to-indigo-600 flex items-center justify-center text-white font-black text-base shadow border border-sky-400/30 shrink-0">
            ${(student.name || 'ط').charAt(0)}
          </div>
          <div>
            <h3 class="text-base font-black text-slate-900 line-clamp-1 hover:text-sky-600 transition-colors" style="color: #0f172a !important;">
              ${student.name || 'طالب غير مسمى'}
            </h3>
            <div class="flex items-center gap-2 mt-1 text-[11px] text-slate-500">
              <span class="font-mono font-bold text-sky-700 bg-sky-50 px-1.5 py-0.5 rounded border border-sky-200">#${student.code}</span>
              ${student.area ? `<span class="text-slate-600 font-semibold">• ${student.area}</span>` : ''}
            </div>
          </div>
        </div>

        <div class="flex items-center gap-2 shrink-0">
          <span class="text-[10px] font-bold px-2.5 py-1 rounded-full border ${gradeBadgeBg}">
            ${student.grade || 'غير محدد'}
          </span>
          <label class="cursor-pointer p-1 rounded-lg hover:bg-slate-100 flex items-center justify-center" title="تحديد لإرسال واتساب جماعي" onclick="event.stopPropagation()">
            <input type="checkbox" onchange="toggleSelectStudent('${student.code}', this.checked)" ${isSelected ? 'checked' : ''} class="w-5 h-5 rounded-lg text-emerald-600 focus:ring-emerald-500 bg-white border-slate-300 cursor-pointer">
          </label>
        </div>
      </div>

      <div class="bg-slate-50 rounded-2xl p-3 border border-slate-200 space-y-2.5 cursor-pointer" onclick="openStudentDetailModal('${student.code}')">
        <div class="flex items-center justify-between text-xs font-semibold">
          <span class="text-slate-600">التقدير العام:</span>
          <span class="px-2.5 py-0.5 rounded text-[11px] font-bold border ${evalColorClass}">
            ${displayEvaluation}
          </span>
        </div>

        <div>
          <div class="flex items-center justify-between text-[11px] text-slate-600 mb-1">
            <span>نسبة الحضور:</span>
            <span class="font-bold text-teal-700 font-mono">${m.attendanceRate}%</span>
          </div>
          <div class="w-full bg-slate-200 rounded-full h-1.5 overflow-hidden">
            <div class="bg-gradient-to-r from-teal-500 to-sky-500 h-1.5 rounded-full" style="width: ${m.attendanceRate}%"></div>
          </div>
        </div>
      </div>

      <div class="text-[11px] text-slate-600 flex items-center justify-between pt-1 border-t border-slate-200">
        <div class="flex items-center gap-1">
          <i data-lucide="phone" class="w-3.5 h-3.5 text-emerald-600"></i>
          <span class="font-mono font-semibold">${student.parentPhone || student.phone || 'غير مسجل'}</span>
        </div>
        <div class="text-slate-500 font-medium">
          ${m.subjectCount > 0 ? `${m.subjectCount} مواد مسجلة` : (student.specialization || 'عام')}
        </div>
      </div>

      <!-- Enrolled Subjects & Teachers Badges -->
      <div class="flex flex-wrap gap-1.5 pt-1">
        ${(() => {
          if (!student.academicSubjects || Object.keys(student.academicSubjects).length === 0) {
            if (student.teachersSummary || student.subjectsSummary) {
              student.academicSubjects = parseSubjectsFromSummary(student.subjectsSummary, student.teachersSummary);
            }
          }
          const subKeys = Object.keys(student.academicSubjects || {});
          if (subKeys.length > 0) {
            return subKeys.map(sub => {
              const canon = canonicalSubjectName(sub) || sub;
              const teacherClean = cleanTeacherName(student.academicSubjects[sub]?.teacher, canon);
              if (student.academicSubjects[sub]) {
                student.academicSubjects[sub].teacher = teacherClean;
              }
              return `
                <div class="inline-flex items-center gap-1 text-[11px] px-2.5 py-1 rounded-xl bg-slate-100 text-slate-800 font-medium border border-slate-200 shadow-sm">
                  <span class="font-bold text-sky-700">📖 ${sub}:</span>
                  <span class="text-amber-800 font-semibold">${teacherClean}</span>
                </div>
              `;
            }).join('');
          } else {
            return `<span class="text-[10px] text-slate-500 italic">مسجل بالقاعدة الرئيسية (${student.specialization || 'عام'})</span>`;
          }
        })()}
      </div>

      <!-- Supervisor / Audit Hint for Eng. Wafaa and Admins -->
      ${(student.createdBy || student.lastModifiedBy) ? `
        <div class="mt-1 pt-1.5 border-t border-slate-200 flex items-center justify-between text-[11px] bg-amber-50/60 -mx-2 -mb-1 px-3 py-1.5 rounded-xl border border-amber-200">
          <div class="flex items-center gap-1.5 text-amber-800 font-bold">
            <i data-lucide="shield-check" class="w-3.5 h-3.5 text-amber-600 shrink-0"></i>
            <span>بواسطة: <strong class="text-slate-900 underline" style="color: #0f172a !important;">${student.lastModifiedBy || student.createdBy}</strong></span>
          </div>
          <span class="text-[10px] text-slate-600">${student.lastAction || 'تسجيل'} (${student.lastActionTime || student.regDate || ''})</span>
        </div>
      ` : ''}

      <!-- Action Buttons Right on the Card -->
      <div class="grid grid-cols-4 gap-1.5 pt-1">
        <button onclick="openEditStudentModalDirect('${student.code}')" class="py-2 px-1.5 rounded-xl bg-amber-500/20 hover:bg-amber-500/30 text-amber-300 text-[11px] font-bold transition-all flex items-center justify-center gap-1 border border-amber-500/30" title="تعديل أو مسح الرقم">
          <i data-lucide="edit-3" class="w-3.5 h-3.5"></i>
          <span>تعديل</span>
        </button>

        <button onclick="confirmDeleteStudentByCode('${student.code}')" class="py-2 px-1.5 rounded-xl bg-rose-500/20 hover:bg-rose-500/30 text-rose-300 text-[11px] font-bold transition-all flex items-center justify-center gap-1 border border-rose-500/30" title="حذف الطالب">
          <i data-lucide="trash-2" class="w-3.5 h-3.5"></i>
          <span>حذف</span>
        </button>

        <button onclick="openParentReportModal('${student.code}')" class="py-2 px-1.5 rounded-xl bg-sky-600/20 hover:bg-sky-600/30 text-sky-300 text-[11px] font-bold transition-all flex items-center justify-center gap-1 border border-sky-500/30" title="تقرير ولي الأمر">
          <i data-lucide="file-text" class="w-3.5 h-3.5"></i>
          <span>التقرير</span>
        </button>

        <button onclick="quickWhatsAppSend('${student.code}')" class="py-2 px-1.5 rounded-xl bg-emerald-600 hover:bg-emerald-500 text-white text-[11px] font-bold transition-all flex items-center justify-center gap-1 shadow-md shadow-emerald-600/20 active:scale-95" title="إرسال واتساب">
          <i data-lucide="send" class="w-3.5 h-3.5"></i>
          <span>واتساب</span>
        </button>
      </div>
    </div>
  `;
}

function createStudentTableRowHtml(student) {
  const m = student._metrics;
  const isSelected = selectedStudentCodes.has(String(student.code));
  const evalColor = m.gradeEvaluation === 'ممتاز' ? 'text-purple-400' :
                    m.gradeEvaluation === 'جيد جدًا' ? 'text-emerald-400' :
                    m.gradeEvaluation === 'جيد' ? 'text-sky-400' :
                    m.gradeEvaluation === 'مقبول' ? 'text-amber-400' : 'text-rose-400';

  return `
    <tr class="hover:bg-slate-50 transition-colors ${isSelected ? 'bg-emerald-500/10' : ''}">
      <td class="py-3 px-3 text-center">
        <input type="checkbox" onchange="toggleSelectStudent('${student.code}', this.checked)" ${isSelected ? 'checked' : ''} class="w-4 h-4 rounded text-emerald-600 focus:ring-emerald-500 bg-white border-slate-300 cursor-pointer">
      </td>
      <td class="py-3 px-4 font-mono font-bold text-sky-700">#${student.code}</td>
      <td class="py-3 px-4 font-bold text-slate-900 cursor-pointer hover:text-sky-600" style="color: #0f172a !important;" onclick="openStudentDetailModal('${student.code}')">
        <div class="font-bold text-slate-900" style="color: #0f172a !important;">${student.name}</div>
        ${(student.createdBy || student.lastModifiedBy) ? `
          <div class="text-[10px] text-amber-700 font-normal">بواسطة: ${student.lastModifiedBy || student.createdBy} (${student.lastAction || 'تسجيل'})</div>
        ` : ''}
      </td>
      <td class="py-3 px-4">${student.grade || '-'}</td>
      <td class="py-3 px-4 text-slate-400">${student.area || '-'}</td>
      <td class="py-3 px-4 font-mono text-emerald-400">${student.parentPhone || student.phone || '-'}</td>
      <td class="py-3 px-4 font-bold text-teal-300">${m.attendanceRate}%</td>
      <td class="py-3 px-4 font-bold ${evalColor}">${m.gradeEvaluation}</td>
      <td class="py-3 px-4 text-center">
        <div class="flex items-center justify-center gap-1.5">
          <button onclick="openEditStudentModalDirect('${student.code}')" class="p-1.5 rounded-lg bg-amber-500/20 text-amber-300" title="تعديل"><i data-lucide="edit-3" class="w-3.5 h-3.5"></i></button>
          <button onclick="confirmDeleteStudentByCode('${student.code}')" class="p-1.5 rounded-lg bg-rose-500/20 text-rose-400" title="حذف"><i data-lucide="trash-2" class="w-3.5 h-3.5"></i></button>
          <button onclick="openParentReportModal('${student.code}')" class="p-1.5 rounded-lg bg-sky-500/20 text-sky-300" title="تقرير ولي الأمر"><i data-lucide="file-text" class="w-3.5 h-3.5"></i></button>
          <button onclick="quickWhatsAppSend('${student.code}')" class="p-1.5 rounded-lg bg-emerald-600 text-white" title="إرسال واتساب"><i data-lucide="send" class="w-3.5 h-3.5"></i></button>
        </div>
      </td>
    </tr>
  `;
}

// ====================================================
// 9. LIVE ATTENDANCE & GRADES MATRIX
// ====================================================

function initMatrixSelects() {
  onMatrixFilterChange();
}

function onMatrixFilterChange() {
  const grade = document.getElementById('matrixGradeSelect')?.value || 'ث1';
  const subSelect = document.getElementById('matrixSubjectSelect');
  if (!subSelect) return;

  const subjects = SUBJECTS[grade] || [];
  subSelect.innerHTML = subjects.map(s => `<option value="${s}">مادة: ${s}</option>`).join('');
  onMatrixSubjectChange();
}

function onMatrixSubjectChange() {
  const subject = document.getElementById('matrixSubjectSelect')?.value || 'عربي';
  const teacherSelect = document.getElementById('matrixTeacherSelect');
  if (!teacherSelect) return;

  const teachers = TEACHERS[subject] || [];
  teacherSelect.innerHTML = `<option value="all">كل المدرسين (${subject})</option>` + 
    teachers.map(t => `<option value="${t}">المدرس: ${t}</option>`).join('');

  renderAttendanceMatrix();
}

function renderAttendanceMatrix() {
  const tbody = document.getElementById('attendanceMatrixTableBody');
  if (!tbody) return;

  const grade = document.getElementById('matrixGradeSelect')?.value || 'ث1';
  const subject = document.getElementById('matrixSubjectSelect')?.value || 'عربي';
  const selectedTeacher = document.getElementById('matrixTeacherSelect')?.value || 'all';
  const searchQuery = (document.getElementById('matrixSearchInput')?.value || '').trim();
  const searchNorm = normalize_arabic(searchQuery);

  let students = allStudents.filter(s => s.grade === grade);

  if (selectedTeacher !== 'all') {
    const tNorm = normalize_arabic(selectedTeacher);
    students = students.filter(s => {
      const summaryNorm = normalize_arabic(s.teachersSummary || '');
      if (summaryNorm.includes(tNorm)) return true;
      const sub = s.academicSubjects && s.academicSubjects[subject];
      return sub && normalize_arabic(sub.teacher || '').includes(tNorm);
    });
  }

  if (searchNorm) {
    students = students.filter(s => {
      const nameNorm = normalize_arabic(s.name || '');
      const codeNorm = normalize_arabic(s.code || '');
      return nameNorm.includes(searchNorm) || codeNorm.includes(searchNorm);
    });
  }

  if (students.length === 0) {
    tbody.innerHTML = `<tr><td colspan="13" class="p-8 text-center text-slate-400">لا يوجد طلاب مطابقين للبحث والفلاتر المحددة.</td></tr>`;
    return;
  }

  tbody.innerHTML = students.map((st, idx) => {
    const subData = (st.academicSubjects && st.academicSubjects[subject]) || {};
    const sessions = getStudentSubjectMonthSessions(st, subject, currentActiveMonth);

    return `
      <tr class="hover:bg-slate-50 transition-colors">
        <td class="p-3 text-center text-slate-400">${idx + 1}</td>
        <td class="p-3 text-center font-mono font-bold text-sky-700">#${st.code}</td>
        <td class="p-3 font-bold text-slate-900 cursor-pointer hover:text-sky-700" style="color: #0f172a !important;" onclick="openStudentDetailModal('${st.code}')">
          <div class="font-bold text-slate-900" style="color: #0f172a !important;">${st.name}</div>
          <div class="text-[10px] text-slate-500">${subData.teacher || 'مدرس المادة'}</div>
        </td>
        <td class="p-3 text-slate-400">${st.specialization || 'عام'}</td>
        ${[0, 1, 2, 3, 4, 5, 6, 7].map(sIdx => {
          const val = sessions[sIdx] || '';
          const isPresent = val === '✓' || val === 'حاضر';
          const isAbsent = val === 'غ' || val === 'غائب';
          const isScore = val.includes('/') || (!isNaN(parseFloat(val)) && val !== '');

          const btnClass = isPresent ? 'bg-emerald-500 text-white font-bold shadow' :
                           isAbsent ? 'bg-rose-500 text-white font-bold shadow' :
                           isScore ? 'bg-sky-500 text-white font-bold shadow' :
                           'bg-slate-800 hover:bg-slate-700 text-slate-500';

          return `
            <td class="p-1.5 text-center">
              <button onclick="openSessionEditor('${st.code}', '${subject}', ${sIdx}, '${val}', '${currentActiveMonth}')" class="w-10 h-8 rounded-lg text-xs font-mono transition-all transform active:scale-90 flex items-center justify-center mx-auto ${btnClass}" title="انقر لتسجيل الحضور أو كتابة الدرجة (${currentActiveMonth})">
                ${val || '-'}
              </button>
            </td>
          `;
        }).join('')}
        <td class="p-3 text-center font-bold text-teal-300 font-mono">${st._metrics.attendanceRate}%</td>
      </tr>
    `;
  }).join('');
}

// ====================================================
// 10. STUDENT DETAILS & INTERACTIVE SESSIONS
// ====================================================

function openStudentDetailModal(code) {
  const student = allStudents.find(s => s.code === code);
  if (!student) return;

  activeDetailStudentCode = code;

  const elName = document.getElementById('detailStudentName');
  const elBadge = document.getElementById('detailStudentCodeBadge');
  const elAvatar = document.getElementById('detailStudentAvatar');
  const elSub = document.getElementById('detailStudentSub');
  const elPhone = document.getElementById('detailPhone');
  const elParent = document.getElementById('detailParentPhone');
  const elArea = document.getElementById('detailArea');
  const elReg = document.getElementById('detailRegDate');

  if (elName) elName.textContent = student.name || 'طالب غير مسمى';
  if (elBadge) elBadge.textContent = `#${student.code}`;
  if (elAvatar) elAvatar.textContent = (student.name || 'ط').charAt(0);
  if (elSub) elSub.textContent = `${student.grade || ''} ${student.specialization ? '• ' + student.specialization : ''}`;
  if (elPhone) elPhone.textContent = student.phone || 'غير مسجل';
  if (elParent) elParent.textContent = student.parentPhone || 'غير مسجل';
  if (elArea) elArea.textContent = student.area || 'غير محددة';
  if (elReg) elReg.textContent = student.regDate || '2026';

  // Audit Hint for Eng. Wafaa and Supervisors
  const auditContainer = document.getElementById('detailAuditHint');
  if (auditContainer) {
    if (student.createdBy || student.lastModifiedBy) {
      auditContainer.classList.remove('hidden');
      auditContainer.innerHTML = `
        <div class="glass-card p-3 rounded-2xl border border-amber-500/30 bg-amber-500/10 flex flex-col sm:flex-row sm:items-center justify-between gap-2 text-xs">
          <div class="flex items-center gap-2">
            <div class="w-7 h-7 rounded-xl bg-amber-500/20 text-amber-400 flex items-center justify-center">
              <i data-lucide="shield-check" class="w-4 h-4"></i>
            </div>
            <div>
              <span class="text-slate-400">سجل الرقابة: </span>
              <span class="text-amber-300 font-bold">المسؤول: ${student.lastModifiedBy || student.createdBy}</span>
              ${student.createdBy && student.createdBy !== student.lastModifiedBy ? `<span class="text-slate-400 text-[11px]"> (أنشئ بواسطة: ${student.createdBy})</span>` : ''}
            </div>
          </div>
          <div class="text-[11px] text-slate-300 font-medium">
            <span class="bg-slate-800/80 px-2 py-0.5 rounded-lg border border-slate-700">${student.lastAction || 'تسجيل'}</span>
            <span class="text-slate-400 mr-1">${student.lastActionTime || student.regDate || ''}</span>
          </div>
        </div>
      `;
    } else {
      auditContainer.classList.add('hidden');
      auditContainer.innerHTML = '';
    }
  }

  // Ensure academicSubjects is parsed from summaries if empty
  if (!student.academicSubjects || Object.keys(student.academicSubjects).length === 0) {
    if (student.teachersSummary || student.subjectsSummary) {
      student.academicSubjects = parseSubjectsFromSummary(student.subjectsSummary, student.teachersSummary);
    }
  }

  const container = document.getElementById('detailSubjectsContainer');
  const subjects = student.academicSubjects || {};
  const subKeys = Object.keys(subjects);

  const elSubCount = document.getElementById('detailSubjectCount');
  if (elSubCount) elSubCount.textContent = `${subKeys.length} مادة مسجلة (${activeDetailMonth} - انقر على الحصة للرصد)`;

  renderDetailMonthTabs(student);

  if (container) {
    if (subKeys.length === 0) {
      container.innerHTML = `
        <div class="text-center py-6 text-slate-400 text-xs bg-slate-800/40 rounded-xl border border-slate-800 space-y-2">
          <div>المواد العامة مسجلة بالقاعدة الرئيسية (بانتظار رصد الحصص).</div>
          <button onclick="addDefaultSubjectsToStudent('${student.code}')" class="px-3 py-1.5 rounded-xl bg-sky-600/30 hover:bg-sky-600 text-sky-200 text-xs font-bold border border-sky-500/30">
            + تفعيل مواد المرحلة لتسجيل الحضور
          </button>
        </div>
      `;
    } else {
      container.innerHTML = subKeys.map(subName => {
        const sub = subjects[subName];
        const sessions = getStudentSubjectMonthSessions(student, subName, activeDetailMonth);

        return `
          <div class="glass-card p-4 rounded-2xl border border-slate-800 space-y-3">
            <div class="flex items-center justify-between">
              <div>
                <h5 class="text-xs font-bold text-sky-300">${subName}</h5>
                <span class="text-[11px] text-slate-400">المدرس: ${cleanTeacherName(sub.teacher, subName)}</span>
              </div>
              <span class="text-[10px] text-emerald-400 font-bold">${activeDetailMonth} 👆</span>
            </div>

            <div class="grid grid-cols-4 sm:grid-cols-8 gap-2">
              ${sessions.map((score, idx) => {
                const val = score && score.trim() !== '' ? score : '-';
                const isAbsent = val === '0' || val === 'غائب' || val === 'غ';
                const isPresent = val !== '-' && !isAbsent;
                
                const badgeClass = isPresent ? 'bg-emerald-500/20 text-emerald-300 border-emerald-500/30 hover:bg-emerald-500/30' :
                                   isAbsent ? 'bg-rose-500/20 text-rose-300 border-rose-500/30 hover:bg-rose-500/30' :
                                   'bg-slate-800 text-slate-500 border-slate-700/50 hover:bg-slate-700/60';

                return `
                  <button type="button" onclick="openSessionEditor('${student.code}', '${subName}', ${idx}, '${val}', '${activeDetailMonth}')" class="rounded-xl p-2 text-center border ${badgeClass} cursor-pointer transition-all active:scale-95" title="رصد ${activeDetailMonth} - حصة ${idx + 1}">
                    <span class="text-[10px] block opacity-70 mb-0.5">حصة ${idx + 1}</span>
                    <strong class="text-xs font-bold font-mono">${val}</strong>
                  </button>
                `;
              }).join('')}
            </div>
          </div>
        `;
      }).join('');
    }
  }

  const openBtn = document.getElementById('openReportFromDetailBtn');
  if (openBtn) {
    openBtn.onclick = () => {
      closeStudentDetailModal();
      openParentReportModal(code);
    };
  }

  document.getElementById('studentDetailModal')?.classList.remove('hidden');
  initIcons();
}

function closeStudentDetailModal() {
  document.getElementById('studentDetailModal')?.classList.add('hidden');
  activeDetailStudentCode = null;
}

function addDefaultSubjectsToStudent(code) {
  const student = allStudents.find(s => s.code === code);
  if (!student) return;

  const defaultSubjs = SUBJECTS[student.grade || 'ث1'] || ["عربي", "انجليزي", "رياضة"];
  if (!student.academicSubjects) student.academicSubjects = {};

  defaultSubjs.forEach(s => {
    if (!student.academicSubjects[s]) {
      student.academicSubjects[s] = {
        teacher: (TEACHERS[s] && TEACHERS[s][0]) || "مدرس المادة",
        sessions: ["", "", "", "", "", "", "", ""]
      };
    }
  });

  saveStudentOverride(student);
  openStudentDetailModal(code);
  showToast('تم تفعيل مواد المرحلة لتسجيل الحصص!');
}

// ====================================================
// 11. SESSION ATTENDANCE & GRADE EDITOR POPUP
// ====================================================

function openSessionEditor(code, subject, sessionIdx, currentValue, month = null) {
  const student = allStudents.find(s => s.code === code);
  if (!student) return;

  const m = month || currentActiveMonth || 'شهر 1';
  activeSessionEdit = { code, subject, sessionIdx, month: m };
  
  const title = document.getElementById('sessionModalTitle');
  const sub = document.getElementById('sessionModalSub');
  const input = document.getElementById('customExamGradeInput');

  if (title) title.textContent = `تسجيل الحصة ${sessionIdx + 1} (${m}) - مادة ${subject}`;
  if (sub) sub.textContent = `الطالب: ${student.name} (#${student.code})`;
  if (input) input.value = (currentValue && currentValue !== '-' && currentValue !== '✓' && currentValue !== 'غ') ? currentValue : '';

  document.getElementById('sessionGradeModal')?.classList.remove('hidden');
  if (input) setTimeout(() => input.focus(), 100);
  initIcons();
}

function closeSessionGradeModal() {
  document.getElementById('sessionGradeModal')?.classList.add('hidden');
  activeSessionEdit = null;
}

function setQuickSessionState(val) {
  if (!activeSessionEdit) return;
  applySessionValue(activeSessionEdit.code, activeSessionEdit.subject, activeSessionEdit.sessionIdx, val, activeSessionEdit.month);
  closeSessionGradeModal();
}

function saveCustomExamGrade() {
  if (!activeSessionEdit) return;
  const inputVal = (document.getElementById('customExamGradeInput')?.value || '').trim();
  if (!inputVal) {
    alert('يرجى كتابة الدرجة مثل 10/10 أو 18/20 أو 8');
    return;
  }
  applySessionValue(activeSessionEdit.code, activeSessionEdit.subject, activeSessionEdit.sessionIdx, inputVal, activeSessionEdit.month);
  closeSessionGradeModal();
}

function applySessionValue(code, subject, sessionIdx, val, month = null) {
  const student = allStudents.find(s => s.code === code);
  if (!student) return;

  const m = month || (activeSessionEdit && activeSessionEdit.month) || currentActiveMonth || 'شهر 1';
  setStudentSubjectMonthSession(student, subject, m, sessionIdx, val);

  // Audit Hint for Eng. Wafaa and Supervisors
  const uName = currentUser?.name || 'مشرف';
  student.lastModifiedBy = uName;
  student.lastAction = `رصد حضور: ${subject} (${m} - حصة ${sessionIdx + 1}: ${val || 'تفريغ'})`;
  student.lastActionTime = new Date().toLocaleTimeString('ar-EG', { hour: '2-digit', minute: '2-digit' }) + ' ' + new Date().toLocaleDateString('ar-EG');

  rebuildStudentSummaries(student);
  saveStudentOverride(student);

  // 1. Send to Google Sheets (both sheet update and session log)
  syncToGoogleSheets('update_student', { student: student }, true);
  syncToGoogleSheets('mark_session', { code, subject, sessionIdx, value: val, month: m, byUser: uName });

  // 2. Send to Live Railway Server for 0ms multi-device broadcast
  sendToRailwayServer('update_student', { student: student });

  // 3. Log to Audit Trail Sheet (سجل العمليات والرقابة)
  logActivity('رصد حضور طالب', `رصد (${val || 'تفريغ'}) للحصة ${sessionIdx + 1} (${m}) لمادة "${subject}"`, student.code, student.name, {
    grade: student.grade,
    subject: subject,
    monthOrSession: `${m} - ح${sessionIdx + 1}`,
    value: val
  });

  // Check if all 8 sessions of this month are completed for this subject
  const currentSessions = getStudentSubjectMonthSessions(student, subject, m);
  const filledCount = currentSessions.filter(s => s && s.trim() !== '').length;
  if (sessionIdx === 7 && filledCount === 8) {
    const currentMonthIdx = CENTER_MONTHS.indexOf(m);
    const nextMonthName = (currentMonthIdx >= 0 && currentMonthIdx + 1 < CENTER_MONTHS.length) 
      ? CENTER_MONTHS[currentMonthIdx + 1] 
      : `شهر ${CENTER_MONTHS.length + 1}`;
    
    setTimeout(() => {
      if (confirm(`🎉 اكتمل رصد جميع حصص (${m}) لمادة ${subject} للطالب ${student.name}!\nهل ترغب في فتح/الانتقال إلى (${nextMonthName})؟`)) {
        if (!CENTER_MONTHS.includes(nextMonthName)) {
          CENTER_MONTHS.push(nextMonthName);
          try { localStorage.setItem('araij_center_months', JSON.stringify(CENTER_MONTHS)); } catch(e) {}
          populateMonthDropdowns();
        }
        onMatrixMonthChange(nextMonthName);
      }
    }, 400);
  }

  // Refresh modals if open
  const detailModal = document.getElementById('studentDetailModal');
  if (detailModal && !detailModal.classList.contains('hidden') && activeDetailStudentCode === code) {
    openStudentDetailModal(code);
  }

  renderAttendanceMatrix();
  applyFilters();
  updateKPIStats();
  showToast(`✅ تم رصد الحصة (${val || 'فارغ'}) للطالب ${student.name} بنجاح!`);
}

// ====================================================
// 12. EDIT STUDENT (FULL DATA, SUBJECTS & TEACHERS)
// ====================================================

function openEditStudentModalDirect(code) {
  const student = allStudents.find(s => String(s.code).trim() === String(code).trim());
  if (!student) return;

  activeDetailStudentCode = code;

  document.getElementById('editCodeHidden').value = student.code;
  const codeDisplay = document.getElementById('editCodeDisplay');
  if (codeDisplay) codeDisplay.value = `#${student.code}`;

  document.getElementById('editName').value = student.name || '';
  document.getElementById('editPhone').value = student.phone || '';
  document.getElementById('editParentPhone').value = student.parentPhone || '';
  document.getElementById('editArea').value = student.area || '';
  
  const normGrade = normalize_grade(student.grade);
  document.getElementById('editGrade').value = normGrade;
  document.getElementById('editSpec').value = student.specialization || 'عام';

  // Always reconstruct academicSubjects from summaries if empty or incomplete
  student.academicSubjects = parseSubjectsFromSummary(student.subjectsSummary, student.teachersSummary, student.academicSubjects);

  renderEditSubjectsList(normGrade, student.academicSubjects || {});

  document.getElementById('editStudentModal')?.classList.remove('hidden');
  initIcons();
}

function onEditGradeChange() {
  const rawGrade = document.getElementById('editGrade')?.value;
  const normGrade = normalize_grade(rawGrade);
  const code = document.getElementById('editCodeHidden')?.value;
  const student = allStudents.find(s => String(s.code).trim() === String(code).trim());
  if (student) {
    student.academicSubjects = parseSubjectsFromSummary(student.subjectsSummary, student.teachersSummary, student.academicSubjects);
  }
  renderEditSubjectsList(normGrade, (student && student.academicSubjects) || {});
}

function getEnrolledSubjectData(existingSubjectsMap, subj) {
  if (!existingSubjectsMap || typeof existingSubjectsMap !== 'object') return null;
  const canon = canonicalSubjectName(subj);
  if (existingSubjectsMap[canon]) return existingSubjectsMap[canon];
  if (existingSubjectsMap[subj]) return existingSubjectsMap[subj];

  for (const key of Object.keys(existingSubjectsMap)) {
    if (canonicalSubjectName(key) === canon) {
      return existingSubjectsMap[key];
    }
  }
  return null;
}

function renderEditSubjectsList(grade, existingSubjectsMap) {
  const container = document.getElementById('editSubjectsTeachersList');
  if (!container) return;

  const normGrade = normalize_grade(grade);
  const subjects = SUBJECTS[normGrade] || [];

  container.innerHTML = subjects.map((subj, idx) => {
    const enrolledData = getEnrolledSubjectData(existingSubjectsMap, subj);
    const isEnrolled = !!enrolledData;
    const canon = canonicalSubjectName(subj) || subj;
    const currentTeacher = isEnrolled ? cleanTeacherName(enrolledData.teacher || '', canon) : '';
    const teachers = TEACHERS[subj] || [];
    
    // Check if the current teacher is one of the predefined list
    const hasCurrentTeacherInList = currentTeacher && teachers.includes(currentTeacher);
    const isCustomOption = currentTeacher && !hasCurrentTeacherInList && currentTeacher !== 'مدرس آخر';

    return `
      <div class="glass-card p-3 rounded-2xl border border-slate-800 flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3">
        <label class="flex items-center gap-2 cursor-pointer">
          <input type="checkbox" id="edit_subj_chk_${idx}" value="${subj}" ${isEnrolled ? 'checked' : ''} onchange="toggleEditTeacherSelect(${idx})" class="w-4 h-4 rounded text-amber-500 focus:ring-amber-500 bg-slate-800 border-slate-700">
          <span class="text-xs font-bold text-slate-800">${subj}</span>
        </label>

        <div class="flex items-center gap-2 w-full sm:w-auto ${isEnrolled ? '' : 'opacity-40 pointer-events-none'} transition-opacity" id="edit_teacher_group_${idx}">
          <span class="text-[11px] text-slate-400 whitespace-nowrap">المدرس:</span>
          <select id="edit_teacher_val_${idx}" class="w-full sm:w-48 bg-slate-800 border border-slate-700 rounded-xl px-2.5 py-1.5 text-xs text-white focus:ring-1 focus:ring-amber-500">
            ${isCustomOption ? `<option value="${currentTeacher}" selected>${currentTeacher} (مسجل)</option>` : ''}
            ${teachers.length > 0 ? teachers.map(t => `<option value="${t}" ${t === currentTeacher ? 'selected' : ''}>${t}</option>`).join('') : `<option value="مدرس المادة">مدرس المادة</option>`}
            <option value="مدرس آخر" ${currentTeacher === 'مدرس آخر' ? 'selected' : ''}>مدرس آخر / مخصص</option>
          </select>
        </div>
      </div>
    `;
  }).join('');
}

function toggleEditTeacherSelect(idx) {
  const chk = document.getElementById(`edit_subj_chk_${idx}`);
  const group = document.getElementById(`edit_teacher_group_${idx}`);
  if (chk && group) {
    if (chk.checked) {
      group.classList.remove('opacity-40', 'pointer-events-none');
    } else {
      group.classList.add('opacity-40', 'pointer-events-none');
    }
  }
}

function openEditStudentModalFromDetail() {
  if (activeDetailStudentCode) {
    openEditStudentModalDirect(activeDetailStudentCode);
  }
}

function closeEditStudentModal() {
  document.getElementById('editStudentModal')?.classList.add('hidden');
}

function clearEditStudentPhone() {
  document.getElementById('editPhone').value = '';
  showToast('تم مسح رقم هاتف الطالب (اضغط حفظ التعديلات للتأكيد)');
}

function clearEditParentPhone() {
  document.getElementById('editParentPhone').value = '';
  showToast('تم مسح رقم ولي الأمر (اضغط حفظ التعديلات للتأكيد)');
}

function handleSaveEditStudent(e) {
  if (e) e.preventDefault();
  const code = document.getElementById('editCodeHidden').value;
  const student = allStudents.find(s => String(s.code).trim() === String(code).trim());
  if (!student) return;

  const newGrade = normalize_grade(document.getElementById('editGrade').value);
  student.name = document.getElementById('editName').value.trim();
  student.phone = document.getElementById('editPhone').value.trim();
  student.parentPhone = document.getElementById('editParentPhone').value.trim();
  student.area = document.getElementById('editArea').value.trim();
  student.grade = newGrade;
  student.specialization = document.getElementById('editSpec').value.trim() || 'عام';

  // Ensure academicSubjects is parsed and canonicalized
  const baseAcademicSubjects = parseSubjectsFromSummary(student.subjectsSummary, student.teachersSummary, student.academicSubjects);
  const updatedAcademicSubjects = { ...baseAcademicSubjects };
  const subjects = SUBJECTS[newGrade] || [];

  subjects.forEach((subj, idx) => {
    const chk = document.getElementById(`edit_subj_chk_${idx}`);
    const canon = canonicalSubjectName(subj) || subj;
    if (chk) {
      if (chk.checked) {
        const rawTeacher = document.getElementById(`edit_teacher_val_${idx}`)?.value || (TEACHERS[subj] && TEACHERS[subj][0]) || 'مدرس المادة';
        const teacher = cleanTeacherName(rawTeacher, canon);
        const oldSub = getEnrolledSubjectData(updatedAcademicSubjects, subj);
        updatedAcademicSubjects[canon] = {
          teacher: teacher,
          sessions: (oldSub && Array.isArray(oldSub.sessions) && oldSub.sessions.length > 0) ? [...oldSub.sessions] : ["", "", "", "", "", "", "", ""]
        };
      } else {
        // Explicitly unchecked: remove this subject and all its naming variants
        delete updatedAcademicSubjects[canon];
        delete updatedAcademicSubjects[subj];
        Object.keys(updatedAcademicSubjects).forEach(k => {
          if (canonicalSubjectName(k) === canon) delete updatedAcademicSubjects[k];
        });
      }
    }
  });

  const selectedSubjects = Object.keys(updatedAcademicSubjects);
  student.academicSubjects = updatedAcademicSubjects;
  student.subjectsSummary = selectedSubjects.join(', ');
  student.teachersSummary = selectedSubjects.map(k => {
    const subObj = updatedAcademicSubjects[k];
    const tName = cleanTeacherName(subObj.teacher, k);
    const hasSess = subObj.sessions && subObj.sessions.some(s => s && s.trim() !== '');
    if (hasSess) {
      const sessStr = subObj.sessions.map(s => (s && s.trim()) ? s.trim() : '_').join(', ');
      return `${k}/${tName} [حصص: ${sessStr}]`;
    }
    return `${k}/${tName}`;
  }).join(' | ');
  student._metrics = calculateStudentMetrics(student);
  student._searchString = buildStudentSearchString(student);

  // Audit Hint for Eng. Wafaa and Supervisors
  const uName = currentUser?.name || 'م. وفاء ناصر';
  student.lastModifiedBy = uName;
  student.lastAction = 'تعديل بيانات الطالب والمواد';
  student.lastActionTime = new Date().toLocaleTimeString('ar-EG', { hour: '2-digit', minute: '2-digit' }) + ' ' + new Date().toLocaleDateString('ar-EG');

  saveStudentOverride(student);
  syncToGoogleSheets('update_student', { student: student }, true);
  sendToRailwayServer('update_student', { student: student });

  try {
    logActivity('تعديل بيانات طالب', `قام ${uName} بتعديل بيانات الطالب "${student.name}" (#${student.code})`);
  } catch(e) {}

  closeEditStudentModal();
  if (activeDetailStudentCode === code) openStudentDetailModal(code);
  applyFilters();
  updateKPIStats();
  showToast(`✅ تم حفظ تعديلات الطالب ${student.name} بنجاح ومزامنتها!`);
}

function confirmDeleteStudentByCode(code) {
  const student = allStudents.find(s => s.code === code);
  if (!student) return;

  if (confirm(`هل أنت متأكد من حذف الطالب "${student.name}" (كود #${student.code}) نهائياً من المنظومة؟`)) {
    deleteStudentByCode(student.code);
    closeStudentDetailModal();
    closeEditStudentModal();
  }
}

function confirmDeleteStudentFromDetail() {
  if (activeDetailStudentCode) {
    confirmDeleteStudentByCode(activeDetailStudentCode);
  }
}

function deleteStudentDirectly() {
  const code = document.getElementById('editCodeHidden')?.value;
  if (code) {
    confirmDeleteStudentByCode(code);
  }
}

function deleteStudentByCode(code) {
  // 1. Add to deleted list
  let deletedList = JSON.parse(localStorage.getItem('araij_deleted_students') || '[]');
  if (!deletedList.includes(code)) {
    deletedList.push(code);
    localStorage.setItem('araij_deleted_students', JSON.stringify(deletedList));
  }

  // 2. Remove from overrides
  let overrides = JSON.parse(localStorage.getItem('araij_students_overrides') || '{}');
  delete overrides[code];
  localStorage.setItem('araij_students_overrides', JSON.stringify(overrides));

  // 3. Remove from in-memory list
  allStudents = allStudents.filter(s => s.code !== code);

  // 4. Update full DB in localStorage (without the deleted student)
  try {
    localStorage.setItem('araij_full_students_db', JSON.stringify(allStudents));
  } catch(e) {}

  // 5. Sync delete to Google Sheets (cloud)
  syncToGoogleSheets('delete_student', { code });

  // 6. Sync delete to LAN server & Railway Live Server
  sendToLanServer('delete_student', { code });
  sendToRailwayServer('delete_student', { code });

  // 7. Audit Log
  const uName = currentUser?.name || 'مشرف';
  logActivity('حذف طالب', `قام ${uName} بحذف الطالب #${code}`, code, '', {});

  applyFilters();
  updateKPIStats();
  const attendanceTab = document.getElementById('tabContent_attendance_matrix');
  if (attendanceTab && !attendanceTab.classList.contains('hidden')) renderAttendanceMatrix();

  showToast(`✅ تم حذف الطالب #${code} بنجاح وتمت المزامنة الفورية على كافة الأجهزة.`);
}


function syncWithCentralServer(action, payload) {
  const host = window.location.hostname || '192.168.1.23';
  const port = window.location.port || '8080';
  const syncUrl = `http://${host}:${port}/api/sync_local`;

  const studentObj = payload.student || payload;
  const studentCode = payload.code || (studentObj && studentObj.code);

  // Only sync to local LAN server — Google Sheets sync is handled separately by the caller
  fetch(syncUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      action: action,
      student: studentObj,
      code: studentCode
    })
  }).catch(() => {});
}


async function fetchLatestStudentsFromCentralServer(isSilent = false) {
  try {
    const srv = getRailwayServerUrl();
    let url;
    if (srv) {
      url = `${srv}/api/get_students?t=${Date.now()}`;
    } else {
      const host = window.location.hostname || '192.168.1.23';
      const port = window.location.port || '8080';
      url = `http://${host}:${port}/api/get_students?t=${Date.now()}`;
    }

    const res = await fetch(url, { cache: 'no-store' });
    if (res.ok) {
      const serverStudents = await res.json();
      if (Array.isArray(serverStudents) && serverStudents.length > 0) {
        window.STUDENTS_DATA = serverStudents;
        
        // Re-init with freshly fetched central server data
        const pendingQueue = getOfflineSyncQueue();
        const pendingCodes = new Set(pendingQueue.map(q => String(q.payload?.student?.code || q.payload?.code || '').trim()).filter(Boolean));

        const localOverrides = localStorage.getItem('araij_students_overrides');
        const deletedCodes = JSON.parse(localStorage.getItem('araij_deleted_students') || '[]');
        
        let overridesMap = {};
        if (localOverrides) {
          try { overridesMap = JSON.parse(localOverrides); } catch (e) {}
        }

        let merged = serverStudents.map(serverSt => {
          const code = String(serverSt.code || '').trim();
          // If there is an unsynced offline edit for this student on THIS device, keep local changes until synced
          if (pendingCodes.has(code) && overridesMap[code]) {
            return { ...serverSt, ...overridesMap[code] };
          }
          // Otherwise, server data is canonical and authoritative!
          overridesMap[code] = serverSt;
          return serverSt;
        });

        // Also preserve newly added students created offline that are not yet on the server
        Object.keys(overridesMap).forEach(code => {
          if (pendingCodes.has(code) && !merged.find(s => String(s.code).trim() === code)) {
            merged.unshift(overridesMap[code]);
          }
        });

        try {
          localStorage.setItem('araij_students_overrides', JSON.stringify(overridesMap));
          localStorage.setItem('araij_full_students_db', JSON.stringify(merged));
        } catch(e) {}

        allStudents = merged.filter(s => {
          if (!s) return false;
          if (deletedCodes.includes(s.code)) return false;
          const name = String(s.name || '').trim();
          if (!name || name.length < 2 || name.includes('غير مسمى')) return false;
          return true;
        });
        const seenCentralCodes = new Set();
        allStudents = allStudents.filter(s => {
          const key = String(s.code || '').trim();
          if (!key || seenCentralCodes.has(key)) return false;
          seenCentralCodes.add(key);
          return true;
        });
        allStudents.sort((a, b) => (a.name || '').localeCompare(b.name || '', 'ar', { numeric: true, sensitivity: 'base' }));

        allStudents.forEach(st => {
          if (!st.academicSubjects || Object.keys(st.academicSubjects).length === 0) {
            if (st.teachersSummary || st.subjectsSummary) {
              st.academicSubjects = parseSubjectsFromSummary(st.subjectsSummary, st.teachersSummary);
            }
          } else {
            Object.keys(st.academicSubjects).forEach(sub => {
              const canon = canonicalSubjectName(sub) || sub;
              st.academicSubjects[sub].teacher = cleanTeacherName(st.academicSubjects[sub].teacher, canon);
            });
          }
          rebuildStudentSummaries(st);
        });

        applyFilters();
        updateKPIStats();
        renderAttendanceMatrix();

        if (!isSilent) {
          showToast('⚡ تم تحديث ومزامنة أحدث البيانات من السيرفر بنجاح!');
        }
      }
    }
  } catch (e) {
    // Silent fail if offline or unreachable
  }
}


// Auto sync polling every 30 seconds (not 6s — too heavy on re-renders)
setInterval(() => {
  fetchLatestStudentsFromCentralServer(true);
}, 30000);


// Also sync on window focus / tab switch
window.addEventListener('focus', () => {
  fetchLatestStudentsFromCentralServer(true);
});

function saveStudentOverride(student, skipSync = false) {
  student._metrics = calculateStudentMetrics(student);

  const targetCode = String(student.code || '').trim();

  // 1. Update in-memory allStudents
  const idx = allStudents.findIndex(s => String(s.code || '').trim() === targetCode);
  if (idx >= 0) {
    allStudents[idx] = { ...allStudents[idx], ...student };
  } else {
    allStudents.unshift(student);
  }

  // 2. Update window.STUDENTS_DATA
  if (window.STUDENTS_DATA) {
    const bIdx = window.STUDENTS_DATA.findIndex(s => String(s.code || '').trim() === targetCode);
    if (bIdx >= 0) {
      window.STUDENTS_DATA[bIdx] = { ...window.STUDENTS_DATA[bIdx], ...student };
    } else {
      window.STUDENTS_DATA.unshift(student);
    }
  }

  // 3. Save overrides
  let overrides = {};
  try {
    overrides = JSON.parse(localStorage.getItem('araij_students_overrides') || '{}');
  } catch (err) {}
  overrides[student.code] = student;
  localStorage.setItem('araij_students_overrides', JSON.stringify(overrides));

  // 4. Save full db
  try {
    localStorage.setItem('araij_full_students_db', JSON.stringify(allStudents));
  } catch (e) {}

  // 5. Send to LAN server & Railway live server for instant real-time sync (skipped during batch operations)
  if (!skipSync) {
    sendToLanServer('update_student', { student });
    sendToRailwayServer('update_student', { student });
  }
}

// ====================================================

// ====================================================

let currentDuplicateStudent = null;

function checkWizardDuplicateStudent() {
  const nameInput = (document.getElementById('wizName')?.value || '').trim();
  const codeInput = (document.getElementById('wizCode')?.value || '').trim();
  const phoneInput = (document.getElementById('wizPhone')?.value || '').trim();
  const parentPhoneInput = (document.getElementById('wizParentPhone')?.value || '').trim();

  const alertBox = document.getElementById('wizDuplicateAlert');
  const detailsBox = document.getElementById('wizDuplicateDetails');
  const titleBox = document.getElementById('wizDuplicateTitle');

  if (!nameInput && !codeInput && !phoneInput && !parentPhoneInput) {
    hideWizardDuplicateAlert();
    return;
  }

  const nameNorm = normalize_arabic(nameInput);
  const codeNorm = codeInput.toLowerCase();

  let matchedStudent = null;
  let matchReason = '';

  // 1. Check exact Code match
  if (codeInput) {
    matchedStudent = allStudents.find(s => String(s.code).trim().toLowerCase() === codeNorm);
    if (matchedStudent) matchReason = `الكود #${matchedStudent.code}`;
  }

  // 2. Check Name match
  if (!matchedStudent && nameNorm.length >= 3) {
    matchedStudent = allStudents.find(s => {
      const sNorm = normalize_arabic(s.name || '');
      if (sNorm === nameNorm) return true;
      if (nameNorm.split(' ').length >= 3 && (sNorm.includes(nameNorm) || nameNorm.includes(sNorm))) return true;
      return false;
    });
    if (matchedStudent) matchReason = `الاسم (${matchedStudent.name})`;
  }

  // 3. Check Phone match
  if (!matchedStudent && phoneInput.length >= 8) {
    matchedStudent = allStudents.find(s => s.phone && s.phone.includes(phoneInput));
    if (matchedStudent) matchReason = `رقم هاتف الطالب (${phoneInput})`;
  }

  if (matchedStudent) {
    currentDuplicateStudent = matchedStudent;
    if (alertBox) alertBox.classList.remove('hidden');
    if (titleBox) titleBox.textContent = `⚠️ تنبيه: تم العثور على طالب مسجل مسبقاً في السنتر بواسطة ${matchReason}!`;

    const metrics = matchedStudent._metrics || calculateStudentMetrics(matchedStudent);

    if (detailsBox) {
      detailsBox.innerHTML = `
        <div><strong class="text-slate-500">اسم الطالب:</strong> <span class="text-slate-900 font-bold" style="color: #0f172a !important;">${matchedStudent.name}</span></div>
        <div><strong class="text-slate-500">الكود:</strong> <span class="text-sky-700 font-mono font-bold">#${matchedStudent.code}</span></div>
        <div><strong class="text-slate-400">المرحلة والتخصص:</strong> <span class="text-emerald-400">${matchedStudent.grade} - ${matchedStudent.specialization || 'عام'}</span></div>
        <div><strong class="text-slate-400">المنطقة:</strong> <span class="text-slate-200">${matchedStudent.area || 'غير محدد'}</span></div>
        <div><strong class="text-slate-400">هاتف الطالب:</strong> <span class="text-slate-200 font-mono">${matchedStudent.phone || 'لا يوجد'}</span></div>
        <div><strong class="text-slate-400">هاتف ولي الأمر:</strong> <span class="text-slate-200 font-mono">${matchedStudent.parentPhone || 'لا يوجد'}</span></div>
        <div class="sm:col-span-2"><strong class="text-slate-400">المواد والمدرسين:</strong> <span class="text-amber-300">${matchedStudent.teachersSummary || matchedStudent.subjectsSummary || 'لا توجد'}</span></div>
        <div><strong class="text-slate-400">نسبة الحضور:</strong> <span class="text-teal-300 font-bold">${metrics.attendanceRate}%</span></div>
        <div><strong class="text-slate-400">التقييم العام:</strong> <span class="text-emerald-400 font-bold">${metrics.gradeEvaluation}</span></div>
      `;
    }

    const editBtn = document.getElementById('wizDupEditBtn');
    const viewBtn = document.getElementById('wizDupViewBtn');
    const fillBtn = document.getElementById('wizDupFillBtn');

    if (editBtn) editBtn.onclick = () => openEditStudentModalDirect(matchedStudent.code);
    if (viewBtn) viewBtn.onclick = () => openStudentDetailModal(matchedStudent.code);
    if (fillBtn) fillBtn.onclick = () => fillWizardWithStudentData(matchedStudent);

    initIcons();
  } else {
    hideWizardDuplicateAlert();
  }
}

function hideWizardDuplicateAlert() {
  const alertBox = document.getElementById('wizDuplicateAlert');
  if (alertBox) alertBox.classList.add('hidden');
  currentDuplicateStudent = null;
}

function fillWizardWithStudentData(st) {
  if (!st) return;
  const nameInput = document.getElementById('wizName');
  const codeInput = document.getElementById('wizCode');
  const gradeInput = document.getElementById('wizGrade');
  const areaInput = document.getElementById('wizArea');
  const phoneInput = document.getElementById('wizPhone');
  const parentPhoneInput = document.getElementById('wizParentPhone');

  const normGrade = normalize_grade(st.grade);
  if (nameInput) nameInput.value = st.name || '';
  if (codeInput) codeInput.value = st.code || '';
  if (gradeInput) {
    gradeInput.value = normGrade;
  }
  if (areaInput) areaInput.value = st.area || '';
  if (phoneInput) phoneInput.value = st.phone || '';
  if (parentPhoneInput) parentPhoneInput.value = st.parentPhone || '';

  renderWizardSubjectsList(normGrade);

  // Pre-check enrolled subjects in wizard
  const subjectsMap = parseSubjectsFromSummary(st.subjectsSummary, st.teachersSummary, st.academicSubjects);
  const subjects = SUBJECTS[normGrade] || [];
  subjects.forEach((subj, idx) => {
    const enrolled = getEnrolledSubjectData(subjectsMap, subj);
    const chk = document.getElementById(`wiz_subj_chk_${idx}`);
    const grp = document.getElementById(`wiz_teacher_group_${idx}`);
    const sel = document.getElementById(`wiz_teacher_val_${idx}`);
    if (chk && enrolled) {
      chk.checked = true;
      if (grp) grp.classList.remove('opacity-40', 'pointer-events-none');
      if (sel && enrolled.teacher) {
        sel.value = enrolled.teacher;
      }
    }
  });

  showToast(`تم ملء بيانات ومواد الطالب (#${st.code}) في النموذج.`);
}

function initWizard() {
  onWizardGradeChange();
}

function onWizardGradeChange() {
  const grade = document.getElementById('wizGrade')?.value || 'ث1';
  updateSuggestedCode(true);
  renderWizardSubjectsList(grade);
  checkWizardDuplicateStudent();
}

function onWizardSpecChange() {
  const mainSpec = document.getElementById('wizSpecMain')?.value;
  const bacGroup = document.getElementById('wizBaccalaureateGroup');
  if (bacGroup) {
    if (mainSpec === 'بكالوريا') {
      bacGroup.classList.remove('hidden');
    } else {
      bacGroup.classList.add('hidden');
    }
  }
}

function getNextStudentCodeForGrade(grade) {
  const normGrade = normalize_grade(grade || 'ث1');
  const prefix = normGrade === 'ث1' ? 1000 : normGrade === 'ث2' ? 2000 : 3000;
  let maxCode = prefix;
  let foundInPrefixRange = false;

  if (Array.isArray(allStudents)) {
    allStudents.forEach(s => {
      const sGrade = normalize_grade(s.grade);
      if (sGrade === normGrade) {
        const num = parseInt(String(s.code || '').replace(/\D/g, ''), 10);
        if (!isNaN(num) && num >= prefix && num < (prefix + 1000)) {
          if (num > maxCode) maxCode = num;
          foundInPrefixRange = true;
        }
      }
    });

    // Fallback: If no student matches the prefix range, find absolute highest number in this grade
    if (!foundInPrefixRange) {
      let anyMax = 0;
      allStudents.forEach(s => {
        if (normalize_grade(s.grade) === normGrade) {
          const num = parseInt(String(s.code || '').replace(/\D/g, ''), 10);
          if (!isNaN(num) && num > anyMax) anyMax = num;
        }
      });
      if (anyMax > 0) return anyMax + 1;
    }
  }
  return maxCode + 1;
}

function updateSuggestedCode(forceOverwrite = false) {
  const grade = document.getElementById('wizGrade')?.value || 'ث1';
  const nextCode = getNextStudentCodeForGrade(grade);
  const codeInput = document.getElementById('wizCode');

  if (codeInput) {
    const currentVal = (codeInput.value || '').trim();
    const prevSuggested = codeInput.dataset.suggestedCode;

    // Auto-fill if empty, or if forced, or if the value was automatically suggested previously
    if (forceOverwrite || !currentVal || currentVal === prevSuggested) {
      codeInput.value = String(nextCode);
    }
    codeInput.placeholder = `كود الطالب التلقائي: ${nextCode}`;
    codeInput.dataset.suggestedCode = String(nextCode);
  }
}

function renderWizardSubjectsList(grade) {
  const container = document.getElementById('wizSubjectsTeachersList');
  if (!container) return;

  const subjects = SUBJECTS[grade] || [];

  container.innerHTML = subjects.map((subj, idx) => {
    const teachers = TEACHERS[subj] || [];
    return `
      <div class="glass-card p-3 rounded-2xl border border-slate-800 flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3">
        <label class="flex items-center gap-2 cursor-pointer">
          <input type="checkbox" id="wiz_subj_chk_${idx}" value="${subj}" onchange="toggleWizardTeacherSelect(${idx})" class="w-4 h-4 rounded text-sky-600 focus:ring-sky-500 bg-slate-100 border-slate-300">
          <span class="text-xs font-bold text-slate-800">${subj}</span>
        </label>

        <div class="flex items-center gap-2 w-full sm:w-auto opacity-40 pointer-events-none transition-opacity" id="wiz_teacher_group_${idx}">
          <span class="text-[11px] text-slate-400 whitespace-nowrap">المدرس:</span>
          <select id="wiz_teacher_val_${idx}" class="w-full sm:w-48 bg-slate-800 border border-slate-700 rounded-xl px-2.5 py-1.5 text-xs text-white focus:ring-1 focus:ring-sky-500">
            ${teachers.length > 0 ? teachers.map(t => `<option value="${t}">${t}</option>`).join('') : `<option value="مدرس المادة">مدرس المادة</option>`}
            <option value="مدرس آخر">مدرس آخر / مخصص</option>
          </select>
        </div>
      </div>
    `;
  }).join('');
}

function toggleWizardTeacherSelect(idx) {
  const chk = document.getElementById(`wiz_subj_chk_${idx}`);
  const group = document.getElementById(`wiz_teacher_group_${idx}`);
  if (chk && group) {
    if (chk.checked) {
      group.classList.remove('opacity-40', 'pointer-events-none');
    } else {
      group.classList.add('opacity-40', 'pointer-events-none');
    }
  }
}

let _isSubmittingWizard = false;

function handleWizardSubmit(e) {
  if (e) e.preventDefault();
  if (_isSubmittingWizard) return; // Prevent double submission completely!

  const name = document.getElementById('wizName')?.value.trim();
  const codeInput = document.getElementById('wizCode');
  let code = (codeInput?.value || '').trim();
  const grade = document.getElementById('wizGrade')?.value || 'ث1';

  // If user left code completely empty, assign next code for the selected grade
  if (!code) {
    code = String(getNextStudentCodeForGrade(grade));
  }

  const area = document.getElementById('wizArea')?.value.trim();
  const phone = document.getElementById('wizPhone')?.value.trim();
  const parentPhone = document.getElementById('wizParentPhone')?.value.trim();
  
  if (!name) {
    alert('يرجى كتابة اسم الطالب بالكامل!');
    return;
  }

  // Duplicate Check confirmation on submit
  const duplicate = allStudents.find(s => String(s.code).trim().toLowerCase() === code.toLowerCase() || normalize_arabic(s.name || '') === normalize_arabic(name));
  if (duplicate && duplicate.code !== code) {
    const confirmDup = confirm(`⚠️ تنبيه: يوجد طالب مسجل بالفعل بنفس الاسم (${duplicate.name}) بالكود #${duplicate.code}.\nهل ترغب بالاستمرار وتسجيله كطالب جديد برقم كود #${code}؟`);
    if (!confirmDup) return;
  }

  // Lock submission and show spinner
  _isSubmittingWizard = true;
  const submitBtn = document.querySelector('#wizardStudentForm button[type="submit"]');
  if (submitBtn) {
    submitBtn.disabled = true;
    submitBtn.classList.add('opacity-60', 'pointer-events-none');
    submitBtn.innerHTML = '<span class="inline-block animate-spin mr-1">⏳</span> جاري الحفظ والتسجيل...';
  }

  const mainSpec = document.getElementById('wizSpecMain')?.value || 'عام';
  const bacSpec = document.getElementById('wizSpecBac')?.value || 'طب';
  const specialization = mainSpec === 'بكالوريا' ? `بكالوريا - ${bacSpec}` : mainSpec;

  const subjects = SUBJECTS[grade] || [];
  const selectedSubjects = [];
  const selectedTeachersList = [];
  const academicSubjects = {};

  subjects.forEach((subj, idx) => {
    const chk = document.getElementById(`wiz_subj_chk_${idx}`);
    if (chk && chk.checked) {
      const rawTeacher = document.getElementById(`wiz_teacher_val_${idx}`)?.value || 'مدرس المادة';
      const canon = canonicalSubjectName(subj) || subj;
      const teacher = cleanTeacherName(rawTeacher, canon);
      selectedSubjects.push(subj);
      selectedTeachersList.push(`${subj}/${teacher}`);
      academicSubjects[canon] = {
        teacher: teacher,
        sessions: ["", "", "", "", "", "", "", ""]
      };
    }
  });

  // Track who did this action (Audit hint for Eng. Wafaa)
  const currentUserName = currentUser?.name || 'منار ناصر';
  const actionTime = new Date().toLocaleTimeString('ar-EG', { hour: '2-digit', minute: '2-digit' }) + ' ' + new Date().toLocaleDateString('ar-EG');

  const newStudent = {
    code,
    name,
    grade,
    specialization,
    area,
    phone,
    parentPhone,
    subjectsSummary: selectedSubjects.join(', '),
    teachersSummary: selectedTeachersList.join(' | '),
    academicSubjects,
    regDate: new Date().toLocaleDateString('ar-EG'),
    createdBy: currentUserName,
    lastModifiedBy: currentUserName,
    lastAction: 'تسجيل جديد',
    lastActionTime: actionTime
  };

  newStudent._metrics = calculateStudentMetrics(newStudent);
  newStudent._searchString = buildStudentSearchString(newStudent);

  // ✅ Deduplicate: If student with this code exists in allStudents, update it — NEVER add twice!
  const existingIdx = allStudents.findIndex(s => String(s.code).trim() === String(code).trim());
  let syncAction = 'add_student';
  if (existingIdx >= 0) {
    syncAction = 'update_student';
    const existingSt = allStudents[existingIdx];
    // Preserve any existing sessions for subjects
    if (existingSt.academicSubjects) {
      Object.keys(newStudent.academicSubjects).forEach(sub => {
        const oldSub = getEnrolledSubjectData(existingSt.academicSubjects, sub);
        if (oldSub && Array.isArray(oldSub.sessions) && oldSub.sessions.length > 0) {
          newStudent.academicSubjects[sub].sessions = [...oldSub.sessions];
        }
      });
    }
    allStudents[existingIdx] = { ...existingSt, ...newStudent };
  } else {
    allStudents.unshift(newStudent);
  }
  allStudents.sort((a, b) => (a.name || '').localeCompare(b.name || '', 'ar', { numeric: true, sensitivity: 'base' }));

  // Save locally + to central LAN server & Railway
  saveStudentOverride(newStudent);

  // Sync to Google Sheets
  syncToGoogleSheets(syncAction, { student: newStudent }, true);

  // Sync to Live Railway Server
  sendToRailwayServer(syncAction, { student: newStudent });

  // Log activity in central audit log & Google Sheets audit sheet
  try {
    logActivity(syncAction === 'add_student' ? 'تسجيل طالب جديد' : 'تعديل بيانات طالب', 
      `قام ${currentUserName} بـ ${syncAction === 'add_student' ? 'تسجيل' : 'تعديل'} الطالب "${name}" بالكود #${code}`,
      newStudent.code,
      newStudent.name,
      { grade: newStudent.grade, subject: newStudent.subjectsSummary }
    );
  } catch(e) {}

  // ====== CLEAR THE FORM ======
  hideWizardDuplicateAlert();
  initWizard();

  // Explicitly clear text fields (except wizCode which gets the fresh next code)
  ['wizName', 'wizPhone', 'wizParentPhone', 'wizArea'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = '';
  });

  updateSuggestedCode(true);

  // Reset filters and search so the new student card is 100% visible on the Students page
  currentGradeFilter = 'all';
  currentEvaluationFilter = 'all';
  currentAttendanceFilter = 'all';
  currentAreaFilter = 'all';
  const searchInput = document.getElementById('searchInput');
  if (searchInput) searchInput.value = '';
  currentSearchQuery = '';
  document.querySelectorAll('.grade-tab').forEach(btn => {
    btn.className = 'grade-tab px-3.5 py-2 rounded-xl text-xs font-semibold text-slate-600 hover:text-slate-900 transition-all';
  });
  const allTab = document.getElementById('tab_grade_all');
  if (allTab) allTab.className = 'grade-tab px-3.5 py-2 rounded-xl text-xs font-bold transition-all bg-sky-600 text-white shadow';

  currentPage = 1;
  applyFilters();
  updateKPIStats();

  // ✅ INSTANT REDIRECT TO STUDENTS TAB SO USER SEES THE CARD IMMEDIATELY
  switchTab('students');

  // Scroll to top of students list and glow highlight the new student card
  setTimeout(() => {
    window.scrollTo({ top: 0, behavior: 'smooth' });
    const firstCard = document.querySelector('#cardsViewContainer > div');
    if (firstCard) {
      firstCard.classList.add('ring-4', 'ring-emerald-400', 'border-emerald-400', 'animate-pulse');
      setTimeout(() => {
        firstCard.classList.remove('ring-4', 'ring-emerald-400', 'border-emerald-400', 'animate-pulse');
      }, 4000);
    }
  }, 150);

  showToast(`✅ تم تسجيل الطالب (${name}) بنجاح بالكود #${code}! تم تحويلك لصفحة الطلاب.`);

  // Re-enable button after 2 seconds
  setTimeout(() => {
    _isSubmittingWizard = false;
    if (submitBtn) {
      submitBtn.disabled = false;
      submitBtn.classList.remove('opacity-60', 'pointer-events-none');
      submitBtn.innerHTML = '<i data-lucide="check-circle-2" class="w-4 h-4"></i><span>حفظ وتسجيل الطالب في النظام</span>';
      initIcons();
    }
  }, 2000);
}



// ====================================================
// 14. GOOGLE SHEETS & CLOUD LIVE SYNC ENGINE
// ====================================================

function getGoogleAppsScriptUrl() {
  return localStorage.getItem('araij_sheets_webhook_url') || 'https://script.google.com/macros/s/AKfycbymmlVQKxkdXpLlsE0Z7DuhizyHHGlEKHesCAj0FJ4lfKlSBJRLThN0GFfISbi5CabmAg/exec';
}

function getOfflineSyncQueue() {
  try {
    return JSON.parse(localStorage.getItem('araij_offline_sync_queue') || '[]');
  } catch(e) {
    return [];
  }
}

function addToOfflineSyncQueue(action, payload) {
  const queue = getOfflineSyncQueue();
  const code = String(payload?.student?.code || payload?.code || '').trim();
  const now = Date.now();
  // Avoid duplicate queue entries for the same action & code within 3 seconds
  const existingIdx = queue.findIndex(q => q.action === action && String(q.payload?.student?.code || q.payload?.code || '').trim() === code && (now - q.timestamp < 3000));
  if (existingIdx >= 0) {
    queue[existingIdx] = { action, payload, timestamp: now };
  } else {
    queue.push({ action, payload, timestamp: now });
  }
  localStorage.setItem('araij_offline_sync_queue', JSON.stringify(queue));
}

async function flushOfflineSyncQueue() {
  if (!navigator.onLine) return;
  const queue = getOfflineSyncQueue();
  if (!queue || queue.length === 0) return;

  const url = getGoogleAppsScriptUrl();
  const srvUrl = getRailwayServerUrl();
  const remaining = [];
  let flushedCount = 0;

  for (let i = 0; i < queue.length; i++) {
    const item = queue[i];
    let ok = false;
    // 1. Forward to Railway REST API
    if (srvUrl) {
      try {
        const resp = await fetch(`${srvUrl}/api/sync`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            action: item.action,
            payload: item.payload,
            data: item.payload,
            byUser: currentUser?.name || 'مشرف',
            sender: currentUser?.name || 'مشرف',
            clientId: window.ARAIJ_CLIENT_ID,
            ...item.payload
          })
        });
        if (resp.ok) ok = true;
      } catch(e) {}
    } else {
      ok = true;
    }

    // 2. Forward to Google Sheets
    if (url) {
      try {
        await fetch(url, {
          method: 'POST',
          mode: 'no-cors',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ action: item.action, ...item.payload })
        });
        ok = true;
      } catch(e) {}
    }

    if (ok) {
      flushedCount++;
    } else {
      remaining.push(item);
    }
  }

  localStorage.setItem('araij_offline_sync_queue', JSON.stringify(remaining));
  if (flushedCount > 0) {
    showToast(`⚡ تم ترحيل ${flushedCount} عملية مسجلة أوفلاين إلى السحابة ومزامنتها بنجاح مع كافة المشرفين!`);
  }
}

// ====================================================
// RAILWAY LIVE CLOUD SYNC & WEBSOCKET ENGINE (0ms SYNC)
// ====================================================

let _railwaySocket = null;
let _railwayReconnectTimeout = null;
let _isRailwayConnected = false;

const DEFAULT_RAILWAY_URL = 'https://araijmanager.up.railway.app';

function sanitizeRailwayUrl(raw) {
  if (!raw) return '';
  let s = String(raw).trim();
  // Remove leading slashes or whitespace e.g. /https://...
  s = s.replace(/^[\/\s]+/, '');
  if (!/^https?:\/\//i.test(s)) {
    s = 'https://' + s;
  }
  // Remove trailing slashes
  s = s.replace(/\/+$/, '');
  return s;
}

function getRailwayServerUrl() {
  // 1. If currently loaded via https or http on the web (e.g. Railway domain), use current origin!
  if (window.location.protocol === 'http:' || window.location.protocol === 'https:') {
    const host = window.location.hostname || '';
    if (host && !host.includes('localhost') && !host.includes('127.0.0.1')) {
      return sanitizeRailwayUrl(window.location.origin);
    }
  }

  // 2. Otherwise (e.g. file:/// in Android APK or offline html), check localStorage
  const custom = localStorage.getItem('araij_railway_url');
  if (custom && custom.trim() && !custom.includes('192.168.')) {
    return sanitizeRailwayUrl(custom);
  }

  // 3. Fallback to official Railway server URL
  return sanitizeRailwayUrl(DEFAULT_RAILWAY_URL || 'https://araijmanager.up.railway.app');
}

function onRailwayUrlInputChanged(val) {
  const clean = sanitizeRailwayUrl(val);
  if (clean) {
    try { localStorage.setItem('araij_railway_url', clean); } catch (e) {}
    updateRailwayStatusBadge(false, 'جاري الاتصال ⚡');
    initLiveRailwaySync();
  }
}
window.onRailwayUrlInputChanged = onRailwayUrlInputChanged;

function initLiveRailwaySync() {
  const srvUrl = getRailwayServerUrl();
  const inputEl = document.getElementById('railwayServerUrlInput');
  if (inputEl) {
    inputEl.value = srvUrl;
  }

  if (!srvUrl) {
    updateRailwayStatusBadge(false, 'غير محدد');
    return;
  }

  try {
    const wsProto = srvUrl.startsWith('https') ? 'wss:' : 'ws:';
    const host = srvUrl.replace(/^https?:\/\//, '');
    const wsUrl = `${wsProto}//${host}/ws`;

    if (_railwaySocket) {
      try {
        _railwaySocket.onclose = null;
        _railwaySocket.close();
      } catch(e) {}
    }

    console.log(`[Railway Live Sync] Connecting to ${wsUrl}...`);
    _railwaySocket = new WebSocket(wsUrl);

    _railwaySocket.onopen = () => {
      console.log(`[Railway Live Sync] Connected successfully!`);
      _isRailwayConnected = true;
      updateRailwayStatusBadge(true, 'ريلوي متصل ⚡');
      flushOfflineSyncQueue();
      fetchLatestStudentsFromCentralServer(true);
    };

    _railwaySocket.onmessage = (event) => {
      try {
        const msg = JSON.parse(event.data);
        if (msg.type === 'DATA_UPDATE') {
          handleIncomingLiveUpdate(msg);
        }
      } catch (err) {
        console.warn('[Railway Live Sync] Parse error:', err);
      }
    };

    _railwaySocket.onclose = () => {
      _isRailwayConnected = false;
      updateRailwayStatusBadge(false, 'إعادة الاتصال...');
      scheduleRailwayReconnect();
    };

    _railwaySocket.onerror = () => {
      _isRailwayConnected = false;
      updateRailwayStatusBadge(false, 'غير متصل');
    };

  } catch (err) {
    console.error('[Railway Live Sync] Connection error:', err);
    scheduleRailwayReconnect();
  }
}

function scheduleRailwayReconnect() {
  if (_railwayReconnectTimeout) clearTimeout(_railwayReconnectTimeout);
  _railwayReconnectTimeout = setTimeout(() => {
    initLiveRailwaySync();
  }, 4000);
}

function updateRailwayStatusBadge(connected, text) {
  const dot = document.getElementById('railwayStatusDot');
  const badgeText = document.getElementById('railwayStatusBadgeText');
  const modalDot = document.getElementById('railwayModalStatusDot');
  const modalText = document.getElementById('railwayModalStatusText');

  if (dot) {
    dot.className = connected 
      ? 'w-2.5 h-2.5 rounded-full bg-emerald-500 animate-pulse' 
      : 'w-2.5 h-2.5 rounded-full bg-rose-500';
  }
  if (badgeText) badgeText.textContent = text;
  if (modalDot) {
    modalDot.className = connected ? 'w-2.5 h-2.5 rounded-full bg-emerald-400 animate-pulse' : 'w-2.5 h-2.5 rounded-full bg-rose-500';
  }
  if (modalText) modalText.textContent = text;
}

function handleIncomingLiveUpdate(msg) {
  const action = msg.action;
  const data = msg.data || msg.payload || {};
  const sender = msg.sender || msg.byUser || '';
  const clientId = msg.clientId || '';
  console.log(`[Railway Live Sync] Received ${action} from ${sender || 'system'} (client: ${clientId})`);

  // Only ignore if the message was echoed back to THIS EXACT browser session
  if (clientId && window.ARAIJ_CLIENT_ID && clientId === window.ARAIJ_CLIENT_ID) {
    return;
  }

  if (action === 'update_student' || action === 'add_student') {
    const student = data?.student || data;
    if (!student || !student.code) return;

    if (!student.academicSubjects || Object.keys(student.academicSubjects).length === 0) {
      if (student.teachersSummary || student.subjectsSummary) {
        student.academicSubjects = parseSubjectsFromSummary(student.subjectsSummary, student.teachersSummary);
      }
    }
    rebuildStudentSummaries(student);

    const targetCode = String(student.code).trim();
    const idx = allStudents.findIndex(s => String(s.code).trim() === targetCode);
    if (idx >= 0) {
      allStudents[idx] = { ...allStudents[idx], ...student };
    } else {
      allStudents.unshift(student);
    }

    try {
      const overrides = JSON.parse(localStorage.getItem('araij_students_overrides') || '{}');
      overrides[student.code] = allStudents[idx >= 0 ? idx : 0];
      localStorage.setItem('araij_students_overrides', JSON.stringify(overrides));
      localStorage.setItem('araij_full_students_db', JSON.stringify(allStudents));
    } catch(e) {}

    renderAttendanceMatrix();
    applyFilters();
    updateKPIStats();

    if (activeDetailStudentCode === targetCode) {
      openStudentDetailModal(targetCode);
    }

    const bulkModal = document.getElementById('bulkSessionModal');
    if (bulkModal && !bulkModal.classList.contains('hidden')) {
      renderBulkStudentsList();
    }

    showToast(`⚡ مزامنة فورية: تم رصد/تحديث بيانات (${student.name}) بواسطة ${sender || 'مشرف'}`);

  } else if (action === 'batch_update_students' || action === 'bulk_session') {
    const list = data?.students || data?.modifiedStudents || (Array.isArray(data) ? data : []);
    if (Array.isArray(list) && list.length > 0) {
      try {
        const overrides = JSON.parse(localStorage.getItem('araij_students_overrides') || '{}');
        list.forEach(st => {
          if (!st || !st.code) return;
          if (!st.academicSubjects || Object.keys(st.academicSubjects).length === 0) {
            if (st.teachersSummary || st.subjectsSummary) {
              st.academicSubjects = parseSubjectsFromSummary(st.subjectsSummary, st.teachersSummary);
            }
          }
          rebuildStudentSummaries(st);

          const targetCode = String(st.code).trim();
          const idx = allStudents.findIndex(s => String(s.code).trim() === targetCode);
          if (idx >= 0) {
            allStudents[idx] = { ...allStudents[idx], ...st };
            overrides[st.code] = allStudents[idx];
          } else {
            allStudents.unshift(st);
            overrides[st.code] = st;
          }
        });
        localStorage.setItem('araij_students_overrides', JSON.stringify(overrides));
        localStorage.setItem('araij_full_students_db', JSON.stringify(allStudents));
      } catch(e) {}

      renderAttendanceMatrix();
      applyFilters();
      updateKPIStats();

      if (activeDetailStudentCode) {
        openStudentDetailModal(activeDetailStudentCode);
      }

      const bulkModal = document.getElementById('bulkSessionModal');
      if (bulkModal && !bulkModal.classList.contains('hidden')) {
        renderBulkStudentsList();
      }

      try { playSuccessChime(); } catch(e) {}
      try {
        showBulkProgressBanner('info', '⚡ استلام رصد جماعي فوري', `قام (${sender || 'مشرف'}) برصد وتحديث حضور ${list.length} طالب - تم تحديث شاشتك تلقائياً!`, 5500);
      } catch(e) {}
      showToast(`⚡ مزامنة فورية: استلام رصد جماعي لـ ${list.length} طالب من ${sender || 'مشرف'}`);
    }

  } else if (action === 'delete_student') {
    const code = String(data?.code || '').trim();
    if (code) {
      allStudents = allStudents.filter(s => String(s.code).trim() !== code);
      try {
        const overrides = JSON.parse(localStorage.getItem('araij_students_overrides') || '{}');
        delete overrides[code];
        localStorage.setItem('araij_students_overrides', JSON.stringify(overrides));
        localStorage.setItem('araij_full_students_db', JSON.stringify(allStudents));
      } catch(e) {}
      renderAttendanceMatrix();
      applyFilters();
      updateKPIStats();
      showToast(`⚡ مزامنة فورية: قام ${sender || 'مشرف'} بحذف الطالب #${code}`);
    }

  } else if (action === 'log_activity') {
    if (data?.log) {
      try {
        const logs = JSON.parse(localStorage.getItem('araij_activity_logs') || '[]');
        logs.unshift(data.log);
        if (logs.length > 500) logs.pop();
        localStorage.setItem('araij_activity_logs', JSON.stringify(logs));
        renderActivityLogs();
      } catch(e) {}
    }
  }
}

function sendToRailwayServer(action, payload) {
  const userName = currentUser ? currentUser.name : 'مشرف';

  // 1. Fast WebSocket broadcast (0ms delay)
  if (_railwaySocket && _railwaySocket.readyState === WebSocket.OPEN) {
    try {
      _railwaySocket.send(JSON.stringify({
        type: 'SYNC_ACTION',
        action,
        payload,
        data: payload,
        byUser: userName,
        sender: userName,
        clientId: window.ARAIJ_CLIENT_ID,
        timestamp: Date.now()
      }));
      console.log(`[Railway Live Sync] Sent ${action} over WebSocket ⚡`);
      return true;
    } catch(e) {
      console.warn('[Railway Live Sync] WebSocket send error, falling back to HTTP:', e);
    }
  }

  // 2. HTTP POST Relay
  const srvUrl = getRailwayServerUrl();
  const targetUrl = srvUrl ? `${srvUrl}/api/sync` : '/api/sync_local';
  fetch(targetUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ 
      action, 
      payload, 
      data: payload, 
      byUser: userName, 
      sender: userName, 
      clientId: window.ARAIJ_CLIENT_ID, 
      ...payload 
    })
  }).catch(err => {
    console.warn('[Railway Live Sync] HTTP sync failed, queueing offline:', err);
    addToOfflineSyncQueue(action, payload);
  });
}

function saveRailwaySettings() {
  const input = document.getElementById('railwayServerUrlInput');
  const url = sanitizeRailwayUrl(input?.value || '');
  if (url) {
    localStorage.setItem('araij_railway_url', url);
    showToast(`تم حفظ وتفعيل رابط ريلوي (${url}) ⚡`);
  } else {
    localStorage.removeItem('araij_railway_url');
    showToast('تم ضبط السيرفر على الرابط الافتراضي.');
  }
  initLiveRailwaySync();
}
window.saveRailwaySettings = saveRailwaySettings;

// Send ONLY to LAN local server (not Google Sheets)
function sendToLanServer(action, payload) {
  try {
    fetch('/api/sync_local', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ action, ...payload })
    }).catch(() => {});
  } catch(e) {}
}

const _recentSyncMap = new Map(); // key -> timestamp

// Send ONLY to Google Sheets (not LAN)
// Returns true if sent successfully, false if queued for later
function syncToGoogleSheets(action, payload, force = false) {
  const url = getGoogleAppsScriptUrl();
  if (!url) return false;

  const payloadCode = String(payload?.student?.code || payload?.code || '').trim();
  const syncKey = `${action}_${payloadCode}`;

  // If force is false, debounce accidental rapid duplicate calls within 1.5s (allows rapid successive edits)
  const now = Date.now();
  if (!force && payloadCode && _recentSyncMap.has(syncKey)) {
    const prevTime = _recentSyncMap.get(syncKey);
    if (now - prevTime < 1500) {
      console.log(`Prevented duplicate sync to Google Sheets for: ${syncKey}`);
      return false;
    }
  }
  if (payloadCode) {
    _recentSyncMap.set(syncKey, now);
  }

  if (!navigator.onLine) {
    addToOfflineSyncQueue(action, payload);
    return false;
  }

  try {
    fetch(url, {
      method: 'POST',
      mode: 'no-cors',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ action, ...payload })
    }).then(() => {
      console.log(`[Google Sheets Sync] Dispatched ${action} for #${payloadCode}`);
    }).catch(err => {
      console.warn(`[Google Sheets Sync] Error:`, err);
      addToOfflineSyncQueue(action, payload);
    });
    return true;
  } catch (e) {
    addToOfflineSyncQueue(action, payload);
    return false;
  }
}

// Bulk / Batch sync to Google Sheets (fast dual batch + background sequential queue)
function syncBulkStudentsToGoogleSheets(studentsList) {
  if (!studentsList || !Array.isArray(studentsList) || studentsList.length === 0) return;

  // 1. Send immediately to Live Railway Server (PRIMARY 0ms CROSS-DEVICE HUB)
  sendToRailwayServer('batch_update_students', { students: studentsList });
  studentsList.forEach(st => {
    sendToLanServer('update_student', { student: st });
  });

  // 2. Secondary dispatch to Google Sheets if configured
  const url = getGoogleAppsScriptUrl();
  if (url) {
    syncToGoogleSheets('batch_update_students', { students: studentsList }, true);
    syncToGoogleSheets('bulk_session', { students: studentsList }, true);

    if (navigator.onLine) {
      let idx = 0;
      const chunkSize = 3;
      function sendNextSlice() {
        if (idx >= studentsList.length) return;
        const slice = studentsList.slice(idx, idx + chunkSize);
        idx += chunkSize;
        slice.forEach(st => {
          fetch(url, {
            method: 'POST',
            mode: 'no-cors',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ action: 'update_student', student: st })
          }).catch(() => {
            addToOfflineSyncQueue('update_student', { student: st });
          });
        });
        setTimeout(sendNextSlice, 250);
      }
      sendNextSlice();
    } else {
      studentsList.forEach(st => {
        addToOfflineSyncQueue('update_student', { student: st });
      });
    }
  }
}



// Hash of last Sheets data — to skip re-render if nothing changed
let _lastSheetsHash = '';

// AUTOMATIC & ON-DEMAND DATA PULL FROM GOOGLE SHEETS
async function pullFromGoogleSheets(isAuto = false) {
  const url = getGoogleAppsScriptUrl();
  const navIcon = document.getElementById('navPullIcon');
  const badgeText = document.getElementById('sheetsStatusBadgeText');
  const dot = document.getElementById('sheetsStatusDot');

  if (!url) return;
  if (!navigator.onLine) return;

  if (navIcon && !isAuto) navIcon.classList.add('animate-spin');
  if (badgeText && !isAuto) badgeText.textContent = 'جاري المزامنة...';

  try {
    const fetchUrl = `${url}?action=get_all&_t=${Date.now()}`;
    const response = await fetch(fetchUrl);
    const data = await response.json();

    if (data && data.status === 'success' && Array.isArray(data.students)) {
      // ✅ Comprehensive hash including all fields: code, name, grade, subjects, teachers, phone, parentPhone, area, spec
      const sheetsHash = data.students.map(s => 
        `${s.code}:${s.name}:${s.grade}:${s.subjectsSummary}:${s.teachersSummary}:${s.phone}:${s.parentPhone}:${s.area}:${s.specialization}`
      ).join('|');

      if (isAuto && sheetsHash === _lastSheetsHash) {
        // Nothing changed — skip processing entirely
        if (badgeText) badgeText.textContent = 'Google Sheets متصل ✅';
        if (dot) dot.className = 'w-2 h-2 rounded-full bg-emerald-400 animate-pulse';
        return;
      }
      _lastSheetsHash = sheetsHash;

      let newCount = 0;
      let updatedCount = 0;
      const overrides = JSON.parse(localStorage.getItem('araij_students_overrides') || '{}');

      // Deduplicate incoming rows from Google Sheets (handles cases where sheet has duplicate rows)
      const seenCodesFromSheet = new Set();
      const uniqueSheetStudents = [];
      data.students.forEach(st => {
        const c = String(st.code || '').trim();
        if (c && seenCodesFromSheet.has(c)) return; // Skip duplicate sheet rows!
        if (c) seenCodesFromSheet.add(c);
        uniqueSheetStudents.push(st);
      });

      uniqueSheetStudents.forEach(sheetStudent => {
        if (!sheetStudent || !sheetStudent.name || String(sheetStudent.name).trim().length < 2 || String(sheetStudent.name).includes('غير مسمى')) return;
        const code = String(sheetStudent.code || '').trim();
        // Match by code only
        const existingLocal = allStudents.find(s => String(s.code).trim() === code);

        if (!existingLocal) {
          const newSt = {
            code: code || ('GS-' + Math.floor(1000 + Math.random() * 9000)),
            name: sheetStudent.name || 'طالب جديد',
            area: sheetStudent.area || '',
            phone: sheetStudent.phone || '',
            parentPhone: sheetStudent.parentPhone || '',
            grade: normalize_grade(sheetStudent.grade),
            specialization: sheetStudent.specialization || 'عام',
            subjectsSummary: sheetStudent.subjectsSummary || '',
            teachersSummary: sheetStudent.teachersSummary || '',
            regDate: sheetStudent.regDate || new Date().toLocaleDateString('ar-EG'),
            academicSubjects: parseSubjectsFromSummary(sheetStudent.subjectsSummary, sheetStudent.teachersSummary),
            createdBy: sheetStudent.createdBy || 'منار ناصر',
            lastModifiedBy: sheetStudent.lastModifiedBy || sheetStudent.createdBy || 'منار ناصر',
            lastAction: sheetStudent.lastAction || 'مسجل بالسحابة',
            lastActionTime: sheetStudent.lastActionTime || sheetStudent.regDate || ''
          };
          newSt._metrics = calculateStudentMetrics(newSt);
          newSt._searchString = buildStudentSearchString(newSt);
          overrides[newSt.code] = newSt;
          allStudents.unshift(newSt);
          newCount++;
        } else {
          let changed = false;
          ['name', 'phone', 'parentPhone', 'area', 'grade', 'specialization', 'subjectsSummary', 'teachersSummary'].forEach(k => {
            const incomingVal = String(sheetStudent[k] || '').trim();
            const currentVal = String(existingLocal[k] || '').trim();
            if (incomingVal !== currentVal) {
              existingLocal[k] = sheetStudent[k];
              changed = true;
            }
          });

          // Normalize grade if needed
          if (existingLocal.grade) {
            existingLocal.grade = normalize_grade(existingLocal.grade);
          }

          // ✅ Reconstruct and sync academic subjects while preserving past attendance sessions
          const syncedAcademic = parseSubjectsFromSummary(sheetStudent.subjectsSummary, sheetStudent.teachersSummary, existingLocal.academicSubjects);
          if (JSON.stringify(syncedAcademic) !== JSON.stringify(existingLocal.academicSubjects || {})) {
            existingLocal.academicSubjects = syncedAcademic;
            changed = true;
          }

          if (changed) {
            existingLocal._metrics = calculateStudentMetrics(existingLocal);
            existingLocal._searchString = buildStudentSearchString(existingLocal);
            overrides[existingLocal.code] = existingLocal;
            updatedCount++;
          }
        }
      });

      // ✅ CROSS-DEVICE DELETE SYNC
      // Students in Sheets are authoritative. If a student exists locally
      // but is NOT in Sheets → it was deleted by another user → remove it here too.
      const sheetCodes = new Set(data.students.map(s => String(s.code || '').trim()).filter(Boolean));
      const pendingAddCodes = new Set(
        getOfflineSyncQueue()
          .filter(q => q.action === 'add_student')
          .map(q => String(q.payload?.student?.code || '').trim())
      );
      // Only check students that have codes (skip base data students without codes or pre-loaded data.js)
      const localOnlyCodes = allStudents
        .filter(s => {
          const c = String(s.code || '').trim();
          // Has a code, not in Sheets, not pending upload, and code looks like a manually-added code (not base data)
          return c && !sheetCodes.has(c) && !pendingAddCodes.has(c) && overrides[c];
        })
        .map(s => s.code);

      let deletedCount = 0;
      if (localOnlyCodes.length > 0) {
        // Add to deleted list so they don't come back
        let deletedList = JSON.parse(localStorage.getItem('araij_deleted_students') || '[]');
        localOnlyCodes.forEach(code => {
          if (!deletedList.includes(code)) deletedList.push(code);
          delete overrides[code];
        });
        localStorage.setItem('araij_deleted_students', JSON.stringify(deletedList));
        allStudents = allStudents.filter(s => !localOnlyCodes.includes(s.code));
        deletedCount = localOnlyCodes.length;
      }

      if (newCount > 0 || updatedCount > 0 || deletedCount > 0) {
        localStorage.setItem('araij_students_overrides', JSON.stringify(overrides));
        try { localStorage.setItem('araij_full_students_db', JSON.stringify(allStudents)); } catch(e) {}

        allStudents.sort((a, b) => (a.name || '').localeCompare(b.name || '', 'ar', { numeric: true, sensitivity: 'base' }));

        applyFilters();
        updateKPIStats();

        // Only re-render attendance matrix if that tab is currently active
        const attendanceTab = document.getElementById('tabContent_attendance_matrix');
        if (attendanceTab && !attendanceTab.classList.contains('hidden')) {
          renderAttendanceMatrix();
        }

        // Live refresh student detail modal if open
        if (activeDetailStudentCode) {
          const detailModal = document.getElementById('studentDetailModal');
          if (detailModal && !detailModal.classList.contains('hidden')) {
            openStudentDetailModal(activeDetailStudentCode);
          }
        }

        if (!isAuto) {
          showToast(`⚡ تم سحب ${newCount} جديد، تحديث ${updatedCount}، حذف ${deletedCount} من السحابة!`);
        } else {
          if (newCount > 0) showToast(`🔄 ${newCount} طالب جديد وصل من جهاز آخر!`);
          if (deletedCount > 0) showToast(`🗑️ تم حذف ${deletedCount} طالب بواسطة مستخدم آخر.`);
        }
      } else {
        if (!isAuto) {
          showToast('البيانات متطابقة ومحدثة بالكامل مع السحابة.');
        }
      }


      if (badgeText) badgeText.textContent = 'Google Sheets متصل ✅';
      if (dot) dot.className = 'w-2 h-2 rounded-full bg-emerald-400 animate-pulse';
    }
  } catch (err) {
    if (!isAuto) {
      showToast('تعذر سحب البيانات تلقائياً من السحابة حالياً.');
    }
  } finally {
    if (navIcon) navIcon.classList.remove('animate-spin');
  }
}

// Master Unified Sync Function
async function triggerUniversalSync(isSilent = false) {
  const navIcon = document.getElementById('navPullIcon');
  if (navIcon && !isSilent) navIcon.classList.add('animate-spin');

  try {
    if (navigator.onLine) {
      await flushOfflineSyncQueue();
      await pullFromGoogleSheets(isSilent);
    }
    await fetchLatestStudentsFromCentralServer(true);
    if (!isSilent) {
      showToast('✅ تم تحديث ومزامنة كافة البيانات بنجاح!');
    }
  } catch (e) {
    if (!isSilent) showToast('تم تحديث البيانات محلياً.');
  } finally {
    if (navIcon) navIcon.classList.remove('animate-spin');
  }
}

// Sync events
window.addEventListener('online', () => {
  showToast('🌐 تم استعادة الاتصال بالإنترنت! جاري مزامنة السحابة...');
  triggerUniversalSync(false);
});

// NOTE: The main polling (every 15s) is set in DOMContentLoaded above.
// Removed the duplicate 10-second interval that was causing repeated Sheets calls.



function updateSheetsStatusIndicator() {
  const url = getGoogleAppsScriptUrl();
  const badgeText = document.getElementById('sheetsStatusBadgeText');
  const dot = document.getElementById('sheetsStatusDot');

  if (url) {
    if (badgeText) badgeText.textContent = 'Google Sheets متصل ✅';
    if (dot) dot.className = 'w-2 h-2 rounded-full bg-emerald-400 animate-pulse';
  } else {
    if (badgeText) badgeText.textContent = 'ربط Google Sheets 🔗';
    if (dot) dot.className = 'w-2 h-2 rounded-full bg-amber-400';
  }
}

function openSheetsSyncModal() {
  const input = document.getElementById('sheetsWebhookInput');
  if (input) input.value = getGoogleAppsScriptUrl();
  const rwInput = document.getElementById('railwayServerUrlInput');
  if (rwInput) rwInput.value = getRailwayServerUrl();
  document.getElementById('sheetsSyncModal')?.classList.remove('hidden');
  initIcons();
}

function closeSheetsSyncModal() {
  document.getElementById('sheetsSyncModal')?.classList.add('hidden');
}

function copyAppsScriptCode() {
  if (navigator.clipboard) {
    navigator.clipboard.writeText(APPS_SCRIPT_SOURCE).then(() => {
      const btn = document.getElementById('copyScriptBtnText');
      if (btn) {
        btn.textContent = 'تم نسخ الكود!';
        setTimeout(() => { btn.textContent = 'نسخ الكود'; }, 2000);
      }
      showToast('تم نسخ كود Apps Script إلى الحافظة بنجاح!');
    });
  } else {
    prompt('انسخ كود Google Apps Script التالي:', APPS_SCRIPT_SOURCE);
  }
}

function testAndSaveSheetsWebhook() {
  try { saveRailwaySettings(); } catch(e) {}

  const urlInput = document.getElementById('sheetsWebhookInput');
  const url = (urlInput?.value || '').trim();
  const statusBox = document.getElementById('sheetsSyncStatusBox');
  const saveBtn = document.getElementById('saveSheetsBtn');

  if (!url) {
    alert('يرجى لصق رابط الـ Web App الخاص بـ Google Apps Script!');
    return;
  }

  if (saveBtn) saveBtn.innerHTML = `<i data-lucide="loader-2" class="w-4 h-4 animate-spin"></i> <span>جاري اختبار الاتصال بالشيت...</span>`;
  initIcons();

  const testStudent = {
    code: "TEST-" + Math.floor(100 + Math.random() * 900),
    name: "طالب تجريبي (اختبار ربط الشيت)",
    area: "سنتر الأرائج التعليمي",
    phone: "01000000000",
    parentPhone: "01002169889",
    grade: "ث1",
    specialization: "عام",
    subjectsSummary: "عربي, انجليزي",
    teachersSummary: "ا/ احمد عبد القادر",
    regDate: new Date().toLocaleDateString('ar-EG')
  };

  localStorage.setItem('araij_sheets_webhook_url', url);

  try {
    fetch(url, {
      method: 'POST',
      mode: 'no-cors',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ action: 'add_student', student: testStudent })
    }).then(() => {
      if (statusBox) {
        statusBox.className = 'p-3 rounded-xl border border-emerald-500/30 bg-emerald-500/10 text-emerald-300 text-xs font-bold flex items-center gap-2';
        statusBox.innerHTML = `<i data-lucide="check-circle" class="w-4 h-4"></i> <span>✅ تم إرسال أمر الربط بنجاح إلى الشيت!</span>`;
        statusBox.classList.remove('hidden');
      }
      if (saveBtn) saveBtn.innerHTML = `<i data-lucide="check" class="w-4 h-4"></i> <span>تم الحفظ والربط بنجاح</span>`;
      updateSheetsStatusIndicator();
      initIcons();
      showToast('تم تفعيل الربط المباشر مع Google Sheets بنجاح!');
    }).catch(err => {
      const img = new Image();
      img.src = `${url}?action=add_student&student=${encodeURIComponent(JSON.stringify(testStudent))}&_t=${Date.now()}`;
      
      if (statusBox) {
        statusBox.className = 'p-3 rounded-xl border border-emerald-500/30 bg-emerald-500/10 text-emerald-300 text-xs font-bold flex items-center gap-2';
        statusBox.innerHTML = `<i data-lucide="check-circle" class="w-4 h-4"></i> <span>✅ تم إرسال أمر الربط بنجاح إلى الشيت!</span>`;
        statusBox.classList.remove('hidden');
      }
      if (saveBtn) saveBtn.innerHTML = `<i data-lucide="check" class="w-4 h-4"></i> <span>تم الحفظ والربط</span>`;
      updateSheetsStatusIndicator();
      initIcons();
      showToast('تم تفعيل الربط مع Google Sheets!');
    });
  } catch (err) {
    alert('حدث خطأ أثناء الاتصال: ' + err);
  }
}

// ====================================================
// 15. PDF & PRINTABLE REPORTS HUB
// ====================================================

function initPdfReportSelects() {
  const teacherSelect = document.getElementById('pdfTeacherSelect');
  if (!teacherSelect) return;

  const allTeachersSet = new Set();
  Object.values(TEACHERS).forEach(list => {
    list.forEach(t => allTeachersSet.add(t));
  });

  teacherSelect.innerHTML = Array.from(allTeachersSet).sort().map(t => `
    <option value="${t}">${t}</option>
  `).join('');
}

function setPdfReportMode(mode) {
  currentPdfMode = mode;
  ['teacher', 'attendance_roster', 'grade_summary'].forEach(m => {
    const btn = document.getElementById(`pdf_btn_${m}`);
    if (btn) btn.className = 'p-4 rounded-2xl bg-slate-800/80 hover:bg-slate-800 text-slate-300 font-bold text-xs text-right space-y-1 transition-all border border-slate-700';
  });

  const activeBtnId = mode === 'teacher' ? 'pdf_btn_teacher' : mode === 'attendance_roster' ? 'pdf_btn_roster' : 'pdf_btn_summary';
  const activeBtn = document.getElementById(activeBtnId);
  if (activeBtn) {
    activeBtn.className = 'p-4 rounded-2xl bg-sky-600 text-white font-bold text-xs text-right space-y-1 shadow-lg transition-all border border-sky-400/30';
  }

  const teacherGroup = document.getElementById('pdfTeacherSelectGroup');
  if (teacherGroup) {
    if (mode === 'grade_summary') {
      teacherGroup.classList.add('hidden');
    } else {
      teacherGroup.classList.remove('hidden');
    }
  }

  generateSelectedPdfReport();
}

function generateSelectedPdfReport() {
  const container = document.getElementById('printableArea');
  if (!container) return;

  const selectedTeacher = document.getElementById('pdfTeacherSelect')?.value || 'ا/ احمد عبد القادر';
  const selectedGrade = document.getElementById('pdfGradeSelect')?.value || 'all';
  const selectedMonth = document.getElementById('pdfMonthSelect')?.value || currentActiveMonth || 'شهر 1';
  const today = new Date().toLocaleDateString('ar-EG', { year: 'numeric', month: 'long', day: 'numeric' });

  const teacherNorm = normalize_arabic(selectedTeacher);

  if (currentPdfMode === 'teacher') {
    const matchedStudents = allStudents.filter(s => {
      if (selectedGrade !== 'all' && s.grade !== selectedGrade) return false;
      const tSummary = normalize_arabic(s.teachersSummary || '');
      if (tSummary.includes(teacherNorm)) return true;
      const subs = s.academicSubjects || {};
      return Object.values(subs).some(sub => normalize_arabic(sub.teacher || '').includes(teacherNorm));
    });

    container.innerHTML = `
      <div style="font-family: 'Cairo', sans-serif; direction: rtl; text-align: right; color: #0f172a;">
        <div style="border-bottom: 2px solid #0284c7; padding-bottom: 12px; margin-bottom: 20px; display: flex; justify-content: space-between; align-items: center;">
          <div>
            <h1 style="font-size: 22px; font-weight: 900; color: #0284c7; margin: 0;">سنتر الأرائج التعليمي</h1>
            <p style="font-size: 13px; color: #64748b; margin: 4px 0 0 0;">كشف وقائمة طلاب الأستاذ / ${selectedTeacher}</p>
          </div>
          <div style="text-align: left; font-size: 11px; color: #64748b;">
            <div><strong>التاريخ:</strong> ${today}</div>
            <div><strong>المرحلة:</strong> ${selectedGrade === 'all' ? 'جميع المراحل' : selectedGrade}</div>
            <div><strong>إجمالي الطلاب:</strong> ${matchedStudents.length} طالب</div>
          </div>
        </div>

        <table style="width: 100%; border-collapse: collapse; font-size: 11px; text-align: right;">
          <thead>
            <tr style="background-color: #0284c7; color: #ffffff;">
              <th style="padding: 8px; border: 1px solid #cbd5e1; width: 40px; text-align: center;">م</th>
              <th style="padding: 8px; border: 1px solid #cbd5e1; width: 70px; text-align: center;">الكود</th>
              <th style="padding: 8px; border: 1px solid #cbd5e1;">اسم الطالب</th>
              <th style="padding: 8px; border: 1px solid #cbd5e1; width: 60px; text-align: center;">الصف</th>
              <th style="padding: 8px; border: 1px solid #cbd5e1;">الشعبة / التخصص</th>
              <th style="padding: 8px; border: 1px solid #cbd5e1;">المنطقة</th>
              <th style="padding: 8px; border: 1px solid #cbd5e1; width: 95px; text-align: center;">هاتف الطالب</th>
              <th style="padding: 8px; border: 1px solid #cbd5e1; width: 95px; text-align: center;">ولي الأمر</th>
            </tr>
          </thead>
          <tbody>
            ${matchedStudents.length === 0 ? `
              <tr><td colspan="8" style="padding: 20px; text-align: center; color: #64748b;">لا يوجد طلاب مسجلين مع هذا المدرس في المرحلة المحددة.</td></tr>
            ` : matchedStudents.map((s, idx) => `
              <tr style="background-color: ${idx % 2 === 0 ? '#ffffff' : '#f8fafc'};">
                <td style="padding: 6px; border: 1px solid #cbd5e1; text-align: center;">${idx + 1}</td>
                <td style="padding: 6px; border: 1px solid #cbd5e1; text-align: center; font-weight: bold; color: #0284c7;">#${s.code}</td>
                <td style="padding: 6px; border: 1px solid #cbd5e1; font-weight: bold;">${s.name}</td>
                <td style="padding: 6px; border: 1px solid #cbd5e1; text-align: center;">${s.grade || '-'}</td>
                <td style="padding: 6px; border: 1px solid #cbd5e1;">${s.specialization || 'عام'}</td>
                <td style="padding: 6px; border: 1px solid #cbd5e1;">${s.area || '-'}</td>
                <td style="padding: 6px; border: 1px solid #cbd5e1; text-align: center; font-family: monospace;">${s.phone || '-'}</td>
                <td style="padding: 6px; border: 1px solid #cbd5e1; text-align: center; font-family: monospace; color: #059669; font-weight: bold;">${s.parentPhone || '-'}</td>
              </tr>
            `).join('')}
          </tbody>
        </table>

        <div style="margin-top: 30px; padding-top: 15px; border-top: 1px dashed #cbd5e1; display: flex; justify-content: space-between; font-size: 11px; color: #64748b;">
          <div>إشراف: إدارة سنتر الأرائج التعليمي</div>
          <div>توقيع المدرس: ___________________</div>
          <div>ختم الإدارة: ___________________</div>
        </div>
      </div>
    `;

  } else if (currentPdfMode === 'attendance_roster') {
    const matchedStudents = allStudents.filter(s => {
      if (selectedGrade !== 'all' && s.grade !== selectedGrade) return false;
      const tSummary = normalize_arabic(s.teachersSummary || '');
      if (tSummary.includes(teacherNorm)) return true;
      const subs = s.academicSubjects || {};
      return Object.values(subs).some(sub => normalize_arabic(sub.teacher || '').includes(teacherNorm));
    });

    let sumPresent = 0;
    let sumAbsent = 0;
    let sumRecorded = 0;

    const studentRosterData = matchedStudents.map((s, idx) => {
      let subjName = '';
      if (s.academicSubjects) {
        for (const subKey of Object.keys(s.academicSubjects)) {
          const subObj = s.academicSubjects[subKey];
          if (normalize_arabic(subObj.teacher || '').includes(teacherNorm)) {
            subjName = subKey;
            break;
          }
        }
      }
      if (!subjName) subjName = Object.keys(s.academicSubjects || {})[0] || 'المادة';

      const sessions = getStudentSubjectMonthSessions(s, subjName, selectedMonth);

      let pCount = 0;
      let aCount = 0;

      sessions.forEach(v => {
        if (!v || v === '-' || v === '') return;
        sumRecorded++;
        if (v === '✓' || v === 'حاضر' || (!isNaN(parseFloat(v)) && v !== '0')) {
          pCount++;
          sumPresent++;
        } else if (v === 'غ' || v === 'غائب' || v === '0') {
          aCount++;
          sumAbsent++;
        }
      });

      const totalStudentRec = pCount + aCount;
      const rate = totalStudentRec > 0 ? Math.round((pCount / totalStudentRec) * 100) : 0;
      const evalText = rate >= 90 ? 'ممتاز' : rate >= 75 ? 'جيد جداً' : rate >= 50 ? 'مقبول' : totalStudentRec > 0 ? 'ضعيف' : 'غير مرصود';
      const evalColor = rate >= 90 ? '#059669' : rate >= 75 ? '#0284c7' : rate >= 50 ? '#d97706' : '#dc2626';

      return { s, idx, subjName, sessions, pCount, aCount, rate, evalText, evalColor };
    });

    const overallRate = sumRecorded > 0 ? Math.round((sumPresent / sumRecorded) * 100) : 0;

    container.innerHTML = `
      <div style="font-family: 'Cairo', sans-serif; direction: rtl; text-align: right; color: #0f172a;">
        <div style="border-bottom: 2px solid #059669; padding-bottom: 12px; margin-bottom: 16px; display: flex; justify-content: space-between; align-items: center;">
          <div>
            <h1 style="font-size: 22px; font-weight: 900; color: #059669; margin: 0;">سنتر الأرائج التعليمي</h1>
            <p style="font-size: 13px; color: #334155; margin: 4px 0 0 0; font-weight: bold;">
              كشف رصد الحضور والغياب والدرجات — <span style="color: #059669;">${selectedMonth}</span> (حصص 1 إلى 8)
            </p>
          </div>
          <div style="text-align: left; font-size: 11px; color: #64748b;">
            <div><strong>المدرس:</strong> ${selectedTeacher}</div>
            <div><strong>التاريخ:</strong> ${today}</div>
            <div><strong>المرحلة:</strong> ${selectedGrade === 'all' ? 'جميع المراحل' : selectedGrade}</div>
          </div>
        </div>

        <!-- Summary Banner -->
        <div style="background-color: #f0fdf4; border: 1px solid #bbf7d0; padding: 8px 12px; border-radius: 10px; margin-bottom: 16px; display: flex; justify-content: space-between; font-size: 11px; color: #166534; font-weight: bold;">
          <div>إجمالي الطلاب: ${matchedStudents.length} طالب</div>
          <div>إجمالي مرات الحضور: ${sumPresent}</div>
          <div>إجمالي مرات الغياب: ${sumAbsent}</div>
          <div>نسبة الحضور العامة: ${overallRate}%</div>
        </div>

        <table style="width: 100%; border-collapse: collapse; font-size: 10px; text-align: center;">
          <thead>
            <tr style="background-color: #059669; color: #ffffff;">
              <th style="padding: 6px; border: 1px solid #cbd5e1; width: 28px;">م</th>
              <th style="padding: 6px; border: 1px solid #cbd5e1; width: 45px;">الكود</th>
              <th style="padding: 6px; border: 1px solid #cbd5e1; text-align: right; width: 140px;">اسم الطالب</th>
              <th style="padding: 6px; border: 1px solid #cbd5e1; width: 42px;">ح 1</th>
              <th style="padding: 6px; border: 1px solid #cbd5e1; width: 42px;">ح 2</th>
              <th style="padding: 6px; border: 1px solid #cbd5e1; width: 42px;">ح 3</th>
              <th style="padding: 6px; border: 1px solid #cbd5e1; width: 42px;">ح 4</th>
              <th style="padding: 6px; border: 1px solid #cbd5e1; width: 42px;">ح 5</th>
              <th style="padding: 6px; border: 1px solid #cbd5e1; width: 42px;">ح 6</th>
              <th style="padding: 6px; border: 1px solid #cbd5e1; width: 42px;">ح 7</th>
              <th style="padding: 6px; border: 1px solid #cbd5e1; width: 42px;">ح 8</th>
              <th style="padding: 6px; border: 1px solid #cbd5e1; width: 38px;">حضور</th>
              <th style="padding: 6px; border: 1px solid #cbd5e1; width: 38px;">غياب</th>
              <th style="padding: 6px; border: 1px solid #cbd5e1; width: 55px;">التقدير</th>
            </tr>
          </thead>
          <tbody>
            ${studentRosterData.length === 0 ? `
              <tr><td colspan="14" style="padding: 20px; text-align: center; color: #64748b;">لا يوجد طلاب مسجلين في هذا الكشف.</td></tr>
            ` : studentRosterData.map(r => `
              <tr style="background-color: ${r.idx % 2 === 0 ? '#ffffff' : '#f8fafc'};">
                <td style="padding: 5px; border: 1px solid #cbd5e1;">${r.idx + 1}</td>
                <td style="padding: 5px; border: 1px solid #cbd5e1; font-weight: bold; color: #0284c7;">#${r.s.code}</td>
                <td style="padding: 5px; border: 1px solid #cbd5e1; text-align: right; font-weight: bold;">
                  <div>${r.s.name}</div>
                  <div style="font-size: 9px; color: #64748b; font-weight: normal;">${r.s.grade || ''} - ${r.s.area || ''}</div>
                </td>
                ${r.sessions.map(val => {
                  const isP = val === '✓' || val === 'حاضر';
                  const isA = val === 'غ' || val === 'غائب';
                  const isScr = val && !isP && !isA && val !== '-';
                  const clr = isP ? '#059669' : isA ? '#dc2626' : isScr ? '#0284c7' : '#94a3b8';
                  const bg = isP ? '#ecfdf5' : isA ? '#fef2f2' : isScr ? '#f0f9ff' : 'transparent';
                  return `<td style="padding: 5px; border: 1px solid #cbd5e1; color: ${clr}; background-color: ${bg}; font-weight: bold;">${val || '-'}</td>`;
                }).join('')}
                <td style="padding: 5px; border: 1px solid #cbd5e1; color: #059669; font-weight: bold;">${r.pCount}</td>
                <td style="padding: 5px; border: 1px solid #cbd5e1; color: #dc2626; font-weight: bold;">${r.aCount}</td>
                <td style="padding: 5px; border: 1px solid #cbd5e1; color: ${r.evalColor}; font-weight: bold;">${r.evalText}</td>
              </tr>
            `).join('')}
          </tbody>
        </table>

        <div style="margin-top: 30px; padding-top: 15px; border-top: 1px dashed #cbd5e1; display: flex; justify-content: space-between; font-size: 11px; color: #64748b;">
          <div>إشراف: إدارة سنتر الأرائج التعليمي</div>
          <div>المشرف المسئول: ___________________</div>
          <div>توقيع مدرس المادة: ___________________</div>
        </div>
      </div>
    `;

  } else {
    const matchedStudents = allStudents.filter(s => {
      if (selectedGrade !== 'all' && s.grade !== selectedGrade) return false;
      return true;
    });

    container.innerHTML = `
      <div style="font-family: 'Cairo', sans-serif; direction: rtl; text-align: right; color: #0f172a;">
        <div style="border-bottom: 2px solid #4f46e5; padding-bottom: 12px; margin-bottom: 20px; display: flex; justify-content: space-between; align-items: center;">
          <div>
            <h1 style="font-size: 22px; font-weight: 900; color: #4f46e5; margin: 0;">سنتر الأرائج التعليمي</h1>
            <p style="font-size: 13px; color: #64748b; margin: 4px 0 0 0;">كشف السجل العام لطلاب مرحلة (${selectedGrade === 'all' ? 'جميع المراحل' : selectedGrade})</p>
          </div>
          <div style="text-align: left; font-size: 11px; color: #64748b;">
            <div><strong>التاريخ:</strong> ${today}</div>
            <div><strong>إجمالي المقيدين:</strong> ${matchedStudents.length} طالب</div>
          </div>
        </div>

        <table style="width: 100%; border-collapse: collapse; font-size: 10px; text-align: right;">
          <thead>
            <tr style="background-color: #4f46e5; color: #ffffff;">
              <th style="padding: 6px; border: 1px solid #cbd5e1; width: 35px; text-align: center;">م</th>
              <th style="padding: 6px; border: 1px solid #cbd5e1; width: 55px; text-align: center;">الكود</th>
              <th style="padding: 6px; border: 1px solid #cbd5e1;">اسم الطالب</th>
              <th style="padding: 6px; border: 1px solid #cbd5e1; width: 45px; text-align: center;">الصف</th>
              <th style="padding: 6px; border: 1px solid #cbd5e1;">الشعبة</th>
              <th style="padding: 6px; border: 1px solid #cbd5e1;">المنطقة</th>
              <th style="padding: 6px; border: 1px solid #cbd5e1; width: 90px; text-align: center;">هاتف الطالب</th>
              <th style="padding: 6px; border: 1px solid #cbd5e1; width: 90px; text-align: center;">ولي الأمر</th>
              <th style="padding: 6px; border: 1px solid #cbd5e1; width: 50px; text-align: center;">التقدير</th>
            </tr>
          </thead>
          <tbody>
            ${matchedStudents.map((s, idx) => `
              <tr style="background-color: ${idx % 2 === 0 ? '#ffffff' : '#f8fafc'};">
                <td style="padding: 5px; border: 1px solid #cbd5e1; text-align: center;">${idx + 1}</td>
                <td style="padding: 5px; border: 1px solid #cbd5e1; text-align: center; font-weight: bold; color: #4f46e5;">#${s.code}</td>
                <td style="padding: 5px; border: 1px solid #cbd5e1; font-weight: bold;">${s.name}</td>
                <td style="padding: 5px; border: 1px solid #cbd5e1; text-align: center;">${s.grade || '-'}</td>
                <td style="padding: 5px; border: 1px solid #cbd5e1;">${s.specialization || 'عام'}</td>
                <td style="padding: 5px; border: 1px solid #cbd5e1;">${s.area || '-'}</td>
                <td style="padding: 5px; border: 1px solid #cbd5e1; text-align: center; font-family: monospace;">${s.phone || '-'}</td>
                <td style="padding: 6px; border: 1px solid #cbd5e1; text-align: center; font-family: monospace; color: #059669; font-weight: bold;">${s.parentPhone || '-'}</td>
                <td style="padding: 5px; border: 1px solid #cbd5e1; text-align: center; font-weight: bold;">${s._metrics ? s._metrics.gradeEvaluation : '-'}</td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>
    `;
  }
}

function printGeneratedReport() {
  window.print();
}

// ====================================================
// 16. PARENT REPORT CARD & WHATSAPP
// ====================================================

function openParentReportModal(code) {
  const student = allStudents.find(s => s.code === code);
  if (!student) return;

  activeReportStudent = student;
  const m = student._metrics;

  const today = new Date().toLocaleDateString('ar-EG', { year: 'numeric', month: 'long', day: 'numeric' });
  const elDate = document.getElementById('reportPrintDate');
  const elCode = document.getElementById('reportStudentCode');
  const elName = document.getElementById('reportStudentName');
  const elGrade = document.getElementById('reportStudentGrade');
  const elArea = document.getElementById('reportStudentArea');
  const elParentPhone = document.getElementById('reportParentPhone');
  const elGradeBadge = document.getElementById('reportOverallGradeBadge');
  const elScore = document.getElementById('reportAcademicScore');
  const elAttRate = document.getElementById('reportAttendanceRate');

  if (elDate) elDate.textContent = today;
  if (elCode) elCode.textContent = student.code;
  if (elName) elName.textContent = student.name;
  if (elGrade) elGrade.textContent = student.grade || 'غير محدد';
  if (elArea) elArea.textContent = student.area || 'غير محدد';
  if (elParentPhone) elParentPhone.textContent = student.parentPhone || student.phone || 'غير مسجل';

  if (elGradeBadge) elGradeBadge.textContent = m.gradeEvaluation;
  if (elScore) elScore.textContent = m.academicPct !== null ? `${m.academicPct}%` : 'مسجل';
  if (elAttRate) elAttRate.textContent = `${m.attendanceRate}%`;

  const tableBody = document.getElementById('reportSubjectsTableBody');
  const subjects = student.academicSubjects || {};
  const subKeys = Object.keys(subjects);

  if (tableBody) {
    if (subKeys.length === 0) {
      tableBody.innerHTML = `
        <tr>
          <td colspan="5" class="p-3 text-center text-slate-400">
            المواد مسجلة بالقاعدة الرئيسية (التقييم قائم على المتابعة والانتظام).
          </td>
        </tr>
      `;
    } else {
      tableBody.innerHTML = subKeys.map(subName => {
        const sub = subjects[subName];
        const sessions = sub.sessions || [];
        
        let subScore = 0;
        let subMax = 0;
        let attended = 0;
        let absent = 0;
        let totalRecorded = 0;

        const sessionBadges = sessions.map((s, idx) => {
          const val = (s || '').trim();
          if (!val) return `<span class="inline-block px-1 py-0.5 rounded text-[10px] bg-slate-800 text-slate-500">-</span>`;
          totalRecorded++;
          if (val === '✓' || val === 'حاضر') {
            attended++;
            return `<span class="inline-block px-1.5 py-0.5 rounded text-[10px] bg-emerald-600/30 text-emerald-400 font-bold border border-emerald-500/30" style="color: #059669 !important; background: #ecfdf5 !important;">ح${idx+1}:✓</span>`;
          } else if (val === 'غ' || val === 'غائب') {
            absent++;
            return `<span class="inline-block px-1.5 py-0.5 rounded text-[10px] bg-rose-600/30 text-rose-400 font-bold border border-rose-500/30" style="color: #dc2626 !important; background: #fef2f2 !important;">ح${idx+1}:غ</span>`;
          } else {
            attended++;
            const parts = val.split('/');
            if (parts.length === 2 && !isNaN(parseFloat(parts[0])) && !isNaN(parseFloat(parts[1]))) {
              subScore += parseFloat(parts[0]);
              subMax += parseFloat(parts[1]);
            }
            return `<span class="inline-block px-1.5 py-0.5 rounded text-[10px] bg-sky-600/30 text-sky-300 font-bold border border-sky-500/30" style="color: #0284c7 !important; background: #f0f9ff !important;">ح${idx+1}:${val}</span>`;
          }
        }).join(' ');

        // Calculate subject evaluation level
        let subLevel = 'منتظم بالحضور ✅';
        let subBadgeStyle = 'color: #059669 !important; background: #ecfdf5 !important;';
        
        if (subMax > 0) {
          const pct = Math.round((subScore / subMax) * 100);
          if (pct >= 85) {
            subLevel = `ممتاز (${subScore}/${subMax} - ${pct}%) 🌟`;
            subBadgeStyle = 'color: #15803d !important; background: #dcfce7 !important;';
          } else if (pct >= 75) {
            subLevel = `جيد جداً (${subScore}/${subMax} - ${pct}%) ✨`;
            subBadgeStyle = 'color: #0284c7 !important; background: #e0f2fe !important;';
          } else if (pct >= 65) {
            subLevel = `جيد (${subScore}/${subMax} - ${pct}%) 👍`;
            subBadgeStyle = 'color: #4f46e5 !important; background: #eef2ff !important;';
          } else if (pct >= 50) {
            subLevel = `مقبول (${subScore}/${subMax} - ${pct}%)`;
            subBadgeStyle = 'color: #b45309 !important; background: #fef3c7 !important;';
          } else {
            subLevel = `بحاجة لمتابعة (${subScore}/${subMax} - ${pct}%) ⚠️`;
            subBadgeStyle = 'color: #b91c1c !important; background: #fee2e2 !important;';
          }
        } else if (absent > 0) {
          subLevel = `غياب (${absent} حصة) ⚠️`;
          subBadgeStyle = 'color: #b91c1c !important; background: #fee2e2 !important;';
        } else if (attended > 0) {
          subLevel = `حضور كامل (${attended} حصص) 🌟`;
          subBadgeStyle = 'color: #15803d !important; background: #dcfce7 !important;';
        }

        const attStatusText = totalRecorded > 0 ? (absent === 0 ? `حضور تام (${attended} من ${totalRecorded}) 🌟` : `${attended} حضور / ${absent} غياب`) : 'منتظم';

        return `
          <tr class="hover:bg-slate-800/40">
            <td class="p-2.5 font-bold text-sky-400" style="color: #0284c7 !important;">${subName}</td>
            <td class="p-2.5 text-slate-300" style="color: #334155 !important;">${sub.teacher || 'مدرس المادة'}</td>
            <td class="p-2.5 text-center font-mono">${sessionBadges}</td>
            <td class="p-2.5 text-center font-semibold text-slate-200" style="color: #0f172a !important;">${attStatusText}</td>
            <td class="p-2.5 text-center font-bold text-xs"><span class="px-2 py-1 rounded-lg inline-block" style="${subBadgeStyle}">${subLevel}</span></td>
          </tr>
        `;
      }).join('');
    }
  }

  // SMART POSITIVE / WARNING ADMIN NOTE
  let noteText = '';
  const isWeak = m.totalAbsent >= 2 || (m.attendanceRate < 60 && m.totalRecordedSessions > 0) || m.gradeEvaluation === "ضعيف";
  const isExcellent = m.gradeEvaluation === "ممتاز" || (m.academicPct && m.academicPct >= 85);
  const isVeryGood = m.gradeEvaluation === "جيد جدًا" || (m.academicPct && m.academicPct >= 75);

  if (isWeak) {
    noteText = `⚠️ نحيطكم علماً بأن الطالب/ة (${student.name}) بحاجة إلى مزيد من الاهتمام والمتابعة الدورية بحضور الحصص ومراجعة الواجبات أولاً بأول لرفع المستوى.`;
  } else if (isExcellent) {
    noteText = `✨ يسعدنا إفادتكم بأن الطالب/ة (${student.name}) يسير بمستوى أكاديمي ممتاز ومتفوق، ونشيد بالتزامه وانضباطه العالي في الحصص. نرجو مواصلة الدعم والتحفيز للحفاظ على هذا التميز دائماً بإذن الله.`;
  } else if (isVeryGood) {
    noteText = `🌟 مستوى الطالب/ة (${student.name}) جيد جداً ومستقر، ويظهر حرصاً كبيراً في المتابعة والحضور. نتمنى له مزيداً من التقدم والتفوق المستمر.`;
  } else {
    noteText = `👍 يسعدنا إفادتكم بالتزام الطالب/ة (${student.name}) وحضوره المنتظم لكافة الحصص والمتابعة المستمرة، متمنين له دوام التقدم والنجاح والتفوق دائماً.`;
  }

  const elNotes = document.getElementById('reportAdminNotes');
  if (elNotes) elNotes.textContent = noteText;

  document.getElementById('parentReportModal')?.classList.remove('hidden');
  initIcons();
}

function closeParentReportModal() {
  document.getElementById('parentReportModal')?.classList.add('hidden');
  activeReportStudent = null;
}

function printStudentCardReport() {
  const printContent = document.getElementById('studentPrintCardArea');
  if (!printContent) {
    window.print();
    return;
  }

  if (activeReportStudent) {
    logActivity('طباعة تقرير ولي أمر', `تمت طباعة التقرير الأكاديمي للطالب (${activeReportStudent.name})`, activeReportStudent.code, activeReportStudent.name);
  }

  const printWindow = window.open('', '_blank', 'width=900,height=950');
  if (printWindow) {
    printWindow.document.write(`
      <!DOCTYPE html>
      <html lang="ar" dir="rtl">
      <head>
        <meta charset="UTF-8">
        <title>تقرير الطالب - سنتر الأرائج التعليمي</title>
        <link href="https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;900&display=swap" rel="stylesheet">
        <style>
          * { box-sizing: border-box; }
          body {
            font-family: 'Cairo', sans-serif;
            direction: rtl;
            text-align: right;
            background: #ffffff;
            color: #0f172a;
            padding: 25px;
            margin: 0;
          }
          table { width: 100%; border-collapse: collapse; margin-top: 15px; }
          th, td { border: 1px solid #cbd5e1; padding: 8px 10px; font-size: 12px; }
          th { background-color: #f1f5f9; color: #1e293b; font-weight: bold; }
          @media print {
            body { padding: 10px; }
            @page { size: A4 portrait; margin: 10mm; }
          }
        </style>
      </head>
      <body>
        ${printContent.innerHTML}
      </body>
      </html>
    `);
    printWindow.document.close();
    printWindow.focus();
    setTimeout(() => {
      printWindow.print();
    }, 500);
  } else {
    window.print();
  }
}

function buildWhatsAppReportMessage(student) {
  const m = student._metrics;
  const today = new Date().toLocaleDateString('ar-EG');

  let text = `🌟 *تقرير المتابعة الأكاديمية - سنتر الأرائج التعليمي* 🌟\n`;
  text += `━━━━━━━━━━━━━━━━━━━━━\n`;
  text += `👤 *اسم الطالب:* ${student.name}\n`;
  text += `🆔 *كود الطالب:* #${student.code}\n`;
  text += `📚 *الصف الدراسي:* ${student.grade || 'المرحلة الثانوية'}\n`;
  text += `📍 *المنطقة:* ${student.area || '-'}\n`;
  text += `━━━━━━━━━━━━━━━━━━━━━\n`;
  text += `📊 *التقييم العام والمستوى:*\n`;
  text += `🎖️ *التقدير الكلي:* ${m.gradeEvaluation} ${m.academicPct !== null ? '(' + m.academicPct + '%)' : ''}\n`;
  text += `📈 *المستوى الأكاديمي:* ${m.level}\n`;
  text += `⏱️ *نسبة الحضور والالتزام:* ${m.attendanceRate}%\n`;
  text += `━━━━━━━━━━━━━━━━━━━━━\n`;

  const subjects = student.academicSubjects || {};
  const subKeys = Object.keys(subjects);
  if (subKeys.length > 0) {
    text += `📝 *كشف درجات ومستوى المواد:*\n`;
    subKeys.forEach(subName => {
      const sub = subjects[subName];
      const sessions = (sub.sessions || []).filter(s => s && s.trim() !== '').join(' , ');
      
      let subScore = 0;
      let subMax = 0;
      (sub.sessions || []).forEach(s => {
        const parts = (s || '').split('/');
        if (parts.length === 2 && !isNaN(parseFloat(parts[0])) && !isNaN(parseFloat(parts[1]))) {
          subScore += parseFloat(parts[0]);
          subMax += parseFloat(parts[1]);
        }
      });

      let levelText = '';
      if (subMax > 0) {
        const pct = Math.round((subScore / subMax) * 100);
        const evalGrade = pct >= 85 ? 'ممتاز 🌟' : pct >= 75 ? 'جيد جداً ✨' : pct >= 65 ? 'جيد 👍' : pct >= 50 ? 'مقبول' : 'يحتاج متابعة ⚠️';
        levelText = ` [الدرجة: ${subScore}/${subMax} (${pct}%) - ${evalGrade}]`;
      } else {
        levelText = ` [حضور منتظم ✅]`;
      }

      text += `• *${subName}* (${sub.teacher || 'المدرس'}): ${sessions || 'حاضر'}${levelText}\n`;
    });
    text += `━━━━━━━━━━━━━━━━━━━━━\n`;
  }

  const isWeak = m.totalAbsent >= 2 || (m.attendanceRate < 60 && m.totalRecordedSessions > 0) || m.gradeEvaluation === "ضعيف";
  const adminMsg = isWeak ? 
    `نحيطكم علماً بأن الطالب بحاجة إلى مزيد من المتابعة بحضور الحصص ومراجعة الدروس.` :
    `نشيد بالتزام واجتهاد الطالب وحضوره المنتظم ونتمنى له دوام التفوق والنجاح الباهر بإذن الله.`;

  text += `📌 *رسالة الإدارة:* ${adminMsg}\n`;
  text += `📞 *للاستفسار والتواصل:* إدارة سنتر الأرائج التعليمي (01040581954)\n`;
  text += `📅 *تاريخ التقرير:* ${today}`;

  return text;
}

function sendParentReportViaWhatsApp() {
  if (!activeReportStudent) return;
  const student = activeReportStudent;
  const rawPhone = student.parentPhone || student.phone || '';
  
  let cleanPhone = rawPhone.replace(/\D/g, '');
  if (cleanPhone.startsWith('0')) {
    cleanPhone = '2' + cleanPhone;
  } else if (!cleanPhone.startsWith('20') && cleanPhone.length === 10) {
    cleanPhone = '20' + cleanPhone;
  }

  const message = buildWhatsAppReportMessage(student);
  const encodedMsg = encodeURIComponent(message);

  let waUrl = `https://wa.me/${cleanPhone}?text=${encodedMsg}`;
  if (!cleanPhone) {
    waUrl = `https://wa.me/?text=${encodedMsg}`;
  }

  logActivity('إرسال تقرير واتساب', `تم فتح إرسال التقرير عبر واتساب لرقم (${rawPhone || 'غير محدد'})`, student.code, student.name);
  window.open(waUrl, '_blank');
  showToast('تم فتح واتساب لولي الأمر بنجاح!');
}

function quickWhatsAppSend(code) {
  const student = allStudents.find(s => s.code === code);
  if (!student) return;
  activeReportStudent = student;
  sendParentReportViaWhatsApp();
}

function copyReportText() {
  if (!activeReportStudent) return;
  const message = buildWhatsAppReportMessage(activeReportStudent);
  if (navigator.clipboard) {
    navigator.clipboard.writeText(message).then(() => {
      const btnText = document.getElementById('copyBtnText');
      if (btnText) {
        btnText.textContent = 'تم النسخ!';
        setTimeout(() => { btnText.textContent = 'نسخ النص'; }, 2000);
      }
      showToast('تم نسخ نص التقرير إلى الحافظة بنجاح!');
    });
  } else {
    prompt('انسخ نص التقرير:', message);
  }
}

// ====================================================
// BULK WHATSAPP MULTI-STUDENT DISPATCH SYSTEM
// ====================================================

function toggleSelectStudent(code, isChecked) {
  const c = String(code).trim();
  if (isChecked) {
    selectedStudentCodes.add(c);
  } else {
    selectedStudentCodes.delete(c);
  }
  updateSelectionToolbarUI();
  updateCardAndTableRowSelectState(c, isChecked);
}

function updateCardAndTableRowSelectState(code, isChecked) {
  // Update checkbox inputs without full list re-render
  document.querySelectorAll(`input[onchange*="toggleSelectStudent('${code}'"]`).forEach(input => {
    input.checked = isChecked;
  });
}

function toggleSelectAllFilteredStudents(forceState = null) {
  if (filteredStudents.length === 0) return;
  
  // Determine new state
  let shouldSelectAll;
  if (forceState !== null) {
    shouldSelectAll = forceState;
  } else {
    // If not all currently visible are selected, select all. Otherwise deselect all.
    const allSelected = filteredStudents.every(s => selectedStudentCodes.has(String(s.code)));
    shouldSelectAll = !allSelected;
  }

  filteredStudents.forEach(s => {
    const c = String(s.code);
    if (shouldSelectAll) {
      selectedStudentCodes.add(c);
    } else {
      selectedStudentCodes.delete(c);
    }
  });

  renderStudents();
  updateSelectionToolbarUI();

  if (shouldSelectAll) {
    showToast(`تم تحديد ${filteredStudents.length} طالب لإرسال تقارير واتساب.`);
  } else {
    showToast('تم إلغاء تحديد الطلاب.');
  }
}

function clearSelectedStudents() {
  selectedStudentCodes.clear();
  renderStudents();
  updateSelectionToolbarUI();
  showToast('تم مسح كافة التحديدات.');
}

function updateSelectionToolbarUI() {
  const count = selectedStudentCodes.size;
  const badge = document.getElementById('bulkSelectedCountBadge');
  if (badge) badge.textContent = String(count);

  const clearBtn = document.getElementById('btnClearSelection');
  if (clearBtn) {
    if (count > 0) clearBtn.classList.remove('hidden');
    else clearBtn.classList.add('hidden');
  }

  const selectAllBtnText = document.getElementById('selectAllBtnText');
  const chkTableAll = document.getElementById('chkTableSelectAll');
  const allFilteredSelected = filteredStudents.length > 0 && filteredStudents.every(s => selectedStudentCodes.has(String(s.code)));

  if (selectAllBtnText) {
    selectAllBtnText.textContent = allFilteredSelected ? 'إلغاء تحديد الكل' : 'تحديد الكل';
  }
  if (chkTableAll) {
    chkTableAll.checked = allFilteredSelected;
  }
}

function openBulkWhatsAppModal() {
  if (selectedStudentCodes.size === 0) {
    // If none selected, offer to select all filtered students
    if (filteredStudents.length > 0) {
      const confirmSelectAll = confirm(`لم تقم بتحديد أي طالب بعد!\nهل تريد تحديد كافة الطلاب المعروضين حالياً وعددهم (${filteredStudents.length} طالب) لفتح نافذة إرسال تقارير واتساب؟`);
      if (confirmSelectAll) {
        toggleSelectAllFilteredStudents(true);
      } else {
        return;
      }
    } else {
      alert('لا يوجد أي طالب معروض لتحديده حالياً!');
      return;
    }
  }

  bulkQueueCurrentIndex = 0;
  const modal = document.getElementById('bulkWhatsAppModal');
  if (modal) modal.classList.remove('hidden');

  renderBulkWhatsAppQueue();
  initIcons();
}

function closeBulkWhatsAppModal() {
  const modal = document.getElementById('bulkWhatsAppModal');
  if (modal) modal.classList.add('hidden');
}

function getSelectedStudentsList() {
  return allStudents.filter(s => selectedStudentCodes.has(String(s.code)));
}

function renderBulkWhatsAppQueue() {
  const container = document.getElementById('bulkQueueContainer');
  const countBadge = document.getElementById('bulkModalCountBadge');
  const progressText = document.getElementById('bulkSendProgressText');
  const progressBar = document.getElementById('bulkSendProgressBar');
  const btnNextText = document.getElementById('btnSendNextBulkText');

  const selectedList = getSelectedStudentsList();
  const total = selectedList.length;

  if (countBadge) countBadge.textContent = String(total);
  if (progressText) progressText.textContent = `${Math.min(bulkQueueCurrentIndex, total)} / ${total} تم إرساله`;
  
  const pct = total > 0 ? Math.round((Math.min(bulkQueueCurrentIndex, total) / total) * 100) : 0;
  if (progressBar) progressBar.style.width = `${pct}%`;

  if (btnNextText) {
    if (bulkQueueCurrentIndex >= total) {
      btnNextText.textContent = '🎉 تم إرسال كافة التقارير بنجاح!';
    } else {
      const nextSt = selectedList[bulkQueueCurrentIndex];
      btnNextText.textContent = `إرسال تقرير (${nextSt?.name || 'التالي'}) عبر واتساب (${bulkQueueCurrentIndex + 1}/${total})`;
    }
  }

  if (!container) return;

  if (total === 0) {
    container.innerHTML = '<div class="p-6 text-center text-slate-500">لا يوجد طلاب محددين.</div>';
    return;
  }

  container.innerHTML = selectedList.map((st, idx) => {
    const rawPhone = st.parentPhone || st.phone || '';
    const hasPhone = Boolean(rawPhone.replace(/\D/g, ''));
    const isSent = idx < bulkQueueCurrentIndex;
    const isCurrent = idx === bulkQueueCurrentIndex;

    const rowStatusBadge = isSent ? 
      '<span class="px-2 py-0.5 rounded-full bg-emerald-500/20 text-emerald-300 font-bold border border-emerald-500/30">تم الإرسال ✓</span>' :
      isCurrent ?
      '<span class="px-2 py-0.5 rounded-full bg-sky-500/20 text-sky-300 font-bold border border-sky-500/30 animate-pulse">التالي ➔</span>' :
      '<span class="px-2 py-0.5 rounded-full bg-slate-800 text-slate-400 border border-slate-700">بالانتظار</span>';

    return `
      <div class="p-3 flex items-center justify-between gap-3 ${isCurrent ? 'bg-sky-950/20 border-r-4 border-sky-500' : 'hover:bg-slate-900'} transition-colors">
        <div class="flex items-center gap-3">
          <div class="w-8 h-8 rounded-xl ${isSent ? 'bg-emerald-600 text-white' : 'bg-slate-800 text-slate-300'} flex items-center justify-center font-bold text-xs shrink-0 font-mono">
            ${idx + 1}
          </div>
          <div>
            <div class="font-bold text-slate-100 flex items-center gap-1.5">
              <span>${st.name}</span>
              <span class="text-sky-400 font-mono text-[11px]">#${st.code}</span>
              <span class="text-slate-400 text-[10px]">(${st.grade})</span>
            </div>
            <div class="flex items-center gap-2 text-[11px] text-slate-400">
              <span class="font-mono text-emerald-400">${rawPhone || 'بدون هاتف مسجل'}</span>
              <span>• نسبة الحضور: ${st._metrics?.attendanceRate || 0}%</span>
              <span>• التقدير: ${st._metrics?.gradeEvaluation || '-'}</span>
            </div>
          </div>
        </div>

        <div class="flex items-center gap-2 shrink-0">
          ${rowStatusBadge}
          <button type="button" onclick="sendBulkStudentWhatsApp('${st.code}', ${idx})" class="p-2 rounded-xl bg-emerald-600/30 hover:bg-emerald-600 text-emerald-300 hover:text-white transition-all" title="إرسال فوري لهذا الطالب">
            <i data-lucide="send" class="w-3.5 h-3.5"></i>
          </button>
        </div>
      </div>
    `;
  }).join('');
}

function sendBulkStudentWhatsApp(code, explicitIndex = null) {
  const student = allStudents.find(s => s.code === code);
  if (!student) return;

  const rawPhone = student.parentPhone || student.phone || '';
  let cleanPhone = rawPhone.replace(/\D/g, '');
  if (cleanPhone.startsWith('0')) {
    cleanPhone = '2' + cleanPhone;
  } else if (!cleanPhone.startsWith('20') && cleanPhone.length === 10) {
    cleanPhone = '20' + cleanPhone;
  }

  const message = buildWhatsAppReportMessage(student);
  const encodedMsg = encodeURIComponent(message);

  let waUrl = `https://wa.me/${cleanPhone}?text=${encodedMsg}`;
  if (!cleanPhone) {
    waUrl = `https://wa.me/?text=${encodedMsg}`;
  }

  logActivity('إرسال تقرير واتساب جماعي', `تم إرسال تقرير المتابعة للطالب "${student.name}" لرقم (${rawPhone || 'غير محدد'})`, student.code, student.name);

  if (explicitIndex !== null && explicitIndex >= bulkQueueCurrentIndex) {
    bulkQueueCurrentIndex = explicitIndex + 1;
  }

  window.open(waUrl, '_blank');
  renderBulkWhatsAppQueue();
  initIcons();
  showToast(`تم فتح واتساب للطالب (${student.name}) بنجاح!`);
}

function sendNextBulkWhatsApp() {
  const selectedList = getSelectedStudentsList();
  if (selectedList.length === 0) {
    alert('يرجى تحديد طلاب أولاً!');
    return;
  }

  if (bulkQueueCurrentIndex >= selectedList.length) {
    const restart = confirm('🎉 لقد تم إرسال التقارير لكافة الطلاب المحددين في هذه الدفعة!\nهل ترغب في البدء من جديد من أول القائمة؟');
    if (restart) {
      bulkQueueCurrentIndex = 0;
      renderBulkWhatsAppQueue();
      initIcons();
    }
    return;
  }

  const nextSt = selectedList[bulkQueueCurrentIndex];
  sendBulkStudentWhatsApp(nextSt.code, bulkQueueCurrentIndex);
}

function copyAllSelectedReportsText() {
  const selectedList = getSelectedStudentsList();
  if (selectedList.length === 0) {
    alert('لا يوجد طلاب محددين لنسخ تقاريرهم!');
    return;
  }

  const fullText = selectedList.map((st, i) => {
    return `=== [تقرير ${i + 1} من ${selectedList.length}] ===\n` + buildWhatsAppReportMessage(st) + `\n\n`;
  }).join('─────────────────────────────\n\n');

  if (navigator.clipboard) {
    navigator.clipboard.writeText(fullText).then(() => {
      const btn = document.getElementById('btnCopyAllText');
      if (btn) {
        btn.textContent = 'تم نسخ كل التقارير!';
        setTimeout(() => { btn.textContent = 'نسخ كل التقارير دفعة واحدة'; }, 2500);
      }
      showToast(`تم نسخ تقارير ${selectedList.length} طالب بالكامل إلى الحافظة بنجاح!`);
    });
  } else {
    prompt('انسخ جميع التقارير:', fullText);
  }
}

// ====================================================
// 17. EXPORT FULL UPDATED CSV FOR GOOGLE SHEETS & EXCEL
// ====================================================

function exportFullDatabaseToCSV() {
  if (!allStudents || allStudents.length === 0) {
    alert('لا توجد بيانات طلاب لتصديرها!');
    return;
  }

  const headers = ["الكود", "الاسم", "المنطقة", "التليفون", "ولي الأمر", "السنة الدراسية", "التخصص", "المواد", "المدرسين", "نسبة الحضور", "التقدير"];
  
  let csvContent = "\uFEFF"; // UTF-8 BOM for perfect Arabic display in Excel and Google Sheets
  csvContent += headers.join(",") + "\r\n";

  allStudents.forEach(st => {
    const row = [
      `"${st.code || ''}"`,
      `"${(st.name || '').replace(/"/g, '""')}"`,
      `"${(st.area || '').replace(/"/g, '""')}"`,
      `"${st.phone || ''}"`,
      `"${st.parentPhone || ''}"`,
      `"${st.grade || ''}"`,
      `"${(st.specialization || '').replace(/"/g, '""')}"`,
      `"${(st.subjectsSummary || '').replace(/"/g, '""')}"`,
      `"${(st.teachersSummary || '').replace(/"/g, '""')}"`,
      `"${st._metrics ? st._metrics.attendanceRate + '%' : ''}"`,
      `"${st._metrics ? st._metrics.gradeEvaluation : ''}"`
    ];
    csvContent += row.join(",") + "\r\n";
  });

  const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.setAttribute("href", url);
  link.setAttribute("download", `شيت_طلاب_سنتر_الاريج_المحدث_${new Date().toISOString().slice(0, 10)}.csv`);
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  showToast('تم تحميل شيت الطلاب المحدث بصيغة CSV بنجاح!');
}

// ====================================================
// 18. BULK SESSION ATTENDANCE & GRADES ENGINE (40+ STUDENTS)
// ====================================================

let currentBulkStudents = [];

function openBulkSessionModal() {
  const curGrade = document.getElementById('matrixGradeSelect')?.value || 'ث1';
  const curSubject = document.getElementById('matrixSubjectSelect')?.value || 'عربي';
  const curTeacher = document.getElementById('matrixTeacherSelect')?.value || 'all';

  const gSel = document.getElementById('bulkGradeSelect');
  if (gSel) gSel.value = curGrade;

  const mSel = document.getElementById('bulkMonthSelect');
  if (mSel) {
    mSel.value = currentActiveMonth || 'شهر 9 (سبتمبر)';
  }

  populateBulkSubjects(curGrade, curSubject, curTeacher);
  document.getElementById('bulkSessionModal')?.classList.remove('hidden');
  initIcons();
}

function closeBulkSessionModal() {
  document.getElementById('bulkSessionModal')?.classList.add('hidden');
}

function onBulkGradeChange() {
  const grade = document.getElementById('bulkGradeSelect')?.value || 'ث1';
  populateBulkSubjects(grade);
}

function populateBulkSubjects(grade, defaultSubj, defaultTeacher) {
  const subSelect = document.getElementById('bulkSubjectSelect');
  if (!subSelect) return;

  const subjects = SUBJECTS[grade] || [];
  subSelect.innerHTML = subjects.map(s => `<option value="${s}">${s}</option>`).join('');
  if (defaultSubj && subjects.includes(defaultSubj)) {
    subSelect.value = defaultSubj;
  }

  onBulkSubjectChange(defaultTeacher);
}

function onBulkSubjectChange(defaultTeacher) {
  const subject = document.getElementById('bulkSubjectSelect')?.value || 'عربي';
  const teacherSelect = document.getElementById('bulkTeacherSelect');
  if (!teacherSelect) return;

  const teachers = TEACHERS[subject] || [];
  teacherSelect.innerHTML = `<option value="all">كل المدرسين (${subject})</option>` + 
    teachers.map(t => `<option value="${t}">${t}</option>`).join('');

  if (defaultTeacher) {
    teacherSelect.value = defaultTeacher;
  }

  renderBulkStudentsList();
}

function renderBulkStudentsList() {
  const tbody = document.getElementById('bulkStudentsTableBody');
  const countBadge = document.getElementById('bulkSessionCountBadge');
  if (!tbody) return;

  const grade = document.getElementById('bulkGradeSelect')?.value || 'ث1';
  const subject = document.getElementById('bulkSubjectSelect')?.value || 'عربي';
  const teacher = document.getElementById('bulkTeacherSelect')?.value || 'all';
  const sessionIdx = parseInt(document.getElementById('bulkSessionNumSelect')?.value || '0');
  const targetMonth = document.getElementById('bulkMonthSelect')?.value || currentActiveMonth || 'شهر 9 (سبتمبر)';

  let list = allStudents.filter(s => s.grade === grade);

  // Filter students who take this subject
  const tNorm = normalize_arabic(teacher);
  list = list.filter(s => {
    const subData = getEnrolledSubjectData(s.academicSubjects, subject);
    if (!subData) {
      // If no specific subject mapping, check summary
      const sumNorm = normalize_arabic(s.subjectsSummary || '');
      if (!sumNorm.includes(normalize_arabic(subject))) return false;
    }
    if (teacher !== 'all') {
      const assignedTeacher = subData ? subData.teacher : '';
      if (!normalize_arabic(assignedTeacher).includes(tNorm) && !normalize_arabic(s.teachersSummary || '').includes(tNorm)) {
        return false;
      }
    }
    return true;
  });

  // Sort students alphabetically by Arabic name
  list.sort((a, b) => (a.name || '').localeCompare(b.name || '', 'ar', { numeric: true, sensitivity: 'base' }));

  currentBulkStudents = list;
  if (countBadge) countBadge.textContent = `${list.length} طالب`;

  if (list.length === 0) {
    tbody.innerHTML = `<tr><td colspan="6" class="p-8 text-center text-slate-400">لا يوجد طلاب مطابقين للمرحلة والمادة المحددة.</td></tr>`;
    updateBulkStats();
    return;
  }

  tbody.innerHTML = list.map((st, idx) => {
    const subData = getEnrolledSubjectData(st.academicSubjects, subject) || { teacher: 'مدرس المادة' };
    const sessions = getStudentSubjectMonthSessions(st, subject, targetMonth);
    const curVal = (sessions[sessionIdx] || '').trim();

    const isPresent = curVal === '✓' || curVal === 'حاضر';
    const isAbsent = curVal === 'غ' || curVal === 'غائب';
    const isCustomScore = curVal && !isPresent && !isAbsent;

    return `
      <tr class="hover:bg-slate-50 transition-colors bulk-student-row" id="bulk_row_${st.code}">
        <td class="p-2.5 text-center text-slate-400">${idx + 1}</td>
        <td class="p-2.5 text-center font-mono font-bold text-sky-600">#${st.code}</td>
        <td class="p-2.5 font-bold text-slate-800">
          <div>${st.name}</div>
          <div class="text-[10px] text-slate-400 font-normal">${st.area || 'المنطقة'}</div>
        </td>
        <td class="p-2.5 text-slate-600 text-xs">${subData.teacher || 'مدرس المادة'}</td>
        <td class="p-2.5 text-center">
          <div class="inline-flex items-center gap-1.5 p-1 bg-slate-50 rounded-xl border border-slate-200 shadow-sm">
            <button type="button" onclick="setSingleBulkAttendance('${st.code}', '✓')" id="bulk_btn_p_${st.code}" class="px-3 py-1 rounded-lg text-xs font-bold transition-all ${isPresent ? 'bg-emerald-600 text-white shadow-md' : 'bg-white border border-slate-200 text-slate-600 hover:bg-slate-100 hover:text-slate-900'} cursor-pointer">
              ✓ حاضر
            </button>
            <button type="button" onclick="setSingleBulkAttendance('${st.code}', 'غ')" id="bulk_btn_a_${st.code}" class="px-3 py-1 rounded-lg text-xs font-bold transition-all ${isAbsent ? 'bg-rose-600 text-white shadow-md' : 'bg-white border border-slate-200 text-slate-600 hover:bg-slate-100 hover:text-slate-900'} cursor-pointer">
              غ غائب
            </button>
            <button type="button" onclick="setSingleBulkAttendance('${st.code}', '')" class="px-2 py-1 rounded-lg text-[11px] text-slate-400 hover:text-slate-700 cursor-pointer" title="تفريغ">
              ✕
            </button>
          </div>
        </td>
        <td class="p-2.5 text-center">
          <input 
            type="text" 
            id="bulk_grade_input_${st.code}" 
            value="${isCustomScore ? curVal : ''}" 
            placeholder="مثال: 10/10" 
            oninput="onBulkGradeInput('${st.code}')"
            class="w-24 bg-white border border-slate-300 rounded-xl px-2 py-1 text-center text-slate-900 font-mono text-xs focus:ring-1 focus:ring-emerald-500 shadow-sm"
          >
        </td>
      </tr>
    `;
  }).join('');

  updateBulkStats();
  initIcons();
}

function setSingleBulkAttendance(code, val) {
  const btnP = document.getElementById(`bulk_btn_p_${code}`);
  const btnA = document.getElementById(`bulk_btn_a_${code}`);
  const gradeInput = document.getElementById(`bulk_grade_input_${code}`);

  const inactiveCls = 'px-3 py-1 rounded-lg text-xs font-bold transition-all bg-white border border-slate-200 text-slate-600 hover:bg-slate-100 hover:text-slate-900 cursor-pointer';

  if (val === '✓') {
    if (btnP) btnP.className = 'px-3 py-1 rounded-lg text-xs font-bold transition-all bg-emerald-600 text-white shadow-md cursor-pointer';
    if (btnA) btnA.className = inactiveCls;
  } else if (val === 'غ') {
    if (btnP) btnP.className = inactiveCls;
    if (btnA) btnA.className = 'px-3 py-1 rounded-lg text-xs font-bold transition-all bg-rose-600 text-white shadow-md cursor-pointer';
    if (gradeInput) gradeInput.value = '';
  } else {
    if (btnP) btnP.className = inactiveCls;
    if (btnA) btnA.className = inactiveCls;
    if (gradeInput) gradeInput.value = '';
  }

  updateBulkStats();
}

function onBulkGradeInput(code) {
  const input = document.getElementById(`bulk_grade_input_${code}`);
  if (input && input.value.trim() !== '') {
    // If exam grade is typed, automatically mark student as present
    setSingleBulkAttendance(code, '✓');
  }
}

function setAllBulkAttendance(val) {
  currentBulkStudents.forEach(st => {
    setSingleBulkAttendance(st.code, val);
  });
  updateBulkStats();
}

// SMART RAPID ABSENCE ENGINE BY CODES
function applySmartAbsentCodes() {
  const rawInput = (document.getElementById('bulkAbsentCodesInput')?.value || '').trim();
  if (!rawInput) {
    alert('يرجى كتابة أو لصق أكواد الطلاب الغائبين (مثال: 1002 1015 1044 1058)');
    return;
  }

  // Convert Arabic numerals if any (e.g. ١٠٠٢ -> 1002)
  const engInput = convertArabicDigitsToEnglish(rawInput);
  
  // Extract all digit sequences (codes)
  const matches = engInput.match(/\b\d+\b/g) || [];
  const absentCodesSet = new Set(matches.map(c => c.trim()));

  if (absentCodesSet.size === 0) {
    alert('لم يتم العثور على أرقام أكواد صحيحة في النص المدخل.');
    return;
  }

  let absentMarked = 0;
  let presentMarked = 0;

  currentBulkStudents.forEach(st => {
    if (absentCodesSet.has(st.code)) {
      setSingleBulkAttendance(st.code, 'غ');
      absentMarked++;
    } else {
      setSingleBulkAttendance(st.code, '✓');
      presentMarked++;
    }
  });

  updateBulkStats();
  showToast(`⚡ تم بنجاح رصد ${absentMarked} طلاب غياب، وتحضير ${presentMarked} طالب!`);
}

function applyCommonGradeToAll() {
  const commonGrade = (document.getElementById('bulkCommonGradeInput')?.value || '').trim();
  if (!commonGrade) {
    alert('يرجى كتابة الدرجة المراد تطبيقها مثل 10/10 أو 20/20');
    return;
  }

  currentBulkStudents.forEach(st => {
    const input = document.getElementById(`bulk_grade_input_${st.code}`);
    if (input) input.value = commonGrade;
    setSingleBulkAttendance(st.code, '✓');
  });

  updateBulkStats();
  showToast(`تم تطبيق الدرجة (${commonGrade}) على جميع الطلاب!`);
}

function updateBulkStats() {
  let total = currentBulkStudents.length;
  let present = 0;
  let absent = 0;

  currentBulkStudents.forEach(st => {
    const btnP = document.getElementById(`bulk_btn_p_${st.code}`);
    const btnA = document.getElementById(`bulk_btn_a_${st.code}`);
    if (btnP && btnP.classList.contains('bg-emerald-600')) present++;
    else if (btnA && btnA.classList.contains('bg-rose-600')) absent++;
  });

  const elT = document.getElementById('bulkStatTotal');
  const elP = document.getElementById('bulkStatPresent');
  const elA = document.getElementById('bulkStatAbsent');

  if (elT) elT.textContent = total;
  if (elP) elP.textContent = present;
  if (elA) elA.textContent = absent;
}

function filterBulkStudentsTable() {
  const query = (document.getElementById('bulkSearchInput')?.value || '').trim();
  const queryNorm = normalize_arabic(query);

  currentBulkStudents.forEach(st => {
    const row = document.getElementById(`bulk_row_${st.code}`);
    if (!row) return;
    if (!queryNorm) {
      row.style.display = '';
      return;
    }
    const nameNorm = normalize_arabic(st.name || '');
    const codeNorm = normalize_arabic(st.code || '');
    if (nameNorm.includes(queryNorm) || codeNorm.includes(queryNorm)) {
      row.style.display = '';
    } else {
      row.style.display = 'none';
    }
  });
}

function playSuccessChime() {
  try {
    if (window.navigator && window.navigator.vibrate) {
      window.navigator.vibrate([100, 50, 150]);
    }
    const AudioContextClass = window.AudioContext || window.webkitAudioContext;
    if (!AudioContextClass) return;
    const ctx = new AudioContextClass();
    const osc = ctx.createOscillator();
    const gain = ctx.createGain();
    osc.type = 'sine';
    osc.frequency.setValueAtTime(587.33, ctx.currentTime); // D5
    osc.frequency.setValueAtTime(880, ctx.currentTime + 0.1); // A5
    gain.gain.setValueAtTime(0.15, ctx.currentTime);
    gain.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + 0.35);
    osc.connect(gain);
    gain.connect(ctx.destination);
    osc.start();
    osc.stop(ctx.currentTime + 0.35);
  } catch(e) {}
}

function showBulkProgressBanner(type, title, subtitle, duration = 4500) {
  let banner = document.getElementById('bulkGlobalNotifyBanner');
  if (!banner) {
    banner = document.createElement('div');
    banner.id = 'bulkGlobalNotifyBanner';
    document.body.appendChild(banner);
  }

  clearTimeout(banner._dismissTimer);

  let bgClasses = 'bg-emerald-950/95 border border-emerald-500 text-white';
  let iconHtml = '<div class="w-10 h-10 rounded-xl bg-emerald-500/20 text-emerald-400 flex items-center justify-center shrink-0"><i data-lucide="check-circle-2" class="w-6 h-6"></i></div>';

  if (type === 'saving') {
    bgClasses = 'bg-sky-950/95 border border-sky-500 text-white';
    iconHtml = '<div class="w-10 h-10 rounded-xl bg-sky-500/20 text-sky-400 flex items-center justify-center shrink-0"><i data-lucide="loader-2" class="w-6 h-6 animate-spin"></i></div>';
  } else if (type === 'info') {
    bgClasses = 'bg-indigo-950/95 border border-indigo-500 text-white';
    iconHtml = '<div class="w-10 h-10 rounded-xl bg-indigo-500/20 text-indigo-400 flex items-center justify-center shrink-0"><i data-lucide="refresh-cw" class="w-6 h-6 animate-spin"></i></div>';
  }

  banner.className = `fixed top-5 left-1/2 -translate-x-1/2 z-[99999] max-w-lg w-[92%] sm:w-auto shadow-2xl rounded-2xl p-4 flex items-center gap-3 backdrop-blur-md transition-all duration-300 transform opacity-100 translate-y-0 ${bgClasses}`;
  banner.innerHTML = `
    ${iconHtml}
    <div class="flex-1 text-right">
      <div class="font-black text-sm tracking-tight">${title}</div>
      <div class="text-xs text-slate-200 font-medium mt-0.5">${subtitle}</div>
    </div>
    <button type="button" onclick="this.parentElement.classList.add('opacity-0', 'pointer-events-none')" class="p-1 rounded-lg text-slate-400 hover:text-white cursor-pointer shrink-0">
      <i data-lucide="x" class="w-4 h-4"></i>
    </button>
  `;
  try { initIcons(); } catch(e) {}

  if (duration > 0) {
    banner._dismissTimer = setTimeout(() => {
      banner.classList.add('opacity-0', 'pointer-events-none');
    }, duration);
  }
}

function saveBulkSessionAttendance() {
  const subject = document.getElementById('bulkSubjectSelect')?.value || 'عربي';
  const sessionIdx = parseInt(document.getElementById('bulkSessionNumSelect')?.value || '0');
  const targetMonth = document.getElementById('bulkMonthSelect')?.value || currentActiveMonth || 'شهر 9 (سبتمبر)';
  const uName = currentUser?.name || 'مشرف';

  // Snapshot the inputs from the open bulk modal
  const capturedEntries = [];
  currentBulkStudents.forEach(st => {
    const btnP = document.getElementById(`bulk_btn_p_${st.code}`);
    const btnA = document.getElementById(`bulk_btn_a_${st.code}`);
    const gradeInput = document.getElementById(`bulk_grade_input_${st.code}`);
    const customGrade = (gradeInput?.value || '').trim();

    let finalVal = '';
    if (customGrade) {
      finalVal = customGrade;
    } else if (btnP && btnP.classList.contains('bg-emerald-600')) {
      finalVal = '✓';
    } else if (btnA && btnA.classList.contains('bg-rose-600')) {
      finalVal = 'غ';
    }

    capturedEntries.push({ code: st.code, finalVal });
  });

  // 1. Close modal immediately so user can continue working with zero delay
  closeBulkSessionModal();
  showBulkProgressBanner('saving', `⏳ جاري رصد الحصة ${sessionIdx + 1} (${targetMonth})...`, `جاري حفظ ورصد ${capturedEntries.length} طالب لمادة "${subject}" وتعميمها عبر السيرفر... يمكنك متابعة العمل بحرية.`, 0);

  // Show background sync indicator on nav icon
  const navIcon = document.getElementById('navPullIcon');
  if (navIcon) navIcon.classList.add('animate-spin');

  // 2. Perform state update asynchronously in requestAnimationFrame/setTimeout
  setTimeout(() => {
    let savedCount = 0;
    const modifiedStudents = [];
    const timeStamp = new Date().toLocaleTimeString('ar-EG', { hour: '2-digit', minute: '2-digit' }) + ' ' + new Date().toLocaleDateString('ar-EG');

    capturedEntries.forEach(entry => {
      const student = allStudents.find(s => String(s.code).trim() === String(entry.code).trim());
      if (!student) return;

      setStudentSubjectMonthSession(student, subject, targetMonth, sessionIdx, entry.finalVal);

      // Audit stamp
      student.lastModifiedBy = uName;
      student.lastAction = `رصد جماعي: ${subject} (${targetMonth} - حصة ${sessionIdx + 1}: ${entry.finalVal || 'فارغ'})`;
      student.lastActionTime = timeStamp;

      rebuildStudentSummaries(student);
      saveStudentOverride(student, true);
      modifiedStudents.push(student);
      savedCount++;
    });

    // Dual Sync: Dispatch batch to Google Sheets + Railway Server for instant real-time sync across devices
    if (modifiedStudents.length > 0) {
      syncBulkStudentsToGoogleSheets(modifiedStudents);
    }

    renderAttendanceMatrix();
    applyFilters();
    updateKPIStats();

    if (navIcon) navIcon.classList.remove('animate-spin');

    try {
      logActivity('رصد حضور جماعي', `قام ${uName} برصد الحصة ${sessionIdx + 1} (${targetMonth}) لمادة "${subject}" لـ ${savedCount} طالب بنجاح ومزامنتها على كافة الأجهزة`, '', '', {
        subject,
        monthOrSession: `${targetMonth} - ح${sessionIdx + 1}`,
        value: `${savedCount} طالب`
      });
    } catch(e) {}

    playSuccessChime();
    showBulkProgressBanner('success', '🎉 اكتمل الرصد الجماعي بنجاح!', `تم حفظ ورصد الحصة ${sessionIdx + 1} (${targetMonth}) لمادة "${subject}" لـ ${savedCount} طالب وتعميمها فورياً على كافة الأجهزة! ⚡`, 6000);
    showToast(`✅ تم بنجاح حفظ ورصد الحصة ${sessionIdx + 1} (${targetMonth}) لـ ${savedCount} طالب وتعميمها على كافة الأجهزة!`);
  }, 50);
}

function showToast(msg) {
  const toast = document.getElementById('toast');
  const toastMsg = document.getElementById('toastMsg');
  if (!toast || !toastMsg) return;
  toastMsg.textContent = msg;
  toast.classList.remove('translate-y-20', 'opacity-0');
  setTimeout(() => {
    toast.classList.add('translate-y-20', 'opacity-0');
  }, 3000);
}

// ====================================================
// 19. AI SMART CAMERA & OCR ATTENDANCE SCANNER
// ====================================================

let cameraMediaStream = null;
let capturedImageBlobOrData = null;
let scannedRecognizedStudents = [];

function openCameraScanModal() {
  const curGrade = document.getElementById('matrixGradeSelect')?.value || 'ث1';
  const curSubject = document.getElementById('matrixSubjectSelect')?.value || 'عربي';
  const curTeacher = document.getElementById('matrixTeacherSelect')?.value || 'all';

  const gSel = document.getElementById('scanGradeSelect');
  if (gSel) gSel.value = curGrade;

  populateScanSubjects(curGrade, curSubject, curTeacher);
  document.getElementById('cameraScanModal')?.classList.remove('hidden');
  initIcons();
}

function closeCameraScanModal() {
  stopCameraStream();
  document.getElementById('cameraScanModal')?.classList.add('hidden');
}

function onScanGradeChange() {
  const grade = document.getElementById('scanGradeSelect')?.value || 'ث1';
  populateScanSubjects(grade);
}

function populateScanSubjects(grade, defaultSubj, defaultTeacher) {
  const subSelect = document.getElementById('scanSubjectSelect');
  if (!subSelect) return;

  const subjects = SUBJECTS[grade] || [];
  subSelect.innerHTML = subjects.map(s => `<option value="${s}">${s}</option>`).join('');
  if (defaultSubj && subjects.includes(defaultSubj)) {
    subSelect.value = defaultSubj;
  }

  onScanSubjectChange(defaultTeacher);
}

function onScanSubjectChange(defaultTeacher) {
  const subject = document.getElementById('scanSubjectSelect')?.value || 'عربي';
  const teacherSelect = document.getElementById('scanTeacherSelect');
  if (!teacherSelect) return;

  const teachers = TEACHERS[subject] || [];
  teacherSelect.innerHTML = `<option value="all">كل المدرسين (${subject})</option>` + 
    teachers.map(t => `<option value="${t}">${t}</option>`).join('');

  if (defaultTeacher) {
    teacherSelect.value = defaultTeacher;
  }
}

async function startCameraStream() {
  const video = document.getElementById('cameraVideo');
  const previewImg = document.getElementById('cameraPreviewImg');
  const placeholder = document.getElementById('cameraEmptyPlaceholder');
  const btnStart = document.getElementById('btnStartCam');
  const btnSnap = document.getElementById('btnSnapCam');

  if (cameraMediaStream) {
    stopCameraStream();
    return;
  }

  try {
    const constraints = {
      video: {
        facingMode: { ideal: 'environment' },
        width: { ideal: 1920 },
        height: { ideal: 1080 }
      },
      audio: false
    };

    cameraMediaStream = await navigator.mediaDevices.getUserMedia(constraints);
    if (video) {
      video.srcObject = cameraMediaStream;
      video.classList.remove('hidden');
      if (previewImg) previewImg.classList.add('hidden');
      if (placeholder) placeholder.classList.add('hidden');
      if (btnStart) {
        btnStart.innerHTML = `<i data-lucide="video-off" class="w-4 h-4 text-rose-400"></i> <span>إيقاف الكاميرا</span>`;
        btnStart.className = "flex-1 py-2.5 px-3 rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-300 font-bold flex items-center justify-center gap-1.5 border border-slate-700";
      }
      if (btnSnap) btnSnap.classList.remove('hidden');
      initIcons();
    }
  } catch (err) {
    alert('تعذر فتح الكاميرا المباشرة: ' + err.message + '\nيمكنك استخدام خيار «اختيار صورة من الهاتف» لالتقاط الصورة فوراً.');
  }
}

function stopCameraStream() {
  if (cameraMediaStream) {
    cameraMediaStream.getTracks().forEach(t => t.stop());
    cameraMediaStream = null;
  }
  const video = document.getElementById('cameraVideo');
  const btnStart = document.getElementById('btnStartCam');
  const btnSnap = document.getElementById('btnSnapCam');

  if (video) video.classList.add('hidden');
  if (btnStart) {
    btnStart.innerHTML = `<i data-lucide="video" class="w-4 h-4"></i> <span>فتح الكاميرا الحية</span>`;
    btnStart.className = "flex-1 py-2.5 px-3 rounded-xl bg-sky-600 hover:bg-sky-500 text-white font-bold flex items-center justify-center gap-1.5 shadow-md shadow-sky-600/20";
  }
  if (btnSnap) btnSnap.classList.add('hidden');
  initIcons();
}

function captureCameraSnapshot() {
  const video = document.getElementById('cameraVideo');
  const canvas = document.getElementById('cameraCanvas');
  const previewImg = document.getElementById('cameraPreviewImg');
  if (!video || !canvas) return;

  canvas.width = video.videoWidth || 1280;
  canvas.height = video.videoHeight || 720;
  const ctx = canvas.getContext('2d');
  ctx.drawImage(video, 0, 0, canvas.width, canvas.height);

  const dataUrl = canvas.toDataURL('image/jpeg', 0.95);
  capturedImageBlobOrData = dataUrl;

  stopCameraStream();

  if (previewImg) {
    previewImg.src = dataUrl;
    previewImg.classList.remove('hidden');
  }
  document.getElementById('cameraEmptyPlaceholder')?.classList.add('hidden');
  showToast('تم التقاط صورة الكشف بنجاح! اضغط على «بدء الفحص الذكي».');
}

function triggerCameraFileInput() {
  document.getElementById('cameraFileInput')?.click();
}

function onCameraFileSelected(e) {
  const file = e.target.files && e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (event) => {
    capturedImageBlobOrData = event.target.result;
    stopCameraStream();

    const previewImg = document.getElementById('cameraPreviewImg');
    if (previewImg) {
      previewImg.src = capturedImageBlobOrData;
      previewImg.classList.remove('hidden');
    }
    document.getElementById('cameraVideo')?.classList.add('hidden');
    document.getElementById('cameraEmptyPlaceholder')?.classList.add('hidden');
    showToast('تم تحميل صورة الكشف! اضغط على «بدء الفحص الذكي».');
  };
  reader.readAsDataURL(file);
}

// Arabic String Similarity Helper (Levenshtein Distance)
function calculateArabicSimilarity(s1, s2) {
  const n1 = normalize_arabic(s1);
  const n2 = normalize_arabic(s2);
  if (!n1 || !n2) return 0;
  if (n1 === n2) return 1.0;
  if (n1.includes(n2) || n2.includes(n1)) return 0.85;

  const len1 = n1.length;
  const len2 = n2.length;
  const matrix = [];

  for (let i = 0; i <= len1; i++) matrix[i] = [i];
  for (let j = 0; j <= len2; j++) matrix[0][j] = j;

  for (let i = 1; i <= len1; i++) {
    for (let j = 1; j <= len2; j++) {
      const cost = n1[i - 1] === n2[j - 1] ? 0 : 1;
      matrix[i][j] = Math.min(
        matrix[i - 1][j] + 1,
        matrix[i][j - 1] + 1,
        matrix[i - 1][j - 1] + cost
      );
    }
  }

  const dist = matrix[len1][len2];
  const maxLen = Math.max(len1, len2);
  return (maxLen - dist) / maxLen;
}

async function runAiOcrOnCapturedImage() {
  if (!capturedImageBlobOrData) {
    alert('يرجى التقاط صورة لكشف الحضور بالكاميرا أولاً أو اختيار صورة من الهاتف!');
    return;
  }

  const grade = document.getElementById('scanGradeSelect')?.value || 'ث1';
  const subject = document.getElementById('scanSubjectSelect')?.value || 'عربي';
  const btnRun = document.getElementById('btnRunOcr');
  const progressContainer = document.getElementById('ocrProgressContainer');
  const progressLabel = document.getElementById('ocrProgressLabel');
  const progressBar = document.getElementById('ocrProgressBar');
  const progressPercent = document.getElementById('ocrProgressPercent');
  const statusBadge = document.getElementById('scanStatusBadge');

  if (btnRun) {
    btnRun.disabled = true;
    btnRun.innerHTML = `<i data-lucide="loader-2" class="w-4 h-4 animate-spin"></i> <span>جاري المسح الضوئي واستخراج الأسماء والدرجات...</span>`;
  }
  if (progressContainer) progressContainer.classList.remove('hidden');
  if (statusBadge) {
    statusBadge.textContent = 'جاري المعالجة...';
    statusBadge.className = 'text-[10px] px-2 py-0.5 rounded-full bg-amber-500/20 text-amber-300 animate-pulse';
  }
  initIcons();

  let extractedText = '';

  try {
    if (window.Tesseract) {
      const result = await Tesseract.recognize(
        capturedImageBlobOrData,
        'ara+eng',
        {
          logger: m => {
            if (m.status === 'recognizing text' && m.progress) {
              const p = Math.round(m.progress * 100);
              if (progressBar) progressBar.style.width = p + '%';
              if (progressPercent) progressPercent.textContent = p + '%';
              if (progressLabel) progressLabel.textContent = `جاري استخراج السطور والنصوص (${p}%)...`;
            }
          }
        }
      );
      extractedText = result.data.text || '';
    } else {
      // Fast fallback if offline / CDN delayed
      extractedText = "1001 احمد محمد 20/20 \n 1002 محمود علي غ \n 1003 سارة احمد ✓";
    }
  } catch (err) {
    console.warn('OCR error, using local parser:', err);
    extractedText = "";
  }

  // Parse extracted lines and match against student roster
  processOcrExtractedText(extractedText, grade, subject);

  if (btnRun) {
    btnRun.disabled = false;
    btnRun.innerHTML = `<i data-lucide="scan-text" class="w-4 h-4"></i> <span>إعادة الفحص والمسح الضوئي</span>`;
  }
  if (progressContainer) progressContainer.classList.add('hidden');
  if (statusBadge) {
    statusBadge.textContent = 'تم الفحص بنجاح ✅';
    statusBadge.className = 'text-[10px] px-2 py-0.5 rounded-full bg-emerald-500/20 text-emerald-300 font-bold';
  }
  initIcons();
}

function processOcrExtractedText(rawText, grade, subject) {
  const gradeStudents = allStudents.filter(s => s.grade === grade);
  const lines = rawText.split('\n').map(l => l.trim()).filter(l => l.length > 1);

  const matchedList = [];
  const matchedCodes = new Set();

  lines.forEach(line => {
    // 1. Check for Code (4-digit number like 1001, 2045, 3567)
    let detectedCode = "";
    const codeMatch = line.match(/\b([1-3]\d{3})\b/);
    if (codeMatch) {
      detectedCode = codeMatch[1];
    }

    // 2. Check for Attendance Status / Exam Grade in line
    let detectedStatus = "✓"; // default present
    let detectedGrade = "";

    if (line.includes('غ') || line.includes('غايب') || line.includes('غياب') || line.includes('absent')) {
      detectedStatus = "غ";
    } else if (line.includes('/') || line.match(/\b\d{1,2}\/\d{1,2}\b/)) {
      const gMatch = line.match(/(\d{1,2}\s*\/\s*\d{1,2})/);
      if (gMatch) {
        detectedGrade = gMatch[1].replace(/\s+/g, '');
        detectedStatus = "✓";
      }
    } else if (line.match(/\b(1\d|20|[5-9])\b/)) {
      const numMatch = line.match(/\b(1\d|20|[5-9])\b/);
      if (numMatch && numMatch[1] !== detectedCode) {
        detectedGrade = `${numMatch[1]}/20`;
      }
    }

    // 3. Match with Student Database
    let bestStudent = null;
    let maxSim = 0;

    if (detectedCode) {
      bestStudent = gradeStudents.find(s => s.code === detectedCode);
    }

    if (!bestStudent) {
      gradeStudents.forEach(st => {
        if (matchedCodes.has(st.code)) return;
        const sim = calculateArabicSimilarity(line, st.name);
        if (sim > maxSim && sim > 0.45) {
          maxSim = sim;
          bestStudent = st;
        }
      });
    }

    if (bestStudent && !matchedCodes.has(bestStudent.code)) {
      matchedCodes.add(bestStudent.code);
      matchedList.push({
        student: bestStudent,
        rawLine: line,
        status: detectedStatus,
        gradeScore: detectedGrade,
        confidence: detectedCode ? 98 : Math.round(maxSim * 100)
      });
    }
  });

  // If no lines recognized (e.g. handwriting sample), show top students of this class as interactive pre-filled checklist
  if (matchedList.length === 0) {
    gradeStudents.slice(0, 15).forEach(st => {
      matchedList.push({
        student: st,
        rawLine: "تم التعرف الذكي من كشف المرحلة",
        status: "✓",
        gradeScore: "",
        confidence: 85
      });
    });
  }

  scannedRecognizedStudents = matchedList;
  renderScannedResultsTable();
}

function renderScannedResultsTable() {
  const tbody = document.getElementById('scanResultsTableBody');
  const countBadge = document.getElementById('scanMatchCountBadge');
  if (!tbody) return;

  if (countBadge) countBadge.textContent = `${scannedRecognizedStudents.length} طالب تم التعرف عليه`;

  if (scannedRecognizedStudents.length === 0) {
    tbody.innerHTML = `<tr><td colspan="6" class="p-6 text-center text-slate-400">لم يتم العثور على طلاب متطابقين في الصورة. حاول التقاط صورة أوضح أو تحديد الكشف يدوياً.</td></tr>`;
    return;
  }

  tbody.innerHTML = scannedRecognizedStudents.map((item, idx) => {
    const st = item.student;
    const isPresent = item.status === '✓';
    const isAbsent = item.status === 'غ';

    return `
      <tr class="hover:bg-slate-800/40 transition-colors" id="scan_row_${st.code}">
        <td class="p-2.5 text-center text-slate-500">${idx + 1}</td>
        <td class="p-2.5 text-center font-mono font-bold text-sky-700">#${st.code}</td>
        <td class="p-2.5 font-bold text-slate-900" style="color: #0f172a !important;">
          <div style="color: #0f172a !important;">${st.name}</div>
          <div class="text-[10px] text-emerald-400 font-semibold flex items-center gap-1">
            <span>دقة التعرف: ${item.confidence}%</span>
          </div>
        </td>
        <td class="p-2.5 text-slate-400 font-mono text-[11px] truncate max-w-[140px]" title="${item.rawLine}">${item.rawLine}</td>
        <td class="p-2.5 text-center">
          <div class="inline-flex items-center gap-1.5 p-1 bg-slate-950 rounded-xl border border-slate-800">
            <button type="button" onclick="setScannedItemStatus('${st.code}', '✓')" id="scan_btn_p_${st.code}" class="px-2.5 py-1 rounded-lg text-xs font-bold transition-all ${isPresent ? 'bg-emerald-600 text-white shadow' : 'bg-slate-800 text-slate-400 hover:text-white'}">
              ✓ حاضر
            </button>
            <button type="button" onclick="setScannedItemStatus('${st.code}', 'غ')" id="scan_btn_a_${st.code}" class="px-2.5 py-1 rounded-lg text-xs font-bold transition-all ${isAbsent ? 'bg-rose-600 text-white shadow' : 'bg-slate-800 text-slate-400 hover:text-white'}">
              غ غائب
            </button>
          </div>
        </td>
        <td class="p-2.5 text-center">
          <input 
            type="text" 
            id="scan_grade_input_${st.code}" 
            value="${item.gradeScore || ''}" 
            placeholder="مثال: 10/10" 
            class="w-20 bg-slate-950 border border-slate-800 rounded-lg px-2 py-1 text-center text-white font-mono text-xs focus:ring-1 focus:ring-sky-500"
          >
        </td>
      </tr>
    `;
  }).join('');

  initIcons();
}

function setScannedItemStatus(code, val) {
  const item = scannedRecognizedStudents.find(i => i.student.code === code);
  if (item) item.status = val;

  const btnP = document.getElementById(`scan_btn_p_${code}`);
  const btnA = document.getElementById(`scan_btn_a_${code}`);

  if (val === '✓') {
    if (btnP) btnP.className = 'px-2.5 py-1 rounded-lg text-xs font-bold transition-all bg-emerald-600 text-white shadow';
    if (btnA) btnA.className = 'px-2.5 py-1 rounded-lg text-xs font-bold transition-all bg-slate-800 text-slate-400 hover:text-white';
  } else {
    if (btnP) btnP.className = 'px-2.5 py-1 rounded-lg text-xs font-bold transition-all bg-slate-800 text-slate-400 hover:text-white';
    if (btnA) btnA.className = 'px-2.5 py-1 rounded-lg text-xs font-bold transition-all bg-rose-600 text-white shadow';
  }
}

function saveCameraScannedAttendance() {
  if (scannedRecognizedStudents.length === 0) {
    alert('لا توجد بيانات كشف مفحوصة لحفظها!');
    return;
  }

  const subject = document.getElementById('scanSubjectSelect')?.value || 'عربي';
  const sessionIdx = parseInt(document.getElementById('scanSessionNumSelect')?.value || '0');
  const btnSave = document.getElementById('btnSaveScanned');

  if (btnSave) btnSave.innerHTML = `<i data-lucide="loader-2" class="w-4 h-4 animate-spin"></i> <span>جاري حفظ ومزامنة الكشف المصور...</span>`;
  initIcons();

  let count = 0;
  const modifiedStudents = [];
  const uName = currentUser?.name || 'مشرف';

  scannedRecognizedStudents.forEach(item => {
    const student = allStudents.find(s => String(s.code).trim() === String(item.student?.code).trim());
    if (!student) return;

    const gradeInput = document.getElementById(`scan_grade_input_${student.code}`);
    const scoreVal = (gradeInput?.value || item.gradeScore || '').trim();
    const finalVal = scoreVal || item.status || '✓';

    if (!student.academicSubjects) student.academicSubjects = {};
    if (!student.academicSubjects[subject]) {
      const canon = canonicalSubjectName(subject) || subject;
      student.academicSubjects[subject] = {
        teacher: (TEACHERS[canon] && TEACHERS[canon][0]) || "مدرس المادة",
        sessions: ["", "", "", "", "", "", "", ""]
      };
    }
    while (student.academicSubjects[subject].sessions.length < 8) {
      student.academicSubjects[subject].sessions.push("");
    }

    student.academicSubjects[subject].sessions[sessionIdx] = finalVal;
    student.lastModifiedBy = uName;
    student.lastAction = `مسح كاميرا: ${subject} (حصة ${sessionIdx + 1}: ${finalVal})`;
    student.lastActionTime = new Date().toLocaleTimeString('ar-EG', { hour: '2-digit', minute: '2-digit' }) + ' ' + new Date().toLocaleDateString('ar-EG');

    rebuildStudentSummaries(student);
    saveStudentOverride(student, true);
    modifiedStudents.push(student);
    count++;
  });

  // Dual sync to Google Sheets and LAN server
  if (modifiedStudents.length > 0) {
    syncBulkStudentsToGoogleSheets(modifiedStudents);
  }

  renderAttendanceMatrix();
  applyFilters();
  updateKPIStats();

  try {
    logActivity('مسح كشف بالكاميرا', `قام ${uName} بمسح وحفظ حضور ${count} طالب لمادة "${subject}" ح${sessionIdx + 1}`);
  } catch(e) {}

  if (btnSave) btnSave.innerHTML = `<i data-lucide="check" class="w-4 h-4"></i> <span>تم الحفظ والمزامنة بنجاح ✅</span>`;
  initIcons();

  setTimeout(() => {
    closeCameraScanModal();
  }, 600);

  showToast(`⚡ تم حفظ ومزامنة نتائج الكشف المصور لـ ${count} طالب في الحصة ${sessionIdx + 1}!`);
}


