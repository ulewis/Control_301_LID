// LID301 v10 · Docentes pueden gestionar sus estudiantes + revisión de funcionalidades de estudiante
const SPREADSHEET_ID = '1Vg3vbUodSbpK5eQ6OAsxrzTaWy9PQPYEPcJSISwyvCA';

const SHEETS = Object.freeze({
  STUDENTS: 'ESTUDIANTES',
  TEACHERS: 'DOCENTES',
  PROJECTS: 'PROYECTOS',
  SCHEDULES: 'HORARIOS',
  ATTENDANCE: 'ASISTENCIAS',
  AUDIT: 'AUDITORIA',
  CONFIG: 'CONFIG',
  CHALLENGES: 'CHALLENGES'
});

const COL = Object.freeze({
  STUDENTS: { ID: 1, TEACHER_EMAIL: 2, FIRST: 3, LAST: 4, EMAIL: 5, CAREER: 6, PROJECT_ID: 7, STATUS: 8, MIN_HOURS: 9, ROLE: 10, CREATED: 11, UPDATED: 12 },
  TEACHERS: { ID: 1, NAME: 2, EMAIL: 3, PHONE: 4, STATUS: 5, CREATED: 6, UPDATED: 7 },
  PROJECTS: { ID: 1, SIDISI: 2, TITLE: 3, ALIAS: 4, FINANCED: 5, FUNDER: 6, AMOUNT: 7, TEACHER_EMAIL: 8, STATUS: 9 },
  SCHEDULES: { ID: 1, EMAIL: 2, WEEK_START: 3, DATE: 4, DAY: 5, START: 6, END: 7, MINUTES: 8, STATUS: 9, CREATED: 10, UPDATED: 11, NOTE: 12 },
  ATTENDANCE: { ID: 1, EMAIL: 2, DATE: 3, IN: 4, OUT: 5, RAW_MIN: 6, VALID_MIN: 7, STATUS: 8, VERIFY: 9, IP: 10, DEVICE: 11, CHALLENGE: 12, NOTE: 13, EDITOR: 14, CREATED: 15, UPDATED: 16 },
  CHALLENGES: { ID: 1, EMAIL: 2, ACTION: 3, CREATED: 4, EXPIRES: 5, USED: 6, USED_AT: 7, DEVICE: 8, IP: 9 }
});

function doGet(e) {
  ensureSystemInitialized_();
  let flash = null;
  if (e && e.parameter && e.parameter.proof) {
    try {
      flash = consumePresenceProof_(e.parameter.proof);
    } catch (err) {
      flash = { success: false, message: err.message || String(err) };
    }
  }
  const template = HtmlService.createTemplateFromFile('Index');
  template.initialFlash = JSON.stringify(flash || null);
  template.accountChooserUrl = JSON.stringify(getAccountChooserUrl_());
  return template
    .evaluate()
    .setTitle('LID 301 · Control de asistencia')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getBootstrapData(weekStartIso) {
  ensureSystemInitialized_();
  const email = requireUser_();
  const config = getConfig_();
  const admin = isAdmin_(email, config);
  const teacher = getTeacherByEmail_(email);
  const student = getStudentByEmail_(email);
  const isTeacher = !!(teacher && teacher.status === 'ACTIVO');
  const isStudent = !!(student && student.status === 'ACTIVO');

  if (!admin && !isTeacher && !isStudent) {
    throw new Error('Tu cuenta UPCH no está registrada como administrador, docente o estudiante activo del LID 301.');
  }

  const weekStart = normalizeWeekStart_(weekStartIso);
  return {
    app: {
      labName: config.LAB_NAME || 'LID 301',
      minimumWeeklyMinutes: Number(config.MIN_WEEKLY_MINUTES || 300),
      calendarStartHour: Number(config.CALENDAR_START_HOUR || 7),
      calendarEndHour: Number(config.CALENDAR_END_HOUR || 21),
      slotMinutes: Number(config.SLOT_MINUTES || 30),
      timezone: config.TIMEZONE || 'America/Lima',
      weekStart,
      accountChooserUrl: getAccountChooserUrl_(),
      rawWebAppUrl: ScriptApp.getService().getUrl(),
      publicFrontendUrl: String(config.PUBLIC_FRONTEND_URL || '')
    },
    user: {
      email,
      isAdmin: admin,
      isTeacher,
      isStudent,
      teacher: teacher || null,
      student: student || null
    },
    mine: isStudent ? getMyWeekData_(email, weekStart) : null
  };
}

function getMyDashboard(weekStartIso) {
  const email = requireActiveStudent_();
  return getMyWeekData_(email, normalizeWeekStart_(weekStartIso));
}

function saveMySchedule(weekStartIso, blocks) {
  const email = requireActiveStudent_();
  const config = getConfig_();
  const weekStart = normalizeWeekStart_(weekStartIso);
  const normalized = validateScheduleBlocks_(weekStart, blocks || []);
  const total = normalized.reduce((sum, b) => sum + b.minutes, 0);
  const student = getStudentByEmail_(email);
  const minimum = student ? Math.round(Number(student.minHours || 5) * 60) : Number(config.MIN_WEEKLY_MINUTES || 300);
  if (total < minimum) {
    throw new Error('Debes programar como mínimo ' + formatMinutes_(minimum) + ' por semana.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = ss_().getSheetByName(SHEETS.SCHEDULES);
    const values = sheet.getDataRange().getValues();
    const now = new Date();
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      if (lower_(row[COL.SCHEDULES.EMAIL - 1]) === email && String(row[COL.SCHEDULES.WEEK_START - 1]) === weekStart && String(row[COL.SCHEDULES.STATUS - 1]) === 'CONFIRMADO') {
        sheet.getRange(i + 1, COL.SCHEDULES.STATUS).setValue('CANCELADO');
        sheet.getRange(i + 1, COL.SCHEDULES.UPDATED).setValue(now);
      }
    }

    if (normalized.length) {
      const rows = normalized.map(b => [
        Utilities.getUuid(), email, weekStart, b.date, b.dayName, b.start, b.end, b.minutes,
        'CONFIRMADO', now, now, b.note || ''
      ]);
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    }
    audit_(email, 'GUARDAR_HORARIO', 'SEMANA', weekStart, JSON.stringify({ blocks: normalized.length, totalMinutes: total }), 'WEB');
  } finally {
    lock.releaseLock();
  }
  return getMyWeekData_(email, weekStart);
}

function copyPreviousWeek(weekStartIso) {
  const email = requireActiveStudent_();
  const weekStart = normalizeWeekStart_(weekStartIso);
  const prev = addDaysIso_(weekStart, -7);
  const sheet = ss_().getSheetByName(SHEETS.SCHEDULES);
  const rows = sheet.getDataRange().getValues();
  const previous = rows.slice(1).filter(r => lower_(r[COL.SCHEDULES.EMAIL - 1]) === email && String(r[COL.SCHEDULES.WEEK_START - 1]) === prev && String(r[COL.SCHEDULES.STATUS - 1]) === 'CONFIRMADO');
  if (!previous.length) throw new Error('No hay un horario confirmado en la semana anterior.');
  const blocks = previous.map(r => ({
    date: addDaysIso_(String(r[COL.SCHEDULES.DATE - 1]), 7),
    start: String(r[COL.SCHEDULES.START - 1]),
    end: String(r[COL.SCHEDULES.END - 1]),
    note: String(r[COL.SCHEDULES.NOTE - 1] || '')
  }));
  return saveMySchedule(weekStart, blocks);
}

function createPresenceChallenge(action) {
  const email = requireActiveStudent_();
  const config = getConfig_();
  const normalizedAction = String(action || '').toUpperCase();
  if (!['ENTRY', 'EXIT'].includes(normalizedAction)) throw new Error('Acción de presencia inválida.');

  const open = findOpenAttendance_(email);
  if (normalizedAction === 'ENTRY' && open) throw new Error('Ya tienes una entrada abierta. Marca salida antes de volver a entrar.');
  if (normalizedAction === 'EXIT' && !open) throw new Error('No tienes una entrada abierta para cerrar.');

  const nowSec = Math.floor(Date.now() / 1000);
  const ttl = Number(config.CHALLENGE_TTL_SECONDS || 60);
  const payload = {
    v: 1,
    cid: Utilities.getUuid(),
    email,
    action: normalizedAction,
    iat: nowSec,
    exp: nowSec + ttl
  };
  const token = signObject_(payload, getSharedSecret_());
  const sheet = ss_().getSheetByName(SHEETS.CHALLENGES);
  sheet.appendRow([
    payload.cid,
    email,
    normalizedAction,
    new Date(payload.iat * 1000),
    new Date(payload.exp * 1000),
    false,
    '',
    '',
    ''
  ]);
  audit_(email, 'CREAR_CHALLENGE', 'CHALLENGE', payload.cid, normalizedAction, 'WEB');

  const base = String(config.RPI_BASE_URL || '').replace(/\/$/, '');
  if (!/^http:\/\//i.test(base)) throw new Error('RPI_BASE_URL debe apuntar a la Raspberry por HTTP local, por ejemplo http://192.168.50.2:8080.');
  return {
    url: base + '/verify?token=' + encodeURIComponent(token),
    expiresAt: new Date(payload.exp * 1000).toISOString(),
    challengeId: payload.cid
  };
}

function getAdminWeek(weekStartIso, teacherFilter, projectFilter) {
  const viewer = requireManager_();
  const weekStart = normalizeWeekStart_(weekStartIso);
  const weekEnd = addDaysIso_(weekStart, 6);
  const teachers = listTeachers_();
  const allProjects = listProjects_();
  const allStudents = listStudents_();
  const selectedTeacherEmail = viewer.isAdmin ? lower_(teacherFilter || '') : viewer.email;
  const visibleProjects = viewer.isAdmin ? allProjects : allProjects.filter(p => p.teacherEmail === viewer.email);
  let selectedProjectId = String(projectFilter || '').trim();
  if (selectedProjectId && !visibleProjects.some(p => p.id === selectedProjectId)) selectedProjectId = '';

  const students = allStudents.filter(st => {
    if (selectedTeacherEmail && st.teacherEmail !== selectedTeacherEmail) return false;
    if (selectedProjectId && st.projectId !== selectedProjectId) return false;
    return true;
  });
  const allowed = new Set(students.map(st => st.email));
  const schedules = listSchedulesForWeek_(weekStart).filter(x => allowed.has(x.email));
  const attendance = listAttendanceForWeek_(weekStart, weekEnd).filter(x => allowed.has(x.email));
  const open = listOpenAttendance_().filter(x => allowed.has(x.email));

  const summary = students.filter(st => st.status === 'ACTIVO').map(st => {
    const planned = schedules.filter(x => x.email === st.email && x.status === 'CONFIRMADO').reduce((a, b) => a + b.minutes, 0);
    const valid = attendance.filter(x => x.email === st.email).reduce((a, b) => a + b.validMinutes, 0);
    const openSession = open.find(x => x.email === st.email) || null;
    const teacher = teachers.find(t => t.email === st.teacherEmail) || null;
    const project = allProjects.find(p => p.id === st.projectId) || null;
    return {
      email: st.email,
      name: fullName_(st),
      career: st.career,
      projectId: st.projectId,
      projectAlias: project ? project.alias : '',
      projectSidisi: project ? project.sidisi : '',
      teacherEmail: st.teacherEmail,
      teacherName: teacher ? teacher.name : st.teacherEmail,
      plannedMinutes: planned,
      validMinutes: valid,
      minimumMinutes: Math.round(Number(st.minHours || 5) * 60),
      presentNow: !!openSession,
      openSince: openSession ? openSession.entry : null
    };
  });

  const teacherSummary = teachers.filter(t => t.status === 'ACTIVO').map(t => {
    const mine = allStudents.filter(st => st.status === 'ACTIVO' && st.teacherEmail === t.email);
    const mineSet = new Set(mine.map(st => st.email));
    const mineAttendance = listAttendanceForWeek_(weekStart, weekEnd).filter(x => mineSet.has(x.email));
    const mineSchedules = listSchedulesForWeek_(weekStart).filter(x => mineSet.has(x.email) && x.status === 'CONFIRMADO');
    return {
      email: t.email,
      name: t.name,
      phone: t.phone,
      studentCount: mine.length,
      plannedMinutes: mineSchedules.reduce((a,b)=>a+b.minutes,0),
      validMinutes: mineAttendance.reduce((a,b)=>a+b.validMinutes,0),
      compliantCount: mine.filter(st => {
        const valid = mineAttendance.filter(x => x.email === st.email).reduce((a,b)=>a+b.validMinutes,0);
        return valid >= Math.round(Number(st.minHours || 5) * 60);
      }).length
    };
  });

  audit_(viewer.email, viewer.isAdmin ? 'VER_DASHBOARD_ADMIN' : 'VER_DASHBOARD_DOCENTE', 'SEMANA', weekStart, JSON.stringify({ teacher: selectedTeacherEmail, project: selectedProjectId }), 'WEB');
  return {
    weekStart,
    viewer,
    selectedTeacherEmail,
    selectedProjectId,
    teachers,
    projects: visibleProjects,
    students,
    schedules,
    attendance,
    open,
    summary,
    teacherSummary
  };
}

function adminListStudents() {
  requireAdmin_();
  return listStudents_();
}

function managerUpsertStudent(student) {
  const viewer = requireManager_();
  const config = getConfig_();
  const email = lower_(student && student.email);
  const requestedTeacherEmail = lower_(student && student.teacherEmail);
  const teacherEmail = viewer.isAdmin ? requestedTeacherEmail : viewer.email;
  const projectId = String(student && student.projectId || '').trim();

  if (!email || !email.endsWith('@' + lower_(config.ALLOWED_DOMAIN || 'upch.pe'))) {
    throw new Error('El correo del estudiante debe pertenecer al dominio institucional ' + (config.ALLOWED_DOMAIN || 'upch.pe') + '.');
  }

  const teacher = getTeacherByEmail_(teacherEmail);
  if (!teacher || teacher.status !== 'ACTIVO') {
    throw new Error(viewer.isAdmin ? 'Selecciona un docente a cargo activo.' : 'Tu cuenta debe estar registrada como docente activo.');
  }

  if (projectId) {
    const project = getProjectById_(projectId);
    if (!project || project.status !== 'ACTIVO') throw new Error('Selecciona un proyecto activo.');
    if (project.teacherEmail !== teacherEmail) throw new Error('El proyecto seleccionado pertenece a otro docente.');
  }

  const status = String(student && student.status || 'ACTIVO').toUpperCase();
  if (!['ACTIVO', 'INACTIVO'].includes(status)) throw new Error('Estado inválido.');
  const minHours = Number(student && student.minHours || 5);
  if (!(minHours > 0 && minHours <= 80)) throw new Error('Horas mínimas inválidas.');

  const sheet = ss_().getSheetByName(SHEETS.STUDENTS);
  const rows = sheet.getDataRange().getValues();
  const now = new Date();
  let rowIndex = -1;
  let existing = null;

  for (let i = 1; i < rows.length; i++) {
    if (lower_(rows[i][COL.STUDENTS.EMAIL - 1]) === email) {
      rowIndex = i + 1;
      existing = rows[i];
      break;
    }
  }

  if (existing && !viewer.isAdmin) {
    const existingTeacher = lower_(existing[COL.STUDENTS.TEACHER_EMAIL - 1]);
    const existingStatus = String(existing[COL.STUDENTS.STATUS - 1] || '').toUpperCase();
    if (existingTeacher !== viewer.email) {
      throw new Error('Ese estudiante ya está asignado a otro docente. Solo un administrador puede reasignarlo.');
    }
    if (existingStatus === 'ELIMINADO') {
      throw new Error('Ese estudiante fue eliminado. Solo un administrador puede reactivarlo.');
    }
  }

  const row = [
    rowIndex > 0 ? rows[rowIndex - 1][COL.STUDENTS.ID - 1] : Utilities.getUuid(),
    teacherEmail,
    String(student && student.firstName || '').trim(),
    String(student && student.lastName || '').trim(),
    email,
    String(student && student.career || '').trim(),
    projectId,
    status,
    minHours,
    'ESTUDIANTE',
    rowIndex > 0 ? rows[rowIndex - 1][COL.STUDENTS.CREATED - 1] || now : now,
    now
  ];

  if (rowIndex > 0) sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
  else sheet.appendRow(row);

  audit_(viewer.email, rowIndex > 0 ? 'EDITAR_ESTUDIANTE' : 'AGREGAR_ESTUDIANTE', 'ESTUDIANTE', email,
    JSON.stringify({ status, minHours, teacherEmail, projectId }), 'WEB');

  const all = listStudents_();
  return viewer.isAdmin ? all : all.filter(st => st.teacherEmail === viewer.email);
}

function adminUpsertStudent(student) {
  requireAdmin_();
  return managerUpsertStudent(student);
}

function managerListProjects() {
  const viewer = requireManager_();
  const projects = listProjects_();
  return viewer.isAdmin ? projects : projects.filter(p => p.teacherEmail === viewer.email);
}

function managerUpsertProject(project) {
  const viewer = requireManager_();
  const id = String(project && project.id || '').trim();
  const requestedTeacherEmail = lower_(project && project.teacherEmail);
  const teacherEmail = viewer.isAdmin ? (requestedTeacherEmail || (viewer.isTeacher ? viewer.email : '')) : viewer.email;
  const teacher = getTeacherByEmail_(teacherEmail);
  if (!teacher || teacher.status !== 'ACTIVO') throw new Error('Selecciona un docente responsable activo.');

  const pendingSidisi = toBool_(project && project.pendingSidisi);
  const sidisi = pendingSidisi ? 'POR REGISTRAR' : String(project && project.sidisi || '').trim();
  if (!sidisi) throw new Error('Ingresa el Código SIDISI o marca “Por registrar”.');
  const title = String(project && project.title || '').trim();
  const alias = String(project && project.alias || '').trim();
  if (!title) throw new Error('Ingresa el título completo del proyecto.');
  if (!alias) throw new Error('Ingresa un alias corto para el proyecto.');

  const financed = String(project && project.financed || 'NO').toUpperCase() === 'SI' ? 'SI' : 'NO';
  const funder = financed === 'SI' ? String(project && project.funder || '').trim() : '';
  const amount = financed === 'SI' ? String(project && project.amount || '').trim() : '';
  if (financed === 'SI' && !funder) throw new Error('Indica el financiador.');
  if (financed === 'SI' && !amount) throw new Error('Indica el monto financiado, incluyendo la moneda.');

  const sheet = ss_().getSheetByName(SHEETS.PROJECTS);
  const rows = sheet.getDataRange().getValues();
  let rowIndex = -1;
  let existing = null;
  if (id) {
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][COL.PROJECTS.ID - 1] || '') === id) { rowIndex = i + 1; existing = rows[i]; break; }
    }
    if (rowIndex < 0) throw new Error('Proyecto no encontrado.');
    const existingTeacher = lower_(existing[COL.PROJECTS.TEACHER_EMAIL - 1]);
    if (!viewer.isAdmin && existingTeacher !== viewer.email) throw new Error('No puedes editar proyectos de otro docente.');
    if (viewer.isAdmin && existingTeacher !== teacherEmail) {
      const assigned = listStudents_().filter(st => st.projectId === id);
      if (assigned.length) throw new Error('No puedes cambiar el docente responsable mientras el proyecto tenga estudiantes asignados. Reasigna primero a esos estudiantes.');
    }
  }

  const finalId = id || Utilities.getUuid();
  const row = [finalId, sidisi, title, alias, financed, funder, amount, teacherEmail, 'ACTIVO'];
  if (rowIndex > 0) sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
  else sheet.appendRow(row);

  audit_(viewer.email, rowIndex > 0 ? 'EDITAR_PROYECTO' : 'AGREGAR_PROYECTO', 'PROYECTO', finalId, JSON.stringify({ sidisi, alias, teacherEmail, financed, funder, amount }), 'WEB');
  return managerListProjects();
}

function managerDeleteProject(projectId) {
  const viewer = requireManager_();
  const id = String(projectId || '').trim();
  const project = getProjectById_(id);
  if (!project) throw new Error('Proyecto no encontrado.');
  if (!viewer.isAdmin && project.teacherEmail !== viewer.email) throw new Error('No puedes eliminar proyectos de otro docente.');

  const assigned = listStudents_().filter(st => st.projectId === id);
  if (assigned.length) {
    throw new Error('No puedes eliminar este proyecto porque tiene ' + assigned.length + ' estudiante(s) asignado(s). Reasígnalos primero.');
  }

  const sheet = ss_().getSheetByName(SHEETS.PROJECTS);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][COL.PROJECTS.ID - 1] || '') === id) {
      sheet.getRange(i + 1, COL.PROJECTS.STATUS).setValue('ELIMINADO');
      audit_(viewer.email, 'ELIMINAR_PROYECTO', 'PROYECTO', id, JSON.stringify({ alias: project.alias, teacherEmail: project.teacherEmail }), 'WEB');
      return managerListProjects();
    }
  }
  throw new Error('Proyecto no encontrado.');
}

function adminListTeachers() {
  requireAdmin_();
  return listTeachers_();
}

function adminDeleteStudent(email) {
  const admin = requireAdmin_();
  const target = lower_(email);
  if (!target) throw new Error('Correo de estudiante inválido.');

  const sheet = ss_().getSheetByName(SHEETS.STUDENTS);
  const rows = sheet.getDataRange().getValues();
  let rowIndex = -1;
  let snapshot = null;
  for (let i = 1; i < rows.length; i++) {
    if (lower_(rows[i][COL.STUDENTS.EMAIL - 1]) === target) {
      rowIndex = i + 1;
      snapshot = rows[i];
      break;
    }
  }
  if (rowIndex < 0) throw new Error('Estudiante no encontrado.');
  if (String(snapshot[COL.STUDENTS.STATUS - 1] || '').toUpperCase() === 'ELIMINADO') {
    throw new Error('El estudiante ya fue eliminado.');
  }

  const now = new Date();
  sheet.getRange(rowIndex, COL.STUDENTS.STATUS).setValue('ELIMINADO');
  sheet.getRange(rowIndex, COL.STUDENTS.UPDATED).setValue(now);
  audit_(admin, 'ELIMINAR_ESTUDIANTE', 'ESTUDIANTE', target, JSON.stringify({
    nombre: [snapshot[COL.STUDENTS.FIRST - 1], snapshot[COL.STUDENTS.LAST - 1]].filter(Boolean).join(' '),
    docente: lower_(snapshot[COL.STUDENTS.TEACHER_EMAIL - 1])
  }), 'WEB');
  return listStudents_();
}

function adminDeleteTeacher(email) {
  const admin = requireAdmin_();
  const target = lower_(email);
  if (!target) throw new Error('Correo de docente inválido.');

  const assigned = listStudents_().filter(st => st.teacherEmail === target);
  if (assigned.length) {
    throw new Error('No puedes eliminar este docente porque todavía tiene ' + assigned.length + ' estudiante(s) asignado(s). Reasígnalos primero.');
  }
  const projects = listProjects_().filter(p => p.teacherEmail === target);
  if (projects.length) {
    throw new Error('No puedes eliminar este docente porque todavía tiene ' + projects.length + ' proyecto(s) registrado(s). Reasígnalos o elimínalos primero.');
  }

  const sheet = ss_().getSheetByName(SHEETS.TEACHERS);
  const rows = sheet.getDataRange().getValues();
  let rowIndex = -1;
  let snapshot = null;
  for (let i = 1; i < rows.length; i++) {
    if (lower_(rows[i][COL.TEACHERS.EMAIL - 1]) === target) {
      rowIndex = i + 1;
      snapshot = rows[i];
      break;
    }
  }
  if (rowIndex < 0) throw new Error('Docente no encontrado.');
  if (String(snapshot[COL.TEACHERS.STATUS - 1] || '').toUpperCase() === 'ELIMINADO') {
    throw new Error('El docente ya fue eliminado.');
  }

  const now = new Date();
  sheet.getRange(rowIndex, COL.TEACHERS.STATUS).setValue('ELIMINADO');
  sheet.getRange(rowIndex, COL.TEACHERS.UPDATED).setValue(now);
  audit_(admin, 'ELIMINAR_DOCENTE', 'DOCENTE', target, JSON.stringify({
    nombre: String(snapshot[COL.TEACHERS.NAME - 1] || '')
  }), 'WEB');
  return listTeachers_();
}

function adminUpsertTeacher(teacher) {
  const admin = requireAdmin_();
  const config = getConfig_();
  const email = lower_(teacher && teacher.email);
  if (!email || !email.endsWith('@' + lower_(config.ALLOWED_DOMAIN || 'upch.pe'))) {
    throw new Error('El correo del docente debe pertenecer al dominio institucional ' + (config.ALLOWED_DOMAIN || 'upch.pe') + '.');
  }
  const name = String(teacher.name || '').trim();
  if (!name) throw new Error('Ingresa el nombre completo del docente.');
  const phone = String(teacher.phone || '').trim();
  const status = String(teacher.status || 'ACTIVO').toUpperCase();
  if (!['ACTIVO', 'INACTIVO'].includes(status)) throw new Error('Estado inválido.');

  const sheet = ss_().getSheetByName(SHEETS.TEACHERS);
  const rows = sheet.getDataRange().getValues();
  const now = new Date();
  let rowIndex = -1;
  for (let i = 1; i < rows.length; i++) {
    if (lower_(rows[i][COL.TEACHERS.EMAIL - 1]) === email) { rowIndex = i + 1; break; }
  }
  const row = [
    rowIndex > 0 ? rows[rowIndex - 1][COL.TEACHERS.ID - 1] : Utilities.getUuid(),
    name,
    email,
    phone,
    status,
    rowIndex > 0 ? rows[rowIndex - 1][COL.TEACHERS.CREATED - 1] || now : now,
    now
  ];
  if (rowIndex > 0) sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
  else sheet.appendRow(row);
  audit_(admin, rowIndex > 0 ? 'EDITAR_DOCENTE' : 'AGREGAR_DOCENTE', 'DOCENTE', email, JSON.stringify({ status, phone }), 'WEB');
  return listTeachers_();
}

function adminCorrectAttendance(id, entryIso, exitIso, note) {
  const admin = requireAdmin_();
  const sheet = ss_().getSheetByName(SHEETS.ATTENDANCE);
  const rows = sheet.getDataRange().getValues();
  let rowIndex = -1;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][COL.ATTENDANCE.ID - 1]) === String(id)) { rowIndex = i + 1; break; }
  }
  if (rowIndex < 0) throw new Error('Registro de asistencia no encontrado.');
  const entry = parseIsoDateTime_(entryIso);
  const exit = parseIsoDateTime_(exitIso);
  if (!entry || !exit || exit <= entry) throw new Error('Entrada/salida inválidas.');
  const config = getConfig_();
  const rawMinutes = Math.max(0, Math.round((exit.getTime() - entry.getTime()) / 60000));
  const validMinutes = calculateValidMinutes_(rawMinutes, config);
  sheet.getRange(rowIndex, COL.ATTENDANCE.DATE).setValue(dateOnly_(entry));
  sheet.getRange(rowIndex, COL.ATTENDANCE.IN).setValue(entry);
  sheet.getRange(rowIndex, COL.ATTENDANCE.OUT).setValue(exit);
  sheet.getRange(rowIndex, COL.ATTENDANCE.RAW_MIN).setValue(rawMinutes);
  sheet.getRange(rowIndex, COL.ATTENDANCE.VALID_MIN).setValue(validMinutes);
  sheet.getRange(rowIndex, COL.ATTENDANCE.STATUS).setValue('CORREGIDA');
  sheet.getRange(rowIndex, COL.ATTENDANCE.NOTE).setValue(String(note || 'Corrección administrativa'));
  sheet.getRange(rowIndex, COL.ATTENDANCE.EDITOR).setValue(admin);
  sheet.getRange(rowIndex, COL.ATTENDANCE.UPDATED).setValue(new Date());
  audit_(admin, 'CORREGIR_ASISTENCIA', 'ASISTENCIA', id, JSON.stringify({ entryIso, exitIso, validMinutes }), 'WEB');
  return true;
}

function getRaspberrySetup() {
  const admin = requireAdmin_();
  ensureSystemInitialized_();
  const config = getConfig_();
  const secret = getSharedSecret_();
  return {
    webAppUrl: ScriptApp.getService().getUrl(),
    accountChooserUrl: getAccountChooserUrl_(),
    sharedSecret: secret,
    deviceId: config.RPI_DEVICE_ID || 'LID301-RPI-01',
    raspberryBaseUrl: config.RPI_BASE_URL || 'http://192.168.50.2:8080',
    allowedSubnetHint: '192.168.50.0/24',
    interfaceName: 'wlan0',
    adminEmail: admin
  };
}

function adminUpdateRaspberryBaseUrl(url) {
  const admin = requireAdmin_();
  const value = String(url || '').trim().replace(/\/$/, '');
  if (!/^http:\/\/(?:\d{1,3}\.){3}\d{1,3}:\d+$/.test(value) && !/^http:\/\/[a-z0-9.-]+:\d+$/i.test(value)) {
    throw new Error('Usa una URL local HTTP, por ejemplo http://192.168.50.2:8080.');
  }
  const sheet = ss_().getSheetByName(SHEETS.CONFIG);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === 'RPI_BASE_URL') {
      sheet.getRange(i + 1, 2).setValue(value);
      audit_(admin, 'ACTUALIZAR_RPI_URL', 'CONFIG', 'RPI_BASE_URL', value, 'WEB');
      return getRaspberrySetup();
    }
  }
  throw new Error('No se encontró RPI_BASE_URL en CONFIG.');
}

function consumePresenceProof_(proofToken) {
  const email = requireActiveStudent_();
  const config = getConfig_();
  const proof = verifySignedObject_(proofToken, getSharedSecret_());
  if (!proof || proof.v !== 1) throw new Error('Prueba de presencia inválida.');
  if (lower_(proof.email) !== email) throw new Error('La validación corresponde a otra cuenta.');
  if (String(proof.deviceId) !== String(config.RPI_DEVICE_ID || 'LID301-RPI-01')) throw new Error('Raspberry no reconocida.');
  if (!['ENTRY', 'EXIT'].includes(String(proof.action))) throw new Error('Acción inválida en la prueba de presencia.');
  const verifiedAt = Number(proof.verifiedAt || 0);
  if (Math.abs(Math.floor(Date.now() / 1000) - verifiedAt) > Number(config.CHALLENGE_TTL_SECONDS || 60) + 30) {
    throw new Error('La prueba de presencia ha expirado. Inténtalo nuevamente.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const challengeSheet = ss_().getSheetByName(SHEETS.CHALLENGES);
    const challenges = challengeSheet.getDataRange().getValues();
    let challengeRow = -1;
    let challenge = null;
    for (let i = challenges.length - 1; i >= 1; i--) {
      if (String(challenges[i][COL.CHALLENGES.ID - 1]) === String(proof.cid)) {
        challengeRow = i + 1; challenge = challenges[i]; break;
      }
    }
    if (!challenge) throw new Error('Challenge no encontrado.');
    if (lower_(challenge[COL.CHALLENGES.EMAIL - 1]) !== email) throw new Error('Challenge asociado a otra cuenta.');
    if (String(challenge[COL.CHALLENGES.ACTION - 1]) !== String(proof.action)) throw new Error('La acción del challenge no coincide.');
    if (toBool_(challenge[COL.CHALLENGES.USED - 1])) throw new Error('Este challenge ya fue utilizado.');
    const expires = new Date(challenge[COL.CHALLENGES.EXPIRES - 1]);
    if (expires.getTime() < Date.now()) throw new Error('El challenge ha expirado.');

    challengeSheet.getRange(challengeRow, COL.CHALLENGES.USED).setValue(true);
    challengeSheet.getRange(challengeRow, COL.CHALLENGES.USED_AT).setValue(new Date());
    challengeSheet.getRange(challengeRow, COL.CHALLENGES.DEVICE).setValue(String(proof.deviceId));
    challengeSheet.getRange(challengeRow, COL.CHALLENGES.IP).setValue(String(proof.clientIp || ''));

    const attendanceSheet = ss_().getSheetByName(SHEETS.ATTENDANCE);
    const now = new Date();
    if (proof.action === 'ENTRY') {
      if (findOpenAttendance_(email)) throw new Error('Ya existe una entrada abierta.');
      const id = Utilities.getUuid();
      attendanceSheet.appendRow([
        id,
        email,
        dateOnly_(now),
        now,
        '',
        0,
        0,
        'ABIERTA',
        'RPI_WIFI',
        String(proof.clientIp || ''),
        String(proof.deviceId),
        String(proof.cid),
        '',
        '',
        now,
        now
      ]);
      audit_(email, 'MARCAR_ENTRADA', 'ASISTENCIA', id, JSON.stringify({ ip: proof.clientIp, device: proof.deviceId }), 'RPI');
      return { success: true, message: 'Entrada registrada correctamente.', action: 'ENTRY', at: now.toISOString() };
    }

    const open = findOpenAttendance_(email);
    if (!open) throw new Error('No existe una entrada abierta para cerrar.');
    const entry = new Date(open.raw[COL.ATTENDANCE.IN - 1]);
    const rawMinutes = Math.max(0, Math.round((now.getTime() - entry.getTime()) / 60000));
    const validMinutes = calculateValidMinutes_(rawMinutes, config);
    const state = rawMinutes > Number(config.MAX_SESSION_HOURS || 12) * 60 ? 'INVALIDA' : 'CERRADA';
    attendanceSheet.getRange(open.rowIndex, COL.ATTENDANCE.OUT).setValue(now);
    attendanceSheet.getRange(open.rowIndex, COL.ATTENDANCE.RAW_MIN).setValue(rawMinutes);
    attendanceSheet.getRange(open.rowIndex, COL.ATTENDANCE.VALID_MIN).setValue(validMinutes);
    attendanceSheet.getRange(open.rowIndex, COL.ATTENDANCE.STATUS).setValue(state);
    attendanceSheet.getRange(open.rowIndex, COL.ATTENDANCE.IP).setValue(String(proof.clientIp || ''));
    attendanceSheet.getRange(open.rowIndex, COL.ATTENDANCE.DEVICE).setValue(String(proof.deviceId));
    attendanceSheet.getRange(open.rowIndex, COL.ATTENDANCE.CHALLENGE).setValue(String(proof.cid));
    attendanceSheet.getRange(open.rowIndex, COL.ATTENDANCE.UPDATED).setValue(now);
    audit_(email, 'MARCAR_SALIDA', 'ASISTENCIA', open.id, JSON.stringify({ rawMinutes, validMinutes, state }), 'RPI');
    return { success: true, message: 'Salida registrada. Tiempo válido: ' + formatMinutes_(validMinutes) + '.', action: 'EXIT', at: now.toISOString() };
  } finally {
    lock.releaseLock();
  }
}

function getMyWeekData_(email, weekStart) {
  const config = getConfig_();
  const weekEnd = addDaysIso_(weekStart, 6);
  const schedule = listSchedulesForWeek_(weekStart).filter(x => x.email === email && x.status === 'CONFIRMADO');
  const attendance = listAttendanceForWeek_(weekStart, weekEnd).filter(x => x.email === email);
  const plannedMinutes = schedule.reduce((a, b) => a + b.minutes, 0);
  const validMinutes = attendance.reduce((a, b) => a + b.validMinutes, 0);
  const open = findOpenAttendance_(email);
  return {
    weekStart,
    weekEnd,
    schedule,
    attendance,
    plannedMinutes,
    validMinutes,
    minimumMinutes: (() => {
      const student = getStudentByEmail_(email);
      return student ? Math.round(Number(student.minHours || 5) * 60) : Number(config.MIN_WEEKLY_MINUTES || 300);
    })(),
    openSession: open ? serializeAttendanceRow_(open.raw, open.rowIndex) : null
  };
}

function listStudents_() {
  const sheet = ss_().getSheetByName(SHEETS.STUDENTS);
  const rows = sheet.getDataRange().getValues();
  return rows.slice(1).filter(r => String(r[COL.STUDENTS.EMAIL - 1] || '').trim() && String(r[COL.STUDENTS.STATUS - 1] || '').toUpperCase() !== 'ELIMINADO').map(r => ({
    id: String(r[COL.STUDENTS.ID - 1] || ''),
    teacherEmail: lower_(r[COL.STUDENTS.TEACHER_EMAIL - 1]),
    firstName: String(r[COL.STUDENTS.FIRST - 1] || ''),
    lastName: String(r[COL.STUDENTS.LAST - 1] || ''),
    email: lower_(r[COL.STUDENTS.EMAIL - 1]),
    career: String(r[COL.STUDENTS.CAREER - 1] || ''),
    projectId: String(r[COL.STUDENTS.PROJECT_ID - 1] || '').trim(),
    status: String(r[COL.STUDENTS.STATUS - 1] || 'INACTIVO'),
    minHours: Number(r[COL.STUDENTS.MIN_HOURS - 1] || 5),
    role: 'ESTUDIANTE'
  }));
}

function listProjects_() {
  const sheet = ss_().getSheetByName(SHEETS.PROJECTS);
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  return rows.slice(1).filter(r => String(r[COL.PROJECTS.ID - 1] || '').trim() && String(r[COL.PROJECTS.STATUS - 1] || '').toUpperCase() !== 'ELIMINADO').map(r => ({
    id: String(r[COL.PROJECTS.ID - 1] || ''),
    sidisi: String(r[COL.PROJECTS.SIDISI - 1] || '').trim(),
    title: String(r[COL.PROJECTS.TITLE - 1] || '').trim(),
    alias: String(r[COL.PROJECTS.ALIAS - 1] || '').trim(),
    financed: String(r[COL.PROJECTS.FINANCED - 1] || 'NO').toUpperCase(),
    funder: String(r[COL.PROJECTS.FUNDER - 1] || '').trim(),
    amount: String(r[COL.PROJECTS.AMOUNT - 1] || '').trim(),
    teacherEmail: lower_(r[COL.PROJECTS.TEACHER_EMAIL - 1]),
    status: String(r[COL.PROJECTS.STATUS - 1] || 'ACTIVO').toUpperCase()
  })).sort((a,b) => a.alias.localeCompare(b.alias, 'es'));
}

function getProjectById_(id) {
  return listProjects_().find(p => p.id === String(id || '').trim()) || null;
}

function listTeachers_() {
  const sheet = ss_().getSheetByName(SHEETS.TEACHERS);
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  return rows.slice(1).filter(r => String(r[COL.TEACHERS.EMAIL - 1] || '').trim() && String(r[COL.TEACHERS.STATUS - 1] || '').toUpperCase() !== 'ELIMINADO').map(r => ({
    id: String(r[COL.TEACHERS.ID - 1] || ''),
    name: String(r[COL.TEACHERS.NAME - 1] || '').trim(),
    email: lower_(r[COL.TEACHERS.EMAIL - 1]),
    phone: String(r[COL.TEACHERS.PHONE - 1] || '').trim(),
    status: String(r[COL.TEACHERS.STATUS - 1] || 'INACTIVO')
  })).sort((a,b) => a.name.localeCompare(b.name, 'es'));
}

function getTeacherByEmail_(email) {
  return listTeachers_().find(t => t.email === lower_(email)) || null;
}

function getStudentByEmail_(email) {
  return listStudents_().find(s => s.email === lower_(email)) || null;
}

function listSchedulesForWeek_(weekStart) {
  const sheet = ss_().getSheetByName(SHEETS.SCHEDULES);
  const rows = sheet.getDataRange().getValues();
  return rows.slice(1).filter(r => String(r[COL.SCHEDULES.WEEK_START - 1]) === weekStart).map(r => ({
    id: String(r[COL.SCHEDULES.ID - 1] || ''),
    email: lower_(r[COL.SCHEDULES.EMAIL - 1]),
    weekStart: String(r[COL.SCHEDULES.WEEK_START - 1] || ''),
    date: String(r[COL.SCHEDULES.DATE - 1] || ''),
    day: String(r[COL.SCHEDULES.DAY - 1] || ''),
    start: String(r[COL.SCHEDULES.START - 1] || ''),
    end: String(r[COL.SCHEDULES.END - 1] || ''),
    minutes: Number(r[COL.SCHEDULES.MINUTES - 1] || 0),
    status: String(r[COL.SCHEDULES.STATUS - 1] || ''),
    note: String(r[COL.SCHEDULES.NOTE - 1] || '')
  }));
}

function listAttendanceForWeek_(weekStart, weekEnd) {
  const sheet = ss_().getSheetByName(SHEETS.ATTENDANCE);
  const rows = sheet.getDataRange().getValues();
  return rows.slice(1).filter(r => {
    const d = String(r[COL.ATTENDANCE.DATE - 1] || '');
    return d >= weekStart && d <= weekEnd && d;
  }).map((r, i) => serializeAttendanceRow_(r, i + 2));
}

function listOpenAttendance_() {
  const sheet = ss_().getSheetByName(SHEETS.ATTENDANCE);
  const rows = sheet.getDataRange().getValues();
  const latest = {};
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (String(r[COL.ATTENDANCE.STATUS - 1]) === 'ABIERTA' && r[COL.ATTENDANCE.IN - 1] && !r[COL.ATTENDANCE.OUT - 1]) {
      const s = serializeAttendanceRow_(r, i + 1);
      latest[s.email] = s;
    }
  }
  return Object.values(latest);
}

function findOpenAttendance_(email) {
  const sheet = ss_().getSheetByName(SHEETS.ATTENDANCE);
  const rows = sheet.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    const r = rows[i];
    if (lower_(r[COL.ATTENDANCE.EMAIL - 1]) === lower_(email) && String(r[COL.ATTENDANCE.STATUS - 1]) === 'ABIERTA' && r[COL.ATTENDANCE.IN - 1] && !r[COL.ATTENDANCE.OUT - 1]) {
      return { rowIndex: i + 1, id: String(r[COL.ATTENDANCE.ID - 1]), raw: r };
    }
  }
  return null;
}

function serializeAttendanceRow_(r, rowIndex) {
  return {
    rowIndex,
    id: String(r[COL.ATTENDANCE.ID - 1] || ''),
    email: lower_(r[COL.ATTENDANCE.EMAIL - 1]),
    date: String(r[COL.ATTENDANCE.DATE - 1] || ''),
    entry: dateTimeIso_(r[COL.ATTENDANCE.IN - 1]),
    exit: dateTimeIso_(r[COL.ATTENDANCE.OUT - 1]),
    rawMinutes: Number(r[COL.ATTENDANCE.RAW_MIN - 1] || 0),
    validMinutes: Number(r[COL.ATTENDANCE.VALID_MIN - 1] || 0),
    status: String(r[COL.ATTENDANCE.STATUS - 1] || ''),
    verification: String(r[COL.ATTENDANCE.VERIFY - 1] || ''),
    ip: String(r[COL.ATTENDANCE.IP - 1] || ''),
    deviceId: String(r[COL.ATTENDANCE.DEVICE - 1] || ''),
    note: String(r[COL.ATTENDANCE.NOTE - 1] || '')
  };
}

function validateScheduleBlocks_(weekStart, blocks) {
  if (!Array.isArray(blocks) || !blocks.length) throw new Error('Agrega al menos un bloque de horario.');
  const weekEnd = addDaysIso_(weekStart, 6);
  const result = blocks.map((b, idx) => {
    const date = String(b.date || '');
    const start = normalizeTime_(b.start);
    const end = normalizeTime_(b.end);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || date < weekStart || date > weekEnd) throw new Error('Bloque ' + (idx + 1) + ': la fecha no pertenece a la semana seleccionada.');
    if (!start || !end) throw new Error('Bloque ' + (idx + 1) + ': hora inválida.');
    const startMin = timeToMinutes_(start);
    const endMin = timeToMinutes_(end);
    if (endMin <= startMin) throw new Error('Bloque ' + (idx + 1) + ': la hora de fin debe ser posterior al inicio.');
    if (endMin - startMin > 12 * 60) throw new Error('Bloque ' + (idx + 1) + ': duración excesiva.');
    return { date, start, end, minutes: endMin - startMin, dayName: dayNameFromIso_(date), note: String(b.note || '').trim() };
  });
  result.sort((a, b) => (a.date + a.start).localeCompare(b.date + b.start));
  for (let i = 1; i < result.length; i++) {
    const a = result[i - 1], b = result[i];
    if (a.date === b.date && timeToMinutes_(b.start) < timeToMinutes_(a.end)) throw new Error('Hay bloques de horario superpuestos el ' + b.date + '.');
  }
  return result;
}

function requireUser_() {
  const email = lower_(Session.getActiveUser().getEmail());
  if (!email) throw new Error('No pude identificar tu cuenta Google institucional. Abre la Web App iniciando sesión con tu cuenta UPCH.');
  const config = getConfig_();
  const domain = lower_(config.ALLOWED_DOMAIN || 'upch.pe');
  if (!email.endsWith('@' + domain)) throw new Error('Debes ingresar con tu cuenta institucional @' + domain + '.');
  return email;
}

function requireActiveStudent_() {
  const email = requireUser_();
  const student = getStudentByEmail_(email);
  if (!student || student.status !== 'ACTIVO') throw new Error('No estás registrado como estudiante activo del LID 301.');
  return email;
}

function requireAdmin_() {
  const email = requireUser_();
  const config = getConfig_();
  if (!isAdmin_(email, config)) throw new Error('Esta acción requiere permisos de administrador.');
  return email;
}

function isAdmin_(email, config) {
  const list = String((config || getConfig_()).ADMIN_EMAILS || '').split(',').map(lower_).filter(Boolean);
  return list.includes(lower_(email));
}

function requireManager_() {
  const email = requireUser_();
  const config = getConfig_();
  const isAdmin = isAdmin_(email, config);
  const teacher = getTeacherByEmail_(email);
  const isTeacher = !!(teacher && teacher.status === 'ACTIVO');
  if (!isAdmin && !isTeacher) throw new Error('Esta vista requiere permisos de administrador o docente activo.');
  return { email, isAdmin, isTeacher, teacher: teacher || null };
}

function getAccountChooserUrl_() {
  const url = ScriptApp.getService().getUrl();
  return 'https://accounts.google.com/AccountChooser?continue=' + encodeURIComponent(url);
}

function getConfig_() {
  const sheet = ss_().getSheetByName(SHEETS.CONFIG);
  const rows = sheet.getDataRange().getValues();
  const out = {};
  for (let i = 1; i < rows.length; i++) {
    const key = String(rows[i][0] || '').trim();
    if (key) out[key] = rows[i][1];
  }
  return out;
}

function ensureSystemInitialized_() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('PRESENCE_SHARED_SECRET')) {
    const seed = Utilities.getUuid() + '|' + Date.now() + '|' + Math.random();
    const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, seed, Utilities.Charset.UTF_8);
    props.setProperty('PRESENCE_SHARED_SECRET', base64UrlBytes_(bytes));
  }
  props.setProperty('SYSTEM_VERSION', '1.4.0');
  ensureProjectsSheet_();
  maybeCleanupAudit_();
}

function ensureProjectsSheet_() {
  const ss = ss_();
  let sheet = ss.getSheetByName(SHEETS.PROJECTS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.PROJECTS, 2);
    sheet.getRange(1, 1, 1, 9).setValues([['ID','Código SIDISI','Título completo','Alias','Financiado','Financiador','Monto','Docente responsable (correo UPCH)','Estado']]);
    sheet.setFrozenRows(1);
    sheet.getRange(1,1,1,9).setFontWeight('bold');
    sheet.setColumnWidth(1, 250);
    sheet.setColumnWidth(2, 140);
    sheet.setColumnWidth(3, 420);
    sheet.setColumnWidth(4, 200);
    sheet.setColumnWidth(5, 110);
    sheet.setColumnWidth(6, 220);
    sheet.setColumnWidth(7, 150);
    sheet.setColumnWidth(8, 260);
    sheet.setColumnWidth(9, 110);
  }
}

function maybeCleanupAudit_() {
  const props = PropertiesService.getScriptProperties();
  const config = getConfig_();
  const intervalDays = Math.max(1, Number(config.AUDIT_CLEANUP_INTERVAL_DAYS || 30));
  const retentionDays = Math.max(1, Number(config.AUDIT_RETENTION_DAYS || 180));
  const nowMs = Date.now();
  const lastMs = Number(props.getProperty('AUDIT_LAST_CLEANUP_MS') || 0);
  if (lastMs && (nowMs - lastMs) < intervalDays * 86400000) return;

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1500)) return;
  try {
    const refreshedLastMs = Number(props.getProperty('AUDIT_LAST_CLEANUP_MS') || 0);
    if (refreshedLastMs && (nowMs - refreshedLastMs) < intervalDays * 86400000) return;
    const deleted = purgeOldAuditRows_(retentionDays);
    props.setProperty('AUDIT_LAST_CLEANUP_MS', String(nowMs));
    if (deleted > 0) {
      audit_('SYSTEM', 'LIMPIEZA_AUDITORIA', 'AUDITORIA', '', JSON.stringify({ retentionDays, deletedRows: deleted }), 'SISTEMA');
    }
  } finally {
    lock.releaseLock();
  }
}

function purgeOldAuditRows_(retentionDays) {
  const sheet = ss_().getSheetByName(SHEETS.AUDIT);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 0;

  const cutoffMs = Date.now() - Math.max(1, Number(retentionDays || 180)) * 86400000;
  const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const blocks = [];
  let blockStart = null;
  let blockCount = 0;

  for (let i = 0; i < values.length; i++) {
    const raw = values[i][0];
    const d = raw instanceof Date ? raw : new Date(raw);
    const stale = raw && !isNaN(d.getTime()) && d.getTime() < cutoffMs;
    const actualRow = i + 2;

    if (stale) {
      if (blockStart === null) blockStart = actualRow;
      blockCount++;
    } else if (blockStart !== null) {
      blocks.push({ start: blockStart, count: blockCount });
      blockStart = null;
      blockCount = 0;
    }
  }
  if (blockStart !== null) blocks.push({ start: blockStart, count: blockCount });

  let deleted = 0;
  for (let i = blocks.length - 1; i >= 0; i--) {
    sheet.deleteRows(blocks[i].start, blocks[i].count);
    deleted += blocks[i].count;
  }
  return deleted;
}

function runAuditCleanupNow() {
  const admin = requireAdmin_();
  const config = getConfig_();
  const retentionDays = Math.max(1, Number(config.AUDIT_RETENTION_DAYS || 180));
  const deleted = purgeOldAuditRows_(retentionDays);
  PropertiesService.getScriptProperties().setProperty('AUDIT_LAST_CLEANUP_MS', String(Date.now()));
  audit_(admin, 'LIMPIEZA_AUDITORIA_MANUAL', 'AUDITORIA', '', JSON.stringify({ retentionDays, deletedRows: deleted }), 'WEB');
  return { success: true, retentionDays, deletedRows: deleted };
}

function getSharedSecret_() {
  ensureSystemInitialized_();
  return PropertiesService.getScriptProperties().getProperty('PRESENCE_SHARED_SECRET');
}

function signObject_(obj, secret) {
  const body = base64UrlText_(JSON.stringify(obj));
  const sig = Utilities.computeHmacSha256Signature(body, secret, Utilities.Charset.UTF_8);
  return body + '.' + base64UrlBytes_(sig);
}

function verifySignedObject_(token, secret) {
  const parts = String(token || '').split('.');
  if (parts.length !== 2) throw new Error('Firma inválida.');
  const expected = base64UrlBytes_(Utilities.computeHmacSha256Signature(parts[0], secret, Utilities.Charset.UTF_8));
  if (!constantTimeEquals_(expected, parts[1])) throw new Error('La firma de presencia no es válida.');
  const json = Utilities.newBlob(Utilities.base64DecodeWebSafe(padBase64_(parts[0]))).getDataAsString('UTF-8');
  return JSON.parse(json);
}

function constantTimeEquals_(a, b) {
  a = String(a || ''); b = String(b || '');
  if (a.length !== b.length) return false;
  let diff = 0;
  for (let i = 0; i < a.length; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diff === 0;
}

function base64UrlText_(text) {
  return Utilities.base64EncodeWebSafe(text, Utilities.Charset.UTF_8).replace(/=+$/g, '');
}
function base64UrlBytes_(bytes) { return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/g, ''); }
function padBase64_(s) { while (s.length % 4) s += '='; return s; }

function audit_(actor, action, entity, entityId, detail, origin) {
  try {
    ss_().getSheetByName(SHEETS.AUDIT).appendRow([new Date(), actor, action, entity, entityId, detail || '', origin || '']);
  } catch (e) {
    console.error('Audit error', e);
  }
}

function calculateValidMinutes_(rawMinutes, config) {
  const min = Number(config.MIN_SESSION_MINUTES || 10);
  const max = Number(config.MAX_SESSION_HOURS || 12) * 60;
  if (rawMinutes < min || rawMinutes > max) return 0;
  return rawMinutes;
}

function ss_() { return SpreadsheetApp.openById(SPREADSHEET_ID); }
function lower_(v) { return String(v || '').trim().toLowerCase(); }
function toBool_(v) { return v === true || String(v).toLowerCase() === 'true'; }
function fullName_(s) { return [s.firstName, s.lastName].filter(Boolean).join(' ').trim() || s.email; }

function normalizeWeekStart_(iso) {
  if (iso && /^\d{4}-\d{2}-\d{2}$/.test(String(iso))) {
    const d = parseIsoDate_(String(iso));
    const day = (d.getUTCDay() + 6) % 7; // Monday=0
    return dateOnlyUtc_(new Date(d.getTime() - day * 86400000));
  }
  const tz = getConfig_().TIMEZONE || 'America/Lima';
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  return normalizeWeekStart_(today);
}

function parseIsoDate_(iso) {
  const p = String(iso).split('-').map(Number);
  return new Date(Date.UTC(p[0], p[1] - 1, p[2], 12, 0, 0));
}
function dateOnlyUtc_(d) { return Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd'); }
function addDaysIso_(iso, days) { return dateOnlyUtc_(new Date(parseIsoDate_(iso).getTime() + days * 86400000)); }
function dayNameFromIso_(iso) {
  const names = ['Domingo','Lunes','Martes','Miércoles','Jueves','Viernes','Sábado'];
  return names[parseIsoDate_(iso).getUTCDay()];
}
function dateOnly_(d) { return Utilities.formatDate(new Date(d), getConfig_().TIMEZONE || 'America/Lima', 'yyyy-MM-dd'); }
function dateTimeIso_(v) { return v ? new Date(v).toISOString() : null; }
function parseIsoDateTime_(v) { const d = new Date(v); return isNaN(d.getTime()) ? null : d; }
function normalizeTime_(t) {
  const m = String(t || '').match(/^(\d{1,2}):(\d{2})$/);
  if (!m) return null;
  const h = Number(m[1]), min = Number(m[2]);
  if (h < 0 || h > 23 || min < 0 || min > 59) return null;
  return String(h).padStart(2, '0') + ':' + String(min).padStart(2, '0');
}
function timeToMinutes_(t) { const p = String(t).split(':').map(Number); return p[0] * 60 + p[1]; }
function formatMinutes_(m) { const h = Math.floor(Number(m || 0) / 60); const min = Number(m || 0) % 60; return h + ' h ' + String(min).padStart(2, '0') + ' min'; }
