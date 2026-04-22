// ─── CONFIG ───────────────────────────────────────────────────────────────────
const SHEET_ID   = "1MM_Lq06euusJvCS7uPdGSYbbM8TM0kBpZODTswFSDBc";
const YOUR_EMAIL = "daniel@hapd.mx";
const TEAM = [
  { name: "Samantha Castillo", email: "gteentrenamiento@fac.mx" },
  { name: "Diego Lopez",       email: "d.lopez@fac.mx" },
  { name: "Julio Flores",      email: "construccion@fac.mx" },
  { name: "Cristobal Salinas", email: "cristobal@fac.mx" },
  { name: "Erick Campuzano",   email: "chefcorporativo@fac.mx" },
  { name: "Nayeli Sanjuan",    email: "talento@fac.mx" },
  { name: "Abraham Focil",     email: "rh1@fac.mx" },
];
const DAYS = ["Lunes","Martes","Miércoles","Jueves","Viernes","Sábado","Domingo"];
const HEADERS = ["Semana","Rango","Día","Descanso","Ubicación","Horario","Actividades","Objetivos","Comentario de Cierre","Última actualización"];

// ─── MERGE HELPER — trims whitespace before deciding ─────────────────────────
function mergeVal(incoming, existing) {
  const s = String(incoming == null ? "" : incoming).trim();
  return s !== "" ? s : String(existing == null ? "" : existing);
}

// ─── RECEIVE DATA FROM APP ────────────────────────────────────────────────────
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    Logger.log(JSON.stringify(data.day));
    const ss = SpreadsheetApp.openById(SHEET_ID);
    writeMasterDay(ss, data);
    writePersonDay(ss, data);
    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: err.message })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput("Plan Semanal hap&d - OK").setMimeType(ContentService.MimeType.TEXT);
}

// ─── MASTER SHEET ─────────────────────────────────────────────────────────────
function writeMasterDay(ss, data) {
  let sheet = ss.getSheetByName("Master");
  if (!sheet) {
    sheet = ss.insertSheet("Master");
    const h = ["Timestamp","Nombre","Rol","Semana","Rango","Día","Descanso","Ubicación","Horario","Actividades","Objetivos","Comentario de Cierre"];
    sheet.appendRow(h);
    sheet.setFrozenRows(1);
    sheet.getRange(1,1,1,h.length).setFontWeight("bold").setBackground("#E8631A").setFontColor("#ffffff");
  }

  const d          = data.day;
  const ts         = Utilities.formatDate(new Date(), "America/Mexico_City", "dd/MM/yyyy HH:mm:ss");
  const weekNum    = String(data.weekNum).trim();
  const dayName    = String(d.day).trim();
  const memberName = String(data.member).trim();

  const inUbicacion   = (d.ubicacion || []).join(", ");
  const inHorario     = String(d.horario     || "");
  const inActividades = String(d.actividades || "");
  const inObjetivos   = String(d.objetivos   || "");
  const inComentario  = String(d.comentario  || "");

  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (
      String(rows[i][1]).trim() === memberName &&
      String(rows[i][3]).trim() === weekNum    &&
      String(rows[i][5]).trim() === dayName
    ) {
      sheet.getRange(i + 1, 1, 1, 12).setValues([[
        ts,
        memberName,
        data.rol,
        weekNum,
        String(data.weekRange),
        dayName,
        d.descanso ? "Sí" : "No",
        mergeVal(inUbicacion,   rows[i][7]),   // H
        mergeVal(inHorario,     rows[i][8]),   // I
        mergeVal(inActividades, rows[i][9]),   // J
        mergeVal(inObjetivos,   rows[i][10]),  // K
        mergeVal(inComentario,  rows[i][11])   // L
      ]]);
      return;
    }
  }

  sheet.appendRow([
    ts, memberName, data.rol, weekNum, String(data.weekRange), dayName,
    d.descanso ? "Sí" : "No",
    inUbicacion, inHorario, inActividades, inObjetivos, inComentario
  ]);
}

// ─── INDIVIDUAL PERSON SHEET ──────────────────────────────────────────────────
function writePersonDay(ss, data) {
  const sheetName = data.member.split(" ")[0];
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.setFrozenRows(1);
    sheet.getRange(1,1,1,HEADERS.length).setValues([HEADERS]).setFontWeight("bold").setBackground("#E8631A").setFontColor("#ffffff");
  }

  const d  = data.day;
  const ts = new Date().toLocaleString("es-MX");

  const inUbicacion   = (d.ubicacion || []).join(", ");
  const inHorario     = String(d.horario     || "");
  const inActividades = String(d.actividades || "");
  const inObjetivos   = String(d.objetivos   || "");
  const inComentario  = String(d.comentario  || "");

  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.weekNum) && rows[i][2] === d.day) {
      sheet.getRange(i + 1, 1, 1, 10).setValues([[
        data.weekNum,
        data.weekRange,
        d.day,
        d.descanso ? "✓ Descanso" : "No",
        mergeVal(inUbicacion,   rows[i][4]),   // Ubicación
        mergeVal(inHorario,     rows[i][5]),   // Horario
        mergeVal(inActividades, rows[i][6]),   // Actividades
        mergeVal(inObjetivos,   rows[i][7]),   // Objetivos
        mergeVal(inComentario,  rows[i][8]),   // Comentario
        ts
      ]]);
      return;
    }
  }

  sheet.appendRow([
    data.weekNum, data.weekRange, d.day,
    d.descanso ? "✓ Descanso" : "No",
    inUbicacion, inHorario, inActividades, inObjetivos, inComentario, ts
  ]);
}

// ─── DAILY REPORT ─────────────────────────────────────────────────────────────
function sendDailyReport() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const master = ss.getSheetByName("Master");
  if (!master) return;

  const now = new Date();
  const dayNames = ["Domingo","Lunes","Martes","Miércoles","Jueves","Viernes","Sábado"];

  const yesterday = new Date(now); yesterday.setDate(now.getDate() - 1);
  const yesterdayName = dayNames[yesterday.getDay()];
  const yesterdayWeek = getWeekNumber(yesterday);
  const todayName = dayNames[now.getDay()];
  const todayWeek = getWeekNumber(now);

  const rows = master.getDataRange().getValues();
  const yesterdayData = {};
  const todayData = {};
  TEAM.forEach(m => { yesterdayData[m.name] = null; todayData[m.name] = null; });

  const toStr = v => (v instanceof Date) ? Utilities.formatDate(v, "America/Mexico_City", "HH:mm") : String(v || "—");

  rows.slice(1).forEach(r => {
    const member = r[1];
    const semana = String(r[3]);
    const dia    = String(r[5]);
    if (semana === String(yesterdayWeek) && dia === yesterdayName)
      yesterdayData[member] = { ubicacion: toStr(r[7]), horario: toStr(r[8]), actividades: toStr(r[9]), objetivos: toStr(r[10]), comentario: r[11] ? String(r[11]) : null };
    if (semana === String(todayWeek) && dia === todayName)
      todayData[member] = { ubicacion: toStr(r[7]), horario: toStr(r[8]), actividades: toStr(r[9]), objetivos: toStr(r[10]) };
  });

  const yesterdayStr = yesterday.toLocaleDateString("es-MX", { weekday:"long", day:"numeric", month:"long" });
  const todayStr     = now.toLocaleDateString("es-MX",       { weekday:"long", day:"numeric", month:"long" });

  let html = `<div style="font-family:Arial,sans-serif;max-width:700px;margin:0 auto;">`;

  html += `
    <div style="background:#E8631A;padding:20px 24px;border-radius:8px 8px 0 0;">
      <h2 style="color:#fff;margin:0;font-size:18px;">The hap&d co. · Reporte Nocturno</h2>
      <p style="color:rgba(255,255,255,0.8);margin:4px 0 0;font-size:13px;">${todayStr}</p>
    </div>
    <div style="background:#1a1a1a;padding:10px 20px;">
      <span style="color:#E8631A;font-weight:bold;font-size:12px;text-transform:uppercase;letter-spacing:1px;">Cierre de ayer · ${yesterdayStr}</span>
    </div>
    <div style="border:1px solid #e5e0d8;border-top:none;overflow:hidden;">`;

  TEAM.forEach((member, i) => {
    const d  = yesterdayData[member.name];
    const bg = i % 2 === 0 ? "#ffffff" : "#faf8f5";
    const replyBody = d ? encodeURIComponent(
      `Hola ${member.name.split(' ')[0]},\n\nSobre tu día del ${yesterdayStr}:\n` +
      `Ubicación: ${d.ubicacion}\nHorario: ${d.horario}\nActividades: ${d.actividades}\nObjetivos: ${d.objetivos}\nCierre: ${d.comentario || 'No enviado'}\n\nMi feedback:\n`
    ) : '';
    const replyBtn = `<a href="mailto:${member.email}?subject=Feedback · ${yesterdayStr}&body=${replyBody}" style="font-size:11px;padding:3px 10px;border:1px solid #E8631A;color:#E8631A;border-radius:12px;text-decoration:none;margin-left:8px;">Responder</a>`;
    if (!d) {
      html += `<div style="padding:14px 20px;background:${bg};border-bottom:1px solid #f0ebe3;display:flex;justify-content:space-between;align-items:center;">
        <div><strong style="font-size:13px;">${member.name}</strong>${replyBtn}</div>
        <span style="background:#fee2e2;color:#991b1b;padding:2px 10px;border-radius:12px;font-size:11px;font-weight:bold;">Sin plan</span>
      </div>`;
    } else {
      const hasComment = !!d.comentario;
      html += `<div style="padding:14px 20px;background:${bg};border-bottom:1px solid #f0ebe3;">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;">
          <div><strong style="font-size:13px;">${member.name}</strong>${replyBtn}</div>
          <span style="color:${hasComment ? '#16a34a' : '#d97706'};font-size:11px;font-weight:bold;">${hasComment ? '✓ Cierre enviado' : '⛔ Sin cierre'}</span>
        </div>
        <table style="width:100%;font-size:12px;color:#555;border-collapse:collapse;">
          <tr><td style="padding:2px 8px 2px 0;color:#999;width:90px;">Ubicación</td><td>${d.ubicacion}</td></tr>
          <tr><td style="padding:2px 8px 2px 0;color:#999;">Horario</td><td>${d.horario}</td></tr>
          <tr><td style="padding:2px 8px 2px 0;color:#999;">Actividades</td><td>${d.actividades}</td></tr>
          <tr><td style="padding:2px 8px 2px 0;color:#999;">Objetivos</td><td>${d.objetivos}</td></tr>
          <tr><td style="padding:6px 8px 2px 0;color:#999;vertical-align:top;">Cierre</td>
            <td style="padding-top:6px;color:${hasComment ? '#333' : '#d97706'};font-style:${hasComment ? 'italic' : 'normal'};">
              ${hasComment ? `"${d.comentario}"` : 'No envió comentario de cierre.'}
            </td>
          </tr>
        </table>
      </div>`;
    }
  });

  html += `</div>
    <div style="background:#1a1a1a;padding:10px 20px;margin-top:16px;">
      <span style="color:#E8631A;font-weight:bold;font-size:12px;text-transform:uppercase;letter-spacing:1px;">Plan de hoy · ${todayStr}</span>
    </div>
    <div style="border:1px solid #e5e0d8;border-top:none;overflow:hidden;">`;

  TEAM.forEach((member, i) => {
    const d  = todayData[member.name];
    const bg = i % 2 === 0 ? "#ffffff" : "#faf8f5";
    const replyBodyToday = d ? encodeURIComponent(
      `Hola ${member.name.split(' ')[0]},\n\nSobre tu plan de hoy ${todayStr}:\n` +
      `Ubicación: ${d.ubicacion}\nHorario: ${d.horario}\nActividades: ${d.actividades}\nObjetivos: ${d.objetivos}\n\nMi feedback:\n`
    ) : '';
    const replyBtn = `<a href="mailto:${member.email}?subject=Plan · ${todayStr}&body=${replyBodyToday}" style="font-size:11px;padding:3px 10px;border:1px solid #E8631A;color:#E8631A;border-radius:12px;text-decoration:none;margin-left:8px;">Responder</a>`;
    if (!d) {
      html += `<div style="padding:14px 20px;background:${bg};border-bottom:1px solid #f0ebe3;display:flex;justify-content:space-between;align-items:center;">
        <div><strong style="font-size:13px;">${member.name}</strong>${replyBtn}</div>
        <span style="background:#fee2e2;color:#991b1b;padding:2px 10px;border-radius:12px;font-size:11px;font-weight:bold;">Sin plan</span>
      </div>`;
    } else {
      html += `<div style="padding:14px 20px;background:${bg};border-bottom:1px solid #f0ebe3;">
        <div style="margin-bottom:6px;"><strong style="font-size:13px;">${member.name}</strong>${replyBtn}</div>
        <table style="width:100%;font-size:12px;color:#555;border-collapse:collapse;">
          <tr><td style="padding:2px 8px 2px 0;color:#999;width:90px;">Ubicación</td><td>${d.ubicacion}</td></tr>
          <tr><td style="padding:2px 8px 2px 0;color:#999;">Horario</td><td>${d.horario}</td></tr>
          <tr><td style="padding:2px 8px 2px 0;color:#999;">Actividades</td><td>${d.actividades}</td></tr>
          <tr><td style="padding:2px 8px 2px 0;color:#999;">Objetivos</td><td>${d.objetivos}</td></tr>
        </table>
      </div>`;
    }
  });

  html += `</div>
    <p style="font-size:11px;color:#aaa;text-align:center;margin-top:12px;">Plan Semanal SC · The hap&d co. · <a href="https://plan.hapd.mx" style="color:#E8631A;">plan.hapd.mx</a></p>
  </div>`;

  GmailApp.sendEmail(YOUR_EMAIL, `Reporte SC · ${todayStr}`, "", { htmlBody: html });

  let abrahamHtml = `<div style="font-family:Arial,sans-serif;max-width:700px;margin:0 auto;">
    <div style="background:#E8631A;padding:20px 24px;border-radius:8px 8px 0 0;">
      <h2 style="color:#fff;margin:0;font-size:18px;">The hap&d co. · Plan de hoy</h2>
      <p style="color:rgba(255,255,255,0.8);margin:4px 0 0;font-size:13px;">${todayStr}</p>
    </div>
    <div style="border:1px solid #e5e0d8;border-top:none;border-radius:0 0 8px 8px;overflow:hidden;">`;

  TEAM.forEach((member, i) => {
    if (member.name === "Abraham Focil") return;
    const d  = todayData[member.name];
    const bg = i % 2 === 0 ? "#ffffff" : "#faf8f5";
    abrahamHtml += `<div style="padding:12px 20px;background:${bg};border-bottom:1px solid #f0ebe3;">
      <strong style="font-size:13px;">${member.name}</strong>
      ${!d ? `<span style="color:#991b1b;font-size:11px;margin-left:8px;">Sin plan</span>` :
      `<div style="font-size:12px;color:#555;margin-top:4px;">
        <span style="color:#999;">Ubicación:</span> ${d.ubicacion} &nbsp;·&nbsp;
        <span style="color:#999;">Horario:</span> ${d.horario}
      </div>`}
    </div>`;
  });

  abrahamHtml += `</div>
    <p style="font-size:11px;color:#aaa;text-align:center;margin-top:12px;">Plan Semanal SC · The hap&d co.</p>
  </div>`;

  GmailApp.sendEmail("rh1@fac.mx", `Plan del equipo · ${todayStr}`, "", { htmlBody: abrahamHtml });
}

// ─── WEEKLY REPORT ────────────────────────────────────────────────────────────
function sendWeeklyReport() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const master = ss.getSheetByName("Master");
  if (!master) return;

  const today      = new Date();
  const lastMonday = new Date(today);
  lastMonday.setDate(today.getDate() - 7);
  const weekNum = getWeekNumber(lastMonday);

  const rows = master.getDataRange().getValues();
  const data = {};
  TEAM.forEach(m => data[m.name] = []);

  const toStr = v => (v instanceof Date) ? Utilities.formatDate(v, "America/Mexico_City", "HH:mm") : String(v || "—");

  rows.slice(1).forEach(r => {
    if (String(r[3]) === String(weekNum)) {
      if (!data[r[1]]) data[r[1]] = [];
      data[r[1]].push({
        day: String(r[5]), descanso: r[6],
        ubicacion: toStr(r[7]), horario: toStr(r[8]),
        actividades: toStr(r[9]), objetivos: toStr(r[10]),
        comentario: r[11] ? String(r[11]) : null
      });
    }
  });

  const weekRange = getWeekRange(lastMonday);

  let html = `<div style="font-family:Arial,sans-serif;max-width:700px;margin:0 auto;">
    <div style="background:#E8631A;padding:20px 24px;border-radius:8px 8px 0 0;">
      <h2 style="color:#fff;margin:0;font-size:18px;">The hap&d co. · Resumen Semanal</h2>
      <p style="color:rgba(255,255,255,0.8);margin:4px 0 0;font-size:13px;">Semana ${weekNum} · ${weekRange}</p>
    </div>
    <div style="border:1px solid #e5e0d8;border-top:none;border-radius:0 0 8px 8px;overflow:hidden;">`;

  TEAM.forEach((member, i) => {
    const days           = data[member.name] || [];
    const bg             = i % 2 === 0 ? "#ffffff" : "#faf8f5";
    const daysWithComment  = days.filter(d => d.comentario).length;
    const totalActiveDays  = days.filter(d => d.descanso !== "Sí").length;

    html += `<div style="padding:16px 20px;background:${bg};border-bottom:1px solid #f0ebe3;">
      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;">
        <strong style="font-size:15px;">${member.name}</strong>
        <span style="font-size:11px;color:#888;">Cierres: ${daysWithComment}/${totalActiveDays} días</span>
      </div>`;

    if (days.length === 0) {
      html += `<p style="color:#d97706;font-size:12px;margin:0;">Sin datos esta semana</p>`;
    } else {
      html += `<table style="width:100%;font-size:11px;border-collapse:collapse;">
        <tr style="background:#f5f3ef;">
          <th style="padding:5px 8px;text-align:left;color:#999;font-weight:600;">Día</th>
          <th style="padding:5px 8px;text-align:left;color:#999;font-weight:600;">Ubicación</th>
          <th style="padding:5px 8px;text-align:left;color:#999;font-weight:600;">Objetivos</th>
          <th style="padding:5px 8px;text-align:left;color:#999;font-weight:600;">Cierre</th>
        </tr>`;
      days.forEach(d => {
        const isRest = d.descanso === "Sí";
        html += `<tr style="border-top:1px solid #f0ebe3;${isRest ? 'background:#f0fdf4;' : ''}">
          <td style="padding:5px 8px;font-weight:600;color:${isRest ? '#16a34a' : '#333'}">${d.day.slice(0,3)}</td>
          <td style="padding:5px 8px;color:#555;">${isRest ? 'Descanso' : d.ubicacion}</td>
          <td style="padding:5px 8px;color:#555;">${isRest ? '—' : d.objetivos}</td>
          <td style="padding:5px 8px;color:${d.comentario ? '#333' : '#d97706'};font-style:${d.comentario ? 'italic' : 'normal'};">${d.comentario ? `"${d.comentario.substring(0,80)}${d.comentario.length > 80 ? '...' : ''}"` : (isRest ? '—' : 'N/A')}</td>
        </tr>`;
      });
      html += `</table>`;
    }
    html += `</div>`;
  });

  html += `</div>
    <p style="font-size:11px;color:#aaa;text-align:center;margin-top:12px;">Plan Semanal SC · The hap&d co. · <a href="https://plan.hapd.mx" style="color:#E8631A;">plan.hapd.mx</a></p>
  </div>`;

  GmailApp.sendEmail(YOUR_EMAIL, `Resumen Semanal SC · Semana ${weekNum} · ${weekRange}`, "", { htmlBody: html });
}

// ─── SETUP TRIGGERS ───────────────────────────────────────────────────────────
function setupTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("sendDailyReport").timeBased().everyDays(1).atHour(0).create();
  ScriptApp.newTrigger("sendWeeklyReport").timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(0).create();
  Logger.log("✅ Triggers set up successfully");
}

// ─── HELPERS ──────────────────────────────────────────────────────────────────
function getWeekNumber(date) {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay() || 7));
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
}

function getWeekRange(date) {
  const monday = new Date(date);
  const day    = monday.getDay() || 7;
  monday.setDate(monday.getDate() - day + 1);
  const sunday = new Date(monday);
  sunday.setDate(monday.getDate() + 6);
  const fmt = d => d.toLocaleDateString("es-MX", { day:"numeric", month:"short" });
  return `${fmt(monday)} – ${fmt(sunday)}`;
}
