/**
 * Espaço Angélica — Sincronização Oficial de Duração Planilha x Google Calendar (Apps Script — Preview)
 *
 * IMPORTANTE:
 * - Este arquivo é a implementação de referência validada no ambiente de preview.
 * - NÃO contém IDs hardcoded de produção, credenciais ou dados de clientes.
 * - Para ativação em produção (somente após aprovação explícita do proprietário),
 *   configure nas Propriedades do Script:
 *     - SPREADSHEET_ID: ID oficial da planilha
 *     - GOOGLE_CALENDAR_ID: ID do calendário compartilhado
 *   e execute installCalendarDurationSyncTrigger() uma única vez.
 *
 * Contrato oficial de colunas (1-based):
 * - Aba "Agendamentos":
 *     Coluna E (5): Procedimento
 *     Coluna F (6): Data Agendada (NUNCA usar como duração!)
 *     Coluna G (7): Horário Agendado
 *     Coluna H (8): Sinal (R$)
 *     Coluna I (9): Status
 *     Coluna J (10): ID Agendamento
 *     Coluna N (14): IDs Serviços Adicionais (JSON opcional)
 *     Coluna P (16): ID Serviço Principal
 *     Coluna T (20): Google Calendar Event ID
 * - Aba "Serviços":
 *     Coluna A (1): ID
 *     Coluna B (2): Título
 *     Coluna C (3): Categoria
 *     Coluna D (4): Preço
 *     Coluna E (5): Preço Formatado
 *     Coluna F (6): Duração (ÚNICA fonte oficial da duração!)
 *     Coluna G (7): Badge
 *     Coluna H (8): Descrição
 */

var CDS_TZ = "America/Sao_Paulo";
var CDS_TZ_OFFSET = "-03:00";
var CDS_APT_TAG_PREFIX = "[ESPACO_ANGELICA_APT_ID:";

// Exceções já autorizadas pelo administrador que devem ser preservadas na sincronização.
var CDS_AUTHORIZED_EXCEPTIONS = {
  excludeAppointmentIds: ["apt-1794006950286-rec-4"],
  durationOverrideByAppointmentId: {
    "apt-1793902440508": 90
  }
};

/**
 * Instala o gatilho instalável onEdit necessário para acessar CalendarApp.
 * (Triggers simples onEdit não têm permissão para chamar CalendarApp.)
 */
function installCalendarDurationSyncTrigger() {
  var ss = _cdsOpenSpreadsheet_();
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "handleSpreadsheetEditForCalendarSync") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger("handleSpreadsheetEditForCalendarSync")
    .forSpreadsheet(ss)
    .onEdit()
    .create();
  Logger.log("Gatilho instalável handleSpreadsheetEditForCalendarSync configurado com sucesso.");
}

/**
 * Gatilho instalável onEdit:
 * - Se editar aba "Serviços" (ou "Depilação"), Coluna F (6): propaga nova duração aos agendamentos futuros.
 * - Se editar aba "Agendamentos", Coluna E (5), F (6), G (7), N (14) ou P (16): recalcula horário final
 *   usando a duração oficial atual da aba "Serviços" Coluna F e atualiza o evento existente.
 */
function handleSpreadsheetEditForCalendarSync(e) {
  if (!e || !e.range) return { status: "ignored", reason: "NO_EVENT_RANGE" };

  var range = e.range;
  var sheet = range.getSheet();
  var sheetName = _cdsNormalize_(sheet.getName());
  var startRow = range.getRow();
  var endRow = range.getLastRow();
  var startCol = range.getColumn();
  var endCol = range.getLastColumn();

  if (endRow < 2) return { status: "ignored", reason: "HEADER_ONLY" };

  var ss = sheet.getParent();
  var calendar = _cdsGetSharedCalendar_();

  // 1. Mudança de duração na aba Serviços (Coluna F = 6)
  if (sheetName === "servicos" || sheetName === "depilacao") {
    if (startCol > 6 || endCol < 6) {
      return { status: "ignored", reason: "NON_DURATION_COLUMN" };
    }
    var summaries = [];
    for (var r = Math.max(2, startRow); r <= endRow; r++) {
      summaries.push(syncFutureAppointmentsForServiceRow_(ss, calendar, sheet, r));
    }
    Logger.log(JSON.stringify({ action: "ON_SERVICES_DURATION_EDIT", summaries: summaries }));
    return { status: "success", summaries: summaries };
  }

  // 2. Mudança de Procedimento (E=5), Data (F=6), Horário (G=7), Adicionais (N=14) ou ID Serviço (P=16) em Agendamentos
  if (sheetName === "agendamentos") {
    var watched = [5, 6, 7, 14, 16];
    var intersects = false;
    for (var w = 0; w < watched.length; w++) {
      if (watched[w] >= startCol && watched[w] <= endCol) {
        intersects = true;
        break;
      }
    }
    if (!intersects) return { status: "ignored", reason: "NON_SYNC_COLUMN" };

    var results = [];
    for (var ar = Math.max(2, startRow); ar <= endRow; ar++) {
      results.push(syncAppointmentRowToCalendar_(ss, calendar, ar, { requireFuture: true, preserveStartOnUpdate: false }));
    }
    Logger.log(JSON.stringify({ action: "ON_AGENDAMENTOS_EDIT", results: results }));
    return { status: "success", results: results };
  }

  return { status: "ignored", reason: "OTHER_SHEET" };
}

/**
 * Hook programático para ser chamado quando o painel admin atualizar um serviço via Web App.
 */
function syncFutureAppointmentsForServiceDurationChange(serviceId, serviceTitle) {
  var ss = _cdsOpenSpreadsheet_();
  var calendar = _cdsGetSharedCalendar_();
  var catalog = _cdsReadServicesCatalog_(ss);
  var target = null;

  if (serviceId && catalog.byId[String(serviceId).trim()] && catalog.byId[String(serviceId).trim()].length === 1) {
    target = catalog.byId[String(serviceId).trim()][0];
  } else if (serviceTitle && catalog.byTitle[_cdsNormalize_(serviceTitle)] && catalog.byTitle[_cdsNormalize_(serviceTitle)].length === 1) {
    target = catalog.byTitle[_cdsNormalize_(serviceTitle)][0];
  }

  if (!target) {
    var msg = "Serviço não encontrado na aba Serviços (ID=" + serviceId + ", Título=" + serviceTitle + ").";
    Logger.log("ERROR: " + msg);
    return { ok: false, errorCode: "SERVICE_NOT_FOUND", message: msg };
  }
  return _cdsPropagateServiceDurationToFutureAppointments_(ss, calendar, target);
}

function syncFutureAppointmentsForServiceRow_(ss, calendar, serviceSheet, rowIndex) {
  var values = serviceSheet.getRange(rowIndex, 1, 1, 8).getValues()[0];
  var target = {
    rowIndex: rowIndex,
    id: String(values[0] || "").trim(),
    title: String(values[1] || "").trim(),
    normalizedTitle: _cdsNormalize_(values[1]),
    rawDuration: values[5],
    durationMinutes: _cdsParseDurationMinutes_(values[5])
  };
  return _cdsPropagateServiceDurationToFutureAppointments_(ss, calendar, target);
}

function _cdsPropagateServiceDurationToFutureAppointments_(ss, calendar, targetService) {
  if (!targetService || targetService.durationMinutes === null) {
    var errMsg = "Duração inválida na aba Serviços, coluna F (serviço: " + (targetService ? targetService.title : "desconhecido") + ", valor: " + (targetService ? targetService.rawDuration : "") + "). Nenhum evento foi alterado.";
    Logger.log("ERROR: " + errMsg);
    return { ok: false, errorCode: "INVALID_DURATION", updated: 0, errors: 1, message: errMsg };
  }

  var agSheet = ss.getSheetByName("Agendamentos");
  if (!agSheet || agSheet.getLastRow() < 2) {
    return { ok: true, updated: 0, unchanged: 0, skipped: 0, errors: 0, logs: [] };
  }

  var rows = agSheet.getDataRange().getValues();
  var logs = [];
  var updated = 0, unchanged = 0, skipped = 0, errors = 0;

  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    var rowServiceId = String(r[15] || "").trim();
    var rowProc = _cdsNormalize_(r[4]);
    var addIds = _cdsParseAdditionalIds_(r[13]);

    var matches =
      (targetService.id && rowServiceId === targetService.id && (!rowProc || rowProc === targetService.normalizedTitle)) ||
      (targetService.normalizedTitle && rowProc === targetService.normalizedTitle) ||
      (targetService.id && addIds.indexOf(targetService.id) >= 0);

    if (!matches) continue;

    var res = syncAppointmentRowToCalendar_(ss, calendar, i + 1, {
      requireFuture: true,
      preserveStartOnUpdate: true
    });
    logs.push(res);

    if (res.status === "UPDATED" || res.status === "CREATED") updated++;
    else if (res.status === "UNCHANGED") unchanged++;
    else if (res.status === "ERROR") errors++;
    else skipped++;
  }

  var summary = { ok: errors === 0, service: targetService.title, durationMinutes: targetService.durationMinutes, updated: updated, unchanged: unchanged, skipped: skipped, errors: errors, logs: logs };
  Logger.log(JSON.stringify(summary));
  return summary;
}

function syncAppointmentRowToCalendar_(ss, calendar, rowIndex, options) {
  options = options || {};
  var agSheet = ss.getSheetByName("Agendamentos");
  if (!agSheet || rowIndex < 2 || rowIndex > agSheet.getLastRow()) {
    return { status: "ERROR", errorCode: "ROW_NOT_FOUND", rowIndex: rowIndex, message: "Linha inválida em Agendamentos." };
  }

  var lastCol = Math.max(agSheet.getLastColumn(), 20);
  var row = agSheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0];
  var display = agSheet.getRange(rowIndex, 1, 1, lastCol).getDisplayValues()[0];

  var client = String(row[1] || "").trim();
  var procedure = String(row[4] || "").trim();
  var statusNorm = _cdsNormalize_(row[8]);
  var appointmentId = String(row[9] || "").trim();
  var additionalIds = _cdsParseAdditionalIds_(row[13]);
  var serviceId = String(row[15] || "").trim();
  var storedEventId = String(row[19] || "").trim();

  if (!client && !appointmentId && !procedure) return { status: "EMPTY_ROW", rowIndex: rowIndex };
  if (statusNorm === "cancelado" || statusNorm === "cancelada" || statusNorm === "cancelled" || statusNorm === "excluido") {
    return { status: "SKIPPED_CANCELLED", rowIndex: rowIndex, appointmentId: appointmentId };
  }
  if (statusNorm.indexOf("bloquei") >= 0 || statusNorm.indexOf("bloquead") >= 0 || statusNorm === "folga" || _cdsNormalize_(procedure).indexOf("bloquei") === 0) {
    return { status: "SKIPPED_SCHEDULE_BLOCK", rowIndex: rowIndex, appointmentId: appointmentId };
  }
  if (CDS_AUTHORIZED_EXCEPTIONS.excludeAppointmentIds.indexOf(appointmentId) >= 0 || statusNorm === "teste sistema" || appointmentId.indexOf("test-") === 0) {
    return { status: "SKIPPED_AUTHORIZED_EXCLUSION", rowIndex: rowIndex, appointmentId: appointmentId };
  }
  if (appointmentId && Object.prototype.hasOwnProperty.call(CDS_AUTHORIZED_EXCEPTIONS.durationOverrideByAppointmentId, appointmentId)) {
    return { status: "SKIPPED_AUTHORIZED_DURATION_EXCEPTION", rowIndex: rowIndex, appointmentId: appointmentId };
  }

  // Ler data (Col F) e horário inicial (Col G) — Col F NUNCA é duração!
  var dateKey = _cdsDateKey_(row[5], display[5]);
  var timeKey = _cdsTimeKey_(row[6], display[6]);
  if (!dateKey || !timeKey) {
    return { status: "ERROR", errorCode: "INVALID_DATE_OR_TIME", rowIndex: rowIndex, appointmentId: appointmentId, message: "Data ou horário inválido nas colunas F/G." };
  }

  var startDate = new Date(dateKey + "T" + timeKey + ":00" + CDS_TZ_OFFSET);
  if (options.requireFuture && startDate.getTime() <= new Date().getTime()) {
    return { status: "SKIPPED_PAST", rowIndex: rowIndex, appointmentId: appointmentId };
  }

  // Ler duração fresca exclusivamente da aba Serviços (Coluna F)
  var lookup = _cdsResolveServiceDuration_(ss, serviceId, procedure, additionalIds);
  if (!lookup.ok) {
    Logger.log("ERROR row " + rowIndex + ": " + lookup.message);
    return { status: "ERROR", errorCode: lookup.errorCode, rowIndex: rowIndex, appointmentId: appointmentId, message: lookup.message };
  }

  var endDate = new Date(startDate.getTime() + lookup.durationMinutes * 60000);
  var title = client ? client + " — " + lookup.service.title : lookup.service.title;
  var existingEvent = storedEventId ? calendar.getEventById(storedEventId) : null;

  // Guarda de idempotência secundária via tag na descrição
  if (!existingEvent && appointmentId) {
    var dayStart = new Date(dateKey + "T00:00:00" + CDS_TZ_OFFSET);
    var dayEnd = new Date(dateKey + "T23:59:59" + CDS_TZ_OFFSET);
    var dayEvents = calendar.getEvents(dayStart, dayEnd);
    var tag = CDS_APT_TAG_PREFIX + appointmentId + "]";
    for (var eIdx = 0; eIdx < dayEvents.length; eIdx++) {
      if (String(dayEvents[eIdx].getDescription() || "").indexOf(tag) >= 0) {
        existingEvent = dayEvents[eIdx];
        break;
      }
    }
  }

  if (existingEvent) {
    if (existingEvent.isAllDayEvent()) {
      return { status: "SKIPPED_ALL_DAY_EVENT", rowIndex: rowIndex, appointmentId: appointmentId, eventId: existingEvent.getId() };
    }
    var targetStart = options.preserveStartOnUpdate ? existingEvent.getStartTime() : startDate;
    var targetEnd = new Date(targetStart.getTime() + lookup.durationMinutes * 60000);

    if (!storedEventId) agSheet.getRange(rowIndex, 20).setValue(existingEvent.getId());
    if (lookup.resolvedBy === "TITLE_OVERRIDE_AFTER_PROCEDURE_EDIT" && lookup.service.id) {
      agSheet.getRange(rowIndex, 16).setValue(lookup.service.id);
    }

    if (existingEvent.getStartTime().getTime() === targetStart.getTime() &&
        existingEvent.getEndTime().getTime() === targetEnd.getTime() &&
        existingEvent.getTitle() === title) {
      return { status: "UNCHANGED", rowIndex: rowIndex, appointmentId: appointmentId, eventId: existingEvent.getId() };
    }

    existingEvent.setTime(targetStart, targetEnd);
    if (existingEvent.getTitle() !== title) existingEvent.setTitle(title);
    return {
      status: "UPDATED",
      rowIndex: rowIndex,
      appointmentId: appointmentId,
      eventId: existingEvent.getId(),
      durationMinutes: lookup.durationMinutes
    };
  }

  var desc = "Procedimento: " + lookup.service.title + (appointmentId ? "\n" + CDS_APT_TAG_PREFIX + appointmentId + "]" : "");
  var created = calendar.createEvent(title, startDate, endDate, { description: desc });
  agSheet.getRange(rowIndex, 20).setValue(created.getId());
  return {
    status: "CREATED",
    rowIndex: rowIndex,
    appointmentId: appointmentId,
    eventId: created.getId(),
    durationMinutes: lookup.durationMinutes
  };
}

function _cdsOpenSpreadsheet_() {
  var id = String(PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID") || "").trim();
  if (!id) throw new Error("Configure SPREADSHEET_ID nas Propriedades do Script antes da ativação em produção.");
  return SpreadsheetApp.openById(id);
}

function _cdsGetSharedCalendar_() {
  var calId = String(PropertiesService.getScriptProperties().getProperty("GOOGLE_CALENDAR_ID") || "").trim();
  if (!calId) throw new Error("Configure GOOGLE_CALENDAR_ID nas Propriedades do Script antes da ativação em produção.");
  var cal = CalendarApp.getCalendarById(calId);
  if (!cal) throw new Error("Google Calendar não encontrado ou sem permissão de edição.");
  return cal;
}

function _cdsReadServicesCatalog_(ss) {
  var byId = {}, byTitle = {};
  ["Serviços", "Servicos", "Depilação", "Depilacao"].forEach(function(tabName) {
    var sh = ss.getSheetByName(tabName);
    if (!sh || sh.getLastRow() < 2) return;
    var values = sh.getDataRange().getValues();
    for (var i = 1; i < values.length; i++) {
      var id = String(values[i][0] || "").trim();
      var title = String(values[i][1] || "").trim();
      if (!id && !title) continue;
      var entry = {
        rowIndex: i + 1,
        id: id,
        title: title,
        normalizedTitle: _cdsNormalize_(title),
        rawDuration: values[i][5],
        durationMinutes: _cdsParseDurationMinutes_(values[i][5])
      };
      if (id) {
        byId[id] = byId[id] || [];
        byId[id].push(entry);
      }
      if (entry.normalizedTitle) {
        byTitle[entry.normalizedTitle] = byTitle[entry.normalizedTitle] || [];
        byTitle[entry.normalizedTitle].push(entry);
      }
    }
  });
  return { byId: byId, byTitle: byTitle };
}

function _cdsResolveServiceDuration_(ss, serviceId, procedureTitle, additionalIds) {
  var catalog = _cdsReadServicesCatalog_(ss);
  var normTitle = _cdsNormalize_(procedureTitle);
  var matched = null;
  var resolvedBy = null;

  if (serviceId) {
    var idList = catalog.byId[serviceId] || [];
    if (idList.length > 1) return { ok: false, errorCode: "AMBIGUOUS_SERVICE_MATCH", message: "ID duplicado em Serviços: " + serviceId };
    if (idList.length === 1) {
      if (normTitle && idList[0].normalizedTitle && idList[0].normalizedTitle !== normTitle) {
        var tList = catalog.byTitle[normTitle] || [];
        if (tList.length === 1) { matched = tList[0]; resolvedBy = "TITLE_OVERRIDE_AFTER_PROCEDURE_EDIT"; }
        else if (tList.length > 1) return { ok: false, errorCode: "AMBIGUOUS_SERVICE_MATCH", message: "Título ambíguo em Serviços: " + procedureTitle };
        else return { ok: false, errorCode: "SERVICE_NOT_FOUND", message: "Procedimento não encontrado em Serviços: " + procedureTitle };
      } else {
        matched = idList[0]; resolvedBy = "ID";
      }
    }
  }
  if (!matched && normTitle) {
    var titleList = catalog.byTitle[normTitle] || [];
    if (titleList.length > 1) return { ok: false, errorCode: "AMBIGUOUS_SERVICE_MATCH", message: "Título ambíguo em Serviços: " + procedureTitle };
    if (titleList.length === 1) { matched = titleList[0]; resolvedBy = "TITLE"; }
  }

  if (!matched) return { ok: false, errorCode: "SERVICE_NOT_FOUND", message: "Serviço não encontrado na aba Serviços." };
  if (matched.durationMinutes === null) return { ok: false, errorCode: "INVALID_DURATION", message: "Duração inválida na coluna F da aba Serviços para " + matched.title };

  var total = matched.durationMinutes;
  for (var a = 0; a < additionalIds.length; a++) {
    var aList = catalog.byId[additionalIds[a]] || [];
    if (aList.length !== 1 || aList[0].durationMinutes === null) {
      return { ok: false, errorCode: "INVALID_DURATION", message: "Serviço adicional ausente ou sem duração válida: " + additionalIds[a] };
    }
    total += aList[0].durationMinutes;
  }
  return { ok: true, service: matched, durationMinutes: total, resolvedBy: resolvedBy };
}

function _cdsParseDurationMinutes_(value) {
  if (value === null || value === undefined || value instanceof Date) return null;
  if (typeof value === "number") return (isFinite(value) && value > 0 && value <= 1440 && Math.floor(value) === value) ? value : null;
  var raw = String(value).trim();
  if (!raw || raw.indexOf("/") >= 0 || raw.indexOf("-") >= 0 || raw.indexOf(":") >= 0) return null;
  var m = /^(?:(\d+)\s*(?:h(?:oras?)?)(?:\s*(?:e\s*)?(?:(\d{1,2})\s*(?:min(?:utos?)?|m)?)?)?|(\d+)\s*(?:min(?:utos?)?|m)?)$/i.exec(raw.toLowerCase());
  if (!m) return null;
  var total = 0;
  if (m[1] !== undefined) {
    total += Number(m[1]) * 60;
    if (m[2] !== undefined) {
      if (Number(m[2]) >= 60) return null;
      total += Number(m[2]);
    }
  } else if (m[3] !== undefined) {
    total += Number(m[3]);
  }
  return (isFinite(total) && total > 0 && total <= 1440) ? total : null;
}

function _cdsNormalize_(value) {
  return String(value || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().replace(/[^a-z0-9]+/g, " ").trim().replace(/\s+/g, " ");
}

function _cdsDateKey_(raw, displayed) {
  if (raw instanceof Date && !isNaN(raw.getTime())) return Utilities.formatDate(raw, CDS_TZ, "yyyy-MM-dd");
  var s = String(displayed || raw || "").trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  var br = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  return br ? br[3] + "-" + br[2] + "-" + br[1] : "";
}

function _cdsTimeKey_(raw, displayed) {
  if (raw instanceof Date && !isNaN(raw.getTime())) return Utilities.formatDate(raw, CDS_TZ, "HH:mm");
  var s = String(displayed || raw || "").trim();
  var m = s.match(/^(\d{1,2}):(\d{2})/);
  return m ? ("0" + m[1]).slice(-2) + ":" + m[2] : "";
}

function _cdsParseAdditionalIds_(raw) {
  try {
    var parsed = JSON.parse(String(raw || "[]"));
    return Array.isArray(parsed) ? parsed.map(String) : [];
  } catch (e) {
    return [];
  }
}
