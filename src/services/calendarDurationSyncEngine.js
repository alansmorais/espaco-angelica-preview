/**
 * Espaço Angélica — Motor Oficial de Sincronização de Duração com Google Calendar
 *
 * Regras oficiais de estrutura:
 * - Aba "Agendamentos":
 *     Coluna E (idx 4): Procedimento
 *     Coluna F (idx 5): Data Agendada (NUNCA usar como duração!)
 *     Coluna G (idx 6): Horário Agendado
 *     Coluna H (idx 7): Sinal (R$)
 *     Coluna I (idx 8): Status
 *     Coluna J (idx 9): ID Agendamento
 *     Coluna N (idx 13): IDs Serviços Adicionais (JSON opcional)
 *     Coluna P (idx 15): ID Serviço Principal (opcional, preferencial quando consistente)
 *     Coluna T (idx 19): ID do Evento no Google Calendar
 *
 * - Aba "Serviços":
 *     Coluna A (idx 0): ID
 *     Coluna B (idx 1): Título
 *     Coluna C (idx 2): Categoria
 *     Coluna D (idx 3): Preço
 *     Coluna E (idx 4): Preço Formatado
 *     Coluna F (idx 5): Duração (ÚNICA fonte oficial da duração!)
 *     Coluna G (idx 6): Badge
 *     Coluna H (idx 7): Descrição
 *
 * Fórmula obrigatória:
 *   HORÁRIO FINAL = HORÁRIO INICIAL + DURAÇÃO DO SERVIÇO
 */

export const AGENDAMENTOS_COL = Object.freeze({
  TIMESTAMP: 0,
  CLIENTE: 1,
  PROCEDIMENTO: 4,      // Coluna E
  DATA_AGENDADA: 5,     // Coluna F (Data — NUNCA duração)
  HORARIO_AGENDADO: 6,  // Coluna G
  SINAL: 7,             // Coluna H
  STATUS: 8,            // Coluna I
  APPOINTMENT_ID: 9,    // Coluna J
  ADDITIONAL_IDS: 13,   // Coluna N
  SERVICE_ID: 15,       // Coluna P
  CALENDAR_EVENT_ID: 19 // Coluna T
});

export const SERVICOS_COL = Object.freeze({
  ID: 0,              // Coluna A
  TITULO: 1,          // Coluna B
  CATEGORIA: 2,       // Coluna C
  PRECO: 3,           // Coluna D
  PRECO_FORMATADO: 4, // Coluna E
  DURACAO: 5,         // Coluna F (ÚNICA fonte oficial da duração)
  BADGE: 6,           // Coluna G
  DESCRICAO: 7        // Coluna H
});

export const TIMEZONE_OFFSET = '-03:00'; // America/Sao_Paulo
export const APT_TAG_PREFIX = '[ESPACO_ANGELICA_APT_ID:';

const CANCELLED_STATUSES = new Set([
  'cancelado',
  'cancelada',
  'cancelled',
  'canceled',
  'excluido',
  'excluído'
]);

const BLOCK_MARKERS = [
  'bloqueio',
  'bloqueado',
  'bloqueada',
  'folga',
  'indisponivel',
  'agenda bloqueada'
];

/**
 * Normaliza texto para comparação insensível a acentos, caixa e espaços extras.
 */
export function normalizeText(value) {
  return String(value ?? '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .trim()
    .replace(/\s+/g, ' ');
}

/**
 * Converte a célula da Coluna F da aba "Serviços" em duração inteira em minutos.
 * Rejeita explicitamente datas, valores negativos, zero, texto inválido e vazios.
 * NUNCA retorna duração padrão.
 */
export function parseOfficialDurationMinutes(value) {
  if (value === null || value === undefined) return null;
  if (value instanceof Date) return null;

  if (typeof value === 'number') {
    if (!Number.isFinite(value) || value <= 0 || value > 1440 || !Number.isInteger(value)) {
      return null;
    }
    return value;
  }

  const raw = String(value).trim();
  if (!raw) return null;

  // Bloqueio explícito contra datas (ex: "25/11/2026", "2026-11-25") ou sinais negativos
  if (raw.includes('/') || raw.includes('-') || raw.includes(':')) {
    return null;
  }

  const lower = raw.toLowerCase();

  // Aceita apenas formatos claros de duração:
  // "45", "45 min", "45 minutos", "1h", "1 hora", "2 horas", "1h 30min", "1h30", "1h30m", "1 hora e 30 minutos"
  const fullPattern = /^(?:(\d+)\s*(?:h(?:oras?)?)(?:\s*(?:e\s*)?(?:(\d{1,2})\s*(?:min(?:utos?)?|m)?)?)?|(\d+)\s*(?:min(?:utos?)?|m)?)$/i;
  const match = fullPattern.exec(lower);
  if (!match) return null;

  let totalMinutes = 0;
  if (match[1] !== undefined) {
    totalMinutes += Number(match[1]) * 60;
    if (match[2] !== undefined) {
      const mins = Number(match[2]);
      if (mins >= 60) return null;
      totalMinutes += mins;
    }
  } else if (match[3] !== undefined) {
    totalMinutes += Number(match[3]);
  }

  if (!Number.isFinite(totalMinutes) || totalMinutes <= 0 || totalMinutes > 1440) {
    return null;
  }
  return totalMinutes;
}

/**
 * Normaliza data lida da Coluna F da aba "Agendamentos" para "yyyy-MM-dd".
 */
export function parseSheetDateKey(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    const yyyy = value.getUTCFullYear();
    const mm = String(value.getUTCMonth() + 1).padStart(2, '0');
    const dd = String(value.getUTCDate()).padStart(2, '0');
    return `${yyyy}-${mm}-${dd}`;
  }
  const s = String(value ?? '').trim();
  if (!s) return null;

  const iso = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s);
  if (iso) {
    const [, y, m, d] = iso;
    if (isValidCalendarDate(Number(y), Number(m), Number(d))) {
      return `${y}-${m}-${d}`;
    }
    return null;
  }

  const br = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(s);
  if (br) {
    const [, d, m, y] = br;
    if (isValidCalendarDate(Number(y), Number(m), Number(d))) {
      return `${y}-${m}-${d}`;
    }
    return null;
  }

  return null;
}

function isValidCalendarDate(year, month, day) {
  if (month < 1 || month > 12 || day < 1 || day > 31) return false;
  const probe = new Date(Date.UTC(year, month - 1, day));
  return (
    probe.getUTCFullYear() === year &&
    probe.getUTCMonth() === month - 1 &&
    probe.getUTCDate() === day
  );
}

/**
 * Normaliza horário lido da Coluna G da aba "Agendamentos" para "HH:mm".
 */
export function parseSheetTimeKey(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    const hh = String(value.getUTCHours()).padStart(2, '0');
    const mm = String(value.getUTCMinutes()).padStart(2, '0');
    return `${hh}:${mm}`;
  }
  const s = String(value ?? '').trim();
  const match = /^(\d{1,2}):(\d{2})(?::\d{2})?$/.exec(s);
  if (!match) return null;
  const hours = Number(match[1]);
  const minutes = Number(match[2]);
  if (hours < 0 || hours > 23 || minutes < 0 || minutes > 59) return null;
  return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}`;
}

/**
 * Calcula início e término exatos a partir de Data (Col F), Horário (Col G) e Duração (Serviços Col F).
 * HORÁRIO FINAL = HORÁRIO INICIAL + DURAÇÃO DO SERVIÇO
 */
export function calculateEventInterval(dateValue, startTimeValue, durationMinutes, tzOffset = TIMEZONE_OFFSET) {
  const dateKey = parseSheetDateKey(dateValue);
  if (!dateKey) {
    return { ok: false, errorCode: 'INVALID_DATE', message: `Data agendada inválida na coluna F: "${dateValue}"` };
  }
  const startTimeKey = parseSheetTimeKey(startTimeValue);
  if (!startTimeKey) {
    return { ok: false, errorCode: 'INVALID_START_TIME', message: `Horário inicial inválido na coluna G: "${startTimeValue}"` };
  }
  if (!Number.isInteger(durationMinutes) || durationMinutes <= 0) {
    return { ok: false, errorCode: 'INVALID_DURATION', message: `Duração inválida para cálculo do horário final: "${durationMinutes}"` };
  }

  const startIso = `${dateKey}T${startTimeKey}:00${tzOffset}`;
  const startDate = new Date(startIso);
  if (Number.isNaN(startDate.getTime())) {
    return { ok: false, errorCode: 'INVALID_INTERVAL', message: `Não foi possível construir data/hora inicial: ${startIso}` };
  }

  const endMs = startDate.getTime() + durationMinutes * 60_000;
  const endDate = new Date(endMs);

  // Calcula representação local no mesmo offset (sem depender de timezone do SO)
  const offsetMinutes = parseOffsetMinutes(tzOffset);
  const localEnd = new Date(endMs + offsetMinutes * 60_000);
  const endDateKey = `${localEnd.getUTCFullYear()}-${String(localEnd.getUTCMonth() + 1).padStart(2, '0')}-${String(localEnd.getUTCDate()).padStart(2, '0')}`;
  const endTimeKey = `${String(localEnd.getUTCHours()).padStart(2, '0')}:${String(localEnd.getUTCMinutes()).padStart(2, '0')}`;
  const endIso = `${endDateKey}T${endTimeKey}:00${tzOffset}`;

  return {
    ok: true,
    dateKey,
    startTimeKey,
    endDateKey,
    endTimeKey,
    durationMinutes,
    startIso,
    endIso,
    startDate,
    endDate
  };
}

function parseOffsetMinutes(tzOffset) {
  const m = /^([+-])(\d{2}):(\d{2})$/.exec(String(tzOffset || '-03:00'));
  if (!m) return -180;
  const sign = m[1] === '-' ? -1 : 1;
  return sign * (Number(m[2]) * 60 + Number(m[3]));
}

/**
 * Lê a aba "Serviços" diretamente (sempre fresca, sem cache antigo).
 * Mantém listas por ID (Col A) e por Título normalizado (Col B) para detectar ambiguidades.
 */
export function indexServicesFromSheet(servicosRows) {
  const byId = new Map();
  const byTitle = new Map();
  const entries = [];

  for (let i = 1; i < servicosRows.length; i++) {
    const row = servicosRows[i] || [];
    const id = String(row[SERVICOS_COL.ID] ?? '').trim();
    const title = String(row[SERVICOS_COL.TITULO] ?? '').trim();
    const rawDuration = row[SERVICOS_COL.DURACAO];
    if (!id && !title) continue;

    const entry = {
      rowIndex: i + 1, // 1-based sheet row
      id,
      title,
      normalizedTitle: normalizeText(title),
      rawDuration,
      durationMinutes: parseOfficialDurationMinutes(rawDuration)
    };
    entries.push(entry);

    if (id) {
      const list = byId.get(id) || [];
      list.push(entry);
      byId.set(id, list);
    }
    if (entry.normalizedTitle) {
      const list = byTitle.get(entry.normalizedTitle) || [];
      list.push(entry);
      byTitle.set(entry.normalizedTitle, list);
    }
  }

  return { entries, byId, byTitle };
}

/**
 * Localiza o serviço oficial na aba "Serviços" e retorna sua duração oficial da Coluna F.
 *
 * Regras:
 * 1. Preferencialmente pelo ID (Col P de Agendamentos -> Col A de Serviços).
 * 2. Se o Procedimento (Col E) foi alterado e não corresponde mais ao título do ID antigo,
 *    busca o novo serviço pelo título na Col B para respeitar alteração de procedimento.
 * 3. Caso contrário (sem ID), busca pelo título normalizado na Col B.
 * 4. Em serviço inexistente, correspondência ambígua ou duração inválida: retorna erro,
 *    NUNCA assume duração padrão.
 */
export function resolveServiceDurationFromSheet(servicosRows, { serviceId = '', procedureTitle = '', additionalServiceIds = [] }) {
  const catalog = indexServicesFromSheet(servicosRows);
  const cleanId = String(serviceId ?? '').trim();
  const cleanTitle = String(procedureTitle ?? '').trim();
  const normTitle = normalizeText(cleanTitle);

  let matchedService = null;
  let resolvedBy = null;

  if (cleanId) {
    const idMatches = catalog.byId.get(cleanId) || [];
    if (idMatches.length > 1) {
      return {
        ok: false,
        errorCode: 'AMBIGUOUS_SERVICE_MATCH',
        message: `Correspondência ambígua na aba Serviços para ID "${cleanId}" (${idMatches.length} registros).`
      };
    }

    if (idMatches.length === 1) {
      const candidate = idMatches[0];
      // Se Coluna E (Procedimento) foi alterada para outro serviço diferente do ID salvo em Col P:
      if (normTitle && candidate.normalizedTitle && candidate.normalizedTitle !== normTitle) {
        const titleMatches = catalog.byTitle.get(normTitle) || [];
        if (titleMatches.length === 1) {
          matchedService = titleMatches[0];
          resolvedBy = 'TITLE_OVERRIDE_AFTER_PROCEDURE_EDIT';
        } else if (titleMatches.length > 1) {
          return {
            ok: false,
            errorCode: 'AMBIGUOUS_SERVICE_MATCH',
            message: `Correspondência ambígua na aba Serviços para procedimento "${cleanTitle}".`
          };
        } else {
          return {
            ok: false,
            errorCode: 'SERVICE_NOT_FOUND',
            message: `Procedimento "${cleanTitle}" não encontrado na aba Serviços (diverge do ID anterior "${cleanId}").`
          };
        }
      } else {
        matchedService = candidate;
        resolvedBy = 'ID';
      }
    } else if (normTitle) {
      const titleMatches = catalog.byTitle.get(normTitle) || [];
      if (titleMatches.length === 1) {
        matchedService = titleMatches[0];
        resolvedBy = 'TITLE_FALLBACK';
      } else if (titleMatches.length > 1) {
        return {
          ok: false,
          errorCode: 'AMBIGUOUS_SERVICE_MATCH',
          message: `Correspondência ambígua na aba Serviços para procedimento "${cleanTitle}".`
        };
      }
    }
  } else if (normTitle) {
    const titleMatches = catalog.byTitle.get(normTitle) || [];
    if (titleMatches.length > 1) {
      return {
        ok: false,
        errorCode: 'AMBIGUOUS_SERVICE_MATCH',
        message: `Correspondência ambígua na aba Serviços para título "${cleanTitle}" (${titleMatches.length} registros).`
      };
    }
    if (titleMatches.length === 1) {
      matchedService = titleMatches[0];
      resolvedBy = 'TITLE';
    }
  }

  if (!matchedService) {
    return {
      ok: false,
      errorCode: 'SERVICE_NOT_FOUND',
      message: `Serviço não encontrado na aba Serviços (ID: "${cleanId}", Procedimento: "${cleanTitle}"). Nenhuma duração padrão foi assumida.`
    };
  }

  if (matchedService.durationMinutes === null) {
    return {
      ok: false,
      errorCode: 'INVALID_DURATION',
      service: matchedService,
      message: `Duração inválida na aba Serviços, coluna F (linha ${matchedService.rowIndex}, serviço "${matchedService.title}", valor: "${matchedService.rawDuration ?? ''}").`
    };
  }

  let totalMinutes = matchedService.durationMinutes;
  const extraServices = [];

  for (const rawAddId of additionalServiceIds || []) {
    const addId = String(rawAddId ?? '').trim();
    if (!addId) continue;
    const addMatches = catalog.byId.get(addId) || [];
    if (addMatches.length === 0) {
      return {
        ok: false,
        errorCode: 'ADDITIONAL_SERVICE_NOT_FOUND',
        message: `Serviço adicional "${addId}" não encontrado na aba Serviços.`
      };
    }
    if (addMatches.length > 1) {
      return {
        ok: false,
        errorCode: 'AMBIGUOUS_SERVICE_MATCH',
        message: `Correspondência ambígua para serviço adicional "${addId}".`
      };
    }
    if (addMatches[0].durationMinutes === null) {
      return {
        ok: false,
        errorCode: 'INVALID_DURATION',
        message: `Duração inválida para serviço adicional "${addMatches[0].title}" na coluna F da aba Serviços.`
      };
    }
    totalMinutes += addMatches[0].durationMinutes;
    extraServices.push(addMatches[0]);
  }

  return {
    ok: true,
    service: matchedService,
    additionalServices: extraServices,
    durationMinutes: totalMinutes,
    resolvedBy
  };
}

function parseAdditionalIds(rawValue) {
  if (Array.isArray(rawValue)) return rawValue.map(String).filter(Boolean);
  const s = String(rawValue ?? '').trim();
  if (!s) return [];
  try {
    const parsed = JSON.parse(s);
    return Array.isArray(parsed) ? parsed.map(String).filter(Boolean) : [];
  } catch {
    return [];
  }
}

/**
 * Verifica se a linha de Agendamentos é cancelada, bloqueio de agenda ou exceção autorizada.
 */
export function classifyRowGuards(row, authorizedExceptions = {}) {
  const client = String(row[AGENDAMENTOS_COL.CLIENTE] ?? '').trim();
  const procedure = String(row[AGENDAMENTOS_COL.PROCEDIMENTO] ?? '').trim();
  const status = String(row[AGENDAMENTOS_COL.STATUS] ?? '').trim();
  const appointmentId = String(row[AGENDAMENTOS_COL.APPOINTMENT_ID] ?? '').trim();

  const normStatus = normalizeText(status);
  const normClient = normalizeText(client);
  const normProcedure = normalizeText(procedure);

  if (!client && !appointmentId && !procedure) {
    return { skip: true, reason: 'EMPTY_ROW', message: 'Linha vazia ignorada.' };
  }

  // 1. Cancelados não devem ser alterados
  if (CANCELLED_STATUSES.has(normStatus)) {
    return { skip: true, reason: 'SKIPPED_CANCELLED', message: `Agendamento "${appointmentId || client}" cancelado — não alterado.` };
  }

  // 2. Bloqueios de agenda / folgas não devem ser alterados
  const isBlock = BLOCK_MARKERS.some(
    (marker) => normStatus === marker || normStatus.includes(marker) || normProcedure.startsWith(marker) || normClient.startsWith(marker)
  );
  if (isBlock) {
    return { skip: true, reason: 'SKIPPED_SCHEDULE_BLOCK', message: `Bloqueio de agenda "${appointmentId || procedure || client}" preservado.` };
  }

  // 3. Exceções de sincronização já autorizadas no sistema (ex: EXCLUDE_FROM_SYNC / teste_sistema)
  const excludeIds = new Set((authorizedExceptions.excludeAppointmentIds || []).map(String));
  if (
    (appointmentId && excludeIds.has(appointmentId)) ||
    normStatus === 'teste sistema' ||
    appointmentId.startsWith('test-')
  ) {
    return {
      skip: true,
      reason: 'SKIPPED_AUTHORIZED_EXCLUSION',
      message: `Exceção autorizada de exclusão preservada para "${appointmentId || client}".`
    };
  }

  // 4. Exceção autorizada de duração/calendário fixo (ex: KEEP_CALENDAR_AND_ALIGN_SHEET)
  const durationOverrides = authorizedExceptions.durationOverrideByAppointmentId || {};
  if (appointmentId && Object.prototype.hasOwnProperty.call(durationOverrides, appointmentId)) {
    return {
      skip: true,
      reason: 'SKIPPED_AUTHORIZED_DURATION_EXCEPTION',
      message: `Exceção autorizada de duração preservada para "${appointmentId}".`
    };
  }

  return { skip: false };
}

export function buildEventTitle(client, procedure) {
  if (client && procedure) return `${client} — ${procedure}`;
  return procedure || client || 'Agendamento Espaço Angélica';
}

export function buildIdempotencyDescription(appointmentId, procedure, existingDescription = '') {
  const tag = appointmentId ? `${APT_TAG_PREFIX}${appointmentId}]` : '';
  const base = String(existingDescription || '').trim();
  if (tag && base.includes(tag)) return base;
  const summary = `Procedimento: ${procedure}`;
  return [base, summary, tag].filter(Boolean).join('\n');
}

/**
 * Sincroniza uma única linha da aba "Agendamentos" com o Google Calendar.
 * - Sempre lê duração fresca da aba "Serviços" Coluna F.
 * - Nunca usa Coluna F de "Agendamentos" como duração.
 * - Usa o ID do evento na Coluna T (e tag de idempotência como proteção secundária) para evitar duplicação.
 */
export function syncAppointmentRow({
  rowIndex,
  row,
  servicosRows,
  calendarAdapter,
  now = new Date(),
  authorizedExceptions = {},
  requireFuture = false,
  preserveStartOnUpdate = false,
  onWriteCalendarEventId = null,
  onWriteServiceId = null
}) {
  const appointmentId = String(row[AGENDAMENTOS_COL.APPOINTMENT_ID] ?? '').trim();
  const client = String(row[AGENDAMENTOS_COL.CLIENTE] ?? '').trim();
  const procedure = String(row[AGENDAMENTOS_COL.PROCEDIMENTO] ?? '').trim();
  const dateRaw = row[AGENDAMENTOS_COL.DATA_AGENDADA];
  const timeRaw = row[AGENDAMENTOS_COL.HORARIO_AGENDADO];
  const serviceId = String(row[AGENDAMENTOS_COL.SERVICE_ID] ?? '').trim();
  const additionalServiceIds = parseAdditionalIds(row[AGENDAMENTOS_COL.ADDITIONAL_IDS]);
  const storedEventId = String(row[AGENDAMENTOS_COL.CALENDAR_EVENT_ID] ?? '').trim();

  // 1. Regras de proteção: cancelados, bloqueios, exceções autorizadas
  const guard = classifyRowGuards(row, authorizedExceptions);
  if (guard.skip) {
    return {
      status: guard.reason,
      rowIndex,
      appointmentId,
      eventId: storedEventId || null,
      changed: false,
      message: guard.message
    };
  }

  // 2. Validar data (Col F) e horário inicial (Col G) antes de checar passado
  const dateKey = parseSheetDateKey(dateRaw);
  const startTimeKey = parseSheetTimeKey(timeRaw);
  if (!dateKey || !startTimeKey) {
    return {
      status: 'ERROR',
      errorCode: !dateKey ? 'INVALID_DATE' : 'INVALID_START_TIME',
      rowIndex,
      appointmentId,
      eventId: storedEventId || null,
      changed: false,
      message: `Data (${dateRaw}) ou horário inicial (${timeRaw}) inválido na linha ${rowIndex} de Agendamentos. Evento não alterado.`
    };
  }

  const startProbe = new Date(`${dateKey}T${startTimeKey}:00${TIMEZONE_OFFSET}`);
  const referenceNow = now instanceof Date ? now : new Date(now);
  if (requireFuture && startProbe.getTime() <= referenceNow.getTime()) {
    return {
      status: 'SKIPPED_PAST',
      rowIndex,
      appointmentId,
      eventId: storedEventId || null,
      changed: false,
      message: `Agendamento passado (${dateKey} ${startTimeKey}) — não alterado.`
    };
  }

  // 3. Resolver duração oficial exclusivamente a partir da aba Serviços, Coluna F
  const durationLookup = resolveServiceDurationFromSheet(servicosRows, {
    serviceId,
    procedureTitle: procedure,
    additionalServiceIds
  });

  if (!durationLookup.ok) {
    return {
      status: 'ERROR',
      errorCode: durationLookup.errorCode,
      rowIndex,
      appointmentId,
      eventId: storedEventId || null,
      changed: false,
      message: durationLookup.message
    };
  }

  // 4. Calcular intervalo: HORÁRIO FINAL = HORÁRIO INICIAL + DURAÇÃO DO SERVIÇO
  const interval = calculateEventInterval(dateKey, startTimeKey, durationLookup.durationMinutes);
  if (!interval.ok) {
    return {
      status: 'ERROR',
      errorCode: interval.errorCode,
      rowIndex,
      appointmentId,
      eventId: storedEventId || null,
      changed: false,
      message: interval.message
    };
  }

  // 5. Localizar evento existente pelo ID do Google Calendar (Col T) ou tag de idempotência
  let existingEvent = storedEventId ? calendarAdapter.getEventById(storedEventId) : null;
  if (!existingEvent && appointmentId && typeof calendarAdapter.findEventByAppointmentId === 'function') {
    existingEvent = calendarAdapter.findEventByAppointmentId(appointmentId);
  }

  // Nunca alterar eventos de dia inteiro (ex: Folga)
  if (existingEvent && existingEvent.allDay) {
    return {
      status: 'SKIPPED_ALL_DAY_EVENT',
      rowIndex,
      appointmentId,
      eventId: existingEvent.id,
      changed: false,
      message: `Evento de dia inteiro ("${existingEvent.title}") preservado.`
    };
  }

  const desiredTitle = buildEventTitle(client, durationLookup.service.title || procedure);

  if (existingEvent) {
    const targetStart = preserveStartOnUpdate ? new Date(existingEvent.startIso) : interval.startDate;
    const targetStartIso = preserveStartOnUpdate ? existingEvent.startIso : interval.startIso;
    const targetEndDate = preserveStartOnUpdate
      ? new Date(targetStart.getTime() + durationLookup.durationMinutes * 60_000)
      : interval.endDate;
    const targetInterval = preserveStartOnUpdate
      ? calculateEventInterval(
          existingEvent.startIso.slice(0, 10),
          existingEvent.startIso.slice(11, 16),
          durationLookup.durationMinutes
        )
      : interval;

    const alreadySynced =
      existingEvent.startIso === targetStartIso &&
      existingEvent.endIso === targetInterval.endIso &&
      existingEvent.title === desiredTitle;

    // Garante vínculo salvo em Col T caso recuperado pela tag de idempotência
    if (!storedEventId && typeof onWriteCalendarEventId === 'function') {
      onWriteCalendarEventId(rowIndex, existingEvent.id);
    }
    if (
      durationLookup.resolvedBy === 'TITLE_OVERRIDE_AFTER_PROCEDURE_EDIT' &&
      durationLookup.service.id &&
      typeof onWriteServiceId === 'function'
    ) {
      onWriteServiceId(rowIndex, durationLookup.service.id);
    }

    if (alreadySynced) {
      return {
        status: 'UNCHANGED',
        rowIndex,
        appointmentId,
        eventId: existingEvent.id,
        durationMinutes: durationLookup.durationMinutes,
        startIso: existingEvent.startIso,
        endIso: existingEvent.endIso,
        changed: false,
        message: `Evento ${existingEvent.id} já está sincronizado (${existingEvent.startIso} → ${existingEvent.endIso}).`
      };
    }

    const before = {
      title: existingEvent.title,
      startIso: existingEvent.startIso,
      endIso: existingEvent.endIso
    };

    const updated = calendarAdapter.updateEvent(existingEvent.id, {
      title: desiredTitle,
      startDate: targetStart,
      endDate: targetEndDate,
      startIso: targetStartIso,
      endIso: targetInterval.endIso,
      description: buildIdempotencyDescription(appointmentId, durationLookup.service.title || procedure, existingEvent.description)
    });

    return {
      status: 'UPDATED',
      rowIndex,
      appointmentId,
      eventId: updated.id,
      serviceId: durationLookup.service.id,
      serviceTitle: durationLookup.service.title,
      durationMinutes: durationLookup.durationMinutes,
      before,
      after: {
        title: updated.title,
        startIso: updated.startIso,
        endIso: updated.endIso
      },
      changed: true,
      message: `Evento ${updated.id} atualizado: ${before.startIso.slice(11, 16)}–${before.endIso.slice(11, 16)} → ${updated.startIso.slice(11, 16)}–${updated.endIso.slice(11, 16)} (${durationLookup.durationMinutes} min).`
    };
  }

  // 6. Criar novo evento (quando ainda não existe) e gravar Event ID na Coluna T
  const created = calendarAdapter.createEvent({
    title: desiredTitle,
    startDate: interval.startDate,
    endDate: interval.endDate,
    startIso: interval.startIso,
    endIso: interval.endIso,
    appointmentId,
    description: buildIdempotencyDescription(appointmentId, durationLookup.service.title || procedure)
  });

  if (typeof onWriteCalendarEventId === 'function') {
    onWriteCalendarEventId(rowIndex, created.id);
  }
  if (
    durationLookup.resolvedBy === 'TITLE_OVERRIDE_AFTER_PROCEDURE_EDIT' &&
    durationLookup.service.id &&
    typeof onWriteServiceId === 'function'
  ) {
    onWriteServiceId(rowIndex, durationLookup.service.id);
  }

  return {
    status: 'CREATED',
    rowIndex,
    appointmentId,
    eventId: created.id,
    serviceId: durationLookup.service.id,
    serviceTitle: durationLookup.service.title,
    durationMinutes: durationLookup.durationMinutes,
    after: {
      title: created.title,
      startIso: created.startIso,
      endIso: created.endIso
    },
    changed: true,
    message: `Evento ${created.id} criado para ${dateKey} ${interval.startTimeKey}–${interval.endTimeKey} (${durationLookup.durationMinutes} min).`
  };
}

/**
 * Identifica se uma linha de Agendamentos está vinculada ao serviço editado na aba Serviços.
 */
function appointmentMatchesService(row, targetService) {
  const rowServiceId = String(row[AGENDAMENTOS_COL.SERVICE_ID] ?? '').trim();
  const rowProcedure = normalizeText(row[AGENDAMENTOS_COL.PROCEDIMENTO]);
  const additionalIds = parseAdditionalIds(row[AGENDAMENTOS_COL.ADDITIONAL_IDS]);

  if (targetService.id && rowServiceId && rowServiceId === targetService.id) {
    // Se o procedimento em Col E foi editado para outro título, não vincular pelo ID obsoleto
    if (!rowProcedure || rowProcedure === targetService.normalizedTitle) {
      return true;
    }
  }
  if (targetService.normalizedTitle && rowProcedure === targetService.normalizedTitle) {
    return true;
  }
  if (targetService.id && additionalIds.includes(targetService.id)) {
    return true;
  }
  return false;
}

/**
 * Executado quando a Coluna F (Duração) da aba "Serviços" é alterada.
 * - Valida a nova duração imediatamente (se inválida, registra erro e não altera nenhum evento).
 * - Localiza todos os agendamentos futuros vinculados àquele serviço.
 * - Recalcula o horário final preservando data e horário inicial.
 * - Atualiza eventos existentes via Event ID sem duplicar.
 * - Ignora passados, cancelados, bloqueios e exceções autorizadas.
 */
export function syncFutureAppointmentsForServiceDurationChange({
  editedServiceRowIndex,
  serviceId = '',
  serviceTitle = '',
  servicosRows,
  agendamentosRows,
  calendarAdapter,
  now = new Date(),
  authorizedExceptions = {},
  onWriteCalendarEventId = null
}) {
  const logs = [];
  const catalog = indexServicesFromSheet(servicosRows);

  let targetService = null;
  if (editedServiceRowIndex && editedServiceRowIndex >= 2) {
    targetService = catalog.entries.find((e) => e.rowIndex === editedServiceRowIndex) || null;
  }
  if (!targetService && serviceId) {
    const matches = catalog.byId.get(String(serviceId).trim()) || [];
    if (matches.length === 1) targetService = matches[0];
  }
  if (!targetService && serviceTitle) {
    const matches = catalog.byTitle.get(normalizeText(serviceTitle)) || [];
    if (matches.length === 1) targetService = matches[0];
  }

  if (!targetService) {
    const err = {
      status: 'ERROR',
      errorCode: 'SERVICE_NOT_FOUND',
      message: `Serviço editado não localizado na aba Serviços (row=${editedServiceRowIndex}, id="${serviceId}", título="${serviceTitle}").`
    };
    logs.push(err);
    return { ok: false, updated: 0, unchanged: 0, skipped: 0, errors: 1, logs };
  }

  if (targetService.durationMinutes === null) {
    const err = {
      status: 'ERROR',
      errorCode: 'INVALID_DURATION',
      serviceId: targetService.id,
      serviceTitle: targetService.title,
      message: `Nova duração inválida na aba Serviços, coluna F (linha ${targetService.rowIndex}, serviço "${targetService.title}", valor: "${targetService.rawDuration ?? ''}"). Nenhum evento foi alterado.`
    };
    logs.push(err);
    return { ok: false, updated: 0, unchanged: 0, skipped: 0, errors: 1, logs };
  }

  let updated = 0;
  let created = 0;
  let unchanged = 0;
  let skipped = 0;
  let errors = 0;

  for (let i = 1; i < agendamentosRows.length; i++) {
    const row = agendamentosRows[i] || [];
    if (!appointmentMatchesService(row, targetService)) continue;

    const result = syncAppointmentRow({
      rowIndex: i + 1,
      row,
      servicosRows,
      calendarAdapter,
      now,
      authorizedExceptions,
      requireFuture: true,
      preserveStartOnUpdate: true,
      onWriteCalendarEventId
    });
    logs.push(result);

    if (result.status === 'UPDATED') updated++;
    else if (result.status === 'CREATED') created++;
    else if (result.status === 'UNCHANGED') unchanged++;
    else if (result.status === 'ERROR') errors++;
    else skipped++;
  }

  return {
    ok: errors === 0,
    service: {
      id: targetService.id,
      title: targetService.title,
      durationMinutes: targetService.durationMinutes
    },
    updated,
    created,
    unchanged,
    skipped,
    errors,
    logs
  };
}

/**
 * Despachante de detecção de alterações nas abas "Serviços" e "Agendamentos".
 * - Edição na aba "Serviços", Coluna F (col 6, linha >= 2) -> propaga nova duração para agendamentos futuros.
 * - Edição na aba "Agendamentos", Coluna E (Procedimento, col 5), Coluna F (Data, col 6),
 *   Coluna G (Horário, col 7), Coluna N (Adicionais, col 14) ou Coluna P (ID Serviço, col 16)
 *   -> recalcula horário final usando duração atual de "Serviços" Col F e atualiza o evento.
 */
export function handleSpreadsheetEdit({
  sheetName,
  editedRow,
  editedColumn,
  servicosRows,
  agendamentosRows,
  calendarAdapter,
  now = new Date(),
  authorizedExceptions = {},
  onWriteCalendarEventId = null,
  onWriteServiceId = null
}) {
  const normSheet = normalizeText(sheetName);
  if (!editedRow || editedRow < 2) {
    return { ok: true, action: 'IGNORED_HEADER_EDIT', logs: [] };
  }

  // Aba Serviços -> Coluna F (1-based: 6)
  if (normSheet === 'servicos' || normSheet === 'depilacao') {
    if (editedColumn !== SERVICOS_COL.DURACAO + 1) {
      return { ok: true, action: 'IGNORED_NON_DURATION_COLUMN', logs: [] };
    }
    const summary = syncFutureAppointmentsForServiceDurationChange({
      editedServiceRowIndex: editedRow,
      servicosRows,
      agendamentosRows,
      calendarAdapter,
      now,
      authorizedExceptions,
      onWriteCalendarEventId
    });
    return { ...summary, action: 'SYNC_SERVICE_DURATION_CHANGE' };
  }

  // Aba Agendamentos -> Coluna E (5), Coluna F (6), Coluna G (7), Coluna N (14), Coluna P (16)
  if (normSheet === 'agendamentos') {
    const watchedCols = new Set([
      AGENDAMENTOS_COL.PROCEDIMENTO + 1,
      AGENDAMENTOS_COL.DATA_AGENDADA + 1,
      AGENDAMENTOS_COL.HORARIO_AGENDADO + 1,
      AGENDAMENTOS_COL.ADDITIONAL_IDS + 1,
      AGENDAMENTOS_COL.SERVICE_ID + 1
    ]);
    if (!watchedCols.has(editedColumn)) {
      return { ok: true, action: 'IGNORED_NON_SYNC_COLUMN', logs: [] };
    }

    const row = agendamentosRows[editedRow - 1];
    if (!row) {
      return {
        ok: false,
        action: 'SYNC_APPOINTMENT_EDIT',
        logs: [{ status: 'ERROR', errorCode: 'ROW_NOT_FOUND', rowIndex: editedRow, message: `Linha ${editedRow} não encontrada em Agendamentos.` }]
      };
    }

    const result = syncAppointmentRow({
      rowIndex: editedRow,
      row,
      servicosRows,
      calendarAdapter,
      now,
      authorizedExceptions,
      requireFuture: true,
      preserveStartOnUpdate: false,
      onWriteCalendarEventId,
      onWriteServiceId
    });

    return {
      ok: result.status !== 'ERROR',
      action: 'SYNC_APPOINTMENT_EDIT',
      result,
      logs: [result]
    };
  }

  return { ok: true, action: 'IGNORED_OTHER_SHEET', logs: [] };
}

/**
 * Adaptador de Google Calendar em memória para ambiente de preview e testes automatizados.
 * Garante que nenhuma chamada real ao Google Calendar seja feita.
 */
export function createInMemoryCalendarAdapter(initialEvents = []) {
  const eventsById = new Map();
  let seq = 1;
  let createCalls = 0;
  let updateCalls = 0;

  for (const ev of initialEvents) {
    eventsById.set(ev.id, { ...ev });
  }

  return {
    getEventById(id) {
      const found = eventsById.get(id);
      return found ? { ...found } : null;
    },
    findEventByAppointmentId(appointmentId) {
      const marker = `${APT_TAG_PREFIX}${appointmentId}]`;
      for (const ev of eventsById.values()) {
        if (ev.appointmentId === appointmentId || String(ev.description || '').includes(marker)) {
          return { ...ev };
        }
      }
      return null;
    },
    createEvent(payload) {
      createCalls++;
      const id = `mock-cal-evt-${seq++}`;
      const record = {
        id,
        title: payload.title,
        startIso: payload.startIso,
        endIso: payload.endIso,
        description: payload.description || '',
        appointmentId: payload.appointmentId || '',
        allDay: false
      };
      eventsById.set(id, record);
      return { ...record };
    },
    updateEvent(id, payload) {
      const existing = eventsById.get(id);
      if (!existing) throw new Error(`Evento não encontrado no adaptador em memória: ${id}`);
      updateCalls++;
      const updated = {
        ...existing,
        title: payload.title ?? existing.title,
        startIso: payload.startIso ?? existing.startIso,
        endIso: payload.endIso ?? existing.endIso,
        description: payload.description ?? existing.description
      };
      eventsById.set(id, updated);
      return { ...updated };
    },
    listEvents() {
      return Array.from(eventsById.values()).map((e) => ({ ...e }));
    },
    getStats() {
      return {
        totalEvents: eventsById.size,
        createCalls,
        updateCalls
      };
    }
  };
}
