import test from 'node:test';
import assert from 'node:assert/strict';
import {
  AGENDAMENTOS_COL,
  SERVICOS_COL,
  createInMemoryCalendarAdapter,
  handleSpreadsheetEdit,
  parseOfficialDurationMinutes,
  syncAppointmentRow,
  syncFutureAppointmentsForServiceDurationChange
} from '../../src/services/calendarDurationSyncEngine.js';

/**
 * Suíte oficial dos 11 testes obrigatórios — Sincronização da Duração com Google Calendar
 *
 * TODOS os dados abaixo são 100% fictícios (ambiente de preview).
 * Nenhuma alteração real na planilha oficial ou no Google Calendar é executada.
 */

const REFERENCE_NOW = new Date('2026-11-21T10:00:00-03:00');

function createMockServicosSheet() {
  return [
    // Col A(0): ID | Col B(1): Título | Col C(2): Categoria | Col D(3): Preço | Col E(4): Preço Formatado | Col F(5): Duração | Col G(6): Badge | Col H(7): Descrição
    ['ID', 'Título', 'Categoria', 'Preço', 'Preço Formatado', 'Duração', 'Badge', 'Descrição'],
    ['manicure-tradicional-mao', 'Manicure Tradicional (Mão)', 'Manicure', 40, 'R$ 40,00', '45 min', '', 'Cutilagem e esmaltação tradicional'],
    ['pedicure-spa', 'Pedicure Spa', 'Pedicure', 65, 'R$ 65,00', '60 min', '', 'Cuidados completos para os pés'],
    ['alongamento-fibra', 'Alongamento em Fibra de Vidro', 'Alongamento', 180, 'R$ 180,00', '2h', 'Destaque', 'Alongamento slim em fibra de vidro'],
    ['servico-duracao-invalida', 'Serviço Duração Inválida', 'Manicure', 50, 'R$ 50,00', 'invalida', '', 'Registro para teste de validação']
  ];
}

function makeAgendamentoRow({
  client = 'Cliente Fictícia',
  procedure = 'Manicure Tradicional (Mão)',
  date = '25/11/2026',
  time = '14:00',
  deposit = 20,
  status = 'confirmado',
  appointmentId = 'apt-mock-001',
  additionalIds = '',
  serviceId = 'manicure-tradicional-mao',
  calendarEventId = ''
} = {}) {
  const row = new Array(20).fill('');
  row[AGENDAMENTOS_COL.TIMESTAMP] = '2026-11-20T12:00:00-03:00';
  row[AGENDAMENTOS_COL.CLIENTE] = client;
  row[AGENDAMENTOS_COL.PROCEDIMENTO] = procedure;      // Coluna E
  row[AGENDAMENTOS_COL.DATA_AGENDADA] = date;          // Coluna F (Data — NÃO duração)
  row[AGENDAMENTOS_COL.HORARIO_AGENDADO] = time;       // Coluna G
  row[AGENDAMENTOS_COL.SINAL] = deposit;               // Coluna H
  row[AGENDAMENTOS_COL.STATUS] = status;               // Coluna I
  row[AGENDAMENTOS_COL.APPOINTMENT_ID] = appointmentId;// Coluna J
  row[AGENDAMENTOS_COL.ADDITIONAL_IDS] = additionalIds;// Coluna N
  row[AGENDAMENTOS_COL.SERVICE_ID] = serviceId;        // Coluna P
  row[AGENDAMENTOS_COL.CALENDAR_EVENT_ID] = calendarEventId; // Coluna T
  return row;
}

function createAgendamentosHeader() {
  const header = new Array(20).fill('');
  header[0] = 'Timestamp';
  header[1] = 'Cliente';
  header[4] = 'Procedimento';
  header[5] = 'Data Agendada';
  header[6] = 'Horário Agendado';
  header[7] = 'Sinal (R$)';
  header[8] = 'Status';
  header[9] = 'ID Agendamento';
  header[13] = 'IDs Serviços Adicionais';
  header[15] = 'ID Serviço Principal';
  header[19] = 'Google Calendar Event ID';
  return header;
}

// ============================================================================
// TESTE 1: Alteração da duração de 45 para 60 minutos
// ============================================================================
test('1. Alteração da duração de 45 para 60 minutos atualiza o evento existente (14:00–14:45 -> 14:00–15:00)', () => {
  const servicos = createMockServicosSheet();
  const agendamentos = [
    createAgendamentosHeader(),
    makeAgendamentoRow({
      client: 'Clara Fictícia',
      procedure: 'Manicure Tradicional (Mão)',
      date: '25/11/2026',
      time: '14:00',
      appointmentId: 'apt-mock-001',
      serviceId: 'manicure-tradicional-mao',
      calendarEventId: 'evt-manicure-001'
    })
  ];

  const calendar = createInMemoryCalendarAdapter([
    {
      id: 'evt-manicure-001',
      title: 'Clara Fictícia — Manicure Tradicional (Mão)',
      startIso: '2026-11-25T14:00:00-03:00',
      endIso: '2026-11-25T14:45:00-03:00',
      description: '[ESPACO_ANGELICA_APT_ID:apt-mock-001]',
      allDay: false
    }
  ]);

  // Altera duração na aba Serviços, Coluna F (de 45 min para 60 min)
  servicos[1][SERVICOS_COL.DURACAO] = '60 min';

  const res = handleSpreadsheetEdit({
    sheetName: 'Serviços',
    editedRow: 2,
    editedColumn: SERVICOS_COL.DURACAO + 1,
    servicosRows: servicos,
    agendamentosRows: agendamentos,
    calendarAdapter: calendar,
    now: REFERENCE_NOW
  });

  assert.equal(res.ok, true);
  assert.equal(res.updated, 1);
  assert.equal(calendar.getStats().totalEvents, 1);
  assert.equal(calendar.getStats().createCalls, 0);

  const evt = calendar.getEventById('evt-manicure-001');
  assert.equal(evt.startIso, '2026-11-25T14:00:00-03:00');
  assert.equal(evt.endIso, '2026-11-25T15:00:00-03:00');
});

// ============================================================================
// TESTE 2: Alteração da duração de 60 para 90 minutos
// ============================================================================
test('2. Alteração da duração de 60 para 90 minutos atualiza o mesmo evento (14:00–15:00 -> 14:00–15:30)', () => {
  const servicos = createMockServicosSheet();
  servicos[1][SERVICOS_COL.DURACAO] = '60 min';

  const agendamentos = [
    createAgendamentosHeader(),
    makeAgendamentoRow({
      client: 'Clara Fictícia',
      procedure: 'Manicure Tradicional (Mão)',
      date: '25/11/2026',
      time: '14:00',
      appointmentId: 'apt-mock-002',
      serviceId: 'manicure-tradicional-mao',
      calendarEventId: 'evt-manicure-002'
    })
  ];

  const calendar = createInMemoryCalendarAdapter([
    {
      id: 'evt-manicure-002',
      title: 'Clara Fictícia — Manicure Tradicional (Mão)',
      startIso: '2026-11-25T14:00:00-03:00',
      endIso: '2026-11-25T15:00:00-03:00',
      description: '[ESPACO_ANGELICA_APT_ID:apt-mock-002]',
      allDay: false
    }
  ]);

  // Altera duração na aba Serviços, Coluna F (de 60 min para 90 min / 1h 30min)
  servicos[1][SERVICOS_COL.DURACAO] = '1h 30min';

  const res = handleSpreadsheetEdit({
    sheetName: 'Serviços',
    editedRow: 2,
    editedColumn: SERVICOS_COL.DURACAO + 1,
    servicosRows: servicos,
    agendamentosRows: agendamentos,
    calendarAdapter: calendar,
    now: REFERENCE_NOW
  });

  assert.equal(res.ok, true);
  assert.equal(res.updated, 1);
  assert.equal(calendar.getStats().totalEvents, 1);
  assert.equal(calendar.getStats().createCalls, 0);

  const evt = calendar.getEventById('evt-manicure-002');
  assert.equal(evt.startIso, '2026-11-25T14:00:00-03:00');
  assert.equal(evt.endIso, '2026-11-25T15:30:00-03:00');
});

// ============================================================================
// TESTE 3: Alteração do horário inicial na aba Agendamentos
// ============================================================================
test('3. Alteração do horário inicial em Agendamentos recalcula o horário final pela duração oficial de Serviços', () => {
  const servicos = createMockServicosSheet(); // Manicure Tradicional (Mão) = 45 min
  const agendamentos = [
    createAgendamentosHeader(),
    makeAgendamentoRow({
      client: 'Beatriz Fictícia',
      procedure: 'Manicure Tradicional (Mão)',
      date: '25/11/2026',
      time: '16:30', // alterado de 14:00 para 16:30
      appointmentId: 'apt-mock-003',
      serviceId: 'manicure-tradicional-mao',
      calendarEventId: 'evt-time-003'
    })
  ];

  const calendar = createInMemoryCalendarAdapter([
    {
      id: 'evt-time-003',
      title: 'Beatriz Fictícia — Manicure Tradicional (Mão)',
      startIso: '2026-11-25T14:00:00-03:00',
      endIso: '2026-11-25T14:45:00-03:00',
      description: '[ESPACO_ANGELICA_APT_ID:apt-mock-003]',
      allDay: false
    }
  ]);

  const res = handleSpreadsheetEdit({
    sheetName: 'Agendamentos',
    editedRow: 2,
    editedColumn: AGENDAMENTOS_COL.HORARIO_AGENDADO + 1, // Coluna G
    servicosRows: servicos,
    agendamentosRows: agendamentos,
    calendarAdapter: calendar,
    now: REFERENCE_NOW
  });

  assert.equal(res.ok, true);
  assert.equal(res.result.status, 'UPDATED');
  assert.equal(calendar.getStats().totalEvents, 1);

  const evt = calendar.getEventById('evt-time-003');
  assert.equal(evt.startIso, '2026-11-25T16:30:00-03:00');
  assert.equal(evt.endIso, '2026-11-25T17:15:00-03:00');
});

// ============================================================================
// TESTE 4: Alteração da data na aba Agendamentos
// ============================================================================
test('4. Alteração da data (Coluna F de Agendamentos) preserva horário inicial e recalcula fim pela Coluna F de Serviços', () => {
  const servicos = createMockServicosSheet(); // Pedicure Spa = 60 min
  const agendamentos = [
    createAgendamentosHeader(),
    makeAgendamentoRow({
      client: 'Helena Fictícia',
      procedure: 'Pedicure Spa',
      date: '28/11/2026', // alterado de 25/11/2026 para 28/11/2026
      time: '10:00',
      appointmentId: 'apt-mock-004',
      serviceId: 'pedicure-spa',
      calendarEventId: 'evt-date-004'
    })
  ];

  const calendar = createInMemoryCalendarAdapter([
    {
      id: 'evt-date-004',
      title: 'Helena Fictícia — Pedicure Spa',
      startIso: '2026-11-25T10:00:00-03:00',
      endIso: '2026-11-25T11:00:00-03:00',
      description: '[ESPACO_ANGELICA_APT_ID:apt-mock-004]',
      allDay: false
    }
  ]);

  // Verifica que a Coluna F de Agendamentos (data) NUNCA é interpretada como duração
  assert.equal(parseOfficialDurationMinutes(agendamentos[1][AGENDAMENTOS_COL.DATA_AGENDADA]), null);

  const res = handleSpreadsheetEdit({
    sheetName: 'Agendamentos',
    editedRow: 2,
    editedColumn: AGENDAMENTOS_COL.DATA_AGENDADA + 1, // Coluna F
    servicosRows: servicos,
    agendamentosRows: agendamentos,
    calendarAdapter: calendar,
    now: REFERENCE_NOW
  });

  assert.equal(res.ok, true);
  assert.equal(res.result.status, 'UPDATED');
  assert.equal(calendar.getStats().totalEvents, 1);

  const evt = calendar.getEventById('evt-date-004');
  assert.equal(evt.startIso, '2026-11-28T10:00:00-03:00');
  assert.equal(evt.endIso, '2026-11-28T11:00:00-03:00');
});

// ============================================================================
// TESTE 5: Alteração do procedimento na aba Agendamentos
// ============================================================================
test('5. Alteração do procedimento em Agendamentos (Coluna E) busca a duração do novo serviço em Serviços antes de atualizar', () => {
  const servicos = createMockServicosSheet(); // Alongamento em Fibra de Vidro = 2h (120 min)
  const agendamentos = [
    createAgendamentosHeader(),
    makeAgendamentoRow({
      client: 'Marina Fictícia',
      procedure: 'Alongamento em Fibra de Vidro', // alterado em Col E (mesmo com ID antigo em Col P!)
      date: '26/11/2026',
      time: '14:00',
      appointmentId: 'apt-mock-005',
      serviceId: 'manicure-tradicional-mao', // ID antigo ainda em Col P
      calendarEventId: 'evt-proc-005'
    })
  ];

  const calendar = createInMemoryCalendarAdapter([
    {
      id: 'evt-proc-005',
      title: 'Marina Fictícia — Manicure Tradicional (Mão)',
      startIso: '2026-11-26T14:00:00-03:00',
      endIso: '2026-11-26T14:45:00-03:00',
      description: '[ESPACO_ANGELICA_APT_ID:apt-mock-005]',
      allDay: false
    }
  ]);

  let updatedServiceIdInSheet = null;
  const res = handleSpreadsheetEdit({
    sheetName: 'Agendamentos',
    editedRow: 2,
    editedColumn: AGENDAMENTOS_COL.PROCEDIMENTO + 1, // Coluna E
    servicosRows: servicos,
    agendamentosRows: agendamentos,
    calendarAdapter: calendar,
    now: REFERENCE_NOW,
    onWriteServiceId: (_rowIdx, newId) => {
      updatedServiceIdInSheet = newId;
    }
  });

  assert.equal(res.ok, true);
  assert.equal(res.result.status, 'UPDATED');
  assert.equal(updatedServiceIdInSheet, 'alongamento-fibra');
  assert.equal(calendar.getStats().totalEvents, 1);

  const evt = calendar.getEventById('evt-proc-005');
  assert.equal(evt.title, 'Marina Fictícia — Alongamento em Fibra de Vidro');
  assert.equal(evt.startIso, '2026-11-26T14:00:00-03:00');
  assert.equal(evt.endIso, '2026-11-26T16:00:00-03:00');
});

// ============================================================================
// TESTE 6: Vários agendamentos futuros do mesmo serviço
// ============================================================================
test('6. Vários agendamentos futuros do mesmo serviço são recalculados juntos sem alterar outros serviços', () => {
  const servicos = createMockServicosSheet();
  const agendamentos = [
    createAgendamentosHeader(),
    makeAgendamentoRow({
      client: 'Cliente A',
      procedure: 'Manicure Tradicional (Mão)',
      date: '25/11/2026',
      time: '09:00',
      appointmentId: 'apt-multi-1',
      serviceId: 'manicure-tradicional-mao',
      calendarEventId: 'evt-multi-1'
    }),
    makeAgendamentoRow({
      client: 'Cliente B',
      procedure: 'Manicure Tradicional (Mão)',
      date: '27/11/2026',
      time: '15:30',
      appointmentId: 'apt-multi-2',
      serviceId: '', // sem ID em Col P: resolve pelo título em Col B
      calendarEventId: 'evt-multi-2'
    }),
    makeAgendamentoRow({
      client: 'Cliente C',
      procedure: 'Manicure Tradicional (Mão)',
      date: '02/12/2026',
      time: '11:00',
      appointmentId: 'apt-multi-3',
      serviceId: 'manicure-tradicional-mao',
      calendarEventId: 'evt-multi-3'
    }),
    makeAgendamentoRow({
      client: 'Cliente Outro Serviço',
      procedure: 'Pedicure Spa',
      date: '25/11/2026',
      time: '13:00',
      appointmentId: 'apt-other-1',
      serviceId: 'pedicure-spa',
      calendarEventId: 'evt-other-1'
    })
  ];

  const calendar = createInMemoryCalendarAdapter([
    { id: 'evt-multi-1', title: 'Cliente A — Manicure Tradicional (Mão)', startIso: '2026-11-25T09:00:00-03:00', endIso: '2026-11-25T09:45:00-03:00', allDay: false },
    { id: 'evt-multi-2', title: 'Cliente B — Manicure Tradicional (Mão)', startIso: '2026-11-27T15:30:00-03:00', endIso: '2026-11-27T16:15:00-03:00', allDay: false },
    { id: 'evt-multi-3', title: 'Cliente C — Manicure Tradicional (Mão)', startIso: '2026-12-02T11:00:00-03:00', endIso: '2026-12-02T11:45:00-03:00', allDay: false },
    { id: 'evt-other-1', title: 'Cliente Outro Serviço — Pedicure Spa', startIso: '2026-11-25T13:00:00-03:00', endIso: '2026-11-25T14:00:00-03:00', allDay: false }
  ]);

  // Altera Manicure Tradicional (Mão) de 45 min para 60 min
  servicos[1][SERVICOS_COL.DURACAO] = '60 min';

  const res = syncFutureAppointmentsForServiceDurationChange({
    editedServiceRowIndex: 2,
    servicosRows: servicos,
    agendamentosRows: agendamentos,
    calendarAdapter: calendar,
    now: REFERENCE_NOW
  });

  assert.equal(res.ok, true);
  assert.equal(res.updated, 3);
  assert.equal(calendar.getStats().totalEvents, 4);
  assert.equal(calendar.getStats().createCalls, 0);

  assert.equal(calendar.getEventById('evt-multi-1').endIso, '2026-11-25T10:00:00-03:00');
  assert.equal(calendar.getEventById('evt-multi-2').endIso, '2026-11-27T16:30:00-03:00');
  assert.equal(calendar.getEventById('evt-multi-3').endIso, '2026-12-02T12:00:00-03:00');
  // Outro serviço permanece intacto
  assert.equal(calendar.getEventById('evt-other-1').endIso, '2026-11-25T14:00:00-03:00');
});

// ============================================================================
// TESTE 7: Agendamento passado não deve ser alterado
// ============================================================================
test('7. Agendamento passado não é alterado quando a duração do serviço muda', () => {
  const servicos = createMockServicosSheet();
  const agendamentos = [
    createAgendamentosHeader(),
    makeAgendamentoRow({
      client: 'Cliente Passado',
      procedure: 'Manicure Tradicional (Mão)',
      date: '15/11/2026', // anterior a REFERENCE_NOW (21/11/2026)
      time: '14:00',
      appointmentId: 'apt-past-007',
      serviceId: 'manicure-tradicional-mao',
      calendarEventId: 'evt-past-007'
    })
  ];

  const calendar = createInMemoryCalendarAdapter([
    {
      id: 'evt-past-007',
      title: 'Cliente Passado — Manicure Tradicional (Mão)',
      startIso: '2026-11-15T14:00:00-03:00',
      endIso: '2026-11-15T14:45:00-03:00',
      allDay: false
    }
  ]);

  servicos[1][SERVICOS_COL.DURACAO] = '60 min';

  const res = syncFutureAppointmentsForServiceDurationChange({
    editedServiceRowIndex: 2,
    servicosRows: servicos,
    agendamentosRows: agendamentos,
    calendarAdapter: calendar,
    now: REFERENCE_NOW
  });

  assert.equal(res.ok, true);
  assert.equal(res.updated, 0);
  assert.equal(res.skipped, 1);
  assert.equal(res.logs[0].status, 'SKIPPED_PAST');
  assert.equal(calendar.getStats().updateCalls, 0);
  assert.equal(calendar.getEventById('evt-past-007').endIso, '2026-11-15T14:45:00-03:00');
});

// ============================================================================
// TESTE 8: Agendamento cancelado (e bloqueio de agenda) não devem ser alterados
// ============================================================================
test('8. Agendamento cancelado e bloqueio de agenda não são alterados quando a duração muda', () => {
  const servicos = createMockServicosSheet();
  const agendamentos = [
    createAgendamentosHeader(),
    makeAgendamentoRow({
      client: 'Cliente Cancelada',
      procedure: 'Manicure Tradicional (Mão)',
      date: '26/11/2026',
      time: '10:00',
      status: 'Cancelado',
      appointmentId: 'apt-cancel-008',
      serviceId: 'manicure-tradicional-mao',
      calendarEventId: 'evt-cancel-008'
    }),
    makeAgendamentoRow({
      client: 'Bloqueio Almoço',
      procedure: 'Manicure Tradicional (Mão)',
      date: '26/11/2026',
      time: '12:00',
      status: 'Bloqueado',
      appointmentId: 'apt-block-008',
      serviceId: 'manicure-tradicional-mao',
      calendarEventId: 'evt-block-008'
    })
  ];

  const calendar = createInMemoryCalendarAdapter([
    { id: 'evt-cancel-008', title: 'Cliente Cancelada — Manicure Tradicional (Mão)', startIso: '2026-11-26T10:00:00-03:00', endIso: '2026-11-26T10:45:00-03:00', allDay: false },
    { id: 'evt-block-008', title: 'Bloqueio Almoço', startIso: '2026-11-26T12:00:00-03:00', endIso: '2026-11-26T13:00:00-03:00', allDay: false }
  ]);

  servicos[1][SERVICOS_COL.DURACAO] = '60 min';

  const res = syncFutureAppointmentsForServiceDurationChange({
    editedServiceRowIndex: 2,
    servicosRows: servicos,
    agendamentosRows: agendamentos,
    calendarAdapter: calendar,
    now: REFERENCE_NOW
  });

  assert.equal(res.ok, true);
  assert.equal(res.updated, 0);
  assert.equal(res.skipped, 2);
  assert.equal(res.logs[0].status, 'SKIPPED_CANCELLED');
  assert.equal(res.logs[1].status, 'SKIPPED_SCHEDULE_BLOCK');
  assert.equal(calendar.getStats().updateCalls, 0);
  assert.equal(calendar.getEventById('evt-cancel-008').endIso, '2026-11-26T10:45:00-03:00');
  assert.equal(calendar.getEventById('evt-block-008').endIso, '2026-11-26T13:00:00-03:00');
});

// ============================================================================
// TESTE 9: Serviço não encontrado (e correspondência ambígua)
// ============================================================================
test('9. Serviço não encontrado registra erro, não assume duração padrão e não altera o evento existente', () => {
  const servicos = createMockServicosSheet();
  const row = makeAgendamentoRow({
    client: 'Cliente Serviço Inexistente',
    procedure: 'Procedimento Fantasma',
    date: '29/11/2026',
    time: '15:00',
    appointmentId: 'apt-missing-009',
    serviceId: 'id-inexistente',
    calendarEventId: 'evt-missing-009'
  });

  const calendar = createInMemoryCalendarAdapter([
    {
      id: 'evt-missing-009',
      title: 'Cliente Serviço Inexistente — Antigo',
      startIso: '2026-11-29T15:00:00-03:00',
      endIso: '2026-11-29T15:45:00-03:00',
      allDay: false
    }
  ]);

  const res = syncAppointmentRow({
    rowIndex: 2,
    row,
    servicosRows: servicos,
    calendarAdapter: calendar,
    now: REFERENCE_NOW
  });

  assert.equal(res.status, 'ERROR');
  assert.equal(res.errorCode, 'SERVICE_NOT_FOUND');
  assert.equal(res.changed, false);
  assert.equal(calendar.getStats().updateCalls, 0);
  assert.equal(calendar.getStats().createCalls, 0);
  assert.equal(calendar.getEventById('evt-missing-009').endIso, '2026-11-29T15:45:00-03:00');

  // Também valida correspondência ambígua (dois serviços com mesmo título)
  const ambiguousServicos = [
    ...createMockServicosSheet(),
    ['manicure-duplicada', 'Manicure Tradicional (Mão)', 'Manicure', 45, 'R$ 45,00', '50 min', '', 'Duplicado']
  ];
  const ambRow = makeAgendamentoRow({
    client: 'Cliente Ambígua',
    procedure: 'Manicure Tradicional (Mão)',
    date: '29/11/2026',
    time: '16:00',
    appointmentId: 'apt-amb-009',
    serviceId: '',
    calendarEventId: 'evt-missing-009'
  });
  const ambRes = syncAppointmentRow({
    rowIndex: 3,
    row: ambRow,
    servicosRows: ambiguousServicos,
    calendarAdapter: calendar,
    now: REFERENCE_NOW
  });
  assert.equal(ambRes.status, 'ERROR');
  assert.equal(ambRes.errorCode, 'AMBIGUOUS_SERVICE_MATCH');
  assert.equal(calendar.getStats().updateCalls, 0);
});

// ============================================================================
// TESTE 10: Duração inválida na aba Serviços Coluna F
// ============================================================================
test('10. Duração inválida na aba Serviços (Coluna F) registra erro e não altera o evento existente', () => {
  const servicos = createMockServicosSheet();
  const agendamentos = [
    createAgendamentosHeader(),
    makeAgendamentoRow({
      client: 'Cliente Duração Inválida',
      procedure: 'Manicure Tradicional (Mão)',
      date: '25/11/2026',
      time: '14:00',
      appointmentId: 'apt-invalid-010',
      serviceId: 'manicure-tradicional-mao',
      calendarEventId: 'evt-invalid-010'
    })
  ];

  const calendar = createInMemoryCalendarAdapter([
    {
      id: 'evt-invalid-010',
      title: 'Cliente Duração Inválida — Manicure Tradicional (Mão)',
      startIso: '2026-11-25T14:00:00-03:00',
      endIso: '2026-11-25T14:45:00-03:00',
      allDay: false
    }
  ]);

  const invalidSamples = ['', '0', '-30 min', 'abc', '25/11/2026', 'indisponivel'];
  for (const badDuration of invalidSamples) {
    servicos[1][SERVICOS_COL.DURACAO] = badDuration;
    const res = handleSpreadsheetEdit({
      sheetName: 'Serviços',
      editedRow: 2,
      editedColumn: SERVICOS_COL.DURACAO + 1,
      servicosRows: servicos,
      agendamentosRows: agendamentos,
      calendarAdapter: calendar,
      now: REFERENCE_NOW
    });

    assert.equal(res.ok, false, `Esperava erro para duração inválida "${badDuration}"`);
    assert.equal(res.errors, 1);
    assert.equal(res.logs[0].errorCode, 'INVALID_DURATION');
  }

  // Evento permanece 100% intacto após todas as tentativas inválidas
  assert.equal(calendar.getStats().updateCalls, 0);
  assert.equal(calendar.getStats().createCalls, 0);
  assert.equal(calendar.getEventById('evt-invalid-010').endIso, '2026-11-25T14:45:00-03:00');
});

// ============================================================================
// TESTE 11: Execução repetida sem duplicação de eventos (Idempotência + Exceções Autorizadas)
// ============================================================================
test('11. Execução repetida é idempotente (zero duplicação) e preserva exceções já autorizadas', () => {
  const servicos = createMockServicosSheet();
  servicos[1][SERVICOS_COL.DURACAO] = '60 min';

  const agendamentos = [
    createAgendamentosHeader(),
    makeAgendamentoRow({
      client: 'Cliente Idempotente',
      procedure: 'Manicure Tradicional (Mão)',
      date: '25/11/2026',
      time: '14:00',
      appointmentId: 'apt-idem-011',
      serviceId: 'manicure-tradicional-mao',
      calendarEventId: '' // começa sem ID na Col T -> cria uma única vez e salva ID
    }),
    makeAgendamentoRow({
      client: 'Cliente Exceção Autorizada',
      procedure: 'Manicure Tradicional (Mão)',
      date: '25/11/2026',
      time: '16:00',
      appointmentId: 'apt-exception-keep-011',
      serviceId: 'manicure-tradicional-mao',
      calendarEventId: 'evt-exception-011'
    })
  ];

  const calendar = createInMemoryCalendarAdapter([
    {
      id: 'evt-exception-011',
      title: 'Cliente Exceção Autorizada — Manicure Tradicional (Mão)',
      startIso: '2026-11-25T16:00:00-03:00',
      endIso: '2026-11-25T17:30:00-03:00', // exceção autorizada de 90 min
      allDay: false
    }
  ]);

  const authorizedExceptions = {
    excludeAppointmentIds: ['apt-test-excluded'],
    durationOverrideByAppointmentId: { 'apt-exception-keep-011': 90 }
  };

  const writeEventId = (rowIdx, eventId) => {
    agendamentos[rowIdx - 1][AGENDAMENTOS_COL.CALENDAR_EVENT_ID] = eventId;
  };

  // 1ª execução: cria 1 evento para apt-idem-011 e preserva a exceção autorizada
  const run1 = syncFutureAppointmentsForServiceDurationChange({
    editedServiceRowIndex: 2,
    servicosRows: servicos,
    agendamentosRows: agendamentos,
    calendarAdapter: calendar,
    now: REFERENCE_NOW,
    authorizedExceptions,
    onWriteCalendarEventId: writeEventId
  });
  assert.equal(run1.ok, true);
  assert.equal(run1.created, 1);
  assert.equal(run1.skipped, 1);
  const createdId = agendamentos[1][AGENDAMENTOS_COL.CALENDAR_EVENT_ID];
  assert.ok(createdId, 'Event ID deve ser gravado na Coluna T após criação');

  // 2ª e 3ª execuções repetidas (inclusive simulando perda temporária da Coluna T na 3ª!)
  const run2 = syncFutureAppointmentsForServiceDurationChange({
    editedServiceRowIndex: 2,
    servicosRows: servicos,
    agendamentosRows: agendamentos,
    calendarAdapter: calendar,
    now: REFERENCE_NOW,
    authorizedExceptions,
    onWriteCalendarEventId: writeEventId
  });
  assert.equal(run2.ok, true);
  assert.equal(run2.created, 0);
  assert.equal(run2.updated, 0);
  assert.equal(run2.unchanged, 1);

  // Simula retry mesmo se Col T não tivesse sido salva (guarda por tag [ESPACO_ANGELICA_APT_ID:...])
  agendamentos[1][AGENDAMENTOS_COL.CALENDAR_EVENT_ID] = '';
  const run3 = syncFutureAppointmentsForServiceDurationChange({
    editedServiceRowIndex: 2,
    servicosRows: servicos,
    agendamentosRows: agendamentos,
    calendarAdapter: calendar,
    now: REFERENCE_NOW,
    authorizedExceptions,
    onWriteCalendarEventId: writeEventId
  });
  assert.equal(run3.ok, true);
  assert.equal(run3.created, 0);
  assert.equal(run3.unchanged, 1);
  assert.equal(agendamentos[1][AGENDAMENTOS_COL.CALENDAR_EVENT_ID], createdId);

  // Zero duplicação: total de eventos = 2 (1 criado + 1 exceção inicial)
  assert.equal(calendar.getStats().totalEvents, 2);
  assert.equal(calendar.getStats().createCalls, 1);
  assert.equal(calendar.getStats().updateCalls, 0);
  // Exceção autorizada permanece com 90 min intactos
  assert.equal(calendar.getEventById('evt-exception-011').endIso, '2026-11-25T17:30:00-03:00');
});
