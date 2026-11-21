export declare const AGENDAMENTOS_COL: Readonly<{
  TIMESTAMP: 0;
  CLIENTE: 1;
  PROCEDIMENTO: 4;
  DATA_AGENDADA: 5;
  HORARIO_AGENDADO: 6;
  SINAL: 7;
  STATUS: 8;
  APPOINTMENT_ID: 9;
  ADDITIONAL_IDS: 13;
  SERVICE_ID: 15;
  CALENDAR_EVENT_ID: 19;
}>;

export declare const SERVICOS_COL: Readonly<{
  ID: 0;
  TITULO: 1;
  CATEGORIA: 2;
  PRECO: 3;
  PRECO_FORMATADO: 4;
  DURACAO: 5;
  BADGE: 6;
  DESCRICAO: 7;
}>;

export declare const TIMEZONE_OFFSET: string;
export declare const APT_TAG_PREFIX: string;

export declare function normalizeText(value: unknown): string;
export declare function parseOfficialDurationMinutes(value: unknown): number | null;
export declare function parseSheetDateKey(value: unknown): string | null;
export declare function parseSheetTimeKey(value: unknown): string | null;
export declare function calculateEventInterval(
  dateValue: unknown,
  startTimeValue: unknown,
  durationMinutes: number,
  tzOffset?: string
): Record<string, unknown>;
export declare function indexServicesFromSheet(servicosRows: unknown[][]): Record<string, unknown>;
export declare function resolveServiceDurationFromSheet(
  servicosRows: unknown[][],
  query: { serviceId?: string; procedureTitle?: string; additionalServiceIds?: string[] }
): Record<string, unknown>;
export declare function classifyRowGuards(
  row: unknown[],
  authorizedExceptions?: Record<string, unknown>
): { skip: boolean; reason?: string; message?: string };
export declare function syncAppointmentRow(options: Record<string, unknown>): Record<string, unknown>;
export declare function syncFutureAppointmentsForServiceDurationChange(
  options: Record<string, unknown>
): Record<string, unknown>;
export declare function handleSpreadsheetEdit(options: Record<string, unknown>): Record<string, unknown>;
export declare function createInMemoryCalendarAdapter(initialEvents?: Array<Record<string, unknown>>): Record<string, unknown>;
