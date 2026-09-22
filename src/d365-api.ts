/**
 * Создание Time Entries через D365 Web API из авторизованной вкладки.
 *
 * Зачем в обход UI: Quick Create — это ~15–20 с на запись, а здесь месяц
 * закрывается за пару минут. Ни Stage 1, ни Stage 2, ни переключения недель:
 * дата записи задаётся полем, а не тем, какую неделю показывает грид.
 *
 * Форма payload'а выверена эмпирически (scripts/d365-bisect-create.ts,
 * 2026-09-22). Тенант отбивал создание ошибкой Field Service
 * "0x80040265 ... solution upgrade did not successfully complete", и настоящая
 * причина пряталась в её же хвосте: "Time entry is missing the required field
 * Date". Без msdyn_date запись не создаётся, с ним — HTTP 201.
 *
 * Всё, чего нет в payload, тенант проставляет сам (проверено чтением созданной
 * записи): msdyn_bookableresource, msdyn_timeentrysettingId, msdyn_type=Work,
 * msdyn_entrystatus=Draft, amc_payableduration, amc_enablecutoff. Поэтому шлём
 * ровно проверенный минимум — чем меньше полей, тем меньше поводов для плагинов
 * тенанта сработать неожиданно.
 */
import { Page } from 'playwright';

export const HOURS_TYPE_PAYABLE = 100000000;
export const HOURS_TYPE_BILLABLE = 100000001;
const AMC_TYPE_WORK = 100000000;
const TASK_TYPE_GENERAL = 100000000;

export interface D365Context {
  projectId: string;
  projectName: string;
  teamId: string;
  teamName: string;
  userId: string;
}

export interface EntrySpec {
  /** YYYY-MM-DD */
  date: string;
  task: string;
  minutes: number;
}

interface ODataError {
  error?: { code?: string; message?: string };
}

interface ODataList<T> {
  value?: T[];
  '@odata.nextLink'?: string;
}

export type Log = (msg: string) => void;

export class D365Api {
  private readonly page: Page;
  private readonly log: Log;

  constructor(page: Page, log: Log) {
    this.page = page;
    this.log = log;
  }

  // ── транспорт ─────────────────────────────────────────────────────────────

  /** GET по OData-пути (или по полному nextLink). */
  private async get<T>(query: string): Promise<T> {
    const res = (await this.page.evaluate(
      `(async () => {
        const q = ${JSON.stringify(query)};
        const url = q.startsWith('http') ? q : window.location.origin + '/api/data/v9.2/' + q;
        const r = await fetch(url, {
          credentials: 'include',
          headers: {
            'Accept': 'application/json',
            'OData-MaxVersion': '4.0',
            'OData-Version': '4.0',
            'Prefer': 'odata.include-annotations="*"',
          },
        });
        const text = await r.text();
        let parsed = null; try { parsed = JSON.parse(text); } catch (e) {}
        if (!r.ok) return { __http: r.status, __body: parsed || text.slice(0, 1500) };
        return parsed;
      })()`,
    )) as T & { __http?: number; __body?: unknown };

    if (res && res.__http) {
      throw new Error(`GET ${query.slice(0, 80)} → HTTP ${res.__http}: ${describeError(res.__body)}`);
    }
    return res;
  }

  /** GET всех страниц: D365 отдаёт максимум 5000 записей за раз. */
  private async getAll<T>(query: string, limit = 5000): Promise<T[]> {
    const out: T[] = [];
    let next: string | undefined = query;
    while (next && out.length < limit) {
      const page: ODataList<T> = await this.get<ODataList<T>>(next);
      out.push(...(page.value ?? []));
      next = page['@odata.nextLink'];
    }
    return out;
  }

  private async post(entitySet: string, body: Record<string, unknown>): Promise<string> {
    const res = (await this.page.evaluate(
      `(async () => {
        const r = await fetch(window.location.origin + '/api/data/v9.2/' + ${JSON.stringify(entitySet)}, {
          method: 'POST',
          credentials: 'include',
          headers: {
            'Content-Type': 'application/json; charset=utf-8',
            'Accept': 'application/json',
            'OData-MaxVersion': '4.0',
            'OData-Version': '4.0',
            'Prefer': 'return=representation',
          },
          body: JSON.stringify(${JSON.stringify(body)}),
        });
        const text = await r.text();
        let parsed = null; try { parsed = JSON.parse(text); } catch (e) {}
        return { ok: r.ok, status: r.status, body: parsed || text.slice(0, 1500) };
      })()`,
    )) as { ok: boolean; status: number; body: Record<string, unknown> | string };

    if (!res.ok) {
      throw new Error(`POST ${entitySet} → HTTP ${res.status}: ${describeError(res.body)}`);
    }
    const rec = res.body as Record<string, unknown>;
    const idKey = Object.keys(rec).find((k) => k.endsWith('id') && typeof rec[k] === 'string');
    return (idKey ? (rec[idKey] as string) : '') || '';
  }

  // ── контекст ──────────────────────────────────────────────────────────────

  /**
   * Проект/команда берутся из последней собственной записи, а не из .env: если
   * человека переведут на другой проект, инструмент подхватит это сам, и в
   * конфиге не заведётся протухший GUID.
   */
  async resolveContext(): Promise<D365Context> {
    const who = await this.get<{ UserId: string }>('WhoAmI');
    const userId = who.UserId;

    const F = '@OData.Community.Display.V1.FormattedValue';
    const recent = await this.get<ODataList<Record<string, string>>>(
      `msdyn_timeentries?$top=1&$orderby=createdon desc` +
        `&$filter=_owninguser_value eq ${userId} and _amc_projectteam_value ne null` +
        `&$select=_msdyn_project_value,_amc_projectteam_value`,
    );
    const last = recent.value?.[0];
    if (!last) {
      throw new Error(
        'Не нашёл ни одной вашей Time Entry — не из чего взять проект и команду. ' +
          'Создайте одну запись вручную через Quick Create, дальше инструмент подхватит контекст.',
      );
    }

    return {
      userId,
      projectId: last['_msdyn_project_value'],
      projectName: last['_msdyn_project_value' + F] ?? '(без имени)',
      teamId: last['_amc_projectteam_value'],
      teamName: last['_amc_projectteam_value' + F] ?? '(без имени)',
    };
  }

  // ── задачи проекта ────────────────────────────────────────────────────────

  /** Индекс «имя задачи в нижнем регистре → id» по проекту. */
  async loadTaskIndex(projectId: string): Promise<Map<string, string>> {
    const rows = await this.getAll<{ amc_importedprojecttaskid: string; amc_name: string }>(
      `amc_importedprojecttasks?$filter=_amc_project_value eq ${projectId} and statecode eq 0` +
        `&$select=amc_name&$top=5000`,
    );
    const index = new Map<string, string>();
    for (const row of rows) {
      const key = (row.amc_name ?? '').trim().toLowerCase();
      // Дубли по имени в тенанте есть; берём первую — какая именно, неважно,
      // это лишь ярлык для отчёта.
      if (key && !index.has(key)) index.set(key, row.amc_importedprojecttaskid);
    }
    return index;
  }

  async createTask(name: string, projectId: string): Promise<string> {
    return this.post('amc_importedprojecttasks', {
      amc_name: name,
      amc_tasktype: TASK_TYPE_GENERAL,
      'amc_project@odata.bind': `/msdyn_projects(${projectId})`,
    });
  }

  // ── записи времени ────────────────────────────────────────────────────────

  /**
   * Ключи уже существующих записей за период: `дата|taskId|тип часов`.
   * Нужно, чтобы повторный запуск не задваивал табель.
   */
  async existingEntryKeys(userId: string, fromDate: string, toDate: string): Promise<Set<string>> {
    const rows = await this.getAll<Record<string, string | number>>(
      `msdyn_timeentries?$filter=_owninguser_value eq ${userId}` +
        ` and msdyn_datetimezoneindependent ge ${fromDate}` +
        ` and msdyn_datetimezoneindependent le ${toDate}` +
        `&$select=msdyn_datetimezoneindependent,amc_hourstype,_amc_projecttask_value&$top=5000`,
    );
    const keys = new Set<string>();
    for (const row of rows) {
      const date = String(row['msdyn_datetimezoneindependent'] ?? '').slice(0, 10);
      keys.add(`${date}|${row['_amc_projecttask_value']}|${row['amc_hourstype']}`);
    }
    return keys;
  }

  async createEntry(spec: {
    date: string;
    minutes: number;
    taskId: string;
    hoursType: number;
    ctx: D365Context;
  }): Promise<string> {
    return this.post('msdyn_timeentries', {
      msdyn_datetimezoneindependent: spec.date,
      // Без этого поля тенант отвечает 0x80040265 — см. шапку файла.
      msdyn_date: `${spec.date}T00:00:00Z`,
      msdyn_duration: spec.minutes,
      amc_hourstype: spec.hoursType,
      amc_type: AMC_TYPE_WORK,
      'msdyn_project@odata.bind': `/msdyn_projects(${spec.ctx.projectId})`,
      'amc_projecttask@odata.bind': `/amc_importedprojecttasks(${spec.taskId})`,
      'amc_ProjectTeam@odata.bind': `/amc_projectteams(${spec.ctx.teamId})`,
    });
  }
}

function describeError(body: unknown): string {
  const err = (body as ODataError)?.error;
  if (err?.message) return `${err.code ?? ''} ${err.message}`.trim();
  return typeof body === 'string' ? body : JSON.stringify(body).slice(0, 400);
}

/** "9/22/2026" -> "2026-09-22" (D365 Web API принимает только ISO). */
export function toIsoDate(d365Date: string): string {
  const [m, d, y] = d365Date.split('/').map((p) => parseInt(p, 10));
  if (!m || !d || !y) throw new Error(`Не разобрал дату "${d365Date}" (ожидал M/D/YYYY)`);
  return `${y}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
}

export function hoursToMinutes(hours: number): number {
  return Math.round(hours * 60);
}
