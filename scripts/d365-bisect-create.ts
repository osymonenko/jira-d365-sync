/**
 * Бисект payload'а для создания msdyn_timeentry.
 *
 *   node --require ts-node/register/transpile-only scripts/d365-bisect-create.ts [Имя задачи]
 *
 * Зачем: тенант отбивал создание записи ошибкой Field Service
 * "0x80040265 ... solution upgrade did not successfully complete" — сообщение
 * уводит в сторону, настоящая причина в его хвосте ("Time entry is missing the
 * required field Date"). Скрипт гоняет варианты payload'а от минимального к
 * полному и останавливается на первом успехе, так что при следующей поломке
 * виновное поле находится за один прогон, а не за день гаданий.
 *
 * Создаёт максимум ОДНУ запись (15 минут, статус Draft) — на первом успешном
 * варианте. Проект, команда, ресурс и задача берутся из вашей последней записи
 * в D365: никаких GUID'ов в исходниках.
 */
import { openD365, odata } from './d365-query';
import { D365Api } from '../src/d365-api';

const TASK_NAME = process.argv[2] ?? 'Internal Daily meeting';
const MIN = 15;

type Body = Record<string, unknown>;

async function post(
  page: import('playwright').Page,
  body: Body,
): Promise<{ status: number; ok: boolean; body: unknown }> {
  return (await page.evaluate(
    `(async () => {
      const r = await fetch(window.location.origin + '/api/data/v9.2/msdyn_timeentries', {
        method: 'POST', credentials: 'include',
        headers: {
          'Content-Type': 'application/json; charset=utf-8', 'Accept': 'application/json',
          'OData-MaxVersion': '4.0', 'OData-Version': '4.0', 'Prefer': 'return=representation',
        },
        body: JSON.stringify(${JSON.stringify(body)}),
      });
      const t = await r.text();
      let p = null; try { p = JSON.parse(t); } catch (e) {}
      return { status: r.status, ok: r.ok, body: p || t.slice(0, 1200) };
    })()`,
  )) as { status: number; ok: boolean; body: unknown };
}

async function main(): Promise<void> {
  const { page, close } = await openD365();
  try {
    const api = new D365Api(page, (m) => console.log(m));
    const ctx = await api.resolveContext();
    console.log(`Проект: ${ctx.projectName} | команда: ${ctx.teamName}`);

    const tasks = await api.loadTaskIndex(ctx.projectId);
    const taskId = tasks.get(TASK_NAME.trim().toLowerCase());
    if (!taskId) {
      throw new Error(`Задача "${TASK_NAME}" не найдена на проекте — передайте другое имя аргументом`);
    }

    // Дата — завтра: свежая, но не задним числом, чтобы тестовая запись не
    // затерялась среди настоящих и её легко было удалить.
    const d = new Date();
    d.setDate(d.getDate() + 1);
    const DATE = d.toISOString().slice(0, 10);
    console.log(`Задача: ${TASK_NAME} | дата: ${DATE}\n`);

    const base: Body = {
      msdyn_datetimezoneindependent: DATE,
      msdyn_duration: MIN,
      amc_hourstype: 100000000,
      amc_type: 100000000,
      'msdyn_project@odata.bind': `/msdyn_projects(${ctx.projectId})`,
      'amc_projecttask@odata.bind': `/amc_importedprojecttasks(${taskId})`,
      'amc_ProjectTeam@odata.bind': `/amc_projectteams(${ctx.teamId})`,
    };

    const variants: { name: string; body: Body }[] = [
      { name: '1_minimal', body: { ...base } },
      { name: '2_plus_msdyn_type', body: { ...base, msdyn_type: 192350000 } },
      { name: '3_plus_msdyn_date', body: { ...base, msdyn_date: `${DATE}T00:00:00Z` } },
      {
        name: '4_plus_durations',
        body: {
          ...base,
          msdyn_date: `${DATE}T00:00:00Z`,
          amc_payableduration: MIN,
          amc_billableduration: 0,
          amc_enablecutoff: true,
        },
      },
      {
        name: '5_full',
        body: {
          ...base,
          msdyn_type: 192350000,
          msdyn_date: `${DATE}T00:00:00Z`,
          msdyn_entrystatus: 192350000,
          amc_payableduration: MIN,
          amc_billableduration: 0,
          amc_copytobillableduration: true,
          amc_enablecutoff: true,
        },
      },
    ];

    for (const v of variants) {
      process.stdout.write(`  ${v.name.padEnd(20)} `);
      const res = await post(page, v.body);
      if (res.ok) {
        const rec = res.body as Record<string, unknown>;
        console.log(`✓ HTTP ${res.status}  id=${rec['msdyn_timeentryid']}`);
        console.log(`\nУСПЕХ на варианте ${v.name}. Payload:`);
        console.log(JSON.stringify(v.body, null, 2));

        console.log('\nЖду 6с (вдруг плагин создаст Billable-двойник) и читаю дату...');
        await page.waitForTimeout(6000);
        const F = '@OData.Community.Display.V1.FormattedValue';
        const back = (await odata(
          page,
          `msdyn_timeentries?$filter=_owninguser_value eq ${ctx.userId} and msdyn_datetimezoneindependent eq ${DATE}` +
            `&$select=msdyn_timeentryid,msdyn_datetimezoneindependent,msdyn_duration,amc_hourstype,` +
            `amc_payableduration,amc_billableduration,msdyn_entrystatus,amc_copytobillableduration`,
        )) as { value?: Record<string, unknown>[] };
        for (const r of back.value ?? []) {
          console.log(
            `    ${r['msdyn_datetimezoneindependent']} | ${r['amc_hourstype' + F]} | dur=${r['msdyn_duration']} ` +
              `pay=${r['amc_payableduration']} bill=${r['amc_billableduration']} | ${r['msdyn_entrystatus' + F]}`,
          );
        }
        console.log(`Итого на ${DATE}: ${(back.value ?? []).length} — удалите тестовую запись в D365.`);
        return;
      }
      const err = (res.body as { error?: { code?: string; message?: string } })?.error;
      console.log(`✗ HTTP ${res.status}  ${err?.code ?? ''} ${String(err?.message ?? res.body).slice(0, 110)}`);
    }
    console.log('\nНи один вариант не прошёл — создание через Web API заблокировано плагином тенанта.');
  } finally {
    await close();
  }
}

main().catch((e) => {
  console.error('Bisect failed:', e);
  process.exit(1);
});
