/**
 * Переиспользуемый запросник к D365 Web API через авторизованную вкладку Edge.
 *
 *   node --require ts-node/register/transpile-only scripts/d365-query.ts <queries.json> [outDir]
 *
 * queries.json: { "<имя>": "<OData-путь после /api/data/v9.2/>", ... }
 * Результат: <outDir>/<имя>.json. Только чтение — ничего не создаёт.
 */
import * as dotenv from 'dotenv';
dotenv.config();

import { chromium, Page } from 'playwright';
import * as fs from 'fs';
import * as path from 'path';

const d365Url = (process.env['D365_URL'] ?? '').replace(/\/?$/, '/');
const userDataDir = path.resolve(process.env['USER_DATA_DIR'] ?? './edge-profile');
const queriesFile = process.argv[2];
const outDir = path.resolve(process.argv[3] ?? path.join('logs', `query-${Date.now()}`));

export async function openD365(): Promise<{ page: Page; close: () => Promise<void> }> {
  const context = await chromium.launchPersistentContext(userDataDir, {
    channel: 'msedge',
    headless: false,
    args: ['--start-maximized'],
    ignoreDefaultArgs: ['--enable-automation'],
  });
  const page = context.pages()[0] ?? (await context.newPage());
  await page.goto(d365Url, { waitUntil: 'domcontentloaded', timeout: 60000 });

  const waitForXrm = (timeout: number): Promise<unknown> =>
    page.waitForFunction(
      '!!(window.Xrm && window.Xrm.Utility && window.Xrm.Utility.getGlobalContext)',
      undefined,
      { timeout, polling: 1000 },
    );

  // Признак "клиент загрузился" — window.Xrm, а не разметка навбара: SSO делает
  // цепочку редиректов (login → pagetype=apps → appid=…), и aria-label шелла
  // отличается между сборками D365.
  await waitForXrm(5 * 60 * 1000);

  // Хаб выбора приложения: без приложения Xrm есть, но контекста сущностей нет.
  if (page.url().includes('pagetype=apps')) {
    console.log(`Выбираю приложение "Team Member's PO Hub"...`);
    await page
      .locator('[aria-label*="Team Member" i], h2:has-text("Team Member")')
      .first()
      .click({ timeout: 30000 });
    await page.waitForTimeout(5000);
    await waitForXrm(120000);
  }
  console.log(`D365 готов: ${page.url().slice(0, 110)}`);
  return { page, close: () => context.close() };
}

export async function odata(page: Page, query: string): Promise<unknown> {
  return page.evaluate(
    `(async () => {
      const r = await fetch(window.location.origin + '/api/data/v9.2/' + ${JSON.stringify(query)}, {
        credentials: 'include',
        headers: {
          'Accept': 'application/json',
          'OData-MaxVersion': '4.0',
          'OData-Version': '4.0',
          'Prefer': 'odata.include-annotations="*"',
        },
      });
      const text = await r.text();
      if (!r.ok) return { __error: r.status + ' ' + r.statusText, body: text.slice(0, 3000) };
      try { return JSON.parse(text); } catch (e) { return { __parseError: text.slice(0, 3000) }; }
    })()`,
  );
}

async function main(): Promise<void> {
  if (!queriesFile) throw new Error('Укажите путь к queries.json');
  const queries = JSON.parse(fs.readFileSync(queriesFile, 'utf8')) as Record<string, string>;
  fs.mkdirSync(outDir, { recursive: true });

  const { page, close } = await openD365();
  try {
    for (const [name, q] of Object.entries(queries)) {
      process.stdout.write(`  ${name} ... `);
      const res = await odata(page, q);
      const file = path.join(outDir, `${name}.json`);
      fs.writeFileSync(file, JSON.stringify(res, null, 2), 'utf8');
      const err = (res as { __error?: string }).__error;
      const n = (res as { value?: unknown[] }).value?.length;
      console.log(err ? `✗ ${err}` : `✓ ${n ?? 1} запись(ей)`);
    }
  } finally {
    await close();
  }
  console.log(`\nГотово: ${outDir}`);
}

if (require.main === module) {
  main().catch((e) => {
    console.error('Query failed:', e);
    process.exit(1);
  });
}
