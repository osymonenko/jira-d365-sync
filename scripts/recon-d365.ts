/**
 * Одноразовая разведка D365 для Stage 2 (Web API путь).
 *
 * Запускает Edge с тем же профилем, что и основной CLI (USER_DATA_DIR),
 * ждёт SSO, затем:
 *   1. дергает Web API из контекста страницы (cookie-сессия уже есть):
 *      - последние msdyn_timeentry пользователя со ВСЕМИ полями,
 *      - метаданные msdyn_timeentry (какие атрибуты обязательны),
 *      - bookableresource текущего пользователя, проекты и задачи проекта;
 *   2. снимает структуру weekly-grid iframe (строки/ячейки/переключатель недели).
 *
 * Результат: logs/recon-<ts>/*.json + grid-dom.html. Ничего не создаёт и не меняет.
 */
import * as dotenv from 'dotenv';
dotenv.config();

import { chromium } from 'playwright';
import * as fs from 'fs';
import * as path from 'path';

const d365Url = (process.env['D365_URL'] ?? '').replace(/\/?$/, '/');
const userDataDir = path.resolve(process.env['USER_DATA_DIR'] ?? './edge-profile');

const outDir = path.resolve('logs', `recon-${new Date().toISOString().replace(/[:.]/g, '-')}`);

function save(name: string, data: unknown): void {
  const file = path.join(outDir, name);
  fs.writeFileSync(file, typeof data === 'string' ? data : JSON.stringify(data, null, 2), 'utf8');
  console.log(`  saved ${name} (${fs.statSync(file).size} bytes)`);
}

async function main(): Promise<void> {
  if (!d365Url.startsWith('http')) throw new Error('D365_URL не задан в .env');
  fs.mkdirSync(outDir, { recursive: true });
  console.log(`Recon output: ${outDir}`);

  const context = await chromium.launchPersistentContext(userDataDir, {
    channel: 'msedge',
    headless: false,
    args: ['--start-maximized'],
    ignoreDefaultArgs: ['--enable-automation'],
  });
  const page = context.pages()[0] ?? (await context.newPage());

  console.log('Открываю D365, жду навбар (до 5 мин, при необходимости залогиньтесь вручную)...');
  await page.goto(`${d365Url}main.aspx?pagetype=entitylist&etn=msdyn_timeentry`, {
    waitUntil: 'domcontentloaded',
    timeout: 60000,
  });
  await page
    .locator('[aria-label="Main Navigation"], nav[aria-label*="navigation" i], [data-id="navbar-container"]')
    .first()
    .waitFor({ state: 'visible', timeout: 5 * 60 * 1000 });
  console.log('Навбар найден.');

  if (page.url().includes('pagetype=apps')) {
    console.log('Страница выбора приложения — кликаю "Team Member\'s PO Hub"...');
    await page.locator('[aria-label*="Team Member" i], h2:has-text("Team Member")').first().click();
    await page.waitForTimeout(4000);
    await page.goto(`${d365Url}main.aspx?pagetype=entitylist&etn=msdyn_timeentry`, {
      waitUntil: 'domcontentloaded',
      timeout: 60000,
    });
  }
  await page.waitForTimeout(6000);

  // ── 1. Web API из контекста страницы ────────────────────────────────────
  console.log('\n[1] Web API probes...');
  const api = (await page.evaluate(`(async () => {
    const base = window.location.origin + '/api/data/v9.2/';
    const get = async (q) => {
      const r = await fetch(base + q, {
        credentials: 'include',
        headers: {
          'Accept': 'application/json',
          'OData-MaxVersion': '4.0',
          'OData-Version': '4.0',
          'Prefer': 'odata.include-annotations="*"',
        },
      });
      const text = await r.text();
      if (!r.ok) return { __error: r.status + ' ' + r.statusText, body: text.slice(0, 2000) };
      try { return JSON.parse(text); } catch (e) { return { __parseError: text.slice(0, 2000) }; }
    };
    const out = {};
    out.origin = window.location.origin;
    out.url = window.location.href;
    try {
      const g = window.Xrm && window.Xrm.Utility && window.Xrm.Utility.getGlobalContext
        ? window.Xrm.Utility.getGlobalContext() : null;
      out.userId = g ? g.userSettings.userId : null;
      out.userName = g ? g.userSettings.userName : null;
      out.clientUrl = g ? g.getClientUrl() : null;
      out.orgUniqueName = g ? g.organizationSettings.uniqueName : null;
    } catch (e) { out.xrmError = String(e); }

    out.whoami = await get('WhoAmI');
    out.recentTimeEntries = await get("msdyn_timeentries?$top=5&$orderby=createdon desc");
    out.bookableResources = await get("bookableresources?$top=10&$select=bookableresourceid,name,_userid_value,resourcetype");
    out.projects = await get("msdyn_projects?$top=20&$select=msdyn_projectid,msdyn_subject,statecode");
    out.projectTasks = await get("msdyn_projecttasks?$top=10&$select=msdyn_projecttaskid,msdyn_subject,_msdyn_project_value");
    out.timeEntryMeta = await get(
      "EntityDefinitions(LogicalName='msdyn_timeentry')/Attributes?$select=LogicalName,SchemaName,AttributeType,RequiredLevel,IsValidForCreate,DisplayName"
    );
    out.timeEntrySetName = await get(
      "EntityDefinitions(LogicalName='msdyn_timeentry')?$select=EntitySetName,PrimaryIdAttribute,PrimaryNameAttribute"
    );
    out.timeEntryRelations = await get(
      "EntityDefinitions(LogicalName='msdyn_timeentry')/ManyToOneRelationships?$select=ReferencingAttribute,ReferencedEntity,ReferencingEntityNavigationPropertyName"
    );
    return out;
  })()`)) as Record<string, unknown>;
  save('webapi.json', api);

  // ── 2. Weekly grid DOM ──────────────────────────────────────────────────
  console.log('\n[2] Weekly grid DOM...');
  const frames = page.frames().map((f) => ({ name: f.name(), url: f.url() }));
  save('frames.json', frames);

  const gridFrame =
    page.frames().find((f) => f.url().includes('TimeEntryGridControl')) ??
    page.frames().find((f) => f !== page.mainFrame() && f.url().startsWith('http'));

  if (!gridFrame) {
    console.log('  ⚠ Grid iframe не найден');
  } else {
    console.log(`  grid frame: ${gridFrame.url().slice(0, 120)}`);
    const dump = (await gridFrame.evaluate(`(() => {
      const nodes = [];
      document.querySelectorAll('*').forEach((el) => {
        const aria = el.getAttribute('aria-label');
        const role = el.getAttribute('role');
        const did = el.getAttribute('data-id');
        const title = el.getAttribute('title');
        if (!aria && !role && !did && !title) return;
        nodes.push({
          tag: el.tagName.toLowerCase(),
          role, aria, did, title,
          cls: (el.className && el.className.toString ? el.className.toString() : '').slice(0, 80),
          text: (el.textContent || '').trim().slice(0, 70),
        });
      });
      return {
        title: document.title,
        count: nodes.length,
        nodes: nodes.slice(0, 1200),
        html: document.body.innerHTML.slice(0, 200000),
      };
    })()`)) as { nodes: unknown[]; html: string; count: number };
    save('grid-nodes.json', { count: dump.count, nodes: dump.nodes });
    save('grid-dom.html', dump.html);
  }

  console.log(`\nГотово. Смотрите ${outDir}`);
  await context.close();
}

main().catch((e) => {
  console.error('Recon failed:', e);
  process.exit(1);
});
