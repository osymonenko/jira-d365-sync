import * as dotenv from 'dotenv';
dotenv.config();

import { Command } from 'commander';
import { firstEntriesPerTask, parseTimesheet, ParseDiagnostics } from './excel';
import { D365Client } from './d365';
import { D365Api, hoursToMinutes, toIsoDate, HOURS_TYPE_PAYABLE, HOURS_TYPE_BILLABLE } from './d365-api';
import { runPreflight } from './preflight';
import { TimeEntry } from './types';
import { createLogger, createRunId, Logger } from './logger';

const program = new Command();

program
  .name('d365-time-entry')
  .description('Automatically fill Dynamics 365 time entries from an Excel timesheet')
  .requiredOption('-f, --file <path>', 'Path to timesheet.xlsx')
  .option('-w, --week <date>', 'Only process the week containing this Monday (YYYY-MM-DD). Omit to process all weeks.')
  .option('--stage1-only', '[kept for GUI compatibility] alias for default behaviour — Stage 2 is not implemented yet')
  .option('--stage2-only', '[deprecated] Stage 2 not implemented yet — see TODO in src/d365.ts')
  .option('-m, --mode <mode>', 'api = Web API (fast, fills every day in one pass) | ui = Quick Create panel', 'api')
  .option('--preflight-only', 'Run preflight checks only (no browser automation)')
  .option('--list-tasks', 'Parse Excel and list all tasks without running automation')
  .parse(process.argv);

const opts = program.opts<{
  file: string;
  week?: string;
  mode?: string;
  stage1Only?: boolean;
  stage2Only?: boolean;
  preflightOnly?: boolean;
  listTasks?: boolean;
}>();

async function main(): Promise<void> {
  // Parse week filter early — used by both --list-tasks and normal run
  let weekFilter: Date | undefined;
  if (opts.week) {
    weekFilter = new Date(opts.week);
    if (isNaN(weekFilter.getTime())) {
      console.error(`Error: --week "${opts.week}" is not a valid date (expected YYYY-MM-DD)`);
      process.exit(1);
    }
  }

  // --list-tasks: parse Excel and print task names, no browser or D365_URL needed
  if (opts.listTasks) {
    let entries: TimeEntry[];
    const diag = {} as ParseDiagnostics;
    try {
      entries = parseTimesheet(opts.file, weekFilter, diag);
    } catch (err) {
      console.error('Failed to parse Excel file:', (err as Error).message);
      process.exit(1);
    }
    printParseReport(diag);
    const taskEntries = firstEntriesPerTask(entries);
    if (taskEntries.length === 0) {
      console.log('No tasks found.');
    } else {
      console.log(`${taskEntries.length} task(s):`);
      taskEntries.forEach((e) => console.log(`  • ${e.task}`));
    }
    return;
  }

  const d365Url = process.env['D365_URL'];
  const browserMode = process.env['BROWSER_MODE'] ?? 'chrome-profile';
  const cdpUrl = process.env['CDP_URL'] ?? 'http://localhost:9222';
  const userDataDir = process.env['USER_DATA_DIR'] ?? './browser-profile';
  // COPY_TO_BILLABLE_DURATION=true makes each new Time Entry set its "Copy to
  // Billable Duration" toggle to Yes. Anything other than a truthy string
  // ("true"/"1"/"yes") leaves the toggle at the D365 default (current behaviour).
  const copyToBillable = /^(true|1|yes)$/i.test(process.env['COPY_TO_BILLABLE_DURATION'] ?? '');

  if (!d365Url) {
    console.error('Error: D365_URL is not set. Add it to .env or set the environment variable.');
    process.exit(1);
  }

  const mode = (opts.mode ?? 'api').toLowerCase();
  if (mode !== 'api' && mode !== 'ui') {
    console.error(`Error: --mode "${opts.mode}" is not supported. Use "api" or "ui".`);
    process.exit(1);
  }

  if (opts.stage2Only) {
    console.error('Error: --stage2-only is no longer supported. Stage 2 (grid-cell fill) is not implemented yet — see TODO in src/d365.ts.');
    process.exit(1);
  }

  // Preflight checks
  const preflightOk = await runPreflight({ d365Url, browserMode, cdpUrl, userDataDir, excelFile: opts.file });
  if (!preflightOk || opts.preflightOnly) {
    process.exit(preflightOk ? 0 : 1);
  }

  // ── Stage: parse Excel ────────────────────────────────────────────────────
  console.log(`\n[1/4] Reading timesheet: ${opts.file}`);
  if (weekFilter) console.log(`      Week filter: --week ${opts.week} (matches the week whose Monday = ${opts.week})`);
  else console.log(`      Week filter: none (processing all weeks in the file)`);
  let entries: TimeEntry[];
  const diag = {} as ParseDiagnostics;
  try {
    entries = parseTimesheet(opts.file, weekFilter, diag);
  } catch (err) {
    console.error('      ✗ Failed to parse Excel file:', (err as Error).message);
    process.exit(1);
  }

  // Always report what the parser saw — makes "0 entries" self-explanatory.
  printParseReport(diag);

  if (entries.length === 0) {
    console.log('\n[!] No time entries to submit for this selection. Browser will NOT launch.');
    if (diag.weekBlocks === 0) {
      console.log('    Reason: no week blocks were recognised in the sheet. Expected a row whose');
      console.log('    column A holds a date (week start), followed by task rows (task name in');
      console.log(`    column C, hours in columns D–H). Sheet read: "${diag.sheetName}", ${diag.totalRows} rows.`);
    } else if (weekFilter && !diag.weeks.some((w) => w.matchedFilter)) {
      console.log(`    Reason: no week in the file matches --week ${opts.week}.`);
      console.log('    The file has these week(s) — submit using one of their dates:');
      diag.weeks.forEach((w) => console.log(`      • week of ${w.monday} (${w.entries} entr${w.entries === 1 ? 'y' : 'ies'})`));
    } else {
      console.log('    Reason: the selected week has task rows but no hour cells > 0');
      console.log('    (columns D–H must contain numbers, e.g. 0.5, 1, 2).');
    }
    return;
  }

  const stage1Entries = firstEntriesPerTask(entries);
  console.log(`\n[2/4] Found ${entries.length} time entries across ${stage1Entries.length} unique task(s):`);
  stage1Entries.forEach((e) => console.log(`  • ${e.date} | ${e.task} | ${e.hours}h`));
  // TODO: Stage 2 (оставшиеся дни многодневных задач) будет реализован отдельно
  // через клик по ячейкам в "All Weekly Time Entries" weekly grid — после
  // Stage 1 строки задач уже существуют в гриде, нужно только проставить часы
  // в ячейках пересечения день × задача.

  // Logger writes to logs/<runId>/run.log alongside Playwright trace + error screenshots.
  const runId = createRunId();
  const logger: Logger = createLogger(runId);
  logger.log(`Run ${runId} started — logs in ${logger.runDir}`);

  // ── Stage: launch browser ─────────────────────────────────────────────────
  console.log(`\n[3/4] Launching browser (mode: ${browserMode})...`);
  console.log(`      Copy to Billable Duration: ${copyToBillable ? 'ENABLED (toggles set to Yes)' : 'off (D365 default)'}`);
  const client = new D365Client(d365Url, userDataDir, browserMode, cdpUrl, logger, copyToBillable);
  await client.launch();

  // Режим api не открывает Quick Create вообще: записи создаются через Web API
  // из уже авторизованной вкладки. Дата — это поле записи, а не то, какую
  // неделю показывает грид, поэтому и первый день задачи, и все остальные
  // закрываются одним проходом, без переключения недель.
  if (mode === 'api') {
    let ok = false;
    try {
      ok = await runApiMode(client, entries, runId, logger.runDir);
    } finally {
      await client.close();
      await logger.close();
    }
    process.exit(ok ? 0 : 1);
  }

  const results = { created: 0, existing: 0, failed: 0 };
  const failures: { entry: string; error: string }[] = [];

  try {
    // ── Stage 1: First-day entry per task (with inline task creation) ────────
    console.log(`\n[4/4] Stage 1: Submitting first-day entry for ${stage1Entries.length} task(s) ──`);
    for (let i = 0; i < stage1Entries.length; i++) {
      const entry = stage1Entries[i];
      const label = `${entry.date} | ${entry.task} | ${entry.hours}h`;
      // For all but the last entry, keep the Quick Create panel open via
      // "Save & Create New" so the next task can be typed straight in.
      const keepPanelOpen = i < stage1Entries.length - 1;
      const t0 = Date.now();
      try {
        const status = await withRetry(() => client.fillEntryWithTaskLookup(entry, { keepPanelOpen }));
        const elapsed = ((Date.now() - t0) / 1000).toFixed(1);
        console.log(`  ✓ [${i + 1}/${stage1Entries.length}] ${label} [${status === 'created' ? 'task created' : 'task found'}] (${elapsed}s)`);
        if (status === 'created') results.created++; else results.existing++;
      } catch (err) {
        const elapsed = ((Date.now() - t0) / 1000).toFixed(1);
        console.error(`  ✗ [${i + 1}/${stage1Entries.length}] ${label} — ${(err as Error).message} (${elapsed}s)`);
        failures.push({ entry: label, error: (err as Error).message });
        results.failed++;
      }
    }
  } finally {
    await client.close();
    await logger.close();
  }

  // ── Summary ─────────────────────────────────────────────────────────────
  console.log('\n─────────────────────────────────────');
  console.log('Summary:');
  console.log(`  Run ID:                 ${runId}`);
  console.log(`  Logs/trace:             ${logger.runDir}`);
  console.log(`  Stage 1 — task found:   ${results.existing}`);
  console.log(`  Stage 1 — task created: ${results.created}`);
  if (failures.length > 0) {
    console.log(`  Failed:                 ${results.failed}`);
    console.log('\nFailed entries:');
    failures.forEach((f) => console.log(`  • ${f.entry}\n    ${f.error}`));
    process.exit(1);
  } else {
    console.log('\nAll done! ✓');
  }
}

/**
 * Заливка табеля через Web API. Один проход по ВСЕМ записям Excel — деления на
 * Stage 1 / Stage 2 здесь нет, потому что у записи есть поле даты и не нужно
 * ни открывать Quick Create, ни листать недели в гриде.
 *
 * На каждую ячейку Excel создаются ДВЕ записи D365: Payable и Billable. В этом
 * тенанте это отдельные строки (amc_hourstype), а не два поля одной записи —
 * см. src/d365-api.ts.
 */
async function runApiMode(
  client: D365Client,
  entries: TimeEntry[],
  runId: string,
  runDir: string,
): Promise<boolean> {
  const api = new D365Api(client.getCurrentPage(), (m) => console.log(m));

  console.log(`\n[4/4] Web API: creating ${entries.length * 2} record(s) for ${entries.length} cell(s) ──`);
  const ctx = await api.resolveContext();
  console.log(`      Project: ${ctx.projectName}`);
  console.log(`      Team:    ${ctx.teamName}`);

  const isoDates = entries.map((e) => toIsoDate(e.date)).sort();
  const from = isoDates[0];
  const to = isoDates[isoDates.length - 1];
  console.log(`      Period:  ${from} … ${to}`);

  const taskIndex = await api.loadTaskIndex(ctx.projectId);
  console.log(`      Project tasks known to D365: ${taskIndex.size}`);

  // Ключи уже существующих записей: повторный запуск не задваивает табель.
  const existing = await api.existingEntryKeys(ctx.userId, from, to);
  console.log(`      Entries already in this period: ${existing.size}\n`);

  const stats = { created: 0, skipped: 0, tasksCreated: 0, failed: 0 };
  const failures: string[] = [];

  for (let i = 0; i < entries.length; i++) {
    const entry = entries[i];
    const label = `${entry.date} | ${entry.task} | ${entry.hours}h`;
    const date = toIsoDate(entry.date);
    const nameKey = entry.task.trim().toLowerCase();

    let taskId = taskIndex.get(nameKey);
    if (!taskId) {
      try {
        taskId = await api.createTask(entry.task.trim(), ctx.projectId);
        taskIndex.set(nameKey, taskId);
        stats.tasksCreated++;
        console.log(`  + project task created: ${entry.task}`);
      } catch (err) {
        stats.failed += 2;
        failures.push(`${label} — не смог создать задачу: ${(err as Error).message}`);
        console.error(`  ✗ [${i + 1}/${entries.length}] ${label} — task create failed`);
        continue;
      }
    }

    const done: string[] = [];
    for (const hoursType of [HOURS_TYPE_PAYABLE, HOURS_TYPE_BILLABLE]) {
      const kind = hoursType === HOURS_TYPE_PAYABLE ? 'payable' : 'billable';
      const key = `${date}|${taskId}|${hoursType}`;
      if (existing.has(key)) {
        stats.skipped++;
        done.push(`${kind}: already there`);
        continue;
      }
      try {
        await api.createEntry({ date, minutes: hoursToMinutes(entry.hours), taskId, hoursType, ctx });
        existing.add(key);
        stats.created++;
        done.push(kind);
      } catch (err) {
        stats.failed++;
        failures.push(`${label} (${kind}) — ${(err as Error).message}`);
        done.push(`${kind}: FAILED`);
      }
    }
    const mark = done.some((d) => d.includes('FAILED')) ? '✗' : '✓';
    console.log(`  ${mark} [${i + 1}/${entries.length}] ${label}  [${done.join(', ')}]`);
  }

  console.log('\n─────────────────────────────────────');
  console.log('Summary (Web API mode):');
  console.log(`  Run ID:                 ${runId}`);
  console.log(`  Logs:                   ${runDir}`);
  console.log(`  Records created:        ${stats.created}`);
  console.log(`  Already existed:        ${stats.skipped}`);
  console.log(`  Project tasks created:  ${stats.tasksCreated}`);
  if (failures.length > 0) {
    console.log(`  Failed:                 ${stats.failed}`);
    console.log('\nFailures:');
    failures.forEach((f) => console.log(`  • ${f}`));
    return false;
  }
  console.log('\nAll done! ✓  Записи созданы в статусе Draft — проверьте и нажмите Submit в D365.');
  return true;
}

// Print a compact, human-readable account of what the Excel parser saw. Makes
// a "0 entries" result diagnosable at a glance: which weeks exist, how many
// task rows / hours each has, and which one the --week filter selected.
function printParseReport(diag: ParseDiagnostics): void {
  console.log(`      Sheet "${diag.sheetName}": ${diag.totalRows} rows, ${diag.weekBlocks} week block(s) detected.`);
  if (diag.weekBlocks === 0) return;
  console.log('      Weeks in file:');
  diag.weeks.forEach((w) => {
    const mark = diag.filterApplied ? (w.matchedFilter ? '► ' : '  ') : '  ';
    console.log(
      `      ${mark}${w.weekStart}–${w.weekEnd} (Mon ${w.monday}): ` +
        `${w.taskRows} task row(s), ${w.taskRowsWithHours} with hours, ${w.entries} entr${w.entries === 1 ? 'y' : 'ies'}` +
        `${diag.filterApplied && w.matchedFilter ? '  ← selected by filter' : ''}`,
    );
  });
}

async function withRetry<T>(fn: () => Promise<T>, attempts = 2): Promise<T> {
  let lastError: Error | undefined;
  for (let i = 0; i < attempts; i++) {
    try {
      return await fn();
    } catch (err) {
      lastError = err as Error;
      if (i < attempts - 1) {
        console.warn(`  ⚠ Retrying after error: ${lastError.message}`);
        await new Promise((r) => setTimeout(r, 1500));
      }
    }
  }
  throw lastError;
}

main()
  .then(() => {
    // Explicit exit: Playwright browser handles and (before login) the stdin
    // listener can keep the event loop alive, so a bare `return` from main()
    // leaves the CLI hanging after "All done!". Failures below call
    // process.exit(1) directly before this resolves.
    process.exit(0);
  })
  .catch((err) => {
    console.error('Unexpected error:', err);
    process.exit(1);
  });
