import * as dotenv from 'dotenv';
dotenv.config();

import { Command } from 'commander';
import { firstEntriesPerTask, parseTimesheet } from './excel';
import { D365Client } from './d365';
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
  .option('--preflight-only', 'Run preflight checks only (no browser automation)')
  .option('--list-tasks', 'Parse Excel and list all tasks without running automation')
  .parse(process.argv);

const opts = program.opts<{
  file: string;
  week?: string;
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
    try {
      entries = parseTimesheet(opts.file, weekFilter);
    } catch (err) {
      console.error('Failed to parse Excel file:', (err as Error).message);
      process.exit(1);
    }
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

  if (!d365Url) {
    console.error('Error: D365_URL is not set. Add it to .env or set the environment variable.');
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

  // Parse Excel
  console.log(`\nReading timesheet: ${opts.file}`);
  let entries: TimeEntry[];
  try {
    entries = parseTimesheet(opts.file, weekFilter);
  } catch (err) {
    console.error('Failed to parse Excel file:', (err as Error).message);
    process.exit(1);
  }

  if (entries.length === 0) {
    console.log('No time entries found (all cells are empty or zero). Nothing to submit.');
    return;
  }

  const stage1Entries = firstEntriesPerTask(entries);
  console.log(`\nFound ${entries.length} time entries across ${stage1Entries.length} unique task(s):`);
  stage1Entries.forEach((e) => console.log(`  • ${e.task}`));
  // TODO: Stage 2 (оставшиеся дни многодневных задач) будет реализован отдельно
  // через клик по ячейкам в "All Weekly Time Entries" weekly grid — после
  // Stage 1 строки задач уже существуют в гриде, нужно только проставить часы
  // в ячейках пересечения день × задача.

  // Logger writes to logs/<runId>/run.log alongside Playwright trace + error screenshots.
  const runId = createRunId();
  const logger: Logger = createLogger(runId);
  logger.log(`Run ${runId} started — logs in ${logger.runDir}`);

  // Launch browser
  const client = new D365Client(d365Url, userDataDir, browserMode, cdpUrl, logger);
  await client.launch();

  const results = { created: 0, existing: 0, failed: 0 };
  const failures: { entry: string; error: string }[] = [];

  try {
    // ── Stage 1: First-day entry per task (with inline task creation) ────────
    console.log(`\n── Stage 1: Submitting first-day entry for ${stage1Entries.length} task(s) ──`);
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
