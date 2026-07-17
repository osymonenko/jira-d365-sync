import * as path from 'path';
import { test, expect } from './fixtures';
import { D365Client } from '../src/d365';
import { TimeEntry } from '../src/types';
import { firstEntriesPerTask, parseTimesheet } from '../src/excel';

const D365_URL = process.env.D365_URL ?? 'https://amcbridge.crm.dynamics.com/';
const TIMESHEET_FILE = process.env.D365_TIMESHEET_FILE ?? path.resolve('data/timesheet.xlsx');

// The fill-flow test drives the real timesheet — picks the first parsed task on its earliest day.
// "check sum" rows are skipped by parseTimesheet itself (see src/excel.ts).
function loadFirstStage1Entry(): TimeEntry {
  const stage1 = firstEntriesPerTask(parseTimesheet(TIMESHEET_FILE));
  if (stage1.length === 0) {
    throw new Error(`No time entries parsed from ${TIMESHEET_FILE} — check the Excel structure`);
  }
  return stage1[0];
}

const SAMPLE_ENTRY: TimeEntry = loadFirstStage1Entry();

/**
 * Discrete steps of the D365 time-entry flow as individual tests so the
 * VS Code Playwright Test Explorer can run them in isolation and the user
 * can drop breakpoints inside D365Client methods.
 *
 * Tests run serially in a single Edge session (worker-scoped fixture);
 * SSO cookies persist in ./edge-profile across runs.
 */
test.describe.serial('D365 time-entry flow', () => {
  test('signs into D365 and reaches the nav bar', async ({ edgePage }) => {
    await edgePage.goto(D365_URL, { waitUntil: 'domcontentloaded' });
    await edgePage
      .locator('[aria-label="Main Navigation"], nav[aria-label*="navigation" i], [data-id="navbar-container"]')
      .first()
      .waitFor({ state: 'visible', timeout: 5 * 60 * 1000 });
    expect(edgePage.url()).toContain('dynamics.com');
  });

  test('navigates to the Time Entries list', async ({ edgePage }) => {
    const client = D365Client.fromPage(edgePage, D365_URL);
    await client.navigateToTimeEntries();
    expect(edgePage.url()).toContain('msdyn_timeentry');
  });

  test('opens the Quick Create: Time Entry dialog', async ({ edgePage }) => {
    const client = D365Client.fromPage(edgePage, D365_URL);
    await client.openNewTimeEntry();
    await expect(
      edgePage.locator('div[aria-label="Quick Create: Time Entry"], .ms-Dialog-main, [data-id="quickCreateFlyout"]'),
    ).toBeVisible();
  });

  test('fills a single entry via the task-lookup flow', async ({ edgePage }, testInfo) => {
    const client = D365Client.fromPage(edgePage, D365_URL);
    const result = await client.fillEntryWithTaskLookup(SAMPLE_ENTRY);

    // Both branches are valid outcomes:
    //   'existing' — задача уже была в D365, выбрана из lookup и сохранена
    //   'created'  — задача отсутствовала, пройден sub-panel "Imported Project Task"
    expect(['existing', 'created']).toContain(result);

    testInfo.annotations.push({
      type: 'flow-branch',
      description: `${result} (task="${SAMPLE_ENTRY.task}", date=${SAMPLE_ENTRY.date}, hours=${SAMPLE_ENTRY.hours})`,
    });
    console.log(`    → flow branch: ${result} for "${SAMPLE_ENTRY.task}"`);

    // Quick Create панель должна быть закрыта после save — общий пост-кондишн для обеих веток.
    await expect(edgePage.locator('div[aria-label="Quick Create: Time Entry"]')).toBeHidden();
  });
});
