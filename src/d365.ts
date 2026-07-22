import { chromium, Browser, BrowserContext, Page, Locator } from 'playwright';
import * as path from 'path';
import { TimeEntry } from './types';
import { Logger } from './logger';

async function selectAll(locator: Locator): Promise<void> {
  await locator.click({ clickCount: 3 });
}

const SELECTORS = {
  // The Weekly Time Entries view embeds its command bar inside a PCF iframe.
  // Stable anchors (confirmed via MCP DOM dump 2026-05-20):
  //   1. src path: …/webresources/msdyn_/TimeEntryGridControl/html/index.html
  //   2. aria-label: "Time Entry Grid"  (localizable — fragile in non-EN tenants)
  //   3. parent PCF container: div[data-id="entity_control_container"]
  // The "+ New" button inside is a native <button aria-label="New"> with
  // textContent " New" (Fluent UI icon glyph + word) — match on
  // aria-label, NOT on textContent.
  weeklyGridIframe: 'iframe[src*="TimeEntryGridControl"], iframe[aria-label="Time Entry Grid"], div[data-id="entity_control_container"] iframe',
  newTimeEntryBtn: 'button[aria-label="New Time Entry"], button[title="New Time Entry"], button:has-text("New Time Entry"), [aria-label="New Time Entry"], button:has-text("New")',
  quickCreatePanel: 'div[aria-label="Quick Create: Time Entry"], .ms-Dialog-main, [data-id="quickCreateFlyout"]',
  dateField: 'input[aria-label="Date"], [data-id*="date"] input, input[placeholder*="date" i]',
  payableDurationContainer: '[data-id*="duration" i][data-id*="payable" i], [aria-label*="Payable Duration" i]',
  // "Copy to Billable Duration" two-option field in the Quick Create panel.
  // D365 renders a two-option/boolean field with the "Toggle" display as a
  // Fluent UI switch (role="switch", aria-checked reflects on/off) — that's the
  // pill control visible in the screenshot ("No"). Some tenants/builds render
  // it as a checkbox instead, so we OR both role variants plus a data-id
  // fallback (logical name usually contains "billable"). NOTE: this selector is
  // written defensively from the field label — not yet confirmed against a live
  // trace like the others. If it misses, widen the OR list or run playwright-healer.
  copyToBillableToggle:
    '[role="switch"][aria-label*="Copy to Billable" i], ' +
    '[role="checkbox"][aria-label*="Copy to Billable" i], ' +
    'input[type="checkbox"][aria-label*="Copy to Billable" i], ' +
    '[data-id*="billable" i][role="switch"], ' +
    '[data-id*="billable" i] [role="switch"], ' +
    '[aria-label*="Copy to Billable Duration" i]',
  projectTaskField: '[data-id*="projecttask" i] input, [aria-label*="Project Task" i] input',
  // Task Lookup results dropdown — контейнер всегда имеет aria-label="Lookup
  // results" (верифицировано через trace.zip 2026-05-22), и для "No records",
  // и для случая когда есть совпадения. Видимая надпись "Imported Project Tasks"
  // — это просто заголовок-секция, не aria-label.
  // Скопим все under-lookup селекторы внутрь этого региона, чтобы не цеплять
  // случайные [role="treeitem"] / "New"-кнопки на остальной странице.
  taskLookupRegion: '[aria-label*="Lookup results" i]',
  // Стабильный data-id для "No records found" — есть и span, и role="status",
  // подходят обе для проверки факта "нет записей".
  noRecordsText: '[data-id$="_No_Records_Text"], [aria-label*="No records" i]',
  newTaskBtn: 'button:has-text("New"), a:has-text("Create a new record"), button:has-text("Create a new record")',
  importedTaskPanel: '[aria-label*="Imported Project Task" i], [aria-label*="Quick Create: Imported" i]',
  // Name field in the "Imported Project Task" sub-panel.
  // D365 renders it as a textbox-role element with aria-label="Name" (NOT a
  // nested <input>) — verified via accessibility snapshot 2026-05-22:
  //   textbox "Name" [ref=e345] (placeholder "---")
  // The previous selector `[aria-label="Name"] input` looked for an <input>
  // *descendant* of [aria-label="Name"] — which never matched because the
  // textbox itself carries the aria-label. Keep the legacy descendant pattern
  // as a fallback in case the DOM regresses to a wrapped-input layout.
  taskNameField:
    '[data-id*="taskname" i] input, [data-id*="subject" i] input, ' +
    'input[aria-label="Name"], textarea[aria-label="Name"], [role="textbox"][aria-label="Name"], ' +
    '[aria-label="Name"] input, [placeholder="Name"]',
  saveAndCloseBtn: 'button:has-text("Save and Close"), button:has-text("Save & Close")',
  // Split-button chevron next to "Save and Close" (Fluent UI). The chevron is a
  // separate <button aria-haspopup="true"> rendered as a sibling — its
  // aria-label usually includes "Save and Close" plus "more options" wording,
  // but exact wording varies by D365 build, so we OR several patterns.
  saveSplitButtonChevron:
    'button[aria-haspopup="true"][aria-label*="Save and Close" i], ' +
    'button[aria-haspopup="menu"][aria-label*="Save and Close" i], ' +
    'button[aria-haspopup="true"][aria-label*="more options" i], ' +
    'button[aria-haspopup="menu"][aria-label*="more options" i]',
  // "Save & Create New" menu item — Fluent UI uses role="menuitem" by default,
  // but some D365 versions render plain <button>; cover both.
  saveAndCreateNewMenuItem:
    '[role="menuitem"]:has-text("Save & Create New"), ' +
    '[role="menuitemcheckbox"]:has-text("Save & Create New"), ' +
    'button:has-text("Save & Create New")',
  // Fluent UI DatePicker popup that opens when typing into / clicking the Date
  // field. Verified via trace.zip 2026-05-22:
  //   <div role="dialog" aria-modal="true" aria-label="Calendar"
  //        id="datePicker-popupSurface1"
  //        class="fui-DatePicker__popupSurface ...">
  // The popup is `aria-modal="true"`, so Tab is trapped *inside* the calendar
  // — that's why `keyboard.press('Tab')` after `dateField.fill()` did NOT
  // dismiss it. Subsequent clicks on Payable Duration / Type / Project Task
  // hit the calendar's modal overlay and time out with
  //   "<div role='dialog' aria-label='Calendar'> subtree intercepts pointer events".
  // OR multiple anchors for resilience: D365 occasionally renames the dialog
  // (e.g. "Date Picker") between releases.
  datePickerPopup:
    '[role="dialog"][aria-label="Calendar"], ' +
    '[role="dialog"][aria-label*="Date Picker" i], ' +
    '.fui-DatePicker__popupSurface',
  // D365's native navigation-away guard ("Unsaved changes — Do you want to save
  // your changes before leaving this page?"). Confirmed via error screenshots
  // 2026-07-18: pressing Escape on the "Imported Project Task" sub-panel (when it
  // hasn't visibly closed within 20s) does NOT dismiss the panel — it instead
  // triggers this confirmation dialog, which is a genuine second modal
  // (`div[id^="modalDialogRoot_"]`) that intercepts pointer events on the entire
  // page indefinitely until explicitly answered.
  unsavedChangesDialog: '[role="dialog"]:has-text("Unsaved changes"), div[id^="modalDialogRoot_"]:has-text("Unsaved changes")',
  unsavedChangesSaveBtn: 'button:has-text("Save and continue")',
  unsavedChangesDiscardBtn: 'button:has-text("Discard changes")',
};

// Map decimal hours to the exact label used in the Payable Duration dropdown.
// Options are 15-minute increments up to 8 hours max:
//   <1h: "15 minutes", "30 minutes", "45 minutes"
//   1h : "1 hour"  (singular)
//   >1h: "1.5 hours", "2 hours", "1.25 hours" ...  (decimal, plural)
function hoursToD365DurationOption(hours: number): string {
  if (hours < 1) {
    const minutes = Math.round(hours * 60);
    return `${minutes} minutes`;
  }
  if (hours === 1) return '1 hour';
  const formatted = parseFloat(hours.toFixed(2)).toString();
  return `${formatted} hours`;
}

function defaultLog(msg: string): void {
  const ts = new Date().toTimeString().slice(0, 8);
  console.log(`[${ts}] ${msg}`);
}

// Wait for a line equal to `signal` on stdin (used for the GUI "CONTINUE" ack).
// Returns a `cancel` alongside the promise: attaching a 'data' listener puts
// stdin into flowing mode and keeps the event loop alive, so the caller MUST
// cancel() once it no longer needs the signal — otherwise the process never
// exits after the run finishes.
function waitForStdinSignal(signal: string): { promise: Promise<void>; cancel: () => void } {
  let onData: (chunk: Buffer) => void = () => {};
  const promise = new Promise<void>((resolve) => {
    onData = (chunk: Buffer): void => {
      if (chunk.toString().split(/\r?\n/).some((line) => line.trim() === signal)) {
        resolve();
      }
    };
    process.stdin.on('data', onData);
  });
  const cancel = (): void => {
    process.stdin.off('data', onData);
  };
  return { promise, cancel };
}

export class D365Client {
  private browser: Browser | null = null;
  private context: BrowserContext | null = null;
  private page: Page | null = null;
  private readonly d365Url: string;
  private readonly userDataDir: string;
  private readonly browserMode: string;
  private readonly cdpUrl: string;
  private readonly logger: Logger | null;
  // When true, each new Time Entry has its "Copy to Billable Duration" toggle
  // set to Yes during the pass. When false (default) the toggle is left at the
  // D365 default and never touched. Driven by COPY_TO_BILLABLE_DURATION in .env.
  private readonly copyToBillable: boolean;
  private tracingActive = false;

  constructor(
    d365Url: string,
    userDataDir: string,
    browserMode = 'chrome-profile',
    cdpUrl = 'http://localhost:9222',
    logger: Logger | null = null,
    copyToBillable = false,
  ) {
    this.d365Url = d365Url;
    this.userDataDir = path.resolve(userDataDir);
    this.browserMode = browserMode;
    this.cdpUrl = cdpUrl;
    this.logger = logger;
    this.copyToBillable = copyToBillable;
  }

  private log(msg: string): void {
    if (this.logger) this.logger.log(msg);
    else defaultLog(msg);
  }

  private async captureFailure(stepLabel: string, err: Error): Promise<void> {
    if (!this.logger || !this.page) return;
    const safeLabel = stepLabel.replace(/[^a-z0-9]+/gi, '_').slice(0, 60);
    const file = path.join(this.logger.runDir, `error-${Date.now()}-${safeLabel}.png`);
    try {
      await this.page.screenshot({ path: file, fullPage: true });
      this.logger.error(`Screenshot saved: ${file}`);
    } catch (screenshotErr) {
      this.logger.error(`Screenshot failed: ${(screenshotErr as Error).message}`);
    }
    this.logger.error(`Step "${stepLabel}" failed: ${err.message}`);
  }

  private async startTracing(): Promise<void> {
    if (!this.context || this.tracingActive) return;
    try {
      await this.context.tracing.start({ screenshots: true, snapshots: true, sources: false });
      this.tracingActive = true;
      this.log('Playwright tracing started');
    } catch (err) {
      // CDP-attached contexts don't always support tracing — degrade gracefully
      this.log(`Tracing not started: ${(err as Error).message}`);
    }
  }

  private async stopTracing(): Promise<void> {
    if (!this.context || !this.tracingActive) return;
    try {
      const tracePath = this.logger
        ? path.join(this.logger.runDir, 'trace.zip')
        : path.resolve('logs', `trace-${Date.now()}.zip`);
      await this.context.tracing.stop({ path: tracePath });
      this.log(`Trace saved: ${tracePath}`);
    } catch (err) {
      this.log(`Trace stop failed: ${(err as Error).message}`);
    } finally {
      this.tracingActive = false;
    }
  }

  async launch(): Promise<void> {
    if (this.browserMode === 'cdp') {
      this.log(`Connecting to Chrome via CDP at ${this.cdpUrl} ...`);
      this.browser = await chromium.connectOverCDP(this.cdpUrl);
      this.context = this.browser.contexts()[0];
      if (!this.context) throw new Error('No browser context found. Make sure Chrome is open.');

      const pages = this.context.pages();
      this.log(`Found ${pages.length} open tab(s) in Chrome`);

      const d365Host = new URL(this.d365Url).hostname;
      const d365Page = pages.find(p => p.url().includes(d365Host));
      const anyRealPage = pages.find(
        p => !p.url().startsWith('chrome://') && !p.url().startsWith('chrome-extension://'),
      );
      this.page = d365Page ?? anyRealPage ?? await this.context.newPage();
      this.log(`Using tab: ${this.page.url() || '(new tab)'}`);
      await this.page.bringToFront();

      if (!this.page.url().includes(d365Host)) {
        this.log(`Navigating to D365: ${this.d365Url}`);
        await this.page.goto(this.d365Url, { waitUntil: 'domcontentloaded', timeout: 30000 });
      } else {
        this.log(`Already on D365 — skipping navigation`);
      }

    } else if (this.browserMode === 'chrome-profile') {
      this.log(`Launching Edge with profile: ${this.userDataDir}`);
      this.context = await chromium.launchPersistentContext(this.userDataDir, {
        channel: 'msedge',
        headless: false,
        args: ['--start-maximized'],
        ignoreDefaultArgs: ['--enable-automation'],
      });
      const pages = this.context.pages();
      this.page = pages.length > 0 ? pages[0] : await this.context.newPage();
      this.log(`Navigating to D365: ${this.d365Url}`);
      await this.page.goto(this.d365Url, { waitUntil: 'domcontentloaded', timeout: 30000 });

    } else {
      this.log(`Launching bundled Chromium with profile: ${this.userDataDir}`);
      this.context = await chromium.launchPersistentContext(this.userDataDir, {
        headless: false,
        args: ['--start-maximized'],
      });
      const pages = this.context.pages();
      this.page = pages.length > 0 ? pages[0] : await this.context.newPage();
      this.log(`Navigating to D365: ${this.d365Url}`);
      await this.page.goto(this.d365Url, { waitUntil: 'domcontentloaded', timeout: 30000 });
    }

    await this.startTracing();

    this.log(`Waiting for D365 to load... Sign in if prompted, then click "I'm logged in" (auto-continues when nav bar appears, max 5 min)`);
    console.log('__AWAIT_LOGIN__');
    // Two ways to proceed: auto-detect the nav bar, or the user manually acks via
    // the GUI ("CONTINUE" on stdin). Whichever wins the race, the loser must not
    // crash the process: if the user acks first, navBarWait would otherwise reject
    // with a timeout ~5 min later (unhandled rejection → process death *mid-run*),
    // so we swallow it. And the stdin listener is cancelled once we're done, or it
    // keeps stdin flowing and the process never exits after the run.
    const stdinSignal = waitForStdinSignal('CONTINUE');
    const navBarWait = this.page!.locator(
      '[aria-label="Main Navigation"], nav[aria-label*="navigation" i], [data-id="navbar-container"]',
    ).first().waitFor({ state: 'visible', timeout: 5 * 60 * 1000 })
      .then(() => this.log('D365 nav bar detected'))
      .catch(() => {});
    try {
      await Promise.race([
        navBarWait,
        stdinSignal.promise.then(() => this.log('User confirmed login (manual signal)')),
      ]);
    } finally {
      stdinSignal.cancel();
    }
    this.log(`D365 loaded — proceeding`);
  }

  async navigateToTimeEntries(): Promise<void> {
    const page = this.getPage();
    this.log(`Navigating to Time Entries list...`);
    await page.goto(
      `${this.d365Url}main.aspx?pagetype=entitylist&etn=msdyn_timeentry`,
      { waitUntil: 'domcontentloaded', timeout: 30000 },
    );

    // D365 may redirect to the app-selection hub if no appid context is set.
    // Detect by URL and click "Team Member's PO Hub" automatically.
    if (page.url().includes('pagetype=apps')) {
      this.log(`App selection page detected — clicking "Team Member's PO Hub"...`);
      const appTile = page.locator(
        '[aria-label*="Team Member" i], h2:has-text("Team Member"), .appTile:has-text("Team Member")',
      ).first();
      await appTile.waitFor({ state: 'visible', timeout: 15000 });
      await appTile.click();
      this.log(`Clicked app tile — waiting for app to load...`);
      await page.waitForTimeout(3000);
      await page.goto(
        `${this.d365Url}main.aspx?pagetype=entitylist&etn=msdyn_timeentry`,
        { waitUntil: 'domcontentloaded', timeout: 30000 },
      );
    }

    // No explicit wait here — D365 networkidle never settles (constant
    // polling) and the "New Time Entry" text only appears in a hover tooltip.
    // openNewTimeEntry() already waits for the real iframe button (≤20s).
    this.log(`Time Entries page loaded`);
  }

  async fillEntryWithTaskLookup(
    entry: TimeEntry,
    options: { keepPanelOpen?: boolean } = {},
  ): Promise<'created' | 'existing'> {
    try {
      return await this.fillEntryWithTaskLookupInner(entry, options);
    } catch (err) {
      await this.captureFailure(`task-lookup ${entry.task}`, err as Error);
      throw err;
    }
  }

  private async fillEntryWithTaskLookupInner(
    entry: TimeEntry,
    options: { keepPanelOpen?: boolean },
  ): Promise<'created' | 'existing'> {
    const page = this.getPage();
    const label = `${entry.date} | ${entry.task} | ${entry.hours}h`;
    await this.dismissMicrosoftSurvey();

    // 1. Открыть Quick Create dialog — или переиспользовать, если он уже открыт
    //    после предыдущего "Save & Create New" из bulk-цикла.
    const panelAlreadyOpen = await page
      .locator(SELECTORS.quickCreatePanel)
      .first()
      .isVisible()
      .catch(() => false);
    if (panelAlreadyOpen) {
      this.log(`  Quick Create panel already open — reusing existing dialog`);
    } else {
      await this.openNewTimeEntry();
    }

    // 2. Дата
    this.log(`    Filling date: ${entry.date}`);
    const dateField = page.locator(SELECTORS.dateField).first();
    await dateField.waitFor({ state: 'visible', timeout: 10000 });
    await selectAll(dateField);
    await dateField.fill(entry.date);
    await page.keyboard.press('Tab');
    // D365 renders the Date field as a Fluent UI DatePicker. Typing into it (or
    // even clicking it after a "Save & Create New" round-trip) can leave the
    // calendar popup open: <div role="dialog" aria-modal="true"
    // aria-label="Calendar" class="fui-DatePicker__popupSurface ...">. Because
    // it's aria-modal, the Tab above is trapped *inside* the calendar and
    // doesn't dismiss it, so the popup's overlay intercepts subsequent clicks
    // on Type / Payable Duration / Project Task. Escape triggers the popup's
    // onDismiss handler (verified for Fluent UI DatePicker) and works even
    // while focus is trapped. We treat dismissal as best-effort: if the popup
    // never appeared (legacy non-Fluent build) the isVisible probe returns
    // false within ~200ms and we move on.
    await this.dismissDatePickerPopup();

    // 2.5. Copy to Billable Duration — only when explicitly enabled via settings.
    if (this.copyToBillable) {
      await this.setCopyToBillableToggle(true);
    }

    // 3. Payable Duration: открыть dropdown, выбрать опцию по label'у ("30 minutes" / "1 hour" / ...)
    this.log(`    Filling duration: ${entry.hours}h`);
    const durationContainer = page.locator(SELECTORS.payableDurationContainer).first();
    await durationContainer.waitFor({ state: 'visible', timeout: 10000 });
    await durationContainer.click();
    await page.waitForTimeout(500);

    const optionLabel = hoursToD365DurationOption(entry.hours);
    this.log(`    Selecting duration option: "${optionLabel}"`);
    const option = page.getByRole('option', { name: optionLabel, exact: true }).first();
    await option.waitFor({ state: 'visible', timeout: 5000 });
    await option.click();
    await page.waitForTimeout(300);

    // 4. Поиск задачи в лукапе
    this.log(`    Searching for task: "${entry.task}"`);
    // On retry after a failed "Add" click the previous chip may still be set in
    // the Project Task field, hiding the text input. Clear it first if present.
    // The chip's remove button has aria-label containing the task name followed
    // by something like " (remove)" or is a generic button inside the chip
    // container. We try the common D365 patterns: [aria-label="<task> (remove)"]
    // and the generic chip-dismiss button scoped to the projecttask container.
    const chipRemoveBtn = page.locator(
      `[data-id*="projecttask" i] button[aria-label*="${entry.task.replace(/"/g, '\\"')}" i], ` +
      '[data-id*="projecttask" i] button[aria-label*="remove" i], ' +
      '[data-id*="projecttask" i] button[aria-label*="delete" i], ' +
      '[data-id*="projecttask" i] button[aria-label="×"], ' +
      '[data-id*="projecttask" i] [role="option"] button',
    ).first();
    const chipVisible = await chipRemoveBtn.isVisible({ timeout: 500 }).catch(() => false);
    if (chipVisible) {
      this.log(`    Clearing existing Project Task chip before re-filling...`);
      await chipRemoveBtn.click();
      await page.waitForTimeout(300);
    }
    const taskInput = page.locator(SELECTORS.projectTaskField).first();
    await taskInput.waitFor({ state: 'visible', timeout: 10000 });
    await taskInput.click();
    await taskInput.fill(entry.task);

    // Дожидаемся стабильного состояния lookup-результатов. Два сигнала:
    //   1. "No records found" — контейнер с aria-label="Lookup results"
    //      (см. SELECTORS.taskLookupRegion + SELECTORS.noRecordsText).
    //   2. Подсказка существующей задачи — <li role="treeitem"> с aria-label
    //      вида "<Task Name>, <Project> - <Milestone>", внутри role="tree"
    //      (верифицировано через trace.zip 2026-05-22).
    // Подсказку фильтруем по тексту задачи case-insensitive: D365 может вернуть
    // запись в другом регистре чем в Excel (напр. "Internal Daily meeting" vs
    // "Internal daily meeting"). Совпадение по aria-label избегает зависимости
    // от inner-text и от регистра. Скоупим [role="treeitem"] внутрь lookupRegion,
    // чтобы не цеплять стрейловые treeitem других полей Quick Create.
    const lookupRegion = page.locator(SELECTORS.taskLookupRegion).first();
    const noRecordsLoc = lookupRegion.locator(SELECTORS.noRecordsText).first();
    const escapedTask = entry.task.replace(/\\/g, '\\\\').replace(/"/g, '\\"');
    const suggestionLoc = lookupRegion.locator(
      `[role="treeitem"][aria-label*="${escapedTask}" i]`,
    ).first();
    this.log(`    Waiting for lookup dropdown to settle (No records OR suggestion)...`);
    await noRecordsLoc
      .or(suggestionLoc)
      .first()
      .waitFor({ state: 'visible', timeout: 20000 });
    const noRecords = await noRecordsLoc.isVisible();
    this.log(`    Lookup settled — noRecords=${noRecords}`);

    if (!noRecords) {
      // 5а. Задача найдена → выбрать из подсказок и сохранить.
      // Тот же селектор что и для wait — гарантирует что кликаем ровно тот же
      // элемент, появление которого мы дождались.
      this.log(`    ✓ Task found — selecting: "${entry.task}"`);
      const hasSuggestion = await suggestionLoc.isVisible({ timeout: 3000 }).catch(() => false);
      if (hasSuggestion) {
        await suggestionLoc.click();
      } else {
        await page.keyboard.press('Enter');
      }
      // D365 lookup dialogs (full record-picker, not just inline dropdown) require
      // an explicit "Add" button click after selecting a row from the results table.
      // Scope to the Quick Create panel so we don't pick up unrelated "Add" buttons
      // elsewhere on the page (e.g. command bar). Remove `:visible` from has-text
      // selectors — it is not a standard CSS pseudo-class and causes false positives
      // in Playwright's CSS engine (matches hidden elements transiently).
      const quickCreatePanel = page.locator(SELECTORS.quickCreatePanel).first();
      const addBtn = quickCreatePanel.locator(
        'button[aria-label="Add"], button[data-id*="add" i]',
      ).first();
      const addVisible = await addBtn.isVisible({ timeout: 500 }).catch(() => false);
      if (addVisible) {
        this.log(`    Clicking "Add" to confirm lookup selection...`);
        await addBtn.waitFor({ state: 'visible', timeout: 3000 });
        await addBtn.click();
        await page.waitForTimeout(500);
      }
      await page.waitForTimeout(300);

      this.log(`    Saving entry: ${label}`);
      await this.saveTimeEntry({ keepPanelOpen: !!options.keepPanelOpen, prevTask: entry.task });
      return 'existing';
    }

    // 5б. Задачи нет → создать через под-панель "Imported Project Task"
    this.log(`    + Task not found — creating: "${entry.task}"`);
    // Скопим "+ New" к региону lookup-результатов, чтобы не кликнуть
    // command-bar "New" в основной панели.
    const newBtn = lookupRegion.locator(SELECTORS.newTaskBtn).first();
    await newBtn.waitFor({ state: 'visible', timeout: 5000 });
    await newBtn.click();

    this.log(`    Waiting for Imported Project Task panel...`);
    const importedPanel = page.locator(SELECTORS.importedTaskPanel).first();
    await importedPanel.waitFor({ state: 'visible', timeout: 10000 });

    // Scope the Name lookup INTO the Imported Project Task panel: the main
    // Quick Create dialog above can also expose generic "Name"/textbox nodes,
    // and an un-scoped first-match would pick the wrong one.
    const nameInput = importedPanel.locator(SELECTORS.taskNameField).first();
    await nameInput.waitFor({ state: 'visible', timeout: 5000 });
    await nameInput.click();
    await nameInput.fill(entry.task);
    this.log(`    Filled task name — saving sub-panel...`);

    // Scope Save & Close to the sub-panel — the main Quick Create dialog also
    // has its own Save & Close button and an un-scoped .first() could hit it.
    const saveSubPanel = importedPanel.locator(SELECTORS.saveAndCloseBtn).first();
    await saveSubPanel.waitFor({ state: 'visible', timeout: 5000 });
    await saveSubPanel.click();
    // Wait up to 20s for the sub-panel to close. If it doesn't close (D365 is slow
    // or showed a post-save state), press Escape to dismiss the top-most dialog
    // overlay — the Imported Task panel is always on top of Quick Create: Time Entry,
    // so Escape targets it specifically and leaves the main panel intact.
    const subPanelClosed = await importedPanel
      .waitFor({ state: 'hidden', timeout: 20000 })
      .then(() => true)
      .catch(() => false);
    if (!subPanelClosed) {
      this.log(`    ⚠ Sub-panel did not close after 20s — pressing Escape to dismiss overlay`);
      await page.keyboard.press('Escape');
      // Escape on the still-open sub-panel does not necessarily close it — it can
      // instead trigger D365's native "Unsaved changes" navigation-away confirmation
      // dialog (confirmed via error screenshots 2026-07-18). That dialog is a real
      // second modal that blocks every subsequent click page-wide until answered, so
      // we must explicitly detect and answer it rather than assume Escape worked.
      const unsavedDialog = page.locator(SELECTORS.unsavedChangesDialog).first();
      const unsavedDialogShown = await unsavedDialog
        .waitFor({ state: 'visible', timeout: 3000 })
        .then(() => true)
        .catch(() => false);
      if (unsavedDialogShown) {
        this.log(`    ⚠ "Unsaved changes" dialog appeared — clicking "Save and continue" to clear it`);
        const saveContinueBtn = unsavedDialog.locator(SELECTORS.unsavedChangesSaveBtn).first();
        await saveContinueBtn.click({ timeout: 5000 }).catch(() => {});
        await unsavedDialog.waitFor({ state: 'hidden', timeout: 10000 }).catch(() => {});
      }
      await importedPanel.waitFor({ state: 'hidden', timeout: 5000 }).catch(() => {});
    }
    // Extra wait for D365's DialogContainer overlay to fully detach from the DOM.
    // Without this, the section[id*="popupContainer"] element lingers and intercepts
    // pointer events on the underlying Quick Create: Time Entry panel.
    await page.waitForTimeout(1500);
    this.log(`    ✓ Task created: "${entry.task}" — saving time entry...`);
    await this.saveTimeEntry({ keepPanelOpen: !!options.keepPanelOpen, prevTask: entry.task });
    return 'created';
  }

  // Dismiss the Microsoft OBF feedback survey ("Microsoft would love your perspective").
  // The survey overlay is aria-hidden but still intercepts pointer events — clicking
  // Duration / Task fields fails with "obf-FloodgateDynamicUxContainer subtree
  // intercepts pointer events". Removing it from the DOM via evaluate() is the
  // simplest and most reliable dismissal — no close-button selector guessing needed.
  private async dismissMicrosoftSurvey(): Promise<void> {
    const page = this.getPage();
    const removed = (await page
      .evaluate(
        `(() => {
          const el = document.querySelector(
            '#obf-FloodgateDynamicUxContainer, [id*="obf-Floodgate"], [id*="obf-DxT"]'
          );
          if (!el) return false;
          el.remove();
          return true;
        })()`,
      )
      .catch(() => false)) as boolean;
    if (removed) {
      this.log(`    Dismissed Microsoft Feedback survey overlay`);
      await page.waitForTimeout(300);
    }
  }

  // Dismiss the "Your changes were saved" notification bar that D365 shows after
  // each save. It blocks clicks for several seconds if left open, so we proactively
  // close it. Non-fatal: if the selector doesn't match we just proceed.
  //
  // DOM structure verified via MCP browser_snapshot + evaluate 2026-06-19:
  //   D365 renders a Fluent UI BAR notification (type 2) into the main page DOM:
  //     div[data-id="notificationWrapper"]
  //       ul
  //         li#notificationWrapperglobal-notification-list
  //           … message content …
  //           button[aria-label="Close"][title="Close"]  ← × close button
  //   Type-1 notifications use button[aria-label="Dismiss notification"] instead.
  //   The old selectors ([data-id="notificationBar"], .ms-MessageBar-dismissal)
  //   never matched — the actual data-id is "notificationWrapper" and the close
  //   button is a bare fui-Button with aria-label="Close".
  private async dismissSavedNotification(): Promise<void> {
    const page = this.getPage();
    const closeBtn = page.locator(
      // Primary: BAR notification (type 2) — the × button inside the wrapper div.
      // data-id="notificationWrapper" is stable; aria-label="Close" is the Fluent UI
      // label D365 sets on the close × button (confirmed 2026-06-19).
      '[data-id="notificationWrapper"] button[aria-label="Close"], ' +
      // Type-1 (dialog) notification dismiss button — seen when addGlobalNotification
      // type:1 is used; aria-label differs from the BAR variant.
      'button[aria-label="Dismiss notification"], ' +
      // ID-pattern fallback: the LI wrapper always has id="notificationWrapperXxx"
      '[id*="notificationWrapper"] button[aria-label="Close"], ' +
      // Legacy fallbacks kept for resilience in case D365 reverts to older patterns.
      '[data-id="notificationBar"] button, ' +
      'button[aria-label="Dismiss this notification"], ' +
      '.ms-MessageBar-dismissal button',
    ).first();
    const visible = await closeBtn.isVisible({ timeout: 1500 }).catch(() => false);
    if (visible) {
      this.log(`    Dismissing "changes saved" notification...`);
      await closeBtn.click();
      await page.waitForTimeout(200);
    }
  }

  // Dismiss the Fluent UI DatePicker popup if it is currently open.
  // Why this exists: the popup is `aria-modal="true"`, so Tab is trapped
  // inside it — `keyboard.press('Tab')` after filling the date field does
  // NOT close it. Its overlay then intercepts every subsequent click on
  // sibling fields. We send `Escape` (the standard onDismiss trigger for
  // Fluent UI popups) and wait up to 2s for the dialog to detach. If the
  // popup wasn't open in the first place, the initial visibility probe
  // returns false within ~200ms and the method is a no-op — so it's cheap
  // to call defensively after every date interaction.
  private async dismissDatePickerPopup(): Promise<void> {
    const page = this.getPage();
    const popup = page.locator(SELECTORS.datePickerPopup).first();
    const isOpen = await popup.isVisible({ timeout: 200 }).catch(() => false);
    if (!isOpen) return;
    this.log(`    Dismissing Calendar popup (Escape)...`);
    await page.keyboard.press('Escape');
    await popup.waitFor({ state: 'hidden', timeout: 2000 }).catch(() => {
      this.log(`    ⚠ Calendar popup still visible after Escape — proceeding anyway`);
    });
  }

  // Set the "Copy to Billable Duration" toggle to `desired`. Best-effort: reads
  // the current state via aria-checked and only clicks when it differs, so we
  // never accidentally flip an already-correct toggle. If the control isn't
  // found (label/markup differs on this tenant) we log a warning and continue —
  // a missing toggle must not fail the whole entry.
  private async setCopyToBillableToggle(desired: boolean): Promise<void> {
    const page = this.getPage();
    const toggle = page.locator(SELECTORS.copyToBillableToggle).first();
    const visible = await toggle.isVisible({ timeout: 2000 }).catch(() => false);
    if (!visible) {
      this.log(`    ⚠ "Copy to Billable Duration" toggle not found — skipping (leaving D365 default)`);
      return;
    }
    const checkedAttr = await toggle.getAttribute('aria-checked').catch(() => null);
    // Fluent switch exposes aria-checked; a bare <input type=checkbox> may not,
    // so fall back to the DOM `checked` property in that case.
    const current =
      checkedAttr !== null
        ? checkedAttr === 'true'
        : await toggle.isChecked().catch(() => false);
    if (current === desired) {
      this.log(`    Copy to Billable Duration already ${desired ? 'Yes' : 'No'} — no change`);
      return;
    }
    this.log(`    Setting Copy to Billable Duration → ${desired ? 'Yes' : 'No'}`);
    await toggle.click();
    await page.waitForTimeout(200);
  }

  // Save the current Quick Create: Time Entry. When `keepPanelOpen` is true,
  // uses the split-button "Save & Create New" so the next bulk entry can be
  // typed into the same panel without reopening via "+ New". Falls back to a
  // regular Save and Close (panel will close) if the split-button path fails —
  // the next iteration of `fillEntryWithTaskLookup` will detect the closed
  // panel and reopen it.
  private async saveTimeEntry(opts: { keepPanelOpen: boolean; prevTask: string }): Promise<void> {
    const page = this.getPage();
    if (opts.keepPanelOpen) {
      const ok = await this.trySaveAndCreateNew(opts.prevTask);
      if (ok) return;
      this.log(`    Falling back to Save and Close — panel will reopen on next entry`);
    }
    const saveBtn = page.locator(SELECTORS.saveAndCloseBtn).first();
    await saveBtn.waitFor({ state: 'visible', timeout: 8000 });
    await saveBtn.click();
    await page.locator(SELECTORS.quickCreatePanel).waitFor({ state: 'hidden', timeout: 15000 });
    await page.waitForTimeout(500);
    await this.dismissSavedNotification();
  }

  private async trySaveAndCreateNew(prevTask: string): Promise<boolean> {
    const page = this.getPage();
    try {
      this.log(`    Opening Save split-button → "Save & Create New" to keep panel open...`);
      // Scope chevron to the main Quick Create dialog so we don't pick up
      // a chevron in any other split-button on the page (e.g. command bar).
      const quickCreate = page.locator(SELECTORS.quickCreatePanel).first();

      // Try specific aria-label selectors first; if none visible within 3s,
      // fall back to any aria-haspopup button in the panel (the split chevron
      // is the only one there).
      let chevron = quickCreate.locator(SELECTORS.saveSplitButtonChevron).first();
      const specificVisible = await chevron.isVisible({ timeout: 3000 }).catch(() => false);
      if (!specificVisible) {
        // Precise fallback: button[aria-haspopup] whose parent also contains
        // the "Save and Close" button — uniquely targets the split-button chevron
        // without accidentally clicking Type / Duration / Project dropdowns.
        this.log(`    Specific chevron selector not found — trying sibling-based approach`);
        chevron = quickCreate.locator(
          ':has(button:has-text("Save and Close")) > button[aria-haspopup], ' +
          ':has(button:has-text("Save & Close")) > button[aria-haspopup]',
        ).first();
      }
      await chevron.waitFor({ state: 'visible', timeout: 5000 });
      await chevron.click();

      // Menu item renders in a Fluent UI portal (outside the panel), so search
      // page-wide. Cover both "Save & Create New" and "Save and Create New".
      const menuItem = page.locator(
        SELECTORS.saveAndCreateNewMenuItem + ', [role="menuitem"]:has-text("Save and Create New")',
      ).first();
      await menuItem.waitFor({ state: 'visible', timeout: 5000 });
      await menuItem.click();

      // After D365 saves the entry the panel stays visible but renders a fresh
      // empty form. Detect this by polling the project-task lookup input — it
      // held `prevTask` during the previous fill and should clear (or no longer
      // contain `prevTask`) once the round-trip is done. 10s is plenty for a
      // single save; longer would mask real failures.
      this.log(`    Waiting for Quick Create form to reset...`);
      const taskInput = page.locator(SELECTORS.projectTaskField).first();
      const deadline = Date.now() + 10000;
      while (Date.now() < deadline) {
        const val = (await taskInput.inputValue().catch(() => prevTask)) ?? '';
        if (val.trim() === '' || !val.includes(prevTask)) {
          this.log(`    ✓ Form reset — panel ready for next entry`);
          await this.dismissSavedNotification();
          return true;
        }
        await page.waitForTimeout(200);
      }
      throw new Error('Quick Create form did not reset within 10s');
    } catch (err) {
      this.log(`    ⚠ "Save & Create New" failed: ${(err as Error).message}`);
      return false;
    }
  }

  // TODO: Stage 2 — заполнение оставшихся дней многодневных задач через клик
  // по ячейкам "All Weekly Time Entries" grid. После Stage 1 строка задачи уже
  // существует в гриде → находим ячейку (день × задача) → клик → вводим часы.
  // Старый подход (открытие Quick Create для каждого дня) был удалён —
  // см. git history если нужна основа.

  async close(): Promise<void> {
    await this.stopTracing();
    this.page = null;
    if (this.browserMode === 'cdp') {
      if (this.browser) await this.browser.close();
      this.browser = null;
      this.context = null;
    } else {
      if (this.context) {
        await this.context.close();
        this.context = null;
      }
    }
  }

  static fromPage(page: Page, d365Url: string, logger: Logger | null = null): D365Client {
    const client = new D365Client(d365Url, '', 'external', '', logger);
    client.page = page;
    client.context = page.context();
    return client;
  }

  getCurrentPage(): Page {
    return this.getPage();
  }

  async openNewTimeEntry(): Promise<void> {
    const page = this.getPage();
    await page.bringToFront();

    const currentUrl = page.url();
    if (!currentUrl.includes('msdyn_timeentry') && !currentUrl.includes('timeentry')) {
      await this.navigateToTimeEntries();
    }

    this.log(`  Clicking "New Time Entry"...`);
    // The Weekly Time Entries view renders its command bar inside a PCF iframe.
    // Primary path: scoped frameLocator on the stable src path.
    // Defensive fallback: walk every same-origin frame and pick the first with
    // a visible button[aria-label="New"]. Last resort: top-level page locator.
    let frameClicked = false;
    const primaryBtn = page
      .frameLocator(SELECTORS.weeklyGridIframe)
      .locator('button[aria-label="New"]')
      .first();
    try {
      await primaryBtn.waitFor({ state: 'visible', timeout: 15000 });
      this.log(`  Found "New" button via primary frameLocator`);
      await primaryBtn.click();
      frameClicked = true;
    } catch {
      this.log(`  Primary frameLocator did not resolve — scanning all frames`);
      for (const frame of page.frames()) {
        if (frame === page.mainFrame()) continue;
        this.log(`    frame url=${frame.url().slice(0, 80)}`);
        const btn = frame.locator('button[aria-label="New"]').first();
        if (await btn.isVisible().catch(() => false)) {
          this.log(`  Found "New" button while walking frames`);
          await btn.click();
          frameClicked = true;
          break;
        }
      }
    }
    if (!frameClicked) {
      this.log(`  No iframe button found — falling back to top-level locator`);
      const newBtn = page.locator(SELECTORS.newTimeEntryBtn).first();
      await newBtn.waitFor({ state: 'visible', timeout: 15000 });
      await newBtn.click();
    }

    this.log(`  Waiting for Quick Create panel...`);
    await page.locator(SELECTORS.quickCreatePanel).waitFor({ state: 'visible', timeout: 15000 });
    await page.waitForTimeout(300);
    this.log(`  Panel opened`);
  }

  private getPage(): Page {
    if (!this.page) throw new Error('Browser not launched. Call launch() first.');
    return this.page;
  }
}
