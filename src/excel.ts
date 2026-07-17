import * as XLSX from 'xlsx';
import { DAY_KEYS, DayKey, TimeEntry, WeekBlock } from './types';

// Excel stores dates as days since Dec 30, 1899 (with the 1900 leap-year bug offset).
// The result is a UTC-midnight Date — all downstream date math and formatting MUST
// use UTC getters/setters (getUTCDate/setUTCDate/...), otherwise a machine in a
// UTC-negative timezone reads the day back as the *previous* calendar day.
function excelSerialToDate(serial: number): Date {
  return new Date((serial - 25569) * 86400 * 1000);
}

function toD365Date(d: Date): string {
  return `${d.getUTCMonth() + 1}/${d.getUTCDate()}/${d.getUTCFullYear()}`;
}

// D365's Payable Duration dropdown only offers 15-minute increments up to 8h.
// Anything outside that grid can't be selected and would fail late (browser
// timeout on a non-existent option), so we reject it here — at parse time — with
// enough context (task + date) for the user to fix the Excel cell.
function validateHours(hours: number, task: string, dateLabel: string): void {
  const onGrid = Math.abs(hours * 4 - Math.round(hours * 4)) < 1e-9;
  if (!onGrid || hours <= 0 || hours > 8) {
    throw new Error(
      `Invalid hours ${hours} for "${task}" on ${dateLabel}: ` +
        `D365 accepts only 15-minute increments (0.25, 0.5, …) up to 8 hours.`,
    );
  }
}

export function parseTimesheet(filePath: string, weekFilter?: Date): TimeEntry[] {
  const wb = XLSX.readFile(filePath);
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, defval: null });

  const weeks: WeekBlock[] = [];
  let currentWeek: WeekBlock | null = null;

  // Row 0 may be a header (weekstart | weekend | Task | Mon | Tue | Wed | Thu | Fri) or already data.
  // Start at 0 — header rows naturally fail both block checks below (string col A, no currentWeek yet).
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i] as (string | number | null)[];
    const colA = row[0];
    const colB = row[1];
    const colC = row[2];

    // New week block: col A contains a numeric Excel serial date
    if (typeof colA === 'number' && colA > 40000) {
      currentWeek = {
        weekStart: excelSerialToDate(colA),
        weekEnd: excelSerialToDate(typeof colB === 'number' ? colB : colA + 6),
        tasks: [],
      };
      weeks.push(currentWeek);
    }

    // Task row: col C is a non-empty string
    if (typeof colC === 'string' && colC.trim().length > 0 && currentWeek) {
      const taskName = colC.trim();
      // "check sum" is a user-maintained verification row (column totals), not a real task — skip it
      if (taskName.toLowerCase() === 'check sum') continue;
      const hours: Partial<Record<DayKey, number>> = {};

      // Columns D–H (index 3–7) correspond to Mon–Fri
      DAY_KEYS.forEach((day, idx) => {
        const val = row[3 + idx];
        if (typeof val === 'number' && val > 0) {
          hours[day] = val;
        }
      });

      if (Object.keys(hours).length > 0) {
        currentWeek.tasks.push({ task: taskName, hours });
      }
    }
  }

  // Flatten week blocks into individual TimeEntry records
  const entries: TimeEntry[] = [];

  for (const week of weeks) {
    if (weekFilter) {
      // Filter to the week whose Monday matches weekFilter (within 1 day tolerance)
      const weekMonday = new Date(week.weekStart);
      weekMonday.setUTCDate(weekMonday.getUTCDate() + 1);
      const filterMs = weekFilter.getTime();
      const mondayMs = weekMonday.getTime();
      if (Math.abs(filterMs - mondayMs) > 86400 * 1000) continue;
    }

    for (const taskRow of week.tasks) {
      DAY_KEYS.forEach((day, idx) => {
        const hours = taskRow.hours[day];
        if (!hours) return;

        const entryDate = new Date(week.weekStart);
        entryDate.setUTCDate(entryDate.getUTCDate() + 1 + idx); // Mon=+1 … Fri=+5
        const dateLabel = toD365Date(entryDate);

        validateHours(hours, taskRow.task, dateLabel);

        entries.push({
          date: dateLabel,
          task: taskRow.task,
          hours,
          weekStart: week.weekStart,
        });
      });
    }
  }

  return entries;
}

export function uniqueTasks(entries: TimeEntry[]): string[] {
  return [...new Set(entries.map((e) => e.task))];
}

export function firstEntriesPerTask(entries: TimeEntry[]): TimeEntry[] {
  const seen = new Map<string, TimeEntry>();
  for (const entry of [...entries].sort((a, b) => new Date(a.date).getTime() - new Date(b.date).getTime())) {
    if (!seen.has(entry.task)) seen.set(entry.task, entry);
  }
  return [...seen.values()];
}

export function remainingEntries(entries: TimeEntry[]): TimeEntry[] {
  const firstDates = new Map<string, string>();
  for (const entry of [...entries].sort((a, b) => new Date(a.date).getTime() - new Date(b.date).getTime())) {
    if (!firstDates.has(entry.task)) firstDates.set(entry.task, entry.date);
  }
  return entries.filter((e) => e.date !== firstDates.get(e.task));
}
