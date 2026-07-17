import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import * as XLSX from 'xlsx';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { firstEntriesPerTask, parseTimesheet, remainingEntries, uniqueTasks } from './excel';

// Excel serial for a given UTC date (days since 1899-12-30, the offset xlsx uses).
function dateToSerial(year: number, month: number, day: number): number {
  const epoch = Date.UTC(1899, 11, 30);
  const target = Date.UTC(year, month - 1, day);
  return (target - epoch) / 86400000;
}

let fixtureFile: string;
let tmpDir: string;

beforeAll(() => {
  tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'd365-test-'));
  fixtureFile = path.join(tmpDir, 'fixture.xlsx');

  // Two weeks: Sun 2026-04-26 → Sat 2026-05-02 (Mon 2026-04-27)
  //            Sun 2026-05-03 → Sat 2026-05-09 (Mon 2026-05-04)
  const week1Start = dateToSerial(2026, 4, 26);
  const week1End = dateToSerial(2026, 5, 2);
  const week2Start = dateToSerial(2026, 5, 3);
  const week2End = dateToSerial(2026, 5, 9);

  const aoa: (string | number | null)[][] = [
    ['weekstart', 'weekend', 'Task', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri'],
    [week1Start, week1End, null, null, null, null, null, null],
    [null, null, '  ProjectAlpha  ', 8, 8, 0, null, 4],
    [null, null, 'ProjectBeta', null, null, 8, 8, 4],
    [week2Start, week2End, null, null, null, null, null, null],
    [null, null, 'ProjectAlpha', 8, 8, 8, 8, 8],
    [null, null, '', 1, 1, 1, 1, 1], // empty task name → must be skipped
  ];

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
  XLSX.writeFile(wb, fixtureFile);
});

afterAll(() => {
  fs.rmSync(tmpDir, { recursive: true, force: true });
});

describe('parseTimesheet', () => {
  it('parses both weeks into flat TimeEntry array, skipping zero / empty hours', () => {
    const entries = parseTimesheet(fixtureFile);
    // Week 1: Alpha Mon=8, Tue=8, Fri=4 (Wed=0 skipped, Thu=null skipped) → 3
    //         Beta  Wed=8, Thu=8, Fri=4 → 3
    // Week 2: Alpha Mon-Fri=8 each → 5
    //         '' task → skipped entirely
    expect(entries).toHaveLength(11);
  });

  it('trims whitespace from task names', () => {
    const entries = parseTimesheet(fixtureFile);
    const tasks = new Set(entries.map((e) => e.task));
    expect(tasks.has('ProjectAlpha')).toBe(true);
    expect(tasks.has('  ProjectAlpha  ')).toBe(false);
  });

  it('formats dates as M/D/YYYY (no leading zeros)', () => {
    const entries = parseTimesheet(fixtureFile);
    // Week 1 Mon = 2026-04-27
    const week1Mon = entries.find((e) => e.date === '4/27/2026');
    expect(week1Mon).toBeDefined();
    // Week 2 Fri = 2026-05-08
    const week2Fri = entries.find((e) => e.date === '5/8/2026');
    expect(week2Fri).toBeDefined();
  });

  it('Mon..Fri map to weekStart+1 .. weekStart+5 (weekStart is Sunday in the sheet)', () => {
    const entries = parseTimesheet(fixtureFile);
    // Filter to week 2 (weekStart Sun 2026-05-03) — Alpha has all 5 weekdays.
    const week2Sun = new Date(Date.UTC(2026, 4, 3));
    const alphaWeek2 = entries.filter(
      (e) => e.task === 'ProjectAlpha' && e.weekStart.getTime() === week2Sun.getTime(),
    );
    const dates = alphaWeek2.map((e) => e.date).sort();
    expect(dates).toEqual(['5/4/2026', '5/5/2026', '5/6/2026', '5/7/2026', '5/8/2026']);
  });

  it('weekFilter narrows to a single Monday with ±1 day tolerance', () => {
    const monday = new Date(Date.UTC(2026, 4, 4)); // 2026-05-04 (Monday of week 2)
    const entries = parseTimesheet(fixtureFile, monday);
    expect(entries.every((e) => e.date.startsWith('5/'))).toBe(true);
    expect(entries).toHaveLength(5); // only Alpha, all 5 weekdays
  });

  it('weekFilter that does not match any week returns empty array', () => {
    const noMatch = new Date(Date.UTC(2030, 0, 6));
    const entries = parseTimesheet(fixtureFile, noMatch);
    expect(entries).toEqual([]);
  });

  it('throws on non-existent file', () => {
    expect(() => parseTimesheet(path.join(tmpDir, 'does-not-exist.xlsx'))).toThrow();
  });
});

describe('parseTimesheet — "check sum" row is skipped', () => {
  let checkSumFile: string;

  beforeAll(() => {
    checkSumFile = path.join(tmpDir, 'with-check-sum.xlsx');
    const weekStart = dateToSerial(2026, 5, 3);
    const weekEnd = dateToSerial(2026, 5, 9);
    const aoa: (string | number | null)[][] = [
      ['weekstart', 'weekend', 'Task', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri'],
      [weekStart, weekEnd, 'Internal daily meeting', 0.5, 0.5, 0.5, 0.5, 0.5],
      [null, null, 'check sum', 8, 8, 8, 8, 8],
      [null, null, 'Check Sum', 8, 8, 8, 8, 8], // different casing
      [null, null, '  check sum  ', 8, 8, 8, 8, 8], // surrounding whitespace
    ];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    XLSX.writeFile(wb, checkSumFile);
  });

  it('does not emit entries for "check sum" rows regardless of casing/whitespace', () => {
    const entries = parseTimesheet(checkSumFile);
    expect(entries.every((e) => e.task.toLowerCase().trim() !== 'check sum')).toBe(true);
    // Only the real "Internal daily meeting" row produces entries → 5 weekdays
    expect(entries).toHaveLength(5);
    expect(new Set(entries.map((e) => e.task))).toEqual(new Set(['Internal daily meeting']));
  });
});

describe('firstEntriesPerTask', () => {
  it('returns one entry per task — the earliest date', () => {
    const entries = [
      { date: '4/28/2026', task: 'Alpha', hours: 8, weekStart: new Date() },
      { date: '4/27/2026', task: 'Alpha', hours: 8, weekStart: new Date() },
      { date: '4/29/2026', task: 'Alpha', hours: 4, weekStart: new Date() },
      { date: '4/27/2026', task: 'Beta', hours: 8, weekStart: new Date() },
    ];
    const result = firstEntriesPerTask(entries);
    expect(result).toHaveLength(2);
    const alpha = result.find((e) => e.task === 'Alpha');
    expect(alpha?.date).toBe('4/27/2026');
    const beta = result.find((e) => e.task === 'Beta');
    expect(beta?.date).toBe('4/27/2026');
  });

  it('returns single-day task as-is', () => {
    const entries = [{ date: '5/1/2026', task: 'Solo', hours: 8, weekStart: new Date() }];
    expect(firstEntriesPerTask(entries)).toHaveLength(1);
  });
});

describe('remainingEntries', () => {
  it('excludes first day per task, keeps all other days', () => {
    const entries = [
      { date: '4/27/2026', task: 'Alpha', hours: 8, weekStart: new Date() },
      { date: '4/28/2026', task: 'Alpha', hours: 8, weekStart: new Date() },
      { date: '4/29/2026', task: 'Alpha', hours: 4, weekStart: new Date() },
      { date: '4/27/2026', task: 'Beta', hours: 8, weekStart: new Date() },
    ];
    const result = remainingEntries(entries);
    expect(result).toHaveLength(2);
    expect(result.every((e) => e.task === 'Alpha')).toBe(true);
    const dates = result.map((e) => e.date).sort();
    expect(dates).toEqual(['4/28/2026', '4/29/2026']);
  });

  it('returns empty array when every task has only one entry', () => {
    const entries = [
      { date: '5/1/2026', task: 'A', hours: 8, weekStart: new Date() },
      { date: '5/1/2026', task: 'B', hours: 4, weekStart: new Date() },
    ];
    expect(remainingEntries(entries)).toHaveLength(0);
  });
});

describe('uniqueTasks', () => {
  it('dedupes preserving first-seen order, case-sensitively', () => {
    const entries = [
      { date: '1/1/2026', task: 'Alpha', hours: 1, weekStart: new Date() },
      { date: '1/2/2026', task: 'Beta', hours: 1, weekStart: new Date() },
      { date: '1/3/2026', task: 'Alpha', hours: 1, weekStart: new Date() },
      { date: '1/4/2026', task: 'alpha', hours: 1, weekStart: new Date() },
    ];
    expect(uniqueTasks(entries)).toEqual(['Alpha', 'Beta', 'alpha']);
  });
});
