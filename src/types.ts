export interface WeekBlock {
  weekStart: Date;
  weekEnd: Date;
  tasks: TaskRow[];
}

export interface TaskRow {
  task: string;
  hours: Partial<Record<DayKey, number>>;
}

export type DayKey = 'Mon' | 'Tue' | 'Wed' | 'Thu' | 'Fri';

export const DAY_KEYS: DayKey[] = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri'];

export interface TimeEntry {
  /** M/D/YYYY — the format D365 expects in the Date field */
  date: string;
  task: string;
  hours: number;
  weekStart: Date;
}
