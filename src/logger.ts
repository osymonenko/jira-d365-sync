import * as fs from 'fs';
import * as path from 'path';

export interface Logger {
  log(msg: string): void;
  error(msg: string): void;
  runDir: string;
  close(): Promise<void>;
}

function timestamp(): string {
  return new Date().toTimeString().slice(0, 8);
}

export function createRunId(): string {
  return new Date().toISOString().replace(/[:.]/g, '-');
}

export function createLogger(runId: string): Logger {
  const runDir = path.resolve('logs', runId);
  fs.mkdirSync(runDir, { recursive: true });
  const stream = fs.createWriteStream(path.join(runDir, 'run.log'), { flags: 'a' });

  const write = (level: 'INFO' | 'ERROR', msg: string): void => {
    const line = `[${timestamp()}] [${level}] ${msg}`;
    if (level === 'ERROR') console.error(line);
    else console.log(line);
    stream.write(line + '\n');
  };

  return {
    runDir,
    log: (msg) => write('INFO', msg),
    error: (msg) => write('ERROR', msg),
    close: () =>
      new Promise<void>((resolve) => {
        stream.end(resolve);
      }),
  };
}
