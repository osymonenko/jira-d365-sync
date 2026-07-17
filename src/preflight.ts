import * as http from 'http';
import * as fs from 'fs';
import { parseTimesheet } from './excel';

function httpGet(url: string, timeoutMs = 4000): Promise<string> {
  return new Promise((resolve, reject) => {
    const req = http.get(url, { timeout: timeoutMs }, (res) => {
      let data = '';
      res.on('data', (chunk: Buffer) => (data += chunk.toString()));
      res.on('end', () => resolve(data));
    });
    req.on('error', reject);
    req.on('timeout', () => {
      req.destroy();
      reject(new Error('timeout'));
    });
  });
}

export async function runPreflight(config: {
  d365Url: string;
  browserMode: string;
  cdpUrl: string;
  userDataDir: string;
  excelFile: string;
}): Promise<boolean> {
  let allOk = true;
  console.log('\n── Preflight checks ──────────────────────────');

  // 1. Excel file
  process.stdout.write(`[1/3] Excel file ... `);
  if (!fs.existsSync(config.excelFile)) {
    console.log('✗ NOT FOUND');
    console.log(`      Path: ${config.excelFile}`);
    allOk = false;
  } else {
    try {
      const entries = parseTimesheet(config.excelFile);
      console.log(`✓  ${entries.length} time entries found`);
    } catch (e) {
      console.log(`✗ PARSE ERROR`);
      console.log(`      ${(e as Error).message}`);
      allOk = false;
    }
  }

  // 2. D365 URL
  process.stdout.write(`[2/3] D365 URL ... `);
  if (!config.d365Url || !config.d365Url.startsWith('http')) {
    console.log('✗ NOT SET — add D365_URL to .env');
    allOk = false;
  } else {
    console.log(`✓  ${config.d365Url}`);
  }

  // 3. Browser
  process.stdout.write(`[3/3] Browser ... `);
  if (config.browserMode === 'cdp') {
    try {
      const raw = await httpGet(`${config.cdpUrl}/json/version`);
      const info = JSON.parse(raw) as { Browser?: string };
      console.log(`✓  CDP connected — ${info.Browser ?? 'Chrome'}`);
    } catch {
      console.log(`✗ CANNOT REACH Chrome at ${config.cdpUrl}`);
      console.log(`      TIP: Switch to chrome-profile mode (BROWSER_MODE=chrome-profile in .env)`);
      allOk = false;
    }
  } else if (config.browserMode === 'chrome-profile') {
    const profileExists = fs.existsSync(config.userDataDir);
    if (profileExists) {
      console.log(`✓  chrome-profile — using saved profile at ${config.userDataDir}`);
    } else {
      console.log(`✓  chrome-profile — first run (you'll need to sign in once; cookies will be saved to ${config.userDataDir})`);
    }
  } else {
    console.log(`✓  Persistent Chromium — ${config.userDataDir}`);
  }

  console.log('──────────────────────────────────────────────');
  if (allOk) {
    console.log('✓  All checks passed — starting automation...\n');
  } else {
    console.log('✗  Fix the issues above before running.\n');
  }
  return allOk;
}
