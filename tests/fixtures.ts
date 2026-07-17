import { test as base, BrowserContext, Page, chromium } from '@playwright/test';
import * as path from 'path';

type WorkerFixtures = {
  edgeContext: BrowserContext;
  edgePage: Page;
};

/**
 * Worker-scoped persistent Edge context using ./edge-profile.
 *
 * The CLI (`npm start`) uses the same profile, so SSO cookies are shared.
 * NB: only one process can hold the profile lock at a time — don't run
 * `npm start` and `npx playwright test` simultaneously.
 */
export const test = base.extend<{}, WorkerFixtures>({
  edgeContext: [
    async ({}, use) => {
      const profileDir = path.resolve(process.env.USER_DATA_DIR || './edge-profile');
      const ctx = await chromium.launchPersistentContext(profileDir, {
        channel: 'msedge',
        headless: false,
        args: ['--start-maximized'],
        ignoreDefaultArgs: ['--enable-automation'],
        viewport: null,
      });
      await use(ctx);
      await ctx.close();
    },
    { scope: 'worker' },
  ],

  edgePage: [
    async ({ edgeContext }, use) => {
      const page = edgeContext.pages()[0] ?? (await edgeContext.newPage());
      await use(page);
    },
    { scope: 'worker' },
  ],
});

export { expect } from '@playwright/test';
