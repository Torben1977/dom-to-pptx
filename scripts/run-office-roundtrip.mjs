import { spawnSync } from 'node:child_process';
import path from 'node:path';
import process from 'node:process';

const vitest = path.resolve('node_modules/vitest/vitest.mjs');
const result = spawnSync(process.execPath, [vitest, 'run', 'src/__tests__/office-roundtrip.integration.test.js'], {
  stdio: 'inherit',
  env: { ...process.env, DOM_TO_PPTX_OFFICE_ROUNDTRIP: '1' },
});

if (result.error) throw result.error;
process.exitCode = result.status ?? 1;
