import { readFile, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import assert from 'node:assert/strict';
import { development, directoryHashes, rejectLinks, spikeOutput, spikeSources } from './build-guard.mjs';

// Called by scripts/development.ps1 around a full check: --begin before the build, --finish after every stage passed.
const [mode, runId] = process.argv.slice(2);
if (!['--begin', '--finish'].includes(mode) || !/^[0-9a-f]{32}$/.test(runId ?? ''))
  throw new Error('Usage: record-build.mjs --begin <run-id> | --finish <run-id>');

const started = join(development, 'runs', runId, 'sources.json');
if (mode === '--begin') {
  await writeFile(started, JSON.stringify(await spikeSources(), null, 2));
} else {
  const sources = await spikeSources();
  assert.deepEqual(sources, JSON.parse(await readFile(started, 'utf8')), 'Spike0 sources changed during the check; run it again.');
  const lastPassed = join(development, 'last-passed.json');
  await rejectLinks(lastPassed);
  await writeFile(lastPassed, JSON.stringify({ runId, sources, files: await directoryHashes(spikeOutput) }, null, 2));
}
