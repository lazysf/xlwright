import { readFile, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import assert from 'node:assert/strict';
import { development, directoryHashes, productOutputs, productSources, rejectLinks, spikeOutput, spikeSources } from './build-guard.mjs';

// Called by scripts/development.ps1 around a full check: --begin before the build, --finish after every stage passed.
const [mode, runId] = process.argv.slice(2);
if (!['--begin', '--finish'].includes(mode) || !/^[0-9a-f]{32}$/.test(runId ?? ''))
  throw new Error('Usage: record-build.mjs --begin <run-id> | --finish <run-id>');

const started = join(development, 'runs', runId, 'sources.json');
const readSources = async () => ({ spike: await spikeSources(), product: await productSources() });
if (mode === '--begin') {
  await writeFile(started, JSON.stringify(await readSources(), null, 2));
} else {
  const sources = await readSources();
  const recorded = JSON.parse(await readFile(started, 'utf8'));
  assert.deepEqual(sources.spike, recorded.spike, 'Spike0 sources changed during the check; run it again.');
  assert.deepEqual(sources.product, recorded.product, 'Product sources changed during the check; run it again.');
  const product = { sources: sources.product, outputs: {} };
  for (const [project, output] of Object.entries(productOutputs)) product.outputs[project] = await directoryHashes(output);
  const lastPassed = join(development, 'last-passed.json');
  await rejectLinks(lastPassed);
  await writeFile(lastPassed, JSON.stringify({ runId, sources: sources.spike, files: await directoryHashes(spikeOutput), product }, null, 2));
}
