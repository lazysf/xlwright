import { cp, lstat, open, readFile, rename, rm } from 'node:fs/promises';
import { execFileSync } from 'node:child_process';
import { randomUUID } from 'node:crypto';
import { join, relative, resolve } from 'node:path';
import { isDeepStrictEqual } from 'node:util';
import { development, directoryHashes, productOutputs, productSources, rejectLinks, repository, sha256File, spikeOutput, spikeSources } from './build-guard.mjs';
import { writeManifest } from './write-manifest.mjs';

// No argument copies Spike0. --product copies the server and the real-Excel scenarios.
const args = process.argv.slice(2);
if (!(args.length === 0 || args.length === 1 && args[0] === '--product'))
  throw new Error('Usage: prepare-experiment.mjs [--product]');
const product = args[0] === '--product';
if (resolve(process.cwd()) !== repository) throw new Error('Run from the repository root; the copies resolve their assets from there.');

const git = (...args) => execFileSync('git', args, { cwd: repository, encoding: 'utf8' }).trim();
const id = randomUUID().replaceAll('-', '');
const prepared = join(development, 'prepared');
const incoming = join(prepared, `incoming-${id}`);
const target = join(prepared, `${product ? 'product' : 'experiment'}-${id}`);
const outputs = product ? Object.entries(productOutputs).map(([project, output]) => ({ project, output, copy: join(incoming, 'bin', project, 'release') }))
  : [{ project: 'Spike0', output: spikeOutput, copy: incoming }];
for (const path of [prepared, ...outputs.map(item => item.output), join(development, 'check.lock'), join(development, 'last-passed.json')]) await rejectLinks(path);

const copyOutput = (source, destination) => cp(source, destination, { recursive: true, filter: async path => {
  if ((await lstat(path)).isSymbolicLink()) throw new Error('Build outputs may not contain symbolic links.');
  return true;
} });

// The integrated check holds this lock exclusively, so a running build cannot be copied.
let buildLock;
try { buildLock = await open(join(development, 'check.lock'), 'a+'); }
catch (error) { if (error.code === 'EBUSY') throw new Error('A check is running; wait for it to finish.'); throw error; }
try {
  let lastPassed;
  try { lastPassed = JSON.parse(await readFile(join(development, 'last-passed.json'), 'utf8')); }
  catch (error) { if (error.code === 'ENOENT') throw new Error('The latest full check did not pass; run npm run check.'); throw error; }
  const sources = product ? await productSources() : await spikeSources();
  const passed = product ? lastPassed.product : { sources: lastPassed.sources, outputs: { Spike0: lastPassed.files } };
  if (passed === undefined || !isDeepStrictEqual(sources, passed.sources))
    throw new Error('Sources changed since the last passed check; run npm run check.');
  for (const { project, output, copy } of outputs) {
    if (!isDeepStrictEqual(await directoryHashes(output), passed.outputs[project])) throw new Error('Build differs from the last passed check; run npm run check.');
    await copyOutput(output, copy);
    if (!isDeepStrictEqual(await directoryHashes(copy), passed.outputs[project])) throw new Error('The copy differs from the last passed check.');
  }
  // The manifest must name exactly the verified build, so the whole copy is compared once more before it is written.
  const files = await directoryHashes(incoming);
  const verified = outputs.flatMap(({ project, copy }) => passed.outputs[project].map(file =>
    ({ path: [relative(incoming, copy).replaceAll('\\', '/'), file.path].filter(Boolean).join('/'), sha256: file.sha256 })));
  const byPath = list => [...list].sort((a, b) => a.path < b.path ? -1 : a.path > b.path ? 1 : 0);
  if (!isDeepStrictEqual(byPath(files), byPath(verified)))
    throw new Error('The copy differs from the last passed check.');
  await writeManifest(incoming, lastPassed.runId, repository, files, {
    commit: git('rev-parse', 'HEAD'),
    worktreeDirty: git('status', '--porcelain') !== '',
    ...(product ? {} : { fixtureHash: await sha256File(join(repository, 'tests', 'fixtures', 'lifecycle.xlsx')) }),
    sources
  });
  await rename(incoming, target);
} finally {
  await buildLock.close();
  await rm(incoming, { recursive: true, force: true });
}

const entry = product ? join(target, 'bin', 'XlWright.RealExcel.Tests', 'release', 'XlWright.RealExcel.Tests.dll') : join(target, 'Spike0.dll');
console.log(`Prepared ${relative(repository, target)}`);
console.log(`Run from the repository root: dotnet ${relative(repository, entry).replaceAll('\\', '/')}${product ? '' : ' <experiment argument>'}`);
