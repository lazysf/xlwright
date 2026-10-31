import { cp, lstat, open, readFile, rename, rm, writeFile } from 'node:fs/promises';
import { execFileSync } from 'node:child_process';
import { randomUUID } from 'node:crypto';
import { join, relative, resolve } from 'node:path';
import { isDeepStrictEqual } from 'node:util';
import { development, directoryHashes, rejectLinks, repository, sha256File, spikeOutput, spikeSources } from './build-guard.mjs';

if (resolve(process.cwd()) !== repository) throw new Error('Run from the repository root; Spike0 resolves its assets from there.');

const git = (...args) => execFileSync('git', args, { cwd: repository, encoding: 'utf8' }).trim();
const id = randomUUID().replaceAll('-', '');
const prepared = join(development, 'prepared');
const incoming = join(prepared, `incoming-${id}`);
const target = join(prepared, `experiment-${id}`);
for (const path of [prepared, spikeOutput, join(development, 'check.lock'), join(development, 'last-passed.json')]) await rejectLinks(path);

// The integrated check holds this lock exclusively, so a running build cannot be copied.
let buildLock;
try { buildLock = await open(join(development, 'check.lock'), 'a+'); }
catch (error) { if (error.code === 'EBUSY') throw new Error('A check is running; wait for it to finish.'); throw error; }
try {
  let lastPassed;
  try { lastPassed = JSON.parse(await readFile(join(development, 'last-passed.json'), 'utf8')); }
  catch (error) { if (error.code === 'ENOENT') throw new Error('The latest full check did not pass; run npm run check.'); throw error; }
  const sources = await spikeSources();
  if (!isDeepStrictEqual(sources, lastPassed.sources)) throw new Error('Spike0 sources changed since the last passed check; run npm run check.');
  if (!isDeepStrictEqual(await directoryHashes(spikeOutput), lastPassed.files)) throw new Error('Build differs from the last passed check; run npm run check.');
  await cp(spikeOutput, incoming, { recursive: true, filter: async path => {
    if ((await lstat(path)).isSymbolicLink()) throw new Error('Build outputs may not contain symbolic links.');
    return true;
  } });
  const files = await directoryHashes(incoming);
  if (!isDeepStrictEqual(files, lastPassed.files)) throw new Error('The copy differs from the last passed check.');
  await writeFile(join(incoming, 'manifest.json'), `${JSON.stringify({
    checkRun: lastPassed.runId,
    commit: git('rev-parse', 'HEAD'),
    worktreeDirty: git('status', '--porcelain') !== '',
    fixtureHash: await sha256File(join(repository, 'tests', 'fixtures', 'lifecycle.xlsx')),
    files,
    sources
  }, null, 2)}\n`);
  await rename(incoming, target);
} finally {
  await buildLock.close();
  await rm(incoming, { recursive: true, force: true });
}

console.log(`Prepared ${relative(repository, target)}`);
console.log(`Run from the repository root: dotnet ${relative(repository, join(target, 'Spike0.dll')).replaceAll('\\', '/')} <experiment argument>`);
