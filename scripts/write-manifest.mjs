import { writeFile } from 'node:fs/promises';
import { join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { development, directoryHashes } from './build-guard.mjs';

// The only writer of manifest.json, so the tests check the fixtures against the format real copies use.
export async function writeManifest(directory, checkRun, repository, files, details = {}) {
  await writeFile(join(directory, 'manifest.json'), `${JSON.stringify({ checkRun, repository, ...details, files }, null, 2)}\n`);
}

// npm run check cannot run prepare-experiment.mjs, so tests use this entry for a synthetic copy.
if (process.argv[1] !== undefined && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const [directory, checkRun, repository] = process.argv.slice(2);
  if (repository === undefined) throw new Error('Usage: write-manifest.mjs <directory> <check-run> <repository>');
  // This entry skips the checks of prepare-experiment.mjs, so it must not be able to bless a real copy.
  if (resolve(directory).startsWith(join(development, 'prepared'))) throw new Error('Use npm run prepare:product for prepared copies.');
  await writeManifest(directory, checkRun, repository, await directoryHashes(directory));
}
