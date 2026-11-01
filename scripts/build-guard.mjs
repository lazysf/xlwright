import { lstat, readdir, readFile, stat } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { dirname, join, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

export const repository = resolve(dirname(fileURLToPath(import.meta.url)), '..');
export const development = join(repository, '.xlwright', 'development');

export async function rejectLinks(path) {
  for (let current = path; ; current = dirname(current)) {
    try { if ((await lstat(current)).isSymbolicLink()) throw new Error('Development paths may not use symbolic links or junctions.'); }
    catch (error) { if (error.code !== 'ENOENT') throw error; }
    if (dirname(current) === current) break;
  }
}

export async function sourceFiles(sources) {
  const entries = await readdir(join(repository, sources), { recursive: true, withFileTypes: true });
  return entries.map(entry => relative(repository, join(entry.parentPath, entry.name)))
    .filter(path => /\.(cs|csproj)$/.test(path) && !/[\\/](bin|obj)[\\/]/.test(path));
}

// A quick guard for inspection. It sees edits, not deletions; experiments compare hashes instead.
export async function rejectStaleBuild(project, sources) {
  const built = (await stat(join(development, 'artifacts', 'bin', project, 'release', `${project}.dll`))).mtimeMs;
  for (const path of await sourceFiles(sources))
    if ((await stat(join(repository, path))).mtimeMs > built) throw new Error(`Build is older than ${path}; run npm run check.`);
}

export async function sha256File(path) {
  return createHash('sha256').update(await readFile(path)).digest('hex').toUpperCase();
}

const hashEntries = (root, paths) => Promise.all(paths.sort().map(async path =>
  ({ path: path.replaceAll('\\', '/'), sha256: await sha256File(join(root, path)) })));

export const spikeSources = async () => hashEntries(repository, await sourceFiles('spikes/Spike0'));

export const spikeOutput = join(development, 'artifacts', 'bin', 'Spike0', 'release');

// The product copy holds the server and the real-Excel scenarios, which start that server from ../../XlWright.McpServer.
export const productSources = async () =>
  hashEntries(repository, [...await sourceFiles('src'), ...await sourceFiles('tests')]);

export const productOutputs = Object.fromEntries(['XlWright.McpServer', 'XlWright.RealExcel.Tests']
  .map(project => [project, join(development, 'artifacts', 'bin', project, 'release')]));

export async function directoryHashes(root) {
  const entries = await readdir(root, { recursive: true, withFileTypes: true });
  return hashEntries(root, entries.filter(entry => entry.isFile())
    .map(entry => relative(root, join(entry.parentPath, entry.name))));
}
