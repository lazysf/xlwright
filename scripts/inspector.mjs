import { cp, lstat, mkdir, open, readFile, writeFile } from 'node:fs/promises';
import { spawn, spawnSync } from 'node:child_process';
import { createHash, randomUUID } from 'node:crypto';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import assert from 'node:assert/strict';

const repository = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const mode = process.argv[2];
if (!['--check', '--cli', '--web'].includes(mode) || (mode !== '--cli' && process.argv.length !== 3))
  throw new Error('Usage: inspector.mjs --check | --web | --cli [Inspector CLI flags]');
if (resolve(process.cwd()) !== repository) throw new Error('Run the Inspector from the repository root.');
const development = join(repository, '.xlwright', 'development');
async function rejectLinks(path) {
  for (let current = path; ; current = dirname(current)) {
    try { if ((await lstat(current)).isSymbolicLink()) throw new Error('Inspector paths may not use symbolic links or junctions.'); }
    catch (error) { if (error.code !== 'ENOENT') throw error; }
    if (dirname(current) === current) break;
  }
}

const sessions = join(development, 'inspector');
await rejectLinks(sessions);
await mkdir(sessions, { recursive: true });
await rejectLinks(join(development, 'check.lock'));
// The integrated check already holds this lock. Standalone snapshots must not race a build.
const buildLock = mode !== '--check' ? await open(join(development, 'check.lock'), 'a+') : undefined;
const session = join(sessions, randomUUID());
await mkdir(session);
try {
  for (const project of ['XlWright.McpServer', 'Spike0']) {
    const source = join(development, 'artifacts', 'bin', project, 'release');
    await rejectLinks(source);
    await cp(source, join(session, project), { recursive: true, filter: async path => {
      if ((await lstat(path)).isSymbolicLink()) throw new Error('Build outputs may not contain symbolic links.');
      return true;
    } });
  }
} finally { await buildLock?.close(); }

const server = join(session, 'XlWright.McpServer', 'XlWright.McpServer.dll');
const spike = join(session, 'Spike0', 'Spike0.dll');
const config = join(session, 'config.json');
await writeFile(config, JSON.stringify({ mcpServers: { xlwright: {
  type: 'stdio', command: 'dotnet', args: [server, '--development'], cwd: repository,
  env: { XLWRIGHT_DEVELOPMENT_SPIKE_DLL: spike }
} } }, null, 2));

const env = { ...process.env, MCP_INSPECTOR_SECRET_STORE: 'memory', MCP_STORAGE_DIR: session,
  MCP_CLIENT_CONFIG_PATH: join(session, 'client.json'), MCP_INSPECTOR_OAUTH_STATE_PATH: join(session, 'oauth.json'),
  MCP_LOG_FILE: join(session, 'inspector.log'), HOST: '127.0.0.1' };
delete env.DANGEROUSLY_OMIT_AUTH;
delete env.DANGEROUSLY_BIND_ALL_INTERFACES;
delete env.MCP_INSPECTOR_API_TOKEN;
delete env.MCP_PROXY_AUTH_TOKEN;
const launcher = join(repository, 'node_modules', '@modelcontextprotocol', 'inspector', 'clients', 'launcher', 'build', 'index.js');
const cliArguments = ['--cli', '--config', config, '--server', 'xlwright', '--stored-auth-only', '--format', 'json'];
if (mode === '--cli') {
  const argumentsToForward = process.argv.length > 3 ? process.argv.slice(3) : ['--method', 'tools/list'];
  const result = spawnSync(process.execPath, [launcher, ...cliArguments, ...argumentsToForward],
    { cwd: repository, env, stdio: 'inherit', windowsHide: true, timeout: 60_000 });
  if (result.error) throw result.error;
  process.exitCode = result.status ?? 1;
} else if (mode === '--web') {
  const child = spawn(process.execPath, [launcher, '--web', '--config', config], { cwd: repository, env, stdio: 'inherit', windowsHide: true });
  child.on('error', error => { console.error(error); process.exitCode = 1; });
  child.on('exit', code => { process.exitCode = code ?? 1; });
} else {
  const invoke = args => {
    const result = spawnSync(process.execPath, [launcher, ...cliArguments, ...args],
      { cwd: repository, env, encoding: 'utf8', windowsHide: true, timeout: 60_000 });
    if (result.error || result.status !== 0) throw new Error(result.stderr || result.error?.message || 'Inspector failed.');
    return JSON.parse(result.stdout);
  };
  const list = invoke(['--method', 'tools/list']);
  await writeFile(join(session, 'tools.json'), JSON.stringify(list, null, 2));
  assert.equal(list.result.tools.length, 7);
  assert.ok(list.result.tools.some(tool => tool.name === 'dev_status'));
  const response = invoke(['--method', 'tools/call', '--tool-name', 'dev_status']);
  await writeFile(join(session, 'status.json'), JSON.stringify(response, null, 2));
  assert.notEqual(response.result.isError, true);
  const status = JSON.parse(response.result.content[0].text);
  const hash = async path => createHash('sha256').update(await readFile(path)).digest('hex').toUpperCase();
  assert.equal(status.build, await hash(server));
  assert.equal(status.spikeHash, await hash(spike));
  assert.equal(status.state, 'idle');
  console.log('PASS Inspector tools/list, dev_status and server/Spike hashes. No Excel operation requested.');
}
