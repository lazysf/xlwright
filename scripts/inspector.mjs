import { cp, lstat, mkdir, open, readFile, rm, writeFile } from 'node:fs/promises';
import { spawn, spawnSync } from 'node:child_process';
import { randomUUID } from 'node:crypto';
import { join, relative, resolve } from 'node:path';
import assert from 'node:assert/strict';
import { development, rejectLinks, rejectStaleBuild, repository, sha256File } from './build-guard.mjs';

const mode = process.argv[2];
if (!['--check', '--cli', '--snapshot', '--web'].includes(mode) || (mode !== '--cli' && process.argv.length !== 3))
  throw new Error('Usage: inspector.mjs --check | --snapshot | --web | --cli [Inspector CLI flags]');
if (resolve(process.cwd()) !== repository) throw new Error('Run the Inspector from the repository root.');

const sessions = join(development, 'inspector');
await rejectLinks(sessions);
await mkdir(sessions, { recursive: true });
await rejectLinks(join(development, 'check.lock'));
// The integrated check already holds this lock. Standalone snapshots must not race a build.
const buildLock = mode !== '--check' ? await open(join(development, 'check.lock'), 'a+') : undefined;
const session = join(sessions, randomUUID());
try {
  // Each DLL is compared with its own sources: an internal Core change does not always recompile the server.
  await rejectStaleBuild('XlWright.McpServer', 'src/XlWright.McpServer');
  await rejectStaleBuild('XlWright.Core', 'src/XlWright.Core');
  await mkdir(session);
  const source = join(development, 'artifacts', 'bin', 'XlWright.McpServer', 'release');
  await rejectLinks(source);
  await cp(source, join(session, 'XlWright.McpServer'), { recursive: true, filter: async path => {
    if ((await lstat(path)).isSymbolicLink()) throw new Error('Build outputs may not contain symbolic links.');
    return true;
  } });
} finally { await buildLock?.close(); }

// Repeated loop runs would otherwise accumulate build copies. Failed checks keep theirs for investigation.
async function discardSession() {
  try { await rm(session, { recursive: true, maxRetries: 5 }); }
  catch { console.error(`Inspector session retained: ${relative(repository, session)}`); }
}

const server = join(session, 'XlWright.McpServer', 'XlWright.McpServer.dll');
const config = join(session, 'config.json');
await writeFile(config, JSON.stringify({ mcpServers: { xlwright: {
  type: 'stdio', command: 'dotnet', args: [server, '--development'], cwd: repository
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
  await discardSession();
  if (result.error) throw result.error;
  process.exitCode = result.status ?? 1;
} else if (mode === '--web') {
  const child = spawn(process.execPath, [launcher, '--web', '--config', config], { cwd: repository, env, stdio: 'inherit', windowsHide: true });
  child.on('error', error => { console.error(error); process.exitCode = 1; });
  child.on('exit', code => { process.exitCode = code ?? 1; });
} else {
  const invoke = (era, args) => {
    const result = spawnSync(process.execPath, [launcher, ...cliArguments, '--protocol-era', era, ...args],
      { cwd: repository, env, encoding: 'utf8', windowsHide: true, timeout: 60_000 });
    if (result.error || result.status !== 0) throw new Error(result.stderr || result.error?.message || 'Inspector failed.');
    return JSON.parse(result.stdout);
  };
  const listTools = era => invoke(era, ['--method', 'tools/list', '--strict']).result.tools;
  const serialize = value => `${JSON.stringify(value, null, 2)}\n`;
  const toolsSnapshot = join(repository, 'tests', 'XlWright.McpServer.Tests', 'tools.snapshot.json');
  if (mode === '--snapshot') {
    await writeFile(toolsSnapshot, serialize(listTools('legacy')));
    console.log(`Updated ${relative(repository, toolsSnapshot)}. Review its diff; models read these definitions.`);
  } else {
    const expectedTools = JSON.parse(await readFile(toolsSnapshot, 'utf8'));
    for (const era of ['legacy', 'modern']) {
      const tools = listTools(era);
      await writeFile(join(session, `tools-${era}.json`), serialize(tools));
      assert.deepEqual(tools, expectedTools,
        `tools/list (${era}) differs from the snapshot. Compare ${relative(repository, join(session, `tools-${era}.json`))}; if intended, run npm run snapshot:tools.`);
      const response = invoke(era, ['--method', 'tools/call', '--tool-name', 'dev_status']);
      await writeFile(join(session, `status-${era}.json`), serialize(response));
      assert.notEqual(response.result.isError, true);
      assert.equal(JSON.parse(response.result.content[0].text).build, await sha256File(server));
    }
    console.log('PASS Inspector tools/list snapshot, dev_status and server hash on both protocol eras.');
  }
  await discardSession();
}
