#!/usr/bin/env node
'use strict';

const fs = require('fs');
const path = require('path');
const childProcess = require('child_process');

function parseArgs(argv) {
  const args = [...argv];
  const out = { _: [] };
  while (args.length) {
    const token = args.shift();
    if (token.startsWith('--')) {
      const key = token.slice(2);
      const value = args[0] && !args[0].startsWith('--') ? args.shift() : true;
      out[key] = value;
    } else {
      out._.push(token);
    }
  }
  return out;
}

function boolArg(args, key, defaultValue) {
  if (args[key] === undefined) return defaultValue;
  if (typeof args[key] === 'boolean') return args[key];
  const value = String(args[key]).toLowerCase();
  if (value === 'false' || value === '0' || value === 'no') return false;
  return true;
}

function fileExists(filePath) {
  try {
    fs.accessSync(filePath, fs.constants.F_OK);
    return true;
  } catch (err) {
    return false;
  }
}

function ensureDir(dirPath) {
  fs.mkdirSync(dirPath, { recursive: true });
}

function hasUnescapedAmp(value) {
  for (let i = 0; i < value.length; i += 1) {
    if (value[i] === '&') {
      const tail = value.slice(i);
      if (
        tail.startsWith('&amp;') ||
        tail.startsWith('&lt;') ||
        tail.startsWith('&gt;') ||
        tail.startsWith('&quot;') ||
        tail.startsWith('&apos;')
      ) {
        continue;
      }
      return true;
    }
  }
  return false;
}

function validateManifestText(xmlText) {
  const issues = [];
  if (!xmlText.includes('<OfficeApp')) issues.push('missing <OfficeApp>');
  if (!xmlText.includes('<Id>')) issues.push('missing <Id>');
  if (!xmlText.includes('SourceLocation')) issues.push('missing SourceLocation');
  const attrRegex = /DefaultValue="([^"]*)"/g;
  let match;
  while ((match = attrRegex.exec(xmlText))) {
    if (hasUnescapedAmp(match[1])) {
      issues.push(`unescaped & in DefaultValue: ${match[1].slice(0, 120)}`);
      break;
    }
  }
  return issues;
}

function validateManifestFile(filePath) {
  const xmlText = fs.readFileSync(filePath, 'utf8');
  const issues = validateManifestText(xmlText);
  if (issues.length) {
    throw new Error(`Manifest invalid: ${issues.join('; ')}`);
  }
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function waitForHealth(url, attempts, delayMs) {
  const maxAttempts = attempts || 12;
  const waitMs = delayMs || 400;
  for (let i = 0; i < maxAttempts; i += 1) {
    try {
      const res = await fetch(url);
      if (res.ok) return true;
    } catch (err) {
      // ignore
    }
    await sleep(waitMs);
  }
  throw new Error(`server not healthy at ${url}`);
}

function killPidFile(pidPath) {
  if (!pidPath || !fileExists(pidPath)) return;
  const raw = fs.readFileSync(pidPath, 'utf8').trim();
  const pid = Number(raw);
  if (!Number.isFinite(pid)) return;
  try {
    process.kill(pid);
  } catch (err) {
    // ignore
  }
}

function killPort(port) {
  if (!port) return;
  try {
    const raw = childProcess.execFileSync('lsof', ['-nP', `-iTCP:${port}`, '-sTCP:LISTEN', '-t'], { encoding: 'utf8' }).trim();
    if (!raw) return;
    raw.split(/\r?\n/).map((line) => Number(line.trim())).filter(Number.isFinite).forEach((pid) => {
      try {
        process.kill(pid);
      } catch (err) {
        // ignore
      }
    });
  } catch (err) {
    // lsof not available or no listeners
  }
}

function startServerDetached(serverCmd, pidPath, logPath) {
  if (!serverCmd) return null;
  ensureDir(path.dirname(pidPath));
  ensureDir(path.dirname(logPath));
  const out = fs.openSync(logPath, 'a');
  const child = childProcess.spawn(serverCmd, {
    shell: true,
    detached: true,
    stdio: ['ignore', out, out],
  });
  child.unref();
  fs.writeFileSync(pidPath, String(child.pid));
  return child.pid;
}

function resolveAppTargets(appValue) {
  const app = (appValue || 'word').toLowerCase();
  const targets = [];
  if (app === 'word' || app === 'both') targets.push('Microsoft Word');
  if (app === 'excel' || app === 'both') targets.push('Microsoft Excel');
  if (!targets.length) throw new Error('app must be word, excel, or both');
  return targets;
}

function restartApps(appValue) {
  const apps = resolveAppTargets(appValue);
  apps.forEach((app) => {
    try {
      childProcess.execFileSync('osascript', ['-e', `quit app \"${app}\"`], { stdio: 'ignore' });
    } catch (err) {
      // ignore if not running
    }
    childProcess.execFileSync('open', ['-a', app]);
  });
}

function findTrashBin() {
  try {
    childProcess.execFileSync('which', ['trash'], { stdio: 'ignore' });
    return 'trash';
  } catch (err) {
    return null;
  }
}

function trashPaths(pathsToTrash) {
  const bin = findTrashBin();
  if (!bin) {
    throw new Error('trash command not found (install or skip --clear-cache)');
  }
  pathsToTrash.forEach((target) => {
    if (fileExists(target)) {
      childProcess.execFileSync(bin, [target], { stdio: 'ignore' });
    }
  });
}

function clearOfficeCache(appValue) {
  if (process.platform !== 'darwin') return;
  const home = process.env.HOME || '';
  if (!home) throw new Error('HOME not set');
  const apps = resolveAppTargets(appValue);
  apps.forEach((app) => {
    if (app === 'Microsoft Word') {
      trashPaths([
        path.join(home, 'Library', 'Containers', 'com.microsoft.Word', 'Data', 'Library', 'Caches'),
        path.join(home, 'Library', 'Containers', 'com.microsoft.Word', 'Data', 'Library', 'WebKit'),
      ]);
    }
    if (app === 'Microsoft Excel') {
      trashPaths([
        path.join(home, 'Library', 'Containers', 'com.microsoft.Excel', 'Data', 'Library', 'Caches'),
        path.join(home, 'Library', 'Containers', 'com.microsoft.Excel', 'Data', 'Library', 'WebKit'),
      ]);
    }
  });
}

function sideloadMacManifest(manifestPath, appValue) {
  const home = process.env.HOME || '';
  if (!home) throw new Error('HOME not set');
  const targets = [];
  if (appValue === 'word' || appValue === 'both') {
    targets.push(path.join(home, 'Library', 'Containers', 'com.microsoft.Word', 'Data', 'Documents', 'wef'));
  }
  if (appValue === 'excel' || appValue === 'both') {
    targets.push(path.join(home, 'Library', 'Containers', 'com.microsoft.Excel', 'Data', 'Documents', 'wef'));
  }
  if (!targets.length) throw new Error('app must be word, excel, or both');
  targets.forEach((dir) => {
    ensureDir(dir);
    fs.copyFileSync(manifestPath, path.join(dir, path.basename(manifestPath)));
  });
}

function usage() {
  return [
    'word-addin-dev (deterministic dev loop for Office add-ins)',
    '',
    'Usage:',
    '  node word-addin-dev.js --manifest ./manifest.xml --server-cmd "node dist/your-app.js serve --port 4321"',
    '',
    'Options:',
    '  --manifest <path>        Manifest XML to validate/sideload',
    '  --server-cmd <command>   Server command to run (quoted)',
    '  --health-url <url>       Health check URL (default: http://localhost:4321/api/health)',
    '  --app word|excel|both    Target app (default: word)',
    '  --port <n>               Port used for kill-port (default: 4321)',
    '  --pid-file <path>        PID file for server (default: .cache/verify/server.pid)',
    '  --log-file <path>        Log file for server (default: .cache/verify/server.log)',
    '  --sideload true|false    Copy manifest to wef (default: true if manifest provided)',
    '  --restart true|false     Restart app after sideload (default: true if sideload true)',
    '  --clear-cache true|false Clear Office caches (default: false)',
    '  --validate true|false    Validate manifest before sideload (default: true)',
    '  --health true|false      Wait for health check (default: true)',
    '  --kill-port true|false   Kill any listener on port (default: true)',
  ].join('\n');
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  if (args._[0] === 'help' || args.help) {
    process.stdout.write(`${usage()}\n`);
    return;
  }

  const manifestPath = args.manifest ? path.resolve(String(args.manifest)) : null;
  const appValue = (args.app || 'word').toLowerCase();
  const port = Number(args.port || 4321);
  const pidPath = args['pid-file'] || path.join(process.cwd(), '.cache', 'verify', 'server.pid');
  const logPath = args['log-file'] || path.join(process.cwd(), '.cache', 'verify', 'server.log');
  const serverCmd = args['server-cmd'] ? String(args['server-cmd']) : null;
  const healthUrl = args['health-url'] || `http://localhost:${port}/api/health`;

  const sideload = boolArg(args, 'sideload', Boolean(manifestPath));
  const restart = boolArg(args, 'restart', sideload);
  const clearCache = boolArg(args, 'clear-cache', false);
  const validate = boolArg(args, 'validate', true);
  const health = boolArg(args, 'health', true);
  const killPortFlag = boolArg(args, 'kill-port', true);

  if (manifestPath && validate) {
    validateManifestFile(manifestPath);
  }

  if (clearCache) {
    clearOfficeCache(appValue);
  }

  if (sideload) {
    if (!manifestPath) throw new Error('manifest is required when sideload=true');
    sideloadMacManifest(manifestPath, appValue);
  }

  if (killPortFlag) {
    killPort(port);
  } else {
    killPidFile(pidPath);
  }

  const pid = startServerDetached(serverCmd, pidPath, logPath);

  if (health) {
    await waitForHealth(healthUrl, 12, 400);
  }

  if (restart) {
    restartApps(appValue);
  }

  process.stdout.write(`ok (pid=${pid || 'none'} port=${port})\n`);
}

main().catch((err) => {
  process.stderr.write(`${err && err.message ? err.message : err}\n`);
  process.exit(1);
});
