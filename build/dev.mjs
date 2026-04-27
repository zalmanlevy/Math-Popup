import { spawn } from 'node:child_process';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = dirname(fileURLToPath(import.meta.url));
const root = join(__dirname, '..');

// One-shot build, then launch electron. (For incremental dev use `npm run build:watch` in another terminal.)
const build = spawn(process.execPath, [join(__dirname, 'build.mjs')], { stdio: 'inherit' });
build.on('exit', code => {
  if (code !== 0) process.exit(code ?? 1);
  const electronBin = process.platform === 'win32'
    ? join(root, 'node_modules', 'electron', 'dist', 'electron.exe')
    : join(root, 'node_modules', 'electron', 'dist', 'electron');
  const child = spawn(electronBin, ['.'], { stdio: 'inherit', cwd: root });
  child.on('exit', c => process.exit(c ?? 0));
});
