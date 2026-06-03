/**
 * Taskpane log utility — shared across entry point and commands.
 */
const logEl = document.getElementById('log')!;

export function log(msg: string, cls?: string): void {
  const line = document.createElement('div');
  line.className = cls || '';
  line.textContent = new Date().toLocaleTimeString('en-GB', { hour12: false }) + ' ' + msg;
  logEl.appendChild(line);
  logEl.scrollTop = logEl.scrollHeight;
  if (logEl.children.length > 200) logEl.removeChild(logEl.firstChild!);
}
