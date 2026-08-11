/**
 * Client for the PowerShell OneNote COM bridge.
 *
 * Spawns bridge/onenote_bridge.ps1 once and speaks JSON-lines RPC over its
 * stdio. PowerShell rather than Python: with x64 Click-to-Run Office the
 * OneNote typelib is only registered under the Win32 key, which breaks
 * pywin32 dispatch from 64-bit Python; .NET's COM binder is unaffected.
 */

import { spawn, type ChildProcessWithoutNullStreams } from "node:child_process";
import { createInterface } from "node:readline";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const BRIDGE_SCRIPT = join(
  dirname(fileURLToPath(import.meta.url)),
  "..",
  "bridge",
  "onenote_bridge.ps1",
);

const CALL_TIMEOUT_MS = 60_000;

interface Pending {
  resolve: (value: unknown) => void;
  reject: (err: Error) => void;
  timer: NodeJS.Timeout;
}

export class OneNoteBridge {
  private child: ChildProcessWithoutNullStreams | null = null;
  private nextId = 1;
  private pending = new Map<number, Pending>();

  start(): void {
    if (this.child) return;
    this.child = spawn(
      "powershell.exe",
      ["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", BRIDGE_SCRIPT],
      { stdio: ["pipe", "pipe", "pipe"], windowsHide: true },
    );
    const rl = createInterface({ input: this.child.stdout });
    rl.on("line", (line) => this.onLine(line));
    this.child.on("exit", (code) => {
      const err = new Error(`OneNote bridge exited (code ${code ?? "unknown"})`);
      for (const p of this.pending.values()) {
        clearTimeout(p.timer);
        p.reject(err);
      }
      this.pending.clear();
      this.child = null;
    });
  }

  private onLine(line: string): void {
    let msg: { id?: number; ok?: boolean; result?: unknown; error?: string };
    try {
      msg = JSON.parse(line);
    } catch {
      return; // stray PowerShell noise; responses are always valid JSON lines
    }
    if (typeof msg.id !== "number") return;
    const p = this.pending.get(msg.id);
    if (!p) return;
    this.pending.delete(msg.id);
    clearTimeout(p.timer);
    if (msg.ok) p.resolve(msg.result);
    else p.reject(new Error(msg.error ?? "OneNote bridge error"));
  }

  call(op: string, args: Record<string, unknown> = {}): Promise<unknown> {
    this.start();
    const child = this.child;
    if (!child) return Promise.reject(new Error("OneNote bridge failed to start"));
    const id = this.nextId++;
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        this.pending.delete(id);
        reject(new Error(`OneNote bridge call '${op}' timed out after ${CALL_TIMEOUT_MS}ms`));
      }, CALL_TIMEOUT_MS);
      this.pending.set(id, { resolve, reject, timer });
      child.stdin.write(JSON.stringify({ id, op, args }) + "\n");
    });
  }

  stop(): void {
    if (!this.child) return;
    this.child.stdin.end();
    this.child.kill();
    this.child = null;
  }
}
