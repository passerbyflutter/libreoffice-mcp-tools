/**
 * Simple async mutex for serializing LibreOffice subprocess calls.
 * soffice --headless does not support concurrent execution.
 */
export class Mutex {
  private queue: Array<() => void> = [];
  private locked = false;

  async acquire(): Promise<MutexGuard> {
    if (!this.locked) {
      this.locked = true;
      return new MutexGuard(() => this.release());
    }
    return new Promise<MutexGuard>(resolve => {
      this.queue.push(() => {
        this.locked = true;
        resolve(new MutexGuard(() => this.release()));
      });
    });
  }

  private release(): void {
    const next = this.queue.shift();
    if (next) {
      next();
    } else {
      this.locked = false;
    }
  }
}

export class MutexGuard {
  private released = false;

  constructor(private readonly releaseFn: () => void) {}

  dispose(): void {
    if (!this.released) {
      this.released = true;
      this.releaseFn();
    }
  }
}
