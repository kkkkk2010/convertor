import { AppError } from "./errors";

type ReleaseFn = () => void;

type QueueEntry = {
  requestId: string;
  resolve: (release: ReleaseFn) => void;
  reject: (error: AppError) => void;
  timeoutId: NodeJS.Timeout;
};

export type QueueOptions = {
  maxConcurrent: number;
  maxQueue: number;
  queueWaitTimeoutMs: number;
};

export class RequestQueue {
  private readonly maxConcurrent: number;
  private readonly maxQueue: number;
  private readonly queueWaitTimeoutMs: number;
  private readonly queue: QueueEntry[] = [];
  private active = 0;

  constructor(options: QueueOptions) {
    this.maxConcurrent = options.maxConcurrent;
    this.maxQueue = options.maxQueue;
    this.queueWaitTimeoutMs = options.queueWaitTimeoutMs;
  }

  async acquire(requestId: string): Promise<ReleaseFn> {
    if (this.active < this.maxConcurrent) {
      this.active += 1;
      return () => this.release();
    }
    if (this.queue.length >= this.maxQueue) {
      throw new AppError("QUEUE_FULL", "Queue is full.");
    }
    return new Promise<ReleaseFn>((resolve, reject) => {
      const timeoutId = setTimeout(() => {
        this.removeFromQueue(requestId);
        reject(new AppError("QUEUE_TIMEOUT", "Queue wait timeout exceeded."));
      }, this.queueWaitTimeoutMs);
      this.queue.push({ requestId, resolve, reject, timeoutId });
    });
  }

  cancel(requestId: string): void {
    const entryIndex = this.queue.findIndex(
      (entry) => entry.requestId === requestId,
    );
    if (entryIndex === -1) {
      return;
    }
    const [entry] = this.queue.splice(entryIndex, 1);
    clearTimeout(entry.timeoutId);
    entry.reject(new AppError("INTERNAL", "Request canceled."));
  }

  private release(): void {
    this.active = Math.max(0, this.active - 1);
    if (this.queue.length === 0) {
      return;
    }
    const entry = this.queue.shift();
    if (!entry) {
      return;
    }
    clearTimeout(entry.timeoutId);
    this.active += 1;
    entry.resolve(() => this.release());
  }

  private removeFromQueue(requestId: string): void {
    const entryIndex = this.queue.findIndex(
      (entry) => entry.requestId === requestId,
    );
    if (entryIndex === -1) {
      return;
    }
    const [entry] = this.queue.splice(entryIndex, 1);
    clearTimeout(entry.timeoutId);
  }
}
