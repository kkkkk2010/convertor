declare namespace NodeJS {
  type Timeout = number;
  type Signals = string;
}

declare const process: {
  argv: string[];
  env: Record<string, string | undefined>;
  exit: (code?: number) => void;
};

declare const require: any;
declare const module: any;

declare class Buffer extends Uint8Array {
  static from: (...args: unknown[]) => Buffer;
  static concat: (list: Buffer[]) => Buffer;
  static alloc: (size: number) => Buffer;
  toString: (encoding?: string) => string;
  equals: (otherBuffer: Buffer) => boolean;
}

declare const setTimeout: (
  handler: (...args: unknown[]) => void,
  timeout?: number,
  ...args: unknown[]
) => number;

declare const clearTimeout: (timeoutId?: number) => void;
