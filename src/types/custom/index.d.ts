declare module "jszip" {
  class JSZip {
    static loadAsync: (...args: unknown[]) => Promise<any>;
    file: (...args: unknown[]) => any;
    files: Record<string, any>;
    generateAsync: (...args: unknown[]) => Promise<any>;
  }
  namespace JSZip {
    type JSZipObject = any;
  }
  export default JSZip;
}

declare module "fast-xml-parser" {
  export const XMLParser: new (...args: unknown[]) => any;
}

declare module "@xmldom/xmldom" {
  export const DOMParser: new (...args: unknown[]) => any;
  export const XMLSerializer: new (...args: unknown[]) => any;
}

declare module "node:fs/promises" {
  const fs: any;
  export default fs;
}

declare module "node:fs" {
  const fs: any;
  export default fs;
}

declare module "node:path" {
  const path: any;
  export default path;
}

declare module "node:os" {
  const os: any;
  export default os;
}

declare module "node:stream" {
  export const PassThrough: any;
}

declare module "node:perf_hooks" {
  export const performance: { now: () => number };
}

declare module "node:child_process" {
  export const spawn: any;
  export const spawnSync: any;
}

declare module "node:http" {
  export type IncomingMessage = any;
  export type Server = any;
  export type IncomingHttpHeaders = Record<string, string | string[] | undefined>;
  const http: any;
  export default http;
}

declare module "node:crypto" {
  export const randomUUID: () => string;
}
