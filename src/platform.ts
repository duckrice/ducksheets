// File access: native dialogs and the Rust side in the app; a file picker
// and a download in a plain browser (used during development).

import { invoke } from "@tauri-apps/api/core";
import { listen } from "@tauri-apps/api/event";
import { getCurrentWindow } from "@tauri-apps/api/window";
import { open, save } from "@tauri-apps/plugin-dialog";

export const isDesktop = "__TAURI_INTERNALS__" in window;

export interface OpenedFile {
  name: string;
  /** Absolute path; missing in the browser, where files can't be written back. */
  path: string | null;
  bytes: Uint8Array;
}

export const OPEN_FILTERS = [
  { name: "Workbooks", extensions: ["xlsx", "xlsm", "csv"] },
  { name: "Excel workbook", extensions: ["xlsx", "xlsm"] },
  { name: "CSV (comma-separated values)", extensions: ["csv"] },
];

export const baseName = (path: string) => path.split(/[\\/]/).pop() ?? path;

export async function readPath(path: string): Promise<OpenedFile> {
  const buffer = await invoke<ArrayBuffer>("read_file", { path });
  return { name: baseName(path), path, bytes: new Uint8Array(buffer) };
}

export async function pickFile(): Promise<OpenedFile | null> {
  if (isDesktop) {
    const path = await open({ multiple: false, directory: false, filters: OPEN_FILTERS });
    return path ? readPath(path) : null;
  }
  return new Promise((resolve) => {
    const input = document.createElement("input");
    input.type = "file";
    input.accept = ".xlsx,.xlsm,.csv";
    input.onchange = async () => {
      const file = input.files?.[0];
      resolve(file ? { name: file.name, path: null, bytes: new Uint8Array(await file.arrayBuffer()) } : null);
    };
    input.click();
  });
}

/** Asks where to save; null if cancelled. In the browser there is nothing to ask. */
export async function pickSavePath(suggested: string): Promise<string | null> {
  if (!isDesktop) return suggested;
  const csv = /\.csv$/i.test(suggested);
  return save({
    defaultPath: suggested,
    filters: csv ? [OPEN_FILTERS[2], OPEN_FILTERS[1]] : [OPEN_FILTERS[1], OPEN_FILTERS[2]],
  });
}

export async function writePath(path: string, bytes: Uint8Array): Promise<void> {
  if (isDesktop) {
    await invoke("write_file", bytes, { headers: { "x-path": encodeURIComponent(path) } });
    return;
  }
  const url = URL.createObjectURL(new Blob([bytes as BlobPart]));
  const link = document.createElement("a");
  link.href = url;
  link.download = baseName(path);
  link.click();
  setTimeout(() => URL.revokeObjectURL(url), 10_000);
}

/** The workbooks DuckSheets was started with. */
export async function startupFiles(): Promise<string[]> {
  return isDesktop ? invoke<string[]>("startup_files") : [];
}

/** Calls `handler` with the files someone opens while DuckSheets is already running. */
export function onFilesOpened(handler: (paths: string[]) => void) {
  if (isDesktop) void listen<string[]>("open-files", (event) => handler(event.payload));
}

export function setTitle(title: string) {
  document.title = title;
  if (isDesktop) void getCurrentWindow().setTitle(title).catch(() => {});
}

/** Lets `confirm` veto closing the window (to offer saving first); it resolves to true when it may close. */
export function onCloseRequested(confirm: () => Promise<boolean>) {
  if (!isDesktop) return;
  const appWindow = getCurrentWindow();
  void appWindow.onCloseRequested(async (event) => {
    event.preventDefault();
    if (await confirm()) await appWindow.destroy();
  });
}

export function openUrl(url: string): Promise<void> {
  return invoke("open_url", { url });
}
