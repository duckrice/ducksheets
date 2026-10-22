// Small pieces of interface: pop-ups under a button, menus, questions,
// asking for a name, and a note at the bottom with an Undo button.

import { h } from "./dom";

let closeOpen: (() => void) | null = null;

/** Shows `make`'s content under `anchor`; closes on Escape or a click elsewhere. */
export function popover(anchor: HTMLElement, make: (close: () => void) => HTMLElement): () => void {
  closeOpen?.();
  const box = h("div", { class: "popover" });
  const close = () => {
    if (closeOpen !== close) return;
    closeOpen = null;
    box.remove();
    document.removeEventListener("pointerdown", outside, true);
    document.removeEventListener("keydown", escape, true);
  };
  const outside = (event: Event) => {
    if (!box.contains(event.target as Node) && !anchor.contains(event.target as Node)) close();
  };
  const escape = (event: KeyboardEvent) => {
    if (event.key === "Escape") {
      event.preventDefault();
      close();
    }
  };
  closeOpen = close;
  box.append(make(close));
  document.body.append(box);
  const at = anchor.getBoundingClientRect();
  box.style.left = `${Math.max(8, Math.min(at.left, window.innerWidth - box.offsetWidth - 8))}px`;
  const below = at.bottom + 4;
  box.style.top = `${below + box.offsetHeight > window.innerHeight - 8 ? Math.max(8, at.top - box.offsetHeight - 4) : below}px`;
  document.addEventListener("pointerdown", outside, true);
  document.addEventListener("keydown", escape, true);
  return close;
}

export interface MenuItem {
  /** Colours shown before the label, as a little preview. */
  swatch?: string[];
  label: string;
  action?: () => void;
  /** A menu inside the menu. */
  items?: MenuItem[];
  danger?: boolean;
}

/** A menu of actions under `anchor` (or at the mouse, for a right-click). */
export function menu(anchor: HTMLElement, items: MenuItem[]) {
  popover(anchor, (close) => {
    const list = h("div", { class: "menu", role: "menu" });
    const fill = (entries: MenuItem[], back?: MenuItem[]) => {
      list.replaceChildren();
      if (back) list.append(h("button", { type: "button", class: "menu-item", role: "menuitem", onclick: () => fill(back) }, "← Back"));
      for (const item of entries) {
        list.append(
          h(
            "button",
            {
              type: "button",
              class: `menu-item${item.danger ? " danger" : ""}`,
              role: "menuitem",
              disabled: !item.action && !item.items?.length,
              onclick: () => {
                if (item.items) fill(item.items, entries);
                else {
                  close();
                  item.action?.();
                }
              },
            },
            item.swatch ? h("span", { class: "menu-swatch", "aria-hidden": "true" }, ...item.swatch.map((color) => h("span", { style: `background:${color}` }))) : null,
            item.label,
            item.items ? h("span", { class: "menu-more" }, "›") : null,
          ),
        );
      }
      list.querySelector<HTMLButtonElement>("button:not(:disabled)")?.focus();
    };
    fill(items);
    return list;
  });
}

export function modal(content: HTMLElement, onKey: (event: KeyboardEvent) => void): () => void {
  const backdrop = h("div", { class: "backdrop" }, content);
  const previous = document.activeElement as HTMLElement | null;
  backdrop.addEventListener("keydown", (event) => {
    event.stopPropagation();
    onKey(event);
  });
  document.body.append(backdrop);
  return () => {
    backdrop.remove();
    previous?.focus?.();
  };
}

/** Asks a yes-or-no question. */
export function ask(title: string, message: string, ok: string): Promise<boolean> {
  return new Promise((resolve) => {
    const yes = h("button", { type: "button", class: "btn primary danger" }, ok);
    const no = h("button", { type: "button", class: "btn" }, "Cancel");
    const close = modal(h("div", { class: "dialog", role: "alertdialog", "aria-label": title }, h("h2", {}, title), h("p", {}, message), h("div", { class: "dialog-buttons" }, no, yes)), (event) => {
      if (event.key === "Escape") done(false);
    });
    const done = (answer: boolean) => {
      close();
      resolve(answer);
    };
    yes.addEventListener("click", () => done(true));
    no.addEventListener("click", () => done(false));
    no.focus();
  });
}

/** Asks whether to save changes first: save, don't save, or cancel what was being done. */
export function askSave(name: string): Promise<"save" | "discard" | "cancel"> {
  return new Promise((resolve) => {
    const save = h("button", { type: "button", class: "btn primary" }, "Save");
    const discard = h("button", { type: "button", class: "btn" }, "Don't save");
    const cancel = h("button", { type: "button", class: "btn subtle" }, "Cancel");
    const close = modal(
      h("div", { class: "dialog", role: "alertdialog", "aria-label": "Save changes?" }, h("h2", {}, `Save the changes to ${name}?`), h("p", {}, "If you don't save, the changes are lost."), h("div", { class: "dialog-buttons" }, cancel, discard, save)),
      (event) => {
        if (event.key === "Escape") done("cancel");
      },
    );
    const done = (answer: "save" | "discard" | "cancel") => {
      close();
      resolve(answer);
    };
    save.addEventListener("click", () => done("save"));
    discard.addEventListener("click", () => done("discard"));
    cancel.addEventListener("click", () => done("cancel"));
    save.focus();
  });
}

/** Asks for a name; null if cancelled. */
export function askName(title: string, initial: string, ok: string): Promise<string | null> {
  return new Promise((resolve) => {
    const input = h("input", { type: "text", class: "tb-input wide", value: initial, "aria-label": title }) as HTMLInputElement;
    const yes = h("button", { type: "button", class: "btn primary" }, ok);
    const no = h("button", { type: "button", class: "btn" }, "Cancel");
    const close = modal(h("div", { class: "dialog", role: "dialog", "aria-label": title }, h("h2", {}, title), input, h("div", { class: "dialog-buttons" }, no, yes)), (event) => {
      if (event.key === "Escape") done(null);
      if (event.key === "Enter") done(input.value.trim() || null);
    });
    const done = (answer: string | null) => {
      close();
      resolve(answer);
    };
    yes.addEventListener("click", () => done(input.value.trim() || null));
    no.addEventListener("click", () => done(null));
    input.focus();
    input.select();
  });
}

let toastTimer = 0;

/** A short note at the bottom of the window, with an optional Undo. */
export function toast(text: string, undo?: () => void) {
  document.querySelector(".toast")?.remove();
  clearTimeout(toastTimer);
  const el = h("div", { class: "toast", role: "status" }, h("span", {}, text));
  if (undo)
    el.append(
      h("button", {
        type: "button",
        class: "btn subtle small",
        onclick: () => {
          el.remove();
          undo();
        },
      }, "Undo"),
    );
  document.body.append(el);
  toastTimer = window.setTimeout(() => el.remove(), 8000);
}
