import { EventEmitter } from "events";

export const bus = new EventEmitter();

export function notifyChange(table: string, maxRowVersion: number) {
    bus.emit("change", { table, maxRowVersion, at: Date.now() });
}

export function waitForChange(table: string, since: number, timeoutMs: number): Promise<void> {
    return new Promise((resolve) => {
        const on = (e: any) => {
            if (e.table === table && e.maxRowVersion > since) {
                cleanup(); resolve();
            }
        };
        const timer = setTimeout(() => { cleanup(); resolve(); }, timeoutMs);
        const cleanup = () => { clearTimeout(timer); bus.off("change", on); };
        bus.on("change", on);
    });
}
