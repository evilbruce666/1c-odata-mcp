import { AsyncLocalStorage } from "node:async_hooks";
import { createHash } from "node:crypto";

interface WriteOperationContext {
  operationId: string;
  requestHash: string;
}

const operationContext = new AsyncLocalStorage<WriteOperationContext>();

/** Runs one MCP write-tool invocation with its preview/confirmation operation id. */
export function withWriteOperation<T>(
  operationId: string,
  requestHash: string,
  fn: () => Promise<T>,
): Promise<T> {
  return operationContext.run({ operationId, requestHash }, fn);
}

export function currentWriteOperationId(): string | undefined {
  return operationContext.getStore()?.operationId;
}

export function currentWriteRequestHash(): string | undefined {
  return operationContext.getStore()?.requestHash;
}

/** Hashes stable caller arguments, excluding the confirmation toggle and idempotency token. */
export function fingerprintWriteInput(toolName: string, input: Record<string, unknown>): string {
  const stableInput = Object.fromEntries(
    Object.entries(input).filter(([key]) => key !== "confirm" && key !== "operationId"),
  );
  return createHash("sha256")
    .update(toolName)
    .update("\n")
    .update(JSON.stringify(canonicalize(stableInput)))
    .digest("hex");
}

function canonicalize(value: unknown): unknown {
  if (Array.isArray(value)) return value.map((item) => canonicalize(item));
  if (value && typeof value === "object") {
    return Object.fromEntries(
      Object.entries(value as Record<string, unknown>)
        .sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0))
        .map(([key, nested]) => [key, canonicalize(nested)]),
    );
  }
  return value;
}
