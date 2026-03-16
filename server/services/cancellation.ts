/**
 * Simple in-memory cancellation registry.
 * Migration loops call isCancelled(itemId) at every iteration boundary.
 * When true they throw CancelledError which the outer catch handles.
 */

const cancelledItems = new Set<number>();

export function requestCancellation(itemId: number): void {
  cancelledItems.add(itemId);
}

export function isCancelled(itemId: number): boolean {
  return cancelledItems.has(itemId);
}

export function clearCancellation(itemId: number): void {
  cancelledItems.delete(itemId);
}

export class CancelledError extends Error {
  constructor(itemId: number) {
    super(`Migration cancelled by user (item ${itemId})`);
    this.name = 'CancelledError';
  }
}

export function checkCancellation(itemId: number): void {
  if (isCancelled(itemId)) {
    clearCancellation(itemId);
    throw new CancelledError(itemId);
  }
}
