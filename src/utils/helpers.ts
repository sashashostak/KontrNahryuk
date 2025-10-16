/**
 * Допоміжні функції
 * FIXED: Винесено з main.ts
 */

export function byId<T extends HTMLElement>(id: string): T | null {
  return document.getElementById(id) as T | null;
}
