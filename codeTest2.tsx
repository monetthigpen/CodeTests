// ---- localStorage settings (match what getUserIdsFromSelection reads) ----
const localStorageVar = "Knowledge Services.peoplePickerIDs"; // <-- use YOUR exact key
const storageLimit = 25;
const storageExpirationDays = 30;

type PeoplePickerCache<T> = {
  savedAt: number;      // Date.now()
  expiresAt: number;    // Date.now() + N days
  items: T[];
};

function readCache<T>(key: string): T[] {
  const raw = localStorage.getItem(key);
  if (!raw) return [];

  try {
    const parsed = JSON.parse(raw) as PeoplePickerCache<T> | T[];

    // Back-compat: if you previously stored just an array
    if (Array.isArray(parsed)) return parsed;

    // New shape with expiration
    if (parsed.expiresAt && Date.now() > parsed.expiresAt) {
      localStorage.removeItem(key);
      return [];
    }

    return Array.isArray(parsed.items) ? parsed.items : [];
  } catch {
    return [];
  }
}

function writeCache<T>(key: string, items: T[], limit: number, expirationDays: number): void {
  const limited = items.slice(0, limit);
  const now = Date.now();
  const payload: PeoplePickerCache<T> = {
    savedAt: now,
    expiresAt: now + expirationDays * 24 * 60 * 60 * 1000,
    items: limited,
  };
  localStorage.setItem(key, JSON.stringify(payload));
}


