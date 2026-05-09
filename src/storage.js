const ARCHIVE_KEY = 'the-earth-written-by-us-archive';
const LEGACY_ARCHIVE_KEYS = ['earth-at-the-threshold-archive', 'a-website-that-tips-archive'];

export function loadArchive() {
  try {
    const rawArchive = window.localStorage.getItem(ARCHIVE_KEY) ?? LEGACY_ARCHIVE_KEYS
      .map((key) => window.localStorage.getItem(key))
      .find(Boolean);
    if (rawArchive && !window.localStorage.getItem(ARCHIVE_KEY)) {
      window.localStorage.setItem(ARCHIVE_KEY, rawArchive);
    }
    return {
      available: true,
      entries: rawArchive ? JSON.parse(rawArchive) : [],
      error: null
    };
  } catch (error) {
    return {
      available: false,
      entries: [],
      error
    };
  }
}

export function saveArchiveEntry(entry) {
  try {
    const current = loadArchive().entries;
    const next = [entry, ...current].slice(0, 20);
    window.localStorage.setItem(ARCHIVE_KEY, JSON.stringify(next));
    return {
      available: true,
      entries: next,
      error: null
    };
  } catch (error) {
    return {
      available: false,
      entries: [],
      error
    };
  }
}

export function clearArchiveEntries() {
  try {
    window.localStorage.removeItem(ARCHIVE_KEY);
    LEGACY_ARCHIVE_KEYS.forEach((key) => window.localStorage.removeItem(key));
    return {
      available: true,
      entries: [],
      error: null
    };
  } catch (error) {
    return {
      available: false,
      entries: [],
      error
    };
  }
}
