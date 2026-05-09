const ARCHIVE_KEY = 'the-earth-written-by-us-archive';

export function loadArchive() {
  try {
    const rawArchive = window.localStorage.getItem(ARCHIVE_KEY);
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
