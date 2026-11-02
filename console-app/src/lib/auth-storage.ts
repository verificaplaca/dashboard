export const AUTH_STORAGE_KEY = 'vp-production-auth';
type SessionStore = Pick<Storage, 'getItem' | 'setItem' | 'removeItem'>;

/** One SDK storage adapter: refreshes keep the persistence chosen at sign-in. */
export function createAuthStorage(persistent: SessionStore, tab: SessionStore) {
  let nextRemember: boolean | undefined;
  return {
    setRemember(value: boolean) { nextRemember = value; },
    getItem(key: string) { return persistent.getItem(key) ?? tab.getItem(key); },
    setItem(key: string, value: string) {
      const remember = nextRemember ?? (persistent.getItem(AUTH_STORAGE_KEY) !== null);
      const target = remember ? persistent : tab;
      const other = remember ? tab : persistent;
      target.setItem(key, value);
      other.removeItem(key);
      if (key === AUTH_STORAGE_KEY) nextRemember = undefined;
    },
    removeItem(key: string) {
      persistent.removeItem(key);
      tab.removeItem(key);
    },
  };
}
