import { useState, useEffect, useCallback } from 'react';
import { getListItems, getListItem, type ListQueryOptions } from './client';

interface QueryState<T> {
  data: T | null;
  loading: boolean;
  error: Error | null;
  refetch: () => void;
}

/**
 * Fetch a SharePoint list. Auto-fetches on mount and when query options change.
 *
 * Usage:
 *   const { data, loading, error } = useSharePointList<Property>(
 *     LIST_NAMES.Properties,
 *     { orderBy: 'fields/Title asc' }
 *   );
 */
export function useSharePointList<TItem>(
  listName: string,
  options: ListQueryOptions = {}
): QueryState<TItem[]> {
  const [data, setData] = useState<TItem[] | null>(null);
  const [loading, setLoading] = useState<boolean>(true);
  const [error, setError] = useState<Error | null>(null);
  const [refetchTrigger, setRefetchTrigger] = useState(0);

  // Memoize the options into a stable key so the effect re-runs on actual change,
  // not on every render.
  const optionsKey = JSON.stringify(options);

  useEffect(() => {
    let cancelled = false;
    setLoading(true);
    setError(null);

    getListItems<TItem>(listName, options)
      .then((items) => {
        if (!cancelled) {
          setData(items);
          setLoading(false);
        }
      })
      .catch((err) => {
        if (!cancelled) {
          setError(err instanceof Error ? err : new Error(String(err)));
          setLoading(false);
        }
      });

    return () => {
      cancelled = true;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [listName, optionsKey, refetchTrigger]);

  const refetch = useCallback(() => setRefetchTrigger((n) => n + 1), []);
  return { data, loading, error, refetch };
}

/**
 * Fetch a single SharePoint list item by ID.
 */
export function useSharePointItem<TItem>(
  listName: string,
  itemId: string | null | undefined
): QueryState<TItem> {
  const [data, setData] = useState<TItem | null>(null);
  const [loading, setLoading] = useState<boolean>(true);
  const [error, setError] = useState<Error | null>(null);
  const [refetchTrigger, setRefetchTrigger] = useState(0);

  useEffect(() => {
    if (!itemId) {
      setData(null);
      setLoading(false);
      return;
    }
    let cancelled = false;
    setLoading(true);
    setError(null);

    getListItem<TItem>(listName, itemId)
      .then((item) => {
        if (!cancelled) {
          setData(item);
          setLoading(false);
        }
      })
      .catch((err) => {
        if (!cancelled) {
          setError(err instanceof Error ? err : new Error(String(err)));
          setLoading(false);
        }
      });

    return () => {
      cancelled = true;
    };
  }, [listName, itemId, refetchTrigger]);

  const refetch = useCallback(() => setRefetchTrigger((n) => n + 1), []);
  return { data, loading, error, refetch };
}
