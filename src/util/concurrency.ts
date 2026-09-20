/**
 * Run `fn` over `items` with at most `limit` promises in flight.
 *
 * OneNote reporting is request-per-section, so a whole-account scan is 170+
 * round trips. Sequential is painfully slow; an unbounded `Promise.all` earns a
 * 429. This keeps throughput useful while staying polite — and `graphRequest`
 * already retries 429/5xx with backoff if we still trip the limiter.
 */
export const mapWithConcurrency = async <T, R>(
  items: readonly T[],
  limit: number,
  fn: (item: T, index: number) => Promise<R>,
): Promise<R[]> => {
  const size = Math.max(1, Math.floor(limit));
  const results = new Array<R>(items.length);
  let cursor = 0;

  const worker = async (): Promise<void> => {
    for (;;) {
      const index = cursor;
      cursor += 1;
      if (index >= items.length) return;
      results[index] = await fn(items[index] as T, index);
    }
  };

  const workers = Array.from({ length: Math.min(size, items.length) }, worker);
  await Promise.all(workers);
  return results;
};

/** Default fan-out for Graph scans. */
export const DEFAULT_CONCURRENCY = 5;
