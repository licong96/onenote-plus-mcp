import { describe, expect, it } from 'vitest';
import { DEFAULT_CONCURRENCY, mapWithConcurrency } from '@/util/concurrency.js';

describe('mapWithConcurrency', () => {
  it('returns results in input order regardless of completion order', async () => {
    const slowFirst = async (n: number): Promise<number> => {
      await new Promise((resolve) => setTimeout(resolve, (5 - n) * 4));
      return n * 2;
    };
    const result = await mapWithConcurrency([1, 2, 3, 4, 5], 3, slowFirst);
    expect(result).toEqual([2, 4, 6, 8, 10]);
  });

  it('never runs more than `limit` callbacks at once', async () => {
    let active = 0;
    let peak = 0;

    await mapWithConcurrency([1, 2, 3, 4, 5, 6, 7, 8, 9, 10], 3, async () => {
      active += 1;
      peak = Math.max(peak, active);
      await new Promise((resolve) => setTimeout(resolve, 1));
      active -= 1;
    });

    expect(peak).toBe(3);
  });

  it('passes the item and its index to the callback', async () => {
    const seen: Array<[string, number]> = [];
    await mapWithConcurrency(['a', 'b', 'c'], 2, async (item, index) => {
      seen.push([item, index]);
      return index;
    });

    expect(seen.sort((left, right) => left[1] - right[1])).toEqual([
      ['a', 0],
      ['b', 1],
      ['c', 2],
    ]);
  });

  it('handles an empty input list', async () => {
    const callback = async (): Promise<number> => 1;
    await expect(mapWithConcurrency([], 4, callback)).resolves.toEqual([]);
  });

  it('falls back to serial execution for a limit below 1', async () => {
    let active = 0;
    let peak = 0;

    await mapWithConcurrency([1, 2, 3], 0, async () => {
      active += 1;
      peak = Math.max(peak, active);
      await new Promise((resolve) => setTimeout(resolve, 1));
      active -= 1;
    });

    expect(peak).toBe(1);
  });

  it('does not spawn more workers than there are items', async () => {
    let started = 0;
    await mapWithConcurrency([1, 2], 8, async (n) => {
      started += 1;
      return n;
    });
    expect(started).toBe(2);
  });

  it('propagates a callback rejection', async () => {
    await expect(
      mapWithConcurrency([1, 2, 3], 2, async (n) => {
        if (n === 2) throw new Error('boom');
        return n;
      }),
    ).rejects.toThrow('boom');
  });

  it('exposes a default concurrency of 5', () => {
    expect(DEFAULT_CONCURRENCY).toBe(5);
  });
});
