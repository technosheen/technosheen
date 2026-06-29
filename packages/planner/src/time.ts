export interface Interval {
  start: string;
  end: string;
}

export function toMillis(value: string): number {
  return new Date(value).getTime();
}

export function durationMinutes(interval: Interval): number {
  return Math.round((toMillis(interval.end) - toMillis(interval.start)) / 60_000);
}

export function overlaps(a: Interval, b: Interval): boolean {
  return toMillis(a.start) < toMillis(b.end) && toMillis(b.start) < toMillis(a.end);
}

export function assertNoOverlaps(intervals: Interval[]): void {
  const sorted = [...intervals].sort((a, b) => toMillis(a.start) - toMillis(b.start));
  for (let index = 1; index < sorted.length; index += 1) {
    const previous = sorted[index - 1];
    const current = sorted[index];
    if (previous && current && overlaps(previous, current)) {
      throw new Error(`Overlapping intervals: ${previous.start} and ${current.start}`);
    }
  }
}
