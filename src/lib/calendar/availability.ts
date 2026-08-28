import type { CalendarDateTime, CalendarEvent } from "@/types/calendar";

export interface AvailabilityWindow {
  start: Date;
  end: Date;
}

interface BusyInterval {
  start: number;
  end: number;
}

export function calendarDateTimeInstant(value: CalendarDateTime) {
  const milliseconds = value.dateTime.replace(/\.(\d{3})\d+/, ".$1");
  const hasOffset = /(?:Z|[+-]\d{2}:\d{2})$/i.test(milliseconds);
  const instant = Date.parse(
    value.timeZone.toUpperCase() === "UTC" && !hasOffset
      ? `${milliseconds}Z`
      : milliseconds,
  );
  return Number.isFinite(instant) ? instant : null;
}

function eventMetadata(event: CalendarEvent) {
  return {
    id: event.id,
    subject: event.subject,
    start: event.start,
    end: event.end,
    location: event.location,
    organizer: event.organizer,
    isAllDay: event.isAllDay,
    showAs: event.showAs,
  };
}

export function calculateAvailability(
  events: CalendarEvent[],
  windows: AvailabilityWindow[],
  minimumFreeMinutes: number,
) {
  return windows.map((window) => {
    const windowStart = window.start.getTime();
    const windowEnd = window.end.getTime();
    const conflicts = events.flatMap((event) => {
      if (event.isCancelled || event.showAs.toLowerCase() === "free") return [];
      const start = calendarDateTimeInstant(event.start);
      const end = calendarDateTimeInstant(event.end);
      if (start === null || end === null || end <= windowStart || start >= windowEnd) {
        return [];
      }
      return [{ event, start: Math.max(start, windowStart), end: Math.min(end, windowEnd) }];
    }).sort((a, b) => a.start - b.start || a.end - b.end);

    const mergedBusy: BusyInterval[] = [];
    for (const conflict of conflicts) {
      const previous = mergedBusy.at(-1);
      if (previous && conflict.start <= previous.end) {
        previous.end = Math.max(previous.end, conflict.end);
      } else {
        mergedBusy.push({ start: conflict.start, end: conflict.end });
      }
    }

    const freeWindows: BusyInterval[] = [];
    let cursor = windowStart;
    for (const busy of mergedBusy) {
      if (busy.start > cursor) freeWindows.push({ start: cursor, end: busy.start });
      cursor = Math.max(cursor, busy.end);
    }
    if (cursor < windowEnd) freeWindows.push({ start: cursor, end: windowEnd });

    const qualifyingFreeWindows = freeWindows.filter(
      (free) => free.end - free.start >= minimumFreeMinutes * 60_000,
    );

    return {
      requestedStart: window.start.toISOString(),
      requestedEnd: window.end.toISOString(),
      isEntireWindowFree: mergedBusy.length === 0,
      conflicts: conflicts.map(({ event }) => eventMetadata(event)),
      busyWindows: mergedBusy.map((busy) => ({
        start: new Date(busy.start).toISOString(),
        end: new Date(busy.end).toISOString(),
      })),
      freeWindows: qualifyingFreeWindows.map((free) => ({
        start: new Date(free.start).toISOString(),
        end: new Date(free.end).toISOString(),
        durationMinutes: Math.round((free.end - free.start) / 60_000),
      })),
    };
  });
}
