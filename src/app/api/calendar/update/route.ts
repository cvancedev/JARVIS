import { NextResponse } from "next/server";
import { calculateAvailability, calendarDateTimeInstant } from "@/lib/calendar/availability";
import {
  conflictFingerprint,
  eventSnapshot,
  isRecord,
  snapshotFingerprint,
  validateUpdateProposal,
} from "@/lib/calendar/proposals";
import {
  getCalendarEvent,
  getCalendarEvents,
  OutlookGraphError,
  updateCalendarEvent,
} from "@/lib/outlook/graph";
import { getValidAccessToken, OutlookScopeError } from "@/lib/outlook/oauth";
import type { CalendarProposalConflict } from "@/types/calendarProposal";

export const runtime = "nodejs";

function conflictMetadata(
  conflicts: ReturnType<typeof calculateAvailability>[number]["conflicts"],
): CalendarProposalConflict[] {
  return conflicts.flatMap((event): CalendarProposalConflict[] => {
    const start = calendarDateTimeInstant(event.start);
    const end = calendarDateTimeInstant(event.end);
    return start === null || end === null ? [] : [{
      id: event.id,
      subject: event.subject,
      start: new Date(start).toISOString(),
      end: new Date(end).toISOString(),
      isAllDay: event.isAllDay,
    }];
  });
}

export async function POST(request: Request) {
  let body: unknown;
  try {
    body = await request.json();
  } catch {
    return NextResponse.json({ error: "Invalid request body." }, { status: 400 });
  }
  const proposal = validateUpdateProposal(isRecord(body) ? body.proposal : null);
  if (!proposal) {
    return NextResponse.json({ error: "Invalid calendar update proposal." }, { status: 400 });
  }

  const tokenResponse = NextResponse.json({});
  try {
    const accessToken = await getValidAccessToken(tokenResponse);
    if (!accessToken) {
      return NextResponse.json({ error: "Microsoft is not connected." }, { status: 401, headers: tokenResponse.headers });
    }
    const currentEvent = await getCalendarEvent(accessToken, proposal.original.id);
    if (currentEvent.eventType !== "singleInstance") {
      return NextResponse.json(
        { error: "Recurring calendar events cannot be updated in this milestone." },
        { status: 400, headers: tokenResponse.headers },
      );
    }
    const currentSnapshot = eventSnapshot(currentEvent, proposal.original.timeZone);
    if (!currentSnapshot || snapshotFingerprint(currentSnapshot) !== snapshotFingerprint(proposal.original)) {
      return NextResponse.json(
        { error: "The event changed after review. Prepare and review a fresh update.", code: "calendar_event_changed" },
        { status: 409, headers: tokenResponse.headers },
      );
    }

    const timeChanged = proposal.proposed.start !== proposal.original.start ||
      proposal.proposed.end !== proposal.original.end;
    let currentConflicts: CalendarProposalConflict[] = [];
    if (timeChanged) {
      const start = new Date(proposal.proposed.start);
      const end = new Date(proposal.proposed.end);
      const events = (await getCalendarEvents(accessToken, start, end, 200))
        .filter((event) => event.id !== proposal.original.id);
      const availability = calculateAvailability(events, [{ start, end }], 1)[0];
      currentConflicts = conflictMetadata(availability.conflicts);
    }
    const reviewed = proposal.conflicts.map(conflictFingerprint).sort();
    const current = currentConflicts.map(conflictFingerprint).sort();
    if (JSON.stringify(reviewed) !== JSON.stringify(current)) {
      return NextResponse.json(
        { error: "Calendar conflicts changed. Review the updated conflicts before updating.", code: "calendar_conflicts_changed", conflicts: currentConflicts },
        { status: 409, headers: tokenResponse.headers },
      );
    }

    await updateCalendarEvent(accessToken, proposal.original.id, currentEvent.etag, proposal);
    return NextResponse.json({ updated: true }, { headers: tokenResponse.headers });
  } catch (error: unknown) {
    if (error instanceof OutlookScopeError) {
      return NextResponse.json(
        { error: "Microsoft calendar write permission is not configured.", code: "calendar_write_permission_required" },
        { status: 403, headers: tokenResponse.headers },
      );
    }
    if (error instanceof OutlookGraphError && [404, 412].includes(error.upstreamStatus)) {
      return NextResponse.json(
        { error: "The event changed or no longer exists. Prepare and review a fresh update.", code: "calendar_event_changed" },
        { status: 409, headers: tokenResponse.headers },
      );
    }
    return NextResponse.json({ error: "Unable to update the calendar event." }, { status: 502, headers: tokenResponse.headers });
  }
}
