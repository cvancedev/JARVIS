import { NextResponse } from "next/server";
import {
  eventSnapshot,
  isRecord,
  snapshotFingerprint,
  validateCancellationProposal,
} from "@/lib/calendar/proposals";
import { cancelCalendarEvent, getCalendarEvent, OutlookGraphError } from "@/lib/outlook/graph";
import { getValidAccessToken, OutlookScopeError } from "@/lib/outlook/oauth";

export const runtime = "nodejs";

export async function POST(request: Request) {
  let body: unknown;
  try {
    body = await request.json();
  } catch {
    return NextResponse.json({ error: "Invalid request body." }, { status: 400 });
  }
  const proposal = validateCancellationProposal(isRecord(body) ? body.proposal : null);
  if (!proposal) {
    return NextResponse.json({ error: "Invalid calendar cancellation proposal." }, { status: 400 });
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
        { error: "Recurring calendar events cannot be cancelled in this milestone." },
        { status: 400, headers: tokenResponse.headers },
      );
    }
    const currentSnapshot = eventSnapshot(currentEvent, proposal.original.timeZone);
    if (!currentSnapshot || snapshotFingerprint(currentSnapshot) !== snapshotFingerprint(proposal.original)) {
      return NextResponse.json(
        { error: "The event changed after review. Prepare and review a fresh cancellation.", code: "calendar_event_changed" },
        { status: 409, headers: tokenResponse.headers },
      );
    }

    await cancelCalendarEvent(accessToken, proposal.original.id, currentEvent.etag);
    return NextResponse.json({ cancelled: true }, { headers: tokenResponse.headers });
  } catch (error: unknown) {
    if (error instanceof OutlookScopeError) {
      return NextResponse.json(
        { error: "Microsoft calendar write permission is not configured.", code: "calendar_write_permission_required" },
        { status: 403, headers: tokenResponse.headers },
      );
    }
    if (error instanceof OutlookGraphError && [404, 412].includes(error.upstreamStatus)) {
      return NextResponse.json(
        { error: "The event changed or no longer exists. Prepare and review a fresh cancellation.", code: "calendar_event_changed" },
        { status: 409, headers: tokenResponse.headers },
      );
    }
    return NextResponse.json({ error: "Unable to cancel the calendar event." }, { status: 502, headers: tokenResponse.headers });
  }
}
