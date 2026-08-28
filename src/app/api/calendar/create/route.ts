import { NextResponse } from "next/server";
import {
  calculateAvailability,
  calendarDateTimeInstant,
} from "@/lib/calendar/availability";
import {
  createCalendarEvent,
  getCalendarEvents,
  OutlookGraphError,
} from "@/lib/outlook/graph";
import {
  getValidAccessToken,
  OutlookScopeError,
} from "@/lib/outlook/oauth";
import type {
  CalendarEventProposal,
  CalendarProposalConflict,
} from "@/types/calendarProposal";

export const runtime = "nodejs";

const MAX_SUBJECT_LENGTH = 200;
const MAX_LOCATION_LENGTH = 500;
const MAX_DESCRIPTION_LENGTH = 2_000;
const MAX_ATTENDEES = 20;
const EMAIL_PATTERN = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
const UUID_PATTERN =
  /^[0-9a-f]{8}-[0-9a-f]{4}-[1-8][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function isValidTimeZone(timeZone: string) {
  try {
    new Intl.DateTimeFormat("en-US", { timeZone }).format();
    return true;
  } catch {
    return false;
  }
}

function validateProposal(value: unknown): CalendarEventProposal | null {
  if (!isRecord(value)) return null;
  const attendeeEmails = Array.isArray(value.attendeeEmails)
    ? value.attendeeEmails
    : null;
  if (
    typeof value.id !== "string" ||
    !UUID_PATTERN.test(value.id) ||
    typeof value.subject !== "string" ||
    !value.subject.trim() ||
    value.subject.length > MAX_SUBJECT_LENGTH ||
    typeof value.start !== "string" ||
    typeof value.end !== "string" ||
    typeof value.timeZone !== "string" ||
    value.timeZone.length > 100 ||
    !isValidTimeZone(value.timeZone) ||
    typeof value.location !== "string" ||
    value.location.length > MAX_LOCATION_LENGTH ||
    typeof value.description !== "string" ||
    value.description.length > MAX_DESCRIPTION_LENGTH ||
    !attendeeEmails ||
    attendeeEmails.length > MAX_ATTENDEES ||
    !attendeeEmails.every(
      (email) => typeof email === "string" && email.length <= 254 && EMAIL_PATTERN.test(email),
    ) ||
    !Array.isArray(value.conflicts)
  ) return null;

  const start = new Date(value.start);
  const end = new Date(value.end);
  if (
    !Number.isFinite(start.getTime()) ||
    !Number.isFinite(end.getTime()) ||
    end <= start ||
    end.getTime() - start.getTime() > 24 * 60 * 60 * 1_000
  ) return null;

  return {
    id: value.id,
    subject: value.subject.trim(),
    start: start.toISOString(),
    end: end.toISOString(),
    timeZone: value.timeZone,
    location: value.location.trim(),
    attendeeEmails: [...new Set(attendeeEmails.map((email) => email.toLowerCase()))],
    description: value.description.trim(),
    conflicts: value.conflicts.flatMap((conflict): CalendarProposalConflict[] => {
      if (!isRecord(conflict) || typeof conflict.id !== "string") return [];
      return [{
        id: conflict.id,
        subject: typeof conflict.subject === "string" ? conflict.subject : "(Busy)",
        start: typeof conflict.start === "string" ? conflict.start : "",
        end: typeof conflict.end === "string" ? conflict.end : "",
        isAllDay: conflict.isAllDay === true,
      }];
    }),
  };
}

function conflictMetadata(
  conflicts: ReturnType<typeof calculateAvailability>[number]["conflicts"],
): CalendarProposalConflict[] {
  return conflicts.flatMap((event): CalendarProposalConflict[] => {
    const start = calendarDateTimeInstant(event.start);
    const end = calendarDateTimeInstant(event.end);
    if (start === null || end === null) return [];
    return [{
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
  const proposal = validateProposal(isRecord(body) ? body.proposal : null);
  if (!proposal) {
    return NextResponse.json({ error: "Invalid calendar event proposal." }, { status: 400 });
  }

  const tokenResponse = NextResponse.json({});
  try {
    const accessToken = await getValidAccessToken(tokenResponse);
    if (!accessToken) {
      return NextResponse.json(
        { error: "Microsoft is not connected." },
        { status: 401 },
      );
    }

    const start = new Date(proposal.start);
    const end = new Date(proposal.end);
    const events = await getCalendarEvents(accessToken, start, end, 200);
    const availability = calculateAvailability(events, [{ start, end }], 1)[0];
    const currentConflicts = conflictMetadata(availability.conflicts);
    const fingerprint = (conflict: CalendarProposalConflict) =>
      `${conflict.id}|${conflict.start}|${conflict.end}|${conflict.isAllDay}`;
    const reviewedConflicts = proposal.conflicts.map(fingerprint).sort();
    const currentConflictSet = currentConflicts.map(fingerprint).sort();
    if (JSON.stringify(reviewedConflicts) !== JSON.stringify(currentConflictSet)) {
      return NextResponse.json(
        {
          error: "Calendar conflicts changed. Review the updated proposal before creating it.",
          code: "calendar_conflicts_changed",
          conflicts: currentConflicts,
        },
        { status: 409, headers: tokenResponse.headers },
      );
    }

    const event = await createCalendarEvent(accessToken, proposal);
    return NextResponse.json(
      { created: true, eventId: event.id },
      { status: 201, headers: tokenResponse.headers },
    );
  } catch (error: unknown) {
    if (error instanceof OutlookScopeError) {
      return NextResponse.json(
        {
          error: "Microsoft calendar write permission is not configured. Reconnect after granting delegated Calendars.ReadWrite.",
          code: "calendar_write_permission_required",
          ...(process.env.NODE_ENV === "development"
            ? { diagnostic: { missingScopes: error.missingScopes } }
            : {}),
        },
        { status: 403, headers: tokenResponse.headers },
      );
    }

    if (error instanceof OutlookGraphError && error.upstreamStatus === 403) {
      return NextResponse.json(
        {
          error: "Microsoft Graph refused calendar event creation.",
          code: "calendar_create_forbidden",
          ...(process.env.NODE_ENV === "development"
            ? {
                diagnostic: {
                  upstreamStatus: error.upstreamStatus,
                  graphErrorCode: error.graphErrorCode,
                  graphErrorMessage: error.graphErrorMessage,
                  requestId: error.requestId,
                },
              }
            : {}),
        },
        { status: 403, headers: tokenResponse.headers },
      );
    }

    return NextResponse.json(
      { error: "Unable to create the calendar event." },
      { status: 502, headers: tokenResponse.headers },
    );
  }
}
