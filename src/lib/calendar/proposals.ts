import type {
  CalendarEventCancellationProposal,
  CalendarEventSnapshot,
  CalendarEventUpdateProposal,
  CalendarProposalConflict,
} from "@/types/calendarProposal";
import { calendarDateTimeInstant } from "@/lib/calendar/availability";
import type { CalendarEventDetail } from "@/types/calendar";

const MAX_SUBJECT_LENGTH = 200;
const MAX_LOCATION_LENGTH = 500;
const MAX_DESCRIPTION_LENGTH = 2_000;
const MAX_SNAPSHOT_ATTENDEES = 100;
const EMAIL_PATTERN = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
const UUID_PATTERN =
  /^[0-9a-f]{8}-[0-9a-f]{4}-[1-8][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;

export function isRecord(value: unknown): value is Record<string, unknown> {
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

function validateSnapshot(value: unknown): CalendarEventSnapshot | null {
  if (!isRecord(value) || !Array.isArray(value.attendeeEmails)) return null;
  if (
    typeof value.id !== "string" || !value.id ||
    typeof value.subject !== "string" || !value.subject.trim() ||
    value.subject.length > MAX_SUBJECT_LENGTH ||
    typeof value.start !== "string" || typeof value.end !== "string" ||
    typeof value.timeZone !== "string" || !isValidTimeZone(value.timeZone) ||
    typeof value.location !== "string" || value.location.length > MAX_LOCATION_LENGTH ||
    typeof value.description !== "string" || value.description.length > MAX_DESCRIPTION_LENGTH ||
    typeof value.organizerName !== "string" || typeof value.organizerEmail !== "string" ||
    typeof value.etag !== "string" || !value.etag ||
    value.attendeeEmails.length > MAX_SNAPSHOT_ATTENDEES ||
    !value.attendeeEmails.every((email) =>
      typeof email === "string" && email.length <= 254 && EMAIL_PATTERN.test(email))
  ) return null;
  const start = new Date(value.start);
  const end = new Date(value.end);
  if (!Number.isFinite(start.getTime()) || !Number.isFinite(end.getTime()) || end <= start) {
    return null;
  }
  return {
    id: value.id,
    subject: value.subject.trim(),
    start: start.toISOString(),
    end: end.toISOString(),
    timeZone: value.timeZone,
    location: value.location.trim(),
    attendeeEmails: [...new Set(value.attendeeEmails.map((email) => email.toLowerCase()))],
    description: value.description.trim(),
    organizerName: value.organizerName.trim(),
    organizerEmail: value.organizerEmail.trim(),
    etag: value.etag,
  };
}

function validateProposed(value: unknown): CalendarEventUpdateProposal["proposed"] | null {
  if (!isRecord(value) || !Array.isArray(value.attendeeEmails)) return null;
  if (
    typeof value.subject !== "string" || !value.subject.trim() ||
    value.subject.length > MAX_SUBJECT_LENGTH ||
    typeof value.start !== "string" || typeof value.end !== "string" ||
    typeof value.timeZone !== "string" || !isValidTimeZone(value.timeZone) ||
    typeof value.location !== "string" || value.location.length > MAX_LOCATION_LENGTH ||
    typeof value.description !== "string" || value.description.length > MAX_DESCRIPTION_LENGTH ||
    value.attendeeEmails.length > MAX_SNAPSHOT_ATTENDEES ||
    !value.attendeeEmails.every((email) =>
      typeof email === "string" && email.length <= 254 && EMAIL_PATTERN.test(email))
  ) return null;
  const start = new Date(value.start);
  const end = new Date(value.end);
  if (!Number.isFinite(start.getTime()) || !Number.isFinite(end.getTime()) || end <= start ||
      end.getTime() - start.getTime() > 86_400_000) return null;
  return {
    subject: value.subject.trim(),
    start: start.toISOString(),
    end: end.toISOString(),
    timeZone: value.timeZone,
    location: value.location.trim(),
    attendeeEmails: [...new Set(value.attendeeEmails.map((email) => email.toLowerCase()))],
    description: value.description.trim(),
  };
}

export function validateConflict(value: unknown): CalendarProposalConflict | null {
  if (!isRecord(value) || typeof value.id !== "string" ||
      typeof value.subject !== "string" || typeof value.start !== "string" ||
      typeof value.end !== "string" || typeof value.isAllDay !== "boolean") return null;
  return {
    id: value.id,
    subject: value.subject,
    start: value.start,
    end: value.end,
    isAllDay: value.isAllDay,
  };
}

export function validateUpdateProposal(value: unknown): CalendarEventUpdateProposal | null {
  if (!isRecord(value) || typeof value.id !== "string" || !UUID_PATTERN.test(value.id) ||
      !Array.isArray(value.conflicts)) return null;
  const original = validateSnapshot(value.original);
  const proposed = validateProposed(value.proposed);
  const conflicts = value.conflicts.map(validateConflict);
  if (!original || !proposed || conflicts.some((conflict) => !conflict)) return null;
  const changed = proposed.subject !== original.subject || proposed.start !== original.start ||
    proposed.end !== original.end || proposed.location !== original.location ||
    proposed.description !== original.description ||
    JSON.stringify([...proposed.attendeeEmails].sort()) !==
      JSON.stringify([...original.attendeeEmails].sort());
  if (!changed) return null;
  return { id: value.id, original, proposed, conflicts: conflicts as CalendarProposalConflict[] };
}

export function validateCancellationProposal(
  value: unknown,
): CalendarEventCancellationProposal | null {
  if (!isRecord(value) || typeof value.id !== "string" || !UUID_PATTERN.test(value.id)) return null;
  const original = validateSnapshot(value.original);
  return original ? { id: value.id, original } : null;
}

export function snapshotFingerprint(snapshot: CalendarEventSnapshot) {
  return JSON.stringify({
    id: snapshot.id,
    subject: snapshot.subject,
    start: snapshot.start,
    end: snapshot.end,
    location: snapshot.location,
    attendeeEmails: [...snapshot.attendeeEmails].sort(),
    description: snapshot.description,
    etag: snapshot.etag,
  });
}

export function conflictFingerprint(conflict: CalendarProposalConflict) {
  return `${conflict.id}|${conflict.start}|${conflict.end}|${conflict.isAllDay}`;
}

export function eventSnapshot(
  event: CalendarEventDetail,
  timeZone: string,
): CalendarEventSnapshot | null {
  const start = calendarDateTimeInstant(event.start);
  const end = calendarDateTimeInstant(event.end);
  if (start === null || end === null) return null;
  return {
    id: event.id,
    subject: event.subject,
    start: new Date(start).toISOString(),
    end: new Date(end).toISOString(),
    timeZone,
    location: event.location,
    attendeeEmails: event.attendees.map((attendee) => attendee.email).filter(Boolean),
    description: event.description.slice(0, MAX_DESCRIPTION_LENGTH),
    organizerName: event.organizer.name,
    organizerEmail: event.organizer.email,
    etag: event.etag,
  };
}
