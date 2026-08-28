"use client";

import type {
  CalendarEventCancellationProposal,
  CalendarEventSnapshot,
  CalendarEventUpdateProposal,
} from "@/types/calendarProposal";

interface Props {
  update: CalendarEventUpdateProposal | null;
  cancellation: CalendarEventCancellationProposal | null;
  isSubmitting: boolean;
  error: string | null;
  onUpdate: () => void;
  onCancelEvent: () => void;
  onDismiss: () => void;
}

function displayDateTime(value: string, timeZone: string) {
  return new Intl.DateTimeFormat(undefined, {
    dateStyle: "medium",
    timeStyle: "short",
    timeZone,
  }).format(new Date(value));
}

function EventDetails({ event }: { event: CalendarEventSnapshot }) {
  return (
    <dl className="mt-2 grid gap-2 text-zinc-300">
      <div><dt className="text-zinc-500">Title</dt><dd>{event.subject}</dd></div>
      <div><dt className="text-zinc-500">Start</dt><dd>{displayDateTime(event.start, event.timeZone)}</dd></div>
      <div><dt className="text-zinc-500">End</dt><dd>{displayDateTime(event.end, event.timeZone)}</dd></div>
      <div><dt className="text-zinc-500">Time zone</dt><dd>{event.timeZone}</dd></div>
      {event.location ? <div><dt className="text-zinc-500">Location</dt><dd>{event.location}</dd></div> : null}
      {event.organizerName || event.organizerEmail ? (
        <div><dt className="text-zinc-500">Organizer</dt><dd>{event.organizerName || event.organizerEmail}</dd></div>
      ) : null}
      <div><dt className="text-zinc-500">Attendees</dt><dd>{event.attendeeEmails.length ? event.attendeeEmails.join(", ") : "None"}</dd></div>
      {event.description ? <div><dt className="text-zinc-500">Description</dt><dd className="whitespace-pre-wrap">{event.description}</dd></div> : null}
    </dl>
  );
}

export default function CalendarChangeProposalCard({
  update,
  cancellation,
  isSubmitting,
  error,
  onUpdate,
  onCancelEvent,
  onDismiss,
}: Props) {
  if (!update && !cancellation) return null;
  const original = update?.original ?? cancellation?.original;
  if (!original) return null;

  return (
    <section className={`mt-3 rounded-lg border p-4 text-sm ${
      cancellation ? "border-red-900 bg-red-950/20" : "border-blue-800 bg-blue-950/20"
    }`}>
      <h3 className={`font-semibold ${cancellation ? "text-red-300" : "text-blue-300"}`}>
        {cancellation ? "Review Calendar Cancellation" : "Review Calendar Update"}
      </h3>
      {update ? (
        <div className="mt-3 grid gap-4 md:grid-cols-2">
          <div className="rounded-md border border-zinc-800 p-3">
            <h4 className="font-medium text-zinc-400">Original</h4>
            <EventDetails event={update.original} />
          </div>
          <div className="rounded-md border border-blue-900 p-3">
            <h4 className="font-medium text-blue-300">Proposed</h4>
            <EventDetails event={{
              ...update.original,
              ...update.proposed,
            }} />
          </div>
        </div>
      ) : (
        <EventDetails event={original} />
      )}

      {update ? (
        <div className={`mt-3 rounded-md border p-3 ${
          update.conflicts.length
            ? "border-amber-800 bg-amber-950/30 text-amber-200"
            : "border-emerald-900 bg-emerald-950/30 text-emerald-300"
        }`}>
          {update.conflicts.length ? (
            <>
              <p className="font-medium">Conflict detected</p>
              <ul className="mt-2 space-y-1">
                {update.conflicts.map((conflict) => (
                  <li key={conflict.id}>
                    {conflict.subject}: {displayDateTime(conflict.start, update.proposed.timeZone)}–
                    {displayDateTime(conflict.end, update.proposed.timeZone)}
                    {conflict.isAllDay ? " (all day)" : ""}
                  </li>
                ))}
              </ul>
            </>
          ) : "No calendar conflicts found."}
        </div>
      ) : (
        <p className="mt-3 rounded-md border border-red-900 bg-red-950/30 p-3 text-red-200">
          This permanently removes the selected calendar event.
        </p>
      )}

      {error ? <p className="mt-3 text-red-400" role="alert">{error}</p> : null}
      <div className="mt-4 flex justify-end gap-2">
        <button type="button" onClick={onDismiss} disabled={isSubmitting}
          className="rounded-md border border-zinc-700 px-3 py-2 text-zinc-300 hover:border-zinc-500 disabled:opacity-50">
          {cancellation ? "Keep Event" : "Dismiss"}
        </button>
        <button type="button" onClick={update ? onUpdate : onCancelEvent} disabled={isSubmitting}
          className={`rounded-md px-3 py-2 font-medium text-white disabled:opacity-50 ${
            cancellation ? "bg-red-700 hover:bg-red-600" : "bg-blue-600 hover:bg-blue-500"
          }`}>
          {isSubmitting ? (cancellation ? "Cancelling..." : "Updating...") : (cancellation ? "Cancel Event" : "Update Event")}
        </button>
      </div>
    </section>
  );
}
