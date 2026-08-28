"use client";

import type { CalendarEventProposal } from "@/types/calendarProposal";

interface CalendarProposalCardProps {
  proposal: CalendarEventProposal;
  isCreating: boolean;
  error: string | null;
  onCreate: () => void;
  onCancel: () => void;
}

function displayDateTime(value: string, timeZone: string) {
  return new Intl.DateTimeFormat(undefined, {
    dateStyle: "medium",
    timeStyle: "short",
    timeZone,
  }).format(new Date(value));
}

export default function CalendarProposalCard({
  proposal,
  isCreating,
  error,
  onCreate,
  onCancel,
}: CalendarProposalCardProps) {
  return (
    <section className="mt-3 rounded-lg border border-blue-800 bg-blue-950/20 p-4 text-sm">
      <h3 className="font-semibold text-blue-300">Review Calendar Event</h3>
      <dl className="mt-3 grid gap-2 text-zinc-300">
        <div><dt className="text-zinc-500">Title</dt><dd>{proposal.subject}</dd></div>
        <div>
          <dt className="text-zinc-500">Start</dt>
          <dd>{displayDateTime(proposal.start, proposal.timeZone)}</dd>
        </div>
        <div>
          <dt className="text-zinc-500">End</dt>
          <dd>{displayDateTime(proposal.end, proposal.timeZone)}</dd>
        </div>
        <div><dt className="text-zinc-500">Time zone</dt><dd>{proposal.timeZone}</dd></div>
        {proposal.location ? (
          <div><dt className="text-zinc-500">Location</dt><dd>{proposal.location}</dd></div>
        ) : null}
        <div>
          <dt className="text-zinc-500">Attendees</dt>
          <dd>{proposal.attendeeEmails.length ? proposal.attendeeEmails.join(", ") : "None"}</dd>
        </div>
        {proposal.description ? (
          <div>
            <dt className="text-zinc-500">Description</dt>
            <dd className="whitespace-pre-wrap">{proposal.description}</dd>
          </div>
        ) : null}
      </dl>

      <div className={`mt-3 rounded-md border p-3 ${
        proposal.conflicts.length
          ? "border-amber-800 bg-amber-950/30 text-amber-200"
          : "border-emerald-900 bg-emerald-950/30 text-emerald-300"
      }`}>
        {proposal.conflicts.length ? (
          <>
            <p className="font-medium">Conflict detected</p>
            <ul className="mt-2 space-y-1">
              {proposal.conflicts.map((conflict) => (
                <li key={conflict.id}>
                  {conflict.subject}: {displayDateTime(conflict.start, proposal.timeZone)}–
                  {displayDateTime(conflict.end, proposal.timeZone)}
                  {conflict.isAllDay ? " (all day)" : ""}
                </li>
              ))}
            </ul>
          </>
        ) : (
          "No calendar conflicts found."
        )}
      </div>

      {error ? <p className="mt-3 text-red-400" role="alert">{error}</p> : null}

      <div className="mt-4 flex justify-end gap-2">
        <button
          type="button"
          className="rounded-md border border-zinc-700 px-3 py-2 text-zinc-300 hover:border-zinc-500 disabled:opacity-50"
          onClick={onCancel}
          disabled={isCreating}
        >
          Cancel
        </button>
        <button
          type="button"
          className="rounded-md bg-blue-600 px-3 py-2 font-medium text-white hover:bg-blue-500 disabled:opacity-50"
          onClick={onCreate}
          disabled={isCreating}
        >
          {isCreating ? "Creating..." : "Create Event"}
        </button>
      </div>
    </section>
  );
}
