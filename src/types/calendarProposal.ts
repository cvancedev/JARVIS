export interface CalendarProposalConflict {
  id: string;
  subject: string;
  start: string;
  end: string;
  isAllDay: boolean;
}

export interface CalendarEventProposal {
  id: string;
  subject: string;
  start: string;
  end: string;
  timeZone: string;
  location: string;
  attendeeEmails: string[];
  description: string;
  conflicts: CalendarProposalConflict[];
}

export interface CalendarEventSnapshot {
  id: string;
  subject: string;
  start: string;
  end: string;
  timeZone: string;
  location: string;
  attendeeEmails: string[];
  description: string;
  organizerName: string;
  organizerEmail: string;
  etag: string;
}

export interface CalendarEventUpdateProposal {
  id: string;
  original: CalendarEventSnapshot;
  proposed: Omit<CalendarEventSnapshot, "id" | "etag" | "organizerName" | "organizerEmail">;
  conflicts: CalendarProposalConflict[];
}

export interface CalendarEventCancellationProposal {
  id: string;
  original: CalendarEventSnapshot;
}
