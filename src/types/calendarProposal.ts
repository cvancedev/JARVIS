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
