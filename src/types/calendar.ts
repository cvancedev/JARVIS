export interface CalendarParticipant {
  name: string;
  email: string;
}

export interface CalendarAttendee extends CalendarParticipant {
  type: "required" | "optional" | "resource";
  response: string;
}

export interface CalendarDateTime {
  dateTime: string;
  timeZone: string;
}

export interface CalendarEvent {
  id: string;
  subject: string;
  start: CalendarDateTime;
  end: CalendarDateTime;
  location: string;
  organizer: CalendarParticipant;
  attendees: CalendarAttendee[];
  isAllDay: boolean;
  isCancelled: boolean;
  showAs: string;
}

export interface CalendarEventDetail extends CalendarEvent {
  description: string;
  eventType: "singleInstance" | "occurrence" | "exception" | "seriesMaster";
  seriesMasterId: string;
  etag: string;
}
