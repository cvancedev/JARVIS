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
}
