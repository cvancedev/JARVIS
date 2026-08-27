export type OutlookConnectionState =
  | "not_connected"
  | "connecting"
  | "connected"
  | "authentication_error";

export interface OutlookMessage {
  id: string;
  senderName: string;
  senderEmail: string;
  subject: string;
  receivedAt: string;
  isRead: boolean;
  preview: string;
}
