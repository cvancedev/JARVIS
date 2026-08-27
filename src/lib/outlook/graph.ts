import type { OutlookMessage } from "@/types/outlook";

export class OutlookGraphError extends Error {
  constructor(
    public readonly upstreamStatus: number,
    public readonly graphErrorCode: string,
    public readonly graphErrorMessage: string,
    public readonly requestId: string | null,
  ) {
    super("Microsoft Graph inbox request failed.");
    this.name = "OutlookGraphError";
  }
}

function sanitizedString(value: unknown, fallback: string, maxLength: number) {
  return typeof value === "string" && value.trim()
    ? value.replace(/[\r\n]+/g, " ").trim().slice(0, maxLength)
    : fallback;
}

async function createGraphError(response: Response) {
  let body: unknown = null;
  try {
    body = await response.json();
  } catch {
    // Some upstream failures do not include a JSON Graph error body.
  }

  const error =
    typeof body === "object" && body !== null && "error" in body &&
    typeof body.error === "object" && body.error !== null
      ? (body.error as Record<string, unknown>)
      : null;
  const innerError =
    error && typeof error.innerError === "object" && error.innerError !== null
      ? (error.innerError as Record<string, unknown>)
      : null;
  const graphErrorCode = sanitizedString(error?.code, "unknown_error", 100);
  const graphErrorMessage = sanitizedString(
    error?.message,
    "Microsoft Graph returned an error without a message.",
    500,
  );
  const requestIdValue =
    response.headers.get("request-id") ??
    response.headers.get("client-request-id") ??
    innerError?.["request-id"] ??
    innerError?.requestId;
  const requestId =
    typeof requestIdValue === "string"
      ? sanitizedString(requestIdValue, "", 200) || null
      : null;

  console.error("Microsoft Graph inbox request failed.", {
    upstreamStatus: response.status,
    graphErrorCode,
    graphErrorMessage,
    requestId,
  });

  return new OutlookGraphError(
    response.status,
    graphErrorCode,
    graphErrorMessage,
    requestId,
  );
}

export async function getRecentInboxMessages(accessToken: string) {
  const query = new URLSearchParams({
    "$top": "10",
    "$select": "id,from,subject,receivedDateTime,isRead,bodyPreview",
    "$orderby": "receivedDateTime desc",
  });
  const response = await fetch(
    `https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?${query}`,
    { headers: { Authorization: `Bearer ${accessToken}` }, cache: "no-store" },
  );
  if (!response.ok) throw await createGraphError(response);

  const body: unknown = await response.json();
  if (typeof body !== "object" || body === null || !("value" in body) || !Array.isArray(body.value)) {
    throw new Error("Invalid Microsoft Graph response.");
  }

  return body.value.flatMap((item): OutlookMessage[] => {
    if (typeof item !== "object" || item === null) return [];
    const message = item as Record<string, unknown>;
    if (
      typeof message.id !== "string" ||
      typeof message.subject !== "string" ||
      typeof message.receivedDateTime !== "string" ||
      typeof message.isRead !== "boolean" ||
      typeof message.bodyPreview !== "string"
    ) return [];

    const from = message.from as Record<string, unknown> | undefined;
    const address = from?.emailAddress as Record<string, unknown> | undefined;
    return [{
      id: message.id,
      senderName: typeof address?.name === "string" ? address.name : "Unknown sender",
      senderEmail: typeof address?.address === "string" ? address.address : "",
      subject: message.subject || "(No subject)",
      receivedAt: message.receivedDateTime,
      isRead: message.isRead,
      preview: message.bodyPreview.slice(0, 500),
    }];
  });
}
