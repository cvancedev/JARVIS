import type { OutlookMessage, OutlookMessageDetail } from "@/types/outlook";

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

function decodeHtmlEntities(value: string) {
  const namedEntities: Record<string, string> = {
    amp: "&",
    apos: "'",
    gt: ">",
    lt: "<",
    nbsp: " ",
    quot: '"',
  };

  return value.replace(/&(#x?[0-9a-f]+|[a-z]+);/gi, (entity, code: string) => {
    if (!code.startsWith("#")) return namedEntities[code.toLowerCase()] ?? entity;
    const hexadecimal = code[1]?.toLowerCase() === "x";
    const point = Number.parseInt(code.slice(hexadecimal ? 2 : 1), hexadecimal ? 16 : 10);
    try {
      return Number.isFinite(point) ? String.fromCodePoint(point) : entity;
    } catch {
      return entity;
    }
  });
}

function htmlToReadableText(html: string) {
  return decodeHtmlEntities(
    html
      .replace(/<(script|style|head|noscript|svg)[^>]*>[\s\S]*?<\/\1>/gi, " ")
      .replace(/<(?:br|hr)\s*\/?>/gi, "\n")
      .replace(/<\/(?:p|div|li|tr|h[1-6])\s*>/gi, "\n")
      .replace(/<[^>]+>/g, " "),
  )
    .replace(/\r/g, "")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n[ \t]+/g, "\n")
    .replace(/[ \t]{2,}/g, " ")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function readableBody(content: string, contentType: string) {
  if (contentType.toLowerCase() === "html") return htmlToReadableText(content);
  return content
    .replace(/\0/g, "")
    .replace(/\r\n?/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

async function parseMessageList(response: Response) {
  if (!response.ok) throw await createGraphError(response);

  const body: unknown = await response.json();
  if (
    typeof body !== "object" ||
    body === null ||
    !("value" in body) ||
    !Array.isArray(body.value)
  ) {
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
  }).sort(
    (a, b) => Date.parse(b.receivedAt) - Date.parse(a.receivedAt),
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
  return parseMessageList(response);
}

export async function searchOutlookMessages(
  accessToken: string,
  searchQuery: string,
) {
  const escapedQuery = searchQuery
    .replace(/\\/g, "\\\\")
    .replace(/"/g, '\\"');
  const query = new URLSearchParams({
    "$search": `"${escapedQuery}"`,
    "$top": "20",
    "$select": "id,from,subject,receivedDateTime,isRead,bodyPreview",
  });
  const response = await fetch(
    `https://graph.microsoft.com/v1.0/me/messages?${query}`,
    { headers: { Authorization: `Bearer ${accessToken}` }, cache: "no-store" },
  );
  return parseMessageList(response);
}

export async function getOutlookMessage(
  accessToken: string,
  messageId: string,
): Promise<OutlookMessageDetail> {
  const query = new URLSearchParams({
    "$select": "id,from,toRecipients,subject,receivedDateTime,body",
  });
  const response = await fetch(
    `https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(messageId)}?${query}`,
    {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        Prefer: 'outlook.body-content-type="text"',
      },
      cache: "no-store",
    },
  );
  if (!response.ok) throw await createGraphError(response);

  const body: unknown = await response.json();
  if (typeof body !== "object" || body === null) {
    throw new Error("Invalid Microsoft Graph message response.");
  }
  const message = body as Record<string, unknown>;
  const messageBody = message.body as Record<string, unknown> | undefined;
  if (
    typeof message.id !== "string" ||
    typeof message.subject !== "string" ||
    typeof message.receivedDateTime !== "string" ||
    typeof messageBody?.content !== "string" ||
    typeof messageBody.contentType !== "string"
  ) {
    throw new Error("Invalid Microsoft Graph message response.");
  }

  const from = message.from as Record<string, unknown> | undefined;
  const sender = from?.emailAddress as Record<string, unknown> | undefined;
  const recipients = Array.isArray(message.toRecipients)
    ? message.toRecipients.flatMap((recipient): string[] => {
        if (typeof recipient !== "object" || recipient === null) return [];
        const emailAddress = (recipient as Record<string, unknown>).emailAddress;
        if (typeof emailAddress !== "object" || emailAddress === null) return [];
        const address = (emailAddress as Record<string, unknown>).address;
        return typeof address === "string" ? [address] : [];
      })
    : [];

  return {
    id: message.id,
    senderName: typeof sender?.name === "string" ? sender.name : "Unknown sender",
    senderEmail: typeof sender?.address === "string" ? sender.address : "",
    recipients,
    subject: message.subject || "(No subject)",
    receivedAt: message.receivedDateTime,
    body: readableBody(messageBody.content, messageBody.contentType),
    bodyContentType: messageBody.contentType.toLowerCase() === "html" ? "html" : "text",
  };
}

export async function sendOutlookReply(
  accessToken: string,
  messageId: string,
  replyBody: string,
) {
  const response = await fetch(
    `https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(messageId)}/reply`,
    {
      method: "POST",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ comment: replyBody }),
      cache: "no-store",
    },
  );
  if (!response.ok) throw await createGraphError(response);
  if (response.status !== 202) {
    throw new Error("Unexpected Microsoft Graph reply response.");
  }
}
