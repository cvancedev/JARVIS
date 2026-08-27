import { NextResponse } from "next/server";
import {
  getOutlookMessage,
  OutlookGraphError,
  sendOutlookReply,
} from "@/lib/outlook/graph";
import { getValidAccessToken } from "@/lib/outlook/oauth";

export const runtime = "nodejs";

const MAX_REPLY_LENGTH = 50_000;

export async function POST(request: Request) {
  let requestBody: unknown;
  try {
    requestBody = await request.json();
  } catch {
    return NextResponse.json({ error: "Invalid request body." }, { status: 400 });
  }

  if (typeof requestBody !== "object" || requestBody === null) {
    return NextResponse.json({ error: "Invalid request body." }, { status: 400 });
  }
  const body = requestBody as Record<string, unknown>;
  const messageId = typeof body.messageId === "string" ? body.messageId : "";
  const replyBody = typeof body.replyBody === "string" ? body.replyBody.trim() : "";
  if (!messageId || messageId.length > 2_048) {
    return NextResponse.json({ error: "Invalid message ID." }, { status: 400 });
  }
  if (!replyBody || replyBody.length > MAX_REPLY_LENGTH) {
    return NextResponse.json(
      { error: `Reply body must be between 1 and ${MAX_REPLY_LENGTH} characters.` },
      { status: 400 },
    );
  }

  const tokenResponse = NextResponse.json({});
  try {
    const accessToken = await getValidAccessToken(tokenResponse);
    if (!accessToken) {
      return NextResponse.json({ error: "Outlook is not connected." }, { status: 401 });
    }

    // Re-fetching verifies that this ID still resolves in the connected mailbox.
    await getOutlookMessage(accessToken, messageId);
    await sendOutlookReply(accessToken, messageId, replyBody);
    return NextResponse.json({ sent: true }, { headers: tokenResponse.headers });
  } catch (error: unknown) {
    if (error instanceof OutlookGraphError) {
      const permissionRequired =
        error.upstreamStatus === 403 ||
        /accessdenied|authorization/i.test(error.graphErrorCode);

      if (permissionRequired) {
        return NextResponse.json(
          {
            error: "Outlook send permission is not configured. Reconnect after granting delegated Mail.Send.",
            code: "mail_send_permission_required",
            ...(process.env.NODE_ENV === "development"
              ? {
                  diagnostic: {
                    upstreamStatus: error.upstreamStatus,
                    graphErrorCode: error.graphErrorCode,
                    graphErrorMessage: error.graphErrorMessage,
                    requestId: error.requestId,
                  },
                }
              : {}),
          },
          { status: 403, headers: tokenResponse.headers },
        );
      }
    }

    return NextResponse.json(
      { error: "Unable to send the Outlook reply." },
      { status: 502, headers: tokenResponse.headers },
    );
  }
}
