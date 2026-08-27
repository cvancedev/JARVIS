import { NextResponse } from "next/server";
import {
  getRecentInboxMessages,
  OutlookGraphError,
} from "@/lib/outlook/graph";
import { getValidAccessToken } from "@/lib/outlook/oauth";

export const runtime = "nodejs";

export async function GET() {
  const tokenResponse = NextResponse.json({ messages: [] });
  try {
    const accessToken = await getValidAccessToken(tokenResponse);
    if (!accessToken) {
      return NextResponse.json({ error: "Outlook is not connected." }, { status: 401 });
    }
    const messages = await getRecentInboxMessages(accessToken);
    return NextResponse.json({ messages }, { headers: tokenResponse.headers });
  } catch (error: unknown) {
    if (
      process.env.NODE_ENV === "development" &&
      error instanceof OutlookGraphError
    ) {
      return NextResponse.json(
        {
          error: "Unable to retrieve the Outlook inbox.",
          diagnostic: {
            upstreamStatus: error.upstreamStatus,
            graphErrorCode: error.graphErrorCode,
            graphErrorMessage: error.graphErrorMessage,
            requestId: error.requestId,
          },
        },
        { status: 502, headers: tokenResponse.headers },
      );
    }

    return NextResponse.json(
      { error: "Unable to retrieve the Outlook inbox." },
      { status: 502, headers: tokenResponse.headers },
    );
  }
}
