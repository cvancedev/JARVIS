import { NextResponse } from "next/server";
import { getOutlookMessage, OutlookGraphError } from "@/lib/outlook/graph";
import { getValidAccessToken } from "@/lib/outlook/oauth";

export const runtime = "nodejs";

export async function GET(
  _request: Request,
  { params }: { params: Promise<{ id: string }> },
) {
  const { id } = await params;
  if (!id || id.length > 2_048) {
    return NextResponse.json({ error: "Invalid message ID." }, { status: 400 });
  }

  const tokenResponse = NextResponse.json({ message: null });
  try {
    const accessToken = await getValidAccessToken(tokenResponse);
    if (!accessToken) {
      return NextResponse.json({ error: "Outlook is not connected." }, { status: 401 });
    }
    const message = await getOutlookMessage(accessToken, id);
    return NextResponse.json({ message }, { headers: tokenResponse.headers });
  } catch (error: unknown) {
    const status = error instanceof OutlookGraphError && error.upstreamStatus === 404
      ? 404
      : 502;
    return NextResponse.json(
      { error: status === 404 ? "Email not found." : "Unable to retrieve the email." },
      { status, headers: tokenResponse.headers },
    );
  }
}
