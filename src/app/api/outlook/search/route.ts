import { NextRequest, NextResponse } from "next/server";
import { searchOutlookMessages } from "@/lib/outlook/graph";
import { getValidAccessToken } from "@/lib/outlook/oauth";

export const runtime = "nodejs";

const MAX_SEARCH_LENGTH = 200;

export async function GET(request: NextRequest) {
  const query = request.nextUrl.searchParams.get("q")?.trim() ?? "";
  if (!query) {
    return NextResponse.json({ error: "Search query is required." }, { status: 400 });
  }
  if (query.length > MAX_SEARCH_LENGTH) {
    return NextResponse.json(
      { error: `Search query must be ${MAX_SEARCH_LENGTH} characters or fewer.` },
      { status: 400 },
    );
  }

  const tokenResponse = NextResponse.json({ messages: [] });
  try {
    const accessToken = await getValidAccessToken(tokenResponse);
    if (!accessToken) {
      return NextResponse.json({ error: "Outlook is not connected." }, { status: 401 });
    }
    const messages = await searchOutlookMessages(accessToken, query);
    return NextResponse.json({ messages }, { headers: tokenResponse.headers });
  } catch {
    return NextResponse.json(
      { error: "Unable to search Outlook." },
      { status: 502, headers: tokenResponse.headers },
    );
  }
}
