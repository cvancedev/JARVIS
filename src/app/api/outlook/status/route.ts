import { NextResponse } from "next/server";
import { getValidAccessToken } from "@/lib/outlook/oauth";
import { clearTokenCookies } from "@/lib/outlook/session";

export const runtime = "nodejs";

export async function GET() {
  const connectedResponse = NextResponse.json({ state: "connected" });
  try {
    const accessToken = await getValidAccessToken(connectedResponse);
    return accessToken
      ? connectedResponse
      : NextResponse.json({ state: "not_connected" });
  } catch {
    const errorResponse = NextResponse.json(
      { state: "authentication_error" },
      { status: 401 },
    );
    clearTokenCookies(errorResponse);
    return errorResponse;
  }
}
