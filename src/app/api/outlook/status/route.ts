import { NextResponse } from "next/server";
import {
  delegatedScopeStatus,
  getValidAccessTokenDetails,
} from "@/lib/outlook/oauth";
import { clearTokenCookies } from "@/lib/outlook/session";

export const runtime = "nodejs";

export async function GET() {
  const tokenResponse = NextResponse.json({});
  try {
    const token = await getValidAccessTokenDetails(tokenResponse);
    return token
      ? NextResponse.json(
          {
            state: "connected",
            permissions: delegatedScopeStatus(token.grantedScopes),
          },
          { headers: tokenResponse.headers },
        )
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
