import { NextResponse } from "next/server";
import { createAuthorizationRequest } from "@/lib/outlook/oauth";
import { setOAuthCookies } from "@/lib/outlook/session";

export const runtime = "nodejs";

export function GET(request: Request) {
  try {
    const { url, state, verifier } = createAuthorizationRequest();
    const response = NextResponse.redirect(url);
    setOAuthCookies(response, state, verifier);
    return response;
  } catch {
    return NextResponse.redirect(new URL("/?outlook=configuration_error", request.url));
  }
}
