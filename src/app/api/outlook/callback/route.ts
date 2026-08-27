import { NextRequest, NextResponse } from "next/server";
import { exchangeAuthorizationCode, matchesOAuthState } from "@/lib/outlook/oauth";
import {
  clearOAuthCookies,
  OUTLOOK_STATE_COOKIE,
  OUTLOOK_VERIFIER_COOKIE,
  setTokenCookies,
} from "@/lib/outlook/session";

export const runtime = "nodejs";

export async function GET(request: NextRequest) {
  const code = request.nextUrl.searchParams.get("code");
  const state = request.nextUrl.searchParams.get("state");
  const expectedState = request.cookies.get(OUTLOOK_STATE_COOKIE)?.value;
  const verifier = request.cookies.get(OUTLOOK_VERIFIER_COOKIE)?.value;
  const redirect = (result: string) =>
    NextResponse.redirect(new URL(`/?outlook=${result}`, request.url));

  if (
    request.nextUrl.searchParams.has("error") ||
    !code ||
    !state ||
    !expectedState ||
    !verifier ||
    !matchesOAuthState(expectedState, state)
  ) {
    const response = redirect("authentication_error");
    clearOAuthCookies(response);
    return response;
  }

  try {
    const token = await exchangeAuthorizationCode(code, verifier);
    const response = redirect("connected");
    clearOAuthCookies(response);
    setTokenCookies(response, token.access_token, token.expires_in, token.refresh_token);
    return response;
  } catch {
    const response = redirect("authentication_error");
    clearOAuthCookies(response);
    return response;
  }
}
