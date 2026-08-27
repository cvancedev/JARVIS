import { createHash, randomBytes, timingSafeEqual } from "node:crypto";
import type { NextResponse } from "next/server";
import { getOutlookConfig } from "@/lib/outlook/config";
import { readAccessSession, readRefreshToken, setTokenCookies } from "@/lib/outlook/session";

interface TokenResponse {
  access_token: string;
  expires_in: number;
  refresh_token?: string;
}

function isTokenResponse(value: unknown): value is TokenResponse {
  if (typeof value !== "object" || value === null) return false;
  const token = value as Record<string, unknown>;
  return typeof token.access_token === "string" &&
    typeof token.expires_in === "number" &&
    (token.refresh_token === undefined || typeof token.refresh_token === "string");
}

async function requestToken(parameters: URLSearchParams) {
  const response = await fetch(getOutlookConfig().tokenUrl, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: parameters,
    cache: "no-store",
  });
  if (!response.ok) throw new Error("Microsoft token exchange failed.");
  const body: unknown = await response.json();
  if (!isTokenResponse(body)) throw new Error("Invalid Microsoft token response.");
  return body;
}

export function createAuthorizationRequest() {
  const config = getOutlookConfig();
  const state = randomBytes(32).toString("base64url");
  const verifier = randomBytes(48).toString("base64url");
  const challenge = createHash("sha256").update(verifier).digest("base64url");
  const url = new URL(config.authorizeUrl);
  url.search = new URLSearchParams({
    client_id: config.clientId,
    response_type: "code",
    redirect_uri: config.redirectUri,
    response_mode: "query",
    scope: config.scopes.join(" "),
    state,
    code_challenge: challenge,
    code_challenge_method: "S256",
  }).toString();
  return { url, state, verifier };
}

export function matchesOAuthState(expected: string, received: string) {
  const a = Buffer.from(expected);
  const b = Buffer.from(received);
  return a.length === b.length && timingSafeEqual(a, b);
}

export async function exchangeAuthorizationCode(code: string, verifier: string) {
  const config = getOutlookConfig();
  return requestToken(new URLSearchParams({
    client_id: config.clientId,
    client_secret: config.clientSecret,
    grant_type: "authorization_code",
    code,
    redirect_uri: config.redirectUri,
    code_verifier: verifier,
    scope: config.scopes.join(" "),
  }));
}

export async function getValidAccessToken(response: NextResponse) {
  const accessSession = await readAccessSession();
  if (accessSession && accessSession.expiresAt > Date.now()) return accessSession.token;

  const refreshToken = await readRefreshToken();
  if (!refreshToken) return null;
  const config = getOutlookConfig();
  const token = await requestToken(new URLSearchParams({
    client_id: config.clientId,
    client_secret: config.clientSecret,
    grant_type: "refresh_token",
    refresh_token: refreshToken,
    scope: config.scopes.join(" "),
  }));
  setTokenCookies(response, token.access_token, token.expires_in, token.refresh_token);
  return token.access_token;
}
