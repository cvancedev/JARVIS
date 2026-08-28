import { createHash, randomBytes, timingSafeEqual } from "node:crypto";
import type { NextResponse } from "next/server";
import {
  getOutlookConfig,
  normalizedDelegatedScopes,
  outlookScopeVersion,
} from "@/lib/outlook/config";
import { readAccessSession, readRefreshToken, setTokenCookies } from "@/lib/outlook/session";

interface RawTokenResponse {
  access_token: string;
  expires_in: number;
  refresh_token?: string;
  scope?: string;
}

export class OutlookScopeError extends Error {
  constructor(public readonly missingScopes: string[]) {
    super("Microsoft access token is missing required delegated scopes.");
    this.name = "OutlookScopeError";
  }
}

function tokenClaimScopes(accessToken: string) {
  try {
    const payload = accessToken.split(".")[1];
    if (!payload) return [];
    const decoded: unknown = JSON.parse(Buffer.from(payload, "base64url").toString("utf8"));
    if (typeof decoded !== "object" || decoded === null) return [];
    const scope = (decoded as Record<string, unknown>).scp;
    return typeof scope === "string"
      ? scope.split(/\s+/).filter(Boolean).map((value) => value.toLowerCase()).sort()
      : [];
  } catch {
    return [];
  }
}

function grantedScopes(token: Pick<RawTokenResponse, "access_token" | "scope">) {
  const responseScopes = token.scope
    ? token.scope.split(/\s+/).filter(Boolean).map((value) =>
        value.slice(value.lastIndexOf("/") + 1).toLowerCase())
    : [];
  return [...new Set([...responseScopes, ...tokenClaimScopes(token.access_token)])].sort();
}

function missingRequiredScopes(granted: string[], requested: string[]) {
  const grantedSet = new Set(granted.map((scope) => scope.toLowerCase()));
  return normalizedDelegatedScopes(requested).filter((scope) => !grantedSet.has(scope));
}

export function delegatedScopeStatus(scopes: string[]) {
  return {
    calendarsReadWrite: scopes.includes("calendars.readwrite"),
  };
}

function isTokenResponse(value: unknown): value is RawTokenResponse {
  if (typeof value !== "object" || value === null) return false;
  const token = value as Record<string, unknown>;
  return typeof token.access_token === "string" &&
    typeof token.expires_in === "number" &&
    (token.refresh_token === undefined || typeof token.refresh_token === "string") &&
    (token.scope === undefined || typeof token.scope === "string");
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
  const config = getOutlookConfig();
  const resolvedScopes = grantedScopes(body);
  const missingScopes = missingRequiredScopes(resolvedScopes, config.scopes);
  if (missingScopes.length) throw new OutlookScopeError(missingScopes);
  return {
    ...body,
    grantedScopes: resolvedScopes,
    scopeVersion: outlookScopeVersion(config.scopes),
  };
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

export async function getValidAccessTokenDetails(response: NextResponse) {
  const config = getOutlookConfig();
  const currentScopeVersion = outlookScopeVersion(config.scopes);
  const accessSession = await readAccessSession();
  if (
    accessSession &&
    accessSession.expiresAt > Date.now() &&
    accessSession.scopeVersion === currentScopeVersion &&
    missingRequiredScopes(accessSession.grantedScopes, config.scopes).length === 0
  ) {
    return {
      accessToken: accessSession.token,
      grantedScopes: accessSession.grantedScopes,
    };
  }

  const refreshToken = await readRefreshToken();
  if (!refreshToken) return null;
  const token = await requestToken(new URLSearchParams({
    client_id: config.clientId,
    client_secret: config.clientSecret,
    grant_type: "refresh_token",
    refresh_token: refreshToken,
    scope: config.scopes.join(" "),
  }));
  setTokenCookies(
    response,
    token.access_token,
    token.expires_in,
    token.grantedScopes,
    token.scopeVersion,
    token.refresh_token,
  );
  return {
    accessToken: token.access_token,
    grantedScopes: token.grantedScopes,
  };
}

export async function getValidAccessToken(response: NextResponse) {
  return (await getValidAccessTokenDetails(response))?.accessToken ?? null;
}
