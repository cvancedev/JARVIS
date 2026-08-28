import { createCipheriv, createDecipheriv, createHash, randomBytes } from "node:crypto";
import { cookies } from "next/headers";
import type { NextResponse } from "next/server";
import { getOutlookConfig } from "@/lib/outlook/config";

export const OUTLOOK_ACCESS_COOKIE = "jarvis_outlook_access";
export const OUTLOOK_REFRESH_COOKIE = "jarvis_outlook_refresh";
export const OUTLOOK_STATE_COOKIE = "jarvis_outlook_state";
export const OUTLOOK_VERIFIER_COOKIE = "jarvis_outlook_verifier";

interface AccessSession {
  token: string;
  expiresAt: number;
  grantedScopes: string[];
  scopeVersion: string;
}

function encryptionKey() {
  return createHash("sha256").update(getOutlookConfig().sessionSecret).digest();
}

function encrypt(value: string) {
  const iv = randomBytes(12);
  const cipher = createCipheriv("aes-256-gcm", encryptionKey(), iv);
  const ciphertext = Buffer.concat([cipher.update(value, "utf8"), cipher.final()]);
  return [iv, cipher.getAuthTag(), ciphertext]
    .map((part) => part.toString("base64url"))
    .join(".");
}

function decrypt(value: string) {
  try {
    const [iv, tag, ciphertext] = value.split(".");
    if (!iv || !tag || !ciphertext) return null;
    const decipher = createDecipheriv(
      "aes-256-gcm",
      encryptionKey(),
      Buffer.from(iv, "base64url"),
    );
    decipher.setAuthTag(Buffer.from(tag, "base64url"));
    return Buffer.concat([
      decipher.update(Buffer.from(ciphertext, "base64url")),
      decipher.final(),
    ]).toString("utf8");
  } catch {
    return null;
  }
}

const secureCookie = {
  httpOnly: true,
  sameSite: "lax" as const,
  secure: process.env.NODE_ENV === "production",
  path: "/",
};

export function setOAuthCookies(response: NextResponse, state: string, verifier: string) {
  response.cookies.set(OUTLOOK_STATE_COOKIE, state, { ...secureCookie, maxAge: 600 });
  response.cookies.set(OUTLOOK_VERIFIER_COOKIE, verifier, { ...secureCookie, maxAge: 600 });
}

export function clearOAuthCookies(response: NextResponse) {
  response.cookies.set(OUTLOOK_STATE_COOKIE, "", { ...secureCookie, maxAge: 0 });
  response.cookies.set(OUTLOOK_VERIFIER_COOKIE, "", { ...secureCookie, maxAge: 0 });
}

export function setTokenCookies(
  response: NextResponse,
  accessToken: string,
  expiresIn: number,
  grantedScopes: string[],
  scopeVersion: string,
  refreshToken?: string,
) {
  const accessSession: AccessSession = {
    token: accessToken,
    expiresAt: Date.now() + Math.max(expiresIn - 60, 60) * 1_000,
    grantedScopes,
    scopeVersion,
  };
  response.cookies.set(
    OUTLOOK_ACCESS_COOKIE,
    encrypt(JSON.stringify(accessSession)),
    { ...secureCookie, maxAge: Math.max(expiresIn, 60) },
  );
  if (refreshToken) {
    response.cookies.set(OUTLOOK_REFRESH_COOKIE, encrypt(refreshToken), {
      ...secureCookie,
      maxAge: 30 * 24 * 60 * 60,
    });
  }
}

export function clearTokenCookies(response: NextResponse) {
  response.cookies.set(OUTLOOK_ACCESS_COOKIE, "", { ...secureCookie, maxAge: 0 });
  response.cookies.set(OUTLOOK_REFRESH_COOKIE, "", { ...secureCookie, maxAge: 0 });
}

export async function readAccessSession() {
  const encrypted = (await cookies()).get(OUTLOOK_ACCESS_COOKIE)?.value;
  const decrypted = encrypted ? decrypt(encrypted) : null;
  if (!decrypted) return null;
  try {
    const parsed: unknown = JSON.parse(decrypted);
    if (typeof parsed !== "object" || parsed === null) return null;
    const session = parsed as Record<string, unknown>;
    return typeof session.token === "string" &&
      typeof session.expiresAt === "number" &&
      Array.isArray(session.grantedScopes) &&
      session.grantedScopes.every((scope) => typeof scope === "string") &&
      typeof session.scopeVersion === "string"
      ? ({
          token: session.token,
          expiresAt: session.expiresAt,
          grantedScopes: session.grantedScopes,
          scopeVersion: session.scopeVersion,
        } satisfies AccessSession)
      : null;
  } catch {
    return null;
  }
}

export async function readRefreshToken() {
  const encrypted = (await cookies()).get(OUTLOOK_REFRESH_COOKIE)?.value;
  return encrypted ? decrypt(encrypted) : null;
}
