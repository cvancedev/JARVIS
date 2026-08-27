import { NextResponse } from "next/server";
import { clearTokenCookies } from "@/lib/outlook/session";

export function POST() {
  const response = NextResponse.json({ state: "not_connected" });
  clearTokenCookies(response);
  return response;
}
