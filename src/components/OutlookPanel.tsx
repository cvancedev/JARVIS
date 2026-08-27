"use client";

import { useEffect, useState } from "react";
import type { OutlookConnectionState, OutlookMessage } from "@/types/outlook";

export default function OutlookPanel() {
  const [connectionState, setConnectionState] =
    useState<OutlookConnectionState>("connecting");
  const [messages, setMessages] = useState<OutlookMessage[]>([]);
  const [isLoadingInbox, setIsLoadingInbox] = useState(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    const loadStatus = async () => {
      const callbackResult = new URLSearchParams(window.location.search).get(
        "outlook",
      );
      if (
        callbackResult === "authentication_error" ||
        callbackResult === "configuration_error"
      ) {
        setConnectionState("authentication_error");
        return;
      }

      try {
        const response = await fetch("/api/outlook/status", { cache: "no-store" });
        const body: unknown = await response.json();
        const state =
          typeof body === "object" && body !== null &&
          "state" in body && typeof body.state === "string"
            ? body.state
            : "authentication_error";
        setConnectionState(
          state === "connected" || state === "not_connected"
            ? state
            : "authentication_error",
        );
      } catch {
        setConnectionState("authentication_error");
      }
    };
    void loadStatus();
  }, []);

  const loadInbox = async () => {
    setIsLoadingInbox(true);
    setError(null);
    try {
      const response = await fetch("/api/outlook/inbox", { cache: "no-store" });
      if (!response.ok) throw new Error("Inbox request failed.");
      const body: unknown = await response.json();
      if (
        typeof body !== "object" || body === null ||
        !("messages" in body) || !Array.isArray(body.messages)
      ) throw new Error("Invalid inbox response.");
      setMessages(body.messages as OutlookMessage[]);
    } catch {
      setError("Could not load the inbox.");
    } finally {
      setIsLoadingInbox(false);
    }
  };

  const disconnect = async () => {
    await fetch("/api/outlook/disconnect", { method: "POST" });
    setConnectionState("not_connected");
    setMessages([]);
    setError(null);
  };

  const stateLabel = connectionState === "not_connected"
    ? "Not Connected"
    : connectionState === "connecting"
      ? "Connecting"
      : connectionState === "connected"
        ? "Connected"
        : "Authentication Error";

  return (
    <section className="mt-6 border-t border-zinc-800 pt-4">
      <div className="flex items-center justify-between gap-2">
        <h2 className="text-xs font-semibold uppercase tracking-wide text-zinc-500">
          Outlook
        </h2>
        <span className="text-xs text-zinc-500">{stateLabel}</span>
      </div>

      {connectionState === "connected" ? (
        <div className="mt-3 flex gap-2">
          <button
            type="button"
            className="rounded-md border border-zinc-700 px-2 py-1.5 text-xs text-zinc-300 hover:border-zinc-500 disabled:opacity-50"
            onClick={loadInbox}
            disabled={isLoadingInbox}
          >
            {isLoadingInbox ? "Loading..." : "Recent Inbox"}
          </button>
          <button
            type="button"
            className="text-xs text-zinc-500 hover:text-zinc-300"
            onClick={disconnect}
          >
            Disconnect
          </button>
        </div>
      ) : (
        <button
          type="button"
          className="mt-3 w-full rounded-md border border-zinc-700 px-3 py-2 text-left text-xs text-zinc-300 hover:border-zinc-500 disabled:opacity-50"
          onClick={() => {
            setConnectionState("connecting");
            window.location.assign("/api/outlook/connect");
          }}
          disabled={connectionState === "connecting"}
        >
          Connect Outlook
        </button>
      )}

      {error ? <p className="mt-2 text-xs text-red-400">{error}</p> : null}

      {messages.length > 0 ? (
        <ul className="mt-3 space-y-2">
          {messages.map((message) => (
            <li key={message.id} className="rounded-md bg-zinc-900 p-2 text-xs">
              <p className={message.isRead ? "text-zinc-400" : "font-medium text-white"}>
                {message.subject}
              </p>
              <p className="mt-1 truncate text-zinc-500" title={message.senderEmail}>
                {message.senderName}
              </p>
              <p className="mt-1 line-clamp-2 text-zinc-600">{message.preview}</p>
              <time className="mt-1 block text-zinc-600" dateTime={message.receivedAt}>
                {new Date(message.receivedAt).toLocaleString()}
              </time>
            </li>
          ))}
        </ul>
      ) : null}
    </section>
  );
}
