"use client";

import { useEffect, useRef, useState } from "react";
import type {
  OutlookConnectionState,
  OutlookMessage,
  OutlookMessageDetail,
} from "@/types/outlook";

export default function OutlookPanel() {
  const [connectionState, setConnectionState] =
    useState<OutlookConnectionState>("connecting");
  const [messages, setMessages] = useState<OutlookMessage[]>([]);
  const [selectedMessage, setSelectedMessage] =
    useState<OutlookMessageDetail | null>(null);
  const [isLoadingInbox, setIsLoadingInbox] = useState(false);
  const [searchQuery, setSearchQuery] = useState("");
  const [isSearching, setIsSearching] = useState(false);
  const [resultsMode, setResultsMode] = useState<"recent" | "search">("recent");
  const [isLoadingMessage, setIsLoadingMessage] = useState(false);
  const [isSummarizing, setIsSummarizing] = useState(false);
  const [summary, setSummary] = useState("");
  const [draftInstruction, setDraftInstruction] = useState("");
  const [draftReply, setDraftReply] = useState("");
  const [isDrafting, setIsDrafting] = useState(false);
  const [isConfirmingSend, setIsConfirmingSend] = useState(false);
  const [isSending, setIsSending] = useState(false);
  const [sendSuccess, setSendSuccess] = useState(false);
  const sendActiveRef = useRef(false);
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
    if (isDrafting || isSummarizing || sendActiveRef.current) return;
    setIsLoadingInbox(true);
    setError(null);
    setSelectedMessage(null);
    setSummary("");
    setDraftInstruction("");
    setDraftReply("");
    setIsConfirmingSend(false);
    setSendSuccess(false);
    try {
      const response = await fetch("/api/outlook/inbox", { cache: "no-store" });
      if (!response.ok) throw new Error("Inbox request failed.");
      const body: unknown = await response.json();
      if (
        typeof body !== "object" || body === null ||
        !("messages" in body) || !Array.isArray(body.messages)
      ) throw new Error("Invalid inbox response.");
      setMessages(body.messages as OutlookMessage[]);
      setResultsMode("recent");
    } catch {
      setError("Could not load the inbox.");
    } finally {
      setIsLoadingInbox(false);
    }
  };

  const searchInbox = async () => {
    const trimmedQuery = searchQuery.trim();
    if (
      !trimmedQuery ||
      trimmedQuery.length > 200 ||
      isSearching ||
      isDrafting ||
      isSummarizing ||
      sendActiveRef.current
    ) return;

    setIsSearching(true);
    setError(null);
    setSelectedMessage(null);
    setSummary("");
    setDraftInstruction("");
    setDraftReply("");
    setIsConfirmingSend(false);
    setSendSuccess(false);
    try {
      const response = await fetch(
        `/api/outlook/search?q=${encodeURIComponent(trimmedQuery)}`,
        { cache: "no-store" },
      );
      if (!response.ok) throw new Error("Search request failed.");
      const body: unknown = await response.json();
      if (
        typeof body !== "object" || body === null ||
        !("messages" in body) || !Array.isArray(body.messages)
      ) throw new Error("Invalid search response.");
      setMessages(body.messages as OutlookMessage[]);
      setResultsMode("search");
    } catch {
      setError("Could not search Outlook.");
    } finally {
      setIsSearching(false);
    }
  };

  const disconnect = async () => {
    if (isDrafting || isSummarizing || sendActiveRef.current) return;
    await fetch("/api/outlook/disconnect", { method: "POST" });
    setConnectionState("not_connected");
    setMessages([]);
    setSelectedMessage(null);
    setSummary("");
    setDraftInstruction("");
    setDraftReply("");
    setIsConfirmingSend(false);
    setSendSuccess(false);
    setError(null);
  };

  const openMessage = async (messageId: string) => {
    if (isDrafting || isSummarizing || sendActiveRef.current) return;
    setIsLoadingMessage(true);
    setSelectedMessage(null);
    setSummary("");
    setDraftInstruction("");
    setDraftReply("");
    setIsConfirmingSend(false);
    setSendSuccess(false);
    setError(null);
    try {
      const response = await fetch(
        `/api/outlook/messages/${encodeURIComponent(messageId)}`,
        { cache: "no-store" },
      );
      if (!response.ok) throw new Error("Message request failed.");
      const body: unknown = await response.json();
      if (
        typeof body !== "object" || body === null ||
        !("message" in body) || typeof body.message !== "object" ||
        body.message === null
      ) throw new Error("Invalid message response.");
      setSelectedMessage(body.message as OutlookMessageDetail);
    } catch {
      setError("Could not open the email.");
    } finally {
      setIsLoadingMessage(false);
    }
  };

  const summarizeMessage = async () => {
    if (!selectedMessage || isSummarizing) return;
    setIsSummarizing(true);
    setSummary("");
    setError(null);
    try {
      const response = await fetch("/api/outlook/summarize", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ messageId: selectedMessage.id }),
      });
      if (!response.ok || !response.body) throw new Error("Summary request failed.");

      const reader = response.body.getReader();
      const decoder = new TextDecoder();
      let summaryText = "";
      while (true) {
        const { done, value } = await reader.read();
        if (done) {
          summaryText += decoder.decode();
          setSummary(summaryText);
          break;
        }
        summaryText += decoder.decode(value, { stream: true });
        setSummary(summaryText);
      }
    } catch {
      setSummary("");
      setError("Could not summarize the email.");
    } finally {
      setIsSummarizing(false);
    }
  };

  const draftMessageReply = async () => {
    if (!selectedMessage || isDrafting) return;
    setIsDrafting(true);
    setDraftReply("");
    setIsConfirmingSend(false);
    setSendSuccess(false);
    setError(null);
    try {
      const response = await fetch("/api/outlook/draft-reply", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          messageId: selectedMessage.id,
          instruction: draftInstruction.trim(),
        }),
      });
      if (!response.ok || !response.body) throw new Error("Draft request failed.");

      const reader = response.body.getReader();
      const decoder = new TextDecoder();
      let draftText = "";
      while (true) {
        const { done, value } = await reader.read();
        if (done) {
          draftText += decoder.decode();
          setDraftReply(draftText);
          break;
        }
        draftText += decoder.decode(value, { stream: true });
        setDraftReply(draftText);
      }
    } catch {
      setDraftReply("");
      setError("Could not generate a reply draft.");
    } finally {
      setIsDrafting(false);
    }
  };

  const sendReply = async () => {
    if (
      !selectedMessage ||
      !isConfirmingSend ||
      !draftReply.trim() ||
      sendActiveRef.current
    ) return;

    sendActiveRef.current = true;
    setIsSending(true);
    setError(null);
    try {
      const response = await fetch("/api/outlook/send-reply", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          messageId: selectedMessage.id,
          replyBody: draftReply,
        }),
      });
      const responseBody: unknown = await response.json();
      if (!response.ok) {
        const permissionRequired =
          typeof responseBody === "object" &&
          responseBody !== null &&
          "code" in responseBody &&
          responseBody.code === "mail_send_permission_required";
        throw new Error(permissionRequired ? "mail_send_permission_required" : "send_failed");
      }

      setDraftReply("");
      setDraftInstruction("");
      setIsConfirmingSend(false);
      setSendSuccess(true);
    } catch (sendError: unknown) {
      setError(
        sendError instanceof Error &&
        sendError.message === "mail_send_permission_required"
          ? "Outlook needs delegated Mail.Send permission. Grant it in Entra, then reconnect Outlook."
          : "Could not send the Outlook reply. Your draft was preserved.",
      );
    } finally {
      sendActiveRef.current = false;
      setIsSending(false);
    }
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
        <div className="mt-3">
          <div className="flex gap-2">
            <button
              type="button"
              className="rounded-md border border-zinc-700 px-2 py-1.5 text-xs text-zinc-300 hover:border-zinc-500 disabled:opacity-50"
              onClick={loadInbox}
              disabled={isLoadingInbox || isSearching || isDrafting || isSummarizing || isSending}
            >
              {isLoadingInbox ? "Loading..." : "Recent Inbox"}
            </button>
            <button
              type="button"
              className="text-xs text-zinc-500 hover:text-zinc-300"
              onClick={disconnect}
              disabled={isDrafting || isSummarizing || isSending}
            >
              Disconnect
            </button>
          </div>
          <form
            className="mt-2 flex gap-1"
            onSubmit={(event) => {
              event.preventDefault();
              void searchInbox();
            }}
          >
            <input
              type="search"
              className="min-w-0 flex-1 rounded-md border border-zinc-700 bg-zinc-900 px-2 py-1.5 text-xs text-white outline-none placeholder:text-zinc-600 focus:border-zinc-500"
              value={searchQuery}
              onChange={(event) => setSearchQuery(event.target.value)}
              placeholder="Search Outlook"
              maxLength={200}
              aria-label="Search Outlook messages"
            />
            <button
              type="submit"
              className="rounded-md border border-zinc-700 px-2 py-1.5 text-xs text-zinc-300 hover:border-zinc-500 disabled:opacity-50"
              disabled={
                !searchQuery.trim() ||
                isSearching ||
                isLoadingInbox ||
                isDrafting ||
                isSummarizing ||
                isSending
              }
            >
              {isSearching ? "Searching..." : "Search"}
            </button>
          </form>
          {resultsMode === "search" ? (
            <button
              type="button"
              className="mt-2 text-xs text-zinc-500 hover:text-zinc-300"
              onClick={loadInbox}
              disabled={isLoadingInbox || isDrafting || isSummarizing || isSending}
            >
              Back to Recent Inbox
            </button>
          ) : null}
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
      {isLoadingMessage ? (
        <p className="mt-3 text-xs text-zinc-500">Opening email...</p>
      ) : null}

      {messages.length > 0 ? (
        <ul className="mt-3 space-y-2">
          {messages.map((message) => (
            <li key={message.id}>
              <button
                type="button"
                className="w-full rounded-md bg-zinc-900 p-2 text-left text-xs hover:bg-zinc-800"
                onClick={() => openMessage(message.id)}
                disabled={isLoadingMessage || isDrafting || isSummarizing || isSending}
              >
                <span className={message.isRead ? "text-zinc-400" : "font-medium text-white"}>
                  {message.subject}
                </span>
                <span className="mt-1 block truncate text-zinc-500" title={message.senderEmail}>
                  {message.senderName}
                </span>
                <span className="mt-1 line-clamp-2 text-zinc-600">{message.preview}</span>
                <time className="mt-1 block text-zinc-600" dateTime={message.receivedAt}>
                  {new Date(message.receivedAt).toLocaleString()}
                </time>
              </button>
            </li>
          ))}
        </ul>
      ) : resultsMode === "search" && !isSearching ? (
        <p className="mt-3 text-xs text-zinc-500">No matching messages.</p>
      ) : null}

      {selectedMessage ? (
        <article className="mt-4 rounded-md border border-zinc-700 bg-zinc-900 p-3 text-xs">
          <div className="flex items-start justify-between gap-2">
            <h3 className="font-medium text-white">{selectedMessage.subject}</h3>
            <button
              type="button"
              className="text-zinc-500 hover:text-white"
              onClick={() => {
                setSelectedMessage(null);
                setSummary("");
                setDraftInstruction("");
                setDraftReply("");
                setIsConfirmingSend(false);
                setSendSuccess(false);
              }}
              aria-label="Close email"
              disabled={isDrafting || isSummarizing || isSending}
            >
              Close
            </button>
          </div>
          <p className="mt-2 text-zinc-400">
            From: {selectedMessage.senderName} &lt;{selectedMessage.senderEmail}&gt;
          </p>
          {selectedMessage.recipients.length > 0 ? (
            <p className="mt-1 break-words text-zinc-500">
              To: {selectedMessage.recipients.join(", ")}
            </p>
          ) : null}
          <time className="mt-1 block text-zinc-500" dateTime={selectedMessage.receivedAt}>
            {new Date(selectedMessage.receivedAt).toLocaleString()}
          </time>
          <div className="mt-3 max-h-80 overflow-y-auto whitespace-pre-wrap border-t border-zinc-800 pt-3 leading-relaxed text-zinc-300">
            {selectedMessage.body || "(Empty message body)"}
          </div>
          <div className="mt-3 flex gap-2">
            <button
              type="button"
              className="rounded-md bg-blue-600 px-3 py-1.5 text-xs text-white hover:bg-blue-500 disabled:opacity-50"
              onClick={summarizeMessage}
              disabled={isSummarizing || isDrafting || isSending}
            >
              {isSummarizing ? "Summarizing..." : "Summarize"}
            </button>
            <button
              type="button"
              className="rounded-md border border-zinc-600 px-3 py-1.5 text-xs text-zinc-200 hover:border-zinc-400 disabled:opacity-50"
              onClick={draftMessageReply}
              disabled={isDrafting || isSummarizing || isSending}
            >
              {isDrafting ? "Drafting..." : "Draft Reply"}
            </button>
          </div>

          <input
            type="text"
            className="mt-2 w-full rounded-md border border-zinc-700 bg-zinc-950 px-2 py-1.5 text-xs text-white outline-none placeholder:text-zinc-600 focus:border-zinc-500"
            value={draftInstruction}
            onChange={(event) => setDraftInstruction(event.target.value)}
            placeholder="Optional: Make it shorter, keep it friendly..."
            maxLength={1_000}
            aria-label="Optional reply drafting instruction"
            disabled={isDrafting || isConfirmingSend || isSending}
          />

          {summary ? (
            <section className="mt-3 rounded-md border border-blue-900/60 bg-blue-950/30 p-3">
              <h4 className="font-medium text-blue-300">JARVIS Summary</h4>
              <p className="mt-2 whitespace-pre-wrap leading-relaxed text-zinc-300">
                {summary}
              </p>
            </section>
          ) : null}

          {draftReply || isDrafting ? (
            <section className="mt-3 rounded-md border border-emerald-900/60 bg-emerald-950/20 p-3">
              <h4 className="font-medium text-emerald-300">Draft Reply</h4>
              <textarea
                className="mt-2 min-h-40 w-full resize-y rounded-md border border-zinc-700 bg-zinc-950 p-2 text-xs leading-relaxed text-zinc-200 outline-none focus:border-zinc-500"
                value={draftReply}
                onChange={(event) => setDraftReply(event.target.value)}
                placeholder={isDrafting ? "Generating reply draft..." : "Reply draft"}
                disabled={isDrafting || isConfirmingSend || isSending}
              />
              <div className="mt-2 flex gap-2">
                <button
                  type="button"
                  className="rounded-md border border-zinc-700 px-2 py-1.5 text-xs text-zinc-300 hover:border-zinc-500 disabled:opacity-50"
                  onClick={draftMessageReply}
                  disabled={isDrafting || isSummarizing || isSending}
                >
                  Regenerate
                </button>
                <button
                  type="button"
                  className="text-xs text-zinc-500 hover:text-zinc-300 disabled:opacity-50"
                  onClick={() => {
                    setDraftReply("");
                    setIsConfirmingSend(false);
                    setSendSuccess(false);
                  }}
                  disabled={isDrafting || isSending}
                >
                  Clear Draft
                </button>
                <button
                  type="button"
                  className="ml-auto rounded-md bg-emerald-700 px-3 py-1.5 text-xs text-white hover:bg-emerald-600 disabled:opacity-50"
                  onClick={() => {
                    setError(null);
                    setSendSuccess(false);
                    setIsConfirmingSend(true);
                  }}
                  disabled={!draftReply.trim() || isDrafting || isSummarizing || isSending}
                >
                  Send
                </button>
              </div>

              {isConfirmingSend ? (
                <section className="mt-3 rounded-md border border-amber-700/70 bg-amber-950/20 p-3">
                  <h5 className="font-medium text-amber-300">Confirm Outlook Reply</h5>
                  <p className="mt-2 break-words text-zinc-300">
                    Recipient: {selectedMessage.senderName} &lt;{selectedMessage.senderEmail}&gt;
                  </p>
                  <p className="mt-1 break-words text-zinc-400">
                    In reply to: {selectedMessage.subject}
                  </p>
                  <div className="mt-3 max-h-64 overflow-y-auto whitespace-pre-wrap rounded-md border border-zinc-700 bg-zinc-950 p-2 leading-relaxed text-zinc-200">
                    {draftReply}
                  </div>
                  <div className="mt-3 flex justify-end gap-2">
                    <button
                      type="button"
                      className="rounded-md border border-zinc-700 px-3 py-1.5 text-zinc-300 hover:border-zinc-500 disabled:opacity-50"
                      onClick={() => setIsConfirmingSend(false)}
                      disabled={isSending}
                    >
                      Cancel
                    </button>
                    <button
                      type="button"
                      className="rounded-md bg-red-700 px-3 py-1.5 text-white hover:bg-red-600 disabled:opacity-50"
                      onClick={() => void sendReply()}
                      disabled={isSending}
                    >
                      {isSending ? "Sending..." : "Send Reply"}
                    </button>
                  </div>
                </section>
              ) : null}
            </section>
          ) : null}

          {sendSuccess ? (
            <p className="mt-3 rounded-md border border-emerald-800 bg-emerald-950/30 p-3 text-emerald-300">
              Outlook reply sent.
            </p>
          ) : null}
        </article>
      ) : null}
    </section>
  );
}
