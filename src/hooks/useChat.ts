"use client";

import { useEffect, useRef, useState } from "react";
import type { Message } from "@/types/message";

const STORAGE_KEY = "jarvis-chat-history";
const THINKING_MESSAGE = "Thinking...";
const INITIAL_MESSAGES: Message[] = [
  {
    id: "1",
    role: "assistant",
    content: "Welcome back, Curt. Awaiting instructions.",
  },
];

function isStoredMessage(value: unknown): value is Message {
  if (typeof value !== "object" || value === null) return false;

  const message = value as Record<string, unknown>;

  return (
    typeof message.id === "string" &&
    message.id.trim() !== "" &&
    (message.role === "user" || message.role === "assistant") &&
    typeof message.content === "string" &&
    message.content.trim() !== ""
  );
}

export function useChat() {
  const [input, setInput] = useState("");
  const [isLoading, setIsLoading] = useState(false);
  const [isThinking, setIsThinking] = useState(false);
  const [hasRestoredHistory, setHasRestoredHistory] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [messages, setMessages] = useState<Message[]>(INITIAL_MESSAGES);
  const requestActiveRef = useRef(false);
  const activeStreamingMessageIdRef = useRef<string | null>(null);
  const skipNextSaveRef = useRef(false);

  useEffect(() => {
    try {
      const savedHistory = localStorage.getItem(STORAGE_KEY);

      if (savedHistory) {
        const parsedHistory: unknown = JSON.parse(savedHistory);

        if (!Array.isArray(parsedHistory) || !parsedHistory.every(isStoredMessage)) {
          localStorage.removeItem(STORAGE_KEY);
        } else {
          const completeMessages = parsedHistory.filter(
            (message) =>
              message.role !== "assistant" || message.content !== THINKING_MESSAGE,
          );

          // Client-only storage cannot be read during the server render.
          // eslint-disable-next-line react-hooks/set-state-in-effect
          setMessages(completeMessages);
        }
      }
    } catch {
      try {
        localStorage.removeItem(STORAGE_KEY);
      } catch {
        // Browser storage may be unavailable.
      }
    }

    // Gate saving until the client-only restoration attempt has completed.
    setHasRestoredHistory(true);
  }, []);

  useEffect(() => {
    if (!hasRestoredHistory) return;

    if (skipNextSaveRef.current) {
      skipNextSaveRef.current = false;
      return;
    }

    const completedMessages = activeStreamingMessageIdRef.current
      ? messages.filter(
          (message) => message.id !== activeStreamingMessageIdRef.current,
        )
      : messages;

    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(completedMessages));
    } catch {
      // Keep chat usable if browser storage is unavailable.
    }
  }, [hasRestoredHistory, isLoading, messages]);

  const handleSend = async () => {
    const trimmedInput = input.trim();

    if (!trimmedInput || requestActiveRef.current) return;

    requestActiveRef.current = true;
    setIsLoading(true);
    setIsThinking(true);
    setError(null);

    const userMessage: Message = {
      id: crypto.randomUUID(),
      role: "user",
      content: trimmedInput,
    };
    const nextMessages = [...messages, userMessage];

    setMessages(nextMessages);
    setInput("");

    let assistantMessageId: string | null = null;

    try {
      const response = await fetch("/api/chat", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          messages: nextMessages.map(({ role, content }) => ({ role, content })),
        }),
      });

      if (!response.ok) {
        throw new Error("The chat request failed.");
      }

      if (!response.body) {
        throw new Error("The chat response did not include a stream.");
      }

      assistantMessageId = crypto.randomUUID();
      activeStreamingMessageIdRef.current = assistantMessageId;
      const assistantMessage: Message = {
        id: assistantMessageId,
        role: "assistant",
        content: THINKING_MESSAGE,
      };

      setMessages((prev) => [...prev, assistantMessage]);
      setIsThinking(false);

      const reader = response.body.getReader();
      const decoder = new TextDecoder();
      let assistantText = "";

      const appendChunk = (chunk: string) => {
        if (!chunk) return;

        assistantText += chunk;
        setMessages((prev) =>
          prev.map((message) =>
            message.id === assistantMessageId
              ? { ...message, content: assistantText }
              : message,
          ),
        );
      };

      while (true) {
        const { done, value } = await reader.read();

        if (done) {
          appendChunk(decoder.decode());
          break;
        }

        appendChunk(decoder.decode(value, { stream: true }));
      }

      if (!assistantText.trim()) {
        throw new Error("The chat response was empty.");
      }
    } catch {
      if (assistantMessageId) {
        setMessages((prev) =>
          prev.filter((message) => message.id !== assistantMessageId),
        );
      }

      setError("JARVIS could not respond. Please try again.");
    } finally {
      activeStreamingMessageIdRef.current = null;
      requestActiveRef.current = false;
      setIsLoading(false);
      setIsThinking(false);
    }
  };

  const clearConversation = () => {
    if (requestActiveRef.current) return;

    skipNextSaveRef.current = true;
    setMessages([]);
    setError(null);

    try {
      localStorage.removeItem(STORAGE_KEY);
    } catch {
      // Keep chat usable if browser storage is unavailable.
    }
  };

  return {
    input,
    setInput,
    isLoading,
    isThinking,
    messages,
    error,
    handleSend,
    clearConversation,
  };
}
