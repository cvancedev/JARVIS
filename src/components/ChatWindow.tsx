"use client";

import { useEffect, useRef, useState } from "react";
import Message from "./Message";
import { useSpeechOutput } from "@/hooks/useSpeechOutput";
import { Message as MessageType } from "@/types/message";

const thinkingMessage: MessageType = {
  id: "thinking",
  role: "assistant",
  content: "Thinking...",
};

const AUTO_SPEAK_STORAGE_KEY = "jarvis-auto-speak";

interface ChatWindowProps {
  messages: MessageType[];
  isThinking: boolean;
  isStreaming: boolean;
}

export default function ChatWindow({
  messages,
  isThinking,
  isStreaming,
}: ChatWindowProps) {
  const [speakingMessageId, setSpeakingMessageId] = useState<string | null>(null);
  const [autoSpeak, setAutoSpeak] = useState(false);
  const { isSupported, isSpeaking, speak, stop } = useSpeechOutput();
  const previousIsStreamingRef = useRef(false);
  const streamingAssistantIdRef = useRef<string | null>(null);
  const autoSpokenMessageIdsRef = useRef(new Set<string>());
  const streamingMessageId =
    isStreaming && !isThinking
      ? messages.findLast((message) => message.role === "assistant")?.id ?? null
      : null;

  useEffect(() => {
    try {
      if (localStorage.getItem(AUTO_SPEAK_STORAGE_KEY) === "true") {
        // Client-only storage cannot be read during the server render.
        // eslint-disable-next-line react-hooks/set-state-in-effect
        setAutoSpeak(true);
      }
    } catch {
      // Keep speech controls usable if browser storage is unavailable.
    }
  }, []);

  useEffect(() => {
    if (isStreaming && !isThinking && streamingMessageId) {
      streamingAssistantIdRef.current = streamingMessageId;
    }

    if (previousIsStreamingRef.current && !isStreaming) {
      const completedMessageId = streamingAssistantIdRef.current;
      const completedMessage = messages.find(
        (message) => message.id === completedMessageId,
      );

      if (
        autoSpeak &&
        isSupported &&
        completedMessage?.role === "assistant" &&
        !autoSpokenMessageIdsRef.current.has(completedMessage.id)
      ) {
        autoSpokenMessageIdsRef.current.add(completedMessage.id);
        setSpeakingMessageId(completedMessage.id);
        speak(completedMessage.content);
      }

      streamingAssistantIdRef.current = null;
    }

    previousIsStreamingRef.current = isStreaming;
  }, [autoSpeak, isStreaming, isSupported, isThinking, messages, speak, streamingMessageId]);

  const toggleAutoSpeak = () => {
    const nextAutoSpeak = !autoSpeak;

    if (!nextAutoSpeak) stop();
    setAutoSpeak(nextAutoSpeak);

    try {
      localStorage.setItem(AUTO_SPEAK_STORAGE_KEY, String(nextAutoSpeak));
    } catch {
      // Keep speech controls usable if browser storage is unavailable.
    }
  };

  const handleSpeak = (message: MessageType) => {
    if (isSpeaking && speakingMessageId === message.id) {
      stop();
      return;
    }

    setSpeakingMessageId(message.id);
    speak(message.content);
  };

  return (
    <div className="space-y-4">
      <div className="flex justify-end">
        <label
          className={`flex items-center gap-2 text-xs ${
            isSupported ? "text-zinc-400" : "text-zinc-600"
          }`}
          title={isSupported ? undefined : "Speech output unavailable"}
        >
          <input
            type="checkbox"
            className="accent-blue-500"
            checked={autoSpeak}
            onChange={toggleAutoSpeak}
            disabled={!isSupported}
          />
          Auto-Speak
        </label>
      </div>
      {messages.map((message) => (
        <Message
          key={message.id}
          message={message}
          canSpeak={
            isSupported &&
            message.role === "assistant" &&
            message.id !== streamingMessageId
          }
          isSpeaking={isSpeaking && speakingMessageId === message.id}
          onSpeak={() => handleSpeak(message)}
        />
      ))}
      {isThinking ? <Message message={thinkingMessage} /> : null}
    </div>
  );
}
