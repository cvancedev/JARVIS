"use client";

import { useEffect, useRef } from "react";
import { useVoiceInput } from "@/hooks/useVoiceInput";

interface PromptInputProps {
  value: string;
  onChange: (value: string) => void;
  onSend: () => void;
  isLoading: boolean;
}

export default function PromptInput({
  value,
  onChange,
  onSend,
  isLoading,
}: PromptInputProps) {
  const inputRef = useRef<HTMLInputElement>(null);
  const voiceBaseTextRef = useRef("");
  const previousVoiceStateRef = useRef("mic-off");
  const {
    isSupported: isVoiceSupported,
    isListening,
    isWakeWordEnabled,
    isAlwaysListening,
    isMicOff,
    voiceState,
    transcript,
    error: voiceError,
    startListening,
    stopListening,
    toggleWakeWord,
    toggleAlwaysListening,
    toggleMicOff,
    resetTranscript,
  } = useVoiceInput();

  useEffect(() => {
    if (!isLoading) {
      inputRef.current?.focus();
    }
  }, [isLoading]);

  useEffect(() => {
    if (
      voiceState === "command" &&
      previousVoiceStateRef.current !== "command"
    ) {
      voiceBaseTextRef.current = value;
    }
    previousVoiceStateRef.current = voiceState;
  }, [value, voiceState]);

  useEffect(() => {
    if (!transcript) return;

    const baseText = voiceBaseTextRef.current.trimEnd();
    onChange(`${baseText}${baseText ? " " : ""}${transcript}`);
  }, [onChange, transcript]);

  const toggleListening = () => {
    if (isWakeWordEnabled) return;

    if (isListening) {
      stopListening();
      return;
    }

    voiceBaseTextRef.current = value;
    resetTranscript();
    startListening();
  };

  return (
    <div className="mt-8">
      <div className="flex gap-3">
        <input
          className="min-w-0 flex-1 rounded-lg border border-zinc-700 bg-zinc-900 px-4 py-3 text-white outline-none disabled:opacity-60"
          placeholder="Type a command..."
          value={value}
          onChange={(event) => onChange(event.target.value)}
          onKeyDown={(event) => {
            if (event.key === "Enter") onSend();
          }}
          disabled={isLoading || isListening}
          ref={inputRef}
        />

        <button
          type="button"
          onClick={toggleListening}
          className={`rounded-lg border px-4 py-3 text-sm transition-colors disabled:cursor-not-allowed disabled:opacity-50 ${
            isListening
              ? "border-red-500 bg-red-950 text-red-300"
              : "border-zinc-700 bg-zinc-900 text-zinc-300 hover:border-zinc-500 hover:text-white"
          }`}
          disabled={
            isLoading || !isVoiceSupported || isWakeWordEnabled || isMicOff
          }
          aria-pressed={isListening}
          aria-label={isListening ? "Stop listening" : "Start voice input"}
          title={isVoiceSupported ? "Voice input" : "Voice input unavailable"}
        >
          {isListening ? "Listening…" : "Mic"}
        </button>

        <label
          className={`flex items-center gap-2 rounded-lg border border-zinc-800 px-3 text-xs ${
            isVoiceSupported ? "text-zinc-400" : "text-zinc-600"
          }`}
          title={isVoiceSupported ? undefined : "Wake word unavailable"}
        >
          <input
            type="checkbox"
            className="accent-blue-500"
            checked={isWakeWordEnabled}
            onChange={toggleWakeWord}
            disabled={isLoading || !isVoiceSupported || isMicOff}
          />
          Wake Word
        </label>

        <label
          className={`flex items-center gap-2 rounded-lg border border-zinc-800 px-3 text-xs ${
            isVoiceSupported && !isMicOff ? "text-zinc-400" : "text-zinc-600"
          }`}
          title={isVoiceSupported ? undefined : "Always Listening unavailable"}
        >
          <input
            type="checkbox"
            className="accent-blue-500"
            checked={isAlwaysListening}
            onChange={toggleAlwaysListening}
            disabled={isLoading || !isVoiceSupported || isMicOff}
          />
          Always Listening
        </label>

        <button
          type="button"
          className={`rounded-lg border px-3 py-3 text-xs transition-colors ${
            isMicOff
              ? "border-red-800 bg-red-950 text-red-300"
              : "border-zinc-700 bg-zinc-900 text-zinc-400 hover:text-white"
          }`}
          onClick={toggleMicOff}
          disabled={isLoading || !isVoiceSupported}
        >
          {isMicOff ? "Enable Mic" : "Mic Off"}
        </button>

        <button
          type="button"
          onClick={onSend}
          className="rounded-lg bg-blue-600 px-6 py-3 hover:bg-blue-500 disabled:cursor-not-allowed disabled:opacity-50"
          disabled={isLoading || isListening}
        >
          {isLoading ? "Thinking..." : "Send"}
        </button>
      </div>

      {isVoiceSupported ? (
        <p className="mt-2 text-xs text-zinc-500">
          {voiceState === "mic-off"
            ? "Mic Off"
            : voiceState === "push-to-talk"
              ? "Push-to-Talk"
            : voiceState === "wake-word"
              ? "Wake Word"
              : voiceState === "always-listening"
                ? "Always Listening"
              : "Listening for Command"}
        </p>
      ) : (
        <p className="mt-2 text-xs text-zinc-500">Voice input unavailable.</p>
      )}
      {voiceError ? (
        <p className="mt-2 text-xs text-red-400" role="alert">
          {voiceError}
        </p>
      ) : null}
    </div>
  );
}
