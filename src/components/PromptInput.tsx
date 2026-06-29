"use client";

import { useEffect, useRef } from "react";

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

  useEffect(() => {
    if (!isLoading) {
      inputRef.current?.focus();
    }
  }, [isLoading]);
  
  return (
    <div className="mt-8 flex gap-3">
      <input
        className="flex-1 rounded-lg border border-zinc-700 bg-zinc-900 px-4 py-3 text-white outline-none"
        placeholder="Type a command..."
        value={value}
        onChange={(e) => onChange(e.target.value)}
        onKeyDown={(e) => {
          if (e.key === "Enter") onSend();
        }}
        disabled={isLoading}
        ref={inputRef}

      />

      <button
        onClick={onSend}
        className="rounded-lg bg-blue-600 px-6 py-3 hover:bg-blue-500"
        disabled={isLoading}
      >
        {isLoading ? "Thinking..." : "Send"}
      </button>
    </div>
  );
}