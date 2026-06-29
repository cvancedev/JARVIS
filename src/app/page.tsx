"use client";

import { useState } from "react";
import ChatWindow from "@/components/ChatWindow";
import PromptInput from "@/components/PromptInput";
import { Message } from "@/types/message";

export default function Home() {
  const [input, setInput] = useState("");
  const [isLoading, setIsLoading] = useState(false);
  const [messages, setMessages] = useState<Message[]>([
    {
      id: "1",
      role: "assistant",
      content: "Welcome back, Curt. Awaiting instructions.",
    },
  ]);

  const handleSend = async () => {
  const trimmedInput = input.trim();

  if (!trimmedInput || isLoading) return;

  setIsLoading(true);

  const userMessage: Message = {
    id: crypto.randomUUID(),
    role: "user",
    content: trimmedInput,
  };

  setMessages((prev) => [...prev, userMessage]);

  setInput("");

  // Fake thinking delay
  setTimeout(() => {
    const assistantMessage: Message = {
      id: crypto.randomUUID(),
      role: "assistant",
      content: "Command received. AI connection coming next.",
    };

    setMessages((prev) => [...prev, assistantMessage]);

    setIsLoading(false);
  }, 700);
};

  return (
    <main className="min-h-screen bg-zinc-950 text-white">
      <header className="border-b border-zinc-800 p-5">
        <div className="mx-auto flex max-w-5xl items-center justify-between">
          <div>
            <h1 className="text-3xl font-bold">JARVIS</h1>
            <p className="text-sm text-zinc-400">
              Personal AI Operating System
            </p>
          </div>

          <div className="flex items-center gap-2">
            <div className="h-3 w-3 animate-pulse rounded-full bg-green-500" />
            <span className="text-sm text-zinc-400">ONLINE</span>
          </div>
        </div>
      </header>

      <section className="mx-auto flex max-w-5xl flex-col px-6 py-10">
        <div className="mb-8">
          <h2 className="text-4xl font-bold">Good Afternoon, Curt.</h2>
          <p className="mt-3 text-zinc-400">
            Awaiting your instructions...
          </p>
        </div>

        <div className="min-h-[420px] rounded-xl border border-zinc-800 bg-zinc-900/40 p-6">
          <ChatWindow messages={messages} />
        </div>

        <PromptInput
          value={input}
          onChange={setInput}
          onSend={handleSend}
          isLoading={isLoading}
        />
      </section>
    </main>
  );
}