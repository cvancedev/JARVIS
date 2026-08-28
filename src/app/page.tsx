"use client";

import ChatWindow from "@/components/ChatWindow";
import CalendarProposalCard from "@/components/CalendarProposalCard";
import PromptInput from "@/components/PromptInput";
import Sidebar from "@/components/Sidebar";
import { useChat } from "@/hooks/useChat";
import { useMemory } from "@/hooks/useMemory";

export default function Home() {
  const { memories, addMemory, deleteMemory, clearMemories } = useMemory();
  const {
    input,
    setInput,
    isLoading,
    isThinking,
    messages,
    conversations,
    activeConversationId,
    error,
    handleSend,
    clearConversation,
    createConversation,
    switchConversation,
    renameConversation,
    deleteConversation,
    pendingMemorySuggestion,
    dismissMemorySuggestion,
    pendingCalendarProposal,
    isCreatingCalendarEvent,
    calendarCreationError,
    calendarCreationSuccess,
    createCalendarEvent,
    cancelCalendarProposal,
  } = useChat(memories);

  const handleClearConversation = () => {
    if (window.confirm("Clear this conversation?")) {
      clearConversation();
    }
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

          <div className="flex items-center gap-4">
            <button
              type="button"
              className="text-sm text-zinc-400 hover:text-white disabled:cursor-not-allowed disabled:opacity-50"
              onClick={handleClearConversation}
              disabled={isLoading}
            >
              Clear
            </button>
            <div className="flex items-center gap-2">
              <div className="h-3 w-3 animate-pulse rounded-full bg-green-500" />
              <span className="text-sm text-zinc-400">ONLINE</span>
            </div>
          </div>
        </div>
      </header>

      <div className="flex flex-1">
        <Sidebar
          conversations={conversations}
          activeConversationId={activeConversationId}
          isLoading={isLoading}
          onCreateConversation={createConversation}
          onSwitchConversation={switchConversation}
          onRenameConversation={renameConversation}
          onDeleteConversation={deleteConversation}
          memories={memories}
          onAddMemory={addMemory}
          onDeleteMemory={deleteMemory}
          onClearMemories={clearMemories}
        />

        <section className="mx-auto flex w-full max-w-5xl flex-col px-6 py-10">
          <div className="mb-8">
            <h2 className="text-4xl font-bold">Good Afternoon, Curt.</h2>
            <p className="mt-3 text-zinc-400">
              Awaiting your instructions...
            </p>
          </div>

          <div className="min-h-[420px] rounded-xl border border-zinc-800 bg-zinc-900/40 p-6">
            <ChatWindow
              messages={messages}
              isThinking={isThinking}
              isStreaming={isLoading}
            />
          </div>

          {pendingMemorySuggestion ? (
            <div className="mt-3 flex flex-wrap items-center gap-3 rounded-lg border border-amber-900/60 bg-amber-950/20 px-4 py-3 text-sm">
              <p className="min-w-0 flex-1 text-zinc-300">
                <span className="font-medium text-amber-300">
                  Save to memory?
                </span>{" "}
                {pendingMemorySuggestion.content}
              </p>
              <div className="flex gap-2">
                <button
                  type="button"
                  className="rounded-md bg-amber-500 px-3 py-1.5 text-xs font-medium text-zinc-950 hover:bg-amber-400"
                  onClick={() => {
                    if (addMemory(pendingMemorySuggestion.content)) {
                      dismissMemorySuggestion(pendingMemorySuggestion.id);
                    }
                  }}
                >
                  Save
                </button>
                <button
                  type="button"
                  className="rounded-md border border-zinc-700 px-3 py-1.5 text-xs text-zinc-300 hover:border-zinc-500 hover:text-white"
                  onClick={() =>
                    dismissMemorySuggestion(pendingMemorySuggestion.id)
                  }
                >
                  Dismiss
                </button>
              </div>
            </div>
          ) : null}

          {pendingCalendarProposal ? (
            <CalendarProposalCard
              proposal={pendingCalendarProposal}
              isCreating={isCreatingCalendarEvent}
              error={calendarCreationError}
              onCreate={() => void createCalendarEvent()}
              onCancel={cancelCalendarProposal}
            />
          ) : null}

          {calendarCreationSuccess ? (
            <p className="mt-3 rounded-lg border border-emerald-900 bg-emerald-950/30 px-4 py-3 text-sm text-emerald-300">
              {calendarCreationSuccess}
            </p>
          ) : null}

          <PromptInput
            value={input}
            onChange={setInput}
            onSend={handleSend}
            isLoading={isLoading || isCreatingCalendarEvent}
          />
          {error ? (
            <p className="mt-3 text-sm text-red-400" role="alert">{error}</p>
          ) : null}
        </section>
      </div>
    </main>
  );
}
