"use client";

import { useState } from "react";
import OutlookPanel from "@/components/OutlookPanel";
import type { Conversation } from "@/types/conversation";
import type { Memory } from "@/types/memory";

interface SidebarProps {
  conversations: Conversation[];
  activeConversationId: string | null;
  isLoading: boolean;
  onCreateConversation: () => void;
  onSwitchConversation: (id: string) => void;
  onRenameConversation: (id: string, title: string) => void;
  onDeleteConversation: (id: string) => void;
  memories: Memory[];
  onAddMemory: (content: string) => void;
  onDeleteMemory: (id: string) => void;
  onClearMemories: () => void;
}

export default function Sidebar({
  conversations,
  activeConversationId,
  isLoading,
  onCreateConversation,
  onSwitchConversation,
  onRenameConversation,
  onDeleteConversation,
  memories,
  onAddMemory,
  onDeleteMemory,
  onClearMemories,
}: SidebarProps) {
  const [editingId, setEditingId] = useState<string | null>(null);
  const [editingTitle, setEditingTitle] = useState("");
  const [memoryInput, setMemoryInput] = useState("");
  const sortedConversations = [...conversations].sort(
    (a, b) => Date.parse(b.updatedAt) - Date.parse(a.updatedAt),
  );

  const startRenaming = (conversation: Conversation) => {
    if (isLoading) return;
    setEditingId(conversation.id);
    setEditingTitle(conversation.title);
  };

  const finishRenaming = () => {
    if (!editingId) return;

    const trimmedTitle = editingTitle.trim();
    if (!trimmedTitle) return;

    onRenameConversation(editingId, trimmedTitle);
    setEditingId(null);
    setEditingTitle("");
  };

  const confirmDelete = (conversation: Conversation) => {
    if (
      !isLoading &&
      window.confirm(`Delete "${conversation.title}"?`)
    ) {
      if (editingId === conversation.id) {
        setEditingId(null);
        setEditingTitle("");
      }
      onDeleteConversation(conversation.id);
    }
  };

  return (
    <aside className="hidden w-64 shrink-0 overflow-y-auto border-r border-zinc-800 bg-zinc-950 p-4 md:block">
      <button
        type="button"
        className="w-full rounded-lg border border-zinc-700 px-4 py-2.5 text-left text-sm font-medium text-zinc-200 transition-colors hover:border-zinc-500 hover:bg-zinc-900 disabled:cursor-not-allowed disabled:opacity-50"
        onClick={onCreateConversation}
        disabled={isLoading}
      >
        + New Conversation
      </button>

      <nav className="mt-4 space-y-1" aria-label="Conversations">
        {sortedConversations.map((conversation) => {
          const isActive = conversation.id === activeConversationId;

          return editingId === conversation.id ? (
            <form
              key={conversation.id}
              className={`flex items-center gap-1 rounded-lg p-1 ${
                isActive ? "bg-zinc-800" : "bg-zinc-900"
              }`}
              onSubmit={(event) => {
                event.preventDefault();
                finishRenaming();
              }}
            >
              <input
                className="min-w-0 flex-1 rounded bg-zinc-950 px-2 py-1.5 text-sm text-white outline-none ring-1 ring-zinc-600 focus:ring-zinc-400"
                value={editingTitle}
                onChange={(event) => setEditingTitle(event.target.value)}
                aria-label={`Rename ${conversation.title}`}
                autoFocus
                disabled={isLoading}
              />
              <button
                type="submit"
                className="rounded px-2 py-1.5 text-xs text-zinc-200 hover:bg-zinc-700 disabled:opacity-50"
                disabled={isLoading || !editingTitle.trim()}
                aria-label="Save conversation title"
              >
                Save
              </button>
              <button
                type="button"
                className="rounded px-2 py-1.5 text-xs text-zinc-400 hover:bg-zinc-700 hover:text-white"
                onClick={() => setEditingId(null)}
                aria-label="Cancel rename"
              >
                Cancel
              </button>
            </form>
          ) : (
            <div
              key={conversation.id}
              className={`group flex items-center rounded-lg transition-colors ${
                isActive
                  ? "bg-zinc-800 text-white"
                  : "text-zinc-400 hover:bg-zinc-900 hover:text-zinc-200"
              }`}
            >
              <button
                type="button"
                className="min-w-0 flex-1 truncate px-3 py-2.5 text-left text-sm disabled:cursor-not-allowed disabled:opacity-50"
                onClick={() => onSwitchConversation(conversation.id)}
                disabled={isLoading}
                aria-current={isActive ? "page" : undefined}
                title={conversation.title}
              >
                {conversation.title}
              </button>
              <button
                type="button"
                className="rounded px-1.5 py-1 text-xs text-zinc-500 hover:bg-zinc-700 hover:text-white disabled:cursor-not-allowed disabled:opacity-50"
                onClick={() => startRenaming(conversation)}
                disabled={isLoading}
                aria-label={`Rename ${conversation.title}`}
                title="Rename"
              >
                Edit
              </button>
              <button
                type="button"
                className="mr-1 rounded px-1.5 py-1 text-xs text-zinc-500 hover:bg-red-950 hover:text-red-300 disabled:cursor-not-allowed disabled:opacity-50"
                onClick={() => confirmDelete(conversation)}
                disabled={isLoading}
                aria-label={`Delete ${conversation.title}`}
                title="Delete"
              >
                ×
              </button>
            </div>
          );
        })}
      </nav>

      <section className="mt-6 border-t border-zinc-800 pt-4">
        <div className="flex items-center justify-between gap-2">
          <h2 className="text-xs font-semibold uppercase tracking-wide text-zinc-500">
            Memory
          </h2>
          {memories.length > 0 ? (
            <button
              type="button"
              className="text-xs text-zinc-500 hover:text-red-300"
              onClick={() => {
                if (window.confirm("Clear all saved memories?")) {
                  onClearMemories();
                }
              }}
            >
              Clear all
            </button>
          ) : null}
        </div>

        <form
          className="mt-3 flex gap-1"
          onSubmit={(event) => {
            event.preventDefault();
            if (!memoryInput.trim()) return;
            onAddMemory(memoryInput);
            setMemoryInput("");
          }}
        >
          <input
            className="min-w-0 flex-1 rounded-md border border-zinc-700 bg-zinc-900 px-2 py-1.5 text-xs text-white outline-none placeholder:text-zinc-600 focus:border-zinc-500"
            value={memoryInput}
            onChange={(event) => setMemoryInput(event.target.value)}
            placeholder="Add a memory"
            aria-label="New memory"
          />
          <button
            type="submit"
            className="rounded-md border border-zinc-700 px-2 py-1.5 text-xs text-zinc-300 hover:border-zinc-500 hover:text-white disabled:cursor-not-allowed disabled:opacity-50"
            disabled={!memoryInput.trim()}
            aria-label="Save memory"
          >
            Add
          </button>
        </form>

        <ul className="mt-3 space-y-1">
          {memories.map((memory) => (
            <li
              key={memory.id}
              className="group flex items-start gap-1 rounded-md px-2 py-1.5 text-xs text-zinc-400 hover:bg-zinc-900"
            >
              <span className="min-w-0 flex-1 break-words">{memory.content}</span>
              <button
                type="button"
                className="shrink-0 rounded px-1 text-zinc-600 hover:bg-red-950 hover:text-red-300"
                onClick={() => onDeleteMemory(memory.id)}
                aria-label={`Delete memory: ${memory.content}`}
                title="Delete memory"
              >
                ×
              </button>
            </li>
          ))}
        </ul>
      </section>

      <OutlookPanel />
    </aside>
  );
}
