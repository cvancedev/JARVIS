"use client";

import { useEffect, useState } from "react";
import type { Memory } from "@/types/memory";

const MEMORY_STORAGE_KEY = "jarvis-memories";

function isStoredMemory(value: unknown): value is Memory {
  if (typeof value !== "object" || value === null) return false;

  const memory = value as Record<string, unknown>;

  return (
    typeof memory.id === "string" &&
    memory.id.trim() !== "" &&
    typeof memory.content === "string" &&
    memory.content.trim() !== "" &&
    typeof memory.createdAt === "string" &&
    !Number.isNaN(Date.parse(memory.createdAt)) &&
    typeof memory.updatedAt === "string" &&
    !Number.isNaN(Date.parse(memory.updatedAt))
  );
}

function persistMemories(memories: Memory[]) {
  try {
    localStorage.setItem(MEMORY_STORAGE_KEY, JSON.stringify(memories));
  } catch {
    // Keep memory controls usable if browser storage is unavailable.
  }
}

export function useMemory() {
  const [memories, setMemories] = useState<Memory[]>([]);

  useEffect(() => {
    try {
      const storedMemories = localStorage.getItem(MEMORY_STORAGE_KEY);
      if (!storedMemories) return;

      const parsedMemories: unknown = JSON.parse(storedMemories);
      if (
        !Array.isArray(parsedMemories) ||
        !parsedMemories.every(isStoredMemory)
      ) {
        localStorage.removeItem(MEMORY_STORAGE_KEY);
        return;
      }

      // Client-only storage cannot be read during the server render.
      // eslint-disable-next-line react-hooks/set-state-in-effect
      setMemories(parsedMemories);
    } catch {
      try {
        localStorage.removeItem(MEMORY_STORAGE_KEY);
      } catch {
        // Browser storage may be unavailable.
      }
    }
  }, []);

  const addMemory = (content: string) => {
    const trimmedContent = content.trim();

    if (
      !trimmedContent ||
      memories.some((memory) => memory.content === trimmedContent)
    ) {
      return false;
    }

    const now = new Date().toISOString();
    const memory: Memory = {
      id: crypto.randomUUID(),
      content: trimmedContent,
      createdAt: now,
      updatedAt: now,
    };
    const nextMemories = [memory, ...memories];

    setMemories(nextMemories);
    persistMemories(nextMemories);
    return true;
  };

  const deleteMemory = (id: string) => {
    if (!memories.some((memory) => memory.id === id)) return;

    const nextMemories = memories.filter((memory) => memory.id !== id);
    setMemories(nextMemories);
    persistMemories(nextMemories);
  };

  const clearMemories = () => {
    setMemories([]);
    persistMemories([]);
  };

  return {
    memories,
    addMemory,
    deleteMemory,
    clearMemories,
  };
}
