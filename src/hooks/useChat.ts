"use client";

import { useEffect, useRef, useState } from "react";
import type { Conversation } from "@/types/conversation";
import type {
  CalendarEventProposal,
  CalendarProposalConflict,
} from "@/types/calendarProposal";
import type { Memory } from "@/types/memory";
import type { MemorySuggestion } from "@/types/memorySuggestion";
import type { Message } from "@/types/message";

const CONVERSATIONS_STORAGE_KEY = "jarvis-conversations";
const ACTIVE_CONVERSATION_STORAGE_KEY = "jarvis-active-conversation-id";
const LEGACY_STORAGE_KEY = "jarvis-chat-history";
const THINKING_MESSAGE = "Thinking...";
const DEFAULT_TITLE = "New Conversation";
const MAX_TITLE_LENGTH = 48;
const CALENDAR_PROPOSAL_CONTROL = "\n\u001eJARVIS_CALENDAR_PROPOSAL:";
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

function isStoredConversation(value: unknown): value is Conversation {
  if (typeof value !== "object" || value === null) return false;

  const conversation = value as Record<string, unknown>;

  return (
    typeof conversation.id === "string" &&
    conversation.id.trim() !== "" &&
    typeof conversation.title === "string" &&
    conversation.title.trim() !== "" &&
    Array.isArray(conversation.messages) &&
    conversation.messages.every(isStoredMessage) &&
    typeof conversation.createdAt === "string" &&
    !Number.isNaN(Date.parse(conversation.createdAt)) &&
    typeof conversation.updatedAt === "string" &&
    !Number.isNaN(Date.parse(conversation.updatedAt))
  );
}

function isCalendarProposalConflict(value: unknown): value is CalendarProposalConflict {
  if (typeof value !== "object" || value === null) return false;
  const conflict = value as Record<string, unknown>;
  return typeof conflict.id === "string" &&
    typeof conflict.subject === "string" &&
    typeof conflict.start === "string" &&
    typeof conflict.end === "string" &&
    typeof conflict.isAllDay === "boolean";
}

function isCalendarEventProposal(value: unknown): value is CalendarEventProposal {
  if (typeof value !== "object" || value === null) return false;
  const proposal = value as Record<string, unknown>;
  return typeof proposal.id === "string" &&
    typeof proposal.subject === "string" &&
    typeof proposal.start === "string" &&
    typeof proposal.end === "string" &&
    typeof proposal.timeZone === "string" &&
    typeof proposal.location === "string" &&
    Array.isArray(proposal.attendeeEmails) &&
    proposal.attendeeEmails.every((email) => typeof email === "string") &&
    typeof proposal.description === "string" &&
    Array.isArray(proposal.conflicts) &&
    proposal.conflicts.every(isCalendarProposalConflict);
}

function parseStoredArray(
  value: string,
  validator: (item: unknown) => boolean,
): unknown[] | null {
  try {
    const parsed: unknown = JSON.parse(value);
    return Array.isArray(parsed) && parsed.every(validator) ? parsed : null;
  } catch {
    return null;
  }
}

function createConversation(messages: Message[] = INITIAL_MESSAGES): Conversation {
  const now = new Date().toISOString();

  return {
    id: crypto.randomUUID(),
    title: DEFAULT_TITLE,
    messages,
    createdAt: now,
    updatedAt: now,
  };
}

function removeIncompleteMessages(conversation: Conversation): Conversation {
  return {
    ...conversation,
    messages: conversation.messages.filter(
      (message) =>
        message.role !== "assistant" || message.content !== THINKING_MESSAGE,
    ),
  };
}

function capitalizeTitle(title: string): string {
  const minorWords = new Set([
    "a",
    "an",
    "and",
    "as",
    "at",
    "by",
    "for",
    "from",
    "in",
    "of",
    "on",
    "or",
    "the",
    "to",
    "with",
  ]);

  return title
    .split(" ")
    .map((word, index) =>
      word
        .split("-")
        .map((part, partIndex) => {
          if (/[a-z][A-Z]|[A-Z].*[A-Z]/.test(part)) return part;

          const lowercasePart = part.toLowerCase();
          if (index > 0 && partIndex === 0 && minorWords.has(lowercasePart)) {
            return lowercasePart;
          }

          return lowercasePart.charAt(0).toUpperCase() + lowercasePart.slice(1);
        })
        .join("-"),
    )
    .join(" ");
}

function generateConversationTitle(message: string): string {
  let title = message.replace(/\s+/g, " ").trim().replace(/[.!?,;:]+$/, "");

  title = title
    .replace(
      /^(?:please\s+)?help me (?:write|create|draft|make|build)\s+(?:an?\s+|the\s+)?/i,
      "",
    )
    .replace(/^how (?:do|can|should) i\s+/i, "")
    .replace(
      /^(?:please\s+)?(?:write|create|draft|make|build)\s+(?:me\s+)?(?:an?\s+|the\s+)?/i,
      "",
    )
    .trim();

  const aboutIndex = title.toLowerCase().indexOf(" about ");
  if (aboutIndex >= 18) title = title.slice(0, aboutIndex);

  let wasShortened = false;
  if (title.length > MAX_TITLE_LENGTH) {
    const candidate = title.slice(0, MAX_TITLE_LENGTH + 1);
    const lastSpace = candidate.lastIndexOf(" ");
    title = title.slice(0, lastSpace > 0 ? lastSpace : MAX_TITLE_LENGTH);
    wasShortened = true;
  }

  if (wasShortened) title = title.replace(/[.!?,;:\-]+$/, "");

  return capitalizeTitle(title.trim()) || DEFAULT_TITLE;
}

function detectMemorySuggestion(message: string): string | null {
  const normalized = message.replace(/\s+/g, " ").trim();
  const sensitivePattern =
    /\b(api[ -]?key|password|passcode|secret|access[ -]?token|refresh[ -]?token|bearer|private[ -]?key|credit[ -]?card|cvv|social security|ssn)\b|\bsk-[a-z0-9_-]+|\b[A-Za-z0-9_-]{32,}\b/i;
  const temporaryPattern =
    /\b(today|tomorrow|tonight|this (?:morning|afternoon|evening|week|month)|next (?:hour|week|month)|for now|just this once|temporarily)\b/i;

  if (sensitivePattern.test(normalized) || temporaryPattern.test(normalized)) {
    return null;
  }

  const rememberMatch = normalized.match(/^remember(?: that)?[,:]?\s+(.+)$/i);
  const hasDurableIntent =
    rememberMatch ||
    /^(?:from now on|going forward)[,:]?\s+.+$/i.test(normalized) ||
    /^(?:my preference is|i prefer)\s+.+$/i.test(normalized);

  if (!hasDurableIntent) return null;

  const content = (rememberMatch?.[1] ?? normalized)
    .trim()
    .replace(/[.!?]+$/, "");
  const words = content.split(/\s+/);
  const oneTimeTaskPattern =
    /^to\s+(?:call|email|send|buy|pick up|schedule|book|submit|pay|cancel|order|complete|finish)\b/i;

  if (
    content.length < 12 ||
    words.length < 3 ||
    oneTimeTaskPattern.test(content)
  ) {
    return null;
  }

  return content;
}

export function useChat(memories: Memory[] = []) {
  const [input, setInput] = useState("");
  const [isLoading, setIsLoading] = useState(false);
  const [isThinking, setIsThinking] = useState(false);
  const [hasRestoredHistory, setHasRestoredHistory] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [conversations, setConversations] = useState<Conversation[]>([]);
  const [memorySuggestions, setMemorySuggestions] = useState<
    MemorySuggestion[]
  >([]);
  const [calendarProposalState, setCalendarProposalState] = useState<{
    conversationId: string;
    proposal: CalendarEventProposal;
  } | null>(null);
  const [isCreatingCalendarEvent, setIsCreatingCalendarEvent] = useState(false);
  const [calendarCreationError, setCalendarCreationError] = useState<string | null>(null);
  const [calendarCreationSuccess, setCalendarCreationSuccess] = useState<string | null>(null);
  const [activeConversationId, setActiveConversationId] = useState<string | null>(
    null,
  );
  const requestActiveRef = useRef(false);
  const activeStreamingMessageIdRef = useRef<string | null>(null);
  const calendarCreationActiveRef = useRef(false);

  const activeConversation = conversations.find(
    (conversation) => conversation.id === activeConversationId,
  );
  const messages = activeConversation?.messages ?? INITIAL_MESSAGES;
  const pendingMemorySuggestion =
    memorySuggestions.find(
      (suggestion) => suggestion.conversationId === activeConversationId,
    ) ?? null;
  const pendingCalendarProposal =
    calendarProposalState?.conversationId === activeConversationId
      ? calendarProposalState.proposal
      : null;

  const persistConversationState = (
    nextConversations: Conversation[],
    nextActiveConversationId: string,
  ) => {
    try {
      localStorage.setItem(
        CONVERSATIONS_STORAGE_KEY,
        JSON.stringify(nextConversations),
      );
      localStorage.setItem(
        ACTIVE_CONVERSATION_STORAGE_KEY,
        nextActiveConversationId,
      );
    } catch {
      // Keep chat usable if browser storage is unavailable.
    }
  };

  useEffect(() => {
    let restoredConversations: Conversation[] | null = null;
    let restoredActiveId: string | null = null;

    try {
      const savedConversations = localStorage.getItem(CONVERSATIONS_STORAGE_KEY);

      if (savedConversations) {
        const parsedConversations = parseStoredArray(
          savedConversations,
          isStoredConversation,
        );

        if (parsedConversations && parsedConversations.length > 0) {
          restoredConversations = (
            parsedConversations as Conversation[]
          ).map(removeIncompleteMessages);
          const savedActiveId = localStorage.getItem(
            ACTIVE_CONVERSATION_STORAGE_KEY,
          );
          restoredActiveId = restoredConversations.some(
            (conversation) => conversation.id === savedActiveId,
          )
            ? savedActiveId
            : restoredConversations[0].id;
        } else {
          localStorage.removeItem(CONVERSATIONS_STORAGE_KEY);
          localStorage.removeItem(ACTIVE_CONVERSATION_STORAGE_KEY);
        }
      }

      if (!restoredConversations) {
        const savedHistory = localStorage.getItem(LEGACY_STORAGE_KEY);
        let migratedMessages: Message[] | null = null;

        if (savedHistory) {
          const parsedHistory = parseStoredArray(savedHistory, isStoredMessage);

          if (parsedHistory) {
            migratedMessages = (parsedHistory as Message[]).filter(
              (message) =>
                message.role !== "assistant" ||
                message.content !== THINKING_MESSAGE,
            );
          } else {
            localStorage.removeItem(LEGACY_STORAGE_KEY);
          }
        }

        const conversation = createConversation(
          migratedMessages ?? INITIAL_MESSAGES,
        );
        restoredConversations = [conversation];
        restoredActiveId = conversation.id;

        if (migratedMessages) {
          localStorage.setItem(
            CONVERSATIONS_STORAGE_KEY,
            JSON.stringify(restoredConversations),
          );
          localStorage.setItem(ACTIVE_CONVERSATION_STORAGE_KEY, restoredActiveId);
          localStorage.removeItem(LEGACY_STORAGE_KEY);
        }
      }
    } catch {
      try {
        localStorage.removeItem(CONVERSATIONS_STORAGE_KEY);
        localStorage.removeItem(ACTIVE_CONVERSATION_STORAGE_KEY);
      } catch {
        // Browser storage may be unavailable.
      }

      const conversation = createConversation();
      restoredConversations = [conversation];
      restoredActiveId = conversation.id;
    }

    // Client-only storage cannot be read during the server render.
    // eslint-disable-next-line react-hooks/set-state-in-effect
    setConversations(restoredConversations);
    setActiveConversationId(restoredActiveId);
    setHasRestoredHistory(true);
  }, []);

  useEffect(() => {
    if (!hasRestoredHistory || !activeConversationId) return;

    const conversationsToStore = conversations.map((conversation) => ({
      ...conversation,
      messages: activeStreamingMessageIdRef.current
        ? conversation.messages.filter(
            (message) => message.id !== activeStreamingMessageIdRef.current,
          )
        : conversation.messages,
    }));

    try {
      localStorage.setItem(
        CONVERSATIONS_STORAGE_KEY,
        JSON.stringify(conversationsToStore),
      );
      localStorage.setItem(
        ACTIVE_CONVERSATION_STORAGE_KEY,
        activeConversationId,
      );
    } catch {
      // Keep chat usable if browser storage is unavailable.
    }
  }, [activeConversationId, conversations, hasRestoredHistory, isLoading]);

  const setConversationMessages = (
    conversationId: string,
    update: Message[] | ((previous: Message[]) => Message[]),
  ) => {
    setConversations((previous) =>
      previous.map((conversation) => {
        if (conversation.id !== conversationId) return conversation;

        const nextMessages =
          typeof update === "function" ? update(conversation.messages) : update;

        return {
          ...conversation,
          messages: nextMessages,
          updatedAt: new Date().toISOString(),
        };
      }),
    );
  };

  const handleSend = async () => {
    const trimmedInput = input.trim();

    if (!trimmedInput || requestActiveRef.current || !activeConversation) return;

    const requestConversationId = activeConversation.id;

    requestActiveRef.current = true;
    setIsLoading(true);
    setIsThinking(true);
    setError(null);
    setCalendarProposalState((previous) =>
      previous?.conversationId === requestConversationId ? null : previous,
    );
    setCalendarCreationError(null);
    setCalendarCreationSuccess(null);

    const userMessage: Message = {
      id: crypto.randomUUID(),
      role: "user",
      content: trimmedInput,
    };
    const nextMessages = [...messages, userMessage];
    const suggestedMemory = detectMemorySuggestion(trimmedInput);

    if (
      suggestedMemory &&
      !memories.some((memory) => memory.content === suggestedMemory)
    ) {
      setMemorySuggestions((previous) => [
        ...previous.filter(
          (suggestion) =>
            suggestion.conversationId !== requestConversationId,
        ),
        {
          id: crypto.randomUUID(),
          content: suggestedMemory,
          sourceMessageId: userMessage.id,
          conversationId: requestConversationId,
          createdAt: new Date().toISOString(),
        },
      ]);
    }

    const isFirstUserMessage = !activeConversation.messages.some(
      (message) => message.role === "user",
    );

    if (activeConversation.title === DEFAULT_TITLE && isFirstUserMessage) {
      const now = new Date().toISOString();
      const nextConversations = conversations.map((conversation) =>
        conversation.id === requestConversationId
          ? {
              ...conversation,
              title: generateConversationTitle(trimmedInput),
              messages: nextMessages,
              updatedAt: now,
            }
          : conversation,
      );

      setConversations(nextConversations);
      persistConversationState(nextConversations, requestConversationId);
    } else {
      setConversationMessages(requestConversationId, nextMessages);
    }
    setInput("");

    let assistantMessageId: string | null = null;

    try {
      const response = await fetch("/api/chat", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          messages: nextMessages.map(({ role, content }) => ({ role, content })),
          memories: memories.map(({ content }) => ({ content })),
          timeZone: Intl.DateTimeFormat().resolvedOptions().timeZone,
        }),
      });

      if (!response.ok) throw new Error("The chat request failed.");
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

      setConversationMessages(requestConversationId, (previous) => [
        ...previous,
        assistantMessage,
      ]);
      setIsThinking(false);

      const reader = response.body.getReader();
      const decoder = new TextDecoder();
      let assistantText = "";
      let responseBuffer = "";
      let proposalPayload = "";
      let readingProposal = false;

      const appendChunk = (chunk: string) => {
        if (!chunk) return;

        assistantText += chunk;
        setConversationMessages(requestConversationId, (previous) =>
          previous.map((message) =>
            message.id === assistantMessageId
              ? { ...message, content: assistantText }
              : message,
          ),
        );
      };

      const processResponseChunk = (chunk: string, isFinal = false) => {
        if (readingProposal) {
          proposalPayload += chunk;
          return;
        }
        responseBuffer += chunk;
        const controlIndex = responseBuffer.indexOf(CALENDAR_PROPOSAL_CONTROL);
        if (controlIndex >= 0) {
          appendChunk(responseBuffer.slice(0, controlIndex));
          proposalPayload = responseBuffer.slice(
            controlIndex + CALENDAR_PROPOSAL_CONTROL.length,
          );
          responseBuffer = "";
          readingProposal = true;
          return;
        }
        const safeLength = isFinal
          ? responseBuffer.length
          : Math.max(0, responseBuffer.length - CALENDAR_PROPOSAL_CONTROL.length + 1);
        appendChunk(responseBuffer.slice(0, safeLength));
        responseBuffer = responseBuffer.slice(safeLength);
      };

      while (true) {
        const { done, value } = await reader.read();

        if (done) {
          processResponseChunk(decoder.decode(), true);
          break;
        }

        processResponseChunk(decoder.decode(value, { stream: true }));
      }

      if (readingProposal && proposalPayload) {
        try {
          const proposal: unknown = JSON.parse(proposalPayload);
          if (isCalendarEventProposal(proposal)) {
            setCalendarProposalState({
              conversationId: requestConversationId,
              proposal,
            });
          }
        } catch {
          // Ignore malformed control data without affecting the chat response.
        }
      }

      if (!assistantText.trim()) {
        throw new Error("The chat response was empty.");
      }
    } catch {
      if (assistantMessageId) {
        setConversationMessages(requestConversationId, (previous) =>
          previous.filter((message) => message.id !== assistantMessageId),
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
    if (requestActiveRef.current || calendarCreationActiveRef.current || !activeConversation) return;

    setConversationMessages(activeConversation.id, []);
    setMemorySuggestions((previous) =>
      previous.filter(
        (suggestion) => suggestion.conversationId !== activeConversation.id,
      ),
    );
    setError(null);
    setCalendarProposalState((previous) =>
      previous?.conversationId === activeConversation.id ? null : previous,
    );
    setCalendarCreationError(null);
    setCalendarCreationSuccess(null);
  };

  const createNewConversation = () => {
    if (requestActiveRef.current || calendarCreationActiveRef.current) return;

    const conversation = createConversation([]);
    const nextConversations = [conversation, ...conversations];

    setConversations(nextConversations);
    setActiveConversationId(conversation.id);
    setError(null);
    setCalendarCreationError(null);
    setCalendarCreationSuccess(null);
    persistConversationState(nextConversations, conversation.id);
  };

  const switchConversation = (id: string) => {
    if (
      requestActiveRef.current ||
      calendarCreationActiveRef.current ||
      id === activeConversationId ||
      !conversations.some((conversation) => conversation.id === id)
    ) {
      return;
    }

    setActiveConversationId(id);
    setError(null);
    setCalendarCreationError(null);
    setCalendarCreationSuccess(null);

    try {
      localStorage.setItem(ACTIVE_CONVERSATION_STORAGE_KEY, id);
    } catch {
      // Keep chat usable if browser storage is unavailable.
    }
  };

  const renameConversation = (id: string, title: string) => {
    const trimmedTitle = title.trim();

    if (
      requestActiveRef.current ||
      calendarCreationActiveRef.current ||
      !trimmedTitle ||
      !conversations.some((conversation) => conversation.id === id) ||
      !activeConversationId
    ) {
      return;
    }

    const nextConversations = conversations.map((conversation) =>
      conversation.id === id
        ? {
            ...conversation,
            title: trimmedTitle,
            updatedAt: new Date().toISOString(),
          }
        : conversation,
    );

    setConversations(nextConversations);
    persistConversationState(nextConversations, activeConversationId);
  };

  const deleteConversation = (id: string) => {
    if (
      requestActiveRef.current ||
      calendarCreationActiveRef.current ||
      !conversations.some((conversation) => conversation.id === id)
    ) {
      return;
    }

    let nextConversations = conversations.filter(
      (conversation) => conversation.id !== id,
    );
    let nextActiveConversationId = activeConversationId;

    if (nextConversations.length === 0) {
      const conversation = createConversation([]);
      nextConversations = [conversation];
      nextActiveConversationId = conversation.id;
    } else if (id === activeConversationId) {
      nextActiveConversationId = [...nextConversations].sort(
        (a, b) => Date.parse(b.updatedAt) - Date.parse(a.updatedAt),
      )[0].id;
    }

    if (!nextActiveConversationId) return;

    setConversations(nextConversations);
    setMemorySuggestions((previous) =>
      previous.filter((suggestion) => suggestion.conversationId !== id),
    );
    setCalendarProposalState((previous) =>
      previous?.conversationId === id ? null : previous,
    );
    setActiveConversationId(nextActiveConversationId);
    setError(null);
    persistConversationState(nextConversations, nextActiveConversationId);
  };

  const dismissMemorySuggestion = (id: string) => {
    setMemorySuggestions((previous) =>
      previous.filter((suggestion) => suggestion.id !== id),
    );
  };

  const cancelCalendarProposal = () => {
    if (calendarCreationActiveRef.current) return;
    setCalendarProposalState(null);
    setCalendarCreationError(null);
  };

  const createCalendarEvent = async () => {
    if (!pendingCalendarProposal || calendarCreationActiveRef.current) return;
    calendarCreationActiveRef.current = true;
    setIsCreatingCalendarEvent(true);
    setCalendarCreationError(null);
    setCalendarCreationSuccess(null);
    try {
      const response = await fetch("/api/calendar/create", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ proposal: pendingCalendarProposal }),
      });
      const body: unknown = await response.json();
      if (
        response.status === 409 &&
        typeof body === "object" &&
        body !== null &&
        "conflicts" in body &&
        Array.isArray(body.conflicts) &&
        body.conflicts.every(isCalendarProposalConflict)
      ) {
        const updatedConflicts = body.conflicts as CalendarProposalConflict[];
        setCalendarProposalState((previous) => previous ? {
          ...previous,
          proposal: { ...previous.proposal, conflicts: updatedConflicts },
        } : previous);
        setCalendarCreationError(
          "Your calendar changed. Review the updated conflicts, then click Create Event again if you still want to proceed.",
        );
        return;
      }
      if (!response.ok) {
        const permissionRequired = typeof body === "object" && body !== null &&
          "code" in body && body.code === "calendar_write_permission_required";
        throw new Error(permissionRequired ? "permission" : "creation");
      }

      setCalendarProposalState(null);
      setCalendarCreationSuccess("Calendar event created successfully.");
    } catch (creationError: unknown) {
      setCalendarCreationError(
        creationError instanceof Error && creationError.message === "permission"
          ? "Microsoft needs delegated Calendars.ReadWrite permission. Reconnect Microsoft after granting it."
          : "The calendar event could not be created. Your proposal was preserved.",
      );
    } finally {
      calendarCreationActiveRef.current = false;
      setIsCreatingCalendarEvent(false);
    }
  };

  return {
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
    createConversation: createNewConversation,
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
  };
}
