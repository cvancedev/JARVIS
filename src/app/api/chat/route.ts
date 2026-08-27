import OpenAI from "openai";

export const runtime = "nodejs";

const MODEL = "gpt-5.6";
const MAX_MEMORY_COUNT = 100;
const MAX_MEMORY_LENGTH = 1_000;
const MAX_TOTAL_MEMORY_LENGTH = 10_000;

type ChatMessage = {
  role: "developer" | "user" | "assistant";
  content: string;
};

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function errorResponse(error: string, status: number, details?: string) {
  return Response.json(
    details ? { error, details } : { error },
    { status },
  );
}

export async function POST(request: Request) {
  if (!process.env.OPENAI_API_KEY) {
    return errorResponse(
      "Server configuration error.",
      500,
      "OPENAI_API_KEY is not configured.",
    );
  }

  let body: unknown;

  try {
    body = await request.json();
  } catch {
    return errorResponse("Invalid request body.", 400, "The body must be valid JSON.");
  }

  if (!isRecord(body)) {
    return errorResponse("Invalid request body.", 400, "Expected a JSON object.");
  }

  if (!Array.isArray(body.messages)) {
    return errorResponse(
      "Invalid messages.",
      400,
      "The messages field must be an array.",
    );
  }

  if (body.messages.length === 0) {
    return errorResponse(
      "Invalid messages.",
      400,
      "The messages array must contain at least one message.",
    );
  }

  const messages: ChatMessage[] = [];

  for (const [index, message] of body.messages.entries()) {
    if (!isRecord(message)) {
      return errorResponse(
        "Invalid message.",
        400,
        `Message at index ${index} must be an object.`,
      );
    }

    if (message.role !== "user" && message.role !== "assistant") {
      return errorResponse(
        "Invalid message role.",
        400,
        `Message at index ${index} must have a role of user or assistant.`,
      );
    }

    if (typeof message.content !== "string" || message.content.trim() === "") {
      return errorResponse(
        "Invalid message content.",
        400,
        `Message at index ${index} must have non-empty string content.`,
      );
    }

    messages.push({ role: message.role, content: message.content });
  }

  const memoryContents: string[] = [];

  if (body.memories !== undefined) {
    if (!Array.isArray(body.memories)) {
      return errorResponse(
        "Invalid memories.",
        400,
        "The memories field must be an array when provided.",
      );
    }

    if (body.memories.length > MAX_MEMORY_COUNT) {
      return errorResponse(
        "Memory payload too large.",
        413,
        `At most ${MAX_MEMORY_COUNT} memories may be sent.`,
      );
    }

    for (const memory of body.memories) {
      if (!isRecord(memory) || typeof memory.content !== "string") continue;

      const content = memory.content.replace(/\s+/g, " ").trim();
      if (!content) continue;

      if (content.length > MAX_MEMORY_LENGTH) {
        return errorResponse(
          "Memory payload too large.",
          413,
          `Each memory must be at most ${MAX_MEMORY_LENGTH} characters.`,
        );
      }

      memoryContents.push(content);
    }

    const totalMemoryLength = memoryContents.reduce(
      (total, content) => total + content.length,
      0,
    );

    if (totalMemoryLength > MAX_TOTAL_MEMORY_LENGTH) {
      return errorResponse(
        "Memory payload too large.",
        413,
        `Combined memory content must be at most ${MAX_TOTAL_MEMORY_LENGTH} characters.`,
      );
    }
  }

  const memoryContext = memoryContents.length
    ? [
        "Saved user facts and preferences:",
        ...memoryContents.map((content) => `- ${content}`),
        "",
        "Use these details only when relevant. Do not mention that they came from localStorage or force unrelated memories into an answer. Current user instructions always take priority.",
      ].join("\n")
    : null;

  const input: ChatMessage[] = memoryContext
    ? [{ role: "developer", content: memoryContext }, ...messages]
    : messages;

  try {
    const openai = new OpenAI({ apiKey: process.env.OPENAI_API_KEY });
    const responseStream = await openai.responses.create({
      model: MODEL,
      input,
      stream: true,
    });

    const encoder = new TextEncoder();
    const stream = new ReadableStream<Uint8Array>({
      async start(controller) {
        try {
          for await (const event of responseStream) {
            if (event.type === "response.output_text.delta") {
              controller.enqueue(encoder.encode(event.delta));
            }
          }

          controller.close();
        } catch (error: unknown) {
          console.error("OpenAI response stream failed.", error);
          controller.error(error);
        }
      },
    });

    return new Response(stream, {
      headers: {
        "Cache-Control": "no-cache, no-transform",
        "Content-Type": "text/plain; charset=utf-8",
        "X-Content-Type-Options": "nosniff",
      },
    });
  } catch (error: unknown) {
    console.error("OpenAI chat request failed.", error);

    if (error instanceof OpenAI.APIError) {
      const status = error.status >= 400 && error.status <= 599 ? error.status : 502;
      return Response.json(
        {
          error: "OpenAI API request failed.",
          code: error.code ?? "openai_api_error",
          message: error.message,
        },
        { status },
      );
    }

    return errorResponse(
      "OpenAI API request failed.",
      502,
      "Unable to complete the request.",
    );
  }
}
