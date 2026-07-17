import OpenAI from "openai";

export const runtime = "nodejs";

const MODEL = "gpt-5.6";

type ChatMessage = {
  role: "user" | "assistant";
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

  try {
    const openai = new OpenAI({ apiKey: process.env.OPENAI_API_KEY });
    const responseStream = await openai.responses.create({
      model: MODEL,
      input: messages,
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
