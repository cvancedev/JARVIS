import { Message as MessageType } from "@/types/message";

interface MessageProps {
  message: MessageType;
  canSpeak?: boolean;
  isSpeaking?: boolean;
  onSpeak?: () => void;
}

export default function Message({
  message,
  canSpeak = false,
  isSpeaking = false,
  onSpeak,
}: MessageProps) {
  const isUser = message.role === "user";

  return (
    <div
      className={`mb-4 flex ${
        isUser ? "justify-end" : "justify-start"
      }`}
    >
      <div
        className={`max-w-xl rounded-lg px-4 py-3 ${
          isUser
            ? "bg-blue-600 text-white"
            : "bg-zinc-800 text-white"
        }`}
      >
        <p>{message.content}</p>
        {canSpeak ? (
          <button
            type="button"
            className={`mt-2 text-xs transition-colors ${
              isSpeaking
                ? "text-blue-300 hover:text-blue-200"
                : "text-zinc-500 hover:text-zinc-300"
            }`}
            onClick={onSpeak}
            aria-pressed={isSpeaking}
          >
            {isSpeaking ? "Stop" : "Speak"}
          </button>
        ) : null}
      </div>
    </div>
  );
}
