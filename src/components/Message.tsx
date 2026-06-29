import { Message as MessageType } from "@/types/message";

interface MessageProps {
  message: MessageType;
}

export default function Message({ message }: MessageProps) {
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
        {message.content}
      </div>
    </div>
  );
}