import Message from "./Message";
import { Message as MessageType } from "@/types/message";

const thinkingMessage: MessageType = {
  id: "thinking",
  role: "assistant",
  content: "Thinking...",
};

interface ChatWindowProps {
  messages: MessageType[];
  isLoading: boolean;
}

export default function ChatWindow({
  messages,
  isLoading,
}: ChatWindowProps) {
  return (
    <div className="space-y-4">
      {messages.map((message) => (
        <Message
          key={message.id}
          message={message}
        />
      ))}
      {isLoading ? <Message message={thinkingMessage} /> : null}
    </div>
  );
}