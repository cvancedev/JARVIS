import Message from "./Message";
import { Message as MessageType } from "@/types/message";

interface ChatWindowProps {
  messages: MessageType[];
}

export default function ChatWindow({
  messages,
}: ChatWindowProps) {
  return (
    <div className="space-y-4">
      {messages.map((message) => (
        <Message
          key={message.id}
          message={message}
        />
      ))}
    </div>
  );
}