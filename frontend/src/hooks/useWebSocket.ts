import { useEffect, useState, useRef } from 'react';

export function useWebSocket(url: string) {
  const [messages, setMessages] = useState<string[]>([]);
  const ws = useRef<WebSocket | null>(null);

  useEffect(() => {
    ws.current = new WebSocket(url);

    ws.current.onmessage = (event) => {
      setMessages((prev) => [...prev, event.data]);
    };

    ws.current.onclose = () => {
      console.log('WebSocket closed');
      // Optional: reconnect logic here
    };

    return () => {
      ws.current?.close();
    };
  }, [url]);

  return { messages };
}
