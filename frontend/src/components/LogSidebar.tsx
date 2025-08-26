import React, { useEffect, useState, useRef } from "react";

const WS_URL = "ws://localhost:5000/ws/logs"; // Adjust backend websocket URL

const LogSidebar: React.FC = () => {
  const [logs, setLogs] = useState<string[]>([]);
  const logEndRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const ws = new WebSocket(WS_URL);

    ws.onmessage = (event) => {
      setLogs((prev) => [...prev, event.data]);
    };

    ws.onclose = () => {
      console.log("WebSocket closed");
    };

    return () => {
      ws.close();
    };
  }, []);

  useEffect(() => {
    logEndRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [logs]);

  return (
    <div className="bg-black text-green-400 font-mono p-4 h-screen w-80 overflow-y-auto rounded-l-lg shadow-lg fixed top-0 left-0">
      <h2 className="text-green-200 text-lg mb-2 border-b border-green-700 pb-1">Server Logs</h2>
      <div className="flex-1 overflow-y-auto space-y-1">
        {logs.map((line, idx) => (
          <pre key={idx} className="break-words">{line}</pre>
        ))}
        <div ref={logEndRef} />
      </div>
    </div>
  );
};

export default LogSidebar;
