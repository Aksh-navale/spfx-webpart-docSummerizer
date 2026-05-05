import * as React from "react";
import styles from "./CopilotSummary.module.scss";
import { ICopilotSummaryProps } from "./ICopilotSummaryProps";
import CopilotService from "../services/CopilotService";

interface IMessage {
  role: "user" | "assistant";
  text: string;
}

export default function CopilotSummary(props: ICopilotSummaryProps) {
  const [messages, setMessages] = React.useState<IMessage[]>([]);
  const [input, setInput] = React.useState("");

  const [summaryLoading, setSummaryLoading] = React.useState(false);
  const [chatLoading, setChatLoading] = React.useState(false);

  const [file, setFile] = React.useState<File | null>(null);
  const [instruction, setInstruction] = React.useState("");

  const serviceRef = React.useRef(new CopilotService(props.context));

  // ✅ SCROLL REF (END OF CHAT)
  const chatEndRef = React.useRef<HTMLDivElement | null>(null);

  // ✅ AUTO SCROLL TO LATEST MESSAGE
  React.useEffect(() => {
    chatEndRef.current?.scrollIntoView({
      behavior: "smooth"
    });
  }, [messages, chatLoading, summaryLoading]);

  // CHAT MESSAGE
  const send = async () => {
    if (!input) return;

    const userMsg: IMessage = { role: "user", text: input };
    setMessages((m) => [...m, userMsg]);
    setInput("");

    setChatLoading(true);

    const reply = await serviceRef.current.sendMessage(input);

    setMessages((m) => [...m, { role: "assistant", text: reply }]);

    setChatLoading(false);
  };

  // READ FILE
  const readFileText = (file: File): Promise<string> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();

      reader.onload = () => resolve(reader.result as string);
      reader.onerror = () => reject("Error reading file");

      reader.readAsText(file);
    });
  };

  // DOCUMENT ANALYSIS
  const handleAnalyze = async () => {
    if (!file || !instruction) return;

    setSummaryLoading(true);

    try {
      const fileText = (await readFileText(file)).slice(0, 12000);

      const prompt = `
${instruction}

Document Content:
${fileText}
`;

      setMessages((m) => [
        ...m,
        { role: "user", text: `📄 ${file.name}\n${instruction}` }
      ]);

      const reply = await serviceRef.current.sendMessage(prompt);

      setMessages((m) => [...m, { role: "assistant", text: reply }]);
    } catch {
      setMessages((m) => [
        ...m,
        { role: "assistant", text: "Error analyzing document." }
      ]);
    }

    setSummaryLoading(false);
  };

  const formatCopilotText = (text: string) => {
    return text
      .replace(/\*\*(.*?)\*\*/g, "<strong>$1</strong>")
      .replace(/\n/g, "<br/>")
      .replace(/•/g, "&#8226; ");
  };

  return (
    <div className={styles.wrapper}>
      {/* HEADER */}
      <div className={styles.header}>
        <div className={styles.title}>🤖 Copilot Assistant</div>
        <div className={styles.subtitle}>Chat • Analyze • Summarize</div>
      </div>

      {/* FILE ANALYZER */}
      <div className={styles.card}>
        <div className={styles.cardTitle}>📄 Document Analyzer</div>

        <input
          type="file"
          className={styles.fileInput}
          onChange={(e) => {
            const f = e.target.files?.[0];
            if (f) setFile(f);
          }}
        />

        {file && (
          <textarea
            className={styles.textarea}
            placeholder="Enter instruction"
            value={instruction}
            onChange={(e) => setInstruction(e.target.value)}
          />
        )}

        {file && instruction && (
          <button className={styles.primaryBtn} onClick={handleAnalyze}>
            Analyze File
          </button>
        )}

        {summaryLoading && (
          <div className={styles.loading}>Analyzing document...</div>
        )}
      </div>

      {/* CHAT AREA */}
      <div className={styles.chatBox}>
        {messages.map((m, i) => (
          <div key={i} className={`${styles.message} ${styles[m.role]}`}>
            <strong>{m.role === "user" ? "You" : "Copilot"}:</strong>
            <div
              dangerouslySetInnerHTML={{
                __html: formatCopilotText(m.text)
              }}
            />
          </div>
        ))}

        {(chatLoading || summaryLoading) && (
          <div className={`${styles.message} ${styles.assistant}`}>
            <strong>Copilot:</strong>
            <div className={styles.typingIndicator}>
              <span></span>
              <span></span>
              <span></span>
            </div>
          </div>
        )}

        {/* ✅ IMPORTANT: SCROLL TARGET */}
        <div ref={chatEndRef} />
      </div>

      {/* INPUT */}
      <div className={styles.inputRow}>
        <input
          className={styles.input}
          value={input}
          onChange={(e) => setInput(e.target.value)}
          placeholder="Ask Copilot..."
        />

        <button className={styles.sendBtn} onClick={send}>
          Send
        </button>
      </div>
    </div>
  );
}


// List the days of the week
// Explain what Microsoft 365 Copilot is
// Write 3 bullet points about productivity
// Summarize teamwork benefits
// Explain SharePoint in 3 bullets
//  summerize in 5 key points
//  List days of week
