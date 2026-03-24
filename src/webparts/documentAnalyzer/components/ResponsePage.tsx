import * as React from "react";
import { useContext, useState, useEffect, useRef } from "react";
import styles from "./ResponsePage.module.scss";
import { PackageContext } from "./PackageContext";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import CopilotService from "./services/CopilotService";

interface IResponsePageProps {
  goBack: () => void;
  context: WebPartContext;
}

const ResponsePage: React.FC<IResponsePageProps> = ({ goBack, context }) => {
  const { packageFiles, togglePackage } = useContext(PackageContext);

  const [prompt, setPrompt] = useState<string>("");
  const [response, setResponse] = useState<string>("");
  const [loading, setLoading] = useState<boolean>(false);

  const copilotRef = useRef<CopilotService | null>(null);

  useEffect(() => {
    if (context) {
      copilotRef.current = new CopilotService(context);
    }
  }, [context]);

  const generateResponse = async () => {
    if (!prompt.trim()) {
      alert("Please enter a prompt");
      return;
    }

    if (packageFiles.length === 0) {
      alert("Please select at least one document");
      return;
    }

    if (!copilotRef.current) {
      alert("Copilot not initialized");
      return;
    }

    setLoading(true);
    setResponse("");

    try {
      // 🔥 Combine selected documents into prompt
      const docsText = packageFiles
        .map((f) => `${f.name} → ${f.url}`)
        .join("\n");

        const fileUrls = packageFiles.map(f => f.url).join("\n");

        const finalPrompt = `
            ${prompt}

            Selected Documents:
            ${fileUrls}

            👉 Provide clear structured response.
            👉 Use bullet points where helpful.
           `;
      console.log("docsText", docsText)
      // ✅ CALL YOUR EXISTING SERVICE
        const result = await copilotRef.current.sendMessage(finalPrompt);

      setResponse(result);
    } catch (error) {
      console.error("Copilot Error:", error);
      setResponse("Error generating response");
    }

    setLoading(false);
  };

  return (
    <div className={styles.container}>
      
      {/* LEFT SIDE */}
      <div className={styles.leftSection}>
        
        {/* BACK */}
        <div className={styles.back} onClick={goBack}>
          ← Back
        </div>

        {/* PROMPT */}
        <div className={styles.promptBox}>
          <input
            type="text"
            placeholder="Ask something about selected documents..."
            value={prompt}
            onChange={(e) => setPrompt(e.target.value)}
            onKeyDown={(e) => {
                if (e.key === "Enter") {
                void generateResponse();
                }
            }}
        />

          <button onClick={() => void generateResponse()}>
            Generate
          </button>
        </div>

        {/* RESPONSE */}
        <div className={styles.responseBox}>
          {loading ? (
            <p>🤖 Copilot is generating response...</p>
          ) : (
            <textarea value={response} readOnly />
          )}
        </div>
      </div>

      {/* RIGHT SIDE */}
      <div className={styles.rightSection}>
        <h3>Selected Documents</h3>

        {packageFiles.length > 0 ? (
          packageFiles.map((file: string, i: number) => (
            <div key={i} className={styles.docItem}>
              <span>📄 {file}</span>
              <span onClick={() => togglePackage(file)}>🗑️</span>
            </div>
          ))
        ) : (
          <p>No documents selected</p>
        )}
      </div>
    </div>
  );
};

export default ResponsePage;