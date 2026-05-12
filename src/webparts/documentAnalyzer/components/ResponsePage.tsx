import * as React from "react";
import { useContext, useState, useEffect, useRef } from "react";
import styles from "./ResponsePage.module.scss";
import { PackageContext } from "./PackageContext";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import CopilotService from "./services/CopilotService";
import * as pdfjsLib from "pdfjs-dist";
import ReactMarkdown from "react-markdown";
import remarkGfm from "remark-gfm";

(pdfjsLib as any).GlobalWorkerOptions.workerSrc =
  "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";

interface IResponsePageProps {
  goBack: () => void;
  context: WebPartContext;
}

const ResponsePage: React.FC<IResponsePageProps> = ({ goBack, context }) => {
  const { packageFiles, togglePackage } = useContext(PackageContext);

  const [prompt, setPrompt] = useState<string>("");
  const [response, setResponse] = useState<string>("");
  const [loading, setLoading] = useState<boolean>(false);
  const [selectedAction, setSelectedAction] = useState<string>("");

  const copilotRef = useRef<CopilotService | null>(null);

  useEffect(() => {
    if (context) {
      copilotRef.current = new CopilotService(context);
    }
  }, [context]);

  const responseRef = useRef<HTMLDivElement | null>(null);

  useEffect(() => {
    if (responseRef.current) {
      responseRef.current.scrollTo({
        top: responseRef.current.scrollHeight,
        behavior: "smooth",
      });
    }
  }, [response, loading]);



  const actionRef = useRef<HTMLSelectElement | null>(null);

  // const readFileContent = async (file: File): Promise<string> => {
  //   const fileType = file.type;

  //   console.log("📄 Reading file:", file.name, fileType);

  //   try {
  //     // ✅ TEXT FILE (SUPER RELIABLE)
  //     if (fileType === "text/plain") {
  //       const text = await file.text();
  //       return text || "Text file is empty";
  //     }

  //     // ✅ PDF FILE (FIXED FOR SPFx)
  //     if (fileType === "application/pdf") {
  //       const arrayBuffer = await file.arrayBuffer();

  //       const pdf = await (pdfjsLib as any).getDocument({
  //         data: new Uint8Array(arrayBuffer),
  //         disableWorker: true,
  //       }).promise;

  //       let fullText = "";

  //       for (let i = 1; i <= pdf.numPages; i++) {
  //         const page = await pdf.getPage(i);
  //         const content = await page.getTextContent();

  //         if (!content.items || content.items.length === 0) {
  //           console.warn(`⚠️ Page ${i} has no text`);
  //           continue;
  //         }

  //         const pageText = content.items
  //           .map((item: any) => item.str || "")
  //           .join(" ");

  //         // ✅ Clean text
  //         const cleaned = pageText.replace(/\s+/g, " ").trim();

  //         fullText += `\n--- Page ${i} ---\n${cleaned}\n`;
  //       }

  //       console.log("✅ Extracted PDF length:", fullText.length);

  //       if (!fullText.trim()) {
  //         return "⚠️ PDF has no readable text (might be scanned)";
  //       }

  //       return fullText;
  //     }

  //     return "Unsupported file type (only TXT & PDF supported)";
  //   } catch (error) {
  //     console.error("❌ FILE READ ERROR:", error);
  //     return "Error reading file";
  //   }
  // };

   const generateResponse = async (): Promise<string | null> => {
  if (!prompt.trim()) {
    alert("Please enter a prompt");
    return null; // ✅ FIX
  }

  if (packageFiles.length === 0) {
    alert("Please select at least one document");
    return null; // ✅ FIX
  }

  if (!copilotRef.current) {
    alert("Copilot not initialized");
    return null; // ✅ FIX
  }

  setLoading(true);
  setResponse("");

  try {
    const fileContents = packageFiles.map((f) => {
      if (!f.content || f.content.length < 50) {
        return `📄 ${f.name}: ⚠️ No usable content`;
      }

      return `📄 ${f.name}:\n${f.content.substring(0, 2000)}`;
    });

    const combinedFilesText = fileContents.join("\n\n");

    const trimmedText =
      combinedFilesText.length > 6000
        ? combinedFilesText.substring(0, 6000) + "\n...[truncated]"
        : combinedFilesText;

    const finalPrompt = `
${prompt}

Selected Documents Content:
${trimmedText}

👉 Provide a clear, structured response.
`;

    const result = await copilotRef.current.sendMessage(finalPrompt);

    setResponse(result);
    return result; // ✅ OK
  } catch (error: any) {
    console.error("❌ ERROR:", error);

    const message =
      error?.message ||
      error?.response?.data ||
      "Unknown error";

    setResponse(`❌ ${message}`);
    return null; // ✅ MUST RETURN
  } finally {
    setLoading(false);
  }
};

  // const handleAction = (action: string, text: string) => {
  //   const encodedText = encodeURIComponent(text);

  //   switch (action) {
  //     case "teamsChannel":
  //       window.open(
  //         `https://teams.microsoft.com/l/channel/YOUR_CHANNEL_ID/General?groupId=YOUR_TEAM_ID&tenantId=YOUR_TENANT_ID`,
  //         "_blank"
  //       );
  //       break;

  //     case "teamsMessage":
  //       window.open(
  //         `https://teams.microsoft.com/l/chat/0/0?users=&message=${encodedText}`,
  //         "_blank"
  //       );
  //       break;

  //     case "email":
  //       window.open(
  //         `mailto:?subject=Document Summary&body=${encodedText}`,
  //         "_blank"
  //       );
  //       break;

  //     case "sharepoint":
  //       alert("SharePoint integration can be added via API");
  //       break;

  //     default:
  //       alert("Please select an action");
  //   }
  // };

  return (
    <div className={styles.wrapper}>
      {/* TOP BAR */}
      <div className={styles.topBar}>
        <input
          type="text"
          placeholder="Summarize these documents to generate a..."
          value={prompt}
          onChange={(e) => setPrompt(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === "Enter") {
              void generateResponse();
            }
          }}
        />
        <button onClick={() => void generateResponse()}>➤</button>
      </div>

      <div className={styles.mainLayout}>
        {/* LEFT SECTION */}
        <div className={styles.leftPanel}>
          <div className={styles.responseContainer} ref={responseRef}>
            {loading ? (
              <p className={styles.loader}>⏳ Opening...</p>
            ) : (
              <div className={styles.markdown}>
                <ReactMarkdown remarkPlugins={[remarkGfm]}>
                  {response || "Your generated response will appear here"}
                </ReactMarkdown>
              </div>
            )}
          </div>
        </div>

        {/* RIGHT SECTION */}
        <div className={styles.rightPanel}>
          {/* BACK LINK */}
          <div className={styles.backLink} onClick={goBack}>
            ← Back to Document Search
          </div>

          {/* SELECTED DOCUMENTS */}
          <div className={styles.card}>
            <h3>Selected Documents :</h3>

            <div className={styles.docList}>
              {packageFiles.length > 0 ? (
                packageFiles.map((file, i) => (
                  <div key={i} className={styles.docItem}>
                    <span className={styles.docName}>📄 {file.name}</span>

                    <span
                      className={styles.deleteIcon}
                      onClick={() => togglePackage(file)}
                    >
                      🗑️
                    </span>
                  </div>
                ))
              ) : (
                <p>No documents selected</p>
              )}
            </div>
          </div>

          {/* ACTIONS */}
          <div className={styles.card}>
            <h3>Actions:</h3>

            <select
              ref={actionRef}
              className={styles.dropdown}
              value={selectedAction}
              onChange={(e) => setSelectedAction(e.target.value)}
            >
              <option value="">Select an action:</option>
              <option value="teamsChannel">Post to Teams Channel</option>
              <option value="teamsMessage">Send as Teams Message</option>
              <option value="email">Send In an OutlookEmail</option>
              <option value="sharepoint">Post as SharePoint News</option>
            </select>
          </div>

          {/* SUBMIT BUTTON */}
          <button
            className={styles.submitBtn}
            onClick={async () => {
              if (!selectedAction) {
                alert("Please select an action first");
                return;
              }

              if (!response) {
                alert("Please generate response first");
                return;
              }

              setLoading(true); // ✅ show loader while opening

              try {
                const message = `Here is the generated summary:\n\n${response}`;
                const encodedText = encodeURIComponent(message);

                switch (selectedAction) {
                  case "teamsMessage":
                    window.open(
                      `https://teams.microsoft.com/l/chat/0/0?users=&message=${encodedText}`,
                      "_blank"
                    );
                    break;

                  case "teamsChannel":
                    window.open(
                      `https://teams.microsoft.com/l/channel/YOUR_CHANNEL_ID/General?groupId=YOUR_TEAM_ID&tenantId=YOUR_TENANT_ID`,
                      "_blank"
                    );
                    break;

                  case "email":
                    window.location.href = `mailto:?subject=Document Summary&body=${encodedText}`;
                    break;

                  case "sharepoint":
                    alert("SharePoint integration can be added via API");
                    break;
                }
              } finally {
                setLoading(false); // ✅ stop loader
              }
            }}
            >
            Submit
          </button>
        </div>
      </div>
    </div>
  );
};

export default ResponsePage;
