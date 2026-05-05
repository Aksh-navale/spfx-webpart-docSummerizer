import * as React from "react";
import { useState, useEffect, useContext, useRef } from "react";
import styles from "./DocumentPage.module.scss";
import { PackageContext } from "./PackageContext";
import { WebPartContext } from "@microsoft/sp-webpart-base";
// import { SPHttpClient } from "@microsoft/sp-http";
import * as pdfjsLib from "pdfjs-dist";
import pdfWorker from "pdfjs-dist/build/pdf.worker.entry";
import { GlobalWorkerOptions } from "pdfjs-dist";
import Tesseract from "tesseract.js";
import * as XLSX from "xlsx";
import mammoth from "mammoth";
import ReactJson from "react-json-view";

(pdfjsLib as any).GlobalWorkerOptions.workerSrc = pdfWorker;

GlobalWorkerOptions.workerSrc = require("pdfjs-dist/build/pdf.worker.js");

import {
  FilePicker,
  IFilePickerResult,
} from "@pnp/spfx-controls-react/lib/FilePicker";

/* PROPS */
export interface IDocumentPageProps {
  goToResponse: () => void;
  context: WebPartContext;
}

/* FILE TYPE */
interface IFile {
  id: number;
  name: string;
  url?: string;
  fileObject?: File;
  content?: string;
  source?: "local" | "sharepoint";
  isProcessing?: boolean;
  previewUrl?: string;
}

const DocumentPage: React.FC<IDocumentPageProps> = ({
  goToResponse,
  context,
}) => {
  const [searchValue, setSearchValue] = useState("");
  const [allFiles, setAllFiles] = useState<IFile[]>([]);
  const [results, setResults] = useState<IFile[]>([]);

  // const [graphFiles, setGraphFiles] = useState<IFile[]>([]);


  const { packageFiles, togglePackage, resetAll } = useContext(PackageContext);
  const fileInputRef = useRef<HTMLInputElement | null>(null);

  /* ---------------- PDF TEXT EXTRACTION ---------------- */
    const extractPdfText = async (file: File): Promise<string> => {
  try {
    const buffer = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: buffer }).promise;

    let text = ""; // ✅ MUST be here (top-level inside function)

    // STEP 1: normal extraction
    for (let i = 1; i <= pdf.numPages; i++) {
      const page = await pdf.getPage(i);
      const content = await page.getTextContent();

      const pageText = content.items
        .map((i: any) => i.str || "")
        .join(" ");

      text += pageText + "\n";
    }

    // STEP 2: OCR fallback
    if (!text.trim() || text.length < 30) {
      text = ""; // ✅ reuse same variable (DO NOT redeclare)

      for (let i = 1; i <= pdf.numPages; i++) {
        const page = await pdf.getPage(i);
        const viewport = page.getViewport({ scale: 4 });

        const canvas = document.createElement("canvas");
        const context = canvas.getContext("2d");

        canvas.height = viewport.height;
        canvas.width = viewport.width;

        await page.render({
          canvasContext: context!,
          viewport,
        }).promise;

        const result = await Tesseract.recognize(canvas, "eng", {
          logger: m => console.log(m),
        });
        console.log("OCR RESULT FOR PAGE " + i, result);
        const ocrText = result.data.text || "";

        console.log("OCR TEXT LENGTH:", ocrText.length);
        console.log("OCR TEXT SAMPLE:", ocrText.substring(0, 100));

        text += ocrText + "\n";   // ✅ THIS LINE IS MISSING

      }
    }

    // ✅ FINAL RETURN (inside function)
    if (!text.trim()) {
      return "";
    }

    return text.trim();
  } catch (err) {
    console.error("❌ PDF extraction failed:", err);
    return "❌ Error reading PDF";
  }
};

  /* ---------------- PROCESS FILE ---------------- */
   const processFile = async (
  fileObject: File,
  name: string,
  url?: string
) => {
  let content = "";
  const fileName = name.toLowerCase();

  try {
    // ✅ PDF
    if (fileName.endsWith(".pdf")) {
      content = await extractPdfText(fileObject);
    }

    // ✅ TXT
    else if (fileName.endsWith(".txt")) {
      content = await fileObject.text();
    }

    // ✅ CSV
    else if (fileName.endsWith(".csv")) {
      content = await fileObject.text();
    }

    // ✅ JSON (NEW)
    else if (fileName.endsWith(".json")) {
      const text = await fileObject.text();
      try {
        const parsed = JSON.parse(text);
        content = JSON.stringify(parsed); // store clean JSON
      } catch {
        content = text; // fallback if invalid JSON
      }
    }

    // ✅ DOCX (Word)
    else if (fileName.endsWith(".docx")) {
      const arrayBuffer = await fileObject.arrayBuffer();

      const result = await mammoth.extractRawText({
        arrayBuffer,
      });

      content = result.value || "";
    }

    // ✅ EXCEL (xlsx / xls)
    else if (fileName.endsWith(".xlsx") || fileName.endsWith(".xls")) {
      const data = await fileObject.arrayBuffer();

      const workbook = XLSX.read(data, { type: "array" });

      let text = "";

      workbook.SheetNames.forEach((sheetName) => {
        const sheet = workbook.Sheets[sheetName];
        const sheetText = XLSX.utils.sheet_to_csv(sheet);
        text += sheetText + "\n";
      });

      content = text;
    }

    // ✅ IMAGE OCR (NEW 🔥)
    else if (fileName.match(/\.(png|jpg|jpeg|webp)$/)) {
      try {
        const result = await Tesseract.recognize(fileObject, "eng", {
          logger: (m) => console.log("OCR:", m),
        });

        content = result.data.text || "";
      } catch (err) {
        console.error("❌ OCR failed:", err);
        content = "Error reading image";
      }
    }

    // ❌ Unsupported
    else {
      content = "⚠️ Unsupported file type";
    }
  } catch (err) {
    console.error("❌ File processing error:", err);
    content = "❌ Error reading file";
  }

  return {
    id: Date.now(),
    name,
    url: url || "",
    fileObject,
    content,

    // ✅ Needed for image preview
    previewUrl: fileObject
      ? URL.createObjectURL(fileObject)
      : undefined,
  };
};

  /* ---------------- LOCAL UPLOAD ---------------- */
  const handleFileUpload = async (
    event: React.ChangeEvent<HTMLInputElement>,
  ) => {
    const files = Array.from(event.target.files || []);
    if (!files.length) return;

    files.forEach((file) => {
      // ✅ Show immediately
      const tempFile: IFile = {
        id: Date.now() + Math.random(), // ensure unique id
        name: file.name,
        source: "local",
        isProcessing: true,
      };

      setAllFiles((prev) => [...prev, tempFile]);
      setResults((prev) => [...prev, tempFile]);

      // ✅ Process in background
      void (async () => {
        const processed = await processFile(file, file.name);

        const updated: IFile = {
          ...processed,
          source: "local",
          isProcessing: false,
        };

        setAllFiles((prev) =>
          prev.map((f) => (f.id === tempFile.id ? updated : f)),
        );

        setResults((prev) =>
          prev.map((f) => (f.id === tempFile.id ? updated : f)),
        );
      })();
    });

    // ✅ Reset input (important for re-uploading same files)
    event.target.value = "";
  };
  /* ---------------- SHAREPOINT FETCH ---------------- */
  //  const handleSharePointFileSelect = async (file: IFile) => {
  //   try {
  //     // ✅ Convert absolute → server relative
  //     const serverRelativeUrl = file.url?.replace(
  //       window.location.origin,
  //       ""
  //     );

  //     const fileUrl = `${context.pageContext.web.absoluteUrl}/_api/web/getfilebyserverrelativeurl('${serverRelativeUrl}')/$value`;

  //     const res = await context.spHttpClient.get(
  //       fileUrl,
  //       SPHttpClient.configurations.v1
  //     );

  //     const blob = await res.blob();

  //     const fileObject = new File([blob], file.name, {
  //       type: blob.type || "application/octet-stream",
  //     });

  //     const processed = await processFile(fileObject, file.name, file.url);

  //     togglePackage(processed);
  //   } catch (err) {
  //     console.error("❌ SharePoint fetch failed", err);
  //   }
  // };

  /* ---------------- FILE PICKER ---------------- */
  const handleFilePicker = async (pickerResults: IFilePickerResult[]) => {
    if (!pickerResults.length) return;

    pickerResults.forEach((file) => {
      // ✅ show immediately
      const tempFile: IFile = {
        id: Date.now() + Math.random(),
        name: file.fileName,
        url: file.fileAbsoluteUrl,
        source: "sharepoint",
      };

      setAllFiles((prev) => [...prev, tempFile]);
      setResults((prev) => [...prev, tempFile]);

      // ✅ process in background
      void (async () => {
        try {
          const blob = await file.downloadFileContent();

          const fileObject = new File([blob], file.fileName, {
            type: blob.type,
          });

          const processed = await processFile(
            fileObject,
            file.fileName,
            file.fileAbsoluteUrl,
          );

          const updated: IFile = {
            ...processed,
            source: "sharepoint",
          };

          setAllFiles((prev) =>
            prev.map((f) => (f.id === tempFile.id ? updated : f)),
          );

          setResults((prev) =>
            prev.map((f) => (f.id === tempFile.id ? updated : f)),
          );
        } catch (err) {
          console.error("FilePicker error", err);
        }
      })();
    });
  };

  /* ---------------- fetchSharePointFiles organization level ---------------- */



  const fetchGraphFiles = async (query: string): Promise<IFile[]> => {
  try {
    const client = await context.msGraphClientFactory.getClient("3");

    const response = await client.api("/search/query").post({
      requests: [
        {
          entityTypes: ["driveItem"],
          query: {
            queryString: query || "*"
          },
          from: 0,
          size: 20
        }
      ]
    });

    const hits =
      response.value[0]?.hitsContainers[0]?.hits || [];

    return hits.map((hit: any) => {
      const item = hit.resource;

      return {
        id: Date.now() + Math.random(),
        name: item.name,
        url: item.webUrl,
        source: "sharepoint",
        content: "",
      };
    });
  } catch (err) {
    console.error("Graph search failed", err);
    return [];
  }
};

useEffect(() => {
  const runSearch = async () => {
    const localFiltered = allFiles.filter((f) =>
      f.name.toLowerCase().includes(searchValue.toLowerCase())
    );

    let external: IFile[] = [];

    if (searchValue) {
      external = await fetchGraphFiles(searchValue);
    }

    // setGraphFiles(external);
    setResults([...localFiltered, ...external]);
  };

  void  runSearch();
}, [searchValue, allFiles]);

  /* ---------------- UI ---------------- */
  return (
    <div className={styles.docContainer}>
      {/* HEADER */}
      <div className={styles.docHeader}>
        <div className={styles.headerContent}>
          <div className={styles.icon}>📄</div>
          <div>
            <h2>Document Processing</h2>
            <p>Search and manage your document packages with ease</p>
          </div>
        </div>
      </div>
      <div className={styles.docBody}>
        {/* LEFT PANEL */}
        <div className={styles.leftPanel}>
          <div className={styles.card}>
            <h3>🔍 Search Documents</h3>
            <p>
              Find and process your documents quickly with our advanced search
            </p>

            <input
              type="text"
              placeholder="Search documents by name, type, or content..."
              className={styles.searchInput}
              value={searchValue}
              onChange={(e) => setSearchValue(e.target.value)}
            />

            <div className={styles.btnGroup}>
              <div className={styles.leftBtns}>
              <button
                className={`${styles.btn} ${styles.reset}`}
                onClick={resetAll}
              >
                Reset
              </button>

              <button
                className={`${styles.btn} ${styles.upload}`}
                onClick={() => fileInputRef.current?.click()}
              >
                Upload
              </button>
              <div className={styles.sharepoint}>
                <FilePicker
                  context={context}
                  onSave={handleFilePicker}
                  buttonLabel="Browse SharePoint"
                />
              </div>
              </div>
              <button
                className={`${styles.btn} ${styles.process}`}
                onClick={goToResponse}
              >
                Proceed to Document Processing
              </button>
            </div>

            <input
              type="file"
              multiple
              ref={fileInputRef}
              style={{ display: "none" }}
              onChange={handleFileUpload}
            />

            {/* ✅ RESULTS LIST */}
            <div>
              {results.map((file) => {
                const isSelected = packageFiles.some(
                  (f) => f.name === file.name,
                );

                return (
                  <div
                    key={file.id}
                    style={{
                      padding: "10px",
                      borderBottom: "1px solid #eee",
                      display: "flex",
                      justifyContent: "space-between",
                      alignItems: "center",
                    }}
                  >
                    {/* FILE INFO */}
                    <div>
                      <div>
                        {file.name.endsWith(".xlsx") ? "📊" :
                        file.name.endsWith(".docx") ? "📝" :
                        "📄"} <a
                                  href={file.url}
                                  target="_blank"
                                  rel="noopener noreferrer"
                                  style={{ textDecoration: "none", color: "inherit" }}
                                >
                                  {file.name}
                                </a>
                      </div>

                       {/* ✅ JSON VIEWER HERE */}
                        {file.name.endsWith(".json") && file.content && (
                          <div style={{ marginTop: "10px", maxWidth: "500px" }}>
                            {(() => {
                              try {
                                return (
                                  <ReactJson
                                    src={JSON.parse(file.content)}
                                    collapsed={1}
                                    enableClipboard={true}
                                    displayDataTypes={false}
                                  />
                                );
                              } catch {
                                return <pre>{file.content}</pre>;
                              }
                            })()}
                          </div>
                        )}


                        {file.name.match(/\.(png|jpg|jpeg|webp)$/) && file.previewUrl && (
                          <div style={{ marginTop: "10px" }}>
                            <img
                              src={file.previewUrl}
                              alt="preview"
                              style={{
                                width: "120px",
                                borderRadius: "6px",
                                border: "1px solid #ddd",
                              }}
                            />

                            <p style={{ fontSize: "12px", marginTop: "5px" }}>
                              {file.content || "Processing OCR..."}
                            </p>
                          </div>
                        )}

                      {/* ✅ SOURCE LABEL */}
                      <div
                        style={{
                          fontSize: "12px",
                          color: "#6b7280",
                          marginTop: "2px",
                        }}
                      >
                        {file.source === "local"
                          ? "📁 Local Upload"
                          : "🌐 SharePoint / OneDrive"}
                      </div>
                    </div>

                    {/* ACTION BUTTONS */}
                    <div style={{ display: "flex", gap: "8px" }}>
                      <button
                        style={{
                          background: "rgb(10, 129, 60)",
                          color: "white",
                          border: "none",
                          padding: "5px 10px",
                          borderRadius: "3px",
                          cursor: "pointer",
                        }}
                        disabled={isSelected || file.isProcessing}
                        onClick={() => {
                         if (file.isProcessing){
                            alert("File is still processing. Please wait.");
                            return;
                          }

                          togglePackage({
                            ...file,
                            url: file.url || "",
                          });
                        }}
                      >
                        {isSelected ? "Added" : "Add"}
                      </button>

                      <button
                        style={{
                          background: "#05059c",
                          color: "white",
                          border: "none",
                          padding: "5px 10px",
                          borderRadius: "3px",
                          cursor: "pointer",
                        }}
                        onClick={() => {
                          setAllFiles((prev) =>
                            prev.filter((f) => f.id !== file.id),
                          );
                          setResults((prev) =>
                            prev.filter((f) => f.id !== file.id),
                          );
                        }}
                      >
                        Remove
                      </button>
                    </div>
                  </div>
                );
              })}
            </div>

            <div className={styles.cardRow}>
              <div className={`${styles.miniCard} ${styles.blue}`}>
                <h4>OneDrive Storage</h4>
                <p>
                  Access documents from SharePoint, OneDrive, and other cloud
                  platforms
                </p>

                <FilePicker
                  context={context}
                  onSave={handleFilePicker}
                  buttonLabel="Connect Now →"
                />
              </div>

              <div
                className={`${styles.miniCard} ${styles.green}`}
                onClick={() => fileInputRef.current?.click()}
              >
                <h4>Bulk Upload</h4>
                <p>Upload multiple documents at once for batch processing</p>
                <span className={styles.bulkupload}>Start Upload →</span>
              </div>
            </div>
          </div>
        </div>

        {/* RIGHT PANEL */}

        <div className={styles.rightPanel}>
          <div className={styles.card}>
            <div className={styles.selectedHeader}>
              <h3>Selected Documents</h3>
              <span className={styles.count}>{packageFiles.length}</span>
            </div>

            <div className={styles.selectedList}>
              {packageFiles.length === 0 ? (
                <div className={styles.emptyState}>
                  <div className={styles.icon}>📄</div>
                  <p>No documents selected yet</p>
                  <small>
                    Search and select documents to add them to your processing
                    queue.
                  </small>
                </div>
              ) : (
                packageFiles.map((file, index) => (
                  <div key={index} className={styles.fileRow}>
                    <span className={styles.fileName}>📄 {file.name}</span>

                    <button
                      className={styles.removeBtn}
                      onClick={() => togglePackage(file)}
                    >
                      Remove
                    </button>
                  </div>
                ))
              )}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

export default DocumentPage;
