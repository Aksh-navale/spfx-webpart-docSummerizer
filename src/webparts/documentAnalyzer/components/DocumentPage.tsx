import * as React from "react";
import { useState, useEffect, useContext } from "react";
import styles from "./DocumentPage.module.scss";
import { PackageContext } from "./PackageContext";
import { WebPartContext } from "@microsoft/sp-webpart-base";


/* ✅ DEFINE TYPES AT TOP (VERY IMPORTANT) */
export interface IDocumentPageProps {
  goToResponse: () => void;
  context: WebPartContext;
}

interface IFile {
  id: number;
  name: string;
  url?: string;
}

/* ✅ USE TYPE HERE */
const DocumentPage: React.FC<IDocumentPageProps> = ({ goToResponse,context  }) => {
  const [searchValue, setSearchValue] = useState<string>("");
  const [files, setFiles] = useState<IFile[]>([]);
  const [results, setResults] = useState<IFile[]>([]);

  const { packageFiles, togglePackage, resetAll } =
    useContext(PackageContext);

  useEffect(() => {
    fetch("http://localhost:3001/files")
      .then((res) => res.json())
      .then((data: IFile[]) => setFiles(data))
      .catch(() => setFiles([]));
  }, []);

  const search = () => {
    const filtered = files.filter((f) =>
      f.name.toLowerCase().includes(searchValue.toLowerCase())
    );
    setResults(filtered);
  };

    const fileInputRef = React.useRef<HTMLInputElement | null>(null);

    const handleFileUpload = async (
      event: React.ChangeEvent<HTMLInputElement>
    ) => {
      if (event.target.files && event.target.files.length > 0) {
        const file = event.target.files[0];

        try {
          // 🔥 Upload to SharePoint
          const uploadUrl = `${context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('Shared Documents')/Files/add(url='${file.name}',overwrite=true)`;

          const response = await fetch(uploadUrl, {
            method: "POST",
            headers: {
              "Accept": "application/json;odata=verbose"
            },
            body: file
          });

          const data = await response.json();

          const fileUrl =
            context.pageContext.web.absoluteUrl +
            data.d.ServerRelativeUrl;

          console.log("Uploaded File URL:", fileUrl);

          // ✅ Store file with URL
          togglePackage({
            name: file.name,
            url: fileUrl
          });

        } catch (error) {
          console.error("Upload error:", error);
          alert("File upload failed");
        }
      }
    };


  return (
  <div className={styles.container}>
    <div className={styles.search_selected}>
      
      {/* LEFT SIDE */}
      <div className={styles.search_upload}>
        <h2>Document Search</h2>

        <div className={styles.uploadSection}>
            <input
                type="file"
                ref={fileInputRef}
                style={{ display: "none" }}
                onChange={handleFileUpload}
            />

            <button
                className={styles.uploadBtn}
                onClick={() => fileInputRef.current?.click()}
            >
                ⬆️ Upload File
            </button>
        </div>


        <div className={styles.searchBox}>
          <input
            value={searchValue}
            onChange={(e) => setSearchValue(e.target.value)}
            placeholder="Search documents..."
          />
          <button onClick={search}>Search</button>
          <button onClick={resetAll}>Reset</button>
        </div>

        {/* RESULTS */}
        <div className={styles.results}>
          {results.map((file) => (
            <div key={file.id} className={styles.resultCard}>
              <span>{file.name}</span>

              <button
                onClick={() => togglePackage(file.name)}
                className={styles.btn}
              >
                {packageFiles.includes(file.name)
                  ? "Remove"
                  : "Add"}
              </button>
            </div>
          ))}
        </div>

        <button
          className={styles.processBtn}
          onClick={goToResponse}
        >
          Proceed →
        </button>
      </div>

      {/* RIGHT SIDE */}
      <div className={styles.selected_docs}>
        <div className={styles.rightSection}>
          <h3>Selected Documents</h3>

          <div className={styles.documentsList}>
            {packageFiles.length > 0 ? (
              packageFiles.map((f: string, i: number) => (
                <div key={i} className={styles.docItem}>
                  📄 {f}
                </div>
              ))
            ) : (
              <p>No documents added</p>
            )}
          </div>
        </div>
      </div>

    </div>
  </div>
);
};

export default DocumentPage;