import * as React from "react";
import { useState } from "react";
import { IDocumentAnalyzerProps } from "./IDocumentAnalyzerProps";

import DocumentPage from "./DocumentPage";
import ResponsePage from "./ResponsePage";
import { PackageProvider } from "./PackageContext";

const DocumentAnalyzer: React.FC<IDocumentAnalyzerProps> = ({ context }) => {
  const [page, setPage] = useState<"document" | "response">("document");

  return (
    <PackageProvider>
      {page === "document" && (
        <DocumentPage goToResponse={() => setPage("response")} context={context}/>
      )}

      {page === "response" && (
        <ResponsePage
          goBack={() => setPage("document")}
          context={context}
        />
      )}
    </PackageProvider>
  );
};

export default DocumentAnalyzer;