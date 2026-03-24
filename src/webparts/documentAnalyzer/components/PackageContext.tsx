import React, { createContext, useState } from "react";

export const PackageContext = createContext<any>(null);

export const PackageProvider = ({ children }) => {
  const [packageFiles, setPackageFiles] = useState<string[]>([]);

  const togglePackage = (fileName: string) => {
    setPackageFiles((prev) =>
      prev.includes(fileName)
        ? prev.filter((f) => f !== fileName)
        : [...prev, fileName]
    );
  };

  const resetAll = () => {
    setPackageFiles([]);
  };

  return (
    <PackageContext.Provider
      value={{
        packageFiles,
        togglePackage,
        resetAll
      }}
    >
      {children}
    </PackageContext.Provider>
  );
};