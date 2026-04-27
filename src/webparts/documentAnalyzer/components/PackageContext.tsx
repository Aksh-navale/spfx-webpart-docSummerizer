import * as React from "react";
import { createContext, useState, ReactNode } from "react";

/*  FILE TYPE */
export interface IPackageFile {
  name: string;
  url: string;
  fileObject?: File; 
  content?: string; 
}

/*  CONTEXT TYPE */
interface IPackageContext {
  packageFiles: IPackageFile[];
  togglePackage: (file: IPackageFile) => void;
  resetAll: () => void;
}

/* CONTEXT */
export const PackageContext = createContext<IPackageContext>({
  packageFiles: [],
  togglePackage: () => {},
  resetAll: () => {}
});

/*  PROVIDER PROPS */
interface IPackageProviderProps {
  children: ReactNode;
}

/*  PROVIDER */
export const PackageProvider: React.FC<IPackageProviderProps> = ({
  children
}) => {
  const [packageFiles, setPackageFiles] = useState<IPackageFile[]>([]);

  /*  ADD / REMOVE FILE */
const togglePackage = (file: IPackageFile) => {
  setPackageFiles((prev) => {
    const exists = prev.find((f) => f.name === file.name);

    if (exists) {
      return prev.filter((f) => f.name !== file.name);
    }

    return [
      ...prev,
      {
        name: file.name,
        url: file.url || "",
        fileObject: file.fileObject,
        content: file.content, // ✅ THIS IS THE FIX
      },
    ];
  });
};

  /*  RESET */
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