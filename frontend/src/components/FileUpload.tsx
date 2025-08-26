// import React, { useRef, useState } from "react";
// import JSZip from "jszip";

// export type PreviewData = { [fileName: string]: string };

// interface Props {
//   onPreview: (data: PreviewData) => void;
//   onDownload: (url: string) => void;
//   companyName: string;
//   projectTitle: string;
//   logoFile: File | null;
// }

// const FileUpload: React.FC<Props> = ({
//   onPreview,
// }) => {
//   const fileInputRef = useRef<HTMLInputElement>(null);
//   const [selectedFile, setSelectedFile] = useState<File | null>(null);
//   const [uploading] = useState(false);
//   const [error, setError] = useState<string | null>(null);
//   const [dragOver, setDragOver] = useState(false);

//   const allowedTypes = [
//     "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
//   ];
//   const maxSizeMB = 5;

//   function validateAndSetFile(file: File) {
//     if (!allowedTypes.includes(file.type)) {
//       setError("Only .xlsx files are supported.");
//       setSelectedFile(null);
//       onPreview({});
//       return;
//     }
//     if (file.size > maxSizeMB * 1024 * 1024) {
//       setError("File size must be less than 5MB.");
//       setSelectedFile(null);
//       onPreview({});
//       return;
//     }

//     setError(null);
//     setSelectedFile(file);

//     // Read Excel as ZIP and extract XML files for preview
//     const reader = new FileReader();
//     reader.onload = async (e) => {
//       const data = e.target?.result;
//       if (!data) return;
//       try {
//         const zip = await JSZip.loadAsync(data);
//         const preview: PreviewData = {};
//         // Extract first 3 XML files text content for preview
//         const xmlFiles = Object.keys(zip.files).filter((name) =>
//           name.endsWith(".xml")
//         );
//         for (let i = 0; i < Math.min(3, xmlFiles.length); i++) {
//           preview[xmlFiles[i]] = await zip.files[xmlFiles[i]].async("text");
//         }
//         onPreview(preview);
//       } catch {
//         setError("Failed to read Excel file as ZIP.");
//         onPreview({});
//       }
//     };
//     reader.readAsArrayBuffer(file);
//   }

//   const handleDrop = (event: React.DragEvent<HTMLDivElement>) => {
//     event.preventDefault();
//     setDragOver(false);
//     const files = event.dataTransfer.files;
//     if (files && files[0]) validateAndSetFile(files[0]);
//   };

//   const handleDragOver = (event: React.DragEvent<HTMLDivElement>) => {
//     event.preventDefault();
//     setDragOver(true);
//   };

//   const handleDragLeave = () => setDragOver(false);

//   const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
//     const files = event.target.files;
//     if (files && files[0]) validateAndSetFile(files[0]);
//   };

//   function handleGenerateReport() {
//     if (!selectedFile) {
//       setError("Please upload a valid Excel file.");
//       return;
//     }
//     // TODO: Add your report generation API call here,
//     // you may set uploading, progress, call backend and then 
//     // call onDownload with the download URL returned by backend
//   }

//   return (
//     <div className="max-w-xl mx-auto p-6 bg-white shadow-lg rounded-2xl">
//       <h2 className="text-xl font-bold text-gray-800 mb-4">Upload Excel File</h2>
//       <div
//         className={`border-2 border-dashed rounded-lg p-8 text-center cursor-pointer transition-colors duration-200 ${
//           dragOver ? "border-blue-500 bg-blue-50" : "border-gray-300 bg-gray-50"
//         }`}
//         onDrop={handleDrop}
//         onDragOver={handleDragOver}
//         onDragLeave={handleDragLeave}
//         onClick={() => fileInputRef.current?.click()}
//       >
//         <input
//           ref={fileInputRef}
//           type="file"
//           accept=".xlsx"
//           onChange={handleFileChange}
//           hidden
//         />
//         <p className="text-gray-600 font-medium">
//           {selectedFile ? (
//             <span className="text-green-700">{selectedFile.name}</span>
//           ) : (
//             "Drag & drop Excel file here, or click to browse"
//           )}
//         </p>
//         <p className="text-sm text-gray-500 mt-1">(.xlsx only, max 5MB)</p>
//       </div>

//       {error && <p className="text-red-600 text-sm mt-3">⚠ {error}</p>}
//       {uploading && (
//         <div className="w-full bg-gray-200 rounded-full h-3 mt-4 overflow-hidden">
//           <div
//             className="bg-blue-500 h-3 text-xs text-white flex items-center justify-center transition-all duration-300"
//             style={{ width: `${onprogress}%` }}
//           >
//           </div>
//         </div>
//       )}

//       <div className="flex justify-end mt-6">
//         <button
//           disabled={!selectedFile || uploading}
//           onClick={handleGenerateReport}
//           className={`px-5 py-2 rounded-lg font-semibold text-white transition-colors duration-200 ${
//             !selectedFile || uploading
//               ? "bg-gray-400 cursor-not-allowed"
//               : "bg-blue-600 hover:bg-blue-700"
//           }`}
//         >
//           {uploading ? "Generating Report..." : "Generate Report"}
//         </button>
//       </div>
//     </div>
//   );
// };

// export default FileUpload;

import React, { useRef, useState } from "react";
import * as XLSX from "xlsx";

export type SheetPreviewData = {
  sheetName: string;
  data: Array<Array<string | number | null>>;
};

interface Props {
  onPreview: (data: SheetPreviewData[]) => void;
  onDownload: (url: string) => void;
  companyName: string;
  projectTitle: string;
  logoFile: File | null;
}

const FileUpload: React.FC<Props> = ({ onPreview }) => {
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [dragOver, setDragOver] = useState(false);

  const allowedTypes = [
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  ];
  const maxSizeMB = 5;

  function validateAndSetFile(file: File) {
    if (!allowedTypes.includes(file.type)) {
      setError("Only .xlsx files are supported.");
      setSelectedFile(null);
      onPreview([]);
      return;
    }
    if (file.size > maxSizeMB * 1024 * 1024) {
      setError("File size must be less than 5MB.");
      setSelectedFile(null);
      onPreview([]);
      return;
    }

    setError(null);
    setSelectedFile(file);

    const reader = new FileReader();
    reader.onload = (e) => {
      const data = e.target?.result;
      if (!data) return;

      try {
        const workbook = XLSX.read(data, { type: "array" });
        const sheetsPreview: SheetPreviewData[] = workbook.SheetNames.map(sheetName => {
          const worksheet = workbook.Sheets[sheetName];
          const sheetData = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as (string | number | null)[][];
          return { sheetName, data: sheetData };
        });
        onPreview(sheetsPreview);
      } catch {
        setError("Failed to parse Excel file.");
        onPreview([]);
      }
    };
    reader.readAsArrayBuffer(file);
  }

  // ... drag & drop handlers remain unchanged ...

  const handleDrop = (event: React.DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    setDragOver(false);
    const files = event.dataTransfer.files;
    if (files && files[0]) validateAndSetFile(files[0]);
  };

  const handleDragOver = (event: React.DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    setDragOver(true);
  };

  const handleDragLeave = () => setDragOver(false);

  const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    const files = event.target.files;
    if (files && files[0]) validateAndSetFile(files[0]);
  };

  return (
    <div className="max-w-xl mx-auto p-6 bg-white shadow-lg rounded-2xl">
      <h2 className="text-xl font-bold text-gray-800 mb-4">Upload Excel File</h2>
      <div
        className={`border-2 border-dashed rounded-lg p-8 text-center cursor-pointer transition-colors duration-200 ${
          dragOver ? "border-blue-500 bg-blue-50" : "border-gray-300 bg-gray-50"
        }`}
        onDrop={handleDrop}
        onDragOver={handleDragOver}
        onDragLeave={handleDragLeave}
        onClick={() => fileInputRef.current?.click()}
      >
        <input
          ref={fileInputRef}
          type="file"
          accept=".xlsx"
          onChange={handleFileChange}
          hidden
        />
        <p className="text-gray-600 font-medium">
          {selectedFile ? (
            <span className="text-green-700">{selectedFile.name}</span>
          ) : (
            "Drag & drop Excel file here, or click to browse"
          )}
        </p>
        <p className="text-sm text-gray-500 mt-1">(.xlsx only, max 5MB)</p>
      </div>
      {error && <p className="text-red-600 text-sm mt-3">⚠ {error}</p>}
    </div>
  );
};

export default FileUpload;
