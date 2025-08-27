import React, { useRef, useState } from "react";
import * as XLSX from "xlsx";
import axios from "axios";

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

const FileUpload: React.FC<Props> = ({ onPreview, onDownload }) => {
  const fileInputRef = useRef<HTMLInputElement>(null);
  const pptInputRef = useRef<HTMLInputElement>(null);

  const [selectedFile, setSelectedFile] = useState<File | null>(null); //Excel
  const [pptFile, setPptFile] = useState<File | null>(null); // PPT template

  const [error, setError] = useState<string | null>(null);
  const [dragOver, setDragOver] = useState(false);
  const [uploading, setUploading] = useState(false);
  const [uploadProgress, setUploadProgress] = useState(0);

  const allowedExcelTypes = [
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  ];
  const allowedPptTypes = [
    "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  ];
  const maxSizeMB = 5;

  // ================= Excel Upload =================
  function validateAndSetFile(file: File) {
    if (!allowedExcelTypes.includes(file.type)) {
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

    // Parse file for preview instantly using SheetJS
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = e.target?.result;
      if (!data) return;

      try {
        const workbook = XLSX.read(data, { type: "array" });
        const sheetsPreview: SheetPreviewData[] = workbook.SheetNames.map(
          (sheetName) => {
            const worksheet = workbook.Sheets[sheetName];
            const sheetData = XLSX.utils.sheet_to_json(worksheet, {
              header: 1,
              defval: "", // fill empty cells
              raw: false, // format dates properly
            }) as (string | number | null)[][];
            return { sheetName, data: sheetData };
          }
        );
        onPreview(sheetsPreview);
      } catch (error) {
        console.error("Failed to parse Excel file:", error);
        setError("Failed to parse Excel file.");
        onPreview([]);
      }
    };
    reader.readAsArrayBuffer(file);
  }

  // ================= PPT Upload =================
  function validateAndSetPptFile(file: File) {
    if (!allowedPptTypes.includes(file.type)) {
      setError("Only .pptx files are supported.");
      setPptFile(null);
      return;
    }
    if (file.size > maxSizeMB * 1024 * 1024) {
      setError("PPT file size must be less than 5MB.");
      setPptFile(null);
      return;
    }
    setError(null);
    setPptFile(file);
  }

  // Handlers for drag & drop etc.
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

  const handlePptChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    const files = event.target.files;
    if (files && files[0]) validateAndSetPptFile(files[0]);
  };

  // ================= API Upload =================
  async function handleGenerateReport() {
    if (!selectedFile) {
      setError("Please upload a valid Excel file.");
      return;
    }
    if (!pptFile) {
      setError("Please upload a PPT template.");
      return;
    }

    setError(null);
    setUploading(true);
    setUploadProgress(0);

    const formData = new FormData();
    formData.append("excel", selectedFile); // Excel
    formData.append("ppt", pptFile); // PPT

    try {
      const response = await axios.post(
        "http://localhost:5000/api/upload-report",
        formData,
        {
          headers: { "Content-Type": "multipart/form-data" },
          onUploadProgress: (progressEvent) => {
            let percentComplete = 0;
            if (progressEvent.total) {
              percentComplete = Math.round(
                (progressEvent.loaded * 100) / progressEvent.total
              );
            }
            setUploadProgress(percentComplete);
          },
        }
      );

      // Backend responds with a download URL
      const { excelDownloadUrl, pptDownloadUrl } = response.data;

      // Auto-download both
      [excelDownloadUrl, pptDownloadUrl].forEach((url) => {
        const link = document.createElement("a");
        link.href = url;
        link.setAttribute("download", ""); // backend already sets filename
        document.body.appendChild(link);
        link.click();
        link.remove();
      });

    } catch (error) {
      console.error("Upload error:", error);
      setError("Failed to upload file and generate report.");
    } finally {
      setUploading(false);
      setUploadProgress(0);
    }
  }

  // ================= UI =================
  return (
    <div className="max-w-xl mx-auto p-6 bg-white shadow-lg rounded-2xl">
      <h2 className="text-xl font-bold text-gray-800 mb-4">
        Upload Files
      </h2>

      {/* Excel Upload Box */}
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
          disabled={uploading}
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

      {/* PPT Upload Box */}
      <div
        className="border-2 border-dashed rounded-lg p-6 text-center cursor-pointer mt-4"
        onClick={() => pptInputRef.current?.click()}
      >
        <input
          ref={pptInputRef}
          type="file"
          accept=".pptx"
          onChange={handlePptChange}
          disabled={uploading}
          hidden
        />
        <p className="text-gray-600 font-medium">
          {pptFile ? (
            <span className="text-green-700">{pptFile.name}</span>
          ) : (
            "Click to upload PPT Template (.pptx)"
          )}
        </p>
        <p className="text-sm text-gray-500 mt-1">(.pptx only, max 5MB)</p>
      </div>

      {/* Errors */}
      {error && <p className="text-red-600 text-sm mt-3">⚠ {error}</p>}

      {/* Progress Bar */}
      {uploading && (
        <div className="w-full bg-gray-200 rounded-full h-3 mt-4 overflow-hidden">
          <div
            className="bg-blue-600 h-3 text-xs text-white flex items-center justify-center transition-all duration-300"
            style={{ width: `${uploadProgress}%` }}
          >
            {uploadProgress}%
          </div>
        </div>
      )}

      {/* Generate Report Button */}
      <div className="flex justify-end mt-6">
        <button
          disabled={!selectedFile || !pptFile || uploading}
          onClick={handleGenerateReport}
          className={`px-5 py-2 rounded-lg font-semibold text-white transition-colors duration-200 ${
            !selectedFile || !pptFile || uploading
              ? "bg-gray-400 cursor-not-allowed"
              : "bg-blue-600 hover:bg-blue-700"
          }`}
        >
          {uploading ? "Generating Report..." : "Generate Report"}
        </button>
      </div>
    </div>
  );
};

export default FileUpload;
