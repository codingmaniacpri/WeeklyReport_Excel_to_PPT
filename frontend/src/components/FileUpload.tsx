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
  logoFile: File | null;
}

const FileUpload: React.FC<Props> = ({ onPreview, onDownload, companyName, logoFile }) => {
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [projectTitle, setProjectTitle] = useState("");
  const [weekRange, setWeekRange] = useState("");
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [dragOver, setDragOver] = useState(false);
  const [uploading, setUploading] = useState(false);
  const [uploadProgress, setUploadProgress] = useState(0);
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
        const sheetsPreview: SheetPreviewData[] = workbook.SheetNames.map(
          (sheetName) => {
            const worksheet = workbook.Sheets[sheetName];
            const sheetData = XLSX.utils.sheet_to_json(worksheet, {
              header: 1,
              defval: "",
              raw: false,
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

  async function handleGenerateReport() {
    if (!selectedFile) {
      setError("Please upload a valid Excel file.");
      return;
    }
    setError(null);
    setUploading(true);
    setUploadProgress(0);
    const formData = new FormData();
    formData.append("file", selectedFile);
    formData.append("projectTitle", projectTitle);
    formData.append("weekRange", weekRange);
    formData.append("companyName", companyName);
    if (logoFile) formData.append("logoFile", logoFile);
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
      const downloadUrl = response.data.downloadUrl;
      onDownload(downloadUrl);
    } catch (error) {
      console.error("Upload error:", error);
      setError("Failed to upload file and generate report.");
    } finally {
      setUploading(false);
      setUploadProgress(0);
    }
  }

  return (
    <div className="max-w-xl mx-auto p-6 bg-white shadow-lg rounded-2xl mt-4">
      <h2 className="text-xl font-bold text-gray-800 mb-4">Weekly Report Generation</h2>
      <div className="mb-4">
        <label className="block font-semibold mb-1">Project Title</label>
        <input
          className="w-full border rounded px-3 py-2"
          type="text"
          value={projectTitle}
          onChange={e => setProjectTitle(e.target.value)}
          required
        />
      </div>
      <div className="mb-4">
        <label className="block font-semibold mb-1">Week Range</label>
        <input
          className="w-full border rounded px-3 py-2"
          type="text"
          value={weekRange}
          onChange={e => setWeekRange(e.target.value)}
          placeholder="e.g. 19 Aug - 25 Aug 2025"
          required
        />
      </div>
      <div
        className={`border-2 border-dashed rounded-lg p-8 text-center cursor-pointer transition-colors duration-200 ${dragOver ? "border-blue-500 bg-blue-50" : "border-gray-300 bg-gray-50"}`}
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
      {error && <p className="text-red-600 text-sm mt-3">⚠ {error}</p>}
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
      <div className="flex justify-end mt-6">
        <button
          disabled={!selectedFile || uploading}
          onClick={handleGenerateReport}
          className={`px-5 py-2 rounded-lg font-semibold text-white transition-colors duration-200 ${
            !selectedFile || uploading
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
