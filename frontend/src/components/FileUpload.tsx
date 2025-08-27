import React, { useRef, useState } from "react";
import * as XLSX from "xlsx";
import axios from "axios";
import DatePicker from "react-datepicker";
import "react-datepicker/dist/react-datepicker.css";

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

const FileUpload: React.FC<Props> = ({
  onPreview,
  companyName,
  logoFile,
}) => {
  // refs for hidden file inputs
  const excelInputRef = useRef<HTMLInputElement>(null);
  const pptInputRef = useRef<HTMLInputElement>(null);

  // input states
  const [projectTitle, setProjectTitle] = useState("");
  const [startDate, setStartDate] = useState<Date | null>(null);
  const [endDate, setEndDate] = useState<Date | null>(null);
  const [weekRange, setWeekRange] = useState<string>("");

  // file states
  const [excelFile, setExcelFile] = useState<File | null>(null);
  const [pptFile, setPptFile] = useState<File | null>(null);

  // UI states for drag & drop, progress and errors
  const [dragOver, setDragOver] = useState(false);
  const [uploading, setUploading] = useState(false);
  const [uploadProgress, setUploadProgress] = useState(0);
  const [error, setError] = useState<string | null>(null);

  // allowed mime types and max sizes for validation
  const allowedExcelTypes = [
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  ];
  const allowedPptTypes = [
    "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  ];
  const maxExcelSizeMB = 5;
  const maxPptSizeMB = 5;

  // Validates and sets Excel file, also parses the sheet for preview
  function validateAndSetExcelFile(file: File) {
    if (!allowedExcelTypes.includes(file.type)) {
      setError("Only .xlsx files are supported.");
      setExcelFile(null);
      onPreview([]);
      return;
    }
    if (file.size > maxExcelSizeMB * 1024 * 1024) {
      setError("Excel file must be less than 5MB.");
      setExcelFile(null);
      onPreview([]);
      return;
    }
    setError(null);
    setExcelFile(file);

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
  // Validates and sets PPT template file
  function validateAndSetPptFile(file: File) {
    if (!allowedPptTypes.includes(file.type)) {
      setError("Only .pptx files are supported.");
      setPptFile(null);
      return;
    }
    if (file.size > maxPptSizeMB * 1024 * 1024) {
      setError("PPT file must be less than 5MB.");
      setPptFile(null);
      return;
    }
    setError(null);
    setPptFile(file);
  }

  // Handlers to manage drag & drop and file input changes for Excel
  const handleDropExcel = (event: React.DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    setDragOver(false);
    const files = event.dataTransfer.files;
    if (files && files[0]) validateAndSetExcelFile(files[0]);
  };
  const handleDragOverExcel = (event: React.DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    setDragOver(true);
  };
  const handleDragLeaveExcel = () => setDragOver(false);
  const handleExcelChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    const files = event.target.files;
    if (files && files[0]) validateAndSetExcelFile(files[0]);
  };
  // Handler for PPT file change
  const handlePptChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    const files = event.target.files;
    if (files && files[0]) validateAndSetPptFile(files[0]);
  };

  // Handle report generation upload with validation and progress
  async function handleGenerateReport() {
    if (!excelFile) {
      setError("Please upload a valid Excel file.");
      return;
    }
    if (!pptFile) {
      setError("Please upload a PPT template.");
      return;
    }
    if (!projectTitle.trim()) {
      setError("Please enter the project title.");
      return;
    }
    if (!weekRange.trim()) {
      setError("Please select a week range.");
      return;
    }
    setError(null);
    setUploading(true);
    setUploadProgress(0);
    const formData = new FormData();
    formData.append("excel", excelFile);
    formData.append("ppt", pptFile);
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
            if (progressEvent.total) {
              setUploadProgress(
                Math.round((progressEvent.loaded * 100) / progressEvent.total)
              );
            }
          },
        }
      );
      const { excelDownloadUrl, pptDownloadUrl } = response.data;
      [excelDownloadUrl, pptDownloadUrl].forEach((url) => {
        if (url) {
          const link = document.createElement("a");
          link.href = url;
          link.setAttribute("download", ""); // Let server suggest filename
          document.body.appendChild(link);
          link.click();
          link.remove();
        }
      });
    } catch (error) {
      console.error("Upload error:", error);
      setError("Failed to upload file and generate report.");
    } finally {
      setUploading(false);
      setUploadProgress(0);
    }
  }

  // JSX rendering with labels, inputs, file upload areas, error & progress feedback
  return (
    <div className="max-w-xl mx-auto p-6 bg-white shadow-lg rounded-2xl mt-4">
      <h2 className="text-xl font-bold text-gray-800 mb-4">
        Weekly Report Generation
      </h2>

      {/* Text input: Project Title */}
      <div className="mb-4">
        <label className="block font-semibold mb-1">Project Title</label>
        <input
          type="text"
          className="w-full border rounded px-3 py-2"
          value={projectTitle}
          onChange={(e) => setProjectTitle(e.target.value)}
          required
          placeholder="Enter your project title"
        />
      </div>

      {/* Date range picker: Week Range */}
      <div className="mb-4">
        <label className="block font-semibold mb-1">Week Range</label>
        <DatePicker
          wrapperClassName="w-full"
          selectsRange
          startDate={startDate}
          endDate={endDate}
          onChange={(dates: [Date | null, Date | null]) => {
            const [start, end] = dates;
            setStartDate(start);
            setEndDate(end);
            if (start && end) {
              const formatDate = (d: Date) =>
                d.toLocaleDateString("en-GB", {
                  day: "2-digit",
                  month: "short",
                  year: "numeric",
                });
              setWeekRange(`${formatDate(start)} - ${formatDate(end)}`);
            } else {
              setWeekRange("");
            }
          }}
          isClearable
          placeholderText="Select a date range"
          className="w-full border rounded px-3 py-2"
        />
      </div>

      {/* Upload Excel file */}
      <p className="font-semibold mb-2 text-center text-gray-800 flex items-center justify-center gap-2">
        📑 Upload Excel File
      </p>
      <div
        className={`border-2 border-dashed rounded-lg p-8 text-center cursor-pointer transition-colors duration-200 ${
          dragOver ? "border-blue-500 bg-blue-50" : "border-gray-300 bg-gray-50"
        }`}
        onDrop={handleDropExcel}
        onDragOver={handleDragOverExcel}
        onDragLeave={handleDragLeaveExcel}
        onClick={() => excelInputRef.current?.click()}
      >
        <input
          ref={excelInputRef}
          type="file"
          accept=".xlsx"
          onChange={handleExcelChange}
          disabled={uploading}
          hidden
        />
        <p className="text-gray-600 font-medium">
          {excelFile ? (
            <span className="text-green-700">{excelFile.name}</span>
          ) : (
            "Drag & drop Excel file here, or click to browse"
          )}
        </p>
        <p className="text-sm text-gray-500 mt-1">(.xlsx only, max 5MB)</p>
      </div>

      {/* Divider */}
      <div className="my-6">
        <div className="w-full h-0.5 bg-gradient-to-r from-transparent via-blue-200 to-transparent rounded"></div>
      </div>

      {/* Upload PPT template */}
      <p className="font-semibold mb-2 text-center text-gray-800 flex items-center justify-center gap-2">
        📊 Upload PPT Template
      </p>
      <div
        className="border-2 border-dashed rounded-lg p-6 text-center cursor-pointer mt-4 transition-colors duration-200 bg-purple-50 hover:shadow-lg"
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

      {/* Error Message */}
      {error && <p className="text-red-600 text-sm mt-3">⚠ {error}</p>}

      {/* Upload Progress Bar */}
      {uploading && (
        <div className="w-full bg-gray-200 rounded-full h-3 mt-4 overflow-hidden">
          <div
            className="bg-gradient-to-r from-blue-500 to-blue-700 h-3 text-xs text-white flex items-center justify-center transition-all duration-300"
            style={{ width: `${uploadProgress}%` }}
          >
            {uploadProgress}%
          </div>
        </div>
      )}

      {/* Generate Report Button */}
      <div className="flex justify-end mt-6">
        <button
          disabled={!excelFile || !pptFile || uploading}
          onClick={handleGenerateReport}
          className={`px-5 py-2 rounded-lg font-semibold text-white transition-colors duration-200 ${
            !excelFile || !pptFile || uploading
              ? "bg-gray-400 cursor-not-allowed"
              : "bg-blue-600 hover:bg-blue-700 shadow-lg"
          }`}
        >
          {uploading ? "Generating Report..." : "Generate Report"}
        </button>
      </div>
    </div>
  );
};

export default FileUpload;
