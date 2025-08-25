import React, { useRef, useState } from 'react';

type Props = {
  onDownload: (url: string) => void;
};

const FileUpload: React.FC<Props> = ({ onDownload }) => {
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const [uploading, setUploading] = useState(false);
  const [progress, setProgress] = useState(0);
  const [error, setError] = useState<string | null>(null);

  // Drag & drop handler
  const handleDrop = (event: React.DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    const files = event.dataTransfer.files;
    if (files && files[0]) {
      validateAndSetFile(files[0]);
    }
  };

  // File picker handler
  const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    const files = event.target.files;
    if (files && files[0]) {
      validateAndSetFile(files[0]);
    }
  };

  // File validation logic
  function validateAndSetFile(file: File) {
    const allowedTypes = ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'];
    const maxSizeMB = 5;
    if (!allowedTypes.includes(file.type)) {
      setError('Only .xlsx files supported');
      setSelectedFile(null);
      return;
    }
    if (file.size > maxSizeMB * 1024 * 1024) {
      setError('File size must be less than 5MB');
      setSelectedFile(null);
      return;
    }
    setError(null);
    setSelectedFile(file);
  }

  // API integration: Upload
  async function handleUpload() {
    if (!selectedFile) return;
    setUploading(true);
    setProgress(0);
    setError(null);

    const formData = new FormData();
    formData.append('excel_file', selectedFile);

    const xhr = new XMLHttpRequest();
    xhr.open('POST', 'http://localhost:5000/upload_excel', true);

    xhr.upload.onprogress = (event) => {
      if (event.lengthComputable) {
        setProgress(Math.round((event.loaded / event.total) * 100));
      }
    };
    xhr.onload = function () {
      setUploading(false);
      if (xhr.status === 200) {
        // API returns a download URL or file
        onDownload(xhr.responseText);
      } else {
        setError('Upload failed. Try again.');
      }
    };
    xhr.onerror = function () {
      setUploading(false);
      setError('Network error. Try again.');
    };

    xhr.send(formData);
  }

  return (
    <div
      className="border-2 border-dashed border-gray-400 p-6 rounded-lg flex flex-col items-center mb-4"
      onDrop={handleDrop}
      onDragOver={e => e.preventDefault()}
    >
      <input
        ref={fileInputRef}
        type="file"
        accept=".xlsx"
        onChange={handleFileChange}
        hidden
      />
      <button
        className="bg-blue-600 text-white px-4 py-2 rounded mb-2"
        type="button"
        onClick={() => fileInputRef.current?.click()}
      >
        Select Excel File
      </button>
      <p className="text-gray-700">Or drag and drop your .xlsx file here</p>
      {selectedFile && (
        <p className="mt-2 text-green-700">{selectedFile.name}</p>
      )}
      {error && (
        <p className="mt-2 text-red-600">{error}</p>
      )}
      {uploading && (
        <div className="w-full bg-gray-100 rounded mt-4">
          <div className="bg-blue-500 text-xs leading-none py-1 text-center text-white rounded"
            style={{ width: `${progress}%` }}>
            {progress}%
          </div>
        </div>
      )}
      <button
        className="bg-green-600 text-white px-4 py-2 rounded mt-4"
        disabled={!selectedFile || uploading}
        type="button"
        onClick={handleUpload}
      >
        Upload
      </button>
    </div>
  );
};

export default FileUpload;
