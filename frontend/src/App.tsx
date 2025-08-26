// // export default App
// import { useState } from 'react';
// import FileUpload from './components/FileUpload';
// // import Dashboard from './components/Dashboard/Dashboard';
// // import Progress from './components/Progress';

// function App() {
  
//   // Dashboard inputs state
//   const [companyName, ] = useState('');
//   const [projectTitle, ] = useState('');
//   const [logoFile, ] = useState<File | null>(null);

//   // URL of generated PPT download
//   const [downloadUrl, setDownloadUrl] = useState('');

//   const handleDownload = (url: string) => {
//     setDownloadUrl(url);
//     window.open(url, '_blank');
//   };

//   return (
//     <>
//       <div className="container mx-auto p-6">
//         <h1 className="text-3xl font-bold mb-6">Welcome to Weekly Report Generator</h1>

//         {/* Dashboard to input company/project details
//         <Dashboard
//           companyName={companyName}
//           setCompanyName={setCompanyName}
//           projectTitle={projectTitle}
//           setProjectTitle={setProjectTitle}
//           logoFile={logoFile}
//           setLogoFile={setLogoFile}
//         /> */}

//         <FileUpload
//           companyName={companyName}
//           projectTitle={projectTitle}
//           logoFile={logoFile}
//           onDownload={handleDownload}
//         />

//         {/* Download link */}
//         {downloadUrl && (
//           <a
//             href={downloadUrl}
//             className="mt-4 text-blue-600 underline block"
//             download
//           >
//             Download Generated PPTX
//           </a>
//         )}
//       </div>
//     </>
//   );
// }

// export default App;

import { useState } from 'react';
import FileUpload from './components/FileUpload';
import ExcelPreview from './components/ExcelPreview';
import type { SheetPreviewData } from './components/FileUpload';
// import LogSidebar from './components/LogSidebar';

function App() {
  // For excel file preview data from FileUpload
  const [preview, setPreview] = useState<SheetPreviewData[]>([]);

  // Download url after backend processing
  const [downloadUrl, setDownloadUrl] = useState('');

  // Handle download link
  const handleDownload = (url: string) => {
    setDownloadUrl(url);
    window.open(url, '_blank');
  };

  return (
    <div className="flex min-h-screen bg-gradient-to-tr from-gray-200 via-blue-100 to-purple-100">
      {/* Sidebar with live logs */}
      

      {/* Main area */}
      <main className="flex-1 p-8">
        <h1 className="text-4xl font-bold mb-8">
          Welcome to Weekly Report Generator
        </h1>

        {/* Excel file upload */}
        <FileUpload onPreview={setPreview} onDownload={handleDownload} companyName={''} projectTitle={''} logoFile={null} />

        {/* Excel preview table */}
        {preview.length > 0 && (
          <ExcelPreview preview={preview} onGenerateReport={function (): void {
            throw new Error('Function not implemented.');
          } } onGeneratePpt={function (): void {
            throw new Error('Function not implemented.');
          } } loadingReport={false} loadingPpt={false} />
        )}

        {/* Download link */}
        {downloadUrl && (
          <a
            href={downloadUrl}
            className="mt-4 text-blue-600 underline block"
            download
          >
            Download Generated PPTX
          </a>
        )}
      </main>
    </div>
  );
}

export default App;
