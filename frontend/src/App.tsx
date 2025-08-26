// // export default App
// import { useState } from 'react';
// import FileUpload from './components/FileUpload/FileUpload';
// import Dashboard from './components/Dashboard/Dashboard';
// // import Progress from './components/Progress';

// function App() {
//   // Dashboard inputs state
//   const [companyName, setCompanyName] = useState('');
//   const [projectTitle, setProjectTitle] = useState('');
//   const [logoFile, setLogoFile] = useState<File | null>(null);

//   // URL of generated PPT download
//   const [downloadUrl, setDownloadUrl] = useState('');

//   const handleDownload = (url: string) => {
//     setDownloadUrl(url);
//     window.open(url, '_blank');
//   };

//   return (
//     <>
//       <div className="container mx-auto p-6">
//         <h1 className="text-3xl font-bold mb-6">Vite + React + Tailwind CSS</h1>

//         {/* Dashboard to input company/project details */}
//         <Dashboard
//           companyName={companyName}
//           setCompanyName={setCompanyName}
//           projectTitle={projectTitle}
//           setProjectTitle={setProjectTitle}
//           logoFile={logoFile}
//           setLogoFile={setLogoFile}
//         />

//         {/* File upload with API integration */}
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
