// import React from "react";

// export type PreviewData = { [fileName: string]: string };

// interface Props {
//   preview: PreviewData;
// }

// const ExcelPreview: React.FC<Props> = ({ preview }) => {
//   if (!preview || Object.keys(preview).length === 0) return null;

//   return (
//     <div className="mt-10 space-y-6 max-w-4xl mx-auto overflow-auto bg-white p-4 rounded-xl shadow">
//       {Object.entries(preview).map(([xmlFile, content]) => (
//         <div key={xmlFile} className="mb-4">
//           <h3 className="font-semibold mb-2">{xmlFile}</h3>
//           <pre className="text-xs max-h-40 overflow-auto p-2 bg-gray-100 rounded whitespace-pre-wrap border border-gray-300">
//             {content}
//           </pre>
//         </div>
//       ))}
//     </div>
//   );
// };

// export default ExcelPreview;

import React, { useState } from "react";

export type SheetPreviewData = {
  sheetName: string;
  data: Array<Array<string | number | null>>;
};

interface Props {
  preview: SheetPreviewData[];
  onGenerateReport: () => void;
  onGeneratePpt: () => void;
  loadingReport: boolean;
  loadingPpt: boolean;
}

const ExcelPreviewTabs: React.FC<Props> = ({
  preview,
  onGenerateReport,
  loadingReport,
  loadingPpt
}) => {
  const [activeIndex, setActiveIndex] = useState(0);

  if (!preview || preview.length === 0) return null;

  return (
    <div className="max-w-6xl mx-auto mt-8">
      {/* Page Title */}
      <h1 className="text-3xl font-extrabold text-center text-gray-900 mb-6">
        Preview
      </h1>

      {/* Card */}
      <div className="bg-white rounded-2xl shadow-lg p-6">
        {/* Upload & Button Panel */}
        <div className="flex flex-col md:flex-row items-center justify-between mb-4 gap-2">
          <span className="text-lg font-semibold text-gray-700">
            Excel File Preview
          </span>
          <div className="flex gap-2">
            <button
              className={`px-5 py-2 rounded-xl font-semibold text-white ${
                loadingReport
                  ? "bg-gray-400 cursor-wait"
                  : "bg-blue-600 hover:bg-blue-700"
              } shadow`}
              disabled={loadingReport}
              onClick={onGenerateReport}
            >
              {loadingPpt ? "Generating..." : "Generate PPT"}
            </button>
          </div>
        </div>

        {/* Sheet Tabs */}
        <div className="overflow-x-auto border-b border-gray-200 mb-4">
          <ul className="flex space-x-2">
            {preview.map((sheet, idx) => (
              <li key={sheet.sheetName}>
                <button
                  className={`px-4 py-2 rounded-t-lg font-semibold ${
                    idx === activeIndex
                      ? "bg-blue-100 border-l border-r border-t border-blue-400 text-blue-700"
                      : "text-gray-600 hover:bg-gray-100"
                  }`}
                  onClick={() => setActiveIndex(idx)}
                >
                  {sheet.sheetName}
                </button>
              </li>
            ))}
          </ul>
        </div>

        {/* Sheet Table Panel */}
        <div className="overflow-x-auto">
          <table className="min-w-full border border-gray-300">
            <thead className="bg-gray-100 sticky top-0">
              <tr>
                {preview[activeIndex].data[0]?.map((header, idx) => (
                  <th
                    key={idx}
                    className="px-4 py-2 border-r border-gray-200 font-bold text-gray-800"
                  >
                    {header ?? ""}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {preview[activeIndex].data.slice(1).map((row, i) => (
                <tr key={i} className="border-b border-gray-100">
                  {row.map((cell, j) => (
                    <td
                      key={j}
                      className="px-4 py-2 border-r border-gray-100 whitespace-nowrap"
                    >
                      {cell ?? ""}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};

export default ExcelPreviewTabs;
