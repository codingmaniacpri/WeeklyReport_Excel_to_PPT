import React, { useState } from "react";

// Data type for a single sheet's preview
export type SheetPreviewData = {
  sheetName: string;
  data: Array<Array<string | number | null>>;
};

interface Props {
  preview: SheetPreviewData[];
  onGenerateReport?: () => void;
  onGeneratePpt?: () => void;
  loadingReport?: boolean;
  loadingPpt?: boolean;
}

// Table component for rendering sheet data neatly
const SheetTable: React.FC<{ data: Array<Array<string | number | null>> }> = ({ data }) => {
  // Handle empty or undefined data
  if (!data || data.length === 0) {
    return <p className="text-center text-gray-500 p-6">No data available</p>;
  }

  // Separate headers and body rows
  const [headers, ...rows] = data;

  return (
    <div className="overflow-x-auto">
      <table className="min-w-full border border-gray-300 text-sm">
        <thead className="bg-blue-50 sticky top-0 z-10">
          <tr>
            {headers.map((header, idx) => (
              <th
                key={idx}
                className="px-4 py-2 border-r border-gray-200 font-semibold text-blue-800 whitespace-nowrap"
              >
                {header ?? ""}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.map((row, i) => (
            <tr
              key={i}
              className={i % 2 === 0 ? "bg-white" : "bg-blue-50"} // alternating row colors
            >
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
  );
};

const ExcelPreviewTabs: React.FC<Props> = ({
  preview
}) => {
  // Track which sheet is currently active (default first)
  const [activeIndex, setActiveIndex] = useState(0);

  // Hide component if no preview data
  if (!preview || preview.length === 0) return null;

  return (
    <div className="max-w-7xl mx-auto mt-10 p-4 transition-opacity duration-300">
      <h2 className="text-3xl font-extrabold text-center text-gray-900 mb-6">
        Excel Preview
      </h2>
      <div className="bg-white shadow-lg rounded-2xl p-6">
        {/* Sheet Tabs */}
        <div className="mb-6 overflow-x-auto border-b border-gray-200">
          <nav className="flex space-x-4">
            {preview.map((sheet, idx) => (
              <button
                key={sheet.sheetName}
                className={`px-4 py-2 rounded-t-lg font-semibold whitespace-nowrap transition-colors duration-150 ${
                  idx === activeIndex
                    ? "bg-blue-100 text-blue-700 border border-blue-400 shadow"
                    : "text-gray-500 hover:bg-gray-100"
                }`}
                onClick={() => setActiveIndex(idx)}
                aria-current={idx === activeIndex ? "page" : undefined}
              >
                {sheet.sheetName}
              </button>
            ))}
          </nav>
        </div>

        {/* Active Sheet Table */}
        <SheetTable data={preview[activeIndex].data} />
      </div>
    </div>
  );
};

export default ExcelPreviewTabs;
