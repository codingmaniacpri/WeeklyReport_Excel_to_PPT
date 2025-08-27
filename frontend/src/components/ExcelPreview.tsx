import React, { useState } from "react";

/**
 * Type definition for sheet preview data
 */
export type SheetPreviewData = {
  sheetName: string;
  data: Array<Array<string | number | null>>;
};

/**
 * Props for ExcelPreviewTabs component
 */
interface Props {
  preview: SheetPreviewData[];
  onGenerateReport: () => void;
  onGeneratePpt: () => void;
  loadingReport: boolean;
  loadingPpt: boolean;
}

/**
 * Table view for sheet data
 */
const SheetTable: React.FC<{ data: Array<Array<string | number | null>> }> = ({
  data,
}) => {
  if (!data || data.length === 0) {
    return (
      <p className="text-gray-500 text-center py-6">No data available</p>
    );
  }

  const [headers, ...rows] = data;

  return (
    <div className="overflow-x-auto">
      <table className="min-w-full border border-gray-300 text-sm">
        <thead className="bg-gray-100 sticky top-0">
          <tr>
            {headers.map((header, idx) => (
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
          {rows.map((row, i) => (
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
  );
};

/**
 * Main Excel Preview Tabs component
 */
const ExcelPreviewTabs: React.FC<Props> = ({
  preview
}) => {
  const [activeIndex, setActiveIndex] = useState(0);

  if (!preview || preview.length === 0) return null;

  return (
    <div className="max-w-6xl mx-auto mt-8">
      {/* Page Title */}
      <h1 className="text-3xl font-extrabold text-center text-gray-900 mb-6">
        Excel Preview
      </h1>

      {/* Card */}
      <div className="bg-white rounded-2xl shadow-lg p-6">
        {/* Info Panel */}
        <div className="flex flex-col md:flex-row items-center justify-between mb-4 gap-2">
          <span className="text-lg font-semibold text-gray-700">
            Excel File Preview
          </span>
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

        {/* Active Sheet Table */}
        <SheetTable data={preview[activeIndex].data} />
      </div>
    </div>
  );
};

export default ExcelPreviewTabs;
