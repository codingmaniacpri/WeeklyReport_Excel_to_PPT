import { useState } from 'react';
import * as XLSX from 'xlsx';

export type PreviewData = { [sheetName: string]: string[][] };

export function useExcelPreview() {
  const [preview, setPreview] = useState<PreviewData>({});

  function generatePreview(file: File) {
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = e.target?.result;
      if (data) {
        const workbook = XLSX.read(data, { type: 'binary' });
        const previewData: PreviewData = {};
        workbook.SheetNames.forEach((sheet) => {
          const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheet], { header: 1 });
          previewData[sheet] = (rows as string[][]).slice(0, 6);
        });
        setPreview(previewData);
      }
    };
    reader.readAsBinaryString(file);
  }

  return { preview, generatePreview };
}
