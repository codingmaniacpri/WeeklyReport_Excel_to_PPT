import openpyxl
import json
import os

def read_visible_excel_sheets_to_json(file_path, output_folder="extracted_json"):
    wb = openpyxl.load_workbook(file_path, data_only=True)
    os.makedirs(output_folder, exist_ok=True)
    saved_files = []

    # Only process visible sheets (skip hidden sheets like "QA Review")
    visible_sheets = [ws for ws in wb.worksheets if ws.sheet_state == "visible"]

    for ws in visible_sheets:
        sheet_name = ws.title

        # Find the first row with actual headers (often row 1, but sometimes more)
        header_row_idx = 1
        for i in range(1, ws.max_row + 1):
            values = [cell.value for cell in ws[i]]
            # A good header row should have at least one non-blank and not all "Unnamed..."
            cleaned_header = [
                str(val).strip() if val is not None else "" for val in values
            ]
            valid_headers = [
                h for h in cleaned_header
                if h and not h.lower().startswith("unnamed")
            ]
            if len(valid_headers) >= 2:  # two or more real headers
                header_row_idx = i
                headers = cleaned_header
                break
        else:
            # No header found, skip
            print(f"Sheet '{sheet_name}' skipped: header not found.")
            continue

        # Clean headers: drop columns if header is blank or startswith Unnamed
        keep_indices = [
            idx for idx, h in enumerate(headers)
            if h and not h.lower().startswith("unnamed")
        ]
        final_headers = [headers[idx] for idx in keep_indices]

        # Debug print
        print(f"Extracting '{sheet_name}': using header row {header_row_idx}: {final_headers}")

        # Go through the data rows as in Excel
        all_rows = []
        for row in ws.iter_rows(min_row=header_row_idx+1, max_row=ws.max_row, max_col=len(headers)):
            # Only keep the columns as per final_headers
            row_data = {}
            for pos, idx in enumerate(keep_indices):
                cell = row[idx]
                cell_val = cell.value
                style_info = {}
                if cell.has_style:
                    font = cell.font
                    fill = cell.fill
                    style_info = {
                        "font_bold": font.bold,
                        "font_italic": font.italic,
                        "font_underline": font.underline,
                        "font_color": font.color.rgb if font.color else None,
                        "fill_color": fill.fgColor.rgb if fill.fgColor.type == 'rgb' else None,
                        "number_format": cell.number_format,
                        "alignment_horizontal": cell.alignment.horizontal,
                        "alignment_vertical": cell.alignment.vertical,
                    }
                row_data[final_headers[pos]] = {
                    "value": cell_val,
                    "style": style_info
                }
            # Only add row if it has at least one non-blank value in kept columns
            if any(
                row_data[header]["value"] is not None and str(row_data[header]["value"]).strip() != ""
                for header in final_headers
            ):
                all_rows.append(row_data)

        json_path = os.path.join(output_folder, f"{sheet_name}.json")
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(all_rows, f, indent=2, default=str)
        saved_files.append(json_path)
        print(f"Saved JSON for visible sheet '{sheet_name}' at: {json_path}")

    return saved_files
