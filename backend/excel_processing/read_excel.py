import pandas as pd
import re

def get_latest_date_comments(file_path):
    df = pd.read_excel(file_path, sheet_name="Dependencies", header=1)

    if df.columns[0].startswith("Unnamed"):
        df = df.drop(df.columns[0], axis=1)

    def filter_latest_comment(text):
        if pd.isna(text):
            return text
        comments = str(text).split("\n")
        date_map = {}
        for c in comments:
            dates = re.findall(r"\b\d{2}/\d{2}\b", c)
            if dates:
                for d in dates:
                    date_map[d] = date_map.get(d, []) + [c]
        if not date_map:
            return text
        latest_date = max(date_map.keys(), key=lambda d: (int(d.split("/")[0]), int(d.split("/")[1])))
        return "\n".join(date_map[latest_date])

    processed_df = df.copy()
    processed_df["Comments"] = processed_df["Comments"].apply(filter_latest_comment)
    return processed_df, processed_df["Comments"]
