import os
import pandas as pd
from utils.text_utils import norm_text

def merge_if_updated(old_path, new_df):
    new_df = new_df.copy()
    if new_df.empty:
        if os.path.exists(old_path):
            old_df = pd.read_excel(old_path)
            return old_df, False
        return new_df, False

    new_df["item"] = new_df["item"].map(norm_text)
    new_df["amount"] = pd.to_numeric(new_df["amount"], errors="coerce").round(2)

    if not os.path.exists(old_path):
        return new_df, True

    old_df = pd.read_excel(old_path)
    if old_df.empty:
        return new_df, True

    old_df["item"] = old_df["item"].map(norm_text)
    old_df["amount"] = pd.to_numeric(old_df["amount"], errors="coerce").round(2)

    merged = (
        pd.concat([old_df, new_df], ignore_index=True)
        .sort_values(["company_id", "year", "item"])
        .drop_duplicates(["company_id", "year", "item"], keep="last")
        .reset_index(drop=True)
    )

    return merged, not merged.equals(old_df)