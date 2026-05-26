import pandas as pd
import re
from utils.logging_utils import log
from utils.text_utils import norm_text

def _pick(df, patterns, label="", prefer_exact=True):
    if df is None or df.empty:
        s = pd.Series(dtype="float")
        s.index.name = "year"
        return s

    tmp = df.copy()
    tmp["item"] = tmp["item"].astype(str).map(norm_text)
    tmp["year"] = pd.to_numeric(tmp["year"], errors="coerce")
    tmp["amount"] = pd.to_numeric(tmp["amount"], errors="coerce")

    best = None
    best_score = -1

    for pat in patterns:
        use_regex = pat.startswith("re:")
        expr = pat[3:] if use_regex else re.escape(pat)

        m = tmp[tmp["item"].str.contains(expr, na=False, regex=True)]
        if m.empty:
            continue

        for it in m["item"].unique():
            score = 0
            it_norm = norm_text(it)

            if use_regex:
                score += 2
            if prefer_exact and (it_norm == (pat[3:] if use_regex else pat)):
                score += 10

            for p in patterns:
                p_expr = p[3:] if p.startswith("re:") else p
                if re.search(p_expr if p.startswith("re:") else re.escape(p_expr), it_norm):
                    score += 1

            if score > best_score:
                best_score = score
                best = it

    if best is None:
        if label:
            log(f"⚠️ _pick missing: {label}")
        s = pd.Series(dtype="float")
        s.index.name = "year"
        return s

    s = tmp[tmp["item"] == best].groupby("year")["amount"].last().sort_index()
    s.index.name = "year"
    if label:
        log(f"✔ _pick {label}: '{best}'")
    return s