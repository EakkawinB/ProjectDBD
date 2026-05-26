import re, pandas as pd

def norm_text(s):
    if s is None:
        return ""
    s = str(s).replace("\u00a0", " ").replace("\u200b", " ")
    return re.sub(r"\s+", " ", s).strip()

def to_float(txt):
    if txt is None:
        return None
    t = norm_text(txt)
    if t in ["", "-", "-"]:
        return None
    neg = False
    if t.startswith("(") and t.endswith(")"):
        neg = True
        t = t[1:-1].strip()
    t = t.replace(",", "")
    m = re.search(r"-?\d+(\.\d+)?", t)
    if not m:
        return None
    try:
        v = float(m.group(0))
        if neg:
            v = -v
        return round(v, 2)
    except:
        return None

def looks_like_person_name(name):
    if not name:
        return False
    name = name.strip()
    prefixes = ["นาย", "นาง", "นางสาว", "Mr.", "Mrs.", "Ms."]
    return any(name.startswith(p) for p in prefixes)



def normalize_company_id(cid_raw):
    if cid_raw is None or pd.isna(cid_raw):
        return None

    cid = str(cid_raw).strip()

    if cid.endswith(".0"):
        cid = cid[:-2]

    if not cid.isdigit() or len(cid) != 13:
        return None

    return cid