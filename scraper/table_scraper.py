from utils.text_utils import norm_text, to_float
import pandas as pd

def get_years(page):
    ths = page.locator("table thead tr").first.locator("th")
    years = []
    for i in range(ths.count()):
        txt = norm_text(ths.nth(i).text_content())
        if txt.isdigit():
            years.append(int(txt))
    return years


def scrape_table(page, cid, years):
    records = []
    rows = page.locator("table tbody tr")

    for r in range(rows.count()):
        row = rows.nth(r)

        ths = row.locator("th")
        tds = row.locator("td")

        item = ""
        offset = 0

        if ths.count() > 0:
            item = norm_text(ths.first.text_content())
            offset = 0
        else:
            if tds.count() == 0:
                continue
            item = norm_text(tds.first.text_content())
            offset = 1

        if not item:
            continue

        td_count = tds.count()
        if td_count <= offset:
            continue

        data_count = td_count - offset
        stride = 2 if data_count >= len(years) * 2 else 1
        usable = min(len(years), data_count // stride)

        for i in range(usable):
            idx = offset + i * stride
            txt = tds.nth(idx).text_content()
            records.append({
                "company_id": cid,
                "year": years[i],
                "item": item,
                "amount": to_float(txt)
            })

    df = pd.DataFrame(records)
    if not df.empty:
        df["item"] = df["item"].map(norm_text)
        df = df.drop_duplicates(["company_id", "year", "item"])
    return df