"""
STEP 1 - Extract the fields we need from the two openFDA zip files into small CSV files.

Why: the recall file unzips to about 1.5 GB of JSON, too big to open in Excel or load at once.
This script reads each zip as a stream, one record at a time, and keeps only the useful fields.

Input : data/device-recall-0001-of-0001_json.zip
        data/device-enforcement-0001-of-0001_json.zip
Output: data/recalls.csv       (one row per recalled product, with root cause)
        data/enforcement.csv   (one row per recalled product, with Class I / II / III)
Run   : python step1_extract.py        (takes 1-3 minutes)
"""
import csv
import zipfile

import ijson

RECALL_ZIP = "data/device-recall-0001-of-0001_json.zip"
ENFORCE_ZIP = "data/device-enforcement-0001-of-0001_json.zip"


def stream_records(zip_path):
    """Yield each record in the 'results' list of the single JSON file inside a zip."""
    with zipfile.ZipFile(zip_path) as z:
        name = z.namelist()[0]
        with z.open(name) as f:
            yield from ijson.items(f, "results.item")


def first(value):
    """openFDA stores some fields as lists; keep the first item as text."""
    if isinstance(value, list):
        return value[0] if value else ""
    return value if value is not None else ""


# ---- recall file: root cause, product code, device class, dates
recall_cols = ["recall_number", "event_id", "date_initiated", "recalling_firm", "product_code",
               "device_name", "device_class", "medical_specialty", "root_cause", "recall_status",
               "product_quantity", "reason_for_recall"]
n = 0
with open("data/recalls.csv", "w", newline="", encoding="utf-8") as out:
    w = csv.writer(out)
    w.writerow(recall_cols)
    for r in stream_records(RECALL_ZIP):
        o = r.get("openfda", {}) or {}
        w.writerow([
            r.get("product_res_number", ""),
            r.get("res_event_number", ""),
            r.get("event_date_initiated", ""),
            r.get("recalling_firm", ""),
            r.get("product_code", ""),
            first(o.get("device_name")),
            first(o.get("device_class")),
            first(o.get("medical_specialty_description")),
            r.get("root_cause_description", ""),
            r.get("recall_status", ""),
            r.get("product_quantity", ""),
            " ".join(str(r.get("reason_for_recall", "")).split())[:300],
        ])
        n += 1
print(f"recalls.csv written: {n:,} rows")

# ---- enforcement file: FDA classification (Class I / II / III)
enf_cols = ["recall_number", "classification", "recall_initiation_date", "status", "voluntary_mandated"]
n = 0
with open("data/enforcement.csv", "w", newline="", encoding="utf-8") as out:
    w = csv.writer(out)
    w.writerow(enf_cols)
    for r in stream_records(ENFORCE_ZIP):
        w.writerow([r.get("recall_number", ""), r.get("classification", ""),
                    r.get("recall_initiation_date", ""), r.get("status", ""),
                    r.get("voluntary_mandated", "")])
        n += 1
print(f"enforcement.csv written: {n:,} rows")
