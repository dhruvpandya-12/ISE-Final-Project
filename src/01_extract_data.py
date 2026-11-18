import csv
import zipfile
from pathlib import Path

import ijson

DATA = Path(__file__).resolve().parents[1] / "data"
RECALL_ZIP = DATA / "device-recall-0001-of-0001_json.zip"
ENFORCE_ZIP = DATA / "device-enforcement-0001-of-0001_json.zip"


def stream_records(zip_path):
    with zipfile.ZipFile(zip_path) as z:
        name = z.namelist()[0]
        with z.open(name) as f:
            yield from ijson.items(f, "results.item")


def first(value):
    if isinstance(value, list):
        return value[0] if value else ""
    return value if value is not None else ""


recall_cols = ["recall_number", "event_id", "date_initiated", "recalling_firm", "product_code",
               "device_name", "device_class", "medical_specialty", "root_cause", "recall_status",
               "product_quantity", "reason_for_recall"]
n = 0
with open(DATA / "recalls.csv", "w", newline="", encoding="utf-8") as out:
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


enf_cols = ["recall_number", "classification", "recall_initiation_date", "status", "voluntary_mandated"]
n = 0
with open(DATA / "enforcement.csv", "w", newline="", encoding="utf-8") as out:
    w = csv.writer(out)
    w.writerow(enf_cols)
    for r in stream_records(ENFORCE_ZIP):
        w.writerow([r.get("recall_number", ""), r.get("classification", ""),
                    r.get("recall_initiation_date", ""), r.get("status", ""),
                    r.get("voluntary_mandated", "")])
        n += 1
print(f"enforcement.csv written: {n:,} rows")
