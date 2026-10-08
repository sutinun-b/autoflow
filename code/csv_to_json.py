"""แปลงไฟล์ CSV ของ AUTO FLOW (autoflow_log.csv / results.csv) เป็น JSON

วิธีใช้:
    python csv_to_json.py autoflow_log.csv            -> ได้ autoflow_log.json ข้างกัน
    python csv_to_json.py a.csv b.csv -o ออกที่โฟลเดอร์นี้

รองรับ:
- ไฟล์ UTF-8 ที่มี BOM (แบบที่ AUTO FLOW เขียน เพื่อให้ Excel อ่านภาษาไทยได้)
- แถวสรุปท้ายไฟล์ (หลังบรรทัดว่าง เช่น "เวลารวม, 581.3 วินาที") -> แยกไปไว้ใน "summary"
- ตัวเลขถูกแปลงเป็น number อัตโนมัติ ช่องว่างเป็น null
"""
import argparse
import csv
import json
import os
import re

_NUM = re.compile(r"^-?\d+(\.\d+)?$")


def _val(s):
    s = (s or "").strip()
    if s == "":
        return None
    if _NUM.match(s):
        return float(s) if "." in s else int(s)
    return s


def csv_to_obj(path):
    with open(path, encoding="utf-8-sig", newline="") as f:
        rows = list(csv.reader(f))
    if not rows:
        return {"source": os.path.basename(path), "rows": [], "summary": {}}
    head = [h.strip() for h in rows[0]]
    data, summary, tail, fresh = [], {}, False, False
    for r in rows[1:]:
        if not any(c.strip() for c in r):
            tail, fresh = True, True         # บรรทัดว่าง = เริ่มส่วนสรุป
            continue
        filled = sum(1 for c in r if c.strip())
        if tail and filled <= 2 and r[0].strip():
            summary[r[0].strip()] = _val(r[1] if len(r) > 1 else "")   # แถวสรุปแบบ "ชื่อ, ค่า"
            fresh = False
            continue
        if tail and fresh:
            head = [h.strip() for h in r]        # ตารางที่ 2 ในไฟล์เดียวกัน (เช่น ตารางสรุปรุ่น)
            summary.setdefault("tables", []).append({"columns": head, "rows": []})
            fresh = False
            continue
        item = {head[i] if i < len(head) else f"col{i+1}": _val(c) for i, c in enumerate(r)}
        if tail and summary.get("tables"):
            summary["tables"][-1]["rows"].append(item)
        else:
            data.append(item)
    return {"source": os.path.basename(path), "columns": [h for h in rows[0]], "rows": data, "summary": summary}


def main():
    ap = argparse.ArgumentParser(description="แปลง CSV ของ AUTO FLOW เป็น JSON")
    ap.add_argument("files", nargs="+")
    ap.add_argument("-o", "--out", default=None, help="โฟลเดอร์ปลายทาง (ไม่ใส่ = ข้างไฟล์ต้นทาง)")
    a = ap.parse_args()
    for p in a.files:
        obj = csv_to_obj(p)
        out_dir = a.out or os.path.dirname(os.path.abspath(p))
        os.makedirs(out_dir, exist_ok=True)
        out = os.path.join(out_dir, os.path.splitext(os.path.basename(p))[0] + ".json")
        with open(out, "w", encoding="utf-8") as f:
            json.dump(obj, f, ensure_ascii=False, indent=2)
        print(f"{p} -> {out}  ({len(obj['rows'])} แถว)")


if __name__ == "__main__":
    main()
