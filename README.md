# AUTO FLOW · AI คัดรูปจอโฆษณา → รายงาน PowerPoint

หน้าเว็บโปรเจกต์: https://freerannnnnn01-maker.github.io/autoflow/

AUTO FLOW เป็นโปรแกรมบน Windows (Python/Tkinter) ที่ใช้ Google Gemini ดูรูปถ่ายจอโฆษณา
แล้วทำงานต่อจนได้ PowerPoint รายงานประจำเดือนในปุ่มเดียว:

- เลือกรูปที่ดีที่สุดของแต่ละจอ
- เช็กไม่ให้ใช้ซีนซ้ำกับจอก่อนหน้า
- แนะนำการแต่งสี
- ทำ PowerPoint รายงานประจำเดือน

| โฟลเดอร์ | เนื้อหา |
|---|---|
| `index.html` | หน้าเว็บโปรเจกต์ (Use case 6 องค์ประกอบ · Workflow · Prompt · ผลทดลอง · log) |
| `code/auto_flow.py` | ซอร์สโค้ด AUTO FLOW |
| `code/csv_to_json.py` | สคริปต์แปลง log CSV เป็น JSON |
| `data/workflow.json` | Workflow 8 ขั้น · AI · กฎ · Human-in-the-loop |
| `data/model_test_results.*` | ผลทดลองเทียบรุ่น AI 12 จอ × 3 รุ่น (7 ต.ค. 2026) |
| `data/run*_autoflow_log.*` | log ที่โปรแกรมบันทึกจากการรันจริง |
| `assets/` | ภาพหน้าจอจริง · ภาพออกแบบ · วิดีโอจำลองขั้นตอน |

ไม่มี API key และไม่มีที่อยู่ NAS อยู่ในไฟล์ใดๆ ของ repository นี้
