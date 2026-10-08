# -*- coding: utf-8 -*-
"""
AUTO FLOW · Demo (โมดูลแยกของ InspectPilot)
=============================================
สายพานอัตโนมัติ: ดึงรูปจาก NAS → ตรวจเบลอ/มืด → AI (Gemini) เลือกรูป + แต่งสี →
คนยืนยันเฉพาะจอที่ไม่ชัวร์ → ครอป → รีไซส์ → Stamp → ชื่อไฟล์ → PowerPoint → CSV

- ไฟล์นี้ต้องวางข้าง InspectPilot.py (โฟลเดอร์เดียวกับ campaign_creator.py)
- InspectPilot.py จะ import ไฟล์นี้ตอนเปิดโปรแกรม แล้วเรียก install() เพื่อผูกหน้าต่าง
  AUTO FLOW เข้ากับโปรแกรม และโชว์ปุ่ม "✦ AUTO FLOW · Demo" บนหัวโปรแกรม
- ไม่ใช้แล้ว → ลบไฟล์นี้ทิ้งได้เลย ปุ่มจะหายไปเอง ส่วนอื่นของโปรแกรมไม่กระทบ
- ใช้ตัวช่วยของ InspectPilot (สี ฟอนต์ การ์ด ปุ่ม render_stamp python-pptx ฯลฯ)
  ผ่าน install() จึงเปิดไฟล์นี้เดี่ยวๆ ไม่ได้
"""
import os
import sys
import re
import io
import json
import math
import time
import base64
import shutil
import logging
import datetime
import threading
import traceback
import urllib.request
import urllib.error
import tkinter as tk
from tkinter import filedialog
from tkinter import font as tkfont
from PIL import Image, ImageDraw, ImageFont, ImageTk

_HOST = None
AF_CACHE_FILE = None     # ตั้งค่าใน install() (อยู่ในโฟลเดอร์ข้อมูลของ InspectPilot)


# ตัวแปรของ InspectPilot ที่โหลดทีหลัง (lazy) — ห้ามคัดลอกตอน install ให้ดึงสดผ่าน _af_sync_libs()
_HOST_LAZY = {"Presentation", "Inches", "Pt", "PP_ALIGN", "pd", "_report_libs_loaded", "_site_aliases_cache"}
Presentation = Inches = Pt = PP_ALIGN = None


def _af_sync_libs():
    """โหลด python-pptx ผ่าน InspectPilot แล้วดึงชื่อมาใช้ในไฟล์นี้"""
    global Presentation, Inches, Pt, PP_ALIGN
    _HOST._ensure_report_libs()
    Presentation, Inches, Pt, PP_ALIGN = _HOST.Presentation, _HOST.Inches, _HOST.Pt, _HOST.PP_ALIGN
    if Presentation is None:
        raise RuntimeError("ไม่พบไลบรารี python-pptx")


def install(host_module, app_class):
    """ผูก AUTO FLOW เข้ากับ InspectPilot: ยืมตัวช่วยจากโปรแกรมหลัก แล้วเพิ่มเมธอดให้คลาสแอป"""
    global _HOST, AF_CACHE_FILE
    _HOST = host_module
    g = globals()
    for k, v in vars(host_module).items():
        if not k.startswith("__") and k not in g and k not in _HOST_LAZY:
            g[k] = v
    AF_CACHE_FILE = os.path.join(host_module.DATA_DIR, "autoflow_ai_cache.json")
    for name in ("_open_auto_flow", "_af_worker", "_af_review"):
        setattr(app_class, name, g[name])
    AF_STATIONS[:] = _af_build_stations()
    app_class.AF_ACCENT = AF_ACCENT
    app_class.AF_STATIONS = AF_STATIONS
    app_class.AF_CHIPS = AF_CHIPS


# ======================================================================
# ส่วนที่ 1 · ฟังก์ชันระดับโมดูล (ตรวจรูป · เรียก Gemini · ภาพเคลื่อนไหว)
# ======================================================================
def af_fallback_font():
    """[DEMO #auto-flow] ฟอนต์สำรองเมื่อไม่มีฟอนต์ประทับวันที่ของโปรแกรม"""
    cands = [os.path.join(os.environ.get("WINDIR", "C:\\Windows"), "Fonts", n) for n in ("arialbd.ttf", "tahomabd.ttf", "arial.ttf")]
    cands += ["/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", "/System/Library/Fonts/Supplemental/Arial Bold.ttf"]
    for c in cands:
        if os.path.exists(c):
            return c
    return None


# ============================================================
# [DEMO #auto-flow] AUTO FLOW — สายพานเดียว ดึงรูป → ตรวจเบลอ/มืด → AI เลือก →
# คนยืนยัน → Stamp → Clean → PowerPoint → CSV
# ทำไว้สาธิตส่งประเมิน AI Literacy (Level 4–5) เท่านั้น ไม่ใช่งานจริง
# ============================================================
AF_BLUR_MIN = 60.0        # ความแปรปรวนของ Laplacian (ที่ภาพย่อกว้าง 640px) ต่ำกว่านี้ = เบลอ
AF_DARK_MAX = 50.0        # ความสว่างเฉลี่ย (0–255) ต่ำกว่านี้ = มืดเกินไป
AF_BRIGHT_MIN = 215.0     # สูงกว่านี้ = สว่างจ้าเกินไป
# [CHANGE #af-lite-main] ต.ค. 2026: ทดลองเทียบ 12 จอ (ai_compare) → 3.5 Flash-Lite ตัดสินซีนตรงกติกากว่า
# เร็วกว่า ~16 เท่า (4 วิ vs 69 วิ/จอ) ถูกกว่า ~9 เท่า → ใช้เป็นตัวหลัก · 3.5 Flash เป็นตัวสำรอง
AF_DEFAULT_MODEL = "gemini-3.5-flash-lite"
AF_MODEL_CFG_VER = 2                      # เครื่องที่บันทึก 3.5 Flash ไว้ (ค่าเริ่มต้นเดิม) → เปลี่ยนเป็น Lite ให้ครั้งเดียว
AF_DEFAULT_THRESHOLD = 70
AF_MAX_AI_IMAGES = 8          # 1 จอมีได้ 1–8 รูป (ตามเดือน)
AF_AI_PROMPT = (
    "คุณคือผู้ตรวจรูปถ่ายป้ายโฆษณา (จอ LED/บิลบอร์ด) สำหรับรายงานประจำเดือน\n"
    "มีรูปของจอเดียวกัน {n} รูป เรียงเป็นรูปที่ 1 ถึง {n}\n"
    "ให้เลือกรูปที่ดีที่สุด 1 รูป โดยดูว่า: เห็นหน้าจอ/สินค้าชัดเจน แสงพอดี ไม่เบลอ "
    "ไม่มีสิ่งกีดขวางบังจอ และมุมมองเห็นทั้งจอ\n"
    "ตอบเป็น JSON เท่านั้น รูปแบบ: "
    '{{"best": <เลขรูป 1-{n}>, "confidence": <0-100>, "reason": "<เหตุผลภาษาไทย 1 ประโยคสั้นๆ>", '
    '"color": [<ปุ่มแต่งสี 0-3 ปุ่ม เรียงตามลำดับ>], "color_reason": "<เหตุผลการแต่งสี ภาษาไทยสั้นๆ>"}}\n'
    "ปุ่มแต่งสีที่เลือกได้ (ชุดเดียวกับที่ทีมใช้) มีแค่: \"Auto\" = ปรับสมดุลแสงสีอัตโนมัติ, "
    "\"Balance\" = แก้สีอมเหลือง/อมเขียวเล็กน้อย, \"Brighter\" = สว่างขึ้นเล็กน้อย, \"Darker\" = มืดลงเล็กน้อย "
    "ถ้ารูปดีอยู่แล้วให้ color เป็น []\n"
    "ถ้ารูปทั้งหมดใช้ไม่ได้ ให้ confidence ต่ำกว่า 40"
)
# [FEATURE #af-scene] รุ่นใหม่: AI หากรอบจอ + ตั้งชื่อซีน + ไม่ให้ซีนซ้ำกับจอก่อนหน้า (ไม่ใช้ YOLO)
AF_AI_PROMPT_SCENE = (
    "คุณคือผู้ตรวจรูปถ่ายป้ายโฆษณา (จอ LED/บิลบอร์ด) สำหรับรายงานประจำเดือน\n"
    "งานนี้คือสินค้าเดียวกันหลายจอ แต่ละจอต้องใช้ 'ซีน' ไม่ซ้ำกัน · ตอนนี้เป็นรูปของจอเดียวกัน {n} รูป (รูปที่ 1–{n})\n"
    "{used}"
    "ทำทุกรูป:\n"
    "1) box = กรอบของส่วนจอภาพที่แสดงโฆษณา [ymin, xmin, ymax, xmax] สเกล 0–1000\n"
    "2) scene = ชื่อซีนสั้นๆ ภาษาไทย 3–8 คำ บรรยาย 'สิ่งที่เห็นบนจอ' ในจังหวะนั้น "
    "ห้ามใช้ชื่อสินค้า/แบรนด์/แคมเปญ (ทุกจอเป็นสินค้าเดียวกันอยู่แล้ว) "
    "ให้บอก: คน (จำนวน/ลักษณะ) หรือวัตถุหลัก + สีพื้นหลัง + ข้อความพาดหัวใหญ่ถ้ามี "
    "เช่น \"ชาย 2 คนคู่ขวด พื้นฟ้า\", \"ขวดเดี่ยว ข้อความ 9.9 พื้นแดง\", \"หญิงถือกล่อง พื้นชมพู\" · "
    "ไม่ต้องบอกท่าทาง (ท่าทางไม่ใช่ตัวแบ่งซีน) · "
    "ซีนเดียวกัน = คน/วัตถุหลักชุดเดิม + พื้นหลัง/โทนสีเดิม + ข้อความพาดหัวเดิม "
    "แม้ท่าทาง มุมกล้อง ขนาด หรือจังหวะภาพเคลื่อนไหวต่างกัน · "
    "คนละซีน = เปลี่ยนคน/วัตถุหลัก หรือเปลี่ยนพื้นหลัง/ฉาก หรือเปลี่ยนข้อความพาดหัว · "
    "ถ้าจอแบ่งหลายช่อง ให้วิเคราะห์เองว่าส่วนไหนเปลี่ยนตามจังหวะ แล้วใช้ส่วนนั้นตัดสินซีน ไม่ใช้ส่วนที่ค้างเหมือนกันทุกรูป\n"
    "3) no_product = true ถ้าจอดำ/กำลังเปลี่ยนภาพ/ไม่เห็นสินค้า\n"
    "4) score 0–100 = คุณภาพทั้งภาพ: เห็นจอ/สินค้าชัด แสงพอดี ไม่เบลอ ไม่มีสิ่งกีดขวางบังจอ เห็นทั้งจอ องค์ประกอบดี\n"
    "{dup_rule}"
    "แล้วเลือก best 1 รูป: {pick_rule}\n"
    "ตอบเป็น JSON เท่านั้น รูปแบบ: "
    '{{"best": <เลขรูป 1-{n}>, "confidence": <0-100>, "reason": "<เหตุผลภาษาไทย 1 ประโยคสั้นๆ>", '
    '"color": [<ปุ่มแต่งสี 0-3 ปุ่ม>], "color_reason": "<เหตุผลการแต่งสี สั้นๆ>", '
    '"images": [{{"i": <เลขรูป>, "box": [ymin, xmin, ymax, xmax], "scene": "<ชื่อซีน>", "dup": <เลขซีนที่ใช้แล้ว หรือ 0>, '
    '"no_product": <true/false>, "score": <0-100>}}], "scene_note": "<ถ้าจอแบ่งหลายช่อง บอกสั้นๆ ว่าเทียบส่วนไหน ไม่มีให้เว้นว่าง>"}}\n'
    "ปุ่มแต่งสีที่เลือกได้ มีแค่: \"Auto\" = ปรับสมดุลแสงสีอัตโนมัติ, \"Balance\" = แก้สีอมเหลือง/อมเขียว, "
    "\"Brighter\" = สว่างขึ้นเล็กน้อย, \"Darker\" = มืดลงเล็กน้อย · ถ้ารูปดีอยู่แล้วให้ color เป็น []\n"
    "ถ้ารูปทั้งหมดใช้ไม่ได้ ให้ confidence ต่ำกว่า 40"
)
AF_AI_PROMPT_VER = "scene2"    # [FIX #af-cache-ver] เปลี่ยนคำสั่ง AI เมื่อไร เปลี่ยนเลขนี้ → ไม่ใช้ผลที่จำไว้จากคำสั่งเก่า
AF_SCENE_NAME_MAX = 60         # [FIX #af-scene-len] เดิม 40 ตัวอักษร ชื่อซีนโดนตัดกลางคำ ("…พื้นฟ้")
# [FEATURE #af-fallback] รุ่นหลักยุ่ง (503) / โควต้าหมด → สลับไปรุ่น Lite ให้เอง (ลองตามลำดับ · รุ่นที่ไม่มีจะข้ามเอง)
AF_FALLBACK_MODELS = ("gemini-3.5-flash", "gemini-3.5-flash-lite", "gemini-3.1-flash-lite",
                      "gemini-3.1-flash-lite-preview", "gemini-2.5-flash-lite")   # รุ่นหลักถูกตัดออกจากรายการเอง
AF_SCENE_REF_MAX = 12          # ส่งซีนที่ใช้แล้วไปเทียบได้สูงสุดกี่ซีน (ภาพจอย่อ)
AF_COLOR_CODES = {"auto": "A", "balance": "S", "brighter": "D", "darker": "F"}
AF_COLOR_LABEL = {"A": "✨ Auto", "S": "🎨 Balance", "D": "☀ สว่างขึ้น", "F": "🌙 มืดลง"}


def af_collect_groups(src, screens=None):
    """จัดกลุ่มรูปเป็น (ชื่อจอ, [ไฟล์]) — src เป็นโฟลเดอร์เดียวหรือหลายโฟลเดอร์ก็ได้
    1) โฟลเดอร์ย่อย 1 โฟลเดอร์ = 1 จอ (ข้ามโฟลเดอร์ที่ขึ้นต้น #)
    2) ไฟล์รวมโฟลเดอร์เดียว: ถ้ามีรายชื่อจอ (Master File) → จับจากชื่อไฟล์ที่ขึ้นต้นด้วยชื่อจอ
       (วิธีเดียวกับ SmartPicker) ไม่มี → ตัดเลขท้าย '(1)' แล้วใช้ชื่อที่เหลือ"""
    srcs = list(src) if isinstance(src, (list, tuple)) else [src]
    search = sorted({x.strip().upper(): x.strip() for x in (screens or []) if x and x.strip()}.items(),
                    key=lambda kv: len(kv[0]), reverse=True)
    groups = {}

    def add(key, path):
        groups.setdefault(key, []).append(path)

    for sdir in srcs:
        if not sdir or not os.path.isdir(sdir):
            continue
        entries = sorted(os.scandir(sdir), key=lambda e: natural_name_key(e.name))
        for d in entries:
            if d.is_dir() and not d.name.startswith("#"):
                for f in os.listdir(d.path):
                    if f.lower().endswith(SUPPORTED_FORMATS):
                        add(d.name, os.path.join(d.path, f))
        for e in entries:
            if not (e.is_file() and e.name.lower().endswith(SUPPORTED_FORMATS)):
                continue
            up = e.name.upper()
            key = None
            for m_up, m in search:
                if up.startswith(m_up) and (len(up) == len(m_up) or not up[len(m_up)].isalnum()):
                    key = m
                    break
            if key is None:
                key = os.path.splitext(CLEAN_NUMBER_SUFFIX_RE.sub("", e.name))[0].strip()
            add(key, e.path)
    return [(k, sorted(v, key=lambda x: natural_name_key(os.path.basename(x))))
            for k, v in sorted(groups.items(), key=lambda kv: natural_name_key(kv[0]))]


def af_quality(path):
    """ตรวจคุณภาพรูปในเครื่อง (PIL ล้วน ไม่ต้องใช้ OpenCV)
    คืน dict: sharp (ความคม), bright (ความสว่าง), flags [..], ok"""
    from PIL import ImageStat
    with Image.open(path) as im:
        try:
            im.draft("L", (1280, 1280))
        except Exception:
            pass
        g = im.convert("L")
    if g.width > 640:
        g = g.resize((640, max(1, int(g.height * 640 / g.width))), Image.BILINEAR)
    bright = ImageStat.Stat(g).mean[0]
    lap = g.filter(ImageFilter.Kernel((3, 3), [0, 1, 0, 1, -4, 1, 0, 1, 0], scale=1, offset=128))
    sharp = ImageStat.Stat(lap).var[0]
    flags = []
    if sharp < AF_BLUR_MIN:
        flags.append("เบลอ")
    if bright < AF_DARK_MAX:
        flags.append("มืดเกินไป")
    elif bright > AF_BRIGHT_MIN:
        flags.append("สว่างจ้าเกินไป")
    return {"sharp": sharp, "bright": bright, "flags": flags, "ok": not flags}


def af_local_pick(qs):
    """โหมดไม่ใช้ AI: เลือกด้วยเกณฑ์ในเครื่อง (คมชัด + แสงใกล้กลาง) — qs: [(idx, q)]"""
    def score(q):
        expo = max(0.0, 1.0 - abs(q["bright"] - 128.0) / 128.0)
        return math.log1p(q["sharp"]) * (0.5 + 0.5 * expo)
    ranked = sorted(qs, key=lambda t: score(t[1]), reverse=True)
    best = ranked[0]
    if len(ranked) == 1:
        return best[0], 100, "ผ่านเกณฑ์รูปเดียว"
    s1, s2 = score(ranked[0][1]), score(ranked[1][1])
    q = best[1]
    conf = 55
    conf += 20 if q["sharp"] >= 3 * AF_BLUR_MIN else 0
    conf += 10 if 70 <= q["bright"] <= 190 else 0
    conf += int(min(14, 200 * (s1 - s2) / max(s1, 1e-6)))
    conf = max(0, min(99, conf))
    return best[0], conf, "คมชัดที่สุดและแสงพอดีที่สุด (เกณฑ์ในเครื่อง ไม่ใช้ AI)"


def _af_jpeg_b64(path, max_side=640):
    import base64, io
    with Image.open(path) as im:
        try:
            im.draft("RGB", (max_side * 2, max_side * 2))
        except Exception:
            pass
        im = im.convert("RGB")
        im.thumbnail((max_side, max_side), Image.LANCZOS)
        buf = io.BytesIO()
        im.save(buf, "JPEG", quality=80)
    return base64.b64encode(buf.getvalue()).decode("ascii")


class AFQuotaDayError(RuntimeError):
    """โควต้า AI รายวันหมด (429 แบบ PerDay) — ไม่ต้องลองซ้ำ"""


class AFBusyError(RuntimeError):
    """Google ยุ่ง/ขัดข้อง/ชนโควต้าต่อนาที/หมดเวลา (ลองใหม่ครบแล้ว) — สลับไปรุ่นอื่นได้"""


class AFModelMissing(RuntimeError):
    """ไม่มีรุ่นนี้ (404) — ข้ามไปรุ่นถัดไป"""


def _af_parse_429(body):
    """แยกชนิด 429: (ต่อวัน?, วินาทีที่ให้รอ, ข้อความ)"""
    try:
        err = json.loads(body).get("error", {})
    except Exception:
        return False, 30, ""
    per_day, delay = False, 30
    for d in err.get("details") or []:
        for v in d.get("violations") or []:
            if "perday" in str(v.get("quotaId", "")).lower():
                per_day = True
        rd = d.get("retryDelay")
        if rd:
            m = re.match(r"([\d.]+)s", str(rd))
            if m:
                delay = float(m.group(1))
    return per_day, delay, err.get("message", "")


def _af_gemini_call(api_key, model, parts, timeout=150, meta=None):
    """เรียก Gemini ผ่าน REST (urllib ล้วน ไม่ต้องติดตั้ง SDK) คืนข้อความคำตอบ
    [FIX #af-retry] เดิมเอาข้อความ error ไปทับข้อมูลคำขอ → รอบที่ลองใหม่พังทุกครั้ง ("POST data should be bytes")
    ตอนนี้แยกกัน + รอนานขึ้น (รูปหลายรูป) + หมดเวลา/เน็ตสะดุด ลองใหม่ให้เอง · ข้อความ error บอกสาเหตุชัด"""
    import urllib.request, urllib.error, socket
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
    payload = json.dumps({"contents": [{"role": "user", "parts": parts}],
                          "generationConfig": {"temperature": 0.2, "responseMimeType": "application/json"}}).encode("utf-8")
    last = None
    for attempt in range(3):
        req = urllib.request.Request(url, data=payload, method="POST",
                                     headers={"Content-Type": "application/json", "x-goog-api-key": api_key})
        try:
            with urllib.request.urlopen(req, timeout=timeout) as r:
                data = json.loads(r.read().decode("utf-8"))
            if meta is not None:                  # [FEATURE #ai-compare] เก็บจำนวน token + จำนวนครั้งที่ลอง
                meta["usage"] = data.get("usageMetadata") or {}
                meta["tries"] = attempt + 1
            return data["candidates"][0]["content"]["parts"][0]["text"]
        except urllib.error.HTTPError as e:
            try:
                ebody = e.read().decode("utf-8")
            except Exception:
                ebody = ""
            if e.code == 429:
                per_day, delay, msg = _af_parse_429(ebody)
                if per_day:
                    raise AFQuotaDayError("โควต้า AI วันนี้หมดแล้ว (Free tier ต่อวัน)")
                last = AFBusyError(f"ชนโควต้าต่อนาที (429) {msg[:60]}".strip())
                if attempt < 2:
                    time.sleep(min(65.0, max(5.0, delay + 1)))    # รอตามที่ Google บอก แล้วลองใหม่
                    continue
                raise last
            try:
                msg = json.loads(ebody).get("error", {}).get("message", "")
            except Exception:
                msg = ""
            label = {500: "Google ขัดข้อง (500)", 503: "Google ยุ่ง (503)", 400: "คำขอไม่ถูกต้อง (400)",
                     403: "API Key ไม่มีสิทธิ์ (403)", 404: f"ไม่มีรุ่น {model} (404)"}.get(e.code, f"HTTP {e.code}")
            if e.code == 404:
                raise AFModelMissing(label)
            last = (AFBusyError if e.code in (500, 502, 503, 504) else RuntimeError)(f"{label} {msg[:80]}".strip())
            if e.code in (500, 502, 503, 504) and attempt < 2:
                time.sleep(6 * (attempt + 1))
                continue
            raise last
        except (socket.timeout, TimeoutError) as e:
            last = AFBusyError(f"หมดเวลารอ AI ({timeout} วิ)")
            if attempt < 2:
                continue
            raise last
        except urllib.error.URLError as e:
            reason = getattr(e, "reason", e)
            if isinstance(reason, (socket.timeout, TimeoutError)) or "timed out" in str(reason):
                last = AFBusyError(f"หมดเวลารอ AI ({timeout} วิ)")
            else:
                last = RuntimeError(f"เชื่อมต่อไม่ได้ ({str(reason)[:60]})")
            if attempt < 2:
                time.sleep(3)
                continue
            raise last
        except (KeyError, IndexError):
            raise RuntimeError("AI ตอบกลับไม่ครบ")
    raise last or RuntimeError("AI ไม่ตอบ")


def af_ai_pacing(model):
    """เว้นระยะระหว่างคำขอ (วินาที) ให้ไม่ชนโควต้าต่อนาทีของรุ่นฟรี
    (วัดจริง ก.ย. 2026: Flash 5 คำขอ/นาที · Flash-Lite 15 คำขอ/นาที)"""
    return 4.5 if "lite" in (model or "").lower() else 12.5


def af_gemini_ping(api_key, model):
    txt = _af_gemini_call(api_key, model, [{"text": 'ตอบเป็น JSON {"ok": true}'}], timeout=30)
    return "ok" in txt.lower()


def af_img_b64(pil, max_side=256):
    """ภาพ PIL → base64 JPEG (ใช้ส่งภาพจอย่อของซีนที่ใช้แล้ว)"""
    import base64, io
    im = pil.convert("RGB").copy()
    im.thumbnail((max_side, max_side), Image.LANCZOS)
    buf = io.BytesIO()
    im.save(buf, "JPEG", quality=82)
    return base64.b64encode(buf.getvalue()).decode("ascii")


def _af_box_from_ai(v):
    """[ymin, xmin, ymax, xmax] 0–1000 → (x0, y0, x1, y1) สัดส่วน 0–1 · ผิดรูปแบบ → None"""
    try:
        y0, x0, y1, x1 = [max(0.0, min(1000.0, float(t))) / 1000.0 for t in list(v)[:4]]
    except Exception:
        return None
    if x1 - x0 < 0.02 or y1 - y0 < 0.02:
        return None
    return (x0, y0, x1, y1)


def af_scene_pick(infos, ai_best, counts=None):
    """กฎตัดสิน (บังคับในโปรแกรมอีกชั้น): เลือกจากรูปที่ซีนยังไม่ถูกใช้ + เห็นสินค้า ก่อน
    ไม่มีเลย → ยอมซีนซ้ำ: ซีนที่ใช้ไปน้อยจอที่สุดก่อน แล้วค่อยคะแนนภาพรวม (12 จอ 5 ซีน = ซีนละ 2–3 จอ)
    counts[j] = ใช้ซีนที่ j (เริ่ม 1) ไปแล้วกี่จอ · คืน (index, forced_dup, overridden)"""
    n = len(infos)
    if n == 0:
        return ai_best, False, False
    fresh = [k for k, x in enumerate(infos) if x and not x.get("dup") and not x.get("noprod")]
    if fresh:
        if ai_best in fresh:
            return ai_best, False, False
        return max(fresh, key=lambda k: infos[k].get("score", 0)), False, True
    ok = [k for k, x in enumerate(infos) if x and not x.get("noprod")] or list(range(n))
    counts = counts or {}

    def use_n(k):
        return counts.get((infos[k] or {}).get("dup") or 0, 0)
    least = min(use_n(k) for k in ok)
    ok = [k for k in ok if use_n(k) == least]
    if ai_best in ok:
        return ai_best, True, False
    return max(ok, key=lambda k: (infos[k] or {}).get("score", 0)), True, True


def _af_saved_model(cfg):
    """รุ่น AI ที่บันทึกไว้ · ค่าเริ่มต้นเก่า (3.5 Flash / 2.5 Flash) ที่ยังไม่เคยย้าย → เปลี่ยนเป็นรุ่นหลักใหม่ครั้งเดียว
    (ถ้าผู้ใช้พิมพ์เลือก 3.5 Flash เองหลังจากนี้ จะจำไว้ตามนั้น เพราะ model_v ถูกบันทึกเป็นเวอร์ชันใหม่แล้ว)"""
    m = (cfg.get("model") or "").strip()
    if not m or m == "gemini-2.5-flash":
        return AF_DEFAULT_MODEL
    if m == "gemini-3.5-flash" and int(cfg.get("model_v", 1) or 1) < AF_MODEL_CFG_VER:
        return AF_DEFAULT_MODEL
    return m


_AF_NEW_CLAIM = re.compile(r"ไม่ซ้ำ|ซีนใหม่|ยังไม่(?:เคย)?(?:ถูก)?(?:เลือก)?ใช้")
_AF_DUP_CLAIM = re.compile(r"(?<!ไม่)ซ้ำ")


def af_reason_fix(reason, k, ai_best, info, over):
    """[FIX #af-reason] เหตุผลของ AI บางครั้งขัดกับผลจริง (เช่น เขียน "ไม่ซ้ำ" แต่รูปที่ได้ซีนซ้ำ
    หรือโปรแกรมเปลี่ยนรูปให้ แต่เหตุผลยังพูดถึงรูปที่ AI เสนอ) → เขียนเหตุผลใหม่จากข้อมูลจริง
    คืน (เหตุผล, ถูกแทนไหม) · ส่วน "ซีน: … / ซ้ำกับ …" โปรแกรมต่อท้ายให้ภายหลังจากผลเช็กซีนจริง"""
    reason = (reason or "").strip()
    info = info or {}
    sc = info.get("score")
    sc_txt = f" · คะแนนภาพรวม {sc}/100" if sc else ""
    if over:
        return f"โปรแกรมเลือกรูปที่ {k + 1} แทนรูปที่ {ai_best + 1} ที่ AI เสนอ (ตามกฎกันซีนซ้ำ){sc_txt}", True
    is_dup = bool(info.get("dup"))
    says_new = bool(_AF_NEW_CLAIM.search(reason))
    says_dup = bool(_AF_DUP_CLAIM.search(_AF_NEW_CLAIM.sub("", reason)))
    if (is_dup and says_new) or (not is_dup and says_dup and not says_new and info):
        return f"รูปที่ {k + 1} เห็นจอชัดและภาพรวมดีที่สุดในรูปที่เลือกได้{sc_txt}", True
    return reason or "—", False


def af_gemini_pick(api_key, model, paths, used=None, meta=None):
    """ส่งรูป (ย่อแล้ว) ของ 1 จอ + ภาพจอย่อของซีนที่ใช้แล้ว ให้ AI เลือก
    คืน (index, confidence, reason, acts, creason, infos[ต่อรูป], scene_note)
    infos[k] = {"box": (x0,y0,x1,y1)|None, "scene": str, "dup": int, "noprod": bool, "score": int}"""
    n = len(paths)
    used = list(used or [])[-AF_SCENE_REF_MAX:]
    if used:
        u_txt = ("ซีนที่ใช้แล้ว (จอก่อนหน้าเลือกไปแล้ว ดูภาพจอย่อที่แนบก่อนรูปของจอนี้):\n"
                 + "".join(f"  ซีนที่ใช้แล้ว {j}: {u['name']} (จอ {u['screen']} · ใช้แล้ว {u.get('count', 1)} จอ)\n"
                           for j, u in enumerate(used, 1)))
        dup_rule = ("5) dup = เลขซีนที่ใช้แล้วที่เป็นซีนเดียวกับรูปนี้ (เทียบกับภาพจอย่อที่แนบ ไม่ใช่เทียบแค่ชื่อ · "
                    "คนชุดเดิม + พื้นหลังเดิม = ซ้ำ แม้ท่าทางต่างกัน) · ไม่ซ้ำกับข้อใด dup = 0\n")
        pick_rule = ("เลือกจากรูปที่ dup = 0 และ no_product = false ที่คะแนนภาพรวมสูงสุด · "
                     "ถ้าไม่มีรูปแบบนั้นเลย (ต้องซ้ำ) ให้เลือกซีนที่ 'ใช้แล้วน้อยจอที่สุด' ก่อน แล้วค่อยดูคะแนนภาพรวม")
    else:
        u_txt, dup_rule = "นี่คือจอแรก ยังไม่มีซีนที่ใช้แล้ว\n", "5) dup = 0 ทุกรูป\n"
        pick_rule = "รูปที่ no_product = false และคะแนนภาพรวมสูงสุด"
    parts = [{"text": AF_AI_PROMPT_SCENE.format(n=n, used=u_txt, dup_rule=dup_rule, pick_rule=pick_rule)}]
    for j, u in enumerate(used, 1):
        if u.get("b64"):
            parts.append({"text": f"ซีนที่ใช้แล้ว {j}: {u['name']}"})
            parts.append({"inline_data": {"mime_type": "image/jpeg", "data": u["b64"]}})
    parts.append({"text": "— รูปของจอนี้ —"})
    for i, p in enumerate(paths, 1):
        parts.append({"text": f"รูปที่ {i}"})
        parts.append({"inline_data": {"mime_type": "image/jpeg", "data": _af_jpeg_b64(p)}})
    txt = _af_gemini_call(api_key, model, parts, meta=meta)
    m = re.search(r"\{.*\}", txt, re.S)
    if not m:
        raise RuntimeError("AI ตอบไม่เป็น JSON")
    obj = json.loads(m.group(0))
    best = int(obj.get("best", 0)) - 1
    if not 0 <= best < n:
        raise RuntimeError(f"AI ตอบเลขรูปไม่ถูกต้อง ({obj.get('best')})")
    conf = int(max(0, min(100, float(obj.get("confidence", 0)))))
    reason = str(obj.get("reason", "")).strip() or "—"
    acts = []
    for a in (obj.get("color") or [])[:3]:
        code = AF_COLOR_CODES.get(str(a).strip().lower())
        if code:
            acts.append(code)
    creason = str(obj.get("color_reason", "")).strip()
    infos = [None] * n
    for it in (obj.get("images") or []):
        try:
            k = int(it.get("i", 0)) - 1
        except Exception:
            continue
        if not 0 <= k < n:
            continue
        try:
            dup = int(it.get("dup") or 0)
        except Exception:
            dup = 0
        if not 0 <= dup <= len(used):
            dup = 0
        try:
            score = int(max(0, min(100, float(it.get("score", 0)))))
        except Exception:
            score = 0
        infos[k] = {"box": _af_box_from_ai(it.get("box") or []), "scene": str(it.get("scene", "")).strip()[:AF_SCENE_NAME_MAX],
                    "dup": dup, "noprod": bool(it.get("no_product")), "score": score}
        if dup:                                        # ใช้ชื่อเดิมของซีนที่ใช้แล้ว (ชื่อตรงกันทุกจอ)
            infos[k]["scene"] = used[dup - 1]["name"]
    note = str(obj.get("scene_note", "")).strip()[:120]
    return best, conf, reason, acts, creason, infos, note


def af_generate_pptx(template, image_paths, start_page, out_path, order="row", quality="standard",
                     on_progress=None, start_auto=True):
    """[DEMO #auto-flow] ขั้น 8 — ใช้ตัวสร้างเดียวกับโหมด Generate:
    วิเคราะห์หน้า → ย่อรูป Standard/HD → ก๊อปหน้าต้นแบบเองจนครบ (แทรกก่อนหน้าปิด)
    ไม่มีช่องรูป → 2 รูป/หน้า · คืน dict: placed, leftover, pages {path เดิม: หน้า}, how, added, size"""
    import tempfile
    _af_sync_libs()
    ana = pp_analyze_template(template, None if start_auto else start_page, order)
    total = len(image_paths)
    tmp = tempfile.mkdtemp(prefix="af_pptx_")
    try:
        prepared, budget, _shrunk = pp_prepare_images(image_paths, quality, ana["size"], tmp)
        back = dict(zip(prepared, image_paths))
        items = [(p_, os.path.splitext(os.path.basename(o_))[0]) for p_, o_ in zip(prepared, image_paths)]
        prs = Presentation(template)
        cb = (lambda i, n, pno: on_progress(i, n)) if on_progress else None
        if ana["end"]:
            r = pp_fill_report(prs, ana, items, order, cb)
            how = "ตามช่องในเทมเพลต"
        else:
            r = pp_fill_legacy(prs, items, start_page, "client", cb)
            how = "2 รูป/หน้า (ไม่พบช่องรูป)"
        prs.save(out_path)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    try:
        size = os.path.getsize(out_path)
    except OSError:
        size = 0
    return {"placed": r["placed"], "leftover": total - r["placed"],
            "pages": {back.get(k_, k_): v_ for k_, v_ in r["page_of"].items()},
            "how": how, "added": r["pages_added"], "size": size, "budget": budget,
            "n_slides": len(prs.slides), "start": ana["start"], "end": ana["end"]}


def af_ppt_plan(ana, n, start_page=1):
    """แผนหน้าใน PowerPoint สำหรับ n รูป → [(หน้า, ช่อง, ก๊อปเพิ่มไหม)]"""
    if ana and ana.get("end"):
        plan, _extra = pp_plan_pages(ana, n)
        return plan
    s0 = max(1, int(start_page or 1))
    n_sl = (ana or {}).get("n_slides") or s0
    return [(s0 + k, 2, s0 + k > n_sl) for k in range(max(1, -(-n // 2)))]


def af_local_color(path):
    """แต่งสีแบบไม่ใช้ AI — เลือกปุ่มของทีมจากค่าเฉลี่ยสี/ความสว่าง (ใช้ตอนปิด AI หรือคนเลือกเอง)"""
    from PIL import ImageStat
    with Image.open(path) as im:
        try:
            im.draft("RGB", (640, 640))
        except Exception:
            pass
        sm = im.convert("RGB")
    sm.thumbnail((320, 320))
    r, g, b = ImageStat.Stat(sm).mean[:3]
    y = 0.299 * r + 0.587 * g + 0.114 * b
    acts, why = ["A"], ["ปรับสมดุลแสงสีอัตโนมัติ"]
    if (r - b) > 18 and (g - b) > 10:
        acts.append("S")
        why.append("ภาพอมเหลือง")
    if y < 95:
        acts.append("D")
        why.append("แสงน้อย")
    elif y > 185:
        acts.append("F")
        why.append("สว่างจ้า")
    return acts, " · ".join(why) + " (เกณฑ์ในเครื่อง)"


def af_apply_color(img, acts):
    """ใช้ปุ่มแต่งสีชุดเดียวกับ SmartPicker (A/S/D/F) — ไม่มีโมดูล Fetch → คืนรูปเดิม"""
    if not acts or smart_fetch_module is None:
        return img
    try:
        return smart_fetch_module.apply_effect_stack(img.convert("RGB"), list(acts))
    except Exception:
        return img


def af_crop_box(w, h):
    """กล่องครอปกลางให้ได้สัดส่วนเท่า TARGET (16:9) — คืน (l, t, r, b) และ % ที่ตัด"""
    tr = TARGET_WIDTH / TARGET_HEIGHT
    if w / h > tr:
        nw = int(round(h * tr))
        l = (w - nw) // 2
        return (l, 0, l + nw, h), ("ซ้าย/ขวา", 100.0 * (w - nw) / w / 2)
    nh = int(round(w / tr))
    t = (h - nh) // 2
    return (0, t, w, t + nh), ("บน/ล่าง", 100.0 * (h - nh) / h / 2)


def af_crop_box_focus(w, h, sbox):
    """[FEATURE #af-scene] กล่อง 16:9 ที่เลื่อนให้จออยู่กลางภาพที่สุด (ภายในขอบภาพ)
    sbox = (x0,y0,x1,y1) สัดส่วนของจอ · คืน (box, ข้อความ, box กลางภาพเดิม)"""
    box0, _info = af_crop_box(w, h)
    if not sbox:
        return box0, None, box0
    tr = TARGET_WIDTH / TARGET_HEIGHT
    x0, y0, x1, y1 = sbox
    if w / h > tr:
        nw = box0[2] - box0[0]
        cx = (x0 + x1) / 2 * w
        l = int(round(min(max(cx - nw / 2, 0), w - nw)))
        box = (l, 0, l + nw, h)
        a_, b_ = 100.0 * l / w, 100.0 * (w - l - nw) / w
        txt = f"ตัดซ้าย {a_:.0f}% · ขวา {b_:.0f}%"
    else:
        nh = box0[3] - box0[1]
        cy = (y0 + y1) / 2 * h
        t = int(round(min(max(cy - nh / 2, 0), h - nh)))
        box = (0, t, w, t + nh)
        a_, b_ = 100.0 * t / h, 100.0 * (h - t - nh) / h
        txt = f"ตัดบน {a_:.0f}% · ล่าง {b_:.0f}%"
    return box, txt, box0


def af_quota_date():
    """วันของโควต้า (Google รีเซ็ตโควต้ารายวันตามเวลาแปซิฟิก — ประมาณ UTC-8)"""
    return (datetime.datetime.utcnow() - datetime.timedelta(hours=8)).strftime("%Y-%m-%d")


def af_quota_limit(model):
    return 500 if "lite" in (model or "").lower() else 20




def af_cache_key(model, paths):
    parts = []
    for p in paths:
        try:
            parts.append(f"{os.path.basename(p)}:{os.path.getsize(p)}")
        except OSError:
            parts.append(os.path.basename(p))
    return AF_AI_PROMPT_VER + "|" + (model or "") + "|" + ";".join(parts)


def af_cache_load():
    try:
        with open(AF_CACHE_FILE, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def af_cache_save(cache):
    try:
        os.makedirs(DATA_DIR, exist_ok=True)
        if len(cache) > 3000:
            for k in list(cache)[:len(cache) - 3000]:
                cache.pop(k, None)
        with open(AF_CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump(cache, f, ensure_ascii=False)
    except Exception:
        pass


# ---- ภาพสำหรับ AI Stage (PIL ล้วน ไม่แตะ Tk — สร้างในเธรดงานได้) ----
AF_STAGE_W, AF_STAGE_H = 400, 225


def af_open_small(path, max_side=900):
    with Image.open(path) as im:
        try:
            im.draft("RGB", (max_side * 2, max_side * 2))
        except Exception:
            pass
        im = im.convert("RGB")
    im.thumbnail((max_side, max_side), Image.BILINEAR)
    return im


def af_cover(img, w, h):
    s = max(w / img.width, h / img.height)
    nw, nh = max(w, int(img.width * s + 0.5)), max(h, int(img.height * s + 0.5))
    im = img.resize((nw, nh), Image.BILINEAR)
    l, t = (nw - w) // 2, (nh - h) // 2
    return im.crop((l, t, l + w, t + h))


def af_contain(img, w, h, bg=(12, 15, 24)):
    s = min(w / img.width, h / img.height)
    im = img.resize((max(1, int(img.width * s)), max(1, int(img.height * s))), Image.BILINEAR)
    out = Image.new("RGB", (w, h), bg)
    out.paste(im, ((w - im.width) // 2, (h - im.height) // 2))
    return out


def _af_rgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def _af_badge(d, x, y, kind, r=9):
    col = {"ok": "#34D399", "bad": "#FF8C42", "pick": "#FFB800", "scan": "#22D3EE"}.get(kind, "#8FA3BF")
    d.ellipse([x - r, y - r, x + r, y + r], fill=col)
    ink = (20, 14, 0)
    if kind == "ok":
        d.line([(x - 4, y), (x - 1, y + 4), (x + 5, y - 4)], fill=ink, width=2)
    elif kind == "bad":
        d.line([(x, y - 5), (x, y + 1)], fill=ink, width=2)
        d.ellipse([x - 1.3, y + 3, x + 1.3, y + 5.6], fill=ink)
    elif kind == "pick":
        pts = []
        for i in range(10):
            ang = math.pi / 5 * i - math.pi / 2
            rr = 6 if i % 2 == 0 else 2.6
            pts.append((x + rr * math.cos(ang), y + rr * math.sin(ang)))
        d.polygon(pts, fill=ink)


def af_grid_frame(tiles, states, sweep=None, shimmer=None, W=None, H=None):
    """ขั้น ② — รูปทั้งหมดของจอ: states[i] = ok/bad/scan/pick/dim/wait
    sweep=(i, p) เส้นสแกน · shimmer=x (0..1) แสงวิ่งผ่านรูปที่รอ AI"""
    W, H = W or AF_STAGE_W, H or AF_STAGE_H
    bg = Image.new("RGB", (W, H), (12, 15, 24))
    n = max(1, len(tiles))
    cols = 3 if n > 4 else (2 if n > 1 else 1)
    rows = (n + cols - 1) // cols
    gap = 8
    tw = min(int((W - gap * (cols + 1)) / cols), int(((H - gap * (rows + 1)) / rows) * 16 / 9))
    th = int(tw * 9 / 16)
    ox = (W - (cols * tw + (cols - 1) * gap)) // 2
    oy = (H - (rows * th + (rows - 1) * gap)) // 2
    d = ImageDraw.Draw(bg)
    rects = []
    for i, t in enumerate(tiles):
        r_, c_ = divmod(i, cols)
        x, y = ox + c_ * (tw + gap), oy + r_ * (th + gap)
        im = af_cover(t, tw, th)
        stt = states[i] if i < len(states) else "wait"
        if stt in ("bad", "dim"):
            im = Image.blend(im, Image.new("RGB", im.size, (12, 15, 24)), 0.58)
        bg.paste(im, (x, y))
        rects.append((x, y, x + tw, y + th))
        col = {"pick": "#FFB800", "scan": "#22D3EE", "bad": "#FF8C42"}.get(stt)
        if col:
            d.rectangle([x - 2, y - 2, x + tw + 1, y + th + 1], outline=_af_rgb(col), width=3 if stt == "pick" else 2)
        if stt in ("ok", "bad", "pick"):
            _af_badge(d, x + 12, y + 12, stt)
    if sweep is not None and sweep[0] < len(rects):
        x0, y0, x1, y1 = rects[sweep[0]]
        sx = int(x0 + (x1 - x0) * max(0.0, min(1.0, sweep[1])))
        ov = Image.new("RGBA", bg.size, (0, 0, 0, 0))
        od = ImageDraw.Draw(ov)
        od.rectangle([x0, y0, sx, y1], fill=(34, 211, 238, 40))
        od.rectangle([sx - 1, y0, sx + 1, y1], fill=(170, 245, 255, 255))
        bg = Image.alpha_composite(bg.convert("RGBA"), ov).convert("RGB")
    if shimmer is not None:
        ov = Image.new("RGBA", bg.size, (0, 0, 0, 0))
        od = ImageDraw.Draw(ov)
        for i, (x0, y0, x1, y1) in enumerate(rects):
            if (states[i] if i < len(states) else "") != "ok":
                continue
            cx = x0 - 40 + (x1 - x0 + 80) * ((shimmer + i * 0.12) % 1.0)
            for k in range(24):
                a = int(70 * (1 - abs(k - 12) / 12))
                xx = int(cx - 12 + k)
                if x0 <= xx <= x1:
                    od.line([(xx, y0), (xx - 10, y1)], fill=(255, 211, 77, a), width=1)
        bg = Image.alpha_composite(bg.convert("RGBA"), ov).convert("RGB")
    return bg


def af_wipe_frame(before, after, p, W=None, H=None):
    """ขั้น ③ — BEFORE | AFTER ; เส้นแบ่งเลื่อนจากซ้ายไปกลาง (p 0..1)"""
    W, H = W or AF_STAGE_W, H or AF_STAGE_H
    b = af_contain(before, W, H)
    a = af_contain(after, W, H)
    x = int(W * 0.5 * max(0.0, min(1.0, p))) if p < 1 else W // 2
    out = a.copy()
    out.paste(b.crop((0, 0, x, H)), (0, 0)) if x > 0 else None
    d = ImageDraw.Draw(out)
    if x > 0:
        d.rectangle([x - 1, 0, x + 1, H], fill=(255, 255, 255))
        d.ellipse([x - 11, H // 2 - 11, x + 11, H // 2 + 11], fill=(255, 255, 255))
        d.polygon([(x - 6, H // 2), (x - 2, H // 2 - 4), (x - 2, H // 2 + 4)], fill=(20, 20, 20))
        d.polygon([(x + 6, H // 2), (x + 2, H // 2 - 4), (x + 2, H // 2 + 4)], fill=(20, 20, 20))
    return out


def af_crop_frame(img, box, W=None, H=None):
    """ขั้น ④ — พื้นที่นอกกรอบครอปมืดลง + กรอบเส้นประสีทอง"""
    W, H = W or AF_STAGE_W, H or AF_STAGE_H
    base = af_contain(img, W, H)
    s = min(W / img.width, H / img.height)
    ox, oy = (W - img.width * s) / 2, (H - img.height * s) / 2
    l, t, r, b = [ox + box[0] * s, oy + box[1] * s, ox + box[2] * s, oy + box[3] * s]
    ov = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    od = ImageDraw.Draw(ov)
    for rr in ((0, 0, W, t), (0, b, W, H), (0, t, l, b), (r, t, W, b)):
        if rr[2] > rr[0] and rr[3] > rr[1]:
            od.rectangle(rr, fill=(0, 0, 0, 140))
    out = Image.alpha_composite(base.convert("RGBA"), ov).convert("RGB")
    d = ImageDraw.Draw(out)
    gold = (255, 184, 0)
    for x in range(int(l), int(r), 10):
        d.line([(x, t), (min(x + 5, r), t)], fill=gold, width=2)
        d.line([(x, b), (min(x + 5, r), b)], fill=gold, width=2)
    for y in range(int(t), int(b), 10):
        d.line([(l, y), (l, min(y + 5, b))], fill=gold, width=2)
        d.line([(r, y), (r, min(y + 5, b))], fill=gold, width=2)
    k = 12
    for (cx, cy, sx, sy) in ((l, t, 1, 1), (r, t, -1, 1), (l, b, 1, -1), (r, b, -1, -1)):
        d.line([(cx, cy), (cx + sx * k, cy)], fill=gold, width=4)
        d.line([(cx, cy), (cx, cy + sy * k)], fill=gold, width=4)
    return out


def af_label_frame(img, text, font, W=None, H=None, pos="br", col=(0, 0, 0)):
    """รูป + ป้ายข้อความมุมภาพ (ขนาดไฟล์ / ชื่อไฟล์ / หน้า PPT)"""
    W, H = W or AF_STAGE_W, H or AF_STAGE_H
    out = af_contain(img, W, H) if img.size != (W, H) else img.copy()
    d = ImageDraw.Draw(out)
    try:
        tw = d.textlength(text, font=font)
    except Exception:
        tw = len(text) * 7
    ph = 24
    x = W - tw - 26 if pos == "br" else (W - tw) / 2 - 10
    y = H - ph - 10 if pos in ("br", "bc") else (H - ph) / 2
    d.rounded_rectangle([x, y, x + tw + 20, y + ph], radius=12, fill=col)
    d.text((x + 10, y + 4), text, font=font, fill=(255, 255, 255))
    return out


def af_slide_mini(filled, slots, w=52, h=30, hl=None, cols=None):
    """สไลด์ย่อ: ช่องที่วางแล้วเป็นสีฟ้า ช่องว่างเทา ช่องล่าสุดขอบทอง"""
    im = Image.new("RGB", (w, h), (255, 255, 255))
    d = ImageDraw.Draw(im)
    n = max(1, slots)
    cols = cols or (2 if n > 1 else 1)
    rows = (n + cols - 1) // cols
    pad, gap = 3, 2
    cw = (w - pad * 2 - gap * (cols - 1)) / cols
    ch = (h - pad * 2 - gap * (rows - 1)) / rows
    for i in range(n):
        r_, c_ = divmod(i, cols)
        x0, y0 = pad + c_ * (cw + gap), pad + r_ * (ch + gap)
        d.rectangle([x0, y0, x0 + cw, y0 + ch], fill=(111, 168, 220) if i < filled else (214, 219, 228),
                    outline=(255, 184, 0) if hl == i else None, width=2 if hl == i else 1)
    return im



# ---- [DEMO #auto-flow v3] เฟรมภาพเคลื่อนไหวแบบใหม่ (มุมมน · เรืองแสง · easing) ----
AF_STAGE_BG = (18, 23, 38)
_AF_MASKS = {}


def af_ease(p):
    p = max(0.0, min(1.0, p))
    return p * p * (3 - 2 * p)


def af_ease_out(p):
    p = max(0.0, min(1.0, p))
    return 1 - (1 - p) ** 3


def _af_mask(w, h, r):
    k = (w, h, r)
    m = _AF_MASKS.get(k)
    if m is None:
        S = 3
        big = Image.new("L", (w * S, h * S), 0)
        ImageDraw.Draw(big).rounded_rectangle([0, 0, w * S - 1, h * S - 1], radius=r * S, fill=255)
        m = big.resize((w, h), Image.LANCZOS)
        _AF_MASKS[k] = m
    return m


def af_rounded(img, r=14, bg=AF_STAGE_BG, border=None, bw=2):
    """วางรูปบนพื้น bg แบบมุมมน (+ขอบสี) — ใช้กับทุกภาพที่โชว์บน AI Stage"""
    w, h = img.size
    out = Image.new("RGB", (w, h), bg)
    out.paste(img.convert("RGB"), (0, 0), _af_mask(w, h, r))
    if border:
        S = 3
        ov = Image.new("RGBA", (w * S, h * S), (0, 0, 0, 0))
        ImageDraw.Draw(ov).rounded_rectangle([bw * S // 2, bw * S // 2, w * S - 1 - bw * S // 2, h * S - 1 - bw * S // 2],
                                             radius=r * S, outline=border, width=bw * S)
        out = Image.alpha_composite(out.convert("RGBA"), ov.resize((w, h), Image.LANCZOS)).convert("RGB")
    return out


def _af_tile_layout(n, W, H, gap=10):
    cols = 3 if n > 4 else (2 if n > 1 else 1)
    rows = (n + cols - 1) // cols
    tw = min(int((W - gap * (cols + 1)) / cols), int(((H - gap * (rows + 1)) / rows) * 16 / 9))
    th = int(tw * 9 / 16)
    ox = (W - (cols * tw + (cols - 1) * gap)) // 2
    oy = (H - (rows * th + (rows - 1) * gap)) // 2
    return [(ox + (i % cols) * (tw + gap), oy + (i // cols) * (th + gap), tw, th) for i in range(n)]


def _af_badge2(d, x, y, kind, r=11):
    col = {"ok": (52, 211, 153), "bad": (255, 140, 66), "pick": (255, 184, 0)}.get(kind, (143, 163, 191))
    d.ellipse([x - r - 3, y - r - 3, x + r + 3, y + r + 3], fill=(0, 0, 0, 110))
    d.ellipse([x - r, y - r, x + r, y + r], fill=col + (255,))
    ink = (18, 14, 4, 255)
    if kind == "ok":
        d.line([(x - 5, y), (x - 1, y + 4), (x + 6, y - 5)], fill=ink, width=3)
    elif kind == "bad":
        d.line([(x, y - 6), (x, y + 1)], fill=ink, width=3)
        d.ellipse([x - 1.8, y + 3.5, x + 1.8, y + 7], fill=ink)
    elif kind == "pick":
        pts = []
        for i in range(10):
            ang = math.pi / 5 * i - math.pi / 2
            rr = 7.5 if i % 2 == 0 else 3.2
            pts.append((x + rr * math.cos(ang), y + rr * math.sin(ang)))
        d.polygon(pts, fill=ink)


def af_grid2(tiles, states, W, H, sweep=None, shimmer=None, pulse=0.0, bg=AF_STAGE_BG):
    """ขั้น ② — รูปทั้งหมดของจอ (มุมมน) · sweep=(i,p) · shimmer=t · pulse 0..1 (เรืองแสงรูปที่เลือก)"""
    out = Image.new("RGBA", (W, H), bg + (255,))
    lay = _af_tile_layout(len(tiles), W, H)
    ov = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    od = ImageDraw.Draw(ov)
    for i, (t, (x, y, tw, th)) in enumerate(zip(tiles, lay)):
        stt = states[i] if i < len(states) else "wait"
        im = af_cover(t, tw, th)
        if stt in ("bad", "dim", "dup"):
            im = Image.blend(im, Image.new("RGB", im.size, bg), 0.62 if stt == "bad" else 0.5)
        if stt == "wait":
            im = Image.blend(im, Image.new("RGB", im.size, bg), 0.25)
        out.paste(im, (x, y), _af_mask(tw, th, 10))
        if stt == "pick":
            a = int(90 + 120 * pulse)
            for k in range(6, 0, -1):
                od.rounded_rectangle([x - k * 2, y - k * 2, x + tw + k * 2, y + th + k * 2], radius=12 + k * 2,
                                     outline=(255, 184, 0, max(0, a // (k + 1))), width=2)
            od.rounded_rectangle([x - 1, y - 1, x + tw, y + th], radius=11, outline=(255, 196, 40, 255), width=3)
        elif stt == "scan":
            od.rounded_rectangle([x - 1, y - 1, x + tw, y + th], radius=11, outline=(34, 211, 238, 255), width=2)
        elif stt == "bad":
            od.rounded_rectangle([x - 1, y - 1, x + tw, y + th], radius=11, outline=(255, 140, 66, 160), width=2)
        if stt in ("ok", "bad", "pick"):
            _af_badge2(od, x + 16, y + 16, stt)
        if sweep is not None and sweep[0] == i:
            sx = int(x + tw * af_ease(sweep[1]))
            od.rectangle([x, y, sx, y + th], fill=(34, 211, 238, 34))
            for g in range(10, 0, -1):
                od.rectangle([sx - g, y, sx + 1, y + th], fill=(120, 235, 255, int(12 * (11 - g))))
            od.rectangle([sx - 1, y, sx + 1, y + th], fill=(220, 252, 255, 255))
        if shimmer is not None and stt == "ok":
            ph = ((shimmer * 0.9 + i * 0.17) % 1.0)
            cx = x - 60 + (tw + 120) * ph
            for k in range(40):
                xx = int(cx - 20 + k)
                if x <= xx < x + tw:
                    aa = int(85 * (1 - abs(k - 20) / 20))
                    od.line([(xx, y), (xx - 18, y + th)], fill=(255, 216, 102, aa), width=1)
    out = Image.alpha_composite(out, ov)
    return out.convert("RGB"), lay


AF_GEM_COLS = [(66, 133, 244), (155, 114, 203), (234, 67, 53), (251, 188, 5), (52, 168, 83), (66, 133, 244)]


def af_gem_col(u):
    """สีไล่แบบ Gemini (ฟ้า–ม่วง–แดง–เหลือง–เขียว) · u วนรอบ 0..1"""
    u = u % 1.0
    n = len(AF_GEM_COLS) - 1
    k = int(u * n)
    f_ = u * n - k
    a, b = AF_GEM_COLS[k], AF_GEM_COLS[k + 1]
    return tuple(int(a[j] + (b[j] - a[j]) * f_) for j in range(3))


def af_think_aura(base, lay, states, t, W, H):
    """[FEATURE #af-aura] ระหว่างรอ AI — ขอบสี Gemini วิ่งรอบรูปที่ผ่านเกณฑ์ (เหลื่อมจังหวะกัน) + เรืองแสงนุ่ม
    base = ภาพกริดที่วาดไว้แล้ว (af_grid2) · วาดแค่ขอบทับ ทุกเฟรมจึงเบา"""
    ov = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    d = ImageDraw.Draw(ov)
    steps = 72
    for i, (x, y, tw, th) in enumerate(lay):
        if (states[i] if i < len(states) else "wait") != "ok":
            continue
        x0, y0, x1, y1 = x - 1, y - 1, x + tw, y + th
        pw, ph_ = x1 - x0, y1 - y0
        per = 2 * (pw + ph_)
        phs = t * 0.45 + i * 0.13

        def pt(dd):
            dd %= per
            if dd < pw:
                return (x0 + dd, y0)
            dd -= pw
            if dd < ph_:
                return (x1, y0 + dd)
            dd -= ph_
            if dd < pw:
                return (x1 - dd, y1)
            return (x0, y1 - (dd - pw))
        for s_ in range(steps):
            u = s_ / steps
            dist = u * per
            a = int(235 * (0.30 + 0.70 * (0.5 + 0.5 * math.sin((u + phs) * math.tau))))
            d.line([pt(dist), pt(dist + per / steps + 1)], fill=af_gem_col(u + phs) + (a,), width=3)
    sm = ov.resize((max(1, W // 2), max(1, H // 2)), Image.BILINEAR).filter(ImageFilter.GaussianBlur(3))
    glow = sm.resize((W, H), Image.BILINEAR)
    out = Image.alpha_composite(base.convert("RGBA"), glow)
    return Image.alpha_composite(out, ov).convert("RGB")


def _af_cover_map(iw, ih, tw, th):
    """แปลงพิกัดสัดส่วนในรูปเดิม → พิกัดในไทล์ที่ครอปแบบ af_cover"""
    s = max(tw / iw, th / ih)
    nw, nh = max(tw, int(iw * s + 0.5)), max(th, int(ih * s + 0.5))
    l, t = (nw - tw) // 2, (nh - th) // 2
    return lambda fx_, fy_: (fx_ * nw - l, fy_ * nh - t)


def af_grid_scenes(tiles, states, marks, W, H, alpha=1.0, font=None, bg=AF_STAGE_BG):
    """[FEATURE #af-scene] ขั้น ② — กรอบแดงรอบจอที่ AI หาเจอ + ป้ายชื่อซีน (ซ้ำ = จาง + ป้ายส้ม)
    marks[i] = {"box": (x0,y0,x1,y1)|None, "label": str, "kind": "ok"|"dup"|"bad"|"pick"} · alpha 0..1 (ค่อยๆ ขึ้น/หาย)"""
    img, lay = af_grid2(tiles, states, W, H, bg=bg)
    if alpha <= 0.01:
        return img
    ov = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    od = ImageDraw.Draw(ov)
    a = int(255 * max(0.0, min(1.0, alpha)))
    for i, (t, (x, y, tw, th)) in enumerate(zip(tiles, lay)):
        mk = marks[i] if i < len(marks) else None
        if not mk:
            continue
        kind = mk.get("kind", "ok")
        box = mk.get("box")
        if box:
            mp = _af_cover_map(t.width, t.height, tw, th)
            bx0, by0 = mp(box[0], box[1])
            bx1, by1 = mp(box[2], box[3])
            bx0, by0 = max(0, bx0) + x, max(0, by0) + y
            bx1, by1 = min(tw, bx1) + x, min(th, by1) + y
            if bx1 - bx0 > 4 and by1 - by0 > 4:
                col = (255, 90, 103) if kind != "dup" else (255, 140, 66)
                od.rounded_rectangle([bx0, by0, bx1, by1], radius=4, outline=col + (int(a * (0.55 if kind in ("dup", "bad") else 1)),), width=3)
        lab = mk.get("label")
        if lab and font is not None:
            col = {"dup": (255, 140, 66), "bad": (255, 90, 103), "pick": (255, 184, 0)}.get(kind, (10, 14, 26))
            fg = (255, 255, 255) if kind == "ok" else (18, 14, 4)
            tw_ = int(od.textlength(lab, font=font)) + 16
            hh = int(font.size * 1.55)
            lx, ly = x + 8, y + th - hh - 8
            tw_ = min(tw_, tw - 16)
            od.rounded_rectangle([lx, ly, lx + tw_, ly + hh], radius=hh // 2, fill=(col + (int(a * 0.9),)) if kind != "ok" else (10, 14, 26, int(a * 0.82)))
            od.text((lx + 8, ly + hh / 2), lab, font=font, fill=fg + (a,), anchor="lm")
    return Image.alpha_composite(img.convert("RGBA"), ov).convert("RGB")


def af_zoom2(grid_img, rect, full, p, W, H, bg=AF_STAGE_BG):
    """รูปที่เลือกขยายจากตำแหน่งในกริดจนเต็มเวที"""
    e = af_ease_out(p)
    x, y, tw, th = rect
    fx, fy, fw, fh = 0, 0, W, H
    cx, cy = x + (fx - x) * e, y + (fy - y) * e
    cw, ch = tw + (fw - tw) * e, th + (fh - th) * e
    base = Image.blend(grid_img, Image.new("RGB", grid_img.size, bg), min(0.85, 0.85 * e))
    im = af_cover(full, max(2, int(cw)), max(2, int(ch)))
    r = int(10 + 4 * e)
    base.paste(im, (int(cx), int(cy)), _af_mask(im.width, im.height, r))
    ov = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    ImageDraw.Draw(ov).rounded_rectangle([cx, cy, cx + cw - 1, cy + ch - 1], radius=r, outline=(255, 196, 40, 255), width=3)
    return Image.alpha_composite(base.convert("RGBA"), ov).convert("RGB")


def af_wipe2(b, a, p, W, H):
    """BEFORE → AFTER เส้นแบ่งกวาดจากซ้ายไปขวาแล้วกลับมาหยุดกลางภาพ"""
    b = af_cover(b, W, H)
    a = af_cover(a, W, H)
    if p < 0.7:
        x = int(W * af_ease(p / 0.7))
    else:
        x = int(W - W * 0.5 * af_ease((p - 0.7) / 0.3))
    out = a.copy()
    if x > 0:
        out.paste(b.crop((0, 0, x, H)), (0, 0))
    ov = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    d = ImageDraw.Draw(ov)
    for g in range(14, 0, -1):
        d.rectangle([x - g, 0, x + g, H], fill=(255, 255, 255, int(5 * (15 - g) / 2)))
    d.rectangle([x - 1, 0, x + 1, H], fill=(255, 255, 255, 255))
    d.ellipse([x - 15, H // 2 - 15, x + 15, H // 2 + 15], fill=(255, 255, 255, 255))
    d.polygon([(x - 8, H // 2), (x - 3, H // 2 - 5), (x - 3, H // 2 + 5)], fill=(20, 20, 30, 255))
    d.polygon([(x + 8, H // 2), (x + 3, H // 2 - 5), (x + 3, H // 2 + 5)], fill=(20, 20, 30, 255))
    return Image.alpha_composite(out.convert("RGBA"), ov).convert("RGB")


def af_crop2(img, box, p, W, H, bg=AF_STAGE_BG):
    """ขั้น ④ — แถบมืดเลื่อนเข้าหากรอบ 16:9 + มุมกรอบสีทอง"""
    base = af_contain(img, W, H, bg)
    s = min(W / img.width, H / img.height)
    ox, oy = (W - img.width * s) / 2, (H - img.height * s) / 2
    e = af_ease(p)
    L0, T0, R0, B0 = ox, oy, ox + img.width * s, oy + img.height * s
    l = L0 + (ox + box[0] * s - L0) * e
    t = T0 + (oy + box[1] * s - T0) * e
    r = R0 + (ox + box[2] * s - R0) * e
    b = B0 + (oy + box[3] * s - B0) * e
    ov = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    od = ImageDraw.Draw(ov)
    for rr in ((0, 0, W, t), (0, b, W, H), (0, t, l, b), (r, t, W, b)):
        if rr[2] > rr[0] + 0.5 and rr[3] > rr[1] + 0.5:
            od.rectangle(rr, fill=(0, 0, 0, int(160 * e)))
    gold = (255, 196, 40, 255)
    for x in range(int(l), int(r), 3):
        pass
    od.rectangle([l, t, r, b], outline=(255, 196, 40, int(120 * e)), width=1)
    for gx in (1 / 3, 2 / 3):
        od.line([(l + (r - l) * gx, t), (l + (r - l) * gx, b)], fill=(255, 255, 255, int(55 * e)), width=1)
        od.line([(l, t + (b - t) * gx), (r, t + (b - t) * gx)], fill=(255, 255, 255, int(55 * e)), width=1)
    k = 22
    for (cx, cy, sx, sy) in ((l, t, 1, 1), (r, t, -1, 1), (l, b, 1, -1), (r, b, -1, -1)):
        od.line([(cx, cy), (cx + sx * k, cy)], fill=gold, width=5)
        od.line([(cx, cy), (cx, cy + sy * k)], fill=gold, width=5)
    return Image.alpha_composite(base.convert("RGBA"), ov).convert("RGB")


def af_crop3(img, box0, box1, p, W, H, bg=AF_STAGE_BG):
    """[FEATURE #af-scene] ขั้น ④ — กรอบทองขึ้นกลางภาพก่อน แล้วค่อยๆ เลื่อนไปให้จออยู่กลาง (box ในพิกัดของ img)"""
    base = af_contain(img, W, H, bg)
    s = min(W / img.width, H / img.height)
    ox, oy = (W - img.width * s) / 2, (H - img.height * s) / 2
    fade = af_ease(min(1.0, p / 0.3))
    mv = af_ease(max(0.0, min(1.0, (p - 0.35) / 0.45)))
    bx = [box0[k] + (box1[k] - box0[k]) * mv for k in range(4)]
    l, t, r, b = ox + bx[0] * s, oy + bx[1] * s, ox + bx[2] * s, oy + bx[3] * s
    ov = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    od = ImageDraw.Draw(ov)
    for rr in ((0, 0, W, t), (0, b, W, H), (0, t, l, b), (r, t, W, b)):
        if rr[2] > rr[0] + 0.5 and rr[3] > rr[1] + 0.5:
            od.rectangle(rr, fill=(0, 0, 0, int(160 * fade)))
    gold = (255, 196, 40, int(255 * fade))
    od.rectangle([l, t, r, b], outline=(255, 196, 40, int(120 * fade)), width=1)
    for gx in (1 / 3, 2 / 3):
        od.line([(l + (r - l) * gx, t), (l + (r - l) * gx, b)], fill=(255, 255, 255, int(55 * fade)), width=1)
        od.line([(l, t + (b - t) * gx), (r, t + (b - t) * gx)], fill=(255, 255, 255, int(55 * fade)), width=1)
    k = 22
    for (cx, cy, sx, sy) in ((l, t, 1, 1), (r, t, -1, 1), (l, b, 1, -1), (r, b, -1, -1)):
        od.line([(cx, cy), (cx + sx * k, cy)], fill=gold, width=5)
        od.line([(cx, cy), (cx, cy + sy * k)], fill=gold, width=5)
    return Image.alpha_composite(base.convert("RGBA"), ov).convert("RGB")


def _af_split_check(text):
    """ตัดเครื่องหมาย ✓ นำหน้าออก (ฟอนต์ไทยบางตัวไม่มีตัวนี้ → ขึ้นเป็นสี่เหลี่ยม) · คืน (มีไหม, ข้อความที่เหลือ)"""
    t = text or ""
    if t.startswith("✓"):
        return True, t[1:].lstrip()
    return False, t


def _af_check(d, x, y, size, fill, width=None):
    """วาดเครื่องหมายถูกเป็นเส้น (ไม่พึ่งฟอนต์) · x,y = มุมซ้ายบน"""
    w_ = width or max(2, int(size * 0.16))
    d.line([(x + size * 0.08, y + size * 0.55), (x + size * 0.38, y + size * 0.84), (x + size * 0.92, y + size * 0.22)],
           fill=fill, width=w_, joint="curve")


def af_text(d, xy, text, font, fill):
    """d.text ที่วาด ✓ นำหน้าเป็นเส้นแทนตัวอักษร · คืนความกว้างที่ใช้"""
    x, y = xy
    has, rest = _af_split_check(text)
    dx = 0
    if has:
        sz = font.size * 0.9
        _af_check(d, x, y + font.size * 0.12, sz, fill)
        dx = sz + font.size * 0.45
    d.text((x + dx, y), rest, font=font, fill=fill)
    try:
        return dx + d.textlength(rest, font=font)
    except Exception:
        return dx + len(rest) * 8


def _af_text_w(d, text, font):
    has, rest = _af_split_check(text)
    try:
        w_ = d.textlength(rest, font=font)
    except Exception:
        w_ = len(rest) * 8
    return w_ + (font.size * 1.35 if has else 0)


def _af_pill(d, x, y, text, font, fill=(0, 0, 0, 190), fg=(255, 255, 255, 255), pad=14, h=34):
    tw = _af_text_w(d, text, font)
    d.rounded_rectangle([x, y, x + tw + pad * 2, y + h], radius=h // 2, fill=fill)
    af_text(d, (x + pad, y + (h - font.size) / 2 - 2), text, font, fg)
    return tw + pad * 2


def af_resize2(img, p, src_wh, dst_wh, font, W, H, bg=AF_STAGE_BG):
    """ขั้น ⑤ — ภาพหดเข้าแล้วเด้งกลับ + ตัวเลขขนาดนับลง"""
    e = af_ease(p)
    sc = 1 - 0.14 * math.sin(math.pi * min(1.0, p * 1.1))
    base = Image.new("RGB", (W, H), bg)
    iw, ih = int(W * sc), int(H * sc)
    im = af_cover(img, iw, ih)
    base.paste(im, ((W - iw) // 2, (H - ih) // 2), _af_mask(iw, ih, 12))
    ov = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    d = ImageDraw.Draw(ov)
    cw = int(src_wh[0] + (dst_wh[0] - src_wh[0]) * e)
    chh = int(src_wh[1] + (dst_wh[1] - src_wh[1]) * e)
    txt = f"{cw} × {chh}"
    try:
        tw = d.textlength(txt, font=font)
    except Exception:
        tw = 120
    _af_pill(d, (W - tw - 28) / 2, H - 56, txt, font, fill=(10, 12, 22, 210), fg=(255, 216, 102, 255))
    return Image.alpha_composite(base.convert("RGBA"), ov).convert("RGB")


def af_stamp2(base_img, stamped, p, W, H):
    """ขั้น ⑥ — ภาพจางเป็นเวอร์ชันประทับวันที่ + วงแสงตรงจุดประทับ"""
    a = af_cover(base_img, W, H)
    b = af_cover(stamped, W, H)
    e = af_ease(min(1.0, p * 1.4))
    out = Image.blend(a, b, e).convert("RGBA")
    ov = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    d = ImageDraw.Draw(ov)
    cx = W * (1 - (STAMP_MARGIN_RIGHT + 330) / TARGET_WIDTH)
    cy = H * (1 - (STAMP_MARGIN_BOTTOM + 70) / TARGET_HEIGHT)
    rr = 20 + 90 * af_ease_out(p)
    al = int(200 * (1 - af_ease(p)))
    d.ellipse([cx - rr * 1.6, cy - rr * 0.7, cx + rr * 1.6, cy + rr * 0.7], outline=(254, 98, 2, al), width=3)
    return Image.alpha_composite(out, ov).convert("RGB")


def af_rename2(img, old, new, p, font, W, H):
    """ขั้น ⑦ — [FEATURE #af-rename-b] แถวเดียว: ส่วนที่ตัดออก (เลขท้าย) ขึ้นกรอบแดง + ป้าย "ตัดเลขท้ายออก"
    → หดหายไป ตัวหนังสือเลื่อนชิดกัน → ทั้งชื่อเปลี่ยนเป็นสีเขียว + ติ๊ก (ชื่อใหม่สั้นกว่าเดิมเสมอ ไม่ล้น)"""
    import difflib
    out = af_cover(img, W, H).convert("RGBA")
    ov = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    d = ImageDraw.Draw(ov)
    x = 24
    tl = lambda t_: (d.textlength(t_, font=font) if t_ else 0)
    ck_w = font.size * 1.35
    avail = W - x * 2 - ck_w
    if tl(old) > avail and hasattr(font, "font_variant"):     # ชื่อยาวมาก → ย่อตัวอักษรให้พอดีแถว
        try:
            font = font.font_variant(size=max(10, int(font.size * avail / max(1, tl(old)))))
            ck_w = font.size * 1.35
        except Exception:
            pass
    ph = max(40, int(font.size * 2.7))
    e0 = af_ease(min(1.0, p / 0.12))
    top = H - ph * e0
    d.rectangle([0, top, W, H], fill=(8, 10, 18, int(225 * e0)))
    d.rectangle([0, top, W, top + 2], fill=(236, 72, 153, int(255 * e0)))
    y = top + (ph - font.size) / 2 - 3
    ops = difflib.SequenceMatcher(None, old, new, autojunk=False).get_opcodes()
    morph = all(op[0] in ("equal", "delete") for op in ops)
    eh = af_ease(max(0.0, min(1.0, (p - 0.14) / 0.25)))      # กรอบแดงขยาย
    ec = af_ease(max(0.0, min(1.0, (p - 0.45) / 0.30)))      # ส่วนที่ตัดหดหาย
    eg = af_ease(max(0.0, min(1.0, (p - 0.75) / 0.20)))      # เปลี่ยนเป็นสีเขียว + ติ๊ก
    g0, g1 = (200, 208, 224), (52, 211, 153)
    fg = tuple(int(g0[j] + (g1[j] - g0[j]) * eg) for j in range(3)) + (int(255 * e0),)
    off = 0.0
    if eg > 0:
        _af_check(d, x, y + font.size * 0.12, font.size * 0.9, g1 + (int(255 * eg),))
        off = ck_w * eg
    if not morph:                                            # ชื่อเปลี่ยนแบบอื่น (ไม่ใช่แค่ตัดออก) → จางสลับ
        if eg < 1:
            d.text((x + off, y), old, font=font, fill=g0 + (int(255 * e0 * (1 - eg)),))
        d.text((x + off, y), new, font=font, fill=g1 + (int(255 * eg),))
        return Image.alpha_composite(out, ov).convert("RGB")
    cx = x + off
    pill_x = None
    for tag, i1, i2, _j1, _j2 in ops:
        seg = old[i1:i2]
        if not seg:
            continue
        if tag == "equal":
            d.text((cx, y), seg, font=font, fill=fg)
            cx += tl(seg)
            continue
        full = tl(seg)
        wc = full * (1 - ec)
        if wc > 1:
            cimg = Image.new("RGBA", (int(full) + 6, int(font.size * 1.6)), (0, 0, 0, 0))
            cd = ImageDraw.Draw(cimg)
            cd.rounded_rectangle([0, 0, 3 + (full + 2) * eh, cimg.height - 1], radius=6,
                                 fill=(255, 90, 103, int(110 * eh)), outline=(255, 90, 103, int(255 * eh)), width=2)
            cd.text((3, 3), seg, font=font, fill=(255, 200, 205, int(255 * (1 - ec) * e0)))
            cimg = cimg.crop((0, 0, max(1, int(wc) + 3), cimg.height))
            ov.alpha_composite(cimg, (int(cx - 3), int(y - 3)))
            if pill_x is None:
                pill_x = cx
        cx += wc
    if pill_x is not None and eh > 0.6 and ec < 0.9:
        k_ = 1 - ec
        pw_ = _af_text_w(d, "ตัดเลขท้ายออก", font) + 20
        _af_pill(d, max(8, min(pill_x - 10, W - pw_ - 8)), top - font.size * 1.9, "ตัดเลขท้ายออก", font, fill=(255, 90, 103, int(230 * k_)),
                 fg=(255, 255, 255, int(255 * k_)), pad=10, h=int(font.size * 1.6))
    return Image.alpha_composite(out, ov).convert("RGB")


def af_caption2(img, text, font, W, H, col=(10, 12, 22, 215), fg=(255, 255, 255, 255), pos="bl"):
    """ภาพเต็มเวที + ป้ายข้อความทรงแคปซูล (มุมล่าง)"""
    out = af_cover(img, W, H).convert("RGBA")
    ov = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    d = ImageDraw.Draw(ov)
    for y in range(H // 3):
        d.line([(0, H - y), (W, H - y)], fill=(0, 0, 0, int(120 * (1 - y / (H / 3)))))
    hh = max(30, int(font.size * 1.9))
    tw = _af_text_w(d, text, font)
    x = 18 if pos == "bl" else (W - tw - 28) / 2
    _af_pill(d, x, H - hh - 16, text, font, fill=col, fg=fg, h=hh)
    return Image.alpha_composite(out, ov).convert("RGB")


def af_fly2(img, p, slide_img, W, H, bg=AF_STAGE_BG, target=None):
    """ขั้น ⑧ — รูปหดลอยลงช่องในสไลด์ (target = กรอบช่องในพิกัดของ slide_img)"""
    e = af_ease(p)
    base = Image.new("RGB", (W, H), bg)
    sw_, sh_ = slide_img.size
    sx, sy = (W - sw_) // 2, (H - sh_) // 2
    # เงาสไลด์
    ov = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    od = ImageDraw.Draw(ov)
    for k in range(10, 0, -1):
        od.rounded_rectangle([sx - k, sy - k + 4, sx + sw_ + k, sy + sh_ + k + 4], radius=10 + k, fill=(0, 0, 0, 10))
    base = Image.alpha_composite(base.convert("RGBA"), ov).convert("RGB")
    base.paste(slide_img, (sx, sy), _af_mask(sw_, sh_, 8))
    if target is None:
        target = (sw_ * 0.55, sh_ * 0.55, sw_ * 0.4, sh_ * 0.4)
    tx, ty, tw, th = sx + target[0], sy + target[1], target[2], target[3]
    cw, ch = W + (tw - W) * e, H + (th - H) * e
    cx, cy = tx * e, ty * e
    im = af_cover(img, max(2, int(cw)), max(2, int(ch)))
    base.paste(im, (int(cx), int(cy)), _af_mask(im.width, im.height, max(3, int(14 * (1 - e) + 3))))
    if e > 0.85:
        ov = Image.new("RGBA", (W, H), (0, 0, 0, 0))
        a = int(255 * (e - 0.85) / 0.15)
        ImageDraw.Draw(ov).rounded_rectangle([tx - 3, ty - 3, tx + tw + 3, ty + th + 3], radius=5,
                                             outline=(255, 184, 0, a), width=3)
        base = Image.alpha_composite(base.convert("RGBA"), ov).convert("RGB")
    return base


def af_slide_big(filled, slots, w, h, hl=None, img=None, title=None, font=None):
    """สไลด์จำลองขนาดใหญ่ (ขั้น ⑧) — คืน (ภาพ, กรอบช่องทั้งหมด)
    ช่องที่วางแล้ว = ฟ้า · ช่อง hl = ขอบทอง (ถ้าส่ง img มาจะใส่รูปในช่องนั้น)"""
    S = 2
    im = Image.new("RGB", (w * S, h * S), (250, 251, 253))
    d = ImageDraw.Draw(im)
    bar = int(h * 0.13) * S
    d.rectangle([0, 0, w * S, bar], fill=(40, 44, 110))
    d.rectangle([0, bar, w * S, bar + 3 * S], fill=(236, 72, 153))
    n = max(1, slots)
    cols = 2 if n > 1 else 1
    rows = (n + cols - 1) // cols
    pad, gap = int(w * 0.05) * S, int(w * 0.03) * S
    top = bar + int(h * 0.06) * S
    cw = (w * S - pad * 2 - gap * (cols - 1)) / cols
    ch = (h * S - top - pad - gap * (rows - 1)) / rows
    if ch * 16 / 9 < cw:
        cw2 = ch * 16 / 9
        pad += int((cw - cw2) * cols / 2)
        cw = cw2
    else:
        ch = cw * 9 / 16
    rects = []
    for i in range(n):
        r_, c_ = divmod(i, cols)
        x0, y0 = pad + c_ * (cw + gap), top + r_ * (ch + gap)
        d.rounded_rectangle([x0, y0, x0 + cw, y0 + ch], radius=4 * S,
                            fill=(111, 168, 220) if i < filled and i != hl else (222, 226, 234))
        if i >= filled and i != hl:
            d.text((x0 + cw / 2 - 4 * S, y0 + ch / 2 - 7 * S), "+", fill=(170, 176, 190))
        rects.append((x0 / S, y0 / S, cw / S, ch / S))
    im = im.resize((w, h), Image.LANCZOS)
    if hl is not None and 0 <= hl < len(rects):
        x0, y0, cw, ch = rects[hl]
        if img is not None:
            t = af_cover(img, max(2, int(cw)), max(2, int(ch)))
            im.paste(t, (int(x0), int(y0)), _af_mask(t.width, t.height, 3))
        ImageDraw.Draw(im).rounded_rectangle([x0 - 2, y0 - 2, x0 + cw + 2, y0 + ch + 2], radius=5,
                                             outline=(255, 184, 0), width=3)
    if title and font:
        ImageDraw.Draw(im).text((int(w * 0.04), int(h * 0.13 / 2 - font.size * 0.62)), title, font=font, fill=(255, 255, 255))
    return im, rects


_AF_CHIP_GRAD = {}
# [FEATURE #af-logo] โลโก้ Gemini (PNG 112px ฝังในโค้ด ไม่ต้องแนบไฟล์) แทนดาวสีเหลืองบนป้าย AI ENGINE
AF_GEMINI_LOGO_B64 = (
    "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"
    "lFTkSuQmCC"
)
_AF_LOGO = {}


def af_gemini_logo(size, gray=False):
    """โลโก้ Gemini ขนาด size (RGBA) · gray = สีเทา (โหมดไม่ใช้ AI) · เก็บแคชตามขนาด"""
    size = max(8, int(size))
    key = (size, gray)
    im = _AF_LOGO.get(key)
    if im is None:
        import base64, io
        src = _AF_LOGO.get("src")
        if src is None:
            src = _AF_LOGO["src"] = Image.open(io.BytesIO(base64.b64decode(AF_GEMINI_LOGO_B64))).convert("RGBA")
        im = src.resize((size, size), Image.LANCZOS)
        if gray:
            a = im.getchannel("A")
            im = Image.merge("RGBA", (im.convert("L"),) * 3 + (a.point(lambda v: v * 0.7),))
        if len(_AF_LOGO) > 40:
            _AF_LOGO.clear()
        _AF_LOGO[key] = im
    return im


def af_ai_chip(state, t, model, msg, W, H, font_s, font_b, bg=AF_STAGE_BG):
    """[DEMO #auto-flow] ป้าย AI ENGINE บนเวที — ขอบไล่สีวิ่ง · โลโก้ Gemini เต้นตอนคิด · แท่งเสียงเต้นตามสถานะ
    state: off / idle / think / done / cache / quota / err"""
    g = _AF_CHIP_GRAD.get(W)
    if g is None:
        stops = [(236, 72, 153), (139, 92, 246), (34, 211, 238), (255, 184, 0), (236, 72, 153)]
        g = Image.new("RGB", (W * 2, 1))
        px = g.load()
        for x in range(W * 2):
            u = (x % W) / W * (len(stops) - 1)
            k = int(u)
            f_ = u - k
            a, b = stops[k], stops[min(k + 1, len(stops) - 1)]
            px[x, 0] = tuple(int(a[j] + (b[j] - a[j]) * f_) for j in range(3))
        _AF_CHIP_GRAD[W] = g
    think = state == "think"
    off = state == "off"
    speed = {"think": 320, "done": 90, "cache": 60}.get(state, 30)
    off_x = int(t * speed) % W
    ring_col = g.crop((off_x, 0, off_x + W, 1)).resize((W, H))
    if off:
        ring_col = Image.new("RGB", (W, H), (70, 80, 104))
    elif state == "quota" or state == "err":
        ring_col = Image.new("RGB", (W, H), (255, 90, 103))
    out = Image.new("RGB", (W, H), bg)
    bw = 3 if think else 2
    outer = _af_mask(W, H, H // 2)
    inner = Image.new("L", (W, H), 0)
    inner.paste(_af_mask(W - bw * 2, H - bw * 2, (H - bw * 2) // 2), (bw, bw))
    out.paste(ring_col, (0, 0), outer)
    pulse = 0.5 + 0.5 * math.sin(t * (9 if think else 2.4))
    fill = (20, 24, 40)
    if think:
        fill = tuple(int(fill[j] + ((58, 44, 10)[j] - fill[j]) * (0.35 + 0.35 * pulse)) for j in range(3))
    elif state in ("done", "cache"):
        fill = (14, 40, 34)
    out.paste(Image.new("RGB", (W, H), fill), (0, 0), inner)
    d = ImageDraw.Draw(out, "RGBA")
    # [FEATURE #af-logo] โลโก้ Gemini (เดิมดาวเหลือง) · ตอนคิด = ขยาย–หดช้าๆ + แสงฟุ้ง (ไม่หมุน)
    cx, cy, r = H // 2 + 4, H // 2, H * 0.30
    cc = (255, 196, 40) if not off else (120, 130, 150)
    if state in ("done", "cache"):
        cc = (52, 211, 153)
    lp = 0.5 + 0.5 * math.sin(t * 3.2)
    sz = int(H * (0.60 + (0.12 * lp if think else 0.0)))
    lg = af_gemini_logo(sz, gray=off)
    lx, ly = int(cx - sz / 2), int(cy - sz / 2)
    if think:
        gsz = int(sz * 1.5)
        glow = af_gemini_logo(gsz).filter(ImageFilter.GaussianBlur(max(2, H // 12)))
        a_ = glow.getchannel("A").point(lambda v, k=0.45 + 0.4 * lp: int(v * k))
        glow.putalpha(a_)
        out.paste(glow, (int(cx - gsz / 2), int(cy - gsz / 2)), glow)
    out.paste(lg, (lx, ly), lg)
    # ข้อความ
    tx = int(H * 1.05) + 6
    top = {"think": "AI กำลังวิเคราะห์" + "." * (int(t * 3) % 4), "done": "✓ AI ตอบแล้ว",
           "cache": "✓ ใช้ผลเดิม", "quota": "โควต้าหมด", "err": "AI ไม่ตอบ", "off": "ไม่ใช้ AI",
           }.get(state, "AI ENGINE · GEMINI")
    if msg and state in ("done", "cache", "quota", "err"):
        top = msg
    tc = {"think": (255, 216, 102), "done": (110, 231, 183), "cache": (110, 231, 183),
          "quota": (255, 150, 160), "err": (255, 150, 160), "off": (140, 150, 170)}.get(state, (155, 239, 255))
    af_text(d, (tx, H * 0.14), top, font_s, tc + (255,))
    d.text((tx, H * 0.14 + font_s.size * 1.25), model if not off else "เกณฑ์ในเครื่อง",
           font=font_b, fill=(255, 255, 255, 255) if not off else (170, 178, 196, 255))
    # แท่งเสียง (equalizer)
    nb, bwid, gap = 5, max(3, H // 12), max(2, H // 16)
    bx0 = W - H * 0.45 - nb * (bwid + gap)
    for i in range(nb):
        if think:
            v = 0.25 + 0.75 * abs(math.sin(t * (7 + i * 1.7) + i * 1.3))
        elif state in ("done", "cache"):
            v = 0.35 + 0.25 * abs(math.sin(t * 2 + i))
        elif off:
            v = 0.12
        else:
            v = 0.18 + 0.12 * abs(math.sin(t * 1.5 + i * 0.9))
        bh = (H * 0.56) * v
        x0 = bx0 + i * (bwid + gap)
        d.rounded_rectangle([x0, cy - bh / 2, x0 + bwid, cy + bh / 2], radius=bwid // 2,
                            fill=(cc if not off else (90, 100, 120)) + (255,))
    return out


def af_fit_target(img):
    """ปรับรูปให้เป็นขนาดมาตรฐาน TARGET (ครอปกลาง) ก่อนประทับวันที่"""
    if img.size == (TARGET_WIDTH, TARGET_HEIGHT):
        return img, False
    tw, th = TARGET_WIDTH, TARGET_HEIGHT
    s = max(tw / img.width, th / img.height)
    nw, nh = max(tw, int(round(img.width * s))), max(th, int(round(img.height * s)))
    img = img.resize((nw, nh), Image.LANCZOS)
    l, t = (nw - tw) // 2, (nh - th) // 2
    return img.crop((l, t, l + tw, t + th)), True


# ======================================================================
# ส่วนที่ 2 · หน้าต่าง AUTO FLOW (ผูกเป็นเมธอดของ InspectPilotApp ผ่าน install())
# ======================================================================
# ================================================================
# [DEMO #auto-flow v3] หน้าต่าง AUTO FLOW · AI Stage (จัดเต็ม)
# ขนาดปรับตามจอ · การ์ดมุมมนเรืองแสง · ภาพเคลื่อนไหวทุกขั้น (easing 30fps)
# ================================================================
AF_ACCENT = "#EC4899"
AF_STATIONS = []       # สร้างใน install() เพราะใช้ขนาดรูปมาตรฐานจาก InspectPilot


def _af_build_stations():
    return [("ดึงจาก NAS", "ระบบ Fetch"), ("ตรวจ + เลือก", "เบลอ/มืด · AI"), ("แต่งสี", "ปุ่มของทีม"),
            ("ครอป 16:9", "ตัดขอบ"), ("รีไซส์", f"{TARGET_WIDTH}×{TARGET_HEIGHT}"), ("ประทับวันที่", "EXIF"),
            ("ตั้งชื่อไฟล์", "ตัดเลขท้าย"), ("ลง PPT", "ตามช่อง")]
AF_CHIPS = ["ดึง", "เลือก", "แต่งสี", "ครอป", "รีไซส์", "Stamp", "ชื่อไฟล์", "PPT"]

def _open_auto_flow(self):
    old = getattr(self, "_af_win", None)
    if old is not None:
        try:
            if old.winfo_exists():
                old.lift()
                old.focus_force()
                return
        except tk.TclError:
            pass
    global AF_STAGE_W, AF_STAGE_H
    PINK, GOLD_, CY, GRN, ORG, VIO = self.AF_ACCENT, "#FFB800", "#22D3EE", "#34D399", "#FF8C42", "#8B5CF6"
    BG0, PAN, CARDC, LINE = "#0C0F1A", "#121726", "#161C2D", "#252D45"
    SUBC, DIMC = "#A6B0C8", "#6C7791"
    cfg0 = dict(self.config_data.get("autoflow", {}) or {})
    sw_, sh_ = self.winfo_screenwidth(), self.winfo_screenheight()
    W = min(1580, int(sw_ * 0.92))
    H = min(960, int(sh_ * 0.87))      # เผื่อ taskbar ของ Windows
    k = max(0.78, min(1.08, W / 1500))
    f = lambda n: max(8, int(round(n * k)))
    LW = int(W * 0.25)
    SWd = W - 20 - 44 - LW - 18 - 18
    small = H < 800
    TGAP = f(12)
    TWt = int((SWd - 28 - 36 - 7 * TGAP) / 8)
    THt = int(TWt * 9 / 16)
    others = f(58) + 12 + 14 + (18 + THt + f(52) + f(34)) + 12 + (0 if small else 74 + f(20)) + f(34) + 30
    BHmax = (H - 20 - 44 - f(74) - 14 - 18) - others
    BW = int(min(SWd * 0.58, BHmax * 16 / 9))
    BH = int(BW * 9 / 16)
    AF_STAGE_W, AF_STAGE_H = BW, BH
    EXW = SWd - 28 - BW - 16
    win = tk.Toplevel(self)
    self._af_win = win
    win.withdraw()
    win.title("AUTO FLOW · Demo")
    win.configure(bg=BG_MAIN)
    win.transient(self)
    win.resizable(False, False)
    outer, card = self._glow_card(win, PINK, radius=18, bg=BG0, border_alpha=170, glow_alpha=40)
    outer.configure(width=W, height=H)
    outer.pack(padx=10, pady=10)
    body = tk.Frame(card, bg=BG0)
    body.pack(fill="both", expand=True, padx=22, pady=16)
    st = {"running": False, "stop": False, "out": None, "ppt": None, "csv": None, "png": None,
          "log": [], "ph": {}, "anim": None, "t0": 0, "active": None, "pulse": 0.0, "think": False,
          "cnt": {"auto": 0, "human": 0, "bad": 0}}
    F_TITLE, F_H, F_B, F_S, F_XS = (FONT_TH, f(17), "bold"), (FONT_TH, f(11), "bold"), (FONT_TH, f(10)), (FONT_TH, f(9)), (FONT_TH, f(8))
    pil_f = self._ss_pil_font(max(12, int(BH / 17)), True)
    pil_fs = self._ss_pil_font(max(11, int(BH / 21)), True)

    def photo(key, pil):
        ph_ = ImageTk.PhotoImage(pil)
        st["ph"][key] = ph_
        return ph_

    def gcard(parent, col, bg=CARDC, radius=14, alpha=110, glow=0):
        o, c_ = self._glow_card(parent, col, radius=radius, bg=bg, border_alpha=alpha, glow_alpha=glow)
        return o, c_

    def fit(o, c_, r=14):
        try:
            self._autosize_card_height(o, c_, radius=r)
        except Exception:
            pass

    # ---------- หัว ----------
    head = tk.Frame(body, bg=BG0)
    head.pack(fill="x")
    self._neon_icon(head, f(48), BG0, self._draw_rocket_glyph, PINK).pack(side="left")
    hc = tk.Frame(head, bg=BG0)
    hc.pack(side="left", padx=(f(14), 0))
    tr_ = tk.Frame(hc, bg=BG0)
    tr_.pack(anchor="w")
    tk.Label(tr_, text="AUTO FLOW", font=F_TITLE, bg=BG0, fg="#FFFFFF").pack(side="left")
    tk.Label(tr_, text="  AI เลือกรูป · แต่งสี · ลงสไลด์ ในปุ่มเดียว", font=F_S, bg=BG0, fg=SUBC).pack(side="left", pady=(f(4), 0))
    chips_f = tk.Frame(hc, bg=BG0)
    chips_f.pack(anchor="w", pady=(f(6), 0))
    chips = []
    for i, t_ in enumerate(self.AF_CHIPS):
        if i:
            tk.Label(chips_f, text="›", font=F_S, bg=BG0, fg=DIMC).pack(side="left", padx=2)
        lb = tk.Label(chips_f, text=f" {i + 1}  {t_} ", font=(FONT_TH, f(8), "bold"), bg=CARDC, fg=DIMC, padx=f(6), pady=1)
        lb.pack(side="left")
        chips.append(lb)
    hr = tk.Frame(head, bg=BG0)
    hr.pack(side="right", anchor="n")
    tk.Label(hr, text="  DEMO · สาธิตเท่านั้น  ", font=(FONT_TH, f(8), "bold"), bg=_lerp_hex(BG0, PINK, 0.35), fg="#FFD1E8").pack(anchor="e")
    mchip = tk.Label(hr, text="", font=("Consolas", f(9), "bold"), bg=_lerp_hex(BG0, CY, 0.16), fg="#9BEFFF", padx=f(10), pady=2)
    gl = tk.Canvas(body, height=3, bg=BG0, highlightthickness=0)
    gl.pack(fill="x", pady=(f(10), f(12)))
    gl_ph = {}

    def draw_gl(e=None):
        w_ = max(10, gl.winfo_width())
        im = Image.new("RGB", (w_, 3))
        a_, b_, c_ = _af_rgb(PINK), _af_rgb(VIO), _af_rgb(CY)
        px = im.load()
        for x in range(w_):
            t_ = x / max(1, w_ - 1)
            col = tuple(int(a_[j] + (b_[j] - a_[j]) * min(1, t_ * 2)) if t_ < .5 else int(b_[j] + (c_[j] - b_[j]) * (t_ - .5) * 2) for j in range(3))
            for y in range(3):
                px[x, y] = col
        gl_ph["i"] = ImageTk.PhotoImage(im)
        gl.delete("all")
        gl.create_image(0, 0, anchor="nw", image=gl_ph["i"])
    gl.bind("<Configure>", draw_gl)

    main = tk.Frame(body, bg=BG0)
    main.pack(fill="both", expand=True)
    left = tk.Frame(main, bg=BG0, width=LW)
    left.pack(side="left", fill="y")
    left.pack_propagate(False)
    so, sc = gcard(main, GOLD_, bg=PAN, radius=18, alpha=140, glow=26)
    so.pack(side="left", fill="both", expand=True, padx=(18, 0))
    SB = PAN

    # ---------- ซ้าย: ตั้งค่า ----------
    setup = tk.Frame(left, bg=BG0)
    runf = tk.Frame(left, bg=BG0)
    widgets_lock = []
    vars_ = {
        "tpl": tk.StringVar(value=cfg0.get("tpl") or self.pptx_template_path.get()),
        "out": tk.StringVar(value=cfg0.get("out", "")),
        "p1": tk.StringVar(value=str(cfg0.get("start") or self.pptx_start_page.get() or "3")),
        # [FEATURE #pptx-auto-pages] ขนาดรูปในไฟล์ Standard / Original (HD)
        "q": tk.StringVar(value=(cfg0.get("quality") if cfg0.get("quality") in PP_QUALITY
                                 else getattr(self, "pptx_quality_var", None) and self.pptx_quality_var.get() or "standard")),
        "key": tk.StringVar(value=cfg0.get("api_key", "")),
        "model": tk.StringVar(value=_af_saved_model(cfg0)),
        "thr": tk.StringVar(value=str(cfg0.get("threshold", AF_DEFAULT_THRESHOLD))),
        "prod": tk.StringVar(value=cfg0.get("product", "")),
    }
    use_ai = {"on": bool(cfg0.get("use_ai", True))}

    def sec_head(parent, glyph, col, title, bg=CARDC, right=None):
        h_ = tk.Frame(parent, bg=bg)
        h_.pack(fill="x", padx=f(14), pady=(f(10), f(5)))
        self._neon_icon(h_, f(24), bg, glyph, col).pack(side="left")
        tk.Label(h_, text=title, font=F_H, bg=bg, fg="#FFFFFF").pack(side="left", padx=(f(8), 0))
        if right:
            right(h_)
        return h_

    def entry_row(parent, label, var, browse=None, show=None, bg=CARDC, padx=None, inline=None):
        px_ = f(14) if padx is None else padx
        if label and not inline:
            tk.Label(parent, text=label, font=(FONT_TH, f(8), "bold"), bg=bg, fg=SUBC, anchor="w").pack(fill="x", padx=px_)
        row = tk.Frame(parent, bg=bg)
        row.pack(fill="x", padx=px_, pady=(2, f(4) if inline else f(6)))
        if label and inline:       # ป้ายอยู่แถวเดียวกับช่อง (ประหยัดความสูง)
            tk.Label(row, text=label, font=(FONT_TH, f(8), "bold"), bg=bg, fg=SUBC, anchor="w",
                     width=inline).pack(side="left", padx=(0, f(4)))
        box = self._round_field(row)
        box.pack(side="left", fill="x", expand=True)
        e = tk.Entry(box, textvariable=var, font=("Consolas", f(9)), bg=INPUT_BG, fg=TEXT_PRIMARY,
                     insertbackground=TEXT_PRIMARY, relief="flat", bd=0, highlightthickness=0, show=show or "")
        e.pack(fill="x", padx=f(10), pady=f(5))
        widgets_lock.append(e)
        if browse:
            b = self._pill_button(row, "เลือก", _lerp_hex(bg, BRAND, 0.6), "#FFFFFF", command=browse,
                                  font=(FONT_TH, f(8), "bold"), padx=f(10), pady=f(5))
            b.pack(side="left", padx=(f(6), 0))
            widgets_lock.append(b)
        return e

    def pick_dir(var, title):
        d = filedialog.askdirectory(parent=win, title=title)
        if d:
            var.set(d)

    def pick_ppt():
        fl = filedialog.askopenfilename(parent=win, title="เลือกเทมเพลต PowerPoint", filetypes=[("PowerPoint", "*.pptx")])
        if fl:
            vars_["tpl"].set(fl)

    # การ์ด ① NAS
    o1, c1 = gcard(setup, CY)
    o1.pack(fill="x")

    def _ref(h_):
        lb = tk.Label(h_, text="↻ อ่านใหม่", font=(FONT_TH, f(8), "bold"), bg=CARDC, fg="#9DA0FF", cursor="hand2")
        lb.pack(side="right")
        lb.bind("<Button-1>", lambda e: refresh_nas())
    sec_head(c1, self._draw_folder_glyph, CY, "แหล่งรูป · NAS", right=_ref)
    # ช่องกรอกสินค้า/Playlist — กรอกที่นี่ได้เลย (ค่าอื่นอ่านจากแท็บ Fetch)
    tk.Label(c1, text="สินค้า / Playlist  ·  หลายตัวคั่นด้วย ,", font=(FONT_TH, f(8), "bold"), bg=CARDC,
             fg="#9BEFFF", anchor="w").pack(fill="x", padx=f(14))
    prow = tk.Frame(c1, bg=CARDC)
    prow.pack(fill="x", padx=f(14), pady=(3, f(8)))
    pbox = self._round_field(prow)
    pbox.pack(fill="x")
    prod_e = tk.Entry(pbox, textvariable=vars_["prod"], font=(FONT_TH, f(11), "bold"), bg=INPUT_BG, fg="#FFFFFF",
                      insertbackground="#FFFFFF", relief="flat", bd=0, highlightthickness=0)
    prod_e.pack(fill="x", padx=f(10), pady=f(6))
    widgets_lock.append(prod_e)
    prod_e.bind("<Return>", lambda e: refresh_nas(False))
    prod_e.bind("<FocusOut>", lambda e: refresh_nas(False))
    _pj = {"id": None}

    def _prod_changed(*_a):          # พิมพ์แล้วอัปเดตคำเตือนเอง (หน่วง 0.35 วิ)
        if _pj["id"]:
            try:
                win.after_cancel(_pj["id"])
            except Exception:
                pass
        _pj["id"] = win.after(350, lambda: (not st["running"]) and refresh_nas(False))
    vars_["prod"].trace_add("write", _prod_changed)
    try:
        self._bind_clipboard(prod_e)
    except Exception:
        pass
    nas_grid = tk.Frame(c1, bg=CARDC)
    nas_grid.pack(fill="x", padx=f(14), pady=(0, f(12)))
    nas_grid.columnconfigure(1, weight=1)
    nas_vals = {}
    for i_, (k_, t_) in enumerate((("drv", "Master Path"), ("scr", "จอ"), ("out", "Output"))):
        tk.Label(nas_grid, text=t_, font=F_XS, bg=CARDC, fg=DIMC, anchor="w").grid(row=i_, column=0, sticky="w", pady=1)
        v_ = tk.Label(nas_grid, text="—", font=(FONT_TH, f(9), "bold"), bg=CARDC, fg="#E6EAF4", anchor="w")
        v_.grid(row=i_, column=1, sticky="w", padx=(f(10), 0), pady=1)
        nas_vals[k_] = v_
    nas_err = tk.Label(c1, text="", font=F_XS, bg=CARDC, fg=WARNING, anchor="w", justify="left", wraplength=LW - 40)

    # การ์ด ② PowerPoint
    o2, c2 = gcard(setup, VIO)
    o2.pack(fill="x", pady=(f(8), 0))
    sec_head(c2, self._draw_document_glyph, VIO, "PowerPoint · แบบ Generate")
    entry_row(c2, "เทมเพลต", vars_["tpl"], pick_ppt, inline=7)
    out_e = entry_row(c2, "ผลลัพธ์", vars_["out"], lambda: pick_dir(vars_["out"], "เลือกโฟลเดอร์ผลลัพธ์"), inline=7)
    # ช่องว่าง = สร้างให้เอง → คำใบ้จางๆ ทับในช่อง (ไม่ได้ใส่ลงค่าจริง)
    out_ph = tk.Label(out_e.master, text="เว้นว่าง = สร้างให้เอง", font=("Consolas", f(9)), bg=INPUT_BG, fg=DIMC,
                      cursor="xterm")
    out_ph.bind("<Button-1>", lambda e: out_e.focus_set())

    def _ph_upd(*_a):
        try:
            if not vars_["out"].get() and win.focus_get() is not out_e:
                out_ph.place(in_=out_e, x=2, rely=0.5, anchor="w")
            else:
                out_ph.place_forget()
        except (tk.TclError, KeyError):
            pass
    out_e.bind("<FocusIn>", _ph_upd, add="+")
    out_e.bind("<FocusOut>", _ph_upd, add="+")
    vars_["out"].trace_add("write", _ph_upd)
    win.after(150, _ph_upd)
    # [FEATURE #pptx-auto-pages] แถบหน้า (เปิด/ต้นแบบ/ปิด) + หน้าต้นแบบอัตโนมัติ + ขนาดรูป
    pr_ = tk.Frame(c2, bg=CARDC)
    pr_.pack(fill="x", padx=f(14), pady=(0, f(4)))
    pp_strip = tk.Canvas(pr_, width=int(LW * 0.56), height=f(68), bg=CARDC, highlightthickness=0)
    if not small:          # จอเตี้ย: ไม่แสดงแถบหน้า เหลือบรรทัดเดียว (ประหยัดความสูง)
        pp_strip.pack(side="left", anchor="n")
    pr_r = tk.Frame(pr_, bg=CARDC)
    pr_r.pack(side="left", fill="both", expand=True, padx=(f(6), 0))
    r1_ = tk.Frame(pr_r, bg=CARDC)
    r1_.pack(fill="x")
    tk.Label(r1_, text="หน้าต้นแบบ", font=(FONT_TH, f(8), "bold"), bg=CARDC, fg="#FFFFFF").pack(side="left", padx=(0, f(6)))
    bx_ = self._round_field(r1_)
    bx_.pack(side="left")
    e_ = tk.Entry(bx_, textvariable=vars_["p1"], width=3, font=("Consolas", f(10), "bold"), bg=INPUT_BG, fg=TEXT_PRIMARY,
                  insertbackground=TEXT_PRIMARY, relief="flat", bd=0, highlightthickness=0, justify="center")
    e_.pack(padx=f(6), pady=f(3))
    widgets_lock.append(e_)
    pp_auto_l = tk.Label(r1_ if small else pr_r, text="", font=F_XS, bg=CARDC, fg=GRN, anchor="w")
    pp_hint_l = tk.Label(r1_ if small else pr_r, text="", font=(FONT_TH, f(8), "bold"), bg=CARDC, fg="#2DD4BF", anchor="w")
    if small:
        pp_hint_l.pack(side="right")
        pp_auto_l.pack(side="left", padx=(f(6), 0))
    else:
        pp_auto_l.pack(fill="x", pady=(f(3), 0))
        pp_hint_l.pack(fill="x")
    pq_ = tk.Frame(c2, bg=CARDC)
    pq_.pack(fill="x", padx=f(14), pady=(0, f(3)))
    pq_.columnconfigure((0, 1), weight=1, uniform="afq")
    q_btns, q_subs = {}, {}

    def set_q(k_):
        vars_["q"].set(k_)
        for kk, b_ in q_btns.items():
            b_.configure(bg=_lerp_hex(CARDC, "#2DD4BF", 0.55) if kk == k_ else CARD_BG_2,
                         fg="#FFFFFF" if kk == k_ else TEXT_SECONDARY)
        pp_refresh()
    for i_, k_ in enumerate(("standard", "hd")):
        q_ = PP_QUALITY[k_]
        b_ = self._pill_button(pq_, q_["label"], CARD_BG_2, TEXT_SECONDARY,
                               command=lambda k=k_: set_q(k), font=(FONT_TH, f(8), "bold"), padx=4, pady=f(4))
        b_.grid(row=0, column=i_, sticky="ew", padx=(0, 3) if i_ == 0 else (3, 0))
        q_btns[k_] = b_
        widgets_lock.append(b_)
        q_subs[k_] = tk.Label(pq_, text=f"≤{q_['limit_kb']:,} KB/240 จอ", font=(FONT_TH, max(7, f(7))), bg=CARDC, fg=DIMC)
        # [UI #pptx-gen3] ปุ่มขนาดรูปแสดงแค่ชื่อ Standard / Original (HD) — ไม่แสดงรายละเอียดใต้ปุ่ม
    pq_.pack_configure(pady=(0, f(8)))
    pp_est_l = tk.Label(c2, text="", font=F_XS, bg=CARDC, fg=DIMC, anchor="w")   # (ไม่ได้แสดง — ขนาดอยู่ในปุ่มแล้ว)
    pp_st = {"ana": None, "tpl": vars_["tpl"].get().strip(), "auto": bool(cfg0.get("start_auto", True)),
             "setting": False, "gen": 0, "n": None}

    def pp_n_screens():
        fa = getattr(self, "_smart_fetch_app", None)
        try:
            n_ = len(fa._get_screens()) if fa is not None else 0
        except Exception:
            n_ = 0
        return n_ or None

    def pp_refresh():
        ana = pp_st["ana"]
        n_ = pp_st["n"]
        try:
            self._pp_draw_strip(pp_strip, ana, n_, accent="#2DD4BF", summary=False)
        except Exception:
            pass
        if not ana:
            pp_auto_l.configure(text="เลือกเทมเพลตเพื่อวิเคราะห์หน้า" if not vars_["tpl"].get().strip() else "", fg=DIMC)
            pp_hint_l.configure(text="")
            pp_est_l.configure(text="")
            fit(o2, c2)
            return
        if not ana["end"]:
            pp_auto_l.configure(text="⚠ ไม่พบช่องรูป → 2 รูป/หน้า", fg=WARNING)
        else:
            pp_auto_l.configure(text=("✓ อัตโนมัติ" if small else "✓ อัตโนมัติ · แก้ได้") if pp_st["auto"] else "✎ กำหนดเอง",
                                fg=GRN if pp_st["auto"] else SUBC)
        budget = pp_image_budget(vars_["q"].get(), ana["size"])
        for kk, b_ in q_btns.items():        # เพดาน + ขนาดไฟล์ที่คาดไว้ ใต้ปุ่มแต่ละแบบ
            bud_ = pp_image_budget(kk, ana["size"])
            sub_ = f"≤{PP_QUALITY[kk]['limit_kb']:,} KB/240 จอ"
            if n_:
                sub_ = f"≤{PP_QUALITY[kk]['limit_kb']:,} KB · ~{(ana['size'] + n_ * (bud_ * 0.9 + 6000)) / 1048576:,.1f} MB"
            try:
                q_subs[kk].configure(text=sub_, fg=SUBC if kk == vars_["q"].get() else DIMC)
            except Exception:
                pass
        if n_:
            try:
                plan = af_ppt_plan(ana, n_, int(vars_["p1"].get() or 1))
            except ValueError:
                plan = []
            pp_hint_l.configure(text=f"{n_} จอ → {len(plan)} หน้า")
            est = (ana["size"] + n_ * (budget * 0.9 + 6000)) / 1048576
            pp_est_l.configure(text=f"เพดาน {PP_QUALITY[vars_['q'].get()]['limit_kb']:,} KB / 240 จอ · รูปละ ≤ {budget // 1024:,} KB · คาดว่า ~{est:,.1f} MB")
        else:
            pp_hint_l.configure(text=f"{ana['per_page']} ช่อง/หน้า" if ana["end"] else "")
            pp_est_l.configure(text=f"เพดาน {PP_QUALITY[vars_['q'].get()]['limit_kb']:,} KB / 240 จอ · รูปละ ≤ {budget // 1024:,} KB")
        fit(o2, c2)

    def pp_scan(*_a):
        tpl_ = vars_["tpl"].get().strip()
        if tpl_ != pp_st["tpl"]:
            pp_st["auto"] = True
            pp_st["tpl"] = tpl_
        pp_st["gen"] += 1
        g_ = pp_st["gen"]
        pp_st["n"] = pp_n_screens()
        if not tpl_ or not os.path.isfile(tpl_):
            pp_st["ana"] = None
            pp_refresh()
            return
        start_ = None
        if not pp_st["auto"]:
            try:
                start_ = max(1, int(vars_["p1"].get()))
            except ValueError:
                start_ = None
        order_ = getattr(self, "pptx_order_var", None)
        order_ = order_.get() if order_ is not None else "row"

        def work():
            try:
                a_ = pp_analyze_template(tpl_, start_, order_)
            except Exception:
                a_ = None

            def done():
                if g_ != pp_st["gen"] or not win.winfo_exists():
                    return
                pp_st["ana"] = a_
                if a_ and pp_st["auto"] and a_["auto_start"] and vars_["p1"].get() != str(a_["auto_start"]):
                    pp_st["setting"] = True
                    try:
                        vars_["p1"].set(str(a_["auto_start"]))
                    finally:
                        pp_st["setting"] = False
                pp_refresh()
            post(done)
        threading.Thread(target=work, daemon=True).start()

    def _p1_edit(*_a):
        if pp_st["setting"]:
            return
        pp_st["auto"] = False
        if pp_st.get("job"):
            try:
                win.after_cancel(pp_st["job"])
            except Exception:
                pass
        pp_st["job"] = win.after(400, pp_scan)
    vars_["p1"].trace_add("write", _p1_edit)
    vars_["tpl"].trace_add("write", lambda *_a: win.after(200, pp_scan))
    set_q(vars_["q"].get())

    # การ์ด ③ AI
    o3, c3 = gcard(setup, GOLD_)
    o3.pack(fill="x", pady=(f(8), 0))
    _qbox = {}

    def _q_right(h_):
        # [FEATURE #pptx-auto-pages] โควต้า + ผลทดสอบ ย้ายมาหัวการ์ด (เผื่อที่ให้การ์ด PowerPoint)
        qrow = tk.Frame(h_, bg=CARDC)
        qrow.pack(side="right")
        _qbox["lbl"] = tk.Label(qrow, text="", font=F_XS, bg=CARDC, fg=SUBC, anchor="e")
        _qbox["lbl"].pack(side="top", anchor="e")
        _qbox["bar"] = tk.Canvas(qrow, height=6, width=int(LW * 0.24), bg=CARDC, highlightthickness=0)
        _qbox["bar"].pack(side="top", anchor="e", pady=(2, 0))
    sec_head(c3, self._draw_sparkle_glyph, GOLD_, "AI · Gemini Vision", right=_q_right)
    ai_seg = tk.Frame(c3, bg=CARDC)
    ai_seg.pack(fill="x", padx=f(14), pady=(0, f(8)))
    ai_seg.columnconfigure((0, 1), weight=1, uniform="afai")
    ai_btns = {}
    for i_, (k_, t_) in enumerate(((True, "ใช้ AI เลือก + แต่งสี"), (False, "ไม่ใช้ AI"))):
        b_ = self._pill_button(ai_seg, t_, CARD_BG_2, TEXT_SECONDARY, command=lambda k=k_: toggle_ai(k),
                               font=(FONT_TH, f(8), "bold"), padx=8, pady=f(6))
        b_.grid(row=0, column=i_, sticky="ew", padx=(0, 4) if i_ == 0 else (4, 0))
        ai_btns[k_] = b_
        widgets_lock.append(b_)
    kr = tk.Frame(c3, bg=CARDC)
    kr.pack(fill="x", padx=f(14), pady=(0, f(6)))
    tk.Label(kr, text="API Key", font=(FONT_TH, f(8), "bold"), bg=CARDC, fg=SUBC).pack(side="left", padx=(0, f(6)))
    bxk = self._round_field(kr)
    bxk.pack(side="left", fill="x", expand=True)
    ek = tk.Entry(bxk, textvariable=vars_["key"], font=("Consolas", f(9)), bg=INPUT_BG, fg=TEXT_PRIMARY, show="•",
                  insertbackground=TEXT_PRIMARY, relief="flat", bd=0, highlightthickness=0)
    ek.pack(fill="x", padx=f(8), pady=f(5))
    widgets_lock.append(ek)
    ping_btn = self._pill_button(kr, "ทดสอบ", _lerp_hex(CARDC, GOLD_, 0.28), "#FFE08A",
                                 command=lambda: do_ping(), font=(FONT_TH, f(8), "bold"), padx=f(8), pady=f(4))
    ping_btn.pack(side="left", padx=(f(6), 0))
    widgets_lock.append(ping_btn)
    mr = tk.Frame(c3, bg=CARDC)
    mr.pack(fill="x", padx=f(14), pady=(0, f(8)))
    tk.Label(mr, text="Model", font=(FONT_TH, f(8), "bold"), bg=CARDC, fg=SUBC).pack(side="left", padx=(0, f(6)))
    bxm = self._round_field(mr)
    bxm.pack(side="left", fill="x", expand=True)
    em = tk.Entry(bxm, textvariable=vars_["model"], font=("Consolas", f(9)), bg=INPUT_BG, fg=TEXT_PRIMARY, width=10,
                  insertbackground=TEXT_PRIMARY, relief="flat", bd=0, highlightthickness=0)
    em.pack(fill="x", padx=f(8), pady=f(5))
    tk.Label(mr, text="มั่นใจ ≥", font=(FONT_TH, f(8), "bold"), bg=CARDC, fg=SUBC).pack(side="left", padx=(f(8), f(6)))
    bxt = self._round_field(mr)
    bxt.pack(side="left")
    et = tk.Entry(bxt, textvariable=vars_["thr"], width=4, font=("Consolas", f(9), "bold"), bg=INPUT_BG, fg=TEXT_PRIMARY,
                  insertbackground=TEXT_PRIMARY, relief="flat", bd=0, highlightthickness=0, justify="center")
    et.pack(padx=f(6), pady=f(5))
    widgets_lock.extend([em, et])
    mr.pack_configure(pady=(0, f(12)))
    q_lbl, q_bar = _qbox["lbl"], _qbox["bar"]
    _pmsg = {"id": None}

    class _PingMsg:
        """ผลทดสอบการเชื่อมต่อ แสดงแทนข้อความโควต้าชั่วคราว 5 วินาที"""
        def configure(self_, cnf=None, **kw):
            kw.update(cnf or {})
            try:
                q_lbl.configure(**kw)
                if _pmsg["id"]:
                    win.after_cancel(_pmsg["id"])
                _pmsg["id"] = win.after(5000, refresh_quota) if kw.get("text") != "กำลังทดสอบ…" else None
            except tk.TclError:
                pass
    ping_lbl = _PingMsg()
    start_btn = self._pill_button(setup, "▶   START AUTO FLOW", PINK, "#FFFFFF", command=lambda: do_start(),
                                  font=(FONT_TH, f(13), "bold"), padx=16, pady=f(12))
    start_btn.pack(side="bottom", fill="x", pady=(f(10), 0), before=o1)

    def quota_used(model):
        q = self.config_data.get("autoflow_quota") or {}
        if q.get("date") != af_quota_date():
            return 0
        return int((q.get("counts") or {}).get(model, 0))

    def quota_add(model, n=1):
        q = self.config_data.get("autoflow_quota") or {}
        if q.get("date") != af_quota_date():
            q = {"date": af_quota_date(), "counts": {}}
        q.setdefault("counts", {})[model] = int(q["counts"].get(model, 0)) + n
        self.config_data["autoflow_quota"] = q
        self._save_config()
        refresh_quota()

    def refresh_quota():
        m = vars_["model"].get().strip() or AF_DEFAULT_MODEL
        used, lim = quota_used(m), af_quota_limit(m)
        q_bar.delete("all")
        wq = int(q_bar["width"])
        q_bar.create_rectangle(0, 0, wq, 6, fill=INPUT_BG, width=0)
        if use_ai["on"]:
            q_lbl.configure(text=f"โควต้าวันนี้ {used}/{lim}", fg=WARNING if used < lim else ERROR)
            q_bar.create_rectangle(0, 0, int(wq * min(1, used / max(1, lim))), 6, fill=GOLD_ if used < lim else ERROR, width=0)
        else:
            q_lbl.configure(text="ใช้เกณฑ์ในเครื่อง", fg=DIMC)
        try:
            cnt_v["quota"].configure(text=f"{used}/{lim}" if use_ai["on"] else "—")
        except Exception:
            pass

    def toggle_ai(v=None):
        use_ai["on"] = (not use_ai["on"]) if v is None else v
        for kk, bb in ai_btns.items():
            on = kk == use_ai["on"]
            bb.configure(bg=(_lerp_hex(CARDC, GOLD_, 0.62) if kk else _lerp_hex(CARDC, "#8FA3BF", 0.38)) if on else CARD_BG_2,
                         fg=("#1F1400" if kk else "#FFFFFF") if on else TEXT_SECONDARY)
        refresh_quota()
        st["ai_state"] = "idle" if use_ai["on"] else "off"
        try:
            draw_chip()
        except Exception:
            pass

    def read_fetch(autofill=True):
        try:
            self._ensure_panel_built("fetch")
        except Exception:
            pass
        fa = getattr(self, "_smart_fetch_app", None)
        if fa is None:
            return "โหลดโหมด Fetch ไม่ได้"
        try:
            info = {"masters": fa._get_master_paths(), "output": fa.var_output.get().strip(),
                    "products": fa._get_products(), "screens": fa._get_screens(),
                    "busy": bool(getattr(fa, "_running", False))}
        except Exception as e:
            return f"อ่านค่า Fetch ไม่ได้: {e}"
        typed = [x.strip() for x in re.split(r"[,\n]+", vars_["prod"].get()) if x.strip()]
        if typed:
            info["products"] = list(dict.fromkeys(typed))      # ใช้สินค้าที่กรอกในหน้านี้
        elif info["products"] and autofill:
            vars_["prod"].set(", ".join(info["products"]))     # ยังไม่กรอก → ดึงค่าจากแท็บ Fetch มาใส่ให้
        errs_ = []
        miss = [n for n, v in (("Master Path", info["masters"]), ("Output", info["output"])) if not v]
        if miss:
            errs_.append("ยังไม่ได้ตั้งในแท็บ Fetch: " + ", ".join(miss))
        if not info["products"]:
            errs_.append("กรอกชื่อสินค้า / Playlist ในช่องด้านบน")
        if errs_:
            info["error"] = " · ".join(errs_)
        return info

    def short(t, n=30):
        return t if len(t) <= n else t[:n - 1] + "…"

    def refresh_nas(autofill=True):
        info = read_fetch(autofill)
        if isinstance(info, str):
            nas_err.configure(text="⚠ " + info)
            nas_err.pack(fill="x", padx=f(14), pady=(0, f(10)))
            fit(o1, c1)
            return info
        nas_vals["drv"].configure(text=f"{len(info['masters'])} ไดรฟ์")
        nas_vals["scr"].configure(text=f"{len(info['screens'])} จอ (Master File)" if info["screens"] else "ทุกจอใต้ Master Path")
        nas_vals["out"].configure(text=short("…" + info["output"][-40:] if len(info["output"]) > 40 else (info["output"] or "—"), int(LW / 9)))
        if info.get("error"):
            nas_err.configure(text="⚠ " + info["error"])
            nas_err.pack(fill="x", padx=f(14), pady=(0, f(10)))
        else:
            nas_err.pack_forget()
        fit(o1, c1)
        try:
            win.after(60, pp_scan)      # จำนวนจออาจเปลี่ยน → คำนวณหน้า/ขนาดไฟล์ใหม่
        except Exception:
            pass
        return info

    # ---------- ซ้าย: ระหว่างรัน ----------
    ro1, rc1 = gcard(runf, CY)
    ro1.pack(fill="x")
    rh = tk.Frame(rc1, bg=CARDC)
    rh.pack(fill="x", padx=f(14), pady=f(12))
    self._neon_icon(rh, f(30), CARDC, self._draw_folder_glyph, CY).pack(side="left")
    rhc = tk.Frame(rh, bg=CARDC)
    rhc.pack(side="left", padx=(f(10), 0))
    rn1 = tk.Label(rhc, text="", font=(FONT_TH, f(10), "bold"), bg=CARDC, fg="#FFFFFF", anchor="w")
    rn1.pack(anchor="w")
    rn2 = tk.Label(rhc, text="", font=F_XS, bg=CARDC, fg=SUBC, anchor="w")
    rn2.pack(anchor="w")
    ro2, rc2 = gcard(runf, PINK)
    ro2.pack(fill="x", pady=(f(10), 0))
    pt = tk.Frame(rc2, bg=CARDC)
    pt.pack(fill="x", padx=f(14), pady=(f(12), 0))
    p_big = tk.Label(pt, text="0%", font=("Consolas", f(26), "bold"), bg=CARDC, fg="#FFFFFF")
    p_big.pack(side="left")
    ptc = tk.Frame(pt, bg=CARDC)
    ptc.pack(side="left", padx=(f(12), 0), pady=(f(6), 0))
    p_l1 = tk.Label(ptc, text="เตรียมงาน…", font=(FONT_TH, f(10), "bold"), bg=CARDC, fg="#FFFFFF", anchor="w")
    p_l1.pack(anchor="w")
    p_l2 = tk.Label(ptc, text="", font=F_XS, bg=CARDC, fg=SUBC, anchor="w")
    p_l2.pack(anchor="w")
    pbar = tk.Canvas(rc2, height=f(10), bg=CARDC, highlightthickness=0)
    pbar.pack(fill="x", padx=f(14), pady=(f(8), f(14)))
    ro3, rc3 = gcard(runf, GOLD_)
    ro3.pack(fill="x", pady=(f(10), 0))
    cgrid = tk.Frame(rc3, bg=CARDC)
    cgrid.pack(fill="x", padx=f(12), pady=f(12))
    cgrid.columnconfigure((0, 1), weight=1, uniform="afc")
    cnt_v = {}
    for i_, (k_, t_, col_, gly) in enumerate((("auto", "AI เลือกเอง", GOLD_, self._draw_sparkle_glyph),
                                               ("human", "คนเลือก", ORG, self._draw_person_glyph),
                                               ("bad", "รูปเสียตัดออก", ERROR, self._draw_exclaim_glyph),
                                               ("quota", "โควต้า AI วันนี้", CY, self._draw_bolt_glyph))):
        tf = tk.Frame(cgrid, bg=_lerp_hex(CARDC, col_, 0.08), highlightthickness=1, highlightbackground=_lerp_hex(CARDC, col_, 0.35))
        tf.grid(row=i_ // 2, column=i_ % 2, sticky="nsew", padx=(0 if i_ % 2 == 0 else 5, 0 if i_ % 2 else 5), pady=(0 if i_ < 2 else 5, 0))
        tfi = tk.Frame(tf, bg=tf["bg"])
        tfi.pack(fill="x", padx=f(10), pady=f(8))
        self._neon_icon(tfi, f(22), tf["bg"], gly, col_).pack(side="left", anchor="n", pady=(f(4), 0))
        tfc = tk.Frame(tfi, bg=tf["bg"])
        tfc.pack(side="left", padx=(f(8), 0))
        cnt_v[k_] = tk.Label(tfc, text="0", font=("Consolas", f(17), "bold"), bg=tf["bg"], fg=col_, anchor="w")
        cnt_v[k_].pack(anchor="w")
        tk.Label(tfc, text=t_, font=F_XS, bg=tf["bg"], fg=SUBC, anchor="w").pack(anchor="w")
    ro4, rc4 = gcard(runf, VIO)
    ro4.pack(fill="x", pady=(f(10), 0))
    sec_head(rc4, self._draw_document_glyph, VIO, "PowerPoint ที่กำลังสร้าง")
    pp_row = tk.Frame(rc4, bg=CARDC)
    pp_row.pack(fill="x", padx=f(14))
    pp_txt = tk.Label(rc4, text="—", font=F_XS, bg=CARDC, fg=SUBC, anchor="w")
    pp_txt.pack(fill="x", padx=f(14), pady=(f(6), 0))
    pp_txt2 = tk.Label(rc4, text="", font=F_XS, bg=CARDC, fg="#5EEAD4", anchor="w")
    pp_txt2.pack(fill="x", padx=f(14), pady=(0, f(12)))
    ro5, rc5 = gcard(runf, CY, bg=CARDC, alpha=80)
    ro5.pack(fill="both", expand=True, pady=(f(10), 0))
    sec_head(rc5, self._draw_bolt_glyph, CY, "กิจกรรมล่าสุด")
    feed = tk.Frame(rc5, bg=CARDC)
    feed.pack(fill="both", expand=True, padx=f(14), pady=(0, f(10)))
    feed_rows = []
    rbot = tk.Frame(runf, bg=BG0)
    rbot.pack(side="bottom", fill="x", pady=(f(10), 0), before=ro1)
    ro5.configure(height=f(60))
    timer = tk.Label(rbot, text="0:00", font=("Consolas", f(15), "bold"), bg=_lerp_hex(BG0, PINK, 0.5), fg="#FFFFFF", padx=f(16), pady=f(7))
    timer.pack(side="left")
    stop_btn = self._pill_button(rbot, "■  หยุด", _lerp_hex(BG0, ERROR, 0.35), "#FFD0D0", command=lambda: do_stop(),
                                 font=(FONT_TH, f(10), "bold"), padx=f(16), pady=f(8))
    back_btn = self._pill_button(rbot, "◀  ตั้งค่า / รันใหม่", _lerp_hex(BG0, BRAND, 0.45), "#FFFFFF", command=lambda: show_setup(),
                                 font=(FONT_TH, f(10), "bold"), padx=f(14), pady=f(8))

    # ---------- ขวา: AI Stage ----------
    live = tk.Frame(sc, bg=SB)
    summ = tk.Frame(sc, bg=SB)
    sh = tk.Frame(live, bg=SB)
    sh.pack(fill="x", padx=14, pady=(f(12), f(10)))
    self._neon_icon(sh, f(38), SB, self._draw_sparkle_glyph, GOLD_).pack(side="left")
    shc = tk.Frame(sh, bg=SB)
    shc.pack(side="left", padx=(f(12), 0))
    s_title = tk.Label(shc, text="AI Stage", font=(FONT_TH, f(15), "bold"), bg=SB, fg="#FFFFFF", anchor="w")
    s_title.pack(anchor="w")
    s_sub = tk.Label(shc, text="", font=F_S, bg=SB, fg=SUBC, anchor="w")
    s_sub.pack(anchor="w")
    s_stat = tk.Label(sh, text="", font=(FONT_TH, f(9), "bold"), bg=_lerp_hex(SB, GOLD_, 0.18), fg="#FFE08A", padx=f(8), pady=f(4),
                      width=15)
    s_stat.pack(side="right", before=shc)
    CHW, CHH = f(300), f(50)
    ai_cv = tk.Label(sh, bg=SB, bd=0)
    ai_cv.pack(side="right", padx=(0, f(12)), before=shc)
    chip_f = (self._ss_pil_font(max(10, int(CHH * 0.24)), True), self._ss_pil_font(max(12, int(CHH * 0.34)), True))
    st.update({"ai_state": "idle", "ai_msg": "", "model": ""})

    def ai_chip(state, msg=""):
        st["ai_state"], st["ai_msg"] = state, msg
        draw_chip()

    def draw_chip():
        try:
            m = st.get("model") or (vars_["model"].get().strip() or AF_DEFAULT_MODEL)
            stt = st["ai_state"] if use_ai["on"] or st["running"] else "off"
            if st["running"] and not st.get("use_ai", True):
                stt = "off"
            ai_cv.configure(image=photo("aichip", af_ai_chip(stt, time.time(), m, st["ai_msg"], CHW, CHH,
                                                               chip_f[0], chip_f[1], bg=_af_rgb(SB))))
        except Exception:
            pass
    br = tk.Frame(live, bg=SB)
    br.pack(fill="x", padx=14)
    big = tk.Label(br, bg=SB, bd=0)
    big.pack(side="left", anchor="n")
    exp = tk.Frame(br, bg=SB, width=EXW, height=BH)
    exp.pack(side="left", fill="y", padx=(16, 0))
    exp.pack_propagate(False)
    cards = {}

    def exp_card(key, num, title, col, glyph):
        o_, c_ = gcard(exp, col, bg=_lerp_hex(SB, col, 0.06), radius=14, alpha=70)
        o_.pack(fill="x", pady=(0, f(8)))
        h_ = tk.Frame(c_, bg=c_["bg"])
        h_.pack(fill="x", padx=f(12), pady=(f(10), f(4)))
        nb = tk.Label(h_, text=num, font=(FONT_TH, f(9), "bold"), bg=col, fg="#12100A", width=2)
        nb.pack(side="left")
        tl = tk.Label(h_, text=title, font=(FONT_TH, f(10), "bold"), bg=c_["bg"], fg=col)
        tl.pack(side="left", padx=(f(8), 0))
        cards[key] = {"o": o_, "c": c_, "col": col, "title": tl, "nb": nb}
        return c_
    e2 = exp_card(2, "2", "ทำไมเลือกรูปนี้", GOLD_, None)
    e2_txt = tk.Label(e2, text="—", font=(FONT_TH, f(10)), bg=e2["bg"], fg="#F1F4FA", anchor="nw", justify="left", wraplength=EXW - 40, height=3)
    e2_txt.pack(fill="x", padx=f(12))
    e2r = tk.Frame(e2, bg=e2["bg"])
    e2r.pack(fill="x", padx=f(12), pady=(f(6), f(12)))
    tk.Label(e2r, text="มั่นใจ", font=F_XS, bg=e2["bg"], fg=SUBC).pack(side="left")
    cbar = tk.Canvas(e2r, height=f(8), bg=e2["bg"], highlightthickness=0)
    cbar.pack(side="left", fill="x", expand=True, padx=f(8))
    cpct = tk.Label(e2r, text="—", font=("Consolas", f(11), "bold"), bg=e2["bg"], fg=GOLD_, width=5, anchor="e")
    cpct.pack(side="left")
    e3 = exp_card(3, "3", "แต่งสี · ปุ่มเดียวกับทีม", PINK, None)
    e3_chips = tk.Frame(e3, bg=e3["bg"])
    e3_chips.pack(fill="x", padx=f(12))
    e3_txt = tk.Label(e3, text="—", font=F_S, bg=e3["bg"], fg=SUBC, anchor="nw", justify="left", wraplength=EXW - 40, height=3)
    e3_txt.pack(fill="x", padx=f(12), pady=(f(6), f(12)))

    e3_chips.configure(height=f(26))
    e3_chips.pack_propagate(False)
    for cd in cards.values():
        fit(cd["o"], cd["c"])
    bo, bc = gcard(live, LINE, bg=CARDC, radius=14, alpha=255)
    bo.pack(fill="x", padx=14, pady=(f(12), 0))
    bt_ = tk.Frame(bc, bg=CARDC)
    bt_.pack(fill="x", padx=12, pady=(f(8), 4))
    tk.Label(bt_, text="สายพานของจอนี้", font=(FONT_TH, f(9), "bold"), bg=CARDC, fg=SUBC).pack(side="left")
    belt_cv = tk.Canvas(bc, height=6, bg=CARDC, highlightthickness=0)
    belt_cv.pack(fill="x", padx=12 + TWt // 2, pady=(2, 4))
    bl = tk.Frame(bc, bg=CARDC)
    bl.pack(fill="x", padx=12, pady=(0, f(8)))
    stations = []
    for i_, (t_, d_) in enumerate(self.AF_STATIONS):
        cf = tk.Frame(bl, bg=CARDC, width=TWt, height=THt + f(52))
        cf.pack(side="left", anchor="n", padx=(0 if i_ == 0 else TGAP, 0))
        cf.pack_propagate(False)
        th_ = tk.Label(cf, bg=CARDC, bd=0)
        th_.pack()
        tr2 = tk.Frame(cf, bg=CARDC)
        tr2.pack(anchor="w", pady=(4, 0))
        ic_ = tk.Label(tr2, image=self._status_icon_photo("info", f(13), CARDC), bg=CARDC)
        ic_.pack(side="left")
        tl_ = tk.Label(tr2, text=t_, font=(FONT_TH, f(8), "bold"), bg=CARDC, fg=DIMC)
        tl_.pack(side="left", padx=(3, 0))
        dl_ = tk.Label(cf, text=d_, font=(FONT_TH, f(7)), bg=CARDC, fg=DIMC, anchor="w", justify="left", wraplength=TWt)
        dl_.pack(fill="x")
        stations.append({"th": th_, "ic": ic_, "tl": tl_, "dl": dl_, "d0": d_, "thumb": None, "state": "wait"})

    fit(bo, bc)
    strip_f = tk.Frame(live, bg=SB)
    if not small:
        strip_f.pack(fill="x", padx=14, pady=(f(10), 0))
    strip_t = tk.Label(strip_f, text="เสร็จแล้ว 0 จอ", font=(FONT_TH, f(8), "bold"), bg=SB, fg=SUBC, anchor="w")
    strip_t.pack(fill="x")
    strip_row = tk.Frame(strip_f, bg=SB, height=54)
    strip_row.pack(fill="x", pady=(4, 0))
    strip_row.pack_propagate(False)
    tick = tk.Frame(live, bg=CARDC, highlightthickness=1, highlightbackground=LINE)
    tick.pack(fill="x", padx=14, pady=(f(10), f(12)))
    tick_dot = tk.Label(tick, text="●", font=(FONT_TH, f(8)), bg=CARDC, fg=DIMC)
    tick_dot.pack(side="left", padx=(f(10), 4))
    tick_l = tk.Label(tick, text="พร้อม", font=F_S, bg=CARDC, fg=SUBC, anchor="w")
    tick_l.pack(side="left", fill="x", expand=True, pady=f(5))
    tick_b = tk.Label(tick, text="ดู Log ทั้งหมด ›", font=(FONT_TH, f(8), "bold"), bg=CARDC, fg="#9DA0FF", cursor="hand2")
    tick_b.pack(side="right", padx=f(10))
    tick_b.bind("<Button-1>", lambda e: open_log())

    # ---------- ตัวช่วยภาพ ----------
    def show_big(pil):
        big.configure(image=photo("big", af_rounded(pil, 16, _af_rgb(SB))))

    def idle_art():
        im = Image.new("RGB", (BW, BH), _af_rgb(CARDC))
        d = ImageDraw.Draw(im)
        for x in range(0, BW, 26):
            d.line([(x, 0), (x, BH)], fill=(28, 34, 52))
        for y in range(0, BH, 26):
            d.line([(0, y), (BW, y)], fill=(28, 34, 52))
        cx, cy = BW // 2, BH // 2
        for rr, a in ((70, 60), (48, 110), (26, 200)):
            d.ellipse([cx - rr, cy - rr, cx + rr, cy + rr], outline=(236, 72, 153), width=2)
        d.polygon([(cx - 8, cy - 12), (cx - 8, cy + 12), (cx + 14, cy)], fill=(255, 255, 255))
        show_big(im)

    def st_img(i, pulse=0.0):
        s_ = stations[i]
        base = s_["thumb"]
        im = af_cover(base, TWt, THt) if base is not None else Image.new("RGB", (TWt, THt), (14, 18, 30))
        if s_["state"] == "wait" and base is None:
            d = ImageDraw.Draw(im)
            d.text((TWt // 2 - 5, THt // 2 - 9), str(i + 1), font=pil_fs, fill=(60, 70, 96))
        col = {"run": (255, 196, 40), "done": (52, 211, 153), "warn": (251, 191, 36), "skip": (248, 113, 113)}.get(s_["state"])
        out = af_rounded(im, 10, _af_rgb(CARDC), border=(col + (int(150 + 105 * pulse) if s_["state"] == "run" else 200,)) if col else (60, 70, 96, 160), bw=2)
        return out

    def station(i, state, detail=None, thumb=None):
        s_ = stations[i]
        s_["state"] = state
        if thumb is not None:
            s_["thumb"] = thumb
        if state == "wait":
            s_["thumb"] = None
        kind = {"done": "ok", "run": "run", "warn": "warn", "skip": "err"}.get(state, "info")
        s_["ic"].configure(image=self._status_icon_photo(kind, f(13), CARDC))
        s_["tl"].configure(fg={"done": "#FFFFFF", "run": GOLD_, "warn": WARNING, "skip": ERROR}.get(state, DIMC))
        s_["dl"].configure(text=detail if detail is not None else s_["d0"],
                           fg={"run": "#FFE08A", "warn": WARNING, "skip": ERROR, "done": SUBC}.get(state, DIMC))
        s_["th"].configure(image=photo(f"st{i}", st_img(i)))
        if state == "run":
            st["active"] = i
            chip(i, "run")
            s_stat.configure(text=f"{i + 1}  {self.AF_STATIONS[i][0]}")
        elif st.get("active") == i:
            st["active"] = None
        if state in ("done", "skip", "warn") and state != "warn":
            chip(i, "done" if state == "done" else "skip")
        draw_belt()

    def draw_belt():
        belt_cv.update_idletasks()
        w_ = max(10, belt_cv.winfo_width())
        belt_cv.delete("all")
        belt_cv.create_line(0, 3, w_, 3, fill=LINE, width=2)
        n = len(stations)
        done_n = sum(1 for s_ in stations if s_["state"] in ("done", "skip"))
        act = st.get("active")
        prog_ = (done_n + (0.5 if act is not None else 0)) / max(1, n - 1)
        belt_cv.create_line(0, 3, int(w_ * min(1, prog_)), 3, fill=GOLD_, width=3)
        for i in range(n):
            x = int(w_ * i / max(1, n - 1))
            s_ = stations[i]["state"]
            col = {"done": GRN, "run": GOLD_, "warn": WARNING, "skip": ERROR}.get(s_, LINE)
            r = 4 if s_ != "run" else 5
            belt_cv.create_oval(x - r, 3 - r, x + r, 3 + r, fill=col, width=0)

    def belt_reset():
        for i in range(len(stations)):
            stations[i]["state"] = "wait"
            stations[i]["thumb"] = None
            station(i, "wait")
        for i in range(len(chips)):
            chip(i, "wait")

    def chip(i, state):
        lb = chips[i]
        if state == "run":
            lb.configure(bg=_lerp_hex(BG0, GOLD_, 0.55), fg="#1F1400")
        elif state == "done":
            lb.configure(bg=_lerp_hex(BG0, GRN, 0.25), fg="#BFF5DD")
        elif state == "skip":
            lb.configure(bg=_lerp_hex(BG0, ERROR, 0.25), fg="#FFD0D0")
        else:
            lb.configure(bg=CARDC, fg=DIMC)

    def active_card(k_):
        for kk, cd in cards.items():
            on = kk == k_
            cd["nb"].configure(bg=cd["col"] if on or kk == 0 else _lerp_hex(SB, cd["col"], 0.4))

    def typewrite(lbl, text, dur=0.7):
        n = max(1, len(text))
        steps = max(1, int(dur / 0.03))

        def stepf(i=1):
            if not lbl.winfo_exists():
                return
            lbl.configure(text=text[:int(n * i / steps)] + ("▍" if i < steps else ""))
            if i < steps:
                lbl.after(30, lambda: stepf(i + 1))
        stepf()

    def conf_bar(v, dur=0.8):
        cbar.update_idletasks()
        w_ = max(10, cbar.winfo_width())
        hh = int(cbar["height"])
        steps = max(1, int(dur / 0.03))

        def stepf(i=1):
            if not cbar.winfo_exists():
                return
            e_ = af_ease_out(i / steps)
            cbar.delete("all")
            cbar.create_rectangle(0, 0, w_, hh, fill=INPUT_BG, width=0)
            if v:
                col = GOLD_ if v >= 70 else WARNING
                cbar.create_rectangle(0, 0, int(w_ * v / 100 * e_), hh, fill=col, width=0)
                cpct.configure(text=f"{int(v * e_)}%")
            else:
                cpct.configure(text="—")
            if i < steps:
                cbar.after(30, lambda: stepf(i + 1))
        stepf()

    def card_pick(conf, reason, by):
        active_card(2)
        if by in ("AI", "AI (ใช้ผลเดิม)") and st.get("model"):
            by = f"{by} · {st['model']}"
        cards[2]["title"].configure(text=f"ทำไมเลือกรูปนี้ · {by}")
        typewrite(e2_txt, reason or "—")
        conf_bar(conf)

    def card_color(acts, reason):
        active_card(3)
        for ch in e3_chips.winfo_children():
            ch.destroy()
        if not acts:
            tk.Label(e3_chips, text="✓ รูปดีอยู่แล้ว ไม่ต้องแต่ง", font=(FONT_TH, f(9), "bold"), bg=e3["bg"], fg=GRN).pack(side="left")
        for j, a_ in enumerate(acts):
            lb = tk.Label(e3_chips, text=AF_COLOR_LABEL.get(a_, a_), font=(FONT_TH, f(9), "bold"),
                          bg=_lerp_hex(e3["bg"], PINK, 0.35), fg="#FFE3F1", padx=f(10), pady=2)
            lb.pack(side="left", padx=(0, f(6)))
        typewrite(e3_txt, reason or "—", 0.5)

    def card_crop(text):
        active_card(None)
        typewrite(s_sub, text, 0.6)

    def cards_reset():
        active_card(None)
        cards[2]["title"].configure(text="ทำไมเลือกรูปนี้")
        e2_txt.configure(text="—")
        conf_bar(None, 0.03)
        for ch in e3_chips.winfo_children():
            ch.destroy()
        e3_txt.configure(text="—")

    # ---------- ตัวเล่นภาพเคลื่อนไหว (main thread) ----------
    def anim_stop():
        j = st.get("anim")
        if j:
            try:
                win.after_cancel(j)
            except Exception:
                pass
        st["anim"] = None

    def play(frame_fn, dur, loop=False, fps=30):
        anim_stop()
        t0 = time.time()

        def tickf():
            if not win.winfo_exists():
                return
            t_ = time.time() - t0
            p_ = t_ if loop else min(1.0, t_ / max(0.01, dur))
            try:
                show_big(frame_fn(p_))
            except Exception:
                pass
            if loop or p_ < 1.0:
                st["anim"] = win.after(int(1000 / fps), tickf)
            else:
                st["anim"] = None
        tickf()

    def fx(kind, *a):
        if kind == "sweep":
            tiles, states, i = a
            ss = list(states)
            ss[i] = "scan"
            play(lambda p: af_grid2(tiles, ss, BW, BH, sweep=(i, p))[0], 0.32)
        elif kind == "grid":
            tiles, states = a
            anim_stop()
            show_big(af_grid2(tiles, states, BW, BH)[0])
        elif kind == "think":
            tiles, states = a
            st["think"] = True
            base_, lay_ = af_grid2(tiles, states, BW, BH)          # [FEATURE #af-aura] วาดกริดครั้งเดียว แล้ววาดแค่ขอบวิ่ง
            play(lambda t: af_think_aura(base_, lay_, states, t, BW, BH), 0, loop=True, fps=20)
        elif kind == "pick":
            tiles, states, idx, full = a
            g_, lay = af_grid2(tiles, states, BW, BH)
            rect = lay[idx] if idx < len(lay) else (0, 0, BW, BH)

            def fr(p):
                if p < 0.45:
                    return af_grid2(tiles, states, BW, BH, pulse=0.5 + 0.5 * math.sin(p / 0.45 * math.pi * 3))[0]
                return af_zoom2(g_, rect, full, (p - 0.45) / 0.55, BW, BH)
            play(fr, 1.3)
        elif kind == "wipe":
            b_, a_ = a
            b_ = af_cover(b_, BW, BH)
            a_ = af_cover(a_, BW, BH)
            play(lambda p: af_wipe2(b_, a_, p, BW, BH), 1.4)
        elif kind == "crop":
            img, box = a
            play(lambda p: af_crop2(img, box, p, BW, BH), 0.9)
        elif kind == "crop3":                      # [FEATURE #af-scene] กรอบทองกลางภาพ → เลื่อนให้จออยู่กลาง
            img, b0, b1 = a
            play(lambda p: af_crop3(img, b0, b1, p, BW, BH), 1.3)
        elif kind == "scenes":                     # [FEATURE #af-scene] กรอบแดงรอบจอ + ชื่อซีน ค่อยๆ ขึ้น / หาย
            tiles, states, marks, fade_in = a
            fnt_s = st.get("_scene_font")
            if fnt_s is None:
                fnt_s = st["_scene_font"] = self._ss_pil_font(max(11, int(BH / 26)), True)
            if fade_in:
                play(lambda p: af_grid_scenes(tiles, states, marks, BW, BH, af_ease(p), fnt_s), 0.45)
            else:
                play(lambda p: af_grid_scenes(tiles, states, marks, BW, BH, 1.0 - af_ease(p), fnt_s), 0.35)
        elif kind == "resize":
            img, s_, d_ = a
            play(lambda p: af_resize2(img, p, s_, d_, pil_f, BW, BH), 0.9)
        elif kind == "stamp":
            b_, s2 = a
            play(lambda p: af_stamp2(b_, s2, p, BW, BH), 0.9)
        elif kind == "rename":
            img, o_, n_ = a
            f_half = st.get("_rename_font")
            if f_half is None:                     # [FEATURE #af-rename-b] แถวเดียว ตัวใหญ่ขึ้น (เดิม BH/34)
                f_half = st["_rename_font"] = self._ss_pil_font(max(12, int(BH / 22)), True)
            play(lambda p: af_rename2(img, o_, n_, p, f_half, BW, BH), 2.4)
        elif kind == "fly":
            img, slide, tgt = a
            play(lambda p: af_fly2(img, p, slide, BW, BH, target=tgt), 0.95)
        elif kind == "static":
            anim_stop()
            show_big(a[0])

    def think_stop():
        st["think"] = False

    # หัวใจเต้น: ขอบสถานีที่กำลังทำ + จุดสถานะ + ข้อความ AI กำลังคิด
    def heartbeat():
        if not win.winfo_exists():
            return
        st["pulse"] = 0.5 + 0.5 * math.sin(time.time() * 5)
        i = st.get("active")
        if i is not None and st["running"]:
            stations[i]["th"].configure(image=photo(f"st{i}", st_img(i, st["pulse"])))
        if st["running"]:
            tick_dot.configure(fg=PINK if st["pulse"] > 0.5 else _lerp_hex(CARDC, PINK, 0.4))
            if st.get("think"):
                n_ = int(time.time() * 3) % 4
                s_stat.configure(text="AI กำลังวิเคราะห์" + "." * n_ + " " * (3 - n_))
            draw_pbar()
        if live.winfo_ismapped():
            draw_chip()
        win.after(60 if st.get("ai_state") == "think" else 90, heartbeat)

    prog_state = {"i": 0, "n": 0, "eta": None, "shown": 0.0}

    def draw_pbar():
        pbar.update_idletasks()
        w_ = max(10, pbar.winfo_width())
        hh = int(pbar["height"])
        n = prog_state["n"]
        target = prog_state["i"] / n if n else 0
        prog_state["shown"] += (target - prog_state["shown"]) * 0.25
        v = prog_state["shown"]
        key = (w_, round(v, 3), round(st["pulse"], 1))
        if pbar.__dict__.get("_k") == key:
            return
        pbar._k = key
        im = Image.new("RGB", (w_, hh), _af_rgb(INPUT_BG))
        fw = int(w_ * v)
        if fw > 2:
            g = Image.new("RGB", (fw, hh))
            a_, b_ = _af_rgb(PINK), _af_rgb(GOLD_)
            gd = ImageDraw.Draw(g)
            for x in range(fw):
                t_ = x / max(1, fw - 1)
                gd.line([(x, 0), (x, hh)], fill=tuple(int(a_[j] + (b_[j] - a_[j]) * t_) for j in range(3)))
            sx = int((time.time() * 180) % (fw + 80)) - 40
            for dx in range(-20, 21):
                xx = sx + dx
                if 0 <= xx < fw:
                    px = g.getpixel((xx, hh // 2))
                    al = 0.35 * (1 - abs(dx) / 20)
                    gd.line([(xx, 0), (xx, hh)], fill=tuple(int(px[j] + (255 - px[j]) * al) for j in range(3)))
            im.paste(g, (0, 0))
        im = af_rounded(im, hh // 2, _af_rgb(CARDC))
        pbar._ph = ImageTk.PhotoImage(im)
        pbar.delete("all")
        pbar.create_image(0, 0, anchor="nw", image=pbar._ph)
        p_big.configure(text=f"{int(round(v * 100))}%")

    def progress(i, n, eta=None, label=None):
        prog_state.update({"i": i, "n": n, "eta": eta})
        p_l1.configure(text=label or (f"จอ {i} / {n}" if n else "เตรียมงาน…"))
        p_l2.configure(text=(f"เหลือประมาณ {int(eta // 60)}:{int(eta % 60):02d} นาที" if eta else ("กำลังทำงาน" if n else "")))

    def count_to(lbl, target, dur=0.45):
        try:
            start = int(str(lbl.cget("text")).split("/")[0])
        except ValueError:
            start = 0
        if start == target:
            lbl.configure(text=str(target))
            return
        steps = max(1, int(dur / 0.03))

        def stepf(i=1):
            if not lbl.winfo_exists():
                return
            lbl.configure(text=str(int(start + (target - start) * af_ease_out(i / steps))))
            if i < steps:
                lbl.after(30, lambda: stepf(i + 1))
        stepf()

    def counters(d):
        for k_, v_ in d.items():
            if k_ in cnt_v:
                count_to(cnt_v[k_], int(v_))

    def ppt_update(pages, filled, cur=None):
        """[FEATURE #pptx-auto-pages] สไลด์ย่อ (สูงสุด 6 หน้า เลื่อนตามหน้าที่กำลังวาง)
        หน้าที่ก๊อปจากหน้าต้นแบบมีป้าย + มุมขวาบน"""
        for ch in pp_row.winfo_children():
            ch.destroy()
        dup = st.get("pp_dup") or set()
        total = sum(n_ for _p, n_ in pages)
        starts, k_ = [], 0
        for _p, n_ in pages:
            starts.append(k_)
            k_ += n_
        cur_i = 0
        for i_, s_ in enumerate(starts):
            if filled > s_:
                cur_i = i_
        nshow = max(1, min(6, len(pages)))
        first = max(0, min(cur_i - nshow + 2, len(pages) - nshow))
        sw2 = int(min((LW - 60 - 8 * (nshow - 1)) / nshow, (LW - 60) * 0.3))
        shh = min(int(sw2 * st.get("pp_ratio", 1.35)), f(84))
        for idx in range(first, first + nshow):
            pno, n_ = pages[idx]
            s_ = starts[idx]
            f_here = max(0, min(n_, filled - s_))
            hl = (filled - 1 - s_) if (cur and 0 <= filled - 1 - s_ < n_) else None
            im = af_slide_mini(f_here, n_, sw2, shh, hl=hl, cols=1 if n_ <= 3 else 2)
            if pno in dup:
                d_ = ImageDraw.Draw(im)
                r_ = max(5, sw2 // 8)
                d_.ellipse([sw2 - 2 * r_ - 1, 1, sw2 - 2, 2 * r_], fill=(45, 212, 191))
                d_.line([sw2 - r_ - 1, 4, sw2 - r_ - 1, 2 * r_ - 3], fill=(255, 255, 255), width=2)
                d_.line([sw2 - 2 * r_ + 2, r_, sw2 - 5, r_], fill=(255, 255, 255), width=2)
            im = af_rounded(im, 5, _af_rgb(CARDC))
            tk.Label(pp_row, image=photo(f"pp{idx - first}", im), bg=CARDC, bd=0).pack(side="left", padx=(0, 8))
        if first + nshow < len(pages):
            tk.Label(pp_row, text=f"+{len(pages) - first - nshow}", font=F_S, bg=CARDC, fg=DIMC).pack(side="left")
        n_dup = len([p_ for p_, _n in pages if p_ in dup])
        if pages:
            base = f"หน้า {pages[0][0]}" + (f" + ก๊อปเพิ่มเอง {n_dup} หน้า" if n_dup else
                                            (f"–{pages[-1][0]}" if len(pages) > 1 else ""))
            pp_txt.configure(text=f"{base} · วางแล้ว {min(filled, total)} / {total} ช่อง")
        else:
            pp_txt.configure(text="—")
        try:
            pp_txt2.configure(text="+ = หน้าที่โปรแกรมก๊อปจากหน้าต้นแบบ" if n_dup else "")
        except Exception:
            pass
        guard(pp_row)
        fit(ro4, rc4)

    def strip_add(pil, by, n, scene="", dup=False):
        """แถบ "เสร็จแล้ว N จอ": ภาพจอย่อ + ชื่อซีน (ซีนซ้ำ = ขอบส้ม)"""
        kids = strip_row.winfo_children()
        maxn = max(4, int((SWd - 40) / 94))
        if len(kids) >= maxn:
            kids[0].destroy()
        col = (255, 140, 66) if (dup or by == "คน") else (255, 184, 0)
        if scene:
            base = af_contain(pil, 86, 48, _af_rgb("#06080F"))
            dd = ImageDraw.Draw(base, "RGBA")
            fs_ = st.get("_strip_font")
            if fs_ is None:
                fs_ = st["_strip_font"] = self._ss_pil_font(10, True)
            lab = ("ซ้ำ · " if dup else "") + scene
            while lab and dd.textlength(lab, font=fs_) > 80:
                lab = lab[:-1]
            dd.rectangle([0, 34, 86, 48], fill=(8, 10, 20, 205))
            dd.text((43, 41), lab, font=fs_, fill=(255, 200, 150) if dup else (240, 244, 252), anchor="mm")
        else:
            base = af_cover(pil, 86, 48)
        im = af_rounded(base, 8, _af_rgb(SB), border=col + (255,), bw=2)
        lb = tk.Label(strip_row, image=photo(f"strip{n}", im), bg=SB, bd=0)
        lb.pack(side="left", padx=(0, 8))
        strip_t.configure(text=f"เสร็จแล้ว {n} จอ")
        guard(strip_row)

    def screen_start(i, n, name, nfiles):
        anim_stop()
        st["think"] = False
        s_title.configure(text=f"จอ {i + 1} / {n}  ·  {name}")
        s_sub.configure(text=f"{nfiles} รูปจาก NAS · ทำครบทุกขั้นก่อนไปจอถัดไป")
        belt_reset()
        cards_reset()
        s_stat.configure(text="เริ่มจอใหม่")
        if st.get("use_ai"):
            st["ai_state"], st["ai_msg"] = "idle", ""
        progress(i, n, prog_state.get("eta"), f"กำลังทำจอ {i + 1} / {n}")

    def stage_msg(title, sub):
        s_title.configure(text=title)
        s_sub.configure(text=sub)

    def status(text):
        s_stat.configure(text=text)

    feed_font = tkfont.Font(family=FONT_TH, size=f(8))
    feed_tfont = tkfont.Font(family="Consolas", size=f(8))

    def _fit_one_line(t_, width):
        if feed_font.measure(t_) <= width:
            return t_
        lo, hi = 0, len(t_)
        while lo < hi:
            mid = (lo + hi + 1) // 2
            if feed_font.measure(t_[:mid] + "…") <= width:
                lo = mid
            else:
                hi = mid - 1
        return t_[:lo].rstrip(" ·:") + "…"

    def feed_add(text, kind):
        """กิจกรรมล่าสุด — 1 งาน = 1 บรรทัด (ยาวเกินตัดท้ายด้วย …) รายการใหม่อยู่บนสุด"""
        col = {"ok": GRN, "warn": WARNING, "err": ERROR, "ai": "#FFD34D", "h": "#FFFFFF"}.get(kind, SUBC)
        kids = feed.pack_slaves()
        rw = tk.Frame(feed, bg=CARDC)
        rw._col = col
        if kids:
            rw.pack(fill="x", before=kids[0], pady=(0, 2))
        else:
            rw.pack(fill="x", pady=(0, 2))
        tm = tk.Label(rw, text=time.strftime("%H:%M:%S"), font=feed_tfont, bg=CARDC, fg=DIMC)
        tm.pack(side="left")
        tk.Label(rw, text="●", font=(FONT_TH, f(7)), bg=CARDC, fg=col).pack(side="left", padx=(f(6), 0))
        t_ = " ".join(text.split()).lstrip("✓⚠✕ ")
        fw = feed.winfo_width()
        avail = (fw if fw > 50 else LW - f(56)) - feed_tfont.measure("00:00:00") - f(28)
        lb = tk.Label(rw, text=_fit_one_line(t_, max(60, avail)), font=feed_font, bg=CARDC,
                      fg=_lerp_hex(CARDC, col, 0.0), anchor="w")
        lb.pack(side="left", fill="x", padx=(f(5), 0))

        def fade(i=1, n=8):       # ค่อยๆ เรืองขึ้น
            if lb.winfo_exists():
                lb.configure(fg=_lerp_hex(CARDC, col, i / n))
                if i < n:
                    lb.after(35, lambda: fade(i + 1))
        fade()
        feed.update_idletasks()
        fh = feed.winfo_height()
        row_h = rw.winfo_reqheight() + 2
        n_max = max(2, min(10, fh // row_h)) if fh > 20 else 6      # ไม่ให้แถวล่างโดนตัดครึ่ง
        for k_, old in enumerate(feed.pack_slaves()):     # รายการเก่าจางลง
            if k_ >= n_max:
                old.destroy()
            elif k_ >= 1:
                for ch in old.winfo_children()[2:]:
                    try:
                        ch.configure(fg=_lerp_hex(CARDC, getattr(old, "_col", SUBC), max(0.5, 1 - k_ * 0.07)))
                    except Exception:
                        pass
        guard(rw)

    def log(text, kind="m"):
        st["log"].append((time.strftime("%H:%M:%S"), text, kind))
        try:
            feed_add(text, kind)
        except Exception:
            pass
        tick_l.configure(text=text if len(text) <= 120 else text[:119] + "…",
                         fg={"ok": GRN, "warn": WARNING, "err": ERROR, "ai": "#FFD34D", "h": "#FFFFFF"}.get(kind, SUBC))

    def open_log():
        lw = tk.Toplevel(win)
        lw.title("AUTO FLOW · Log")
        lw.configure(bg=BG_MAIN)
        lw.transient(win)
        tx = tk.Text(lw, bg=INPUT_BG, fg=TEXT_SECONDARY, font=(FONT_TH, f(10)), width=110, height=30, relief="flat",
                     padx=12, pady=10, wrap="word")
        tx.pack(fill="both", expand=True, padx=10, pady=10)
        for tg, fg_ in (("ok", GRN), ("warn", WARNING), ("err", ERROR), ("ai", "#FFD34D"), ("h", "#FFFFFF"), ("m", SUBC)):
            tx.tag_configure(tg, foreground=fg_)
        for ts, t_, k_ in st["log"]:
            tx.insert("end", f"{ts}  {t_}\n", (k_,))
        tx.configure(state="disabled")
        tx.see("end")

    def post(fn, *a):
        try:
            win.after(0, lambda: win.winfo_exists() and fn(*a))
        except Exception:
            pass

    U = {"post": post, "station": station, "screen_start": screen_start, "card_pick": card_pick,
         "card_color": card_color, "card_crop": card_crop, "fx": fx, "think_stop": think_stop, "anim_stop": anim_stop,
         "ai_chip": ai_chip,
         "log": log, "counters": counters, "progress": progress, "ppt_update": ppt_update, "strip_add": strip_add,
         "stage_msg": stage_msg, "status": status, "quota_add": quota_add, "idle": lambda: idle_art(),
         "BW": BW, "BH": BH, "pil_f": pil_f, "pil_fs": pil_fs}

    # ---------- ตั้งค่า / เริ่ม / หยุด ----------
    def save_cfg():
        try:
            thr = int(float(vars_["thr"].get() or AF_DEFAULT_THRESHOLD))
        except ValueError:
            thr = AF_DEFAULT_THRESHOLD
        c_ = dict(self.config_data.get("autoflow", {}) or {})
        c_.update({"tpl": vars_["tpl"].get().strip(), "out": vars_["out"].get().strip(),
                   "start": vars_["p1"].get().strip(), "quality": vars_["q"].get(),
                   "start_auto": pp_st["auto"],
                   "api_key": vars_["key"].get().strip(), "model": vars_["model"].get().strip() or AF_DEFAULT_MODEL,
                   "model_v": AF_MODEL_CFG_VER,
                   "threshold": max(0, min(100, thr)), "use_ai": use_ai["on"],
                   "product": vars_["prod"].get().strip()})
        self.config_data["autoflow"] = c_
        self._save_config()
        return dict(c_)

    def do_ping():
        k_, m = vars_["key"].get().strip(), vars_["model"].get().strip() or AF_DEFAULT_MODEL
        if not k_:
            ping_lbl.configure(text="ใส่ API Key ก่อน", fg=WARNING)
            return
        ping_lbl.configure(text="กำลังทดสอบ…", fg=DIMC)

        def work():
            try:
                ok = af_gemini_ping(k_, m)
                post(ping_lbl.configure, {"text": "✓ เชื่อมต่อได้" if ok else "ตอบกลับผิดรูปแบบ", "fg": GRN if ok else WARNING})
            except AFQuotaDayError:
                post(ping_lbl.configure, {"text": "✕ โควต้าวันนี้หมด", "fg": ERROR})
            except Exception as e:
                post(ping_lbl.configure, {"text": f"✕ {str(e)[:40]}", "fg": ERROR})
            post(quota_add, m)
        threading.Thread(target=work, daemon=True).start()

    def show_setup():
        if st["running"]:
            return
        runf.pack_forget()
        setup.pack(fill="both", expand=True)
        summ.pack_forget()
        live.pack(fill="both", expand=True)
        belt_reset()
        cards_reset()
        idle_art()
        stage_msg("AI Stage", "กด START เพื่อดูทุกขั้นตอน ตั้งแต่ดึงรูปจาก NAS จนลงสไลด์")
        status("พร้อมเริ่ม")
        refresh_nas()
        for o_, c_ in ((o2, c2), (o3, c3)):
            fit(o_, c_)

    def show_run():
        setup.pack_forget()
        runf.pack(fill="both", expand=True)
        back_btn.pack_forget()
        stop_btn.configure(text="■  หยุด")
        stop_btn.pack(side="left", padx=(8, 0))
        summ.pack_forget()
        live.pack(fill="both", expand=True)
        for ch in list(strip_row.winfo_children()) + list(feed.winfo_children()):
            ch.destroy()
        strip_t.configure(text="เสร็จแล้ว 0 จอ")
        for k_ in ("auto", "human", "bad"):
            cnt_v[k_].configure(text="0")
        prog_state.update({"i": 0, "n": 0, "shown": 0.0})
        progress(0, 0)
        for o_, c_ in ((ro1, rc1), (ro2, rc2), (ro3, rc3), (ro4, rc4)):
            fit(o_, c_)

    def tick_timer():
        if not st["running"] or not win.winfo_exists():
            return
        t = int(time.time() - st["t0"])
        timer.configure(text=f"{t // 60}:{t % 60:02d}")
        win.after(500, tick_timer)

    def do_start():
        if st["running"]:
            return
        c = save_cfg()
        errs = []
        info = refresh_nas()
        if isinstance(info, str):
            errs.append(info)
        elif info.get("error"):
            errs.append(info["error"])
        elif info["busy"]:
            errs.append("แท็บ Fetch กำลังดึงรูปอยู่ รอให้เสร็จก่อน")
        if not c["tpl"] or not os.path.isfile(c["tpl"]):
            errs.append("เลือกเทมเพลต PowerPoint")
        if not (str(c["start"]).isdigit() and int(c["start"]) >= 1):
            errs.append("หน้าต้นแบบต้องเป็นตัวเลขตั้งแต่ 1 ขึ้นไป")
        if c["use_ai"] and not c["api_key"]:
            errs.append("ใส่ API Key หรือเลือก \"ไม่ใช้ AI\"")
        if errs:
            self._show_alert("ยังเริ่มไม่ได้", "\n".join(errs), kind="warning")
            return
        c["fetch"] = info
        if c["use_ai"]:
            used, lim = quota_used(c["model"]), af_quota_limit(c["model"])
            need = len(info["screens"]) if info["screens"] else None
            if need and need > lim - used:
                ok_ = self._show_confirm_dialog(
                    title="โควต้า AI อาจไม่พอ", subtitle=f"{c['model']} · วันนี้ใช้ไปแล้ว {used}/{lim} ครั้ง",
                    rows=[("warn", "จอที่จะให้ AI เลือก", f"{need} จอ"), ("info", "โควต้าที่เหลือ", f"{max(0, lim - used)} ครั้ง")],
                    note="จอที่เกินโควต้าจะสลับไปรุ่น Lite ก่อน หมดอีกค่อยใช้เกณฑ์ในเครื่อง (ใช้ผลเดิม ไม่นับโควต้า)",
                    confirm_text="เริ่มต่อ", cancel_text="ยกเลิก", accent=WARNING)
                if not ok_:
                    return
            c["quota_used"], c["quota_limit"] = used, lim
            c["fb_used"] = {m: quota_used(m) for m in AF_FALLBACK_MODELS}
        if not c["out"]:
            c["out"] = os.path.join(info["output"], "AutoFlow_" + datetime.datetime.now().strftime("%Y%m%d_%H%M"))
            vars_["out"].set(c["out"])
        try:
            c["baseline"] = float((self.config_data.get("autoflow") or {}).get("baseline", 3))
        except Exception:
            c["baseline"] = 3.0
        rn1.configure(text=f"ดึงจาก NAS · {len(info['masters'])} ไดรฟ์")
        rn2.configure(text=short(f"{str(len(info['screens'])) + ' จอ · ' if info['screens'] else ''}{', '.join(info['products'][:1])}", int(LW / 8)))
        st.update({"model": c["model"], "use_ai": c["use_ai"], "ai_state": "idle" if c["use_ai"] else "off", "ai_msg": ""})
        st.update({"running": True, "stop": False, "t0": time.time(), "log": []})
        show_run()
        refresh_quota()
        tick_timer()
        threading.Thread(target=self._af_worker, args=(c, U, st), daemon=True).start()

    def do_stop():
        if not st["running"]:
            return
        st["stop"] = True
        stop_btn.configure(text="กำลังหยุด…")
        log("กดหยุด — จะหยุดหลังจอที่กำลังทำ", "warn")
        fa = getattr(self, "_smart_fetch_app", None)
        try:
            if fa is not None and getattr(fa, "_running", False):
                fa._cancel()
        except Exception:
            pass

    def on_done(result):
        if prog_state["n"]:
            prog_state["shown"] = prog_state["i"] / prog_state["n"]
            draw_pbar()
        p_l1.configure(text="เสร็จแล้ว" if result.get("ok") else "หยุดแล้ว")
        p_l2.configure(text=f"{prog_state['i']} / {prog_state['n']} จอ")
        tick_dot.configure(fg=GRN if result.get("ok") else WARNING)
        st["running"] = False
        st["think"] = False
        anim_stop()
        st.update({k_: result.get(k_) for k_ in ("out", "ppt", "csv")})
        stop_btn.pack_forget()
        back_btn.pack(side="left", padx=(8, 0))
        if result.get("summary"):
            build_summary(result["summary"])
        try:
            self._play_notify_sound("success" if result.get("ok") else "error")
        except Exception:
            pass

    # ---------- หน้าสรุปผล ----------
    def build_summary(S_):
        for ch in summ.winfo_children():
            ch.destroy()
        live.pack_forget()
        summ.pack(fill="both", expand=True)
        HW, HH = SWd - 28, f(118)
        hero = tk.Canvas(summ, width=HW, height=HH, bg=SB, highlightthickness=0)
        hero.pack(padx=14, pady=(f(12), f(10)))
        gim = Image.new("RGB", (HW, HH))
        gd = ImageDraw.Draw(gim)
        a_, b_ = (58, 22, 60), (16, 46, 52)
        for x in range(HW):
            t_ = x / max(1, HW - 1)
            gd.line([(x, 0), (x, HH)], fill=tuple(int(a_[j] + (b_[j] - a_[j]) * t_) for j in range(3)))
        gd2 = ImageDraw.Draw(gim, "RGBA")
        ccx, ccy, rr = HW - f(70), HH // 2, int(HH * 0.32)
        okc = (52, 211, 153) if not S_.get("stopped") else (251, 191, 36)
        for g in range(8, 0, -1):
            gd2.ellipse([ccx - rr - g * 3, ccy - rr - g * 3, ccx + rr + g * 3, ccy + rr + g * 3], fill=okc + (6,))
        gd2.ellipse([ccx - rr, ccy - rr, ccx + rr, ccy + rr], fill=okc + (255,))
        if not S_.get("stopped"):
            gd2.line([(ccx - rr * 0.45, ccy + rr * 0.02), (ccx - rr * 0.1, ccy + rr * 0.36), (ccx + rr * 0.48, ccy - rr * 0.32)],
                     fill=(6, 30, 20, 255), width=max(4, rr // 5), joint="curve")
        else:
            gd2.rectangle([ccx - rr * 0.32, ccy - rr * 0.4, ccx - rr * 0.1, ccy + rr * 0.4], fill=(40, 26, 0, 255))
            gd2.rectangle([ccx + rr * 0.1, ccy - rr * 0.4, ccx + rr * 0.32, ccy + rr * 0.4], fill=(40, 26, 0, 255))
        hero._bg = ImageTk.PhotoImage(af_rounded(gim, 18, _af_rgb(SB), border=(255, 184, 0, 150)))
        hero.create_image(0, 0, anchor="nw", image=hero._bg)
        el = int(S_["elapsed"])
        ttl = "AUTO FLOW เสร็จแล้ว" if not S_.get("stopped") else "AUTO FLOW หยุดกลางทาง"
        hero.create_text(f(30), HH // 2 - f(16), text=ttl, anchor="w", font=(FONT_TH, f(20), "bold"), fill="#FFFFFF")
        hero.create_text(f(30), HH // 2 + f(20), anchor="w", font=(FONT_TH, f(10)), fill="#D6DCEA",
                         text=f"{S_['done']} จอ · {S_['imgs']} รูป · ใช้เวลา {el // 60}:{el % 60:02d} นาที · {S_.get('model') or 'ไม่ใช้ AI'}")
        if S_.get("scene_known"):       # [FEATURE #af-scene]
            hero.create_text(f(30), HH // 2 + f(44), anchor="w", font=(FONT_TH, f(10), "bold"),
                             fill="#FFD27A" if S_.get("scene_dup") else "#6EE7B7",
                             text=f"ซีนไม่ซ้ำ {S_['scene_new']} / {S_['scene_known']} จอ"
                                  + (f" · ซีนซ้ำ {S_['scene_dup']} จอ" if S_.get("scene_dup") else "")
                                  + (f" · ไม่ได้เช็กซีน {S_['done'] - S_['scene_known']} จอ" if S_['done'] > S_['scene_known'] else ""))
        # คอนเฟตติ
        parts = []
        import random as _rnd
        cols = ["#FFB800", "#EC4899", "#22D3EE", "#34D399", "#8B5CF6", "#FFFFFF"]
        for _ in range(46):
            x0 = _rnd.uniform(HW * 0.45, HW - 10)
            sz = _rnd.uniform(4, 8)
            it = hero.create_rectangle(x0, -10, x0 + sz, -10 + sz * 0.6, fill=_rnd.choice(cols), width=0)
            parts.append([it, _rnd.uniform(-1.2, 1.2), _rnd.uniform(1.6, 4.2)])
        t0c = time.time()

        def confetti():
            if not hero.winfo_exists():
                return
            alive = time.time() - t0c < 2.6
            for p_ in parts:
                hero.move(p_[0], p_[1], p_[2])
            if alive:
                hero.after(30, confetti)
            else:
                for p_ in parts:
                    hero.delete(p_[0])
        if not S_.get("stopped"):
            confetti()
        kp = tk.Frame(summ, bg=SB)
        kp.pack(fill="x", padx=14)
        kp.columnconfigure((0, 1, 2, 3, 4, 5), weight=1, uniform="afk")
        kvals = {}
        dn = max(1, S_["done"])

        def kpi(i, n, t, sub, col, gly):
            o_, c_ = gcard(kp, col, bg=_lerp_hex(SB, col, 0.07), radius=14, alpha=130)
            o_.grid(row=0, column=i, sticky="nsew", padx=(0 if i == 0 else 10, 0))
            r_ = tk.Frame(c_, bg=c_["bg"])
            r_.pack(fill="x", padx=f(14), pady=(f(12), 0))
            self._neon_icon(r_, f(26), c_["bg"], gly, col).pack(side="left")
            v_ = tk.Label(r_, text="0", font=("Consolas", f(22), "bold"), bg=c_["bg"], fg=col)
            v_.pack(side="left", padx=(f(10), 0))
            tk.Label(c_, text=t, font=(FONT_TH, f(10), "bold"), bg=c_["bg"], fg="#FFFFFF", anchor="w").pack(fill="x", padx=f(14), pady=(f(4), 0))
            s_ = tk.Label(c_, text=sub, font=F_XS, bg=c_["bg"], fg=SUBC, anchor="w")
            s_.pack(fill="x", padx=f(14), pady=(0, f(12)))
            kvals[i] = (v_, s_, o_, c_)
            return v_
        count_to(kpi(0, S_["auto"], "AI เลือกเอง", f"{100 * S_['auto'] // dn}% ไม่ต้องใช้คน", GOLD_, self._draw_sparkle_glyph), S_["auto"], 0.9)
        count_to(kpi(1, S_["human"], "คนเลือก", f"ตรงกับที่ AI เสนอ {S_['same']}" if S_["human"] else "ไม่ต้องใช้คนเลย", ORG, self._draw_person_glyph), S_["human"], 0.9)
        count_to(kpi(2, S_["bad"], "รูปเสียที่ตัดออก", f"เบลอ {S_['blur']} · มืด {S_['dark']} · จ้า {S_['bright']}", ERROR, self._draw_exclaim_glyph), S_["bad"], 0.9)
        kpi(3, 0, "ประหยัดเวลา", "", GRN, self._draw_clock_glyph if hasattr(self, "_draw_clock_glyph") else self._draw_calendar_glyph)
        # [FEATURE #pptx-auto-pages] หน้าที่ก๊อปเอง + ขนาดไฟล์ PPT จริง เทียบเพดาน Standard / HD
        TEAL_ = "#2DD4BF"
        v4 = kpi(4, 0, "หน้าที่ก๊อปเอง", f"จากหน้าต้นแบบ {S_.get('pp_start', '—')}", TEAL_, self._draw_document_glyph)
        v4.configure(text=f"+{S_.get('pp_added', 0)}")
        size_mb = S_.get("pp_size", 0) / 1048576
        over = S_.get("pp_size", 0) / 1024 > S_.get("pp_limit_kb", 1e12) * 1.02 and S_.get("done", 0) >= 20
        v5 = kpi(5, 0, "ไฟล์ PPT", f"{S_.get('pp_quality', '')} " + ("⚠ เกินเพดาน" if over else "✓"),
                 WARNING if over else CY, self._draw_folder_glyph)
        v5.configure(text=(f"{size_mb:,.1f}MB" if size_mb < 100 else f"{size_mb:,.0f}MB") if S_.get("pp_size") else "—",
                     font=("Consolas", f(15), "bold"))
        for i in range(6):
            fit(kvals[i][2], kvals[i][3])
        bl_ = tk.Frame(summ, bg=SB)
        bl_.pack(fill="x", padx=14, pady=(f(10), 0))
        tk.Label(bl_, text="เทียบแบบเดิม (คัด + แต่งสี + ครอป + ลงสไลด์ด้วยมือ)", font=F_S, bg=SB, fg=SUBC).pack(side="left")
        bvar = tk.StringVar(value=f"{S_.get('baseline', 3):g}")
        bb_ = self._round_field(bl_)
        bb_.pack(side="left", padx=8)
        be = tk.Entry(bb_, textvariable=bvar, width=4, font=("Consolas", f(10)), bg=INPUT_BG, fg=TEXT_PRIMARY,
                      insertbackground=TEXT_PRIMARY, relief="flat", bd=0, highlightthickness=0, justify="center")
        be.pack(padx=8, pady=4)
        tk.Label(bl_, text="นาที / จอ", font=F_S, bg=SB, fg=SUBC).pack(side="left")

        def upd_save(*_a):
            try:
                b = max(0.0, float(bvar.get()))
            except ValueError:
                return
            saved = b * S_["done"] - S_["elapsed"] / 60.0
            kvals[3][0].configure(text=f"~{saved:.0f}" if saved >= 1 else "—")
            kvals[3][1].configure(text=f"นาที · เดิม {b * S_['done']:.0f} → {S_['elapsed'] / 60:.1f}")
            S_["baseline"] = b
            S_["saved"] = saved
            self.config_data.setdefault("autoflow", {})["baseline"] = b
        bvar.trace_add("write", upd_save)
        upd_save()
        bt = tk.Frame(summ, bg=SB)
        bt.pack(side="bottom", fill="x", padx=14, pady=(f(8), f(14)))
        for text_, key_, solid in (("เปิด PowerPoint", "ppt", True), ("เปิด CSV", "csv", False), ("เปิดโฟลเดอร์", "out", False)):
            self._pill_button(bt, text_, BRAND if solid else _lerp_hex(SB, BRAND, 0.4), "#FFFFFF",
                              command=lambda k=key_: st.get(k) and self._open_path(st[k]),
                              font=(FONT_TH, f(10), "bold"), padx=f(16), pady=f(8)).pack(side="left", padx=(0, 10))
        self._pill_button(bt, "⬇  บันทึกภาพสรุป (PNG)", GRN, "#06140E", command=lambda: save_png(S_),
                          font=(FONT_TH, f(10), "bold"), padx=f(16), pady=f(8)).pack(side="right")
        gr = tk.Frame(summ, bg=SB)
        gr.pack(fill="both", expand=True, padx=14, pady=(f(12), 0))
        summ.update_idletasks()
        avail_h = max(120, H - 20 - 44 - f(74) - 14 - 18 - (HH + f(22)) - f(130) - f(40) - f(70))
        thumbs = S_["thumbs"]
        ncols = 5 if len(thumbs) > 8 else 4
        tw = int((SWd - 28 - (ncols - 1) * 12) / ncols)
        th = int(tw * 9 / 16)
        nrows = max(1, int(avail_h // (th + f(28))))
        show_n = min(len(thumbs), ncols * nrows)
        cells = []
        for i, (pil, cap, by) in enumerate(thumbs[:show_n]):
            cf = tk.Frame(gr, bg=SB)
            cf.grid(row=i // ncols, column=i % ncols, padx=(0, 12), pady=(0, f(8)), sticky="nw")
            col = (255, 184, 0, 255) if by != "คน" else (255, 140, 66, 255)
            im = af_rounded(af_cover(pil, tw, th), 10, _af_rgb(SB), border=col, bw=2)
            lb = tk.Label(cf, bg=SB, bd=0, image=photo(f"sumph{i}", Image.new("RGB", (tw, th), _af_rgb(SB))))
            lb.pack()
            rw = tk.Frame(cf, bg=SB)
            rw.pack(fill="x", pady=(4, 0))
            tk.Label(rw, text=short(cap, 22), font=(FONT_TH, f(8), "bold"), bg=SB, fg="#E6EAF4").pack(side="left")
            tk.Label(rw, text=f" {by} ", font=(FONT_TH, f(7), "bold"), bg=_lerp_hex(SB, "#FFB800" if by != "คน" else ORG, 0.4),
                     fg="#FFF1C8").pack(side="right")
            cells.append((lb, im, i))
        if len(thumbs) > show_n:
            tk.Label(gr, text=f"+ อีก {len(thumbs) - show_n} จอ · ดูทั้งหมดในโฟลเดอร์ 03_Final", font=F_S, bg=SB,
                     fg=DIMC).grid(row=nrows, column=0, columnspan=ncols, sticky="w")

        def reveal(j=0):
            if j < len(cells) and cells[j][0].winfo_exists():
                lb, im, i = cells[j]
                lb.configure(image=photo(f"sum{i}", im))
                win.after(70, lambda: reveal(j + 1))
        reveal()
        guard(summ)
        win._af_save_png = lambda: save_png(S_)

    def save_png(S_):
        try:
            Wp, Hp = 1600, 900
            im = Image.new("RGB", (Wp, Hp), (12, 15, 26))
            d = ImageDraw.Draw(im)
            Fp = lambda s, b=False: self._ss_pil_font(s, b)
            for x in range(Wp):
                t_ = x / (Wp - 1)
                d.line([(x, 0), (x, 150)], fill=(int(58 + (16 - 58) * t_), int(22 + (46 - 22) * t_), int(60 + (52 - 60) * t_)))
            d.text((60, 36), "AUTO FLOW · สรุปผล", font=Fp(44, True), fill=(255, 255, 255))
            el = int(S_["elapsed"])
            d.text((62, 100), f"{S_['done']} จอ · {S_['imgs']} รูป · ใช้เวลา {el // 60}:{el % 60:02d} นาที · "
                              f"{S_.get('model') or 'ไม่ใช้ AI'} · {datetime.datetime.now():%d/%m/%Y %H:%M}", font=Fp(20), fill=(214, 220, 234))
            if S_.get("scene_known"):
                d.text((62, 126), f"ซีนไม่ซ้ำ {S_['scene_new']} / {S_['scene_known']} จอ"
                       + (f" · ซีนซ้ำ {S_['scene_dup']} จอ" if S_.get("scene_dup") else "")
                       + (f" · ไม่ได้เช็กซีน {S_['done'] - S_['scene_known']} จอ" if S_['done'] > S_['scene_known'] else ""),
                       font=Fp(17, True), fill=(255, 210, 122) if S_.get("scene_dup") else (110, 231, 183))
            kp_ = [(str(S_["auto"]), "AI เลือกเอง", f"{100 * S_['auto'] // max(1, S_['done'])}% ไม่ต้องใช้คน", (255, 184, 0)),
                   (str(S_["human"]), "คนเลือก", f"ตรงกับที่ AI เสนอ {S_['same']}", (255, 140, 66)),
                   (str(S_["bad"]), "รูปเสียที่ตัดออก", f"เบลอ {S_['blur']} · มืด {S_['dark']}", (255, 90, 103)),
                   (f"~{max(0, S_.get('saved', 0)):.0f} นาที", "ประหยัดเวลา", f"เทียบแบบเดิม {S_.get('baseline', 3):g} นาที/จอ", (52, 211, 153))]
            for i, (n, t, sub, col) in enumerate(kp_):
                x = 60 + i * 380
                d.rounded_rectangle([x, 185, x + 350, 335], radius=20, fill=(18, 23, 38), outline=col, width=3)
                d.text((x + 26, 200), n, font=Fp(48, True), fill=col)
                d.text((x + 26, 264), t, font=Fp(22, True), fill=(255, 255, 255))
                d.text((x + 26, 298), sub, font=Fp(17), fill=(166, 176, 200))
            if S_.get("pp_size"):
                d.text((60, 343), f"PowerPoint · ก๊อปหน้าต้นแบบเพิ่มเอง {S_.get('pp_added', 0)} หน้า · "
                                  f"{S_['pp_size'] / 1048576:,.1f} MB ({S_.get('pp_quality', '')})",
                       font=Fp(17, True), fill=(45, 212, 191))
            for i, (pil, cap, by) in enumerate(S_["thumbs"][:12]):
                x, y = 60 + (i % 6) * 250, 375 + (i // 6) * 240
                t_ = af_rounded(af_cover(pil, 230, 130), 12, (12, 15, 26),
                                border=(255, 184, 0, 255) if by != "คน" else (255, 140, 66, 255), bw=3)
                im.paste(t_, (x, y))
                d.text((x, y + 140), f"{cap[:18]} · {by}", font=Fp(17), fill=(166, 176, 200))
            d.text((60, Hp - 46), "InspectPilot · AUTO FLOW (Demo) — NAS → AI เลือก + แต่งสี → ครอป → รีไซส์ → Stamp → ชื่อไฟล์ → PowerPoint",
                   font=Fp(16), fill=(108, 119, 145))
            pth = os.path.join(st["out"] or ".", "autoflow_summary.png")
            im.save(pth)
            st["png"] = pth
            self._open_path(pth)
        except Exception as e:
            self._show_toast("บันทึกภาพไม่สำเร็จ", str(e), kind="error")

    win._af_done = on_done
    win._af_start = do_start
    win._af_vars = vars_
    win._af_toggle_ai = toggle_ai
    win._af_state = st
    win._af_stop = lambda: do_stop()

    def close():
        if st["running"]:
            self._show_toast("AUTO FLOW กำลังทำงาน", "กดหยุดก่อน แล้วค่อยปิดหน้าต่างครับ", kind="warning")
            return
        save_cfg()
        anim_stop()
        self._af_win = None
        win.destroy()
    win.protocol("WM_DELETE_WINDOW", close)

    def guard(w_):
        for seq in ("<MouseWheel>", "<Shift-MouseWheel>", "<Button-4>", "<Button-5>"):
            w_.bind(seq, lambda e: "break")
        for ch in w_.winfo_children():
            guard(ch)

    toggle_ai(use_ai["on"])
    show_setup()
    # ปรับขนาดภาพบนเวทีให้พอดีพื้นที่จริง (วัดหลังจัดวาง) — ไม่มีช่องว่างเหลือ/ไม่ล้น
    win.update_idletasks()
    try:
        main_h = H - 18 - 32 - head.winfo_reqheight() - 3 - f(10) - f(12)
        sc_h = main_h - 2 * (int(math.ceil(18 * 0.4)) + 1)
        slack = sc_h - live.winfo_reqheight() - 6
        if small and slack > strip_f.winfo_reqheight() + f(10) + 90:
            strip_f.pack(fill="x", padx=14, pady=(f(10), 0), before=tick)
            win.update_idletasks()
            slack = sc_h - live.winfo_reqheight() - 6
        if abs(slack) > 4:
            nBW = min(int((BH + slack) * 16 / 9), int(SWd * 0.66))
            BW, BH = nBW, int(nBW * 9 / 16)
            EXW = SWd - 28 - BW - 16
            AF_STAGE_W, AF_STAGE_H = BW, BH
            U["BW"], U["BH"] = BW, BH
            pil_f = self._ss_pil_font(max(12, int(BH / 17)), True)
            pil_fs = self._ss_pil_font(max(11, int(BH / 21)), True)
            U["pil_f"], U["pil_fs"] = pil_f, pil_fs
            exp.configure(width=EXW, height=BH)
            for lb_ in (e2_txt, e3_txt):
                lb_.configure(wraplength=EXW - 40)
            for cd in cards.values():
                fit(cd["o"], cd["c"])
            idle_art()
    except Exception as e:
        logging.warning("AUTO FLOW relayout: %s", e)
    guard(win)
    win.update_idletasks()
    x = max(0, (sw_ - (W + 20)) // 2)
    y = max(0, (sh_ - (H + 20)) // 2 - 22)
    win.geometry(f"+{x}+{y}")
    win.deiconify()
    win.focus_set()
    heartbeat()

def _af_worker(self, c, U, st):
    post = U["post"]

    def call(name, *a):
        post(U[name], *a)

    def L(text, kind="m"):
        call("log", text, kind)

    t_all = time.time()
    rows, result = [], {"out": c["out"], "ppt": None, "csv": None, "ok": False}
    fnt = U["pil_f"]
    BW, BH = U["BW"], U["BH"]

    def fx(kind, *a, wait=0.0):
        call("fx", kind, *a)
        if wait:
            time.sleep(wait)
    stopped = False
    try:
        out = c["out"]
        d_raw, d_stamp, d_final = (os.path.join(out, n) for n in ("01_นำเข้ารูป", "02_แต่งสี_Stamp", "03_Final"))
        for d in (d_raw, d_stamp, d_final):
            os.makedirs(d, exist_ok=True)
        L(f"เริ่ม AUTO FLOW · {('AI ' + c['model']) if c['use_ai'] else 'ไม่ใช้ AI'} · มั่นใจขั้นต่ำ {c['threshold']}%", "h")
        # ── ① ดึงจาก NAS (ระบบ Fetch เดิม ดึงทุกจอพร้อมกัน) ──
        info = c["fetch"]
        fa = self._smart_fetch_app
        call("stage_msg", "① กำลังดึงรูปจาก NAS…", "ใช้ระบบ Fetch เดิม · ดึงทุกจอพร้อมกันก่อน แล้วทำทีละจอ")
        call("idle")
        call("station", 0, "run", "เชื่อมต่อไดรฟ์…")
        screens_list = info["screens"]
        if not screens_list:
            call("station", 0, "run", "หารายชื่อจอ…")
            screens_list = fa._get_screens_from_master_paths(info["masters"], info["products"])
        if not screens_list:
            raise RuntimeError("ไม่พบโฟลเดอร์จอใต้ Master Path")
        L(f"① ดึงจาก NAS · {len(screens_list)} จอ · {len(info['products'])} สินค้า", "h")
        prod_dirs = [os.path.join(info["output"], p_) for p_ in info["products"]]
        th = threading.Thread(target=fa._logic, args=(list(screens_list), list(info["products"]),
                                                      list(info["masters"]), info["output"], False), daemon=True)
        th.start()
        t_f = time.time()
        last_n = -1
        while th.is_alive():
            th.join(0.8)
            n_now, newest = 0, None
            for pd_ in prod_dirs:
                try:
                    fs_ = [os.path.join(pd_, f_) for f_ in os.listdir(pd_) if f_.lower().endswith(SUPPORTED_FORMATS)]
                except OSError:
                    fs_ = []
                n_now += len(fs_)
                if fs_ and n_now != last_n:
                    newest = max(fs_, key=lambda p: os.path.getmtime(p) if os.path.exists(p) else 0)
            call("station", 0, "run", f"{n_now} รูป · {time.time() - t_f:.0f}s")
            if newest and n_now != last_n:
                try:
                    fx("static", af_caption2(af_open_small(newest, 900), f"NAS  →  {n_now} รูป", fnt, BW, BH, col=(8, 60, 80, 225)))
                except Exception:
                    pass
            last_n = n_now
        if getattr(fa, "_cancel_event", None) is not None and fa._cancel_event.is_set():
            stopped = True
            raise RuntimeError("หยุดระหว่างดึงรูป")
        groups = af_collect_groups(prod_dirs, screens_list)
        n_img = 0
        for screen, files in groups:
            dst = os.path.join(d_raw, screen)
            os.makedirs(dst, exist_ok=True)
            cp = []
            for f in files:
                p = os.path.join(dst, os.path.basename(f))
                try:
                    same = os.path.exists(p) and os.path.getsize(p) == os.path.getsize(f)
                except OSError:
                    same = False
                if not same:
                    if os.path.exists(p):
                        p = safe_dest_path(dst, os.path.basename(f))
                    shutil.copy2(f, p)
                cp.append(p)      # รันซ้ำโฟลเดอร์เดิม → ใช้ไฟล์เดิม (ชื่อคงที่ ใช้ผล AI เดิมได้)
            n_img += len(cp)
            rows.append({"screen": screen, "files": cp, "q": [], "ai": None, "final": None, "by": "", "note": "",
                         "acts": [], "creason": ""})
        have = {r["screen"].upper() for r in rows}
        for sc_ in screens_list:
            if sc_.upper() not in have:
                rows.append({"screen": sc_, "files": [], "q": [], "ai": None, "final": None, "by": "ไม่มีรูป",
                             "note": "ไม่พบรูปบน NAS", "acts": [], "creason": ""})
        rows.sort(key=lambda r: natural_name_key(r["screen"]))
        n_none = sum(1 for r in rows if not r["files"])
        fetch_s = time.time() - t_f
        L(f"① ดึงแล้ว {n_img} รูป · {len(rows) - n_none} จอ ({fetch_s:.0f}s)" + (f" · ไม่พบรูป {n_none} จอ" if n_none else ""),
          "warn" if n_none else "ok")
        todo = [r for r in rows if r["files"]]
        if not todo:
            raise RuntimeError("ไม่พบรูปของจอใดเลยบน NAS")
        # แผนช่องใน PowerPoint (ลำดับเดียวกับตอนสร้างจริง)
        order = getattr(self, "pptx_order_var", None)
        order = order.get() if order is not None else "row"
        # [FEATURE #pptx-auto-pages] แผนหน้าจากตัววิเคราะห์เดียวกับ Generate (รวมหน้าที่จะก๊อปเพิ่ม)
        start_auto = bool(c.get("start_auto", True))
        ana = pp_analyze_template(c["tpl"], None if start_auto else int(c["start"]), order)
        plan = af_ppt_plan(ana, len(todo), int(c["start"]))
        pages = [(pno, n_) for pno, n_, _dup in plan]
        dup_pages = {pno for pno, _n, dup in plan if dup}
        st["pp_dup"] = dup_pages
        st["pp_ratio"] = ana["slide_h"] / max(1, ana["slide_w"])
        if ana["end"]:
            L(f"⑧ เทมเพลต: หน้าต้นแบบ {ana['start']}" + (f"–{ana['end']}" if ana["end"] != ana["start"] else "")
              + f" · {ana['per_page']} ช่อง/หน้า · {len(todo)} จอ → {len(plan)} หน้า"
              + (f" (ก๊อปเพิ่ม {len(dup_pages)})" if dup_pages else ""), "ok")
        else:
            L(f"⑧ เทมเพลตไม่มีช่องรูป → 2 รูป/หน้า เริ่มหน้า {c['start']}", "warn")
        slot_plan = [(pno, k + 1) for pno, n_ in pages for k in range(n_)]
        call("ppt_update", pages, 0)
        cache = af_cache_load()
        used_sc = []            # [FEATURE #af-scene] ซีนที่จอก่อนหน้าเลือกไปแล้ว [{name, screen, thumb, b64}]
        q_used = int(c.get("quota_used", 0))
        q_lim = int(c.get("quota_limit", af_quota_limit(c["model"])))
        # [FEATURE #af-fallback] รุ่นสำรอง (Lite) — ใช้เมื่อรุ่นหลักยุ่ง (503) หรือโควต้าวันนี้หมด
        fb = {"list": [m for m in AF_FALLBACK_MODELS if m != c["model"]],
              "used": {m: int(v) for m, v in (c.get("fb_used") or {}).items()}, "main_dead": False, "said": False}

        def ai_avail():
            return (not fb["main_dead"] and q_used < q_lim) or any(
                fb["used"].get(m, 0) < af_quota_limit(m) for m in fb["list"])

        def ai_pick(paths_):
            """เรียก AI · รุ่นหลักยุ่ง/โควต้าหมด → สลับรุ่น Lite ให้เอง · คืน (ผล, รุ่นที่ตอบ)"""
            nonlocal q_used
            last_e = None
            if not fb["main_dead"] and q_used < q_lim:
                try:
                    res_ = af_gemini_pick(c["api_key"], c["model"], paths_, used_sc)
                    q_used += 1
                    call("quota_add", c["model"])
                    return res_, c["model"]
                except AFQuotaDayError as e:
                    fb["main_dead"] = True
                    call("quota_add", c["model"], max(0, q_lim - q_used))
                    q_used = q_lim
                    L(f"⚠ โควต้า {c['model']} วันนี้หมดแล้ว → สลับไปรุ่น Lite", "warn")
                    last_e = e
                except AFBusyError as e:
                    q_used += 1
                    call("quota_add", c["model"])
                    L(f"   {r['screen']}: {c['model']} {str(e)[:40]} → ลองรุ่น Lite", "warn")
                    last_e = e
                except Exception:
                    q_used += 1
                    call("quota_add", c["model"])
                    raise
            for m_ in list(fb["list"]):
                if fb["used"].get(m_, 0) >= af_quota_limit(m_):
                    continue
                call("ai_chip", "think", f"สลับไป {m_}")
                try:
                    res_ = af_gemini_pick(c["api_key"], m_, paths_, used_sc)
                except AFModelMissing as e:
                    fb["list"].remove(m_)
                    last_e = last_e or e
                    continue
                except AFQuotaDayError as e:
                    fb["used"][m_] = af_quota_limit(m_)
                    last_e = e
                    continue
                except AFBusyError as e:
                    fb["used"][m_] = fb["used"].get(m_, 0) + 1
                    call("quota_add", m_)
                    last_e = e
                    continue
                fb["used"][m_] = fb["used"].get(m_, 0) + 1
                call("quota_add", m_)
                if not fb["said"]:
                    fb["said"] = True
                    L(f"   ใช้รุ่นสำรอง {m_} แทน", "warn")
                return res_, m_
            if isinstance(last_e, (AFQuotaDayError, AFModelMissing)) or last_e is None:
                raise AFQuotaDayError("โควต้า AI วันนี้หมดทุกรุ่น")
            raise last_e
        cnt = {"auto": 0, "human": 0, "bad": 0}
        done_n = 0
        t_loop = time.time()
        final_paths = []
        used_names = set()      # รันซ้ำโฟลเดอร์เดิม → เขียนทับไฟล์ของรอบก่อน ชนกันเองในรอบนี้เท่านั้นถึงเติมเลข

        def dest(folder, name):
            p_ = os.path.join(folder, name)
            if p_.lower() in used_names:
                p_ = safe_dest_path(folder, name)
            used_names.add(p_.lower())
            return p_
        for si, r in enumerate(todo):
            if st.get("stop"):
                stopped = True
                break
            call("screen_start", si, len(todo), r["screen"], len(r["files"]))
            call("station", 0, "done", f"{len(r['files'])} รูป · {fetch_s:.0f}s")
            smalls = []
            for f in r["files"][:AF_MAX_AI_IMAGES]:
                try:
                    smalls.append(af_open_small(f, 640))
                except Exception:
                    smalls.append(Image.new("RGB", (160, 90), (40, 40, 50)))
            call("station", 0, "done", f"{len(r['files'])} รูป · {fetch_s:.0f}s", af_grid2(smalls, ["ok"] * len(smalls), 320, 180)[0])
            fx("grid", smalls, ["wait"] * len(smalls), wait=0.5)
            # ── ② ตรวจ + เลือก ──
            states = ["wait"] * len(smalls)
            call("station", 1, "run", "ตรวจเบลอ/มืด…")
            for k, f in enumerate(r["files"]):
                if k < len(smalls):
                    fx("sweep", smalls, list(states), k, wait=0.36)
                try:
                    q = af_quality(f)
                except Exception as e:
                    q = {"sharp": 0, "bright": 0, "flags": [f"เปิดไม่ได้"], "ok": False}
                r["q"].append(q)
                if k < len(smalls):
                    states[k] = "ok" if q["ok"] else "bad"
                fx("grid", smalls, list(states))
            bad_n = sum(1 for q in r["q"] if not q["ok"])
            cnt["bad"] += bad_n
            call("counters", {"bad": cnt["bad"]})
            passed = [(i, q) for i, q in enumerate(r["q"]) if q["ok"]]
            if bad_n:
                L(f"   {r['screen']}: ตัดออก " + ", ".join(f"({k + 1}) {'/'.join(q['flags'])}" for k, q in enumerate(r["q"]) if not q["ok"]), "warn")
            ai_acts, ai_creason = None, ""
            r["scn"], r["scene_note"] = {}, ""

            def _apply_scene(pp_, infos_, ai_best_):
                """ผล AI ต่อรูป → r["scn"] (index ของ r["files"]) + กฎไม่ให้ซีนซ้ำ · คืน (index ใน pp_, forced, overridden)"""
                for j_, inf_ in enumerate(infos_):
                    if inf_:
                        r["scn"][pp_[j_][0]] = inf_
                if not any(infos_):
                    return ai_best_, False, False
                return af_scene_pick(infos_, ai_best_, {j + 1: u.get("count", 1) for j, u in enumerate(used_sc)})
            if not passed:
                r["note"] = "รูปไม่ผ่านเกณฑ์ทั้งหมด"
            elif len(passed) == 1 and not c["use_ai"]:
                r["ai"] = (passed[0][0], 100, "ผ่านเกณฑ์รูปเดียว", "เกณฑ์")
            elif c["use_ai"]:
                pp = passed[:AF_MAX_AI_IMAGES]
                paths = [r["files"][i] for i, _q in pp]
                names = [os.path.basename(p) for p in paths]
                ck = af_cache_key(c["model"], paths)
                hit = cache.get(ck)
                if hit and hit.get("best_name") in names and (hit.get("images") is not None or len(pp) == 1):
                    # ใช้ผลเดิม: ชื่อซีน/กรอบจอ/คะแนนจากรอบก่อน แล้วเช็กซีนซ้ำกับงานนี้ใหม่ในเครื่อง (ไม่เสียโควต้า)
                    sc_names = [u["name"] for u in used_sc]
                    infos = []
                    for nm in names:
                        inf = dict((hit.get("images") or {}).get(nm) or {})
                        if inf:
                            inf["box"] = tuple(inf["box"]) if inf.get("box") else None
                            inf["dup"] = (sc_names.index(inf["scene"]) + 1) if inf.get("scene") in sc_names else 0
                        infos.append(inf or None)
                    bi = names.index(hit["best_name"])
                    ai_bi = bi
                    bi, forced, over = _apply_scene(pp, infos, bi)
                    r["scene_forced"] = forced
                    reason, fixed_ = af_reason_fix(hit.get("reason_raw", hit.get("reason", "")), bi, ai_bi,
                                                   infos[bi] if bi < len(infos) else None, over)
                    r["reason_fixed"] = fixed_
                    r["ai"] = (pp[bi][0], 100 if len(pp) == 1 else int(hit.get("conf", 0)), reason, "ใช้ผลเดิม")
                    r["scene_note"] = hit.get("note", "")
                    ai_acts, ai_creason = list(hit.get("acts") or []), hit.get("creason", "")
                    r["cache"] = True
                    call("ai_chip", "cache", "✓ ใช้ผลเดิม · ไม่เสียโควต้า")
                elif not ai_avail():
                    if len(passed) == 1:
                        r["ai"] = (passed[0][0], 100, "ผ่านเกณฑ์รูปเดียว · โควต้า AI หมด (ไม่ได้เช็กซีน)", "เกณฑ์")
                    else:
                        bi, conf, reason = af_local_pick(passed)
                        r["ai"] = (bi, conf, reason + " · โควต้า AI หมด (ไม่ได้เช็กซีน)", "เกณฑ์ในเครื่อง")
                    call("ai_chip", "quota", "โควต้าหมด · ใช้เกณฑ์ในเครื่อง")
                else:
                    call("station", 1, "run", "AI หากรอบจอ + ดูซีน…")
                    fx("think", smalls, list(states))
                    call("ai_chip", "think", "")
                    t_ai = time.time()
                    m_used = c["model"]
                    try:
                        (bi, conf, reason_raw, acts, creason, infos, note), m_used = ai_pick(paths)
                        ai_bi = bi
                        bi, forced, over = _apply_scene(pp, infos, bi)
                        r["scene_forced"] = forced
                        reason, fixed_ = af_reason_fix(reason_raw, bi, ai_bi, infos[bi] if bi < len(infos) else None, over)
                        r["reason_fixed"] = fixed_
                        by_ = "AI" if m_used == c["model"] else "AI สำรอง"
                        r["ai"] = (pp[bi][0], 100 if len(pp) == 1 else conf, reason, by_)
                        r["ai_model"] = m_used
                        r["scene_note"] = note
                        ai_acts, ai_creason = acts, creason
                        cache[ck] = {"best_name": names[ai_bi], "conf": conf, "reason": reason, "reason_raw": reason_raw,
                                     "acts": acts,
                                     "creason": creason, "note": note, "model": m_used,
                                     "images": {nm: {"box": list(inf["box"]) if inf.get("box") else None,
                                                     "scene": inf.get("scene", ""), "noprod": inf.get("noprod", False),
                                                     "score": inf.get("score", 0)}
                                                for nm, inf in zip(names, infos) if inf}}
                        af_cache_save(cache)
                        tag_ = "" if m_used == c["model"] else " · รุ่นสำรอง"
                        call("ai_chip", "done", f"✓ ตอบใน {time.time() - t_ai:.1f} วิ · มั่นใจ {conf}%{tag_}")
                    except AFQuotaDayError:
                        call("ai_chip", "quota", "โควต้าหมด · ใช้เกณฑ์ในเครื่อง")
                        bi, conf, reason = af_local_pick(passed)
                        r["ai"] = (bi, conf, reason + " · โควต้า AI หมด", "เกณฑ์ในเครื่อง")
                        L("⚠ โควต้า AI วันนี้หมดทุกรุ่น — จอที่เหลือใช้เกณฑ์ในเครื่องแทน (ไม่ต้องรอคน)", "warn")
                    except Exception as e:
                        r["note"] = f"AI ไม่ตอบ: {str(e)[:70]}"
                        call("ai_chip", "err", "AI ไม่ตอบ → ให้คนเลือก")
                        L(f"   {r['screen']}: AI ไม่ตอบ ({str(e)[:60]}) → ให้คนเลือก", "err")
                    call("think_stop")
                    call("anim_stop")
                    wait_left = af_ai_pacing(m_used) - (time.time() - t_ai)
                    r["_pace"] = max(0.0, wait_left)
            else:
                bi, conf, reason = af_local_pick(passed)
                r["ai"] = (bi, conf, reason, "เกณฑ์ในเครื่อง")
            # [FEATURE #af-scene] จอดำ/ไม่เห็นสินค้า → นับเป็นรูปตัดออก · แสดงกรอบแดง + ชื่อซีนบนเวที แล้วค่อยๆ หาย
            for k_, inf_ in r["scn"].items():
                if inf_.get("noprod") and k_ < len(r["q"]) and r["q"][k_]["ok"] and (not r["ai"] or r["ai"][0] != k_):
                    r["q"][k_]["ok"] = False
                    r["q"][k_]["flags"] = list(r["q"][k_]["flags"]) + ["ไม่เห็นสินค้า"]
                    cnt["bad"] += 1
                    bad_n += 1
                    if k_ < len(states):
                        states[k_] = "bad"
            if r["scn"]:
                call("counters", {"bad": cnt["bad"]})
                marks, st_sc = [], list(states)
                for k_ in range(len(smalls)):
                    inf_ = r["scn"].get(k_)
                    if not inf_:
                        marks.append(None)
                        continue
                    if inf_.get("noprod"):
                        kind_, lab_ = "bad", "ไม่เห็นสินค้า"
                    elif inf_.get("dup"):
                        kind_ = "dup"
                        lab_ = f"ซ้ำกับ {used_sc[inf_['dup'] - 1]['screen']}" if inf_["dup"] <= len(used_sc) else "ซีนซ้ำ"
                        if st_sc[k_] == "ok":
                            st_sc[k_] = "dup"
                    else:
                        kind_, lab_ = "ok", inf_.get("scene") or ""
                    marks.append({"box": inf_.get("box"), "label": lab_, "kind": kind_})
                call("station", 1, "run", "เช็กซีนไม่ให้ซ้ำ…")
                fx("scenes", smalls, st_sc, marks, True, wait=0.5)
                time.sleep(1.5)
                fx("scenes", smalls, st_sc, marks, False, wait=0.4)
                states = st_sc
            need_h = (r["ai"] is None) or (r["ai"][1] < c["threshold"])
            if need_h:
                if r["ai"] and not r["note"]:
                    r["note"] = f"มั่นใจ {r['ai'][1]}% ต่ำกว่าเกณฑ์ {c['threshold']}%"
                call("station", 1, "warn", "รอคนเลือก…")
                call("status", "รอคนยืนยัน")
                call("think_stop")
                fx("grid", smalls, [("pick" if (r["ai"] and i == r["ai"][0]) else s_) for i, s_ in enumerate(states)])
                L(f"   {r['screen']}: ส่งให้คนเลือก ({r['note']})", "warn")
                ev = threading.Event()
                post(self._af_review, [r], ev, list(used_sc))
                ev.wait()
                if r["final"] is None:
                    r["by"] = "ข้าม"
                    call("station", 1, "skip", "ข้ามจอนี้")
                    L(f"   {r['screen']}: ข้าม", "warn")
                    done_n += 1
                    call("progress", si + 1, len(todo))
                    continue
                cnt["human"] += 1
            else:
                r["final"], r["by"] = r["ai"][0], r["ai"][3]
                cnt["auto"] += 1
            call("counters", {"auto": cnt["auto"], "human": cnt["human"]})
            fi = r["final"]
            if ai_acts is not None and r["ai"] and fi == r["ai"][0]:
                r["acts"], r["creason"] = list(ai_acts), ai_creason or "AI แนะนำ"
            else:
                r["acts"], r["creason"] = af_local_color(r["files"][fi])
            st2 = [("pick" if i == fi else ("bad" if states[i] == "bad" else "dim")) for i in range(len(smalls))]
            ai = r["ai"]
            try:
                pick_full = af_open_small(r["files"][fi], 1100)
            except Exception:
                pick_full = smalls[fi] if fi < len(smalls) else Image.new("RGB", (160, 90))
            call("status", "★ เลือกรูป (%d)" % (fi + 1))
            fx("pick", smalls, st2, min(fi, len(smalls) - 1), pick_full)
            by_txt = {"AI": "AI", "ใช้ผลเดิม": "AI (ใช้ผลเดิม)", "คน": "คนเลือก",
                      "AI สำรอง": "AI สำรอง · " + r.get("ai_model", "").replace("gemini-", "")}.get(r["by"], "เกณฑ์ในเครื่อง")
            reason_txt = (ai[2] if (ai and fi == ai[0]) else "คนเลือกรูปนี้เอง")
            # [FEATURE #af-scene] ซีนของรูปที่ใช้จริง → จำไว้ให้จอถัดไป (ซีนซ้ำไม่เพิ่มรายการ)
            inf_f = r["scn"].get(fi) or {}
            r["scene"] = inf_f.get("scene", "")
            r["scene_box"] = inf_f.get("box")
            r["scene_dup"] = ""
            r["scene_crop"] = None
            if inf_f:
                try:
                    bx_ = inf_f.get("box")
                    if bx_:
                        W_, H_ = pick_full.size
                        r["scene_crop"] = pick_full.crop((int(bx_[0] * W_), int(bx_[1] * H_), int(bx_[2] * W_), int(bx_[3] * H_)))
                except Exception:
                    r["scene_crop"] = None
                if inf_f.get("dup") and inf_f["dup"] <= len(used_sc):
                    r["scene_dup"] = used_sc[inf_f["dup"] - 1]["screen"]
                    used_sc[inf_f["dup"] - 1]["count"] = used_sc[inf_f["dup"] - 1].get("count", 1) + 1
                    why_dup = "ไม่มีตัวเลือก" if r.get("scene_forced", True) or r["by"] != "คน" else "คนเลือกเอง"
                    reason_txt += f" · ซีนซ้ำกับ {r['scene_dup']} ({why_dup})"
                    r["scene_why"] = f"ซ้ำกับ {r['scene_dup']} · {why_dup}"
                elif r["scene"]:
                    used_sc.append({"name": r["scene"], "screen": r["screen"], "count": 1,
                                    "b64": af_img_b64(r["scene_crop"]) if r["scene_crop"] is not None else ""})
                    r["scene_why"] = "ซีนใหม่"
                if r["scene"]:
                    reason_txt += f" · ซีน: {r['scene']}"
                if r.get("scene_note"):
                    reason_txt += f" · {r['scene_note']}"
            call("card_pick", (ai[1] if (ai and fi == ai[0]) else None), reason_txt, by_txt)
            pick_small = smalls[fi] if fi < len(smalls) else af_open_small(r["files"][fi], 640)
            call("station", 1, "done", f"ตัดเสีย {bad_n} · เลือก ({fi + 1})" + (f" {ai[1]}%" if ai and fi == ai[0] else ""), pick_small)
            L(f"   {r['screen']}: เลือกรูป ({fi + 1}) · {by_txt}" + (f" · {ai[1]}% · {ai[2]}" if ai and fi == ai[0] else ""), "ai")
            time.sleep(max(1.6, r.pop("_pace", 0.0)) if c["use_ai"] else 1.5)
            # ── ③ แต่งสี ──
            src_p = r["files"][fi]
            call("station", 2, "run", "กำลังแต่งสี…")
            prev = af_open_small(src_p, 1100)
            prev_after = af_apply_color(prev, r["acts"])
            call("card_color", r["acts"], r["creason"])
            fx("wipe", prev, prev_after, wait=1.9)
            half = prev_after.copy()
            half.paste(prev.crop((0, 0, prev.width // 2, prev.height)), (0, 0))
            call("station", 2, "done", " → ".join(AF_COLOR_LABEL.get(a_, a_).split(" ", 1)[-1] for a_ in r["acts"]) or "ไม่ต้องแต่ง", half)
            # ── ④ ครอป ──
            with Image.open(src_p) as im0:
                exif = im0.info.get("exif")
                full = im0.convert("RGB")
            ow, oh = full.size
            full = af_apply_color(full, r["acts"])
            box0, (side, pct) = af_crop_box(ow, oh)
            # [FEATURE #af-scene] มีกรอบจอจาก AI → เลื่อนกรอบ 16:9 ให้จออยู่กลางภาพ
            box, ftxt, _b0 = af_crop_box_focus(ow, oh, r.get("scene_box"))
            sx, sy = prev_after.width / ow, prev_after.height / oh
            pbox0 = (box0[0] * sx, box0[1] * sy, box0[2] * sx, box0[3] * sy)
            pbox = (box[0] * sx, box[1] * sy, box[2] * sx, box[3] * sy)
            moved = ftxt and box != box0
            call("station", 3, "run", "กำลังครอป…")
            if pct < 0.5:
                ctxt = "④ สัดส่วน 16:9 พอดีแล้ว ไม่ต้องตัด"
            elif moved:
                ctxt = f"④ ครอป 16:9 ให้จออยู่กลาง · {ftxt}"
            else:
                ctxt = f"④ ตัดขอบ{side} {pct:.1f}% ให้เป็น 16:9"
            call("card_crop", ctxt + f" → ⑤ ย่อจาก {ow}×{oh} เป็น {TARGET_WIDTH}×{TARGET_HEIGHT} → ⑥ ประทับวันที่จาก EXIF")
            if moved:
                fx("crop3", prev_after, pbox0, pbox)
                time.sleep(1.7)
            else:
                fx("crop", prev_after, pbox)
                time.sleep(1.3)
            cropped = full.crop(box)
            call("station", 3, "done", ("จอกลาง · " + ftxt) if moved else (f"ตัด{side} {pct:.0f}%" if pct >= 0.5 else "สัดส่วนพอดีแล้ว"),
                 prev_after.crop(tuple(int(v) for v in pbox)))
            # ── ⑤ รีไซส์ ──
            call("station", 4, "run", "กำลังย่อ…")
            resized = cropped.resize((TARGET_WIDTH, TARGET_HEIGHT), Image.LANCZOS) if cropped.size != (TARGET_WIDTH, TARGET_HEIGHT) else cropped
            rs_small = resized.resize((BW, BH), Image.BILINEAR)
            fx("resize", rs_small, (ow, oh), (TARGET_WIDTH, TARGET_HEIGHT), wait=1.15)
            call("station", 4, "done", f"{ow}×{oh} → {TARGET_WIDTH}×{TARGET_HEIGHT}", rs_small)
            # ── ⑥ ประทับวันที่ ──
            call("station", 5, "run", "กำลังประทับ…")
            dt = extract_date_from_image(src_p)
            r["date"] = f"{dt[0]:02d}.{dt[1]:02d}.{dt[2]}" if dt else ""
            fpath = STAMP_FONT_PATH or af_fallback_font()
            stamped = render_stamp(resized.convert("RGBA"), r["date"], font_path=fpath) if (r["date"] and fpath) else resized
            dst = dest(d_stamp, os.path.basename(src_p))
            kw = {"quality": 95}
            if exif:
                kw["exif"] = exif
            stamped.save(dst, **kw)
            r["stamped"] = dst
            st_small = stamped.convert("RGB").resize((BW, BH), Image.BILINEAR)
            fx("stamp", rs_small, st_small, wait=1.15)
            call("station", 5, "done", r["date"] or "ไม่มีวันที่", st_small)
            # ── ⑦ ล้างชื่อไฟล์ ──
            call("station", 6, "run", "กำลังตั้งชื่อ…")
            name = os.path.basename(dst)
            clean = CLEAN_NUMBER_SUFFIX_RE.sub("", name)
            stem = os.path.splitext(clean)[0]
            sc_key = find_dgt_in_text(r["screen"])
            if not (find_dgt_in_text(stem) or slugify_name(r["screen"]) in slugify_name(stem)) or \
                    (sc_key and find_dgt_in_text(stem) != sc_key):
                clean = r["screen"] + os.path.splitext(name)[1]
            fdst = dest(d_final, clean)
            shutil.copy2(dst, fdst)
            r["final_file"] = fdst
            final_paths.append(fdst)
            fx("rename", st_small, name, os.path.basename(fdst), wait=2.8)
            call("station", 6, "done", os.path.splitext(os.path.basename(fdst))[0][:22], st_small)
            # ── ⑧ ลง PowerPoint (วางจริงตอนจบ ลำดับเดียวกัน) ──
            slot = slot_plan[len(final_paths) - 1] if len(final_paths) <= len(slot_plan) else None
            r["slot"] = slot
            call("station", 7, "run", "กำลังวาง…")
            call("ppt_update", pages, len(final_paths), True)
            if slot:
                n_pg = dict(pages).get(slot[0], 4)
                sl_w, sl_h = int(BW * 0.78), int(BH * 0.78)
                sl_empty, rects = af_slide_big(slot[1] - 1, n_pg, sl_w, sl_h, hl=slot[1] - 1,
                                               title=f"หน้า {slot[0]}", font=U["pil_fs"])
                fx("fly", st_small, sl_empty, rects[slot[1] - 1], wait=1.05)
                sl_full, _r = af_slide_big(slot[1] - 1, n_pg, sl_w, sl_h, hl=slot[1] - 1, img=st_small,
                                           title=f"หน้า {slot[0]}", font=U["pil_fs"])
                fin_fr = af_fly2(st_small, 1.0, sl_full, BW, BH, target=rects[slot[1] - 1])
                fx("static", af_caption2(fin_fr, f"✓  หน้า {slot[0]} · ช่อง {slot[1]}", fnt, BW, BH,
                                         col=(6, 95, 70, 230), pos="bc"), wait=0.7)
            else:
                fin_fr = af_caption2(st_small, "ลง PowerPoint ตอนจบงาน", fnt, BW, BH, col=(40, 60, 110, 230), pos="bc")
                fx("static", fin_fr, wait=0.7)
            call("station", 7, "done", f"หน้า {slot[0]} · ช่อง {slot[1]}" if slot else "ต่อท้าย",
                 fin_fr if slot else None)
            done_n += 1
            call("strip_add", r.get("scene_crop") if r.get("scene_crop") is not None else st_small,
                 "คน" if r["by"] == "คน" else "AI", done_n, r.get("scene", ""), bool(r.get("scene_dup")))
            el = time.time() - t_loop
            eta = el / (si + 1) * (len(todo) - si - 1)
            call("progress", si + 1, len(todo), eta)
            L(f"✓ {r['screen']} → {os.path.basename(fdst)}" + (f" · หน้า {slot[0]} ช่อง {slot[1]}" if slot else ""), "ok")
            time.sleep(0.3)
        # ── บันทึก PowerPoint ──
        call("anim_stop")
        qk = c.get("quality") if c.get("quality") in PP_QUALITY else "standard"
        call("stage_msg", "กำลังบันทึก PowerPoint…",
             f"ย่อรูป {PP_QUALITY[qk]['label']} · วาง {len(final_paths)} รูป · ก๊อปหน้าต้นแบบเพิ่มเองถ้าไม่พอ")
        ppt_out = os.path.join(out, os.path.splitext(os.path.basename(c["tpl"]))[0] + "_AutoFlow.pptx")
        gres = None
        if final_paths:
            gres = af_generate_pptx(c["tpl"], final_paths, int(c["start"]), ppt_out, order=order,
                                    quality=qk, start_auto=start_auto)
            result["ppt"] = ppt_out
            L(f"⑧ PowerPoint: วาง {gres['placed']} รูป {gres['how']}"
              + (f" · ก๊อปหน้าเพิ่ม {gres['added']} หน้า" if gres["added"] else "")
              + f" · {gres['size'] / 1048576:,.1f} MB ({PP_QUALITY[qk]['label']}) → {os.path.basename(ppt_out)}", "ok")
            if gres["leftover"]:
                L(f"   วางไม่ได้ {gres['leftover']} รูป — ตรวจเทมเพลตอีกครั้ง", "warn")
        pages_of = {os.path.basename(k_): v_ for k_, v_ in (gres["pages"].items() if gres else [])}
        # ── CSV ──
        import csv
        csv_path = os.path.join(out, "autoflow_log.csv")
        with open(csv_path, "w", newline="", encoding="utf-8-sig") as fh:
            w = csv.writer(fh)
            w.writerow(["จอ", "จำนวนรูป", "ผ่านเกณฑ์", "รูปที่ตัดออก", "ตัวเลือก (AI/เกณฑ์)", "ความมั่นใจ %", "เหตุผล",
                        "รูปที่ใช้จริง", "ซีน", "ซ้ำกับจอ", "เหตุผลซีน", "ตัดสินโดย", "ตรงกับ AI", "ปุ่มแต่งสี", "เหตุผลแต่งสี", "วันที่ประทับ",
                        "ไฟล์สุดท้าย", "หน้า PPT", "หมายเหตุ"])
            for r in rows:
                ai = r["ai"]
                rej = "; ".join(f"({k + 1}) {'/'.join(q['flags'])}" for k, q in enumerate(r["q"]) if not q["ok"])
                fin = r["final"]
                ff = os.path.basename(r.get("final_file", "")) if r.get("final_file") else ""
                w.writerow([r["screen"], len(r["files"]), sum(1 for q in r["q"] if q["ok"]), rej,
                            (f"({ai[0] + 1}) {os.path.basename(r['files'][ai[0]])}" if ai else ""),
                            (ai[1] if ai else ""), (ai[2] if ai else ""),
                            (f"({fin + 1}) {os.path.basename(r['files'][fin])}" if fin is not None else ""),
                            r.get("scene", ""), r.get("scene_dup", ""),
                            (r.get("scene_why", "") + (f" · {r['scene_note']}" if r.get("scene_note") else "")).strip(" ·"),
                            (f"{r['by']} ({r['ai_model']})" if r["by"] == "AI สำรอง" else r["by"]),
                            ("ใช่" if (ai and fin == ai[0]) else ("ไม่ใช่" if fin is not None and ai else "")),
                            " → ".join(AF_COLOR_LABEL.get(a_, a_) for a_ in r.get("acts") or []), r.get("creason", ""),
                            r.get("date", ""), ff, (pages_of.get(ff, "—") if ff else ""),
                            (r["note"] + (" · ใช้ผลเดิม" if r.get("cache") else "")
                             + (" · เหตุผลเขียนใหม่ (ของ AI ขัดกับผลเช็กซีน)" if r.get("reason_fixed") else "")).strip(" ·")])
            w.writerow([])
            w.writerow(["เวลาดึงจาก NAS", f"{fetch_s:.1f} วินาที"])
            w.writerow(["เวลารวม", f"{time.time() - t_all:.1f} วินาที"])
            w.writerow(["โมเดล", c["model"] if c["use_ai"] else "ไม่ใช้ AI"])
            fb_n = {}
            for r in rows:
                if r.get("by") == "AI สำรอง":
                    fb_n[r.get("ai_model", "")] = fb_n.get(r.get("ai_model", ""), 0) + 1
            for m_, n_ in fb_n.items():
                w.writerow(["รุ่นสำรอง", f"{m_} · {n_} จอ"])
        result["csv"] = csv_path
        flags = [f for r in rows for q in r["q"] if not q["ok"] for f in q["flags"]]
        hum = [r for r in rows if r["by"] == "คน"]
        result["summary"] = {
            "done": len(final_paths), "imgs": n_img, "elapsed": time.time() - t_all,
            "auto": sum(1 for r in rows if r.get("final_file") and r["by"] != "คน"),
            "human": len(hum), "same": sum(1 for r in hum if r["ai"] and r["final"] == r["ai"][0]),
            "bad": len(flags), "blur": flags.count("เบลอ"), "dark": flags.count("มืดเกินไป"),
            "bright": flags.count("สว่างจ้าเกินไป"), "baseline": c.get("baseline", 3),
            "model": c["model"] if c["use_ai"] else "", "stopped": stopped,
            "pp_added": (gres or {}).get("added", 0), "pp_size": (gres or {}).get("size", 0),
            "pp_quality": PP_QUALITY[qk]["label"], "pp_start": (gres or {}).get("start") or c["start"],
            "pp_limit_kb": PP_QUALITY[qk]["limit_kb"] * max(1, len(final_paths)) / PP_QUALITY_REF_SCREENS,
            "scene_new": sum(1 for r in rows if r.get("final_file") and r.get("scene") and not r.get("scene_dup")),
            "scene_dup": sum(1 for r in rows if r.get("final_file") and r.get("scene_dup")),
            "scene_known": sum(1 for r in rows if r.get("final_file") and r.get("scene")),
            "thumbs": [(af_open_small(r["final_file"], 320), r["screen"], "คน" if r["by"] == "คน" else "AI")
                       for r in rows if r.get("final_file")]}
        result["ok"] = not stopped
        L(f"สรุป · {len(final_paths)} จอ · {time.time() - t_all:.0f} วินาที · CSV → autoflow_log.csv", "h")
    except Exception as e:
        logging.error("AUTO FLOW error: %s\n%s", e, traceback.format_exc())
        L(f"✕ หยุด: {e}", "err")
        call("stage_msg", "หยุดทำงาน", str(e)[:90])
    post(lambda: self._af_win is not None and self._af_win._af_done(result))

def _af_review(self, items, ev, used=None):
    """[DEMO #auto-flow] ขั้น ④ คนยืนยัน — ทีละจอ คลิกรูปที่จะใช้ (กรอบทอง = ที่ AI เสนอ)
    [FEATURE #af-scene] used = ซีนที่จอก่อนหน้าใช้แล้ว → แถบด้านบน + ป้าย "ซ้ำกับ …" บนรูป (เตือน ไม่ห้าม)"""
    used = list(used or [])
    A = self.AF_ACCENT
    parent = getattr(self, "_af_win", None) or self
    dlg = tk.Toplevel(parent)
    dlg.withdraw()
    dlg.title("ขั้น ④ คนยืนยัน")
    dlg.configure(bg=BG_MAIN)
    dlg.transient(parent)
    dlg.resizable(False, False)
    try:
        pw_ = parent.winfo_width()
    except Exception:
        pw_ = 1200
    W = max(900, min(1240, int(pw_ * 0.74)))
    kf = W / 1000
    outer, card = self._glow_card(dlg, WARNING, radius=18, bg=CARD_BG, border_alpha=170, glow_alpha=36)
    outer.configure(width=W, height=560)
    outer.pack(padx=10, pady=10)
    body = tk.Frame(card, bg=CARD_BG)
    body.pack(fill="both", expand=True, padx=18, pady=14)
    head = tk.Frame(body, bg=CARD_BG)
    head.pack(fill="x")
    self._neon_icon(head, int(42 * kf), CARD_BG, self._draw_person_glyph, WARNING).pack(side="left")
    hc = tk.Frame(head, bg=CARD_BG)
    hc.pack(side="left", padx=(10, 0))
    t1 = tk.Label(hc, text="", font=(FONT_TH, int(14 * kf), "bold"), bg=CARD_BG, fg=TEXT_PRIMARY)
    t1.pack(anchor="w")
    t2 = tk.Label(hc, text="", font=FONT_SMALL, bg=CARD_BG, fg=WARNING)
    t2.pack(anchor="w")
    cnt = tk.Label(head, text="", font=("Consolas", 10, "bold"), bg=_lerp_hex(CARD_BG, WARNING, 0.2), fg="#FFE08A", padx=10, pady=2)
    cnt.pack(side="right")
    tk.Frame(body, bg=WARNING, height=2).pack(fill="x", pady=(10, 12))
    keep_ph = []
    if used:
        ur = tk.Frame(body, bg=CARD_BG)
        ur.pack(fill="x", pady=(0, 10))
        tk.Label(ur, text="ซีนที่ใช้แล้ว", font=(FONT_TH, 9, "bold"), bg=CARD_BG, fg=TEXT_SECONDARY).pack(side="left", padx=(0, 10))
        for u in used[-8:]:
            cell_u = tk.Frame(ur, bg=CARD_BG)
            cell_u.pack(side="left", padx=(0, 8))
            thumb = None
            try:
                if u.get("b64"):
                    import base64, io
                    thumb = Image.open(io.BytesIO(base64.b64decode(u["b64"]))).convert("RGB")
            except Exception:
                thumb = None
            if thumb is not None:
                ph_u = ImageTk.PhotoImage(af_rounded(af_contain(thumb, 64, 36, _af_rgb("#06080F")), 6, _af_rgb(CARD_BG)))
                keep_ph.append(ph_u)
                tk.Label(cell_u, image=ph_u, bg=CARD_BG, bd=0).pack(side="left")
            tk.Label(cell_u, text=f"{u['name']}\n{u['screen']}", font=(FONT_TH, 8), bg=CARD_BG, fg=TEXT_SECONDARY,
                     justify="left").pack(side="left", padx=(4, 0))
    grid = tk.Frame(body, bg=CARD_BG)
    grid.pack(fill="both", expand=True)
    hint = tk.Label(body, text=f"คลิกรูปที่จะใช้ · กดเลข 1–{AF_MAX_AI_IMAGES} · Enter = ใช้ตามที่เสนอ · ซีนซ้ำเลือกได้ (แค่เตือน)", font=(FONT_TH, 8),
                    bg=CARD_BG, fg=TEXT_MUTED)
    hint.pack(side="left", pady=(10, 0))
    st = {"i": 0, "photos": []}
    skip_b = self._pill_button(body, "ข้ามจอนี้", CARD_BG_2, TEXT_SECONDARY, command=lambda: choose(None),
                               font=FONT_SMALL, padx=14, pady=5)
    skip_b.pack(side="right", pady=(10, 0))
    use_b = self._pill_button(body, "ใช้ตามที่เสนอ", _lerp_hex(CARD_BG, WARNING, 0.6), "#1F1400",
                              command=lambda: choose("ai"), font=(FONT_TH, 9, "bold"), padx=14, pady=5)
    use_b.pack(side="right", padx=(0, 8), pady=(10, 0))

    def show():
        for ch in grid.winfo_children():
            ch.destroy()
        st["photos"] = []
        r = items[st["i"]]
        cnt.configure(text=f"{st['i'] + 1} / {len(items)}")
        t1.configure(text=f"เลือกรูปเอง · {r['screen']}")
        ai = r["ai"]
        t2.configure(text=(r["note"] or "") + (f" · เสนอรูป ({ai[0] + 1}) เพราะ {ai[2]}" if ai else ""))
        if ai:
            use_b.pack(side="right", padx=(0, 8), pady=(10, 0))
        else:
            use_b.pack_forget()
        n = min(len(r["files"]), AF_MAX_AI_IMAGES)
        cols = 4 if n > 6 else (3 if n > 4 else max(1, n))
        if n in (3, 4):
            cols = n if n == 3 else 4
        tw = (W - 60) // cols - 18
        th = int(tw * 9 / 16)
        bgc = _af_rgb(CARD_BG)
        for k, f in enumerate(r["files"][:AF_MAX_AI_IMAGES]):
            q = r["q"][k]
            is_ai = bool(ai and ai[0] == k)
            bc = (255, 184, 0, 255) if is_ai else ((255, 140, 66, 220) if not q["ok"] else (70, 80, 108, 255))
            cell = tk.Frame(grid, bg=CARD_BG, cursor="hand2")
            cell.grid(row=k // cols, column=k % cols, padx=9, pady=8)
            try:
                im = af_cover(af_open_small(f, max(tw, th) * 2), tw, th)
                if not q["ok"]:
                    im = Image.blend(im, Image.new("RGB", im.size, (12, 16, 24)), 0.5)
                ph = ImageTk.PhotoImage(af_rounded(im, 12, bgc, border=bc, bw=3 if is_ai else 2))
                ph_h = ImageTk.PhotoImage(af_rounded(im, 12, bgc, border=(34, 211, 238, 255), bw=3))
            except Exception:
                ph = ph_h = None
            st["photos"].extend([ph, ph_h])
            lb = tk.Label(cell, image=ph, bg=CARD_BG, bd=0)
            lb.pack()
            cr = tk.Frame(cell, bg=CARD_BG)
            cr.pack(fill="x", pady=(6, 0))
            tk.Label(cr, text=f" {k + 1} ", font=("Consolas", 9, "bold"), bg=WARNING if is_ai else CARD_BG_2,
                     fg="#1F1400" if is_ai else TEXT_PRIMARY).pack(side="left")
            cap = ("★ AI เสนอ · " if is_ai else "") + ("⚠ " + "/".join(q["flags"]) if not q["ok"] else "ผ่านเกณฑ์")
            inf = (r.get("scn") or {}).get(k) or {}
            is_dup = bool(inf.get("dup")) and inf["dup"] <= len(used)
            if inf.get("scene") and q["ok"]:
                cap += f" · ซีน: {inf['scene']}"
            if is_dup:
                cap = cap + f" · ซ้ำกับ {used[inf['dup'] - 1]['screen']}"
            cl = tk.Label(cr, text=cap, font=(FONT_TH, 9, "bold"), bg=CARD_BG, wraplength=max(120, tw - 30), justify="left",
                          fg=WARNING if is_ai else ("#FFB38A" if (is_dup or not q["ok"]) else TEXT_SECONDARY))
            cl.pack(side="left", padx=(8, 0))

            def _hov(e, lb=lb, a=ph, b=ph_h, on=True):
                if a is not None:
                    lb.configure(image=b if on else a)
            lb.bind("<Enter>", lambda e, h=_hov: h(e, on=True))
            lb.bind("<Leave>", lambda e, h=_hov: h(e, on=False))
            for w_ in (cell, lb, cl, cr):
                w_.bind("<Button-1>", lambda e, k=k: choose(k))
        try:
            self._autosize_card_height(outer, card, radius=18)
        except Exception:
            pass

    def choose(k):
        r = items[st["i"]]
        if k == "ai":
            if not r["ai"]:
                return
            k = r["ai"][0]
        if k is None:
            r["final"], r["by"] = None, "ข้าม"
        else:
            if k >= len(r["files"]):
                return
            r["final"], r["by"] = k, "คน"
        st["i"] += 1
        if st["i"] >= len(items):
            finish()
        else:
            show()

    def finish():
        try:
            dlg.destroy()
        except Exception:
            pass
        ev.set()

    def on_close():
        for r in items[st["i"]:]:
            r["final"], r["by"] = None, "ข้าม"
        finish()
    dlg.protocol("WM_DELETE_WINDOW", on_close)
    dlg.bind("<Return>", lambda e: choose("ai"))
    dlg.bind("<Escape>", lambda e: choose(None))
    for d in range(1, AF_MAX_AI_IMAGES + 1):
        dlg.bind(str(d), lambda e, d=d: choose(d - 1))
    dlg._keep_ph = keep_ph
    show()
    dlg.update_idletasks()
    x = parent.winfo_rootx() + (parent.winfo_width() - (W + 20)) // 2
    y = parent.winfo_rooty() + max(40, (parent.winfo_height() - dlg.winfo_reqheight()) // 2)
    dlg.geometry(f"+{max(0, x)}+{max(0, y)}")
    try:
        dlg.attributes("-alpha", 0.0)
    except Exception:
        pass
    dlg.deiconify()

    def _fade(i=1):
        try:
            dlg.attributes("-alpha", min(1.0, i / 8))
            if i < 8:
                dlg.after(22, lambda: _fade(i + 1))
        except Exception:
            pass
    _fade()
    dlg.grab_set()
    dlg.focus_force()
    dlg._af_choose = choose
    self._af_review_dlg = dlg
