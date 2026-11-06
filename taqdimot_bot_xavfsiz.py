# -*- coding: utf-8 -*-
"""
TAQDIMOT BOT  -  Telegram'da AI yordamida premium dizaynli PPTX taqdimot yaratuvchi pullik bot.
Bitta fayl: foydalanuvchi qismi + to'lov (chek orqali) + admin panel.

O'rnatish (Pydroid / Termux / VPS):
    pip install pyTelegramBotAPI python-pptx requests pillow

Sozlash: BOT_TOKEN va GEMINI_API_KEY ni environment variables orqali bering; so'ng python taqdimot_bot.py
"""
import os
import sys
import io
import re
import json
import time
import html
import sqlite3
import logging
import threading
import traceback
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor

import requests
import telebot
from telebot import types
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

# ============================================================================
#                               SOZLAMALAR
# ============================================================================
BOT_TOKEN = os.getenv("BOT_TOKEN", "").strip()
ADMIN_IDS = [123456789]                      # admin Telegram ID lari (@userinfobot dan oling)

# Matn yaratish uchun Gemini API kaliti (bepul: https://aistudio.google.com/apikey)
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "").strip()
GEMINI_MODEL = "gemini-2.5-flash"

# Slaydlardagi rasmlar uchun Pexels kaliti (bepul: https://www.pexels.com/api/). Bo'sh qoldirsangiz rasmsiz dizayn ishlaydi.
PEXELS_API_KEY = os.getenv("PEXELS_API_KEY", "").strip()

CARD_NUMBER = "8600 0000 0000 0000"          # to'lov qabul qilinadigan karta
CARD_OWNER = "ISM FAMILIYA"
FREE_CREDITS = 1                             # yangi foydalanuvchiga sovg'a taqdimot soni

PACKAGES = [
    {"id": "p1",  "title": "1 ta taqdimot",  "credits": 1,  "price": 7000},
    {"id": "p5",  "title": "5 ta taqdimot",  "credits": 5,  "price": 29000},
    {"id": "p15", "title": "15 ta taqdimot", "credits": 15, "price": 75000},
]
SLIDE_COUNTS = [8, 10, 12, 15]               # slaydlar soni variantlari (titul, reja va yakunni ham hisoblaydi)
MAX_PARALLEL = 2                             # bir vaqtda nechta taqdimot yaratilsin
DB_FILE = "taqdimot_bot.db"
# ============================================================================

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger("taqdimot")

# Maxfiy kalitlar kodda saqlanmaydi; ishga tushirishdan oldin environment variables orqali bering.
if not BOT_TOKEN:
    raise SystemExit("XATO: BOT_TOKEN environment variable sozlanmagan.")
if not GEMINI_API_KEY:
    raise SystemExit("XATO: GEMINI_API_KEY environment variable sozlanmagan.")

bot = telebot.TeleBot(BOT_TOKEN, parse_mode="HTML", threaded=True, num_threads=8)
gen_sem = threading.Semaphore(MAX_PARALLEL)

# ----------------------------------------------------------------------------
#                                  MA'LUMOTLAR BAZASI
# ----------------------------------------------------------------------------
db = sqlite3.connect(DB_FILE, check_same_thread=False)
db.row_factory = sqlite3.Row
db_lock = threading.Lock()


def dbx(sql, args=(), fetch=False):
    with db_lock:
        cur = db.execute(sql, args)
        rows = cur.fetchall() if fetch else None
        db.commit()
        return rows if fetch else cur


dbx("""CREATE TABLE IF NOT EXISTS users(
    id INTEGER PRIMARY KEY, name TEXT, username TEXT, credits INTEGER DEFAULT 0,
    made INTEGER DEFAULT 0, banned INTEGER DEFAULT 0, joined TEXT)""")
dbx("""CREATE TABLE IF NOT EXISTS payments(
    id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER, package TEXT, credits INTEGER,
    amount INTEGER, file_id TEXT, status TEXT DEFAULT 'pending', created TEXT)""")


def now():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def ensure_user(u):
    rows = dbx("SELECT id FROM users WHERE id=?", (u.id,), fetch=True)
    if not rows:
        dbx("INSERT INTO users(id,name,username,credits,joined) VALUES(?,?,?,?,?)",
            (u.id, u.first_name or "", u.username or "", FREE_CREDITS, now()))
        return True
    dbx("UPDATE users SET name=?, username=? WHERE id=?", (u.first_name or "", u.username or "", u.id))
    return False


def get_user(uid):
    rows = dbx("SELECT * FROM users WHERE id=?", (uid,), fetch=True)
    return rows[0] if rows else None


def add_credits(uid, n):
    dbx("UPDATE users SET credits=credits+? WHERE id=?", (n, uid))


def take_credit(uid):
    cur = dbx("UPDATE users SET credits=credits-1 WHERE id=? AND credits>0", (uid,))
    return cur.rowcount > 0


def is_admin(uid):
    return uid in ADMIN_IDS


def money(n):
    return f"{n:,}".replace(",", " ") + " so'm"


# ----------------------------------------------------------------------------
#                                  DIZAYN MAVZULARI
# ----------------------------------------------------------------------------
THEMES = {
    "gold": dict(name="Qora va Oltin", bg="0F1115", panel="1A1D24", text="F5F1E8", muted="A8A39A",
                 accent="D4AF37", accent2="F3E2A0", on_accent="0F1115", head="Playfair Display", body="Montserrat"),
    "ocean": dict(name="Okean", bg="0B1D33", panel="12304F", text="EAF4FF", muted="8FB3D9",
                  accent="38BDF8", accent2="7DD3FC", on_accent="0B1D33", head="Poppins", body="Inter"),
    "emerald": dict(name="Zumrad", bg="0C2620", panel="143A30", text="ECFDF5", muted="8FBFAE",
                    accent="34D399", accent2="A7F3D0", on_accent="0C2620", head="Raleway", body="Open Sans"),
    "royal": dict(name="Shoh binafsha", bg="1B1033", panel="2A1B4D", text="F3EEFF", muted="B3A3D9",
                  accent="C084FC", accent2="E9D5FF", on_accent="1B1033", head="Cormorant Garamond", body="Lato"),
    "sunset": dict(name="Quyosh botishi", bg="1E1118", panel="2D1A25", text="FFF1F2", muted="D9A3B0",
                   accent="FB7185", accent2="FDBA74", on_accent="1E1118", head="DM Serif Display", body="DM Sans"),
    "light": dict(name="Minimal oq", bg="FAFAF7", panel="EFEDE6", text="1F2328", muted="6B7280",
                  accent="E4572E", accent2="F4A261", on_accent="FFFFFF", head="Lora", body="Inter"),
}

LABELS = {
    "uz": dict(agenda="Reja", thanks="E'tiboringiz uchun rahmat!", author="Muallif", deck="TAQDIMOT", lang_name="o'zbek (lotin)"),
    "ru": dict(agenda="План", thanks="Спасибо за внимание!", author="Автор", deck="ПРЕЗЕНТАЦИЯ", lang_name="русский"),
    "en": dict(agenda="Agenda", thanks="Thank you!", author="Author", deck="PRESENTATION", lang_name="English"),
}

W, H = 13.333, 7.5


def rgb(h):
    return RGBColor.from_string(h)


def shp(slide, kind, x, y, w, h, fill=None, line=None, lw=1.25):
    s = slide.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill:
        s.fill.solid()
        s.fill.fore_color.rgb = rgb(fill)
    else:
        s.fill.background()
    if line:
        s.line.color.rgb = rgb(line)
        s.line.width = Pt(lw)
    else:
        s.line.fill.background()
    s.shadow.inherit = False
    return s


def tb(slide, x, y, w, h, text, size=14, color="FFFFFF", bold=False, font="Calibri",
       align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, italic=False, spacing=None):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    p = tf.paragraphs[0]
    p.alignment = align
    if spacing:
        p.line_spacing = spacing
    r = p.add_run()
    r.text = text
    f = r.font
    f.size = Pt(size)
    f.bold = bold
    f.italic = italic
    f.name = font
    f.color.rgb = rgb(color)
    return box


def set_bg(slide, color):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = rgb(color)


def fit_image(raw, w, h):
    from PIL import Image
    im = Image.open(io.BytesIO(raw)).convert("RGB")
    target = w / h
    iw, ih = im.size
    if iw / ih > target:
        nw = int(ih * target)
        left = (iw - nw) // 2
        im = im.crop((left, 0, left + nw, ih))
    else:
        nh = int(iw / target)
        top = (ih - nh) // 2
        im = im.crop((0, top, iw, top + nh))
    if im.size[0] > 1400:
        k = 1400 / im.size[0]
        im = im.resize((1400, int(im.size[1] * k)))
    out = io.BytesIO()
    im.save(out, "JPEG", quality=85)
    out.seek(0)
    return out


def get_image(query, w, h):
    if not PEXELS_API_KEY or not query:
        return None
    try:
        r = requests.get("https://api.pexels.com/v1/search",
                         params={"query": query, "per_page": 1, "orientation": "portrait" if h > w else "landscape"},
                         headers={"Authorization": PEXELS_API_KEY}, timeout=12)
        photos = r.json().get("photos") or []
        if not photos:
            return None
        raw = requests.get(photos[0]["src"]["large"], timeout=20).content
        return fit_image(raw, w, h)
    except Exception as e:
        log.warning("rasm olinmadi: %s", e)
        return None


def footer(slide, T, topic, idx, total):
    shp(slide, MSO_SHAPE.RECTANGLE, 0.8, 6.92, W - 1.6, 0.012, fill=T["muted"])
    tb(slide, 0.8, 7.0, 8, 0.3, topic[:70], size=10, color=T["muted"], font=T["body"])
    tb(slide, W - 3.8, 7.0, 3.0, 0.3, f"{idx:02d} / {total:02d}", size=10, color=T["muted"],
       font=T["body"], align=PP_ALIGN.RIGHT)


def slide_title(slide, T, text, x=0.8, w=11.7):
    shp(slide, MSO_SHAPE.RECTANGLE, x, 0.62, 0.9, 0.07, fill=T["accent"])
    size = 32 if len(text) <= 38 else (26 if len(text) <= 60 else 22)
    tb(slide, x, 0.85, w, 0.9, text, size=size, color=T["text"], bold=True, font=T["head"], anchor=MSO_ANCHOR.TOP)


def clip(s, n):
    s = (s or "").strip()
    return s if len(s) <= n else s[:n - 1].rsplit(" ", 1)[0] + "..."


# ----------------------------------------------------------------------------
#                                  SLAYD TURLARI
# ----------------------------------------------------------------------------
def s_title(prs, T, data, author, lang, img):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s, T["bg"])
    L = LABELS[lang]
    if img:
        s.shapes.add_picture(img, Inches(W * 0.55), 0, Inches(W * 0.45), Inches(H))
        shp(s, MSO_SHAPE.RECTANGLE, W * 0.55 - 0.12, 0, 0.12, H, fill=T["accent"])
    else:
        shp(s, MSO_SHAPE.OVAL, 8.3, -1.8, 8.2, 8.2, fill=T["panel"])
        shp(s, MSO_SHAPE.OVAL, 9.4, 1.2, 4.6, 4.6, line=T["accent"], lw=1.5)
        shp(s, MSO_SHAPE.OVAL, 11.3, 5.2, 1.5, 1.5, fill=T["accent"])
        shp(s, MSO_SHAPE.OVAL, 8.6, 5.9, 0.5, 0.5, fill=T["accent2"])
    tw = 7.0
    shp(s, MSO_SHAPE.RECTANGLE, 0.9, 1.3, 0.9, 0.07, fill=T["accent"])
    tb(s, 0.9, 1.5, tw, 0.4, L["deck"], size=12, color=T["accent"], bold=True, font=T["body"])
    title = clip(data.get("title", ""), 90)
    size = 44 if len(title) <= 28 else (36 if len(title) <= 50 else 28)
    tb(s, 0.9, 2.1, tw, 2.6, title, size=size, color=T["text"], bold=True, font=T["head"], spacing=1.05)
    tb(s, 0.9, 4.75, tw, 1.0, clip(data.get("subtitle", ""), 140), size=16, color=T["muted"], font=T["body"])
    if author:
        shp(s, MSO_SHAPE.RECTANGLE, 0.9, 6.15, 0.06, 0.62, fill=T["accent"])
        tb(s, 1.15, 6.15, 6, 0.3, L["author"].upper(), size=9, color=T["muted"], bold=True, font=T["body"])
        tb(s, 1.15, 6.4, 6, 0.4, clip(author, 40), size=16, color=T["text"], bold=True, font=T["body"])
    return s


def s_agenda(prs, T, titles, lang, topic, idx, total):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s, T["bg"])
    shp(s, MSO_SHAPE.RECTANGLE, 0, 0, 4.6, H, fill=T["panel"])
    shp(s, MSO_SHAPE.RECTANGLE, 0.8, 2.6, 0.9, 0.07, fill=T["accent"])
    tb(s, 0.8, 2.85, 3.6, 1.2, LABELS[lang]["agenda"], size=40, color=T["text"], bold=True, font=T["head"])
    items = titles[:8]
    row = min(0.72, 5.6 / max(len(items), 1))
    y0 = 0.9 + (5.6 - row * len(items)) / 2
    for i, t in enumerate(items):
        y = y0 + i * row
        tb(s, 5.3, y, 0.9, row, f"{i + 1:02d}", size=20, color=T["accent"], bold=True, font=T["head"],
           anchor=MSO_ANCHOR.MIDDLE)
        tb(s, 6.3, y, 6.3, row, clip(t, 60), size=18, color=T["text"], font=T["body"], anchor=MSO_ANCHOR.MIDDLE)
        shp(s, MSO_SHAPE.RECTANGLE, 5.3, y + row - 0.02, 7.3, 0.01, fill=T["muted"])
    footer(s, T, topic, idx, total)
    return s


def s_bullets(prs, T, sd, img, left, topic, idx, total, num):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s, T["bg"])
    pts = (sd.get("points") or [])[:4]
    img_w, img_h = 4.0, 4.9
    px = 0.8 if left else W - 0.8 - img_w
    tx = (0.8 + img_w + 0.7) if left else 0.8
    tw = W - 1.6 - img_w - 0.7
    slide_title(s, T, clip(sd.get("title", ""), 80), x=tx, w=tw)
    # o'ng/chap panel
    if img:
        shp(s, MSO_SHAPE.RECTANGLE, px + 0.18, 1.18, img_w, img_h, fill=T["accent"])
        s.shapes.add_picture(img, Inches(px), Inches(1.0), Inches(img_w), Inches(img_h))
    else:
        shp(s, MSO_SHAPE.RECTANGLE, px, 1.0, img_w, img_h, fill=T["panel"])
        shp(s, MSO_SHAPE.RECTANGLE, px, 1.0, img_w, 0.08, fill=T["accent"])
        tb(s, px, 1.8, img_w, 2.4, f"{num:02d}", size=110, color=T["accent"], bold=True, font=T["head"],
           align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        tb(s, px + 0.4, 4.4, img_w - 0.8, 1.2, clip(sd.get("title", ""), 50), size=13, color=T["muted"],
           font=T["body"], align=PP_ALIGN.CENTER)
    n = max(len(pts), 1)
    row = min(1.2, 4.7 / n)
    y0 = 2.0
    for i, p in enumerate(pts):
        y = y0 + i * row
        shp(s, MSO_SHAPE.OVAL, tx, y + 0.04, 0.58, 0.58, fill=T["accent"])
        tb(s, tx, y + 0.04, 0.58, 0.58, str(i + 1), size=15, color=T["on_accent"], bold=True, font=T["body"],
           align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        tb(s, tx + 0.85, y, tw - 0.85, 0.38, clip(p.get("head", ""), 60), size=17, color=T["text"], bold=True,
           font=T["body"])
        body = clip(p.get("text", ""), 200)
        tb(s, tx + 0.85, y + 0.4, tw - 0.85, row - 0.4, body, size=13 if len(body) > 120 else 14,
           color=T["muted"], font=T["body"])
    footer(s, T, topic, idx, total)
    return s


def s_cards(prs, T, sd, topic, idx, total):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s, T["bg"])
    slide_title(s, T, clip(sd.get("title", ""), 80))
    pts = (sd.get("points") or [])[:4]
    n = max(len(pts), 1)
    gap = 0.3
    cw = (W - 1.6 - gap * (n - 1)) / n
    for i, p in enumerate(pts):
        x = 0.8 + i * (cw + gap)
        shp(s, MSO_SHAPE.RECTANGLE, x, 2.0, cw, 4.5, fill=T["panel"])
        shp(s, MSO_SHAPE.RECTANGLE, x, 2.0, cw, 0.09, fill=T["accent"])
        tb(s, x + 0.35, 2.4, cw - 0.7, 0.7, f"{i + 1:02d}", size=34, color=T["accent"], bold=True, font=T["head"])
        tb(s, x + 0.35, 3.25, cw - 0.7, 0.9, clip(p.get("head", ""), 50), size=17 if n < 4 else 15, color=T["text"],
           bold=True, font=T["body"])
        body = clip(p.get("text", ""), 190)
        tb(s, x + 0.35, 4.2, cw - 0.7, 2.2, body, size=13 if n < 4 else 12, color=T["muted"], font=T["body"])
    footer(s, T, topic, idx, total)
    return s


def s_end(prs, T, author, lang):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s, T["bg"])
    shp(s, MSO_SHAPE.OVAL, -2.2, -2.2, 6.0, 6.0, fill=T["panel"])
    shp(s, MSO_SHAPE.OVAL, W - 4.0, H - 3.4, 6.0, 6.0, fill=T["panel"])
    shp(s, MSO_SHAPE.OVAL, W - 2.6, 1.0, 0.6, 0.6, fill=T["accent"])
    shp(s, MSO_SHAPE.OVAL, 1.4, H - 1.9, 0.35, 0.35, fill=T["accent2"])
    shp(s, MSO_SHAPE.RECTANGLE, W / 2 - 0.55, 2.35, 1.1, 0.07, fill=T["accent"])
    t = LABELS[lang]["thanks"]
    tb(s, 1.0, 2.7, W - 2.0, 1.6, t, size=44 if len(t) < 24 else 36, color=T["text"], bold=True, font=T["head"],
       align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    if author:
        tb(s, 1.0, 4.6, W - 2.0, 0.5, clip(author, 40), size=16, color=T["accent"], bold=True, font=T["body"],
           align=PP_ALIGN.CENTER)
    return s


def build_pptx(data, theme_key, author, lang, images):
    T = THEMES.get(theme_key, THEMES["gold"])
    prs = Presentation()
    prs.slide_width = Inches(W)
    prs.slide_height = Inches(H)
    slides = data["slides"]
    total = len(slides) + 3
    topic = data.get("title", "")
    s_title(prs, T, data, author, lang, images.get("title"))
    s_agenda(prs, T, [x.get("title", "") for x in slides], lang, topic, 2, total)
    for i, sd in enumerate(slides):
        idx = i + 3
        kind = i % 3
        n_pts = len(sd.get("points") or [])
        if kind == 2 and n_pts >= 3:
            s_cards(prs, T, sd, topic, idx, total)
        else:
            s_bullets(prs, T, sd, images.get(i), left=(kind == 1), topic=topic, idx=idx, total=total, num=i + 1)
    s_end(prs, T, author, lang)
    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out


# ----------------------------------------------------------------------------
#                                  AI MATN YARATISH
# ----------------------------------------------------------------------------
def ai_content(topic, count, lang):
    body_n = count - 3
    L = LABELS[lang]["lang_name"]
    prompt = f"""Sen professional taqdimot muallifisan. Mavzu: "{topic}".
Taqdimot tili: {L}. Aynan {body_n} ta mazmunli slayd tayyorla.
Faqat JSON qaytar, boshqa matn yo'q. Format:
{{"title": "qisqa jozibali sarlavha (max 60 belgi)", "subtitle": "bir jumlalik izoh (max 110 belgi)",
 "slides": [{{"title": "slayd sarlavhasi (max 55 belgi)", "image_query": "2-4 ta INGLIZCHA so'z, rasm qidirish uchun",
   "points": [{{"head": "qisqa tezis (max 45 belgi)", "text": "1-2 jumlali aniq izoh (max 150 belgi)"}}]}}]}}
Qoidalar: har slaydda 3 yoki 4 ta point; faktlar aniq va foydali bo'lsin; mantiqiy tartib: kirish, asosiy qismlar, xulosa;
emoji ishlatma; to'qima raqam va manba keltirma."""
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{GEMINI_MODEL}:generateContent"
    last = None
    for _ in range(3):
        try:
            r = requests.post(url, params={"key": GEMINI_API_KEY}, timeout=150, json={
                "contents": [{"parts": [{"text": prompt}]}],
                "generationConfig": {"responseMimeType": "application/json", "temperature": 0.7}})
            txt = r.json()["candidates"][0]["content"]["parts"][0]["text"]
            txt = re.sub(r"^```(?:json)?|```$", "", txt.strip()).strip()
            data = json.loads(txt)
            slides = [s for s in data.get("slides", []) if s.get("title") and s.get("points")]
            if len(slides) < 3:
                raise ValueError("kam slayd")
            data["slides"] = slides[:body_n]
            return data
        except Exception as e:
            last = e
            log.warning("AI xato: %s", e)
            time.sleep(2)
    raise RuntimeError(f"AI javob bermadi: {last}")


def fetch_images(data):
    jobs = [("title", data.get("title", ""), W * 0.45, H)]
    for i, sd in enumerate(data["slides"]):
        if i % 3 != 2 or len(sd.get("points") or []) < 3:
            jobs.append((i, sd.get("image_query") or sd.get("title", ""), 4.0, 4.9))
    res = {}
    if not PEXELS_API_KEY:
        return res
    with ThreadPoolExecutor(max_workers=4) as ex:
        futs = {k: ex.submit(get_image, q, w, h) for k, q, w, h in jobs}
        for k, f in futs.items():
            try:
                res[k] = f.result()
            except Exception:
                res[k] = None
    return res


# ----------------------------------------------------------------------------
#                                  KLAVIATURALAR
# ----------------------------------------------------------------------------
BTN_NEW, BTN_ME, BTN_BUY, BTN_HELP, BTN_CANCEL = "Yangi taqdimot", "Hisobim", "Tarif sotib olish", "Yordam", "Bekor qilish"
states = {}


def main_kb():
    kb = types.ReplyKeyboardMarkup(resize_keyboard=True)
    kb.row(BTN_NEW)
    kb.row(BTN_ME, BTN_BUY)
    kb.row(BTN_HELP)
    return kb


def cancel_kb():
    kb = types.ReplyKeyboardMarkup(resize_keyboard=True)
    kb.row(BTN_CANCEL)
    return kb


def ik(rows):
    kb = types.InlineKeyboardMarkup()
    for r in rows:
        kb.row(*[types.InlineKeyboardButton(t, callback_data=d) for t, d in r])
    return kb


def e(s):
    return html.escape(str(s))


# ----------------------------------------------------------------------------
#                                  FOYDALANUVCHI QISMI
# ----------------------------------------------------------------------------
def guard(m):
    u = get_user(m.from_user.id)
    if u and u["banned"]:
        bot.send_message(m.chat.id, "Siz botdan foydalanishdan chetlatilgansiz.")
        return False
    return True


@bot.message_handler(commands=["start"])
def cmd_start(m):
    new = ensure_user(m.from_user)
    if not guard(m):
        return
    states.pop(m.from_user.id, None)
    u = get_user(m.from_user.id)
    txt = (f"Assalomu alaykum, <b>{e(m.from_user.first_name)}</b>.\n\n"
           "Men mavzu bo'yicha professional va premium dizaynli <b>PowerPoint taqdimot</b> tayyorlab beraman: "
           "tayyor matn, chiroyli ranglar, zamonaviy shriftlar va rasmlar bilan.\n\n")
    if new and FREE_CREDITS:
        txt += f"Sovg'a: sizga <b>{FREE_CREDITS} ta bepul taqdimot</b> berildi.\n"
    txt += f"Hisobingizda: <b>{u['credits']}</b> ta taqdimot.\n\nBoshlash uchun \"{BTN_NEW}\" tugmasini bosing."
    bot.send_message(m.chat.id, txt, reply_markup=main_kb())


@bot.message_handler(commands=["admin"])
def cmd_admin(m):
    if not is_admin(m.from_user.id):
        return
    states.pop(m.from_user.id, None)
    bot.send_message(m.chat.id, "<b>Admin panel</b>", reply_markup=admin_kb())


def start_new(m):
    u = get_user(m.from_user.id)
    if u["credits"] <= 0:
        bot.send_message(m.chat.id, "Hisobingizda taqdimot qolmagan. Davom etish uchun tarif tanlang.",
                         reply_markup=buy_kb())
        return
    states[m.from_user.id] = {"step": "topic"}
    bot.send_message(m.chat.id, "Taqdimot <b>mavzusini</b> yozing.\n\nMasalan: <i>Sun'iy intellektning kelajagi</i>",
                     reply_markup=cancel_kb())


def buy_kb():
    return ik([[(f"{p['title']}  -  {money(p['price'])}", f"pk:{p['id']}")] for p in PACKAGES])


@bot.message_handler(content_types=["text"])
def on_text(m):
    ensure_user(m.from_user)
    uid = m.from_user.id
    txt = (m.text or "").strip()
    if not guard(m):
        return

    if txt == BTN_CANCEL:
        states.pop(uid, None)
        bot.send_message(m.chat.id, "Bekor qilindi.", reply_markup=main_kb())
        return
    if txt == BTN_NEW:
        return start_new(m)
    if txt == BTN_ME:
        u = get_user(uid)
        bot.send_message(m.chat.id, f"<b>Hisobingiz</b>\n\nQolgan taqdimotlar: <b>{u['credits']}</b>\n"
                                    f"Yaratilgan taqdimotlar: <b>{u['made']}</b>\nID: <code>{uid}</code>")
        return
    if txt == BTN_BUY:
        bot.send_message(m.chat.id, "<b>Tariflar</b>\nO'zingizga mosini tanlang:", reply_markup=buy_kb())
        return
    if txt == BTN_HELP:
        bot.send_message(m.chat.id,
                         "<b>Qanday ishlaydi</b>\n1. \"Yangi taqdimot\" tugmasini bosing\n2. Mavzu, slaydlar soni, til va dizaynni tanlang\n"
                         "3. 1-2 daqiqada tayyor .pptx fayl olasiz\n\nFayl PowerPoint, Google Slides va telefon ilovalarida ochiladi. "
                         "Dizayndagi shriftlar (Montserrat, Playfair Display va h.k.) qurilmada o'rnatilgan bo'lsa, eng chiroyli ko'rinadi.\n\n"
                         "Savollar uchun adminga murojaat qiling.")
        return

    st = states.get(uid)
    # --- admin holatlari
    if st and is_admin(uid) and st["step"].startswith("adm_"):
        return admin_state(m, st, txt)
    # --- taqdimot oqimi
    if st and st["step"] == "topic":
        if len(txt) < 3 or len(txt) > 200:
            bot.send_message(m.chat.id, "Mavzu 3 dan 200 belgigacha bo'lsin. Qayta yozing.")
            return
        st["topic"] = txt
        st["step"] = "count"
        rows = [[(f"{c} slayd", f"n:{c}") for c in SLIDE_COUNTS[:2]], [(f"{c} slayd", f"n:{c}") for c in SLIDE_COUNTS[2:]]]
        bot.send_message(m.chat.id, "Nechta slayd bo'lsin?", reply_markup=ik(rows))
        return
    if st and st["step"] == "author":
        st["author"] = clip(txt, 40)
        return confirm_and_generate(m.chat.id, uid)
    bot.send_message(m.chat.id, "Menyudan tanlang.", reply_markup=main_kb())


@bot.callback_query_handler(func=lambda c: True)
def on_cb(c):
    uid = c.from_user.id
    d = c.data or ""
    cid = c.message.chat.id
    mid = c.message.message_id
    try:
        bot.answer_callback_query(c.id)
    except Exception:
        pass
    ensure_user(c.from_user)
    u = get_user(uid)
    if u and u["banned"] and not is_admin(uid):
        return
    st = states.get(uid)

    # ---- taqdimot oqimi
    if d.startswith("n:") and st and st["step"] == "count":
        st["count"] = int(d[2:])
        st["step"] = "lang"
        bot.edit_message_text("Taqdimot tilini tanlang:", cid, mid,
                              reply_markup=ik([[("O'zbekcha", "l:uz"), ("Русский", "l:ru"), ("English", "l:en")]]))
        return
    if d.startswith("l:") and st and st["step"] == "lang":
        st["lang"] = d[2:]
        st["step"] = "theme"
        keys = list(THEMES)
        rows = [[(THEMES[k]["name"], f"t:{k}") for k in keys[i:i + 2]] for i in range(0, len(keys), 2)]
        bot.edit_message_text("Dizayn mavzusini tanlang:", cid, mid, reply_markup=ik(rows))
        return
    if d.startswith("t:") and st and st["step"] == "theme":
        st["theme"] = d[2:]
        st["step"] = "author"
        bot.edit_message_text("Muallif ismini yozing (titul va yakuniy slaydda chiqadi) yoki o'tkazib yuboring.",
                              cid, mid, reply_markup=ik([[("O'tkazib yuborish", "a:skip")]]))
        return
    if d == "a:skip" and st and st["step"] == "author":
        st["author"] = ""
        try:
            bot.delete_message(cid, mid)
        except Exception:
            pass
        return confirm_and_generate(cid, uid)

    # ---- to'lov
    if d.startswith("pk:"):
        pk = next((p for p in PACKAGES if p["id"] == d[3:]), None)
        if not pk:
            return
        states[uid] = {"step": "receipt", "pk": pk}
        bot.send_message(cid,
                         f"<b>{pk['title']}</b>\nTo'lov summasi: <b>{money(pk['price'])}</b>\n\n"
                         f"Karta: <code>{CARD_NUMBER}</code>\nQabul qiluvchi: <b>{e(CARD_OWNER)}</b>\n\n"
                         "To'lovni amalga oshirib, <b>chek rasmini (skrinshot)</b> shu yerga yuboring. "
                         "Admin tekshirgach hisobingiz to'ldiriladi.", reply_markup=cancel_kb())
        return
    if d.startswith(("ap:", "rj:")) and is_admin(uid):
        return payment_decision(c, d)

    # ---- admin
    if d.startswith("adm_") and is_admin(uid):
        return admin_cb(c, d)


def confirm_and_generate(chat_id, uid):
    st = states.pop(uid, None)
    if not st:
        return
    if not take_credit(uid):
        bot.send_message(chat_id, "Hisobingizda taqdimot qolmagan.", reply_markup=main_kb())
        return
    bot.send_message(chat_id, "Qabul qilindi. Taqdimot tayyorlanmoqda, taxminan 1-2 daqiqa kuting...",
                     reply_markup=main_kb())
    threading.Thread(target=generate_job, args=(chat_id, uid, st), daemon=True).start()


def generate_job(chat_id, uid, st):
    status = bot.send_message(chat_id, "1/3  Mazmun yozilmoqda...")
    with gen_sem:
        try:
            data = ai_content(st["topic"], st["count"], st["lang"])
            bot.edit_message_text("2/3  Dizayn va rasmlar tayyorlanmoqda...", chat_id, status.message_id)
            images = fetch_images(data)
            bot.edit_message_text("3/3  Fayl yig'ilmoqda...", chat_id, status.message_id)
            buf = build_pptx(data, st["theme"], st.get("author", ""), st["lang"], images)
            name = re.sub(r"[^\w\-]+", "_", data.get("title", "taqdimot"))[:40].strip("_") or "taqdimot"
            buf.name = f"{name}.pptx"
            u = get_user(uid)
            bot.send_document(chat_id, buf, caption=f"<b>{e(data.get('title', ''))}</b>\nTaqdimot tayyor. "
                                                    f"Qolgan hisob: {u['credits']} ta.")
            dbx("UPDATE users SET made=made+1 WHERE id=?", (uid,))
            try:
                bot.delete_message(chat_id, status.message_id)
            except Exception:
                pass
        except Exception:
            log.error(traceback.format_exc())
            add_credits(uid, 1)
            try:
                bot.edit_message_text("Taqdimot yaratishda xatolik bo'ldi. Hisobingiz qaytarildi, birozdan so'ng qayta urinib ko'ring.",
                                      chat_id, status.message_id)
            except Exception:
                pass
            for a in ADMIN_IDS:
                try:
                    bot.send_message(a, f"Xatolik (user {uid}):\n<code>{e(traceback.format_exc()[-600:])}</code>")
                except Exception:
                    pass


# ----------------------------------------------------------------------------
#                                  TO'LOV (CHEK)
# ----------------------------------------------------------------------------
@bot.message_handler(content_types=["photo"])
def on_photo(m):
    uid = m.from_user.id
    ensure_user(m.from_user)
    if not guard(m):
        return
    st = states.get(uid)
    if not st or st["step"] != "receipt":
        bot.send_message(m.chat.id, "Rasm qabul qilinmadi. To'lov qilish uchun avval tarifni tanlang.", reply_markup=main_kb())
        return
    pk = st["pk"]
    fid = m.photo[-1].file_id
    cur = dbx("INSERT INTO payments(user_id,package,credits,amount,file_id,created) VALUES(?,?,?,?,?,?)",
              (uid, pk["title"], pk["credits"], pk["price"], fid, now()))
    pid = cur.lastrowid
    states.pop(uid, None)
    bot.send_message(m.chat.id, "Chek qabul qilindi. Admin tekshirgach sizga xabar beramiz.", reply_markup=main_kb())
    cap = (f"<b>Yangi to'lov #{pid}</b>\nFoydalanuvchi: {e(m.from_user.first_name)} "
           f"(@{e(m.from_user.username or '-')}) <code>{uid}</code>\nTarif: {pk['title']}\nSumma: {money(pk['price'])}")
    kb = ik([[("Tasdiqlash", f"ap:{pid}"), ("Rad etish", f"rj:{pid}")]])
    for a in ADMIN_IDS:
        try:
            bot.send_photo(a, fid, caption=cap, reply_markup=kb)
        except Exception:
            log.warning("admin %s ga yuborilmadi", a)


def payment_decision(c, d):
    pid = int(d[3:])
    rows = dbx("SELECT * FROM payments WHERE id=?", (pid,), fetch=True)
    if not rows:
        return
    p = rows[0]
    if p["status"] != "pending":
        bot.answer_callback_query(c.id, f"Bu to'lov allaqachon: {p['status']}", show_alert=True)
        return
    ok = d.startswith("ap:")
    dbx("UPDATE payments SET status=? WHERE id=?", ("approved" if ok else "rejected", pid))
    if ok:
        add_credits(p["user_id"], p["credits"])
        msg = f"To'lovingiz tasdiqlandi. Hisobingizga <b>{p['credits']}</b> ta taqdimot qo'shildi."
    else:
        msg = "To'lovingiz tasdiqlanmadi. Chek noto'g'ri bo'lsa, qayta yuboring yoki adminga yozing."
    try:
        bot.send_message(p["user_id"], msg)
    except Exception:
        pass
    mark = "TASDIQLANDI" if ok else "RAD ETILDI"
    try:
        bot.edit_message_caption(f"{c.message.caption or ''}\n\n<b>{mark}</b> (admin {c.from_user.id})",
                                 c.message.chat.id, c.message.message_id)
    except Exception:
        pass


# ----------------------------------------------------------------------------
#                                  ADMIN PANEL
# ----------------------------------------------------------------------------
def admin_kb():
    return ik([
        [("Statistika", "adm_stats"), ("Kutilayotgan to'lovlar", "adm_pend")],
        [("Taqdimot berish", "adm_give"), ("Xabar yuborish", "adm_bc")],
        [("Bloklash", "adm_ban"), ("Blokdan chiqarish", "adm_unban")],
        [("Foydalanuvchini qidirish", "adm_find")],
    ])


def admin_cb(c, d):
    uid, cid = c.from_user.id, c.message.chat.id
    if d == "adm_stats":
        total = dbx("SELECT COUNT(*) n, COALESCE(SUM(made),0) m FROM users", fetch=True)[0]
        today = dbx("SELECT COUNT(*) n FROM users WHERE joined LIKE ?", (datetime.now().strftime("%Y-%m-%d") + "%",), fetch=True)[0]["n"]
        inc = dbx("SELECT COALESCE(SUM(amount),0) s, COUNT(*) n FROM payments WHERE status='approved'", fetch=True)[0]
        pend = dbx("SELECT COUNT(*) n FROM payments WHERE status='pending'", fetch=True)[0]["n"]
        bot.send_message(cid, f"<b>Statistika</b>\n\nFoydalanuvchilar: <b>{total['n']}</b> (bugun +{today})\n"
                              f"Yaratilgan taqdimotlar: <b>{total['m']}</b>\nTasdiqlangan to'lovlar: <b>{inc['n']}</b> ta, "
                              f"jami <b>{money(inc['s'])}</b>\nKutilayotgan to'lovlar: <b>{pend}</b>")
    elif d == "adm_pend":
        rows = dbx("SELECT * FROM payments WHERE status='pending' ORDER BY id LIMIT 20", fetch=True)
        if not rows:
            bot.send_message(cid, "Kutilayotgan to'lov yo'q.")
        for p in rows:
            bot.send_photo(cid, p["file_id"], caption=f"<b>To'lov #{p['id']}</b>\nFoydalanuvchi: <code>{p['user_id']}</code>\n"
                                                      f"Tarif: {p['package']}\nSumma: {money(p['amount'])}",
                           reply_markup=ik([[("Tasdiqlash", f"ap:{p['id']}"), ("Rad etish", f"rj:{p['id']}")]]))
    elif d == "adm_give":
        states[uid] = {"step": "adm_give"}
        bot.send_message(cid, "Format: <code>ID miqdor</code>\nMasalan: <code>123456789 5</code> (manfiy son ayiradi)")
    elif d == "adm_bc":
        states[uid] = {"step": "adm_bc"}
        bot.send_message(cid, "Hammaga yuboriladigan xabar matnini yozing (HTML belgilari ishlaydi). Bekor qilish: " + BTN_CANCEL)
    elif d == "adm_ban":
        states[uid] = {"step": "adm_ban"}
        bot.send_message(cid, "Bloklanadigan foydalanuvchi ID sini yuboring.")
    elif d == "adm_unban":
        states[uid] = {"step": "adm_unban"}
        bot.send_message(cid, "Blokdan chiqariladigan foydalanuvchi ID sini yuboring.")
    elif d == "adm_find":
        states[uid] = {"step": "adm_find"}
        bot.send_message(cid, "Foydalanuvchi ID sini yuboring.")


def admin_state(m, st, txt):
    uid, cid, step = m.from_user.id, m.chat.id, st["step"]
    states.pop(uid, None)
    try:
        if step == "adm_give":
            a, b = txt.split()
            add_credits(int(a), int(b))
            bot.send_message(cid, f"Bajarildi: {a} hisobiga {b} ta qo'shildi.")
            if int(b) > 0:
                try:
                    bot.send_message(int(a), f"Hisobingizga <b>{int(b)}</b> ta taqdimot qo'shildi.")
                except Exception:
                    pass
        elif step in ("adm_ban", "adm_unban"):
            dbx("UPDATE users SET banned=? WHERE id=?", (1 if step == "adm_ban" else 0, int(txt)))
            bot.send_message(cid, "Bajarildi.")
        elif step == "adm_find":
            u = get_user(int(txt))
            if not u:
                bot.send_message(cid, "Topilmadi.")
            else:
                bot.send_message(cid, f"<b>{e(u['name'])}</b> @{e(u['username'] or '-')}\nID: <code>{u['id']}</code>\n"
                                      f"Hisob: {u['credits']}\nYaratgan: {u['made']}\nBlok: {'ha' if u['banned'] else "yo'q"}\nQo'shilgan: {u['joined']}")
        elif step == "adm_bc":
            ids = [r["id"] for r in dbx("SELECT id FROM users WHERE banned=0", fetch=True)]
            bot.send_message(cid, f"Yuborish boshlandi: {len(ids)} ta foydalanuvchi...")
            threading.Thread(target=broadcast, args=(cid, ids, m.text), daemon=True).start()
    except Exception:
        bot.send_message(cid, "Format noto'g'ri. Qayta urinib ko'ring: /admin")


def broadcast(admin_chat, ids, text):
    ok = bad = 0
    for i in ids:
        try:
            bot.send_message(i, text)
            ok += 1
        except Exception:
            bad += 1
        time.sleep(0.05)
    bot.send_message(admin_chat, f"Xabar yuborildi. Muvaffaqiyatli: {ok}, xato: {bad}")


# ----------------------------------------------------------------------------
if __name__ == "__main__":
    log.info("Bot ishga tushdi")
    bot.infinity_polling(skip_pending=True, timeout=30, long_polling_timeout=30)
