import os
import asyncio
from io import BytesIO

from aiogram import Bot, Dispatcher, types, F
from aiogram.filters import CommandStart
from aiogram.types import Message
from aiohttp import web

import pandas as pd

from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.lib.colors import black

from PyPDF2 import PdfReader, PdfWriter
import fitz  # PyMuPDF



# ================== НАСТРОЙКИ ==================
TOKEN = os.getenv("TELEGRAM_TOKEN")
bot = Bot(token=TOKEN)
dp = Dispatcher()

# Шрифт: или встроенный Helvetica-Bold, или подключаем TTF.
FONT_NAME = "Helvetica-Bold"
FONT_TTF  = None  # например: "fonts/DejaVuSans-Bold.ttf"
TEXT_COLOR = black

# Прямоугольник под словом PROMOCODE (в долях ширины/высоты страницы)
BOX = dict(
    x=0.10,   # ширину не трогаем
    y=0.150,  # было 0.37 → опустили ниже (если нужно ещё ниже — уменьшайте)
    w=0.80,
    h=0.050   # было 0.13 → делаем ниже бокс, значит и шрифт станет меньше
)

# Память на пользователя: последние загруженные Template и Codes
USER_STATE = {}  # user_id -> {"template": bytes, "codes": [..]}

# =================================================


# ---------- сервис: регистрация шрифта -----------
def ensure_font():
    global FONT_NAME
    if FONT_TTF:
        try:
            pdfmetrics.registerFont(TTFont("CustomFont", FONT_TTF))
            FONT_NAME = "CustomFont"
        except Exception as e:
            print(f"[FONT] Не удалось подключить TTF ({FONT_TTF}): {e}. Использую Helvetica-Bold.")


# ---------- автоподбор размера шрифта ------------
def fit_font_size(text, box_w_pt, box_h_pt, font_name):
    lo, hi = 4, 300
    best = lo
    while lo <= hi:
        mid = (lo + hi) // 2
        width = stringWidth(text, font_name, mid)
        if width <= box_w_pt and mid <= box_h_pt:
            best = mid
            lo = mid + 1
        else:
            hi = mid - 1
    return best


# ---------- рисуем одну страницу-оверлей ---------
def make_overlay(page_w, page_h, code_text):
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=(page_w, page_h))
    c.setFillColor(TEXT_COLOR)
    c.setStrokeColor(TEXT_COLOR)

    # прямоугольник в пунктах
    box_x = BOX["x"] * page_w
    box_y = BOX["y"] * page_h
    box_w = BOX["w"] * page_w
    box_h = BOX["h"] * page_h

    font_size = fit_font_size(code_text, box_w, box_h, FONT_NAME)
    c.setFont(FONT_NAME, font_size)

    text_w = stringWidth(code_text, FONT_NAME, font_size)
    x = box_x + (box_w - text_w) / 2.0
    y = box_y + (box_h - font_size) / 2.0

    c.drawString(x, y, code_text)
    c.showPage()
    c.save()
    buf.seek(0)
    return PdfReader(buf).pages[0]


# ---------- сборка итогового PDF ------------------
def build_pdf(template_bytes: bytes, codes: list[str]) -> bytes:
    """
    Надёжная сборка через PyMuPDF:
    - на каждую страницу выводим шаблон;
    - поверх выводим промокод, подгоняя размер шрифта в BOX.
    """
    tpl = fitz.open(stream=template_bytes, filetype="pdf")
    if tpl.page_count == 0:
        raise ValueError("Пустой PDF-шаблон")

    page0 = tpl[0]
    page_w, page_h = page0.rect.width, page0.rect.height

    out = fitz.open()

    # Подготовим шрифт
    if FONT_TTF:
        # регистрируем TTF и используем его
        fontname = out.insert_font(file=FONT_TTF)
    else:
        # встроенный Helvetica (helv)
        fontname = "helv"

    def fit_font_size_fitz(text, bw, bh):
        lo, hi = 4, 300
        best = lo
        while lo <= hi:
            mid = (lo + hi) // 2
            width = out.get_text_length(text, fontname=fontname, fontsize=mid)
            if width <= bw and mid <= bh:
                best = mid
                lo = mid + 1
            else:
                hi = mid - 1
        return best

    # координаты «бокса» под код
    box_x = BOX["x"] * page_w
    box_y = BOX["y"] * page_h
    box_w = BOX["w"] * page_w
    box_h = BOX["h"] * page_h

    for code in codes:
        # 1) новая страница нужного размера
        p = out.new_page(width=page_w, height=page_h)
        # 2) рисуем на неё шаблон
        p.show_pdf_page(p.rect, tpl, 0)

        # 3) рассчитываем размер шрифта и позицию
        text = str(code).strip()
        fs = fit_font_size_fitz(text, box_w, box_h)
        text_w = out.get_text_length(text, fontname=fontname, fontsize=fs)

        x = box_x + (box_w - text_w) / 2.0
        # в PyMuPDF y — это БАЗОВАЯ линия; чтобы центрировать по высоте «бокса»:
        y = box_y + (box_h + fs) / 2.0

        # 4) печатаем код (чёрным)
        p.insert_text(
            fitz.Point(x, y),
            text,
            fontname=fontname,
            fontsize=fs,
            color=(0, 0, 0),  # black
        )

    return out.tobytes()


# ---------- чтение Excel / CSV --------------------
def read_codes_from_bytes(file_name: str, raw: bytes) -> list[str]:
    name = (file_name or "").lower()
    if name.endswith(".csv"):
        df = pd.read_csv(BytesIO(raw), dtype=str)
    else:
        # xlsx/xls/неизвестно — пробуем как Excel
        df = pd.read_excel(BytesIO(raw), dtype=str)

    # ищем столбец code/код или берём первый непустой
    cols = [c for c in df.columns]
    col_name = None
    for cand in cols:
        lc = str(cand).strip().lower()
        if lc in {"code", "код", "promocode", "promo", "коды", "codes"}:
            col_name = cand
            break
    if col_name is None:
        # берём первый столбец
        col_name = cols[0]

    series = df[col_name].dropna().astype(str).str.strip()
    codes = [s for s in series.tolist() if s]
    # убираем дубликаты, сохраняя порядок
    seen = set()
    uniq = []
    for c in codes:
        if c not in seen:
            seen.add(c)
            uniq.append(c)
    return uniq


# ================== HANDLERS ======================
@dp.message(CommandStart())
async def on_start(m: Message):
    await m.answer(
        "Привет! Пришли мне два файла:\n"
        "1) PDF-шаблон (страница с PROMOCODE)\n"
        "2) Excel/CSV с промокодами (столбец code/код или просто первый столбец)\n\n"
        "Как только оба файла получу — соберу и вышлю итоговый PDF."
    )


@dp.message(F.document)
async def on_document(m: Message):
    user_id = m.from_user.id
    doc = m.document
    file_name = doc.file_name or ""

    # скачиваем байты
    buf = BytesIO()
    try:
        # aiogram v3 умеет так:
        await bot.download(doc, destination=buf)
    except Exception:
        # запасной путь через file_path
        try:
            tg_file = await bot.get_file(doc.file_id)
            await bot.download_file(tg_file.file_path, buf)
        except Exception as e:
            await m.reply(f"Не удалось скачать файл: {e}")
            return
    raw = buf.getvalue()

    state = USER_STATE.setdefault(user_id, {"template": None, "codes": None})

    if (doc.mime_type or "").lower() == "application/pdf" or file_name.lower().endswith(".pdf"):
        state["template"] = raw
        await m.reply("Шаблон PDF сохранён ✅")
    elif any(file_name.lower().endswith(ext) for ext in (".xlsx", ".xls", ".csv")):
        try:
            codes = read_codes_from_bytes(file_name, raw)
            if not codes:
                await m.reply("В этом файле не нашёл ни одного кода 😕")
                return
            state["codes"] = codes
            await m.reply(f"Файл с кодами сохранён ✅\nНайдено кодов: {len(codes)}")
        except Exception as e:
            await m.reply(f"Не смог прочитать коды: {e}")
            return
    else:
        await m.reply("Мне нужен PDF (шаблон) или Excel/CSV (коды).")
        return

    # если есть оба — генерим
    if state.get("template") and state.get("codes"):
        await m.reply("Собираю итоговый PDF… ⏳")
        try:
            ensure_font()
            out_bytes = build_pdf(state["template"], state["codes"])
            out = BytesIO(out_bytes)
            out.name = "promo_cards.pdf"
            await m.answer_document(types.BufferedInputFile(out.getvalue(), filename="promo_cards.pdf"))
            # сбрасываем состояние (чтобы следующий запуск был чистым)
            USER_STATE[user_id] = {"template": None, "codes": None}
        except Exception as e:
            await m.reply(f"Упс, не собралось: {e}")


# ============ Веб-сервер для Render Free ==========
async def handle(request):
    return web.Response(text="Bot is alive!")

async def start_web_server():
    app = web.Application()
    app.router.add_get("/", handle)
    runner = web.AppRunner(app)
    await runner.setup()
    site = web.TCPSite(runner, "0.0.0.0", int(os.getenv("PORT", 5000)))
    await site.start()


# ================== ENTRYPOINT ====================
async def main():
    ensure_font()
    await start_web_server()
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
