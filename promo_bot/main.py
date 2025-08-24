import os
import asyncio
import logging
import tempfile
from typing import List

import pandas as pd
from aiogram import Bot, Dispatcher, F, types
from aiogram.filters import CommandStart, Command
from aiogram.types import FSInputFile
from aiohttp import web

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


# ---------- ЛОГИ ----------
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)
# --------------------------

# ---------- TELEGRAM ----------
TOKEN = os.getenv("TELEGRAM_TOKEN")  # из Render → Environment
bot = Bot(token=TOKEN)
dp = Dispatcher()
# ------------------------------


# ---------- /start /help ----------
@dp.message(CommandStart())
async def cmd_start(message: types.Message):
    text = (
        "Привет! Я соберу PDF с промо-картами.\n\n"
        "Отправь мне Excel (.xlsx) c колонкой `code` (или `код`). "
        "Если такой колонки нет — возьму первую колонку.\n\n"
        "Команды:\n"
        "• /help — короткая справка"
    )
    await message.answer(text)


@dp.message(Command("help"))
async def cmd_help(message: types.Message):
    await message.answer(
        "1) Пришли Excel (.xlsx) с кодами (колонка `code`/`код`/первая колонка)\n"
        "2) Я верну PDF: по одной карточке на страницу."
    )
# ----------------------------------


# ---------- ГЕНЕРАЦИЯ PDF ----------
FONT_PATH = os.path.join(os.path.dirname(__file__), "fonts", "DejaVuSans-Bold.ttf")
FONT_NAME = "DejaVuSansBold"

def ensure_font_registered():
    if FONT_NAME not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont(FONT_NAME, FONT_PATH))

def build_pdf(codes: List[str], out_path: str):
    """Простой вариант: 1 код = 1 страница, код по центру."""
    ensure_font_registered()
    c = canvas.Canvas(out_path, pagesize=A4)
    w, h = A4
    c.setFont(FONT_NAME, 48)

    for code in codes:
        code_str = str(code).strip()
        if not code_str:
            continue
        c.drawCentredString(w / 2, h / 2, code_str)
        c.showPage()

    c.save()
# -----------------------------------


# ---------- ПРИЁМ EXCEL ФАЙЛОВ ----------
def pick_codes_column(df: pd.DataFrame) -> List[str]:
    cols_lower = {col.lower(): col for col in df.columns}
    for candidate in ("code", "код"):
        if candidate in cols_lower:
            return df[cols_lower[candidate]].astype(str).fillna("").tolist()
    # если спец-колонок нет — взять первую
    first_col = df.columns[0]
    return df[first_col].astype(str).fillna("").tolist()


@dp.message(F.document)
async def handle_document(message: types.Message):
    doc: types.Document = message.document
    filename = (doc.file_name or "").lower()

    if not filename.endswith(".xlsx"):
        await message.answer("Пришлите, пожалуйста, *Excel* файл с расширением `.xlsx`.", parse_mode="Markdown")
        return

    await message.answer("Файл получил, обрабатываю…")

    # Временная папка на диске
    with tempfile.TemporaryDirectory() as tmpdir:
        src_path = os.path.join(tmpdir, "input.xlsx")
        out_pdf = os.path.join(tmpdir, "promo_cards.pdf")

        # Скачиваем файл
        file = await bot.get_file(doc.file_id)
        await bot.download(file, destination=src_path)

        # Читаем Excel
        try:
            df = pd.read_excel(src_path, dtype=str)
        except Exception as e:
            logger.exception("Ошибка чтения Excel: %s", e)
            await message.answer("Не удалось прочитать Excel. Убедитесь, что это .xlsx и файл не повреждён.")
            return

        if df.empty or len(df.columns) == 0:
            await message.answer("В Excel нет данных. Нужна хотя бы одна колонка с кодами.")
            return

        codes = pick_codes_column(df)
        codes = [c for c in codes if str(c).strip()]
        if not codes:
            await message.answer("Кодов не нашёл. Проверьте колонку `code`/`код` или первую колонку.")
            return

        # Генерируем PDF
        try:
            build_pdf(codes, out_pdf)
        except Exception as e:
            logger.exception("Ошибка генерации PDF: %s", e)
            await message.answer("Что-то пошло не так при сборке PDF 😔")
            return

        # Отправляем пользователю
        await message.answer_document(
            FSInputFile(out_pdf),
            caption=f"Готово! Сформировано {len(codes)} страниц(ы) с кодами.",
        )
# ---------------------------------------


# ---------- ВЕБ-ПИНГ ДЛЯ RENDER ----------
async def ping_handler(_request: web.Request):
    logger.info("✅ Ping received")
    return web.Response(text="Bot is alive!")

async def start_web_server():
    app = web.Application()
    app.router.add_get("/", ping_handler)
    runner = web.AppRunner(app)
    await runner.setup()
    site = web.TCPSite(runner, "0.0.0.0", int(os.getenv("PORT", "5000")))
    await site.start()
# ----------------------------------------


# ---------- ЗАПУСК ----------
async def main():
    await start_web_server()
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
# ---------------------------
