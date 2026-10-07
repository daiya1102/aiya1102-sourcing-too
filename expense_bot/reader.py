"""領収書画像 → Receipt。Claude のビジョンで読み取る。"""
from __future__ import annotations

import base64
import io
import os
from datetime import date, datetime
from pathlib import Path

from .receipt import CATEGORIES, Receipt

DEFAULT_MODEL = "claude-sonnet-5-5"
IMAGE_SUFFIXES = {".jpg", ".jpeg", ".png", ".webp", ".gif", ".pdf"}
MAX_EDGE = 2000  # API の容量制限対策。レシートの文字は十分読める大きさ

TOOL = {
    "name": "record_receipt",
    "description": "領収書・レシートから読み取った内容を記録する",
    "input_schema": {
        "type": "object",
        "properties": {
            "is_receipt": {"type": "boolean", "description": "領収書・レシート・切符等の支払証憑なら true"},
            "date": {"type": "string", "description": "支払日 YYYY-MM-DD。和暦は西暦に直す。年が無ければ今年"},
            "vendor": {"type": "string", "description": "店名・会社名(株式会社等も含め印字どおり)"},
            "item": {
                "type": "string",
                "description": "物品名・用件を短く(例: 駐車場代 / 有料道路代 / 消耗品 / ガソリン代 / 会食費)",
            },
            "amount": {"type": "integer", "description": "支払総額(税込・円)。ポイント値引き後の実支払額"},
            "category": {
                "type": "string",
                "enum": list(CATEGORIES),
                "description": "transport=交通費(電車・バス・タクシー・駐車場・有料道路・ガソリン) / "
                "site=現場経費(消耗品・材料・工具・備品) / entertainment=交際費(接待飲食・会費・慶弔・贈答)",
            },
            "unit_price": {"type": "integer", "description": "明細が1品のみで数量が2以上のときの税込単価"},
            "quantity": {"type": "integer", "description": "同上の数量"},
            "notes": {"type": "string", "description": "読み取りに自信がない点があれば簡潔に"},
        },
        "required": ["is_receipt", "date", "vendor", "item", "amount", "category"],
    },
}

SYSTEM = (
    "あなたは経理担当です。送られた領収書・レシートの画像から、経費精算書に転記する項目を読み取り、"
    "record_receipt を呼んでください。金額は必ず税込の支払総額(お預り・お釣りではなく『合計』)。"
    "読めない項目を推測で作らないこと(自信がなければ notes に書く)。"
)


class ReceiptError(Exception):
    pass


def _prepare(path: Path) -> tuple[str, str, str]:
    """(block type, media type, base64) を返す。大きい画像は縮小。"""
    suffix = path.suffix.lower()
    raw = path.read_bytes()
    if suffix == ".pdf":
        return "document", "application/pdf", base64.standard_b64encode(raw).decode()
    from PIL import Image, ImageOps

    img = ImageOps.exif_transpose(Image.open(io.BytesIO(raw)))  # スマホ写真の回転を補正
    if max(img.size) > MAX_EDGE:
        img.thumbnail((MAX_EDGE, MAX_EDGE))
    buf = io.BytesIO()
    img.convert("RGB").save(buf, "JPEG", quality=88)
    return "image", "image/jpeg", base64.standard_b64encode(buf.getvalue()).decode()


def read_receipt(path: Path, *, client=None, model: str | None = None, today: date | None = None) -> Receipt:
    if client is None:
        import anthropic

        if not os.environ.get("ANTHROPIC_API_KEY"):
            raise ReceiptError("環境変数 ANTHROPIC_API_KEY が設定されていません")
        client = anthropic.Anthropic()
    today = today or date.today()
    kind, media, data = _prepare(path)
    msg = client.messages.create(
        model=model or os.environ.get("EXPENSE_MODEL", DEFAULT_MODEL),
        max_tokens=1024,
        system=SYSTEM,
        tools=[TOOL],
        tool_choice={"type": "tool", "name": TOOL["name"]},
        messages=[
            {
                "role": "user",
                "content": [
                    {"type": kind, "source": {"type": "base64", "media_type": media, "data": data}},
                    {"type": "text", "text": f"今日は {today.isoformat()} です。この領収書を読み取ってください。"},
                ],
            }
        ],
    )
    block = next((b for b in msg.content if b.type == "tool_use"), None)
    if block is None:
        raise ReceiptError("読み取り結果が得られませんでした")
    return parse(block.input, source=path.name, today=today)


def parse(d: dict, *, source: str = "", today: date | None = None) -> Receipt:
    if not d.get("is_receipt", True):
        raise ReceiptError("領収書として認識できませんでした")
    try:
        paid = datetime.strptime(d["date"], "%Y-%m-%d").date()
        amount = int(d["amount"])
    except (KeyError, ValueError, TypeError) as e:
        raise ReceiptError(f"日付または金額を読み取れませんでした: {e}") from e
    if amount <= 0:
        raise ReceiptError("金額を読み取れませんでした")
    if d.get("category") not in CATEGORIES:
        raise ReceiptError(f"種別が不明です: {d.get('category')!r}")
    qty, unit = d.get("quantity"), d.get("unit_price")
    if not (qty and unit and qty > 1 and qty * unit == amount):
        qty = unit = None  # 単価×数量が合計と一致しないときは合計だけ書く
    return Receipt(
        date=paid,
        vendor=d["vendor"].strip(),
        item=d["item"].strip(),
        amount=amount,
        category=d["category"],
        unit_price=unit,
        quantity=qty,
        source=source,
    )
