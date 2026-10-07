import shutil
from datetime import date
from pathlib import Path
from types import SimpleNamespace

import pytest

from expense_bot.reader import ReceiptError, parse, read_receipt
from expense_bot.receipt import Receipt

TEMPLATE = Path(__file__).parent.parent / "data" / "expense_report.xls"


def _rc(**kw):
    base = dict(date=date(2026, 10, 3), vendor="タイムズ24", item="駐車場代", amount=800, category="transport")
    return Receipt(**{**base, **kw})


def test_parse_ok_and_unit_price_dropped_when_inconsistent():
    d = {"is_receipt": True, "date": "2026-10-03", "vendor": " ABC ", "item": "ボンド", "amount": 786,
         "category": "site", "unit_price": 262, "quantity": 3}
    rc = parse(d)
    assert (rc.vendor, rc.mark, rc.unit_price, rc.quantity) == ("ABC", "経", 262, 3)
    rc = parse({**d, "amount": 800})  # 262*3 != 800
    assert rc.unit_price is None and rc.quantity is None


@pytest.mark.parametrize("bad", [{"is_receipt": False}, {"date": "x", "amount": 1}, {"date": "2026-01-01", "amount": 0}])
def test_parse_rejects(bad):
    with pytest.raises(ReceiptError):
        parse({"vendor": "a", "item": "b", "category": "site", **bad})


def test_read_receipt_with_fake_client(tmp_path):
    from PIL import Image

    img = tmp_path / "r.png"
    Image.new("RGB", (3000, 2000), "white").save(img)
    block = SimpleNamespace(type="tool_use", input={"is_receipt": True, "date": "2026-10-03", "vendor": "X",
                                                    "item": "駐車場代", "amount": 400, "category": "transport"})
    seen = {}

    class Msgs:
        def create(self, **kw):
            seen.update(kw)
            return SimpleNamespace(content=[block])

    rc = read_receipt(img, client=SimpleNamespace(messages=Msgs()))
    assert rc.amount == 400 and rc.source == "r.png"
    assert seen["tool_choice"]["name"] == "record_receipt"


@pytest.mark.skipif(not (shutil.which("soffice") and TEMPLATE.exists()), reason="LibreOffice / template が必要")
def test_workbook_fill_and_new_sheet(tmp_path):
    pytest.importorskip("uno")
    from expense_bot.workbook import Duplicate, ExpenseWorkbook, FIRST_ROW, LAST_ROW

    book = tmp_path / "book.xls"
    shutil.copy(TEMPLATE, book)
    with ExpenseWorkbook(book) as wb:
        sheets_before = wb.sheets.Count
        last = wb.current_sheet().Name
        p = wb.add(_rc())
        assert (p.sheet, p.row, p.new_sheet) == (last, 11, False)  # 原簿(6) は10行目まで使用済み
        with pytest.raises(Duplicate):
            wb.add(_rc())
        # 単価×数量は式で入る
        wb.add(_rc(vendor="ホームセンター", item="ボンド", amount=786, category="site", unit_price=262, quantity=3))
        assert wb.current_sheet().getCellByPosition(25, 11).Value == 786
        # 15行を埋めると新しいシートへ
        for i in range(LAST_ROW - FIRST_ROW + 1 - 3):  # 既存1+追加2 で3行使用済み
            wb.add(_rc(vendor=f"店{i}", amount=100 + i))
        p = wb.add(_rc(vendor="溢れ", amount=999))
        assert p.new_sheet and p.row == 10 and wb.sheets.Count == sheets_before + 1
        assert wb.current_sheet().getCellByPosition(5, 10).String == ""  # 新シートの2行目は空
    assert book.with_name("book.xls.bak").exists()
    # 保存後に開き直して確認
    with ExpenseWorkbook(book, backup=False) as wb:
        s = wb.current_sheet()
        assert s.getCellByPosition(5, 9).String == "溢れ" and s.getCellByPosition(25, 9).Value == 999
