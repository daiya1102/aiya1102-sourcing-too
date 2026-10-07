"""経費精算書(.xls/.xlsx)への転記。

openpyxl 等では .xls を扱えず、印・丸囲み・注釈などの図形も失われるため、
LibreOffice(UNO)でブックを直接開き、元の書式・図形を保ったまま書き込む。
"""
from __future__ import annotations

import re
import shutil
import socket
import subprocess
import tempfile
import time
from dataclasses import dataclass
from datetime import date
from pathlib import Path

from .receipt import Receipt

# 原簿シートのレイアウト(0始まり)。明細は 10〜24 行目の15行
FIRST_ROW, LAST_ROW = 9, 23
COL_DATE, COL_KIND, COL_VENDOR, COL_ITEM, COL_PURPOSE = 0, 2, 5, 10, 15
COL_RECEIPT_YES, COL_UNIT, COL_QTY, COL_TOTAL = 20, 22, 24, 25
KIND_COLS = {"交": 2, "経": 3, "際": 4}
APPLY_DATE = "R5"
TEMPLATE_SHEET = "記入例"
CIRCLE_COLOR = 0xFF0000
FILTERS = {".xls": "MS Excel 97", ".xlsx": "Calc MS Excel 2007 XML"}


class WorkbookError(Exception):
    pass


@dataclass
class Placed:
    sheet: str
    row: int  # 1始まりのExcel行番号
    new_sheet: bool = False


class Duplicate(WorkbookError):
    pass


def _serial(d: date) -> int:
    return (d - date(1899, 12, 30)).days


def _wareki(d: date) -> str:
    return f"令和{d.year - 2018}年{d.month}月{d.day}日"


def _free_port() -> int:
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


class ExpenseWorkbook:
    def __init__(self, path: Path, *, backup: bool = True):
        self.path = Path(path).resolve()
        if self.path.suffix.lower() not in FILTERS:
            raise WorkbookError("対応形式は .xls / .xlsx です")
        if not self.path.exists():
            raise WorkbookError(f"ファイルがありません: {self.path}")
        self.backup = backup
        self.dirty = False

    # ---- LibreOffice の起動・終了 ----
    def __enter__(self) -> "ExpenseWorkbook":
        import uno  # LibreOffice 同梱の python3-uno が必要
        from com.sun.star.beans import PropertyValue

        self._uno = uno
        binary = shutil.which("soffice") or shutil.which("libreoffice")
        if not binary:
            raise WorkbookError("LibreOffice が見つかりません(soffice をインストールしてください)")
        port = _free_port()
        self._profile = tempfile.mkdtemp(prefix="expense_lo_")  # 利用者の LibreOffice と干渉しない専用プロファイル
        self._proc = subprocess.Popen(
            [binary, "--headless", "--invisible", "--norestore", "--nologo",
             f"--accept=socket,host=127.0.0.1,port={port};urp;",
             f"-env:UserInstallation={Path(self._profile).as_uri()}"],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        ctx = uno.getComponentContext()
        resolver = ctx.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", ctx)
        for _ in range(90):
            try:
                remote = resolver.resolve(f"uno:socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext")
                break
            except Exception:
                time.sleep(1)
        else:
            self._shutdown()
            raise WorkbookError("LibreOffice に接続できませんでした")
        self._desktop = remote.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", remote)
        hidden = PropertyValue()
        hidden.Name, hidden.Value = "Hidden", True
        self.doc = self._desktop.loadComponentFromURL(uno.systemPathToFileUrl(str(self.path)), "_blank", 0, (hidden,))
        if self.doc is None:
            self._shutdown()
            raise WorkbookError("ブックを開けませんでした")
        return self

    def __exit__(self, exc_type, exc, tb):
        try:
            if exc_type is None and self.dirty:
                self._save()
        finally:
            try:
                self.doc.close(True)
            except Exception:
                pass
            self._shutdown()

    def _shutdown(self):
        try:
            self._desktop.terminate()
        except Exception:
            pass
        try:
            self._proc.wait(timeout=15)
        except Exception:
            self._proc.kill()
        shutil.rmtree(self._profile, ignore_errors=True)

    def _save(self):
        from com.sun.star.beans import PropertyValue

        if self.backup:
            shutil.copy2(self.path, self.path.with_name(self.path.name + ".bak"))
        f = PropertyValue()
        f.Name, f.Value = "FilterName", FILTERS[self.path.suffix.lower()]
        self.doc.storeToURL(self._uno.systemPathToFileUrl(str(self.path)), (f,))

    # ---- シート操作 ----
    @property
    def sheets(self):
        return self.doc.Sheets

    def current_sheet(self):
        """最新の原簿シート(末尾のシート。『記入例』は対象外)。"""
        names = [n for n in self.sheets.ElementNames if n != TEMPLATE_SHEET]
        if not names:
            raise WorkbookError("原簿シートがありません")
        return self.sheets.getByName(names[-1])

    def _row_used(self, sheet, r: int) -> bool:
        return any(sheet.getCellByPosition(c, r).String.strip()
                   for c in (COL_DATE, COL_VENDOR, COL_ITEM, COL_PURPOSE, COL_TOTAL))

    def _free_row(self, sheet) -> int | None:
        return next((r for r in range(FIRST_ROW, LAST_ROW + 1) if not self._row_used(sheet, r)), None)

    def _data_band(self, sheet) -> tuple[int, int]:
        top = sheet.getCellByPosition(0, FIRST_ROW).Position.Y
        last = sheet.getCellByPosition(0, LAST_ROW)
        return top, last.Position.Y + last.Size.Height

    def _clear_circles(self, sheet):
        """明細行にある丸囲み(中身が空の図形)を消す。ヘッダの丸や注釈は残す。"""
        top, bottom = self._data_band(sheet)
        page = sheet.DrawPage
        for i in reversed(range(page.Count)):
            s = page.getByIndex(i)
            if (s.ShapeType.endswith(("CustomShape", "EllipseShape")) and not s.String.strip()
                    and s.Size.Width < 1500 and top <= s.Position.Y < bottom):
                page.remove(s)

    def _new_sheet(self, base):
        last = self.sheets.ElementNames[-1]
        m = re.search(r"\((\d+)\)\s*$", last)
        n = int(m.group(1)) + 1 if m else 2
        name = f"原簿 ({n})"
        while self.sheets.hasByName(name):
            n += 1
            name = f"原簿 ({n})"
        self.sheets.copyByName(base.Name, name, self.sheets.Count)
        sheet = self.sheets.getByName(name)
        for r in range(FIRST_ROW, LAST_ROW + 1):
            for c in (COL_DATE, COL_VENDOR, COL_ITEM, COL_PURPOSE, COL_UNIT, COL_QTY, COL_TOTAL, 28):
                cell = sheet.getCellByPosition(c, r)
                cell.setString("")
        for r in range(27, 36):  # 前回の申請メモ(「立替えております…」等)を消す。27行目の社名・合計は残す
            for c in range(0, 30):
                cell = sheet.getCellByPosition(c, r)
                if cell.Type.value == "TEXT":
                    cell.setString("")
        self._clear_circles(sheet)
        return sheet

    # ---- 書き込み ----
    def _circle(self, sheet, cell):
        from com.sun.star.awt import Point, Size
        from com.sun.star.drawing.FillStyle import NONE

        shape = self.doc.createInstance("com.sun.star.drawing.EllipseShape")
        sheet.DrawPage.add(shape)
        p, s = cell.Position, cell.Size
        dx, dy = s.Width // 14, s.Height // 14
        shape.setPosition(Point(p.X + dx, p.Y + dy))
        shape.setSize(Size(s.Width - 2 * dx, s.Height - 2 * dy))
        shape.FillStyle = NONE
        shape.LineColor = CIRCLE_COLOR
        shape.LineWidth = 25
        shape.Anchor = cell

    def _find_duplicate(self, sheet, rc: Receipt) -> bool:
        for r in range(FIRST_ROW, LAST_ROW + 1):
            d = sheet.getCellByPosition(COL_DATE, r)
            if (d.Value == _serial(rc.date) and sheet.getCellByPosition(COL_TOTAL, r).Value == rc.amount
                    and sheet.getCellByPosition(COL_VENDOR, r).String.strip() == rc.vendor):
                return True
        return False

    def add(self, rc: Receipt, *, new_sheet: bool = False, today: date | None = None) -> Placed:
        sheet = self.current_sheet()
        if self._find_duplicate(sheet, rc):
            raise Duplicate(f"登録済みの可能性があります({rc.date} {rc.vendor} {rc.amount:,}円)")
        created = False
        row = None if new_sheet else self._free_row(sheet)
        if row is None:
            sheet, created = self._new_sheet(sheet), True
            row = FIRST_ROW

        cell = lambda c: sheet.getCellByPosition(c, row)
        cell(COL_DATE).setValue(_serial(rc.date))
        for mark, col in KIND_COLS.items():
            cell(col).setString(mark)
        cell(COL_VENDOR).setString(rc.vendor)
        cell(COL_ITEM).setString(rc.item)
        cell(COL_PURPOSE).setString(rc.purpose)
        cell(COL_RECEIPT_YES).setString("有")
        cell(COL_RECEIPT_YES + 1).setString("無")
        if rc.unit_price and rc.quantity:  # 現場経費の単価は税込。合計は 単価×数量 の式
            cell(COL_UNIT).setValue(rc.unit_price)
            cell(COL_QTY).setValue(rc.quantity)
            cell(COL_TOTAL).setFormula(f"=W{row + 1}*Y{row + 1}")
        else:
            cell(COL_TOTAL).setValue(rc.amount)
        self._circle(sheet, cell(KIND_COLS[rc.mark]))
        self._circle(sheet, cell(COL_RECEIPT_YES))

        today = today or date.today()
        ap = sheet.getCellRangeByName(APPLY_DATE)
        if ap.Type.value == "VALUE" or ap.getFormula().startswith("="):
            ap.setValue(_serial(today))
        else:
            ap.setString(_wareki(today))
        self.dirty = True
        return Placed(sheet.Name, row + 1, created)

    def set_applicant(self, name: str):
        self.current_sheet().getCellRangeByName("C7").setString(name)
        self.dirty = True
