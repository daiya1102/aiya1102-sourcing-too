from __future__ import annotations

import argparse
import shutil
import sys
import time
from pathlib import Path

from .reader import IMAGE_SUFFIXES, ReceiptError, read_receipt
from .receipt import CATEGORIES
from .workbook import Duplicate, ExpenseWorkbook, WorkbookError

DEFAULT_WORKBOOK = Path("data/expense_report.xls")
DEFAULT_INBOX = Path("inbox")


def _purpose_for(img: Path) -> str:
    """領収書と同名の .txt があれば、その1行目を『目的/現場・行程/内容』として使う。"""
    side = img.with_suffix(".txt")
    return side.read_text(encoding="utf8").strip().splitlines()[0] if side.exists() and side.stat().st_size else ""


def process(images: list[Path], args) -> dict[Path, str | None]:
    """画像を読み取り、1回のブック操作でまとめて転記する。戻り値: {画像: エラー文 or None}"""
    result: dict[Path, str | None] = {}
    receipts = []
    for img in images:
        try:
            rc = read_receipt(img)
            rc.purpose = args.purpose or _purpose_for(img)
            if args.category:
                rc.category = args.category
            receipts.append((img, rc))
            print(f"読取: {img.name} → {rc.date:%Y/%m/%d} {rc.vendor} / {rc.item} / {rc.amount:,}円 [{rc.mark}]")
        except ReceiptError as e:
            result[img] = str(e)
            print(f"失敗: {img.name}: {e}", file=sys.stderr)
    if args.dry_run or not receipts:
        result.update({img: None for img, _ in receipts})
        return result
    receipts.sort(key=lambda t: t[1].date)
    with ExpenseWorkbook(args.workbook) as wb:
        if args.applicant:
            wb.set_applicant(args.applicant)
        for i, (img, rc) in enumerate(receipts):
            try:
                at = wb.add(rc, new_sheet=args.new_sheet and i == 0)
                result[img] = None
                note = "(新しいシートを作成)" if at.new_sheet else ""
                print(f"転記: {img.name} → 『{at.sheet}』{at.row}行目 {note}")
                if not rc.purpose:
                    print("  ※ 目的/現場・行程/内容 が空欄です。Excelで入力してください", file=sys.stderr)
            except Duplicate as e:
                result[img] = f"スキップ: {e}"
                print(f"スキップ: {img.name}: {e}", file=sys.stderr)
    return result


def _scan(inbox: Path) -> list[Path]:
    return sorted(p for p in inbox.iterdir() if p.is_file() and p.suffix.lower() in IMAGE_SUFFIXES)


def _archive(img: Path, inbox: Path, error: str | None):
    dest = inbox / ("failed" if error else "processed")
    dest.mkdir(exist_ok=True)
    for f in (img, img.with_suffix(".txt")):
        if f.exists():
            shutil.move(f, dest / f.name)
    if error:
        (dest / f"{img.name}.error.txt").write_text(error + "\n", encoding="utf8")


def cmd_add(args) -> int:
    res = process(args.images, args)
    return 1 if any(res.values()) else 0


def cmd_watch(args) -> int:
    args.inbox.mkdir(exist_ok=True)
    print(f"監視中: {args.inbox}/ に領収書画像を置くと自動で転記します(Ctrl+C で終了)")
    seen: dict[Path, int] = {}
    while True:
        # 書き込み途中のファイルを避けるため、サイズが前回から変わっていないものだけ処理
        ready = []
        for p in _scan(args.inbox):
            size = p.stat().st_size
            if seen.get(p) == size:
                ready.append(p)
            seen[p] = size
        if ready:
            try:
                for img, err in process(ready, args).items():
                    if not args.dry_run:
                        _archive(img, args.inbox, err)
                    seen.pop(img, None)
            except WorkbookError as e:
                print(f"エラー: {e}(ファイルは受信箱に残します)", file=sys.stderr)
        time.sleep(args.interval)


def main(argv=None) -> int:
    common = argparse.ArgumentParser(add_help=False)
    common.add_argument("--workbook", type=Path, default=DEFAULT_WORKBOOK, help=f"経費精算書(既定 {DEFAULT_WORKBOOK})")
    common.add_argument("--purpose", default="", help="目的/現場・行程/内容(全画像に共通で入れる)")
    common.add_argument("--category", choices=list(CATEGORIES), help="種別を固定(通常はAIが判定)")
    common.add_argument("--applicant", help="申請者名を書き換える")
    common.add_argument("--new-sheet", action="store_true", help="新しい原簿シートに書き始める")
    common.add_argument("--dry-run", action="store_true", help="読み取り結果を表示するだけで書き込まない")
    ap = argparse.ArgumentParser(prog="expense_bot", description="領収書画像を経費精算書へ自動転記")
    sub = ap.add_subparsers(dest="cmd", required=True)
    a = sub.add_parser("add", parents=[common], help="指定した画像を転記")
    a.add_argument("images", nargs="+", type=Path)
    a.set_defaults(fn=cmd_add)
    w = sub.add_parser("watch", parents=[common], help="受信箱フォルダを監視して自動転記")
    w.add_argument("--inbox", type=Path, default=DEFAULT_INBOX)
    w.add_argument("--interval", type=float, default=5.0)
    w.set_defaults(fn=cmd_watch)
    args = ap.parse_args(argv)
    try:
        return args.fn(args)
    except (WorkbookError, KeyboardInterrupt) as e:
        print(f"エラー: {e}" if isinstance(e, WorkbookError) else "終了します", file=sys.stderr)
        return 1
