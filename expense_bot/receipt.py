from __future__ import annotations

from dataclasses import dataclass
from datetime import date

# 経費精算書の「種別」欄(交 / 経 / 際)に対応
CATEGORIES = {
    "transport": "交",  # 交通費: 電車・バス・タクシー・駐車場・有料道路・ガソリン
    "site": "経",  # 現場経費: 消耗品・材料・備品など
    "entertainment": "際",  # 交際費: 飲食接待・会費・慶弔・贈答
}


@dataclass
class Receipt:
    date: date
    vendor: str  # 購入仕入先・出張先
    item: str  # 物品名・用件
    amount: int  # 税込合計(円)
    category: str  # CATEGORIES のキー
    purpose: str = ""  # 目的/現場・行程/内容(領収書からは分からないので通常は手入力)
    unit_price: int | None = None  # 単価(税込)。数量が複数のときだけ
    quantity: int | None = None
    source: str = ""  # 元の画像ファイル名(ログ用)

    @property
    def mark(self) -> str:
        return CATEGORIES[self.category]
