"""
台帳の各列がどう読まれているかを、列ごとに全行スキャンして報告する。

    python tools/check_ledger_dates.py "C:\\path\\to\\台帳.xlsx"
    python tools/check_ledger_dates.py data\\references          # フォルダごと

前の版は各シート 1 行しか見ておらず、行によって型が違う列
（複数ファイルを結合した台帳では普通に起きる）を取り逃がしていた。
この版は全行を読み、列ごとに型の内訳を出す。

読み方：
  「型の内訳」に str と int/float が混在している列は要注意。
  同じ列なのに行によって保存され方が違うということで、
  日付が数値（シリアル値）のまま残っている可能性が高い。
"""
from __future__ import annotations

import datetime as dt
import re
import sys
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

try:
    import openpyxl
except ImportError:
    print("openpyxl が入っていません。pip install openpyxl")
    raise SystemExit(1)

try:
    from app.cell_format import cell_to_text, is_date_format, serial_to_datetime
except Exception as e:  # noqa: BLE001
    print(f"app/cell_format.py を読み込めません: {e}")
    print("プロジェクト直下から実行してください。")
    raise SystemExit(1)

#: 日付シリアル値としてありえる範囲（1995-05-28 〜 2035-01-01 あたり）
SERIAL_MIN, SERIAL_MAX = 34_800, 49_500

#: 文字列として入っている日付の代表的な形
_STR_DATE_PATTERNS = [
    re.compile(r"^\s*(\d{4})[/\-年](\d{1,2})[/\-月](\d{1,2})日?"),        # 2020/05/28
    re.compile(r"^\s*(\d{1,2})[/\-](\d{1,2})[/\-](\d{4})"),               # 05/28/2020
]


def looks_like_date_string(text: str) -> bool:
    return any(p.match(str(text)) for p in _STR_DATE_PATTERNS)


def classify(value, number_format) -> str:
    """このセルが何に見えるかを一言で。"""
    if value is None or str(value).strip() == "":
        return "empty"
    if isinstance(value, (dt.datetime, dt.date, dt.time)):
        return "日付型"
    if isinstance(value, bool):
        return "真偽値"
    if isinstance(value, (int, float)):
        if is_date_format(number_format):
            return "数値+日付書式"
        if SERIAL_MIN <= float(value) <= SERIAL_MAX:
            return "★シリアル値疑い"
        return "数値"
    if looks_like_date_string(value):
        return "日付文字列"
    return "文字列"


def check_workbook(path: Path, sample_rows: int = 3) -> None:
    print("=" * 92)
    print(f"■ {path.name}")
    print("=" * 92)

    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    except Exception as e:  # noqa: BLE001
        print(f"  開けません: {type(e).__name__}: {e}")
        return

    for ws in wb.worksheets:
        if getattr(ws, "sheet_state", "visible") != "visible":
            continue

        rows = list(ws.iter_rows())
        if len(rows) < 2:
            continue

        headers = ["" if c.value is None else str(c.value).strip() for c in rows[0]]
        n_cols = len(headers)
        kinds: list[Counter] = [Counter() for _ in range(n_cols)]
        fmts: list[Counter] = [Counter() for _ in range(n_cols)]
        samples: list[list[tuple[int, object, str]]] = [[] for _ in range(n_cols)]
        suspect_rows: list[list[int]] = [[] for _ in range(n_cols)]

        for r_idx, row in enumerate(rows[1:], start=2):
            for c_idx in range(min(n_cols, len(row))):
                cell = row[c_idx]
                v = cell.value
                fmt = getattr(cell, "number_format", None)
                k = classify(v, fmt)
                kinds[c_idx][k] += 1
                if k == "empty":
                    continue
                fmts[c_idx][str(fmt)] += 1
                if len(samples[c_idx]) < sample_rows:
                    samples[c_idx].append((r_idx, v, k))
                if k == "★シリアル値疑い":
                    if len(suspect_rows[c_idx]) < 5:
                        suspect_rows[c_idx].append(r_idx)

        total = len(rows) - 1
        print(f"\n  --- シート「{ws.title}」  データ {total} 行 × {n_cols} 列 ---\n")

        problem_cols = []
        for c_idx in range(n_cols):
            counts = kinds[c_idx]
            filled = sum(n for k, n in counts.items() if k != "empty")
            if filled == 0:
                continue
            head = headers[c_idx] or f"({c_idx + 1}列目)"
            breakdown = " / ".join(f"{k}×{n}" for k, n in counts.most_common() if k != "empty")
            fmt_top = ", ".join(f for f, _ in fmts[c_idx].most_common(2))

            date_like = counts["日付型"] + counts["数値+日付書式"] + counts["日付文字列"]
            serial_like = counts["★シリアル値疑い"]
            flag = ""
            if serial_like and date_like:
                flag = "  ← ★★ 日付列に数値が混在（最優先）"
                problem_cols.append((head, "混在", serial_like, date_like))
            elif serial_like:
                flag = "  ← ★ 全部数値。日付列かどうか判断できない"
                problem_cols.append((head, "全数値", serial_like, 0))
            elif counts["日付文字列"]:
                flag = "  ← 日付が文字列で入っている"
                problem_cols.append((head, "文字列日付", 0, counts["日付文字列"]))

            print(f"  [{c_idx + 1:>2}] {head[:22]:<22} {filled:>4}件  {breakdown}{flag}")
            print(f"       書式: {fmt_top[:60]}")
            for r_idx, v, k in samples[c_idx]:
                out = cell_to_text(v, fmts[c_idx].most_common(1)[0][0]
                                   if fmts[c_idx] else None)
                extra = ""
                if k == "★シリアル値疑い":
                    d = serial_to_datetime(v)
                    extra = f"   → 日付なら {d.date()}" if d else ""
                print(f"       {r_idx}行目: {str(v)[:34]:<34} → 出力 {out!r}{extra}")
            if suspect_rows[c_idx]:
                print(f"       数値のままの行: {suspect_rows[c_idx]} …")
            print()

        if problem_cols:
            print("  " + "!" * 86)
            print("  対応が必要な列")
            print("  " + "!" * 86)
            for head, kind, serial_n, date_n in problem_cols:
                if kind == "混在":
                    print(f"    「{head}」: 日付 {date_n} 件 / 数値のまま {serial_n} 件")
                    print(f"       同じ列で行ごとに保存され方が違います。"
                          f"列内の日付を根拠に数値側も日付として復元できます。")
                elif kind == "全数値":
                    print(f"    「{head}」: 全 {serial_n} 件が数値。")
                    print(f"       日付列なのか金額・件数なのか、ファイルからは判断できません。"
                          f"Excel で表示を確認してください。")
                else:
                    print(f"    「{head}」: 日付が文字列 {date_n} 件。"
                          f"形式を統一すれば揃えられます。")
            print()

    try:
        wb.close()
    except Exception:  # noqa: BLE001
        pass


def main() -> None:
    if len(sys.argv) < 2:
        print(__doc__)
        raise SystemExit(1)

    target = Path(sys.argv[1])
    if target.is_dir():
        files = sorted(p for p in target.rglob("*.xls*") if not p.name.startswith("~$"))
    elif target.is_file():
        files = [target]
    else:
        print(f"見つかりません: {target}")
        raise SystemExit(1)

    if not files:
        print(f"{target} に Excel ファイルがありません")
        raise SystemExit(1)

    for f in files:
        check_workbook(f)


if __name__ == "__main__":
    main()
