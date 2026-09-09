#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
完了日の年度誤り検出・修正スクリプト

前提:
  - 依頼日は正しい
  - 完了日は「月・日」は正しいが「年」だけ誤っている場合がある

方針:
  月日を固定したまま年を振り直し、「完了日 >= 依頼日」を満たす中で
  間隔が最小になる年を採用する。採用後に再検証を行い、怪しいものは
  自動修正せず「要確認」として残す。

使い方:
  # まず確認（ファイルは変更されない）
  python fix_kanryobi.py data.xlsx

  # 実際に上書き（.bak を自動作成）
  python fix_kanryobi.py data.xlsx --apply

  # 列名やシートを明示、ログをファイルへ
  python fix_kanryobi.py data.xlsx --apply \
      --sheet "案件一覧" --request-col 依頼日 --done-col 完了日 \
      --log fix_log.txt
"""

import argparse
import csv
import datetime as dt
import os
import re
import shutil
import sys

# ---------------------------------------------------------------- 設定既定値

# 列名の自動判定に使う候補（部分一致・全角半角ゆれを吸収）
REQUEST_KEYS = ["依頼日", "依赖日", "依頼", "受付日", "開始日", "request", "start"]
DONE_KEYS = ["完了日", "完了", "終了日", "納品日", "done", "end", "complete"]

# 間隔しきい値の自動決定に使うパーセンタイルと安全係数
PCTL = 0.95
SAFETY = 2.0
MIN_THRESHOLD_DAYS = 180   # これより短いしきい値は使わない（誤検出防止）
MAX_THRESHOLD_DAYS = 1000  # 上限

# 年の探索範囲（依頼日の年に対する相対）
YEAR_OFFSETS = [-1, 0, 1, 2, 3]

EXCEL_EPOCH = dt.date(1899, 12, 30)   # Excel 1900 日付システム
SERIAL_MIN, SERIAL_MAX = 20000, 60000  # 1954年〜2064年あたり


# ---------------------------------------------------------------- 日付パース

def serial_to_date(n):
    """Excel シリアル値 -> date"""
    try:
        return EXCEL_EPOCH + dt.timedelta(days=int(n))
    except (ValueError, OverflowError):
        return None


def strip_time(s):
    """末尾の時刻部分を取り除く（'06/23/2020 0:00:00' -> '06/23/2020'）"""
    return re.sub(r"[\sT]+\d{1,2}:\d{2}(:\d{2})?(\.\d+)?\s*([APap][Mm])?$", "", s).strip()


SLASH_PAT = re.compile(r"^(\d{1,4})[/\-.](\d{1,2})[/\-.](\d{1,4})$")


def detect_date_order(raw_values):
    """
    列の値全体を見て、区切り記号つき日付の並び順を判定する。

    05/06/2020 のように両方 12 以下だと月日の区別がつかないため、
    列内に 13 以上の値が現れる行を手がかりにする。
    戻り値: ("mdy" | "dmy" | "ymd", 判定根拠の文字列)
    """
    first_gt12 = second_gt12 = ymd = total = 0
    for v in raw_values:
        if not isinstance(v, str):
            continue
        m = SLASH_PAT.match(strip_time(v.strip()))
        if not m:
            continue
        a, b, _c = int(m.group(1)), int(m.group(2)), int(m.group(3))
        total += 1
        if a > 31:
            ymd += 1
        elif a > 12:
            first_gt12 += 1
        elif b > 12:
            second_gt12 += 1

    if total == 0:
        return "mdy", "対象なし"
    if ymd > 0 and ymd >= total * 0.5:
        return "ymd", f"年が先頭の形式 {ymd}/{total} 件"
    if first_gt12 and second_gt12:
        return "mdy", (f"⚠ 月日の並びが混在（1番目>12 が {first_gt12}件、"
                       f"2番目>12 が {second_gt12}件）。mm/dd として処理します")
    if second_gt12:
        return "mdy", f"2番目に13以上が {second_gt12}/{total} 件 → mm/dd/yyyy と判定"
    if first_gt12:
        return "dmy", f"1番目に13以上が {first_gt12}/{total} 件 → dd/mm/yyyy と判定"
    return "mdy", ("⚠ すべて12以下で月日の判別不能。mm/dd と仮定します"
                   "（違う場合は --date-order dmy を指定）")


def parse_date(value, order="mdy"):
    """
    セルの値を date に変換する。
    datetime / date / Excel シリアル値 / 文字列 いずれにも対応。
    order は区切り記号つき日付の並び順（"mdy" / "dmy" / "ymd"）。
    変換できない場合は None。
    """
    if value is None:
        return None

    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value

    # 数値 = シリアル値の可能性
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if SERIAL_MIN <= value <= SERIAL_MAX:
            return serial_to_date(value)
        # 20230501 のように yyyymmdd が数値で入っている場合
        if float(value).is_integer() and 19000101 <= value <= 21001231:
            try:
                return dt.datetime.strptime(str(int(value)), "%Y%m%d").date()
            except ValueError:
                return None
        return None

    s = str(value).strip()
    if not s:
        return None
    # 未完了などの文字列を除外
    if re.search(r"(未完了|未定|進行中|作業中|保留|TBD|N/?A|[-—ー―])$", s, re.I):
        return None

    # 数字だけの場合: 8桁は yyyymmdd、それ以外はシリアル値として扱う
    if re.fullmatch(r"\d+(\.\d+)?", s):
        if re.fullmatch(r"\d{8}", s):
            try:
                return dt.datetime.strptime(s, "%Y%m%d").date()
            except ValueError:
                return None
        v = float(s)
        if SERIAL_MIN <= v <= SERIAL_MAX:
            return serial_to_date(v)
        return None

    # 和暦・漢字表記を正規化し、時刻部分を落とす
    s = s.replace("年", "/").replace("月", "/").replace("日", "")
    s = s.replace("．", ".").replace("－", "-").replace("／", "/").replace("．", ".")
    s = re.sub(r"\s+", " ", s).strip()
    s = re.sub(r"\([月火水木金土日]\)", "", s).strip()   # (水) などの曜日表記
    s = strip_time(s).rstrip("/").strip()

    # 年が先頭の形式は並び順に関わらず一意に決まるので先に試す
    fmts = ["%Y/%m/%d", "%Y-%m-%d", "%Y.%m.%d", "%Y%m%d"]
    if order == "dmy":
        fmts += ["%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%d/%m/%y"]
    elif order == "ymd":
        fmts += ["%y/%m/%d", "%y-%m-%d"]
    else:
        fmts += ["%m/%d/%Y", "%m-%d-%Y", "%m.%d.%Y", "%m/%d/%y"]
    fmts += ["%y/%m/%d", "%y-%m-%d"]

    for f in fmts:
        try:
            return dt.datetime.strptime(s, f).date()
        except ValueError:
            continue
    return None


def fiscal_year(d):
    """日本の年度（4月始まり）"""
    return d.year if d.month >= 4 else d.year - 1


def safe_replace_year(d, year):
    """月日を保ったまま年を差し替える。2/29 が存在しない年なら None。"""
    try:
        return d.replace(year=year)
    except ValueError:
        return None


# ---------------------------------------------------------------- 修正ロジック

def percentile(sorted_vals, p):
    if not sorted_vals:
        return None
    k = (len(sorted_vals) - 1) * p
    lo, hi = int(k), min(int(k) + 1, len(sorted_vals) - 1)
    return sorted_vals[lo] + (sorted_vals[hi] - sorted_vals[lo]) * (k - lo)


def decide_threshold(pairs, override=None):
    """
    正常に見える行の工期分布からしきい値を決める。

    誤りのある行自身が分布を押し上げてしまうため、
    「しきい値を出す -> 超える行を除外 -> 再計算」を収束するまで繰り返す
    （トリム反復）。これで外れ値の影響を受けにくくする。
    """
    if override:
        return override, "手動指定"

    gaps = sorted((done - req).days for req, done in pairs if done >= req)
    if len(gaps) < 5:
        return MIN_THRESHOLD_DAYS, "サンプル不足のため既定値"

    work = gaps
    th = None
    for _ in range(8):
        p = percentile(work, PCTL)
        new_th = int(max(MIN_THRESHOLD_DAYS, min(MAX_THRESHOLD_DAYS, p * SAFETY)))
        kept = [g for g in work if g <= new_th]
        if new_th == th or len(kept) < 5 or len(kept) == len(work):
            th = new_th
            break
        th, work = new_th, kept

    return th, (f"P{int(PCTL*100)}={int(percentile(work, PCTL))}日 x{SAFETY}"
                f" / 母数 {len(work)}件（トリム後）")


def build_candidates(req, done):
    """年を振り直した候補リスト [(date, gap), ...] を gap 昇順で返す。"""
    out = []
    for off in YEAR_OFFSETS:
        c = safe_replace_year(done, req.year + off)
        if c is None:
            continue          # 2/29 が存在しない年
        gap = (c - req).days
        if gap < 0:
            continue          # 完了日が依頼日より前になる候補は除外
        out.append((c, gap))
    out.sort(key=lambda x: x[1])
    return out


def evaluate(req, done, threshold):
    """
    1 行を評価する。
    戻り値: (判定, 修正後の日付 or None, 理由)
      判定 ∈ {"OK", "FIX", "REVIEW"}
    """
    gap = (done - req).days

    reversed_ = gap < 0
    too_long = gap > threshold
    # 年度差による補助検出。しきい値に依存しないので、
    # データ件数が少なくて自動しきい値が甘くなった場合の保険になる。
    fy_gap = fiscal_year(done) - fiscal_year(req)
    odd_fy = fy_gap >= 2

    if not reversed_ and not too_long and not odd_fy:
        return "OK", None, ""

    if reversed_:
        reason = "完了日が依頼日より前"
    elif too_long:
        reason = f"間隔が異常に長い({gap}日)"
    else:
        reason = f"年度が {fy_gap} 年度先"

    cands = build_candidates(req, done)
    if not cands:
        return "REVIEW", None, reason + " / 妥当な候補年なし"

    best, best_gap = cands[0]

    # ---- 再検証 -------------------------------------------------
    problems = []

    if best == done:
        # 年を変えても最良が元のままなら、年の誤りではない可能性が高い
        return "REVIEW", None, reason + " / 年の振替では解消せず（長期案件の可能性）"

    if (best.month, best.day) != (done.month, done.day):
        problems.append("月日が保持されていない")

    if best < req:
        problems.append("依頼日より前")

    if best_gap > threshold:
        problems.append(f"修正後も間隔が長い({best_gap}日)")

    # 年度チェック: 修正後は依頼日と同年度 or 翌年度に収まるのが自然
    fy_diff = fiscal_year(best) - fiscal_year(req)
    if fy_diff not in (0, 1):
        problems.append(f"年度差 {fy_diff}")

    # 曖昧さチェック:
    # 候補年は 1 年（約365日）刻みなので通常は最良候補が明確に決まる。
    # 2位との差が小さい（=判断が割れる）ケースだけを警告する。
    if len(cands) > 1 and (cands[1][1] - best_gap) < 200 and cands[1][1] <= threshold:
        problems.append(f"2位候補 {cands[1][0]} と接近（要判断）")

    if problems:
        return "REVIEW", best, reason + " / 候補=" + best.isoformat() + " ただし " + "、".join(problems)

    # 年度差だけが根拠の場合は、長期案件との区別がつかないため自動修正しない
    if odd_fy and not reversed_ and not too_long:
        return "REVIEW", best, (f"{reason} / 候補={best.isoformat()}"
                                f"（間隔 {gap}日 -> {best_gap}日）長期案件の可能性あり")

    # 確信度:
    #   高 = 完了日が依頼日より前（正当な解釈が存在しない明白な誤り）
    #   中 = 間隔が長いことによる判定（長期案件と原理的に区別できない）
    conf = "高" if reversed_ else "中"
    return "FIX", best, (f"[確信度{conf}] {reason} -> 年を {done.year} から "
                         f"{best.year} へ（間隔 {best_gap}日）")


# ---------------------------------------------------------------- 列の特定

def find_col(headers, keys, explicit=None):
    """ヘッダー行から対象列のインデックスを返す。見つからなければ None。"""
    norm = [(str(h).strip() if h is not None else "") for h in headers]

    if explicit:
        for i, h in enumerate(norm):
            if h == explicit:
                return i
        for i, h in enumerate(norm):
            if explicit in h:
                return i
        return None

    for k in keys:                       # 完全一致優先
        for i, h in enumerate(norm):
            if h == k:
                return i
    for k in keys:                       # 部分一致
        for i, h in enumerate(norm):
            if h and k.lower() in h.lower():
                return i
    return None


# ---------------------------------------------------------------- 実行本体

def backup_path(path):
    """元の拡張子を保ったバックアップ名を作る（Excel でそのまま開ける）"""
    root, ext = os.path.splitext(path)
    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"{root}_backup_{stamp}{ext}"


class Logger:
    def __init__(self, path=None):
        self.lines = []
        self.path = path

    def __call__(self, msg=""):
        print(msg)
        self.lines.append(msg)

    def file_only(self, msg=""):
        """画面には出さず、ログファイルにだけ残す（大量の明細用）"""
        self.lines.append(msg)

    def save(self):
        if self.path:
            with open(self.path, "w", encoding="utf-8-sig") as f:
                f.write("\n".join(self.lines) + "\n")


def process_rows(rows, ci_req, ci_done, threshold_override, log, header_row_no,
                 date_order="auto"):
    """
    rows: [(行番号, [セル値...]), ...]
    戻り値: [(行番号, 修正後date, 元date), ...]  実際に直す対象
    """
    # --- 日付の並び順を列全体から判定 -------------------------------
    raw = []
    for _, cells in rows:
        if ci_req < len(cells):
            raw.append(cells[ci_req])
        if ci_done < len(cells):
            raw.append(cells[ci_done])

    if date_order == "auto":
        order, why = detect_date_order(raw)
        log(f"日付の並び順: {order} ({why})")
    else:
        order, why = date_order, "手動指定"
        log(f"日付の並び順: {order} (手動指定)")
    log("")

    parsed = []
    for rowno, cells in rows:
        req = parse_date(cells[ci_req], order) if ci_req < len(cells) else None
        done = parse_date(cells[ci_done], order) if ci_done < len(cells) else None
        parsed.append((rowno, cells, req, done))

    pairs = [(r, d) for _, _, r, d in parsed if r and d]
    threshold, how = decide_threshold(pairs, threshold_override)

    log(f"有効な日付ペア: {len(pairs)} 件")
    log(f"間隔しきい値  : {threshold} 日 ({how})")
    log(f"  ※ 実在する工期が {threshold} 日を超える案件は誤って修正対象になります。")
    log(f"     心当たりがあれば --threshold で業務実態に合わせてください。")
    log("")

    fixes, reviews, skipped = [], [], []

    for rowno, cells, req, done in parsed:
        if req is None or done is None:
            raw_req = cells[ci_req] if ci_req < len(cells) else None
            raw_done = cells[ci_done] if ci_done < len(cells) else None
            if raw_req is None and raw_done is None:
                continue          # 完全な空行なので無視
            if (str(raw_req or "").strip() == "" and
                    str(raw_done or "").strip() == ""):
                continue
            detail = []
            if req is None and str(raw_req or "").strip():
                detail.append(f"依頼日を解釈できない: {raw_req!r}")
            elif req is None:
                detail.append("依頼日が空欄")
            if done is None and str(raw_done or "").strip():
                detail.append(f"完了日を解釈できない: {raw_done!r}")
            elif done is None:
                detail.append("完了日が空欄")
            skipped.append((rowno, " / ".join(detail)))
            continue

        verdict, newdate, reason = evaluate(req, done, threshold)
        if verdict == "FIX":
            fixes.append((rowno, newdate, done, req, reason))
        elif verdict == "REVIEW":
            reviews.append((rowno, newdate, done, req, reason))

    log("=" * 72)
    log(f"自動修正の対象: {len(fixes)} 件")
    log("=" * 72)
    for rowno, newdate, done, req, reason in fixes:
        log(f"  行{rowno:>5}  依頼 {req}  完了 {done} -> {newdate}   {reason}")
    if not fixes:
        log("  （なし）")

    log("")
    log("=" * 72)
    log(f"要確認（自動修正しない）: {len(reviews)} 件")
    log("=" * 72)
    for rowno, newdate, done, req, reason in reviews:
        log(f"  行{rowno:>5}  依頼 {req}  完了 {done}   {reason}")
    if not reviews:
        log("  （なし）")

    if skipped:
        log("")
        log(f"スキップ: {len(skipped)} 件")
        for rowno, why in skipped[:30]:
            log(f"  行{rowno:>5}  {why}")
        if len(skipped) > 30:
            log(f"  ... 他 {len(skipped)-30} 件（全件は --log のファイルに出力）")
            log.file_only("")
            log.file_only("--- スキップ全件 ---")
            for rowno, why in skipped:
                log.file_only(f"  行{rowno:>5}  {why}")

    return [(r, nd, od) for r, nd, od, _, _ in fixes]


def run_excel(path, args, log):
    try:
        import openpyxl
    except ImportError:
        sys.exit("openpyxl が必要です:  pip install openpyxl")

    wb = openpyxl.load_workbook(path)          # 数式は文字列のまま保持
    ws = wb[args.sheet] if args.sheet else wb.active
    log(f"シート: {ws.title}")

    hr = args.header_row
    headers = [c.value for c in ws[hr]]
    ci_req = find_col(headers, REQUEST_KEYS, args.request_col)
    ci_done = find_col(headers, DONE_KEYS, args.done_col)

    if ci_req is None or ci_done is None:
        log(f"ヘッダー行({hr}): {headers}")
        sys.exit("依頼日 / 完了日 の列を特定できません。--request-col / --done-col で指定してください。")

    log(f"依頼日列: {headers[ci_req]} (列{ci_req+1})")
    log(f"完了日列: {headers[ci_done]} (列{ci_done+1})")
    log("")

    rows = []
    for r in range(hr + 1, ws.max_row + 1):
        rows.append((r, [ws.cell(row=r, column=c).value
                         for c in range(1, ws.max_column + 1)]))

    fixes = process_rows(rows, ci_req, ci_done, args.threshold, log, hr,
                         args.date_order)

    if not fixes:
        log("\n修正対象がないため、ファイルは変更しません。")
        return

    if not args.apply:
        log(f"\n[確認モード] {len(fixes)} 件の修正候補。実際に反映するには --apply を付けてください。")
        return

    if not args.no_backup:
        bak = backup_path(path)
        shutil.copy2(path, bak)
        log(f"\nバックアップ作成: {bak}")

    for rowno, newdate, _ in fixes:
        cell = ws.cell(row=rowno, column=ci_done + 1)
        fmt = cell.number_format          # 元の表示形式を保持
        cell.value = dt.datetime(newdate.year, newdate.month, newdate.day)
        if fmt and fmt != "General":
            cell.number_format = fmt
        else:
            cell.number_format = "yyyy/mm/dd"

    wb.save(path)
    log(f"上書き完了: {path}  ({len(fixes)} 件修正)")


def run_csv(path, args, log):
    enc = args.encoding
    with open(path, newline="", encoding=enc) as f:
        table = list(csv.reader(f))

    hr = args.header_row
    if len(table) < hr:
        sys.exit("行数が足りません。")
    headers = table[hr - 1]

    ci_req = find_col(headers, REQUEST_KEYS, args.request_col)
    ci_done = find_col(headers, DONE_KEYS, args.done_col)
    if ci_req is None or ci_done is None:
        log(f"ヘッダー行: {headers}")
        sys.exit("依頼日 / 完了日 の列を特定できません。--request-col / --done-col で指定してください。")

    log(f"依頼日列: {headers[ci_req]}")
    log(f"完了日列: {headers[ci_done]}")
    log("")

    rows = [(i + 1, table[i]) for i in range(hr, len(table))]
    fixes = process_rows(rows, ci_req, ci_done, args.threshold, log, hr,
                         args.date_order)

    if not fixes:
        log("\n修正対象がないため、ファイルは変更しません。")
        return
    if not args.apply:
        log(f"\n[確認モード] {len(fixes)} 件の修正候補。実際に反映するには --apply を付けてください。")
        return

    if not args.no_backup:
        bak = backup_path(path)
        shutil.copy2(path, bak)
        log(f"\nバックアップ作成: {bak}")

    for rowno, newdate, _ in fixes:
        table[rowno - 1][ci_done] = newdate.strftime(args.csv_date_format)

    with open(path, "w", newline="", encoding=enc) as f:
        csv.writer(f).writerows(table)
    log(f"上書き完了: {path}  ({len(fixes)} 件修正)")


def main():
    ap = argparse.ArgumentParser(description="完了日の年度誤りを検出・修正する")
    ap.add_argument("file", help="xlsx / xlsm / csv ファイル")
    ap.add_argument("--sheet", help="シート名（既定: 先頭シート）")
    ap.add_argument("--header-row", type=int, default=1, help="ヘッダー行番号（既定: 1）")
    ap.add_argument("--request-col", help="依頼日の列名を明示")
    ap.add_argument("--done-col", help="完了日の列名を明示")
    ap.add_argument("--threshold", type=int, help="工期しきい値（日）。未指定なら自動算出")
    ap.add_argument("--date-order", default="auto", choices=["auto", "mdy", "dmy", "ymd"],
                    help="文字列日付の並び順。既定は auto（列全体から判定）")
    ap.add_argument("--apply", action="store_true", help="実際にファイルを上書きする")
    ap.add_argument("--no-backup", action="store_true", help=".bak を作らない")
    ap.add_argument("--log", help="ログの出力先ファイル")
    ap.add_argument("--encoding", default="utf-8-sig", help="CSV の文字コード（既定: utf-8-sig）")
    ap.add_argument("--csv-date-format", default="%Y/%m/%d", help="CSV 書き戻し時の日付書式")
    args = ap.parse_args()

    if not os.path.exists(args.file):
        sys.exit(f"ファイルが見つかりません: {args.file}")

    log = Logger(args.log)
    log(f"対象ファイル: {args.file}")
    log(f"実行日時    : {dt.datetime.now():%Y-%m-%d %H:%M:%S}")
    log(f"モード      : {'上書き' if args.apply else '確認のみ'}")
    log("")

    ext = os.path.splitext(args.file)[1].lower()
    try:
        if ext in (".xlsx", ".xlsm", ".xltx"):
            run_excel(args.file, args, log)
        elif ext in (".csv", ".tsv", ".txt"):
            run_csv(args.file, args, log)
        else:
            sys.exit(f"未対応の拡張子: {ext}")
    finally:
        log.save()
        if args.log:
            print(f"\nログ保存: {args.log}")


if __name__ == "__main__":
    main()
