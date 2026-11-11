"""Claude가 영수증을 읽어 쓴 receipts.json 을 검산하고 경비 엑셀을 만든다.

    python check.py out/10월경비 samples

검사기는 원본 이미지를 보지 못한다. 그래서 두 가지로 나눈다.
  형식 오류  — 필드 빠짐, 날짜 형식, 금액이 정수가 아님, 입력 파일 누락·중복 기재
              → 종료 코드 1. receipts.json 을 고쳐서 다시 돌린다.
  검산 걸림  — 품목 금액 != 수량×단가, 품목 합 != 합계(또는 공급가액), 공급가액+부가세 != 합계,
              같은 날짜·가맹점·합계 중복, 읽은 쪽이 "확인 필요"로 표시한 것
              → 종료 코드 0. 엑셀 "확인 필요" 시트로 간다. 숫자를 맞추려고 고치지 않는다.
"""
import json
import re
import sys
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

PAYMENTS = {"카드", "현금", "계좌이체", "기타"}
INPUT_EXT = {".png", ".jpg", ".jpeg", ".pdf"}


def is_int(v):
    return isinstance(v, int) and not isinstance(v, bool)


def format_errors(r):
    """한 건의 형식 오류. 사람이 확인할 건이라도 file·status·note 는 있어야 한다."""
    e = []
    if r.get("status") not in ("ok", "확인 필요"):
        return [f"status 는 ok 또는 확인 필요: {r.get('status')!r}"]
    if r["status"] == "확인 필요":
        if not str(r.get("note") or "").strip():
            e.append("확인 필요인데 note(이유)가 없다")
        # 못 읽은 칸은 null 이어도 되지만, 적었다면 형식은 맞아야 한다(엑셀 확인 필요 시트에 그대로 들어간다)
        d = r.get("date")
        if d is not None and not (isinstance(d, str) and re.fullmatch(r"\d{4}-\d{2}-\d{2}", d)):
            e.append(f"date 는 YYYY-MM-DD 또는 null: {d!r}")
        if r.get("total") is not None and not is_int(r["total"]):
            e.append(f"total 은 정수 또는 null: {r['total']!r}")
        if r.get("merchant") is not None and not isinstance(r["merchant"], str):
            e.append(f"merchant 는 문자열 또는 null: {r['merchant']!r}")
        return e
    d = r.get("date")
    if not (isinstance(d, str) and re.fullmatch(r"\d{4}-\d{2}-\d{2}", d)):
        e.append(f"date 는 YYYY-MM-DD: {d!r}")
    else:
        try:
            date.fromisoformat(d)
        except ValueError:
            e.append(f"없는 날짜: {d}")
    if not str(r.get("merchant") or "").strip():
        e.append("merchant 가 비었다")
    if r.get("payment") not in PAYMENTS:
        e.append(f"payment 는 {sorted(PAYMENTS)} 중 하나: {r.get('payment')!r}")
    if not (is_int(r.get("total")) and r["total"] > 0):
        e.append(f"total 은 0보다 큰 정수(원): {r.get('total')!r}")
    for k in ("supply", "vat"):
        if r.get(k) is not None and not is_int(r[k]):
            e.append(f"{k} 는 정수 또는 null: {r[k]!r}")
    if (r.get("supply") is None) != (r.get("vat") is None):
        e.append("supply·vat 는 둘 다 있거나 둘 다 null")
    items = r.get("items")
    if not (isinstance(items, list) and items):
        e.append("items 가 비었다")
    else:
        for i, it in enumerate(items, 1):
            if not str(it.get("name") or "").strip():
                e.append(f"품목 {i}: name 이 비었다")
            for k in ("qty", "unit_price", "amount"):
                if not is_int(it.get(k)):
                    e.append(f"품목 {i}: {k} 는 정수: {it.get(k)!r}")
    return e


def arithmetic(r):
    """검산. 원본이 원래 안 맞을 수도 있으니 오류가 아니라 '확인 필요' 사유로 돌려준다."""
    why = []
    for i, it in enumerate(r["items"], 1):
        if it["qty"] * it["unit_price"] != it["amount"]:
            why.append(f"품목 {i} 금액 {it['amount']:,} != {it['qty']:,}×{it['unit_price']:,}")
    s = sum(it["amount"] for it in r["items"])
    # 카드 영수증은 부가세 포함 금액(품목 합 = 합계), 청구서는 공급가 기준(품목 합 = 공급가액)
    if s not in (r["total"], r.get("supply")):
        target = f"공급가액 {r['supply']:,}" if r.get("supply") is not None else f"합계 {r['total']:,}"
        why.append(f"품목 합 {s:,} != {target}")
    if r.get("supply") is not None and r["supply"] + r["vat"] != r["total"]:
        why.append(f"공급가액 {r['supply']:,} + 부가세 {r['vat']:,} != 합계 {r['total']:,}")
    return why


def review(rows, files):
    """rows: receipts.json, files: 입력 폴더의 파일 이름들 → (형식 오류, 통과 행, [(행, 사유)])"""
    if not isinstance(rows, list):
        return ["receipts.json 은 리스트여야 한다"], [], []
    errors, ok, flagged, seen = [], [], [], {}
    named = [Path(str(r.get("file", ""))).name for r in rows]
    for f in sorted(set(files) - set(named)):
        errors.append(f"{f}: receipts.json 에 없다 — 모든 파일을 한 줄씩 기록한다(못 읽으면 확인 필요)")
    for f in sorted({n for n in named if named.count(n) > 1}):
        errors.append(f"{f}: 두 번 기록했다 — 파일 하나에 한 줄")
    for name, r in zip(named, rows):
        if name not in files:
            errors.append(f"{name or '(file 없음)'}: 입력 폴더에 없는 파일")
            continue
        e = format_errors(r)
        errors += [f"{name}: {x}" for x in e]
        if e:
            continue
        if r["status"] == "확인 필요":
            flagged.append((r, [f"읽은 쪽 표시: {r['note']}"]))
            continue
        why = arithmetic(r)
        key = (r["date"], r["merchant"].replace(" ", ""), r["total"])
        if key in seen:
            why.append(f"중복 의심: {seen[key]} 와 날짜·가맹점·합계가 같다")
        else:
            seen[key] = name
        (flagged.append((r, why)) if why else ok.append(r))
    return errors, ok, flagged


def safe(row):
    """영수증에서 읽은 글자가 = + - @ 로 시작하면 엑셀이 수식으로 실행한다 — 앞에 ' 를 붙여 글자로 둔다."""
    return [f"'{v}" if isinstance(v, str) and v[:1] in ("=", "+", "-", "@", "\t", "\r") else v for v in row]


def write_xlsx(path, ok, flagged):
    wb = Workbook()
    ws = wb.active
    ws.title = "경비"
    head = ["날짜", "가맹점", "품목", "공급가액", "부가세", "합계", "결제수단", "파일"]
    ws.append(head)
    for r in sorted(ok, key=lambda r: r["date"]):
        items = r["items"]
        desc = items[0]["name"] + (f" 외 {len(items) - 1}건" if len(items) > 1 else "")
        ws.append(safe([r["date"], r["merchant"], desc, r.get("supply"), r.get("vat"), r["total"], r["payment"], Path(r["file"]).name]))
    n = ws.max_row
    ws.append([])
    ws.append(["합계", f"{len(ok)}건", "", sum(r.get("supply") or 0 for r in ok), sum(r.get("vat") or 0 for r in ok),
               sum(r["total"] for r in ok)])
    for p in sorted({r["payment"] for r in ok}):
        ws.append(["", p, "", "", "", sum(r["total"] for r in ok if r["payment"] == p)])
    for c in ws[n + 2]:
        c.font = Font(bold=True)

    wf = wb.create_sheet("확인 필요")
    wf.append(["파일", "날짜", "가맹점", "합계", "사유"])
    for r, why in flagged:
        wf.append(safe([Path(r["file"]).name, r.get("date"), r.get("merchant"), r.get("total"), " / ".join(why)]))

    for sh in (ws, wf):
        for c in sh[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="1F3A5F")
        for row in sh.iter_rows(min_row=2):
            for c in row:
                if is_int(c.value):
                    c.number_format = "#,##0"
        for col, w in zip("ABCDEFGH", ([12, 18, 22, 12, 10, 12, 10, 14] if sh is ws else [14, 12, 18, 12, 70])):
            sh.column_dimensions[col].width = w
    if flagged:
        wf.sheet_properties.tabColor = "E8590C"
    wb.save(path)


def main(out_dir, input_dir):
    sys.stdout.reconfigure(encoding="utf-8")  # Windows cp949 콘솔에서 ✗ 출력
    out = Path(out_dir)
    files = {p.name for p in Path(input_dir).iterdir() if p.suffix.lower() in INPUT_EXT}
    try:
        rows = json.loads((out / "receipts.json").read_text(encoding="utf-8"))
    except (OSError, ValueError) as e:
        print(f"  ✗ receipts.json 을 읽지 못했다: {e}")
        return 1
    errors, ok, flagged = review(rows, files)
    if errors:
        for e in errors:
            print("  ✗", e)
        print(f"형식 오류 {len(errors)}건 — receipts.json 을 고치고 다시 돌린다 (엑셀은 만들지 않았다)")
        return 1
    xlsx = out / f"{out.name}.xlsx"
    write_xlsx(xlsx, ok, flagged)
    print(f"{xlsx}  (통과 {len(ok)}건 / 확인 필요 {len(flagged)}건, 통과분 합계 {sum(r['total'] for r in ok):,}원)")
    for r, why in flagged:
        print(f"  ? {Path(r['file']).name}: {' / '.join(why)}")
    return 0


if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    sys.exit(main(*sys.argv[1:]))
