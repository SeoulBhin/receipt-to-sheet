"""python test_check.py — 검산이 맞게 잡고 맞게 통과시키는지. 값은 예시 영수증과 무관하게 지어낸 것."""
import copy
import tempfile
from pathlib import Path

from openpyxl import load_workbook

from check import review, write_xlsx

OK = {"file": "samples/a.png", "date": "2026-01-05", "merchant": "가상상점", "payment": "카드",
      "items": [{"name": "물품", "qty": 2, "unit_price": 1100, "amount": 2200}, {"name": "기타", "qty": 1, "unit_price": 1100, "amount": 1100}],
      "supply": 3000, "vat": 300, "total": 3300, "status": "ok", "note": ""}
FILES = {"a.png"}


def run(*rows, files=FILES):
    return review([copy.deepcopy(r) for r in rows], files)


def edit(**kw):
    r = copy.deepcopy(OK)
    r.update(kw)
    return r


# 통과: 부가세 포함 품목 합 = 합계, 공급가+부가세 = 합계
assert run(OK) == ([], [OK], [])
# 통과: 청구서처럼 품목 합 = 공급가액
inv = edit(items=[{"name": "x", "qty": 3, "unit_price": 1000, "amount": 3000}])
assert run(inv)[2] == []
# 통과: 부가세 표기 없는 영수증(supply·vat 둘 다 null)
assert run(edit(supply=None, vat=None))[2] == []

# 형식 오류: 날짜 형식, 없는 날짜, 금액이 문자열, 결제수단, supply 만 있음
assert run(edit(date="26.01.05"))[0] == ["a.png: date 는 YYYY-MM-DD: '26.01.05'"]
assert "없는 날짜" in run(edit(date="2026-02-30"))[0][0]
assert "total 은" in run(edit(total="3,300"))[0][0]
assert "payment" in run(edit(payment="법인카드"))[0][0]
assert "둘 다" in run(edit(vat=None))[0][0]
# 형식 오류: 입력 파일을 빼먹음 / 같은 파일 두 번 / 없는 파일
assert "없다" in run(OK, files={"a.png", "b.png"})[0][0]
assert any("두 번" in e for e in run(OK, OK)[0])
assert any("입력 폴더에 없는" in e for e in run(edit(file="samples/zzz.png"))[0])
# 형식 오류: 확인 필요인데 이유 없음
assert run(edit(status="확인 필요", note=""))[0] == ["a.png: 확인 필요인데 note(이유)가 없다"]

# 확인 필요(오류 아님): 품목 합이 합계·공급가 어느 쪽과도 안 맞음
e, ok, fl = run(edit(total=3400, supply=3100))
assert e == [] and ok == [] and "품목 합 3,300 != 합계 3,400 · 공급가액 3,100" in fl[0][1]
# 확인 필요: 공급가 + 부가세 != 합계
assert "공급가액 3,000 + 부가세 310 != 합계 3,300" in run(edit(vat=310))[2][0][1]
# 확인 필요: 수량×단가 != 금액
bad = edit(items=[{"name": "물품", "qty": 2, "unit_price": 1100, "amount": 2300}, OK["items"][1]])
assert any("품목 1 금액" in w for w in run(bad)[2][0][1])
# 확인 필요: 같은 날짜·가맹점·합계 두 장(파일 이름은 다름, 가맹점 띄어쓰기 차이 무시)
e, ok, fl = run(OK, edit(file="samples/b.png", merchant="가상 상점"), files={"a.png", "b.png"})
assert len(ok) == 1 and "중복 의심" in fl[0][1][0]
# 확인 필요: 읽은 쪽이 판독 불가로 표시 — 금액이 null 이어도 형식 오류 아님
e, ok, fl = run({"file": "samples/a.png", "status": "확인 필요", "note": "합계 부분이 가려짐", "total": None})
assert e == [] and fl[0][1] == ["읽은 쪽 표시: 합계 부분이 가려짐"]
# 형식 오류: 확인 필요라도 적은 칸은 형식이 맞아야 한다
assert any("total 은 정수 또는 null" in e for e in run({"file": "samples/a.png", "status": "확인 필요", "note": "x", "total": "1,000"})[0])
assert any("date 는 YYYY-MM-DD 또는 null" in e for e in run({"file": "samples/a.png", "status": "확인 필요", "note": "x", "date": "1/5"})[0])

# 엑셀: 영수증 글자가 = 로 시작해도 수식으로 들어가지 않는다
evil = edit(merchant='=HYPERLINK("http://x","y")', items=[{"name": "+1+1", "qty": 1, "unit_price": 3300, "amount": 3300}])
with tempfile.TemporaryDirectory() as d:
    _, ok, fl = run(evil)
    write_xlsx(Path(d) / "t.xlsx", ok, fl)
    ws = load_workbook(Path(d) / "t.xlsx")["경비"]
    assert ws["B2"].data_type == "s" and ws["B2"].value.startswith("'=") and ws["C2"].value == "'+1+1", (ws["B2"].value, ws["C2"].value)
print("ok")
