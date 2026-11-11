# receipt-to-sheet

폴더에 모아 둔 영수증·청구서(사진·PDF)를 Claude Code가 읽어 **경비 정리 엑셀(xlsx)**로 만든다.
Claude는 읽고 옮겨 적기만 하고, 금액이 맞는지는 코드가 다시 계산한다.

```
samples/*.png·*.pdf ──▶ Claude Code (receipt-to-sheet 스킬) ──▶ out/<이름>/receipts.json
                                                                  │
                          python check.py ◀───────────────────────┘
                          ├ 형식 검사: 날짜 YYYY-MM-DD, 금액 정수, 필수 칸, 입력 파일 전부 한 줄씩  → 걸리면 다시
                          ├ 검산: 수량×단가 = 금액, 품목 합 = 합계(또는 공급가액), 공급가액+부가세 = 합계, 중복
                          └ out/<이름>/<이름>.xlsx — 통과분은 "경비" 시트, 검산에 걸린 건 "확인 필요" 시트
```

## 쓰는 법

```bash
pip install openpyxl
claude    # 이 폴더에서 열면 .claude/skills/receipt-to-sheet 가 잡힌다
> samples 폴더 영수증으로 10월 경비 정리 엑셀 만들어줘
```

비대화형:

```bash
claude -p "samples 폴더 영수증들로 10월 경비 정리 엑셀 만들어줘. 이름은 10월경비" \
  --allowedTools "Read,Glob,Grep,Skill,Bash(python:*),Edit(out/**)"
```

**끝나면 `verify.py` 로 확인한다.** 작업 폴더는 세션이 고칠 수 있어서 믿지 않는다.
원본 저장소를 임시 폴더에 새로 받아 `receipts.json` 하나만 가져가 원본 검사기로 다시 검사하고,
작업 폴더의 검사기(`check.py`·`test_check.py`·`verify.py`)가 원본과 다르면 덮어쓰고 실패 처리한다.

```bash
python verify.py out/10월경비 samples
```

## 검사기가 하는 일과 못 하는 일

검사기는 원본 이미지를 보지 못하고 Claude가 쓴 JSON만 본다. 그래서

- **형식 오류**(빠진 칸, `26.10.05` 같은 날짜, `"14,000"` 같은 문자열 금액, 입력 파일 누락)는 실패(종료 코드 1) — Claude가 고쳐서 다시 돌린다.
- **검산 걸림**(합이 안 맞음, 중복 의심, Claude가 판독 불가로 표시)은 실패가 아니라 엑셀의 "확인 필요" 시트로 보낸다. 원본 청구서 자체가 틀렸을 수도 있어서, 합을 맞추려고 숫자를 고치면 안 된다.
- **못 하는 일:** 숫자를 잘못 읽었는데 합까지 맞게 잘못 읽으면 못 잡는다. 품목 이름·가맹점 이름 오타도 못 잡는다.

`python test_check.py` 로 검사기 자체를 확인한다.

## 폴더

| 경로 | 내용 |
|---|---|
| `.claude/skills/receipt-to-sheet/SKILL.md` | Claude가 따르는 읽기 규칙과 JSON 스키마 |
| `check.py` | 형식 검사 + 검산 + xlsx |
| `verify.py` | 세션이 끝난 뒤 원본 저장소를 새로 받아 그쪽 검사기로 재검사 |
| `samples/` | 예시 영수증·청구서 10장 — **가상 회사 '예시랩'의 지어낸 자료**(가맹점·사업자번호·카드번호 모두 가짜) |
| `samples_blur/` | 흐림 강도 실험용 사진 3장(같은 영수증을 더 흐리게) |
| `out/` | 위 자료로 실제 만든 결과(receipts.json·xlsx) — `10월경비`, `흐림시험` |

예시 영수증 생성기와 정답표는 이 저장소에 넣지 않았다. 같은 폴더에 있으면 영수증을 읽는 세션이 정답을 읽을 수 있어서다.
