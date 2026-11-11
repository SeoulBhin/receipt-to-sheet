"""영수증 정리 세션이 끝난 뒤 사람이(또는 CI가) 돌리는 최종 확인.

    python verify.py out/10월경비 samples

작업 폴더는 세션이 고칠 수 있다(검사기·커밋·git 설정까지). 그래서 작업 폴더의 검사기는 믿지 않는다.

1. 원본 저장소(REPO)를 임시 폴더에 새로 받는다. 작업 폴더의 git 설정·훅은 타지 않는다.
2. 작업 폴더의 검사기(check.py·test_check.py·verify.py)가 원본과 다르면 원본으로 덮어쓰고 실패 처리한다.
3. receipts.json 한 파일만 원본 쪽에 복사해서, 원본 검사기로 test_check.py 와 check.py 를 돌린다.
   통과하면 만들어진 xlsx 를 결과 폴더로 가져온다.

REPO 는 환경변수 RECEIPT_REPO 로 바꿀 수 있다(로컬 경로도 된다). verify.py 자체도 세션이 고칠 수 있으니
확실히 하려면 세션 밖에서 받은 사본으로 돌린다.
"""
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

REPO = os.environ.get("RECEIPT_REPO", "https://github.com/SeoulBhin/receipt-to-sheet.git")
CHECKER = ["check.py", "test_check.py", "verify.py"]
HERE = Path(__file__).parent.resolve()


def same(a: Path, b: Path) -> bool:
    return a.exists() and a.read_bytes().replace(b"\r\n", b"\n") == b.read_bytes().replace(b"\r\n", b"\n")


def put(src: Path, dst: Path):
    """dst 에 쓴다. 링크를 따라 저장소 밖에 쓰지 않게, 같은 폴더에 임시 파일로 쓴 뒤 이름을 바꿔 끼운다."""
    if not dst.parent.resolve().is_relative_to(HERE):
        sys.exit(f"✗ {dst} 가 저장소 밖을 가리킨다(링크된 폴더) — 쓰지 않고 멈춘다")
    fd, tmp = tempfile.mkstemp(dir=dst.parent)
    os.close(fd)
    shutil.copyfile(src, tmp)
    os.replace(tmp, dst)


def main(out_dir, input_dir):
    sys.stdout.reconfigure(encoding="utf-8")
    out = Path(out_dir).absolute()
    src = out / "receipts.json"
    if src.is_symlink() or not src.resolve().is_relative_to(HERE):
        sys.exit("✗ receipts.json 이 링크다 — 멈춘다")
    given = Path(os.path.abspath(input_dir))   # 링크는 따라가지 않고 .. 만 정리
    for d in (given, HERE / "samples"):
        if d.is_symlink() or d.is_junction():
            sys.exit(f"✗ {d} 가 링크다(다른 폴더를 원본처럼 보이게 할 수 있다) — 멈춘다")
    inputs = str(given)
    use_clean_samples = given == HERE / "samples" or given.resolve().is_relative_to((HERE / "samples").resolve())
    with tempfile.TemporaryDirectory() as tmp:
        clean = Path(tmp) / "repo"
        r = subprocess.run(["git", "-c", "core.hooksPath=", "clone", "-q", "--depth", "1", REPO, str(clean)],
                           cwd=tmp, capture_output=True, text=True, encoding="utf-8")
        if r.returncode:
            sys.exit("✗ 원본 저장소를 받지 못했다(주소·네트워크 확인)")
        changed = [f for f in CHECKER if not same(HERE / f, clean / f)]
        if changed:
            print("✗ 검사기가 바뀌었다:", ", ".join(changed), "→ 원본으로 덮어쓴다", flush=True)
            for f in changed:
                put(clean / f, HERE / f)
        # 저장소에 든 예시 영수증으로 검사할 때는 작업 폴더 쪽 파일을 믿지 않는다 — 세션이 원본 이미지를 바꿔
        # 잘못 읽은 값을 "맞게" 만들 수 있어서. 원본 쪽 samples 로 검사하고, 다르면 실패 처리한다.
        if use_clean_samples:
            bad = sorted({p.name for p in (HERE / "samples").iterdir()} ^ {p.name for p in (clean / "samples").iterdir()})
            bad += [p.name for p in (clean / "samples").iterdir() if not same(HERE / "samples" / p.name, p)]
            if bad:
                print("✗ 예시 영수증이 원본과 다르다:", ", ".join(sorted(set(bad))), "→ 원본 쪽 파일로 검사한다", flush=True)
                changed.append("samples")
            inputs = str(clean / "samples")
        dst = clean / "out" / out.name
        dst.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(src, dst / "receipts.json")
        # -E: PYTHONPATH 같은 환경변수로 끼어드는 코드를 막는다. -s·-I 는 안 쓴다(openpyxl 이 사용자 site-packages 에 있음).
        # 한계: 세션은 같은 Windows 사용자로 돈다. 사용자 site-packages·python.exe 를 고칠 수 있는 세션이면
        # 어떤 플래그로도 못 막는다 — 그 수준은 세션을 다른 계정·샌드박스에서 돌려서 막을 일이다
        py = [sys.executable, "-E"]
        ok = subprocess.run([*py, "test_check.py"], cwd=clean).returncode == 0
        ok = subprocess.run([*py, "check.py", f"out/{out.name}", inputs], cwd=clean).returncode == 0 and ok
        if ok:
            put(dst / f"{out.name}.xlsx", out / f"{out.name}.xlsx")
    return 0 if ok and not changed else 1


if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    sys.exit(main(*sys.argv[1:]))
