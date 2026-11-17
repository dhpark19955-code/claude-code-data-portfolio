# -*- coding: utf-8 -*-
"""공개 전 점검. 저장소에 올리면 안 되는 것이 섞였는지 본다.

  1) 비밀값 모양의 문자열 (값은 출력하지 않고 파일과 줄 번호만)
  2) 사용자 이름이 든 절대경로, 메일 주소
  3) 금지어: tools/denylist.local.txt 에 한 줄에 하나씩 (회사명, 종목명, 사람 이름 등)
     이 파일은 .gitignore 에 있어 저장소에 올라가지 않는다
  4) 올리면 안 되는 형식의 파일 (허용 목록 밖의 xlsx, pdf 등)

실행:  python -I -X utf8 tools/prepublish_scan.py     (문제가 있으면 종료 코드 1)
"""
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SKIP_DIRS = {".git", "__pycache__", ".venv", "node_modules"}
TEXT_EXT = {".md", ".py", ".html", ".txt", ".csv", ".json", ".yaml", ".yml", ".example", ".ps1", ".bat", ".cmd", ""}
BLOCK_EXT = {".xlsx", ".xlsm", ".xls", ".pdf", ".docx", ".pptx", ".env", ".pem", ".key", ".zip"}
ALLOW_FILES = {"demo/output/sample_dataset_ib.xlsx", "demo/output/sample_model_dcf.xlsx"}

PATTERNS = [
    ("비밀값 대입", re.compile(r"(?i)(api[_-]?key|secret|token|passw(or)?d)\s*[:=]\s*['\"]?[A-Za-z0-9_\-+/=]{12,}")),
    ("긴 키 모양 문자열", re.compile(r"\b(sk-[A-Za-z0-9_\-]{20,}|AKIA[0-9A-Z]{16}|ghp_[A-Za-z0-9]{30,}|AIza[0-9A-Za-z_\-]{30,})")),
    ("사용자 절대경로", re.compile(r"(?i)([A-Z]:\\Users\\[^\\\s<%]+|/Users/[^/\s<]+|/home/[^/\s<]+)")),
    ("메일 주소", re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")),
]
ALLOW_MAIL = {"noreply@anthropic.com", "dhpark19955@gmail.com"}  # 작성자 연락처는 공개용


def load_denylist():
    p = os.path.join(ROOT, "tools", "denylist.local.txt")
    if not os.path.exists(p):
        return None
    with open(p, encoding="utf-8") as f:
        return [s.strip() for s in f if s.strip() and not s.startswith("#")]


def main():
    deny = load_denylist()
    findings = []
    for base, dirs, files in os.walk(ROOT):
        dirs[:] = [d for d in dirs if d not in SKIP_DIRS]
        for name in files:
            full = os.path.join(base, name)
            rel = os.path.relpath(full, ROOT).replace("\\", "/")
            ext = os.path.splitext(name)[1].lower()
            if rel == "tools/denylist.local.txt":
                continue
            if ext in BLOCK_EXT and rel not in ALLOW_FILES:
                findings.append((rel, 0, "올리면 안 되는 형식"))
                continue
            if ext not in TEXT_EXT:
                continue
            try:
                with open(full, encoding="utf-8") as f:
                    lines = f.readlines()
            except (UnicodeDecodeError, OSError):
                continue
            for i, line in enumerate(lines, 1):
                for label, rx in PATTERNS:
                    m = rx.search(line)
                    if m and not (label == "메일 주소" and m.group(0) in ALLOW_MAIL):
                        if rel == "tools/prepublish_scan.py":
                            continue  # 이 파일의 패턴 정의
                        findings.append((rel, i, label))
                for word in deny or []:
                    if word.lower() in line.lower():
                        findings.append((rel, i, "금지어"))

    for rel, i, label in findings:
        print("%s:%d  %s" % (rel, i, label))
    print("금지어 목록: %s" % ("없음 (tools/denylist.local.txt 를 만들면 함께 검사)" if deny is None else "%d개" % len(deny)))
    print("findings=%d" % len(findings))
    return 1 if findings else 0


if __name__ == "__main__":
    sys.exit(main())
