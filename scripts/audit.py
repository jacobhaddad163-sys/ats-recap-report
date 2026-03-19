"""
Daily Audit Script — ATS Recap Report Builder
Outputs a clean Markdown report (audit-report.md) for GitHub Issues.
"""

import os
import sys
import py_compile
import re
import importlib
from datetime import datetime
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parent.parent

SECRET_PATTERNS = [
    (r'(?i)password\s*=\s*["\'][^"\']+["\']', "Hardcoded password"),
    (r'(?i)api_key\s*=\s*["\'][^"\']+["\']', "Hardcoded API key"),
    (r'(?i)token\s*=\s*["\'][^"\']+["\']', "Hardcoded token"),
    (r'(?i)secret\s*=\s*["\'][^"\']+["\']', "Hardcoded secret"),
    (r'sk-ant-[a-zA-Z0-9_-]{20,}', "Exposed Anthropic API key"),
    (r'(?i)aws_access_key_id\s*=\s*["\'][^"\']+["\']', "AWS access key"),
]

SECURITY_PATTERNS = [
    (r'\beval\s*\(', "eval()"),
    (r'\bexec\s*\(', "exec()"),
    (r'\bos\.system\s*\(', "os.system()"),
    (r'subprocess\.\w+\s*\(.*shell\s*=\s*True', "subprocess with shell=True"),
    (r'\b__import__\s*\(', "dynamic __import__()"),
    (r'\bpickle\.loads?\s*\(', "pickle deserialization"),
]

SKIP_DIRS = {"__pycache__", ".git", "venv", ".venv", "node_modules", "scripts"}


def find_py_files(root):
    return sorted(
        p for p in root.rglob("*.py")
        if not any(skip in p.parts for skip in SKIP_DIRS)
    )


def check_syntax(py_files):
    failures = []
    for f in py_files:
        try:
            py_compile.compile(str(f), doraise=True)
        except py_compile.PyCompileError as exc:
            failures.append((f.relative_to(REPO_ROOT), str(exc)))
    return failures


def check_secrets(py_files):
    findings = []
    for f in py_files:
        try:
            lines = f.read_text(encoding="utf-8", errors="ignore").splitlines()
        except Exception:
            continue
        for lineno, line in enumerate(lines, 1):
            if line.lstrip().startswith("#"):
                continue
            for pattern, label in SECRET_PATTERNS:
                if re.search(pattern, line):
                    findings.append((f.relative_to(REPO_ROOT), lineno, label, line.strip()))
    return findings


def check_security(py_files):
    findings = []
    for f in py_files:
        try:
            lines = f.read_text(encoding="utf-8", errors="ignore").splitlines()
        except Exception:
            continue
        for lineno, line in enumerate(lines, 1):
            if line.lstrip().startswith("#"):
                continue
            for pattern, label in SECURITY_PATTERNS:
                if re.search(pattern, line):
                    findings.append((f.relative_to(REPO_ROOT), lineno, label, line.strip()))
    return findings


def check_gitignore():
    gitignore = REPO_ROOT / ".gitignore"
    if not gitignore.exists():
        return False, "`.gitignore` not found"
    content = gitignore.read_text(encoding="utf-8", errors="ignore")
    for line in content.splitlines():
        if "secrets.toml" in line and not line.strip().startswith("#"):
            return True, "secrets.toml is in .gitignore"
    return False, "secrets.toml is **NOT** in .gitignore"


def check_imports():
    sys.path.insert(0, str(REPO_ROOT))
    failures = []
    for mod in ["utils", "utils.auth", "utils.ats_parser", "utils.excel_generator"]:
        try:
            importlib.import_module(mod)
        except Exception as exc:
            failures.append((mod, f"{type(exc).__name__}: {exc}"))
    return failures


def check_requirements():
    req = REPO_ROOT / "requirements.txt"
    if not req.exists():
        return False, "requirements.txt not found"
    lines = [l for l in req.read_text().splitlines() if l.strip() and not l.startswith("#")]
    if not lines:
        return False, "requirements.txt is empty"
    return True, f"{len(lines)} dependencies listed"


def main():
    now = datetime.now()
    py_files = find_py_files(REPO_ROOT)
    lines = []
    has_issues = False

    def out(s=""):
        lines.append(s)

    # --- Header ---
    out(f"## ATS Recap Report Builder — Daily Audit")
    out(f"**Date:** {now.strftime('%B %d, %Y at %I:%M %p')}")
    out(f"**Files scanned:** {len(py_files)}")
    out("")

    # --- Summary Table ---
    results = {}

    # 1. Syntax
    syntax_fails = check_syntax(py_files)
    results["Syntax Check"] = len(syntax_fails) == 0

    # 2. Secrets
    secret_finds = check_secrets(py_files)
    results["Hardcoded Secrets"] = len(secret_finds) == 0

    # 3. Gitignore
    gi_ok, gi_msg = check_gitignore()
    results["Secrets Protected"] = gi_ok

    # 4. Imports
    import_fails = check_imports()
    results["Module Imports"] = len(import_fails) == 0

    # 5. Security
    security_finds = check_security(py_files)
    results["Security Scan"] = len(security_finds) == 0

    # 6. Requirements
    req_ok, req_msg = check_requirements()
    results["Dependencies"] = req_ok

    out("| Check | Status |")
    out("|-------|--------|")
    for check, passed in results.items():
        icon = "✅ Pass" if passed else "❌ Fail"
        out(f"| {check} | {icon} |")
        if not passed:
            has_issues = True
    out("")

    # --- Details (only if issues) ---
    if syntax_fails:
        out("### ❌ Syntax Errors")
        out("These files have Python syntax errors and will crash on import:")
        out("")
        for path, err in syntax_fails:
            out(f"- **`{path}`** — {err}")
        out("")
        out("**How to fix:** Open the file, find the line mentioned in the error, fix the typo or missing bracket.")
        out("")

    if secret_finds:
        out("### ❌ Hardcoded Secrets Found")
        out("Passwords or API keys found directly in the code:")
        out("")
        for path, lineno, label, code in secret_finds:
            out(f"- **`{path}` line {lineno}** — {label}")
            out(f"  ```python")
            out(f"  {code}")
            out(f"  ```")
        out("")
        out("**How to fix:** Move these to `.streamlit/secrets.toml` and use `st.secrets[\"KEY_NAME\"]` instead.")
        out("")

    if not gi_ok:
        out("### ❌ Secrets Not Protected")
        out(f"{gi_msg}")
        out("")
        out("**How to fix:** Add `.streamlit/secrets.toml` to your `.gitignore` file immediately.")
        out("")

    if import_fails:
        out("### ⚠️ Import Failures")
        out("These modules failed to import:")
        out("")
        for mod, err in import_fails:
            out(f"- **`{mod}`** — {err}")
        out("")
        out("**Note:** Import failures in CI may be expected if Streamlit isn't running. Check if this also happens locally with `python -c \"import {module}\"`.")
        out("")

    if security_finds:
        out("### ❌ Security Concerns")
        out("Potentially dangerous function calls found:")
        out("")
        for path, lineno, label, code in security_finds:
            out(f"- **`{path}` line {lineno}** — {label}")
            out(f"  ```python")
            out(f"  {code}")
            out(f"  ```")
        out("")
        out("**How to fix:** Replace with safer alternatives. `eval()`/`exec()` should almost never be used. `subprocess` with `shell=True` is a command injection risk.")
        out("")

    if not req_ok:
        out(f"### ❌ Dependencies Issue")
        out(f"{req_msg}")
        out("")

    # --- All clear ---
    if not has_issues and not import_fails:
        out("### ✅ Everything looks good!")
        out("No issues found. All checks passed. Your app is healthy.")
        out("")

    out("---")
    out("*This report was generated automatically by GitHub Actions.*")

    report = "\n".join(lines)
    try:
        print(report)
    except UnicodeEncodeError:
        print(report.encode("ascii", errors="replace").decode("ascii"))

    # Write markdown file for the workflow
    (REPO_ROOT / "audit-report.md").write_text(report, encoding="utf-8")

    sys.exit(1 if has_issues else 0)


if __name__ == "__main__":
    main()
