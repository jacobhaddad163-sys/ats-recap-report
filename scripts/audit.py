"""
Daily Audit Script — All Haddad Brands Apps
Outputs a clean Markdown report (audit-report.md) for GitHub Issues.

Usage:
  python audit.py          # Audit only ATS Recap Report
  python audit.py --all    # Audit all 4 apps
"""

import os
import sys
import py_compile
import re
import importlib
from datetime import datetime
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parent
REPO_ROOT = SCRIPT_DIR.parent

# When --all flag is used, scan all 4 app directories (relative to workspace root)
APPS = {
    "ATS Recap Report": {
        "dir": "ats-recap-report",
        "modules": ["utils", "utils.auth", "utils.ats_parser", "utils.excel_generator"],
    },
    "Confirmed Deals Recap": {
        "dir": "confirmed-deals-recap",
        "modules": ["utils", "utils.auth", "utils.po_parser", "utils.recap_builder"],
    },
    "GM Sheet Builder": {
        "dir": "gm-sheet-builder",
        "modules": ["utils", "utils.auth"],
    },
    "Store Recap Builder": {
        "dir": "store-recap-builder",
        "modules": [],
    },
}

SECRET_PATTERNS = [
    (r'(?i)password\s*=\s*["\'][^"\']+["\']', "Hardcoded password"),
    (r'(?i)api_key\s*=\s*["\'][^"\']+["\']', "Hardcoded API key"),
    (r'(?i)token\s*=\s*["\'][^"\']+["\']', "Hardcoded token"),
    (r'(?i)secret\s*=\s*["\'][^"\']+["\']', "Hardcoded secret"),
    (r'sk-ant-[a-zA-Z0-9_-]{20,}', "Exposed Anthropic API key"),
]

SECURITY_PATTERNS = [
    (r'\beval\s*\(', "eval()"),
    (r'\bexec\s*\(', "exec()"),
    (r'\bos\.system\s*\(', "os.system()"),
    (r'subprocess\.\w+\s*\(.*shell\s*=\s*True', "subprocess with shell=True"),
    (r'\b__import__\s*\(', "dynamic __import__()"),
    (r'\bpickle\.loads?\s*\(', "pickle deserialization"),
]

SKIP_DIRS = {"__pycache__", ".git", "venv", ".venv", "node_modules", "scripts",
             ".claude", "swatch_component"}


def find_py_files(root):
    return sorted(
        p for p in root.rglob("*.py")
        if not any(skip in p.parts for skip in SKIP_DIRS)
    )


def check_syntax(py_files, root):
    failures = []
    for f in py_files:
        try:
            py_compile.compile(str(f), doraise=True)
        except py_compile.PyCompileError as exc:
            failures.append((f.relative_to(root), str(exc)))
    return failures


def check_secrets(py_files, root):
    findings = []
    for f in py_files:
        if f.name == "audit.py":
            continue
        try:
            lines = f.read_text(encoding="utf-8", errors="ignore").splitlines()
        except Exception:
            continue
        for lineno, line in enumerate(lines, 1):
            if line.lstrip().startswith("#"):
                continue
            for pattern, label in SECRET_PATTERNS:
                if re.search(pattern, line):
                    # Skip st.secrets.get() calls — those are safe
                    if "st.secrets" in line or "secrets.get" in line:
                        continue
                    findings.append((f.relative_to(root), lineno, label, line.strip()))
    return findings


def check_security(py_files, root):
    findings = []
    for f in py_files:
        if f.name == "audit.py":
            continue
        try:
            lines = f.read_text(encoding="utf-8", errors="ignore").splitlines()
        except Exception:
            continue
        for lineno, line in enumerate(lines, 1):
            if line.lstrip().startswith("#"):
                continue
            for pattern, label in SECURITY_PATTERNS:
                if re.search(pattern, line):
                    findings.append((f.relative_to(root), lineno, label, line.strip()))
    return findings


def check_gitignore(root):
    gitignore = root / ".gitignore"
    if not gitignore.exists():
        return False, "`.gitignore` not found"
    content = gitignore.read_text(encoding="utf-8", errors="ignore")
    for line in content.splitlines():
        if "secrets.toml" in line and not line.strip().startswith("#"):
            return True, "Protected"
    return False, "secrets.toml is **NOT** in .gitignore"


def check_imports(root, modules):
    if not modules:
        return []
    sys.path.insert(0, str(root))
    failures = []
    for mod in modules:
        try:
            importlib.import_module(mod)
        except Exception as exc:
            failures.append((mod, f"{type(exc).__name__}: {exc}"))
    if str(root) in sys.path:
        sys.path.remove(str(root))
    return failures


def check_requirements(root):
    req = root / "requirements.txt"
    if not req.exists():
        return False, "Not found"
    lines = [l for l in req.read_text().splitlines() if l.strip() and not l.startswith("#")]
    if not lines:
        return False, "Empty"
    return True, f"{len(lines)} deps"


def audit_app(app_name, app_root, modules):
    """Audit one app. Returns (results_dict, details_lines, has_issues)."""
    py_files = find_py_files(app_root)
    results = {}
    details = []
    has_issues = False

    # 1. Syntax
    syntax_fails = check_syntax(py_files, app_root)
    results["Syntax"] = "✅" if not syntax_fails else "❌"

    # 2. Secrets
    secret_finds = check_secrets(py_files, app_root)
    results["Secrets"] = "✅" if not secret_finds else "❌"

    # 3. Gitignore
    gi_ok, gi_msg = check_gitignore(app_root)
    results["Protected"] = "✅" if gi_ok else "❌"

    # 4. Imports
    import_fails = check_imports(app_root, modules)
    results["Imports"] = "✅" if not import_fails else "⚠️"

    # 5. Security
    security_finds = check_security(py_files, app_root)
    results["Security"] = "✅" if not security_finds else "❌"

    # 6. Requirements
    req_ok, req_msg = check_requirements(app_root)
    results["Deps"] = "✅" if req_ok else "❌"

    # Collect details for failures
    if syntax_fails:
        has_issues = True
        details.append(f"**Syntax Errors in {app_name}:**")
        for path, err in syntax_fails:
            details.append(f"- `{path}` — {err}")
        details.append(f"  *How to fix:* Open the file, fix the typo or missing bracket on that line.")
        details.append("")

    if secret_finds:
        has_issues = True
        details.append(f"**Hardcoded Secrets in {app_name}:**")
        for path, lineno, label, code in secret_finds:
            details.append(f"- `{path}` line {lineno} — {label}")
            details.append(f"  ```python\n  {code}\n  ```")
        details.append(f"  *How to fix:* Move to `.streamlit/secrets.toml` and use `st.secrets[\"KEY\"]` instead.")
        details.append("")

    if not gi_ok:
        has_issues = True
        details.append(f"**Secrets Not Protected in {app_name}:** {gi_msg}")
        details.append(f"  *How to fix:* Add `.streamlit/secrets.toml` to `.gitignore`.")
        details.append("")

    if import_fails:
        details.append(f"**Import Warnings in {app_name}:**")
        for mod, err in import_fails:
            details.append(f"- `{mod}` — {err}")
        details.append(f"  *Note:* May be expected in CI without Streamlit runtime.")
        details.append("")

    if security_finds:
        has_issues = True
        details.append(f"**Security Concerns in {app_name}:**")
        for path, lineno, label, code in security_finds:
            details.append(f"- `{path}` line {lineno} — {label}")
            details.append(f"  ```python\n  {code}\n  ```")
        details.append(f"  *How to fix:* Replace with safer alternatives. Avoid `eval()`, `exec()`, `shell=True`.")
        details.append("")

    if not req_ok:
        has_issues = True
        details.append(f"**Dependencies Issue in {app_name}:** {req_msg}")
        details.append("")

    return results, details, has_issues, len(py_files)


def main():
    now = datetime.now()
    use_all = "--all" in sys.argv
    lines = []
    any_issues = False

    def out(s=""):
        lines.append(s)

    out("## 🔍 Daily App Audit — Haddad Brands")
    out(f"**Date:** {now.strftime('%B %d, %Y at %I:%M %p')}")
    out("")

    # Determine which apps to audit
    if use_all:
        workspace = REPO_ROOT.parent  # parent of ats-recap-report checkout
        apps_to_audit = {}
        for name, info in APPS.items():
            app_dir = workspace / info["dir"]
            if app_dir.exists():
                apps_to_audit[name] = (app_dir, info["modules"])
            else:
                out(f"> ⚠️ **{name}** directory not found at `{app_dir}`")
    else:
        apps_to_audit = {"ATS Recap Report": (REPO_ROOT, APPS["ATS Recap Report"]["modules"])}

    out(f"**Apps audited:** {len(apps_to_audit)}")
    out("")

    # --- Summary Table ---
    out("| App | Syntax | Secrets | Protected | Imports | Security | Deps | Files |")
    out("|-----|--------|---------|-----------|---------|----------|------|-------|")

    all_details = []

    for app_name, (app_root, modules) in apps_to_audit.items():
        results, details, has_issues, file_count = audit_app(app_name, app_root, modules)
        if has_issues:
            any_issues = True
        all_details.extend(details)

        out(f"| {app_name} | {results['Syntax']} | {results['Secrets']} | {results['Protected']} | {results['Imports']} | {results['Security']} | {results['Deps']} | {file_count} |")

    out("")

    # --- Details ---
    if all_details:
        out("---")
        out("### Issues & Warnings")
        out("")
        for line in all_details:
            out(line)
    else:
        out("### ✅ All apps are healthy!")
        out("No issues found across any app. Everything is working correctly.")
        out("")

    out("---")
    out("*Generated automatically by GitHub Actions. Runs daily at 7:00 AM EST.*")

    report = "\n".join(lines)
    try:
        print(report)
    except UnicodeEncodeError:
        print(report.encode("ascii", errors="replace").decode("ascii"))

    # Write report file — in workspace root for CI, or repo root for local
    if use_all:
        report_path = REPO_ROOT.parent / "audit-report.md"
    else:
        report_path = REPO_ROOT / "audit-report.md"
    report_path.write_text(report, encoding="utf-8")

    sys.exit(1 if any_issues else 0)


if __name__ == "__main__":
    main()
