# -*- coding: utf-8 -*-
"""Cross-check self-test that does not depend on PowerShell execution policy.

Runs the tool's own `--selftest` entry point in a subprocess, then asserts on
the structured log it writes. This is the portable path: `tools/selftest.ps1`
does the same thing but is a convenience wrapper for PowerShell users.

    python tools/selftest.py SAMPLE.docx
    python tools/selftest.py SAMPLE.docx --engine word
    python tools/selftest.py FOLDER --engine wps --keep
    python tools/selftest.py --check                # environment only

Exit code 0 = all assertions passed, 1 = at least one failed.
"""
import os
import subprocess
import sys
import tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ENTRY = os.path.join(ROOT, "src", "pdf_tool.py")
LOG = os.path.join(tempfile.gettempdir(), "pdf_tool_selftest.log")


def run_check(engine):
    cmd = [sys.executable, os.path.join(ROOT, "src", "wps_to_pdf.py"),
           "--engine", engine, "--check"]
    print("Running: %s" % " ".join(cmd))
    proc = subprocess.run(cmd, cwd=ROOT)
    return proc.returncode


def run_selftest(sample, engine, keep):
    if os.path.exists(LOG):
        os.remove(LOG)
    cmd = [sys.executable, ENTRY, "--selftest", os.path.abspath(sample), engine]
    if keep:
        cmd.append("keep")
    print("Running: %s" % " ".join(cmd))
    proc = subprocess.run(cmd, cwd=ROOT)
    if not os.path.exists(LOG):
        print("FAILED: the tool produced no selftest log at %s" % LOG,
              file=sys.stderr)
        return 1, None
    with open(LOG, encoding="utf-8", errors="replace") as handle:
        lines = [line.rstrip("\n") for line in handle]
    return proc.returncode, lines


def assert_log(lines, keep):
    failures = []
    if "RESULT: OK" not in lines:
        failures.append("RESULT is not OK")

    reports = [l for l in lines if l.startswith("[") and "] out = " in l]
    if not reports:
        failures.append("no per-document report found")

    # Documents without pictures skip optimisation: the exported PDF *is* the
    # deliverable, so pdf_removed = False is expected for them.
    skipped = set()
    for line in lines:
        if line.startswith("[") and "] skipped = True" in line:
            skipped.add(line[1:line.index("]")])

    for line in lines:
        if not line.startswith("["):
            continue
        idx = line[1:line.index("]")]
        key, _, value = line.partition(" = ")
        key = key.split("] ", 1)[-1].strip()
        value = value.strip()
        if key == "diff" and value.isdigit() and int(value) != 0:
            failures.append("document %s reports %s pixel differences"
                            % (idx, value))
        elif key == "final_exists" and value == "False":
            failures.append("document %s produced no final PDF" % idx)
        elif key == "interim_exists_after" and value == "True" and not keep:
            failures.append("document %s left the interim PDF behind" % idx)
        elif key == "pages" and value.isdigit() and int(value) <= 0:
            failures.append("document %s exported zero pages" % idx)
        elif (key == "pdf_removed" and value == "False" and not keep
              and idx not in skipped):
            failures.append("document %s kept the interim PDF without being "
                            "asked to" % idx)
    return failures


def main(argv):
    if not argv or argv[0] in ("-h", "--help"):
        print(__doc__)
        return 0

    engine = "wps"
    keep = False
    if "--engine" in argv:
        engine = argv[argv.index("--engine") + 1]
    if "--keep" in argv:
        keep = True

    if "--check" in argv:
        return run_check(engine)

    sample = argv[0]
    if not os.path.exists(sample):
        print("sample not found: %s" % sample, file=sys.stderr)
        return 1

    code, lines = run_selftest(sample, engine, keep)
    if lines is None:
        return 1

    for line in lines:
        print("  " + line)

    failures = assert_log(lines, keep)
    documents = len([l for l in lines if "] out = " in l])
    print()
    if failures:
        print("SELFTEST FAILED:")
        for f in failures:
            print("  - %s" % f)
        return 1
    print("SELFTEST PASSED (%d document(s))" % documents)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
