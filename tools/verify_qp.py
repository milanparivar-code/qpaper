#!/usr/bin/env python3
"""
Verification harness for the T1 Python-I question paper built by build_t1_qp.py.

Three real checks, no logic re-implemented:

1. RUNTIME  - every Practice-Book MCQ that is a runnable snippet is executed and
              the marked option is compared with the real output.
2. STATIC   - the "invalid statement" MCQ has every option compiled; exactly one
              must raise SyntaxError and it must be the marked one.
3. DOCUMENT - the three generated .docx sets are re-opened: question marks are
              summed (offline 16 / online 09), the position of the correct
              option is read back out of the paper and compared with the answer
              key file, and the three sets must hold the identical question set
              with the MCQs and their options reshuffled.
"""

import contextlib
import io
import re
import sys
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

sys.path.insert(0, str(Path(__file__).resolve().parent))
from build_t1_qp import MCQS, OUTDIR, Q2, Q3, SET_ORDER  # noqa: E402

FAIL = []

# what kind of proof each Practice Book question needs
RUNTIME = {54: "6", 107: "4", 165: "12", 167: "2 5", 176: "1 5"}
EXPRESSIONS = {64: [512, 64, 512]}        # values of the three expressions
ERRORS = {86: "ZeroDivisionError"}
COMPILE = {45}                            # only one option may be invalid code
MANUAL = {8, 21}                          # mnemonic translator / flow trace


def check(cond, msg):
    print(("  OK   " if cond else "  FAIL ") + msg)
    if not cond:
        FAIL.append(msg)


def run_snippet(code):
    buf = io.StringIO()
    try:
        with contextlib.redirect_stdout(buf):
            exec(code, {})
        return buf.getvalue(), None
    except Exception as exc:                       # noqa: BLE001
        return buf.getvalue(), type(exc).__name__


def compiles(stmt):
    try:
        compile(stmt, "<verify>", "exec")
        return True
    except SyntaxError:
        return False


def verify_answers():
    print("== 1/2. are the marked answers actually correct? ==")
    covered = set(RUNTIME) | set(EXPRESSIONS) | set(ERRORS) | set(COMPILE) | MANUAL
    pbs = [m["pb"] for m in MCQS]
    check(set(pbs) == covered and len(pbs) == len(set(pbs)),
          f"all {len(pbs)} MCQs are covered by exactly one proof method")

    for m in MCQS:
        pb, correct = m["pb"], m["opts"][m["ans"]]
        check(len(m["opts"]) == 6, f"PB{pb}: 6 options (T1 rule)")
        check(len(set(m["opts"])) == 6, f"PB{pb}: all options distinct")

        if pb in MANUAL:
            print(f"  ---- PB{pb}: hand-traced, marked {correct!r}")
            continue
        if pb in COMPILE:
            bad = [o for o in m["opts"] if not compiles(o)]
            check(len(bad) == 1 and bad[0] == correct,
                  f"PB{pb}: exactly one option fails to compile and it is the "
                  f"marked one ({correct!r})")
            continue
        code = "\n".join(m["code"])
        if pb in EXPRESSIONS:
            vals = [eval(line, {}) for line in m["code"]]      # noqa: S307
            marked = [int(x) for x in correct.split(",")]
            check(vals == EXPRESSIONS[pb] and marked == vals,
                  f"PB{pb}: expressions evaluate to {vals}, marked {correct!r}")
            continue
        out, err = run_snippet(code)
        if pb in ERRORS:
            check(err == ERRORS[pb] and correct == "Error",
                  f"PB{pb}: snippet raises {err}, marked option {correct!r}")
            continue
        check(out.strip() == RUNTIME[pb] and correct == RUNTIME[pb],
              f"PB{pb}: printed {out.strip()!r}, marked option {correct!r}")


# --------------------------------------------------------------------------- #
# document checks
# --------------------------------------------------------------------------- #
def set_box_letters(doc):
    out = []
    for txbx in doc.element.body.iter(qn("w:txbxContent")):
        txt = "".join(t.text or "" for t in txbx.iter(qn("w:t")))
        if txt.strip().startswith("Set"):
            out.append(txt.strip())
    return out


def parse_options(cell_text):
    return dict(re.findall(r"\(([A-F])\)\s*([^\t\n]+)", cell_text))


def option_set(cell_text):
    """frozenset of the six option texts (the option *values*, not A-F)."""
    return frozenset(parse_options(cell_text).values())


def verify_documents():
    print("\n== 3. the generated documents ==")
    key_txt = (OUTDIR / "ANSWER KEY_T1_PYTHON-I_MDP_do not circulate.txt").read_text()
    # options are unique per MCQ, the stems are not (several say
    # "What will be the output of the following program on execution?")
    by_opts = {frozenset(m["opts"]): m for m in MCQS}
    check(len(by_opts) == len(MCQS), "option sets identify each MCQ uniquely")

    for letter in ("A", "B", "C"):
        path = OUTDIR / f"SET {letter}_T1_PYTHON-1_TEST PAPER_SEM I_MDP.docx"
        doc = Document(str(path))
        print(f"-- SET {letter} ({path.name})")

        boxes = set_box_letters(doc)
        check(len(boxes) == 4 and all(b == f"Set: {letter}" for b in boxes),
              f"all 'Set' boxes (drawing + VML, both pages) read 'Set: {letter}'")

        head = " ".join(c.text for r in doc.tables[0].rows for c in r.cells)
        check("TEST 1 ( CO1 ) (OFFLINE)" in head, "offline heading = TEST 1 (CO1)")
        check("2/3" not in head and "CO2" not in head,
              "no leftover 2/3 or CO2 in the offline heading")
        check("MAX MARKS: 16 marks" in head, "offline MAX MARKS = 16")
        check("2:15 to 3:30 pm" in head and "29-Sep-2026" in head,
              "offline date/time = 29-Sep-2026, 2:15 to 3:30 pm")
        check("1.25" in head, "offline duration = 1.25 hours")
        check("017012194" in head and "017212194" in head,
              "subject codes of all branches printed")
        check("CE/IT/CSD" in head, "branches printed")

        head_on = " ".join(c.text for r in doc.tables[3].rows for c in r.cells)
        check("TEST 1 ( CO1 ) (ONLINE)" in head_on, "online heading = TEST 1 (CO1)")
        check("MAX MARKS: 09 marks" in head_on, "online MAX MARKS = 09")
        check("4:15 to 5:30 pm" in head_on, "online time = 4:15 to 5:30 pm")

        # ---------------- offline question table ----------------------------
        rows = doc.tables[1].rows
        mcq_marks, desc_marks, mcq_rows = 0, 0, []
        for r in rows[1:]:
            sub, marks, body = (r.cells[1].text.strip(), r.cells[3].text.strip(),
                                r.cells[2].text)
            if not sub.endswith(")"):
                continue
            if re.search(r"\([A-F]\)", body):
                mcq_rows.append((sub, body))
                mcq_marks += int(marks)
            else:
                desc_marks += int(marks)
        check(len(mcq_rows) == 10 and mcq_marks == 10,
              f"Q-1 = {len(mcq_rows)} MCQs, {mcq_marks} marks (from Practice Book)")
        check(desc_marks == 6, f"Q-2 descriptive = {desc_marks} marks (out of PB)")
        check(mcq_marks + desc_marks == 16,
              f"offline total = {mcq_marks + desc_marks} marks")
        check(all(len(parse_options(b)) == 6 for _, b in mcq_rows),
              "every MCQ prints 6 options")
        check(all(option_set(b) in by_opts for _, b in mcq_rows),
              "every MCQ in the paper comes from the Practice Book selection")
        check([by_opts[option_set(b)]["pb"] for _, b in mcq_rows]
              == [MCQS[i]["pb"] for i in SET_ORDER[letter]],
              "MCQ order matches SET_ORDER (reshuffled per set)")

        # ---------------- online question table -----------------------------
        on_rows = doc.tables[4].rows[1:]
        on_marks = sum(int(r.cells[3].text.strip()) for r in on_rows)
        check(len(on_rows) == 2 and on_marks == 9,
              f"online Q-3 = {len(on_rows)} parts, {on_marks} marks (out of PB)")
        on_body = " ".join(r.cells[2].text for r in on_rows)
        check("Collatz" in on_body and "ternary" in on_body,
              "online questions are the out-of-PB ones")

        # ---------------- answer key ----------------------------------------
        block = key_txt.split(f"--- SET {letter} ---")[1].split("--- SET")[0]
        key = dict(re.findall(r"(Q-1 \d+\))\s+(\([A-F]\))", block))
        for sub, body in mcq_rows:
            opts = parse_options(body)
            m = by_opts[frozenset(opts.values())]
            right = [L for L, o in opts.items() if o == m["opts"][m["ans"]]]
            check(len(right) == 1 and key.get(f"Q-1 {sub}") == f"({right[0]})",
                  f"key Q-1 {sub} = ({right[0] if right else '?'}) matches the "
                  f"printed position of the correct option (PB{m['pb']})")

        # ---------------- format rules --------------------------------------
        mains = {r.cells[0].text.strip() for r in rows[1:] if r.cells[0].text.strip()}
        mains |= {r.cells[0].text.strip() for r in on_rows
                  if r.cells[0].text.strip()}
        check(mains == {"Q-1", "Q-2", "Q-3"}, f"main questions = {sorted(mains)}")
        all_text = " ".join(c.text for t in doc.tables for r in t.rows
                            for c in r.cells)
        check(not re.search(r"\bOR\b", all_text), "no OR-type options in the paper")
        check(sum(q["marks"] for q in Q2) == 6 and sum(q["marks"] for q in Q3) == 9,
              "authoring marks: Q-2 = 6, Q-3 = 9")
        check(len(rows) == 14 and len(doc.tables[4].rows) == 3,
              f"tables rebuilt ({len(rows)} and {len(doc.tables[4].rows)} rows)")
        floating = [i for i, t in enumerate(doc.tables)
                    if t._tbl.find(qn("w:tblPr")).find(qn("w:tblpPr")) is not None]
        check(not floating,
              "no floating question table left (a floating table cannot split "
              f"over a page) - floating tables: {floating}")
        check(page_break_present(doc),
              "explicit page break separates the offline and the online paper")
        hdr = doc.tables[1]._tbl.findall(qn("w:tr"))[0].find(qn("w:trPr"))
        check(hdr is not None and hdr.find(qn("w:tblHeader")) is not None,
              "Marks/Bloom header row repeats on every page")
        layout_report(doc, f"SET {letter}")


def para_height(para, width):
    """Height of one paragraph in twips: wrapped line count x leading."""
    text = "".join(t.text or "" for t in para.iter(qn("w:t")))
    size = 22
    run = para.find(qn("w:r"))
    if run is not None:
        rPr = run.find(qn("w:rPr"))
        sz = rPr.find(qn("w:sz")) if rPr is not None else None
        if sz is not None:
            size = int(sz.get(qn("w:val")))
    pPr = para.find(qn("w:pPr"))
    spacing = pPr.find(qn("w:spacing")) if pPr is not None else None
    exact = spacing is not None and spacing.get(qn("w:lineRule")) == "exact"
    line_h = int(spacing.get(qn("w:line"))) if exact else int(size * 1.25 * 10)
    indent = pPr.find(qn("w:ind")) if pPr is not None else None
    avail = width - (int(indent.get(qn("w:left"))) if indent is not None else 0)
    per_line = max(10, avail // max(20, size * 5))
    nlines = max(1, -(-len(text) // per_line)) if text else 1
    return nlines * line_h


def layout_report(doc, label):
    """Estimate the printed height of a question table from the XML.
    Cell margins of the TableGrid style are 0 top/bottom, so a row is as tall
    as its tallest cell.  Estimate only - there is no renderer here."""
    usable_h = 15840 - 446 - 432            # letter page minus top/bottom margin
    for idx, name in ((1, "offline"), (4, "online")):
        tbl = doc.tables[idx]._tbl
        grid = tbl.find(qn("w:tblGrid"))
        widths = [int(g.get(qn("w:w"))) for g in grid.findall(qn("w:gridCol"))]
        total = 0
        for tr in tbl.findall(qn("w:tr")):
            row = 0
            for tc, w in zip(tr.findall(qn("w:tc")), widths):
                row = max(row, sum(para_height(p, w - 216)
                                   for p in tc.findall(qn("w:p"))))
            total += row
        print(f"     {label} {name} question table ~{total / 1440:.1f} in "
              f"(one page of body text is ~{usable_h / 1440:.1f} in)")


def page_break_present(doc):
    body = doc.element.body
    tables = body.findall(qn("w:tbl"))
    node = tables[2].getnext()
    while node is not None and node is not tables[3]:
        for br in node.iter(qn("w:br")):
            if br.get(qn("w:type")) == "page":
                return True
        node = node.getnext()
    return False


def verify_set_differences():
    print("\n== the three sets differ only by MCQ reshuffling ==")
    by_opts = {frozenset(m["opts"]): m for m in MCQS}
    orders, stems, opt_order = {}, {}, {}
    for letter in ("A", "B", "C"):
        path = OUTDIR / f"SET {letter}_T1_PYTHON-1_TEST PAPER_SEM I_MDP.docx"
        doc = Document(str(path))
        pbs, stm, oo = [], [], {}
        for r in doc.tables[1].rows[2:12]:
            body = r.cells[2].text
            opts = parse_options(body)
            m = by_opts[frozenset(opts.values())]
            pbs.append(m["pb"])
            oo[m["pb"]] = [opts[L] for L in "ABCDEF"]
            stem = "\n".join(l for l in body.splitlines()
                             if not re.match(r"\s*\([A-F]\)", l))
            stm.append(re.sub(r"\s+", " ", stem).strip())
        orders[letter], stems[letter], opt_order[letter] = pbs, stm, oo

    check(sorted(orders["A"]) == sorted(orders["B"]) == sorted(orders["C"]),
          "all three sets contain the same 10 MCQs")
    check(sorted(stems["A"]) == sorted(stems["B"]) == sorted(stems["C"]),
          "all three sets print the same question stems (word for word)")
    check(orders["A"] != orders["B"] and orders["B"] != orders["C"]
          and orders["A"] != orders["C"], "MCQ order differs in every set")
    for x, y in (("A", "B"), ("B", "C"), ("A", "C")):
        diff = sum(opt_order[x][pb] != opt_order[y][pb] for pb in opt_order[x])
        check(diff >= 8, f"option order reshuffled for {diff}/10 MCQs "
                         f"between SET {x} and SET {y}")
    # Q-2 / Q-3 text must be byte identical in the three sets
    q23 = {}
    for letter in ("A", "B", "C"):
        path = OUTDIR / f"SET {letter}_T1_PYTHON-1_TEST PAPER_SEM I_MDP.docx"
        doc = Document(str(path))
        q23[letter] = ([r.cells[2].text for r in doc.tables[1].rows[12:]]
                       + [r.cells[2].text for r in doc.tables[4].rows[1:]])
    check(q23["A"] == q23["B"] == q23["C"],
          "Q-2 and Q-3 are identical in all three sets (only MCQs reshuffled)")


def main():
    verify_answers()
    verify_documents()
    verify_set_differences()
    print()
    if FAIL:
        print(f"{len(FAIL)} CHECK(S) FAILED")
        for f in FAIL:
            print("   -", f)
        sys.exit(1)
    print("ALL CHECKS PASSED")


if __name__ == "__main__":
    main()
