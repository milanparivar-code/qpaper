#!/usr/bin/env python3
"""
Verification harness for the T1 Python-I question paper (guidelines
PYTHON-1_T1_QUESTION_PAPER_GENERATION_GUIDELINES.pdf).

1. RUNTIME   every output MCQ is executed and the marked option compared with
             the real output; the descriptive questions' stated examples are
             executed; a reference solution of the online question is run to
             prove the specification is well posed.
2. RULES     §3 unit blueprint, §4/§7 set rules, §5 PB split, §6 six options
             with e) Error and f) None of the above, §9 one integrated online
             question, §10 template controls, §11 confidentiality.
"""

import contextlib
import io
import re
import sys
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

sys.path.insert(0, str(Path(__file__).resolve().parent))
from build_t1_qp import (MCQS, OFFLINE_UNIT_TARGET, ONLINE_NAME, OUTDIR,  # noqa: E402
                         Q2, Q3, SET_ORDER, T1_UNIT_BLUEPRINT, offline_name)

FAIL = []

RUNTIME = {54: "6", 107: "4", 165: "12", 167: "2 5", 176: "1 5", 160: "2 5 7"}
EXPRESSIONS = {64: [512, 64, 512]}
ERRORS = {86: "ZeroDivisionError"}
MANUAL = {8, 21}          # concept question and paper algorithm, traced by hand


def check(cond, msg):
    print(("  OK   " if cond else "  FAIL ") + msg)
    if not cond:
        FAIL.append(msg)


def run(src):
    buf = io.StringIO()
    try:
        with contextlib.redirect_stdout(buf):
            exec(src, {})
        return buf.getvalue(), None
    except Exception as exc:                       # noqa: BLE001
        return buf.getvalue(), type(exc).__name__


def numbers(s):
    return [int(n) for n in re.findall(r"-?\d+", s)]


# --------------------------------------------------------------------------- #
def verify_answers():
    print("== 1. every MCQ answer re-derived from the code itself ==")
    covered = set(RUNTIME) | set(EXPRESSIONS) | set(ERRORS) | MANUAL
    pbs = [m["pb"] for m in MCQS]
    check(set(pbs) == covered and len(pbs) == len(set(pbs)),
          f"all {len(pbs)} MCQs covered by exactly one proof method")
    for m in MCQS:
        pb, correct = m["pb"], m["opts"][m["ans"]]
        check(len(m["opts"]) == 6 and len(set(m["opts"])) == 6,
              f"PB{pb}: exactly 6 distinct options (§6)")
        if m["is_output"]:
            check(m["opts"][4] == "Error" and m["opts"][5] == "None of the above",
                  f"PB{pb}: e) Error and f) None of the above (§6)")
            check(m["ans"] < 5 or pb == 86,
                  f"PB{pb}: correct answer is a real option, not f)")
        if pb in MANUAL:
            print(f"  ---- PB{pb}: hand traced, marked {correct!r}")
            continue
        if pb in EXPRESSIONS:
            vals = [eval(line, {}) for line in m["code"]]     # noqa: S307
            check(vals == EXPRESSIONS[pb] == [int(x) for x in correct.split(",")],
                  f"PB{pb}: expressions = {vals}, marked {correct!r}")
            continue
        out, err = run("\n".join(m["code"]))
        if pb in ERRORS:
            check(err == ERRORS[pb] and correct == "Error",
                  f"PB{pb}: raises {err}, marked {correct!r}")
            continue
        check(out.strip().rstrip(",") == RUNTIME[pb] and
              numbers(out) == numbers(correct),
              f"PB{pb}: printed {out.strip()!r}, marked {correct!r}")

    print("\n== descriptive / online questions are well posed ==")
    out, err = run("total = 7 / 2\ncount = int(total)\n"
                   "result = count ** 2 + 9 % 4\n"
                   "print(result, type(result), type(total))")
    check(out.strip() == "10 <class 'int'> <class 'float'>",
          f"Q-2 1) cell really prints {out.strip()!r}")
    n = 9875
    while n >= 10:
        s = 0
        while n:
            s += n % 10
            n //= 10
        n = s
    check(n == 2, "Q-2 2) digital root of 9875 is 2, as the paper states")

    def bill(prev, cur, kind):
        units = cur - prev
        if kind == 1:
            if units <= 100:
                energy = units * 3
            elif units <= 200:
                energy = 300 + (units - 100) * 5
            else:
                energy = 300 + 500 + (units - 200) * 7
            service = 50
        else:
            energy, service = units * 9, 150
        return units, energy, service, energy + service
    cases = [(1000, 1450, 1), (1000, 1450, 2), (500, 650, 1), (500, 500, 1)]
    for prev, cur, kind in cases:
        u, e, s, t = bill(prev, cur, kind)
        assert t == e + s
    check(bill(1000, 1450, 1) == (450, 2550, 50, 2600),
          "Q-3 reference solution: domestic 450 units -> bill 2600 (HIGH USAGE)")
    check(bill(1000, 1450, 2) == (450, 4050, 150, 4200),
          "Q-3 reference solution: commercial 450 units -> bill 4200")
    check(bill(500, 650, 1) == (150, 550, 50, 600),
          "Q-3 reference solution: 150 units straddles the 100-unit slab")


# --------------------------------------------------------------------------- #
def set_box_letters(doc):
    return ["".join(t.text or "" for t in tb.iter(qn("w:t"))).strip()
            for tb in doc.element.body.iter(qn("w:txbxContent"))
            if "".join(t.text or "" for t in tb.iter(qn("w:t"))).strip()
            .startswith("Set")]


def parse_options(cell_text):
    return dict(re.findall(r"\(([a-f])\)\s*([^\t\n]+)", cell_text))


# the two template question tables have different geometries - both must be
# reproduced exactly (§10)
OFFLINE_GEOM = ({"624", "677"}, ["642", "497", "8396", "810", "841"])
ONLINE_GEOM = ({"260", "4832"}, ["620", "625", "8110", "900", "1348"])


def template_controls(doc, qtable, label, geom):
    exp_h, exp_w = geom
    floating = [i for i, t in enumerate(doc.tables)
                if t._tbl.find(qn("w:tblPr")).find(qn("w:tblpPr")) is not None]
    check(not floating, f"{label}: no floating question table {floating}")
    heights = {tr.find(qn("w:trPr")).find(qn("w:trHeight")).get(qn("w:val"))
               for tr in qtable._tbl.findall(qn("w:tr"))
               if tr.find(qn("w:trPr")) is not None
               and tr.find(qn("w:trPr")).find(qn("w:trHeight")) is not None}
    check(heights == exp_h,
          f"{label}: template row heights preserved {sorted(exp_h)}, "
          f"got {sorted(heights)}")
    widths = [g.get(qn("w:w"))
              for g in qtable._tbl.find(qn("w:tblGrid")).findall(qn("w:gridCol"))]
    check(widths == exp_w,
          f"{label}: template column widths preserved {widths}")
    check(len(doc.tables) == 3 and len(doc.tables[2].rows) == 3,
          f"{label}: exactly one page - header, questions, Bloom table (§10)")
    check(doc.tables[0].rows[0].cells[0].text.strip() == "Enrollment No:",
          f"{label}: enrollment-number row untouched (§10)")
    check(not doc.element.body.findall(".//" + qn("w:br")),
          f"{label}: no stray page break inside the single-page paper")
    check(len(doc.element.body.findall(qn("w:sectPr"))) == 1,
          f"{label}: one section, template page setup intact")


def student_text(doc):
    return " ".join(c.text for t in doc.tables for r in t.rows for c in r.cells)


def no_faculty_leak(doc, label):
    student = student_text(doc)
    for leak in ("Practice Book", "PB Sr", "blueprint", "Answer", "faculty",
                 "audit"):
        check(leak not in student, f"{label}: no {leak!r} in the student copy (§11)")
    check(not re.search(r"\bOR\b", student), f"{label}: no OR-type options")


def verify_offline(letter, audit_txt):
    path = OUTDIR / offline_name(letter)
    doc = Document(str(path))
    label = f"OFFLINE SET {letter}"
    print(f"-- {path.name}")
    check(set_box_letters(doc) == [f"Set: {letter}"] * 2,
          f"both 'Set' boxes read 'Set: {letter}'")

    head = " ".join(c.text for r in doc.tables[0].rows for c in r.cells)
    for probe in ("TEST 1 ( CO1 ) (OFFLINE)", "MAX MARKS: 16 marks",
                  "2:15 to 3:30 pm", "29-Sep-2026", "1.25", "017012194",
                  "017212194", "CE/IT/CSD",
                  "Computer Programming using Python-I"):
        check(probe in head, f"{label}: header contains {probe!r}")
    check("2/3" not in head and "CO2" not in head,
          f"{label}: no leftover 2/3 or CO2 in the heading")
    check("4:15 to 5:30" not in student_text(doc) and "Q-3" not in student_text(doc),
          f"{label}: carries no online content")

    rows = doc.tables[1].rows
    block = rows[1]
    check(block.cells[0].text.strip() == "Q-1"
          and "1 Mark each" in block.cells[2].text
          and block.cells[3].text.strip() == "(10)",
          f"{label}: Q-1 block row = compulsory MCQ, 1 mark each, (10)")
    mcq_rows, desc = [], 0
    for r in rows[2:]:
        sub, marks, body = (r.cells[1].text.strip(),
                            r.cells[3].text.strip(), r.cells[2].text)
        if re.search(r"\([a-f]\)", body):
            mcq_rows.append((sub, body, marks))
        else:
            desc += int(marks)
    check(len(mcq_rows) == 10 and sum(int(m) for _, _, m in mcq_rows) == 10,
          f"{label}: {len(mcq_rows)} MCQs worth "
          f"{sum(int(m) for _, _, m in mcq_rows)} marks from the PB (§5)")
    check(desc == 6, f"{label}: outside-PB descriptive = {desc} marks (§5)")
    check(all(len(parse_options(b)) == 6 for _, b, _ in mcq_rows),
          f"{label}: every MCQ prints six options (§6)")
    check([by_opts[frozenset(parse_options(b).values())]["pb"]
           for _, b, _ in mcq_rows] == [MCQS[i]["pb"] for i in SET_ORDER[letter]],
          f"{label}: MCQ sequence matches SET_ORDER")
    for sub, body, _ in mcq_rows:
        m = by_opts[frozenset(parse_options(body).values())]
        if m["is_output"]:
            o = parse_options(body)
            check(o["e"] == "Error" and o["f"] == "None of the above",
                  f"{label}: {sub} (PB{m['pb']}) prints e) Error and "
                  f"f) None of the above")

    nested = next(r for r in rows if "elif a > 5:" in r.cells[2].text)
    code_paras, text_paras = [], []
    for para in nested.cells[2].paragraphs:
        pPr = para._p.find(qn("w:pPr"))
        ind = pPr.find(qn("w:ind")) if pPr is not None else None
        fonts = {r._r.find(qn("w:rPr")).find(qn("w:rFonts")).get(qn("w:ascii"))
                 for r in para.runs if r._r.find(qn("w:rPr")) is not None
                 and r._r.find(qn("w:rPr")).find(qn("w:rFonts")) is not None}
        (code_paras if ind is not None else text_paras).append(
            (int(ind.get(qn("w:left"))) if ind is not None else None, fonts))
    check(code_paras and all(f == {"Courier New"} for _, f in code_paras),
          f"{label}: all {len(code_paras)} code lines are monospaced (§10)")
    check(all(f == {"Times New Roman"} for _, f in text_paras),
          f"{label}: stem and options keep the template font")
    levels = sorted({i for i, _ in code_paras})
    check(len(levels) >= 3,
          f"{label}: code indentation preserved, levels {levels}")

    block_txt = audit_txt.split(f"--- ANSWER KEY SET {letter}")[1] \
                         .split("--- ANSWER KEY")[0]
    key = dict(re.findall(r"(Q-1 \d+\))\s+(\([a-f]\))", block_txt))
    for sub, body, _ in mcq_rows:
        o = parse_options(body)
        m = by_opts[frozenset(o.values())]
        right = [L for L, v in o.items() if v == m["opts"][m["ans"]]]
        check(len(right) == 1 and key.get(f"Q-1 {sub}") == f"({right[0]})",
              f"{label}: key Q-1 {sub} = ({right[0]}) is the printed position of "
              f"the correct option (PB{m['pb']})")
    mains = {r.cells[0].text.strip() for r in rows[1:] if r.cells[0].text.strip()}
    check(mains == {"Q-1", "Q-2"}, f"{label}: main questions = {sorted(mains)}")
    no_faculty_leak(doc, label)
    template_controls(doc, doc.tables[1], label, OFFLINE_GEOM)
    return [parse_options(r.cells[2].text) for r in rows[2:12]], \
           [r.cells[2].text for r in rows[12:]]


def verify_online():
    path = OUTDIR / ONLINE_NAME
    doc = Document(str(path))
    label = "ONLINE (common)"
    print(f"-- {path.name}")
    check(set_box_letters(doc) == ["Set: 1"] * 2,
          "both 'Set' boxes read 'Set: 1' (one common online paper)")
    head = " ".join(c.text for r in doc.tables[0].rows for c in r.cells)
    for probe in ("TEST 1 ( CO1 ) (ONLINE)", "MAX MARKS: 09 marks",
                  "4:15 to 5:30 pm", "29-Sep-2026", "1.25", "017012194",
                  "017212194"):
        check(probe in head, f"{label}: header contains {probe!r}")
    check("2:15 to 3:30" not in student_text(doc),
          f"{label}: carries no offline content")

    on = doc.tables[1].rows[1:]
    check(len(on) == 1 and on[0].cells[1].text.strip() == "(A)"
          and on[0].cells[3].text.strip() == "(09)"
          and on[0].cells[4].text.strip() == "C",
          f"{label}: ONE integrated Q-3(A), marks (09), Bloom C (§9)")
    body = on[0].cells[2].text
    check("algorithm" in body.lower() and "flowchart" in body.lower(),
          f"{label}: requires an algorithm and a flowchart (Unit-1 link, §9)")
    for probe in ("loop", "again", "how many", "Rs. 2000"):
        check(probe in body, f"{label}: requires {probe!r}")
    spec = body.split("Use of functions")[0]
    for banned in ("def ", "list(", "tuple(", "dict(", "set(", "import ",
                   "open(", "[", "lambda"):
        check(banned not in spec, f"{label}: avoids {banned!r} (§9)")
    check("ternary" not in body.lower(), f"{label}: avoids the ternary form (§9)")
    mains = {r.cells[0].text.strip() for r in on if r.cells[0].text.strip()}
    check(mains == {"Q-3"}, f"{label}: main questions = {sorted(mains)}")
    no_faculty_leak(doc, label)
    template_controls(doc, doc.tables[1], label, ONLINE_GEOM)


def verify_documents():
    global by_opts
    print("\n== 2. the generated student papers ==")
    by_opts = {frozenset(m["opts"]): m for m in MCQS}
    check(len(by_opts) == len(MCQS), "option sets identify each MCQ uniquely")
    audit_txt = (OUTDIR
                 / "AUDIT SHEET AND ANSWER KEY_T1_PYTHON-I_MDP_faculty only.txt"
                 ).read_text()

    opts, tail = {}, {}
    for letter in ("A", "B", "C"):
        opts[letter], tail[letter] = verify_offline(letter, audit_txt)
    verify_online()

    print("\n== sets differ only in MCQ sequence (§4, §7) ==")
    from collections import Counter
    for letter in ("A", "B", "C"):
        order = [tuple(sorted(o.values())) for o in opts[letter]]
        for other in ("A", "B", "C"):
            if other <= letter:
                continue
            o2 = [tuple(sorted(o.values())) for o in opts[other]]
            check(sorted(order) == sorted(o2),
                  f"SET {letter} and SET {other} contain the same ten questions")
            check(order != o2,
                  f"SET {letter} and SET {other} use a different MCQ sequence")
        check(tail["A"] == tail[letter],
              f"Q-2 is byte-identical in SET A and SET {letter}")
        letters = [L for o in opts[letter]
                   for L, v in o.items()
                   if v == by_opts[frozenset(o.values())]
                   ["opts"][by_opts[frozenset(o.values())]["ans"]]]
        c = Counter(letters)
        check(max(c.values()) <= 3 and len(set(letters)) >= 4,
              f"SET {letter} answer-key spread {dict(sorted(c.items()))}")


def verify_blueprint():
    print("\n== 3. the unit blueprint (§3) ==")
    unit_offline = {1: 0, 2: 0, 3: 0}
    for m in MCQS:
        unit_offline[m["unit"]] += 1
    for q in Q2:
        unit_offline[q["unit"]] += q["marks"]
    on_u2, on_u3 = Q3[0]["unit_online"]
    total = {1: unit_offline[1], 2: unit_offline[2] + on_u2,
             3: unit_offline[3] + on_u3}
    check(unit_offline == OFFLINE_UNIT_TARGET,
          f"offline unit split {unit_offline[1]}/{unit_offline[2]}/"
          f"{unit_offline[3]} = target 2/6/8 (§3)")
    check(total == T1_UNIT_BLUEPRINT,
          f"T1 unit total {total[1]}/{total[2]}/{total[3]} = blueprint 2/10/13 (§3)")
    names = sorted(f.name for f in OUTDIR.glob("*.docx"))
    check(len(names) == 4 and sum(n.startswith("SET") for n in names) == 3
          and sum(n.startswith("ONLINE") for n in names) == 1,
          f"deliverable = 3 offline sets + 1 common online paper: {names}")


def main():
    verify_answers()
    verify_documents()
    verify_blueprint()
    print()
    if FAIL:
        print(f"{len(FAIL)} CHECK(S) FAILED")
        for f in FAIL:
            print("   -", f)
        sys.exit(1)
    print("ALL CHECKS PASSED")


if __name__ == "__main__":
    main()
