#!/usr/bin/env python3
"""
Verification harness for the T1 Python-I question paper (guidelines
PYTHON-1_T1_QUESTION_PAPER_GENERATION_GUIDELINES.pdf).

1. RUNTIME   every MCQ is executed and the marked option compared with the
             real output; the descriptive questions' worked examples and a
             reference solution of the online specification are executed to
             prove both are well posed.
2. FORMAT    the generated header is compared run by run and paragraph by
             paragraph with the untouched template, so a collapsed or
             re-flowed template field cannot pass.
3. RULES     §3 blueprint, §4/§7 sets, §5 PB split, §6 six options with
             e) Error and f) None of the above, §9 one integrated online
             question, §10 template geometry and marks format, §11 audit.
"""

import contextlib
import io
import re
import sys
from collections import Counter
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

sys.path.insert(0, str(Path(__file__).resolve().parent))
from build_t1_qp import (BRANCHES, MCQS, OFFLINE_UNIT_TARGET, ONLINE_NAME,  # noqa: E402
                         OUTDIR, Q2, Q3, SET_ORDER, SUBJECT, SUBJECT_CODES,
                         T1_UNIT_BLUEPRINT, mark, offline_name)

FAIL = []
TEMPLATE = (Path(__file__).resolve().parents[1] / "QP REQUIRED DOCUMENTS"
            / "04 QP_FORMAT/QP_FORMAT/01 FORMAT_PYTHON-1"
            / "SET A_T1 to T3_PYTHON-1_TEST PAPER_SEM I_FACUTLY SHORT NAME_FORMAT.docx")

# every MCQ answer is re-derived by executing its code; the value below is what
# the code really prints (trailing separators normalised away)
RUNTIME = {66: "11.0", 74: "46.0", 84: "3.0", 68: "27.2", 166: "77",
           161: "81", 158: "26,27,28,29,", 164: "28", 170: "1 2 4",
           160: "2 5 7,", 71: "False", 76: "9", 177: "11 5", 114: "5\n11",
           165: "12"}


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


def norm(s):
    return re.sub(r"[\s,]+$", "", s.strip())


def marks_value(text):
    return float(re.sub(r"[][ ]", "", text))


# --------------------------------------------------------------------------- #
def verify_answers():
    print("== 1. every MCQ answer re-derived by executing the code ==")
    pbs = [m["pb"] for m in MCQS]
    check(set(pbs) == set(RUNTIME) and len(pbs) == len(set(pbs)),
          f"all {len(pbs)} MCQs are execution-verified, none copied")
    for m in MCQS:
        pb, correct = m["pb"], m["opts"][m["ans"]]
        check(len(m["opts"]) == 6 and len(set(m["opts"])) == 6,
              f"PB{pb}: exactly 6 distinct options (§6)")
        check(m["opts"][4] == "Error" and m["opts"][5] == "None of the above",
              f"PB{pb}: e) Error and f) None of the above (§6)")
        check(m["ans"] < 5, f"PB{pb}: the answer is a real option, not f)")
        out, err = run("\n".join(m["code"]))
        check(err is None, f"PB{pb}: the snippet runs without raising ({err})")
        check(norm(out) == norm(RUNTIME[pb]),
              f"PB{pb}: really prints {out.strip()!r}, expected {RUNTIME[pb]!r}")
        check(norm(out) == norm(correct) or
              [float(x) for x in re.findall(r"-?\d+\.?\d*", norm(out))] ==
              [float(x) for x in re.findall(r"-?\d+\.?\d*", correct)],
              f"PB{pb}: printed value equals the marked option {correct!r}")

    print("\n== the descriptive and online questions are well posed ==")
    src = "\n".join(line for kind, block in Q2[1]["body"] if kind == "code"
                    for line in block)
    out, err = run(src)
    lines = out.splitlines()
    check(err is None and lines and lines[0] == "17 <class 'int'> <class 'float'>",
          f"Q-2 2) cell really gives {lines[0] if lines else None!r}")

    def rows(n):
        total, out = 0, []
        for i in range(1, n + 1):
            row = [j for j in range(1, i + 1) if i % j == 0]
            total += len(row)
            out.append(f"Row {i} : " + " ".join(str(j) for j in row))
        return out, total
    exp, total = rows(4)
    check(exp == ["Row 1 : 1", "Row 2 : 1 2", "Row 3 : 1 3", "Row 4 : 1 2 4"],
          f"Q-2 3) worked example (n = 4) really gives {exp}")
    check(total == 8, "Q-2 3) example is consistent (8 divisors in total)")

    def bill(prev, cur, kind):
        units = cur - prev
        if kind == 1:
            energy = (units * 3 if units <= 100 else
                      300 + (units - 100) * 5 if units <= 200 else
                      800 + (units - 200) * 7)
            fixed = 50
        else:
            energy = units * 9 if units <= 200 else 1800 + (units - 200) * 12
            fixed = 150
        sur = energy * 5 // 100 if units > 500 else 0
        tot = energy + fixed + sur
        cat = "HIGH" if tot > 5000 else "MEDIUM" if tot > 2000 else "LOW"
        return units, energy, fixed, sur, tot, cat
    check(bill(1000, 1450, 1) == (450, 2550, 50, 0, 2600, "MEDIUM"),
          "Q-3 reference: domestic 450 units -> 2600, MEDIUM")
    check(bill(1000, 1700, 1) == (700, 4300, 50, 215, 4565, "MEDIUM"),
          "Q-3 reference: 700 units cross the 500-unit surcharge line")
    check(bill(100, 900, 2) == (800, 9000, 150, 450, 9600, "HIGH"),
          "Q-3 reference: commercial 800 units -> 9600, HIGH")
    check(bill(500, 650, 1) == (150, 550, 50, 0, 600, "LOW"),
          "Q-3 reference: 150 units straddle the 100-unit slab -> LOW")
    check(bill(0, 60, 1)[5] == "LOW" and bill(0, 0, 1) == (0, 0, 50, 0, 50, "LOW"),
          "Q-3 reference: zero consumption is defined, not a division by zero")


# --------------------------------------------------------------------------- #
def paras(cell):
    return cell.findall(qn("w:p"))


def para_text(p):
    """Paragraph text without the floating 'Set:' textbox, which is anchored
    inside the CCE paragraph and is checked on its own."""
    out = []
    for t in p.iter(qn("w:t")):
        anc = t.getparent()
        while anc is not None and anc.tag != qn("w:txbxContent"):
            anc = anc.getparent()
        if anc is None:
            out.append(t.text or "")
    return "".join(out)


def verify_header(doc, tpl, offline, label):
    """The header must differ from the template ONLY in the variable runs."""
    print(f"   header layout vs template ({label})")
    # the template carries the offline header in table 0 and the online header
    # in table 3 - each split file must be compared with its own
    gt = doc.tables[0]._tbl
    tt = (tpl.tables[0] if offline else tpl.tables[3])._tbl
    grows, trows = gt.findall(qn("w:tr")), tt.findall(qn("w:tr"))
    check(len(grows) == len(trows) == 9,
          f"{label}: header still has 9 rows, got {len(grows)}")
    for ri in (0, 1, 2, 3, 8):
        gp = [para_text(p) for p in paras(grows[ri].findall(qn("w:tc"))[0])]
        tp = [para_text(p) for p in paras(trows[ri].findall(qn("w:tc"))[0])]
        if ri == 3:
            tp = [t.replace("TEST 1/2/3 ( CO1/CO2/CO3 )", "TEST 1 ( CO1 )")
                  for t in tp]
        check(gp == tp, f"{label}: header row {ri} untouched")
    for ri in (4, 5):
        gp = paras(grows[ri].findall(qn("w:tc"))[0])
        tp = paras(trows[ri].findall(qn("w:tc"))[0])
        check(len(gp) == len(tp) == 2,
              f"{label}: header row {ri} keeps the template's TWO paragraphs "
              f"(got {len(gp)})")
        check([len(p.findall(qn('w:r'))) for p in gp] ==
              [len(p.findall(qn('w:r'))) for p in tp],
              f"{label}: header row {ri} keeps the template's run count")
    g4 = [para_text(p) for p in paras(grows[4].findall(qn("w:tc"))[0])]
    check(g4[0].startswith("B.E. SEMESTER-I") and "BRANCH" not in g4[0],
          f"{label}: 'B.E. SEMESTER-I' is still on its own line")
    check(g4[1] == "BRANCH: " + BRANCHES,
          f"{label}: the branch list is on the template's BRANCH: line")
    g5 = [para_text(p) for p in paras(grows[5].findall(qn("w:tc"))[0])]
    check(g5[0] == "SUBJECT- " + SUBJECT and g5[1] == "SUBJECT CODE: " + SUBJECT_CODES,
          f"{label}: subject and subject codes on their own two lines")
    for ri, probes in ((6, ("DATE: 29-Sep-2026", "DURATION:    1.25  Hours")),
                       (7, ("TIME:  " + ("2:15 to 3:30 pm" if offline
                                          else "4:15 to 5:30 pm"),
                            "MAX MARKS: " + ("16" if offline else "09")))):
        txt = para_text(paras(grows[ri].findall(qn("w:tc"))[0])[0])
        tpl_txt = para_text(paras(trows[ri].findall(qn("w:tc"))[0])[0])
        for pr in probes:
            check(pr in txt, f"{label}: row {ri} contains {pr!r}")
        # the label that follows must start in the template's own column
        key = "DURATION" if ri == 6 else "MAX"
        check(txt.index(key) == tpl_txt.index(key),
              f"{label}: row {ri} keeps {key} in the template's column "
              f"({txt.index(key)} == {tpl_txt.index(key)})")


def is_mcq(cell_text):
    """An MCQ row is one whose option block starts on its own '(a) ' line.

    Matching a bare '(a)' anywhere would also hit code such as print(c).
    """
    return "\n(a) " in cell_text


def parse_options(cell_text):
    """Read the option block only - code lines may contain '(c)' too."""
    tail = cell_text[cell_text.rindex("\n(a) ") + 1:]
    return dict(re.findall(r"\(([a-f])\)\s*([^\t\n]+)", tail))


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
    check(widths == exp_w, f"{label}: template column widths preserved {widths}")
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
    check(not re.search(r"\bOR\b", student), f"{label}: no OR-type options (§4)")


OFFLINE_GEOM = ({"624", "677"}, ["642", "497", "8396", "810", "841"])
ONLINE_GEOM = ({"260", "4832"}, ["620", "625", "8110", "900", "1348"])


def verify_offline(letter, tpl, by_opts, audit_txt):
    path = OUTDIR / offline_name(letter)
    doc = Document(str(path))
    label = f"OFFLINE SET {letter}"
    print(f"-- {path.name}")
    check([t for t in ("".join(x.text or "" for x in tb.iter(qn("w:t"))).strip()
                       for tb in doc.element.body.iter(qn("w:txbxContent")))
           if t.startswith("Set")] == [f"Set: {letter}"] * 2,
          f"both 'Set' boxes read 'Set: {letter}'")
    verify_header(doc, tpl, True, label)

    rows = doc.tables[1].rows
    blocks = [r for r in rows[1:] if r.cells[0].text.strip() == "Q-1"]
    check([r.cells[1].text.strip() for r in blocks] == ["A)", "B)"],
          f"{label}: Q-1 has an A) and a B) block, in that order")
    check("0.5 Marks each" in blocks[0].cells[2].text
          and blocks[0].cells[3].text.strip() == mark(5),
          f"{label}: Q-1(A) is the 0.5-mark block, total {mark(5)}")
    check("1 Mark each" in blocks[1].cells[2].text
          and blocks[1].cells[3].text.strip() == mark(5),
          f"{label}: Q-1(B) is the 1-mark block, total {mark(5)}")

    mcq_rows = [(r.cells[1].text.strip(), r.cells[2].text,
                 r.cells[3].text.strip())
                for r in rows[2:] if is_mcq(r.cells[2].text)]
    check(len(mcq_rows) == len(MCQS),
          f"{label}: {len(mcq_rows)} MCQs printed (expected {len(MCQS)})")
    n_a = sum(1 for m in MCQS if m["group"] == "A")
    check(all(m == mark(0.5) for _, _, m in mcq_rows[:n_a])
          and all(m == mark(1) for _, _, m in mcq_rows[n_a:]),
          f"{label}: first {n_a} MCQs marked {mark(0.5)}, the rest {mark(1)}")
    check(sum(marks_value(m) for _, _, m in mcq_rows) == 10,
          f"{label}: MCQ marks total "
          f"{sum(marks_value(m) for _, _, m in mcq_rows):g} = the PB quota (§5)")

    # a Q-2 row is neither an MCQ row nor one of the Q-1 block rows
    q2_rows = [r for r in rows[2:]
               if not is_mcq(r.cells[2].text)
               and r.cells[1].text.strip() not in ("A)", "B)")]
    check(len(q2_rows) == len(Q2), f"{label}: Q-2 has {len(Q2)} sub-questions")
    for r, q in zip(q2_rows, Q2):
        check(marks_value(r.cells[3].text) == q["marks"]
              and r.cells[3].text.strip() == mark(q["marks"]),
              f"{label}: Q-2 {r.cells[1].text.strip()} marks "
              f"{r.cells[3].text.strip()} (Unit {q['unit']})")
    check(sum(q["marks"] for q in Q2) == 6,
          f"{label}: Q-2 carries 6 marks, all outside the PB (§5)")
    # FY instruction Case 1 / guidelines §6: mixed marks need >= 7 per block
    check(len(blocks[0:1]) and len(mcq_rows[:n_a]) == 10 and len(mcq_rows[n_a:]) == 5,
          f"{label}: 10 x 0.5 in Q-1(A) and 5 x 1 in Q-1(B), as requested")
    print("  ---- §6 NOTE: Q-1(B) has 5 questions, below the FY instruction's")
    print("       minimum of 7 per block for mixed 0.5/1-mark MCQs; written")
    print("       HOD clearance is required for this pattern.")
    check(all(len(parse_options(b)) == 6 for _, b, _ in mcq_rows),
          f"{label}: every MCQ prints six options (§6)")
    check([by_opts[frozenset(parse_options(b).values())]["pb"]
           for _, b, _ in mcq_rows]
          == [MCQS[i]["pb"]
              for i in SET_ORDER[letter]["A"] + SET_ORDER[letter]["B"]],
          f"{label}: MCQ sequence matches SET_ORDER (A block then B block)")
    for sub, body, _ in mcq_rows:
        o = parse_options(body)
        check(o["e"] == "Error" and o["f"] == "None of the above",
              f"{label}: {sub} prints e) Error and f) None of the above")

    nested = next(r for r in rows if "elif count==5:" in r.cells[2].text)
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
    check(len(levels) >= 4,
          f"{label}: nested indentation preserved, levels {levels}")
    rule = {p._p.find(qn("w:pPr")).find(qn("w:spacing")).get(qn("w:lineRule"))
            for p in nested.cells[2].paragraphs
            if p._p.find(qn("w:pPr")) is not None
            and p._p.find(qn("w:pPr")).find(qn("w:spacing")) is not None}
    check("exact" not in rule,
          f"{label}: no 'exact' leading that could clip code {sorted(rule)}")
    stops = {t.get(qn("w:pos"))
             for p in rows[2].cells[2].paragraphs
             for tabs in [p._p.find(qn("w:pPr")).find(qn("w:tabs"))
                          if p._p.find(qn("w:pPr")) is not None else None]
             if tabs is not None for t in tabs.findall(qn("w:tab"))}
    check(stops == {"2800", "5600"},
          f"{label}: options sit on explicit tab stops {sorted(stops)} (§13)")

    block_txt = audit_txt.split(f"--- ANSWER KEY SET {letter}")[1] \
                         .split("--- ANSWER KEY")[0]
    key = dict(re.findall(r"(Q-1\([AB]\) \d+\))\s+(\([a-f]\))", block_txt))
    seq = []
    for i, (sub, body, _) in enumerate(mcq_rows):
        o = parse_options(body)
        m = by_opts[frozenset(o.values())]
        right = [L for L, v in o.items() if v == m["opts"][m["ans"]]]
        grp = "A" if i < n_a else "B"
        check(len(right) == 1 and key.get(f"Q-1({grp}) {sub}") == f"({right[0]})",
              f"{label}: key Q-1({grp}) {sub} = ({right[0]}) is the printed "
              f"position of the correct option (PB{m['pb']})")
        seq.append(right[0])
    for grp, part in (("A", seq[:n_a]), ("B", seq[n_a:])):
        check(all(x != y for x, y in zip(part, part[1:])),
              f"{label}: Q-1({grp}) has no two consecutive equal answers {part}")
    mains = {r.cells[0].text.strip() for r in rows[1:] if r.cells[0].text.strip()}
    check(mains == {"Q-1", "Q-2"}, f"{label}: main questions = {sorted(mains)}")
    no_faculty_leak(doc, label)
    template_controls(doc, doc.tables[1], label, OFFLINE_GEOM)
    return ([parse_options(b) for _, b, _ in mcq_rows],
            [r.cells[2].text for r in q2_rows], n_a)


def verify_online(tpl):
    path = OUTDIR / ONLINE_NAME
    doc = Document(str(path))
    label = "ONLINE (common)"
    print(f"-- {path.name}")
    check([t for t in ("".join(x.text or "" for x in tb.iter(qn("w:t"))).strip()
                       for tb in doc.element.body.iter(qn("w:txbxContent")))
           if t.startswith("Set")] == ["Set: 1"] * 2,
          "both 'Set' boxes read 'Set: 1' (one common online paper)")
    verify_header(doc, tpl, False, label)
    on = doc.tables[1].rows[1:]
    check(len(on) == 1 and on[0].cells[1].text.strip() == "(A)"
          and on[0].cells[3].text.strip() == mark(9)
          and on[0].cells[4].text.strip() == "C",
          f"{label}: ONE integrated Q-3(A), marks {mark(9)}, Bloom C (§9)")
    body = on[0].cells[2].text
    check("algorithm" in body.lower() and "flowchart" in body.lower(),
          f"{label}: requires an algorithm and a flowchart (Unit-1 link, §9)")
    for probe in ("eight consumers", "INVALID READING", "INVALID TYPE",
                  "surcharge", "HIGH", "MEDIUM", "LOW", "highest total bill",
                  "how many" if False else "number of consumers"):
        check(probe in body, f"{label}: requires {probe!r}")
    spec = body.split("Use of functions")[0]
    for banned in ("def ", "list(", "tuple(", "dict(", "set(", "import ",
                   "open(", "[", "lambda"):
        check(banned not in spec, f"{label}: avoids {banned!r} (§9)")
    check("ternary" in body.lower(), f"{label}: bars the ternary form (§9)")
    mains = {r.cells[0].text.strip() for r in on if r.cells[0].text.strip()}
    check(mains == {"Q-3"}, f"{label}: main questions = {sorted(mains)}")
    no_faculty_leak(doc, label)
    template_controls(doc, doc.tables[1], label, ONLINE_GEOM)


def verify_documents():
    print("\n== 2. the generated student papers ==")
    tpl = Document(str(TEMPLATE))
    by_opts = {frozenset(m["opts"]): m for m in MCQS}
    check(len(by_opts) == len(MCQS), "option sets identify each MCQ uniquely")
    audit_txt = (OUTDIR
                 / "AUDIT SHEET AND ANSWER KEY_T1_PYTHON-I_MDP_faculty only.txt"
                 ).read_text()
    opts, tail = {}, {}
    for letter in ("A", "B", "C"):
        opts[letter], tail[letter], _ = verify_offline(letter, tpl, by_opts,
                                                     audit_txt)
    verify_online(tpl)

    print("\n== sets differ only in MCQ sequence (§4, §7) ==")
    for letter in ("A", "B", "C"):
        order = [tuple(sorted(o.values())) for o in opts[letter]]
        for other in ("A", "B", "C"):
            if other <= letter:
                continue
            o2 = [tuple(sorted(o.values())) for o in opts[other]]
            check(sorted(order) == sorted(o2),
                  f"SET {letter} and SET {other} contain the same questions")
            check(order != o2,
                  f"SET {letter} and SET {other} use a different MCQ sequence")
        check(tail["A"] == tail[letter],
              f"Q-2 is byte-identical in SET A and SET {letter}")
        seq = [L for o in opts[letter] for L, v in o.items()
               if v == by_opts[frozenset(o.values())]["opts"]
               [by_opts[frozenset(o.values())]["ans"]]]
        c = Counter(seq)
        check(max(c.values()) <= len(seq) * 0.3 and len(set(seq)) >= 4,
              f"SET {letter} answer-key spread {dict(sorted(c.items()))} over "
              f"{len(seq)} questions - no letter above 30%")


def verify_blueprint():
    print("\n== 3. blueprint, PB split and file inventory (§3, §5) ==")
    unit_offline = {1: 0, 2: 0, 3: 0}
    for m in MCQS:
        unit_offline[m["unit"]] += m["marks"]
    for q in Q2:
        unit_offline[q["unit"]] += q["marks"]
    on_u2, on_u3 = Q3[0]["unit_online"]
    total = {1: unit_offline[1], 2: unit_offline[2] + on_u2,
             3: unit_offline[3] + on_u3}
    check(sum(unit_offline.values()) == 16,
          f"offline total = {sum(unit_offline.values()):g} (§5)")
    check(unit_offline == OFFLINE_UNIT_TARGET,
          f"offline unit split {unit_offline[1]:g}/{unit_offline[2]:g}/"
          f"{unit_offline[3]:g} = target 2/6/8 (§3)")
    check(total == T1_UNIT_BLUEPRINT,
          f"T1 unit total {total[1]:g}/{total[2]:g}/{total[3]:g} = blueprint "
          f"2/10/13 (§3)")
    pb = sum(m["marks"] for m in MCQS) + sum(q["marks"] for q in Q2
                                             if q["pb"])
    outside = sum(q["marks"] for q in Q2 if not q["pb"])
    check(pb == 10 and outside == 6,
          f"offline PB = {pb:g} and outside PB = {outside:g} (§5)")
    names = sorted(f.name for f in OUTDIR.glob("*.docx"))
    check(len(names) == 4 and sum(n.startswith("SET") for n in names) == 3
          and sum(n.startswith("ONLINE") for n in names) == 1,
          f"deliverable = 3 offline sets + 1 common online paper")
    mains = 2 + len(Q3)
    check(mains <= 7, f"main questions across the paper = {mains} (max 7, §4)")


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
