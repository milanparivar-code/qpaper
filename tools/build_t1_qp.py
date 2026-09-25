#!/usr/bin/env python3
"""
Build the T1 question paper (SET A / B / C) for
  Computer Programming using Python-I, SEM-I, Batch 2026     QP Setter : MDP

Driven by, in order of authority:
  QP REQUIRED DOCUMENTS/PYTHON-1_T1_QUESTION_PAPER_GENERATION_GUIDELINES.pdf
  QP REQUIRED DOCUMENTS/03 .../FY_QUESTION PAPER SETTING INSTRUCTIONS_T1 TO T3.pdf
  QP REQUIRED DOCUMENTS/02 ... Marks Distribution ... For Faculty.pdf
  QP REQUIRED DOCUMENTS/Test-1 SCHEDULE_... .pdf
  QP REQUIRED DOCUMENTS/04 .../01 FORMAT_PYTHON-1/SET A_T1 to T3_...FORMAT.docx

Guideline points implemented here
  §1  16 offline / 09 online / 25 total, 29-Sep-2026, 2:15-3:30 and 4:15-5:30
  §2  all 15 branch subject codes in the header
  §3  offline unit blueprint 2 / 6 / 8, online 4 / 5  ->  T1 total 2 / 10 / 13
  §4  conceptual, above average, no recall / T-F / fill in the blanks, Q-1 MCQ,
      max 7 main questions, 3 sets, same questions, only MCQ order reshuffled
  §5  offline 10 marks from PB + 6 marks outside PB; online 9 marks outside PB;
      faculty-only traceability sheet kept out of the student paper
  §6  every MCQ 1 mark, exactly 6 options, output MCQs use e) Error and
      f) None of the above, every output MCQ executed before finalising
  §7  outside-PB questions test precedence, type conversion, logical operators,
      nested decisions, while/for trace, break/continue, nested loops
  §9  ONE integrated 9 mark question under Q-3(A), Bloom C, Unit-1 connection
      through the required algorithm + flowchart
  §10 built on the original template file, monospaced code with indentation,
      department marks style, no structural redesign

OPEN ITEM (§6): the guidelines' *preferred* 10 x 0.5 + 5 x 1 MCQ pattern needs
written HOD clearance because it clashes with the "at least seven questions in
Q-1(A) and Q-1(B)" rule.  Until that clearance exists this script builds the
non-conflicting pattern (Case 2: ten 1-mark MCQs, all compulsory in Q-1).
"""

import copy
import io
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = (
    ROOT
    / "QP REQUIRED DOCUMENTS/04 QP_FORMAT/QP_FORMAT/01 FORMAT_PYTHON-1"
    / "SET A_T1 to T3_PYTHON-1_TEST PAPER_SEM I_FACUTLY SHORT NAME_FORMAT.docx"
)
OUTDIR = ROOT / "QP_MDP_T1_PYTHON-I_SEM-I_2026"

# --------------------------------------------------------------------------- #
# paper identity (guidelines §1 and §2)
# --------------------------------------------------------------------------- #
BRANCHES = (
    "CE/IT/CSD/AIML/AIDS/CSE/CST/CS&IT/CEA/MA&CP/CSE(AI)/"
    "CSE(DS)/CSE(CS)/CS&BIO/BIOTECH"
)
SUBJECT = "Computer Programming using Python-I"
SUBJECT_CODES = (
    "017012194/017022194/017032194/017042194/017052194/017122194/"
    "017132194/017142194/017152194/017162194/017172194/017182194/"
    "017192194/017202194/017212194"
)
DATE = "29-Sep-2026"
OFFLINE_TIME = "2:15 to 3:30 pm"
ONLINE_TIME = "4:15 to 5:30 pm"
DURATION = "1.25  Hours"
OFFLINE_MAX = "16 marks"
ONLINE_MAX = "09 marks"

# The template's question tables are floating (w:tblpPr); a floating table is
# not split over a page boundary by Word, which would push the whole offline
# paper onto one page.  Setting this to True keeps the template exactly as
# shipped, at the cost of that pagination risk.
KEEP_TEMPLATE_FLOATING_TABLE = False

# --------------------------------------------------------------------------- #
# typography (guidelines §4: 11 or 12 pt as appropriate to the template)
# --------------------------------------------------------------------------- #
Q_FONT = "Times New Roman"
CODE_FONT = "Courier New"
Q_SIZE = 22       # 11 pt question text
CODE_SIZE = 18    #  9 pt monospaced code, indentation preserved
CODE_LINE = 216   # code lines at 0.9 line spacing (proportional, never clips)
OPT_SIZE = 20     # 10 pt options
MARK_SIZE = 20    # 10 pt marks / Bloom columns
HEADER_SIZE = 24  # 12 pt institute header block (template size)

# --------------------------------------------------------------------------- #
# Q-1 : eight 1-mark MCQs, all from the T1 Practice Book (guidelines §5, §6)
#
# Every item below is an execution / trace question whose STEM is unchanged
# from the Practice Book - §5 keeps such a question classified as PB even when
# the options are changed.  The options were rebuilt to exactly six with
# e) Error and f) None of the above (§6), their order is FIXED (§4 allows only
# the MCQ *sequence* to be reshuffled) and every answer was re-derived by
# executing the code here rather than copied from the PB key.
# Recall-only PB items (Sr 1-20) and the trivial algorithm trace (Sr 21) are
# deliberately not used: §4 forbids easy recall questions.
# --------------------------------------------------------------------------- #
MCQS = [
    dict(pb=66, unit=2, bloom="A", ans=2, is_output=True,
         stem=["What will be the output of this program?"],
         code=["print(6 + 5 - 4 * 3 / 2 % 1)"],
         opts=["11", "15", "11.0", "7", "Error", "None of the above"]),
    dict(pb=74, unit=2, bloom="A", ans=3, is_output=True,
         stem=["What will be the output of the following program on execution?"],
         code=["a=4", "b=6", "c=3", "d=2", "print(a+d**b*c/a-b)"],
         opts=["46", "52.0", "48.0", "46.0", "Error", "None of the above"]),
    dict(pb=71, unit=2, bloom="A", ans=1, is_output=True,
         stem=["What should be the output of the following python code snippet:"],
         code=["x=0.0", "y=48>0", "z=11<7",
               "print(not(float(x or y or z)))"],
         opts=["True", "False", "0.0", "No output", "Error",
               "None of the above"]),
    dict(pb=114, unit=3, bloom="N", ans=1, is_output=True,
         stem=["What will be the output of given Python code?"],
         code=["n=7", "c=0", "while(n):", "    if(n>5):", "        c=c+n-1",
               "        n=n-1", "    else:", "        break",
               "print(n)", "print(c)"],
         opts=["4 16", "5 11", "6 7", "5 10", "Error", "None of the above"]),
    dict(pb=150, unit=3, bloom="N", ans=2, is_output=True,
         stem=["What will be the output of the following snippet?"],
         code=["a = True", "b = False", "c = False",
               "if not a or b:", "    print (1)",
               "elif not a or not b and c:", "    print (2)",
               "elif not a or b or not b and a:", "    print (3)",
               "else:", "    print (4)"],
         opts=["1", "2", "3", "4", "Error", "None of the above"]),
    dict(pb=166, unit=3, bloom="N", ans=2, is_output=True,
         stem=["What is the output of the following code?"],
         code=["val = 154", "while(not(val)):", "    val**=2",
               "else:", "    val//=2", "print(val)"],
         opts=["154", "11", "77", "23716", "Error", "None of the above"]),
    dict(pb=170, unit=3, bloom="N", ans=0, is_output=True,
         stem=["What will be the output of the following program on execution?"],
         code=["for x in range(0,15):", "    if(x%3==0):", "        continue",
               "    if(x%5==0):", "        continue", "    if(x%7==0):",
               "        break", "    print(x,end=\" \")"],
         opts=["1 2 4", "0 1 2 4", "1 2 3 4", "1 2 4 7", "Error",
               "None of the above"]),
    dict(pb=177, unit=3, bloom="N", ans=1, is_output=True,
         stem=["What will be the output of the following program on execution?"],
         code=["x=0", "count=0", "for x in range(10):", "    while x<15:",
               "        if x<0:", "            pass",
               "        elif x%2==0:", "            x+=1", "            continue",
               "        elif x%3==0:", "            x+=1", "            continue",
               "        elif count==5:", "            break",
               "        count+=1", "print(x,count)"],
         opts=["10 5", "11 5", "16 5", "14 5", "Error", "None of the above"]),
]

# MCQ sequence per set - the questions and their options are identical in all
# three sets, only this order changes (guidelines §4).  Orders are chosen so
# that no two consecutive answers carry the same option letter.
SET_ORDER = {
    "A": [0, 1, 2, 4, 3, 6, 5, 7],
    "B": [6, 2, 0, 7, 1, 5, 3, 4],
    "C": [1, 4, 3, 0, 6, 7, 5, 2],
}

# --------------------------------------------------------------------------- #
# Q-2 : the 6 outside-PB marks plus the 2 Unit-1 PB marks (guidelines §3, §5, §7)
#       1) PB Sr. No. 30 (algorithm + flowchart), varied to 2 marks - Unit 1
#       2) outside PB, 3 marks - Unit 2 : precedence, type conversion, logical
#       3) outside PB, 3 marks - Unit 3 : nested loops and counters
#       giving the offline unit split 2 / 6 / 8 and PB 10 + outside PB 6
# --------------------------------------------------------------------------- #
Q2 = [
    dict(unit=1, bloom="C", marks=2, pb=30,
         body=[
             ("p", "Write an algorithm and draw a flowchart to arrange the "
                   "three input numbers x, y and z in descending order."),
         ]),
    dict(unit=2, bloom="A", marks=3, pb=None,
         body=[
             ("p", "The following statements are executed in one cell of a "
                   "Jupyter Notebook:"),
             ("code", ["p = 17", "q = 5", "r = p / q",
                       "s = int(r) ** 2 + p % q * 2", "t = s / 2",
                       "print(s, type(s), t, type(t))",
                       "print(bool(p % q) and not (s > 20))"]),
             ("p", "(i)  Write the exact output produced by both print "
                   "statements."),
             ("p", "(ii) State the order in which the operators ** , % , * and "
                   "+ are applied while evaluating s, and explain why s is of "
                   "type int while t is of type float."),
         ]),
    dict(unit=3, bloom="C", marks=3, pb=None,
         body=[
             ("p", "Write a Python program that reads a positive integer n "
                   "from the user and, for every row i from 1 to n, prints the "
                   "row number followed by all the numbers from 1 to i that "
                   "divide i exactly. After all the rows have been printed, "
                   "print the total number of divisors printed. Do not use any "
                   "built-in function other than input(), int() and print()."),
             ("p", "Example :  Input  n = 4"),
             ("code", ["Row 1 : 1", "Row 2 : 1 2", "Row 3 : 1 3",
                       "Row 4 : 1 2 4", "Total divisors printed = 8"]),
         ]),
]

# --------------------------------------------------------------------------- #
# Q-3(A) : ONE integrated 9-mark question, entirely outside the Practice Book
#          (guidelines §9).  Bloom C.  Unit-1 connection = the required
#          algorithm and flowchart; Unit 2 through variables, input, int(),
#          arithmetic and output; Unit 3 through nested conditions, loops,
#          validation and counters.  Restricted to the §9 scope list - no
#          functions, collections, classes, files, modules, ternary operator.
# --------------------------------------------------------------------------- #
Q3 = [
    dict(unit_online=(4, 5), bloom="C", marks=9,
         body=[
             ("p", "An electricity distribution company wants a billing "
                   "system for its consumers."),
             ("p", "First write an algorithm and draw a flowchart for the "
                   "system described below, and then write a Python program "
                   "for the same system."),
             ("p", "System specification : the program processes the bills of "
                   "exactly eight consumers, one after another, using a loop. "
                   "For every consumer it reads the consumer number, the "
                   "previous month meter reading, the current month meter "
                   "reading (all integers) and the type of connection "
                   "(1 for domestic, 2 for commercial). The program must"),
             ("p", "a)     validate the readings : the previous reading must "
                   "not be negative and the current reading must not be "
                   "smaller than the previous reading; if either test fails, "
                   "print INVALID READING and read the two readings again, "
                   "repeating this until a valid pair is entered;"),
             ("p", "b)     validate the type of connection : if the value "
                   "entered is neither 1 nor 2, print INVALID TYPE and read "
                   "the type again until it is 1 or 2;"),
             ("p", "c)     compute the units consumed as current reading "
                   "minus previous reading;"),
             ("p", "d)     compute the energy charge : for a domestic "
                   "connection the first 100 units are charged at Rs. 3 per "
                   "unit, the next 100 units at Rs. 5 per unit and every "
                   "remaining unit at Rs. 7 per unit; for a commercial "
                   "connection the first 200 units are charged at Rs. 9 per "
                   "unit and every remaining unit at Rs. 12 per unit;"),
             ("p", "e)     add a fixed charge of Rs. 50 for a domestic "
                   "connection and Rs. 150 for a commercial connection;"),
             ("p", "f)     if the units consumed are more than 500, add a "
                   "demand surcharge equal to 5 percent of the energy charge, "
                   "computed with integer arithmetic only; otherwise the "
                   "surcharge is zero;"),
             ("p", "g)     compute the total bill as energy charge plus fixed "
                   "charge plus surcharge, and decide the category : HIGH if "
                   "the total bill is more than Rs. 5000, MEDIUM if it is "
                   "more than Rs. 2000 but not more than Rs. 5000, and LOW "
                   "otherwise;"),
             ("p", "h)     print the consumer number, the units consumed, the "
                   "energy charge, the fixed charge, the surcharge, the total "
                   "bill and the category for that consumer;"),
             ("p", "i)     after all eight consumers have been processed, "
                   "print the number of consumers in each of the three "
                   "categories, the total units consumed by all eight "
                   "consumers together, and the highest total bill along "
                   "with the consumer number that received it."),
             ("p", "Use of functions, lists, tuples, dictionaries, sets, "
                   "files or modules is not allowed. Do not use the ternary "
                   "(conditional) operator form."),
         ]),
]

OFFLINE_UNIT_TARGET = {1: 2, 2: 6, 3: 8}    # guidelines §3
T1_UNIT_BLUEPRINT = {1: 2, 2: 10, 3: 13}    # guidelines §3


# --------------------------------------------------------------------------- #
# OOXML helpers
# --------------------------------------------------------------------------- #
def _w(tag, **attrs):
    e = OxmlElement(tag)
    for k, v in attrs.items():
        e.set(qn("w:" + k), str(v))
    return e


def add_run(parent, text, size=Q_SIZE, bold=False, italic=False, font=Q_FONT):
    r = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    rf = OxmlElement("w:rFonts")
    for a in ("ascii", "hAnsi", "cs"):
        rf.set(qn("w:" + a), font)
    rPr.append(rf)
    if bold:
        rPr.append(_w("w:b"))
        rPr.append(_w("w:bCs"))
    if italic:
        rPr.append(_w("w:i"))
        rPr.append(_w("w:iCs"))
    rPr.append(_w("w:sz", val=size))
    rPr.append(_w("w:szCs", val=size))
    r.append(rPr)
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    r.append(t)
    parent.append(r)
    return r


def add_tab(parent, size=OPT_SIZE, font=Q_FONT):
    r = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    rf = OxmlElement("w:rFonts")
    for a in ("ascii", "hAnsi", "cs"):
        rf.set(qn("w:" + a), font)
    rPr.append(rf)
    rPr.append(_w("w:sz", val=size))
    rPr.append(_w("w:szCs", val=size))
    r.append(rPr)
    r.append(OxmlElement("w:tab"))
    parent.append(r)
    return r


def add_para(parent, align=None, indent=None, tabstops=None, spacing_after=0,
             line=240, exact=False):
    p = OxmlElement("w:p")
    pPr = OxmlElement("w:pPr")
    if tabstops:
        tabs = OxmlElement("w:tabs")
        for pos in tabstops:
            tabs.append(_w("w:tab", val="left", pos=pos))
        pPr.append(tabs)
    pPr.append(_w("w:spacing", after=spacing_after, line=line,
                  lineRule="exact" if exact else "auto"))
    if indent:
        pPr.append(_w("w:ind", left=indent))
    if align:
        pPr.append(_w("w:jc", val=align))
    p.append(pPr)
    parent.append(p)
    return p


def clear_cell(tc):
    for child in list(tc):
        if child.tag in (qn("w:p"), qn("w:tbl")):
            tc.remove(child)


def cell_paras(tc):
    return tc.findall(qn("w:p"))


def _insert_before_rpr(pPr, element):
    """pPr children are order sensitive: ... spacing, ind, jc, rPr."""
    rpr = pPr.find(qn("w:rPr"))
    if rpr is not None:
        rpr.addprevious(element)
    else:
        pPr.append(element)


def set_line(tc, text, bold=False, size=Q_SIZE, align=None):
    """Replace the whole content of a header cell by one formatted line."""
    proto = cell_paras(tc)[0]
    pPr = proto.find(qn("w:pPr"))
    clear_cell(tc)
    p = OxmlElement("w:p")
    if pPr is not None:
        p.append(copy.deepcopy(pPr))
    tc.append(p)
    add_run(p, text, size=size, bold=bold)
    if align:
        _insert_before_rpr(p.find(qn("w:pPr")), _w("w:jc", val=align))
    return p


# --------------------------------------------------------------------------- #
# content writers
# --------------------------------------------------------------------------- #
def write_blocks(tc, blocks):
    clear_cell(tc)
    for kind, payload in blocks:
        if kind == "p":
            add_run(add_para(tc), payload, size=Q_SIZE)
        elif kind == "code":
            for line in payload:
                stripped = line.lstrip(" ")
                depth = (len(line) - len(stripped)) // 4
                p = add_para(tc, indent=180 + depth * 180, line=CODE_LINE)
                add_run(p, stripped, size=CODE_SIZE, font=CODE_FONT)


def write_options(tc, opts):
    """Six options, horizontal, on tab stops (guidelines §4, §10)."""
    letters = "(a)", "(b)", "(c)", "(d)", "(e)", "(f)"
    longest = max(len(o) for o in opts)
    if longest <= 18:
        per_line, stops = 3, [2800, 5600]
    elif longest <= 42:
        per_line, stops = 2, [4200]
    else:
        per_line, stops = 1, []
    for i in range(0, len(opts), per_line):
        p = add_para(tc, tabstops=stops)
        for k, opt in enumerate(opts[i:i + per_line]):
            if k:
                add_tab(p)
            add_run(p, f"{letters[i + k]} {opt}", size=OPT_SIZE)


def write_mcq_cell(tc, mcq, opts):
    blocks = [("p", s) for s in mcq["stem"]]
    if mcq["code"]:
        blocks.append(("code", mcq["code"]))
    write_blocks(tc, blocks)
    write_options(tc, opts)


# --------------------------------------------------------------------------- #
# table assembly
# --------------------------------------------------------------------------- #
def rebuild_question_table(tbl, rows_spec, online=False):
    """rows_spec entries: q, sub, blocks | (stem, code, opts), marks, bloom"""
    trs = tbl.findall(qn("w:tr"))
    if online:
        # The ONLINE table has no separate header row: row 0 is
        # 'Q-3 | (A) | <empty> | Marks | Bloom Taxonomy' and row 1 is the
        # template's own empty question row.  Row 0 therefore becomes the
        # header - its 'Q-3' / '(A)' placeholders are cleared so that the
        # number is not printed twice - and row 1 supplies the prototype,
        # which keeps the template's own question row height (§10).
        header = trs[0]
        for tc in header.findall(qn("w:tc"))[:3]:
            clear_cell(tc)
            add_para(tc)
        proto_q = trs[1]
    else:
        header = trs[0]
        proto_q = trs[1]
    for tr in trs[1:]:
        tbl.remove(tr)

    for spec in rows_spec:
        tr = copy.deepcopy(proto_q)          # template row heights preserved
        tcs = tr.findall(qn("w:tc"))
        set_line(tcs[0], spec["q"], bold=True)
        set_line(tcs[1], spec["sub"], bold=True)
        if spec.get("opts") is not None:
            write_mcq_cell(tcs[2], spec["mcq"], spec["opts"])
        else:
            write_blocks(tcs[2], spec["blocks"])
        set_line(tcs[3], spec["marks"], bold=True, size=MARK_SIZE, align="center")
        set_line(tcs[4], spec["bloom"], bold=True, size=MARK_SIZE, align="center")
        tbl.append(tr)


def make_table_flowable(tbl):
    tblPr = tbl.find(qn("w:tblPr"))
    floating = tblPr.find(qn("w:tblpPr"))
    if floating is not None:
        tblPr.remove(floating)
    if tblPr.find(qn("w:jc")) is None:
        tblPr.find(qn("w:tblW")).addnext(_w("w:jc", val="center"))


def spacer_before(tbl):
    p = OxmlElement("w:p")
    pPr = OxmlElement("w:pPr")
    pPr.append(_w("w:spacing", after=0, line=240, lineRule="auto"))
    p.append(pPr)
    add_run(p, "", size=HEADER_SIZE)
    tbl.addprevious(p)


def repeat_header_row(tr):
    trPr = tr.find(qn("w:trPr"))
    if trPr is None:
        trPr = OxmlElement("w:trPr")
        tr.insert(0, trPr)
    if trPr.find(qn("w:tblHeader")) is None:
        trPr.append(_w("w:tblHeader"))


def page_break_between_sections(doc):
    body = doc.element.body
    tables = body.findall(qn("w:tbl"))
    bloom_offline, header_online = tables[2], tables[3]
    node = bloom_offline.getnext()
    while node is not None and node is not header_online:
        nxt = node.getnext()
        if node.tag == qn("w:p"):
            body.remove(node)
        node = nxt
    p = OxmlElement("w:p")
    pPr = OxmlElement("w:pPr")
    pPr.append(_w("w:spacing", after=0, line=240, lineRule="auto"))
    p.append(pPr)
    r = OxmlElement("w:r")
    r.append(_w("w:br", type="page"))
    p.append(r)
    header_online.addprevious(p)


# --------------------------------------------------------------------------- #
# header block
# --------------------------------------------------------------------------- #
def replace_text_everywhere(doc, old, new):
    n = 0
    for t in doc.element.body.iter(qn("w:t")):
        if t.text and old in t.text:
            t.text = t.text.replace(old, new)
            n += 1
    return n


def mark(n):
    """Department bracketed marks format - guidelines §10: [0.5] [1] [03] [09]."""
    if n == 1:
        return "[1]"
    if n == 0.5:
        return "[0.5]"
    return f"[{int(n):02d}]"


def set_run_text(r, text):
    """Rewrite a run's text and leave its formatting (w:rPr) untouched."""
    ts = r.findall(qn("w:t"))
    if not ts:
        t = OxmlElement("w:t")
        r.append(t)
        ts = [t]
    for extra in ts[1:]:
        r.remove(extra)
    ts[0].text = text
    ts[0].set(qn("xml:space"), "preserve")


def fill_header(tbl, offline):
    """Populate only the template's variable fields, run by run.

    The template keeps 'B.E. SEMESTER-I' and 'BRANCH:' on two separate
    paragraphs - and 'SUBJECT-' / 'SUBJECT CODE:' likewise - and pads the DATE
    and TIME lines with spaces so that DURATION and MAX MARKS begin at a fixed
    column.  Rebuilding those cells collapses that layout, so only the runs
    that actually carry variable text are rewritten (guidelines §10: populate
    only the existing template fields).
    """
    rows = tbl.findall(qn("w:tr"))

    def runs(ri, pi):
        return rows[ri].findall(qn("w:tc"))[0].findall(qn("w:p"))[pi] \
                      .findall(qn("w:r"))

    # the branch list goes on the template's own 'BRANCH:' line, and the
    # subject codes on the template's own 'SUBJECT CODE:' line
    set_run_text(runs(4, 1)[1], " " + BRANCHES)
    set_run_text(runs(5, 0)[1], SUBJECT)
    set_run_text(runs(5, 1)[0], "SUBJECT CODE: " + SUBJECT_CODES)

    # date: only the three runs holding the sample date change, so the padding
    # runs that align DURATION stay exactly as the template ships them
    d = runs(6, 0)
    day, mon, year = DATE.split("-")
    set_run_text(d[1], day)
    set_run_text(d[3], mon)
    set_run_text(d[5], year[-1])

    # time: the run keeps its original length so MAX MARKS stays in the
    # template's column
    t = runs(7, 0)
    width = len(t[0].find(qn("w:t")).text)
    label = "TIME:  " + (OFFLINE_TIME if offline else ONLINE_TIME)
    set_run_text(t[0], label + " " * max(0, width - len(label)))
    line = "".join(x.text or "" for r in t for x in r.findall(qn("w:t")))
    if offline:
        # the template's offline page is printed with the full T1 total
        set_run_text(t[6], t[6].find(qn("w:t")).text.replace("25", "16"))
    else:
        assert "09" in line, "online header does not carry 09 marks"
    assert "1.25" in "".join(x.text or "" for r in runs(6, 0)
                             for x in r.findall(qn("w:t"))), "duration changed"


def set_set_letter(doc, letter, current="A"):
    n = 0
    for txbx in doc.element.body.iter(qn("w:txbxContent")):
        for t in txbx.iter(qn("w:t")):
            if t.text == current:
                t.text = letter
                n += 1
    return n


def page_break_index(doc):
    """body index of the paragraph that carries the section page break."""
    for i, el in enumerate(doc.element.body):
        if el.tag == qn("w:p"):
            for br in el.findall(".//" + qn("w:br")):
                if br.get(qn("w:type")) == "page":
                    return i
    raise AssertionError("template page break not found")


def split_pages(doc, keep):
    """Cut a combined offline+online paper down to one of its two pages.

    The template holds both pages in a single w:sectPr section, separated by
    one explicit page-break paragraph, so a page is removed by dropping the
    body elements on the other side of that paragraph.  Everything else -
    borders, merged cells, row heights, column widths, the Bloom table and the
    section properties - is left exactly as the template defines it (§10).
    """
    body = doc.element.body
    br = page_break_index(doc)
    sect = body.find(qn("w:sectPr"))
    kids = [el for el in list(body) if el is not sect]
    for el in (kids[br:] if keep == "offline" else kids[:br + 1]):
        body.remove(el)
    return doc


# --------------------------------------------------------------------------- #
# build
# --------------------------------------------------------------------------- #
def build_set(letter):
    """Build the complete two-page paper in memory and return it as bytes."""
    doc = Document(str(TEMPLATE))

    assert replace_text_everywhere(doc, "TEST 1/2/3 ( CO1/CO2/CO3 )",
                                   "TEST 1 ( CO1 )") == 2
    assert set_set_letter(doc, letter) >= 2

    page_break_between_sections(doc)
    tables = doc.tables
    fill_header(tables[0]._tbl, offline=True)
    fill_header(tables[3]._tbl, offline=False)
    for tbl in (tables[1]._tbl, tables[4]._tbl):
        if not KEEP_TEMPLATE_FLOATING_TABLE:
            make_table_flowable(tbl)
        spacer_before(tbl)
    repeat_header_row(tables[1]._tbl.findall(qn("w:tr"))[0])

    # ---- offline ----------------------------------------------------------
    rows_spec = [dict(q="Q-1", sub="Sr. No.",
                      blocks=[("p", "MCQ ( 1 Mark each )")],
                      marks=mark(len(MCQS)), bloom="")]
    key = {}
    for n, src in enumerate(SET_ORDER[letter], start=1):
        mcq = MCQS[src]
        rows_spec.append(dict(q="", sub=f"{n})", mcq=mcq, opts=mcq["opts"],
                              marks=mark(1), bloom=mcq["bloom"]))
        key[f"Q-1 {n})"] = "(%s)" % "abcdef"[mcq["ans"]]
    for n, q in enumerate(Q2, start=1):
        rows_spec.append(dict(q="Q-2" if n == 1 else "", sub=f"{n})",
                              blocks=q["body"], marks=mark(q["marks"]),
                              bloom=q["bloom"]))
        key[f"Q-2 {n})"] = f"descriptive, {q['marks']} marks"
    rebuild_question_table(tables[1]._tbl, rows_spec)

    # ---- online -----------------------------------------------------------
    rows_spec = []
    for n, q in enumerate(Q3, start=1):
        rows_spec.append(dict(q="Q-3" if n == 1 else "", sub="(A)",
                              blocks=q["body"], marks=mark(q["marks"]),
                              bloom=q["bloom"]))
        key["Q-3 (A)"] = "integrated 9 mark question, outside PB"
    rebuild_question_table(tables[4]._tbl, rows_spec, online=True)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf, key


# --------------------------------------------------------------------------- #
# faculty-only documents (never part of the student paper - guidelines §5, §11)
# --------------------------------------------------------------------------- #
def write_audit_sheet(keys, path):
    unit_marks = {1: 0, 2: 0, 3: 0}
    for m in MCQS:
        unit_marks[m["unit"]] += 1
    for q in Q2:
        unit_marks[q["unit"]] += q["marks"]
    on_u2, on_u3 = Q3[0]["unit_online"]
    pb_marks = len(MCQS) + sum(q["marks"] for q in Q2 if q["pb"])
    out_marks = sum(q["marks"] for q in Q2 if not q["pb"])
    total_unit = {1: unit_marks[1], 2: unit_marks[2] + on_u2,
                  3: unit_marks[3] + on_u3}

    L = ["FACULTY-ONLY AUDIT SHEET - do not circulate, do not attach to the QP",
         "T1 (CO1) Computer Programming using Python-I | SEM-I Batch 2026",
         "QP setter MDP | exam 29-Sep-2026 | deadline 25-Sep-2026 12:00 midnight",
         "",
         "MARKS AND BLUEPRINT CHECK",
         f"  offline total ....... {sum(1 for _ in MCQS) + sum(q['marks'] for q in Q2)}"
         f" (required 16)",
         f"  from Practice Book .. {pb_marks}  (required 10)",
         f"  outside PB .......... {out_marks}  (required 6)",
         f"  online total ........ {sum(q['marks'] for q in Q3)} (required 9,"
         f" entirely outside PB)",
         f"  offline unit split .. {unit_marks[1]} / {unit_marks[2]} /"
         f" {unit_marks[3]}   (target 2 / 6 / 8)",
         f"  online unit split ... 0 / {on_u2} / {on_u3}",
         f"  T1 unit total ....... {total_unit[1]} / {total_unit[2]} /"
         f" {total_unit[3]}   (blueprint 2 / 10 / 13)",
         "",
         "QUESTION TRACEABILITY (guidelines §5)",
         "  position | unit | source            | marks | class      | bloom",
         ]
    for i, m in enumerate(MCQS, start=1):
        L.append(f"  MCQ {i:>2}    |  {m['unit']}   | PB Sr. No. {m['pb']:>3}    |"
                 f"   1   | PB         | {m['bloom']}")
    for i, q in enumerate(Q2, start=1):
        src = f"PB Sr. No. {q['pb']:>3}" if q["pb"] else "original"
        cls = "PB        " if q["pb"] else "outside PB"
        L.append(f"  Q-2 {i})    |  {q['unit']}   | {src:<17} |"
                 f"   {q['marks']}   | {cls} | {q['bloom']}")
    L.append(f"  Q-3 (A)  | 2,3  | original          |   9   | outside PB |"
             f" {Q3[0]['bloom']}")
    L.append("")
    for letter in ("A", "B", "C"):
        L.append(f"--- ANSWER KEY SET {letter} (MCQ order {SET_ORDER[letter]}) ---")
        for k, v in keys[letter].items():
            L.append(f"  {k:<12} {v}")
        L.append("")
    L.append("OPEN ITEM (§6): the preferred 10 x 0.5 + 5 x 1 MCQ pattern needs")
    L.append(f"written HOD clearance; this paper uses {len(MCQS)} 1-mark MCQs "
             "in Q-1.")
    L.append("")
    L.append("PB Sr. No. 159 was rejected: its printed code cannot be")
    L.append("re-indented unambiguously from the Practice Book and it does not")
    L.append("reproduce the printed key, so it could not be verified (§6).")
    L.append("")
    L.append("OUTSTANDING (§10): a PDF preview must be generated and every page")
    L.append("visually checked in Word before these files are called final.")
    path.write_text("\n".join(L), encoding="utf-8")


def write_compliance_note(path):
    L = [f"""L. J. INSTITUTE OF ENGINEERING AND TECHNOLOGY, AHMEDABAD
TEST 1 (CO1) - Computer Programming using Python-I - SEM-I Batch 2026
QP SETTER : MDP      EXAM : 29-Sep-2026 (Tue)
Compliance map for PYTHON-1_T1_QUESTION_PAPER_GENERATION_GUIDELINES.pdf

S1  Identity/schedule : 16 offline / 09 online / 25 total; 2:15-3:30 and
    4:15-5:30; 1.25 hrs each; deadline 25-Sep-2026; qp.fyall@gmail.com.
S2  Branches/codes    : all 15 subject codes printed in both headers.
S3  Blueprint         : offline {OFFLINE_UNIT_TARGET[1]} / {OFFLINE_UNIT_TARGET[2]} / {OFFLINE_UNIT_TARGET[3]}; online 0 / 4 / 5;
    T1 total 2 / 10 / 13 = the issued blueprint exactly. Nothing outside
    Units 1-3 is used.
S4  Setting rules     : conceptual/reasoning questions only; no definitions,
    formulae, short notes, True/False, Yes/No or fill in the blanks; Q-1 is
    the compulsory MCQ; 3 main questions in total; 3 sets with identical
    questions and only the MCQ sequence reshuffled; no faculty name used.
S5  PB split          : offline = 8 MCQ marks + 2 marks of PB Sr. No. 30 = 10
    from the Practice Book, and 6 marks outside it; online 9 marks entirely
    outside it. Every PB question keeps its stem unchanged (only the options
    were rebuilt to six), so each remains a PB question.
S6  MCQ rules         : every MCQ is 1 mark with exactly 6 options; every
    output MCQ uses e) Error and f) None of the above; every output MCQ was
    executed (see tools/verify_qp.py) before finalising.
    OPEN ITEM: the preferred 10 x 0.5 + 5 x 1 pattern was NOT used because it
    needs written HOD clearance (it clashes with the "at least seven
    questions in Q-1(A) and Q-1(B)" rule). This paper uses the compliant
    Case-2 pattern: eight 1-mark MCQs, all compulsory in Q-1.
    Recall-only PB items (Sr 1-20) and the trivial algorithm trace (Sr 21) are
    deliberately not used, because §4 forbids easy recall questions; PB
    Sr. No. 159 was dropped because its code cannot be re-indented
    unambiguously and it does not reproduce the printed key.
S7  Workflow          : blueprint frozen first; PB serial numbers recorded in
    the faculty-only audit sheet; outside-PB questions test operator
    precedence, type conversion, logical operators, nested decisions,
    while/for tracing, break/continue, nested loops, validation, counters.
S9  Online paper      : ONE integrated 9-mark question under Q-3(A), outside
    the Practice Book, Bloom C, strictly inside the T1 scope (variables,
    input/output, int(), arithmetic and comparison operators, if/elif/else,
    nested conditions, for, while, counters, validation). Unit 1 is connected
    through the required algorithm and flowchart; no functions, collections,
    classes, files or modules.
S10 Format controls   : built on the original Python-I T1-T3 template file;
    enrollment boxes, merged cells, borders, row heights, column widths, tab
    stops and the Bloom table untouched; only the template's own fields are
    populated and they are populated RUN BY RUN, so the header keeps its own
    paragraph layout - 'B.E. SEMESTER-I' and 'BRANCH:' stay on separate lines,
    as do 'SUBJECT-' and 'SUBJECT CODE:', and DURATION / MAX MARKS stay in the
    template's columns. Code is Courier New with indentation preserved; the
    marks column uses the bracketed form named in §10 ([1] [02] [03] [08]
    [09]). NOTE: §10 makes that format conditional on the approved sample, and
    no approved sample paper is in the QP documents folder - confirm it against
    your sample, or switch it in one place (the mark() helper).
    ONE DEVIATION, deliberate: the template's question tables are floating
    (w:tblpPr) and a floating table is not split across pages by Word, so they
    were converted to inline centred tables of identical width. Set
    KEEP_TEMPLATE_FLOATING_TABLE = True in tools/build_t1_qp.py to restore the
    shipped structure exactly.
    OUTSTANDING: a PDF preview must be produced and every page visually
    checked in Word before these files are treated as final - that step cannot
    be performed in this environment.
S11 Audit             : run tools/verify_qp.py; it re-executes the MCQ code,
    checks 6 options and the e)/f) convention, sums the marks, checks the unit
    split, the PB/outside-PB split, the three sets and the confidentiality of
    the student copies.

FILES  - offline and online are issued as separate papers, as the test itself
is held in two sittings (offline 2:15-3:30 pm, online 4:15-5:30 pm):
  SET A/B/C_T1_PYTHON-1_TEST PAPER (OFFLINE)_SEM I_MDP.docx -> three sets
  ONLINE (COMMON)_T1_PYTHON-1_TEST PAPER_SEM I_MDP.docx     -> one common paper
      §4 requires the three sets to carry identical questions with only the
      MCQ sequence reshuffled.  The online paper contains no MCQ, so a single
      online paper serves all three sets.  Its 'Set:' box reads 'Set: 1'
      because that box is 0.86 inch wide at 16 pt and cannot hold a word -
      change it if the department prefers a different label.
  AUDIT SHEET AND ANSWER KEY (faculty only) -> never send with the QP
"""]
    path.write_text(L[0], encoding="utf-8")


def offline_name(letter):
    return f"SET {letter}_T1_PYTHON-1_TEST PAPER (OFFLINE)_SEM I_MDP.docx"


ONLINE_NAME = "ONLINE (COMMON)_T1_PYTHON-1_TEST PAPER_SEM I_MDP.docx"


def main():
    OUTDIR.mkdir(exist_ok=True)
    keys = {}
    online = None
    for letter in ("A", "B", "C"):
        buf, keys[letter] = build_set(letter)
        off = split_pages(Document(buf), "offline")
        off.save(str(OUTDIR / offline_name(letter)))
        print("written:", OUTDIR / offline_name(letter))
        if online is None:
            # §4: the three sets carry identical questions and only the MCQ
            # sequence is reshuffled, so the online paper - which holds no
            # MCQ - is one common paper for all three sets.
            buf.seek(0)
            online = split_pages(Document(buf), "online")
            set_set_letter(online, "1", current=letter)
            online.save(str(OUTDIR / ONLINE_NAME))
            print("written:", OUTDIR / ONLINE_NAME)
    write_audit_sheet(keys, OUTDIR / "AUDIT SHEET AND ANSWER KEY_T1_PYTHON-I_MDP_faculty only.txt")
    write_compliance_note(OUTDIR / "QP SETTING NOTE_T1_PYTHON-I_MDP.txt")
    print("written: faculty-only audit sheet and compliance note")


if __name__ == "__main__":
    main()
