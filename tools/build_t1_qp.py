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
import shutil
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
CODE_LINE = 216   # 10.8 pt exact leading for code lines
OPT_SIZE = 20     # 10 pt options
MARK_SIZE = 20    # 10 pt marks / Bloom columns
HEADER_SIZE = 24  # 12 pt institute header block (template size)

# --------------------------------------------------------------------------- #
# Q-1 : ten 1-mark MCQs, all from the T1 Practice Book (guidelines §5, §6)
#
# output MCQs carry e) Error and f) None of the above and the option order is
# FIXED - guidelines §4 allow only the MCQ *sequence* to be reshuffled.
# `is_output` marks the questions the e)/f) convention applies to.
# --------------------------------------------------------------------------- #
MCQS = [
    dict(pb=8, unit=1, bloom="U", ans=2, is_output=False,
         stem=["A program that reads each of the instructions in mnemonic form "
               "and translates it into the machine-language equivalent is ______"],
         code=[],
         opts=["Machine language", "Interpreter", "Assembler",
               "Compiler", "Linker", "Loader"]),
    dict(pb=21, unit=1, bloom="A", ans=1, is_output=True,
         stem=["What will be the output of the following algorithm if the "
               "input is 4?"],
         code=["Algorithm:", "1. Start", "2. Set x = input value",
               "3. Set y = x * 2",
               "4. If y > 5, then print \"Large\", otherwise print \"Small\".",
               "5. Stop"],
         opts=["Small", "Large", "8", "4", "Error", "None of the above"]),
    dict(pb=64, unit=2, bloom="A", ans=2, is_output=True,
         stem=["What are the values of the following Python expressions?"],
         code=["2**(3**2)", "(2**3)**2", "2**3**2"],
         opts=["64, 512, 64", "64, 64, 64", "512, 64, 512",
               "512, 512, 512", "Error", "None of the above"]),
    dict(pb=54, unit=2, bloom="A", ans=3, is_output=True,
         stem=["What will be the output of the following program on execution?"],
         code=["a=0", "b=6", "x=(a or b) or ((a and a) or (a and b))", "print(x)"],
         opts=["0", "True", "False", "6", "Error", "None of the above"]),
    dict(pb=86, unit=2, bloom="N", ans=4, is_output=True,
         stem=["What will be the output of the following program on execution?"],
         code=["a=50", "b=60", "print((a and b)/False)"],
         opts=["0", "60", "50", "60.0", "Error", "None of the above"]),
    dict(pb=107, unit=3, bloom="N", ans=0, is_output=True,
         stem=["Given the nested if-else structure below, what will be the "
               "value of x after code execution completes?"],
         code=["x = 0", "a = 0", "b = -5",
               "if a > 0:", "    if b < 0:", "        x = x + 5",
               "    elif a > 5:", "        x = x + 4",
               "    else:", "        x = x + 3",
               "else:", "    x = x + 4", "print(x)"],
         opts=["4", "0", "3", "2", "Error", "None of the above"]),
    dict(pb=165, unit=3, bloom="A", ans=1, is_output=True,
         stem=["What will be the output of the following program on execution?"],
         code=["x=0", "while x<10:", "    if x%3==0:", "        x+=5",
               "        continue", "    if x%2==0:", "        x+=14",
               "    else:", "        x+=1", "else:", "    x+=1", "print(x)"],
         opts=["10", "12", "11", "0", "Error", "None of the above"]),
    dict(pb=167, unit=3, bloom="N", ans=1, is_output=True,
         stem=["What is the output of the following code?"],
         code=["n=10", "i=1", "while(i<=n):", "    k=0", "    if(n%i==0):",
               "        j=1", "        while(j<=i):", "            if(i%j==0):",
               "                k=k+1", "            j=j+1",
               "        if(k==2):", '            print(i,end=" ")', "    i=i+1"],
         opts=["1 2 5 10", "2 5", "2 3 5 7", "2 5 10",
               "Error", "None of the above"]),
    dict(pb=176, unit=3, bloom="N", ans=2, is_output=True,
         stem=["What will be the output of the following program on execution?"],
         code=["x=0", "count=0", "while x<15:", "    if x%2==0:", "        x+=1",
               "        continue", "    if x%3==0:", "        x+=1",
               "        continue", "    if count==5:", "        break",
               "    count+=1", "print(x,count)"],
         opts=["5 5", "0 5", "1 5", "15 5", "Error", "None of the above"]),
    dict(pb=160, unit=3, bloom="A", ans=3, is_output=True,
         stem=["What should be the output of the following python code snippet?"],
         code=["a=5", "b=7", "c=2",
               "if a>b:", "    a,b = b,a",
               "if a>c:", "    a,c = c,a",
               "if b>c:", "    b,c = c,b",
               'print(a,b,c,end=",")'],
         opts=["7 5 2", "2 7 5", "5 2 7", "2 5 7",
               "Error", "None of the above"]),
]

# Q-1 sequence per set - only the order changes, never the options (§4, §7)
SET_ORDER = {
    "A": [5, 6, 2, 9, 4, 7, 8, 3, 0, 1],
    "B": [7, 0, 3, 5, 8, 1, 4, 2, 9, 6],
    "C": [2, 9, 1, 4, 6, 0, 5, 8, 3, 7],
}

# --------------------------------------------------------------------------- #
# Q-2 : 6 marks outside the Practice Book (guidelines §5, §7)
#       unit split 3 + 3 so that the offline paper is 2 / 6 / 8
# --------------------------------------------------------------------------- #
Q2 = [
    dict(unit=2, bloom="A", marks=3,
         body=[
             ("p", "The following statements are executed in one cell of a "
                   "Jupyter Notebook:"),
             ("code", ["total = 7 / 2",
                       "count = int(total)",
                       "result = count ** 2 + 9 % 4",
                       "print(result, type(result), type(total))"]),
             ("p", "(i)  Write the exact output produced by the cell."),
             ("p", "(ii) Explain the result by stating the order in which the "
                   "operators / , ** , % and + are applied and what the int() "
                   "conversion does to the value."),
         ]),
    dict(unit=3, bloom="A", marks=3,
         body=[
             ("p", "Write a Python program that reads a positive integer from "
                   "the user and prints its digital root, i.e. the single digit "
                   "obtained by repeatedly adding the digits of the number "
                   "until only one digit remains. Use of string, list, tuple or "
                   "any other data structure and their built-in functions is "
                   "not allowed."),
             ("p", "Example :  Input : 9875   then  9+8+7+5 = 29  ->  2+9 = 11  "
                   "->  1+1 = 2,  so the output is  2"),
         ]),
]

# --------------------------------------------------------------------------- #
# Q-3(A) : ONE integrated 9-mark question, entirely outside the Practice Book
#          (guidelines §9).  Bloom C.  Unit-1 connection = the required
#          algorithm and flowchart; Units 2 and 3 through the program.
#          Restricted to the T1 scope list of §9 - no functions, collections,
#          classes, files, modules and no ternary operator.
# --------------------------------------------------------------------------- #
Q3 = [
    dict(unit_online=(4, 5), bloom="C", marks=9,
         body=[
             ("p", "An electricity distribution company wants a billing system "
                   "for its consumers."),
             ("p", "First write an algorithm and draw a flowchart for the "
                   "system described below, and then write a Python program "
                   "for the same system."),
             ("p", "System specification : the program processes the bills of "
                   "exactly five consumers, one after another, using a loop. "
                   "For every consumer it reads the consumer number, the "
                   "previous month meter reading, the current month meter "
                   "reading (all integers) and the type of connection "
                   "(1 for domestic, 2 for commercial). The program must"),
             ("p", "a)     validate the readings : if the current reading is "
                   "smaller than the previous reading, print INVALID READING "
                   "and read the two readings again, repeating this until a "
                   "valid pair is entered;"),
             ("p", "b)     compute the units consumed as current reading minus "
                   "previous reading;"),
             ("p", "c)     compute the energy charge : for a domestic "
                   "connection the first 100 units are charged at Rs. 3 per "
                   "unit, the next 100 units at Rs. 5 per unit and every "
                   "remaining unit at Rs. 7 per unit; for a commercial "
                   "connection every unit is charged at Rs. 9 per unit;"),
             ("p", "d)     add a service charge of Rs. 50 for a domestic "
                   "connection and Rs. 150 for a commercial connection, and "
                   "compute the total bill;"),
             ("p", "e)     print the consumer number, the units consumed, the "
                   "energy charge, the service charge and the total bill;"),
             ("p", "f)     print HIGH USAGE if the total bill is more than "
                   "Rs. 2000 and NORMAL USAGE otherwise;"),
             ("p", "g)     after all five consumers have been processed, print "
                   "how many of them were billed more than Rs. 2000."),
             ("p", "Use of functions, lists, tuples, dictionaries, sets, files "
                   "or modules is not allowed."),
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
                p = add_para(tc, indent=180 + depth * 180, line=CODE_LINE,
                             exact=True)
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
        # the ONLINE table has no separate header row: row 0 is
        # 'Q-3 | (A) | <empty> | Marks | Bloom Taxonomy'
        header = trs[0]
        for tc in header.findall(qn("w:tc"))[:3]:
            clear_cell(tc)
            add_para(tc)
        proto_q = copy.deepcopy(header)
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


def fill_header(tbl, offline):
    rows = tbl.findall(qn("w:tr"))
    set_line(rows[4].findall(qn("w:tc"))[0],
             "B.E. SEMESTER-I" + " " * 42 + "BRANCH: " + BRANCHES,
             bold=True, size=HEADER_SIZE)
    set_line(rows[5].findall(qn("w:tc"))[0],
             "SUBJECT- " + SUBJECT + " " * 24 + "SUBJECT CODE: " + SUBJECT_CODES,
             bold=True, size=HEADER_SIZE)
    set_line(rows[6].findall(qn("w:tc"))[0],
             "DATE: " + DATE + " " * 48 + "DURATION:    " + DURATION,
             bold=True, size=HEADER_SIZE)
    time = OFFLINE_TIME if offline else ONLINE_TIME
    marks = OFFLINE_MAX if offline else ONLINE_MAX
    set_line(rows[7].findall(qn("w:tc"))[0],
             "TIME:  " + time + " " * 44 + "MAX MARKS: " + marks,
             bold=True, size=HEADER_SIZE)


def set_set_letter(doc, letter):
    n = 0
    for txbx in doc.element.body.iter(qn("w:txbxContent")):
        for t in txbx.iter(qn("w:t")):
            if t.text == "A":
                t.text = letter
                n += 1
    return n


# --------------------------------------------------------------------------- #
# build
# --------------------------------------------------------------------------- #
def build_set(letter, outpath):
    shutil.copyfile(TEMPLATE, outpath)
    doc = Document(str(outpath))

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
                      marks="(10)", bloom="")]
    key = {}
    for n, src in enumerate(SET_ORDER[letter], start=1):
        mcq = MCQS[src]
        rows_spec.append(dict(q="", sub=f"{n})", mcq=mcq, opts=mcq["opts"],
                              marks="1", bloom=mcq["bloom"]))
        key[f"Q-1 {n})"] = "(%s)" % "abcdef"[mcq["ans"]]
    for n, q in enumerate(Q2, start=1):
        rows_spec.append(dict(q="Q-2" if n == 1 else "", sub=f"{n})",
                              blocks=q["body"], marks=str(q["marks"]),
                              bloom=q["bloom"]))
        key[f"Q-2 {n})"] = f"descriptive, {q['marks']} marks"
    rebuild_question_table(tables[1]._tbl, rows_spec)

    # ---- online -----------------------------------------------------------
    rows_spec = []
    for n, q in enumerate(Q3, start=1):
        rows_spec.append(dict(q="Q-3" if n == 1 else "", sub="(A)",
                              blocks=q["body"], marks="(09)", bloom=q["bloom"]))
        key["Q-3 (A)"] = "integrated 9 mark question, outside PB"
    rebuild_question_table(tables[4]._tbl, rows_spec, online=True)

    doc.save(str(outpath))
    return key


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
    total_unit = {1: unit_marks[1], 2: unit_marks[2] + on_u2,
                  3: unit_marks[3] + on_u3}

    L = ["FACULTY-ONLY AUDIT SHEET - do not circulate, do not attach to the QP",
         "T1 (CO1) Computer Programming using Python-I | SEM-I Batch 2026",
         "QP setter MDP | exam 29-Sep-2026 | deadline 25-Sep-2026 12:00 midnight",
         "",
         "MARKS AND BLUEPRINT CHECK",
         f"  offline total ....... {sum(1 for _ in MCQS) + sum(q['marks'] for q in Q2)}"
         f" (required 16)",
         f"  from Practice Book .. {len(MCQS)}  (required 10)",
         f"  outside PB .......... {sum(q['marks'] for q in Q2)}"
         f"  (required 6)",
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
        L.append(f"  Q-2 {i})    |  {q['unit']}   | original          |"
                 f"   {q['marks']}   | outside PB | {q['bloom']}")
    L.append(f"  Q-3 (A)  | 2,3  | original          |   9   | outside PB |"
             f" {Q3[0]['bloom']}")
    L.append("")
    for letter in ("A", "B", "C"):
        L.append(f"--- ANSWER KEY SET {letter} (MCQ order {SET_ORDER[letter]}) ---")
        for k, v in keys[letter].items():
            L.append(f"  {k:<12} {v}")
        L.append("")
    L.append("OPEN ITEM (§6): the preferred 10 x 0.5 + 5 x 1 MCQ pattern needs")
    L.append("written HOD clearance; this paper uses ten 1-mark MCQs in Q-1.")
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
S5  PB split          : offline = 10 MCQ marks from the Practice Book + 6
    marks outside it; online 9 marks entirely outside it. PB questions are
    used with their stems unchanged, so they remain PB questions.
S6  MCQ rules         : every MCQ is 1 mark with exactly 6 options; every
    output MCQ uses e) Error and f) None of the above; every output MCQ was
    executed (see tools/verify_qp.py) before finalising.
    OPEN ITEM: the preferred 10 x 0.5 + 5 x 1 pattern was NOT used because it
    needs written HOD clearance (it clashes with the "at least seven
    questions in Q-1(A) and Q-1(B)" rule). This paper uses the compliant
    Case-2 pattern: ten 1-mark MCQs, all compulsory in Q-1.
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
    enrollment boxes, merged cells, borders, column widths, tab stops and the
    Bloom table untouched; only the template's own fields populated; code in
    Courier New with indentation preserved; department marks style.
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

FILES
  SET A/B/C_T1_PYTHON-1_TEST PAPER_SEM I_MDP.docx  -> student papers
  ANSWER KEY + AUDIT SHEET (faculty only)          -> never send with the QP
"""]
    path.write_text(L[0], encoding="utf-8")


def main():
    OUTDIR.mkdir(exist_ok=True)
    keys = {}
    for letter in ("A", "B", "C"):
        name = f"SET {letter}_T1_PYTHON-1_TEST PAPER_SEM I_MDP.docx"
        keys[letter] = build_set(letter, OUTDIR / name)
        print("written:", OUTDIR / name)
    write_audit_sheet(keys, OUTDIR / "AUDIT SHEET AND ANSWER KEY_T1_PYTHON-I_MDP_faculty only.txt")
    write_compliance_note(OUTDIR / "QP SETTING NOTE_T1_PYTHON-I_MDP.txt")
    print("written: faculty-only audit sheet and compliance note")


if __name__ == "__main__":
    main()
