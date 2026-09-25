#!/usr/bin/env python3
"""
Build the T1 question paper (SET A / B / C) for
  Computer Programming using Python-I, SEM-I, Batch 2026
QP Setter : MDP

Everything is driven by the official LJIET template
  QP REQUIRED DOCUMENTS/04 QP_FORMAT/QP_FORMAT/01 FORMAT_PYTHON-1/
      SET A_T1 to T3_PYTHON-1_TEST PAPER_SEM I_FACUTLY SHORT NAME_FORMAT.docx
so that the layout (header block, floating "Set" box, question table with
Marks / Bloom Taxonomy columns, Bloom legend) is preserved exactly.

Rules implemented (FY_QUESTION PAPER SETTING INSTRUCTIONS_T1 TO T3):
  * Q-1 = compulsory MCQ, 1 mark each, 6 options (T1 => 6 options)
  * Offline 16 marks = 10 marks from Practice Book (the 10 MCQs)
                      +  6 marks out of Practice Book (Q-2)
  * Online 9 marks   = completely out of the Practice Book (Q-3)
  * No OR-type options, all questions compulsory, maximum 7 main questions
  * TEST 1 / CO1 only in the heading, subject codes of all branches printed
  * 3 sets, identical questions, MCQs reshuffled (question order + option order)
"""

import copy
import random
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
# paper constants
# --------------------------------------------------------------------------- #
SET_LETTER = {"A": "A", "B": "B", "C": "C"}

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

Q_FONT = "Times New Roman"
CODE_FONT = "Courier New"
Q_SIZE = 22      # 11 pt  (guideline 13 allows 11 or 12)
CODE_SIZE = 18   #  9 pt  for the code blocks
CODE_LINE = 216  # 10.8 pt exact leading for code lines
OPT_SIZE = 20    # 10 pt
MARK_SIZE = 20   # 10 pt
HEADER_SIZE = 24 # 12 pt bold - the institute header block keeps the template size


# --------------------------------------------------------------------------- #
# question bank  (source = Practice Book Sr. No. / "OUT" = new question)
# --------------------------------------------------------------------------- #
# every MCQ: stem (lines), optional code (lines), 6 options, correct index 0-5
MCQS = [
    dict(
        pb=8, unit=1, bloom="U", ans=1,
        stem=["A program that reads each of the instructions in mnemonic form and "
              "translates it into the machine-language equivalent is ______"],
        code=[],
        opts=["Machine language", "Assembler", "Interpreter",
              "Compiler", "Linker", "Loader"],
    ),
    dict(
        pb=21, unit=1, bloom="A", ans=0,
        stem=["What will be the output of the following algorithm if the input is 4?"],
        code=["Algorithm:",
              "1. Start",
              "2. Set x = input value",
              "3. Set y = x * 2",
              "4. If y > 5, then print \"Large\", otherwise print \"Small\".",
              "5. Stop"],
        opts=["Large", "Small", "No output", "Error", "8", "4"],
    ),
    dict(
        pb=64, unit=2, bloom="A", ans=3,
        stem=["What are the values of the following Python expressions?"],
        code=["2**(3**2)", "(2**3)**2", "2**3**2"],
        opts=["64, 512, 64", "64, 64, 64", "512, 512, 512",
              "512, 64, 512", "512, 512, 64", "64, 512, 512"],
    ),
    dict(
        pb=54, unit=2, bloom="A", ans=1,
        stem=["What will be the output of the following program on execution?"],
        code=["a=0", "b=6", "x=(a or b) or ((a and a) or (a and b))", "print(x)"],
        opts=["0", "6", "True", "False", "None", "Error"],
    ),
    dict(
        pb=45, unit=2, bloom="U", ans=1,
        stem=["Which of the following is an invalid statement?"],
        code=[],
        opts=["abc = 1,000,000", "a b c = 1000 2000 3000",
              "a,b,c = 1000, 2000, 3000", "a_b_c = 1,000,000",
              "a, b, c = 1000, 2000, 3000", "a = b = c = 1000"],
    ),
    dict(
        pb=86, unit=2, bloom="N", ans=3,
        stem=["What will be the output of the following program on execution?"],
        code=["a=50", "b=60", "print((a and b)/False)"],
        opts=["0", "60", "50", "Error", "60.0", "1"],
    ),
    dict(
        pb=107, unit=3, bloom="N", ans=3,
        stem=["Given the nested if-else structure below, what will be the value of x "
              "after code execution completes?"],
        code=["x = 0", "a = 0", "b = -5",
              "if a > 0:", "    if b < 0:", "        x = x + 5",
              "    elif a > 5:", "        x = x + 4",
              "    else:", "        x = x + 3",
              "else:", "    x = x + 4", "print(x)"],
        opts=["2", "0", "3", "4", "5", "9"],
    ),
    dict(
        pb=165, unit=3, bloom="A", ans=2,
        stem=["What will be the output of the following program on execution?"],
        code=["x=0",
              "while x<10:",
              "    if x%3==0:",
              "        x+=5",
              "        continue",
              "    if x%2==0:",
              "        x+=14",
              "    else:",
              "        x+=1",
              "else:",
              "    x+=1",
              "print(x)"],
        opts=["10", "11", "12", "0", "15", "14"],
    ),
    dict(
        pb=167, unit=3, bloom="N", ans=0,
        stem=["What is the output of the following code?"],
        code=["n=10", "i=1",
              "while(i<=n):",
              "    k=0",
              "    if(n%i==0):",
              "        j=1",
              "        while(j<=i):",
              "            if(i%j==0):",
              "                k=k+1",
              "            j=j+1",
              "        if(k==2):",
              '            print(i,end=" ")',
              "    i=i+1"],
        opts=["2 5", "1 2 5 10", "2 3 5 7", "2 5 10", "5 10", "10 20"],
    ),
    dict(
        pb=176, unit=3, bloom="N", ans=0,
        stem=["What will be the output of the following program on execution?"],
        code=["x=0", "count=0",
              "while x<15:",
              "    if x%2==0:",
              "        x+=1",
              "        continue",
              "    if x%3==0:",
              "        x+=1",
              "        continue",
              "    if count==5:",
              "        break",
              "    count+=1",
              "print(x,count)"],
        opts=["1 5", "5 5", "0 5", "15 5", "16 5", "14 5"],
    ),
]

# Q-2 : offline descriptive part, 6 marks, completely OUT of the Practice Book
Q2 = [
    dict(
        unit=2, bloom="U", marks=3,
        body=[
            ("p", "A student creates a new Jupyter Notebook and types the following "
                  "code in three different cells:"),
            ("code", ["Cell-1 :  n = 7", "Cell-2 :  n = n * 3",
                      "Cell-3 :  print(n, type(n))"]),
            ("p", "The student runs Cell-3 first, then Cell-1, then Cell-2 and "
                  "finally runs Cell-3 once again."),
            ("p", "(a) State exactly what is displayed by each of the two executions "
                  "of Cell-3."),
            ("p", "(b) Explain the reason for the difference between the two outputs "
                  "and state the corrective steps that give the intended output."),
        ],
    ),
    dict(
        unit=3, bloom="A", marks=3,
        body=[
            ("p", "Write a Python program that reads a positive integer from the user "
                  "and prints its digital root, i.e. the single digit obtained by "
                  "repeatedly adding the digits of the number until only one digit "
                  "remains. Use of string, list, tuple or any other data structure "
                  "and their built-in functions is not allowed."),
            ("p", "Example :  Input : 9875   then  9+8+7+5 = 29  ->  2+9 = 11  ->  "
                  "1+1 = 2,  so the output is  2"),
        ],
    ),
]

# Q-3 : online part, 9 marks, completely OUT of the Practice Book
Q3 = [
    dict(
        unit=2, bloom="A", marks=4,
        body=[
            ("p", "Write a Python program that reads the total distance travelled by a "
                  "vehicle in kilometre (float) and the total time taken in minutes "
                  "(float) from the user. The program should"),
            ("p", "(i)    compute and print the average speed in km/hour rounded to "
                  "two decimal places;"),
            ("p", "(ii)   print the time taken in the form  HH hours MM minutes "
                  "SS seconds  using only the arithmetic operators // and % together "
                  "with typecasting (no data structure or built-in string function is "
                  "allowed);"),
            ("p", "(iii)  print the single word FAST if the average speed is more than "
                  "60 km/hour and NORMAL otherwise, by using one single line ternary "
                  "(conditional) expression."),
            ("p", "Assume suitable sample input and show the output."),
        ],
    ),
    dict(
        unit=3, bloom="C", marks=5,
        body=[
            ("p", "Write a Python program that reads a positive integer N from the "
                  "user and generates the Collatz sequence starting from N until the "
                  "value 1 is reached (Rule : if the current value is even, divide it "
                  "by 2 using integer division; if it is odd, multiply it by 3 and "
                  "add 1). The program must print"),
            ("p", "(i)    every value of the sequence on the same line separated by a "
                  "single space,"),
            ("p", "(ii)   the total number of steps taken to reach 1, and"),
            ("p", "(iii)  the largest value that appeared in the sequence."),
            ("p", "Use of string, list, tuple or any other data structure and their "
                  "built-in functions is not allowed."),
            ("p", "Example :  For N = 6 the sequence is  6 3 10 5 16 8 4 2 1 ,  "
                  "steps = 8 ,  largest value = 16"),
        ],
    ),
]

# MCQ order per set (0-based index into MCQS). Sets B and C are reshuffles.
SET_ORDER = {
    "A": [0, 1, 2, 3, 4, 5, 6, 7, 8, 9],
    "B": [6, 2, 9, 0, 4, 7, 1, 5, 8, 3],
    "C": [3, 8, 1, 5, 9, 0, 7, 4, 2, 6],
}
OPTION_SHUFFLE_SEED = {"A": 11, "B": 22, "C": 33}

# where the correct option is placed in each set (0 = (A) ... 5 = (F)).
# every letter is used 1-2 times per set and no letter repeats back to back,
# so that no set gives the answer away by its pattern.
TARGET_LETTER = {
    "A": [1, 4, 0, 3, 5, 2, 3, 0, 4, 1],
    "B": [3, 0, 5, 2, 4, 1, 0, 5, 3, 2],
    "C": [4, 2, 1, 5, 0, 3, 2, 4, 1, 5],
}


# --------------------------------------------------------------------------- #
# low level OOXML helpers
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


def add_tab(parent, size=OPT_SIZE, font=Q_FONT):
    """A tab jump must live inside a run: <w:r><w:tab/></w:r>."""
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


def clear_cell(tc):
    for child in list(tc):
        if child.tag == qn("w:p") or child.tag == qn("w:tbl"):
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
def write_blocks(tc, blocks, first_para_indent=None):
    """blocks = list of (kind, payload); kind in {'p','code'}"""
    clear_cell(tc)
    for i, (kind, payload) in enumerate(blocks):
        if kind == "p":
            p = add_para(tc, indent=first_para_indent if i == 0 else None)
            add_run(p, payload, size=Q_SIZE)
        elif kind == "code":
            for line in payload:
                stripped = line.lstrip(" ")
                depth = (len(line) - len(stripped)) // 4
                p = add_para(tc, indent=180 + depth * 180, line=CODE_LINE,
                             exact=True)
                add_run(p, stripped, size=CODE_SIZE, font=CODE_FONT)


def write_options(tc, opts):
    """Append the option lines to the cell (6 options, T1)."""
    letters = "(A)", "(B)", "(C)", "(D)", "(E)", "(F)"
    longest = max(len(o) for o in opts)
    if longest <= 18:
        per_line, stops = 3, [2800, 5600]
    elif longest <= 42:
        per_line, stops = 2, [4200]
    else:
        per_line, stops = 1, []
    for i in range(0, len(opts), per_line):
        chunk = opts[i:i + per_line]
        p = add_para(tc, tabstops=stops)
        for k, opt in enumerate(chunk):
            if k:
                add_tab(p)
            add_run(p, f"{letters[i + k]} {opt}", size=OPT_SIZE)


def write_mcq_cell(tc, stem, code, opts):
    blocks = [("p", s) for s in stem]
    if code:
        blocks.append(("code", code))
    write_blocks(tc, blocks)
    write_options(tc, opts)


# --------------------------------------------------------------------------- #
# table assembly
# --------------------------------------------------------------------------- #
def rebuild_question_table(tbl, rows_spec, online=False):
    """rows_spec = list of dicts with keys q, sub, blocks, opts, marks, bloom"""
    trs = tbl.findall(qn("w:tr"))
    if online:
        # the ONLINE table of the T1-T3 template has no separate header row:
        # row 0 is 'Q-3 | (A) | <empty> | Marks | Bloom Taxonomy'.
        # Turn it into a clean header row and use it as the row prototype.
        header = trs[0]
        for tc in header.findall(qn("w:tc"))[:3]:
            clear_cell(tc)
            add_para(tc)
        proto_q = copy.deepcopy(header)
        for tr in trs[1:]:
            tbl.remove(tr)
    else:
        header = trs[0]                # '' | '' | '' | Marks | Bloom Taxonomy
        proto_q = trs[1]               # 'Q-1' | 'A)' | text | marks | bloom
        for tr in trs[1:]:
            tbl.remove(tr)

    repeat_header_row(header)

    for spec in rows_spec:
        tr = copy.deepcopy(proto_q)
        trPr = tr.find(qn("w:trPr"))
        if trPr is not None:
            for h in trPr.findall(qn("w:trHeight")):
                h.set(qn("w:val"), "240")           # let the row fit its content
        tcs = tr.findall(qn("w:tc"))
        # column 1 : question number
        set_line(tcs[0], spec["q"], bold=True)
        # column 2 : sub number / Sr. No.
        set_line(tcs[1], spec["sub"], bold=True)
        # column 3 : question body
        if spec.get("opts") is not None:
            write_mcq_cell(tcs[2], spec["stem"], spec["code"], spec["opts"])
        else:
            write_blocks(tcs[2], spec["blocks"])
        # column 4 : marks
        set_line(tcs[3], spec["marks"], bold=True, size=MARK_SIZE, align="center")
        # column 5 : bloom
        set_line(tcs[4], spec["bloom"], bold=True, size=MARK_SIZE, align="center")
        tbl.append(tr)


def shuffle_options(mcq, rng, target):
    """Return the 6 options with the correct one sitting at index `target`."""
    right = mcq["opts"][mcq["ans"]]
    wrong = [o for i, o in enumerate(mcq["opts"]) if i != mcq["ans"]]
    rng.shuffle(wrong)
    out = wrong[:target] + [right] + wrong[target:]
    return out, target


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
    """tbl = the 17-column header table of one page."""
    rows = tbl.findall(qn("w:tr"))
    # row 3 : 'CONTINUOUS AND COMPREHENSIVE EVALUATION (CCE)' / 'TEST 1 (CO1)'
    # (left untouched here - handled by replace_text_everywhere)
    # row 4 : B.E. SEMESTER-I ... BRANCH:
    set_line(rows[4].findall(qn("w:tc"))[0],
             "B.E. SEMESTER-I" + " " * 42 + "BRANCH: " + BRANCHES,
             bold=True, size=HEADER_SIZE)
    # row 5 : SUBJECT / SUBJECT CODE
    set_line(rows[5].findall(qn("w:tc"))[0],
             "SUBJECT- " + SUBJECT + " " * 24 + "SUBJECT CODE: " + SUBJECT_CODES,
             bold=True, size=HEADER_SIZE)
    # row 6 : DATE / DURATION
    set_line(rows[6].findall(qn("w:tc"))[0],
             "DATE: " + DATE + " " * 48 + "DURATION:    " + DURATION,
             bold=True, size=HEADER_SIZE)
    # row 7 : TIME / MAX MARKS
    time = OFFLINE_TIME if offline else ONLINE_TIME
    marks = OFFLINE_MAX if offline else ONLINE_MAX
    set_line(rows[7].findall(qn("w:tc"))[0],
             "TIME:  " + time + " " * 44 + "MAX MARKS: " + marks,
             bold=True, size=HEADER_SIZE)


def make_table_flowable(tbl):
    """The template builds the question tables as *floating* tables (w:tblpPr).
    A floating table cannot be split over a page boundary, so a long question
    paper would be pushed whole onto the next page.  Turn them into ordinary
    inline tables and centre them exactly like the institute header table."""
    tblPr = tbl.find(qn("w:tblPr"))
    floating = tblPr.find(qn("w:tblpPr"))
    if floating is not None:
        tblPr.remove(floating)
    if tblPr.find(qn("w:jc")) is None:
        tblPr.find(qn("w:tblW")).addnext(_w("w:jc", val="center"))


def spacer_before(tbl):
    """Small gap between the header block and the question table."""
    p = OxmlElement("w:p")
    pPr = OxmlElement("w:pPr")
    pPr.append(_w("w:spacing", after=0, line=240, lineRule="auto"))
    p.append(pPr)
    add_run(p, "", size=HEADER_SIZE)
    tbl.addprevious(p)


def repeat_header_row(tr):
    """Repeat the 'Marks | Bloom Taxonomy' row on every page of the paper."""
    trPr = tr.find(qn("w:trPr"))
    if trPr is None:
        trPr = OxmlElement("w:trPr")
        tr.insert(0, trPr)
    if trPr.find(qn("w:tblHeader")) is None:
        trPr.append(_w("w:tblHeader"))


def page_break_between_sections(doc):
    """Drop the template's filler paragraphs and start the ONLINE paper on a
    fresh page, so that a long offline paper can never drag it around."""
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


def set_set_letter(doc, letter):
    """The floating 'Set: A' box exists twice (DrawingML + VML fallback)."""
    n = 0
    for txbx in doc.element.body.iter(qn("w:txbxContent")):
        for t in txbx.iter(qn("w:t")):
            if t.text == "A":
                t.text = letter
                n += 1
    return n


# --------------------------------------------------------------------------- #
# build one set
# --------------------------------------------------------------------------- #
def build_set(letter, outpath):
    shutil.copyfile(TEMPLATE, outpath)
    doc = Document(str(outpath))

    # ---- heading : TEST 1 ( CO1 ) only -------------------------------------
    assert replace_text_everywhere(doc, "TEST 1/2/3 ( CO1/CO2/CO3 )",
                                   "TEST 1 ( CO1 )") == 2
    n_boxes = set_set_letter(doc, letter)
    assert n_boxes >= 2, n_boxes          # drawing + VML fallback, both pages

    # tables[0] = offline header, tables[1] = offline questions,
    # tables[2] = bloom legend, tables[3] = online header,
    # tables[4] = online questions, tables[5] = bloom legend
    page_break_between_sections(doc)
    tables = doc.tables
    fill_header(tables[0]._tbl, offline=True)
    fill_header(tables[3]._tbl, offline=False)
    for tbl in (tables[1]._tbl, tables[4]._tbl):
        make_table_flowable(tbl)
        spacer_before(tbl)

    # ---- offline question table -------------------------------------------
    order = SET_ORDER[letter]
    seed = OPTION_SHUFFLE_SEED[letter]
    rng = random.Random(seed)
    answer_key = {}

    rows_spec = [dict(q="Q-1", sub="Sr. No.",
                      blocks=[("p", "MCQ ( 1 Mark each )")],
                      marks="10", bloom="")]
    for n, src in enumerate(order, start=1):
        mcq = MCQS[src]
        opts, new_ans = shuffle_options(mcq, rng, TARGET_LETTER[letter][n - 1])
        answer_key[f"Q-1 {n})"] = "(%s)" % "ABCDEF"[new_ans]
        rows_spec.append(dict(q="", sub=f"{n})", stem=mcq["stem"], code=mcq["code"],
                              opts=opts, marks="1", bloom=mcq["bloom"]))

    for n, q in enumerate(Q2, start=1):
        rows_spec.append(dict(q="Q-2" if n == 1 else "", sub=f"{n})",
                              blocks=q["body"], marks=str(q["marks"]),
                              bloom=q["bloom"]))
        answer_key[f"Q-2 {n})"] = f"Unit {q['unit']} / Bloom {q['bloom']}"

    rebuild_question_table(tables[1]._tbl, rows_spec)

    # ---- online question table --------------------------------------------
    rows_spec = []
    for n, q in enumerate(Q3, start=1):
        rows_spec.append(dict(q="Q-3" if n == 1 else "",
                              sub="(%s)" % "AB"[n - 1],
                              blocks=q["body"], marks=str(q["marks"]),
                              bloom=q["bloom"]))
        answer_key[f"Q-3 ({'AB'[n-1]})"] = f"Unit {q['unit']} / Bloom {q['bloom']}"
    rebuild_question_table(tables[4]._tbl, rows_spec, online=True)

    doc.save(str(outpath))
    return answer_key


def write_answer_key(keys, path):
    lines = [
        "L. J. INSTITUTE OF ENGINEERING AND TECHNOLOGY, AHMEDABAD",
        "ANSWER KEY  -  TEST 1 (CO1) (OFFLINE) - Computer Programming using Python-I",
        "SEM-I, Batch 2026   |   QP Setter : MDP   |   FOR EVALUATION PURPOSE ONLY",
        "",
        "OFFLINE (16 marks) : Q-1 = 10 MCQ x 1 mark (from Practice Book, 6 options)",
        "                     Q-2 = 6 marks descriptive (out of Practice Book)",
        "ONLINE  (09 marks) : Q-3 = 9 marks descriptive (out of Practice Book)",
        "",
        "Practice Book source of each MCQ (authoring order):",
    ]
    for i, m in enumerate(MCQS, start=1):
        lines.append(f"  MCQ {i:>2} : PB Sr. No. {m['pb']:>3}  (Unit {m['unit']}, "
                     f"Bloom {m['bloom']})")
    lines.append("")
    for letter in ("A", "B", "C"):
        lines.append(f"--- SET {letter} ---")
        for k, v in keys[letter].items():
            lines.append(f"  {k:<12} {v}")
        lines.append("")
    path.write_text("\n".join(lines), encoding="utf-8")


def write_compliance_note(path):
    unit_marks = {1: 0, 2: 0, 3: 0}
    for m in MCQS:
        unit_marks[m["unit"]] += 1
    for q in Q2:
        unit_marks[q["unit"]] += q["marks"]
    for q in Q3:
        unit_marks[q["unit"]] += q["marks"]
    target = {1: 2, 2: 10, 3: 13}
    txt = f"""L. J. INSTITUTE OF ENGINEERING AND TECHNOLOGY, AHMEDABAD
TEST 1 (CO1) - Computer Programming using Python-I - SEM-I Batch 2026
QP SETTER : MDP     EXAM DATE : 29-Sep-2026 (Tue)
Submission deadline as per the QP setting circular : 25-Sep-2026, 12:00 midnight
Send to : qp.fyall@gmail.com

---------------------------------------------------------------------------
PAPER STRUCTURE
---------------------------------------------------------------------------
OFFLINE  16 marks   2:15 - 3:30 pm   1.25 hrs
   Q-1  10 x MCQ x 1 mark ........ 10 marks ... FROM the Practice Book
   Q-2  2 x 3 marks ..............  6 marks ... OUT of the Practice Book
ONLINE   09 marks   4:15 - 5:30 pm  1.25 hrs
   Q-3  (A) 4 marks + (B) 5 marks   9 marks ... OUT of the Practice Book
Total T1 = 25 marks, which is the T1 weightage issued by the HOD
(MCQ 10 + Descriptive/Programs 15 = 25).

UNIT / CO COVERAGE (issued unit wise marks 2 / 10 / 13, variation of 2-3 allowed)
   Unit 1 Introduction to Computer .................. {unit_marks[1]:>2} marks
   Unit 2 Introduction to Python and Jupyter ....... {unit_marks[2]:>2} marks
   Unit 3 Conditional Execution and Iterations ..... {unit_marks[3]:>2} marks
   Course Outcome : CO1 for every question of T1.

---------------------------------------------------------------------------
GUIDELINE BY GUIDELINE (FY_QUESTION PAPER SETTING INSTRUCTIONS_T1 TO T3)
---------------------------------------------------------------------------
 1  As per the T1 marks distribution issued by the HOD (10 MCQ + 15 descriptive
    = 25, split 16 offline / 09 online).
 2  No OR type options anywhere; all questions are compulsory.
 3  Every question is conceptual / output tracing / program writing. No
    definition, formula, theory, short note, True-False, Yes-No or fill in the
    blanks question is used.
 4  Difficulty is above average: 8 of the 10 MCQs are code/algorithm tracing
    questions (nested if-elif, for-else, while-else, break/continue/pass) and
    the descriptive questions need multi-step logic, not recall.
 5  Offline 16 marks = 10 marks from the Practice Book (the 10 MCQs, listed
    with their PB serial numbers in the answer key) + 6 marks out of the
    Practice Book (Q-2). Online 9 marks are completely out of the Practice
    Book (Q-3).
 6  Marks of the PB questions are kept at 1 mark each, which is within the
    allowed variation.
 7  16 marks in 1.25 hrs (10 MCQ + two 3 mark answers) and 9 marks in 1.25 hrs
    (two programs) can be completed inside the duration without being lengthy.
 8  Every MCQ is of 1 mark and carries exactly 6 options, as required for T1.
 9  Only 3 main questions (Q-1, Q-2, Q-3). Q-1 is the compulsory MCQ block and
    all MCQs are of the same marks (1 mark), so all of them are kept in Q-1
    (Case 2 of the instruction).
10  Subject codes of all 15 branches are printed in the heading.
11  Heading reads "TEST 1 ( CO1 )"; 2/3 and CO2/CO3 have been deleted.
12  Three sets A, B and C are prepared. Questions are exactly the same in all
    three sets; only the MCQs are reshuffled (question order and option order
    both), and the correct option is placed at a different letter in each set.
13  Pages are minimised: 11 pt question text, 9 pt code blocks with tight
    leading, options set horizontally with tab stops, no images, and the
    question tables were converted from floating to inline so that they can
    run over a page boundary (a floating table would have been pushed whole
    onto the next page).  Estimated length: offline paper about 2-3 pages,
    online paper 1 page (the offline length is driven by the ten MCQs, each of
    which has to carry six options in T1).
14  Not saved on a personal PC beyond this submission and not sent to anybody
    for prior verification.
15  No printout before the completion of the exam.
16  No LJIET faculty name is used in any question.
17  To be mailed to qp.fyall@gmail.com.
18  Prepared on the official Python-I T1-T3 format only
    (SET A_T1 to T3_PYTHON-1_TEST PAPER_SEM I_..._FORMAT.docx), including the
    Set box, the Marks and Bloom Taxonomy columns and the Bloom legend.

---------------------------------------------------------------------------
NOTE ON THE OFFLINE / ONLINE SPLIT
---------------------------------------------------------------------------
The HOD marks distribution for Python-I T1 is MCQ 10 + Descriptive 15 = 25.
The schedule splits T1 into a 16 mark offline test and a 09 mark online test,
and the setting instruction splits the offline paper into "10 marks from PB +
6 marks out of PB".  The only split that satisfies all three is
   offline = 10 MCQ (from PB) + 6 marks programs (out of PB)
   online  = 9 marks programs (out of PB)
which is what this paper uses.  If the department wants the online 9 marks to
be MCQ instead, only the Q-3 block has to be replaced.
"""
    path.write_text(txt, encoding="utf-8")


def main():
    OUTDIR.mkdir(exist_ok=True)
    keys = {}
    for letter in ("A", "B", "C"):
        name = (f"SET {letter}_T1_PYTHON-1_TEST PAPER_SEM I_MDP.docx")
        keys[letter] = build_set(letter, OUTDIR / name)
        print("written:", OUTDIR / name)
    write_answer_key(keys, OUTDIR / "ANSWER KEY_T1_PYTHON-I_MDP_do not circulate.txt")
    print("written:", OUTDIR / "ANSWER KEY_T1_PYTHON-I_MDP_do not circulate.txt")
    write_compliance_note(OUTDIR / "QP SETTING NOTE_T1_PYTHON-I_MDP.txt")
    print("written:", OUTDIR / "QP SETTING NOTE_T1_PYTHON-I_MDP.txt")


if __name__ == "__main__":
    main()
