#!/usr/bin/env python3
"""
Regenerate the faculty-only note for the 9-mark practical problem statement.

The reference solution and the sample input are embedded here so the note is
reproducible: the sample output printed into the note is produced by actually
executing the reference solution every time this script runs.
"""

import subprocess
import sys
from pathlib import Path

REFERENCE = '''two = 0
four = 0
longstay = 0
collected = 0
maxfee = -1
maxno = 0
for i in range(1, 9):
    no = int(input())
    vt = int(input())
    while vt != 1 and vt != 2:
        print("INVALID TYPE")
        vt = int(input())
    eh = int(input()); em = int(input())
    xh = int(input()); xm = int(input())
    parked = (xh * 60 + xm) - (eh * 60 + em)
    while parked < 0:
        print("INVALID TIME")
        eh = int(input()); em = int(input())
        xh = int(input()); xm = int(input())
        parked = (xh * 60 + xm) - (eh * 60 + em)
    if parked <= 10:
        hours = 0
    else:
        hours = (parked + 59) // 60
    if vt == 1:
        if hours <= 3:
            fee = hours * 10
        else:
            fee = 3 * 10 + (hours - 3) * 15
        two = two + 1
    else:
        if hours <= 3:
            fee = hours * 20
        else:
            fee = 3 * 20 + (hours - 3) * 30
        four = four + 1
    penalty = 0
    if parked > 480:
        penalty = 50
        longstay = longstay + 1
    total = fee + penalty
    collected = collected + total
    if total > maxfee:
        maxfee = total
        maxno = no
    if parked > 480:
        print(no, parked, hours, fee, penalty, total, "LONG STAY")
    else:
        print(no, parked, hours, fee, penalty, total, "NORMAL")
print("two-wheelers", two, "four-wheelers", four, "long stays", longstay)
print("collected", collected)
print("highest", maxfee, maxno)
'''

SAMPLE = [
    [101, 1, 9, 45, 11, 20],
    [102, 2, 8, 0, 18, 30],
    [103, 1, 10, 0, 10, 8],
    [104, 3, 2, 0, 0, 12, 5],
    [105, 2, 23, 30, 1, 15, 13, 14, 9, 0, 13, 14, 23, 13],
    [106, 1, 7, 0, 7, 5],
    [107, 2, 6, 15, 9, 45],
    [108, 1, 14, 0, 23, 59],
]

NOTE = """FACULTY NOTE - do not circulate to students
Q-3(A) PRACTICAL PROBLEM STATEMENT, 9 MARKS - T1 (CO1) Computer Programming using Python-I
QP setter MDP | online test 29-Sep-2026, 4:15 to 5:30 pm

--------------------------------------------------------------------------
1. SYLLABUS MAPPING - every clause uses Units 1-3 only
--------------------------------------------------------------------------
clause   requirement                          syllabus clause used
------   ---------------------------------    ------------------------------
intro    write an algorithm                   Unit 1  1.3 Algorithm and Flowchart
intro    draw a flowchart                     Unit 1  1.3 Algorithm and Flowchart
spec     read vehicle number, type, times     Unit 2  2.2 reading input from users
spec     all values are integers              Unit 2  2.1 int type, typecasting
a)       type must be 1 or 2, repeat          Unit 3  3.1 if / while validation loop
b)       hour*60+minute, exit minus entry     Unit 2  2.3 arithmetic *, +, -
b)       reject negative parked minutes       Unit 2  2.3 comparison <
b)       repeat until a valid pair            Unit 3  3.2 while loop
c)       parked <= 10 gives zero hours        Unit 3  3.1 if-else
c)       round UP with (parked + 59) // 60    Unit 2  2.3 // and %, operator precedence
d)       first 3 hours, rest at another rate  Unit 3  3.1 if-else, nested if
d)       3*10 + (hours-3)*15 style charging   Unit 2  2.3 arithmetic, precedence
e)       parked > 480 adds Rs. 50             Unit 2  2.3 comparison; Unit 3 3.1 if
f)       total = fee + penalty, then print    Unit 2  2.2 printing output
g)       LONG STAY / NORMAL                   Unit 3  3.1 if-else
h)       counters, totals, highest fee        Unit 3  3.2 for loop, counters
h)       track the highest fee and its number Unit 3  3.1 nested if, Unit 2 2.3 >
overall  eight vehicles in one loop           Unit 3  3.2 for using range

NOTHING from Unit 4 (functions, scope, recursion) or later units is required,
and the question forbids functions, collections, files, modules and the
ternary form, so a student cannot drift outside Units 1-3.

--------------------------------------------------------------------------
2. SUGGESTED MARKING SCHEME (total 9)
--------------------------------------------------------------------------
  algorithm, correct sequence of steps ..................... 2
  flowchart, correct symbols and decision branches ......... 1
  input, type validation loop and time validation loop ..... 2
  minutes conversion and rounding up to a whole hour ....... 1
  fee slabs, penalty and total ............................. 2
  counters, totals and highest fee with vehicle number ..... 1
The student paper shows only [09]; the split above is for evaluation only.

--------------------------------------------------------------------------
3. REFERENCE SOLUTION - executed, not hand-written
--------------------------------------------------------------------------
"""

TAIL = """
--------------------------------------------------------------------------
5. NOTES
--------------------------------------------------------------------------
- This is an alternative to the electricity-billing question currently in
  ONLINE (COMMON)_T1_PYTHON-1_TEST PAPER_SEM I_MDP.docx. It is entirely
  outside the Practice Book, Bloom level C, and is one integrated question
  under Q-3(A) as the guidelines require.
- Unit attribution for the blueprint: the required algorithm and flowchart
  connect Unit 1 as the guidelines ask, but no marks are separately allotted
  to Unit 1, so the online paper stays at Unit 1 = 0, Unit 2 = 4, Unit 3 = 5
  and the T1 total stays 2 / 10 / 13.
- OUTSTANDING: a PDF preview must be produced and every page visually checked
  in Word before this file is called final.
"""


def indent(text, pad="    "):
    return "\n".join(pad + line if line.strip() else line
                     for line in text.splitlines()) + "\n"


def main(outpath):
    stdin = "".join(str(v) + "\n" for row in SAMPLE for v in row)
    run = subprocess.run([sys.executable, "-c", REFERENCE], input=stdin,
                         capture_output=True, text=True, check=True)
    text = NOTE + indent(REFERENCE)
    text += ("\n--------------------------------------------------------------------"
             "------\n4. VERIFIED SAMPLE RUN (produced by executing the "
             "reference solution above\n   at the moment this note was "
             "generated)\n----------------------------------------------"
             "----------------------\ninput, one value per line: vehicle "
             "number, type, entry hour, entry minute,\nexit hour, exit "
             "minute - plus a deliberately invalid type (104) and two "
             "invalid\ntime pairs (105) to exercise the validation loops:\n\n"
             "  101 1  9 45 11 20 | 102 2  8 0 18 30 | 103 1 10  0 10  8\n"
             "  104 3 -> 2  0  0 12  5   (3 is invalid, so the type is read "
             "again)\n  105 2 23 30  1 15 -> 13 14 9 0 -> 13 14 23 13   "
             "(two invalid pairs)\n  106 1  7  0  7  5 | 107 2  6 15  9 45 "
             "| 108 1 14  0 23 59\n\noutput:\n\n" + indent(run.stdout))
    text += ("  worked checks: 101 -> 95 min, (95+59)//60 = 2 h, 2 x 10 = "
             "Rs. 20\n                 102 -> 630 min, 11 h, 3x20 + 8x30 = "
             "Rs. 300, +50 = Rs. 350\n                 103 -> 8 min, free, "
             "Rs. 0\n                 104 -> 725 min, 13 h, 3x20 + 10x30 = "
             "Rs. 360, +50 = Rs. 410\n                 108 -> 599 min, 10 h, "
             "3x10 + 7x15 = Rs. 135, +50 = Rs. 185\n")
    text += TAIL
    Path(outpath).write_text(text)
    return run.stdout


if __name__ == "__main__":
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(
        "Q-3(A) PRACTICAL PROBLEM_FACULTY NOTE_MDP.txt")
    printed = main(out)
    print("written:", out)
    print(printed)
