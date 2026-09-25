#!/usr/bin/env python3
"""
Build the 9-mark practical problem statement for T1 (CO1) Python-I as a
standalone Word file, laid out on the ONLINE page of the official Python-I
T1-T3 template so that it can be issued as-is or dropped into the set.

Scope is restricted to Units 1-3 of the syllabus:
  Unit 1  1.3 Algorithm and Flowchart
  Unit 2  2.1 data types / typecasting, 2.2 variables, input and output,
          2.3 arithmetic, comparison and logical operators, precedence
  Unit 3  3.1 if / if-else / if-elif-else / nested if,
          3.2 for and while loops, nested loops, 3.3 break, continue, pass
Nothing from Unit 4 (functions) or later is used or permitted.
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from docx import Document                                       # noqa: E402
from build_t1_qp import (TEMPLATE, fill_header, make_table_flowable,  # noqa: E402
                         mark, page_break_between_sections,
                         rebuild_question_table, replace_text_everywhere,
                         set_set_letter, spacer_before, split_pages)

BODY = [
    ("p", "A college campus has an automated parking system that charges every "
          "vehicle when it leaves the campus."),
    ("p", "First write an algorithm and draw a flowchart for the system "
          "described below, and then write a Python program for the same "
          "system."),
    ("p", "System specification : the program processes exactly eight "
          "vehicles, one after another, using a loop. For every vehicle it "
          "reads the vehicle number, the type of vehicle (1 for two-wheeler, "
          "2 for four-wheeler), the entry time as hour and minute and the "
          "exit time as hour and minute, all on a 24-hour clock. The program "
          "must"),
    ("p", "a)     validate the type of vehicle : if the value entered is "
          "neither 1 nor 2, print INVALID TYPE and read the type again until "
          "it is 1 or 2;"),
    ("p", "b)     convert the entry time and the exit time into the number of "
          "minutes from midnight and find the parked minutes as exit time "
          "minus entry time; if the parked minutes are negative, print "
          "INVALID TIME and read the entry and exit times again, repeating "
          "this until a valid pair is entered;"),
    ("p", "c)     the first 10 minutes of parking are free : if the parked "
          "minutes are 10 or less, the billable time is zero; otherwise "
          "compute the billable hours by rounding the parked minutes UP to a "
          "whole hour, using integer arithmetic only;"),
    ("p", "d)     compute the parking fee : for a two-wheeler the first 3 "
          "billable hours are charged at Rs. 10 per hour and every hour after "
          "that at Rs. 15 per hour; for a four-wheeler the first 3 billable "
          "hours are charged at Rs. 20 per hour and every hour after that at "
          "Rs. 30 per hour;"),
    ("p", "e)     if the parked minutes are more than 480, add an overstay "
          "penalty of Rs. 50;"),
    ("p", "f)     compute the total fee as parking fee plus penalty, and print "
          "the vehicle number, the parked minutes, the billable hours, the "
          "parking fee, the penalty and the total fee for that vehicle;"),
    ("p", "g)     print LONG STAY for a vehicle whose parked minutes are more "
          "than 480 and NORMAL for every other vehicle;"),
    ("p", "h)     after all eight vehicles have been processed, print the "
          "number of two-wheelers, the number of four-wheelers, the number of "
          "LONG STAY vehicles, the total fee collected from all eight "
          "vehicles, and the highest total fee together with the vehicle "
          "number that paid it."),
    ("p", "Use of functions, lists, tuples, dictionaries, sets, files or "
          "modules is not allowed. Do not use the ternary (conditional) "
          "operator form."),
]


def build(outpath):
    doc = Document(str(TEMPLATE))
    assert replace_text_everywhere(doc, "TEST 1/2/3 ( CO1/CO2/CO3 )",
                                   "TEST 1 ( CO1 )") == 2
    page_break_between_sections(doc)
    tables = doc.tables
    fill_header(tables[3]._tbl, offline=False)
    tbl = tables[4]._tbl
    make_table_flowable(tbl)
    spacer_before(tbl)
    rebuild_question_table(tbl, [dict(q="Q-3", sub="(A)", blocks=BODY,
                                      marks=mark(9), bloom="C")], online=True)
    split_pages(doc, "online")
    set_set_letter(doc, "1", current="A")
    doc.save(str(outpath))
    return outpath


if __name__ == "__main__":
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(
        "Q-3(A) PRACTICAL PROBLEM_9 MARKS_T1_PYTHON-I_MDP.docx")
    print("written:", build(out))
