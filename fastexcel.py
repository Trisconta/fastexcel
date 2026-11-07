#!/usr/bin/env python3
#-*- coding: utf-8 -*-

""" fastexcel.py script allows you to learn Excel openpyxl usage

(c)2020..2026  Henrique Moreira
"""

# pylint: disable=consider-using-f-string, too-many-arguments

### Notes:
###	unnecessary-comprehension, see:
###		https://github.com/PyCQA/pylint/issues/3164

import sys
import warnings
from openpyxl import load_workbook
import europattern

__version__ = "1.00 51"

DATA_ONLY = True

_SEP_CELL = "; "
_DEBUG = 0


def main():
    """ Main script """
    code = reader(sys.stdout, sys.stderr, sys.argv[1:])
    if code is None:
        print(f"""fastexcel.py Excel-file [options] file.xlsx [sheet [...]]

fastexcel script allows you to quickly learn how to use openpyxl,
by dumping the Excel 'active-sheet' (or the corresponding sheet) into text.

Options are:
	-s	Skip column heading
	-m	Show formulas

Version: {__version__}
""")
        code = 0
    sys.exit(code)


def reader(out, err, args):
    """ Reader function """
    debug = _DEBUG
    opts = {
        "skip": False,
        "data-only": DATA_ONLY,
    }
    if not args:
        return None
    param = args
    while param and param[0].startswith("-"):
        if param[0] in ("-s", "--skip"):
            opts["skip"] = True
            del param[0]
            continue
        if param[0] in ("-m", "--formulas"):
            opts["data-only"] = False
            del param[0]
            continue
        return None
    path = param[0]
    del param[0]
    if param:
        sheets = param
    else:
        sheets = []
    active = sheets == []
    suite = xcel_reader(path, sheets, active, data_only=opts["data-only"])
    if suite is None:
        err.write(f"Uops, cannot read: {path}\n")
    dump_xcel(out, suite, opts, debug)
    return 0


def xcel_reader(path, sheets=None, active=True, data_only=True):
    """ Excel reader function """
    assert isinstance(path, str), f"Wrong path, not a string: {[path]}"
    suite = {
        "path": path,
        "wb": None,
    }
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        wbk = wrap_load_book(path, data_only=data_only)
        suite["wb"] = wbk
    sheet_names = wbk.sheetnames
    ws_active = wbk.active
    if active:
        ws_sel = [ws_active,]
    else:
        ws_sel = [wbk.get_sheet_by_name(sheets[0])]
    assert suite["wb"], "Null!"
    suite["req-sheets"] = sheets
    suite["sheet-names"] = sheet_names
    suite["active"] = ws_active
    suite["ws-sel"] = ws_sel
    return suite


def dump_xcel(out, suite, opts, debug=0):
    """ Dump one Excel sheet """
    #sheets = suite["req-sheets"]
    sheet = suite["ws-sel"][0]
    cont, info = dump_one(suite, sheet, out, opts, debug=debug)
    (title, _), (max_col, max_row), _ = info
    if debug > 0:
        print("Debug: '{}': cols={}, rows={}, {}"
              "".format(title, max_col, max_row, cont if cont else "(empty)"))
    return 0


def dump_one(suite, sheet, out=None, opts=None, y_start=1, debug=0):
    """ Dump sheet content to stdout (or another text stream) """
    output = out if out is not None else sys.stdout
    ws = sheet
    skip = opts.get("skip", False)
    for idx, row in enumerate(ws.rows, y_start):
        s = "" if skip else "{}#\t".format(idx)
        w = ""
        for cell in row:
            if w:
                w += _SEP_CELL
            if debug > 0:
                val = "({}.{}='{}')".format(cell.coordinate, shown_cell_type(cell), cell_repr(cell))
            else:
                val = cell_repr(cell)
            w += val
        if output:
            output.write("{}{}\n".format(s, w))
    return (
        (ws.max_row,),
        (
            (ws.title, ws.path),
            (ws.max_column, ws.max_row),
            ws.sheet_properties,
        )
    )


def cell_repr(cell, when_null="---"):
    """ Cell representation """
    assert isinstance(when_null, str), "when_null invalid"
    s = None
    if cell.value is None:
        s = when_null
    elif cell.data_type == "n":
        num_fmt = cell.number_format
        spl = num_fmt.split(".")
        if num_fmt.startswith(("G", "General")):
            s = _generalis_format(cell, num_fmt)
        if len(spl) > 1:
            zeros = spl[-1].count("0")
        else:
            zeros = 0
        if s is None:
            fmt = "{:0."+str(zeros)+"f}"
            s = fmt.format(cell.value)
        # Debug:	s = repr([num_fmt, cell.value, s, spl, zeros])
        #s = repr([num_fmt, cell.value, s, spl, zeros])
    elif cell.data_type == "s":
        s = mystring(cell.value)
    else:
        s = "{}".format(cell.value)
        if s.endswith(" 00:00:00"):
            s = s[:-len(" 00:00:00")]
    return s

def _generalis_format(cell, num_fmt):
    if not isinstance(cell.value, float):
        return str(cell.value)
    if cell.value.is_integer():
        s = str(int(cell.value))
    else:
        s = f"{cell.value:g}"
    return s


def mystring(astr, do_normalize=True):
    """ Simple string conversion. """
    res = astr.replace("\t", "\\t")
    res = res.replace("\u20AC", "")
    if do_normalize:
        pat = europattern.normalize_number(res)
    else:
        pat = None
    if pat is None:
        return string_7bit(res)
    return pat

def string_7bit(astr: str, alt: str="."):
    """ Returns dots instead of chars ASCII above or equal to 127d. """
    assert isinstance(astr, str), f"Bad string_7bit(): {[astr]}"
    myset = [
        aba if aba <= '~' else alt for aba in astr
        if aba in ("\t", "\n") or aba >= ' '
    ]
    return ''.join(myset)


def shown_cell_type(cell):
    """ Return a single letter for the cell data_type.
        's' means string.
    """
    # [openpyxl.cell package](https://openpyxl.readthedocs.io/en/2.0/openpyxl.cell.html)
    # openpyxl.cell.cell.VALID_TYPES = ('s', 'f', 'n', 'b', 'n', 'inlineStr', 'e', 'str')
    # openpyxl.cell.cell.TYPE_BOOL = 'b'
    # openpyxl.cell.cell.TYPE_STRING = 's'
    #  ...TYPE_NULL = 'n' ? (TYPE_NUMERIC too!)
    #  ...TYPE_FORMULA = 'f'
    s = cell.data_type
    if cell.is_date:
        letter = "D"
    else:
        letter = "c" if s == "str" else s[0]
    return letter


def wrap_load_book(path: str, data_only=True):
    """ Wrapper for load_workbook """
    # The following breaks at 'ws.path'!
    #wbk = load_workbook(path, data_only=True, read_only=True)
    # Loading normally:
    wbk = load_workbook(path, data_only=data_only)
    return wbk


#
# Main script
#
if __name__ == '__main__':
    main()
