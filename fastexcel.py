# (c)2020  Henrique Moreira  -- fast Excel read-out

import sys
from openpyxl import load_workbook


_SEP_CELL = "; "


def main():
    """ Main script """
    reader(sys.stdout, sys.stderr, sys.argv[1:])

    
def reader(out, err, args):
    param = args
    path = param[0]
    del param[0]
    if param:
        sheets = param
    else:
        sheets = []
    active = sheets == []
    suite = xcel_reader(path, sheets, active)
    dump_xcel(out, suite)
    return 0


def xcel_reader(path, sheets=None, active=True):
    assert isinstance(path, str)
    suite = {"path": path}
    wb = load_workbook(path)
    if wb:
        sheet_names = wb.sheetnames
        ws_active = wb.active
        if active:
            ws_sel = [ws_active,]
        else:
            ws_sel = [wb.get_sheet_by_name(sheets[0])]
        suite["wb"] = wb
        suite["req-sheets"] = sheets
        suite["sheet-names"] = sheet_names
        suite["active"] = ws_active
        suite["ws-sel"] = ws_sel
    return suite


def dump_xcel(out, suite, debug=0):
    #sheets = suite["req-sheets"]
    cont, info = dump_one(out, suite, suite["ws-sel"][0], debug)
    (title, _), (max_col, max_row), _ = info
    if debug >= 0:
        print("Debug: '{}': cols={}, rows={}, {}"
              "".format(title, max_col, max_row, cont if cont else "(empty)"))
    return 0


def dump_one(out, suite, sheet, debug=0):
    ws = sheet
    rows = [row for row in ws.rows]
    info = ((ws.title, ws.path), (ws.max_column, ws.max_row), ws.sheet_properties,
            )
    if not rows:
        return tuple(), inf
    idx = 0
    for row in rows:
        idx += 1
        s = "{}#\t".format(idx)
        w = ""
        for cell in row:
            if w:
                w += _SEP_CELL
            if debug > 0:
                val = "({}.{}='{}')".format(cell.coordinate, shown_cell_type(cell), cell_repr(cell))
            else:
                val = cell_repr(cell)
            w += val
        out.write("{}{}\n".format(s, w))
    return (ws.max_row,), info


def cell_repr(cell, when_null="---"):
    assert isinstance(when_null, str)
    if cell.value is None:
        s = when_null
    elif cell.data_type == "n":
        num_fmt = cell.number_format
        spl = num_fmt.split(".")
        if len(spl) > 1:
            zeros = spl[-1].count("0")
        else:
            zeros = 0
        fmt = "{:0."+str(zeros)+"f}"
        s = fmt.format(cell.value)
    else:
        s = "{}".format(cell.value)
    return s


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


#
# Main script
#
if __name__ == '__main__':
    main()