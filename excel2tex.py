import argparse
from openpyxl import load_workbook

class Cell:
    def __init__(self, text, cell_type, parameters):
        self.text = text
        self.cell_type = cell_type
        self.parameters = parameters
        if parameters.get("height") is None:
            parameters["height"] = 1
        if parameters.get("width") is None:
            parameters["width"] = 1

    def output(self):
        begin_line = '|' if self.parameters['begin'] else ''
        str_dic = {
                'plain': self.text,
                'multirowcell_begin': '\\multirowcell{' + str(self.parameters["height"]) + '}{' + self.text + '}',
                'multirowcell_other': '',
#                  'multicolumn_begin': '\\multicolumn{' + str(self.parameters["width"]) + '}{' + '|' if self.parameters['begin'] else '' + 'c|}{' + self.text + '}',
                'multicolumn_begin': '\\multicolumn{' + str(self.parameters["width"]) + '}{' + begin_line + 'c|}{' + self.text + '}',
                'block_firstline_begin': '\\multicolumn{' + str(self.parameters["width"]) + '}{' + begin_line + 'c|}{' + '\\multirowcell{' + str(self.parameters["height"]) + '}{' + self.text + '}' + '}',
                'block_placeholder_begin': '\\multicolumn{1}{|c}{}',
                'block_placeholder_end': '\\multicolumn{1}{c|}{}',
                'block_placeholder_other': '\\multicolumn{1}{c}{}'
                }
        res = str_dic[self.cell_type]
        if not self.parameters["begin"]:
            res = '& ' + res
        if self.parameters["end"]:
            res += r' \\'
        return res.strip()

class MergedCell:
    def __init__(self, bounds):
        self.y1, self.x1, self.y2, self.x2 = bounds

    def is_merged(self, row, col):
        return True if row >= self.x1 and row <= self.x2 and col >= self.y1 and col <= self.y2 else False

    def get_type(self, row, col):
        if self.x1 != self.x2 and self.y1 != self.y2:
            if row == self.x1:
                return "block_firstline_begin" if col == self.y1 else "block_firstline_other"
            if col == self.y1:
                return "block_placeholder_begin"
            return "block_placeholder_end" if col == self.y2 else "block_placeholder_other"
        if self.x1 == self.x2:
            return "multicolumn_begin" if col == self.y1 else "multicolumn_other"
        if self.y1 == self.y2:
            return "multirowcell_begin" if row == self.x1 else "multirowcell_other"

    def is_end(self, col_max):
        return self.y2 == col_max

class Table:
    def __init__(self, ws):
        self.ws = ws
        self.min_row = ws.min_row
        self.max_row = ws.max_row
        self.min_column = ws.min_column
        self.max_column = ws.max_column
        self.merged_cells = [MergedCell(cell.bounds) for cell in ws.merged_cells.ranges]
        self.cline_ranges = []
        self.cells = []
        self.tex = '% Please add the following required packages to your document preamble:\n% \\usepackage{multirow, makecell}\n'
        self.tex += r'\begin{tabular}{*{' + str(self.max_column - self.min_column + 1) + '}{|c}|}\n'
        self.tex += '\\hline\n'
        self.row_texs = []
        self.convert()
        self.tex += ''.join(self.row_texs)
        self.tex += '\\hline\n'
        self.tex += '\\end{tabular}'

    def is_empty(self, row, col):
        return self.ws.cell(row, col).value is not None or self.cells[row -
                self.ws.min_row][col - self.ws.min_column].parameters['merged_count']

    def set_bounds(self):
        # remove top empty rows
        i = self.ws.min_row
        while 1:
            empty = True
            for j in range(i, self.ws.max_column + 1):
                if self.is_empty(i, j):
                    empty = False
                    break
            if empty:
                self.min_row += 1
            else:
                break
            i += 1
        j = self.ws.min_column
        # remove bottom empty rows
        i = self.ws.max_row
        while 1:
            empty = True
            for j in range(self.ws.max_column, self.ws.min_column - 1, -1):
                if self.is_empty(i, j):
                    empty = False
                    break
            if empty:
                self.max_row -= 1
            else:
                break
            i -= 1
        # remove left empty cols
        j = self.ws.min_column
        while 1:
            empty = True
            for j in range(j, self.ws.max_row + 1):
                if self.is_empty(i, j):
                    empty = False
                    break
            if empty:
                self.min_column += 1
            else:
                break
            j += 1
        # remove right empty cols
        j = self.ws.max_column
        while 1:
            empty = True
            for i in range(self.ws.max_row, self.ws.min_row - 1, -1):
                if self.is_empty(i, j):
                    empty = False
                    break
            if empty:
                self.max_column -= 1
            else:
                break
            j -= 1

    def set_parameters(self):
        self.cells = []
        # check type
        for i in range(self.min_row, self.max_row + 1):
            row_cells = []
            row_cline = []
            cline_begin = self.min_column - 1
            cline_end = self.min_column - 1
            cline_start = True
            for j in range(self.min_column, self.max_column + 1):
                text = self.ws.cell(i, j).value
                if text is None:
                    text = ''
                if type(text) != str:
                    text = str(text)
                text = text.replace('\n', r' \\')
                cell_type = "plain"
                parameters = {}
                parameters["begin"] = True if j == self.min_column else False
                # real end or block end
                parameters["end"] = True if j == self.max_column else False
                parameters["merged_count"] = 0
                merged_count = 1
                for m in self.merged_cells:
                    if m.is_merged(i, j):
                        parameters["merged_count"] = merged_count
                        cell_type = m.get_type(i, j)
                        if cell_type == "multirowcell_begin":
                            parameters["height"] = m.x2 - m.x1 + 1
                        if cell_type == "multicolumn_begin":
                            parameters["width"] = m.y2 - m.y1 + 1
                        if cell_type == "block_firstline_begin":
                            parameters["height"] = m.x2 - m.x1 + 1
                            parameters["width"] = m.y2 - m.y1 + 1
                        parameters["end"] = m.is_end(self.max_column)
                        break
                    merged_count += 1
                cell = Cell(text, cell_type, parameters)
                row_cells.append(cell)
                # cline
                if i == self.min_row:
                    continue
                up_cell = self.cells[i - self.min_row - 1][j - self.min_column]
                if cell_type == 'plain' or parameters["merged_count"] != up_cell.parameters["merged_count"]:
                    if cline_start:
                        cline_begin = j
                        cline_start = False
                    cline_end = j
                    if j == self.max_column:
                        row_cline.append((cline_begin, cline_end))
                else:
                    if not cline_start:
                        row_cline.append((cline_begin, cline_end))
                        cline_start = True
            self.cells.append(row_cells)
            if i > self.min_row:
                self.cline_ranges.append(row_cline)

    def convert(self):
        self.set_parameters()
        # remove empty rows and cols
        self.set_bounds()
        # correct parameters
        self.set_parameters()
        # set output text
        n = 1
        print(self.min_row, self.max_row, self.min_column, self.max_column)
        print(len(self.cells))
        for i in range(self.min_row, self.max_row + 1):
#          for r in self.cells:
#              self.tex += f'% row {n}\n'
            row_tex = f'% row {n}\n'
#              for cell in r:
            for j in range(self.min_column, self.max_column + 1):
                cell = self.cells[i - self.min_row][j - self.min_column]
                if cell.cell_type != "multicolumn_other" and cell.cell_type != "block_firstline_other":
                    out_text = cell.output()
                    if out_text:
                        out_text = '  ' + out_text + '\n'
                    # cline
                    if n < self.max_row - self.min_row + 1 and cell.parameters['end']:
                        row_cline = self.cline_ranges[n - 1]
                        if len(row_cline) == 1 and row_cline[0][0] == self.min_column and row_cline[0][1] == self.max_column:
                            out_text += '\\hline'
                        else:
                            for cline_range in row_cline:
                                out_text += '\\cline{' + str(cline_range[0]) + '-' + str(cline_range[1]) + '}\n'
                    row_tex += out_text
            n += 1
            self.row_texs.append(row_tex)

if __name__ == '__main__':
    parser = argparse.ArgumentParser(
            description='Convert excel table to latex table.',
            formatter_class=argparse.RawTextHelpFormatter,
            epilog="""Note:
1. The height of every merged cell must be greater than the number of lines in your text.
2. Make sure that the height of every merged cell is even.
3. Make the size of the merged cell suitable, i.e., do not give more space than that you need.
"""
            )
    parser.add_argument('-s', default='table.xlsx', dest='source', help='source file (default: %(default)s)')
    parser.add_argument('-o', default='table.tex', dest='target', help='target file (default: %(default)s)')
    parser.add_argument('-e', default='utf-8', dest='encoding',
            choices=['utf-8', 'utf-8-sig'],
            help='file encoding (default: %(default)s), if there is mess code, set it to utf-8-sig')
    args = parser.parse_args()
    wb = load_workbook(args.source)
    ws = wb.active
    t = Table(ws)
    with open(args.target, 'w', encoding=args.encoding) as f:
        f.write(t.tex.replace('  &\n', '  &'))
