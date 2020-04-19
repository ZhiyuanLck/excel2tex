from .Cell import Cell
from .MergedCell import MergedCell
from .help import scan_all
from .line import LineMatrix

class Table:
    def __init__(self, ws, args):
        self.ws = ws
        self.args = args
        self.x1 = 1
        self.x2 = ws.max_row + 3
        self.y1 = 1
        self.y2 = ws.max_column + 3
        self.colors = set()
        self.vspace = 0
        self.cells = [
                [Cell(self.ws.cell(i + 1, j + 1), self)
                for j in range(self.y2)]
                for i in range(self.x2)]
        self.merged_cells = [MergedCell(idx + 1, cell.bounds) for idx, cell in enumerate(ws.merged_cells.ranges)]
        self.init_cell_matrix()
        self.clines = LineMatrix(self, 'cline')
        self.hlines = LineMatrix(self, 'hline')
        self.set_bounds()
        self.clines.set_lines(self, 'cline')
        self.clines.set_vspace()
        self.hlines.set_lines(self, 'hline')
        if self.args.excel_format:
            self.set_line_bounds()
        self.set_props()
        self.set_clines()
        self.set_texs()

    def init_cell_matrix(self):
        for i in range(1, self.x2 + 1):
            for j in range(1, self.y2 + 1):
                cell = self.cells[i - 1][j - 1]
                r, c = cell.get_head(self.merged_cells, i, j)
                cell.head = self.cells[r - 1][c - 1]
                cell.merged_idx = cell.get_merged_idx(self.merged_cells, i, j)
                cell.text_prop.text = self.ws.cell(i, j).value
                cell.first_row = True if i == self.x1 else False
                cell.last_row = True if i == self.x2 else False
                cell.first_col = True if j == self.y1 else False
                cell.last_col = True if j == self.y2 else False
                # set border_prop
                cell.border.set_prop(cell)

    def set_bounds(self):
        # remove top empty rows
        old = self.x1
        for i in range(old - 1, self.x2):
            cell_list = self.cells[i]
            if scan_all(cell_list, 'not_empty'):
                self.x1 += 1
            else:
                break
        # remove bottom empty rows
        old = self.x2
        for i in range(old, self.x1 - 1, -1):
            cell_list = self.cells[i - 1]
            if scan_all(cell_list, 'not_empty'):
                self.x2 -= 1
            else:
                break
        # remove left empty cols
        old = self.y1
        for j in range(old - 1, self.y2):
            cell_list = [row[j] for row in self.cells]
            if scan_all(cell_list, 'not_empty'):
                self.y1 += 1
            else:
                break
        # remove right empty cols
        old = self.y2
        for j in range(old - 1, self.y1, -1):
            cell_list = [row[j] for row in self.cells]
            if scan_all(cell_list, 'not_empty'):
                self.y2 -= 1
            else:
                break

    def set_line_bounds(self):
        cline_transpose = list(zip(*self.clines.borders))
        hline_transpose = list(zip(*self.hlines.borders))
        # cline x
        cline_x_min = 1
        cline_x_max = len(self.clines.borders)
        # cell y
        cell_y_min = 1
        cell_y_max = len(cline_transpose)
        # cell x
        cell_x_min = 1
        cell_x_max = len(hline_transpose)
        # hline y
        hline_y_min = 1
        hline_y_max = len(self.hlines.borders)
        # cline remove empty row line
        for cline in self.clines.borders:
            if self.clines.is_empty(cline):
                cline_x_min += 1
            else:
                break
        for cline in self.clines.borders[::-1]:
            if self.clines.is_empty(cline):
                cline_x_max -= 1
            else:
                break
        # cline remove empty col cell
        for cline in cline_transpose:
            if self.clines.is_empty(cline):
                cell_y_min += 1
            else:
                break
        for cline in cline_transpose[::-1]:
            if self.clines.is_empty(cline):
                cell_y_max -= 1
            else:
                break
        # hline remove empty col line
        for hline in self.hlines.borders:
            if self.hlines.is_empty(hline):
                hline_y_min += 1
            else:
                break
        for hline in self.hlines.borders[::-1]:
            if self.hlines.is_empty(hline):
                hline_y_max -= 1
            else:
                break
        # hline remove empty row cell
        for hline in hline_transpose:
            if self.hlines.is_empty(hline):
                cell_x_min += 1
            else:
                break
        for hline in hline_transpose[::-1]:
            if self.hlines.is_empty(hline):
                cell_x_max -= 1
            else:
                break
        # combine line range with origin cell range
        if cline_x_min <= cline_x_max:
            x1 = cline_x_min
            x2 = cline_x_max if cline_x_min == cline_x_max else cline_x_max - 1
            self.x1 = min(self.x1, x1)
            self.x2 = max(self.x2, x2)
        if hline_y_min <= hline_y_max:
            y1 = hline_y_min
            y2 = hline_y_max if hline_y_min == hline_y_max else hline_y_max - 1
            self.y1 = min(self.y1, y1)
            self.y2 = max(self.y2, y2)
        # combine new cell range with the range befor
        self.x1 = min(self.x1, cell_x_min)
        self.x2 = max(self.x2, cell_x_max)
        self.y1 = min(self.y1, cell_y_min)
        self.y2 = max(self.y2, cell_y_max)

    def set_props(self):
        for i in range(self.x1, self.x2 + 1):
            for j in range(self.y1, self.y2 + 1):
                cell = self.cells[i - 1][j - 1]
                self.set_cell_type()
                cell.set_prop()

    def set_cell_type(self):
        '''
        set `begin`, `end`, `height`, `width`, `cell_type`
        '''
        for i in range(self.x1, self.x2 + 1):
            for j in range(self.y1, self.y2 + 1):
                cell = self.cells[i - 1][j - 1]
                cell.cell_type = "plain"
                cell.begin = True if j == self.y1 else False
                # real end or block end
                cell.end = True if j == self.y2 else False
                if cell.merged_idx:
                    merged_cell = self.merged_cells[cell.merged_idx - 1]
                    cell.cell_type = merged_cell.get_type(i, j)
                    cell.height = merged_cell.x2 - merged_cell.x1 + 1
                    cell.width = merged_cell.y2 - merged_cell.y1 + 1
                    cell.end = merged_cell.is_end(self.y2)

    def set_clines(self):
        self.cline_ranges = []
        for i in range(self.x1, self.x2 + 1):
            row_cline = []
            cline_begin = self.y1 - 1
            cline_end = self.y1 - 1
            cline_start = True
            for j in range(self.y1, self.y2 + 1):
                if i == self.x1:
                    continue
                cell = self.cells[i - 1][j - 1]
                up_cell = self.cells[i - 2][j - 1]
                if cell.cell_type == 'plain' or cell.merged_idx != up_cell.merged_idx:
                    if cline_start:
                        cline_begin = j
                        cline_start = False
                    cline_end = j
                    if j == self.y2:
                        row_cline.append((cline_begin, cline_end))
                else:
                    if not cline_start:
                        row_cline.append((cline_begin, cline_end))
                        cline_start = True
            if i > self.x1:
                self.cline_ranges.append(row_cline)

    def set_texs(self):
        self.row_texs = []
        self.tex = '% Please add the following required packages to your document preamble:\n'
        if self.colors:
            self.tex += '% \\usepackage{xcolor}'
        self.tex += '% \\usepackage{multirow, makecell}\n'
        if self.args.excel_format:
            self.tex += '% \\usepackage{booktabs}'
            self.tex += '% \\usepackage{colortbl}'
            self.tex += '% \\usepackage{arydshln}'
        self.tex += '\n% color definition\n' if self.colors else ''
        for color in self.colors:
            self.tex += '\\definecolor{' + color + '}{HTML}{' + color + '}\n'
        if self.args.excel_format:
            self.tex += '''
{
\\setlength\\abovetopsep{0pt}
\\setlength\\belowbottomsep{0pt}
\\setlength\\aboverulesep{0pt}
\\setlength\\belowrulesep{0pt}
\\newcommand\\colorwrap[2]{
  \\arrayrulecolor{#1}#2
  \\arrayrulecolor{black}
}
'''
        self.tex += '\n\\begin{tabular}{*{' + str(self.y2 - self.y1 + 1) + '}{|c}|}\n'
        if self.args.excel_format:
            self.convert_excel()
            self.tex += ''.join(self.row_texs)
        else:
            self.tex += '\\hline\n'
            self.convert()
            self.tex += ''.join(self.row_texs)
            self.tex += '\\hline\n'
        self.tex += '\\end{tabular}'
        if self.args.excel_format:
            self.tex += '\n}'

    def get_all_cline_tex(self, cell, n, row, col):
        res = ''
        if n < self.x2 - self.x1 + 1 and cell.end:
            row_cline = self.cline_ranges[n - 1]
            if len(row_cline) == 1 and row_cline[0][0] == self.y1 and row_cline[0][1] == self.y2:
                res += '\\hline'
            else:
                for cline_range in row_cline:
                    res += '\\cline{' + str(cline_range[0]) + '-' + str(cline_range[1]) + '}\n'
        return res

    def convert(self):
        n = 1
        for i in range(self.x1, self.x2 + 1):
            row_tex = f'\n% row {n}\n'
            for j in range(self.y1, self.y2 + 1):
                cell = self.cells[i - 1][j - 1]
                if cell.skip():
                    out_text = cell.text_prop.text
                    # cline
                    out_text += self.get_all_cline_tex(cell, n, i, j)
                    row_tex += out_text
            n += 1
            self.row_texs.append(row_tex.replace('  &\n', '  &'))

    def convert_excel(self):
        print(self.clines.get_cline_range(self.clines.borders[4])[1].style.style)
        print(len(self.clines.borders[4]))
        print(self.clines.borders[4][5].style)
        n = 1
        for i in range(self.x1, self.x2 + 1):
            row_tex = f'\n% row {n}\n'
            if n == 1:
                row_tex += self.clines.get_row_cline_tex(self.clines.borders[i - 1])
                row_tex += '\\vspace{-' + str(self.vspace) + 'pt}'
            for j in range(self.y1, self.y2 + 1):
                cell = self.cells[i - 1][j - 1]
                if cell.skip():
                    out_text = cell.text_prop.text
                    row_tex += out_text
            # cline
            cline_tex = self.clines.get_row_cline_tex(self.clines.borders[i])
            row_tex += cline_tex
            n += 1
            self.row_texs.append(row_tex.replace('  &\n', '  &'))
