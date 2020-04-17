from .Cell import Cell
from .MergedCell import MergedCell
from .help import scan_all

class Table:
    def __init__(self, ws, args):
        self.ws = ws
        self.args = args
        self.x1 = ws.min_row
        self.x2 = ws.max_row
        self.y1 = ws.min_column
        self.y2 = ws.max_column
        self.cells = [
                [Cell(self.ws.cell(i, j), self.args)
                    for j in range(self.y1, self.y2 + 1)]
                for i in range(self.x1, self.x2 + 1)]
        self.merged_cells = [MergedCell(idx + 1, cell.bounds) for idx, cell in enumerate(ws.merged_cells.ranges)]
        self.init_cell_matrix()
        self.set_bounds()
        self.set_props()
        self.set_clines()
        self.set_texs()

    def init_cell_matrix(self):
        for i in range(self.x1 - 1, self.x2):
            for j in range(self.y1 - 1, self.y2):
                cell = self.cells[i][j]
                r, c = cell.get_head(self.merged_cells, i, j)
                cell.head = self.cells[r - 1][c - 1]
                cell.cell_prop.merged_idx = cell.get_merged_idx(self.merged_cells, i + 1, j + 1)
#                  cell.text_prop.text = cell.head.text_prop.text if cell.cell_prop.merged_idx else self.ws.cell(i + 1, j + 1).value
                cell.text_prop.text = self.ws.cell(i + 1, j + 1).value

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
                cell.cell_prop.cell_type = "plain"
                cell.cell_prop.begin = True if j == self.y1 else False
                # real end or block end
                cell.cell_prop.end = True if j == self.y2 else False
                if cell.cell_prop.merged_idx:
                    merged_cell = self.merged_cells[cell.cell_prop.merged_idx - 1]
                    cell.cell_prop.cell_type = merged_cell.get_type(i, j)
                    cell.cell_prop.height = merged_cell.x2 - merged_cell.x1 + 1
                    cell.cell_prop.width = merged_cell.y2 - merged_cell.y1 + 1
                    cell.cell_prop.end = merged_cell.is_end(self.y2)

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
                if cell.cell_prop.cell_type == 'plain' or cell.cell_prop.merged_idx != up_cell.cell_prop.merged_idx:
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
        self.tex = '% Please add the following required packages to your document preamble:\n% \\usepackage{multirow, makecell}\n'
        self.tex += r'\begin{tabular}{*{' + str(self.y2 - self.y1 + 1) + '}{|c}|}\n'
        self.tex += '\\hline\n'
        self.row_texs = []
        self.convert()
        self.tex += ''.join(self.row_texs)
        self.tex += '\\hline\n'
        self.tex += '\\end{tabular}'


    def convert(self):
        n = 1
        for i in range(self.x1, self.x2 + 1):
            row_tex = f'% row {n}\n'
            for j in range(self.y1, self.y2 + 1):
                cell = self.cells[i - 1][j - 1]
                if cell.cell_prop.cell_type != "multicolumn_other" and cell.cell_prop.cell_type != "block_firstline_other":
                    out_text = cell.text_prop.text
                    # cline
                    if n < self.x2 - self.x1 + 1 and cell.cell_prop.end:
                        row_cline = self.cline_ranges[n - 1]
                        if len(row_cline) == 1 and row_cline[0][0] == self.y1 and row_cline[0][1] == self.y2:
                            out_text += '\\hline'
                        else:
                            for cline_range in row_cline:
                                out_text += '\\cline{' + str(cline_range[0]) + '-' + str(cline_range[1]) + '}\n'
                    row_tex += out_text
            n += 1
            self.row_texs.append(row_tex.replace('  &\n', '  &'))
