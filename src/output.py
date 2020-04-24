from .help import wrap_excl, wrap_ge

class OutputBase:
    def __init__(self, table):
        self.table = table
        self.cells = table.cells
        self.hlines = table.hlines.borders
        self.vlines = table.vlines.borders
        self.max_width = table.hlines.max_width

    def get_vline(self, vline):
        if vline.is_none or vline.ignored:
            return ''
        return f'\\vsl{{{vline.color}}}{{{vline.width}pt}}'

class OutputHhline(OutputBase):
    def __init__(self, table):
        super().__init__(table)

    def get_hhline(self, i):
        res = '\\hhline{\n'
        ignore_vhline = i == self.table.x1 - 1 or i == self.table.x2
        for j in range(self.table.y1 - 1, self.table.y2):
            if not ignore_vhline and j == self.table.y1 - 1:
                res += self.get_vhline(self.vlines[i - 1][j])
            res += self.get_hline(i, j)
            if not ignore_vhline:
                res += self.get_vhline(self.vlines[i - 1][j + 1])
        res += '}\n'
        return res

    def get_vhline(self, vline):
        res = self.get_vline(vline)
        return wrap_ge(res) if res else res

    def get_hline(self, i, j):
        hline = self.hlines[i][j]
        if hline.first_hline:
            return self.get_base_hline(hline)
        max_w = self.max_width[i]
        cell = self.cells[i - 1][j]
        cell_color = cell.color
        if hline.ignored:
            return self.get_ignored_hline(max_w, hline, cell_color)
        return self.get_other_hline(max_w, hline, cell_color)

    def get_base_hline(self, hline):
        if hline.is_none:
            return '  ~\n'
        color = hline.color if hline.is_colored else 'black'
        if hline.is_dash:
            res = self.dfill(color, hline.get_dash_width(), hline.width, hline.format_dash())
            return wrap_excl(res)
        res = self.sfill(hline.color, hline.width)
        return wrap_excl(res)

    def get_other_hline(self, max_w, hline, cell_color):
        if cell_color == 'white':
            return self.get_base_hline(hline)
        if hline.is_none:
            return '  ~\n'
        div = max_w - hline.width
        common_width = hline.get_dash_width()
        res = None
        if div and hline.is_dash:
            res = self.sdfill(cell_color, div, common_width, hline.color, hline.width, hline.format_dash())
        if not div and hline.is_dash:
            res = self.dfill(hline.color, common_width, hline.width, hline.format_dash())
        if div and not hline.is_dash:
            res = self.ssfill(cell_color, div, hline.color, hline.width)
        if not div and not hline.is_dash:
            res = self.sfill(hline.color, hline.width)
        return wrap_excl(res)

    def get_ignored_hline(self, max_width, hline, cell_color):
        if cell_color == 'white':
            return '  ~\n'
        res = self.sfill(cell_color, max_width)
        return wrap_excl(res)

    def dfill(self, color, width, height, style):
        return f'\\dfill{{{color}}}{{{width}}}{{{height}pt}}{{{style}}}'

    def sfill(self, color, height):
        return f'\\sfill{{{color}}}{{{height}pt}}'

    def ssfill(self, tcolor, theight, bcolor, bheight):
        return f'\\ssfill{{{tcolor}}}{{{theight}pt}}{{{bcolor}}}{{{bheight}pt}}'

    def sdfill(self, tcolor, theight, common_width, bcolor, bheight, style):
        return f'\\sdfill{{{tcolor}}}{{{theight}pt}}{{{common_width}}}{{{bcolor}}}{{{bheight}pt}}{{{style}}}'

class OutputCell(OutputBase):
    def __init__(self, table):
        super().__init__(table)

    def get_row(self, i):
        row_tex = ''
        vlines = self.vlines[i]
#          print(self.table.x1, self.table.x2)
        for j in range(self.table.y1 - 1, self.table.y2):
            cell = self.cells[i][j]
            lvline = vlines[j]
            rvline = vlines[j + 1]
            if not cell.ignored:
                row_tex += self.wrap_cell(cell, self.get_cell(cell, lvline, rvline))
        row_tex += '\\\\\n'
        return row_tex

    def wrap_cell(self, cell, tex):
        before = '  '
        if not cell.first_col:
            before += '&'
            if tex:
                before += ' '
        return before + tex

    def get_cell(self, cell, lvline, rvline):
        if cell.ignored:
            return ''
        left = '' if not cell.begin else self.get_cell_vline(lvline)
        right = self.get_cell_vline(rvline)
        return f'\\multicolumn{{{cell.width}}}{{{left}X{right}}}{{{self.get_col(cell)}}}\n'

    def get_cell_vline(self, vline):
        res = self.get_vline(vline)
        return f'!{{{res}}}' if res else res

    def get_col(self, cell):
        res = ''
        if cell.color != 'white':
            res += f'\\cellcolor{{{cell.color}}}'
        if cell.control_cell and not cell.one_row:
            res += f'\n    \\multirowcell{{-{cell.height}}}[0ex][{cell.align}]{{{cell.text_prop.text}}}'
        elif not cell.merged_idx:
            res += cell.text_prop.text
        return res

class Output:
    def __init__(self, table):
        self.table = table
        self.args = table.args
        hhline = OutputHhline(table)
        outputcell = OutputCell(table)
        tex = ''
        tex += hhline.get_hhline(0)
        for i in range(table.x1 - 1, table.x2):
            tex += f'% row {i - table.x1 + 2}\n'
            tex += outputcell.get_row(i)
            tex += hhline.get_hhline(i + 1)
        self.tex = self.wrap_tex(tex)

    def wrap_tex(self, tex):
        width = self.table.y2 - self.table.y1 + 1
        before = f'\\begin{{tabularx}}{{{self.args.width}}}{{*{{{width}}}{{X}}}}\n'
        end = f'\\end{{tabularx}}'
        return before + tex + end
