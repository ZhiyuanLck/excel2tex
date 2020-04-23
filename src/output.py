from .help import wrap_excl, wrap_ge

class OutputBase:
    def __init__(self, table):
        self.table = table
        self.cells = table.cells
        self.hlines = table.hlines.borders
        self.vlines = table.vlines.borders
        self.max_width = table.hlines.max_width

    def get_vline(self, vline):
        return f'\\vsl{{{vline.color}}}{{{vline.width}}}'

class OutputHhline(OutputBase):
    def __init__(self, table):
        super().__init__(table)

    def get_hhline(self, i):
        res = '\\hhline{\n'
        for j in range(self.table.y1 - 1, self.table.y2 + 1):
#          for j, hline in enumerate(self.hlines[i]):
            if j == self.table.y1 - 1:
                res += self.get_vhline(self.vlines[i - 1][j])
            res += self.get_hline(i, j)
            res += self.get_vhline(self.vlines[i - 1][j + 1])
        res += '}\n'
        return res

    def get_vhline(self, vline):
        if vline.is_none:
            return ''
        return wrap_ge(self.get_vline(vline))

    def get_hline(self, i, j):
        hline = self.hlines[i][j]
        if hline.first_hline:
            return self.get_base_hline(hline)
        max_w = self.max_width[j]
        cell = self.cells[i - 1][j]
        cell_color = cell.color if cell.is_colored else 'white'
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

    def dfill(self, color, width, height, style):
        return f'\\dfill{{{color}}}{{{width}}}{{{height}}}{{{style}}}'

    def sfill(self, color, height):
        return f'\\sfill{{{color}}}{{{height}}}'

    def ssfill(self, tcolor, theight, bcolor, bheight):
        return f'\\ssfill{{{tcolor}}}{{{theight}}}{{{bcolor}}}{{{bheight}}}'

    def sdfill(self, tcolor, theight, common_width, bcolor, bheight, style):
        return f'\\sdfill{{{tcolor}}}{{{theight}}}{{{common_width}}}{{{bcolor}}}{{{bheight}}}{{{style}}}'

class Output:
    def __init__(self, table):
        hhline = OutputHhline(table)
        for i in range(table.x2 + 1):
            if i == 4:
                print(hhline.get_hhline(i))
