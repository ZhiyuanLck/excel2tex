from .help import wrap_excl, wrap_ge

class OutputBase:
    def __init__(self, table):
        self.table = table
        self.cells = table.cells
        self.hlines = table.hlines.borders
        self.vlines = table.vlines.borders
        self.max_width = table.hlines.max_width
        self.vline_max_width = table.vlines.vline_max_width

    def get_vline(self, vline):
        if vline.is_none or vline.ignored:
            return ''
        return f'\\vsl{{{vline.color}}}{{{vline.width}pt}}'

class OutputHhline(OutputBase):
    def __init__(self, table):
        super().__init__(table)

    def get_tbvline(self, i, j):
        '''
        get top right and bottom right vlines of hlines[i][j]
        '''
        tvline = None
        bvline = None
        if i > self.table.x1 - 1:
            tvline = self.vlines[i - 1][j + 1]
        if i < self.table.x2:
            bvline = self.vlines[i][j + 1]
        return tvline, bvline

    def get_hhline(self, i):
        res = '\\hhline{\n'
        for j in range(self.table.y1 - 1, self.table.y2):
            rhline = self.hlines[i][j]
            lhline = rhline.get_lhline()
            # first vhline
            if j == self.table.y1 - 1:
                tvline, bvline = self.get_tbvline(i, j - 1)
                res += self.get_vhline(lhline, rhline, tvline, bvline)
            res += self.get_hline(i, j)
            tvline, bvline = self.get_tbvline(i, j)
            res += self.get_vhline(lhline, rhline, tvline, bvline)
        res += '}\n'
        return res

    def has_style(self, vline):
        if vline is None:
            return False
        return not vline.is_none

    # used when 't' and 'b' has only one 'true'
#      def get_single_vline(self, vline, vline_cal_width):
#          lcell = vline.get_cell('l')
#          rcell = vline.get_cell('r')
#          lcolor = 'white'
#          rcolor = 'white'
#          if lcell is not None:
#              lcolor = lcell.color
#          if rcell is not None:
#              rcolor = rcell.color
#          l = lcolor == 'white'
#          r = rcolor == 'white'
#          if l and r:
#              return ''
#          return f'\\vsl{{{rcolor}}}{{{vline.width}pt}}'

    def get_vhline(self, lhline, rhline, tvline, bvline):
        l, r, t, b = (self.has_style(vline) for vline in (lhline, rhline, tvline, bvline))
        if not t and not b:
            return ''
        elif t and not b:
            res = self.get_vline(tvline)
        elif not t and b:
            res = self.get_vline(bvline)
        else:
            lcolor = 'white'
            rcolor = 'white'
            if l:
                lcolor = lhline.color
            if r:
                rcolor = rhline.color
            bcolor = bvline.color
            if bcolor == lcolor or bcolor == rcolor:
                res = self.get_vline(bvline)
            else:
                res = self.get_vline(tvline)
        return wrap_ge(res) if res else res

    def get_hline(self, i, j):
        hline = self.hlines[i][j]
        max_w = self.max_width[i]
        if hline.first_hline:
            return self.get_base_hline(max_w, hline)
        cell = self.cells[i - 1][j]
        cell_color = cell.color
        if hline.ignored:
            return self.get_ignored_hline(max_w, hline, cell_color)
        return self.get_other_hline(max_w, hline, cell_color)

    # 没有样式的横线，对半填充
    def get_none_hline(self, max_w, hline):
        if not max_w:
            return '  ~\n'
        width = max_w / 2
        tcell = hline.get_cell('t')
        bcell = hline.get_cell('b')
        tcolor = None
        bcolor = None
        if tcell is not None and tcell.color != 'white':
            tcolor = tcell.color
        if bcell is not None and bcell.color != 'white':
            bcolor = bcell.color
        if tcolor is None and bcolor is None:
            return '  ~\n'
        if tcolor is None and bcolor is not None:
            res = self.sfill(bcolor, width)
        if bcolor is None:
            bcolor = 'white'
        res = self.ssfill(tcolor, width, bcolor, width)
        return wrap_excl(res)

    def get_base_hline(self, max_w, hline):
        if hline.is_none:
#              return '  ~\n'
            return self.get_none_hline(max_w, hline)
        color = hline.color if hline.is_colored else 'black'
        if hline.is_dash:
            res = self.dfill(color, hline.get_dash_width(), hline.width, hline.format_dash())
            return wrap_excl(res)
        res = self.sfill(hline.color, hline.width)
        return wrap_excl(res)

    def get_other_hline(self, max_w, hline, cell_color):
        if cell_color == 'white':
            return self.get_base_hline(max_w, hline)
        if hline.is_none:
#              return '  ~\n'
            return self.get_none_hline(max_w, hline)
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
        for j in range(self.table.y1 - 1, self.table.y2):
            cell = self.cells[i][j]
            lvline = vlines[j]
            rvline = cell.get_rvline()
            if not cell.ignored:
                row_tex += self.wrap_cell(cell, self.get_cell(cell, lvline, rvline))
        row_tex += '\\\\\n'
        return row_tex

    def wrap_cell(self, cell, tex):
        before = '  '
        if not cell.begin:
            before += '&'
            if tex:
                before += ' '
        return before + tex

    def get_cell(self, cell, lvline, rvline):
        left = '' if not cell.begin else self.get_cell_vline(lvline)
        right = self.get_cell_vline(rvline)
        align = self.get_align(cell)
        return f'\\multicolumn{{{cell.width}}}{{{left}{align}{right}}}{{{self.get_col(cell)}}}\n'

    def get_align(self, cell):
        if cell.merged_idx and not cell.control_cell:
            align = 'c'
        else:
            align = cell.align
        before = ''
        prop = cell.text_prop
        if prop.color != '000000':
            before += f'\\color{{{prop.color}}}'
        if prop.i:
            before += '\\itshape'
        if prop.b:
            before += '\\bfseries'
        if before:
            before = f'>{{{before}}}'
        return before + align

    def get_cell_vline(self, vline):
        res = self.get_vline(vline)
        return f'!{{{res}}}' if res else res
#          if vline.is_none:
#              lcell = vline.get_cell('l')
#              rcell = vline.get_cell('r')
#              lcolor = 'white'
#              rcolor = 'white'
#              if lcell is not None:
#                  lcolor = lcell.color
#              if rcell is not None:
#                  rcolor = rcell.color
#              l = lcolor == 'white'
#              r = rcolor == 'white'
#              if l and r:
#                  return ''
#              x, y = vline.idx
#              width = self.vline_max_width[y] / 2
#              res = f'\\vsl{{{lcolor}}}{{{width}pt}}'
#              res += f'\\vsl{{{rcolor}}}{{{width}pt}}'
#          else:
#              res = self.get_vline(vline)
#          return f'!{{{res}}}' if res else res

    def get_col(self, cell):
        res = ''
        text = cell.text_prop.text
        if cell.color != 'white':
            res += f'\\cellcolor{{{cell.color}}}'
        if cell.control_cell:
            if cell.one_row:
                res += text
            else:
                res += f'\n    \\multirow{{-{cell.height}}}*{{{text}}}'
        elif not cell.merged_idx:
            res += text
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
        before = f'\\begin{{tabular}}{{*{{{width}}}{{c}}}}\n'
        end = f'\\end{{tabular}}'
        return before + tex + end
