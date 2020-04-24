from copy import copy

class ClineRange:
    def __init__(self, start, end, style):
        self.start = start
        self.end = end
        self.style = style

class Line:
    style_dic = {
            #                         line_number
            #                    is_dash | l_n | width  |  dash_style
            'thin':             (False,    1,    0.4,      [1, 1]),
            'medium':           (False,    1,    0.6,      [1, 1]),
            'thick':            (False,    1,    0.8,      [1, 1]),
            'double':           (False,    2,    0.4,      [1, 1]),
            'hair':             (True,     1,    0.4,      [1, 1]),
            'dotted':           (True,     1,    0.4,      [1, 1]),
            'dashed':           (True,     1,    0.4,      [1, 1]),
            'dashDot':          (True,     1,    0.4,      [1, 1]),
            'dashDotDot':       (True,     1,    0.4,      [1, 1]),
            'mediumDashed':     (True,     1,    0.4,      [1, 1]),
            'mediumDashDot':    (True,     1,    0.4,      [1, 1]),
            'mediumDashDotDot': (True,     1,    0.4,      [1, 1]),
            'slantDashDot':     (True,     1,    0.4,      [1, 1]),
            }
    def __init__(self):
        self.style = None
        self.is_none = True
        self.is_dash = False
        self.line_number = 1
        self.width = 0
        self.dash_style = []
        self.is_colored = False
        self.color = None
        # 第一行的顶端的横线
        self.first_hline = False
        # 表格末尾的横线
        self.last_hline = False
        # 表格左侧的竖线
        self.first_vline = False
        # 竖线是否不画
        self.ignored = False

    def format_dash(self):
        return 'pt '.join([str(x) for x in self.dash_style]) + 'pt'

    def get_dash_width(self):
        return str(sum(self.dash_style)) + 'pt'

    def set_style(self, style):
        self.style = style
        (self.is_dash, self.line_number, self.width, self.dash_style) = self.style_dic[style]

    def get_vline(self, table, i, j):
        if j == 0:
            return table.cells[i][j].border.left
        left = table.cells[i][j - 1].border.right
        right = table.cells[i][j].border.left
        return left if left.style is not None else right

    def get_hline(self, table, i, j):
        if i == 0:
            return table.cells[i][j].border.top
        up = table.cells[i - 1][j].border.bottom
        down = table.cells[i][j].border.top
        return up if up.style is not None else down

    def set_line(self, table, i, j, type):
        if i + 1 == table.x1:
            self.first_hline = True
        if i == table.x2:
            self.last_hline = True
        if j + 1 == table.y1:
            self.first_vline = True
        if type == 'hline' and i > table.x1 - 1 and i < table.x2:
            up = table.cells[i - 1][j].merged_idx
            down = table.cells[i][j].merged_idx
            if up == down and up > 0:
                self.ignored = True
        if type == 'vline' and j > table.y1 - 1 and j < table.y2:
            left = table.cells[i][j - 1].merged_idx
            right = table.cells[i][j].merged_idx
            if left == right and left > 0:
                self.ignored = True
        border = self.get_hline(table, i, j) if type == 'hline' else self.get_vline(table, i, j)
        if border.style is not None:
            self.is_none = False
            self.set_style(border.style)
            if border.color is not None:
                self.is_colored = True
                self.color = border.color
                table.colors.add(self.color)
            else:
                self.color = 'black'

    def is_eql(self, line):
        if self.style != line.style:
            return False
        if self.is_colored != line.is_colored:
            return False
        if self.color != line.color:
            return False
        return True

class LineMatrix:
    def __init__(self, table, type):
        # befor set_bounds
        self.table = table
        self.x_max = table.x2 - 1
        self.y_max = table.y2 - 1
        self.max_width = []
        self.borders = [[Line()
            for j in range(self.y_max)]
            for i in range(self.x_max)]

    # after set_props
    def set_lines(self, type):
        for i in range(self.x_max):
            max_w = 0
            for j in range(self.y_max):
                border = self.borders[i][j]
                border.set_line(self.table, i, j, type)
                max_w = max(border.width, max_w)
            self.max_width.append(max_w)

    def is_empty(self, line):
        for border in line:
            if border.style is not None:
                return False
        return True

    def is_full(self, line):
        style = line[0]
        for border in line:
            if not border.is_eql(style):
                return False
        return True

    def cut_range(self, hline):
        start_idx = len(hline) - 1
        end_idx = 0
        for i in range(len(hline)):
            if hline[i].style:
                start_idx = i
                break
        for i in range(len(hline) - 1, -1, -1):
            if hline[i].style:
                end_idx = i
                break
        return start_idx, end_idx

    def get_hline_range(self, hline):
        start_idx, end_idx = self.cut_range(hline)
        # one line
        if start_idx == end_idx:
            return [ClineRange(start_idx, end_idx, hline[0])]
        # none line
        if end_idx == 0:
            return []
        i = start_idx
        start = start_idx
        end = end_idx
        res = []
        pre = Line()
        for border in hline[start_idx:end_idx + 1]:
            if not border.is_eql(pre):
                # previous group end
                if i > start_idx and pre.style:
                    end = i - 1
                    res.append(ClineRange(start + 1, end + 1, pre))
                # new group begin
                if border.style:
                    start = i
                # current group end
            if i == end_idx:
                res.append(ClineRange(start + 1, i + 1, border))
            pre = copy(border)
            i += 1
        return res

    def get_hline_tex(self, hline_range):
        tex = ''
        style = hline_range.style
        if style.is_dash:
            tex += '\\cdashline{' + str(hline_range.start) + '-' + str(hline_range.end) + '}'
        elif not style.is_none:
            tex += '\\cmidrule[' + str(style.width) + 'pt' + ']{' + str(hline_range.start) + '-' + str(hline_range.end) + '}'
        if style.is_colored:
            tex = '\\colorwrap{' + style.color + '}' + '{' + tex + '}'
        # not support
#          if style.line_number == 2:
#              tex = tex + '\n' + tex
        return tex

    # not used
    def set_vspace(self):
        for hline in self.borders:
            hline_ranges = self.get_hline_range(hline)
            max_width = 0
            for hline_range in hline_ranges:
                if not hline_range.style.is_dash:
                    max_width = max(0.4, max_width, hline_range.style.width)
            self.table.vspace += max_width

    # not used
    def get_row_hline_tex(self, hline):
        tex = ''
        hline_ranges = self.get_hline_range(hline)
        if not hline_ranges:
            return tex
#          max_width = 0
        for hline_range in hline_ranges:
#              if not hline_range.style.is_dash:
#                  max_width = max(0.4, max_width, hline_range.style.width)
            tex += self.get_hline_tex(hline_range) + '\n'
#          self.table.vspace += max_width
        return tex
