from copy import copy

class ClineRange:
    def __init__(self, start, end, style):
        self.start = start
        self.end = end
        self.style = style

class Line:
    style_dic = {
            #                         line_number
            #                    is_dash | l_n | width  |  dash_width | gap_width
            'thin':             (False,    1,    0.4,      0.4,         0.4),
            'medium':           (False,    1,    0.6,      0.4,         0.4),
            'thick':            (False,    1,    0.8,      0.4,         0.4),
            'double':           (False,    2,    0.4,      0.4,         0.4),
            'hair':             (True,     1,    0.4,      0.4,         0.4),
            'dotted':           (True,     1,    0.4,      0.4,         0.4),
            'dashed':           (True,     1,    0.4,      0.4,         0.4),
            'dashDot':          (True,     1,    0.4,      0.4,         0.4),
            'dashDotDot':       (True,     1,    0.4,      0.4,         0.4),
            'mediumDashed':     (True,     1,    0.4,      0.4,         0.4),
            'mediumDashDot':    (True,     1,    0.4,      0.4,         0.4),
            'mediumDashDotDot': (True,     1,    0.4,      0.4,         0.4),
            'slantDashDot':     (True,     1,    0.4,      0.4,         0.4),
            }
    def __init__(self):
        self.style = None
        self.is_none = True
        self.is_dash = False
        self.line_number = 1
        self.width = 0.4
        self.dash_width = 0.4
        self.gap_width = 0.4
        self.is_colored = False
        self.color = None

    def set_style(self, style):
        self.style = style
        (self.is_dash, self.line_number, self.width, self.dash_width,
                self.gap_width) = self.style_dic[style]

    def get_hline(self, table, i, j):
        if j == 0:
            return table.cells[i][j].border.left
        left = table.cells[i][j - 1].border.right
        right = table.cells[i][j].border.left
        return left if left.style is not None else right

    def get_cline(self, table, i, j):
        if i == 0:
            return table.cells[i][j].border.top
        up = table.cells[i - 1][j].border.bottom
        down = table.cells[i][j].border.top
        return up if up.style is not None else down

    def set_line(self, table, i, j, type):
        border = self.get_cline(table, i, j) if type == 'cline' else self.get_hline(table, i, j)
        if border.style is not None:
            self.is_none = False
            self.set_style(border.style)
            if border.color is not None:
                self.is_colored = True
                self.color = border.color

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
        x_max = table.x2 - 1
        y_max = table.y2 - 1
        self.row_axis = x_max if type == 'cline' else y_max
        self.col_axis = y_max if type == 'cline' else x_max
        self.borders = [[Line()
            for j in range(self.col_axis)]
            for i in range(self.row_axis)]

    def get_coor(self, i, j, type):
        return (i, j) if type == 'cline' else (j, i)

    # after set_props
    def set_lines(self, table, type):
        for i in range(self.row_axis):
            for j in range(self.col_axis):
                x, y = self.get_coor(i, j, type)
                self.borders[i][j].set_line(table, x, y, type)
#                  if type == 'cline' and i == 4 and j == 5:
#                      print(self.borders[i][j].style)

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

    def cut_range(self, cline):
        start_idx = len(cline) - 1
        end_idx = 0
        for i in range(len(cline)):
            if cline[i].style:
                start_idx = i
                break
        for i in range(len(cline) - 1, -1, -1):
            if cline[i].style:
                end_idx = i
                break
        return start_idx, end_idx

    def get_cline_range(self, cline):
        start_idx, end_idx = self.cut_range(cline)
        # one line
        if start_idx == end_idx:
            return [ClineRange(start_idx, end_idx, cline[0])]
        # none line
        if end_idx == 0:
            return []
        i = start_idx
        start = start_idx
        end = end_idx
        res = []
        pre = Line()
        for border in cline[start_idx:end_idx + 1]:
            if not border.is_eql(pre):
                # previous group end
                if i > start_idx and pre.style:
                    end = i - 1
                    res.append(ClineRange(start + 1, end + 1, pre))
                # new group begin
                if border.style:
#                      if i == end_idx:
#                          print(i, pre.style, border.style)
                    start = i
                # current group end
            if i == end_idx:
                print(start, len(res), pre.style, border.style, border.is_eql(pre))
                res.append(ClineRange(start + 1, i + 1, border))
                print('test', res[-1].style.style)
            pre = copy(border)
            i += 1
        print('test', res[-1].style.style)
        return res

    def get_cline_tex(self, cline_range):
        tex = ''
        style = cline_range.style
        if style.is_dash:
            tex += '\\cdashline{' + str(cline_range.start) + '-' + str(cline_range.end) + '}'
        elif not style.is_none:
            tex += '\\cmidrule[' + str(style.width) + 'pt' + ']{' + str(cline_range.start) + '-' + str(cline_range.end) + '}'
        if style.is_colored:
            tex = '\\colorwrap{' + style.color + '}' + '{' + tex + '}'
        # not support
#          if style.line_number == 2:
#              tex = tex + '\n' + tex
        return tex

    def set_vspace(self):
        for cline in self.borders:
            cline_ranges = self.get_cline_range(cline)
            max_width = 0
            for cline_range in cline_ranges:
                if not cline_range.style.is_dash:
                    max_width = max(0.4, max_width, cline_range.style.width)
            self.table.vspace += max_width

    def get_row_cline_tex(self, cline):
        tex = ''
        cline_ranges = self.get_cline_range(cline)
        if not cline_ranges:
            return tex
#          max_width = 0
        for cline_range in cline_ranges:
#              if not cline_range.style.is_dash:
#                  max_width = max(0.4, max_width, cline_range.style.width)
            tex += self.get_cline_tex(cline_range) + '\n'
#          self.table.vspace += max_width
        return tex
