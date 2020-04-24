from .Prop import TextProp, Border

class Cell:
    def __init__(self, cell, table):
        self.table = table
        self.cell = cell
        self.args = table.args
        self.head = self
        self.coor = (1, 1)
        self.align = 'center'
        self.cell_type = 'plain'
        self.merged_idx = 0
        self.height = 1
        self.width = 1
        self.color = 'white'
        self.begin = False
        self.end = False
        # block控制语句
        self.control_cell = False
        self.one_row = False
        # block中与控制单元格在同一行的其他单元格忽略
        self.ignored = False
        self.first_row = False
        self.first_col = False
        self.last_col = False
        self.text_prop = TextProp(self.table)
        self.border = Border(self.table)

    def set_prop(self):
        align_dic = {
                "center": "c",
                "left": "l",
                "right": "r",
                }
        self.coor = (self.cell.row, self.cell.column)
        self.align = self.cell.alignment.horizontal
        if self.align is None:
            self.align = 'center'
        self.align = align_dic[self.align]
        if self.merged_idx:
            self.align = self.head.align
            row, col = self.coor
            self.control_cell = self.table.merged_cells[self.merged_idx-1].is_control(row, col)
            self.ignored = self.table.merged_cells[self.merged_idx-1].is_ignored(row, col)
            self.one_row = self.table.merged_cells[self.merged_idx-1].is_one_row(row, col)
        self.text_prop.set_prop(self)

    def set_color(self):
        color = self.cell.fill.fgColor.rgb
        if color is not None and color != '00000000' and isinstance(color, str):
            self.color = color[2:]
            self.table.colors.add(self.color)
        elif self.merged_idx:
            self.color = self.head.color
        if self.color == '000000':
            self.color = 'white'

    def get_merged_idx(self, merged_cells, row, col):
        for m in merged_cells:
            if m.is_merged(row, col):
                return m.idx
        return 0

    def get_head(self, merged_cells, row, col):
        for m in merged_cells:
            if m.is_merged(row, col):
                return m.get_head()
        return row, col

    def skip(self):
        return self.cell_type != "multicolumn_other" and self.cell_type != "block_firstline_other"


    def not_empty(self):
        res = self.text_prop.text is not None or self.merged_idx
        if self.table.args.excel_format:
            res = res or self.color != 'white'
        return res
