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
        self.color = 'FFFFFF'
        self.begin = False
        self.end = False
        self.first_row = False
        self.last_row = False
        self.first_col = False
        self.last_col = False
        self.text_prop = TextProp(self.table)
        self.border = Border(self.table)

    def set_prop(self):
        self.coor = (self.cell.row, self.cell.column)
        self.align = self.cell.alignment.horizontal
        if self.merged_idx:
            self.color = self.head.color
        else:
            color = self.cell.fill.bgColor.rgb
            if color != '00000000' and color is not None and isinstance(color, str):
                self.table.colors.add(color)
                self.color = color[2:]
        self.text_prop.set_prop(self)
#          self.border.set_prop(self)

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
        return self.text_prop.text is not None or self.merged_idx
