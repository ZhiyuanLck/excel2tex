from .Prop import *

class Cell:
    def __init__(self, cell, args):
        self.cell = cell
        self.args = args
        self.head = self
        self.cell_prop = CellProp()
        self.text_prop = TextProp()
        self.border = Border()

    def set_prop(self):
        self.cell_prop.set_prop(self)
        self.text_prop.set_prop(self)
        self.border.set_prop(self)

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

    def not_empty(self):
        return self.text_prop.text is not None or self.cell_prop.merged_idx
