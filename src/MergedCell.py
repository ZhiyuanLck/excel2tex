class MergedCell:
    def __init__(self, idx, bounds):
        self.y1, self.x1, self.y2, self.x2 = bounds
        # idx'th mergedcell
        self.idx = idx

    def is_merged(self, row, col):
        return True if row >= self.x1 and row <= self.x2 and col >= self.y1 and col <= self.y2 else False

    def get_type(self, row, col):
        if self.x1 != self.x2 and self.y1 != self.y2:
            if row == self.x1:
                return "block_firstline_begin" if col == self.y1 else "block_firstline_other"
            if col == self.y1:
                return "block_placeholder_begin"
            return "block_placeholder_end" if col == self.y2 else "block_placeholder_other"
        if self.x1 == self.x2:
            return "multicolumn_begin" if col == self.y1 else "multicolumn_other"
        if self.y1 == self.y2:
            return "multirowcell_begin" if row == self.x1 else "multirowcell_other"

    def is_end(self, col_max):
        return self.y2 == col_max

    def get_head(self):
        return self.x1, self.y1

    def is_control(self, row, col):
        return row == self.x2 and col == self.y1

    def is_ignored(self, row, col):
        return row == self.x2 and col > self.y1

    def is_one_row(self, row, col):
        return self.x1 == self.x2
