class TextProp:
    def __init__(self, table):
        self.table = table
        self.text = ''
        self.i = False
        self.b = False
        self.color = '000000'

    def set_prop(self, cell):
        self.i = cell.cell.font.i
        self.b = cell.cell.font.b
        color = cell.cell.font.color.rgb
        if color is not None and isinstance(color, str):
            self.color = color[2:]
        if cell.merged_idx:
            self.i = cell.head.text_prop.i
            self.b = cell.head.text_prop.b
            self.color = cell.head.text_prop.color
            self.text = cell.head.text_prop.text
        self.text = self.format_text(cell) if cell.args.excel_format else self.get_cell_tex(cell)

    def format_text(self, cell):
        text = self.text
        # no text
        if text is None:
            text = ''
        # not string
        if not isinstance(text, str):
            text = str(text)
        # deal with escape character
        text = text.replace('#', r'\#')
        text = text.replace('%', r'\%')
        # inline math
        if not cell.args.math:
            text = text.replace('$', r'\$')
        # line break
        text = text.replace('\n', r' \\')
        # excel format
        if text:
            if cell.args.excel_format:
                if self.i:
                    text = '\\textit{' + text + '}'
                if self.b:
                    text = '\\textbf{' + text + '}'
                if self.color != '000000':
                    self.table.colors.add(self.color)
                    text = '\\textcolor{' + self.color + '}' + '{' + text + '}'
        return text

    # not used by excel format
    def get_cell_tex(self, cell):
        text = self.format_text(cell)
        begin_line = '|' if cell.begin else ''
        str_dic = {
                'plain': text,
                'multirowcell_begin': '\\multirowcell{' + str(cell.height) + '}{' + text + '}',
                'multirowcell_other': '',
#                  'multicolumn_begin': '\\multicolumn{' + str(cell.width) + '}{' + '|' if self.parameters['begin'] else '' + 'c|}{' + text + '}',
                'multicolumn_begin': '\\multicolumn{' + str(cell.width) + '}{' + begin_line + 'c|}{' + text + '}',
                'multicolumn_other': '',
                'block_firstline_begin': '\\multicolumn{' + str(cell.width) + '}{' + begin_line + 'c|}{' + '\\multirowcell{' + str(cell.height) + '}{' + text + '}' + '}',
                'block_firstline_other': '',
                'block_placeholder_begin': '\\multicolumn{1}{|c}{}',
                'block_placeholder_end': '\\multicolumn{1}{c|}{}',
                'block_placeholder_other': '\\multicolumn{1}{c}{}'
                }
        text = str_dic[cell.cell_type]
        if not cell.begin:
            text = '& ' + text
        if cell.end:
            text += r' \\'
        if text:
            text = '  ' + text.strip() + '\n'
        return text

class BorderLine:
    def __init__(self, table):
        self.table = table
        self.style = None
        self.color = None

    def set_prop(self, cell, pos):
        border = getattr(cell.cell.border, pos)
        if border.style is not None:
            self.style = border.style
            color = border.color.rgb
            if color is not None and isinstance(color, str):
                self.color = color[2:]
                if self.color != '000000':
                    self.table.colors.add(self.color)

class Border:
    def __init__(self, table):
        self.left = BorderLine(table)
        self.right = BorderLine(table)
        self.top = BorderLine(table)
        self.bottom = BorderLine(table)

    def set_prop(self, cell):
        for pos in ['left', 'right', 'top', 'bottom']:
            getattr(self, pos).set_prop(cell, pos)
