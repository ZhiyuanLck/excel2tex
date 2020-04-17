from .help import colors

class CellProp:
    def __init__(self):
        self.coor = (1, 1)
        self.align = 'center'
        self.cell_type = 'plain'
        self.merged_idx = 0
        self.height = 1
        self.width = 1
        self.color = 'FFFFFF'
        self.begin = False
        self.end = False

    def set_prop(self, cell):
        self.coor = (cell.cell.row, cell.cell.column)
        self.align = cell.cell.alignment.horizontal
        if cell.cell_prop.merged_idx:
            self.color = cell.head.cell_prop.color
        else:
            color = cell.cell.fill.bgColor.rgb
            if color != '00000000' and color is not None and isinstance(color, str):
                self.color = color[2:]

class TextProp:
    def __init__(self):
        self.text = ''
        self.i = False
        self.b = False
        self.color = '000000'

    def set_prop(self, cell):
        if cell.cell_prop.merged_idx:
            self.i = cell.head.text_prop.i
            self.b = cell.head.text_prop.b
            self.color = cell.head.text_prop.color
        else:
            self.i = cell.cell.font.i
            self.b = cell.cell.font.b
            color = cell.cell.font.color.rgb
            if color is not None and isinstance(color, str):
                self.color = color[2:]
        self.format_text(cell)

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
                    text + '\\textit{' + text + '}'
                if self.b:
                    text + '\\textbf{' + text + '}'
                if self.color != '000000':
                    colors.add(self.color)
                    text = '\\textcolor{\\' + self.color + '}' + '{' + text + '}'
        begin_line = '|' if cell.cell_prop.begin else ''
        str_dic = {
                'plain': text,
                'multirowcell_begin': '\\multirowcell{' + str(cell.cell_prop.height) + '}{' + text + '}',
                'multirowcell_other': '',
#                  'multicolumn_begin': '\\multicolumn{' + str(cell.cell_prop.width) + '}{' + '|' if self.parameters['begin'] else '' + 'c|}{' + text + '}',
                'multicolumn_begin': '\\multicolumn{' + str(cell.cell_prop.width) + '}{' + begin_line + 'c|}{' + text + '}',
                'multicolumn_other': '',
                'block_firstline_begin': '\\multicolumn{' + str(cell.cell_prop.width) + '}{' + begin_line + 'c|}{' + '\\multirowcell{' + str(cell.cell_prop.height) + '}{' + text + '}' + '}',
                'block_firstline_other': '',
                'block_placeholder_begin': '\\multicolumn{1}{|c}{}',
                'block_placeholder_end': '\\multicolumn{1}{c|}{}',
                'block_placeholder_other': '\\multicolumn{1}{c}{}'
                }
        text = str_dic[cell.cell_prop.cell_type]
        if not cell.cell_prop.begin:
            text = '& ' + text
        if cell.cell_prop.end:
            text += r' \\'
        if text:
            text = '  ' + text.strip() + '\n'
        self.text = text

class BorderLine:
    def __init__(self):
        self.style = ''
        self.color = '000000'

    def set_prop(self, cell, pos):
        border = getattr(cell.cell.border, pos)
        if border.style is not None:
            self.style = border.style
            color = border.color.rgb
            if color is not None and isinstance(color, str):
                self.color = color[2:]

class Border:
    def __init__(self):
        self.left = BorderLine()
        self.right = BorderLine()
        self.top = BorderLine()
        self.bottom = BorderLine()

    def set_prop(self, cell):
        for pos in ['left', 'right', 'top', 'bottom']:
            getattr(self, pos).set_prop(cell, pos)
