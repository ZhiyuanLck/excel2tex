class Setting:
    def __init__(self, table):
        self.args = table.args
        self.output_setting(table.colors)

    def output_setting(self, colors):
        out = self.get_setting() + self.get_color(colors)
#          with open(self.args.setting, 'w', encoding=self.args.encoding) as f:
        with open('abc.tex', 'w', encoding=self.args.encoding) as f:
            f.write(out)

    def get_color(self, colors):
        res = ''
        for color in colors:
            res += self.set_color(color)
        return res

    def set_color(self, color):
        return f'\\definecolor{{{color}}}{{HTML}}{{{color}}}\n'

    def get_setting(self):
        return r'''\usepackage[table]{xcolor}
\usepackage{tabularx}
\usepackage{multirow, makecell}
\usepackage{colortbl}
\usepackage{dashrule}
\usepackage{ehhline}
\usepackage{arydshln}

%% vertical line
% vertical colored line #1 color #2 width
\newcommand{\vsl}[2]{\color{#1}\vrule width #2}
% doubled vline
% #1 first color #2 first width #3 sep #4 second color #5 second sep
\newcommand{\dvsl}[5]{%
  \vsl{#1}{#2}\hspace{#3}\vsl{#4}{#5}%
}

%% horizontal line
% colored solid line pattern
% #1 color #2 width #3 height
\newcommand{\hsp}[3]{\hbox{\textcolor{#1}{\rule{#2}{#3}}}}
% colored dash line pattern
% #1 color #2 width #3 height #4 style
\newcommand{\hdp}[4]{\hbox{\textcolor{#1}{\hdashrule{#2}{#3}{#4}}}}
% fill line
% #1 top fill #2 bottom fill
\newcommand{\leaderfill}[1]{%
  \xleaders\hbox{%
    \vbox{\baselineskip=0pt\lineskip=0pt#1}%
  }\hfill%
}
% top: solid, bottom: solid
% #1 top color #2 top height #3 bottom color#4 bottom height
\newcommand{\ssfill}[4]{%
  \leaderfill{\hsp{#1}{0.01pt}{#2}\hsp{#3}{0.01pt}{#4}}%
}
% top: solid, bottom: dashed
% #1 top color, #2 top height, #3 common width to expand
% #4 bottom color #5 bottom height, #6 bottom dash line style
\newcommand{\sdfill}[6]{%
  \leaderfill{\hsp{#1}{#3}{#2}\hdp{#4}{#3}{#5}{#6}}%
}
% single solid
% #1 color #2 height
\newcommand{\sfill}[2]{%
  \leaderfill{\hsp{#1}{0.01pt}{#2}}%
}
% single dash
% #1 color #2 width #3 height #4 style
\newcommand{\dfill}[4]{%
  \leaderfill{\hdp{#1}{#2}{#3}{#4}}%
}

%% vline settings
% #1 rule width #2 color
\newcommand{\beforevline}[2]{%
  \global\setlength\arrayrulewidth{#1}\arrayrulecolor{#2}%
}
\newcommand{\aftervline}{%
  \global\setlength\arrayrulewidth{0.4pt}\arrayrulecolor{black}%
}
'''
