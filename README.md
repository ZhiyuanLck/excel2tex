# Convert excel table to latex table

This tool convert excel table to latex table in human-readable format.

## Requirements

Please add the following required packages to your document preamble:

```tex
\usepackage{multirow, makecell}
```

Following python packages are needed:

```shell
pip install openpyxl
```

## Usage

```text
usage: excel2tex.py [-h] -s SOURCE -o TARGET

optional arguments:
  -h, --help  show this help message and exit
  -s SOURCE   source file (default: table.xlsx)
  -o TARGET   target file (default: table.tex)
  -e {utf-8,utf-8-sig}  file encoding (default: utf-8), if there is mess code, set it to utf-8-sig
```

## Note

1. The height of every merged cell must be greater than the number of lines in your text.
2. Make sure that the height of every merged cell is **even**.
3. Make the size of the merged cell suitable, i.e., do not give more space than that you need.

## Trouble shooting

### Mess code

Try to set encoding to `utf-8-sig`, for example

```shell
python excel2tex.py -s table.xlsx -o table.tex -e utf-8-sig
```

### Missing vertical line or redundant empty row

Please check the space of every merged cell whether they satisfy the conditions in **Note**.
