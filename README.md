# Convert excel table to latex table

A powerful tool that converts excel table to latex table in human-readable format.

Documentation for simplified Chinese: [简体中文](../master/doc/zh_cn.md)

<!--ts-->
   * [Convert excel table to latex table](#convert-excel-table-to-latex-table)
      * [Supported excel's styles](#supported-excels-styles)
      * [Before conversion](#before-conversion)
      * [Requirements](#requirements)
      * [Usage](#usage)
         * [Simple usage](#simple-usage)
         * [Enable all formats](#enable-all-formats)
      * [Trouble shooting](#trouble-shooting)
         * [Mess code](#mess-code)
      * [Support](#support)
      * [Buy me a coffee](#buy-me-a-coffee)

<!-- Added by: zhiyuan, at: Sat 25 Apr 2020 08:05:09 PM UTC -->

<!--te-->

## Supported excel's styles

- horizontal line
  - none
  - solid
  - dash (coming soon)
  - double line (coming soon)
- vertical line
  - none
  - solid
  - double line (coming soon)
- colored line
- colored cell
- colored text
- horizontal alignment
- text shape
  - italic
  - bold

## Before conversion

- Don't use **`theme colors`** that excel show in the panel, otherwise your color will not be recognized. `standard colors` or `more colors` will be ok.
- Font style is specified for the whole text, not a single character, i.e., if you have two characters one of which is bold and another is italic, then they will be both bold and italic.
- The height of every merged cell must not be less than the number of lines in your text.

## Requirements

If you are not using a `-e` option, please add the following required packages to your preamble:
```tex
\usepackage{multirow, makecell}
```

Following python package is required:
```shell
pip install openpyxl
```

## Usage

```text
usage: excel2tex.py [-h] [-s SOURCE] [-o TARGET] [--setting SETTING] [--sig] [-m] [-e]

optional arguments:
  -h, --help            show this help message and exit
  -s SOURCE             source file (default: table.xlsx)
  -o TARGET             target file (default: table.tex)
  --setting SETTING     setting file (default: setting.tex)
  --sig                 set file encoding to utf-8-sig, only use when there is mess code.
  -m, --math            enabel inline math
  -e, --excel-format    enabel all formats
```

If you are using windows and have no python installed, an executable file is provided [here (outdated)](https://github.com/ZhiyuanLck/excel2tex/releases/tag/0.1)

### Simple usage

We have the following excel table to be converted to latex table.

![Excel table](img/excel_table.png)

`python excel2tex.py` is the simplest method to do this, which means converting an excel file of name `table.xlsx` to a tex file of name `table.tex`. And because you are not using the `-e` option, the table is resolved in the simplest way:
- All lines are drawn
- No element will be colored
- Text are all centered
- All characters are converted to what they have been.

So this is the converted table drawn in latex. Generated code is in [`simple.tex`](../master/examples/simple.tex)

![latex table of simple format](img/simple.png)

### Enable all formats

If you want more styles to be resolved, try to use the `-e` option. The following command will convert the table with most of the styles you have set in excel. And a setting file of name `setting.tex` by default (change it by `--setting mysetting.tex`). Please input the setting file in your preamble. The `-m` option ensure that the character `$` is not escaped so that you can enter the math mode in excel just as in latex.

```shell
python excel2tex.py -e -m
```

Here is the result. The generated code is in [`all.tex`](../master/examples/all.tex)

![latex table of all format](img/all.png)

## Trouble shooting

### Mess code

Try to set encoding to `utf-8-sig`, for example

```shell
python excel2tex.py -s table.xlsx -o table.tex --sig
```

## Support
If you have some emergency trouble with this tool, send me your code to my email: lichangkai225@qq.com

## Buy me a coffee

Does this tool help you? You can buy me coffee!

![wechat](img/wechat.jpg) ![alipay](img/alipay.jpg)
