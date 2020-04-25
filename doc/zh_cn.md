# excel表格转换工具

将excel中的表格转换为latex中的表格，支持识别多种格式。

<!--ts-->
   * [excel表格转换工具](#excel表格转换工具)
      * [支持识别的excel样式](#支持识别的excel样式)
      * [准备表格时的注意事项](#准备表格时的注意事项)
      * [依赖的python包和tex包](#依赖的python包和tex包)
      * [使用](#使用)
         * [简单用法](#简单用法)
         * [识别更多格式](#识别更多格式)
      * [可能出现的问题](#可能出现的问题)
         * [乱码](#乱码)
      * [联系邮箱](#联系邮箱)
      * [打赏](#打赏)

<!-- Added by: zhiyuan, at: Sat 25 Apr 2020 08:05:27 PM UTC -->

<!--te-->

## 支持识别的excel样式

- 横线
  - 无样式
  - 实线
  - 虚线 (即将上线)
  - 双线 (即将上线)
- 竖线
  - 无样式
  - 实线
  - 双线 (即将上线)
- 线条颜色
- 单元格颜色
- 文本颜色
- 水平对齐
- 文本形式
  - 斜体
  - 粗体

## 准备表格时的注意事项

- 设置颜色时不要选择**`主题颜色`**中的颜色，否则不能识别。选择`标准颜色`或者`更多颜色`通过色盘选取即可
- 不支持单个字符文本样式设置，一个单元格中出现的文本样式最终都会被设置成这个单元格所有文本的样式
- 你可以在excel中通过`alt + enter`进行手动换行，但是要注意最终文本的行数不能超过合并单元格的高度

## 依赖的python包和tex包

如果你没有使用`-e`选项，则把下面的tex包添加到导言区：
```tex
\usepackage{multirow, makecell}
```

安装以下python包
```shell
pip install openpyxl
```

## 使用

```text
usage: excel2tex.py [-h] [-s SOURCE] [-o TARGET] [--setting SETTING] [--sig] [-m] [-e]

可选参数
  -h, --help            显示命令帮助并退出
  -s SOURCE             要转换的excel文件名（默认table.xlsx）
  -o TARGET             要生成的tex文件名（默认table.tex）
  --setting SETTING     要放到导言区的设置文件名（默认setting.tex）
  --sig                 将编码设置为utf-8-sig，乱码的时候使用
  -m, --math            使得$...$被识别为数学模式
  -e, --excel-format    使用所有格式
```

如果你是windows用户且系统中没有python环境，这里提供了一个[二进制可执行文件](https://github.com/ZhiyuanLck/excel2tex/releases/tag/0.1)（落后于当前版本）使用

### 简单用法

下面是将要被转换的excel表格

![Excel table](../img/excel_table.png)

使用命令`python excel2tex.py`，你将得到一个只有必要格式的表格。这条命令的含义是将`table.xlsx`转换为`table.tex`。因为`-e`选项并没有被使用，生成的表格是最简的：
- 绘制所有的线条
- 文本和线条颜色默认黑色，无单元格填充
- 所有文本居中对齐
- 所有字符以本来的样子在latex生成的表格中显示

下面是转换生成的latex表格，生成的代码见[`simple.tex`](../master/simple.tex)

![latex table of simple format](../img/simple.png)

### 识别更多格式

如果你希望尽可能的识别你在excel中设置的格式，使用`-e`选项。下面的命令会尽可能的识别你在excel中设置的格式然后转换成latex表格，同时也会生成一个设置文件，里面加载了需要的宏包，定义了一些命令和颜色，请在导言区导入这个配置文件（`\input{setting.tex}`）。`-m`选项保证了字符`$`不会被转义（`\$`），以便识别数学模式。

```shell
python excel2tex.py -e -m
```

下面是转换生成的latex表格，生成的代码见[`all.tex`](../master/all.tex)

![latex table of all format](../img/all.png)

## 可能出现的问题

### 乱码

使用`--sig`选项将编码设置为`utf-8-sig`

```shell
python excel2tex.py -s table.xlsx -o table.tex --sig
```

## 联系邮箱

如果你有和这个脚本工具有关的紧急问题需要解决，把代码发到我的邮箱：lichangkai225@qq.com

## 打赏

如果这个工具帮助到了你，你喜欢它的话，欢迎打赏！

<!-- ![wechat](../img/wechat.png) ![alipay](../img/alipay.jpg) -->
<img src="https://github.com/ZhiyuanLck/excel2tex/blob/master/img/wechat.png" width="100">
<img src="https://github.com/ZhiyuanLck/excel2tex/blob/master/img/alipay.jpg" width="100">
