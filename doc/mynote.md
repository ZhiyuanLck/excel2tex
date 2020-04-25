## 要实现的功能

### 基础功能

文本

- 颜色
- 斜体、粗体
- 对齐

单元格

- 颜色

框线

- 横线
  - 宽度
  - 颜色
  - 样式
- 竖线
  - 宽度
  - 颜色
  - 样式

### 组合
合并单元格

### 绘制过程
1. 每一行hhline
  - 基础属性设置
2. 每一行单元格

### 单个单元格语法

- 每个单元格写成`\muticolumn`形式
- 所有右侧竖线样式由`\muticolumn`控制
- 每行第一个单元格控制左侧竖线样式
- 单元格颜色在`\multicolumn`中设置而不是在`\multirowcell`中

```tex
\multicolumn{<cnum>}{<start vline> <align> <right vline>}{<cell color> <content>}
```

### 合并单元格语法

- 合并单元格布局由合并单元格左下角单元格控制
- `\multirowcell`中的对齐与`\multicolumn`中的对齐保持一致，即文本对齐方式
- 除去左下角单元格，其他单元格文本为空，单元格颜色一致

右侧竖线跳过的情形：
- block列数大于1且不是一行最后一个单元格

```tex
% <content>内容
\multirowcell{<rnum}[0ex][<align>]{<text>}
```

### 横线语法

- 所有下侧横线样式在当前行换行符`\\`后由`\hhline`控制
- 第一行单元格上侧的横线单独设置

下侧横线填充颜色

1. 无(无横线且无填充颜色)
2. 虚线(线宽最大 或 单元格无填充颜色)
3. 实线(线宽最大 或 单元格无填充颜色)
4. 实线填充(单元与下面的单元格属于一个合并单元,且有填充颜色)
5. 虚线+实线填充(线宽不是最大，且有填充颜色)
6. 实线+实线填充(线宽不是最大，且有填充颜色)

实际可以归纳为填充两个pattern，都用`\xleaders`实现，top pattern一定是实线，分别判断两个pattern是否需要填充（高度0pt）即可

```tex
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
```

采取竖线覆盖横线的方式（第一行和最后一行除外），`hhline`中对`|`需要有`\beforevline`设置，最后需要有`\aftervline`设置

```tex
\hhline{
  <line style spec>
  <line style spec>
  …
  <after settings>
}
```

`<line style spec>`

```tex
>{<before settings>}
!{<line style>}
```

相关宏命令
```tex
%% vline settings
% #1 rule width #2 color
\newcommand{\beforevline}[2]{%
  \global\setlength\arrayrulewidth{#1}\arrayrulecolor{#2}%
}
\newcommand{\aftervline}{%
  \global\setlength\arrayrulewidth{0.4pt}\arrayrulecolor{black}%
}
```

# 步骤

1. 初始化单元格矩阵: 确定idx, text属性
2. 移除空行、空列: 得到有效的范围bounds
3. 重新扫描单元格矩阵，确定其他属性
4. 转换
  - 单元格
  - 边框

# 属性列表

- head 属性head
- cell 真实的cell
- coor
- merged_idx
- 对齐
- 类型
- 高
- 宽
- 颜色
- 行首
- 行尾
- first_row
- last_row
- firt_col
- last_col
- 文本
  - 内容
  - 斜体
  - 粗体
  - 颜色
- 边框
  - 四条边框

# 表格转换

- 矩阵每一列分成了若干等宽的长方形或正方行
- 扫描每一行，如果是正方形输出1, 如果是长方形
  - 如果是长方形的开头，输出2
  - 否则输出0

单元格类型

- plain
- multirowcell_begin
  - height
- multirowcell_other
- multicolumn_begin
  - width
- multicolumn_other: skip
- block_firstline_begin
  - height
  - width
- block_firstline_other: skip
- block_placeholder_begin
- block_placeholder_end
- block_placeholder_other

扫描每一个单元格

1. plain cell
2. multirowcell
3. multicolumn
4. block
  a. 第一行
    - 第一个插入语句
    - 其他的跳过
  b. 其他行
    - 最后一个multicolumn加|
    - 其他不加

<!-- 1. 判断multicolumn是否在末尾 -->

## cline处理

### 属性设置

1. 线段类型
- is_none
- is_dash
2. 宽度
  - thin
  - medium
  - thick
3. 颜色
- 有颜色
- 无颜色
4. dash line gap

### 设置cline
- 检查每行是否一致

### 通过cline移除空行

cline_x_min
hcell_x_min
self.x1

cline_x_max
hcell_x_max
self.x2

ccell_y_min
hline_y_min
self.y1

ccell_y_max
hline_y_max
self.y2

### 分类

开始分类的条件：

- 行第一个非空style
- cur不是None且cur != pre

继续分类的条件：

- cur不是None且cur = pre

结束分类的条件：
- cur != pre
- 倒数第一个非空style

### cline实现宏包
- 实线使用`booktabs`包，统一用`\cmidrule`实现不同线宽的cline。由于会使用竖线，将会导致产生顶端间距的sep设置为0，且在顶层减去每条cline的线宽
```tex
\setlength\abovetopsep{0pt}
\setlength\belowbottomsep{0pt}
\setlength\aboverulesep{0pt}
\setlength\belowrulesep{0pt}
```
- 颜色使用`colortbl`，定义宏命令
```tex
\newcommand\colorwrap[2]{
  \arrayrulecolor{#1}#2
  \arrayrulecolor{black}
}

# todo
- 检查异常的border
- text颜色放到>{}中
- vhline使用下面的竖线填充
- 没有样式的横线，填充下方单元格的颜色

### note
math宽度不要太高
已经知道的不兼容的包：`arydshln`, `stackengine`
