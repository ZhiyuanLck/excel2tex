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
```
