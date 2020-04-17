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
- 单元格
  - coor
  - merged_idx
  - 对齐
  - 类型
  - 高
  - 宽
  - 颜色
  - 行首
  - 行尾
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

公共属性

- begin
- end

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
