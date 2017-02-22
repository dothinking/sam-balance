# sam-balance
社会核算矩阵的账户集结与平衡

首先使用`0-1矩阵法`集结账户，然后使用`直接交叉熵法`平衡SAM，具体流程为：

* 使用Matlab软件从Excel文件读入原始数据
* 分别左乘和右乘相应的矩阵集结账户
* 输出集结后的SAM数据到Excel文件
* 使用GAMS软件读入上一步输出的SAM数据
* 优化得到平衡后的SAM数据
* 输出到Excel文件

## 合并矩阵

假设$\alpha$是一个元素为0或1的行向量，则根据矩阵乘法运算规律可知，矩阵$M$左乘$\alpha$的效果为：将矩阵$M$的某些行相加，而这些行恰好是行向量$\alpha$中元素1对应的位置。同理，矩阵$M$右乘元素为0或1的列向量$\beta$，可以将$\beta$中元素1位置对应矩阵$M$的列求和。因此，得到如下结论：

> 左乘0-1行向量可以合并行，右乘0-1列向量可以合并列。

于是，将合并不同行的行向量$(\alpha_1, \alpha_2, \cdots)$组合为矩阵$A$；由于需要合并的列与行是相对应的，所以右乘的合并矩阵即为$A$的转置矩阵$A^T$。最终，合并后的矩阵$N=A\,M\,A^T$。

## 平衡矩阵

平衡SAM表使用的是直接交叉熵法，模型及原理不在此阐述，GAMS求解程序可以参考相关专业的材料，例如《可计算一般均衡模型的基本原理与编程》的第五章《SAM表的平衡》。实际操作中需要注意SAM表的读写，这可以参考GAMS相关文档：

> [Customized data interchange links for spreadsheets](http://www.gams.com/help/index.jsp?topic=%2Fgams.doc%2Fuserguides%2Fmccarl%2Findex.html)

### SAM表的读入

对于数据较少的情况可以直接以`table`的形式输入，如果数据量很大，则可以直接读入Excel文件，代码如下：

```
set  i  "定义项目集合，假设有60个账户"  /data1*data60/ ;
alias(i,j) ;

Parameters    sam(i,j)   "核算矩阵"  ;
* 读入合并后的SAM表，赋值给sam参数
$libinclude xlimport sam C:\Users\Administrator\GAMS\SAM.xlsx sheet2!A1:BI61
```

### SAM表的写出

类似地，将平衡后的SAM表数据写出到Excel文件的代码如下：

```
* 写Excel文件, Q为平衡调整后的数据
$LIBInclude Xlexport Q.l C:\Users\Administrator\Desktop\temp\GAMS\SAM.xlsx sheet3!A1:BI61
```

最后特别注意：

读入的SAM表中必须包含与GAMS程序中的`set`一致的索引（表头），写出Excel文件的工作表中也必须事先存在与SAM表一致的索引。
