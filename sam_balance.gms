set  i  "定义项目集合"  /data1*data60/ ;
alias(i,j) ;

Parameters
         sam(i,j)   "核算矩阵"
         Q0(i,j)    "初始数据"
         H0         "初始数据之和"   ;
* 读入合并后的SAM表，赋值给sam参数
$libinclude xlimport sam C:\Users\Administrator\Desktop\temp\GAMS\SAM.xlsx sheet2!A1:BI61
Q0(i,j) = sam(i,j);
H0 = sum((i,j),sam(i,j));

variables
         Q(i,j)     "调整后的SAM表数据"
         H          "调整后SAM表数据之和"
         Hratio     "调整前后数据之和的比值"
         z          "需要优化的目标函数"  ;
positive variable Q(i,j);

equations
         totalsum   "调整后数据之和的计算函数"
         Hratiodef  "调整前后数据和比值的计算函数"
         balance    "行、列和相等的约束函数"
         object     "优化的目标函数" ;
* 函数定义
totalsum..   H =e= sum((i,j),Q(i,j));
Hratiodef..  Hratio =e= H/H0;
balance(i).. sum(j,Q(i,j)) =e= sum(j,Q(j,i));
object..     z =e=sum((i,j)$sam(i,j),(1/H)*Q(i,j)*log(Q(i,j)/sam(i,j))-log(Hratio));
* 给定初始值
Q.l(i,j) = Q0(i,j);
H.l = H0;
* 约束调整范围
Hratio.lo = 0.95;
Hratio.up = 1.05;

model samsolving /all/;
solve samsolving using nlp minimizing z;

* 写入Excel文件
$LIBInclude Xlexport Q.l C:\Users\Administrator\Desktop\temp\GAMS\SAM.xlsx sheet3!A1:BI61
