set  i  "������Ŀ����"  /data1*data60/ ;
alias(i,j) ;

Parameters
         sam(i,j)   "�������"
         Q0(i,j)    "��ʼ����"
         H0         "��ʼ����֮��"   ;
* ����ϲ����SAM����ֵ��sam����
$libinclude xlimport sam C:\Users\Administrator\Desktop\temp\GAMS\SAM.xlsx sheet2!A1:BI61
Q0(i,j) = sam(i,j);
H0 = sum((i,j),sam(i,j));

variables
         Q(i,j)     "�������SAM������"
         H          "������SAM������֮��"
         Hratio     "����ǰ������֮�͵ı�ֵ"
         z          "��Ҫ�Ż���Ŀ�꺯��"  ;
positive variable Q(i,j);

equations
         totalsum   "����������֮�͵ļ��㺯��"
         Hratiodef  "����ǰ�����ݺͱ�ֵ�ļ��㺯��"
         balance    "�С��к���ȵ�Լ������"
         object     "�Ż���Ŀ�꺯��" ;
* ��������
totalsum..   H =e= sum((i,j),Q(i,j));
Hratiodef..  Hratio =e= H/H0;
balance(i).. sum(j,Q(i,j)) =e= sum(j,Q(j,i));
object..     z =e=sum((i,j)$sam(i,j),(1/H)*Q(i,j)*log(Q(i,j)/sam(i,j))-log(Hratio));
* ������ʼֵ
Q.l(i,j) = Q0(i,j);
H.l = H0;
* Լ��������Χ
Hratio.lo = 0.95;
Hratio.up = 1.05;

model samsolving /all/;
solve samsolving using nlp minimizing z;

* д��Excel�ļ�
$LIBInclude Xlexport Q.l C:\Users\Administrator\Desktop\temp\GAMS\SAM.xlsx sheet3!A1:BI61
