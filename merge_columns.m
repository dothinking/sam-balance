%% �ϲ����У�ע�⿪ʼ��Ϊ02��
% 01��ũҵ��02 ��
% 02��ú̿��03�� 
% 03��ʯ�ͣ�04��
% 04����Ȼ����05�� 
% 05���Ṥҵ��07-10��
% 06���ع�ҵ��06+11-18�� 
% 07����磨19��
% 08��ˮ�硢�˵��������20��
% 09��ȼ��������21��
% 10��ˮ��������22��
% 11������ҵ��23��
% 12����ͨ���估�ִ���24��
% 13��ס�޺Ͳ���ҵ��26��
% 14����������ҵ��25+27��
% 20160517
% ú 7��41   
% �Ṥҵ��13-39 45��48��49��50��70��84��85��94��99   
% �ع�ҵ��9-12��42-44��46��47��51-69��71-83��86-93��95��96  
% ������ 100-103
% ��ͨ��105-109��111��112
% ������ҵ��104��110��113-140  
clear;
%% ����׼��
filename = 'ȫ��2012��.xlsx';     % �������ݵ�excel�ļ���
input_sheet = 'sheet1';    % �������ݵĹ�������
range = 'B2:KB288';         % �������ݷ�Χ
output1_sheet = 'sheet2';  % ������ݵĹ�������
output2_sheet = 'sheet3';  % �����ƽ�����ݵĹ�������
final_sheet = 'sheet4';    % �����ƽ�����ݵĹ�������
label = 'data';            % �Զ������ݱ�ͷ��������Ҫ��GAMS��setһ��
% �ϲ�����
col = {1:5,7,39,96,97,[6,40],[12:38,44,47:49,69,83,84,93,98],[8:11,41:43,45,46,50:68,70:82,85:92,94,95],99:102,[104:108,110,111],[103,109,112:139]}; 
% �ϲ������Ŀ
title1 = {'ũ������ҵ','ʯ�ͺ���Ȼ�����ɲ�Ʒ','����ʯ�ͺͺ�ȼ�ϼӹ�Ʒ','���������������͹�Ӧ','ȼ�������͹�Ӧ','ú','�Ṥҵ','�ع�ҵ','����','��ͨ','������ҵ'};   
title2 = {'�Ͷ�','�ʱ�','����','��ҵ','����','̼˰','Ͷ��-����','���','����'};

%% ����ָ�������ԭʼ����:Ҫ����������Ԫ������Ϊ����
data = xlsread(filename,input_sheet,range);
data(isnan(data)) = 0.0;   % ����ֵ��Ϊ0

%% ����Ҫ����ϲ�����
N = length(cell2mat(col)); % �ϲ�ǰ��Ʒ/���
T = length(col);           % �ϲ�����Ʒ/���
M = size(data,1) - 2*N;    % ������Ŀ�����Ͷ�������ҵ�������ȣ�
sub_a = zeros(T, N);
for i =1:T
    sub_a(i,col{i}) = 1;
end
% ����Ϊ�����Ʒ&��ľ����
sub_a = [sub_a zeros(size(sub_a)); zeros(size(sub_a)) sub_a];
% ��װΪ�ܵĺϲ�����
A = [sub_a zeros(2*T,M);zeros(M,2*N) eye(M)];

%% �ϲ�����
ydata = A*data*A';

%% ����� excel
% 1 �ϲ����SAM��
% ��ͷ
sam_title = cell(1,2*T+M);
for i = 1:length(sam_title)
    sam_title{i} = [label, num2str(i)]; 
end
xlswrite(filename, sam_title, output1_sheet,'B1');  % д�б�ͷ
xlswrite(filename, sam_title', output1_sheet,'A2'); % д�б�ͷ
xlswrite(filename, ydata, output1_sheet,'B2');      % д����
% 2 ΪGAMS׼���Ĵ����ƽ�����ݵĹ�����
xlswrite(filename, sam_title, output2_sheet,'B1');  % д�б�ͷ
xlswrite(filename, sam_title', output2_sheet,'A2'); % д�б�ͷ
% 3 ���չ���������ʵ�ʱ�ͷ��
real_title = cell(1,2*T+M);
for i = 1:T
    real_title{i} = title1{i}; 
end
for i = 1:T
    real_title{T+i} = title1{i}; 
end
for i = 1:M
    real_title{2*T+i} = title2{i}; 
end
xlswrite(filename, real_title, final_sheet,'B1');  % д�б�ͷ
xlswrite(filename, real_title', final_sheet,'A2'); % д�б�ͷ