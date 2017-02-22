%% 合并的列（注意开始行为02）
% 01、农业（02 ）
% 02、煤炭（03） 
% 03、石油（04）
% 04、天然气（05） 
% 05、轻工业（07-10）
% 06、重工业（06+11-18） 
% 07、火电（19）
% 08、水电、核电等其他（20）
% 09、燃气生产（21）
% 10、水的生产（22）
% 11、建筑业（23）
% 12、交通运输及仓储（24）
% 13、住宿和餐饮业（26）
% 14、其他服务业（25+27）
% 20160517
% 煤 7、41   
% 轻工业：13-39 45、48、49、50、70、84、85、94、99   
% 重工业：9-12、42-44、46、47、51-69、71-83、86-93、95、96  
% 建筑： 100-103
% 交通：105-109、111、112
% 第三产业：104、110、113-140  
clear;
%% 数据准备
filename = '全国2012年.xlsx';     % 读入数据的excel文件名
input_sheet = 'sheet1';    % 读入数据的工作表名
range = 'B2:KB288';         % 读入数据范围
output1_sheet = 'sheet2';  % 输出数据的工作表名
output2_sheet = 'sheet3';  % 存放配平后数据的工作表名
final_sheet = 'sheet4';    % 存放配平后数据的工作表名
label = 'data';            % 自定义数据表头，但是需要与GAMS的set一致
% 合并规则
col = {1:5,7,39,96,97,[6,40],[12:38,44,47:49,69,83,84,93,98],[8:11,41:43,45,46,50:68,70:82,85:92,94,95],99:102,[104:108,110,111],[103,109,112:139]}; 
% 合并后的项目
title1 = {'农林牧渔业','石油和天然气开采产品','精炼石油和核燃料加工品','电力、热力生产和供应','燃气生产和供应','煤','轻工业','重工业','建筑','交通','第三产业'};   
title2 = {'劳动','资本','居民','企业','政府','碳税','投资-储蓄','存货','国外'};

%% 读入指定区域的原始数据:要求数据区域单元格类型为数字
data = xlsread(filename,input_sheet,range);
data(isnan(data)) = 0.0;   % 将空值设为0

%% 根据要求构造合并矩阵
N = length(cell2mat(col)); % 合并前商品/活动数
T = length(col);           % 合并后商品/活动数
M = size(data,1) - 2*N;    % 其他项目数（劳动力、企业、政府等）
sub_a = zeros(T, N);
for i =1:T
    sub_a(i,col{i}) = 1;
end
% 扩充为针对商品&活动的矩阵块
sub_a = [sub_a zeros(size(sub_a)); zeros(size(sub_a)) sub_a];
% 组装为总的合并矩阵
A = [sub_a zeros(2*T,M);zeros(M,2*N) eye(M)];

%% 合并矩阵
ydata = A*data*A';

%% 输出到 excel
% 1 合并后的SAM表
% 表头
sam_title = cell(1,2*T+M);
for i = 1:length(sam_title)
    sam_title{i} = [label, num2str(i)]; 
end
xlswrite(filename, sam_title, output1_sheet,'B1');  % 写行表头
xlswrite(filename, sam_title', output1_sheet,'A2'); % 写列表头
xlswrite(filename, ydata, output1_sheet,'B2');      % 写数据
% 2 为GAMS准备的存放配平后数据的工作表
xlswrite(filename, sam_title, output2_sheet,'B1');  % 写行表头
xlswrite(filename, sam_title', output2_sheet,'A2'); % 写列表头
% 3 最终工作表（带有实际表头）
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
xlswrite(filename, real_title, final_sheet,'B1');  % 写行表头
xlswrite(filename, real_title', final_sheet,'A2'); % 写列表头