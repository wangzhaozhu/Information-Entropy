clc;clear;
%----------参数设置--------------------
dataPath='./textdata1-01-12.xlsx';%数据文件路径，根据自己的电脑改下该路径即可
resultPath='./';%结果输出路径

%----------开始计算--------------------
[data,txt]=xlsread(dataPath);%数据读取
columnName=txt(2,2:end)';
if size(txt,1)==2
    tag=data(1,2:end);%指标属性数据
    planName=num2cell(data(3:end,1));
    data=data(:,2:end);  
else
    tag=data(1,:);%指标属性数据
    planName=txt(3:end,1);
end
x=[];%标准化数据存储矩阵
%数据标准化，指标属性1代表正向指标，-1代表负向指标，其它值代表适度指标的最佳值
for i=1:length(tag)
    if tag(i)==1
        pData=data(3:end,i);
        pDataSt=mapminmax(pData',0,1)';%正向指标标准化
        x=[x,pDataSt];
    elseif tag(i)==-1
        nData=data(3:end,i);
        nDataSt=1-mapminmax(nData',0,1)';%负向指标正向标准化
        x=[x,nDataSt];
	else
		pnData=data(3:end,i);
		pnDataSt=1-abs(pnData-tag(i))/max(abs(pnData-tag(i)));%适度指标正向标准化
		x=[x,pnDataSt];
    end
end
[m,n]=size(x);

%计算第j项指标下第i个方案占该指标的比重
for i=1:m
    for j=1:n
   p(i,j)=x(i,j)/sum(x(:,j));
    end 
end

%将0替换为0.001，不然log(0)会出错
p(p==0)=0.0001;

%计算第j项指标的熵值
k=1/log(m);
for i=1:m
    for j=1:n
 e0(i,j)=p(i,j)*log(p(i,j));
    end 
end
for j=1:n
  e(j)=-k*sum(e0(:,j));
end

%计算第j项指标的差异系数
g=1-e;

%求权重
for j=1:n;
    w(j,1)=g(j)/sum(g);
end

%计算各方案的综合得分

for i=1:m
    score(i,1)=-sum(e0(i,:).*w');
end

%结果导出
scoreCell=[[{'方案'},columnName',{'加权熵'}];[planName,num2cell(x),num2cell(score)]];%标准化数据与分数
wCell=[{'指标','权重'};[columnName,num2cell(w)]];%指标名称与权重
eCell=[{'指标','信息熵'};[columnName,num2cell(e')]];%指标名称与权重

dataPathSplit= regexp(dataPath, '\', 'split');
fileName=regexp(dataPathSplit{end}, '.xls', 'split');

xlswrite(cat(2,resultPath,fileName{1},'熵值法计算结果12.xlsx'),scoreCell,'标准化数据与分数')%将标准化数据与分数写入excel
xlswrite(cat(2,resultPath,fileName{1},'熵值法计算结果12.xlsx'),wCell,'权重')%将权重写入excel
xlswrite(cat(2,resultPath,fileName{1},'熵值法计算结果12.xlsx'),eCell,'信息熵')%将信息熵写入excel

disp(cat(2,'Finish!结果所在路径为:'))
disp(cat(2,resultPath,fileName{1},'熵值法计算结果12.xlsx'))
%------画图-------
% columns=txt(2,2:end);
 %bar(w,0.8)
 %for i = 1:length(w)
  %   text(i,w(i),num2str(w(i),'%g'),...
   %  'HorizontalAlignment','center',...
    % 'VerticalAlignment','bottom')
 %end
 %xlabel('指标');ylabel('权重');
 %set(gca, 'XTickLabels', columns)