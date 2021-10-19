clc;clear;
%----------��������--------------------
dataPath='./textdata1-01-12.xlsx';%�����ļ�·���������Լ��ĵ��Ը��¸�·������
resultPath='./';%������·��

%----------��ʼ����--------------------
[data,txt]=xlsread(dataPath);%���ݶ�ȡ
columnName=txt(2,2:end)';
if size(txt,1)==2
    tag=data(1,2:end);%ָ����������
    planName=num2cell(data(3:end,1));
    data=data(:,2:end);  
else
    tag=data(1,:);%ָ����������
    planName=txt(3:end,1);
end
x=[];%��׼�����ݴ洢����
%���ݱ�׼����ָ������1��������ָ�꣬-1������ָ�꣬����ֵ�����ʶ�ָ������ֵ
for i=1:length(tag)
    if tag(i)==1
        pData=data(3:end,i);
        pDataSt=mapminmax(pData',0,1)';%����ָ���׼��
        x=[x,pDataSt];
    elseif tag(i)==-1
        nData=data(3:end,i);
        nDataSt=1-mapminmax(nData',0,1)';%����ָ�������׼��
        x=[x,nDataSt];
	else
		pnData=data(3:end,i);
		pnDataSt=1-abs(pnData-tag(i))/max(abs(pnData-tag(i)));%�ʶ�ָ�������׼��
		x=[x,pnDataSt];
    end
end
[m,n]=size(x);

%�����j��ָ���µ�i������ռ��ָ��ı���
for i=1:m
    for j=1:n
   p(i,j)=x(i,j)/sum(x(:,j));
    end 
end

%��0�滻Ϊ0.001����Ȼlog(0)�����
p(p==0)=0.0001;

%�����j��ָ�����ֵ
k=1/log(m);
for i=1:m
    for j=1:n
 e0(i,j)=p(i,j)*log(p(i,j));
    end 
end
for j=1:n
  e(j)=-k*sum(e0(:,j));
end

%�����j��ָ��Ĳ���ϵ��
g=1-e;

%��Ȩ��
for j=1:n;
    w(j,1)=g(j)/sum(g);
end

%������������ۺϵ÷�

for i=1:m
    score(i,1)=-sum(e0(i,:).*w');
end

%�������
scoreCell=[[{'����'},columnName',{'��Ȩ��'}];[planName,num2cell(x),num2cell(score)]];%��׼�����������
wCell=[{'ָ��','Ȩ��'};[columnName,num2cell(w)]];%ָ��������Ȩ��
eCell=[{'ָ��','��Ϣ��'};[columnName,num2cell(e')]];%ָ��������Ȩ��

dataPathSplit= regexp(dataPath, '\', 'split');
fileName=regexp(dataPathSplit{end}, '.xls', 'split');

xlswrite(cat(2,resultPath,fileName{1},'��ֵ��������12.xlsx'),scoreCell,'��׼�����������')%����׼�����������д��excel
xlswrite(cat(2,resultPath,fileName{1},'��ֵ��������12.xlsx'),wCell,'Ȩ��')%��Ȩ��д��excel
xlswrite(cat(2,resultPath,fileName{1},'��ֵ��������12.xlsx'),eCell,'��Ϣ��')%����Ϣ��д��excel

disp(cat(2,'Finish!�������·��Ϊ:'))
disp(cat(2,resultPath,fileName{1},'��ֵ��������12.xlsx'))
%------��ͼ-------
% columns=txt(2,2:end);
 %bar(w,0.8)
 %for i = 1:length(w)
  %   text(i,w(i),num2str(w(i),'%g'),...
   %  'HorizontalAlignment','center',...
    % 'VerticalAlignment','bottom')
 %end
 %xlabel('ָ��');ylabel('Ȩ��');
 %set(gca, 'XTickLabels', columns)