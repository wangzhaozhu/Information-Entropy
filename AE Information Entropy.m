clc;clear;
warning off
%----------��������--------------------
dataPath='./textData1-01-12.xlsx';%�����ļ�·���������Լ��ĵ��Ը��¸�·������

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
 e(i,j)=-p(i,j)*log(p(i,j));
    end 
end


%�������
eCell=[['����';planName],[columnName';num2cell(e)]];%��Ϣ��

dataPathSplit= regexp(dataPath, '\', 'split');
fileName=regexp(dataPathSplit{end}, '.xls', 'split');

xlswrite(cat(2,fileName{1},'��ֵ��������-������12.xlsx'),eCell,'������')%����Ϣ��д��excel

disp(cat(2,'Finish!�������·��Ϊ:'))