%% This code read CSV file from B1505 and generate Excel file with data and parameters
%  Author: Yunwei Ma, Virginia Tech, 2022
%  Library: none
%  Version: Malab 2022b

%% Initialize program
clc;
clear;
%% Read file names and paths
%open file dialog
[filename,path] = uigetfile('*.*','MultiSelect', 'on');

% If only 1 file selected, still put into cell with length=1
if isa(filename,'double')
    disp('no file selected');
    return
end
if isa(filename,'char')
    LenOfFile=1;
    temp=filename;
    filename={1};
    name_and_path={1};
    filename{1}=temp;
    name_and_path{1} = fullfile(path, filename{1});
    clear temp temp2;
else
    LenOfFile=length(filename);
    name_and_path = fullfile(path, filename);
end
%% Read each .CSV file

SummaryName=strings(LenOfFile,1);
SummaryModule=strings(LenOfFile,1);
SummaryStartV=strings(LenOfFile,1);
SummaryStopV=strings(LenOfFile,1);
SummaryCompliance=strings(LenOfFile,1);
SummaryDateTime=strings(LenOfFile,1);
SummaryArea=ones(LenOfFile,1);
SummaryPerimeter=ones(LenOfFile,1);
SummaryData=cell(LenOfFile,1);


%writetable(table(),excelname_raw,'Sheet',1,'Range','A1','WriteVariableNames',false)
%writetable(table(),excelname,'Sheet',1,'Range','A1','WriteVariableNames',false)
for ifile=1:LenOfFile
    opts = detectImportOptions(name_and_path{ifile});
    opts = setvartype(opts,'char');
    opts.VariableNamingRule='Preserve';
    opts.DataLines=[1,Inf];% or 'string'
    rawdata = readtable(name_and_path{ifile},opts);
    %convert file name to follow excel sheetname rule
    temp=strrep(filename{ifile},'[','$');
    temp=strrep(temp,']','$');
    temp=temp(1:31);
    %writetable(rawdata,excelname_raw,'Sheet',temp,'Range','A1','WriteVariableNames',false)
    SummaryName{ifile}=filename{ifile};
    SummaryName{ifile}=filename{ifile};
    [~,index] = ismember('Measurement.Port.Unit',rawdata{:,2});
    SummaryModule{ifile}=rawdata{index,3}{1};
    [~,index] = ismember('Measurement.Primary.Start',rawdata{:,2});
    SummaryStartV{ifile}=rawdata{index,3}{1};
    [~,index] = ismember('Measurement.Primary.Stop',rawdata{:,2});
    SummaryStopV{ifile}=rawdata{index,3}{1};
    [~,index] = ismember('Measurement.Primary.Compliance',rawdata{:,2});
    SummaryCompliance{ifile}=rawdata{index,3}{1};
    [~,index] = ismember('TestRecord.RecordTime',rawdata{:,2});
    SummaryDateTime{ifile}=rawdata{index,3}{1};
    [~,index] = ismember('Channel.IName',rawdata{:,2});
    Name_I = rawdata{index,3}{1};
    [~,index] = ismember('Channel.VName',rawdata{:,2});
    Name_V = rawdata{index,3}{1};
    [~,index] = ismember('DataName',rawdata{:,1});
    LocDataName=index;
    [~,rowV] = ismember(Name_V,rawdata{LocDataName,:});
    [~,rowI] = ismember(Name_I,rawdata{LocDataName,:});
    V = table2array(rawdata(LocDataName+1:end,rowV));
    I = table2array(rawdata(LocDataName+1:end,rowI));
    data=table(V,I);
    data.V = str2double(data.V);
    data.I = str2double(data.I);
    %Define abs I
    words=num2str(transpose(2:height(data)+1));
    Eqn="=abs(B"+words+")";
    Eqn = strrep(Eqn,' ','');
    data.abs_I=Eqn;
    %Define area
    Eqn=strrep("=C"+words,' ','');
    Eqn=Eqn+"/sheet1!F"+num2str(ifile+1);
    data.area_I=Eqn;
    %writetable(data,excelname,'Sheet',temp,'Range','A1','UseExcel',true)
    SummaryData{ifile,1}=data;
    hold on
    %plot(data.V,data.I,'DisplayName',filename{ifile});
    clear temp
end
Summary=table(SummaryName,SummaryModule,SummaryDateTime,SummaryStartV,SummaryStopV,SummaryArea,SummaryPerimeter);
SummaryData=cell2table(SummaryData,"VariableNames","Data");
Summary=[Summary SummaryData];

save("result.mat","Summary","filename")
%% Plot datas
%set normalization 
LenOfFile=height(Summary);
hold on
for ifile=1:LenOfFile
    plot(Summary.Data{ifile,1}.V,abs(1e8*Summary.Data{ifile,1}.I/Summary.SummaryArea(ifile)),'DisplayName',filename{ifile});
end
set(gca, 'YScale', 'log')


