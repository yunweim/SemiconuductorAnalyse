%% This code read CSV file from B1505 and generate Excel file with data and parameters
%  Author: Yunwei Ma, Virginia Tech, 2022
%  Library: none
%  Version: Malab 2022b

%% Initialize program
clc;
clear;
%Delete name overlapping
excelname = 'B1505_CV_summary.xlsx';
excelname_raw = 'B1505_CV_data.xlsx';
if isfile(excelname)
    delete B1505_CV_summary.xlsx
end
if isfile(excelname_raw)
    delete B1505_CV_data.xlsx
end

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
SummaryFrequency=strings(LenOfFile,1);
SummaryDelay=strings(LenOfFile,1);
SummaryAC=strings(LenOfFile,1);
SummaryArea=ones(LenOfFile,1);
SummaryPerimeter=ones(LenOfFile,1);


writetable(table(),excelname_raw,'Sheet',1,'Range','A1','WriteVariableNames',false)
writetable(table(),excelname,'Sheet',1,'Range','A1','WriteVariableNames',false)
for ifile=1:LenOfFile
    opts = detectImportOptions(name_and_path{ifile});
    opts = setvartype(opts,'char');
    opts.VariableNamingRule='Preserve';
    opts.DataLines=[1,Inf];% or 'string'
    rawdata = readtable(name_and_path{ifile},opts);
    %convert file name to follow excel sheetname rule
    temp=strrep(filename{ifile},'[','$');
    temp=strrep(temp,']','$');
    if length(temp)>31
     temp=temp(1:31);
    end
    writetable(rawdata,excelname_raw,'Sheet',temp,'Range','A1','WriteVariableNames',false)
    SummaryName{ifile}=filename{ifile};
    SummaryName{ifile}=filename{ifile};
    [~,index] = ismember('Channel.Unit',rawdata{:,2});
    SummaryModule{ifile}=strcat(rawdata{index,3}{1},'_',rawdata{index,4}{1});
    [~,index] = ismember('Measurement.Primary.Start',rawdata{:,2});
    SummaryStartV{ifile}=rawdata{index,3}{1};
    [~,index] = ismember('Measurement.Primary.Stop',rawdata{:,2});
    SummaryStopV{ifile}=rawdata{index,3}{1};
    [~,index] = ismember('Measurement.Primary.Compliance',rawdata{:,2});
    SummaryCompliance{ifile}=rawdata{index,3}{1};
    [~,index] = ismember('TestRecord.RecordTime',rawdata{:,2});
    SummaryDateTime{ifile}=rawdata{index,3}{1};
    [~,index] = ismember('Measurement.Secondary.Frequency',rawdata{:,2});
    SummaryFrequency{ifile}=rawdata{index,3}{1};
    [~,index] = ismember('Measurement.Timing.Delay',rawdata{:,2});
    SummaryDelay{ifile}=rawdata{index,3}{1};
    [~,index] = ismember('Measurement.Secondary.ACLevel',rawdata{:,2});
    SummaryAC{ifile}=rawdata{index,3}{1};
    [~,index] = ismember('Channel.IName',rawdata{:,2});
    Name_I = rawdata{index,3}{1};
    [~,index] = ismember('Channel.VName',rawdata{:,2});
    Name_V = rawdata{index,3}{1};
    [~,index] = ismember('DataName',rawdata{:,1});
    data = rawdata(index+1:end,2:5);
    data.Properties.VariableNames = {'I','V','C','G'};
    data.V = str2double(data.V);
    data.I = str2double(data.I);
    %Define abs I
    words=num2str(transpose(2:height(data)+1));
    Eqn="=1/(C"+words+")^2";
    Eqn = strrep(Eqn,' ','');
    data.inv_C2=Eqn;
    %Define area 
    Eqn='=sheet1!I'+strings(height(data),1)+num2str(ifile+1);
    data.area=Eqn;
    %Define area C
    words=num2str(transpose(2:height(data)+1));
    Eqn="=C"+words+"/F"+words;
    Eqn = strrep(Eqn,' ','');
    data.areaC=Eqn;
    %Define 1/areaC^2
    words=num2str(transpose(2:height(data)+1));
    Eqn="=1/G"+words+"^2";
    Eqn = strrep(Eqn,' ','');
    data.inv_areaC2=Eqn;    
    %Write to file
    writetable(data,excelname,'Sheet',temp,'Range','A1','UseExcel',true)
    clear temp
end
Summary=table(SummaryName,SummaryModule,SummaryDateTime,SummaryStartV,SummaryStopV,SummaryFrequency,SummaryDelay,SummaryAC,SummaryArea,SummaryPerimeter);
writetable(Summary,excelname_raw,'Sheet',1,'Range','A1')
writetable(Summary,excelname,'Sheet',1,'Range','A1','UseExcel',true)
%% Find key parameters position


