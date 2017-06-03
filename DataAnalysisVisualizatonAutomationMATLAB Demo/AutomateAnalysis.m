%% Automating Solar Data Analysis
% Once we have interatively created our analysis, it is useful to now
% automate it in the form of a function so we can do similar analysis on
% a variety of locations and compare the results

% Copyright 2012-2016 The MathWorks, Inc.

%% Pick a Folder of Data
% To select a folder interactively, uncomment the first line and comment
% out the second line.

% folderName = uigetdir();
folderName = fullfile(pwd,'SolarData');
allFileNames = dir(fullfile(folderName,'*Daily.xlsx'));

%% Run Analysis on all Excel Files in Folder

% Preallocating variables
coeffs = zeros(length(allFileNames),4);
stationName = cell(length(allFileNames),1);

tic
% Looping through all Excel Files
for ii = 1:length(allFileNames)
   
    % Display file processing
    disp(allFileNames(ii).name)
    
    % Create full file name and extract name of station
    fileName = fullfile(folderName,allFileNames(ii).name);
    [~, stationName{ii,1}, ~] = fileparts(allFileNames(ii).name);
    
    % Run Analysis Function
    model = SolarAnalysisFcn(fileName);
    
    % Extract Coefficents
    coeffs(ii,:) = coeffvalues(model);
end
toc

%% Write Results to an Excel File

xlswrite('SolarFitResults.xlsx',stationName,'Sheet1','A1');
xlswrite('SolarFitResults.xlsx',coeffs,'Sheet1','B1');
