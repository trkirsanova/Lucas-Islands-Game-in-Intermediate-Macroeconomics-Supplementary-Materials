%% Collect Pe(C5:C19), W(B5:B19), Is(B1) and export combined.xlsx (with Is letter+code)

folder = uigetdir(pwd, 'Select the folder with student Excel files');
if isequal(folder,0); error('No folder selected.'); end

rangeStrPe = 'C5:C19';
rangeStrW  = 'B5:B19';
rangeStrIs = 'B1';
sheetName  = 1;   % use 1st sheet; set to [] or 'Sheet1' if needed

%% Gather files (sorted for reproducibility)
exts = ["*.xlsx","*.xlsm","*.xls"];
files = [];
for e = exts
    files = [files; dir(fullfile(folder, e))]; %#ok<AGROW>
end
if isempty(files), error('No Excel files found in %s', folder); end
[~, idx] = sort({files.name}); files = files(idx);

%% Containers
PeCols = {};  WCols = {};  IDs = [];  FN = strings(0,1);
IsLetters = {};                     % cellstr column (char), e.g., {'A';'B';...}
IsCodes   = [];                     % numeric column (A->1,...,J->10)

nRowsPe = 15; nRowsW = 15;
validLetters = 'A':'J';

id = 0;
for k = 1:numel(files)
    fpath = fullfile(files(k).folder, files(k).name);
    try
        if isempty(sheetName)
            Pe = readmatrix(fpath, 'Range', rangeStrPe);
            W  = readmatrix(fpath, 'Range', rangeStrW);
            IsCell = readcell( fpath, 'Range', rangeStrIs);
        else
            Pe = readmatrix(fpath, 'Sheet', sheetName, 'Range', rangeStrPe);
            W  = readmatrix(fpath, 'Sheet', sheetName, 'Range', rangeStrW);
            IsCell = readcell( fpath, 'Sheet', sheetName, 'Range', rangeStrIs);
        end

        % Coerce numeric and pad/trim
        Pe = double(Pe(:));  if numel(Pe) < nRowsPe, Pe(end+1:nRowsPe,1) = NaN; end; Pe = Pe(1:nRowsPe);
        W  = double(W(:));   if numel(W)  < nRowsW,  W(end+1:nRowsW,1)  = NaN; end; W  = W(1:nRowsW);

        % Island letter from B1 -> clean char
        if iscell(IsCell), IsVal = IsCell{1}; else, IsVal = IsCell; end
        if isempty(IsVal) || (isstring(IsVal) && strlength(IsVal)==0) || (ischar(IsVal) && all(isspace(IsVal)))
            islChar = '';
        else
            islChar = upper(char(string(IsVal)));
            islChar = strtrim(islChar);
            if ~isempty(islChar), islChar = islChar(1); end   % take first char if longer
            if ~ismember(islChar, validLetters), islChar = ''; end
        end

        % Map letter to code A=1,...,J=10 (else NaN)
        if isempty(islChar)
            islCode = NaN;
        else
            islCode = double(islChar) - double('A') + 1;
        end

        % Accept this file
        id = id + 1;
        IDs(end+1,1)       = id;            %#ok<AGROW>
        PeCols{end+1}      = Pe;            %#ok<AGROW>
        WCols{end+1}       = W;             %#ok<AGROW>
        IsLetters{end+1,1} = islChar;       %#ok<AGROW>  % ensure column cellstr
        IsCodes(end+1,1)   = islCode;       %#ok<AGROW>
        FN(end+1,1)        = string(files(k).name); %#ok<AGROW>

    catch ME
        warning('Skipping %s: %s', files(k).name, ME.message);
    end
end

nOK = numel(IDs);
if nOK == 0, error('No readable files in %s', folder); end

%% Build numeric matrices
PeMat = NaN(nRowsPe, nOK);  WMat = NaN(nRowsW, nOK);
for j = 1:nOK, PeMat(:,j) = PeCols{j}; WMat(:,j) = WCols{j}; end

% Headers/labels
topHeader = [{' '}, num2cell(IDs.')];                   % 1 × (nOK+1)
rowLabels = arrayfun(@(x) sprintf('r%02d', x), 5:19, 'UniformOutput', false)'; 

% Assemble outputs as cell arrays (mixed text + numbers)
PeOut = [topHeader; [rowLabels, num2cell(PeMat)]];
WOut  = [topHeader; [rowLabels, num2cell(WMat)]];

% Island sheet: ID | IsLetter | IsCode
IsOut = [{'ID','IsLetter','IsCode'}; num2cell(IDs), IsLetters, num2cell(IsCodes)];

% Mapping table
MapTbl = table(IDs, FN, 'VariableNames', {'ID','OriginalFilename'});

%% Write
folder_write = folder; % change here
outFile = fullfile(folder_write, 'combined.xlsx');
writecell(PeOut, outFile, 'Sheet', 'Pe');
writecell(WOut,  outFile, 'Sheet', 'Wi');
writecell(IsOut, outFile, 'Sheet', 'Is');
writetable(MapTbl, outFile, 'Sheet', 'Mapping', 'WriteVariableNames', true);

fprintf('Done. Wrote %s with sheets: Pe, W, Is (letter+code), Mapping. IDs 1..%d\n', outFile, max(IDs));
