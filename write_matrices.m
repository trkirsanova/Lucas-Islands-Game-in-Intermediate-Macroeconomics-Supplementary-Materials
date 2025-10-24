%% Append Wmatr, Pagg, Nb, Nr, Ei, Eu, Mc, Me to an existing Excel file
% Writes sheets:
%   - W_table:   columns [t, 1..N] from Wmatr (T x N)
%   - P:         columns [t, P]     from Pagg   (T x 1)
%   - Nb, Nr:    columns [t, 1..N]  from Nb/Nr  (T x N)
%   - Ei, Eu:    columns [t, Ei] / [t, Eu]  (T x 1)
%   - Mc, Me:    columns [t, Mc] / [t, Me]  (T x 1)

% ---- Your inputs (as provided) ------------------------------------------


Wmatr = [10	10	10	10	10	10	10	10	10	10
9	12	9	10	9	12	10	9	10	10
8.5	11	9	9.5	8.5	11	9.5	9	9.5	9.5
8	10	9	9	8	10	9	9	9	9
9	11	10	10	9	11	10	10	10	10
10	11	10.5	10.5	10	11	10.5	10.5	10.5	10.5
10	11	10.5	10.5	10	11	10.5	10.5	10.5	10.5
10.5	10.5	10.5	10.5	10.5	10.5	10.5	10.5	10.5	10.5
8.5	10.5	11	11	8.5	11	9	10.5	11	9
9	10.5	10.5	10.5	9	10.5	9.5	10.5	11	9
9	9.5	10	9.5	9	10	9	10	10	9
9.5	9.5	9.5	9.5	9.5	10	9	9.5	10	9
9.5	9.5	9.5	9.5	9.5	9.5	9.5	9.5	9.5	9.5
10	10	9.5	10	10	10	10	10.5	10	10
10.5	10.5	10.5	10.5	10	10.5	10.5	10.5	11	10.5];

Pagg = [1
1
0.95
0.9
1
1.05
1.05
1.05
1
1
0.95
0.95
0.95
1
1.05];

Nb = [0	0	0	0	0	0	0	0	0	0
0	1	0	0	0	1	0	0	0	0
0	0	0	0	0	0	0	0	0	0
0	0	0	0	0	0	0	0	0	0
0	0	0	0	0	0	0	0	0	0
0	0	0	0	0	0	0	0	0	0
0	0	0	0	0	0	0	0	0	0
0	0	0	0	0	0	0	0	0	0
0	0	0	0	0	0	0	0	0	0
0	0	0	0	0	0	0	0	0	0
0	0	0	0	0	0	0	0	0	0
0	0	0	0	0	0	0	0	0	0
0	0	0	0	0	0	0	0	0	0
0	0	0	0	0	0	0	0	0	0
0	0	0	0	0	0	0	0	0	0];

Nr = [0	0	0	0	0	0	0	0	0	0
0	0	0	0	0	0	0	0	0	0
0	0	0	0	0	0	0	0	0	0
0	0	0	0	0	0	0	0	0	0
0	0	0	0	0	0	0	0	0	0
0	0	0	0	0	0	0	0	0	0
0	0	0	0	0	0	0	0	0	0
0	0	0	0	0	0	0	0	0	0
1	0	0	0	1	0	1	0	0	1
0	0	0	0	0	0	0	0	0	0
0	0	0	0	0	0	0	0	0	0
0	0	0	0	0	0	0	0	0	0
0	0	0	0	0	0	0	0	0	0
0	0	0	0	0	0	0	0	0	0
0	0	0	0	0	0	0	0	0	0];

Ei = [0
0
0
0
1
1
0
0
0
0
0
0
0
1
1];

Eu = [0
0
1
1
0
0
0
1
1
0
1
1
0
0
0];

Mc = [0
0
0
0
0
1
0
0
0
0
0
0
0
0
0];

Me = [0
0
0
0
0
0
0
1
1
0
0
1
0
0
0];






% --- Settings ---
sheetNameW  = 'W_table';  % islands' wage
sheetNameP  = 'P';        % aggregate price
sheetNameNb = 'Nb';       % demand boom in island
sheetNameNr = 'Nr';       % demand reduction in island
sheetNameEi = 'Ei';       % inflation is up (indicator over time)
sheetNameEu = 'Eu';       % unemployment is up (indicator over time)
sheetNameMc = 'Mc';       % contractionary monetary policy
sheetNameMe = 'Me';       % expansionary monetary policy

tStart      = 1;            % time starts at 1
% ----------------------------------------------------

% Validate dimensions
[nT_W, nIslands] = size(Wmatr);
nT_P = numel(Pagg);
if nT_P ~= nT_W
    error('Time dimension mismatch: rows(Wmatr)=%d, length(Pagg)=%d', nT_W, nT_P);
end

% Validate Nb/Nr sizes (T x N)
if ~isequal(size(Nb), [nT_W, nIslands])
    error('Nb must be %d x %d (got %d x %d).', nT_W, nIslands, size(Nb,1), size(Nb,2));
end
if ~isequal(size(Nr), [nT_W, nIslands])
    error('Nr must be %d x %d (got %d x %d).', nT_W, nIslands, size(Nr,1), size(Nr,2));
end

% Make Ei/Eu/Mc/Me column vectors of length T
Ei = Ei(:); Eu = Eu(:); Mc = Mc(:); Me = Me(:);
if any([numel(Ei) numel(Eu) numel(Mc) numel(Me)] ~= nT_W)
    error('Ei/Eu/Mc/Me must each have %d rows (time periods).', nT_W);
end


% Build headers and tables (as cell arrays so we can mix text & numbers)
t = (tStart:(tStart+nT_W-1)).';
islandCodes = 1:nIslands;

W_out = [{'t'}, num2cell(islandCodes); num2cell(t), num2cell(Wmatr)];
P_out = [{'t','P'}; num2cell(t), num2cell(Pagg)];
Nb_out = [{'t'}, num2cell(islandCodes); num2cell(t), num2cell(Nb)];
Nr_out = [{'t'}, num2cell(islandCodes); num2cell(t), num2cell(Nr)];
Ei_out = [{'t','Ei'}; num2cell(t), num2cell(Ei)];
Eu_out = [{'t','Eu'}; num2cell(t), num2cell(Eu)];
Mc_out = [{'t','Mc'}; num2cell(t), num2cell(Mc)];
Me_out = [{'t','Me'}; num2cell(t), num2cell(Me)];


% Pick the Excel file to append (e.g., combined.xlsx from the previous script)
[xf, xp] = uigetfile('*.xlsx', 'Select Excel file to append (e.g., combined.xlsx)');
if isequal(xf,0)
    error('No file selected.');
end
outFile = fullfile(xp, xf);

% Write (overwrites sheets with same names)
writecell(W_out,  outFile, 'Sheet', sheetNameW);
writecell(P_out,  outFile, 'Sheet', sheetNameP);
writecell(Nb_out, outFile, 'Sheet', sheetNameNb);
writecell(Nr_out, outFile, 'Sheet', sheetNameNr);
writecell(Ei_out, outFile, 'Sheet', sheetNameEi);
writecell(Eu_out, outFile, 'Sheet', sheetNameEu);
writecell(Mc_out, outFile, 'Sheet', sheetNameMc);
writecell(Me_out, outFile, 'Sheet', sheetNameMe);

fprintf(['Appended to %s:\n',...
    '  - "%s": wage table (t, islands 1..%d)\n',...
    '  - "%s": aggregate price (t,P)\n',...
    '  - "%s": demand booms per island (t, islands)\n',...
    '  - "%s": demand reductions per island (t, islands)\n',...
    '  - "%s": inflation-up indicator (t,Ei)\n',...
    '  - "%s": unemployment-up indicator (t,Eu)\n',...
    '  - "%s": contractionary monetary policy (t,Mc)\n',...
    '  - "%s": expansionary monetary policy (t,Me)\n'], ...
    outFile, sheetNameW, nIslands, sheetNameP, sheetNameNb, sheetNameNr, ...
    sheetNameEi, sheetNameEu, sheetNameMc, sheetNameMe);