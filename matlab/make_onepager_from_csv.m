%% make_onepager_from_csv.m  (FIXED)
% Reads CSV exports and creates a 1-page (A4 landscape) summary (PNG + PDF).
% Expected repo structure:
%   excel/THERMAL-ANALYSIS2.csv   (WINTER)
%   excel/THERMAL-ANALYSIS.csv    (SUMMER)
%   outputs/                      (generated files saved here)

clear; clc;

% --- Paths (script lives in /matlab)
thisDir  = fileparts(mfilename('fullpath'));
repoRoot = fullfile(thisDir, '..');

winterCsv = fullfile(repoRoot, 'excel', 'THERMAL-ANALYSIS2.csv'); % Winter
summerCsv = fullfile(repoRoot, 'excel', 'THERMAL-ANALYSIS.csv');  % Summer
outDir    = fullfile(repoRoot, 'outputs');

assert(exist(winterCsv,'file')==2, "Missing: %s", winterCsv);
assert(exist(summerCsv,'file')==2, "Missing: %s", summerCsv);
if ~exist(outDir,'dir'), mkdir(outDir); end

% --- Read CSVs as cells (robust with mixed text/numbers)
W = readcell(winterCsv, 'Delimiter', ',', 'TextType', 'string');
S = readcell(summerCsv, 'Delimiter', ',', 'TextType', 'string');

%% =======================
% WINTER extraction
% =======================

% Walls
rWalls = findRowFirstColContains(W, "Walls:");
[rH, cQ] = findHeaderInSection(W, rWalls, "Q_loss");
wallRows = collectDataRows(W, rH+1, 1, 10);
wWallVals = cellfun(@(r) toNum(W{r,cQ}), num2cell(wallRows));
wWalls = sum(wWallVals, 'omitnan');

% Windows
rWin = findRowFirstColContains(W, "Windows:");
[rH, cQ] = findHeaderInSection(W, rWin, "Q_loss");
rData = firstNumericRow(W, rH+1, cQ, 10);
wWindows = toNum(W{rData,cQ});

% Floor
rFloor = findRowFirstColContains(W, "Floor:");
[rH, cQ] = findHeaderInSection(W, rFloor, "Q_loss");
rData = firstNumericRow(W, rH+1, cQ, 10);
wFloor = toNum(W{rData,cQ});

% Ceiling
rCeil = findRowFirstColContains(W, "Ceiling:");
[rH, cQ] = findHeaderInSection(W, rCeil, "Q_loss");
rData = firstNumericRow(W, rH+1, cQ, 10);
wCeiling = toNum(W{rData,cQ});

% Air exchange + totals
wAir           = valueRightOfLabel(W, "Q_tot,loss (by air exchange)");
wTransferTotal = valueRightOfLabel(W, "Q_tot,loss (by heat transfer)");
wTotal         = valueRightOfLabelLast(W, "Q_tot,loss");

% Assumptions (winter)
wACH  = valueRightOfLabel(W, "n (ACH)");
wVol  = valueRightOfLabel(W, "Volume of heated space (m³)");
wTout = valueRightOfLabel(W, "External Temperature (°C)");
wTin  = valueRightOfLabel(W, "Internal Temperature (°C)");

winterLabels = {'Walls','Windows','Floor','Ceiling','Air exchange'};
winterVals_W = [wWalls, wWindows, wFloor, wCeiling, wAir];

%% =======================
% SUMMER extraction
% =======================

% Walls
rWalls = findRowFirstColContains(S, "Walls:");
[rH, cQwall] = findHeaderInSection(S, rWalls, "Q_wall");
wallRows = collectDataRows(S, rH+1, 1, 10);
sWallVals = cellfun(@(r) toNum(S{r,cQwall}), num2cell(wallRows));
sWalls = sum(sWallVals, 'omitnan');

% Floor
rFloor = findRowFirstColContains(S, "Floor:");
[rH, cQfloor] = findHeaderInSection(S, rFloor, "Q_floor");
rData = firstNumericRow(S, rH+1, cQfloor, 10);
sFloor = toNum(S{rData,cQfloor}); % can be negative

% Windows
rWin = findRowFirstColContains(S, "Windows:");
[rH, cQwin] = findHeaderInSection(S, rWin, "Q_win");
[~,  cQsol] = findHeaderInSection(S, rWin, "Q_solar");
rData = firstRowFirstColContainsBelow(S, rH+1, "Wall", 10);
sWinTotal = toNum(S{rData,cQwin});
sWinSolar = toNum(S{rData,cQsol});
sWinCond  = sWinTotal - sWinSolar;

% Ceiling
rCeil = findRowFirstColContains(S, "Ceiling:");
[rH, cQceil] = findHeaderInSection(S, rCeil, "Q_ceiling");
rData = firstNumericRow(S, rH+1, cQceil, 10);
sCeiling = toNum(S{rData,cQceil});

% People (FIX: read value UNDER the 'Q_people' column)
rPeople = findRowFirstColContains(S, "People:");
[rH, cQpeople] = findHeaderInSection(S, rPeople, "Q_people");
rData = firstNumericRow(S, rH+1, cQpeople, 10);
sPeople = toNum(S{rData,cQpeople});

% Lighting (FIX: read value UNDER the 'Q_light' column)
rLight = findRowAnyColContains(S, "Heat Exchange by Lighting");
rHdrL  = findRowInRangeAnyColContains(S, rLight, 12, "Q_light");
cQlight = findColInRowContains(S, rHdrL, "Q_light");
rData  = firstNumericRow(S, rHdrL+1, cQlight, 10);
sLight = toNum(S{rData,cQlight});

% Summer total
sTotal = valueRightOfLabel(S, "Q_gain_tot");

% Summer assumptions (from Windows row)
rWin = findRowFirstColContains(S, "Windows:");
rHdr = findRowInRangeAnyColContains(S, rWin, 12, "I-value");
cTout = findColInRowContains(S, rHdr, "External Temperature");
cTin  = findColInRowContains(S, rHdr, "Internal Temperature");
cI    = findColInRowContains(S, rHdr, "I-value");
cTau  = findColInRowContains(S, rHdr, "τ-value");
rData = firstRowFirstColContainsBelow(S, rHdr+1, "Wall", 10);
sTout = toNum(S{rData,cTout});
sTin  = toNum(S{rData,cTin});
sI    = toNum(S{rData,cI});
sTau  = toNum(S{rData,cTau});

% Ground temperature from Floor table
rFloor = findRowFirstColContains(S, "Floor:");
rHdrF  = findRowInRangeAnyColContains(S, rFloor, 10, "External (ground) Temperature");
cTg    = findColInRowContains(S, rHdrF, "External (ground) Temperature");
rDataF = firstNumericRow(S, rHdrF+1, cTg, 10);
sTground = toNum(S{rDataF,cTg});

summerLabels = {'Walls','Windows','Ceiling','People','Lighting','Floor (ground)'};
summerVals_W = [sWalls, sWinTotal, sCeiling, sPeople, sLight, sFloor];

%% =======================
% One-page plot
% =======================

fig = figure('Color','w');
set(fig,'Units','inches','Position',[0.3 0.3 11.69 8.27]); % A4 landscape
t = tiledlayout(3,2,'TileSpacing','compact','Padding','compact');

axTitle = nexttile(t,[1 2]);
axis(axTitle,'off');
text(0,0.80,'Building Thermal Balance — Winter Loss vs Summer Gains','FontSize',16,'FontWeight','bold');
text(0,0.52,'Source: CSV exports (winter + summer)','FontSize',10);

axW = nexttile(t);
bar(axW, categorical(winterLabels), winterVals_W/1000);
grid(axW,'on'); ylabel(axW,'kW');
title(axW,'Winter — Heat Loss Breakdown');
xtickangle(axW,20);

axS = nexttile(t);
base  = [sWalls, sWinCond, sCeiling, sPeople, sLight, sFloor] / 1000;
solar = [0,      sWinSolar, 0,        0,       0,      0]      / 1000;
bar(axS, categorical(summerLabels), [base(:), solar(:)], 'stacked');
grid(axS,'on'); ylabel(axS,'kW');
title(axS,'Summer — Heat Gain Breakdown (Windows split: conduction + solar)');
xtickangle(axS,20);
legend(axS, {'Base','Solar (windows)'}, 'Location','best');

axInfo = nexttile(t,[1 2]);
axis(axInfo,'off');

w_total_pos = sum(max(winterVals_W,0),'omitnan');
w_share_air = 100*(wAir/w_total_pos);

summerPos = max([sWalls, sWinTotal, sCeiling, sPeople, sLight],0);
s_sumPos  = sum(summerPos,'omitnan');
s_share_win   = 100*(sWinTotal/s_sumPos);
s_share_solar = 100*(sWinSolar/s_sumPos);

winterInfo = sprintf([ ...
    'WINTER assumptions: Tin=%.1f C | Tout=%.1f C | ACH=%.2f 1/h | V=%.1f m^3\n' ...
    'WINTER totals: Transfer=%.2f kW | Air exchange=%.2f kW | Total=%.2f kW (Air exchange = %.0f%%)\n' ...
    ], ...
    wTin, wTout, wACH, wVol, ...
    wTransferTotal/1000, wAir/1000, wTotal/1000, w_share_air);

summerInfo = sprintf([ ...
    'SUMMER assumptions: Tin=%.1f C | Tout=%.1f C | Tground=%.1f C | I=%.0f W/m^2 | tau=%.3f\n' ...
    'SUMMER total (net): %.2f kW (Windows = %.0f%% of positive gains; Solar = %.0f%%)\n' ...
    'Note: Floor term may be negative (cooling contribution).\n' ...
    ], ...
    sTin, sTout, sTground, sI, sTau, ...
    sTotal/1000, s_share_win, s_share_solar);
    
text(0,0.78,winterInfo,'FontSize',10);
text(0,0.42,summerInfo,'FontSize',10);

pngOut = fullfile(outDir,'thermal_balance_onepager.png');
pdfOut = fullfile(outDir,'thermal_balance_onepager.pdf');

try
    exportgraphics(fig, pngOut, 'Resolution', 300);
    exportgraphics(fig, pdfOut);
catch
    print(fig, pngOut, '-dpng', '-r300');
    print(fig, pdfOut, '-dpdf', '-bestfit');
end

fprintf("Saved:\n  %s\n  %s\n", pngOut, pdfOut);

%% =======================
% Helpers
% =======================
function r = findRowFirstColContains(C, pattern)
    col1 = string(C(:,1));
    r = find(contains(lower(col1), lower(pattern)), 1, 'first');
    if isempty(r), error("Could not find section '%s' in first column.", pattern); end
end

function r = findRowAnyColContains(C, pattern)
    S = lower(string(C));
    r = find(any(contains(S, lower(pattern)),2), 1, 'first');
    if isempty(r), error("Could not find '%s' in any column.", pattern); end
end

function [rHdr, cKey] = findHeaderInSection(C, rStart, key)
    rHdr = findRowInRangeAnyColContains(C, rStart, 12, key);
    cKey = findColInRowContains(C, rHdr, key);
end

function r = findRowInRangeAnyColContains(C, rStart, maxRows, pattern)
    rEnd = min(size(C,1), rStart+maxRows);
    block = lower(string(C(rStart:rEnd,:)));
    idx = find(any(contains(block, lower(pattern)),2), 1, 'first');
    if isempty(idx), error("Header '%s' not found near row %d.", pattern, rStart); end
    r = rStart + idx - 1;
end

function c = findColInRowContains(C, r, pattern)
    row = lower(string(C(r,:)));
    c = find(contains(row, lower(pattern)), 1, 'first');
    if isempty(c), error("Column containing '%s' not found in row %d.", pattern, r); end
end

function rows = collectDataRows(C, rStart, nameCol, maxRows)
    rows = [];
    rEnd = min(size(C,1), rStart+maxRows-1);
    for r = rStart:rEnd
        v = string(C{r,nameCol});
        if strlength(strtrim(v))==0 || ismissing(v), break; end
        if contains(v, ":"), break; end
        rows(end+1) = r; %#ok<AGROW>
    end
end

function r = firstNumericRow(C, rStart, c, maxRows)
    rEnd = min(size(C,1), rStart+maxRows-1);
    for rr = rStart:rEnd
        if ~isnan(toNum(C{rr,c}))
            r = rr; return;
        end
    end
    error("No numeric value found in column %d near row %d.", c, rStart);
end

function r = firstRowFirstColContainsBelow(C, rStart, pattern, maxRows)
    rEnd = min(size(C,1), rStart+maxRows-1);
    col1 = lower(string(C(rStart:rEnd,1)));
    idx = find(contains(col1, lower(pattern)), 1, 'first');
    if isempty(idx), error("Row containing '%s' not found below row %d.", pattern, rStart); end
    r = rStart + idx - 1;
end

function v = valueRightOfLabel(C, label)
    S = lower(string(C));
    [r,c] = find(contains(S, lower(label)), 1, 'first');
    if isempty(r), error("Label '%s' not found.", label); end
    v = firstNumericToRight(C, r, c);
end

function v = valueRightOfLabelLast(C, label)
    S = lower(string(C));
    idx = find(contains(S, lower(label)));
    if isempty(idx), error("Label '%s' not found.", label); end
    [r,c] = ind2sub(size(S), idx(end));
    v = firstNumericToRight(C, r, c);
end

function v = firstNumericToRight(C, r, cStart)
    for c = cStart+1:size(C,2)
        vv = toNum(C{r,c});
        if ~isnan(vv), v = vv; return; end
    end
    error("No numeric value found to the right of label at row %d.", r);
end

function x = toNum(v)
    if isnumeric(v)
        x = v; if isempty(x), x = NaN; end; return;
    end
    s = strtrim(string(v));
    if strlength(s)==0 || ismissing(s), x = NaN; return; end
    s = erase(s, '"');
    x = str2double(s);
    if isnan(x)
        x = str2double(strrep(s, ',', '.'));
    end
end
