%% make_onepager_from_csv.m  
% Creates a 1-page (A4 landscape) Winter vs Summer thermal balance from CSV exports.
%
% Repo structure expected:
%   excel/THERMAL-ANALYSIS2.csv   -> WINTER (heat loss)
%   excel/THERMAL-ANALYSIS.csv    -> SUMMER (heat gain)
%   outputs/                      -> script writes PNG + PDF here
%
% Output files:
%   outputs/thermal_balance_onepager.png
%   outputs/thermal_balance_onepager.pdf

clear; clc;

%% Paths
thisDir  = fileparts(mfilename('fullpath'));
repoRoot = fullfile(thisDir, '..');

winterCsv = fullfile(repoRoot, 'excel', 'THERMAL-ANALYSIS2.csv'); % WINTER
summerCsv = fullfile(repoRoot, 'excel', 'THERMAL-ANALYSIS.csv');  % SUMMER
outDir    = fullfile(repoRoot, 'outputs');

assert(exist(winterCsv,'file')==2, 'Missing winter CSV: %s', winterCsv);
assert(exist(summerCsv,'file')==2, 'Missing summer CSV: %s', summerCsv);
if ~exist(outDir,'dir'), mkdir(outDir); end

%% Read CSVs (comma OR semicolon)
W = readCsvFlexible(winterCsv);
S = readCsvFlexible(summerCsv);

%% =========================
% WINTER (Heat loss)
% =========================

% Walls: sum Q_loss from the 4 wall rows
rWalls = findRowColContains(W, 1, "Walls:");
[rHdrW, cQlossW] = findHeaderRowAndCol(W, rWalls, 15, "Q_loss");
wallRows = collectDataRows(W, rHdrW+1, 1, 12);
wWallVals = nan(size(wallRows));
for i = 1:numel(wallRows)
    wWallVals(i) = toNum(W{wallRows(i), cQlossW});
end
wWalls = sum(wWallVals, 'omitnan');

% Windows (single row with Q_loss)
rWin = findRowColContains(W, 1, "Windows:");
[rHdrWin, cQlossWin] = findHeaderRowAndCol(W, rWin, 15, "Q_loss");
rWinData = firstNumericRow(W, rHdrWin+1, cQlossWin, 15);
wWindows = toNum(W{rWinData, cQlossWin});

% Floor
rFloor = findRowColContains(W, 1, "Floor:");
[rHdrF, cQlossF] = findHeaderRowAndCol(W, rFloor, 15, "Q_loss");
rFData = firstNumericRow(W, rHdrF+1, cQlossF, 15);
wFloor = toNum(W{rFData, cQlossF});

% Ceiling
rCeil = findRowColContains(W, 1, "Ceiling:");
[rHdrC, cQlossC] = findHeaderRowAndCol(W, rCeil, 15, "Q_loss");
rCData = firstNumericRow(W, rHdrC+1, cQlossC, 15);
wCeiling = toNum(W{rCData, cQlossC});

% Air exchange + totals (labels are somewhere in the table)
wAir           = valueRightOfLabelAny(W, "Q_tot,loss (by air exchange)");
wTransferTotal = valueRightOfLabelAny(W, "Q_tot,loss (by heat transfer)");
wTotal         = valueRightOfLabelLastAny(W, "Q_tot,loss");

% Assumptions (winter)
wACH  = valueRightOfLabelAny(W, "n (ACH)");
wVol  = valueRightOfLabelAny(W, "Volume of heated space");
wTout = valueRightOfLabelAny(W, "External Temperature");
wTin  = valueRightOfLabelAny(W, "Internal Temperature");

winterLabels = {'Walls','Windows','Floor','Ceiling','Air exchange'};
winterVals_W = [wWalls, wWindows, wFloor, wCeiling, wAir];

%% =========================
% SUMMER (Heat gain)
% =========================

% Walls: sum Q_wall (W)
rWallsS = findRowColContains(S, 1, "Walls:");
[rHdrWS, cQwall] = findHeaderRowAndCol(S, rWallsS, 15, "Q_wall");
wallRowsS = collectDataRows(S, rHdrWS+1, 1, 12);
sWallVals = nan(size(wallRowsS));
for i = 1:numel(wallRowsS)
    sWallVals(i) = toNum(S{wallRowsS(i), cQwall});
end
sWalls = sum(sWallVals, 'omitnan');

% Floor: Q_floor (can be negative)
rFloorS = findRowColContains(S, 1, "Floor:");
[rHdrFS, cQfloor] = findHeaderRowAndCol(S, rFloorS, 15, "Q_floor");
rFloorData = firstNumericRow(S, rHdrFS+1, cQfloor, 15);
sFloor = toNum(S{rFloorData, cQfloor});

% Windows: Q_win and Q_solar from the windows row
rWinS = findRowColContains(S, 1, "Windows:");
[rHdrWinS, cQwin] = findHeaderRowAndColPrefer(S, rWinS, 20, "Q_win", "(w)", "w/m");
[~,         cQsol] = findHeaderRowAndColPrefer(S, rWinS, 20, "Q_solar", "(w)", "w/m");

rWinDataS = firstRowColContainsBelow(S, rHdrWinS+1, 1, "Wall", 20);
sWinTotal = toNum(S{rWinDataS, cQwin});
sWinSolar = toNum(S{rWinDataS, cQsol});
sWinCond  = sWinTotal - sWinSolar;

% Ceiling: Q_ceiling
rCeilS = findRowColContains(S, 1, "Ceiling:");
[rHdrCeilS, cQceil] = findHeaderRowAndCol(S, rCeilS, 15, "Q_ceiling");
rCeilData = firstNumericRow(S, rHdrCeilS+1, cQceil, 15);
sCeiling = toNum(S{rCeilData, cQceil});

% People: FIX (read value UNDER column Q_people, not to the right)
rPeople = findRowColContains(S, 1, "People:");
[rHdrPeople, cQpeople] = findHeaderRowAndColPrefer(S, rPeople, 20, "Q_people", "(w)", "w/m");
rPeopleData = firstNumericRow(S, rHdrPeople+1, cQpeople, 20);
sPeople = toNum(S{rPeopleData, cQpeople});

% Lighting: FIX (must pick Q_light (W), not q_light (W/m^2))
rLight = findRowAnyColContains(S, "Heat Exchange by Lighting");
rHdrLight = findRowInRangeAnyColContains(S, rLight, 20, "light");
cQlight = findColPreferPowerNotFlux(S, rHdrLight, "light");  % <= this is the key fix
rLightData = firstNumericRow(S, rHdrLight+1, cQlight, 20);
sLight = toNum(S{rLightData, cQlight});

% Total summer gain
sTotal = valueRightOfLabelAny(S, "Q_gain_tot");

% Summer assumptions (from the windows header row)
rWinHdr = findRowInRangeAnyColContains(S, rWinS, 25, "External Temperature");
cToutS  = findColInRowContains(S, rWinHdr, "External Temperature");
cTinS   = findColInRowContains(S, rWinHdr, "Internal Temperature");
cIS     = findColInRowContains(S, rWinHdr, "I-value");
% tau can appear as "tau" or "t-value" depending on export
cTauS   = findColInRowContainsAny(S, rWinHdr, ["tau", "transmittance", "t-value"]);

sTout = toNum(S{rWinDataS, cToutS});
sTin  = toNum(S{rWinDataS, cTinS});
sI    = toNum(S{rWinDataS, cIS});
sTau  = toNum(S{rWinDataS, cTauS});

% Ground temperature from floor table
rFloorHdr = findRowInRangeAnyColContains(S, rFloorS, 20, "External (ground) Temperature");
cTg = findColInRowContains(S, rFloorHdr, "External (ground) Temperature");
rTgData = firstNumericRow(S, rFloorHdr+1, cTg, 20);
sTground = toNum(S{rTgData, cTg});

summerLabels = {'Walls','Windows','Ceiling','People','Lighting','Floor (ground)'};
summerVals_W = [sWalls, sWinTotal, sCeiling, sPeople, sLight, sFloor];

%% =========================
% Build one-pager (A4 landscape)
% =========================

fig = figure('Color','w');
set(fig,'Units','inches','Position',[0.3 0.3 11.69 8.27]); % A4 landscape

tlo = tiledlayout(3,2,'TileSpacing','compact','Padding','compact');

% Title
axTitle = nexttile(tlo,[1 2]);
axis(axTitle,'off');
text(0,0.80,'Building Thermal Balance — Winter Loss vs Summer Gains','FontSize',16,'FontWeight','bold');
text(0,0.52,'Source: CSV exports (winter + summer)','FontSize',10);

% Winter bar (keep order)
axW = nexttile(tlo);
catsW = categorical(winterLabels, winterLabels, 'Ordinal', true);
bar(axW, catsW, winterVals_W/1000);
grid(axW,'on'); ylabel(axW,'kW');
title(axW,'Winter — Heat Loss Breakdown');
xtickangle(axW,20);

% Summer stacked bar (keep order)
axS = nexttile(tlo);
base  = [sWalls, sWinCond, sCeiling, sPeople, sLight, sFloor] / 1000;
solar = [0,      sWinSolar, 0,        0,       0,      0]      / 1000;
catsS = categorical(summerLabels, summerLabels, 'Ordinal', true);
bar(axS, catsS, [base(:), solar(:)], 'stacked');
grid(axS,'on'); ylabel(axS,'kW');
title(axS,'Summer — Heat Gain Breakdown (Windows split: conduction + solar)');
xtickangle(axS,20);
legend(axS, {'Base','Solar (windows)'}, 'Location','best');

% Info box (ASCII-safe sprintf: no °, τ, m³)
axInfo = nexttile(tlo,[1 2]);
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

% Export
pngOut = fullfile(outDir,'thermal_balance_onepager.png');
pdfOut = fullfile(outDir,'thermal_balance_onepager.pdf');

try
    exportgraphics(fig, pngOut, 'Resolution', 300);
    exportgraphics(fig, pdfOut);
catch
    print(fig, pngOut, '-dpng', '-r300');
    print(fig, pdfOut, '-dpdf', '-bestfit');
end

fprintf('Saved:\n  %s\n  %s\n', pngOut, pdfOut);

%% =========================
% Helper functions
% =========================

function C = readCsvFlexible(fp)
    % Try comma, if it results in a single column, retry with semicolon.
    C = readcell(fp, 'Delimiter', ',', 'TextType', 'string');
    if size(C,2) == 1
        C = readcell(fp, 'Delimiter', ';', 'TextType', 'string');
    end
end

function r = findRowColContains(C, col, pattern)
    v = lower(string(C(:,col)));
    r = find(contains(v, lower(pattern)), 1, 'first');
    if isempty(r)
        error("Could not find '%s' in column %d.", pattern, col);
    end
end

function r = findRowAnyColContains(C, pattern)
    M = lower(string(C));
    r = find(any(contains(M, lower(pattern)),2), 1, 'first');
    if isempty(r), error("Could not find '%s' in any column.", pattern); end
end

function r = findRowInRangeAnyColContains(C, rStart, maxRows, pattern)
    rEnd = min(size(C,1), rStart+maxRows);
    blk = lower(string(C(rStart:rEnd,:)));
    idx = find(any(contains(blk, lower(pattern)),2), 1, 'first');
    if isempty(idx)
        error("Could not find '%s' near row %d.", pattern, rStart);
    end
    r = rStart + idx - 1;
end

function c = findColInRowContains(C, r, pattern)
    row = lower(string(C(r,:)));
    c = find(contains(row, lower(pattern)), 1, 'first');
    if isempty(c), error("Could not find column containing '%s' in row %d.", pattern, r); end
end

function c = findColInRowContainsAny(C, r, patterns)
    row = lower(string(C(r,:)));
    c = [];
    for k = 1:numel(patterns)
        cc = find(contains(row, lower(patterns(k))), 1, 'first');
        if ~isempty(cc), c = cc; return; end
    end
    error("Could not find any of the requested columns in row %d.", r);
end

function [rHdr, cKey] = findHeaderRowAndCol(C, rStart, maxRows, key)
    rHdr = findRowInRangeAnyColContains(C, rStart, maxRows, key);
    cKey = findColInRowContains(C, rHdr, key);
end

function [rHdr, cKey] = findHeaderRowAndColPrefer(C, rStart, maxRows, key, mustInclude, mustNotInclude)
    rHdr = findRowInRangeAnyColContains(C, rStart, maxRows, key);
    row  = string(C(rHdr,:));
    % candidates containing key
    cand = find(contains(lower(row), lower(key)));
    if isempty(cand), error("No column found for '%s'.", key); end

    % prefer those that include mustInclude and do not include mustNotInclude
    best = cand(contains(lower(row(cand)), lower(mustInclude)) & ~contains(lower(row(cand)), lower(mustNotInclude)));
    if ~isempty(best)
        cKey = best(1);
        return;
    end

    % fallback: first candidate
    cKey = cand(1);
end

function c = findColPreferPowerNotFlux(C, rHdr, lightKey)
    % In lighting header row, there is usually:
    %   q_light (w/m^2)  AND  Q_light (w)
    % We want the POWER column: Q_light (w) not the flux q_light (w/m^2).
    row = string(C(rHdr,:));
    low = lower(row);

    % Candidate columns containing 'light'
    cand = find(contains(low, lower(lightKey)));
    if isempty(cand)
        error("Lighting header row found but no 'light' columns detected.");
    end

    % Prefer header that contains '(w)' AND NOT 'w/m'
    pref = cand(contains(low(cand), '(w)') & ~contains(low(cand), 'w/m'));
    if ~isempty(pref)
        c = pref(1);
        return;
    end

    % Next best: contains 'q_light' vs 'Q_light' ambiguity (case may be lost in CSV)
    % Prefer any column that contains 'q_light' AND NOT 'w/m'
    pref2 = cand(contains(low(cand), 'q_light') & ~contains(low(cand), 'w/m'));
    if ~isempty(pref2)
        c = pref2(1);
        return;
    end

    % Last resort: pick the last lighting-related column (often Q_light is last)
    c = cand(end);
end

function rows = collectDataRows(C, rStart, nameCol, maxRows)
    rows = [];
    rEnd = min(size(C,1), rStart+maxRows-1);
    for r = rStart:rEnd
        s = strtrim(string(C{r,nameCol}));
        if strlength(s)==0 || ismissing(s), break; end
        if contains(s, ":"), break; end
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

function r = firstRowColContainsBelow(C, rStart, col, pattern, maxRows)
    rEnd = min(size(C,1), rStart+maxRows-1);
    v = lower(string(C(rStart:rEnd,col)));
    idx = find(contains(v, lower(pattern)), 1, 'first');
    if isempty(idx)
        error("Could not find '%s' below row %d in column %d.", pattern, rStart, col);
    end
    r = rStart + idx - 1;
end

function v = valueRightOfLabelAny(C, label)
    low = lower(string(C));
    idx = find(contains(low, lower(label)), 1, 'first');
    if isempty(idx), error("Label '%s' not found.", label); end
    [r,c] = ind2sub(size(low), idx);
    v = firstNumericToRight(C, r, c);
end

function v = valueRightOfLabelLastAny(C, label)
    low = lower(string(C));
    idx = find(contains(low, lower(label)));
    if isempty(idx), error("Label '%s' not found.", label); end
    [r,c] = ind2sub(size(low), idx(end));
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
        x = v;
        if isempty(x), x = NaN; end
        return;
    end
    s = strtrim(string(v));
    if strlength(s)==0 || ismissing(s)
        x = NaN; return;
    end
    s = erase(s, '"');
    x = str2double(s);
    if isnan(x)
        x = str2double(strrep(s, ',', '.'));
    end
end
