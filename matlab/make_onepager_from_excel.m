%% make_onepager_from_excel.m
% One-page Winter vs Summer thermal balance from THERMAL-ANALYSIS.xlsx
%
% Expected repo structure:
%   excel/THERMAL-ANALYSIS.xlsx
%   outputs/   (created if missing)
%
% Output:
%   outputs/thermal_balance_onepager.png
%   outputs/thermal_balance_onepager.pdf

clear; clc;

%% Paths (script is inside /matlab)
thisDir  = fileparts(mfilename('fullpath'));
repoRoot = fullfile(thisDir, '..');

xlsFile = fullfile(repoRoot, 'excel', 'THERMAL-ANALYSIS.xlsx');
outDir  = fullfile(repoRoot, 'outputs');

if ~exist(xlsFile, 'file')
    error('Excel file not found: %s', xlsFile);
end
if ~exist(outDir, 'dir')
    mkdir(outDir);
end

%% Sheet names (change only if your workbook uses different names)
sheetWinter = 'Winter';
sheetSummer = 'Summer';

%% ----------------------------
% WINTER (Heat loss)
% (cell addresses based on your current workbook layout)
% ----------------------------
wWalls   = sum(readnum(xlsFile, sheetWinter, {'K6','K7','K8','K9'}), 'omitnan');
wWindows = readnum(xlsFile, sheetWinter, 'K15');
wFloor   = readnum(xlsFile, sheetWinter, 'K21');
wCeiling = readnum(xlsFile, sheetWinter, 'K27');

wTransferTotal = readnum(xlsFile, sheetWinter, 'K30');  % heat transfer total
wAir           = readnum(xlsFile, sheetWinter, 'C38');  % air exchange total
wTotal         = readnum(xlsFile, sheetWinter, 'C41');  % final total

% Assumptions block
wACH  = readnum(xlsFile, sheetWinter, 'C34');
wVol  = readnum(xlsFile, sheetWinter, 'C35');
wTout = readnum(xlsFile, sheetWinter, 'C36');
wTin  = readnum(xlsFile, sheetWinter, 'C37');

winterLabels = {'Walls','Windows','Floor','Ceiling','Air exchange'};
winterVals_W = [wWalls, wWindows, wFloor, wCeiling, wAir];

%% ----------------------------
% SUMMER (Heat gain)
% ----------------------------
sWalls   = sum(readnum(xlsFile, sheetSummer, {'H6','H7','H8','H9'}), 'omitnan');
sFloor   = readnum(xlsFile, sheetSummer, 'J15');      % may be negative (ground cooling)

% Windows: total window gain and solar term
sWinSolar = readnum(xlsFile, sheetSummer, 'L20');     % Q_solar (W)
sWinTotal = readnum(xlsFile, sheetSummer, 'M20');     % Q_win (W) = conduction + solar
sWinCond  = sWinTotal - sWinSolar;

sCeiling = readnum(xlsFile, sheetSummer, 'H26');
sPeople  = readnum(xlsFile, sheetSummer, 'F33');
sLight   = readnum(xlsFile, sheetSummer, 'E38');      % IMPORTANT: Q_light (W), not q_light (W/m^2)

sTotal   = readnum(xlsFile, sheetSummer, 'H42');      % Q_gain_tot (W)

% Assumptions
sTground = readnum(xlsFile, sheetSummer, 'G15');
sTout    = readnum(xlsFile, sheetSummer, 'G20');
sTin     = readnum(xlsFile, sheetSummer, 'H20');
sI       = readnum(xlsFile, sheetSummer, 'J20');
sTau     = readnum(xlsFile, sheetSummer, 'K20');

summerLabels = {'Walls','Windows','Ceiling','People','Lighting','Floor (ground)'};
summerVals_W = [sWalls, sWinTotal, sCeiling, sPeople, sLight, sFloor];

%% ----------------------------
% Build one-page figure (A4 landscape)
% ----------------------------
fig = figure('Color','w');
set(fig,'Units','inches','Position',[0.3 0.3 11.69 8.27]); % A4 landscape

t = tiledlayout(3,2,'TileSpacing','compact','Padding','compact');

% Title tile (spans)
axTitle = nexttile(t,[1 2]);
axis(axTitle,'off');
text(0,0.80,'Building Thermal Balance — Winter Loss vs Summer Gains', ...
    'FontSize',16,'FontWeight','bold');
text(0,0.52,'Source: excel/THERMAL-ANALYSIS.xlsx (Winter & Summer sheets)', ...
    'FontSize',10);

% Winter bar (keep order)
axW = nexttile(t);
catsW = categorical(winterLabels, winterLabels, 'Ordinal', true);
bar(axW, catsW, winterVals_W/1000);
grid(axW,'on');
ylabel(axW,'kW');
title(axW,'Winter — Heat Loss Breakdown');
xtickangle(axW,20);

% Summer stacked bar (Windows split: conduction + solar)
axS = nexttile(t);
base  = [sWalls, sWinCond, sCeiling, sPeople, sLight, sFloor] / 1000;
solar = [0,      sWinSolar, 0,        0,       0,      0]      / 1000;
catsS = categorical(summerLabels, summerLabels, 'Ordinal', true);
bar(axS, catsS, [base(:), solar(:)], 'stacked');
grid(axS,'on');
ylabel(axS,'kW');
title(axS,'Summer — Heat Gain Breakdown (Windows split: conduction + solar)');
xtickangle(axS,20);
legend(axS, {'Base','Solar (windows)'}, 'Location','best');

% Info tile (spans)
axInfo = nexttile(t,[1 2]);
axis(axInfo,'off');

% Shares (for short interpretation line)
w_total_pos = sum(max(winterVals_W,0),'omitnan');
w_share_air = 100 * (wAir / w_total_pos);

summerPos = max([sWalls, sWinTotal, sCeiling, sPeople, sLight], 0);
s_sumPos  = sum(summerPos,'omitnan');
s_share_win   = 100 * (sWinTotal / s_sumPos);
s_share_solar = 100 * (sWinSolar / s_sumPos);

winterInfo = sprintf([ ...
    'WINTER assumptions: Tin=%.1f C | Tout=%.1f C | ACH=%.2f 1/h | V=%.1f m^3\n' ...
    'WINTER totals: Transfer=%.2f kW | Air exchange=%.2f kW | Total=%.2f kW (Air exchange = %.0f%%)\n' ...
    ], ...
    wTin, wTout, wACH, wVol, wTransferTotal/1000, wAir/1000, wTotal/1000, w_share_air);

summerInfo = sprintf([ ...
    'SUMMER assumptions: Tin=%.1f C | Tout=%.1f C | Tground=%.1f C | I=%.0f W/m^2 | tau=%.3f\n' ...
    'SUMMER total (net): %.2f kW (Windows = %.0f%% of positive gains; Solar = %.0f%%)\n' ...
    'Note: Floor term may be negative (cooling contribution).\n' ...
    ], ...
    sTin, sTout, sTground, sI, sTau, sTotal/1000, s_share_win, s_share_solar);

text(0,0.78,winterInfo,'FontSize',10);
text(0,0.42,summerInfo,'FontSize',10);

%% Export
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

%% ---- helper: read numeric cell(s) ----
function v = readnum(file, sheet, ref)
    if iscell(ref)
        v = nan(1,numel(ref));
        for i = 1:numel(ref)
            v(i) = readnum(file, sheet, ref{i});
        end
        return;
    end

    c = readcell(file, 'Sheet', sheet, 'Range', ref);
    x = c{1};

    if isempty(x)
        v = NaN; return;
    end
    if isnumeric(x)
        v = x; return;
    end

    s = strtrim(string(x));
    if strlength(s)==0 || ismissing(s)
        v = NaN; return;
    end

    % handle comma decimals if needed
    s = strrep(s, ',', '.');
    v = str2double(s);
end
