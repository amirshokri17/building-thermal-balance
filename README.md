# Building Thermal Balance (Excel + MATLAB)

Steady-state winter heat-loss and summer heat-gain breakdown for a residential case study.

## Scope
- **Winter**: transmission losses (U·A·ΔT) + ventilation/air exchange (ACH)
- **Summer**: equivalent temperature method for opaque elements + window conduction & solar gains + internal loads (people, lighting)

## How to use
1. Open `excel/THERMAL-ANALYSIS.xlsx` to review inputs and calculations.
2. Run `matlab/winter_heat_loss.m` for winter component losses.
3. Run `matlab/summer_heat_gain.m` for summer component gains.

## Repository structure
- `excel/` – calculation model (XLSX)
- `matlab/` – scripts (winter/summer breakdown)
- `report/` – documentation/report
- `outputs/` – exported plots/screenshots
