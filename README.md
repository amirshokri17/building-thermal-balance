# Building Thermal Balance (Excel + MATLAB)

Steady-state winter heat-loss and summer heat-gain breakdown for a residential case study.

## What’s included
- **Winter**: envelope transmission + ventilation/air exchange
- **Summer**: equivalent temperature method for opaque elements + window conduction/solar gains + internal loads (people, lighting)
- Component-wise totals for quick winter vs summer comparison

## Repository structure
- `excel/` – Excel calculation model (XLSX)
- `report/` – analysis/report (PDF)
- `matlab/` – scripts for winter and summer component breakdown

## How to run
1. Open `matlab/winter_heat_loss.m` and run
2. Open `matlab/summer_heat_gain.m` and run
3. Use the Excel file in `excel/` for the full calculation sheet and checks

## Tools
Microsoft Excel, MATLAB
