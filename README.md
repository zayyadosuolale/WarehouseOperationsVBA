# Warehouse Operations Automation Dashboard

## Overview

An Excel VBA application that automates the intake, cleaning, analysis, and visualization of simulated warehouse order-fulfillment data.

The synthetic dataset was designed using operational concepts and observations gained during a food bank operations internship. No proprietary organizational data is used.

## Features
- Automated CSV import
- Data cleaning and validation
- Row-level operational metric calculations
- Unique-order KPI aggregation
- Process bottleneck identification
- Labor cost and throughput analysis
- PivotTable and PivotChart automation
- Interactive management dashboard
- One-click analysis workflow
- Automated PDF reporting

## Dashboard
Example:
<img width="1297" height="653" alt="image" src="https://github.com/user-attachments/assets/0bd760c1-8875-4b7d-927b-275c22398600" />


## Metrics
- Average order processing time
- Labor cost per order
- Labor cost per unit
- Process-step throughput
- Target achievement rate
- Bottleneck process step
- Process-path performance

## Workflow
CSV Data
→ Import
→ Clean
→ Calculate Metrics
→ Build KPIs
→ Refresh PivotTables
→ Update Dashboard
→ Export PDF

## Technologies
- Microsoft Excel
- VBA
- Excel Tables
- PivotTables
- PivotCharts
- Dictionaries
- Automated PDF reporting

## Dataset
The primary synthetic dataset contains:
- 150 simulated orders
- 900 process-step records
- 3 fulfillment paths
- 6 processing steps

## Validation

A separate controlled six-order dataset was used to manually
verify KPI calculations and dashboard outputs.

## Project Structure
warehouse-operations-vba/
│
├── README.md
├── Warehouse_Operations_Dashboard.xlsm
├── Warehouse_Operations_Report.pdf
│
├── data/
│   ├── warehouse_orders.csv
│   └── warehouse_dashboard_test.csv
│
└── src/
    ├── modMain.bas
    ├── modDataImport.bas
    ├── modDataCleaning.bas
    ├── modCalculations.bas
    ├── modKPI.bas
    └── modDashboard.bas

## VBA Architecture
The application is separated into modules based on responsibility:
- **modMain** — Controls the complete analysis workflow and user-facing actions.
- **modDataImport** — Imports user-selected CSV files into the workbook.
- **modDataCleaning** — Cleans, validates, and prepares imported records.
- **modCalculations** — Calculates row-level operational metrics.
- **modKPI** — Aggregates order-level KPIs and identifies process bottlenecks.
- **modDashboard** — Refreshes PivotTables, updates dashboard outputs, and exports PDF reports.

## How to Use
1. Download and open `Warehouse_Operations_Dashboard.xlsm`.
2. Enable macros when prompted by Excel.
3. Navigate to the **Dashboard** sheet.
4. Click **Run Full Analysis**.
5. Select `warehouse_orders.csv` or another CSV using the required structure.
6. Review the automatically updated KPIs, charts, and process analysis.
7. Click **Export Report** to save the dashboard as a PDF.

## Key VBA Concepts Demonstrated
- Modular VBA programming
- File selection and CSV importing
- Dynamic last-row and last-column detection
- Loops and conditional logic
- Scripting Dictionaries
- Excel Table manipulation
- Automated data validation
- Unique-order aggregation
- PivotTable and PivotChart refreshing
- Dynamic dashboard updates
- Error handling
- Automated PDF export

## Disclaimer

This project uses simulated warehouse fulfillment data created for portfolio and educational purposes. The dataset was modeled using general operational knowledge gained through a food bank operations internship. No proprietary, confidential, or organizational data is included.
