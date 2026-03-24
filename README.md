## Overview

This project quantifies the carbon emission reduction contribution of offshore wind power in China.

It uses a grid-average displacement methodology to analyse historical trends (2006–2024) and project future mitigation potential (2025–2034).

The study also examines regional disparities across coastal provinces.

## Repository Structure

data/        - datasets used in the analysis  
Manuscript/  - research paper 
scripts/     - Python code for analysis  

## Data Description

This project uses multiple datasets including:

- MEIC carbon emission data  
- Offshore wind capacity data  
- National electricity production and GDP data  

Some data were manually collected and do not have associated scripts.  
Certain datasets were adjusted (e.g. coordinate alignment and formatting corrections).

These modifications are documented for transparency.

## Methodology

The project applies the grid-average displacement method to estimate carbon reduction.

Carbon reduction is calculated as:

Carbon Reduction = Electricity Generation × Grid Emission Factor

Electricity generation is estimated using installed capacity and capacity factors.

## How to Run

1. Install required packages:
pip install pandas numpy matplotlib

2. Run the script:
python scripts/code.py

## Key Findings

- Grid emission factor decreased significantly from 2006 to 2024  
- Offshore wind contributed substantial carbon reductions  
- A major increase occurred in 2021 due to policy changes  
- Future projections show strong growth in mitigation potential  
