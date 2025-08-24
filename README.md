# The Hidden Costs of Subsidized Housing in Bogotá

This repository contains a comprehensive analysis of VIS/VIP (Vivienda de Interés Social/Prioritario) subsidized housing in Colombia, examining the true costs of homeownership versus renting over a 20-year period.

## Key Findings

Under realistic assumptions, **VIS ownership never beats renting within 20 years**. The analysis includes often-overlooked costs such as:

- Gray-delivery finishing costs (obra gris) financed at consumer rates
- Uncovered down payment portions financed at high interest rates
- HOA fees growing with inflation
- Maintenance costs growing above inflation (CPI+2%)
- Commuting time costs due to peripheral locations
- Transaction costs and illiquidity
- 10-year lockup period with clawback provisions

## Repository Contents

### Paper
- `vis_paper.tex` - LaTeX source for the analysis paper
- `vis_paper.pdf` - Compiled PDF of the paper
- `cash_outflows_dominant.pdf` - Figure showing cumulative cash outflows
- `net_if_sold_dominant.pdf` - Figure showing net position if property is sold

### Analysis Tools
- `vis_housing_analysis.py` - Python script that generates scenario analysis
- `vis_housing_analysis.xlsx` - Excel spreadsheet with all scenarios
- `vis_housing_spreadsheet.py` - Advanced Excel generator with formulas

### Supporting Files
- `bogota_cranes.jpg` - Image of Bogotá construction
- `vis_paper_regen.zip` - Archive of regenerated figures

## How to Use

### Compile the Paper
```bash
pdflatex vis_paper.tex
```

### Run the Analysis
```bash
python3 vis_housing_analysis.py
```

This will generate an Excel file with all scenarios comparing different CPI rates (1%, 3%, 5%) and property appreciation rates (0%, 1%, 2%).

### Create Interactive Spreadsheet
```bash
python3 vis_housing_spreadsheet.py
```

This creates a more detailed Excel model with transparent formulas and adjustable parameters.

## Key Parameters

- **Property Price**: 135M COP (135 SMMLV)
- **Down Payment**: 20% (10% subsidized, 10% financed)
- **Finishing Costs**: 30M COP financed at 13% for 5 years
- **Mortgage**: 9% APR for 20 years
- **Commuting Time**: 10 hrs/week (rent) vs 20 hrs/week (own)
- **Lockup Period**: 10 years

## Authors

**Igor Rivin**  
*with assistance from Claude and ChatGPT*

## License

This analysis is provided for educational and research purposes. Please cite appropriately if using this work.

## Acknowledgments

Thanks to Claude (Anthropic) and ChatGPT (OpenAI) for assistance in developing the analysis framework and calculations.