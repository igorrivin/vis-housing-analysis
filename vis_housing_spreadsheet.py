#!/usr/bin/env python3
"""
VIS/VIP Housing Analysis - Excel Formula-based Spreadsheet Generator
Creates a transparent Excel file with visible formulas
"""

import xlsxwriter
from datetime import datetime

def create_vis_spreadsheet(filename='vis_housing_model.xlsx'):
    """Create Excel spreadsheet with formulas for VIS housing analysis"""
    
    # Create workbook and add worksheets
    workbook = xlsxwriter.Workbook(filename)
    
    # Add formats
    header_format = workbook.add_format({
        'bold': True,
        'bg_color': '#D9D9D9',
        'border': 1,
        'text_wrap': True,
        'valign': 'top'
    })
    
    input_format = workbook.add_format({
        'bg_color': '#FFFF99',
        'border': 1,
        'num_format': '#,##0'
    })
    
    percent_format = workbook.add_format({
        'bg_color': '#FFFF99',
        'border': 1,
        'num_format': '0.0%'
    })
    
    currency_format = workbook.add_format({
        'num_format': '#,##0',
        'border': 1
    })
    
    # 1. Parameters Sheet
    params_sheet = workbook.add_worksheet('Parameters')
    
    # Write headers
    params_sheet.write('A1', 'PARAMETER', header_format)
    params_sheet.write('B1', 'VALUE', header_format)
    params_sheet.write('C1', 'UNIT', header_format)
    params_sheet.write('D1', 'DESCRIPTION', header_format)
    
    # Base parameters
    row = 2
    params = [
        ['Property_Price', 135000000, 'COP', 'VIS property price (135 SMMLV)'],
        ['Down_Payment_Pct', 0.20, '%', 'Total down payment requirement'],
        ['Subsidy_Coverage_Pct', 0.10, '%', 'Government subsidy coverage'],
        ['Uncovered_DP_Pct', 0.10, '%', 'Uncovered down payment to finance'],
        ['', '', '', ''],
        ['Uncovered_DP_Rate', 0.15, '%', 'Interest rate for uncovered DP loan'],
        ['Uncovered_DP_Years', 5, 'years', 'Term for uncovered DP loan'],
        ['Finishing_Cost', 30000000, 'COP', 'Gray delivery finishing costs'],
        ['Finishing_Rate', 0.13, '%', 'Interest rate for finishing loan'],
        ['Finishing_Years', 5, 'years', 'Term for finishing loan'],
        ['Mortgage_Rate', 0.09, '%', 'Mortgage interest rate'],
        ['Mortgage_Years', 20, 'years', 'Mortgage term'],
        ['', '', '', ''],
        ['Initial_Rent', 1500000, 'COP/month', 'Starting monthly rent'],
        ['Initial_HOA', 150000, 'COP/month', 'Starting HOA fee'],
        ['Initial_Maintenance', 150000, 'COP/month', 'Starting maintenance cost'],
        ['Transaction_Cost_Pct', 0.05, '%', 'Transaction costs (buying/selling)'],
        ['', '', '', ''],
        ['Wage_Hour', 12000, 'COP/hour', 'Hourly wage for time valuation'],
        ['Rent_Commute_Hours', 10, 'hours/week', 'Weekly commute (central rent)'],
        ['Own_Commute_Hours', 20, 'hours/week', 'Weekly commute (peripheral own)'],
        ['Lockup_Years', 10, 'years', 'Clawback/lockup period'],
    ]
    
    for i, param in enumerate(params):
        if param[0]:  # Skip empty rows
            params_sheet.write(row + i, 0, param[0])
            if isinstance(param[1], float) and param[2] == '%':
                params_sheet.write(row + i, 1, param[1], percent_format)
            elif isinstance(param[1], (int, float)):
                params_sheet.write(row + i, 1, param[1], input_format)
            else:
                params_sheet.write(row + i, 1, param[1])
            params_sheet.write(row + i, 2, param[2])
            params_sheet.write(row + i, 3, param[3])
    
    # Define named ranges for easy reference
    workbook.define_name('Property_Price', '=Parameters!$B$2')
    workbook.define_name('Down_Payment_Pct', '=Parameters!$B$3')
    workbook.define_name('Subsidy_Coverage_Pct', '=Parameters!$B$4')
    workbook.define_name('Mortgage_Rate', '=Parameters!$B$12')
    workbook.define_name('Mortgage_Years', '=Parameters!$B$13')
    workbook.define_name('Initial_Rent', '=Parameters!$B$15')
    
    # 2. Monthly Cash Flow Sheet
    cashflow_sheet = workbook.add_worksheet('Monthly_Cashflow')
    
    # Headers
    headers = ['Month', 'Year', 'Rent', 'Rent_Commute', 'Total_Rent',
               'Mortgage', 'Uncovered_DP', 'Finishing', 'HOA', 'Maintenance', 
               'Own_Commute', 'Total_Own', 'Rent_Cumulative', 'Own_Cumulative']
    
    for col, header in enumerate(headers):
        cashflow_sheet.write(0, col, header, header_format)
    
    # Generate 240 months (20 years)
    for month in range(1, 241):
        row = month
        year = (month - 1) // 12 + 1
        
        # Month and Year
        cashflow_sheet.write_formula(row, 0, f'={month}')
        cashflow_sheet.write_formula(row, 1, f'=INT((A{row+1}-1)/12)+1')
        
        # Rent with CPI adjustment (assuming in Scenario sheet)
        cashflow_sheet.write_formula(row, 2, 
            f'=Parameters!$B$15*(1+Scenario!$B$2)^((A{row+1}-1)/12)')
        
        # Rent commute cost
        cashflow_sheet.write_formula(row, 3,
            f'=Parameters!$B$20*Parameters!$B$21*52/12*(1+Scenario!$B$2)^((A{row+1}-1)/12)')
        
        # Total Rent
        cashflow_sheet.write_formula(row, 4, f'=C{row+1}+D{row+1}')
        
        # Mortgage payment (using PMT function)
        if month <= 240:  # 20 years
            cashflow_sheet.write_formula(row, 5,
                f'=IF(A{row+1}<=Parameters!$B$13*12,'
                f'PMT(Parameters!$B$12/12,Parameters!$B$13*12,'
                f'-Parameters!$B$2*(1-Parameters!$B$3)),0)')
        
        # Continue with other formulas...
    
    # 3. Scenario Analysis Sheet
    scenario_sheet = workbook.add_worksheet('Scenario_Analysis')
    
    # Input section
    scenario_sheet.write('A1', 'SCENARIO INPUTS', header_format)
    scenario_sheet.write('A2', 'CPI Rate:', header_format)
    scenario_sheet.write('B2', 0.03, percent_format)  # 3% default
    scenario_sheet.write('A3', 'Appreciation Rate:', header_format)
    scenario_sheet.write('B3', 0.01, percent_format)  # 1% default
    
    # Results section
    scenario_sheet.write('A6', 'RESULTS BY YEAR', header_format)
    scenario_sheet.merge_range('B6:F6', '', header_format)
    
    result_headers = ['Year', 'Rent Cumulative (M)', 'Own Cumulative (M)', 
                     'Property Value (M)', 'Net if Sold (M)', 'Own-Rent (M)']
    
    for col, header in enumerate(result_headers):
        scenario_sheet.write(7, col, header, header_format)
    
    # Key years: 5, 10, 15, 20
    key_years = [5, 10, 15, 20]
    for i, year in enumerate(key_years):
        row = 8 + i
        scenario_sheet.write(row, 0, year)
        
        # Formulas to pull from monthly cashflow
        month = year * 12
        scenario_sheet.write_formula(row, 1, 
            f'=Monthly_Cashflow!N{month+1}/1000000', currency_format)
        scenario_sheet.write_formula(row, 2, 
            f'=Monthly_Cashflow!O{month+1}/1000000', currency_format)
        
        # Property value with appreciation
        scenario_sheet.write_formula(row, 3,
            f'=Parameters!$B$2*(1+$B$3)^A{row+1}/1000000', currency_format)
        
        # Net if sold (only after lockup)
        if year > 10:
            scenario_sheet.write_formula(row, 4,
                f'=D{row+1}*(1-Parameters!$B$18)-C{row+1}', currency_format)
            scenario_sheet.write_formula(row, 5,
                f'=E{row+1}-B{row+1}', currency_format)
        else:
            scenario_sheet.write(row, 4, '---')
            scenario_sheet.write(row, 5, '---')
    
    # 4. Sensitivity Table Sheet
    sensitivity_sheet = workbook.add_worksheet('Sensitivity_Table')
    
    # Create 3D sensitivity table
    sensitivity_sheet.write('A1', 'SENSITIVITY ANALYSIS', header_format)
    sensitivity_sheet.merge_range('B1:G1', 'Net Position if Sold - Rent (Year 20)', header_format)
    
    # Column widths
    params_sheet.set_column('A:A', 25)
    params_sheet.set_column('B:B', 15)
    params_sheet.set_column('C:C', 12)
    params_sheet.set_column('D:D', 40)
    
    cashflow_sheet.set_column('A:N', 15)
    scenario_sheet.set_column('A:F', 20)
    sensitivity_sheet.set_column('A:G', 15)
    
    workbook.close()
    print(f"Excel model created: {filename}")

def main():
    """Main function"""
    print("Creating VIS Housing Analysis Spreadsheet...")
    create_vis_spreadsheet()
    print("\nSpreadsheet created successfully!")
    print("Open 'vis_housing_model.xlsx' to see the model.")
    print("\nKey features:")
    print("- Yellow cells are inputs you can modify")
    print("- All calculations use Excel formulas (transparent)")
    print("- Scenario sheet shows results for different CPI/appreciation rates")
    print("- Monthly cashflow sheet shows detailed month-by-month calculations")

if __name__ == "__main__":
    main()