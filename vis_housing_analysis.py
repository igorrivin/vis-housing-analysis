#!/usr/bin/env python3
"""
VIS/VIP Housing Analysis - Rent vs Buy Model
Based on the paper "The Hidden Costs of Subsidized Housing in Bogotá"
"""

import pandas as pd
import numpy as np
from datetime import datetime
import xlsxwriter

class VISHousingAnalysis:
    def __init__(self):
        # Base parameters
        self.property_price = 135_000_000  # 135M COP (135 SMMLV)
        self.down_payment_pct = 0.20
        self.subsidy_coverage_pct = 0.10
        self.uncovered_down_payment_pct = 0.10
        
        # Financing parameters
        self.uncovered_dp_rate = 0.15  # 15% APR
        self.uncovered_dp_years = 5
        self.finishing_cost = 30_000_000  # 30M COP
        self.finishing_rate = 0.13  # 13% APR
        self.finishing_years = 5
        self.mortgage_rate = 0.09  # 9% APR
        self.mortgage_years = 20
        
        # Other parameters
        self.initial_rent = 1_500_000  # 1.5M COP/month
        self.initial_hoa = 150_000  # 150K COP/month
        self.initial_maintenance = 150_000  # 150K COP/month
        self.transaction_cost_pct = 0.05  # 5% of property value
        
        # Time costs
        self.wage_per_hour = 12_000  # 12K COP/hour
        self.rent_commute_hours_week = 10
        self.own_commute_hours_week = 20
        
        # Lockup period
        self.lockup_years = 10
        
    def calculate_monthly_payment(self, principal, annual_rate, years):
        """Calculate monthly payment for a loan"""
        monthly_rate = annual_rate / 12
        n_payments = years * 12
        if monthly_rate == 0:
            return principal / n_payments
        return principal * (monthly_rate * (1 + monthly_rate)**n_payments) / \
               ((1 + monthly_rate)**n_payments - 1)
    
    def calculate_scenario(self, cpi_rate, appreciation_rate, years=20):
        """Calculate one scenario for given CPI and appreciation rates"""
        results = []
        
        # Initial calculations
        down_payment = self.property_price * self.down_payment_pct
        subsidy = self.property_price * self.subsidy_coverage_pct
        uncovered_dp = self.property_price * self.uncovered_down_payment_pct
        mortgage_principal = self.property_price * (1 - self.down_payment_pct)
        
        # Monthly payments
        uncovered_dp_payment = self.calculate_monthly_payment(
            uncovered_dp, self.uncovered_dp_rate, self.uncovered_dp_years)
        finishing_payment = self.calculate_monthly_payment(
            self.finishing_cost, self.finishing_rate, self.finishing_years)
        mortgage_payment = self.calculate_monthly_payment(
            mortgage_principal, self.mortgage_rate, self.mortgage_years)
        
        # Cumulative tracking
        rent_cumulative = 0
        own_cumulative = 0
        
        # Initial ownership costs (at closing)
        own_cumulative += subsidy  # This is what the buyer effectively pays
        own_cumulative += self.property_price * self.transaction_cost_pct
        
        for year in range(1, years + 1):
            # Annual values
            rent_annual = 0
            own_annual = 0
            
            # Calculate monthly values for this year
            for month in range(12):
                total_months = (year - 1) * 12 + month
                
                # Rent with CPI adjustment
                current_rent = self.initial_rent * (1 + cpi_rate) ** (total_months / 12)
                rent_annual += current_rent
                
                # Ownership costs
                # 1. Mortgage (fixed)
                if total_months < self.mortgage_years * 12:
                    own_annual += mortgage_payment
                
                # 2. Uncovered down payment loan
                if total_months < self.uncovered_dp_years * 12:
                    own_annual += uncovered_dp_payment
                
                # 3. Finishing loan
                if total_months < self.finishing_years * 12:
                    own_annual += finishing_payment
                
                # 4. HOA with CPI adjustment
                current_hoa = self.initial_hoa * (1 + cpi_rate) ** (total_months / 12)
                own_annual += current_hoa
                
                # 5. Maintenance with CPI+2% adjustment
                current_maintenance = self.initial_maintenance * \
                                    (1 + cpi_rate + 0.02) ** (total_months / 12)
                own_annual += current_maintenance
            
            # Add commuting time costs (CPI-indexed wage)
            current_wage = self.wage_per_hour * (1 + cpi_rate) ** year
            rent_time_cost = current_wage * self.rent_commute_hours_week * 52
            own_time_cost = current_wage * self.own_commute_hours_week * 52
            
            rent_annual += rent_time_cost
            own_annual += own_time_cost
            
            # Update cumulatives
            rent_cumulative += rent_annual
            own_cumulative += own_annual
            
            # Calculate net if sold (only after lockup)
            if year > self.lockup_years:
                property_value = self.property_price * (1 + appreciation_rate) ** year
                sale_costs = property_value * self.transaction_cost_pct
                net_if_sold = property_value - sale_costs - own_cumulative
            else:
                net_if_sold = None
            
            # Store results for key years (10, 15, 20)
            if year in [10, 15, 20]:
                results.append({
                    'CPI': f"{cpi_rate*100:.0f}%",
                    'Appr': f"{appreciation_rate*100:.0f}%",
                    'Year': year,
                    'Rent_cum_M': round(rent_cumulative / 1_000_000, 1),
                    'Own_cum_M': round(own_cumulative / 1_000_000, 1),
                    'Net_if_sold_M': '---' if net_if_sold is None else round(net_if_sold / 1_000_000, 1),
                    'Own_minus_Rent_M': '---' if net_if_sold is None else round((net_if_sold - rent_cumulative) / 1_000_000, 1)
                })
        
        return results
    
    def generate_all_scenarios(self):
        """Generate all scenarios from the paper"""
        scenarios = []
        
        # CPI rates: 1%, 3%, 5%
        # Appreciation rates: 0%, 1%, 2%
        for cpi in [0.01, 0.03, 0.05]:
            for appr in [0.00, 0.01, 0.02]:
                scenario_results = self.calculate_scenario(cpi, appr)
                scenarios.extend(scenario_results)
        
        return pd.DataFrame(scenarios)
    
    def create_excel_report(self, filename='vis_housing_analysis.xlsx'):
        """Create Excel report with all scenarios"""
        # Generate scenarios
        df = self.generate_all_scenarios()
        
        # Create Excel writer
        with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
            # Write main results
            df.to_excel(writer, sheet_name='Sensitivity Analysis', index=False)
            
            # Get workbook and worksheet
            workbook = writer.book
            worksheet = writer.sheets['Sensitivity Analysis']
            
            # Add formats
            header_format = workbook.add_format({
                'bold': True,
                'text_wrap': True,
                'valign': 'top',
                'fg_color': '#D9D9D9',
                'border': 1
            })
            
            # Format columns
            worksheet.set_column('A:C', 10)  # CPI, Appr, Year
            worksheet.set_column('D:G', 15)  # Monetary values
            
            # Apply header format
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)
            
            # Add parameters sheet
            params_df = pd.DataFrame([
                ['Parameter', 'Value', 'Unit'],
                ['Property Price', self.property_price, 'COP'],
                ['Down Payment %', self.down_payment_pct * 100, '%'],
                ['Subsidy Coverage %', self.subsidy_coverage_pct * 100, '%'],
                ['Uncovered DP Rate', self.uncovered_dp_rate * 100, '% APR'],
                ['Uncovered DP Years', self.uncovered_dp_years, 'years'],
                ['Finishing Cost', self.finishing_cost, 'COP'],
                ['Finishing Rate', self.finishing_rate * 100, '% APR'],
                ['Finishing Years', self.finishing_years, 'years'],
                ['Mortgage Rate', self.mortgage_rate * 100, '% APR'],
                ['Mortgage Years', self.mortgage_years, 'years'],
                ['Initial Rent', self.initial_rent, 'COP/month'],
                ['Initial HOA', self.initial_hoa, 'COP/month'],
                ['Initial Maintenance', self.initial_maintenance, 'COP/month'],
                ['Transaction Cost %', self.transaction_cost_pct * 100, '%'],
                ['Wage per Hour', self.wage_per_hour, 'COP/hour'],
                ['Rent Commute Hours/Week', self.rent_commute_hours_week, 'hours'],
                ['Own Commute Hours/Week', self.own_commute_hours_week, 'hours'],
                ['Lockup Period', self.lockup_years, 'years']
            ])
            
            params_df.to_excel(writer, sheet_name='Parameters', index=False, header=False)
            
            # Format parameters sheet
            params_worksheet = writer.sheets['Parameters']
            params_worksheet.set_column('A:A', 25)
            params_worksheet.set_column('B:B', 15)
            params_worksheet.set_column('C:C', 15)
            
            # Apply header format to first row
            for col in range(3):
                params_worksheet.write(0, col, params_df.iloc[0, col], header_format)
        
        print(f"Excel report created: {filename}")
        return df

def main():
    """Main function to run the analysis"""
    print("VIS/VIP Housing Analysis - Rent vs Buy Model")
    print("=" * 50)
    
    # Create analysis instance
    analysis = VISHousingAnalysis()
    
    # Generate Excel report
    df = analysis.create_excel_report()
    
    # Print summary
    print("\nScenario Summary:")
    print(df.to_string(index=False))
    
    print("\n" + "=" * 50)
    print("Analysis complete. Check 'vis_housing_analysis.xlsx' for detailed results.")

if __name__ == "__main__":
    main()