#!/usr/bin/env python
# coding: utf-8

# In[1]:

# pip install pandas plotly kaleido

import pandas as pd
import plotly.graph_objs as go
import plotly.express as px
import plotly.io as pio
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Alignment
from typing import Any


# In[2]:


class TestAutomationAnalyzer:
    def __init__(self, csv_path: str):
        # Read and process CSV
        self.df = pd.read_csv(csv_path)
        self.consolidated_df = self._consolidate_epics()

    def _consolidate_epics(self) -> pd.DataFrame:
        # Aggregate test results by EPIC
        consolidated = self.df.groupby('Epic').agg({
            'PASSED': 'sum',
            'FAILED': 'sum',
            'BROKEN': 'sum',
            'SKIPPED': 'sum',
            'UNKNOWN': 'sum'
        }).reset_index()

        # Calculate total tests and pass rate
        consolidated['totalTests'] = consolidated[['PASSED', 'FAILED', 'BROKEN', 'SKIPPED', 'UNKNOWN']].sum(axis=1)
        consolidated['passRate'] = (consolidated['PASSED'] / consolidated['totalTests'] * 100).round(2)

        # Determine status
        consolidated['status'] = consolidated['passRate'].apply(self._determine_status)

        # Sort by status
        consolidated.sort_values(by='status', inplace=True)
        
        return consolidated

    def _determine_status(self, pass_rate: float) -> str:
        if pass_rate >= 95:
            return 'Acceptable'
        elif pass_rate >= 80:
            return 'Maintenance Advised'
        else:
            return 'Review Required'

    def generate_epic_summary_table(self) -> go.Figure:
        # Add totals row
        totals = self.consolidated_df[['PASSED', 'FAILED', 'BROKEN', 'SKIPPED', 'UNKNOWN', 'totalTests']].sum()
        totals_row = pd.DataFrame({
            'Epic': ['TOTAL'],
            'PASSED': [totals['PASSED']],
            'FAILED': [totals['FAILED']],
            'BROKEN': [totals['BROKEN']],
            'SKIPPED': [totals['SKIPPED']],
            'UNKNOWN': [totals['UNKNOWN']],
            'totalTests': [totals['totalTests']],
            'passRate': [''],
            'status': ['']
        })

        final_df = pd.concat([self.consolidated_df, totals_row], ignore_index=True)

        # Create interactive table with Plotly
        fig = go.Figure(data=[go.Table(
            header=dict(values=['Epic', 'Total Tests', 'Passed', 'Failed', 'Broken', 'Skipped', 'Pass Rate', 'Status'],
                        fill_color='paleturquoise',
                        align='center'),
            cells=dict(values=[
                final_df['Epic'],
                final_df['totalTests'],
                final_df['PASSED'],
                final_df['FAILED'],
                final_df['BROKEN'],
                final_df['SKIPPED'],
                final_df['passRate'].apply(lambda x: f'{x}%' if x != '' else ''),
                final_df['status']
            ],
                fill_color=[
                    'white',
                    'white',
                    'lightgreen',
                    'lightcoral',
                    'lightsalmon',
                    'lightblue',
                    'white',
                    final_df['status'].apply(lambda x: 'lightgreen' if x == 'Acceptable' 
                                              else 'yellow' if x == 'Maintenance Advised' 
                                              else 'lightpink' if x == 'Review Required' else 'white')
                ],
                align='center')
        )])

        fig.update_layout(
            title='Test Automation EPIC Summary',
            height=4000,
            width=1800
        )

        return fig

    def save_epic_summary_table(self, output_path: str = 'epic_summary_table_ChatGPT_3.png'):
        # Save table as static image
        fig = self.generate_epic_summary_table()
        pio.write_image(fig, output_path)
        print(f"Table image saved to {output_path}")

    def save_epic_summary_to_excel(self, output_excel_path: str = 'epic_summary_ChatGPT_3.xlsx'):
        # Add totals row
        totals = self.consolidated_df[['PASSED', 'FAILED', 'BROKEN', 'SKIPPED', 'UNKNOWN', 'totalTests']].sum()
        totals_row = pd.DataFrame({
            'Epic': ['TOTAL'],
            'PASSED': [totals['PASSED']],
            'FAILED': [totals['FAILED']],
            'BROKEN': [totals['BROKEN']],
            'SKIPPED': [totals['SKIPPED']],
            'UNKNOWN': [totals['UNKNOWN']],
            'totalTests': [totals['totalTests']],
            'passRate': [''],
            'status': ['']
        })

        final_df = pd.concat([self.consolidated_df, totals_row], ignore_index=True)

        # Write to Excel
        wb = Workbook()
        ws = wb.active
        ws.title = "EPIC Summary"

        # Write headers and data
        for r_idx, row in enumerate(dataframe_to_rows(final_df, index=False, header=True), 1):
            for c_idx, value in enumerate(row, 1):
                cell = ws.cell(row=r_idx, column=c_idx, value=value)

                # Bold headers
                if r_idx == 1:
                    cell.font = Font(bold=True)
                    cell.alignment = Alignment(horizontal='center')

        # Save workbook
        wb.save(output_excel_path)
        print(f"Excel file saved to {output_excel_path}")


# In[3]:


if __name__ == '__main__':
    # Replace with your actual CSV path
    analyzer = TestAutomationAnalyzer('IH Application weekly test automation results grouped by JIRA EPIC.csv')

    # Generate and save table image
    analyzer.save_epic_summary_table()

    # Generate and save Excel summary
    analyzer.save_epic_summary_to_excel()

