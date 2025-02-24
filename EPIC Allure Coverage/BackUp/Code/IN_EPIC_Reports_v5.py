import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Alignment
from typing import Any

class TestAutomationAnalyzer:
    def __init__(self, csv_path: str):
        print("Reading CSV file...")
        self.df = pd.read_csv(csv_path)
        self.consolidated_df = self._consolidate_epics()
        print("CSV file read and data consolidated.")

    def _consolidate_epics(self) -> pd.DataFrame:
        consolidated = self.df.groupby('Epic').agg({
            'PASSED': 'sum',
            'FAILED': 'sum',
            'BROKEN': 'sum',
            'SKIPPED': 'sum',
            'UNKNOWN': 'sum'
        }).reset_index()

        consolidated['totalTests'] = consolidated[['PASSED', 'FAILED', 'BROKEN', 'SKIPPED', 'UNKNOWN']].sum(axis=1)
        consolidated['passRate'] = (consolidated['PASSED'] / consolidated['totalTests'] * 100).round(2)
        consolidated['status'] = consolidated['passRate'].apply(self._determine_status)
        consolidated.sort_values(by='status', inplace=True)

        return consolidated

    def _determine_status(self, pass_rate: float) -> str:
        if pass_rate >= 95:
            return 'Acceptable'
        elif pass_rate >= 80:
            return 'Maintenance Advised'
        else:
            return 'Review Required'

    def generate_epic_summary_table_plot(self):
        
        print("Generating table plot...")
        
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
        
        # Debugging output
        print("Final DataFrame Columns:", final_df.columns.tolist())
        print("Final DataFrame Shape:", final_df.shape)
        
        fig, ax = plt.subplots(figsize=(12, 8))
        ax.axis('tight')
        ax.axis('off')
        
        column_labels = ['Epic', 'Passed', 'Failed', 'Broken', 'Skipped', 'Unknown', 'Total Tests', 'Pass Rate %', 'Status']
        if len(final_df.columns) != len(column_labels):
            raise ValueError("Mismatch between DataFrame columns and column labels")

        table_data = final_df.values.tolist()

        table = ax.table(cellText=table_data, colLabels=column_labels, cellLoc='center', loc='center', edges='closed')
        table.auto_set_font_size(False)
        table.set_fontsize(10)
        table.scale(1.2, 1.2)

        table_colors = {
            'Acceptable': 'lightgreen',
            'Maintenance Advised': 'yellow',
            'Review Required': 'lightpink'
        }

        for i, row in enumerate(table_data):
            if i < len(self.consolidated_df):
                status = row[-1]
                for j in range(len(row)):
                    if status in table_colors:
                        table[(i+1, j)].set_facecolor(table_colors.get(status, 'white'))
        
        for key, cell in table.get_celld().items():
            if key[0] == 0:
                cell.set_fontsize(10)
                cell.set_text_props(fontweight='bold')
                cell.set_facecolor('paleturquoise')
            if key[1] != 0:
                cell.set_text_props(ha='center')

        print("Table plot generated.")

    
    def save_epic_summary_table_plot(self, output_path: str = 'epic_summary_table.png'):
        self.generate_epic_summary_table_plot()
        plt.savefig(output_path, bbox_inches='tight')
        plt.close()
        print(f"Table image saved to {output_path}")

    def save_epic_summary_to_excel(self, output_excel_path: str = 'epic_summary.xlsx'):
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

        wb = Workbook()
        ws = wb.active
        ws.title = "EPIC Summary"

        for r_idx, row in enumerate(dataframe_to_rows(final_df, index=False, header=True), 1):
            for c_idx, value in enumerate(row, 1):
                cell = ws.cell(row=r_idx, column=c_idx, value=value)
                if r_idx == 1:
                    cell.font = Font(bold=True)
                    cell.alignment = Alignment(horizontal='center')

        wb.save(output_excel_path)
        print(f"Excel file saved to {output_excel_path}")

if __name__ == '__main__':
    print("Initializing analyzer with the CSV file...")
    analyzer = TestAutomationAnalyzer('IH Application weekly test automation results grouped by JIRA EPIC.csv')
    print("Analyzer initialized.")
    
    print("Generating and saving table image...")
    analyzer.save_epic_summary_table_plot()
    print("Table image saved.")
    
    print("Generating and saving Excel summary...")
    analyzer.save_epic_summary_to_excel()
    print("Excel summary saved.")