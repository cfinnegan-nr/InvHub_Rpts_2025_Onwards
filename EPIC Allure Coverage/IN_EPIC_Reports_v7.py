import pandas as pd
import matplotlib.pyplot as plt
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
        # First, handle rows with Epic values
        epic_rows = self.df[self.df['Epic'].notna()]
        consolidated_epics = epic_rows.groupby('Epic').agg({
            'PASSED': 'sum',
            'FAILED': 'sum',
            'BROKEN': 'sum',
            'SKIPPED': 'sum'
        }).reset_index()

        # Handle rows with no Epic/Feature but has Story
        story_only_rows = self.df[
            (self.df['Epic'].isna()) & 
            (self.df['Feature'].isna()) & 
            (self.df['Story'].notna())
        ]
        
        # Handle completely untagged rows (no Epic, Feature, or Story)
        untagged_rows = self.df[
            (self.df['Epic'].isna()) & 
            (self.df['Feature'].isna()) & 
            (self.df['Story'].isna())
        ]

        consolidated_list = [consolidated_epics]

        if not story_only_rows.empty:
            consolidated_stories = story_only_rows.groupby('Story').agg({
                'PASSED': 'sum',
                'FAILED': 'sum',
                'BROKEN': 'sum',
                'SKIPPED': 'sum'
            }).reset_index()
            
            # Rename Story column to Epic and add suffix
            consolidated_stories = consolidated_stories.rename(columns={'Story': 'Epic'})
            consolidated_stories['Epic'] = consolidated_stories['Epic'] + ' - No EPIC Tagged'
            consolidated_list.append(consolidated_stories)

        if not untagged_rows.empty:
            untagged_summary = pd.DataFrame({
                'Epic': ['Test Cases Not Tagged'],
                'PASSED': [untagged_rows['PASSED'].sum()],
                'FAILED': [untagged_rows['FAILED'].sum()],
                'BROKEN': [untagged_rows['BROKEN'].sum()],
                'SKIPPED': [untagged_rows['SKIPPED'].sum()]
            })
            consolidated_list.append(untagged_summary)

        # Combine all dataframes
        consolidated = pd.concat(consolidated_list, ignore_index=True)

        # Calculate metrics
        consolidated['totalTests'] = consolidated[['PASSED', 'FAILED', 'BROKEN', 'SKIPPED']].sum(axis=1)
        consolidated['passRate'] = (consolidated['PASSED'] / consolidated['totalTests'] * 100).round(2)
        consolidated['status'] = consolidated['passRate'].apply(self._determine_status)
        
        # Modified sorting to order by status and then by passRate in descending order
        consolidated.sort_values(
            by=['status', 'passRate'], 
            ascending=[True, False], 
            inplace=True
        )

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
        # totals = self.consolidated_df[['PASSED', 'FAILED', 'BROKEN', 'SKIPPED', 'UNKNOWN', 'totalTests']].sum()
        totals = self.consolidated_df[['PASSED', 'FAILED', 'BROKEN', 'SKIPPED', 'totalTests']].sum()
        totals_row = pd.DataFrame({
            'Epic': ['TOTAL'],
            'PASSED': [totals['PASSED']],
            'FAILED': [totals['FAILED']],
            'BROKEN': [totals['BROKEN']],
            'SKIPPED': [totals['SKIPPED']],
            #'UNKNOWN': [totals['UNKNOWN']],
            'totalTests': [totals['totalTests']],
            'passRate': [''],
            'status': ['']
        })

        final_df = pd.concat([self.consolidated_df, totals_row], ignore_index=True)

        # Set up the plot with wider figure size for better Epic column display
        fig, ax = plt.subplots(figsize=(24, 10))  # Increased overall width
        ax.axis('tight')
        ax.axis('off')

        column_labels = ['Epic', 'Passed', 'Failed', 'Broken', 'Skipped', 'Total Tests', 'Pass Rate %', 'Status']
        if len(final_df.columns) != len(column_labels):
            raise ValueError("Mismatch between DataFrame columns and column labels")

        table_data = final_df.values.tolist()

        # Create table with processed data
        table = ax.table(cellText=table_data, colLabels=column_labels, cellLoc='center', loc='center', edges='closed')
        table.auto_set_font_size(False)
        table.set_fontsize(9)
        table.scale(1.2, 1.2)

        # Calculate widths for Epic and Status columns based on content
        max_epic_length = max(len(str(x)) for x in final_df['Epic'])
        max_status_length = max(len(str(x)) for x in final_df['status'])
        
        # Dynamic width calculation with bounds
        epic_width = min(0.5, max(0.3, max_epic_length * 0.004))  # Epic column
        status_width = min(0.25, max(0.05, max_status_length * 0.005))  # Status column

        # Distribute remaining width among other columns
        remaining_width = 1.0 - (epic_width + status_width)
        middle_cols_width = remaining_width / (len(column_labels) - 2)  # -2 for Epic and Status

        # Set column widths
        for col in range(len(column_labels)):
            for row in range(len(table_data) + 1):
                cell = table[(row, col)]
                if col == 0:  # Epic column
                    cell.set_width(epic_width)
                elif col == len(column_labels) - 1:  # Status column
                    cell.set_width(status_width)
                else:  # Other columns
                    cell.set_width(middle_cols_width)

        # Color cells based on status
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

        # Header formatting
        for key, cell in table.get_celld().items():
            if key[0] == 0:
                cell.set_fontsize(10)
                cell.set_text_props(fontweight='bold')
                cell.set_facecolor('paleturquoise')
            if key[1] not in [0, len(column_labels)-1]:  # Align all but first and last column to center
                cell.set_text_props(ha='center')

        print("Table plot generated.")

    def save_epic_summary_table_plot(self, output_path: str = None):
        if output_path is None:
            from datetime import datetime
            date_suffix = datetime.now().strftime('%d%m%y')
            output_path = f'IH_Epic_Summary_Table_cf_v1-0_{date_suffix}.png'
        
        self.generate_epic_summary_table_plot()
        plt.savefig(output_path, bbox_inches='tight')
        plt.close()
        print(f"Table image saved to {output_path}")

    def save_epic_summary_to_excel(self, output_excel_path: str = None):
        if output_excel_path is None:
            from datetime import datetime
            date_suffix = datetime.now().strftime('%d%m%y')
            output_excel_path = f'IH_Epic_Summary_XL_cf_v1-0_{date_suffix}.xlsx'


        #totals = self.consolidated_df[['PASSED', 'FAILED', 'BROKEN', 'SKIPPED', 'UNKNOWN', 'totalTests']].sum()
        totals = self.consolidated_df[['PASSED', 'FAILED', 'BROKEN', 'SKIPPED', 'totalTests']].sum()
        totals_row = pd.DataFrame({
            'Epic': ['TOTAL'],
            'PASSED': [totals['PASSED']],
            'FAILED': [totals['FAILED']],
            'BROKEN': [totals['BROKEN']],
            'SKIPPED': [totals['SKIPPED']],
            #clear'UNKNOWN': [totals['UNKNOWN']],
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

        # Save XL file out to another EXCEL file with a name that does not change based on date
        # This will be the standard approach for a PowerBI report to read from
        sPBI_Report_Source_Path = 'IH_Epic_ALLURE_Summary_cf_2025.xlsx'   
        wb.save(sPBI_Report_Source_Path)
        print(f"Excel file (for PowerBI use) saved to {sPBI_Report_Source_Path}")

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