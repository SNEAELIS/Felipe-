import os
import sys
from _datetime import datetime

import pandas as pd
import numpy as np
from openpyxl.styles import Font


def analyze_attendance_data(dir_path):
    """
    Analyzes attendance data ('P' for Present, 'F' for Fault) from Excel files.

    Args:
        dir_path: The parent directory path where all the .xlsx files are stored
    """


    def helper_calculate(df, overall_atend_avg):
        """
        Calculate comprehensive attendance statistics.
        Returns a dictionary of DataFrames with statistics.
        """
        df_cleaned = df.map(lambda x: str(x).strip().upper() if pd.notna(x) else '')
        df_cleaned = df_cleaned[df_cleaned.iloc[:, 1].str.strip().astype(bool)]

        pf_cols = [col for col in df_cleaned.columns if (df_cleaned[col].isin(['P', 'F'])).any()]

        if not pf_cols:
            return None

        # Student-level stats
        student_stats = pd.DataFrame({
            'Student': df_cleaned.iloc[:, 1] if len(df_cleaned.columns) > 1 else np.arange(1,
                                                                                     len(df_cleaned) + 1),
            'Total_presents': (df_cleaned[pf_cols] == 'P').sum(axis=1),
            'Total_absences': (df_cleaned[pf_cols] == 'F').sum(axis=1),
            'Attendance_rate': ((df_cleaned[pf_cols] == 'P').mean(axis=1) * 100).round(1)
        })

        # Summary stats
        total_days = len(pf_cols)
        total_students = len(df_cleaned)
        total_presents = (df_cleaned[pf_cols] == 'P').sum().sum()
        total_absences = (df_cleaned[pf_cols] == 'F').sum().sum()

        summary_stats = pd.DataFrame({
            'Metric': ['Total Students', 'Total Days', 'Total Presents', 'Total Absences',
                       'Overall Attendance Rate'],
            'Value': [
                total_students,
                total_days,
                total_presents,
                total_absences,
                (total_presents / (total_students * total_days) * 100).round(1) if (
                 total_students * total_days) > 0 else 0
            ]
        })

        # Append to NumPy array
        overall_atend_avg.append(summary_stats['Value'].iloc[-1])

        return {
            'summary': summary_stats,
            'students': student_stats
        }


    def helper_write_to_sheet(hwriter, hstats, hsheet_name):
        """Write formatted statistics to a new sheet using only pandas"""
        try:
            # Validate input data
            if not hstats or 'summary' not in hstats or 'students' not in hstats:
                print(f"Skipping {hsheet_name} - invalid stats format")
                return

            # Combine all stats into one DataFrame with sections
            summary_df = hstats['summary']
            students_df = hstats['students']

            # Check for empty DataFrames
            if summary_df is None or len(summary_df) == 0:
                print(f"Skipping {hsheet_name} - no summary data")
                return

            if students_df is None or len(students_df) == 0:
                print(f"Skipping {hsheet_name} - no student data")
                return

            # Create section headers with safe length calculation
            summary_header = pd.DataFrame([["SUMMARY STATISTICS"]], columns=['Metric'])
            student_header = pd.DataFrame([["STUDENT ATTENDANCE"]], columns=['Student'])

            # Create a spacer row
            spacer = pd.DataFrame([['---'] * len(summary_df.columns)], columns=summary_df.columns)

            # Combine all components
            final_df = pd.concat([
                summary_header,
                summary_df,
                spacer,
                student_header,
                students_df,
                spacer
                ],
                ignore_index=True)

            # Write to Excel
            final_df.to_excel(
                hwriter,
                sheet_name=hsheet_name,
                index=False,
                header=True
            )
            workbook = hwriter.book
            worksheet = workbook[hsheet_name]

            # Bold headers (pandas-accessible formatting)
            for cell in worksheet[1]:  # Header row
                cell.font = Font(bold=True)
            worksheet['A1'].font = Font(bold=True)
            worksheet.cell(row=len(summary_df) + 3, column=1).font = Font(bold=True)

            # Auto-adjust column widths (still requires openpyxl access)
            for column in worksheet.columns:
                max_length = max(len(str(cell.value)) for cell in column)
                worksheet.column_dimensions[column[0].column_letter].width = max_length + 2
        except Exception as e:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"Error occurred at line: {exc_tb.tb_lineno}")
            print(f"Error writing to sheet {filename}:\n {str(e)[:100]}")

    overall_avg = []
    summary_file = os.path.join(dir_path, "average_attendance.txt")

    try:
        for root, dirs, files in os.walk(dir_path):
            for filename in files:
                if filename.endswith('.xlsx'):
                    xlsx_path = os.path.join(root, filename)
                    print(f"Processing {filename}...")

                    try:
                        with pd.ExcelFile(xlsx_path) as xls:
                            sheet_names = xls.sheet_names

                            with pd.ExcelWriter(xlsx_path, engine='openpyxl', mode='a',
                                                if_sheet_exists='replace') as writer:
                                for sheet_name in sheet_names:
                                    df = pd.read_excel(xls, sheet_name=sheet_name)

                                    stats = helper_calculate(df, overall_avg)

                                    new_sheet_name = f'Stats_{sheet_name}'

                                    helper_write_to_sheet(writer, stats, new_sheet_name)
                        print(f"Added statistics sheet to {filename}")
                    except Exception as e:
                        exc_type, exc_value, exc_tb = sys.exc_info()
                        print(f"Error occurred at line: {exc_tb.tb_lineno}")
                        print(f"Error processing {filename}: {str(e)}")
    except FileNotFoundError:
        print(f"Error: The directory '{dir_path}' was not found.")
    except Exception as e:
        print(f'An unexpected error occurred:{type(e).__name__}\n{str(e)[:100]}')
    try:
        # Calculate average using NumPy
        print(overall_avg)
        if overall_avg:  # Check if array is not empty
            avg_attendance = sum(overall_avg)/len(overall_avg)
            with open(summary_file, 'w') as f:
                f.write(f"Average Overall Attendance Rate: {avg_attendance:.2f}%\n")
                f.write(f"Calculated from {len(overall_avg)} classes/sheets\n")
                f.write(f"Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
    except Exception as e:
        print(f'An unexpected error occurred:{type(e).__name__}\n{str(e)[:100]}')
def main():
    dir_path = (r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência "
                r"Social\SNEAELIS - Termo Fomento Inst. Léo Moura")
    analyze_attendance_data(dir_path)


if __name__ == "__main__":
    main()