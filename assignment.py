import pandas as pd
from datetime import timedelta

def calculate_hours(row):
    """Calculate the number of hours worked in a shift."""
    return (row['Time Out'] - row['Time']).total_seconds() / 3600

def check_criteria(df):
    """Check the specified criteria for each employee."""
    results = {
        "seven_consecutive_days": [],
        "less_than_10_hours_between_shifts": [],
        "more_than_14_hours_single_shift": []
    }

    current_employee = None
    consecutive_days = 1
    previous_day = None
    previous_shift_end = None

    for index, row in df.iterrows():
        employee = row['Employee Name']
        shift_start = row['Time']
        shift_end = row['Time Out']
        hours_worked = row['Hours Worked']

        # New employee
        if employee != current_employee:
            consecutive_days = 1
            previous_day = None
            previous_shift_end = None
            current_employee = employee

        # More than 14 hours in a single shift
        if hours_worked > 14:
            results["more_than_14_hours_single_shift"].append(employee)

        # Consecutive days check
        if previous_day is not None and (shift_start.date() - previous_day).days == 1:
            consecutive_days += 1
            if consecutive_days == 7:
                results["seven_consecutive_days"].append(employee)
        else:
            consecutive_days = 1

        # Less than 10 hours but more than 1 hour between shifts
        if previous_shift_end is not None:
            hours_between_shifts = (shift_start - previous_shift_end).total_seconds() / 3600
            if 1 < hours_between_shifts < 10:
                results["less_than_10_hours_between_shifts"].append(employee)

        previous_day = shift_start.date()
        previous_shift_end = shift_end

    return results

def main(file_path):
    # Load the Excel file
    timecard_df = pd.read_excel(file_path)

    # Preprocess: Remove rows with NaN in critical columns
    preprocessed_df = timecard_df.dropna(subset=['Time', 'Time Out', 'Employee Name'])

    # Add a column for hours worked per shift
    preprocessed_df['Hours Worked'] = preprocessed_df.apply(calculate_hours, axis=1)

    # Sort by Employee Name and Time for sequential analysis
    preprocessed_df.sort_values(by=['Employee Name', 'Time'], inplace=True)

    # Analyze the data
    results = check_criteria(preprocessed_df)

    # Write the results to a file
    with open('output.txt', 'w') as file:
        file.write(str(results))

# Replace 'your_file_path_here.xlsx' with the path to your Excel file
file_path = 'your_file_path_here.xlsx'
main(file_path)
