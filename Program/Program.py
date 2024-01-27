import pandas as pd
from datetime import datetime, timedelta

class Employee:
    def __init__(self, name, position, start_datetime, end_datetime):
        self.name = name
        self.position = position
        self.start_datetime = start_datetime
        self.end_datetime = end_datetime

    def get_duration(self):
        return self.end_datetime - self.start_datetime

class EmployeeAnalyzer:
    def __init__(self):
        self.employee_data = []
        self.end_datetime_previous_shift = None

    def read_employee_data(self, file_path):
        try:
            df = pd.read_excel(file_path, engine='openpyxl')
            df.columns = map(str.lower, df.columns)  # Convert column names to lowercase

            for index, row in df.iterrows():
                name = row['employee name']
                position = row['position id']

                # Check if 'time' and 'time out' are not missing (NaT)
                if not pd.isna(row['time']) and not pd.isna(row['time out']):
                    start_datetime_str = str(row['time'])
                    end_datetime_str = str(row['time out'])

                    try:
                        start_datetime = datetime.strptime(start_datetime_str, '%Y-%m-%d %H:%M:%S')
                        end_datetime = datetime.strptime(end_datetime_str, '%Y-%m-%d %H:%M:%S')
                    except ValueError as ve:
                        print(f"Error parsing datetime in row {index}: {ve}")
                        continue

                    employee = Employee(name, position, start_datetime, end_datetime)
                    self.employee_data.append(employee)

        except Exception as e:
            print(f"Error reading Excel file: {e}")

    def analyze_data(self):
        for employee in self.employee_data:
            # Analyze the data
            if employee.get_duration() >= timedelta(days=7):
                print(f"{employee.name} ({employee.position}) has worked for 7 consecutive days.")

            time_between_shifts = employee.start_datetime - (self.end_datetime_previous_shift or employee.start_datetime)
            if timedelta(hours=1) < time_between_shifts < timedelta(hours=10):
                print(f"{employee.name} ({employee.position}) has less than 10 hours between shifts but greater than 1 hour.")

            if employee.get_duration() > timedelta(hours=14):
                print(f"{employee.name} ({employee.position}) has worked for more than 14 hours in a single shift.")

            self.end_datetime_previous_shift = employee.end_datetime

analyzer = EmployeeAnalyzer()
analyzer.read_employee_data('excel.xlsx')
analyzer.analyze_data()
