import pandas as pd
from datetime import datetime
from typing import List, Dict, Optional

class MonetaryDataPoint:
    """Class to store monetary data with calculated score"""
    def __init__(self, date: datetime, value: float):
        self.date = date
        self.value = value
        self.three_month_avg = None
        self.percent_change = None
        self.score = None
    
    def calculate_metrics(self, previous_points: List['MonetaryDataPoint']):
        """Calculate 3-month average and score"""
        if len(previous_points) >= 2:
            # Get current and previous 2 months' values
            recent_values = [point.value for point in previous_points[-2:]] + [self.value]
            
            # Calculate simple 3-month moving average
            self.three_month_avg = sum(recent_values) / 3
            
            # Calculate score based on current value
            self.calculate_score()  # Call the score calculation here
            
            # Debug print to verify calculations
            print(f"\nCalculating for {self.date.strftime('%Y-%m')}:")
            print(f"Current value: {self.value}")
            print(f"Previous 2 values: {[point.value for point in previous_points[-2:]]}")
            print(f"3-month average: {self.three_month_avg:.2f}")
            print(f"Calculated score: {self.score}")

    def calculate_score(self):
        """Calculate score based on the absolute value"""
        # Remove percent_change check since we're using absolute values
        if self.value is None:
            return None
        
        # Adjust these thresholds based on your scoring needs
        if self.value >= 0.50: score = 4
        elif self.value >= 0.40: score = 6
        elif self.value >= 0.30: score = 10
        elif self.value >= 0.20: score = 8
        elif self.value >= 0.15: score = 6
        elif self.value >= 0.10: score = 4
        elif self.value >= 0.5: score = 0
        elif self.value >= 0.0: score = -4
        elif self.value >= -0.5: score = -8
        elif self.value >= -0.10: score = -10
        elif self.value >= -0.15: score = 8
        elif self.value >= -0.20: score = 10
        elif self.value >= -0.30: score = 10
        elif self.value >= -0.40: score = 10
        else: score = 0.10
        
        self.score = score
        return score
        
    def to_dict(self) -> Dict:
        """Convert data point to dictionary"""
        return {
            'date': self.date.strftime('%Y-%m-%d'),
            'value': self.value,
            '3_month_avg': round(self.three_month_avg, 2) if self.three_month_avg else None,
            'percent_change': round(self.percent_change, 2) if self.percent_change else None,
            'score': self.score
        }

class MonetaryData:
    """Class to store and analyze all monetary data"""
    def __init__(self):
        self.data_points: List[MonetaryDataPoint] = []
    
    def add_data_point(self, date: datetime, value: float):
        """Add a new data point to the collection"""
        new_point = MonetaryDataPoint(date, value)
        new_point.calculate_metrics(self.data_points)
        if new_point.percent_change is not None:
            new_point.calculate_score()
        self.data_points.append(new_point)
    
    def load_from_excel(self, file_path: str, sheet_name: str = None) -> bool:
        """Load data from Excel file with specific column mapping"""
        try:
            # Read Excel file with headers in row 8 (0-indexed 7)
            # Explicitly specify which columns to read (A for date, F for value)
            df = pd.read_excel(
                file_path,
                sheet_name=sheet_name,
                header=7,
                usecols="A,F",  # Column A (DATE) and F (VALUE)
                names=['date', 'value']  # Name the columns explicitly
            )
            
            # Convert dates
            df['date'] = pd.to_datetime(df['date'], errors='coerce')
            
            # Convert values to numeric
            df['value'] = pd.to_numeric(df['value'], errors='coerce')
            
            # Drop rows with invalid data
            df = df.dropna(subset=['date', 'value'])
            
            # Add data points
            for _, row in df.iterrows():
                self.add_data_point(
                    date=row['date'],
                    value=row['value']
                )
            return True
        except Exception as e:
            print(f"ERROR during data loading: {str(e)}")
            return False
    
    def get_all_data(self) -> List[Dict]:
        return [point.to_dict() for point in self.data_points if point.percent_change is not None]
    
    def get_last_n_months(self, n: int = 12) -> List[Dict]:
        valid_points = [p for p in self.data_points if p.percent_change is not None]
        return [point.to_dict() for point in valid_points[-n:]]
    
    def get_current_score(self) -> Optional[Dict]:
        valid_points = [p for p in self.data_points if p.percent_change is not None]
        return valid_points[-1].to_dict() if valid_points else None

if __name__ == "__main__":
    print("=== Starting Monetary Data Analysis ===")
    monetary_data = MonetaryData()
    
    print("\n=== Loading M2 Data ===")
    file_path = r'C:\Users\emres\Desktop\USEndoData(1).xlsx'
    sheet_name = "M2"  # Use the exact sheet name from your file
    
    try:
        with pd.ExcelFile(file_path) as xl:
            print(f"Available sheets: {xl.sheet_names}")
            
            if sheet_name not in xl.sheet_names:
                print(f"Error: Sheet '{sheet_name}' not found")
            else:
                success = monetary_data.load_from_excel(file_path, sheet_name=sheet_name)
                
                if success:
                    
                        
                    print("\n=== First 5 Calculated Months ===")
                    for data in monetary_data.get_all_data()[:5]:
                        print(data)
                    
                    print("\n=== Last 5 Calculated Months ===")
                    for data in monetary_data.get_all_data()[-5:]:
                        print(data)
                    
                    # Save to JSON
                    import json
                    with open('m2_results.json', 'w') as f:
                        json.dump(monetary_data.get_all_data(), f, indent=2)
                    print("\nResults saved to m2_results.json")
                else:
                    print("\nFailed to load data. Please check:")
                    print("1. Data starts at row 8 with headers in row 8")
                    print("2. There is a date column and a value column")
                    print("3. No merged cells or other formatting issues")
    except Exception as e:
        print(f"Error opening Excel file: {str(e)}")