import pandas as pd
from datetime import datetime
from typing import List, Dict, Optional

class UMCSIDataPoint:
    """Class to store a single month's UMCSI data with calculated score"""
    def __init__(self, date: datetime, value: float, change: Optional[float] = None):
        self.date = date
        self.value = value
        self.change = change
        self.score = None
    
    def calculate_score(self):
        """Calculate score based on the UMCSI scoring system"""
        if self.value is None:
            return None
        
        # UMCSI scoring mapping
        score_map = {
            55: 10, 56: 10, 57: 9, 58: 9, 59: 8, 60: 7, 61: 6, 62: 5,
            63: -10, 64: -9, 65: -8, 66: -7, 67: -6, 68: -5, 69: -4,
            70: -3, 71: -2, 72: -1, 73: -1, 74: 0, 75: 0, 76: 1, 77: 1,
            78: 2, 79: 2, 80: 3, 81: 3, 82: 4, 83: 4, 84: 5, 85: 5,
            86: 6, 87: 6, 88: 7, 89: 7, 90: 8, 91: 8, 92: 9, 93: 9,
            94: 10, 95: 10, 96: -5, 97: -5, 98: -6, 99: -6, 100: -7,
            101: -8, 102: -9, 103: -9, 104: -10, 105: -10
        }
        
        # Round to nearest integer and clamp within 55-105 range
        rounded_value = int(round(self.value))
        clamped_value = max(55, min(105, rounded_value))
        self.score = score_map.get(clamped_value, 0)
        return self.score
    
    def to_dict(self) -> Dict:
        """Convert data point to dictionary"""
        return {
            'date': self.date.strftime('%Y-%m-%d'),
            'value': self.value,
            'change': self.change,
            'score': self.score
        }

class UMCSIData:
    """Class to store and analyze all UMCSI data"""
    def __init__(self):
        self.data_points: List[UMCSIDataPoint] = []
    
    def add_data_point(self, date: datetime, value: float, change: Optional[float] = None):
        """Add a new data point to the collection"""
        if change is None and len(self.data_points) > 0:
            previous_value = self.data_points[-1].value
            change = value - previous_value
        
        new_point = UMCSIDataPoint(date, value, change)
        new_point.calculate_score()
        self.data_points.append(new_point)
    
    def load_from_excel(self, file_path: str, sheet_name: str = 'UMCSI') -> bool:
        """Load data from Excel file"""
        try:
            # Read Excel file
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            
            # Clean column names
            df.columns = df.columns.str.strip().str.lower()
            
            # Convert dates - handle US format (MM/DD/YYYY)
            df['date'] = pd.to_datetime(df['date'])
            
            # Sort by date and drop any rows with missing values
            df = df.sort_values('date').dropna(subset=['date', 'umcsi'])
            
            # Add data points
            for _, row in df.iterrows():
                self.add_data_point(
                    date=row['date'],
                    value=row['umcsi'],
                    change=row.get('change')
                )
            return True
        except Exception as e:
            print(f"ERROR during data loading: {str(e)}")
            return False
    
    def get_all_data(self) -> List[Dict]:
        """Get all data as list of dictionaries"""
        return [point.to_dict() for point in self.data_points]
    
    def get_last_n_months(self, n: int = 12) -> List[Dict]:
        """Get data for last N months"""
        return [point.to_dict() for point in self.data_points[-n:]]
    
    def get_current_score(self) -> Optional[Dict]:
        """Get the most recent score"""
        if not self.data_points:
            return None
        return self.data_points[-1].to_dict()

if __name__ == "__main__":
    print("=== Starting UMCSI Data Analysis ===")
    umcsi_data = UMCSIData()
    
    # Test with Excel loading
    print("\n=== Loading UMCSI Data ===")
    file_path = 'C:\\Users\\emres\Desktop\\USEndoData(1).xlsx'
    success = umcsi_data.load_from_excel(file_path, sheet_name='UMCSI')
    
    if success:
        print("\n=== First 5 Months ===")
        for data in umcsi_data.get_all_data()[:5]:
            print(data)
        
        print("\n=== Last 5 Months ===")
        for data in umcsi_data.get_all_data()[-5:]:
            print(data)
        
        # Save to JSON
        import json
        with open('umcsi_results.json', 'w') as f:
            json.dump(umcsi_data.get_all_data(), f, indent=2)
        print("\nResults saved to umcsi_results.json")
    else:
        print("\nFailed to load UMCSI data. Please check:")
        print("1. File exists at specified path")
        print("2. Sheet named 'UMCSI' exists")
        print("3. Columns named 'Date' and 'UMCSI' exist")
        print("4. Date format is consistent (MM/DD/YYYY)")