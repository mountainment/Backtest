import pandas as pd
from datetime import datetime
from typing import List, Dict, Optional

class ISMDataPoint:
    """Class to store a single month's ISM data with calculated score"""
    def __init__(self, date: datetime, value: float, change: Optional[float] = None):
        print(f"\nCreating new ISMDataPoint: date={date}, value={value}, change={change}")
        self.date = date
        self.value = value
        self.change = change
        self.score_info = None
        print(f"Initialized ISMDataPoint: {self.__dict__}")
    
    def calculate_score(self, previous_ism: Optional[float] = None):
        """Calculate score based on the scoring system rules"""
        print(f"\nCalculating score: value={self.value}, change={self.change}, previous_ism={previous_ism}")
        
        if self.change is None:
            print("No change value available - cannot calculate score")
            return None
        
        current_ism = self.value
        ism_change = self.change
        
        # Initialize default values
        score = 0
        bias = "Neutral"
        phase = "Normal"
        scenario = "Standard"
        
        # Check for special transitions first (scenarios 5 and 6)
        if previous_ism is not None:
            # Scenario 5: Transition from >50 to <=50 (entering contraction)
            if previous_ism > 50 and current_ism <= 50:
                scenario = "ISM troughs below 50"
                score = +8 + min(2, abs(ism_change)/2)  # -8 to -10
                bias = "Long"
                phase = "Project Diffusion"
                print(f"Scenario 5: Trough transition - raw score: {score}")
                
                self.score_info = {
                    'score': round(score, 1),
                    'bias': f"{bias} Bias",
                    'phase': phase,
                    'scenario': scenario
                }
                return self.score_info
                
            # Scenario 6: Transition from <50 to >=50 (entering expansion)
            if previous_ism < 50 and current_ism >= 50:
                scenario = "ISM peaks above 50"
                score = -8 - min(2, ism_change/2)  # +8 to +10
                bias = "Short"
                phase = "Project Initiation"
                print(f"Scenario 6: Peak transition - raw score: {score}")
                
                self.score_info = {
                    'score': round(score, 1),
                    'bias': f"{bias} Bias",
                    'phase': phase,
                    'scenario': scenario
                }
                return self.score_info
        
        # Standard scenarios (1-4)
        if current_ism > 50:
            print("ISM > 50 (Expansion territory)")
            if ism_change > 0:  # Scenario 1
                scenario = "Above 50 and growing"
                score = 5 + min(3, (ism_change / 2))  # +5 to +8
                bias = "Short"
                phase = "Initiation"
                print(f"Scenario 1: Growing expansion - raw score: {score}")
            else:  # Scenario 2
                scenario = "Above 50 and slowing"
                score = max(0, 5 + ism_change)  # 0 to +5
                bias = "Short"
                phase = "Initiation"
                print(f"Scenario 2: Slowing expansion - raw score: {score}")
        else:
            print("ISM ≤ 50 (Contraction territory)")
            if ism_change < 0:  # Scenario 3
                scenario = "Below 50 and slowing"
                score = -5 - min(3, abs(ism_change)/2)  # -5 to -8
                bias = "Long"
                phase = "Definition"
                print(f"Scenario 3: Accelerating contraction - raw score: {score}")
            else:  # Scenario 4
                scenario = "Below 50 and growing"
                score = min(0, -5 + ism_change)  # 0 to -5
                bias = "Long"
                phase = "Definition"
                print(f"Scenario 4: Improving contraction - raw score: {score}")
        
        # Round final score
        score = round(score, 1)
        print(f"Final calculated score: {score}")
        
        self.score_info = {
            'score': score,
            'bias': f"{bias} Bias",
            'phase': phase,
            'scenario': scenario
        }
        return self.score_info
    
    def to_dict(self) -> Dict:
        """Convert data point to dictionary"""
        return {
            'date': self.date.strftime('%Y-%m-%d'),
            'value': self.value,
            'change': self.change,
            'score': self.score_info['score'] if self.score_info else None,
            'bias': self.score_info['bias'] if self.score_info else None,
            'phase': self.score_info['phase'] if self.score_info else None,
            'scenario': self.score_info['scenario'] if self.score_info else None
        }

class ISMData:
    """Class to store and analyze all ISM data"""
    def __init__(self):
        print("\nInitializing new ISMData collection")
        self.data_points: List[ISMDataPoint] = []
    
    def add_data_point(self, date: datetime, value: float, change: Optional[float] = None):
        """Add a new data point to the collection"""
        print(f"\nAdding new data point: date={date}, value={value}, change={change}")
        
        # Calculate change if not provided and we have previous data
        if change is None and len(self.data_points) > 0:
            previous_value = self.data_points[-1].value
            change = value - previous_value
            print(f"Calculated change: {value} - {previous_value} = {change}")
        
        new_point = ISMDataPoint(date, value, change)
        
        # Calculate score with previous ISM value if available
        previous_ism = self.data_points[-1].value if len(self.data_points) > 0 else None
        new_point.calculate_score(previous_ism=previous_ism)
        
        self.data_points.append(new_point)
        print(f"Added new point. Total points now: {len(self.data_points)}")
    
    def load_from_excel(self, file_path: str, sheet_name: str = 'ISM') -> bool:
        """Load data from Excel file"""
        print(f"\nAttempting to load data from {file_path}, sheet {sheet_name}")
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            df.columns = df.columns.str.strip().str.lower()
            
            # Convert dates and sort
            df['date'] = pd.to_datetime(df['date'])
            df = df.sort_values('date')
            
            # Add all data points
            for _, row in df.iterrows():
                self.add_data_point(
                    date=row['date'],
                    value=row['ism'],
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

# Example usage
# ... (keep all the existing code above) ...

# Example usage
if __name__ == "__main__":
    print("=== Starting ISM Data Analysis ===")
    ism_data = ISMData()
    
    # Test with sample data
    print("\n=== Testing with sample data ===")
    ism_data.add_data_point(datetime(2023, 1, 1), 52.0)
    ism_data.add_data_point(datetime(2023, 2, 1), 53.5)
    ism_data.add_data_point(datetime(2023, 3, 1), 49.0)
    
    print("\nSample data results:")
    for data in ism_data.get_all_data():
        print(data)
    
    # Test with Excel loading
    print("\n=== Testing with Excel loading ===")
    success = ism_data.load_from_excel('C:\\Users\\emres\\Desktop\\USEndoData(1).xlsx')
    
    if success:
        print("\n=== All Monthly Data ===")
        print("[\n" + ",\n".join([str(data) for data in ism_data.get_all_data()]) + "\n]")
        
        print("\n=== Current Month Summary ===")
        current = ism_data.get_current_score()
        print(f"Date: {current['date']}")
        print(f"ISM Value: {current['value']}")
        print(f"Month-to-Month Change: {current['change']}")
        print(f"Score: {current['score']}")
        print(f"Bias: {current['bias']}")
        print(f"Phase: {current['phase']}")
        print(f"Scenario: {current['scenario']}")
    else:
        print("\nFailed to load data from Excel file")