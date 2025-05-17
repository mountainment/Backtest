import pandas as pd
from datetime import datetime
from typing import List, Dict, Optional

class NMIDataPoint:
    """Class to store a single month's NMI data with calculated score"""
    def __init__(self, date: datetime, value: float, change: Optional[float] = None):
        self.date = date
        self.value = value
        self.change = change
        self.score_info = None
    
    def calculate_score(self, previous_NMI: Optional[float] = None):
        """Calculate score based on the scoring system rules"""
        if self.change is None:
            return None
        
        current_NMI = self.value
        NMI_change = self.change
        
        # Initialize default values
        score = 0
        bias = "Neutral"
        phase = "Normal"
        scenario = "Standard"
        
        # Check for special transitions first (scenarios 5 and 6)
        if previous_NMI is not None:
            # Scenario 5: Transition from >50 to <=50 (entering contraction)
            if previous_NMI > 50 and current_NMI <= 50:
                scenario = "NMI troughs below 50"
                score = -8 - min(2, abs(NMI_change)/2)
                bias = "Long"
                phase = "Project Diffusion"
                
                self.score_info = {
                    'score': round(score, 1),
                    'bias': f"{bias} Bias",
                    'phase': phase,
                    'scenario': scenario
                }
                return self.score_info
                
            # Scenario 6: Transition from <50 to >=50 (entering expansion)
            if previous_NMI < 50 and current_NMI >= 50:
                scenario = "NMI peaks above 50"
                score = 8 + min(2, NMI_change/2)
                bias = "Short"
                phase = "Project Initiation"
                
                self.score_info = {
                    'score': round(score, 1),
                    'bias': f"{bias} Bias",
                    'phase': phase,
                    'scenario': scenario
                }
                return self.score_info
        
        # Standard scenarios (1-4)
        if current_NMI > 50:
            if NMI_change > 0:
                scenario = "Above 50 and growing"
                score = 5 + min(3, (NMI_change / 2))
                bias = "Short"
                phase = "Initiation"
            else:
                scenario = "Above 50 and slowing"
                score = max(0, 5 + NMI_change)
                bias = "Short"
                phase = "Initiation"
        else:
            if NMI_change < 0:
                scenario = "Below 50 and slowing"
                score = -5 - min(3, abs(NMI_change)/2)
                bias = "Long"
                phase = "Definition"
            else:
                scenario = "Below 50 and growing"
                score = min(0, -5 + NMI_change)
                bias = "Long"
                phase = "Definition"
        
        self.score_info = {
            'score': round(score, 1),
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

class NMIData:
    """Class to store and analyze all NMI data"""
    def __init__(self):
        self.data_points: List[NMIDataPoint] = []
    
    def add_data_point(self, date: datetime, value: float, change: Optional[float] = None):
        """Add a new data point to the collection"""        
        if change is None and len(self.data_points) > 0:
            previous_value = self.data_points[-1].value
            change = value - previous_value
        
        new_point = NMIDataPoint(date, value, change)
        previous_NMI = self.data_points[-1].value if len(self.data_points) > 0 else None
        new_point.calculate_score(previous_NMI=previous_NMI)
        self.data_points.append(new_point)
    
    def load_from_excel(self, file_path: str, sheet_name: str = 'NMI') -> bool:
        """Load data from Excel file with enhanced error handling"""
        try:
            # Read Excel file
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            
            # Clean column names
            df.columns = df.columns.str.strip().str.lower()
            
            # Find value column
            value_col = next((col for col in ['ism nmi', 'nmi', 'ism'] if col in df.columns), None)
            if not value_col:
                raise ValueError("Could not find NMI value column (looking for 'ISM NMI', 'NMI', or 'ISM')")
            
            # Store original index for error reporting
            df['original_index'] = df.index
            
            # Convert dates - try multiple formats
            df['date'] = pd.to_datetime(df['date'], dayfirst=True, errors='coerce')
            
            # Identify problematic rows
            bad_rows = df[df['date'].isnull()]
            if not bad_rows.empty:
                print(f"\nWarning: Could not parse dates in rows {bad_rows['original_index'].tolist()}")
                print("Problematic values:")
                for idx, row in bad_rows.iterrows():
                    print(f"Row {row['original_index']}: '{row['date']}'")
                
                # Drop problematic rows but keep processing
                df = df.dropna(subset=['date'])
            
            # Sort by date and drop any rows with missing NMI values
            df = df.sort_values('date').dropna(subset=[value_col])
            
            # Add data points
            for _, row in df.iterrows():
                try:
                    self.add_data_point(
                        date=row['date'],
                        value=row[value_col],
                        change=row.get('change')
                    )
                except Exception as e:
                    print(f"Error processing row {row['original_index']}: {str(e)}")
                    continue
                    
            return True
            
        except Exception as e:
            print(f"ERROR during data loading: {str(e)}")
            return False
    
    def get_all_data(self) -> List[Dict]:
        return [point.to_dict() for point in self.data_points]
    
    def get_last_n_months(self, n: int = 12) -> List[Dict]:
        return [point.to_dict() for point in self.data_points[-n:]]
    
    def get_current_score(self) -> Optional[Dict]:
        if not self.data_points:
            return None
        return self.data_points[-1].to_dict()

if __name__ == "__main__":
    print("=== Starting NMI Data Analysis ===")
    nmi_data = NMIData()
    
    # Test with Excel loading
    print("\n=== Loading NMI Data ===")
    file_path = r'C:\Users\emres\Desktop\USEndoData(1).xlsx'
    success = nmi_data.load_from_excel(file_path)
    
    if success:
        print("\n=== First 5 Months ===")
        for data in nmi_data.get_all_data()[:5]:
            print(data)
        
        print("\n=== Last 5 Months ===")
        for data in nmi_data.get_all_data()[-5:]:
            print(data)
        
        # Save to JSON
        import json
        with open('nmi_results.json', 'w') as f:
            json.dump(nmi_data.get_all_data(), f, indent=2)
        print("\nResults saved to nmi_results.json")
    else:
        print("\nFailed to load NMI data. Please check:")
        print("1. File exists at specified path")
        print("2. Sheet named 'NMI' exists")
        print("3. Columns named 'Date' and 'NMI' (or similar) exist")
        print("4. Date format is consistent (DD.MM.YYYY)")