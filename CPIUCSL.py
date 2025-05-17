import pandas as pd
from datetime import datetime
from typing import List, Dict

class CPIDataPoint:
    """Class to store CPI data with calculated 12-month average and score"""
    def __init__(self, date: datetime, cpi_change: float):
        self.date = date
        self.cpi_change = float(cpi_change)
        self.score = None
    
    def calculate_score(self, cpi_12ma: float):
        """Calculate score based on the 12-month average CPI change"""
        # Scoring table (CPI thresholds and corresponding scores)
        score_table = [
            (1.4, 3), (1.2, 6), (1.0, 10), (0.8, 8), (0.6, 6),
            (0.4, 4), (0.3, 0), (0.0, 0), (-0.1, -6), (-0.2, -8),
            (-0.3, -10), (-0.4, 6), (-0.5, 8), (-0.6, 10), (-0.7, 10)
        ]
        
        # Handle zero values explicitly
        if abs(cpi_12ma) < 0.01:
            self.score = 0
            return self.score
            
        # Find the closest threshold
        closest = min(score_table, key=lambda x: abs(x[0] - cpi_12ma))
        self.score = closest[1]
        return self.score
    
    def to_dict(self) -> Dict:
        """Convert to dictionary"""
        return {
            'date': self.date.strftime('%Y-%m-%d'),
            'cpi_change': self.cpi_change,
            'score': self.score
        }

def load_cpi_data(file_path: str) -> List[Dict]:
    """Load and process CPI data from Excel"""
    try:
        # Read Excel file
        df = pd.read_excel(file_path, sheet_name='CPIAUCSL')
        
        # Clean column names and find relevant columns
        df.columns = df.columns.str.strip()
        date_col = 'DATE'
        change_col = 'CPI_CHANGE'
        
        # Convert and clean data
        df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
        df = df.dropna(subset=[date_col])
        
        # Convert percentage changes to numeric (remove % sign)
        df[change_col] = pd.to_numeric(
            df[change_col].astype(str).str.replace('%', ''),
            errors='coerce'
        )
        df = df.dropna(subset=[change_col])
        df = df.sort_values(date_col)
        
        # Calculate 12-month moving average of changes
        df['12ma'] = df[change_col].rolling(12).mean()
        
        # Process data points (only those with 12ma values)
        results = []
        for i, row in df.iterrows():
            if not pd.isna(row['12ma']):
                point = CPIDataPoint(row[date_col], row[change_col])
                point.calculate_score(row['12ma'])
                results.append(point.to_dict())
                
        return results
        
    except Exception as e:
        print(f"Error processing file: {str(e)}")
        return []

def print_all_scores(data: List[Dict]):
    """Print all scores in a clean table format"""
    if not data:
        print("No data to display")
        return
    
    # Print header
    print("\nDate\t\tCPI Change\t12-Month Avg\tScore")
    print("------------------------------------------------")
    
    # Print rows
    for record in data:
        print(f"{record['date']}\t{record['cpi_change']:.2f}%\t\t{record.get('12ma', 'N/A'):.2f}%\t\t{record['score']}")

if __name__ == "__main__":
    # Configuration
    input_file = "C:\\Users\\emres\\Desktop\\USEndoData(1).xlsx"
    
    # Process data
    print("Processing CPI data...")
    processed_data = load_cpi_data(input_file)
    
    if processed_data:
        print(f"\nSuccessfully processed {len(processed_data)} records")
        print_all_scores(processed_data)
    else:
        print("No data was processed. Please check your input file.")