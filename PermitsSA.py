import pandas as pd
from datetime import datetime
from typing import List, Dict

class PermitDataPoint:
    """Class to store permit data with calculated score"""
    def __init__(self, date: datetime, value: float):
        self.date = date
        self.value = float(value)
        self.score = self._calculate_score()
    
    def _calculate_score(self) -> float:
        """Calculate score based on the permits value"""
        score_table = {
            400: 10, 500: 10, 600: 10, 700: 7, 800: -8,
            900: -5, 1000: -3, 1100: 0, 1200: 2, 1300: 3,
            1400: 4, 1500: 5, 1600: 8, 1700: 9, 1800: 10,
            1900: -7, 2000: -9, 2100: -10, 2200: -10, 2300: -10,
            2400: -10
        }
        
        # Find the closest threshold
        closest = min(score_table.keys(), key=lambda x: abs(x - self.value))
        return score_table[closest]
    
    def to_dict(self) -> Dict:
        """Convert to dictionary"""
        return {
            'date': self.date.strftime('%Y-%m-%d'),
            'value': self.value,
            'score': self.score
        }

def load_permits_data(file_path: str) -> List[Dict]:
    """Load and process permits data from Excel"""
    try:
        # Read Excel file
        df = pd.read_excel(file_path, sheet_name='PermitsSA')
        
        # Print columns for debugging
        print("Available columns:", df.columns.tolist())
        
        # Find the right columns
        date_col = None
        value_col = None
        
        for col in df.columns:
            if df[col].dtype == 'datetime64[ns]':
                date_col = col
            elif 'date' in str(col).lower():
                date_col = col
            elif 'value' in str(col).lower() or 'permits' in str(col).lower():
                value_col = col
        
        # Fallback to first two columns if automatic detection fails
        if date_col is None or value_col is None:
            print("Warning: Couldn't automatically detect columns, trying fallback...")
            if len(df.columns) >= 2:
                date_col = df.columns[0]
                value_col = df.columns[1]
                print(f"Using first column as date ({date_col}), second as value ({value_col})")
            else:
                raise ValueError("Not enough columns in the data")
        
        # Clean and convert data
        df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
        df = df.dropna(subset=[date_col])
        
        # Convert values to numeric, handling commas
        df[value_col] = pd.to_numeric(df[value_col].astype(str).str.replace(',', ''), errors='coerce')
        df = df.dropna(subset=[value_col])
        
        # Process data points
        results = []
        for _, row in df.iterrows():
            try:
                point = PermitDataPoint(row[date_col], row[value_col])
                results.append(point.to_dict())
            except Exception as e:
                print(f"Error processing row {row}: {str(e)}")
                continue
                
        return results
        
    except Exception as e:
        print(f"Error processing file: {str(e)}")
        return []

def print_all_scores(data: List[Dict]):
    """Print all scores in a formatted table without tabulate"""
    if not data:
        print("No data to display")
        return
    
    # Determine column widths
    date_width = max(len("Date"), max(len(d['date']) for d in data))
    value_width = max(len("Value"), max(len(f"{d['value']:,.0f}") for d in data))
    score_width = max(len("Score"), max(len(str(d['score'])) for d in data))
    
    # Create header
    header = (f"| {'Date'.ljust(date_width)} | {'Value'.rjust(value_width)} "
              f"| {'Score'.rjust(score_width)} |")
    separator = "+-" + "-"*date_width + "-+-" + "-"*value_width + "-+-" + "-"*score_width + "-+"
    
    print("\n" + separator)
    print(header)
    print(separator)
    
    # Print rows
    for record in data:
        row = (f"| {record['date'].ljust(date_width)} | "
               f"{f'{record['value']:,.0f}'.rjust(value_width)} | "
               f"{str(record['score']).rjust(score_width)} |")
        print(row)
    
    print(separator)
    print(f"\nTotal records: {len(data)}")

def save_to_csv(data: List[Dict], output_path: str):
    """Save processed data to CSV"""
    if not data:
        print("No data to save")
        return
        
    df = pd.DataFrame(data)
    df.to_csv(output_path, index=False)
    print(f"Data saved to {output_path}")

if __name__ == "__main__":
    # Configuration
    input_file = "C:\\Users\\emres\\Desktop\\USEndoData(1).xlsx"  # Update with your file path
    output_file = "permits_processed.csv"
    
    # Process data
    print("Processing permits data...")
    processed_data = load_permits_data(input_file)
    
    if processed_data:
        print(f"\nSuccessfully processed {len(processed_data)} records")
        
        # Print all scores in table format
        print_all_scores(processed_data)
        
        # Save to CSV
        save_to_csv(processed_data, output_file)
    else:
        print("No data was processed. Please check your input file.")