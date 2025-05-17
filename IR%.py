import pandas as pd
from datetime import datetime
from typing import List, Dict

class InterestRateDataPoint:
    """Class to store interest rate data with calculated score"""
    def __init__(self, date: datetime, rate_value: float):
        self.date = date
        self.rate_value = float(rate_value)
        self.score = self._calculate_score()
    
    def _calculate_score(self) -> float:
        """Calculate score based on the raw rate value (not percentage)"""
        # Scoring table (raw rate value thresholds and corresponding scores)
        score_table = {
            60: -10, 50: -10, 40: -10, 30: -10,
            20: -10, 15: -9, 10: -6, 5: -3,
            0: 0, -5: 0, -10: 3, -15: 5,
            -20: 7, -30: -8, -40: -9, -50: -10, -60: -10
        }
        
        # Find the closest threshold
        closest = min(score_table.keys(), key=lambda x: abs(x - self.rate_value))
        return score_table[closest]
    
    def to_dict(self) -> Dict:
        """Convert to dictionary"""
        return {
            'date': self.date.strftime('%Y-%m-%d'),
            'rate_value': self.rate_value,
            'score': self.score
        }

def load_interest_rate_data(file_path: str) -> List[Dict]:
    """Load and process interest rate data from Excel"""
    try:
        # Read Excel file
        df = pd.read_excel(file_path, sheet_name='IR%')
        
        # Print columns for debugging
        print("Available columns:", df.columns.tolist())
        
        # Use column A for date and column E for RateValue
        if len(df.columns) < 5:
            raise ValueError("Data doesn't have enough columns (need at least 5 columns)")
        
        date_col = df.columns[0]  # Column A (first column)
        rate_value_col = df.columns[4]  # Column E (5th column)
        
        print(f"Using column '{date_col}' for dates and column '{rate_value_col}' for rate values")
        
        # Clean and convert data
        df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
        df = df.dropna(subset=[date_col])
        
        # Convert rate values to numeric (treating them as raw numbers, not percentages)
        df[rate_value_col] = pd.to_numeric(
            df[rate_value_col].astype(str).str.replace('%', ''),
            errors='coerce'
        )*100
        df = df.dropna(subset=[rate_value_col])
        
        # Process data points
        results = []
        for _, row in df.iterrows():
            try:
                point = InterestRateDataPoint(row[date_col], row[rate_value_col])
                results.append(point.to_dict())
            except Exception as e:
                print(f"Error processing row {row}: {str(e)}")
                continue
                
        return results
        
    except Exception as e:
        print(f"Error processing file: {str(e)}")
        return []

def print_all_scores(data: List[Dict]):
    """Print all scores in a formatted table"""
    if not data:
        print("No data to display")
        return
    
    # Determine column widths
    date_width = max(len("Date"), max(len(d['date']) for d in data))
    rate_width = max(len("Rate Value"), max(len(f"{d['rate_value']:.2f}") for d in data))
    score_width = max(len("Score"), max(len(str(d['score'])) for d in data))
    
    # Create header
    separator = "+-" + "-"*date_width + "-+-" + "-"*rate_width + "-+-" + "-"*score_width + "-+"
    header = (f"| {'Date'.ljust(date_width)} | {'Rate Value'.center(rate_width)} "
              f"| {'Score'.rjust(score_width)} |")
    
    print("\n" + separator)
    print(header)
    print(separator)
    
    # Print rows
    for record in data:
        row = (f"| {record['date'].ljust(date_width)} | "
               f"{f'{record['rate_value']:.2f}'.rjust(rate_width)} | "
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
    output_file = "interest_rates_processed.csv"
    
    # Process data
    print("Processing interest rate data...")
    processed_data = load_interest_rate_data(input_file)
    
    if processed_data:
        print(f"\nSuccessfully processed {len(processed_data)} records")
        
        # Print all scores in table format
        print_all_scores(processed_data)
        
        # Save to CSV
        save_to_csv(processed_data, output_file)
    else:
        print("No data was processed. Please check your input file.")