import pandas as pd

def extract_data_from_excel(filepath):
    """Reads the ticket data and extracts all columns."""
    try:
        print(f"Opening and parsing Excel file: {filepath}")
        
        df = pd.read_excel(filepath, header=0)
        
        # Clean the data: Drop completely empty rows and fill blanks
        df.dropna(how='all', inplace=True)
        df = df.astype(object)
        df.fillna('', inplace=True)
        
        # Convert to dictionary and grab all headers
        records = df.to_dict(orient='records')
        headers = df.columns.tolist() 
        
        print(f"[EXTRACTED] Successfully pulled {len(records)} rows across {len(headers)} columns.")
        return records, headers
        
    except Exception as e:
        print(f"[ERROR] Failed to extract data from {filepath}: {e}")
        return None, None