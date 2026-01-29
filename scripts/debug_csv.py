
import pandas as pd
import os

def check_csv():
    csv_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'test_policy.csv'))
    print(f"Reading: {csv_path}")
    try:
        df = pd.read_csv(csv_path, sep=None, engine='python', dtype=str)
        print("Columns found:", df.columns.tolist())
        col_id = next((c for c in df.columns if c.upper() in ['CODIGO', 'USERPRINCIPALNAME', 'UPN', 'EMAIL']), None)
        print("Identified ID column:", col_id)
    except Exception as e:
        print("Error:", e)

if __name__ == "__main__":
    check_csv()
