import pandas as pd
import io
import os
from report import generate_report

def create_broken_excel():
    # Create a DataFrame missing the 'Work Status' column
    df = pd.DataFrame({
        'Cluster Name': ['Central', 'Eastern'],
        'PARENT_TICKET_ID': ['TT1', 'TT2'],
        'TT Type': ['P1', 'P2'],
        'Action Type': ['Type A', 'Type B'],
        'SITE_ID': ['Site 1', 'Site 2'],
        'Date': ['2026-01-01', '2026-01-02']
    })
    
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    buf.seek(0)
    return buf

def test_fix():
    print("Testing fix for missing 'Work Status' column...")
    broken_excel_buf = create_broken_excel()
    
    # Save to a temp file because generate_report expects a filepath
    tmp_path = "broken_test.xlsx"
    with open(tmp_path, "wb") as f:
        f.write(broken_excel_buf.read())
    
    try:
        # This used to raise IndexError or fail
        result = generate_report(tmp_path)
        print("SUCCESS: generate_report handled the missing column gracefully.")
    except Exception as e:
        print(f"FAILURE: generate_report failed with: {e}")
        import traceback
        traceback.print_exc()
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)

if __name__ == "__main__":
    test_fix()
