import os
import pandas as pd

REPORT_PATH = r"C:\Kubo\kubo_document_matching_report.csv"
OUTPUT_MISSING = r"C:\Kubo\action_missing_pdfs.csv"
OUTPUT_UNTRACKED = r"C:\Kubo\action_untracked_pdfs.csv"

def run_automation():
    print("Starting Kubo automated discrepancy processor...")
    
    if not os.path.exists(REPORT_PATH):
        print(f"[ERROR] Main report not found at {REPORT_PATH}. Run Kubo_daily_check.py first.")
        return

    # Load the matching report CSV
    df = pd.read_csv(REPORT_PATH)
    
    # Filter missing PDFs
    missing_df = df[df['Status'] == 'MISSING_PDF']
    missing_df.to_csv(OUTPUT_MISSING, index=False)
    print(f"[OK] Exported {len(missing_df)} missing PDF records to: {OUTPUT_MISSING}")
    
    # Filter untracked PDFs (including duplicates if any)
    untracked_df = df[df['Status'].str.contains('UNTRACKED', na=False)]
    untracked_df.to_csv(OUTPUT_UNTRACKED, index=False)
    print(f"[OK] Exported {len(untracked_df)} untracked PDF records to: {OUTPUT_UNTRACKED}")
    
    print("\n--- AUTOMATION SUMMARY ---")
    print(f"Total Missing SAP Docs requiring physical scans: {len(missing_df)}")
    print(f"Total Untracked PDFs needing review/re-labeling: {len(untracked_df)}")
    print("Action lists are ready for team handling!")

if __name__ == "__main__":
    run_automation()