import os
import pandas as pd
from datetime import datetime

PDF_DIR = r"\\192.168.1.4\B1_Shr\Attachments\Acknowledged Deliveries\2026\AUGUST"
TODAY_DATE = datetime.now().date()

def check_folder_scans():
    print(f"Scanning folder directly for today ({TODAY_DATE})...")
    
    if not os.path.exists(PDF_DIR):
        print(f"[ERROR] Network directory not accessible: {PDF_DIR}")
        return

    todays_scans = []
    
    # Walk through the directory to catch files
    for root, dirs, files in os.walk(PDF_DIR):
        for file in files:
            if file.lower().endswith('.pdf'):
                fpath = os.path.join(root, file)
                stat = os.stat(fpath)
                
                # Check modification or creation date against today
                mtime = datetime.fromtimestamp(stat.st_mtime).date()
                ctime = datetime.fromtimestamp(stat.st_ctime).date()
                
                if mtime == TODAY_DATE or ctime == TODAY_DATE:
                    todays_scans.append({
                        'File_Name': file,
                        'Folder': root,
                        'Time_Added': datetime.fromtimestamp(max(stat.st_mtime, stat.st_ctime)).strftime('%H:%M:%S')
                    })

    print("\n" + "=" * 65)
    print(f"DIRECT FOLDER SCAN REPORT FOR TODAY ({TODAY_DATE})")
    print("=" * 65)
    print(f"Total PDFs with Today's Timestamp in Folder: {len(todays_scans)}")
    print("-" * 65)
    
    if todays_scans:
        df_today = pd.DataFrame(todays_scans)
        print(df_today[['File_Name', 'Time_Added']].head(25).to_string(index=False))
        
        output_csv = r"C:\Kubo\action_today_folder_scans.csv"
        df_today.to_csv(output_csv, index=False)
        print(f"\n[OK] Saved today's folder scan list to: {output_csv}")
    else:
        print("Hata kwenye folda yenyewe hakuna faili lililojichapisha na tarehe ya leo.")
        print("(Kama zimescanwa leo lakini zimekuja na tarehe ya jana au siku iliyopita kutokana na settings za scanner, zipo kwenye list kuu ya 1,317 matched).")
    print("=" * 65)

if __name__ == "__main__":
    check_folder_scans()