import os
import pandas as pd

def run_offline_reconciliation():
    delivery_folder = r"C:\Kubo\delivery"
    excel_path = "New SAP Combine (HANA).xlsx"
    report_path = "reconciliation_audit_report.csv"
    
    if not os.path.exists(delivery_folder):
        print(f"[Error] Delivery folder not found: {delivery_folder}")
        return

    if not os.path.exists(excel_path):
        print(f"[Error] Master Excel file not found: {excel_path}")
        return

    print(f"Loading master SAP HANA export: {excel_path}...")
    try:
        df = pd.read_excel(excel_path)
    except Exception as e:
        print(f"[Error] Could not read Excel file: {e}")
        return
    
    # Standard column name checks (adjust if your headers differ)
    date_column = "Date"
    doc_id_column = "DOCUMENT_ID"
    
    if date_column not in df.columns or doc_id_column not in df.columns:
        print(f"[Error] Expected columns '{date_column}' or '{doc_id_column}' not found.")
        print(f"Available columns in your Excel: {list(df.columns)}")
        return

    # Filter for August 24, 2026
    target_date = "2026-08-24"
    filtered_df = df[df[date_column].astype(str).str.contains(target_date)]
    valid_doc_numbers = set(filtered_df[doc_id_column].astype(str).str.strip())
    
    print(f"Loaded {len(valid_doc_numbers)} valid document records for {target_date}.")

    audit_results = []
    pdf_files = [f for f in os.listdir(delivery_folder) if f.lower().endswith(".pdf")]
    
    if not pdf_files:
        print("No PDF files found in the C:\\Kubo\\delivery folder.")
        return

    print(f"Scanning {len(pdf_files)} local delivery PDFs...")
    for filename in pdf_files:
        # Extract document number from filename (assuming format like INV_10023.pdf)
        try:
            doc_number = filename.split("_")[1].split(".")[0].strip()
        except IndexError:
            audit_results.append({
                "Filename": filename,
                "Document_ID": "UNKNOWN",
                "Status": "Invalid Filename Format"
            })
            continue
        
        # Cross-reference against HANA export data
        if doc_number in valid_doc_numbers:
            status = "Verified Match (August 24)"
        else:
            status = "Not Found in August 24 Records"
            
        audit_results.append({
            "Filename": filename,
            "Document_ID": doc_number,
            "Status": status
        })

    # Save output audit report
    audit_df = pd.DataFrame(audit_results)
    audit_df.to_csv(report_path, index=False)
    
    print("\n" + "="*50)
    print("           RECONCILIATION AUDIT SUMMARY")
    print("="*50)
    print(audit_df.to_string(index=False))
    print("="*50)
    print(f"\nAudit report successfully exported to: C:\\Kubo\\{report_path}\n")

if __name__ == "__main__":
    run_offline_reconciliation()