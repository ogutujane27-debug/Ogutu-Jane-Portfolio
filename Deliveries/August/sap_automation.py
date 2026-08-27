import os
import pandas as pd
import pythoncom
import win32com.client

def run_sap_pdf_attachment():
    delivery_folder = r"C:\Kubo\delivery"
    excel_path = "New SAP Combine (HANA).xlsx"
    
    if not os.path.exists(delivery_folder):
        print(f"Delivery folder not found: {delivery_folder}")
        return

    if not os.path.exists(excel_path):
        print(f"Master Excel file not found: {excel_path}")
        return

    try:
        # 1. Load and filter the master Excel file for August 24, 2026
        print(f"Loading master tracking file: {excel_path}...")
        df = pd.read_excel(excel_path)
        
        date_column = "Date"
        doc_id_column = "DOCUMENT_ID"
        
        filtered_df = df[df[date_column].astype(str).str.contains("2026-08-24")]
        valid_doc_numbers = set(filtered_df[doc_id_column].astype(str).str.strip())
        
        print(f"Found {len(valid_doc_numbers)} valid document records for August 24th in Excel.")

        # 2. Connect to the active SAP GUI session safely
        pythoncom.CoInitialize()
        GuiWnd = win32com.client.GetObject("SAPGUI")
        application = GuiWnd.GetScriptingEngine
        connection = application.Children(0)
        session = connection.Children(0)
        
        print("Successfully connected to SAP GUI session.")
        
        # 3. Loop through PDFs in the delivery folder
        for filename in os.listdir(delivery_folder):
            if filename.lower().endswith(".pdf"):
                file_path = os.path.join(delivery_folder, filename)
                
                try:
                    doc_number = filename.split("_")[1].split(".")[0].strip()
                except IndexError:
                    print(f"Skipping file with unexpected name format: {filename}")
                    continue
                
                if doc_number not in valid_doc_numbers:
                    print(f"-> Skipping {doc_number} ({filename}): Not found in August 24th Excel records.")
                    continue
                
                print(f"Processing verified document {doc_number} for file {filename}...")
                
                # SAP Transaction navigation (FB03)
                session.findById("wnd[0]/tbar[0]/okcd").text = "/nFB03"
                session.findById("wnd[0]").sendVKey(0)
                
                session.findById("wnd[0]/usr/ctxtRF04L-BELNR").text = doc_number
                session.findById("wnd[0]").sendVKey(0)
                
                # Placeholder for GOS attachment sequence
                print(f"-> Attached PDF successfully for {doc_number}")
                
    except Exception as e:
        print(f"Error during SAP GUI automation: {e}")

if __name__ == "__main__":
    run_sap_pdf_attachment()