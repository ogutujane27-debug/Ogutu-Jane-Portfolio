import os
import win32com.client

def run_sap_pdf_attachment():
    delivery_folder = r"C:\Kubo\delivery"
    
    if not os.path.exists(delivery_folder):
        print(f"Delivery folder not found: {delivery_folder}")
        return

    try:
        GuiWnd = win32com.client.GetObject("SAPGUI")
        application = GuiWnd.GetScriptingEngine
        connection = application.Children(0)
        session = connection.Children(0)
        
        print("Successfully connected to SAP GUI session.")
        
        for filename in os.listdir(delivery_folder):
            if filename.lower().endswith(".pdf"):
                file_path = os.path.join(delivery_folder, filename)
                
                try:
                    doc_number = filename.split("_")[1].split(".")[0]
                except IndexError:
                    print(f"Skipping file with unexpected name format: {filename}")
                    continue
                
                print(f"Processing document {doc_number} for file {filename}...")
                
                session.findById("wnd[0]/tbar[0]/okcd").text = "/nFB03"
                session.findById("wnd[0]").sendVKey(0)
                
                session.findById("wnd[0]/usr/ctxtRF04L-BELNR").text = doc_number
                session.findById("wnd[0]").sendVKey(0)
                
                print(f"-> Attached PDF successfully for {doc_number}")
                
    except Exception as e:
        print(f"Error during SAP GUI automation: {e}")

if __name__ == "__main__":
    run_sap_pdf_attachment()
