import pdfplumber
import pandas as pd
import os

def extract_tables_from_pdfs(pdf_folder):
    for pdf_file in os.listdir(pdf_folder):
        if pdf_file.endswith(".pdf"):
            pdf_path = os.path.join(pdf_folder, pdf_file)
            all_tables = []
            with pdfplumber.open(pdf_path) as pdf:
                for page_num, page in enumerate(pdf.pages, start=1):
                    tables = page.extract_table()
                    if tables:
                        df = pd.DataFrame(tables)
                        df.insert(0, "Page_Number", page_num)  # Add column to track page number
                        
                        # Ensure unique column names
                        df.columns = [f"Column_{i}" if col is None else col for i, col in enumerate(df.columns)]
                        all_tables.append(df)
            
            if all_tables:
                final_df = pd.concat(all_tables, ignore_index=True, sort=False)
                output_excel = os.path.join(pdf_folder, f"{os.path.splitext(pdf_file)[0]}.xlsx")
                final_df.to_excel(output_excel, index=False)
                print(f"Extracted tables from {pdf_file} and saved as {output_excel}")
            else:
                print(f"No tables found in {pdf_file}.")

if __name__ == "__main__":
    pdf_folder = "iob"  # Change this to the path of your folder containing PDFs
    extract_tables_from_pdfs(pdf_folder)