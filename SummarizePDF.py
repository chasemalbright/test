import pandas as pd
import fitz  # PyMuPDF
import re
import os
import math
from tkinter import Tk, filedialog

def select_pdf_file():
    root = Tk()
    root.withdraw()  # Hide the main window
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    return file_path

def select_csv_file():
    root = Tk()
    root.withdraw()  # Hide the main window
    file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
    return file_path


def calculate_distance(x1, y1, x2, y2):
    distance = math.sqrt((x2 - x1)**2 + (y2 - y1)**2)
    return distance

def extract_annotations(pdf_path):
    
     # get structure friom pdf
    doc = fitz.open(pdf_path)
    annotations = []
    cat = doc.pdf_catalog() 
    bsi_definition = (doc.xref_get_key(cat, "BSIAnnotColumns"))
    bsi_xref_str=bsi_definition[1].split()
    print(bsi_xref_str)
    bsi_xref=int(bsi_xref_str[0])
    ColumnName = doc.xref_object(bsi_xref)

    pattern = r'/Name \((.*?)\)\n'
    
    # Extract values using the pattern
    xref_values = re.findall(pattern, ColumnName)
    
    

    for page_number, page in enumerate(doc, start=1):
        for annot in page.annots():
            if annot.type[1] == 'Line':  # Check if it's a line annotation (type 2)
                info = annot.info  # Extract basic info about the annotation
                points = annot.vertices  # Line annotation points
                annot_xref_obj = doc.xref_object(annot.xref)
                annot_name = str(doc.xref_get_key(annot.xref,"NM")[1])
                annot_xref_str = str(annot.xref)
                # Expect two tuples (x1, y1) and (x2, y2)
                if len(points) == 2:  # Corrected to check for two tuples
                    x1, y1 = points[0]
                    x2, y2 = points[1]
                    # Convert points to inches assuming 72 DPI (native PDF units are in points)
                    x1_inches = x1 / 72
                    y1_inches = y1 / 72
                    x2_inches = x2 / 72
                    y2_inches = y2 / 72
                    length = calculate_distance(x1_inches,y1_inches,x2_inches,y2_inches)
                    subject = info.get('subject', 'Unknown')
                    BsiData=doc.xref_get_key(annot.xref, "BSIColumnData")
                    BsiString=BsiData[1]
                    pattern = r'\((.*?)\)'
                    Bsi_values = re.findall(pattern, BsiString)
                    # Create a dictionary of keys and values
                    xref_dict = dict(zip(xref_values, Bsi_values))
                    scalefactor=96
                    scaledlength=length*scalefactor
                    page_label = page.get_label()

                    annotations.append({
                        'page_number': page_number,
                        'Sheet Number': page_label,
                        'subject': subject.upper(),
                        'Scaled Length':scaledlength,
                        **xref_dict,
                        'x1': x1_inches,
                        'y1': y1_inches,
                        'x2': x2_inches,
                        'y2': y2_inches,
                        'NM': annot_name,
                        'ScaleFactor':96,
                        'author': info.get('title', 'Unknown'),
                        'modified_date': info.get('modDate', 'Unknown'),
                        'xref':annot_xref_str
                        
                    })

    doc.close()
    return annotations

def save_to_csv(data, csv_path):
    df = pd.DataFrame(data)
    file_path = r"I:\Shared drives\United Structural\Reference\aisc-shapes-database-v15.0.xlsx"
    Wide_Flange_NoExt = os.path.splitext(csv_path)[0]
    Wide_Flange_Path= Wide_Flange_NoExt + "Wide Flange.csv"

    # Read the additional data spreadsheet into another pandas dataframe
    additional_df = pd.read_excel(file_path, sheet_name='Database v15.0', usecols=['AISC_Manual_Label', 'W','Type','PB'])

    # Merge the two dataframes based on the 'subject' column in df and 'AISC_Manual_Label' column in additional_df
    merged_df = pd.merge(df, additional_df, left_on='subject', right_on='AISC_Manual_Label', how='left')

    filtered_df = merged_df[merged_df['Type'] == 'W'].copy()
    filtered_df['Weight Each'] = filtered_df['W'] * filtered_df['Scaled Length'] /12  # Example calculation
    filtered_df.to_csv(Wide_Flange_Path,index=False)



    merged_df.to_csv(csv_path, index=False)
    print("Annotations have been successfully saved to", csv_path)


def main():
    #pdf_path = select_pdf_file()
    pdf_path = r"I:\Shared drives\United Structural\Projects\2024\24-027 - Stellant - v2022\Drawings\Other\Grids from 2024-03-14 - Stellant - Struct CDs.pdf"
    if not pdf_path:
        print("No PDF file selected. Exiting...")
        return

    annotations = extract_annotations(pdf_path)
    if not annotations:
        print("No line annotations found in the PDF. Exiting...")
        return

    csv_path = select_csv_file()
    if not csv_path:
        print("No CSV file selected. Exiting...")
        return

    save_to_csv(annotations, csv_path)

if __name__ == "__main__":
    main()
