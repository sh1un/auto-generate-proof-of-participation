import os

import fitz  # PyMuPDF
import pandas as pd

# Ensure the results directory exists
results_dir = "results"
if not os.path.exists(results_dir):
    os.makedirs(results_dir)

# Read the participants list
participants_df = pd.read_excel("test回饋表單NTNUxAWS.xlsx", sheet_name="sheet1")

# PDF template path
template_pdf_path = "[template] AWS Educate certificate_FCU.pdf"

# Iterate through each participant
for index, row in participants_df.iterrows():
    # Open the PDF template
    doc = fitz.open(template_pdf_path)
    # Assuming the template has only one page
    page = doc[0]
    # Define text style
    text = row["您的姓名"]  # Name read from Excel
    font_size = 36
    font_color = (0, 0, 0)  # Black
    x_coord = 440  # Horizontal position, needs adjustment
    y_coord = 250  # Vertical position, needs adjustment

    # Insert text
    # Note: This assumes you have a font file that supports Chinese characters
    # If your PDF template requires a specific font to support Chinese, you should specify the `fontfile` parameter
    page.insert_text(
        (x_coord, y_coord),
        text,
        fontsize=font_size,
        fontfile="C:\\Windows\\Fonts\\msjh.ttf",
        fontname="china-ts",
        color=font_color,
        encoding=0,
    )

    # Define the output PDF file path
    output_pdf_path = os.path.join(
        results_dir, f'certificate_{index+1}_{row["您的姓名"]}.pdf'
    )

    # Save the modified PDF
    doc.save(output_pdf_path)
    # Close the document
    doc.close()

print("All certificate files have been generated.")
