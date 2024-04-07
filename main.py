import os

import fitz  # PyMuPDF
import pandas as pd
from dotenv import load_dotenv

# Load .env file
load_dotenv()

# Ensure the results directory exists
results_dir = "results"
if not os.path.exists(results_dir):
    os.makedirs(results_dir)

# Read the participants list
participants_file = os.environ.get("PARTICIPANTS_FILE")
participants_sheet = os.environ.get("PARTICIPANTS_SHEET")
participant_name_column = os.environ.get("PARTICIPANT_NAME_COLUMN")
participants_df = pd.read_excel(participants_file, sheet_name=participants_sheet)

# PDF template path
template_pdf_path = os.environ.get("TEMPLATE_PDF_PATH")

# Font file path
font_file_path = os.path.join(
    os.path.dirname(__file__), os.environ.get("FONT_FILE_PATH")
)

# Iterate through each participant
for index, row in participants_df.iterrows():
    # Open the PDF template
    doc = fitz.open(template_pdf_path)
    # Assuming the template has only one page
    page = doc[0]
    # Define text style
    text = row[participant_name_column]  # Name read from Excel
    font_size = 32
    font_color = (0, 0, 0)  # Black

    # Define the position and size of the text box
    x_coord = 350  # Horizontal position
    y_coord = 210  # Vertical position
    rect = fitz.Rect(x_coord, y_coord, x_coord + 300, y_coord + 300)

    # Choose the font based on the text content
    if all(ord(char) < 128 for char in text):  # If the text is all English characters
        page.insert_textbox(
            rect,
            text,
            fontsize=font_size,
            fontfile=font_file_path,
            color=font_color,
            align=fitz.TEXT_ALIGN_CENTER,  # Set text alignment to center
            overlay=True,
        )
    else:  # If the text contains Chinese characters
        page.insert_textbox(
            rect,
            text,
            fontsize=font_size,
            fontname="china-t",
            color=font_color,
            align=fitz.TEXT_ALIGN_CENTER,  # Set text alignment to center
            overlay=True,
        )

    # Define the output PDF file path
    output_pdf_path = os.path.join(
        results_dir, f"certificate_{index+1}_{row[participant_name_column]}.pdf"
    )

    # Save the modified PDF
    doc.save(output_pdf_path)

    # Close the document
    doc.close()

print("All certificate files have been generated.")
