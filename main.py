import os

import fitz  # PyMuPDF
import pandas as pd
from dotenv import load_dotenv

# Load .env file
load_dotenv(verbose=True, override=True)

# Ensure the results directory exists
print("現在工作目錄:" + os.getcwd())
results_dir = "results"
if not os.path.exists(results_dir):
    os.makedirs(results_dir)

# Read the participants list
PARTICIPANTS_FILE = os.environ.get("PARTICIPANTS_FILE")
print("名單: " + PARTICIPANTS_FILE)
PARTICIPANTS_SHEET = "清單"
PARTICIPANT_NAME_COLUMN = "姓名"
EVENT_NAME_COLUMN = "event_name"
EVENT_DATE_COLUMN = "event_date"
EVENT_LOCATION_COLUMN = "event_location"
participants_df = pd.read_excel(PARTICIPANTS_FILE, sheet_name=PARTICIPANTS_SHEET)

# PDF template path
TEMPLATE_PDF_PATH = "[template] AWS Educate certificate_FCU-Amazon Ember.pdf"

# Font file path
font_file_path_english = os.path.join(
    os.path.dirname(__file__), "fonts/Ember/AmazonEmber_Rg.ttf"
)
font_file_path_chinese = os.path.join(
    os.path.dirname(__file__), "fonts/static/NotoSansTC-Regular.ttf"
)

# Iterate through each participant
for index, row in participants_df.iterrows():
    # Open the PDF template
    doc = fitz.open(TEMPLATE_PDF_PATH)
    # Assuming the template has only one page
    page = doc[0]

    # Define text style
    name_text = row[PARTICIPANT_NAME_COLUMN]  # Name read from Excel
    event_text = f"in recognition of your participation in the {row[EVENT_NAME_COLUMN]} on {row[EVENT_DATE_COLUMN]} at {row[EVENT_LOCATION_COLUMN]}."
    font_size_name = 32
    font_size_event = 18
    font_color = (0, 0, 0)  # Black

    # Define the position and size of the text box for the name
    x_coord_name = 365  # Horizontal position
    y_coord_name = 210  # Vertical position
    rect_name = fitz.Rect(
        x_coord_name, y_coord_name, x_coord_name + 300, y_coord_name + 50
    )

    # Define the position and size of the text box for the event information
    x_coord_event = 250  # Horizontal position
    y_coord_event = 265  # Vertical position
    rect_event = fitz.Rect(
        x_coord_event, y_coord_event, x_coord_event + 525, y_coord_event + 50
    )

    # Use TextWriter to insert text
    tw = fitz.TextWriter(page.rect)

    # Insert the name text
    if all(
        ord(char) < 128 for char in name_text
    ):  # If the text is all English characters
        font = fitz.Font(fontfile=font_file_path_english)
    else:  # If the text contains Chinese characters
        font = fitz.Font(fontfile=font_file_path_chinese)

    tw.fill_textbox(rect_name, name_text, font=font, fontsize=font_size_name, align=1)

    # Insert the event information text
    font = fitz.Font(fontfile=font_file_path_english)
    tw.fill_textbox(
        rect_event, event_text, font=font, fontsize=font_size_event, align=1
    )

    # Write text to the page
    tw.write_text(page)

    # Define the output PDF file path
    output_pdf_path = os.path.join(
        results_dir, f"certificate_{index+1}_{row[PARTICIPANT_NAME_COLUMN]}.pdf"
    )

    # Save the modified PDF
    doc.save(output_pdf_path)

    # Close the document
    doc.close()

print("All certificate files have been generated.")
