# Proof of participation Generator

This project is a Python script that generates PDF certificates from an Excel file containing participant names. It reads the participant names from the Excel file, inserts them into a PDF template, and saves the modified PDF files with the participant names as filenames.

## Requirements

- Python 3.11

## Installation

1. Clone the repository or download the project files.
2. Create a virtual environment and activate it:

```bash
python -m venv .venv
source .venv/bin/activate  # On Windows, use `.venv\Scripts\activate`
```

3. Install the required Python packages by running the following command:

```bash
pip install -r requirements.txt
```

## Usage

1. Create a `.env` file in the project root directory and populate it with the following environment variables, For Example:

```shell
# 參與者名單檔案路徑 (.xlsx)
PARTICIPANTS_FILE="test回饋表單NTNUxAWS.xlsx"

# 參與者名單的工作表
PARTICIPANTS_SHEET="sheet1"

# PDF模板檔案路徑
TEMPLATE_PDF_PATH="[template] AWS Educate certificate_FCU.pdf"

# 英文字體檔案路徑，這裡預設是微軟正黑體
FONT_FILE_PATH="fonts/msjh.ttf"

# 參與者名字是工作表的哪個 Column，請打上那個 Column 內的文字
PARTICIPANT_NAME_COLUMN="您的姓名"
```

- Run the script:

```bash
python main.py
```

The script will generate PDF certificates for each participant and save them in the `results` directory within the project folder. The filenames will be in the format `certificate_<number>_<participant_name>.pdf`.

## Note

The script assumes that the PDF template has only one page. If your PDF template has multiple pages, you may need to modify the script accordingly.
