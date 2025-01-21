# docx2pdf
Python script to batch convert .docx files to PDFs with flexible output options and overwrite prompts.

# Description
This Python script converts all .docx files in a specified directory and its subdirectories into PDFs. It provides options for handling output locations and file overwrites, while preserving basic formatting and content from the .docx files.

## Features
- Converts .docx files to PDFs.
- Option to output PDFs in the same directory as the .docx files or in a centralized pdf_output folder.
- Prompts the user whether to overwrite existing PDF files.
- Supports recursive directory traversal to process .docx files in subdirectories.

## Requirements
Python 3.6+

## Required Python libraries:
```bash
python-docx
pywin32 (for access to win32com.client - native Windows PDF exporter)
```
Install the required libraries using pip:
```bash
pip install python-docx pywin32
```

## Usage
Running the Script
To use the script, save it as docx_to_pdf.py and run it from the command line with the directory containing .docx files as an argument:
```bash
python docx2pdf.py <directory_path>
```
Replace <directory_path> with the path to the directory containing the .docx files you want to convert.

## Script Workflow
1. The script prompts whether to save PDFs in the same directory as the .docx files or in a pdf_output folder.
2. If a PDF file with the same name already exists, the script asks whether to overwrite it, skip it, or apply the same decision to all subsequent files.
3. The script processes all .docx files in the specified directory and its subdirectories.

## Example
```bash
python docx2pdf.py "C:\Users\User\Documents\DocxFiles"
```

## Error Handling
1. If a directory does not exist, the script will terminate with an error message.
2. If a .docx file fails to convert, an error message will indicate the problematic file and the reason.

## License
This project is licensed under the MIT License.
