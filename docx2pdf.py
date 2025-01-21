import os
import sys
import argparse
import win32com.client

def prompt_output_directory():
    """
    Prompts the user to decide where to output the PDF files.

    Returns:
        str: Either 'same' to output PDFs in the same directory as .docx files, or 'pdf_output' for a central folder.
    """
    while True:
        response = input("Output PDFs in the same directory as .docx files? (y/n): ").strip().lower()
        if response in ['y', 'yes']:
            return "same"
        elif response in ['n', 'no']:
            return "pdf_output"

def prompt_overwrite(file_path, overwrite_all):
    """
    Prompts the user whether to overwrite an existing file or use a global decision.

    Args:
        file_path (str): Path to the file to potentially overwrite.
        overwrite_all (str): Global overwrite decision ('yes', 'no', or None).

    Returns:
        bool: True if the file should be overwritten, False otherwise.
    """
    if overwrite_all == "yes":
        return True
    elif overwrite_all == "no":
        return False

    while True:
        response = input(f"File '{file_path}' already exists. Overwrite? (y/n/all): ").strip().lower()
        if response in ['y', 'yes']:
            return True
        elif response in ['n', 'no']:
            return False
        elif response == 'all':
            return "all"

def convert_docx_to_pdf(docx_file, pdf_file):
    """
    Converts a single .docx file to a PDF using Microsoft Word.

    Args:
        docx_file (str): Path to the input .docx file.
        pdf_file (str): Path to the output PDF file.
    """
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(docx_file)
        doc.SaveAs(pdf_file, FileFormat=17)  # 17 corresponds to wdFormatPDF
        doc.Close()
        word.Quit()
        print(f"SUCCESS: Converted '{os.path.basename(docx_file)}' to PDF.")
    except Exception as e:
        print(f"ERROR: Failed to convert '{os.path.basename(docx_file)}' to PDF. Reason: {e}")

def convert_all_docx_to_pdf(directory):
    """
    Converts all .docx files in the given directory and its subfolders to PDFs.

    Args:
        directory (str): Path to the directory containing .docx files.
    """
    if not os.path.isdir(directory):
        print(f"ERROR: Directory '{directory}' does not exist.")
        sys.exit(1)

    output_choice = prompt_output_directory()
    output_directory = os.path.join(directory, "pdf_output") if output_choice == "pdf_output" else None
    if output_directory:
        os.makedirs(output_directory, exist_ok=True)

    print("Starting conversion of .docx files to PDFs...")

    overwrite_all = None

    for root, _, files in os.walk(directory):
        for file in files:
            if file.endswith('.docx'):
                docx_path = os.path.join(root, file)
                if output_choice == "same":
                    pdf_path = os.path.splitext(docx_path)[0] + ".pdf"
                else:
                    relative_path = os.path.relpath(root, directory)
                    pdf_subdir = os.path.join(output_directory, relative_path)
                    os.makedirs(pdf_subdir, exist_ok=True)
                    pdf_filename = os.path.splitext(file)[0] + ".pdf"
                    pdf_path = os.path.join(pdf_subdir, pdf_filename)

                if os.path.exists(pdf_path):
                    if overwrite_all not in ["yes", "no"]:
                        decision = prompt_overwrite(pdf_path, overwrite_all)
                        if decision == "all":
                            overwrite_all = "yes"
                        elif decision == "no":
                            overwrite_all = "no"
                        elif not decision:
                            print(f"SKIPPED: '{file}' not overwritten.")
                            continue
                    elif overwrite_all == "no":
                        print(f"SKIPPED: '{file}' not overwritten.")
                        continue

                print(f"Processing: {file}")
                convert_docx_to_pdf(docx_path, pdf_path)

    print(f"Conversion complete! All PDFs saved in: {output_directory if output_directory else 'original directories'}")

def main():
    parser = argparse.ArgumentParser(description="Convert all .docx files in a directory to PDFs.")
    parser.add_argument("directory", type=str, help="Path to the directory containing .docx files.")

    args = parser.parse_args()
    convert_all_docx_to_pdf(args.directory)

if __name__ == "__main__":
    main()