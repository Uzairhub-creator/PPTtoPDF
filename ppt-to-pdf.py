import os
import comtypes.client

def ppt_to_pdf(ppt_path, pdf_filename=None):
    # Initialize PowerPoint application
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1

    # Open the PowerPoint file
    presentation = powerpoint.Presentations.Open(ppt_path)

    # Set the PDF filename if not provided
    if pdf_filename is None:
        pdf_filename = os.path.splitext(ppt_path)[0] + '.pdf'

    # Export the presentation as a PDF
    presentation.SaveAs(pdf_filename, 32)  # 32 is the code for PDF format
    presentation.Close()

    print(f"Converted '{ppt_path}' to '{pdf_filename}'")

    # Quit the PowerPoint application
    powerpoint.Quit()


def convert_all_pptx_in_directory(directory_path):
    # Check if the provided directory exists
    if not os.path.isdir(directory_path):
        print("The specified path is not a directory or does not exist.")
        return

    # Find all .pptx files in the directory
    pptx_files = [f for f in os.listdir(directory_path) if f.endswith('.pptx')]

    if not pptx_files:
        print("No .pptx files found in the specified directory.")
        return

    # Convert each .pptx file to PDF
    for pptx_file in pptx_files:
        ppt_path = os.path.join(directory_path, pptx_file)
        try:
            ppt_to_pdf(ppt_path)
        except Exception as e:
            print(f"Failed to convert {pptx_file}: {e}")

# Prompt the user for the directory path
directory_path = input("Enter the path to the directory containing the .pptx files: ")
convert_all_pptx_in_directory(directory_path)
