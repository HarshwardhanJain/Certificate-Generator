import pandas as pd
from PyPDF2 import PdfWriter, PdfReader
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import io
import os

def create_certificates(participants, number, output_dir):
    for participant in participants:
        name = participant
        
        # Create a PDF canvas in memory
        packet = io.BytesIO()
        can = canvas.Canvas(packet, pagesize=letter)
        
        # Set font style, size, and make it bold for the name
        font = "Helvetica-Bold"
        font_size_name = 24
        can.setFont(font, font_size_name)
        
        # Calculate the width of the name string
        name_width = can.stringWidth(name, font, font_size_name)
        
        # Set coordinates for the participant's name
        page_width, page_height = letter
        x_name = (page_width - name_width + 233) / 2  # Center the name horizontally
        y_name = 226  # Y-coordinate for the name
        
        # Draw the participant's name at the specified coordinates
        can.drawString(x_name, y_name, name)
        
        # Set font style, size, and make it bold for the number
        font_size_number = 15
        can.setFont(font, font_size_number)
        
        # Set coordinates for the participant's number
        x_number = 530  # X-coordinate for the number
        y_number = 190  # Y-coordinate for the number
        
        # Draw the participant's number at the specified coordinates
        can.drawString(x_number, y_number, number)
        
        can.save()

        # Move the pointer to the beginning of the StringIO buffer
        packet.seek(0)

        # Read the new PDF with the name and number
        new_pdf = PdfReader(packet)
        
        # Ensure the template path is correct
        template_path = "Quality Week Celebration\(Coordinator) Poster Presentation - World Quality Week Celebration.pdf"
        existing_pdf = PdfReader(open(template_path, "rb"))
        
        # Create a writer object
        output = PdfWriter()

        # Merge the name and number PDF with the existing template
        for i in range(len(existing_pdf.pages)):
            page = existing_pdf.pages[i]
            if i == 0:  # Assuming you want to modify the first page
                page.merge_page(new_pdf.pages[0])
            output.add_page(page)

        # Save the final PDF to a file
        os.makedirs(output_dir, exist_ok=True)
        output_path = os.path.join(output_dir, f"{name}.pdf")
        with open(output_path, "wb") as outputStream:
            output.write(outputStream)

        # Print the absolute path of the saved PDF
        absolute_path = os.path.abspath(output_path)
        print(f"Certificate saved as {absolute_path}")

def read_excel_and_generate_certificates(file_path, base_output_dir):
    df = pd.read_excel(file_path)
    print(f"Columns in {file_path}: {df.columns.tolist()}")  # Print column names
    
    while True:
        folder_name = input(f"Enter the folder name to create inside the defined path for {file_path}: ")
        output_dir = os.path.join(base_output_dir, folder_name)
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
            break
        else:
            print("Folder already exists. Please try again.")
    
    name_column = input("Enter the column name for the participant's name: ")
    roll_no_column = input("Enter the column name for the participant's roll number: ")
    number = input(f"Enter the number for {file_path}: ")
    
    participants = df.apply(lambda row: f"{row[name_column]} {row[roll_no_column]}", axis=1).tolist()
    create_certificates(participants, number, output_dir)

# Example usage
excel_files = [r"Quality Week Celebration\Test.xlsx"]

base_output_dir = r"Quality Week Celebration\Processed Certificates"

for file in excel_files:
    read_excel_and_generate_certificates(file, base_output_dir)