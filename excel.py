from PIL import Image
import pytesseract
import pandas as pd

# Set the full path to the Tesseract executable
pytesseract.pytesseract.tesseract_cmd = r'C:/Users/alimu/OneDrive/Desktop/New folder (3)/tesseract-main/tesseract.exe'


 

def image_to_text(image_path):
    # Open the image file
    image = Image.open(image_path)

    # Use pytesseract to extract text from the image
    text = pytesseract.image_to_string(image)

    return text

def text_to_excel(text, output_excel_path):
    # Split the text into lines
    lines = [line.split() for line in text.splitlines()]

    # Convert the data to a DataFrame using pandas
    df = pd.DataFrame(lines)

    # Save the DataFrame to an Excel file
    df.to_excel(output_excel_path, index=False, header=False)

if __name__ == "__main__":
    image_path = 'sheet.jpg'

    output_excel_path = 'output_excel_file.xlsx'

    # Perform OCR on the image
    extracted_text = image_to_text(image_path)

    # Convert the extracted text to an Excel file
    text_to_excel(extracted_text, output_excel_path)

    print(f"Conversion complete. Excel file saved at: {output_excel_path}")
