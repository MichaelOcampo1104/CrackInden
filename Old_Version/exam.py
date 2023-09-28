from openpyxl import Workbook
from openpyxl.drawing.image import Image as OpenpyxlImage
import os

def main():
    # Create a new Excel workbook and select the active worksheet
    wb = Workbook()
    ws = wb.active

    # Add some dummy data
    ws.append(['Page Number', 'Full Text', 'Image Reference'])
    ws.append([1, 'Some text here', 'Image_1'])

    # Define the image path (replace with the path to an actual image on your system)
    img_path = r'C:\Users\hp\Documents\Python Scripts\crackinden\file_filtered\Page_7_Image_1.png'

    # Check if the image exists
    if os.path.exists(img_path):
        # Add the image to Excel
        img = OpenpyxlImage(img_path)
        img.width = 60
        img.height = 80
        cell_num = 'C2'  # The cell where the image will appear
        ws.column_dimensions['C'].width = 15
        ws.row_dimensions[2].height = 90
        ws.add_image(img, cell_num)
        print(f"Image added at {cell_num}")
    else:
        print(f"Image not found at {img_path}")

    # Save the workbook
    wb.save('test_image_embedding.xlsx')

if __name__ == "__main__":
    main()
