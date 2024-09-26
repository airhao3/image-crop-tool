import os
import uuid
import argparse
from PIL import Image
import pyheif
import zipfile
import io

# Convert HEIC to JPG
def convert_heic_to_jpg(heic_file, output_folder):
    heif_file = pyheif.read(heic_file)
    img = Image.frombytes(
        heif_file.mode,
        heif_file.size,
        heif_file.data,
        "raw",
        heif_file.mode,
        heif_file.stride,
    )
    unique_filename = str(uuid.uuid4()) + '.jpg'
    output_path = os.path.join(output_folder, unique_filename)
    img.save(output_path, "JPEG")
    return output_path

# Crop images based on top and bottom crop values
def crop_images(input_folder, output_folder, top_crop, bottom_crop):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    image_files = [f for f in os.listdir(input_folder) if f.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.gif', '.heic')) and not f.startswith('._')]
    total_files = len(image_files)

    for index, filename in enumerate(image_files):
        img_path = os.path.join(input_folder, filename)
        try:
            if filename.lower().endswith('.heic'):
                img_path = convert_heic_to_jpg(img_path, output_folder)  # Convert HEIC to JPG

            with Image.open(img_path) as img:
                width, height = img.size
                crop_box = (0, top_crop, width, height - bottom_crop)
                cropped_img = img.crop(crop_box)
                unique_filename = str(uuid.uuid4()) + os.path.splitext(filename)[1]
                output_path = os.path.join(output_folder, unique_filename)
                cropped_img.save(output_path)
            print(f"Processed: {filename}")
        except Exception as e:
            print(f"Failed to process file {filename}: {e}")

# Extract images from Excel files
def extract_images_from_excel(input_folder):
    temp_folder = os.path.join(input_folder, "tmp_storage")
    if not os.path.exists(temp_folder):
        os.makedirs(temp_folder)

    excel_files = [f for f in os.listdir(input_folder) if f.lower().endswith('.xlsx')]

    for excel_file in excel_files:
        zip_file_path = os.path.join(input_folder, excel_file)
        try:
            with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
                for file_info in zip_ref.infolist():
                    if file_info.filename.startswith('xl/media/'):
                        with zip_ref.open(file_info) as img_file:
                            img_data = img_file.read()
                            img = Image.open(io.BytesIO(img_data))
                            img_path = os.path.join(temp_folder, f"{excel_file}_{file_info.filename.split('/')[-1]}")
                            img.save(img_path)
            print(f"Extracted images from: {excel_file}")
        except Exception as e:
            print(f"Failed to process Excel file {excel_file}: {e}")

# Main function to handle command-line arguments
def main():
    parser = argparse.ArgumentParser(description="Image Processing Tool")
    
    # Add options for different functionalities
    parser.add_argument('--input_folder', type=str, required=True, help="Input folder containing images or Excel files.")
    parser.add_argument('--output_folder', type=str, required=True, help="Output folder for processed images.")
    parser.add_argument('--top_crop', type=int, default=0, help="Pixels to crop from the top of the image.")
    parser.add_argument('--bottom_crop', type=int, default=0, help="Pixels to crop from the bottom of the image.")
    parser.add_argument('--operation', choices=['crop', 'extract_excel', 'all'], default='all', help="Operation to perform (crop images, extract from Excel, or both).")

    args = parser.parse_args()

    # Perform the selected operation
    if args.operation == 'crop' or args.operation == 'all':
        crop_images(args.input_folder, args.output_folder, args.top_crop, args.bottom_crop)
    if args.operation == 'extract_excel' or args.operation == 'all':
        extract_images_from_excel(args.input_folder)

if __name__ == '__main__':
    main()
