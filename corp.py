import os
import uuid
from PIL import Image
import pyheif  # 新增导入
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.ttk import Progressbar
from openpyxl.drawing.image import Image as ExcelImage  # 新增导入
import zipfile
import io
import threading  # For running tasks in the background


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


def crop_images(input_folder, output_folder, top_crop, bottom_crop):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    image_files = [f for f in os.listdir(input_folder) if f.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.gif', '.heic')) and not f.startswith('._')]
    total_files = len(image_files)

    progress_bar['maximum'] = total_files

    for index, filename in enumerate(image_files):
        img_path = os.path.join(input_folder, filename)
        try:
            if filename.lower().endswith('.heic'):
                img_path = convert_heic_to_jpg(img_path, output_folder)
                filename = os.path.basename(img_path)

            with Image.open(img_path) as img:
                width, height = img.size
                crop_box = (0, top_crop, width, height - bottom_crop)
                cropped_img = img.crop(crop_box)
                unique_filename = str(uuid.uuid4()) + os.path.splitext(filename)[1]
                output_path = os.path.join(output_folder, unique_filename)
                cropped_img.save(output_path)
        except Exception as e:
            print(f"无法处理文件 {filename}: {e}")

        # Update progress bar and label
        progress_bar['value'] = index + 1
        progress_label['text'] = f"进度: {index + 1}/{total_files}"
        root.update_idletasks()


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
        except Exception as e:
            print(f"无法处理文件 {excel_file}: {e}")


def extract_images():
    input_folder = input_folder_var.get()
    if not input_folder:
        messagebox.showwarning("警告", "请确保选择输入文件夹。")
        return

    extract_images_from_excel(input_folder)
    messagebox.showinfo("完成", "图像提取完成，存放在 tmp_storage 文件夹。")


def select_input_folder():
    folder_selected = filedialog.askdirectory()
    if folder_selected:
        input_folder_var.set(folder_selected)


def select_output_folder():
    folder = filedialog.askdirectory()
    output_folder_var.set(folder)


def start_cropping():
    input_folder = input_folder_var.get()
    output_folder = output_folder_var.get()
    top_crop = int(top_crop_var.get())
    bottom_crop = int(bottom_crop_var.get())

    if not input_folder or not output_folder:
        messagebox.showwarning("警告", "请确保选择输入和输出文件夹。")
        return

    threading.Thread(target=lambda: crop_images(input_folder, output_folder, top_crop, bottom_crop)).start()  # Run cropping in a separate thread


def start_processing():
    input_folder = input_folder_var.get()
    output_folder = output_folder_var.get()
    operation = operation_var.get()

    if not input_folder or not output_folder:
        messagebox.showwarning("警告", "请确保选择输入和输出文件夹。")
        return

    if operation in ["提取图像", "均需执行"]:
        threading.Thread(target=lambda: extract_images_from_excel(input_folder)).start()
    if operation in ["进行剪裁", "均需执行"]:
        threading.Thread(target=lambda: crop_images(input_folder, output_folder, int(top_crop_var.get()), int(bottom_crop_var.get()))).start()
    if operation == "图像格式转换":
        threading.Thread(target=lambda: convert_image_format(input_folder, output_folder)).start()


# Create main window
root = tk.Tk()
root.title("图像处理工具")

# Input folder
tk.Label(root, text="选择文件或路径:").grid(row=0, column=0, padx=10, pady=10)
input_folder_var = tk.StringVar()
input_folder_entry = tk.Entry(root, textvariable=input_folder_var, width=50)
input_folder_entry.grid(row=0, column=1, padx=10, pady=10)
tk.Button(root, text="选择文件夹", command=select_input_folder).grid(row=0, column=2, padx=10, pady=10)

# Operation selection
tk.Label(root, text="选择操作:").grid(row=1, column=0, padx=10, pady=10)
operation_var = tk.StringVar(value="提取图像")
operation_options = ["提取图像", "进行剪裁", "图像格式转换", "均需执行"]
operation_menu = tk.OptionMenu(root, operation_var, *operation_options)
operation_menu.grid(row=1, column=1, padx=10, pady=10)

# Output folder
tk.Label(root, text="存放路径:").grid(row=2, column=0, padx=10, pady=10)
output_folder_var = tk.StringVar()
output_folder_entry = tk.Entry(root, textvariable=output_folder_var, width=50)
output_folder_entry.grid(row=2, column=1, padx=10, pady=10)
tk.Button(root, text="选择存放路径", command=select_output_folder).grid(row=2, column=2, padx=10, pady=10)

# Crop settings
tk.Label(root, text="上裁剪像素:").grid(row=3, column=0, padx=10, pady=10)
top_crop_var = tk.StringVar(value="0")
top_crop_entry = tk.Entry(root, textvariable=top_crop_var)
top_crop_entry.grid(row=3, column=1, padx=10, pady=10)

tk.Label(root, text="下裁剪像素:").grid(row=4, column=0, padx=10, pady=10)
bottom_crop_var = tk.StringVar(value="0")
bottom_crop_entry = tk.Entry(root, textvariable=bottom_crop_var)
bottom_crop_entry.grid(row=4, column=1, padx=10, pady=10)

# Progress bar
progress_bar = Progressbar(root, length=300)
progress_bar.grid(row=5, column=1, padx=10, pady=10)
progress_label = tk.Label(root, text="进度: 0/0")
progress_label.grid(row=6, column=1, padx=10, pady=10)

# Start processing button
tk.Button(root, text="执行操作", command=start_processing).grid(row=7, column=1, padx=10, pady=10)

root.mainloop()
