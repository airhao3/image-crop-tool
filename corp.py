import os
import uuid
from PIL import Image
import pyheif  # 新增导入
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.ttk import Progressbar

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

    # 获取所有图像文件
    image_files = [f for f in os.listdir(input_folder) if f.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.gif', '.heic')) and not f.startswith('._')]
    total_files = len(image_files)

    # 更新进度条
    progress_bar['maximum'] = total_files

    for index, filename in enumerate(image_files):
        img_path = os.path.join(input_folder, filename)
        try:
            if filename.lower().endswith('.heic'):
                img_path = convert_heic_to_jpg(img_path, output_folder)  # 转换 HEIC 文件
                filename = os.path.basename(img_path)  # 更新文件名为转换后的文件名

            with Image.open(img_path) as img:
                width, height = img.size
                crop_box = (0, top_crop, width, height - bottom_crop)
                cropped_img = img.crop(crop_box)
                unique_filename = str(uuid.uuid4()) + os.path.splitext(filename)[1]
                output_path = os.path.join(output_folder, unique_filename)
                cropped_img.save(output_path)
        except Exception as e:
            print(f"无法处理文件 {filename}: {e}")

        # 更新进度条
        progress_bar['value'] = index + 1
        progress_label['text'] = f"进度: {index + 1}/{total_files}"  # 更新进度文本
        root.update_idletasks()  # 更新界面

    messagebox.showinfo("完成", "图像裁剪完成。")

def select_input_folder():
    folder = filedialog.askdirectory()
    input_folder_var.set(folder)

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

    crop_images(input_folder, output_folder, top_crop, bottom_crop)

# 创建主窗口
root = tk.Tk()
root.title("图像裁剪工具")

# 输入文件夹选择
input_folder_var = tk.StringVar()
tk.Label(root, text="输入文件夹:").grid(row=0, column=0)
tk.Entry(root, textvariable=input_folder_var, width=50).grid(row=0, column=1)
tk.Button(root, text="选择", command=select_input_folder).grid(row=0, column=2)

# 输出文件夹选择
output_folder_var = tk.StringVar()
tk.Label(root, text="输出文件夹:").grid(row=1, column=0)
tk.Entry(root, textvariable=output_folder_var, width=50).grid(row=1, column=1)
tk.Button(root, text="选择", command=select_output_folder).grid(row=1, column=2)

# 裁剪参数输入
top_crop_var = tk.StringVar(value="50")
tk.Label(root, text="上边界裁剪像素数:").grid(row=2, column=0)
tk.Entry(root, textvariable=top_crop_var, width=10).grid(row=2, column=1)

bottom_crop_var = tk.StringVar(value="50")
tk.Label(root, text="下边界裁剪像素数:").grid(row=3, column=0)
tk.Entry(root, textvariable=bottom_crop_var, width=10).grid(row=3, column=1)

# 进度条
progress_bar = Progressbar(root, orient='horizontal', length=400, mode='determinate')
progress_bar.grid(row=5, column=0, columnspan=3, pady=10)

# 进度标签
progress_label = tk.Label(root, text="进度: 0/0")
progress_label.grid(row=6, column=0, columnspan=3)

# 开始裁剪按钮
tk.Button(root, text="开始裁剪", command=start_cropping).grid(row=4, column=1)

# 运行主循环
root.mainloop()