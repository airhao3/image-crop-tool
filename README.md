# 图像处理工具

## 简介
该工具用于处理图像文件，包括图像格式转换、裁剪和从Excel文件中提取图像。

## 功能
- **选择文件夹**: 选择输入和输出文件夹。
- **图像格式转换**: 支持将HEIC格式转换为JPEG。
- **裁剪图像**: 根据用户输入的上裁剪和下裁剪像素裁剪图像。
- **从Excel提取图像**: 从指定的Excel文件中提取嵌入的图像。

## 使用说明
1. 运行 `corp.py` 启动图形界面。
2. 选择输入文件夹和输出文件夹。
3. 选择所需的操作（提取图像、进行剪裁、图像格式转换）。
4. 输入裁剪像素（如果选择裁剪）。
5. 点击“执行操作”按钮开始处理。

## 命令行使用
`corp_cli.py` 提供了命令行界面，允许用户通过命令行参数进行图像处理。

#### 命令格式
```bash
python corp_cli.py --input_folder <输入文件夹> --output_folder <输出文件夹> [--top_crop <上裁剪像素>] [--bottom_crop <下裁剪像素>] [--operation <操作>]
```

#### 参数说明
- `--input_folder`: 输入文件夹，包含待处理的图像或Excel文件（必需）。
- `--output_folder`: 输出文件夹，处理后的图像将保存到此文件夹（必需）。
- `--top_crop`: 从图像顶部裁剪的像素数（可选，默认为0）。
- `--bottom_crop`: 从图像底部裁剪的像素数（可选，默认为0）。
- `--operation`: 要执行的操作，选项包括 `crop`（裁剪图像）、`extract_excel`（从Excel提取图像）或 `all`（同时执行裁剪和提取，默认为 `all`）。

#### 示例
1. 裁剪图像：
   ```bash
   python corp_cli.py --input_folder ./input_images --output_folder ./output_images --top_crop 10 --bottom_crop 10 --operation crop
   ```

2. 从Excel提取图像：
   ```bash
   python corp_cli.py --input_folder ./input_excel --output_folder ./output_images --operation extract_excel
   ```

3. 同时裁剪图像和提取Excel中的图像：
   ```bash
   python corp_cli.py --input_folder ./input_images --output_folder ./output_images --top_crop 10 --bottom_crop 10 --operation all
   ```

## 依赖
- Python 3.x
- Pillow
- pyheif
- openpyxl
- tkinter

## 注意事项
- 确保输入文件夹中包含支持的图像格式（.png, .jpg, .jpeg, .bmp, .gif, .heic）。
- 输出文件夹必须存在，程序会在该文件夹中保存处理后的图像。

## 错误处理
- 如果出现错误，程序会在控制台输出错误信息，确保输入的文件夹和文件格式正确。

## 更新日志
- 新增HEIC格式支持。
- 优化图像裁剪和提取功能。