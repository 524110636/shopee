
#   pyinstaller --onefile --icon="C:\Users\Intel Core i7\Desktop\python(2)\image_path\logo.ico" 正式版.py  打包成exe文件


"""
关键词智能推荐：根据已有的热词库，结合当前电商行业的趋势和用户搜索习惯，动态推荐高相关性、高转化率的关键词。

SEO优化提示：在生成标题时，可以添加一个SEO友好度评分或建议，帮助用户优化标题以提高搜索引擎排名。

自定义模板功能：用户可以设置多种标题格式模板，并与导入的前缀关键词自由组合，以便适应不同的商品类型或营销策略。

多语言支持：针对跨境电商平台，提供多语言标题生成能力，可以根据目标市场自动翻译或适配。

图片压缩：在图片处理模块中，增加图片压缩功能，保证裁剪后的图片大小满足电商平台上传限制。

图片水印添加：允许用户在拼接或裁剪图片过程中添加自己的店铺Logo或其他水印信息。

一键同步到店铺：将生成的商品标题和图片直接同步到对应的电商平台店铺，减少手动上传的繁琐步骤。

数据分析反馈：分析生成标题后的产品销售数据，如点击率、转化率等，进一步优化标题生成算法。


2.8.2:修复了拼接图片尺寸不对的问题
"""
import random
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox, Toplevel, Text
from ttkthemes import ThemedTk
import pandas as pd
import jieba
from collections import Counter
from PIL import Image, ImageTk
from tkinter import messagebox
import re
import os
from PIL import Image
from tkinter import filedialog, messagebox
from tkinter import filedialog
from PIL import Image, ImageDraw, ImageFont
import threading
from PIL import Image, ImageFilter




# 主窗口类
class DataAnalysisApp:
    def __init__(self, master):
        self.master = master
        master.title("虾皮多功能工具集2.8.2——by唐大帅---有问题请点击联系作者添加我")
        master.geometry("600x500")

        # 新增的主题列表
        self.themes = ['radiance', 'arc']

        #主题
        bottom_buttons_frame = ttk.Frame(master)
        self.change_theme_button = ttk.Button(bottom_buttons_frame, text="更换主题", command=self.change_theme)
        # 使用说明
        self.instructions_button = ttk.Button(bottom_buttons_frame, text="使用说明", command=self.show_instructions)
        
        # 添加联系作者按钮
        self.contact_author_button = tk.Button(master, text="联系作者", command=self.show_author_info)
        self.contact_author_button.pack(pady=10)

        # 使用 grid 布局并设置居中对齐
        self.change_theme_button.grid(row=0, column=0, sticky=tk.E)  # 紧贴在右边框
        self.instructions_button.grid(row=0, column=1, sticky=tk.W)  # 紧贴在左边框
        bottom_buttons_frame.grid_rowconfigure(0, weight=1)  # 让这一行填充满整个高度
        bottom_buttons_frame.grid_columnconfigure((0, 1), weight=1)  # 让这两列宽度相等

        # 设置 bottom_buttons_frame 在其父窗口（master）中底部对齐
        bottom_buttons_frame.pack(side=tk.BOTTOM, pady=10)

        self.upload_excel_button = ttk.Button(master, text="上传Excel文件并分析", command=self.upload_and_analyze_excel)
        self.upload_excel_button.pack(pady=10)

        self.upload_title_prefix_frame = ttk.Frame(master)
        self.upload_title_prefix_button = ttk.Button(self.upload_title_prefix_frame, text="导入标题前缀", command=self.upload_title_prefix)
        self.upload_title_prefix_button.pack(side=tk.LEFT)
        self.open_title_prefix_button = ttk.Button(self.upload_title_prefix_frame, text="打开标题前缀文件", command=self.open_title_prefix_file)
        self.open_title_prefix_button.pack(side=tk.LEFT)
        self.upload_title_prefix_frame.pack(pady=10)

        self.upload_filter_frame = ttk.Frame(master)
        self.upload_filter_button = ttk.Button(self.upload_filter_frame, text="导入过滤关键词", command=self.upload_filter_keywords)
        self.upload_filter_button.pack(side=tk.LEFT)
        self.open_filter_button = ttk.Button(self.upload_filter_frame, text="打开过滤关键词文件", command=self.open_filter_keywords_file)
        self.open_filter_button.pack(side=tk.LEFT)
        self.upload_filter_frame.pack(pady=10)

        self.user_input_custom_frame = ttk.Frame(master)
        self.user_input_custom_label = ttk.Label(self.user_input_custom_frame, text="手动输入要排除的词语（空格分隔）：")
        self.user_input_custom_label.pack(side=tk.LEFT)
        self.user_input_custom_entry = ttk.Entry(self.user_input_custom_frame)
        self.user_input_custom_entry.pack(side=tk.LEFT)
        self.user_input_custom_frame.pack(pady=10)

        self.display_hot_words_button = ttk.Button(master, text="显示数据热词", state=tk.DISABLED, command=self.display_hot_words)
        self.display_hot_words_button.pack(pady=5)

        self.generate_titles_button = ttk.Button(master, text="生成标题", state=tk.DISABLED, command=self.generate_titles)
        self.generate_titles_button.pack(pady=5)

        self.save_hot_words_button = ttk.Button(master, text="保存热词", state=tk.DISABLED, command=self.save_hot_words)
        self.save_hot_words_button.pack(pady=5)

        self.more_features_button = ttk.Button(master, text="更多功能", command=self.show_more_features)
        self.more_features_button.pack(pady=10)

        self.word_freq = []  # 保存词频统计结果
        self.title_prefixes = []  # 保存导入的标题前缀
        self.filter_words = []  # 保存导入的过滤关键词
        self.title_prefixes_file_path = ""  # 保存标题前缀文件的路径
        self.filter_words_file_path = ""  # 保存过滤关键词文件的路径

# 更多功能
    def show_more_features(self):
        new_window = Toplevel(self.master)
        new_window.transient(self.master)  # 设置 transient 属性
        new_window.title("更多功能")
        new_window.geometry("+200+100")  # 调整窗口位置,避免遮挡
        new_window.attributes("-topmost", True)  # 设置子窗口在顶层显示

        # 图片拼接按钮
        image_stitch_button = ttk.Button(new_window, text="图片拼接", command=self.handle_image_stitching)
        image_stitch_button.pack(pady=5)

        # 图片裁剪按钮
        image_crop_button = ttk.Button(new_window, text="图片裁剪", command=self.handle_image_cropping)
        image_crop_button.pack(pady=5)

        # 添加水印按钮
        add_watermark_button = ttk.Button(new_window, text="添加水印", command=self.handle_add_watermark)
        add_watermark_button.pack(side=tk.LEFT, pady=5)

        new_window.focus_force()  # 强制设置焦点到新窗口

    # 图片拼接
    def handle_image_stitching(self):
        file_paths = filedialog.askopenfilenames(title="选择要拼接的图片", filetypes=[("图片文件", "*.png *.jpg *.jpeg")])
        
        if not file_paths:
            messagebox.showinfo(title="提示", message="未选择任何图片!")
            return
        
        # 禁用一些控件,避免用户在图片拼接时进行其他操作
        self.disable_controls()
        
        # 创建并启动一个新线程来执行图片拼接操作
        stitching_thread = threading.Thread(target=self.stitch_images, args=(file_paths,))
        stitching_thread.start()

    def stitch_images(self, file_paths):
        try:
            # 按照文件名排序图片
            sorted_file_paths = sorted(file_paths)
            total_images = len(sorted_file_paths)  # 总共需要处理的图片文件数量

            total_height = 0
            target_width = 800  # 设置目标宽度为 800 像素

            # 让用户选择保存图片的目录
            save_path = filedialog.askdirectory(title="选择保存图片的目录")
            if not save_path:
                messagebox.showinfo(title="提示", message="未选择保存目录!")
                return

            # 在用户选择的目录下创建裁剪img文件夹
            cut_folder_path = os.path.join(save_path, '裁剪img')
            os.makedirs(cut_folder_path, exist_ok=True)

            # 创建进度窗口
            progress_window = Toplevel(self.master)
            progress_window.title("图片拼接进度")
            progress_window.geometry("400x100+500+300")
            progress_bar = ttk.Progressbar(progress_window, length=300, mode='determinate')
            progress_bar.pack(pady=10)
            progress_window.attributes("-topmost", True)

            # 初始化进度条
            progress_bar['value'] = 0
            progress_bar['maximum'] = total_images

            current_step = 0

            resized_images = []

            # 循环遍历所有图片,将它们等比缩放至目标宽度
            for path in sorted_file_paths:
                img = Image.open(path)
                width, height = img.size

                # 计算等比缩放后的高度
                aspect_ratio = target_width / width
                new_height = int(height * aspect_ratio)

                # 等比缩放图片
                resized_img = img.resize((target_width, new_height), Image.Resampling.LANCZOS)
                resized_images.append(resized_img)

                current_step += 1
                progress_bar['value'] = current_step
                progress_bar.update()

            # 计算拼接后的总高度
            total_height = sum(img.height for img in resized_images)

            # 创建一个新的图像,用于拼接所有等比缩放后的图片
            result_img = Image.new('RGB', (target_width, total_height), color=(255, 255, 255))
            y_offset = 0

            # 将所有等比缩放后的图片粘贴到结果图像上
            for img in resized_images:
                result_img.paste(img, (0, y_offset))
                y_offset += img.height

            # 将拼接后的图片保存到裁剪img文件夹中
            output_path = os.path.join(cut_folder_path, 'stitched_image.png')
            result_img.save(output_path)

            # 关闭进度窗口
            progress_window.destroy()

            messagebox.showinfo(title="成功", message=f"图片已成功拼接,并保存至:{output_path}")
        except (IOError, OSError) as e:
            messagebox.showerror(title="错误", message=f"图片拼接失败!\n错误信息: {str(e)}")
        except Exception as e:
            messagebox.showerror(title="错误", message=f"发生未知错误:\n{str(e)}")
        finally:
            # 启用之前被禁用的控件
            self.enable_controls()
                
    def disable_controls(self):
        # 禁用一些控件,避免用户在图片拼接时进行其他操作
        self.upload_excel_button.config(state=tk.DISABLED)
        self.upload_title_prefix_button.config(state=tk.DISABLED)
        self.open_title_prefix_button.config(state=tk.DISABLED)
        self.upload_filter_button.config(state=tk.DISABLED)
        self.open_filter_button.config(state=tk.DISABLED)
        self.user_input_custom_entry.config(state=tk.DISABLED)
        self.display_hot_words_button.config(state=tk.DISABLED)
        self.generate_titles_button.config(state=tk.DISABLED)
        self.save_hot_words_button.config(state=tk.DISABLED)

    def enable_controls(self):
        # 启用之前被禁用的控件
        self.upload_excel_button.config(state=tk.NORMAL)
        self.upload_title_prefix_button.config(state=tk.NORMAL)
        self.open_title_prefix_button.config(state=tk.NORMAL)
        self.upload_filter_button.config(state=tk.NORMAL)
        self.open_filter_button.config(state=tk.NORMAL)
        self.user_input_custom_entry.config(state=tk.NORMAL)
        self.display_hot_words_button.config(state=tk.NORMAL)
        self.generate_titles_button.config(state=tk.NORMAL)
        self.save_hot_words_button.config(state=tk.NORMAL)
        self.more_features_button.config(state=tk.NORMAL)


    # 图片裁剪
    def handle_image_cropping(self):
        file_paths = filedialog.askopenfilenames(title="选择要裁剪的图片", filetypes=[("图片文件", "*.png *.jpg *.jpeg")])

        if not file_paths:
            messagebox.showinfo(title="提示", message="未选择任何图片!")
            return

        try:
            # 统计总共需要裁剪的图片数量
            total_crops = 0
            for path in file_paths:
                img = Image.open(path)
                width, height = img.size

                crop_height = 1400
                num_crops = int(height / crop_height)
                if height % crop_height != 0:
                    num_crops += 1

                total_crops += num_crops

            # 创建进度窗口
            progress_window = Toplevel(self.master)
            progress_window.title("图片裁剪进度")
            progress_window.geometry("400x100+500+300")
            progress_label = ttk.Label(progress_window, text=f"总共需要裁剪 {total_crops} 张图片...", font=("Arial", 14))
            progress_label.pack(pady=10)
            progress_window.attributes("-topmost", True)

            current_image = 1

            for i, path in enumerate(sorted(file_paths)):
                img = Image.open(path)
                width, height = img.size

                crop_height = 1400
                num_crops = int(height / crop_height)
                if height % crop_height != 0:
                    num_crops += 1

                for j in range(num_crops):
                    top = j * crop_height
                    bottom = (j + 1) * crop_height
                    if bottom > height:
                        bottom = height

                    cropped_img = img.crop((0, top, width, bottom))
                    file_name = f"img_{i + 1}_{j + 1}"
                    output_path = os.path.join(os.path.dirname(path), f"{file_name}.png")
                    cropped_img.save(output_path)

                    progress_label.config(text=f"正在裁剪第 {current_image}/{total_crops} 张图片...")
                    current_image += 1
                    progress_window.update()

                if i == len(file_paths) - 1:
                    messagebox.showinfo(title="成功", message=f"图片 {os.path.basename(path)} 已成功裁剪并保存至原文件夹!")

            progress_window.destroy()
        except Exception as e:
            messagebox.showerror(title="错误", message=f"图片裁剪过程中发生错误: {e}")

    # 添加水印元素
    def handle_add_watermark(self):
        watermark_window = Toplevel(self.master)
        watermark_window.transient(self.master)
        watermark_window.title("添加水印")
        watermark_window.geometry("+300+200")
        watermark_window.attributes("-topmost", True)

        # 文本编辑框
        watermark_text_frame = ttk.Frame(watermark_window)
        watermark_text_label = ttk.Label(watermark_text_frame, text="水印文字:")
        watermark_text_label.pack(side=tk.LEFT)
        watermark_text_entry = Text(watermark_text_frame, height=3, width=30)  # 将 ttk.Entry 改为 Text,可以换行
        watermark_text_entry.pack(side=tk.LEFT)
        watermark_text_frame.pack(pady=5)

        # 上传图片按钮
        upload_image_button = ttk.Button(watermark_window, text="上传图片", command=lambda: self.upload_images(watermark_window))
        upload_image_button.pack(pady=5)

        # 透明度滑块
        transparency_label = ttk.Label(watermark_window, text="透明度:")
        transparency_label.pack(pady=5)
        transparency_scale = ttk.Scale(watermark_window, from_=0, to=100, orient=tk.HORIZONTAL, command=lambda val: transparency_value.set(f"当前透明度: {int(float(val))}%"))
        transparency_scale.pack(pady=5)
        transparency_value = tk.StringVar()
        transparency_value.set("当前透明度: 0%")
        transparency_value_label = ttk.Label(watermark_window, textvariable=transparency_value)
        transparency_value_label.pack(pady=5)

        # 保存图片按钮
        save_image_button = ttk.Button(watermark_window, text="保存图片", command=lambda: self.save_watermarked_images(watermark_window, watermark_text_entry.get("1.0", "end-1c"), transparency_scale.get()))
        save_image_button.pack(pady=5)

        watermark_window.focus_force()

    # 上传水印图片图片
    def upload_images(self, watermark_window):
        file_paths = filedialog.askopenfilenames(title="选择图片文件", filetypes=[("Image Files", "*.png *.jpg *.jpeg")])
        if file_paths:
            self.image_paths = list(file_paths)
            messagebox.showinfo("成功", f"已成功上传 {len(self.image_paths)} 张图片")

    # 保存水印图片
    def save_watermarked_images(self, watermark_window, watermark_text, transparency):
        if not hasattr(self, 'image_paths') or not self.image_paths:
            messagebox.showerror("错误", "请先上传图片!")
            return

        try:
            # 获取用户选择的保存目录
            save_dir = filedialog.askdirectory(title="选择保存目录")
            if not save_dir:
                messagebox.showinfo("提示", "未选择保存目录!")
                return

            for image_path in self.image_paths:
                image = Image.open(image_path)
                watermark = Image.new("RGBA", image.size, (0, 0, 0, 0))
                draw = ImageDraw.Draw(watermark)

                # 构建字体文件的完整路径
                font_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'TaipeiSansTCBeta-Bold.ttf')
                font = ImageFont.truetype(font_path, 36)

                # 自动换行
                lines = []
                words = watermark_text.split()
                line = ""
                for word in words:
                    new_line = line + word + " "
                    line_width, _ = font.getbbox(new_line)[2:]
                    if line_width > image.width - 50:
                        lines.append(line.strip())
                        line = word + " "
                    else:
                        line = new_line
                if line:
                    lines.append(line.strip())

                y = 10
                for line in lines:
                    line_width, line_height = font.getbbox(line)[2:]
                    x = (image.width - line_width) / 2
                    draw.text((x, y), line, font=font, fill=(255, 255, 255, int(255 * (100 - transparency) / 100)))
                    y += line_height + 5

                watermarked_image = Image.alpha_composite(image.convert("RGBA"), watermark)

                # 保存水印图片到用户选择的目录下的 watermarked_images 文件夹中
                filename = os.path.splitext(os.path.basename(image_path))[0] + "_Watermark" + os.path.splitext(image_path)[1]
                output_dir = os.path.join(save_dir, 'watermarked_images')
                os.makedirs(output_dir, exist_ok=True)
                output_path = os.path.join(output_dir, filename)
                watermarked_image.save(output_path)

            messagebox.showinfo("成功", f"已成功保存 {len(self.image_paths)} 张水印图片到 {os.path.join(save_dir, 'watermarked_images')}")
        except Exception as e:
            messagebox.showerror("错误", f"添加水印失败: {e}")

        watermark_window.destroy()
    # 联系作者
    def show_author_info(self):
        # 获取图片文件的路径
        image_path = os.path.join(os.path.dirname(__file__), 'wx.png')

        # 检查图片文件是否存在
        if not os.path.exists(image_path):
            messagebox.showerror("错误", "找不到图片文件")
            return

        # 创建一个新的窗口
        new_window = tk.Toplevel(self.master)
        new_window.title("联系作者")

        # 打开图片文件
        image = Image.open(image_path)

        # 创建一个PhotoImage对象
        photo = ImageTk.PhotoImage(image)

        # 在新窗口中显示图片
        label = tk.Label(new_window, image=photo)
        label.image = photo  # 保持对photo的引用
        label.pack()

        # 在图片下方添加文本
        text = tk.Label(new_window, text="有问题请联系我！\n 本土货代，本土1对1账号，\n1对1陪跑，本土OBS无人直播\n2024最新延迟出货清分\n.....\n\n\n只有你想不到，没有我做不到！！", font=("Arial", 15))
        text.pack()
    
    # 更换主题    
    def change_theme(self):
        # 随机选择一个主题
        theme = random.choice(self.themes)
        self.master.set_theme(theme)
    
    # 显示使用说明
    def show_instructions(self):
            instructions_text = """
                      使用说明：—————— by 唐大帅\n\n
                \n 软件没有在线更新的功能，功能只会越来越多越来越好用，如果想获取最新软件，请点击最上方的联系作者按钮\n\n
                1. “上传Excel文件并分析”：\n点击此按钮选择一个Excel文件进行数据分析和词频统计，将分析出哪些关键词的热度比较高(可以点击保存热词查看)。\n\n
                
                2. “导入标题前缀”和“打开标题前缀文件”：\n用于导入或选择标题前缀文本文件，这些前缀将被用于生成的标题中。提示：(请在标题前缀.txt中进行更改【一行一个】)\n\n
                
                3. “导入过滤关键词”和“打开过滤关键词文件”：\n用于导入或选择过滤关键词文本文件，这些关键词在生成标题时会被排除。\n\n
                
                4. 输入需要排除的词语：\n在此文本框中手动输入要排除的词语（用空格分隔），这些词语在生成标题时也会被排除。提示：每次手动输入都会比较麻烦，建议每次手动输入后，将需要过滤的关键字写进过滤关键词.txt文件中\n\n
                
                5. “显示数据热词”：\n点击后会在新窗口中显示从Excel文件中分析出的数据热词。将会分析出哪些关键词的热度比较高及出现的次数，并且按照从高到低排序。\n\n
                
                6. “生成标题”：\n随机生成5个商品标题供你复制，如果不满意重新点击此按钮。\n\n
                
                7. “保存热词”：\n将热词列表保存为Excel文件，配合数据分析软件使用。\n\n

                9. “联系作者”：\n点击之后，你可以联系我。\n\n

                10.更多功能：目前新增了图片拼接，图片裁剪的功能，两者搭配可以快速裁剪出适合虾皮详情的尺寸的图片。\n\n

                8. “使用说明”：\n字面意思。\n
                
            """
            self.show_text_in_new_window("使用说明", instructions_text)
            
    # 上传Excel文件并分析
    def upload_and_analyze_excel(self):
        file_path = filedialog.askopenfilename(
            title="请选择Excel数据文件", 
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        if not file_path:
            return
        
        try:
            data_frame = pd.read_excel(file_path)
            text = " ".join(re.findall(r'[\u4e00-\u9fff]+', data_frame.to_string()))
            words = jieba.lcut(text)
            self.word_freq = Counter(word for word in words if len(word) > 1).most_common()

            if self.word_freq:
                self.display_hot_words_button['state'] = tk.NORMAL
                self.generate_titles_button['state'] = tk.NORMAL
                self.save_hot_words_button['state'] = tk.NORMAL
                messagebox.showinfo(title="成功", message="数据分析和词频统计完成。")
            else:
                messagebox.showwarning(title="警告", message="Excel文件中未找到足够的文本数据。")
        except Exception as e:
            messagebox.showerror(title="错误", message=f"读取或分析Excel文件失败: {e}")


    # 上传标题前缀
    def upload_title_prefix(self):
        self.title_prefixes_file_path = filedialog.askopenfilename(
            title="请选择标题前缀文本文件", 
            filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")]
        )
        if not self.title_prefixes_file_path:
            return

        try:
            with open(self.title_prefixes_file_path, 'r', encoding='utf-8') as file:
                self.title_prefixes = [line.strip() for line in file if line.strip()]
                messagebox.showinfo(title="成功", message="标题前缀文本导入成功。")
        except Exception as e:
            messagebox.showerror(title="错误", message=f"读取标题前缀文本文件失败: {e}")

    # 打开标题前缀文件
    def open_title_prefix_file(self):
        if self.title_prefixes_file_path:
            os.startfile(self.title_prefixes_file_path)
        else:
            messagebox.showinfo(title="提示", message="尚未导入标题前缀文件。")

    # 上传过滤关键词
    def upload_filter_keywords(self):
        self.filter_words_file_path = filedialog.askopenfilename(
            title="请选择过滤关键词文本文件", 
            filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")]
        )
        if not self.filter_words_file_path:
            return

        try:
            with open(self.filter_words_file_path, 'r', encoding='utf-8') as file:
                self.filter_words = [line.strip() for line in file if line.strip()]
                messagebox.showinfo(title="成功", message="过滤关键词文本导入成功。")
        except Exception as e:
            messagebox.showerror(title="错误", message=f"读取过滤关键词文本文件失败: {e}")

    # 打开过滤关键词文件
    def open_filter_keywords_file(self):
        if self.filter_words_file_path:
            os.startfile(self.filter_words_file_path)
        else:
            messagebox.showinfo(title="提示", message="尚未导入过滤关键词文件。")

    # 显示数据热词
    def display_hot_words(self):
        hot_words_text = "\n".join([f"{word}: {count}次" for word, count in self.word_freq])
        self.show_text_in_new_window("数据热词展示", hot_words_text)

    # 生成标题
    def generate_titles(self):
        exclude_words_set = set(self.filter_words)
        # 添加用户自定义过滤词
        custom_words = self.user_input_custom_entry.get().split()
        exclude_words_set.update(custom_words)

        available_words = [word for word, count in self.word_freq if word not in exclude_words_set and len(word) > 1][:20]
        titles = []
        for _ in range(5):
            if self.title_prefixes:
                title_prefix = random.choice(self.title_prefixes)
            else:
                title_prefix = ""
            random.shuffle(available_words)
            title = f"{title_prefix} {' '.join(available_words)}".strip()
            titles.append(title)
        generated_titles = "\n".join(titles)
        self.show_text_in_new_window("随机生成的标题", generated_titles)

    # 保存热词
    def save_hot_words(self):
        save_path = filedialog.asksaveasfilename(
            title="保存热词Excel文件",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")]
        )
        if not save_path:
            return

        try:
            hot_words_data = pd.DataFrame(self.word_freq, columns=['词语', '出现次数'])
            hot_words_data.to_excel(save_path, index=False)
            messagebox.showinfo(title="成功", message="热词已成功保存到Excel文件。")
        except Exception as e:
            messagebox.showerror(title="错误", message=f"保存热词到Excel文件失败: {e}")

    # 显示文本
    def show_text_in_new_window(self, title, text):
        new_window = Toplevel(self.master)
        new_window.title(title)
        text_widget = Text(new_window, wrap='word')
        text_widget.insert(tk.END, text)
        text_widget.config(state='disabled')
        text_widget.pack(padx=10, pady=10)

# 主函数
def main():
    root = tk.Tk()
    app = DataAnalysisApp(root)
    root.mainloop()

# 主函数
def main():
    root = ThemedTk(theme="radiance")  # 使用ThemedTk替换tk.Tk
    app = DataAnalysisApp(root)
    root.mainloop()
if __name__ == "__main__":
    main()
