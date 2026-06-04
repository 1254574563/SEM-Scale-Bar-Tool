"""
SEM图片自动排版生成PPT脚本

功能说明：
1. 选择母文件夹，并检查 PlainImages 子文件夹是否存在
2. 按同名 .tif 文件进行母文件夹 / 子文件夹图片配对
3. 选择倍率并获取标尺文字、长度、线宽
4. 创建16:9比例PPT
5. 每页按 3 列排版，每列显示一组配对图片：
   - 上方独立框：母文件夹图片
   - 下方独立框：PlainImages 子文件夹图片
6. 每个独立框左上角添加白底文本框
7. 每个独立框左下角添加标尺
8. 自动分页并保存为 PPT 文件
"""
"""更新日志2026年3月19日
1、生成ppt路径修改
2、生成ppt名称修改
3、 图片左上角编号字体大小位置调整
4、图片大小调整
5、标尺字体调整(18号、新罗马)
6、 不同放大倍数的标尺长度宽度数据调整;
7、
8k对应标尺文字3um, 标尺线的几何信息宽度3.5磅, 长度2.25cm
12k对应标尺文字2um, 标尺线的几何信息宽度3.5磅, 长度2.25cm
16k对应标尺文字1.5um, 标尺线的几何信息宽度3.5磅, 长度2,25cm
24k对应标尺文字1um, 标尺线的几何信息宽度3.5磅, 长度2.25cm
80k对应标尺文字300nm, 标尺线的几何信息宽度3.5磅, 长度2.25cm
40k对应标尺文字600nm, 标尺线的几何信息宽度3.5磅, 长度2.25cm
"""

from pathlib import Path
from tkinter import Tk, filedialog, simpledialog, messagebox

from PIL import Image

from pptx import Presentation
from pptx.util import Cm, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_VERTICAL_ANCHOR


# =============================
# 标尺配置
# =============================

SCALE_OPTIONS = {
    "8k":  {"label": "3 μm",   "length_cm": 2.25, "line_width_pt": 3.5},
    "12k": {"label": "2 μm",   "length_cm": 2.25, "line_width_pt": 3.5},
    "16k": {"label": "1 μm",   "length_cm": 1.5, "line_width_pt": 3.5},
    "24k": {"label": "1 μm",   "length_cm": 2.25, "line_width_pt": 3.5},
    "40k": {"label": "600 nm", "length_cm": 2.25, "line_width_pt": 3.5},
    "80k": {"label": "300 nm", "length_cm": 2.25, "line_width_pt": 3.5},
}


# =============================
# 选择母文件夹
# =============================

def choose_folder():
    root = Tk()
    root.withdraw()

    folder = filedialog.askdirectory(title="选择母文件夹")

    if not folder:
        print("未选择文件夹")
        return None

    return Path(folder)


# =============================
# 检查并配对图片
# =============================

def find_and_pair(parent: Path):
    parent = Path(parent)
    sub = parent / "PlainImages"

    if not sub.exists():
        print("没有找到 PlainImages 子文件夹")
        return None

    parent_imgs = list(parent.glob("*.tif"))
    sub_imgs = list(sub.glob("*.tif"))

    print("母文件夹 tif 数量:", len(parent_imgs))
    print("子文件夹 tif 数量:", len(sub_imgs))

    sub_dict = {f.name: f for f in sub_imgs}
    pairs = []

    for f in parent_imgs:
        if f.name in sub_dict:
            pairs.append((f, sub_dict[f.name]))
        else:
            print("未找到配对:", f.name)

    if not pairs:
        print("没有成功配对的图片")
        return None

    print("成功配对数量:", len(pairs))
    return pairs


# =============================
# 选择倍率
# =============================

def choose_scale():
    root = Tk()
    root.withdraw()

    prompt = (
        "请输入倍率：\n"
        "8k, 12k, 16k, 24k, 40k, 80k"
    )

    while True:
        mag = simpledialog.askstring(
            title="选择倍率",
            prompt=prompt
        )

        if mag is None:
            print("用户取消倍率输入")
            return None

        mag = mag.strip().lower()

        if mag in SCALE_OPTIONS:
            cfg = SCALE_OPTIONS[mag]
            return mag, cfg["label"], cfg["length_cm"], cfg["line_width_pt"]

        messagebox.showerror(
            "输入错误",
            "请输入正确倍率：8k, 12k, 16k, 24k, 40k, 80k"
        )


# =============================
# 图片插入工具：按区域自适应放图，保持比例，不裁剪
# =============================

def add_image_fit_box(slide, img_path: Path, left_cm, top_cm, box_w_cm, box_h_cm):
    with Image.open(img_path) as img:
        img_w_px, img_h_px = img.size

    img_ratio = img_w_px / img_h_px
    box_ratio = box_w_cm / box_h_cm

    if img_ratio >= box_ratio:
        target_w = box_w_cm
        target_h = target_w / img_ratio
    else:
        target_h = box_h_cm
        target_w = target_h * img_ratio

    final_left = left_cm + (box_w_cm - target_w) / 2
    final_top = top_cm + (box_h_cm - target_h) / 2

    slide.shapes.add_picture(
        str(img_path),
        Cm(final_left),
        Cm(final_top),
        width=Cm(target_w),
        height=Cm(target_h),
    )


# =============================
# 绘制单个独立框
# =============================

def draw_single_box(
    slide,
    img_path: Path,
    file_name: str,
    box_left: float,
    box_top: float,
    box_w: float,
    box_h: float,
    scale_label: str,
    scale_len_cm: float,
    scale_line_width_pt: float,
):
    border_color = RGBColor(58, 90, 150)
    white_color = RGBColor(255, 255, 255)
    black_color = RGBColor(0, 0, 0)

    border_pt = 4.4
    inner_offset = 0.10

    label_box_w = 2.5
    label_box_h = 1.16

    # 先放图片，再放边框和其他元素，避免遮挡
    add_image_fit_box(
        slide,
        img_path,
        box_left,
        box_top,
        box_w,
        box_h,
    )

    # 外部独立矩形框
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Cm(box_left),
        Cm(box_top),
        Cm(box_w),
        Cm(box_h),
    )
    shape.fill.background()
    shape.line.width = Pt(border_pt)
    shape.line.color.rgb = border_color
    shape.shadow.inherit = False

    # 左上角白底文本框
    text_left = box_left + inner_offset
    text_top = box_top + inner_offset

    text_bg = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Cm(text_left),
        Cm(text_top),
        Cm(label_box_w),
        Cm(label_box_h),
    )
    text_bg.fill.solid()
    text_bg.fill.fore_color.rgb = white_color
    text_bg.line.fill.background()
    text_bg.shadow.inherit = False

    tf = text_bg.text_frame
    tf.clear()
    tf.margin_left = Cm(0)
    tf.margin_right = Cm(0)
    tf.margin_top = Cm(0)
    tf.margin_bottom = Cm(0)
    tf.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE

    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER

    run = p.add_run()
    run.text = file_name
    run.font.size = Pt(18)
    run.font.name = "Times New Roman"
    run.font.color.rgb = black_color

    # 左下角标尺位置
    scale_margin_left = 0.75
    scale_margin_bottom = 0.65

    scale_left = box_left + scale_margin_left
    scale_top = box_top + box_h - scale_margin_bottom - 0.10

    # 标尺线：用细矩形模拟，兼容性更好
    # 0.12 cm 视觉上接近 3.5 磅
    line_height_cm = 0.12

    line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Cm(scale_left),
        Cm(scale_top - line_height_cm / 2),
        Cm(scale_len_cm),
        Cm(line_height_cm),
    )
    line.fill.solid()
    line.fill.fore_color.rgb = white_color
    line.line.fill.background()
    line.shadow.inherit = False

    # 标尺文字
    scale_text_box = slide.shapes.add_textbox(
        Cm(scale_left + 0.10),
        Cm(scale_top - 0.85),
        Cm(scale_len_cm + 1.2),
        Cm(0.6),
    )
    scale_text_box.shadow.inherit = False

    tf2 = scale_text_box.text_frame
    tf2.clear()
    tf2.margin_left = Cm(0)
    tf2.margin_right = Cm(0)
    tf2.margin_top = Cm(0)
    tf2.margin_bottom = Cm(0)
    tf2.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE

    p2 = tf2.paragraphs[0]
    p2.alignment = PP_ALIGN.LEFT

    run2 = p2.add_run()
    run2.text = scale_label
    run2.font.size = Pt(18)
    run2.font.name = "Times New Roman"
    run2.font.color.rgb = white_color


# =============================
# 生成 PPT
# =============================

def create_ppt_from_pairs(
    pairs,
    scale_label,
    scale_len_cm,
    scale_line_width_pt,
    save_folder: Path
):
    prs = Presentation()
    prs.slide_width = Cm(33.867)   # 16:9
    prs.slide_height = Cm(19.05)

    # 页面尺寸（cm）
    slide_w = 33.867
    slide_h = 19.05

    # 一页 3 列，每列上下两个独立框
    cols = 3
    box_w = 10.66
    box_h = 8.0

    # 横向均匀分布
    gap_x = (slide_w - cols * box_w) / (cols + 1)

    # 纵向：上框 + 下框 + 三段间距（上、中央、下）
    gap_y = (slide_h - 2 * box_h) / 3

    for page_start in range(0, len(pairs), 3):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        page_pairs = pairs[page_start:page_start + 3]

        for idx, pair in enumerate(page_pairs):
            parent_img, child_img = pair
            file_name = parent_img.stem

            col_left = gap_x + idx * (box_w + gap_x)

            top_box_top = gap_y
            bottom_box_top = gap_y * 2 + box_h

            # 上框：母文件夹图片
            draw_single_box(
                slide=slide,
                img_path=parent_img,
                file_name=file_name,
                box_left=col_left,
                box_top=top_box_top,
                box_w=box_w,
                box_h=box_h,
                scale_label=scale_label,
                scale_len_cm=scale_len_cm,
                scale_line_width_pt=scale_line_width_pt,
            )

            # 下框：子文件夹图片
            draw_single_box(
                slide=slide,
                img_path=child_img,
                file_name=file_name,
                box_left=col_left,
                box_top=bottom_box_top,
                box_w=box_w,
                box_h=box_h,
                scale_label=scale_label,
                scale_len_cm=scale_len_cm,
                scale_line_width_pt=scale_line_width_pt,
            )

    folder_name = save_folder.name
    output_file_name = f"{folder_name}_标尺已添加.pptx"
    output_path = save_folder / output_file_name
    prs.save(output_path)
    print(f"完成！PPT 已保存到：\n{output_path}")


# =============================
# 主程序入口
# =============================

if __name__ == "__main__":
    folder = choose_folder()
    if not folder:
        exit()

    pairs = find_and_pair(folder)
    if not pairs:
        exit()

    scale_info = choose_scale()
    if not scale_info:
        exit()

    mag, scale_label, scale_len_cm, scale_line_width_pt = scale_info

    create_ppt_from_pairs(
        pairs=pairs,
        scale_label=scale_label,
        scale_len_cm=scale_len_cm,
        scale_line_width_pt=scale_line_width_pt,
        save_folder=folder
    )
