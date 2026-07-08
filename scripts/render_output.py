"""图片渲染器模块：将运行输出 + 图形渲染为白底 PNG 图片。

职责：
- 将运行输出（stdout/stderr）和可选的图形输出渲染为一张 PNG
- 支持中文字体 fallback
- 长输出行自动换行（80 字符/行）
"""

import logging
import os
import textwrap
from typing import Optional

from PIL import Image, ImageDraw, ImageFont

logger = logging.getLogger(__name__)

# ─── 渲染规格常量 ───
CANVAS_WIDTH = 1400
CANVAS_BG = "#FFFFFF"
SEPARATOR_COLOR = "#E0E0E0"
OUTPUT_FONT_SIZE = 24
TITLE_FONT_SIZE = 28
LINE_WRAP_WIDTH = 62  # 字符/行
PADDING = 32
LINE_HEIGHT_EXTRA = 10  # 行间距额外像素


# ─── 字体 Fallback 列表 ───
FONT_CANDIDATES_MONO = [
    r"C:\Windows\Fonts\NotoSansSC-VF.ttf",
    r"C:\Windows\Fonts\msyh.ttc",
    r"C:\Windows\Fonts\simsun.ttc",
    "Noto Sans Mono CJK SC",
    "Microsoft YaHei",
    "SimSun",
    "Noto Sans Mono",
    "DejaVu Sans Mono",
]

FONT_CANDIDATES_SANS = [
    r"C:\Windows\Fonts\NotoSansSC-VF.ttf",
    r"C:\Windows\Fonts\msyh.ttc",
    r"C:\Windows\Fonts\simsun.ttc",
    r"C:\Windows\Fonts\simhei.ttf",
    "Noto Sans CJK SC",
    "Microsoft YaHei",
    "SimSun",
    "SimHei",
    "Noto Sans",
    "WenQuanYi Micro Hei",
    "WenQuanYi Zen Hei",
    "DejaVu Sans",
]


def _find_font(font_names: list[str], size: int) -> ImageFont.FreeTypeFont | None:
    """在系统中查找第一个可用字体，返回 ImageFont 对象。

    Args:
        font_names: 字体名称列表（按优先级排列）
        size: 字体大小

    Returns:
        ImageFont 对象，若全部不可用则返回 None
    """
    for name in font_names:
        try:
            if os.path.exists(name):
                return ImageFont.truetype(name, size)
            return ImageFont.truetype(name, size)
        except (OSError, IOError):
            try:
                # 尝试通过 fc-list 查找
                import subprocess
                result = subprocess.run(
                    ["fc-list", f":family={name}", "file"],
                    capture_output=True, text=True, timeout=5,
                )
                for line in result.stdout.strip().split("\n"):
                    if ":" in line:
                        font_path = line.split(":")[0].strip()
                        if font_path:
                            return ImageFont.truetype(font_path, size)
            except Exception:
                pass
            continue
    return None


def _get_fonts() -> tuple[ImageFont.FreeTypeFont, ImageFont.FreeTypeFont, ImageFont.FreeTypeFont]:
    """获取三类字体：标题字体、等宽代码字体、普通输出字体。

    Returns:
        (title_font, code_font, output_font) 三元组

    Raises:
        RuntimeError: 若找不到任何可用字体
    """
    title_font = _find_font(FONT_CANDIDATES_SANS, TITLE_FONT_SIZE)
    code_font = _find_font(FONT_CANDIDATES_MONO, OUTPUT_FONT_SIZE)
    output_font = _find_font(FONT_CANDIDATES_SANS, OUTPUT_FONT_SIZE)

    if title_font is None:
        # 最后尝试系统默认字体
        title_font = ImageFont.load_default()
    if code_font is None:
        code_font = ImageFont.load_default()
    if output_font is None:
        output_font = ImageFont.load_default()

    if all(isinstance(f, ImageFont.ImageFont) and not hasattr(f, 'getbbox')
           for f in [title_font, code_font, output_font]):
        # 全部是默认字体且没有中文字体，给出警告
        logger.warning(
            "未找到中文字体，输出可能无法正确显示中文。"
            "请安装 fonts-noto-cjk：sudo apt install fonts-noto-cjk"
        )

    return title_font, code_font, output_font


def _wrap_text(text: str, width: int = LINE_WRAP_WIDTH) -> list[str]:
    """将长文本按指定宽度自动换行。

    Args:
        text: 原始文本
        width: 每行最大字符数

    Returns:
        换行后的文本行列表
    """
    lines = []
    for line in text.split("\n"):
        if len(line) <= width:
            lines.append(line)
        else:
            lines.extend(textwrap.wrap(line, width=width))
    return lines


def _draw_text_block(
    draw: ImageDraw.ImageDraw,
    text: str,
    x: int,
    y: int,
    font: ImageFont.FreeTypeFont,
    color: str = "#333333",
    bg_color: str | None = None,
    max_width: int | None = None,
    line_height: int | None = None,
) -> int:
    """绘制文本块，返回绘制后的 y 坐标。

    Args:
        draw: ImageDraw 对象
        text: 文本内容
        x: 起始 x 坐标
        y: 起始 y 坐标
        font: 字体
        color: 文本颜色
        bg_color: 背景色（None 则透明）
        max_width: 最大宽度（用于换行），默认使用画布宽度减去边距
        line_height: 行高，默认根据字体大小计算

    Returns:
        绘制后的 y 坐标（下一行起始位置）
    """
    if not text:
        return y

    if max_width is None:
        max_width = LINE_WRAP_WIDTH

    if line_height is None:
        line_height = font.size + LINE_HEIGHT_EXTRA

    wrapped = _wrap_text(text, max_width)

    for line in wrapped:
        if bg_color:
            # 获取文本宽度用于绘制背景
            bbox = draw.textbbox((0, 0), line, font=font)
            text_width = bbox[2] - bbox[0]
            draw.rectangle(
                [(x - 2, y), (x + text_width + 2, y + line_height)],
                fill=bg_color,
            )
        draw.text((x, y), line, font=font, fill=color)
        y += line_height

    return y


def _get_text_height(
    text: str,
    font: ImageFont.FreeTypeFont,
    max_width: int = LINE_WRAP_WIDTH,
    line_height: int | None = None,
) -> int:
    """计算文本块的总高度。

    Args:
        text: 文本内容
        font: 字体
        max_width: 最大宽度
        line_height: 行高

    Returns:
        文本块总高度（像素）
    """
    if not text:
        return 0
    if line_height is None:
        line_height = font.size + LINE_HEIGHT_EXTRA
    wrapped = _wrap_text(text, max_width)
    return len(wrapped) * line_height


def render_code_result(
    task_name: str,
    code: str,
    output_text: str,
    extra_image_path: Optional[str],
    save_path: str,
) -> str:
    """将代码任务的运行结果渲染为白底 PNG 图片。

    图片结构：
    ┌────────────────────────────┐
    │  任务名称（标题）             │
    ├────────────────────────────┤
    │  运行输出（文本）             │
    ├────────────────────────────┤
    │  图形输出（如有）             │
    └────────────────────────────┘

    Args:
        task_name: 任务名称（作为标题）
        code: 保留兼容参数；代码在 Word 中以文字形式展示，不再渲染到图片
        output_text: 运行输出文本
        extra_image_path: 图形输出路径（如 matplotlib 生成的图），可为 None
        save_path: 输出 PNG 路径

    Returns:
        save_path（生成的 PNG 文件路径）

    Raises:
        RuntimeError: 若找不到可用字体
    """
    title_font, _code_font, output_font = _get_fonts()

    output_line_height = output_font.size + LINE_HEIGHT_EXTRA
    title_line_height = title_font.size + LINE_HEIGHT_EXTRA

    # 计算各区域高度
    # 标题区
    title_height = _get_text_height(task_name, title_font, LINE_WRAP_WIDTH, title_line_height)
    title_height += PADDING  # 标题下方间距

    # 输出区
    display_output = output_text.strip() or "（无标准输出）"
    output_height = _get_text_height(display_output, output_font, LINE_WRAP_WIDTH, output_line_height)
    output_height += PADDING

    # 图形区
    extra_image_height = 0
    extra_img = None
    if extra_image_path and os.path.exists(extra_image_path):
        try:
            extra_img = Image.open(extra_image_path)
            # 缩放图形以适应画布宽度
            img_w, img_h = extra_img.size
            scale = (CANVAS_WIDTH - PADDING * 2) / img_w
            if scale < 1:
                img_w = int(img_w * scale)
                img_h = int(img_h * scale)
                extra_img = extra_img.resize((img_w, img_h), Image.LANCZOS)
            extra_image_height = img_h + PADDING
        except Exception as e:
            logger.warning(f"无法加载图形输出 {extra_image_path}: {e}")
            extra_img = None

    # 总画布高度
    total_height = (
        PADDING  # 顶部边距
        + title_height
        + output_height
        + extra_image_height
        + PADDING  # 底部边距
    )

    # 创建画布
    canvas = Image.new("RGB", (CANVAS_WIDTH, total_height), CANVAS_BG)
    draw = ImageDraw.Draw(canvas)

    y = PADDING

    # 绘制标题
    draw.text((PADDING, y), task_name, font=title_font, fill="#1A1A1A")
    y += title_height

    # 绘制分隔线
    draw.line(
        [(PADDING, y), (CANVAS_WIDTH - PADDING, y)],
        fill=SEPARATOR_COLOR, width=1,
    )
    y += PADDING

    # 绘制输出区
    _draw_text_block(
        draw, display_output, PADDING, y, output_font,
        color="#333333",
        max_width=LINE_WRAP_WIDTH, line_height=output_line_height,
    )
    y += output_height

    # 绘制图形输出
    if extra_img:
        y += PADDING
        paste_x = (CANVAS_WIDTH - extra_img.width) // 2
        canvas.paste(extra_img, (paste_x, y))
        extra_img.close()

    # 保存
    save_dir = os.path.dirname(save_path)
    if save_dir:
        os.makedirs(save_dir, exist_ok=True)
    canvas.save(save_path, "PNG")
    logger.info(f"渲染图片已保存: {save_path}")

    return save_path
