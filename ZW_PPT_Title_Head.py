from datetime import datetime
from pptx import Presentation
from pptx.util import Cm, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from PIL import Image, ImageDraw, ImageFont, ImageFilter

def PPT_Title_Head():
    # 获取当前日期
    current_date = datetime.now()
    # 获取中文星期的名称
    weekday_cn = {0: '星期一', 1: '星期二', 2: '星期三', 3: '星期四', 4: '星期五', 5: '星期六', 6: '星期日'}
    weekday_name = weekday_cn[current_date.weekday()]
    # 格式化日期为"YYYY年M月D日 星期N"格式
    formatted_date = f"{current_date.year}年{current_date.month}月{current_date.day}日 {weekday_name}"
    # 打开图片与设定文本内容
    background = Image.new('RGB', (22150,3500), color=(255, 255, 255))
    title_head = "孝感公众客户装维工单日通报"
    title_time = formatted_date
    # 字体使用微软黑体
    font_title_head = ImageFont.truetype('msyhbd.ttc', 1200,)
    font_title_time = ImageFont.truetype('msyhbd.ttc', 800)

    title_head_box = font_title_head.getbbox(title_head)
    title_time_box = font_title_time.getbbox(title_time)
    # 绘画模板
    draw_title_head = ImageDraw.Draw(background)
    draw_title_time = ImageDraw.Draw(background)
    draw_title_head_shadow = ImageDraw.Draw(background)
    draw_title_time_shadow = ImageDraw.Draw(background)
    # 标题头与标题时间定位
    draw_title_head_x = (22150 - title_head_box[2] + title_head_box[0]) / 2
    draw_title_head_y = (title_head_box[3] - title_head_box[1]) / 4
    draw_title_time_x = (22150 - title_time_box[2] + title_time_box[0]) / 2
    draw_title_time_y = (title_time_box[3] - title_time_box[1]) + draw_title_head_y * 3 + 200
    # 先绘制阴影，让其在图层之下
    draw_title_head_shadow.text((draw_title_head_x, draw_title_head_y + 100), title_head, fill=(128, 128, 128 ,102), font = font_title_head)
    draw_title_time_shadow.text((draw_title_time_x, draw_title_time_y + 100), title_time, fill=(128, 128, 128 ,102), font = font_title_time)
    # 分离阴影图像层并应用高斯模糊
    shadow_layer = background.copy()
    shadow_layer.paste(shadow_layer, box=None, mask=None)
    shadow_layer = shadow_layer.filter(ImageFilter.GaussianBlur(radius=50))
    # 将模糊的阴影层合并回背景图像
    background.paste(shadow_layer, box=None, mask=None)
    # 文本绘制
    draw_title_head.text((draw_title_head_x, draw_title_head_y), title_head, fill=(255, 0, 0), font = font_title_head)
    draw_title_time.text((draw_title_time_x, draw_title_time_y), title_time, fill=(255, 0, 0), font = font_title_time)
    # 保存
    background.save("./Pic/Title_Head.png", "PNG")