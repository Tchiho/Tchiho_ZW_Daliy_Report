from spire.presentation import Presentation

def Trans():
    presentation = Presentation()
    # 加载一个 PowerPoint 演示文稿
    presentation.LoadFromFile("./Output/ZW_Daily_Report.pptx")

    # 遍历演示文稿中的幻灯片
    for i in range(presentation.Slides.Count):
        # 获取幻灯片
        slide = presentation.Slides.get_Item(i)
        # 指定输出文件名
        fileName ="./Output/ZW_Daily_Report.png"
        # 将每个幻灯片保存为 PNG 图像
        image = slide.SaveAsImage()
        # 或者使用指定的宽度和高度保存幻灯片为图像
        # image = slide.SaveAsImageByWH(800, 600)
        image.Save(fileName)
        image.Dispose()
    presentation.Dispose()
