import os
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm

def generate_report(template_path, output_path, data_dict):
    """
    读取指定路径的图片和文本，生成 Word 报告
    :param template_path: Word 模板路径 (例如: report/templates/template_default.docx)
    :param output_path: 生成报告的保存路径
    :param data_dict: 字典，键为模板中的占位符，值为替换文本或图片绝对路径
    """
    doc = DocxTemplate(template_path)
    context = {}

    IMG_SIZES = {
        "img_spectrumSNR": (Mm(135), Mm(95)),
        "img_linewidth": (Mm(140), Mm(82))
    }
    DEFAULT_IMG_SIZE = (Mm(145), Mm(87))

    for key, value in data_dict.items():
        if key.startswith('img_') and isinstance(value, str) and os.path.exists(value):
            w, h = IMG_SIZES.get(key, DEFAULT_IMG_SIZE)
            context[key] = InlineImage(doc, value, width=w, height=h)
        else:
            # 图片不存在或者是普通文本，直接填入
            if key.startswith('img_') and not os.path.exists(value):
                 context[key] = "【暂无】"
            else:
                 context[key] = value

    doc.render(context)
    doc.save(output_path)