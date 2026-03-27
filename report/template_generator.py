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

    for key, value in data_dict.items():
        # 如果键名以 img_ 开头，且值是一个存在的本地路径，则作为图片插入
        if key.startswith('img_') and isinstance(value, str) and os.path.exists(value):
            # width=Mm(150), height=Mm(80) 设置图片宽度和高度（会变形）
            context[key] = InlineImage(doc, value, width=Mm(145), height=Mm(87))
        else:
            # 图片不存在或者是普通文本，直接填入
            if key.startswith('img_') and not os.path.exists(value):
                 context[key] = "【暂无】"
            else:
                 context[key] = value

    doc.render(context)
    doc.save(output_path)