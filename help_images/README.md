# help_images/

此目录存放用户指南中的截图/示意图。

## 使用方法

1. 将截图（PNG / JPG）放到此目录
2. 在 `utils/user_guide.py` 的 `USER_GUIDE_SECTIONS` 相应位置添加：
   ```python
   {"type": "image", "file": "你的截图.png", "width": 580},
   ```
3. 可选 `"caption"` 字段添加图片说明文字

## 推荐截图尺寸

- 宽度 580px（匹配文档窗口宽度）
- 格式 PNG
