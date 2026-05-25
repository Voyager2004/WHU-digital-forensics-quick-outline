# 数字取证快查大纲

本项目整理了一份用于武汉大学数字取证课程开卷考试复习参考的快查大纲，并提供生成 DOCX 文档的脚本。

> 非官方项目：本仓库仅为个人课程复习资料整理，不代表武汉大学、课程组、教材作者或出版社立场。

## 对应教材

本快查大纲对应教材：

《数字取证》  
陈晶、张俊、何琨、郭永健、朱勇宇 编著  
清华大学出版社，2023 年 8 月  
ISBN：9787302639749  
丛书：网络空间安全学科系列教材

## 仓库内容

- `data/outline.zh-CN.json`：快查大纲结构化数据。
- `tools/build_docx.py`：从结构化数据生成 DOCX 和 Markdown。
- `dist/数字取证快查大纲.docx`：已生成的快查大纲文档。
- `docs/copyright-boundary.md`：版权边界说明。

## 版权边界

本仓库不包含：

- 教材 PDF 或扫描件
- 教材 OCR / MinerU 提取文本
- 教材全文、长段释义或原文摘录
- 课后题
- 课后题答案
- 课程考试题或未公开资料

仓库中的大纲仅用于帮助合法持有教材的学习者快速定位知识点。页码引用仅作为查书索引使用。

## 生成文档

建议使用 Python 3.10+。

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python tools/build_docx.py
```

默认会生成：

- `dist/数字取证快查大纲.docx`
- `dist/数字取证快查大纲.md`

也可以指定输出路径：

```bash
python tools/build_docx.py --output dist/数字取证快查大纲.docx --markdown dist/数字取证快查大纲.md
```

## 许可证

代码部分使用 MIT License，见 `LICENSE-CODE`。

大纲内容使用 CC BY-NC-SA 4.0 License，见 `LICENSE-CONTENT`。
