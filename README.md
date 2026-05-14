# Word Comments Guide

Python 操作 Word 批注（创建 / 读取 / 回复）的完整技术方案。

## 核心结论

| 操作 | 方案 |
|------|------|
| 创建新批注 | python-docx + 手工注入 XML |
| 读取批注 | 解析 `word/comments.xml` + `commentsExtended.xml` |
| 回复批注（嵌套显示） | **必须用 Word COM 接口** |
| 修改批注作者 | COM 完成后 zipfile 正则后处理 |

> ⚠️ Word 365 对回复批注的 XML 校验极严格，手工拼 XML 只会显示为独立批注，无法嵌套。

## 快速开始

```bash
pip install python-docx lxml
python py/create_doc_with_comments.py   # 生成带批注的示例文档
python py/read_comments.py output.docx  # 读取批注
```

## 项目结构

```
py/
├── create_doc_with_comments.py   # 创建带批注的文档（XML注入）
├── read_comments.py              # 读取批注 + 回复关系
├── verify_replies.py             # 验证批注父子关系
├── locate_comments.py            # 定位批注标注的具体文字
├── inspect_xml.py                # 输出批注相关 XML 原始内容
└── reply_via_word_com.py         # ⭐ COM加回复 + 后处理改作者（需Windows）
docs/
└── Word批注操作完整指南.md         # 完整技术文档
```

## 详细文档

👉 [完整技术指南](docs/Word批注操作完整指南.md) — XML结构、命名空间、调试技巧、踩坑记录

## License

MIT
