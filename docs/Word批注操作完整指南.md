# Word 批注（Comments）操作完整指南

> 本文档面向 AI Agent，记录在 Python 中创建/读取/回复 Word 批注的全部技术细节。
> 基于 2026-05-13 的实战调试结果整理。

---

## 1. 核心结论（先看这个）

| 操作               | 推荐方案                                              | 备注                          |
| ---------------- | ------------------------------------------------- | --------------------------- |
| **创建新批注**        | python-docx + 手工注入 XML                            | 可控、轻量                       |
| **读取批注**         | 直接解析 `word/comments.xml` + `commentsExtended.xml` | 简单                          |
| **回复已有批注（嵌套显示）** | ⚠️ **必须用 Word COM 接口**                            | 手工拼 XML 在 Word 365 中显示为独立批注 |
| **修改批注作者**       | COM 添加完成后，用 zipfile + 正则后处理                       | Word COM 强制使用登录账号名          |

**最重要的教训：** Word 365 对回复批注的 XML 校验非常严格，手工构造看似完整的 XML（comments.xml + commentsExtended.xml + commentsIds.xml + commentsExtensible.xml + document.xml 中的 commentRangeStart/End/Reference 全部齐全），仍可能被识别为"独立批注"而非"回复"。**唯一稳定可靠的回复创建方式是通过 Word COM 接口让 Word 自己生成 XML。**

---

## 2. Word 批注的 XML 结构总览

一份带批注的 `.docx`（本质是 zip 包）涉及以下文件：

```
[Content_Types].xml          ← 声明 comments.xml 等的内容类型
_rels/.rels
word/
  document.xml               ← 正文，含 commentRangeStart/End/Reference 标记
  comments.xml               ← 批注主体内容（必须）
  commentsExtended.xml       ← 父子回复关系（Word 2013+）
  commentsIds.xml            ← paraId ↔ durableId 映射（Word 2016+）
  commentsExtensible.xml     ← UTC 时间戳等扩展信息（Word 365）
  people.xml                 ← 作者信息（Word 2013+）
  _rels/document.xml.rels    ← 声明上述文件的关系
```

### 2.1 各文件之间的关联键

```
comments.xml          : <w:comment w:id="N"> 内嵌 <w:p w14:paraId="XXXXXXXX">
commentsExtended.xml  : <w15:commentEx w15:paraId="XXXXXXXX" w15:paraIdParent="YYYYYYYY">
commentsIds.xml       : <w16cid:commentId w16cid:paraId="XXXXXXXX" w16cid:durableId="ZZZZZZZZ">
document.xml          : <w:commentRangeStart w:id="N">、<w:commentRangeEnd w:id="N">、<w:commentReference w:id="N">
```

- `w:id`（数字）：comments.xml 内部 ID，document.xml 通过它引用批注
- `w14:paraId`（8位十六进制）：跨文件主键，commentsExtended/commentsIds 用它
- `w16cid:durableId`（8位十六进制）：Word 内部稳定 ID

---

## 3. 创建新批注（手工注入）

### 3.1 批注主体 `comments.xml`

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="..." xmlns:w14="..." xmlns:w15="..." mc:Ignorable="w14 w15">
  <w:comment w:id="0" w:author="Claude" w:date="2026-05-13T10:00:00Z" w:initials="C">
    <w:p w14:paraId="1D780707" w14:textId="77777777">
      <w:r><w:annotationRef/></w:r>
      <w:r><w:t>批注内容文字</w:t></w:r>
    </w:p>
  </w:comment>
</w:comments>
```

**注意：**

- `w:date` 是 ISO 8601 格式，用 UTC 带 `Z`，或本地时间不带 Z
- `w14:paraId` 必须唯一，用 8 位十六进制
- `<w:annotationRef/>` 必须放在第一个 `<w:r>` 里，否则 Word 不显示锚点

### 3.2 在 `document.xml` 中标记批注范围

在正文 `<w:p>` 中，被批注文字的前后用三个标签包裹：

```xml
<w:p>
  <w:commentRangeStart w:id="0"/>     ← 批注范围开始
  <w:r><w:t>被批注的文字</w:t></w:r>
  <w:commentRangeEnd w:id="0"/>       ← 批注范围结束
  <w:r>
    <w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>
    <w:commentReference w:id="0"/>    ← 批注锚点
  </w:r>
</w:p>
```

**关键：** `commentRangeStart` 和 `commentRangeEnd` 是兄弟节点，**必须出现在同一个父节点（`<w:p>` 或 `<w:tc>`）下**。不能跨段落配对，否则 Word 显示异常。

### 3.3 注册关系文件

`[Content_Types].xml`：

```xml
<Override PartName="/word/comments.xml"
          ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
```

`word/_rels/document.xml.rels`：

```xml
<Relationship Id="rIdComments"
              Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
              Target="comments.xml"/>
```

### 3.4 完整代码骨架

```python
import zipfile, shutil
from docx import Document

# 1. 用 python-docx 创建基础文档
doc = Document()
doc.add_paragraph('这是被批注的文字所在段落。')
doc.save('base.docx')

# 2. 复制后，操作 zip 包注入 comments.xml 和 document.xml
shutil.copy2('base.docx', 'final.docx')
tmp = 'final.docx.tmp'

with zipfile.ZipFile('final.docx', 'r') as zin, \
     zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == 'word/document.xml':
            data = inject_comment_markers(data, ...)   # 加 commentRangeStart/End/Reference
        elif item.filename == '[Content_Types].xml':
            data = inject_content_type(data)
        elif item.filename == 'word/_rels/document.xml.rels':
            data = inject_rel(data)
        zout.writestr(item, data)
    zout.writestr('word/comments.xml', build_comments_xml(...))

import os
os.replace(tmp, 'final.docx')
```

完整可运行版本见：`py/create_doc_with_comments.py`

---

## 4. 读取批注

简单直接，关键是要同时读 `commentsExtended.xml` 才能判断回复关系：

```python
import zipfile
from lxml import etree

def W(tag):  return '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}' + tag
def W15(tag): return '{http://schemas.microsoft.com/office/word/2012/wordml}' + tag
def W14(tag): return '{http://schemas.microsoft.com/office/word/2010/wordml}' + tag

with zipfile.ZipFile(docx_path, 'r') as z:
    comments_tree = etree.fromstring(z.read('word/comments.xml'))
    ext_tree      = etree.fromstring(z.read('word/commentsExtended.xml'))

# 1. 构建 paraId -> 父paraId 映射
parent_map = {}
for ex in ext_tree.iter(W15('commentEx')):
    pid  = ex.get(W15('paraId'), '')
    ppid = ex.get(W15('paraIdParent'), '')
    parent_map[pid] = ppid

# 2. 遍历所有批注
for c in comments_tree.findall('.//' + W('comment')):
    cid     = c.get(W('id'))
    author  = c.get(W('author'))
    text    = ''.join(t.text for t in c.iter(W('t')) if t.text)
    p       = c.find('.//' + W('p'))
    para_id = p.get(W14('paraId'), '') if p is not None else ''
    parent  = parent_map.get(para_id, '')

    if parent:
        print(f'回复批注 [ID={cid}] {author}: {text} (父 paraId={parent})')
    else:
        print(f'独立批注 [ID={cid}] {author}: {text}')
```

### 4.1 定位批注标注的具体文字

```python
# 在 document.xml 中找 commentRangeStart 和 commentRangeEnd 之间的所有 <w:t>
for start in tree.iter(W('commentRangeStart')):
    cid = start.get(W('id'))
    parent = start.getparent()
    siblings = list(parent)
    idx = siblings.index(start)
    texts = []
    for sibling in siblings[idx:]:
        if etree.QName(sibling.tag).localname == 'commentRangeEnd' \
           and sibling.get(W('id')) == cid:
            break
        for t in sibling.iter(W('t')):
            if t.text: texts.append(t.text)
    print(f'ID={cid} 标注文字：{"".join(texts)}')
```

---

## 5. 回复批注（⚠️ 关键章节）

### 5.1 失败的方案：手工拼 XML

我尝试过完整地手工注入下列内容来构造回复批注，**Word 365 不识别为回复**：

1. ✅ 在 `comments.xml` 添加 `<w:comment>` 主体
2. ✅ 在 `commentsExtended.xml` 添加 `<w15:commentEx ... w15:paraIdParent="父paraId"/>` 建立父子关系
3. ✅ 在 `commentsIds.xml` 添加映射
4. ✅ 在 `commentsExtensible.xml` 添加 UTC 时间戳
5. ✅ 在 `document.xml` 中嵌套插入 `commentRangeStart` + `commentRangeEnd` + `commentReference`
6. ✅ 把 `rsidR` 改成已注册的值
7. ✅ 在 `people.xml` 中注册新作者

**结果：批注都正常显示，但回复显示为独立气泡，不嵌套在父批注下面。**

### 5.2 失败原因分析

Word 365 用了多重机制校验回复关系，至少包括：

- `commentsExtended.xml` 中的 `paraIdParent`
- `commentsExtensible.xml` 中的 `dateUtc` 时区对应关系
- `commentsIds.xml` 中 `durableId` 的内部哈希校验（推测）
- 某些情况下还会校验 `people.xml` 中 `providerId`

手工无法精确复现所有校验项，**即使 XML 结构看起来"完全一致"也不行**。

### 5.3 正确方案：用 Word COM 接口添加回复

```python
import win32com.client as win32

word = win32.gencache.EnsureDispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = False

doc = word.Documents.Open(os.path.abspath(docx_path))

# 关键 API: Comment.Replies.Add(Range, Text)
for i in range(1, doc.Comments.Count + 1):
    c = doc.Comments(i)
    if '需要回复的关键词' in c.Range.Text:
        c.Replies.Add(c.Scope, '回复内容')   # ← 必须传 c.Scope 作为 Range

doc.Save()
doc.Close(SaveChanges=False)
word.Quit()
```

**重要细节：**

1. **`Replies.Add` 签名**：`Add(Range, Text)`，Range 是必填的，传父批注的 `c.Scope`
2. **批注总数会动态变化**：每加一条回复，`doc.Comments.Count` +1。如果用 `for i in range(1, count+1)` 循环，索引会漂移
3. **正确循环模式**：先扫描一遍记录"待回复"清单，再按 offset 修正索引依次添加

```python
# 推荐写法
pending = []
original_count = doc.Comments.Count
for i in range(1, original_count + 1):
    c = doc.Comments(i)
    if 关键词匹配:
        pending.append((i, reply_text))

offset = 0
for original_idx, reply_text in pending:
    idx = original_idx + offset
    c = doc.Comments(idx)
    c.Replies.Add(c.Scope, reply_text)
    offset += 1
```

### 5.4 修正批注作者

Word COM 强制使用登录账号名作为作者（即使设了 `word.UserName = 'Claude'` 也无效）。

**解决方案：** COM 添加完成、保存关闭后，再用 zipfile 后处理：

```python
# 1. 在 COM 之前记录现有批注 ID
with zipfile.ZipFile(SRC, 'r') as z:
    src_tree = etree.fromstring(z.read('word/comments.xml'))
existing_ids = {c.get(W('id')) for c in src_tree.findall('.//' + W('comment'))}

# 2. COM 添加回复，保存关闭

# 3. 后处理：新出现的 ID 就是 Claude 加的回复，把 author 改写
import re
def replace_author(m):
    cid = m.group(1)
    if cid not in existing_ids:
        return f'<w:comment w:id="{cid}" w:author="Claude" w:date="{m.group(3)}" w:initials="C"'
    return m.group(0)

pattern = r'<w:comment w:id="(\d+)" w:author="([^"]+)" w:date="([^"]+)" w:initials="([^"]+)"'
text = re.sub(pattern, replace_author, comments_xml_text)
```

---

## 6. 关键命名空间速查

```
w     = http://schemas.openxmlformats.org/wordprocessingml/2006/main
w14   = http://schemas.microsoft.com/office/word/2010/wordml
w15   = http://schemas.microsoft.com/office/word/2012/wordml
w16cid= http://schemas.microsoft.com/office/word/2016/wordml/cid
w16cex= http://schemas.microsoft.com/office/word/2018/wordml/cex
```

---

## 7. 调试技巧

### 7.1 文件被占用错误

Word 有时不会立刻释放文件锁。表现为：

```
PermissionError: [Errno 13] Permission denied
```

**解决：** 让用户手动关闭 Word 中已打开的文档，或换一个输出文件名。

### 7.2 Word 缓存看不到改动

Word 365 会缓存已打开的文档。修改 docx 后即使重新双击打开，仍可能看到旧内容。

**解决：** 完全退出 Word（任务管理器确认没有 WINWORD.EXE 进程），再打开。

### 7.3 验证 XML 写入正确性

```python
# 检查 document.xml 中所有 commentReference 的 ID 集合
import re
ids = set(re.findall(r'<w:commentReference w:id="(\d+)"/>', doc_xml))
print(f'document.xml 中引用了批注 ID：{sorted(ids, key=int)}')
```

每条批注（**包括回复**）都必须在 `document.xml` 中有对应的 `commentReference`，否则 Word 不显示。

---

## 8. 完整脚本清单

本项目目录 `py/` 下的可参考脚本：

| 脚本                            | 功能                        |
| ----------------------------- | ------------------------- |
| `create_doc_with_comments.py` | 从零创建带批注的文档（手工注入 XML）      |
| `read_comments.py`            | 读取文档所有批注                  |
| `verify_replies.py`           | 验证批注及父子关系                 |
| `locate_comments.py`          | 找出批注标注的具体文字               |
| `inspect_xml.py`              | 输出批注相关 XML 原始内容           |
| `reply_via_word_com.py`       | ⭐ **推荐**：COM 加回复 + 后处理改作者 |
| `minimal_test.py`             | 最小化重现测试                   |
| `reply_nested.py`             | （失败的）手工拼嵌套回复 XML          |

---

## 9. 一句话总结

**读和写独立批注用 zipfile + XML 操作即可；写回复批注必须用 Word COM 接口让 Word 亲自生成 XML，然后用 zipfile 后处理改作者。**
