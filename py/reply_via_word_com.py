"""
⭐ 推荐方案：用 Word COM 接口添加回复批注（需 Windows + Word）
然后用 zipfile 后处理修改作者名
"""
import os, zipfile, re
from lxml import etree

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

def reply_comments(src_docx, output_docx, replies):
    """
    replies: [(关键词或ID匹配函数, 回复内容), ...]
    
    注意：必须在 Windows 上安装了 Word 才能运行
    """
    try:
        import win32com.client as win32
    except ImportError:
        print('❌ 需要 Windows + pywin32: pip install pywin32')
        return

    # 1. 记录现有批注 ID
    with zipfile.ZipFile(src_docx, 'r') as z:
        src_tree = etree.fromstring(z.read('word/comments.xml'))
    existing_ids = {c.get(W + 'id') for c in src_tree.findall('.//' + W + 'comment')}

    # 2. 用 COM 打开文档并添加回复
    word = win32.gencache.EnsureDispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = False
    doc = word.Documents.Open(os.path.abspath(src_docx))

    pending = []
    for i in range(1, doc.Comments.Count + 1):
        c = doc.Comments(i)
        for matcher, reply_text in replies:
            if callable(matcher) and matcher(c.Range.Text):
                pending.append((i, reply_text))
            elif isinstance(matcher, str) and matcher in c.Range.Text:
                pending.append((i, reply_text))

    offset = 0
    for orig_idx, reply_text in pending:
        c = doc.Comments(orig_idx + offset)
        c.Replies.Add(c.Scope, reply_text)
        offset += 1

    doc.Save()
    doc.Close(SaveChanges=False)
    word.Quit()

    # 3. 后处理：修改新批注的作者
    import shutil
    shutil.copy2(src_docx, output_docx)
    tmp = output_docx + '.tmp'

    with zipfile.ZipFile(output_docx, 'r') as zin, \
         zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == 'word/comments.xml':
                text = data.decode('utf-8')
                def fix_author(m):
                    cid = m.group(1)
                    if cid not in existing_ids:
                        return f'<w:comment w:id="{cid}" w:author="AI Agent" w:date="{m.group(3)}" w:initials="A"'
                    return m.group(0)
                text = re.sub(
                    r'<w:comment w:id="(\d+)" w:author="([^"]+)" w:date="([^"]+)" w:initials="([^"]+)"',
                    fix_author, text)
                data = text.encode('utf-8')
            zout.writestr(item, data)
    os.replace(tmp, output_docx)
    print(f'✅ 已添加回复并保存: {output_docx}')

if __name__ == '__main__':
    print('⚠️ 此脚本需要 Windows + Word 环境运行')
    print('用法示例:')
    print('  reply_via_word_com.py("input.docx", "output.docx", [("关键词", "回复内容")])')
