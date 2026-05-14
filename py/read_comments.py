"""读取 Word 文档中的所有批注（含回复关系）"""
import zipfile, sys
from lxml import etree

W  = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
W14 = '{http://schemas.microsoft.com/office/word/2010/wordml}'
W15 = '{http://schemas.microsoft.com/office/word/2012/wordml}'

def read(docx_path):
    with zipfile.ZipFile(docx_path, 'r') as z:
        comments_tree = etree.fromstring(z.read('word/comments.xml'))
        try:
            ext_tree = etree.fromstring(z.read('word/commentsExtended.xml'))
        except KeyError:
            ext_tree = None

    parent_map = {}
    if ext_tree is not None:
        for ex in ext_tree.iter(W15 + 'commentEx'):
            parent_map[ex.get(W15 + 'paraId', '')] = ex.get(W15 + 'paraIdParent', '')

    results = []
    for c in comments_tree.findall('.//' + W + 'comment'):
        cid = c.get(W + 'id')
        author = c.get(W + 'author')
        text = ''.join(t.text for t in c.iter(W + 't') if t.text)
        p = c.find('.//' + W + 'paraId')
        pid = ''
        for pp in c.iter(W + 'p'):
            pid = pp.get(W14 + 'paraId', '')
            break
        parent = parent_map.get(pid, '')
        results.append({
            'id': cid, 'author': author, 'text': text,
            'para_id': pid, 'parent_para_id': parent, 'is_reply': bool(parent)
        })
    return results

if __name__ == '__main__':
    path = sys.argv[1] if len(sys.argv) > 1 else 'output_with_comments.docx'
    for c in read(path):
        prefix = '  ↳ 回复' if c['is_reply'] else '●'
        print(f"{prefix} [ID={c['id']}] {c['author']}: {c['text']}")
