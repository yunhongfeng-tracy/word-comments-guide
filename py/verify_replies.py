"""验证批注的父子关系是否正确"""
import zipfile, sys
from lxml import etree

W   = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
W14 = '{http://schemas.microsoft.com/office/word/2010/wordml}'
W15 = '{http://schemas.microsoft.com/office/word/2012/wordml}'

def verify(docx_path):
    with zipfile.ZipFile(docx_path, 'r') as z:
        comments_tree = etree.fromstring(z.read('word/comments.xml'))
        try:
            ext_tree = etree.fromstring(z.read('word/commentsExtended.xml'))
        except KeyError:
            print('⚠️ 无 commentsExtended.xml'); return

    para_ids = set()
    for p in comments_tree.iter(W + 'p'):
        pid = p.get(W14 + 'paraId', '')
        if pid: para_ids.add(pid)

    print('=== 批注父子关系 ===')
    for ex in ext_tree.iter(W15 + 'commentEx'):
        pid = ex.get(W15 + 'paraId', '')
        ppid = ex.get(W15 + 'paraIdParent', '')
        if ppid:
            ok = '✅' if ppid in para_ids else '❌ 父不存在'
            print(f'  {pid} → 父 {ppid} {ok}')

if __name__ == '__main__':
    verify(sys.argv[1] if len(sys.argv) > 1 else 'output_with_comments.docx')
