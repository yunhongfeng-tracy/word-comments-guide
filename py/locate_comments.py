"""定位批注标注的具体文字"""
import zipfile, sys
from lxml import etree

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

def locate(docx_path):
    with zipfile.ZipFile(docx_path, 'r') as z:
        tree = etree.fromstring(z.read('word/document.xml'))

    for start in tree.iter(W + 'commentRangeStart'):
        cid = start.get(W + 'id')
        parent = start.getparent()
        siblings = list(parent)
        idx = siblings.index(start)
        texts = []
        for sib in siblings[idx:]:
            if etree.QName(sib.tag).localname == 'commentRangeEnd' and sib.get(W + 'id') == cid:
                break
            for t in sib.iter(W + 't'):
                if t.text: texts.append(t.text)
        print(f'ID={cid} 标注文字：{"".join(texts)}')

if __name__ == '__main__':
    locate(sys.argv[1] if len(sys.argv) > 1 else 'output_with_comments.docx')
