"""输出批注相关 XML 原始内容（调试用）"""
import zipfile, sys

FILES = ['word/comments.xml', 'word/commentsExtended.xml',
         'word/commentsIds.xml', 'word/commentsExtensible.xml', 'word/people.xml']

def inspect(docx_path):
    with zipfile.ZipFile(docx_path, 'r') as z:
        for f in FILES:
            try:
                data = z.read(f).decode('utf-8')
                print(f'\n{"="*60}\n📄 {f}\n{"="*60}')
                print(data[:3000])
                if len(data) > 3000: print(f'\n... ({len(data)} chars total)')
            except KeyError:
                print(f'\n⚠️ {f} — 不存在')

if __name__ == '__main__':
    inspect(sys.argv[1] if len(sys.argv) > 1 else 'output_with_comments.docx')
