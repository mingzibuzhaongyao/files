from docx import Document
import re

# 标题样式映射
HEADING_MAP = {
    'Heading 1': 'h2',
    'Heading 2': 'h3',
    'Heading 3': 'h4',
    'Heading 4': 'h5',
}

# 回退用的编号正则：一级(1. / 1．)，二级(1) / 1）)
RE_TOP = re.compile(r'^\d+\s*[\.．]\s*')
RE_SUB = re.compile(r'^\d+\s*[)\uff09]\s*')

def get_ilvl_from_numpr(para):
    """优先用 Word 的自动编号层级；拿不到就返回 None。"""
    try:
        pPr = para._p.pPr
        if pPr is None or pPr.numPr is None:
            return None
        ilvl = pPr.numPr.ilvl
        # ilvl 可能不存在，按 Word 习惯默认 0
        if ilvl is None or ilvl.val is None:
            return 0
        return int(ilvl.val)
    except Exception:
        return None

def docx_to_html(input_path, output_path):
    doc = Document(input_path)
    html = []
    ul_stack = []  # 用来管理当前 <ul> 深度
  
    def open_to(level):
        # 把 <ul> 开到指定层级
        while len(ul_stack) <= level:
            html.append('<ul>')
            ul_stack.append(True)

    def close_to(level):
        # 关闭到指定层级（不含该层级）
        while len(ul_stack) > level:
            html.append('</ul>')
            ul_stack.pop()

    for para in doc.paragraphs:
        print(para.text, para._p.pPr.numPr if para._p.pPr is not None else None)
        raw = para.text.strip()
        if not raw:
            continue

        # 1) 先判定是否为自动编号
        ilvl = get_ilvl_from_numpr(para)
        content = raw

        # 2) 若不是自动编号，再用正则兜底（两级）
        if ilvl is None:
            if RE_TOP.match(raw):
                ilvl = 0
                content = RE_TOP.sub('', raw)
            elif RE_SUB.match(raw):
                ilvl = 1
                content = RE_SUB.sub('', raw)

        if ilvl is not None:
            # 列表项
            if ilvl >= len(ul_stack):
                open_to(ilvl)
            elif ilvl < len(ul_stack) - 1:
                close_to(ilvl + 1)
            html.append(f'<li>{content}</li>')
            continue

        # 普通段落/标题 —— 先把所有列表闭合
        close_to(0)
        style = para.style.name if para.style is not None else ''
        if style in HEADING_MAP:
            tag = HEADING_MAP[style]
            html.append(f'<{tag}>{raw}</{tag}>')
        else:
            html.append(f'<p>{raw}</p>')

    # 文件末尾收尾
    close_to(0)

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(html))
    print(f'HTML saved to {output_path}')
    

if __name__ == '__main__':
    input_file = 'temp.docx'
    output_file = 'output.html'
    docx_to_html(input_file, output_file)
