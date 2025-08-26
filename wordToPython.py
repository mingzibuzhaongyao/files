from docx import Document

# Heading 样式到 HTML 的映射
heading_map = {
    'Heading 1': 'h2',
    'Heading 2': 'h3',
    'Heading 3': 'h4',
    'Heading 4': 'h5',
}

def docx_to_html(input_path, output_path):
    doc = Document(input_path)
    html_lines = []
    ul_open = False

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        style_name = para.style.name

        # 检查段落是否是自动编号
        numPr = getattr(getattr(para._p, "pPr", None), "numPr", None)
        if numPr is not None:
            # 自动编号段落
            if not ul_open:
                html_lines.append('<ul>')
                ul_open = True
            html_lines.append(f'<li>{text}</li>')
        else:
            # 普通段落或标题
            if ul_open:
                html_lines.append('</ul>')
                ul_open = False

            if style_name in heading_map:
                tag = heading_map[style_name]
                html_lines.append(f'<{tag}>{text}</{tag}>')
            else:
                html_lines.append(f'<p>{text}</p>')

    if ul_open:
        html_lines.append('</ul>')

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(html_lines))
    print(f'HTML saved to {output_path}')

if __name__ == '__main__':
    input_file = 'temp.docx'  # Word 文件路径
    output_file = 'output.html'  # 输出 HTML 文件路径
    docx_to_html(input_file, output_file)
