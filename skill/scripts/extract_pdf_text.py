# -*- coding: utf-8 -*-
"""
把 PDF 逐页文本抽取为 UTF-8 文本文件，供 Claude 通读、切句、对齐。

为什么要单独抽取：
- 直接在终端打印中文常因 Windows 代码页(GBK)显示成乱码，
  写到 UTF-8 文件里再用 Read 工具看就正常。
- 英文一般可正常抽取；扫描/图片型 PDF 可能出现 OCR 讹误（人名、专有词），
  抽取后要人工核对，必要时在文档里加注说明。

用法：
    python extract_pdf_text.py 试题.pdf exam.txt
    python extract_pdf_text.py 答案.pdf answer.txt
"""
import sys, io
import fitz  # PyMuPDF


def extract(pdf_path, out_path):
    d = fitz.open(pdf_path)
    with io.open(out_path, 'w', encoding='utf-8') as w:
        for i, page in enumerate(d):
            w.write('\n\n===== PAGE %d =====\n' % (i + 1))
            w.write(page.get_text())
    print('extracted %d pages -> %s' % (d.page_count, out_path))


if __name__ == '__main__':
    if len(sys.argv) < 3:
        print('usage: python extract_pdf_text.py in.pdf out.txt'); sys.exit(1)
    extract(sys.argv[1], sys.argv[2])
