# -*- coding: utf-8 -*-
"""
make_bilingual.py —— 便捷命令行入口（封装抽取/渲染/校验三步）

把分散的两个脚本收到一个 CLI 下，方便手动使用：

    # 1) 把 PDF 抽成文本（供你通读、切句、写 spec）
    python make_bilingual.py extract 真题.pdf 真题.txt

    # 2) 校验 spec.json 是否合法、统计各类块数量
    python make_bilingual.py check spec.json

    # 3) 把 spec.json 渲染成排版好的中英对照 Word
    python make_bilingual.py render spec.json 输出.docx

说明：从"两个 PDF"到"成品 docx"中间，有一步是**逐句翻译并组装成 spec.json**——
这一步需要语言模型的理解力（推荐用 Claude Code 安装本仓库的 skill 后直接对话完成，
见 README）。本 CLI 负责前后两端确定性的、不需要模型的部分。
"""
import sys, os, json

HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.join(HERE, 'skill', 'scripts')
sys.path.insert(0, SCRIPTS)


def cmd_extract(args):
    from extract_pdf_text import extract
    if len(args) != 2:
        print('用法: python make_bilingual.py extract 输入.pdf 输出.txt'); return 1
    extract(args[0], args[1]); return 0


def cmd_render(args):
    from render_bilingual_docx import build
    if len(args) != 2:
        print('用法: python make_bilingual.py render spec.json 输出.docx'); return 1
    with open(args[0], encoding='utf-8') as f:
        spec = json.load(f)
    build(spec, args[1]); return 0


def cmd_check(args):
    if len(args) != 1:
        print('用法: python make_bilingual.py check spec.json'); return 1
    with open(args[0], encoding='utf-8') as f:
        spec = json.load(f)
    blocks = spec.get('blocks', [])
    from collections import Counter
    c = Counter(b.get('type') for b in blocks)
    print('JSON 合法 ✓  标题:', spec.get('title'))
    print('总块数:', len(blocks))
    for t, n in sorted(c.items()):
        print('  %-12s %d' % (t, n))
    cloze = c.get('cloze_opts', 0)
    if cloze and cloze != 20:
        print('  ⚠ 完形选项块为 %d（通常应为 20）' % cloze)
    return 0


def main():
    if len(sys.argv) < 2 or sys.argv[1] in ('-h', '--help', 'help'):
        print(__doc__); return 0
    cmd, args = sys.argv[1], sys.argv[2:]
    table = {'extract': cmd_extract, 'render': cmd_render, 'check': cmd_check}
    if cmd not in table:
        print('未知命令: %s（可用: extract / render / check）' % cmd); return 1
    return table[cmd](args)


if __name__ == '__main__':
    sys.exit(main())
