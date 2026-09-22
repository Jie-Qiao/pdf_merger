#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""pdf_merger_cli.py —— 「全能文档转PDF合并器」的命令行入口

给这个小工具加一个命令行接口，方便脚本、批处理以及 AI Agent 调用。
转换与合并的核心逻辑与 GUI 版**完全共用**（都来自 pdf_merger_app），
所以同样的输入在两种入口下产出的 PDF 完全一致。

输入方式（可组合，前三种互斥，-i 可以叠加在任意一种之上）
  --list-file FILE   读取「复制列表」导出的 JSON 文件
  --list JSON        直接传 JSON 字符串
  --stdin            从标准输入读 JSON（或每行一项的纯文本）
  -i/--input VALUE   一个输入项，可重复。VALUE 既可以是纯文件路径，
                     也可以是列表项文本，例如：
                         -i D:\\a.docx
                         -i "[2x] [单面] D:\\b.pdf"
                         -i "[空白页] A4纵向"
                         -i "[分隔符] 切分 名为 附件二"

常用选项
  -o/--out-dir DIR   输出目录（默认当前目录，不存在会自动创建）
  -n/--name NAME     输出文件名（不含 .pdf）；多段时可用 # 作序号占位符
  --no-auto-pad      关闭「自动分叠」，即不插入空白页
  --json             以 JSON 输出结果，便于程序解析
  --dry-run          只解析并列出计划，不做实际的转换与合并
  -q/--quiet         不打印进度信息

退出码
  0  全部成功
  1  用法或输入有误（没给输入、目录建不出来等）
  2  部分成功（有文件被跳过或某一段失败，但至少产出了一个 PDF）
  3  失败（没有产出任何 PDF）

示例
  # 几个文件按默认双面合并成一个 PDF
  python pdf_merger_cli.py -i 试卷.docx -i 答案.pdf -o D:\\out

  # 其中一个文件要单面打印（每页独占一张纸，背面留白）
  python pdf_merger_cli.py -i 试卷.pdf -i "[1x] [单面] 答案.pdf" -o D:\\out

  # 复用 GUI 里「复制列表」保存下来的配置
  python pdf_merger_cli.py --list-file list.json -o D:\\out

  # 机器可读输出，便于 Agent 解析
  python pdf_merger_cli.py --list-file list.json -o D:\\out --json

  # 先看看会怎么处理，不实际生成
  python pdf_merger_cli.py --list-file list.json --dry-run
"""

import argparse
import io
import json
import os
import shutil
import sys
import tempfile

# 允许直接 `python pdf_merger_cli.py` 运行：把脚本所在目录加入模块搜索路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    from pdf_merger_app import (        # noqa: E402
        BLANK_PREFIX, SEP_PREFIX, FILE_ITEM_RE,
        SIDE_DUPLEX, SIDE_LABELS,
        build_segments, build_tasks_from_lines, format_file_item,
        load_list_json, write_outputs,
    )
except ImportError as _exc:
    sys.stderr.write(
        "无法导入核心模块 pdf_merger_app：%s\n\n"
        "它依赖以下库，请先安装：\n"
        "    pip install pypdf PyQt5 Pillow pywin32\n"
        "（本工具用 COM 调用本机 Office / WPS 做格式转换，因此只能在 Windows 上运行）\n"
        % _exc)
    sys.exit(1)


# ---------------------------------------------------------------- 辅助

def _describe_task(task):
    """把一个任务转成便于阅读/序列化的 dict（用于 --dry-run 与 --json）"""
    kind = task[0]
    if kind == 'file':
        return {
            'type': 'file',
            'path': task[1],
            'copies': task[2],
            'side': (task[3] if len(task) > 3 else None) or SIDE_DUPLEX,
            # dry-run 不做存在性校验，这里如实标出来，方便排查路径写错
            'exists': os.path.isfile(task[1]),
        }
    if kind == 'blank':
        return {'type': 'blank', 'width_cm': task[1], 'height_cm': task[2]}
    if kind == 'sep':
        return {
            'type': 'separator',
            'mode': task[1],
            'width_cm': task[2],
            'height_cm': task[3],
            'name': task[4],
            'side': task[5],
        }
    return {'type': str(kind)}


def _normalize_input(raw, default_side=SIDE_DUPLEX):
    """把 -i 的值规范成一条标准列表项文本。

    已经是列表项（文件项 / 空白页 / 分隔符）就原样返回；
    否则当成纯文件路径，补上「1 份 + 默认面向」的标记。
    """
    text = (raw or '').strip()
    if not text:
        return ''
    if text.startswith(SEP_PREFIX) or text.startswith(BLANK_PREFIX):
        return text
    if FILE_ITEM_RE.match(text):
        return text
    # 纯路径：统一成和界面一致的显示格式，这样份数/面向语义完全一样
    return format_file_item(1, text, None, default_side=default_side)


# --example 输出的模板：故意覆盖了全部列表项语法，照抄改成自己的路径即可。
# 路径用正斜杠示范 —— JSON 里写反斜杠必须双写，用正斜杠更省事。
_EXAMPLE_JSON = r'''{
  "kind": "pdf_merger_list",
  "version": 1,
  "items": [
    "D:/docs/试卷.docx",
    "[2x] [双面] D:/docs/答案.pdf",
    "[1x] [单面] D:/docs/答题卡.pdf",
    "[空白页] A4纵向",
    "[分隔符] 切分 名为 附件二",
    "D:/docs/附件二.pdf"
  ],
  "output_name": "合并结果"
}
'''


def _make_progress(quiet):
    """返回一个 progress 回调；quiet 时什么都不做。

    进度写到 stderr，这样 --json 模式下 stdout 依然只有干净的 JSON。
    """
    if quiet:
        return lambda _text: None

    def _progress(text):
        sys.stderr.write(text + '\n')
        sys.stderr.flush()

    return _progress


def _emit(payload, as_json, human_lines, stream=None):
    """按模式输出结果：JSON 或人类可读文本"""
    stream = stream or sys.stdout
    if as_json:
        stream.write(json.dumps(payload, ensure_ascii=False, indent=2) + '\n')
    else:
        for line in human_lines:
            stream.write(line + '\n')
    stream.flush()


def _fail(code, message, as_json=False, extra=None):
    """统一的错误出口：打印错误并返回退出码。

    --json 模式把 JSON 写到 **stdout**（调用方只解析一个流就够了）；
    普通模式把提示写到 stderr，不污染正常输出。
    """
    payload = {'ok': False, 'error': message, 'exit_code': code}
    if extra:
        payload.update(extra)
    if as_json:
        _emit(payload, True, [], stream=sys.stdout)
    else:
        _emit(payload, False, ['错误：' + message], stream=sys.stderr)
    return code


# ---------------------------------------------------------------- 参数

def build_parser():
    parser = argparse.ArgumentParser(
        prog='pdf_merger_cli',
        description='把 Word / Excel / PPT / 图片 / OFD / PDF 转成 PDF 并合并（命令行版）',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            '示例：\n'
            '  python pdf_merger_cli.py -i 试卷.docx -i 答案.pdf -o D:\\out\n'
            '  python pdf_merger_cli.py -i 试卷.pdf -i "[1x] [单面] 答案.pdf" -o D:\\out\n'
            '  python pdf_merger_cli.py --list-file list.json -o D:\\out --json\n'
            '  python pdf_merger_cli.py --list-file list.json --dry-run\n'
            '\n'
            '列表项语法（-i 的取值，以及 JSON 里 items 的每一项）：\n'
            '  <路径>                          1 份、双面（默认）\n'
            '  [2x] <路径>                     重复 2 份\n'
            '  [1x] [单面] <路径>               单面：每页独占一张纸、背面留白\n'
            '  [1x] [双面] <路径>               双面\n'
            '  [空白页] A4纵向                  插入一张空白页（也可写 21x29.7）\n'
            '  [分隔符] 切分 名为 附件二          在此处拆成另一个 PDF 并命名为「附件二」\n'
            '  [分隔符] 插入分隔纸 21x29.7cm     插入一页纸作为物理隔断\n'
            '\n'
            'JSON 列表格式（用于 --list / --list-file / --stdin）：\n'
            '  最简形式 —— 一个字符串数组，每项是路径或上面那种列表项文本：\n'
            '      ["a.docx", "[2x] b.pdf", "[1x] [单面] c.pdf"]\n'
            '\n'
            '  完整形式 —— 对象，可顺便指定输出文件名：\n'
            '      {\n'
            '        "kind": "pdf_merger_list",\n'
            '        "version": 1,\n'
            '        "items": ["a.docx", "[空白页] A4纵向", "[1x] [单面] b.pdf"],\n'
            '        "output_name": "合并结果"\n'
            '      }\n'
            '      其中 kind / version / output_name 都可省略，只有 items 必需；\n'
            '      kind 若填写则必须等于 "pdf_merger_list"，否则报错；\n'
            '      output_name 等价于命令行 -n（命令行给的 -n 优先级更高）。\n'
            '\n'
            '  也接受「每行一项」的纯文本（方便从记事本直接粘贴）。\n'
            '\n'
            '  拿不准格式时，直接运行  python pdf_merger_cli.py --example\n'
            '  就会打印一份可直接复制修改的 JSON 模板。\n'
            '\n'
            '  注意：JSON 里的 Windows 路径，反斜杠要写成两个或改用正斜杠 ——\n'
            '      "D:\\\\a.pdf"  或  "D:/a.pdf"    （单个 \\ 在 JSON 里是非法转义）\n'
        ),
    )

    src = parser.add_mutually_exclusive_group()
    src.add_argument('--list-file', metavar='FILE',
                     help='从 JSON 列表文件读取（「复制列表」导出的格式）')
    src.add_argument('--list', dest='list_json', metavar='JSON',
                     help='直接传入 JSON 列表字符串')
    src.add_argument('--stdin', action='store_true',
                     help='从标准输入读取 JSON 或逐行文本')

    parser.add_argument('-i', '--input', action='append', default=[],
                        metavar='VALUE',
                        help='输入项，可重复。可为纯路径，也可为列表项文本，'
                             '如 "[2x] [单面] D:\\\\a.pdf"')
    parser.add_argument('-o', '--out-dir', default='.', metavar='DIR',
                        help='输出目录（默认当前目录，不存在会自动创建）')
    parser.add_argument('-n', '--name', default='', metavar='NAME',
                        help='输出文件名（不含 .pdf）；多段可用 # 作序号占位符')
    parser.add_argument('--no-auto-pad', dest='auto_pad', action='store_false',
                        default=True,
                        help='关闭「自动分叠」（不插入空白页）')
    parser.add_argument('--json', dest='as_json', action='store_true',
                        help='以 JSON 输出结果，便于程序解析')
    parser.add_argument('--dry-run', action='store_true',
                        help='只列出将要处理的内容，不实际转换/合并')
    parser.add_argument('-q', '--quiet', action='store_true',
                        help='不输出进度信息')
    parser.add_argument('-V', '--version', action='version',
                        version='pdf_merger_cli 1.0')
    parser.add_argument('--example', action='store_true',
                        help='打印一份 JSON 列表格式的示例（可直接复制修改）')
    return parser


# ---------------------------------------------------------------- 主流程

def main(argv=None):
    """CLI 入口；返回退出码（0/1/2/3），见文件头的说明。"""
    # 统一按 UTF-8 输出，保证中文在管道/重定向下也不会乱码
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding='utf-8', errors='replace')
        except Exception:
            pass

    parser = build_parser()
    args = parser.parse_args(argv)

    # --example：只打印一份 JSON 模板，供复制修改（输出纯 JSON，不带任何装饰）
    if args.example:
        sys.stdout.write(_EXAMPLE_JSON)
        sys.stdout.flush()
        return 0

    # ---- 1) 收集列表项 ----
    lines = []
    output_name = (args.name or '').strip()

    try:
        if args.list_file:
            if not os.path.isfile(args.list_file):
                return _fail(1, f"找不到列表文件：{args.list_file}", args.as_json)
            with io.open(args.list_file, encoding='utf-8') as f:
                got, oname = load_list_json(f.read())
            lines.extend(got)
            if oname and not output_name:
                output_name = oname
        elif args.list_json:
            got, oname = load_list_json(args.list_json)
            lines.extend(got)
            if oname and not output_name:
                output_name = oname
        elif args.stdin:
            got, oname = load_list_json(sys.stdin.read())
            lines.extend(got)
            if oname and not output_name:
                output_name = oname
    except ValueError as exc:
        return _fail(1, f"列表内容解析失败：{exc}", args.as_json)

    # -i 的值追加在列表之后（两者可以一起用）
    for raw in args.input:
        item = _normalize_input(raw)
        if item:
            lines.append(item)

    lines = [str(x) for x in lines if str(x).strip()]
    if not lines:
        return _fail(
            1,
            '没有输入内容。请用 -i 指定文件，或用 --list-file / --list / --stdin 提供列表。',
            args.as_json)

    # ---- 2) 解析任务与分段 ----
    # dry-run 不做文件存在性校验：它的用途正是「先看看写的内容对不对」，
    # 路径写错了也应该能列出来，而不是直接报错退出。
    tasks, warnings = build_tasks_from_lines(lines, check_exists=not args.dry_run)
    file_count = sum(1 for t in tasks if t[0] == 'file')
    if not file_count:
        return _fail(1, '解析后没有任何可处理的文件（可能路径都不存在或格式不支持）。',
                     args.as_json, extra={'warnings': warnings})

    segments = build_segments(tasks)
    out_dir = os.path.abspath(args.out_dir)

    # ---- 3) dry-run：只报告计划 ----
    if args.dry_run:
        plan = {
            'ok': True,
            'dry_run': True,
            'out_dir': out_dir,
            'output_name': output_name,
            'segment_count': len(segments),
            'file_count': file_count,
            'segments': [
                {
                    'index': i,
                    'name': name,
                    'side': side or SIDE_DUPLEX,
                    'items': [_describe_task(t) for t in seg],
                }
                for i, (name, seg, side) in enumerate(segments, start=1)
            ],
            'warnings': warnings,
        }
        human = [
            f"输出目录：{out_dir}",
            f"共 {len(segments)} 段 / {file_count} 个文件",
        ]
        for seg in plan['segments']:
            tag = SIDE_LABELS.get(seg['side'], seg['side'])
            human.append(f"  第 {seg['index']} 段（{tag}）"
                         + (f" 名称={seg['name']}" if seg['name'] else ''))
            for item in seg['items']:
                if item['type'] == 'file':
                    side_tag = SIDE_LABELS.get(item['side'], item['side'])
                    miss = '' if item.get('exists') else '   ← 文件不存在'
                    human.append(f"      [{item['copies']}x] [{side_tag}] "
                                 f"{item['path']}{miss}")
                elif item['type'] == 'blank':
                    human.append(f"      [空白页] {item['width_cm']:g}x{item['height_cm']:g}cm")
                else:
                    human.append(f"      [分隔符] {item.get('name') or ''}")
        for w in warnings:
            human.append(f"  警告：{w}")
        _emit(plan, args.as_json, human)
        return 0

    # ---- 4) 实际执行 ----
    try:
        os.makedirs(out_dir, exist_ok=True)
    except Exception as exc:
        return _fail(1, f"无法创建输出目录：{exc}", args.as_json)

    progress = _make_progress(args.quiet)
    temp_dir = tempfile.mkdtemp(prefix='pdf_merger_cli_')
    pdf_cache = {}
    failed = []
    try:
        results, errors = write_outputs(
            out_dir, segments, pdf_cache, failed, temp_dir,
            user_name=output_name, auto_pad=args.auto_pad,
            default_side=SIDE_DUPLEX, progress=progress)
    except Exception as exc:
        return _fail(3, f"处理失败：{exc}", args.as_json,
                     extra={'warnings': warnings, 'failed': failed})
    finally:
        # 转换用的中间文件全部收在这个临时目录里，用完即弃
        shutil.rmtree(temp_dir, ignore_errors=True)

    # ---- 5) 汇总输出 ----
    total_pages = sum(p for _, p, _s, _pad in results)
    payload = {
        'ok': bool(results) and not errors and not failed,
        'out_dir': out_dir,
        'files': [
            {
                'name': name,
                'path': os.path.join(out_dir, name),
                'pages': pages,
                'side': side or SIDE_DUPLEX,
                'side_label': SIDE_LABELS.get(side or SIDE_DUPLEX, ''),
                'blank_pages': pad,
            }
            for name, pages, side, pad in results
        ],
        'total_pages': total_pages,
        'warnings': warnings,
        'failed': failed,
        'errors': [{'name': n, 'error': str(e)} for n, e, _p, _t in errors],
    }

    human = [f"输出目录：{out_dir}"]
    for f in payload['files']:
        extra = f"，插 {f['blank_pages']} 页空白" if f['blank_pages'] else ''
        human.append(f"  {f['name']}（{f['pages']} 页，{f['side_label']}{extra}）")
    human.append(f"合计 {len(payload['files'])} 个文件 / {total_pages} 页")
    for w in warnings:
        human.append(f"  警告：{w}")
    for f in failed:
        human.append(f"  跳过：{f}")
    for e in payload['errors']:
        human.append(f"  失败：{e['name']} -> {e['error']}")

    _emit(payload, args.as_json, human)

    # ---- 6) 退出码 ----
    if not results:
        return 3
    if warnings or failed or errors:
        return 2
    return 0


if __name__ == '__main__':
    sys.exit(main())
