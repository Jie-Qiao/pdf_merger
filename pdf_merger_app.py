import sys
import os
import json
import tempfile
import base64
import re
import win32com.client
from PIL import Image
from pypdf import PdfWriter, PdfReader
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
                             QListWidget, QLabel, QMessageBox, QAbstractItemView, QLineEdit,
                             QFileDialog, QCheckBox, QInputDialog, QMenu)
from PyQt5.QtCore import Qt, QStandardPaths
# 根据官方示例引入 OFD
try:
    from easyofd.ofd import OFD
    HAS_EASYOFD = True
except ImportError:
    HAS_EASYOFD = False

# 全局定义支持的扩展名，方便拖拽文件夹时进行过滤
SUPPORTED_EXTS = ['.docx', '.doc', '.xlsx', '.xls', '.ppt', '.pptx', 
                  '.jpg', '.jpeg', '.png', '.bmp', '.ofd', '.pdf']

# ============ 新增功能相关常量 ============
BLANK_PREFIX = '[空白页] '                 # 空白页列表项的文本前缀（用于识别与解析）
DEFAULT_BLANK_KEY = 'A4纵向'               # 默认插入的空白页规格

SEP_PREFIX = '[分隔符]'                    # 分隔符列表项的文本前缀
SEP_SPLIT = 'SPLIT'                        # 模式一：在此处切分成多个独立 PDF（分隔符本身不产生页面）
SEP_SHEET = 'SHEET'                        # 模式二：插入一页“分隔纸”作为物理隔断
SEP_DEFAULT_MODE = SEP_SPLIT

# 分隔符模式说明：内部值 -> (中文名, 说明)
SEP_MODES = [
    (SEP_SPLIT, '切分成独立PDF', '分隔符本身不产生页面，只是把它上下两段拆成两个文件'),
    (SEP_SHEET, '插入分隔纸', '额外插入一页纸（默认 A4 纵向）作为两个 PDF 之间的物理隔断'),
]

# ============ 单面/双面打印相关常量 ============
# 打印面向标记（挂在文件/分隔符上，用于「自动分叠」与输出提示）
# 注意：没有「默认」这个取值 —— 未标记一律按双面处理，见 PDFMergerApp._default_side
SIDE_DUPLEX = 'DUPLEX'    # 双面
SIDE_SIMPLEX = 'SIMPLEX'  # 单面

SIDE_LABELS = {
    SIDE_DUPLEX: '双面',
    SIDE_SIMPLEX: '单面',
}

# 纸张尺寸表：名称 -> (宽, 高)，单位厘米
BLANK_PRESETS = {
    'A3纵向':     (29.7, 42.0),
    'A3横向':     (42.0, 29.7),
    'A4纵向':     (21.0, 29.7),
    'A4横向':     (29.7, 21.0),
    'A5纵向':     (14.8, 21.0),
    'A5横向':     (21.0, 14.8),
    'Letter纵向': (21.59, 27.94),
    'Letter横向': (27.94, 21.59),
}

CM_TO_PT = 28.3465   # 1 厘米 = 28.3465 点


def cm_to_pt(v):
    """厘米转 PDF 点"""
    return float(v) * CM_TO_PT


def find_blank_preset(name):
    """按名称模糊匹配纸张规格，返回 (显示名, 宽cm, 高cm)；匹配不到返回 None"""
    if not name:
        return None
    key_target = name.replace(' ', '').lower()
    for k, (w, h) in BLANK_PRESETS.items():
        if k.lower() == key_target:
            return k, w, h
    for k, (w, h) in BLANK_PRESETS.items():
        if key_target in k.lower() or k.lower() in key_target:
            return k, w, h
    return None


def parse_blank_spec(text):
    """解析空白页规格文本，返回 (宽cm, 高cm)。
    支持三种写法：
      1) 预置名称: A4纵向 / A5横向 / Letter纵向 ...
      2) 宽x高:    21x29.7  或  21*29.7
      3) 仅数字:   21  (同时作为宽和高，即正方形)
    """
    preset = find_blank_preset(text)
    if preset:
        return preset[1], preset[2]

    nums = re.findall(r'(\d+(?:\.\d+)?)', text or '')
    if len(nums) >= 2:
        return float(nums[0]), float(nums[1])
    if len(nums) == 1:
        v = float(nums[0])
        return v, v
    return None


SEP_RE = re.compile(
    r'^\s*' + re.escape(SEP_PREFIX) + r'\s*(.*)$', re.IGNORECASE)

# 模式别名（全部小写，比对时大小写不敏感）
SPLIT_ALIASES = ('切分', '拆分', '分件', '独立', 'split', '新文件', '分开')
SHEET_ALIASES = ('分隔纸', '插页', '隔页', 'sheet', '插入分隔纸', '分隔页')
NAME_MARKERS = ('名为', '名称', '文件名', 'named', 'name')

# 打印面向别名
SIDE_ALIASES_DUPLEX = ('双面', '双面打印', 'duplex', '双')
SIDE_ALIASES_SIMPLEX = ('单面', '单面打印', 'simplex', '单')


def parse_sep_spec(text):
    """解析分隔符规格文本。

    返回 (mode, 名称, 宽cm, 高cm, 文件名)：
      - mode = 'SPLIT' : 只切分，不产生页面，尺寸为 None
      - mode = 'SHEET' : 插入一页分隔纸，宽高为 cm
      - 第 2 个返回值只是模式的中文说明（“切分”/“分隔纸”），不是文件名
      - 第 5 个返回值才是用户指定的输出文件名，未指定时为 None

    支持的写法（大小写不敏感，方括号可省略，各段可用空格/逗号分隔）：
      '分隔符'                          -> SPLIT
      '分隔符 切分'                      -> SPLIT
      '分隔符 插入分隔纸'                 -> SHEET，A4 纵向
      '分隔符 插入分隔纸 A4横向'           -> SHEET，A4 横向
      '分隔符 插入分隔纸 A4横向 附件二'     -> SHEET，A4 横向，命名“附件二”
      '分隔符 名为 附件二'                -> SPLIT，命名“附件二”
    """
    raw = (text or '').strip()
    m = SEP_RE.match(raw)
    inner = m.group(1).strip() if m else raw

    # 兜底：ensure_ascii=False 打包/复制粘贴可能带出 \uXXXX 字面量，先还原
    if '\\u' in inner:
        try:
            inner = inner.encode('utf-8').decode('unicode_escape')
        except Exception:
            pass

    if not inner:
        return SEP_DEFAULT_MODE, None, None, None, None

    # 1) 剥离显式的“名为 xxx”后缀，xxx 一律视为文件名
    doc_name = None
    for marker in NAME_MARKERS:
        idx = inner.lower().find(marker)
        if idx >= 0:
            doc_name = inner[idx + len(marker):].strip(' \t:：') or None
            inner = inner[:idx].strip()
            break

    # 2) 剩余部分按分隔符切词
    tokens = [t for t in re.split(r'[\s,，、|]+', inner) if t]

    mode = None
    size_tokens = []
    leftovers = []

    for tok in tokens:
        low = tok.lower()
        if low in SPLIT_ALIASES:
            mode = SEP_SPLIT
        elif low in SHEET_ALIASES:
            mode = SEP_SHEET
        elif parse_blank_spec(tok):
            size_tokens.append(tok)
        else:
            leftovers.append(tok)

    # 3) 尺寸：优先用识别出的规格词，否则把剩余词合起来试一次
    width = height = None
    if size_tokens:
        size = parse_blank_spec(' '.join(size_tokens))
        if size:
            width, height = size
            mode = SEP_SHEET
    if width is None and leftovers:
        joined = ' '.join(leftovers)
        size = parse_blank_spec(joined)
        if size and not find_blank_preset(joined) is None:
            width, height = size
            mode = SEP_SHEET
            leftovers = []

    # 4) 剩下的无法识别的词，当作文件名（先摘出单/双面关键字）
    side = None
    rest = []
    for tok in leftovers:
        low = tok.lower()
        if low in SIDE_ALIASES_DUPLEX:
            side = SIDE_DUPLEX
        elif low in SIDE_ALIASES_SIMPLEX:
            side = SIDE_SIMPLEX
        else:
            rest.append(tok)
    leftovers = rest

    if leftovers and doc_name is None:
        doc_name = ' '.join(leftovers)

    if mode is None:
        mode = SEP_DEFAULT_MODE

    if mode == SEP_SHEET and width is None:
        width, height = BLANK_PRESETS[DEFAULT_BLANK_KEY]

    label = '切分' if mode == SEP_SPLIT else '分隔纸'
    return mode, label, width, height, doc_name, side


def format_sep_text(mode, doc_name=None, width=None, height=None, side=None):
    """把分隔符配置格式化成列表项文本。

    注意：不要写成 `format_sep_text(*parse_sep_spec(x))`。
    parse_sep_spec 的第二个返回值是它内部识别出的**模式关键字**（如「切分」），
    不是名称；直接展开会把它塞进 doc_name 位置，产出「名为 切分」这种脏数据。
    正确用法是按位解包：mode, _, w, h, name, side = parse_sep_spec(x)
    """
    parts = [SEP_PREFIX]
    if mode == SEP_SHEET:
        parts.append('插入分隔纸')
        if width and height:
            parts.append(f'{width:g}x{height:g}cm')
    else:
        parts.append('切分')
    if side:
        parts.append(SIDE_LABELS.get(side, ''))
    if doc_name:
        parts.append(f'名为 {doc_name}')
    return ' '.join(parts)


def parse_list_item(text):
    """把列表项文本解析成 (类型, 值)。
    文件项: '[2x] D:\\a.pdf'      -> ('file', 'D:\\a.pdf')
    空白页: '[空白页] A4纵向 (14.8x21cm)' -> ('blank', 'A4纵向 (14.8x21cm)')
    分隔符: '[分隔符] 切分'        -> ('sep', '切分')
    """
    if text.startswith(SEP_PREFIX):
        return 'sep', text[len(SEP_PREFIX):].strip()
    if text.startswith(BLANK_PREFIX):
        return 'blank', text[len(BLANK_PREFIX):].strip()
    m = FILE_ITEM_RE.match(text)
    if m:
        return 'file', m.group(2).strip() if m.group(2) else text.strip()
    return 'file', text.strip()


def sanitize_filename(name):
    """把用户输入的名字清洗成合法文件名（去掉非法字符与首尾空白点）"""
    if not name:
        return ''
    name = re.sub(r'[\\/:*?"<>|\r\n\t]', '_', name).strip().strip('.')
    return name[:120]


# 文件项形如： [2x] [双面] D:\a.pdf  —— 份数必填，面向标记可选
# 面向标记的「=」前缀是历史格式（早期的「跟随默认面」标记），现在只用于兼容旧文本
FILE_ITEM_RE = re.compile(r'^\[(\d+)\s*x\]\s*(?:\[(=?)(双面|单面)\]\s*)?(.+)$',
                          re.IGNORECASE)


def parse_file_item(text):
    """解析文件列表项，返回 (份数, side 或 None, 路径)。

    支持的写法：
      '[1x] D:\\a.pdf'         -> 未标记，side=None（按默认的双面处理）
      '[2x] [双面] D:\\a.pdf'   -> side=SIDE_DUPLEX
      '[2x] [单面] D:\\a.pdf'   -> side=SIDE_SIMPLEX
      '[2x] [=双面] D:\\a.pdf'  -> side=None

    带「=」的是**历史格式**：早期版本用它表示「跟随全局默认面」。
    现在默认面已固定为双面、显示上不再生成「=」，但旧文本（比如以前
    复制出去的列表 JSON）仍然要能正确解析成「未标记」。
    """
    t = (text or '').strip()
    # FILE_ITEM_RE 的分组：1=份数  2=「=」号（可选）  3=双面/单面  4=路径
    m = FILE_ITEM_RE.match(t)
    if m and m.group(4) is not None:
        copies = max(1, int(m.group(1)))
        explicit = (m.group(2) or '') != '='   # 带等号的是历史默认标记，不算显式指定
        mark = m.group(3)
        path = m.group(4).strip()
        side = None
        if explicit and mark == '双面':
            side = SIDE_DUPLEX
        elif explicit and mark == '单面':
            side = SIDE_SIMPLEX
        return copies, side, path
    # 不带份数的裸路径
    side = None
    m2 = re.match(r'^(?:\[(=?)(双面|单面)\]\s*)?(.+)$', t)
    if m2:
        explicit = m2.group(1) != '='
        if explicit and m2.group(2) == '双面':
            side = SIDE_DUPLEX
        elif explicit and m2.group(2) == '单面':
            side = SIDE_SIMPLEX
        return 1, side, m2.group(3).strip()
    return 1, None, t


def format_file_item(copies, path, side=None, display=True, default_side=None):
    """把文件项格式化成列表文本。

    display=True（默认，供界面显示）时，未标记面向的项也会补上 [双面]，
    方便一眼看清每个文件按什么方式打印。
    display=False 用于内部临时构造，需要时再显式带标记。
    """
    parts = [f"[{copies}x]"]
    if side:
        parts.append(f"[{SIDE_LABELS[side]}]")
    elif display:
        parts.append(f"[{side_display_token(None, default_side)}]")
    parts.append(path)
    return ' '.join(parts)


# 自动编号的占位符：写 # 的地方会被替换成递增序号
SEQ_PLACEHOLDER = '#'


# ============ 面向的显式标记与「列表显示」 ============
def side_display_token(side, default_side=None):
    """把内部面向值转成列表里显示的短标记（双面 / 单面）。

    未显式标记的文件按默认面（双面）显示，这样列表里每一项都带标记，
    一眼就能看清每个文件会怎么打印。

    历史说明：早期版本在未标记时显示成 [=双面]，用「=」表示「跟随全局默认」。
    后来把全局默认固定成了双面（见 _default_side），「=」就没意义了，
    这里统一成不带前缀的 [双面]。旧文本里的 [=双面] 仍然能被正确解析，
    见 parse_file_display。
    """
    if side is None:
        return SIDE_LABELS[default_side or SIDE_DUPLEX]
    return SIDE_LABELS.get(side, SIDE_LABELS[SIDE_DUPLEX])


def parse_file_display(text):
    """解析显示文本，返回 (份数, side 或 None, 路径)。

    带「=」的标记（如 [=双面]）是历史格式，解析回 None（等价于未标记）。
    保留这段是为了让早期用「复制列表」导出的文本仍能正确恢复。
    """
    copies, side, path = parse_file_item(text)
    m = FILE_ITEM_RE.match((text or '').strip())
    if m and m.group(4) is not None and (m.group(2) or '') == '=':
        return copies, None, path
    return copies, side, path


def resolve_output_name(user_name, index, total):
    """把用户填写的输出文件名模板解析成实际文件名（不含扩展名）。

    user_name 为空        -> None，交给调用方用默认名（合并输出_打印预览 / 分卷NN）
    user_name 含 # 占位符  -> 把 # 替换成序号；序号按最终段数补零对齐
    user_name 不含占位符：
        - 只有一段      -> 原样使用
        - 有多段        -> 自动在末尾补序号，避免几段互相覆盖

    返回清洗后的文件名（不含 .pdf），或 None 表示用默认名。
    """
    raw = (user_name or '').strip()
    if not raw:
        return None

    # 去掉用户可能顺手打上的 .pdf 后缀（统一由程序补）
    if raw.lower().endswith('.pdf'):
        raw = raw[:-4]

    if SEQ_PLACEHOLDER in raw:
        width = max(len(str(total)), 2)
        try:
            name = raw.replace(SEQ_PLACEHOLDER, f"{index:0{width}d}")
        except Exception:
            name = raw.replace(SEQ_PLACEHOLDER, str(index))
    elif total > 1:
        # 多段但没写占位符：补序号，避免覆盖
        name = f"{raw}_{index}"
    else:
        name = raw

    name = sanitize_filename(name)
    if not name:
        return None
    # 编号后仍可能撞名，交给调用方的去重逻辑处理
    return name


def plan_auto_pad(segment_page_counts, enabled=True):
    """旧接口，保留兼容：等价于 plan_pads(..., segment_sides=None)。

    新代码请直接用 plan_pads，它才认单/双面标记。
    """
    return plan_pads(segment_page_counts, None, enabled)


# ============ 剪切板（复制 / 粘贴整份列表） ============
CLIP_KIND = 'pdf_merger_list'   # 自建格式的标识，防止误粘别人的 JSON
CLIP_VERSION = 1


def plan_pads(segment_page_counts, segment_sides=None, enabled=True):
    """计算「自动分叠」要给每段补几页空白 —— 面向感知版。

    为什么需要它：最初只按「页数是不是奇数」补页，压根没看这一段的
    单/双面标记。结果标了「单面」的文档常常一页空白都不补，
    用户看着标记以为会插页，实际什么也没发生。

    统一前提：**打印时全程只选「双面」**，靠补页达成单面效果。
    （单面/双面是打印驱动里的运行时选项，写不进 PDF 文件，只能这么绕。）
    一张纸 = 正反两个页位。

    两种段的规则不同：

      【双面段】内容正反面都印，只需让它独占整张纸、从正面开始
        · 占偶数页且从正面开始 -> 天然独占整张纸，不补
        · 占奇数页             -> 末尾补 1 页，否则下一段会印到它背面

      【单面段】每一页内容都只占一张纸的**正面**，背面必须空白
        · 做法是把整段「撑」成每页一张纸：页数 N -> 物理占 2N 个页位
        · 具体是「内容页、空白页、内容页、空白页…」交替排布，
          由 _write_segment 的 simplex 模式逐页插空白实现
        · 所以这里返回的 pad 表示「额外要补的空白页数」，
          单面段为 N（偶数页的段也照样补），补完 2N 必为偶数

    参数：
      segment_page_counts: [段1页数, 段2页数, ...]
      segment_sides:       [段1面向, ...]，元素为 SIDE_* 或 None；可为 None
      enabled:             是否启用（False 时全部补 0 页）

    返回 (pads, layout)：
      pads   每段要额外补的空白页数
      layout 推算表，含每段起止页码、占用纸张编号、是否独占整张纸
    """
    if segment_sides is None:
        segment_sides = [None] * len(segment_page_counts)

    pads = []
    layout = []
    cursor = 1  # 下一段的起始页码（1-based）

    for idx, pages in enumerate(segment_page_counts, start=1):
        side = segment_sides[idx - 1] if idx - 1 < len(segment_sides) else None
        is_simplex = (side == SIDE_SIMPLEX)

        if not enabled:
            pad = 0
        elif is_simplex:
            # 单面：每页都要独占一张纸，背面全空 -> 再补出等量的空白页
            pad = pages
        else:
            # 双面：奇数页补 1 页凑偶数，偶数页不动
            pad = 1 if pages % 2 == 1 else 0

        padded = pages + pad
        start, end = cursor, cursor + padded - 1
        # 双面打印，每张纸装 2 页，据此推算占用哪些纸张
        sheets = (start - 1) // 2 + 1, (end - 1) // 2 + 1
        layout.append({
            'index': idx,
            'pages': pages,
            'pad': pad,
            'padded': padded,
            'start': start,
            'end': end,
            'side': side,
            'simplex': is_simplex,
            'sheets': sheets,
            # 起始页为奇数 且 占偶数页 -> 独占整张纸，从正面开始
            'clean': (start % 2 == 1) and (padded % 2 == 0),
        })
        pads.append(pad)
        cursor = end + 1

    return pads, layout


def build_list_snapshot(lines, output_name=''):
    """把列表的每一行文本打包成可复制的 JSON 结构。"""
    return {
        'kind': CLIP_KIND,
        'version': CLIP_VERSION,
        'items': [str(x) for x in lines],
        'output_name': output_name or '',
    }


def dump_list_json(lines, output_name='', pretty=True):
    """导出成 JSON 文本（保存到剪切板）。"""
    snap = build_list_snapshot(lines, output_name)
    return json.dumps(snap, ensure_ascii=False,
                      indent=2 if pretty else None)


def load_list_json(text):
    """解析粘贴的文本，返回 (lines, output_name)。

    容错处理：
      · 标准 JSON（含 items 数组）      -> 直接采用
      · 只有 kind 对不上的 JSON         -> 抛 ValueError
      · 纯文本，每行一项                -> 直接当列表用（兼容从记事本粘回来）
    解析不出来会抛 ValueError，由调用方提示用户。
    """
    raw = (text or '').strip()
    if not raw:
        raise ValueError('剪切板是空的。')

    lines = None
    output_name = ''

    # 判断是不是本程序的 JSON：必须是对象，且带 kind/items 之类的特征。
    # 注意不能只看开头是不是 '[' —— 列表项本身就形如 "[1x] D:\a.pdf"，
    # 那是纯文本而不是 JSON 数组。
    looks_json = False
    if raw.startswith('{'):
        looks_json = True
    elif raw.startswith('['):
        # 只有整段是合法 JSON 数组（形如 ["...","..."]）才当 JSON
        try:
            json.loads(raw)
            looks_json = True
        except Exception:
            looks_json = False

    if looks_json:
        try:
            data = json.loads(raw)
        except Exception as exc:
            raise ValueError(f'JSON 解析失败：{exc}')

        if isinstance(data, list):
            lines = data
        elif isinstance(data, dict):
            kind = data.get('kind')
            if kind and kind != CLIP_KIND:
                raise ValueError(f'这不是本程序导出的列表（kind={kind}）。')
            if not isinstance(data.get('items'), list):
                raise ValueError('JSON 里缺少 items 数组。')
            lines = data['items']
            output_name = str(data.get('output_name') or '')
        else:
            raise ValueError('JSON 结构无法识别。')

        if lines and not isinstance(lines[0], str):
            raise ValueError('items 里应该是字符串。')

    if lines is None:
        # 兜底：按行文本处理（兼容从记事本 / 聊天窗口粘回来的情况）
        lines = [ln.strip() for ln in raw.splitlines() if ln.strip()]

    if not lines:
        raise ValueError('没有解析出任何列表项。')

    return [str(x) for x in lines], output_name


class DragDropListWidget(QListWidget):
    """支持拖拽的列表组件"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        # 修改为 DragDrop 模式，同时支持外部拖入和内部移动
        self.setDragDropMode(QAbstractItemView.DragDrop)
        self.setDefaultDropAction(Qt.MoveAction) # 内部默认动作为移动
        self.setSelectionMode(QAbstractItemView.ExtendedSelection)
        # 允许双击编辑列表项文本（用于修改份数 / 修改空白页尺寸）
        self.setEditTriggers(QAbstractItemView.DoubleClicked | QAbstractItemView.EditKeyPressed)

    def dragEnterEvent(self, event):
        # 1. 如果是内部拖拽排序，交回给系统默认逻辑处理
        if event.source() == self:
            super().dragEnterEvent(event)
        # 2. 如果是从外部拖入文件/文件夹
        elif event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if event.source() == self:
            super().dragMoveEvent(event)
        elif event.mimeData().hasUrls():
            event.setDropAction(Qt.CopyAction)
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        # 1. 处理内部拖拽排序
        if event.source() == self:
            super().dropEvent(event)
        # 2. 处理外部文件拖入
        elif event.mimeData().hasUrls():
            event.setDropAction(Qt.CopyAction)
            event.accept()
            for url in event.mimeData().urls():
                if url.isLocalFile():
                    path = str(url.toLocalFile())
                    
                    # 文件夹递归处理
                    if os.path.isdir(path):
                        for root, dirs, files in os.walk(path):
                            for file in files:
                                ext = os.path.splitext(file)[1].lower()
                                if ext in SUPPORTED_EXTS:
                                    full_path = os.path.join(root, file)
                                    self._add_item_if_unique(full_path)
                    # 单文件处理
                    else:
                        ext = os.path.splitext(path)[1].lower()
                        if ext in SUPPORTED_EXTS:
                            self._add_item_if_unique(path)
        else:
            event.ignore()

    def _add_item_if_unique(self, file_path):
        """避免重复添加文件的辅助方法（重复是通过份数实现的，不重复入列）"""
        items = [self.item(i).text() for i in range(self.count())]
        for text in items:
            kind, value = parse_list_item(text)
            if kind == 'file' and value == file_path:
                return
        # 统一带上面向标记：未标记的显示成 [双面]，方便一眼看清每个文档怎么打
        parent = self.parent()
        if isinstance(parent, PDFMergerApp):
            self.addItem(format_file_item(1, file_path, None,
                                          default_side=parent._default_side()))
        else:
            self.addItem(format_file_item(1, file_path, None))

    def keyPressEvent(self, event):
        """Delete 删除选中项；Ctrl+C / Ctrl+V 复制粘贴整个列表"""
        parent = self.parent()
        is_app = isinstance(parent, PDFMergerApp)

        if (event.modifiers() & Qt.ControlModifier) and is_app:
            if event.key() == Qt.Key_C:
                parent.copy_list()
                return
            if event.key() == Qt.Key_V:
                parent.paste_list()
                return

        if event.key() == Qt.Key_Delete and self.state() != QAbstractItemView.EditingState:
            for item in self.selectedItems():
                self.takeItem(self.row(item))
            if is_app:
                parent.sync_buttons()
            return
        super().keyPressEvent(event)

class PDFMergerApp(QWidget):
    def __init__(self):
        super().__init__()
        self.temp_dir = tempfile.gettempdir()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('全能文档转PDF合并器')
        self.resize(550, 450)
        layout = QVBoxLayout()

        # 提示标签
        self.label = QLabel("请将 Word, Excel, PPT, 图片 或 OFD 的 文件/文件夹 拖入下方列表中\n"
                            "拖动可排序，按 Delete 键删除\n"
                            "双击某项可修改份数（[1x]）或空白页尺寸；右键查看更多操作（含复制/粘贴列表）\n"
                            "每个文档前都标了 [双面]/[单面]：带「=」的表示跟随下方默认设置，"
                            "不带的是你单独指定的")
        self.label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.label)

        # 拖拽列表
        self.file_list = DragDropListWidget(self)
        self.file_list.itemChanged.connect(self._on_item_changed)
        self.file_list.itemSelectionChanged.connect(self.sync_buttons)
        self.file_list.setContextMenuPolicy(Qt.CustomContextMenu)
        self.file_list.customContextMenuRequested.connect(self.show_list_menu)
        layout.addWidget(self.file_list)

        # --- 列表操作按钮：清空 / 份数 / 空白页 / 分隔符 ---
        list_btn_layout = QHBoxLayout()
        self.btn_clear = QPushButton("清空列表")
        self.btn_clear.clicked.connect(self.clear_all)

        self.btn_plus = QPushButton("份数 +1")
        self.btn_plus.setToolTip("让选中的文件在最终 PDF 中重复出现更多次")
        self.btn_plus.clicked.connect(lambda: self.change_copies(1))

        self.btn_minus = QPushButton("份数 -1")
        self.btn_minus.clicked.connect(lambda: self.change_copies(-1))

        self.btn_blank = QPushButton("插入空白页")
        self.btn_blank.setToolTip("在选中项之后插入一张空白页，可拖动到任意位置")
        self.btn_blank.clicked.connect(self.insert_blank_page)

        self.btn_sep = QPushButton("插入分隔符")
        self.btn_sep.setToolTip("标记此处分属上下两个不同的 PDF\n"
                                "默认「切分」：拆成多个文件\n"
                                "右键可改成「插入分隔纸」：多插一页作为物理隔断")
        self.btn_sep.clicked.connect(self.insert_separator)

        self.btn_side = QPushButton("单/双面")
        self.btn_side.setToolTip("在「双面 / 单面」之间切换选中的文件\n"
                                 "可按住 Ctrl / Shift 多选，每个文档独立标记，互不影响\n"
                                 "没标记的文件一律按双面处理")
        self.btn_side.clicked.connect(self.toggle_side)

        self.btn_copy = QPushButton("复制列表")
        self.btn_copy.setToolTip("把整个列表复制成 JSON 文本，可粘贴到记事本保存\n"
                                 "以后用「粘贴列表」原样恢复")
        self.btn_copy.clicked.connect(self.copy_list)

        self.btn_paste = QPushButton("粘贴列表")
        self.btn_paste.setToolTip("从剪切板恢复列表（支持替换或追加）")
        self.btn_paste.clicked.connect(self.paste_list)

        list_btn_layout.addWidget(self.btn_clear)
        list_btn_layout.addWidget(self.btn_plus)
        list_btn_layout.addWidget(self.btn_minus)
        list_btn_layout.addWidget(self.btn_blank)
        list_btn_layout.addWidget(self.btn_sep)
        list_btn_layout.addWidget(self.btn_side)
        list_btn_layout.addWidget(self.btn_copy)
        list_btn_layout.addWidget(self.btn_paste)
        layout.addLayout(list_btn_layout)

        # --- 输出设置区域 ---
        settings_layout = QVBoxLayout()
        
        # 1. 目录选择
        dir_layout = QHBoxLayout()
        self.dir_label = QLabel("输出目录:")
        self.dir_input = QLineEdit(QStandardPaths.writableLocation(QStandardPaths.DesktopLocation))
        self.dir_btn = QPushButton("选择...")
        self.dir_btn.clicked.connect(self.select_directory)
        dir_layout.addWidget(self.dir_label)
        dir_layout.addWidget(self.dir_input)
        dir_layout.addWidget(self.dir_btn)
        settings_layout.addLayout(dir_layout)

        # 2. 输出文件名（留空则用默认名 / 自动编号）
        name_layout = QHBoxLayout()
        self.name_label = QLabel("输出文件名:")
        self.name_input = QLineEdit()
        self.name_input.setPlaceholderText("留空则用默认名（单个：合并输出_打印预览；多个：分卷01、分卷02…）")
        self.name_input.setClearButtonEnabled(True)
        self.cb_ask_name = QCheckBox("每次合并前询问")
        self.cb_ask_name.setToolTip("勾选后，点击合并时会先弹出输入框确认文件名")
        name_layout.addWidget(self.name_label)
        name_layout.addWidget(self.name_input)
        name_layout.addWidget(self.cb_ask_name)
        settings_layout.addLayout(name_layout)

        # 3. 预览选项
        self.cb_open = QCheckBox("合并完成后自动打开 PDF 进行打印预览")
        self.cb_open.setChecked(True) # 默认勾选
        settings_layout.addWidget(self.cb_open)

        # 4. 打印面向（单/双面）设置
        self.cb_auto_pad = QCheckBox("自动分叠（按单/双面自动插空白页）")
        self.cb_auto_pad.setToolTip(
            "勾选后一律只按「双面」打印，靠插空白页达成单面/分叠效果。\n"
            "粒度是**每个文件**，各按自己的标记处理：\n\n"
            "· 单面文件：每一页后面自动插一张空白页，\n"
            "  让内容全落在纸的正面、背面留白 —— 效果等同单面打印\n"
            "· 双面文件：它自己页数为奇数时末尾补 1 页凑成偶数，\n"
            "  让它独占整张纸、从正面开始，不会和相邻内容挤在一张纸上\n"
            "· 单个文件标了单面只影响它自己，不影响列表里别的文件\n"
            "· 没标记的文件一律按双面处理（可以在列表里单独改成单面）\n\n"
            "注意：「单面/双面」本身是打印驱动里的选项，写不进 PDF 文件，\n"
            "只能靠插空白页来达成单面的效果。")
        self.cb_auto_pad.setChecked(True)
        settings_layout.addWidget(self.cb_auto_pad)

        layout.addLayout(settings_layout)

        # 合并按钮
        self.btn_generate = QPushButton("一键转换并合并")
        self.btn_generate.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold; height: 40px;")
        self.btn_generate.clicked.connect(self.process_files)
        layout.addWidget(self.btn_generate)

        self.setLayout(layout)
        # 分隔符文本较长（含尺寸与名称），自动换行防截断
        self.file_list.setWordWrap(True)
        self.sync_buttons()

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Delete:
            for item in self.file_list.selectedItems():
                self.file_list.takeItem(self.file_list.row(item))
            self.sync_buttons()

    def select_directory(self):
        """选择输出目录"""
        folder = QFileDialog.getExistingDirectory(self, "选择输出目录", self.dir_input.text())
        if folder:
            self.dir_input.setText(folder)

    # ================== 新增：份数 / 空白页 相关逻辑 ==================

    def sync_buttons(self):
        """根据当前选中项类型，启用/禁用份数与空白页按钮"""
        selected = self.file_list.selectedItems()
        has_file = any(parse_list_item(i.text())[0] == 'file' for i in selected)
        has_sel = len(selected) > 0
        self.btn_plus.setEnabled(has_file)
        self.btn_minus.setEnabled(has_file)
        self.btn_blank.setEnabled(has_sel)
        self.btn_sep.setEnabled(has_sel)
        self.btn_side.setEnabled(has_sel)

    def clear_all(self):
        self.file_list.clear()
        self.sync_buttons()

    def _on_item_changed(self, item):
        """用户编辑完列表项后，把文本规范化（份数 / 空白页尺寸 / 分隔符模式）"""
        self._normalize_item(item, warn_missing=True)

    def _normalize_item(self, item, warn_missing=True):
        """把一项的文本规范化，可复用于编辑、粘贴与恢复。

        warn_missing=False 用于批量恢复历史列表：文件不存在只记录、不逐条弹窗。
        """
        text = item.text()

        # --- 分隔符项 ---
        if text.startswith(SEP_PREFIX):
            mode, _, width, height, doc_name, side = parse_sep_spec(text)
            normalized = format_sep_text(mode, doc_name, width, height, side)
            if text != normalized:
                self._block_change(True)
                item.setText(normalized)
                self._block_change(False)
            return

        # --- 空白页项 ---
        if text.startswith(BLANK_PREFIX):
            spec = text[len(BLANK_PREFIX):].strip()
            size = parse_blank_spec(spec)
            if size is None:
                if warn_missing:
                    QMessageBox.warning(self, "格式不正确",
                                        f"无法识别的空白页尺寸：{spec}\n\n"
                                        "可用写法：A4纵向 / A5横向 / Letter纵向\n"
                                        "或自定义宽x高（单位厘米）：21x29.7")
            else:
                w, h = size
                normalized = f"{w:g}x{h:g}cm"
                if spec != normalized:
                    self._block_change(True)
                    item.setText(f"{BLANK_PREFIX}{normalized}")
                    self._block_change(False)
            return

        # --- 文件项：解析 [Nx] [面向] 路径 ---
        copies, side, path = parse_file_display(text)

        if not os.path.isfile(path):
            if warn_missing:
                QMessageBox.warning(self, "文件不存在", f"找不到该文件，已忽略本次修改：\n{path}")
            path = getattr(item, '_last_path', path) or path
            copies = getattr(item, '_last_copies', 1) or 1

        item._last_path = path
        item._last_copies = copies
        normalized = format_file_item(copies, path, side,
                                     default_side=self._default_side())
        if text != normalized:
            self._block_change(True)
            item.setText(normalized)
            self._block_change(False)

    def _block_change(self, block):
        """临时屏蔽 itemChanged 信号，避免 setText 造成递归"""
        self.file_list.blockSignals(block)

    def _set_stage(self, text):
        """把进度文字显示到合并按钮上。

        模块级的转换/合并函数通过 progress 回调调用它，
        这样那些函数就不必知道界面的存在（CLI 同样复用它们）。
        """
        self.btn_generate.setText(text)
        QApplication.processEvents()

    def change_copies(self, delta):
        """让选中文件在最终 PDF 中重复出现 delta 次"""
        selected = [i for i in self.file_list.selectedItems()
                    if parse_list_item(i.text())[0] == 'file']
        if not selected:
            return
        self._block_change(True)
        for item in selected:
            _kind, path = parse_list_item(item.text())
            copies, side, real_path = parse_file_display(item.text())
            copies = max(1, copies + delta)
            item.setText(format_file_item(copies, real_path, side,
                                          default_side=self._default_side()))
            item._last_path = real_path
            item._last_copies = copies
        self._block_change(False)
        self.sync_buttons()

    def insert_blank_page(self):
        """在当前选中项之后插入一张空白页；没有选中则追加到末尾"""
        selected = self.file_list.selectedItems()
        spec = DEFAULT_BLANK_KEY
        if selected:
            row = self.file_list.row(selected[-1])
        else:
            row = self.file_list.count() - 1

        if selected:
            # 若选中的是空白页，沿用它的尺寸，连续插入更省事
            kind, value = parse_list_item(selected[-1].text())
            if kind == 'blank':
                spec = value

        text = f"{BLANK_PREFIX}{spec}"

        self._block_change(True)
        self.file_list.insertItem(row + 1, text)
        self._block_change(False)

        new_item = self.file_list.item(row + 1)
        new_item._last_path = None
        new_item._last_copies = 1
        self.file_list.clearSelection()
        new_item.setSelected(True)
        self.file_list.setCurrentItem(new_item)
        self.file_list.scrollToItem(new_item)
        self.sync_buttons()

    # ================== 新增：分隔符相关逻辑 ==================

    def insert_separator(self):
        """在选中项之后插入一个分隔符；没有选中则追加到末尾。
        分隔符默认沿用同类型项的设置，方便连续插多个同样的分隔符。
        """
        selected = self.file_list.selectedItems()
        if selected:
            row = self.file_list.row(selected[-1])
        else:
            row = self.file_list.count() - 1

        mode, width, height, doc_name, side = SEP_DEFAULT_MODE, None, None, None, None
        if selected:
            kind, value = parse_list_item(selected[-1].text())
            if kind == 'sep':
                mode, _, width, height, doc_name, side = parse_sep_spec(selected[-1].text())

        text = format_sep_text(mode, doc_name, width, height, side)

        self._block_change(True)
        self.file_list.insertItem(row + 1, text)
        self._block_change(False)

        new_item = self.file_list.item(row + 1)
        new_item._last_path = None
        new_item._last_copies = 1
        self.file_list.clearSelection()
        new_item.setSelected(True)
        self.file_list.setCurrentItem(new_item)
        self.file_list.scrollToItem(new_item)
        self.sync_buttons()

    def set_sep_mode(self, mode):
        """切换选中分隔符的模式（切分 / 插入分隔纸）"""
        targets = [i for i in self.file_list.selectedItems()
                   if parse_list_item(i.text())[0] == 'sep']
        if not targets:
            return
        self._block_change(True)
        for item in targets:
            _, _, old_w, old_h, doc_name, side = parse_sep_spec(item.text())
            if mode == SEP_SHEET:
                w, h = (old_w, old_h) if old_w else BLANK_PRESETS[DEFAULT_BLANK_KEY]
            else:
                w, h = None, None
            item.setText(format_sep_text(mode, doc_name, w, h, side))
        self._block_change(False)

    def set_sep_size(self):
        """设置选中分隔符「插入分隔纸」时的纸张规格"""
        targets = [i for i in self.file_list.selectedItems()
                   if parse_list_item(i.text())[0] == 'sep']
        if not targets:
            QMessageBox.information(self, "提示", "请先在列表中选择一个分隔符。")
            return
        text, ok = QInputDialog.getText(
            self, "设置分隔纸尺寸",
            "输入纸张规格或自定义宽x高（单位厘米）：\n"
            "预置可选：A3纵向 / A3横向 / A4纵向 / A4横向 / A5纵向 / A5横向 / Letter纵向 / Letter横向\n"
            "自定义示例：21x29.7",
            QLineEdit.Normal, DEFAULT_BLANK_KEY)
        if not ok or not text.strip():
            return
        size = parse_blank_spec(text.strip())
        if size is None:
            QMessageBox.warning(self, "格式不正确", f"无法识别的尺寸：{text}")
            return
        w, h = size
        self._block_change(True)
        for item in targets:
            _, _, _, _, doc_name, side = parse_sep_spec(item.text())
            # 明确设置了纸张规格，自动切到「插入分隔纸」模式
            item.setText(format_sep_text(SEP_SHEET, doc_name, w, h, side))
        self._block_change(False)

    def set_sep_name(self):
        """给选中分隔符命名，命名后这段会输出成“名字.pdf”"""
        targets = [i for i in self.file_list.selectedItems()
                   if parse_list_item(i.text())[0] == 'sep']
        if not targets:
            QMessageBox.information(self, "提示", "请先在列表中选择一个分隔符。")
            return

        _, _, _, _, current, _side = parse_sep_spec(targets[-1].text())
        text, ok = QInputDialog.getText(
            self, "分隔符名称",
            "给这一段的输出文件起个名字（留空则不单独命名，按序号自动编号）：",
            QLineEdit.Normal, current or '')
        if not ok:
            return
        name = sanitize_filename(text)
        self._block_change(True)
        for item in targets:
            mode, _, w, h, _, side = parse_sep_spec(item.text())
            item.setText(format_sep_text(mode, name or None, w, h, side))
        self._block_change(False)

    # ================== 新增：单面 / 双面打印标记 ==================

    def _cycle_side(self, cur, default_side=None):
        """在「双面 / 单面」之间切换，取当前**生效**的面向作为起点。

        只做两态切换（不绕「默认」）：
          · 列表里每个文档都显示 [双面] 或 [单面]，三态循环时切到「默认」
            在界面上看起来像没反应，反而让人以为标记丢了
          · 没标记的文件一律按双面处理，点一下即变成单面
        """
        effective = self._effective_side(cur, default_side)
        return SIDE_SIMPLEX if effective == SIDE_DUPLEX else SIDE_DUPLEX

    def _default_side(self):
        """未标记的文件按什么处理。

        固定为双面 —— 这是「全程只选双面打印」这个方案的前提：
        单面的文件会靠插空白页来达成效果，所以默认面必然是双面。
        """
        return SIDE_DUPLEX

    def _effective_side(self, explicit, default_side=None):
        """把「未标记」折算成实际生效的面向"""
        if explicit:
            return explicit
        return default_side or self._default_side()

    def _apply_side_to_item(self, item, side):
        """把面向写进一个列表项；side 为 None 表示未标记（按默认的双面处理）。"""
        kind, _ = parse_list_item(item.text())
        if kind == 'file':
            copies, _old, path = parse_file_display(item.text())
            item.setText(format_file_item(copies, path, side,
                                          default_side=self._default_side()))
        elif kind == 'sep':
            mode, _, w, h, name, _ = parse_sep_spec(item.text())
            item.setText(format_sep_text(mode, name, w, h, side))

    def _selected_side_sources(self):
        """取出当前选中项里承载面向的项（文件项 + 分隔符项）。

        分隔符的面向作用于它下方那一段；空白页不承载面向。
        """
        out = []
        for item in self.file_list.selectedItems():
            if parse_list_item(item.text())[0] in ('file', 'sep'):
                out.append(item)
        return sorted(out, key=self.file_list.row)

    def toggle_side(self):
        """点击按钮：把选中项（可多选）的打印面向在 双面 / 单面 之间切换。

        每一项独立切换，互不影响 —— 之前是「整段共用一个标记 + 每次先清空全部」，
        所以标记一个文档会把另一个文档的标记抹掉。
        """
        targets = self._selected_side_sources()
        if not targets:
            QMessageBox.information(
                self, "提示",
                "请先在列表中选择要标记的文件或分隔符（可按住 Ctrl / Shift 多选）。\n\n"
                "点击在 双面 ↔ 单面 之间切换。")
            return

        default = self._default_side()
        self._block_change(True)
        for item in targets:
            kind, _ = parse_list_item(item.text())
            if kind == 'sep':
                cur = parse_sep_spec(item.text())[5]
            else:
                cur = parse_file_display(item.text())[1]
            self._apply_side_to_item(item, self._cycle_side(cur, default))
        self._block_change(False)
        self.sync_buttons()

    def set_blank_size(self):
        """弹窗让用户输入空白页尺寸，应用到所有选中的空白页"""
        targets = [i for i in self.file_list.selectedItems()
                   if parse_list_item(i.text())[0] == 'blank']
        if not targets:
            QMessageBox.information(self, "提示", "请先在列表中选择一个空白页。")
            return
        text, ok = QInputDialog.getText(
            self, "设置空白页尺寸",
            "输入纸张规格或自定义宽x高（单位厘米）：\n"
            "预置可选：A3纵向 / A3横向 / A4纵向 / A4横向 / A5纵向 / A5横向 / Letter纵向 / Letter横向\n"
            "自定义示例：21x29.7",
            QLineEdit.Normal, DEFAULT_BLANK_KEY)
        if not ok or not text.strip():
            return
        size = parse_blank_spec(text.strip())
        if size is None:
            QMessageBox.warning(self, "格式不正确", f"无法识别的尺寸：{text}")
            return
        w, h = size
        new_text = f"{BLANK_PREFIX}{w:g}x{h:g}cm"
        self._block_change(True)
        for item in targets:
            item.setText(new_text)
        self._block_change(False)

    def show_list_menu(self, pos):
        """列表右键菜单"""
        item = self.file_list.itemAt(pos)
        if item and not item.isSelected():
            self.file_list.clearSelection()
            item.setSelected(True)
            self.file_list.setCurrentItem(item)

        selected = self.file_list.selectedItems()
        kind = parse_list_item(item.text())[0] if item else None

        menu = QMenu(self)
        if kind == 'file':
            menu.addAction("份数 +1", lambda: self.change_copies(1))
            menu.addAction("份数 -1", lambda: self.change_copies(-1))
            menu.addSeparator()
            self._add_side_actions(menu)
        elif kind == 'blank':
            menu.addAction("设置空白页尺寸...", self.set_blank_size)
        elif kind == 'sep':
            cur_mode, _, _, _, doc_name, cur_side = parse_sep_spec(item.text())
            menu.addAction("分隔符模式：", None).setEnabled(False)
            for value, label, tip in SEP_MODES:
                act = menu.addAction(("● " if value == cur_mode else "○ ") + label)
                act.setToolTip(tip)
                act.triggered.connect(lambda _=False, v=value: self.set_sep_mode(v))
            menu.addSeparator()
            menu.addAction("设置分隔纸尺寸...", self.set_sep_size)
            menu.addAction("重命名这一段...", self.set_sep_name)
            if doc_name:
                menu.addAction("清除名称", self._clear_sep_name)
            menu.addSeparator()
            self._add_side_actions(menu)
        menu.addAction("在此项之后插入空白页", self.insert_blank_page)
        menu.addAction("在此项之后插入分隔符", self.insert_separator)
        menu.addSeparator()
        menu.addAction("复制整个列表", self.copy_list)
        menu.addAction("粘贴列表", self.paste_list)
        menu.addAction("查看/编辑列表文本...", self.show_list_text)
        if selected:
            menu.addSeparator()
            menu.addAction("删除选中项", self.delete_selected)
        menu.exec_(self.file_list.viewport().mapToGlobal(pos))

    def _add_side_actions(self, menu):
        """给右键菜单加单/双面子菜单（对当前选中项生效）"""
        item = self.file_list.currentItem()
        cur = None
        if item:
            if parse_list_item(item.text())[0] == 'sep':
                cur = parse_sep_spec(item.text())[5]
            else:
                cur = parse_file_display(item.text())[1]

        n = len(self._selected_side_sources())
        # 未标记的项按默认面（双面）显示，菜单标题才和列表里看到的一致
        title = "打印面向（当前：%s）" % SIDE_LABELS.get(cur or self._default_side())
        if n > 1:
            title = f"打印面向（已选 {n} 项）"
        sub = menu.addMenu(title)
        for value, label in [(SIDE_DUPLEX, '双面'), (SIDE_SIMPLEX, '单面')]:
            act = sub.addAction(("● " if value == cur else "○ ") + label)
            act.triggered.connect(lambda _=False, v=value: self._set_side(v))

    def _set_side(self, side):
        """把选中的每一项显式设为某个面向（只影响选中的，不动其他文档）。"""
        targets = self._selected_side_sources()
        if not targets:
            QMessageBox.information(
                self, "提示",
                "请先在列表中选择要标记的文件或分隔符（可多选）。")
            return

        self._block_change(True)
        for item in targets:
            self._apply_side_to_item(item, side)
        self._block_change(False)
        self.sync_buttons()

    def _clear_sep_name(self):
        """清除选中分隔符的名称"""
        targets = [i for i in self.file_list.selectedItems()
                   if parse_list_item(i.text())[0] == 'sep']
        self._block_change(True)
        for item in targets:
            mode, _, w, h, _, side = parse_sep_spec(item.text())
            item.setText(format_sep_text(mode, None, w, h, side))
        self._block_change(False)

    def delete_selected(self):
        for item in self.file_list.selectedItems():
            self.file_list.takeItem(self.file_list.row(item))
        self.sync_buttons()

    # ================== 列表的复制 / 粘贴 ==================

    def _list_lines(self):
        """当前列表的每一行文本"""
        return [self.file_list.item(i).text() for i in range(self.file_list.count())]

    def _restore_list(self, lines):
        """用给定文本重建列表，并规范化每一项"""
        self._block_change(True)
        self.file_list.clear()
        for text in lines:
            self.file_list.addItem(text)
        for i in range(self.file_list.count()):
            self._normalize_item(self.file_list.item(i), warn_missing=False)
        self._block_change(False)
        self.sync_buttons()

    def copy_list(self):
        """把整个列表复制成 JSON 文本（粘贴到任何地方都能存下来）。"""
        if self.file_list.count() == 0:
            QMessageBox.information(self, "提示", "列表是空的，没有可复制的内容。")
            return
        payload = dump_list_json(self._list_lines(), self.name_input.text().strip())
        QApplication.clipboard().setText(payload)
        n = self.file_list.count()
        QMessageBox.information(
            self, "已复制",
            f"已把 {n} 项复制到剪切板（JSON 格式）。\n\n"
            "把它粘贴到记事本、聊天窗口或笔记里保存，\n"
            "以后用「粘贴列表」就能原样恢复。")

    def paste_list(self):
        """从剪切板恢复列表，可选择替换或追加。"""
        text = QApplication.clipboard().text()
        if not (text or '').strip():
            QMessageBox.information(self, "提示", "剪切板里没有文本内容。")
            return

        try:
            lines, output_name = load_list_json(text)
        except ValueError as exc:
            QMessageBox.warning(
                self, "无法解析",
                f"{exc}\n\n"
                "请确认复制的是本程序「复制整个列表」生成的 JSON，\n"
                "或者每行一项的纯文本。")
            return

        mode = 'replace'
        if self.file_list.count() > 0:
            box = QMessageBox(self)
            box.setWindowTitle("粘贴列表")
            box.setText(f"解析出 {len(lines)} 项。要如何粘贴？")
            btn_replace = box.addButton("替换当前列表", QMessageBox.AcceptRole)
            btn_append = box.addButton("追加到末尾", QMessageBox.ActionRole)
            box.addButton("取消", QMessageBox.RejectRole)
            box.exec_()
            clicked = box.clickedButton()
            if clicked is btn_replace:
                mode = 'replace'
            elif clicked is btn_append:
                mode = 'append'
            else:
                return

        if mode == 'replace':
            self._restore_list(lines)
        else:
            self._block_change(True)
            for text_line in lines:
                self.file_list.addItem(text_line)
            for i in range(self.file_list.count()):
                self._normalize_item(self.file_list.item(i), warn_missing=False)
            self._block_change(False)
            self.sync_buttons()

        if output_name and not self.name_input.text().strip():
            self.name_input.setText(output_name)
        self._warn_missing_files()

    def show_list_text(self):
        """弹出一个可编辑文本框，直接看/改整份列表（最灵活的备份与恢复方式）。"""
        from PyQt5.QtWidgets import QDialog, QPlainTextEdit, QDialogButtonBox

        dlg = QDialog(self)
        dlg.setWindowTitle("列表文本（可直接编辑后确定应用）")
        dlg.resize(640, 420)
        v = QVBoxLayout(dlg)
        v.addWidget(QLabel("每行一项，格式与列表里显示的一致：\n"
                           "[份数x] [双面|单面] 文件路径\n"
                           "[空白页] 21x29.7cm\n"
                           "[分隔符] 切分 单面 名为 附件二"))
        edit = QPlainTextEdit()
        edit.setPlainText(dump_list_json(self._list_lines(),
                                         self.name_input.text().strip()))
        v.addWidget(edit)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("应用")
        buttons.button(QDialogButtonBox.Cancel).setText("取消")
        buttons.accepted.connect(dlg.accept)
        buttons.rejected.connect(dlg.reject)
        v.addWidget(buttons)

        if dlg.exec_() != QDialog.Accepted:
            return

        try:
            lines, output_name = load_list_json(edit.toPlainText())
        except ValueError as exc:
            QMessageBox.warning(self, "无法解析", str(exc))
            return

        self._restore_list(lines)
        if output_name and not self.name_input.text().strip():
            self.name_input.setText(output_name)
        self._warn_missing_files()

    def _warn_missing_files(self):
        """粘贴后提示哪些文件在本机找不到（换机器时最常见的问题）。"""
        missing = []
        for i in range(self.file_list.count()):
            kind, value = parse_list_item(self.file_list.item(i).text())
            if kind == 'file':
                path = parse_file_display(self.file_list.item(i).text())[2]
                if not os.path.isfile(path):
                    missing.append(path)
        if missing:
            head = '\n'.join(missing[:10])
            more = f"\n…还有 {len(missing) - 10} 个" if len(missing) > 10 else ''
            QMessageBox.warning(
                self, "有文件找不到",
                f"以下 {len(missing)} 个文件在当前电脑上不存在，"
                f"合并时会自动跳过：\n\n{head}{more}")

    # ================== 合并主流程 ==================


    def get_merge_tasks(self):
        """读取界面列表，生成合并任务序列。

        解析规则见模块级 build_tasks_from_lines（CLI 共用同一实现），
        这里只负责把列表内容取出来交给它。
        """
        return build_tasks_from_lines(self._list_lines())





    def process_files(self):
        count = self.file_list.count()
        if count == 0:
            QMessageBox.warning(self, "警告", "请先拖入需要处理的文件！")
            return

        tasks, warnings = self.get_merge_tasks()
        file_tasks = [t for t in tasks if t[0] == 'file']
        if not file_tasks:
            QMessageBox.warning(self, "警告", "列表中没有可处理的文件！")
            return

        if warnings:
            QMessageBox.warning(self, "部分内容已跳过",
                                "以下内容不会被合并：\n\n" + "\n".join(warnings[:10]))

        out_dir = self.dir_input.text()
        if not os.path.exists(out_dir):
            try:
                os.makedirs(out_dir)
            except Exception as e:
                QMessageBox.critical(self, "错误", f"无法创建输出目录:\n{str(e)}")
                return

        segments = build_segments(tasks)
        total_segments = len(segments)
        has_sep = any(t[0] == 'sep' for t in tasks)

        # --- 解析输出文件名 ---
        user_name = self.name_input.text().strip()
        if self.cb_ask_name.isChecked():
            hint = ("留空使用默认名。可用 # 作为序号占位符。"
                    if total_segments > 1 else "留空使用默认名。")
            if total_segments > 1:
                hint += f"\n本次将输出 {total_segments} 个文件。"
            user_name, ok = QInputDialog.getText(
                self, "输出文件名", hint, QLineEdit.Normal, user_name)
            if not ok:
                return
            user_name = user_name.strip()
            self.name_input.setText(user_name)
        else:
            # 多段但没写占位符时提醒一下会加序号
            if total_segments > 1 and user_name and SEQ_PLACEHOLDER not in user_name:
                QMessageBox.information(
                    self, "提示",
                    f"本次会输出 {total_segments} 个文件，\n"
                    f"你填的名字没有序号占位符，程序会自动加上 _1、_2… 以区分。\n\n"
                    f"如果想让序号接在名字中间，可用 # 占位，例如：\n"
                    f"合同#_已签字")

        default_single = sanitize_filename(user_name) or '合并输出_打印预览'
        if total_segments > 1:
            default_single = '合并输出_打印预览'

        self.btn_generate.setText("正在处理中，请稍候...")
        self.btn_generate.setEnabled(False)
        QApplication.processEvents()

        # 同一文件只转换一次，重复份数复用转换结果
        pdf_cache = {}
        failed = []
        auto_pad = self.cb_auto_pad.isChecked()
        default_side = self._default_side()

        try:
            results, errors = write_outputs(
                out_dir, segments, pdf_cache, failed, self.temp_dir,
                user_name=user_name, default_single=default_single,
                auto_pad=auto_pad, default_side=default_side,
                progress=self._set_stage)

            # 某一段写盘失败（比如原文件被占用）时，降级为单段输出并给出提示
            if errors and len(segments) > 1:
                reason = "\n".join(f"{name}：{err}" for name, err, _, _ in errors[:5])
                if len(errors) > 5:
                    reason += f"\n... 共 {len(errors)} 段失败"
                fallback_name = (sanitize_filename(user_name) or '合并输出_打印预览')
                if SEQ_PLACEHOLDER in fallback_name:
                    fallback_name = fallback_name.replace(SEQ_PLACEHOLDER, '')
                    fallback_name = sanitize_filename(fallback_name) or '合并输出_打印预览'
                fallback_file = fallback_name + ".pdf"
                merged_path = os.path.join(out_dir, fallback_file)
                try:
                    merged_pages = write_segment(
                        tasks, merged_path, pdf_cache, failed, self.temp_dir,
                        progress=self._set_stage)
                    results = [(fallback_file, merged_pages, None, 0)]
                    errors = []
                    QMessageBox.warning(self, "已降级为单文件输出",
                                        f"分卷时出现错误，已改为输出一个完整文件：\n{reason}")
                except Exception as e:
                    raise Exception(f"分卷失败，且降级合并也失败：\n{reason}\n\n{e}")

            # 注意：这里不再清空列表，方便你微调后重跑
            if not results:
                raise Exception("没有任何内容被写出，请检查文件是否有效。")

            total_pages = sum(p for _, p, _s, _pad in results)

            # 面向标记：没标的按默认面（双面）处理（default_side 已在上面取好）
            lines = []
            simplex_segs = []
            for name, pages, side, pad in results:
                eff = side or default_side
                tag = SIDE_LABELS.get(eff, '')
                if eff == SIDE_SIMPLEX and pad:
                    extra = f"，每页后插 1 页空白（共 {pad} 页）"
                    simplex_segs.append(name)
                elif pad:
                    extra = f"，末尾补 {pad} 页空白"
                else:
                    extra = ""
                lines.append(f"{name}（{pages} 页，{tag}{extra}）")

            fail_msg = ""
            if failed:
                fail_msg = "\n\n以下内容处理失败（已跳过）：\n" + "\n".join(failed[:10])
                if len(failed) > 10:
                    fail_msg += f"\n... 共 {len(failed)} 项失败"

            summary = f"输出目录：\n{out_dir}\n\n"
            if len(results) == 1 and not has_sep:
                summary += f"文件：{results[0][0]}\n共 {total_pages} 页"
            else:
                summary += f"共生成 {len(results)} 个文件（合计 {total_pages} 页）：\n" + "\n".join(lines)

            if auto_pad:
                summary += ("\n\n已启用自动分叠（逐个文件按它自己的标记处理）：\n"
                            "· 单面文件：每一页后面自动插一张空白页，\n"
                            "  让内容全落在纸的正面、背面留白，效果等同单面打印\n"
                            "· 双面文件：它自己页数为奇数时末尾补 1 页，\n"
                            "  凑成偶数独占整张纸、从正面开始\n"
                            "· 所以一个文件标了单面，只影响它自己，不影响别的文件")
                if simplex_segs:
                    names = '、'.join(simplex_segs[:5])
                    if len(simplex_segs) > 5:
                        names += f" 等 {len(simplex_segs)} 个"
                    summary += ("\n\n本次按单面插页的输出段：\n"
                                f"{names}")
                summary += ("\n\n提醒：「单面/双面」本身是打印驱动里的选项，写不进 PDF；\n"
                            "这里靠插空白页来达到单面的效果，所以打印时依然选「双面」即可。")
            else:
                summary += "\n\n提示：未启用自动分叠。不同文件若混排，注意按上面的单/双面分别打印。"

            detail = summary + fail_msg

            if self.cb_open.isChecked() and len(results) == 1:
                os.startfile(os.path.join(out_dir, results[0][0]))
                if fail_msg or len(results) > 1:
                    QMessageBox.warning(self, "完成", detail)
            else:
                QMessageBox.information(self, "完成", detail)

        except Exception as e:
            QMessageBox.critical(self, "错误", f"处理过程中出错:\n{str(e)}")
        finally:
            # 清理临时文件
            for p in pdf_cache.values():
                if p != out_dir and os.path.dirname(p) == self.temp_dir:
                    try:
                        os.remove(p)
                    except Exception:
                        pass
            self.btn_generate.setText("一键转换并合并")
            self.btn_generate.setEnabled(True)
            self.sync_buttons()

    # --- 各种格式转换核心逻辑 ---
    



        


# ============ 转换与合并核心（不依赖 GUI，供 GUI 与 CLI 共用） ============
#
# 这些函数原本是 PDFMergerApp 的方法，但逻辑上完全不需要界面 ——
# 唯一与 GUI 有关的只是「进度提示」，所以统一改成 progress 回调：
# GUI 传一个更新按钮文字的函数，CLI 传 None（默认什么都不做）。
# 这样 CLI 与 GUI 共用同一套实现，不会出现两份各自腐化的逻辑。


def _noop_progress(_text):
    """progress 回调的默认实现：什么都不做"""
    pass


def convert_powerpoint(input_path, output_path):
    """PowerPoint -> PDF（优先 MS PowerPoint，其次 WPS 演示）"""
    app = None
    presentation = None
    try:
        try:
            app = win32com.client.DispatchEx("PowerPoint.Application")
        except Exception:
            try:
                app = win32com.client.DispatchEx("kwpp.Application")
            except Exception:
                app = win32com.client.DispatchEx("WPP.Application")
        # ReadOnly=1, Untitled=0, WithWindow=0 —— 避免弹出前台窗口
        presentation = app.Presentations.Open(os.path.abspath(input_path), 1, 0, 0)
        # 32 = ppSaveAsPDF
        presentation.SaveAs(os.path.abspath(output_path), 32)
    finally:
        if presentation:
            presentation.Close()
        if app:
            app.Quit()


def convert_word(input_path, output_path):
    """Word -> PDF（优先 MS Word，其次 WPS 文字）"""
    app = None
    try:
        try:
            app = win32com.client.DispatchEx("Word.Application")
        except Exception:
            app = win32com.client.DispatchEx("kwps.Application")
        app.Visible = False
        app.DisplayAlerts = 0
        doc = app.Documents.Open(os.path.abspath(input_path), ReadOnly=True)
        # 17 = wdExportFormatPDF，WPS 也通用
        doc.SaveAs(os.path.abspath(output_path), FileFormat=17)
        doc.Close()
    finally:
        if app:
            app.Quit()


def convert_excel(input_path, output_path):
    """Excel -> PDF（优先 MS Excel，其次 WPS 表格）"""
    app = None
    try:
        try:
            app = win32com.client.DispatchEx("Excel.Application")
        except Exception:
            try:
                app = win32com.client.DispatchEx("ket.Application")
            except Exception:
                app = win32com.client.DispatchEx("ET.Application")
        app.Visible = False
        app.DisplayAlerts = False
        wb = app.Workbooks.Open(os.path.abspath(input_path), ReadOnly=True)
        # 把每个工作表压到 1 页宽，避免导出时出现大量冗余空白页
        for sheet in wb.Worksheets:
            try:
                sheet.PageSetup.Zoom = False
                sheet.PageSetup.FitToPagesWide = 1
                sheet.PageSetup.FitToPagesTall = False
            except Exception:
                pass
        wb.ExportAsFixedFormat(0, os.path.abspath(output_path))
        # SaveChanges=False 非常重要：绝不回写源文件
        wb.Close(SaveChanges=False)
    finally:
        if app:
            app.Quit()


def convert_image(input_path, output_path):
    """图片 -> PDF"""
    image = Image.open(input_path)
    if image.mode != 'RGB':
        image = image.convert('RGB')
    image.save(output_path)


def convert_ofd(input_path, output_path):
    """OFD -> PDF（依赖 easyofd）"""
    if not HAS_EASYOFD:
        raise Exception("未安装 easyofd 库！请检查环境。")

    with open(input_path, "rb") as f:
        ofdb64 = str(base64.b64encode(f.read()), "utf-8")

    ofd = OFD()            # 初始化 OFD 工具类
    ofd.read(ofdb64)       # 读取 ofdb64，此处不生成多余的 xml
    pdf_bytes = ofd.to_pdf()
    ofd.del_data()         # 清理内存

    with open(output_path, "wb") as f:
        f.write(pdf_bytes)


def convert_to_pdf(file_path, temp_pdf):
    """按扩展名把单个文件转成 PDF；PDF 直接返回原路径。"""
    ext = os.path.splitext(file_path)[1].lower()
    if ext in ['.docx', '.doc']:
        convert_word(file_path, temp_pdf)
    elif ext in ['.xlsx', '.xls']:
        convert_excel(file_path, temp_pdf)
    elif ext in ['.ppt', '.pptx']:
        convert_powerpoint(file_path, temp_pdf)
    elif ext in ['.jpg', '.jpeg', '.png', '.bmp']:
        convert_image(file_path, temp_pdf)
    elif ext == '.ofd':
        convert_ofd(file_path, temp_pdf)
    elif ext == '.pdf':
        return file_path
    else:
        raise Exception(f"不支持的文件格式: {ext}")
    return temp_pdf


def build_tasks_from_lines(lines, check_exists=True):
    """把列表项文本转成合并任务序列。

    返回 (tasks, warnings)：
      tasks: [('file', 路径, 份数, side) / ('blank', 宽cm, 高cm)
              / ('sep', 模式, 宽cm, 高cm, 名称, side), ...]
      warnings: 被跳过的项及原因

    check_exists=True 时校验文件是否真的存在 —— GUI 与 CLI 都需要，
    避免把「路径打错」静默地变成「内容莫名少了一截」。
    """
    tasks = []
    warnings = []

    for i, text in enumerate(lines):
        kind, value = parse_list_item(text)

        if kind == 'sep':
            mode, _, w, h, doc_name, side = parse_sep_spec(text)
            tasks.append(('sep', mode, w, h, doc_name, side))
            continue

        if kind == 'blank':
            size = parse_blank_spec(value)
            if size is None:
                warnings.append(f"第 {i + 1} 项空白页尺寸无法识别，已跳过：{value}")
            else:
                tasks.append(('blank', size[0], size[1]))
            continue

        # 文件项：[1x] [双面] D:\a.pdf —— 中间的面向标记可选
        # 用 parse_file_display：把历史格式的「=双面」折回 None，不影响打印结果
        copies, side, path = parse_file_display(text)

        if check_exists and not os.path.isfile(path):
            warnings.append(f"第 {i + 1} 项文件不存在，已跳过：{path}")
            continue
        ext = os.path.splitext(path)[1].lower()
        if ext not in SUPPORTED_EXTS:
            warnings.append(f"第 {i + 1} 项格式不支持，已跳过：{path}")
            continue

        tasks.append(('file', path, copies, side))

    return tasks, warnings


def build_segments(tasks):
    """按分隔符把任务序列切成若干段。
    返回 [(段名, 任务列表, side), ...]；「插入分隔纸」模式会往对应段追加一张空白页。
    若列表中没有任何分隔符，则整体算作一段（段名为 None）。

    关于命名：一个分隔符处只能写一个名字，而它同时是上一段的终点和
    下一段的起点。这里的规则是——名字统一落在「下一段」（即分隔符下方那段），
    符合「给它下面那份文件起名」的直觉。

    关于面向（side）：分隔符上写的面向作用于它**下方**那一段，与命名规则一致。
    这里返回的段级 side **只用于结果汇报**；真正决定插页的是每个文件自己的
    标记，见 write_segment。
    """
    def head_side_of(seg):
        """取一段里第一个文件的显式面向（没有就返回 None）"""
        for t in seg:
            if t[0] == 'file':
                return t[3] if len(t) > 3 else None
        return None

    if not any(t[0] == 'sep' for t in tasks):
        # 无分隔符：整段就是一段，面向取段内第一个文件的标记
        return [(None, list(tasks), head_side_of(tasks))]

    # 源列表里开头就没有文件（一上来就是分隔符）时，前面的空段直接丢弃
    normalized = []
    for t in tasks:
        if t[0] == 'sep':
            if normalized:
                normalized.append(t)
        else:
            normalized.append(t)
    tasks = normalized

    segments = []
    current = []
    pending_name = None
    pending_side = None

    for task in tasks:
        if task[0] != 'sep':
            current.append(task)
            continue

        _, mode, w_cm, h_cm, doc_name, side = task
        if mode == SEP_SHEET:
            # 物理隔断：这一页纸属于它上面的那份 PDF
            current.append(('blank', w_cm, h_cm))

        if current:
            # 该段的面向：优先用上方分隔符指定的，否则看段内第一个文件的标记
            segments.append((pending_name, current,
                             pending_side or head_side_of(current)))
        current = []
        pending_name = doc_name or None
        pending_side = side

    if current:
        # 末段的面向：优先用分隔符指定的，否则看段内第一个文件的标记
        seg_side = pending_side or head_side_of(current)
        segments.append((pending_name, current, seg_side))

    return [(name, seg, side) for name, seg, side in segments if seg]


def write_segment(segment_tasks, out_path, pdf_cache, failed, temp_dir,
                  pad_pages=0, pad_size=None, simplex=False, auto_pad=False,
                  default_side=None, progress=None):
    """把一段任务写成一个 PDF 文件；返回写入的页数（含插入的空白页）。
    至少保证输出一页（全空时补一张 A4 空白页），避免出现 0 页的无效 PDF。
    任何异常都会向上抛出，由调用方决定是整批失败还是降级跳过。

    pad_pages: 末尾补几张空白页（段级兜底，一般由文件级逻辑接管后传 0）
    pad_size:  (宽cm, 高cm)，空白页尺寸；None 则用 A4 纵向
    simplex:   段级单面开关（兼容旧调用；新逻辑以文件级标记为准）
    auto_pad:  是否启用「自动分叠」。开启时**逐个文件**按它自己的
               单/双面标记处理：
                 · 单面文件：每页内容后面插一张空白页，
                   让内容全部落在纸张正面、背面留白
                 · 双面文件：它自己页数为奇数时末尾补 1 页，
                   凑成偶数让它独占整张纸、从正面开始
               「单面/双面」是打印驱动的运行时选项、写不进 PDF，
               只能靠这种插页方式达成单面效果（打印时全程选双面即可）。
    default_side: 文件未显式标记时按它折算
    progress:  进度回调 callable(text)，可为 None
    """
    progress = progress or _noop_progress
    writer = PdfWriter()

    def blank_size():
        if pad_size:
            return pad_size
        return BLANK_PRESETS[DEFAULT_BLANK_KEY]

    def add_blank():
        w_cm, h_cm = blank_size()
        writer.add_blank_page(width=cm_to_pt(w_cm), height=cm_to_pt(h_cm))

    def side_of(task):
        """取一个文件任务实际生效的面向：显式标记 > 默认面"""
        raw = task[3] if len(task) > 3 else None
        return raw or default_side

    for task in segment_tasks:
        if task[0] == 'blank':
            _, w_cm, h_cm = task
            writer.add_blank_page(width=cm_to_pt(w_cm), height=cm_to_pt(h_cm))
            continue

        _, file_path, copies = task[:3]
        base_name = os.path.basename(file_path)

        if file_path not in pdf_cache:
            progress(f"转换中：{base_name}")
            temp_pdf = os.path.join(temp_dir, f"temp_{len(pdf_cache)}.pdf")
            try:
                result = convert_to_pdf(file_path, temp_pdf)
            except Exception as e:
                failed.append(f"{base_name}：{e}")
                continue
            if not os.path.isfile(result) or os.path.getsize(result) == 0:
                failed.append(f"{base_name}：转换结果为空")
                continue
            pdf_cache[file_path] = result

        src_pdf = pdf_cache[file_path]

        # 这个文件该按单面还是双面处理：
        # auto_pad 开启时以**文件自己的标记**为准；否则退回段级的 simplex 开关
        if auto_pad:
            file_simplex = (side_of(task) == SIDE_SIMPLEX)
        else:
            file_simplex = simplex

        # 每份都重新打开一次，保证重复内容完整且不共用 reader 缓存
        for n in range(copies):
            try:
                reader = PdfReader(src_pdf)
                page_count = len(reader.pages)

                if file_simplex:
                    # 单面：逐页写入，每页后面跟一张空白，让内容都落在正面
                    for page in reader.pages:
                        writer.add_page(page)
                        add_blank()
                else:
                    writer.append(reader)
                    # 双面：这一份页数为奇数时补 1 页，
                    # 免得下一份/下一个文件印到它的背面
                    if auto_pad and page_count % 2 == 1:
                        add_blank()
            except Exception as e:
                failed.append(f"{base_name}（第 {n + 1} 份）：{e}")
                break

    if len(writer.pages) == 0:
        w, h = BLANK_PRESETS[DEFAULT_BLANK_KEY]
        writer.add_blank_page(width=cm_to_pt(w), height=cm_to_pt(h))

    # 段级末尾补页（一般已由上面的文件级逻辑处理，pad_pages 传 0）
    for _ in range(max(0, int(pad_pages))):
        w_cm, h_cm = blank_size()
        writer.add_blank_page(width=cm_to_pt(w_cm), height=cm_to_pt(h_cm))

    try:
        with open(out_path, 'wb') as f:
            writer.write(f)
        pages = len(writer.pages)
    finally:
        writer.close()
    return pages


def estimate_segment_pads(segment_tasks, pdf_cache, failed, temp_dir,
                          default_side=None, progress=None):
    """预估这一段会插入多少张空白页（与 write_segment 的规则保持一致）。

    规则：单面文件补「它自己的页数 × 份数」张；双面文件每份奇数页补 1 张。
    仅用于结果汇报。
    """
    progress = progress or _noop_progress
    total = 0

    for task in segment_tasks:
        if task[0] != 'file':
            continue
        _, file_path, copies = task[:3]
        raw_side = task[3] if len(task) > 3 else None
        side = raw_side or default_side

        if file_path in pdf_cache:
            src_pdf = pdf_cache[file_path]
        else:
            base_name = os.path.basename(file_path)
            progress(f"统计中：{base_name}")
            temp_pdf = os.path.join(temp_dir, f"temp_{len(pdf_cache)}.pdf")
            try:
                result = convert_to_pdf(file_path, temp_pdf)
            except Exception as e:
                failed.append(f"{base_name}：{e}")
                continue
            if not os.path.isfile(result) or os.path.getsize(result) == 0:
                failed.append(f"{base_name}：转换结果为空")
                continue
            pdf_cache[file_path] = result
            src_pdf = result

        try:
            n = len(PdfReader(src_pdf).pages)
        except Exception as e:
            failed.append(f"{os.path.basename(file_path)}：{e}")
            continue

        if side == SIDE_SIMPLEX:
            # 每页后面插一张空白
            total += n * copies
        elif n % 2 == 1:
            # 双面：每份奇数页补 1 张
            total += copies

    return total


def write_outputs(out_dir, segments, pdf_cache, failed, temp_dir,
                  user_name='', default_single='合并输出_打印预览',
                  auto_pad=False, pad_size=None, default_side=None,
                  progress=None):
    """把各段写盘。返回 (结果列表, 错误列表)。
    结果列表: [(文件名, 页数, side, pad), ...]
    错误列表: [(文件名, 异常对象, 输出路径, 段任务), ...] —— 调用方决定是否降级重试

    命名优先级（从高到低）：
      1. 分隔符上写的名字（右键「重命名这一段」）
      2. 界面上填的「输出文件名」（支持用 # 做序号占位符）
      3. 默认名：单段用 default_single，多段用 分卷NN

    auto_pad: 启用「自动分叠」时，**逐个文件**按它自己的单/双面标记处理：
              · 单面文件：每页内容后面插一张空白页，让内容全部落在
                        纸张正面、背面留白（配合「全程双面打印」）
              · 双面文件：它自己页数为奇数时末尾补 1 页，凑成偶数让它
                        独占整张纸、从正面开始，不会和相邻内容挤在一张纸上
              注意粒度是**文件**而不是段：一个文件标了单面，只影响它自己。
    """
    results = []
    errors = []
    used = set()

    # 先按段把面向理清（仅用于结果汇报）：分隔符上没写的，用段内第一个文件的标记
    seg_sides = []
    for doc_name, segment_tasks, seg_side in segments:
        side = seg_side
        if side is None:
            for t in segment_tasks:
                if t[0] == 'file' and len(t) > 3 and t[3]:
                    side = t[3]
                    break
        seg_sides.append(side)

    total = len(segments)
    for idx, (doc_name, segment_tasks, _side) in enumerate(segments, start=1):
        if doc_name:
            # 1) 分隔符自带的名字优先
            base = sanitize_filename(doc_name)
        else:
            # 2) 界面上的自定义文件名
            base = resolve_output_name(user_name, idx, total)

        if not base:
            # 3) 默认名
            base = default_single if total == 1 else f"分卷{idx:02d}"

        filename = base + ".pdf"
        # 重名保护：本次任务内部去重，同时避开目录里已有的同名文件
        dup = 2
        while filename.lower() in used or os.path.exists(os.path.join(out_dir, filename)):
            filename = f"{base}_{dup}.pdf"
            dup += 1
            if dup > 999:
                break
        used.add(filename.lower())
        out_path = os.path.join(out_dir, filename)

        # 统计这一段会插入多少张空白页（仅用于结果汇报）。
        # 真正插页在 write_segment 里按**逐个文件**的标记完成。
        pad = 0
        if auto_pad:
            try:
                pad = estimate_segment_pads(segment_tasks, pdf_cache, failed,
                                            temp_dir, default_side, progress)
            except Exception:
                pad = 0

        try:
            pages = write_segment(segment_tasks, out_path, pdf_cache, failed,
                                  temp_dir, pad_pages=0, pad_size=pad_size,
                                  auto_pad=auto_pad, default_side=default_side,
                                  progress=progress)
            results.append((filename, pages, seg_sides[idx - 1], pad))
        except Exception as e:
            errors.append((filename, e, out_path, segment_tasks))

    return results, errors


if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle('Fusion') 
    ex = PDFMergerApp()
    ex.show()
    sys.exit(app.exec_())