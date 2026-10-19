import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import threading
import os
import json
import re
import glob
import shutil
import datetime
import urllib.request
from pypdf import PdfReader
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ============================================================
#  配置区
# ============================================================
API_KEY = ""          # 程序启动时会从本地配置文件读取（界面里可以填/改）

FONT = "微软雅黑"
COLOR_WARN = RGBColor(0xC0, 0x00, 0x00)
COLOR_GRAY = RGBColor(0x60, 0x60, 0x60)
COLOR_HEAD = RGBColor(0x1F, 0x3B, 0x6B)
COLOR_SRC = RGBColor(0x70, 0x70, 0x70)

MAX_TOKENS = 8000


# ============================================================
#  本地配置管理（保存 API Key）
#  路径自动适配：Windows 用 %APPDATA%，其他系统用用户主目录
# ============================================================
APP_NAME = "投标决策助手"
CONFIG_FILE = "config.json"


def get_config_dir():
    """拿到配置目录（不写死路径，自动适配任何电脑和用户）"""
    base = os.environ.get("APPDATA")          # Windows: C:\Users\xxx\AppData\Roaming
    if not base:
        base = os.path.expanduser("~")        # 兜底：用户主目录
    folder = os.path.join(base, APP_NAME)
    os.makedirs(folder, exist_ok=True)
    return folder


def get_config_path():
    return os.path.join(get_config_dir(), CONFIG_FILE)


def load_config():
    path = get_config_path()
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def save_config(cfg):
    path = get_config_path()
    with open(path, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)
    return path


# ============================================================
#  发动机
# ============================================================
def ask_ai(messages):
    url = "https://api.deepseek.com/chat/completions"
    data = {
        "model": "deepseek-chat",
        "messages": messages,
        "max_tokens": MAX_TOKENS
    }
    body = json.dumps(data).encode("utf-8")
    req = urllib.request.Request(url, data=body, method="POST")
    req.add_header("Content-Type", "application/json")
    req.add_header("Authorization", "Bearer " + API_KEY)
    with urllib.request.urlopen(req) as resp:
        result = json.loads(resp.read().decode("utf-8"))
    return result["choices"][0]["message"]["content"]


# ============================================================
#  读文件（pdf / docx / txt）
# ============================================================
def read_docx_file(path):
    doc = Document(path)
    parts = []
    for p in doc.paragraphs:
        if p.text.strip():
            parts.append(p.text.strip())
    for table in doc.tables:
        for row in table.rows:
            cells = [c.text.strip().replace("\n", " ") for c in row.cells]
            line = " | ".join(cells)
            if line.replace("|", "").strip():
                parts.append(line)
    return "\n".join(parts)


def read_any_file(path):
    lower = path.lower()
    if lower.endswith(".pdf"):
        reader = PdfReader(path)
        text = ""
        for page in reader.pages:
            text = text + page.extract_text()
        if len(text.strip()) < 100:
            return "【提示】这份PDF可能是扫描件，无法提取文字，需要OCR识别。"
    elif lower.endswith(".docx"):
        text = read_docx_file(path)
    elif lower.endswith(".doc"):
        return "【提示】这是旧版 .doc 格式，请先用 WPS/Word 另存为 .docx 后再运行。"
    else:
        try:
            with open(path, "r", encoding="utf-8") as f:
                text = f.read()
        except UnicodeDecodeError:
            with open(path, "r", encoding="gbk", errors="ignore") as f:
                text = f.read()
    return text


# ============================================================
#  工程量清单：代码直接提取（不经过AI）
# ============================================================
def collect_boq(cache_dir):
    pat = re.compile(
        r"(\d{1,4})\s*\|\s*(\d{9,14})\s*\|\s*([^|\n]+?)\s*\|\s*([^|\n]+?)\s*\|\s*([\d,\.]+)"
    )
    items = []
    seen = set()
    for fp in sorted(glob.glob(os.path.join(cache_dir, "块*.txt"))):
        with open(fp, "r", encoding="utf-8") as f:
            for line in f:
                m = pat.search(line)
                if m:
                    code = m.group(2).strip()
                    if code not in seen:
                        seen.add(code)
                        items.append([
                            "",
                            code,
                            m.group(3).strip(),
                            m.group(4).strip(),
                            m.group(5).strip()
                        ])
    return items


# ============================================================
#  Word 格式工具
# ============================================================
def clean(s):
    s = str(s)
    for ch in ["**", "`", "###", "##", "&gt;"]:
        s = s.replace(ch, "")
    s = s.replace("&amp;", "&")
    return s.strip()


def set_run(run, size=10.5, bold=False, color=None):
    run.font.name = FONT
    run.font.size = Pt(size)
    run.bold = bold
    if color is not None:
        run.font.color.rgb = color
    run._element.rPr.rFonts.set(qn('w:eastAsia'), FONT)


def add_para(doc, text, size=10.5, bold=False, color=None,
             align=None, space_after=4, indent=0):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(space_after)
    if indent:
        p.paragraph_format.left_indent = Pt(indent)
    r = p.add_run(clean(text))
    set_run(r, size, bold, color)
    return p


def add_para_mixed(doc, number, text, size=10.5, indent=12, space_after=2):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.left_indent = Pt(indent)

    src = ""
    body = text
    idx = text.rfind("（出处：")
    if idx == -1:
        idx = text.rfind("(出处：")
    if idx != -1:
        body = text[:idx].rstrip()
        src = text[idx:]

    is_key = body.startswith("⚠️")

    r1 = p.add_run(f"{number}. ")
    set_run(r1, size, False, None)
    r2 = p.add_run(clean(body))
    set_run(r2, size, is_key, COLOR_WARN if is_key else None)
    if src:
        r3 = p.add_run("  " + clean(src))
        set_run(r3, size - 1, False, COLOR_SRC)
    return p


def shade_cell(cell, fill="D9E2F3"):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill)
    tcPr.append(shd)


def style_table(table):
    src_col = None
    if table.rows:
        headers = [c.text.strip() for c in table.rows[0].cells]
        for i, h in enumerate(headers):
            if "出处" in h:
                src_col = i
                break
    for ri, row in enumerate(table.rows):
        for ci, cell in enumerate(row.cells):
            if ri == 0:
                shade_cell(cell)
            for p in cell.paragraphs:
                p.paragraph_format.space_before = Pt(1)
                p.paragraph_format.space_after = Pt(1)
                for r in p.runs:
                    if ri == 0:
                        set_run(r, size=10, bold=True)
                    elif src_col is not None and ci == src_col:
                        set_run(r, size=9, bold=False, color=COLOR_SRC)
                    else:
                        set_run(r, size=10, bold=False)


# ============================================================
#  文件占用检测 + 自动备份 + 安全保存
# ============================================================
def is_locked(path):
    """检查文件是否被 WPS/Word 等程序占用"""
    if not os.path.exists(path):
        return False
    try:
        with open(path, "a+b"):
            return False
    except PermissionError:
        return True
    except OSError:
        return False


def backup_existing(path, log, keep=3):
    """目标已存在 → 改名成"旧版备份"；只保留最近 keep 个"""
    if not os.path.exists(path):
        return

    folder = os.path.dirname(path)
    name = os.path.splitext(os.path.basename(path))[0]
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = os.path.join(folder, f"{name}_旧版备份_{stamp}.docx")

    try:
        os.replace(path, backup_path)
        log(f"📦 检测到已存在同名结果，已备份为：{os.path.basename(backup_path)}")
    except Exception as e:
        log(f"⚠️ 备份失败（{e}），将尝试直接覆盖")

    try:
        pattern = os.path.join(folder, f"{name}_旧版备份_*.docx")
        olds = sorted(glob.glob(pattern))
        for old in olds[:-keep]:
            os.remove(old)
            log(f"🧹 已清理更早的备份：{os.path.basename(old)}")
    except Exception:
        pass


def safe_save_docx(doc, out_path, log, keep_backups=3):
    """
    ① 生成到临时文件 ② 目标存在则先备份 ③ 写回固定文件名
    ④ 被占用 → 另存为"xxx_最新_时间戳.docx"
    """
    tmp_path = out_path + ".saving.docx"

    try:
        doc.save(tmp_path)
    except Exception as e:
        log(f"❌ 生成文档失败：{e}")
        return None, False

    folder = os.path.dirname(out_path)
    name = os.path.splitext(os.path.basename(out_path))[0]

    if os.path.exists(out_path):
        backup_existing(out_path, log, keep_backups)

    try:
        os.replace(tmp_path, out_path)
        try:
            for old in glob.glob(os.path.join(folder, f"{name}_最新_*.docx")):
                os.remove(old)
        except Exception:
            pass
        return out_path, False
    except (PermissionError, OSError):
        pass

    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    new_path = os.path.join(folder, f"{name}_最新_{stamp}.docx")

    try:
        os.replace(tmp_path, new_path)
    except Exception as e:
        log(f"❌ 保存失败：{e}")
        return None, False

    log("⚠️ 结果文件正被 WPS/Word 打开，无法覆盖原文件")
    log(f"   已另存为最新版：{os.path.basename(new_path)}")
    log("   （下次运行前先关掉 WPS/Word 里打开的旧文件，就能写回原文件名）")
    return new_path, True


# ============================================================
#  提示词
# ============================================================
PROMPT_EXTRACT = "你是招标文件信息提取助手。请从用户给出的文本片段中，尽可能完整地提取以下所有信息，并且【每一条都要标注它在原文中的位置（出处）】，格式如“（出处：第二章 投标人须知前附表）”。原文照抄，不要改写，不要编造：\n1. 项目基本信息：项目名称、招标人、招标代理机构、招标编号、建设地点、建设规模、招标范围、标段划分\n2. 资格要求：企业资质、项目负责人资格、业绩要求、财务要求、是否接受联合体、其他资格限制条件（只提取“投标人须具备什么条件”的内容，不要混入开标参会要求）\n3. 时间节点：招标文件获取时间、投标截止时间、保证金缴纳截止时间、投标有效期、投标人提问截止、澄清答疑时间\n4. 开标安排：开标时间、开标地点、参会要求、开标方式\n5. 费用与付款：最高投标限价、投标保证金金额与缴纳方式、履约保证金、预付款比例、进度款比例、结算款比例、质保金比例与期限、质量保修期、暂列金额、暂估价、不可竞争费用\n6. 工期：计划工期、计划开竣工日期、工期延误违约金\n7. 废标条款：列出全部“否决投标/废标”条款（保留原文编号）\n8. 不予受理情形：逾期送达、未送达指定地点、未按要求加密或密封标记等\n9. 评标办法：评标方法、评分标准及各分值、评标基准价计算方式、中标候选人确定方式\n10. 工程量清单：请严格按照“序号 | 项目编码 | 项目名称 | 计量单位 | 工程量”的格式逐行输出，一条都不能少，不要省略、不要写“等等”\n11. 投标文件要求：组成材料、份数、加密与密封要求、编制要求、递交与撤回\n12. 中标后风险条款：违约责任、工期违约金、分包限制、工程变更、质量标准、人员到岗、农民工工资等\n【重要】如果这段文本中有工程量清单表格，必须逐行完整输出清单的每一行，这是硬性要求。\n如果这段文本中确实没有上述任何一项信息，才回答：本段无关键信息。但只要出现了资质、时间、金额、比例、条款、清单等任何相关内容，就必须逐项提取出来，不得省略。不要使用**等markdown符号。"

PROMPT_MERGE_GROUP = "你是招标文件信息整理助手。用户会给你几段从同一份招标文件中提取的零散信息，请把它们合并，并严格遵守：\n1. 必须去重：同一件事只能保留一条，不要因为措辞不同就当成两件事\n2. 保留每条信息末尾的（出处：xxx）标注\n3. 工程量清单的行（格式为“序号 | 项目编码 | 项目名称 | 单位 | 数量”）必须原样完整保留，一行都不能删、不能合并\n4. 只输出整理后的内容，不要编造，不要使用**等markdown符号"

PROMPT_MERGE_FINAL = "你是投标决策助手。请把用户提供的零散信息，整理成严格符合以下结构的JSON，不要输出任何其他文字，不要用```包裹：\n{\n \"标题\": \"投标决策要点清单\",\n \"副标题信息\": [{\"项目\": \"...\", \"内容\": \"...\"}],\n \"章节\": [\n  {\"标题\": \"【一、资格门槛】\", \"提示\": \"⚠️ ...\", \"类型\": \"表\", \"表头\": [\"自查项\",\"招标文件要求\",\"注意\",\"出处\"], \"行\": [[\"企业资质\",\"...\",\"⚠️ 常见门槛\",\"第二章 前附表\"]]},\n  {\"标题\": \"【四、废标红线】\", \"提示\": \"⚠️ ...\", \"类型\": \"列表\", \"内容\": [\"⚠️ ...（出处：第三章）\"]}\n ],\n \"行动清单\": [\"9月26日前：购买招标文件\"]\n}\n\n要求：\n1.「类型」只能是 \"表\" 或 \"列表\"\n2.「表」类型必须给出「表头」和「行」：「表头」是列名数组，「行」是二维数组，每行列数必须与表头一致\n3.「表」的最后一列必须是「出处」列；无法确定填“—”\n4.「列表」类型的「内容」是字符串数组，每条末尾用（出处：xxx）标注\n5. 每个章节必须有「提示」字段；有风险的以 ⚠️ 开头\n6. 章节顺序依次为：【一、资格门槛】（表，表头：自查项/招标文件要求/注意/出处）、【二、关键时间节点】（表，表头：时间/事项/要做什么/出处）、【三、费用与现金流】（表，表头：项目/金额或比例/注意/出处）、【四、废标红线】（列表）、【五、不予受理情形】（列表）、【六、怎么中标】（表，表头：项目/内容/出处）、【七、中标后的风险条款】（表，表头：条款/内容/出处）、【八、投标文件编制要求】（列表）\n7.【极其重要】不要在JSON里输出工程量清单的具体内容！工程量清单由程序另行处理，你完全跳过它。章节里不要出现工程量清单这一节\n8. 必须去重：同一件事只能出现一次，跨章节也不得重复\n9. ⚠️ 标记必须克制，但条款不得删减：【四、废标红线】和【五、不予受理情形】这两个列表，必须把原文中的全部条款一条不漏地列出（通常是十几条），其中只有“最容易忽略且后果最严重”的 3-5 条加 ⚠️ 标记，其余条款照常列出但不加 ⚠️。绝对不允许为了控制⚠️数量而删除条款\n10. 表格中凡是“容易出错”“会导致废标”的内容，在“注意”列简要说明原因\n11.「副标题信息」必须包含：项目名称、招标人、招标编号、最高投标限价、计划工期\n12.「行动清单」按时间顺序排列，格式为“X月X日前：具体要做的事”\n13. 关键数字必须准确完整\n14. 没有信息的章节，「内容」写 [\"未找到\"]\n15. 所有字符串中禁止出现 ** 和 ` 和 # 等markdown符号，字符串内不要换行\n16. 所有语句必须通顺、符合中文表达习惯，例如应写“不得超过最高投标限价，否则废标”，不要写“超过最高投标限价，否则废标”\n17. 只输出JSON本身"


# ============================================================
#  渲染成 Word（只生成，不保存）
# ============================================================
def render_boq_table(doc, items):
    if not items:
        add_para(doc, "（未从原文中提取到工程量清单）", size=10.5, indent=12)
        return
    headers = ["序号", "项目编码", "项目名称", "计量单位", "工程量"]
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    for i, h in enumerate(headers):
        table.rows[0].cells[i].text = h
    for n, it in enumerate(items, 1):
        cells = table.add_row().cells
        cells[0].text = str(n)
        for i in range(1, len(headers)):
            cells[i].text = clean(it[i])
    style_table(table)


def build_doc(data, FILE_PATH, boq_items):
    doc = Document()

    add_para(doc, data.get("标题", "投标决策要点清单"),
             size=18, bold=True, color=COLOR_HEAD,
             align=WD_ALIGN_PARAGRAPH.CENTER, space_after=14)

    for item in data.get("副标题信息", []):
        add_para(doc, f'{item.get("项目","")}：{item.get("内容","")}', size=11)

    add_para(doc, "", size=6, space_after=2)
    add_para(doc, f'生成日期：{datetime.date.today()}', size=9, color=COLOR_GRAY, space_after=1)
    add_para(doc, f'数据来源：{os.path.basename(FILE_PATH)}', size=9, color=COLOR_GRAY, space_after=1)
    add_para(doc, "⚠️ 本清单由 AI 自动生成，关键信息请务必与招标文件原文核对",
             size=9, bold=True, color=COLOR_WARN, space_after=12)

    for sec in data.get("章节", []):
        add_para(doc, sec.get("标题", ""), size=13, bold=True,
                 color=COLOR_HEAD, space_after=2)

        tip = sec.get("提示", "")
        if tip:
            add_para(doc, tip, size=10, bold=True, color=COLOR_WARN,
                     indent=12, space_after=6)

        typ = str(sec.get("类型", "列表")).strip()

        if typ == "表":
            headers = sec.get("表头", [])
            rows = sec.get("行", [])
            if headers:
                table = doc.add_table(rows=1, cols=len(headers))
                table.style = "Table Grid"
                for i, h in enumerate(headers):
                    table.rows[0].cells[i].text = clean(h)
                for row in rows:
                    cells = table.add_row().cells
                    for i, v in enumerate(row):
                        if i < len(cells):
                            cells[i].text = clean(v)
                style_table(table)
                add_para(doc, "", size=6, space_after=8)
        else:
            for i, line in enumerate(sec.get("内容", []), 1):
                add_para_mixed(doc, i, clean(line), size=10.5,
                               indent=12, space_after=2)
            add_para(doc, "", size=6, space_after=8)

    add_para(doc, f"【九、工程量清单】（共 {len(boq_items)} 条）", size=13, bold=True,
             color=COLOR_HEAD, space_after=2)
    add_para(doc, "⚠️ 本表由程序直接从原文提取，未经过 AI 改写；报价前请与原文清单逐项核对，防止漏项",
             size=10, bold=True, color=COLOR_WARN, indent=12, space_after=6)
    render_boq_table(doc, boq_items)
    add_para(doc, "", size=6, space_after=8)

    actions = data.get("行动清单", [])
    if actions:
        add_para(doc, "【十、行动清单】建议照着做", size=13, bold=True,
                 color=COLOR_HEAD, space_after=2)
        add_para(doc, "⚠️ 按时间倒排，越早动手越安全", size=10, bold=True,
                 color=COLOR_WARN, indent=12, space_after=6)
        for a in actions:
            add_para(doc, "□ " + clean(a), size=11, indent=12, space_after=3)

    return doc


# ============================================================
#  主流程（供界面调用）
# ============================================================
def process_file(FILE_PATH, log, stop_event, on_progress=None):
    def prog(pct):
        if on_progress:
            on_progress(pct)

    log("=" * 52)
    log(f"待处理文件：{FILE_PATH}")

    if not os.path.exists(FILE_PATH):
        log("❌ 找不到这个文件，请检查路径")
        return None

    try:
        text = read_any_file(FILE_PATH)
    except Exception as e:
        log(f"❌ 读取文件失败：{e}")
        return None

    log(f"读出来一共 {len(text)} 个字")

    if len(text.strip()) < 100:
        log("⚠️ 读出来的内容太少，程序已停止。")
        log("可能原因：扫描件PDF / 文件是空的 / 格式不支持")
        return None

    folder = os.path.dirname(FILE_PATH)
    base = os.path.splitext(os.path.basename(FILE_PATH))[0]

    lines = text.split("\n")
    log(f"一共 {len(lines)} 行")

    out_docx = os.path.join(folder, base + "_投标决策要点清单.docx")
    if is_locked(out_docx):
        log("⚠️ 提示：上次的结果文件正被 WPS/Word 打开，本次将自动另存为最新版")
        log("   建议先关闭它，这样可以写回原文件名（不影响继续运行）")

    # ---------- 切分 ----------
    chunk_size = 3000
    chunks = []
    current = ""
    for line in lines:
        if len(current) + len(line) > chunk_size:
            chunks.append(current)
            current = ""
        current = current + line + "\n"
    if current != "":
        chunks.append(current)

    log(f"一共切成了 {len(chunks)} 块")
    log(f"预计需要 {len(chunks) + 8} 次AI调用，请耐心等待")
    log("-" * 52)

    cache_dir = os.path.join(folder, base + "_缓存")
    os.makedirs(cache_dir, exist_ok=True)

    total_steps = len(chunks) + 8
    done_steps = 0

    # ---------- 第一遍：逐块提取 ----------
    all_results = []

    for idx, chunk in enumerate(chunks, 1):
        if stop_event.is_set():
            log("⏹ 已取消")
            return None

        cache_file = os.path.join(cache_dir, f"块{idx:03d}.txt")
        raw_file = os.path.join(cache_dir, f"原文{idx:03d}.txt")

        if not os.path.exists(raw_file):
            with open(raw_file, "w", encoding="utf-8") as f:
                f.write(chunk)

        try:
            if os.path.exists(cache_file):
                with open(cache_file, "r", encoding="utf-8") as f:
                    one = f.read()
                log(f"[{idx}/{len(chunks)}] 已缓存，跳过")
            else:
                history = [
                    {"role": "system", "content": PROMPT_EXTRACT},
                    {"role": "user", "content": chunk}
                ]
                one = ask_ai(history)
                if len(chunk) > 500 and "无关键信息" in one:
                    log(f"[{idx}/{len(chunks)}] ⚠️ 疑似失败，重试...")
                    one = ask_ai(history)
                with open(cache_file, "w", encoding="utf-8") as f:
                    f.write(one)
                log(f"[{idx}/{len(chunks)}] 完成，提取 {len(one)} 字")
        except Exception as e:
            log(f"[{idx}/{len(chunks)}] ❌ 调用失败：{e}")
            one = ""

        all_results.append(one)
        done_steps += 1
        prog(int(done_steps * 100 / total_steps))

    # ---------- 代码直接抓清单 ----------
    boq_items = collect_boq(cache_dir)
    log(f"📋 程序直接提取到工程量清单：{len(boq_items)} 条（未经AI，不丢失）")

    # ---------- 分组汇总 ----------
    group_size = 8
    groups = [all_results[i:i + group_size] for i in range(0, len(all_results), group_size)]
    log("-" * 52)
    log(f"开始分组汇总：{len(all_results)} 块 → {len(groups)} 组")

    merged_groups = []
    for gi, g in enumerate(groups, 1):
        if stop_event.is_set():
            log("⏹ 已取消")
            return None

        cache_file = os.path.join(cache_dir, f"汇总组{gi:02d}.txt")
        try:
            if os.path.exists(cache_file):
                with open(cache_file, "r", encoding="utf-8") as f:
                    m = f.read()
                log(f"[组{gi}/{len(groups)}] 已缓存，跳过")
            else:
                history = [
                    {"role": "system", "content": PROMPT_MERGE_GROUP},
                    {"role": "user", "content": "\n".join(g)}
                ]
                m = ask_ai(history)
                with open(cache_file, "w", encoding="utf-8") as f:
                    f.write(m)
                log(f"[组{gi}/{len(groups)}] 汇总完成，{len(m)} 字")
        except Exception as e:
            log(f"[组{gi}/{len(groups)}] ❌ 失败：{e}")
            m = ""

        merged_groups.append(m)
        done_steps += 1
        prog(int(done_steps * 100 / total_steps))

    # ---------- 最终汇总成 JSON ----------
    log("-" * 52)
    log("正在生成最终结构化结果（这一步可能要1-2分钟）...")

    if stop_event.is_set():
        log("⏹ 已取消")
        return None

    try:
        history = [
            {"role": "system", "content": PROMPT_MERGE_FINAL},
            {"role": "user", "content": "\n\n".join(merged_groups)}
        ]
        final = ask_ai(history)
    except Exception as e:
        log(f"❌ 最终汇总失败：{e}")
        return None

    done_steps += 1
    prog(int(done_steps * 100 / total_steps))

    # ---------- 解析 JSON ----------
    raw = final.strip()
    if raw.startswith("```"):
        raw = raw.split("\n", 1)[1].rsplit("```", 1)[0]

    try:
        data = json.loads(raw)
    except Exception as e:
        log(f"❌ JSON 解析失败：{e}")
        fallback = os.path.join(folder, base + "_整理结果_原始.txt")
        with open(fallback, "w", encoding="utf-8") as f:
            f.write(final)
        log(f"原始返回内容已存到：{fallback}")
        return None

    # ---------- 渲染 + 保存 ----------
    log("正在渲染 Word 文档...")
    try:
        doc = build_doc(data, FILE_PATH, boq_items)
    except Exception as e:
        log(f"❌ 生成Word失败：{e}")
        return None

    saved_path, renamed = safe_save_docx(doc, out_docx, log)
    if not saved_path:
        return None

    prog(100)
    log("=" * 52)
    log("✅ 处理完成！")
    log(f"📄 结果文件：{saved_path}")
    log(f"📋 工程量清单：{len(boq_items)} 条")
    return saved_path


# ============================================================
#  【新增】清理缓存对话框
# ============================================================
class CacheDialog(tk.Toplevel):
    def __init__(self, parent, caches, log):
        super().__init__(parent)
        self.title("清理缓存")
        self.caches = caches
        self.log = log

        self.transient(parent)
        self.grab_set()
        self.resizable(False, False)

        tk.Label(self, text="检测到该文件夹下有如下缓存：",
                 font=(FONT, 10, "bold")).pack(anchor="w", padx=16, pady=(14, 2))
        tk.Label(self,
                 text="缓存只用于加速重跑，删除不影响已生成的报告；\n但下次重跑会重新调用AI，会慢一些、也会再花一点钱。",
                 font=(FONT, 9), fg="#666666", justify="left").pack(anchor="w", padx=16, pady=(0, 10))

        body = tk.Frame(self)
        body.pack(fill="both", expand=True, padx=16)

        self.vars = []
        for name, path, n in caches:
            var = tk.BooleanVar(value=True)
            self.vars.append((var, path, name, n))
            tk.Checkbutton(body, text=f"{name}    （{n} 个文件）",
                           variable=var, font=(FONT, 10),
                           anchor="w").pack(fill="x", pady=1)

        btns = tk.Frame(self)
        btns.pack(fill="x", padx=16, pady=14)

        tk.Button(btns, text="全选", command=lambda: self.set_all(True),
                  font=(FONT, 10), width=8).pack(side="left")
        tk.Button(btns, text="全不选", command=lambda: self.set_all(False),
                  font=(FONT, 10), width=8).pack(side="left", padx=6)

        tk.Button(btns, text="清理选中", command=self.do_clear,
                  font=(FONT, 10, "bold"), width=10,
                  bg="#C00000", fg="white").pack(side="right")
        tk.Button(btns, text="取消", command=self.destroy,
                  font=(FONT, 10), width=8).pack(side="right", padx=8)

        # 居中显示
        self.update_idletasks()
        x = parent.winfo_rootx() + (parent.winfo_width() - self.winfo_width()) // 2
        y = parent.winfo_rooty() + (parent.winfo_height() - self.winfo_height()) // 3
        self.geometry(f"+{max(x, 0)}+{max(y, 0)}")

    def set_all(self, value):
        for var, _, _, _ in self.vars:
            var.set(value)

    def do_clear(self):
        targets = [(p, name, n) for (v, p, name, n) in self.vars if v.get()]
        if not targets:
            messagebox.showwarning("提示", "请至少选择一个缓存", parent=self)
            return

        total_files = 0
        for p, name, n in targets:
            shutil.rmtree(p, ignore_errors=True)
            total_files += n
            self.log(f"🧹 已清理缓存：{name}（{n} 个文件）")

        messagebox.showinfo("完成",
                            f"已清理 {len(targets)} 个缓存文件夹，共 {total_files} 个文件。",
                            parent=self)
        self.destroy()


# ============================================================
#  界面
# ============================================================
class App:
    def __init__(self, root):
        self.root = root
        root.title("投标决策要点清单生成器")
        root.geometry("880x720")
        root.minsize(780, 620)

        # ==================== API Key 设置区 ====================
        fkey = tk.LabelFrame(root, text=" API Key 设置（只保存在你自己电脑上） ",
                             font=(FONT, 10), padx=8, pady=6)
        fkey.pack(fill="x", padx=14, pady=(12, 4))

        tk.Label(fkey, text="API Key：", font=(FONT, 10)).pack(side="left")

        self.key_var = tk.StringVar()
        self.key_entry = tk.Entry(fkey, textvariable=self.key_var,
                                  font=(FONT, 10), show="●", width=42)
        self.key_entry.pack(side="left", padx=6)

        self.show_key = tk.BooleanVar(value=False)
        tk.Checkbutton(fkey, text="显示", variable=self.show_key,
                       command=self.toggle_key, font=(FONT, 9)).pack(side="left")

        tk.Button(fkey, text="保存", command=self.save_key,
                  font=(FONT, 9), width=6).pack(side="left", padx=6)

        self.key_status = tk.Label(fkey, text="", font=(FONT, 9))
        self.key_status.pack(side="left", padx=6)

        # ==================== 文件选择区 ====================
        f1 = tk.Frame(root)
        f1.pack(fill="x", padx=14, pady=(8, 4))
        tk.Label(f1, text="招标文件：", font=(FONT, 11)).pack(side="left")
        self.path_var = tk.StringVar()
        tk.Entry(f1, textvariable=self.path_var, font=(FONT, 10)).pack(
            side="left", fill="x", expand=True, padx=8)
        tk.Button(f1, text="选择文件", command=self.choose_file,
                  font=(FONT, 10), width=10).pack(side="left")

        # ==================== 按钮区 ====================
        f2 = tk.Frame(root)
        f2.pack(fill="x", padx=14, pady=8)
        self.start_btn = tk.Button(f2, text="开始处理", command=self.start,
                                   font=(FONT, 12, "bold"), width=13, height=2,
                                   bg="#2E75B6", fg="white")
        self.start_btn.pack(side="left")
        self.stop_btn = tk.Button(f2, text="停止", command=self.stop,
                                  font=(FONT, 11), width=8, height=2, state="disabled")
        self.stop_btn.pack(side="left", padx=8)
        self.clear_btn = tk.Button(f2, text="清理缓存…", command=self.clear_cache,
                                   font=(FONT, 11), width=11, height=2)
        self.clear_btn.pack(side="left", padx=8)
        self.open_btn = tk.Button(f2, text="打开结果文件夹", command=self.open_folder,
                                  font=(FONT, 11), height=2, state="disabled")
        self.open_btn.pack(side="left", padx=8)

        # ==================== 进度条 ====================
        self.progress = ttk.Progressbar(root, mode="determinate", maximum=100)
        self.progress.pack(fill="x", padx=14, pady=(4, 6))

        # ==================== 日志区 ====================
        tk.Label(root, text="运行日志：", font=(FONT, 10)).pack(anchor="w", padx=14)
        self.log_box = scrolledtext.ScrolledText(
            root, height=20, font=("Consolas", 10), wrap="word")
        self.log_box.pack(fill="both", expand=True, padx=14, pady=(2, 6))

        tk.Label(root,
                 text="支持 PDF / Word(.docx) / TXT ｜ 文件不上传第三方，Key 只存在你本机",
                 font=(FONT, 9), fg="#666666").pack(pady=(0, 10))

        self.stop_event = threading.Event()
        self.result_path = None

        # 启动时读取已保存的 Key
        self.load_key()

    # ---------- Key 相关 ----------
    def toggle_key(self):
        self.key_entry.configure(show="" if self.show_key.get() else "●")

    def load_key(self):
        cfg = load_config()
        key = cfg.get("api_key", "").strip()
        if key:
            self.key_var.set(key)
            self.key_status.configure(text="✅ 已配置", fg="#2E7D32")
        else:
            self.key_status.configure(text="⚠️ 未配置，请填写后点保存", fg="#C00000")

    def save_key(self):
        key = self.key_var.get().strip()
        if not key:
            messagebox.showwarning("提示", "请输入 API Key")
            return
        try:
            path = save_config({"api_key": key})
        except Exception as e:
            messagebox.showerror("保存失败", str(e))
            return
        self.key_status.configure(text="✅ 已配置", fg="#2E7D32")
        messagebox.showinfo("已保存",
                            f"API Key 已保存到：\n{path}\n\n下次打开会自动读取，不用再填。")

    # ---------- 文件相关 ----------
    def choose_file(self):
        p = filedialog.askopenfilename(
            title="选择招标文件",
            filetypes=[("招标文件", "*.pdf *.docx *.txt"),
                       ("PDF", "*.pdf"),
                       ("Word", "*.docx"),
                       ("文本", "*.txt"),
                       ("所有文件", "*.*")]
        )
        if p:
            self.path_var.set(p)

    # ---------- 【改动】清理缓存：扫描全部，弹窗勾选 ----------
    def clear_cache(self):
        """扫描该文件夹下所有缓存，弹窗让用户勾选要清理哪些"""
        path = self.path_var.get().strip().strip('"')
        if not path:
            messagebox.showwarning("提示", "请先选择任意一个招标文件（用来定位文件夹）")
            return

        folder = os.path.dirname(path)
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning("提示", "文件夹路径无效")
            return

        # 扫描该文件夹下所有 "*_缓存" 文件夹
        caches = []
        try:
            for d in sorted(os.listdir(folder)):
                p = os.path.join(folder, d)
                if os.path.isdir(p) and d.endswith("_缓存"):
                    try:
                        n = len(os.listdir(p))
                    except Exception:
                        n = 0
                    caches.append((d, p, n))
        except Exception as e:
            messagebox.showerror("错误", f"扫描文件夹失败：{e}")
            return

        if not caches:
            messagebox.showinfo("提示", "该文件夹下没有找到缓存文件夹，无需清理")
            return

        CacheDialog(self.root, caches, self.log)

    # ---------- 日志与进度 ----------
    def log(self, msg):
        self.root.after(0, self._log_ui, msg)

    def _log_ui(self, msg):
        self.log_box.insert(tk.END, msg + "\n")
        self.log_box.see(tk.END)

    def set_progress(self, pct):
        self.root.after(0, lambda: self.progress.configure(value=pct))

    # ---------- 开始处理 ----------
    def start(self):
        global API_KEY

        key = self.key_var.get().strip()
        if not key:
            messagebox.showerror("错误", "请先在上方填写 API Key 并点『保存』")
            return
        API_KEY = key

        path = self.path_var.get().strip().strip('"')
        if not path:
            messagebox.showwarning("提示", "请先选择招标文件")
            return
        if not os.path.exists(path):
            messagebox.showerror("错误", "找不到这个文件，请重新选择")
            return

        self.log_box.delete("1.0", tk.END)
        self.progress.configure(value=0)
        self.start_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self.clear_btn.configure(state="disabled")
        self.open_btn.configure(state="disabled")
        self.stop_event.clear()
        self.result_path = None

        threading.Thread(target=self.worker, args=(path,), daemon=True).start()

    def worker(self, path):
        try:
            result = process_file(path, self.log, self.stop_event, self.set_progress)
            self.result_path = result
        except Exception as e:
            self.log(f"❌ 出现未预期的错误：{e}")
            self.result_path = None
        finally:
            self.root.after(0, self.finish)

    def finish(self):
        self.start_btn.configure(state="normal")
        self.stop_btn.configure(state="disabled")
        self.clear_btn.configure(state="normal")
        if self.result_path:
            self.open_btn.configure(state="normal")
            messagebox.showinfo("完成", "处理完成！\n\n结果文件：\n" + self.result_path)
        else:
            messagebox.showwarning("未完成", "处理没有完成，请查看日志。")

    def stop(self):
        if messagebox.askyesno("确认", "确定要停止处理吗？\n（已完成的缓存会保留，下次可继续）"):
            self.stop_event.set()
            self.log("⏹ 正在停止...")

    def open_folder(self):
        if self.result_path and os.path.exists(self.result_path):
            os.startfile(os.path.dirname(self.result_path))


# ============================================================
#  启动
# ============================================================
if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()