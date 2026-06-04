#!/usr/bin/env python3
# /home/bid/bid_api_server.py
# 标书AI生成服务 v3.0（标题层级编号/黑体宋体/首行缩进/1.5行距/中文符号）
import os
import uuid
import asyncio
import logging
import subprocess
from pathlib import Path
from datetime import datetime
from logging.handlers import RotatingFileHandler
from typing import Dict, Any, Optional, List

from fastapi import FastAPI, BackgroundTasks, HTTPException
from pydantic import BaseModel
import uvicorn

import pdfplumber
from openai import OpenAI
from langchain_community.document_loaders import Docx2txtLoader
from langchain_text_splitters import RecursiveCharacterTextSplitter
from langchain_community.vectorstores import Chroma
from langchain_community.embeddings import HuggingFaceEmbeddings
from langchain_core.documents import Document

# ==================== 配置 ====================
KB_DIR = "/home/bid/quote"
CHROMA_DIR = "/home/bid/chroma_quote"
PRODUCT_DIR = "/home/bid/product"
LOG_DIR = "/home/bid/logs"

os.makedirs(LOG_DIR, exist_ok=True)
os.makedirs(PRODUCT_DIR, exist_ok=True)

# ==================== 日志配置 ====================
LOG_FILE = os.path.join(LOG_DIR, "bid_api_server.log")

logger = logging.getLogger("bid_server")
logger.setLevel(logging.INFO)
logger.propagate = False

if not logger.handlers:
    file_handler = RotatingFileHandler(LOG_FILE, maxBytes=10*1024*1024, backupCount=5, encoding="utf-8")
    file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
    logger.addHandler(file_handler)

    console_handler = logging.StreamHandler()
    console_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
    logger.addHandler(console_handler)

# vLLM 配置
OPENAI_BASE_URL = os.getenv("OPENAI_BASE_URL", "http://10.10.130.107:7890/v1")
MODEL_NAME = os.getenv("MODEL_NAME", "Qwen/Qwen3-30B-A3B")
API_KEY="qanwen2025"

client = OpenAI(base_url=OPENAI_BASE_URL, api_key=API_KEY)

# ==================== 模块级初始化 ====================
logger.info("正在加载 BAAI/bge-m3 embedding 模型（离线模式）...")
EMBEDDINGS = HuggingFaceEmbeddings(
    model_name="BAAI/bge-m3",
    model_kwargs={"device": "cpu", "local_files_only": True},
    encode_kwargs={"normalize_embeddings": True}
)
logger.info("BAAI/bge-m3 加载完成")

app = FastAPI(title="Bid Document AI Server", version="3.0.0")

tasks: Dict[str, Dict[str, Any]] = {}
VECTORSTORE = None


# ==================== Pydantic 模型 ====================
class GenerateBidRequest(BaseModel):
    tender_path: str
    uuid: str

class TaskStatus(BaseModel):
    task_id: str
    status: str
    progress: int
    message: str
    result: Optional[Dict[str, Any]] = None
    created_at: str
    updated_at: str


# ==================== 中文标点修复 ====================
def fix_chinese_punctuation(text: str) -> str:
    """将英文标点替换为中文全角标点（在不误伤数字的情况下）"""
    replacements = [
        (', ', '，'), (',', '，'),
        ('; ', '；'), (';', '；'),
        (': ', '：'), (':', '：'),
        ('! ', '！'), ('!', '！'),
        ('? ', '？'), ('?', '？'),
        ('( ', '（'), ('( ', '（'),
        (') ', '）'), (')', '）'),
        # 英文引号 → 中文引号（简单替换）
        ('"', '“'),  # 第二次会被替换为”
    ]
    for old, new in replacements:
        text = text.replace(old, new)
    # 成对引号处理：单数位置 " 变"，双数位置 " 变"
    result = []
    quote_open = True
    for ch in text:
        if ch == '"':
            result.append('"' if quote_open else '"')
            quote_open = not quote_open
        else:
            result.append(ch)
    return ''.join(result)


# ==================== DOCX 格式化 ====================
def format_docx(docx_path: str):
    """用 python-docx 做中文排版：黑体标题/宋体正文/首行缩进/1.5 行距"""
    from docx import Document
    from docx.shared import Pt, Cm
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    doc = Document(docx_path)

    def set_run_font(run, font_name, font_size_pt, bold=False):
        """设置 run 的字体（同时设置中/英文字体）"""
        run.font.name = font_name
        run.font.size = Pt(font_size_pt)
        run.font.bold = bold
        rPr = run._element.get_or_add_rPr()
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is None:
            rFonts = OxmlElement('w:rFonts')
            rPr.insert(0, rFonts)
        rFonts.set(qn('w:eastAsia'), font_name)
        rFonts.set(qn('w:ascii'), font_name)
        rFonts.set(qn('w:hAnsi'), font_name)

    def set_paragraph_format(para, first_line_indent=True, line_spacing=1.5):
        """设置段落格式"""
        pf = para.paragraph_format
        pf.line_spacing = line_spacing
        if first_line_indent:
            pf.first_line_indent = Cm(0.74)  # 约2个中文字符
        else:
            pf.first_line_indent = Cm(0)

    # ---------- 修改 Normal 样式 ----------
    normal_style = doc.styles['Normal']
    normal_style.font.name = '宋体'
    normal_style.font.size = Pt(12)
    nrPr = normal_style.element.get_or_add_rPr()
    nrFonts = nrPr.find(qn('w:rFonts'))
    if nrFonts is None:
        nrFonts = OxmlElement('w:rFonts')
        nrPr.insert(0, nrFonts)
    nrFonts.set(qn('w:eastAsia'), '宋体')
    nrFonts.set(qn('w:ascii'), '宋体')
    nrFonts.set(qn('w:hAnsi'), '宋体')
    set_paragraph_format(normal_style, first_line_indent=True)

    # ---------- 修改 Heading 样式 ----------
    heading_config = {
        'Heading 1': {'font': '黑体', 'size': 16, 'bold': True,  'indent': False},
        'Heading 2': {'font': '黑体', 'size': 14, 'bold': True,  'indent': False},
        'Heading 3': {'font': '黑体', 'size': 13, 'bold': True,  'indent': False},
    }
    for style_name, cfg in heading_config.items():
        if style_name in [s.name for s in doc.styles]:
            hs = doc.styles[style_name]
            hs.font.name = cfg['font']
            hs.font.size = Pt(cfg['size'])
            hs.font.bold = cfg['bold']
            hs.font.color.rgb = None
            hrPr = hs.element.get_or_add_rPr()
            hrFonts = hrPr.find(qn('w:rFonts'))
            if hrFonts is None:
                hrFonts = OxmlElement('w:rFonts')
                hrPr.insert(0, hrFonts)
            hrFonts.set(qn('w:eastAsia'), cfg['font'])
            hrFonts.set(qn('w:ascii'), cfg['font'])
            hrFonts.set(qn('w:hAnsi'), cfg['font'])
            set_paragraph_format(hs, first_line_indent=cfg['indent'])

    # ---------- 逐段强制字体修正 ----------
    for para in doc.paragraphs:
        style_name = para.style.name if para.style else ''
        is_heading = style_name.startswith('Heading')

        if is_heading:
            level = style_name.replace('Heading ', '')
            cfg = heading_config.get(style_name, {'font': '黑体', 'size': 14, 'bold': True})
            for run in para.runs:
                set_run_font(run, cfg['font'], cfg['size'], cfg['bold'])
            set_paragraph_format(para, first_line_indent=False)
        else:
            for run in para.runs:
                set_run_font(run, '宋体', 12)
            set_paragraph_format(para, first_line_indent=True)

    doc.save(docx_path)
    logger.info(f"DOCX 格式化完成: {docx_path}")


# ==================== 核心函数 ====================
def load_and_split_knowledge(kb_dir: str, chunk_size: int = 800, chunk_overlap: int = 150) -> List[Document]:
    raw_docs = []
    kb_path = Path(kb_dir)

    if not kb_path.exists():
        logger.error(f"知识库目录不存在: {kb_dir}")
        return []

    for file_path in kb_path.rglob("*"):
        if not file_path.is_file():
            continue
        ext = file_path.suffix.lower()
        try:
            if ext == ".pdf":
                with pdfplumber.open(file_path) as pdf:
                    for page_num, page in enumerate(pdf.pages, 1):
                        text = page.extract_text() or ""
                        for t_idx, table in enumerate(page.extract_tables()):
                            table_str = "\n".join([" | ".join(str(c) if c else "" for c in row) for row in table])
                            text += f"\n\n【表格 {t_idx+1}】\n{table_str}"
                        if text.strip():
                            raw_docs.append(Document(page_content=text, metadata={
                                "source": str(file_path), "page": page_num, "type": "pdf", "filename": file_path.name
                            }))

            elif ext == ".docx":
                for doc in Docx2txtLoader(str(file_path)).load():
                    doc.metadata.update({"source": str(file_path), "type": "word", "filename": file_path.name})
                    if doc.page_content.strip():
                        raw_docs.append(doc)

            elif ext == ".doc":
                try:
                    for doc in Docx2txtLoader(str(file_path)).load():
                        doc.metadata.update({"source": str(file_path), "type": "word", "filename": file_path.name})
                        if doc.page_content.strip():
                            raw_docs.append(doc)
                except Exception:
                    pass
                try:
                    result = subprocess.run(
                        ["antiword", str(file_path)],
                        capture_output=True, text=True, timeout=30
                    )
                    if result.returncode == 0 and result.stdout.strip():
                        raw_docs.append(Document(
                            page_content=result.stdout.strip(),
                            metadata={"source": str(file_path), "type": "word", "filename": file_path.name}
                        ))
                        logger.info(f"通过 antiword 解析旧 .doc 文件: {file_path.name}")
                    else:
                        logger.warning(f"无法解析旧 .doc 文件: {file_path.name}")
                except (FileNotFoundError, subprocess.TimeoutExpired):
                    logger.warning(f"无法解析旧 .doc 文件（antiword 不可用）: {file_path.name}")

        except Exception as e:
            logger.warning(f"跳过文件 {file_path.name}: {e}")

    if not raw_docs:
        logger.error("知识库目录下没有成功加载任何有效文档！")
        return []

    splitter = RecursiveCharacterTextSplitter(
        chunk_size=chunk_size,
        chunk_overlap=chunk_overlap,
        separators=["\n\n", "\n", "。", "！", "？", "；", "，", "、", " ", ""],
        keep_separator=True
    )
    split_docs = splitter.split_documents(raw_docs)
    logger.info(f"知识库加载完成：原始文档 {len(raw_docs)} 个，切分后 chunks {len(split_docs)} 个")
    return split_docs


def get_vectorstore():
    global VECTORSTORE
    if VECTORSTORE is not None:
        return VECTORSTORE

    if os.path.exists(CHROMA_DIR) and os.path.exists(os.path.join(CHROMA_DIR, "chroma.sqlite3")) and os.path.getsize(os.path.join(CHROMA_DIR, "chroma.sqlite3")) > 0:
        logger.info("从磁盘加载已有向量数据库")
        VECTORSTORE = Chroma(persist_directory=CHROMA_DIR, embedding_function=EMBEDDINGS)
        return VECTORSTORE

    logger.info("首次构建知识库向量数据库（来源：/home/bid/quote）...")
    docs = load_and_split_knowledge(KB_DIR)
    if not docs:
        raise ValueError("知识库为空，无法构建向量数据库。请确认 /home/bid/quote 目录下有可读取的 .pdf 或 .docx 文件。")

    VECTORSTORE = Chroma.from_documents(docs, EMBEDDINGS, persist_directory=CHROMA_DIR)
    logger.info(f"向量数据库构建并持久化完成，共 {len(docs)} 个 chunks")
    return VECTORSTORE


def extract_criteria(tender_path: str) -> List[Dict]:
    if not os.path.exists(tender_path):
        raise FileNotFoundError(f"招标文件不存在: {tender_path}")

    ext = tender_path.lower().split('.')[-1]

    try:
        if ext == "pdf":
            with pdfplumber.open(tender_path) as pdf:
                full_text = "\n".join([p.extract_text() or "" for p in pdf.pages])

        elif ext == "docx":
            loader = Docx2txtLoader(tender_path)
            docs = loader.load()
            full_text = "\n".join([doc.page_content for doc in docs])

        else:
            raise ValueError(f"目前仅支持 PDF 和 DOCX 格式的招标文件，当前文件类型: {ext}")

    except Exception as e:
        raise ValueError(f"无法解析招标文件: {tender_path}，错误: {str(e)}")

    prompt = f"""请从以下招标文件文本中提取所有评分标准，输出严格 JSON 数组。
每个元素包含：id, category, item_name, max_score, description, requirements。
只输出 JSON，不要其他文字。

招标文件内容：
{full_text[:16000]}"""

    result = call_qwen(prompt)
    try:
        import json
        return json.loads(result)
    except:
        return [{"id": "1", "category": "技术", "item_name": "整体响应", "max_score": 100,
                 "description": "完整响应招标要求", "requirements": "详见招标文件"}]


def call_qwen(prompt: str) -> str:
    try:
        response = client.chat.completions.create(
            model=MODEL_NAME,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.2,
            max_tokens=4096,
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        logger.error(f"模型调用失败: {e}")
        return f"[模型调用失败] {str(e)}"


def generate_section_markdown(criterion: Dict, vectorstore, section_num: int) -> str:
    """生成单个评分项的 Markdown，带自动编号"""
    query = f"{criterion.get('item_name', '')} {criterion.get('description', '')}"
    docs = vectorstore.similarity_search(query, k=6)
    context = "\n\n".join([f"（来源：{d.metadata.get('filename', '未知')}）\n{d.page_content[:900]}" for d in docs])

    prompt = f"""你是资深标书专家。

**评分项编号**：2.{section_num}
**评分项名称**：{criterion.get('item_name')}
**满分**：{criterion.get('max_score')}
**评分标准**：{criterion.get('description')}
**具体要求**：{criterion.get('requirements')}

**公司知识库相关材料**：
{context}

请用正式商务中文撰写针对此项的响应（350-550字）。

格式要求（严格遵守）：
1. 你的输出必须以"## 2.{section_num} {criterion.get('item_name')}"开头（这是Markdown二级标题，序号已由系统指定，不要修改）
2. 标题之后直接写正文段落，不要再重复出现评分项名称（不要用粗体、不要用标题格式二次出现）
3. 正文必须是纯文本段落，包含：直接回应、优势亮点（含量化数据）、证据引用（标注来源文件名）、符合性承诺
4. 正文之后，紧接着一行"### 2.{section_num}.1 支撑材料"（三级标题）
5. 支撑材料下列出项目符号列表（- 开头），每条引用一个来源文件
6. 全文必须使用中文全角标点符号（，。！？：；""''），严禁英文半角标点"""

    content = call_qwen(prompt)
    content = fix_chinese_punctuation(content)

    # 确保有二级标题前缀
    expected_heading = f"## 2.{section_num}"
    if not content.strip().startswith(expected_heading):
        content = f"## 2.{section_num} {criterion.get('item_name')}\n\n{content}"

    return content + "\n\n"


def build_bid_markdown(tender_path: str, vectorstore) -> str:
    criteria = extract_criteria(tender_path)

    # v3.0：不生成顶部文件路径和生成时间，直接进入标书正文
    md = "# 一、符合性声明\n\n我公司完全响应本次招标要求，承诺所提供产品及服务完全满足招标文件规定的全部技术与商务条款。\n\n"
    md += "# 二、评分项详细响应\n\n"

    for idx, crit in enumerate(criteria, 1):
        md += generate_section_markdown(crit, vectorstore, idx)

    md += "# 三、结束语\n\n以上为我公司针对本次招标的完整响应，期待贵方审阅。\n"
    return md


def md_to_docx(md_content: str, output_path: str):
    md_file = output_path.replace(".docx", ".md")
    with open(md_file, "w", encoding="utf-8") as f:
        f.write(md_content)

    # pandoc 生成裸 DOCX
    subprocess.run(["pandoc", md_file, "-o", output_path,
                    "--from", "markdown+smart",
                    "--to", "docx"], check=True)
    os.remove(md_file)

    # python-docx 格式化
    try:
        format_docx(output_path)
    except Exception as e:
        logger.warning(f"DOCX 格式化失败（非致命）: {e}")

    logger.info(f"标书已生成: {output_path}")


# ==================== 后台任务 ====================
async def process_generate_bid(task_id: str, req: GenerateBidRequest):
    try:
        tasks[task_id].update({"status": "running", "progress": 10, "message": "正在加载知识库向量数据库..."})
        vectorstore = get_vectorstore()

        tasks[task_id].update({"progress": 35, "message": "正在解析招标文件并提取评分标准..."})
        md_content = build_bid_markdown(req.tender_path, vectorstore)

        tasks[task_id].update({"progress": 75, "message": "正在生成 Word 文档并排版..."})
        output_path = os.path.join(PRODUCT_DIR, f"{req.uuid}.docx")
        await asyncio.to_thread(md_to_docx, md_content, output_path)

        tasks[task_id].update({
            "status": "completed",
            "progress": 100,
            "message": "标书生成成功",
            "result": {"docx_path": output_path, "uuid": req.uuid},
            "updated_at": datetime.now().isoformat()
        })
        logger.info(f"标书生成完成: {output_path}")
    except Exception as e:
        logger.error(f"生成标书失败: {e}", exc_info=True)
        tasks[task_id].update({
            "status": "failed", "progress": 100,
            "message": f"生成失败: {str(e)}",
            "updated_at": datetime.now().isoformat()
        })


# ==================== HTTP 接口 ====================
@app.get("/health")
async def health():
    return {
        "status": "ok",
        "base_url": OPENAI_BASE_URL,
        "model": MODEL_NAME,
        "log_file": LOG_FILE,
        "time": datetime.now().isoformat()
    }


@app.post("/generate_bid")
async def generate_bid(req: GenerateBidRequest, background_tasks: BackgroundTasks):
    task_id = str(uuid.uuid4())
    now = datetime.now().isoformat()
    tasks[task_id] = {
        "task_id": task_id, "status": "pending", "progress": 0,
        "message": "标书生成任务已创建", "result": None,
        "created_at": now, "updated_at": now
    }
    background_tasks.add_task(process_generate_bid, task_id, req)
    logger.info(f"收到生成标书请求，task_id={task_id}, uuid={req.uuid}")
    return {"task_id": task_id, "message": "任务已提交，请通过 /bid_progress/{task_id} 查询进度"}


@app.get("/bid_progress/{task_id}", response_model=TaskStatus)
async def get_bid_progress(task_id: str):
    if task_id not in tasks:
        raise HTTPException(404, "任务不存在")
    return TaskStatus(**tasks[task_id])


# ==================== 主入口 ====================
if __name__ == "__main__":
    logger.info("Bid AI Server HTTP 启动中...")
    uvicorn.run("bid_api_server:app", host="0.0.0.0", port=8000, reload=False)