# bidbydoc
根据上传的招标文件，生成标书文件

```markdown
# 标书AI生成服务 — 技术方案文档

> 版本：v3.0.0  
> 最后更新：2026-06-04  
> 适用环境：Linux (Ubuntu 22.04+ / CentOS 7+) + Python 3.10+

---

## 目录

1. [项目概述](#1-项目概述)
2. [系统架构](#2-系统架构)
3. [服务器与网络拓扑](#3-服务器与网络拓扑)
4. [文件目录结构](#4-文件目录结构)
5. [依赖清单](#5-依赖清单)
6. [安装部署步骤](#6-安装部署步骤)
7. [配置参数说明](#7-配置参数说明)
8. [API 接口文档](#8-api-接口文档)
9. [知识库管理](#9-知识库管理)
10. [运维管理](#10-运维管理)
11. [常见问题与故障排查](#11-常见问题与故障排查)
12. [版本演进与变更记录](#12-版本演进与变更记录)

---

## 1. 项目概述

### 1.1 项目背景

为满足企业投标需求，利用大语言模型（LLM）与本地知识库（公司资质、历史业绩、技术方案等），实现**招标文件的自动解析**与**标书文档的自动生成**。减少人工编制标书的工作量，保证响应内容的一致性和专业性。

### 1.2 核心能力

| 能力 | 说明 |
|------|------|
| 招标文件解析 | 支持 PDF / DOCX 格式，提取评分标准与技术要求 |
| 知识库检索 | 基于向量语义搜索，从公司历史文档中匹配最相关的响应材料 |
| AI 标书生成 | 调用 Qwen3-30B-A3B 大模型，自动撰写商务中文标书正文 |
| 文档排版输出 | 输出 .docx 格式，含黑体标题、宋体正文、首行缩进、1.5 倍行距 |
| 中文规范处理 | 全角标点符号、层级标题编号（一/1.1/1.1.1） |

### 1.3 技术栈

```
┌─────────────────────────────────────────────────────┐
│                    标书生成服务                       │
│  FastAPI + Uvicorn  (HTTP API, async 后台任务)       │
│  LangChain + Chroma  (向量知识库)                     │
│  HuggingFace bge-m3  (中文 Embedding 模型)            │
│  pdfplumber / docx2txt  (文档解析)                    │
│  python-docx  (DOCX 格式化排版)                       │
│  pandoc  (Markdown → DOCX 转换)                       │
└─────────────────────────────────────────────────────┘
```

---

## 2. 系统架构

### 2.1 数据流

```
用户上传招标文件 (.docx/.pdf)
        │
        ▼
┌──────────────────┐
│  FastAPI 服务     │  POST /generate_bid
│  (端口 8000)     │
└──────┬───────────┘
       │
       ├──(1)──► extract_criteria()     解析招标文件 → 提取评分标准 JSON
       │              │
       │              └── 调用 Qwen3-30B-A3B (vLLM API)
       │
       ├──(2)──► get_vectorstore()      加载/构建 Chroma 向量数据库
       │              │
       │              └── BAAI/bge-m3  (离线 Embedding)
       │              └── 知识库来源: /home/bid/quote/
       │
       ├──(3)──► generate_section_markdown()  对每条评分标准：
       │              │                          ① 语义检索知识库 (k=6)
       │              │                          ② 调用 LLM 生成商务响应
       │              │                          ③ 中文标点规范化
       │              └── build_bid_markdown()  拼接完整 Markdown
       │
       └──(4)──► md_to_docx()
                      │
                      ├── pandoc: Markdown → 裸 DOCX
                      └── format_docx(): 字体/缩进/行距排版
                              │
                              ▼
                      /home/bid/product/{uuid}.docx
```

### 2.2 模块职责

| 模块 | 文件位置 | 职责 |
|------|----------|------|
| HTTP 服务层 | `bid_api_server.py` (顶层) | FastAPI 路由、后台任务调度 |
| Pydantic 模型 | `bid_api_server.py` (类定义) | 请求/响应校验 |
| 知识库加载 | `load_and_split_knowledge()` | 遍历 /home/bid/quote/，解析 PDF/DOCX/DOC |
| 向量数据库 | `get_vectorstore()` | Chroma 持久化管理 |
| 招标解析 | `extract_criteria()` | 用 LLM 提取评分标准 JSON |
| LLM 调用 | `call_qwen()` | OpenAI 兼容 API 调用 vLLM |
| 标书生成 | `generate_section_markdown()` + `build_bid_markdown()` | 逐项生成 Markdown |
| DOCX 排版 | `format_docx()` | python-docx 后处理 |
| 启动入口 | `__main__` | 预加载 Embedding → 启动 uvicorn |

---

## 3. 服务器与网络拓扑

### 3.1 服务器清单

| 角色 | IP 地址 | 说明 |
|------|---------|------|
| **应用服务器** | `10.10.130.102` | 运行 FastAPI 标书生成服务 |
| **vLLM 推理服务器** | `10.10.130.107` | 部署 Qwen3-30B-A3B 模型，提供 OpenAI 兼容 API |
| **HuggingFace 缓存** | 本地 (`/root/.cache/huggingface/`) | BAAI/bge-m3 离线加载 |

### 3.2 网络连通性要求

```
应用服务器 (10.10.130.102)  ──HTTP──►  vLLM 服务器 (10.10.130.107:7890)
应用服务器 (10.10.130.102)  ──无需──►  huggingface.co (离线模式)
```

- 应用服务器 **必须**能访问 vLLM 服务器的 7890 端口
- 应用服务器 **不需要**能访问外网（HuggingFace Embedding 模型已预缓存）

### 3.3 端口分配

| 服务 | 端口 | 协议 |
|------|------|------|
| FastAPI 标书服务 | 8000 | HTTP |
| vLLM 推理 API | 7890 | HTTP (OpenAI 兼容) |

---

## 4. 文件目录结构

```
/home/bid/
├── bid_api_server.py              # 主服务程序 (v3.0)
├── start_bid_server.sh            # systemd 启动脚本
├── requirements.txt               # Python 依赖清单 (可选)
│
├── quote/                         # 知识库目录 (原始文档)
│   ├── *.pdf                      #   公司资质、历史标书等 PDF 文件
│   ├── *.docx                     #   技术方案等 DOCX 文件
│   ├── *.doc                      #   旧格式 .doc 文件 (通过 antiword 解析)
│   └── *.xlsx                     #   电子表格 (当前跳过，不解析)
│
├── chroma_quote/                  # Chroma 向量数据库 (持久化)
│   └── chroma.sqlite3             #   向量索引文件
│
├── product/                       # 输出目录 (生成的标书文件)
│   └── {uuid}.docx                #   生成的 Word 标书文档
│
├── logs/                          # 日志目录
│   └── bid_api_server.log         #   滚动日志 (10MB × 5)
│
├── test/                          # 测试招标文件
│   └── *.docx                     #   待响应的招标文件
│
└── venv/                          # Python 虚拟环境 (可选)
    └── ...
```

### 关键路径常量

| 常量 | 路径 | 用途 |
|------|------|------|
| `KB_DIR` | `/home/bid/quote` | 知识库源文件目录 |
| `CHROMA_DIR` | `/home/bid/chroma_quote` | Chroma 向量库持久化目录 |
| `PRODUCT_DIR` | `/home/bid/product` | 标书生成输出目录 |
| `LOG_DIR` | `/home/bid/logs` | 日志存放目录 |
| `LOG_FILE` | `/home/bid/logs/bid_api_server.log` | 主日志文件 |

---

## 5. 依赖清单

### 5.1 系统依赖

| 软件 | 最低版本 | 用途 | 安装命令 |
|------|----------|------|----------|
| Python | 3.10+ | 运行环境 | `apt install python3` |
| pandoc | 2.10+ | Markdown → DOCX 转换 | `apt install pandoc` |
| antiword | 0.37 | 解析旧 .doc 文件 (可选) | `apt install antiword` |

### 5.2 Python 依赖

```
# requirements.txt
fastapi>=0.110.0              # Web 框架
uvicorn[standard]>=0.29.0     # ASGI 服务器
pydantic>=2.0.0               # 数据校验
openai>=1.30.0                # vLLM OpenAI 兼容客户端
pdfplumber>=0.10.0            # PDF 解析
langchain>=0.2.0              # LLM 框架
langchain-community>=0.2.0    # 社区集成 (文档加载器、Embedding、Chroma)
langchain-text-splitters>=0.2.0  # 文本分割
langchain-core>=0.2.0         # 核心类型
chromadb>=0.5.0               # 向量数据库
sentence-transformers>=2.7.0 # Embedding 模型
docx2txt>=0.8                 # DOCX 文本提取
python-docx>=1.1.0            # DOCX 格式化排版
```

### 5.3 预训练模型

| 模型 | 用途 | 大小 | 缓存路径 |
|------|------|------|----------|
| `BAAI/bge-m3` | 中文文本向量化 (Embedding) | ~2.4 GB | `/root/.cache/huggingface/hub/models--BAAI--bge-m3/` |

**首次部署必须预下载模型**（服务器无外网时尤为重要）。

---

## 6. 安装部署步骤

### 6.1 环境准备

```bash
# 1. 安装系统依赖
apt-get update
apt-get install -y pandoc antiword python3 python3-pip python3-venv

# 2. 创建目录结构
mkdir -p /home/bid/{quote,chroma_quote,product,logs,test}

# 3. 创建 Python 虚拟环境 (推荐)
python3 -m venv /home/bid/venv
source /home/bid/venv/bin/activate
```

### 6.2 安装 Python 依赖

```bash
# 激活虚拟环境后
pip install --upgrade pip
pip install fastapi uvicorn[standard] pydantic openai pdfplumber
pip install langchain langchain-community langchain-text-splitters langchain-core
pip install chromadb sentence-transformers docx2txt python-docx
```

如果使用 conda 环境，跳过 venv，直接用 conda 的 pip：

```bash
/root/miniconda3/bin/pip install fastapi uvicorn[standard] pydantic openai pdfplumber ...
```

### 6.3 预下载 Embedding 模型

```bash
# 如有外网访问权限
huggingface-cli download BAAI/bge-m3

# 或设置国内镜像
HF_ENDPOINT=https://hf-mirror.com huggingface-cli download BAAI/bge-m3

# 验证缓存
ls /root/.cache/huggingface/hub/models--BAAI--bge-m3/snapshots/
```

### 6.4 放置文件

```bash
# 1. 上传 bid_api_server.py
scp bid_api_server.py root@10.10.130.102:/home/bid/

# 2. 上传知识库文件
scp *.pdf *.docx root@10.10.130.102:/home/bid/quote/

# 3. 上传测试招标文件
scp 招标文件.docx root@10.10.130.102:/home/bid/test/
```

### 6.5 创建启动脚本

```bash
cat > /home/bid/start_bid_server.sh << 'EOF'
#!/bin/bash
# 请根据你的 Python 环境修改以下路径
CONDA_PYTHON="/root/miniconda3/bin/python"
VENV_PYTHON="/home/bid/venv/bin/python"

if [ -f "$CONDA_PYTHON" ]; then
    PYTHON="$CONDA_PYTHON"
elif [ -f "$VENV_PYTHON" ]; then
    PYTHON="$VENV_PYTHON"
else
    echo "未找到可用的 Python 环境"
    exit 1
fi

cd /home/bid
exec "$PYTHON" bid_api_server.py
EOF

chmod +x /home/bid/start_bid_server.sh
```

### 6.6 配置 systemd 服务

```bash
cat > /etc/systemd/system/bid_api_server.service << 'EOF'
[Unit]
Description=Bid Document AI Server
After=network.target

[Service]
Type=simple
User=root
ExecStart=/home/bid/start_bid_server.sh
Restart=always
RestartSec=5
StandardOutput=journal
StandardError=journal

[Install]
WantedBy=multi-user.target
EOF

systemctl daemon-reload
systemctl enable bid_api_server
systemctl start bid_api_server
```

### 6.7 验证部署

```bash
# 1. 检查服务状态
systemctl status bid_api_server

# 2. 健康检查
curl http://localhost:8000/health

# 3. 预期输出
# {"status":"ok","base_url":"http://10.10.130.107:7890/v1","model":"Qwen/Qwen3-30B-A3B",...}

# 4. 端到端测试
curl -X POST "http://localhost:8000/generate_bid" \
  -H "Content-Type: application/json" \
  -d '{
    "tender_path":"/home/bid/test/招标文件.docx",
    "uuid":"test-001"
  }'

# 5. 查询进度
curl "http://localhost:8000/bid_progress/{task_id}"
```

---

## 7. 配置参数说明

### 7.1 环境变量

服务通过 `os.getenv()` 读取环境变量，支持默认值：

| 环境变量 | 默认值 | 说明 |
|----------|--------|------|
| `OPENAI_BASE_URL` | `http://10.10.130.107:7890/v1` | vLLM 推理服务地址 |
| `MODEL_NAME` | `Qwen/Qwen3-30B-A3B` | 使用的模型名称 |
| `API_KEY` | `qanwen2025` | vLLM API 密钥 |

### 7.2 硬编码常量

| 代码常量 | 值 | 说明 |
|----------|-----|------|
| `chunk_size` | 800 | 知识库文本切分块大小 |
| `chunk_overlap` | 150 | 相邻文本块重叠字符数 |
| `k` | 6 | 向量搜索返回的相关文档数 |
| `temperature` | 0.2 | LLM 生成温度（低=保守/一致） |
| `max_tokens` | 4096 | LLM 单次生成最大 Token 数 |
| `full_text[:16000]` | 16K 字符 | 招标文件截断长度 |

### 7.3 日志配置

| 参数 | 值 |
|------|-----|
| 日志级别 | `INFO` |
| 单文件最大 | 10 MB |
| 备份保留 | 5 个历史文件 |
| 编码 | UTF-8 |

---

## 8. API 接口文档

### 8.1 健康检查

```http
GET /health
```

**响应示例：**

```json
{
  "status": "ok",
  "base_url": "http://10.10.130.107:7890/v1",
  "model": "Qwen/Qwen3-30B-A3B",
  "log_file": "/home/bid/logs/bid_api_server.log",
  "time": "2026-06-04T12:00:00"
}
```

### 8.2 生成标书

```http
POST /generate_bid
Content-Type: application/json
```

**请求体：**

| 字段 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `tender_path` | string | 是 | 招标文件的绝对路径（支持 .pdf / .docx） |
| `uuid` | string | 是 | 生成的标书文件名（不含扩展名） |

**请求示例：**

```json
{
  "tender_path": "/home/bid/test/CHDTDZ14717-SB-72302-小纪汗煤矿抗冲击地压智能综采放顶煤液压支架购置项目.docx",
  "uuid": "20240604-012"
}
```

**响应示例：**

```json
{
  "task_id": "e234e55c-7f7e-4d99-902c-c89ba8bf027e",
  "message": "任务已提交，请通过 /bid_progress/{task_id} 查询进度"
}
```

**注意：** 生成是异步的，立即返回 `task_id`，需要轮询进度接口。生成的标书保存在 `/home/bid/product/{uuid}.docx`。

### 8.3 查询任务进度

```http
GET /bid_progress/{task_id}
```

**状态码说明：**

| status | progress | 含义 |
|--------|----------|------|
| `pending` | 0 | 任务已提交，等待执行 |
| `running` | 10 | 正在加载知识库向量数据库 |
| `running` | 35 | 正在解析招标文件并提取评分标准 |
| `running` | 75 | 正在调用 pandoc 生成 Word 文档并格式化 |
| `completed` | 100 | 标书生成成功 |
| `failed` | 100 | 生成失败，见 `message` 字段 |

**成功响应：**

```json
{
  "task_id": "e234e55c-7f7e-4d99-902c-c89ba8bf027e",
  "status": "completed",
  "progress": 100,
  "message": "标书生成成功",
  "result": {
    "docx_path": "/home/bid/product/20240604-012.docx",
    "uuid": "20240604-012"
  },
  "created_at": "2026-06-04T11:44:33.706000",
  "updated_at": "2026-06-04T11:47:00.123456"
}
```

**失败响应：**

```json
{
  "task_id": "...",
  "status": "failed",
  "progress": 100,
  "message": "生成失败: 知识库为空，无法构建向量数据库。请确认 /home/bid/quote 目录下有可读取的文件。",
  "result": null,
  "created_at": "...",
  "updated_at": "..."
}
```

---

## 9. 知识库管理

### 9.1 支持的文件格式

| 格式 | 解析方式 | 状态 |
|------|----------|------|
| `.pdf` | `pdfplumber` 提取文本 + 表格 | ✅ 完整支持 |
| `.docx` | `Docx2txtLoader` (langchain) | ✅ 完整支持 |
| `.doc` | `Docx2txtLoader` 尝试 → `antiword` 回退 | ⚠️ 需安装 antiword |
| `.xlsx` | 不支持 | ❌ 跳过 |

### 9.2 知识库更新流程

当 `/home/bid/quote/` 中新增或修改文件后，**必须重建向量数据库**：

```bash
# 方法1：删除旧向量库，重启服务（推荐）
systemctl stop bid_api_server
rm -rf /home/bid/chroma_quote
systemctl start bid_api_server
# 服务启动时会自动重建

# 方法2：调用 API 触发首次请求时自动重建
# 但耗时较长，不推荐大批量更新时使用
```

### 9.3 知识库文档量参考

| 文档量 | 预估 chunks | Chroma 大小 | 构建耗时 |
|--------|-------------|-------------|----------|
| 3 个 DOCX | ~30 | ~200 KB | < 1 分钟 |
| 7 个 XLSX + 3 个 DOCX + 4 个 DOC | ~150 | ~2 MB | 2-5 分钟 |
| 50+ 文档 | ~2000+ | ~50 MB+ | 10-30 分钟 |

---

## 10. 运维管理

### 10.1 常用命令速查

```bash
# 查看服务状态
systemctl status bid_api_server

# 重启服务
systemctl restart bid_api_server

# 停止服务
systemctl stop bid_api_server

# 查看实时日志
journalctl -u bid_api_server -f

# 查看应用日志
tail -f /home/bid/logs/bid_api_server.log

# 检查端口占用
ss -tlnp | grep 8000

# 强制释放端口
fuser -k 8000/tcp
```

### 10.2 日志分析

日志格式：`时间 - 级别 - 消息`

关键日志节点：

```
正常启动：
  正在加载 BAAI/bge-m3 embedding 模型（离线模式）...
  BAAI/bge-m3 加载完成
  从磁盘加载已有向量数据库
  Bid AI Server HTTP 启动中...

请求到达：
  收到生成标书请求，task_id=xxx, uuid=xxx

生成完成：
  标书生成完成: /home/bid/product/xxx.docx

异常情况：
  生成标书失败: Cannot send a request, as the client has been closed.
  → 原因：Embedding 模型加载失败或网络问题
```

### 10.3 性能参考

| 环节 | 耗时 (参考) |
|------|-------------|
| Embedding 模型加载（首次） | 30-60 秒 |
| Chroma 向量库加载（已有） | 1-3 秒 |
| 招标文件解析 (50页 DOCX) | 3-10 秒 |
| LLM 评分标准提取 | 10-30 秒 |
| 每条评分项生成（LLM） | 15-45 秒 |
| DOCX 排版格式化 | 1-5 秒 |
| **总耗时（5 个评分项）** | **3-6 分钟** |

---

## 11. 常见问题与故障排查

### 11.1 `Address already in use` (端口 8000 被占用)

```bash
fuser -k 8000/tcp
systemctl restart bid_api_server
```

### 11.2 `Cannot send a request, as the client has been closed`

**根因：** Embedding 模型在 async 上下文中尝试网络请求，httpx 客户端已关闭。

**确认步骤：**

```bash
# 1. 检查模型是否已缓存
ls /root/.cache/huggingface/hub/models--BAAI--bge-m3/snapshots/

# 2. 检查代码是否包含 local_files_only=True
grep local_files_only /home/bid/bid_api_server.py

# 3. 确认 EMBEDDINGS 在模块顶层初始化（不在 if __name__ 或 async 事件中）
grep -n "EMBEDDINGS =" /home/bid/bid_api_server.py
```

**解决方案：** 使用 v3.0 版本代码（已在模块顶层初始化 + `local_files_only=True`）。

### 11.3 `SyntaxError: unmatched ')'`

**根因：** 第 55 行 `API_KEY` 变量在传输时被截断。

**解决方案：**

```bash
# 检查并修复
sed -n '55p' /home/bid/bid_api_server.py
# 应显示：API_KEY=*** "qanwen2025")
# 如不完整，手动编辑第 55 行
vim /home/bid/bid_api_server.py
```

### 11.4 服务反复崩溃重启

**现象：** `systemctl status` 显示 `activating (auto-restart)`，日志只有"启动中"。

**排查：**

```bash
# 查看 journal 错误栈
journalctl -u bid_api_server --no-pager -n 50

# 前台运行观察
cd /home/bid
python3 bid_api_server.py
```

### 11.5 `No /Root object! - Is this really a PDF?`

**根因：** 知识库中的文件被误判为 PDF 格式处理。

**排查：**

```bash
# 检查实际文件类型
file /home/bid/quote/*
```

### 11.6 HuggingFace 模型下载卡死

**现象：** 启动时卡在"正在加载 BAAI/bge-m3"，最后超时被 systemd 杀掉。

**解决方案：**

```bash
# 确认缓存存在
ls /root/.cache/huggingface/hub/models--BAAI--bge-m3/snapshots/

# 如果不存在，离线预下载（在有外网的机器上）
HF_ENDPOINT=https://hf-mirror.com huggingface-cli download BAAI/bge-m3

# 确保代码中有 local_files_only=True
```

---

## 12. 版本演进与变更记录

| 版本 | 日期 | 关键变更 |
|------|------|----------|
| v2.5 | 2026-06-03 | 初始版本，支持 PDF/DOCX 解析，Chroma 向量库，基础 Markdown → DOCX |
| v2.6 | 2026-06-04 | 新增 `startup_event` 预加载，`.doc` 容错（antiword 回退），修复重复构建 |
| v2.7 | 2026-06-04 | Embedding 加载移至 `if __name__` 主线程，避免 async httpx 冲突 |
| v2.8 | 2026-06-04 | Embedding 移至**模块顶层**初始化，加 `local_files_only=True` 离线加载 |
| v2.9 | 2026-06-04 | 新增 `format_docx()` 排版后处理：黑体标题、宋体正文、首行缩进、1.5 行距 |
| **v3.0** | **2026-06-04** | 标题层级（H1/H2/H3）带自动编号；`fix_chinese_punctuation()` 中文标点修复；删除顶部冗余信息；LLM prompt 优化防重复标题 |

---

```
