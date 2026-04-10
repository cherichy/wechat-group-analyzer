# 微信群成员活跃度分析与清理工具

基于 [Wetrace](https://github.com/afumu/wetrace) 本地微信数据库 API + 聊天记录 CSV 导出，对微信群进行成员活跃度分析，生成带头像的综合 Excel 花名册，辅助群主进行不活跃成员清理。

---

## 前置条件

| 依赖 | 说明 |
|------|------|
| **[Wetrace](https://github.com/afumu/wetrace)** | 本地运行，API 地址 `http://127.0.0.1:5200/api/v1` |
| **uv** | Python 包管理器（`scoop install uv`），用于创建本地虚拟环境 |
| **聊天记录 CSV** | 从 Wetrace 导出的群聊天记录（UTF-8 BOM 编码） |
| **消息 JSON**（可选） | 从 Wetrace 导出的 `messages_*.json`，用于提取聊天中出现的头像 URL |

### 环境初始化

```powershell
uv init
uv venv
uv pip install openpyxl Pillow
```

---

## 完整流程概览

```
┌──────────────────────────────────────────────────────────────┐
│  Step 1: fetch_members.py                                    │
│  从 Wetrace API 拉取群成员列表 → chatroom_info_raw.json      │
└────────────────────┬─────────────────────────────────────────┘
                     ▼
┌──────────────────────────────────────────────────────────────┐
│  Step 2: export_members.py                                   │
│  逐个查询联系人详情 + 活跃昵称 → group_members_full.csv      │
└────────────────────┬─────────────────────────────────────────┘
                     ▼
┌──────────────────────────────────────────────────────────────┐
│  Step 3: analyze.py                                          │
│  分析聊天记录 CSV → member_stats.csv + 控制台报告            │
└────────────────────┬─────────────────────────────────────────┘
                     ▼
┌──────────────────────────────────────────────────────────────┐
│  Step 4: build_final_excel.py                                │
│  汇总所有数据源 + 清理建议 → group_members_comprehensive.csv │
└────────────────────┬─────────────────────────────────────────┘
                     ▼
┌──────────────────────────────────────────────────────────────┐
│  Step 5: build_excel_with_avatars.py                         │
│  下载头像 + 嵌入图片 → group_members_comprehensive.xlsx      │
└──────────────────────────────────────────────────────────────┘
```

---

## 各脚本详解

### Step 1: `fetch_members.py` — 拉取群成员列表

**作用**：调用 Wetrace API `GET /chatrooms/{chatroom_id}` 获取完整群成员列表。

**输入**：无（直接调 API）
**输出**：`chatroom_info_raw.json`

**适配其他群时需修改**：
```python
# 把 chatroom ID 换成目标群的 ID
url = 'http://127.0.0.1:5200/api/v1/chatrooms/57327409534@chatroom'
#                                              ^^^^^^^^^^^^^^^^^^^^^^^^
#                                              改成你的群 chatroom ID
```

> **如何找到 chatroom ID？**
> 调用 `GET /sessions?keyword=群名关键词` 搜索，返回结果中的 `UserName` 字段即是。

---

### Step 2: `export_members.py` — 导出成员详细信息

**作用**：
1. 调用 `GET /analysis/member_activity/{chatroom_id}` 获取所有发言者的活跃昵称
2. 逐个调用 `GET /contacts/{user_id}` 获取每个成员的微信昵称、微信号、备注

**输入**：`chatroom_info_raw.json`（Step 1 输出）
**输出**：`group_members_full.csv`

**适配其他群时需修改**：
```python
url = f'{BASE}/analysis/member_activity/57327409534@chatroom'
#                                       ^^^^^^^^^^^^^^^^^^^^^^^^
```

> 注意：`/contacts` 批量接口可能返回空结果，但逐个 `/contacts/:id` 始终可用。

---

### Step 3: `analyze.py` — 分析聊天记录

**作用**：解析 Wetrace 导出的聊天记录 CSV，统计每人发言数、首/末次发言时间、活跃得分，输出分布图和 Top 30。

**输入**：聊天记录 CSV 文件
**输出**：`member_stats.csv` + 控制台报告

**适配其他群时需修改**：
```python
filename = r'chat_export_CuTeDSL & cuTile 交流群_57327409534@chatroom.csv'
#            ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
#            换成你导出的 CSV 文件名
```

**活跃得分公式**：
```
score = log2(发言数 + 1) × 10 + max(0, (365 - 距今天数)) / 365 × 50
```
- 发言越多、越近期活跃，得分越高
- 满分约 200（1000+条消息且最近刚发言）

---

### Step 4: `build_final_excel.py` — 汇总综合 CSV

**作用**：合并所有数据源（成员列表、联系人详情、聊天统计、消息头像），交叉比对生成清理建议，输出一个包含全部信息的综合 CSV。

**输入**：
- `chatroom_info_raw.json`（成员列表）
- `group_members_full.csv`（联系人详情）
- `member_stats.csv`（聊天统计）
- `messages_*.json`（聊天消息中的头像 URL，可选）

**输出**：`group_members_comprehensive.csv`

**清理逻辑**（内置）：
1. **从未发言** → 建议移除
2. **有发言记录** → 保留
3. **群主** → 始终保留

**输出字段**：

| 字段 | 说明 |
|------|------|
| 序号 | 排序后的编号 |
| 微信ID | `wxid_xxx` 或自定义微信号 |
| 可辨识昵称 | 优先取：备注 > 活跃昵称 > 微信昵称 > 群昵称 > ID |
| 微信昵称 | 来自 `/contacts/:id` API |
| 群昵称 | 来自 `/chatrooms/:id` API 的 displayName |
| 活跃昵称 | 来自 `/analysis/member_activity` |
| 微信号 | alias 字段 |
| 备注 | remark 字段 |
| 头像URL(小/大) | 来自消息 JSON 或 contacts API |
| 发言总数 | 聊天记录中的消息条数 |
| 首次/最后发言 | 日期 |
| 距今天数 | 最后发言距数据截止日的天数 |
| 活跃得分 | 综合评分 |
| 活跃状态 | "活跃" 或 "从未发言" |
| 清理建议 | "保留" / "保留(群主)" / "建议移除" |
| 是否群主 | "是" 或空 |

---

### Step 5: `build_excel_with_avatars.py` — 生成带头像的 Excel（最终产物）

**作用**：
1. 读取 `group_members_comprehensive.csv`
2. 对缺头像的成员，调 `/contacts/:id` API 补全头像 URL
3. 下载所有头像到 `avatars_cache/` 目录（有缓存，不重复下载）
4. 缩放头像到 40x40px，嵌入 Excel 单元格
5. 粉色背景 = 建议移除，黄色背景 = 从未发言
6. 首行冻结 + 自动筛选

**输入**：`group_members_comprehensive.csv`（Step 4）
**输出**：`group_members_comprehensive.xlsx`

**适配其他群时**：不需要改，它读的是通用格式的 CSV。

---

## 快速复用到其他群（Checklist）

### 1. 获取群 chatroom ID

```powershell
# 在浏览器或 curl 中访问
http://127.0.0.1:5200/api/v1/sessions?keyword=群名关键词
```

记下返回的 `UserName` 字段，格式为 `数字@chatroom`。

### 2. 从 Wetrace 导出聊天记录 CSV

使用 Wetrace 的导出功能，导出目标群的 CSV 聊天记录。文件名格式通常为：
```
chat_export_群名_chatroom_id.csv
```

### 3. 修改脚本中的 chatroom ID 和文件名

需要修改的文件和位置：

| 文件 | 行号 | 修改内容 |
|------|------|----------|
| `fetch_members.py` | 第 5 行 | URL 中的 chatroom ID |
| `export_members.py` | 第 22 行 | URL 中的 chatroom ID |
| `analyze.py` | 第 9 行 | CSV 文件名 |
| `build_final_excel.py` | 第 66 行 | messages JSON 文件名（如有） |

### 4. 按顺序执行

```powershell
.venv\Scripts\python.exe fetch_members.py
.venv\Scripts\python.exe export_members.py
.venv\Scripts\python.exe analyze.py
.venv\Scripts\python.exe build_final_excel.py
.venv\Scripts\python.exe build_excel_with_avatars.py
```

> **PowerShell 中文乱码问题**：脚本的控制台输出在 PowerShell 中可能乱码，可以管道重定向：
> ```powershell
> .venv\Scripts\python.exe xxx.py 2>&1 | Out-File -Encoding utf8 output.txt
> ```

### 5. 调整清理参数

在 `build_final_excel.py` 中，清理建议逻辑默认为"从未发言 → 建议移除"。如需更精细的控制（如按活跃得分阈值移除低活跃发言者），可修改该脚本中的清理判断部分。

---

## 目录结构

```
.
├── README.md                              # 本文档
├── pyproject.toml                         # uv 项目配置（需自行 uv init）
├── .venv/                                 # Python 虚拟环境（需自行创建）
│
├── 输入数据（需自行准备）
│   ├── chat_export_*@chatroom.csv         # Wetrace 导出的聊天记录
│   ├── messages_*@chatroom.json           # Wetrace 导出的消息 JSON（含头像，可选）
│   └── chatroom_info_raw.json             # Step 1 生成
│
├── 核心脚本（按执行顺序）
│   ├── fetch_members.py                   # Step 1: 拉取群成员列表
│   ├── export_members.py                  # Step 2: 查询联系人详情
│   ├── analyze.py                         # Step 3: 分析聊天活跃度
│   ├── build_final_excel.py               # Step 4: 汇总综合 CSV + 清理建议
│   └── build_excel_with_avatars.py        # Step 5: 生成带头像 Excel
│
├── 输出文件（脚本自动生成）
│   ├── group_members_comprehensive.xlsx   # ★ 最终产物：带头像的综合 Excel
│   ├── group_members_comprehensive.csv    # 综合 CSV（Excel 的数据源）
│   ├── group_members_full.csv             # 全员联系人信息
│   └── member_stats.csv                   # 聊天活跃度统计
│
└── avatars_cache/                         # 头像缓存目录（自动创建）
```

---

## Wetrace API 备忘

| 接口 | 用途 | 注意事项 |
|------|------|----------|
| `GET /sessions?keyword=` | 搜索会话，找 chatroom ID | — |
| `GET /chatrooms/:id` | 获取群成员列表 + 群主 | 返回 `users[].userName/displayName` |
| `GET /contacts/:id` | 查询单个联系人详情 | 字段小驼峰：`nickName`, `smallHeadImgUrl` |
| `GET /contacts` | 批量查询联系人 | 可能返回空，建议用逐个查询 |
| `GET /analysis/member_activity/:id` | 群发言者活跃数据 | 返回 `name`, `platformId`, `messageCount` |

---

## 已知问题与 Workaround

1. **PowerShell 中文乱码**：Python 脚本中需要 `sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')`，输出仍可能乱码，用 `Out-File -Encoding utf8` 重定向后读取。

2. **CSV 编码**：Wetrace 导出的 CSV 是 UTF-8 with BOM，偶有非法字节，读取时需 `errors='replace'`。

3. **CSV 脏数据**：部分行的 `sender_id` 字段包含空格、换行或超长内容（消息内容溢出），需过滤：
   ```python
   if ' ' in sender_id or len(sender_id) > 50 or '\n' in sender_id:
       continue
   ```

4. **`/contacts` 批量接口**：可能返回 0 结果，原因不明。逐个 `/contacts/:id` 始终可靠。

5. **头像来源优先级**：消息 JSON 只包含发过言的人的头像；从未发言者需要通过 `/contacts/:id` 的 `smallHeadImgUrl` 字段补全（`build_excel_with_avatars.py` 已自动处理）。
