# Agent 的记忆管理 · 第四讲课后任务

中山大学《软工新生研讨实践》第四讲课后任务：**跑一个带记忆的对话程序，问它两个问题，第二个问题在第一个的答案基础上追问**（课件示例：先给名字让它写藏头诗，再问它解读诗意）。

本仓库实现了一个命令行记忆机器人 `memory_bot.py`，把课件讲的五段记忆方案全部做了进去，另附课件讲解与 Agent 功能说明两份 Word 文档（`docs/`）。

## 仓库结构

| 文件 | 说明 |
|---|---|
| `memory_bot.py` | 记忆机器人：完整历史 / 滑动窗口 / 压缩摘要三档 + 按天持久化 + jieba 关键词选取 |
| `llm.py` | 大模型调用封装（OpenAI 兼容接口，可换任意服务商） |
| `demo_run1.txt` | 演示记录一：藏头诗 → 追问诗意 → 回忆名字（同一会话内的记忆） |
| `demo_run2.txt` | 演示记录二：重启后仍记得名字与藏头诗（持久化）+ 关键词命中往期文件（选取） |
| `docs/第四讲-讲解.docx` | 课件逐节讲解 + 术语解释 |
| `docs/Agent功能说明.docx` | 本 Agent 的主要功能说明 |
| `gen_docs.py` | 生成上述两份 Word 的脚本 |

## 快速开始

```bash
pip install -r requirements.txt
```

### 1. 配置大模型密钥

任选其一：

- 环境变量：`LLM_KEY`（可选 `LLM_BASE`、`LLM_MODEL`，默认 OpenAI `gpt-4o-mini`）
- 当前目录新建 `llm_key.txt`，粘贴密钥（已在 `.gitignore`）

换国产模型示例（DeepSeek）：

```bash
export LLM_BASE=https://api.deepseek.com/v1
export LLM_MODEL=deepseek-chat
```

### 2. 运行记忆机器人

```bash
python memory_bot.py                    # 默认：完整历史模式
MEMORY_MODE=window python memory_bot.py # 滑动窗口：只带最近 3 轮
MEMORY_MODE=compress python memory_bot.py # 压缩摘要：超过 6 条先总结
```

输入 `quit` 退出。每轮对话追加写入 `memory/今天日期.txt`，重启后自动读回；
提问时会先用 jieba 分词提取关键词，命中哪个往期日期文件才把哪天带进请求。

## 五段记忆方案对照（课件总结表）

| 做法 | 解决了什么 | 留下的问题 | 本仓库对应 |
|---|---|---|---|
| 完整历史 | 模型能记住对话 | 终会撞到上下文窗口上限 | `MODE=full`（默认） |
| 滑动窗口 | 请求长度可控 | 丢失更早的内容 | `MODE=window`（`N=3`） |
| 压缩摘要 | 长对话压短 | 丢失具体细节 | `MODE=compress`（`LIMIT=6`） |
| 持久化与选取 | 重启不丢、按需取用 | 关键词匹配程度有限（只认字面） | `memory/*.txt` + `pick_past_days()` |
| 向量检索 / Agentic Search（延伸） | 按语义找、让模型自己搜 | 需要向量数据库 / 工具调用能力 | 课件延伸内容，未实现 |

## 安全须知

`llm_key.txt` 等密钥文件全部在 `.gitignore` 中，**永远不会被提交**。
密钥就是密码，不要发给任何人、不要截图外传。
