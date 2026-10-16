# -*- coding: utf-8 -*-
"""
memory_bot.py —— 第四讲《Agent 的记忆管理》课后任务

大模型接口本身是无状态的：每次请求互不相干，"记忆"全靠调用方自己维护。
本程序把课件里的五段记忆方案做进一个命令行机器人：

    1. 完整历史    —— 每轮把 user/assistant 全量追加进 messages 重发（默认 MODE=full）
    2. 滑动窗口    —— MODE=window，只带最近 N 轮（messages[-2*N:]），更早的丢弃
    3. 压缩摘要    —— MODE=compress，历史超过 LIMIT 条就先让模型总结成一段话
    4. 持久化存储  —— 每轮对话追加写入 memory/今天日期.txt，程序重启后读回，记忆还在
    5. 关键词选取  —— 新问题先经 jieba 分词取关键词，命中哪个往期日期文件，
                      才把那个文件的内容带进请求（没装 jieba 自动降级为二字滑窗匹配）

用法：
    python memory_bot.py            # 聊天，输入 quit 退出
    MEMORY_MODE=window python memory_bot.py     # 滑动窗口模式
    MEMORY_MODE=compress python memory_bot.py   # 压缩摘要模式
"""
import datetime
import glob
import os
import sys

# Windows 终端默认 GBK：打印生僻字/emoji 会 UnicodeEncodeError，
# 管道喂进来的 UTF-8 中文也会被错解成 GBK 并产生孤立代理字符（API 直接 400）。
# 输入输出都改成 UTF-8 容错。
try:
    sys.stdout.reconfigure(errors="replace")
    sys.stdin.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

import requests

from llm import LLM_BASE, LLM_MODEL, load_key

# 预热 jieba 词典并吞掉它的加载日志，保持演示输出干净
try:
    import contextlib
    import io

    import jieba
    with contextlib.redirect_stderr(io.StringIO()):
        list(jieba.cut("预热"))
except ImportError:
    jieba = None

# ---------- 记忆策略开关 ----------
MODE = os.environ.get("MEMORY_MODE", "full")  # full | window | compress
N = 3      # 滑动窗口：只携带最近 N 轮（一轮 = user + assistant 两条）
LIMIT = 6  # 压缩摘要：历史条数超过 LIMIT 就先压缩

MEMORY_DIR = "memory"
TODAY_PATH = os.path.join(
    MEMORY_DIR, datetime.date.today().isoformat() + ".txt")

SYSTEM_PROMPT = "你是一个乐于助人的中文智能助手，回答简洁友好。"


# ---------- 大模型调用（OpenAI 兼容接口，与课件同构） ----------

def ask(messages):
    """整段 messages 发出去，返回 (回复文本, usage)。
    usage 里能看到 prompt_tokens 计费与 DeepSeek 的缓存命中字段。"""
    r = requests.post(
        LLM_BASE + "/chat/completions",
        headers={"Authorization": "Bearer " + load_key()},
        json={"model": LLM_MODEL, "messages": messages},
        timeout=60,
    )
    r.raise_for_status()
    data = r.json()
    return data["choices"][0]["message"]["content"], data.get("usage", {})


# ---------- 4. 持久化：当天一个文件，逐行 "role: content" ----------

def load_today():
    """读回今天的记忆文件，还原成 messages 列表。程序重启靠它找回记忆。"""
    os.makedirs(MEMORY_DIR, exist_ok=True)
    history = []
    if os.path.exists(TODAY_PATH):
        with open(TODAY_PATH, encoding="utf-8") as f:
            for line in f:
                line = line.rstrip("\n")
                if ": " in line:
                    role, content = line.split(": ", 1)
                    history.append({"role": role, "content": content})
    return history


def append_today(role, content):
    """一轮结束立刻落盘：哪怕程序马上崩溃，这一轮也不丢。"""
    with open(TODAY_PATH, "a", encoding="utf-8") as f:
        f.write(role + ": " + content.replace("\n", " ") + "\n")


# ---------- 5. 关键词选取：往期文件只挑命中的带进请求 ----------

def extract_keys(text):
    """从新问题里提取关键词：jieba 分词后去掉单字和标点。
    没装 jieba 时降级为连续汉字的二字滑窗，效果差一些但能用。"""
    if jieba is not None:
        return [w for w in jieba.cut(text) if len(w) > 1]
    else:
        keys, run = [], ""
        for ch in text:
            if "一" <= ch <= "鿿":
                run += ch
                if len(run) >= 2:
                    keys.append(run[-2:])
            else:
                run = ""
        return keys


def pick_past_days(question):
    """遍历 memory/ 下今天以外的日期文件，任一关键词命中就整天带回。"""
    keys = extract_keys(question)
    picked = []
    for path in sorted(glob.glob(os.path.join(MEMORY_DIR, "*.txt"))):
        if os.path.abspath(path) == os.path.abspath(TODAY_PATH):
            continue  # 今天的记忆已经在 history 里，不重复带
        with open(path, encoding="utf-8") as f:
            text = f.read()
        if any(k in text for k in keys):
            picked.append(os.path.basename(path) + "\n" + text.strip())
    return picked


# ---------- 3. 压缩摘要：太长就先让模型总结 ----------

def compress(history):
    req = history + [{"role": "user", "content": "把以上对话总结成一段话"}]
    summary, _ = ask(req)
    print("[已压缩成摘要]")
    return [{"role": "system", "content": "之前的对话摘要: " + summary}]


# ---------- 主循环 ----------

def main():
    history = load_today()
    print("记忆机器人已启动（模式: %s，模型: %s）" % (MODE, LLM_MODEL))
    if history:
        print("[记忆] 从今天文件读回 %d 条历史" % len(history))
    print("输入 quit 退出。\n")

    while True:
        try:
            text = input("你: ").strip()
        except (EOFError, KeyboardInterrupt):
            break
        if not text:
            continue
        if text == "quit":
            break
        if not sys.stdin.isatty():
            print(text)  # 管道喂题时回显（input 的提示符已打印），演示记录才完整

        # 5. 往期记忆按需选取：关键词命中哪天就只带哪天
        picked = pick_past_days(text)
        if picked:
            print("[记忆] 关键词命中往期文件，已带入请求（共 %d 段）"
                  % len(picked))

        # 3. 压缩摘要模式：历史过长先压缩
        if MODE == "compress" and len(history) > LIMIT:
            history = compress(history)

        # 组装本次请求
        outgoing = [{"role": "system", "content": SYSTEM_PROMPT}]
        if picked:
            outgoing.append({"role": "system",
                             "content": "往期的历史记录：\n" + "\n\n".join(picked)})
        # 1. 完整历史 / 2. 滑动窗口（只带最近 N 轮）
        window = history if MODE != "window" else history[-2 * N:]
        outgoing += window
        outgoing.append({"role": "user", "content": text})

        try:
            reply, usage = ask(outgoing)
        except Exception as e:
            print("机器人: 大模型出错了：" + str(e))
            continue

        print("机器人:", reply)
        # 9/10 页的观察点：token 计费与缓存命中
        print("[tokens] 本次携带 %d 条 | prompt=%s completion=%s 缓存命中=%s\n" % (
            len(outgoing),
            usage.get("prompt_tokens", "?"),
            usage.get("completion_tokens", "?"),
            usage.get("prompt_cache_hit_tokens", "?")))

        # 1+4. 全量记入历史并落盘
        history.append({"role": "user", "content": text})
        history.append({"role": "assistant", "content": reply})
        append_today("user", text)
        append_today("assistant", reply)

    print("再见！今天的对话已保存在", TODAY_PATH)


if __name__ == "__main__":
    main()
