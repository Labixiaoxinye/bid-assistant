# -*- coding: utf-8 -*-
r"""
第 1 步：切块（Chunking）

把 docx 切成适合检索的小块：400 字/块，段落级重叠 1 段（约 80 字）。

为什么和主程序的 3000 字/块不一样？
· 3000 字是"塞给模型做提取"，块越大调用次数越少
· 400 字是"给检索用的"，块越小语义越纯 —— 一块里混了 20 件事，
  它的向量就是这 20 件事的平均值，反而不像任何一个具体问题

用法：python 01_切块.py
输出：可以_chunks.json（一个字符串数组）
"""
import json
from docx import Document

# ============ 改这里 ============
FILE = r"D:\投标决策助手跑数据文件\可以.docx"
OUT  = r"可以_chunks.json"
# ===============================

CHUNK_SIZE = 400     # 每块目标字数
OVERLAP_PARAS = 1    # 重叠：保留上一块最后几个段落


def main():
    doc = Document(FILE)
    paras = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    print(f"段落数：{len(paras)}")

    chunks = []
    cur, cur_len = [], 0

    for p in paras:
        cur.append(p)
        cur_len += len(p)
        if cur_len >= CHUNK_SIZE:
            chunks.append("\n".join(cur))
            # 重叠：最后一段留作下一块的开头，防止句子被切断
            keep = cur[-OVERLAP_PARAS:] if len(cur) > OVERLAP_PARAS else []
            cur = list(keep)
            cur_len = sum(len(x) for x in cur)

    if cur:
        chunks.append("\n".join(cur))

    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(chunks, f, ensure_ascii=False, indent=1)

    total = sum(len(c) for c in chunks)
    print(f"总字数：{total}")
    print(f"切成：{len(chunks)} 块，平均 {total // len(chunks)} 字")


if __name__ == "__main__":
    main()
