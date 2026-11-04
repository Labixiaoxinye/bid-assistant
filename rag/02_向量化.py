# -*- coding: utf-8 -*-
r"""
第 2 步：向量化（Embedding）

把每一块文字变成一串数字（1024 维向量）。
用 bge-m3（硅基流动免费额度可用），批量提交，每批 32 条。

⚠️ 建库和查询必须用同一个模型。换模型 = 换坐标系，旧向量全部作废。

用法：python 02_向量化.py
输入：可以_chunks.json
输出：可以_vectors.json  {"chunks": [...], "vectors": [[...], ...]}
"""
import json, urllib.request

# ============ 改这里 ============
EMBED_KEY = "你的硅基流动KEY填这里"
IN_FILE   = r"可以_chunks.json"
OUT_FILE  = r"可以_vectors.json"
# ===============================

URL   = "https://api.siliconflow.cn/v1/embeddings"
MODEL = "BAAI/bge-m3"
BATCH = 32          # 每批多少条


def embed_batch(texts):
    body = json.dumps({"model": MODEL, "input": texts}).encode("utf-8")
    req = urllib.request.Request(URL, body, method="POST")
    req.add_header("Content-Type", "application/json")
    req.add_header("Authorization", "Bearer " + EMBED_KEY)
    with urllib.request.urlopen(req, timeout=120) as r:
        return [x["embedding"] for x in json.loads(r.read().decode("utf-8"))["data"]]


def main():
    chunks = json.load(open(IN_FILE, encoding="utf-8"))
    print(f"待向量化：{len(chunks)} 块")

    vectors = []
    for i in range(0, len(chunks), BATCH):
        batch = chunks[i:i + BATCH]
        vectors += embed_batch(batch)
        print(f"  {i + len(batch)}/{len(chunks)}")

    json.dump({"chunks": chunks, "vectors": vectors},
              open(OUT_FILE, "w", encoding="utf-8"), ensure_ascii=False)

    print(f"完成：{len(vectors)} 个向量，每个 {len(vectors[0])} 维")


if __name__ == "__main__":
    main()
