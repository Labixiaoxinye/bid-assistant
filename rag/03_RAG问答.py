# -*- coding: utf-8 -*-
r"""
RAG 第 3 步：完整问答（检索 + 生成）  —— 防闪退版
只改下面 EMBED_KEY 这一行。DeepSeek 的 Key 会自动从投标助手配置里读。
"""
import sys, os, json, math, traceback, urllib.request, urllib.error

sys.stdout.reconfigure(encoding="utf-8")

# ========================= 只改这里 =========================
EMBED_KEY = "你的硅基流动KEY填这里"
# ===========================================================

VEC_FILE = r"可以_vectors.json"

EMBED_URL   = "https://api.siliconflow.cn/v1/embeddings"
EMBED_MODEL = "BAAI/bge-m3"

DS_URL   = "https://api.deepseek.com/chat/completions"
DS_MODEL = "deepseek-chat"

FINAL_K   = 8
MIN_SCORE = 0.55
BONUS     = 0.05


def load_ds_key():
    p = os.path.join(os.environ.get("APPDATA", ""), "投标决策助手", "config.json")
    if os.path.exists(p):
        try:
            return json.load(open(p, encoding="utf-8")).get("api_key", "")
        except Exception:
            return ""
    return ""


DS_KEY = load_ds_key()


def http_post(url, payload, key):
    body = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(url, body, method="POST")
    req.add_header("Content-Type", "application/json")
    req.add_header("Authorization", "Bearer " + key)
    try:
        with urllib.request.urlopen(req, timeout=60) as r:
            return json.loads(r.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        msg = e.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"接口返回 HTTP {e.code}\n服务器原话：{msg[:400]}") from None


def embed_many(texts):
    resp = http_post(EMBED_URL, {"model": EMBED_MODEL, "input": texts}, EMBED_KEY)
    return [x["embedding"] for x in resp["data"]]


def cosine(a, b):
    dot = sum(x * y for x, y in zip(a, b))
    na = math.sqrt(sum(x * x for x in a))
    nb = math.sqrt(sum(y * y for y in b))
    return 0.0 if na == 0 or nb == 0 else dot / (na * nb)


EXPAND = {
    "废标": ["废标", "否决投标", "投标无效", "不予受理", "取消投标资格", "视为放弃"],
    "保证金": ["投标保证金", "保证金缴纳截止时间", "保证金退还", "保证金不予退还情形"],
    "时间": ["投标截止时间", "开标时间", "澄清答疑时间", "招标文件获取时间", "投标有效期"],
    "限价": ["最高投标限价", "投标报价上限", "报价不得超过"],
    "资质": ["企业资质要求", "安全生产许可证", "项目经理资格", "类似业绩要求"],
    "清单": ["工程量清单", "分部分项工程量清单", "措施项目清单"],
}
LIST_KWS = ["废标", "否决投标", "投标无效", "不予受理", "取消投标资格",
            "视为不参加投标", "视为放弃", "不予接收", "拒绝接收",
            "将被拒绝", "不接受", "不予退还", "不得参加投标", "取消中标资格"]


def keyword_hits(chunks, question):
    """列举型问题（有哪些/列出/全部）→ 把所有字面命中的块都捞出来"""
    is_list_q = any(x in question for x in ["哪些", "那些", "列出", "全部", "所有", "条款"])
    if not is_list_q:
        return []
    hits = []
    for i, c in enumerate(chunks):
        if any(kw in c for kw in LIST_KWS):
            hits.append(i)
    return hits

def build_queries(q):
    qs = [q]
    for k, words in EXPAND.items():
        if k in q:
            qs.extend(words)
    return list(dict.fromkeys(qs))


def source_of(text):
    if "补充通知" in text or "补充说明" in text:
        return "补充通知"
    return "正文"


def retrieve(question, chunks, vectors):
    queries = build_queries(question)
    print(f"\n📌 扩展出 {len(queries)} 个查询：")
    for i, q in enumerate(queries, 1):
        print(f"   {i}. {q}")

    qvecs = embed_many(queries)

    best = {}
    for qv in qvecs:
        for i, v in enumerate(vectors):
            s = cosine(qv, v)
            if s < MIN_SCORE:
                continue
            if i not in best or s > best[i][0]:
                cnt = best[i][1] if i in best else 0
                best[i] = (s, cnt + 1)

    scored = []
    for i, (s, hits) in best.items():
        bonus = BONUS if source_of(chunks[i]) != "正文" else 0
        scored.append((s + bonus, s, hits, i))
    scored.sort(reverse=True)    
    top = scored[:FINAL_K]

    # —— 列举型问题：关键词命中的块也补进候选 ——
    kw_idx = keyword_hits(chunks, question)
    if kw_idx:
        exist = {i for _, _, _, i in top}
        extra = [(MIN_SCORE, MIN_SCORE, 1, i) for i in kw_idx if i not in exist]
        print(f"   ⌨️ 列举型问题：关键词再补 {len(extra)} 块（合计 {len(top) + len(extra)} 块）")
        top = top + extra

    return top


SYSTEM = """你是投标决策助手。用户会给你「问题」和「从招标文件中检索到的原文片段」。

【回答规则 —— 必须严格遵守】
1. 只使用片段中的信息回答，不许用常识补充，不许编造
2. 每条结论后标注出处（片段里给出的章节名），格式必须是"章节名 + 条款号"
   （例如：第一章 招标公告 三、第7项 / 投标人须知前附表 3.4.3）
   【禁止】只写"片段1""片段2"这类编号，编号不是出处。
   如果片段里确实找不到章节号，写"原文未见明确条款号"。
3. 如果片段之间互相冲突（例如正文写 80 万，补充通知写 100 万），
   以补充通知 / 澄清答疑为准，并明确提示"原正文条款已被补充通知修改"
4. 如果片段不足以回答，直接回答："检索到的资料不足以回答"，并说明还缺什么。
   绝对不许硬凑一个看起来合理的答案
5. 用要点列，简洁，不要客套话。
6. 业务常识（仅在必要时使用，并必须注明是推断）：
   招标文件若表述为"投标保证金应在投标截止时间前到账"，
   而补充通知顺延了投标截止时间，则保证金缴纳截止时间同样顺延。
   按此推断时，请写明"依据投标截止时间顺延推断"，不要当成原文结论。"""


def ask_ds(question, picked, chunks):
    if not DS_KEY or DS_KEY.startswith("填"):
        return "（DeepSeek Key 为空，跳过生成）"

    ctx = []
    for n, (score, raw, hits, i) in enumerate(picked, 1):
        ctx.append(f"【片段{n}｜来源：{source_of(chunks[i])}｜相似度 {raw:.3f}】\n{chunks[i]}")
    user = "【问题】\n" + question + "\n\n【检索到的原文片段】\n" + "\n\n".join(ctx)

    resp = http_post(
        DS_URL,
        {
            "model": DS_MODEL,
            "messages": [
                {"role": "system", "content": SYSTEM},
                {"role": "user", "content": user},
            ],
            "max_tokens": 1500,
        },
        DS_KEY,
    )
    return resp["choices"][0]["message"]["content"]


def self_check():
    print("─" * 60)
    print("① 检查向量文件 …")
    if not os.path.exists(VEC_FILE):
        print(f"   ❌ 找不到：{VEC_FILE}")
        return None
    d = json.load(open(VEC_FILE, encoding="utf-8"))
    print(f"   ✅ {len(d['chunks'])} 块，{len(d['vectors'][0])} 维")

    print("② 检查硅基流动 Key …")
    if EMBED_KEY.startswith("填"):
        print("   ❌ 你还没填 EMBED_KEY，去第 18 行改成你的真 Key")
        return None
    try:
        v = embed_many(["自检"])
        print(f"   ✅ 调用成功，向量维度 {len(v[0])}")
    except Exception as e:
        print(f"   ❌ {e}")
        return None

    print("③ 检查 DeepSeek Key …")
    ok = DS_KEY and not DS_KEY.startswith("填")
    print(f"   {'✅ 已从配置文件读取' if ok else '❌ 没读到，生成环节会跳过'}")
    print("─" * 60)
    return d


def main():
    d = self_check()
    if d is None:
        return
    chunks, vectors = d["chunks"], d["vectors"]

    while True:
        question = input("\n你的问题（q 退出）：").strip()
        if not question or question.lower() == "q":
            break
        try:
            picked = retrieve(question, chunks, vectors)
            if not picked:
                print("\n⚠️ 没有任何块的相似度超过阈值 → 库中没有相关内容")
                continue

            print("\n📊 检索结果：")
            for n, (final, raw, hits, i) in enumerate(picked, 1):
                print(f"  {n}. 第{i+1}块  相似度 {raw:.3f}  命中{hits}次  来源：{source_of(chunks[i])}  {len(chunks[i])}字")

            if picked[0][1] < MIN_SCORE + 0.03:
                print(f"\n⚠️ 最高相似度只有 {picked[0][1]:.3f}，结果可能不可靠")

            print("\n" + "=" * 62)
            print("🤖 回答：\n")
            print(ask_ds(question, picked, chunks))
        except Exception:
            print("\n❌ 出错了，完整信息如下（把这段发给 AI）：\n")
            traceback.print_exc()


if __name__ == "__main__":
    try:
        main()
    except Exception:
        print("\n❌ 启动就出错了：\n")
        traceback.print_exc()
    finally:
        input("\n按回车键退出…")
