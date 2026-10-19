import json
import os
import glob
import urllib.request
import urllib.error

API_KEY = "你的KEY填这里"


# ============================================================
#  工具们（模型可以调用的函数）
# ============================================================

def read_tender_summary(file_path):
    """读取已生成的招标文件提取结果（从缓存读，速度快）"""
    folder = os.path.dirname(file_path)
    base = os.path.splitext(os.path.basename(file_path))[0]
    cache_dir = os.path.join(folder, base + "_缓存")

    if not os.path.exists(cache_dir):
        return ("未找到该文件的提取缓存。请先在主程序中处理过这个文件，"
                "或告诉用户先运行完整分析。")

    parts = []
    # 拼接所有分组汇总结果
    for fp in sorted(glob.glob(os.path.join(cache_dir, "汇总组*.txt"))):
        with open(fp, "r", encoding="utf-8") as f:
            parts.append(f.read())
    # 再补上未被分组的块结果
    if not parts:
        for fp in sorted(glob.glob(os.path.join(cache_dir, "块*.txt"))):
            with open(fp, "r", encoding="utf-8") as f:
                parts.append(f.read())

    text = "\n".join(parts)
    if len(text.strip()) < 50:
        return "缓存内容为空，可能上次未提取成功。"
    # 控制长度，避免超出上下文
    return text[:12000]


def check_company_qualification():
    """查询企业资质档案。没有档案时明确告知无法判断，绝不编造"""
    profile_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                "company_profile.json")

    # ① 没有档案 → 明确拒绝判断
    if not os.path.exists(profile_path):
        return ("【无企业档案】当前没有录入企业资质信息，我无法判断是否满足投资格。\n"
                "请先提供以下信息：\n"
                "1. 施工资质（类别 / 等级 / 有效期）\n"
                "2. 安全生产许可证（有效期）\n"
                "3. 类似业绩（项目名称 / 合同额 / 竣工时间）\n"
                "4. 注册建造师（专业 / 等级 / 是否有在建项目）\n"
                "5. 近三年财务状况\n"
                "6. 信用记录\n"
                "提供后我可以帮你逐条对照招标文件要求。")

    # ② 有档案 → 返回真实数据，并提示核对
    try:
        with open(profile_path, "r", encoding="utf-8") as f:
            profile = json.load(f)
    except Exception as e:
        return f"读取企业档案失败：{e}"

    return ("【以下为企业档案数据，请确认是否为最新】\n"
            + json.dumps(profile, ensure_ascii=False, indent=2))


def calc_bidding_margin(limit_wan, cost_ratio=0.85):
    """粗略测算投标利润空间
    limit_wan: 最高投标限价（万元）
    cost_ratio: 预估成本占限价的比例（默认85%）
    """
    try:
        limit = float(limit_wan)
        cost = limit * float(cost_ratio)
        profit = limit - cost
        return (f"最高投标限价 {limit:.1f} 万元\n"
                f"按成本率 {float(cost_ratio)*100:.0f}% 估算成本 {cost:.1f} 万元\n"
                f"预估毛利 {profit:.1f} 万元（毛利率 {profit/limit*100:.1f}%）\n"
                f"提示：这是粗估，实际需结合工程量清单逐项测算。")
    except Exception as e:
        return f"测算失败：{e}"


# ============================================================
#  工具说明书
# ============================================================
tools = [
    {
        "type": "function",
        "function": {
            "name": "read_tender_summary",
            "description": "读取某个招标文件已提取的关键信息（资格要求、时间节点、限价、废标条款、工程量清单等）。用户询问某个招标文件的具体内容时使用。",
            "parameters": {
                "type": "object",
                "properties": {
                    "file_path": {
                        "type": "string",
                        "description": "招标文件的完整路径"
                    }
                },
                "required": ["file_path"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "check_company_qualification",
            "description": "查询本公司当前的资质、业绩、建造师、财务等情况，用于判断是否满足招标文件的资格要求。",
            "parameters": {
                "type": "object",
                "properties": {},
                "required": []
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "calc_bidding_margin",
            "description": "根据最高投标限价粗略测算利润空间",
            "parameters": {
                "type": "object",
                "properties": {
                    "limit_wan": {
                        "type": "number",
                        "description": "最高投标限价，单位万元"
                    },
                    "cost_ratio": {
                        "type": "number",
                        "description": "预估成本占限价的比例，0-1之间，默认0.85"
                    }
                },
                "required": ["limit_wan"]
            }
        }
    }
]

# 函数名 → 真实函数 的映射
TOOL_FUNCS = {
    "read_tender_summary": read_tender_summary,
    "check_company_qualification": check_company_qualification,
    "calc_bidding_margin": calc_bidding_margin,
}


# ============================================================
#  调模型
# ============================================================
def call_llm(messages):
    url = "https://api.deepseek.com/chat/completions"
    data = {
        "model": "deepseek-chat",
        "messages": messages,
        "tools": tools
    }
    body = json.dumps(data).encode("utf-8")
    req = urllib.request.Request(url, data=body, method="POST")
    req.add_header("Content-Type", "application/json")
    req.add_header("Authorization", "Bearer " + API_KEY)
    with urllib.request.urlopen(req) as resp:
        return json.loads(resp.read().decode("utf-8"))


# ============================================================
#  ⭐ ReAct 循环（这就是 Agent 的心脏）
# ============================================================
SYSTEM_PROMPT = """你是一个投标辅助助手。你可以调用工具获取信息。

【工具使用】
1. 用户问招标文件内容 → 调用 read_tender_summary
2. 需要判断资质是否满足 → 调用 check_company_qualification
3. 需要算利润空间 → 调用 calc_bidding_margin

【严谨性要求 —— 这是最高优先级】
4. 如果 check_company_qualification 返回"无企业档案"，
   你必须明确回答："我无法判断是否满足要求，因为缺少企业资质信息。"
   绝对不允许用假设、示例或默认数据代替真实数据，
   绝对不允许在这种情况下给出"建议投标/不建议投标"的结论。

5. 任何判断都必须说明依据。例如：
   "招标文件要求市政叁级及以上，档案显示为市政叁级 → 满足"

6. 工具没有返回的信息，就说"未获取到"，不要编造。

7. 最终回答必须包含一句提示：
   "本结论基于所提供的信息，请以原始文件和企业证件核对。"

【回答格式】
· 能判断时：明确结论 + 逐条依据 + 风险提示
· 不能判断时：明确说明缺少什么数据 + 请用户补充
"""


def run_agent(question, max_rounds=6):
    messages = [
        {"role": "system", "content": SYSTEM_PROMPT},
        {"role": "user", "content": question}
    ]

    for rnd in range(1, max_rounds + 1):
        print(f"\n--- 第 {rnd} 轮 ---")
        try:
            resp = call_llm(messages)
        except urllib.error.HTTPError as e:
            print("❌ API 报错:", e.read().decode("utf-8"))
            return

        msg = resp["choices"][0]["message"]

        # 模型不再需要工具 → 输出最终答案
        if not msg.get("tool_calls"):
            print("\n【最终回答】")
            print(msg.get("content"))
            return msg.get("content")

        # 模型要调工具 → 逐个执行
        messages.append(msg)
        for tc in msg["tool_calls"]:
            name = tc["function"]["name"]
            args = json.loads(tc["function"]["arguments"]) if tc["function"]["arguments"] else {}
            print(f"  🔧 调用工具：{name}({args})")

            func = TOOL_FUNCS.get(name)
            if func is None:
                result = f"错误：没有名为 {name} 的工具"
            else:
                try:
                    result = func(**args)
                except Exception as e:
                    result = f"工具执行出错：{e}"

            print(f"     → 返回 {len(str(result))} 字")
            messages.append({
                "role": "tool",
                "tool_call_id": tc["id"],
                "content": str(result)
            })

    print("\n⚠️ 达到最大轮数，模型仍未给出最终答案")
    return None


# ============================================================
#  运行
# ============================================================
if __name__ == "__main__":
    print("=" * 55)
    print("投标决策 Agent（输入 quit 退出）")
    print("示例：帮我看看 D:\\新建文件夹\\可以.docx 值不值得投")
    print("=" * 55)

    while True:
        q = input("\n你：").strip()
        if q.lower() == "quit":
            break
        if not q:
            continue
        run_agent(q)