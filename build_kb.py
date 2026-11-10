# -*- coding: utf-8 -*-
"""天津财经大学会计学博士面试知识库 — 构建脚本
从两份 docx 导出的 txt 中解析结构化数据，生成自包含单文件 index.html。
"""
import re
import json
import os

HERE = os.path.dirname(os.path.abspath(__file__))
HANDBOOK = "/tmp/handbook.txt"
ZHISHIDIAN = "/tmp/zhishidian.txt"


# ---------------------------------------------------------------- 解析手册
def parse_handbook(path):
    with open(path, encoding="utf-8") as f:
        raw = f.read()
    raw = raw.replace("\ufeff", "").replace("\x0c", "")
    lines = raw.split("\n")

    modules = []
    bank = []
    corrections = []
    mocks = []
    plan7 = []

    cur_module = None
    cur_q = None
    cur_section = None
    cur_corr = None
    cur_mock = None

    section = "head"  # head | modules | bank | corr | mocks | plan

    i = 0
    n = len(lines)
    while i < n:
        line = lines[i].strip()

        # ---- 段落分界 ----
        if line.startswith("六、上传"):
            section = "bank"
            i += 1
            continue
        if line.startswith("七、原始题库"):
            section = "corr"
            i += 1
            continue
        if line.startswith("八、12道现场"):
            section = "mocks"
            i += 1
            continue
        if line.startswith("九、入校面试前7天"):
            section = "plan"
            i += 1
            continue
        if line.startswith("十、公开资料"):
            section = "end"
            i += 1
            continue

        if section == "head":
            # 模块标题
            m = re.match(r"^模块([一二三四五六七八九十]+)｜(.+)$", line)
            if m:
                # 下一非空行为描述
                desc = ""
                j = i + 1
                while j < n and not lines[j].strip():
                    j += 1
                if j < n:
                    desc = lines[j].strip()
                cur_module = {"name": line, "desc": desc, "questions": []}
                modules.append(cur_module)
                i = j + 1
                section = "modules"
                continue
            i += 1
            continue

        if section == "modules":
            m = re.match(r"^模块([一二三四五六七八九十]+)｜(.+)$", line)
            if m:
                desc = ""
                j = i + 1
                while j < n and not lines[j].strip():
                    j += 1
                if j < n:
                    desc = lines[j].strip()
                cur_module = {"name": line, "desc": desc, "questions": []}
                modules.append(cur_module)
                i = j + 1
                continue
            m = re.match(r"^(\d{2}-\d{2})\s+(.+)$", line)
            if m:
                cur_q = {"id": m.group(1), "title": m.group(2).strip(),
                         "answer": "", "followup": ""}
                cur_module["questions"].append(cur_q)
                i += 1
                continue
            m = re.match(r"^答题要点｜(.*)$", line)
            if m and cur_q is not None:
                cur_q["answer"] = m.group(1).strip()
                i += 1
                continue
            m = re.match(r"^追问/易错｜(.*)$", line)
            if m and cur_q is not None:
                cur_q["followup"] = m.group(1).strip()
                i += 1
                continue
            i += 1
            continue

        if section == "bank":
            if line in ("会计篇", "审计篇", "财管篇", "税法篇", "其他情境题补充"):
                cur_section = {"name": line, "questions": []}
                bank.append(cur_section)
                i += 1
                continue
            if cur_section is not None:
                m = re.match(r"^(\d{3})\.\s*(.+)$", line)
                if m:
                    cur_section["questions"].append(m.group(2).strip())
                    i += 1
                    continue
                m = re.match(r"^(\d+)、(.+)$", line)
                if m and cur_section["name"] == "其他情境题补充":
                    cur_section["questions"].append(m.group(2).strip())
                    i += 1
                    continue
            i += 1
            continue

        if section == "corr":
            m = re.match(r"^纠错(\d+)｜(.+)$", line)
            if m:
                cur_corr = {"id": m.group(1), "title": m.group(2).strip(), "text": ""}
                corrections.append(cur_corr)
                i += 1
                continue
            m = re.match(r"^规范表述：(.*)$", line)
            if m and cur_corr is not None:
                cur_corr["text"] = m.group(1).strip()
                i += 1
                continue
            i += 1
            continue

        if section == "mocks":
            m = re.match(r"^模拟题(\d+)｜(.+)$", line)
            if m:
                cur_mock = {"id": m.group(1), "title": m.group(2).strip(), "text": ""}
                mocks.append(cur_mock)
                i += 1
                continue
            m = re.match(r"^回答路线：(.*)$", line)
            if m and cur_mock is not None:
                cur_mock["text"] = m.group(1).strip()
                i += 1
                continue
            i += 1
            continue

        if section == "plan":
            # 表头：时间/任务/可验收成果
            if line in ("时间", "任务", "可验收成果"):
                i += 1
                continue
            m = re.match(r"^(D-\d+)$", line)
            if m:
                day = m.group(1)
                # 收集后续两行：任务 / 成果
                task = ""
                result = ""
                j = i + 1
                while j < n and not lines[j].strip():
                    j += 1
                if j < n:
                    task = lines[j].strip()
                k = j + 1
                while k < n and not lines[k].strip():
                    k += 1
                if k < n:
                    result = lines[k].strip()
                plan7.append({"day": day, "task": task, "result": result})
                i = k + 1
                continue
            i += 1
            continue

        i += 1

    return modules, bank, corrections, mocks, plan7


# ---------------------------------------------------------------- 解析知识点解释
NOISE_ZH = {
    "财政部考试中心", "国家税务总局", "财政部", "维基文库，自由的图书馆",
    "中国证券监督管理委员会", "中国注册会计师协会", "哥伦比亚大学",
    "推荐重点阅读", "建议重点掌握的经典文献", "建议重点掌握的英文顶刊文献",
    "建议重点阅读", "经典文献（建议面试前重点阅读）", "权威参考：",
    "官方准则原文：",
}


def parse_explanations(path):
    with open(path, encoding="utf-8") as f:
        raw = f.read()
    raw = raw.replace("\ufeff", "").replace("\x0c", "")
    blocks = raw.split("[网页搜索]")
    exp = {}
    for b in blocks:
        m = re.search(r"(\d{2}-\d{2})", b)
        if not m:
            continue
        qid = m.group(1)
        if qid not in ("03-15", "05-12", "10-04", "10-05", "10-09", "10-10"):
            continue
        clean = []
        for l in b.split("\n"):
            s = l.strip()
            if not s:
                continue
            if "chatgpt-system-hint" in s:
                continue
            if s.startswith("思考了"):
                continue
            if s in NOISE_ZH:
                continue
            if len(s) < 20 and s.isascii():
                continue
            s = re.sub(r"^[•\-–—\t ]+", "", s).strip()
            if not s:
                continue
            s = re.sub(r"\\text\{([^}]*)\}", r"\1", s)
            s = s.replace("\\[", "").replace("\\]", "").replace("\\(", "").replace("\\)", "")
            s = s.replace("\\boxed", "").replace("\\begin{aligned}", "").replace("\\end{aligned}", "")
            s = s.strip()
            if not s:
                continue
            clean.append(s)
        exp[qid] = "\n".join(clean)
    return exp


# ---------------------------------------------------------------- 概览（静态）
OVERVIEW = {
    "weights": [
        {"k": "财务会计/准则/合并", "w": "30%", "v": "定义、准则条件、分录、三表影响及高频易错点"},
        {"k": "成本管理会计/财管", "w": "25%", "v": "作业成本、标准成本、投资决策、估值及预算"},
        {"k": "审计/税法/内控", "w": "20%", "v": "审计证据、风险导向、2026税法、治理"},
        {"k": "会计理论/研究方法", "w": "15%", "v": "可检验机制、研究识别、文章贡献、科研规划"},
        {"k": "试讲/英语/综合素质", "w": "10%", "v": "教学设计、表达、互动、英文术语及学院匹配"},
    ],
    "struct30": "①概念或结论（10秒）→ ②核心机制/公式/适用条件（15秒）→ ③一个边界或实践意义（5秒）。",
    "struct90": "①问题来自什么业务情境 → ②准则或理论如何界定 → ③明确公式/分录或模型 → ④给一个简明例子 → ⑤指出容易误判的例外、与研究或治理的联系。",
    "lectures": [
        {"t": "投资性房地产：两种计量模式", "d": "出租写字楼情境→成本法/公允价值法、转换规则及账务案例→政策选择与利润影响"},
        {"t": "作业成本法 ABC", "d": "传统工时扭曲成本→资源→作业→动因→产品、12万元算例→ABC与传统法对照"},
        {"t": "标准成本法", "d": "标准/实际用量场景→材料价格与数量差异→责任中心和行为扭曲"},
        {"t": "盈余管理", "d": "业绩阈值→应计与真实盈余管理对比→研究测度、合法选择与舞弊边界"},
        {"t": "NPV与IRR", "d": "投资回收情境→现金流、NPV、IRR冲突→跨期价值和不确定性"},
    ],
}


# ---------------------------------------------------------------- 去学校化
_REPLACEMENTS = [
    ("天津财经大学会计学博士教师岗位面试", "会计学博士高校教职面试"),
    ("天津财经大学会计学博士教师面试", "会计学博士高校教职面试"),
    ("天津财经大学会计学教师面试", "会计学高校教职面试"),
    ("天津财经大学教师面试", "高校教师面试"),
    ("天津财经大学教师试讲", "高校教师试讲"),
    ("天津财经大学会计学院", "会计学院"),
    ("为什么选择天津财经大学", "为什么选择本学院"),
    ("天津财经大学", "高校"),
    ("针对天财高校教师岗位", "针对高校教师岗位"),
    ("结合天财一流课程", "结合一流课程"),
    ("天财一流课程", "一流课程"),
    ("天财", "高校"),
    ("天津离家近", "离家近"),
]


def normalize(text):
    if not isinstance(text, str):
        return text
    for a, b in _REPLACEMENTS:
        text = text.replace(a, b)
    return text


def normalize_tree(obj):
    if isinstance(obj, str):
        return normalize(obj)
    if isinstance(obj, list):
        return [normalize_tree(x) for x in obj]
    if isinstance(obj, dict):
        return {k: normalize_tree(v) for k, v in obj.items()}
    return obj


def main():
    modules, bank, corrections, mocks, plan7 = parse_handbook(HANDBOOK)
    explanations = parse_explanations(ZHISHIDIAN)

    # 把详解挂到对应问题
    for mod in modules:
        for q in mod["questions"]:
            if q["id"] in explanations:
                q["detail"] = explanations[q["id"]]

    data = {
        "overview": OVERVIEW,
        "modules": modules,
        "bank": bank,
        "corrections": corrections,
        "mocks": mocks,
        "plan7": plan7,
    }
    data = normalize_tree(data)

    # 校验统计
    total_q = sum(len(m["questions"]) for m in modules)
    total_bank = sum(len(s["questions"]) for s in bank)
    print("模块数:", len(modules))
    print("规范答题题目数:", total_q)
    print("详解数:", len(explanations), list(explanations.keys()))
    print("真题各篇:", {s["name"]: len(s["questions"]) for s in bank}, "共", total_bank)
    print("纠错卡:", len(corrections), "模拟题:", len(mocks), "7天计划:", len(plan7))

    json_str = json.dumps(data, ensure_ascii=False).replace("</", "<\\/")

    template = open(os.path.join(HERE, "template.html"), encoding="utf-8").read()
    out = template.replace("__DATA_JSON__", json_str)

    out_path = os.path.join(HERE, "index.html")
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(out)
    print("已生成:", out_path, "大小 %.1f KB" % (os.path.getsize(out_path) / 1024))


if __name__ == "__main__":
    main()
