#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
合同审查脚本（合同审查智能体 / C002 FR-07）
纯标准库实现，无外部依赖，满足"数据不出域、开箱即用、离线可用"。

用法：
  python review.py <合同.txt|目录/> <output_dir> [--docx]
    输入可为单个 .txt/.md 文件，或包含多个 .txt/.md 的目录（批量）。
    --docx  额外生成可打印的 .docx 审查报告（零依赖 OpenXML）。

功能：条款拆解 -> 规则引擎（风险表述/权利义务失衡/缺失条款/表述歧义）
      -> 风险分级(高/中/低) -> 条款级批注审查报告 + 批次汇总。
依据规则见 references/risk_rules.md（须律师核对现行有效条文）。
"""
import os
import re
import sys
import glob
import datetime
import zipfile

# ---------- 风险规则（与 references/risk_rules.md 保持一致） ----------
# 每条：(名称, 正则, 级别, 风险说明, 修改建议, 依据)
PATTERN_RULES = [
    ("绝对化免责", r"概不负责|不承担任何责任|免除全部责任|一切责任由.*承担", "高",
     "一方完全免责，可能构成加重对方责任、排除对方主要权利的格式条款", "明确责任范围与限额，删除绝对化免责表述", "《民法典》第497条"),
    ("最终解释权", r"最终解释权|保留最终解释", "高",
     "'最终解释权归本方'属无效格式条款，损害对方权益", "删除；改为'双方协商或依法解释'", "《民法典》第498条"),
    ("概括放弃权利", r"无条件.*放弃|不可撤销地放弃|放弃一切.*权利", "高",
     "概括性放弃权利，可能显失公平", "限定具体情形与范围", "《民法典》第151条"),
    ("空白金额", r"_{2,}\s*元|[xX×＊*]{2,}\s*元|金额待定|价款待定|(?<![0-9])若干元", "高",
     "关键金额未确定，履行无依据、易生争议", "填写确定金额或明确计算方式", "《民法典》第470条"),
    ("空白日期", r"_{2,}\s*年|_{2,}\s*月|_{2,}\s*日|日期待定|[xX×]{4}\s*年", "高",
     "关键时间未确定，履行期限不明", "填写确定日期或约定确定方法", "《民法典》第470条、第511条"),
    ("单方解除", r"有权单方解除|单方随时终止|可随时解除本合同", "中",
     "单方解除权配置失衡，需审查是否对等", "明确解除条件、通知期与对等安排", "《民法典》第562/563条"),
    ("单方变更", r"有权单方变更|单方(调整|修改)本?合同|有权随时调整", "中",
     "单方变更条款失衡", "变更须双方书面同意", "《民法典》第543条"),
    ("定金过高", r"定金", "中",
     "定金比例若超过主合同标的额20%，超过部分不产生定金效力", "定金比例不超过标的额20%", "《民法典》第586条"),
    ("自动续约", r"自动续(约|签)|自动展期|到期自动延长", "低",
     "自动续约缺少退出通知机制", "增加到期前书面通知/异议机制", "《民法典》第510条"),
]

# 违约金/利息缺标准：命中"违约金/逾期利息"但同段无比例数字
PENALTY_KEY = re.compile(r"违约金|逾期利息")
PENALTY_STD = re.compile(r"\d+\s*(%|％|‰|元|倍)|日\s*\d|月\s*\d|年利率|百分之")

# ---------- 缺失条款规则 ----------
# 关键条款 -> (检测正则, 级别, 说明)
MISSING_RULES = [
    ("违约责任", r"违约责任|违约金|承担.*责任", "高", "无违约责任，守约方救济困难"),
    ("争议解决/管辖", r"争议(解决|处理)|管辖|仲裁|诉讼", "高", "未约定管辖或仲裁，争议解决路径不明"),
    ("付款/履行期限", r"付款|支付.*期|履行期|交付.*(日|时间)|付清", "高", "无期限约定，易生履行争议"),
    ("标的/数量/质量", r"标的|数量|质量|规格|货物|服务内容", "高", "合同主要内容可能缺失"),
    ("不可抗力", r"不可抗力", "中", "未约定不可抗力，风险分担不明"),
    ("保密条款", r"保密|商业秘密", "中", "涉及商业信息时缺失保密约定"),
    ("合同生效条件", r"生效|自.*签(字|章)之日起|盖章.*生效", "中", "生效时点/条件不明"),
    ("通知与送达", r"通知|送达|电子邮件|地址.*变更", "低", "未约定有效送达方式"),
]

LEVEL_ORDER = {"高": 0, "中": 1, "低": 2}
LEVEL_TAG = {"高": "🔴高", "中": "🟠中", "低": "🟡低"}

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
PKG_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
DOC_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def safe_filename(name):
    return re.sub(r'[\\/:*?"<>|\r\n\t ]+', '_', name)


def split_clauses(text):
    """按 '第X条' / '一、' / '1.' / 空行段落 切分条款。返回 [(编号标签, 内容)]。"""
    lines = [l.rstrip() for l in text.split("\n")]
    clauses = []
    cur_label, cur_buf = None, []
    head_re = re.compile(r'^\s*(第\s*[一二三四五六七八九十百零0-9]+\s*条|[一二三四五六七八九十]+、|\d+[\.、]\s*)')

    def flush():
        if cur_buf:
            content = "\n".join(cur_buf).strip()
            if content:
                clauses.append((cur_label or "前言/未编号", content))

    for l in lines:
        if not l.strip():
            continue
        m = head_re.match(l)
        if m:
            flush()
            cur_label = m.group(1).strip()
            cur_buf = [l.strip()]
        else:
            cur_buf.append(l.strip())
    flush()
    if not clauses:  # 无任何条款头，则整体作为一个单元
        clauses = [("全文", text.strip())]
    return clauses


def review_clause(label, content):
    """对单个条款执行规则，返回 findings 列表。"""
    findings = []
    for name, pat, level, desc, sug, basis in PATTERN_RULES:
        if re.search(pat, content):
            findings.append((level, name, desc, sug, basis))
    # 违约金/利息缺标准
    if PENALTY_KEY.search(content) and not PENALTY_STD.search(content):
        findings.append(("中", "违约金/利息标准缺失",
                         "约定了违约金/逾期利息但未见明确比例或计算标准",
                         "明确违约金/利息的计算标准，避免过高(超损失30%可被调整)或过低",
                         "《民法典》第585条"))
    return findings


def review_contract(text):
    """返回 (clause_findings, missing_findings, counts)。"""
    clauses = split_clauses(text)
    clause_findings = []  # [(label, content_snippet, [findings])]
    for label, content in clauses:
        fs = review_clause(label, content)
        if fs:
            snippet = content if len(content) <= 60 else content[:60] + "…"
            clause_findings.append((label, snippet, fs))

    # 缺失条款检测（全文层面）
    missing = []
    for name, pat, level, desc in MISSING_RULES:
        if not re.search(pat, text):
            missing.append((level, name, desc))

    counts = {"高": 0, "中": 0, "低": 0}
    for _, _, fs in clause_findings:
        for f in fs:
            counts[f[0]] += 1
    for m in missing:
        counts[m[0]] += 1
    return clauses, clause_findings, missing, counts


def build_report_md(name, clauses, clause_findings, missing, counts):
    total = counts["高"] + counts["中"] + counts["低"]
    lines = [
        "# 合同审查报告（AI 初审）",
        "",
        "**合同名称**：%s" % name,
        "**审查时间**：%s" % datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
        "**条款总数**：%d ｜ **风险条款数**：%d" % (len(clauses), len(clause_findings)),
        "**风险统计**：%s 高 ｜ %s 中 ｜ %s 低（合计 %d）"
        % (counts["高"], counts["中"], counts["低"], total),
        "",
        "> ⚠️ 本报告为 AI 辅助初审，仅提示风险与建议，**不替代法律意见**，须由承办律师复核定稿。",
        "",
        "## 一、审查结论",
    ]
    if counts["高"] > 0:
        lines.append("- **存在 %d 项高风险，必须修改后方可签署。**" % counts["高"])
    if counts["中"] > 0:
        lines.append("- 存在 %d 项中风险，建议修改。" % counts["中"])
    if counts["低"] > 0:
        lines.append("- 存在 %d 项低风险，提示关注。" % counts["低"])
    if total == 0:
        lines.append("- 未命中已知风险规则；仍建议律师人工复核（规则库覆盖有限）。")

    lines += ["", "## 二、条款级风险批注"]
    if clause_findings:
        for label, snippet, fs in clause_findings:
            fs_sorted = sorted(fs, key=lambda x: LEVEL_ORDER[x[0]])
            lines.append("")
            lines.append("### %s" % label)
            lines.append("> 原文摘录：%s" % snippet)
            for level, rname, desc, sug, basis in fs_sorted:
                lines.append("- **[%s] %s**：%s" % (LEVEL_TAG[level], rname, desc))
                lines.append("  - 修改建议：%s" % sug)
                lines.append("  - 依据（须律师核对）：%s" % basis)
    else:
        lines.append("")
        lines.append("- 未在条款中命中风险表述规则。")

    lines += ["", "## 三、缺失关键条款"]
    if missing:
        missing_sorted = sorted(missing, key=lambda x: LEVEL_ORDER[x[0]])
        for level, mname, desc in missing_sorted:
            lines.append("- **[%s] 缺失「%s」**：%s" % (LEVEL_TAG[level], mname, desc))
    else:
        lines.append("- 关键条款均已覆盖。")

    lines += [
        "",
        "## 四、总体修改建议",
        "1. 优先处理全部高风险项（免责/最终解释权/空白金额日期/关键缺失条款）。",
        "2. 逐项落实中风险修改（单方权利对等化、违约金标准明确化）。",
        "3. 由承办律师结合交易背景与现行法律，核对上述依据并定稿。",
        "",
        "---",
        "*本报告由合同审查智能体自动生成，为 AI 初审，须经承办律师终审确认。规则依据见 references/risk_rules.md。*",
    ]
    return "\n".join(lines)


# ---------- 零依赖 .docx 生成（OpenXML，复用 lawyer-letter-agent 框架） ----------
def xml_escape(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def bold_runs(text):
    parts = re.split(r'(\*\*.+?\*\*)', text)
    out = []
    for p in parts:
        if len(p) >= 4 and p.startswith("**") and p.endswith("**"):
            out.append('<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">%s</w:t></w:r>'
                       % xml_escape(p[2:-2]))
        elif p:
            out.append('<w:r><w:t xml:space="preserve">%s</w:t></w:r>' % xml_escape(p))
    return "".join(out)


def md_paragraph(text, style=None):
    ppr = ('<w:pPr><w:pStyle w:val="%s"/></w:pPr>' % xml_escape(style)) if style else ""
    return "<w:p>%s%s</w:p>" % (ppr, bold_runs(text))


def md_to_docx_body(md_text):
    body = []
    for raw in md_text.split("\n"):
        line = raw.rstrip()
        if not line.strip():
            continue
        if line.startswith("### "):
            body.append(md_paragraph(line[4:], "Heading3"))
        elif line.startswith("## "):
            body.append(md_paragraph(line[3:], "Heading2"))
        elif line.startswith("# "):
            body.append(md_paragraph(line[2:], "Heading1"))
        elif line.startswith("> "):
            body.append(md_paragraph(line[2:], "Quote"))
        elif re.match(r'^\d+\.\s', line):
            body.append(md_paragraph(re.sub(r'^\d+\.\s', '', line), "ListNumber"))
        elif re.match(r'^\s+- ', line):
            body.append(md_paragraph(line.strip()[2:], "ListBullet2"))
        elif line.startswith("- ") or line.startswith("* "):
            body.append(md_paragraph(line[2:], "ListBullet"))
        else:
            body.append(md_paragraph(line, None))
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr>')
    return "".join(body) + sect


def build_docx(md_text, out_path):
    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="%s"><w:body>%s</w:body></w:document>'
        % (W_NS, md_to_docx_body(md_text))
    )
    styles_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:styles xmlns:w="%s">'
        '<w:docDefaults><w:rPrDefault><w:rPr>'
        '<w:rFonts w:ascii="Calibri" w:eastAsia="\u5b8b\u4f53" w:hAnsi="Calibri" w:cs="Times New Roman"/>'
        '<w:sz w:val="21"/></w:rPr></w:rPrDefault></w:docDefaults>'
        '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>'
        '<w:pPr><w:keepNext/><w:spacing w:before="240" w:after="60"/><w:outlineLvl w:val="0"/></w:pPr>'
        '<w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/>'
        '<w:pPr><w:keepNext/><w:spacing w:before="160" w:after="40"/><w:outlineLvl w:val="1"/></w:pPr>'
        '<w:rPr><w:b/><w:sz w:val="26"/></w:rPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/>'
        '<w:pPr><w:keepNext/><w:spacing w:before="120" w:after="40"/><w:outlineLvl w:val="2"/></w:pPr>'
        '<w:rPr><w:b/><w:sz w:val="23"/></w:rPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/>'
        '<w:pPr><w:ind w:left="480"/><w:spacing w:before="60" w:after="60"/></w:pPr>'
        '<w:rPr><w:i/><w:color w:val="555555"/></w:rPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="ListNumber"><w:name w:val="List Number"/>'
        '<w:pPr><w:ind w:left="480" w:hanging="360"/></w:pPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="ListBullet"><w:name w:val="List Bullet"/>'
        '<w:pPr><w:ind w:left="480" w:hanging="360"/></w:pPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="ListBullet2"><w:name w:val="List Bullet 2"/>'
        '<w:pPr><w:ind w:left="900" w:hanging="360"/></w:pPr></w:style>'
        '</w:styles>' % W_NS
    )
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="%s">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
        '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
        '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
        '</Types>' % PKG_NS
    )
    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="%s">'
        '<Relationship Id="rId1" Type="%s/officeDocument" Target="word/document.xml"/>'
        '<Relationship Id="rId2" Type="%s/metadata/core-properties" Target="docProps/core.xml"/>'
        '<Relationship Id="rId3" Type="%s/extended-properties" Target="docProps/app.xml"/>'
        '</Relationships>' % (PKG_NS, DOC_NS, PKG_NS, PKG_NS)
    )
    doc_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="%s">'
        '<Relationship Id="rId1" Type="%s/styles" Target="styles.xml"/>'
        '</Relationships>' % (PKG_NS, DOC_NS)
    )
    now = datetime.datetime.now().strftime("%Y-%m-%dT%H:%M:%SZ")
    core_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        '<dc:title>\u5408\u540c\u5ba1\u67e5\u62a5\u544a</dc:title>'
        '<dcterms:created xsi:type="dcterms:W3CDTF">%s</dcterms:created>'
        '<dcterms:modified xsi:type="dcterms:W3CDTF">%s</dcterms:modified>'
        '</cp:coreProperties>' % (now, now)
    )
    app_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
        '<Application>ContractReviewAgent</Application></Properties>'
    )
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", root_rels)
        z.writestr("word/document.xml", document_xml)
        z.writestr("word/styles.xml", styles_xml)
        z.writestr("word/_rels/document.xml.rels", doc_rels)
        z.writestr("docProps/core.xml", core_xml)
        z.writestr("docProps/app.xml", app_xml)


# ---------- 主流程 ----------
def collect_inputs(path):
    if os.path.isdir(path):
        files = []
        for ext in ("*.txt", "*.md"):
            files.extend(glob.glob(os.path.join(path, ext)))
        return sorted(files)
    return [path]


def main():
    args = sys.argv[1:]
    make_docx = "--docx" in args
    args = [a for a in args if a != "--docx"]
    if len(args) < 2:
        print("用法: python review.py <合同.txt|目录/> <output_dir> [--docx]")
        sys.exit(1)
    inp, outdir = args[0], args[1]
    if not os.path.exists(inp):
        print("输入不存在: %s" % inp)
        sys.exit(1)
    os.makedirs(outdir, exist_ok=True)

    files = collect_inputs(inp)
    if not files:
        print("未找到 .txt/.md 合同文件: %s" % inp)
        sys.exit(1)

    summary = []
    for f in files:
        name = os.path.splitext(os.path.basename(f))[0]
        with open(f, encoding="utf-8-sig") as fh:
            text = fh.read()
        clauses, clause_findings, missing, counts = review_contract(text)
        report = build_report_md(name, clauses, clause_findings, missing, counts)
        base = safe_filename(name + "_审查报告")
        with open(os.path.join(outdir, base + ".md"), "w", encoding="utf-8") as w:
            w.write(report)
        if make_docx:
            build_docx(report, os.path.join(outdir, base + ".docx"))
        summary.append((name, counts["高"], counts["中"], counts["低"]))

    lines = [
        "# 合同审查批次汇总",
        "",
        "审查时间：%s" % datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
        "合同份数：%d" % len(files),
        "格式：%s" % ("Markdown + DOCX" if make_docx else "Markdown"),
        "",
        "| 合同 | 高风险 | 中风险 | 低风险 | 建议 |",
        "|---|---|---|---|---|",
    ]
    for name, h, m, l in summary:
        advice = "必须修改后签署" if h > 0 else ("建议修改" if m > 0 else "提示关注")
        lines.append("| %s | %d | %d | %d | %s |" % (name, h, m, l, advice))
    lines += ["", "---", "*本批次均为 AI 初审，须经承办律师终审确认。*"]
    with open(os.path.join(outdir, "_审查汇总.md"), "w", encoding="utf-8") as w:
        w.write("\n".join(lines))

    print("已审查 %d 份合同至 %s%s" % (len(files), outdir, "（含 .docx）" if make_docx else ""))
    for name, h, m, l in summary:
        print("  - %s：高%d 中%d 低%d" % (name, h, m, l))


if __name__ == "__main__":
    main()
