#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""机械审计：对照报价提取文本，扫描客户行程 DOCX 整个包（正文、页眉页脚、docProps、批注…）。

用法：
    python -X utf8 scripts/audit_docx.py 行程.docx _quote_extract.txt [--allow-pricing] [--forbid 词1 词2 ...]

检查：
  1. 供应商联系方式：报价里出现的电话 / QQ / 微信 / 邮箱 / 地址关键字是否泄露到 DOCX
  2. 价格：RMB / 元 / 单房差 / 小费 / 16免1 …；报价「餐名+数字」的餐标数字是否残留
  3. 承诺类用语：含机票 / 纯玩 / 无购物 / 赠送 … 仅当报价原文有才放行
  4. 【景点】名（去括号附注）必须出现在报价中
  5. 列出正文中所有 2–4 位数字，供人工逐个确认来源

这只是机械初筛，不能代替第三方独立审计（见 SKILL.md「独立审计」）。退出码 1 = 有命中。
"""
import argparse
import re
import sys
import zipfile

PRICE_TERMS = ['RMB', '人民币', '元/人', '元/', '团费', '单房差', '单间差', '房差', '小费', 'FOC',
               '16免1', '报价', '自费', '代开费', '车差']
PROMISE_TERMS = ['含机票', '含保险', '含高铁', '含门票', '含于团费', '无必加', '纯玩', '无购物', '不进店',
                 '赠送', '矿泉水', '啤酒', '饮料', '特产', '购物', '导游', '领队', '升级', '免费']
SUPPLIER_KEYS = ['旅行社', 'Travel Agency', 'TEL', '微信', 'QQ', 'ADD', 'OP', '负责人']


def package_text(path):
    out = {}
    with zipfile.ZipFile(path) as z:
        for n in z.namelist():
            if n.endswith('.xml'):
                xml = z.read(n).decode('utf-8', 'ignore')
                out[n] = re.sub(r'<[^>]+>', '', xml)
    return out


def body_text(path):
    from docx import Document
    return '\n'.join(p.text for p in Document(path).paragraphs)


def supplier_tokens(quote):
    toks = set(re.findall(r'\d{7,}', quote))                       # 电话 / QQ / 微信号
    toks |= set(re.findall(r'[\w.+-]+@[\w-]+\.[\w.]+', quote))       # 邮箱
    # 抬头通常是供应商信息；「贵 州 蓝 狮 旅 行 社」这类字间空格先去掉。
    # 「致(TO)」后面是收件方（华宇 / 我们自己），不算供应商。
    head = re.sub(r'(?<=[一-鿿]) (?=[一-鿿])', '', quote[:800])
    for name in re.findall(r'[一-鿿]{2,}?(?:旅行社|旅游公司)', head):
        core = re.sub(r'(?:旅行社|旅游公司)$', '', name)   # 贵州蓝狮旅行社 → 贵州蓝狮
        toks |= {name, core, core[-2:]}                     # 再加品牌短名「蓝狮」
    toks |= set(re.findall(r'[A-Z][A-Za-z ]{3,}?(?:Travel Agency|Travel|Tours)', head))
    return {t.strip() for t in toks if len(t.strip()) >= 2}


def meal_standards(quote):
    # 「午：夜郎风味50」→ 夜郎风味50；客户稿只能出现餐名，不能带数字
    return set(re.findall(r'[早午晚]：\s*([一-鿿]+\d+)', quote))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('docx')
    ap.add_argument('quote')
    ap.add_argument('--allow-pricing', action='store_true')
    ap.add_argument('--forbid', nargs='*', default=[])
    a = ap.parse_args()

    quote = open(a.quote, encoding='utf-8').read()
    parts = package_text(a.docx)
    body = body_text(a.docx)
    hits = []

    def scan(term, why, only_body=False):
        for name, text in parts.items():
            if only_body and name != 'word/document.xml':
                continue
            i = text.find(term)
            if i >= 0:
                hits.append(f'[{why}] {name}: …{text[max(0, i - 15):i + len(term) + 15]}…')

    for t in sorted(supplier_tokens(quote)):
        scan(t, '供应商信息')
    for k in SUPPLIER_KEYS:
        scan(k, '供应商关键字?', only_body=True)
    if not a.allow_pricing:
        for t in PRICE_TERMS:
            scan(t, '价格')
        for m in sorted(meal_standards(quote)):
            scan(m, '餐标数字')
    for t in PROMISE_TERMS:
        if t in body:
            hits.append(f'[承诺用语，报价{"有" if t in quote else "无"}此词] {t}')
    for t in a.forbid:
        scan(t, '自定义禁词')

    for name in re.findall(r'【(.*?)】', body):
        core = re.sub(r'（.*?）|\(.*?\)', '', name).strip()
        if core and core not in quote:
            hits.append(f'[景点不在报价] 【{name}】')

    nums = sorted({m.group() for m in re.finditer(r'(?<![\d.])\d{2,4}(?![\d])', body)})

    print('=== 机械审计 ===')
    print('\n'.join(hits) if hits else '无命中')
    print('--- 正文数字（人工确认来源：车次/时刻/景点事实 OK，餐标/价格 不 OK）---')
    print(' '.join(nums))
    sys.exit(1 if hits else 0)


if __name__ == '__main__':
    main()
