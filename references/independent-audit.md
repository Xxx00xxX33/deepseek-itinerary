# Third-party independent audit

Mandatory before any itinerary DOCX is delivered (SKILL.md step 7).

## Why a separate auditor

Whoever wrote the copy reads it the way they meant it. A fresh agent with no generation context reads it the way a client will — and catches implied services, meal-slot drift and leaked metadata the writer's own checklist passed. (贵州7天 case: the self-check passed; a closer read still found 「还能穿行水帘洞」「银饰叮当」「入夜后灯火通明」「早：酒店早餐」 — none in the quote.)

## Procedure

1. Run the mechanical scan first and fix its hits:
   `python -X utf8 scripts/audit_docx.py "行程.docx" _quote_extract.txt`
2. Spawn a **fresh** subagent (Agent tool, general-purpose, background is fine). Do **not** fork the conversation and do **not** pass your own reasoning or "what I changed" — only the inputs below.
3. While it runs, don't edit the DOCX it is reading.
4. Fix every 违规 / 风险, regenerate, rerun step 1, then send the auditor the changed passages (SendMessage to the same agent) until it returns 通过.
5. Tell the user: verdict, changes made because of the audit, and items the auditor flagged for human verification.

## Auditor prompt (fill in the brackets)

```text
你是一名独立的第三方审计员，没有参与这份行程的制作。请只读核查，不要修改任何文件。

## 材料
- 来源报价（唯一事实来源）：[报价 .doc/.pdf 路径]。已提取文本：[_quote_extract.txt 路径]（UTF-8）。
- 被审计的客户行程：[行程 .docx 路径]。用 python-docx 读正文，并用 zipfile 读整个包内所有 XML
  （页眉、页脚、docProps/core.xml、app.xml、批注、脚注等）。Python 加 -X utf8。

## 客户的硬性要求（逐条判定）
[把用户本次的要求原文贴在这里，例如：禁止修改行程内容；禁止承诺报价中没有的服务；
 禁止把供应商信息和价格透露到行程中；只需要每天的行程安排以及景点介绍。]
并按以下口径执行：
1. 每天的路线、景点及其附注、顺序、车程、航班/车次与时刻（含"暂定"）、餐名、酒店名与等级必须与报价一致，不漏不改。
2. 正文、概述、景点介绍不得新增或暗示报价未列的景点、子景点、活动、餐食、门票、交通、导游、赠送、
   "纯玩/无购物/含XX"，也不得暗示报价未写的时段（夜游）或必然效果（保证花期）。客观描述景点本身可以，
   但措辞不能让客人以为这是包含的服务。
3. 整个 docx 包（含元数据）不得出现供应商名称/地址/电话/微信/QQ/联系人，以及团费、单房差、餐标金额、
   票价、小费、自费、购物推荐、16免1 等（除非用户明确要求含报价）。

## 输出（中文）
- 结论：通过 / 不通过
- 发现清单：严重度（违规 / 风险 / 建议）、位置（Day N、段落）、原文引文、报价依据或缺失依据、理由、建议改法
- 已检查且无问题的项目（说明覆盖面）
- 需要人工核实的事实（景点年代、数据、不熟悉的景点等）
请严格、具体，不要泛泛而谈。
```

## What counts as pass

- Zero 违规.
- Every 风险 fixed, or explicitly reported to the user with the reason it stays (e.g. 「景中店」 kept because it is verbatim quote text).
- 建议 are optional; mention the ones you didn't take.
