# Pre-delivery fact-check checklist

Run this **before** saving the final DOCX. Quote text (or `_quote_extract.txt`) is the sole source.

## Pricing & policy

**Default = no pricing block in the client DOCX.**

```
[ ] Default path: 正文无团费/单房差/FOC/5850 等价格段、「报价与说明」标题可无
[ ] Only if user explicitly asked to include 报价:
[ ]   团费 / 人 数字与报价一致（含币种 RMB）
[ ]   单间差 一致
[ ]   成团人数 / 16免1 / 领队单间规则 一致
[ ]   儿童占床 / 不占床比例 一致
[ ]   小费（如 RMB30/人天）一致
[ ]   报价有效期 + 避开节假日说明 一致
[ ] 未写报价没有的「含机票」「含保险」「含于团费」等
```

## Days & logistics

```
[ ] 天数与报价标题一致（如「6天」）；未擅自加「5晚」除非报价写了晚数
[ ] 每日路线、城市顺序与报价一致
[ ] 车程 / 车次 / 时段标注与来源一致
[ ] 每日酒店名 + 星级（五星/准五星/五钻）与报价一致
[ ] 返程日无酒店时不写宿：
[ ] Day 条文案（含可选 06Dec 日期前缀）与指定修改版一致
```

## Meals

```
[ ] 每日餐名与报价一致（含「自理」）
[ ] 餐标数字一致（如 烤鸭60、晚宴150、火锅100）— 或按指定草稿省略数字
[ ] 餐次代码（三餐 / 早午 / 仅早 / 仅晚）与报价膳食栏一致
[ ] 膳食汇总段（若有）与分日内容不矛盾
```

## Sights & services

```
[ ] 景点列表 ⊆ 报价（无新增景点）
[ ] 含电瓶车 / 含门票 等附注仅在报价出现时写入
[ ] 雨天取消、周一闭馆、季节替换、升船机停运等限制条件保留
[ ] 外国游客景点介绍：是「地点说明」，不是整句「游览/参观+景点名」
[ ] 景点介绍字数符合用户要求（常见 40–50 汉字不含标点）；用 CJK 计数校验
[ ] 介绍未新增报价没有的服务 / 门票 / 表演
```

## 赠送 vs 行程内

```
[ ] 赠送栏只含报价「赠送」项或正文明确「赠送…」
[ ] 行程内 DIY / 表演 / 品茗 等未误标为「赠送」
[ ] 矿泉水、温泉门票、米酒、无米粿 等与报价分类一致
```

## Shopping & inclusions language

```
[ ] 购物句与报价一致（如「全程不进店」）— 报价未写则不发明
[ ] 未发明「无购物店外费用」「无必加」等（除非来源有）
[ ] 独家安排 / 亮点句均可在报价或用户指定前部文案中找到依据
```

## Contacts & hotels summary

```
[ ] 联系人、电话、传真、Skype、Email、QQ 与报价一致；无则不编造
[ ] 精选酒店总述（若有）与分日酒店一致
```

## Style fidelity (when user asked to 复刻 / 学习修改版)

```
[ ] 模板壳 = 用户指定 DOCX 或已知 Voyage 样例（页眉 logo 在）
[ ] Day 条底色与样例一致（常见 #2E8B7A）
[ ] Day 条字体 Arial Unicode MS、白字
[ ] 标题 #1A6B5E、副标 #555555、亮点/美食常见 #008080
[ ] 【景点名】与 宿： 为深绿加粗
[ ] 未混用错误 Day 底色（如样例是 2E8B7A 却写成 008080 底）
```

## Readability (not facts, still block delivery)

```
[ ] 无多余空格（· / 两侧、Day 条超长空格填充）
[ ] 标点统一（→ ｜ 、）
[ ] 文件可在 Word / WPS 打开
[ ] 输出在源文件同目录（+ 可选 ASCII 文件名副本）
```

## Quick side-by-side method

1. Save extract: `python scripts/extract_doc.py "报价.doc" > _quote_extract.txt`  
   (or unzip/python-docx dump for `.docx`)
2. Open DOCX and quote extract side by side
3. Check each day hotel + meal line first (highest error rate)
4. Check pricing block second — or confirm it is absent if forbidden
5. Spot-check 2–3 sight intros: place fact present, CJK length in range
6. Grep DOCX text for words that must not appear unless in quote:  
   `含机票` `含保险` `无必加` `含于团费` `适合拍照` `值得一游`
