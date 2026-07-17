# Pre-delivery fact-check checklist

Run this **before** saving the final DOCX. Quote text (or `_quote_extract.txt`) is the sole source.

## Pricing & policy

```
[ ] 团费 / 人 数字与报价一致（含币种 RMB）
[ ] 单间差 一致
[ ] 成团人数 / 16免1 / 领队单间规则 一致
[ ] 儿童占床 / 不占床比例 一致
[ ] 小费（如 RMB30/人天）一致
[ ] 报价有效期 + 避开节假日说明 一致
[ ] 未写报价没有的「含机票」「含保险」「含于团费」等
```

## Days & logistics

```
[ ] 天数与报价标题一致（如「6天」）；未擅自加「5晚」除非报价写了晚数
[ ] 每日路线、城市顺序与报价一致
[ ] 车程标注（如 3H、1.5H+3.5H）与报价括号一致
[ ] 每日酒店名 + 星级（五星/准五星）与报价一致
[ ] 返程日无酒店时不写宿：
```

## Meals

```
[ ] 每日餐名与报价一致（含「自理」）
[ ] 餐标数字一致（如 烤鸭60、晚宴150、火锅100）
[ ] 餐次代码（三餐 / 早午 / 仅早 / 仅晚）与报价膳食栏一致
[ ] 膳食汇总段与分日内容不矛盾
```

## Sights & services

```
[ ] 景点列表 ⊆ 报价（无新增景点）
[ ] 含电瓶车 / 含门票 等附注仅在报价出现时写入
[ ] 雨天取消、周一闭馆等限制条件保留
[ ] 可选「外国游客扩写」仅重述报价事实，不新增服务
```

## 赠送 vs 行程内

```
[ ] 赠送栏只含报价「赠送」项或正文明确「赠送…」
[ ] 行程内 DIY / 表演 / 品茗 等未误标为「赠送」
[ ] 矿泉水、温泉门票、米酒、无米粿 等与报价分类一致
```

## Shopping & inclusions language

```
[ ] 购物句与报价一致（如「全程不进店」）
[ ] 未发明「无购物店外费用」「无必加」等
[ ] 独家安排 / 亮点句均可在报价中找到依据
```

## Contacts & hotels summary

```
[ ] 联系人、电话、传真、Skype、Email、QQ 与报价一致
[ ] 精选酒店总述与分日酒店一致
```

## Readability (not facts, still block delivery)

```
[ ] 无多余空格（· / 两侧、Day 条超长空格填充）
[ ] 标点统一（→ ｜ 、）
[ ] 文件可在 Word / WPS 打开
[ ] 输出在源 .doc 同目录（+ 可选 ASCII 文件名副本）
```

## Quick side-by-side method

1. Save extract: `python scripts/extract_doc.py "报价.doc" > _quote_extract.txt`
2. Open DOCX and quote extract side by side
3. Check each day hotel + meal line first (highest error rate)
4. Check pricing block second
5. Grep DOCX text for words that must not appear unless in quote:  
   `含机票` `含保险` `无必加` `含于团费` `适合拍照` `值得一游`
