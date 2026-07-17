#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DEPRECATED — pure OOXML / non-Voyage path. Do not use for new itineraries.

Canonical style: generate.py (template) or generate_chaoshan6_voyage_style.py
(worked example). See ../SKILL.md.

Historical: 跟着阿嬷游潮汕 pure-OOXML attempt; Word-open issues led to
template-shell Voyage approach instead.
"""
from __future__ import annotations

import os
import zipfile
from xml.sax.saxutils import escape
from datetime import datetime, timezone

# ---------- 输出路径 ----------
OUTPUT_DIR = r'C:\Users\DELL\Downloads'
OUTPUT_CN = os.path.join(OUTPUT_DIR, '跟着阿嬷游潮汕_厦门进出6天行程.docx')
OUTPUT_EN = os.path.join(OUTPUT_DIR, 'Chaoshan_Xiamen_6D_Itinerary.docx')

# ---------- 颜色 ----------
DEEP_BLUE = '1A3C6E'
DARK_GREEN = '2E6B4F'
WARM_GOLD = 'B8862C'
DARK_GRAY = '333333'
MEDIUM_GRAY = '666666'
LIGHT_GRAY = '999999'
WHITE = 'FFFFFF'
BG_LIGHT = 'EDF2F9'
BG_WARM = 'FEF6E9'
BG_GREEN = 'E8F5ED'
BG_RED = 'FFF1F0'
BG_GOLD = 'FBF6EB'

FONT = 'Microsoft YaHei'


def cn_count(s: str) -> int:
    return sum(1 for c in s if '\u4e00' <= c <= '\u9fff')


def t(s: str) -> str:
    return escape(s, {'"': '&quot;', "'": '&apos;'})


# ---------- 行程数据（严格依据报价） ----------
DAYS = [
    {
        'num': 1,
        'route': '新加坡 / 厦门',
        'sights': [
            {
                'time': '全天',
                'name': '抵达厦门 · 环岛路黄金海岸 / 演武大桥 / 世贸双子塔 / 沙坡尾',
                'desc': '本日自新加坡乘机前往福建滨海城市厦门。抵达后按行程参观环岛路黄金海岸，登上演武大桥海上观景平台远眺海面与岸线，并可远眺厦门地标建筑世贸双子塔，随后打卡网红景点沙坡尾。本日报价餐食仅列晚餐：龙虾鲍鱼红酒晚宴，不含午间用餐安排。',
                'meal': '晚餐：龙虾鲍鱼红酒晚宴\n餐标 RMB150',
            },
        ],
        'meals_summary': '晚餐（龙虾鲍鱼红酒晚宴，餐标 RMB150）；午不列',
        'hotel': '丽华酒店（五星）',
    },
    {
        'num': 2,
        'route': '厦门 / 永定（车程约 3H）',
        'sights': [
            {
                'time': '上午',
                'name': '世界文化遗产 · 海上花园鼓浪屿（龙头路 / 万国建筑 / 最美转角）',
                'desc': '早餐后前往被称作海上花园的鼓浪屿，该岛列入世界文化遗产。行程安排龙头路商业街与万国建筑博览，并途经当地所称最美转角。游客可沿街巷步行，观察岛上多样建筑立面与商业街区风貌，了解这座与厦门本岛隔海相望、以近代建筑群著称的历史岛屿。',
                'meal': '早餐：酒店',
            },
            {
                'time': '下午',
                'name': '初溪土楼群（含电瓶车）· 集庆楼 / 善庆楼 / 绳庆楼',
                'desc': '自厦门前往永定，车程约三小时。游览客家土楼中被称作最美丽、最原生态的初溪土楼群，行程含电瓶车接驳。重点参观世界文化遗产集庆楼、善庆楼、绳庆楼，了解客家人以大型围合土楼聚族而居的传统建筑形式、村落格局与土楼群整体风貌，完成本日土楼参观。',
                'meal': '午餐：北京烤鸭\n餐标 RMB60',
            },
            {
                'time': '晚间',
                'name': '天子温泉（赠温泉门票 / SPA，请自备泳衣）· 客家料理',
                'desc': '当晚入住当地温泉酒店天子温泉（五星），为行程中提升的超五星温泉酒店一晚。报价赠送温泉门票，可尽情享受泡汤乐趣并体验温泉SPA，请自备泳衣。晚餐品尝正宗客家料理，并赠送品尝客家米酒。本日含早午晚三餐，午餐为北京烤鸭，晚餐为客家料理。',
                'meal': '晚餐：客家料理\n餐标 RMB60\n赠客家米酒',
            },
        ],
        'meals_summary': '三餐：早 + 北京烤鸭（餐标60）+ 客家料理（餐标60，赠客家米酒）',
        'hotel': '天子温泉（五星）· 赠温泉门票 / SPA',
    },
    {
        'num': 3,
        'route': '永定 / 揭阳 / 潮州（车程约 3H+1H）',
        'sights': [
            {
                'time': '上午',
                'name': '《给阿嬷的情书》取景地 · 揭阳西淇村（赠南枝同款无米粿）',
                'desc': '早餐后车程前往揭阳，打卡电影《给阿嬷的情书》取景地揭阳西淇村。行程可感受古村田园诗意与乡愁氛围，了解一纸侨批、半世守望的故事主题，以及《石板桥》所承载的牵挂与等待；并赠送品尝南枝同款无米粿，体验与影片相关的潮汕侨乡饮食安排。',
                'meal': '早餐：酒店',
            },
            {
                'time': '中午',
                'name': '揭阳古城西马路历史文化名街（唐人街 / 暹罗客栈取景地）',
                'desc': '随后打卡揭阳古城西马路历史文化名街，此处亦为电影中唐人街、暹罗客栈取景地。街区内南洋风情骑楼错落有致，青石板路保留旧城岁月最深记忆。游客可步行浏览骑楼商铺与街巷布局，感受揭阳侨乡历史街区的建筑立面、街巷尺度与市井氛围。',
                'meal': '午餐：揭阳风味\n餐标 RMB60',
            },
            {
                'time': '下午',
                'name': '广东十大最美古村 · 千年龙湖古寨（还原电影约 90% 泰国街景）',
                'desc': '后往历史古都潮州，游览列入广东十大最美古村的千年古村龙湖古寨。报价称其街景可还原电影中约百分之九十的泰国街景元素。游客可沿古寨巷弄参观传统民居与街巷格局，了解潮汕古村落的空间肌理、生活痕迹与影片取景之间的对应关系。',
                'meal': '晚餐：潮式十八碟\n餐标 RMB80\n含手工DIY红桃粿',
            },
        ],
        'meals_summary': '三餐：早 + 揭阳风味（餐标60）+ 潮式十八碟（餐标80，手工DIY红桃粿）',
        'hotel': '愉羽酒店（准五星）',
    },
    {
        'num': 4,
        'route': '潮州 / 澄海 / 汕头 / 潮州（车程约 1H+1H+1H）',
        'sights': [
            {
                'time': '上午',
                'name': '岭南第一侨宅 · 陈慈黉故居（潮汕英歌舞）',
                'desc': '早餐后前往澄海，打卡《给阿嬷的情书》取景地、被称作岭南第一侨宅的陈慈黉故居。行程中可欣赏非遗文化潮汕英歌舞。游客可参观侨宅院落与建筑布局，了解潮汕华侨家族住宅的历史背景，以及影片取景地点与潮汕侨乡文化、华侨生活记忆之间的直接关联。',
                'meal': '早餐：酒店',
            },
            {
                'time': '中午',
                'name': '汕头小公园老城骑楼 · 暹南电影院 / 裕丰银信局 / 侨批文物馆',
                'desc': '前往美食之乡汕头，参观小公园老城骑楼建筑群。行程含海平路七十七号还原片中暹南电影院场景，以及万安街三号裕丰银信局所还原的寄侨批名场面，并参观侨批文物馆以了解侨批历史。请注意侨批文物馆逢周一闭馆，当天参观安排须留意闭馆说明。',
                'meal': '午餐：汕头全鹅宴\n餐标 RMB60',
            },
            {
                'time': '下午/晚',
                'name': '潮州古城（含电瓶车）· 西湖 / 广济门 / 牌坊街 / 茶馆潮剧 / 灯光秀',
                'desc': '后返潮州，游览千年潮州古城，行程含电瓶车。安排打卡西湖湖心亭与泰佛殿，参观广济门古城楼与牌坊街，于百年茶馆品茗听潮剧，漫步滨江长廊并远眺湘子桥；晚上欣赏古城灯光秀。若遇雨天，灯光秀安排取消；其余古城参观内容以当日报价行程为准。',
                'meal': '晚餐：自理\n潮汕小吃',
            },
        ],
        'meals_summary': '早午（汕头全鹅宴，餐标60）；晚餐自理潮汕小吃',
        'hotel': '愉羽酒店（准五星）',
    },
    {
        'num': 5,
        'route': '潮州 / 南澳 / 厦门（车程约 1.5H+3.5H）',
        'sights': [
            {
                'time': '全天',
                'name': '南澳岛 · 南澳大桥 / 前江湾 / 启航广场 / 北回归线标志塔 / 青澳湾',
                'desc': '早餐后前往被称作广东最美岛屿的南澳岛，途经南澳大桥，车程与后续返厦合计见当日交通。车游新南澳外滩前江湾海滨路，打卡网红景点启航广场，参观自然之门北回归线标志塔，以及被称作东方夏威夷的青澳湾。岛上行程结束后返回厦门入住。',
                'meal': '午餐：南澳小海鲜\n餐标 RMB60\n晚餐：自助海鲜烤肉火锅\n餐标 RMB100',
            },
        ],
        'meals_summary': '三餐：早 + 南澳小海鲜（餐标60）+ 自助海鲜烤肉火锅（餐标100）',
        'hotel': '丽华酒店（五星）',
    },
    {
        'num': 6,
        'route': '厦门 / 新加坡',
        'sights': [
            {
                'time': '上午',
                'name': '自由活动至集合 · 送机返程',
                'desc': '早餐后在厦门自由活动，直至集合时间。随后前往机场搭乘航班返回新加坡温暖家园。本日餐食仅含早餐，不含午晚餐及任何景点参观安排；请按领队通知准时集合，并预留前往机场、办理值机与登机手续所需时间，确保顺利登机并完成本次六日行程返程。',
                'meal': '早餐：酒店',
            },
        ],
        'meals_summary': '仅早餐；午晚餐不含',
        'hotel': '本日行程结束（无酒店）',
    },
]

HIGHLIGHTS = [
    '独家：跟着阿嬷游潮汕，走进《给阿嬷的情书》故里，领略潮汕华侨文化',
    '漫步千年潮州古城，百年茶馆品茗听潮剧',
    '欣赏非遗文化潮汕英歌舞；手工DIY红桃粿',
    '广东十大最美古村 · 龙湖古寨',
    '广东最美岛屿、东方夏威夷 · 南澳岛',
    '海上花园鼓浪屿、永定土楼 · 两大世界文化遗产',
    '精选酒店：当地五星为主，提升 1 晚超五星温泉酒店（温泉 SPA）',
    '美食：龙虾鲍鱼红酒晚宴、北京烤鸭、客家风味、潮式十八碟、汕头全鹅宴、南澳海鲜、自助海鲜烤肉火锅等',
]


# ---------- OOXML helpers ----------

def run_xml(text: str, size_half_points: int = 20, bold: bool = False,
            color: str = DARK_GRAY, font: str = FONT) -> str:
    b = '<w:b/>' if bold else ''
    return (
        f'<w:r>'
        f'<w:rPr>'
        f'<w:rFonts w:ascii="{t(font)}" w:hAnsi="{t(font)}" w:eastAsia="{t(font)}"/>'
        f'{b}'
        f'<w:sz w:val="{size_half_points}"/>'
        f'<w:szCs w:val="{size_half_points}"/>'
        f'<w:color w:val="{color}"/>'
        f'</w:rPr>'
        f'<w:t xml:space="preserve">{t(text)}</w:t>'
        f'</w:r>'
    )


def para_xml(runs_xml: str, align: str = 'left',
             before: int = 0, after: int = 60, line: int = 276) -> str:
    jc = f'<w:jc w:val="{align}"/>' if align != 'left' else ''
    return (
        f'<w:p>'
        f'<w:pPr>'
        f'<w:spacing w:before="{before}" w:after="{after}" w:line="{line}" w:lineRule="auto"/>'
        f'{jc}'
        f'</w:pPr>'
        f'{runs_xml}'
        f'</w:p>'
    )


def empty_para() -> str:
    return '<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr></w:p>'


def shd(fill: str) -> str:
    return f'<w:shd w:val="clear" w:color="auto" w:fill="{fill}"/>'


def cell_xml(paragraphs: list[str], width_dxa: int, fill: str | None = None,
             v_align: str | None = None) -> str:
    props = [f'<w:tcW w:w="{width_dxa}" w:type="dxa"/>']
    if fill:
        props.append(shd(fill))
    if v_align:
        props.append(f'<w:vAlign w:val="{v_align}"/>')
    body = ''.join(paragraphs) if paragraphs else empty_para()
    return f'<w:tc><w:tcPr>{"".join(props)}</w:tcPr>{body}</w:tc>'


def row_xml(cells: list[str]) -> str:
    return f'<w:tr>{"".join(cells)}</w:tr>'


def table_xml(rows: list[str], col_widths: list[int]) -> str:
    grid = ''.join(f'<w:gridCol w:w="{w}"/>' for w in col_widths)
    return (
        f'<w:tbl>'
        f'<w:tblPr>'
        f'<w:tblW w:w="{sum(col_widths)}" w:type="dxa"/>'
        f'<w:tblBorders>'
        f'<w:top w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>'
        f'<w:left w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>'
        f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>'
        f'<w:right w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>'
        f'<w:insideH w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>'
        f'<w:insideV w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>'
        f'</w:tblBorders>'
        f'<w:tblCellMar>'
        f'<w:top w:w="40" w:type="dxa"/>'
        f'<w:left w:w="80" w:type="dxa"/>'
        f'<w:bottom w:w="40" w:type="dxa"/>'
        f'<w:right w:w="80" w:type="dxa"/>'
        f'</w:tblCellMar>'
        f'</w:tblPr>'
        f'<w:tblGrid>{grid}</w:tblGrid>'
        f'{"".join(rows)}'
        f'</w:tbl>'
    )


def multi_line_paras(text: str, size: int = 18, color: str = MEDIUM_GRAY,
                     align: str = 'center', bold: bool = False) -> list[str]:
    lines = text.split('\n') if text else ['']
    out = []
    for i, line in enumerate(lines):
        out.append(para_xml(
            run_xml(line, size, bold=bold, color=color),
            align=align, before=40 if i == 0 else 20, after=20, line=260
        ))
    return out


def build_document_body() -> str:
    parts: list[str] = []

    # Title page
    parts.append(empty_para())
    parts.append(empty_para())
    parts.append(para_xml(run_xml('*  *  *  *  *', 28, color=WARM_GOLD), align='center', after=120))
    parts.append(para_xml(run_xml('跟着阿嬷游潮汕', 56, bold=True, color=DEEP_BLUE), align='center', after=120))
    parts.append(para_xml(
        run_xml('舌尖上的福建/潮汕美食团 · 厦门进出6天', 26, color=MEDIUM_GRAY),
        align='center', after=80))
    parts.append(para_xml(
        run_xml('厦门 · 鼓浪屿 · 永定土楼 · 揭阳 · 潮州 · 汕头 · 澄海 · 南澳岛', 22, color=WARM_GOLD),
        align='center', after=120))
    parts.append(para_xml(run_xml('*  *  *  *  *', 28, color=WARM_GOLD), align='center', after=200))
    parts.append(para_xml(run_xml('行程亮点（据报价）', 28, bold=True, color=DEEP_BLUE), after=120))

    for h in HIGHLIGHTS:
        parts.append(para_xml(run_xml('[*]  ' + h, 21, color=DARK_GRAY), after=40, line=280))

    parts.append(para_xml(
        run_xml('说明：本文档景点、餐食、酒店、价格与服务均严格依据华宇报价原文整理，未增加报价外项目。',
                18, color=LIGHT_GRAY),
        after=120))

    # Page break
    parts.append(
        '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
    )

    # Days
    col_w = [1247, 6520, 1871]  # ~2.2cm, 11.5cm, 3.3cm in dxa (1cm≈567)
    total_w = sum(col_w)

    for day in DAYS:
        # Day header table (1 col)
        badge = f'D-{day["num"]:02d}'
        header_runs = (
            run_xml(f'  {badge}  ', 24, bold=True, color=WHITE) +
            run_xml('  ' + day['route'], 24, bold=True, color=WHITE)
        )
        header_row = row_xml([
            cell_xml(
                [para_xml(header_runs, align='left', before=80, after=80)],
                total_w, fill=DEEP_BLUE, v_align='center'
            )
        ])
        parts.append(table_xml([header_row], [total_w]))
        parts.append(empty_para())

        # Detail table
        header_cells = [
            cell_xml([para_xml(run_xml('时段', 18, bold=True, color=WHITE),
                               align='center', before=60, after=60)],
                     col_w[0], fill=DEEP_BLUE, v_align='center'),
            cell_xml([para_xml(run_xml('行程亮点', 18, bold=True, color=WHITE),
                               align='center', before=60, after=60)],
                     col_w[1], fill=DEEP_BLUE, v_align='center'),
            cell_xml([para_xml(run_xml('餐食/备注', 18, bold=True, color=WHITE),
                               align='center', before=60, after=60)],
                     col_w[2], fill=DEEP_BLUE, v_align='center'),
        ]
        rows = [row_xml(header_cells)]

        for sight in day['sights']:
            time_cell = cell_xml(
                [para_xml(run_xml(sight['time'], 18, bold=True, color=DEEP_BLUE),
                          align='center', before=60, after=60)],
                col_w[0], fill=BG_LIGHT, v_align='center'
            )
            name_p = para_xml(run_xml(sight['name'], 22, bold=True, color=DARK_GREEN),
                              before=40, after=40)
            desc_p = para_xml(run_xml(sight['desc'], 18, color=MEDIUM_GRAY),
                              before=20, after=60, line=280)
            meal_paras = multi_line_paras(sight.get('meal') or '', size=17,
                                          color=MEDIUM_GRAY, align='center')
            sight_cell = cell_xml([name_p, desc_p], col_w[1])
            meal_cell = cell_xml(meal_paras, col_w[2], fill=BG_WARM, v_align='center')
            rows.append(row_xml([time_cell, sight_cell, meal_cell]))

        # meals summary
        rows.append(row_xml([
            cell_xml(
                [para_xml(run_xml('[餐食]', 18, bold=True, color=DARK_GREEN),
                          align='center', before=40, after=40)],
                col_w[0], fill=BG_GREEN, v_align='center'
            ),
            cell_xml(
                [para_xml(run_xml(day['meals_summary'], 18, color=DARK_GRAY),
                          before=40, after=40)],
                col_w[1], fill=BG_GREEN
            ),
            cell_xml([empty_para()], col_w[2], fill=BG_GREEN),
        ]))

        # hotel
        rows.append(row_xml([
            cell_xml(
                [para_xml(run_xml('[住宿]', 18, bold=True, color='CC3333'),
                          align='center', before=40, after=40)],
                col_w[0], fill=BG_RED, v_align='center'
            ),
            cell_xml(
                [para_xml(run_xml(day['hotel'], 18, color=DARK_GRAY),
                          before=40, after=40)],
                col_w[1], fill=BG_RED
            ),
            cell_xml([empty_para()], col_w[2], fill=BG_RED),
        ]))

        parts.append(table_xml(rows, col_w))
        parts.append(empty_para())

    # Pricing section
    parts.append(para_xml(run_xml('报价与说明', 32, bold=True, color=DEEP_BLUE),
                          before=120, after=160))

    price_lines = [
        ('成人报价', 'RMB 2480 / 人'),
        ('单间差', 'RMB 800'),
        ('成团人数', '16 人以上报价；16 免 1'),
        ('领队单间', '领队若用单间需补单间差'),
        ('儿童', '占床与大人同价；不占床按 60%'),
        ('小费', 'RMB 30 / 人天'),
        ('赠送', '每人每天矿泉水两瓶'),
        ('购物', '全程不进店'),
        ('有效期', '2026年6月1日 至 2027年12月31日'),
        ('避开档期', '需避开 1–5/5「五一节」、1–7/10「国庆节」、春节期间'),
    ]
    price_paras = []
    for label, val in price_lines:
        price_paras.append(para_xml(
            run_xml(label + '：', 22, bold=True, color=DEEP_BLUE) +
            run_xml(val, 22, color=DARK_GRAY),
            before=40, after=40
        ))
    price_row = row_xml([
        cell_xml(price_paras, total_w, fill=BG_GOLD)
    ])
    parts.append(table_xml([price_row], [total_w]))
    parts.append(empty_para())

    parts.append(para_xml(run_xml('酒店安排（据报价）', 26, bold=True, color=DEEP_BLUE), after=80))
    for h in [
        'D1 / D5  丽华酒店（五星）',
        'D2       天子温泉（五星；赠温泉门票，享受温泉 SPA，请自备泳衣）',
        'D3 / D4  愉羽酒店（准五星）',
        '说明：报价写明「全程当地 5* 酒店，提升 1 晚超五星温泉酒店（享受温泉 SPA）」；每日酒店以行程表列名为准。',
    ]:
        parts.append(para_xml(run_xml('[*]  ' + h, 20, color=DARK_GRAY), after=40))

    parts.append(para_xml(run_xml('美食安排（据报价）', 26, bold=True, color=DEEP_BLUE),
                          before=120, after=80))
    for m in [
        'D1 晚餐：龙虾鲍鱼红酒晚宴（餐标 150）',
        'D2 午餐：北京烤鸭（餐标 60）；晚餐：客家料理（餐标 60，赠客家米酒）',
        'D3 午餐：揭阳风味（餐标 60）；晚餐：潮式十八碟（餐标 80，手工 DIY 红桃粿）',
        'D4 午餐：汕头全鹅宴（餐标 60）；晚餐：自理潮汕小吃',
        'D5 午餐：南澳小海鲜（餐标 60）；晚餐：自助海鲜烤肉火锅（餐标 100）',
        'D6：仅早餐',
        '亮点菜式亦列：客家风味、梅菜扣肉+盐酒鸡、揭阳小吃、南澳海鲜、潮汕小吃等（以当日餐食安排为准）',
        '说明：上列金额为报价原文中的餐标，含于团费 RMB2480/人，非另行加价（自理餐除外）。',
    ]:
        parts.append(para_xml(run_xml('[*]  ' + m, 20, color=DARK_GRAY), after=40))

    parts.append(para_xml(run_xml('联系方式（据报价）', 26, bold=True, color=DEEP_BLUE),
                          before=120, after=80))
    for c in [
        '联系人：姚文跃（13600926388）；李珏（13950127910）',
        '电话：0592-8121184 / 8121185　　传真：0592-8121195',
        'Skype：xiaoyao_1976　　Email：yao19760126@aliyun.com　　QQ：499756025',
    ]:
        parts.append(para_xml(run_xml(c, 20, color=DARK_GRAY), after=40))

    parts.append(para_xml(
        run_xml('本行程文档由华宇报价原文整理生成，仅供阅读参考；最终以签约确认的行程与报价单为准。',
                18, color=LIGHT_GRAY),
        before=120, after=40))

    # sectPr required at end of body
    sect = (
        '<w:sectPr>'
        '<w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1134" w:right="1134" w:bottom="1020" w:left="1134" '
        'w:header="720" w:footer="720" w:gutter="0"/>'
        '<w:cols w:space="720"/>'
        '<w:docGrid w:linePitch="360"/>'
        '</w:sectPr>'
    )
    return ''.join(parts) + sect


def document_xml() -> str:
    body = build_document_body()
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:body>{body}</w:body>'
        '</w:document>'
    )


def content_types_xml() -> str:
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>
  <Override PartName="/word/webSettings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
'''


def rels_xml() -> str:
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
'''


def document_rels_xml() -> str:
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings" Target="webSettings.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>
</Relationships>
'''


def styles_xml() -> str:
    # Minimal styles with Normal + TableGrid
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:eastAsia="{FONT}"/>
        <w:sz w:val="20"/>
        <w:szCs w:val="20"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:spacing w:after="0" w:line="276" w:lineRule="auto"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
    <w:rPr>
      <w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:eastAsia="{FONT}"/>
      <w:sz w:val="20"/>
      <w:szCs w:val="20"/>
    </w:rPr>
  </w:style>
  <w:style w:type="table" w:default="1" w:styleId="TableNormal">
    <w:name w:val="Normal Table"/>
    <w:tblPr>
      <w:tblInd w:w="0" w:type="dxa"/>
      <w:tblCellMar>
        <w:top w:w="0" w:type="dxa"/>
        <w:left w:w="108" w:type="dxa"/>
        <w:bottom w:w="0" w:type="dxa"/>
        <w:right w:w="108" w:type="dxa"/>
      </w:tblCellMar>
    </w:tblPr>
  </w:style>
  <w:style w:type="table" w:styleId="TableGrid">
    <w:name w:val="Table Grid"/>
    <w:basedOn w:val="TableNormal"/>
    <w:uiPriority w:val="59"/>
    <w:tblPr>
      <w:tblBorders>
        <w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
        <w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>
        <w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
        <w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
        <w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>
        <w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      </w:tblBorders>
    </w:tblPr>
  </w:style>
</w:styles>
'''


def settings_xml() -> str:
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:zoom w:percent="100"/>
  <w:defaultTabStop w:val="720"/>
  <w:characterSpacingControl w:val="doNotCompress"/>
  <w:compat>
    <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
  </w:compat>
</w:settings>
'''


def web_settings_xml() -> str:
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:optimizeForBrowser/>
  <w:allowPNG/>
</w:webSettings>
'''


def font_table_xml() -> str:
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:font w:name="{FONT}">
    <w:panose1 w:val="020B0604030504040204"/>
    <w:charset w:val="86"/>
    <w:family w:val="swiss"/>
    <w:pitch w:val="variable"/>
  </w:font>
  <w:font w:name="Arial">
    <w:panose1 w:val="020B0604020202020204"/>
    <w:charset w:val="00"/>
    <w:family w:val="swiss"/>
    <w:pitch w:val="variable"/>
  </w:font>
</w:fonts>
'''


def core_xml() -> str:
    now = datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:dcterms="http://purl.org/dc/terms/"
  xmlns:dcmitype="http://purl.org/dc/dcmitype/"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>跟着阿嬷游潮汕 · 厦门进出6天</dc:title>
  <dc:subject>华宇报价行程</dc:subject>
  <dc:creator>华宇报价整理</dc:creator>
  <cp:lastModifiedBy>itinerary-generator</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>
</cp:coreProperties>
'''


def app_xml() -> str:
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
  xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>OOXML Itinerary Generator</Application>
  <DocSecurity>0</DocSecurity>
  <ScaleCrop>false</ScaleCrop>
  <SharedDoc>false</SharedDoc>
  <HyperlinksChanged>false</HyperlinksChanged>
  <AppVersion>16.0000</AppVersion>
</Properties>
'''


def write_docx(path: str) -> None:
    parts = {
        '[Content_Types].xml': content_types_xml(),
        '_rels/.rels': rels_xml(),
        'word/document.xml': document_xml(),
        'word/_rels/document.xml.rels': document_rels_xml(),
        'word/styles.xml': styles_xml(),
        'word/settings.xml': settings_xml(),
        'word/webSettings.xml': web_settings_xml(),
        'word/fontTable.xml': font_table_xml(),
        'docProps/core.xml': core_xml(),
        'docProps/app.xml': app_xml(),
    }
    # Validate XML well-formedness
    from xml.etree import ElementTree as ET
    for name, xml in parts.items():
        try:
            ET.fromstring(xml)
        except ET.ParseError as e:
            raise RuntimeError(f'Invalid XML in {name}: {e}') from e

    # Write ZIP with stored compression for maximum compatibility
    if os.path.exists(path):
        os.remove(path)
    with zipfile.ZipFile(path, 'w', compression=zipfile.ZIP_DEFLATED) as zf:
        # Content_Types first (convention)
        for name, xml in parts.items():
            # Use UTF-8 without BOM
            data = xml.encode('utf-8')
            info = zipfile.ZipInfo(name)
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = 0o644 << 16
            zf.writestr(info, data)

    # Verify
    with zipfile.ZipFile(path, 'r') as zf:
        bad = zf.testzip()
        if bad:
            raise RuntimeError(f'Corrupt zip entry: {bad}')
        # Ensure Word can parse document
        ET.fromstring(zf.read('word/document.xml'))


def verify_descriptions() -> None:
    print('--- Character count check (target 100-150 CJK) ---')
    ok = True
    for day in DAYS:
        for s in day['sights']:
            n = cn_count(s['desc'])
            flag = 'OK' if 100 <= n <= 150 else 'WARN'
            if flag != 'OK':
                ok = False
            print(f'{flag} D{day["num"]} {s["name"][:24]}... -> {n}')
    if not ok:
        print('WARNING: some descriptions outside 100-150 range')


def main() -> None:
    verify_descriptions()
    for path in (OUTPUT_CN, OUTPUT_EN):
        write_docx(path)
        print(f'OK - Saved: {path} ({os.path.getsize(path)} bytes)')


if __name__ == '__main__':
    main()
