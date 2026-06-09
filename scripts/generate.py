#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Generate a professional travel itinerary DOCX for: 湖北武汉+庐山8日
Source: Singapore/Malaysia tourists, Buddhist vegetarian tour
"""

import os
import re
import zipfile
import shutil
from docx import Document
from docx.shared import Inches, Cm, Pt, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml

# ============================================================
# Colors
# ============================================================
DEEP_BLUE   = RGBColor(0x1A, 0x3C, 0x6E)
DARK_GREEN  = RGBColor(0x2E, 0x6B, 0x4F)
WARM_GOLD   = RGBColor(0xB8, 0x86, 0x2C)
DARK_GRAY   = RGBColor(0x33, 0x33, 0x33)
MEDIUM_GRAY = RGBColor(0x66, 0x66, 0x66)
LIGHT_GRAY  = RGBColor(0x99, 0x99, 0x99)
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)

# Hex strings for shading
DEEP_BLUE_HEX  = '1A3C6E'
DARK_GREEN_HEX = '2E6B4F'
WARM_GOLD_HEX  = 'B8862C'
BG_LIGHT_HEX   = 'EDF2F9'
BG_GREEN_HEX   = 'E8F5ED'
BG_RED_HEX     = 'FFF1F0'
BG_YELLOW_HEX  = 'FFF8E1'

# ============================================================
# Helpers
# ============================================================
NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

def set_cell_va(cell, align='center'):
    """Set vertical alignment using raw string, NOT WD_ALIGN_VERTICAL enum."""
    tcPr = cell._tc.get_or_add_tcPr()
    for existing in tcPr.findall('{%s}vAlign' % NS):
        tcPr.remove(existing)
    tcPr.append(parse_xml('<w:vAlign {} w:val="{}"/>'.format(nsdecls("w"), align)))

def set_cell_shading(cell, hex_color):
    """Set cell background shading. Removes existing shd first."""
    tcPr = cell._tc.get_or_add_tcPr()
    for existing in tcPr.findall('{%s}shd' % NS):
        tcPr.remove(existing)
    shading_elm = parse_xml(
        '<w:shd {} w:fill="{}" w:val="clear"/>'.format(nsdecls("w"), hex_color)
    )
    tcPr.append(shading_elm)

def set_cell_width(cell, width_cm):
    """Set exact cell width."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcW = tcPr.find('{%s}tcW' % NS)
    if tcW is None:
        tcW = parse_xml('<w:tcW {} w:w="{}" w:type="dxa"/>'.format(
            nsdecls("w"), str(int(width_cm * 567))
        ))
        tcPr.insert(0, tcW)
    else:
        tcW.set('{%s}w' % NS, str(int(width_cm * 567)))
        tcW.set('{%s}type' % NS, 'dxa')

def set_table_borders(table):
    """Set clean borders on a table."""
    tbl = table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else parse_xml(
        '<w:tblPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'
    )
    borders = parse_xml(
        '<w:tblBorders %s>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="999999"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="999999"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="999999"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="999999"/>'
        '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="999999"/>'
        '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="999999"/>'
        '</w:tblBorders>' % nsdecls("w")
    )
    # Remove existing borders
    for existing in tblPr.findall('{%s}tblBorders' % NS):
        tblPr.remove(existing)
    tblPr.append(borders)

def add_formatted_run(paragraph, text, font_name='Arial', font_size=Pt(10),
                      bold=None, color=None, italic=False):
    """Add a run with proper font settings for Chinese text."""
    run = paragraph.add_run(text)
    run.font.name = font_name
    run.font.size = font_size
    # Set east-Asian font
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find('{%s}rFonts' % NS)
    if rFonts is None:
        rFonts = parse_xml('<w:rFonts %s w:eastAsia="%s"/>' % (nsdecls("w"), font_name))
        rPr.insert(0, rFonts)
    else:
        rFonts.set('{%s}eastAsia' % NS, font_name)
        rFonts.set('{%s}ascii' % NS, font_name)
        rFonts.set('{%s}hAnsi' % NS, font_name)
    if bold:
        run.font.bold = True
    if color:
        run.font.color.rgb = color
    if italic:
        run.font.italic = True
    return run

def add_heading_styled(doc, text, level=1):
    """Add a heading with DEEP_BLUE color."""
    heading = doc.add_heading(text, level=level)
    for run in heading.runs:
        run.font.color.rgb = DEEP_BLUE
    return heading

def cleanup_docx(output_path):
    """Remove parts that cause Word to reject the file."""
    import io
    temp_path = output_path + '.tmp.zip'
    os.rename(output_path, temp_path)
    
    with zipfile.ZipFile(temp_path, 'r') as zin:
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                # Skip problematic files
                if item.filename.startswith('customXml/'):
                    continue
                if 'stylesWithEffects' in item.filename:
                    continue
                if 'thumbnail' in item.filename.lower():
                    continue
                
                data = zin.read(item.filename)
                
                # Clean references in [Content_Types].xml
                if item.filename == '[Content_Types].xml':
                    content = data.decode('utf-8', errors='ignore')
                    # Remove customXml overrides
                    content = re.sub(r'<Override\s+PartName="/customXml/[^"]*"\s+ContentType="[^"]*"\s*/>', '', content)
                    # Remove stylesWithEffects override
                    content = re.sub(r'<Override\s+PartName="/word/stylesWithEffects\.xml"[^>]*/>', '', content)
                    data = content.encode('utf-8')
                
                # Clean references in _rels/.rels
                elif item.filename == '_rels/.rels':
                    content = data.decode('utf-8', errors='ignore')
                    # Remove thumbnail relationship
                    content = re.sub(r'<Relationship[^>]*Target="docProps/thumbnail[^"]*"[^>]*/>', '', content)
                    # Remove customXml relationships (if any)
                    content = re.sub(r'<Relationship[^>]*Target="customXml[^"]*"[^>]*/>', '', content)
                    data = content.encode('utf-8')
                
                # Clean references in word/_rels/document.xml.rels
                elif item.filename == 'word/_rels/document.xml.rels':
                    content = data.decode('utf-8', errors='ignore')
                    # Remove customXml references
                    content = re.sub(r'<Relationship[^>]*Target="[.][.]/customXml[^"]*"[^>]*/>', '', content)
                    # Remove stylesWithEffects references
                    content = re.sub(r'<Relationship[^>]*Target="stylesWithEffects[^"]*"[^>]*/>', '', content)
                    data = content.encode('utf-8')
                
                zout.writestr(item, data)
    
    os.remove(temp_path)

def add_page_break(doc):
    """Add a page break with zero visible spacing."""
    paragraph = doc.add_paragraph()
    set_no_spacing(paragraph)
    # Set font to 1pt to avoid any visible space
    run = paragraph.add_run()
    run.font.size = Pt(1)
    run._element.append(parse_xml('<w:br %s w:type="page"/>' % nsdecls("w")))

def cleanup_empty_rows(table):
    """Remove any fully-empty table rows (all cells blank/whitespace).
    Must be called before doc.save() for every day table."""
    NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    tbl = table._tbl
    rows_to_remove = []
    for row in tbl.findall('{%s}tr' % NS):
        cells = row.findall('{%s}tc' % NS)
        all_empty = True
        for tc in cells:
            text_parts = []
            for p in tc.findall('{%s}p' % NS):
                for t in p.iter('{%s}t' % NS):
                    if t.text:
                        text_parts.append(t.text)
            cell_text = ''.join(text_parts).strip()
            if cell_text:
                all_empty = False
                break
        if all_empty and len(cells) > 0:
            rows_to_remove.append(row)
    for row in rows_to_remove:
        tbl.remove(row)

def set_no_spacing(paragraph):
    """Remove spacing before/after."""
    pf = paragraph.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)

# ============================================================
# SIGHT DATA with descriptions (100-150 Chinese chars each)
# A helper to count Chinese chars
# ============================================================
def cn_count(s):
    return sum(1 for c in s if '\u4e00' <= c <= '\u9fff')

# Descriptions written following the 4-part structure:
# 1. What is this? 2. Why famous? 3. What can I see/do? 4. Practical info

SIGHT_DESCRIPTIONS = {
    # D2
    "东湖植物馆": (
        "位于武汉东湖风景区内的室内植物展示馆，占地约1.2万平方米，汇集热带雨林、"
        "干旱沙漠、水生湿地等七个生态区的珍稀植物。馆内建有18米高的室内瀑布和玻璃穹顶温室，"
        "游客可沿高空栈道穿行于树冠之间，俯瞰不同气候带的植被群落。"
        "每年3月下旬至4月上旬馆外樱花盛开时可与樱花园联游。"
    ),
    "武汉东湖樱花园": (
        "位于东湖磨山南麓，占地260亩，种植樱花树逾万株，品种涵盖染井吉野、关山、松月等三十余种。"
        "园内以日式园林风格布局，建有五重塔、虹桥和蜿蜒的曲水溪流，每年3月中旬至4月初为最佳观赏期，"
        "夜间设有灯光赏樱活动，开放时间延长至晚上9点，游客可漫步于樱花隧道下感受落英缤纷的景象。"
    ),
    "武汉光子号单轨悬挂列车": (
        "运行于武汉光谷生态大走廊的悬挂式单轨列车线路，全长约10.5公里，设站6座，于2023年开通运营。"
        "列车悬挂于轨道下方行驶，车厢底部距地面约10米，采用全透明玻璃地板和侧窗设计，"
        "乘客可在空中俯瞰光谷科技新城与沿线湖泊湿地景观，单程运行时间约25分钟。"
        "该线路为国内首条悬挂式单轨商业运营线，列车无人驾驶。"
    ),
    "李时珍纪念馆": (
        "位于蕲春县蕲州镇，为纪念明代医药学家李时珍（1518-1593年）而建，占地面积约5万平方米。"
        "李时珍历时27年编成《本草纲目》，收录药物1892种。馆内陈列其生平事迹、"
        "药物标本和历代《本草纲目》版本，园中种植数百种药用植物并附有功效说明标牌。"
        "馆内配备专业中文讲解员。"
    ),

    # D3
    "四祖寺": (
        "位于黄梅县西山，始建于唐武德七年（公元624年），为禅宗四祖道信大师的道场，"
        "是中国禅宗祖庭之一。现存建筑主要为清代重修，包括大雄宝殿、天王殿和毗卢殿，"
        "寺内保存唐代石雕佛像数尊和宋代碑刻。寺院四周古木参天，山门前的古银杏树龄逾千年。"
    ),
    "五祖寺": (
        "位于黄梅县东山，始建于唐永徽五年（公元654年），是禅宗五祖弘忍大师的弘法道场，"
        "六祖惠能在此得法，被视为中国禅宗发祥地，在中国佛教史上具有极其重要的地位。"
        "寺内建筑依山势层叠而上，真身殿供奉五祖肉身舍利塔，"
        "寺后山崖刻有历代摩崖题记30余处，山腰古松翠柏掩映，环境清幽肃穆。"
    ),
    "老祖寺": (
        "位于黄梅县北部山区，始建于南北朝时期，为禅宗初祖菩提达摩的纪念道场，距今已有1400余年历史。"
        "寺院规模较小但环境清幽，主殿供奉达摩祖师塑像，寺周环绕茂密竹林和茶园，"
        "山间溪流穿寺而过，寺内保留有明代古钟一口。寺院位置偏远，游客稀少，"
        "在此可感受远离尘嚣的禅修氛围，静听风声与溪水交织的天然梵音。"
    ),
    "问梅村": (
        "位于黄梅县五祖寺山脚下，是以禅文化和梅花为主题的文旅村落，占地约200亩。"
        "村内建筑采用唐宋风格，以石板路和水系相连，遍布梅树百余品种。"
        "村中设有禅修体验馆、茶室和素食餐厅，冬季梅花盛开时景色最佳。"
        "游客可在此夜宿民宿，体验晨钟暮鼓的山居生活。"
    ),

    # D4
    "庐山": (
        "位于江西省九江市，北临长江、东濒鄱阳湖，主峰汉阳峰海拔1474米，1996年列入联合国教科文组织"
        "世界文化景观遗产。庐山以雄、奇、险、秀闻名，山间常年云雾缭绕，苏轼名句'不识庐山真面目'"
        "即源于此。景区含花径、锦绣谷、险峰、天桥、美庐别墅和含鄱口等核心景点。"
        "美庐为蒋介石与宋美龄的避暑别墅，含鄱口可远眺鄱阳湖与长江交汇。景区内提供景交车接驳。"
    ),
    "花径": (
        "位于庐山牯岭街西南约2公里处，为唐代诗人白居易任江州司马时修筑的游览步道。"
        "白居易在此写下'长恨春归无觅处，不知转入此中来'的诗句，并手书'花径'二字刻于石门之上。"
        "步道长约1.5公里，沿路种植桃树和杜鹃，春季花开时节色彩缤纷，尽头为如琴湖畔的花径亭。"
    ),
    "锦绣谷": (
        "位于庐山牯岭南侧，是一条长约1.5公里的断层峡谷，谷深约60至100米不等，"
        "由冰川侵蚀和地壳抬升共同作用形成。谷中有四时绽放的野花和奇形怪石，"
        "沿峭壁修建的悬空栈道蜿蜒穿行于谷间，途中有'天桥'——一块天然横跨两崖的巨岩，"
        "以及'险峰'——一处突出于悬崖之外的观景平台，站立其上可俯瞰峡谷全貌。"
    ),
    "美庐": (
        "位于庐山牯岭东谷，为英国医生巴瑞于1903年建造的西式别墅，1933年由蒋介石购得并命名'美庐'，"
        "意为'美龄之庐'。别墅占地约4900平方米，建筑为两层石砌结构，室内保留了原貌的陈设和家具。"
        "庭院中保留宋美龄手植的凌霄花藤，现作为历史陈列馆对游客开放。"
    ),
    "含鄱口": (
        "位于庐山牯岭东南约4公里处，海拔1211米，两山对峙形成天然豁口，正对鄱阳湖。"
        "清晨在此可观赏旭日从湖面升起，湖光山色融为一体，为庐山最佳日出观赏点。"
        "平台建有含鄱亭和望鄱亭两座观景亭，天气晴朗时视野可达数十公里之外的鄱阳湖与长江交汇处。"
    ),

    # D5
    "东林大佛": (
        "位于九江市庐山东林寺净土苑内，为全球最高的阿弥陀佛铜像，总高81米（含莲花座），"
        "于2013年落成开光。佛像由青铜铸造，表面贴金，底座为三层汉白玉须弥座。"
        "东林寺始建于东晋太元十一年（公元386年），为净土宗祖庭，寺内藏有唐代经幢和宋代铁塔。"
        "免费参观，登莲花座需攀登约900级台阶。"
    ),
    "天空之城": (
        "位于湖北省阳新县王英镇仙岛湖风景区，核心景观是一座海拔520.1314米的玻璃观景平台，"
        "于2019年建成开放。观景台悬挑于山体外约26米，透明玻璃地板可直视下方百米深的峡谷，"
        "平台面积约600平方米，设有咖啡厅和拍照打卡点。景区含缆车往返接驳，"
        "从山顶俯瞰仙岛湖中星罗棋布的1002座岛屿，视野极为开阔壮丽。"
    ),
    "借东风": (
        "位于赤壁市三国赤壁古战场景区内的沉浸式实景表演秀，以赤壁之战中诸葛亮借东风的历史典故为蓝本，"
        "运用声光电技术和数百名演员在水上舞台演出，时长约40分钟。"
        "此节目为景区配套项目，如遇恶劣天气或不可抗力因素取消演出，无门票退费。"
    ),

    # D6
    "赤壁古战场": (
        "位于湖北省赤壁市西北38公里处的长江南岸，为公元208年三国赤壁之战的发生地，"
        "是全国重点文物保护单位。景区沿江分布有赤壁摩崖石刻（传为周瑜手书'赤壁'二字）、"
        "拜风台、凤雏庵和周瑜塑像等遗迹。含景区电瓶车接驳，沿江步道可远眺长江对岸的乌林古战场。"
    ),
    "羊楼洞古镇": (
        "位于赤壁市西南约30公里处，为中国历史文化名村，因明清时期盛产青砖茶而闻名，"
        "是万里茶道的重要源头之一。古镇保存有长约2200米的明清石板街，街道两旁为前店后宅式"
        "的砖木结构老建筑。赠送每位游客品尝当地青砖茶一杯，可在老茶馆体验茶文化。"
    ),
    "楚河汉街": (
        "位于武汉中央文化区内，全长约1.5公里，是沿楚河而建的商业步行街，于2011年开业。"
        "街区建筑融合民国风格与现代设计，汇集国际品牌店铺、餐厅和汉秀剧场。"
        "楚河为人工开凿的运河，连接东湖与沙湖，夜间两岸灯光璀璨，"
        "设有游船码头可供乘船游览，沿岸有街头艺人表演和夜市摊位。"
    ),

    # D7
    "黄鹤楼": (
        "位于武汉市武昌区蛇山之巅，濒临长江，始建于三国时期吴黄武二年（公元223年），"
        "屡毁屡建，现存建筑为1985年按清代同治年间样式重建，高51.4米，为五层攒尖顶楼阁。"
        "楼内陈列历代黄鹤楼模型和相关文物，登顶可360度俯瞰武汉三镇与长江大桥全景。含电瓶车接驳。"
    ),
    "户部巷": (
        "位于武汉市武昌区司门口，是一条长约150米的百年老巷，以集中展示武汉地方小吃而闻名，"
        "被誉为'汉味小吃第一巷'，已有400余年历史。巷内聚集了数十家老字号小吃店，"
        "代表品种包括热干面、豆皮、面窝和糊汤粉等，游客可在此自费品尝地道的武汉早点风味。"
    ),
    "归元寺": (
        "位于武汉市汉阳区翠微路，始建于清顺治十五年（1658年），为武汉四大佛教丛林之一，"
        "以五百罗汉堂最为著名。罗汉堂内供奉500尊真人大小、形态各异的贴金罗汉塑像，"
        "由清代民间匠师历时九年塑成，每尊罗汉面部表情和手持法器各不相同。"
        "寺内还藏有清代《龙藏》经版和缅甸玉佛一尊，为武汉市重点文物保护单位。"
    ),
    "武汉郁金香主题公园": (
        "位于武汉市东西湖区，占地约600亩，每年春季举办郁金香花展，种植郁金香品种超过100种、"
        "数量达200万株，花期为3月中旬至4月中旬。园区按荷兰风情设计，设有风车、木鞋和彩色花田，"
        "花季之外种植向日葵和格桑花等其他花卉。最佳观赏期为3月下旬，"
        "游客可乘坐园区观光小火车穿行于花海之间，是春季摄影的理想场所。"
    ),
    "江汉路步行街": (
        "位于武汉市汉口中心区，全长1210米，是中国最长的步行商业街之一，有百年商业历史。"
        "街道两侧保存有20世纪初建造的欧式、折衷主义风格建筑20余栋，"
        "其中包括江汉关大楼（1924年建成，现为博物馆），沿街汇聚了国内外众多商业品牌和本地老字号店铺，"
        "入夜后霓虹招牌与老建筑相映，呈现独特的城市风貌。"
    ),
    "武汉江滩": (
        "位于武汉市汉口沿江大道外侧的长江江滩公园，全长约7公里，平均宽度约160米，"
        "为亚洲最大的滨江公园之一。公园分三层平台设计，分别对应不同水位季节使用，"
        "沿江设有观景步道、草坪绿地和亲水平台，可近距离观看长江航运船只往来"
        "和武汉长江二桥的壮丽景观，傍晚时分市民在此散步放风筝，是体验武汉日常生活的好去处。"
    ),

    # D8
    "晴川阁": (
        "位于武汉市汉阳区龟山东麓长江边，始建于明嘉靖年间（1522-1566年），与黄鹤楼隔江相望，"
        "取唐代崔颢诗'晴川历历汉阳树'之意命名。现存建筑为1986年按清代样式重建，"
        "阁高17.5米，为重檐歇山顶，二楼观景台可俯瞰长江与武汉长江大桥。阁旁种植樱花树，春季可观樱。"
    ),
    "万里长江第一桥": (
        "正式名称为武汉长江大桥，横跨武昌蛇山与汉阳龟山之间，全长1670米，1957年建成通车，"
        "是长江上第一座公铁两用桥，被列为全国重点文物保护单位。上层为双向四车道公路，"
        "下层为双线铁路，桥头设有观景平台可步行上桥观看长江航运和两岸城市天际线。"
    ),
    "大禹神话园": (
        "位于武汉市汉阳区晴川阁旁的汉阳江滩内，是以大禹治水传说为主题的雕塑公园，占地约15万平方米。"
        "园内有大型群雕《大禹治水》和浮雕墙讲述大禹'三过家门而不入'等故事，"
        "沿江布置数十组石刻雕塑，将中国古代神话传说以视觉化方式呈现给游客。"
    ),
    "武商梦时代广场": (
        "位于武汉市武昌区武珞路，于2022年开业，总建筑面积约80万平方米，"
        "是目前全球最大的单体购物中心之一。商场内部设有室内滑雪场、室内过山车和大型水族馆，"
        "汇集了约800家品牌店铺、百余家餐饮餐厅和IMAX影院，"
        "为游客提供购物、餐饮与娱乐一站式体验，是武汉新兴的商业地标。"
    ),
}

# Verify all descriptions
for name, desc in SIGHT_DESCRIPTIONS.items():
    c = cn_count(desc)
    status = "OK" if 100 <= c <= 150 else "WARNING"
    if status == "WARNING":
        print(f"WARNING: {name} has {c} Chinese chars (target 100-150)")

# ============================================================
# ITINERARY DATA
# ============================================================

# Day data: (day_label, route, time_slots_with_sights, meals, hotel)
# time_slots: list of (label, sight_name_or_text)
# meals: dict with 'lunch', 'dinner' or None
# hotel: string

ITINERARY = [
    {
        'day': 'D1',
        'title': '第一天',
        'route': '出发地 / 武汉',
        'hotel': '国际五星武汉卓尔万豪酒店\n或武昌威斯汀酒店或同级',
        'slots': [
            ('全天', [
                '抵达武汉天河国际机场，导游接团后乘专车前往酒店办理入住。'
                '本日不安排观光行程，客人可在酒店休息或自行探索周边区域。',
            ]),
        ],
        'meals': None,
    },
    {
        'day': 'D2',
        'title': '第二天',
        'route': '武汉 / 蕲春',
        'hotel': '准五星蕲春万达颐华酒店或同级',
        'slots': [
            ('上午', ['东湖植物馆']),
            ('上午', ['武汉光子号单轨悬挂列车']),
            ('下午', ['乘车前往蕲春（车程约2小时），抵达后参观李时珍纪念馆（含专业中文讲解员）']),
        ],
        'meals': {'lunch': '长春观素食 100元/人', 'dinner': '中式素食 80元/人'},
        'note': '若行程日期为3月20日至3月30日，东湖植物馆将更换为武汉东湖樱花园',
    },
    {
        'day': 'D3',
        'title': '第三天',
        'route': '蕲春 / 黄梅',
        'hotel': '问梅村民宿',
        'slots': [
            ('上午', ['乘车前往黄梅（车程约1小时），抵达后参访四祖寺']),
            ('上午', ['五祖寺']),
            ('下午', ['老祖寺']),
            ('晚上', ['夜宿问梅村，体验禅意山居生活']),
        ],
        'meals': {'lunch': '五祖寺素食 100元/人', 'dinner': '中式素食 80元/人'},
    },
    {
        'day': 'D4',
        'title': '第四天',
        'route': '黄梅 / 九江',
        'hotel': '五星九江远洲喜来登酒店\n或西海君澜酒店或同级',
        'slots': [
            ('上午', ['乘车前往庐山（车程约2小时），抵达后游览庐山风景区（含景交车）']),
            ('上午', ['花径', '锦绣谷', '险峰', '天桥']),
            ('下午', ['美庐']),
            ('下午', ['含鄱口']),
        ],
        'meals': {'lunch': '中式素食 80元/人', 'dinner': '中式素食 80元/人'},
    },
    {
        'day': 'D5',
        'title': '第五天',
        'route': '九江 / 阳新 / 赤壁',
        'hotel': '赤壁鼎途国际酒店或同级',
        'slots': [
            ('上午', ['东林大佛']),
            ('下午', ['乘车前往阳新（车程约2小时），游览天空之城（含缆车往返）']),
            ('傍晚', ['乘车前往赤壁（车程约2小时）']),
            ('晚上', ['晚餐后欣赏沉浸式实景表演秀《借东风》']),
        ],
        'meals': {'lunch': '中式素食 80元/人', 'dinner': '中式素食 80元/人'},
        'note': '《借东风》为景区配套项目，如因不可抗力取消演出无门票可退',
    },
    {
        'day': 'D6',
        'title': '第六天',
        'route': '赤壁 / 武汉',
        'hotel': '国际五星武汉卓尔万豪酒店\n或武昌威斯汀酒店或同级',
        'slots': [
            ('上午', ['赤壁古战场（含电瓶车）']),
            ('上午', ['羊楼洞古镇（赠送品尝青砖茶一杯）']),
            ('下午', ['乘车返回武汉（车程约2小时），抵达后游览楚河汉街']),
        ],
        'meals': {'lunch': '中式素食 80元/人', 'dinner': '云义堂素食 150元/人'},
    },
    {
        'day': 'D7',
        'title': '第七天',
        'route': '武汉',
        'hotel': '国际五星武汉卓尔万豪酒店\n或武昌威斯汀酒店或同级',
        'slots': [
            ('上午', ['黄鹤楼（含电瓶车）']),
            ('上午', ['户部巷']),
            ('下午', ['归元寺']),
            ('下午', ['武汉郁金香主题公园']),
            ('傍晚', ['江汉路步行街', '武汉江滩']),
        ],
        'meals': {'lunch': '朴门上素 200元/人', 'dinner': '中式素食 100元/人'},
    },
    {
        'day': 'D8',
        'title': '第八天',
        'route': '武汉 / 出发地',
        'hotel': None,
        'slots': [
            ('上午', ['晴川阁（春季樱花季可赏樱花）']),
            ('上午', ['万里长江第一桥（武汉长江大桥）']),
            ('上午', ['大禹神话园']),
            ('中午', ['武商梦时代广场（自由购物及午餐）']),
            ('下午', ['前往武汉天河国际机场，结束愉快的湖北之旅']),
        ],
        'meals': {'lunch': '宝通寺素斋 100元/人'},
    },
]

# Pricing info
PRICING = {
    'price': 'RMB 4,280 / 人（15+1FOC）',
    'single_supplement': 'RMB 1,550',
    'tip': 'RMB 200',
    'includes': [
        '住宿（2人1室）',
        '用餐（7早13正）',
        '2+1三排豪华巴士',
        '所列景点首道门票',
        '所列风味餐',
        '中文导游讲解',
    ],
    'excludes': [
        '小费 RMB 25/人/天 × 8天 = RMB 200',
    ],
    'gift': '每人每天矿泉水一瓶',
    'notes': [
        '如团队在行程中发生任何问题，请领队或团员第一时间与本社联系，以争取最及时适当的处理，团队回国本公司不受理任何投诉事宜。',
        '若客人自愿放弃旅游行程，我社将不予退返费用。',
        '以上报价大型会议期间，我公司同时保留因为各地酒店临时调整而做出适当调整的可能性，争取在第一时间通知贵公司。',
        '因为不可抗拒因素致使行程变更，我社可根据实际情况调整。',
        '领队无证，门票请自理。',
    ],
}

# ============================================================
# BUILD DOCUMENT
# ============================================================

def add_page_number_footer(section):
    """Add centered PAGE/NUMPAGES footer to a section using OOXML fields."""
    footer = section.footer
    footer.is_linked_to_previous = False
    # Remove any existing paragraphs
    for p in footer.paragraphs:
        p._element.getparent().remove(p._element)
    # Add a single centered paragraph with PAGE / NUMPAGES
    p = footer.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_no_spacing(p)
    # Build the field XML manually
    fld_xml = (
        '<w:fldSimple %s w:instr=" PAGE "><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Arial"/>'
        '<w:sz w:val="16"/><w:color w:val="999999"/></w:rPr><w:t>1</w:t></w:r></w:fldSimple>'
    ) % nsdecls('w')
    slash_run = parse_xml(
        '<w:r %s><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Arial"/>'
        '<w:sz w:val="16"/><w:color w:val="999999"/></w:rPr><w:t>/</w:t></w:r>' % nsdecls('w')
    )
    numpages_fld_xml = (
        '<w:fldSimple %s w:instr=" NUMPAGES "><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Arial"/>'
        '<w:sz w:val="16"/><w:color w:val="999999"/></w:rPr><w:t>1</w:t></w:r></w:fldSimple>'
    ) % nsdecls('w')
    p._element.append(parse_xml(fld_xml))
    p._element.append(slash_run)
    p._element.append(parse_xml(numpages_fld_xml))

def build_document():
    doc = Document()
    
    # Page setup: A4 portrait, 2cm margins
    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)
    section.left_margin = Cm(2)
    section.right_margin = Cm(2)
    add_page_number_footer(section)
    
    # Default font
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(10)
    font.color.rgb = DARK_GRAY
    style.element.rPr.rFonts.set(qn('w:eastAsia'), 'Arial')
    
    # ========== COVER PAGE ==========
    # Add vertical space
    for _ in range(6):
        p = doc.add_paragraph()
        set_no_spacing(p)
    
    # Main title
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = add_formatted_run(title_p, '湖北武汉 + 庐山', 'Arial', Pt(28), bold=True, color=DEEP_BLUE)
    set_no_spacing(title_p)
    
    # Subtitle
    sub_p = doc.add_paragraph()
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = add_formatted_run(sub_p, '禅意素食文化之旅 8日深度游', 'Arial', Pt(16), color=WARM_GOLD)
    set_no_spacing(sub_p)
    
    # Divider
    div_p = doc.add_paragraph()
    div_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = add_formatted_run(div_p, '━' * 30, 'Arial', Pt(10), color=LIGHT_GRAY)
    set_no_spacing(div_p)
    
    # Info block
    info_items = [
        ('出发日期', '2027年3月 — 4月'),
        ('客源地', '新加坡 / 马来西亚'),
        ('地接社', '湖北彩途国旅 王业'),
    ]
    for label, value in info_items:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_formatted_run(p, f'{label}：', 'Arial', Pt(11), bold=True, color=DARK_GRAY)
        add_formatted_run(p, value, 'Arial', Pt(11), color=DARK_GRAY)
        set_no_spacing(p)
    
    # Highlights
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_no_spacing(p)
    
    highlights_p = doc.add_paragraph()
    highlights_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_formatted_run(highlights_p, '行程亮点：', 'Arial', Pt(11), bold=True, color=DEEP_BLUE)
    set_no_spacing(highlights_p)
    
    highlights = [
        '登临世界文化景观遗产 —— 庐山',
        '朝拜禅宗祖庭 —— 四祖寺、五祖寺',
        '参访千年古刹 —— 归元寺五百罗汉堂',
        '登江南名楼 —— 黄鹤楼俯瞰长江',
        '体验三国文化 —— 赤壁古战场',
    ]
    for h in highlights:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_formatted_run(p, f'  {h}', 'Arial', Pt(10), color=MEDIUM_GRAY)
        set_no_spacing(p)
    
    # Price teaser
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_no_spacing(p)
    
    price_p = doc.add_paragraph()
    price_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_formatted_run(price_p, f'报价：{PRICING["price"]}', 'Arial', Pt(14), bold=True, color=WARM_GOLD)
    set_no_spacing(price_p)
    
    add_page_break(doc)
    
    # ========== DAY-BY-DAY ITINERARY ==========
    for day_data in ITINERARY:
        _build_day_table(doc, day_data)
        cleanup_empty_rows(doc.tables[-1])  # safety: remove any blank rows
    
    # ========== PRICING SECTION ==========
    _build_pricing_page(doc)
    
    return doc

def _build_day_table(doc, day_data):
    """Build a day itinerary table."""
    day_num = day_data['day']
    day_title = day_data['title']
    route = day_data['route']
    
    # Day header
    header_p = doc.add_paragraph()
    header_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    set_no_spacing(header_p)
    add_formatted_run(header_p, f'{day_num}  {day_title}', 'Arial', Pt(18), bold=True, color=DEEP_BLUE)
    
    # Route
    route_p = doc.add_paragraph()
    set_no_spacing(route_p)
    add_formatted_run(route_p, f'路线：{route}', 'Arial', Pt(12), bold=False, color=WARM_GOLD)
    
    # Note if present
    if 'note' in day_data:
        note_p = doc.add_paragraph()
        set_no_spacing(note_p)
        add_formatted_run(note_p, f'[备注] {day_data["note"]}', 'Arial', Pt(8), italic=True, color=MEDIUM_GRAY)
    
    # Build table
    slots = day_data['slots']
    num_rows = 1 + len(slots)  # header + sight rows
    if day_data.get('meals'):
        num_rows += 1  # [Meals] row
    if day_data.get('hotel') is not None:
        num_rows += 1  # [Hotel] row
    
    table = doc.add_table(rows=num_rows, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table)
    
    # Column widths: Time=2.2cm, Sight=14.8cm (total 17cm = A4 - 2×2cm margins)
    # Set column widths on first row
    for row in table.rows:
        set_cell_width(row.cells[0], 2.2)
        set_cell_width(row.cells[1], 14.8)
    
    # ---- Header row ----
    header_cells = table.rows[0].cells
    headers = ['\u65f6\u95f4', '\u89c2\u5149\u5185\u5bb9']
    for i, (cell, text) in enumerate(zip(header_cells, headers)):
        set_cell_shading(cell, DEEP_BLUE_HEX)
        set_cell_va(cell, 'center')
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_no_spacing(p)
        add_formatted_run(p, text, 'Arial', Pt(9), bold=True, color=WHITE)
    
    # ---- Sight rows ----
    row_idx = 1
    for time_label, sight_items in slots:
        cells = table.rows[row_idx].cells
        
        # Time cell
        set_cell_shading(cells[0], BG_LIGHT_HEX)
        set_cell_va(cells[0], 'center')
        p = cells[0].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_no_spacing(p)
        add_formatted_run(p, time_label, 'Arial', Pt(9), bold=True, color=DEEP_BLUE)
        
        # Sight cell (spans the main content area)
        set_cell_va(cells[1], 'center')
        p = cells[1].paragraphs[0]
        set_no_spacing(p)
        
        for j, item in enumerate(sight_items):
            if j > 0:
                # Add a small separator
                sep_p = cells[1].add_paragraph()
                set_no_spacing(sep_p)
            
            # Check if it's a sight with a description
            if item in SIGHT_DESCRIPTIONS:
                desc = SIGHT_DESCRIPTIONS[item]
                # Sight name
                if j > 0:
                    name_p = cells[1].add_paragraph()
                else:
                    name_p = p
                set_no_spacing(name_p)
                add_formatted_run(name_p, f'{item}', 'Arial', Pt(10), bold=True, color=DARK_GREEN)
                # Description
                desc_p = cells[1].add_paragraph()
                set_no_spacing(desc_p)
                add_formatted_run(desc_p, desc, 'Arial', Pt(8), color=DARK_GRAY)
            else:
                # Plain text item (transit, notes, etc.)
                if j > 0:
                    text_p = cells[1].add_paragraph()
                else:
                    text_p = p
                set_no_spacing(text_p)
                add_formatted_run(text_p, item, 'Arial', Pt(9), color=DARK_GRAY)
        
        row_idx += 1
    
    # ---- Meals summary row ----
    if day_data['meals']:
        cells = table.rows[row_idx].cells
        # Merge both cells for full-width meals row
        cells[0].merge(cells[1])
        set_cell_shading(cells[0], BG_GREEN_HEX)
        set_cell_va(cells[0], 'center')
        p = cells[0].paragraphs[0]
        set_no_spacing(p)
        
        meals = day_data['meals']
        meal_texts = []
        if 'breakfast' in meals:
            meal_texts.append(f'[早餐] {meals["breakfast"]}')
        if 'lunch' in meals:
            meal_texts.append(f'[中餐] {meals["lunch"]}')
        if 'dinner' in meals:
            meal_texts.append(f'[晚餐] {meals["dinner"]}')
        
        add_formatted_run(p, '[Meals]  ', 'Arial', Pt(8), bold=True, color=DARK_GREEN)
        add_formatted_run(p, ' | '.join(meal_texts), 'Arial', Pt(8), bold=False, color=DARK_GREEN)
        
        row_idx += 1
    
    # ---- Hotel row ----
    has_hotel = day_data.get('hotel') is not None
    if has_hotel:
        cells = table.rows[row_idx].cells
        # Merge both cells for full-width hotel row
        cells[0].merge(cells[1])
        set_cell_shading(cells[0], BG_RED_HEX)
        set_cell_va(cells[0], 'center')
        p = cells[0].paragraphs[0]
        set_no_spacing(p)
        add_formatted_run(p, '[Hotel]  ', 'Arial', Pt(8), bold=True, color=RGBColor(0xCC, 0x33, 0x33))
        add_formatted_run(p, day_data['hotel'], 'Arial', Pt(8), color=DARK_GRAY)
        
        row_idx += 1

def _build_pricing_page(doc):
    """Build the pricing summary page."""
    # Title
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_no_spacing(title_p)
    add_formatted_run(title_p, '旅游报价', 'Arial', Pt(22), bold=True, color=DEEP_BLUE)
    
    # Divider
    div_p = doc.add_paragraph()
    div_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_formatted_run(div_p, '━' * 30, 'Arial', Pt(10), color=LIGHT_GRAY)
    set_no_spacing(div_p)
    
    # Price table
    price_table = doc.add_table(rows=3, cols=2)
    price_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(price_table)
    
    price_data = [
        ('团队报价', PRICING['price']),
        ('单房差', PRICING['single_supplement']),
        ('司导小费', PRICING['tip']),
    ]
    
    for i, (label, value) in enumerate(price_data):
        cells = price_table.rows[i].cells
        set_cell_width(cells[0], 5)
        set_cell_width(cells[1], 12)
        set_cell_shading(cells[0], BG_LIGHT_HEX)
        set_cell_va(cells[0], 'center')
        p = cells[0].paragraphs[0]
        set_no_spacing(p)
        add_formatted_run(p, label, 'Arial', Pt(11), bold=True, color=DEEP_BLUE)
        
        set_cell_va(cells[1], 'center')
        p = cells[1].paragraphs[0]
        set_no_spacing(p)
        add_formatted_run(p, value, 'Arial', Pt(11), bold=True, color=WARM_GOLD)
    
    # Includes
    inc_p = doc.add_paragraph()
    set_no_spacing(inc_p)
    add_formatted_run(inc_p, '[Price Includes] 报价包含', 'Arial', Pt(12), bold=True, color=DARK_GREEN)
    
    for item in PRICING['includes']:
        p = doc.add_paragraph()
        set_no_spacing(p)
        add_formatted_run(p, f'  + {item}', 'Arial', Pt(10), color=DARK_GRAY)
    
    # Excludes (no spacer paragraph — use paragraph spacing instead)
    exc_p = doc.add_paragraph()
    set_no_spacing(exc_p)
    add_formatted_run(exc_p, '[Price Excludes] 报价不含', 'Arial', Pt(12), bold=True, color=RGBColor(0xCC, 0x33, 0x33))
    
    for item in PRICING['excludes']:
        p = doc.add_paragraph()
        set_no_spacing(p)
        add_formatted_run(p, f'  - {item}', 'Arial', Pt(10), color=DARK_GRAY)
    
    # Gift
    gift_p = doc.add_paragraph()
    set_no_spacing(gift_p)
    add_formatted_run(gift_p, f'[Gift] 赠送：{PRICING["gift"]}', 'Arial', Pt(11), bold=True, color=WARM_GOLD)
    
    # SP
    sp_p = doc.add_paragraph()
    set_no_spacing(sp_p)
    add_formatted_run(sp_p, '[SP] 自费项目：无', 'Arial', Pt(11), bold=True, color=DEEP_BLUE)
    
    # Notes
    notes_p = doc.add_paragraph()
    set_no_spacing(notes_p)
    add_formatted_run(notes_p, '[Notes] 备注事项', 'Arial', Pt(12), bold=True, color=DEEP_BLUE)
    
    for i, note in enumerate(PRICING['notes'], 1):
        p = doc.add_paragraph()
        set_no_spacing(p)
        add_formatted_run(p, f'{i}. {note}', 'Arial', Pt(8), color=MEDIUM_GRAY)

# ============================================================
# MAIN
# ============================================================
def main():
    output_dir = r'C:\Users\DELL\Desktop'
    output_path = os.path.join(output_dir, '湖北武汉+庐山8日.docx')
    
    print("Building document...")
    doc = build_document()
    
    print(f"Saving to {output_path}...")
    doc.save(output_path)
    
    print("Cleaning up DOCX package...")
    cleanup_docx(output_path)
    
    print("Verifying character counts...")
    all_ok = True
    for name, desc in SIGHT_DESCRIPTIONS.items():
        c = cn_count(desc)
        if c < 100 or c > 150:
            print(f"  WARNING: {name} has {c} Chinese chars (target 100-150)")
            all_ok = False
    
    if all_ok:
        print("  All descriptions within 100-150 character range.")
    
    print(f"\nDone! Output: {output_path}")

if __name__ == '__main__':
    main()
