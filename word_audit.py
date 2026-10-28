#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""文档校阅 v1.1：Word 格式批量初筛器，标准库实现，无第三方依赖。"""
from __future__ import annotations
import argparse
import csv
import datetime as dt
import html
import json
import re
import sys
import zipfile
from collections import Counter
from pathlib import Path
from xml.etree import ElementTree as ET

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
Q = lambda name: f'{{{W}}}{name}'
ATTR = lambda e, key: e.get(Q(key)) if e is not None else None
A4_W, A4_H = 11906, 16838  # dxa/twips；打印机/软件可出现微小取整差异
EXPECTED_MARGIN = 1440  # 1in=2.54cm
TITLE_PT, TEXT_PT, CAPTION_PT = 16, 12, 10
MAX_PER_CATEGORY = 18
SEV = ['错误', '待复核', '提示']
COLUMNS = ['文件', '等级', '类别', '位置', '检测值', '规范/建议', '说明']


def tag(parent, name):
    return parent.find(Q(name)) if parent is not None else None


def value(parent, name, attr='val'):
    return ATTR(tag(parent, name), attr)


def first_value(parents, parent_tag, tagname, attr='val'):
    for item in parents:
        elm = tag(tag(item, parent_tag), tagname) if parent_tag else tag(item, tagname)
        if elm is not None and ATTR(elm, attr) is not None:
            return ATTR(elm, attr)
    return None


def onoff(s):
    if s is None: return None
    return str(s).lower() not in ('false', '0', 'off', 'no')


def read_zip_xml(z, path):
    try:
        with z.open(path) as f:
            return ET.parse(f).getroot()
    except KeyError:
        return None


def text_of(p):
    # 跳过修订删除的文本及域指令，保留常规 w:t。
    parts=[]
    def walk(node):
        if node.tag in (Q('del'), Q('moveFrom')):
            return
        if node.tag==Q('t'):
            parts.append(node.text or '')
        elif node.tag==Q('tab'):
            parts.append('\t')
        for child in node:
            walk(child)
    walk(p)
    return ''.join(parts).strip()


def contains_han(s):
    return any('\u3400' <= c <= '\u9fff' or '\uf900' <= c <= '\ufaff' for c in s)


def contains_latin(s):
    return bool(re.search(r'[A-Za-z]', s))


def small_excerpt(s):
    return (s[:50] + '…') if len(s)>50 else s


def numeric(s):
    try: return int(s)
    except (TypeError, ValueError): return None


class Styles:
    def __init__(self, styles_root):
        self.styles={}; self.doc_p={}; self.doc_r={}
        if styles_root is None: return
        for e in styles_root.findall(Q('style')):
            sid = ATTR(e,'styleId')
            if sid:
                self.styles[sid] = {'based':value(e,'basedOn'), 'pPr':tag(e,'pPr'),
                                    'rPr':tag(e,'rPr'), 'type':ATTR(e,'type') or 'paragraph',
                                    'name':value(e,'name') or sid}
        defaults=tag(styles_root,'docDefaults')
        pd=tag(defaults,'pPrDefault'); rd=tag(defaults,'rPrDefault')
        self.doc_p=tag(pd,'pPr')
        self.doc_r=tag(rd,'rPr')

    def chain(self, sid):
        seen=set()
        while sid and sid in self.styles and sid not in seen:
            seen.add(sid)
            yield self.styles[sid]
            sid=self.styles[sid]['based']

    def p_props(self,p):
        ppr=tag(p,'pPr'); sid=value(ppr,'pStyle') or 'Normal'
        return ([ppr] if ppr is not None else []) + [s['pPr'] for s in self.chain(sid) if s['pPr'] is not None] + ([self.doc_p] if self.doc_p is not None else []), sid

    def r_props(self,p,run):
        rr=tag(run,'rPr')
        rp=value(rr,'rStyle')
        pp,sid=self.p_props(p)
        chain=[rr] if rr is not None else []
        chain += [s['rPr'] for s in self.chain(rp) if s['rPr'] is not None]
        chain += [s['rPr'] for s in self.chain(sid) if s['rPr'] is not None]
        if self.doc_r is not None: chain.append(self.doc_r)
        return chain

    def r_prop(self,p,run,prop,attr='val'):
        for rpr in self.r_props(p,run):
            el=tag(rpr,prop)
            if el is not None:
                v=ATTR(el,attr)
                if v is not None: return v
                if prop=='b': return 'true'  # <w:b/> 代表加粗
        return None

    def font(self,p,run,script):
        attr='eastAsia' if script=='zh' else 'ascii'
        theme='eastAsiaTheme' if script=='zh' else 'asciiTheme'
        for rpr in self.r_props(p,run):
            rf=tag(rpr,'rFonts')
            if rf is None: continue
            found=ATTR(rf,attr)
            if found: return ('known',found)
            if script=='en' and ATTR(rf,'hAnsi'): return ('known',ATTR(rf,'hAnsi'))
            if ATTR(rf,theme) or (script=='en' and ATTR(rf,'hAnsiTheme')):
                return ('theme',ATTR(rf,theme) or ATTR(rf,'hAnsiTheme'))
        return ('unknown',None)

    def paragraph_format(self,p,kind):
        props,_=self.p_props(p)
        for parent in props:
            e=tag(parent,kind)
            if e is not None: return e
        return None


def category_of(p, index, ref_state, styles):
    s=text_of(p)
    if index==0: return '文章标题'
    if index==1: return '作者姓名'
    if index==2: return '所属单位'
    if re.match(r'^\s*(?:图\s*\d+|表\s*\d+|Figure\s+\d+|Table\s+\d+)\s*[.．:：、\-]?',s,re.I):
        return '图表标题'
    if re.match(r'^\s*(?:\d+(?:\.\d+)+)\s+\S',s): return '二级标题'
    if re.match(r'^\s*\d+[.．、]?\s+\S',s): return '一级标题'
    pprops,sid=styles.p_props(p)
    name=(styles.styles.get(sid,{}).get('name') or '').lower()
    if ('heading 2' in name or '标题 2' in name or 'heading2' in sid.lower()): return '二级标题'
    if ('heading 1' in name or '标题 1' in name or 'heading1' in sid.lower()): return '一级标题'
    if re.match(r'^\s*(?:参考文献|references)\s*$',s,re.I): return '一级标题'
    if re.match(r'^\s*\[\s*\d+\s*\]',s) and ref_state: return '参考文献条目'
    return '正文'


def dxa_to_pt(s): return (numeric(s)/20) if numeric(s) is not None else None


class Auditor:
    def __init__(self,filename):
        self.filename=Path(filename).name; self.rows=[]; self.limits=Counter()
        self.paras=[]; self.all_text=''; self.ref_start=-1

    def issue(self,severity,category,where,observed,expect,explain):
        if self.limits[category] >= MAX_PER_CATEGORY:
            if self.limits[category] == MAX_PER_CATEGORY:
                self.rows.append([self.filename,'提示',category,'全文','同类问题较多',
                                  '仅显示前 18 处','其余同类位置未逐条列出'])
            self.limits[category]+=1; return
        self.rows.append([self.filename,severity,category,where,str(observed),str(expect),explain])
        self.limits[category]+=1

    def expect_num(self,cat,loc,found,expected,tol,units='pt'):
        if found is None:
            self.issue('待复核',cat,loc,'未取得明确数值',f'{expected:g} {units}',
                       '可能使用了未解析的样式、主题或默认值，请在 Word 中核对')
        elif abs(found-expected)>tol:
            self.issue('错误',cat,loc,f'{found:g} {units}',f'{expected:g} {units}',
                       '检测到数值不符合本次编辑团队的格式规范')

    def check_document(self,path):
        try:
            with zipfile.ZipFile(path,'r') as z:
                if z.testzip() is not None: raise ValueError('ZIP 数据损坏')
                root=read_zip_xml(z,'word/document.xml')
                if root is None: raise ValueError('缺少 word/document.xml')
                styles=Styles(read_zip_xml(z,'word/styles.xml'))
                self._page(root)
                self._body(root,styles)
                self._structure(z,root)
        except (zipfile.BadZipFile, ET.ParseError, OSError, ValueError, RuntimeError) as e:
            self.issue('错误','无法解析','文件',type(e).__name__,'有效的 .docx 文件',str(e))
        return self.rows

    def _page(self,root):
        sects=list(root.iter(Q('sectPr')))
        if not sects:
            self.issue('待复核','页面设置','全文','未解析到 section 定义','A4，2.54 cm','需人工检查')
            return
        for i,sect in enumerate(sects,1):
            loc=f'第 {i} 节'
            pg=tag(sect,'pgSz'); mar=tag(sect,'pgMar')
            w=numeric(ATTR(pg,'w')); h=numeric(ATTR(pg,'h'))
            if w is None or h is None:
                self.issue('待复核','页面大小',loc,'未知','A4 210×297 mm','页面尺寸未显式保存')
            else:
                if not (abs(w-A4_W)<=22 and abs(h-A4_H)<=22):
                    self.issue('错误','页面大小',loc,f'{w}×{h} twips','A4 约 11906×16838 twips','请设置 A4 页面，不能使用 Letter')
            for key,cn in [('top','上'),('bottom','下'),('left','左'),('right','右')]:
                x=numeric(ATTR(mar,key))
                if x is None:
                    self.issue('待复核','页边距',f'{loc}／{cn}','未显式给出','2.54 cm','需人工查看')
                elif abs(x-EXPECTED_MARGIN)>8:
                    self.issue('错误','页边距',f'{loc}／{cn}',f'{x/567:.2f} cm','2.54 cm','页边距与通知不一致')

    def _body(self,root,styles):
        body=tag(root,'body')
        if body is None: return
        # 正文 + 表格段落。表格单元格与标题段落不同；不用普通正文首行缩进规则检查表格。
        all_para=list(body.iter(Q('p')))
        table_paras={id(p) for t in body.iter(Q('tbl')) for p in t.iter(Q('p'))}
        self.paras=[(p,text_of(p),id(p) in table_paras) for p in all_para if text_of(p)]
        self.all_text='\n'.join(txt for p,txt,_ in self.paras)
        self.ref_start=next((i for i,(p,t,_) in enumerate(self.paras) if re.match(r'^(?:\d+[.．、]?\s*)?(?:参考文献|References)\s*$',t,re.I)),len(self.paras))
        if len(self.paras)<3:
            self.issue('错误','论文结构','文档开头',f'仅 {len(self.paras)} 个非空段落','标题＋作者＋单位','请核实前三个部分是否存在')
        first3=self.paras[:3]
        for i,(p,s,in_table) in enumerate(self.paras):
            if in_table:
                # 表格内容的对齐/缩进由表格自身决定，通知未统一规定，避免误报
                continue
            role=category_of(p,i,i>self.ref_start,styles)
            loc=f'正文第 {i+1} 段：{small_excerpt(s)}'
            if role in ('文章标题','作者姓名','所属单位') and i>2: continue
            self._paragraph_check(p,s,role,loc,styles,i)
            self._run_check(p,s,role,loc,styles)
        # 姓名统一中文：中文字符可识别，但拼音或留学生英文姓名涉及例外，只能提示人工检查
        if len(first3)>=2 and not contains_han(first3[1][1]):
            self.issue('待复核','姓名语言','正文第 2 段',small_excerpt(first3[1][1]),'中文姓名',
                       '如为留学生或需要外文姓名，请编辑团队人工确认')
        for idx in [0,1,2]:
            if len(self.paras)>idx and idx!=0 and re.match(r'^\d+[.．、]?\s',self.paras[idx][1]):
                self.issue('待复核','文章开头',f'正文第 {idx+1} 段','疑似章节标题','标题/姓名/所属单位','前三个非空段落可能不是所要求的结构')

    def _paragraph_check(self,p,s,role,loc,styles,index):
        props,_=styles.p_props(p)
        def field(tagname,key):
            for pp in props:
                e=tag(pp,tagname)
                if e is not None and ATTR(e,key) is not None:
                    return ATTR(e,key)
            return None
        expected_align={'文章标题':'center','作者姓名':'center','所属单位':'center',
                        '一级标题':'left','二级标题':'left','图表标题':'center',
                        '正文':'both','参考文献条目':'both'}[role]
        align=field('jc','val') or 'left'   # Word 的未定义默认通常左对齐
        if role=='正文' and index>self.ref_start: # 参考文献中的非编号文字不强制对齐
            pass
        elif align!=expected_align:
            self.issue('错误','段落对齐',loc,align,expected_align,'请使用段落属性，而非空格排版')
        spacing_before=field('spacing','before')
        spacing_after=field('spacing','after')
        # 不设置段前后 = 0pt（没有规定的覆盖继承）；Word可能存在上下文间距自动压缩，报告仍需复核
        before=dxa_to_pt(spacing_before) if spacing_before is not None else 0.0
        after=dxa_to_pt(spacing_after) if spacing_after is not None else 0.0
        goal={'文章标题':(0,12),'作者姓名':(0,0),'所属单位':(0,12),
              '一级标题':(12,6),'二级标题':(6,0),'图表标题':(None,None),
              '正文':(0,0),'参考文献条目':(0,0)}[role]
        for v,target,what in [(before,goal[0],'段前'),(after,goal[1],'段后')]:
            if target is not None and abs(v-target)>1.0:
                self.issue('错误','段落间距',loc,f'{what} {v:g} pt',f'{what} {target:g} pt',
                           '使用 Word 段落设置，不要手动插入空行')
        # 单倍行距: auto 240 = 1.0；未显式定义时通常也是单倍
        spval=field('spacing','line'); rule=field('spacing','lineRule') or 'auto'
        if spval is not None:
            if rule != 'auto':
                self.issue('错误','行距',loc,f'{spval} / {rule}','单倍行距（auto / 240）',
                           '固定值/最小值不等于单倍行距')
            elif abs((numeric(spval) or 0)-240)>3:
                self.issue('错误','行距',loc,f'{spval} / auto','240 / auto','此值与单倍行距不符')
        if role=='正文' and index<self.ref_start:
            fc=field('ind','firstLineChars'); fl=field('ind','firstLine')
            if fc is not None and numeric(fc)!=200:
                self.issue('错误','首行缩进',loc,f'{fc}/100 字符','2 字符','请使用首行缩进 2 字符')
            elif fc is None and fl is None:
                self.issue('错误','首行缩进',loc,'0 或未设置','2 字符','正文段落应首行缩进 2 字符')
            elif fc is None and fl is not None:
                if abs((numeric(fl) or 0)-480)>60:
                    self.issue('错误','首行缩进',loc,f'{fl} twips','2 字符','缩进与 12pt 中文正文两字符近似值不符')
                else:
                    self.issue('待复核','首行缩进',loc,f'{fl} twips','2 字符（字符单位）','只给出 twips，不能直接证明为 2 字符')
        elif role in ('一级标题','二级标题','文章标题','作者姓名','所属单位'):
            fc=field('ind','firstLineChars'); fl=field('ind','firstLine')
            if (numeric(fc) or 0)!=0 or (numeric(fl) or 0)!=0:
                self.issue('错误','标题/署名缩进',loc,f'firstLineChars={fc}, firstLine={fl}',
                           '不缩进','标题、作者与单位不应沿用正文首行缩进')

    def _run_check(self,p,s,role,loc,styles):
        pt = TITLE_PT if role=='文章标题' else (CAPTION_PT if role=='图表标题' else TEXT_PT)
        need_bold = True if role in ('文章标题','一级标题','二级标题') else False if role in ('作者姓名','所属单位') else None
        found_runs=[]
        for run in p.iter(Q('r')):
            rtext=''.join(t.text or '' for t in run.iter(Q('t')))
            if not rtext.strip(): continue
            found_runs.append((run,rtext))
        if not found_runs:
            self.issue('待复核','字体',loc,'无普通文字运行片段','检查 TextBox/域/特殊对象','段落可能主要由特殊对象构成')
            return
        # 对每个不同类别的错误进行合并，避免每个字/每个 run 一行刷屏。
        problems=set()
        for run,rt in found_runs:
            sz=styles.r_prop(p,run,'sz')
            if sz is not None and numeric(sz) is not None:
                actual=numeric(sz)/2
                if abs(actual-pt)>0.1: problems.add(('字号','错误',f'{actual:g} pt',f'{pt:g} pt'))
            else:
                problems.add(('字号','待复核','字号由主题/未知默认决定',f'{pt:g} pt'))
            if need_bold is not None:
                b=styles.r_prop(p,run,'b')
                if (onoff(b) if b is not None else False)!=need_bold:
                    problems.add(('粗体','错误',str(onoff(b)),'加粗' if need_bold else '不加粗'))
            # 参考文献条目中保留英文原文；中文正文/标题字体按最新通知
            if contains_han(rt):
                status,font=styles.font(p,run,'zh')
                expected='黑体' if role=='文章标题' else '宋体'
                if status=='known' and font and font not in (expected, 'SimHei' if expected=='黑体' else 'SimSun'):
                    problems.add(('中文字体','错误',font,expected))
                elif status!='known': problems.add(('中文字体','待复核',font or '未知',expected))
            if contains_latin(rt):
                status,font=styles.font(p,run,'en')
                expected='Arial' if role=='文章标题' else 'Times New Roman'
                if status=='known' and font and font.lower()!=expected.lower():
                    problems.add(('英文字体','错误',font,expected))
                elif status!='known': problems.add(('英文字体','待复核',font or '未知',expected))
        for typ,sev,actual,expected in sorted(problems):
            self.issue(sev,typ,loc,actual,expected,
                       '已检查样式继承与局部格式；主题字体、文本框等特殊情况仍需人工复核')

    def _structure(self,z,root):
        t=self.all_text
        for label,pattern in [('研究领域与论文选择',r'研究领域|论文选择'),
                              ('高水平论文分析',r'高水平论文|论文\s*A\s*[（(]高水平'),
                              ('中等水平论文分析',r'中等水平论文|论文\s*B\s*[（(]中等'),
                              ('相对普通水平论文分析',r'普通水平论文|论文\s*C\s*[（(]普通'),
                              ('比较分析',r'横向比较|比较分析'),('总结与启示',r'总结与启示|收获与启示'),
                              ('参考文献',r'参考文献')]:
            if not re.search(pattern,t,re.I):
                self.issue('待复核','内容结构',label,'未检出约定中文标题','包含该主题的章节',
                           '程序只做关键词筛查，可能存在同义标题，需人工确认')
        # 期末通知要求章节标题原则上中文；如果存在明显英文 Heading，提示修改
        for i,(p,s,is_table) in enumerate(self.paras):
            if is_table: continue
            role=category_of(p,i,i>self.ref_start,Styles(read_zip_xml(z,'word/styles.xml')))
            if role in ('一级标题','二级标题') and contains_latin(s) and not contains_han(s):
                self.issue('错误','章节语言',f'正文第 {i+1} 段',small_excerpt(s),'中文章节标题',
                           '本次最新版补充通知要求章节标题使用中文')
        # 文内与文末方括号式数字引用；这是启发式，不验证引用来源真伪
        body='\n'.join(x[1] for x in self.paras[:self.ref_start])
        end='\n'.join(x[1] for x in self.paras[self.ref_start:])
        cited={int(x) for x in re.findall(r'\[\s*(\d+)\s*\]',body)}
        listed={int(x) for x in re.findall(r'(?m)^\s*\[\s*(\d+)\s*\]',end)}
        if self.ref_start<len(self.paras) and not listed:
            self.issue('待复核','参考文献','参考文献章节','未识别到 [1] 形式的文献条目',
                       '按引用顺序编号','可能采用了不兼容的文献编号方式')
        for n in sorted(cited-listed)[:15]:
            self.issue('待复核','引用编号','正文/参考文献',f'文内 [{n}] 无对应条目',
                       '文献编号一一对应','请确认是否丢失、错号或出现交叉引用')
        for n in sorted(listed-cited)[:15]:
            self.issue('待复核','引用编号','参考文献',f'条目 [{n}] 未在正文匹配到',
                       '文献应在正文适当位置被引用','可能为正文自动域、复杂引用或多引文写法')
        if listed and listed!=set(range(1,max(listed)+1)):
            self.issue('待复核','参考文献编号','参考文献',','.join(map(str,sorted(listed)[:20])),
                       '从 [1] 开始连续编号','这里只检查编号是否连续，不验证按出现顺序排列')
        # 图、表编号的顺序仅作潜在问题提示
        for typ,pattern in [('图片',r'(?:图|Figure)\s*(\d+)\s*[.．:：]'),
                            ('表格',r'(?:表|Table)\s*(\d+)\s*[.．:：]')]:
            cap=[]
            for _,s,_ in self.paras:
                mm=re.match(r'^\s*'+pattern,s,re.I)
                if mm: cap.append(int(mm.group(1)))
            if cap and (cap!=list(range(1,len(cap)+1))):
                self.issue('待复核','图表编号',typ,','.join(map(str,cap)),
                           '按出现次序从 1 连续编号','可能缺图、跳号、重复编号，需对照图片与表格人工核对')
        if any(Q('ins')==e.tag or Q('del')==e.tag for e in root.iter()):
            self.issue('待复核','修订模式','全文','检测到修订记录','提交清洁版本','请在 Word 中审查并接受/拒绝修订')
        page_fields=[]
        for fn in z.namelist():
            if re.match(r'word/(?:header|footer)\d*\.xml$',fn):
                el=read_zip_xml(z,fn)
                if el is None: continue
                for it in el.iter():
                    if it.tag==Q('instrText') and it.text and re.search(r'\bPAGE\b',it.text,re.I): page_fields.append(fn)
                    if it.tag==Q('fldSimple') and re.search(r'\bPAGE\b',ATTR(it,'instr') or '',re.I): page_fields.append(fn)
        if page_fields:
            self.issue('待复核','页码','页眉页脚',','.join(sorted(set(page_fields))),
                       '个人稿件暂不添加页码','检测到 PAGE 字段，需要核实是否为页码')


def csv_safe(x):
    # 防止学生文件名/段落文本在 Excel 中被当成公式执行（CSV Formula Injection）。
    s=str(x)
    return "'"+s if s.lstrip().startswith(('=', '+', '-', '@')) else s


def safe_write_csv(path,rows,header):
    with path.open('w',newline='',encoding='utf-8-sig') as f:
        w=csv.writer(f); w.writerow(header)
        for row in rows: w.writerow([csv_safe(x) for x in row])


def make_html(folder,summary,issues):
    table_head=''.join('<th>'+html.escape(k)+'</th>' for k in COLUMNS)
    table_rows=[]
    for r in issues:
        cells=''.join('<td>'+html.escape(str(x))+'</td>' for x in r)
        table_rows.append('<tr class="'+('bad' if r[1]=='错误' else 'warn' if r[1]=='待复核' else 'note')+'">'+cells+'</tr>')
    ss=''.join('<tr>'+''.join('<td>'+html.escape(str(x))+'</td>' for x in row)+'</tr>' for row in summary)
    doc=f'''<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><title>论文合集格式审校报告</title>
<style>body{{font-family:Segoe UI,Microsoft YaHei,Arial,sans-serif;background:#f6f8fb;color:#1e293b;margin:30px auto;max-width:1400px;padding:0 20px}}
h1{{font-size:26px}}p{{color:#475569}}.frame{{background:white;border:1px solid #e5e7eb;border-radius:12px;padding:20px;margin:18px 0;overflow:auto}}table{{border-collapse:collapse;width:100%;font-size:13px}}td,th{{border-bottom:1px solid #e5e7eb;text-align:left;padding:9px 10px;vertical-align:top}}th{{background:#f1f5f9;position:sticky;top:0}}.bad td:nth-child(2){{color:#bd2424;font-weight:700}}.warn td:nth-child(2){{color:#b26a00;font-weight:700}}.note td:nth-child(2){{color:#475569}}.badge{{border-radius:5px;background:#edf2ff;padding:3px 9px}}small{{color:#64748b}}code{{background:#f1f5f9;padding:2px}}</style></head><body>
<h1>论文合集 · Word 格式自动初筛报告</h1><p>本报告在本地生成。<b>检测结果不等于最终审稿结论</b>；需要人工审查科研内容、真实排版、文献来源和图表清晰度。</p>
<div class="frame"><h2>文件汇总</h2><table><thead><tr>{''.join('<th>'+x+'</th>' for x in ['文件','错误数','待复核数','提示数','初筛状态'])}</tr></thead><tbody>{ss}</tbody></table></div>
<div class="frame"><h2>具体问题</h2><table><thead><tr>{table_head}</tr></thead><tbody>{''.join(table_rows) if table_rows else '<tr><td colspan="7">未检测到明显问题。注意：不代表全部格式正确。</td></tr>'}</tbody></table></div>
<p><small>单击浏览器的 Ctrl+F 可查找文件名。建议用 Excel 打开随附 CSV 并筛选“等级”。</small></p></body></html>'''
    (folder/'审校报告.html').write_text(doc,encoding='utf-8')


def audit_dir(input_dir, output_dir=None):
    input_dir=Path(input_dir).expanduser().resolve()
    if not input_dir.is_dir(): raise ValueError('输入路径必须是存放 .docx 的文件夹')
    files=sorted(f for f in input_dir.glob('*.docx') if not f.name.startswith('~$'))
    if not files: raise ValueError(f'该目录下找不到 .docx 文件：{input_dir}')
    parent=Path(output_dir).expanduser().resolve() if output_dir else input_dir.parent
    parent.mkdir(parents=True,exist_ok=True)
    out=parent/('审校结果_'+dt.datetime.now().strftime('%Y%m%d_%H%M%S'))
    out.mkdir(exist_ok=True)
    detail=[]; summary=[]
    for path in files:
        auditor=Auditor(path.name)
        detail.extend(auditor.check_document(path))
        count=Counter(r[1] for r in auditor.rows)
        status='有明确异常' if count['错误'] else '需要人工复核' if count['待复核'] else '未发现明显异常（仍需人工复核）'
        summary.append([path.name,count['错误'],count['待复核'],count['提示'],status])
    safe_write_csv(out/'稿件汇总.csv',summary,['文件','错误数','待复核数','提示数','初筛状态'])
    safe_write_csv(out/'问题明细.csv',detail,COLUMNS)
    make_html(out,summary,detail)
    (out/'审校元数据.json').write_text(json.dumps({'version':'1.1','created_at':dt.datetime.now().isoformat(),
         'count_files':len(files),'count_errors':sum(r[1]=='错误' for r in detail),
         'count_review':sum(r[1]=='待复核' for r in detail)},ensure_ascii=False,indent=2),encoding='utf-8')
    return out,summary


def gui():
    try:
        import tkinter as tk
        from tkinter import filedialog,messagebox
        root=tk.Tk(); root.withdraw()
        src=filedialog.askdirectory(title='选择存放学生 .docx 稿件的文件夹')
        if not src: return
        dst=filedialog.askdirectory(title='选择审校报告保存的位置（取消则保存至稿件文件夹的上级目录）')
        try:
            out,summary=audit_dir(src,dst or None)
        except Exception as e:
            messagebox.showerror('检测失败',str(e)); return
        messagebox.showinfo('检测完成',f'已处理 {len(summary)} 份稿件。\n\n输出目录：\n{out}\n\n请打开「审校报告.html」查看结果。')
    except ImportError:
        print('当前 Python 没有 tkinter，请使用命令行模式 --input 文件夹',file=sys.stderr)
        raise SystemExit(2)
    finally:
        try: root.destroy()
        except Exception: pass


def main():
    ap=argparse.ArgumentParser(description='论文合集 Word 格式批量初筛器；Python 3.9+，零第三方依赖')
    ap.add_argument('--input',help='含多个 .docx 的文件夹')
    ap.add_argument('--output',help='报告存放的上级目录（可选）')
    ap.add_argument('--gui',action='store_true',help='打开图形界面')
    args=ap.parse_args()
    if args.gui or not args.input:
        gui(); return
    try:
        out,summary=audit_dir(args.input,args.output)
        print('处理文件数:',len(summary),'\n报告目录:',out)
        for row in summary: print(f'{row[0]} | 错误 {row[1]} | 待复核 {row[2]} | {row[4]}')
    except Exception as e:
        print('错误:',e,file=sys.stderr); sys.exit(2)


if __name__=='__main__': main()