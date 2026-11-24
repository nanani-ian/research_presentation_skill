"""Read-only PPTX/XLSX summary, style inventory and structural audit. No Office needed."""
import argparse, collections, json, posixpath, re, sys, zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

VERSION=3
NS = dict(a='http://schemas.openxmlformats.org/drawingml/2006/main',
          p='http://schemas.openxmlformats.org/presentationml/2006/main',
          s='http://schemas.openxmlformats.org/spreadsheetml/2006/main',
          c='http://schemas.openxmlformats.org/drawingml/2006/chart',
          r='http://schemas.openxmlformats.org/officeDocument/2006/relationships',
          x='http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing')
RID = '{'+NS['r']+'}id'

def xml(z, part):
    return ET.fromstring(z.read(part))

def rels(z, part):
    rp = posixpath.join(posixpath.dirname(part), '_rels', posixpath.basename(part)+'.rels')
    if rp not in z.namelist(): return {}
    result = {}
    for e in xml(z, rp):
        target = e.get('Target', '')
        external = e.get('TargetMode') == 'External'
        resolved = target if external else (target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join(posixpath.dirname(part), target)))
        result[e.get('Id')] = dict(target=resolved, external=external, kind=e.get('Type','').rsplit('/',1)[-1])
    return result

def top(counter, n=12):
    return [{'value':k,'count':v} for k,v in counter.most_common(n)]

def style(z, parts):
    colors, fonts, sizes = (collections.Counter() for _ in range(3))
    for part in parts:
        rt = xml(z, part)
        colors.update(e.get('val') for e in rt.findall('.//a:srgbClr',NS) if e.get('val'))
        fonts.update(e.get('typeface') for e in rt.findall('.//a:latin',NS)+rt.findall('.//a:ea',NS) if e.get('typeface'))
        sizes.update(int(e.get('sz'))/100 for e in rt.findall('.//a:rPr',NS) if e.get('sz'))
    themes=[]
    for n in z.namelist():
        if re.fullmatch(r'ppt/theme/theme\d+\.xml',n):
            rt=xml(z,n); cs=rt.find('.//a:clrScheme',NS)
            themes.append({'part':n,'colors':{e.tag.rsplit('}',1)[-1]:(list(e)[0].get('val') or list(e)[0].get('lastClr')) for e in cs} if cs is not None else {}})
    return dict(explicit_colors=top(colors), explicit_fonts=top(fonts), explicit_run_sizes_pt=top(sizes),themes=themes,
                note='Explicit formatting frequency is an observation, not resolved inherited styling or visual approval.')

def slide_selection(spec, count):
    if not spec: return set(range(1,count+1))
    result=set()
    for t in spec.split(','):
        t=t.strip()
        if '-' in t:
            a,b=map(int,t.split('-',1));result.update(range(a,b+1))
        else:result.add(int(t))
    if not result or min(result)<1 or max(result)>count:raise ValueError('Slide selection outside presentation range')
    return result

def pptx(z,args):
    pres=xml(z,'ppt/presentation.xml'); pr=rels(z,'ppt/presentation.xml')
    ids=pres.find('p:sldIdLst',NS)
    parts=[pr[e.get(RID)]['target'] for e in ids] if ids is not None else []
    selected=slide_selection(args.slides,len(parts)); sz=pres.find('p:sldSz',NS)
    report={'kind':'pptx','slide_count':len(parts),'size_emu':dict(sz.attrib) if sz is not None else {},'slides':[]}
    for i,part in enumerate(parts,1):
        if i not in selected:continue
        rt=xml(z,part); rr=rels(z,part); shapes=rt.findall('.//p:sp',NS)
        alltext=[''.join(t.text or '' for t in p.findall('.//a:t',NS)) for p in rt.findall('.//a:p',NS)]
        alltext=[t for t in alltext if t]
        row=dict(page=i,part=part,text_excerpt='\n'.join(alltext)[:args.max_text],text_shapes=len([s for s in shapes if s.findall('.//a:t',NS)]),
                 pictures=len(rt.findall('.//p:pic',NS)),tables=len(rt.findall('.//a:tbl',NS)),
                 charts=[rr[e.get(RID)]['target'] for e in rt.findall('.//c:chart',NS) if e.get(RID) in rr])
        if args.mode=='audit':
            text_rows=[]
            for s in shapes:
                paras=s.findall('p:txBody/a:p',NS)
                text='\n'.join(''.join(t.text or '' for t in p.findall('.//a:t',NS)) for p in paras).strip()
                if not text:continue
                nv=s.find('p:nvSpPr/p:cNvPr',NS)
                text_rows.append(dict(id=nv.get('id') if nv is not None else '',name=nv.get('name') if nv is not None else '',chars=len(text),paragraphs=len(paras),runs=sum(len(p.findall('a:r',NS)) for p in paras),excerpt=text[:80]))
            row['text_objects']=text_rows
            row['short_text_objects']=sum(t['chars']<=6 for t in text_rows)
            row['relationships']=[dict(id=k,**v) for k,v in rr.items() if v['external'] or v['kind'] in ('chart','oleObject','package')]
            row['note']='Short objects may be labels or diagram nodes; inspect visually. These counts do not prove editability or refresh.'
        report['slides'].append(row)
    chart_parts=sorted(n for n in z.namelist() if re.fullmatch(r'ppt/charts/chart\d+\.xml',n))
    report['chart_count']=len(chart_parts)
    if args.mode=='audit':report['charts']=[dict(part=n,refs=[e.text for e in xml(z,n).findall('.//c:f',NS)],relationships=[dict(id=k,**v) for k,v in rels(z,n).items()]) for n in chart_parts]
    if args.mode=='style':report['style']=style(z,[parts[i-1] for i in sorted(selected)])
    report['other_chart_parts']=sum('/charts/' in n and n.endswith('.xml') and '/_rels/' not in n and 'chartEx' in n for n in z.namelist())
    return report

def column_name(number):
    text=''
    while number:
        number,remainder=divmod(number-1,26);text=chr(65+remainder)+text
    return text

def drawing_objects(z,part,root):
    result=[];sheet_rels=rels(z,part)
    for drawing in root.findall('s:drawing',NS):
        link=sheet_rels.get(drawing.get(RID))
        if not link or link['external']:continue
        dp=link['target'];dr=rels(z,dp)
        for anchor in xml(z,dp):
            def address(tag):
                node=anchor.find('x:'+tag,NS)
                if node is None:return None
                col=node.find('x:col',NS);row=node.find('x:row',NS)
                return column_name(int(col.text)+1)+str(int(row.text)+1) if col is not None and row is not None else None
            base={'from':address('from'),'to':address('to')}
            extent=anchor.find('x:ext',NS)
            if extent is not None:base['extent_emu']=dict(extent.attrib)
            for chart in anchor.findall('.//c:chart',NS):
                target=dr.get(chart.get(RID))
                result.append(dict(kind='chart',part=target['target'] if target else None,**base))
            for pic in anchor.findall('.//x:pic',NS):result.append(dict(kind='image',**base))
    return result

def xlsx(z,args):
    wb=xml(z,'xl/workbook.xml');wr=rels(z,'xl/workbook.xml'); wanted=set(args.sheets or [])
    sheets=list(wb.find('s:sheets',NS)); names={s.get('name') for s in sheets}
    if wanted-names:raise ValueError('Unknown sheets: '+', '.join(sorted(wanted-names)))
    strings=[''.join(item.itertext()) for item in xml(z,'xl/sharedStrings.xml')] if args.mode=='layout' and 'xl/sharedStrings.xml' in z.namelist() else []
    shown_charts=set()
    report=dict(kind='xlsx',sheet_count=len(sheets),sheets=[],chart_count=0,media_count=sum(n.startswith('xl/media/') for n in z.namelist()),external_workbook_parts=[])
    for s in sheets:
        if wanted and s.get('name') not in wanted:continue
        part=wr[s.get(RID)]['target']; rt=xml(z,part); dim=rt.find('s:dimension',NS)
        cells=rt.findall('s:sheetData/s:row/s:c',NS); formulas=[c for c in cells if c.find('s:f',NS) is not None]
        row=dict(name=s.get('name'),state=s.get('state','visible'),dimension=dim.get('ref') if dim is not None else '',populated_cells=sum(c.find('s:v',NS) is not None or c.find('s:is',NS) is not None or c.find('s:f',NS) is not None for c in cells),formula_cells=len(formulas),merges=len(rt.findall('s:mergeCells/s:mergeCell',NS)),hyperlinks=len(rt.findall('s:hyperlinks/s:hyperlink',NS)))
        objects=drawing_objects(z,part,rt)
        row['charts_on_sheet']=sum(x['kind']=='chart' for x in objects)
        row['images_on_sheet']=sum(x['kind']=='image' for x in objects)
        shown_charts.update(x['part'] for x in objects if x['kind']=='chart' and x.get('part'))
        if args.mode=='layout':
            row['drawing_objects']=objects
            row['merge_ranges']=[e.get('ref') for e in rt.findall('s:mergeCells/s:mergeCell',NS)][:40]
            row['column_widths']=[dict(e.attrib) for e in rt.findall('s:cols/s:col',NS)][:20]
            pane=rt.find('s:sheetViews/s:sheetView/s:pane',NS);row['freeze_or_split_pane']=dict(pane.attrib) if pane is not None else {}
            labels=[]
            for c in cells:
                if c.get('t')=='s':
                    value=c.find('s:v',NS);text=strings[int(value.text)] if value is not None and value.text else ''
                elif c.get('t')=='inlineStr':text=''.join(c.find('s:is',NS).itertext())
                else:continue
                if text:labels.append({'cell':c.get('r'),'text':text[:min(args.max_text,160)]})
            row['opening_labels']=labels[:16]
            row['left_side_labels']=[x for x in labels if re.fullmatch(r'[A-C]\d+',x['cell'])][:32]
        if args.mode=='audit':
            row['formula_examples']=[dict(cell=c.get('r'),formula=c.find('s:f',NS).text) for c in formulas if c.find('s:f',NS).text][:5]
            row['cached_errors']=[dict(cell=c.get('r'),value=c.find('s:v',NS).text if c.find('s:v',NS) is not None else '') for c in cells if c.get('t')=='e'][:20]
        report['sheets'].append(row)
    charts=[]
    for n in z.namelist():
        if re.fullmatch(r'xl/(?:drawings/)?charts/chart\d+\.xml',n):
            rt=xml(z,n);pa=rt.find('c:chart/c:plotArea',NS)
            charts.append(dict(part=n,types=[e.tag.rsplit('}',1)[-1] for e in pa if e.tag.rsplit('}',1)[-1].endswith('Chart')] if pa is not None else [],title=''.join(e.text or '' for e in rt.findall('c:chart/c:title//a:t',NS)),refs=[e.text for e in rt.findall('.//c:f',NS)]))
        elif re.fullmatch(r'xl/externalLinks/externalLink\d+\.xml',n):report['external_workbook_parts'].append(dict(part=n,relationships=[dict(id=k,**v) for k,v in rels(z,n).items()]))
    report['chart_count']=len(charts)
    if args.mode!='summary':
        if wanted and args.mode=='layout':charts=[c for c in charts if c['part'] in shown_charts]
        elif wanted:charts=[c for c in charts if any((name+'!') in (f or '') or ("'"+name.replace("'","''")+"'!") in (f or '') for name in wanted for f in c['refs'])]
        report['charts']=charts
    if args.mode=='style':
        st=xml(z,'xl/styles.xml');report['cell_style_fonts']=[{e.tag.rsplit('}',1)[-1]:e.attrib for e in f} for f in st.find('s:fonts',NS)][:20]
    report['note']='Read-only stored structure. Cached errors are not newly recalculated; links are not refresh-tested. Sheet dimensions may include formatting-only cells.'
    return report

def main():
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('input',type=Path);p.add_argument('--mode',choices=['summary','style','audit','layout'],default='summary')
    p.add_argument('--slides',help='Presentation positions, for example 2,4-6')
    p.add_argument('--sheets',nargs='+');p.add_argument('--max-text',type=int,default=160)
    p.add_argument('--out',type=Path,help='Optional generated JSON report; matching unchanged inputs reuse it')
    args=p.parse_args(); source=args.input.resolve(strict=True)
    if source.suffix.lower() not in ('.pptx','.xlsx'):p.error('Only PPTX/XLSX supported')
    if args.mode=='layout' and source.suffix.lower()!='.xlsx':p.error('layout mode currently supports XLSX only')
    if args.max_text<0 or args.max_text>20000:p.error('max-text must be between 0 and 20000')
    if (source.suffix.lower()=='.xlsx' and args.slides) or (source.suffix.lower()=='.pptx' and args.sheets):p.error('Slide/sheet selection does not match file type')
    if args.out:
        args.out=args.out.resolve()
        if args.out==source or args.out.suffix.lower()!='.json':p.error('Output must be a separate JSON file')
        if not args.out.parent.is_dir():p.error('Output parent must already exist')
    stat=source.stat(); key=dict(version=VERSION,path=str(source),size=stat.st_size,mtime_ns=stat.st_mtime_ns,mode=args.mode,slides=args.slides,sheets=args.sheets,max_text=args.max_text)
    if args.out and args.out.exists():
        saved=json.loads(args.out.read_text(encoding='utf8'))
        if saved.get('source_signature')==key:
            print(json.dumps(dict(cached=True,report=str(args.out),summary=saved['summary']),ensure_ascii=False));return
        if saved.get('generator')!='research-presentation/inspect_ooxml':p.error('Refusing to overwrite a JSON file not created by this script')
    with zipfile.ZipFile(source,'r') as z:report=pptx(z,args) if source.suffix.lower()=='.pptx' else xlsx(z,args)
    end=source.stat()
    if end.st_size!=stat.st_size or end.st_mtime_ns!=stat.st_mtime_ns:raise RuntimeError('Source changed during inspection; discard results and retry')
    summary={k:report[k] for k in ('kind','slide_count','sheet_count','chart_count','media_count') if k in report}
    result=dict(generator='research-presentation/inspect_ooxml',source_signature=key,summary=summary,report=report)
    if args.out:
        args.out.write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf8')
        print(json.dumps(dict(cached=False,report=str(args.out),summary=summary),ensure_ascii=False))
    else:print(json.dumps(result,ensure_ascii=False,indent=2))

if __name__=='__main__':
    try:main()
    except (ValueError,KeyError,OSError,zipfile.BadZipFile,ET.ParseError) as e:
        print('Error: '+str(e),file=sys.stderr);sys.exit(2)
