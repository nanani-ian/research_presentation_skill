"""Discover preview candidates or pack explicitly adopted image versions. Never deletes or writes PPT."""
import argparse,json,math,re,sys
from pathlib import Path
from PIL import Image,ImageDraw,ImageFont,ImageOps

PATTERN=re.compile(r'^(?:页面|page)[ _-]?(\d+)[ _-]V(\d+)\.(png|jpg|jpeg|webp)$',re.I)

def json_out(path,data):
    if path.suffix.lower()!='.json' or not path.parent.is_dir():raise ValueError('JSON output requires an existing parent and .json suffix')
    if path.exists():raise ValueError('Output already exists; use a new report name')
    path.write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding='utf8')

def main():
    p=argparse.ArgumentParser(description=__doc__); sub=p.add_subparsers(dest='command',required=True)
    d=sub.add_parser('discover');d.add_argument('folder',type=Path);d.add_argument('--out',type=Path,required=True)
    c=sub.add_parser('pack');c.add_argument('manifest',type=Path);c.add_argument('--out',type=Path,required=True);c.add_argument('--columns',type=int,default=4);c.add_argument('--width',type=int,default=480);c.add_argument('--font',type=Path)
    args=p.parse_args()
    if args.command=='discover':
        folder=args.folder.resolve(strict=True)
        if not folder.is_dir():raise ValueError('Discovery input must be a directory')
        rows=[]
        for f in folder.iterdir():
            m=PATTERN.match(f.name)
            if f.is_file() and m:rows.append(dict(page=int(m[1]),version=int(m[2]),file=str(f),status='pending'))
        if not rows:raise ValueError('No matching page-version images found')
        rows.sort(key=lambda x:(x['page'],x['version']))
        json_out(args.out,dict(generator='research-presentation/preview_pack',note='Candidates only. Latest does not mean adopted. Mark explicit user-adopted versions; rejected alternatives may remain pending or rejected.',pages=rows))
        print(json.dumps(dict(candidates=len(rows),pages=len(set(x['page'] for x in rows)),manifest=str(args.out.resolve()),auto_adopted=0),ensure_ascii=False));return
    manifest=args.manifest.resolve(strict=True); data=json.loads(manifest.read_text(encoding='utf8'))
    rows=data.get('pages',[])
    if not rows:raise ValueError('Manifest has no pages')
    if any(x.get('status') not in ('pending','adopted','rejected') for x in rows):raise ValueError('Unknown adoption status')
    if any(not isinstance(x.get('page'),int) or isinstance(x.get('page'),bool) or x['page']<1 for x in rows):raise ValueError('Page numbers must be positive integers')
    expected=set(range(1,max(x['page'] for x in rows)+1));adopted=[x for x in rows if x.get('status')=='adopted'];counts={i:sum(x['page']==i for x in adopted) for i in expected}
    if any(n!=1 for n in counts.values()):raise ValueError('Each page needs exactly one explicitly adopted version: '+str({i:n for i,n in counts.items() if n!=1}))
    selected=sorted(adopted,key=lambda x:x['page']); sources=[]
    for row in selected:
        f=Path(row['file']);f=(manifest.parent/f).resolve(strict=True) if not f.is_absolute() else f.resolve(strict=True)
        if f.suffix.lower() not in ('.png','.jpg','.jpeg','.webp'):raise ValueError('Preview inputs must be raster images')
        sources.append(f)
    output=args.out.resolve()
    if output in sources or output==manifest or output.suffix.lower() not in ('.png','.jpg','.jpeg','.webp'):raise ValueError('Output must be a separate raster image')
    if output.exists() or not output.parent.is_dir():raise ValueError('Output must be new and its parent must already exist')
    if not 1<=args.columns<=8 or not 160<=args.width<=1600:raise ValueError('Columns or thumbnail width outside supported range')
    font=ImageFont.truetype(str(args.font),16) if args.font else ImageFont.load_default(size=16)
    thumb_h=round(args.width*9/16);pad=16;label_h=28;columns=min(args.columns,len(selected));rows_n=math.ceil(len(selected)/columns)
    canvas=Image.new('RGB',(columns*(args.width+pad)+pad,rows_n*(thumb_h+label_h+pad)+pad),'#eeeeee');draw=ImageDraw.Draw(canvas)
    for i,(row,source) in enumerate(zip(selected,sources)):
        with Image.open(source) as im:
            im=ImageOps.exif_transpose(im).convert('RGB');im=ImageOps.contain(im,(args.width,thumb_h))
            x=pad+(i%columns)*(args.width+pad);y=pad+(i//columns)*(thumb_h+label_h+pad)
            draw.rectangle((x,y,x+args.width,y+thumb_h),fill='white');canvas.paste(im,(x+(args.width-im.width)//2,y+(thumb_h-im.height)//2))
        draw.text((x,y+thumb_h+5),'Page '+str(row['page'])+'  V'+str(row.get('version','?')),font=font,fill='black')
    canvas.save(output)
    print(json.dumps(dict(output=str(output),pages=len(selected),size=canvas.size,note='Manifest adoption is supplied by the user/agent; the script cannot certify human approval.'),ensure_ascii=False))

if __name__=='__main__':
    try:main()
    except (ValueError,KeyError,OSError,TypeError) as e:
        print('Error: '+str(e),file=sys.stderr);sys.exit(2)
