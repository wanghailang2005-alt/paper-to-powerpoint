"""Rebuild supported flat HTML layout as editable SVG with complete embedded figures."""
from pathlib import Path
import argparse,base64,hashlib,html,json,mimetypes,urllib.parse

def num(n):return f'{n:.3f}'.rstrip('0').rstrip('.')
def esc(s):return html.escape(str(s),quote=True)
def color(s):
    if s.startswith('rgb('):return '#'+''.join(f'{int(v.strip()):02X}'for v in s[4:-1].split(','))
    if s.startswith('rgba('):
        rgb=s[5:-1].split(',')
        if float(rgb[3])!=1:raise ValueError(f'Partial alpha must be manually represented: {s}')
        return '#'+''.join(f'{int(v.strip()):02X}'for v in rgb[:3])
    return s
def attrs(r):return f'font-family="{esc(r["font"])}" font-size="{round(r["size"])}" font-weight="{"bold" if r["weight"]==700 else "normal"}" fill="{color(r["fill"])}"'
def asset_path(base,src):
    parsed=urllib.parse.urlparse(src)
    if parsed.scheme=='file':return Path(urllib.parse.unquote(parsed.path.lstrip('/') if len(parsed.path)>2 and parsed.path[2]==':'else parsed.path))
    if parsed.scheme:raise ValueError('Use local source figures; remote assets must be resolved explicitly before conversion')
    return (base/urllib.parse.unquote(src)).resolve()
def main():
    ap=argparse.ArgumentParser();ap.add_argument('--layout',type=Path,required=True);ap.add_argument('--html-dir',type=Path,required=True);ap.add_argument('--out',type=Path,required=True);args=ap.parse_args()
    pages=json.loads(args.layout.read_text(encoding='utf-8-sig'));args.out.mkdir(parents=True,exist_ok=True);manifest=[]
    for page in pages:
        if page.get('warnings'):raise ValueError(f'Page {page["number"]}: inspect and fix unsupported styles: {page["warnings"]}')
        width,height=page.get('canvas',[1280,720])
        xml=['<?xml version="1.0" encoding="UTF-8"?>',f'<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" lang="zh-CN" width="{width}" height="{height}" viewBox="0 0 {width} {height}">',f'<title>{esc(page["title"])}</title>']
        for p in page['paints']:
            c=color(p['fill'])
            if p['kind']=='rect':xml.append(f'<rect x="{num(p["x"])}" y="{num(p["y"])}" width="{num(p["w"])}" height="{num(p["h"])}" fill="{c}"/>')
            elif p['kind']=='polygon':xml.append('<polygon points="'+' '.join(f'{num(x)},{num(y)}'for x,y in p['points'])+f'" fill="{c}"/>')
            else:xml.append(f'<line x1="{num(p["x1"])}" y1="{num(p["y1"])}" x2="{num(p["x2"])}" y2="{num(p["y2"])}" stroke="{c}" stroke-width="{num(p["width"])}"/>')
        image_records=[]
        for i,p in enumerate(page['images'],1):
            file=asset_path(args.html_dir,p['src']);raw=file.read_bytes();mime=mimetypes.guess_type(file.name)[0]
            if mime not in ['image/png','image/jpeg']:raise ValueError('Use complete PNG/JPEG source figures')
            b64=base64.b64encode(raw).decode();xml.append(f'<image id="figure-{i}" x="{num(p["x"])}" y="{num(p["y"])}" width="{num(p["w"])}" height="{num(p["h"])}" preserveAspectRatio="xMidYMid meet" xlink:href="data:{mime};base64,{b64}"/>')
            image_records.append({'source':str(file),'sha256':hashlib.sha256(raw).hexdigest()})
        objects=[];splits=[]
        for i,line in enumerate(page['lines'],1):
            groups=[]
            for run in line['runs']:
                key=(run['font'],run['size'],run['weight'])
                if groups and groups[-1][0]==key and abs(groups[-1][1][-1]['x']+groups[-1][1][-1]['w']-run['x'])<2:groups[-1][1].append(run)
                else:groups.append((key,[run]))
            if len(groups)>1:splits.append({'line':i,'parts':len(groups),'reason':'mixed weight/font or discontinuous math spacing'})
            for j,(_,runs)in enumerate(groups,1):
                text=''.join(r['text']for r in runs);body=esc(text)if len(runs)==1 else ''.join(f'<tspan fill="{color(r["fill"])}">{esc(r["text"])}</tspan>'for r in runs)
                ident=f'line-{i:03}-part-{j}';xml.append(f'<text id="{ident}" x="{num(runs[0]["x"])}" y="{num(line["y"])}" {attrs(runs[0])} text-anchor="start" xml:space="preserve">{body}</text>')
                objects.append({'id':ident,'text':text,'font':runs[0]['font'],'svg_size':round(runs[0]['size']),'weight':runs[0]['weight']})
        xml.append('</svg>');file=args.out/f'slide-{page["number"]:02}.svg';file.write_text('\n'.join(xml),encoding='utf-8')
        manifest.append({'page':page['number'],'file':file.name,'title':page['title'],'canvas':[width,height],'text_objects':objects,'images':image_records,'split_lines':splits,'sha256':hashlib.sha256(file.read_bytes()).hexdigest()})
    (args.out/'manifest.json').write_text(json.dumps(manifest,ensure_ascii=False,indent=2),encoding='utf-8')
    print(f'Built {len(manifest)} standalone editable SVG pages: {args.out.resolve()}')
if __name__=='__main__':main()
