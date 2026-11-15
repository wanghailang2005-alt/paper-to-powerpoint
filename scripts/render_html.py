"""Render agent-authored, evidence-grounded slide content; no paper content is bundled."""
from pathlib import Path
import argparse,html,json,shutil,re

def main():
    ap=argparse.ArgumentParser();ap.add_argument('--deck',type=Path,required=True);ap.add_argument('--out',type=Path,required=True)
    args=ap.parse_args();deck=json.loads(args.deck.read_text(encoding='utf-8-sig'));slides=deck['slides']
    if not slides:raise ValueError('slides must be nonempty')
    args.out.mkdir(parents=True,exist_ok=True);assets=args.out/'assets';assets.mkdir(exist_ok=True)
    shutil.copy2(Path(__file__).resolve().parents[1]/'assets/deck.css',assets/'deck.css')
    # Assets belong to the new deck; keep original bytes and names to preserve figure provenance.
    for src in deck.get('assets',[]):
        p=(args.deck.parent/src).resolve();target=assets/p.name
        if target.exists() and target.read_bytes()!=p.read_bytes():raise ValueError(f'Asset name collision: {p.name}')
        if p!=target.resolve():shutil.copy2(p,target)
    esc=html.escape;parts=[];metadata=[]
    for i,s in enumerate(slides,1):
        refs=s.get('references',deck.get('references',[]))
        if not refs:raise ValueError(f'Slide {i} needs a source reference')
        if any(not r.get('title') for r in refs):raise ValueError('Reference requires full title')
        if re.search(r'<(?:script|iframe|svg)\b|\bon\w+\s*=',s['body'],re.I):raise ValueError('Body accepts static HTML only')
        source=''.join('<div class="reference-title">'+esc(r['title'])+'</div><div class="reference-locator">'+esc('DOI: '+r['doi'] if r.get('doi') else 'DOI：未提供 / 未标注')+'　 |　'+esc(r.get('locator',s.get('locator','')))+'</div>'for r in refs)
        if len(refs)>1:source='<div class="reference-title">'+esc(refs[0]['title'])+'</div><div class="reference-locator">'+esc('DOI: '+refs[0]['doi'] if refs[0].get('doi')else'DOI：未提供 / 未标注')+'　 |　'+esc(refs[0].get('locator',s.get('locator','')))+'</div><div class="review-reference">'+esc('；'.join(r['title']+' · '+('DOI: '+r['doi'] if r.get('doi')else'未标DOI')for r in refs[1:]))+'</div>'
        parts.append(f'<div class="slide-wrap"><article class="slide v2-slide" data-number="{i}" data-title="{esc(s["title"],quote=True)}"><header class="slide-header"><h1>{i:02} {esc(s["title"])}</h1></header><div class="v2-claim"><div>{s.get("claim","")}</div></div><main class="slide-main">{s["body"]}</main><div class="takeaway">{esc(s.get("takeaway",""))}</div><footer class="slide-footer">{source}</footer><div class="slide-number">{i:02} / {len(slides):02}</div><div class="bottom-frame"></div></article></div>')
        metadata.append({'number':i,'title':s['title'],'note':s.get('note',''),'references':refs})
    options=''.join(f'<option value="{i}">{i:02} · {esc(s["title"])}</option>'for i,s in enumerate(slides,1))
    runtime='''const slides=[...document.querySelectorAll('.slide')],picker=document.querySelector('#slide-picker');let current=1;function go(n){current=Math.max(1,Math.min(slides.length,n));slides.forEach((s,i)=>s.classList.toggle('active',i===current-1));picker.value=current;scale();}function scale(){const st=document.querySelector('.stage'),factor=Math.min(st.clientWidth/1280,st.clientHeight/720);for(const s of slides){s.style.transform=`scale(${factor})`;s.parentElement.style.width=`${1280*factor}px`;s.parentElement.style.height=`${720*factor}px`;}}picker.onchange=()=>go(+picker.value);document.querySelector('#prev').onclick=()=>go(current-1);document.querySelector('#next').onclick=()=>go(current+1);document.addEventListener('keydown',e=>{if(e.key==='ArrowRight')go(current+1);if(e.key==='ArrowLeft')go(current-1);});if(new URLSearchParams(location.search).has('embed')){document.querySelector('.app-toolbar').style.display='none';document.querySelector('.stage').style.cssText='height:100vh;padding:0';}window.addEventListener('resize',scale);go(1);'''
    runtime=runtime.replace('factor=Math.min(st.clientWidth/1280,st.clientHeight/720)', 'cs=getComputedStyle(st),factor=Math.min((st.clientWidth-parseFloat(cs.paddingLeft)-parseFloat(cs.paddingRight))/1280,(st.clientHeight-parseFloat(cs.paddingTop)-parseFloat(cs.paddingBottom))/720)')
    doc='<!doctype html><html lang="zh-CN"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>'+esc(deck.get('title','文献汇报'))+'</title><link rel="stylesheet" href="assets/deck.css"><body><nav class="app-toolbar"><b class="brand">'+esc(deck.get('title','文献汇报'))+'</b><button id="prev">上一页</button><select id="slide-picker">'+options+'</select><button id="next">下一页</button></nav><div class="stage">'+''.join(parts)+'</div><script>'+runtime+'</script></body></html>'
    (args.out/'index.html').write_text(doc,encoding='utf-8');(args.out/'slides.json').write_text(json.dumps(metadata,ensure_ascii=False,indent=2),encoding='utf-8')
    print(f'Rendered {len(slides)} pages: {(args.out/"index.html").resolve()}')
if __name__=='__main__':main()
