(async()=>{
await document.fonts.ready;
const canvas=document.createElement('canvas'),ctx=canvas.getContext('2d');
const out=[];
const count=document.querySelectorAll('.slide').length;
if(!count||!document.querySelector('#slide-picker'))throw new Error('Expected .slide elements and #slide-picker');
for(let num=1;num<=count;num++){
 const picker=document.querySelector('#slide-picker');picker.value=String(num);picker.dispatchEvent(new Event('change'));
 await new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r)));
 const slide=document.querySelector('.slide.active');
 await Promise.all([...slide.querySelectorAll('img')].map(i=>i.decode()));
 const sr=slide.getBoundingClientRect(),width=slide.offsetWidth,height=slide.offsetHeight,scale=sr.width/width;
 const box=r=>({x:(r.left-sr.left)/scale,y:(r.top-sr.top)/scale,w:r.width/scale,h:r.height/scale});
 const paints=[],images=[],lines=[],warnings=[];const groups=new Map();let ownerId=0;
 for(const e of [slide,...slide.querySelectorAll('*')]){
  const cs=getComputedStyle(e),r=box(e.getBoundingClientRect());if(!r.w||!r.h||cs.display==='none')continue;
  if(e.classList.contains('bottom-frame')){paints.push({kind:'polygon',points:[[0,r.y],[width*.04,r.y],[width*.06,r.y+r.h*.65],[width*.94,r.y+r.h*.65],[width*.96,r.y],[width,r.y],[width,height],[0,height]],fill:cs.backgroundColor});continue;}
  for(const pseudo of ['::before','::after']){const pc=getComputedStyle(e,pseudo);if(!['none','normal'].includes(pc.content)&&pc.display!=='none')warnings.push(`Unsupported pseudo-element: ${e.tagName}.${e.className}${pseudo}`);}
  if(e!==slide&&cs.transform!=='none')warnings.push(`Unsupported transform: ${e.tagName}.${e.className}`);
  if(cs.opacity!=='1'||cs.visibility!=='visible')warnings.push(`Unsupported opacity/visibility: ${e.tagName}.${e.className}`);
  if(cs.backgroundImage!=='none'||cs.filter!=='none'||cs.boxShadow!=='none'||cs.clipPath!=='none'||parseFloat(cs.borderRadius)>0||cs.fontStyle==='italic'||cs.direction==='rtl')warnings.push(`Unsupported painted/text style: ${e.tagName}.${e.className}`);
  if(cs.backgroundColor!=='rgba(0, 0, 0, 0)' && cs.backgroundColor!=='transparent')paints.push({kind:'rect',...r,fill:cs.backgroundColor});
  for(const [side,x1,y1,x2,y2] of [['Top',r.x,r.y,r.x+r.w,r.y],['Bottom',r.x,r.y+r.h,r.x+r.w,r.y+r.h],['Left',r.x,r.y,r.x,r.y+r.h],['Right',r.x+r.w,r.y,r.x+r.w,r.y+r.h]]){
   const width=parseFloat(cs['border'+side+'Width']);if(width&&cs['border'+side+'Style']!=='none')paints.push({kind:'line',x1,y1,x2,y2,width,fill:cs['border'+side+'Color']});
   if(width&&!['solid','none'].includes(cs['border'+side+'Style']))warnings.push(`Unsupported border style: ${e.tagName}.${e.className}`);
  }
  if(e.tagName==='IMG'){
   if(cs.objectFit!=='contain'||cs.objectPosition!=='50% 50%')warnings.push(`Images require object-fit:contain and centered object-position: ${e.getAttribute('src')}`);
   const k=Math.min(r.w/e.naturalWidth,r.h/e.naturalHeight),w=e.naturalWidth*k,h=e.naturalHeight*k;
   images.push({src:e.getAttribute('src'),x:r.x+(r.w-w)/2,y:r.y+(r.h-h)/2,w,h,naturalWidth:e.naturalWidth,naturalHeight:e.naturalHeight});
  }
 }
 const walker=document.createTreeWalker(slide,NodeFilter.SHOW_TEXT);let node;
 while(node=walker.nextNode()){
  const el=node.parentElement,cs=getComputedStyle(el);if(cs.display==='none'||!el.getBoundingClientRect().width)continue;
  let owner=el;while(owner.parentElement!==slide && getComputedStyle(owner).display==='inline')owner=owner.parentElement;
  if(!groups.has(owner))groups.set(owner,{id:ownerId++,parts:[]});
  const size=parseFloat(cs.fontSize),family=cs.fontFamily.split(',')[0].replace(/["']/g,''),weight=Number(cs.fontWeight)>=600?700:400;
  ctx.font=`${weight} ${size}px "${family}"`;const metrics=ctx.measureText('Hg国');
  const ascent=metrics.fontBoundingBoxAscent,descent=metrics.fontBoundingBoxDescent;
  for(let i=0;i<node.length;i++){
   const range=document.createRange();range.setStart(node,i);range.setEnd(node,i+1);const cr=range.getBoundingClientRect();if(!cr.width||!cr.height)continue;
   const r=box(cr),baseline=r.y+(r.h-ascent-descent)/2+ascent;
   groups.get(owner).parts.push({text:node.textContent[i],x:r.x,y:baseline,w:r.w,h:r.h,top:r.y,font:family,size,weight,fill:cs.color});
  }
 }
 for(const [owner,{id,parts}] of groups){
  const rows=[];
  for(const part of parts){let line=rows.find(l=>Math.abs(l.y-part.y)<1.8);if(!line){line={y:part.y,chars:[]};rows.push(line);}line.chars.push(part);}
  for(const line of rows){
   line.chars.sort((a,b)=>a.x-b.x);const runs=[];
   for(const c of line.chars){const last=runs.at(-1);if(last&&last.font===c.font&&last.size===c.size&&last.weight===c.weight&&last.fill===c.fill&&Math.abs(last.x+last.w-c.x)<2){last.text+=c.text;last.w=c.x+c.w-last.x;}else runs.push({...c});}
   if(!runs.map(r=>r.text).join('').trim())continue;
   const minX=Math.min(...runs.map(r=>r.x)),maxX=Math.max(...runs.map(r=>r.x+r.w));
   lines.push({owner:id,tag:owner.tagName,x:minX,y:line.y,w:maxX-minX,top:Math.min(...runs.map(r=>r.top)),h:Math.max(...runs.map(r=>r.h)),runs});
  }
 }
 for(const li of slide.querySelectorAll('li')){
  const cs=getComputedStyle(li),list=li.parentElement,ls=getComputedStyle(list);if(ls.listStyleType==='none')continue;
  const r=box(li.getBoundingClientRect()),sz=parseFloat(cs.fontSize);ctx.font=`400 ${sz}px "Microsoft YaHei"`;const m=ctx.measureText('国');
  const own=groups.get(li),first=lines.find(l=>l.owner===own?.id),y=first?.y ?? r.y+(parseFloat(cs.lineHeight)-m.fontBoundingBoxAscent-m.fontBoundingBoxDescent)/2+m.fontBoundingBoxAscent;
  const text=list.tagName==='OL'?String([...list.children].indexOf(li)+1)+'.':'•';
  lines.push({owner:ownerId++,tag:'MARKER',x:r.x-23,y,w:20,top:y-sz,h:sz,runs:[{text,font:'Microsoft YaHei',size:sz,weight:400,fill:cs.color,x:r.x-23,w:20}]});
 }
 out.push({number:num,title:slide.dataset.title||slide.querySelector('h1')?.textContent||String(num),canvas:[width,height],paints,images,lines,warnings:[...new Set(warnings)]});
}
return out;
})()
