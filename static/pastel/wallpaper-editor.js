(()=>{
const image=document.getElementById('pageWallpaper'),frame=document.getElementById('wallpaperFrame'),toggle=document.getElementById('wallpaperToolButton'),controls=document.getElementById('wallpaperControls');
const sizeSlider=document.getElementById('wallpaperSize'),sizeValue=document.getElementById('wallpaperSizeValue');
const KEY='pastel-v24-offline-wallpaper'; let imgW=1671,imgH=941,state={scale:100,x:0,y:0},gesture=null;
try{const saved=JSON.parse(localStorage.getItem(KEY));if(saved&&Number.isFinite(saved.scale))state={scale:saved.scale,x:saved.x||0,y:saved.y||0}}catch{}
const fit=()=>Math.min(innerWidth/imgW,innerHeight/imgH);const clamp=(v,min,max)=>Math.max(min,Math.min(max,v));
function geometry(){const ratio=fit()*state.scale/100,w=imgW*ratio,h=imgH*ratio;return {w,h,x:(innerWidth-w)/2+state.x,y:(innerHeight-h)/2+state.y}}
function draw(){const g=geometry();Object.assign(image.style,{width:g.w+'px',height:g.h+'px',left:g.x+'px',top:g.y+'px'});Object.assign(frame.style,{width:g.w+'px',height:g.h+'px',left:g.x+'px',top:g.y+'px'});sizeSlider.value=clamp(state.scale,25,200);sizeValue.textContent=Math.round(state.scale)+'%';try{localStorage.setItem(KEY,JSON.stringify(state))}catch{}}
image.onload=()=>{imgW=image.naturalWidth;imgH=image.naturalHeight;draw()};draw();
toggle.onclick=()=>{const open=frame.hidden;frame.hidden=!open;controls.hidden=!open;toggle.setAttribute('aria-pressed',String(open));toggle.textContent=open?'✓ Hoàn tất chỉnh nền':'⛶ Chỉnh vị trí nền';draw()};
frame.addEventListener('pointerdown',e=>{if(e.pointerType==='mouse'&&e.button!==0)return;e.preventDefault();const g=geometry(),corner=e.target.closest('[data-handle]')?.dataset.handle||null;const sign={tl:[-1,-1],tr:[1,-1],bl:[-1,1],br:[1,1]}[corner];gesture={id:e.pointerId,startX:e.clientX,startY:e.clientY,s:{...state},g,corner,sign,ox:sign?(sign[0]>0?g.x:g.x+g.w):0,oy:sign?(sign[1]>0?g.y:g.y+g.h):0};frame.setPointerCapture(e.pointerId)});
frame.addEventListener('pointermove',e=>{if(!gesture||gesture.id!==e.pointerId)return;const a=gesture,dx=e.clientX-a.startX,dy=e.clientY-a.startY;if(!a.corner){state.x=a.s.x+dx;state.y=a.s.y+dy}else{const sx=a.sign[0],sy=a.sign[1],cx=a.g.x+(sx>0?a.g.w:0),cy=a.g.y+(sy>0?a.g.h:0),oldX=cx-a.ox,oldY=cy-a.oy,projection=((oldX+dx)*oldX+(oldY+dy)*oldY)/(oldX*oldX+oldY*oldY);state.scale=clamp(a.s.scale*projection,20,250);const w=imgW*fit()*state.scale/100,h=imgH*fit()*state.scale/100;state.x=a.ox-(sx>0?w:0)-(innerWidth-w)/2;state.y=a.oy-(sy>0?h:0)-(innerHeight-h)/2}draw()});
frame.addEventListener('pointerup',()=>gesture=null);frame.addEventListener('pointercancel',()=>gesture=null);
sizeSlider.oninput=()=>{state.scale=Number(sizeSlider.value);draw()};
document.getElementById('wallpaperReset').onclick=()=>{state={scale:100,x:0,y:0};draw()};
document.getElementById('wallpaperCopy').onclick=()=>{const s=JSON.stringify(state);if(navigator.clipboard?.writeText)navigator.clipboard.writeText(s).then(()=>alert('Đã sao chép thông số')).catch(()=>prompt('Sao chép thông số:',s));else prompt('Sao chép thông số:',s)};
addEventListener('resize',draw);
})();