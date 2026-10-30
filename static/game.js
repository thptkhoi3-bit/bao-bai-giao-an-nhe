/* Trò chơi tương tác cho lớp học: Rung chuông vàng, Ô chữ, Vòng quay may mắn, Ai nhanh ai đúng.
   Chạy hoàn toàn trên trình duyệt sau khi tải (không cần mạng), âm thanh WebAudio (bật/tắt), không blur. */
(function(){
const L='ABCD',E=s=>String(s==null?'':s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const norm=s=>String(s||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[đĐ]/g,'D').toUpperCase().replace(/[^A-Z0-9]/g,'');
let G=null,AC=null;
function snd(){try{return localStorage.getItem('gp_snd')!=='0'}catch(e){return true}}
function tone(f,d,type,vol,at){if(!snd())return;try{AC=AC||new (window.AudioContext||window.webkitAudioContext)();const t=AC.currentTime+(at||0),o=AC.createOscillator(),g=AC.createGain();o.type=type||'sine';o.frequency.value=f;g.gain.setValueAtTime(vol||.18,t);g.gain.exponentialRampToValueAtTime(.001,t+d);o.connect(g);g.connect(AC.destination);o.start(t);o.stop(t+d)}catch(e){}}
const S={tick:()=>tone(1200,.05,'square',.08),beep:()=>tone(880,.12,'square',.12),ding:()=>{tone(784,.18,'sine',.25);tone(1047,.3,'sine',.25,.15);tone(1319,.4,'sine',.22,.3)},
 buzz:()=>tone(140,.6,'sawtooth',.2),pop:()=>tone(660,.08,'triangle',.15),win:()=>[523,659,784,1047].forEach((f,i)=>tone(f,.25,'triangle',.2,i*.12))};
function prep(d){d=d||{};const qs=(d.questions||[]).filter(q=>q&&q.q&&q.options&&q.options.filter(Boolean).length>=2&&L.indexOf(q.a)>=0&&L.indexOf(q.a)<q.options.length);
 let oc=null;const o=d.ochu;if(o&&o.rows&&o.rows.length){const kw=norm(o.keyword);const rows=o.rows.map((r,i)=>{const a=norm(r.answer);return{clue:r.clue||'',a,pos:i<kw.length?a.indexOf(kw[i]):-1}});
  if(kw.length>=2&&rows.length===kw.length&&rows.every(r=>r.pos>=0&&r.a))oc={kwRaw:o.keyword||kw,kw,clue:o.clue||'',rows}}
 return{title:d.title||'Trò chơi',questions:qs,ochu:oc}}
const MODES=[['rcv','🔔 Rung chuông vàng'],['ochu','🔤 Ô chữ'],['vq','🎡 Vòng quay'],['an','⚡ Ai nhanh ai đúng']];
function open(data,opts){close();opts=opts||{};G={d:prep(data),names:(opts.names||[]).slice(0,80),standalone:!!opts.standalone,mode:opts.mode||'rcv',timer:null,
 rcv:{i:0,sel:-1,shown:false,dur:20,left:null},oc:{open:{},sel:0,kw:false},vq:{mode:(opts.names||[]).length?'n':'q',used:{},ang:0,spinning:false,res:null,shown:false,rm:true,sel:-1},
 an:{n:2,teams:['Đội 1','Đội 2','Đội 3','Đội 4'],sc:[0,0,0,0],i:0,shown:false,dur:15,left:null}};
 const el=document.createElement('div');el.id='gp';el.className='gp';el.setAttribute('role','dialog');el.setAttribute('aria-label','Trò chơi');
 el.innerHTML=`<div class="gptop"><div class="gptitle">🎮 ${E(G.d.title)}</div><nav class="gptabs">${MODES.map(([k,t])=>`<button class="gpb gptab" data-m="${k}">${t}</button>`).join('')}</nav>
 <div class="gptabs"><button class="gpb" id="gpsnd" title="Bật/tắt âm thanh (M)"></button><button class="gpb" id="gpfs" title="Toàn màn hình (F)">⛶ Toàn màn hình</button>${G.standalone?'<a class="gpb" href="/" style="text-decoration:none">↩ Trang chính</a>':'<button class="gpb red" id="gpx" title="Đóng (Esc)">✕ Đóng</button>'}</div></div>
 <div class="gpbody" id="gpbody"></div><div class="gphelp" id="gphelp"></div>`;
 document.body.appendChild(el);el.addEventListener('click',onClick);document.addEventListener('keydown',onKey);
 sndBtn();mode(G.mode);if(opts.fullscreen)fs();return el}
function close(){if(!G)return;stopT();document.removeEventListener('keydown',onKey);const el=document.getElementById('gp');if(el)el.remove();if(document.fullscreenElement)document.exitFullscreen().catch(()=>{});G=null}
function fs(){const el=document.getElementById('gp');if(!el)return;if(document.fullscreenElement)document.exitFullscreen().catch(()=>{});else if(el.requestFullscreen)el.requestFullscreen().catch(()=>{})}
function sndBtn(){const b=document.getElementById('gpsnd');if(b)b.textContent=snd()?'🔊 Âm thanh':'🔇 Tắt tiếng'}
function stopT(){if(G&&G.timer){clearInterval(G.timer);G.timer=null}}
function mode(m){stopT();G.mode=m;document.querySelectorAll('#gp .gptab').forEach(b=>b.classList.toggle('on',b.dataset.m===m));draw()}
const HELP={rcv:'Phím: 1–4 / A–D chọn · Space: xem đáp án · T: bấm giờ · ← →: câu trước/sau · F: toàn màn hình · M: âm thanh',
 ochu:'Phím: 1–9 chọn hàng · Enter: mở hàng · K: mở từ khoá · F: toàn màn hình · M: âm thanh',
 vq:'Phím: Space quay · Enter: xem đáp án · F: toàn màn hình · M: âm thanh',
 an:'Phím: 1–4 cộng 10 điểm cho đội · Shift+1–4 trừ 10 · Space: xem đáp án · T: bấm giờ · ← →: câu trước/sau'};
function draw(){const b=document.getElementById('gpbody');if(!b)return;document.getElementById('gphelp').textContent=HELP[G.mode];
 if(G.mode==='ochu')return drawOc(b);if(G.mode==='vq')return drawVq(b);
 if(!G.d.questions.length){b.innerHTML='<div class="gpempty">Chưa có câu hỏi trắc nghiệm nào (cần đủ 4 phương án và đáp án). Thầy cô soạn thêm ở mục “Câu hỏi trò chơi” nhé.</div>';return}
 if(G.mode==='rcv')drawRcv(b);else drawAn(b)}
function qHtml(q,n,tot,sel,shown){return `<div class="gpq" id="gpq"><span class="gpno">Câu ${n}/${tot}</span>${E(q.q)}</div><div class="gpopts">${q.options.map((o,j)=>o?`<button class="gpopt${sel===j?' sel':''}${shown?(L[j]===q.a?' right':(sel===j?' wrong':' dim')):''}" data-o="${j}"><b>${L[j]}</b><span>${E(o)}</span></button>`:'').join('')}</div>
 ${shown?`<div class="gpex" id="gpex">✅ Đáp án: ${q.a}. ${E(q.options[L.indexOf(q.a)])}${q.explain?' — '+E(q.explain):''}</div>`:''}`}
function tHtml(s){return `<span class="gptimer${s.left===0?' end':(s.left!=null&&s.left<=5?' low':'')}" id="gptimer">${s.left==null?s.dur:(s.left===0?'Hết giờ!':s.left)}</span>`}
function drawRcv(b){const s=G.rcv,qs=G.d.questions,q=qs[s.i];
 b.innerHTML=`${qHtml(q,s.i+1,qs.length,s.sel,s.shown)}<div class="gprow">${tHtml(s)}<button class="gpb" data-a="prev" ${s.i?'':'disabled'}>◀ Câu trước</button><button class="gpb warn" data-a="time" id="gptime">⏱ ${s.left==null||s.left===0?'Bắt đầu giờ':'Dừng giờ'}</button>
 <select class="gpsel" data-a="dur" aria-label="Thời gian">${[10,15,20,30,45,60].map(x=>`<option ${x===s.dur?'selected':''} value="${x}">${x} giây</option>`).join('')}</select>
 <button class="gpb ok" data-a="show" id="gpreveal">👁 Xem đáp án</button><button class="gpb pri" data-a="next" id="gpnext" ${s.i<qs.length-1?'':'disabled'}>Câu sau ▶</button></div>`}
function timer(s){if(G.timer){stopT();s.left=null;draw();return}s.left=s.dur;draw();G.timer=setInterval(()=>{if(!G){return}s.left--;if(s.left<=0){s.left=0;stopT();S.buzz()}else if(s.left<=5)S.beep();else S.tick();draw()},1000)}
function reveal(s,q){if(s.shown)return;stopT();s.shown=true;if(s.sel<0||L[s.sel]===q.a)S.ding();else S.buzz();draw()}
function go(s,d){const n=s.i+d;if(n<0||n>=G.d.questions.length)return;stopT();s.i=n;s.sel=-1;s.shown=false;s.left=null;S.pop();draw()}
/* ô chữ */
function drawOc(b){const o=G.d.ochu,s=G.oc;if(!o){b.innerHTML='<div class="gpempty">Chưa có ô chữ hợp lệ. Ở mục “Ô chữ”, thầy cô nhập từ khoá và mỗi hàng ngang phải chứa chữ cái tương ứng của từ khoá.</div>';return}
 const L0=Math.max(...o.rows.map(r=>r.pos)),W=Math.max(...o.rows.map(r=>L0-r.pos+r.a.length));const bw=b.clientWidth*.58,bh=b.clientHeight-40;
 const cs=Math.max(26,Math.min(64,Math.floor(Math.min((bw-60)/(W+1.6),(bh-20)/o.rows.length-6))));
 let g=`<div class="ocg" id="ocgrid" style="--cs:${cs}px">`;o.rows.forEach((r,i)=>{const off=L0-r.pos,op=s.open[i]||s.kw&&false;g+=`<div class="ocr${s.sel===i?' sel':''}" data-r="${i}"><button class="ocn" data-r="${i}" title="Hàng ${i+1}">${i+1}</button>`;
  for(let c=0;c<W;c++){const j=c-off;if(j<0||j>=r.a.length){g+='<span class="occ e"></span>';continue}const k=c===L0,show=s.open[i]||(k&&s.kw);g+=`<span class="occ${k?' k':''}${show?' o':''}" data-r="${i}">${show?r.a[j]:''}</span>`}
  g+='</div>'});g+='</div>';const r=o.rows[s.sel];
 b.innerHTML=`<div class="oc">${g}<div class="ocside"><div class="occlue" id="occlue"><small style="color:#ea580c;font-size:.6em">Hàng ${s.sel+1} · ${r.a.length} chữ cái</small><br>${E(r.clue)}</div>
 <div class="gprow" style="justify-content:flex-start"><button class="gpb pri" data-a="ocopen" id="ocopen">🔓 Mở hàng ${s.sel+1}</button><button class="gpb ok" data-a="ockw" id="ockwb">🔑 Mở từ khoá</button></div>
 ${s.kw?`<div class="ockw" id="ockw">${E(o.kwRaw.toUpperCase())}</div>`:`<div class="occlue" style="font-size:clamp(16px,1.5vw,24px);box-shadow:0 5px 0 #16a34a">🔑 Từ khoá hàng dọc (${o.kw.length} chữ cái)${o.clue?': '+E(o.clue):''}</div>`}</div></div>`}
/* vòng quay */
function vqItems(){const s=G.vq;const all=s.mode==='n'?G.names.map((n,i)=>({k:'n'+i,t:n})):G.d.questions.map((q,i)=>({k:'q'+i,t:'Câu '+(i+1),q}));return all.filter(x=>!s.rm||!s.used[x.k])}
const COL=['#ef4444','#f97316','#eab308','#22c55e','#06b6d4','#3b82f6','#8b5cf6','#ec4899'];
function drawVq(b){const s=G.vq,it=vqItems();const sz=Math.max(240,Math.min(b.clientHeight-30,b.clientWidth*.5));
 b.innerHTML=`<div class="vq"><div class="vqw"><canvas id="vqc" width="${sz*2}" height="${sz*2}" style="width:${sz}px;height:${sz}px"></canvas></div><div class="ocside">
 <div class="gprow" style="justify-content:flex-start"><button class="gpb ${s.mode==='q'?'on':''}" data-a="vqq">❓ Quay câu hỏi</button><button class="gpb ${s.mode==='n'?'on':''}" data-a="vqn">🧒 Quay tên học sinh</button></div>
 <label style="font-weight:700;font-size:16px"><input type="checkbox" data-a="vqrm" ${s.rm?'checked':''}> Bỏ mục đã quay</label>
 ${s.mode==='n'&&!G.names.length?'<textarea class="vqnames" id="vqnames" placeholder="Chưa có danh sách lớp (tab Nhận xét). Dán tên học sinh, mỗi dòng 1 em…"></textarea><button class="gpb" data-a="vqload">Dùng danh sách này</button>':''}
 <button class="gpb pri" data-a="spin" id="vqspin" style="font-size:clamp(20px,2vw,32px)" ${it.length&&!s.spinning?'':'disabled'}>🎡 QUAY!</button>
 <div class="vqres" id="vqres">${s.res?vqRes(s):(it.length?'Bấm QUAY (hoặc phím Space) nhé!':'Đã quay hết! <button class="gpb" data-a="vqreset">↺ Quay lại từ đầu</button>')}</div></div></div>`;
 wheel(it)}
function vqRes(s){const r=s.res;if(!r.q)return `🎉<span class="big" id="vqwin">${E(r.t)}</span>`;
 return `<span class="big" id="vqwin">${E(r.t)}</span>${E(r.q.q)}<div class="gpopts">${r.q.options.map((o,j)=>o?`<button class="gpopt${s.sel===j?' sel':''}${s.shown?(L[j]===r.q.a?' right':' dim'):''}" data-vo="${j}"><b>${L[j]}</b><span>${E(o)}</span></button>`:'').join('')}</div>
 ${s.shown?`<div class="gpex">✅ ${r.q.a}. ${E(r.q.options[L.indexOf(r.q.a)])}</div>`:'<div class="gprow"><button class="gpb ok" data-a="vqshow" id="vqshow">👁 Xem đáp án</button></div>'}`}
function wheel(it,ang){const c=document.getElementById('vqc');if(!c)return;const x=c.getContext('2d'),R=c.width/2,n=Math.max(1,it.length);ang=ang==null?G.vq.ang:ang;x.clearRect(0,0,c.width,c.height);
 for(let i=0;i<n;i++){const a0=ang+i*2*Math.PI/n-Math.PI/2,a1=a0+2*Math.PI/n;x.beginPath();x.moveTo(R,R);x.arc(R,R,R-6,a0,a1);x.closePath();x.fillStyle=it.length?COL[i%COL.length]:'#e5e7eb';x.fill();x.strokeStyle='#fff';x.lineWidth=4;x.stroke();
  if(!it.length)continue;x.save();x.translate(R,R);x.rotate((a0+a1)/2);x.fillStyle='#fff';x.font=`800 ${Math.max(14,Math.min(46,R*1.6/Math.max(6,n)))}px "Be Vietnam Pro",sans-serif`;x.textAlign='right';x.textBaseline='middle';
  let t=it[i].t;if(t.length>18)t=t.slice(0,17)+'…';x.fillText(t,R-24,0);x.restore()}
 x.beginPath();x.arc(R,R,R*.12,0,2*Math.PI);x.fillStyle='#fff';x.fill();x.lineWidth=8;x.strokeStyle='#f97316';x.stroke()}
function spin(){const s=G.vq,it=vqItems();if(s.spinning||!it.length)return;s.spinning=true;s.res=null;s.shown=false;s.sel=-1;draw();
 const n=it.length,pick=Math.floor(Math.random()*n),seg=2*Math.PI/n;const target=-(pick*seg+seg/2)+(Math.random()-.5)*seg*.6;const a0=s.ang;let end=target;while(end<a0+2*Math.PI*5)end+=2*Math.PI;
 const T=window.GP_FAST?60:4200,t0=performance.now();let lastSeg=-1;
 function f(t){if(!G||G.mode!=='vq'){if(G)s.spinning=false;return}const p=Math.min(1,(t-t0)/T),e=1-Math.pow(1-p,3),a=a0+(end-a0)*e;s.ang=a;wheel(it,a);const sg=Math.floor(((a%(2*Math.PI))+2*Math.PI)/seg);if(sg!==lastSeg){lastSeg=sg;S.tick()}
  if(p<1)requestAnimationFrame(f);else{s.ang=end%(2*Math.PI);s.spinning=false;s.res=it[pick];s.used[it[pick].k]=1;S.win();draw()}}
 requestAnimationFrame(f)}
/* ai nhanh ai đúng */
function drawAn(b){const s=G.an,qs=G.d.questions,q=qs[s.i];const mx=Math.max(...s.sc.slice(0,s.n));
 b.innerHTML=`<div class="an"><div>${qHtml(q,s.i+1,qs.length,-1,s.shown)}<div class="gprow">${tHtml(s)}<button class="gpb" data-a="prev" ${s.i?'':'disabled'}>◀</button><button class="gpb warn" data-a="time" id="gptime">⏱ ${s.left==null||s.left===0?'Bắt đầu giờ':'Dừng giờ'}</button>
 <select class="gpsel" data-a="dur" aria-label="Thời gian">${[10,15,20,30,45,60].map(x=>`<option ${x===s.dur?'selected':''} value="${x}">${x} giây</option>`).join('')}</select><button class="gpb ok" data-a="show" id="gpreveal">👁 Đáp án</button><button class="gpb pri" data-a="next" id="gpnext" ${s.i<qs.length-1?'':'disabled'}>▶</button></div></div>
 <div class="ants"><div class="gprow" style="margin:0;justify-content:flex-start"><b style="font-size:17px">Số đội:</b>${[2,3,4].map(k=>`<button class="gpb ${s.n===k?'on':''}" data-a="teams" data-n="${k}">${k}</button>`).join('')}<button class="gpb" data-a="reset" title="Về 0 điểm">↺</button></div>
 ${s.sc.slice(0,s.n).map((v,i)=>`<div class="ant${v===mx&&v>0?' lead':''}" data-t="${i}"><input value="${E(s.teams[i])}" data-team="${i}" aria-label="Tên đội ${i+1}" maxlength="24"><span class="sc" id="sc${i}">${v}</span><div class="pm"><button data-a="plus" data-t="${i}" aria-label="Cộng 10 điểm">+10</button><button data-a="minus" data-t="${i}" aria-label="Trừ 10 điểm">−10</button></div></div>`).join('')}</div></div>`}
function score(i,d){const s=G.an;if(i>=s.n)return;s.sc[i]+=d;d>0?S.ding():S.buzz();draw()}
function onClick(ev){const t=ev.target.closest('[data-a],[data-m],[data-o],[data-vo],[data-r],#gpx,#gpfs,#gpsnd');if(!t||!G)return;
 if(t.id==='gpx')return close();if(t.id==='gpfs')return fs();if(t.id==='gpsnd'){try{localStorage.setItem('gp_snd',snd()?'0':'1')}catch(e){}sndBtn();return S.pop()}
 if(t.dataset.m)return mode(t.dataset.m);const a=t.dataset.a,cur=G.mode==='an'?G.an:G.rcv,q=G.d.questions[cur.i];
 if(t.dataset.o!=null&&G.mode==='rcv'){if(!G.rcv.shown){G.rcv.sel=+t.dataset.o;S.pop();draw()}return}
 if(t.dataset.vo!=null){if(!G.vq.shown){G.vq.sel=+t.dataset.vo;S.pop();draw()}return}
 if(t.dataset.r!=null&&!a){G.oc.sel=+t.dataset.r;S.pop();return draw()}
 switch(a){case'prev':return go(cur,-1);case'next':return go(cur,1);case'time':return timer(cur);case'show':return reveal(cur,q);
  case'ocopen':G.oc.open[G.oc.sel]=1;S.ding();const nx=G.d.ochu.rows.findIndex((r,i)=>!G.oc.open[i]);return draw();
  case'ockw':G.oc.kw=true;S.win();return draw();
  case'spin':return spin();case'vqq':G.vq.mode='q';G.vq.res=null;return draw();case'vqn':G.vq.mode='n';G.vq.res=null;return draw();
  case'vqshow':G.vq.shown=true;S.ding();return draw();case'vqreset':G.vq.used={};G.vq.res=null;return draw();
  case'vqload':{const v=(document.getElementById('vqnames').value||'').split(/\n/).map(x=>x.replace(/^\s*\d+[.)]?\s*/,'').trim()).filter(Boolean);if(v.length){G.names=v.slice(0,80);G.vq.used={}}return draw()}
  case'teams':G.an.n=+t.dataset.n;return draw();case'reset':G.an.sc=[0,0,0,0];return draw();case'plus':return score(+t.dataset.t,10);case'minus':return score(+t.dataset.t,-10)}}
document.addEventListener('change',ev=>{if(!G)return;const t=ev.target;if(t.dataset&&t.dataset.a==='dur'){const s=G.mode==='an'?G.an:G.rcv;s.dur=+t.value;stopT();s.left=null;draw()}
 if(t.dataset&&t.dataset.a==='vqrm'){G.vq.rm=t.checked;draw()}if(t.dataset&&t.dataset.team!=null)G.an.teams[+t.dataset.team]=t.value.trim()||('Đội '+(+t.dataset.team+1))});
function onKey(e){if(!G)return;const tg=e.target;if(tg&&/INPUT|TEXTAREA|SELECT/.test(tg.tagName))return;const k=e.key,m=G.mode;
 if(k==='f'||k==='F'){e.preventDefault();return fs()}if(k==='m'||k==='M'){try{localStorage.setItem('gp_snd',snd()?'0':'1')}catch(_){}return sndBtn()}
 if(k==='Escape'&&!document.fullscreenElement&&!G.standalone)return close();
 if(m==='rcv'||m==='an'){const s=m==='an'?G.an:G.rcv,q=G.d.questions[s.i];if(!q)return;
  if(k===' '||k==='Enter'){e.preventDefault();return reveal(s,q)}if(k==='ArrowRight')return go(s,1);if(k==='ArrowLeft')return go(s,-1);if(k==='t'||k==='T')return timer(s);
  if(m==='rcv'){const j='1234'.indexOf(k)>=0?'1234'.indexOf(k):L.indexOf(k.toUpperCase());if(j>=0&&j<q.options.length&&!s.shown){s.sel=j;S.pop();draw()}}
  else{const c=e.code&&/^Digit[1-4]$/.test(e.code)?+e.code.slice(5)-1:'1234'.indexOf(k);if(c>=0)score(c,e.shiftKey?-10:10)}}
 else if(m==='ochu'&&G.d.ochu){const n=parseInt(k,10);if(n>=1&&n<=G.d.ochu.rows.length){G.oc.sel=n-1;S.pop();draw()}else if(k==='Enter'){G.oc.open[G.oc.sel]=1;S.ding();draw()}else if(k==='k'||k==='K'){G.oc.kw=true;S.win();draw()}}
 else if(m==='vq'){if(k===' '){e.preventDefault();spin()}else if(k==='Enter'&&G.vq.res&&G.vq.res.q){G.vq.shown=true;S.ding();draw()}}}
window.addEventListener('resize',()=>{if(G&&(G.mode==='ochu'||G.mode==='vq')&&!(G.vq&&G.vq.spinning))draw()});
window.GamePlayer={open,close,prep,norm,state:()=>G};
})();
