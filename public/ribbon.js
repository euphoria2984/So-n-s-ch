/* ribbon.js: thanh công cụ kiểu Word + chỉnh kích thước bảng.
   Cách dùng: đặt file này cạnh index.html, thêm 1 dòng
   <script src="ribbon.js"></script>
   ngay TRƯỚC thẻ </body> (sau thẻ </script> của script chính). */
(()=>{
const $=s=>document.querySelector(s),$$=s=>[...document.querySelectorAll(s)];
const clamp=(v,a,b)=>Math.max(a,Math.min(b,v));

/* ================= CSS ================= */
let css=`
/* --- thanh tab kiểu Word --- */
#rtabs{flex:none;display:flex;align-items:flex-end;gap:2px;padding:4px 12px 0;border-bottom:1px solid var(--line);background:var(--bg);overflow-x:auto;scrollbar-width:none}
.rt{position:relative;border:0;background:none;padding:7px 16px;border-radius:5px 5px 0 0;color:var(--g1);cursor:pointer;font-weight:500;white-space:nowrap;transition:background-color .15s,color .15s}
.rt:hover{color:var(--ink)}
.rt:focus{outline:0}
.rt.on{color:var(--ink);box-shadow:inset 0 -2px 0 var(--acc)}
.rt.ctx{display:none}.rt.ctx.vis{display:block}
/* --- ribbon Trang chủ: nhóm có nhãn bên dưới --- */
#main[data-rt]:not([data-rt=home]) #tb{display:none}
#tb .gi{display:flex;align-items:center;gap:1px;flex-wrap:wrap}
/* --- ribbon ngữ cảnh: Bảng / Ảnh / Đối tượng --- */
#rbody{flex:none;display:none;padding:6px 10px 4px;background:var(--soft);border-bottom:1px solid var(--line)}
#rbody.on{display:block}
#rbody .ctxg{display:none!important;flex-wrap:wrap;align-items:stretch;gap:4px 0}
#rbody .ctxg.act:not([hidden]){display:flex!important}
.rg{display:flex;flex-direction:column;flex:0 0 auto;align-items:stretch;justify-content:space-between;padding:2px 12px;border-right:1px solid var(--line)}
.rg:last-child{border-right:0}
.rgi{display:flex;flex:1;flex-wrap:wrap;align-items:center;justify-content:center;gap:4px;min-height:34px}
.capc{display:flex;align-items:center;gap:4px}
#rbody .dd{min-width:130px}
#rbody .ctxg button{background:var(--bg);border:1px solid var(--line);border-radius:4px;padding:6px 10px;white-space:nowrap}
#rbody .ctxg button:hover{background:var(--hov)}
#rbody .ctxg button.on{background:var(--accs);border-color:var(--accb);color:var(--acc)}
#rbody .ctxg button:disabled{opacity:.4;pointer-events:none}
#rbody .ctxg .ddb{background:var(--bg);border:1px solid var(--line);padding:6px 10px}
#rbody .nl{display:flex;align-items:center;gap:6px;color:var(--g1);font-size:12.5px;white-space:nowrap}
#rbody .nl input{width:62px;padding:5px 7px;background:var(--bg);border:1px solid var(--line);border-radius:4px}
#rbody .capin{width:auto;flex:1 1 120px;min-width:80px;max-width:200px}
#rbody #gImg,#rbody #gFo{flex-wrap:wrap;align-items:center}
#rbody #gImg .rg,#rbody #gFo .rg{flex:1 1 auto;justify-content:center}
#rbody #gImg .rgi,#rbody #gFo .rgi{align-items:center;justify-content:center;gap:8px}
#rbody .capc{flex:none}
#rbody #gImg .dd,#rbody #gFo .dd{flex:none}
#rbody #gImg button.bgb{flex-direction:row;gap:6px;padding:6px 10px;min-width:0}
#rbody #gImg button.bgb .lucide{width:18px;height:18px}
#rbody .rg{padding:2px 8px}
#rbody #gImg .dd{min-width:92px}
#rbody #gImg button.bgb{min-width:48px}

#rbody .rgi{flex-wrap:nowrap}
#rbody .capc .dd{min-width:84px}
#rbody .ctxg button.del{color:#e5484d;padding:6px 8px;display:grid;place-items:center}
#rbody .ctxg button.bgb.del{display:flex;color:inherit}
#rbody .ctxg button.bgb.del .lucide{color:#e5484d}
#rbody .capc .nl input{width:50px}
#rbody input.capcl{width:30px;height:30px;padding:2px;border:1px solid var(--line);border-radius:4px;background:var(--bg);cursor:pointer}
#rbody #cellc{width:30px;height:30px;padding:2px;border:1px solid var(--line);border-radius:4px;background:var(--bg)}
#gFo[data-k=im] .g-tx,#gFo[data-k=tx] .g-im{display:none}
#ctxbar{display:none!important}
.hid{display:none!important}
/* nút lớn có menu thả xuống (kiểu Word) */
#rbody .ctxg button.bgb{display:flex;flex-direction:column;align-items:center;justify-content:center;gap:2px;min-width:58px;padding:5px 8px;font-size:12px}
#rbody .ctxg button.bgb .lucide{width:22px;height:22px}
#rbody .rg.cs .rgi{display:grid;grid-template-columns:auto auto;gap:4px 12px;justify-content:center}
#rbody .rg.cs button{display:flex;align-items:center;gap:6px}
#rbody .ctxg .agrid{display:grid;grid-template-columns:repeat(3,30px);gap:2px}
#rbody .ctxg .agrid button{width:30px;height:26px;padding:0;display:grid;place-items:center}
#rmenu{position:fixed;z-index:92;display:none;min-width:160px;padding:4px;background:var(--bg);border:1px solid var(--line);border-radius:6px;box-shadow:0 10px 28px #0003}
#rmenu button{display:block;width:100%;text-align:left;border:0;background:none;padding:7px 12px;border-radius:4px;cursor:pointer;color:var(--g1);white-space:nowrap}
#rmenu button:hover{background:var(--soft)}
#rmenu button.dis{opacity:.4;pointer-events:none}
.prev #rtabs,.prev #rbody{display:none!important}
/* --- ảnh trong ô bảng: giới hạn chiều cao để không đẩy cả bảng sang trang khác --- */
.body td img,.body th img{max-width:100%;max-height:calc(var(--bh)*.4);height:auto}
.body td p,.body th p{text-align:inherit}
/* --- khung biến đổi ảnh: tròn trắng, 8 điểm + xoay (giống Word) --- */
#tf{border:1px solid #8a8a8a}
#tf i{background:#fff;border:1.5px solid #6f6f6f}
#tf .n{left:50%;top:-7px;margin-left:-6px;cursor:ns-resize}
#tf .s{left:50%;bottom:-7px;margin-left:-6px;cursor:ns-resize}
#tf .e{right:-7px;top:50%;margin-top:-6px;cursor:ew-resize}
#tf .w{left:-7px;top:50%;margin-top:-6px;cursor:ew-resize}
/* --- khung kéo chỉnh kích thước bảng: cũng là hình tròn trắng --- */
#tbf{position:fixed;z-index:39;display:none;border:1px solid #8a8a8a;pointer-events:none}
#tbf i{position:absolute;width:12px;height:12px;background:#fff;border:1.5px solid #6f6f6f;border-radius:50%;pointer-events:auto;touch-action:none}
#tbf .e{right:-7px;top:50%;margin-top:-6px;cursor:ew-resize}
#tbf .s{left:50%;bottom:-7px;margin-left:-6px;cursor:ns-resize}
#tbf .se{right:-7px;bottom:-7px;cursor:nwse-resize}
#tbf .mv{left:-24px;top:-24px;width:18px;height:18px;cursor:move;display:grid;place-items:center;font:700 12px/1 sans-serif;color:#555;font-style:normal}
#tbf em{position:absolute;left:50%;top:100%;margin-top:12px;transform:translateX(-50%);background:var(--acc);color:var(--bg);font:600 11px 'Google Sans',Inter,sans-serif;padding:1px 7px;border-radius:3px;white-space:nowrap;display:none;font-style:normal}
/* --- thanh Ảnh / Đối tượng: đúng 1 dòng, dàn đều, mọi phần tử co lại chứ không đè nhau --- */
#rbody #gImg,#rbody #gFo{flex-wrap:nowrap;align-items:center;justify-content:space-between;gap:0}
#rbody #gImg .rg,#rbody #gFo .rg{flex:1 1 auto;min-width:0;padding:2px 6px;justify-content:center}
#rbody #gImg .rgi,#rbody #gFo .rgi{flex-wrap:nowrap;min-width:0;justify-content:space-evenly;gap:6px;align-items:center}
#rbody #gImg .dd,#rbody #gFo .dd{flex:1 1 70px;min-width:44px;max-width:150px}
#rbody #gImg .capc,#rbody #gFo .capc{flex:1 1 auto;min-width:0;gap:6px}
#rbody #gImg .capc .dd,#rbody #gFo .capc .dd{flex:1 1 70px;min-width:40px}
#rbody #gImg .capc .nl,#rbody #gFo .capc .nl{flex:none}
#rbody #gImg .capc .nl input,#rbody #gFo .capc .nl input{width:46px}
#rbody #gImg .capcl,#rbody #gFo .capcl{flex:none}
#rbody #gImg .capin,#rbody #gFo .capin{flex:1 1 90px;min-width:40px;max-width:none;width:auto}
#rbody #gImg button.bgb{flex:1 1 auto;min-width:0;max-width:150px}
#rbody #gImg button.bgb span{overflow:hidden;text-overflow:ellipsis;min-width:0}
#rbody #gImg button.del,#rbody #gFo button.del{flex:none}
/* --- thanh Ảnh: Bố cục + Kích thước sát nhau, mô tả rộng, Xóa sát phải --- */
#rbody #gImg .rg.sz{flex:0 1 auto}
#rbody #gImg .rg.sz .rgi{justify-content:flex-start;gap:6px}
#rbody #gImg .rg.sz .dd{flex:0 1 130px;min-width:44px;max-width:none}
#rbody #gImg .rg.sz button.bgb{flex:0 1 auto;max-width:none}
#rbody #gImg .rg.cap{flex:1 1 auto}
#rbody #gImg .rg.cap .capin{flex:1 1 200px;min-width:60px}
#rbody #gImg .rg.cap .capc{flex:0 1 auto}
#rbody #gImg .rg.cap .capc .dd{flex:0 1 150px;min-width:60px}
#rbody #gImg .rg:last-child{flex:0 0 auto;margin-left:auto;padding-right:2px}
/* --- Mục lục: bảng nổi bên phải như menu bên trái --- */
aside{position:fixed!important;right:0;top:0;bottom:0;z-index:63;width:304px!important;padding:22px 18px!important;border-left:1px solid var(--line)!important;background:var(--bg);overflow-x:hidden;overflow-y:auto;opacity:1!important;display:block!important;transform:none}
aside.hide{transform:translate3d(100%,0,0);visibility:hidden;box-shadow:none}
aside:not(.hide){box-shadow:-12px 0 40px #0003}
.ready aside{transition:transform .36s cubic-bezier(.22,.85,.3,1),box-shadow .36s,visibility 0s .36s}
.ready aside:not(.hide){transition:transform .36s cubic-bezier(.22,.85,.3,1),box-shadow .36s,visibility 0s}
aside .aclose{position:absolute;right:10px;top:12px;width:30px;height:30px;min-width:0;border:0;border-radius:4px;background:none;display:grid;place-items:center;cursor:pointer;color:var(--g1)}
aside .aclose:hover{background:var(--soft)}
#scrim2{position:fixed;inset:0;z-index:62;background:rgba(0,0,0,.28);opacity:0;pointer-events:none;transition:opacity .25s}
#scrim2.on{opacity:1;pointer-events:auto}
html[data-theme=dark] #scrim2{background:rgba(0,0,0,.45)}
@media print{#rtabs,#rbody,#tbf,#rmenu,aside,#scrim2{display:none!important}}
`;
/* độ rộng cột: mỗi bảng lưu --c1, --c2... (%) trong style của chính nó */
for(let i=1;i<=30;i++)css+=`.body table>tbody>tr>:nth-child(${i}){width:var(--c${i},auto)}`;
const stl=document.createElement('style');stl.id='ribbon-css';stl.textContent=css;document.head.appendChild(stl);

/* ================= Bỏ 2 nút +/- cỡ chữ ================= */
['#fsm','#fsp'].forEach(s=>{const b=$(s);b&&b.remove()});

/* ================= Dựng ribbon ================= */
const main=$('#main'),tb=$('#tb');
const rt=document.createElement('div');rt.id='rtabs';
rt.innerHTML='<button class="rt on" data-r="home">Trang chủ</button><button class="rt ctx" data-r="tbl">Bảng</button><button class="rt ctx" data-r="img">Ảnh</button><button class="rt ctx" data-r="fo">Đối tượng</button>';
const rb=document.createElement('div');rb.id='rbody';
tb.before(rt);tb.after(rb);main.dataset.rt='home';

/* nhóm của ribbon Trang chủ */
const wrap=(g,lb)=>{if(!g||g.dataset.rb)return;g.dataset.rb=1;
 const gi=document.createElement('div');gi.className='gi';while(g.firstChild)gi.appendChild(g.firstChild);
 g.append(gi)};
const G=s=>{const e=$(s);return e&&e.closest('.grp')};
const ins=G('#lnk'),mb=$('#mathbtn');if(ins&&mb)ins.appendChild(mb);
wrap(G('#blk'),'Kiểu & cỡ chữ');
wrap(G('#tb [data-c=bold]'),'Phông chữ');
wrap(G('#tb [data-c=insertUnorderedList]'),'Đoạn văn');
wrap(G('#tb [data-c=justifyLeft]'),'Căn chỉnh');
wrap(ins,'Chèn');
const pa=$('.pgact');if(pa){pa.classList.add('grp');wrap(pa,'Trang')}

/* nhóm của ribbon ngữ cảnh */
const T=$('#gTbl'),I=$('#gImg'),F=$('#gFo');
rb.append(T,I,F);
$$('#rbody .lb').forEach(x=>x.remove());
const mk=(root,lb,els,cls)=>{const g=document.createElement('div');g.className='rg'+(cls?' '+cls:'');
 const i=document.createElement('div');i.className='rgi';els.filter(Boolean).forEach(e=>i.appendChild(e));
 g.append(i);root.appendChild(g);return i};
const ddw=s=>{const e=$(s);return e&&(e.closest('.dd')||e)};

/* --- menu thả xuống dùng chung --- */
const RM=document.createElement('div');RM.id='rmenu';document.body.appendChild(RM);
let rmFor=null;
const rmClose=()=>{RM.style.display='none';rmFor=null};
function redDel(b){b.classList.add('del');b.title='Xóa';b.querySelector('span').textContent='▾';return b}
function rmOpen(btn,items){if(rmFor===btn){rmClose();return}
 RM.innerHTML='';
 items.forEach(it=>{const b=document.createElement('button');b.type='button';b.textContent=it.t;if(it.dis)b.classList.add('dis');
  b.onmousedown=e=>e.preventDefault();b.onclick=()=>{rmClose();if(!it.dis)it.f()};RM.appendChild(b)});
 const r=btn.getBoundingClientRect();RM.style.display='block';rmFor=btn;
 RM.style.left=Math.max(6,Math.min(r.left,innerWidth-RM.offsetWidth-6))+'px';RM.style.top=r.bottom+4+'px'}
document.addEventListener('mousedown',e=>{if(rmFor&&!e.target.closest('#rmenu')&&e.target.closest('button')!==rmFor)rmClose()},true);
addEventListener('resize',rmClose);$('.canvas').addEventListener('scroll',rmClose);
const mbtn=(host,icon,label,items)=>{const b=document.createElement('button');b.type='button';b.className='bgb';
 b.innerHTML=`<i data-lucide="${icon}"></i><span>${label} ▾</span>`;
 b.onmousedown=e=>e.preventDefault();b.onclick=()=>rmOpen(b,typeof items==='function'?items():items);host.appendChild(b);return b};

/* ===== Ribbon BẢNG (theo Table Layout của Word) ===== */
const tq=k=>T.querySelector('[data-t='+k+']');
const hid=document.createElement('div');hid.className='hid';T.appendChild(hid);
['rowA','rowB','rowD','colL','colR','colD','tblD'].forEach(k=>hid.appendChild(tq(k)));
hid.insertAdjacentHTML('beforeend','<button data-ta="left"></button><button data-ta="center"></button><button data-ta="right"></button>');
const clkT=k=>()=>{const b=hid.querySelector('[data-t='+k+']');b&&b.click()};
const clkA=k=>()=>{const b=hid.querySelector('[data-ta='+k+']');b&&b.click()};
const selCells=(a,b)=>{fcs();const r=document.createRange();r.setStart(a,0);r.setEnd(b,b.childNodes.length);const s=getSelection();s.removeAllRanges();s.addRange(r)};

{const h=mk(T,'Hàng & cột',[]);
 mbtn(h,'box-select','Chọn',[
  {t:'Chọn ô',f:()=>curCell&&selCells(curCell,curCell)},
  {t:'Chọn hàng',f:()=>{const r=curCell&&curCell.parentElement;r&&selCells(r.cells[0],r.cells[r.cells.length-1])}},
  {t:'Chọn bảng',f:()=>{const t=curTbl();if(t){const c=t.querySelectorAll('td,th');selCells(c[0],c[c.length-1])}}}]);
 redDel(mbtn(h,'trash-2','Xóa',[{t:'Xóa cột',f:clkT('colD')},{t:'Xóa hàng',f:clkT('rowD')},{t:'Xóa bảng',f:clkT('tblD')}]));
 mbtn(h,'square-plus','Chèn',[{t:'Chèn hàng phía trên',f:clkT('rowA')},{t:'Chèn hàng phía dưới',f:clkT('rowB')},{t:'Chèn cột bên trái',f:clkT('colL')},{t:'Chèn cột bên phải',f:clkT('colR')}])}

mk(T,'Kích thước ô',[],'cs').innerHTML='<label class="nl">Cao <input type="number" id="tbh" min="3" step="1"> mm</label><button data-t2="eqr" title="Các hàng cao bằng nhau"><i data-lucide="rows-3"></i>Chia đều hàng</button><label class="nl">Rộng <input type="number" id="cw" min="3" step="1"> mm</label><button data-t2="eqc" title="Các cột rộng bằng nhau"><i data-lucide="columns-3"></i>Chia đều cột</button>';

/* căn chữ trong ô: lưới 3x3 */
const HA={L:'left',C:'center',R:'right'},VA={T:'top',M:'middle',B:'bottom'},HN={L:'trái',C:'giữa',R:'phải'},VN={T:'Trên',M:'Giữa',B:'Dưới'};
const aSvg=(h,v)=>{const ws=[10,6],x=w=>h==='L'?2:h==='R'?14-w:(16-w)/2,y0=v==='T'?2:v==='M'?5.5:9;
 return `<svg viewBox="0 0 16 16" width="16" height="16" fill="currentColor">${ws.map((w,i)=>`<rect x="${x(w)}" y="${y0+i*3.5}" width="${w}" height="2" rx="1"/>`).join('')}</svg>`};
const AG=mk(T,'Căn chỉnh',[]);
AG.innerHTML='<div class="agrid">'+['T','M','B'].map(v=>['L','C','R'].map(h=>`<button type="button" data-h="${HA[h]}" data-v="${VA[v]}" title="${VN[v]} ${HN[h]}">${aSvg(h,v)}</button>`).join('')).join('')+'</div>';
const cellsSel=()=>{const t=curTbl();if(!t)return[];const s=getSelection(),all=[...t.querySelectorAll('td,th')];
 if(s.rangeCount&&!s.isCollapsed&&pv.contains(s.anchorNode)){const r=s.getRangeAt(0),l=all.filter(c=>r.intersectsNode(c));if(l.length)return l}
 return curCell?[curCell]:[]};
AG.querySelectorAll('button').forEach(b=>{b.onmousedown=e=>e.preventDefault();
 b.onclick=()=>{const l=cellsSel();if(!l.length)return;l.forEach(c=>{c.style.textAlign=b.dataset.h;c.style.verticalAlign=b.dataset.v});later();save();ctx()}});

mk(T,'Kiểu bảng',[tq('head'),tq('border'),$('#cellc'),tq('nocolor')]);

{const h=mk(T,'Bảng',[]);
 h.innerHTML='<label class="nl">Rộng bảng <input type="number" id="tbw" min="10" max="100" step="1"> %</label>';
 const al=document.createElement('button');al.type='button';al.textContent='Căn bảng ▾';al.onmousedown=e=>e.preventDefault();
 al.onclick=()=>rmOpen(al,[{t:'Sát trái',f:clkA('left')},{t:'Giữa',f:clkA('center')},{t:'Sát phải',f:clkA('right')}]);
 const fit=document.createElement('button');fit.type='button';fit.dataset.t2='fit';fit.textContent='Đặt lại';fit.title='Rộng hết khung, cột đều, bỏ chiều cao hàng đã chỉnh';
 h.append(al,fit)}

/* ===== Ribbon ẢNH: gọn, 4 nhóm ===== */
const iq=k=>I.querySelector('[data-ia='+k+']');
const ihid=document.createElement('div');ihid.className='hid';I.appendChild(ihid);
['inline','left','center','right','float-l','float-r','free'].forEach(k=>ihid.appendChild(iq(k)));
const LN={inline:'Cùng dòng',left:'Trái',center:'Giữa',right:'Phải','float-l':'Chữ bao phải','float-r':'Chữ bao trái',free:'Tự do'};
const szH=mk(I,'Bố cục',[ddw('#imgw')]);szH.parentElement.classList.add('sz');
const layB=mbtn(szH,'wrap-text','Bố cục',()=>Object.keys(LN).map(k=>({t:LN[k],dis:iq(k).disabled,f:()=>iq(k).click()})));
szH.prepend(layB);
mk(I,'Mô tả ảnh',[I.querySelector('.capin')],'cap');
{const d=iq('del');d.className='del';d.title='Xóa ảnh';d.innerHTML='<i data-lucide="trash-2"></i>';mk(I,'Xóa',[d])}

const fq=k=>F.querySelector('[data-fa='+k+']');
mk(F,'Sắp xếp',[fq('front'),fq('back')]);
mk(F,'Hộp chữ',[F.querySelector('.txo')],'g-tx');
mk(F,'Ảnh tự do',[F.querySelector('.imx'),F.querySelector('.imo')],'g-im');
{const d=fq('del');d.className='del';d.title='Xóa';d.innerHTML='<i data-lucide="trash-2"></i>';mk(F,'Xóa',[d])}

I.querySelectorAll('[data-ia=free],[data-ia=float-l],[data-ia=float-r]').forEach(b=>b.dataset.t0=b.title||'');

/* ===== Phông & cỡ chữ của mô tả ảnh ===== */
const capEl=i=>{const fo=i.closest('.fo.im');if(fo)return fo.querySelector(':scope>.cap');const g=i.closest('.fg');return g&&g.querySelector('.cap')};
const capStyle=i=>{const c=capEl(i);if(!c)return;c.style.fontSize=i.dataset.capfs?i.dataset.capfs+'pt':'';c.style.fontFamily=i.dataset.capff||'';c.style.color=i.dataset.capcl||''};
{const _s=setCap;setCap=function(i){const r=_s.apply(this,arguments);capStyle(i);return r}}
const capc=[];let lastCap=null;
const capInit=host=>{
 host.insertAdjacentHTML('beforeend','<span class="capc"><select title="Phông chữ mô tả ảnh"><option value="">Phông mô tả</option></select><label class="nl" title="Cỡ chữ mô tả (pt)"><input type="number" min="6" max="72" step=".5" placeholder="pt"></label><input type="color" class="capcl" title="Màu chữ mô tả" value="#000000"></span>');
 const box=host.lastElementChild,sel=box.querySelector('select'),inp=box.querySelector('input[type=number]'),col=box.querySelector('.capcl');
 ddEnh(sel);capc.push({sel,inp,col});
 col.addEventListener('input',()=>{const i=lastCap;if(!i||!i.isConnected)return;i.dataset.capcl=col.value;capStyle(i);later();save()});
 sel.addEventListener('change',()=>{const i=lastCap;if(!i||!i.isConnected)return;
  if(sel.value)i.dataset.capff=sel.value;else delete i.dataset.capff;capStyle(i);later();save()});
 inp.addEventListener('change',()=>{const i=lastCap;if(!i||!i.isConnected)return;const v=parseFloat(inp.value);
  if(v>=6)i.dataset.capfs=Math.min(72,v);else delete i.dataset.capfs;capStyle(i);later();save()});
 box.addEventListener('pointerdown',()=>{lastCap=capImg()},true)};
const fillCap=()=>capc.forEach(({sel})=>{const v=sel.value;while(sel.options.length>1)sel.remove(1);
 [...$('#sf').options].forEach(o=>o.value&&sel.add(new Option(o.textContent,o.value)));sel.value=v});
capInit(I.querySelector('.capin').parentElement);
capInit(F.querySelector('.imx'));
fillCap();new MutationObserver(fillCap).observe($('#sf'),{childList:true});

/* ================= Chuyển tab ================= */
let cur='home',was={tbl:0,img:0,fo:0};
const panels={tbl:T,img:I,fo:F};
function setRT(n){cur=n;main.dataset.rt=n;
 rt.querySelectorAll('.rt').forEach(b=>b.classList.toggle('on',b.dataset.r===n));
 rb.classList.toggle('on',n!=='home');
 Object.entries(panels).forEach(([k,e])=>e.classList.toggle('act',k===n));
 rmClose();
 requestAnimationFrame(()=>{typeof tf==='function'&&tf();tbfPos()})}
rt.addEventListener('mousedown',e=>e.preventDefault());
rt.addEventListener('click',e=>{const b=e.target.closest('.rt');if(b)setRT(b.dataset.r)});
function rSync(){const v={tbl:!T.hidden,img:!I.hidden,fo:!F.hidden};
 rt.querySelectorAll('.ctx').forEach(b=>b.classList.toggle('vis',v[b.dataset.r]));
 if(cur!=='home'&&!v[cur])setRT('home');
 if(cur==='home'){if(v.img&&!was.img)setRT('img');else if(v.fo&&!was.fo)setRT('fo')}
 was=v}

/* ================= Cột / hàng của bảng ================= */
const nCols=t=>Math.max(1,...[...t.rows].map(r=>r.cells.length));
function setCols(t,w){const s=w.reduce((a,b)=>a+b,0)||1;
 w.forEach((x,i)=>t.style.setProperty('--c'+(i+1),(Math.round(x/s*1000)/10)+'%'));
 for(let i=w.length+1;i<=30;i++)t.style.removeProperty('--c'+i)}
function colsOf(t){const n=nCols(t),w=[];
 for(let i=1;i<=n;i++)w.push(parseFloat(t.style.getPropertyValue('--c'+i)));
 if(w.every(x=>x>0)&&!t.style.getPropertyValue('--c'+(n+1)))return w;
 const r0=t.rows[0],tw=t.getBoundingClientRect().width||1,m=[];
 for(let i=0;i<n;i++)m.push(r0&&r0.cells[i]?r0.cells[i].getBoundingClientRect().width/tw*100:100/n);
 setCols(t,m);const s=m.reduce((a,b)=>a+b,0)||1;return m.map(x=>x/s*100)}
const curTbl=()=>curCell&&curCell.isConnected&&pv.contains(curCell)?curCell.closest('table'):null;

let pend=null;
T.addEventListener('click',e=>{pend=null;const b=e.target.closest('[data-t]');
 if(!b||!curTbl())return;const a=b.dataset.t;if(!/^col[LRD]$/.test(a))return;
 const t=curTbl();pend={t,a,ix:curCell.cellIndex,w:colsOf(t)}},true);
T.addEventListener('mousedown',e=>{if(e.target.closest('button'))e.preventDefault()});
T.addEventListener('click',e=>{
 if(pend){const{t,a,ix,w}=pend;pend=null;
  if(t.isConnected&&t.rows.length){const n=nCols(t),x=w.slice(),avg=100/Math.max(1,w.length);
   if(a==='colL')x.splice(ix,0,avg);else if(a==='colR')x.splice(ix+1,0,avg);else x.splice(ix,1);
   if(x.length===n)setCols(t,x);else for(let i=1;i<=30;i++)t.style.removeProperty('--c'+i)}
  return}
 const b=e.target.closest('button');if(!b)return;const t=curTbl();if(!t)return;
 const a=b.dataset.t2,al=b.dataset.ta,k=sc(t.closest('.page'));
 if(a==='eqc')setCols(t,Array(nCols(t)).fill(1));
 else if(a==='eqr'){const h=[...t.rows].map(r=>r.getBoundingClientRect().height/k),m=Math.round(h.reduce((x,y)=>x+y,0)/h.length);
  [...t.rows].forEach(r=>r.style.height=m+'px')}
 else if(a==='fit'){t.style.removeProperty('width');t.style.removeProperty('margin-left');t.style.removeProperty('margin-right');for(let i=1;i<=30;i++)t.style.removeProperty('--c'+i);[...t.rows].forEach(r=>r.style.removeProperty('height'))}
 else if(al){t.style.marginLeft=al==='left'?'0':'auto';t.style.marginRight=al==='right'?'0':'auto'}
 else return;
 later();save();ctx()});
$('#tbw').addEventListener('change',e=>{const t=curTbl();if(!t)return;
 t.style.width=clamp(parseFloat(e.target.value)||100,10,100)+'%';later();save()});
$('#tbh').addEventListener('change',e=>{const t=curTbl();if(!t)return;
 const v=Math.max(3,parseFloat(e.target.value)||3);curCell.parentElement.style.height=Math.round(v*MM)+'px';later();save()});
$('#cw').addEventListener('change',e=>{const t=curTbl();if(!t)return;
 const k=sc(t.closest('.page')),W=t.getBoundingClientRect().width/k/MM,v=parseFloat(e.target.value),w=colsOf(t),ix=curCell.cellIndex;
 if(!(v>0)||w.length<2||!(W>0))return;
 const p=clamp(v/W*100,3,100-3*(w.length-1)),o=w.reduce((a,b,i)=>i===ix?a:a+b,0)||1,f=(100-p)/o;
 setCols(t,w.map((x,i)=>i===ix?p:x*f));later();save();ctx()});

function syncCont(){const ps=pgs();ps.forEach((p,i)=>{if(!i)return;
 const kids=b=>b?[...b.children].filter(x=>!x.classList.contains('exc')):[];
 const f=kids(p.querySelector('.body'))[0];if(!f||f.tagName!=='TABLE'||!f.dataset.cont)return;
 const l=kids(ps[i-1].querySelector('.body')).pop();
 if(l&&l.tagName==='TABLE'&&f.style.cssText!==l.style.cssText)f.style.cssText=l.style.cssText})}

/* ================= Kéo để chỉnh kích thước bảng ================= */
let drag=0,hot=null;
pv.addEventListener('pointermove',e=>{if(drag||!pv.isContentEditable)return;
 const c=e.target.closest&&e.target.closest('td,th');let h=null;
 if(c&&c.closest('.body')&&!e.buttons){const r=c.getBoundingClientRect(),t=c.closest('table'),ci=c.cellIndex,tr=c.parentElement;
  if(r.right-e.clientX<=4)h={t,type:'c',i:ci};
  else if(e.clientX-r.left<=4&&ci>0)h={t,type:'c',i:ci-1};
  else if(r.bottom-e.clientY<=4)h={t,type:'r',i:tr.rowIndex};
  else if(e.clientY-r.top<=4&&tr.rowIndex>0)h={t,type:'r',i:tr.rowIndex-1}}
 hot=h;pv.style.cursor=h?(h.type==='c'?'col-resize':'row-resize'):''});
pv.addEventListener('pointerleave',()=>{if(!drag){hot=null;pv.style.cursor=''}});
pv.addEventListener('mousedown',e=>{if(hot&&e.button===0)e.preventDefault()},true);
pv.addEventListener('pointerdown',e=>{if(!hot||e.button!==0||!hot.t.isConnected)return;
 const{t,type,i}=hot,k=sc(t.closest('.page')),x0=e.clientX,y0=e.clientY;let mv;drag=1;
 if(type==='c'){const W=colsOf(t),n=W.length,r0=t.getBoundingClientRect(),p0=parseFloat(t.style.width)||100;
  mv=ev=>{const dx=ev.clientX-x0;
   if(i>=n-1){t.style.width=Math.round(clamp(p0*(r0.width+dx)/r0.width,10,100)*10)/10+'%';return}
   const d=clamp(dx/r0.width*100,3-W[i],W[i+1]-3),w=W.slice();w[i]+=d;w[i+1]-=d;setCols(t,w)}}
 else{const tr=t.rows[i],h0=tr.getBoundingClientRect().height/k;
  mv=ev=>{tr.style.height=Math.max(12,Math.round(h0+(ev.clientY-y0)/k))+'px'}}
 const mv2=ev=>{mv(ev);tbfPos()};
 const up=()=>{removeEventListener('pointermove',mv2);removeEventListener('pointerup',up);drag=0;hot=null;pv.style.cursor='';later();save();ctx()};
 addEventListener('pointermove',mv2);addEventListener('pointerup',up)});

/* khung có núm tròn trắng (phải / dưới / góc) + núm ✥ ở góc trên trái để di chuyển bảng */
const TBF=document.createElement('div');TBF.id='tbf';
TBF.innerHTML='<i class="mv" data-h="mv" title="Kéo để di chuyển bảng">✥</i><i class="e" data-h="e"></i><i class="s" data-h="s"></i><i class="se" data-h="se"></i><em></em>';
document.body.appendChild(TBF);
const frameT=()=>{if(!pv.isContentEditable||!curCell||!curCell.isConnected||!pv.contains(curCell))return null;
 if(selFo&&pv.contains(selFo))return null;if(selImg&&pv.contains(selImg))return null;
 return curCell.closest('table')};
function tbfPos(){const t=frameT();if(!t){TBF.style.display='none';return}
 const r=t.getBoundingClientRect(),cv=$('.canvas').getBoundingClientRect();
 if(!r.width||r.bottom<cv.top||r.top>cv.bottom){TBF.style.display='none';return}
 Object.assign(TBF.style,{display:'block',left:r.left+'px',top:r.top+'px',width:r.width+'px',height:r.height+'px'})}

/* di chuyển tự do: kéo ngang để dời lề trái, kéo dọc/sang trang khác để đổi vị trí giữa các khối */
function moveTable(t,e){drag=1;
 const k=sc(t.closest('.page')),x0=e.clientX,ml0=parseFloat(t.style.marginLeft)||0;
 const ind=document.createElement('div');
 ind.style.cssText='position:fixed;z-index:95;height:3px;border-radius:2px;background:var(--acc);pointer-events:none;display:none';
 document.body.appendChild(ind);t.style.opacity='.55';let tg=null;
 const mv=ev=>{
  const pg=pgs().find(p=>{const b=p.getBoundingClientRect();return ev.clientX>=b.left&&ev.clientX<=b.right&&ev.clientY>=b.top&&ev.clientY<=b.bottom});
  tg=null;ind.style.display='none';
  if(pg){const body=pg.querySelector('.body'),br=body.getBoundingClientRect(),
    kids=[...body.children].filter(x=>!x.classList.contains('exc')&&x!==t);
   if(!kids.length)tg={body};
   else{let best=null,bd=1e9;
    kids.forEach(x=>{const r=x.getBoundingClientRect(),d=ev.clientY<r.top?r.top-ev.clientY:ev.clientY>r.bottom?ev.clientY-r.bottom:0;if(d<bd){bd=d;best=x}});
    const r=best.getBoundingClientRect();tg={el:best,after:ev.clientY>(r.top+r.bottom)/2}}
   const y=tg.el?(tg.after?tg.el.getBoundingClientRect().bottom:tg.el.getBoundingClientRect().top):br.top;
   Object.assign(ind.style,{display:'block',left:br.left+'px',width:br.width+'px',top:(y-1)+'px'})}
  const bd2=t.closest('.body');
  if(bd2){const bw=bd2.getBoundingClientRect().width/k,tw=t.getBoundingClientRect().width/k;
   t.style.marginLeft=Math.round(clamp(ml0+(ev.clientX-x0)/k,0,Math.max(0,bw-tw)))+'px';t.style.marginRight='auto'}
  tbfPos()};
 const up=()=>{removeEventListener('pointermove',mv);removeEventListener('pointerup',up);ind.remove();t.style.opacity='';
  if(tg){delete t.dataset.cont;
   if(tg.body)tg.body.appendChild(t);else tg.after?tg.el.after(t):tg.el.before(t);
   const sib=[...t.parentElement.children].filter(x=>!x.classList.contains('exc'));
   if(sib[sib.length-1]===t){const p=document.createElement('p');p.innerHTML='<br>';t.after(p)}}
  drag=0;run(0,null);save();ctx()};
 addEventListener('pointermove',mv);addEventListener('pointerup',up)}

TBF.addEventListener('pointerdown',e=>{const h=e.target.dataset&&e.target.dataset.h,t=frameT();if(!h||!t)return;
 e.preventDefault();e.stopPropagation();
 if(h==='mv'){moveTable(t,e);return}
 drag=1;
 const k=sc(t.closest('.page')),r0=t.getBoundingClientRect(),p0=parseFloat(t.style.width)||100,
  hs=[...t.rows].map(r=>r.getBoundingClientRect().height/k),lab=TBF.querySelector('em'),x0=e.clientX,y0=e.clientY;
 lab.style.display='block';
 const mv=ev=>{
  if(h.includes('e'))t.style.width=Math.round(clamp(p0*(r0.width+ev.clientX-x0)/r0.width,10,100)*10)/10+'%';
  if(h.includes('s')){const f=Math.max(.25,(r0.height+ev.clientY-y0)/r0.height);[...t.rows].forEach((r,i)=>r.style.height=Math.max(14,Math.round(hs[i]*f))+'px')}
  const r=t.getBoundingClientRect();lab.textContent=Math.round(r.width/k/MM)+' × '+Math.round(r.height/k/MM)+' mm';tbfPos()};
 const up=()=>{removeEventListener('pointermove',mv);removeEventListener('pointerup',up);lab.style.display='none';drag=0;later();save();ctx()};
 addEventListener('pointermove',mv);addEventListener('pointerup',up)});
$('.canvas').addEventListener('scroll',tbfPos);addEventListener('resize',tbfPos);

/* ================= Khung ảnh: thêm 4 núm cạnh (đủ 8 núm tròn như Word) ================= */
TF.insertAdjacentHTML('afterbegin','<i class="n" data-h="n"></i><i class="s" data-h="s"></i><i class="e" data-h="e"></i><i class="w" data-h="w"></i>');
TF.addEventListener('pointerdown',e=>{const h=e.target.dataset&&e.target.dataset.h;if(!/^[nsew]$/.test(h||''))return;
 const t=tgt();if(!t)return;e.preventDefault();e.stopPropagation();
 const g=geo(t),pg=t.closest('.page'),k=sc(pg),fo=t.classList.contains('fo'),a=g.a*Math.PI/180,cs=Math.cos(a),sn=Math.sin(a),lab=TF.querySelector('em');
 const ux=h==='e'?1:h==='w'?-1:0,uy=h==='s'?1:h==='n'?-1:0,len=ux?g.w:g.h;
 const o={x:g.cx+(-ux*g.w/2)*cs-(-uy*g.h/2)*sn,y:g.cy+(-ux*g.w/2)*sn+(-uy*g.h/2)*cs};
 lab.style.display='block';
 const mv=ev=>{
  const dx=ev.clientX-o.x,dy=ev.clientY-o.y,lx=dx*cs+dy*sn,ly=-dx*sn+dy*cs;
  let s=Math.max(.02,(ux?lx*ux:ly*uy)/len);if(g.w*s<24)s=24/g.w;
  const w=g.w*s,hh=g.h*s;
  if(fo){const pr=pg.getBoundingClientRect();t.style.width=Math.round(w/k*10)/10+'px';
   const G2=geo(t),w2=G2.w,h2=G2.h,cx=o.x+(ux*w2/2)*cs-(uy*h2/2)*sn,cy=o.y+(ux*w2/2)*sn+(uy*h2/2)*cs;
   t.style.left=Math.round((cx-w2/2-pr.left)/k*10)/10+'px';t.style.top=Math.round((cy-h2/2-pr.top)/k*10)/10+'px'}
  else{const bw=t.closest('.body').getBoundingClientRect().width/k;t.style.width=Math.round(Math.min(w/k,bw))+'px'}
  lab.textContent=Math.round(w/k/MM)+' × '+Math.round(hh/k/MM)+' mm';tf()};
 const up=()=>{removeEventListener('pointermove',mv);removeEventListener('pointerup',up);lab.style.display='none';run(0,null);save();thumbs();ctx()};
 addEventListener('pointermove',mv);addEventListener('pointerup',up)},true);

/* ================= Sửa lỗi: ảnh và mô tả bị tách nhau ================= */
/* (a) khi chia đoạn sang trang, không bao giờ cắt ngang giữa ảnh và dòng mô tả */
splitBlock=function(e,limit,rl=1e9){
 const tn=[],w=document.createTreeWalker(e,NodeFilter.SHOW_TEXT);let n;while(n=w.nextNode())tn.push(n);
 const total=tn.reduce((a,t)=>a+t.length,0);if(total<2)return null;
 const pos=k=>{for(const t of tn){if(k<=t.length)return[t,k];k-=t.length}};
 const fits=k=>{const[t,o]=pos(k),r=document.createRange();r.setStart(t,o);r.collapse(true);const rc=r.getClientRects(),q=rc.length?rc[rc.length-1]:r.getBoundingClientRect();return q.bottom<=limit&&q.left<rl};
 let lo=0,hi=total-1;while(lo<hi){const m=(lo+hi+1)>>1;fits(m)?lo=m:hi=m-1}
 const text=e.textContent;let k=lo;
 if(k>0&&text[k-1]!==' '){const j=text.lastIndexOf(' ',k-1);if(j>=0)k=j+1}
 if(k<=0)return null;
 const[t,o]=pos(k),r=document.createRange(),mm=t.parentElement.closest('.mth,.fg');
 mm?r.setStartBefore(mm):r.setStart(t,o);
 if(mm){const pre=document.createRange();pre.setStart(e,0);pre.setEnd(r.startContainer,r.startOffset);
  const f=pre.cloneContents();if(!f.textContent.trim()&&!f.querySelector('img'))return null}
 r.setEndAfter(e.lastChild);
 const rest=e.cloneNode(false);rest.removeAttribute('data-brk');rest.dataset.cont=1;rest.appendChild(r.extractContents());ch=1;return rest};
/* (b) tự vá: mô tả mồ côi bị xóa, ảnh thiếu mô tả được dựng lại đúng chữ đã đặt */
function heal(){
 pv.querySelectorAll('.cap').forEach(c=>{const p=c.parentElement;if(!p||!p.querySelector(':scope>img'))c.remove()});
 pv.querySelectorAll('.fg').forEach(g=>{if(!g.querySelector('img'))g.remove()});
 pv.querySelectorAll('img[data-cap]').forEach(i=>{const txt=i.dataset.cap,fo=i.closest('.fo.im'),g=i.closest('.fg'),
  c=fo?fo.querySelector(':scope>.cap'):g&&g.querySelector(':scope>.cap');
  if(!c||c.textContent!==txt)setCap(i,txt)})}

/* ================= Sửa lỗi: đổi chỗ 2 trang trong danh sách trang ================= */
/* thứ tự trang do người dùng tự quyết (kể cả bìa / credit / bìa cuối), không tự ép lại */
fixOrder=function(){};
reorder=function(a,b){const ps=pgs(),pa=ps[a],pb=ps[b];if(!pa||!pb||pa===pb)return;
 const m=document.createComment('');pa.replaceWith(m);pb.replaceWith(pa);m.replaceWith(pb);
 const L=pgs();new Set([a-1,a,a+1,b-1,b,b+1]).forEach(x=>{if(x>=0&&x<L.length)brkFirst(L[x])});
 run(0,null);goPage(b)};

/* ================= Ảnh chèn vào bảng nằm trong ô được chọn ================= */
{let sr=null;
 const go=()=>{const s=getSelection();
  sr=s.rangeCount&&pv.contains(s.anchorNode)?s.getRangeAt(0).cloneRange():null;
  const f=document.createElement('input');f.type='file';f.accept='image/*';
  f.onchange=()=>{const file=f.files[0];if(!file)return;const r=new FileReader();
   r.onload=()=>{pv.focus({preventScroll:true});const s2=getSelection();
    if(sr&&sr.startContainer.isConnected&&pv.contains(sr.startContainer)){s2.removeAllRanges();s2.addRange(sr)}
    document.execCommand('insertImage',false,r.result);later()};
   r.readAsDataURL(file)};
  f.click()};
 [$('#img'),$('#img2')].forEach(b=>{if(!b)return;b.onmousedown=e=>e.preventDefault();b.onclick=go})}

/* ================= Nối vào ctx() / run() ================= */
{const _c=ctx;
 ctx=function(){_c();
  const ed=pv.isContentEditable,foOn=!!(selFo&&pv.contains(selFo)),
   cell=ed&&!foOn&&curCell&&curCell.isConnected&&pv.contains(curCell),
   im=ed&&!foOn&&selImg&&pv.contains(selImg)?selImg:null;
  T.hidden=!cell;
  const inC=!!(im&&im.closest('td,th'));
  I.querySelectorAll('[data-ia=free],[data-ia=float-l],[data-ia=float-r]').forEach(b=>{
   b.disabled=inC;b.title=inC?'Không dùng được khi ảnh nằm trong ô bảng':(b.dataset.t0||'')});
  const on=I.querySelector('[data-ia].on');
  layB.querySelector('span').textContent=(on?LN[on.dataset.ia]:'Bố cục')+' ▾';
  if(cell){const t=curCell.closest('table'),w=$('#tbw'),h=$('#tbh'),cw=$('#cw'),k=sc(t.closest('.page'));
   if(document.activeElement!==w)w.value=Math.round(parseFloat(t.style.width)||100);
   if(document.activeElement!==h)h.value=Math.round(curCell.parentElement.getBoundingClientRect().height/k/MM*10)/10;
   if(document.activeElement!==cw)cw.value=Math.round(curCell.getBoundingClientRect().width/k/MM*10)/10;
   const ha=curCell.style.textAlign||'left',va=curCell.style.verticalAlign||'top';
   AG.querySelectorAll('button').forEach(b=>b.classList.toggle('on',b.dataset.h===ha&&b.dataset.v===va))}
  const ci=foOn&&selFo.classList.contains('im')?selFo.querySelector('img'):im;
  if(ci)capc.forEach(({sel,inp,col})=>{
   if(document.activeElement!==inp){const c=capEl(ci);
    inp.value=ci.dataset.capfs||(c?Math.round(parseFloat(getComputedStyle(c).fontSize)*.75*2)/2:'')}
   sel.value=ci.dataset.capff||'';
   if(document.activeElement!==col){const c=capEl(ci);let h=ci.dataset.capcl;
    if(!h&&c){const m=getComputedStyle(c).color.match(/\d+/g);if(m)h='#'+m.slice(0,3).map(n=>(+n).toString(16).padStart(2,'0')).join('')}
    col.value=h||'#000000'}});
  rSync();tbfPos()}}
/* màu số trang / đầu trang: luôn đọc được trên nền hiện tại */
const hexOf=v=>{if(!v)return'';if(v[0]==='#')return v.length===7?v:'';const m=v.match(/\d+/g);return m?'#'+m.slice(0,3).map(n=>(+n).toString(16).padStart(2,'0')).join(''):''};
function numColor(){const c=cfg(),root=hexOf($('#pgc').value)||'#ffffff';
 pgs().forEach(p=>{const bg=hexOf(p.style.backgroundColor)||root,auto=lum(bg)<.45?'#eee':'#222',
  col=c.pnca?auto:(Math.abs(lum(c.pnc)-lum(bg))<.3?auto:c.pnc);
  p.querySelectorAll(':scope>.num').forEach(e=>e.style.color=col);
  p.querySelectorAll(':scope>.hdr').forEach(e=>e.style.color=c.pnca?auto:'')})}
{const _a=apply;apply=function(){const r=_a.apply(this,arguments);try{numColor()}catch(e){}return r}}
{const _b=setBgColor;setBgColor=function(){const r=_b.apply(this,arguments);try{numColor()}catch(e){}return r}}
$('#bgreset').addEventListener('click',()=>numColor());
{const _r=run;run=function(...a){try{heal()}catch(e){}const r=_r.apply(this,a);try{syncCont();tbfPos();numColor()}catch(e){}return r}}

/* ===== Mục lục dạng menu nổi ===== */
const SCR2=document.createElement('div');SCR2.id='scrim2';document.body.appendChild(SCR2);
ASIDE.insertAdjacentHTML('afterbegin','<button type="button" class="aclose" title="Đóng (Esc)"><i data-lucide="x"></i></button>');
setAside=function(v){v=!!v;ASIDE.classList.toggle('hide',!v);SCR2.classList.toggle('on',v);$('#asidebtn').classList.toggle('on',v);
 try{localStorage.setItem('aside',v?1:0)}catch(e){}
 if(v){const o=$('#side').dataset.open;if(o)showPanel(o)}
 requestAnimationFrame(()=>{typeof tf==='function'&&tf();tbfPos()})};
{const _sp=showPanel;showPanel=function(n){const was=$('#side').dataset.open===n;
 if(!was&&!ASIDE.classList.contains('hide'))setAside(false);return _sp.apply(this,arguments)}}
SCR2.onclick=()=>setAside(false);
ASIDE.querySelector('.aclose').onclick=()=>setAside(false);
addEventListener('keydown',e=>{if(e.key==='Escape'&&!ASIDE.classList.contains('hide')&&!ddCur&&!$('#mdl.on'))setAside(false)});
setAside(false);

window.lucide&&lucide.createIcons();
setRT('home');ctx();later();
})();