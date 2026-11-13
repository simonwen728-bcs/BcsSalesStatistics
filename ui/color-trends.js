/* BCS_COLOR_TRENDS_1 — fixed native tables; no spreadsheet style replay. */
(function(){
'use strict';
const empty=()=>({qty:0,amount:0,rows:0});
const add=(a,q,v)=>{a.qty+=q;a.amount+=v;a.rows++;};
const key=v=>String(v||'').trim().replace(/\s+/g,' ').toUpperCase();
const number=v=>{const n=Number(String(v??'').replace(/[$,]/g,''));return Number.isFinite(n)?n:0;};
const validDate=v=>{const d=String(v||'').slice(0,10);return /^\d{4}-\d{2}-\d{2}$/.test(d)?d:'';};
function compute(sales,reps,opts){
 const group=new Map((reps?.s1||[]).map(r=>[key(r.code),key(r.group)]));
 const latest=sales.reduce((m,r)=>{const d=validDate(r.soldDate);return d>m?d:m;},'');
 const cutoff=opts.end && opts.end<latest?opts.end:latest;
 const colors=new Map(),months=new Map(),years=new Map();let total=empty();
 for(const r of sales){
  const d=validDate(r.soldDate),type=key(r.type).replace(/\s/g,'');
  if(!d||!cutoff||d>cutoff||!['CABINET','CABINETOTHERS'].includes(type)||key(r.rep1)==='IC'||group.get(key(r.rep1))==='IC')continue;
  const name=key(r.color)||'(空白)';if(name==='UNIVERSAL'&&!opts.includeUniversal)continue;
  let c=colors.get(name);if(!c){c={name,period:empty(),months:new Map(),years:new Map()};colors.set(name,c);}
  const q=number(r.qty),a=number(r.amount),m=d.slice(0,7),y=d.slice(0,4);
  for(const [map,k] of [[months,m],[years,y],[c.months,m],[c.years,y]]){if(!map.has(k))map.set(k,empty());add(map.get(k),q,a);}
  if((!opts.start||d>=opts.start)&&(!opts.end||d<=opts.end)){add(c.period,q,a);add(total,q,a);}
 }
 const ratio=(value,denom)=>denom && denom!==0?value/denom:null;
 const metrics=(v,denom)=>({qty:v.qty,amount:v.amount,qtyShare:ratio(v.qty,denom.qty),amountShare:ratio(v.amount,denom.amount),rows:v.rows});
 const list=[...colors.values()].filter(c=>c.period.rows).map(c=>({name:c.name,...metrics(c.period,total)})).sort((a,b)=>a.name==='UNIVERSAL'?1:b.name==='UNIVERSAL'?-1:b.qty-a.qty||a.name.localeCompare(b.name));
 const monthKeys=[];
 if(cutoff){const y=+cutoff.slice(0,4),m=+cutoff.slice(5,7);for(let back=11;back>=0;back--){const d=new Date(Date.UTC(y,m-1-back,1));monthKeys.push(d.toISOString().slice(0,7));}}
 return {latest,cutoff,total,list,colors,months,years,monthKeys,metrics};
}
function detail(data,name){
 const c=data.colors.get(name);if(!c)return {years:[],months:[]};
 const line=(k,map,denom)=>({period:k,...data.metrics(map.get(k)||empty(),denom.get(k)||empty())});
 return {years:[...data.years.keys()].sort().map(k=>line(k,c.years,data.years)),months:data.monthKeys.map(k=>line(k,c.months,data.months))};
}
const esc=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
let cfg,host,ctx=null,revision='',ctxToken,seq=0,data,selected='',chart,observer;
const $=id=>host.querySelector('#ct-'+id),zh=()=>cfg.lang()==='zh',text=(a,b)=>zh()?a:b;
const fmt=(n,kind)=>n==null?'—':kind==='pct'?(n*100).toFixed(2)+'%':Number(n).toLocaleString('en-US',{minimumFractionDigits:kind==='money'?2:0,maximumFractionDigits:2});
const cols=(first)=>[first,text('销售数量','Quantity'),text('数量占比','Quantity %'),text('销售金额','Sales Amount'),text('金额占比','Amount %')];
function table(headers,rows,foot){
 const head='<thead><tr>'+headers.map(h=>'<th>'+esc(h)+'</th>').join('')+'</tr></thead>';
 const cells=r=>r.map((v,i)=>'<td'+(i?' class="ct-n"':'')+'>'+esc(v)+'</td>').join('');
 return '<table><colgroup><col style="width:38%">'+Array(4).fill('<col style="width:15.5%">').join('')+'</colgroup>'+head+'<tbody>'+rows.map(r=>'<tr>'+cells(r)+'</tr>').join('')+'</tbody>'+(foot?'<tfoot><tr>'+cells(foot)+'</tr></tfoot>':'')+'</table>';
}
const row=r=>[r.period??r.name,fmt(r.qty),fmt(r.qtyShare,'pct'),fmt(r.amount,'money'),fmt(r.amountShare,'pct')];
function drawList(){
 const search=key($('search').value),list=data.list.filter(r=>r.name.includes(search));
 const total=data.metrics(data.total,data.total);
 $('list').innerHTML=table(cols(text('颜色','Color')),list.map(row),[text('合计','Total'),...row(total).slice(1)]);
 $('list').querySelectorAll('tbody tr').forEach((tr,i)=>{
  const name=list[i].name;tr.classList.toggle('ct-selected',name===selected);tr.tabIndex=0;tr.setAttribute('aria-selected',String(name===selected));
  tr.onclick=()=>{selected=name;drawList();drawDetail();};tr.onkeydown=e=>{if(e.key==='Enter'||e.key===' '){e.preventDefault();tr.click();}};
 });
 if(!list.length)$('list').querySelector('tbody').innerHTML='<tr><td colspan="5">'+text('没有符合条件的颜色','No matching colors')+'</td></tr>';
}
function drawDetail(){
 const d=detail(data,selected);$('selected').textContent=selected||text('请选择颜色','Select a color');
 $('years').innerHTML=table(cols(text('年份','Year')),d.years.map(row));
 $('months').innerHTML=table(cols(text('月份','Month')),d.months.map(row));
 $('chart-title').textContent=text('最近12个月占比趋势','Share trend — last 12 months')+' · '+selected;
 if(!window.echarts){$('chart').textContent=text('图表组件未加载，请刷新页面','Chart unavailable; refresh the page');return;}
 if(!chart)chart=window.echarts.init($('chart'));
 chart.setOption({animation:false,tooltip:{trigger:'axis',valueFormatter:v=>fmt(v,'pct')},legend:{top:4},grid:{left:15,right:20,top:40,bottom:10,containLabel:true},xAxis:{type:'category',boundaryGap:false,data:data.monthKeys,axisLabel:{rotate:35}},yAxis:{type:'value',axisLabel:{formatter:v=>(v*100).toFixed(0)+'%'}},series:[['qtyShare',text('数量占比','Quantity %'),'#2456a6'],['amountShare',text('金额占比','Amount %'),'#e6a23c']].map(([k,name,color])=>({name,type:'line',connectNulls:false,symbol:'circle',symbolSize:5,lineStyle:{width:2},itemStyle:{color},data:d.months.map(r=>r[k])}))},true);chart.resize();
}
function recalc(){
 if(!ctx)return;const start=$('start').value,end=$('end').value;
 if(start&&end&&start>end){$('state').textContent=text('开始日期不能晚于结束日期','Start date must precede end date');return;}
 const t=performance.now();data=compute(ctx.sales,ctx.reps,{start,end,includeUniversal:$('universal').checked});
 if(!data.list.some(r=>r.name===selected))selected=data.list[0]?.name||'';
 $('list-title').textContent=text('颜色列表','Color list')+' · '+(start||text('最早','First'))+' — '+(end||data.latest||'—');
 drawList();drawDetail();
 $('state').textContent=text('销售数据截止 ','Sales data through ')+(data.latest||'—')+' · '+text('右侧统计截止 ','Detail through ')+(data.cutoff||'—')+' · '+text('已更新 ','Updated ')+((performance.now()-t)/1000).toFixed(2)+'s';
}
async function refresh(){
 const request=++seq,token=cfg.token();$('refresh').disabled=true;$('state').textContent=text('正在检查最新数据…','Checking latest data…');
 try{
  const res=await cfg.normalize(await cfg.api('get_product_context','','',ctx&&ctxToken===token?revision:''));
  if(request!==seq)return;
  if(!res.unchanged){ctx=res;revision=res.revision;ctxToken=token;}
  if(!ctx)throw Error('No source data');recalc();
 }catch(e){if(request===seq){$('state').textContent=text('读取失败，请重试：','Load failed; retry: ')+(e.message||e);$('list').replaceChildren();$('years').replaceChildren();$('months').replaceChildren();if(chart)chart.clear();}}
 finally{if(request===seq)$('refresh').disabled=false;}
}
function quick(kind){
 const latest=ctx?.sales.reduce((m,r)=>{const d=validDate(r.soldDate);return d>m?d:m;},'')||'';
 if(!latest)return;const y=+latest.slice(0,4),m=+latest.slice(5,7),monthStart=latest.slice(0,7)+'-01';
 $('end').value=latest;$('start').value=kind==='last'?monthStart:kind==='year'?y+'-01-01':new Date(Date.UTC(y,m-12,1)).toISOString().slice(0,10);recalc();
}
function mount(){
 host=document.createElement('section');host.id='page-color-dashboard';host.className='page';host.style.display='none';
 host.innerHTML='<div class="ct-toolbar"><b>'+text('橱柜颜色动态','Cabinet Color Trends')+'</b><input id="ct-search" type="search" placeholder="'+text('搜索颜色…','Search colors…')+'"><button class="btn sm" id="ct-last">'+text('上月','Last month')+'</button><button class="btn sm" id="ct-year">'+text('今年','This year')+'</button><button class="btn sm" id="ct-ttm">TTM</button><button class="btn sm" id="ct-all">'+text('全部历史','All history')+'</button><label>'+text('开始','Start')+' <input id="ct-start" type="date"></label><label>'+text('结束','End')+' <input id="ct-end" type="date"></label><label><input id="ct-universal" type="checkbox"> '+text('包含 Universal','Include Universal')+'</label><button class="btn sm" id="ct-refresh">'+text('刷新数据','Refresh')+'</button></div><div class="ct-note">'+text('排除 IC；默认排除 Universal，不作分摊。占比分母为同一期间全部参与统计的橱柜颜色；退货按负数计入。','IC excluded. Universal excluded by default, without allocation. Shares use all included colors in the same period; returns deducted.')+'</div><div id="ct-state" role="status"></div><div class="ct-grid"><div class="ct-left"><b id="ct-list-title"></b><div id="ct-list" class="ct-table"></div></div><div class="ct-right"><b id="ct-selected"></b><div class="ct-analysis"><div id="ct-years" class="ct-table"></div><div id="ct-months" class="ct-table"></div></div><div class="ct-chart-panel"><div id="ct-chart-title"></div><div id="ct-chart"></div></div></div></div>';
 document.querySelector('main').appendChild(host);
 const style=document.createElement('style');style.textContent=`#page-color-dashboard{padding:10px 12px;gap:8px;min-height:0;flex:1;overflow:hidden}.ct-toolbar{display:flex;gap:8px;align-items:center;flex-wrap:wrap;background:#edf3fb;padding:8px;border-radius:6px}.ct-toolbar input{border:1px solid #cbd5e1;border-radius:4px;height:28px;padding:3px}.ct-toolbar label{white-space:nowrap}.ct-toolbar input[type=checkbox]{vertical-align:middle}.ct-note,#ct-state{font-size:12px;color:#52657d}.ct-grid{display:flex;gap:10px;flex:1;min-height:0}.ct-left{flex:0 0 35%;display:flex;flex-direction:column;gap:8px;min-height:0;min-width:0}#ct-list{flex:1;min-height:0}#ct-list-title{font-size:13px;color:#174779}.ct-right{display:flex;flex-direction:column;gap:8px;min-width:0;flex:1;min-height:0}#ct-selected{font-size:17px;color:#174779}.ct-analysis{display:flex;gap:8px;flex:0 1 425px;max-height:57%;min-height:180px}.ct-analysis>.ct-table{flex:1;min-width:0}.ct-table{overflow:auto;border:1px solid #cbd5e1;border-radius:6px;background:white}.ct-table table{width:100%;min-width:450px;table-layout:fixed;border-collapse:separate;border-spacing:0;font:13px Arial,sans-serif;color:#182330}.ct-table th,.ct-table td{height:30px;padding:4px 6px;border-right:1px solid #cbd5e1;border-bottom:1px solid #cbd5e1;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.ct-table th{position:sticky;top:0;background:#d9eaf7;text-align:center;z-index:2;font-size:12px}.ct-table tfoot td{position:sticky;bottom:0;background:#d9e6f3;font-weight:bold}.ct-table .ct-n{text-align:right}#ct-list tbody tr{cursor:pointer}#ct-list tbody tr:hover{background:#f0f5ff}#ct-list .ct-selected{background:#dbeafe}.ct-chart-panel{display:flex;flex:1;min-height:160px;flex-direction:column;background:white;border:1px solid #cbd5e1;border-radius:6px;padding:8px}#ct-chart{flex:1;min-height:140px}#ct-chart-title{font-weight:bold;color:#174779;font-size:13px}@media(max-width:1000px){.ct-grid{overflow:auto;flex-direction:column}.ct-left{flex:0 0 280px}.ct-right{min-height:650px}.ct-analysis{max-height:none;flex:0 0 400px}}`;
 document.head.appendChild(style);
 $('search').oninput=()=>data&&drawList();$('refresh').onclick=refresh;
 for(const k of ['start','end','universal'])$(k).onchange=recalc;
 for(const k of ['last','year','ttm'])$(k).onclick=()=>quick(k);
 $('all').onclick=()=>{$('start').value='';$('end').value='';recalc();};
 if(window.ResizeObserver){observer=new ResizeObserver(()=>{if(chart&&host.style.display!=='none')chart.resize();});observer.observe($('chart'));}
}
window.BCSColorTrends={compute,detail,show(options){cfg=options;if(!host)mount();host.style.display='flex';if(ctx&&ctxToken===cfg.token())recalc();return refresh();}};
})();
