/* BCS_DYNAMIC_BARS_1; BCS_COLOR_TRENDS_8_RECONCILE; BCS_COLOR_TRENDS_7_ALLOCATION; BCS_COLOR_TRENDS_6_TABLES; BCS_COLOR_TRENDS_5_ALIGN; BCS_COLOR_TRENDS_4_WIDTHS; BCS_COLOR_TRENDS_3_CHARTS; BCS_COLOR_TRENDS_2_SUMMARY; BCS_COLOR_TRENDS_1 — fixed native tables; no spreadsheet style replay. */
(function(){
'use strict';
const empty=()=>({qty:0,amount:0,rows:0});
const add=(a,q,v)=>{a.qty+=q;a.amount+=v;a.rows++;};
const key=v=>String(v||'').trim().replace(/\s+/g,' ').toUpperCase();
const number=v=>{const n=Number(String(v??'').replace(/[$,]/g,''));return Number.isFinite(n)?n:0;};
const validDate=v=>{const d=String(v||'').slice(0,10);return /^\d{4}-\d{2}-\d{2}$/.test(d)?d:'';};
// BCS_COLOR_ALLOCATION_MONTHLY_LEDGER_1: allocate once; every view sums the same dated records.
function allocateCabinetRows(rows,include){
 if(include)return rows;
 const groups=new Map(),out=rows.filter(r=>r.color!=='UNIVERSAL');
 const aggregate=list=>{const map=new Map();for(const r of list){let v=map.get(r.color);if(!v){v={qty:0,amount:0};map.set(r.color,v);}v.qty+=r.qty;v.amount+=r.amount;}return map;};
 const all=aggregate(out),items=new Map();
 for(const r of out){if(!items.has(r.item))items.set(r.item,[]);items.get(r.item).push(r);}
 for(const r of rows){const m=r.date.slice(0,7);if(!groups.has(m))groups.set(m,[]);groups.get(m).push(r);}
 const split=(weights,target)=>{
  const total=weights.reduce((s,w)=>s+w[1],0),sign=target<0?-1:1,n=Math.abs(target);
  if(!n)return weights.map(w=>[w[0],0]);
  if(!total)throw Error('Universal分摊缺少有效颜色参照');
  const result=weights.map(([c,w],i)=>{const x=n*w/total;return {c,i,n:Math.floor(x),f:x-Math.floor(x)};});
  let rest=n-result.reduce((s,r)=>s+r.n,0);const order=[...result].sort((a,b)=>b.f-a.f||a.i-b.i);
  for(let i=0;i<rest;i++)order[i%order.length].n++;
  return result.map(r=>[r.c,r.n*sign]);
 };
 for(const [month,list] of groups){
  const normal=list.filter(r=>r.color!=='UNIVERSAL'),global=aggregate(normal),perItem=new Map(),universal=new Map();
  for(const r of normal){if(!perItem.has(r.item))perItem.set(r.item,[]);perItem.get(r.item).push(r);}
  for(const r of list.filter(r=>r.color==='UNIVERSAL')){
   const k=r.item+'\u0000'+(r.qty<0?'-':'+');if(!universal.has(k))universal.set(k,[]);universal.get(k).push(r);
  }
  for(const us of universal.values()){
   const item=us[0].item;
   const candidates=[aggregate(perItem.get(item)||[]),global,aggregate(items.get(item)||[]),all];
   const weights=field=>{
    for(const map of candidates){const w=[...map].sort((a,b)=>a[0].localeCompare(b[0])).map(([c,v])=>[c,Math.abs(field==='amount'?(v.amount||v.qty):(v.qty||v.amount))]);if(w.some(x=>x[1]>0))return w;}
    throw Error('Universal分摊缺少颜色参照：'+month+' / '+item);
   };
   if(!us.some(r=>r.qty||r.amount))continue;
   // Positive and negative quantities stay separate, preserving both sales and returns.
   us.sort((a,b)=>a.date.localeCompare(b.date)||a.order.localeCompare(b.order)||a.qty-b.qty||a.amount-b.amount);
   const qw=weights('qty'),aw=weights('amount'),sign=us[0].qty<0?-1:1;
   const remaining=new Map(split(qw,us.reduce((s,r)=>s+Math.round(Math.abs(r.qty)),0)));
   for(const r of us){
    const qa=split([...remaining],Math.round(Math.abs(r.qty))),aa=split(aw,Math.round(r.amount*100));
    const allocated=new Map();
    for(const [c,q] of qa){remaining.set(c,remaining.get(c)-q);allocated.set(c,{qty:q*sign,amount:0});}
    for(const [c,cents] of aa){if(!allocated.has(c))allocated.set(c,{qty:0,amount:0});allocated.get(c).amount=cents/100;}
    for(const [color,v] of allocated)if(v.qty||v.amount)out.push({...r,color,qty:v.qty,amount:v.amount});
   }
  }
 }
 return out;
}

function compute(sales,reps,opts){
 const group=new Map((reps?.s1||[]).map(r=>[key(r.code),key(r.group)]));
 const latest=sales.reduce((m,r)=>{const d=validDate(r.soldDate);return d>m?d:m;},'');
 const cutoff=opts.end && opts.end<latest?opts.end:latest;
 const colors=new Map(),months=new Map(),years=new Map();let total=empty();
 const addMap=(map,k,q,a)=>{if(!map.has(k))map.set(k,empty());add(map.get(k),q,a);};
 const source=[];
 for(const r of sales){
  const d=validDate(r.soldDate),type=key(r.type).replace(/\s/g,'');
  if(!d||!['CABINET','CABINETOTHERS'].includes(type)||key(r.rep1)==='IC'||group.get(key(r.rep1))==='IC')continue;
  source.push({date:d,color:key(r.color)||'(空白)',item:key(r.itemNo)||'(NO ITEM)',qty:number(r.qty),amount:number(r.amount),order:String(r.invoiceNo||r.invoice||'')});
 }
 for(const r of allocateCabinetRows(source,opts.includeUniversal)){
  if(!cutoff||r.date>cutoff)continue;
  if(!colors.has(r.color))colors.set(r.color,{name:r.color,period:empty(),months:new Map(),years:new Map()});
  const c=colors.get(r.color),m=r.date.slice(0,7),y=r.date.slice(0,4);
  for(const [map,k] of [[months,m],[years,y],[c.months,m],[c.years,y]])addMap(map,k,r.qty,r.amount);
  if(!opts.start||r.date>=opts.start){add(c.period,r.qty,r.amount);add(total,r.qty,r.amount);}
 }
 const ratio=(value,denom)=>denom && denom!==0?value/denom:null;
 const metrics=(v,denom)=>({qty:v.qty,amount:v.amount,qtyShare:ratio(v.qty,denom.qty),amountShare:ratio(v.amount,denom.amount),rows:v.rows});
 const list=[...colors.values()].filter(c=>c.period.rows).map(c=>({name:c.name,...metrics(c.period,total)})).sort((a,b)=>a.name==='UNIVERSAL'?1:b.name==='UNIVERSAL'?-1:b.qty-a.qty||a.name.localeCompare(b.name));
 const monthKeys=[];
 if(cutoff){const y=+cutoff.slice(0,4),m=+cutoff.slice(5,7);for(let back=11;back>=0;back--){const d=new Date(Date.UTC(y,m-1-back,1));monthKeys.push(d.toISOString().slice(0,7));}}
 return {latest,cutoff,total,list,colors,months,years,monthKeys,metrics};
}
function detail(data,name){
 const c=data.colors.get(name);
 const yearKeys=[...data.years.keys()].sort();
 const line=(k,map,denom)=>({period:k,...data.metrics(map.get(k)||empty(),denom.get(k)||empty())});
 const aggregate=(keys,map,denom)=>{
  const value=empty(),base=empty();
  for(const k of keys){
   const v=map.get(k)||empty(),b=denom.get(k)||empty();
   value.qty+=v.qty;value.amount+=v.amount;value.rows+=v.rows;
   base.qty+=b.qty;base.amount+=b.amount;base.rows+=b.rows;
  }
  return data.metrics(value,base);
 };
 if(!c)return {years:[],months:[],yearTotal:data.metrics(empty(),empty()),monthTotal:data.metrics(empty(),empty())};
 return {years:yearKeys.map(k=>line(k,c.years,data.years)),months:data.monthKeys.map(k=>line(k,c.months,data.months)),
  yearTotal:aggregate(yearKeys,c.years,data.years),monthTotal:aggregate(data.monthKeys,c.months,data.months)};
}
const esc=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
let cfg,host,ctx=null,revision='',ctxToken,seq=0,data,selected='',charts={},observer,chartType='bar';
const $=id=>host.querySelector('#ct-'+id),zh=()=>cfg.lang()==='zh',text=(a,b)=>zh()?a:b;
const fmt=(n,kind)=>n==null?'—':kind==='pct'?(n*100).toFixed(2)+'%':Number(n).toLocaleString('en-US',{minimumFractionDigits:kind==='money'?2:0,maximumFractionDigits:2});
const cols=(first)=>[first,text('销售数量','Quantity'),text('数量占比','Quantity %'),text('销售金额','Sales Amount'),text('金额占比','Amount %')];
function table(headers,rows,foot,widths=[20,20,20,20,20]){
 const head='<thead><tr>'+headers.map(h=>'<th>'+esc(h)+'</th>').join('')+'</tr></thead>';
 const cells=r=>r.map((v,i)=>'<td title="'+esc(v)+'"'+(i===3?' class="ct-money"':i>0?' class="ct-center"':'')+'>'+esc(v)+'</td>').join('');
 const columns='<colgroup>'+widths.map(w=>'<col style="width:'+w+'%">').join('')+'</colgroup>';
 return '<div class="ct-scroll"><table>'+columns+head+'<tbody>'+rows.map(r=>'<tr>'+cells(r)+'</tr>').join('')+'</tbody></table></div>'+(foot?'<div class="ct-footer"><table>'+columns+'<tfoot><tr>'+cells(foot)+'</tr></tfoot></table></div>':'');
}
const row=r=>[r.period??r.name,fmt(r.qty),fmt(r.qtyShare,'pct'),fmt(r.amount,'money'),fmt(r.amountShare,'pct')];
function drawList(){
 const search=key($('search').value),list=data.list.filter(r=>r.name.includes(search));
 const total=data.metrics(data.total,data.total);
 $('list').innerHTML=table(cols(text('颜色','Color')),list.map(row),[text('合计','Total'),...row(total).slice(1)],[34,16.5,16.5,16.5,16.5]);
 $('list').querySelectorAll('tbody tr').forEach((tr,i)=>{
  const name=list[i].name;tr.classList.toggle('ct-selected',name===selected);tr.tabIndex=0;tr.setAttribute('aria-selected',String(name===selected));
  tr.onclick=()=>{selected=name;drawList();drawDetail();};tr.onkeydown=e=>{if(e.key==='Enter'||e.key===' '){e.preventDefault();tr.click();}};
 });
 if(!list.length)$('list').querySelector('tbody').innerHTML='<tr><td colspan="5">'+text('没有符合条件的颜色','No matching colors')+'</td></tr>';
}
function drawDetail(){
 const d=detail(data,selected);$('selected').textContent=selected||text('请选择颜色','Select a color');
 const pad=rows=>rows.concat(Array.from({length:Math.max(0,12-rows.length)},()=>['','','','','']));
 const totalRow=v=>[text('合计','Total'),...row(v).slice(1)];
 $('years').innerHTML=table(cols(text('年份','Year')),pad(d.years.map(row)),totalRow(d.yearTotal));
 $('months').innerHTML=table(cols(text('月份','Month')),pad(d.months.map(row)),totalRow(d.monthTotal));
 const plots=[
  {id:'qty',field:'qty',share:'qtyShare',name:text('销售数量','Quantity'),shareName:text('数量占比','Quantity %'),title:text('最近12个月销售数量','Quantity — last 12 months'),color:'#2456a6',shareColor:'#67a3e8',money:false},
  {id:'amount',field:'amount',share:'amountShare',name:text('销售金额','Sales Amount'),shareName:text('金额占比','Amount %'),title:text('最近12个月销售金额','Sales amount — last 12 months'),color:'#dc921f',shareColor:'#f0b65e',money:true}
 ];
 for(const p of plots){
  $('chart-'+p.id+'-title').textContent=p.title+' · '+selected;
  const el=$('chart-'+p.id);
  if(!window.echarts){el.textContent=text('图表组件未加载，请刷新页面','Chart unavailable; refresh the page');continue;}
  const chart=charts[p.id]||(charts[p.id]=window.echarts.init(el));
  chart.setOption({animation:false,tooltip:{trigger:'axis',axisPointer:{type:chartType==='bar'?'shadow':'line'}},legend:{top:4},grid:{left:12,right:12,top:55,bottom:10,containLabel:true},
   xAxis:{type:'category',boundaryGap:chartType==='bar',data:data.monthKeys,axisLabel:{rotate:40}},
   yAxis:[{type:'value',name:p.name,axisLabel:{formatter:v=>Number(v).toLocaleString('en-US')},nameTextStyle:{fontSize:11}},
    {type:'value',name:p.shareName,axisLabel:{formatter:v=>(v*100).toFixed(0)+'%'},splitLine:{show:false},nameTextStyle:{fontSize:11}}],
   series:[
    {name:p.name,type:chartType,barMaxWidth:42,yAxisIndex:0,connectNulls:false,symbol:'circle',symbolSize:5,lineStyle:{width:2},itemStyle:{color:p.color},tooltip:{valueFormatter:v=>fmt(v,p.money?'money':'qty')},data:d.months.map(r=>(r.rows||r.qtyShare!=null||r.amountShare!=null)?r[p.field]:null)},
    {name:p.shareName,type:'line',yAxisIndex:1,connectNulls:false,symbol:'circle',symbolSize:4,lineStyle:{width:2,type:'dashed'},itemStyle:{color:p.shareColor},tooltip:{valueFormatter:v=>fmt(v,'pct')},data:d.months.map(r=>r[p.share])}
   ]},true);chart.resize();
 }

}
function recalc(){
 if(!ctx)return;const start=$('start').value,end=$('end').value;
 if(start&&end&&start>end){$('state').textContent=text('开始日期不能晚于结束日期','Start date must precede end date');return;}
 const t=performance.now();data=compute(ctx.sales,ctx.reps,{start,end,includeUniversal:$('universal').checked});
 if(!data.list.some(r=>r.name===selected))selected=data.list[0]?.name||'';
 $('list-title').textContent=text('颜色列表','Color list')+' · '+(start||text('最早','First'))+' — '+(end||data.latest||'—');
 drawList();drawDetail();
 $('state').textContent=text('耗时 ','Time ')+((performance.now()-t)/1000).toFixed(2)+'s';
}
async function refresh(){
 const started=performance.now(),request=++seq,token=cfg.token();$('refresh').disabled=true;$('state').textContent=text('正在检查最新数据…','Checking latest data…');
 try{
  const res=await cfg.normalize(await cfg.api('get_product_context','','',ctx&&ctxToken===token?revision:''));
  if(request!==seq)return;
  if(!res.unchanged){ctx=res;revision=res.revision;ctxToken=token;}
  if(!ctx)throw Error('No source data');recalc();$('state').textContent=text('耗时 ','Time ')+((performance.now()-started)/1000).toFixed(2)+'s';
 }catch(e){if(request===seq){$('state').textContent=text('读取失败，请重试：','Load failed; retry: ')+(e.message||e);$('list').replaceChildren();$('years').replaceChildren();$('months').replaceChildren();Object.values(charts).forEach(chart=>chart.clear());}}
 finally{if(request===seq)$('refresh').disabled=false;}
}
function quick(kind){
 const latest=ctx?.sales.reduce((m,r)=>{const d=validDate(r.soldDate);return d>m?d:m;},'')||'';
 if(!latest)return;const y=+latest.slice(0,4),m=+latest.slice(5,7),monthStart=latest.slice(0,7)+'-01';
 $('end').value=latest;$('start').value=kind==='last'?monthStart:kind==='year'?y+'-01-01':new Date(Date.UTC(y,m-12,1)).toISOString().slice(0,10);recalc();
}
function mount(){
 host=document.createElement('section');host.id='page-color-dashboard';host.className='page';host.style.display='none';
 host.innerHTML='<div class="ct-toolbar"><b>'+text('橱柜颜色动态','Cabinet Color Trends')+'</b><input id="ct-search" type="search" placeholder="'+text('搜索颜色…','Search colors…')+'"><button class="btn sm" id="ct-last">'+text('上月','Last month')+'</button><button class="btn sm" id="ct-year">'+text('今年','This year')+'</button><button class="btn sm" id="ct-ttm">TTM</button><button class="btn sm" id="ct-all">'+text('全部历史','All history')+'</button><label>'+text('开始','Start')+' <input id="ct-start" type="date"></label><label>'+text('结束','End')+' <input id="ct-end" type="date"></label><label><input id="ct-universal" type="checkbox"> '+text('包含 Universal','Include Universal')+'</label><button class="btn sm" id="ct-refresh">'+text('刷新数据','Refresh')+'</button><span id="ct-state" role="status"></span><span id="ct-chart-types" style="display:inline-flex;gap:4px"><button class="btn sm primary" id="ct-bar" aria-pressed="true">'+text('柱状图','Bar')+'</button><button class="btn sm" id="ct-line" aria-pressed="false">'+text('折线图','Line')+'</button></span></div><div class="ct-grid"><div class="ct-left"><b id="ct-list-title"></b><div id="ct-list" class="ct-table"></div></div><div class="ct-right"><b id="ct-selected"></b><div class="ct-analysis"><div id="ct-years" class="ct-table"></div><div id="ct-months" class="ct-table"></div></div><div class="ct-charts"><div class="ct-chart-panel"><div id="ct-chart-qty-title" class="ct-chart-title"></div><div id="ct-chart-qty" class="ct-chart"></div></div><div class="ct-chart-panel"><div id="ct-chart-amount-title" class="ct-chart-title"></div><div id="ct-chart-amount" class="ct-chart"></div></div></div></div></div>';
 document.querySelector('main').appendChild(host);
 const style=document.createElement('style');style.textContent=`#page-color-dashboard{padding:10px 12px;gap:8px;min-height:0;flex:1;overflow:hidden}.ct-toolbar{display:flex;gap:8px;align-items:center;flex-wrap:wrap;background:#edf3fb;padding:8px;border-radius:6px}.ct-toolbar input{border:1px solid #cbd5e1;border-radius:4px;height:28px;padding:3px}.ct-toolbar label{white-space:nowrap}.ct-toolbar input[type=checkbox]{vertical-align:middle}#ct-state{font-size:12px;color:#52657d;white-space:nowrap}.ct-grid{display:flex;gap:10px;flex:1;min-height:0}.ct-left{flex:0 0 30%;display:flex;flex-direction:column;gap:8px;min-height:0;min-width:0}#ct-list{flex:1;min-height:0}#ct-list-title{font-size:13px;color:#174779}.ct-right{display:flex;flex-direction:column;gap:8px;min-width:0;flex:1;min-height:0}#ct-selected{font-size:17px;color:#174779}.ct-analysis{display:flex;gap:8px;flex:0 1 425px;max-height:57%;min-height:180px}.ct-analysis>.ct-table{flex:1;min-width:0}.ct-table{display:flex;flex-direction:column;overflow:hidden;border:1px solid #cbd5e1;border-radius:6px;background:white}.ct-scroll{flex:1;min-height:0;overflow:auto;scrollbar-gutter:stable}.ct-footer{flex:0 0 auto;overflow:hidden;scrollbar-gutter:stable}.ct-table table{width:100%;min-width:0;table-layout:fixed;border-collapse:separate;border-spacing:0;font:13px Arial,sans-serif;color:#182330}.ct-table th,.ct-table td{height:30px;padding:4px 6px;border-right:1px solid #cbd5e1;border-bottom:1px solid #cbd5e1;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.ct-table th{white-space:normal;overflow-wrap:break-word;line-height:15px;position:sticky;top:0;background:#d9eaf7;text-align:center;z-index:2;font-size:12px}.ct-table tfoot td{position:sticky;bottom:0;background:#d9e6f3;font-weight:bold}.ct-table .ct-money{text-align:right}.ct-table .ct-center{text-align:center}#ct-list table{min-width:0}.ct-analysis td:first-child{text-align:center}#ct-list tbody tr{cursor:pointer}#ct-list tbody tr:hover{background:#f0f5ff}#ct-list .ct-selected{background:#dbeafe}.ct-charts{display:flex;flex:1;gap:8px;min-height:190px}.ct-chart-panel{display:flex;flex:1;min-width:0;min-height:180px;flex-direction:column;background:white;border:1px solid #cbd5e1;border-radius:6px;padding:8px}.ct-chart{flex:1;min-height:160px}.ct-chart-title{font-weight:bold;color:#174779;font-size:13px}@media(max-width:1000px){.ct-grid{overflow:auto;flex-direction:column}.ct-left{flex:0 0 280px}.ct-right{min-height:650px}.ct-analysis{max-height:none;flex:0 0 400px}.ct-charts{flex-direction:column;flex:0 0 500px}.ct-right{min-height:960px}}`;
 document.head.appendChild(style);
 for(const type of ['bar','line'])$(type).onclick=()=>{chartType=type;for(const k of ['bar','line']){$(k).className='btn sm'+(chartType===k?' primary':'');$(k).setAttribute('aria-pressed',String(chartType===k));}if(data)drawDetail();};
 $('search').oninput=()=>data&&drawList();$('refresh').onclick=refresh;
 for(const k of ['start','end','universal'])$(k).onchange=recalc;
 for(const k of ['last','year','ttm'])$(k).onclick=()=>quick(k);
 $('all').onclick=()=>{$('start').value='';$('end').value='';recalc();};
 if(window.ResizeObserver){observer=new ResizeObserver(()=>{if(host.style.display!=='none')Object.values(charts).forEach(chart=>chart.resize());});observer.observe($('chart-qty'));observer.observe($('chart-amount'));}
}
window.BCSColorTrends={compute,detail,show(options){cfg=options;if(!host)mount();host.style.display='flex';if(ctx&&ctxToken===cfg.token())recalc();return refresh();}};
})();
