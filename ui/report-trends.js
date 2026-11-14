/* BCS_REPORT_TRENDS_2_BAR: amounts use the annual report's existing attribution. */
(function(){
 'use strict';
 const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
 const money=n=>Number(n||0).toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2});
 function company(rows){
   const years={},months={},present=new Set();
   for(const r of rows){const d=String(r.soldDate||'').slice(0,7);if(!/^\d{4}-(0[1-9]|1[0-2])$/.test(d))continue;present.add(d);if(String(r.rep1||'').trim().toUpperCase()==='IC')continue;const v=Math.round(Number(r.amount||0)*100);years[d.slice(0,4)]=(years[d.slice(0,4)]||0)+v;months[d]=(months[d]||0)+v;}
   for(const d of present)if(!(d.slice(0,4) in years))years[d.slice(0,4)]=0;
   return {years,months,present};
 }
 let host,chart,observer,sequence=0,config,currentSpec=null;
 function hide(){sequence++;if(chart){chart.dispose();chart=null;}if(observer){observer.disconnect();observer=null;}if(host)host.remove();host=null;currentSpec=null;config=null;for(const id of ['univerMount','stat-cond-fields','stat-extra']){const e=document.getElementById(id);if(e&&e.dataset.trendDisplay!==undefined){e.style.display=e.dataset.trendDisplay;delete e.dataset.trendDisplay;}}}
 async function show(c){
   hide();config=c;const token=sequence;const zh=c.zh,word=(a,b)=>zh?a:b;let ctx,spec,year=String(c.year||new Date().getFullYear()),dept='',person='',model,request=0,chartType='bar';
   host=document.createElement('div');host.id='bcs-report-trends';host.style.cssText='flex:1;min-height:0;overflow:auto;padding:12px;background:#f3f7fc;color:#173c60;display:flex;flex-direction:column';
   host.innerHTML='<div class="rt-message">'+word('正在读取曲线数据…','Loading chart data…')+'</div>';document.getElementById('page-sheet').appendChild(host);
   for(const id of ['univerMount','stat-cond-fields','stat-extra']){const e=document.getElementById(id);if(e){e.dataset.trendDisplay=e.style.display;e.style.display='none';}}
   const valid=()=>token===sequence&&c.active();
   const select=(key,label,options,value)=>'<label>'+label+' <select data-rt="'+key+'">'+options.map(o=>'<option value="'+esc(o[0])+'"'+(String(o[0])===String(value)?' selected':'')+'>'+esc(o[1])+'</option>').join('')+'</select></label>';
   async function load(){
     try{
       const serial=++request,start=performance.now();const result=await c.load(c.kind==='performance'?year:null);if(!valid()||serial!==request)return;ctx=result;
       if(c.kind==='performance'){spec=c.build(ctx,year);if(spec.error)throw Error(spec.error);}else model=company(ctx.sales||[]);
       draw();const state=document.getElementById('stat-state');if(state)state.textContent=word('曲线已更新 · 耗时 ','Chart updated · ')+((performance.now()-start)/1000).toFixed(1)+'s';
     }catch(e){if(valid())host.innerHTML='<div role="alert">'+esc(word('读取失败：','Unable to load: ')+e.message)+'</div>';}
   }
   function draw(){
     if(!valid())return;
     let controls,labels,values,title,cards,details,heading;
     if(c.kind==='performance'){
       const detail=spec.performanceYearRows.filter(r=>r.kind==='detail').map(r=>({id:String(r.row),v:spec.values[r.row]}));
       const groups=['House','Other Branch','Outside'];if(!groups.includes(dept))dept=groups[0];const people=detail.filter(r=>r.v[0]===dept);if(!people.some(r=>r.id===person))person=people[0]?.id||'';const row=people.find(r=>r.id===person)?.v;
       controls=select('year',word('年份','Year'),c.years.map(y=>[y,y]),year)+select('dept',word('部门','Department'),groups.map(g=>[g,g]),dept)+select('person',word('销售员','Salesperson'),people.map(r=>[r.id,r.v[1]]),person);
       labels=Array.from({length:12},(_,i)=>year+'-'+String(i+1).padStart(2,'0'));const present=new Set((ctx.sales||[]).map(r=>String(r.soldDate||'').slice(0,7)));
       values=labels.map((m,i)=>present.has(m)?Number(row?.[5+i]||0):null);title=dept+' · '+(row?.[1]||word('暂无销售员','No salesperson'));
       cards=[[word('年度合计','Annual sales'),money(row?.[2])],[word('部门占比','Department share'),(Number(row?.[3]||0)*100).toFixed(2)+'%'],[word('总额占比','Total share'),(Number(row?.[4]||0)*100).toFixed(2)+'%']];heading=word('月份','Month');
     }else{
       const ys=Object.keys(model.years).sort();if(year!=='all'&&!ys.includes(year))year=ys[ys.length-1]||'all';
       controls=select('year',word('查看','View'),[['all',word('历年趋势','All years')],...ys.slice().reverse().map(y=>[y,y])],year);
       if(year==='all'){labels=ys;values=ys.map(y=>model.years[y]/100);title=word('全公司 · 历年销售趋势','Company · Annual sales');heading=word('年份','Year');}
       else{labels=Array.from({length:12},(_,i)=>year+'-'+String(i+1).padStart(2,'0'));values=labels.map(m=>model.present.has(m)?(model.months[m]||0)/100:null);title=word('全公司 · ','Company · ')+year;heading=word('月份','Month');}
       cards=[[word('销售合计','Total sales'),money(values.reduce((a,b)=>a+(b||0),0))],[word('统计口径','Basis'),word('Net（扣除 IC）','Net (excluding IC)')]];
     }
     controls+='<div class="rt-types">'+[['bar',word('柱状图','Bar')],['line',word('折线图','Line')]].map(([type,label])=>'<button type="button" data-chart-type="'+type+'" aria-pressed="'+(chartType===type)+'">'+label+'</button>').join('')+'</div>';
     currentSpec={sheet:title,values:[[heading,word('销售金额','Sales amount')],...labels.map((m,i)=>[m,values[i]===null?'':values[i]]),[word('合计','Total'),values.reduce((a,b)=>a+(b||0),0)]]};
     if(chart){chart.dispose();chart=null;}if(observer)observer.disconnect();
     host.innerHTML='<style>#bcs-report-trends .rt-toolbar{display:flex;flex-wrap:wrap;gap:18px;align-items:center;margin-bottom:12px}#bcs-report-trends select{padding:6px 10px;border:1px solid #becfe3;border-radius:5px;background:white;color:#173c60;max-width:300px}#bcs-report-trends .rt-cards{display:flex;flex-wrap:wrap;gap:12px;margin:12px 0}#bcs-report-trends .rt-card{background:white;border:1px solid #d4e1f0;border-radius:6px;padding:12px 22px;min-width:180px}#bcs-report-trends strong{display:block;font-size:22px;margin-top:5px}#bcs-report-trends .rt-chart{flex:1 0 300px;min-height:300px;background:white;border:1px solid #d4e1f0;border-radius:6px}#bcs-report-trends .rt-types{display:flex;gap:0}#bcs-report-trends .rt-types button{padding:6px 14px;border:1px solid #becfe3;background:white;color:#173c60;cursor:pointer}#bcs-report-trends .rt-types button:first-child{border-radius:5px 0 0 5px}#bcs-report-trends .rt-types button:last-child{border-radius:0 5px 5px 0}#bcs-report-trends .rt-types button[aria-pressed="true"]{background:#245ab5;color:white;border-color:#245ab5}#bcs-report-trends h3{margin:0 0 4px}</style><div class="rt-toolbar">'+controls+'</div><h3>'+esc(title)+'</h3><div class="rt-cards">'+cards.map(([k,v])=>'<div class="rt-card">'+esc(k)+'<strong>'+esc(v)+'</strong></div>').join('')+'</div><div class="rt-chart"></div>';
     host.querySelectorAll('select').forEach(el=>el.onchange=async()=>{const k=el.dataset.rt;if(k==='year'){year=el.value;if(c.kind==='performance'){person='';await load();}else draw();}else{if(k==='dept'){dept=el.value;person='';}else person=el.value;draw();}});
     host.querySelectorAll('[data-chart-type]').forEach(el=>el.onclick=()=>{chartType=el.dataset.chartType;draw();});
     if(!window.echarts){host.querySelector('.rt-chart').textContent=word('图表组件未加载，请刷新重试。','Chart unavailable. Refresh to retry.');return;}
     chart=window.echarts.init(host.querySelector('.rt-chart'));chart.setOption({color:['#245ab5'],tooltip:{trigger:'axis',axisPointer:{type:chartType==='bar'?'shadow':'line'},valueFormatter:v=>v==null?'—':money(v)},grid:{left:90,right:35,top:55,bottom:55},xAxis:{type:'category',data:labels},yAxis:{type:'value',name:word('销售金额','Sales amount'),axisLabel:{formatter:v=>Number(v).toLocaleString()}},series:[{name:word('销售金额','Sales amount'),type:chartType,barMaxWidth:64,label:{show:chartType==='bar',position:'top',formatter:p=>p.value==null?'':Number(p.value).toLocaleString(undefined,{notation:'compact',maximumFractionDigits:1})},labelLayout:{hideOverlap:true},itemStyle:{borderRadius:chartType==='bar'?[4,4,0,0]:0},data:values,connectNulls:false,symbolSize:7,lineStyle:{width:3},areaStyle:{opacity:0}}]});
     if(window.ResizeObserver){observer=new ResizeObserver(()=>chart&&chart.resize());observer.observe(host);}
   }
   if(c.kind==='annual')year='all';await load();
 }
 window.BCSReportTrends={show,hide,exportSpec:()=>currentSpec,company};
})();
