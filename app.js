/* ============================================================
   Supply Chain Intelligence Dashboard — application logic
   RAW is no longer hardcoded — see the data-loader block at the
   bottom of this file, which parses an .xlsx file client-side
   (via SheetJS) and calls initDashboard(rows) with the result.
   ============================================================ */

/* ---------- palette ---------- */
const COLORS = {
  gold:'#E3A857', teal:'#45C4B0', blue:'#5B8DEF', red:'#E2645A', violet:'#9C87E0',
  ink:'#E9EDF6', dim:'#97A3BC', faint:'#5E6B87', line:'#28324A', panel:'#161E2C'
};
const CAT_COLORS = ['#E3A857','#45C4B0','#5B8DEF','#E2645A','#9C87E0','#6FCF97','#F2C94C'];
Chart.defaults.font.family = "'IBM Plex Sans', sans-serif";
Chart.defaults.color = COLORS.dim;
Chart.defaults.borderColor = COLORS.line;
Chart.defaults.plugins.legend.labels.boxWidth = 10;
Chart.defaults.plugins.legend.labels.font = {size:11};

/* ---------- helpers ---------- */
const sum = arr => arr.reduce((a,b)=>a+b,0);
const mean = arr => arr.length ? sum(arr)/arr.length : 0;
const std = arr => { const m = mean(arr); return Math.sqrt(mean(arr.map(x=>(x-m)**2))); };
const fmt0 = n => Math.round(n).toLocaleString('en-US');
const fmt1 = n => n.toLocaleString('en-US',{maximumFractionDigits:1,minimumFractionDigits:1});
const fmt2 = n => n.toLocaleString('en-US',{maximumFractionDigits:2,minimumFractionDigits:2});
const usd = n => '$'+fmt0(n);
const usd1 = n => '$'+fmt1(n);
const pct = n => fmt1(n)+'%';
function groupBy(arr, key){
  const m = new Map();
  arr.forEach(r=>{ const k=r[key]; if(!m.has(k)) m.set(k,[]); m.get(k).push(r); });
  return m;
}
function pearson(xs, ys){
  const mx=mean(xs), my=mean(ys);
  let num=0, dx=0, dy=0;
  for(let i=0;i<xs.length;i++){ num += (xs[i]-mx)*(ys[i]-my); dx += (xs[i]-mx)**2; dy += (ys[i]-my)**2; }
  return num / Math.sqrt(dx*dy);
}
function alpha(hex, a){
  const r=parseInt(hex.slice(1,3),16), g=parseInt(hex.slice(3,5),16), b=parseInt(hex.slice(5,7),16);
  return `rgba(${r},${g},${b},${a})`;
}
function heatColor(t){ // t 0..1 -> gold gradient on dark
  const stops = [[22,29,44],[61,54,42],[130,90,45],[201,132,58],[227,168,87]];
  const seg = Math.min(3, Math.floor(t*4));
  const f = (t*4)-seg;
  const c0=stops[seg], c1=stops[seg+1];
  const r = Math.round(c0[0]+(c1[0]-c0[0])*f), g=Math.round(c0[1]+(c1[1]-c0[1])*f), b=Math.round(c0[2]+(c1[2]-c0[2])*f);
  return `rgb(${r},${g},${b})`;
}

/* ---------- field aliases ---------- */
const F = {
  type:'Product type', sku:'SKU', price:'Price', avail:'Availability', sold:'Number of products sold',
  rev:'Revenue generated', demo:'Customer demographics', stock:'Stock levels', leadTimes:'Lead times',
  orderQty:'Order quantities', shipTime:'Shipping times', carrier:'Shipping carriers', shipCost:'Shipping costs',
  supplier:'Supplier name', loc:'Location', supLeadTime:'Lead time', prodVol:'Production volumes',
  mfgLeadTime:'Manufacturing lead time', mfgCost:'Manufacturing costs', inspect:'Inspection results',
  defect:'Defect rates', mode:'Transportation modes', route:'Routes', cost:'Costs'
};

/* ============================================================
   initDashboard — everything below depends on RAW (the parsed
   worksheet rows) and only runs once real data has been loaded.
   ============================================================ */
function initDashboard(RAW){

/* ============================================================
   PRECOMPUTE — one pass over the data building every aggregate
   ============================================================ */
const TYPES = [...new Set(RAW.map(r=>r[F.type]))];
const SUPPLIERS = [...new Set(RAW.map(r=>r[F.supplier]))].sort();
const LOCATIONS = [...new Set(RAW.map(r=>r[F.loc]))];
const CARRIERS = [...new Set(RAW.map(r=>r[F.carrier]))];
const MODES = [...new Set(RAW.map(r=>r[F.mode]))];
const ROUTES = [...new Set(RAW.map(r=>r[F.route]))];

const totalRevenue = sum(RAW.map(r=>r[F.rev]));
const totalSold = sum(RAW.map(r=>r[F.sold]));
const avgDefect = mean(RAW.map(r=>r[F.defect]));
const totalMfgCost = sum(RAW.map(r=>r[F.mfgCost]));
const totalShipCost = sum(RAW.map(r=>r[F.shipCost]));
const avgLeadTime = mean(RAW.map(r=>r[F.leadTimes]));
const avgMfgLeadTime = mean(RAW.map(r=>r[F.mfgLeadTime]));
const totalStock = sum(RAW.map(r=>r[F.stock]));
const approxMargin = mean(RAW.map(r=> (r[F.rev]-r[F.mfgCost]-r[F.shipCost]) / r[F.rev] * 100 ));

function revenueBy(key){
  const g = groupBy(RAW, key);
  return [...g.entries()].map(([k,rows])=>({label:k, value:sum(rows.map(r=>r[F.rev]))})).sort((a,b)=>b.value-a.value);
}

/* ---- inventory ---- */
const avgStock = mean(RAW.map(r=>r[F.stock]));
const invTurnover = totalSold / avgStock;
const invDays = 365 / invTurnover;
const fillRate = mean(RAW.map(r=>r[F.avail]));

/* ---- ABC classification (by revenue, per doc: top 20% count = A, next 30% = B, rest = C) ---- */
const byRevDesc = [...RAW].sort((a,b)=>b[F.rev]-a[F.rev]);
const nA = Math.round(byRevDesc.length*0.2), nB = Math.round(byRevDesc.length*0.3);
byRevDesc.forEach((r,i)=>{ r._abc = i<nA ? 'A' : (i<nA+nB ? 'B' : 'C'); });
const abcSummary = ['A','B','C'].map(cls=>{
  const rows = byRevDesc.filter(r=>r._abc===cls);
  return {cls, count:rows.length, revenue:sum(rows.map(r=>r[F.rev]))};
});

/* ---- supplier scorecard ---- */
const supplierStats = SUPPLIERS.map(sup=>{
  const rows = RAW.filter(r=>r[F.supplier]===sup);
  const pass = rows.filter(r=>r[F.inspect]==='Pass').length;
  return {
    supplier: sup,
    avgLead: mean(rows.map(r=>r[F.supLeadTime])),
    avgDefect: mean(rows.map(r=>r[F.defect])),
    avgShipCost: mean(rows.map(r=>r[F.shipCost])),
    revenue: sum(rows.map(r=>r[F.rev])),
    passRate: pass/rows.length*100,
    avgMfgCost: mean(rows.map(r=>r[F.mfgCost])),
    n: rows.length
  };
});
// composite score: 40% cost, 30% lead time, 20% defect, 10% inspection (lower cost/lead/defect better -> invert)
function norm(vals, invert){
  const mn=Math.min(...vals), mx=Math.max(...vals);
  return vals.map(v=> mx===mn ? 0.5 : (invert ? (mx-v)/(mx-mn) : (v-mn)/(mx-mn)) );
}
{
  const costN = norm(supplierStats.map(s=>s.avgMfgCost+s.avgShipCost), true);
  const leadN = norm(supplierStats.map(s=>s.avgLead), true);
  const defN  = norm(supplierStats.map(s=>s.avgDefect), true);
  const insN  = norm(supplierStats.map(s=>s.passRate), false);
  supplierStats.forEach((s,i)=>{ s.score = costN[i]*0.4 + leadN[i]*0.3 + defN[i]*0.2 + insN[i]*0.1; s.score*=100; });
}
supplierStats.sort((a,b)=>b.score-a.score);

/* ---- manufacturing ---- */
const totalProdVol = sum(RAW.map(r=>r[F.prodVol]));
const inspectCounts = {Pass:0, Fail:0, Pending:0};
RAW.forEach(r=> inspectCounts[r[F.inspect]]++ );
const inspectPassPct = inspectCounts.Pass/RAW.length*100;

/* ---- logistics ---- */
const avgShipTime = mean(RAW.map(r=>r[F.shipTime]));
const avgShipCost = mean(RAW.map(r=>r[F.shipCost]));
const carrierStats = CARRIERS.map(c=>{
  const rows = RAW.filter(r=>r[F.carrier]===c);
  return {label:c, cost:mean(rows.map(r=>r[F.shipCost])), time:mean(rows.map(r=>r[F.shipTime])), n:rows.length};
});
const bestCarrier = [...carrierStats].sort((a,b)=>a.time-b.time)[0];
const routeStats = ROUTES.map(r0=>{
  const rows = RAW.filter(r=>r[F.route]===r0);
  return {label:r0, cost:mean(rows.map(r=>r[F.shipCost])), time:mean(rows.map(r=>r[F.shipTime])), n:rows.length};
});
const bestRoute = [...routeStats].sort((a,b)=>a.time-b.time)[0];

/* ---- Pareto (defects & cost by supplier) ---- */
function paretoSeries(statArr, valueKey){
  const sorted = [...statArr].sort((a,b)=>b[valueKey]-a[valueKey]);
  const total = sum(sorted.map(s=>s[valueKey]));
  let cum = 0;
  return sorted.map(s=>{ cum += s[valueKey]; return {label:s.supplier||s.label, value:s[valueKey], cumPct: cum/total*100}; });
}
const supplierDefectTotal = SUPPLIERS.map(sup=>{
  const rows = RAW.filter(r=>r[F.supplier]===sup);
  return {supplier:sup, value: sum(rows.map(r=>r[F.defect]))};
});
const supplierCostTotal = SUPPLIERS.map(sup=>{
  const rows = RAW.filter(r=>r[F.supplier]===sup);
  return {supplier:sup, value: sum(rows.map(r=>r[F.mfgCost]+r[F.shipCost]))};
});
const paretoDefect = paretoSeries(supplierDefectTotal,'value');
const paretoCost = paretoSeries(supplierCostTotal,'value');

/* ---- XYZ (demand variability proxy: CV of units sold within product type) ---- */
const xyzStats = TYPES.map(t=>{
  const rows = RAW.filter(r=>r[F.type]===t);
  const vals = rows.map(r=>r[F.sold]);
  const cv = std(vals)/mean(vals);
  const cls = cv < 0.5 ? 'X' : (cv < 0.8 ? 'Y' : 'Z');
  return {type:t, cv, cls, n:rows.length};
});

/* ---- EOQ / Safety stock / ROP (assumption-based) ---- */
const ASSUME_ORDER_COST = 50;   // $ per order, stated assumption
const ASSUME_HOLD_PCT = 0.2;    // 20% of unit price annually, stated assumption
const ASSUME_Z = 1.65;          // ~95% service level
const eoqStats = TYPES.map(t=>{
  const rows = RAW.filter(r=>r[F.type]===t);
  const avgPrice = mean(rows.map(r=>r[F.price]));
  const annualDemand = sum(rows.map(r=>r[F.sold])); // treat snapshot as annual proxy
  const holdCost = avgPrice*ASSUME_HOLD_PCT;
  const eoq = Math.sqrt((2*annualDemand*ASSUME_ORDER_COST)/Math.max(holdCost,0.01));
  const avgDemandDaily = annualDemand/365;
  const avgLead = mean(rows.map(r=>r[F.leadTimes]));
  const leadStd = std(rows.map(r=>r[F.leadTimes]));
  const rop = avgDemandDaily*avgLead;
  const safety = ASSUME_Z*leadStd*avgDemandDaily;
  return {type:t, eoq, rop, safety, avgLead, annualDemand};
});

/* ---- correlations ---- */
const corr = {
  volCost: pearson(RAW.map(r=>r[F.prodVol]), RAW.map(r=>r[F.mfgCost])),
  shipCostTime: pearson(RAW.map(r=>r[F.shipCost]), RAW.map(r=>r[F.shipTime])),
  leadDefect: pearson(RAW.map(r=>r[F.supLeadTime]), RAW.map(r=>r[F.defect])),
};

/* ---- AI-style insights (computed, not hardcoded) ---- */
function buildInsights(){
  const worstSupplier = [...supplierStats].sort((a,b)=>b.avgDefect-a.avgDefect)[0];
  const bestLeadSupplier = [...supplierStats].sort((a,b)=>a.avgLead-b.avgLead)[0];
  const modeCost = MODES.map(m=>{
    const rows = RAW.filter(r=>r[F.mode]===m);
    return {mode:m, cost: sum(rows.map(r=>r[F.cost]))};
  }).sort((a,b)=>b.cost-a.cost)[0];
  const typeRevPct = revenueBy(F.type)[0];
  const typeRevShare = typeRevPct.value/totalRevenue*100;
  // high inventory, low sales SKU
  const stockSoldRatio = [...RAW].sort((a,b)=> (b[F.stock]/Math.max(b[F.sold],1)) - (a[F.stock]/Math.max(a[F.sold],1)) )[0];
  const fastestCarrier = [...carrierStats].sort((a,b)=>a.time-b.time)[0];
  const priciestCarrier = [...carrierStats].sort((a,b)=>b.cost-a.cost)[0];
  const carrierPremium = (priciestCarrier.cost-mean(carrierStats.map(c=>c.cost)))/mean(carrierStats.map(c=>c.cost))*100;

  return [
    `<b>${worstSupplier.supplier}</b> carries the highest average defect rate at <b>${pct(worstSupplier.avgDefect)}</b> — a priority candidate for a supplier quality review.`,
    `<b>${modeCost.mode}</b> transport contributes the highest aggregate logistics cost, at <b>${usd(modeCost.cost)}</b> across all shipments.`,
    `<b>${typeRevPct.label}</b> generates <b>${pct(typeRevShare)}</b> of total revenue — the largest single category contribution.`,
    `<b>${bestLeadSupplier.supplier}</b> has the shortest average lead time at <b>${fmt1(bestLeadSupplier.avgLead)} days</b>, a benchmark for the supplier base.`,
    `<b>${stockSoldRatio[F.sku]}</b> (${stockSoldRatio[F.type]}) shows high inventory (${fmt0(stockSoldRatio[F.stock])} units) against low sales (${fmt0(stockSoldRatio[F.sold])} units) — a stagnant-stock risk.`,
    `<b>${fastestCarrier.label}</b> delivers fastest at <b>${fmt1(fastestCarrier.time)} days</b> on average, while <b>${priciestCarrier.label}</b> runs about <b>${pct(Math.abs(carrierPremium))}</b> above the carrier-average cost.`
  ];
}

/* ============================================================
   PAGE DEFINITIONS
   ============================================================ */
const PAGES = [
  {id:'exec', idx:'01 · SOURCE', label:'Executive Overview', tag:'Business health'},
  {id:'inv', idx:'02 · STOCK', label:'Inventory & Demand', tag:'ABC · stockout risk'},
  {id:'sup', idx:'03 · SOURCE', label:'Supplier Scorecard', tag:'Quality · cost · speed'},
  {id:'mfg', idx:'04 · MAKE', label:'Manufacturing & Quality', tag:'Production · inspection'},
  {id:'log', idx:'05 · DELIVER', label:'Logistics & Distribution', tag:'Carriers · routes'},
  {id:'opt', idx:'06 · OPTIMIZE', label:'Optimization & Recs', tag:'Pareto · what-if · insight'}
];

const nav = document.getElementById('flow-nav');
PAGES.forEach((p,i)=>{
  const b = document.createElement('button');
  b.className = 'flow-node' + (i===0?' active':'');
  b.dataset.page = p.id;
  b.innerHTML = `<span class="fn-idx">${p.idx}</span><span class="fn-label">${p.label}</span><span class="fn-tag">${p.tag}</span>`;
  b.onclick = () => showPage(p.id);
  nav.appendChild(b);
});

document.getElementById('meta-rows').textContent = RAW.length;
document.getElementById('meta-sup').textContent = SUPPLIERS.length;
document.getElementById('meta-loc').textContent = LOCATIONS.length;

function showPage(id){
  document.querySelectorAll('.page').forEach(p=>p.classList.toggle('active', p.id==='page-'+id));
  document.querySelectorAll('.flow-node').forEach(n=>n.classList.toggle('active', n.dataset.page===id));
  window.scrollTo({top:0, behavior:'smooth'});
}

/* ============================================================
   RENDER — build DOM for each page
   ============================================================ */
const main = document.getElementById('main');

/* generic chart constructor helpers */
function baseGrid(){ return {color:'rgba(255,255,255,0.045)'}; }
function baseTicks(){ return {color:COLORS.faint, font:{size:10.5}}; }
function mkChart(ctx, config){
  config.options = config.options || {};
  config.options.responsive = true;
  config.options.maintainAspectRatio = false;
  config.options.plugins = Object.assign({legend:{display:false}}, config.options.plugins||{});
  if(config.options.scales){
    Object.values(config.options.scales).forEach(s=>{ s.grid = Object.assign(baseGrid(), s.grid||{}); s.ticks = Object.assign(baseTicks(), s.ticks||{}); });
  }
  return new Chart(ctx, config);
}

function el(html){ const d=document.createElement('div'); d.innerHTML=html.trim(); return d.firstChild; }

/* ---------------------------------------------------------- */
/* PAGE 1 — EXECUTIVE OVERVIEW                                 */
/* ---------------------------------------------------------- */
function pageExec(){
  const p = el(`<section class="page" id="page-exec">
    <div class="page-head"><h2>Executive Supply Chain Overview</h2><p class="q">"How is my supply chain performing overall?"</p></div>
    <div class="section-title">Headline KPIs</div>
    <div class="kpi-row" id="exec-kpis"></div>
    <div class="section-title">Revenue Composition</div>
    <div class="grid cols-2">
      <div class="card"><h3>Revenue Trend</h3><p class="card-sub">By SKU sequence (SKU0 → SKU99)</p><div class="chart-wrap"><canvas id="c-exec-trend"></canvas></div>
        <p class="footnote">No order date field exists in the source data, so this trend is plotted across SKU sequence rather than calendar time.</p></div>
      <div class="card"><h3>Revenue by Product Type</h3><div class="chart-wrap"><canvas id="c-exec-type"></canvas></div></div>
      <div class="card"><h3>Revenue by Supplier</h3><div class="chart-wrap"><canvas id="c-exec-sup"></canvas></div></div>
      <div class="card"><h3>Revenue by Location</h3><div class="chart-wrap"><canvas id="c-exec-loc"></canvas></div></div>
      <div class="card"><h3>Revenue by Transportation Mode</h3><div class="chart-wrap"><canvas id="c-exec-mode"></canvas></div></div>
      <div class="card"><h3>Revenue vs Manufacturing Cost</h3><p class="card-sub">Per SKU</p><div class="chart-wrap"><canvas id="c-exec-scatter"></canvas></div></div>
    </div>
  </section>`);
  main.appendChild(p);

  const kpis = [
    ['Total Revenue', usd(totalRevenue), '', ''],
    ['Total Products Sold', fmt0(totalSold)+' units', '', 'accent-teal'],
    ['Avg Defect Rate', pct(avgDefect), 'across all SKUs', 'accent-red'],
    ['Total Manufacturing Cost', usd(totalMfgCost), '', ''],
    ['Total Shipping Cost', usd(totalShipCost), '', 'accent-blue'],
    ['Avg Lead Time', fmt1(avgLeadTime)+' days', 'order lead time', ''],
    ['Avg Mfg Lead Time', fmt1(avgMfgLeadTime)+' days', '', 'accent-violet'],
    ['Inventory Available', fmt0(totalStock)+' units', 'on hand', 'accent-teal'],
    ['Avg Profit Margin (Approx.)', pct(approxMargin), 'rev − mfg − ship cost', 'accent-gold']
  ];
  const row = document.getElementById('exec-kpis');
  kpis.forEach(([label,value,note,cls])=>{
    row.appendChild(el(`<div class="kpi ${cls||''}"><div class="kpi-label">${label}</div><div class="kpi-value">${value}</div>${note?`<div class="kpi-note">${note}</div>`:''}</div>`));
  });

  // revenue trend by SKU sequence
  const sortedBySku = [...RAW].sort((a,b)=> parseInt(a[F.sku].replace('SKU',''))-parseInt(b[F.sku].replace('SKU','')));
  mkChart(document.getElementById('c-exec-trend'), {
    type:'line',
    data:{ labels: sortedBySku.map(r=>r[F.sku]),
      datasets:[{ data: sortedBySku.map(r=>r[F.rev]), borderColor:COLORS.gold, backgroundColor:alpha(COLORS.gold,0.12), fill:true, tension:0.35, pointRadius:0, borderWidth:2 }]},
    options:{ scales:{ x:{ ticks:{ maxTicksLimit:8 } }, y:{ ticks:{ callback:v=>'$'+v/1000+'k' } } } }
  });

  function catBar(canvasId, seriesArr, horizontal){
    mkChart(document.getElementById(canvasId), {
      type:'bar',
      data:{ labels: seriesArr.map(s=>s.label), datasets:[{ data: seriesArr.map(s=>s.value), backgroundColor: seriesArr.map((_,i)=>CAT_COLORS[i%CAT_COLORS.length]), borderRadius:5, maxBarThickness:38 }]},
      options:{ indexAxis: horizontal?'y':'x', scales:{ x:{ ticks:{ callback:v=> horizontal? '$'+v/1000+'k' : undefined } }, y:{ ticks:{ callback:v=> !horizontal? '$'+v/1000+'k' : undefined } } } }
    });
  }
  catBar('c-exec-type', revenueBy(F.type), false);
  catBar('c-exec-sup', revenueBy(F.supplier), false);
  catBar('c-exec-loc', revenueBy(F.loc), true);
  catBar('c-exec-mode', revenueBy(F.mode), false);

  mkChart(document.getElementById('c-exec-scatter'), {
    type:'scatter',
    data:{ datasets: TYPES.map((t,i)=>({ label:t, data: RAW.filter(r=>r[F.type]===t).map(r=>({x:r[F.mfgCost], y:r[F.rev]})), backgroundColor: alpha(CAT_COLORS[i],0.75), pointRadius:4 })) },
    options:{ plugins:{legend:{display:true, position:'top'}}, scales:{ x:{title:{display:true,text:'Manufacturing Cost ($)',color:COLORS.faint,font:{size:10.5}}}, y:{title:{display:true,text:'Revenue ($)',color:COLORS.faint,font:{size:10.5}}} } }
  });
}

/* ---------------------------------------------------------- */
/* PAGE 2 — INVENTORY & DEMAND                                  */
/* ---------------------------------------------------------- */
function pageInv(){
  const p = el(`<section class="page" id="page-inv">
    <div class="page-head"><h2>Inventory & Demand Analytics</h2><p class="q">Which products are overstocked or likely to stock out — and which suppliers create inventory problems?</p></div>
    <div class="section-title">KPIs</div>
    <div class="kpi-row" id="inv-kpis"></div>
    <div class="section-title">Stock Position</div>
    <div class="grid cols-2">
      <div class="card"><h3>Stock Levels by SKU</h3><p class="card-sub">Top 15 highest-stock SKUs</p><div class="chart-wrap"><canvas id="c-inv-stock"></canvas></div></div>
      <div class="card"><h3>Stock Levels vs Products Sold</h3><div class="chart-wrap"><canvas id="c-inv-vs"></canvas></div></div>
      <div class="card"><h3>Inventory by Product Type</h3><div class="chart-wrap"><canvas id="c-inv-type"></canvas></div></div>
      <div class="card"><h3>Availability Heatmap</h3><p class="card-sub">Avg availability % — product type × location</p><div id="hm-avail"></div></div>
    </div>
    <div class="section-title">Stock vs Sales vs Revenue</div>
    <div class="card"><h3>Bubble Chart — Stock (x) · Sales (y) · Revenue (bubble size)</h3><div class="chart-wrap tall"><canvas id="c-inv-bubble"></canvas></div></div>
    <div class="section-title">ABC Analysis (MBA technique)</div>
    <div class="grid cols-2">
      <div class="card"><h3>SKU Classification by Revenue Contribution</h3><p class="card-sub">A = top 20% of SKUs by revenue · B = next 30% · C = remaining 50%</p><div class="chart-wrap"><canvas id="c-inv-abc"></canvas></div></div>
      <div class="card"><h3>ABC Summary</h3><table class="dt" id="t-abc"></table></div>
    </div>
  </section>`);
  main.appendChild(p);

  const kpis = [
    ['Inventory Turnover', fmt2(invTurnover)+'x', 'units sold ÷ avg stock', ''],
    ['Average Stock Level', fmt0(avgStock)+' units', '', 'accent-blue'],
    ['Inventory Days', fmt0(invDays)+' days', 'to cycle inventory', 'accent-violet'],
    ['Fill Rate (Approx.)', pct(fillRate), 'avg availability score', 'accent-teal']
  ];
  const row = document.getElementById('inv-kpis');
  kpis.forEach(([label,value,note,cls])=> row.appendChild(el(`<div class="kpi ${cls||''}"><div class="kpi-label">${label}</div><div class="kpi-value">${value}</div>${note?`<div class="kpi-note">${note}</div>`:''}</div>`)));

  const top15 = [...RAW].sort((a,b)=>b[F.stock]-a[F.stock]).slice(0,15);
  mkChart(document.getElementById('c-inv-stock'), {
    type:'bar', data:{ labels: top15.map(r=>r[F.sku]), datasets:[{ data: top15.map(r=>r[F.stock]), backgroundColor: alpha(COLORS.blue,0.8), borderRadius:4 }]},
    options:{ scales:{ x:{ticks:{maxTicksLimit:15, font:{size:9}}} } }
  });

  mkChart(document.getElementById('c-inv-vs'), {
    type:'scatter',
    data:{ datasets:[{ label:'SKU', data: RAW.map(r=>({x:r[F.stock], y:r[F.sold]})), backgroundColor: alpha(COLORS.teal,0.7), pointRadius:4 }]},
    options:{ scales:{ x:{title:{display:true,text:'Stock Level',color:COLORS.faint,font:{size:10.5}}}, y:{title:{display:true,text:'Products Sold',color:COLORS.faint,font:{size:10.5}}} } }
  });

  const invByType = TYPES.map(t=>({label:t, value: sum(RAW.filter(r=>r[F.type]===t).map(r=>r[F.stock]))}));
  mkChart(document.getElementById('c-inv-type'), {
    type:'doughnut', data:{ labels: invByType.map(s=>s.label), datasets:[{ data: invByType.map(s=>s.value), backgroundColor: CAT_COLORS, borderColor:COLORS.panel, borderWidth:3 }]},
    options:{ plugins:{legend:{display:true, position:'bottom'}}, cutout:'62%' }
  });

  // availability heatmap: type x location
  const hm = document.getElementById('hm-avail');
  let hmHtml = `<div class="heatmap" style="--cols:${LOCATIONS.length}"><div class="hm-row"><div></div>${LOCATIONS.map(l=>`<div class="hm-head">${l}</div>`).join('')}</div>`;
  TYPES.forEach(t=>{
    hmHtml += `<div class="hm-row"><div class="hm-rowlabel">${t}</div>`;
    LOCATIONS.forEach(l=>{
      const rows = RAW.filter(r=>r[F.type]===t && r[F.loc]===l);
      const v = rows.length ? mean(rows.map(r=>r[F.avail])) : null;
      const t01 = v===null ? 0 : v/100;
      hmHtml += `<div class="hm-cell" style="background:${v===null?'transparent':heatColor(t01)}">${v===null?'—':fmt0(v)}</div>`;
    });
    hmHtml += `</div>`;
  });
  hmHtml += `</div>`;
  hm.innerHTML = hmHtml;

  mkChart(document.getElementById('c-inv-bubble'), {
    type:'bubble',
    data:{ datasets: TYPES.map((t,i)=>({ label:t, data: RAW.filter(r=>r[F.type]===t).map(r=>({x:r[F.stock], y:r[F.sold], r: Math.max(4, Math.sqrt(r[F.rev])/9)})), backgroundColor: alpha(CAT_COLORS[i],0.55) })) },
    options:{ plugins:{legend:{display:true, position:'top'}}, scales:{ x:{title:{display:true,text:'Stock Level',color:COLORS.faint,font:{size:10.5}}}, y:{title:{display:true,text:'Products Sold',color:COLORS.faint,font:{size:10.5}}} } }
  });

  mkChart(document.getElementById('c-inv-abc'), {
    type:'pie',
    data:{ labels: abcSummary.map(a=>`Class ${a.cls} (${a.count} SKUs)`), datasets:[{ data: abcSummary.map(a=>a.revenue), backgroundColor:[COLORS.teal,COLORS.gold,COLORS.red], borderColor:COLORS.panel, borderWidth:3 }]},
    options:{ plugins:{legend:{display:true, position:'bottom'}} }
  });

  const tbl = document.getElementById('t-abc');
  tbl.innerHTML = `<thead><tr><th>Class</th><th># SKUs</th><th>Revenue</th><th>% of Total</th></tr></thead><tbody>
    ${abcSummary.map(a=>`<tr><td class="name"><span class="pill ${a.cls.toLowerCase()}">${a.cls}</span></td><td>${a.count}</td><td>${usd(a.revenue)}</td><td>${pct(a.revenue/totalRevenue*100)}</td></tr>`).join('')}
  </tbody>`;
}

/* ---------------------------------------------------------- */
/* PAGE 3 — SUPPLIER PERFORMANCE                                */
/* ---------------------------------------------------------- */
function pageSup(){
  const p = el(`<section class="page" id="page-sup">
    <div class="page-head"><h2>Supplier Performance Scorecard</h2><p class="q">The most consequential page — who to double down on, and who needs a corrective plan.</p></div>
    <div class="section-title">Supplier Scorecard</div>
    <div class="card"><table class="dt" id="t-supplier"></table></div>
    <div class="section-title">Ranking & Comparison</div>
    <div class="grid cols-2">
      <div class="card"><h3>Supplier Ranking</h3><p class="card-sub">Composite score — 40% cost · 30% lead time · 20% defect · 10% inspection</p><div class="chart-wrap"><canvas id="c-sup-rank"></canvas></div></div>
      <div class="card"><h3>Supplier Comparison Matrix</h3><p class="card-sub">Radar — normalized 0–1, outer edge = better</p><div class="chart-wrap"><canvas id="c-sup-radar"></canvas></div></div>
    </div>
    <div class="section-title">Risk Positioning</div>
    <div class="grid cols-2">
      <div class="card"><h3>Lead Time vs Defect Rate</h3><p class="card-sub">Bubble = revenue supported, color = supplier</p><div class="chart-wrap"><canvas id="c-sup-scatter"></canvas></div></div>
      <div class="card"><h3>Quadrant Analysis</h3><p class="card-sub">Ideal supplier = low lead time + low defect rate</p><div class="chart-wrap"><canvas id="c-sup-quad"></canvas></div>
        <div class="quad-legend">
          <span><i style="background:${COLORS.teal}"></i>Ideal (low lead, low defect)</span>
          <span><i style="background:${COLORS.gold}"></i>Watch</span>
          <span><i style="background:${COLORS.red}"></i>High risk (high lead + high defect)</span>
        </div>
      </div>
    </div>
  </section>`);
  main.appendChild(p);

  const tbl = document.getElementById('t-supplier');
  tbl.innerHTML = `<thead><tr><th>Supplier</th><th>Avg Lead Time</th><th>Avg Defect Rate</th><th>Avg Shipping Cost</th><th>Revenue Supported</th><th>Inspection Pass Rate</th></tr></thead><tbody>
    ${[...supplierStats].sort((a,b)=>b.revenue-a.revenue).map(s=>`<tr><td class="name">${s.supplier}</td><td>${fmt1(s.avgLead)}d</td><td>${pct(s.avgDefect)}</td><td>${usd1(s.avgShipCost)}</td><td>${usd(s.revenue)}</td><td>${pct(s.passRate)}</td></tr>`).join('')}
  </tbody>`;

  mkChart(document.getElementById('c-sup-rank'), {
    type:'bar',
    data:{ labels: supplierStats.map(s=>s.supplier), datasets:[{ data: supplierStats.map(s=>s.score), backgroundColor: supplierStats.map((_,i)=>CAT_COLORS[i%CAT_COLORS.length]), borderRadius:5 }]},
    options:{ indexAxis:'y', scales:{ x:{ suggestedMax:100 } } }
  });

  const radarMetrics = [
    {label:'Cost efficiency', get:s=>1-norm(supplierStats.map(x=>x.avgMfgCost+x.avgShipCost),false)[supplierStats.indexOf(s)] },
  ];
  // build normalized radar dataset directly
  const costsArr = supplierStats.map(s=>s.avgMfgCost+s.avgShipCost);
  const costN = norm(costsArr, true), leadN = norm(supplierStats.map(s=>s.avgLead), true), defN = norm(supplierStats.map(s=>s.avgDefect), true),
        revN = norm(supplierStats.map(s=>s.revenue), false), insN = norm(supplierStats.map(s=>s.passRate), false);
  mkChart(document.getElementById('c-sup-radar'), {
    type:'radar',
    data:{ labels:['Cost eff.','Speed (lead time)','Quality (low defect)','Revenue supported','Inspection pass'],
      datasets: supplierStats.map((s,i)=>({ label:s.supplier, data:[costN[i],leadN[i],defN[i],revN[i],insN[i]], borderColor:CAT_COLORS[i%CAT_COLORS.length], backgroundColor:alpha(CAT_COLORS[i%CAT_COLORS.length],0.08), pointRadius:2, borderWidth:1.6 })) },
    options:{ plugins:{legend:{display:true, position:'bottom'}}, scales:{ r:{ min:0, max:1, ticks:{display:false}, grid:{color:'rgba(255,255,255,0.06)'}, angleLines:{color:'rgba(255,255,255,0.06)'}, pointLabels:{color:COLORS.dim, font:{size:10.5}} } } }
  });

  mkChart(document.getElementById('c-sup-scatter'), {
    type:'bubble',
    data:{ datasets: SUPPLIERS.map((sup,i)=>({ label:sup, data: RAW.filter(r=>r[F.supplier]===sup).map(r=>({x:r[F.supLeadTime], y:r[F.defect], r: Math.max(4, Math.sqrt(r[F.rev])/10)})), backgroundColor: alpha(CAT_COLORS[i%CAT_COLORS.length],0.55) })) },
    options:{ plugins:{legend:{display:true, position:'top'}}, scales:{ x:{title:{display:true,text:'Lead Time (days)',color:COLORS.faint,font:{size:10.5}}}, y:{title:{display:true,text:'Defect Rate (%)',color:COLORS.faint,font:{size:10.5}}} } }
  });

  // quadrant chart with median crosshair
  const medLead = [...supplierStats.map(s=>s.avgLead)].sort((a,b)=>a-b)[Math.floor(supplierStats.length/2)];
  const medDefect = [...supplierStats.map(s=>s.avgDefect)].sort((a,b)=>a-b)[Math.floor(supplierStats.length/2)];
  const quadPlugin = {
    id:'quadLines',
    afterDraw(chart){
      const {ctx, chartArea:{left,right,top,bottom}, scales:{x,y}} = chart;
      ctx.save();
      ctx.strokeStyle='rgba(255,255,255,0.14)'; ctx.setLineDash([4,4]); ctx.lineWidth=1;
      const xm = x.getPixelForValue(medLead), ym = y.getPixelForValue(medDefect);
      ctx.beginPath(); ctx.moveTo(xm, top); ctx.lineTo(xm, bottom); ctx.stroke();
      ctx.beginPath(); ctx.moveTo(left, ym); ctx.lineTo(right, ym); ctx.stroke();
      ctx.restore();
    }
  };
  mkChart(document.getElementById('c-sup-quad'), {
    type:'scatter',
    data:{ datasets: supplierStats.map((s,i)=>({ label:s.supplier, data:[{x:s.avgLead, y:s.avgDefect}],
      backgroundColor: (s.avgLead<=medLead && s.avgDefect<=medDefect) ? COLORS.teal : ((s.avgLead>medLead && s.avgDefect>medDefect) ? COLORS.red : COLORS.gold),
      pointRadius:9, pointHoverRadius:11 })) },
    options:{ plugins:{legend:{display:true, position:'top'}}, scales:{ x:{title:{display:true,text:'Avg Lead Time (days)',color:COLORS.faint,font:{size:10.5}}}, y:{title:{display:true,text:'Avg Defect Rate (%)',color:COLORS.faint,font:{size:10.5}}} } },
    plugins:[quadPlugin]
  });
}

/* ---------------------------------------------------------- */
/* PAGE 4 — MANUFACTURING & QUALITY                              */
/* ---------------------------------------------------------- */
function pageMfg(){
  const p = el(`<section class="page" id="page-mfg">
    <div class="page-head"><h2>Manufacturing & Quality</h2><p class="q">Production efficiency, inspection outcomes and defect drivers.</p></div>
    <div class="section-title">KPIs</div>
    <div class="kpi-row" id="mfg-kpis"></div>
    <div class="section-title">Production & Cost</div>
    <div class="grid cols-2">
      <div class="card"><h3>Production Volume by Supplier</h3><div class="chart-wrap"><canvas id="c-mfg-vol"></canvas></div></div>
      <div class="card"><h3>Manufacturing Cost vs Production Volume</h3><div class="chart-wrap"><canvas id="c-mfg-costvol"></canvas></div></div>
      <div class="card"><h3>Lead Time vs Defect Rate</h3><div class="chart-wrap"><canvas id="c-mfg-leaddefect"></canvas></div></div>
      <div class="card"><h3>Inspection Results</h3><p class="card-sub">Stacked by product type</p><div class="chart-wrap"><canvas id="c-mfg-inspect"></canvas></div></div>
    </div>
    <div class="section-title">Quality Hotspots</div>
    <div class="card"><h3>Heatmap — Supplier × Defect Rate</h3><p class="card-sub">Avg defect rate % by supplier and product type</p><div id="hm-defect"></div></div>
  </section>`);
  main.appendChild(p);

  const kpis = [
    ['Avg Manufacturing Cost', usd1(mean(RAW.map(r=>r[F.mfgCost]))), 'per unit basis', ''],
    ['Total Production Volume', fmt0(totalProdVol)+' units', '', 'accent-blue'],
    ['Avg Manufacturing Lead Time', fmt1(avgMfgLeadTime)+' days', '', 'accent-violet'],
    ['Avg Defect Rate', pct(avgDefect), '', 'accent-red'],
    ['Inspection Pass %', pct(inspectPassPct), `${inspectCounts.Pass} of ${RAW.length} passed`, 'accent-teal']
  ];
  const row = document.getElementById('mfg-kpis');
  kpis.forEach(([label,value,note,cls])=> row.appendChild(el(`<div class="kpi ${cls||''}"><div class="kpi-label">${label}</div><div class="kpi-value">${value}</div>${note?`<div class="kpi-note">${note}</div>`:''}</div>`)));

  const volBySup = SUPPLIERS.map(s=>({label:s, value: sum(RAW.filter(r=>r[F.supplier]===s).map(r=>r[F.prodVol]))}));
  mkChart(document.getElementById('c-mfg-vol'), {
    type:'bar', data:{ labels: volBySup.map(s=>s.label), datasets:[{ data: volBySup.map(s=>s.value), backgroundColor: alpha(COLORS.blue,0.8), borderRadius:5 }]}
  });

  mkChart(document.getElementById('c-mfg-costvol'), {
    type:'scatter', data:{ datasets:[{ label:'SKU', data: RAW.map(r=>({x:r[F.prodVol], y:r[F.mfgCost]})), backgroundColor: alpha(COLORS.gold,0.7), pointRadius:4 }]},
    options:{ scales:{ x:{title:{display:true,text:'Production Volume',color:COLORS.faint,font:{size:10.5}}}, y:{title:{display:true,text:'Manufacturing Cost ($)',color:COLORS.faint,font:{size:10.5}}} } }
  });

  mkChart(document.getElementById('c-mfg-leaddefect'), {
    type:'scatter', data:{ datasets:[{ label:'SKU', data: RAW.map(r=>({x:r[F.mfgLeadTime], y:r[F.defect]})), backgroundColor: alpha(COLORS.red,0.65), pointRadius:4 }]},
    options:{ scales:{ x:{title:{display:true,text:'Manufacturing Lead Time (days)',color:COLORS.faint,font:{size:10.5}}}, y:{title:{display:true,text:'Defect Rate (%)',color:COLORS.faint,font:{size:10.5}}} } }
  });

  const inspectByType = TYPES.map(t=>{
    const rows = RAW.filter(r=>r[F.type]===t);
    return { type:t, Pass: rows.filter(r=>r[F.inspect]==='Pass').length, Fail: rows.filter(r=>r[F.inspect]==='Fail').length, Pending: rows.filter(r=>r[F.inspect]==='Pending').length };
  });
  mkChart(document.getElementById('c-mfg-inspect'), {
    type:'bar',
    data:{ labels: inspectByType.map(t=>t.type), datasets:[
      {label:'Pass', data: inspectByType.map(t=>t.Pass), backgroundColor: COLORS.teal},
      {label:'Fail', data: inspectByType.map(t=>t.Fail), backgroundColor: COLORS.red},
      {label:'Pending', data: inspectByType.map(t=>t.Pending), backgroundColor: COLORS.gold}
    ]},
    options:{ plugins:{legend:{display:true, position:'top'}}, scales:{ x:{stacked:true}, y:{stacked:true} } }
  });

  const hm = document.getElementById('hm-defect');
  let hmHtml = `<div class="heatmap" style="--cols:${TYPES.length}"><div class="hm-row"><div></div>${TYPES.map(t=>`<div class="hm-head">${t}</div>`).join('')}</div>`;
  SUPPLIERS.forEach(s=>{
    hmHtml += `<div class="hm-row"><div class="hm-rowlabel">${s}</div>`;
    TYPES.forEach(t=>{
      const rows = RAW.filter(r=>r[F.supplier]===s && r[F.type]===t);
      const v = rows.length ? mean(rows.map(r=>r[F.defect])) : null;
      const t01 = v===null ? 0 : Math.min(v/5,1);
      hmHtml += `<div class="hm-cell" style="background:${v===null?'transparent':heatColor(t01)}">${v===null?'—':fmt1(v)}</div>`;
    });
    hmHtml += `</div>`;
  });
  hmHtml += `</div>`;
  hm.innerHTML = hmHtml;
}

/* ---------------------------------------------------------- */
/* PAGE 5 — LOGISTICS & DISTRIBUTION                             */
/* ---------------------------------------------------------- */
function pageLog(){
  const p = el(`<section class="page" id="page-log">
    <div class="page-head"><h2>Logistics & Distribution</h2><p class="q">Transportation efficiency across carriers, modes and routes.</p></div>
    <div class="section-title">KPIs</div>
    <div class="kpi-row" id="log-kpis"></div>
    <div class="section-title">Carrier & Mode Performance</div>
    <div class="grid cols-2">
      <div class="card"><h3>Shipping Carrier Comparison</h3><p class="card-sub">Avg cost ($) vs avg time (days)</p><div class="chart-wrap"><canvas id="c-log-carrier"></canvas></div></div>
      <div class="card"><h3>Transportation Mode Analysis</h3><div class="chart-wrap"><canvas id="c-log-mode"></canvas></div></div>
      <div class="card"><h3>Route Analysis</h3><div class="chart-wrap"><canvas id="c-log-route"></canvas></div></div>
      <div class="card"><h3>Cost by Transportation Mode</h3><p class="card-sub">Total logistics cost</p><div class="chart-wrap"><canvas id="c-log-modecost"></canvas></div></div>
    </div>
    <div class="section-title">Distribution Footprint</div>
    <div class="grid cols-2">
      <div class="card"><h3>Shipping Cost Distribution</h3><div class="chart-wrap"><canvas id="c-log-dist"></canvas></div></div>
      <div class="card"><h3>Location-wise Shipment Footprint</h3><p class="card-sub">In place of a geo map — shipment count & avg cost by market</p><div id="log-locations"></div></div>
    </div>
  </section>`);
  main.appendChild(p);

  const kpis = [
    ['Avg Shipping Time', fmt1(avgShipTime)+' days', '', ''],
    ['Avg Shipping Cost', usd1(avgShipCost), '', 'accent-blue'],
    ['Best Carrier Performance', bestCarrier.label, `${fmt1(bestCarrier.time)}d avg`, 'accent-teal'],
    ['Best Route Performance', bestRoute.label, `${fmt1(bestRoute.time)}d avg`, 'accent-gold']
  ];
  const row = document.getElementById('log-kpis');
  kpis.forEach(([label,value,note,cls])=> row.appendChild(el(`<div class="kpi ${cls||''}"><div class="kpi-label">${label}</div><div class="kpi-value">${value}</div>${note?`<div class="kpi-note">${note}</div>`:''}</div>`)));

  mkChart(document.getElementById('c-log-carrier'), {
    type:'bar',
    data:{ labels: carrierStats.map(c=>c.label), datasets:[
      {label:'Avg Cost ($)', data: carrierStats.map(c=>c.cost), backgroundColor: alpha(COLORS.gold,0.85), borderRadius:5, yAxisID:'y'},
      {label:'Avg Time (days)', data: carrierStats.map(c=>c.time), backgroundColor: alpha(COLORS.blue,0.85), borderRadius:5, yAxisID:'y1'}
    ]},
    options:{ plugins:{legend:{display:true, position:'top'}}, scales:{ y:{position:'left'}, y1:{position:'right', grid:{drawOnChartArea:false}} } }
  });

  const modeStats = MODES.map(m=>{ const rows=RAW.filter(r=>r[F.mode]===m); return {label:m, cost: mean(rows.map(r=>r[F.shipCost])), time: mean(rows.map(r=>r[F.shipTime]))}; });
  mkChart(document.getElementById('c-log-mode'), {
    type:'bar', data:{ labels: modeStats.map(m=>m.label), datasets:[{label:'Avg Cost', data: modeStats.map(m=>m.cost), backgroundColor: CAT_COLORS, borderRadius:5}]}
  });

  mkChart(document.getElementById('c-log-route'), {
    type:'bar', data:{ labels: routeStats.map(r=>r.label), datasets:[{label:'Avg Time (days)', data: routeStats.map(r=>r.time), backgroundColor: alpha(COLORS.violet,0.85), borderRadius:5}]}
  });

  const modeCost = MODES.map(m=>({label:m, value: sum(RAW.filter(r=>r[F.mode]===m).map(r=>r[F.cost]))}));
  mkChart(document.getElementById('c-log-modecost'), {
    type:'doughnut', data:{ labels: modeCost.map(m=>m.label), datasets:[{ data: modeCost.map(m=>m.value), backgroundColor: CAT_COLORS, borderColor:COLORS.panel, borderWidth:3 }]},
    options:{ plugins:{legend:{display:true, position:'bottom'}}, cutout:'62%' }
  });

  // shipping cost distribution histogram
  const shipCosts = RAW.map(r=>r[F.shipCost]);
  const binN = 10, mn=Math.min(...shipCosts), mx=Math.max(...shipCosts), w=(mx-mn)/binN;
  const bins = Array.from({length:binN}, (_,i)=>({label:`${fmt0(mn+i*w)}-${fmt0(mn+(i+1)*w)}`, count:0}));
  shipCosts.forEach(v=>{ let i=Math.min(binN-1, Math.floor((v-mn)/w)); bins[i].count++; });
  mkChart(document.getElementById('c-log-dist'), {
    type:'bar', data:{ labels: bins.map(b=>b.label), datasets:[{ data: bins.map(b=>b.count), backgroundColor: alpha(COLORS.teal,0.8), borderRadius:4 }]},
    options:{ scales:{ x:{ ticks:{ font:{size:9} } } } }
  });

  const locStats = LOCATIONS.map(l=>{ const rows=RAW.filter(r=>r[F.loc]===l); return {label:l, n:rows.length, cost: mean(rows.map(r=>r[F.shipCost])), rev: sum(rows.map(r=>r[F.rev]))}; }).sort((a,b)=>b.rev-a.rev);
  const maxN = Math.max(...locStats.map(l=>l.n));
  document.getElementById('log-locations').innerHTML = locStats.map(l=>`
    <div style="display:flex;align-items:center;gap:12px;margin-bottom:10px;">
      <div style="width:82px;font-size:12px;color:var(--ink-dim);">${l.label}</div>
      <div style="flex:1;background:var(--panel-2);border-radius:5px;overflow:hidden;height:16px;">
        <div style="width:${(l.n/maxN*100)}%;height:100%;background:linear-gradient(90deg, var(--gold), var(--teal));"></div>
      </div>
      <div style="width:110px;text-align:right;font-family:var(--mono);font-size:11px;color:var(--ink-faint);">${l.n} SKUs · ${usd1(l.cost)}</div>
    </div>`).join('');
}

/* ---------------------------------------------------------- */
/* PAGE 6 — OPTIMIZATION & RECOMMENDATIONS                       */
/* ---------------------------------------------------------- */
function pageOpt(){
  const p = el(`<section class="page" id="page-opt">
    <div class="page-head"><h2>Optimization & Recommendations</h2><p class="q">MBA-level analysis: Pareto, ABC/XYZ, EOQ, correlation, what-if and root-cause views.</p></div>

    <div class="section-title">AI / Analytics Insights</div>
    <div class="insight-list" id="insights"></div>

    <div class="section-title">Pareto Analysis (80/20)</div>
    <div class="grid cols-2">
      <div class="card"><h3>Suppliers Driving Defects</h3><p class="card-sub">Bars = total defect contribution · line = cumulative %</p><div class="chart-wrap"><canvas id="c-pareto-defect"></canvas></div></div>
      <div class="card"><h3>Suppliers Driving Cost</h3><p class="card-sub">Manufacturing + shipping cost, cumulative %</p><div class="chart-wrap"><canvas id="c-pareto-cost"></canvas></div></div>
    </div>

    <div class="section-title">XYZ Analysis</div>
    <div class="card">
      <p class="card-sub">Demand variability by product type — approximated via coefficient of variation of units sold across SKUs within each category, since the dataset has no time-series demand history.</p>
      <table class="dt" id="t-xyz"></table>
    </div>

    <div class="section-title">EOQ · Safety Stock · Reorder Point</div>
    <div class="card">
      <p class="card-sub">Assumption-based (dataset lacks ordering/holding cost fields): order cost = $${ASSUME_ORDER_COST}/order, holding cost = 20% of unit price/yr, service level Z = ${ASSUME_Z} (~95%). Snapshot demand treated as annual proxy.</p>
      <table class="dt" id="t-eoq"></table>
    </div>

    <div class="section-title">Correlation Analysis</div>
    <div class="grid cols-3">
      <div class="card"><h3>Production Volume ↔ Mfg Cost</h3><div class="chart-wrap short"><canvas id="c-corr1"></canvas></div><p class="footnote">r = <b id="r-1"></b></p></div>
      <div class="card"><h3>Shipping Cost ↔ Shipping Time</h3><div class="chart-wrap short"><canvas id="c-corr2"></canvas></div><p class="footnote">r = <b id="r-2"></b></p></div>
      <div class="card"><h3>Lead Time ↔ Defect Rate</h3><div class="chart-wrap short"><canvas id="c-corr3"></canvas></div><p class="footnote">r = <b id="r-3"></b></p></div>
    </div>

    <div class="section-title">What-if Analysis</div>
    <div class="card">
      <div class="whatif">
        <div>
          <div class="wi-control">
            <label>Shipping cost change <b id="wi-ship-label">0%</b></label>
            <input type="range" id="wi-ship" min="-30" max="30" value="0" step="1">
          </div>
          <div class="wi-result">
            <div class="wr-label">Estimated profit impact</div>
            <div class="wr-value" id="wi-ship-result">$0</div>
          </div>
        </div>
        <div>
          <div class="wi-control">
            <label>Lead time change (days) <b id="wi-lead-label">0</b></label>
            <input type="range" id="wi-lead" min="-10" max="10" value="0" step="1">
          </div>
          <div class="wi-result">
            <div class="wr-label">Estimated total reorder point (all SKUs)</div>
            <div class="wr-value" id="wi-lead-result">—</div>
          </div>
        </div>
      </div>
    </div>

    <div class="section-title">Root Cause Dashboard — High Defect Rate</div>
    <div class="card">
      <div class="rc-controls">
        <select id="rc-supplier"><option value="">All suppliers</option>${SUPPLIERS.map(s=>`<option value="${s}">${s}</option>`).join('')}</select>
      </div>
      <div class="rc-chain" id="rc-chain"></div>
    </div>

    <div class="section-title">KPI Tree</div>
    <div class="card" id="kpi-tree"></div>

    <div class="section-title">MBA Techniques Applied</div>
    <div class="card">
      <div class="insight-list" id="techniques"></div>
    </div>
  </section>`);
  main.appendChild(p);

  // insights
  document.getElementById('insights').innerHTML = buildInsights().map((t,i)=>`<div class="insight"><div class="ii">${String(i+1).padStart(2,'0')}</div><p>${t}</p></div>`).join('');

  // pareto charts
  function paretoChart(canvasId, series){
    mkChart(document.getElementById(canvasId), {
      data:{ labels: series.map(s=>s.label), datasets:[
        {type:'bar', label:'Contribution', data: series.map(s=>s.value), backgroundColor: alpha(COLORS.gold,0.85), borderRadius:4, yAxisID:'y'},
        {type:'line', label:'Cumulative %', data: series.map(s=>s.cumPct), borderColor:COLORS.teal, backgroundColor:COLORS.teal, tension:0.3, yAxisID:'y1', pointRadius:3}
      ]},
      options:{ plugins:{legend:{display:true, position:'top'}}, scales:{ y:{position:'left'}, y1:{position:'right', min:0, max:100, grid:{drawOnChartArea:false}, ticks:{callback:v=>v+'%'}} } }
    });
  }
  paretoChart('c-pareto-defect', paretoDefect);
  paretoChart('c-pareto-cost', paretoCost);

  document.getElementById('t-xyz').innerHTML = `<thead><tr><th>Product Type</th><th>Coefficient of Variation</th><th>Class</th><th>Interpretation</th></tr></thead><tbody>
    ${xyzStats.map(x=>`<tr><td class="name">${x.type}</td><td>${fmt2(x.cv)}</td><td><span class="pill ${x.cls==='X'?'a':(x.cls==='Y'?'b':'c')}">${x.cls}</span></td><td>${x.cls==='X'?'Stable demand — safer for tight inventory':(x.cls==='Y'?'Moderate variability — buffer stock advised':'High variability — build safety stock')}</td></tr>`).join('')}
  </tbody>`;

  document.getElementById('t-eoq').innerHTML = `<thead><tr><th>Product Type</th><th>Annual Demand (proxy)</th><th>EOQ</th><th>Avg Lead Time</th><th>Safety Stock</th><th>Reorder Point</th></tr></thead><tbody>
    ${eoqStats.map(e=>`<tr><td class="name">${e.type}</td><td>${fmt0(e.annualDemand)}</td><td>${fmt0(e.eoq)} units</td><td>${fmt1(e.avgLead)}d</td><td>${fmt0(e.safety)} units</td><td>${fmt0(e.rop)} units</td></tr>`).join('')}
  </tbody>`;

  function corrScatter(id, xs, ys, xl, yl, color){
    mkChart(document.getElementById(id), {
      type:'scatter', data:{ datasets:[{ data: xs.map((x,i)=>({x, y:ys[i]})), backgroundColor: alpha(color,0.65), pointRadius:3 }]},
      options:{ scales:{ x:{title:{display:true,text:xl,color:COLORS.faint,font:{size:9}}}, y:{title:{display:true,text:yl,color:COLORS.faint,font:{size:9}}} } }
    });
  }
  corrScatter('c-corr1', RAW.map(r=>r[F.prodVol]), RAW.map(r=>r[F.mfgCost]), 'Production Vol','Mfg Cost', COLORS.blue);
  corrScatter('c-corr2', RAW.map(r=>r[F.shipCost]), RAW.map(r=>r[F.shipTime]), 'Shipping Cost','Shipping Time', COLORS.gold);
  corrScatter('c-corr3', RAW.map(r=>r[F.supLeadTime]), RAW.map(r=>r[F.defect]), 'Lead Time','Defect Rate', COLORS.red);
  document.getElementById('r-1').textContent = fmt2(corr.volCost) + interp(corr.volCost);
  document.getElementById('r-2').textContent = fmt2(corr.shipCostTime) + interp(corr.shipCostTime);
  document.getElementById('r-3').textContent = fmt2(corr.leadDefect) + interp(corr.leadDefect);
  function interp(r){ const a=Math.abs(r); return a<0.1?' (negligible)':a<0.3?' (weak)':a<0.5?' (moderate)':' (strong)'; }

  // what-if
  const shipSlider = document.getElementById('wi-ship'), leadSlider = document.getElementById('wi-lead');
  function updateWhatIf(){
    const shipPct = +shipSlider.value;
    document.getElementById('wi-ship-label').textContent = (shipPct>0?'+':'')+shipPct+'%';
    const deltaCost = totalShipCost * (shipPct/100);
    const newProfit = -deltaCost;
    const resEl = document.getElementById('wi-ship-result');
    resEl.textContent = (newProfit>=0?'+':'−') + usd(Math.abs(newProfit));
    resEl.className = 'wr-value ' + (newProfit>=0?'up':'down');

    const leadDelta = +leadSlider.value;
    document.getElementById('wi-lead-label').textContent = (leadDelta>0?'+':'')+leadDelta;
    const newRop = sum(eoqStats.map(e=> (e.annualDemand/365) * Math.max(e.avgLead+leadDelta,0) ));
    const baseRop = sum(eoqStats.map(e=> (e.annualDemand/365) * e.avgLead ));
    const ropEl = document.getElementById('wi-lead-result');
    ropEl.innerHTML = `${fmt0(newRop)} units <span style="font-size:11px;color:var(--ink-faint);">(${newRop>=baseRop?'+':''}${fmt0(newRop-baseRop)} vs current)</span>`;
  }
  shipSlider.addEventListener('input', updateWhatIf);
  leadSlider.addEventListener('input', updateWhatIf);
  updateWhatIf();

  // root cause
  function renderRootCause(){
    const supFilter = document.getElementById('rc-supplier').value;
    const rows = supFilter ? RAW.filter(r=>r[F.supplier]===supFilter) : RAW;
    const highDefectRows = [...rows].sort((a,b)=>b[F.defect]-a[F.defect]).slice(0, Math.max(5,Math.round(rows.length*0.2)));
    const avgHigh = mean(highDefectRows.map(r=>r[F.defect]));
    const topSupplier = [...groupBy(highDefectRows, F.supplier).entries()].sort((a,b)=>b[1].length-a[1].length)[0];
    const topFactoryLoc = [...groupBy(highDefectRows, F.loc).entries()].sort((a,b)=>b[1].length-a[1].length)[0];
    const topMode = [...groupBy(highDefectRows, F.mode).entries()].sort((a,b)=>b[1].length-a[1].length)[0];
    const topInspect = [...groupBy(highDefectRows, F.inspect).entries()].sort((a,b)=>b[1].length-a[1].length)[0];
    document.getElementById('rc-chain').innerHTML = `
      <div class="rc-step"><div class="rc-k">High Defect Cohort</div><div class="rc-v">${pct(avgHigh)}</div><div class="rc-v2">${highDefectRows.length} SKUs (top quintile)</div></div>
      <div class="rc-step"><div class="rc-k">Leading Supplier</div><div class="rc-v">${topSupplier[0]}</div><div class="rc-v2">${topSupplier[1].length} of ${highDefectRows.length} SKUs</div></div>
      <div class="rc-step"><div class="rc-k">Leading Market</div><div class="rc-v">${topFactoryLoc[0]}</div><div class="rc-v2">${topFactoryLoc[1].length} SKUs</div></div>
      <div class="rc-step"><div class="rc-k">Leading Transport Mode</div><div class="rc-v">${topMode[0]}</div><div class="rc-v2">${topMode[1].length} SKUs</div></div>
      <div class="rc-step"><div class="rc-k">Inspection Outcome</div><div class="rc-v">${topInspect[0]}</div><div class="rc-v2">${topInspect[1].length} SKUs</div></div>
    `;
  }
  document.getElementById('rc-supplier').addEventListener('change', renderRootCause);
  renderRootCause();

  // KPI tree
  const avgPrice = mean(RAW.map(r=>r[F.price]));
  document.getElementById('kpi-tree').innerHTML = `
    <div class="tree">
      <div class="tree-root"><div class="tk">Revenue</div><div class="tv">${usd(totalRevenue)}</div></div>
      <div class="tree-connector"></div>
      <div class="tree-level">
        <div class="tree-node"><div class="tk">Products Sold</div><div class="tv">${fmt0(totalSold)}</div></div>
        <div class="tree-node"><div class="tk">Avg Price</div><div class="tv">${usd1(avgPrice)}</div></div>
        <div class="tree-node"><div class="tk">Product Types</div><div class="tv">${TYPES.length}</div></div>
      </div>
      <div class="tree-connector"></div>
      <div class="tree-level">
        <div class="tree-node"><div class="tk">Profitability (Approx.)</div><div class="tv">${pct(approxMargin)}</div></div>
      </div>
      <div class="tree-connector"></div>
      <div class="tree-level">
        <div class="tree-node"><div class="tk">Manufacturing Cost</div><div class="tv">${usd(totalMfgCost)}</div></div>
        <div class="tree-node"><div class="tk">Shipping Cost</div><div class="tv">${usd(totalShipCost)}</div></div>
        <div class="tree-node"><div class="tk">Logistics Cost</div><div class="tv">${usd(sum(RAW.map(r=>r[F.cost])))}</div></div>
      </div>
      <div class="tree-connector"></div>
      <div class="tree-level">
        <div class="tree-node"><div class="tk">Operational Efficiency</div><div class="tv">${fmt1(avgLeadTime)}d lead</div></div>
      </div>
      <div class="tree-connector"></div>
      <div class="tree-level">
        <div class="tree-node"><div class="tk">Production</div><div class="tv">${fmt0(totalProdVol)}</div></div>
        <div class="tree-node"><div class="tk">Inspection Pass</div><div class="tv">${pct(inspectPassPct)}</div></div>
        <div class="tree-node"><div class="tk">Transportation Modes</div><div class="tv">${MODES.length}</div></div>
      </div>
    </div>`;

  const techniques = [
    'SCOR (Supply Chain Operations Reference) Model — page structure mirrors Plan/Source/Make/Deliver stages.',
    'ABC Inventory Analysis — SKUs classified by revenue contribution (Inventory page).',
    'Pareto Analysis (80/20) — supplier concentration of defects and cost (this page).',
    'Supplier Performance Scorecard — weighted composite ranking (Supplier page).',
    'Inventory Turnover Analysis — turnover, days-of-inventory, fill rate (Inventory page).',
    'Cost-to-Serve — shipping and logistics cost broken out by carrier, mode, and route (Logistics page).',
    'Lead Time Analysis — order and manufacturing lead time tracked across every page.',
    'Root Cause Analysis — drill-through from high defect rate to supplier / market / mode (this page).',
    'Demand vs Inventory Analysis — stock vs sales scatter and bubble chart (Inventory page).',
    'Capacity Utilization — production volume by supplier (Manufacturing page).',
    'Defect Rate & Quality Control — inspection pass/fail/pending tracking (Manufacturing page).',
    'Transportation Efficiency Analysis — carrier and route comparison (Logistics page).'
  ];
  document.getElementById('techniques').innerHTML = techniques.map(t=>`<div class="insight"><div class="ii">✓</div><p>${t}</p></div>`).join('');
}

/* ---------------------------------------------------------- */
pageExec(); pageInv(); pageSup(); pageMfg(); pageLog(); pageOpt();
document.getElementById('page-exec').classList.add('active');

} // end initDashboard(RAW)

/* ============================================================
   DATA LOADER — reads an .xlsx workbook client-side (SheetJS),
   converts the first sheet to row objects, trims header keys,
   and hands the array to initDashboard(). No data ever leaves
   the browser.
   ============================================================ */
const DEFAULT_FILE = 'supplyChainDataSet.xlsx'; // sits next to this HTML file, used when served over http(s)
const loaderEl = document.getElementById('loader');
const statusEl = document.getElementById('loader-status');
const dropZone = document.getElementById('drop-zone');
const fileInput = document.getElementById('file-input');

function setStatus(msg, isErr){
  statusEl.textContent = msg;
  statusEl.classList.toggle('err', !!isErr);
}

function rowsFromWorkbook(workbook){
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: null });
  // trim header whitespace, in case the source workbook has stray spaces
  return rows.map(row=>{
    const clean = {};
    Object.keys(row).forEach(k=> clean[k.trim()] = row[k]);
    return clean;
  });
}

function launchDashboard(rows){
  if(!rows || !rows.length){ setStatus('That file parsed but contained no rows.', true); return; }
  try{
    initDashboard(rows);
    loaderEl.classList.add('hidden');
    document.getElementById('app').style.display = '';
  }catch(e){
    console.error(e);
    setStatus('Could not build the dashboard from this file — check it matches the expected column layout.', true);
  }
}

function loadFromArrayBuffer(buf){
  try{
    const wb = XLSX.read(buf, { type:'array' });
    launchDashboard(rowsFromWorkbook(wb));
  }catch(e){
    console.error(e);
    setStatus('Could not read that file as an Excel workbook.', true);
  }
}

function loadFromFile(file){
  setStatus('Reading ' + file.name + '…');
  const reader = new FileReader();
  reader.onload = e => loadFromArrayBuffer(new Uint8Array(e.target.result));
  reader.onerror = () => setStatus('Could not read that file.', true);
  reader.readAsArrayBuffer(file);
}

// 1) try to auto-fetch the workbook if this page is served over http(s)
fetch(DEFAULT_FILE).then(res=>{
  if(!res.ok) throw new Error('not found');
  return res.arrayBuffer();
}).then(buf=>{
  loadFromArrayBuffer(new Uint8Array(buf));
}).catch(()=>{
  // 2) fall back to manual drag-and-drop / file picker (needed when opened directly via file://)
  setStatus('Auto-load unavailable here — drop the workbook below or click to browse.');
});

// manual picker wiring (always active, in case auto-fetch loaded the wrong / stale file)
dropZone.addEventListener('click', ()=> fileInput.click());
fileInput.addEventListener('change', e=>{ if(e.target.files[0]) loadFromFile(e.target.files[0]); });
['dragenter','dragover'].forEach(ev=> dropZone.addEventListener(ev, e=>{ e.preventDefault(); dropZone.classList.add('drag'); }));
['dragleave','drop'].forEach(ev=> dropZone.addEventListener(ev, e=>{ e.preventDefault(); dropZone.classList.remove('drag'); }));
dropZone.addEventListener('drop', e=>{ const f = e.dataTransfer.files[0]; if(f) loadFromFile(f); });
