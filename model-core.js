// Shared case data and financial calculations.
const RAW = {"years":[2010,2011,2012,2013,2014,2015,2016,2017,2018,2019,2020],"clothing":{"revenue":[0,4500,6860,8409,9082,9808,10593,11440,12355,13344,14411],"fixed":[0,575,575,587,598,610,622,635,648,660,674],"variable":[0,2035,3404,4291,4669,5078,5521,6000,6519,7079,7685],"depreciation":[0,152,152,152,152,164,178,192,207,224,242],"sga":[1250,1155,1735,2102,2270,2452,2648,2860,3089,3336,3603],"totalProduction":[0,2762,4131,5029,5419,5853,6321,6827,7373,7963,8600],"totalOpEx":[1250,3917,5866,7132,7690,8305,8969,9687,10462,11299,12203],"reportedEbit":[-1250,583,994,1277,1392,1503,1623,1753,1893,2045,2209],"cashPct":[0,0.03,0.03,0.03,0.03,0.03,0.03,0.03,0.03,0.03,0.03],"dso":[0,59.2,59.2,59.2,59.2,59.2,59.2,59.2,59.2,59.2,59.2],"inventoryTurnover":[0,7.7,8.3,12.7,12.7,12.7,12.7,12.7,12.7,12.7,12.7],"dpo":[0,30.8,30.9,31,31,31,31,31,31,31,31],"capex":[1470,952,152,152,334,361,389,421,454,491,530],"startupNwc":[800,null,null,null,null,null,null,null,null,null,null]},"custom":{"revenue":[0,0,6000,14360,20222,21435,22721,24084,25529,27061,28685],"fixed":[0,0,1650,1683,1717,1751,1786,1822,1858,1895,1933],"variable":[0,0,2250,7651,11427,12182,12983,13833,14736,15694,16712],"depreciation":[0,0,310,310,310,436,462,490,520,551,584],"sga":[1201,0,1240,2922,4044,4287,4544,4817,5106,5412,5737],"totalProduction":[0,0,4210,9644,13454,14369,15231,16145,17113,18140,19229],"totalOpEx":[1201,0,5450,12566,17498,18656,19775,20962,22219,23553,24966],"reportedEbit":[-1201,0,550,1794,2724,2779,2946,3123,3310,3509,3719],"cashPct":[0,0,0.03,0.03,0.03,0.03,0.03,0.03,0.03,0.03,0.03],"dso":[0,0,59.2,59.2,59.2,59.2,59.2,59.2,59.2,59.2,59.2],"inventoryTurnover":[0,0,12.2,12.3,12.6,12.7,12.7,12.7,12.7,12.7,12.7],"dpo":[0,0,33.7,33.8,33.9,33.9,33.9,33.9,33.9,33.9,33.9],"capex":[4610,0,310,310,2192,826,875,928,983,1043,1105],"startupNwc":[0,1000,null,null,null,null,null,null,null,null,null]},"caseFacts":{"industrySales2008Bn":42,"industryGrowth":0.046,"dollSales2008Bn":3.1,"dollGrowth":0.03,"taxRate":0.4,"itCost2011":435,"clothingRate":0.084,"customRate":0.09,"lowRiskRate":0.077}};
const EXCEL_BASE = {"clothing":{"initialFCF":-3020,"terminalFCF":883.4845308920296,"terminalValue":16360.824646148694,"pvTerminal":7303.186107170744,"npv":7171.855932412575,"irr":0.240524424032673,"payback":7.375015171547592,"pi":3.3747867325869456,"eaa":1088.1807714656518,"pvTerminalToNPV":1.0183118813311118,"discountedPayback":"Beyond 2020"},"custom":{"initialFCF":-5330.6,"terminalFCF":1487.0803302987806,"terminalValue":24784.672171646343,"pvTerminal":10469.313370670256,"npv":7350.920187847349,"irr":0.18396401063490048,"payback":9.976973935701606,"pi":2.3790042749122704,"eaa":1145.421044584505,"pvTerminalToNPV":1.424218071089696,"discountedPayback":"Beyond 2020"}};
const BASE = {tax:0.40,days:365,taxMode:1,itShare:0,clothingRate:0.084,customRate:0.09,clothingG:0.02,customG:0.02,revDelta:0,marginDelta:0,capexMult:1,nwcMult:1,terminalMethod:0};
const EXCEL_PARAMS={...BASE,clothingG:.03,customG:.03};
const PARAM_META = {
  clothingRate:{label:'Clothing 贴现率',fmt:v=>(v*100).toFixed(1)+'%'}, customRate:{label:'Custom 贴现率',fmt:v=>(v*100).toFixed(1)+'%'},
  clothingG:{label:'Clothing 永续增长率',fmt:v=>(v*100).toFixed(1)+'%'}, customG:{label:'Custom 永续增长率',fmt:v=>(v*100).toFixed(1)+'%'},
  revDelta:{label:'收入增长率变化',fmt:v=>(v*100).toFixed(1)+' pp'}, marginDelta:{label:'营业利润率变化',fmt:v=>(v*100).toFixed(1)+' pp'},
  capexMult:{label:'资本支出倍数',fmt:v=>v.toFixed(2)+'×'}, nwcMult:{label:'营运资本需求倍数',fmt:v=>v.toFixed(2)+'×'}, itShare:{label:'IT 成本计入比例',fmt:v=>(v*100).toFixed(0)+'%'},
  days:{label:'天数口径',fmt:v=>v+' 天'}, taxMode:{label:'亏损抵税',fmt:v=>v===1?'当期抵税':'亏损结转'}, terminalMethod:{label:'终值方法',fmt:v=>v===0?'直接增长末年 FCF':'规范化营运资本'}
};
let state={...BASE};
let scenarioName='Base';
let flowProject='clothing';
let flowYearIndex=3;
let auditProject='clothing';
let heatProject='growth';
let sensitivityState={...BASE,clothingRate:.09,customRate:.09};
let activeFlowStep=6;
let modelMode=false;
const YEARS=RAW.years;
const $=id=>document.getElementById(id);
const clamp=(v,min,max)=>Math.min(max,Math.max(min,v));
const fmtPct=v=>Number.isFinite(v)?(v*100).toFixed(2)+'%':'N/A';
const fmtPct1=v=>Number.isFinite(v)?(v*100).toFixed(1)+'%':'N/A';
const fmtM=v=>Number.isFinite(v)?(v<0?'−':'')+'$'+Math.abs(v/1000).toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2})+'M':'N/A';
const fmtK=v=>Number.isFinite(v)?(v<0?'−':'')+'$'+Math.abs(v).toLocaleString('en-US',{minimumFractionDigits:1,maximumFractionDigits:1})+'K':'N/A';
const fmtNum=v=>Number.isFinite(v)?v.toLocaleString('en-US',{minimumFractionDigits:1,maximumFractionDigits:1}):'N/A';
const projectName=p=>p==='clothing'?'Clothing':'Custom';
function safeNum(v){return Number.isFinite(v)?v:null}
function calcProject(project, p=state, overrides={}){
  const d=RAW[project];
  const rate=overrides.rate!==undefined?overrides.rate:p[project+'Rate'];
  const g=overrides.g!==undefined?overrides.g:p[project+'G'];
  const revDelta=p.revDelta, marginDelta=p.marginDelta, capexMult=p.capexMult, nwcMult=p.nwcMult;
  const days=p.days, tax=p.tax, taxMode=p.taxMode, itShare=p.itShare, terminalMethod=p.terminalMethod;
  const revenue=[];
  for(let i=0;i<YEARS.length;i++){
    const b=d.revenue[i];
    if(b===0 || b===null){revenue.push(0);continue;}
    if(i>0 && d.revenue[i-1]>0){const baseGrowth=b/d.revenue[i-1]-1;revenue.push(revenue[i-1]*(1+baseGrowth+revDelta));}
    else revenue.push(b);
  }
  const fixed=[],variable=[],depreciation=[],sga=[],capex=[],totalProduction=[],totalOpEx=[],ebit=[],taxes=[],nopat=[],openingLoss=[],lossUsed=[],closingLoss=[];
  let losses=0;
  for(let i=0;i<YEARS.length;i++){
    const ratio=d.revenue[i]>0?revenue[i]/d.revenue[i]:1;
    fixed.push(d.fixed[i]);
    variable.push(d.variable[i]*ratio);
    depreciation.push(d.depreciation[i]*capexMult);
    sga.push(d.sga[i]*ratio);
    capex.push(d.capex[i]*capexMult);
    // Preserve source rounding in the published total cost lines at the baseline.
    totalProduction.push(d.totalProduction[i]+d.variable[i]*(ratio-1)+d.depreciation[i]*(capexMult-1));
    totalOpEx.push(d.totalOpEx[i]+d.variable[i]*(ratio-1)+d.sga[i]*(ratio-1)+d.depreciation[i]*(capexMult-1));
    const itCost=(project==='custom' && YEARS[i]===2011)?435*itShare:0;
    ebit.push(revenue[i]-fixed[i]-variable[i]-depreciation[i]-sga[i]+revenue[i]*marginDelta-itCost);
    openingLoss.push(losses);
    let used=0, tx=0, closing=0;
    if(taxMode===1){tx=ebit[i]*tax;}
    else{used=Math.min(losses,Math.max(ebit[i],0));tx=(Math.max(ebit[i],0)-used)*tax;closing=losses-used+Math.max(-ebit[i],0);losses=closing;}
    lossUsed.push(used);closingLoss.push(closing);taxes.push(tx);nopat.push(ebit[i]-tx);
  }
  const cash=[],ar=[],inventory=[],ap=[],startup=[],nwc=[],deltaNwc=[];
  for(let i=0;i<YEARS.length;i++){
    const hasStartup=d.startupNwc[i]!==null && d.startupNwc[i]!==undefined;
    if(hasStartup){cash.push(0);ar.push(0);inventory.push(0);ap.push(0);startup.push(d.startupNwc[i]*nwcMult);}
    else{
      cash.push(revenue[i]*d.cashPct[i]*nwcMult);
      ar.push(revenue[i]/days*d.dso[i]*nwcMult);
      inventory.push((d.inventoryTurnover[i] ? totalProduction[i]/d.inventoryTurnover[i] : 0)*nwcMult);
      ap.push(-(totalOpEx[i]/days)*d.dpo[i]*nwcMult);
      startup.push(0);
    }
    nwc.push(cash[i]+ar[i]+inventory[i]+ap[i]+startup[i]);
    deltaNwc.push(i===0?nwc[i]:nwc[i]-nwc[i-1]);
  }
  const fcf=YEARS.map((_,i)=>nopat[i]+depreciation[i]-capex[i]-deltaNwc[i]);
  const terminalFcf=terminalMethod===0?fcf[fcf.length-1]*(1+g):(nopat[nopat.length-1]+depreciation[depreciation.length-1]-capex[capex.length-1])*(1+g)-nwc[nwc.length-1]*g;
  const validTerminal=Number.isFinite(rate)&&Number.isFinite(g)&&rate>g+1e-10;
  const tv=validTerminal?terminalFcf/(rate-g):null;
  const discount=YEARS.map((_,i)=>1/Math.pow(1+rate,i));
  const pvFcf=fcf.map((v,i)=>v*discount[i]);
  const pvTv=validTerminal?tv*discount[discount.length-1]:null;
  const npv=validTerminal?pvFcf.reduce((a,b)=>a+b,0)+pvTv:null;
  const irrFlows=fcf.slice(); if(validTerminal)irrFlows[irrFlows.length-1]+=tv;
  function npvAt(x){return irrFlows.reduce((s,v,i)=>s+v/Math.pow(1+x,i),0)}
  let irr=null;
  if(validTerminal){let x=.1;for(let j=0;j<250;j++){const fx=npvAt(x);let der=0;for(let i=1;i<irrFlows.length;i++)der+=-i*irrFlows[i]/Math.pow(1+x,i+1);if(!Number.isFinite(fx)||!Number.isFinite(der)||Math.abs(der)<1e-12)break;const nx=x-fx/der;if(!Number.isFinite(nx)||nx<=-.999||nx>100)break;if(Math.abs(nx-x)<1e-12){x=nx;break;}x=nx;}if(Number.isFinite(x)&&Math.abs(npvAt(x))<1e-5)irr=x;}
  const pi=validTerminal&&fcf[0]<0?1+npv/(-fcf[0]):null;
  function getPayback(flows){let running=0;const cumulative=[];for(const f of flows){running+=f;cumulative.push(running);}if(cumulative[0]>=0)return {years:0,cumulative};for(let i=1;i<cumulative.length;i++){if(cumulative[i]>=0&&cumulative[i-1]<0&&flows[i]>0)return {years:(i-1)+(-cumulative[i-1]/flows[i]),cumulative};}return {years:null,cumulative};}
  const pb=getPayback(fcf), dpb=getPayback(pvFcf);
  const eaa=validTerminal?(rate===0?npv/10:npv*rate/(1-Math.pow(1+rate,-10))):null;
  return {project,rate,g,revenue,fixed,variable,depreciation,sga,capex,totalProduction,totalOpEx,ebit,taxes,nopat,openingLoss,lossUsed,closingLoss,cash,ar,inventory,ap,startup,nwc,deltaNwc,fcf,discount,pvFcf,cumulativeFcf:pb.cumulative,cumulativePvFcf:dpb.cumulative,terminalFcf,tv,pvTv,npv,irr,pi,payback:pb.years,discountedPayback:dpb.years,eaa,pvTerminalToNpv:validTerminal&&npv!==0?pvTv/npv:null,validTerminal,irrFlows};
}
function getModels(p=state){return {clothing:calcProject('clothing',p),custom:calcProject('custom',p)}}
function scenarioParams(name){const p={...BASE};if(name==='Bull'){p.clothingG=.03;p.customG=.03;}if(name==='Bear'){p.clothingG=.01;p.customG=.01;p.itShare=1;}return p;}
