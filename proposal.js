/* Cleanearth Site Audit – Word documents (audit proposal and pre-audit report).
   Loaded on demand after vendor/docx.umd.js. Uses the app's own helpers (S, ans, flags, ta, SECTIONS, REGS…)
   and the company settings in S.branding (logo, letterhead, footer, company details). */
(function(){
  'use strict';
  const D=()=>window.docx;
  const TBC='[to be confirmed]';
  const dateTxt=d=>d?fmtDate(d).replace(/^\w+, /,''):TBC;
  const gh=x=>x===null||x===undefined||!isFinite(x)?'–':'GHS '+Math.round(x).toLocaleString('en-GB');
  const qf=x=>x===null||x===undefined?'':Number.isInteger(x)?x.toLocaleString('en-GB'):x.toLocaleString('en-GB',{maximumFractionDigits:2});
  const b64=src=>{const s=atob(String(src).split(',')[1]||'');const u=new Uint8Array(s.length);for(let i=0;i<s.length;i++)u[i]=s.charCodeAt(i);return u};
  const hex=c=>/^#?[0-9a-f]{6}$/i.test(c||'')?String(c).replace('#','').toUpperCase():null;

  /* Built-in Cleanearth letterhead (brand/ folder); anything uploaded in Company settings replaces it. */
  const BUILTIN={header:{url:'brand/header.png',w:974,h:150},footer:{url:'brand/footer.jpg',w:1536,h:293},logo:{url:'brand/logo.png',w:880,h:176}};
  window.CE_BRAND_DEFAULTS={company_name:'Cleanearth',unit:'Water Services',address:'P.O. Box AN 5607 Accra · 7 Neem Close, Comm 18, Lashibi, Tema, Ghana',phone:'+233 (0) 303 403 997 · +233 (0) 203 360 566',email:'info@cleanearthsi.com',website:'www.cleanearth.africa',color:'#233F8C',accent:'#A9D27F'};
  async function asset(stored,key){
    if(stored&&stored.src&&stored.w)return {data:b64(stored.src),w:stored.w,h:stored.h};
    const b=BUILTIN[key];if(!b)return null;
    try{const r=await fetch(b.url);if(!r.ok)return null;return {data:new Uint8Array(await r.arrayBuffer()),w:b.w,h:b.h}}catch(e){return null}
  }
  /* ---------- shared kit: styles, tables, cover, header and footer ---------- */
  async function makeKit(){
    const {Paragraph,TextRun,Table,TableRow,TableCell,WidthType,ShadingType,AlignmentType,BorderStyle,Header,Footer,PageNumber,HeadingLevel,ImageRun,HorizontalPositionRelativeFrom,VerticalPositionRelativeFrom,HorizontalPositionAlign,VerticalPositionAlign,TextWrappingType,VerticalAlign}=D();
    const DEF=window.CE_BRAND_DEFAULTS;const BR={...DEF,...Object.fromEntries(Object.entries(S.branding||{}).filter(([,v])=>v!==''&&v!==null&&v!==undefined))};
    const NAVY=hex(BR.color)||'233F8C',ACCENT=hex(BR.accent)||'A9D27F',GREY='6B7F96',LIGHT='E8EEF1',FONT=(BR.font||'Candara').trim()||'Candara',W=9026;
    const company=BR.company_name;const unit=BR.unit||'';
    const A={header:await asset((S.branding||{}).header_image,'header'),footer:await asset((S.branding||{}).footer_image,'footer'),logo:await asset((S.branding||{}).logo,'logo'),client:await asset((S.audit||{}).client_logo,'none')};
    const letterhead=(BR.header_mode||'Letterhead image')==='Letterhead image'&&A.header;
    const P=(t,o={})=>new Paragraph({spacing:{after:120,line:276},...o,children:[].concat(t).map(x=>typeof x==='string'?new TextRun(x):x)});
    const H=t=>new Paragraph({heading:HeadingLevel.HEADING_1,keepNext:true,spacing:{before:280,after:120},children:[new TextRun(t)]});
    const H2=t=>new Paragraph({heading:HeadingLevel.HEADING_2,keepNext:true,spacing:{before:180,after:80},children:[new TextRun(t)]});
    const bul=t=>new Paragraph({numbering:{reference:'b',level:0},spacing:{after:60},children:[].concat(t).map(x=>typeof x==='string'?new TextRun(x):x)});
    const B_=t=>new TextRun({text:t,bold:true});
    const line={style:BorderStyle.SINGLE,size:4,color:'A9B8C9'};const borders={top:line,bottom:line,left:line,right:line};
    const none={style:BorderStyle.NONE,size:0,color:'FFFFFF'};const nb={top:none,bottom:none,left:none,right:none};
    function table(cols,rows,opt={}){
      const sz=opt.size||19;
      const cell=(txt,i,ri)=>{const hdr=ri===0&&!opt.noHeader;const right=opt.right&&opt.right.includes(i)&&!hdr;const bold=hdr||(opt.boldFirst&&i===0)||(opt.boldRows&&opt.boldRows.includes(ri));
        const lines=String(txt??'').split('\n');
        return new TableCell({width:{size:cols[i],type:WidthType.DXA},borders,margins:{top:60,bottom:60,left:100,right:100},
          shading:hdr?{type:ShadingType.CLEAR,fill:NAVY,color:'auto'}:(opt.shadeFirst&&i===0)||(opt.boldRows&&opt.boldRows.includes(ri))?{type:ShadingType.CLEAR,fill:LIGHT,color:'auto'}:undefined,
          children:lines.map(l=>new Paragraph({alignment:right?AlignmentType.RIGHT:AlignmentType.LEFT,children:[new TextRun({text:l,bold,color:hdr?'FFFFFF':undefined,size:sz})]}))})};
      return new Table({width:{size:W,type:WidthType.DXA},columnWidths:cols,rows:rows.map((r,ri)=>new TableRow({tableHeader:ri===0&&!opt.noHeader,cantSplit:true,children:r.map((c,i)=>cell(c,i,ri))}))});
    }
    const gap=()=>new Paragraph({spacing:{after:80},children:[]});
    const fit=(im,maxW,maxH)=>{let w=maxW,h=Math.round(maxW*im.h/im.w);if(h>maxH){h=maxH;w=Math.round(maxH*im.w/im.h)}return {width:w,height:h}};
    const img=(im,maxW,maxH)=>im?new ImageRun({data:im.data,transformation:fit(im,maxW,maxH)}):null;

    function header(){
      if(letterhead){
        /* same placement as the Cleanearth letterhead: banner at the top right of the page */
        return new Header({children:[new Paragraph({children:[new ImageRun({data:A.header.data,transformation:{width:623,height:Math.round(623*A.header.h/A.header.w)},
          floating:{horizontalPosition:{relative:HorizontalPositionRelativeFrom.PAGE,align:HorizontalPositionAlign.RIGHT},verticalPosition:{relative:VerticalPositionRelativeFrom.PAGE,align:VerticalPositionAlign.TOP},wrap:{type:TextWrappingType.TOP_AND_BOTTOM},allowOverlap:true}})]})]});
      }
      const logo=img(A.logo,170,48);
      const contact=[BR.phone,BR.email,BR.website].filter(Boolean).join('  |  ');
      const right=[new Paragraph({alignment:AlignmentType.RIGHT,children:[new TextRun({text:company,bold:true,size:22,color:NAVY})]}),
        ...(unit?[new Paragraph({alignment:AlignmentType.RIGHT,children:[new TextRun({text:unit,size:17,color:GREY})]})]:[]),
        ...(BR.address?[new Paragraph({alignment:AlignmentType.RIGHT,children:[new TextRun({text:BR.address,size:15,color:GREY})]})]:[]),
        ...(contact?[new Paragraph({alignment:AlignmentType.RIGHT,children:[new TextRun({text:contact,size:15,color:GREY})]})]:[])];
      const t=new Table({width:{size:W,type:WidthType.DXA},columnWidths:[3200,W-3200],borders:{top:none,bottom:none,left:none,right:none,insideHorizontal:none,insideVertical:none},rows:[new TableRow({children:[
        new TableCell({width:{size:3200,type:WidthType.DXA},borders:nb,verticalAlign:VerticalAlign.CENTER,children:[new Paragraph({children:logo?[logo]:[new TextRun({text:company.toUpperCase(),bold:true,size:24,color:NAVY,characterSpacing:30})]})]}),
        new TableCell({width:{size:W-3200,type:WidthType.DXA},borders:nb,children:right})]})]});
      return new Header({children:[t,new Paragraph({spacing:{after:160},border:{bottom:{style:BorderStyle.SINGLE,size:8,color:NAVY,space:4}},children:[]})]});
    }
    function footer(ref,withPage=true){
      const pg=withPage?[new TextRun({text:(ref?ref+'   |   ':'')+'Page ',size:15,color:GREY}),new TextRun({children:[PageNumber.CURRENT],size:15,color:GREY})]:[];
      if(A.footer){
        /* full-width strip across the bottom of the page, as on the letterhead */
        return new Footer({children:[new Paragraph({alignment:AlignmentType.RIGHT,children:[...pg,new ImageRun({data:A.footer.data,transformation:{width:794,height:104},
          floating:{horizontalPosition:{relative:HorizontalPositionRelativeFrom.PAGE,align:HorizontalPositionAlign.LEFT},verticalPosition:{relative:VerticalPositionRelativeFrom.PAGE,align:VerticalPositionAlign.BOTTOM},wrap:{type:TextWrappingType.NONE},allowOverlap:true}})]})]});
      }
      const left=BR.footer_text||[company+(unit?' '+unit:''),BR.reg_no].filter(Boolean).join('  |  ');
      const kids=[new Paragraph({border:{top:{style:BorderStyle.SINGLE,size:4,color:'A9B8C9',space:4}},children:[new TextRun({text:left,size:15,color:GREY})]})];
      if(withPage)kids.push(new Paragraph({alignment:AlignmentType.RIGHT,children:pg}));
      return new Footer({children:kids});
    }
    /* minimal cover: both logos, document type, title, a short accent rule, client, and the key details */
    function cover(docType,title,meta){
      const kids=[];const a=S.audit||{};
      const nob={top:none,bottom:none,left:none,right:none,insideHorizontal:none,insideVertical:none};
      const cl=img(A.logo,200,44),cc=A.client?img(A.client,190,70):null;const half=Math.floor(W/2);
      const cell=(c,al)=>new TableCell({width:{size:half,type:WidthType.DXA},borders:nb,verticalAlign:VerticalAlign.CENTER,children:[new Paragraph({alignment:al,children:c?[c]:[]})]});
      kids.push(new Table({width:{size:W,type:WidthType.DXA},columnWidths:[half,W-half],borders:nob,
        rows:[new TableRow({children:[cell(cl||new TextRun({text:company,bold:true,size:28,color:NAVY}),AlignmentType.LEFT),cell(cc,AlignmentType.RIGHT)]})]}));
      kids.push(new Paragraph({spacing:{before:3800,after:160},children:[new TextRun({text:docType.toUpperCase(),size:18,bold:true,color:GREY,characterSpacing:60})]}));
      kids.push(new Paragraph({spacing:{after:240,line:250},children:[new TextRun({text:title,size:52,bold:true,color:NAVY})]}));
      kids.push(new Table({width:{size:1100,type:WidthType.DXA},columnWidths:[1100],borders:nob,rows:[new TableRow({height:{value:60,rule:'exact'},children:[new TableCell({width:{size:1100,type:WidthType.DXA},borders:nb,shading:{type:ShadingType.CLEAR,fill:ACCENT,color:'auto'},children:[new Paragraph({children:[]})]})]})]}));
      kids.push(new Paragraph({spacing:{before:360,after:60},children:[new TextRun({text:a.client||TBC,size:32,color:NAVY})]}));
      const where=[a.site,a.location].filter(Boolean).join('  ·  ');if(where)kids.push(new Paragraph({spacing:{after:60},children:[new TextRun({text:where,size:22,color:GREY})]}));
      kids.push(new Paragraph({spacing:{before:3200},children:[]}));
      kids.push(new Table({width:{size:W,type:WidthType.DXA},columnWidths:[2200,W-2200],borders:nob,
        rows:meta.map(([k,v])=>new TableRow({children:[
          new TableCell({width:{size:2200,type:WidthType.DXA},borders:nb,margins:{top:50,bottom:50},children:[new Paragraph({children:[new TextRun({text:k.toUpperCase(),size:15,bold:true,color:GREY,characterSpacing:30})]})]}),
          new TableCell({width:{size:W-2200,type:WidthType.DXA},borders:nb,margins:{top:50,bottom:50},children:[new Paragraph({children:[new TextRun({text:String(v||''),size:19,color:'1F2A33'})]})]})]}))}));
      return kids;
    }
    function document(title,ref,coverKids,children){
      const {Document,LevelFormat}=D();
      const lh=!!letterhead,fA=!!A.footer;
      const margin={top:lh?1750:1200,bottom:fA?2150:1100,left:1440,right:1440,header:lh?300:500,footer:fA?1750:400};
      const size={width:11906,height:16838};
      return new Document({creator:company+(unit?' '+unit:''),title,
        styles:{default:{document:{run:{font:FONT,size:21}}},
          paragraphStyles:[
            {id:'Heading1',name:'Heading 1',basedOn:'Normal',next:'Normal',quickFormat:true,run:{size:26,bold:true,color:NAVY,font:FONT},paragraph:{spacing:{before:280,after:120},outlineLevel:0}},
            {id:'Heading2',name:'Heading 2',basedOn:'Normal',next:'Normal',quickFormat:true,run:{size:22,bold:true,color:NAVY,font:FONT},paragraph:{spacing:{before:180,after:80},outlineLevel:1}}]},
        numbering:{config:[{reference:'b',levels:[{level:0,format:LevelFormat.BULLET,text:'•',alignment:AlignmentType.LEFT,style:{paragraph:{indent:{left:540,hanging:300}}}}]}]},
        sections:[
          {properties:{page:{size,margin:{...margin,top:1300,header:300}}},headers:{default:new Header({children:[new Paragraph({children:[]})]})},footers:{default:footer('',false)},children:coverKids},
          {properties:{page:{size,margin,pageNumbers:{start:1}}},headers:{default:header()},footers:{default:footer(ref)},children}]});
    }
    /* photographs */
    async function imgData(id,maxPx=1100){
      const url=await CE.photoURL(id);
      const im=await new Promise((res,rej)=>{const i=new Image();i.onload=()=>res(i);i.onerror=rej;i.src=url});
      const sc=Math.min(1,maxPx/Math.max(im.naturalWidth,im.naturalHeight));
      const c=window.document.createElement('canvas');c.width=Math.max(1,Math.round(im.naturalWidth*sc));c.height=Math.max(1,Math.round(im.naturalHeight*sc));
      c.getContext('2d').drawImage(im,0,0,c.width,c.height);
      const blob=await new Promise(r=>c.toBlob(r,'image/jpeg',0.8));
      return {data:new Uint8Array(await blob.arrayBuffer()),w:c.width,h:c.height};
    }
    async function loadPics(list){const r=await Promise.allSettled(list.map(async p=>({...p,...(await imgData(p.id))})));return {ok:r.filter(x=>x.status==='fulfilled').map(x=>x.value),failed:r.filter(x=>x.status!=='fulfilled').length}}
    let figNo=0;
    function photoGrid(list){
      const half=Math.floor(W/2);
      const cell=p=>new TableCell({width:{size:half,type:WidthType.DXA},borders:nb,margins:{top:80,bottom:120,left:80,right:80},children:p?[
        new Paragraph({alignment:AlignmentType.CENTER,children:[new ImageRun({data:p.data,transformation:(()=>{let w=280,h=Math.round(280*p.h/p.w);if(h>210){h=210;w=Math.round(210*p.w/p.h)}return {width:w,height:h}})()})]}),
        new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:60},children:[new TextRun({text:`Figure ${p.no}: `,bold:true,size:17,color:GREY}),new TextRun({text:p.caption,italics:true,size:17,color:GREY})]})]:[new Paragraph({children:[]})]});
      list.forEach(p=>p.no=++figNo);const rows=[];
      for(let i=0;i<list.length;i+=2)rows.push(new TableRow({cantSplit:true,children:[cell(list[i]),cell(list[i+1])]}));
      return new Table({width:{size:half*2,type:WidthType.DXA},columnWidths:[half,half],borders:{top:none,bottom:none,left:none,right:none,insideHorizontal:none,insideVertical:none},rows});
    }
    return {BR,NAVY,GREY,W,company,unit,P,H,H2,bul,B_,table,gap,cover,document,loadPics,photoGrid};
  }

  /* ---------- turning what was entered into tables ---------- */
  function fmtAns(q,v){
    if(v===undefined||v===null||v===''||(Array.isArray(v)&&!v.length))return null;
    if(Array.isArray(v))return v.join(', ');
    if(q.type==='number'){const n=num(v);return n===null?String(v):qf(n)+(q.unit?' '+q.unit:'')}
    if(q.type==='rating')return `${v} / 5 (${['','Very poor','Poor','Fair','Good','Very good'][+v]||''})`;
    if(q.type==='date')return dateTxt(v);
    return String(v);
  }
  function sectionRows(sid,pick){
    const s=SECTIONS.find(x=>x.id===sid);if(!s)return [];
    return s.qs.filter(q=>!pick||pick.includes(q.id)).map(q=>[q.label,fmtAns(q,ans(sid,q.id))]).filter(r=>r[1]!==null);
  }
  function sectionTable(K,sid,pick,head='Recorded'){const rows=sectionRows(sid,pick);return rows.length?K.table([3700,K.W-3700],[['Item',head],...rows],{boldFirst:true}):null}
  const pushT=(kids,t)=>{if(!t)return false;kids.push(t);kids.push(new (D().Paragraph)({spacing:{after:80},children:[]}));return true};
  const age=y=>{const n=num(y);return n?`${n} (${YEAR-n} yr)`:''};
  const withU=(v,u)=>num(v)!==null?qf(num(v))+(u?' '+u:''):'';

  function equipmentTable(K,list){
    if(!list.length)return null;
    const rows=list.slice().sort((a,b)=>(AREAS.indexOf(a.area)-AREAS.indexOf(b.area))||String(a.category).localeCompare(String(b.category))).map(e=>[
      [e.name||e.category||'Item',e.tag].filter(Boolean).join('\n'),e.area||'',[e.brand,e.model].filter(Boolean).join(' '),age(e.year),
      withU(e.cap,e.cap_unit)+(num(e.power)!==null?`\n${qf(num(e.power))} kW`:''),[e.status,e.duty&&e.duty!=='N/A'?e.duty:''].filter(Boolean).join('\n'),e.condition||'']);
    return K.table([1900,1400,1400,1000,1300,1100,926],[['Item and tag','Area','Make and model','Installed','Capacity','Status','Condition'],...rows],{size:17});
  }
  function chemicalsTable(K,list){
    if(!list.length)return null;
    return K.table([1900,1300,1400,1200,1300,1200,726],[['Chemical','Purpose','Used in / dosing point','Dose','Monthly use','Supplier','SDS'],
      ...list.map(c=>[[c.name,c.trade&&c.trade!==c.name?`(${c.trade})`:''].filter(Boolean).join('\n'),c.purpose||'',[c.used_in,c.point].filter(Boolean).join('\n'),withU(c.dose,c.dose_unit),withU(c.monthly,c.monthly_unit),c.supplier||'',c.sds||''])],{size:17});
  }
  function productsTable(K,list){
    if(!list.length)return null;
    return K.table([1900,1200,1300,1200,900,900,1626],[['Line / product','Type','Capacity','Output','Days/week','CIP/day','Wastes to drain'],
      ...list.map(p=>[p.name||'',p.type||'',withU(p.cap,p.cap_unit),withU(p.output,'L/day'),p.days||'',p.cip||'',p.wastes||''])],{size:17});
  }
  function readingsTable(K,list){
    if(!list.length)return null;
    return K.table([1800,700,700,900,900,1100,1000,1926],[['Point','Time','pH','Temp (°C)','DO (mg/L)','Cond. (µS/cm)','Turb. (NTU)','Notes'],
      ...list.slice().sort((a,b)=>POINT_NAMES.indexOf(a.location)-POINT_NAMES.indexOf(b.location)).map(m=>[[m.location,m.location_detail].filter(Boolean).join('\n'),m.time||'',m.ph||'',m.temp||'',m.do||'',m.ec||'',m.turb||'',[m.sample==='Yes'&&m.sample_id?'Sample '+m.sample_id:'',m.notes].filter(Boolean).join('\n')])],{size:17});
  }
  function observationsTable(K,list){
    if(!list.length)return null;
    const rank={Critical:0,Concern:1,Info:2};
    return K.table([2600,1600,1000,3826],[['Observation','Area','Severity','Detail'],...list.slice().sort((a,b)=>(rank[a.severity]??3)-(rank[b.severity]??3)).map(o=>[o.title||'',o.area||'',o.severity||'',o.detail||''])],{size:17});
  }
  function siteFacts(){
    const a=S.audit||{};const r=[];
    if(a.site)r.push(['Site',a.site]);if(a.location)r.push(['Address / location',a.location]);if(a.gps)r.push(['GPS coordinates',a.gps]);
    if(a.sector)r.push(['Industry / sector',a.sector]);if(a.visitDate)r.push(['Pre-audit visit',dateTxt(a.visitDate)]);if(a.team)r.push(['Cleanearth team',a.team]);
    return r;
  }
  function plantRows(){
    const r=[];const df=num(ans('wwtp','design_flow')),af=num(ans('wwtp','actual_flow')),dc=num(ans('wwtp','design_cod')),ac=num(ans('wwtp','actual_cod')),yr=num(ans('wwtp','commissioned'));
    if((ans('wwtp','technology')||[]).length)r.push(['Treatment stages',ans('wwtp','technology').join(', ')]);
    if(yr)r.push(['Commissioned',`${yr} (${YEAR-yr} years old)`]);
    if(ans('wwtp','designer'))r.push(['Designer / contractor',ans('wwtp','designer')]);
    if(df)r.push(['Design flow',`${qf(df)} m³/day`]);
    if(af)r.push(['Current flow',`${qf(af)} m³/day${df?` (${Math.round(af/df*100)}% of design)`:''}${ans('wwtp','flow_basis')?` – ${String(ans('wwtp','flow_basis')).toLowerCase()}`:''}`]);
    if(dc)r.push(['Design COD load',`${qf(dc)} kg/day`]);
    if(ac)r.push(['Current COD load',`${qf(ac)} kg/day${dc?` (${Math.round(ac/dc*100)}% of design)`:''}`]);
    const bv=num(ans('wwtp','bio_volume'));if(bv&&af)r.push(['Biological retention time',`${(bv/af*24).toFixed(1)} h at current flow`]);
    if(ans('wwtp','bypass'))r.push(['Units bypassed',ans('wwtp','bypass')]);
    return r;
  }
  function photoSets(){
    const sets=[];const seen=new Set();const add=(set,id,caption)=>{if(!id||seen.has(id))return;seen.add(id);set.items.push({id,caption})};
    const mk=t=>{const s={title:t,items:[]};sets.push(s);return s};
    const ob=mk('Observations'),eq=mk('Equipment'),ch=mk('Chemicals and storage'),pl=mk('Production lines'),fr=mk('Sampling points and readings'),wk=mk('Site walk-through and other');
    const rank={Critical:0,Concern:1,Info:2};
    S.regs.observations.slice().sort((a,b)=>(rank[a.severity]??3)-(rank[b.severity]??3)).forEach(o=>(o.photos||[]).forEach(id=>add(ob,id,`${o.title||'Observation'}${o.severity?' ('+o.severity.toLowerCase()+')':''}`)));
    S.regs.equipment.slice().sort((a,b)=>(['Not working','Poor','Fair','Good'].indexOf(a.condition)+1||9)-(['Not working','Poor','Fair','Good'].indexOf(b.condition)+1||9)).forEach(e=>(e.photos||[]).forEach(id=>add(eq,id,[e.name||e.category,e.tag,e.condition&&e.condition.toLowerCase()].filter(Boolean).join(', '))));
    S.regs.chemicals.forEach(c=>(c.photos||[]).forEach(id=>add(ch,id,c.name||c.trade||'Chemical')));
    S.regs.products.forEach(p=>(p.photos||[]).forEach(id=>add(pl,id,p.name||'Production line')));
    S.regs.measurements.forEach(m=>(m.photos||[]).forEach(id=>add(fr,id,`${m.location||'Sampling point'}${num(m.ph)!==null?', pH '+m.ph:''}`)));
    S.regs.tasks.filter(t=>t.phase==='pre').forEach(t=>(t.photos||[]).forEach(id=>add(wk,id,t.title)));
    return sets.filter(s=>s.items.length);
  }

  /* ================= PRE-AUDIT REPORT ================= */
  window.buildSiteReport=async function(){
    const K=await makeKit();const {P,H,H2,bul,B_,table,W}=K;const a=S.audit||{};
    const today=new Date().toISOString().slice(0,10);
    const ref=`${clientCode()}-PAR-${today.slice(2).replace(/-/g,'')}`;
    const F=flags();const crit=F.filter(f=>f[0]==='bad'),conc=F.filter(f=>f[0]!=='bad');
    const coverKids=K.cover('Pre-audit report','Site Assessment Report',[['Prepared for',a.client||TBC],['Prepared by',[a.team||(S.auth&&S.auth.name)||'',K.company+(K.unit?' '+K.unit:'')].filter(Boolean).join(', ')],['Visit date',a.visitDate?dateTxt(a.visitDate):TBC],['Issued',dateTxt(today)],['Reference',ref]]);
    const kids=[];

    let n=0;const sec=t=>H(`${++n}. ${t}`);
    kids.push(sec('Introduction'));
    kids.push(P(`${K.company} carried out a pre-audit visit to ${a.client||'the client'}’s ${a.site||'site'}${a.visitDate?` on ${dateTxt(a.visitDate)}`:''}. ${a.purpose||'The purpose was to understand the existing water and wastewater treatment, identify likely causes of non-compliant effluent, and gather the information needed to scope and cost a full audit.'}`));
    kids.push(P('This report records what was observed and the information collected during the visit. The findings are preliminary and will be confirmed or ruled out by a full audit.'));
    const visit=ta('pre','visit');if(visit.contacts||ans('site','contacts')){kids.push(H2('People met'));kids.push(P(visit.contacts||ans('site','contacts')))}

    kids.push(sec('Summary of findings'));
    const df=num(ans('wwtp','design_flow')),af=num(ans('wwtp','actual_flow'));
    kids.push(P(`The visit recorded ${S.regs.equipment.length} equipment items, ${S.regs.chemicals.length} chemicals, ${S.regs.products.length} production lines, ${S.regs.measurements.length} field readings and ${S.regs.observations.length} observations. ${crit.length} critical issues and ${conc.length} concerns were identified${df&&af?`, and the treatment plant is running at about ${Math.round(af/df*100)}% of its design flow`:''}.`));
    crit.forEach(f=>kids.push(bul([B_('Critical: '),f[1]])));conc.forEach(f=>kids.push(bul([B_('Concern: '),f[1]])));
    if(!F.length)kids.push(P('No significant issues were flagged automatically. See the assessor’s evaluation in section 13.'));

    kids.push(sec('Site and facility'));
    pushT(kids,table([3700,W-3700],[['Item','Recorded'],...siteFacts(),...sectionRows('site').filter(r=>r[0]!=='Client contacts met (name, role, phone)')],{boldFirst:true}));

    kids.push(sec('Water supply and treatment'));
    pushT(kids,sectionTable(K,'water'))||kids.push(P('No information recorded.'));
    const wtpEq=S.regs.equipment.filter(e=>['Water intake & storage','Water treatment'].includes(e.area));if(wtpEq.length){kids.push(H2('Water treatment equipment'));pushT(kids,equipmentTable(K,wtpEq))}

    kids.push(sec('Production and wastewater sources'));
    pushT(kids,sectionTable(K,'production'));
    if(S.regs.products.length){kids.push(H2('Production lines'));pushT(kids,productsTable(K,S.regs.products))}

    kids.push(sec('Existing wastewater treatment plant'));
    const pr=plantRows();if(pr.length)pushT(kids,table([3700,W-3700],[['Item','Recorded'],...pr],{boldFirst:true}));
    pushT(kids,sectionTable(K,'wwtp',['eq_volume','bio_volume','peak_flow','op_hours','docs','upgrades','notes']));
    if(sectionRows('daf').length){kids.push(H2('Pre-treatment and DAF'));pushT(kids,sectionTable(K,'daf'))}
    if(sectionRows('bio').length){kids.push(H2('Biological treatment and clarification'));pushT(kids,sectionTable(K,'bio'))}

    kids.push(sec('Equipment register'));
    if(S.regs.equipment.length){
      const c={};S.regs.equipment.forEach(e=>{const k=e.condition||'Not rated';c[k]=(c[k]||0)+1});
      kids.push(P(`${S.regs.equipment.length} items were recorded. Condition: ${Object.entries(c).map(([k,v])=>`${v} ${k.toLowerCase()}`).join(', ')}.`));
      pushT(kids,equipmentTable(K,S.regs.equipment));
      const iss=S.regs.equipment.filter(e=>e.issues);if(iss.length){kids.push(H2('Issues observed on equipment'));iss.forEach(e=>kids.push(bul([B_(`${e.name||e.category}${e.tag?' ('+e.tag+')':''}: `),e.issues])))}
    }else kids.push(P('No equipment was recorded.'));

    kids.push(sec('Chemicals and dosing'));
    if(S.regs.chemicals.length){
      const cost=S.regs.chemicals.reduce((x,c)=>x+(num(c.monthly)!==null&&num(c.cost)!==null?num(c.monthly)*num(c.cost):0),0);
      kids.push(P(`${S.regs.chemicals.length} chemicals are used on site${cost?`, at an estimated ${gh(cost)} per month based on recorded consumption and unit costs`:''}.`));
      pushT(kids,chemicalsTable(K,S.regs.chemicals));
    }else kids.push(P('No chemicals were recorded.'));
    if(sectionRows('dosing').length){kids.push(H2('Dosing practice'));pushT(kids,sectionTable(K,'dosing'))}

    if(sectionRows('cip').length){kids.push(sec('Cleaning in place (CIP)'));pushT(kids,sectionTable(K,'cip'))}
    if(sectionRows('sludge').length){kids.push(sec('Sludge handling'));pushT(kids,sectionTable(K,'sludge'))}

    kids.push(sec('Compliance and monitoring'));
    pushT(kids,sectionTable(K,'compliance'));
    const L=limitsOf();const lr=PARAMS.filter(p=>has((L[p.id]||{}).max)||has((L[p.id]||{}).min)).map(p=>[p.label,p.unit,has(L[p.id].min)?L[p.id].min:'',has(L[p.id].max)?L[p.id].max:'']);
    if(lr.length){kids.push(H2('Discharge limits'));pushT(kids,table([3200,2000,1900,1926],[['Parameter','Unit','Minimum','Maximum'],...lr]))}

    if(sectionRows('om').length){kids.push(sec('Operation, maintenance and management'));pushT(kids,sectionTable(K,'om'))}

    kids.push(sec('Field readings'));
    pushT(kids,readingsTable(K,S.regs.measurements))||kids.push(P('No field readings were taken.'));

    kids.push(sec('Observations'));
    pushT(kids,observationsTable(K,S.regs.observations))||kids.push(P('No observations were recorded.'));

    const ev=sectionRows('evaluation');
    kids.push(sec('Assessor’s evaluation'));
    if(ev.length){
      const rat=SECTIONS.find(s=>s.id==='evaluation').qs.filter(q=>q.type==='rating').map(q=>[q.label,fmtAns(q,ans('evaluation',q.id))]).filter(r=>r[1]);
      if(rat.length)pushT(kids,table([5500,W-5500],[['Area','Rating'],...rat],{boldFirst:true}));
      [['root_causes','Likely root causes of non-compliance'],['quick_wins','Quick wins identified'],['audit_scope','Items to include in the full audit']].forEach(([id,t])=>{const v=ans('evaluation',id);if(v){kids.push(H2(t));String(v).split(/\n+/).filter(Boolean).forEach(x=>kids.push(bul(x.replace(/^[-•]\s*/,''))))}});
    }else kids.push(P('The evaluation has not been completed.'));

    const co=ta('pre','closeout');
    kids.push(sec('Next steps'));
    if(co.next){String(co.next).split(/\n+/).filter(Boolean).forEach(x=>kids.push(bul(x.replace(/^[-•]\s*/,''))))}
    else kids.push(P(`${K.company} will prepare a proposal for a full audit based on these findings.`));
    if(co.outstanding){kids.push(H2('Information still to be provided by the client'));String(co.outstanding).split(/\n+/).filter(Boolean).forEach(x=>kids.push(bul(x.replace(/^[-•]\s*/,''))))}
    const ir=ta('pre','inforeq');if((ir.requested||[]).length){kids.push(H2('Information requested'));pushT(kids,table([5500,W-5500],[['Item','Status'],...ir.requested.map(x=>[x,(ir.received||[]).includes(x)?'Received':'Outstanding'])]))}

    const sets=photoSets();let failed=0;
    if(sets.length){
      kids.push(H('Appendix: Site photographs'));
      for(const s of sets){const r=await K.loadPics(s.items.slice(0,12));failed+=r.failed;if(r.ok.length){kids.push(H2(s.title));kids.push(K.photoGrid(r.ok))}}
    }
    const doc=K.document('Pre-audit report – '+(a.client||''),ref,coverKids,kids);
    const out=await D().Packer.toBlob(doc);out.photoFailed=failed;return out;
  };

  /* ================= AUDIT PROPOSAL ================= */
  window.buildProposal=async function(){
    const K=await makeKit();const {P,H,H2,bul,B_,table,gap,W}=K;
    const a=S.audit||{};const B=ta('proposal','basis'),SC=ta('proposal','scope'),R=ta('proposal','rates'),LT=ta('proposal','later'),TL=ta('proposal','timeline'),TM=ta('proposal','terms'),DOC=ta('proposal','doc');
    const today=new Date();const todayS=today.toISOString().slice(0,10);
    const wkS=v=>{const n=num(v);return n?`${n} week${n===1?'':'s'}`:TBC};
    const list=arr=>arr&&arr.length?arr.join(', '):TBC;

    /* photos that support the findings */
    const pics=[];const addPic=(id,caption,rank)=>{if(id&&!pics.some(p=>p.id===id))pics.push({id,caption,rank})};
    const sevRank={Critical:0,Concern:1,Info:3};
    S.regs.observations.forEach(o=>(o.photos||[]).slice(0,2).forEach(id=>addPic(id,`${o.title||'Observation'}${o.area?' ('+o.area+')':''}`,sevRank[o.severity]??2)));
    S.regs.equipment.filter(e=>['Poor','Not working'].includes(e.condition)||e.status==='Out of service').forEach(e=>(e.photos||[]).slice(0,1).forEach(id=>addPic(id,`${e.name||e.category||'Equipment'}${e.tag?' '+e.tag:''}: ${String(e.condition||e.status).toLowerCase()}`,e.condition==='Not working'||e.status==='Out of service'?0:1)));
    S.regs.tasks.filter(t=>t.phase==='pre').forEach(t=>(t.photos||[]).slice(0,2).forEach(id=>addPic(id,t.title,2)));
    S.regs.measurements.forEach(m=>(m.photos||[]).slice(0,1).forEach(id=>addPic(id,`${m.location||'Sampling point'}${num(m.ph)!==null?', pH '+m.ph:''}`,2)));
    pics.sort((x,y)=>x.rank-y.rank);
    const PL=await K.loadPics(pics.slice(0,8));

    const F=flags();const crit=F.filter(f=>f[0]==='bad'),conc=F.filter(f=>f[0]!=='bad');
    const failing=ans('compliance','failing')||[];const sc=scopeNumbers();const wk=weeksPlan();const pr=programmeRange();
    const ref=DOC.ref||TBC;const validity=num(TM.validity);
    const validUntil=validity?(()=>{const d=new Date(today);d.setDate(d.getDate()+validity);return d.toISOString().slice(0,10)})():null;
    const audLines=S.regs.costing.filter(c=>(c.phase||'Full audit')==='Full audit'&&costAmount(c)!==null).sort((x,y)=>(x.order??99)-(y.order??99));
    const ohF=1+n0(R.overhead)/100;const sub=audLines.reduce((x,c)=>x+costAmount(c),0);const cont=sub*n0(R.contingency)/100;
    const feeRows=audLines.map(c=>[c.item||c.category||'Item',qf(num(c.qty))+' '+(c.unit||''),gh(num(c.rate)*ohF),gh(costAmount(c)*ohF)]);
    const net=(sub+cont)*ohF;const tax=net*n0(R.tax)/100;
    const METHOD={
      'WWTP performance':'Flow-proportional 24-hour composite samples and grab samples at each treatment stage, so that the removal achieved by every unit can be measured against its design. Field readings of pH, temperature, dissolved oxygen and conductivity will be taken daily.',
      'Water treatment plant':'Raw and treated water will be sampled, and the performance of each treatment unit checked, including where backwash, brine and reject streams are discharged.',
      'Production and CIP source survey':'Individual streams such as CIP rinses, product line washouts and floor drains will be sampled to show where the organic, fat and pH load originates, and where it can be reduced or recovered at source.',
      'Water balance':'Water will be tracked from intake through production, cleaning and utilities to discharge, and benchmarked against water use per litre of product.',
      'Chemicals and dosing':'Chemical use will be compared with actual treatment needs. Jar tests will establish optimum coagulant and flocculant doses, and alternative products will be considered where they offer savings or better performance.',
      'Sludge handling':'Sludge production, dewatering performance and disposal routes will be quantified and checked for compliance.',
      'Asset condition survey':'Every treatment unit and its mechanical, electrical and instrumentation equipment will be assessed for condition, age, remaining life and critical spares.',
      'Operations and management review':'Operator interviews and a review of logs, procedures and maintenance records will identify gaps in operating practice, monitoring and training.'};

    const coverKids=K.cover('Technical and financial proposal','Water and Wastewater Treatment Audit',[['Prepared for',a.client||TBC],['Prepared by',DOC.author?`${DOC.author}, ${K.company}${K.unit?' '+K.unit:''}`:K.company+(K.unit?' '+K.unit:'')],['Issued',dateTxt(todayS)],['Valid until',validUntil?dateTxt(validUntil):TBC],['Reference',ref]]);
    const kids=[];

    kids.push(H('1. Introduction'));
    kids.push(P(`${K.company}’s ${K.unit||'team'} carried out a pre-audit visit to ${a.client||'the client'}’s ${a.site||'site'}${a.location?` (${a.location})`:''}${a.visitDate?` on ${dateTxt(a.visitDate)}`:''} to understand the existing water and wastewater treatment and the reasons for non-compliant effluent.`));
    kids.push(P('This proposal sets out a full audit of the treatment systems. The audit will establish the root causes of non-compliance, confirm the actual flows and loads, and provide the basis for a targeted remediation programme.'));

    /* 2 background – everything relevant from the pre-audit */
    kids.push(H('2. Background and site understanding'));
    kids.push(P(B.problem||'The treated effluent does not consistently meet the applicable discharge limits.'));
    if((B.objectives||[]).length){kids.push(P([B_('The client’s objectives are:')]));B.objectives.forEach(o=>kids.push(bul(o)))}
    if(B.deadline)kids.push(P(`The client is working to a compliance deadline of ${dateTxt(B.deadline)}.`));
    const site=[...siteFacts().filter(r=>r[0]!=='Cleanearth team'),...sectionRows('site',['site_area','wwtp_footprint','free_space','free_space_loc','access','prod_days','shifts','staff','expansion'])];
    if(site.length){kids.push(H2('2.1 Site and operations'));pushT(kids,table([3700,W-3700],[['Item','Recorded'],...site],{boldFirst:true}))}
    const water=sectionRows('water',['sources','consumption','metered','wtp_units','wtp_capacity','wtp_reject','raw_quality']);
    if(water.length){kids.push(H2('2.2 Water supply and treatment'));pushT(kids,table([3700,W-3700],[['Item','Recorded'],...water],{boldFirst:true}))}
    const prod=sectionRows('production',['main_sources','losses','segregation','domestic','storm','grease_traps','hot_streams']);
    if(prod.length||S.regs.products.length){kids.push(H2('2.3 Production and wastewater sources'));if(prod.length)pushT(kids,table([3700,W-3700],[['Item','Recorded'],...prod],{boldFirst:true}));if(S.regs.products.length)pushT(kids,productsTable(K,S.regs.products))}
    const plant=[...plantRows(),...sectionRows('daf',['daf_running','ph_control','daf_float']),...sectionRows('bio',['aeration','do_reading','bulking','ras']),...sectionRows('sludge',['dewatering','disposal']),...sectionRows('cip',['cip_type','cip_cycles','cip_dump','cip_recovery'])];
    if(plant.length){kids.push(H2('2.4 Existing treatment plant'));pushT(kids,table([3700,W-3700],[['Item','Recorded'],...plant],{boldFirst:true}))}
    if(S.regs.equipment.length){
      const bad=S.regs.equipment.filter(e=>['Poor','Not working'].includes(e.condition)||e.status==='Out of service');
      kids.push(H2('2.5 Equipment'));
      kids.push(P(`${S.regs.equipment.length} items of equipment were recorded during the pre-audit${bad.length?`, of which ${bad.length} ${bad.length>1?'are':'is'} in poor condition, not working or out of service`:''}. The full register is included in the pre-audit report.`));
      if(bad.length)pushT(kids,equipmentTable(K,bad));
    }
    if(S.regs.chemicals.length){kids.push(H2('2.6 Chemicals in use'));pushT(kids,chemicalsTable(K,S.regs.chemicals))}
    const comp=sectionRows('compliance',['permit','discharge_to','failing','fail_freq','sampling_freq','lab','epa_actions']);
    if(comp.length){kids.push(H2('2.7 Compliance status'));pushT(kids,table([3700,W-3700],[['Item','Recorded'],...comp],{boldFirst:true}))}
    if(S.regs.measurements.length){kids.push(H2('2.8 Field readings during the visit'));pushT(kids,readingsTable(K,S.regs.measurements))}

    kids.push(H('3. Preliminary findings from the pre-audit'));
    if(F.length){
      kids.push(P('The pre-audit identified the following issues. They are indicative and will be confirmed or ruled out by the audit.'));
      crit.slice(0,12).forEach(f=>kids.push(bul([B_('Critical: '),f[1]])));conc.slice(0,12).forEach(f=>kids.push(bul([B_('Concern: '),f[1]])));
    }else kids.push(P('No significant issues were recorded during the pre-audit.'));
    const rc=ans('evaluation','root_causes');if(rc){kids.push(H2('Likely causes'));String(rc).split(/\n+/).filter(Boolean).forEach(x=>kids.push(bul(x.replace(/^[-•]\s*/,''))))}
    if(PL.ok.length){kids.push(H2('Site photographs'));kids.push(P('Selected photographs from the pre-audit visit that illustrate these findings.'));kids.push(K.photoGrid(PL.ok))}

    kids.push(H('4. Objectives of the audit'));
    ['Establish the root causes of each non-compliant parameter.','Measure actual flows and pollutant loads, and compare them with the plant’s design capacity.','Assess the performance of each treatment unit, the chemicals used and the way the plant is operated.','Identify practical options, from quick operational changes to capital upgrades, with indicative costs.','Provide a confirmed design basis for remediation.'].forEach(t=>kids.push(bul(t)));

    kids.push(H('5. Scope of work'));
    if((SC.areas||[]).length){kids.push(P([B_('The audit will cover:')]));SC.areas.forEach(x=>kids.push(bul(x)))}
    kids.push(H2('Sampling and monitoring programme'));
    kids.push(table([3600,W-3600],[['Item','Planned'],
      ['Sampling points',list(SC.points)],['Monitoring duration',SC.days?`${SC.days} consecutive days`:TBC],
      ['Composite samples',SC.comp?`${SC.comp} per point per day`:TBC],['Grab samples',SC.grab?`${SC.grab} per point per day`:'None'],
      ['Source samples',SC.source?`${SC.source} in total`:'None'],['Total samples',sc.samples||TBC],
      ['Parameters on each sample',list(SC.params)],['Total laboratory analyses',sc.analyses||TBC],
      ['Process tests',(SC.tests||[]).length?SC.tests.join(', '):'None'],['Flow measurement',SC.metering?(SC.metering+(SC.metering==='Hire portable meters'&&SC.meters?` (${SC.meters})`:'')):TBC]],{boldFirst:true}));

    kids.push(H('6. Methodology'));
    kids.push(P('The audit follows the water from source to discharge, so that problems at the treatment plant can be traced back to where they start.'));
    (SC.areas||Object.keys(METHOD)).forEach(x=>{if(METHOD[x]){kids.push(H2(x));kids.push(P(METHOD[x]))}});
    kids.push(H2('Laboratory analysis and data quality'));
    kids.push(P('Samples will be preserved, kept cool and delivered under chain of custody to an accredited laboratory within the holding times for each parameter. Every sample carries a unique code. Results will be checked for consistency before analysis.'));

    kids.push(H('7. Deliverables'));
    ['An audit report with root causes for each failing parameter and a compliance gap analysis.','Measured flows, loads and removal efficiencies for each treatment stage.','A water balance and benchmark of water use.','An asset register with condition ratings.','Prioritised remediation options with indicative capital and running costs.','A presentation of the findings to the client’s team.'].forEach(t=>kids.push(bul(t)));

    kids.push(H('8. Programme'));
    kids.push(table([5000,W-5000],[['Stage','Duration'],['Mobilisation',wkS(TL.mob_weeks)],['Fieldwork and sampling',wkS(TL.field_weeks)],['Laboratory analysis',wkS(TL.lab_weeks)],['Data analysis and reporting',wkS(TL.report_weeks)],['Total',wkS(wk.w)]],{boldRows:[5]}));
    kids.push(P(TL.start?`Subject to approval, work can start on ${dateTxt(TL.start)}${wk.end?`, with the final report expected by ${dateTxt(wk.end)}`:''}.`:'The start date will be agreed on acceptance of this proposal.',{spacing:{before:120,after:120}}));

    kids.push(H('9. Fees for the audit'));
    if(feeRows.length){
      const rows=[['Item','Quantity','Rate','Amount'],...feeRows];
      if(cont)rows.push(['Contingency','','',gh(cont*ohF)]);
      rows.push(['Total excluding taxes','','',gh(net)]);const bold=[rows.length-1];
      if(n0(R.tax)){rows.push([`VAT and levies (${n0(R.tax)}%)`,'','',gh(tax)]);rows.push(['Total','','',gh(net+tax)]);bold.push(rows.length-1)}
      kids.push(table([4026,1700,1500,1800],rows,{right:[1,2,3],boldRows:bold}));
    }else kids.push(P('The fee will be confirmed once the scope and rates are finalised.'));
    if(TM.payment)kids.push(P([B_('Payment terms: '),TM.payment+'.'],{spacing:{before:120,after:80}}));
    if(validity)kids.push(P(`This proposal is valid for ${validity} days from the date above.`));

    kids.push(H('10. Indicative cost of the full programme'));
    kids.push(P('To help with budgeting, the table below shows the likely cost of the whole programme. Only the audit fee is fixed. Remediation and monitoring costs are indicative and will be confirmed by the audit findings.'));
    const remTxt=pr.remHi?(pr.remLo===pr.remHi?gh(pr.remLo):`${gh(pr.remLo)} – ${gh(pr.remHi)}`):TBC;
    kids.push(table([2900,3300,W-6200],[['Phase','Estimate','Basis'],
      ['Full audit',gh(net+tax),'Fixed fee (section 9)'],
      ['Remediation and retrofit',remTxt,(LT.rem_scope||'Scope to be confirmed')+(n0(LT.design_pct)?`, including ${n0(LT.design_pct)}% design and supervision`:'')],
      ['Monitoring and aftercare',pr.mon?gh(pr.mon):TBC,num(LT.mon_months)?`${LT.mon_months} months at ${gh(num(LT.mon_cost))} per month`:''],
      ['Whole programme',pr.hi?(Math.round(pr.lo)===Math.round(pr.hi)?gh(net+tax+pr.remLo+pr.mon):`${gh(net+tax+pr.remLo+pr.mon)} – ${gh(net+tax+pr.remHi+pr.mon)}`):TBC,'Indicative']],{boldRows:[4],right:[1]}));
    if(LT.basis)kids.push(P(LT.basis,{spacing:{before:120,after:120}}));

    kids.push(H('11. Client responsibilities, assumptions and exclusions'));
    if((TM.client_resp||[]).length){kids.push(H2('The client will provide'));TM.client_resp.forEach(x=>kids.push(bul(x)))}
    kids.push(H2('Assumptions'));
    String(TM.assumptions||'Production runs normally during the monitoring period.\nSamples are analysed by an accredited laboratory.\nSafe access to all sampling points is available.').split(/\n+/).filter(Boolean).forEach(x=>kids.push(bul(x.replace(/^[-•]\s*/,''))));
    kids.push(H2('Exclusions'));
    String(TM.exclusions||'Detailed design and construction of remediation works.\nPermit applications and regulatory submissions.').split(/\n+/).filter(Boolean).forEach(x=>kids.push(bul(x.replace(/^[-•]\s*/,''))));

    kids.push(H('12. Acceptance'));
    kids.push(P('To accept this proposal, please sign below and return a copy, or issue a purchase order quoting the proposal reference.'));
    const BR=K.BR;
    kids.push(table([2400,3313,3313],[['','For '+K.company,'For '+(a.client||'the client')],['Name',BR.signatory_name||'',''],['Position',BR.signatory_title||'',''],['Signature','',''],['Date','','']],{boldFirst:true,shadeFirst:true}));

    const doc=K.document('Audit proposal – '+(a.client||''),ref,coverKids,kids);
    const out=await D().Packer.toBlob(doc);out.photoCount=PL.ok.length;out.photoFailed=PL.failed;return out;
  };
})();
