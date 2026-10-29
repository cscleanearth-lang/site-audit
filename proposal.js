/* Cleanearth Site Audit – builds the audit proposal as a Word document from the project data.
   Loaded on demand after vendor/docx.umd.js. Uses the app's own helpers (S, ans, flags, ta, costSummary…). */
(function(){
  'use strict';
  window.buildProposal=async function(){
    const D=window.docx;
    const {Document,Packer,Paragraph,TextRun,Table,TableRow,TableCell,WidthType,ShadingType,AlignmentType,BorderStyle,Footer,PageNumber,LevelFormat,HeadingLevel,ImageRun}=D;
    const NAVY='1F3B5C',GREY='6B7F96',LIGHT='E8EEF5',FONT='Candara',W=9026;
    const a=S.audit||{};const B=ta('proposal','basis'),SC=ta('proposal','scope'),R=ta('proposal','rates'),LT=ta('proposal','later'),TL=ta('proposal','timeline'),TM=ta('proposal','terms'),DOC=ta('proposal','doc');
    const TBC='[to be confirmed]';const today=new Date();const todayS=today.toISOString().slice(0,10);
    const dateTxt=d=>d?fmtDate(d).replace(/^\w+, /,''):TBC;
    const gh=x=>x===null||x===undefined||!isFinite(x)?'–':'GHS '+Math.round(x).toLocaleString('en-GB');
    const qf=x=>x===null||x===undefined?'':Number.isInteger(x)?x.toLocaleString('en-GB'):x.toLocaleString('en-GB',{maximumFractionDigits:2});
    const wkS=v=>{const n=num(v);return n?`${n} week${n===1?'':'s'}`:TBC};
    const list=(arr)=>arr&&arr.length?arr.join(', '):TBC;

    /* ---- building blocks ---- */
    const P=(t,o={})=>new Paragraph({spacing:{after:120,line:276},...o,children:[].concat(t).map(x=>typeof x==='string'?new TextRun(x):x)});
    const H=t=>new Paragraph({heading:HeadingLevel.HEADING_1,keepNext:true,spacing:{before:280,after:120},children:[new TextRun(t)]});
    const H2=t=>new Paragraph({heading:HeadingLevel.HEADING_2,keepNext:true,spacing:{before:180,after:80},children:[new TextRun(t)]});
    const bul=t=>new Paragraph({numbering:{reference:'b',level:0},spacing:{after:60},children:[].concat(t).map(x=>typeof x==='string'?new TextRun(x):x)});
    const B_=t=>new TextRun({text:t,bold:true});
    const line={style:BorderStyle.SINGLE,size:4,color:'A9B8C9'};const borders={top:line,bottom:line,left:line,right:line};
    function table(cols,rows,opt={}){
      const cell=(txt,i,ri)=>{const hdr=ri===0&&!opt.noHeader;const right=opt.right&&opt.right.includes(i)&&!hdr;const bold=hdr||(opt.boldFirst&&i===0)||(opt.boldRows&&opt.boldRows.includes(ri));
        return new TableCell({width:{size:cols[i],type:WidthType.DXA},borders,margins:{top:60,bottom:60,left:100,right:100},
          shading:hdr?{type:ShadingType.CLEAR,fill:NAVY,color:'auto'}:(opt.shadeFirst&&i===0)||(opt.boldRows&&opt.boldRows.includes(ri))?{type:ShadingType.CLEAR,fill:LIGHT,color:'auto'}:undefined,
          children:[new Paragraph({alignment:right?AlignmentType.RIGHT:AlignmentType.LEFT,children:[new TextRun({text:String(txt??''),bold,color:hdr?'FFFFFF':undefined,size:19})]})]})};
      return new Table({width:{size:W,type:WidthType.DXA},columnWidths:cols,rows:rows.map((r,ri)=>new TableRow({tableHeader:ri===0&&!opt.noHeader,cantSplit:true,children:r.map((c,i)=>cell(c,i,ri))}))});
    }
    const gap=()=>new Paragraph({spacing:{after:80},children:[]});

    /* ---- site photographs that support the findings ---- */
    async function imgData(id){
      const url=await CE.photoURL(id);
      const img=await new Promise((res,rej)=>{const i=new Image();i.onload=()=>res(i);i.onerror=rej;i.src=url});
      const sc=Math.min(1,1200/Math.max(img.naturalWidth,img.naturalHeight));
      const c=document.createElement('canvas');c.width=Math.max(1,Math.round(img.naturalWidth*sc));c.height=Math.max(1,Math.round(img.naturalHeight*sc));
      c.getContext('2d').drawImage(img,0,0,c.width,c.height);
      const blob=await new Promise(r=>c.toBlob(r,'image/jpeg',0.82));
      return {data:new Uint8Array(await blob.arrayBuffer()),w:c.width,h:c.height};
    }
    const pics=[];const addPic=(id,caption,rank)=>{if(id&&!pics.some(p=>p.id===id))pics.push({id,caption,rank})};
    const sevRank={Critical:0,Concern:1,Info:3};
    S.regs.observations.forEach(o=>(o.photos||[]).slice(0,2).forEach(id=>addPic(id,`${o.title||'Observation'}${o.area?' ('+o.area+')':''}`,sevRank[o.severity]??2)));
    S.regs.equipment.filter(e=>['Poor','Not working'].includes(e.condition)||e.status==='Out of service').forEach(e=>(e.photos||[]).slice(0,1).forEach(id=>addPic(id,`${e.name||e.category||'Equipment'}${e.tag?' '+e.tag:''}: ${String(e.condition||e.status).toLowerCase()}`,e.condition==='Not working'||e.status==='Out of service'?0:1)));
    S.regs.tasks.filter(t=>t.phase==='pre').forEach(t=>(t.photos||[]).slice(0,2).forEach(id=>addPic(id,t.title,2)));
    S.regs.measurements.forEach(m=>(m.photos||[]).slice(0,1).forEach(id=>addPic(id,`${m.location||'Sampling point'}${num(m.ph)!==null?', pH '+m.ph:''}`,2)));
    pics.sort((x,y)=>x.rank-y.rank);
    const loaded=(await Promise.allSettled(pics.slice(0,8).map(async p=>({...p,...(await imgData(p.id))})))).filter(r=>r.status==='fulfilled').map(r=>r.value);
    const photoFailed=Math.min(pics.length,8)-loaded.length;
    function photoGrid(list,startNo){
      const none={style:BorderStyle.NONE,size:0,color:'FFFFFF'};const nb={top:none,bottom:none,left:none,right:none};const half=Math.floor(W/2);
      const cell=p=>new TableCell({width:{size:half,type:WidthType.DXA},borders:nb,margins:{top:80,bottom:120,left:80,right:80},children:p?[
        new Paragraph({alignment:AlignmentType.CENTER,children:[new ImageRun({data:p.data,transformation:(()=>{let w=280,h=Math.round(280*p.h/p.w);if(h>210){h=210;w=Math.round(210*p.w/p.h)}return {width:w,height:h}})()})]}),
        new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:60},children:[new TextRun({text:`Figure ${p.no}: `,bold:true,size:17,color:GREY}),new TextRun({text:p.caption,italics:true,size:17,color:GREY})]})]:[new Paragraph({children:[]})]});
      const rows=[];list.forEach((p,i)=>p.no=startNo+i);
      for(let i=0;i<list.length;i+=2)rows.push(new TableRow({cantSplit:true,children:[cell(list[i]),cell(list[i+1])]}));
      return new Table({width:{size:half*2,type:WidthType.DXA},columnWidths:[half,half],rows});
    }

    /* ---- facts from the project ---- */
    const F=flags();const crit=F.filter(f=>f[0]==='bad'),conc=F.filter(f=>f[0]!=='bad');
    const failing=ans('compliance','failing')||[];const sc=scopeNumbers();const wk=weeksPlan();const cs=costSummary();const pr=programmeRange();
    const df=num(ans('wwtp','design_flow')),af=num(ans('wwtp','actual_flow')),yr=num(ans('wwtp','commissioned'));
    const ref=DOC.ref||TBC;const validity=num(TM.validity);
    const validUntil=validity?(()=>{const d=new Date(today);d.setDate(d.getDate()+validity);return d.toISOString().slice(0,10)})():null;

    /* ---- fees: overhead spread across the lines, contingency and tax shown ---- */
    const audLines=S.regs.costing.filter(c=>(c.phase||'Full audit')==='Full audit'&&costAmount(c)!==null).sort((x,y)=>(x.order??99)-(y.order??99));
    const ohF=1+n0(R.overhead)/100;const sub=audLines.reduce((x,c)=>x+costAmount(c),0);const cont=sub*n0(R.contingency)/100;
    const feeRows=audLines.map(c=>[c.item||c.category||'Item',qf(num(c.qty))+' '+(c.unit||''),gh(num(c.rate)*ohF),gh(costAmount(c)*ohF)]);
    const net=(sub+cont)*ohF;const tax=net*n0(R.tax)/100;

    /* ---- methodology text per scope area ---- */
    const METHOD={
      'WWTP performance':'Flow-proportional 24-hour composite samples and grab samples at each treatment stage, so that the removal achieved by every unit can be measured against its design. Field readings of pH, temperature, dissolved oxygen and conductivity will be taken daily.',
      'Water treatment plant':'Raw and treated water will be sampled, and the performance of each treatment unit checked, including where backwash, brine and reject streams are discharged.',
      'Production and CIP source survey':'Individual streams such as CIP rinses, product line washouts and floor drains will be sampled to show where the organic, fat and pH load originates, and where it can be reduced or recovered at source.',
      'Water balance':'Water will be tracked from intake through production, cleaning and utilities to discharge, and benchmarked against water use per litre of product.',
      'Chemicals and dosing':'Chemical use will be compared with actual treatment needs. Jar tests will establish optimum coagulant and flocculant doses, and alternative products will be considered where they offer savings or better performance.',
      'Sludge handling':'Sludge production, dewatering performance and disposal routes will be quantified and checked for compliance.',
      'Asset condition survey':'Every treatment unit and its mechanical, electrical and instrumentation equipment will be assessed for condition, age, remaining life and critical spares.',
      'Operations and management review':'Operator interviews and a review of logs, procedures and maintenance records will identify gaps in operating practice, monitoring and training.'};

    const kids=[];
    /* title block */
    kids.push(new Paragraph({spacing:{after:40},children:[new TextRun({text:'CLEANEARTH  |  WATER SERVICES',bold:true,color:GREY,size:18,characterSpacing:40})]}));
    kids.push(new Paragraph({spacing:{after:60},children:[new TextRun({text:'Proposal for a Water and Wastewater Treatment Audit',bold:true,size:38,color:NAVY})]}));
    kids.push(new Paragraph({spacing:{after:200},border:{bottom:{style:BorderStyle.SINGLE,size:12,color:NAVY,space:6}},children:[new TextRun({text:[a.client,a.site].filter(Boolean).join(' – ')||TBC,size:26,color:NAVY})]}));
    kids.push(table([2600,W-2600],[['Client',a.client||TBC],['Site',a.site||TBC],['Proposal reference',ref],['Date',dateTxt(todayS)],['Prepared by',DOC.author?`${DOC.author}, Cleanearth Water Services`:'Cleanearth Water Services'],['Valid until',validUntil?dateTxt(validUntil):TBC]],{noHeader:true,boldFirst:true,shadeFirst:true}));

    /* 1 introduction */
    kids.push(H('1. Introduction'));
    kids.push(P(`Cleanearth’s Water Services team carried out a pre-audit visit to ${a.client||'the client'}’s ${a.site||'site'}${a.visitDate?` on ${dateTxt(a.visitDate)}`:''} to understand the existing water and wastewater treatment and the reasons for non-compliant effluent.`));
    kids.push(P('This proposal sets out a full audit of the treatment systems. The audit will establish the root causes of non-compliance, confirm the actual flows and loads, and provide the basis for a targeted remediation programme.'));

    /* 2 background */
    kids.push(H('2. Background and site understanding'));
    kids.push(P(B.problem||'The treated effluent does not consistently meet the applicable discharge limits.'));
    const facts=[];
    if(a.sector)facts.push(['Industry',a.sector]);
    if(ans('compliance','discharge_to'))facts.push(['Discharge point',ans('compliance','discharge_to')]);
    if(failing.length)facts.push(['Parameters failing',failing.join(', ')+(ans('compliance','fail_freq')?` (${String(ans('compliance','fail_freq')).toLowerCase()})`:'')]);
    if((ans('wwtp','technology')||[]).length)facts.push(['Treatment stages',ans('wwtp','technology').join(', ')]);
    if(df)facts.push(['Design flow',`${fmtN(df)} m³/day`]);
    if(af)facts.push(['Current flow (reported)',`${fmtN(af)} m³/day${df?` – ${Math.round(af/df*100)}% of design`:''}`]);
    if(yr)facts.push(['Plant commissioned',`${yr} (${YEAR-yr} years old)`]);
    if(S.regs.equipment.length)facts.push(['Equipment recorded',`${S.regs.equipment.length} items, ${S.regs.equipment.filter(e=>e.condition==='Poor'||e.condition==='Not working').length} in poor condition or not working`]);
    if(facts.length){kids.push(table([3000,W-3000],[['Item','Pre-audit observation'],...facts],{boldFirst:true}));kids.push(gap())}
    if((B.objectives||[]).length){kids.push(P([B_('The client’s objectives are:')]));B.objectives.forEach(o=>kids.push(bul(o)))}
    if(B.deadline)kids.push(P(`The client is working to a compliance deadline of ${dateTxt(B.deadline)}.`));

    /* 3 findings */
    kids.push(H('3. Preliminary findings from the pre-audit'));
    if(F.length){
      kids.push(P('The pre-audit identified the following issues. They are indicative and will be confirmed or ruled out by the audit.'));
      crit.slice(0,10).forEach(f=>kids.push(bul([B_('Critical: '),f[1]])));
      conc.slice(0,10).forEach(f=>kids.push(bul([B_('Concern: '),f[1]])));
    }else kids.push(P('No significant issues were recorded during the pre-audit.'));
    if(loaded.length){kids.push(H2('Site photographs'));kids.push(P('Selected photographs from the pre-audit visit that illustrate these findings.'));kids.push(photoGrid(loaded,1))}

    /* 4 objectives */
    kids.push(H('4. Objectives of the audit'));
    ['Establish the root causes of each non-compliant parameter.','Measure actual flows and pollutant loads, and compare them with the plant’s design capacity.','Assess the performance of each treatment unit, the chemicals used and the way the plant is operated.','Identify practical options, from quick operational changes to capital upgrades, with indicative costs.','Provide a confirmed design basis for remediation.'].forEach(t=>kids.push(bul(t)));

    /* 5 scope */
    kids.push(H('5. Scope of work'));
    if((SC.areas||[]).length){kids.push(P([B_('The audit will cover:')]));SC.areas.forEach(x=>kids.push(bul(x)))}
    kids.push(H2('Sampling and monitoring programme'));
    kids.push(table([3600,W-3600],[['Item','Planned'],
      ['Sampling points',list(SC.points)],['Monitoring duration',SC.days?`${SC.days} consecutive days`:TBC],
      ['Composite samples',SC.comp?`${SC.comp} per point per day`:TBC],['Grab samples',SC.grab?`${SC.grab} per point per day`:'None'],
      ['Source samples',SC.source?`${SC.source} in total`:'None'],['Total samples',sc.samples||TBC],
      ['Parameters on each sample',list(SC.params)],['Total laboratory analyses',sc.analyses||TBC],
      ['Process tests',(SC.tests||[]).length?SC.tests.join(', '):'None'],['Flow measurement',SC.metering?(SC.metering+(SC.metering==='Hire portable meters'&&SC.meters?` (${SC.meters})`:'')):TBC]],{boldFirst:true}));

    /* 6 methodology */
    kids.push(H('6. Methodology'));
    kids.push(P('The audit follows the water from source to discharge, so that problems at the treatment plant can be traced back to where they start.'));
    (SC.areas||Object.keys(METHOD)).forEach(x=>{if(METHOD[x]){kids.push(H2(x));kids.push(P(METHOD[x]))}});
    kids.push(H2('Laboratory analysis and data quality'));
    kids.push(P('Samples will be preserved, kept cool and delivered under chain of custody to an accredited laboratory within the holding times for each parameter. Every sample carries a unique code. Results will be checked for consistency before analysis.'));

    /* 7 deliverables */
    kids.push(H('7. Deliverables'));
    ['An audit report with root causes for each failing parameter and a compliance gap analysis.','Measured flows, loads and removal efficiencies for each treatment stage.','A water balance and benchmark of water use.','An asset register with condition ratings.','Prioritised remediation options with indicative capital and running costs.','A presentation of the findings to the client’s team.'].forEach(t=>kids.push(bul(t)));

    /* 8 programme */
    kids.push(H('8. Programme'));
    kids.push(table([5000,W-5000],[['Stage','Duration'],['Mobilisation',wkS(TL.mob_weeks)],['Fieldwork and sampling',wkS(TL.field_weeks)],['Laboratory analysis',wkS(TL.lab_weeks)],['Data analysis and reporting',wkS(TL.report_weeks)],['Total',wkS(wk.w)]],{boldRows:[5]}));
    kids.push(P(TL.start?`Subject to approval, work can start on ${dateTxt(TL.start)}${wk.end?`, with the final report expected by ${dateTxt(wk.end)}`:''}.`:'The start date will be agreed on acceptance of this proposal.',{spacing:{before:120,after:120}}));

    /* 9 fees */
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

    /* 10 programme estimate */
    kids.push(H('10. Indicative cost of the full programme'));
    kids.push(P('To help with budgeting, the table below shows the likely cost of the whole programme. Only the audit fee is fixed. Remediation and monitoring costs are indicative and will be confirmed by the audit findings.'));
    const remTxt=pr.remHi?(pr.remLo===pr.remHi?gh(pr.remLo):`${gh(pr.remLo)} – ${gh(pr.remHi)}`):TBC;
    kids.push(table([2900,3300,W-6200],[['Phase','Estimate','Basis'],
      ['Full audit',gh(net+tax),'Fixed fee (section 9)'],
      ['Remediation and retrofit',remTxt,(LT.rem_scope||'Scope to be confirmed')+(n0(LT.design_pct)?`, including ${n0(LT.design_pct)}% design and supervision`:'')],
      ['Monitoring and aftercare',pr.mon?gh(pr.mon):TBC,num(LT.mon_months)?`${LT.mon_months} months at ${gh(num(LT.mon_cost))} per month`:''],
      ['Whole programme',pr.hi?(Math.round(pr.lo)===Math.round(pr.hi)?gh(net+tax+pr.remLo+pr.mon):`${gh(net+tax+pr.remLo+pr.mon)} – ${gh(net+tax+pr.remHi+pr.mon)}`):TBC,'Indicative']],{boldRows:[4],right:[1]}));
    if(LT.basis)kids.push(P(LT.basis,{spacing:{before:120,after:120}}));

    /* 11 responsibilities, assumptions, exclusions */
    kids.push(H('11. Client responsibilities, assumptions and exclusions'));
    if((TM.client_resp||[]).length){kids.push(H2('The client will provide'));TM.client_resp.forEach(x=>kids.push(bul(x)))}
    kids.push(H2('Assumptions'));
    String(TM.assumptions||'Production runs normally during the monitoring period.\nSamples are analysed by an accredited laboratory.\nSafe access to all sampling points is available.').split(/\n+/).filter(Boolean).forEach(x=>kids.push(bul(x.replace(/^[-•]\s*/,''))));
    kids.push(H2('Exclusions'));
    String(TM.exclusions||'Detailed design and construction of remediation works.\nPermit applications and regulatory submissions.').split(/\n+/).filter(Boolean).forEach(x=>kids.push(bul(x.replace(/^[-•]\s*/,''))));

    /* 12 acceptance */
    kids.push(H('12. Acceptance'));
    kids.push(P('To accept this proposal, please sign below and return a copy, or issue a purchase order quoting the proposal reference.'));
    kids.push(table([2400,3313,3313],[['','For Cleanearth','For '+(a.client||'the client')],['Name','',''],['Position','',''],['Signature','',''],['Date','','']],{boldFirst:true,shadeFirst:true}));

    const doc=new Document({
      creator:'Cleanearth Water Services',title:'Audit proposal – '+(a.client||''),
      styles:{default:{document:{run:{font:FONT,size:21}}},
        paragraphStyles:[
          {id:'Heading1',name:'Heading 1',basedOn:'Normal',next:'Normal',quickFormat:true,run:{size:26,bold:true,color:NAVY,font:FONT},paragraph:{spacing:{before:280,after:120},outlineLevel:0}},
          {id:'Heading2',name:'Heading 2',basedOn:'Normal',next:'Normal',quickFormat:true,run:{size:22,bold:true,color:NAVY,font:FONT},paragraph:{spacing:{before:180,after:80},outlineLevel:1}}]},
      numbering:{config:[{reference:'b',levels:[{level:0,format:LevelFormat.BULLET,text:'•',alignment:AlignmentType.LEFT,style:{paragraph:{indent:{left:540,hanging:300}}}}]}]},
      sections:[{properties:{page:{size:{width:11906,height:16838},margin:{top:1100,bottom:1100,left:1440,right:1440}}},
        footers:{default:new Footer({children:[new Paragraph({alignment:AlignmentType.RIGHT,children:[new TextRun({text:`Cleanearth Water Services  |  ${ref}  |  Page `,size:16,color:GREY}),new TextRun({children:[PageNumber.CURRENT],size:16,color:GREY})]})]})},
        children:kids}]});
    const out=await Packer.toBlob(doc);out.photoCount=loaded.length;out.photoFailed=photoFailed;return out;
  };
})();
