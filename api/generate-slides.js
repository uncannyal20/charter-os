import pptxgen from 'pptxgenjs';

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    const S = req.body;
    const pres = new pptxgen();
    pres.layout = 'LAYOUT_16x9';

    const DARK='1F3864', ACCENT='4472C4', GREEN='0A6640', AMBER='D97706';
    const PURPLE='7B2FBE', TEAL='0D9488', WHITE='FFFFFF';
    const TEXT='1A1A2E', SUBTEXT='4A5168', GRAY='6B7280';
    const name = S.charter_name || 'Product Charter';
    const date = new Date().toLocaleDateString('en-GB',{year:'numeric',month:'long',day:'numeric'});
    const team = S.members || [];
    const kpis = (S.kpis && S.kpis.items) || [];
    const phases = (S.roadmap && S.roadmap.phases_data) || [];

    function cleanHtml(html) {
      if (!html) return '';
      return html.replace(/<[^>]*>/g,' ').replace(/\s+/g,' ').trim().slice(0,500);
    }

    function addHeader(slide, title, subtitle, color) {
      slide.addShape(pres.ShapeType.rect,{x:0,y:0,w:'100%',h:0.08,fill:{color:color||ACCENT}});
      slide.addShape(pres.ShapeType.rect,{x:0,y:4.9,w:'100%',h:0.72,fill:{color:'F8FAFF'}});
      slide.addText(title,{x:0.45,y:0.22,w:8.5,h:0.55,fontSize:22,bold:true,color:DARK,fontFace:'Arial'});
      if(subtitle) slide.addText(subtitle,{x:0.45,y:0.78,w:8.5,h:0.3,fontSize:11,color:GRAY,fontFace:'Arial'});
      slide.addText(name+' | '+date,{x:0.45,y:5.0,w:9.1,h:0.32,fontSize:8,color:GRAY,fontFace:'Arial'});
    }

    // Slide 1 - Cover
    const s1 = pres.addSlide();
    s1.addShape(pres.ShapeType.rect,{x:0,y:0,w:'100%',h:'100%',fill:{color:DARK}});
    s1.addShape(pres.ShapeType.rect,{x:0,y:0,w:0.12,h:'100%',fill:{color:ACCENT}});
    s1.addShape(pres.ShapeType.rect,{x:0,y:3.8,w:'100%',h:0.03,fill:{color:ACCENT}});
    s1.addText('PRODUCT CHARTER',{x:0.4,y:0.9,w:9.2,h:0.4,fontSize:11,bold:true,color:'7090C0',charSpacing:4,fontFace:'Arial'});
    s1.addText(name,{x:0.4,y:1.4,w:9.2,h:1.4,fontSize:32,bold:true,color:WHITE,fontFace:'Arial'});
    s1.addText('Management Presentation',{x:0.4,y:2.9,w:9.2,h:0.4,fontSize:14,color:'A0B8D8',fontFace:'Arial'});
    s1.addText('Prepared by Product Office  |  '+date,{x:0.4,y:4.1,w:9.2,h:0.3,fontSize:10,color:'6080A0',fontFace:'Arial'});
    if(team.length) s1.addText('Team: '+team.slice(0,4).map(m=>m.name).join(', ')+(team.length>4?' + '+(team.length-4)+' more':''),{x:0.4,y:4.5,w:9.2,h:0.3,fontSize:10,color:'6080A0',fontFace:'Arial'});

    // Slide 2 - Executive Summary
    const s2 = pres.addSlide();
    s2.addShape(pres.ShapeType.rect,{x:0,y:0,w:'100%',h:'100%',fill:{color:'FAFBFF'}});
    addHeader(s2,'Executive Summary','Charter at a glance',ACCENT);
    const stats=[{l:'Team Members',v:String(team.length),c:ACCENT},{l:'KPIs Defined',v:String(kpis.length),c:TEAL},{l:'Roadmap Phases',v:String(phases.length||3),c:GREEN},{l:'Domain',v:(S.vision&&S.vision.industry)||'—',c:AMBER}];
    stats.forEach((st,i)=>{
      const x=0.45+i*2.35;
      s2.addShape(pres.ShapeType.rect,{x,y:1.2,w:2.15,h:1.3,fill:{color:WHITE},line:{color:'E5E7EB',width:1}});
      s2.addText(st.v,{x,y:1.3,w:2.15,h:0.7,fontSize:st.v.length>8?16:26,bold:true,color:st.c,align:'center',fontFace:'Arial'});
      s2.addText(st.l,{x,y:2.05,w:2.15,h:0.3,fontSize:9,color:SUBTEXT,align:'center',fontFace:'Arial'});
    });
    if(S.vision&&S.vision.product_name){
      s2.addText('Product',{x:0.45,y:2.75,w:1.5,h:0.3,fontSize:9,bold:true,color:GRAY,fontFace:'Arial'});
      s2.addText(S.vision.product_name,{x:2.0,y:2.75,w:7.5,h:0.3,fontSize:10,color:TEXT,fontFace:'Arial'});
    }
    if(S.vision&&S.vision.users){
      s2.addText('Target Users',{x:0.45,y:3.15,w:1.5,h:0.3,fontSize:9,bold:true,color:GRAY,fontFace:'Arial'});
      s2.addText(S.vision.users,{x:2.0,y:3.15,w:7.5,h:0.3,fontSize:10,color:TEXT,fontFace:'Arial'});
    }

    // Slide 3 - Problem Statement
    const s3 = pres.addSlide();
    s3.addShape(pres.ShapeType.rect,{x:0,y:0,w:'100%',h:'100%',fill:{color:'FAFBFF'}});
    addHeader(s3,'Problem Statement','Why this product needs to exist',AMBER);
    if(S.problem&&S.problem.draft){
      const txt = cleanHtml(S.problem.draft);
      const sentences = txt.match(/[^.!?]+[.!?]+/g)||[txt];
      const bullets = sentences.slice(0,5).map(s=>s.trim()).filter(s=>s.length>20);
      s3.addText(bullets.map(b=>({text:'▸  '+b,options:{breakLine:true,paraSpaceBefore:8}})),{x:0.45,y:1.2,w:9.1,h:3.4,fontSize:11,color:TEXT,fontFace:'Arial',valign:'top'});
    } else {
      s3.addText('Problem Statement not yet defined.',{x:0.45,y:2.5,w:9.1,h:0.4,fontSize:12,color:AMBER,align:'center',fontFace:'Arial'});
    }

    // Slide 4 - Vision
    const s4 = pres.addSlide();
    s4.addShape(pres.ShapeType.rect,{x:0,y:0,w:'100%',h:'100%',fill:{color:'FAFBFF'}});
    addHeader(s4,'Vision & Strategy','Where we are going and how we will get there',PURPLE);
    if(S.vision&&S.vision.draft){
      const vtxt = cleanHtml(S.vision.draft);
      const first = vtxt.split(/[.!?]/)[0].trim();
      s4.addShape(pres.ShapeType.rect,{x:0.45,y:1.15,w:9.1,h:0.85,fill:{color:'F5F0FF'},line:{color:'C4B5FD',width:1}});
      s4.addText('"'+first+'"',{x:0.6,y:1.2,w:8.8,h:0.75,fontSize:12,italic:true,bold:true,color:PURPLE,fontFace:'Arial',valign:'middle'});
      s4.addText('STRATEGIC PILLARS',{x:0.45,y:2.15,w:9.1,h:0.28,fontSize:8,bold:true,color:GRAY,charSpacing:2,fontFace:'Arial'});
      const rest = vtxt.slice(first.length+1);
      const pillars = (rest.match(/[A-Z][^.!?]{20,80}[.!?]/g)||[]).slice(0,4);
      const pColors=[ACCENT,TEAL,GREEN,AMBER];
      pillars.forEach((p,i)=>{
        const x=0.45+(i%2)*4.65, y=2.5+Math.floor(i/2)*1.2;
        s4.addShape(pres.ShapeType.rect,{x,y,w:4.5,h:1.05,fill:{color:WHITE},line:{color:'E5E7EB',width:1}});
        s4.addShape(pres.ShapeType.rect,{x,y,w:4.5,h:0.06,fill:{color:pColors[i]}});
        s4.addText(p.slice(0,100),{x:x+0.1,y:y+0.12,w:4.3,h:0.85,fontSize:9,color:TEXT,fontFace:'Arial',valign:'top'});
      });
    } else {
      s4.addText('Vision not yet defined.',{x:0.45,y:2.5,w:9.1,h:0.4,fontSize:12,color:PURPLE,align:'center',fontFace:'Arial'});
    }

    // Slide 5 - KPIs
    const s5 = pres.addSlide();
    s5.addShape(pres.ShapeType.rect,{x:0,y:0,w:'100%',h:'100%',fill:{color:'FAFBFF'}});
    addHeader(s5,'Value Realisation Outcomes','How we measure success',TEAL);
    if(kpis.length){
      const catC={Adoption:ACCENT,Performance:PURPLE,Revenue:GREEN,Satisfaction:AMBER,Technical:TEAL,Efficiency:GREEN,Other:GRAY};
      kpis.slice(0,6).forEach((k,i)=>{
        const x=0.45+(i%2)*4.65, y=1.2+Math.floor(i/2)*1.2;
        const c=catC[k.category]||ACCENT;
        s5.addShape(pres.ShapeType.rect,{x,y,w:4.5,h:1.05,fill:{color:WHITE},line:{color:'E5E7EB',width:1}});
        s5.addShape(pres.ShapeType.rect,{x,y,w:0.06,h:1.05,fill:{color:c}});
        s5.addText(k.name.length>50?k.name.slice(0,48)+'...':k.name,{x:x+0.15,y:y+0.08,w:4.25,h:0.35,fontSize:9.5,bold:true,color:TEXT,fontFace:'Arial'});
        s5.addText('Target: '+(k.target||'TBD'),{x:x+0.15,y:y+0.46,w:2.8,h:0.28,fontSize:8.5,color:SUBTEXT,fontFace:'Arial'});
        s5.addText(k.timeframe||'',{x:x+3.1,y:y+0.46,w:1.3,h:0.28,fontSize:8,color:c,bold:true,align:'right',fontFace:'Arial'});
      });
      if(S.cba_parsed&&S.cba_parsed.totalBenAnnual){
        const p=S.cba_parsed;
        s5.addText('Annual Benefit: SGD '+(p.totalBenAnnual||0).toLocaleString()+'  |  Payback: '+(p.paybackMths||'—')+' months  |  3-Year ROI: '+(p.roi3yr||'—')+'%',{x:0.45,y:4.72,w:9.1,h:0.25,fontSize:8.5,bold:true,color:GREEN,fontFace:'Arial'});
      }
    } else {
      s5.addText('KPIs not yet defined.',{x:0.45,y:2.5,w:9.1,h:0.4,fontSize:12,color:TEAL,align:'center',fontFace:'Arial'});
    }

    // Slide 6 - CBA
    const chartImages = S.chartImages || {};
    if(S.cba && S.cba.draft){
      const s6cba = pres.addSlide();
      s6cba.addShape(pres.ShapeType.rect,{x:0,y:0,w:'100%',h:'100%',fill:{color:'FAFBFF'}});
      addHeader(s6cba,'Cost Benefit Analysis','Investment summary and value breakdown',TEAL);
      const p = S.cba_parsed || {};
      const cbaStats = [
        {l:'Total Investment',v:'SGD '+(p.totalY1||0).toLocaleString(),c:'C53030'},
        {l:'Annual Benefit',v:'SGD '+(p.totalBenAnnual||0).toLocaleString(),c:GREEN},
        {l:'Payback Period',v:(p.paybackMths||'—')+' mths',c:AMBER},
        {l:'3-Year ROI',v:(p.roi3yr||'—')+'%',c:TEAL},
      ];
      cbaStats.forEach((st,i)=>{
        const x=0.45+i*2.35;
        s6cba.addShape(pres.ShapeType.rect,{x,y:1.15,w:2.15,h:0.9,fill:{color:WHITE},line:{color:'E5E7EB',width:1}});
        s6cba.addText(st.v,{x,y:1.2,w:2.15,h:0.5,fontSize:st.v.length>12?13:16,bold:true,color:st.c,align:'center',fontFace:'Arial'});
        s6cba.addText(st.l,{x,y:1.72,w:2.15,h:0.25,fontSize:8.5,color:SUBTEXT,align:'center',fontFace:'Arial'});
      });
      if(chartImages.bar && chartImages.donut){
        const barData = chartImages.bar.replace(/^data:image\/png;base64,/,'');
        const donutData = chartImages.donut.replace(/^data:image\/png;base64,/,'');
        s6cba.addImage({data:'png;base64,'+barData, x:0.45, y:2.25, w:4.5, h:2.3});
        s6cba.addImage({data:'png;base64,'+donutData, x:5.1, y:2.25, w:4.5, h:2.3});
      } else {
        s6cba.addShape(pres.ShapeType.rect,{x:0.45,y:2.2,w:9.1,h:1.4,fill:{color:'FEF9EC'},line:{color:'FCD34D',width:1},rectRadius:0.1});
        s6cba.addText('⚠️  Charts not captured',{x:0.65,y:2.32,w:8.7,h:0.32,fontSize:11,bold:true,color:AMBER,fontFace:'Arial'});
        s6cba.addText('To include charts: Go to Step 4 → Cost Benefit Analysis → CBA View tab (let charts render) → Return here and click Generate Slides again.',{x:0.65,y:2.68,w:8.7,h:0.5,fontSize:9,color:TEXT,fontFace:'Arial'});
        s6cba.addText('The headline figures above are accurate. Only the visual charts are missing.',{x:0.65,y:3.28,w:8.7,h:0.24,fontSize:8.5,color:GRAY,italic:true,fontFace:'Arial'});
      }
    }

    // Slide 7 - Roadmap
    const s7 = pres.addSlide();
    s7.addShape(pres.ShapeType.rect,{x:0,y:0,w:'100%',h:'100%',fill:{color:'FAFBFF'}});
    addHeader(s7,'Product Roadmap','Phased delivery plan',GREEN);
    if(phases.length){
      const pColors=[ACCENT,GREEN,AMBER,TEAL,PURPLE];
      const colW=(9.1-(phases.length-1)*0.15)/phases.length;
      phases.forEach((ph,i)=>{
        const x=0.45+i*(colW+0.15);
        const c=ph.color||pColors[i%5];
        s7.addShape(pres.ShapeType.rect,{x,y:1.15,w:colW,h:0.55,fill:{color:c}});
        s7.addText('Phase '+(i+1),{x,y:1.17,w:colW,h:0.25,fontSize:8,bold:true,color:WHITE,align:'center',fontFace:'Arial'});
        s7.addText(ph.name,{x,y:1.4,w:colW,h:0.25,fontSize:8.5,bold:true,color:WHITE,align:'center',fontFace:'Arial'});
        s7.addText(ph.label||'',{x,y:1.76,w:colW,h:0.22,fontSize:8,color:GRAY,align:'center',fontFace:'Arial'});
        (ph.tasks||[]).slice(0,6).forEach((t,j)=>{
          const ty=2.08+j*0.44;
          s7.addShape(pres.ShapeType.rect,{x:x+0.04,y:ty,w:colW-0.08,h:0.38,fill:{color:WHITE},line:{color:'E5E7EB',width:0.5}});
          s7.addShape(pres.ShapeType.rect,{x:x+0.04,y:ty,w:0.05,h:0.38,fill:{color:c}});
          s7.addText(t.name.length>32?t.name.slice(0,30)+'...':t.name,{x:x+0.13,y:ty+0.04,w:colW-0.22,h:0.3,fontSize:8,color:TEXT,fontFace:'Arial',valign:'middle'});
        });
        if(ph.milestone){
          s7.addShape(pres.ShapeType.rect,{x:x+0.04,y:4.72,w:colW-0.08,h:0.25,fill:{color:'DCFCE7'},line:{color:GREEN,width:0.5}});
          s7.addText('🏁 '+ph.milestone.slice(0,40),{x:x+0.08,y:4.73,w:colW-0.16,h:0.22,fontSize:7,color:GREEN,fontFace:'Arial',valign:'middle'});
        }
      });
    } else {
      s7.addText('Roadmap not yet generated.',{x:0.45,y:2.5,w:9.1,h:0.4,fontSize:12,color:GREEN,align:'center',fontFace:'Arial'});
    }

    // Slide 8 - Team
    const s8 = pres.addSlide();
    s8.addShape(pres.ShapeType.rect,{x:0,y:0,w:'100%',h:'100%',fill:{color:'FAFBFF'}});
    addHeader(s8,'Product Office Team','Who is building this',DARK);
    if(team.length){
      const mColors=[ACCENT,GREEN,AMBER,PURPLE,TEAL,'C53030',ACCENT,GREEN];
      const cols=Math.min(4,team.length);
      const cardW=(9.1-(cols-1)*0.2)/cols;
      team.slice(0,8).forEach((m,i)=>{
        const x=0.45+(i%cols)*(cardW+0.2), y=1.25+Math.floor(i/cols)*1.7;
        const c=mColors[i%8];
        s8.addShape(pres.ShapeType.rect,{x,y,w:cardW,h:1.5,fill:{color:WHITE},line:{color:'E5E7EB',width:1}});
        s8.addShape(pres.ShapeType.ellipse,{x:x+(cardW/2)-0.28,y:y+0.1,w:0.56,h:0.56,fill:{color:c}});
        s8.addText(m.initials||(m.name||'').slice(0,2).toUpperCase(),{x:x+(cardW/2)-0.28,y:y+0.13,w:0.56,h:0.5,fontSize:13,bold:true,color:WHITE,align:'center',fontFace:'Arial'});
        s8.addText(m.name,{x:x+0.08,y:y+0.72,w:cardW-0.16,h:0.3,fontSize:9,bold:true,color:TEXT,align:'center',fontFace:'Arial'});
        s8.addText(m.role,{x:x+0.08,y:y+1.02,w:cardW-0.16,h:0.25,fontSize:8,color:c,align:'center',fontFace:'Arial'});
        s8.addText(m.dept||'',{x:x+0.08,y:y+1.26,w:cardW-0.16,h:0.2,fontSize:7.5,color:GRAY,align:'center',fontFace:'Arial'});
      });
    }

    // Slide 9 - Next Steps
    const s9 = pres.addSlide();
    s9.addShape(pres.ShapeType.rect,{x:0,y:0,w:'100%',h:'100%',fill:{color:'FAFBFF'}});
    addHeader(s9,'Next Steps & Actions Required','What we need from leadership',ACCENT);
    const actions = [
      {l:'Review and endorse Product Charter',o:'Steering Committee',u:'Required',c:DARK},
      {l:'Confirm budget and resource allocation',o:'Programme Director',u:'Required',c:DARK},
      ...(!S.problem||!S.problem.approved?[{l:'Approve Problem Statement',o:'Product Owner',u:'High',c:AMBER}]:[]),
      ...(!S.vision||!S.vision.approved?[{l:'Approve Vision & Strategy',o:'Product Owner',u:'High',c:AMBER}]:[]),
      ...(!kpis.length?[{l:'Define & Approve KPIs',o:'PM + PO',u:'High',c:'C53030'}]:[]),
      ...(!phases.length?[{l:'Generate & Approve Roadmap',o:'Tech Lead + PM',u:'Medium',c:ACCENT}]:[]),
    ];
    actions.slice(0,6).forEach((a,i)=>{
      const y=1.18+i*0.62;
      s9.addShape(pres.ShapeType.rect,{x:0.45,y,w:9.1,h:0.52,fill:{color:WHITE},line:{color:'E5E7EB',width:1}});
      s9.addShape(pres.ShapeType.rect,{x:0.45,y,w:0.06,h:0.52,fill:{color:a.c}});
      s9.addText(a.l,{x:0.65,y:y+0.08,w:6.0,h:0.28,fontSize:10.5,bold:true,color:TEXT,fontFace:'Arial'});
      s9.addText('Owner: '+a.o,{x:0.65,y:y+0.32,w:4.0,h:0.18,fontSize:8.5,color:GRAY,fontFace:'Arial'});
      s9.addShape(pres.ShapeType.rect,{x:7.7,y:y+0.1,w:1.7,h:0.3,fill:{color:a.c+'22'},line:{color:a.c,width:0.5}});
      s9.addText(a.u,{x:7.7,y:y+0.1,w:1.7,h:0.3,fontSize:8.5,bold:true,color:a.c,align:'center',fontFace:'Arial',valign:'middle'});
    });

    // Slide 10 - Closing
    const s10 = pres.addSlide();
    s10.addShape(pres.ShapeType.rect,{x:0,y:0,w:'100%',h:'100%',fill:{color:DARK}});
    s10.addShape(pres.ShapeType.rect,{x:0,y:0,w:0.12,h:'100%',fill:{color:ACCENT}});
    s10.addShape(pres.ShapeType.rect,{x:0,y:2.8,w:'100%',h:0.03,fill:{color:ACCENT}});
    s10.addText('Thank You',{x:0.4,y:1.2,w:9.2,h:0.8,fontSize:36,bold:true,color:WHITE,fontFace:'Arial'});
    s10.addText('Questions & Discussion',{x:0.4,y:2.1,w:9.2,h:0.4,fontSize:16,color:'A0B8D8',fontFace:'Arial'});
    s10.addText(name+' — Product Charter | '+date,{x:0.4,y:3.2,w:9.2,h:0.35,fontSize:11,color:'6080A0',fontFace:'Arial'});
    s10.addText('Generated using Charter OS',{x:0.4,y:4.8,w:9.2,h:0.25,fontSize:8,color:'405070',italic:true,fontFace:'Arial'});

    // Return as buffer
    const buf = await pres.write({outputType:'nodebuffer'});
    res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.presentationml.presentation');
    res.setHeader('Content-Disposition','attachment; filename="management-presentation.pptx"');
    return res.send(buf);

  } catch(err) {
    console.error('Slides error:', err);
    return res.status(500).json({ error: err.message });
  }
}
