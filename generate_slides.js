const pptxgen = require("pptxgenjs");
const fs = require("fs");

// Read charter data passed as argument
const charterData = JSON.parse(process.argv[2]);
const S = charterData;

const pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.title = (S.charter_name || 'Product Charter') + ' — Management Presentation';

// ── Design System ──
const DARK    = '1F3864';
const ACCENT  = '4472C4';
const GREEN   = '0A6640';
const AMBER   = 'D97706';
const PURPLE  = '7B2FBE';
const TEAL    = '0D9488';
const WHITE   = 'FFFFFF';
const LIGHT   = 'F0F4FF';
const TEXT    = '1A1A2E';
const SUBTEXT = '4A5168';
const GRAY    = '6B7280';
const name    = S.charter_name || 'Product Charter';
const date    = new Date().toLocaleDateString('en-GB', {year:'numeric',month:'long',day:'numeric'});
const team    = S.members || [];
const kpis    = (S.kpis && S.kpis.items) || [];
const phases  = (S.roadmap && S.roadmap.phases_data) || [];

function cleanHtml(html) {
  if (!html) return '';
  return html.replace(/<[^>]*>/g,' ').replace(/\s+/g,' ').trim().slice(0,600);
}

function addSlideHeader(slide, title, subtitle, accentColor) {
  slide.addShape(pres.ShapeType.rect, {x:0,y:0,w:'100%',h:0.08,fill:{color:accentColor||ACCENT}});
  slide.addShape(pres.ShapeType.rect, {x:0,y:4.9,w:'100%',h:0.72,fill:{color:'F8FAFF'}});
  slide.addText(title, {x:0.45,y:0.22,w:8.5,h:0.55,fontSize:22,bold:true,color:DARK,fontFace:'Arial'});
  if (subtitle) slide.addText(subtitle, {x:0.45,y:0.78,w:8.5,h:0.3,fontSize:11,color:GRAY,fontFace:'Arial'});
  slide.addText(name + '  |  ' + date, {x:0.45,y:5.0,w:9.1,h:0.32,fontSize:8,color:GRAY,fontFace:'Arial'});
}

function addSlideDivider(slide, y) {
  slide.addShape(pres.ShapeType.rect, {x:0.45,y,w:9.1,h:0.02,fill:{color:'E5E7EB'}});
}

// ════════════════════════════════════════
// SLIDE 1 — COVER
// ════════════════════════════════════════
const s1 = pres.addSlide();
s1.addShape(pres.ShapeType.rect, {x:0,y:0,w:'100%',h:'100%',fill:{color:DARK}});
s1.addShape(pres.ShapeType.rect, {x:0,y:0,w:0.12,h:'100%',fill:{color:ACCENT}});
s1.addShape(pres.ShapeType.rect, {x:0,y:3.8,w:'100%',h:0.03,fill:{color:ACCENT}});
s1.addText('PRODUCT CHARTER', {x:0.4,y:0.9,w:9.2,h:0.4,fontSize:11,bold:true,color:'7090C0',charSpacing:4,fontFace:'Arial'});
s1.addText(name, {x:0.4,y:1.4,w:9.2,h:1.4,fontSize:34,bold:true,color:WHITE,fontFace:'Arial',lineSpacingMultiple:1.1});
s1.addText('Management Presentation', {x:0.4,y:2.9,w:9.2,h:0.4,fontSize:14,color:'A0B8D8',fontFace:'Arial'});
s1.addText('Prepared by Product Office  |  ' + date, {x:0.4,y:4.1,w:9.2,h:0.3,fontSize:10,color:'6080A0',fontFace:'Arial'});
if (team.length) {
  s1.addText('Team: ' + team.slice(0,4).map(m=>m.name).join(', ') + (team.length>4?' + '+(team.length-4)+' more':''), {x:0.4,y:4.5,w:9.2,h:0.3,fontSize:10,color:'6080A0',fontFace:'Arial'});
}
const overall = Math.round([['team',S.members&&S.members.length?60:0],['problem',S.problem&&S.problem.draft?80:0],['vision',S.vision&&S.vision.draft?80:0],['kpis',kpis.length?80:0],['roadmap',phases.length?80:0]].reduce((a,b)=>a+b[1],0)/5);
s1.addText(overall + '% Charter Complete', {x:0.4,y:4.85,w:9.2,h:0.3,fontSize:9,color:'5070A0',italic:true,fontFace:'Arial'});

// ════════════════════════════════════════
// SLIDE 2 — EXECUTIVE SUMMARY
// ════════════════════════════════════════
const s2 = pres.addSlide();
s2.addShape(pres.ShapeType.rect, {x:0,y:0,w:'100%',h:'100%',fill:{color:'FAFBFF'}});
addSlideHeader(s2, 'Executive Summary', 'Charter at a glance', ACCENT);

const stats = [
  {label:'Team Members', val:String(team.length), color:ACCENT},
  {label:'KPIs Defined', val:String(kpis.length), color:TEAL},
  {label:'Roadmap Phases', val:String(phases.length||3), color:GREEN},
  {label:'Charter Progress', val:overall+'%', color:AMBER},
];
stats.forEach((stat,i) => {
  const x = 0.45 + i * 2.35;
  s2.addShape(pres.ShapeType.rect, {x,y:1.2,w:2.15,h:1.3,fill:{color:'FFFFFF'},line:{color:'E5E7EB',width:1},rectRadius:0.1});
  s2.addText(stat.val, {x,y:1.3,w:2.15,h:0.7,fontSize:28,bold:true,color:stat.color,align:'center',fontFace:'Arial'});
  s2.addText(stat.label, {x,y:2.05,w:2.15,h:0.3,fontSize:9,color:SUBTEXT,align:'center',fontFace:'Arial'});
});

// Product name & domain
if (S.vision && S.vision.product_name) {
  s2.addText('Product', {x:0.45,y:2.75,w:1.5,h:0.3,fontSize:9,bold:true,color:GRAY,fontFace:'Arial'});
  s2.addText(S.vision.product_name, {x:2.0,y:2.75,w:7.5,h:0.3,fontSize:10,color:TEXT,fontFace:'Arial'});
  addSlideDivider(s2, 3.1);
}
if (S.vision && S.vision.industry) {
  s2.addText('Domain', {x:0.45,y:3.2,w:1.5,h:0.3,fontSize:9,bold:true,color:GRAY,fontFace:'Arial'});
  s2.addText(S.vision.industry, {x:2.0,y:3.2,w:7.5,h:0.3,fontSize:10,color:TEXT,fontFace:'Arial'});
  addSlideDivider(s2, 3.55);
}
if (S.vision && S.vision.users) {
  s2.addText('Target Users', {x:0.45,y:3.65,w:1.5,h:0.3,fontSize:9,bold:true,color:GRAY,fontFace:'Arial'});
  s2.addText(S.vision.users, {x:2.0,y:3.65,w:7.5,h:0.3,fontSize:10,color:TEXT,fontFace:'Arial'});
}

// ════════════════════════════════════════
// SLIDE 3 — PROBLEM STATEMENT
// ════════════════════════════════════════
const s3 = pres.addSlide();
s3.addShape(pres.ShapeType.rect, {x:0,y:0,w:'100%',h:'100%',fill:{color:'FAFBFF'}});
addSlideHeader(s3, 'Problem Statement', 'Why this product needs to exist', AMBER);

if (S.problem && S.problem.draft) {
  const problemText = cleanHtml(S.problem.draft);
  // Split into lines for bullet-style display
  const sentences = problemText.match(/[^.!?]+[.!?]+/g) || [problemText];
  const bullets = sentences.slice(0,6).map(s => s.trim()).filter(s => s.length > 20);
  s3.addText(bullets.map(b => ({text:'▸  ' + b, options:{breakLine:true, paraSpaceBefore:6}})), {
    x:0.45,y:1.2,w:9.1,h:3.4,fontSize:11,color:TEXT,fontFace:'Arial',valign:'top'
  });
} else {
  s3.addShape(pres.ShapeType.rect, {x:0.45,y:1.5,w:9.1,h:1.5,fill:{color:'FEF9EC'},line:{color:'FCD34D',width:1}});
  s3.addText('Problem Statement not yet defined.\nComplete Step 2 in Charter OS to generate this content.', {x:0.6,y:1.7,w:8.8,h:1.0,fontSize:11,color:AMBER,align:'center',fontFace:'Arial'});
}

// ════════════════════════════════════════
// SLIDE 4 — VISION & STRATEGY
// ════════════════════════════════════════
const s4 = pres.addSlide();
s4.addShape(pres.ShapeType.rect, {x:0,y:0,w:'100%',h:'100%',fill:{color:'FAFBFF'}});
addSlideHeader(s4, 'Vision & Strategy', 'Where we are going and how we will get there', PURPLE);

if (S.vision && S.vision.draft) {
  const visionText = cleanHtml(S.vision.draft);
  // Extract vision statement (usually first sentence)
  const firstSentence = visionText.split(/[.!?]/)[0].trim();
  // Vision callout box
  s4.addShape(pres.ShapeType.rect, {x:0.45,y:1.15,w:9.1,h:0.85,fill:{color:'F5F0FF'},line:{color:'C4B5FD',width:1},rectRadius:0.08});
  s4.addText('"' + firstSentence + '"', {x:0.6,y:1.2,w:8.8,h:0.75,fontSize:12,italic:true,bold:true,color:PURPLE,fontFace:'Arial',valign:'middle'});
  // Strategic pillars
  s4.addText('STRATEGIC PILLARS', {x:0.45,y:2.15,w:9.1,h:0.28,fontSize:8,bold:true,color:GRAY,charSpacing:2,fontFace:'Arial'});
  // Extract pillar-like content
  const remaining = visionText.slice(firstSentence.length+1);
  const pillars = remaining.match(/[A-Z][^.!?]{20,80}[.!?]/g) || [];
  const displayPillars = pillars.slice(0,5);
  const pillarsPerRow = displayPillars.length > 3 ? 3 : displayPillars.length || 3;
  const pillarColors = [ACCENT, TEAL, GREEN, AMBER, PURPLE];
  displayPillars.forEach((pillar,i) => {
    const col = i % pillarsPerRow;
    const row = Math.floor(i / pillarsPerRow);
    const w = (9.1 - (pillarsPerRow-1)*0.15) / pillarsPerRow;
    const x = 0.45 + col*(w+0.15);
    const y = 2.5 + row*1.2;
    s4.addShape(pres.ShapeType.rect, {x,y,w,h:1.05,fill:{color:'FFFFFF'},line:{color:'E5E7EB',width:1},rectRadius:0.08});
    s4.addShape(pres.ShapeType.rect, {x,y,w,h:0.06,fill:{color:pillarColors[i%5]},rectRadius:0.04});
    s4.addText(pillar.slice(0,90), {x:x+0.1,y:y+0.12,w:w-0.2,h:0.85,fontSize:9,color:TEXT,fontFace:'Arial',valign:'top'});
  });
  if (!displayPillars.length) {
    s4.addText(remaining.slice(0,400), {x:0.45,y:2.5,w:9.1,h:2.0,fontSize:10,color:TEXT,fontFace:'Arial',valign:'top'});
  }
} else {
  s4.addShape(pres.ShapeType.rect, {x:0.45,y:1.5,w:9.1,h:1.5,fill:{color:'F5F0FF'},line:{color:'C4B5FD',width:1}});
  s4.addText('Vision & Strategy not yet defined.\nComplete Step 3 in Charter OS.', {x:0.6,y:1.7,w:8.8,h:1.0,fontSize:11,color:PURPLE,align:'center',fontFace:'Arial'});
}

// ════════════════════════════════════════
// SLIDE 5 — VALUE REALISATION & KPIs
// ════════════════════════════════════════
const s5 = pres.addSlide();
s5.addShape(pres.ShapeType.rect, {x:0,y:0,w:'100%',h:'100%',fill:{color:'FAFBFF'}});
addSlideHeader(s5, 'Value Realisation Outcomes', 'How we measure success', TEAL);

if (kpis.length) {
  const catColors = {Adoption:ACCENT,Performance:PURPLE,Revenue:GREEN,Satisfaction:AMBER,Technical:TEAL,Efficiency:GREEN,Other:GRAY};
  const displayKpis = kpis.slice(0,6);
  displayKpis.forEach((k,i) => {
    const col = i % 2;
    const row = Math.floor(i/2);
    const x = 0.45 + col*4.65;
    const y = 1.2 + row*1.2;
    const color = catColors[k.category] || ACCENT;
    s5.addShape(pres.ShapeType.rect, {x,y,w:4.5,h:1.05,fill:{color:'FFFFFF'},line:{color:'E5E7EB',width:1},rectRadius:0.06});
    s5.addShape(pres.ShapeType.rect, {x,y,w:0.06,h:1.05,fill:{color:color},rectRadius:0.04});
    s5.addText(k.name.length>50?k.name.slice(0,48)+'...':k.name, {x:x+0.15,y:y+0.08,w:4.25,h:0.35,fontSize:9.5,bold:true,color:TEXT,fontFace:'Arial'});
    s5.addText('Target: ' + (k.target||'TBD'), {x:x+0.15,y:y+0.46,w:2.8,h:0.28,fontSize:8.5,color:SUBTEXT,fontFace:'Arial'});
    s5.addText(k.timeframe||'', {x:x+3.1,y:y+0.46,w:1.3,h:0.28,fontSize:8,color:color,bold:true,align:'right',fontFace:'Arial'});
    s5.addText(k.category||'', {x:x+0.15,y:y+0.74,w:2.0,h:0.22,fontSize:7.5,color:color,fontFace:'Arial'});
  });
  if (kpis.length > 6) {
    s5.addText('+ ' + (kpis.length-6) + ' additional KPIs — see full charter', {x:0.45,y:4.72,w:9.1,h:0.25,fontSize:8.5,color:GRAY,italic:true,fontFace:'Arial'});
  }
  // CBA headline if available
  if (S.cba_parsed && S.cba_parsed.totalBenAnnual) {
    const p = S.cba_parsed;
    s5.addText('Annual Benefit: SGD ' + (p.totalBenAnnual||0).toLocaleString() + '  |  Payback: ' + (p.paybackMths||'—') + ' months  |  3-Year ROI: ' + (p.roi3yr||'—') + '%', {
      x:0.45,y:4.72,w:9.1,h:0.25,fontSize:8.5,bold:true,color:GREEN,fontFace:'Arial'
    });
  }
} else {
  s5.addShape(pres.ShapeType.rect, {x:0.45,y:1.5,w:9.1,h:1.5,fill:{color:'F0FDFB'},line:{color:'5EEAD4',width:1}});
  s5.addText('KPIs not yet defined.\nComplete Step 4 in Charter OS.', {x:0.6,y:1.7,w:8.8,h:1.0,fontSize:11,color:TEAL,align:'center',fontFace:'Arial'});
}

// ════════════════════════════════════════
// SLIDE 6 — PRODUCT ROADMAP
// ════════════════════════════════════════
const s6 = pres.addSlide();
s6.addShape(pres.ShapeType.rect, {x:0,y:0,w:'100%',h:'100%',fill:{color:'FAFBFF'}});
addSlideHeader(s6, 'Product Roadmap', 'Phased delivery plan', GREEN);

if (phases.length) {
  const phaseColors = [ACCENT, GREEN, AMBER, TEAL, PURPLE];
  const colW = (9.1 - (phases.length-1)*0.15) / phases.length;
  phases.forEach((ph,i) => {
    const x = 0.45 + i*(colW+0.15);
    const color = ph.color || phaseColors[i%5];
    // Phase header
    s6.addShape(pres.ShapeType.rect, {x,y:1.15,w:colW,h:0.55,fill:{color},rectRadius:0.06});
    s6.addText('Phase ' + (i+1), {x,y:1.17,w:colW,h:0.25,fontSize:8,bold:true,color:WHITE,align:'center',fontFace:'Arial'});
    s6.addText(ph.name, {x,y:1.4,w:colW,h:0.25,fontSize:8.5,bold:true,color:WHITE,align:'center',fontFace:'Arial'});
    // Label
    s6.addText(ph.label||'', {x,y:1.76,w:colW,h:0.22,fontSize:8,color:GRAY,align:'center',fontFace:'Arial'});
    // Tasks
    const displayTasks = (ph.tasks||[]).slice(0,6);
    displayTasks.forEach((t,j) => {
      const ty = 2.08 + j*0.44;
      s6.addShape(pres.ShapeType.rect, {x:x+0.04,y:ty,w:colW-0.08,h:0.38,fill:{color:'FFFFFF'},line:{color:'E5E7EB',width:0.5},rectRadius:0.04});
      s6.addShape(pres.ShapeType.rect, {x:x+0.04,y:ty,w:0.05,h:0.38,fill:{color},rectRadius:0.03});
      s6.addText(t.name.length>32?t.name.slice(0,30)+'...':t.name, {x:x+0.13,y:ty+0.04,w:colW-0.22,h:0.3,fontSize:8,color:TEXT,fontFace:'Arial',valign:'middle'});
    });
    if ((ph.tasks||[]).length > 6) {
      s6.addText('+'+(ph.tasks.length-6)+' more', {x,y:2.08+6*0.44,w:colW,h:0.25,fontSize:7.5,color:GRAY,italic:true,align:'center',fontFace:'Arial'});
    }
    // Milestone
    if (ph.milestone) {
      s6.addShape(pres.ShapeType.rect, {x:x+0.04,y:4.72,w:colW-0.08,h:0.25,fill:{color:'DCFCE7'},line:{color:GREEN,width:0.5},rectRadius:0.04});
      s6.addText('🏁 ' + ph.milestone.slice(0,40), {x:x+0.08,y:4.73,w:colW-0.16,h:0.22,fontSize:7,color:GREEN,fontFace:'Arial',valign:'middle'});
    }
  });
} else {
  s6.addShape(pres.ShapeType.rect, {x:0.45,y:1.5,w:9.1,h:1.5,fill:{color:'F0FDF4'},line:{color:'86EFAC',width:1}});
  s6.addText('Roadmap not yet generated.\nComplete Step 5 in Charter OS.', {x:0.6,y:1.7,w:8.8,h:1.0,fontSize:11,color:GREEN,align:'center',fontFace:'Arial'});
}

// ════════════════════════════════════════
// SLIDE 7 — TEAM
// ════════════════════════════════════════
const s7 = pres.addSlide();
s7.addShape(pres.ShapeType.rect, {x:0,y:0,w:'100%',h:'100%',fill:{color:'FAFBFF'}});
addSlideHeader(s7, 'Product Office Team', 'Who is building this', DARK);

if (team.length) {
  const memberColors = [ACCENT,GREEN,AMBER,PURPLE,TEAL,'E53E3E',ACCENT,GREEN];
  const cols = Math.min(4, team.length);
  const cardW = (9.1-(cols-1)*0.2)/cols;
  team.slice(0,8).forEach((m,i) => {
    const col = i % cols;
    const row = Math.floor(i/cols);
    const x = 0.45 + col*(cardW+0.2);
    const y = 1.25 + row*1.7;
    const color = memberColors[i%8];
    s7.addShape(pres.ShapeType.rect, {x,y,w:cardW,h:1.5,fill:{color:'FFFFFF'},line:{color:'E5E7EB',width:1},rectRadius:0.08});
    // Avatar circle
    s7.addShape(pres.ShapeType.ellipse, {x:x+(cardW/2)-0.28,y:y+0.1,w:0.56,h:0.56,fill:{color},line:{color,width:0}});
    s7.addText(m.initials||m.name.slice(0,2).toUpperCase(), {x:x+(cardW/2)-0.28,y:y+0.13,w:0.56,h:0.5,fontSize:13,bold:true,color:WHITE,align:'center',fontFace:'Arial'});
    s7.addText(m.name, {x:x+0.08,y:y+0.72,w:cardW-0.16,h:0.3,fontSize:9,bold:true,color:TEXT,align:'center',fontFace:'Arial'});
    s7.addText(m.role, {x:x+0.08,y:y+1.02,w:cardW-0.16,h:0.25,fontSize:8,color:color,align:'center',fontFace:'Arial'});
    s7.addText(m.dept||'', {x:x+0.08,y:y+1.26,w:cardW-0.16,h:0.2,fontSize:7.5,color:GRAY,align:'center',fontFace:'Arial'});
  });
} else {
  s7.addText('No team members added yet.', {x:0.45,y:2.5,w:9.1,h:0.4,fontSize:12,color:GRAY,align:'center',fontFace:'Arial'});
}

// ════════════════════════════════════════
// SLIDE 8 — NEXT STEPS
// ════════════════════════════════════════
const s8 = pres.addSlide();
s8.addShape(pres.ShapeType.rect, {x:0,y:0,w:'100%',h:'100%',fill:{color:'FAFBFF'}});
addSlideHeader(s8, 'Next Steps & Actions Required', 'What we need from leadership', ACCENT);

const incompleteItems = [];
if (!S.problem || !S.problem.approved) incompleteItems.push({label:'Approve Problem Statement', owner:'Product Owner', urgency:'High', color:AMBER});
if (!S.vision || !S.vision.approved) incompleteItems.push({label:'Approve Vision & Strategy', owner:'Product Owner', urgency:'High', color:AMBER});
if (!kpis.length) incompleteItems.push({label:'Define & Approve KPIs', owner:'PM + PO', urgency:'High', color:'E53E3E'});
if (!phases.length) incompleteItems.push({label:'Generate & Approve Roadmap', owner:'Tech Lead + PM', urgency:'Medium', color:ACCENT});
if (!S.cba || !S.cba.draft) incompleteItems.push({label:'Complete Cost Benefit Analysis', owner:'PM + Finance', urgency:'Medium', color:TEAL});
if (incompleteItems.length === 0) incompleteItems.push({label:'Charter complete — proceed to kick-off!', owner:'All', urgency:'Done', color:GREEN});

// Always add management actions
const mgmtActions = [
  {label:'Review and endorse Product Charter', owner:'Steering Committee', urgency:'Required', color:DARK},
  {label:'Confirm budget and resource allocation', owner:'Programme Director', urgency:'Required', color:DARK},
  {label:'Identify and onboard design partners for beta', owner:'PM', urgency:'Month 1', color:ACCENT},
];

const allItems = [...incompleteItems, ...mgmtActions];
allItems.slice(0,6).forEach((item,i) => {
  const y = 1.18 + i*0.62;
  s8.addShape(pres.ShapeType.rect, {x:0.45,y,w:9.1,h:0.52,fill:{color:'FFFFFF'},line:{color:'E5E7EB',width:1},rectRadius:0.06});
  s8.addShape(pres.ShapeType.rect, {x:0.45,y,w:0.06,h:0.52,fill:{color:item.color},rectRadius:0.04});
  s8.addText(item.label, {x:0.65,y:y+0.08,w:6.0,h:0.28,fontSize:10.5,bold:true,color:TEXT,fontFace:'Arial'});
  s8.addText('Owner: ' + item.owner, {x:0.65,y:y+0.32,w:4.0,h:0.18,fontSize:8.5,color:GRAY,fontFace:'Arial'});
  s8.addShape(pres.ShapeType.rect, {x:7.7,y:y+0.1,w:1.7,h:0.3,fill:{color:item.color+'22'},line:{color:item.color,width:0.5},rectRadius:0.04});
  s8.addText(item.urgency, {x:7.7,y:y+0.1,w:1.7,h:0.3,fontSize:8.5,bold:true,color:item.color,align:'center',fontFace:'Arial',valign:'middle'});
});

// ════════════════════════════════════════
// SLIDE 9 — CLOSING
// ════════════════════════════════════════
const s9 = pres.addSlide();
s9.addShape(pres.ShapeType.rect, {x:0,y:0,w:'100%',h:'100%',fill:{color:DARK}});
s9.addShape(pres.ShapeType.rect, {x:0,y:0,w:0.12,h:'100%',fill:{color:ACCENT}});
s9.addShape(pres.ShapeType.rect, {x:0,y:2.8,w:'100%',h:0.03,fill:{color:ACCENT}});
s9.addText('Thank You', {x:0.4,y:1.2,w:9.2,h:0.8,fontSize:36,bold:true,color:WHITE,fontFace:'Arial'});
s9.addText('Questions & Discussion', {x:0.4,y:2.1,w:9.2,h:0.4,fontSize:16,color:'A0B8D8',fontFace:'Arial'});
s9.addText(name + '  —  Product Charter', {x:0.4,y:3.2,w:9.2,h:0.35,fontSize:11,color:'6080A0',fontFace:'Arial'});
s9.addText('Prepared by Product Office  |  ' + date, {x:0.4,y:3.6,w:9.2,h:0.3,fontSize:10,color:'506080',fontFace:'Arial'});
s9.addText('Generated using Charter OS', {x:0.4,y:4.8,w:9.2,h:0.25,fontSize:8,color:'405070',italic:true,fontFace:'Arial'});

// Save
const filename = '/home/claude/sample-inputs/summary_slides.pptx';
pres.writeFile({fileName: filename})
  .then(() => console.log('DONE:' + filename))
  .catch(e => { console.error('ERROR:' + e.message); process.exit(1); });
