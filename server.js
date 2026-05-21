const express = require('express');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType,
  LevelFormat, ExternalHyperlink, PageNumber, Footer,
  VerticalAlign } = require('docx');

const app = express();
app.use(express.json({ limit: '10mb' }));

const COL = [696, 1424, 1560, 4125, 1545];
const TABLE_W = COL.reduce((a,b)=>a+b,0);
const BLUE='0070C0', PURPLE='7030A0', LINK_BLUE='0000FF', BLACK='000000', RED_FAIL='C00000';

function isRegalado(item){ return /regalado/i.test(item.sponsors||''); }
function accent(item){ return isRegalado(item)?PURPLE:BLUE; }
function bdr(sz=4){ return { style:BorderStyle.SINGLE, size:sz, color:'auto' }; }
function noBdr(){ return { style:BorderStyle.NONE, size:0, color:'FFFFFF' }; }
const STD_BORDERS={ top:bdr(), bottom:bdr(), left:bdr(), right:bdr() };
const NO_BORDERS={ top:noBdr(), bottom:noBdr(), left:noBdr(), right:noBdr() };
const STD_MARGINS={ top:0, bottom:0, left:108, right:108 };
const BRIEF_MARGINS={ top:60, bottom:60, left:108, right:108 };

function txt(text, opts={}){
  return new TextRun({ text, font:'Arial', size:opts.size||20, bold:opts.bold||false,
    italics:opts.italics||false, color:opts.color||BLACK, underline:opts.underline?{}:undefined });
}
function p(children, opts={}){
  const runs=Array.isArray(children)?children:typeof children==='string'?[txt(children,opts)]:[children];
  return new Paragraph({ alignment:opts.align||AlignmentType.LEFT,
    spacing:{ before:opts.before||0, after:opts.after||0 },
    numbering:opts.bullet?{ reference:'bullets', level:opts.level||0 }:undefined, children:runs });
}
function emptyP(){ return p(''); }
function inlineRuns(text, baseOpts={}){
  return text.split(/(\*\*[^*]+\*\*)/g).filter(Boolean).map(part=>
    /^\*\*[^*]+\*\*$/.test(part)
      ? txt(part.replace(/\*\*/g,''),{...baseOpts,bold:true})
      : txt(part,baseOpts));
}
function parseBrief(briefText, color){
  if(!briefText) return [p([txt('Brief not available.',{color:RED_FAIL,italics:true,size:20})])];
  const paras=[];
  for(const raw of briefText.split('\n')){
    const line=raw.trimEnd(), trimmed=line.trim();
    if(!trimmed){ paras.push(p('',{after:60})); continue; }
    if(/^⭐/.test(trimmed)){ paras.push(p(inlineRuns(trimmed,{bold:true,color,size:20}),{before:60,after:40})); continue; }
    if(/^## /.test(line)){ paras.push(p([txt(line.replace(/^## /,'').trim(),{bold:true,color,size:20})],{before:120,after:40})); continue; }
    if(/^### /.test(line)){ paras.push(p([txt(line.replace(/^### /,'').trim(),{bold:true,color,size:20})],{before:80,after:20})); continue; }
    if(/^[\-•\*] /.test(trimmed)){ paras.push(p(inlineRuns(trimmed.replace(/^[\-•\*] /,''),{size:20,color}),{bullet:true,level:0,before:20,after:20})); continue; }
    if(/^\s{2,}[\-•]/.test(line)){ paras.push(p(inlineRuns(trimmed.replace(/^[\-•] /,''),{size:20,color}),{bullet:true,level:1,before:10,after:10})); continue; }
    paras.push(p(inlineRuns(trimmed,{size:20,color}),{before:20,after:20}));
  }
  return paras;
}
function tc(children, opts={}){
  return new TableCell({ width:opts.w!==undefined?{size:opts.w,type:WidthType.DXA}:undefined,
    columnSpan:opts.span, verticalAlign:opts.vAlign||VerticalAlign.TOP,
    shading:opts.fill?{fill:opts.fill,type:ShadingType.CLEAR}:undefined,
    margins:opts.margins||STD_MARGINS, borders:opts.borders||STD_BORDERS,
    children:Array.isArray(children)?children:[children] });
}

async function buildDoc(data){
  const { meetingTitle='Miami-Dade Board of County Commissioners',
    meetingVersion='PRELIMINARY', meetingDate='', meetingTime='',
    meetingLocation='Commission Chambers', items=[] } = data;

  const children=[];

  // Cover header
  const titleParas=[
    p([txt(meetingTitle,{bold:true,color:BLUE,size:26})],{align:AlignmentType.CENTER,after:60}),
    p([txt(meetingVersion+' Version',{bold:true,color:BLUE,size:20,italics:true})],{align:AlignmentType.CENTER,after:80}),
    p([txt(meetingDate,{bold:true,color:BLUE,size:24})],{align:AlignmentType.CENTER,after:40}),
    p([txt(meetingTime,{color:BLUE,size:20})],{align:AlignmentType.CENTER,after:40}),
    p([txt(meetingLocation,{color:BLUE,size:20})],{align:AlignmentType.CENTER,after:0}),
  ];
  children.push(new Table({
    width:{size:9360,type:WidthType.DXA}, columnWidths:[9360],
    borders:{top:noBdr(),bottom:noBdr(),left:noBdr(),right:noBdr(),insideH:noBdr(),insideV:noBdr()},
    rows:[new TableRow({children:[tc(titleParas,{w:9360,borders:NO_BORDERS,margins:{top:0,bottom:0,left:0,right:0}})]})]
  }));
  children.push(emptyP()); children.push(emptyP());

  // Agenda table
  const agendaRows=[];
  let lastSection='';

  for(const item of items){
    const color=accent(item);
    const isFailed=item.status==='failed';
    const isSection=!item.matterId;

    if(item.sectionLabel && item.sectionLabel!==lastSection){
      lastSection=item.sectionLabel;
      const parts=item.sectionLabel.match(/^(\S+)\s+(.+)$/)||['','',item.sectionLabel];
      agendaRows.push(new TableRow({children:[
        tc([p([txt(parts[1],{bold:true,size:20})])],{w:COL[0],borders:STD_BORDERS}),
        tc([p([txt(parts[2],{bold:true,size:20})])],{w:COL[1]+COL[2]+COL[3],span:3,borders:STD_BORDERS}),
        tc([emptyP()],{w:COL[4],borders:STD_BORDERS}),
      ]}));
    }
    if(isSection) continue;

    // Metadata row
    agendaRows.push(new TableRow({children:[
      tc([p([txt(item.itemNumber||'',{bold:true,size:20})])],{w:COL[0],borders:STD_BORDERS,vAlign:VerticalAlign.CENTER}),
      tc([p([new ExternalHyperlink({link:item.matterUrl||'#',children:[txt((item.matterId||'')+'  ',{color:LINK_BLUE,underline:true,size:20})]})])],{w:COL[1],borders:STD_BORDERS,vAlign:VerticalAlign.CENTER}),
      tc([p([txt(item.fileType||'',{bold:true,size:20})])],{w:COL[2],borders:STD_BORDERS,vAlign:VerticalAlign.CENTER}),
      tc([p([txt(item.sponsors||'',{bold:true,size:20})])],{w:COL[3],borders:STD_BORDERS,vAlign:VerticalAlign.CENTER}),
      tc([emptyP()],{w:COL[4],borders:STD_BORDERS}),
    ]}));

    // Title + brief row
    const contentParas=[];
    contentParas.push(p([txt(item.title||'',{color:BLACK,size:20})],{after:80}));
    if(isFailed){
      contentParas.push(emptyP());
      contentParas.push(p([txt('⚠ Document could not be retrieved after 3 attempts. ',{bold:true,color:RED_FAIL,size:20}),txt('Brief not available — edit this document to add manually.',{italics:true,color:RED_FAIL,size:20})],{before:60,after:40}));
      contentParas.push(p([txt('Source URL: ',{bold:true,size:20,color}),new ExternalHyperlink({link:item.matterUrl||'#',children:[txt(item.matterUrl||'',{color:LINK_BLUE,underline:true,size:20})]})],{after:60}));
    } else {
      contentParas.push(...parseBrief(item.brief,color));
    }

    agendaRows.push(new TableRow({children:[
      tc([emptyP()],{w:COL[0],borders:STD_BORDERS}),
      tc(contentParas,{w:COL[1]+COL[2]+COL[3],span:3,borders:STD_BORDERS,margins:BRIEF_MARGINS}),
      tc(item.pdfUrl?[p([new ExternalHyperlink({link:item.pdfUrl,children:[txt('PDF',{color:LINK_BLUE,underline:true,size:20})]})])]:[emptyP()],{w:COL[4],borders:STD_BORDERS,vAlign:VerticalAlign.TOP}),
    ]}));

    // Hearing history
    for(const h of (item.hearingHistory||[])){
      agendaRows.push(new TableRow({children:[
        tc([emptyP()],{w:COL[0],borders:STD_BORDERS}),
        tc([p([txt(h.date||'',{bold:true,size:20})]
