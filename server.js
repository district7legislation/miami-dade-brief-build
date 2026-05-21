const express = require('express');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType,
  LevelFormat, ExternalHyperlink, PageNumber, Footer, VerticalAlign
} = require('docx');

const app = express();
app.use(express.json({ limit: '10mb' }));

const COL = [696, 1424, 1560, 4125, 1545];
const TABLE_W = COL.reduce((a, b) => a + b, 0);

const BLUE      = '0070C0';
const PURPLE    = '7030A0';
const LINK_BLUE = '0000FF';
const BLACK     = '000000';
const RED_FAIL  = 'C00000';

function isRegalado(item) { return /regalado/i.test(item.sponsors || ''); }
function accent(item) { return isRegalado(item) ? PURPLE : BLUE; }

function bdr(sz) {
  return { style: BorderStyle.SINGLE, size: sz || 4, color: 'auto' };
}
function noBdr() {
  return { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
}

const STD_BORDERS = { top: bdr(), bottom: bdr(), left: bdr(), right: bdr() };
const NO_BORDERS  = { top: noBdr(), bottom: noBdr(), left: noBdr(), right: noBdr() };
const STD_MARGINS   = { top: 0,  bottom: 0,  left: 108, right: 108 };
const BRIEF_MARGINS = { top: 60, bottom: 60, left: 108, right: 108 };

function txt(text, opts) {
  opts = opts || {};
  return new TextRun({
    text:      text,
    font:      'Arial',
    size:      opts.size    || 20,
    bold:      opts.bold    || false,
    italics:   opts.italics || false,
    color:     opts.color   || BLACK,
    underline: opts.underline ? {} : undefined
  });
}

function p(children, opts) {
  opts = opts || {};
  var runs = Array.isArray(children) ? children
    : typeof children === 'string'   ? [txt(children, opts)]
    : [children];
  return new Paragraph({
    alignment: opts.align  || AlignmentType.LEFT,
    spacing:   { before: opts.before || 0, after: opts.after || 0 },
    numbering: opts.bullet ? { reference: 'bullets', level: opts.level || 0 } : undefined,
    children:  runs
  });
}

function emptyP() { return p(''); }

function inlineRuns(text, baseOpts) {
  baseOpts = baseOpts || {};
  return text.split(/(\*\*[^*]+\*\*)/g).filter(Boolean).map(function(part) {
    if (/^\*\*[^*]+\*\*$/.test(part)) {
      var o = Object.assign({}, baseOpts, { bold: true });
      return txt(part.replace(/\*\*/g, ''), o);
    }
    return txt(part, baseOpts);
  });
}

function parseBrief(briefText, color) {
  if (!briefText) {
    return [p([txt('Brief not available.', { color: RED_FAIL, italics: true, size: 20 })])];
  }
  var paras = [];
  var lines = briefText.split('\n');
  for (var i = 0; i < lines.length; i++) {
    var line    = lines[i].replace(/\s+$/, '');
    var trimmed = line.trim();

    if (!trimmed) {
      paras.push(p('', { after: 60 }));
      continue;
    }
    if (/^⭐/.test(trimmed)) {
      paras.push(p(inlineRuns(trimmed, { bold: true, color: color, size: 20 }), { before: 60, after: 40 }));
      continue;
    }
    if (/^## /.test(line)) {
      paras.push(p([txt(line.replace(/^## /, '').trim(), { bold: true, color: color, size: 20 })], { before: 120, after: 40 }));
      continue;
    }
    if (/^### /.test(line)) {
      paras.push(p([txt(line.replace(/^### /, '').trim(), { bold: true, color: color, size: 20 })], { before: 80, after: 20 }));
      continue;
    }
    if (/^[\-\u2022\*] /.test(trimmed)) {
      paras.push(p(inlineRuns(trimmed.replace(/^[\-\u2022\*] /, ''), { size: 20, color: color }), { bullet: true, level: 0, before: 20, after: 20 }));
      continue;
    }
    if (/^\s{2,}[\-\u2022]/.test(line)) {
      paras.push(p(inlineRuns(trimmed.replace(/^[\-\u2022] /, ''), { size: 20, color: color }), { bullet: true, level: 1, before: 10, after: 10 }));
      continue;
    }
    paras.push(p(inlineRuns(trimmed, { size: 20, color: color }), { before: 20, after: 20 }));
  }
  return paras;
}

function tc(children, opts) {
  opts = opts || {};
  return new TableCell({
    width:         opts.w !== undefined ? { size: opts.w, type: WidthType.DXA } : undefined,
    columnSpan:    opts.span,
    verticalAlign: opts.vAlign  || VerticalAlign.TOP,
    shading:       opts.fill    ? { fill: opts.fill, type: ShadingType.CLEAR } : undefined,
    margins:       opts.margins || STD_MARGINS,
    borders:       opts.borders || STD_BORDERS,
    children:      Array.isArray(children) ? children : [children]
  });
}

async function buildDoc(data) {
  var meetingTitle    = data.meetingTitle    || 'Miami-Dade Board of County Commissioners';
  var meetingVersion  = data.meetingVersion  || 'PRELIMINARY';
  var meetingDate     = data.meetingDate     || '';
  var meetingTime     = data.meetingTime     || '';
  var meetingLocation = data.meetingLocation || 'Commission Chambers';
  var items           = data.items           || [];

  var children = [];

  // Cover header
  var titleParas = [
    p([txt(meetingTitle, { bold: true, color: BLUE, size: 26 })], { align: AlignmentType.CENTER, after: 60 }),
    p([txt(meetingVersion + ' Version', { bold: true, color: BLUE, size: 20, italics: true })], { align: AlignmentType.CENTER, after: 80 }),
    p([txt(meetingDate, { bold: true, color: BLUE, size: 24 })], { align: AlignmentType.CENTER, after: 40 }),
    p([txt(meetingTime, { color: BLUE, size: 20 })], { align: AlignmentType.CENTER, after: 40 }),
    p([txt(meetingLocation, { color: BLUE, size: 20 })], { align: AlignmentType.CENTER, after: 0 })
  ];

  children.push(new Table({
    width:       { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    borders:     { top: noBdr(), bottom: noBdr(), left: noBdr(), right: noBdr(), insideH: noBdr(), insideV: noBdr() },
    rows: [new TableRow({ children: [
      tc(titleParas, { w: 9360, borders: NO_BORDERS, margins: { top: 0, bottom: 0, left: 0, right: 0 } })
    ]})]
  }));

  children.push(emptyP());
  children.push(emptyP());

  // Agenda table
  var agendaRows  = [];
  var lastSection = '';

  for (var idx = 0; idx < items.length; idx++) {
    var item      = items[idx];
    var color     = accent(item);
    var isFailed  = item.status === 'failed';
    var isSection = !item.matterId;

    // Section divider row
    if (item.sectionLabel && item.sectionLabel !== lastSection) {
      lastSection = item.sectionLabel;
      var parts = item.sectionLabel.match(/^(\S+)\s+(.+)$/) || ['', '', item.sectionLabel];
      agendaRows.push(new TableRow({ children: [
        tc([p([txt(parts[1], { bold: true, size: 20 })])], { w: COL[0], borders: STD_BORDERS }),
        tc([p([txt(parts[2], { bold: true, size: 20 })])], { w: COL[1] + COL[2] + COL[3], span: 3, borders: STD_BORDERS }),
        tc([emptyP()], { w: COL[4], borders: STD_BORDERS })
      ]}));
    }

    if (isSection) continue;

    // Metadata row
    agendaRows.push(new TableRow({ children: [
      tc([p([txt(item.itemNumber || '', { bold: true, size: 20 })])],
         { w: COL[0], borders: STD_BORDERS, vAlign: VerticalAlign.CENTER }),
      tc([p([new ExternalHyperlink({ link: item.matterUrl || '#', children: [
               txt((item.matterId || '') + '  ', { color: LINK_BLUE, underline: true, size: 20 })
             ]})])],
         { w: COL[1], borders: STD_BORDERS, vAlign: VerticalAlign.CENTER }),
      tc([p([txt(item.fileType || '', { bold: true, size: 20 })])],
         { w: COL[2], borders: STD_BORDERS, vAlign: VerticalAlign.CENTER }),
      tc([p([txt(item.sponsors || '', { bold: true, size: 20 })])],
         { w: COL[3], borders: STD_BORDERS, vAlign: VerticalAlign.CENTER }),
      tc([emptyP()], { w: COL[4], borders: STD_BORDERS })
    ]}));

    // Title + brief content row
    var contentParas = [];
    contentParas.push(p([txt(item.title || '', { color: BLACK, size: 20 })], { after: 80 }));

    if (isFailed) {
      contentParas.push(emptyP());
      contentParas.push(p([
        txt('⚠ Document could not be retrieved after 3 attempts. ', { bold: true, color: RED_FAIL, size: 20 }),
        txt('Brief not available — edit this document to add manually.', { italics: true, color: RED_FAIL, size: 20 })
      ], { before: 60, after: 40 }));
      contentParas.push(p([
        txt('Source URL: ', { bold: true, size: 20, color: color }),
        new ExternalHyperlink({ link: item.matterUrl || '#', children: [
          txt(item.matterUrl || '', { color: LINK_BLUE, underline: true, size: 20 })
        ]})
      ], { after: 60 }));
    } else {
      var briefParas = parseBrief(item.brief, color);
      for (var bp = 0; bp < briefParas.length; bp++) {
        contentParas.push(briefParas[bp]);
      }
    }

    var pdfChildren = item.pdfUrl
      ? [p([new ExternalHyperlink({ link: item.pdfUrl, children: [
               txt('PDF', { color: LINK_BLUE, underline: true, size: 20 })
             ]})])]
      : [emptyP()];

    agendaRows.push(new TableRow({ children: [
      tc([emptyP()], { w: COL[0], borders: STD_BORDERS }),
      tc(contentParas, { w: COL[1] + COL[2] + COL[3], span: 3, borders: STD_BORDERS, margins: BRIEF_MARGINS }),
      tc(pdfChildren, { w: COL[4], borders: STD_BORDERS, vAlign: VerticalAlign.TOP })
    ]}));

    // Hearing history rows
    var history = item.hearingHistory || [];
    for (var hi = 0; hi < history.length; hi++) {
      var h = history[hi];
      var actionText = (h.action || '') + (h.detail ? ' ' + h.detail : '');
      agendaRows.push(new TableRow({ children: [
        tc([emptyP()], { w: COL[0], borders: STD_BORDERS }),
        tc([p([txt(h.date || '', { bold: true, size: 20 })])],
           { w: COL[1], borders: STD_BORDERS }),
        tc([p([txt(actionText, { italics: true, size: 20 })])],
           { w: COL[2] + COL[3], span: 2, borders: STD_BORDERS }),
        tc([emptyP()], { w: COL[4], borders: STD_BORDERS })
      ]}));
    }
  }

  if (agendaRows.length > 0) {
    children.push(new Table({
      width:        { size: TABLE_W, type: WidthType.DXA },
      columnWidths: COL,
      borders: { top: bdr(4), bottom: bdr(4), left: bdr(4), right: bdr(4), insideH: bdr(4), insideV: bdr(4) },
      rows: agendaRows
    }));
  }

  var doc = new Document({
    numbering: { config: [{ reference: 'bullets', levels: [
      { level: 0, format: LevelFormat.BULLET, text: '\u2022', alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 480, hanging: 240 } } } },
      { level: 1, format: LevelFormat.BULLET, text: '\u25E6', alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 240 } } } }
    ]}]},
    styles: { default: { document: { run: { font: 'Arial', size: 20 } } } },
    sections: [{
      properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
      footers: { default: new Footer({ children: [
        p([
          txt(meetingTitle + '  \u00B7  ' + meetingDate + '  \u00B7  Page ', { size: 16, color: '888888' }),
          new TextRun({ children: [PageNumber.CURRENT], font: 'Arial', size: 16, color: '888888' }),
          txt(' of ', { size: 16, color: '888888' }),
          new TextRun({ children: [PageNumber.TOTAL_PAGES], font: 'Arial', size: 16, color: '888888' })
        ], { align: AlignmentType.CENTER })
      ]}) },
      children: children
    }]
  });

  return Packer.toBuffer(doc);
}

app.post('/build', async function(req, res) {
  try {
    var buffer   = await buildDoc(req.body);
    var filename = 'agenda-brief-' + (req.body.meetingDate || '').replace(/[^a-zA-Z0-9]/g, '_') + '.docx';
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', 'attachment; filename="' + filename + '"');
    res.send(buffer);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

app.get('/health', function(req, res) {
  res.json({ status: 'ok' });
});

var PORT = process.env.PORT || 3000;
app.listen(PORT, function() {
  console.log('Brief builder running on port ' + PORT);
});
