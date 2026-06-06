// Lew's Tracker v8.0 — Task Report to Word (.docx)
// Usage:  node export-to-docx.js <backup.json>
// Output: lews-tracker-tasks-YYYY-MM-DD.docx in current folder
// docx skill: professional formatting, Arial, color-coded priority rows

const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, WidthType, ShadingType, BorderStyle,
  HeadingLevel, PageNumber, PageBreak } = require('docx');
const fs = require('fs');
const path = require('path');

const arg = process.argv[2];
if (!arg) { console.error('Usage: node export-to-docx.js <backup.json>'); process.exit(1); }
const raw = JSON.parse(fs.readFileSync(arg, 'utf8'));
const tasks = (raw.tasks || []).map(t => {
  t.timeLog = t.timeLog || [];
  t.tags = t.tags || [];
  t.checklist = t.checklist || [];
  t.recurring = t.recurring || 'none';
  t.status = t.status || 'todo';
  t.priority = t.priority || 'medium';
  return t;
});
const notes = raw.notes || [];

const GOLD = 'D8B76D';
const GOLD_DARK = '8B6914';
const PRI_BG  = { high: 'FFEBEB', medium: 'FFFBEB', low: 'EBFFEB' };
const PRI_TXT = { high: 'C00000', medium: '8B6914', low: '006400' };

const todayStr = () => new Date().toISOString().slice(0,10);
const fmtSecs = s => {
  if (!s) return null;
  if (s < 60) return s + 's';
  if (s < 3600) return Math.floor(s/60) + 'm';
  return Math.floor(s/3600) + 'h ' + Math.floor((s%3600)/60) + 'm';
};
const totalSecs = t => t.timeLog.reduce((a,b) => a + (b.seconds||0), 0);
const statusName = s => ({todo:'To Do',progress:'In Progress',waiting:'Waiting',done:'Done'}[s] || s);

const bdr = { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' };
const borders = { top: bdr, bottom: bdr, left: bdr, right: bdr };

function hdrCell(text, w) {
  return new TableCell({
    borders, width: { size: w, type: WidthType.DXA },
    shading: { fill: GOLD, type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({ children: [new TextRun({ text, bold: true, font: 'Arial', size: 18, color: '0A0908' })] })]
  });
}

function cell(text, w, opts) {
  opts = opts || {};
  return new TableCell({
    borders, width: { size: w, type: WidthType.DXA },
    shading: opts.fill ? { fill: opts.fill, type: ShadingType.CLEAR } : undefined,
    margins: { top: 60, bottom: 60, left: 120, right: 120 },
    children: [new Paragraph({ children: [new TextRun({
      text: String(text != null ? text : '—'),
      font: 'Arial', size: 18,
      color: opts.color || '333333',
      bold: opts.bold || false,
      strike: opts.strike || false,
    })] })]
  });
}

const openCnt = tasks.filter(t => t.status !== 'done').length;
const doneCnt = tasks.filter(t => t.status === 'done').length;
const overdue = tasks.filter(t => t.status !== 'done' && t.due && t.due < todayStr()).length;
const timeTotal = fmtSecs(tasks.reduce((a,t) => a + totalSecs(t), 0));

// Task table rows
const taskRows = [
  new TableRow({ children: [
    hdrCell('Title', 3500), hdrCell('Priority', 900), hdrCell('Status', 1100),
    hdrCell('Due', 1000), hdrCell('Time', 800), hdrCell('Tags / Recurring', 2060)
  ]})
].concat(tasks.map(t => {
  const fill = t.status === 'done' ? 'F5F5F5' : (PRI_BG[t.priority] || 'FFFFFF');
  const secs = totalSecs(t);
  const tags = t.tags.map(g => '#' + g).join(', ');
  const rec = t.recurring && t.recurring !== 'none' ? ' ↺' + t.recurring : '';
  return new TableRow({ children: [
    cell(t.title, 3500, { fill, strike: t.status === 'done' }),
    cell(t.priority, 900, { fill, color: PRI_TXT[t.priority] || '333333', bold: true }),
    cell(statusName(t.status), 1100, { fill }),
    cell(t.due || '—', 1000, { fill }),
    cell(fmtSecs(secs) || '—', 800, { fill, color: secs > 0 ? '006400' : '888888' }),
    cell((tags || '—') + rec, 2060, { fill }),
  ]});
}));

// Notes table rows
const noteRows = notes.length ? [
  new TableRow({ children: [hdrCell('Title', 2400), hdrCell('Date', 1400), hdrCell('Content', 5560)] })
].concat(notes.map(n => new TableRow({ children: [
  cell(n.title, 2400, {}),
  cell(n.created || '', 1400, {}),
  cell((n.body || '').substring(0, 280), 5560, {}),
]}))) : [];

const doc = new Document({
  styles: {
    default: { document: { run: { font: 'Arial', size: 22 } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 32, bold: true, font: 'Arial', color: GOLD_DARK },
        paragraph: { spacing: { before: 200, after: 140 }, outlineLevel: 0 } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 26, bold: true, font: 'Arial', color: GOLD_DARK },
        paragraph: { spacing: { before: 180, after: 100 }, outlineLevel: 1 } },
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 }
      }
    },
    headers: { default: new Header({ children: [new Paragraph({
      border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: GOLD, space: 1 } },
      children: [
        new TextRun({ text: "Lew's Tracker v8.0  ·  Task Report", font: 'Arial', size: 18, color: '888888' }),
        new TextRun({ text: '\t' + new Date().toLocaleDateString(), font: 'Arial', size: 18, color: '888888' }),
      ],
      tabStops: [{ type: 'right', position: 10080 }],
    })]}) },
    footers: { default: new Footer({ children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [
        new TextRun({ text: 'Page ', font: 'Arial', size: 16, color: 'AAAAAA' }),
        new TextRun({ children: [PageNumber.CURRENT], font: 'Arial', size: 16, color: 'AAAAAA' }),
        new TextRun({ text: ' of ', font: 'Arial', size: 16, color: 'AAAAAA' }),
        new TextRun({ children: [PageNumber.TOTAL_PAGES], font: 'Arial', size: 16, color: 'AAAAAA' }),
        new TextRun({ text: '   ·   For Melliea ♥', font: 'Arial', size: 16, color: 'CCAA66' }),
      ]
    })]}) },
    children: [
      new Paragraph({ heading: HeadingLevel.HEADING_1,
        children: [new TextRun("Lew’s Tracker — Task Report")] }),
      new Paragraph({ children: [new TextRun({
        text: 'Generated: ' + new Date().toLocaleString(),
        font: 'Arial', size: 18, color: '888888'
      })] }),
      new Paragraph({ children: [] }),

      // Summary bar
      new Paragraph({
        border: { left: { style: BorderStyle.SINGLE, size: 16, color: GOLD, space: 8 } },
        shading: { fill: 'FDF8EE', type: ShadingType.CLEAR },
        spacing: { before: 60, after: 60 },
        indent: { left: 160 },
        children: [new TextRun({
          text: tasks.length + ' total  ·  ' + openCnt + ' open  ·  ' + doneCnt + ' done  ·  ' + overdue + ' overdue' + (timeTotal ? '  ·  Time: ' + timeTotal : ''),
          font: 'Arial', size: 20, bold: true
        })]
      }),
      new Paragraph({ children: [] }),

      new Paragraph({ heading: HeadingLevel.HEADING_2,
        children: [new TextRun('To Do List (' + tasks.length + ')')] }),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [3500, 900, 1100, 1000, 800, 2060],
        rows: taskRows
      }),

      ...(noteRows.length ? [
        new Paragraph({ children: [new PageBreak()] }),
        new Paragraph({ heading: HeadingLevel.HEADING_2,
          children: [new TextRun('Notes (' + notes.length + ')')] }),
        new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [2400, 1400, 5560],
          rows: noteRows
        }),
      ] : []),
    ]
  }]
});

const outFile = 'lews-tracker-tasks-' + todayStr() + '.docx';
Packer.toBuffer(doc).then(function(buf) {
  fs.writeFileSync(outFile, buf);
  console.log('✓ Word report saved: ' + outFile);
  console.log('  Tasks: ' + tasks.length + ' | Notes: ' + notes.length);
}).catch(function(e) {
  console.error('Error:', e.message);
  process.exit(1);
});
