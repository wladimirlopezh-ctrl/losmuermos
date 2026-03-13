const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign
} = require('docx');
const fs = require('fs');

const data = JSON.parse(fs.readFileSync('datos.json'));
const db = data.db;

const HEADER_BG = 'D9D9D9';
const BS = { style: BorderStyle.SINGLE, size: 4, color: '000000' };
const BN = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const borders = { top: BS, bottom: BS, left: BS, right: BS };
const bordersNone = { top: BN, bottom: BN, left: BN, right: BN };

const COL_WIDTHS = [709, 1276, 1067, 1067, 1067, 1068, 1067, 1067, 1067, 1068];
const TABLE_WIDTH = COL_WIDTHS.reduce((a,b)=>a+b,0);

function txt(text, opts={}) {
  return new TextRun({ text: String(text||''), font:'Arial', size:opts.size||16, bold:opts.bold||false, color:opts.color||'000000' });
}

function makeCell(content, opts={}) {
  return new TableCell({
    borders: opts.borders || borders,
    width: { size: opts.width||1067, type: WidthType.DXA },
    shading: opts.fill ? { fill: opts.fill, type: ShadingType.CLEAR } : undefined,
    verticalAlign: VerticalAlign.CENTER,
    columnSpan: opts.span||1,
    rowSpan: opts.rowSpan||1,
    margins: { top:60, bottom:60, left:80, right:80 },
    children: [new Paragraph({ alignment: opts.align||AlignmentType.CENTER, children: Array.isArray(content)?content:[content] })]
  });
}

function getRegistros(n) {
  const t = db[String(n)]||{};
  const hist = (t.historial||[]).filter(h=>h.fecha>='2026-01-01');
  const registros = hist.map(h=>({ enc:h.enc, fechaAsig:h.fechaAsig||h.fecha, fechaComp:h.fecha }));
  if (t.estado==='asignado' && (t.fechaAsig||'')>='2026-01-01') {
    const yaEsta = registros.some(r=>r.enc===t.encargado&&r.fechaAsig===t.fechaAsig);
    if(!yaEsta) registros.push({ enc:t.encargado, fechaAsig:t.fechaAsig, fechaComp:'' });
  }
  return registros;
}

function getUltimaFecha2026(n) {
  const t = db[String(n)]||{};
  const hist = [...(t.historial||[])].filter(h=>h.fecha>='2026-01-01').sort((a,b)=>b.fecha.localeCompare(a.fecha));
  return hist[0] ? fmtDate(hist[0].fecha) : '';
}

function fmtDate(d) {
  if(!d) return '';
  const [y,m,day] = d.split('-');
  return `${day}-${m}-${y.slice(2)}`;
}

function makeHeaderRow1() {
  return new TableRow({ tableHeader:true, children:[
    makeCell(txt('Núm. de terr.',{size:14,color:'404040'}), {width:709, fill:HEADER_BG, rowSpan:2}),
    makeCell(txt('Última fecha en que se completó*',{size:14,color:'404040'}), {width:1276, fill:HEADER_BG, rowSpan:2}),
    ...Array(4).fill(null).map((_,i)=>
      makeCell(txt('Asignado a',{size:14,color:'404040'}), {width:COL_WIDTHS[2+i*2]+COL_WIDTHS[3+i*2], fill:HEADER_BG, span:2})
    )
  ]});
}

function makeHeaderRow2() {
  return new TableRow({ tableHeader:true, children:[
    ...Array(4).fill(null).flatMap((_,i)=>[
      makeCell(txt('Fecha en que se asignó',{size:13,color:'404040'}), {width:COL_WIDTHS[2+i*2], fill:HEADER_BG}),
      makeCell(txt('Fecha en que se completó',{size:13,color:'404040'}), {width:COL_WIDTHS[3+i*2], fill:HEADER_BG}),
    ])
  ]});
}

function makeDataRows(n) {
  const registros = getRegistros(n);
  const ultimaFecha = getUltimaFecha2026(n);
  const slots = Array(4).fill(null).map((_,i)=>registros[i]||null);

  const nameRow = new TableRow({ children:[
    makeCell(txt(String(n),{size:18}), {width:709, rowSpan:2}),
    makeCell(txt(ultimaFecha,{size:18}), {width:1276, rowSpan:2}),
    ...slots.flatMap((r,i)=>[
      makeCell(txt(r?r.enc:'',{size:18}), {width:COL_WIDTHS[2+i*2]+COL_WIDTHS[3+i*2], span:2})
    ])
  ]});

  const dateRow = new TableRow({ children:[
    ...slots.flatMap((r,i)=>[
      makeCell(txt(r?fmtDate(r.fechaAsig):'',{size:18}), {width:COL_WIDTHS[2+i*2]}),
      makeCell(txt(r?fmtDate(r.fechaComp):'',{size:18}), {width:COL_WIDTHS[3+i*2]}),
    ])
  ]});

  return [nameRow, dateRow];
}

const rows = [makeHeaderRow1(), makeHeaderRow2()];
for(let n=1;n<=33;n++) rows.push(...makeDataRows(n));

const doc = new Document({
  sections:[{
    properties:{ page:{ size:{width:11906,height:16838}, margin:{top:1080,right:720,bottom:965,left:720} } },
    children:[
      new Paragraph({ alignment:AlignmentType.CENTER, spacing:{after:160},
        children:[txt('REGISTRO DE ASIGNACIÓN DE TERRITORIO',{size:24,bold:true})] }),
      new Table({
        width:{size:3000,type:WidthType.DXA}, columnWidths:[2000,1000],
        rows:[new TableRow({ children:[
          new TableCell({ borders:bordersNone, width:{size:2000,type:WidthType.DXA},
            children:[new Paragraph({children:[txt('Año de servicio:',{bold:true,size:20})]})] }),
          new TableCell({ borders:{...bordersNone,bottom:BS}, width:{size:1000,type:WidthType.DXA},
            children:[new Paragraph({alignment:AlignmentType.CENTER,children:[txt('2026',{bold:true,size:20})]})] }),
        ]})]
      }),
      new Paragraph({ spacing:{after:160}, children:[] }),
      new Table({ width:{size:TABLE_WIDTH,type:WidthType.DXA}, columnWidths:COL_WIDTHS, rows }),
      new Paragraph({ spacing:{before:120},
        children:[txt('*Cuando comience una nueva página, anote en esta columna la última fecha en que los territorios se completaron.',{size:18})] }),
    ]
  }]
});

Packer.toBuffer(doc).then(buf=>{
  fs.writeFileSync('registro_2026.docx', buf);
  console.log('✅ registro_2026.docx generado correctamente');
});
