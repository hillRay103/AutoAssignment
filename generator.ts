/**
 * AutoAssignment — Final stable generator
 * DOCX: docx npm  |  PDF: pdf-lib npm
 */
import * as fs from 'fs';

const NAVY  = '1A3A6C';
const BLUE  = '1F6FB2';
const BLUE2 = 'EEF1F7';
const DGRAY = '141414';
const MGRAY = '4A4A4A';
const GREEN = '1A6B3A';
const WHITE = 'FFFFFF';

const CODE_BG: Record<string,string> = {
    lightgray:'F7F8FA', lightblue:'EBF5FB', darkgray:'1E1E2E', white:'FFFFFF'
};
const CODE_FG: Record<string,string> = {
    lightgray:DGRAY, lightblue:NAVY, darkgray:'CDD6F4', white:DGRAY
};
const OUT_BG: Record<string,string> = {
    lightgray:'F0FBF4', lightblue:'E8F8EE', darkgray:'0A2010', white:'F0FBF4'
};

// ═══════════════════════════════════════════════════════════════════════════
// DOCX
// ═══════════════════════════════════════════════════════════════════════════
export async function buildDocx(data: any, outPath: string): Promise<void> {
    const {
        Document, Paragraph, TextRun, Table, TableRow, TableCell,
        WidthType, BorderStyle, AlignmentType, ShadingType,
        convertInchesToTwip, Packer, Footer, PageNumber,
        TableLayoutType,
    } = await import('docx');

    const font       = data.fontChoice || 'Calibri';
    const tpl        = data.template   || 'B';
    const name       = data.studentName     || '';
    const reg        = data.regNumber       || '';
    const course     = data.courseName      || '';
    const courseC    = data.courseCode      || '';
    const section    = data.section         || '';
    const instructor = data.submittedTo     || '';
    const dept       = data.department      || '';
    const uni        = data.university      || '';
    const prog       = data.programName     || '';
    const sem        = data.semester        || '';
    const batch      = data.batch           || '';
    const credits    = data.creditHours     || '';
    const assignNo   = data.assignmentNumber || 1;
    const assignTtl  = data.assignmentTitle  || '';
    const subDate    = data.submissionDate   || new Date().toLocaleDateString('en-GB');
    const clo        = data.cloLabel         || '';
    const project    = data.projectTitle     || '';
    const questions  = data.questions        || [];
    const theme      = data.codeTheme        || 'lightgray';
    const codeFs     = parseInt(data.codeFontSize || '10');
    const language   = data.language         || '';
    const fullTitle  = `Assignment No. ${assignNo}` + (assignTtl ? ` - ${assignTtl}` : '');
    const courseStr  = course + (courseC ? ` (${courseC})` : '');

    const children: any[] = [];

    // ── helpers ────────────────────────────────────────────────────────────
    const R = (text: string, o: any = {}) => new TextRun({
        text, font,
        size:    (o.size || 11) * 2,
        bold:    o.bold   || false,
        italics: o.italic || false,
        color:   (o.color || DGRAY).replace('#',''),
    });

    const P = (runs: any[], o: any = {}) => new Paragraph({
        children: Array.isArray(runs) ? runs : [runs],
        alignment: o.align || AlignmentType.LEFT,
        spacing: { before:(o.sb||0)*20, after:(o.sa||3)*20 },
    });

    const noB = () => ({ style: BorderStyle.NONE, size:0, color:'auto' });
    const soB = (c: string, sz = 4) => ({ style: BorderStyle.SINGLE, size:sz, color:c.replace('#','') });

    // Info cell — percentage width
    const iCell = (text: string, o: any = {}) => new TableCell({
        children: [new Paragraph({
            children: [R(text, { bold:o.bold, color:o.color||DGRAY, size:o.size||10 })],
            spacing: { before:55, after:55 },
        })],
        shading: o.bg ? { type:ShadingType.CLEAR, fill:o.bg.replace('#',''), color:'auto' } : undefined,
        width: o.pct != null ? { size:o.pct, type:WidthType.PERCENTAGE } : undefined,
        columnSpan: o.span || 1,
        borders: {
            top:    soB('B0BAC8'),
            bottom: soB('B0BAC8'),
            left:   soB('B0BAC8'),
            right:  soB('B0BAC8'),
        },
        margins: { top:40, bottom:40, left:90, right:90 },
    });

    const infoTable = (rows: any[][]) => new Table({
        width: { size:100, type:WidthType.PERCENTAGE },
        layout: TableLayoutType.FIXED,
        rows: rows.map(cells => new TableRow({ children: cells })),
    });

    const divLine = (color = NAVY, thick = 6) => new Paragraph({
        border: { bottom: { style:BorderStyle.SINGLE, size:thick, color:color.replace('#',''), space:1 } },
        spacing: { before:20, after:20 }, children:[],
    });

    // Question banner
    const qBanner = (label: string, qNum: number) => new Table({
        width: { size:100, type:WidthType.PERCENTAGE },
        layout: TableLayoutType.FIXED,
        rows: [new TableRow({ children: [
            new TableCell({
                children: [P([R(`Q${qNum}`, { bold:true, size:11, color:WHITE })], { align:AlignmentType.CENTER })],
                shading: { type:ShadingType.CLEAR, fill:BLUE.replace('#',''), color:'auto' },
                width: { size:700, type:WidthType.DXA },
                borders: { top:noB(), bottom:noB(), left:noB(), right:noB() },
                margins: { top:60, bottom:60, left:60, right:60 },
            }),
            new TableCell({
                children: [P([R(label, { bold:true, size:11, color:WHITE })], {})],
                shading: { type:ShadingType.CLEAR, fill:NAVY.replace('#',''), color:'auto' },
                width: { size:8660, type:WidthType.DXA },
                borders: { top:noB(), bottom:noB(), left:noB(), right:noB() },
                margins: { top:60, bottom:60, left:120, right:60 },
            }),
        ]})],
    });

    // ── Code block: header row + ONE content row with all lines as paragraphs ──
    const makeCodeBlock = (
        headerLabel: string,
        langTag: string,
        lines: string[],
        headerBg: string,
        contentBg: string,
        contentFg: string,
        borderColor: string,
    ) => {
        return new Table({
            width: { size:100, type:WidthType.PERCENTAGE },
            layout: TableLayoutType.FIXED,
            rows: [
                // Row 1 — header
                new TableRow({ children: [new TableCell({
                    children: [new Paragraph({
                        children: [
                            R(headerLabel, { bold:true, size:10, color:borderColor }),
                            ...(langTag ? [R(`   ${langTag}`, { size:9, color:MGRAY })] : []),
                        ],
                        spacing: { before:80, after:80 },
                    })],
                    shading: { type:ShadingType.CLEAR, fill:headerBg.replace('#',''), color:'auto' },
                    borders: {
                        top:    soB(borderColor, 6),
                        bottom: soB('D0D8E8', 3),
                        left:   soB(borderColor, 6),
                        right:  soB('D0D8E8', 3),
                    },
                    margins: { top:0, bottom:0, left:120, right:120 },
                })]},
                ),
                // Row 2 — ALL code lines as paragraphs in ONE cell
                new TableRow({ children: [new TableCell({
                    children: [
                        ...lines.map(line => new Paragraph({
                            children: [new TextRun({
                                text: line.length === 0 ? '\u00A0' : line,
                                font: 'Courier New',
                                size: Math.max(codeFs, 8) * 2,
                                color: contentFg.replace('#',''),
                            })],
                            spacing: { before:0, after:0 },
                        })),
                        // bottom padding
                        new Paragraph({ children:[], spacing:{ before:100, after:0 } }),
                    ],
                    shading: { type:ShadingType.CLEAR, fill:contentBg.replace('#',''), color:'auto' },
                    borders: {
                        top:    noB(),
                        bottom: soB(borderColor, 6),
                        left:   soB(borderColor, 6),
                        right:  soB('D0D8E8', 3),
                    },
                    margins: { top:80, bottom:0, left:120, right:120 },
                })]},
                ),
            ],
        });
    };

    const codeSection = (label: string, qNum: number, code: string, output: string, qtext: string) => {
        const cbg = CODE_BG[theme] || 'F7F8FA';
        const cfg = CODE_FG[theme] || DGRAY;
        const obg = OUT_BG[theme]  || 'F0FBF4';

        children.push(P([], { sa:6 }));
        children.push(qBanner(label, qNum));
        if (qtext) {
            children.push(P([R(qtext, { italic:true, size:10, color:MGRAY })], { sb:2, sa:1 }));
        }
        children.push(P([], { sa:2 }));
        children.push(makeCodeBlock('Source Code', language, (code||'').split('\n'), BLUE2, cbg, cfg, BLUE));
        children.push(P([], { sa:4 }));
        if (output && output.trim()) {
            children.push(makeCodeBlock('Output', '', output.split('\n'), 'E8F8EE', obg, DGRAY, GREEN));
            children.push(P([], { sa:4 }));
        }
    };

    // ══ TEMPLATE A ═══════════════════════════════════════════════════════════
    if (tpl === 'A') {
        if (uni)  children.push(P([R(uni,  { bold:true, size:15, color:NAVY })], { align:AlignmentType.CENTER, sa:1 }));
        if (dept) children.push(P([R(dept, { size:11,   color:NAVY })],          { align:AlignmentType.CENTER, sa:1 }));
        children.push(divLine(NAVY, 8));
        children.push(divLine(NAVY, 3));
        children.push(P([], { sa:4 }));
        children.push(infoTable([
            [ iCell('Course Title:', {bold:true,color:NAVY,bg:BLUE2,pct:16}), iCell(courseStr,      {pct:22,bg:WHITE}),
              iCell('Course Code:',  {bold:true,color:NAVY,bg:BLUE2,pct:13}), iCell(courseC,        {pct:10,bg:WHITE}),
              iCell('Credit Hours:', {bold:true,color:NAVY,bg:BLUE2,pct:14}), iCell(credits,        {pct:25,bg:WHITE}) ],
            [ iCell('Instructor:',   {bold:true,color:NAVY,bg:BLUE2,pct:16}), iCell(instructor,     {pct:22,bg:WHITE}),
              iCell('Program:',      {bold:true,color:NAVY,bg:BLUE2,pct:13}), iCell(prog||'BS(CS)', {pct:49,bg:WHITE,span:3}) ],
            [ iCell('Semester:',     {bold:true,color:NAVY,bg:BLUE2,pct:16}), iCell(sem,            {pct:10,bg:WHITE}),
              iCell('Batch:',        {bold:true,color:NAVY,bg:BLUE2,pct:10}), iCell(batch,          {pct:12,bg:WHITE}),
              iCell('Section:',      {bold:true,color:NAVY,bg:BLUE2,pct:13}), iCell(section,        {pct:39,bg:WHITE}) ],
            [ iCell('Date:',         {bold:true,color:NAVY,bg:BLUE2,pct:16}), iCell(subDate,        {pct:84,bg:WHITE,span:5}) ],
            [ iCell('Name:',         {bold:true,color:NAVY,bg:BLUE2,pct:16}), iCell(name,           {pct:34,bg:WHITE,span:2}),
              iCell('Reg No:',       {bold:true,color:NAVY,bg:BLUE2,pct:13}), iCell(reg,            {pct:37,bg:WHITE,span:2}) ],
        ]));
        children.push(P([], { sa:6 }));

    // ══ TEMPLATE B ═══════════════════════════════════════════════════════════
    } else if (tpl === 'B') {
        children.push(P([R(fullTitle, { bold:true, size:16, color:NAVY })], { align:AlignmentType.CENTER, sa:2 }));
        if (assignTtl) children.push(P([R(assignTtl, { size:11, color:MGRAY, italic:true })], { align:AlignmentType.CENTER, sa:4 }));
        else children.push(P([], { sa:3 }));
        children.push(infoTable([
            [ iCell('Student Name:', {bold:true,color:NAVY,bg:BLUE2,pct:22}), iCell(name,       {pct:28,bg:WHITE}),
              iCell('Reg No.:',      {bold:true,color:NAVY,bg:BLUE2,pct:22}), iCell(reg,        {pct:28,bg:WHITE}) ],
            [ iCell('Course:',       {bold:true,color:NAVY,bg:BLUE2,pct:22}), iCell(courseStr,  {pct:28,bg:WHITE}),
              iCell('Section:',      {bold:true,color:NAVY,bg:BLUE2,pct:22}), iCell(section,    {pct:28,bg:WHITE}) ],
            [ iCell('Instructor:',   {bold:true,color:NAVY,bg:BLUE2,pct:22}), iCell(instructor, {pct:28,bg:WHITE}),
              iCell('Date:',         {bold:true,color:NAVY,bg:BLUE2,pct:22}), iCell(subDate,    {pct:28,bg:WHITE}) ],
            ...(sem ? [[ iCell('Semester:', {bold:true,color:NAVY,bg:BLUE2,pct:22}), iCell(sem,  {pct:28,bg:WHITE}),
                         iCell('Batch:',   {bold:true,color:NAVY,bg:BLUE2,pct:22}), iCell(batch,{pct:28,bg:WHITE}) ]] : []),
        ]));
        children.push(P([], { sa:6 }));

    // ══ TEMPLATE C ═══════════════════════════════════════════════════════════
    } else if (tpl === 'C') {
        const parts = [courseStr, section?`Section ${section}`:'', instructor].filter(Boolean);
        children.push(P([R(fullTitle, { bold:true, size:16, color:NAVY })], { align:AlignmentType.CENTER, sb:2, sa:6 }));
        children.push(P([R(`${name}  |  ${reg}`, { bold:true, size:11, color:DGRAY })], { align:AlignmentType.CENTER, sa:1 }));
        children.push(P([R(parts.join('  |  '),  { size:10, color:MGRAY })],            { align:AlignmentType.CENTER, sa:1 }));
        children.push(P([R(`Date: ${subDate}`,   { size:10, color:MGRAY })],            { align:AlignmentType.CENTER, sa:4 }));
        children.push(divLine(BLUE, 4));
        children.push(P([], { sa:4 }));

    // ══ TEMPLATE D ═══════════════════════════════════════════════════════════
    } else if (tpl === 'D') {
        children.push(P([], { sa:60 }));
        if (uni)  children.push(P([R(uni,  { bold:true, size:18, color:NAVY })], { align:AlignmentType.CENTER, sa:1 }));
        if (dept) children.push(P([R(dept, { size:11,   color:MGRAY })],         { align:AlignmentType.CENTER, sa:1 }));
        children.push(divLine(NAVY, 6));
        children.push(P([], { sa:40 }));

        if (clo) {
            children.push(new Table({
                width: { size:100, type:WidthType.PERCENTAGE },
                layout: TableLayoutType.FIXED,
                rows: [new TableRow({ children: [
                    new TableCell({
                        children: [P([R(clo, { bold:true, size:10, color:WHITE })], { align:AlignmentType.CENTER, sb:1, sa:1 })],
                        shading: { type:ShadingType.CLEAR, fill:'1A6B8A', color:'auto' },
                        borders: { top:noB(), bottom:noB(), left:noB(), right:noB() },
                    }),
                ]})],
            }));
            children.push(P([], { sa:8 }));
        }
        children.push(P([R(fullTitle, { bold:true, size:20, color:NAVY })], { align:AlignmentType.CENTER, sa:2 }));
        if (assignTtl) children.push(P([R(assignTtl, { size:12, color:MGRAY, italic:true })], { align:AlignmentType.CENTER, sa:2 }));
        if (project) {
            children.push(P([], { sa:8 }));
            children.push(P([R('Project / Topic:', { bold:true, size:11, color:NAVY })], { align:AlignmentType.CENTER, sa:1 }));
            children.push(P([R(project, { bold:true, italic:true, size:13, color:NAVY })], { align:AlignmentType.CENTER, sa:2 }));
        }
        children.push(P([], { sa:40 }));

        children.push(infoTable([
            [ iCell('Submitted By:', {bold:true,color:NAVY,bg:BLUE2,pct:30}), iCell(`${name} (${reg})`, {pct:70,bg:WHITE}) ],
            [ iCell('Submitted To:', {bold:true,color:NAVY,bg:BLUE2,pct:30}), iCell(instructor,         {pct:70,bg:WHITE}) ],
            [ iCell('Course:',       {bold:true,color:NAVY,bg:BLUE2,pct:30}), iCell(courseStr,          {pct:70,bg:WHITE}) ],
            ...(section ? [[ iCell('Section:', {bold:true,color:NAVY,bg:BLUE2,pct:30}), iCell(section, {pct:70,bg:WHITE}) ]] : []),
            ...(prog    ? [[ iCell('Program:', {bold:true,color:NAVY,bg:BLUE2,pct:30}), iCell(prog,    {pct:70,bg:WHITE}) ]] : []),
            ...(sem     ? [[ iCell('Semester:',{bold:true,color:NAVY,bg:BLUE2,pct:30}), iCell(sem,     {pct:70,bg:WHITE}) ]] : []),
            [ iCell('Date:',         {bold:true,color:NAVY,bg:BLUE2,pct:30}), iCell(subDate,            {pct:70,bg:WHITE}) ],
        ]));
        children.push(new Paragraph({ children:[], pageBreakBefore:true }));
    }

    // Questions
    questions.forEach((q: any, idx: number) => {
        codeSection(q.label||`Question ${idx+1}`, idx+1, q.code||'', q.output||'', q.text||'');
        if (q.notes) children.push(P([R(`Note: ${q.notes}`,{italic:true,size:10,color:MGRAY})],{sa:4}));
    });

    // Footer
    const doc = new Document({ sections:[{
        properties:{ page:{
            size:{ width:convertInchesToTwip(8.5), height:convertInchesToTwip(11) },
            margin:{ top:convertInchesToTwip(1), bottom:convertInchesToTwip(1), left:convertInchesToTwip(1), right:convertInchesToTwip(1) },
        }},
        footers:{ default: new Footer({ children:[new Paragraph({
            alignment: AlignmentType.CENTER,
            border:{ top:{ style:BorderStyle.SINGLE, size:4, color:'C8D4E8', space:4 } },
            spacing:{ before:60 },
            children:[
                new TextRun({ text:`${name}  |  ${fullTitle}  |  Page `, font, size:16, color:MGRAY }),
                new TextRun({ children:[PageNumber.CURRENT], font, size:16, color:MGRAY }),
                new TextRun({ text:' of ', font, size:16, color:MGRAY }),
                new TextRun({ children:[PageNumber.TOTAL_PAGES], font, size:16, color:MGRAY }),
            ],
        })]})},
        children,
    }]});
    fs.writeFileSync(outPath, await Packer.toBuffer(doc));
}


// ═══════════════════════════════════════════════════════════════════════════
// PDF — pdf-lib canvas, no cursor issues
// ═══════════════════════════════════════════════════════════════════════════
export async function buildPdf(data: any, outPath: string): Promise<void> {
    const { PDFDocument, rgb, StandardFonts } = await import('pdf-lib');

    const useTimes  = (data.fontChoice||'').includes('Times');
    const tpl       = data.template         || 'B';
    const name      = data.studentName      || '';
    const reg       = data.regNumber        || '';
    const course    = data.courseName       || '';
    const courseC   = data.courseCode       || '';
    const section   = data.section          || '';
    const instructor= data.submittedTo      || '';
    const dept      = data.department       || '';
    const uni       = data.university       || '';
    const prog      = data.programName      || '';
    const sem       = data.semester         || '';
    const batch     = data.batch            || '';
    const credits   = data.creditHours      || '';
    const assignNo  = data.assignmentNumber || 1;
    const assignTtl = data.assignmentTitle  || '';
    const subDate   = data.submissionDate   || new Date().toLocaleDateString('en-GB');
    const clo       = data.cloLabel         || '';
    const project   = data.projectTitle     || '';
    const questions = data.questions        || [];
    const theme     = data.codeTheme        || 'lightgray';
    const codeFs    = parseInt(data.codeFontSize || '9');
    const language  = data.language         || '';

    const fullTitle = `Assignment No. ${assignNo}` + (assignTtl ? ` - ${assignTtl}` : '');
    const courseStr = course + (courseC ? ` (${courseC})` : '');

    const PW=612, PH=792, LM=54, TM=60, BM=60;
    const CW = PW-LM*2;
    const SAFE_BOTTOM = BM + 50; // minimum Y before forcing page break

    const pdfDoc = await PDFDocument.create();
    const fN  = await pdfDoc.embedFont(useTimes ? StandardFonts.TimesRoman      : StandardFonts.Helvetica);
    const fB  = await pdfDoc.embedFont(useTimes ? StandardFonts.TimesRomanBold   : StandardFonts.HelveticaBold);
    const fI  = await pdfDoc.embedFont(useTimes ? StandardFonts.TimesRomanItalic : StandardFonts.HelveticaOblique);
    const fM  = await pdfDoc.embedFont(StandardFonts.Courier);
    const fMB = await pdfDoc.embedFont(StandardFonts.CourierBold);

    const C = (h: string) => {
        const s = h.replace('#','');
        return rgb(parseInt(s.slice(0,2),16)/255, parseInt(s.slice(2,4),16)/255, parseInt(s.slice(4,6),16)/255);
    };

    let pg: any;
    let y = 0;
    const newPage = () => { pg = pdfDoc.addPage([PW,PH]); y = PH - TM; };
    newPage();

    // ── primitives ────────────────────────────────────────────────────────
    const fillR = (x:number,py:number,w:number,h:number,col:string) =>
        pg.drawRectangle({ x, y:py-h, width:w, height:h, color:C(col), borderWidth:0 });

    const strokeR = (x:number,py:number,w:number,h:number,col:string,lw=0.7) =>
        pg.drawRectangle({ x, y:py-h, width:w, height:h, borderColor:C(col), borderWidth:lw, color:undefined });

    const hLine = (py:number,col:string,lw=1) =>
        pg.drawLine({ start:{x:LM,y:py}, end:{x:LM+CW,y:py}, thickness:lw, color:C(col) });

    // Text at exact position — pdf-lib never auto-moves cursor
    const dt = (text:string, x:number, py:number, opts:any={}) => {
        const f = opts.mono ? fM : opts.bold ? fB : opts.italic ? fI : fN;
        pg.drawText(text||' ', { x, y:py, font:f, size:opts.size||10, color:C(opts.color||DGRAY) });
    };

    // Single text line returning next Y
    const tl = (text:string, x:number, py:number, opts:any={}): number => {
        const f   = opts.bold ? fB : opts.italic ? fI : fN;
        const sz  = opts.size || 10;
        const maxW= opts.width ?? (LM+CW-x);
        let t = text;
        while (t.length>1 && f.widthOfTextAtSize(t,sz)>maxW) t=t.slice(0,-1);
        let tx = x;
        if (opts.align==='center') tx = x+(maxW-f.widthOfTextAtSize(t,sz))/2;
        if (opts.align==='right')  tx = x+maxW-f.widthOfTextAtSize(t,sz);
        dt(t, tx, py, opts);
        return py - sz - (opts.gap||4);
    };

    // Table row returning next Y
    const tr = (cols:{t:string,w:number,bold?:boolean,bg?:string,color?:string}[], py:number, rh=18): number => {
        let x=LM;
        for (const c of cols) {
            if (!c.w||c.w<=0) continue;
            if (c.bg) fillR(x,py,c.w,rh,c.bg);
            strokeR(x,py,c.w,rh,'B0BAC8',0.5);
            dt(c.t||'', x+5, py-rh+5, { bold:c.bold, color:c.color||DGRAY, size:8.5 });
            x+=c.w;
        }
        return py-rh;
    };

    const need = (h:number) => { if (y-h < SAFE_BOTTOM) newPage(); };

    // ── PDF code block — draws line by line, handles page breaks cleanly ──
    const pdfCodeBlock = (label:string, qNum:number, code:string, output:string, qtext:string) => {
        const cbg = CODE_BG[theme] || 'F7F8FA';
        const cfg = CODE_FG[theme] || DGRAY;
        const obg = OUT_BG[theme]  || 'F0FBF4';
        const LH  = codeFs + 3; // line height
        const HDR = codeFs + 12; // header height

        // Question banner
        need(30);
        y -= 8;
        fillR(LM,    y, 42, 26, BLUE);
        fillR(LM+42, y, CW-42, 26, NAVY);
        dt(`Q${qNum}`, LM+10, y-18, { bold:true, size:10, color:WHITE });
        dt(label,      LM+56, y-18, { bold:true, size:10, color:WHITE });
        y -= 32;

        if (qtext) {
            need(16);
            y = tl(qtext, LM, y, { italic:true, size:9, color:MGRAY, width:CW, gap:6 });
        }

        // Draw one block (code or output) cleanly across pages
        const drawBlock = (
            headerLabel: string,
            langTag: string,
            lines: string[],
            headerBg: string,
            bodyBg: string,
            bodyFg: string,
            borderCol: string,
        ) => {
            need(HDR + LH + 10);

            // ── Header ────────────────────────────────────────────────────
            fillR(LM, y, CW, HDR, headerBg);
            strokeR(LM, y, CW, HDR, borderCol, 1);
            dt(headerLabel, LM+10, y-HDR+4, { bold:true, size:codeFs, color:borderCol });
            if (langTag) {
                const lx = LM+10+fB.widthOfTextAtSize(headerLabel,codeFs)+6;
                dt(langTag, lx, y-HDR+4, { size:codeFs-1, color:MGRAY });
            }
            y -= HDR;

            // ── Body lines ────────────────────────────────────────────────
            // Track where the current page's body box started
            let boxTopY = y;

            for (let i=0; i<lines.length; i++) {
                if (y - LH < SAFE_BOTTOM) {
                    // Close box on current page with bottom stroke
                    const boxH = boxTopY - y;
                    fillR(LM, boxTopY, CW, boxH, bodyBg);
                    strokeR(LM, boxTopY, CW, boxH, borderCol, 0.8);
                    // Re-draw text on top of the fill (fill covers it otherwise)
                    // Re-render lines that are on this page segment
                    let ry = boxTopY;
                    for (let j=i-(Math.round(boxH/LH)); j<i; j++) {
                        if (j>=0) {
                            dt(lines[j]||' ', LM+10, ry-LH+3, { mono:true, size:codeFs-1, color:bodyFg });
                            ry-=LH;
                        }
                    }
                    newPage();
                    boxTopY = y;
                }
                y -= LH;
            }

            // ── Final padding + close box ─────────────────────────────────
            y -= 8; // bottom padding
            const bodyH = boxTopY - y;
            fillR(LM, boxTopY, CW, bodyH, bodyBg);
            strokeR(LM, boxTopY, CW, bodyH, borderCol, 0.8);
            // Draw all text ON TOP of background
            let ty = boxTopY;
            // Find the start index for this page's lines
            const totalLines = lines.length;
            const linesOnPage = Math.floor(bodyH / LH);
            const startIdx    = Math.max(0, totalLines - linesOnPage);
            for (let i=startIdx; i<totalLines; i++) {
                ty -= LH;
                dt(lines[i]||' ', LM+10, ty+3, { mono:true, size:codeFs-1, color:bodyFg });
            }
            y -= 4;
        };

        drawBlock('Source Code', language, code.split('\n'), BLUE2, cbg, cfg, BLUE);
        if (output && output.trim()) {
            need(20);
            y -= 4;
            drawBlock('Output', '', output.split('\n'), 'E8F8EE', obg, DGRAY, GREEN);
        }
        y -= 10;
    };

    // ══ TEMPLATE HEADERS ══════════════════════════════════════════════════
    if (tpl === 'A') {
        if (uni)  { y = tl(uni,  LM, y, { bold:true, size:13, color:NAVY, align:'center', width:CW, gap:2 }); }
        if (dept) { y = tl(dept, LM, y, { size:10,   color:NAVY, align:'center', width:CW, gap:3 }); }
        hLine(y, NAVY, 3); y-=3; hLine(y, NAVY, 1); y-=8;

        const W=[90,116,72,52,78,0];
        W[5]=CW-W[0]-W[1]-W[2]-W[3]-W[4];
        const aRows=[
            [{t:'Course Title:',w:W[0],bold:true,bg:BLUE2,color:NAVY},{t:courseStr,w:W[1],bg:WHITE},{t:'Course Code:',w:W[2],bold:true,bg:BLUE2,color:NAVY},{t:courseC,w:W[3],bg:WHITE},{t:'Credit Hours:',w:W[4],bold:true,bg:BLUE2,color:NAVY},{t:credits,w:W[5],bg:WHITE}],
            [{t:'Instructor:',w:W[0],bold:true,bg:BLUE2,color:NAVY},{t:instructor,w:W[1],bg:WHITE},{t:'Program:',w:W[2],bold:true,bg:BLUE2,color:NAVY},{t:prog||'BS(CS)',w:W[3]+W[4]+W[5],bg:WHITE}],
            [{t:'Semester:',w:W[0],bold:true,bg:BLUE2,color:NAVY},{t:sem,w:62,bg:WHITE},{t:'Batch:',w:62,bold:true,bg:BLUE2,color:NAVY},{t:batch,w:62,bg:WHITE},{t:'Section:',w:W[2],bold:true,bg:BLUE2,color:NAVY},{t:section,w:CW-W[0]-62-62-62-W[2],bg:WHITE}],
            [{t:'Date:',w:W[0],bold:true,bg:BLUE2,color:NAVY},{t:subDate,w:CW-W[0],bg:WHITE}],
            [{t:'Name:',w:W[0],bold:true,bg:BLUE2,color:NAVY},{t:name,w:W[1]+W[2],bg:WHITE},{t:'Reg No:',w:W[3]+10,bold:true,bg:BLUE2,color:NAVY},{t:reg,w:CW-W[0]-(W[1]+W[2])-(W[3]+10),bg:WHITE}],
        ];
        for (const row of aRows) { y = tr(row as any, y, 18); }
        y-=8;

    } else if (tpl === 'B') {
        y = tl(fullTitle, LM, y, { bold:true, size:14, color:NAVY, align:'center', width:CW, gap:4 });
        if (assignTtl) { y = tl(assignTtl, LM, y, { italic:true, size:10, color:MGRAY, align:'center', width:CW, gap:4 }); }
        y-=2;
        const hw=CW/4;
        const bRows=[
            [{t:'Student Name:',w:hw,bold:true,bg:BLUE2,color:NAVY},{t:name,w:hw,bg:WHITE},{t:'Reg No.:',w:hw,bold:true,bg:BLUE2,color:NAVY},{t:reg,w:hw,bg:WHITE}],
            [{t:'Course:',w:hw,bold:true,bg:BLUE2,color:NAVY},{t:courseStr,w:hw,bg:WHITE},{t:'Section:',w:hw,bold:true,bg:BLUE2,color:NAVY},{t:section,w:hw,bg:WHITE}],
            [{t:'Instructor:',w:hw,bold:true,bg:BLUE2,color:NAVY},{t:instructor,w:hw,bg:WHITE},{t:'Date:',w:hw,bold:true,bg:BLUE2,color:NAVY},{t:subDate,w:hw,bg:WHITE}],
            ...(sem?[[{t:'Semester:',w:hw,bold:true,bg:BLUE2,color:NAVY},{t:sem,w:hw,bg:WHITE},{t:'Batch:',w:hw,bold:true,bg:BLUE2,color:NAVY},{t:batch,w:hw,bg:WHITE}]]:[]),
        ];
        for (const row of bRows) { y = tr(row as any, y, 18); }
        y-=6;

    } else if (tpl === 'C') {
        const parts=[courseStr, section?`Section ${section}`:'', instructor].filter(Boolean);
        y = tl(fullTitle, LM, y, { bold:true, size:14, color:NAVY, align:'center', width:CW, gap:8 });
        y = tl(`${name}  |  ${reg}`, LM, y, { bold:true, size:10, color:DGRAY, align:'center', width:CW, gap:2 });
        y = tl(parts.join('  |  '),  LM, y, { size:9, color:MGRAY, align:'center', width:CW, gap:2 });
        y = tl(`Date: ${subDate}`,   LM, y, { size:9, color:MGRAY, align:'center', width:CW, gap:6 });
        hLine(y, BLUE, 1.5); y-=8;

    } else if (tpl === 'D') {
        y -= 30;
        if (uni)  { y = tl(uni,  LM, y, { bold:true, size:16, color:NAVY, align:'center', width:CW, gap:2 }); }
        if (dept) { y = tl(dept, LM, y, { size:10,   color:MGRAY, align:'center', width:CW, gap:4 }); }
        hLine(y, NAVY, 2); y -= 36;

        if (clo) {
            const cw = fB.widthOfTextAtSize(clo,10)+24;
            const cx = LM+(CW-cw)/2;
            fillR(cx, y, cw, 22, '1A6B8A');
            dt(clo, cx+12, y-15, { bold:true, size:10, color:WHITE });
            y -= 32;
        }
        y = tl(fullTitle, LM, y, { bold:true, size:18, color:NAVY, align:'center', width:CW, gap:6 });
        if (assignTtl) { y = tl(assignTtl, LM, y, { italic:true, size:11, color:MGRAY, align:'center', width:CW, gap:6 }); }
        if (project) {
            y -= 6;
            y = tl('Project / Topic:', LM, y, { bold:true, size:10, color:NAVY, align:'center', width:CW, gap:2 });
            y = tl(project, LM, y, { bold:true, italic:true, size:12, color:NAVY, align:'center', width:CW, gap:6 });
        }
        y -= 30;

        const hw = CW*0.3;
        const dRows: {t:string,w:number,bold?:boolean,bg?:string,color?:string}[][] = [
            [{t:'Submitted By:',w:hw,bold:true,bg:BLUE2,color:NAVY},{t:`${name} (${reg})`,w:CW-hw,bg:WHITE}],
            [{t:'Submitted To:',w:hw,bold:true,bg:BLUE2,color:NAVY},{t:instructor,w:CW-hw,bg:WHITE}],
            [{t:'Course:',w:hw,bold:true,bg:BLUE2,color:NAVY},{t:courseStr,w:CW-hw,bg:WHITE}],
            ...(section?[[{t:'Section:',w:hw,bold:true,bg:BLUE2,color:NAVY},{t:section,w:CW-hw,bg:WHITE}]]:[]),
            ...(prog?[[{t:'Program:',w:hw,bold:true,bg:BLUE2,color:NAVY},{t:prog,w:CW-hw,bg:WHITE}]]:[]),
            ...(sem?[[{t:'Semester:',w:hw,bold:true,bg:BLUE2,color:NAVY},{t:sem,w:CW-hw,bg:WHITE}]]:[]),
            [{t:'Date:',w:hw,bold:true,bg:BLUE2,color:NAVY},{t:subDate,w:CW-hw,bg:WHITE}],
        ];
        for (const row of dRows) { y = tr(row, y, 20); }
        newPage();
    }

    // Questions
    for (let qi=0; qi<questions.length; qi++) {
        const q=questions[qi];
        pdfCodeBlock(q.label||`Question ${qi+1}`, qi+1, q.code||'', q.output||'', q.text||'');
        if (q.notes) { need(16); y=tl(`Note: ${q.notes}`,LM,y,{italic:true,size:9,color:MGRAY,width:CW,gap:4}); }
    }

    // Footer on every page
    const total = pdfDoc.getPageCount();
    for (let i=0; i<total; i++) {
        const p = pdfDoc.getPage(i);
        p.drawLine({ start:{x:LM,y:BM+16}, end:{x:LM+CW,y:BM+16}, thickness:0.5, color:C('C8D4E8') });
        const ft=`${name}  |  ${fullTitle}  |  Page ${i+1} of ${total}`;
        const fw=fN.widthOfTextAtSize(ft,8);
        p.drawText(ft, { x:LM+(CW-fw)/2, y:BM-2, font:fN, size:8, color:C(MGRAY) });
    }

    fs.writeFileSync(outPath, await pdfDoc.save());
}
