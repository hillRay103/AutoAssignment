export function getWebviewContent(saved: any, initCode: string, language: string): string {
    const e  = (s: string) => String(s ?? '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/\"/g,'&quot;');
    const js = (s: string) => String(s ?? '').replace(/\\/g,'\\\\').replace(/`/g,'\\`').replace(/\$/g,'\\$');
    const today = new Date().toISOString().split('T')[0];

    return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<title>AutoAssignment</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',sans-serif;background:#0d1117;color:#cdd6e0;font-size:13px}
.hdr{background:linear-gradient(135deg,#0f2740,#1a4a80);padding:13px 20px;border-bottom:2px solid #1f6fb2}
.hdr h1{font-size:16px;color:#fff;font-weight:700;margin-bottom:2px}
.hdr small{color:#7eb8f7;font-size:11px}
.wrap{padding:14px 18px;max-width:900px}
.tabs{display:flex;gap:2px;margin-bottom:12px;border-bottom:2px solid #1f6fb2}
.tab{padding:7px 15px;border:none;background:#131c2e;color:#5a8ab8;cursor:pointer;border-radius:5px 5px 0 0;font-size:12px;font-weight:600;transition:.15s}
.tab:hover{color:#7eb8f7}.tab.on{background:#1f6fb2;color:#fff}
.pane{display:none}.pane.on{display:block}
.card{background:#131c2e;border:1px solid #1d2d45;border-radius:8px;padding:13px;margin-bottom:10px}
.stl{font-size:10px;font-weight:700;color:#5a8ab8;text-transform:uppercase;letter-spacing:.7px;margin-bottom:9px;padding-bottom:5px;border-bottom:1px solid #1d2d45}
.g2{display:grid;grid-template-columns:1fr 1fr;gap:9px}
.g3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:9px}
.g4{display:grid;grid-template-columns:repeat(4,1fr);gap:9px}
.fld{display:flex;flex-direction:column;gap:3px}
label{font-size:11px;color:#7a94b0;font-weight:500}
.req{color:#f87171;margin-left:2px}.opt{font-size:9px;background:#1d2d45;color:#4a6a8a;padding:1px 5px;border-radius:7px;margin-left:4px}
input,select,textarea{background:#0a1020;border:1px solid #1d2d45;color:#cdd6e0;padding:7px 9px;border-radius:5px;font-size:12px;font-family:inherit;transition:border-color .15s;width:100%}
input:focus,select:focus,textarea:focus{outline:none;border-color:#1f6fb2;background:#0d1628}
textarea{resize:vertical;font-family:'Consolas',monospace;font-size:11px;line-height:1.5}
/* Templates */
.tgrid{display:grid;grid-template-columns:repeat(2,1fr);gap:11px}
.tcard{border:2px solid #1d2d45;border-radius:8px;padding:11px;cursor:pointer;transition:.15s;background:#131c2e;position:relative;z-index:1}
.tcard:hover{border-color:#1f6fb2;z-index:2}.tcard.on{border-color:#1f6fb2;background:#0c1a2e;z-index:2}
.tcard h3{font-size:11px;color:#5a8ab8;margin-bottom:7px;font-weight:700}
.prev{background:#fff;border-radius:4px;padding:6px;font-size:6.5px;color:#111;line-height:1.5;min-height:88px;font-family:'Times New Roman',serif;overflow:hidden;pointer-events:none}
.cu-uni{text-align:center;font-weight:bold;font-size:8px;color:#1a3a6c;border-bottom:2px solid #1a3a6c;padding-bottom:2px;margin-bottom:3px}
.cu-tbl{width:100%;border-collapse:collapse;font-size:6px}
.cu-tbl td{border:1px solid #555;padding:1.5px 3px}
.cu-lb{font-weight:bold;background:#e8e8e8}
.cv-wrap{display:flex;height:82px;overflow:hidden}
.cv-bar{width:14px;background:#1a3a6c;border-radius:2px 0 0 2px;flex-shrink:0}
.cv-body{flex:1;padding:4px 5px;display:flex;flex-direction:column;justify-content:space-between;overflow:hidden}
.cv-tag{background:#1a6b8a;color:#fff;font-size:5px;padding:1px 4px;border-radius:2px;display:inline-block;margin-bottom:2px}
.cv-title{font-size:8px;font-weight:bold;color:#111;border-bottom:1px solid #1a3a6c;padding-bottom:2px;margin-bottom:2px}
.cv-sub{font-size:6px;color:#555;font-style:italic}
.cv-footer{font-size:5.5px;color:#333;text-align:right}
/* Questions */
.qcard{background:#0a1020;border:1px solid #1d2d45;border-radius:8px;padding:13px;margin-bottom:9px}
.qhdr{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:10px}
.qtitle{font-size:13px;font-weight:700;color:#1f6fb2}
.qfiletag{font-size:10px;font-style:italic;margin-top:2px}
.qfiletag.ok{color:#4caf7d}.qfiletag.warn{color:#fca5a5}
.rmv{background:#1e0808;color:#fca5a5;border:1px solid #5c1a1a;padding:4px 10px;border-radius:4px;cursor:pointer;font-size:11px}
.rmv:hover{background:#5c1a1a}
/* File load row */
.load-row{display:flex;gap:6px;align-items:center;margin-bottom:5px;flex-wrap:wrap}
.btn-browse{background:#1a4a80;color:#fff;border:none;padding:6px 14px;border-radius:5px;cursor:pointer;font-size:12px;font-weight:600;transition:.15s}
.btn-browse:hover{background:#2563b0}
.btn-active{background:#0c2e10;color:#4caf7d;border:1px solid #1a6b3a;padding:6px 12px;border-radius:5px;cursor:pointer;font-size:11px;transition:.15s}
.btn-active:hover{background:#1a6b3a;color:#fff}
.load-hint{font-size:10px;color:#4a6a8a;font-style:italic}
/* Output row */
.out-row{display:flex;gap:6px;align-items:center;margin-bottom:5px;flex-wrap:wrap}
.btn-paste{background:#0c2e10;color:#4caf7d;border:1px solid #1a6b3a;padding:5px 12px;border-radius:4px;cursor:pointer;font-size:11px;transition:.15s}
.btn-paste:hover{background:#1a6b3a;color:#fff}
.btn-run{background:#2c1a00;color:#c49a3a;border:1px solid #7a4f1a;padding:5px 12px;border-radius:4px;cursor:pointer;font-size:11px;transition:.15s}
.btn-run:hover{background:#7a4f1a;color:#fff}
.ostatus{font-size:10px;color:#3a6a9a;min-height:14px;font-style:italic;margin-top:2px}
/* Actions */
.acts{display:flex;gap:10px;margin-top:14px;padding-top:11px;border-top:1px solid #1d2d45}
.bp{background:linear-gradient(135deg,#0f2740,#1f6fb2);color:#fff;border:none;padding:10px 22px;border-radius:6px;cursor:pointer;font-size:13px;font-weight:700}
.bp:hover{opacity:.9}
.bs{background:#131c2e;color:#5a8ab8;border:1px solid #1f6fb2;padding:10px 14px;border-radius:6px;cursor:pointer;font-size:12px}
.fmts{display:flex;gap:10px}
.fmt{flex:1;padding:10px;border:2px solid #1d2d45;border-radius:6px;background:#131c2e;color:#4a6a8a;cursor:pointer;font-size:12px;text-align:center;transition:.15s}
.fmt:hover{border-color:#1f6fb2;color:#7eb8f7}.fmt.on{border-color:#1f6fb2;color:#fff;background:#0c1a2e;font-weight:700}
.st{margin-top:11px;padding:10px 13px;border-radius:5px;font-size:12px;display:none;white-space:pre-wrap;word-break:break-word;line-height:1.6}
.st.info{background:#0c1a2e;color:#93c5fd;border:1px solid #1f6fb2;display:block}
.st.error{background:#1a0808;color:#fca5a5;border:1px solid #5c1a1a;display:block}
.st.success{background:#051a10;color:#6ee7b7;border:1px solid #0d4a28;display:block}
</style>
</head>
<body>
<div class="hdr">
  <h1>⚡ AutoAssignment</h1>
  <small>Language: <strong>${e(language)}</strong> — Fill all tabs, then Generate</small>
</div>
<div class="wrap">

<div class="tabs">
  <button class="tab on" onclick="goTab('tpl')">① Template</button>
  <button class="tab"    onclick="goTab('inf')">② Info</button>
  <button class="tab"    onclick="goTab('qst')">③ Questions</button>
  <button class="tab"    onclick="goTab('exp')">④ Export</button>
</div>

<!-- ════ TAB 1: TEMPLATE ════ -->
<div id="pane-tpl" class="pane on">
  <div class="card">
    <div class="stl">Choose Document Template</div>
    <div class="tgrid">
      <div class="tcard ${(saved.defaultTemplate||'B')==='A'?'on':''}" onclick="selTpl('A')" id="tA">
        <h3>🏛️ Template A — University Header</h3>
        <div class="prev">
          <div class="cu-uni">COMSATS University — Sahiwal Campus</div>
          <table class="cu-tbl">
            <tr><td class="cu-lb">Course:</td><td>Data Structures</td><td class="cu-lb">Code:</td><td>CSC211</td><td class="cu-lb">Credits:</td><td>4(3,1)</td></tr>
            <tr><td class="cu-lb">Instructor:</td><td>Mr. Shehzad Ali</td><td class="cu-lb">Program:</td><td colspan="3">BS(CS)</td></tr>
            <tr><td class="cu-lb">Sem:</td><td>03</td><td class="cu-lb">Batch:</td><td>SP25</td><td class="cu-lb">Sec:</td><td>B</td></tr>
            <tr><td class="cu-lb">Name:</td><td colspan="2" style="font-weight:bold">Nisa Khawar</td><td class="cu-lb">Reg:</td><td colspan="2">SP25-BCS-082</td></tr>
          </table>
        </div>
      </div>
      <div class="tcard ${(saved.defaultTemplate||'B')==='B'?'on':''}" onclick="selTpl('B')" id="tB">
        <h3>📋 Template B — Compact Table</h3>
        <div class="prev">
          <div style="text-align:center;font-weight:bold;font-size:9px;color:#1a3a6c;margin-bottom:3px">Assignment No. 1</div>
          <table class="cu-tbl">
            <tr><td class="cu-lb">Name:</td><td>Nisa Khawar</td><td class="cu-lb">Reg:</td><td>SP25-BCS-082</td></tr>
            <tr><td class="cu-lb">Course:</td><td>Prog. Fund.</td><td class="cu-lb">Section:</td><td>B</td></tr>
            <tr><td class="cu-lb">Instructor:</td><td>Sir Ahmed</td><td class="cu-lb">Date:</td><td>07-04-2026</td></tr>
          </table>
          <div style="font-size:6px;color:#888;text-align:center;margin-top:2px">code begins below ↓</div>
        </div>
      </div>
      <div class="tcard ${(saved.defaultTemplate||'B')==='C'?'on':''}" onclick="selTpl('C')" id="tC">
        <h3>📄 Template C — Minimal</h3>
        <div class="prev" style="text-align:center;padding-top:8px">
          <div style="font-weight:bold;font-size:9px;color:#1a3a6c;border-bottom:1.5px solid #1a3a6c;padding-bottom:3px;margin-bottom:3px">Assignment No. 1</div>
          <div style="font-size:7px">Nisa Khawar | SP25-BCS-082</div>
          <div style="font-size:6.5px;color:#555">Prog. Fund. | Section B | Sir Ahmed</div>
        </div>
      </div>
      <div class="tcard ${(saved.defaultTemplate||'B')==='D'?'on':''}" onclick="selTpl('D')" id="tD">
        <h3>🎓 Template D — Cover Page</h3>
        <div class="prev">
          <div class="cv-wrap">
            <div class="cv-bar"></div>
            <div class="cv-body">
              <div>
                <div class="cv-tag">CLO 2-CSC291</div>
                <div class="cv-title">Software Engineering</div>
                <div class="cv-sub">Functional &amp; Non-Functional Requirements</div>
                <div style="font-size:6px;font-style:italic;font-weight:bold;color:#1a3a6c;border-top:1px solid #1a3a6c;padding-top:2px;margin-top:2px">Project Title: Online E-Commerce Website</div>
              </div>
              <div class="cv-footer">Submitted to: ASMA IQBAL<br/>SP25-BCS-082 — NISA KHAWAR<br/>04-02-2026</div>
            </div>
          </div>
        </div>
      </div>
    </div>
  </div>
  <div class="acts"><button class="bp" onclick="goTab('inf')">Next: Info →</button></div>
</div>

<!-- ════ TAB 2: INFO ════ -->
<div id="pane-inf" class="pane">
  <div class="card">
    <div class="stl">University <span class="opt">mainly for Template A &amp; D</span></div>
    <div class="g2">
      <div class="fld" style="grid-column:1/-1"><label>University Name</label>
        <input id="f_uni" value="${e(saved.university)}" placeholder="e.g. COMSATS University Islamabad"/></div>
      <div class="fld"><label>Campus / Department</label>
        <input id="f_dept" value="${e(saved.department)}" placeholder="e.g. Sahiwal Campus — Dept. of CS"/></div>
      <div class="fld"><label>Program</label>
        <input id="f_prog" value="${e(saved.programName)}" placeholder="e.g. BS(CS)"/></div>
    </div>
  </div>
  <div class="card">
    <div class="stl">Course Details</div>
    <div class="g2">
      <div class="fld"><label>Course Title <span class="req">*</span></label>
        <input id="f_course" value="${e(saved.courseName)}" placeholder="e.g. Data Structures"/></div>
      <div class="fld"><label>Course Code</label>
        <input id="f_code" value="${e(saved.courseCode)}" placeholder="e.g. CSC211"/></div>
      <div class="fld"><label>Credit Hours</label>
        <input id="f_credits" value="${e(saved.creditHours)}" placeholder="e.g. 4(3,1)"/></div>
      <div class="fld"><label>Course Instructor <span class="req">*</span></label>
        <input id="f_instructor" value="${e(saved.submittedTo)}" placeholder="e.g. Mr. Shehzad Ali"/></div>
    </div>
    <div class="g4" style="margin-top:9px">
      <div class="fld"><label>Semester</label><input id="f_sem" value="${e(saved.semester)}" placeholder="e.g. 03"/></div>
      <div class="fld"><label>Batch</label><input id="f_batch" value="${e(saved.batch)}" placeholder="e.g. SP25"/></div>
      <div class="fld"><label>Section <span class="req">*</span></label><input id="f_section" value="${e(saved.section)}" placeholder="e.g. B"/></div>
      <div class="fld"><label>Date</label><input id="f_date" type="date" value="${today}"/></div>
    </div>
  </div>
  <div class="card">
    <div class="stl">Student &amp; Assignment</div>
    <div class="g2">
      <div class="fld"><label>Student Name <span class="req">*</span></label>
        <input id="f_name" value="${e(saved.studentName)}" placeholder="e.g. Nisa Khawar"/></div>
      <div class="fld"><label>Registration Number <span class="req">*</span></label>
        <input id="f_reg" value="${e(saved.regNumber)}" placeholder="e.g. SP25-BCS-082"/></div>
      <div class="fld"><label>Assignment Number <span class="req">*</span></label>
        <input id="f_anum" type="number" min="1" value="${e(String(saved.assignmentCounter??1))}"/></div>
      <div class="fld"><label>Assignment Title <span class="opt">optional</span></label>
        <input id="f_atitle" placeholder="e.g. Pointers and Arrays"/></div>
    </div>
    <div class="g2" style="margin-top:9px">
      <div class="fld"><label>CLO Label <span class="opt">Template D only — e.g. CLO 2-CSC291</span></label>
        <input id="f_clo" placeholder="e.g. CLO 2-CSC291"/></div>
      <div class="fld"><label>Project / Topic Title <span class="opt">Template D only</span></label>
        <input id="f_project" placeholder="e.g. Online E-Commerce Website"/></div>
    </div>
  </div>
  <div class="acts">
    <button class="bs" onclick="goTab('tpl')">← Back</button>
    <button class="bp" onclick="goTab('qst')">Next: Questions →</button>
  </div>
</div>

<!-- ════ TAB 3: QUESTIONS ════ -->
<div id="pane-qst" class="pane">
  <div class="card" style="background:#0c1e10;border-color:#1a6b3a">
    <div style="font-size:12px;font-weight:700;color:#4caf7d;margin-bottom:6px">📂 How to load code for each question</div>
    <div style="font-size:11px;color:#7ab89a;line-height:1.8">
      <strong>Method 1 — Browse File (recommended):</strong> Click the blue <strong>📂 Browse &amp; Load File</strong> button → an OS file picker opens → select any .cpp/.py/.c file → code loads instantly. Works for Q1, Q2, Q3 — any file, any name.<br/>
      <strong>Method 2 — Active Editor:</strong> Click on a file tab in VS Code, then click the green button. Useful if the file is already open.<br/>
      <strong>For output:</strong> Run your code in the VS Code terminal → copy all the output text → click <strong>📋 Paste Output</strong>.
    </div>
  </div>
  <div class="card">
    <div class="stl">Questions / Programs</div>
    <div id="qcon"></div>
    <button class="btn-browse" onclick="addQ()" style="margin-top:4px">＋ Add Question</button>
  </div>
  <div class="acts">
    <button class="bs" onclick="goTab('inf')">← Back</button>
    <button class="bp" onclick="goTab('exp')">Next: Export →</button>
  </div>
</div>

<!-- ════ TAB 4: EXPORT ════ -->
<div id="pane-exp" class="pane">
  <div class="card">
    <div class="stl">Output Format</div>
    <div class="fmts">
      <div class="fmt on" id="fdocx" onclick="selFmt('docx')">📄 Word (.docx)</div>
      <div class="fmt"    id="fpdf"  onclick="selFmt('pdf')">📕 PDF</div>
    </div>
  </div>
  <div class="card">
    <div class="stl">Style</div>
    <div class="g3">
      <div class="fld"><label>Font</label>
        <select id="f_font">
          <option value="Calibri" ${(saved.fontChoice||'Calibri')==='Calibri'?'selected':''}>Calibri (default)</option>
          <option value="Times New Roman" ${saved.fontChoice==='Times New Roman'?'selected':''}>Times New Roman</option>
        </select>
      </div>
      <div class="fld"><label>Code Block Background</label>
        <select id="f_theme">
          <option value="lightgray">Light Gray (default)</option>
          <option value="lightblue">Light Blue</option>
          <option value="darkgray">Dark (editor style)</option>
          <option value="white">White / None</option>
        </select>
      </div>
      <div class="fld"><label>Code Font Size</label>
        <select id="f_codesize">
          <option value="10">10pt</option>
          <option value="11" selected>11pt</option>
          <option value="12">12pt</option>
        </select>
      </div>
    </div>
  </div>
  <div class="acts">
    <button class="bs" onclick="goTab('qst')">← Back</button>
    <button class="bp" onclick="doGenerate()">⚡ Generate Document</button>
  </div>
  <div id="statusBox" class="st"></div>
</div>

</div>
<script>
const vscode = acquireVsCodeApi();
const INIT_CODE = \`${js(initCode)}\`;
let curTpl = '${e(saved.defaultTemplate||'B')}';
let curFmt = 'docx';
let qCount = 0;
// Store per-question file paths and labels for auto-run and question detection
const qFilePaths  = {};
const qFileLabels = {};

const TABS = ['tpl','inf','qst','exp'];
function goTab(id){
  document.querySelectorAll('.pane').forEach(p=>p.classList.remove('on'));
  document.querySelectorAll('.tab').forEach(t=>t.classList.remove('on'));
  document.getElementById('pane-'+id).classList.add('on');
  document.querySelectorAll('.tab')[TABS.indexOf(id)].classList.add('on');
}

// FIX 1: Template selection — clears ALL tcard .on states before setting new one
function selTpl(t){
  curTpl = t;
  document.querySelectorAll('.tcard').forEach(c => c.classList.remove('on'));
  const card = document.getElementById('t'+t);
  if(card) card.classList.add('on');
}

function selFmt(f){
  curFmt=f;
  document.getElementById('fdocx').classList.toggle('on',f==='docx');
  document.getElementById('fpdf').classList.toggle('on',f==='pdf');
}

function addQ(){
  qCount++;
  const n = qCount;
  const div = document.createElement('div');
  div.className = 'qcard'; div.id = 'q'+n;
  div.innerHTML =
    '<div class="qhdr">'+
      '<div>'+
        '<div class="qtitle">Question '+n+'</div>'+
        '<div class="qfiletag" id="qft'+n+'">no file loaded</div>'+
      '</div>'+
      '<button class="rmv" onclick="rmQ('+n+')">✕ Remove</button>'+
    '</div>'+
    '<div class="g2" style="margin-bottom:10px">'+
      '<div class="fld"><label>Question Label</label>'+
        '<input id="ql'+n+'" value="Question '+n+'" placeholder="e.g. Q1, Q2 Part (a)"/></div>'+
      '<div class="fld"><label>Question Statement <span class="opt">optional</span></label>'+
        '<input id="qt'+n+'" placeholder="e.g. Write a program to find factorial"/></div>'+
    '</div>'+
    '<div class="fld" style="margin-bottom:10px"><label>Source Code</label>'+
      '<div class="load-row">'+
        '<button class="btn-browse" onclick="browseFile('+n+')">📂 Browse &amp; Load File</button>'+
        '<button class="btn-active" onclick="loadActive('+n+')">↙ From Active Tab</button>'+
        '<span class="load-hint">← Browse opens a file picker — easiest way!</span>'+
      '</div>'+
      '<textarea id="qc'+n+'" rows="8" placeholder="Click \\'Browse &amp; Load File\\' above to select any .cpp/.py file — or paste code directly here"></textarea>'+
    '</div>'+
    '<div class="fld" style="margin-bottom:10px"><label>Program Output</label>'+
      '<div class="out-row">'+
        '<button class="btn-paste" onclick="pasteClip('+n+')">📋 Paste Output</button>'+
        '<button class="btn-run" onclick="autoRun('+n+')">▶ Auto-Run</button>'+
        '<span class="load-hint">Run in terminal → copy output → Paste</span>'+
      '</div>'+
      '<div class="ostatus" id="os'+n+'"></div>'+
      '<textarea id="qo'+n+'" rows="4" placeholder="Paste output here, or type manually"></textarea>'+
    '</div>'+
    '<div class="fld"><label>Notes <span class="opt">optional</span></label>'+
      '<input id="qn'+n+'" placeholder="e.g. Used recursion to solve the problem"/></div>';
  document.getElementById('qcon').appendChild(div);
}

function rmQ(n){ document.getElementById('q'+n)?.remove(); }

// ── Method 1: OS file picker ────────────────────────────────────────────────
function browseFile(n){
  const label = document.getElementById('ql'+n)?.value || 'Question '+n;
  const tag   = document.getElementById('qft'+n);
  tag.textContent = '⏳ Opening file picker...'; tag.className = 'qfiletag';
  // FIX 2: pass the question index and current label so extension can track per-question files
  vscode.postMessage({ command:'browseFile', qIndex:n, label, questionNumber:n });
}

// ── Method 2: Load from active editor tab ────────────────────────────────────
function loadActive(n){
  const tag = document.getElementById('qft'+n);
  tag.textContent = '⏳ Reading active editor...'; tag.className = 'qfiletag';
  vscode.postMessage({ command:'getActiveCode', qIndex:n, questionNumber:n });
}

function pasteClip(n){
  document.getElementById('os'+n).textContent = '⏳ Reading clipboard...';
  vscode.postMessage({ command:'pasteClipboard', qIndex:n });
}

function autoRun(n){
  // FIX 2: use the file path specifically stored for question n
  const fp = qFilePaths[n] || '';
  if(!fp){
    document.getElementById('os'+n).textContent = '⚠️ No file loaded for this question — browse a file first';
    return;
  }
  document.getElementById('os'+n).textContent = '⏳ Running...';
  vscode.postMessage({ command:'runCode', qIndex:n, filePath: fp });
}

// ── Collect & generate ────────────────────────────────────────────────────────
function doGenerate(){
  const name = document.getElementById('f_name').value.trim();
  if(!name){ goTab('inf'); document.getElementById('f_name').focus(); alert('Please enter your Student Name.'); return; }
  const qs = [];
  document.querySelectorAll('.qcard').forEach(card=>{
    const id = card.id.replace('q','');
    qs.push({
      label:  document.getElementById('ql'+id)?.value || 'Q'+id,
      text:   document.getElementById('qt'+id)?.value || '',
      code:   document.getElementById('qc'+id)?.value || '',
      output: document.getElementById('qo'+id)?.value || '',
      notes:  document.getElementById('qn'+id)?.value || '',
    });
  });
  if(!qs.length){ goTab('qst'); alert('Please add at least one question.'); return; }
  goTab('exp');
  setStatus('⏳ Generating — please wait...','info');
  vscode.postMessage({ command:'generate', data:{
    template:        curTpl,
    studentName:     name,
    regNumber:       document.getElementById('f_reg').value.trim(),
    courseName:      document.getElementById('f_course').value.trim(),
    courseCode:      document.getElementById('f_code').value.trim(),
    section:         document.getElementById('f_section').value.trim(),
    submittedTo:     document.getElementById('f_instructor').value.trim(),
    department:      document.getElementById('f_dept').value.trim(),
    university:      document.getElementById('f_uni').value.trim(),
    programName:     document.getElementById('f_prog').value.trim(),
    semester:        document.getElementById('f_sem').value.trim(),
    batch:           document.getElementById('f_batch').value.trim(),
    creditHours:     document.getElementById('f_credits').value.trim(),
    assignmentNumber:parseInt(document.getElementById('f_anum').value)||1,
    assignmentTitle: document.getElementById('f_atitle').value.trim(),
    submissionDate:  document.getElementById('f_date').value,
    cloLabel:        document.getElementById('f_clo').value.trim(),
    projectTitle:    document.getElementById('f_project').value.trim(),
    outputFormat:    curFmt,
    fontChoice:      document.getElementById('f_font').value,
    codeTheme:       document.getElementById('f_theme').value,
    codeFontSize:    document.getElementById('f_codesize').value,
    language:        '${e(language)}',
    questions:       qs,
  }});
}

function setStatus(msg, type){
  const el = document.getElementById('statusBox');
  el.textContent = msg; el.className = 'st ' + type;
}

// ── Messages from extension ───────────────────────────────────────────────────
window.addEventListener('message', ev=>{
  const m = ev.data;

  // FIX 2: fileLoaded uses qIndex to set code in the CORRECT question textarea
  if(m.command === 'fileLoaded'){
    const el  = document.getElementById('qc'+m.qIndex);
    const tag = document.getElementById('qft'+m.qIndex);
    if(m.code === null){
      if(tag){ tag.textContent = m.filename || 'Cancelled'; tag.className = 'qfiletag warn'; }
      return;
    }
    if(el)  el.value = m.code || '';
    if(tag){ tag.textContent = m.filename ? '📄 ' + m.filename : 'loaded'; tag.className = 'qfiletag ok'; }
    // Store file path keyed strictly by qIndex — each question keeps its own path
    if(m.filePath) qFilePaths[m.qIndex] = m.filePath;
  }

  if(m.command === 'clipboardText'){
    const el = document.getElementById('qo'+m.qIndex);
    const st = document.getElementById('os'+m.qIndex);
    if(el) el.value = m.text || '';
    if(st) st.textContent = m.text ? '✅ Pasted ' + m.text.split('\\n').length + ' lines!' : '⚠️ Clipboard was empty';
  }

  if(m.command === 'codeOutput'){
    const el = document.getElementById('qo'+m.qIndex);
    const st = document.getElementById('os'+m.qIndex);
    if(el) el.value = m.output;
    if(st) st.textContent = '✅ Output captured!';
  }
  if(m.command === 'runStatus'){
    const st = document.getElementById('os'+m.qIndex);
    if(st) st.textContent = m.text;
  }
  if(m.command === 'status'){ goTab('exp'); setStatus(m.text, m.type||'info'); }
});

// Auto-load Q1 with the initially open file
addQ();
document.getElementById('qc1').value = INIT_CODE;
const initTag = document.getElementById('qft1');
if(initTag){ initTag.textContent = '📄 (current editor file)'; initTag.className = 'qfiletag ok'; }
// FIX 2: Q1's file path will be filled in by extension via 'fileLoaded' for auto-run;
// set empty string as placeholder so autoRun guard doesn't block it
qFilePaths[1] = '__init__';
</script>
</body>
</html>`;
}
