import * as vscode from 'vscode';
import * as path from 'path';
import * as fs from 'fs';
import { spawn } from 'child_process';
import { getWebviewContent } from './webview';
import { buildDocx, buildPdf } from './generator';

// ══════════════════════════════════════════════════════════════════════════════
export function activate(context: vscode.ExtensionContext) {

    const cmd = vscode.commands.registerCommand('autoassignment.generate', async () => {

        const editor = vscode.window.activeTextEditor;
        if (!editor) {
            vscode.window.showErrorMessage('AutoAssignment: Open a code file first.');
            return;
        }

        const initFile = editor.document.fileName;
        const initExt  = path.extname(initFile).toLowerCase();
        const initCode = editor.document.getText();
        const langMap: Record<string, string> = {
            '.py': 'Python', '.c': 'C', '.cpp': 'C++',
            '.java': 'Java', '.js': 'JavaScript', '.cs': 'C#',
        };
        const language = langMap[initExt] || 'Unknown';

        const cfg = vscode.workspace.getConfiguration('autoassignment');
        const saved = {
            studentName:       cfg.get<string>('studentName',       ''),
            regNumber:         cfg.get<string>('regNumber',         ''),
            courseName:        cfg.get<string>('courseName',        ''),
            courseCode:        cfg.get<string>('courseCode',        ''),
            section:           cfg.get<string>('section',           ''),
            submittedTo:       cfg.get<string>('submittedTo',       ''),
            department:        cfg.get<string>('department',        ''),
            university:        cfg.get<string>('university',        ''),
            programName:       cfg.get<string>('programName',       ''),
            semester:          cfg.get<string>('semester',          ''),
            batch:             cfg.get<string>('batch',             ''),
            creditHours:       cfg.get<string>('creditHours',       ''),
            fontChoice:        cfg.get<string>('fontChoice',        'Calibri'),
            outputFormat:      cfg.get<string>('outputFormat',      'docx'),
            assignmentCounter: cfg.get<number>('assignmentCounter', 1),
            defaultTemplate:   cfg.get<string>('defaultTemplate',   'B'),
        };

        const panel = vscode.window.createWebviewPanel(
            'autoassignment', 'AutoAssignment',
            vscode.ViewColumn.Beside,
            { enableScripts: true, retainContextWhenHidden: true }
        );
        panel.webview.html = getWebviewContent(saved, initCode, language);

        // Register init file path for Q1
        panel.webview.postMessage({
            command: 'fileLoaded', code: null, filename: null,
            filePath: initFile, qIndex: 1,
        });

        // ── Message handler ────────────────────────────────────────────────────
        panel.webview.onDidReceiveMessage(async (msg) => {

            // Browse file
            if (msg.command === 'browseFile') {
                const result = await vscode.window.showOpenDialog({
                    canSelectMany: false,
                    openLabel: `Load code for ${msg.label || 'Question'}`,
                    filters: { 'Code files': ['py','c','cpp','java','js','cs','ts','h','hpp'], 'All files': ['*'] }
                });
                if (!result || result.length === 0) {
                    panel.webview.postMessage({ command: 'fileLoaded', code: null, filename: null, qIndex: msg.qIndex });
                    return;
                }
                try {
                    const fp = result[0].fsPath;
                    panel.webview.postMessage({ command: 'fileLoaded', code: fs.readFileSync(fp,'utf8'), filename: path.basename(fp), filePath: fp, qIndex: msg.qIndex });
                } catch (e) {
                    panel.webview.postMessage({ command: 'fileLoaded', code: null, filename: `Error: ${e}`, qIndex: msg.qIndex });
                }
            }

            // Load from active editor
            if (msg.command === 'getActiveCode') {
                const visibles = vscode.window.visibleTextEditors.filter(e => e.document.uri.scheme === 'file' && !e.document.isUntitled);
                const best = visibles.length > 0 ? visibles[visibles.length-1] : vscode.window.activeTextEditor;
                if (!best) { panel.webview.postMessage({ command:'fileLoaded', code:'', filename:'⚠️ No file open', qIndex:msg.qIndex }); return; }
                panel.webview.postMessage({ command:'fileLoaded', code:best.document.getText(), filename:path.basename(best.document.fileName), filePath:best.document.fileName, qIndex:msg.qIndex });
            }

            // Paste clipboard
            if (msg.command === 'pasteClipboard') {
                try { panel.webview.postMessage({ command:'clipboardText', text: await vscode.env.clipboard.readText()||'', qIndex:msg.qIndex }); }
                catch (_) { panel.webview.postMessage({ command:'clipboardText', text:'', qIndex:msg.qIndex }); }
            }

            // Auto-run
            if (msg.command === 'runCode') {
                const runPath = msg.filePath && msg.filePath !== '__init__' ? msg.filePath : initFile;
                panel.webview.postMessage({ command:'runStatus', text:'⏳ Running...', qIndex:msg.qIndex });
                runAndCapture(runPath, path.extname(runPath).toLowerCase(), output => {
                    panel.webview.postMessage({ command:'codeOutput', output, qIndex:msg.qIndex });
                });
            }

            // Generate document — pure TypeScript, no Python needed
            if (msg.command === 'generate') {
                const d = msg.data;

                const toSave: Record<string,unknown> = {
                    studentName:d.studentName, regNumber:d.regNumber, courseName:d.courseName,
                    courseCode:d.courseCode, section:d.section, submittedTo:d.submittedTo,
                    department:d.department, university:d.university, programName:d.programName,
                    semester:d.semester, batch:d.batch, creditHours:d.creditHours,
                    fontChoice:d.fontChoice, defaultTemplate:d.template,
                    assignmentCounter: d.assignmentNumber + 1,
                };
                for (const [k,v] of Object.entries(toSave)) {
                    cfg.update(k, v, vscode.ConfigurationTarget.Global).then(()=>{},()=>{});
                }

                const saveDir  = path.dirname(initFile);
                const safeName = (d.studentName||'Student').replace(/[^a-zA-Z0-9]/g,'_');
                const ext2     = d.outputFormat === 'pdf' ? 'pdf' : 'docx';
                const padNo    = String(d.assignmentNumber).padStart(2,'0');
                let outPath    = path.join(saveDir, `Assignment${padNo}_${safeName}.${ext2}`);
                let ver = 1;
                while (fs.existsSync(outPath)) { ver++; outPath = path.join(saveDir, `Assignment${padNo}_${safeName}_v${ver}.${ext2}`); }

                panel.webview.postMessage({ command:'status', text:'⏳ Generating...', type:'info' });

                try {
                    if (d.outputFormat === 'pdf') {
                        await buildPdf(d, outPath);
                    } else {
                        await buildDocx(d, outPath);
                    }

                    if (!fs.existsSync(outPath)) { throw new Error('Output file was not created'); }

                    panel.webview.postMessage({ command:'status', text:`✅ Done! Saved: ${path.basename(outPath)}`, type:'success' });
                    const pick = await vscode.window.showInformationMessage(`✅ ${path.basename(outPath)} ready!`, 'Open File', 'Show in Folder');
                    if (pick === 'Open File')      { vscode.env.openExternal(vscode.Uri.file(outPath)); }
                    if (pick === 'Show in Folder') { vscode.commands.executeCommand('revealFileInOS', vscode.Uri.file(outPath)); }

                } catch (err: any) {
                    panel.webview.postMessage({ command:'status', text:`❌ ${err.message || err}`, type:'error' });
                }
            }
        });
    });

    context.subscriptions.push(cmd);
}

function runAndCapture(filePath: string, ext: string, cb: (out: string) => void) {
    const dir  = path.dirname(filePath);
    const base = path.basename(filePath, ext);
    const cr   = (comp: string, ca: string[], re: string, ra: string[]) => {
        const c = spawn(comp, ca, { cwd:dir, shell:true }); let ce = '';
        c.stderr.on('data', (d: Buffer) => { ce += d.toString(); });
        c.on('close', (code: number) => { if (code !== 0) { cb(`[Compile Error]\n${ce}`); return; } capture(re, ra, dir, cb); });
        c.on('error', () => cb(`[${comp} not found in PATH]`));
    };
    if (ext==='.py')   { return capture('python', [filePath], dir, cb); }
    if (ext==='.js')   { return capture('node',   [filePath], dir, cb); }
    if (ext==='.c')    { return cr('gcc',  [filePath,'-o',base+'.exe'], base+'.exe', []); }
    if (ext==='.cpp')  { return cr('g++',  [filePath,'-o',base+'.exe'], base+'.exe', []); }
    if (ext==='.java') {
        const j = spawn('javac', [filePath], { cwd:dir, shell:true }); let je='';
        j.stderr.on('data',(d:Buffer)=>{ je+=d.toString(); });
        j.on('close',(c:number)=>{ if(c!==0){cb(`[Compile Error]\n${je}`);return;} capture('java',[base],dir,cb); });
        return;
    }
    cb('[Auto-run not supported for this file type]');
}

function capture(cmd: string, args: string[], cwd: string, cb: (out: string) => void) {
    let out='', err='';
    const p = spawn(cmd, args, { cwd, timeout:10000, shell:true });
    p.stdout.on('data',(d:Buffer)=>{ out+=d.toString(); });
    p.stderr.on('data',(d:Buffer)=>{ err+=d.toString(); });
    p.on('close',()=> cb((out+err).trim()||'[No output]'));
    p.on('error',()=> cb(`[Could not start: ${cmd}]`));
}

export function deactivate() {}
