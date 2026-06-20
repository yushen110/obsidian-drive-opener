const { Plugin, ItemView, Menu, Notice, MarkdownView, Modal, Setting, TFile, TFolder } = require('obsidian');
const { readdirSync, statSync, existsSync, copyFileSync, renameSync, unlinkSync, rmdirSync, mkdirSync, readFileSync, writeFileSync, createReadStream } = require('fs');
const { join, dirname, extname, basename, resolve } = require('path');
const { exec, execSync, spawn } = require('child_process');
const os = require('os');
const VIEW_TYPE = 'drive-explorer';

/* ===================================================================
   全局文件剪贴板（多点粘贴支持）
   =================================================================== */
const fileClipboard = {
  _items: [], // { sources: string[], action: 'copy'|'cut', timestamp: number }
  get items() { return this._items; },
  push(sources, action) {
    this._items.unshift({ sources: [...sources], action, timestamp: Date.now() });
    if (this._items.length > 10) this._items.pop();
  },
  clear() { this._items = []; },
  remove(idx) { this._items.splice(idx, 1); },
};

/* ===================================================================
   工具函数
   =================================================================== */
/** 安全执行 PowerShell 脚本：写入临时 .ps1 文件 + -File 执行，避免命令注入 */
function runPowershell(script) {
  return new Promise((resolve, reject) => {
    const tmpFile = join(os.tmpdir(), `de_ps_${Date.now()}_${Math.random().toString(36).slice(2)}.ps1`);
    // UTF-8 BOM 防止 Windows PowerShell 5.1 按 ANSI(GBK) 误读中文路径
    writeFileSync(tmpFile, '﻿' + script.replace(/\r\n/g, '\n').replace(/\r/g, '\n'), { encoding: 'utf8' });
    const ps = spawn('powershell', ['-ExecutionPolicy', 'Bypass', '-NoProfile', '-File', tmpFile],
      { windowsHide: true, timeout: 30000 });
    const chunks = []; const errChunks = [];
    ps.stdout.on('data', d => chunks.push(d));
    ps.stderr.on('data', d => errChunks.push(d));
    ps.on('close', code => {
      try { unlinkSync(tmpFile); } catch {}
      const out = Buffer.concat(chunks).toString('utf8').trim();
      if (code === 0) return resolve(out);
      const errDetail = errChunks.length ? Buffer.concat(errChunks).toString('utf8').trim() : '';
      reject(new Error(errDetail || out || `exit code ${code}`));
    });
    ps.on('error', err => { try { unlinkSync(tmpFile); } catch {} reject(err); });
  });
}

function isOfficeFile(name) {
  return ['.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx'].includes(extname(name).toLowerCase());
}
function isPDFFile(name) {
  return extname(name).toLowerCase() === '.pdf';
}
function isImageFile(name) {
  return ['.jpg','.jpeg','.png','.gif','.bmp','.webp','.svg'].includes(extname(name).toLowerCase());
}
function isTextFile(name) {
  return ['.md','.txt','.csv','.json','.xml','.yaml','.yml','.log','.ini','.cfg'].includes(extname(name).toLowerCase());
}

function safeDestPath(srcPath, destDir) {
  let name = basename(srcPath);
  let dest = join(destDir, name);
  if (!existsSync(dest)) return dest;
  const ext = extname(name);
  const base = name.slice(0, -ext.length) || name;
  for (let i = 1; i < 1000; i++) {
    dest = join(destDir, `${base} (${i})${ext}`);
    if (!existsSync(dest)) return dest;
  }
  return join(destDir, `${base} (${Date.now()})${ext}`);
}

function copyDirRecursive(src, dest) {
  if (!existsSync(dest)) mkdirSync(dest, { recursive: true });
  for (const name of readdirSync(src, { encoding: 'utf8' })) {
    try {
      const s = join(src, name), d = join(dest, name);
      statSync(s).isDirectory() ? copyDirRecursive(s, d) : copyFileSync(s, d);
    } catch (e) {
      // 单个文件失败不中断整批，继续处理其余文件
      console.error(`Drive Opener: 复制失败 ${name}: ${e.message}`);
    }
  }
}

function removeDirRecursive(dir) {
  for (const name of readdirSync(dir, { encoding: 'utf8' })) {
    try {
      const p = join(dir, name);
      statSync(p).isDirectory() ? removeDirRecursive(p) : unlinkSync(p);
    } catch (e) {
      console.error(`Drive Opener: 删除失败 ${name}: ${e.message}`);
    }
  }
  try { rmdirSync(dir); } catch {}
}

/** 获取系统可用盘符 */
function getAvailableDrives() {
  const drives = [];
  try {
    const out = execSync('wmic logicaldisk get name,drivetype 2>nul', { encoding: 'utf8', timeout: 5000 });
    for (const line of out.split('\n')) {
      const m = line.match(/^([A-Za-z]):\s+(\d+)/);
      if (m) {
        const letter = m[1] + ':\\';
        const type = parseInt(m[2]);
        let label = '';
        let sizeInfo = '';
        try {
          if (existsSync(letter)) {
            const s = statSync(letter);
            sizeInfo = '';
          }
          // Get volume label
          const volOut = execSync(`vol ${letter} 2>nul`, { encoding: 'utf8', timeout: 2000 });
          const volM = volOut.match(/卷的序列号|Volume in drive/i);
          const labelM = volOut.match(/^(.+?)\s+\(/m) || volOut.match(/^(.+?)\n/m);
          if (labelM) label = labelM[1].trim();
        } catch {}
        const typeNames = { 2: '可移动磁盘', 3: '本地磁盘', 4: '网络磁盘', 5: '光盘' };
        drives.push({ path: letter, type: typeNames[type] || '未知', label });
      }
    }
  } catch {}
  if (drives.length === 0) {
    // Fallback: try common letters
    for (let c = 67; c <= 90; c++) { // C: through Z:
      const letter = String.fromCharCode(c) + ':\\';
      try { if (existsSync(letter)) drives.push({ path: letter, type: '本地磁盘', label: '' }); } catch {}
    }
  }
  return drives;
}

/** 检测 7-Zip 是否安装 */
function is7zInstalled() {
  try {
    execSync('where 7z 2>nul', { encoding: 'utf8', timeout: 2000 });
    return true;
  } catch {
    // 也检查标准安装路径（用户可能未将 7-Zip 加入 PATH）
    const stdPaths = [
      'C:\\Program Files\\7-Zip\\7z.exe',
      'C:\\Program Files (x86)\\7-Zip\\7z.exe',
      join(os.homedir(), 'scoop\\apps\\7zip\\current\\7z.exe'),
    ];
    for (const p of stdPaths) {
      try { if (existsSync(p)) return true; } catch {}
    }
    return false;
  }
}

/** 查找 7z.exe 的完整路径 */
function find7zPath() {
  try {
    const out = execSync('where 7z 2>nul', { encoding: 'utf8', timeout: 2000 });
    const lines = out.trim().split('\n');
    if (lines[0] && existsSync(lines[0].trim())) return lines[0].trim();
  } catch {}
  const stdPaths = [
    'C:\\Program Files\\7-Zip\\7z.exe',
    'C:\\Program Files (x86)\\7-Zip\\7z.exe',
    join(os.homedir(), 'scoop\\apps\\7zip\\current\\7z.exe'),
  ];
  for (const p of stdPaths) {
    try { if (existsSync(p)) return p; } catch {}
  }
  return '7z'; // fallback to PATH
}

/** 使用 Windows Forms API 将实际文件复制到系统剪贴板 */
function copyFilesToOsClipboard(paths) {
  return new Promise((resolve, reject) => {
    if (!paths || paths.length === 0) return reject(new Error('无文件'));
    // 写临时 .ps1 脚本，取巧路径引号转义问题
    const tmpFile = join(os.tmpdir(), `de_clip_${Date.now()}.ps1`);
    const psLines = [
      'Add-Type -AssemblyName System.Windows.Forms',
      '$list = New-Object System.Collections.Specialized.StringCollection',
      ...paths.map(p => `[void]$list.Add('${p.replace(/'/g, "''")}')`),
      'try {',
      '  [System.Windows.Forms.Clipboard]::SetFileDropList($list)',
      "  Write-Output 'OK'",
      '} catch {',
      "  Write-Output ('ERR:' + $_.Exception.Message)",
      '}',
    ];
    // ⚠️ 必须加 UTF-8 BOM：Windows PowerShell 5.1 无 BOM 时按 ANSI (GBK) 读，
    // 中文/特殊字符路径会乱码导致 SetFileDropList 找不到文件
    writeFileSync(tmpFile, '﻿' + psLines.join('\r\n'), { encoding: 'utf8' });
    // -STA 必须：Clipboard API 要求 STA 线程模型
    const ps = spawn('powershell', ['-ExecutionPolicy', 'Bypass', '-STA', '-NoProfile', '-File', tmpFile],
      { windowsHide: true, timeout: 15000 });
    const chunks = []; const errChunks = [];
    ps.stdout.on('data', d => chunks.push(d));
    ps.stderr.on('data', d => errChunks.push(d));
    ps.on('close', code => {
      try { unlinkSync(tmpFile); } catch {}
      const out = Buffer.concat(chunks).toString('utf8').trim();
      if (out === 'OK') return resolve();
      const errDetail = errChunks.length ? Buffer.concat(errChunks).toString('utf8').trim() : '';
      const errMsg = out.startsWith('ERR:') ? out.slice(4) : (errDetail || `exit code ${code}`);
      reject(new Error(errMsg));
    });
    ps.on('error', err => { try { unlinkSync(tmpFile); } catch {} reject(err); });
  });
}
function compressWith7z(targets, archivePath, format = 'zip') {
  const targetList = targets.map(t => `"${t}"`).join(' ');
  const sevenZip = find7zPath();
  return new Promise((resolve, reject) => {
    exec(`"${sevenZip}" a -t${format} "${archivePath}" ${targetList}`, { maxBuffer: 16 * 1024 * 1024 }, (err, stdout, stderr) => {
      if (err && !stdout.includes('Everything is Ok')) return reject(err);
      resolve(stdout);
    });
  });
}

/** 使用 7-Zip 解压（支持 ZIP/7z/RAR/TAR/GZ/BZ2 等几乎所有格式） */
function extractWith7z(archivePath, outDir) {
  const sevenZip = find7zPath();
  return new Promise((resolve, reject) => {
    exec(`"${sevenZip}" x "${archivePath}" -o"${outDir}" -y`, { maxBuffer: 16 * 1024 * 1024 }, (err, stdout, stderr) => {
      if (err && !stdout.includes('Everything is Ok')) {
        // 从 stderr 或 stdout 提取具体错误信息
        const errMsg = (stderr || stdout || '').split('\n').filter(l => l.trim()).slice(-3).join('; ');
        return reject(new Error(errMsg || `7z 退出码: ${err.code || '未知'}`));
      }
      resolve(stdout);
    });
  });
}

/** 使用 PowerShell Expand-Archive 解压 ZIP（Windows 内建，兼容性有限） */
function extractWithPowerShell(archivePath, outDir) {
  return new Promise((resolve, reject) => {
    const escapedArchive = archivePath.replace(/'/g, "''");
    const escapedOut = outDir.replace(/'/g, "''");
    runPowershell(`
      try { Expand-Archive -Path '${escapedArchive}' -DestinationPath '${escapedOut}' -Force -ErrorAction Stop; Write-Output 'OK' } catch { Write-Output ('ERR:' + $_.Exception.Message) }
    `).then(stdout => {
      const result = (stdout || '').trim();
      if (result === 'OK') return resolve(stdout);
      const errMsg = result.replace('ERR:', '');
      if (errMsg.includes('不支持的压缩方法') || errMsg.includes('unsupported compression'))
        return reject(new Error(\`ZIP 使用不支持的压缩算法（\${errMsg}），建议安装 7-Zip 后重试\`));
      reject(new Error(errMsg || 'PowerShell 解压失败'));
    }).catch(reject);
  });
}

/** 使用 Windows 内置 tar 解压 ZIP（Windows 10 17063+，兼容性优于 Expand-Archive） */
function extractWithTar(archivePath, outDir) {
  return new Promise((resolve, reject) => {
    exec(`tar -xf "${archivePath}" -C "${outDir}"`,
      { maxBuffer: 16 * 1024 * 1024 },
      (err, stdout, stderr) => {
        if (err) reject(new Error((stderr || err.message).trim()));
        else resolve(stdout);
      });
  });
}

/* ===================================================================
   WPS 操作
   =================================================================== */
function checkWPSInstalled() {
  try {
    const out = execSync(
      'powershell -NoProfile -Command "try{ $a=New-Object -ComObject Kwps.Application; $a.Quit(); Write-Output 1 }catch{ Write-Output 0 }"',
      { encoding: 'utf8', timeout: 5000 }
    );
    return out.trim() === '1';
  } catch { return false; }
}

async function convertToPDF(srcPath) {
  const ext = extname(srcPath).toLowerCase();
  const pdfPath = srcPath.slice(0, -ext.length) + '.pdf';
  let comType, format;
  if (['.doc','.docx'].includes(ext))      { comType = 'Kwps.Application'; format = 17; }
  else if (['.xls','.xlsx'].includes(ext)) { comType = 'Ket.Application';  format = 0; }
  else if (['.ppt','.pptx'].includes(ext)) { comType = 'Kwpp.Application'; format = 32; }
  else { new Notice('❌ 不支持的文件格式'); return; }
  const script = `
    try {
      $app = New-Object -ComObject ${comType};
      $app.Visible = $false;
      $doc = $app.Documents.Open('${srcPath.replace(/'/g, "''")}');
      $doc.SaveAs('${pdfPath.replace(/'/g, "''")}', ${format});
      $doc.Close();
      $app.Quit();
      Write-Output 'ok';
    } catch { Write-Output $_.Exception.Message; }
  `;
  const result = await runPowershell(script);
  if (result === 'ok') new Notice(`✅ 已转换为 PDF: ${basename(pdfPath)}`);
  else new Notice(`❌ 转换失败: ${result}`);
}

async function mergePDFs(filePaths) {
  if (filePaths.length < 2) { new Notice('⚠️ 请选择至少 2 个 PDF 文件'); return; }
  const outPath = join(dirname(filePaths[0]), '合并_' + Date.now() + '.pdf');
  const pathList = filePaths.map(p => `'${p.replace(/'/g, "''")}'`).join(', ');
  const script = `
    try {
      $app = New-Object -ComObject Kwps.Application;
      $app.Visible = $false;
      $pdf = $app.PdfMergeDocuments(@(${pathList}));
      $pdf.SaveAs('${outPath.replace(/'/g, "''")}');
      $app.Quit();
      Write-Output 'ok';
    } catch { Write-Output $_.Exception.Message; }
  `;
  const result = await runPowershell(script);
  if (result === 'ok') new Notice(`✅ PDF 已合并: ${basename(outPath)}`);
  else new Notice(`❌ 合并失败: ${result}\n提示：WPS 专业版支持 PDF 合并`);
}

async function splitPDF(filePath) {
  const ext = extname(filePath).toLowerCase();
  if (ext !== '.pdf') { new Notice('⚠️ 请选择 PDF 文件'); return; }
  const outDir = join(dirname(filePath), basename(filePath, '.pdf') + '_拆分');
  if (!existsSync(outDir)) mkdirSync(outDir, { recursive: true });
  try {
    // Try WPS COM approach
    const script = `
      try {
        $app = New-Object -ComObject Kwps.Application;
        $app.Visible = $false;
        $pdf = $app.Documents.Open('${filePath.replace(/'/g, "''")}');
        $total = $pdf.Pages.Count;
        for($i=1; $i -le $total; $i++) {
          $out = '${outDir.replace(/'/g, "''")}\\page_' + $i + '.pdf';
          $pdf.SaveAs($out, 17);  # simplified - actual splitting differs by WPS version
        }
        $pdf.Close(); $app.Quit();
        Write-Output $total;
      } catch { Write-Output 'fail'; }
    `;
    const result = await runPowershell(script);
    if (result !== 'fail') {
      new Notice(`✅ PDF 已拆分为 ${result} 页: ${basename(outDir)}`);
      return;
    }
  } catch {}
  // Fallback: use 7-Zip or other method
  new Notice('PDF 拆分需要 WPS 专业版或 Adobe Acrobat');
}

async function batchPrint(filePaths) {
  for (const fp of filePaths) {
    const ext = extname(fp).toLowerCase();
    try {
      if (['.doc','.docx'].includes(ext)) {
        await runPowershell(`
          $app = New-Object -ComObject Kwps.Application;
          $app.Visible = $false;
          $doc = $app.Documents.Open('${fp.replace(/'/g, "''")}');
          $doc.PrintOut(); $doc.Close(); $app.Quit();
        `);
      } else if (['.xls','.xlsx'].includes(ext)) {
        await runPowershell(`
          $app = New-Object -ComObject Ket.Application;
          $app.Visible = $false;
          $wb = $app.Workbooks.Open('${fp.replace(/'/g, "''")}');
          $wb.PrintOut(); $wb.Close(); $app.Quit();
        `);
      } else if (['.ppt','.pptx'].includes(ext)) {
        await runPowershell(`
          $app = New-Object -ComObject Kwpp.Application;
          $app.Visible = $false;
          $ppt = $app.Presentations.Open('${fp.replace(/'/g, "''")}');
          $ppt.PrintOut(); $ppt.Close(); $app.Quit();
        `);
      } else if (ext === '.pdf') {
        exec(`start "" /print "${fp}"`);
      }
    } catch (e) {
      new Notice(`❌ 打印失败: ${basename(fp)}`);
    }
  }
  new Notice(`🖨️ 已发送 ${filePaths.length} 个文件到打印机`);
}

async function batchConvertToPDF(filePaths) {
  let success = 0, fail = 0;
  for (const fp of filePaths) {
    try { await convertToPDF(fp); success++; }
    catch { fail++; }
  }
  new Notice(`✅ 转换完成: ${success} 成功, ${fail} 失败`);
}

/** 格式互转（Office ↔ PDF ↔ 其他） */
async function convertFormat(srcPath, targetExt) {
  const ext = extname(srcPath).toLowerCase();
  const outPath = srcPath.slice(0, -ext.length) + '.' + targetExt;
  let comType, saveFormat;
  if (['.doc','.docx'].includes(ext)) {
    comType = 'Kwps.Application';
    if (targetExt === 'pdf') saveFormat = 17;
    else if (targetExt === 'docx') saveFormat = 16;
    else if (targetExt === 'txt') saveFormat = 7;
    else { new Notice(`❌ 不支持转为 .${targetExt}`); return; }
  } else if (['.xls','.xlsx'].includes(ext)) {
    comType = 'Ket.Application';
    if (targetExt === 'pdf') saveFormat = 0;
    else if (targetExt === 'xlsx') saveFormat = 51;
    else if (targetExt === 'csv') saveFormat = 6;
    else { new Notice(`❌ 不支持转为 .${targetExt}`); return; }
  } else {
    new Notice('❌ 不支持的源格式'); return;
  }
  const script = `
    try {
      $app = New-Object -ComObject ${comType};
      $app.Visible = $false;
      $doc = $app.Documents.Open('${srcPath.replace(/'/g, "''")}');
      $doc.SaveAs('${outPath.replace(/'/g, "''")}', ${saveFormat});
      $doc.Close(); $app.Quit();
      Write-Output 'ok';
    } catch { Write-Output $_.Exception.Message; }
  `;
  const result = await runPowershell(script);
  if (result === 'ok') new Notice(`✅ 已转换: ${basename(outPath)}`);
  else new Notice(`❌ 转换失败: ${result}`);
}

/* ===================================================================
   WPS 工具弹窗（增强版）
   =================================================================== */
class WPSToolModal extends Modal {
  constructor(app, entry, view) {
    super(app);
    this.entry = entry;
    this.view = view;
    this.filePaths = view.getSelectedPaths().length > 0
      ? view.getSelectedPaths() : [entry.path];
    this.wpsInstalled = checkWPSInstalled();
    this.sevenZInstalled = is7zInstalled();
  }

  onOpen() {
    const { contentEl } = this;
    contentEl.empty();
    contentEl.addClass('de-wps-modal');
    contentEl.style.padding = '20px';
    contentEl.createEl('h2', { text: '🛠️ 文件工具' });
    contentEl.createEl('p', { text: `当前 ${this.filePaths.length} 个文件`, cls: 'de-wps-hint' });

    if (!this.wpsInstalled) {
      contentEl.createEl('p', { text: '⚠️ 未检测到 WPS Office', cls: 'de-wps-warn' });
    }

    contentEl.createEl('hr');
    const allOffice = this.filePaths.every(p => isOfficeFile(p));
    const allPDF = this.filePaths.every(p => isPDFFile(p));
    const allImage = this.filePaths.every(p => isImageFile(p));
    const oneFile = this.filePaths.length === 1;

    // === 文档转换 ===
    contentEl.createEl('h3', { text: '📄 文档转换' });
    if (allOffice) {
      this.addBtn('📄 → PDF 转换', `转 PDF`, () => { batchConvertToPDF(this.filePaths); this.close(); });
      if (oneFile) {
        this.addBtn('📄 → TXT', '转为纯文本', () => { convertFormat(this.filePaths[0], 'txt'); this.close(); });
      }
    }
    if (allPDF) {
      this.addBtn('📎 合并 PDF', `合并 ${this.filePaths.length} 个 PDF`, () => { mergePDFs(this.filePaths); this.close(); });
      if (oneFile) {
        this.addBtn('✂️ 拆分 PDF', '按页拆分', () => { splitPDF(this.filePaths[0]); this.close(); });
      }
    }
    if (allOffice && oneFile) {
      this.addBtn('🔄 DOCX ↔ PDF 互转', '格式互转', () => {
        new FormatConvertModal(this.app, this.filePaths[0]).open();
        this.close();
      });
    }

    // === 图片操作 ===
    if (allImage) {
      contentEl.createEl('hr');
      contentEl.createEl('h3', { text: '🖼️ 图片操作' });
      if (oneFile) {
        this.addBtn('👁️ 预览图片', '在新标签页打开预览', () => {
          this.view.previewFile(this.filePaths[0]);
          this.close();
        });
      }
      this.addBtn('🔄 批量旋转', '右键旋转（90°/180°）', () => {
        new Notice('批量旋转功能需要第三方工具支持');
      });
    }

    // === 压缩操作 ===
    contentEl.createEl('hr');
    contentEl.createEl('h3', { text: '🗜️ 压缩操作' });
    this.addBtn('📦 压缩为 ZIP', '使用 Windows 原生压缩', () => {
      this.view.compressToZip(this.filePaths);
      this.close();
    });
    if (this.sevenZInstalled) {
      this.addBtn('📦 压缩为 7z', '使用 7-Zip（更高压缩率）', () => {
        this.view.compressTo7z(this.filePaths);
        this.close();
      });
      this.addBtn('📦 压缩为 RAR', '使用 7-Zip 创建 RAR', () => {
        this.view.compressToRar(this.filePaths);
        this.close();
      });
    } else {
      this.addBtn('📥 安装 7-Zip', '获取更高压缩率支持', () => {
        exec('start https://7-zip.org/');
        new Notice('正在打开 7-Zip 官网');
      });
    }

    if (allPDF || allOffice) {
      contentEl.createEl('hr');
      contentEl.createEl('h3', { text: '🖨️ 批量操作' });
      this.addBtn('🖨️ 批量打印', `打印 ${this.filePaths.length} 个文件`, () => { batchPrint(this.filePaths); this.close(); });
    }

    contentEl.createEl('hr');
    new Setting(contentEl)
      .addButton(btn => btn.setButtonText('关闭').onClick(() => this.close()));
  }

  addBtn(label, desc, cb) {
    new Setting(this.contentEl).setName(label).setDesc(desc)
      .addButton(btn => btn.setButtonText('执行').onClick(cb));
  }
}

/* ===================================================================
   格式互转弹窗
   =================================================================== */
class FormatConvertModal extends Modal {
  constructor(app, filePath) {
    super(app);
    this.filePath = filePath;
  }
  onOpen() {
    const { contentEl } = this;
    contentEl.empty();
    contentEl.createEl('h2', { text: '🔄 格式互转' });
    contentEl.createEl('p', { text: `源文件: ${basename(this.filePath)}` });
    const ext = extname(this.filePath).toLowerCase();
    const targets = ext === '.pdf' ? ['docx', 'txt', 'png'] : ['pdf', 'docx', 'txt'];
    for (const t of targets) {
      new Setting(contentEl)
        .setName(`转为 .${t}`)
        .addButton(btn => btn.setButtonText('转换').onClick(() => {
          convertFormat(this.filePath, t);
          this.close();
        }));
    }
    new Setting(contentEl).addButton(btn => btn.setButtonText('关闭').onClick(() => this.close()));
  }
}

/* ===================================================================
   批量重命名弹窗
   =================================================================== */
class BatchRenameModal extends Modal {
  constructor(app, filePaths, currentDir, onComplete) {
    super(app);
    this.filePaths = filePaths;
    this.currentDir = currentDir;
    this.onComplete = onComplete;
    this.pattern = '';
    this.replacement = '';
    this.useRegex = false;
    this.previewResults = [];
  }

  onOpen() {
    const { contentEl } = this;
    contentEl.empty();
    contentEl.style.padding = '20px';
    contentEl.createEl('h2', { text: '✏️ 批量重命名' });
    contentEl.createEl('p', { text: `已选 ${this.filePaths.length} 个文件`, cls: 'de-wps-hint' });

    // 查找/替换模式
    new Setting(contentEl)
      .setName('查找')
      .addText(tc => tc.setPlaceholder('输入查找文本或正则')
        .onChange(v => { this.pattern = v; this.updatePreview(); })
        .inputEl.style.width = '100%');

    new Setting(contentEl)
      .setName('替换为')
      .addText(tc => tc.setPlaceholder('替换文本（支持 $1, $2 等）')
        .onChange(v => { this.replacement = v; this.updatePreview(); })
        .inputEl.style.width = '100%');

    new Setting(contentEl)
      .setName('使用正则表达式')
      .addToggle(tc => tc.setValue(this.useRegex)
        .onChange(v => { this.useRegex = v; this.updatePreview(); }));

    // 预设规则
    contentEl.createEl('h3', { text: '预设规则' });
    const presets = [
      { label: '序号前缀 (01_名称)', fn: () => this.applyPreset('index') },
      { label: '去除空格', fn: () => { this.pattern = '\\s+'; this.replacement = '_'; this.useRegex = true; this.updatePreview(); }},
      { label: '转小写', fn: () => {
        this.previewResults = this.filePaths.map(fp => {
          const dir = dirname(fp);
          const oldName = basename(fp);
          const newName = oldName.toLowerCase();
          return { oldPath: fp, newPath: join(dir, newName), name: oldName, newName };
        });
        this.renderPreview();
      }},
      { label: '替换日期 (YYYYMMDD)', fn: () => { const d = new Date(); this.pattern = '^'; this.replacement = `${d.getFullYear()}${String(d.getMonth()+1).padStart(2,'0')}${String(d.getDate()).padStart(2,'0')}_`; this.useRegex = false; this.updatePreview(); }},
    ];
    const presetRow = contentEl.createDiv({ cls: 'de-preset-row' });
    presetRow.style.display = 'flex';
    presetRow.style.flexWrap = 'wrap';
    presetRow.style.gap = '6px';
    presetRow.style.marginBottom = '12px';
    for (const p of presets) {
      const btn = presetRow.createEl('button', { text: p.label });
      btn.style.cssText = 'padding:4px 10px;border:1px solid var(--background-modifier-border);border-radius:4px;cursor:pointer;font-size:12px;';
      btn.onclick = p.fn;
    }

    // 预览
    this.previewEl = contentEl.createDiv({ cls: 'de-preview' });
    this.previewEl.style.maxHeight = '300px';
    this.previewEl.style.overflowY = 'auto';
    this.previewEl.style.marginBottom = '12px';
    this.previewEl.style.fontSize = '13px';
    this.updatePreview();

    // 执行按钮
    new Setting(contentEl)
      .addButton(btn => btn.setButtonText('✅ 执行重命名').setCta()
        .onClick(() => this.executeRename()))
      .addButton(btn => btn.setButtonText('取消').onClick(() => this.close()));
  }

  applyPreset(type) {
    if (type === 'index') {
      const digits = String(this.filePaths.length).length;
      this.filePaths.sort((a, b) => a.localeCompare(b));
      this.previewResults = this.filePaths.map((fp, i) => {
        const dir = dirname(fp);
        const ext = extname(fp);
        const idx = String(i + 1).padStart(digits, '0');
        const newName = `${idx}_${basename(fp)}`;
        return { oldPath: fp, newPath: join(dir, newName), name: basename(fp), newName };
      });
      this.renderPreview();
      return;
    }
    this.updatePreview();
  }

  updatePreview() {
    if (!this.pattern) {
      this.previewResults = this.filePaths.map(fp => ({
        oldPath: fp, newPath: fp, name: basename(fp), newName: basename(fp)
      }));
      this.renderPreview();
      return;
    }
    this.previewResults = this.filePaths.map(fp => {
      const dir = dirname(fp);
      const oldName = basename(fp);
      let newName;
      try {
        if (this.useRegex) {
          const regex = new RegExp(this.pattern, 'g');
          newName = oldName.replace(regex, this.replacement);
        } else {
          newName = oldName.split(this.pattern).join(this.replacement);
        }
      } catch (e) {
        newName = oldName + ' [正则错误]';
      }
      return { oldPath: fp, newPath: join(dir, newName), name: oldName, newName };
    });
    this.renderPreview();
  }

  renderPreview() {
    this.previewEl.empty();
    if (this.previewResults.length === 0) return;
    const table = this.previewEl.createEl('table');
    table.style.width = '100%';
    table.style.borderCollapse = 'collapse';
    const thead = table.createEl('thead');
    const hrow = thead.createEl('tr');
    hrow.createEl('th', { text: '原名称' }).style.cssText = 'text-align:left;padding:4px 8px;border-bottom:1px solid var(--background-modifier-border);';
    hrow.createEl('th', { text: '新名称' }).style.cssText = 'text-align:left;padding:4px 8px;border-bottom:1px solid var(--background-modifier-border);';
    const tbody = table.createEl('tbody');
    for (const r of this.previewResults) {
      const changed = r.name !== r.newName && r.newName && !r.newName.includes('[正则错误]');
      const row = tbody.createEl('tr');
      row.createEl('td', { text: r.name }).style.cssText = 'padding:2px 8px;';
      const td2 = row.createEl('td', { text: r.newName || '(空)' });
      td2.style.cssText = `padding:2px 8px;${changed ? 'color:var(--text-accent);font-weight:bold;' : ''}`;
    }
  }

  executeRename() {
    let success = 0, fail = 0;
    for (const r of this.previewResults) {
      if (r.name === r.newName || !r.newName || r.newName.includes('[正则错误]')) continue;
      try {
        if (existsSync(r.newPath)) { fail++; continue; }
        renameSync(r.oldPath, r.newPath);
        success++;
      } catch (e) { fail++; }
    }
    new Notice(`✅ 重命名完成: ${success} 成功, ${fail} 失败`);
    if (this.onComplete) this.onComplete();
    this.close();
  }
}

/* ===================================================================
   重命名弹窗（替代 prompt()，Electron 中 prompt 被阻断）
   =================================================================== */
class RenameModal extends Modal {
  constructor(app, oldName, oldPath, currentDir, onComplete) {
    super(app);
    this.oldName = oldName;
    this.oldPath = oldPath;
    this.onComplete = onComplete;
    this._done = false;
  }

  onOpen() {
    const { contentEl } = this;
    contentEl.empty();
    contentEl.style.padding = '20px';
    contentEl.createEl('h2', { text: '✏️ 重命名' });

    // 纯 DOM，不用 Obsidian Setting（Setting 有内部 Enter 拦截器会先触发 close）
    const row = contentEl.createDiv({ cls: 'setting-item' });
    const info = row.createDiv({ cls: 'setting-item-info' });
    info.createDiv({ cls: 'setting-item-name', text: '新名称' });
    const ctrl = row.createDiv({ cls: 'setting-item-control' });
    const inputEl = ctrl.createEl('input', { type: 'text', cls: 'de-rename-input' });
    inputEl.value = this.oldName;
    inputEl.style.cssText = 'width:100%;padding:6px 8px;border:1px solid var(--background-modifier-border);border-radius:4px;';

    // 单击选中主文件名（不含扩展名）
    const ext = extname(this.oldName);
    const stem = this.oldName.slice(0, -ext.length);
    inputEl.addEventListener('click', () => inputEl.setSelectionRange(0, stem.length));
    // Enter 确认 / Escape 取消
    inputEl.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') { e.preventDefault(); this.doRename(inputEl.value); }
      if (e.key === 'Escape') { e.preventDefault(); this.close(); }
    });
    // 自动聚焦（用 requestAnimationFrame 等 Modal 完全渲染完成）
    requestAnimationFrame(() => {
      requestAnimationFrame(() => { inputEl.focus(); inputEl.setSelectionRange(0, stem.length); });
    });

    // 按钮行
    const btnRow = contentEl.createDiv({ cls: 'setting-item' });
    const btnCtrl = btnRow.createDiv({ cls: 'setting-item-control' });
    btnCtrl.style.justifyContent = 'flex-start';
    const confirmBtn = btnCtrl.createEl('button', { text: '✅ 确认', cls: 'mod-cta' });
    confirmBtn.style.marginRight = '8px';
    confirmBtn.addEventListener('click', () => this.doRename(inputEl.value));
    btnCtrl.createEl('button', { text: '取消' }).addEventListener('click', () => this.close());
  }

  doRename(newName) {
    if (this._done) return;
    const trimmed = (newName || '').trim();
    if (!trimmed || trimmed === this.oldName) {
      new Notice('⚠️ 名称未变化或无效');
      return;
    }
    const newPath = join(dirname(this.oldPath), trimmed);
    try {
      renameSync(this.oldPath, newPath);
      this._done = true;
      new Notice(`✏️ 已重命名为: ${trimmed}`);
      if (this.onComplete) this.onComplete(newPath);
      this.close();
    } catch (e) {
      new Notice(`❌ 重命名失败: ${e.message}`);
    }
  }

  onClose() {
    this.contentEl.empty();
  }
}

/* ===================================================================
   文件预览弹窗/侧边视图
   =================================================================== */
class FilePreviewModal extends Modal {
  constructor(app, filePath) {
    super(app);
    this.filePath = filePath;
  }

  onOpen() {
    const { contentEl } = this;
    contentEl.empty();
    contentEl.style.padding = '20px';
    const name = basename(this.filePath);
    contentEl.createEl('h2', { text: `👁️ ${name}` });

    const ext = extname(this.filePath).toLowerCase();
    const contentArea = contentEl.createDiv();
    contentArea.style.cssText = 'max-height:70vh;overflow:auto;';

    try {
      if (isImageFile(name)) {
        // 图片预览
        const fs = require('fs');
        const data = fs.readFileSync(this.filePath);
        const base64 = data.toString('base64');
        const mimeTypes = { '.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg',
          '.gif': 'image/gif', '.bmp': 'image/bmp', '.webp': 'image/webp', '.svg': 'image/svg+xml' };
        const img = contentArea.createEl('img');
        img.src = `data:${mimeTypes[ext] || 'image/png'};base64,${base64}`;
        img.style.maxWidth = '100%';
        img.style.maxHeight = '65vh';
        img.style.objectFit = 'contain';
      } else if (ext === '.pdf') {
        // PDF 预览 — 使用 iframe
        contentArea.createEl('iframe', { attr: { src: `file:///${this.filePath.replace(/\\/g, '/')}` } })
          .style.cssText = 'width:100%;height:70vh;border:none;';
      } else if (isTextFile(name)) {
        // 文本文件预览
        const text = readFileSync(this.filePath, 'utf-8');
        const pre = contentArea.createEl('pre');
        pre.textContent = text.slice(0, 50000); // 限制50K字符
        pre.style.cssText = 'white-space:pre-wrap;word-break:break-all;font-size:12px;max-height:65vh;overflow:auto;background:var(--background-primary-alt);padding:12px;border-radius:4px;';
        if (text.length > 50000) {
          contentArea.createEl('p', { text: '⚠️ 文件过大，仅显示前 50,000 字符', cls: 'de-wps-warn' });
        }
      } else {
        contentArea.createEl('p', { text: `⚠️ 不支持预览此文件类型 (${ext})\n请在系统中打开查看。` });
      }
    } catch (e) {
      contentArea.createEl('p', { text: `❌ 预览失败: ${e.message}`, cls: 'de-err' });
    }

    new Setting(contentEl)
      .addButton(btn => btn.setButtonText('关闭').onClick(() => this.close()));
  }
}

/* ===================================================================
   粘贴选择弹窗（多点剪贴板）
   =================================================================== */
class PasteSelectModal extends Modal {
  constructor(app, currentDir, view) {
    super(app);
    this.currentDir = currentDir;
    this.view = view;
  }

  onOpen() {
    const { contentEl } = this;
    contentEl.empty();
    contentEl.style.padding = '20px';
    contentEl.createEl('h2', { text: '📋 选择要粘贴的内容' });

    if (fileClipboard.items.length === 0) {
      contentEl.createEl('p', { text: '剪贴板为空' });
      new Setting(contentEl).addButton(btn => btn.setButtonText('关闭').onClick(() => this.close()));
      return;
    }

    for (let i = 0; i < fileClipboard.items.length; i++) {
      const item = fileClipboard.items[i];
      const time = new Date(item.timestamp).toLocaleTimeString();
      const actionLabel = item.action === 'cut' ? '✂️ 剪切' : '📋 复制';
      const names = item.sources.slice(0, 3).map(p => basename(p)).join(', ');
      const more = item.sources.length > 3 ? `…等${item.sources.length}项` : '';

      new Setting(contentEl)
        .setName(`${actionLabel} ${time}`)
        .setDesc(`${names}${more}`)
        .addButton(btn => btn.setButtonText('📝 粘贴').onClick(() => {
          this.pasteItem(item);
          this.close();
        }))
        .addButton(btn => btn.setButtonText('🗑️ 删除').onClick(() => {
          fileClipboard.remove(i);
          this.onOpen();
        }));
    }

    new Setting(contentEl)
      .addButton(btn => btn.setButtonText('清空剪贴板').onClick(() => {
        fileClipboard.clear();
        this.onOpen();
      }))
      .addButton(btn => btn.setButtonText('关闭').onClick(() => this.close()));
  }

  pasteItem(item) {
    let copied = 0;
    let errors = 0;
    for (const src of item.sources) {
      try {
        const dest = safeDestPath(src, this.currentDir);
        if (statSync(src).isDirectory()) { copyDirRecursive(src, dest); }
        else { copyFileSync(src, dest); }
        if (item.action === 'cut') {
          if (statSync(src).isDirectory()) removeDirRecursive(src);
          else unlinkSync(src);
        }
        copied++;
      } catch (e) {
        errors++;
        new Notice(`❌ 操作失败: ${basename(src)} - ${e.message}`);
      }
    }
    // 只有全部成功才从剪贴板移除剪切条目，防止批量失败时丢失源文件
    if (item.action === 'cut' && errors === 0) {
      fileClipboard.remove(fileClipboard.items.indexOf(item));
    }
    new Notice(`✅ 粘贴完成: ${copied} 成功${errors > 0 ? `, ${errors} 失败` : ''}`);
    this.view.render();
  }
}

/* ===================================================================
   DriveExplorerView — 文件浏览器
   =================================================================== */
class DriveExplorerView extends ItemView {
  constructor(leaf) {
    super(leaf);
    this.currentPath = null;
    this.showingDriveSelector = true;
    this.drives = [];
    this.history = [];
    this.historyIndex = -1;
    this.selectedPaths = new Set();
    this.renameTarget = null;
    this.slideDir = null;
    this.searchQuery = '';
    this.showPreview = false;
    this.previewFilePath = null;
    this.thumbnails = new Map(); // path -> base64 thumbnail cache
  }

  getViewType() { return VIEW_TYPE; }
  getIcon() { return 'folder'; }

  getDisplayText() {
    if (this.showingDriveSelector) return '盘符选择';
    if (/^[A-Za-z]:\\$/.test(this.currentPath)) {
      return this.currentPath.charAt(0).toUpperCase() + '盘';
    }
    return basename(this.currentPath) || this.currentPath;
  }

  updateTabTitle() {
    if (this.leaf && this.leaf.tabHeaderInnerTitleEl) {
      this.leaf.tabHeaderInnerTitleEl.textContent = this.getDisplayText();
    }
  }

  getSelectedPaths() { return [...this.selectedPaths]; }

  async onOpen() {
    this.containerEl.setAttribute('tabindex', '0');
    this.containerEl.style.outline = 'none';
    this.containerEl.addEventListener('click', (e) => {
      if (e.target.tagName !== 'INPUT' && e.target.tagName !== 'TEXTAREA') {
        this.containerEl.focus();
      }
    });
    // 空白处右键：显示新建菜单（注册在容器上，确保覆盖整个视图区域）
    this.containerEl.addEventListener('contextmenu', (e) => {
      if (!e.target.closest('.de-item')) {
        e.preventDefault();
        this.showNewMenu(e);
      }
    });
    this.registerMouseButtons();
    this.registerKeyboardShortcuts();
    // 用 Obsidian Scope 注册 F2（比 DOM keydown 优先级更高，能拦截 Obsidian 内置命令）
    if (this.scope) {
      this.scope.register([], 'F2', (evt) => {
        if (this.selectedPaths.size > 0) {
          evt.preventDefault();
          const targetPath = [...this.selectedPaths][0];
          this.renameItem({ name: basename(targetPath), path: targetPath });
          return false; // 告诉 Obsidian：已处理，停止传播
        }
      });
    }
    this.drives = getAvailableDrives();
    this.render();
    this.containerEl.focus();
  }

  registerMouseButtons() {
    this.containerEl.addEventListener('mouseup', (e) => {
      if (e.button === 3) { this.goBack(); e.preventDefault(); }
      if (e.button === 4) { this.goForward(); e.preventDefault(); }
    });
  }

  registerKeyboardShortcuts() {
    this.containerEl.addEventListener('keydown', (e) => {
      if (e.altKey && e.key === 'ArrowLeft') { e.preventDefault(); this.goBack(); return; }
      if (e.altKey && e.key === 'ArrowRight') { e.preventDefault(); this.goForward(); return; }
      if ((e.ctrlKey || e.metaKey) && e.code === 'KeyC' && this.selectedPaths.size > 0) {
        e.preventDefault();
        if (e.shiftKey) { this.copyFilePath({ path: [...this.selectedPaths][0] }); }
        else { this.copyFiles({ path: [...this.selectedPaths][0] }); }
        return;
      }
      if ((e.ctrlKey || e.metaKey) && e.key === 'x' && this.selectedPaths.size > 0) {
        e.preventDefault(); this.cutFiles({ path: [...this.selectedPaths][0] }); return;
      }
      if ((e.ctrlKey || e.metaKey) && e.key === 'v') {
        e.preventDefault();
        if (fileClipboard.items.length === 1) { this.pasteFiles(); }
        else { this.showPasteMenu(); }
        return;
      }
      if ((e.ctrlKey || e.metaKey) && e.key === 'f') {
        e.preventDefault();
        const searchInput = this.containerEl.querySelector('.de-search-input');
        if (searchInput) searchInput.focus();
        return;
      }
      if (e.key === 'F5') {
        e.preventDefault(); this.render(); return;
      }
      if ((e.key === 'Delete' || e.key === 'Del') && this.selectedPaths.size > 0) {
        e.preventDefault();
        this.deleteToRecycle({ path: [...this.selectedPaths][0] });
        return;
      }
    });
  }

  /* =============================================================== */
  /*  渲染                                                           */
  /* =============================================================== */
  render() {
    this.saveScrollPos();
    const { containerEl } = this;
    containerEl.empty();
    if (this.showingDriveSelector) { this.renderDriveSelector(containerEl); return; }
    this.renderExplorer(containerEl);
    this.restoreScrollPos();
  }

  /** 保存并恢复滚动位置 */
  saveScrollPos() {
    const el = this.containerEl.querySelector('.de-list-container') || this.containerEl.querySelector('.de-list');
    if (el) this._savedScrollTop = el.scrollTop;
  }

  restoreScrollPos() {
    if (this._savedScrollTop !== undefined) {
      requestAnimationFrame(() => {
        const el = this.containerEl.querySelector('.de-list-container') || this.containerEl.querySelector('.de-list');
        if (el) el.scrollTop = this._savedScrollTop || 0;
        this._savedScrollTop = undefined;
      });
    }
  }

  /** 盘符选择界面 */
  renderDriveSelector(container) {
    container.addClass('de-drive-selector');
    container.style.padding = '20px';

    container.createEl('h2', { text: '💾 选择盘符' });
    container.createEl('p', { text: '选择要浏览的磁盘或输入路径', cls: 'de-wps-hint' });

    // 自定义路径输入
    const inputRow = container.createDiv({ cls: 'de-path-input-row' });
    inputRow.style.display = 'flex';
    inputRow.style.gap = '8px';
    inputRow.style.marginBottom = '16px';
    const pathInput = inputRow.createEl('input', { cls: 'de-search-input', attr: { type: 'text', placeholder: '输入路径（如 D:\\ 或 \\\\server\\share）' } });
    pathInput.style.flex = '1';
    pathInput.style.padding = '8px';
    pathInput.style.borderRadius = '4px';
    pathInput.style.border = '1px solid var(--background-modifier-border)';
    const goBtn = inputRow.createEl('button', { text: '打开' });
    goBtn.style.cssText = 'padding:8px 16px;border-radius:4px;border:1px solid var(--background-modifier-border);cursor:pointer;';
    goBtn.onclick = () => {
      const input = pathInput.value.trim();
      if (input) this.openPath(input);
    };
    pathInput.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') goBtn.click();
    });

    // 盘符网格
    const grid = container.createDiv({ cls: 'de-drive-grid' });
    grid.style.display = 'grid';
    grid.style.gridTemplateColumns = 'repeat(auto-fill, minmax(200px, 1fr))';
    grid.style.gap = '10px';

    for (const drive of this.drives) {
      const card = grid.createDiv({ cls: 'de-drive-card' });
      card.style.cssText = 'padding:16px;border:1px solid var(--background-modifier-border);border-radius:8px;cursor:pointer;transition:all 0.15s;';
      card.onmouseenter = () => { card.style.borderColor = 'var(--text-accent)'; card.style.background = 'var(--background-modifier-hover)'; };
      card.onmouseleave = () => { card.style.borderColor = 'var(--background-modifier-border)'; card.style.background = 'transparent'; };
      card.onclick = () => this.openPath(drive.path);

      const icon = ['本地磁盘', '本地固定磁盘'].includes(drive.type) ? '💽' :
                   drive.type.includes('可移动') ? '💾' :
                   drive.type.includes('网络') ? '🌐' :
                   drive.type.includes('光盘') ? '💿' : '📁';
      card.createEl('div', { text: icon }).style.fontSize = '32px';
      card.createEl('div', { text: drive.path, cls: 'de-drive-letter' }).style.cssText = 'font-size:18px;font-weight:bold;margin:4px 0;';
      const label = drive.label || drive.type || '';
      card.createEl('div', { text: label, cls: 'de-drive-label' }).style.cssText = 'font-size:13px;color:var(--text-muted);';
    }

    // 刷新盘符按钮
    const refreshRow = container.createDiv();
    refreshRow.style.marginTop = '16px';
    const refreshBtn = refreshRow.createEl('button', { text: '🔄 刷新盘符' });
    refreshBtn.style.cssText = 'padding:6px 14px;border:1px solid var(--background-modifier-border);border-radius:4px;cursor:pointer;';
    refreshBtn.onclick = () => {
      this.drives = getAvailableDrives();
      this.render();
      new Notice('✅ 已刷新盘符列表');
    };

    this.updateTabTitle();
  }

  renderExplorer(container) {
    // ---- 导航栏 ----
    const header = container.createDiv({ cls: 'de-header' });
    const navRow = header.createDiv({ cls: 'de-nav-row' });

    const backBtn = navRow.createEl('button', { cls: 'de-nav-btn', text: '◀' });
    backBtn.title = '后退';
    backBtn.disabled = this.historyIndex <= 0;
    backBtn.addEventListener('click', () => this.goBack());

    const fwdBtn = navRow.createEl('button', { cls: 'de-nav-btn', text: '▶' });
    fwdBtn.title = '前进';
    fwdBtn.disabled = this.historyIndex >= this.history.length - 1;
    fwdBtn.addEventListener('click', () => this.goForward());

    const upBtn = navRow.createEl('button', { cls: 'de-nav-btn', text: '⬆' });
    upBtn.title = '上级目录';
    upBtn.addEventListener('click', () => this.goUp());

    const homeBtn = navRow.createEl('button', { cls: 'de-nav-btn', text: '🏠' });
    homeBtn.title = '盘符选择';
    homeBtn.addEventListener('click', () => {
      this.showingDriveSelector = true;
      this.currentPath = null;
      this.selectedPaths.clear();
      this.render();
    });

    // 搜索框
    const searchInput = navRow.createEl('input', { cls: 'de-search-input', attr: { type: 'text', placeholder: '🔍 搜索文件…' } });
    searchInput.style.cssText = 'flex:1;padding:4px 8px;border:1px solid var(--background-modifier-border);border-radius:4px;font-size:12px;min-width:60px;';
    searchInput.value = this.searchQuery;
    searchInput.addEventListener('input', (e) => {
      this.searchQuery = e.target.value;
      this.renderFileList(this.containerEl.querySelector('.de-list'), true);
    });

    header.createDiv({ cls: 'de-path', text: this.currentPath || '' });

    // 操作按钮行
    const actionRow = header.createDiv({ cls: 'de-header-actions' });
    actionRow.style.display = 'flex';
    actionRow.style.gap = '6px';
    actionRow.style.marginTop = '6px';
    actionRow.style.flexWrap = 'wrap';

    const refreshBtn = actionRow.createEl('button', { cls: 'de-nav-btn', text: '🔄 刷新' });
    refreshBtn.addEventListener('click', () => this.render());

    const previewBtn = actionRow.createEl('button', { cls: 'de-nav-btn', text: this.showPreview ? '📖 隐藏预览' : '📖 预览面板' });
    previewBtn.addEventListener('click', () => {
      this.showPreview = !this.showPreview;
      this.render();
    });

    if (this.selectedPaths.size > 0) {
      const selInfo = actionRow.createSpan({ cls: 'de-sel-info', text: `已选 ${this.selectedPaths.size} 项` });
      selInfo.style.marginLeft = '8px';
    }

    // ---- 主区域：文件列表 + 预览面板 ----
    const mainArea = container.createDiv({ cls: 'de-main-area' });
    mainArea.style.display = 'flex';
    mainArea.style.height = 'calc(100% - 120px)';

    // 文件列表
    const listContainer = mainArea.createDiv({ cls: 'de-list-container' });
    listContainer.style.flex = this.showPreview ? '1' : '1';
    listContainer.style.overflow = 'auto';

    const list = listContainer.createDiv({ cls: 'de-list' });
    if (this.slideDir === 'left')  list.addClass('de-slide-left');
    if (this.slideDir === 'right') list.addClass('de-slide-right');
    this.slideDir = null;

    list.addEventListener('click', (e) => {
      if (e.target === list || e.target === listContainer) {
        this.selectedPaths.clear();
        this.previewFilePath = null;
        this.render();
      }
    });
    // 空白区拖放：放入当前目录
    list.addEventListener('dragover', (e) => {
      if (!e.target.closest('.de-item')) {
        e.preventDefault();
        e.dataTransfer.dropEffect = 'move';
      }
    });
    list.addEventListener('drop', (e) => {
      if (e.target.closest('.de-item')) return;
      e.preventDefault();
      this.dropFilesInto(this.currentPath, e.ctrlKey || e.metaKey);
    });

    this.renderFileList(list);

    // 预览面板
    if (this.showPreview && this.previewFilePath) {
      const previewPanel = mainArea.createDiv({ cls: 'de-preview-panel' });
      previewPanel.style.cssText = 'width:300px;min-width:200px;border-left:1px solid var(--background-modifier-border);padding:8px;overflow:auto;';
      this.renderPreviewPanel(previewPanel, this.previewFilePath);
    }

    this.updateTabTitle();
  }

  renderFileList(container, keepSearch = false) {
    container.empty();
    if (!this.currentPath || !existsSync(this.currentPath)) {
      container.createDiv({ cls: 'de-msg de-err', text: '❌ 路径不存在' });
      return;
    }
    let names;
    try { names = readdirSync(this.currentPath, { encoding: 'utf8' }); }
    catch (err) {
      container.createDiv({ cls: 'de-msg de-err', text: `⚠️ 读取失败: ${err.message}` });
      return;
    }

    // 搜索过滤
    if (this.searchQuery) {
      const q = this.searchQuery.toLowerCase();
      names = names.filter(n => n.toLowerCase().includes(q));
    }

    const entries = [];
    for (const name of names) {
      try {
        const fp = join(this.currentPath, name);
        const s = statSync(fp);
        entries.push({ name, path: fp, isDir: s.isDirectory(), size: s.size, mtime: s.mtime });
      } catch (_) {}
    }
    if (entries.length === 0) {
      container.createDiv({ cls: 'de-msg', text: this.searchQuery ? '🔍 无匹配结果' : '📂 空文件夹' });
      return;
    }

    entries.sort((a, b) => {
      if (a.isDir !== b.isDir) return a.isDir ? -1 : 1;
      return a.name.localeCompare(b.name, 'zh-CN');
    });

    for (const entry of entries) this.renderEntry(container, entry);
  }

  renderEntry(container, entry) {
    const isSelected = this.selectedPaths.has(entry.path);
    const item = container.createDiv({ cls: `de-item${isSelected ? ' selected' : ''}` });
    item.dataset.path = entry.path;
    item.draggable = true;

    // 图标（支持缩略图）
    if (isImageFile(entry.name) && entry.size < 5 * 1024 * 1024) {
      // 生成缩略图
      try {
        const data = readFileSync(entry.path);
        const b64 = data.toString('base64');
        const ext = extname(entry.name).toLowerCase();
        const mime = { '.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg',
          '.gif': 'image/gif', '.bmp': 'image/bmp', '.webp': 'image/webp', '.svg': 'image/svg+xml' };
        const thumb = item.createEl('img', { cls: 'de-thumb' });
        thumb.src = `data:${mime[ext] || 'image/png'};base64,${b64}`;
        thumb.style.cssText = 'width:32px;height:32px;object-fit:cover;border-radius:3px;flex-shrink:0;';
      } catch { item.createSpan({ cls: 'de-icon', text: '🖼️' }); }
    } else {
      item.createSpan({ cls: 'de-icon', text: entry.isDir ? '📁' : this.getIcon(entry.name) });
    }

    const nameSpan = item.createSpan({ cls: 'de-name', text: entry.name });

    // 拖拽 — 兼容 Obsidian 编辑器和外部窗口（如聊天对话）
    item.addEventListener('dragstart', (e) => {
      const path = entry.path;
      const name = entry.name;
      // text/plain 放第一位，兼容性最广（对话框/编辑器优先读此格式）
      e.dataTransfer.setData('text/plain', path);
      // 标准 file:// URI（text/uri-list 规范要求每行以 CRLF 结尾）
      const fileUri = 'file:///' + path.replace(/\\/g, '/');
      e.dataTransfer.setData('text/uri-list', fileUri + '\r\n');
      // HTML 超链接格式（支持富文本编辑器和浏览器）
      e.dataTransfer.setData('text/html',
        `<a href="${fileUri}">${name}</a>`);
      e.dataTransfer.effectAllowed = 'all';
      // 自定义拖拽图像（使用文件名）
      if (e.dataTransfer.setDragImage) {
        const canvas = document.createElement('canvas');
        canvas.width = 240; canvas.height = 32;
        const ctx = canvas.getContext('2d');
        ctx.fillStyle = 'var(--interactive-accent, #555)';
        ctx.roundRect?.(0, 0, canvas.width, canvas.height, 4);
        ctx.fill();
        ctx.fillStyle = '#fff';
        ctx.font = '13px sans-serif';
        ctx.fillText('📄 ' + name.slice(0, 35), 12, 22);
        e.dataTransfer.setDragImage(canvas, 0, 0);
      }
      // 记录拖拽的源路径（供内部文件夹拖放使用）
      this._dragSourcePaths = this.selectedPaths.size > 0
        ? new Set(this.selectedPaths) : new Set([entry.path]);
      item.classList.add('de-dragging');
    });
    item.addEventListener('dragend', () => {
      item.classList.remove('de-dragging');
      this._dragSourcePaths = null;
    });

    // 文件夹拖放目标：拖文件/文件夹到文件夹上时自动移入
    if (entry.isDir) {
      item.addEventListener('dragover', (e) => {
        e.preventDefault();
        e.dataTransfer.dropEffect = 'move';
        item.classList.add('de-drag-over');
      });
      item.addEventListener('dragleave', () => {
        item.classList.remove('de-drag-over');
      });
      item.addEventListener('drop', (e) => {
        e.preventDefault();
        item.classList.remove('de-drag-over');
        const isCopy = e.ctrlKey || e.metaKey;
        this.dropFilesInto(entry.path, isCopy);
      });
    }

    // 单击
    item.addEventListener('click', (e) => {
      if (e.ctrlKey || e.metaKey) {
        if (this.selectedPaths.has(entry.path)) this.selectedPaths.delete(entry.path);
        else this.selectedPaths.add(entry.path);
      } else {
        this.selectedPaths.clear();
        this.selectedPaths.add(entry.path);
        this.previewFilePath = entry.path;
      }
      this.render();
    });

    // 双击
    item.addEventListener('dblclick', () => {
      if (entry.isDir) { this.navigateTo(entry.path); }
      else { this.openFile(entry); }
    });

    // 右键
    item.addEventListener('contextmenu', (e) => {
      e.preventDefault();
      if (!this.selectedPaths.has(entry.path)) {
        this.selectedPaths.clear();
        this.selectedPaths.add(entry.path);
        this.render();
      }
      this.showContextMenu(e, entry);
    });
  }

  getIcon(name) {
    const map = {
      '.md': '📝',  '.txt': '📄',  '.pdf': '📕',
      '.doc': '📘', '.docx': '📘',
      '.xls': '📊', '.xlsx': '📊',
      '.ppt': '📙', '.pptx': '📙',
      '.jpg': '🖼️','.jpeg': '🖼️','.png': '🖼️','.gif': '🖼️','.svg': '🖼️','.bmp':'🖼️',
      '.mp4': '🎬',  '.avi': '🎬',  '.mkv': '🎬',  '.mov': '🎬',
      '.mp3': '🎵',  '.wav': '🎵',  '.flac':'🎵',
      '.zip': '📦',  '.rar': '📦',  '.7z': '📦',
      '.exe': '⚙️',  '.msi': '⚙️',
      '.html':'🌐',  '.css': '🎨',  '.js': '🟨',  '.py': '🐍',  '.json':'📋',
      '.torrent':'🧲',  '.dll':'🔧',
    };
    return map[extname(name).toLowerCase()] || '📄';
  }

  /** 预览面板 */
  renderPreviewPanel(container, filePath) {
    if (!filePath || !existsSync(filePath)) {
      container.createEl('p', { text: '选择文件以预览', cls: 'de-wps-hint' });
      return;
    }
    const name = basename(filePath);
    container.createEl('div', { text: `📄 ${name}`, cls: 'de-preview-title' }).style.cssText = 'font-weight:bold;margin-bottom:8px;font-size:14px;';
    const ext = extname(filePath).toLowerCase();

    try {
      if (isImageFile(name)) {
        const data = readFileSync(filePath);
        const b64 = data.toString('base64');
        const mimeTypes = { '.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg',
          '.gif': 'image/gif', '.bmp': 'image/bmp', '.webp': 'image/webp', '.svg': 'image/svg+xml' };
        const img = container.createEl('img');
        img.src = `data:${mimeTypes[ext] || 'image/png'};base64,${b64}`;
        img.style.cssText = 'max-width:100%;max-height:300px;object-fit:contain;border-radius:4px;';
      } else if (isTextFile(name)) {
        const text = readFileSync(filePath, 'utf-8');
        const pre = container.createEl('pre');
        pre.textContent = text.slice(0, 2000);
        pre.style.cssText = 'white-space:pre-wrap;word-break:break-all;font-size:11px;max-height:300px;overflow:auto;';
      } else {
        const s = statSync(filePath);
        container.createEl('p', { text: `📏 大小: ${formatSize(s.size)}`, cls: 'de-wps-hint' });
        container.createEl('p', { text: `📅 修改: ${s.mtime.toLocaleString()}`, cls: 'de-wps-hint' });
        container.createEl('p', { text: `🔤 类型: ${ext.toUpperCase() || '未知'}`, cls: 'de-wps-hint' });
      }
    } catch (e) {
      container.createEl('p', { text: `❌ ${e.message}`, cls: 'de-err' });
    }
  }

  /* =============================================================== */
  /*  导航                                                           */
  /* =============================================================== */
  openPath(path) {
    path = path.replace(/^"(.*)"$/, '$1').trim();
    if (!existsSync(path)) { new Notice(`❌ 路径不存在: ${path}`); return; }
    this.showingDriveSelector = false;
    this.currentPath = path;
    this.history = [path];
    this.historyIndex = 0;
    this.selectedPaths.clear();
    this.previewFilePath = null;
    new Notice(`📂 已打开: ${path}`);
    this.render();
  }

  navigateTo(path) {
    this.slideDir = 'right';
    this.history = this.history.slice(0, this.historyIndex + 1);
    this.history.push(path);
    this.historyIndex = this.history.length - 1;
    this.currentPath = path;
    this.selectedPaths.clear();
    this.previewFilePath = null;
    this._dragSourcePaths = null; // 导航时清理拖拽状态
    this.render();
  }

  goUp() {
    const p = dirname(this.currentPath);
    if (p !== this.currentPath) {
      this.slideDir = 'left';
      this.history = this.history.slice(0, this.historyIndex + 1);
      this.history.push(p);
      this.historyIndex = this.history.length - 1;
      this.currentPath = p;
      this.selectedPaths.clear();
      this.previewFilePath = null;
      this.render();
    }
  }

  goBack() {
    if (this.historyIndex > 0) {
      this.slideDir = 'left';
      this.historyIndex--;
      this.currentPath = this.history[this.historyIndex];
      this.selectedPaths.clear();
      this.previewFilePath = null;
      this._dragSourcePaths = null;
      this.render();
    }
  }

  goForward() {
    if (this.historyIndex < this.history.length - 1) {
      this.slideDir = 'right';
      this.historyIndex++;
      this.currentPath = this.history[this.historyIndex];
      this.selectedPaths.clear();
      this.previewFilePath = null;
      this._dragSourcePaths = null;
      this.render();
    }
  }

  /* =============================================================== */
  /*  文件操作                                                       */
  /* =============================================================== */
  openFile(entry) {
    // .md 文件在 Obsidian 内打开
    if (extname(entry.name).toLowerCase() === '.md') {
      try {
        // 如果文件在 vault 内，直接用 Obsidian API
        const vaultPath = this.app.vault.adapter.getBasePath?.() || '';
        if (vaultPath && entry.path.startsWith(vaultPath)) {
          const relativePath = entry.path.slice(vaultPath.length).replace(/^[/\\]/, '').replace(/\\/g, '/');
          const file = this.app.vault.getAbstractFileByPath(relativePath);
          if (file instanceof TFile) {
            this.app.workspace.getLeaf('tab').openFile(file);
            return;
          }
        }
        // 文件在 vault 外，使用额外命令打开
        const leaf = this.app.workspace.getLeaf('tab');
        // 使用 Obsidian 的 markdown 阅读能力：通过创建临时 vault 链接
        // 最佳方式：用 child_process 调用 obsidian 或使用系统打开
        // 这里使用 Obsidian 的 MarkdownView
        const content = readFileSync(entry.path, 'utf-8');
        leaf.setEphemeralState?.({ source: content });
        // 直接在新标签页显示内容
        leaf.setViewState({
          type: 'markdown',
          state: { content, basePath: entry.path }
        });
        this.app.workspace.revealLeaf(leaf);
        return;
      } catch (e) {
        // Fallback: 系统打开
        try { exec(`start "" "${entry.path}"`); } catch {}
        return;
      }
    }

    // 其他文件: 系统打开
    try {
      new Notice(`📂 正在打开: ${entry.name}`);
      exec(`start "" "${entry.path}"`, (error) => {
        if (error) new Notice(`❌ 打开失败: ${error.message}`);
      });
    } catch (e) {
      new Notice(`❌ 无法打开文件: ${e.message}`);
    }
  }

  /** 预览文件（在新弹窗中） */
  previewFile(filePath) {
    new FilePreviewModal(this.app, filePath).open();
  }

  getTargetPaths(entry) {
    if (this.selectedPaths.size > 0) return [...this.selectedPaths];
    return [entry.path];
  }

  /* =============================================================== */
  /*  右键菜单                                                       */
  /* =============================================================== */
  showContextMenu(event, entry) {
    const menu = new Menu();
    const targets = this.getTargetPaths(entry);
    const isSingle = targets.length === 1;
    const ext = extname(entry.name).toLowerCase();

    // 复制/剪切/粘贴
    menu.addItem(i => i.setTitle('📋 复制').setIcon('copy').onClick(() => this.copyFiles(entry)));
    menu.addItem(i => i.setTitle('✂️ 剪切').setIcon('scissors').onClick(() => this.cutFiles(entry)));
    if (fileClipboard.items.length > 0) {
      menu.addItem(i => i.setTitle('📝 粘贴到这里').setIcon('paste').onClick(() => this.pasteFiles()));
      if (fileClipboard.items.length > 1) {
        menu.addItem(i => i.setTitle('📋 选择粘贴内容…').setIcon('clipboard-list').onClick(() => this.showPasteMenu()));
      }
    }

    menu.addSeparator();

    // 压缩
    menu.addItem(i => i.setTitle('🗜️ 压缩为 ZIP').setIcon('package').onClick(() => this.compressToZip(targets)));
    if (is7zInstalled()) {
      menu.addItem(i => i.setTitle('🗜️ 压缩为 7z').setIcon('package').onClick(() => this.compressTo7z(targets)));
    }
    if (['.zip','.7z','.rar','.tar','.gz','.bz2','.xz','.tgz'].includes(ext)) {
      menu.addItem(i => i.setTitle('📂 解压到当前目录').setIcon('folder-plus').onClick(() => this.extractArchive(entry)));
    }

    menu.addSeparator();

    // 重命名
    if (isSingle) {
      menu.addItem(i => i.setTitle('✏️ 重命名').setIcon('pencil').onClick(() => this.renameItem(entry)));
    }
    if (targets.length > 1) {
      menu.addItem(i => i.setTitle('✏️ 批量重命名…').setIcon('pencil').onClick(() => {
        new BatchRenameModal(this.app, targets, this.currentPath, () => this.render()).open();
      }));
    }

    // 删除
    menu.addItem(i => i.setTitle('🗑️ 移动到回收站').setIcon('trash').onClick(() => this.deleteToRecycle(entry)));

    menu.addSeparator();

    // 预览
    if (!entry.isDir) {
      menu.addItem(i => i.setTitle('👁️ 预览').setIcon('eye').onClick(() => {
        if (isImageFile(entry.name) || isTextFile(entry.name) || isPDFFile(entry.name)) {
          this.previewFile(entry.path);
        } else {
          this.openFile(entry);
        }
      }));
    }

    // 引用
    if (entry.isDir) {
      menu.addItem(i => i.setTitle('插入文件夹索引').setIcon('list').onClick(() => this.insertFolderIndex(entry)));
    } else {
      menu.addItem(i => i.setTitle('插入文件引用').setIcon('link').onClick(() => this.insertFileRef(entry)));
    }
    menu.addItem(i => i.setTitle('🔗 复制路径').setIcon('copy').onClick(() => this.copyText(entry.path)));

    // WPS工具
    if ((isOfficeFile(entry.name) || isPDFFile(entry.name)) && targets.length === 1) {
      menu.addSeparator();
      menu.addItem(i => i.setTitle('🛠️ 文件工具').setIcon('wrench').onClick(() => {
        new WPSToolModal(this.app, entry, this).open();
      }));
    } else if (targets.length > 1) {
      menu.addSeparator();
      menu.addItem(i => i.setTitle('🛠️ 批量工具').setIcon('wrench').onClick(() => {
        new WPSToolModal(this.app, entry, this).open();
      }));
    }

    menu.addSeparator();
    menu.addItem(i => i.setTitle('🖥️ 在系统中打开').setIcon('external-link').onClick(() => this.openFile(entry)));

    menu.showAtMouseEvent(event);
  }

  /* =============================================================== */
  /*  新建菜单                                                       */
  /* =============================================================== */
  showNewMenu(event) {
    const menu = new Menu();
    menu.addItem(i => i.setTitle('📁 新建文件夹').setIcon('folder').onClick(() => this.createNewItem('folder')));
    menu.addItem(i => i.setTitle('📝 新建文本文档').setIcon('file').onClick(() => this.createNewItem('text')));
    menu.addSeparator();
    menu.addItem(i => i.setTitle('📄 新建 Word 文档').setIcon('file-text').onClick(() => this.createNewItem('word')));
    menu.addItem(i => i.setTitle('📊 新建 Excel 表格').setIcon('table').onClick(() => this.createNewItem('excel')));

    if (fileClipboard.items.length > 0) {
      menu.addSeparator();
      menu.addItem(i => i.setTitle('📝 粘贴').setIcon('paste').onClick(() => this.pasteFiles()));
      if (fileClipboard.items.length > 1) {
        menu.addItem(i => i.setTitle('📋 选择粘贴内容…').setIcon('clipboard-list').onClick(() => this.showPasteMenu()));
      }
    }

    menu.showAtMouseEvent(event);
  }

  async createNewItem(type) {
    let baseName, createFn;
    switch (type) {
      case 'folder': baseName = '新建文件夹'; createFn = async (p) => { mkdirSync(p); }; break;
      case 'text': baseName = '新建文本文档.txt'; createFn = async (p) => { writeFileSync(p, '', 'utf8'); }; break;
      case 'word': baseName = '新建 Word 文档.docx'; createFn = (p) => this.createWordDoc(p); break;
      case 'excel': baseName = '新建 Excel 表格.xlsx'; createFn = (p) => this.createExcelDoc(p); break;
      default: return;
    }
    let targetPath = join(this.currentPath, baseName);
    if (existsSync(targetPath)) {
      const ext = extname(baseName);
      const stem = baseName.slice(0, -ext.length);
      for (let i = 1; i < 999; i++) {
        targetPath = join(this.currentPath, `${stem} (${i})${ext}`);
        if (!existsSync(targetPath)) break;
      }
    }
    try {
      await createFn(targetPath);
      new Notice(`✅ 已创建: ${basename(targetPath)}`);
      this.selectedPaths.clear();
      this.selectedPaths.add(targetPath);
      this.render();
      setTimeout(() => {
        this.renameItem({ name: basename(targetPath), path: targetPath });
      }, 50);
    } catch (e) {
      new Notice(`❌ 创建失败: ${e.message}`);
    }
  }

  async createWordDoc(path) {
    const script = `
      try {
        $app = New-Object -ComObject Kwps.Application; $app.Visible = $false;
        $doc = $app.Documents.Add();
        $doc.SaveAs('${path.replace(/'/g, "''")}'); $doc.Close(); $app.Quit(); Write-Output 'ok';
      } catch { Write-Output 'fail' }
    `;
    const result = await runPowershell(script);
    if (result !== 'ok') { writeFileSync(path, '', 'utf8'); new Notice('⚠️ WPS 未检测到，已创建空白文件'); }
  }

  async createExcelDoc(path) {
    const script = `
      try {
        $app = New-Object -ComObject Ket.Application; $app.Visible = $false;
        $wb = $app.Workbooks.Add();
        $wb.SaveAs('${path.replace(/'/g, "''")}'); $wb.Close(); $app.Quit(); Write-Output 'ok';
      } catch { Write-Output 'fail' }
    `;
    const result = await runPowershell(script);
    if (result !== 'ok') { writeFileSync(path, '', 'utf8'); new Notice('⚠️ WPS 未检测到，已创建空白文件'); }
  }

  /* =============================================================== */
  /*  系统剪贴板操作                                                   */
  /* =============================================================== */
  copyFiles(entry) {
    const sources = this.getTargetPaths(entry);
    // ① 内部多点剪贴板（用于 Drive Opener 内部粘贴）
    fileClipboard.push(sources, 'copy');
    // ② 系统剪贴板：复制文件本体（粘贴到资源管理器等处时是实际文件）
    copyFilesToOsClipboard(sources).then(() => {
      if (sources.length === 1) new Notice(`📋 已复制: ${basename(sources[0])}`);
      else new Notice(`📋 已复制 ${sources.length} 项`);
    }).catch((err) => {
      // 复制失败时只显示错误，不回退写文本路径
      new Notice(`❌ 复制文件失败: ${err.message}`);
    });
  }

  /** 复制文件路径（纯文本）— 粘贴到对话框/笔记中引用文件 */
  copyFilePath(entry) {
    const sources = this.getTargetPaths(entry);
    const text = sources.join('\n');
    navigator.clipboard.writeText(text).then(() => {
      if (sources.length === 1) new Notice(`🔗 已复制路径: ${basename(sources[0])}`);
      else new Notice(`🔗 已复制 ${sources.length} 个路径`);
    }).catch(() => {
      // Electron 中 navigator.clipboard 常静默失败，降级到 textarea 方案
      const ta = document.createElement('textarea');
      ta.value = text;
      document.body.appendChild(ta);
      ta.select();
      document.execCommand('copy');
      document.body.removeChild(ta);
      if (sources.length === 1) new Notice(`🔗 已复制路径: ${basename(sources[0])}`);
      else new Notice(`🔗 已复制 ${sources.length} 个路径`);
    });
  }

  cutFiles(entry) {
    const sources = this.getTargetPaths(entry);
    fileClipboard.push(sources, 'cut');
    new Notice(`✂️ 已剪切 ${sources.length} 项 (剪贴板: ${fileClipboard.items.length} 项)`);
  }

  pasteFiles() {
    if (fileClipboard.items.length === 0) return;
    // 使用最近的一项
    const item = fileClipboard.items[0];
    const destDir = this.currentPath;
    let copied = 0;
    for (const src of item.sources) {
      try {
        const dest = safeDestPath(src, destDir);
        if (statSync(src).isDirectory()) { copyDirRecursive(src, dest); }
        else { copyFileSync(src, dest); }
        if (item.action === 'cut') {
          if (statSync(src).isDirectory()) removeDirRecursive(src);
          else unlinkSync(src);
        }
        copied++;
      } catch (e) {
        new Notice(`❌ 操作失败: ${basename(src)} - ${e.message}`);
      }
    }
    if (item.action === 'cut') fileClipboard.items.shift();
    new Notice(`✅ 粘贴完成: ${copied} 项`);
    this.render();
  }

  showPasteMenu() {
    if (fileClipboard.items.length <= 1) { this.pasteFiles(); return; }
    new PasteSelectModal(this.app, this.currentPath, this).open();
  }

  /* =============================================================== */
  /*  压缩/解压                                                      */
  /* =============================================================== */
  async compressToZip(targets) {
    if (targets.length === 0) return;
    let zipName = targets.length === 1 ? basename(targets[0]) + '.zip' : 'archive_' + Date.now() + '.zip';
    const zipPath = join(this.currentPath, zipName);
    try {
      // 有 7-Zip 时优先使用（兼容性更好，支持 Deflate64/BZip2/LZMA 等算法）
      if (is7zInstalled()) {
        await compressWith7z(targets, zipPath, 'zip');
        new Notice(`🗜️ 压缩完成: ${zipName}`);
        this.render();
        return;
      }
      // 无 7-Zip 时使用 PowerShell
      // 注意：PowerShell 5.1 的 Compress-Archive -Path 只接受文件，不支持目录
      // 当目标包含文件夹时，需先将所有内容复制到临时目录再压缩
      const escapedZipPath = zipPath.replace(/'/g, "''");
      const pathsArray = targets.map(p => `'${p.replace(/'/g, "''")}'`).join(',');
      const psScript = `
        $sources = @(${pathsArray});
        $dest = '${escapedZipPath}';
        # 检查是否有目录
        $hasDir = $false;
        foreach ($src in $sources) {
          if (Test-Path $src -PathType Container) { $hasDir = $true; break; }
        }
        if ($hasDir) {
          # 有目录：先将所有源复制到临时目录，再压缩临时目录内容
          $tempDir = Join-Path $env:TEMP ('de_zip_' + [System.Guid]::NewGuid().ToString('N'));
          New-Item -ItemType Directory -Path $tempDir -Force | Out-Null;
          try {
            foreach ($src in $sources) {
              $name = Split-Path $src -Leaf;
              Copy-Item -Path $src -Destination (Join-Path $tempDir $name) -Recurse -Force;
            }
            Get-ChildItem -Path $tempDir | Compress-Archive -DestinationPath $dest -Force;
          } finally {
            Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue;
          }
        } else {
          # 纯文件：直接压缩
          Compress-Archive -Path $sources -DestinationPath $dest -Force;
        }
        Write-Output 'ok';
      `;
      const result = await runPowershell(psScript);
      if (result === 'ok') { new Notice(`🗜️ 压缩完成: ${zipName}`); this.render(); }
      else { new Notice(`❌ 压缩失败`); }
    } catch (e) { new Notice(`❌ 压缩失败: ${e.message}`); }
  }

  async compressTo7z(targets) {
    if (!is7zInstalled()) { new Notice('⚠️ 请先安装 7-Zip'); return; }
    const name = targets.length === 1 ? basename(targets[0]) : 'archive_' + Date.now();
    const archivePath = join(this.currentPath, name + '.7z');
    try {
      await compressWith7z(targets, archivePath, '7z');
      new Notice(`🗜️ 7z 压缩完成: ${name}.7z`);
      this.render();
    } catch (e) { new Notice(`❌ 7z 压缩失败: ${e.message}`); }
  }

  async compressToRar(targets) {
    if (!is7zInstalled()) { new Notice('⚠️ 请先安装 7-Zip'); return; }
    const name = targets.length === 1 ? basename(targets[0]) : 'archive_' + Date.now();
    const archivePath = join(this.currentPath, name + '.rar');
    try {
      await compressWith7z(targets, archivePath, 'rar');
      new Notice(`🗜️ RAR 压缩完成: ${name}.rar`);
      this.render();
    } catch (e) { new Notice(`❌ RAR 压缩失败: ${e.message}`); }
  }

  async extractArchive(entry) {
    const ext = extname(entry.name).toLowerCase();
    const nameNoExt = basename(entry.name, ext);
    const outDir = join(this.currentPath, nameNoExt);

    // 支持的格式
    const supportedFormats = ['.zip', '.7z', '.rar', '.tar', '.gz', '.bz2', '.xz', '.tgz'];

    if (!supportedFormats.includes(ext)) {
      new Notice(`⚠️ 不支持的压缩格式: ${ext}`);
      return;
    }

    try {
      // 方案一：7-Zip（兼容性最好，支持几乎所有格式和压缩算法）
      if (is7zInstalled()) {
        await extractWith7z(entry.path, outDir);
        new Notice(`📂 已解压到: ${nameNoExt}`);
        this.render();
        return;
      }

      // 以下方案仅适用于 .zip
      if (ext !== '.zip') {
        new Notice('⚠️ 请先安装 7-Zip 以解压此格式');
        return;
      }

      // 方案二：Windows tar 命令（Windows 10 17063+，比 Expand-Archive 更可靠）
      try {
        if (!existsSync(outDir)) mkdirSync(outDir, { recursive: true });
        await extractWithTar(entry.path, outDir);
        new Notice(`📂 已解压到: ${nameNoExt}`);
        this.render();
        return;
      } catch (tarErr) {
        // tar 不可用时降级到 PowerShell
      }

      // 方案三：PowerShell Expand-Archive（局限性最大，作为最后手段）
      await extractWithPowerShell(entry.path, outDir);
      new Notice(`📂 已解压到: ${nameNoExt}`);
      this.render();
    } catch (e) {
      new Notice(`❌ 解压失败: ${e.message}`);
    }
  }

  /** 拖放文件到目标目录（移动或复制） */
  dropFilesInto(targetDir, isCopy = false) {
    const sources = this._dragSourcePaths;
    if (!sources || sources.size === 0) {
      new Notice('⚠️ 没有可拖放的文件');
      return;
    }
    // 过滤掉自身和子目录
    const normTarget = targetDir.replace(/\\/g, '/') + '/';
    const valid = [...sources].filter(p => {
      if (p === targetDir) return false;
      if (p.replace(/\\/g, '/') + '/' === normTarget) return false; // 自己的子目录
      if (p.replace(/\\/g, '/').startsWith(normTarget)) return false;
      return true;
    });
    if (valid.length === 0) {
      new Notice('⚠️ 不能将文件移动到自身或子目录');
      return;
    }
    let success = 0;
    for (const src of valid) {
      try {
        const dest = safeDestPath(src, targetDir);
        if (isCopy) {
          if (statSync(src).isDirectory()) copyDirRecursive(src, dest);
          else copyFileSync(src, dest);
        } else {
          // 移动：先用 renameSync（同盘符），失败则回退到复制+删除
          try {
            renameSync(src, dest);
          } catch (renameErr) {
            if (statSync(src).isDirectory()) {
              copyDirRecursive(src, dest);
              removeDirRecursive(src);
            } else {
              copyFileSync(src, dest);
              unlinkSync(src);
            }
          }
        }
        success++;
      } catch (e) {
        new Notice(`❌ 操作失败: ${basename(src)}`);
      }
    }
    if (success > 0) {
      new Notice(`${isCopy ? '📋 已复制' : '📂 已移动'} ${success} 项到 ${basename(targetDir)}`);
      this._dragSourcePaths = null;
      this.selectedPaths.clear();
      this.render();
    }
  }

  /* =============================================================== */
  /*  删除                                                           */
  /* =============================================================== */
  async deleteToRecycle(entry) {
    const targets = this.getTargetPaths(entry);
    const plural = targets.length > 1 ? `这 ${targets.length} 项` : basename(targets[0]);
    new Notice(`🗑️ 正在将 ${plural} 移入回收站...`);
    let success = 0;
    for (const p of targets) {
      try {
        const isDir = statSync(p).isDirectory();
        const psPath = p.replace(/\\/g, '\\\\');
        const script = isDir
          ? `Add-Type -AssemblyName Microsoft.VisualBasic; [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteDirectory('${psPath.replace(/'/g, "''")}', 'OnlyErrorDialogs', 'SendToRecycleBin');`
          : `Add-Type -AssemblyName Microsoft.VisualBasic; [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile('${psPath.replace(/'/g, "''")}', 'OnlyErrorDialogs', 'SendToRecycleBin');`;
        await runPowershell(script);
        success++;
      } catch (e) { new Notice(`❌ 删除失败: ${basename(p)}`); }
    }
    new Notice(`🗑️ 已将 ${success} 项移入回收站`);
    this.selectedPaths.clear();
    this.previewFilePath = null;
    this.render();
  }

  /* =============================================================== */
  /*  重命名                                                         */
  /* =============================================================== */
  renameItem(entry) {
    new RenameModal(this.app, entry.name, entry.path, this.currentPath, (newPath) => {
      this.selectedPaths.clear();
      this.selectedPaths.add(newPath);
      this.render();
    }).open();
  }

  /* =============================================================== */
  /*  插入引用                                                       */
  /* =============================================================== */
  insertFileRef(entry) {
    const mdView = this.app.workspace.getActiveViewOfType(MarkdownView);
    if (!mdView) { new Notice('⚠️ 请先打开一个笔记文件'); return; }
    const editor = mdView.editor;
    const link = `[${entry.name}](file:///${entry.path.replace(/\\/g, '/')})`;
    editor.replaceSelection(link);
    new Notice('✅ 已插入文件引用');
  }

  insertFolderIndex(entry) {
    const mdView = this.app.workspace.getActiveViewOfType(MarkdownView);
    if (!mdView) { new Notice('⚠️ 请先打开一个笔记文件'); return; }
    let names;
    try { names = readdirSync(entry.path, { encoding: 'utf8' }); }
    catch { new Notice('❌ 读取文件夹失败'); return; }
    const items = names.map(n => {
      try { return { name: n, isDir: statSync(join(entry.path, n)).isDirectory() }; }
      catch { return null; }
    }).filter(Boolean).sort((a, b) => {
      if (a.isDir !== b.isDir) return a.isDir ? -1 : 1;
      return a.name.localeCompare(b.name, 'zh-CN');
    });
    let content = `### 📁 ${entry.name}\n\n`;
    for (const it of items) content += `- ${it.isDir ? '📁' : '📄'} \`${it.name}\`\n`;
    mdView.editor.replaceSelection(content);
    new Notice('✅ 已插入文件夹索引');
  }

    copyText(text) {
    navigator.clipboard.writeText(text)
      .then(() => new Notice('📋 已复制到剪贴板'))
      .catch(() => {
        const ta = document.createElement('textarea');
        ta.value = text;
        document.body.appendChild(ta);
        ta.select();
        document.execCommand('copy');
        document.body.removeChild(ta);
        new Notice('📋 已复制到剪贴板');
      });
  }
}

/* =============================================================== */
/*   工具：格式化文件大小                                           */
/* =============================================================== */
function formatSize(bytes) {
  if (bytes === 0) return '0 B';
  const units = ['B', 'KB', 'MB', 'GB', 'TB'];
  const k = 1024;
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + units[i];
}

/* ===================================================================
   插件主入口
   =================================================================== */
module.exports = class DriveOpenerPlugin extends Plugin {
  async onload() {
    this.injectStyles();

    this.registerView(VIEW_TYPE, (leaf) => new DriveExplorerView(leaf));

    this.addRibbonIcon('folder', '文件浏览器', () => {
      this.openNewTab();
    });

    this.addCommand({
      id: 'open-drive-explorer',
      name: '打开文件浏览器',
      callback: () => this.openNewTab(),
    });
  }

  async openNewTab() {
    const leaf = this.app.workspace.getLeaf('tab');
    await leaf.setViewState({ type: VIEW_TYPE, active: true });
    this.app.workspace.revealLeaf(leaf);
  }

  injectStyles() {
    const style = document.createElement('style');
    style.textContent = `
      .de-header { padding: 12px 16px 8px; border-bottom: 1px solid var(--background-modifier-border); }
      .de-nav-row { display: flex; gap: 6px; margin-bottom: 8px; align-items: center; }
      .de-nav-btn { background: none; border: 1px solid var(--background-modifier-border); border-radius: 4px; padding: 2px 8px; cursor: pointer; font-size: 13px; color: var(--text-muted); }
      .de-nav-btn:disabled { opacity: 0.35; cursor: default; }
      .de-nav-btn:hover:not(:disabled) { background: var(--background-modifier-hover); color: var(--text-normal); }
      .de-path { font-family: var(--font-monospace); font-size: 13px; color: var(--text-accent); padding: 4px 0; word-break: break-all; }
      .de-refresh { background: none; border: 1px solid var(--background-modifier-border); border-radius: 4px; padding: 2px 10px; cursor: pointer; font-size: 12px; color: var(--text-muted); margin-top: 6px; }
      .de-refresh:hover { background: var(--background-modifier-hover); }
      .de-sel-info { display: inline-block; margin-left: 8px; font-size: 12px; color: var(--text-accent); }
      .de-list { padding: 4px 0; overflow-y: auto; }
      .de-item { display: flex; align-items: center; gap: 8px; padding: 5px 16px; cursor: pointer; border-radius: 3px; user-select: none; }
      .de-item:hover { background: var(--background-modifier-hover); }
      .de-item.selected { background: var(--text-accent-hover, var(--interactive-accent-hover)) !important; color: var(--text-on-accent, var(--text-normal)); }
      .de-item.de-dragging { opacity: 0.5; }
      .de-item.de-drag-over { outline: 2px solid var(--interactive-accent); outline-offset: -2px; background: var(--background-modifier-hover); }
      .de-icon { font-size: 16px; line-height: 1; flex-shrink: 0; width: 22px; text-align: center; }
      .de-thumb { width: 32px; height: 32px; object-fit: cover; border-radius: 3px; flex-shrink: 0; }
      .de-name { font-size: 14px; color: var(--text-normal); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; flex: 1; }
      .de-msg { padding: 24px 16px; color: var(--text-muted); text-align: center; font-size: 14px; }
      .de-err { color: var(--text-error); }
      .de-wps-modal h2 { margin: 0 0 8px; }
      .de-wps-modal h3 { margin: 12px 0 4px; font-size: 15px; }
      .de-wps-hint { color: var(--text-muted); font-size: 13px; }
      .de-wps-warn { color: var(--text-warning); font-size: 13px; background: var(--background-modifier-error-rgb); padding: 8px 12px; border-radius: 4px; }
      .de-drive-card { padding: 16px; border: 1px solid var(--background-modifier-border); border-radius: 8px; cursor: pointer; transition: all 0.15s; }
      .de-drive-card:hover { border-color: var(--text-accent); background: var(--background-modifier-hover); }
      .de-header-actions { display: flex; gap: 6px; margin-top: 6px; flex-wrap: wrap; }
      .de-search-input { padding: 4px 8px; border: 1px solid var(--background-modifier-border); border-radius: 4px; font-size: 12px; background: var(--background-primary); color: var(--text-normal); }
      .de-main-area { display: flex; height: calc(100% - 130px); }
      .de-list-container { flex: 1; overflow: auto; }
      .de-preview-panel { width: 300px; min-width: 200px; border-left: 1px solid var(--background-modifier-border); padding: 8px; overflow: auto; font-size: 13px; }
      .de-preset-row button:hover { background: var(--background-modifier-hover); }
      @keyframes de-slide-in-right { from { transform: translateX(28px); opacity: 0.3; } to { transform: translateX(0); opacity: 1; } }
      @keyframes de-slide-in-left { from { transform: translateX(-28px); opacity: 0.3; } to { transform: translateX(0); opacity: 1; } }
      .de-slide-right { animation: de-slide-in-right 180ms cubic-bezier(0.15, 0.85, 0.35, 1); }
      .de-slide-left { animation: de-slide-in-left 180ms cubic-bezier(0.15, 0.85, 0.35, 1); }
    `;
    document.head.appendChild(style);
  }
};
