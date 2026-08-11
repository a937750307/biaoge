// main.js
// 由 index.html 拆分出的全部 JS 逻辑

// 字体切换功能，默认黑体
function changeFont() {
  const font = document.getElementById('fontSelector').value;
  document.querySelector('.container').style.fontFamily = font + ',sans-serif';
  // 预览区slip也同步字体
  document.getElementById('slipWrap').style.fontFamily = font + ',sans-serif';
}
// 页面加载时默认宋体
window.addEventListener('DOMContentLoaded', function() {
  document.querySelector('.container').style.fontFamily = 'SimSun,sans-serif';
  document.getElementById('slipWrap').style.fontFamily = 'SimSun,sans-serif';
});
// 设置“年月”默认值为上一个月
(function setDefaultMonth() {
  const dateInput = document.getElementById('customDate');
  if (dateInput) {
    const now = new Date();
    let year = now.getFullYear();
    let month = now.getMonth(); // 0-11, 上一个月
    if (month === 0) {
      year--;
      month = 12;
    }
    const val = year + '年' + (month < 10 ? '0' : '') + month + '月';
    dateInput.value = val;
  }
})();
let tableData = [];
let tableHeader = [];
let filteredData = [];
let fileName = '';
let merges = [];

// 展开/收起原始表格
function toggleTable() {
  const tableDiv = document.getElementById('tableWrap');
  const btn = document.getElementById('toggleTableBtn');
  if (tableDiv.style.display === 'none') {
    tableDiv.style.display = '';
    btn.textContent = '收起原始表格';
  } else {
    tableDiv.style.display = 'none';
    btn.textContent = '展开原始表格';
  }
}
// 上传/拖拽
const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');
// 只保留按钮的点击事件，去除整个区域的点击事件，避免弹窗两次
uploadArea.addEventListener('dragover', e => { e.preventDefault(); uploadArea.classList.add('dragover'); });
uploadArea.addEventListener('dragleave', e => { e.preventDefault(); uploadArea.classList.remove('dragover'); });
uploadArea.addEventListener('drop', e => {
  e.preventDefault();
  uploadArea.classList.remove('dragover');
  handleFile(e.dataTransfer.files[0]);
});
fileInput.addEventListener('change', e => {
  handleFile(e.target.files[0]);
  // 解决连续上传同一文件不触发change的问题
  fileInput.value = '';
});

function handleFile(file) {
  if (!file) return;
  fileName = file.name;
  // 自动将文件名（去扩展名）填入大标题，但可编辑
  const titleInput = document.getElementById('customTitle');
  if (titleInput) {
    const nameNoExt = file.name.replace(/\.[^.]+$/, '');
    titleInput.value = nameNoExt;
  }
  const reader = new FileReader();
  const isCsv = /\.csv$/i.test(file.name);
  reader.onload = function(e) {
    let data = e.target.result;
    let sheet, json;
    if (isCsv) {
      // 自动检测分隔符
      let delimiter = ',';
      if (data.indexOf(';') > -1 && (data.indexOf(';') < data.indexOf(',') || data.indexOf(',') === -1)) {
        delimiter = ';';
      } else if (data.indexOf('\t') > -1) {
        delimiter = '\t';
      }
      // 逐行分割为二维数组
      let lines = data.split(/\r?\n/).filter(line => line.trim() !== '');
      let aoa = lines.map(line => line.split(delimiter));
      sheet = XLSX.utils.aoa_to_sheet(aoa);
      json = aoa;
    } else {
      // xls/xlsx用二进制
      let workbook = XLSX.read(data, { type: 'binary' });
      sheet = workbook.Sheets[workbook.SheetNames[0]];
      json = XLSX.utils.sheet_to_json(sheet, { header: 1, blankrows: false, raw: false });
    }
    if (!json || json.length === 0) return;
    window.lastJsonRaw = json; // 保存原始json用于多行表头
    tableHeader = json[0];
    // 自动忽略空白行
    tableData = json.slice(1).filter(row => Array.isArray(row) && row.some(cell => cell !== undefined && String(cell).trim() !== ''));
    filteredData = tableData;
    merges = (sheet['!merges'] || []).map(m => ({
      s: { r: m.s.r, c: m.s.c },
      e: { r: m.e.r, c: m.e.c }
    }));
    renderTable();
    document.getElementById('actionsBar').style.display = '';
    // 显示展开原始表格按钮
    var btnWrap = document.getElementById('toggleTableBtnWrap');
    if(btnWrap) btnWrap.style.display = '';
  };
  if (isCsv) {
    reader.readAsText(file, 'utf-8');
  } else {
    reader.readAsBinaryString(file);
  }
}

// 根据 merges 计算每个单元格的 skip 状态和起始单元格的 rowspan/colspan
function buildMergeInfo(rowCount, colCount) {
  const skip = Array.from({ length: rowCount }, () => Array(colCount).fill(false));
  const mergeMap = {};
  merges.forEach(m => {
    if (m.s.r >= rowCount || m.s.c >= colCount) return;
    const key = `${m.s.r},${m.s.c}`;
    let rowspan = Math.min(m.e.r, rowCount - 1) - m.s.r + 1;
    let colspan = m.e.c - m.s.c + 1;
    if (rowspan < 1 || colspan < 1) return;
    mergeMap[key] = { rowspan, colspan };
    for (let r = m.s.r; r <= Math.min(m.e.r, rowCount - 1); r++) {
      for (let c = m.s.c; c <= m.e.c; c++) {
        if (!(r === m.s.r && c === m.s.c)) {
          skip[r][c] = true;
        }
      }
    }
  });
  return { skip, mergeMap };
}

// 渲染表格
function renderTable() {
  // 原始表格区：严格还原用户上传的表格内容
  let raw = window.lastJsonRaw ? window.lastJsonRaw.slice() : [];
  if (raw.length === 0) return;

  // 计算最大列数
  let colCount = Math.max(
    ...raw.map(row => Array.isArray(row) ? row.length : 0)
  );

  // 补齐每一行
  raw = raw.map(row => {
    if (!Array.isArray(row)) row = [];
    if (row.length < colCount) return row.concat(Array(colCount - row.length).fill(''));
    return row;
  });

  // 去除末尾连续空白行
  let lastNotEmpty = raw.length - 1;
  for (; lastNotEmpty >= 0; lastNotEmpty--) {
    if (raw[lastNotEmpty].some(cell => cell !== undefined && String(cell).trim() !== '')) break;
  }
  let rowCount = lastNotEmpty + 1;
  raw = raw.slice(0, rowCount);

  // 表头行数：至少为设置值，且覆盖所有合并单元格
  const headerRowsInput = Math.max(1, parseInt(document.getElementById('customHeaderRows').value) || 1);
  const mergeMaxRow = merges.length > 0
    ? Math.max(...merges.map(m => Math.min(m.e.r, rowCount - 1)))
    : 0;
  const headerRows = Math.min(Math.max(headerRowsInput, mergeMaxRow + 1), rowCount);

  // 构建完整表格的合并信息
  const { skip, mergeMap } = buildMergeInfo(rowCount, colCount);

  // 渲染原始表格
  let html = '<table><thead>';
  for (let r = 0; r < headerRows; r++) {
    html += '<tr>';
    for (let c = 0; c < colCount; c++) {
      if (skip[r][c]) continue;
      const key = `${r},${c}`;
      let attrs = '';
      if (mergeMap[key]) {
        if (mergeMap[key].rowspan > 1) attrs += ` rowspan="${mergeMap[key].rowspan}"`;
        if (mergeMap[key].colspan > 1) attrs += ` colspan="${mergeMap[key].colspan}"`;
      }
      const val = raw[r][c] !== undefined ? raw[r][c] : '';
      html += `<th${attrs} title="${val}">${val}</th>`;
    }
    html += '</tr>';
  }
  html += '</thead><tbody>';
  for (let r = headerRows; r < rowCount; r++) {
    html += '<tr>';
    for (let c = 0; c < colCount; c++) {
      if (skip[r][c]) continue;
      const key = `${r},${c}`;
      let attrs = '';
      if (mergeMap[key]) {
        if (mergeMap[key].rowspan > 1) attrs += ` rowspan="${mergeMap[key].rowspan}"`;
        if (mergeMap[key].colspan > 1) attrs += ` colspan="${mergeMap[key].colspan}"`;
      }
      const val = raw[r][c] !== undefined ? raw[r][c] : '';
      html += `<td${attrs} title="${val}">${val}</td>`;
    }
    html += '</tr>';
  }
  html += '</tbody></table>';
  document.getElementById('tableWrap').innerHTML = html;

  // 工资条/成绩单 slip 区
  const title = document.getElementById('customTitle').value;
  const company = document.getElementById('customCompany').value;
  const date = document.getElementById('customDate').value;

  // 准备 slip 表头行
  let jsonHeaderRows = raw.slice(0, headerRows);

  let slips = '';
  for (let r = headerRows; r < rowCount; r++) {
    slips += `<div class="single-slip">`;
    slips += `<div class="slip-title">${title}</div>`;
    slips += `<div class="slip-meta"><span>${company}</span><span>${date}</span></div>`;
    slips += '<table class="slip-table">';
    // 渲染表头（支持合并单元格，多行表头统一放在一个 thead 中）
    slips += '<thead>';
    for (let h = 0; h < jsonHeaderRows.length; h++) {
      slips += '<tr>';
      for (let c = 0; c < colCount; c++) {
        if (skip[h][c]) continue;
        const key = `${h},${c}`;
        let attrs = '';
        if (mergeMap[key]) {
          if (mergeMap[key].rowspan > 1) attrs += ` rowspan="${mergeMap[key].rowspan}"`;
          if (mergeMap[key].colspan > 1) attrs += ` colspan="${mergeMap[key].colspan}"`;
        }
        const val = jsonHeaderRows[h][c] !== undefined ? jsonHeaderRows[h][c] : '';
        slips += `<th${attrs}>${val}</th>`;
      }
      slips += '</tr>';
    }
    slips += '</thead>';
    // 渲染数据行
    slips += '<tbody><tr>';
    for (let c = 0; c < colCount; c++) {
      if (skip[r][c]) continue;
      const key = `${r},${c}`;
      let attrs = '';
      if (mergeMap[key]) {
        if (mergeMap[key].rowspan > 1) attrs += ` rowspan="${mergeMap[key].rowspan}"`;
        if (mergeMap[key].colspan > 1) attrs += ` colspan="${mergeMap[key].colspan}"`;
      }
      const val = raw[r][c] !== undefined ? raw[r][c] : '';
      slips += `<td${attrs} title="${val}">${val}</td>`;
    }
    slips += '</tr></tbody>';
    slips += '</table>';
    slips += '<hr class="cut-line" />';
    slips += '</div>';
  }
  document.getElementById('slipWrap').innerHTML = slips;
  // 预览后自动滚动到slip区
  if (slips) {
    setTimeout(() => {
      document.getElementById('slipWrap').scrollIntoView({ behavior: 'smooth' });
    }, 100);
  }
}

// 导出为PDF（A4自适应，导出slipWrap内容）
function exportPDF() {
  const slipWrap = document.getElementById('slipWrap');
  if (!slipWrap) return;
  // 显示导出中提示
  let tip = document.createElement('div');
  tip.id = 'exportingTip';
  tip.innerText = '正在导出中，请稍等...';
  tip.style.position = 'fixed';
  tip.style.left = '50%';
  tip.style.top = '30%';
  tip.style.transform = 'translate(-50%, -50%)';
  tip.style.background = 'rgba(34,34,34,0.95)';
  tip.style.color = '#fff';
  tip.style.fontSize = '20px';
  tip.style.padding = '22px 38px';
  tip.style.borderRadius = '10px';
  tip.style.zIndex = '9999';
  tip.style.boxShadow = '0 2px 16px #0005';
  document.body.appendChild(tip);
  // 获取所有单条模板
  const slips = Array.from(slipWrap.querySelectorAll('.single-slip'));
  if (slips.length === 0) { document.body.removeChild(tip); return; }
  const pdf = new window.jspdf.jsPDF({ unit: 'mm', format: 'a4' });
  const pageWidth = pdf.internal.pageSize.getWidth();
  const pageHeight = pdf.internal.pageSize.getHeight();
  let pagePending = false;
  let y = 0;
  // 递归渲染每条 slip，保证每条 slip 不被截断
  function renderSlipToCanvas(slip, cb) {
    html2canvas(slip, { scale: 2, useCORS: true, backgroundColor: '#fff' }).then(canvas => {
      cb(canvas);
    });
  }
  function addSlipToPDF(idx) {
    if (idx >= slips.length) {
      // 文件名：大标题+年月日时分秒
      const title = (document.getElementById('customTitle')?.value || '导出')
        .replace(/[/\\:*?"<>|]/g, '') // 去除非法字符
        .trim(); 
      const now = new Date();
      const pad = n => n < 10 ? '0' + n : n;
      const timeStr = `${now.getFullYear()}${pad(now.getMonth()+1)}${pad(now.getDate())}${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`;
      pdf.save(`${title}_${timeStr}.pdf`);
      // 移除导出中提示
      document.body.removeChild(tip);
      return;
    }
    renderSlipToCanvas(slips[idx], canvas => {
      const imgWidth = pageWidth;
      const imgHeight = canvas.height * imgWidth / canvas.width;
      // 如果当前页剩余空间不足，自动分页
      if (y + imgHeight > pageHeight - 2) {
        pdf.addPage();
        y = 0;
      }
      pdf.addImage(canvas.toDataURL('image/jpeg', 1.0), 'JPEG', 0, y, imgWidth, imgHeight);
      y += imgHeight;
      addSlipToPDF(idx + 1);
    });
  }
  // 首页不addPage
  y = 0;
  addSlipToPDF(0);
}

// 搜索
function searchTable() {
  const val = document.getElementById('searchInput').value.trim();
  if (!val) {
    filteredData = tableData;
    renderTable();
    document.getElementById('slipWrap').style.display = '';
    return;
  }
  // 支持多个关键词（逗号分隔），所有关键词都需匹配
  const keywords = val.split(',').map(s => s.trim()).filter(Boolean);
  filteredData = tableData.filter(row => {
    // 每个关键词都要在前5列任意单元格中出现
    return keywords.every(kw => {
      for (let i = 0; i < Math.min(5, row.length); i++) {
        if ((row[i] + '').includes(kw)) return true;
      }
      return false;
    });
  });
  renderTable();
  // 只显示slipWrap中与搜索结果对应的条目
  // slipWrap每条对应filteredData的索引（row在tableData中的索引+headerRows-1）
  const slipWrap = document.getElementById('slipWrap');
  if (!slipWrap) return;
  if (!filteredData.length) {
    slipWrap.innerHTML = '<div style="color:#f56c6c;text-align:center;margin:30px 0 0 0;font-size:17px;">未找到匹配结果</div>';
    return;
  }
  // 重新渲染slipWrap，只显示匹配的
  const title = document.getElementById('customTitle').value;
  const company = document.getElementById('customCompany').value;
  const date = document.getElementById('customDate').value;
  const headerRows = Math.max(1, parseInt(document.getElementById('customHeaderRows').value) || 1);
  let slips = '';
  let jsonHeaderRows = [];
  if (window.lastJsonRaw && Array.isArray(window.lastJsonRaw)) {
    jsonHeaderRows = window.lastJsonRaw.slice(0, headerRows);
  } else {
    jsonHeaderRows = [tableHeader];
  }
  let colCount = tableHeader.length;
  // 补齐表头行
  jsonHeaderRows = jsonHeaderRows.map(row => {
    if (!Array.isArray(row)) row = [];
    if (row.length < colCount) return row.concat(Array(colCount - row.length).fill(''));
    return row;
  });
  // 计算表头合并信息
  const rawRowCount = window.lastJsonRaw ? window.lastJsonRaw.length : 1;
  const { skip, mergeMap } = buildMergeInfo(Math.min(rawRowCount, headerRows), colCount);
  for (let idx = 0; idx < filteredData.length; idx++) {
    const row = filteredData[idx];
    slips += `<div class="single-slip">`;
    slips += `<div class="slip-title">${title}</div>`;
    slips += `<div class="slip-meta"><span>${company}</span><span>${date}</span></div>`;
    slips += '<table class="slip-table">';
    slips += '<thead>';
    for (let h = 0; h < jsonHeaderRows.length; h++) {
      slips += '<tr>';
      for (let c = 0; c < colCount; c++) {
        if (skip[h] && skip[h][c]) continue;
        const key = `${h},${c}`;
        let attrs = '';
        if (mergeMap[key]) {
          if (mergeMap[key].rowspan > 1) attrs += ` rowspan="${mergeMap[key].rowspan}"`;
          if (mergeMap[key].colspan > 1) attrs += ` colspan="${mergeMap[key].colspan}"`;
        }
        const val = jsonHeaderRows[h][c] !== undefined ? jsonHeaderRows[h][c] : '';
        slips += `<th${attrs}>${val}</th>`;
      }
      slips += '</tr>';
    }
    slips += '</thead>';
    slips += '<tbody><tr>';
    for (let c = 0; c < colCount; c++) {
      const val = row[c] !== undefined ? row[c] : '';
      slips += `<td title="${val}">${val}</td>`;
    }
    slips += '</tr></tbody></table>';
    slips += '<hr class="cut-line" />';
    slips += '</div>';
  }
  slipWrap.innerHTML = slips;
}
