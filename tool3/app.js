/* ============================================================
   海外仓工具箱 - 合并 JS
   模块1: 订单收入匹配工具
   模块2: 海外仓账单自动化上传与核算
   ============================================================ */

/* ── 侧边导航切换 ─────────────────────────────────────────── */
document.querySelectorAll('.nav-item').forEach(function(item) {
  item.addEventListener('click', function() {
    var page = this.getAttribute('data-page');
    // 切换 nav active
    document.querySelectorAll('.nav-item').forEach(function(n) { n.classList.remove('active'); });
    this.classList.add('active');
    // 切换 section
    document.querySelectorAll('.page-section').forEach(function(s) { s.classList.remove('active'); });
    var target = document.getElementById('page-' + page);
    if (target) target.classList.add('active');
  });
});

/* ── 侧边栏折叠/展开 ──────────────────────────────────────── */
(function() {
  var sidebar = document.getElementById('sidebar');
  var toggleBtn = document.getElementById('sidebarToggle');
  var toggleIcon = document.getElementById('toggleIcon');
  var STORAGE_KEY = 'sidebar_collapsed';

  function setCollapsed(collapsed) {
    if (collapsed) {
      sidebar.classList.add('collapsed');
      document.body.classList.add('sidebar-collapsed');
      toggleIcon.textContent = '▶';
      localStorage.setItem(STORAGE_KEY, '1');
    } else {
      sidebar.classList.remove('collapsed');
      document.body.classList.remove('sidebar-collapsed');
      toggleIcon.textContent = '◀';
      localStorage.setItem(STORAGE_KEY, '0');
    }
  }

  // 恢复上次状态
  setCollapsed(localStorage.getItem(STORAGE_KEY) === '1');

  toggleBtn.addEventListener('click', function() {
    setCollapsed(!sidebar.classList.contains('collapsed'));
  });
})();

/* ============================================================
   模块1: 订单收入匹配工具
   ============================================================ */
(function () {
  'use strict';

  var multiFileInput   = document.getElementById('multiFileInput');
  var chooseFilesBtn   = document.getElementById('chooseFilesBtn');
  var uploadMainArea   = document.getElementById('uploadMainArea');
  var uploadStatus     = document.getElementById('uploadStatus');

  var slotBase         = document.getElementById('slotBase');
  var slotIncome       = document.getElementById('slotIncome');
  var dotBase          = document.getElementById('dotBase');
  var dotIncome        = document.getElementById('dotIncome');
  var slotBaseEmpty    = document.getElementById('slotBaseEmpty');
  var slotIncomeEmpty  = document.getElementById('slotIncomeEmpty');
  var slotBaseName     = document.getElementById('slotBaseName');
  var slotIncomeName   = document.getElementById('slotIncomeName');
  var slotBaseTag      = document.getElementById('slotBaseTag');
  var slotIncomeTag    = document.getElementById('slotIncomeTag');
  var slotBaseRemove   = document.getElementById('slotBaseRemove');
  var slotIncomeRemove = document.getElementById('slotIncomeRemove');

  var sheetSelectWrap  = document.getElementById('sheetSelectWrap');
  var sheetSelect      = document.getElementById('sheetSelect');
  var matchBtn         = document.getElementById('matchBtn');
  var resetAllBtn      = document.getElementById('resetAllBtn');
  var downloadMainBtn  = document.getElementById('downloadMainBtn');
  var downloadRemindBtn= document.getElementById('downloadRemindBtn');
  var progressBar      = document.getElementById('progressBar');
  var processStatus    = document.getElementById('processStatus');
  var logBox           = document.getElementById('logBox');
  var statTotal        = document.getElementById('statTotal');
  var statNeed         = document.getElementById('statNeed');
  var statStatus       = document.getElementById('statStatus');
  var rNeed            = document.getElementById('rNeed');
  var rRound1          = document.getElementById('rRound1');
  var rRound2          = document.getElementById('rRound2');
  var rUnmatched       = document.getElementById('rUnmatched');
  var remindCount      = document.getElementById('remindCount');
  var remindBody       = document.getElementById('remindBody');
  var previewBody      = document.getElementById('previewBody');
  var helpBtnMatch     = document.getElementById('helpBtnMatch');
  var helpModalMatch   = document.getElementById('helpModalMatch');
  var closeHelpMatch   = document.getElementById('closeHelpMatch');

  var baseFile   = null, incomeFile  = null;
  var baseWb     = null, incomeWb    = null;
  var outMainWb  = null, outRemindWb = null;
  var logs       = [];
  var processing = false;

  function setStatus(el, msg, type) {
    el.textContent = msg;
    el.className = 'status' + (type ? ' ' + type : '');
  }

  function setProgress(pct, msg) {
    progressBar.style.width = pct + '%';
    if (msg != null) setStatus(processStatus, msg, '');
  }

  function addLog(msg, type) {
    var now = new Date();
    var ts = now.toTimeString().slice(0, 8);
    var prefix = type === 'ok' ? '✓' : type === 'warn' ? '⚠' : type === 'err' ? '✗' : '·';
    logs.push('[' + ts + '] ' + prefix + ' ' + msg);
    logBox.textContent = logs.join('\n');
    logBox.scrollTop = logBox.scrollHeight;
  }

  function clearLogs() {
    logs = [];
    logBox.textContent = '暂无日志，请先上传文件。';
  }

  function isBlank(v) {
    return v === null || v === undefined || v === '' || (typeof v === 'string' && v.trim() === '');
  }

  function toNum(v) {
    if (v === null || v === undefined || v === '') return null;
    if (typeof v === 'number') return isNaN(v) ? null : v;
    var s = String(v).replace(/,/g, '').trim();
    var m = s.match(/-?[\d]+(?:\.[\d]+)?/);
    return m ? parseFloat(m[0]) : null;
  }

  function fmt2(v) {
    var n = toNum(v);
    return n == null ? '' : n.toFixed(2);
  }

  function readWorkbook(file) {
    return file.arrayBuffer().then(function(buf) {
      return XLSX.read(buf, { type: 'array' });
    });
  }

  function sheetToRows(ws) {
    return XLSX.utils.sheet_to_json(ws, { defval: null });
  }

  function normalizeRows(rows) {
    return rows.map(function(r) {
      var out = {};
      Object.keys(r).forEach(function(k) { out[typeof k === 'string' ? k.trim() : k] = r[k]; });
      return out;
    });
  }

  function stripDash(s) {
    return String(s).replace(/-/g, '').trim();
  }

  function downloadWb(wb, filename) {
    var out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    var blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url; a.download = filename;
    document.body.appendChild(a); a.click(); a.remove();
    URL.revokeObjectURL(url);
  }

  function detectFileRole(wb) {
    for (var i = 0; i < wb.SheetNames.length; i++) {
      var ws = wb.Sheets[wb.SheetNames[i]];
      var aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
      for (var j = 0; j < Math.min(10, aoa.length); j++) {
        var row = (aoa[j] || []).map(function(v) { return typeof v === 'string' ? v.trim() : v; });
        if (row.includes('卖家订单号') && row.includes('订单收入')) return 'base';
        if (row.includes('卖家订单号') && row.includes('人民币'))  return 'income';
      }
    }
    return 'unknown';
  }

  function updateSlotUI() {
    if (baseFile) {
      slotBase.classList.add('filled'); slotBase.classList.remove('error');
      dotBase.classList.add('green');
      slotBaseEmpty.style.display = 'none';
      slotBaseName.textContent = baseFile.name;
      slotBaseName.style.display = 'block';
      slotBaseTag.style.display = 'block';
      slotBaseRemove.style.display = 'inline-block';
    } else {
      slotBase.classList.remove('filled', 'error');
      dotBase.classList.remove('green', 'red');
      slotBaseEmpty.style.display = 'block';
      slotBaseName.style.display = 'none';
      slotBaseTag.style.display = 'none';
      slotBaseRemove.style.display = 'none';
    }
    if (incomeFile) {
      slotIncome.classList.add('filled'); slotIncome.classList.remove('error');
      dotIncome.classList.add('green');
      slotIncomeEmpty.style.display = 'none';
      slotIncomeName.textContent = incomeFile.name;
      slotIncomeName.style.display = 'block';
      slotIncomeTag.style.display = 'block';
      slotIncomeRemove.style.display = 'inline-block';
    } else {
      slotIncome.classList.remove('filled', 'error');
      dotIncome.classList.remove('green', 'red');
      slotIncomeEmpty.style.display = 'block';
      slotIncomeName.style.display = 'none';
      slotIncomeTag.style.display = 'none';
      slotIncomeRemove.style.display = 'none';
    }
    if (baseFile && incomeFile) {
      uploadMainArea.classList.add('all-ready');
    } else {
      uploadMainArea.classList.remove('all-ready');
    }
    updateMatchBtn();
    updateStatStatus();
  }

  function updateMatchBtn() {
    matchBtn.disabled = !(baseFile && incomeFile) || processing;
  }

  function updateStatStatus() {
    if (processing) {
      statStatus.textContent = '处理中'; statStatus.className = 'stat-value warn';
    } else if (outMainWb) {
      statStatus.textContent = '已完成'; statStatus.className = 'stat-value success';
    } else if (baseFile || incomeFile) {
      statStatus.textContent = '待匹配'; statStatus.className = 'stat-value';
    } else {
      statStatus.textContent = '待上传'; statStatus.className = 'stat-value';
    }
  }

  function populateSheetSelect(wb) {
    sheetSelect.innerHTML = '';
    wb.SheetNames.forEach(function(name) {
      var opt = document.createElement('option');
      opt.value = name; opt.textContent = name;
      sheetSelect.appendChild(opt);
    });
    var preferred = wb.SheetNames.find(function(n) {
      return n.includes('美国') || n.includes('休斯顿') || n.includes('订单');
    });
    if (preferred) sheetSelect.value = preferred;
    sheetSelectWrap.style.display = wb.SheetNames.length > 1 ? 'block' : 'none';
  }

  function handleFiles(fileList) {
    var files = Array.from(fileList).filter(function(f) { return f.name.toLowerCase().endsWith('.xlsx'); });
    if (files.length === 0) { setStatus(uploadStatus, '请选择 .xlsx 格式文件', 'error'); return; }
    setStatus(uploadStatus, '识别文件中…', '');

    var chain = Promise.resolve();
    files.forEach(function(file) {
      chain = chain.then(function() {
        return readWorkbook(file).then(function(wb) {
          var role = detectFileRole(wb);
          if (role === 'base') {
            baseFile = { name: file.name, file: file }; baseWb = wb;
            populateSheetSelect(wb);
            addLog('基准表已识别：' + file.name, 'ok');
          } else if (role === 'income') {
            incomeFile = { name: file.name, file: file }; incomeWb = wb;
            addLog('C 端收入表已识别：' + file.name, 'ok');
          } else {
            if (!baseFile) {
              baseFile = { name: file.name, file: file }; baseWb = wb;
              populateSheetSelect(wb);
              addLog(file.name + ' 无法自动识别，已暂存为基准表', 'warn');
            } else if (!incomeFile) {
              incomeFile = { name: file.name, file: file }; incomeWb = wb;
              addLog(file.name + ' 无法自动识别，已暂存为 C 端收入表', 'warn');
            }
          }
        }).catch(function(e) {
          setStatus(uploadStatus, file.name + ' 读取失败：' + e.message, 'error');
        });
      });
    });

    chain.then(function() {
      updateSlotUI();
      if (baseFile && incomeFile) {
        setStatus(uploadStatus, '两个文件均已识别，可以开始匹配', 'success');
      } else if (baseFile) {
        setStatus(uploadStatus, '基准表已就绪，还需上传 C 端收入表', 'warn');
      } else if (incomeFile) {
        setStatus(uploadStatus, 'C 端收入表已就绪，还需上传基准表', 'warn');
      }
    });
  }

  chooseFilesBtn.addEventListener('click', function(e) {
    e.stopPropagation(); multiFileInput.value = ''; multiFileInput.click();
  });
  uploadMainArea.addEventListener('click', function() { multiFileInput.value = ''; multiFileInput.click(); });
  multiFileInput.addEventListener('change', function(e) { handleFiles(e.target.files || []); });

  ['dragenter','dragover'].forEach(function(ev) {
    uploadMainArea.addEventListener(ev, function(e) { e.preventDefault(); uploadMainArea.classList.add('dragover'); });
  });
  ['dragleave','drop'].forEach(function(ev) {
    uploadMainArea.addEventListener(ev, function(e) { e.preventDefault(); uploadMainArea.classList.remove('dragover'); });
  });
  uploadMainArea.addEventListener('drop', function(e) { handleFiles(e.dataTransfer.files || []); });

  slotBaseRemove.addEventListener('click', function(e) {
    e.stopPropagation();
    baseFile = null; baseWb = null;
    sheetSelectWrap.style.display = 'none';
    setStatus(uploadStatus, '基准表已移除', '');
    statTotal.textContent = '—'; statNeed.textContent = '—';
    updateSlotUI(); addLog('基准表已移除', 'warn');
  });

  slotIncomeRemove.addEventListener('click', function(e) {
    e.stopPropagation();
    incomeFile = null; incomeWb = null;
    setStatus(uploadStatus, 'C 端收入表已移除', '');
    updateSlotUI(); addLog('C 端收入表已移除', 'warn');
  });

  resetAllBtn.addEventListener('click', function() {
    baseFile = null; baseWb = null;
    incomeFile = null; incomeWb = null;
    multiFileInput.value = '';
    outMainWb = null; outRemindWb = null;
    sheetSelectWrap.style.display = 'none';
    downloadMainBtn.disabled = true; downloadRemindBtn.disabled = true;
    setProgress(0, ''); setStatus(uploadStatus, '', '');
    clearLogs(); resetResultUI(); updateSlotUI();
  });

  function resetResultUI() {
    rNeed.textContent = '—'; rRound1.textContent = '—';
    rRound2.textContent = '—'; rUnmatched.textContent = '—';
    remindCount.textContent = '0 条';
    remindBody.innerHTML = '<tr class="empty-row"><td colspan="4">匹配完成后显示</td></tr>';
    previewBody.innerHTML = '<tr class="empty-row"><td colspan="4">匹配完成后显示</td></tr>';
    statTotal.textContent = '—'; statNeed.textContent = '—';
  }

  matchBtn.addEventListener('click', runMatch);

  function runMatch() {
    if (!baseWb || !incomeWb) return;
    processing = true; updateMatchBtn(); updateStatStatus();
    setProgress(5, '开始处理...'); clearLogs(); addLog('开始匹配任务', '');

    setTimeout(function() {
      try {
        var selectedSheet = sheetSelect.value || baseWb.SheetNames[0];
        setProgress(10, '读取基准表工作表：' + selectedSheet);
        addLog('读取基准表 Sheet：' + selectedSheet);

        var baseWsObj = baseWb.Sheets[selectedSheet];
        if (!baseWsObj) throw new Error('找不到工作表：' + selectedSheet);

        var baseAoa = XLSX.utils.sheet_to_json(baseWsObj, { header: 1, defval: null });
        var headerIdx = 0;
        for (var i = 0; i < Math.min(10, baseAoa.length); i++) {
          var hrow = (baseAoa[i] || []).map(function(v) { return typeof v === 'string' ? v.trim() : v; });
          if (hrow.includes('卖家订单号') || hrow.includes('订单收入')) { headerIdx = i; break; }
        }

        var headers = (baseAoa[headerIdx] || []).map(function(v) { return typeof v === 'string' ? v.trim() : String(v != null ? v : ''); });
        var colOrderNo = headers.indexOf('卖家订单号');
        var colIncomeIdx = headers.indexOf('订单收入');

        if (colOrderNo === -1) throw new Error('基准表未找到「卖家订单号」列');
        if (colIncomeIdx === -1) throw new Error('基准表未找到「订单收入」列');

        addLog('列映射 → 卖家订单号: 第' + (colOrderNo+1) + '列，订单收入: 第' + (colIncomeIdx+1) + '列', 'ok');

        var dataRows = [];
        for (var di = headerIdx + 1; di < baseAoa.length; di++) {
          var row = baseAoa[di] || [];
          var orderNo = row[colOrderNo];
          var income  = row[colIncomeIdx];
          var orderNoStr = (orderNo == null || orderNo === '') ? null : String(orderNo).trim();
          dataRows.push({ origIdx: di, orderNo: orderNoStr, income: income, source: 'original' });
        }

        statTotal.textContent = dataRows.length.toLocaleString();
        addLog('基准表数据行数：' + dataRows.length, 'ok');

        var needFillIdx = dataRows.map(function(r,i){return i;}).filter(function(i){ return dataRows[i].orderNo && isBlank(dataRows[i].income); });
        var hasIncomeIdx = dataRows.map(function(r,i){return i;}).filter(function(i){ return dataRows[i].orderNo && !isBlank(dataRows[i].income); });

        statNeed.textContent = needFillIdx.length.toLocaleString();
        addLog('需填充行：' + needFillIdx.length + '，已有收入行：' + hasIncomeIdx.length, 'ok');

        setProgress(25, '读取 C 端收入表...');

        var incomeWsName = incomeWb.SheetNames.find(function(n) {
          var ws2 = incomeWb.Sheets[n];
          var rows2 = XLSX.utils.sheet_to_json(ws2, { header: 1, defval: null, range: 0 });
          var firstRow = (rows2[0] || []).map(function(v) { return typeof v === 'string' ? v.trim() : v; });
          return firstRow.includes('卖家订单号') || firstRow.includes('人民币');
        });
        if (!incomeWsName) incomeWsName = incomeWb.SheetNames[0];

        addLog('C 端收入 Sheet：' + incomeWsName);
        var incomeWsObj = incomeWb.Sheets[incomeWsName];
        var incomeRows = normalizeRows(sheetToRows(incomeWsObj));
        addLog('C 端收入表行数：' + incomeRows.length, 'ok');

        if (incomeRows.length > 0) {
          var sampleRow = incomeRows[0];
          if (!('卖家订单号' in sampleRow)) throw new Error('C 端收入表未找到「卖家订单号」列');
          if (!('人民币' in sampleRow))    throw new Error('C 端收入表未找到「人民币」列');
        }

        setProgress(40, '构建匹配字典...');
        var cMapDirect = {}, cMapNoDash = {};
        incomeRows.forEach(function(r) {
          var key = String(r['卖家订单号'] != null ? r['卖家订单号'] : '').trim();
          if (!key) return;
          var amt = toNum(r['人民币']);
          if (amt == null) return;
          if (!(key in cMapDirect)) cMapDirect[key] = amt;
          var keyND = stripDash(key);
          if (keyND && !(keyND in cMapNoDash)) cMapNoDash[keyND] = amt;
        });

        addLog('直接匹配字典：' + Object.keys(cMapDirect).length + ' 条', 'ok');
        addLog('去「-」字典：' + Object.keys(cMapNoDash).length + ' 条', 'ok');

        setProgress(55, '执行第一轮匹配...');
        var round1Count = 0, stillEmpty = [];
        needFillIdx.forEach(function(i) {
          var orderNo2 = dataRows[i].orderNo;
          if (orderNo2 in cMapDirect) {
            dataRows[i].income = cMapDirect[orderNo2]; dataRows[i].source = 'round1'; round1Count++;
          } else { stillEmpty.push(i); }
        });
        addLog('第一轮匹配完成：成功 ' + round1Count + ' 条，剩余 ' + stillEmpty.length + ' 条', 'ok');

        setProgress(70, '执行第二轮匹配（去「-」）...');
        var round2Count = 0, unmatched = [];
        stillEmpty.forEach(function(i) {
          var orderNo2 = dataRows[i].orderNo;
          var keyND = stripDash(orderNo2);
          if (keyND && keyND in cMapNoDash) {
            dataRows[i].income = cMapNoDash[keyND]; dataRows[i].source = 'round2'; round2Count++;
          } else { unmatched.push(i); }
        });
        addLog('第二轮匹配完成：新增 ' + round2Count + ' 条，最终未匹配 ' + unmatched.length + ' 条', round2Count > 0 ? 'ok' : 'warn');

        setProgress(82, '生成跨月收入提醒...');
        var remindRows = [];
        hasIncomeIdx.forEach(function(i) {
          var orderNo2 = dataRows[i].orderNo;
          var origIncome2 = toNum(dataRows[i].income);
          var keyND = stripDash(orderNo2);
          var newAmt = null;
          if (orderNo2 in cMapDirect) newAmt = cMapDirect[orderNo2];
          else if (keyND in cMapNoDash) newAmt = cMapNoDash[keyND];
          if (newAmt !== null) remindRows.push({ orderNo: orderNo2, origIncome: origIncome2, newAmt: newAmt, diff: newAmt - (origIncome2 || 0) });
        });
        addLog('跨月收入提醒：' + remindRows.length + ' 条', remindRows.length > 0 ? 'warn' : 'ok');

        setProgress(90, '生成输出文件...');
        var mainWb2 = XLSX.utils.book_new();
        var mainAoA = [['卖家订单号', '订单收入']];
        dataRows.forEach(function(r) { mainAoA.push([r.orderNo != null ? r.orderNo : null, isBlank(r.income) ? null : toNum(r.income)]); });
        var mainWs2 = XLSX.utils.aoa_to_sheet(mainAoA);
        mainWs2['!cols'] = [{ wch: 30 }, { wch: 15 }];
        XLSX.utils.book_append_sheet(mainWb2, mainWs2, '收入匹配完成');
        outMainWb = mainWb2;

        var remindWb2 = XLSX.utils.book_new();
        var remindAoA = [['卖家订单号', '原有金额', '新匹配金额']];
        remindRows.forEach(function(r) { remindAoA.push([r.orderNo, r.origIncome, r.newAmt]); });
        var remindWs2 = XLSX.utils.aoa_to_sheet(remindAoA);
        remindWs2['!cols'] = [{ wch: 30 }, { wch: 15 }, { wch: 15 }];
        XLSX.utils.book_append_sheet(remindWb2, remindWs2, '跨月收入提醒');
        outRemindWb = remindWb2;

        setProgress(100, '匹配完成 ✓');
        addLog('所有文件已生成，可下载', 'ok');

        rNeed.textContent = needFillIdx.length.toLocaleString();
        rRound1.textContent = round1Count.toLocaleString();
        rRound2.textContent = round2Count.toLocaleString();
        rUnmatched.textContent = unmatched.length.toLocaleString();
        remindCount.textContent = remindRows.length + ' 条';

        if (remindRows.length === 0) {
          remindBody.innerHTML = '<tr class="empty-row"><td colspan="4">无跨月收入提醒记录</td></tr>';
        } else {
          remindBody.innerHTML = '';
          remindRows.slice(0, 100).forEach(function(r) {
            var tr = document.createElement('tr');
            tr.innerHTML = '<td class="td-mono">' + r.orderNo + '</td><td class="td-right">' + fmt2(r.origIncome) + '</td><td class="td-right">' + fmt2(r.newAmt) + '</td><td class="td-right" style="color:' + (r.diff > 0 ? 'var(--success)' : r.diff < 0 ? 'var(--danger)' : '') + '">' + fmt2(r.diff) + '</td>';
            remindBody.appendChild(tr);
          });
          if (remindRows.length > 100) {
            var tr2 = document.createElement('tr');
            tr2.innerHTML = '<td colspan="4" class="td-center" style="color:var(--muted)">… 仅展示前 100 条</td>';
            remindBody.appendChild(tr2);
          }
        }

        previewBody.innerHTML = '';
        var previewCnt = 0;
        for (var pi = 0; pi < dataRows.length && previewCnt < 50; pi++) {
          var pr = dataRows[pi];
          if (!pr.orderNo) continue;
          var tr3 = document.createElement('tr');
          var src = pr.source === 'round1' ? '<span style="color:var(--success)">一轮</span>'
            : pr.source === 'round2' ? '<span style="color:#0284C7">二轮</span>'
            : pr.income != null ? '<span style="color:var(--muted)">原有</span>'
            : '<span style="color:var(--danger)">未匹配</span>';
          tr3.innerHTML = '<td class="td-center" style="color:var(--muted)">' + (pi+1) + '</td><td class="td-mono">' + pr.orderNo + '</td><td class="td-right">' + (isBlank(pr.income) ? '—' : fmt2(pr.income)) + '</td><td class="td-center">' + src + '</td>';
          previewBody.appendChild(tr3);
          previewCnt++;
        }
        if (previewCnt === 0) previewBody.innerHTML = '<tr class="empty-row"><td colspan="4">没有含卖家订单号的行</td></tr>';

        downloadMainBtn.disabled = false; downloadRemindBtn.disabled = false;

      } catch(err) {
        setProgress(0, '');
        setStatus(processStatus, '处理失败：' + err.message, 'error');
        addLog('处理失败：' + err.message, 'err');
      }

      processing = false; updateMatchBtn(); updateStatStatus();
    }, 20);
  }

  downloadMainBtn.addEventListener('click', function() {
    if (!outMainWb) return;
    downloadWb(outMainWb, '收入匹配完成.xlsx');
    addLog('已下载：收入匹配完成.xlsx', 'ok');
  });

  downloadRemindBtn.addEventListener('click', function() {
    if (!outRemindWb) return;
    downloadWb(outRemindWb, '跨月收入提醒.xlsx');
    addLog('已下载：跨月收入提醒.xlsx', 'ok');
  });

  helpBtnMatch.addEventListener('click', function() { helpModalMatch.classList.add('open'); });
  closeHelpMatch.addEventListener('click', function() { helpModalMatch.classList.remove('open'); });
  helpModalMatch.addEventListener('click', function(e) { if (e.target === helpModalMatch) helpModalMatch.classList.remove('open'); });

  window.addEventListener('load', function() {
    if (typeof XLSX === 'undefined') {
      setStatus(processStatus, '⚠ 未检测到 xlsx.full.min.js，请将其与本文件放在同一目录', 'error');
      addLog('xlsx.full.min.js 未加载，功能不可用', 'err');
      matchBtn.disabled = true;
    }
  });
})();

/* ============================================================
   模块2: 海外仓账单自动化上传与核算
   ============================================================ */
(function() {
  var templateInput  = document.getElementById('templateInput');
  var dataInput      = document.getElementById('dataInput');
  var templateBtn    = document.getElementById('templateBtn');
  var dataBtn        = document.getElementById('dataBtn');
  var templateReset  = document.getElementById('templateReset');
  var templateList   = document.getElementById('templateList');
  var dataList       = document.getElementById('dataList');
  var templateStatus = document.getElementById('templateStatus');
  var dataStatus     = document.getElementById('dataStatus');
  var processBtn     = document.getElementById('processBtn');
  var processStatus2 = document.getElementById('processStatus2');
  var progressBar2   = document.getElementById('progressBar2');
  var downloadBtn    = document.getElementById('downloadBtn');
  var logBox2        = document.getElementById('logBox2');
  var clearCacheBtn  = document.getElementById('clearCacheBtn');
  var newTaskBtn     = document.getElementById('newTaskBtn');
  var previewBody2   = document.getElementById('previewBody2');
  var missSkuCount   = document.getElementById('missSkuCount');
  var missQuoteCount = document.getElementById('missQuoteCount');
  var missRateCount  = document.getElementById('missRateCount');
  var missingSkuBody   = document.getElementById('missingSkuBody');
  var missingQuoteBody = document.getElementById('missingQuoteBody');
  var missingRateBody  = document.getElementById('missingRateBody');
  var statFiles      = document.getElementById('statFiles');
  var statOrders     = document.getElementById('statOrders');
  var statStatus2    = document.getElementById('statStatus2');
  var templateDrop   = document.getElementById('templateDrop');
  var dataDrop       = document.getElementById('dataDrop');
  var helpBtnBilling = document.getElementById('helpBtnBilling');
  var helpModalBilling = document.getElementById('helpModalBilling');
  var closeHelpBilling = document.getElementById('closeHelpBilling');
  var keywordBtn     = document.getElementById('keywordBtn');
  var keywordModal   = document.getElementById('keywordModal');
  var keywordList    = document.getElementById('keywordList');
  var keywordInput   = document.getElementById('keywordInput');
  var keywordAddBtn  = document.getElementById('keywordAddBtn');
  var keywordResetBtn= document.getElementById('keywordResetBtn');
  var keywordCloseBtn= document.getElementById('keywordCloseBtn');
  var keywordInline  = document.getElementById('keywordInline');

  var templateFile  = null;
  var dataFiles     = [];
  var processing    = false;
  var processed     = false;
  var outputWorkbook = null;
  var outputFilename = '输出结果.xlsx';

  var CACHE_KEY     = 'table_uploader_cache';
  var CACHE_TTL     = 2 * 60 * 60 * 1000;
  var DEFAULT_KEYWORDS = ['林嘉楠', '智博', '拓泽瑞', '澄晞'];
  var KEYWORDS_STORAGE = 'manager_keywords_v1';
  var KEYWORDS = DEFAULT_KEYWORDS.slice();

  function setStatus2(el, msg, type) { el.textContent = msg; el.className = 'status' + (type ? ' ' + type : ''); }
  function setProgress2(p, msg) { progressBar2.style.width = p + '%'; if (msg) setStatus2(processStatus2, msg, ''); }

  function validateXlsx(file) { return file && file.name.toLowerCase().endsWith('.xlsx'); }

  function uniqKeywords(arr) {
    var set = new Set(), out = [];
    arr.forEach(function(v) { var s = String(v||'').trim(); if (!s || set.has(s)) return; set.add(s); out.push(s); });
    return out;
  }

  function loadKeywords() {
    try {
      var raw = localStorage.getItem(KEYWORDS_STORAGE);
      if (!raw) return DEFAULT_KEYWORDS.slice();
      var data = JSON.parse(raw);
      if (Array.isArray(data)) return uniqKeywords(data);
    } catch(_) {}
    return DEFAULT_KEYWORDS.slice();
  }

  function saveKeywords() {
    KEYWORDS = uniqKeywords(KEYWORDS);
    localStorage.setItem(KEYWORDS_STORAGE, JSON.stringify(KEYWORDS));
    renderKeywords();
  }

  function renderKeywords() {
    if (!keywordList || !keywordInline) return;
    keywordList.innerHTML = '';
    keywordInline.textContent = KEYWORDS.length === 0 ? '客户经理关键字：暂无（请添加）' : '客户经理关键字：' + KEYWORDS.join('、');
    KEYWORDS.forEach(function(k) {
      var pill = document.createElement('div'); pill.className = 'kw-pill';
      var span = document.createElement('span'); span.textContent = k;
      var btn2 = document.createElement('button'); btn2.type='button'; btn2.textContent='×';
      btn2.addEventListener('click', function() { KEYWORDS = KEYWORDS.filter(function(x){return x!==k;}); saveKeywords(); });
      pill.appendChild(span); pill.appendChild(btn2); keywordList.appendChild(pill);
    });
  }

  function addKeywordFromInput() {
    var val = keywordInput.value.trim();
    if (!val) return;
    KEYWORDS.push(val); keywordInput.value = ''; saveKeywords();
  }

  function inferType(name) {
    if (name.includes('模板')||name.includes('模版')) return '模板';
    if (name.includes('账单')) return '账单信息表';
    if (name.includes('订单')) return '订单信息表';
    if (name.includes('支出')) return '支出情况表';
    if (name.includes('收入')) return '收入情况表';
    if (name.includes('汇率')) return '汇率表';
    if (name.includes('产品成本')) return '产品成本表';
    if (name.includes('第一档')) return '第一档收入表';
    return '补充表';
  }

  function updateProcessButton() {
    var actualFiles = dataFiles.filter(function(f){return f.file;});
    processBtn.disabled = !(templateFile && templateFile.file && actualFiles.length > 0);
  }

  function updateStats2(orderCountOverride) {
    var fileCount = (templateFile ? 1 : 0) + dataFiles.length;
    statFiles.textContent = fileCount;
    statOrders.textContent = orderCountOverride != null ? orderCountOverride : dataFiles.length;
    if (processing) statStatus2.textContent = '处理中';
    else if (processed) statStatus2.textContent = '已完成';
    else if (fileCount > 0) statStatus2.textContent = '待处理';
    else statStatus2.textContent = '待上传';
  }

  function cacheState() {
    var payload = { template: templateFile ? {name:templateFile.name} : null, data: dataFiles.map(function(f){return {name:f.name};}), ts: Date.now() };
    localStorage.setItem(CACHE_KEY, JSON.stringify(payload));
  }

  function renderTemplateList() {
    templateList.innerHTML = '';
    if (!templateFile) return;
    var item = document.createElement('div'); item.className = 'file-item';
    var fname = document.createElement('div'); fname.className = 'fname'; fname.textContent = templateFile.name + (templateFile.cached ? '（已缓存）' : '');
    var ftag = document.createElement('div'); ftag.className = 'ftag'; ftag.textContent = inferType(templateFile.name);
    var fremove = document.createElement('div'); fremove.className = 'fremove'; fremove.textContent = '删除';
    fremove.addEventListener('click', function() {
      templateFile = null; templateList.innerHTML = ''; templateReset.disabled = true;
      updateProcessButton(); updateStats2(); cacheState();
    });
    item.append(fname, ftag, fremove); templateList.appendChild(item);
  }

  function renderDataList() {
    dataList.innerHTML = '';
    dataFiles.forEach(function(entry) {
      var item = document.createElement('div'); item.className = 'file-item';
      var fname = document.createElement('div'); fname.className = 'fname'; fname.textContent = entry.name + (entry.cached ? '（已缓存）' : '');
      var ftag = document.createElement('div'); ftag.className = 'ftag'; ftag.textContent = inferType(entry.name);
      var fremove = document.createElement('div'); fremove.className = 'fremove'; fremove.textContent = '删除';
      (function(e) {
        fremove.addEventListener('click', function() {
          dataFiles = dataFiles.filter(function(f){return f.name !== e.name;});
          item.remove(); updateProcessButton(); updateStats2(); cacheState();
        });
      })(entry);
      item.append(fname, ftag, fremove); dataList.appendChild(item);
    });
  }

  function hasCachedFiles() { return (templateFile && templateFile.cached) || dataFiles.some(function(f){return f.cached;}); }

  function clearCachedEntries(showMessage) {
    localStorage.removeItem(CACHE_KEY);
    if (templateFile && templateFile.cached) { templateFile = null; templateReset.disabled = true; setStatus2(templateStatus,'',''); }
    dataFiles = dataFiles.filter(function(f){return !f.cached;});
    renderTemplateList(); renderDataList();
    if (showMessage) setStatus2(dataStatus, '缓存已清空', 'success');
    updateProcessButton(); updateStats2();
  }

  function resetAllState() {
    localStorage.removeItem(CACHE_KEY);
    templateFile = null; dataFiles = [];
    templateInput.value = ''; dataInput.value = '';
    templateReset.disabled = true;
    renderTemplateList(); renderDataList();
    processing = false; processed = false;
    outputWorkbook = null; outputFilename = '输出结果.xlsx';
    progressBar2.style.width = '0%'; downloadBtn.disabled = true;
    setStatus2(templateStatus, '', '');
    setStatus2(dataStatus, '已开始新建任务，缓存已清空', 'success');
    setStatus2(processStatus2, '', '');
    logBox2.textContent = '暂无日志';
    clearMissingLists(); updateProcessButton(); updateStats2();
  }

  function restoreCache() {
    var raw = localStorage.getItem(CACHE_KEY);
    if (!raw) return;
    try {
      var payload = JSON.parse(raw);
      if (Date.now() - payload.ts > CACHE_TTL) return;
      templateFile = null; dataFiles = [];
      if (payload.template) { templateFile = {name:payload.template.name, file:null, cached:true}; templateReset.disabled = false; }
      if (payload.data) {
        payload.data.forEach(function(d) {
          if (!dataFiles.some(function(f){return f.name===d.name;})) dataFiles.push({name:d.name, file:null, cached:true});
        });
      }
      renderTemplateList(); renderDataList();
      if (templateFile || dataFiles.length > 0) setStatus2(dataStatus, '已恢复缓存文件，开始新建任务建议先点"新建任务"', '');
      updateProcessButton(); updateStats2();
    } catch(_) {}
  }

  function handleTemplateFiles(files) {
    var file = files[0]; if (!file) return;
    if (!validateXlsx(file)) { setStatus2(templateStatus, '请上传 .xlsx 格式的模板文件', 'error'); return; }
    if (hasCachedFiles()) clearCachedEntries(false);
    templateFile = {name:file.name, file:file, cached:false}; templateReset.disabled = false;
    renderTemplateList(); setStatus2(templateStatus, file.name + ' 上传成功', 'success');
    updateProcessButton(); updateStats2(); cacheState();
  }

  function handleDataFiles(files) {
    if (files.length === 0) return;
    if (hasCachedFiles()) clearCachedEntries(false);
    var added = 0, replaced = 0;
    files.forEach(function(file) {
      if (!validateXlsx(file)) { setStatus2(dataStatus, '仅支持 .xlsx 文件', 'error'); return; }
      var idx = dataFiles.findIndex(function(f){return f.name===file.name;});
      if (idx >= 0) { dataFiles[idx] = {name:file.name, file:file, cached:false}; replaced++; }
      else { dataFiles.push({name:file.name, file:file, cached:false}); added++; }
    });
    renderDataList();
    if (added > 0 && replaced > 0) setStatus2(dataStatus, added + ' 个文件上传成功，' + replaced + ' 个同名文件已替换', 'success');
    else if (added > 0) setStatus2(dataStatus, added + ' 个文件上传成功', 'success');
    else if (replaced > 0) setStatus2(dataStatus, replaced + ' 个同名文件已替换', 'success');
    updateProcessButton(); updateStats2(); cacheState();
  }

  templateBtn.addEventListener('click', function(){ templateInput.click(); });
  dataBtn.addEventListener('click', function(){ dataInput.click(); });
  clearCacheBtn.addEventListener('click', function(){ clearCachedEntries(true); });
  newTaskBtn.addEventListener('click', function(){ resetAllState(); });
  templateInput.addEventListener('change', function(e){ handleTemplateFiles(Array.from(e.target.files||[])); });
  templateReset.addEventListener('click', function(){ templateInput.value=''; templateFile=null; templateReset.disabled=true; setStatus2(templateStatus,'',''); renderTemplateList(); updateProcessButton(); updateStats2(); cacheState(); });
  dataInput.addEventListener('change', function(e){ handleDataFiles(Array.from(e.target.files||[])); });

  keywordBtn.addEventListener('click', function(){ KEYWORDS=loadKeywords(); renderKeywords(); keywordModal.classList.add('open'); });
  keywordCloseBtn.addEventListener('click', function(){ keywordModal.classList.remove('open'); });
  keywordModal.addEventListener('click', function(e){ if(e.target===keywordModal) keywordModal.classList.remove('open'); });
  keywordAddBtn.addEventListener('click', addKeywordFromInput);
  keywordResetBtn.addEventListener('click', function(){ KEYWORDS=DEFAULT_KEYWORDS.slice(); saveKeywords(); });
  keywordInput.addEventListener('keydown', function(e){ if(e.key==='Enter'){ e.preventDefault(); addKeywordFromInput(); } });

  function setupDrag2(dropArea, handler) {
    ['dragenter','dragover'].forEach(function(evt){ dropArea.addEventListener(evt,function(e){e.preventDefault();dropArea.classList.add('dragover');}); });
    ['dragleave','drop'].forEach(function(evt){ dropArea.addEventListener(evt,function(e){e.preventDefault();dropArea.classList.remove('dragover');}); });
    dropArea.addEventListener('drop', function(e){ handler(Array.from(e.dataTransfer.files||[])); });
  }
  setupDrag2(templateDrop, function(files){ if(files.length>1){setStatus2(templateStatus,'模板仅支持单文件上传','error');return;} handleTemplateFiles(files); });
  setupDrag2(dataDrop, handleDataFiles);

  helpBtnBilling.addEventListener('click', function(){ helpModalBilling.classList.add('open'); });
  closeHelpBilling.addEventListener('click', function(){ helpModalBilling.classList.remove('open'); });
  helpModalBilling.addEventListener('click', function(e){ if(e.target===helpModalBilling) helpModalBilling.classList.remove('open'); });

  /* ── 工具函数 ── */
  function isEmpty2(val){ return val===null||val===undefined||val===''; }
  function isBlank2(val){ return isEmpty2(val)||(typeof val==='string'&&val.trim()===''); }
  function pickValue(){ for(var i=0;i<arguments.length;i++){if(!isBlank2(arguments[i]))return arguments[i];}return null; }
  function setIfBlank2(row,key,val){ if(!row||!key)return; if(isBlank2(row[key])&&!isBlank2(val))row[key]=val; }
  function normalizeKey2(val){ if(val===null||val===undefined)return ''; return String(val).replace(/[\s\u200b\u200d\ufeff-]/g,'').trim(); }
  function stripOrderSuffix2(val){ if(val===null||val===undefined)return ''; return String(val).replace(/-[^-]*$/,'').trim(); }
  function toNumber2(val){ if(val===null||val===undefined||val==='')return null; if(typeof val==='number')return val; var s=String(val).replace(/,/g,'').trim(); var m=s.match(/-?\d+(\.\d+)?/); return m?parseFloat(m[0]):null; }
  function round2fn(val){ if(val===null||val===undefined||isNaN(val))return null; return Math.round(val*100)/100; }
  function format2fn(val){ if(val===null||val===undefined||val==='')return ''; var num=Number(val); if(isNaN(num))return String(val); return num.toFixed(2); }
  function normalizeCurrency2(cur){ if(!cur)return null; var s=String(cur).trim(); if(['人民币','RMB','CNY','¥','￥'].includes(s))return '人民币'; if(['USD','US$','$','美元'].includes(s))return '美元'; return s; }
  function parseAmount2(val){ if(val===null||val===undefined||val==='')return{amount:null,currency:null}; if(typeof val==='number')return{amount:val,currency:'人民币'}; var s=String(val).trim(); if(s==='')return{amount:null,currency:null}; var currency=null; if(s.includes('$')||s.includes('USD'))currency='美元'; if(s.includes('人民币')||s.includes('¥')||s.includes('￥'))currency='人民币'; var m=s.replace(/,/g,'').match(/-?\d+(\.\d+)?/); return{amount:m?parseFloat(m[0]):null,currency:currency}; }
  function qtyLabel2(qty){ if(qty===null||qty===undefined)return null; var q=Number(qty); if(isNaN(q))return null; if(q>=4)return '≥4pcs'; if(q===3)return '3 pcs'; if(q===2)return '2 pcs'; if(q===1)return '1pc'; return null; }
  function hasKeyword2(text){ if(!text)return false; var s=String(text); return KEYWORDS.some(function(k){return s.includes(k);}); }
  function readWorkbook2(file){ return file.arrayBuffer().then(function(buf){return XLSX.read(buf,{type:'array'});}); }
  function sheetToRows2(sheet){ return XLSX.utils.sheet_to_json(sheet,{defval:null}); }
  function normalizeRows2(rows){ return rows.map(function(row){var out={};Object.keys(row).forEach(function(k){out[typeof k==='string'?k.trim():k]=row[k];});return out;}); }
  function rowHasHeader2(row, header){ return (row||[]).some(function(v){ return (typeof v==='string'?v.trim():v)===header; }); }
  function sheetLooksLikeOrder2(sheet){
    if(!sheet) return false;
    var aoa = XLSX.utils.sheet_to_json(sheet,{header:1,defval:null});
    var limit = Math.min(aoa.length, 8);
    for(var i=0;i<limit;i++){
      var row = aoa[i] || [];
      if(!rowHasHeader2(row,'订单编号')) continue;
      if(rowHasHeader2(row,'卖家订单号') || rowHasHeader2(row,'产品SKU') || rowHasHeader2(row,'客户经理') || rowHasHeader2(row,'订单状态')) return true;
    }
    return false;
  }

  function getSheetOrThrow2(wb, name, label) {
    var names = Array.isArray(name) ? name : [name];
    var sheetNames = wb.SheetNames || [];
    var nameMap = new Map(sheetNames.map(function(n){return[String(n).trim(),n];}));
    for(var i=0;i<names.length;i++){ var key=String(names[i]).trim(); if(nameMap.has(key))return wb.Sheets[nameMap.get(key)]; }
    for(var j=0;j<names.length;j++){ var key2=String(names[j]).trim(); var hit=sheetNames.find(function(s){return String(s).includes(key2);}); if(hit)return wb.Sheets[hit]; }
    if(label){ var hit2=sheetNames.find(function(s){return String(s).includes(label);}); if(hit2)return wb.Sheets[hit2]; }
    if(label==='订单信息'){
      for(var k=0;k<sheetNames.length;k++){
        var sn = sheetNames[k];
        var sheet = wb.Sheets[sn];
        if(sheetLooksLikeOrder2(sheet)) return sheet;
      }
    }
    throw new Error((label||'')+'缺少工作表：'+names.join(' / '));
  }

  function findHeaderRowIndex2(aoa) {
    var limit = Math.min(aoa.length,10);
    for(var i=0;i<limit;i++){ var row=(aoa[i]||[]).map(function(v){return typeof v==='string'?v.trim():v;}); if(row.includes('订单编号')&&row.includes('国家'))return i; }
    return 0;
  }

  function buildOutputWorkbook2(templateWb, resultRows, headerRowIndex) {
    var sheetName=templateWb.SheetNames[0]; var ws=templateWb.Sheets[sheetName];
    var headerRows=XLSX.utils.sheet_to_json(ws,{header:1,defval:null});
    var idx=headerRowIndex!=null?headerRowIndex:1;
    var headers=(headerRows[idx]||headerRows[0]||[]);
    var dataAoA=resultRows.map(function(row){return headers.map(function(h){return row[h]!=null?row[h]:null;});});
    XLSX.utils.sheet_add_aoa(ws,dataAoA,{origin:{r:idx+1,c:0}});
    return templateWb;
  }

  function downloadWorkbook2(wb, filename) {
    var wbout=XLSX.write(wb,{bookType:'xlsx',type:'array'});
    var blob=new Blob([wbout],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
    var url=URL.createObjectURL(blob); var a=document.createElement('a');
    a.href=url; a.download=filename; document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
  }

  function renderPreview2(rows) {
    previewBody2.innerHTML = '';
    rows.slice(0,30).forEach(function(r) {
      var tr=document.createElement('tr');
      var cols=['订单编号','国家','期间','订单收入','订单支出','毛利','盈亏情况'];
      cols.forEach(function(c,idx2) {
        var td=document.createElement('td');
        if(idx2>=3&&idx2<=5) td.className='align-right'; else td.className='align-left';
        td.textContent=(idx2>=3&&idx2<=5)?format2fn(r[c]):(r[c]!=null?r[c]:'');
        tr.appendChild(td);
      });
      previewBody2.appendChild(tr);
    });
  }

  function clearMissingLists() {
    if(!missingSkuBody)return;
    missingSkuBody.innerHTML=''; missingQuoteBody.innerHTML=''; missingRateBody.innerHTML='';
    missSkuCount.textContent='0'; missQuoteCount.textContent='0'; missRateCount.textContent='0';
  }

  function renderMissingLists2(missing) {
    if(!missingSkuBody)return;
    var skuRows=[];
    missing.sku.forEach(function(set,orderId){ if(!set||set.size===0)return; skuRows.push({orderId:orderId,skus:Array.from(set).join('、')}); });
    missSkuCount.textContent=skuRows.length; missingSkuBody.innerHTML='';
    skuRows.slice(0,200).forEach(function(row){ var tr=document.createElement('tr'); tr.innerHTML='<td>'+row.orderId+'</td><td>'+row.skus+'</td>'; missingSkuBody.appendChild(tr); });
    var quoteRows=missing.quote||[]; missQuoteCount.textContent=quoteRows.length; missingQuoteBody.innerHTML='';
    quoteRows.slice(0,200).forEach(function(row){ var tr=document.createElement('tr'); tr.innerHTML='<td>'+row.orderId+'</td><td>'+(row.country||'')+'</td><td>'+(row.qty||'')+'</td><td>'+(row.method||'')+'</td><td>'+(row.band||'')+'</td><td>'+(row.reason||'')+'</td>'; missingQuoteBody.appendChild(tr); });
    var rateRows=Array.from(missing.rate.values()); missRateCount.textContent=rateRows.length; missingRateBody.innerHTML='';
    rateRows.slice(0,200).forEach(function(row){ var tr=document.createElement('tr'); tr.innerHTML='<td>'+(row.period||'')+'</td><td>'+(row.currency||'')+'</td><td>'+(row.scene||'')+'</td>'; missingRateBody.appendChild(tr); });
  }

  processBtn.addEventListener('click', async function() {
    if (!window.XLSX) { setStatus2(processStatus2,'未加载 XLSX 库，请确认同目录存在 xlsx.full.min.js','error'); return; }
    if (!templateFile||!templateFile.file) { setStatus2(processStatus2,'请重新上传模板文件（缓存文件不能计算）','error'); return; }
    var actualFiles=dataFiles.filter(function(f){return f.file;});
    if (actualFiles.length===0) { setStatus2(processStatus2,'请至少上传 1 个数据文件（缓存文件不能计算）','error'); return; }

    processing=true; processed=false; outputWorkbook=null; updateStats2(); setProgress2(5,'读取文件中...');
    logBox2.textContent='正在读取并解析表格...';

    try {
      var tmplWb=await readWorkbook2(templateFile.file);
      var dataEntries=[];
      for(var i=0;i<actualFiles.length;i++){ dataEntries.push({name:actualFiles[i].name,file:actualFiles[i].file,wb:await readWorkbook2(actualFiles[i].file)}); }

      var scoreSheets=function(wb,names){ if(!wb)return 0; var sheetNames=wb.SheetNames||[]; var trimmed=sheetNames.map(function(n){return String(n).trim();}); var score=0; names.forEach(function(n){var key=String(n).trim(); if(trimmed.includes(key)||sheetNames.some(function(s){return String(s).includes(key);}))score+=1;}); return score; };
      var pickBySheets=function(entries,names,fallbackKeyword){ var best=null,bestScore=0; entries.forEach(function(e){var s=scoreSheets(e.wb,names); if(s>bestScore){bestScore=s;best=e;}}); if(bestScore>0)return best; if(fallbackKeyword)return entries.find(function(e){return e.name.includes(fallbackKeyword);})||null; return null; };

      var billEntry=pickBySheets(dataEntries,['出库单','出库尾程','退货费用'],'账单');
      var orderEntry=pickBySheets(dataEntries,['订单','订单信息','订单明细','订单表'],'订单');
      var expenseEntry=pickBySheets(dataEntries,['头程费用','入库费用','仓储费用','汇率','产品成本'],'支出');
      var incomeEntry=pickBySheets(dataEntries,['报价','C端收入','第一档收入'],'收入');

      var billWb2=billEntry?billEntry.wb:null, orderWb2=orderEntry?orderEntry.wb:null;
      var expenseWb2=expenseEntry?expenseEntry.wb:null, incomeWb2=incomeEntry?incomeEntry.wb:null;
      var missingFiles=[]; if(!billEntry)missingFiles.push('账单信息'); if(!orderEntry)missingFiles.push('订单信息'); if(!expenseEntry)missingFiles.push('支出情况'); if(!incomeEntry)missingFiles.push('收入情况');

      setProgress2(20,'解析工作表...');
      var tmplSheetName=tmplWb.SheetNames[0]; var tmplWs=tmplWb.Sheets[tmplSheetName];
      var tmplAoA=XLSX.utils.sheet_to_json(tmplWs,{header:1,defval:null});
      var headerRowIndex=findHeaderRowIndex2(tmplAoA);
      var headers=(tmplAoA[headerRowIndex]||tmplAoA[0]||[]).map(function(h){return typeof h==='string'?h.trim():h;});

      var templateRows=[];
      for(var ti=headerRowIndex+1;ti<tmplAoA.length;ti++){
        var rowArr=tmplAoA[ti]||[]; var row={};
        headers.forEach(function(h,idx3){ if(h==null||h==='')return; row[h]=rowArr[idx3]!=null?rowArr[idx3]:null; });
        if(!isBlank2(row['订单编号'])||!isBlank2(row['国家'])||!isBlank2(row['期间']))templateRows.push(row);
      }

      var billOut2=billWb2?normalizeRows2(sheetToRows2(getSheetOrThrow2(billWb2,'出库单','账单信息'))):[];
      var billTail2=billWb2?normalizeRows2(sheetToRows2(getSheetOrThrow2(billWb2,'出库尾程','账单信息'))):[];
      var billRet2=billWb2?normalizeRows2(sheetToRows2(getSheetOrThrow2(billWb2,'退货费用','账单信息'))):[];
      var orders2=orderWb2?normalizeRows2(sheetToRows2(getSheetOrThrow2(orderWb2,['订单','订单信息','订单明细','订单表'],'订单信息'))):[];
      var headFee2=expenseWb2?normalizeRows2(sheetToRows2(getSheetOrThrow2(expenseWb2,'头程费用','支出情况'))):[];
      var entryFee2=expenseWb2?normalizeRows2(sheetToRows2(getSheetOrThrow2(expenseWb2,'入库费用','支出情况'))):[];
      var warehouseFee2=expenseWb2?normalizeRows2(sheetToRows2(getSheetOrThrow2(expenseWb2,'仓储费用','支出情况'))):[];
      var fxRows2=expenseWb2?normalizeRows2(sheetToRows2(getSheetOrThrow2(expenseWb2,'汇率','支出情况'))):[];
      var productCost2=expenseWb2?normalizeRows2(sheetToRows2(getSheetOrThrow2(expenseWb2,'产品成本','支出情况'))):[];
      var quoteRows2=incomeWb2?normalizeRows2(sheetToRows2(getSheetOrThrow2(incomeWb2,'报价','收入情况'))):[];
      var cincomeRows2=incomeWb2?normalizeRows2(sheetToRows2(getSheetOrThrow2(incomeWb2,'C端收入','收入情况'))):[];
      var firstIncomeRows2=incomeWb2?normalizeRows2(sheetToRows2(getSheetOrThrow2(incomeWb2,'第一档收入','收入情况'))):[];

      setProgress2(40,'构建映射关系...');
      var missing2={sku:new Map(),quote:[],rate:new Map(),fee:new Map()};
      var missingOrders2=new Set();
      var recordMissingSku2=function(oid,sku){if(!oid||!sku)return; if(!missing2.sku.has(oid))missing2.sku.set(oid,new Set()); missing2.sku.get(oid).add(sku);};
      var recordMissingQuote2=function(row){if(!row)return; missing2.quote.push(row);};
      var recordMissingRate2=function(currency,period,scene){if(!currency||!period)return; var key=currency+'||'+period+'||'+scene; if(!missing2.rate.has(key))missing2.rate.set(key,{currency:currency,period:period,scene:scene});};
      var recordMissingFee2=function(scene,country,period){if(!scene||!country)return; var key=scene+'||'+country+'||'+(period||''); if(!missing2.fee.has(key))missing2.fee.set(key,{scene:scene,country:country,period:period});};

      var fxMap2={};
      fxRows2.forEach(function(r){var cur=normalizeCurrency2(r['原币']); var period=r['期间']; var rate=toNumber2(r['直接汇率']); if(cur&&period!=null&&rate!=null)fxMap2[cur+'||'+period]=rate;});

      var getRate2=function(currency,period,scene){ var cur=normalizeCurrency2(currency); if(!cur||cur==='人民币')return 1; var rate=fxMap2[cur+'||'+period]; if(rate==null){recordMissingRate2(cur,period,scene||'未知');return null;} return rate; };

      var tailGroup2={};
      billTail2.forEach(function(r){ var id=normalizeKey2(r['订单编号']); var sys=normalizeKey2(r['系统单号']); [id,sys].filter(Boolean).forEach(function(key){if(!tailGroup2[key])tailGroup2[key]={出库费用:0,尾程费用:0,币种:null}; tailGroup2[key].出库费用+=toNumber2(r['出库费用'])||0; tailGroup2[key].尾程费用+=toNumber2(r['尾程费用'])||0; if(!tailGroup2[key].币种&&r['币种'])tailGroup2[key].币种=r['币种'];}); });
      var retGroup2={};
      billRet2.forEach(function(r){ var id=normalizeKey2(r['退货单号']); var sys=normalizeKey2(r['系统单号']); [id,sys].filter(Boolean).forEach(function(key){if(!retGroup2[key])retGroup2[key]={退货金额:0,币种:null}; retGroup2[key].退货金额+=toNumber2(r['退货金额'])||0; if(!retGroup2[key].币种&&r['币种'])retGroup2[key].币种=r['币种'];}); });
      var billOutMap2={}, billOutSysMap2={};
      billOut2.forEach(function(r){ var id=normalizeKey2(r['订单编号']); if(id&&!billOutMap2[id])billOutMap2[id]=r; var sys=normalizeKey2(r['系统单号']); if(sys&&!billOutSysMap2[sys])billOutSysMap2[sys]=r; });

      var orderInfo2={}, orderAlias2={}, orderSkuQty2={}, orderQtySum2={};
      var addAlias2=function(key,id){ var nk=normalizeKey2(key); if(!nk||!id)return; if(!orderAlias2[nk])orderAlias2[nk]=id; };
      orders2.forEach(function(r){
        var id=normalizeKey2(r['订单编号']); if(!id)return;
        if(!orderInfo2[id])orderInfo2[id]={'订单状态':null,'客户经理':null,'店铺':null,'卖家订单号':null,'收件方式':null,'订单实付':null};
        var info2=orderInfo2[id];
        if(!info2['订单状态']&&r['订单状态'])info2['订单状态']=r['订单状态'];
        if(!info2['客户经理']&&r['客户经理'])info2['客户经理']=r['客户经理'];
        if(!info2['店铺']&&r['店铺'])info2['店铺']=r['店铺'];
        if(!info2['卖家订单号']&&r['卖家订单号'])info2['卖家订单号']=r['卖家订单号'];
        if(!info2['收件方式']&&r['收件方式'])info2['收件方式']=r['收件方式'];
        if(!info2['订单实付']&&r['订单实付'])info2['订单实付']=r['订单实付'];
        var sku=String(r['产品SKU']||'').trim(); var qty2=toNumber2(r['购买数量'])||0;
        if(sku){ if(!orderSkuQty2[id])orderSkuQty2[id]={}; orderSkuQty2[id][sku]=(orderSkuQty2[id][sku]||0)+qty2; if(!orderQtySum2[id])orderQtySum2[id]=0; orderQtySum2[id]+=qty2; }
        addAlias2(r['卖家订单号'],id); addAlias2(r['平台交易号'],id); addAlias2(r['快递单号'],id); addAlias2(r['第三方物流单号'],id); addAlias2(r['系统单号'],id);
      });

      var productCostMap2={};
      productCost2.forEach(function(r){ var sku=String(r['聚水潭旧编码']||'').trim(); if(!sku)return; var cost=toNumber2(r['成本价']); if(cost!=null)productCostMap2[sku]=cost; });
      var firstIncomeMap2={};
      firstIncomeRows2.forEach(function(r){ var sku=String(r['聚水潭旧编码']||'').trim(); if(!sku)return; var val=toNumber2(r['第一档价格']); firstIncomeMap2[sku]=val!=null?val:0; });

      var costByOrder2={}, firstIncomeBaseByOrder2={};
      Object.keys(orderSkuQty2).forEach(function(id){
        var totalCost=0,totalFirst=0; var skuMap=orderSkuQty2[id];
        Object.keys(skuMap).forEach(function(sku){ var qty3=skuMap[sku]; var unit=productCostMap2[sku]; if(unit==null){recordMissingSku2(id,sku);}else{totalCost+=unit*qty3;} var first=firstIncomeMap2[sku]||0; totalFirst+=first*qty3; });
        costByOrder2[id]=totalCost; firstIncomeBaseByOrder2[id]=totalFirst;
      });

      var cincomeMap2={};
      cincomeRows2.forEach(function(r){ var key=String(r['卖家订单号']||'').trim(); if(!key)return; var amt=toNumber2(r['人民币']); if(!cincomeMap2[key])cincomeMap2[key]=0; cincomeMap2[key]+=amt||0; });

      var headFeeMap2={}, headCurMap2={};
      headFee2.forEach(function(r){ var key=String(r['国家']||'').trim(); if(!key)return; headFeeMap2[key]=toNumber2(r['头程费用']); headCurMap2[key]=r['币种']; });
      var entryFeeMap2={}, entryCurMap2={};
      entryFee2.forEach(function(r){ var key=String(r['国家']||'').trim(); if(!key)return; entryFeeMap2[key]=toNumber2(r['入库费用']); entryCurMap2[key]=r['币种']; });
      var warehouseFeeMap2={}, warehouseCurMap2={};
      warehouseFee2.forEach(function(r){ var key=r['期间']+'||'+r['国家']; warehouseFeeMap2[key]=toNumber2(r['仓储费用']); warehouseCurMap2[key]=r['币种']; });
      var quoteMap2={};
      quoteRows2.forEach(function(r){ var country=String(r['国家']||'').trim(); var qtyRaw=String(r['件数（n）/订单']||'').replace(/\s/g,''); if(!country||!qtyRaw)return; quoteMap2[country+'||'+qtyRaw]=r; });

      var computeFirstIncomeTotal2=function(oid,add){ if(!Object.prototype.hasOwnProperty.call(firstIncomeBaseByOrder2,oid))return null; var base=firstIncomeBaseByOrder2[oid]||0; var qty4=orderQtySum2[oid]||0; return base+add*qty4; };
      var extractAdd2=function(text){ var m=String(text).match(/\+\s*([0-9]*\.?[0-9]+)/); return m?parseFloat(m[1]):0; };
      var resolveOrderKey2=function(){
        for(var i2=0;i2<arguments.length;i2++){
          var raw=arguments[i2]==null?'':String(arguments[i2]).trim();
          var variants=[raw];
          var stripped=stripOrderSuffix2(raw);
          if(stripped&&stripped!==raw)variants.push(stripped);
          for(var j2=0;j2<variants.length;j2++){
            var nk=normalizeKey2(variants[j2]);
            if(!nk)continue;
            if(orderInfo2[nk])return nk;
            if(orderAlias2[nk])return orderAlias2[nk];
          }
        }
        return normalizeKey2(stripOrderSuffix2(arguments[0])||arguments[0])||'';
      };
      var evalSpecTotal2=function(oid,spec,nonShippingSpec,qty5){
        if(spec===null||spec===undefined||spec==='')return null;
        if(typeof spec==='number')return spec*qty5;
        var s=String(spec).trim();
        if(s.includes('不包邮基础上')){ var add=extractAdd2(s); var baseTotal=evalSpecTotal2(oid,nonShippingSpec,null,qty5); if(baseTotal==null)return null; return baseTotal+add*qty5; }
        if(s.includes('第一档')){ var add2=extractAdd2(s); return computeFirstIncomeTotal2(oid,add2); }
        var n=toNumber2(s); if(n!=null)return n*qty5; return null;
      };

      setProgress2(60,'计算中...');
      var baseRows2=templateRows;
      if(baseRows2.length===0&&billOut2.length>0){
        baseRows2=billOut2.map(function(r){return{'国家':r['国家'],'期间':r['期间'],'系统单号':r['系统单号'],'订单编号':r['订单编号'],'数量':r['数量'],'账单币种':r['账单币种'],'出库费用':r['出库费用'],'尾程费用':r['尾程费用'],'退件费用':r['退件费用']};});
      }
      if(baseRows2.length===0)throw new Error('模板无有效数据且未上传账单信息，无法计算');

      var qtySumMap2={};
      var qtySource2=billOut2.length>0?billOut2:baseRows2;
      qtySource2.forEach(function(r){ var key=r['期间']+'||'+r['国家']; if(!r['期间']||!r['国家'])return; var qv=toNumber2(r['数量']); if(!qtySumMap2[key])qtySumMap2[key]=0; qtySumMap2[key]+=(qv!=null?qv:1); });

      var resultRows2=[];
      baseRows2.forEach(function(base){
        var row=Object.assign({},base);
        var orderIdRaw=pickValue(row['订单编号'],row['系统单号'],row['卖家订单号']);
        var orderId2=normalizeKey2(orderIdRaw);
        var orderKey2=resolveOrderKey2(orderIdRaw,row['卖家订单号'],row['系统单号']);
        var billRow2=orderId2?billOutMap2[orderId2]:(orderKey2?billOutMap2[orderKey2]:null);
        var sysId2=normalizeKey2(pickValue(row['系统单号'],billRow2?billRow2['系统单号']:null));
        if(!billRow2&&sysId2)billRow2=billOutSysMap2[sysId2]||null;
        var orderIdDisp=orderIdRaw==null?'':String(orderIdRaw).trim();
        if(orderIdDisp&&(!orderKey2||!orderInfo2[orderKey2]))missingOrders2.add(orderIdDisp);

        var country2=pickValue(row['国家'],billRow2?billRow2['国家']:null);
        var period2=pickValue(row['期间'],billRow2?billRow2['期间']:null);
        var qtyRaw2=pickValue(row['数量'],billRow2?billRow2['数量']:null);
        var qtyNum2=toNumber2(qtyRaw2);

        setIfBlank2(row,'订单编号',orderIdRaw||null); setIfBlank2(row,'国家',country2); setIfBlank2(row,'期间',period2);
        if(qtyNum2!=null)setIfBlank2(row,'数量',qtyNum2);
        setIfBlank2(row,'系统单号',billRow2?billRow2['系统单号']:null);

        var info3=orderKey2?(orderInfo2[orderKey2]||{}):{};
        var manager2=pickValue(row['客户经理'],info3['客户经理']);
        var shipMethod2=pickValue(row['收件方式'],info3['收件方式']);
        var shipMethodText2=shipMethod2==null?'':String(shipMethod2);

        setIfBlank2(row,'订单状态',pickValue(row['订单状态'],info3['订单状态']));
        setIfBlank2(row,'客户经理',manager2); setIfBlank2(row,'收件方式',shipMethodText2);
        if(hasKeyword2(manager2)){ setIfBlank2(row,'店铺名称',pickValue(row['店铺名称'],info3['店铺'])); setIfBlank2(row,'卖家订单号',pickValue(row['卖家订单号'],info3['卖家订单号'])); }

        var tailInfo2=orderId2?(tailGroup2[orderId2]||{}):{}; if((!tailInfo2||!Object.keys(tailInfo2).length)&&sysId2)tailInfo2=tailGroup2[sysId2]||{};
        var retInfo2=orderId2?(retGroup2[orderId2]||{}):{}; if((!retInfo2||!Object.keys(retInfo2).length)&&sysId2)retInfo2=retGroup2[sysId2]||{};

        var baseOut2=toNumber2(row['出库费用'])!=null?toNumber2(row['出库费用']):(billRow2?toNumber2(billRow2['出库费用']):null);
        var baseTail2=toNumber2(row['尾程费用'])!=null?toNumber2(row['尾程费用']):(billRow2?toNumber2(billRow2['尾程费用']):null);
        var baseRet2=toNumber2(row['退件费用'])!=null?toNumber2(row['退件费用']):(billRow2?toNumber2(billRow2['退件费用']):null);

        var outFee2=baseOut2!=null?baseOut2:(tailInfo2['出库费用']!=null?tailInfo2['出库费用']:null);
        var tailFee2=baseTail2!=null?baseTail2:(tailInfo2['尾程费用']!=null?tailInfo2['尾程费用']:null);
        var retFee2=baseRet2!=null?baseRet2:(retInfo2['退货金额']!=null?retInfo2['退货金额']:null);

        setIfBlank2(row,'出库费用',outFee2!=null?round2fn(outFee2):null);
        setIfBlank2(row,'尾程费用',tailFee2!=null?round2fn(tailFee2):null);
        setIfBlank2(row,'退件费用',retFee2!=null?round2fn(retFee2):null);

        var billCur2=pickValue(row['账单币种'],billRow2?billRow2['账单币种']:null,retInfo2['币种'],tailInfo2['币种']);
        setIfBlank2(row,'账单币种',billCur2);

        var outCur2=baseOut2!=null?billCur2:(tailInfo2['币种']||billCur2);
        var tailCur2=baseTail2!=null?billCur2:(tailInfo2['币种']||billCur2);
        var retCur2=baseRet2!=null?billCur2:(retInfo2['币种']||billCur2);

        var headUnit2=country2?(headFeeMap2[country2]!=null?headFeeMap2[country2]:null):null;
        if(country2&&headUnit2==null)recordMissingFee2('头程费用',country2,period2);
        if(headUnit2!=null&&qtyNum2!=null){ var headRate2=getRate2(headCurMap2[country2],period2,'头程费用'); if(headRate2!=null)setIfBlank2(row,'头程费用\n（人民币）',round2fn(headUnit2*qtyNum2*headRate2)); }
        var entryUnit2=country2?(entryFeeMap2[country2]!=null?entryFeeMap2[country2]:null):null;
        if(country2&&entryUnit2==null)recordMissingFee2('入库费用',country2,period2);
        if(entryUnit2!=null&&qtyNum2!=null){ var entryRate2=getRate2(entryCurMap2[country2],period2,'入库费用'); if(entryRate2!=null)setIfBlank2(row,'入库费用\n（人民币）',round2fn(entryUnit2*qtyNum2*entryRate2)); }

        if(outFee2!=null){ var outRate2=getRate2(outCur2,period2,'出库费用'); if(outRate2!=null)setIfBlank2(row,'出库费用\n（人民币）',round2fn(outFee2*outRate2)); }
        if(tailFee2!=null){ var tailRate2=getRate2(tailCur2,period2,'尾程费用'); if(tailRate2!=null)setIfBlank2(row,'尾程费用\n（人民币）',round2fn(tailFee2*tailRate2)); }
        if(retFee2!=null){ var retRate2=getRate2(retCur2,period2,'退件费用'); if(retRate2!=null)setIfBlank2(row,'退件费用（人民币）',round2fn(retFee2*retRate2)); }

        if(period2!=null&&country2!=null){
          var whKey2=period2+'||'+country2; var whTotal2=toNumber2(warehouseFeeMap2[whKey2]); if(whTotal2==null)recordMissingFee2('仓储费用',country2,period2); var totalQty2=qtySumMap2[whKey2]||0;
          var orderQty2=qtyNum2!=null?qtyNum2:1; var whAlloc2=(whTotal2!=null&&totalQty2>0)?(whTotal2/totalQty2*orderQty2):null;
          if(whAlloc2!=null){ var whCur2=warehouseCurMap2[whKey2]||'人民币'; if(normalizeCurrency2(whCur2)==='人民币'){setIfBlank2(row,'仓储费\n（人民币）',whAlloc2);}else{var wr=getRate2(whCur2,period2,'仓储费用'); if(wr!=null)setIfBlank2(row,'仓储费\n（人民币）',whAlloc2*wr);} }
        }

        var cost2=orderKey2?(costByOrder2[orderKey2]!=null?costByOrder2[orderKey2]:null):null;
        if(cost2!=null)setIfBlank2(row,'产品成本\n（人民币）',round2fn(cost2));

        var orderIncome2=null;
        if(orderKey2){
          if(hasKeyword2(manager2)){
            var sellerId2=pickValue(row['卖家订单号'],info3['卖家订单号']);
            if(sellerId2&&cincomeMap2[String(sellerId2).trim()]!=null)orderIncome2=cincomeMap2[String(sellerId2).trim()];
          } else {
            var pay2=pickValue(row['订单实付'],info3['订单实付']); var parsed2=parseAmount2(pay2);
            if(parsed2.amount!=null){ var rate3=getRate2(parsed2.currency,period2,'订单收入'); if(rate3!=null)orderIncome2=parsed2.amount*rate3; }
          }
        }
        if(orderIncome2!=null)setIfBlank2(row,'订单收入',round2fn(orderIncome2));

        if(country2&&qtyNum2!=null&&quoteRows2.length>0){
          var qLabel2=qtyLabel2(qtyNum2);
          if(qLabel2){
            var qrow2=quoteMap2[String(country2).trim()+'||'+qLabel2.replace(/\s/g,'')];
            if(qrow2){
              var useCol2=shipMethodText2.includes('自提')?'不包邮':(shipMethodText2.includes('快递')?'包邮':'包邮');
              var spec2=qrow2[useCol2]; var nonShip2=qrow2['不包邮'];
              var total2=evalSpecTotal2(orderId2,spec2,nonShip2,qtyNum2);
              var rate4=getRate2(qrow2['币种'],period2,'报价收入');
              if(total2!=null&&rate4!=null)setIfBlank2(row,'报价收入',round2fn(total2*rate4));
              if(total2==null)recordMissingQuote2({orderId:orderId2,country:country2,qty:qtyNum2,method:shipMethodText2,band:qLabel2,reason:'缺少订单信息或第一档收入'});
            } else {
              recordMissingQuote2({orderId:orderId2,country:country2,qty:qtyNum2,method:shipMethodText2,band:qLabel2,reason:'报价表未匹配'});
            }
          }
        }

        var expenseFields2=['头程费用\n（人民币）','产品成本\n（人民币）','入库费用\n（人民币）','出库费用\n（人民币）','尾程费用\n（人民币）','退件费用（人民币）','仓储费\n（人民币）'];
        var expenseVals2=expenseFields2.map(function(f){return toNumber2(row[f]);});
        var hasExpense2=expenseVals2.some(function(v){return v!=null;});
        var orderExpenseCalc2=expenseVals2.reduce(function(s,v){return s+(v!=null&&!isNaN(v)?v:0);},0);
        if(hasExpense2)setIfBlank2(row,'订单支出',round2fn(orderExpenseCalc2));

        var incomeForGross2=toNumber2(pickValue(row['订单收入'],orderIncome2))||0;
        var expenseForGross2=toNumber2(pickValue(row['订单支出'],hasExpense2?orderExpenseCalc2:null))||0;
        var hasGrossInput2=!isBlank2(row['订单收入'])||orderIncome2!=null||!isBlank2(row['订单支出'])||hasExpense2;
        if(hasGrossInput2){ var gross2=incomeForGross2-expenseForGross2; setIfBlank2(row,'毛利',round2fn(gross2)); setIfBlank2(row,'盈亏情况',gross2>0?'盈':'亏'); }

        resultRows2.push(row);
      });

      setProgress2(85,'生成输出表...');
      outputWorkbook=buildOutputWorkbook2(tmplWb,resultRows2,headerRowIndex);
      var first2=resultRows2.find(function(r){return !isBlank2(r['国家'])&&!isBlank2(r['期间']);});
      if(first2&&first2['国家']&&first2['期间'])outputFilename=first2['国家']+first2['期间']+'.xlsx';

      renderPreview2(resultRows2);
      renderMissingLists2(missing2);
      processing=false; processed=true; updateStats2(resultRows2.length); setProgress2(100,'数据处理完成！点击下载表格');

      var missingNote2=missingFiles.length?'提示：缺少'+missingFiles.join('、')+'，已尽力填充空白字段。':'';
      var summary2=['处理完成：共 '+resultRows2.length+' 条订单。'];
      var recognized2=[]; if(billEntry)recognized2.push('账单信息：'+billEntry.name); if(orderEntry)recognized2.push('订单信息：'+orderEntry.name); if(expenseEntry)recognized2.push('支出情况：'+expenseEntry.name); if(incomeEntry)recognized2.push('收入情况：'+incomeEntry.name);
      if(recognized2.length)summary2.unshift('识别文件：'+recognized2.join('；')); if(missingNote2)summary2.push(missingNote2);
      if(missingOrders2.size){ var sample2=Array.from(missingOrders2).slice(0,20); summary2.push('订单信息缺失：共 '+missingOrders2.size+' 条，示例：'+sample2.join('、')+(missingOrders2.size>20?'…':'')); }
      if(missing2.fee.size){ var sampleFee2=Array.from(missing2.fee.values()).slice(0,10).map(function(row){ return row.scene+':'+row.country+(row.period?('('+row.period+')'):''); }); summary2.push('费用配置缺失：共 '+missing2.fee.size+' 项，示例：'+sampleFee2.join('、')+(missing2.fee.size>10?'…':'')); }
      if(missing2.rate.size){ var sampleRate2=Array.from(missing2.rate.values()).slice(0,10).map(function(row){ return row.currency+' '+row.period+' ('+row.scene+')'; }); summary2.push('汇率缺失：共 '+missing2.rate.size+' 项，示例：'+sampleRate2.join('、')+(missing2.rate.size>10?'…':'')); }
      logBox2.textContent=summary2.join('\n');
      downloadBtn.disabled=false;

    } catch(err) {
      processing=false; processed=false;
      setStatus2(processStatus2,'处理失败：'+err.message,'error');
      logBox2.textContent='处理失败，请检查文件是否齐全且表头一致。';
    }
  });

  downloadBtn.addEventListener('click', function() {
    if (!outputWorkbook) { setStatus2(processStatus2,'请先完成计算再下载','error'); return; }
    downloadWorkbook2(outputWorkbook, outputFilename||'输出结果.xlsx');
  });

  KEYWORDS = loadKeywords(); renderKeywords(); restoreCache(); updateStats2();
})();
