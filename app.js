// app.js - 在职人员信息分析平台主逻辑
// 全局状态
let allData = [];
let filteredData = [];
let charts = {};
let currentFileName = '';

// 常量
const EDU_ORDER = ['博士', '硕士', '本科', '大专', '中专', '初中', '/'];
const EDU_LABELS = ['博士', '硕士', '本科', '大专', '中专及以下'];
const COMPANY_COLORS = {
  '武汉宏韧': { bg: 'rgba(79,70,229,0.8)', border: '#4F46E5' },
  '武汉弘质': { bg: 'rgba(220,38,38,0.8)', border: '#DC2626' },
  '上海宏韧': { bg: 'rgba(5,150,105,0.8)', border: '#059669' },
  '广州宏韧': { bg: 'rgba(217,119,6,0.8)', border: '#D97706' }
};
const COMPANY_ORDER = ['武汉宏韧', '武汉弘质', '上海宏韧', '广州宏韧'];
const PALETTE = ['#4F46E5','#DC2626','#059669','#D97706','#7C3AED','#0891B2','#DB2777','#65A30D','#EA580C','#4338CA','#0D9488','#CA8A04','#9333EA','#0284C7','#C026D3'];

// ===== Excel 日期转换 =====
function excelDate(serial) {
  if (!serial || typeof serial !== 'number') return serial;
  var utcDays = Math.floor(serial - 25569);
  var date = new Date(utcDays * 86400 * 1000);
  return date.toISOString().split('T')[0];
}

// ===== 公司归属逻辑 =====
function classifyCompany(dept) {
  if (!dept) return '未知';
  if (dept.indexOf('上海') !== -1) return '上海宏韧';
  if (dept.indexOf('广州') !== -1) return '广州宏韧';
  if (dept.indexOf('数据管理与统计') !== -1) return '武汉弘质';
  if (dept === '/') return '未知';
  return '武汉宏韧';
}

// ===== 部门归一化逻辑 =====
// 将带后缀的部门名归并到父部门，如"化药生物分析部-SA"→"化药生物分析部"
function normalizeDept(dept) {
  if (!dept) return dept;
  var idx = dept.indexOf('-');
  if (idx !== -1) return dept.substring(0, idx);
  return dept;
}

// ===== 岗位归一化逻辑 =====
// 将带前缀/后缀的岗位归并到核心岗位，如"助理SD"→"SD"、"资深SD"→"SD"、"SD经理"→"SD"
// 按优先级从长到短匹配，避免短词误匹配
var POSITION_KEYWORDS = [
  'SD经理','SA经理','MD经理','QC经理','PM经理',
  'SD','SA','MD','QC','PM'
];
function normalizePosition(pos) {
  if (!pos) return pos;
  for (var i = 0; i < POSITION_KEYWORDS.length; i++) {
    if (pos.indexOf(POSITION_KEYWORDS[i]) !== -1) {
      return POSITION_KEYWORDS[i];
    }
  }
  return pos;
}

// ===== 识别最新日期的 Sheet =====
// Sheet 名格式通常为 "2026.6.10"、"2025.5.29" 形式
// 从所有 Sheet 中找到日期最新的那个返回其名称
function findLatestDateSheet(sheetNames) {
  var datePattern = /^(\d{4})\.(\d{1,2})\.(\d{1,2})$/;
  var best = null;
  var bestTime = -1;
  for (var i = 0; i < sheetNames.length; i++) {
    var m = sheetNames[i].match(datePattern);
    if (m) {
      var year = parseInt(m[1], 10);
      var month = parseInt(m[2], 10);
      var day = parseInt(m[3], 10);
      var t = year * 10000 + month * 100 + day;
      if (t > bestTime) {
        bestTime = t;
        best = sheetNames[i];
      }
    }
  }
  return best;
}

// ===== 处理 Excel 原始数据 =====
function processExcelData(rawRows) {
  return rawRows.map(function(r) {
    var dept = normalizeDept(r['部门'] || '');
    var result = {
      编号: r['1'] || '',
      组别: r['组别'] || '',
      分组: r['分组'] || '',
      姓名: r['姓名'] || '',
      年龄: typeof r['年龄'] === 'number' ? Math.floor(r['年龄']) : '',
      部门: dept,
      公司: classifyCompany(dept),
      职级: r['职级'] || '',
      职等序列: r['职等序列'] || '',
      岗位: normalizePosition(r['岗位'] || ''),
      入职时间: excelDate(r['入职时间']),
      年限: typeof r['年限'] === 'number' ? Math.round(r['年限'] * 10) / 10 : '',
      学历: r['学历'] || '',
      毕业学校: r['毕业学校'] || '',
      专业: r['专业'] || '',
      毕业时间: excelDate(r['毕业时间']),
      性别: r['性别'] || ''
    };
    return result;
  }).filter(function(r) {
    return r['姓名'] && r['部门'] && r['部门'] !== '/' && r['姓名'] !== '徐会-吴天悦';
  });
}

// ===== 文件导入 =====
function handleFileImport(input) {
  var file = input.files ? input.files[0] : null;
  if (!file) return;
  currentFileName = file.name;
  var reader = new FileReader();
  reader.onload = function(e) {
    try {
      var wb = XLSX.read(e.target.result, { type: 'array' });
      var targetSheet = findLatestDateSheet(wb.SheetNames) || wb.SheetNames[0];
      var ws = wb.Sheets[targetSheet];
      var rawData = XLSX.utils.sheet_to_json(ws);
      allData = processExcelData(rawData);
      showImportSuccess();
      initDataView();
    } catch(err) {
      alert('文件解析失败：' + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
  input.value = '';
}

function showImportSuccess() {
  var panel = document.getElementById('importPanel');
  var companyCounts = {};
  allData.forEach(function(d) {
    var c = d['公司'];
    companyCounts[c] = (companyCounts[c] || 0) + 1;
  });
  var parts = [];
  Object.keys(companyCounts).forEach(function(k) {
    parts.push(k + ' ' + companyCounts[k] + '人');
  });
  var companyStr = parts.join('，');

  panel.classList.add('has-data');
  panel.innerHTML =
    '<div class="import-success">' +
      '<div class="check">✓</div>' +
      '<div class="info">' +
        '<h3>已导入：' + currentFileName + '</h3>' +
        '<p>共 ' + allData.length + ' 条记录 — ' + companyStr + '</p>' +
      '</div>' +
      '<button class="btn-reimport" onclick="reimport()">重新导入</button>' +
    '</div>' +
    '<div class="mapping-hint">' +
      '<b>📋 公司归属规则：</b>部门含"上海"→上海宏韧，含"广州"→广州宏韧，含"数据管理与统计"→武汉弘质，其余→武汉宏韧' +
    '</div>';
}

function reimport() {
  var panel = document.getElementById('importPanel');
  panel.classList.remove('has-data');
  panel.innerHTML =
    '<div class="import-icon">📁</div>' +
    '<h3>导入在职人员信息 Excel</h3>' +
    '<p>将 .xlsx 文件拖拽到此处，或点击下方按钮选择文件</p>' +
    '<button class="btn-choose" onclick="document.getElementById(\'fileInput\').click()">选择 Excel 文件</button>' +
    '<input type="file" id="fileInput" accept=".xlsx,.xls" onchange="handleFileImport(this)">' +
    '<div class="mapping-hint">' +
      '<b>📋 自动识别规则：</b>根据"部门"字段自动归属公司 — 部门含"上海"→上海宏韧，含"广州"→广州宏韧，含"数据管理与统计"→武汉弘质，其余→武汉宏韧' +
    '</div>';
  setupDragDrop();
  document.getElementById('dataSection').classList.remove('visible');
}

function setupDragDrop() {
  var panel = document.getElementById('importPanel');
  panel.addEventListener('dragover', function(e) { e.preventDefault(); panel.classList.add('drag-over'); });
  panel.addEventListener('dragleave', function(e) { e.preventDefault(); panel.classList.remove('drag-over'); });
  panel.addEventListener('drop', function(e) {
    e.preventDefault();
    panel.classList.remove('drag-over');
    var file = e.dataTransfer.files[0];
    if (file && (file.name.endsWith('.xlsx') || file.name.endsWith('.xls'))) {
      currentFileName = file.name;
      var reader = new FileReader();
      reader.onload = function(ev) {
        try {
          var wb = XLSX.read(ev.target.result, { type: 'array' });
          var targetSheet = findLatestDateSheet(wb.SheetNames) || wb.SheetNames[0];
          var ws = wb.Sheets[targetSheet];
          var rawData = XLSX.utils.sheet_to_json(ws);
          allData = processExcelData(rawData);
          showImportSuccess();
          initDataView();
        } catch(err) { alert('文件解析失败：' + err.message); }
      };
      reader.readAsArrayBuffer(file);
    } else {
      alert('请选择 .xlsx 或 .xls 格式的文件');
    }
  });
}

// ===== 数据展示 =====
function initDataView() {
  document.getElementById('dataSection').classList.add('visible');
  initFilters();
  renderAll();
}

function initFilters() {
  document.getElementById('filterCompany').innerHTML = '<option value="">全部公司</option>';
  document.getElementById('companyTabs').innerHTML = '';

  var companies = [];
  allData.forEach(function(d) {
    if (companies.indexOf(d['公司']) === -1) companies.push(d['公司']);
  });
  companies.sort(function(a, b) {
    var ia = COMPANY_ORDER.indexOf(a);
    var ib = COMPANY_ORDER.indexOf(b);
    if (ia === -1) ia = 999;
    if (ib === -1) ib = 999;
    return ia - ib;
  });

  var selComp = document.getElementById('filterCompany');
  companies.forEach(function(c) {
    var o = document.createElement('option');
    o.value = c;
    o.textContent = c;
    selComp.appendChild(o);
  });

  var tabsDiv = document.getElementById('companyTabs');
  var allTab = document.createElement('span');
  allTab.className = 'company-tab active';
  allTab.dataset.company = '';
  allTab.innerHTML = '全部<span class="count">(' + allData.length + ')</span>';
  allTab.onclick = function() { selectCompanyTab(''); };
  tabsDiv.appendChild(allTab);

  companies.forEach(function(c) {
    var tab = document.createElement('span');
    tab.className = 'company-tab';
    tab.dataset.company = c;
    var count = 0;
    allData.forEach(function(d) { if (d['公司'] === c) count++; });
    tab.innerHTML = c + '<span class="count">(' + count + ')</span>';
    tab.onclick = function() { selectCompanyTab(c); };
    tabsDiv.appendChild(tab);
  });

  updateDependentFilters();
}

function selectCompanyTab(company) {
  document.querySelectorAll('.company-tab').forEach(function(t) { t.classList.remove('active'); });
  var target = document.querySelector('.company-tab[data-company="' + company + '"]');
  if (target) target.classList.add('active');
  document.getElementById('filterCompany').value = company;
  updateDependentFilters();
  renderAll();
}

function updateDependentFilters() {
  var company = document.getElementById('filterCompany').value;
  var subset = company ? allData.filter(function(d) { return d['公司'] === company; }) : allData;

  // 部门
  var depts = [];
  subset.forEach(function(d) { if (d['部门'] && depts.indexOf(d['部门']) === -1) depts.push(d['部门']); });
  depts.sort();
  var selDept = document.getElementById('filterDept');
  var curDept = selDept.value;
  selDept.innerHTML = '<option value="">全部部门</option>';
  depts.forEach(function(d) {
    var o = document.createElement('option');
    o.value = d;
    o.textContent = d;
    selDept.appendChild(o);
  });
  if (depts.indexOf(curDept) !== -1) selDept.value = curDept;

  // 学历
  var edus = [];
  subset.forEach(function(d) { if (d['学历'] && edus.indexOf(d['学历']) === -1) edus.push(d['学历']); });
  edus.sort(function(a, b) { return EDU_ORDER.indexOf(a) - EDU_ORDER.indexOf(b); });
  var selEdu = document.getElementById('filterEdu');
  var curEdu = selEdu.value;
  selEdu.innerHTML = '<option value="">全部学历</option>';
  edus.forEach(function(e) {
    var o = document.createElement('option');
    o.value = e;
    o.textContent = e;
    selEdu.appendChild(o);
  });
  if (edus.indexOf(curEdu) !== -1) selEdu.value = curEdu;

  // 性别
  var selG = document.getElementById('filterGender');
  selG.innerHTML = '<option value="">全部</option><option value="男">男</option><option value="女">女</option>';

  // 岗位
  var positions = [];
  subset.forEach(function(d) { if (d['岗位'] && positions.indexOf(d['岗位']) === -1) positions.push(d['岗位']); });
  positions.sort();
  var selPos = document.getElementById('filterPosition');
  var curPos = selPos.value;
  selPos.innerHTML = '<option value="">全部岗位</option>';
  positions.forEach(function(p) {
    var o = document.createElement('option');
    o.value = p;
    o.textContent = p;
    selPos.appendChild(o);
  });
  if (positions.indexOf(curPos) !== -1) selPos.value = curPos;
}

function applyFilters() {
  var company = document.getElementById('filterCompany').value;
  var dept = document.getElementById('filterDept').value;
  var edu = document.getElementById('filterEdu').value;
  var gender = document.getElementById('filterGender').value;
  var position = document.getElementById('filterPosition').value;
  var ageRange = document.getElementById('filterAge').value;
  var tenureRange = document.getElementById('filterTenure').value;

  filteredData = allData.filter(function(d) {
    if (company && d['公司'] !== company) return false;
    if (dept && d['部门'] !== dept) return false;
    if (edu && d['学历'] !== edu) return false;
    if (gender && d['性别'] !== gender) return false;
    if (position && d['岗位'] !== position) return false;
    if (ageRange) {
      var age = d['年龄'];
      if (!age) return false;
      var lo, hi;
      if (ageRange.indexOf('+') !== -1) { lo = parseInt(ageRange); hi = 999; }
      else { var parts = ageRange.split('-'); lo = Number(parts[0]); hi = Number(parts[1]); }
      if (age < lo || age > hi) return false;
    }
    if (tenureRange) {
      var t = d['年限'];
      if (!t) return false;
      var lo2, hi2;
      if (tenureRange.indexOf('+') !== -1) { lo2 = parseInt(tenureRange); hi2 = 999; }
      else { var parts2 = tenureRange.split('-'); lo2 = Number(parts2[0]); hi2 = Number(parts2[1]); }
      if (t < lo2 || t >= hi2) return false;
    }
    return true;
  });
}

function renderAll() {
  applyFilters();
  renderStats();
  renderCharts();
  renderTable();
}

// ===== 统计卡片 =====
function renderStats() {
  var bar = document.getElementById('statsBar');
  var total = filteredData.length;

  var companyCards = '';
  var seen = {};
  filteredData.forEach(function(d) {
    if (!seen[d['公司']]) {
      seen[d['公司']] = true;
      var cls = d['公司'] === '武汉宏韧' ? 'wh' : d['公司'] === '武汉弘质' ? 'hz' : d['公司'] === '上海宏韧' ? 'sh' : 'gz';
      var cnt = 0;
      filteredData.forEach(function(dd) { if (dd['公司'] === d['公司']) cnt++; });
      companyCards += '<div class="stat-card ' + cls + '"><div class="num">' + cnt + '</div><div class="label">' + d['公司'] + '</div></div>';
    }
  });

  if (Object.keys(seen).length === 0 || document.getElementById('filterCompany').value === '') {
    companyCards = '';
    var allCompanies = [];
    allData.forEach(function(d) {
      if (allCompanies.indexOf(d['公司']) === -1) allCompanies.push(d['公司']);
    });
    allCompanies.sort(function(a, b) {
      var ia = COMPANY_ORDER.indexOf(a);
      var ib = COMPANY_ORDER.indexOf(b);
      if (ia === -1) ia = 999;
      if (ib === -1) ib = 999;
      return ia - ib;
    });
    allCompanies.forEach(function(c) {
      var cls = c === '武汉宏韧' ? 'wh' : c === '武汉弘质' ? 'hz' : c === '上海宏韧' ? 'sh' : 'gz';
      var cnt = 0;
      allData.forEach(function(dd) { if (dd['公司'] === c) cnt++; });
      var displayCnt = total;
      if (document.getElementById('filterCompany').value === '') {
        displayCnt = cnt;
      } else {
        displayCnt = 0;
        filteredData.forEach(function(dd) { if (dd['公司'] === c) displayCnt++; });
      }
      companyCards += '<div class="stat-card ' + cls + '"><div class="num">' + displayCnt + '</div><div class="label">' + c + '</div></div>';
    });
  }

  var maleCount = 0, femaleCount = 0, ageSum = 0, ageCount = 0, masterCount = 0;
  filteredData.forEach(function(d) {
    if (d['性别'] === '男') maleCount++;
    if (d['性别'] === '女') femaleCount++;
    if (d['年龄']) { ageSum += d['年龄']; ageCount++; }
    if (d['学历'] === '博士' || d['学历'] === '硕士') masterCount++;
  });
  var avgAge = ageCount ? (ageSum / ageCount).toFixed(1) : 0;

  bar.innerHTML =
    '<div class="stat-card"><div class="num">' + total + '</div><div class="label">总人数</div></div>' +
    companyCards +
    '<div class="stat-card"><div class="num">' + maleCount + '</div><div class="label">男性</div></div>' +
    '<div class="stat-card"><div class="num">' + femaleCount + '</div><div class="label">女性</div></div>' +
    '<div class="stat-card"><div class="num">' + avgAge + '</div><div class="label">平均年龄</div></div>' +
    '<div class="stat-card"><div class="num">' + masterCount + '</div><div class="label">硕士及以上</div></div>';
}

// ===== 工具函数 =====
function getEduGroup(edu) {
  if (['中专','初中','/'].indexOf(edu) !== -1) return '中专及以下';
  return edu;
}

function countBy(data, key, groupFn) {
  var map = {};
  data.forEach(function(d) {
    var val = groupFn ? groupFn(d[key]) : d[key];
    if (val) map[val] = (map[val] || 0) + 1;
  });
  return map;
}

function getCompanyList() {
  var companies = [];
  allData.forEach(function(d) {
    if (companies.indexOf(d['公司']) === -1) companies.push(d['公司']);
  });
  companies.sort(function(a, b) {
    var ia = COMPANY_ORDER.indexOf(a);
    var ib = COMPANY_ORDER.indexOf(b);
    if (ia === -1) ia = 999;
    if (ib === -1) ib = 999;
    return ia - ib;
  });
  return companies;
}

// ===== 图表渲染 =====
function renderCharts() {
  var compare = document.getElementById('compareMode').checked;
  var company = document.getElementById('filterCompany').value;
  var companies = getCompanyList();
  var dataSets = compare
    ? companies.map(function(c) { return { company: c, data: filteredData.filter(function(d) { return d['公司'] === c; }) }; })
    : [{ company: company || '全部', data: filteredData }];

  renderEduChart(dataSets, compare);
  renderGenderChart(dataSets, compare);
  renderAgeChart(dataSets, compare);
  renderPositionChart();
  renderDeptChart();
  renderTenureChart(dataSets, compare);
}

function renderEduChart(dataSets, compare) {
  var labels = EDU_LABELS;
  if (compare) {
    var datasets = dataSets.map(function(ds, i) {
      var counts = countBy(ds.data, '学历', getEduGroup);
      return {
        label: ds.company,
        data: labels.map(function(l) { return counts[l] || 0; }),
        backgroundColor: (COMPANY_COLORS[ds.company] ? COMPANY_COLORS[ds.company].bg : PALETTE[i]),
        borderColor: (COMPANY_COLORS[ds.company] ? COMPANY_COLORS[ds.company].border : PALETTE[i]),
        borderWidth: 1
      };
    });
    createChart('chartEdu', 'bar', labels, datasets, { indexAxis: 'y', plugins: { legend: { display: true } } });
  } else {
    var counts = countBy(dataSets[0].data, '学历', getEduGroup);
    var data = labels.map(function(l) { return counts[l] || 0; });
    var total = data.reduce(function(a, b) { return a + b; }, 0);
    var colors = ['#7C3AED','#4F46E5','#059669','#D97706','#94A3B8'];
    createChart('chartEdu', 'doughnut', labels, [{ data: data, backgroundColor: colors, borderWidth: 2, borderColor: '#fff' }], {
      plugins: {
        legend: {
          position: 'right',
          labels: {
            font: { size: 12 },
            usePointStyle: true,
            padding: 12,
            generateLabels: function(chart) {
              var ds = chart.data.datasets[0];
              return chart.data.labels.map(function(label, i) {
                var val = ds.data[i];
                var pct = total > 0 ? (val / total * 100).toFixed(1) : '0.0';
                return {
                  text: label + '  ' + val + '人 (' + pct + '%)',
                  fillStyle: ds.backgroundColor[i],
                  strokeStyle: ds.borderColor ? (ds.borderColor[i] || ds.backgroundColor[i]) : ds.backgroundColor[i],
                  lineWidth: 1,
                  index: i,
                  hidden: false
                };
              });
            }
          }
        },
        tooltip: {
          backgroundColor: '#1E293B',
          titleFont: { size: 13 },
          bodyFont: { size: 12 },
          padding: 10,
          cornerRadius: 8,
          callbacks: {
            label: function(ctx) {
              var val = ctx.parsed;
              var pct = total > 0 ? (val / total * 100).toFixed(1) : '0.0';
              return ctx.label + ': ' + val + '人 (' + pct + '%)';
            }
          }
        }
      }
    });
  }
}

function renderGenderChart(dataSets, compare) {
  var labels = ['男', '女'];
  if (compare) {
    var datasets = dataSets.map(function(ds) {
      var counts = countBy(ds.data, '性别');
      return {
        label: ds.company,
        data: labels.map(function(l) { return counts[l] || 0; }),
        backgroundColor: (COMPANY_COLORS[ds.company] ? COMPANY_COLORS[ds.company].bg : '#94A3B8'),
        borderColor: (COMPANY_COLORS[ds.company] ? COMPANY_COLORS[ds.company].border : '#64748B'),
        borderWidth: 1
      };
    });
    createChart('chartGender', 'bar', labels, datasets, { plugins: { legend: { display: true } } });
  } else {
    var counts = countBy(dataSets[0].data, '性别');
    var data = labels.map(function(l) { return counts[l] || 0; });
    var total = data.reduce(function(a, b) { return a + b; }, 0);
    createChart('chartGender', 'doughnut', labels, [{ data: data, backgroundColor: ['rgba(59,130,246,0.8)','rgba(236,72,153,0.8)'], borderWidth: 2, borderColor: '#fff' }], {
      plugins: {
        legend: {
          position: 'right',
          labels: {
            font: { size: 12 },
            usePointStyle: true,
            padding: 12,
            generateLabels: function(chart) {
              var ds = chart.data.datasets[0];
              return chart.data.labels.map(function(label, i) {
                var val = ds.data[i];
                var pct = total > 0 ? (val / total * 100).toFixed(1) : '0.0';
                return {
                  text: label + '  ' + val + '人 (' + pct + '%)',
                  fillStyle: ds.backgroundColor[i],
                  strokeStyle: ds.borderColor ? (ds.borderColor[i] || ds.backgroundColor[i]) : ds.backgroundColor[i],
                  lineWidth: 1,
                  index: i,
                  hidden: false
                };
              });
            }
          }
        },
        tooltip: {
          backgroundColor: '#1E293B',
          titleFont: { size: 13 },
          bodyFont: { size: 12 },
          padding: 10,
          cornerRadius: 8,
          callbacks: {
            label: function(ctx) {
              var val = ctx.parsed;
              var pct = total > 0 ? (val / total * 100).toFixed(1) : '0.0';
              return ctx.label + ': ' + val + '人 (' + pct + '%)';
            }
          }
        }
      }
    });
  }
}

function renderAgeChart(dataSets, compare) {
  var ageRanges = ['20-25','26-30','31-35','36-40','41-45','46-50','50+'];
  function classifyAge(age) {
    if (age <= 25) return '20-25';
    if (age <= 30) return '26-30';
    if (age <= 35) return '31-35';
    if (age <= 40) return '36-40';
    if (age <= 45) return '41-45';
    if (age <= 50) return '46-50';
    return '50+';
  }
  if (compare) {
    var datasets = dataSets.map(function(ds) {
      var counts = {};
      ds.data.forEach(function(d) { if (d['年龄']) { var g = classifyAge(d['年龄']); counts[g] = (counts[g] || 0) + 1; } });
      return {
        label: ds.company,
        data: ageRanges.map(function(r) { return counts[r] || 0; }),
        backgroundColor: (COMPANY_COLORS[ds.company] ? COMPANY_COLORS[ds.company].bg : '#94A3B8'),
        borderColor: (COMPANY_COLORS[ds.company] ? COMPANY_COLORS[ds.company].border : '#64748B'),
        borderWidth: 1
      };
    });
    createChart('chartAge', 'bar', ageRanges.map(function(r) { return r + '岁'; }), datasets, { plugins: { legend: { display: true } } });
  } else {
    var counts = {};
    dataSets[0].data.forEach(function(d) { if (d['年龄']) { var g = classifyAge(d['年龄']); counts[g] = (counts[g] || 0) + 1; } });
    var data = ageRanges.map(function(r) { return counts[r] || 0; });
    createChart('chartAge', 'bar', ageRanges.map(function(r) { return r + '岁'; }), [{ data: data, backgroundColor: 'rgba(79,70,229,0.7)', borderColor: '#4F46E5', borderWidth: 1, borderRadius: 4 }], {});
  }
}

function renderPositionChart() {
  var counts = {};
  filteredData.forEach(function(d) { if (d['岗位']) { counts[d['岗位']] = (counts[d['岗位']] || 0) + 1; } });
  var sorted = Object.keys(counts).map(function(k) { return [k, counts[k]]; }).sort(function(a, b) { return b[1] - a[1]; }).slice(0, 15);
  var labels = sorted.map(function(s) { return s[0]; });
  var data = sorted.map(function(s) { return s[1]; });
  var bgColors = PALETTE.slice(0, data.length).map(function(c) { return c + 'BB'; });
  var bdColors = PALETTE.slice(0, data.length);
  createChart('chartPosition', 'bar', labels, [{ data: data, backgroundColor: bgColors, borderColor: bdColors, borderWidth: 1, borderRadius: 4 }], { indexAxis: 'y', plugins: { legend: { display: false } } });
}

function renderDeptChart() {
  var counts = {};
  filteredData.forEach(function(d) { if (d['部门']) { counts[d['部门']] = (counts[d['部门']] || 0) + 1; } });
  var sorted = Object.keys(counts).map(function(k) { return [k, counts[k]]; }).sort(function(a, b) { return b[1] - a[1]; });
  var labels = sorted.map(function(s) { return s[0]; });
  var data = sorted.map(function(s) { return s[1]; });
  createChart('chartDept', 'bar', labels, [{ data: data, backgroundColor: 'rgba(5,150,105,0.6)', borderColor: '#059669', borderWidth: 1, borderRadius: 4 }], { indexAxis: 'y', plugins: { legend: { display: false } } });
}

function renderTenureChart(dataSets, compare) {
  var ranges = ['<1年','1-3年','3-5年','5-10年','>10年'];
  function classifyTenure(t) {
    if (t < 1) return '<1年';
    if (t < 3) return '1-3年';
    if (t < 5) return '3-5年';
    if (t < 10) return '5-10年';
    return '>10年';
  }
  if (compare) {
    var datasets = dataSets.map(function(ds) {
      var counts = {};
      ds.data.forEach(function(d) { if (d['年限']) { var g = classifyTenure(d['年限']); counts[g] = (counts[g] || 0) + 1; } });
      return {
        label: ds.company,
        data: ranges.map(function(r) { return counts[r] || 0; }),
        backgroundColor: (COMPANY_COLORS[ds.company] ? COMPANY_COLORS[ds.company].bg : '#94A3B8'),
        borderColor: (COMPANY_COLORS[ds.company] ? COMPANY_COLORS[ds.company].border : '#64748B'),
        borderWidth: 1
      };
    });
    createChart('chartTenure', 'bar', ranges, datasets, { plugins: { legend: { display: true } } });
  } else {
    var counts = {};
    dataSets[0].data.forEach(function(d) { if (d['年限']) { var g = classifyTenure(d['年限']); counts[g] = (counts[g] || 0) + 1; } });
    var data = ranges.map(function(r) { return counts[r] || 0; });
    createChart('chartTenure', 'bar', ranges, [{ data: data, backgroundColor: 'rgba(217,119,6,0.7)', borderColor: '#D97706', borderWidth: 1, borderRadius: 4 }], {});
  }
}

function createChart(canvasId, type, labels, datasets, extraOpts) {
  if (charts[canvasId]) { charts[canvasId].destroy(); }
  var ctx = document.getElementById(canvasId).getContext('2d');
  var opts = {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: { display: false, labels: { font: { size: 12 }, usePointStyle: true, padding: 12 } },
      tooltip: { backgroundColor: '#1E293B', titleFont: { size: 13 }, bodyFont: { size: 12 }, padding: 10, cornerRadius: 8 }
    },
    scales: type === 'doughnut' ? {} : {
      x: { grid: { display: false }, ticks: { font: { size: 11 }, maxRotation: 45 } },
      y: { grid: { color: '#F1F5F9' }, ticks: { font: { size: 11 } }, beginAtZero: true }
    }
  };
  if (extraOpts) {
    if (extraOpts.indexAxis) opts.indexAxis = extraOpts.indexAxis;
    if (extraOpts.plugins) {
      Object.keys(extraOpts.plugins).forEach(function(k) { opts.plugins[k] = extraOpts.plugins[k]; });
    }
  }
  if (opts.indexAxis === 'y') {
    opts.scales.x.ticks = { maxRotation: 0 };
    opts.scales.y.ticks = { font: { size: 11 } };
  }
  charts[canvasId] = new Chart(ctx, { type: type, data: { labels: labels, datasets: datasets }, options: opts });
}

// ===== 表格渲染 =====
function renderTable() {
  var tbody = document.getElementById('tableBody');
  var countEl = document.getElementById('resultCount');
  countEl.textContent = '共 ' + filteredData.length + ' 条记录';

  function eduChip(e) {
    if (e === '博士') return 'chip-edu-doctor';
    if (e === '硕士') return 'chip-edu-master';
    if (e === '本科') return 'chip-edu-bachelor';
    if (e === '大专') return 'chip-edu-college';
    return 'chip-edu-other';
  }
  function genderChip(g) { return g === '男' ? 'chip-male' : 'chip-female'; }
  function companyColor(c) { return (COMPANY_COLORS[c] ? COMPANY_COLORS[c].border : '#64748B'); }

  var html = '';
  filteredData.forEach(function(d) {
    html += '<tr>' +
      '<td style="font-weight:600;color:' + companyColor(d['公司']) + '">' + (d['公司'] || '') + '</td>' +
      '<td style="font-weight:600">' + (d['姓名'] || '') + '</td>' +
      '<td><span class="chip ' + genderChip(d['性别']) + '">' + (d['性别'] || '') + '</span></td>' +
      '<td>' + (d['年龄'] || '-') + '</td>' +
      '<td><span class="chip ' + eduChip(d['学历']) + '">' + (d['学历'] || '') + '</span></td>' +
      '<td>' + (d['部门'] || '') + '</td>' +
      '<td>' + (d['岗位'] || '') + '</td>' +
      '<td>' + (d['职等序列'] || '') + '</td>' +
      '<td>' + (d['入职时间'] || '-') + '</td>' +
      '<td>' + (d['年限'] || '-') + '</td>' +
      '<td>' + (d['毕业学校'] || '') + '</td>' +
      '<td>' + (d['专业'] || '') + '</td>' +
    '</tr>';
  });
  tbody.innerHTML = html;
}

// ===== 重置筛选 =====
function resetFilters() {
  document.getElementById('filterCompany').value = '';
  document.getElementById('filterAge').value = '';
  document.getElementById('filterTenure').value = '';
  document.querySelectorAll('.company-tab').forEach(function(t) { t.classList.remove('active'); });
  var allTab = document.querySelector('.company-tab[data-company=""]');
  if (allTab) allTab.classList.add('active');
  updateDependentFilters();
  document.getElementById('filterDept').value = '';
  document.getElementById('filterEdu').value = '';
  document.getElementById('filterGender').value = '';
  document.getElementById('filterPosition').value = '';
  renderAll();
}

// ===== 导出 CSV =====
function exportCSV() {
  var headers = ['公司','姓名','性别','年龄','学历','部门','岗位','职等序列','入职时间','年限','毕业学校','专业'];
  var rows = filteredData.map(function(d) {
    return headers.map(function(h) {
      var v = d[h] || '';
      if (typeof v === 'string' && (v.indexOf(',') !== -1 || v.indexOf('"') !== -1)) v = '"' + v.replace(/"/g, '""') + '"';
      return v;
    }).join(',');
  });
  var csv = '\uFEFF' + headers.join(',') + '\n' + rows.join('\n');
  var blob = new Blob([csv], { type: 'text/csv;charset=utf-8' });
  var url = URL.createObjectURL(blob);
  var a = document.createElement('a');
  a.href = url;
  a.download = '在职人员信息_' + new Date().toISOString().slice(0,10) + '.csv';
  a.click();
  URL.revokeObjectURL(url);
}

// ===== 事件绑定 =====
['filterCompany','filterDept','filterEdu','filterGender','filterPosition','filterAge','filterTenure'].forEach(function(id) {
  document.getElementById(id).addEventListener('change', function() {
    if (id === 'filterCompany') {
      updateDependentFilters();
      document.querySelectorAll('.company-tab').forEach(function(t) { t.classList.remove('active'); });
      var target = document.querySelector('.company-tab[data-company="' + this.value + '"]');
      if (target) target.classList.add('active');
    }
    renderAll();
  });
});

// ===== 初始化 =====
setupDragDrop();
fetch('data.json')
  .then(function(r) { if (!r.ok) throw new Error('no data'); return r.json(); })
  .then(function(data) {
    allData = data;
    currentFileName = '在职人员信息.xlsx（默认）';
    showImportSuccess();
    initDataView();
  })
  .catch(function() { /* 无默认数据，等待用户导入 */ });
