const SOURCE_CLASS_ROW = 5;
const SOURCE_DATA_START_ROW = 6;
const MAX_SCORE = 100;

const DEFAULT_THRESHOLDS = {
  2015: [4, 11, 23, 40, 60, 77, 89, 96, 100],
  2022: [10, 34, 66, 90, 100],
};

const state = {
  fileName: '',
  rows: [],
  records: [],
  classes: [],
  maxNo: 0,
  recordMap: new Map(),
  names: new Map(),
  curriculum: '2015',
  results: [],
  stats: [],
  distribution: [],
  cutoffs: [],
  gradeCount: 9,
  currentIndex: -1,
};

const dom = {
  fileInput: document.querySelector('#sourceFile'),
  dropZone: document.querySelector('#dropZone'),
  loadedCount: document.querySelector('#loadedCount'),
  classCount: document.querySelector('#classCount'),
  nameProgress: document.querySelector('#nameProgress'),
  statusMessage: document.querySelector('#statusMessage'),
  rosterShell: document.querySelector('#rosterShell'),
  thresholdGrid: document.querySelector('#thresholdGrid'),
  metricsGrid: document.querySelector('#metricsGrid'),
  resultCountLabel: document.querySelector('#resultCountLabel'),
  resultsShell: document.querySelector('#resultsShell'),
  classStatsShell: document.querySelector('#classStatsShell'),
  gradeDistributionShell: document.querySelector('#gradeDistributionShell'),
  cutoffShell: document.querySelector('#cutoffShell'),
  histogram: document.querySelector('#histogram'),
  binWidth: document.querySelector('#binWidth'),
  lookupForm: document.querySelector('#lookupForm'),
  searchClass: document.querySelector('#searchClass'),
  searchNo: document.querySelector('#searchNo'),
  reportCard: document.querySelector('#reportCard'),
  printRoot: document.querySelector('#printRoot'),
};

document.addEventListener('click', handleDocumentClick);
document.addEventListener('input', handleDocumentInput);
document.addEventListener('paste', handleRosterPaste);
dom.fileInput.addEventListener('change', handleFileInput);
dom.lookupForm.addEventListener('submit', handleLookupSubmit);
dom.binWidth.addEventListener('input', () => {
  if (state.results.length) renderHistogram();
});

['dragenter', 'dragover'].forEach((eventName) => {
  dom.dropZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dom.dropZone.classList.add('dragging');
  });
});

['dragleave', 'drop'].forEach((eventName) => {
  dom.dropZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dom.dropZone.classList.remove('dragging');
  });
});

dom.dropZone.addEventListener('drop', (event) => {
  const file = event.dataTransfer.files[0];
  if (file) loadFile(file);
});

renderThresholds();
renderAll();

function handleDocumentClick(event) {
  const curriculumEl = event.target.closest('[data-curriculum]');
  if (curriculumEl) {
    setCurriculum(curriculumEl.dataset.curriculum);
    return;
  }

  const actionEl = event.target.closest('[data-action]');
  if (!actionEl) return;

  const action = actionEl.dataset.action;
  if (action === 'sample') loadSampleData();
  if (action === 'analyze') analyze();
  if (action === 'check-names') checkNames();
  if (action === 'clear-names') clearNames();
  if (action === 'download-csv') downloadResultsCsv();
  if (action === 'prev-student') moveStudent(-1);
  if (action === 'next-student') moveStudent(1);
  if (action === 'print-current') printCurrentReport();
  if (action === 'print-all') printAllReports();
}

function handleDocumentInput(event) {
  if (event.target.matches('.name-input')) {
    const key = keyOf(event.target.dataset.class, event.target.dataset.no);
    state.names.set(key, event.target.value.trim());
    updateNameProgress();
  }
}

function handleRosterPaste(event) {
  const target = event.target;
  if (!target.matches('.name-input')) return;

  const text = event.clipboardData.getData('text');
  if (!text.includes('\t') && !text.includes('\n')) return;

  event.preventDefault();
  const startClass = Number(target.dataset.class);
  const startNo = Number(target.dataset.no);
  const startClassIndex = state.classes.indexOf(startClass);
  const pasted = parseDelimitedRows(text);

  pasted.forEach((row, rowOffset) => {
    row.forEach((value, colOffset) => {
      const cls = state.classes[startClassIndex + colOffset];
      const no = startNo + rowOffset;
      if (!cls || !no) return;
      const key = keyOf(cls, no);
      if (state.recordMap.has(key)) state.names.set(key, String(value).trim());
    });
  });

  renderRoster();
  updateNameProgress();
  setStatus('붙여넣은 명렬을 반영했습니다.');
}

async function handleFileInput(event) {
  const file = event.target.files[0];
  if (file) await loadFile(file);
}

async function loadFile(file) {
  try {
    setStatus('성적자료를 읽는 중입니다.');
    const rows = await readWorkbookRows(file);
    setStateFromRows(rows, file.name);
    renderAll();
    setStatus(`${file.name}에서 ${state.records.length}건의 성적 데이터를 불러왔습니다.`);
    document.querySelector('#roster').scrollIntoView({ behavior: 'smooth', block: 'start' });
  } catch (error) {
    setStatus(error.message || '성적자료를 읽지 못했습니다.');
  } finally {
    dom.fileInput.value = '';
  }
}

async function readWorkbookRows(file) {
  const ext = file.name.split('.').pop().toLowerCase();
  const buffer = await file.arrayBuffer();

  if (['xlsx', 'xls'].includes(ext)) {
    if (!window.XLSX) {
      throw new Error('XLSX 해석 라이브러리를 불러오지 못했습니다. 네트워크 연결을 확인하거나 CSV/TSV로 저장해 다시 시도하세요.');
    }

    const workbook = window.XLSX.read(buffer, { type: 'array', cellDates: false });
    const firstSheetName = workbook.SheetNames[0];
    if (!firstSheetName) throw new Error('통합문서에서 시트를 찾을 수 없습니다.');

    return window.XLSX.utils.sheet_to_json(workbook.Sheets[firstSheetName], {
      header: 1,
      raw: false,
      blankrows: false,
      defval: '',
    });
  }

  const text = decodeText(buffer);
  return parseDelimitedRows(text);
}

function decodeText(buffer) {
  const utf8 = new TextDecoder('utf-8').decode(buffer);
  let eucKr = '';

  try {
    eucKr = new TextDecoder('euc-kr').decode(buffer);
  } catch {
    return utf8;
  }

  return koreanScore(eucKr) > koreanScore(utf8) ? eucKr : utf8;
}

function koreanScore(text) {
  const matches = String(text).match(/[가-힣]/g);
  return matches ? matches.length : 0;
}

function parseDelimitedRows(text) {
  const sample = String(text).slice(0, 4096);
  const delimiter = sample.includes('\t') ? '\t' : ',';
  const rows = [];
  let row = [];
  let cell = '';
  let quoted = false;

  for (let i = 0; i < text.length; i += 1) {
    const char = text[i];
    const next = text[i + 1];

    if (char === '"') {
      if (quoted && next === '"') {
        cell += '"';
        i += 1;
      } else {
        quoted = !quoted;
      }
    } else if (char === delimiter && !quoted) {
      row.push(cell);
      cell = '';
    } else if ((char === '\n' || char === '\r') && !quoted) {
      if (char === '\r' && next === '\n') i += 1;
      row.push(cell);
      rows.push(row);
      row = [];
      cell = '';
    } else {
      cell += char;
    }
  }

  if (cell.length || row.length) {
    row.push(cell);
    rows.push(row);
  }

  return rows.filter((line) => line.some((value) => String(value).trim().length));
}

function setStateFromRows(rows, fileName) {
  if (!Array.isArray(rows) || rows.length < SOURCE_DATA_START_ROW) {
    throw new Error('원본 파일에서 학급/번호/점수 구조를 찾을 수 없습니다.');
  }

  const classRow = rows[SOURCE_CLASS_ROW - 1] || [];
  const classSet = new Set();
  const records = [];
  let maxNo = 0;

  for (let col = 1; col < classRow.length; col += 1) {
    const classNo = toWholeNumber(classRow[col]);
    if (!Number.isFinite(classNo)) continue;

    classSet.add(classNo);

    for (let row = SOURCE_DATA_START_ROW - 1; row < rows.length; row += 1) {
      const sourceRow = rows[row] || [];
      const studentNo = toWholeNumber(sourceRow[0]);
      const score = toNumber(sourceRow[col]);

      if (!Number.isFinite(studentNo) || !Number.isFinite(score)) continue;

      records.push({
        classNo,
        studentNo,
        score,
      });

      if (studentNo > maxNo) maxNo = studentNo;
    }
  }

  if (!records.length) throw new Error('불러올 수 있는 점수 데이터가 없습니다.');

  const nextNames = new Map();
  records.forEach((record) => {
    const key = keyOf(record.classNo, record.studentNo);
    if (state.names.has(key)) nextNames.set(key, state.names.get(key));
  });

  state.fileName = fileName;
  state.rows = rows;
  state.records = records;
  state.classes = [...classSet].sort((a, b) => a - b);
  state.maxNo = maxNo;
  state.recordMap = new Map(records.map((record) => [keyOf(record.classNo, record.studentNo), record]));
  state.names = nextNames;
  state.results = [];
  state.stats = [];
  state.distribution = [];
  state.cutoffs = [];
  state.currentIndex = -1;
}

function toWholeNumber(value) {
  const number = toNumber(value);
  return Number.isFinite(number) && number > 0 ? Math.trunc(number) : Number.NaN;
}

function toNumber(value) {
  if (typeof value === 'number') return Number.isFinite(value) ? value : Number.NaN;
  const text = String(value ?? '').replace(/,/g, '').trim();
  if (!text) return Number.NaN;
  const number = Number(text);
  return Number.isFinite(number) ? number : Number.NaN;
}

function loadSampleData() {
  const rows = Array.from({ length: 36 }, () => Array(6).fill(''));
  rows[4] = ['번호', 1, 2, 3, 4, ''];

  for (let no = 1; no <= 30; no += 1) {
    const rowIndex = no + 4;
    rows[rowIndex][0] = no;
    for (let col = 1; col <= 4; col += 1) {
      const base = 58 + ((no * 7 + col * 11) % 39);
      const tieBump = no % 6 === 0 ? 0 : (col % 2) * 0.5;
      rows[rowIndex][col] = Math.min(100, base + tieBump);
    }
  }

  setStateFromRows(rows, 'sample-neis-grade-data.xlsx');

  state.records.forEach((record) => {
    state.names.set(keyOf(record.classNo, record.studentNo), `${record.classNo}반 ${record.studentNo}번`);
  });

  renderAll();
  setStatus('샘플 성적자료와 학생명렬을 불러왔습니다. 분석 시작을 눌러 결과를 확인하세요.');
}

function setCurriculum(curriculum) {
  if (!DEFAULT_THRESHOLDS[curriculum]) return;
  state.curriculum = curriculum;
  state.gradeCount = DEFAULT_THRESHOLDS[curriculum].length;
  renderThresholds();

  document.querySelectorAll('[data-curriculum]').forEach((button) => {
    const active = button.dataset.curriculum === curriculum;
    button.classList.toggle('active', active);
    button.setAttribute('aria-pressed', String(active));
  });

  setStatus(`${curriculum} 개정 기준을 선택했습니다.`);
}

function renderThresholds() {
  const thresholds = DEFAULT_THRESHOLDS[state.curriculum];
  dom.thresholdGrid.innerHTML = thresholds.map((value, index) => `
    <label>
      ${index + 1}등급
      <input class="threshold-input" type="number" min="0" max="100" step="0.01" value="${value}" data-grade="${index + 1}">
    </label>
  `).join('');
}

function renderAll() {
  renderRoster();
  renderAnalysisShells();
  renderReport(null);
  updateNameProgress();
}

function renderRoster() {
  if (!state.records.length) {
    dom.rosterShell.innerHTML = '<div class="empty-state">성적자료를 먼저 불러오세요.</div>';
    return;
  }

  const header = state.classes.map((classNo) => `<th>${classNo}반</th>`).join('');
  const rows = [];

  for (let no = 1; no <= state.maxNo; no += 1) {
    const cells = state.classes.map((classNo) => {
      const key = keyOf(classNo, no);
      if (!state.recordMap.has(key)) return '<td class="inactive-cell">-</td>';

      const value = state.names.get(key) || '';
      const missing = value.trim() ? '' : ' is-missing';
      return `<td><input class="name-input${missing}" data-class="${classNo}" data-no="${no}" value="${escapeAttr(value)}" aria-label="${classNo}반 ${no}번 이름"></td>`;
    }).join('');

    rows.push(`<tr><td>${no}번</td>${cells}</tr>`);
  }

  dom.rosterShell.innerHTML = `
    <table class="roster-table">
      <thead><tr><th>번호 / 반</th>${header}</tr></thead>
      <tbody>${rows.join('')}</tbody>
    </table>
  `;
}

function updateNameProgress() {
  const total = state.records.length;
  const filled = state.records.filter((record) => {
    const name = state.names.get(keyOf(record.classNo, record.studentNo));
    return Boolean(name && name.trim());
  }).length;

  dom.loadedCount.textContent = `${total}명`;
  dom.classCount.textContent = `${state.classes.length}개`;
  dom.nameProgress.textContent = `${filled} / ${total}`;
}

function checkNames() {
  if (!state.records.length) {
    setStatus('성적자료를 먼저 불러오세요.');
    return false;
  }

  const missing = getMissingRecords();
  renderRoster();

  if (missing.length) {
    setStatus(`${missing.length}개의 이름 입력 칸이 비어 있습니다. 파란 테두리 칸을 확인하세요.`);
    return false;
  }

  setStatus('필수 이름 입력이 모두 완료되었습니다.');
  return true;
}

function clearNames() {
  state.names = new Map();
  renderRoster();
  updateNameProgress();
  setStatus('학생명렬을 비웠습니다.');
}

function getMissingRecords() {
  return state.records.filter((record) => {
    const name = state.names.get(keyOf(record.classNo, record.studentNo));
    return !name || !name.trim();
  });
}

function analyze() {
  try {
    if (!state.records.length) {
      setStatus('성적자료를 먼저 불러오세요.');
      return;
    }

    const missing = getMissingRecords();
    if (missing.length && !window.confirm(`이름이 비어 있는 학생 ${missing.length}명은 빈칸으로 분석됩니다. 계속할까요?`)) {
      checkNames();
      return;
    }

    const thresholds = getThresholdsFromInputs();
    const gradeCount = thresholds.length;
    const cutoffRanks = buildGradeCutoffRanks(state.records.length, thresholds);
    const ranked = rankRecords(state.records);
    const stats = new Map();

    state.results = ranked.map((record) => {
      const statKey = String(record.classNo);
      const current = stats.get(statKey) || { classNo: record.classNo, count: 0, sum: 0, max: -1, min: 101 };
      current.count += 1;
      current.sum += record.score;
      current.max = Math.max(current.max, record.score);
      current.min = Math.min(current.min, record.score);
      stats.set(statKey, current);

      const midRank = record.rank + (record.tie - 1) / 2;
      const percent = round2((midRank / state.records.length) * 100);

      return {
        ...record,
        name: state.names.get(keyOf(record.classNo, record.studentNo)) || '',
        midRank,
        percent,
        grade: getGradeFromCutoffRank(midRank, cutoffRanks),
      };
    });

    state.gradeCount = gradeCount;
    state.stats = buildClassStats(stats);
    state.distribution = buildGradeDistribution(state.results, gradeCount);
    state.cutoffs = buildCutoffs(state.results, gradeCount);
    state.currentIndex = -1;

    renderAnalysisShells();
    renderReport(null);
    setStatus('분석 완료. 학생조회에서 개별 성적표를 확인하세요.');
    document.querySelector('#analysis').scrollIntoView({ behavior: 'smooth', block: 'start' });
  } catch (error) {
    setStatus(error.message || '분석 중 오류가 발생했습니다.');
  }
}

function getThresholdsFromInputs() {
  const inputs = [...document.querySelectorAll('.threshold-input')];
  const thresholds = inputs.map((input) => {
    const number = Number(input.value);
    if (!Number.isFinite(number)) throw new Error('환경설정의 누적퍼센트 값이 숫자가 아닙니다.');
    return number / 100;
  });

  if (!thresholds.length) throw new Error('등급 기준을 찾을 수 없습니다.');
  return thresholds;
}

function buildGradeCutoffRanks(studentCount, thresholds) {
  const cutoffRanks = [];
  let previous = 0;

  thresholds.forEach((threshold, index) => {
    let cutoffRank = index === thresholds.length - 1 ? studentCount : Math.round(studentCount * threshold);
    cutoffRank = Math.max(0, Math.min(studentCount, cutoffRank));
    cutoffRank = Math.max(previous, cutoffRank);
    cutoffRanks.push(cutoffRank);
    previous = cutoffRank;
  });

  return cutoffRanks;
}

function rankRecords(records) {
  const scoreCounts = new Map();

  records.forEach((record) => {
    const key = String(record.score);
    scoreCounts.set(key, (scoreCounts.get(key) || 0) + 1);
  });

  const ranks = new Map();
  let greater = 0;

  [...scoreCounts.keys()]
    .map(Number)
    .sort((a, b) => b - a)
    .forEach((score) => {
      const key = String(score);
      const tie = scoreCounts.get(key);
      ranks.set(key, { rank: greater + 1, tie });
      greater += tie;
    });

  return records.map((record) => ({
    ...record,
    rank: ranks.get(String(record.score)).rank,
    tie: ranks.get(String(record.score)).tie,
  }));
}

function getGradeFromCutoffRank(midRank, cutoffRanks) {
  for (let i = 0; i < cutoffRanks.length; i += 1) {
    if (midRank <= cutoffRanks[i] + 0.00000001) return i + 1;
  }
  return cutoffRanks.length;
}

function buildClassStats(statsMap) {
  const rows = [...statsMap.values()]
    .sort((a, b) => a.classNo - b.classNo)
    .map((item) => ({
      classNo: item.classNo,
      count: item.count,
      average: item.count ? round2(item.sum / item.count) : 0,
      max: item.max === -1 ? null : item.max,
      min: item.min === 101 ? null : item.min,
    }));

  const total = rows.reduce((acc, row) => {
    acc.count += row.count;
    acc.sum += row.average * row.count;
    acc.max = Math.max(acc.max, row.max ?? -1);
    acc.min = Math.min(acc.min, row.min ?? 101);
    return acc;
  }, { count: 0, sum: 0, max: -1, min: 101 });

  rows.push({
    classNo: '전체',
    count: total.count,
    average: total.count ? round2(total.sum / total.count) : 0,
    max: total.max === -1 ? null : total.max,
    min: total.min === 101 ? null : total.min,
  });

  return rows;
}

function buildGradeDistribution(results, gradeCount) {
  const counts = Array.from({ length: gradeCount }, (_, index) => ({
    grade: index + 1,
    count: 0,
    ratio: 0,
  }));

  results.forEach((result) => {
    if (result.grade >= 1 && result.grade <= gradeCount) counts[result.grade - 1].count += 1;
  });

  counts.forEach((row) => {
    row.ratio = results.length ? round2((row.count / results.length) * 100) : 0;
  });

  return counts;
}

function buildCutoffs(results, gradeCount) {
  return Array.from({ length: gradeCount }, (_, index) => {
    const grade = index + 1;
    const scores = results.filter((result) => result.grade === grade).map((result) => result.score);
    return {
      grade,
      cutoff: scores.length ? Math.min(...scores) : null,
    };
  });
}

function renderAnalysisShells() {
  renderMetrics();
  renderResultsTable();
  renderClassStats();
  renderGradeDistribution();
  renderCutoffs();
  renderHistogram();
}

function renderMetrics() {
  if (!state.results.length) {
    dom.metricsGrid.innerHTML = `
      <div class="store-utility-card metric-card"><span>응시인원</span><strong>0명</strong></div>
      <div class="store-utility-card metric-card"><span>전체평균</span><strong>-</strong></div>
      <div class="store-utility-card metric-card"><span>최고점</span><strong>-</strong></div>
      <div class="store-utility-card metric-card"><span>최저점</span><strong>-</strong></div>
    `;
    return;
  }

  const scores = state.results.map((result) => result.score);
  const average = round2(scores.reduce((sum, score) => sum + score, 0) / scores.length);

  dom.metricsGrid.innerHTML = `
    <div class="store-utility-card metric-card"><span>응시인원</span><strong>${state.results.length}명</strong></div>
    <div class="store-utility-card metric-card"><span>전체평균</span><strong>${formatNumber(average)}</strong></div>
    <div class="store-utility-card metric-card"><span>최고점</span><strong>${formatNumber(Math.max(...scores))}</strong></div>
    <div class="store-utility-card metric-card"><span>최저점</span><strong>${formatNumber(Math.min(...scores))}</strong></div>
  `;
}

function renderResultsTable() {
  dom.resultCountLabel.textContent = `${state.results.length}건`;

  if (!state.results.length) {
    dom.resultsShell.innerHTML = '<div class="empty-state">분석을 시작하면 결과가 표시됩니다.</div>';
    return;
  }

  const rows = state.results.map((result) => `
    <tr>
      <td>${result.classNo}</td>
      <td>${result.studentNo}</td>
      <td>${escapeHtml(result.name)}</td>
      <td>${formatNumber(result.score)}</td>
      <td>${result.rank}</td>
      <td>${formatNumber(result.midRank)}</td>
      <td>${formatNumber(result.percent)}%</td>
      <td>${result.grade}</td>
      <td>${result.tie}</td>
    </tr>
  `).join('');

  dom.resultsShell.innerHTML = `
    <table>
      <thead>
        <tr>
          <th>학급</th><th>번호</th><th>성명</th><th>점수</th><th>석차</th>
          <th>중간석차</th><th>백분율</th><th>석차등급</th><th>동석차수</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  `;
}

function renderClassStats() {
  if (!state.stats.length) {
    dom.classStatsShell.innerHTML = '<div class="empty-state">분석 후 표시됩니다.</div>';
    return;
  }

  const rows = state.stats.map((row) => `
    <tr>
      <td>${escapeHtml(String(row.classNo))}${row.classNo === '전체' ? '' : '반'}</td>
      <td>${row.count}</td>
      <td>${formatNumber(row.average)}</td>
      <td>${formatNullable(row.max)}</td>
      <td>${formatNullable(row.min)}</td>
    </tr>
  `).join('');

  dom.classStatsShell.innerHTML = `
    <table>
      <thead><tr><th>학급</th><th>응시인원</th><th>평균</th><th>최고점</th><th>최저점</th></tr></thead>
      <tbody>${rows}</tbody>
    </table>
  `;
}

function renderGradeDistribution() {
  if (!state.distribution.length) {
    dom.gradeDistributionShell.innerHTML = '<div class="empty-state">분석 후 표시됩니다.</div>';
    return;
  }

  const rows = state.distribution.map((row) => `
    <tr>
      <td>${row.grade}등급</td>
      <td>${row.count}</td>
      <td>${formatNumber(row.ratio)}%</td>
    </tr>
  `).join('');

  dom.gradeDistributionShell.innerHTML = `
    <table>
      <thead><tr><th>등급</th><th>인원</th><th>비율</th></tr></thead>
      <tbody>${rows}</tbody>
    </table>
  `;
}

function renderCutoffs() {
  if (!state.cutoffs.length) {
    dom.cutoffShell.innerHTML = '<div class="empty-state">분석 후 표시됩니다.</div>';
    return;
  }

  const rows = state.cutoffs.map((row) => `
    <tr>
      <td>${row.grade}등급</td>
      <td>${formatNullable(row.cutoff)}</td>
    </tr>
  `).join('');

  dom.cutoffShell.innerHTML = `
    <table>
      <thead><tr><th>등급</th><th>커트라인</th></tr></thead>
      <tbody>${rows}</tbody>
    </table>
  `;
}

function renderHistogram() {
  if (!state.results.length) {
    dom.histogram.innerHTML = '<div class="empty-state">점수 분포가 여기에 그려집니다.</div>';
    return;
  }

  const bins = buildHistogramBins();
  const maxFreq = Math.max(1, ...bins.map((bin) => bin.count));
  const width = 640;
  const height = 360;
  const plot = { left: 48, top: 40, width: 552, height: 246 };
  const gap = 4;
  const barWidth = Math.max(3, (plot.width - gap * (bins.length - 1)) / bins.length);
  const labelStep = bins.length <= 12 ? 1 : Math.ceil(bins.length / 10);

  const bars = bins.map((bin, index) => {
    const barHeight = bin.count ? Math.max(1, (bin.count / maxFreq) * plot.height) : 0;
    const x = plot.left + index * (barWidth + gap);
    const y = plot.top + plot.height - barHeight;
    const label = index % labelStep === 0 ? `<text x="${x + barWidth / 2}" y="${plot.top + plot.height + 24}" text-anchor="middle" font-size="11" fill="#7a7a7a">${bin.start}</text>` : '';
    const count = bin.count ? `<text x="${x + barWidth / 2}" y="${y - 8}" text-anchor="middle" font-size="11" fill="#333333">${bin.count}</text>` : '';
    return `
      <rect x="${x}" y="${y}" width="${barWidth}" height="${barHeight}" fill="#0066cc"></rect>
      ${count}
      ${label}
    `;
  }).join('');

  dom.histogram.innerHTML = `
    <svg viewBox="0 0 ${width} ${height}" role="img" aria-label="점수 분포도">
      <rect x="0" y="0" width="${width}" height="${height}" rx="8" fill="#ffffff"></rect>
      <text x="24" y="24" font-size="14" font-weight="600" fill="#1d1d1f">점수 분포도 (계급: ${getBinWidth()})</text>
      <line x1="${plot.left}" y1="${plot.top + plot.height}" x2="${plot.left + plot.width}" y2="${plot.top + plot.height}" stroke="#7a7a7a"></line>
      <line x1="${plot.left}" y1="${plot.top}" x2="${plot.left}" y2="${plot.top + plot.height}" stroke="#7a7a7a"></line>
      <text x="20" y="${plot.top + 4}" font-size="11" fill="#7a7a7a">${maxFreq}</text>
      <text x="28" y="${plot.top + plot.height + 4}" font-size="11" fill="#7a7a7a">0</text>
      ${bars}
    </svg>
  `;
}

function buildHistogramBins() {
  const binWidth = getBinWidth();
  const map = new Map();

  for (let start = 0; start <= MAX_SCORE; start += binWidth) {
    const end = Math.min(MAX_SCORE, start + binWidth - 0.1);
    map.set(start, {
      start,
      label: `${formatNumber(start)}~${formatNumber(end)}`,
      count: 0,
    });
  }

  state.results.forEach((result) => {
    const score = Math.max(0, Math.min(MAX_SCORE, result.score));
    let start = Math.floor(score / binWidth) * binWidth;
    if (start > MAX_SCORE) start = MAX_SCORE;
    if (!map.has(start)) map.set(start, { start, label: String(start), count: 0 });
    map.get(start).count += 1;
  });

  return [...map.values()].sort((a, b) => a.start - b.start);
}

function getBinWidth() {
  let width = Math.trunc(Number(dom.binWidth.value));
  if (!Number.isFinite(width) || width <= 0) width = 10;
  width = Math.max(1, Math.min(MAX_SCORE, width));
  dom.binWidth.value = width;
  return width;
}

function handleLookupSubmit(event) {
  event.preventDefault();
  searchStudent();
}

function searchStudent() {
  if (!state.results.length) {
    setStatus('분석을 먼저 완료하세요.');
    return;
  }

  const classNo = Number(dom.searchClass.value);
  const studentNo = Number(dom.searchNo.value);

  if (!Number.isFinite(classNo) || !Number.isFinite(studentNo)) {
    setStatus('학급과 번호를 모두 입력하세요.');
    return;
  }

  const index = findStudentIndex(classNo, studentNo);
  if (index < 0) {
    state.currentIndex = -1;
    renderReport(null);
    setStatus('학생을 찾을 수 없습니다.');
    return;
  }

  showStudent(index);
}

function moveStudent(direction) {
  if (!state.results.length) {
    setStatus('분석을 먼저 완료하세요.');
    return;
  }

  let nextIndex;
  if (state.currentIndex < 0) {
    nextIndex = direction >= 0 ? 0 : state.results.length - 1;
  } else {
    nextIndex = state.currentIndex + direction;
  }

  if (nextIndex < 0 || nextIndex >= state.results.length) {
    setStatus('이동할 학생이 없습니다.');
    return;
  }

  showStudent(nextIndex);
}

function showStudent(index) {
  const student = state.results[index];
  state.currentIndex = index;
  dom.searchClass.value = student.classNo;
  dom.searchNo.value = student.studentNo;
  renderReport(student);
  setStatus(`${student.classNo}반 ${student.studentNo}번을 조회했습니다.`);
}

function findStudentIndex(classNo, studentNo) {
  return state.results.findIndex((result) => result.classNo === Math.trunc(classNo) && result.studentNo === Math.trunc(studentNo));
}

function renderReport(student) {
  if (!student) {
    dom.reportCard.innerHTML = '<div class="report-empty">분석 후 학생을 조회하면 성적표가 표시됩니다.</div>';
    return;
  }

  dom.reportCard.innerHTML = `
    <div class="report-header">
      <div>
        <h3>${escapeHtml(student.name || '이름 없음')}</h3>
        <p>${student.classNo}반 ${student.studentNo}번</p>
      </div>
      <span class="report-pill">${student.grade}등급</span>
    </div>
    <div class="report-grid">
      ${reportMetric('원점수', `${formatNumber(student.score)}점`)}
      ${reportMetric('전체석차', `${student.rank}등`)}
      ${reportMetric('상위 퍼센트', `${formatNumber(student.percent)}%`)}
      ${reportMetric('중간석차', formatNumber(student.midRank))}
      ${reportMetric('동석차 인원', `${student.tie}명`)}
      ${reportMetric('교육과정', `${state.curriculum} 개정`)}
    </div>
    <h3>등급컷</h3>
    ${cutoffTableHtml()}
  `;
}

function reportMetric(label, value) {
  return `<div class="report-metric"><span>${label}</span><strong>${escapeHtml(String(value))}</strong></div>`;
}

function printCurrentReport() {
  if (!state.results.length) {
    setStatus('분석을 먼저 완료하세요.');
    return;
  }

  const student = state.currentIndex >= 0 ? state.results[state.currentIndex] : null;
  if (!student) {
    setStatus('먼저 조회할 학생의 학급과 번호를 입력하세요.');
    return;
  }

  printReports([student]);
}

function printAllReports() {
  if (!state.results.length) {
    setStatus('분석 결과가 없습니다.');
    return;
  }

  printReports(state.results);
}

function printReports(students) {
  dom.printRoot.innerHTML = students.map(printPageHtml).join('');
  window.print();
  setStatus(`${students.length}명의 성적표를 인쇄/PDF 화면으로 보냈습니다.`);
}

function printPageHtml(student) {
  return `
    <section class="print-page">
      <h1 class="print-title">학생 성적 개별 조회</h1>
      <table class="print-table">
        <tbody>
          <tr><th>성명 (학급/번호)</th><td>${escapeHtml(student.name || '이름 없음')} (${student.classNo}-${student.studentNo})</td></tr>
          <tr><th>원점수</th><td>${formatNumber(student.score)}점</td></tr>
          <tr><th>전체석차</th><td>${student.rank}등</td></tr>
          <tr><th>상위 퍼센트</th><td>${formatNumber(student.percent)}%</td></tr>
          <tr><th>석차등급</th><td>${student.grade}등급</td></tr>
          <tr><th>동석차 인원</th><td>${student.tie}명</td></tr>
        </tbody>
      </table>
      <table class="print-table">
        <thead><tr><th>등급</th><th>커트라인</th></tr></thead>
        <tbody>${state.cutoffs.map((row) => `<tr><td>${row.grade}등급</td><td>${formatNullable(row.cutoff)}</td></tr>`).join('')}</tbody>
      </table>
    </section>
  `;
}

function downloadResultsCsv() {
  if (!state.results.length) {
    setStatus('저장할 분석 결과가 없습니다.');
    return;
  }

  const header = ['학급', '번호', '성명', '점수', '석차', '중간석차', '백분율(%)', '석차등급', '동석차수'];
  const rows = state.results.map((result) => [
    result.classNo,
    result.studentNo,
    result.name,
    result.score,
    result.rank,
    result.midRank,
    result.percent,
    result.grade,
    result.tie,
  ]);
  const csv = [header, ...rows].map((row) => row.map(csvEscape).join(',')).join('\n');
  const blob = new Blob([`\uFEFF${csv}`], { type: 'text/csv;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = `grade_analysis_${timestamp()}.csv`;
  document.body.append(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
  setStatus('분석 결과 CSV를 저장했습니다.');
}

function cutoffTableHtml() {
  if (!state.cutoffs.length) return '<div class="empty-state">등급컷이 없습니다.</div>';

  return `
    <div class="table-shell compact">
      <table>
        <thead><tr><th>등급</th><th>커트라인</th></tr></thead>
        <tbody>
          ${state.cutoffs.map((row) => `<tr><td>${row.grade}등급</td><td>${formatNullable(row.cutoff)}</td></tr>`).join('')}
        </tbody>
      </table>
    </div>
  `;
}

function setStatus(message) {
  dom.statusMessage.textContent = message;
}

function keyOf(classNo, studentNo) {
  return `${Math.trunc(Number(classNo))}|${Math.trunc(Number(studentNo))}`;
}

function round2(value) {
  return Math.round((Number(value) + Number.EPSILON) * 100) / 100;
}

function formatNullable(value) {
  return value === null || value === undefined ? '-' : formatNumber(value);
}

function formatNumber(value) {
  if (!Number.isFinite(Number(value))) return '-';
  const number = Number(value);
  return Number.isInteger(number) ? String(number) : String(round2(number));
}

function csvEscape(value) {
  const text = String(value ?? '');
  return /[",\n\r]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
}

function timestamp() {
  const date = new Date();
  const pad = (value) => String(value).padStart(2, '0');
  return [
    date.getFullYear(),
    pad(date.getMonth() + 1),
    pad(date.getDate()),
    '_',
    pad(date.getHours()),
    pad(date.getMinutes()),
  ].join('');
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function escapeAttr(value) {
  return escapeHtml(value).replace(/`/g, '&#096;');
}
