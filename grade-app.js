/**
 * grade-app.js — v3.0
 *
 * UI 구조
 * ────────────────────────────────────
 * [반 선택 ▼] [교재 선택 ▼]  [🔲엑셀 | 👤카드] [📋성적표]
 *
 * ─── 엑셀 모드 ───────────────────────────────────
 * 좌측: 학생 리스트(이름/닉네임)  우측: 가로 스프레드시트
 *   1. 단어 평가: 총문제수 | 재시험 | 통과 | 성취율
 *   2. 리딩 평가: 총문제수 | [정답수 - R1~Rn] | [점수 - R1~Rn] | 성취율
 *   3. Teacher's Comment
 *
 * ─── 카드 모드 ───────────────────────────────────
 * 좌측: 학생 리스트  우측: 선택 학생 독립 입력 카드
 *
 * 계산 규칙
 *   통과 = 총문제수 - 재시험
 *   성취율(단어) = 통과 / 총문제수 × 100 (%)
 *   점수(리딩) = 정답수 / 총문제수 × 100 (소수 1자리, % 표시 없음)
 *   성취율(리딩) = 점수들 평균 (%)
 */
const GradeApp = (() => {
  /* ══ 상태 ══ */
  let _st = {
    classId:   null,
    bookId:    null,
    studentId: null,   // 카드 모드 선택 학생
    viewMode:  'excel', // 'excel' | 'card'
    data: {},          // { studentId: { word:{totalQ,retake,pass}, reading:{R0:{correct,score},...}, comment } }
    saving: new Set(), // 저장 중인 studentId
  };

  /* ══ CSS ══ */
  function _css() {
    if (document.getElementById('gr-styles')) return;
    const s = document.createElement('style');
    s.id = 'gr-styles';
    s.textContent = `
/* ── Page ── */
#page-grade { display:none; flex-direction:column; height:100%; overflow:hidden; }
#page-grade.on { display:flex; }

/* ── 상단 툴바 ── */
.gr-toolbar {
  display:flex; align-items:center; gap:8px; padding:9px 14px;
  background:var(--surf); border-bottom:1.5px solid var(--bdr); flex-shrink:0; flex-wrap:wrap;
}
.gr-sel {
  padding:8px 10px; border-radius:10px; background:var(--surf2);
  border:1.5px solid var(--bdr); font-size:13px; color:var(--tx);
  outline:none; cursor:pointer; font-family:var(--font);
  -webkit-appearance:none; transition:border-color .2s; flex:1; min-width:110px;
}
.gr-sel:focus { border-color:var(--a); }
.gr-sel:disabled { opacity:.5; cursor:default; }

/* 뷰 토글 */
.gr-view-toggle {
  display:flex; border-radius:10px; overflow:hidden;
  border:1.5px solid var(--bdr2); flex-shrink:0;
}
.gr-vbtn {
  padding:7px 12px; font-size:12px; font-weight:700; cursor:pointer;
  background:var(--card2); color:var(--tx3); border:none; font-family:var(--font);
  transition:all .15s; white-space:nowrap;
}
.gr-vbtn.on { background:var(--a); color:#fff; }
.gr-vbtn:active { opacity:.8; }

/* ── 본문 2컬럼 ── */
.gr-main {
  flex:1; display:flex; overflow:hidden;
}

/* 좌측 학생 리스트 */
.gr-stu-panel {
  width:96px; flex-shrink:0; border-right:1.5px solid var(--bdr);
  overflow-y:auto; -webkit-overflow-scrolling:touch; background:var(--surf);
}
.gr-stu-panel::-webkit-scrollbar { width:3px; }
.gr-stu-panel::-webkit-scrollbar-thumb { background:var(--bdr2); border-radius:2px; }
.gr-stu-item {
  padding:10px 8px; border-bottom:1px solid var(--bdr);
  cursor:pointer; transition:background .12s; text-align:center;
}
.gr-stu-item:hover { background:var(--a10); }
.gr-stu-item.on { background:var(--a20); border-left:3px solid var(--a); }
.gr-stu-name { font-size:12px; font-weight:700; color:var(--tx); white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
.gr-stu-nick { font-size:10px; color:var(--tx3); margin-top:2px; }
.gr-stu-dot  { width:6px; height:6px; border-radius:50%; margin:3px auto 0; }

/* 우측 콘텐츠 영역 */
.gr-content {
  flex:1; overflow:auto; -webkit-overflow-scrolling:touch;
  padding-bottom:80px;
}
.gr-content::-webkit-scrollbar { width:4px; height:4px; }
.gr-content::-webkit-scrollbar-thumb { background:var(--bdr2); border-radius:2px; }

/* ── 빈 화면 ── */
.gr-empty {
  display:flex; flex-direction:column; align-items:center; justify-content:center;
  height:100%; padding:40px 24px; text-align:center; color:var(--tx3); font-size:14px; line-height:2.4;
}
.gr-empty-ico { font-size:44px; margin-bottom:10px; }

/* ════════════════════════════════════
   엑셀 모드: 가로 스프레드시트
   ════════════════════════════════════ */
.gr-sheet-wrap { min-width:max-content; }
.gr-sheet { border-collapse:collapse; font-size:12px; width:100%; }

/* 학생 고정 컬럼 */
.gr-sheet .gs-fix {
  position:sticky; left:0; z-index:3; background:var(--surf);
  border:1px solid var(--bdr); padding:7px 10px; min-width:90px;
  font-size:12px; font-weight:700; color:var(--tx2); white-space:nowrap;
}
.gr-sheet thead .gs-fix { z-index:4; background:var(--surf2); }
.gr-sheet .gs-fix.on { background:var(--a10); color:var(--a); cursor:pointer; }
.gr-sheet .gs-fix:hover { background:var(--a10); cursor:pointer; }

/* 헤더 셀 */
.gs-th {
  background:var(--surf2); border:1px solid var(--bdr); padding:6px 8px;
  font-size:10px; font-weight:800; color:var(--tx3); text-align:center;
  white-space:nowrap; letter-spacing:.3px;
}
.gs-th.sec { /* 섹션 헤더: 단어/리딩 */
  background:var(--a10); color:var(--a); font-size:11px; letter-spacing:.5px;
}
.gs-th.sub { font-weight:700; color:var(--tx3); font-size:10px; }
.gs-th.reading-hdr { background:rgba(139,92,246,.08); color:#8b5cf6; }
.gs-th.comment-hdr { background:rgba(5,150,105,.08); color:var(--green); }

/* 데이터 셀 */
.gs-td { border:1px solid var(--bdr); text-align:center; padding:0; vertical-align:middle; }
.gs-td.calc { padding:6px 8px; font-size:13px; font-weight:700; }
.gs-td.pass-cell { color:#16a34a; }
.gs-td.fail-cell { color:#f97316; }
.gs-td.score-cell { color:var(--a); }
.gs-td.achv-cell { color:#8b5cf6; font-weight:800; }
.gs-td.comment-cell { padding:0; min-width:140px; }

/* 입력 셀 */
.gs-inp {
  width:100%; min-width:52px; padding:7px 6px; border:none; outline:none;
  background:transparent; font-size:13px; font-weight:700; color:var(--a);
  text-align:center; font-family:var(--font); -moz-appearance:textfield;
  cursor:text;
}
.gs-inp::-webkit-outer-spin-button,
.gs-inp::-webkit-inner-spin-button { -webkit-appearance:none; }
.gs-inp:focus { background:var(--a10); }
.gs-comment-inp {
  width:100%; min-width:140px; padding:7px 8px; border:none; outline:none;
  background:transparent; font-size:12px; color:var(--tx); font-family:var(--font);
  resize:none; height:36px; cursor:text; line-height:1.5;
}
.gs-comment-inp:focus { background:rgba(5,150,105,.05); }

/* 선택 행 하이라이트 */
.gr-sheet tbody tr.selected-row td { background:var(--a10) !important; }
.gr-sheet tbody tr.selected-row td.gs-fix { background:var(--a20) !important; }

/* 저장 버튼 행 */
.gs-save-btn {
  padding:5px 10px; border-radius:7px; background:var(--a); color:#fff;
  border:none; font-size:11px; font-weight:700; cursor:pointer; font-family:var(--font);
  white-space:nowrap; transition:all .15s;
}
.gs-save-btn:active { transform:scale(.94); }
.gs-save-btn.saving { background:var(--bdr2); cursor:default; }

/* ════════════════════════════════════
   카드 모드: 선택 학생 독립 입력
   ════════════════════════════════════ */
.gr-card-view { padding:16px 16px 80px; min-width:280px; }

.gr-student-hero {
  display:flex; align-items:center; gap:12px; margin-bottom:16px;
  background:linear-gradient(135deg,var(--a10),rgba(5,150,105,.06));
  border:1px solid var(--a40); border-radius:14px; padding:14px 16px;
}
.gr-hero-av {
  width:48px; height:48px; border-radius:50%;
  background:linear-gradient(135deg,var(--a20),rgba(5,150,105,.15));
  display:flex; align-items:center; justify-content:center;
  font-size:20px; font-weight:900; color:var(--a); flex-shrink:0;
}
.gr-hero-name  { font-size:18px; font-weight:900; color:var(--tx); }
.gr-hero-meta  { font-size:12px; color:var(--tx3); margin-top:2px; }
.gr-hero-score { margin-left:auto; text-align:right; }
.gr-hero-pct   { font-size:22px; font-weight:900; color:var(--a); line-height:1; }
.gr-hero-lbl   { font-size:10px; color:var(--tx3); }

/* 카드 섹션 */
.gr-section {
  background:var(--card); border:1px solid var(--bdr);
  border-radius:14px; overflow:hidden; box-shadow:var(--sh);
  margin-bottom:14px; animation:cardIn .2s ease both;
}
.gr-sec-head {
  display:flex; align-items:center; justify-content:space-between;
  padding:11px 14px; background:var(--surf2); border-bottom:1px solid var(--bdr);
}
.gr-sec-title { font-size:14px; font-weight:800; color:var(--tx); display:flex; align-items:center; gap:6px; }
.gr-sec-sub   { font-size:11px; color:var(--tx3); }
.gr-sec-badge {
  padding:3px 10px; border-radius:20px; font-size:11px; font-weight:700;
  background:var(--a10); color:var(--a); border:1px solid var(--a40);
}
.gr-sec-badge.rd { background:rgba(139,92,246,.1); color:#8b5cf6; border-color:rgba(139,92,246,.3); }

/* 카드 폼 그리드 */
.gr-form-grid {
  display:grid; gap:0; /* border-collapse 효과 */
}
.gr-form-row {
  display:grid; border-bottom:1px solid var(--bdr);
  align-items:center;
}
.gr-form-row:last-child { border-bottom:none; }
.gr-form-lbl {
  padding:10px 14px; font-size:12px; font-weight:700; color:var(--tx2);
  background:var(--surf2); border-right:1px solid var(--bdr);
}
.gr-form-val { padding:6px 10px; display:flex; align-items:center; gap:6px; }
.gr-form-inp {
  flex:1; padding:8px 10px; border-radius:9px; background:var(--surf2);
  border:1.5px solid var(--bdr); font-size:15px; font-weight:700;
  color:var(--a); text-align:center; outline:none; font-family:var(--font);
  -moz-appearance:textfield; transition:border-color .2s;
}
.gr-form-inp::-webkit-outer-spin-button,
.gr-form-inp::-webkit-inner-spin-button { -webkit-appearance:none; }
.gr-form-inp:focus { border-color:var(--a); background:var(--a10); }
.gr-form-calc {
  min-width:52px; text-align:center; padding:8px 10px;
  font-size:15px; font-weight:800; border-radius:9px;
  background:var(--card2); border:1px solid var(--bdr);
}
.gr-form-calc.pass  { color:#16a34a; background:rgba(22,163,74,.08); border-color:rgba(22,163,74,.25); }
.gr-form-calc.fail  { color:#f97316; background:rgba(249,115,22,.08); border-color:rgba(249,115,22,.25); }
.gr-form-calc.score { color:var(--a); background:var(--a10); border-color:var(--a40); }
.gr-form-calc.achv  { color:#8b5cf6; background:rgba(139,92,246,.1); border-color:rgba(139,92,246,.3); }

/* 카드 저장 버튼 */
.gr-card-save {
  width:100%; padding:13px; border:none; background:var(--a); color:#fff;
  font-size:14px; font-weight:800; cursor:pointer; font-family:var(--font);
  transition:all .15s; margin-top:2px;
}
.gr-card-save:active { opacity:.85; }

/* Teacher's Comment */
.gr-comment-ta {
  width:100%; box-sizing:border-box; padding:12px 14px;
  border:none; outline:none; background:transparent;
  font-size:13px; color:var(--tx); font-family:var(--font);
  resize:none; min-height:80px; line-height:1.8;
}
.gr-comment-ta:focus { background:rgba(5,150,105,.04); }

/* 이력 */
.gr-hist { padding:10px 14px; border-top:1px solid var(--bdr); }
.gr-hist-title { font-size:10px; font-weight:800; color:var(--tx3); letter-spacing:1px; margin-bottom:6px; }
.gr-hist-row {
  display:flex; align-items:center; gap:8px; padding:5px 0;
  border-bottom:1px solid var(--bdr); font-size:12px;
}
.gr-hist-row:last-child { border-bottom:none; }
.gr-hist-date { color:var(--tx3); font-weight:700; min-width:68px; flex-shrink:0; }
.gr-hist-body { flex:1; color:var(--tx2); font-size:11px; }

/* 성적표 공유 */
.gr-share-box { background:var(--surf2); border-radius:10px; padding:12px 14px; font-size:12px; line-height:1.9; color:var(--tx); white-space:pre-wrap; word-break:break-all; border:1px solid var(--bdr); max-height:300px; overflow-y:auto; font-family:var(--font); margin:8px 0; }
.gr-sacts { display:flex; gap:8px; flex-wrap:wrap; }
.gr-sbtn { flex:1; min-width:80px; padding:11px 8px; border-radius:10px; border:none; font-size:13px; font-weight:700; cursor:pointer; font-family:var(--font); transition:all .15s; }
.gr-sbtn.copy  { background:var(--a10); color:var(--a); border:1px solid var(--a40); }
.gr-sbtn.share { background:var(--a); color:#fff; box-shadow:0 3px 10px var(--a40); }
.gr-sbtn:active { transform:scale(.96); }
`;
    document.head.appendChild(s);
  }

  /* ══ INIT ══ */
  async function init() {
    _css();
    if (typeof GradeDB === 'undefined') { console.warn('[GradeApp] GradeDB not loaded'); return; }
    await GradeDB.init();
    console.log('[GradeApp] ✅ v3');
  }

  /* ══ RENDER ══ */
  function render() {
    const pg = document.getElementById('page-grade'); if (!pg) return;
    pg.innerHTML = _shell();
    _fillClass();
    if (_st.classId) _fillBooks();
    _renderStudents();
    _renderContent();
  }

  function _shell() {
    return `
      <div class="ph">
        <div class="phl">
          <div style="width:36px;height:36px;border-radius:10px;background:linear-gradient(135deg,#f59e0b,#d97706);display:flex;align-items:center;justify-content:center;font-size:18px;flex-shrink:0;box-shadow:0 3px 10px rgba(245,158,11,.4)">📝</div>
          <div style="min-width:0">
            <div class="ph-title">성적 관리 <span class="admin-badge">🔑 관리자</span></div>
            <div class="ph-sub" id="gr-sub">반 · 교재를 선택하세요</div>
          </div>
        </div>
        <div class="phr">
          <button class="ibtn" id="gr-rpt-btn" title="전체 성적표"
                  onclick="GradeApp.openReport()" style="display:none">📋</button>
        </div>
      </div>

      <!-- 툴바 -->
      <div class="gr-toolbar">
        <select class="gr-sel" id="gr-csel" onchange="GradeApp._onCls(this.value)">
          <option value="">— 반 선택 —</option>
        </select>
        <select class="gr-sel" id="gr-bsel" onchange="GradeApp._onBk(this.value)" disabled>
          <option value="">— 교재 선택 —</option>
        </select>
        <div class="gr-view-toggle">
          <button class="gr-vbtn ${_st.viewMode==='excel'?'on':''}"
                  onclick="GradeApp._setView('excel')" title="엑셀 모드">🔲 엑셀</button>
          <button class="gr-vbtn ${_st.viewMode==='card'?'on':''}"
                  onclick="GradeApp._setView('card')" title="카드 모드">👤 카드</button>
        </div>
      </div>

      <!-- 본문 2컬럼 -->
      <div class="gr-main">
        <div class="gr-stu-panel" id="gr-stu-panel"></div>
        <div class="gr-content" id="gr-content"></div>
      </div>

      <!-- 성적표 모달 -->
      <div id="gr-rpt-ov" class="ov hidden"
           onclick="if(event.target.id==='gr-rpt-ov')GradeApp.closeReport()">
        <div class="sh" id="gr-rpt-sh" onclick="event.stopPropagation()"
             style="max-height:92vh;display:flex;flex-direction:column;"></div>
      </div>`;
  }

  /* ── 셀렉트 채우기 ── */
  function _fillClass() {
    const sel = document.getElementById('gr-csel'); if (!sel) return;
    const cls = typeof DB !== 'undefined' ? DB.getActiveClasses() : [];
    sel.innerHTML = `<option value="">— 반 선택 —</option>` +
      cls.map(c => `<option value="${c.id}" ${_st.classId===c.id?'selected':''}>${_e(c.name)}</option>`).join('');
  }

  function _fillBooks() {
    const sel = document.getElementById('gr-bsel'); if (!sel) return;
    if (!_st.classId) { sel.innerHTML = `<option value="">— 교재 선택 —</option>`; sel.disabled = true; return; }
    const books = typeof BookLibDB !== 'undefined' ? BookLibDB.getBooksForClass(_st.classId) : [];
    sel.disabled = false;
    sel.innerHTML = `<option value="">— 교재 선택 —</option>` +
      books.map(b => `<option value="${b.id}" ${_st.bookId===b.id?'selected':''}>${_e(b.name)}</option>`).join('');
    if (!books.length) sel.innerHTML = `<option value="">이 반에 배정된 교재 없음</option>`;
  }

  /* ── 학생 패널 ── */
  function _renderStudents() {
    const panel = document.getElementById('gr-stu-panel'); if (!panel) return;
    const students = _getStudents();
    if (!students.length) { panel.innerHTML = ''; return; }

    panel.innerHTML = students.map(s => {
      const rec = _st.classId && _st.bookId
        ? GradeDB.getLatest(_st.classId, s.id, _st.bookId) : null;
      const config = _st.bookId ? GradeDB.getReportConfig(_st.bookId) : null;
      const hasData = !!rec;
      const ach = rec?.word?.totalQ > 0 && rec?.word?.pass != null
        ? Math.round(rec.word.pass / rec.word.totalQ * 100) : null;
      const dotColor = hasData ? (ach != null && ach >= 80 ? '#16a34a' : ach != null ? '#f97316' : '#a78bfa') : 'var(--bdr2)';

      return `
        <div class="gr-stu-item ${_st.studentId===s.id?'on':''}"
             onclick="GradeApp._onStu('${s.id}')">
          <div class="gr-stu-name">${_e(s.name)}</div>
          ${s.nickname ? `<div class="gr-stu-nick">${_e(s.nickname)}</div>` : ''}
          <div class="gr-stu-dot" style="background:${dotColor};opacity:${hasData?1:.3}"></div>
        </div>`;
    }).join('');
  }

  /* ── 우측 컨텐츠 ── */
  function _renderContent() {
    const cnt = document.getElementById('gr-content'); if (!cnt) return;
    if (!_st.classId || !_st.bookId) {
      cnt.innerHTML = `<div class="gr-empty"><div class="gr-empty-ico">📝</div>반과 교재를 선택하세요</div>`;
      return;
    }
    const students = _getStudents();
    if (!students.length) {
      cnt.innerHTML = `<div class="gr-empty"><div class="gr-empty-ico">👨‍🎓</div>재원 학생이 없습니다<br><small>학생 탭에서 엑셀을 가져오세요</small></div>`;
      return;
    }
    if (_st.viewMode === 'excel') _renderExcel(cnt, students);
    else                          _renderCard(cnt, students);
  }

  /* ════════════════════════════════════
   * 엑셀 모드
   * ════════════════════════════════════ */
  function _renderExcel(cnt, students) {
    const config  = GradeDB.getReportConfig(_st.bookId);
    const totalWQ = config.word?.totalQ || 0;
    const actRevs = GradeDB.getActiveReviews(_st.bookId);
    const hasRd   = config.reading?.enabled && actRevs.length > 0;
    const totalRQ = config.reading?.totalQ || 0;
    const rvN     = actRevs.length;

    /* 데이터 확보 */
    students.forEach(s => {
      if (!_st.data[s.id]) {
        const rec = GradeDB.getLatest(_st.classId, s.id, _st.bookId);
        _st.data[s.id] = rec
          ? JSON.parse(JSON.stringify(rec))
          : { word: { totalQ:totalWQ, retake:'', pass:'' }, reading: {}, comment:'' };
        _st.data[s.id].comment = _st.data[s.id].comment || '';
      }
    });

    /* 헤더 행 계산 */
    // 단어: 총문제수 | 재시험 | 통과 | 성취율 = 4컬럼
    // 리딩: 총문제수 | R×N(정답수) | R×N(점수) | 성취율 = 1+N+N+1
    const rdCols = hasRd ? 1 + rvN * 2 + 1 : 0;
    const totalCols = 4 + rdCols + 1; // +1 comment

    const rvHeaderCells = actRevs.map(rv => `<th class="gs-th sub">${_e(rv.name)}</th>`).join('');
    const rvHeaderCells2= actRevs.map(rv => `<th class="gs-th sub">${_e(rv.name)}</th>`).join('');

    const rdHeaderRow1 = hasRd ? `
      <th class="gs-th reading-hdr" colspan="${rdCols}">📖 리딩 평가</th>` : '';
    const rdHeaderRow2 = hasRd ? `
      <th class="gs-th sub reading-hdr">총문제</th>
      <th class="gs-th reading-hdr" colspan="${rvN}">정답 수</th>
      <th class="gs-th reading-hdr" colspan="${rvN}">점수</th>
      <th class="gs-th sub reading-hdr">성취율</th>` : '';

    const html = `
      <div class="gr-sheet-wrap">
        <table class="gr-sheet">
          <thead>
            <tr>
              <th class="gs-fix gs-th" rowspan="2">학생</th>
              <th class="gs-th sec" colspan="4">🔤 단어 평가</th>
              ${rdHeaderRow1}
              <th class="gs-th comment-hdr" rowspan="2">💬 Teacher's Comment</th>
            </tr>
            <tr>
              <th class="gs-th sub">총문제</th>
              <th class="gs-th sub">재시험</th>
              <th class="gs-th sub">통과</th>
              <th class="gs-th sub">성취율</th>
              ${rdHeaderRow2}
            </tr>
          </thead>
          <tbody>
            ${students.map(s => _excelRow(s, config, totalWQ, actRevs, totalRQ, hasRd)).join('')}
          </tbody>
        </table>
      </div>`;
    cnt.innerHTML = html;
  }

  function _excelRow(s, config, totalWQ, actRevs, totalRQ, hasRd) {
    const d  = _st.data[s.id] || {};
    const wd = d.word || {};
    const rd = d.reading || {};
    const isSelected = _st.studentId === s.id;

    const retake  = wd.retake !== undefined && wd.retake !== '' ? Number(wd.retake) : '';
    const pass    = retake !== '' ? Math.max(0, totalWQ - retake) : '';
    const achW    = pass !== '' && totalWQ > 0 ? Math.round(pass / totalWQ * 100) : '';
    const isGoodW = achW !== '' && achW >= 80;

    /* 리딩 셀 */
    const rdCells = hasRd ? `
      <td class="gs-td calc score-cell">${totalRQ||'—'}</td>
      ${actRevs.map((rv, i) => {
        const key = `R${i}`;
        const v = rd[key]?.correct !== undefined && rd[key]?.correct !== '' ? rd[key].correct : '';
        return `<td class="gs-td" style="min-width:52px">
          <input class="gs-inp" type="number" min="0" max="${totalRQ}" step="1"
                 value="${v}" placeholder="—"
                 oninput="GradeApp._excelRdInput('${s.id}','${key}',this.value)">
        </td>`;
      }).join('')}
      ${actRevs.map((rv, i) => {
        const key = `R${i}`;
        const sc = rd[key]?.score !== undefined && rd[key]?.score !== '' ? rd[key].score : '';
        return `<td class="gs-td calc score-cell" id="gr-sc-${s.id}-${key}">${sc!==''?sc:'—'}</td>`;
      }).join('')}
      <td class="gs-td calc achv-cell" id="gr-achvrd-${s.id}">
        ${_calcRdAchieve(rd, actRevs, config)||'—'}
      </td>` : '';

    return `
      <tr class="${isSelected?'selected-row':''}" id="gr-row-${s.id}">
        <td class="gs-fix ${isSelected?'on':''}"
            onclick="GradeApp._onStu('${s.id}')">
          <div style="font-weight:700;font-size:12px">${_e(s.name)}</div>
          ${s.nickname?`<div style="font-size:10px;color:var(--tx3)">${_e(s.nickname)}</div>`:''}
        </td>
        <td class="gs-td calc score-cell">${totalWQ||'—'}</td>
        <td class="gs-td" style="min-width:56px">
          <input class="gs-inp" type="number" min="0" max="${totalWQ}" step="1"
                 value="${retake}" placeholder="0"
                 id="gr-retake-${s.id}"
                 oninput="GradeApp._excelWordInput('${s.id}',this.value)">
        </td>
        <td class="gs-td calc ${pass!==''?(isGoodW?'pass-cell':'fail-cell'):''}" id="gr-pass-${s.id}">
          ${pass!==''?pass:'—'}
        </td>
        <td class="gs-td calc ${achW!==''?(isGoodW?'pass-cell':'fail-cell'):''}" id="gr-achvw-${s.id}">
          ${achW!==''?achW+'%':'—'}
        </td>
        ${rdCells}
        <td class="gs-td comment-cell">
          <textarea class="gs-comment-inp" rows="1" placeholder="선생님 코멘트…"
                    id="gr-cmt-${s.id}"
                    oninput="GradeApp._excelComment('${s.id}',this.value)"
                    >${_e(d.comment||'')}</textarea>
        </td>
        <td class="gs-td" style="padding:4px 6px;white-space:nowrap;border:1px solid var(--bdr)">
          <button class="gs-save-btn" id="gr-sbtn-${s.id}"
                  onclick="GradeApp.saveOne('${s.id}')">저장</button>
        </td>
      </tr>`;
  }

  /* 엑셀 모드 입력 핸들러 */
  function _excelWordInput(sid, val) {
    _ensureData(sid);
    const config = GradeDB.getReportConfig(_st.bookId);
    const totalQ = config.word?.totalQ || 0;
    const retake = val === '' ? '' : Math.max(0, Math.min(totalQ, Number(val)));
    const pass   = retake !== '' ? Math.max(0, totalQ - retake) : '';
    const achW   = pass !== '' && totalQ > 0 ? Math.round(pass / totalQ * 100) : '';
    _st.data[sid].word = { totalQ, retake, pass };

    const passEl = document.getElementById(`gr-pass-${sid}`);
    const achEl  = document.getElementById(`gr-achvw-${sid}`);
    const isGood = achW !== '' && achW >= 80;
    if (passEl) { passEl.textContent = pass !== '' ? pass : '—'; passEl.className = `gs-td calc ${pass!==''?(isGood?'pass-cell':'fail-cell'):''}`; }
    if (achEl)  { achEl.textContent  = achW!==''?achW+'%':'—';  achEl.className  = `gs-td calc ${achW!==''?(isGood?'pass-cell':'fail-cell'):''}`; }
  }

  function _excelRdInput(sid, key, val) {
    _ensureData(sid);
    const config = GradeDB.getReportConfig(_st.bookId);
    const totalQ = config.reading?.totalQ || 0;
    const actRevs= GradeDB.getActiveReviews(_st.bookId);
    const correct = val === '' ? '' : Math.max(0, Math.min(totalQ, Number(val)));
    const score   = correct !== '' ? Math.round(correct / totalQ * 100 * 10) / 10 : '';
    if (!_st.data[sid].reading) _st.data[sid].reading = {};
    _st.data[sid].reading[key] = { correct, score };

    const scEl   = document.getElementById(`gr-sc-${sid}-${key}`);
    const achvEl = document.getElementById(`gr-achvrd-${sid}`);
    if (scEl)   scEl.textContent = score !== '' ? score : '—';
    if (achvEl) achvEl.textContent = _calcRdAchieve(_st.data[sid].reading, actRevs, config) || '—';
  }

  function _excelComment(sid, val) {
    _ensureData(sid);
    _st.data[sid].comment = val;
  }

  function _calcRdAchieve(rd, actRevs, config) {
    if (!actRevs.length) return null;
    const scores = actRevs.map((_, i) => rd[`R${i}`]?.score).filter(s => s !== '' && s != null);
    if (!scores.length) return null;
    return Math.round(scores.reduce((a,b)=>a+b,0) / scores.length * 10) / 10 + '%';
  }

  /* ════════════════════════════════════
   * 카드 모드
   * ════════════════════════════════════ */
  function _renderCard(cnt, students) {
    const s = students.find(s => s.id === _st.studentId) || students[0];
    if (!s) { cnt.innerHTML = `<div class="gr-empty"><div class="gr-empty-ico">👆</div>좌측에서 학생을 선택하세요</div>`; return; }
    // 자동 선택
    if (!_st.studentId) { _st.studentId = s.id; _renderStudents(); }

    const config  = GradeDB.getReportConfig(_st.bookId);
    const totalWQ = config.word?.totalQ || 0;
    const actRevs = GradeDB.getActiveReviews(_st.bookId);
    const hasRd   = config.reading?.enabled && actRevs.length > 0;
    const totalRQ = config.reading?.totalQ || 0;

    _ensureData(s.id);
    const d = _st.data[s.id];
    const wd= d.word || {};
    const rd= d.reading || {};

    const retake  = wd.retake !== undefined && wd.retake !== '' ? wd.retake : '';
    const pass    = retake !== '' ? Math.max(0, totalWQ - Number(retake)) : '';
    const achW    = pass !== '' && totalWQ > 0 ? Math.round(pass / totalWQ * 100) : '';
    const isGoodW = achW !== '' && achW >= 80;

    const history = GradeDB.getRecords(_st.classId, s.id, _st.bookId);
    const achRd   = _calcRdAchieve(rd, actRevs, config);

    cnt.innerHTML = `
      <div class="gr-card-view">
        <!-- 학생 히어로 -->
        <div class="gr-student-hero">
          <div class="gr-hero-av">${_e((s.name||'?')[0])}</div>
          <div>
            <div class="gr-hero-name">${_e(s.name)}</div>
            <div class="gr-hero-meta">${s.nickname?`(${_e(s.nickname)}) · `:''}${_getCls(_st.classId)?.name||''}반</div>
          </div>
          ${achW!==''?`<div class="gr-hero-score"><div class="gr-hero-pct" style="color:${isGoodW?'#16a34a':'#f97316'}">${achW}%</div><div class="gr-hero-lbl">단어 성취율</div></div>`:''}
        </div>

        <!-- 1. 단어 평가 -->
        <div class="gr-section">
          <div class="gr-sec-head">
            <div>
              <div class="gr-sec-title">🔤 단어 평가</div>
              <div class="gr-sec-sub">총 문제 수 ${totalWQ}문제</div>
            </div>
            ${achW!==''?`<span class="gr-sec-badge" style="color:${isGoodW?'#16a34a':'#f97316'};background:${isGoodW?'rgba(22,163,74,.08)':'rgba(249,115,22,.08)'};border-color:${isGoodW?'rgba(22,163,74,.25)':'rgba(249,115,22,.25)'}">${achW}%</span>`:''}
          </div>
          <div class="gr-form-grid">
            <div class="gr-form-row" style="grid-template-columns:110px 1fr 1fr 1fr">
              <div class="gr-form-lbl">총 문제 수</div>
              <div class="gr-form-val"><div class="gr-form-calc score">${totalWQ||'—'}</div></div>
              <div class="gr-form-lbl" style="border-left:1px solid var(--bdr)">재시험</div>
              <div class="gr-form-val">
                <input class="gr-form-inp" type="number" min="0" max="${totalWQ}" step="1"
                       id="gr-cd-retake" placeholder="0" value="${retake}"
                       oninput="GradeApp._cardWordInput(this.value)">
              </div>
            </div>
            <div class="gr-form-row" style="grid-template-columns:110px 1fr 110px 1fr">
              <div class="gr-form-lbl">통과</div>
              <div class="gr-form-val">
                <div class="gr-form-calc ${pass!==''?(isGoodW?'pass':'fail'):''}" id="gr-cd-pass">${pass!==''?pass:'—'}</div>
              </div>
              <div class="gr-form-lbl" style="border-left:1px solid var(--bdr)">성취율</div>
              <div class="gr-form-val">
                <div class="gr-form-calc ${achW!==''?(isGoodW?'pass':'fail'):''}" id="gr-cd-achw">${achW!==''?achW+'%':'—'}</div>
              </div>
            </div>
          </div>
        </div>

        <!-- 2. 리딩 평가 -->
        ${hasRd ? `
        <div class="gr-section">
          <div class="gr-sec-head">
            <div>
              <div class="gr-sec-title">📖 리딩 평가</div>
              <div class="gr-sec-sub">총 문제 수 ${totalRQ}문제 · ${actRevs.map(r=>r.name).join(', ')}</div>
            </div>
            ${achRd?`<span class="gr-sec-badge rd">${achRd}</span>`:''}
          </div>
          <div class="gr-form-grid">
            ${actRevs.map((rv, i) => {
              const key = `R${i}`;
              const correct = rd[key]?.correct !== undefined && rd[key]?.correct !== '' ? rd[key].correct : '';
              const score   = rd[key]?.score !== '' && rd[key]?.score != null ? rd[key].score : '';
              return `
                <div class="gr-form-row" style="grid-template-columns:110px 1fr 110px 1fr">
                  <div class="gr-form-lbl">${_e(rv.name)}<br><span style="font-size:10px;color:var(--tx3)">정답 수</span></div>
                  <div class="gr-form-val">
                    <input class="gr-form-inp" type="number" min="0" max="${totalRQ}" step="1"
                           id="gr-cd-rd-${key}" placeholder="0" value="${correct}"
                           oninput="GradeApp._cardRdInput('${key}',this.value)">
                    <span style="font-size:11px;color:var(--tx3)">/ ${totalRQ}</span>
                  </div>
                  <div class="gr-form-lbl" style="border-left:1px solid var(--bdr)">점수</div>
                  <div class="gr-form-val">
                    <div class="gr-form-calc score" id="gr-cd-sc-${key}">${score!==''?score:'—'}</div>
                  </div>
                </div>`;
            }).join('')}
            <div class="gr-form-row" style="grid-template-columns:110px 1fr">
              <div class="gr-form-lbl">성취율<br><span style="font-size:10px;color:var(--tx3)">(점수 평균)</span></div>
              <div class="gr-form-val">
                <div class="gr-form-calc achv" id="gr-cd-achrd">${achRd||'—'}</div>
              </div>
            </div>
          </div>
        </div>` : ''}

        <!-- 3. Teacher's Comment -->
        <div class="gr-section">
          <div class="gr-sec-head">
            <div class="gr-sec-title">💬 Teacher's Comment</div>
          </div>
          <textarea class="gr-comment-ta" id="gr-cd-comment"
                    placeholder="학생의 성적 결과에 대한 코멘트를 입력하세요…"
                    oninput="GradeApp._cardComment(this.value)">${_e(d.comment||'')}</textarea>
          <button class="gr-card-save" onclick="GradeApp.saveOne('${s.id}')">💾 성적 저장</button>
        </div>

        <!-- 이력 -->
        ${history.length ? `
        <div class="gr-section">
          <div class="gr-sec-head"><div class="gr-sec-title">📋 성적 이력</div></div>
          <div class="gr-hist">
            ${history.map(r => {
              const cfg = GradeDB.getReportConfig(_st.bookId);
              const achRdH = r.reading ? _calcRdAchieve(r.reading, actRevs, cfg) : null;
              return `<div class="gr-hist-row">
                <span class="gr-hist-date">${r.date||''}</span>
                <span class="gr-hist-body">${r.word?.pass!=null?`단어 ${r.word.pass}/${r.word.totalQ}`:''}${achRdH?` · 리딩 ${achRdH}`:''}</span>
                <button class="ibtn red" style="width:26px;height:26px;font-size:10px;flex-shrink:0"
                        onclick="GradeApp.deleteRecord('${r.id}','${s.id}')">🗑</button>
              </div>`;
            }).join('')}
          </div>
        </div>` : ''}
      </div>`;
  }

  /* 카드 모드 입력 핸들러 */
  function _cardWordInput(val) {
    const sid = _st.studentId; if (!sid) return;
    _ensureData(sid);
    const config = GradeDB.getReportConfig(_st.bookId);
    const totalQ = config.word?.totalQ || 0;
    const retake = val === '' ? '' : Math.max(0, Math.min(totalQ, Number(val)));
    const pass   = retake !== '' ? Math.max(0, totalQ - retake) : '';
    const achW   = pass !== '' && totalQ > 0 ? Math.round(pass / totalQ * 100) : '';
    const isGood = achW !== '' && achW >= 80;
    _st.data[sid].word = { totalQ, retake, pass };

    const passEl = document.getElementById('gr-cd-pass');
    const achEl  = document.getElementById('gr-cd-achw');
    if (passEl) { passEl.textContent = pass !== '' ? pass : '—'; passEl.className = `gr-form-calc ${pass!==''?(isGood?'pass':'fail'):''}`; }
    if (achEl)  { achEl.textContent  = achW!==''?achW+'%':'—'; achEl.className  = `gr-form-calc ${achW!==''?(isGood?'pass':'fail'):''}`; }
  }

  function _cardRdInput(key, val) {
    const sid = _st.studentId; if (!sid) return;
    _ensureData(sid);
    const config  = GradeDB.getReportConfig(_st.bookId);
    const totalQ  = config.reading?.totalQ || 0;
    const actRevs = GradeDB.getActiveReviews(_st.bookId);
    const correct = val === '' ? '' : Math.max(0, Math.min(totalQ, Number(val)));
    const score   = correct !== '' ? Math.round(correct / totalQ * 100 * 10) / 10 : '';
    if (!_st.data[sid].reading) _st.data[sid].reading = {};
    _st.data[sid].reading[key] = { correct, score };

    const scEl   = document.getElementById(`gr-cd-sc-${key}`);
    const achvEl = document.getElementById('gr-cd-achrd');
    if (scEl)   scEl.textContent = score !== '' ? score : '—';
    if (achvEl) achvEl.textContent = _calcRdAchieve(_st.data[sid].reading, actRevs, config) || '—';
  }

  function _cardComment(val) {
    const sid = _st.studentId; if (!sid) return;
    _ensureData(sid);
    _st.data[sid].comment = val;
  }

  /* ══ 저장 ══ */
  async function saveOne(sid) {
    if (!_st.classId || !_st.bookId) { _toast('⚠️ 반과 교재를 선택해주세요'); return; }
    if (_st.saving.has(sid)) return;
    _st.saving.add(sid);
    const btn = document.getElementById(`gr-sbtn-${sid}`);
    if (btn) { btn.textContent = '저장 중…'; btn.className = 'gs-save-btn saving'; }

    /* 엑셀 모드: DOM에서 최신값 수집 */
    if (_st.viewMode === 'excel') {
      const config  = GradeDB.getReportConfig(_st.bookId);
      const totalWQ = config.word?.totalQ || 0;
      const actRevs = GradeDB.getActiveReviews(_st.bookId);
      const totalRQ = config.reading?.totalQ || 0;

      const retakeEl = document.getElementById(`gr-retake-${sid}`);
      if (retakeEl) {
        const retake = retakeEl.value === '' ? 0 : Number(retakeEl.value);
        _st.data[sid] = _st.data[sid] || {};
        _st.data[sid].word = { totalQ:totalWQ, retake, pass:Math.max(0,totalWQ-retake) };
      }
      const cmtEl = document.getElementById(`gr-cmt-${sid}`);
      if (cmtEl) { _st.data[sid] = _st.data[sid] || {}; _st.data[sid].comment = cmtEl.value; }
    }

    const d = _st.data[sid] || {};
    await GradeDB.saveRecord({
      classId:   _st.classId,
      studentId: sid,
      bookId:    _st.bookId,
      word:      d.word    || null,
      reading:   d.reading || null,
      comment:   d.comment || '',
    });

    _st.saving.delete(sid);
    if (btn) { btn.textContent = '✅ 저장됨'; btn.className = 'gs-save-btn'; setTimeout(()=>{ if(btn)btn.textContent='저장'; }, 2000); }
    _toast('✅ 저장 완료', 'success');
    _renderStudents(); // 닷 색상 업데이트
  }

  async function deleteRecord(recordId, sid) {
    if (!confirm('이 성적 기록을 삭제하시겠습니까?')) return;
    await GradeDB.deleteRecord(_st.classId, sid||_st.studentId, _st.bookId, recordId);
    // 캐시 무효화 후 재렌더
    if (_st.data[sid||_st.studentId]) delete _st.data[sid||_st.studentId];
    _toast('🗑 삭제 완료');
    _renderContent();
    _renderStudents();
  }

  /* ══ 선택 핸들러 ══ */
  function _onCls(clsId) {
    _st.classId = clsId || null;
    _st.bookId = null; _st.studentId = null; _st.data = {};
    _fillBooks();
    _renderStudents();
    _renderContent();
    _updateRptBtn(); _updateSub();
    const bsel = document.getElementById('gr-bsel');
    if (bsel) bsel.disabled = !_st.classId;
  }

  function _onBk(bkId) {
    _st.bookId = bkId || null;
    _st.studentId = null; _st.data = {};
    _renderStudents();
    _renderContent();
    _updateRptBtn(); _updateSub();
  }

  function _onStu(sid) {
    _st.studentId = sid || null;
    _renderStudents();
    if (_st.viewMode === 'card') _renderContent();
    else {
      // 엑셀 모드: 선택 행 하이라이트 + 스크롤
      document.querySelectorAll('.gr-sheet tbody tr').forEach(tr => {
        tr.classList.toggle('selected-row', tr.id === `gr-row-${sid}`);
      });
      document.querySelectorAll('.gs-fix').forEach(td => {
        const match = td.getAttribute('onclick')?.match(/'([^']+)'/);
        td.classList.toggle('on', match?.[1] === sid);
      });
      const row = document.getElementById(`gr-row-${sid}`);
      row?.scrollIntoView({ behavior:'smooth', block:'nearest' });
    }
  }

  function _setView(mode) {
    _st.viewMode = mode;
    _st.data = {}; // 뷰 전환 시 캐시 초기화
    document.querySelectorAll('.gr-vbtn').forEach((b,i)=>b.classList.toggle('on',(i===0&&mode==='excel')||(i===1&&mode==='card')));
    _renderContent();
  }

  function _updateRptBtn() {
    const btn = document.getElementById('gr-rpt-btn');
    if (btn) btn.style.display = (_st.classId && _st.bookId) ? '' : 'none';
  }
  function _updateSub() {
    const sub = document.getElementById('gr-sub'); if (!sub) return;
    const cls = _st.classId ? _getCls(_st.classId) : null;
    const bk  = _st.bookId && typeof BookLibDB !== 'undefined' ? BookLibDB.getBookById(_st.bookId) : null;
    sub.textContent = cls && bk ? `${cls.name}반 · ${bk.name}` : cls ? `${cls.name}반` : '반 · 교재를 선택하세요';
  }

  /* ══ 반 전체 성적표 ══ */
  function openReport() {
    if (!_st.classId || !_st.bookId) { _toast('⚠️ 반과 교재를 선택해주세요'); return; }
    const ov = document.getElementById('gr-rpt-ov');
    const sh = document.getElementById('gr-rpt-sh'); if (!ov||!sh) return;

    const cls     = _getCls(_st.classId);
    const book    = typeof BookLibDB !== 'undefined' ? BookLibDB.getBookById(_st.bookId) : null;
    const config  = GradeDB.getReportConfig(_st.bookId);
    const actRevs = GradeDB.getActiveReviews(_st.bookId);
    const hasRd   = config.reading?.enabled && actRevs.length > 0;
    const students= _getStudents();
    const today   = new Date().toLocaleDateString('ko-KR');

    const rows = students.map(s => ({ s, rec: GradeDB.getLatest(_st.classId, s.id, _st.bookId) }));

    /* 공유 텍스트 */
    let txt = `📝 ${book?.name||''} 성적표\n🏫 ${cls?.name||''}반 · ${today}\n${'─'.repeat(30)}\n`;
    rows.forEach(({s, rec:r}) => {
      if (!r) { txt += `${s.name}: 미입력\n`; return; }
      const achW = r.word?.pass!=null&&r.word?.totalQ>0 ? Math.round(r.word.pass/r.word.totalQ*100) : null;
      const achRd = hasRd ? _calcRdAchieve(r.reading||{}, actRevs, config) : null;
      txt += `${s.name}: 단어 ${r.word?.pass??'—'}/${r.word?.totalQ??'—'}${achW!=null?`(${achW}%)`:''}${achRd?` · 리딩 ${achRd}`:''}`;
      if (r.comment) txt += `\n  💬 ${r.comment}`;
      txt += '\n';
    });

    const rdThs = hasRd ? actRevs.map(rv=>`<th>${_e(rv.name)}<br>정답/점수</th>`).join('')+'<th>성취율</th>' : '';
    const tableRows = rows.map(({s, rec:r}) => {
      if (!r) return `<tr><td style="font-weight:700;padding:6px 10px">${_e(s.name)}</td><td colspan="99" style="color:var(--tx3);font-size:12px;padding:6px 10px">미입력</td></tr>`;
      const achW  = r.word?.pass!=null&&r.word?.totalQ>0 ? Math.round(r.word.pass/r.word.totalQ*100) : null;
      const isGW  = achW!=null&&achW>=80;
      const achRd = hasRd ? _calcRdAchieve(r.reading||{}, actRevs, config) : null;
      const rdTds = hasRd ? actRevs.map((_,i)=>{const k=`R${i}`;const c=r.reading?.[k]?.correct??'—';const sc=r.reading?.[k]?.score??'—';return `<td style="font-size:12px;padding:6px 8px">${c} / ${sc!=='—'?sc:sc}</td>`;}).join('') + `<td style="color:#8b5cf6;font-weight:700;font-size:12px;padding:6px 8px">${achRd||'—'}</td>` : '';
      return `<tr>
        <td style="font-weight:700;padding:6px 10px">${_e(s.name)}</td>
        <td style="color:${isGW?'#16a34a':'#f97316'};font-weight:700;padding:6px 8px">${r.word?.pass??'—'}</td>
        <td style="padding:6px 8px">${r.word?.retake??'—'}</td>
        <td style="color:${isGW?'#16a34a':'#f97316'};font-weight:800;padding:6px 8px">${achW!=null?achW+'%':'—'}</td>
        ${rdTds}
        <td style="font-size:11px;color:var(--tx2);padding:6px 8px;max-width:160px">${_e(r.comment||'')}</td>
      </tr>`;
    }).join('');

    sh.innerHTML = `
      <div class="sh-handle"></div>
      <div class="sh-title">📋 반 전체 성적표</div>
      <div class="sh-sub">${_e(cls?.name||'')}반 · ${_e(book?.name||'')} · ${today}</div>
      <div style="flex:1;overflow:auto;padding:10px 0 8px">
        <div style="overflow-x:auto;padding:0 2px">
          <table style="width:100%;border-collapse:collapse;font-size:13px;min-width:max-content">
            <thead>
              <tr style="background:var(--surf2)">
                <th style="padding:8px 10px;text-align:left;font-size:11px;color:var(--tx3);border:1px solid var(--bdr)">학생</th>
                <th style="padding:8px 10px;font-size:11px;color:var(--tx3);border:1px solid var(--bdr)">통과</th>
                <th style="padding:8px 10px;font-size:11px;color:var(--tx3);border:1px solid var(--bdr)">재시험</th>
                <th style="padding:8px 10px;font-size:11px;color:var(--tx3);border:1px solid var(--bdr)">성취율</th>
                ${rdThs}
                <th style="padding:8px 10px;font-size:11px;color:var(--tx3);border:1px solid var(--bdr)">코멘트</th>
              </tr>
            </thead>
            <tbody style="border:1px solid var(--bdr)">${tableRows}</tbody>
          </table>
        </div>
        <div style="font-size:10px;font-weight:800;color:var(--tx3);letter-spacing:1px;margin:12px 14px 4px">공유용 텍스트</div>
        <div class="gr-share-box" style="margin:0 14px 8px">${_e(txt)}</div>
      </div>
      <div class="gr-sacts" style="padding:0 0 8px">
        <button class="gr-sbtn copy"  onclick="GradeApp._copy(${JSON.stringify(txt)})">📋 복사</button>
        <button class="gr-sbtn share" onclick="GradeApp._share(${JSON.stringify(txt)})">📤 공유</button>
      </div>
      <button class="btn-x" style="width:100%" onclick="GradeApp.closeReport()">닫기</button>`;
    ov.classList.remove('hidden');
    history.pushState({ pg:'grade', modal:'report' }, '');
  }

  async function _copy(text) { try { await navigator.clipboard.writeText(text); _toast('📋 복사됐습니다','success'); } catch { _toast('⚠️ 복사 실패'); } }
  async function _share(text) {
    const sd={title:'성적표',text};
    if(navigator.share&&navigator.canShare?.(sd)){try{await navigator.share(sd);_toast('📤 공유 완료','success');return;}catch(e){if(e.name==='AbortError')return;}}
    _copy(text);
  }
  function closeReport() { document.getElementById('gr-rpt-ov')?.classList.add('hidden'); }

  /* ══ 유틸 ══ */
  function _getStudents() {
    const cls = _st.classId ? _getCls(_st.classId) : null;
    if (!cls || typeof StudentDB === 'undefined') return [];
    return StudentDB.getFiltered({ classCode: cls.name, status: '재원' });
  }

  function _ensureData(sid) {
    if (!_st.data[sid]) {
      const config  = GradeDB.getReportConfig(_st.bookId);
      const totalWQ = config.word?.totalQ || 0;
      const rec = GradeDB.getLatest(_st.classId, sid, _st.bookId);
      _st.data[sid] = rec
        ? JSON.parse(JSON.stringify(rec))
        : { word:{ totalQ:totalWQ, retake:'', pass:'' }, reading:{}, comment:'' };
      _st.data[sid].comment = _st.data[sid].comment || '';
    }
  }

  function _getCls(id) {
    if (typeof DB === 'undefined') return null;
    if (typeof DB.getClassById === 'function') return DB.getClassById(id);
    return (DB.getActiveClasses?.() || []).find(c => c.id === id) || null;
  }

  const _e = v => String(v??'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
  function _toast(msg, type) {
    const el = document.getElementById('toast'); if (!el) return;
    el.textContent = msg; el.className = type==='success'?'success':'';
    el.classList.remove('hidden'); clearTimeout(el._t);
    el._t = setTimeout(() => el.classList.add('hidden'), 3000);
  }

  return {
    init, render,
    _onCls, _onBk, _onStu, _setView,
    _excelWordInput, _excelRdInput, _excelComment,
    _cardWordInput, _cardRdInput, _cardComment,
    saveOne, deleteRecord,
    openReport, closeReport, _copy, _share,
  };
})();
