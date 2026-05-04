/**
 * grade-app.js — v4.1
 *
 * 변경사항
 * ─────────────────────────────────────
 * · 저장 버그 수정 (saving Set 제거, 단순 async/await)
 * · 액션 컬럼 제거 → 우클릭 컨텍스트 메뉴 (저장/초기화)
 * · 일괄저장 → "저장" 버튼으로 통합 (모든 학생 한번에)
 * · 단어순/리딩순 버튼 제거 → 헤더 클릭 정렬 (학생/단어성취율/리딩성취율)
 * · Teacher's Comment 2.5배 확장
 * · 학생 컬럼 정렬 추가, 반/교재 변경 시 정렬 초기화
 * · 평균 행: "반 평균"→"평균", 학생~통과 셀 병합, 실시간 계산
 * · Enter키만 다음 학생 이동, 방향키↑↓ = 입력값 증가/감소
 * · 리딩 헤더: 총문제 2행 병합, 서브컬럼 교재 선택 수 기준
 * · 카드모드: 3:4:3 비율, 여백 없이 fit, 정답수→점수 실시간 동작 수정
 * · 입력/비입력 셀 시각적 구분
 */
const GradeApp = (() => {
  /* ══ 상수 ══ */
  const ANIMALS_M = ['🐯','🦊','🐻','🐼','🦁','🐮','🐸','🐺'];
  const ANIMALS_F = ['🐱','🐰','🐹','🐨','🦋','🦄','🐧','🐔'];
  const CM_W = 340; // Comment column width px

  /* ══ 상태 ══ */
  let _st = {
    classId:   null,
    bookId:    null,
    studentId: null,
    viewMode:  'excel',
    data:      {},
    dirty:     new Set(),
    sortCol:   null,   // null | 'name' | 'wordAch' | 'rdAch'
    sortDesc:  true,
    slideIdx:  0,
    reportLayout: 1,
    reportGraph:  true,
    touchStartX: 0,
  };

  /* ══ CSS ══ */
  function _css() {
    if (document.getElementById('gr-styles')) return;
    const s = document.createElement('style');
    s.id = 'gr-styles';
    s.textContent = `
#page-grade{display:none;flex-direction:column;height:100%;overflow:hidden;}
#page-grade.on{display:flex;}

/* toolbar */
.gr-toolbar{display:flex;align-items:center;gap:7px;padding:8px 12px;background:var(--surf);border-bottom:1.5px solid var(--bdr);flex-shrink:0;flex-wrap:wrap;}
.gr-sel{padding:7px 10px;border-radius:10px;background:var(--surf2);border:1.5px solid var(--bdr);font-size:13px;color:var(--tx);outline:none;cursor:pointer;font-family:var(--font);-webkit-appearance:none;transition:border-color .2s;flex:1;min-width:100px;}
.gr-sel:focus{border-color:var(--a);}
.gr-sel:disabled{opacity:.5;cursor:default;}
.gr-view-toggle{display:flex;border-radius:10px;overflow:hidden;border:1.5px solid var(--bdr2);flex-shrink:0;}
.gr-vbtn{padding:6px 12px;font-size:12px;font-weight:700;cursor:pointer;background:var(--card2);color:var(--tx3);border:none;font-family:var(--font);transition:all .15s;white-space:nowrap;}
.gr-vbtn.on{background:var(--a);color:#fff;}
.gr-save-all-btn{padding:7px 16px;border-radius:9px;background:var(--a);color:#fff;border:none;font-size:13px;font-weight:700;cursor:pointer;font-family:var(--font);white-space:nowrap;box-shadow:0 2px 8px var(--a40);transition:all .15s;flex-shrink:0;}
.gr-save-all-btn:active{transform:scale(.95);}
.gr-dirty-count{font-size:10px;opacity:.8;margin-left:2px;}

/* layout */
.gr-main{flex:1;display:flex;overflow:hidden;}
.gr-stu-panel{width:88px;flex-shrink:0;border-right:1.5px solid var(--bdr);overflow-y:auto;-webkit-overflow-scrolling:touch;background:var(--surf);}
.gr-stu-panel::-webkit-scrollbar{width:3px;}
.gr-stu-panel::-webkit-scrollbar-thumb{background:var(--bdr2);border-radius:2px;}
.gr-content{flex:1;overflow:auto;-webkit-overflow-scrolling:touch;}
.gr-content::-webkit-scrollbar{width:4px;height:4px;}
.gr-content::-webkit-scrollbar-thumb{background:var(--bdr2);border-radius:2px;}

/* student panel */
.gr-stu-item{padding:8px 4px;border-bottom:1px solid var(--bdr);cursor:pointer;transition:background .12s;text-align:center;position:relative;user-select:none;}
.gr-stu-item:hover{background:var(--a10);}
.gr-stu-item.on{background:var(--a20);border-left:3px solid var(--a);}
.gr-stu-item.dirty-item::after{content:'●';position:absolute;top:3px;right:4px;color:#f59e0b;font-size:10px;}
.gr-stu-emoji{font-size:18px;line-height:1;margin-bottom:2px;}
.gr-stu-name{font-size:11px;font-weight:700;color:var(--tx);word-break:keep-all;}
.gr-stu-nick{font-size:9px;color:var(--tx3);}
.gr-stu-dot{width:6px;height:6px;border-radius:50%;margin:2px auto 0;}

/* empty */
.gr-empty{display:flex;flex-direction:column;align-items:center;justify-content:center;height:100%;padding:40px;text-align:center;color:var(--tx3);font-size:14px;line-height:2.4;}
.gr-empty-ico{font-size:44px;margin-bottom:10px;}

/* ══ EXCEL MODE ══ */
.gr-sheet-wrap{min-width:max-content;}
.gr-sheet{border-collapse:collapse;font-size:12px;}

/* fixed student col */
.gr-sheet .gs-fix{position:sticky;left:0;z-index:3;background:var(--surf);border:1px solid var(--bdr);padding:5px 8px;min-width:110px;width:110px;cursor:pointer;}
.gr-sheet thead .gs-fix{z-index:5;background:var(--surf2);}
.gr-sheet .gs-fix.sel,.gr-sheet .gs-fix:hover{background:var(--a10);}

/* header */
.gs-th{background:var(--surf2);border:1px solid var(--bdr);padding:5px 6px;font-size:10px;font-weight:800;color:var(--tx3);text-align:center;white-space:nowrap;}
.gs-th.sec-w{background:var(--a10);color:var(--a);font-size:11px;}
.gs-th.sec-r{background:rgba(139,92,246,.1);color:#8b5cf6;font-size:11px;}
.gs-th.sec-c{background:rgba(5,150,105,.08);color:var(--green);font-size:11px;}
.gs-th.sortable{cursor:pointer;user-select:none;}
.gs-th.sortable:hover{background:var(--a20);}
.gs-th.sort-on{color:var(--a);background:var(--a10);}

/* data cells */
.gs-td{border:1px solid var(--bdr);text-align:center;padding:0;vertical-align:middle;}
/* 계산값 (읽기 전용) - 연한 배경으로 구분 */
.gs-td.ro{background:var(--surf2);}
.gs-td.ro .gs-val{padding:5px 6px;font-size:13px;font-weight:700;display:block;}
.gs-td.ro.pass-c .gs-val{color:#16a34a;}
.gs-td.ro.fail-c .gs-val{color:#f97316;}
.gs-td.ro.score-c .gs-val{color:var(--a);}
.gs-td.ro.achv-c .gs-val{color:#8b5cf6;font-weight:800;}
/* 입력 가능 셀 - 밝은 흰색/강조 배경 */
.gs-td.inp-cell{background:#fff;}
.dark .gs-td.inp-cell{background:rgba(255,255,255,.04);}

/* selected row highlight */
.gr-sheet tbody tr.sel-row .gs-td:not(.ro){background:rgba(99,102,241,.06)!important;}
.gr-sheet tbody tr.sel-row .gs-fix{background:var(--a20)!important;}

/* number input */
.gs-inp{width:100%;min-width:54px;padding:6px 4px;border:none;outline:none;background:transparent;font-size:13px;font-weight:700;color:var(--a);text-align:center;font-family:var(--font);-moz-appearance:textfield;cursor:text;}
.gs-inp::-webkit-outer-spin-button,.gs-inp::-webkit-inner-spin-button{-webkit-appearance:none;}
.gs-inp:focus{background:rgba(99,102,241,.08);border-radius:4px;}

/* comment */
.gs-cm-cell{min-width:${CM_W}px;width:${CM_W}px;}
.gs-cm-inp{width:100%;padding:5px 8px;border:none;outline:none;background:transparent;font-size:11px;color:var(--tx);font-family:var(--font);resize:none;height:52px;line-height:1.5;cursor:text;box-sizing:border-box;}
.gs-cm-inp:focus{background:rgba(5,150,105,.05);}

/* average row */
.gr-avg-row td{background:var(--surf2)!important;}
.gr-avg-row .gs-fix{color:var(--a);font-weight:800;font-size:12px;}

/* chart */
.gr-chart-wrap{padding:10px 12px 6px;border-top:1.5px solid var(--bdr);background:var(--surf2);flex-shrink:0;}
.gr-chart-title{font-size:11px;font-weight:800;color:var(--tx3);letter-spacing:.5px;margin-bottom:6px;}
.gr-chart-canvas{width:100%;height:72px;display:block;}

/* context menu */
.gr-ctx-menu{position:fixed;z-index:999;background:var(--card);border:1.5px solid var(--bdr);border-radius:10px;box-shadow:var(--sh2);padding:4px;min-width:120px;animation:cardIn .12s ease;}
.gr-ctx-item{padding:8px 14px;font-size:13px;font-weight:700;cursor:pointer;border-radius:7px;transition:background .12s;}
.gr-ctx-item:hover{background:var(--a10);color:var(--a);}
.gr-ctx-item.red:hover{background:rgba(220,38,38,.08);color:#ef4444;}
.gr-ctx-divider{height:1px;background:var(--bdr);margin:3px 0;}

/* ══ CARD MODE ══ */
.gr-carousel-wrap{flex:1;display:flex;flex-direction:column;overflow:hidden;}
.gr-carousel{flex:1;display:flex;overflow:hidden;position:relative;}
/* 3:4:3 ratio */
.gr-slide{flex-shrink:0;overflow-y:auto;-webkit-overflow-scrolling:touch;transition:all .3s cubic-bezier(.4,0,.2,1);}
.gr-slide.prev{flex:3;opacity:.5;filter:brightness(.85);cursor:pointer;border-right:1px solid var(--bdr);}
.gr-slide.curr{flex:4;position:relative;z-index:2;}
.gr-slide.next{flex:3;opacity:.5;filter:brightness(.85);cursor:pointer;border-left:1px solid var(--bdr);}
.gr-slide.solo{flex:1;}
@media(max-width:600px){
  .gr-slide.prev{flex:0;width:36px;overflow:hidden;}
  .gr-slide.next{flex:0;width:36px;overflow:hidden;}
  .gr-slide.curr{flex:1;}
}
.gr-carousel-nav{display:flex;justify-content:center;gap:5px;padding:6px;flex-shrink:0;background:var(--surf2);border-top:1px solid var(--bdr);}
.gr-carousel-dot{width:6px;height:6px;border-radius:50%;background:var(--bdr2);cursor:pointer;transition:all .2s;}
.gr-carousel-dot.on{background:var(--a);width:18px;border-radius:3px;}

/* card body */
.gr-card-body{padding:10px;}
.gr-card-hero{display:flex;align-items:center;gap:10px;padding:10px 12px;background:linear-gradient(135deg,var(--a10),rgba(5,150,105,.06));border-radius:12px;margin-bottom:10px;border:1px solid var(--a40);}
.gr-hero-emo{font-size:24px;flex-shrink:0;}
.gr-hero-nm{font-size:14px;font-weight:900;color:var(--tx);}
.gr-hero-sub{font-size:10px;color:var(--tx3);margin-top:1px;}
.gr-hero-score{margin-left:auto;text-align:right;}
.gr-hero-pct{font-size:20px;font-weight:900;line-height:1;}
.gr-hero-lbl{font-size:9px;color:var(--tx3);}

/* card section */
.gr-csec{margin-bottom:10px;}
.gr-csec-head{display:flex;align-items:center;justify-content:space-between;padding:7px 10px;background:var(--surf2);border-radius:9px 9px 0 0;border:1px solid var(--bdr);border-bottom:none;}
.gr-csec-title{font-size:12px;font-weight:800;color:var(--tx);display:flex;align-items:center;gap:4px;}
.gr-csec-badge{padding:2px 8px;border-radius:20px;font-size:10px;font-weight:700;background:var(--a10);color:var(--a);border:1px solid var(--a40);}
.gr-csec-badge.rd{background:rgba(139,92,246,.1);color:#8b5cf6;border-color:rgba(139,92,246,.3);}
.gr-card-grid{border:1px solid var(--bdr);border-radius:0 0 9px 9px;overflow:hidden;}
.gr-crow{display:flex;border-bottom:1px solid var(--bdr);}
.gr-crow:last-child{border-bottom:none;}
.gr-clbl{padding:8px 10px;font-size:11px;font-weight:700;color:var(--tx2);background:var(--surf2);border-right:1px solid var(--bdr);min-width:80px;flex-shrink:0;display:flex;align-items:center;}
.gr-cval{padding:5px 8px;flex:1;display:flex;align-items:center;gap:5px;}
/* 카드 입력 */
.gr-cinp{flex:1;padding:6px 8px;border-radius:7px;background:var(--surf2);border:1.5px solid var(--bdr);font-size:13px;font-weight:700;color:var(--a);text-align:center;outline:none;font-family:var(--font);-moz-appearance:textfield;transition:border-color .2s;}
.gr-cinp::-webkit-outer-spin-button,.gr-cinp::-webkit-inner-spin-button{-webkit-appearance:none;}
.gr-cinp:focus{border-color:var(--a);background:var(--a10);}
/* 카드 계산값 */
.gr-cdisp{min-width:48px;text-align:center;padding:6px 6px;font-size:13px;font-weight:800;border-radius:7px;background:var(--surf2);border:1px solid var(--bdr);color:var(--tx3);}
.gr-cdisp.pass{color:#16a34a;background:rgba(22,163,74,.08);border-color:rgba(22,163,74,.2);}
.gr-cdisp.fail{color:#f97316;background:rgba(249,115,22,.08);border-color:rgba(249,115,22,.2);}
.gr-cdisp.score{color:var(--a);background:var(--a10);border-color:var(--a40);}
.gr-cdisp.achv{color:#8b5cf6;background:rgba(139,92,246,.1);border-color:rgba(139,92,246,.25);}

/* card comment */
.gr-card-cmt{width:100%;box-sizing:border-box;padding:8px 10px;border:none;outline:none;background:transparent;font-size:12px;color:var(--tx);font-family:var(--font);resize:none;min-height:60px;line-height:1.8;}
.gr-card-cmt:focus{background:rgba(5,150,105,.04);}
.gr-card-save-row{display:flex;gap:7px;padding:8px;}
.gr-card-save-btn{flex:1;padding:10px;border:none;background:var(--a);color:#fff;font-size:12px;font-weight:800;cursor:pointer;font-family:var(--font);border-radius:9px;transition:all .15s;}
.gr-card-save-btn:active{opacity:.85;}

/* ══ REPORT ══ */
.gr-report-panel{padding:14px 14px 80px;}
.gr-rpt-cfg{background:var(--card);border:1px solid var(--bdr);border-radius:12px;padding:12px 14px;margin-bottom:12px;box-shadow:var(--sh);}
.gr-rpt-cfg-title{font-size:11px;font-weight:800;color:var(--tx3);letter-spacing:.5px;margin-bottom:8px;}
.gr-rpt-layouts{display:flex;gap:6px;flex-wrap:wrap;margin-bottom:8px;}
.gr-rpt-lbtn{width:40px;height:40px;border-radius:8px;border:2px solid var(--bdr2);background:var(--surf2);font-size:10px;font-weight:800;cursor:pointer;color:var(--tx3);display:flex;align-items:center;justify-content:center;transition:all .15s;font-family:var(--font);}
.gr-rpt-lbtn.on{border-color:var(--a);background:var(--a20);color:var(--a);}
.gr-rpt-toggle{display:flex;align-items:center;gap:8px;font-size:12px;color:var(--tx2);cursor:pointer;}
.gr-rpt-toggle input{accent-color:var(--a);}
.gr-rpt-preview{background:var(--card);border:1px solid var(--bdr);border-radius:12px;overflow:hidden;box-shadow:var(--sh);animation:cardIn .2s ease;}
.rpt-wrap{padding:20px 24px;font-family:'Noto Sans KR',sans-serif;font-size:13px;color:#111;background:#fff;}
.rpt-header{display:flex;align-items:center;gap:14px;margin-bottom:16px;}
.rpt-title{font-size:20px;font-weight:900;color:#111;flex:1;}
.rpt-divider{border:none;border-top:2px solid #e5e7eb;margin:10px 0;}
.rpt-info p{margin:4px 0;font-size:13px;}
.rpt-sec-title{font-size:14px;font-weight:800;color:#111;margin:14px 0 6px;}
.rpt-tbl{width:100%;border-collapse:collapse;font-size:12px;margin-bottom:12px;}
.rpt-tbl th{background:#f1f5f9;padding:7px 10px;text-align:center;font-size:11px;font-weight:800;color:#475569;border:1px solid #e2e8f0;}
.rpt-tbl td{border:1px solid #e2e8f0;padding:7px 10px;text-align:center;font-size:13px;}
.rpt-tbl .rpt-pass{color:#16a34a;font-weight:700;}
.rpt-tbl .rpt-fail{color:#ea580c;font-weight:700;}
.rpt-tbl .rpt-achv{color:#8b5cf6;font-weight:800;}
.rpt-tbl .rpt-avg td{background:#f8fafc;font-weight:700;}
.rpt-comment-box{border:1.5px solid #e2e8f0;border-radius:8px;padding:12px 14px;min-height:60px;font-size:12px;color:#374151;line-height:1.8;background:#fafafa;}
.rpt-graph-wrap{margin:8px 0 12px;}
.rpt-acts{display:flex;gap:8px;flex-wrap:wrap;padding:10px 14px 14px;}
.rpt-btn{flex:1;min-width:60px;padding:10px 6px;border-radius:10px;border:none;font-size:12px;font-weight:700;cursor:pointer;font-family:var(--font);transition:all .15s;}
.rpt-btn.copy{background:var(--a10);color:var(--a);border:1px solid var(--a40);}
.rpt-btn.share{background:var(--a);color:#fff;box-shadow:0 3px 10px var(--a40);}
.rpt-btn.pdf{background:rgba(5,150,105,.1);color:var(--green);border:1px solid rgba(5,150,105,.3);}
.rpt-btn.cap{background:rgba(245,158,11,.1);color:#d97706;border:1px solid rgba(245,158,11,.3);}
.rpt-btn:active{transform:scale(.95);}

/* report modal */
.gr-rpt-sh-scroll{flex:1;overflow-y:auto;}
.gr-share-box{background:var(--surf2);border-radius:10px;padding:12px 14px;font-size:12px;line-height:1.9;color:var(--tx);white-space:pre-wrap;word-break:break-all;border:1px solid var(--bdr);max-height:260px;overflow-y:auto;font-family:var(--font);margin:8px 0;}
.gr-sacts{display:flex;gap:8px;}
.gr-sbtn{flex:1;padding:11px 8px;border-radius:10px;border:none;font-size:13px;font-weight:700;cursor:pointer;font-family:var(--font);transition:all .15s;}
.gr-sbtn.copy{background:var(--a10);color:var(--a);border:1px solid var(--a40);}
.gr-sbtn.share{background:var(--a);color:#fff;box-shadow:0 3px 10px var(--a40);}
.gr-sbtn:active{transform:scale(.96);}

@media print{body>*:not(#gr-pf){display:none!important;}#gr-pf{display:block!important;position:fixed;inset:0;z-index:9999;background:#fff;padding:24px;overflow:auto;}}
#gr-pf{display:none;}
`;
    document.head.appendChild(s);
  }

  /* ══ INIT ══ */
  async function init() {
    _css();
    if (typeof GradeDB === 'undefined') { console.warn('[GradeApp] GradeDB not loaded'); return; }
    await GradeDB.init();
    window.addEventListener('beforeunload', e => { if (_st.dirty.size > 0) { e.preventDefault(); e.returnValue = ''; } });
    document.addEventListener('click', _closeCtxMenu);
    console.log('[GradeApp] ✅ v4.1');
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
    const hasData = _st.classId && _st.bookId;
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
          <button class="ibtn" id="gr-rpt-btn" title="전체 성적표" onclick="GradeApp.openReport()" style="display:none">📋</button>
        </div>
      </div>
      <div class="gr-toolbar">
        <select class="gr-sel" id="gr-csel" onchange="GradeApp._onCls(this.value)">
          <option value="">— 반 선택 —</option>
        </select>
        <select class="gr-sel" id="gr-bsel" onchange="GradeApp._onBk(this.value)" disabled>
          <option value="">— 교재 선택 —</option>
        </select>
        <div class="gr-view-toggle">
          <button class="gr-vbtn ${_st.viewMode==='excel'?'on':''}"  onclick="GradeApp._setView('excel')">🔲 엑셀</button>
          <button class="gr-vbtn ${_st.viewMode==='card'?'on':''}"   onclick="GradeApp._setView('card')">👤 카드</button>
          <button class="gr-vbtn ${_st.viewMode==='report'?'on':''}" onclick="GradeApp._setView('report')">📄 리포트</button>
        </div>
        ${hasData ? `<button class="gr-save-all-btn" onclick="GradeApp.saveAll()">
          💾 저장<span class="gr-dirty-count" id="gr-dirty-cnt">${_st.dirty.size?`(${_st.dirty.size})`:''}</span>
        </button>` : ''}
      </div>
      <div class="gr-main">
        <div class="gr-stu-panel" id="gr-stu-panel"></div>
        <div class="gr-content" id="gr-content"></div>
      </div>
      ${_st.viewMode==='excel' && hasData ? `
        <div class="gr-chart-wrap" id="gr-chart-wrap">
          <div class="gr-chart-title">📊 성취율 현황 <span style="font-size:10px;font-weight:400;color:var(--tx3)">· 학생 클릭 시 하이라이트</span></div>
          <canvas class="gr-chart-canvas" id="gr-chart"></canvas>
        </div>` : ''}
      <!-- 성적표 모달 -->
      <div id="gr-rpt-ov" class="ov hidden" onclick="if(event.target.id==='gr-rpt-ov')GradeApp.closeReport()">
        <div class="sh" id="gr-rpt-sh" onclick="event.stopPropagation()" style="max-height:92vh;display:flex;flex-direction:column;"></div>
      </div>
      <!-- 컨텍스트 메뉴 -->
      <div id="gr-ctx" class="gr-ctx-menu" style="display:none"></div>
    `;
  }

  /* ── 셀렉트 ── */
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
    if (!books.length) { sel.innerHTML = `<option value="">배정된 교재 없음</option>`; sel.disabled = true; }
  }

  /* ── 학생 패널 ── */
  function _renderStudents() {
    const panel = document.getElementById('gr-stu-panel'); if (!panel) return;
    const students = _getSorted();
    if (!students.length) { panel.innerHTML = ''; return; }
    panel.innerHTML = students.map((s, i) => {
      const rec  = _st.classId && _st.bookId ? GradeDB.getLatest(_st.classId, s.id, _st.bookId) : null;
      const achW = rec?.word?.totalQ > 0 ? Math.round(rec.word.pass / rec.word.totalQ * 100) : null;
      const dotClr = achW != null ? (achW >= 80 ? '#16a34a' : '#f97316') : 'var(--bdr2)';
      const isSel  = _st.viewMode === 'card' ? _st.slideIdx === i : _st.studentId === s.id;
      const isDirty= _st.dirty.has(s.id);
      return `<div class="gr-stu-item ${isSel?'on':''} ${isDirty?'dirty-item':''}"
                   onclick="GradeApp._onStu('${s.id}',${i})">
        <div class="gr-stu-emoji">${_emoji(s, achW)}</div>
        <div class="gr-stu-name">${_e(s.name)}</div>
        ${s.nickname?`<div class="gr-stu-nick">(${_e(s.nickname)})</div>`:''}
        <div class="gr-stu-dot" style="background:${dotClr};opacity:${achW!=null?1:.28}"></div>
      </div>`;
    }).join('');
  }

  function _emoji(s, achW) {
    const fem = ['자','연','지','수','은','민','서','현','유','아','나','하','소','예','혜','미','린','진'].some(c=>s.name.includes(c));
    const pool = fem ? ANIMALS_F : ANIMALS_M;
    return pool[(s.name.charCodeAt(0)||0) % pool.length];
  }

  /* ── 정렬된 학생 ── */
  function _getSorted() {
    const all = _getStudents();
    if (!_st.sortCol) return all;
    return [...all].sort((a, b) => {
      let va = 0, vb = 0;
      if (_st.sortCol === 'name') {
        va = a.name; vb = b.name;
        return _st.sortDesc ? vb.localeCompare(va,'ko') : va.localeCompare(vb,'ko');
      }
      const ra = GradeDB.getLatest(_st.classId, a.id, _st.bookId);
      const rb = GradeDB.getLatest(_st.classId, b.id, _st.bookId);
      if (_st.sortCol === 'wordAch') {
        va = ra?.word?.totalQ > 0 ? ra.word.pass / ra.word.totalQ : -1;
        vb = rb?.word?.totalQ > 0 ? rb.word.pass / rb.word.totalQ : -1;
      } else {
        const cfg = GradeDB.getReportConfig(_st.bookId);
        const revs = GradeDB.getActiveReviews(_st.bookId);
        va = _calcRdN(ra?.reading||{}, revs) ?? -1;
        vb = _calcRdN(rb?.reading||{}, revs) ?? -1;
      }
      return _st.sortDesc ? vb - va : va - vb;
    });
  }

  function _toggleSort(col) {
    if (_st.sortCol === col) _st.sortDesc = !_st.sortDesc;
    else { _st.sortCol = col; _st.sortDesc = true; }
    _renderStudents();
    _renderContent();
  }

  /* ── 콘텐츠 ── */
  function _renderContent() {
    const cnt = document.getElementById('gr-content'); if (!cnt) return;
    if (!_st.classId || !_st.bookId) {
      cnt.innerHTML = `<div class="gr-empty"><div class="gr-empty-ico">📝</div>반과 교재를 선택하세요</div>`; return;
    }
    const sts = _getSorted();
    if (!sts.length) { cnt.innerHTML = `<div class="gr-empty"><div class="gr-empty-ico">👨‍🎓</div>재원 학생이 없습니다</div>`; return; }
    if (_st.viewMode === 'excel')  _renderExcel(cnt, sts);
    else if (_st.viewMode === 'card')   _renderCarousel(cnt, sts);
    else if (_st.viewMode === 'report') _renderReportView(cnt, sts);
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

    students.forEach(s => _ensureData(s.id));

    /* 헤더: 정렬 아이콘 */
    const wIcon  = _st.sortCol==='wordAch'?(_st.sortDesc?'↓':'↑'):'↕';
    const rdIcon = _st.sortCol==='rdAch'  ?(_st.sortDesc?'↓':'↑'):'↕';
    const nmIcon = _st.sortCol==='name'   ?(_st.sortDesc?'↓':'↑'):'↕';

    /*
     * 헤더 구조 (3행):
     * Row1: [학생(3행)] [단어평가(4col)] [리딩평가(1+rvN+rvN+1)] [Comment(3행)]
     * Row2: 단어서브4 | [총문제(2행)] [정답수(rvN)] [점수(rvN)] [성취율(2행)]
     * Row3: 단어비어있음 | rvN서브Q | rvN서브S
     */
    const rdH1span = hasRd ? 1 + rvN * 2 + 1 : 0;
    const rdSection = hasRd ? `<th class="gs-th sec-r" colspan="${rdH1span}">📖 리딩 평가</th>` : '';
    const rdRow2 = hasRd ? `
      <th class="gs-th" rowspan="2" style="background:rgba(139,92,246,.06);color:#8b5cf6;vertical-align:middle">총문제</th>
      <th class="gs-th sec-r" colspan="${rvN}">정답 수</th>
      <th class="gs-th sec-r" colspan="${rvN}">점수</th>
      <th class="gs-th sortable sec-r ${_st.sortCol==='rdAch'?'sort-on':''}" rowspan="2"
          onclick="GradeApp._toggleSort('rdAch')" style="vertical-align:middle">성취율 ${rdIcon}</th>` : '';
    const rdRow3Q = hasRd ? actRevs.map(rv=>`<th class="gs-th" style="font-size:9px;background:rgba(139,92,246,.04)">${_e(rv.name)}</th>`).join('') : '';
    const rdRow3S = hasRd ? actRevs.map(rv=>`<th class="gs-th" style="font-size:9px;background:rgba(139,92,246,.04)">${_e(rv.name)}</th>`).join('') : '';

    const html = `
      <div class="gr-sheet-wrap">
        <table class="gr-sheet" oncontextmenu="GradeApp._onCtxTable(event)">
          <thead>
            <tr>
              <th class="gs-fix gs-th sortable ${_st.sortCol==='name'?'sort-on':''}" rowspan="3"
                  onclick="GradeApp._toggleSort('name')">학생 ${nmIcon}</th>
              <th class="gs-th sec-w" colspan="4">🔤 단어 평가</th>
              ${rdSection}
              <th class="gs-th sec-c" rowspan="3" style="min-width:${CM_W}px">💬 Teacher's Comment</th>
            </tr>
            <tr>
              <th class="gs-th" rowspan="2" style="background:var(--a10);color:var(--a);vertical-align:middle">총 테스트<br>(문제) 수</th>
              <th class="gs-th" rowspan="2" style="background:var(--a10);color:var(--a);vertical-align:middle">재시험</th>
              <th class="gs-th" rowspan="2" style="background:var(--a10);vertical-align:middle">통과</th>
              <th class="gs-th sortable ${_st.sortCol==='wordAch'?'sort-on':''}" rowspan="2"
                  onclick="GradeApp._toggleSort('wordAch')" style="background:var(--a10);vertical-align:middle">성취율 ${wIcon}</th>
              ${rdRow2}
            </tr>
            <tr>
              ${rdRow3Q}${rdRow3S}
            </tr>
          </thead>
          <tbody>
            ${students.map((s,ri) => _excelRow(s, ri, config, totalWQ, actRevs, totalRQ, hasRd)).join('')}
            ${_avgRow(students, config, totalWQ, actRevs, hasRd)}
          </tbody>
        </table>
      </div>`;
    cnt.innerHTML = html;
    setTimeout(() => _renderChart(students, actRevs), 30);
  }

  function _excelRow(s, ri, config, totalWQ, actRevs, totalRQ, hasRd) {
    const d   = _st.data[s.id] || {};
    const wd  = d.word || {};
    const rd  = d.reading || {};
    const isSel = _st.studentId === s.id;

    const retake = wd.retake !== undefined && wd.retake !== '' ? Number(wd.retake) : '';
    const pass   = retake !== '' ? Math.max(0, totalWQ - retake) : '';
    const achW   = pass !== '' && totalWQ > 0 ? Math.round(pass / totalWQ * 100) : '';
    const isGW   = achW !== '' && achW >= 80;

    const rdCells = hasRd ? `
      <td class="gs-td ro" style="background:rgba(139,92,246,.04)">
        <span class="gs-val" style="color:#8b5cf6">${totalRQ||'—'}</span>
      </td>
      ${actRevs.map((rv,i) => {
        const key = `R${i}`;
        const v   = rd[key]?.correct !== '' && rd[key]?.correct != null ? rd[key].correct : '';
        return `<td class="gs-td inp-cell"><input class="gs-inp" type="number" min="0" max="${totalRQ}" step="1"
          value="${v}" placeholder="—" data-sid="${s.id}" data-rkey="${key}" data-row="${ri}"
          oninput="GradeApp._excelRdInput('${s.id}','${key}',this.value,'${totalRQ}')"
          onkeydown="GradeApp._onKey(event,'${s.id}',${ri},'rd')"></td>`;
      }).join('')}
      ${actRevs.map((rv,i) => {
        const key = `R${i}`;
        const sc  = rd[key]?.score !== '' && rd[key]?.score != null ? rd[key].score : '';
        return `<td class="gs-td ro" id="gr-sc-${s.id}-${key}"><span class="gs-val score-c">${sc!==''?sc:'—'}</span></td>`;
      }).join('')}
      <td class="gs-td ro" id="gr-achvrd-${s.id}"><span class="gs-val achv-c">${_calcRdStr(rd,actRevs)||'—'}</span></td>` : '';

    return `
      <tr class="${isSel?'sel-row':''}" id="gr-row-${s.id}" data-sid="${s.id}">
        <td class="gs-fix ${isSel?'sel':''}" onclick="GradeApp._onStu('${s.id}',${ri})">
          <div style="display:flex;align-items:center;gap:5px">
            <span style="font-size:15px">${_emoji(s, achW!==''?achW:null)}</span>
            <div>
              <div style="font-weight:700;font-size:12px;color:var(--tx)">${_e(s.name)}${s.nickname?` <span style="font-size:10px;color:var(--tx3)">(${_e(s.nickname)})</span>`:''}</div>
              ${d.savedAt?`<div style="font-size:8px;color:var(--tx3)">${d.savedAt.slice(5,16)}</div>`:''}
            </div>
          </div>
        </td>
        <td class="gs-td ro"><span class="gs-val score-c">${totalWQ||'—'}</span></td>
        <td class="gs-td inp-cell">
          <input class="gs-inp" type="number" min="0" max="${totalWQ}" step="1"
            value="${retake}" placeholder="0"
            id="gr-retake-${s.id}" data-sid="${s.id}" data-row="${ri}" data-col="word"
            oninput="GradeApp._excelWordInput('${s.id}',this.value,'${totalWQ}')"
            onkeydown="GradeApp._onKey(event,'${s.id}',${ri},'word')">
        </td>
        <td class="gs-td ro ${pass!==''?(isGW?'pass-c':'fail-c'):''}" id="gr-pass-${s.id}">
          <span class="gs-val">${pass!==''?pass:'—'}</span>
        </td>
        <td class="gs-td ro ${achW!==''?(isGW?'pass-c':'fail-c'):''}" id="gr-achvw-${s.id}">
          <span class="gs-val">${achW!==''?achW+'%':'—'}</span>
        </td>
        ${rdCells}
        <td class="gs-td inp-cell gs-cm-cell">
          <textarea class="gs-cm-inp" id="gr-cmt-${s.id}"
            oninput="GradeApp._excelComment('${s.id}',this.value)">${_e(d.comment||'')}</textarea>
        </td>
      </tr>`;
  }

  /* 평균 행 */
  function _avgRow(students, config, totalWQ, actRevs, hasRd) {
    const achWs = students.map(s => {
      const d = _st.data[s.id];
      if (d?.word?.pass != null && d?.word?.totalQ > 0) return Math.round(d.word.pass / d.word.totalQ * 100);
      const rec = GradeDB.getLatest(_st.classId, s.id, _st.bookId);
      return rec?.word?.totalQ > 0 ? Math.round(rec.word.pass / rec.word.totalQ * 100) : null;
    }).filter(v => v != null);
    const avgW = achWs.length ? Math.round(achWs.reduce((a,b)=>a+b,0)/achWs.length) : null;

    let rdAvgCells = '';
    if (hasRd) {
      const achRds = students.map(s => {
        const d = _st.data[s.id];
        const rec = GradeDB.getLatest(_st.classId, s.id, _st.bookId);
        const rd = d?.reading || rec?.reading || {};
        return _calcRdN(rd, actRevs);
      }).filter(v => v != null);
      const avgRd = achRds.length ? Math.round(achRds.reduce((a,b)=>a+b,0)/achRds.length) : null;
      rdAvgCells = `
        <td class="gs-td ro"></td>
        ${actRevs.map(()=>'<td class="gs-td ro"></td>').join('')}
        ${actRevs.map(()=>'<td class="gs-td ro"></td>').join('')}
        <td class="gs-td ro"><span class="gs-val achv-c" id="gr-avg-rd">${avgRd!=null?avgRd+'%':'—'}</span></td>`;
    }

    return `<tr class="gr-avg-row">
      <td class="gs-fix" colspan="3" style="text-align:center;font-weight:800;color:var(--a);font-size:12px">📊 평균</td>
      <td class="gs-td ro"></td>
      <td class="gs-td ro"><span class="gs-val achv-c" id="gr-avg-w">${avgW!=null?avgW+'%':'—'}</span></td>
      ${rdAvgCells}
      <td class="gs-td ro gs-cm-cell"></td>
    </tr>`;
  }

  /* 평균 실시간 갱신 */
  function _updateAvg() {
    const students = _getSorted();
    const config   = GradeDB.getReportConfig(_st.bookId);
    const actRevs  = GradeDB.getActiveReviews(_st.bookId);
    const hasRd    = config.reading?.enabled && actRevs.length > 0;

    const achWs = students.map(s => {
      const d = _st.data[s.id];
      if (!d?.word?.totalQ) return null;
      const p = d.word.pass !== '' ? d.word.pass : null;
      return p != null && d.word.totalQ > 0 ? Math.round(p / d.word.totalQ * 100) : null;
    }).filter(v => v != null);
    const avgW = achWs.length ? Math.round(achWs.reduce((a,b)=>a+b,0)/achWs.length) : null;
    const avgWEl = document.getElementById('gr-avg-w');
    if (avgWEl) avgWEl.textContent = avgW != null ? avgW + '%' : '—';

    if (hasRd) {
      const achRds = students.map(s => _calcRdN(_st.data[s.id]?.reading||{}, actRevs)).filter(v=>v!=null);
      const avgRd  = achRds.length ? Math.round(achRds.reduce((a,b)=>a+b,0)/achRds.length) : null;
      const avgRdEl = document.getElementById('gr-avg-rd');
      if (avgRdEl) avgRdEl.textContent = avgRd != null ? avgRd + '%' : '—';
    }
  }

  /* ── 엑셀 입력 핸들러 ── */
  function _excelWordInput(sid, val, totalWQ) {
    _ensureData(sid);
    const tq     = Number(totalWQ) || 0;
    const retake = val === '' ? '' : Math.max(0, Math.min(tq, Number(val)));
    const pass   = retake !== '' ? Math.max(0, tq - retake) : '';
    const achW   = pass !== '' && tq > 0 ? Math.round(pass / tq * 100) : '';
    const isGW   = achW !== '' && achW >= 80;
    _st.data[sid].word = { totalQ:tq, retake, pass };
    _st.dirty.add(sid); _refreshDirtyUI();

    const passEl = document.getElementById(`gr-pass-${sid}`);
    const achEl  = document.getElementById(`gr-achvw-${sid}`);
    if (passEl) { passEl.querySelector('.gs-val').textContent = pass!==''?pass:'—'; passEl.className=`gs-td ro ${pass!==''?(isGW?'pass-c':'fail-c'):''}`; }
    if (achEl)  { achEl.querySelector('.gs-val').textContent  = achW!==''?achW+'%':'—'; achEl.className =`gs-td ro ${achW!==''?(isGW?'pass-c':'fail-c'):''}`; }
    _updateAvg(); _updateChart();
  }

  function _excelRdInput(sid, key, val, totalRQ) {
    _ensureData(sid);
    const tq      = Number(totalRQ) || 0;
    const correct = val === '' ? '' : Math.max(0, Math.min(tq, Number(val)));
    const score   = correct !== '' ? Math.round(correct / tq * 100 * 10) / 10 : '';
    if (!_st.data[sid].reading) _st.data[sid].reading = {};
    _st.data[sid].reading[key] = { correct, score };
    _st.dirty.add(sid); _refreshDirtyUI();

    const actRevs = GradeDB.getActiveReviews(_st.bookId);
    const scEl    = document.getElementById(`gr-sc-${sid}-${key}`);
    const achvEl  = document.getElementById(`gr-achvrd-${sid}`);
    if (scEl) scEl.querySelector('.gs-val').textContent = score !== '' ? score : '—';
    if (achvEl) achvEl.querySelector('.gs-val').textContent = _calcRdStr(_st.data[sid].reading, actRevs) || '—';
    _updateAvg();
  }

  function _excelComment(sid, val) {
    _ensureData(sid); _st.data[sid].comment = val;
    _st.dirty.add(sid); _refreshDirtyUI();
  }

  /* ── 키보드 네비게이션 ── */
  function _onKey(e, sid, rowIdx, type) {
    const students = _getSorted();
    if (e.key === 'Enter') {
      /* Enter → 다음 학생 같은 컬럼 */
      e.preventDefault();
      const next = students[rowIdx + 1];
      if (next) {
        if (type === 'word') document.getElementById(`gr-retake-${next.id}`)?.focus();
        else document.querySelector(`[data-sid="${next.id}"][data-rkey="${type==='rd'?e.target.dataset.rkey:'R0'}"]`)?.focus();
        _onStu(next.id, rowIdx + 1);
      }
    } else if (e.key === 'ArrowUp' || e.key === 'ArrowDown') {
      /* 방향키 ↑↓ → 값 증가/감소 (기본 동작, 이동 안 함) */
      /* 기본 number input 동작 유지 */
    }
  }

  /* ── 우클릭 컨텍스트 메뉴 ── */
  function _onCtxTable(e) {
    e.preventDefault();
    const row = e.target.closest('tr[data-sid]');
    if (!row) return;
    const sid = row.dataset.sid;
    _showCtxMenu(e.clientX, e.clientY, sid);
  }

  function _showCtxMenu(x, y, sid) {
    const menu = document.getElementById('gr-ctx'); if (!menu) return;
    const s = _getSorted().find(s=>s.id===sid);
    menu.innerHTML = `
      <div style="padding:6px 12px;font-size:10px;font-weight:800;color:var(--tx3);letter-spacing:.5px">
        ${s?_e(s.name):'학생'}
      </div>
      <div class="gr-ctx-divider"></div>
      <div class="gr-ctx-item" onclick="GradeApp.saveOne('${sid}');GradeApp._closeCtxMenu()">💾 저장</div>
      <div class="gr-ctx-item red" onclick="GradeApp.resetOne('${sid}');GradeApp._closeCtxMenu()">🗑 초기화</div>`;
    /* 화면 밖으로 나가지 않게 */
    const menuW = 150, menuH = 100;
    menu.style.left = Math.min(x, window.innerWidth - menuW) + 'px';
    menu.style.top  = Math.min(y, window.innerHeight - menuH) + 'px';
    menu.style.display = 'block';
  }

  function _closeCtxMenu(e) {
    const menu = document.getElementById('gr-ctx');
    if (menu && (!e || !menu.contains(e.target))) menu.style.display = 'none';
  }

  /* ── 차트 ── */
  function _renderChart(students, actRevs) {
    const canvas = document.getElementById('gr-chart'); if (!canvas) return;
    const config  = GradeDB.getReportConfig(_st.bookId||'');
    const hasRd   = config.reading?.enabled && actRevs.length > 0;
    const W = canvas.offsetWidth || 300, H = 72;
    canvas.width = W * window.devicePixelRatio; canvas.height = H * window.devicePixelRatio;
    canvas.style.width = W + 'px'; canvas.style.height = H + 'px';
    const ctx = canvas.getContext('2d');
    ctx.scale(window.devicePixelRatio, window.devicePixelRatio);
    ctx.clearRect(0, 0, W, H);

    if (!students.length) return;
    const n = students.length;
    const bw = Math.max(8, Math.min(24, (W / n / (hasRd?2:1)) - 3));
    const gapX = W / n;

    students.forEach((s, i) => {
      const d    = _st.data[s.id] || {};
      const rec  = GradeDB.getLatest(_st.classId, s.id, _st.bookId);
      const wd   = d.word || rec?.word || {};
      const rdD  = d.reading || rec?.reading || {};
      const achW = wd.totalQ > 0 && wd.pass != null ? Math.round(wd.pass / wd.totalQ * 100) : 0;
      const achR = hasRd ? (_calcRdN(rdD, actRevs) ?? 0) : 0;
      const isHL = _st.studentId === s.id;
      const x    = i * gapX + gapX / 2;
      const maxH = H - 18;

      if (hasRd) {
        ctx.fillStyle = isHL ? '#6366f1' : '#a5b4fc';
        ctx.globalAlpha = isHL ? 1 : .8;
        const hw = Math.round(achW / 100 * maxH);
        ctx.fillRect(x - bw - 1, maxH - hw, bw, hw);
        const rw = Math.round(achR / 100 * maxH);
        ctx.fillStyle = isHL ? '#8b5cf6' : '#c4b5fd';
        ctx.fillRect(x + 1, maxH - rw, bw, rw);
      } else {
        const hw = Math.round(achW / 100 * maxH);
        ctx.fillStyle = isHL ? '#6366f1' : (achW >= 80 ? '#10b981' : '#f97316');
        ctx.globalAlpha = isHL ? 1 : .75;
        ctx.fillRect(x - bw/2, maxH - hw, bw, hw);
      }
      ctx.globalAlpha = 1;
      ctx.font = `${isHL?'bold ':''} 9px sans-serif`;
      ctx.fillStyle = isHL ? '#6366f1' : '#9ca3af';
      ctx.textAlign = 'center';
      ctx.fillText(_givN(s.name).slice(0,2), x, H - 3);
    });
  }

  function _updateChart() {
    const sts   = _getSorted();
    const revs  = _st.bookId ? GradeDB.getActiveReviews(_st.bookId) : [];
    _renderChart(sts, revs);
  }

  /* ════════════════════════════════════
   * 카드 캐러셀 모드
   * ════════════════════════════════════ */
  function _renderCarousel(cnt, students) {
    const idx  = Math.max(0, Math.min(_st.slideIdx, students.length - 1));
    _st.slideIdx = idx;
    const prev = students[idx - 1] || null;
    const curr = students[idx];
    const next = students[idx + 1] || null;

    cnt.innerHTML = `
      <div class="gr-carousel-wrap">
        <div class="gr-carousel" id="gr-carousel"
             ontouchstart="GradeApp._ts(event)" ontouchend="GradeApp._te(event)">
          ${prev ? `<div class="gr-slide prev" onclick="GradeApp._slideTo(${idx-1})">${_cardBody(prev, false)}</div>` : ''}
          <div class="gr-slide ${prev||next?'curr':'solo'}">${_cardBody(curr, true)}</div>
          ${next ? `<div class="gr-slide next" onclick="GradeApp._slideTo(${idx+1})">${_cardBody(next, false)}</div>` : ''}
        </div>
        <div class="gr-carousel-nav">
          ${students.map((_,i) => `<div class="gr-carousel-dot ${i===idx?'on':''}" onclick="GradeApp._slideTo(${i})"></div>`).join('')}
        </div>
      </div>`;
  }

  function _cardBody(s, active) {
    _ensureData(s.id);
    const config  = GradeDB.getReportConfig(_st.bookId);
    const totalWQ = config.word?.totalQ || 0;
    const actRevs = GradeDB.getActiveReviews(_st.bookId);
    const hasRd   = config.reading?.enabled && actRevs.length > 0;
    const totalRQ = config.reading?.totalQ || 0;
    const d  = _st.data[s.id] || {};
    const wd = d.word || {}, rd = d.reading || {};

    const retake = wd.retake !== '' && wd.retake != null ? wd.retake : '';
    const pass   = retake !== '' ? Math.max(0, totalWQ - Number(retake)) : '';
    const achW   = pass !== '' && totalWQ > 0 ? Math.round(pass / totalWQ * 100) : '';
    const isGW   = achW !== '' && achW >= 80;
    const achRd  = _calcRdStr(rd, actRevs);

    const rdRows = hasRd ? actRevs.map((rv, i) => {
      const key     = `R${i}`;
      const correct = rd[key]?.correct !== '' && rd[key]?.correct != null ? rd[key].correct : '';
      const score   = rd[key]?.score   !== '' && rd[key]?.score   != null ? rd[key].score   : '';
      return `<div class="gr-crow">
        <div class="gr-clbl" style="font-size:11px">${_e(rv.name)}</div>
        <div class="gr-cval">
          ${active
            ? `<input class="gr-cinp" type="number" min="0" max="${totalRQ}" step="1"
                     value="${correct}" placeholder="0" id="gr-cd-rd-${s.id}-${key}"
                     oninput="GradeApp._cardRdInput('${key}',this.value,'${s.id}','${totalRQ}')">`
            : `<div class="gr-cdisp score">${correct!==''?correct:'—'}</div>`}
          ${active ? `<span style="font-size:10px;color:var(--tx3);flex-shrink:0">/ ${totalRQ}</span>` : ''}
          <div class="gr-cdisp score" id="gr-cd-sc-${s.id}-${key}" style="min-width:44px">${score!==''?score:'—'}</div>
        </div>
      </div>`;
    }).join('') + `
      <div class="gr-crow">
        <div class="gr-clbl">성취율</div>
        <div class="gr-cval"><div class="gr-cdisp achv" id="gr-cd-achrd-${s.id}">${achRd||'—'}</div></div>
      </div>` : '';

    return `<div class="gr-card-body">
      <div class="gr-card-hero">
        <div class="gr-hero-emo">${_emoji(s, achW!==''?achW:null)}</div>
        <div>
          <div class="gr-hero-nm">${_e(s.name)}${s.nickname?` <span style="font-size:11px;color:var(--tx3)">(${_e(s.nickname)})</span>`:''}</div>
          <div class="gr-hero-sub">${_getCls(_st.classId)?.name||''}반</div>
        </div>
        ${achW!==''?`<div class="gr-hero-score">
          <div class="gr-hero-pct" style="color:${isGW?'#16a34a':'#f97316'}">${achW}%</div>
          <div class="gr-hero-lbl">단어 성취율</div>
        </div>`:''}
      </div>
      <!-- 단어 -->
      <div class="gr-csec">
        <div class="gr-csec-head">
          <div class="gr-csec-title">🔤 단어</div>
          <div class="gr-csec-badge">총 ${totalWQ}문제</div>
        </div>
        <div class="gr-card-grid">
          <div class="gr-crow">
            <div class="gr-clbl">재시험</div>
            <div class="gr-cval">
              ${active
                ? `<input class="gr-cinp" type="number" min="0" max="${totalWQ}" step="1"
                         value="${retake}" placeholder="0" id="gr-cd-retake-${s.id}"
                         oninput="GradeApp._cardWordInput(this.value,'${s.id}','${totalWQ}')">`
                : `<div class="gr-cdisp score">${retake!==''?retake:'—'}</div>`}
            </div>
          </div>
          <div class="gr-crow">
            <div class="gr-clbl">통과</div>
            <div class="gr-cval"><div class="gr-cdisp ${pass!==''?(isGW?'pass':'fail'):''}" id="gr-cd-pass-${s.id}">${pass!==''?pass:'—'}</div></div>
          </div>
          <div class="gr-crow">
            <div class="gr-clbl">성취율</div>
            <div class="gr-cval"><div class="gr-cdisp ${achW!==''?(isGW?'pass':'fail'):''}" id="gr-cd-achw-${s.id}">${achW!==''?achW+'%':'—'}</div></div>
          </div>
        </div>
      </div>
      <!-- 리딩 -->
      ${hasRd?`<div class="gr-csec">
        <div class="gr-csec-head">
          <div class="gr-csec-title">📖 리딩</div>
          <div class="gr-csec-badge rd">총 ${totalRQ}문제</div>
        </div>
        <div class="gr-card-grid">${rdRows}</div>
      </div>`:''}
      <!-- 코멘트 -->
      <div class="gr-csec">
        <div class="gr-csec-head"><div class="gr-csec-title">💬 Teacher's Comment</div></div>
        <div style="border:1px solid var(--bdr);border-top:none;border-radius:0 0 9px 9px">
          ${active
            ? `<textarea class="gr-card-cmt" id="gr-cd-cmt-${s.id}"
                         oninput="GradeApp._cardComment(this.value,'${s.id}')">${_e(d.comment||'')}</textarea>`
            : `<div style="padding:8px 10px;font-size:11px;color:var(--tx2);min-height:40px;line-height:1.7">${_e(d.comment||'')}</div>`}
        </div>
      </div>
      ${active?`<div class="gr-card-save-row">
        <button class="gr-card-save-btn" onclick="GradeApp.saveOne('${s.id}')">💾 저장</button>
      </div>`:''}
    </div>`;
  }

  /* 카드 입력 */
  function _cardWordInput(val, sid, totalWQ) {
    _ensureData(sid);
    const tq   = Number(totalWQ) || 0;
    const rt   = val===''?'':Math.max(0,Math.min(tq,Number(val)));
    const pass = rt!==''?Math.max(0,tq-rt):'';
    const achW = pass!==''&&tq>0?Math.round(pass/tq*100):'';
    const isGW = achW!==''&&achW>=80;
    _st.data[sid].word = {totalQ:tq,retake:rt,pass};
    _st.dirty.add(sid); _refreshDirtyUI();
    const pe=document.getElementById(`gr-cd-pass-${sid}`), ae=document.getElementById(`gr-cd-achw-${sid}`);
    if(pe){pe.textContent=pass!==''?pass:'—';pe.className=`gr-cdisp ${pass!==''?(isGW?'pass':'fail'):''}`;}
    if(ae){ae.textContent=achW!==''?achW+'%':'—';ae.className=`gr-cdisp ${achW!==''?(isGW?'pass':'fail'):''}`;}
  }

  function _cardRdInput(key, val, sid, totalRQ) {
    _ensureData(sid);
    const tq   = Number(totalRQ) || 0;
    const corr = val===''?'':Math.max(0,Math.min(tq,Number(val)));
    const sc   = corr!==''?Math.round(corr/tq*100*10)/10:'';
    if (!_st.data[sid].reading) _st.data[sid].reading = {};
    _st.data[sid].reading[key] = {correct:corr, score:sc};
    _st.dirty.add(sid); _refreshDirtyUI();
    const actRevs = GradeDB.getActiveReviews(_st.bookId);
    const scEl  = document.getElementById(`gr-cd-sc-${sid}-${key}`);
    const achEl = document.getElementById(`gr-cd-achrd-${sid}`);
    if(scEl)  scEl.textContent  = sc!==''?sc:'—';
    if(achEl) achEl.textContent = _calcRdStr(_st.data[sid].reading, actRevs)||'—';
  }

  function _cardComment(val, sid) {
    _ensureData(sid); _st.data[sid].comment=val;
    _st.dirty.add(sid); _refreshDirtyUI();
  }

  function _slideTo(i) {
    const sts=_getSorted();
    _st.slideIdx=Math.max(0,Math.min(i,sts.length-1));
    _st.studentId=sts[_st.slideIdx]?.id||null;
    _renderStudents(); _renderContent();
  }
  function _ts(e){_st.touchStartX=e.touches[0].clientX;}
  function _te(e){const dx=e.changedTouches[0].clientX-_st.touchStartX;if(Math.abs(dx)>50)_slideTo(_st.slideIdx+(dx<0?1:-1));}

  /* ════════════════════════════════════
   * 리포트 뷰
   * ════════════════════════════════════ */
  function _renderReportView(cnt, students) {
    const s = students.find(s=>s.id===_st.studentId)||students[0];
    if(!s){cnt.innerHTML=`<div class="gr-empty"><div class="gr-empty-ico">👆</div>좌측에서 학생을 선택하세요</div>`;return;}
    if(!_st.studentId){_st.studentId=s.id;_renderStudents();}
    cnt.innerHTML=`<div class="gr-report-panel">
      <div class="gr-rpt-cfg">
        <div class="gr-rpt-cfg-title">레이아웃</div>
        <div class="gr-rpt-layouts">${[1,2,3,4,5].map(n=>`<button class="gr-rpt-lbtn ${_st.reportLayout===n?'on':''}" onclick="GradeApp._setLayout(${n})">L${n}</button>`).join('')}</div>
        <label class="gr-rpt-toggle"><input type="checkbox" ${_st.reportGraph?'checked':''} onchange="GradeApp._toggleGraph(this.checked)"> 📊 그래프 포함</label>
      </div>
      <div class="gr-rpt-preview">
        <div class="rpt-wrap" id="gr-rpt-preview">${_buildReport(s)}</div>
        <div class="rpt-acts">
          <button class="rpt-btn copy"  onclick="GradeApp._copyReport()">📋 복사</button>
          <button class="rpt-btn share" onclick="GradeApp._shareReport()">📤 공유</button>
          <button class="rpt-btn pdf"   onclick="GradeApp._printReport()">🖨️ PDF</button>
          <button class="rpt-btn cap"   onclick="GradeApp._captureReport()">📸 캡처</button>
        </div>
      </div>
    </div>`;
  }

  function _buildReport(s) {
    const config=GradeDB.getReportConfig(_st.bookId),actRevs=GradeDB.getActiveReviews(_st.bookId);
    const hasRd=config.reading?.enabled&&actRevs.length>0;
    const book=typeof BookLibDB!=='undefined'?BookLibDB.getBookById(_st.bookId):null;
    const cls=_getCls(_st.classId);
    const rec=GradeDB.getLatest(_st.classId,s.id,_st.bookId);
    const students=_getStudents();
    const today=new Date().toLocaleDateString('ko-KR');

    const achWs=students.map(st=>{const r=GradeDB.getLatest(_st.classId,st.id,_st.bookId);return r?.word?.totalQ>0?Math.round(r.word.pass/r.word.totalQ*100):null;}).filter(v=>v!=null);
    const avgW=achWs.length?Math.round(achWs.reduce((a,b)=>a+b,0)/achWs.length):null;
    const achRds=hasRd?students.map(st=>{const r=GradeDB.getLatest(_st.classId,st.id,_st.bookId);return _calcRdN(r?.reading||{},actRevs);}).filter(v=>v!=null):[];
    const avgRd=achRds.length?Math.round(achRds.reduce((a,b)=>a+b,0)/achRds.length):null;
    const achW=rec?.word?.totalQ>0?Math.round(rec.word.pass/rec.word.totalQ*100):null;
    const achRd=hasRd?_calcRdN(rec?.reading||{},actRevs):null;
    const isGW=achW!=null&&achW>=80;

    const logoSrc=typeof LOGO!=='undefined'&&LOGO.small?LOGO.small:'';
    const hdr=`<div class="rpt-header">${logoSrc?`<img src="${logoSrc}" style="width:44px;height:44px;object-fit:contain;border-radius:8px">`:'<div style="font-size:36px">🌳</div>'}<div style="flex:1"><div class="rpt-title">Achievement Report</div></div><div style="font-size:10px;color:#999">${today}</div></div><hr class="rpt-divider">`;
    const info=`<div class="rpt-info"><p><strong>Student :</strong> ${_e(s.name)}${s.nickname?`(${_e(s.nickname)})`:''}</p><p><strong>Book :</strong>&nbsp;&nbsp;&nbsp;&nbsp;${_e(book?.name||'')}</p></div>`;
    const wordTbl=rec?`<div class="rpt-sec-title">단어 Test Result</div><table class="rpt-tbl"><thead><tr><th>총 테스트수</th><th style="color:#4f46e5">통과</th><th style="color:#ea580c">재시</th><th style="color:#8b5cf6">성취율</th></tr></thead><tbody><tr><td>${rec.word?.totalQ??'—'}</td><td class="rpt-pass">${rec.word?.pass??'—'}</td><td class="rpt-fail">${rec.word?.retake??'—'}</td><td class="rpt-achv">${achW!=null?achW+'%':'—'}</td></tr><tr class="rpt-avg"><td colspan="3" style="text-align:center">평균</td><td class="rpt-achv">${avgW!=null?avgW+'%':'—'}</td></tr></tbody></table>`:'';
    const rdTbl=hasRd&&rec?`<div class="rpt-sec-title">리딩 Test Result</div><table class="rpt-tbl"><thead><tr>${actRevs.map(rv=>`<th>${_e(rv.name)}</th>`).join('')}<th style="color:#8b5cf6">성취율</th></tr></thead><tbody><tr>${actRevs.map((_,i)=>{const sc=rec.reading?.[`R${i}`]?.score;return`<td class="rpt-pass">${sc!=null?sc+'점':'—'}</td>`;}).join('')}<td class="rpt-achv">${achRd!=null?Math.round(achRd)+'%':'—'}</td></tr><tr class="rpt-avg"><td colspan="${actRevs.length}" style="text-align:center">평균</td><td class="rpt-achv">${avgRd!=null?avgRd+'%':'—'}</td></tr></tbody></table>`:'';
    const graph=_st.reportGraph&&achW!=null?`<div class="rpt-graph-wrap">${_svgGraph(achW,hasRd?achRd:null)}</div>`:'';
    const comment=`<div style="margin-top:14px"><div class="rpt-sec-title">Teacher's comment</div><div class="rpt-comment-box">${_e(rec?.comment||'')}</div></div>`;
    const L=_st.reportLayout;
    const bodies={1:[hdr,info,wordTbl,rdTbl,graph,comment],2:[hdr,info,graph,wordTbl,rdTbl,comment],3:[hdr,info,wordTbl,comment,rdTbl,graph],4:[hdr,info,rdTbl,wordTbl,graph,comment],5:[hdr,info,graph,comment,wordTbl,rdTbl]};
    return(bodies[L]||bodies[1]).filter(Boolean).join('');
  }

  function _svgGraph(achW, achRd) {
    const bars=[['단어',achW,'#6366f1']];if(achRd!=null)bars.push(['리딩',Math.round(achRd),'#8b5cf6']);
    const bW=36,gap=20,tw=bars.length*(bW+gap)+gap;
    const svgBars=bars.map(([l,p,c],i)=>{const h=Math.round(p*0.55);const x=i*(bW+gap)+gap;return`<rect x="${x}" y="${55-h}" width="${bW}" height="${h}" rx="4" fill="${c}" opacity=".85"/><text x="${x+bW/2}" y="62" text-anchor="middle" font-size="9" fill="#555">${l}</text><text x="${x+bW/2}" y="${55-h-3}" text-anchor="middle" font-size="10" font-weight="bold" fill="${c}">${p}%</text>`;}).join('');
    return`<svg width="${tw}" height="66" style="display:block;margin:0 auto">${svgBars}<line x1="0" y1="55" x2="${tw}" y2="55" stroke="#e2e8f0" stroke-width="1"/></svg>`;
  }

  function _setLayout(n){_st.reportLayout=n;const s=_getStudents().find(s=>s.id===_st.studentId)||_getStudents()[0];if(s){const el=document.getElementById('gr-rpt-preview');if(el)el.innerHTML=_buildReport(s);}document.querySelectorAll('.gr-rpt-lbtn').forEach((b,i)=>b.classList.toggle('on',i+1===n));}
  function _toggleGraph(v){_st.reportGraph=v;_setLayout(_st.reportLayout);}
  async function _copyReport(){const el=document.getElementById('gr-rpt-preview');try{await navigator.clipboard.writeText(el?.innerText||'');_toast('📋 복사됐습니다','success');}catch{_toast('⚠️ 복사 실패');}}
  async function _shareReport(){const text=document.getElementById('gr-rpt-preview')?.innerText||'';const sd={title:'Achievement Report',text};if(navigator.share&&navigator.canShare?.(sd)){try{await navigator.share(sd);_toast('📤 공유 완료','success');return;}catch(e){if(e.name==='AbortError')return;}}_copyReport();}
  function _printReport(){const el=document.getElementById('gr-rpt-preview');if(!el)return;let frame=document.getElementById('gr-pf');if(!frame){frame=document.createElement('div');frame.id='gr-pf';document.body.appendChild(frame);}frame.innerHTML=el.innerHTML;window.print();setTimeout(()=>frame.remove(),1500);}
  async function _captureReport(){const el=document.getElementById('gr-rpt-preview');if(!el)return;if(typeof html2canvas!=='undefined'){const c=await html2canvas(el,{scale:2,backgroundColor:'#fff'});const a=document.createElement('a');a.href=c.toDataURL('image/png');a.download='report.png';a.click();_toast('📸 캡처 완료','success');}else _toast('⚠️ html2canvas 라이브러리가 필요합니다');}

  /* ════════════════════════════════════
   * 저장
   * ════════════════════════════════════ */
  async function saveOne(sid) {
    if (!_st.classId || !_st.bookId) { _toast('⚠️ 반과 교재를 선택해주세요'); return; }
    _ensureData(sid);

    /* 엑셀 모드: DOM 최신값 수집 */
    if (_st.viewMode === 'excel') {
      const cfg = GradeDB.getReportConfig(_st.bookId);
      const tq  = cfg.word?.totalQ || 0;
      const retEl = document.getElementById(`gr-retake-${sid}`);
      if (retEl) {
        const rt = retEl.value === '' ? 0 : Number(retEl.value);
        _st.data[sid].word = { totalQ:tq, retake:rt, pass:Math.max(0,tq-rt) };
      }
      const cmtEl = document.getElementById(`gr-cmt-${sid}`);
      if (cmtEl) _st.data[sid].comment = cmtEl.value;
      /* 리딩 */
      const actRevs = GradeDB.getActiveReviews(_st.bookId);
      const totalRQ = cfg.reading?.totalQ || 0;
      actRevs.forEach((_, i) => {
        const key = `R${i}`;
        const inp = document.querySelector(`[data-sid="${sid}"][data-rkey="${key}"]`);
        if (inp) {
          const corr  = inp.value === '' ? '' : Math.max(0, Math.min(totalRQ, Number(inp.value)));
          const score = corr !== '' ? Math.round(corr / totalRQ * 100 * 10) / 10 : '';
          if (!_st.data[sid].reading) _st.data[sid].reading = {};
          _st.data[sid].reading[key] = { correct:corr, score };
        }
      });
    }

    const now     = new Date();
    const savedAt = `${now.toISOString().slice(0,10)} ${now.toTimeString().slice(0,5)}`;
    _st.data[sid].savedAt = savedAt;

    await GradeDB.saveRecord({
      classId:   _st.classId,
      studentId: sid,
      bookId:    _st.bookId,
      word:      _st.data[sid].word    || null,
      reading:   _st.data[sid].reading || null,
      comment:   _st.data[sid].comment || '',
      savedAt,
    });

    _st.dirty.delete(sid);
    _refreshDirtyUI();
    /* savedAt 표시 갱신 */
    const fixCell = document.querySelector(`#gr-row-${sid} .gs-fix`);
    if (fixCell) {
      const savedEl = fixCell.querySelector('[data-savedat]') || (() => {
        const d = document.createElement('div');
        d.dataset.savedat = '1';
        d.style.cssText = 'font-size:8px;color:var(--tx3)';
        fixCell.appendChild(d); return d;
      })();
      savedEl.textContent = savedAt.slice(5,16);
    }
    _toast('✅ 저장 완료', 'success');
    _renderStudents();
  }

  async function saveAll() {
    const sids = _getSorted().map(s => s.id);
    if (!sids.length) return;
    let ok = 0;
    for (const sid of sids) {
      _ensureData(sid);
      try { await saveOne(sid); ok++; } catch(e) { console.error('[saveAll]', sid, e); }
    }
    _toast(`✅ ${ok}명 저장 완료`, 'success');
    _renderContent();
  }

  async function resetOne(sid) {
    if (!confirm('이 학생의 입력 데이터를 초기화하시겠습니까?')) return;
    const records = GradeDB.getRecords(_st.classId, sid, _st.bookId);
    for (const r of records) await GradeDB.deleteRecord(_st.classId, sid, _st.bookId, r.id);
    delete _st.data[sid]; _st.dirty.delete(sid);
    _refreshDirtyUI(); _renderContent(); _renderStudents();
    _toast('🗑 초기화 완료');
  }

  /* ══ 선택 핸들러 ══ */
  function _onCls(clsId) {
    if (_st.dirty.size > 0 && !confirm('저장하지 않은 성적이 있습니다. 계속하시겠습니까?')) {
      document.getElementById('gr-csel').value = _st.classId||''; return;
    }
    _st.classId=clsId||null; _st.bookId=null; _st.studentId=null; _st.data={}; _st.dirty.clear(); _st.sortCol=null;
    _fillBooks(); _renderStudents(); _renderContent(); _updateRptBtn(); _updateSub();
    const bsel=document.getElementById('gr-bsel'); if(bsel)bsel.disabled=!_st.classId;
  }
  function _onBk(bkId) {
    if (_st.dirty.size > 0 && !confirm('저장하지 않은 성적이 있습니다. 계속하시겠습니까?')) {
      document.getElementById('gr-bsel').value = _st.bookId||''; return;
    }
    _st.bookId=bkId||null; _st.studentId=null; _st.data={}; _st.dirty.clear(); _st.sortCol=null;
    _renderStudents(); _renderContent(); _updateRptBtn(); _updateSub();
  }
  function _onStu(sid, idx) {
    _st.studentId=sid||null; _st.slideIdx=idx??0;
    _renderStudents();
    if (_st.viewMode==='excel') {
      document.querySelectorAll('.gr-sheet tbody tr').forEach(tr=>tr.classList.toggle('sel-row',tr.id===`gr-row-${sid}`));
      document.querySelectorAll('.gs-fix').forEach(td=>{const m=td.getAttribute('onclick')?.match(/'([^']+)'/);td.classList.toggle('sel',m?.[1]===sid);});
      document.getElementById(`gr-row-${sid}`)?.scrollIntoView({behavior:'smooth',block:'nearest'});
      _updateChart();
    } else if (_st.viewMode==='card') {
      _renderContent();
    } else if (_st.viewMode==='report') {
      const el=document.getElementById('gr-rpt-preview');
      if(el&&sid){const s=_getStudents().find(s=>s.id===sid);if(s)el.innerHTML=_buildReport(s);}
    }
  }
  function _setView(mode) {
    if(_st.dirty.size>0&&mode!==_st.viewMode&&!confirm('저장하지 않은 성적이 있습니다. 이동하시겠습니까?'))return;
    _st.viewMode=mode; _st.data={};
    document.querySelectorAll('.gr-vbtn').forEach((b,i)=>b.classList.toggle('on',(['excel','card','report'][i])===mode));
    _renderContent();
  }
  function _updateRptBtn(){const btn=document.getElementById('gr-rpt-btn');if(btn)btn.style.display=(_st.classId&&_st.bookId)?'':'none';}
  function _updateSub(){const sub=document.getElementById('gr-sub');if(!sub)return;const cls=_st.classId?_getCls(_st.classId):null;const bk=_st.bookId&&typeof BookLibDB!=='undefined'?BookLibDB.getBookById(_st.bookId):null;sub.textContent=cls&&bk?`${cls.name}반 · ${bk.name}`:cls?`${cls.name}반`:'반 · 교재를 선택하세요';}
  function _refreshDirtyUI(){const el=document.getElementById('gr-dirty-cnt');if(el)el.textContent=_st.dirty.size?`(${_st.dirty.size})`:'';document.querySelectorAll('.gr-stu-item').forEach(item=>{const m=item.getAttribute('onclick')?.match(/'([^']+)'/);if(m)item.classList.toggle('dirty-item',_st.dirty.has(m[1]));});}

  /* ══ 전체 성적표 ══ */
  function openReport() {
    if(!_st.classId||!_st.bookId){_toast('⚠️ 반과 교재를 선택해주세요');return;}
    const ov=document.getElementById('gr-rpt-ov'),sh=document.getElementById('gr-rpt-sh');if(!ov||!sh)return;
    const cls=_getCls(_st.classId),book=typeof BookLibDB!=='undefined'?BookLibDB.getBookById(_st.bookId):null;
    const config=GradeDB.getReportConfig(_st.bookId),actRevs=GradeDB.getActiveReviews(_st.bookId);
    const hasRd=config.reading?.enabled&&actRevs.length>0;
    const sts=_getSorted(),today=new Date().toLocaleDateString('ko-KR');
    let txt=`📝 ${book?.name||''} 성적표\n🏫 ${cls?.name||''}반 · ${today}\n${'─'.repeat(26)}\n`;
    sts.forEach(s=>{const rec=GradeDB.getLatest(_st.classId,s.id,_st.bookId);if(!rec){txt+=`${s.name}: 미입력\n`;return;}const achW=rec.word?.totalQ>0?Math.round(rec.word.pass/rec.word.totalQ*100):null;const achRd=hasRd?_calcRdN(rec.reading||{},actRevs):null;txt+=`${s.name}${s.nickname?'('+s.nickname+')':''}: 단어 ${rec.word?.pass??'—'}/${rec.word?.totalQ??'—'}${achW!=null?'('+achW+'%)':''}`+(achRd!=null?` · 리딩 ${Math.round(achRd)}%`:'')+'\n';});
    const thRd=hasRd?actRevs.map(rv=>`<th style="border:1px solid var(--bdr);padding:6px">${_e(rv.name)}</th>`).join('')+'<th style="border:1px solid var(--bdr);padding:6px">성취율</th>':'';
    const tdRows=sts.map(s=>{const r=GradeDB.getLatest(_st.classId,s.id,_st.bookId);if(!r)return`<tr><td style="border:1px solid var(--bdr);padding:6px;font-weight:700">${_e(s.name)}</td><td colspan="99" style="border:1px solid var(--bdr);padding:6px;color:var(--tx3)">미입력</td></tr>`;const achW=r.word?.totalQ>0?Math.round(r.word.pass/r.word.totalQ*100):null;const isGW=achW!=null&&achW>=80;const achRd=hasRd?_calcRdN(r.reading||{},actRevs):null;const rdTds=hasRd?actRevs.map((_,i)=>{const sc=r.reading?.[`R${i}`]?.score??'—';return`<td style="border:1px solid var(--bdr);padding:6px;color:var(--a)">${sc}</td>`;}).join('')+`<td style="border:1px solid var(--bdr);padding:6px;color:#8b5cf6;font-weight:700">${achRd!=null?Math.round(achRd)+'%':'—'}</td>`:'';return`<tr><td style="border:1px solid var(--bdr);padding:6px;font-weight:700">${_e(s.name)}${s.nickname?` (${_e(s.nickname)})`':''}</td><td style="border:1px solid var(--bdr);padding:6px;color:${isGW?'#16a34a':'#f97316'};font-weight:700">${r.word?.pass??'—'}</td><td style="border:1px solid var(--bdr);padding:6px">${r.word?.retake??'—'}</td><td style="border:1px solid var(--bdr);padding:6px;color:${isGW?'#16a34a':'#f97316'};font-weight:800">${achW!=null?achW+'%':'—'}</td>${rdTds}<td style="border:1px solid var(--bdr);padding:6px;font-size:11px">${_e(r.comment||'')}</td></tr>`;}).join('');
    sh.innerHTML=`<div class="sh-handle"></div><div class="sh-title">📋 반 전체 성적표</div><div class="sh-sub">${_e(cls?.name||'')}반 · ${_e(book?.name||'')} · ${today}</div><div class="gr-rpt-sh-scroll"><div style="overflow-x:auto;padding:10px 2px"><table style="width:100%;border-collapse:collapse;font-size:12px;min-width:max-content"><thead><tr style="background:var(--surf2)"><th style="border:1px solid var(--bdr);padding:7px 10px;font-size:11px;color:var(--tx3)">학생</th><th style="border:1px solid var(--bdr);padding:7px 10px;font-size:11px;color:var(--tx3)">통과</th><th style="border:1px solid var(--bdr);padding:7px 10px;font-size:11px;color:var(--tx3)">재시험</th><th style="border:1px solid var(--bdr);padding:7px 10px;font-size:11px;color:var(--tx3)">성취율</th>${thRd}<th style="border:1px solid var(--bdr);padding:7px 10px;font-size:11px;color:var(--tx3)">코멘트</th></tr></thead><tbody>${tdRows}</tbody></table></div><div style="font-size:10px;font-weight:800;color:var(--tx3);letter-spacing:1px;margin:10px 14px 4px">공유용 텍스트</div><div class="gr-share-box" style="margin:0 14px 8px">${_e(txt)}</div></div><div class="gr-sacts" style="padding:0 0 8px"><button class="gr-sbtn copy" onclick="GradeApp._copy(${JSON.stringify(txt)})">📋 복사</button><button class="gr-sbtn share" onclick="GradeApp._shr(${JSON.stringify(txt)})">📤 공유</button></div><button class="btn-x" style="width:100%" onclick="GradeApp.closeReport()">닫기</button>`;
    ov.classList.remove('hidden');history.pushState({pg:'grade',modal:'report'},'');
  }
  async function _copy(t){try{await navigator.clipboard.writeText(t);_toast('📋 복사됐습니다','success');}catch{_toast('⚠️ 복사 실패');}}
  async function _shr(t){const sd={title:'성적표',text:t};if(navigator.share&&navigator.canShare?.(sd)){try{await navigator.share(sd);_toast('📤 공유 완료','success');return;}catch(e){if(e.name==='AbortError')return;}}_copy(t);}
  function closeReport(){document.getElementById('gr-rpt-ov')?.classList.add('hidden');}

  /* ══ 유틸 ══ */
  function _getStudents(){const cls=_st.classId?_getCls(_st.classId):null;if(!cls||typeof StudentDB==='undefined')return[];return StudentDB.getFiltered({classCode:cls.name,status:'재원'});}
  function _ensureData(sid){if(!_st.data[sid]){const cfg=GradeDB.getReportConfig(_st.bookId);const tq=cfg.word?.totalQ||0;const rec=GradeDB.getLatest(_st.classId,sid,_st.bookId);_st.data[sid]=rec?JSON.parse(JSON.stringify(rec)):{word:{totalQ:tq,retake:'',pass:''},reading:{},comment:''};_st.data[sid].comment=_st.data[sid].comment||'';}}
  function _givN(n){return n?.length>1?n.slice(1):n||'';}
  function _getCls(id){if(typeof DB==='undefined')return null;if(typeof DB.getClassById==='function')return DB.getClassById(id);return(DB.getActiveClasses?.()||[]).find(c=>c.id===id)||null;}
  function _calcRdN(rd,revs){if(!revs?.length)return null;const sc=revs.map((_,i)=>rd[`R${i}`]?.score).filter(s=>s!=null&&s!=='');if(!sc.length)return null;return Math.round(sc.reduce((a,b)=>a+b,0)/sc.length);}
  function _calcRdStr(rd,revs){const n=_calcRdN(rd,revs);return n!=null?n+'%':null;}
  const _e=v=>String(v??'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
  function _toast(msg,type){const el=document.getElementById('toast');if(!el)return;el.textContent=msg;el.className=type==='success'?'success':'';el.classList.remove('hidden');clearTimeout(el._t);el._t=setTimeout(()=>el.classList.add('hidden'),3000);}

  return {
    init, render,
    _onCls, _onBk, _onStu, _setView, _toggleSort,
    _excelWordInput, _excelRdInput, _excelComment, _onKey,
    _cardWordInput, _cardRdInput, _cardComment,
    _slideTo, _ts, _te,
    _onCtxTable, _closeCtxMenu,
    saveOne, saveAll, resetOne,
    _setLayout, _toggleGraph,
    _copyReport, _shareReport, _printReport, _captureReport,
    openReport, closeReport, _copy, _shr,
  };
})();
