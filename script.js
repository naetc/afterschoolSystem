'use strict';

const $  = id => document.getElementById(id);
const val = id => $(id)?.value.trim() || '';
const num = v  => parseInt(String(v).replace(/,/g,'')) || 0;
const fmt = n  => Number(n||0).toLocaleString();
const uid = (g,b,n,nm) => `${+g||0}-${+b||0}-${+n||0}-${String(nm||'').trim()}`;
const dsp = (g,b,n)    => `${+g||0}-${+b||0}-${+n||0}`;

function readFileAsArrayBuffer(file) { return new Promise((r, j) => { const rd = new FileReader(); rd.onload = e => r(e.target.result); rd.onerror = () => j(new Error('파일 읽기 실패')); rd.readAsArrayBuffer(file); }); }
function readFileAsText(file) { return new Promise((r, j) => { const rd = new FileReader(); rd.onload = e => r(e.target.result); rd.onerror = () => j(new Error('파일 읽기 실패')); rd.readAsText(file, 'utf-8'); }); }
function parseXlsx(buffer) { const wb = XLSX.read(new Uint8Array(buffer), {type:'array'}); return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {defval:''}); }

// [무결성 변수군]
let C={}, M={}, F=[], E=[], Ld={}, Hs=[];
let f_eq='1', f_ec='ALL'; 
let s4_filt='A', curS4Tab='STAT', s4_cFilter='ALL';
let stuSortCol = '', stuSortAsc = false; 
let mdlConsole, mdlCrsSummary, mdlFreeStart; 
let SysSet = { closedSess: {} };
const KEY = 'bgh_v36';

window.gQ = 1;
let cUid = '', cEnrolls = [], cActiveEIdx = -1, curEditFreeIdx = -1;

window.isQuarterLocked = function(q) { return Object.keys(SysSet.closedSess || {}).some(k => k.startsWith(q + '_')); };

function loadData() {
    const raw = localStorage.getItem(KEY); if (!raw) return;
    try {
        const d = JSON.parse(raw); C = d.C || {}; M = d.M || {}; 
        SysSet = d.SysSet || { closedSess: {} }; SysSet.closedSess = SysSet.closedSess || {};
        F = (d.F || []).map(x => ({ g: +(x.g||0), b: +(x.b||0), n: +(x.n||0), name: String(x.name||''), startQ: +(x.startQ||1), startSess: +(x.startSess||0) }));
        E = (d.E || []).map(x => ({ q: +(x.q||1), g: +(x.g||0), b: +(x.b||0), n: +(x.n||0), name: String(x.name||''), course: String(x.course||''), cT: (x.cT != null) ? +x.cT : null, cB: (x.cB != null) ? +x.cB : null, rT: +(x.rT||0), rB: +(x.rB||0), mm: String(x.mm||''), tMemo: String(x.tMemo||''), bMemo: String(x.bMemo||''), refunds: x.refunds || [], adjusts: x.adjusts || [] }));
        Object.keys(M).forEach(dept => { if (M[dept].cnt !== undefined) { const old = M[dept]; M[dept] = {1:{...old}, 2:{...old}, 3:{...old}, 4:{...old}}; } });
    } catch(e) { console.error('로딩 오류', e); }
}

const save = () => { try { localStorage.setItem(KEY, JSON.stringify({C, M, F, E, SysSet})); updateStorageUsage(); } catch(e) { alert('용량 한도 초과'); } };

window.maskNames = function() {
    if(!confirm('전체 학생 이름 마스킹 처리하시겠습니까?')) return;
    F.forEach(s => s.name = maskWord(s.name)); E.forEach(s => s.name = maskWord(s.name)); save(); location.reload();
};
function maskWord(nm) { if (!nm) return nm; if (nm.length === 2) return nm[0]+'*'; return nm[0]+'*'.repeat(nm.length-2)+nm.slice(-1); }

window.sysBackup = function() {
    const blob = new Blob([JSON.stringify({C, M, F, E, SysSet})], {type:'application/json'});
    const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = `방과후정산_백업_${new Date().toISOString().slice(0,10)}.json`; a.click();
};

function updateStorageUsage() {
    const el = $('storageUsage'); if (!el) return; const raw = localStorage.getItem(KEY) || '';
    const pct = ((raw.length * 2) / (5 * 1024 * 1024) * 100).toFixed(1);
    el.innerHTML = `💾 로컬: <span class="text-success">${((raw.length * 2) / 1024).toFixed(1)} KB</span> / 5120 KB (${pct}%)`;
}

// 💡 1분기 기준 가상 시나리오 테스트 데이터 세팅 함수
window.injectTestData = function() {
    if(!confirm("⚠️ 주의!\n기존 데이터가 모두 삭제되고 다양한 시나리오가 반영된 가상의 [테스트 데이터 세트]로 덮어씌워집니다.\n(필요시 미리 '데이터 백업'을 권장합니다.)\n진행하시겠습니까?")) return;

    // 1. 가상 마스터(부서) 생성
    M = {
        "로봇과학": {
            1: {cnt:1, inst_m:30000, mgmt_m:2000, b:45000, unit:1, mh:"4,4,4"},
            2: {cnt:1, inst_m:30000, mgmt_m:2000, b:45000, unit:1, mh:"4,4,4"},
            3: {cnt:1, inst_m:30000, mgmt_m:2000, b:45000, unit:1, mh:"4,4,4"},
            4: {cnt:1, inst_m:30000, mgmt_m:2000, b:45000, unit:1, mh:"4,4,4"}
        },
        "원어민영어": {
            1: {cnt:2, inst_m:35000, mgmt_m:2000, b:15000, unit:2, mh:"8,8,8"},
            2: {cnt:2, inst_m:35000, mgmt_m:2000, b:15000, unit:2, mh:"8,8,8"},
            3: {cnt:2, inst_m:35000, mgmt_m:2000, b:15000, unit:2, mh:"8,8,8"},
            4: {cnt:2, inst_m:35000, mgmt_m:2000, b:15000, unit:2, mh:"8,8,8"}
        },
        "방송댄스": {
            1: {cnt:1, inst_m:28000, mgmt_m:1500, b:0, unit:1, mh:"4,4,4"},
            2: {cnt:1, inst_m:28000, mgmt_m:1500, b:0, unit:1, mh:"4,4,4"},
            3: {cnt:1, inst_m:28000, mgmt_m:1500, b:0, unit:1, mh:"4,4,4"},
            4: {cnt:1, inst_m:28000, mgmt_m:1500, b:0, unit:1, mh:"4,4,4"}
        }
    };
    window.regenerateC();

    // 2. 가상 지원 대상자 (초3 혼합, 시작시점 교차)
    F = [
        {g:1, b:1, n:1, name:'김자유', startQ:1, startSess:0}, 
        {g:3, b:2, n:5, name:'이무상', startQ:1, startSess:0}, 
        {g:5, b:3, n:12, name:'박지원', startQ:1, startSess:1}  // 1분기 2차수부터 지원
    ];

    // 3. 가상 수강생 (할인, 환불, 중도포기 등 포함)
    E = [];
    const rawE = [
        {q:1, g:1, b:1, n:1, name:'김자유', course:'로봇과학', adj:[], ref:[]}, 
        {q:1, g:1, b:1, n:1, name:'김자유', course:'원어민영어(A)', adj:[], ref:[]},
        {q:1, g:3, b:2, n:5, name:'이무상', course:'방송댄스', adj:[{title:'교재비면제', amtT:0, amtB:-10000}], ref:[]}, 
        {q:1, g:3, b:1, n:15, name:'홍길동', course:'로봇과학', adj:[{title:'다자녀할인', amtT:-10000, amtB:0}], ref:[]}, 
        {q:1, g:5, b:3, n:12, name:'박지원', course:'원어민영어(B)', adj:[], ref:[{sessIdx:0, ty:'DISEASE', ah:2, reqBk:false, rt:0, rb:0, tyNm:''}]}, 
        {q:1, g:6, b:1, n:9, name:'최포기', course:'로봇과학', adj:[], ref:[{sessIdx:1, ty:'STUDENT', ah:1, reqBk:true, rt:0, rb:0, tyNm:''}]} 
    ];

    rawE.forEach(s => {
        let ne = {q:s.q, g:s.g, b:s.b, n:s.n, name:s.name, course:s.course, cT:null, cB:null, rT:0, rB:0, mm:'', tMemo:'', bMemo:'', refunds:s.ref, adjusts:s.adj};
        window.recalcEnrollment(ne);
        E.push(ne);
    });

    SysSet = { closedSess: {} }; 
    save();
    alert('✅ 샘플 데이터가 성공적으로 생성되었습니다!\n1분기 기준으로 세팅된 다자녀할인, 결석, 중도포기, 시작시점 교차 테스트를 지금 바로 경험해 보세요.');
    location.reload();
};

window.setQTab = function(q) {
    window.gQ = q; 
    [1,2,3,4].forEach(i => $('btnQTab'+i)?.classList.remove('active'));
    $('btnQTab'+q)?.classList.add('active');
    
    ['exEnQ', 's4_q', 'p_q1', 'p_qInvoice', 'p_q2'].forEach(id => { if($(id)) $(id).value = String(q); });
    f_eq = String(q);
    
    if($('lblMasterTab')) $('lblMasterTab').innerHTML = `<i class="bi bi-building"></i> [${q}분기] 부서 마스터`;
    if($('lblCourseTab')) $('lblCourseTab').innerHTML = `<i class="bi bi-list-check"></i> [${q}분기] 강좌 요금표`;
    
    regenerateC(); 
    renderE();
    autoRunSet(true); 
    renderSetTabs();
};

window.addEventListener('DOMContentLoaded', () => {
    loadData(); updateStorageUsage();
    if($('restoreFile')) {
        $('restoreFile').addEventListener('change', async function() {
            const file = this.files[0]; if (!file) return; if (!confirm('기존 장부를 덮어씁니다. 진행?')) return;
            try {
                const text = await readFileAsText(file); const d = JSON.parse(text);
                C=d.C||{}; M=d.M||{}; SysSet=d.SysSet||{closedSess:{}}; SysSet.closedSess=SysSet.closedSess||{};
                F=(d.F||[]).map(x=>({g:+(x.g??0), b:+(x.b??0), n:+(x.n??0), name:String(x.name||''), startQ: +(x.startQ||1), startSess: +(x.startSess||0) }));
                E=(d.E||[]).map(x=>({q:+(x.q||1), g:+(x.g??0), b:+(x.b??0), n:+(x.n??0), name:String(x.name||''), course:String(x.course||''), cT:(x.cT!=null)?+x.cT:null, cB:(x.cB!=null)?+x.cB:null, rT:+(x.rT||0), rB:+(x.rB||0), mm:String(x.mm||''), tMemo:String(x.tMemo||''), bMemo:String(x.bMemo||''), refunds:x.refunds||[], adjusts:x.adjusts||[]}));
                Object.keys(M).forEach(dept => { if (M[dept].cnt !== undefined) { const old = M[dept]; M[dept] = {1:{...old}, 2:{...old}, 3:{...old}, 4:{...old}}; } });
                save(); location.reload();
            } catch(err) { alert('❌ 데이터 파손'); }
        });
    }
    
    if($('mdlStuConsole') && typeof bootstrap !== 'undefined') mdlConsole = new bootstrap.Modal($('mdlStuConsole'));
    if($('mdlCourseSummary') && typeof bootstrap !== 'undefined') mdlCrsSummary = new bootstrap.Modal($('mdlCourseSummary'));
    if($('mdlFreeStart') && typeof bootstrap !== 'undefined') mdlFreeStart = new bootstrap.Modal($('mdlFreeStart'));

    const crsHead = document.querySelector('#mdlCourseSummary thead');
    if(crsHead) { crsHead.classList.remove('table-dark'); crsHead.classList.add('text-dark'); }

    if($('tabStep4Btn')) { $('tabStep4Btn').addEventListener('shown.bs.tab', () => { autoRunSet(false); }); }

    E.forEach(e => recalcEnrollment(e)); save();
    setQTab(1); renderF(); renderE();
    renderStaticHeaders();
});

function renderStaticHeaders() {
    if($('tbStatHead')) $('tbStatHead').innerHTML = `<tr><th rowspan="2">강좌명</th><th rowspan="2">신청인원</th><th colspan="2" class="table-warning">실부담금 총액</th><th colspan="2" class="bg-cho3">초3 공제합계</th><th colspan="2" class="bg-free">자유수강 공제합계</th><th colspan="2" class="table-danger">최종 징수액(자부담)</th></tr><tr><th class="table-warning">수강료계</th><th class="table-warning">교재비계</th><th class="bg-cho3">수강료</th><th class="bg-cho3">교재비</th><th class="bg-free">수강료</th><th class="bg-free">교재비</th><th class="table-danger">수강료합</th><th class="table-danger">교재비합</th></tr>`;
    if($('tbStuDtlHead')) $('tbStuDtlHead').innerHTML = `<tr><th rowspan="2" class="clickable" onclick="sortStu('DP')">학적 <span id="sort_DP"></span></th><th rowspan="2" class="clickable" onclick="sortStu('NM')">이름 <span id="sort_NM"></span></th><th rowspan="2">대상</th><th colspan="2">지원금 잔여</th><th rowspan="2">분기</th><th rowspan="2">강좌명</th><th colspan="2">실부담금</th><th colspan="2" class="bg-cho3">초3 공제</th><th colspan="2" class="bg-free">자유 공제</th><th colspan="2" class="text-danger fw-bold">최종징수(자부담)</th></tr><tr><th class="clickable text-primary" onclick="sortStu('C')">초3잔액 <span id="sort_C"></span></th><th class="clickable text-success" onclick="sortStu('F')">자유잔액 <span id="sort_F"></span></th><th>수강료</th><th>교재비</th><th class="bg-cho3">수강료</th><th class="bg-cho3">교재비</th><th class="bg-free">수강료</th><th class="bg-free">교재비</th><th class="text-danger">수강료</th><th class="text-danger">교재비</th></tr>`;
    if($('tbCrseDtlHead')) $('tbCrseDtlHead').innerHTML = `<tr><th rowspan="2">분기</th><th rowspan="2">학적</th><th rowspan="2">이름</th><th rowspan="2">강좌명</th><th colspan="2">실부담금</th><th colspan="2" class="bg-cho3">초3 공제</th><th colspan="2" class="bg-free">자유 공제</th><th colspan="2" class="text-danger fw-bold">최종징수(자부담)</th></tr><tr><th>수강료</th><th>교재비</th><th class="bg-cho3">수강료</th><th class="bg-cho3">교재비</th><th class="bg-free">수강료</th><th class="bg-free">교재비</th><th class="text-danger">수강료</th><th class="text-danger">교재비</th></tr>`;
    if($('tbInvHead')) $('tbInvHead').innerHTML = `<tr><th rowspan="2">순번</th><th rowspan="2">강좌명</th><th colspan="3">1인당 기준액</th><th colspan="4" class="table-warning">수익자 부담 (자부담)</th><th colspan="4" class="bg-cho3">초3 지원</th><th colspan="4" class="bg-free">자유수강 지원</th><th colspan="4" class="table-danger">합계 (최종집행액)</th><th rowspan="2" class="table-secondary">차액<br>(환불/조정)</th><th rowspan="2">비고(수강료)</th></tr><tr><th>수강료</th><th>강사료</th><th>수용비</th><th class="table-warning">인원</th><th class="table-warning">수강료</th><th class="table-warning">강사료</th><th class="table-warning">수용비</th><th class="bg-cho3">인원</th><th class="bg-cho3">수강료</th><th class="bg-cho3">강사료</th><th class="bg-cho3">수용비</th><th class="bg-free">인원</th><th class="bg-free">수강료</th><th class="bg-free">강사료</th><th class="bg-free">수용비</th><th class="table-danger">인원</th><th class="table-danger">수강료</th><th class="table-danger">강사료</th><th class="table-danger">수용비</th></tr>`;
}

window.dlSampleCourse = function() { const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet([{부서명:'과학실험', 강좌수:2, '월 강사료':30000, '월 수용비':2000, '분기 기초 교재비':50000, 주간단위:1, 차수별시수:'4,4,4'}]), '부서마스터'); XLSX.writeFile(wb, '부서양식.xlsx'); };

window.upCourse = async function() {
    const file = $('fileCourse').files[0]; if (!file) return;
    try {
        const buf = await readFileAsArrayBuffer(file); const rows = parseXlsx(buf);
        if (rows.some(r => { const d = String(r['부서명']||r['강좌명']||'').trim(); return d && E.some(e => e.course.startsWith(d) && isQuarterLocked(e.q)); })) return alert('🔒 마감 분기의 부서 포함됨. 마감 해제 후 시도하세요.');
        rows.forEach(r => {
            const dept = String(r['부서명']||r['강좌명']||'').trim(); if (!dept) return;
            const cnt = num(r['강좌수'])||1, inst_m = num(r['월 강사료']||r['강사료']), mgmt_m = num(r['월 수용비']||r['수용비']), b = num(r['분기 기초 교재비']||r['교재비']||0), unit = num(r['주간단위'])||1, mh = String(r['차수별시수']||r['시수']||'4,4,4').trim();
            M[dept] = { 1:{cnt,inst_m,mgmt_m,b,unit,mh}, 2:{cnt,inst_m,mgmt_m,b,unit,mh}, 3:{cnt,inst_m,mgmt_m,b,unit,mh}, 4:{cnt,inst_m,mgmt_m,b,unit,mh} };
        });
        regenerateC(); E.forEach(e => { if (!isQuarterLocked(e.q) && M[e.course.replace(/\([A-Z]\)$/, '')]) recalcEnrollment(e); });
        save(); renderE(); alert('✅ 업로드 적용 완료');
    } catch(err) { alert('❌ 엑셀 구조 에러'); } finally { $('fileCourse').value=''; }
};

window.addDeptMaster = function() {
    const dept = val('c_dept'); if (!dept) return;
    const base = { cnt: num($('c_cnt').value)||1, inst_m: num(val('c_inst_m')), mgmt_m: num(val('c_mgmt_m')), b: num(val('c_b')), unit: num($('c_unit').value)||1, mh: val('c_mh')||'4,4,4' };
    M[dept] = { 1:{...base}, 2:{...base}, 3:{...base}, 4:{...base} }; regenerateC();
    ['c_dept','c_inst_m','c_mgmt_m','c_b','c_mh'].forEach(id => { if($(id)) $(id).value=''; }); alert('✅ 복사 등록 완료');
};

window.updateM = function(dept, k, el) {
    if(!M[dept] || !M[dept][window.gQ]) return;
    if (isQuarterLocked(window.gQ)) { alert('🔒 마감 변경 불가'); el.value = (k==='mh')?M[dept][window.gQ][k]:fmt(M[dept][window.gQ][k]); return; }
    if (k === 'mh') M[dept][window.gQ][k] = el.value.trim(); else { M[dept][window.gQ][k] = num(el.value); el.value = fmt(M[dept][window.gQ][k]); }
    regenerateC(); const aff = E.filter(e => e.course.startsWith(dept) && e.q === window.gQ); aff.forEach(e => recalcEnrollment(e)); save(); renderE();
};

window.delDept = function(dept) { if(confirm('삭제?')) { E = E.filter(e => !e.course.startsWith(dept)); delete M[dept]; regenerateC(); } };

window.regenerateC = function() {
    C = {};
    Object.keys(M).forEach(dept => {
        [1,2,3,4].forEach(q => {
            const md = M[dept][q]; if (!md) return; const mhArr = (md.mh||'4,4,4').split(',').map(x=>num(x)).filter(x=>x>0);
            const tH = mhArr.reduce((a,b)=>a+b, 0); const uS = (md.unit||1)*4;
            const qI = Math.round(((md.inst_m/uS)*tH)/10)*10; const qM = Math.round(((md.mgmt_m/uS)*tH)/10)*10;
            for(let i=0; i<md.cnt; i++) { let nm = md.cnt>1 ? `${dept}(${String.fromCharCode(65+i)})` : dept; if (!C[nm]) C[nm]={}; C[nm][q] = { t: qI+qM, b: md.b, mh: md.mh, instTot: qI, mgmtTot: qM }; }
        });
    }); save(); renderM(); renderC();
    if($('e_c')) $('e_c').innerHTML = '<option value="">강좌선택</option>' + Object.keys(C).map(nm => `<option value="${nm}">${nm}</option>`).join('');
};

function renderM() {
    if(!$('tbMaster')) return; const keys = Object.keys(M); if(!keys.length) return $('tbMaster').innerHTML = '<tbody><tr><td class="text-muted py-3">등록 부서 없음</td></tr></tbody>';
    let h = `<thead class="table-light"><tr><th>부서명</th><th>강좌수</th><th>월 강사료</th><th>월 수용비</th><th>기초 교재비</th><th>주간단위</th><th>시수</th><th>삭제</th></tr></thead><tbody>`;
    keys.forEach(dept => {
        const d = M[dept][window.gQ] || {cnt:1,inst_m:0,mgmt_m:0,b:0,unit:1,mh:'4,4,4'}; const safe = dept.replace(/'/g, "\\'");
        h += `<tr><td class="fw-bold align-middle text-primary">${dept}</td><td><input class="form-control form-control-sm text-center mx-auto" style="width:50px" value="${d.cnt}" onblur="updateM('${safe}','cnt',this)"></td><td><input class="fmt-num mx-auto" style="width:70px" value="${fmt(d.inst_m)}" onblur="updateM('${safe}','inst_m',this)"></td><td><input class="fmt-num mx-auto" style="width:70px" value="${fmt(d.mgmt_m)}" onblur="updateM('${safe}','mgmt_m',this)"></td><td><input class="fmt-num mx-auto" style="width:70px" value="${fmt(d.b)}" onblur="updateM('${safe}','b',this)"></td><td><input class="form-control form-control-sm text-center mx-auto" style="width:50px" value="${d.unit}" onblur="updateM('${safe}','unit',this)"></td><td><input class="form-control form-control-sm text-center mx-auto" style="width:60px" value="${d.mh}" onblur="updateM('${safe}','mh',this)"></td><td><button class="btn btn-sm btn-outline-danger py-0" onclick="delDept('${safe}')"><i class="bi bi-trash"></i></button></td></tr>`;
    }); $('tbMaster').innerHTML = h + '</tbody>';
}

function renderC() {
    if(!$('tbCourse')) return; const keys = Object.keys(C); if (!keys.length) return $('tbCourse').innerHTML = '<tbody><tr><td class="text-muted py-3">산출 강좌 없음</td></tr></tbody>';
    let h = `<thead class="table-light"><tr><th>생성 강좌명 (클릭: 팝업정산)</th><th class="table-warning">기초 수강료(분기)</th><th class="table-warning text-primary">강사료</th><th class="table-warning text-danger">수용비</th><th class="table-info">기초 교재비</th><th>시수</th></tr></thead><tbody>`;
    keys.forEach(nm => { const d = C[nm][window.gQ] || {t:0,b:0,mh:'',instTot:0,mgmtTot:0}; h += `<tr><td class="course-link" onclick="openCourseSummary('${nm.replace(/'/g, "\\'")}', window.gQ)">${nm}</td><td class="fw-bold">${fmt(d.t)}</td><td class="text-primary">${fmt(d.instTot)}</td><td class="text-danger">${fmt(d.mgmtTot)}</td><td>${fmt(d.b)}</td><td>${d.mh}</td></tr>`; }); $('tbCourse').innerHTML = h + '</tbody>';
}

window.dlSampleFree = function() { const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet([{학년:1, 반:1, 번호:1, 이름:'홍길동', 시작분기:1, 시작차수:1}]), '명단'); XLSX.writeFile(wb, '자유수강권.xlsx'); };
window.upFree = async function() {
    const file = $('fileFree').files[0]; if (!file) return;
    try {
        const buf = await readFileAsArrayBuffer(file); const rows = parseXlsx(buf); let added = 0;
        rows.forEach(r => { 
            const nm = String(r['이름']||r['성명']||'').trim(); if (!nm) return; 
            const g=num(r['학년']), b=num(r['반']), n=num(r['번호']), k=uid(g,b,n,nm); 
            const sQ=num(r['시작분기'])||1, sS=Math.max(0, (num(r['시작차수'])||1)-1);
            if (!F.some(x => uid(x.g,x.b,x.n,x.name)===k)) { F.push({g,b,n,name:nm, startQ:sQ, startSess:sS}); added++; } 
        });
        save(); renderF(); alert(`✅ 업로드 완료 (신규: ${added}건)`);
    } catch(err) { alert('❌ 에러'); } finally { $('fileFree').value = ''; if($('fileFreNm')) $('fileFreNm').textContent = '선택된 파일 없음'; }
};
window.addFree = function() { 
    const nm = val('f_nm'); if (!nm) return; 
    F.push({g:num(val('f_g')), b:num(val('f_b')), n:num(val('f_n')), name:nm, startQ:num(val('f_sq')), startSess:num(val('f_ss'))-1}); 
    save(); renderF(); ['f_n','f_nm'].forEach(id => { if($(id)) $(id).value = ''; }); if($('f_n')) $('f_n').focus(); 
};
window.delF = function(i) { F.splice(i,1); save(); renderF(); };

window.changeFreeStart = function(i) {
    const f = F[i];
    curEditFreeIdx = i;
    if($('fs_stuName')) $('fs_stuName').textContent = f.name;
    if($('fs_q')) $('fs_q').value = f.startQ || 1;
    if($('fs_s')) $('fs_s').value = (f.startSess || 0) + 1;
    if(mdlFreeStart) mdlFreeStart.show();
};

window.saveFreeStart = function() {
    if (curEditFreeIdx < 0) return;
    F[curEditFreeIdx].startQ = num(val('fs_q'));
    F[curEditFreeIdx].startSess = num(val('fs_s')) - 1;
    if(mdlFreeStart) mdlFreeStart.hide();
    save(); 
    renderF(); 
    autoRunSet(true); 
    renderE();
};

function renderF() { 
    if($('cnt_f')) $('cnt_f').textContent = F.length; if(!$('tbFree')) return; 
    $('tbFree').innerHTML = `<thead class="table-light"><tr><th>학적</th><th>이름</th><th>지원액</th><th>지원시점 (클릭수정)</th><th>관리</th></tr></thead><tbody>` + 
    F.map((f,i) => `<tr>
        <td>${dsp(f.g,f.b,f.n)}</td>
        <td class="fw-bold"><span class="clickable" onclick="openStuConsole('${uid(f.g,f.b,f.n,f.name).replace(/'/g,"\\'")}')">${f.name}</span></td>
        <td class="text-success fw-bold">600,000</td>
        <td><button class="btn btn-sm btn-primary rounded-pill py-0 px-3 fw-bold text-white shadow-sm border-0" onclick="changeFreeStart(${i})" title="클릭하여 지원 시점 변경" style="font-size:0.8rem;">${f.startQ}분기 ${f.startSess+1}차~ <i class="bi bi-pencil-square ms-1"></i></button></td>
        <td><button class="btn btn-sm btn-outline-danger py-0" onclick="delF(${i})">삭제</button></td>
    </tr>`).join('') + `</tbody>`; 
}

window.dlSampleEnroll = function() {
    const wb = XLSX.utils.book_new(); const cKeys = Object.keys(C);
    if (!cKeys.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet([{학년:1, 반:1, 번호:1, 이름:'홍길동', 강좌명:'강좌개설필요'}]), '명단'); else cKeys.forEach(c => XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet([{학년:'', 반:'', 번호:'', 이름:'', 비고:''}]), c.substring(0, 31)));
    XLSX.writeFile(wb, '수강명단양식.xlsx');
};
window.upEnroll = async function() {
    const fs = Array.from($('fileEnroll').files); if (!fs.length) return; const q = num(val('exEnQ')); if (isQuarterLocked(q)) return alert('🔒 마감 분기');
    const exist = new Set(E.map(e => `${e.q}_${e.course}_${uid(e.g,e.b,e.n,e.name)}`));
    for (const f of fs) {
        const buf = await readFileAsArrayBuffer(f); const wb = XLSX.read(new Uint8Array(buf), {type:'array'});
        wb.SheetNames.forEach(sn => {
            const rows = XLSX.utils.sheet_to_json(wb.Sheets[sn], {defval:''}); const c = sn.trim(); if (!C[c]) return;
            for (const r of rows) {
                const nm = String(r['이름']||r['성명']||'').trim(); if (!nm) continue; const g=num(r['학년']), b=num(r['반']), n=num(r['번호']), id=`${q}_${c}_${uid(g,b,n,nm)}`;
                if (exist.has(id)) continue;
                let newE = { q, g, b, n, name: nm, course: c, cT: null, cB: null, rT: 0, rB: 0, mm: String(r['비고']||'').trim(), tMemo:'', bMemo:'', refunds: [], adjusts: [] };
                recalcEnrollment(newE); E.push(newE); exist.add(id);
            }
        });
    } save(); renderE(); alert('✅ 업로드 완료'); $('fileEnroll').value = '';
};

window.addEnroll = function() {
    if(!val('e_c') || !val('e_nm')) return; const q = num(val('e_q')); if (isQuarterLocked(q)) return alert('🔒 마감 분기');
    let newE = { q, g: num(val('e_g')), b: num(val('e_b')), n: num(val('e_n')), name: val('e_nm'), course: val('e_c'), cT: null, cB: null, rT: 0, rB: 0, mm: '', tMemo:'', bMemo:'', refunds: [], adjusts: [] };
    recalcEnrollment(newE); E.push(newE); save(); renderE(); $('e_nm').value = ''; $('e_n').focus();
};
window.delE = function(i) { if(isQuarterLocked(E[i].q)) return alert('🔒 마감 변경 불가'); if(confirm('삭제?')) { E.splice(i,1); save(); renderE(); } };

window.toggleQ = function(q) { f_eq = (f_eq === String(q)) ? 'ALL' : String(q); renderE(); };
window.toggleC = function(c) { f_ec = (f_ec === c) ? 'ALL' : c; renderE(); };

function renderEFilters() {
    const el = $('tbMatrix'); if (!el) return; const cKeys = Object.keys(C).sort();
    if (!cKeys.length) return el.innerHTML = "<tr><td class='text-muted py-2'>강좌 없음</td></tr>";
    const stat = {}; cKeys.forEach(c => stat[c] = {1:0,2:0,3:0,4:0,tot:0}); const qTot = {1:0,2:0,3:0,4:0,tot:0};
    E.forEach(e => { if (stat[e.course]) { stat[e.course][e.q]++; stat[e.course].tot++; qTot[e.q]++; qTot.tot++; } });
    let h = `<thead class="table-light"><tr><th><button class="btn btn-sm btn-dark w-100" onclick="f_eq='ALL';f_ec='ALL';renderE();">전체</button></th>`;
    cKeys.forEach(c => h += `<th><button class="btn btn-sm w-100 ${f_ec===c?'btn-primary fw-bold':'btn-outline-primary'}" onclick="toggleC('${c.replace(/'/g,"\\'")}')">${c}</button></th>`);
    h += `<th class="bg-secondary text-white">계</th></tr></thead><tbody>`;
    [1,2,3,4].forEach(q => {
        h += `<tr><td><button class="btn btn-sm w-100 ${f_eq===String(q)?'btn-primary fw-bold':'btn-outline-primary'}" onclick="window.setQTab(${q});">${q}분기</button></td>`;
        cKeys.forEach(c => h += `<td class="${(f_eq===String(q)&&(f_ec==='ALL'||f_ec===c))?'bg-primary bg-opacity-10 fw-bold':''}">${stat[c][q]||'-'}</td>`);
        h += `<td class="fw-bold bg-light">${qTot[q]}</td></tr>`;
    });
    $('tbMatrix').innerHTML = h + `</tbody>`;
}

function renderE() {
    if($('cnt_e')) $('cnt_e').textContent = E.length; 
    renderEFilters();
    if(!$('tbEnroll')) return;
    
    const oA = $('chkOnlyAdjust')?.checked, oR = $('chkOnlyRefund')?.checked;
    const ls = E.map((e,i)=>({...e,_i:i})).filter(e => {
        if(f_eq !== 'ALL' && String(e.q) !== f_eq) return false; if(f_ec !== 'ALL' && e.course !== f_ec) return false;
        if(oA && (!e.adjusts || e.adjusts.length === 0)) return false; if(oR && (!e.refunds || e.refunds.length === 0)) return false; return true;
    }).sort((a,b) => a.q - b.q || a.name.localeCompare(b.name));
    
    if (ls.length === 0) {
        $('tbEnroll').innerHTML = `<tr><td colspan="7" class="text-muted py-3">조건에 맞는 수강생이 없습니다.</td></tr>`;
        return;
    }

    let h = `<thead class="table-light"><tr><th>분기</th><th>학적/이름 (팝업콘솔)</th><th>강좌명 (팝업명세)</th><th>실부담 수강료</th><th>실부담 교재비</th><th>상세 증빙 적요</th><th>관리</th></tr></thead><tbody>`;
    ls.forEach(e => {
        const locked = isQuarterLocked(e.q), rowCls = locked ? 'locked-row' : '';
        const info = (e.adjusts?.length>0 ? `<span class="badge bg-warning text-dark me-1">조정</span>` : '') + (e.refunds?.length>0 ? `<span class="badge bg-danger">환불</span>` : '');
        h += `<tr class="${rowCls}"><td><span class="badge bg-secondary">${e.q}분기</span></td><td class="fw-bold"><span class="clickable" onclick="openStuConsole('${uid(e.g,e.b,e.n,e.name).replace(/'/g,"\\'")}')">${dsp(e.g,e.b,e.n)} ${e.name}</span></td><td class="course-link" onclick="openCourseSummary('${e.course.replace(/'/g, "\\'")}', ${e.q})">${e.course}</td><td class="text-primary fw-bold">${fmt(e.cT)}</td><td class="text-success fw-bold">${fmt(e.cB)}</td><td class="text-start" style="font-size:0.8rem;">${info} ${e.mm||''}</td><td><button class="btn btn-sm btn-outline-danger py-0" onclick="delE(${e._i})" ${locked?'disabled':''}>삭제</button></td></tr>`;
    }); $('tbEnroll').innerHTML = h + '</tbody>';
}

window.getSessSplit = function(tAmt, sIdx, mhArr) {
    if (tAmt <= 0) return 0; const totalHours = mhArr.reduce((a,b)=>a+b, 0);
    if (sIdx === mhArr.length - 1) { let pSum = 0; for(let j=0; j<sIdx; j++) pSum += Math.floor((tAmt * (mhArr[j]/totalHours))/10)*10; return tAmt - pSum; }
    else return Math.floor((tAmt * (mhArr[sIdx]/totalHours))/10)*10;
};

window.recalcEnrollment = function(e) {
    const base = C[e.course]?.[e.q] || {t:0, b:0, mh:'4,4,4'}; const mhArr = (base.mh || '4,4,4').split(',').map(x=>num(x)).filter(x=>x>0);
    e.adjusts = e.adjusts || []; e.refunds = e.refunds || [];
    let totAdjT = e.adjusts.reduce((s, a) => s + (a.amtT || 0), 0);
    let totAdjB = e.adjusts.reduce((s, a) => s + (a.amtB || 0), 0);
    let tMemos = [], bMemos = [];

    e.refunds.forEach(r => {
        const bT = getSessSplit(base.t, r.sessIdx, mhArr);
        let rt = 0, ratioText = '';
        if (r.ty === 'BEFORE') { rt = bT; ratioText = '전액'; }
        else if (r.ty === 'DISEASE') {
            const md = M[e.course.replace(/\([A-Z]\)$/, '')]?.[e.q] || {};
            const unitFee = Math.ceil(((md.inst_m||0)+(md.mgmt_m||0))/((md.unit||1)*4)/10)*10;
            rt = Math.ceil((unitFee * r.ah)/10)*10; ratioText = `${r.ah}회`;
        }
        else if (r.ty === 'STUDENT') {
            if (r.ah === 0) { rt = bT; ratioText = '0회'; }
            else { const ratio = r.ah/(mhArr[r.sessIdx]||4); if (ratio <= 1/3) rt=Math.ceil(bT*(2/3)/10)*10; else if (ratio <= 1/2) rt=Math.ceil(bT*(1/2)/10)*10; }
            for (let j = r.sessIdx + 1; j < mhArr.length; j++) rt += getSessSplit(base.t, j, mhArr);
        }
        r.rt = rt; r.rb = r.reqBk ? getSessSplit(base.b, r.sessIdx, mhArr) : 0;
        const tyNm = {'BEFORE':'개시전','DISEASE':'결석','STUDENT':'포기'}[r.ty];
        r.tyNm = `[${tyNm}(${r.sessIdx+1}차)] 환:${fmt(r.rt)}`;
        if (r.rt>0) tMemos.push(r.tyNm); if (r.rb>0) bMemos.push(`[교재환불] -${fmt(r.rb)}`);
    });
    e.adjusts.forEach(a => { if(a.amtT!==0) tMemos.push(`[조정]${a.title}:${fmt(a.amtT)}`); if(a.amtB!==0) bMemos.push(`[교재조정]${a.title}:${fmt(a.amtB)}`); });
    e.tMemo = tMemos.join(', '); e.bMemo = bMemos.join(', '); e.mm = [e.tMemo, e.bMemo].filter(Boolean).join(' | ');
    e.rT = e.refunds.reduce((s,r)=>s+r.rt,0); e.rB = e.refunds.reduce((s,r)=>s+r.rb,0);
    e.cT = Math.max(0, base.t + totAdjT - e.rT); e.cB = Math.max(0, base.b + totAdjB - e.rB);
};

window.openStuConsole = function(uidStr) {
    cUid = uidStr; cEnrolls = []; E.forEach((e, idx) => { if (uid(e.g, e.b, e.n, e.name) === uidStr) cEnrolls.push(idx); });
    if(cEnrolls.length === 0) return alert('내역이 없습니다.'); autoRunSet(true);
    cActiveEIdx = cEnrolls[0]; const p = uidStr.split('-');
    if($('consoleTitle')) $('consoleTitle').innerHTML = `<i class="bi bi-person-lines-fill"></i> [${p[0]}학년 ${p[1]}반] ${p[3]} 통합 회계 콘솔`;
    renderConsole(); mdlConsole.show();
};

window.renderConsole = function() {
    const L = Ld[cUid] || { cB:0, fB:0, isC: false, isF: false }; 
    let totActual = 0; cEnrolls.forEach(i => totActual += (E[i].cT||0) + (E[i].cB||0));
    
    const txtC = L.isC ? `${fmt(L.cB)}원` : `<span class="text-muted fs-6 fw-normal">대상아님</span>`;
    const txtF = L.isF ? `${fmt(L.fB)}원` : `<span class="text-muted fs-6 fw-normal">대상아님</span>`;

    $('consoleTop').innerHTML = `<div><span class="small text-muted">초3 잔액</span><h5 class="text-primary fw-bold mb-0">${txtC}</h5></div><div><span class="small text-muted">자유 잔액</span><h5 class="text-success fw-bold mb-0">${txtF}</h5></div><div><span class="small text-muted">분기 총 실부담액</span><h5 class="text-danger fw-bold mb-0">${fmt(totActual)}원</h5></div>`;

    let hLeft = '';
    cEnrolls.forEach(i => {
        const e = E[i]; hLeft += `<div class="p-2 border rounded mb-2 clickable ${i===cActiveEIdx?'bg-primary bg-opacity-10 border-primary text-primary fw-bold':''}" onclick="setConsoleActive(${i})"><span class="badge bg-secondary me-1">${e.q}분기</span> ${e.course}</div>`;
    }); $('consoleLeft').innerHTML = hLeft;

    const e = E[cActiveEIdx], q = e.q, base = C[e.course]?.[q] || {t:0,b:0,mh:'4,4,4'};
    const totAdjT = e.adjusts.reduce((s,a)=>s+a.amtT, 0); 
    const totAdjB = e.adjusts.reduce((s,a)=>s+a.amtB, 0); 
    const locked = isQuarterLocked(q), dis = locked ? 'disabled' : '';

    let hRight = `<h6 class="fw-bold text-dark border-bottom pb-2">[${q}분기] ${e.course} 실부담금 명세 ${locked?'<span class="badge bg-danger ms-2">🔒 마감됨</span>':''}</h6>`;
    
    hRight += `<table class="table table-sm table-bordered text-center align-middle mb-3 bg-white" style="font-size:0.85rem;">
        <thead class="table-light">
            <tr><th>구분</th><th>기초 금액</th><th>분기 조정(할인)</th><th>누적 환불</th><th class="table-primary">최종 실부담금</th></tr>
        </thead>
        <tbody>
            <tr>
                <td class="fw-bold bg-light">수강료</td>
                <td>${fmt(base.t)}</td>
                <td class="${totAdjT !== 0 ? 'text-warning fw-bold' : ''}">${totAdjT > 0 ? '+' : ''}${fmt(totAdjT)}</td>
                <td class="${e.rT > 0 ? 'text-danger fw-bold' : ''}">${e.rT > 0 ? '-' : ''}${fmt(e.rT)}</td>
                <td class="table-primary fw-bold text-primary">${fmt(e.cT)}</td>
            </tr>
            <tr>
                <td class="fw-bold bg-light">교재비</td>
                <td>${fmt(base.b)}</td>
                <td class="${totAdjB !== 0 ? 'text-warning fw-bold' : ''}">${totAdjB > 0 ? '+' : ''}${fmt(totAdjB)}</td>
                <td class="${e.rB > 0 ? 'text-danger fw-bold' : ''}">${e.rB > 0 ? '-' : ''}${fmt(e.rB)}</td>
                <td class="table-primary fw-bold text-success">${fmt(e.cB)}</td>
            </tr>
        </tbody>
    </table>`;

    hRight += `<ul class="nav nav-pills mb-2" id="cTabs"><li class="nav-item"><button class="nav-link active py-1 small" data-bs-toggle="pill" data-bs-target="#tAdj">✍️ 분기 금액 조정</button></li><li class="nav-item"><button class="nav-link py-1 small ms-2" data-bs-toggle="pill" data-bs-target="#tRef">💸 차수별 환불/결석</button></li></ul>`;
    hRight += `<div class="tab-content border p-3 rounded bg-white mb-3">`;
    hRight += `<div class="tab-pane fade show active" id="tAdj"><div class="row g-1 mb-2"><div class="col-6"><input type="text" id="c_adj_title" class="form-control form-control-sm" placeholder="예: 다자녀할인, 면제" ${dis}></div><div class="col-3"><input type="number" id="c_adj_t" class="form-control form-control-sm text-end" placeholder="수강료 증감" ${dis}></div><div class="col-3"><input type="number" id="c_adj_b" class="form-control form-control-sm text-end" placeholder="교재비 증감" ${dis}></div></div><div class="text-end"><button class="btn btn-warning btn-sm fw-bold" onclick="addConsoleAdj()" ${dis}>조정 반영</button></div></div>`;
    
    const options = base.mh.split(',').map((_,idx)=>`<option value="${idx}">${idx+1}차수 환불</option>`).join('');
    hRight += `<div class="tab-pane fade" id="tRef"><div class="row g-1 mb-2"><div class="col-4"><select id="c_ref_idx" class="form-select form-select-sm">${options}</select></div><div class="col-5"><select id="c_ref_ty" class="form-select form-select-sm"><option value="BEFORE">개시전(전액)</option><option value="DISEASE">결석(일할)</option><option value="STUDENT">포기(구간합산)</option></select></div><div class="col-3"><input type="number" id="c_ref_ah" class="form-control form-control-sm text-center" value="0" placeholder="시수" ${dis}></div><div class="col-10 mt-1"><div class="form-check"><input class="form-check-input" type="checkbox" id="c_ref_bk" ${dis}><label class="form-check-label small text-danger fw-bold" for="c_ref_bk">교재비도 함께 반환</label></div></div></div><div class="text-end"><button class="btn btn-danger btn-sm fw-bold" onclick="addConsoleRef()" ${dis}>환불 등록</button></div></div>`;
    hRight += `</div>`;

    hRight += `<div class="small fw-bold text-muted mb-1">⏱️ 처리 이력 타임라인 (메모 격리 연동)</div><table class="table table-sm table-bordered text-center align-middle mb-0" style="font-size:0.75rem;"><thead class="table-light"><tr><th>유형</th><th>사유</th><th>수강료</th><th>교재비</th><th>삭제</th></tr></thead><tbody>`;
    let cnt = 0;
    e.adjusts.forEach((a, i) => { cnt++; hRight += `<tr><td><span class="badge bg-warning text-dark">조정</span></td><td class="text-start">${a.title}</td><td>${fmt(a.amtT)}</td><td>${fmt(a.amtB)}</td><td><button class="btn btn-sm btn-outline-danger py-0 px-1" onclick="delConsoleHist('adj', ${i})" ${dis}><i class="bi bi-x"></i></button></td></tr>`; });
    e.refunds.forEach((r, i) => { cnt++; hRight += `<tr><td><span class="badge bg-danger">환불</span></td><td class="text-start">${r.tyNm} (대상:${r.sessIdx+1}차)</td><td class="text-danger">-${fmt(r.rt)}</td><td class="text-danger">-${fmt(r.rb)}</td><td><button class="btn btn-sm btn-outline-danger py-0 px-1" onclick="delConsoleHist('ref', ${i})" ${dis}><i class="bi bi-x"></i></button></td></tr>`; });
    if(!cnt) hRight += `<tr><td colspan="5" class="text-muted">내역이 없습니다.</td></tr>`;
    hRight += `</tbody></table>`;

    $('consoleRight').innerHTML = hRight;
};

window.setConsoleActive = function(i) { cActiveEIdx = i; renderConsole(); };
window.addConsoleAdj = function() { const e = E[cActiveEIdx]; if (isQuarterLocked(e.q)) return; const t = val('c_adj_title'), aT = num(val('c_adj_t')), aB = num(val('c_adj_b')); if(!t) return alert('조정 사유 필수'); e.adjusts.push({ title:t, amtT:aT, amtB:aB }); recalcEnrollment(e); save(); renderConsole(); renderE(); };
window.addConsoleRef = function() { const e = E[cActiveEIdx]; if (isQuarterLocked(e.q)) return; const si = num($('c_ref_idx').value), ty = val('c_ref_ty'), ah = num(val('c_ref_ah')), bk = $('c_ref_bk').checked; e.refunds.push({ sessIdx:si, ty, ah, reqBk:bk, rt:0, rb:0, tyNm:'' }); recalcEnrollment(e); save(); renderConsole(); renderE(); };
window.delConsoleHist = function(ty, idx) { const e = E[cActiveEIdx]; if (isQuarterLocked(e.q)) return; if(ty==='adj') e.adjusts.splice(idx,1); else e.refunds.splice(idx,1); recalcEnrollment(e); save(); renderConsole(); renderE(); };

window.openCourseSummary = function(cName, q) {
    autoRunSet(true); if(!$('crsSummaryTitle') || !mdlCrsSummary) return;
    $('crsSummaryTitle').innerHTML = `<i class="bi bi-collection-play-fill"></i> [${q}분기] ${cName} 강좌 정산 명세서`;
    
    const list = Hs.filter(h => h.c === cName && h.q === q);
    let base = C[cName]?.[q] || {t:0, b:0};
    
    $('crsSummaryTop').innerHTML = `<div class="p-2 bg-light border rounded d-flex justify-content-around text-center">
        <div><strong>기초 수강료:</strong> ${fmt(base.t)}원</div>
        <div><strong>기초 교재비:</strong> ${fmt(base.b)}원</div>
        <div><strong>수강인원:</strong> <span class="text-primary fw-bold">${list.length}명</span></div>
    </div>`;

    const cSum = {sT:0, sB:0, tc:0, bc:0, tf:0, bf:0, finT:0, finB:0};
    list.forEach(hItem => { cSum.sT+=hItem.sT; cSum.sB+=hItem.sB; cSum.tc+=hItem.tc; cSum.bc+=hItem.bc; cSum.tf+=hItem.tf; cSum.bf+=hItem.bf; cSum.finT+=hItem.finT; cSum.finB+=hItem.finB; });
    
    let h = `<tr class="table-warning fw-bold sticky-total-row">
        <td colspan="3" class="text-end pe-3">총 합계</td>
        <td>${fmt(cSum.sT)}</td>
        <td>${fmt(cSum.sB)}</td>
        <td class="text-primary">${fmt(cSum.tc)}</td>
        <td class="text-primary">${fmt(cSum.bc)}</td>
        <td class="text-success">${fmt(cSum.tf)}</td>
        <td class="text-success">${fmt(cSum.bf)}</td>
        <td class="text-danger">${fmt(cSum.finT)}</td>
        <td class="text-danger">${fmt(cSum.finB)}</td>
    </tr>`;

    if(!list.length) h += `<tr><td colspan="11" class="text-muted py-3">수강생이 없습니다.</td></tr>`;
    else {
        list.sort((a,b)=>a.dp.localeCompare(b.dp)).forEach(hItem => {
            h += `<tr>
                <td>${hItem.dp}</td>
                <td class="fw-bold">${hItem.nm}</td>
                <td><span class="badge ${hItem.isF?'bg-success':(hItem.isC?'bg-primary':'bg-secondary')}">${hItem.isF?'자유':(hItem.isC?'초3':'일반')}</span></td>
                <td>${fmt(hItem.sT)}</td>
                <td>${fmt(hItem.sB)}</td>
                <td class="text-primary">${fmt(hItem.tc)}</td>
                <td class="text-primary">${fmt(hItem.bc)}</td>
                <td class="text-success">${fmt(hItem.tf)}</td>
                <td class="text-success">${fmt(hItem.bf)}</td>
                <td class="text-danger fw-bold bg-warning bg-opacity-10">${fmt(hItem.finT)}</td>
                <td class="text-danger fw-bold bg-warning bg-opacity-10">${fmt(hItem.finB)}</td>
            </tr>`;
        });
    }
    $('crsSummaryBody').innerHTML = h; mdlCrsSummary.show();
};

window.setFilt = function(f) { 
    s4_filt = f; 
    if($('fBtnA')) $('fBtnA').className = f === 'A' ? 'btn btn-sm btn-dark fw-bold' : 'btn btn-sm btn-outline-dark';
    if($('fBtnF')) $('fBtnF').className = f === 'F' ? 'btn btn-sm btn-success fw-bold' : 'btn btn-sm btn-outline-success';
    if($('fBtnC')) $('fBtnC').className = f === 'C' ? 'btn btn-sm btn-primary fw-bold' : 'btn btn-sm btn-outline-primary';
    renderSetTabs(); 
};

window.autoRunSet = function(silent = false) {
    Ld = {}; Hs = []; 
    if (!E.length) { if (!silent) renderSetTabs(); return; }
    
    const freeMap = new Map();
    F.forEach(f => freeMap.set(uid(f.g,f.b,f.n,f.name), f));

    E.forEach(e => {
        const id  = uid(e.g,e.b,e.n,e.name); 
        const fData = freeMap.get(id);
        const isF = !!fData, isC = String(e.g) === '3';
        if (!Ld[id]) Ld[id] = { id, dp: dsp(e.g,e.b,e.n), nm: e.name, isF, isC, cB: 0, fB: isF ? 600000 : 0, fData, ty: (isF?'자유':'')+(isC?'초3':'')||'일반', enrolls: [] };
        Ld[id].enrolls.push(e);
    });
    
    Object.values(Ld).forEach(L => { if (L.isC) { const qs = [...new Set(L.enrolls.map(e=>e.q))]; if (qs.some(q=>q<=2)) L.cB += 250000; if (qs.some(q=>q>=3)) L.cB += 250000; } });
    
    Object.values(Ld).forEach(L => {
        L.enrolls.sort((a,b)=>a.q-b.q).forEach(e => {
            const bs = C[e.course]?.[e.q] || {t:0,b:0};
            const mhArr = (bs.mh || '4,4,4').split(',').map(x=>num(x)).filter(x=>x>0);
            const totAdjT = e.adjusts.reduce((s,a)=>s+a.amtT, 0);
            const totAdjB = e.adjusts.reduce((s,a)=>s+a.amtB, 0);

            let q_tc = 0, q_bc = 0, q_tf = 0, q_bf = 0, q_finT = 0, q_finB = 0;
            let sessDetails = [];

            for(let sIdx = 0; sIdx < mhArr.length; sIdx++) {
                let bT = getSessSplit(bs.t, sIdx, mhArr);
                let bB = getSessSplit(bs.b, sIdx, mhArr);
                let aT = getSessSplit(totAdjT, sIdx, mhArr);
                let aB = getSessSplit(totAdjB, sIdx, mhArr);
                let rT = e.refunds.filter(r=>r.sessIdx === sIdx).reduce((s,r)=>s+r.rt, 0);
                let rB = e.refunds.filter(r=>r.sessIdx === sIdx).reduce((s,r)=>s+r.rb, 0);

                let sRemT = Math.max(0, bT + aT - rT);
                let sRemB = Math.max(0, bB + aB - rB);

                let s_tc = 0, s_bc = 0, s_tf = 0, s_bf = 0;
                const ded = (amt, cho) => { const av = cho ? L.cB : L.fB; const d = Math.min(amt, av); if (cho) L.cB -= d; else L.fB -= d; return d; };

                if (L.isC) {
                    s_tc = ded(sRemT, true); sRemT -= s_tc;
                    s_bc = ded(sRemB, true); sRemB -= s_bc;
                }

                if (L.isF) {
                    const sQ = L.fData.startQ || 1;
                    const sS = L.fData.startSess || 0;
                    let isEligible = false;
                    if (e.q > sQ) isEligible = true;
                    else if (e.q === sQ && sIdx >= sS) isEligible = true;

                    if (isEligible) {
                        s_tf = ded(sRemT, false); sRemT -= s_tf;
                        s_bf = ded(sRemB, false); sRemB -= s_bf;
                    }
                }

                q_tc += s_tc; q_bc += s_bc;
                q_tf += s_tf; q_bf += s_bf;
                q_finT += sRemT; q_finB += sRemB;

                sessDetails.push({ sIdx, tc:s_tc, bc:s_bc, tf:s_tf, bf:s_bf, finT:sRemT, finB:sRemB, bT });
            }

            Hs.push({ 
                q:e.q, id:L.id, dp:L.dp, nm:L.nm, c:e.course, e, origT:bs.t, origB:bs.b, 
                sT:e.cT!==null?e.cT:bs.t, sB:e.cB!==null?e.cB:bs.b, 
                tc:q_tc, bc:q_bc, tf:q_tf, bf:q_bf, finT:q_finT, finB:q_finB, 
                sessDetails, isF:L.isF, isC:L.isC, g:e.g, ban:e.b, num:e.n 
            });
        });
    }); if (!silent) renderSetTabs();
};

window.toggleSessCheck = function(targetQ, sessIdx, isChecked) {
    const key = `${targetQ}_${sessIdx}`;
    if (!isChecked) {
        if (confirm(`해당 분기 ${sessIdx+1}차 마감을 해제하시겠습니까?`)) { delete SysSet.closedSess[key]; save(); alert('마감이 해제되었습니다.'); renderSetTabs(); renderE(); }
        else $('chkClose_'+sessIdx).checked = true;
    } else {
        if (confirm(`해당 분기 ${sessIdx+1}차수를 마감하시겠습니까?\n이 시점의 청구액이 고정되며, 이후 발생하는 환불은 사후 반환용으로 격리됩니다.`)) {
            autoRunSet(true); const snapshot = {}; const ls = Hs.filter(h => h.q === targetQ);
            ls.forEach(h => {
                const mhArr = (C[h.c]?.[targetQ]?.mh || '4,4,4').split(',').map(x => num(x)).filter(x => x > 0);
                if (sessIdx >= mhArr.length) return;
                
                const bT = getSessSplit(h.sT + h.e.rT - (h.e.adjusts.reduce((s,a)=>s+a.amtT,0)), sessIdx, mhArr);
                const aT = Math.floor((h.e.adjusts.reduce((s,a)=>s+a.amtT, 0) / mhArr.length)/10)*10;
                let grossSessAmt = bT + aT;
                let sRef = (h.e.refunds||[]).filter(r=>r.sessIdx===sessIdx).reduce((s,r)=>s+r.rt, 0);
                let actualSessAmt = Math.max(0, grossSessAmt - sRef);

                let selfA = getSessSplit(h.finT, sessIdx, mhArr), cho3A = getSessSplit(h.tc, sessIdx, mhArr), freeA = getSessSplit(h.tf, sessIdx, mhArr);
                let cTot = selfA + cho3A + freeA;
                if (cTot > actualSessAmt) {
                    let toDed = cTot - actualSessAmt;
                    if (toDed > 0 && selfA > 0) { const d = Math.min(toDed, selfA); selfA -= d; toDed -= d; }
                    if (toDed > 0 && cho3A > 0) { const d = Math.min(toDed, cho3A); cho3A -= d; toDed -= d; }
                    if (toDed > 0 && freeA > 0) { const d = Math.min(toDed, freeA); freeA -= d; toDed -= d; }
                } else if (cTot < actualSessAmt) {
                    let toAdd = actualSessAmt - cTot;
                    if (h.finT > 0 || (h.tc === 0 && h.tf === 0)) selfA += toAdd; else if (h.tc > 0) cho3A += toAdd; else freeA += toAdd;
                }
                snapshot[h.id] = { selfAmt: selfA, cho3Amt: cho3A, freeAmt: freeA };
            });
            SysSet.closedSess[key] = snapshot; save(); alert('마감되었습니다.'); renderSetTabs(); renderE();
        } else $('chkClose_'+sessIdx).checked = false;
    }
};

window.renderSetTabs = function() {
    const qVal = val('s4_q') || String(window.gQ); 
    const hList = Hs.filter(h => (String(h.q)===qVal) && (s4_filt==='A' || (s4_filt==='F'&&h.isF) || (s4_filt==='C'&&h.isC)));
    const chkWrap = $('closeSessChecks');
    
    if(chkWrap) {
        chkWrap.style.setProperty('display', 'flex', 'important'); let chks = `<span class="small fw-bold text-dark">🔒 마감:</span>`;
        let maxSess = 1; Object.keys(C).forEach(c => { const m = (C[c]?.[qVal]?.mh || '4,4,4').split(',').filter(x => num(x) > 0).length; if (m > maxSess) maxSess = m; });
        for (let i = 0; i < maxSess; i++) { const key = `${qVal}_${i}`; const isChecked = SysSet.closedSess[key] ? 'checked' : ''; chks += `<div class="form-check form-check-inline mb-0 ms-2"><input class="form-check-input border-warning" type="checkbox" id="chkClose_${i}" ${isChecked} onchange="toggleSessCheck(${qVal}, ${i}, this.checked)"><label class="form-check-label small fw-bold" for="chkClose_${i}">${i+1}차</label></div>`; }
        chkWrap.innerHTML = chks;
    }
    
    if (hList.length === 0) {
        if($('tbStat')) $('tbStat').innerHTML = `<tr><td colspan="10" class="text-muted py-3">해당 분기 및 조건에 맞는 수강 내역이 없습니다.</td></tr>`;
        if($('tbStuDtl')) $('tbStuDtl').innerHTML = `<tr><td colspan="10" class="text-muted py-3">해당 분기 및 조건에 맞는 수강 내역이 없습니다.</td></tr>`;
        if($('tbCrseDtl')) $('tbCrseDtl').innerHTML = `<tr><td colspan="12" class="text-muted py-3">해당 분기 및 조건에 맞는 수강 내역이 없습니다.</td></tr>`;
        return;
    }

    let sH = ''; const st = {}; hList.forEach(h => { if (!st[h.c]) st[h.c] = {cnt:0,sT:0,sB:0,tc:0,bc:0,tf:0,bf:0,fT:0,fB:0}; const s = st[h.c]; s.cnt++; s.sT+=h.sT; s.sB+=h.sB; s.tc+=h.tc; s.bc+=h.bc; s.tf+=h.tf; s.bf+=h.bf; s.fT+=h.finT; s.fB+=h.finB; });
    Object.keys(st).sort().forEach(c => { const s = st[c]; sH += `<tr><td class="course-link" onclick="openCourseSummary('${c.replace(/'/g, "\\'")}', ${qVal})">${c}</td><td class="table-warning fw-bold">${s.cnt}</td><td class="table-warning">${fmt(s.sT)}</td><td class="table-warning">${fmt(s.sB)}</td><td class="bg-cho3 text-primary">${fmt(s.tc)}</td><td class="bg-cho3">${fmt(s.bc)}</td><td class="bg-free text-success">${fmt(s.tf)}</td><td class="bg-free">${fmt(s.bf)}</td><td class="table-danger fw-bold">${fmt(s.fT)}</td><td class="table-danger text-danger fw-bold">${fmt(s.fB)}</td></tr>`; });
    sH += `<tr class="table-dark fw-bold sticky-total-row"><td colspan="2" class="text-warning">총 합계</td><td class="text-warning">${fmt(hList.reduce((s,h)=>s+h.sT,0))}</td><td class="text-warning">${fmt(hList.reduce((s,h)=>s+h.sB,0))}</td><td class="text-primary">${fmt(hList.reduce((s,h)=>s+h.tc,0))}</td><td class="text-primary">${fmt(hList.reduce((s,h)=>s+h.bc,0))}</td><td class="text-success">${fmt(hList.reduce((s,h)=>s+h.tf,0))}</td><td class="text-success">${fmt(hList.reduce((s,h)=>s+h.bf,0))}</td><td class="text-danger">${fmt(hList.reduce((s,h)=>s+h.finT,0))}</td><td class="text-danger">${fmt(hList.reduce((s,h)=>s+h.finB,0))}</td></tr>`;
    if($('tbStat')) $('tbStat').innerHTML = sH;

    let stuH = ''; const lMap = {}; hList.forEach(h => { if (!lMap[h.id]) lMap[h.id] = {L: Ld[h.id], items:[]}; lMap[h.id].items.push(h); });
    Object.values(lMap).forEach(grp => { grp.items.forEach((h, idx) => { stuH += `<tr>`; if (idx === 0) stuH += `<td rowspan="${grp.items.length}">${grp.L.dp}</td><td rowspan="${grp.items.length}" class="fw-bold"><span class="clickable" onclick="openStuConsole('${grp.L.id}')">${grp.L.nm}</span></td><td rowspan="${grp.items.length}">${grp.L.ty}</td><td rowspan="${grp.items.length}">${fmt(grp.L.cB)}</td><td rowspan="${grp.items.length}">${fmt(grp.L.fB)}</td>`; stuH += `<td>${h.q}</td><td class="course-link" onclick="openCourseSummary('${h.c.replace(/'/g, "\\'")}', ${h.q})">${h.c}</td><td>${fmt(h.sT)}</td><td>${fmt(h.sB)}</td><td class="bg-cho3">${fmt(h.tc)}</td><td class="bg-cho3">${fmt(h.bc)}</td><td class="bg-free">${fmt(h.tf)}</td><td class="bg-free">${fmt(h.bf)}</td><td class="text-danger fw-bold">${fmt(h.finT)}</td><td class="text-danger fw-bold">${fmt(h.finB)}</td></tr>`; }); });
    if($('tbStuDtl')) $('tbStuDtl').innerHTML = stuH;

    if($('cFilterBtnGroup')) { 
        let bh = `<button class="btn btn-sm ${s4_cFilter==='ALL'?'btn-primary fw-bold':'btn-outline-secondary'}" onclick="s4_cFilter='ALL';renderSetTabs();">전체강좌</button>`; 
        Object.keys(C).forEach(c => {
            bh += `<button class="btn btn-sm ${s4_cFilter===c?'btn-primary fw-bold':'btn-outline-secondary'} ms-1" onclick="s4_cFilter='${c.replace(/'/g,"\\'")}';renderSetTabs();">${c}</button>`;
        }); 
        $('cFilterBtnGroup').innerHTML = bh; 
    }
    
    let cList = hList; if (s4_cFilter !== 'ALL') cList = cList.filter(h => h.c === s4_cFilter);
    const cSum = {sT:0, sB:0, tc:0, bc:0, tf:0, bf:0, finT:0, finB:0}; 
    cList.forEach(h => { cSum.sT+=h.sT; cSum.sB+=h.sB; cSum.tc+=h.tc; cSum.bc+=h.bc; cSum.tf+=h.tf; cSum.bf+=h.bf; cSum.finT+=h.finT; cSum.finB+=h.finB; });
    let crsH = `<tr class="table-warning fw-bold sticky-total-row"><td colspan="4" class="text-end pe-3">총 합계</td><td>${fmt(cSum.sT)}</td><td>${fmt(cSum.sB)}</td><td class="text-primary">${fmt(cSum.tc)}</td><td class="text-primary">${fmt(cSum.bc)}</td><td class="text-success">${fmt(cSum.tf)}</td><td class="text-success">${fmt(cSum.bf)}</td><td class="text-danger">${fmt(cSum.finT)}</td><td class="text-danger">${fmt(cSum.finB)}</td></tr>`;
    cList.forEach(h => { crsH += `<tr><td>${h.q}분기</td><td>${h.dp}</td><td class="fw-bold"><span class="clickable" onclick="openStuConsole('${h.id}')">${h.nm}</span></td><td class="course-link" onclick="openCourseSummary('${h.c.replace(/'/g, "\\'")}', ${h.q})">${h.c}</td><td>${fmt(h.sT)}</td><td>${fmt(h.sB)}</td><td class="bg-cho3">${fmt(h.tc)}</td><td class="bg-cho3">${fmt(h.bc)}</td><td class="bg-free">${fmt(h.tf)}</td><td class="bg-free">${fmt(h.bf)}</td><td class="text-danger fw-bold">${fmt(h.finT)}</td><td class="text-danger fw-bold">${fmt(h.finB)}</td></tr>`; });
    if($('tbCrseDtl')) $('tbCrseDtl').innerHTML = crsH;
};

// =====================================================
// STEP 5 : 문서 서식 자동생성
// =====================================================
let eduDataCached = []; 
window.initStep5 = function() { autoRunSet(true); buildEduTabs(); renderPreviewInvoice(); renderPreviewRef(); renderPreviewRoster(); };

function buildEduTabs() {
    const q = window.gQ; const ls = Hs.filter(h => h.q === q && (h.finT > 0 || h.finB > 0));
    const grouped = {}; ls.forEach(h => { const baseC = h.c.replace(/\s*\([A-Za-z가-힣0-9]+\)$/, '').trim(); if (!grouped[baseC]) grouped[baseC] = []; grouped[baseC].push(h); });
    eduDataCached = [];
    Object.keys(grouped).forEach(bc => {
        const sub = grouped[bc]; let maxS = 1; sub.forEach(h => { const m = (C[h.c]?.[q]?.mh||'4,4,4').split(',').length; if(m>maxS) maxS=m; });
        for(let i=0; i<maxS; i++) { sub.forEach(h => { if(h.finT<=0) return; const mArr = (C[h.c]?.[q]?.mh||'4,4,4').split(',').map(x=>num(x)).filter(x=>x>0); if(i < mArr.length) { let amt = getSessSplit(h.finT, i, mArr); if(amt>0) eduDataCached.push({ sheet: bc+` ${i+1}차 수강료`, g:h.g, b:h.ban, n:h.num, nm:h.nm, amt:amt }); } }); }
        sub.filter(h => h.finB > 0).forEach(h => { eduDataCached.push({ sheet: bc+' 재료비', g:h.g, b:h.ban, n:h.num, nm:h.nm, amt:h.finB }); });
    });
    const sheetNames = [...new Set(eduDataCached.map(d => d.sheet))];
    let hTabs = sheetNames.map((sn, idx) => `<button class="sheet-pill ${idx===0?'active':''}" onclick="renderEduSheet('${sn}', this)">${sn}</button>`).join('');
    $('eduSheetTabs').innerHTML = hTabs || '<div class="small text-muted">수납 대상 없음</div>';
    if(sheetNames.length) renderEduSheet(sheetNames[0]); else $('prev_edu').innerHTML = '';
}

window.renderEduSheet = function(sn, el) {
    if(el) { Array.from(el.parentNode.children).forEach(b => b.classList.remove('active')); el.classList.add('active'); }
    const filtered = eduDataCached.filter(d => d.sheet === sn);
    const total = filtered.reduce((s, d) => s + d.amt, 0);
    let h = `<tr class="sticky-total-row fw-bold"><td colspan="2" class="text-end">시트 합계</td><td class="text-danger">${fmt(total)}원</td><td></td></tr>`;
    h += filtered.map(d => `<tr><td>${dsp(d.g, d.b, d.n)}</td><td>${d.nm}</td><td>${fmt(d.amt)}</td><td>${sn}</td></tr>`).join('');
    $('prev_edu').innerHTML = h;
};

window.exEdu = function() {
    const q = window.gQ; if (!eduDataCached.length) return alert('추출할 내역이 없습니다.');
    const wb = XLSX.utils.book_new(); const sg = {};
    eduDataCached.forEach(r => { if(!sg[r.sheet]) sg[r.sheet]=[]; sg[r.sheet].push({ '* 학과': r.sheet.replace(/ (1|2|3|4)차 수강료| 재료비/g, ''), '* 학년': r.g, '* 반': r.b, '* 번호': r.n, '* 성명': r.nm, '* 대상금액': r.amt }); });
    Object.keys(sg).forEach(sn => { const total = sg[sn].reduce((sum, r) => sum + r['* 대상금액'], 0); sg[sn].push({ '* 학과': '총계', '* 학년': '', '* 반': '', '* 번호': '', '* 성명': '', '* 대상금액': total }); XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(sg[sn]), sn.substring(0, 31)); });
    XLSX.writeFile(wb, `${q}분기_에듀파인_수납요구서.xlsx`);
};

function getInvoiceData(q, sVal) {
    const ls = Hs.filter(h => h.q === q);
    const grouped = {};
    
    ls.forEach(h => {
        const mhArr = (C[h.c]?.[q]?.mh || '4,4,4').split(',').map(x => num(x)).filter(x => x > 0);
        const baseC = h.c.replace(/\s*\([A-Za-z가-힣0-9]+\)$/, '').trim();
        const cd = C[h.c]?.[q] || { instTot: 0, mgmtTot: 0 };
        
        let baseT=0, baseM=0, baseI=0, sSelf=0, sCho=0, sFree=0;

        if (sVal === 'ALL') {
            baseT = h.origT;
            if (baseT <= 0) return;
            baseM = cd.mgmtTot;
            baseI = baseT - baseM;
            sSelf = h.finT; sCho = h.tc; sFree = h.tf;
        } else {
            const sIdx = num(sVal) - 1;
            if (sIdx >= mhArr.length) return;
            
            baseT = h.sessDetails[sIdx].bT;
            if (baseT <= 0) return;
            baseM = getSessSplit(cd.mgmtTot, sIdx, mhArr);
            baseI = baseT - baseM;

            let closedSnapshot = null; if (SysSet.closedSess && SysSet.closedSess[`${q}_${sIdx}`]) closedSnapshot = SysSet.closedSess[`${q}_${sIdx}`][h.id];

            if (closedSnapshot) {
                sSelf = closedSnapshot.selfAmt || 0; sCho = closedSnapshot.cho3Amt || 0; sFree = closedSnapshot.freeAmt || 0;
            } else {
                sSelf = h.sessDetails[sIdx].finT;
                sCho = h.sessDetails[sIdx].tc;
                sFree = h.sessDetails[sIdx].tf;
            }
        }

        let totF = sSelf + sCho + sFree;
        let mSelf=0, mCho=0, mFree=0, iSelf=0, iCho=0, iFree=0;

        if (totF > 0) {
            let totalM = baseT > 0 ? Math.floor(totF * (baseM / baseT) / 10) * 10 : 0;
            let remM = totalM;
            
            if (sSelf > 0) { mSelf = Math.floor((sSelf / totF) * totalM / 10) * 10; remM -= mSelf; }
            if (sCho > 0) { mCho = Math.floor((sCho / totF) * totalM / 10) * 10; remM -= mCho; }
            
            if (sFree > 0) { mFree = remM; remM = 0; }
            else if (sCho > 0) { mCho += remM; remM = 0; }
            else if (sSelf > 0) { mSelf += remM; remM = 0; }
            
            iSelf = sSelf - mSelf;
            iCho = sCho - mCho;
            iFree = sFree - mFree;
        }

        if(!grouped[baseC]) grouped[baseC] = { 
            c:baseC, baseT, baseI, baseM, 
            selfCnt:0, selfFee:0, selfInst:0, selfMgmt:0, 
            cho3Cnt:0, cho3Fee:0, cho3Inst:0, cho3Mgmt:0, 
            freeCnt:0, freeFee:0, freeInst:0, freeMgmt:0, 
            totCnt:0, totFee:0, totInst:0, totMgmt:0, memos:[] 
        };
        const g = grouped[baseC];

        if(sSelf>0){ g.selfCnt++; g.selfFee+=sSelf; g.selfInst+=iSelf; g.selfMgmt+=mSelf; }
        if(sCho>0){ g.cho3Cnt++; g.cho3Fee+=sCho; g.cho3Inst+=iCho; g.cho3Mgmt+=mCho; }
        if(sFree>0){ g.freeCnt++; g.freeFee+=sFree; g.freeInst+=iFree; g.freeMgmt+=mFree; }

        if (totF > 0 || baseT > 0) {
            g.totCnt++;
            g.totFee += totF;
            g.totInst += (iSelf + iCho + iFree);
            g.totMgmt += (mSelf + mCho + mFree);
            if(h.e.tMemo) g.memos.push(`${h.nm}(${h.e.tMemo})`);
        }
    });
    
    return Object.values(grouped).sort((a,b) => a.c.localeCompare(b.c));
}

window.renderPreviewInvoice = function() {
    const q = window.gQ, sVal = val('p_sInvoice');
    const data = getInvoiceData(q, sVal);
    let h = '';
    if(!data.length) h = `<tr><td colspan=\"23\" class=\"text-muted py-3\">데이터가 없습니다.</td></tr>`;
    else {
        let tSelf=0, tSelfI=0, tSelfM=0, tCho=0, tChoI=0, tChoM=0, tFree=0, tFreeI=0, tFreeM=0, tTot=0, tTotI=0, tTotM=0;
        data.forEach(g => { 
            tSelf+=g.selfFee; tSelfI+=g.selfInst; tSelfM+=g.selfMgmt;
            tCho+=g.cho3Fee; tChoI+=g.cho3Inst; tChoM+=g.cho3Mgmt;
            tFree+=g.freeFee; tFreeI+=g.freeInst; tFreeM+=g.freeMgmt;
            tTot+=g.totFee; tTotI+=g.totInst; tTotM+=g.totMgmt;
        });
        
        h += `<tr class=\"sticky-total-row fw-bold text-center\"><td colspan=\"6\" class=\"text-end pe-3\">총 합계</td><td class=\"text-warning\">${fmt(tSelf)}</td><td class=\"text-warning\">${fmt(tSelfI)}</td><td class=\"text-warning\">${fmt(tSelfM)}</td><td></td><td class=\"text-primary\">${fmt(tCho)}</td><td class=\"text-primary\">${fmt(tChoI)}</td><td class=\"text-primary\">${fmt(tChoM)}</td><td></td><td class=\"text-success\">${fmt(tFree)}</td><td class=\"text-success\">${fmt(tFreeI)}</td><td class=\"text-success\">${fmt(tFreeM)}</td><td></td><td class=\"text-danger fs-6\">${fmt(tTot)}</td><td class=\"text-danger\">${fmt(tTotI)}</td><td class=\"text-danger\">${fmt(tTotM)}</td><td colspan=\"2\"></td></tr>`;
        
        data.forEach((g, idx) => {
            const uniqueMemos = [...new Set(g.memos)]; const diffFee = g.totFee - (g.baseT * g.totCnt);
            h += `<tr><td>${idx+1}</td><td class=\"course-link\" onclick=\"openCourseSummary('${g.c.replace(/'/g, "\\'")}', window.gQ)\">${g.c}</td><td>${fmt(g.baseT)}</td><td>${fmt(g.baseI)}</td><td>${fmt(g.baseM)}</td><td class=\"table-warning\">${g.selfCnt}</td><td class=\"table-warning\">${fmt(g.selfFee)}</td><td class=\"table-warning\">${fmt(g.selfInst)}</td><td class=\"table-warning\">${fmt(g.selfMgmt)}</td><td class=\"table-primary\">${g.cho3Cnt}</td><td class=\"table-primary\">${fmt(g.cho3Fee)}</td><td class=\"table-primary\">${fmt(g.cho3Inst)}</td><td class=\"table-primary\">${fmt(g.cho3Mgmt)}</td><td class=\"table-success\">${g.freeCnt}</td><td class=\"table-success\">${fmt(g.freeFee)}</td><td class=\"table-success\">${fmt(g.freeInst)}</td><td class=\"table-success\">${fmt(g.freeMgmt)}</td><td class=\"table-danger fw-bold\">${g.totCnt}</td><td class=\"table-danger fw-bold\">${fmt(g.totFee)}</td><td class=\"table-danger\">${fmt(g.totInst)}</td><td class=\"table-danger\">${fmt(g.totMgmt)}</td><td class=\"table-secondary text-danger fw-bold\">${fmt(diffFee)}</td><td class=\"text-start small\" style=\"max-width:150px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;\" title=\"${uniqueMemos.join(', ')}\">${uniqueMemos.join(', ')}</td></tr>`;
        });
    } 
    if($('prev_inv')) $('prev_inv').innerHTML = h;
};

window.exInvoice = function() {
    const q = window.gQ, sVal = val('p_sInvoice'); const data = getInvoiceData(q, sVal);
    if (!data.length) return alert('출력할 데이터가 없습니다.');
    const wb = XLSX.utils.book_new();
    const aoa = [
        [`방과후학교 ${q}분기 ${sVal==='ALL'?'전체합산':sVal+'차'} 교육비 청구서`], [], [`1. 교육기간 : `], [`2. 입금계좌 : `], [`3. 청구내용 : `], [],
        [ '순번', '부서명', '1인당수강료', '1인당강사료', '1인당수용비', '수익자인원', '수익자수강료', '수익자강사료', '수익자수용비', '초3인원', '초3수강료', '초3강사료', '초3수용비', '자유인원', '자유수강료', '자유강사료', '자유수용비', '합계인원', '합계수강료', '합계강사료', '합계수용비', '차액(환불/조정)', '비고(수강료적요)' ]
    ];
    let idx = 1; let tSelf=0, tSelfI=0, tSelfM=0, tCho=0, tChoI=0, tChoM=0, tFree=0, tFreeI=0, tFreeM=0, tTot=0, tTotI=0, tTotM=0;
    
    data.forEach(g => {
        const diffFee = g.totFee - (g.baseT * g.totCnt);
        aoa.push([ idx++, g.c, g.baseT, g.baseI, g.baseM, g.selfCnt, g.selfFee, g.selfInst, g.selfMgmt, g.cho3Cnt, g.cho3Fee, g.cho3Inst, g.cho3Mgmt, g.freeCnt, g.freeFee, g.freeInst, g.freeMgmt, g.totCnt, g.totFee, g.totInst, g.totMgmt, diffFee, [...new Set(g.memos)].join(', ') ]);
        
        tSelf+=g.selfFee; tSelfI+=g.selfInst; tSelfM+=g.selfMgmt;
        tCho+=g.cho3Fee; tChoI+=g.cho3Inst; tChoM+=g.cho3Mgmt;
        tFree+=g.freeFee; tFreeI+=g.freeInst; tFreeM+=g.freeMgmt;
        tTot+=g.totFee; tTotI+=g.totInst; tTotM+=g.totMgmt;
    });
    aoa.push(['총계', '', '', '', '', '', tSelf, tSelfI, tSelfM, '', tCho, tChoI, tChoM, '', tFree, tFreeI, tFreeM, '', tTot, tTotI, tTotM, '', '']);
    
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!merges'] = [ {s:{r:0,c:0},e:{r:0,c:22}}, {s:{r:2,c:0},e:{r:2,c:3}}, {s:{r:2,c:4},e:{r:2,c:22}}, {s:{r:3,c:0},e:{r:3,c:3}}, {s:{r:3,c:4},e:{r:3,c:22}}, {s:{r:4,c:0},e:{r:4,c:22}} ];
    XLSX.utils.book_append_sheet(wb, ws, `${sVal==='ALL'?'전체':sVal+'차'} 청구서`); 
    XLSX.writeFile(wb, `${q}분기_${sVal==='ALL'?'전체합산':sVal+'차'}_교육비청구서.xlsx`);
};

window.exRef = function() {
    const q = window.gQ;
    const data = E.filter(e => e.q === q && (e.rT>0 || e.rB>0 || (e.adjusts&&e.adjusts.length>0))).map(e => ({ q: e.q, c: e.course, dp: dsp(e.g,e.b,e.n), nm: e.name, g: e.g, b: e.b, n: e.n, rT: e.rT, rB: e.rB, cT: e.cT, cB: e.cB, mm: e.mm }));
    if (!data.length) return alert('환불 내역이 없습니다.');
    const wb = XLSX.utils.book_new();
    const rows = data.map(r => ({ '분기': r.q+'분기', '학년': r.g, '반': r.b, '번호': r.n, '이름': r.nm, '강좌명': r.c, '환불_수강료': r.rT, '환불_교재비': r.rB, '실부담_수강료': r.cT, '실부담_교재비': r.cB, '사유_상세': r.mm }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), '환불조정내역'); XLSX.writeFile(wb, `${q}분기_환불_조정_사후증빙용_${new Date().toISOString().slice(0,10)}.xlsx`);
};
window.renderPreviewRef = function() {
    const q = window.gQ;
    const data = E.filter(e => e.q === q && (e.rT>0 || e.rB>0 || (e.adjusts&&e.adjusts.length>0))).map(e => ({ q: e.q, c: e.course, dp: dsp(e.g,e.b,e.n), nm: e.name, g: e.g, b: e.b, n: e.n, rT: e.rT, rB: e.rB, cT: e.cT, cB: e.cB, mm: e.mm }));
    let h = ''; if(!data.length) h = `<tr><td colspan=\"7\" class=\"text-muted py-3\">환불/조정 내역이 없습니다.</td></tr>`;
    else data.forEach(r => h += `<tr><td>${r.q}분기</td><td class=\"fw-bold\">${r.c}</td><td>${r.dp}</td><td>${r.nm}</td><td class=\"text-danger\">${fmt(r.rT)}</td><td class=\"text-danger\">${fmt(r.rB)}</td><td class=\"text-start\" style=\"font-size:0.8rem;\">${r.mm}</td></tr>`);
    if($('prev_ref')) $('prev_ref').innerHTML = h;
};

function getRosterData(q) {
    const tg = val('p_tg'), so = val('p_so');
    let rows = Hs.filter(h => h.q === q).map(h => {
        let t=0, b=0; if (tg==='SELF') {t=h.finT;b=h.finB;} else if (tg==='CHO3') {t=h.tc;b=h.bc;} else if (tg==='FREE') {t=h.tf;b=h.bf;} else {t=h.sT;b=h.sB;}
        return { g: h.g, ban: h.ban, n: h.num, nm: h.nm, c: h.c, t, b, tot: t+b };
    }).filter(x => x.tot > 0);
    rows.sort((a,b) => so==='C' ? a.c.localeCompare(b.c)||a.g-b.g||a.ban-b.ban||a.n-b.n : a.g-b.g||a.ban-b.ban||a.n-b.n||a.nm.localeCompare(b.nm));
    return rows;
}
window.renderPreviewRoster = function() {
    const q = window.gQ; const data = getRosterData(q); let h = '';
    if(!data.length) h = `<tr><td colspan=\"7\" class=\"text-muted py-3\">명단이 없습니다.</td></tr>`;
    else {
        let totT=0, totB=0, totAll=0; data.forEach(r => { totT+=r.t; totB+=r.b; totAll+=r.tot; });
        h += `<tr class=\"sticky-total-row fw-bold text-center\"><td colspan=\"4\" class=\"text-end pe-3\">총 합계</td><td>${fmt(totT)}</td><td>${fmt(totB)}</td><td class=\"text-danger fs-6\">${fmt(totAll)}</td></tr>`;
        data.forEach((r, i) => h += `<tr><td>${i+1}</td><td>${dsp(r.g, r.ban, r.n)}</td><td>${r.nm}</td><td class=\"fw-bold\">${r.c}</td><td>${fmt(r.t)}</td><td>${fmt(r.b)}</td><td class=\"fw-bold text-primary\">${fmt(r.tot)}</td></tr>`);
    }
    if($('prev_ros')) $('prev_ros').innerHTML = h;
};
window.exRoster = function() {
    const q = window.gQ, tg = val('p_tg'), data = getRosterData(q); if (!data.length) return alert('명단이 없습니다.');
    const wb = XLSX.utils.book_new(); const sheetData = data.map((x, i) => ({ '연번': i+1, '학년': x.g, '반': x.ban, '번호': x.n, '이름': x.nm, '과목': x.c, '수강료': x.t, '교재비': x.b, '합계': x.tot }));
    const sumT = sheetData.reduce((sum, x) => sum + x['수강료'], 0), sumB = sheetData.reduce((sum, x) => sum + x['교재비'], 0), sumTot = sheetData.reduce((sum, x) => sum + x['합계'], 0);
    sheetData.push({ '연번': '총계', '학년': '', '반': '', '번호': '', '이름': '', '과목': '', '수강료': sumT, '교재비': sumB, '합계': sumTot });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(sheetData), '정산명단'); XLSX.writeFile(wb, `${q}분기_${tg}_정산명단.xlsx`);
};