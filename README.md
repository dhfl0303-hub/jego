[입고검사일지.html](https://github.com/user-attachments/files/27107237/default.html)
<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, user-scalable=no">
<title>원부재료 입고 검사 일지</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
<style>
* { box-sizing: border-box; margin: 0; padding: 0; -webkit-tap-highlight-color: transparent; }
body { font-family: -apple-system, 'Malgun Gothic', sans-serif; background: #f0f0eb; color: #1a1a1a; font-size: 15px; }

header {
  background: #1a1a1a;
  padding: 14px 20px;
  display: flex;
  align-items: center;
  justify-content: space-between;
  position: sticky;
  top: 0;
  z-index: 50;
}
header h1 { color: #f0a500; font-size: 17px; font-weight: 700; }
.header-btns { display: flex; gap: 8px; }

.btn {
  padding: 9px 16px;
  border-radius: 8px;
  border: none;
  cursor: pointer;
  font-family: inherit;
  font-size: 13px;
  font-weight: 600;
  transition: opacity 0.1s, transform 0.1s;
}
.btn:active { opacity: 0.8; transform: scale(0.97); }
.btn-amber { background: #f0a500; color: #1a1a1a; }
.btn-dark { background: #333; color: #fff; }
.btn-outline { background: transparent; border: 1.5px solid #ccc; color: #444; }
.btn-sm { padding: 7px 12px; font-size: 12px; }
.btn-danger { background: #fee2e2; color: #b91c1c; border: 1.5px solid #fca5a5; }
.btn-green { background: #d1fae5; color: #065f46; border: 1.5px solid #6ee7b7; }

/* ── 기본정보 섹션 ── */
.section {
  background: #fff;
  margin: 12px;
  border-radius: 12px;
  padding: 16px;
  border: 1px solid #e5e5e0;
}
.section-title {
  font-size: 12px;
  font-weight: 700;
  color: #888;
  text-transform: uppercase;
  letter-spacing: 0.5px;
  margin-bottom: 14px;
  padding-bottom: 10px;
  border-bottom: 1px solid #f0f0ec;
}

.info-grid { display: flex; gap: 12px; flex-wrap: wrap; }
.info-group { flex: 1; min-width: 140px; }
.info-group label { display: block; font-size: 11px; font-weight: 700; color: #888; margin-bottom: 5px; }
.info-group input, .info-group select {
  width: 100%;
  padding: 10px 12px;
  border: 1.5px solid #e0e0e0;
  border-radius: 8px;
  font-size: 15px;
  font-family: inherit;
  background: #fafafa;
}
.info-group input:focus, .info-group select:focus { outline: none; border-color: #f0a500; background: #fff; }

/* ── 스캐너 ── */
#scannerPanel {
  display: none;
  background: #1a1a1a;
  padding: 16px;
  margin: 0 12px 12px;
  border-radius: 12px;
}
#scannerPanel.active { display: block; }
#video { width: 100%; max-height: 260px; object-fit: cover; border-radius: 8px; background: #333; display: block; }
.scan-status { color: #f0a500; text-align: center; padding: 8px 0 4px; font-size: 13px; }
.manual-bc { display: none; margin-top: 10px; gap: 8px; }
.manual-bc.show { display: flex; }
.manual-bc input {
  flex: 1; padding: 10px 14px; border-radius: 8px;
  border: 1px solid #555; background: #2a2a2a; color: #fff;
  font-size: 15px; font-family: monospace;
}
.manual-bc input:focus { outline: none; border-color: #f0a500; }

/* ── 품목 폼 ── */
#itemForm {
  display: none;
  background: #fff;
  margin: 0 12px 12px;
  border-radius: 12px;
  padding: 16px;
  border: 2px solid #f0a500;
}
#itemForm.active { display: block; }
#itemForm h3 { font-size: 14px; font-weight: 700; margin-bottom: 14px; color: #555; }

.form-row { display: flex; gap: 10px; flex-wrap: wrap; margin-bottom: 12px; }
.form-group { flex: 1; min-width: 130px; }
.form-group label { display: block; font-size: 11px; font-weight: 700; color: #888; margin-bottom: 5px; }
.form-group input, .form-group select {
  width: 100%;
  padding: 10px 12px;
  border: 1.5px solid #e0e0e0;
  border-radius: 8px;
  font-size: 15px;
  font-family: inherit;
  background: #fafafa;
}
.form-group input:focus, .form-group select:focus { outline: none; border-color: #f0a500; background: #fff; }

/* 관능검사 */
.check-grid {
  display: flex;
  gap: 8px;
  flex-wrap: wrap;
  margin-bottom: 12px;
}
.check-item {
  flex: 1;
  min-width: 80px;
  text-align: center;
}
.check-item label { display: block; font-size: 11px; font-weight: 700; color: #888; margin-bottom: 6px; }
.ox-btns { display: flex; gap: 4px; }
.ox-btn {
  flex: 1;
  padding: 10px 0;
  border-radius: 6px;
  border: 1.5px solid #e0e0e0;
  background: #fafafa;
  cursor: pointer;
  font-size: 16px;
  font-weight: 700;
  transition: all 0.1s;
}
.ox-btn.O { background: #d1fae5; border-color: #6ee7b7; color: #065f46; }
.ox-btn.X { background: #fee2e2; border-color: #fca5a5; color: #b91c1c; }

/* 소비기한 */
.expiry-section { margin-bottom: 12px; }
.expiry-label-row {
  display: flex; align-items: center;
  justify-content: space-between; margin-bottom: 8px;
}
.expiry-title { font-size: 11px; font-weight: 700; color: #888; text-transform: uppercase; }
.add-exp-btn {
  font-size: 12px; background: #fff8e8; border: 1.5px solid #f0a500;
  color: #c8880a; padding: 4px 10px; border-radius: 20px; cursor: pointer; font-weight: 600;
}
.add-exp-btn:disabled { opacity: 0.4; cursor: default; }
.expiry-list { display: flex; flex-direction: column; gap: 6px; }
.expiry-row { display: flex; align-items: center; gap: 8px; }
.expiry-tag {
  font-size: 11px; background: #f5f5f0; color: #666;
  padding: 2px 7px; border-radius: 4px; font-weight: 600; white-space: nowrap;
}
.expiry-date-btn {
  flex: 1; padding: 9px 14px; background: #f9f9f6;
  border: 1.5px solid #e0e0e0; border-radius: 8px;
  text-align: left; cursor: pointer; font-size: 15px;
  font-family: 'Courier New', monospace; font-weight: 600; color: #1a1a1a;
}
.expiry-date-btn.empty { color: #bbb; font-weight: 400; font-family: inherit; font-size: 14px; }
.expiry-date-btn.active { border-color: #f0a500; background: #fff8e8; }
.remove-exp-btn { background: none; border: none; color: #f87171; cursor: pointer; font-size: 20px; padding: 0 4px; }

.date-picker-inline { display: none; background: #fff; border: 1.5px solid #f0a500; border-radius: 10px; padding: 14px; margin-top: 6px; }
.date-picker-inline.show { display: block; }
.picker-selects { display: flex; gap: 8px; margin-bottom: 12px; }
.picker-col { display: flex; flex-direction: column; align-items: center; gap: 5px; flex: 1; }
.picker-col label { font-size: 11px; font-weight: 700; color: #888; }
.picker-select {
  width: 100%; padding: 10px 6px; border: 1.5px solid #e0e0e0; border-radius: 8px;
  background: #fafafa; font-size: 15px; font-family: 'Courier New', monospace;
  font-weight: 600; text-align: center; cursor: pointer; color: #1a1a1a;
  -webkit-appearance: none; appearance: none;
}
.picker-select:focus { outline: none; border-color: #f0a500; background: #fff8e8; }
.picker-btns { display: flex; gap: 8px; justify-content: flex-end; }

/* 서류 체크 */
.doc-checks { display: flex; gap: 8px; flex-wrap: wrap; margin-bottom: 12px; }
.doc-check {
  flex: 1; min-width: 100px; padding: 10px 8px;
  border: 1.5px solid #e0e0e0; border-radius: 8px;
  text-align: center; cursor: pointer; background: #fafafa;
  font-size: 12px; font-weight: 600; color: #888;
  transition: all 0.1s;
}
.doc-check.checked { background: #d1fae5; border-color: #6ee7b7; color: #065f46; }
.doc-check.na { background: #f5f5f0; border-color: #ddd; color: #bbb; }

/* 온도기록지 */
.photo-section { margin-bottom: 12px; }
.photo-preview {
  width: 100%; min-height: 120px; border: 2px dashed #e0e0e0;
  border-radius: 10px; display: flex; align-items: center;
  justify-content: center; cursor: pointer; background: #fafafa;
  overflow: hidden; position: relative;
}
.photo-preview img { width: 100%; height: auto; display: block; border-radius: 8px; }
.photo-placeholder { text-align: center; color: #bbb; padding: 20px; }
.photo-placeholder .icon { font-size: 32px; margin-bottom: 8px; }
.photo-placeholder p { font-size: 13px; }

/* ── 액션바 ── */
.action-bar {
  background: #fff; padding: 12px 16px;
  display: flex; gap: 10px; border-bottom: 1px solid #e5e5e0; flex-wrap: wrap;
  margin: 0 12px 12px; border-radius: 12px;
}

/* ── 품목 카드 ── */
.items-wrap { padding: 0 12px 80px; }
.empty { text-align: center; padding: 60px 20px; color: #aaa; }
.empty .icon { font-size: 40px; margin-bottom: 12px; }

.item-card {
  background: #fff; border-radius: 12px; padding: 16px;
  margin-bottom: 10px; border: 1px solid #e5e5e0; position: relative;
}
.item-header { display: flex; align-items: flex-start; justify-content: space-between; margin-bottom: 12px; }
.item-num { font-size: 11px; color: #aaa; font-weight: 700; margin-bottom: 2px; }
.item-name { font-size: 16px; font-weight: 700; }
.item-meta { font-size: 13px; color: #888; margin-top: 2px; }
.item-del { background: none; border: none; cursor: pointer; color: #ccc; font-size: 20px; padding: 2px; }
.item-del:hover { color: #f87171; }

.item-info-row {
  display: flex; gap: 8px; flex-wrap: wrap; margin-bottom: 10px;
}
.info-badge {
  background: #f5f5f0; border-radius: 6px; padding: 6px 10px;
  font-size: 12px; color: #555;
}
.info-badge span { font-weight: 700; color: #1a1a1a; }

.check-row {
  display: flex; gap: 6px; flex-wrap: wrap; margin-bottom: 10px;
}
.check-badge {
  padding: 4px 10px; border-radius: 20px; font-size: 12px; font-weight: 700;
  border: 1.5px solid;
}
.check-badge.O { background: #d1fae5; border-color: #6ee7b7; color: #065f46; }
.check-badge.X { background: #fee2e2; border-color: #fca5a5; color: #b91c1c; }
.check-badge.none { background: #f5f5f0; border-color: #ddd; color: #bbb; }

.expiry-badges { display: flex; gap: 6px; flex-wrap: wrap; margin-bottom: 10px; }
.expiry-badge { background: #fff8e8; border: 1px solid #f0a500; border-radius: 6px; padding: 4px 10px; font-size: 12px; font-family: monospace; font-weight: 700; color: #c8880a; }

.doc-badges { display: flex; gap: 6px; flex-wrap: wrap; }
.doc-badge { padding: 3px 8px; border-radius: 4px; font-size: 11px; font-weight: 700; }
.doc-badge.yes { background: #d1fae5; color: #065f46; }
.doc-badge.no { background: #f5f5f0; color: #bbb; }

.temp-thumb { width: 60px; height: 45px; object-fit: cover; border-radius: 6px; border: 1px solid #e0e0e0; margin-top: 8px; cursor: pointer; }

/* ── 서명 패널 ── */
#signPanel {
  display: none;
  position: fixed; inset: 0; background: rgba(0,0,0,0.6);
  z-index: 100; align-items: flex-end; justify-content: center;
}
#signPanel.active { display: flex; }
.sign-sheet {
  background: #fff; width: 100%; max-width: 600px;
  border-radius: 20px 20px 0 0; padding: 20px;
}
.sign-sheet h3 { font-size: 16px; font-weight: 700; margin-bottom: 14px; }
.sign-tabs { display: flex; gap: 8px; margin-bottom: 12px; }
.sign-tab {
  flex: 1; padding: 10px; border-radius: 8px; border: 1.5px solid #e0e0e0;
  background: #fafafa; cursor: pointer; font-size: 14px; font-weight: 600;
  text-align: center; color: #888;
}
.sign-tab.active { background: #1a1a1a; color: #f0a500; border-color: #1a1a1a; }
canvas#signCanvas {
  width: 100%; height: 180px; border: 2px dashed #e0e0e0;
  border-radius: 10px; background: #fafafa; display: block; touch-action: none;
}
.sign-btns { display: flex; gap: 8px; margin-top: 12px; }

/* ── 토스트 ── */
.toast {
  position: fixed; bottom: 24px; left: 50%; transform: translateX(-50%);
  background: #1a1a1a; color: #fff; padding: 12px 22px;
  border-radius: 24px; font-size: 14px; font-weight: 600;
  z-index: 9999; animation: toastIn 0.25s ease; white-space: nowrap;
}
@keyframes toastIn {
  from { opacity: 0; transform: translateX(-50%) translateY(10px); }
  to   { opacity: 1; transform: translateX(-50%) translateY(0); }
}

/* 이미지 모달 */
#imgModal {
  display: none; position: fixed; inset: 0; background: rgba(0,0,0,0.85);
  z-index: 200; align-items: center; justify-content: center;
}
#imgModal.active { display: flex; }
#imgModal img { max-width: 95%; max-height: 90vh; border-radius: 10px; }
#imgModal .close { position: absolute; top: 20px; right: 20px; color: #fff; font-size: 30px; cursor: pointer; }

@media (max-width: 480px) {
  .form-row { flex-direction: column; }
  .info-grid { flex-direction: column; }
}
</style>
</head>
<body>

<header>
  <h1>📋 입고 검사 일지</h1>
  <div class="header-btns">
    <button class="btn btn-dark btn-sm" onclick="openSignPanel()">서명</button>
    <button class="btn btn-amber btn-sm" onclick="exportPDF()">PDF</button>
  </div>
</header>

<!-- 기본 정보 -->
<div class="section" style="margin:12px;">
  <div class="section-title">기본 정보</div>
  <div class="info-grid">
    <div class="info-group">
      <label>입고일자</label>
      <input type="date" id="inDate">
    </div>
    <div class="info-group">
      <label>입고검사자</label>
      <input type="text" id="inspector" placeholder="이름 입력">
    </div>
  </div>
</div>

<!-- 스캐너 -->
<div id="scannerPanel">
  <video id="video" autoplay playsinline muted></video>
  <div class="scan-status" id="scanStatus">카메라 시작 중...</div>
  <div class="manual-bc" id="manualBcRow">
    <input type="text" id="manualBcInput" placeholder="바코드 번호 직접 입력" inputmode="numeric">
    <button class="btn btn-amber" onclick="processManualBarcode()">확인</button>
    <button class="btn btn-outline" onclick="closeScanner()">취소</button>
  </div>
</div>

<!-- 품목 폼 -->
<div id="itemForm">
  <h3 id="formTitle">품목 추가</h3>

  <div class="form-row">
    <div class="form-group">
      <label>거래처</label>
      <select id="fCompany">
        <option value="">선택</option>
        <option>에이스유통</option>
        <option>엠즈푸드</option>
        <option>덕풍</option>
        <option>내추럴쿡</option>
        <option>남원원예농협</option>
        <option>SPC</option>
        <option>밤마을</option>
        <option>맑은들</option>
        <option>제원인터내쇼날</option>
        <option>하나푸드</option>
        <option>CJ제일제당</option>
        <option>대선제분</option>
      </select>
    </div>
    <div class="form-group">
      <label>품명</label>
      <input type="text" id="fName" placeholder="예: 버터">
    </div>
  </div>

  <div class="form-row">
    <div class="form-group">
      <label>입고량 (kg)</label>
      <input type="number" id="fQty" placeholder="0.000" step="0.001" min="0">
    </div>
    <div class="form-group">
      <label>제품온도 (℃)</label>
      <input type="number" id="fTemp" placeholder="예: 4.5" step="0.1">
    </div>
  </div>

  <!-- 소비기한 -->
  <div class="expiry-section">
    <div class="expiry-label-row">
      <span class="expiry-title">소비기한</span>
      <button class="add-exp-btn" id="addExpBtn" onclick="addFormExpiry()">+ 추가</button>
    </div>
    <div class="expiry-list" id="formExpiryList"></div>
  </div>

  <!-- 관능검사 -->
  <div class="section-title" style="margin-bottom:10px;">관능검사 (O/X)</div>
  <div class="check-grid">
    <div class="check-item">
      <label>성상</label>
      <div class="ox-btns">
        <button class="ox-btn" id="c_ss_O" onclick="setCheck('ss','O')">O</button>
        <button class="ox-btn" id="c_ss_X" onclick="setCheck('ss','X')">X</button>
      </div>
    </div>
    <div class="check-item">
      <label>표시사항</label>
      <div class="ox-btns">
        <button class="ox-btn" id="c_lb_O" onclick="setCheck('lb','O')">O</button>
        <button class="ox-btn" id="c_lb_X" onclick="setCheck('lb','X')">X</button>
      </div>
    </div>
    <div class="check-item">
      <label>포장상태</label>
      <div class="ox-btns">
        <button class="ox-btn" id="c_pk_O" onclick="setCheck('pk','O')">O</button>
        <button class="ox-btn" id="c_pk_X" onclick="setCheck('pk','X')">X</button>
      </div>
    </div>
    <div class="check-item">
      <label>차량위생</label>
      <div class="ox-btns">
        <button class="ox-btn" id="c_cr_O" onclick="setCheck('cr','O')">O</button>
        <button class="ox-btn" id="c_cr_X" onclick="setCheck('cr','X')">X</button>
      </div>
    </div>
    <div class="check-item">
      <label>성적서</label>
      <div class="ox-btns">
        <button class="ox-btn" id="c_ct_O" onclick="setCheck('ct','O')">O</button>
        <button class="ox-btn" id="c_ct_X" onclick="setCheck('ct','X')">X</button>
      </div>
    </div>
  </div>

  <!-- 서류 체크 -->
  <div class="section-title" style="margin-bottom:10px;">첨부서류</div>
  <div class="doc-checks">
    <div class="doc-check" id="doc_taco" onclick="toggleDoc('taco')">타코메타<br>기록지</div>
    <div class="doc-check" id="doc_cert" onclick="toggleDoc('cert')">시험<br>성적서</div>
    <div class="doc-check" id="doc_imp" onclick="toggleDoc('imp')">수입<br>면장</div>
    <div class="doc-check" id="doc_dist" onclick="toggleDoc('dist')">구분유통<br>증명서</div>
    <div class="doc-check" id="doc_ori" onclick="toggleDoc('ori')">원산지<br>증명원</div>
  </div>

  <!-- 온도기록지 사진 -->
  <div class="photo-section">
    <div class="section-title" style="margin-bottom:10px;">온도기록지 사진</div>
    <div class="photo-preview" id="photoPreview" onclick="document.getElementById('photoInput').click()">
      <div class="photo-placeholder">
        <div class="icon">📷</div>
        <p>터치하여 사진 촬영</p>
      </div>
    </div>
    <input type="file" id="photoInput" accept="image/*" capture="environment" style="display:none" onchange="handlePhoto(event)">
  </div>

  <div style="display:flex;gap:8px;">
    <button class="btn btn-amber" style="flex:1" onclick="submitItem()">추가</button>
    <button class="btn btn-outline" onclick="closeForm()">취소</button>
  </div>
</div>

<!-- 액션바 -->
<div class="action-bar">
  <button class="btn btn-amber" onclick="openScanner()">📷 바코드 스캔</button>
  <button class="btn btn-dark" onclick="openForm()">✏️ 수동 추가</button>
</div>

<!-- 품목 목록 -->
<div class="items-wrap">
  <div class="empty" id="emptyMsg">
    <div class="icon">📦</div>
    <p>바코드를 스캔하거나<br>수동으로 품목을 추가해주세요</p>
  </div>
  <div id="itemsList"></div>
</div>

<!-- 서명 패널 -->
<div id="signPanel">
  <div class="sign-sheet">
    <h3>결재 서명</h3>
    <div class="sign-tabs">
      <div class="sign-tab active" id="tab_write" onclick="switchTab('write')">작성자</div>
      <div class="sign-tab" id="tab_approve" onclick="switchTab('approve')">승인자</div>
    </div>
    <canvas id="signCanvas"></canvas>
    <div class="sign-btns">
      <button class="btn btn-outline btn-sm" onclick="clearSign()">지우기</button>
      <button class="btn btn-dark btn-sm" style="flex:1" onclick="saveSign()">저장</button>
      <button class="btn btn-amber btn-sm" onclick="closeSignPanel()">닫기</button>
    </div>
  </div>
</div>

<!-- 이미지 모달 -->
<div id="imgModal" onclick="closeImgModal()">
  <span class="close">×</span>
  <img id="modalImg" src="">
</div>

<script>
// ── 상태 ──
let items = [];
let nextId = 1;
let formExpiries = [null];
let formOpenPicker = null;
let formChecks = { ss: null, lb: null, pk: null, cr: null, ct: null };
let formDocs = { taco: false, cert: false, imp: false, dist: false, ori: false };
let formPhoto = null;
let pendingBarcode = null;
let videoStream = null;
let scanTimer = null;
let currentSignTab = 'write';
let signs = { write: null, approve: null };
let signDrawing = false;
let signLastX = 0, signLastY = 0;

// 매핑 DB
let mappingDB = JSON.parse(localStorage.getItem('inventoryDB') || '{}');
function saveDB() { localStorage.setItem('inventoryDB', JSON.stringify(mappingDB)); }

// 날짜 기본값
const today = new Date();
document.getElementById('inDate').value = today.toISOString().split('T')[0];

// ── 스캐너 ──
function openScanner() {
  closeForm();
  document.getElementById('scannerPanel').classList.add('active');
  const status = document.getElementById('scanStatus');
  const manualRow = document.getElementById('manualBcRow');

  if (!('BarcodeDetector' in window)) {
    status.textContent = '자동 스캔 미지원 — 번호를 직접 입력해주세요';
    manualRow.classList.add('show');
    return;
  }

  navigator.mediaDevices.getUserMedia({ video: { facingMode: { ideal: 'environment' } } })
    .then(stream => {
      videoStream = stream;
      const video = document.getElementById('video');
      video.srcObject = stream;
      status.textContent = '바코드를 카메라에 비춰주세요...';
      const detector = new BarcodeDetector({ formats: ['ean_13','ean_8','code_128','code_39','upc_a','upc_e'] });
      scanTimer = setInterval(async () => {
        if (video.readyState < 2) return;
        try {
          const results = await detector.detect(video);
          if (results.length > 0) {
            stopCamera(); closeScanner();
            processBarcode(results[0].rawValue);
          }
        } catch(e) {}
      }, 400);
    })
    .catch(() => {
      status.textContent = '카메라 접근 실패 — 번호를 직접 입력해주세요';
      manualRow.classList.add('show');
    });
}

function stopCamera() {
  if (scanTimer) { clearInterval(scanTimer); scanTimer = null; }
  if (videoStream) { videoStream.getTracks().forEach(t => t.stop()); videoStream = null; }
}

function closeScanner() {
  stopCamera();
  document.getElementById('scannerPanel').classList.remove('active');
  document.getElementById('manualBcRow').classList.remove('show');
  document.getElementById('manualBcInput').value = '';
}

function processManualBarcode() {
  const val = document.getElementById('manualBcInput').value.trim();
  if (!val) return;
  closeScanner();
  processBarcode(val);
}

function processBarcode(barcode) {
  if (mappingDB[barcode]) {
    const p = mappingDB[barcode];
    openForm(barcode, p.name, p.origin || '');
    toast('✅ ' + p.name + ' 자동입력됨');
  } else {
    pendingBarcode = barcode;
    openForm(barcode, '', '');
    document.getElementById('formTitle').textContent = '품목 추가 (미등록: ' + barcode + ')';
  }
}

// ── 폼 ──
function resetForm() {
  formExpiries = [null];
  formOpenPicker = null;
  formChecks = { ss: null, lb: null, pk: null, cr: null, ct: null };
  formDocs = { taco: false, cert: false, imp: false, dist: false, ori: false };
  formPhoto = null;
  document.getElementById('fCompany').value = '';
  document.getElementById('fName').value = '';
  document.getElementById('fQty').value = '';
  document.getElementById('fTemp').value = '';
  document.getElementById('photoPreview').innerHTML = '<div class="photo-placeholder"><div class="icon">📷</div><p>터치하여 사진 촬영</p></div>';
  ['ss','lb','pk','cr','ct'].forEach(k => {
    document.getElementById('c_'+k+'_O').className = 'ox-btn';
    document.getElementById('c_'+k+'_X').className = 'ox-btn';
  });
  Object.keys(formDocs).forEach(k => document.getElementById('doc_'+k).className = 'doc-check');
  renderFormExpiries();
}

function openForm(barcode, name, origin) {
  closeScanner();
  resetForm();
  pendingBarcode = barcode || null;
  document.getElementById('formTitle').textContent = '품목 추가';
  if (name) document.getElementById('fName').value = name;
  document.getElementById('itemForm').classList.add('active');
  document.getElementById('itemForm').scrollIntoView({ behavior: 'smooth' });
}

function closeForm() {
  document.getElementById('itemForm').classList.remove('active');
  pendingBarcode = null;
}

function setCheck(key, val) {
  formChecks[key] = formChecks[key] === val ? null : val;
  ['O','X'].forEach(v => {
    document.getElementById('c_'+key+'_'+v).className = 'ox-btn' + (formChecks[key] === v ? ' '+v : '');
  });
}

function toggleDoc(key) {
  formDocs[key] = !formDocs[key];
  document.getElementById('doc_'+key).className = 'doc-check' + (formDocs[key] ? ' checked' : '');
}

function handlePhoto(e) {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = ev => {
    formPhoto = ev.target.result;
    document.getElementById('photoPreview').innerHTML = '<img src="'+formPhoto+'">';
  };
  reader.readAsDataURL(file);
}

// 소비기한 (폼)
function addFormExpiry() {
  if (formExpiries.length >= 3) return;
  formExpiries.push(null);
  renderFormExpiries();
}

function removeFormExpiry(idx) {
  formExpiries.splice(idx, 1);
  if (formOpenPicker === idx) formOpenPicker = null;
  renderFormExpiries();
}

function toggleFormPicker(idx) {
  formOpenPicker = (formOpenPicker === idx) ? null : idx;
  renderFormExpiries();
}

function confirmFormDate(idx) {
  const key = 'fp_' + idx;
  formExpiries[idx] = {
    year:  parseInt(document.getElementById('fy_'+idx).value),
    month: parseInt(document.getElementById('fm_'+idx).value),
    day:   parseInt(document.getElementById('fd_'+idx).value)
  };
  formOpenPicker = null;
  renderFormExpiries();
}

function renderFormExpiries() {
  const now = new Date();
  const curY = now.getFullYear(), curM = now.getMonth()+1, curD = now.getDate();
  const list = document.getElementById('formExpiryList');
  list.innerHTML = formExpiries.map((exp, i) => {
    const isOpen = formOpenPicker === i;
    const dateText = exp ? `${exp.year}-${String(exp.month).padStart(2,'0')}-${String(exp.day).padStart(2,'0')}` : null;
    const yearOpts = Array.from({length:12},(_,j)=>{const y=curY-1+j;return`<option value="${y}"${(exp?.year===y||(!exp&&y===curY))?' selected':''}>${y}년</option>`;}).join('');
    const monthOpts = Array.from({length:12},(_,j)=>{const m=j+1;return`<option value="${m}"${(exp?.month===m||(!exp&&m===curM))?' selected':''}>${String(m).padStart(2,'0')}월</option>`;}).join('');
    const dayOpts = Array.from({length:31},(_,j)=>{const d=j+1;return`<option value="${d}"${(exp?.day===d||(!exp&&d===curD))?' selected':''}>${String(d).padStart(2,'0')}일</option>`;}).join('');
    return `
      <div class="expiry-row">
        <span class="expiry-tag">기한 ${i+1}</span>
        <button class="expiry-date-btn ${dateText?'':'empty'} ${isOpen?'active':''}" onclick="toggleFormPicker(${i})">${dateText||'날짜 선택'}</button>
        ${i>0?`<button class="remove-exp-btn" onclick="removeFormExpiry(${i})">×</button>`:''}
      </div>
      <div class="date-picker-inline ${isOpen?'show':''}">
        <div class="picker-selects">
          <div class="picker-col"><label>년도</label><select class="picker-select" id="fy_${i}">${yearOpts}</select></div>
          <div class="picker-col"><label>월</label><select class="picker-select" id="fm_${i}">${monthOpts}</select></div>
          <div class="picker-col"><label>일</label><select class="picker-select" id="fd_${i}">${dayOpts}</select></div>
        </div>
        <div class="picker-btns">
          <button class="btn btn-outline btn-sm" onclick="formOpenPicker=null;renderFormExpiries()">취소</button>
          <button class="btn btn-amber btn-sm" onclick="confirmFormDate(${i})">확인</button>
        </div>
      </div>`;
  }).join('');
  document.getElementById('addExpBtn').disabled = formExpiries.length >= 3;
}

function submitItem() {
  const company = document.getElementById('fCompany').value;
  const name = document.getElementById('fName').value.trim();
  const qty = document.getElementById('fQty').value;
  const temp = document.getElementById('fTemp').value;

  if (!name) { alert('품명을 입력해주세요.'); return; }

  if (pendingBarcode && name) {
    mappingDB[pendingBarcode] = { name, origin: '' };
    saveDB();
  }

  items.push({
    id: nextId++,
    barcode: pendingBarcode,
    company, name, qty, temp,
    expiries: [...formExpiries],
    checks: {...formChecks},
    docs: {...formDocs},
    photo: formPhoto
  });

  closeForm();
  renderItems();
  toast('✅ ' + name + ' 추가됨');
}

// ── 렌더 ──
function renderItems() {
  const empty = document.getElementById('emptyMsg');
  const list = document.getElementById('itemsList');
  if (!items.length) { empty.style.display='block'; list.innerHTML=''; return; }
  empty.style.display = 'none';

  const checkLabels = { ss:'성상', lb:'표시사항', pk:'포장상태', cr:'차량위생', ct:'성적서' };
  const docLabels = { taco:'타코메타', cert:'시험성적서', imp:'수입면장', dist:'구분유통증명', ori:'원산지증명' };

  list.innerHTML = items.map((item, idx) => {
    const expiryBadges = item.expiries.filter(Boolean).map(e =>
      `<span class="expiry-badge">${e.year}-${String(e.month).padStart(2,'0')}-${String(e.day).padStart(2,'0')}</span>`
    ).join('');

    const checkBadges = Object.entries(checkLabels).map(([k,l]) =>
      `<span class="check-badge ${item.checks[k]||'none'}">${l} ${item.checks[k]||'-'}</span>`
    ).join('');

    const docBadges = Object.entries(docLabels).map(([k,l]) =>
      item.docs[k] ? `<span class="doc-badge yes">✓ ${l}</span>` : ''
    ).filter(Boolean).join('');

    return `
      <div class="item-card">
        <button class="item-del" onclick="removeItem(${item.id})">×</button>
        <div class="item-header">
          <div>
            <div class="item-num">NO. ${String(idx+1).padStart(2,'0')} ${item.barcode?'· '+item.barcode:''}</div>
            <div class="item-name">${item.name}</div>
            <div class="item-meta">${item.company||'-'}</div>
          </div>
        </div>
        <div class="item-info-row">
          <div class="info-badge">입고량 <span>${item.qty||'-'} kg</span></div>
          <div class="info-badge">온도 <span>${item.temp||'-'} ℃</span></div>
        </div>
        ${expiryBadges ? `<div class="expiry-badges">${expiryBadges}</div>` : ''}
        <div class="check-row">${checkBadges}</div>
        ${docBadges ? `<div class="doc-badges">${docBadges}</div>` : ''}
        ${item.photo ? `<img class="temp-thumb" src="${item.photo}" onclick="openImgModal('${item.photo}')">` : ''}
      </div>`;
  }).join('');
}

function removeItem(id) {
  items = items.filter(i => i.id !== id);
  renderItems();
}

// ── 서명 ──
function openSignPanel() {
  document.getElementById('signPanel').classList.add('active');
  initCanvas();
}
function closeSignPanel() { document.getElementById('signPanel').classList.remove('active'); }

function switchTab(tab) {
  currentSignTab = tab;
  ['write','approve'].forEach(t => {
    document.getElementById('tab_'+t).className = 'sign-tab' + (t===tab?' active':'');
  });
  clearSign();
}

function initCanvas() {
  const canvas = document.getElementById('signCanvas');
  const rect = canvas.getBoundingClientRect();
  canvas.width = rect.width * window.devicePixelRatio;
  canvas.height = rect.height * window.devicePixelRatio;
  const ctx = canvas.getContext('2d');
  ctx.scale(window.devicePixelRatio, window.devicePixelRatio);
  ctx.strokeStyle = '#1a1a1a';
  ctx.lineWidth = 2.5;
  ctx.lineCap = 'round';

  if (signs[currentSignTab]) {
    const img = new Image();
    img.onload = () => ctx.drawImage(img, 0, 0, rect.width, rect.height);
    img.src = signs[currentSignTab];
  }

  canvas.onpointerdown = e => {
    signDrawing = true;
    const r = canvas.getBoundingClientRect();
    signLastX = e.clientX - r.left;
    signLastY = e.clientY - r.top;
  };
  canvas.onpointermove = e => {
    if (!signDrawing) return;
    const r = canvas.getBoundingClientRect();
    const x = e.clientX - r.left, y = e.clientY - r.top;
    ctx.beginPath();
    ctx.moveTo(signLastX, signLastY);
    ctx.lineTo(x, y);
    ctx.stroke();
    signLastX = x; signLastY = y;
  };
  canvas.onpointerup = () => signDrawing = false;
}

function clearSign() {
  const canvas = document.getElementById('signCanvas');
  const ctx = canvas.getContext('2d');
  ctx.clearRect(0, 0, canvas.width, canvas.height);
}

function saveSign() {
  const canvas = document.getElementById('signCanvas');
  signs[currentSignTab] = canvas.toDataURL();
  toast(currentSignTab === 'write' ? '✅ 작성자 서명 저장' : '✅ 승인자 서명 저장');
}

// ── 이미지 모달 ──
function openImgModal(src) {
  document.getElementById('modalImg').src = src;
  document.getElementById('imgModal').classList.add('active');
}
function closeImgModal() { document.getElementById('imgModal').classList.remove('active'); }

// ── PDF 출력 ──
function exportPDF() {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });

  const inDate = document.getElementById('inDate').value;
  const inspector = document.getElementById('inspector').value || '-';

  // 폰트 설정 (기본 폰트 사용)
  doc.setFont('helvetica');

  // 제목
  doc.setFontSize(16);
  doc.setFont('helvetica', 'bold');
  doc.text('원부재료 입고 검사 일지', 105, 20, { align: 'center' });

  // 기본정보
  doc.setFontSize(10);
  doc.setFont('helvetica', 'normal');
  doc.text('입고일자: ' + inDate, 20, 32);
  doc.text('입고검사자: ' + inspector, 120, 32);

  // 구분선
  doc.setLineWidth(0.5);
  doc.line(15, 36, 195, 36);

  // 테이블 헤더
  let y = 42;
  doc.setFontSize(8);
  doc.setFont('helvetica', 'bold');
  doc.setFillColor(240, 240, 235);
  doc.rect(15, y-4, 180, 8, 'F');
  doc.text('거래처', 17, y);
  doc.text('품명', 45, y);
  doc.text('입고량', 85, y);
  doc.text('소비기한', 105, y);
  doc.text('성상', 135, y);
  doc.text('표시', 143, y);
  doc.text('포장', 151, y);
  doc.text('차량', 159, y);
  doc.text('성적서', 167, y);
  doc.text('온도', 180, y);

  doc.line(15, y+2, 195, y+2);
  y += 8;

  // 품목 목록
  doc.setFont('helvetica', 'normal');
  const checkLabels = { ss:'성상', lb:'표시', pk:'포장', cr:'차량', ct:'성적서' };

  items.forEach((item, idx) => {
    if (y > 240) {
      doc.addPage();
      y = 20;
    }

    doc.setFontSize(8);
    const expText = item.expiries.filter(Boolean).map(e =>
      `${e.year}-${String(e.month).padStart(2,'0')}-${String(e.day).padStart(2,'0')}`
    ).join(', ');

    doc.text((item.company||'-').substring(0,8), 17, y);
    doc.text((item.name||'-').substring(0,12), 45, y);
    doc.text((item.qty||'-')+'kg', 85, y);
    doc.text(expText.substring(0,20)||'-', 105, y);
    doc.text(item.checks.ss||'-', 136, y);
    doc.text(item.checks.lb||'-', 144, y);
    doc.text(item.checks.pk||'-', 152, y);
    doc.text(item.checks.cr||'-', 160, y);
    doc.text(item.checks.ct||'-', 168, y);
    doc.text((item.temp||'-')+'℃', 180, y);

    doc.setDrawColor(220,220,215);
    doc.line(15, y+2, 195, y+2);
    y += 8;
  });

  y += 10;

  // 온도기록지 사진
  const photoItems = items.filter(i => i.photo);
  if (photoItems.length > 0) {
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(10);
    doc.text('온도기록지', 15, y);
    y += 6;

    photoItems.forEach(item => {
      if (y > 220) { doc.addPage(); y = 20; }
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(8);
      doc.text(item.name + ' (' + (item.company||'-') + ')', 15, y);
      y += 4;
      try {
        doc.addImage(item.photo, 'JPEG', 15, y, 80, 50);
        y += 56;
      } catch(e) { y += 4; }
    });
  }

  // 서명
  y += 6;
  if (y > 250) { doc.addPage(); y = 20; }
  doc.setLineWidth(0.5);
  doc.line(15, y, 195, y);
  y += 8;

  doc.setFont('helvetica', 'bold');
  doc.setFontSize(10);
  doc.text('결재', 15, y);
  y += 6;

  doc.setFontSize(8);
  doc.setFont('helvetica', 'normal');
  doc.text('작성자', 30, y);
  doc.text('승인자', 120, y);
  y += 4;

  if (signs.write) {
    try { doc.addImage(signs.write, 'PNG', 15, y, 50, 20); } catch(e) {}
  }
  if (signs.approve) {
    try { doc.addImage(signs.approve, 'PNG', 105, y, 50, 20); } catch(e) {}
  }

  const ds = inDate.replace(/-/g,'');
  doc.save('입고검사일지_' + ds + '.pdf');
  toast('📥 PDF 저장 완료');
}

// ── 알림 ──
function toast(msg) {
  const el = document.createElement('div');
  el.className = 'toast';
  el.textContent = msg;
  document.body.appendChild(el);
  setTimeout(() => el.remove(), 2500);
}

renderFormExpiries();
renderItems();
</script>
</body>
</html>
