import { GoogleGenerativeAI } from 'https://esm.run/@google/generative-ai@0.21.0';
import * as docx from 'docx';

// State variables
let uploadedImages = []; // Array of { name, base64, type }
let extractedData = null; // Store parsed JSON results

// Elements
const apiKeyInput = document.getElementById('api-key');
const saveKeyCheckbox = document.getElementById('save-key');
const patientNameInput = document.getElementById('patient-name');
const patientIdInput = document.getElementById('patient-id');
const patientSeqInput = document.getElementById('patient-seq');
const debondingDateInput = document.getElementById('debonding-date');

const uploadZone = document.getElementById('upload-zone');
const imageInput = document.getElementById('image-input');
const previewGrid = document.getElementById('preview-grid');
const chartTextArea = document.getElementById('chart-text');

const btnGenerate = document.getElementById('btn-generate');
const btnDownload = document.getElementById('btn-download');
const progressLog = document.getElementById('progress-log');
const progressText = document.getElementById('progress-text');

const btnExtractDates = document.getElementById('btn-extract-dates');
const dateReviewSection = document.getElementById('date-review-section');
const dateReviewDesc = document.getElementById('date-review-desc');
const dateListContainer = document.getElementById('date-list');
const btnAddDate = document.getElementById('btn-add-date');

// 1단계에서 추출되고 사용자가 확인/수정한 날짜 목록 (2단계 DB sheet 생성에 그대로 사용됨)
let confirmedDates = [];

const tabButtons = document.querySelectorAll('.tab-btn');
const tabContents = document.querySelectorAll('.tab-content');
const previewEmpty = document.getElementById('preview-empty');
const docPreviewContainer = document.getElementById('doc-preview-container');
const prevHeaderDate = document.getElementById('prev-header-date');
const prevHeaderSeq = document.getElementById('prev-header-seq');
const prevHeaderName = document.getElementById('prev-header-name');
const prevHeaderId = document.getElementById('prev-header-id');
const dbSheetTableBody = document.getElementById('db-sheet-table-body');
const jsonOutput = document.getElementById('json-output');

// =====================================================================================
// 치식 필드 정규화 — 2/4/6개가 아닌 예외적인 pipe 개수에 대한 안전장치
//
// 정식 포맷:
//  - 2개 필드: 화살표/편측 악궁 표기 ({치식:~|~} 등)
//  - 4개 필드: 4분악 — UR|UL|LR|LL
//  - 6개 필드: 6분악 — UR구치|U전치|UL구치|LR구치|L전치|LL구치
// 그 외의 개수(AI가 pipe를 잘못 붙이는 등 정말 예외적인 경우)만 이 함수를 거쳐
// 4분악 형태로 최대한 합리적으로 매핑합니다.
// =====================================================================================
function normalizeToothParts(arr) {
  if (arr.length < 4) {
    var padded = arr.slice();
    while (padded.length < 4) padded.push('');
    return padded;
  }

  // 4개 초과(그리고 6개가 아닌 경우): 값이 있는 위치를 4개 구역에 비례 배분
  var result = ['', '', '', ''];
  arr.forEach(function (v, i) {
    if (v === '') return;
    var bucket = Math.min(3, Math.floor((i / (arr.length - 1)) * 4));
    result[bucket] = result[bucket] ? result[bucket] + v : v;
  });
  return result;
}

// =====================================================================================
// 치식 태그 공용 파서 — HTML 미리보기(convertToPalmer)와 워드 생성(buildRunsFromLine)이
// 동일한 판정 로직을 공유하도록 { 치식: ... } 내부를 구조화된 형태로 파싱합니다.
//
// 반환값:
//  - { type: 'arrow2', parts: [a, b] } : 2-part 화살표/편측 악궁 표기
//  - { type: 'segments', upperParts, lowerParts, hasUpper, hasLower } : 4분악/6분악 표기
// =====================================================================================
function parseToothTag(inner) {
  var rawParts = inner.split('|').map(function (s) { return s.trim(); });

  if (rawParts.length === 2) {
    return { type: 'arrow2', parts: rawParts };
  }

  // 4분악(4개), 6분악(6개)은 그대로 사용. 그 외 개수만 4분악으로 정규화.
  var parts = (rawParts.length === 4 || rawParts.length === 6)
    ? rawParts
    : normalizeToothParts(rawParts);

  var half = parts.length / 2;
  var upperParts = parts.slice(0, half);
  var lowerParts = parts.slice(half);
  var hasUpper = upperParts.some(function (v) { return v !== ''; });
  var hasLower = lowerParts.some(function (v) { return v !== ''; });

  return { type: 'segments', upperParts: upperParts, lowerParts: lowerParts, hasUpper: hasUpper, hasLower: hasLower };
}

// =====================================================================================
// 치식 표기 변환 (HTML 미리보기 전용)
// EMR 텍스트의 {치식:...} 포맷을 화면 표시용 문자열로 변환합니다.
// ※ 워드(docx) 생성 로직은 더 이상 이 함수를 거치지 않고, 아래 buildRunsFromLine()에서
//    parseToothTag()로 원본을 직접 파싱해 docx.Math(MathFraction) 객체로 만듭니다.
// =====================================================================================
function convertToPalmer(text) {
  if (!text) return text;

  return text.replace(/\{치식:([^}]*)\}/g, function (match, inner) {
    var parsed = parseToothTag(inner);

    if (parsed.type === 'arrow2') {
      var a = parsed.parts[0], b = parsed.parts[1];
      // 화살표 케이스(~)는 분수 표기로: ⇧→분자 ~|~, ⇩→분모 ~|~, ⇧⇩→분자·분모 모두 ~|~
      if (a === '~' && b === '~') return "\\underline{~|~} \\overline{~|~}";
      if (a === '~') return "\\underline{~|~}";
      if (b === '~') return "\\overline{~|~}";
      // 화살표가 아닌 일반 텍스트 표기는 그대로 유지
      if (a && !b) return "상악 " + a;
      if (!a && b) return "하악 " + b;
      return "상악 " + a + " / 하악 " + b;
    }

    var upperParts = parsed.upperParts, lowerParts = parsed.lowerParts;
    var hasUpper = parsed.hasUpper, hasLower = parsed.hasLower;

    if (!hasUpper && !hasLower) return '';

    var upperText = upperParts.map(function (v) { return v || ' '; }).join('|');
    var lowerText = lowerParts.map(function (v) { return v || ' '; }).join('|');

    if (hasUpper && hasLower) {
      // 상악·하악 둘 다 있으면 분수(가로선) 표기
      return "\\underline{" + upperText + "} \\overline{" + lowerText + "}";
    } else if (hasUpper) {
      // 상악만 있으면 밑줄(underline)만
      return "\\underline{" + upperText + "}";
    } else {
      // 하악만 있으면 윗줄(overline)만
      return "\\overline{" + lowerText + "}";
    }
  });
}

// Load API Key from localStorage if present
document.addEventListener('DOMContentLoaded', () => {
  const savedKey = localStorage.getItem('ortho_gemini_api_key');
  if (savedKey) {
    apiKeyInput.value = savedKey;
  }
});

// Tab Switcher
tabButtons.forEach(btn => {
  btn.addEventListener('click', () => {
    tabButtons.forEach(b => b.classList.remove('active'));
    tabContents.forEach(c => c.classList.remove('active'));

    btn.classList.add('active');
    const targetTab = btn.getAttribute('data-tab');
    document.getElementById(targetTab).classList.add('active');
  });
});

// Drag and drop event listeners
uploadZone.addEventListener('dragover', (e) => {
  e.preventDefault();
  uploadZone.classList.add('dragover');
});

uploadZone.addEventListener('dragleave', () => {
  uploadZone.classList.remove('dragover');
});

uploadZone.addEventListener('drop', (e) => {
  e.preventDefault();
  uploadZone.classList.remove('dragover');
  const files = e.dataTransfer.files;
  handleImageFiles(files);
});

imageInput.addEventListener('change', (e) => {
  const files = e.target.files;
  if (files && files.length > 0) {
    handleImageFiles(files);
  }
  imageInput.value = '';
});

// Process and compress uploaded image files
function handleImageFiles(files) {
  Array.from(files).forEach(file => {
    if (!file.type.startsWith('image/')) {
      alert('이미지 파일만 업로드할 수 있습니다.');
      return;
    }

    const reader = new FileReader();
    reader.onload = (e) => {
      const img = new Image();
      img.onload = () => {
        const canvas = document.createElement('canvas');
        let width = img.width;
        let height = img.height;

        const MAX_WIDTH = 1200;
        const MAX_HEIGHT = 1200;

        if (width > height) {
          if (width > MAX_WIDTH) {
            height *= MAX_WIDTH / width;
            width = MAX_WIDTH;
          }
        } else {
          if (height > MAX_HEIGHT) {
            width *= MAX_HEIGHT / height;
            height = MAX_HEIGHT;
          }
        }

        canvas.width = width;
        canvas.height = height;

        const ctx = canvas.getContext('2d');
        ctx.drawImage(img, 0, 0, width, height);

        const dataUrl = canvas.toDataURL('image/jpeg', 0.7);
        const base64 = dataUrl.split(',')[1];

        const imageItem = {
          id: Date.now() + Math.random().toString(36).substr(2, 5),
          name: file.name,
          base64: base64,
          type: 'image/jpeg',
          dataUrl: dataUrl
        };
        uploadedImages.push(imageItem);
        renderImagePreviews();
      };
      img.src = e.target.result;
    };
    reader.readAsDataURL(file);
  });
}

// Render image thumbnails in preview grid
function renderImagePreviews() {
  previewGrid.innerHTML = '';
  uploadedImages.forEach((img) => {
    const card = document.createElement('div');
    card.className = 'preview-card';

    const imgEl = document.createElement('img');
    imgEl.src = img.dataUrl;
    imgEl.alt = img.name;

    const nameEl = document.createElement('div');
    nameEl.className = 'image-name';
    nameEl.innerText = img.name;

    const removeBtn = document.createElement('button');
    removeBtn.className = 'btn-remove-img';
    removeBtn.innerHTML = '&times;';
    removeBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      uploadedImages = uploadedImages.filter(item => item.id !== img.id);
      renderImagePreviews();
    });

    card.appendChild(imgEl);
    card.appendChild(nameEl);
    card.appendChild(removeBtn);
    previewGrid.appendChild(card);
  });
}

// Update UI on metadata input changes
[patientNameInput, patientIdInput, patientSeqInput, debondingDateInput].forEach(input => {
  input.addEventListener('input', () => {
    if (extractedData) {
      updatePreviewHeader();
    }
  });
});

function updatePreviewHeader() {
  prevHeaderDate.innerText = debondingDateInput.value || '2026. 3. 12';
  prevHeaderSeq.innerText = patientSeqInput.value || '23-89';
  prevHeaderName.innerText = patientNameInput.value || '이수아';
  prevHeaderId.innerText = patientIdInput.value || '1219345';
}

function setProgress(show, text = '', targetBtn = btnGenerate) {
  if (show) {
    targetBtn.disabled = true;
    targetBtn.querySelector('.btn-text').classList.add('hidden');
    targetBtn.querySelector('.loader').classList.remove('hidden');
    progressLog.classList.remove('hidden');
    progressText.innerText = text;
    // 진행 중에는 다른 단계 버튼도 눌리지 않게 막습니다.
    [btnExtractDates, btnGenerate].forEach(b => { if (b !== targetBtn) b.disabled = true; });
  } else {
    targetBtn.disabled = false;
    targetBtn.querySelector('.btn-text').classList.remove('hidden');
    targetBtn.querySelector('.loader').classList.add('hidden');
    progressLog.classList.add('hidden');
    btnExtractDates.disabled = false;
    // btnGenerate는 확정된 날짜가 있을 때만 다시 활성화합니다.
    btnGenerate.disabled = confirmedDates.length === 0;
  }
}

// =====================================================================================
// 날짜 목록 편집 UI (1단계 결과를 사용자가 확인/수정하는 화면)
// =====================================================================================
function renderDateList() {
  dateListContainer.innerHTML = '';

  confirmedDates.forEach((date, idx) => {
    const row = document.createElement('div');
    row.style.cssText = 'display:flex; align-items:center; gap:8px; margin-bottom:6px;';

    const idxLabel = document.createElement('span');
    idxLabel.innerText = `${idx + 1}.`;
    idxLabel.style.cssText = 'width:28px; flex-shrink:0; color:#888; font-size:13px; text-align:right;';

    const input = document.createElement('input');
    input.type = 'text';
    input.value = date;
    input.placeholder = '예: 2024.01.15';
    input.style.cssText = 'flex:1; padding:6px 8px; border:1px solid #ccc; border-radius:4px; font-size:14px;';
    input.addEventListener('input', () => {
      confirmedDates[idx] = input.value;
    });

    const removeBtn = document.createElement('button');
    removeBtn.type = 'button';
    removeBtn.innerText = '✕';
    removeBtn.title = '이 날짜 삭제';
    removeBtn.style.cssText = 'flex-shrink:0; padding:5px 10px; border:none; background:#f5c2c0; color:#c0392b; border-radius:4px; cursor:pointer; font-size:13px;';
    removeBtn.addEventListener('click', () => {
      confirmedDates.splice(idx, 1);
      renderDateList();
    });

    row.appendChild(idxLabel);
    row.appendChild(input);
    row.appendChild(removeBtn);
    dateListContainer.appendChild(row);
  });

  dateReviewDesc.innerText = `총 ${confirmedDates.length}건. 잘못된 날짜는 직접 수정하고, 빠지거나 초과된 항목은 추가/삭제하세요. 확인이 끝나면 아래 "2단계" 버튼을 눌러주세요.`;
  btnGenerate.disabled = confirmedDates.length === 0;
}

btnAddDate.addEventListener('click', () => {
  confirmedDates.push('');
  renderDateList();
});

// =====================================================================================
// 확정된 날짜 목록과 텍스트에서 나눈 진료 블록을 병합.
//
// 차트 텍스트의 진료 기록 개수와 캡처 이미지의 날짜 헤더 개수가 다를 수 있습니다
// (한쪽이 더 최근 기록까지 포함하는 등). 다만 시작점(가장 오래된 기록)은 항상
// 같다는 전제를 이용해, 앞에서부터(index 0) 순서대로 1:1 매칭합니다.
//  - 텍스트 블록이 더 많으면: 남는 뒤쪽 블록은 date를 빈 값("")으로 둡니다.
//  - 날짜가 더 많으면: 남는 뒤쪽 날짜는 사용하지 않고 별도로 반환해 경고에 사용합니다.
// =====================================================================================
function mergeDatesAndBlocks(dates, blocks) {
  const minLen = Math.min(dates.length, blocks.length);
  const timeline = [];

  for (let i = 0; i < minLen; i++) {
    timeline.push({ date: dates[i], content: blocks[i] });
  }

  let blocksWithoutDate = 0;
  for (let i = minLen; i < blocks.length; i++) {
    timeline.push({ date: '', content: blocks[i] });
    blocksWithoutDate++;
  }

  const unusedDates = dates.slice(minLen);

  return { timeline, unusedDates, blocksWithoutDate };
}

// =====================================================================================
// 1단계: 이미지에서 날짜 헤더만 추출
// =====================================================================================
btnExtractDates.addEventListener('click', async () => {
  const apiKey = apiKeyInput.value.trim();

  if (!apiKey) {
    alert('Gemini API Key를 입력해주세요.');
    return;
  }
  if (uploadedImages.length === 0) {
    alert('날짜 분석을 위해 차트 캡처 이미지를 1개 이상 업로드해주세요.');
    return;
  }

  if (saveKeyCheckbox.checked) {
    localStorage.setItem('ortho_gemini_api_key', apiKey);
  } else {
    localStorage.removeItem('ortho_gemini_api_key');
  }

  try {
    setProgress(true, '이미지에서 날짜 헤더를 추출하는 중입니다 (Gemini API 호출)...', btnExtractDates);

    const genAI = new GoogleGenerativeAI(apiKey);
    const model = genAI.getGenerativeModel({ model: "gemini-3.5-flash" });

    const imageParts = uploadedImages.map(img => ({
      inlineData: {
        data: img.base64,
        mimeType: img.type
      }
    }));

    const datePrompt = `
당신은 한국 치과 교정과 EMR 차트 캡처 이미지에서 진료 날짜만 정확히 추출하는 전문가입니다.

■ 목표
제공된 캡처 이미지들에서, 각 진료 기록의 "강조(하이라이트) 헤더 바"에 있는 날짜만 순서대로 추출하세요. 내용(S/O/Tx 등)은 필요 없습니다. 오직 날짜만 뽑으세요.

■ 날짜 판별 규칙 — 반드시 (A)만 사용하고 (B)는 절대 사용하지 마세요.
(A) 사용할 것: 주황색/노란색 등으로 강조된 헤더 바의 날짜.
    형식 예: "2024.01.15  21세  교정과초진2" 또는 "2025.03.07  23세  교정과재진2"
    → 이 헤더 바의 "YYYY.MM.DD" 부분만 추출하세요.
(B) 사용 금지: "이름 YYMMDD HH:MM / 이름 YYMMDD HH:MM" 형식의 서명 타임스탬프(예: "정석기 240115 17:51 / 정석기 240115 17:51").
    이는 차트를 작성/서명한 시각이며 실제 진료일과 다를 수 있습니다. 절대 날짜로 사용하지 마세요.
- 어떤 블록의 서명 타임스탬프(B)는 그 블록의 맨 아래, 즉 "다음 블록의 헤더 바" 바로 위에 붙어서 나타납니다. 시각적으로 가깝다고 해서 아래쪽(다음) 블록의 날짜로 착각하지 마세요. 서명 타임스탬프는 항상 위쪽(이전) 블록에 속합니다.

■ 작업 방법
1. 제공된 모든 이미지를 페이지/컬럼 순서대로 훑으며, 강조된 헤더 바를 하나도 빠짐없이 찾으세요. 한 이미지 안에 여러 개의 헤더 바가 있을 수 있습니다.
2. 각 헤더 바의 날짜를 "YYYY.MM.DD" 형식으로 통일하세요 (예: "2024.1.15" → "2024.01.15").
3. 발견한 순서(=시간 오름차순) 그대로 하나의 배열로 나열하세요. 이미지 여러 장에 걸쳐 있어도 전체를 하나의 순서로 합치세요.
4. 최종 결과를 만들기 전에, 날짜가 위에서 아래로 갈수록 항상 이전 값과 같거나 이후인지 스스로 재확인하세요. 같은 날짜가 연속으로 나오거나 순서가 거꾸로 간다면, 헤더를 다시 자세히 확인해 실제 값을 다시 읽으세요.

■ 출력 형식 (순수 JSON만 반환, 다른 설명/코드블록 금지):
{
  "dates": ["2024.01.15", "2024.01.19", "2024.02.26"]
}
    `;

    const result = await model.generateContent([datePrompt, ...imageParts]);
    let responseText = result.response.text().trim();

    if (responseText.startsWith('```json')) {
      responseText = responseText.substring(7);
    }
    if (responseText.endsWith('```')) {
      responseText = responseText.substring(0, responseText.length - 3);
    }
    responseText = responseText.trim();

    let parsed;
    try {
      parsed = JSON.parse(responseText);
    } catch (parseErr) {
      console.error('JSON 파싱 오류. 응답 원본:', responseText);
      throw new Error('API가 올바른 JSON 형식을 반환하지 않았습니다. 원본 응답을 확인해주세요.');
    }

    if (!Array.isArray(parsed.dates)) {
      throw new Error('날짜 배열(dates)을 찾을 수 없습니다. 원본 응답을 확인해주세요.');
    }

    confirmedDates = parsed.dates;
    renderDateList();
    dateReviewSection.style.display = 'block';

    setProgress(false, '', btnExtractDates);
    btnGenerate.disabled = confirmedDates.length === 0;

  } catch (error) {
    console.error(error);
    setProgress(false, '', btnExtractDates);
    alert('날짜 추출 중 오류가 발생했습니다: ' + error.message);
  }
});

// =====================================================================================
// 2단계: 확정된 날짜 목록 + 차트 텍스트로 DB sheet 생성 (이미지는 더 이상 사용하지 않음 —
// 날짜는 이미 1단계에서 사람이 확인/확정했으므로, 여기서는 그 값을 그대로 사용합니다.)
// =====================================================================================
btnGenerate.addEventListener('click', async () => {
  const apiKey = apiKeyInput.value.trim();
  const chartText = chartTextArea.value.trim();

  if (!apiKey) {
    alert('Gemini API Key를 입력해주세요.');
    return;
  }
  if (!chartText) {
    alert('차트 복사 텍스트를 입력해주세요.');
    return;
  }
  if (confirmedDates.length === 0) {
    alert('먼저 1단계에서 날짜 헤더를 추출하고 확인해주세요.');
    return;
  }
  // 편집 중 비어있는 날짜 칸이 남아있는지 확인
  if (confirmedDates.some(d => !d || !d.trim())) {
    alert('비어있는 날짜 항목이 있습니다. 값을 채우거나 해당 항목을 삭제해주세요.');
    return;
  }

  if (saveKeyCheckbox.checked) {
    localStorage.setItem('ortho_gemini_api_key', apiKey);
  } else {
    localStorage.removeItem('ortho_gemini_api_key');
  }

  try {
    setProgress(true, '차트 텍스트를 방문 단위로 정리하는 중입니다 (Gemini API 호출)...', btnGenerate);

    const genAI = new GoogleGenerativeAI(apiKey);
    const model = genAI.getGenerativeModel({ model: "gemini-3.5-flash" });

    // AI가 치식을 변환하지 않고 그대로 반환하도록 프롬프트 공백 및 지시문 최적화
    // ※ 여기서는 날짜를 전혀 다루지 않습니다. 차트 텍스트와 캡처 이미지의 진료 기록
    //   개수가 서로 다를 수 있어서(한쪽이 더 최근 기록까지 포함하는 등), 텍스트 블록
    //   개수를 날짜 개수에 강제로 맞추면 오히려 왜곡이 생길 수 있기 때문입니다.
    //   날짜는 1단계에서 확정한 confirmedDates와 여기서 나온 블록을 앞에서부터
    //   순서대로 매칭하는 방식(mergeDatesAndBlocks)으로 별도 처리합니다.
    const prompt = `
당신은 한국 치과 교정과의 숙련된 임상 기록 정리 전문가입니다.
아래 EMR 차트 텍스트(날짜 제외, 오름차순 정렬됨)를 진료 방문 단위로 나누고, 손글씨로 작성하던 'Debonding Sheet(DB Sheet)' 스타일로 간결하게 정리해주세요. 날짜는 신경 쓰지 마세요 — 텍스트에 날짜가 없으므로 내용 정리에만 집중하세요.

■ 복사한 차트 텍스트:
"""
${chartText}
"""

■ 작업 지침
[1] 블록 분리
- 진단명 반복, "진료 및 경과" 구분자 등을 참고해 텍스트를 진료 방문 단위로 정확히 나누세요.
- 블록 개수를 특정 숫자에 맞출 필요는 없습니다. 텍스트 자체의 방문 구분을 있는 그대로 반영하세요.
- 나눈 순서(=시간 오름차순)를 그대로 유지하세요.

[2] 기록 형식 — 반드시 아래 형식을 따르세요:

● 일반 진료 내원:
- S) 있을 때만 작성 (주관적 소견 = 환자가 말한 증상)
- O) 있을 때만 작성 (객관적 소견 = 임상가가 관찰한 소견)
- Tx) 반드시 작성 (치료 처치 내용)
  → Tx) 내용이 없으면 해당 섹션 생략
- N) 다음 진료 계획(원본의 'n)' 등)
- NN) 다다음 진료 계획(원본의 'nn)' 등)

● 첫 내원 / 진단 차트 (Record taking, A-record, 진단명 결정 등이 포함된 경우):
- C.C) 주 호소
- PI) 현병력 요약
- [Tx Plan] 치료 계획
- Tx) 있을 때만 작성

[3] 작성 스타일 — 손글씨 DB sheet처럼 극도로 간결하게:
- 줄글 금지. 항목 당 짧은 키워드/구절 위주로 작성.
- 진단명 반복 제거.
- 예진자/주치의 등 행정 정보 제거.
- ★핵심 지시사항: 텍스트에 포함된 치식 기호(예: {치식:7|||})는 절대 풀어서 쓰거나 변환하지 말고, 기호와 중괄호를 포함해 원본 형태 그대로 출력하세요.
- 교정 약어 유지(NT, SS, TMA, PC, fig, el, bkt, del 등)
- Elastic: "El. 3/16 med. ({치식:7|7|4|4})" 형식.

■ 출력 JSON 형식(순수 JSON만 반환, 다른 설명/코드블록 금지) — 날짜 필드 없이, 방문 순서대로 나열한 배열:
{
  "blocks": [
    "S) 불편함 없음\\nTx) {치식:7|||} bkt re-bonding, Upper NT 교체\\nN) {치식:|4||} ext 고려",
    "Tx) ..."
  ]
}
    `;

    const result = await model.generateContent([prompt]);
    let responseText = result.response.text().trim();

    if (responseText.startsWith('```json')) {
      responseText = responseText.substring(7);
    }
    if (responseText.endsWith('```')) {
      responseText = responseText.substring(0, responseText.length - 3);
    }
    responseText = responseText.trim();

    let parsed;
    try {
      parsed = JSON.parse(responseText);
    } catch (parseErr) {
      console.error('JSON 파싱 오류. 응답 원본:', responseText);
      throw new Error('API가 올바른 JSON 형식을 반환하지 않았습니다. 원본 응답을 확인해주세요.');
    }

    if (!Array.isArray(parsed.blocks)) {
      throw new Error('blocks 배열을 찾을 수 없습니다. 원본 응답을 확인해주세요.');
    }

    // 확정된 날짜 목록과 텍스트 블록을 앞에서부터(=가장 오래된 기록부터) 순서대로 매칭.
    // 둘 중 하나가 더 최근 기록까지 포함해 개수가 달라도, 시작점은 같다는 전제를 이용합니다.
    const { timeline, unusedDates, blocksWithoutDate } = mergeDatesAndBlocks(confirmedDates, parsed.blocks);

    extractedData = { timeline };

    jsonOutput.value = JSON.stringify(extractedData, null, 2);
    renderPreview();
    btnDownload.disabled = false;

    setProgress(false, '', btnGenerate);

    let mismatchMsg = '';
    if (blocksWithoutDate > 0) {
      mismatchMsg += `\n\n⚠️ 차트 텍스트 쪽에 날짜가 없는 진료 기록이 ${blocksWithoutDate}건 더 있습니다(캡처본에 없는 더 최근 기록으로 보임). 표에서 "날짜 없음"으로 표시된 행에 직접 날짜를 입력해주세요.`;
    }
    if (unusedDates.length > 0) {
      mismatchMsg += `\n\n⚠️ 확정한 날짜 중 ${unusedDates.length}건(${unusedDates.join(', ')})은 대응하는 텍스트 기록을 찾지 못해 사용되지 않았습니다(캡처본에만 있는 더 최근 날짜로 보임). 차트 텍스트를 더 최근 기록까지 포함해서 다시 붙여넣어야 할 수 있습니다.`;
    }

    alert('DB sheet 데이터 추출이 완료되었습니다! 프리뷰를 검토하고 워드 파일을 다운로드하세요.' + mismatchMsg);

  } catch (error) {
    console.error(error);
    setProgress(false, '', btnGenerate);
    alert('오류가 발생했습니다: ' + error.message);
  }
});

// =====================================================================================
// 날짜 이상 감지 — AI의 이미지 판독(OCR) 특성상 가끔 헤더 날짜를 잘못 읽어
// 같은 날짜가 연속 중복되거나, 이후 모든 행이 한 칸씩 밀리는 경우가 있습니다.
// 완벽한 자동 교정은 불가능하지만, "여기 확인해보세요"라고 짚어주는 용도로
// 텍스트 블록이 날짜 오름차순이라는 전제를 활용해 의심 지점을 표시합니다.
// =====================================================================================
function parseKoreanDate(str) {
  if (!str) return null;
  const m = String(str).match(/(\d{4})\s*[.\-/]\s*(\d{1,2})\s*[.\-/]\s*(\d{1,2})/);
  if (!m) return null;
  const year = Number(m[1]), month = Number(m[2]), day = Number(m[3]);
  if (month < 1 || month > 12 || day < 1 || day > 31) return null;
  return new Date(year, month - 1, day);
}

function findDateAnomalies(timeline) {
  const anomalies = [];
  let prevDate = null;
  let prevIndex = -1;

  timeline.forEach((row, idx) => {
    const d = parseKoreanDate(row.date);
    if (!d) return; // 날짜 형식을 못 읽으면 검사 대상에서 제외

    if (prevDate) {
      if (d.getTime() < prevDate.getTime()) {
        anomalies.push({ index: idx, date: row.date, type: 'decrease' });
      } else if (d.getTime() === prevDate.getTime()) {
        anomalies.push({ index: idx, date: row.date, type: 'duplicate' });
      }
    }
    prevDate = d;
    prevIndex = idx;
  });

  return anomalies;
}

function renderDateAnomalyBanner(anomalies) {
  let banner = document.getElementById('date-anomaly-banner');
  if (!banner) {
    banner = document.createElement('div');
    banner.id = 'date-anomaly-banner';
    banner.style.cssText = 'margin-bottom:12px; padding:10px 14px; border-radius:6px; font-size:13px; line-height:1.6; border:1px solid transparent;';
    docPreviewContainer.insertBefore(banner, docPreviewContainer.firstChild);
  }

  if (anomalies.length === 0) {
    banner.style.display = 'none';
    return;
  }

  banner.style.display = 'block';
  banner.style.background = '#fdecea';
  banner.style.color = '#c0392b';
  banner.style.borderColor = '#f5c2c0';

  const list = anomalies
    .map(a => `${a.index + 1}번째 행 (${a.date}${a.type === 'decrease' ? ' — 이전 행보다 날짜가 앞섬' : ' — 이전 행과 날짜 중복'})`)
    .join(', ');

  banner.innerText = `⚠️ 날짜 순서가 의심되는 행이 ${anomalies.length}곳 있습니다: ${list}. 원본 캡처본과 대조해서 확인해주세요.`;
}

// Render the preview table and header
function renderPreview() {
  if (!extractedData || !Array.isArray(extractedData.timeline)) {
    console.error('extractedData.timeline이 올바르지 않습니다:', extractedData);
    alert('데이터 형식이 예상과 달라 미리보기를 표시할 수 없습니다.');
    return;
  }

  updatePreviewHeader();
  dbSheetTableBody.innerHTML = '';

  const anomalies = findDateAnomalies(extractedData.timeline);
  const anomalyIndexSet = new Set(anomalies.map(a => a.index));

  extractedData.timeline.forEach((row, idx) => {
    const tr = document.createElement('tr');

    const tdDate = document.createElement('td');
    tdDate.style.fontWeight = '600';
    tdDate.contentEditable = 'true';
    tdDate.spellcheck = false;

    if (!row.date || !row.date.trim()) {
      // 캡처본에 대응하는 날짜가 없어 비워진 행 — 직접 입력하도록 안내
      tdDate.innerText = '';
      tdDate.dataset.placeholder = '날짜 입력';
      tdDate.style.color = '#e67e22';
      tdDate.style.background = '#fff8ec';
      tdDate.title = '⚠️ 대응하는 캡처본 날짜를 찾지 못했습니다. 직접 입력해주세요.';
    } else {
      tdDate.innerText = row.date;
    }

    if (anomalyIndexSet.has(idx)) {
      tdDate.style.color = '#c0392b';
      tdDate.style.textDecoration = 'underline wavy #c0392b';
      tdDate.title = '⚠️ 이전 행과 날짜 순서가 이상합니다. 원본 이미지를 다시 확인해주세요.';
    }

    // 셀을 직접 수정하면 데이터와 JSON 출력에도 반영, 이상 감지도 다시 실행
    tdDate.addEventListener('blur', () => {
      extractedData.timeline[idx].date = tdDate.innerText.trim();
      jsonOutput.value = JSON.stringify(extractedData, null, 2);
      renderPreview();
    });

    const tdContent = document.createElement('td');

    // AI 데이터를 화면에 그리기 전, 정규식 변환 함수를 거칩니다.
    const parsedContent = convertToPalmer(row.content);
    tdContent.innerHTML = formatToothNotationForHTML(parsedContent);

    tr.appendChild(tdDate);
    tr.appendChild(tdContent);
    dbSheetTableBody.appendChild(tr);
  });

  renderDateAnomalyBanner(anomalies);

  previewEmpty.classList.add('hidden');
  docPreviewContainer.classList.remove('hidden');
}

// Convert normal newlines to HTML breaks, \underline{}/\overline{} 토큰을 실제 밑줄/윗줄로 렌더링
function formatToothNotationForHTML(text) {
  let lines = text.split('\n');
  let result = [];

  for (let i = 0; i < lines.length; i++) {
    const escaped = lines[i]
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      // 상악만 있을 때: 밑줄(underline)
      .replace(/\\underline\{([^}]*)\}/g, '<span style="border-bottom:1px solid currentColor; padding-bottom:1px;">$1</span>')
      // 하악만 있을 때: 윗줄(overline)
      .replace(/\\overline\{([^}]*)\}/g, '<span style="border-top:1px solid currentColor; padding-top:1px;">$1</span>');
    result.push(escaped);
  }

  return result.join('<br>');
}

// =====================================================================================
// 워드(.docx) 생성 — 치식 표기를 docx.Math(MathFraction)로 인라인 삽입
//
// 이전 방식(1x1 표를 텍스트 사이에 끼워 넣는 방식)은 Word에서 표가 항상 블록 레벨
// 요소로 취급되어 문장 중간에 줄바꿈이 생기는 문제가 있었습니다.
// docx.MathUnder / docx.MathOver 는 라이브러리에 존재하지 않는 클래스라
// "is not a constructor" 오류가 발생했었고, 실제로 필요한 건 분자/분모 형태의
// docx.MathFraction 하나면 충분합니다 (UR|UL 위, LR|LL 아래, 가운데 가로선).
// docx.Math는 Run과 마찬가지로 Paragraph 안에 인라인으로 들어가므로
// 문장 흐름이 끊기지 않습니다.
// =====================================================================================

// 원본 {치식:...} 패턴과 태그(S) O) Tx) 등)를 직접 파싱해 run 배열을 만듭니다.
function buildRunsFromLine(line) {
  const runs = [];
  const regex = /(\{치식:[^}]*\}|S\)|O\)|Tx\)|N\)|NN\)|C\.C\)|PI\)|\[Tx Plan\])/g;
  const parts = line.split(regex).filter(Boolean);

  parts.forEach(part => {
    const toothMatch = part.match(/^\{치식:([^}]*)\}$/);

    if (toothMatch) {
      const parsed = parseToothTag(toothMatch[1]);

      if (parsed.type === 'arrow2') {
        const a = parsed.parts[0], b = parsed.parts[1];
        const BLANK = ' ';

        // 화살표 케이스(~)는 분수(MathFraction)로 표현 — 4분악/6분악과 동일한 방식
        if (a === '~' && b === '~') {
          // ⇧⇩ : 분자·분모 모두 ~ | ~
          runs.push(new docx.Math({
            children: [
              new docx.MathFraction({
                numerator: [new docx.MathRun('~ | ~')],
                denominator: [new docx.MathRun('~ | ~')],
              }),
            ],
          }));
          return;
        }
        if (a === '~') {
          // ⇧ : 분자만 ~ | ~ (분모는 공백 → 밑줄처럼 보임)
          runs.push(new docx.Math({
            children: [
              new docx.MathFraction({
                numerator: [new docx.MathRun('~ | ~')],
                denominator: [new docx.MathRun(BLANK)],
              }),
            ],
          }));
          return;
        }
        if (b === '~') {
          // ⇩ : 분모만 ~ | ~ (분자는 공백 → 윗줄처럼 보임)
          runs.push(new docx.Math({
            children: [
              new docx.MathFraction({
                numerator: [new docx.MathRun(BLANK)],
                denominator: [new docx.MathRun('~ | ~')],
              }),
            ],
          }));
          return;
        }

        // 화살표가 아닌 일반 텍스트 표기는 그대로 유지
        let text;
        if (a && !b) text = '상악 ' + a;
        else if (!a && b) text = '하악 ' + b;
        else text = '상악 ' + a + ' / 하악 ' + b;
        runs.push(new docx.TextRun({ text, size: 20, font: 'NanumGothic' }));
        return;
      }

      // 4분악(2+2) / 6분악(3+3) 공통 처리
      const { upperParts, lowerParts, hasUpper, hasLower } = parsed;

      // 위/아래 모두 비어있으면 아무것도 출력하지 않음
      if (!hasUpper && !hasLower) return;

      const upperText = upperParts.map(v => v || ' ').join(' | ');
      const lowerText = lowerParts.map(v => v || ' ').join(' | ');
      const BLANK = ' '; // 빈 쪽에 넣을 최소 공백(빈 MathRun 방지용)

      if (hasUpper && hasLower) {
        // 상악·하악 둘 다 있으면 일반 분수(가로선 하나, 위/아래 모두 내용 있음)
        runs.push(new docx.Math({
          children: [
            new docx.MathFraction({
              numerator: [new docx.MathRun(upperText)],
              denominator: [new docx.MathRun(lowerText)],
            }),
          ],
        }));
      } else if (hasUpper) {
        // 상악만 있으면 "밑줄(underline)" — 분모를 빈 공백으로 둔 분수로 표현하면
        // 상악 텍스트 바로 아래에 선 하나만 그려져 언더라인처럼 보입니다.
        runs.push(new docx.Math({
          children: [
            new docx.MathFraction({
              numerator: [new docx.MathRun(upperText)],
              denominator: [new docx.MathRun(BLANK)],
            }),
          ],
        }));
      } else {
        // 하악만 있으면 "윗줄(overline)" — 분자를 빈 공백으로 둔 분수로 표현하면
        // 하악 텍스트 바로 위에 선 하나만 그려져 오버라인처럼 보입니다.
        runs.push(new docx.Math({
          children: [
            new docx.MathFraction({
              numerator: [new docx.MathRun(BLANK)],
              denominator: [new docx.MathRun(lowerText)],
            }),
          ],
        }));
      }
      return;
    }

    if (/^(S\)|O\)|Tx\)|N\)|NN\)|C\.C\)|PI\)|\[Tx Plan\])$/.test(part)) {
      runs.push(new docx.TextRun({ text: part, bold: true, size: 20, font: 'NanumGothic' }));
    } else if (part !== '') {
      runs.push(new docx.TextRun({ text: part, size: 20, font: 'NanumGothic' }));
    }
  });

  return runs;
}

// 여러 줄(\n)로 구성된 content를 줄 단위 Paragraph 배열로 변환
function buildParagraphsFromContent(content) {
  const lines = content.split('\n');
  return lines.map(line => new docx.Paragraph({
    children: buildRunsFromLine(line),
    spacing: { after: 40 },
  }));
}

// Generate and Download .docx File
btnDownload.addEventListener('click', async () => {
  if (!extractedData) return;

  const name = patientNameInput.value.trim();
  const id = patientIdInput.value.trim();
  const seq = patientSeqInput.value.trim();
  const debondDate = debondingDateInput.value.trim();

  try {
    const doc = new docx.Document({
      sections: [
        {
          properties: {
            page: {
              margin: { top: 720, right: 720, bottom: 720, left: 720 }, // 마진 설정
            },
            column: {
              space: 240, // 단 사이 간격
              count: 2    // 2단 레이아웃 분할
            }
          },
          children: [
            new docx.Paragraph({
              alignment: docx.AlignmentType.RIGHT,
              children: [new docx.TextRun({ text: debondDate, size: 22, font: 'NanumGothic' })],
              spacing: { after: 120 }
            }),
            new docx.Paragraph({
              alignment: docx.AlignmentType.CENTER,
              children: [
                new docx.TextRun({ text: `${seq}  ${name}  Debonding sheet  `, bold: true, size: 28, font: 'NanumGothic' }),
                new docx.TextRun({ text: '∅', bold: false, size: 28, font: 'Cambria' }),
                new docx.TextRun({ text: `  ${id}`, bold: true, size: 28, font: 'NanumGothic' })
              ],
              spacing: { after: 400 }
            }),
            createDocxTable(extractedData.timeline, docx)
          ]
        }
      ]
    });

    const blob = await docx.Packer.toBlob(doc);
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${seq || 'DB'}_${name || '환자'}_Debonding_Sheet.docx`;
    a.click();
    URL.revokeObjectURL(url);

  } catch (err) {
    console.error(err);
    alert('워드 파일 생성 도중 오류가 발생했습니다: ' + err.message);
  }
});

function createDocxTable(timeline, docx) {
  const tableRows = [
    // 헤더 행은 그대로 유지
    new docx.TableRow({
      children: [
        new docx.TableCell({ children: [new docx.Paragraph({ children: [new docx.TextRun({ text: "날짜", bold: true })] })] }),
        new docx.TableCell({ children: [new docx.Paragraph({ children: [new docx.TextRun({ text: "진료 내용", bold: true })] })] })
      ]
    })
  ];

  timeline.forEach(row => {
    // 원본 content(치식 원본 표기 포함)를 줄 단위 Paragraph 배열로 변환
    // — 표(mini table) 대신 docx.Math(MathFraction)를 인라인으로 사용하므로
    //   더 이상 convertToPalmer를 거칠 필요가 없습니다.
    const cellChildren = buildParagraphsFromContent(row.content);

    tableRows.push(new docx.TableRow({
      children: [
        new docx.TableCell({ children: [new docx.Paragraph({ children: [new docx.TextRun(row.date)] })] }),
        new docx.TableCell({ children: cellChildren })
      ]
    }));
  });

  return new docx.Table({ width: { size: 100, type: docx.WidthType.PERCENTAGE }, rows: tableRows });
}
