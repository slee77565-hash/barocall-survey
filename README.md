택시 호출 앱 이용 경험 설문조사
국민대학교 비즈니스IT전문대학원 박사과정 이승현 · 지도교수 안현철
불리한 요금 인상 상황에서 가격결정 주체(AI vs. 요금정책 담당자)와 요금 산정 근거(실시간 시장 수요 vs. 개인 이용 데이터)를
교차 조작한 2×2 피험자 간 실험의 설문 도구입니다. 파일은 `index.html` 하나이며 외부 라이브러리를 쓰지 않습니다.
---
1. 지금 바로 해야 할 일
응답이 이메일로 오게 하려면 2번 항목의 설정을 먼저 마쳐야 합니다.
설정 전에는 참가자가 설문을 끝내도 응답이 어디에도 저장되지 않습니다.
---
2. 응답을 이메일과 스프레드시트로 받기
Google 계정만 있으면 됩니다. 10분이면 끝납니다.
2-1. 스프레드시트 만들기
Google 드라이브에서 새 스프레드시트를 만듭니다. 이름은 자유롭게 정하세요.
상단 메뉴에서 확장 프로그램 → Apps Script를 엽니다.
편집기에 있던 내용을 모두 지우고 아래 코드를 붙여넣습니다.
```javascript
/**
 * 바로콜 설문 응답 수집기
 * 열 순서를 고정하므로 문항 제시 순서가 응답자마다 섞여도 값이 어긋나지 않습니다.
 */

const NOTIFY_EMAIL = 'b831028@naver.com';  // 비우면 메일을 보내지 않음
const NOTIFY_EVERY = 1;                    // 1이면 매 응답, 10이면 10건마다

const META_COLS = ['응답ID', '패널ID', '제출시각', '주체', '근거', '매개제시순서',
                   '다시보기', '소요초', '시나리오체류ms', '주의점검', '주체인식', '근거인식'];

// 모든 문항을 사전에 고정합니다. 새 문항을 추가하면 반드시 여기에도 넣으세요.
const ITEM_COLS = [
  'consent', 'S1', 'S2', 'U1', 'U2',
  'DF1', 'DF2', 'DF3', 'PF1', 'PF2', 'PF3',
  'AC1', 'AC2', 'AC3', 'CH1',
  'SI1', 'SI2', 'SI3',
  'JD1', 'JD2', 'JD3',
  'UN1', 'UN2', 'UN3',
  'EX1', 'EX2', 'EX3',
  'MC1', 'MC2', 'MC3',
  'SR1', 'SR2', 'SR3',
  'RL1', 'RL2',
  'RE1', 'RE2', 'IM1',
  'AN1', 'AN2', 'AT1',
  'PR1', 'PR2', 'PR3',
  'GP1', 'GP2', 'GP3',
  'D1', 'D2', 'D3', 'D4',
  'F1', 'F2', 'F3', 'F4', 'F5', 'F6'
];

const TAIL_COLS = ['원본JSON'];  // 목록에 없는 문항이 생겨도 자료가 남도록 보관

function doPost(e) {
  var lock = LockService.getScriptLock();
  var locked = false;
  try {
    locked = lock.tryLock(25000);

    var raw = (e && e.postData && e.postData.contents) ? e.postData.contents : '{}';
    var d = JSON.parse(raw);
    var a = d.answers || {}, m = d.meta || {}, t = d.timing || {};

    // 검토 모드 응답은 저장하지 않습니다.
    if (m.review === true) return ContentService.createTextOutput('skipped');

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    var width = META_COLS.length + ITEM_COLS.length + TAIL_COLS.length;

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(META_COLS.concat(ITEM_COLS).concat(TAIL_COLS));
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, width).setFontWeight('bold').setBackground('#eef0fe');
    }

    var row = [
      m.responseId || '', m.pid || '',
      m.finishedAt || new Date().toISOString(),
      m.agent || '', m.basis || '',
      (m.mediatorOrder || []).join('-'),
      m.replays || 0, m.durationSec || '', t.B3 || '',
      (a.AT1 === 4 ? '통과' : '실패'),
      (a.MC1 === m.agent ? '정답' : '오답'),
      (a.MC2 === m.basis ? '정답' : '오답')
    ].concat(ITEM_COLS.map(function (k) {
      return (a[k] === undefined || a[k] === null) ? '' : a[k];
    })).concat([raw]);

    sheet.appendRow(row);

    var n = sheet.getLastRow() - 1;
    if (NOTIFY_EMAIL && n % NOTIFY_EVERY === 0) {
      safeMail('[바로콜 설문] 응답 ' + n + '건 (' + agentName(m.agent) + ' / ' + basisName(m.basis) + ')',
        '조건       : ' + agentName(m.agent) + ' × ' + basisName(m.basis) + '\n' +
        '주의점검   : ' + (a.AT1 === 4 ? '통과' : '실패') + '\n' +
        '조작점검   : 주체 ' + (a.MC1 === m.agent ? '정답' : '오답') +
                     ' / 근거 ' + (a.MC2 === m.basis ? '정답' : '오답') + '\n' +
        '소요 시간  : ' + (m.durationSec || '-') + '초\n' +
        '누적 응답  : ' + n + '건\n\n' + SpreadsheetApp.getActiveSpreadsheet().getUrl());
    }
    return ContentService.createTextOutput('ok');

  } catch (err) {
    console.error(err);
    safeMail('[바로콜 설문] 응답 저장 오류',
      '오류: ' + err + '\n\n받은 자료:\n' +
      ((e && e.postData && e.postData.contents) ? e.postData.contents : '(없음)'));
    return ContentService.createTextOutput('error');
  } finally {
    if (locked) { try { lock.releaseLock(); } catch (ignore) {} }
  }
}

function agentName(v) { return v === 'ai' ? 'AI 알고리즘' : (v === 'human' ? '요금정책 담당자' : '알 수 없음'); }
function basisName(v) { return v === 'data' ? '개인 이용 데이터' : (v === 'market' ? '실시간 시장 수요' : '알 수 없음'); }

function safeMail(subject, body) {
  if (!NOTIFY_EMAIL) return;
  try { MailApp.sendEmail(NOTIFY_EMAIL, subject, body); }
  catch (err) { console.error('메일 발송 실패: ' + err); }
}

// 편집기에서 실행하면 조건별 응답 수와 제외 대상 건수를 메일로 받습니다.
function 진행현황() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var last = sheet.getLastRow();
  if (last < 2) { safeMail('[바로콜 설문] 현황', '아직 응답이 없습니다.'); return; }

  var v = sheet.getRange(2, 4, last - 1, 9).getValues();  // 주체 ~ 근거인식
  var box = { 'ai|market': 0, 'ai|data': 0, 'human|market': 0, 'human|data': 0 };
  var fail = 0, wrong = 0;
  v.forEach(function (r) {
    var key = r[0] + '|' + r[1];
    if (box[key] !== undefined) box[key]++;
    if (r[6] === '실패') fail++;
    if (r[7] === '오답' || r[8] === '오답') wrong++;
  });

  safeMail('[바로콜 설문] 조건별 진행 현황',
    '전체 ' + (last - 1) + '건\n\n' +
    'AI × 시장수요        : ' + box['ai|market'] + '건\n' +
    'AI × 개인데이터      : ' + box['ai|data'] + '건\n' +
    '담당자 × 시장수요    : ' + box['human|market'] + '건\n' +
    '담당자 × 개인데이터  : ' + box['human|data'] + '건\n\n' +
    '주의점검 실패        : ' + fail + '건\n' +
    '조작점검 오답 포함   : ' + wrong + '건');
}
```
`NOTIFY_EMAIL`이 본인 주소인지 확인합니다. 응답이 많아져 메일이 번거로우면 `NOTIFY_EVERY`를 10이나 20으로 바꾸세요.
> **열 순서를 고정한 이유.** 설문은 매개변수 묶음과 문항 순서를 응답자마다 무작위로 섞고, 의인화 문항(AN)은 AI 조건에만 제시됩니다.
> 응답에서 나온 순서대로 값을 기록하면 두 번째 응답부터 JD2 값이 SI1 열에 들어가는 식으로 자료가 어긋납니다.
> 그래서 `ITEM_COLS`에 모든 문항을 미리 고정하고, 각 응답의 값을 이름으로 찾아 해당 열에만 넣습니다.
> **문항을 추가하면 반드시 `ITEM_COLS`에도 넣어야 합니다.** 빠뜨려도 `원본JSON` 열에는 남으므로 자료가 사라지지는 않습니다.
2-2. 웹앱으로 배포하기
편집기 오른쪽 위 배포 → 새 배포를 누릅니다.
톱니바퀴에서 웹 앱을 선택합니다.
다음 사용자 인증 정보로 실행: 나
액세스 권한이 있는 사용자: 모든 사용자
배포를 누르고 권한 승인을 진행합니다. 경고 화면이 나오면 고급 → 이동을 눌러 계속합니다.
발급된 웹 앱 URL을 복사합니다. `https://script.google.com/macros/s/......../exec` 형태입니다.
2-3. 설문에 주소 넣기
`index.html`을 열어 위쪽 `CONFIG`의 `ENDPOINT`에 방금 복사한 주소를 붙여넣습니다.
```js
const CONFIG = {
  ENDPOINT: "https://script.google.com/macros/s/......../exec",
  RESEARCH_EMAIL: "b831028@naver.com",
  FEEDBACK: true,
  ...
};
```
GitHub에서 `index.html`을 열고 연필 아이콘으로 수정한 뒤 Commit changes를 누르면 1분 안에 반영됩니다.
2-4. 확인
설문을 한 번 끝까지 응답해 보세요. 스프레드시트에 한 줄이 추가되고 메일이 오면 정상입니다.
---
3. 링크
용도	주소
참가자용	`https://slee77565-hash.github.io/barocall-survey/`
동료 검토용	`https://slee77565-hash.github.io/barocall-survey/?review=1`
검토 모드에서는 문항 코드가 표시되고, 우측 하단 조건 전환 버튼으로 네 조건을 바로 비교할 수 있습니다.
검토 모드로 끝까지 응답해도 자료는 서버로 전송되지 않습니다. 브라우저와 서버 양쪽에서 이중으로 막아 두었으므로
동료 검토 응답이 실제 자료에 섞이지 않습니다.
특정 조건만 보려면 `?agent=ai&basis=data` 처럼 붙이면 됩니다. `agent`는 `ai` 또는 `human`, `basis`는 `market` 또는 `data`입니다.
---
4. 설계에 반영된 원칙
표지: 연구자와 지도교수, 소요 시간, 익명 처리, 문의처를 밝히고 동의에 체크해야 시작 버튼이 열립니다.
자극 순차 노출과 진행 잠금: 상황 안내가 한 문단씩 나타나고 마지막에 앱 화면과 안내 문구가 드러납니다. 이 순서가 끝나기 전에는 다음 버튼이 잠깁니다. 단순 대기 타이머와 달리 참가자가 조작에 해당하는 안내 문구를 반드시 보게 됩니다. 약 17초가 걸립니다.
다시 보기: 노출이 끝나면 다시 볼 수 있고, 횟수가 자료에 기록됩니다.
무작위 배정과 조건 고정: 처음 접속할 때 네 조건 중 하나에 배정되며, 배정된 조건은 브라우저 세션에 저장됩니다. 중간에 새로고침해도 다른 조건으로 바뀌지 않습니다.
응답 식별: 응답마다 고유 번호(`responseId`)가 발급되어 중복이나 오류를 가려낼 수 있습니다. 유료 패널을 쓸 경우 주소 뒤에 `?pid=참가자번호`를 붙이면 패널 ID도 함께 저장됩니다.
문항 순서 무작위화: 확증적 매개변수 세 묶음의 순서와 묶음 내 문항 순서를 섞습니다. 배정된 순서는 `mediatorOrder`에 남습니다.
이전 버튼: 측정 문항 사이에서는 돌아갈 수 있지만, 상황 안내 화면으로는 돌아갈 수 없습니다. 문항을 본 뒤 시나리오를 다시 읽으면 기억이 아니라 재독에 근거한 응답이 되기 때문입니다.
응답 누락 방지: 모든 문항에 응답해야 넘어갑니다.
주의점검: `AT1` 문항이 포함되어 있습니다.
자극의 대칭: 네 조건의 안내 문구는 주어와 근거만 다르고 서술어·문장 구조·정보 항목 수가 같습니다. 주체 표기도 "바로콜 AI 요금 알고리즘"과 "바로콜 요금정책 담당자"로 소속을 맞추었습니다. 정보량이나 소속감 차이가 조작 효과에 섞이지 않게 하기 위함입니다. 한국어 조사도 받침에 맞춰 자동 처리됩니다.
매개변수 문항의 조건 중립성: 판단받는 느낌(JD)과 개별 사정 고려 부족(UN)은 측정 대상을 가격결정 주체로 통일했고, 이용 기록을 직접 언급하는 표현을 뺐습니다. 개인 데이터 조건에만 해당하는 단어가 문항에 들어가면 그 조건에서만 점수가 올라가는 문항 효과가 생기기 때문입니다.
측정 순서: 종속변수 → 확증적 매개변수 → 착취성 → 조작점검·현실성 → 프라이버시 우려 → 일반 프라이버시 성향. 프라이버시 문항을 조작점검보다 뒤에 둔 것은, 개인정보라는 틀이 먼저 활성화되면 근거 인식(MC2)과 개인관련성(SR) 응답이 흔들릴 수 있기 때문입니다.
디브리핑: 마지막 화면에서 가상의 앱과 요금제였음을 안내합니다.
수집 자료에는 응답값과 함께 조건, 매개변수 제시 순서, 화면별 체류 시간, 다시 보기 횟수, 총 소요 시간이 포함됩니다.
---
5. 파일럿에서 확정할 것
점검 대상	근거 문항	성격
가격결정 주체 조작	MC1	필수
산정 근거 조작	MC2	필수
요금의 불리성	MC3	필수
셀별 시나리오 현실성	RE1~RE3 (네 조건 평균 비교)	필수
개인관련성	SR1~SR3	부가·탐색
AI 의인화	AN1~AN2	부가·탐색
산정 근거의 적합성	RL1~RL2	부가·탐색
몰입도	IM1	참고
척도 품질	전 척도의 분산·신뢰도·분포	필수
SI와 EX의 변별	SI1~SI3, EX1~EX3	필수
DF와 PF의 변별	DF1~DF3, PF1~PF3	필수
개인관련성(SR)과 AI 의인화(AN)는 자극이 무엇을 보여줬는지를 기억하는지가 아니라, 자극이 유발한 심리적
반응에 가깝습니다. 이를 조작 성공의 기준에 넣으면 원하는 심리 상태가 나올 때까지 자극을 손보게 되므로
부가적 검증으로만 사용합니다. 결과변수의 조건 간 차이나 효과 방향 역시 자극 수정의 기준으로 쓰지 않습니다.
다만 응답 분포, 천장·바닥 효과, 내적 일관성 같은 측정도구의 기초적 특성은 반드시 점검합니다.
**산정 근거의 적합성(RL1~RL2)**은 대안 설명을 확인하기 위한 탐색 변수입니다. 시장 수요는 가격과 직접 관련된 정보로
읽히지만 개인 이용 기록은 요금 산정 근거로 자의적이라고 느껴질 수 있습니다. 그렇다면 조작이 바꾼 것은 개인 데이터 여부만이
아니라 근거의 적합성까지가 됩니다. 이 차이를 완전히 없애기는 어렵고 개인화 가격의 본질이기도 하지만, 결과가 개인 데이터
때문인지 근거가 부적절해 보여서인지는 구분해 두어야 하므로 두 문항으로 측정합니다.
시나리오 현실성과 몰입도는 분리했습니다. RE1·RE2는 상황이 있을 법한가를 묻는 현실성이고, IM1은 얼마나 빠져들었는가를
묻는 몰입도입니다. 셀별 비교의 기준은 현실성 두 문항만 사용합니다.
가격 공정성 지표의 산출 방식을 파일럿에서 확정해야 합니다. 분배적 공정성(DF)과 절차적 공정성(PF)을
하나로 평균낼지, DF를 주 종속변수로 두고 PF를 보조 지표로 분리할지의 문제입니다. 절차적 공정성은
판단받는 느낌·착취성 지각과 개념적으로 가까워, 하나로 합치면 매개변수와 종속변수가 겹쳐 매개효과가
과대 추정될 수 있습니다. 확인적 요인분석으로 두 차원이 실제로 구분되는지 확인하고 산출 방식을
사전등록에 명시하되, 현재로서는 DF를 주 종속변수, PF를 보조 지표로 분리하는 쪽을 권합니다.
---
6. 파일럿이 끝나면
마지막의 설문 참여 소감(F1~F5)은 문항을 다듬기 위한 것입니다. 본조사에서는 `CONFIG.FEEDBACK`을 `false`로 바꾸면 사라집니다.
`F5`("이 설문이 무엇을 알아보려 한다고 생각하셨습니까?")의 답이 실제 연구 목적에 지나치게 가깝다면
요구특성 위험이 있다는 뜻이므로 시나리오 문구를 손볼 필요가 있습니다.
---
7. 아직 확정되지 않은 문항
`JD1`~`JD3`(판단받는 느낌)과 `EX1`~`EX3`(착취성 지각)은 Duani, Barasch, & Morwitz(2024)의 원문 측정문항을
확보하지 못해 구성개념 정의에 근거해 자체 개발한 척도입니다. 특히 `JD`는 확증적 매개변수이므로
원문을 확보하게 되면 본조사 이전에 교체합니다.
---
8. 목표 표본
셀당 30~40명, 총 120~160명입니다. 조건이 넷이므로 셀당 인원에 4를 곱한 값이 전체 표본이 됩니다.
지도교수와 협의해 확정한 규모이며, 자료수집 전에 사전등록하고 이후 중간 결과에 따라 표본을 늘리거나 줄이지 않습니다.
이 규모로 검출할 수 있는 효과의 범위는 미리 밝혀 둡니다. 유의수준 .05, 검정력 .80 기준으로 1자유도 상호작용을
검출하려면 편에타제곱 .06 수준의 중간 효과에서는 약 130명이면 충분하지만, .03에서는 약 250명,
.02에서는 약 390명이 필요합니다. 참고한 선행연구의 상호작용 효과크기가 .02~.03 수준이었으므로,
총 160명은 중간 크기 이상의 상호작용에는 적정하나 작은 크기의 상호작용에는 부족합니다.
이 한계를 알고 표본을 확정한 것이므로, 상호작용이 유의하지 않게 나타나더라도 효과의 부재로 해석하지 않고
효과크기 추정치 및 신뢰구간과 함께 보고합니다. 자료수집 이후 효과크기를 확인해 표본을 추가하지 않으며,
더 큰 표본이 필요하다고 판단되면 별도의 후속 연구로 설계합니다.
