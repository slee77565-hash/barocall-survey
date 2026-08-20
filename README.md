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
const NOTIFY_EMAIL = 'b831028@naver.com';  // 응답 알림을 받을 주소
const NOTIFY_EVERY = 1;                    // 1이면 매 응답마다 메일, 10이면 10건마다

function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const d = JSON.parse(e.postData.contents);
  const a = d.answers, m = d.meta;

  const keys = Object.keys(a);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(
      ['제출시각', '주체', '근거', '매개제시순서', '다시보기', '소요초', '시나리오체류ms']
        .concat(keys)
    );
  }
  sheet.appendRow(
    [m.finishedAt, m.agent, m.basis, (m.mediatorOrder || []).join('-'),
     m.replays || 0, m.durationSec, (d.timing || {}).B3 || '']
      .concat(keys.map(k => a[k]))
  );

  const n = sheet.getLastRow() - 1;
  if (NOTIFY_EMAIL && n % NOTIFY_EVERY === 0) {
    MailApp.sendEmail(
      NOTIFY_EMAIL,
      '[설문] 응답 ' + n + '건 도착 (' + m.agent + ' / ' + m.basis + ')',
      '새 응답이 도착했습니다.\n\n' +
      '조건: ' + m.agent + ' / ' + m.basis + '\n' +
      '소요 시간: ' + m.durationSec + '초\n' +
      '누적 응답: ' + n + '건\n\n' +
      '전체 자료 보기: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl()
    );
  }
  return ContentService.createTextOutput('ok');
}
```
`NOTIFY_EMAIL`이 본인 주소인지 확인합니다. 응답이 많아져 메일이 번거로우면 `NOTIFY_EVERY`를 10이나 20으로 바꾸세요.
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
특정 조건만 보려면 `?agent=ai&basis=data` 처럼 붙이면 됩니다. `agent`는 `ai` 또는 `human`, `basis`는 `market` 또는 `data`입니다.
---
4. 설계에 반영된 원칙
표지: 연구자와 지도교수, 소요 시간, 익명 처리, 문의처를 밝히고 동의에 체크해야 시작 버튼이 열립니다.
자극 순차 노출과 진행 잠금: 상황 안내가 한 문단씩 나타나고 마지막에 앱 화면과 안내 문구가 드러납니다. 이 순서가 끝나기 전에는 다음 버튼이 잠깁니다. 단순 대기 타이머와 달리 참가자가 조작에 해당하는 안내 문구를 반드시 보게 됩니다. 약 17초가 걸립니다.
다시 보기: 노출이 끝나면 다시 볼 수 있고, 횟수가 자료에 기록됩니다.
무작위 배정: 접속할 때마다 네 조건 중 하나에 배정됩니다.
문항 순서 무작위화: 확증적 매개변수 세 묶음의 순서와 묶음 내 문항 순서를 섞습니다. 배정된 순서는 `mediatorOrder`에 남습니다.
이전 버튼: 측정 문항 사이에서는 돌아갈 수 있지만, 상황 안내 화면으로는 돌아갈 수 없습니다. 문항을 본 뒤 시나리오를 다시 읽으면 기억이 아니라 재독에 근거한 응답이 되기 때문입니다.
응답 누락 방지: 모든 문항에 응답해야 넘어갑니다.
주의점검: `AT1` 문항이 포함되어 있습니다.
조건별 문구 치환: 주체 표현이 자동으로 바뀌며 한국어 조사도 받침에 맞춰 처리됩니다. 네 조건의 문장 구조와 서술어는 동일합니다.
디브리핑: 마지막 화면에서 가상의 앱과 요금제였음을 안내합니다.
수집 자료에는 응답값과 함께 조건, 매개변수 제시 순서, 화면별 체류 시간, 다시 보기 횟수, 총 소요 시간이 포함됩니다.
---
5. 파일럿이 끝나면
마지막의 설문 참여 소감(F1~F5)은 문항을 다듬기 위한 것입니다. 본조사에서는 `CONFIG.FEEDBACK`을 `false`로 바꾸면 사라집니다.
`F5`("이 설문이 무엇을 알아보려 한다고 생각하셨습니까?")의 답이 실제 연구 목적에 지나치게 가깝다면
요구특성 위험이 있다는 뜻이므로 시나리오 문구를 손볼 필요가 있습니다.
---
6. 아직 확정되지 않은 문항
`JD1`~`JD3`(판단받는 느낌)과 `EX1`~`EX3`(착취성 지각)은 Duani, Barasch, & Morwitz(2024)의 원문 측정문항을
확보하지 못해 구성개념 정의에 근거해 자체 개발한 척도입니다. 특히 `JD`는 확증적 매개변수이므로
원문을 확보하게 되면 본조사 이전에 교체합니다.
---
7. 목표 표본
셀당 30~40명, 총 120~160명입니다. 조건이 넷이므로 셀당 인원에 4를 곱한 값이 전체 표본이 됩니다.
