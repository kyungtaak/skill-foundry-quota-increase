---
name: "foundry-quota-increase"
description: "Microsoft Foundry Service: Request for Quota Increase 폼을 Playwright로 처리한다. 요청목록·공통정보 Excel 또는 순차 질문으로 정보를 받고, 실제 모델·리전·TPM을 검증해 입력·캡처한다. 실행=Y 행은 정확한 미리보기의 명시적인 일괄 승인을 받은 뒤에만 한 건씩 제출한다. Foundry/Azure OpenAI 쿼터 증액, 입력 테스트, 캡처, 신청에 사용한다."
metadata:
  version: "1.1.0"
---

# Microsoft Foundry Service: Request for Quota Increase

이 스킬을 실행하기 전에 m_get_skill이 반환한 resourceDir에 있는 SKILL.md를 읽고, 그 폴더의 assets, scripts, references를 사용한다. SKILL.md 또는 필수 리소스가 없으면 파일 위치를 확인하고 복구하기 전에는 제출하지 않는다. 다른 Playwright 지원 에이전트에는 같은 폴더를 담은 portable Agent Skills ZIP을 사용한다. 특정 Scout 도구에 종속된 브라우저 코드를 만들지 않는다.

## 기본 동작

공식 주소: https://aka.ms/aoai/quotaincrease
기본 모드는 제출 없는 입력·캡처다. 공통정보 한 세트와 요청목록 한 행이 신청 한 건이다. 지원 범위는 Azure OpenAI chat 또는 Azure Direct Model (ADM)의 Global Standard TPM 요청이다. Audio/Realtime RPM, PTU/Provisioned, Batch, Fine-tuning, Managed Compute, Anthropic, Fireworks는 자동 변환하거나 제출하지 않고 중단·보고한다.

## 템플릿과 정보 수집

assets/quota_increase_request_template.xlsx는 고객 정보를 비운 원본 구조의 템플릿이다. 요청목록의 헤더·I열 수식·숨김 보조 수식은 유지한다. 공통정보는 항목명과 설명을 남기고 B열 값만 비워 둔다. 입력 포맷과 폼 선택값은 참고 자료이며 실제 폼이 우선이다. 원본 및 배포 템플릿은 덮어쓰지 않는다.

Excel이 없으면 references/interactive-intake.json에 따라 의도(입력·캡처/실제 신청), 담당자, 회사·주소, 구독 GUID, 신청 사유, 모델 타입·배포 유형·정확한 모델명, 총 TPM과 필요한 리전을 차례로 받는다. 누락값이나 신청 사유의 사실을 지어내지 않는다. 답을 로컬 JSON에 저장해 prepare-json을 사용한다.

사용자가 파일을 수정해 다시 주면 최신 파일의 담당자·이메일·구독·모델·실행값을 기준으로 삼는다. 이전 첨부 내용이나 이전 실행의 입력값을 섞지 않는다. 삭제한 모델은 되살리지 않고, 요청 ID의 공백은 유지한다. 건수는 마지막 요청 번호나 파일명이 아니라 실제 요청 행과 선택된 Y 행을 세어 계산한다. 이미 제출 중이거나 접수한 요청이 있으면 새 파일로 새 실행을 만들기 전에 기존 처리 기록과 대조해 중복을 막는다.

## 검증과 입력

scripts/quota_workbook.py prepare 또는 prepare-json으로 새 비공개 실행 폴더를 만든다. 요청목록은 3행 헤더, 4행부터 데이터이며 열은 정확한 헤더 이름으로 찾는다. 행별 구독 ID가 공통 Subscription Id보다 우선이다. Y는 제출 후보이며 N/빈칸은 제출 제외다. 사용자가 명시한 입력 테스트에서는 N도 캡처할 수 있지만 제출 모드로 전환하지 않는다.

이메일·GUID 형식, 필수값, 고유 요청 ID, 중복 요청, 입력 수식, 양의 정수 TPM과 1,000의 배수를 확인한다. I열 캐시 대신 H열 총 TPM을 1,000으로 나눈다. 100M TPM은 폼에 100000을 입력한다. 공개 최고 티어나 API availableCapacity를 신청 총량으로 대체하지 않는다.

실제 폼 제목과 Customer Voice 호스트를 확인한다. 질문 번호는 고정하지 않는다. Quota Type → Model Type → 배포 유형 → 정확한 모델 → 실제로 표시되는 리전 → 수량 순서다. 안쪽 listbox의 정확한 option을 선택하고, 없으면 행을 보류한다. 질문이 없는 리전은 '묻지 않음'으로 기록하며 리전 지원 확인으로 해석하지 않는다.

사용자가 지정한 기본 리전은 리전 질문이 실제로 나타날 때만 선택한다. 기본 리전은 실행별 입력값이며 특정 고객의 리전을 스킬의 공통 기본값으로 저장하지 않는다. 모델명·리전·은퇴일은 이전 파일이나 과거 신청 결과만으로 확정하지 않는다. 삭제·제외 요청은 해당 모델에만 적용하고 비슷한 이름의 모델까지 함께 삭제하지 않는다.

render --action fill로 만든 함수를 Playwright 도구에서 실행하고, 반환 JSON을 record-ui로 저장한다. 모든 공통값과 모델별 값을 대조하고 캡처한다. 원본 Excel은 바꾸지 않는다. 생성 JS와 캡처는 비공개 로컬 자료이며 외부 업로드하지 않는다.

공통 항목이 나타난 뒤 폼이 안정될 시간을 두고 입력한다. 각 텍스트 항목은 입력 후 blur한 실제 값을 다시 읽는다. 값이 유지되지 않으면 제출 전 입력만 한 번 더 시도하고, 다시 어긋나면 캡처 후 중단한다. 포커스 이동으로 구독 GUID가 이름에 붙는 등 다른 항목이 바뀐 경우도 전체 대조에서 차단한다. 입력 완료 후 한 번, 제출 동작에서 한 번 전체 값을 비교하고, 제출 전 화면 캡처가 끝난 직후에도 다시 비교한다. 이름·성·이메일·회사·주소·도시·우편번호·국가·구독 GUID·신청 사유 원문 전체와 Quota Type·Model Type·배포 유형·모델·조건부 리전·수량이 모두 일치해야 한다.

## 제출 승인과 일괄 처리

실행=Y는 승인 자체가 아니다. preview의 공통 입력값 전체와 실제 제출할 Y 요청 ID별 모델·리전 또는 묻지 않음·총량·입력값·단위, 보류/제외 행을 정확히 보여준다. 전송 대상이 Microsoft 공식 Foundry 쿼터 신청 폼이라는 점과 실제 URL을 명시한다. 파일의 비공개 정보를 전송하므로 '⚠️ This message contains details from your private data. Confirm before I send?' 경고를 보여주고 실제 사용자 확인을 기다린다.

단건은 한 건, Excel 일괄 요청은 표시한 Y 행 전체를 한 번에 승인받을 수 있다. 방금 받은 승인 원문과 preview fingerprint를 approve 명령으로 기록한다. 이 명령을 실행한 사실을 사용자 승인으로 꾸미지 않는다. 모델·총량·공통값·실제 리전 질문 여부 또는 범위가 달라지면 재승인한다. 승인 뒤 추가된 Y 행은 포함하지 않는다.

정확한 미리보기와 전송 대상·비공개 정보 경고를 보여준 뒤 일괄 승인을 받았다면 각 건마다 다시 승인받지 않는다. 미리보기를 보기 전의 '입력하고 바로 제출해도 된다'는 말은 아직 확인하지 않은 전송 내용의 승인으로 간주하지 않는다.

승인된 행마다 새 폼에 공통정보와 요청값을 채워 확인하고, begin-submit으로 submitting 상태와 단건 ticket을 먼저 저장한다. render --action submit은 유효한 승인·ticket·Y행·현재 폼 값을 확인한 뒤 Submit을 한 번만 누른다. 병렬 제출하거나 직접 HTTP POST로 폼을 우회하지 않는다.

각 요청은 새 탭이나 새 페이지에서 처리한다. 입력된 폼을 이동하다 beforeunload가 뜨면 무시하거나 제출하지 말고, 현재 동작이 입력 점검인지 제출 후 이동인지 확인한다. 접수 완료된 작업용 탭만 닫고 사용자의 다른 탭은 건드리지 않는다. 도구 응답이 중간에 끊겨도 브라우저 작업은 계속 실행될 수 있다. 먼저 실행 중인 작업과 기존 ledger 상태를 확인하고, 진행 중이면 기다린다. 같은 함수를 다시 호출하거나 같은 ticket을 재사용하지 않는다.

명시적인 제출 완료 문구와 완료 캡처가 있어야 finish-submit으로 성공 처리한다. timeout/연결 종료/불확실한 결과는 mark-unknown으로 기록하고 배치를 멈춘다. submitting/unknown/이미 제출한 행을 자동 재시도하거나 새 실행을 만들어 우회하지 않는다. 기존 처리 기록을 재사용하고 사용자 확인으로 실제 접수 여부를 해결한다.

허용 완료 문구는 `Your response was submitted`, `Your response has been submitted`, `응답이 제출되었습니다`, `응답을 제출했습니다`, `Thanks! You've completed the request!`다. 문구가 실제로 보이고 입력 폼은 사라져야 한다. `Thanks!`나 `Thank you`만으로는 접수 완료로 처리하지 않는다. 문구가 새로 바뀌면 실제 완료 화면과 캡처를 확인하고 사용자에게 확인받는다. 네트워크 기록을 확인할 때는 이미 발생한 응답만 읽고, 확인용 POST를 새로 보내지 않는다.

제출 직전 값이 어긋났고 클릭하지 않았음이 확인된 건은 원인을 남긴 뒤 입력만 다시 할 수 있다. 클릭 이후의 오류는 이 경로로 재시도하지 않는다. 실제 접수 완료를 사용자가 확인한 unknown 건은 `reconcile-submitted`로 기존 attempt의 접수 완료만 기록한다. 원래 unknown 증거는 보존되며 새 ticket이나 클릭을 만들지 않는다. 단순히 '계속'이라는 말만으로 unknown을 해제하지 않는다.

## 보고

일괄 제출은 기본적으로 한 건의 접수 기록을 저장할 때마다 `완료 건수/승인 건수 — 요청 ID · 모델명 접수 완료`를 알린다. 사용자가 건별 알림을 요청하면 여러 건을 한 번에 실행한 뒤 묶어서 보고하지 않는다. 진행 알림에 승인 질문을 다시 붙이지 않는다. 입력 점검 성공을 완료 건수에 포함하지 않는다.

입력·캡처·접수 완료·보류·불확정을 구분하고 최종 합계를 보여준다. 요청 ID가 중간에 비면 삭제된 ID와 실제 건수를 설명하되 번호를 새로 매기지 않는다. 요청 ID별 보류·불확정 원인과 캡처를 남긴다. 입력 가능을 접수 성공, 쿼터 승인, 은퇴 모델 지원 또는 실제 배포 가능성으로 설명하지 않는다. 사용자가 요청하면 캡처를 직접 보여준다.

## 로컬 설치와 명령

Python 3.10 이상, openpyxl, Playwright를 실행할 수 있는 에이전트가 필요하다. `python -m pip install -r requirements.txt`로 Python 의존성을 설치하고, 이 폴더를 에이전트의 skill 경로에 등록한다. Scout에서는 resourceDir의 `SKILL.md`를 읽도록 등록하며, 다른 에이전트에서는 같은 폴더를 Agent Skills 디렉터리로 사용한다.

다음 경로는 예시이며 실제 신청 자료는 저장소 밖의 비공개 폴더에 둔다. 아래 명령은 브라우저를 직접 실행하지 않는다.

```text
python scripts/quota_workbook.py prepare --workbook FILE --run-dir PRIVATE_DIR --mode capture --ids REQ-001 --include-disabled
python scripts/quota_workbook.py prepare --workbook FILE --run-dir PRIVATE_DIR --mode submit
python scripts/quota_workbook.py prepare-json --input FILE --run-dir PRIVATE_DIR --mode capture
python scripts/quota_workbook.py render --run-dir PRIVATE_DIR --id ID --action fill --output FILL.js --screenshots-dir PRIVATE_CAPTURES
python scripts/quota_workbook.py record-ui --run-dir PRIVATE_DIR --id ID --evidence UI.json
python scripts/quota_workbook.py preview --run-dir PRIVATE_DIR
python scripts/quota_workbook.py approve --run-dir PRIVATE_DIR --fingerprint HASH --confirmation-text USER_CONFIRMATION
python scripts/quota_workbook.py begin-submit --run-dir PRIVATE_DIR --id ID
python scripts/quota_workbook.py render --run-dir PRIVATE_DIR --id ID --action submit --output SUBMIT.js --screenshots-dir PRIVATE_CAPTURES --ticket TICKET.json
python scripts/quota_workbook.py finish-submit --run-dir PRIVATE_DIR --id ID --evidence RESULT.json
python scripts/quota_workbook.py mark-unknown --run-dir PRIVATE_DIR --id ID --detail REASON
python scripts/quota_workbook.py reconcile-submitted --run-dir PRIVATE_DIR --id ID --evidence CONFIRMED_RESULT.json --confirmation-text USER_RECEIPT_CONFIRMATION
python scripts/quota_workbook.py status --run-dir PRIVATE_DIR
```

`reconcile-submitted`의 evidence에는 기존 `request_id`, `attempt_id`, `payload_hash`와 `submit_clicked=true`, `submission_status="submitted"`, 허용된 `confirmation_text`, 기존 완료 PNG의 절대 경로인 `confirmation_screenshot`이 필요하다. 오류가 없어야 하며, 실제 접수 여부를 확인한 사용자의 원문을 `--confirmation-text`로 별도 기록한다. 증거나 사용자 확인을 만들어내지 않는다.

## 개발과 배포

개발 의존성은 `python -m pip install -r requirements-dev.txt`, 테스트는 `python -m pytest tests -q`다. 브라우저 회귀 점검은 `python tests/make_browser_smoke.py --output PRIVATE_SMOKE.js --captures PRIVATE_CAPTURES`로 생성한 함수를 Playwright 에이전트에서 실행한다. 이 점검은 별도 브라우저 context에서 모든 네트워크를 가로채 합성 폼만 사용하며 실제 신청은 전송하지 않는다.

GitHub에는 스킬 지침·실행 도구·참고 구조·빈 템플릿·합성 테스트만 포함한다. 고객별 신청 파일, 공통정보, 승인 원문, manifest, ticket, 캡처, HAR, 응답 영수증과 로컬 실행 기록은 넣지 않는다. `.gitignore`는 보조 장치이므로 커밋 전에 파일 내용과 Excel 메타데이터도 확인한다. 버전은 `VERSION`과 frontmatter를 함께 관리한다.

배포 ZIP은 `python scripts/build_release.py --output OUTPUT.zip`으로 만든다. 배포 도구는 명시한 파일만 포함하며 기존 ZIP은 덮어쓰지 않는다. 원본 빈 템플릿에도 고객 값이나 작성자 정보가 들어가지 않았는지 별도로 점검한다. GitHub Actions는 실제 폼에 접근하지 않고 합성 pytest만 실행한다.
