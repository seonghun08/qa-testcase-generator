# QA 문서 자동 생성 프로젝트

## 프로젝트 목표

PG사 백오피스 개발팀이 받는 **개발 요청서(PPT 또는 Jira 티켓)** 를 입력받아, 사내 표준 QA 테스트 케이스 문서(xlsx)를 자동 생성하는 도구를 만든다.

수동 작성에 들이는 시간을 줄이는 게 목적이며, **최종적으로 팀 전체가 공용으로 사용**할 수 있도록 한다.

---

## 동작 방식

### 주 사용 방식: 웹 앱 (`app.py`)

브라우저에서 PPT 업로드 또는 텍스트 붙여넣기 → 자동으로 QA xlsx 생성.
내부적으로 **Claude Code CLI (`claude -p`)** 를 subprocess로 호출하므로 **별도 API 키 불필요** (Pro 구독 크레딧 사용).

```bash
streamlit run app.py
```

### 보조 도구

| 파일 | 용도 | API 필요 |
|---|---|---|
| `app.py` | **메인 웹 앱** — PPT 업로드/텍스트 붙여넣기 → xlsx 생성 | ❌ (Claude Code CLI 사용) |
| `ppt_extract.py` | PPT → 텍스트 추출 CLI | ❌ |
| `build_xlsx.py` | JSON → xlsx 변환 CLI | ❌ |
| `template_builder.py` | xlsx 양식 빌더 (라이브러리) | ❌ |
| `qagen.py` | 단독 CLI (직접 API 호출) | ✅ `ANTHROPIC_API_KEY` |

---

## 환경 제약

| 항목 | 상태 |
|---|---|
| Jira(ITS) 도메인 | 사내 Jira 도메인 — robots.txt가 자동 크롤링 차단 (실제 주소는 사내 문서 참고) |
| 사내 PPT | 사내 DRM 솔루션으로 암호화됨 — 자동 복호화 불가, 사용자가 사내 PC에서 수동 복호화 필요 |
| 사내 파일 추출 | 회사 PC에서 파일을 별도로 꺼내면 DRM 암호화될 수 있음 — 파일 복사 대신 내용 직접 쓰기 방식 사용 |
| 생성된 xlsx | 사내 공유 전 DRM 적용 필요 (수동) |
| Claude Code CLI | `claude` 명령어가 PATH에 있어야 함 (Pro 구독) |

PPT 복호화와 DRM 적용은 **자동화 불가**. 항상 사용자가 사내 PC에서 직접 처리한 텍스트/파일을 입력으로 받는다.

---

## 입출력 사양

### 파일명 규칙 (고정)

```
TB_TC_{Jira번호}_{문서제목}.xlsx
```

- 접두사는 항상 `TB_` (구버전 호환을 위해 통일). BO/SHOP/OSS 같은 시스템 타입을 접두사로 쓰지 않음.
- 문서제목의 공백은 언더스코어로 변환

### 내부 제목 (B2 셀)

```
TB_TC_{Jira번호} {문서제목}
```

파일명과 동일하되 공백 그대로.

---

## 시스템 타입

| 약어 | 풀네임 | 표기 (테스트 대상 셀) |
|---|---|---|
| BO | 내부관리자 | `BO\n(내부관리자)` |
| SHOP | 가맹점관리자 | `SHOP\n(가맹점관리자)` |
| OSS | 운영지원시스템 | `OSS\n(운영지원시스템)` |

---

## TestCaseID 체계

```
TB_{기능그룹번호}_{순번}
예: TB_01_001, TB_01_002, TB_02_001
```

### 기능 그룹 분리 기준

- **화면 단위**로 다른 메뉴/탭이면 그룹 분리 (예: 메인 화면=01, 상세 팝업=02)
- **목록 조회와 상세 조회는 가능하면 별도 그룹으로 분리** (목록=검색/필터/정렬/페이징, 상세=항목 클릭 시 팝업·표시항목)
- 한 화면 안에서는 조회→등록→수정→삭제 순으로 같은 그룹 안에 묶음
- 결재/승인 같은 별도 워크플로우는 별도 그룹
- 모호하면 단일 그룹 `01`로 시작

순번은 그룹마다 `001`부터 새로 시작.

---

## QA 작성 관점 (핵심 원칙)

QA는 **비개발자(기획·운영 담당자)** 가 화면을 직접 클릭하며 **UI/UX가 정상 동작하는지** 확인하는 작업이다.

- 수행절차·기대결과에는 **화면에서 눈으로 확인 가능한 것만** 쓴다: 버튼/입력/필터/검색 동작, 목록·팝업·메시지·알럿 노출과 문구, 정렬·페이징, 검색 결과가 조건에 맞는지, 권한/시스템별 화면 차이.
- **제외**(개발자가 로그/DB 섹션에서 확인): 암호화·해시 처리 여부, 완전일치/범위검색 등 내부 매칭 알고리즘 설명, DB 테이블/컬럼명, 적재 테이블, API명, 쿼리, 트랜잭션 처리 등 화면 밖 내부 구현. 검색은 "조건 입력 시 결과가 조건에 맞게 표시되는가" 수준으로만 기술.

---

## xlsx 양식 구조

### 시트 3개

1. **문서변경이력** — 버전 관리 (#/변경내용/작성자/작성일)
2. **TestCase** — 실제 테스트 케이스 (메인 시트)
3. **참고_TestCase 결과** — Pass/Fixed/Fail/Blocked/N/A 정의 + 색상 가이드

### TestCase 시트 컬럼 (B~W, 총 22개)

```
B: TestCaseID         C: 우선순위 (기본 P0)    D: 테스트 대상
E: 사전조건           F: 단계                  G: 구분            H: 항목
I: 수행절차           J: 기대결과              K: 로그            L: DB
M-O: 개발자 (UI/로그/DB)
P-R: 테스터 (UI/로그/DB)
S-U: QA   (UI/로그/DB)
V: JIRA               W: 비고
```

### 행 레이아웃

```
Row 2: 제목 (B2:W2 병합, 연녹색 배경)
Row 3: 노란 띠 데코 (B3:W3 병합, 8pt 높이)
Row 4: 빈 행 (스페이서)
Row 5: 메인 헤더 (테스트 항목, 개발자/테스터/QA는 그룹 헤더)
Row 6: 서브 헤더 (단계/구분/항목, UI/로그/DB)
Row 7: 빈 행 (스페이서)
Row 8+: 테스트 케이스 데이터 (각 1행, 높이 230pt)
```

---

## 컬럼 기본값 규칙

| 컬럼 | 자동 채우기 값 |
|---|---|
| 우선순위 | `P0` 고정 |
| 테스트 대상 | 시스템 타입에 따라 `BO\n(내부관리자)` 등 |
| 사전조건 | `{계정유형} 계정 로그인\n\n{메뉴경로}` |
| 개발자 UI/로그/DB | 모두 `Pass` |
| 테스터 UI | `Pass` / 로그·DB | `N/A` |
| QA UI/로그/DB | **빈 칸** |
| JIRA | 빈 칸 (QA 작성자 영역) |
| 비고 | 빈 칸 (QA 작성자 영역) |

평가자 헤더의 괄호 안 이름은 비워둠 (`개발자 ()`, `테스터 ()`, `QA ()`).

---

## 상태값 색상 (UI/로그/DB 셀 배경)

| 상태 | 배경 | 글자 |
|---|---|---|
| Pass | `#C6EFCE` | `#006100` |
| Fail | `#FFC7CE` | `#9C0006` |
| Fixed | `#FFEB9C` | `#9C5700` |
| Blocked | `#E4DFEC` | `#5B3970` |
| N/A | `#D9D9D9` | `#595959` |

---

## 로그/DB 섹션 — 플레이스홀더 형식

`{}`로 감싼 값은 개발자가 직접 채우는 부분.
서버 IP·로그 경로·DB SID 등 **사내 접속 정보는 `log_db_config.json`(gitignore 대상)** 에 분리한다.
공개 저장소에는 `log_db_config.example.json`(플레이스홀더)만 포함된다. 템플릿의 `@..@` 마커가 config 값으로 치환된다.

### 로그 템플릿

```
- @LOG_APP_LABEL@ (개인계정 사용)
  ▷ @LOG_SERVER@
- 디렉토리 이동
  ▷ @LOG_DIR@
- 로그 검색
  ▷ grep -i "{mapper_id}" ./catalina.out
     > 출력된 로그내 쿼리 확인
  또는
  ▷ tail -f catalina.out
     > tail 상태에서 화면에서 조회 버튼 클릭
     > 출력된 로그내 쿼리 확인
```

### DB 템플릿 (SELECT/INSERT/UPDATE 유형별)

```
- @DB_LABEL@ (개인계정 사용)
  ▷ @DB_SERVER@ (SID:@DB_SID@)
- 로그에서 출력된 쿼리 복사 후 DB툴에서 실행 시 정상 확인 및 데이터 비교

  [쿼리 예시]
  SELECT * FROM {스키마}.{테이블명}
  WHERE {조건컬럼} = '{예시값}'

  [컬럼 설명]
  {컬럼명}: {설명}
```

INSERT는 `WHERE {PK컬럼} = '{자동채번값}'`, UPDATE는 수정 컬럼 + 상태 컬럼 형태로 분기.

---

## 스타일 (원본 사내 양식과 매칭)

| 항목 | 값 |
|---|---|
| 폰트 | 맑은 고딕 |
| 제목 폰트 크기 | 18pt bold |
| 헤더/본문 폰트 크기 | 9pt (헤더만 bold) |
| 헤더 배경색 | `#E2EEDA` (연녹색, theme:9 accent6 lighter 80%) |
| 노란 띠 색 | `#FFD966` (theme:7 accent4 lighter 60%) |
| 보더 | thin gray (`#808080`) |
| 모든 셀 세로 정렬 | `center` |
| 가로 정렬 | 단답형 셀(`center`) / 긴 텍스트(`left` + indent) |
| Wrap text | 모든 셀 활성화 |

---

## 색상 추출 로직 (참고)

원본 사내 xlsx의 색상은 OOXML 테마 색 + tint 형태로 저장되어 있어 직접 RGB가 안 보임. `apply_tint(hex, tint)` 공식으로 변환:

```python
def apply_tint(rgb_hex, tint):
    r, g, b = [int(rgb_hex[i:i+2], 16)/255 for i in (0, 2, 4)]
    if tint > 0:
        r, g, b = [c + (1-c)*tint for c in (r, g, b)]
    elif tint < 0:
        r, g, b = [c * (1+tint) for c in (r, g, b)]
    return f'{int(r*255):02X}{int(g*255):02X}{int(b*255):02X}'

# 헤더: theme:9 (70AD47) + tint 0.8 → E2EEDA
# 노란 띠: theme:7 (FFC000) + tint 0.4 → FFD966
```

---

## 디렉토리 구조

```
qa-gen/
├── CLAUDE.md                 # 이 파일 (Claude Code가 자동으로 읽음)
├── README.md                 # 사람용 사용법
├── requirements.txt          # Python 의존성
├── app.py                    # 메인 웹 앱 (streamlit run app.py)
├── template_builder.py       # 양식 빌더 (라이브러리)
├── ppt_extract.py            # PPT → 텍스트 추출 CLI
├── build_xlsx.py             # JSON → xlsx 변환 CLI
├── qagen.py                  # 단독 CLI (API 키 필요, 선택 사용)
├── samples/
│   └── TB_TC_PGBO-1001_*.xlsx  # 합의된 양식 샘플
├── inputs/                   # 작업 시 텍스트 입력 보관
└── outputs/                  # 생성된 xlsx 출력
```

---

## 현재 진행 상태

- [x] 양식 분석 완료 (사내 BO/OSS/TB QA 문서 6개 검토)
- [x] 양식 규칙 합의 완료 (위 사양)
- [x] 양식 빌더 스크립트 완성 (`template_builder.py`)
- [x] 가상 시나리오 (PGBO-1001 가맹점 정산) 샘플 xlsx 생성
- [x] CLI 도구 (`qagen.py`) 구현 (API 키 직접 호출 방식)
- [x] 웹 앱 (`app.py`) 구현 (Claude Code CLI 기반, API 비용 없음)
- [x] PPT 텍스트 추출기 (`ppt_extract.py`)
- [x] JSON → xlsx 변환기 (`build_xlsx.py`)
- [ ] PPT 실파일로 추출 테스트
- [ ] 실제 Jira 티켓으로 end-to-end 테스트
- [ ] 팀 배포용 패키징

---

## Claude Code에게 (다음 작업 시작 시)

이 프로젝트에서 Claude Code 세션 내에서 직접 QA 문서를 생성해달라는 요청이 올 수 있습니다.

요청 처리 흐름:

1. 사용자가 `inputs/` 에 있는 텍스트나 PPT 추출 결과를 읽어달라고 요청
2. 텍스트를 분석하여 테스트 케이스 JSON 생성
3. JSON 파일을 `outputs/` 에 저장
4. `build_xlsx.py`를 실행하여 xlsx 생성

또는 사용자가 `streamlit run app.py`로 웹 앱을 직접 사용할 수도 있습니다.

생성 시 이 문서의 모든 규칙(컬럼 기본값, 시스템 타입 표기, 상태값 색상, 로그/DB 플레이스홀더, TestCaseID 체계 등)을 준수할 것.
