# QA 문서 자동 생성

PG사 백오피스 개발 요청서(PPT/Jira)를 입력받아 사내 표준 QA 테스트 케이스 문서(xlsx)를 자동 생성하는 웹 앱.

내부적으로 **Claude Code CLI(`claude`)** 를 호출하므로 **별도 API 키가 필요 없고**, 각 사용자가 **자기 Claude(Pro/Max) 계정 크레딧**으로 생성합니다.

---

## 사전 준비 (각 PC에서 1회)

| # | 항목 | 설치 |
|---|---|---|
| 1 | **Python 3.12** | <https://www.python.org/downloads/> · 설치 첫 화면에서 **"Add Python to PATH" 체크** |
| 2 | **Node.js** | <https://nodejs.org> (Claude Code CLI 설치에 필요한 npm 포함) |
| 3 | **Claude Code CLI** | 터미널에서 `npm install -g @anthropic-ai/claude-code` |
| 4 | **Claude 계정 로그인** | 런처가 자동으로 안내합니다 (세션 없으면 로그인 창으로 이동). 수동: `claude auth login` |

> 4번 로그인은 **PC의 로그인 사용자 단위**로 저장됩니다(`~/.claude/.credentials.json`). 한 번 로그인하면 이후엔 그대로 사용됩니다.

---

## 설치 & 실행

폴더를 받은 뒤 **`run_qa_gen.bat` 더블클릭** 한 번이면 됩니다.

```
run_qa_gen.bat 더블클릭
  ├─ (최초 1회) .venv 가상환경 생성 + 의존성 자동 설치   ← 수 분 소요
  ├─ Claude CLI 설치 확인
  ├─ 계정 세션 확인 → 없으면 자동으로 로그인 절차로 이동
  └─ 웹 앱 실행 → 브라우저 자동 열림 (http://localhost:8501)
```

종료: 콘솔 창에서 `Ctrl+C` 또는 창 닫기.

---

## 사용법 (앱 화면)

1. 사이드바에 **Jira 티켓 번호** 입력 (예: `NXTGENBO-2612`)
2. (선택) 사이드바 **담당자 설정**에서 개발자/테스터/QA명 입력
3. **PPT 업로드**(복호화된 .pptx) 또는 **자유 입력칸**에 요청 내용·QA 지침 작성
   - 둘 다 써도 되고, 한쪽만 써도 됩니다.
   - 검색 조건 등은 하나하나 나열하지 않아도 됩니다 — 추상적으로 적어도 규칙대로 간결하게 생성됩니다.
4. **🚀 문서 생성** 클릭 (최대 5분)
5. 생성된 **xlsx 다운로드** → 사내 DRM 적용 후 공유

---

## QA 작성 규칙 커스터마이즈

생성 규칙은 **`rules.json`** 의 `rules` 배열에 한 줄씩 적습니다. 앱이 생성할 때마다 자동 반영됩니다.
(사내 DRM이 `.txt`를 암호화하는 문제 때문에 `.json`으로 관리)

```json
{
  "rules": [
    "수행절차는 보통 3~5단계 내외로 간결하게 쓴다.",
    "..."
  ]
}
```

규칙을 바꾸면 **다음 생성부터** 바로 적용됩니다(앱 재시작 불필요, CLI 별도 실행 불필요).

---

## 업데이트

사내 방화벽에서 GitHub raw/api는 열려 있으므로, 코드 갱신은 받은 방식에 맞게:

- **git으로 받았다면**: `git pull`
- **zip으로 받았다면**: 새 zip을 받아 덮어쓰기 (단, `.venv`·`outputs`·`rules.json`은 보존)

`.venv`는 배포물에 포함되지 않습니다(`.gitignore` 제외). 첫 실행 시 각 PC에서 자동 생성됩니다.

---

## 제약 / 주의

- **PPT 복호화·xlsx DRM 적용은 자동화 불가** — 사내 PC에서 수동 처리. 앱에는 **복호화된 .pptx**를 올립니다.
- PPT는 **텍스트만 추출**합니다(텍스트 상자·표·그룹 도형·발표자 노트). **이미지/화면 캡처 속 글자는 추출되지 않으므로** 자유 입력칸에 직접 적어주세요.
- 생성 입력 텍스트는 Claude 호출 시 전송됩니다. 민감 정보 주의.
- 각 사용자 본인 계정 크레딧을 사용합니다.

---

## 파일 구조

```
qa-gen/
├── run_qa_gen.bat        # ★ 설치 겸 실행 런처 (이것만 더블클릭)
├── app.py                # 웹 앱 (streamlit)
├── rules.json            # QA 작성 규칙 (여기만 수정)
├── settings.json         # 담당자명 등 로컬 설정
├── requirements.json     # 의존성 목록 (DRM 회피용 json)
├── template_builder.py   # xlsx 양식 빌더
├── ppt_extract.py        # PPT→텍스트 추출 CLI (보조)
├── build_xlsx.py         # JSON→xlsx 변환 CLI (보조)
├── qagen.py              # API 키 직접 호출 CLI (보조, 선택)
├── CLAUDE.md             # 프로젝트 컨텍스트 (Claude Code용)
├── inputs/               # 입력 보관
└── outputs/              # 생성된 xlsx 출력
```

> `qagen.py`는 Claude Code CLI 대신 `ANTHROPIC_API_KEY`로 직접 API를 호출하는 보조 경로입니다(종량제 과금). 일반 사용은 위 웹 앱으로 충분합니다.
