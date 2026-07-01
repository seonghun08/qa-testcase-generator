# QA 문서 자동 생성

PG사 백오피스 개발 요청서(PPT/Jira)를 입력받아 사내 표준 QA 테스트 케이스 문서(xlsx)를 자동 생성하는 Web App 입니다.

내부적으로 **Claude Code CLI(`claude`)** 를 호출하므로 **별도 API 키가 필요 없고**, 각 사용자가 **자기 Claude(Pro/Max) 계정 크레딧**으로 생성합니다.

---

## 1. 사전 준비

| # | 항목 | 설치 |
|---|---|---|
| 1 | **Python 3.12** | <https://www.python.org/downloads/> · 설치 첫 화면에서 **"Add Python to PATH" 체크** |
| 2 | **Node.js** | <https://nodejs.org> (Claude Code CLI 설치에 필요한 npm 포함) |
| 3 | **Claude Code CLI** | 터미널에서 `npm install -g @anthropic-ai/claude-code` |
| 4 | **Claude 계정 로그인** | 런처가 자동으로 안내합니다 (세션 없으면 로그인 창으로 이동). 수동: `claude auth login` |

> 4번 로그인은 **PC의 로그인 사용자 단위**로 저장됩니다(`~/.claude/.credentials.json`). 한 번 로그인하면 이후엔 그대로 사용됩니다.
>
> (선택) 생성 xlsx의 **로그/DB 접속 안내**를 실제 값으로 채우려면 `log_db_config.example.json`을 복사해 `log_db_config.json`으로 저장한 뒤 사내 서버 값을 입력하세요. 없으면 플레이스홀더로 동작합니다.

---

## 2. 설치 & 실행

폴더를 받은 뒤 `run_qa_gen.bat`를 통해 local 서버를 실행합니다.

```
run_qa_gen.bat 더블클릭
  ├─ (최초 1회) .venv 가상환경 생성 + 의존성 자동 설치
  ├─ Claude CLI 설치 확인
  ├─ 계정 세션 확인 → 없으면 자동으로 로그인 절차로 이동
  └─ 웹 앱 실행 → 브라우저 자동 열림 (http://localhost:8501)
```

종료: 콘솔 창에서 `Ctrl+C` 또는 창 닫기.

---

## 3. 사용법 (App 화면)

1. **📋 QA 문서 생성** 탭 상단에 **Jira 티켓 번호** 입력 (ex:`TEST-2612`)
2. (선택) 사이드바 **담당자 설정**에서 개발자/테스터/QA명 입력
3. **PPT 업로드**(.pptx) 또는 **자유 입력칸**에 요청 내용 · QA 지침 작성
   - 둘 다 써도 되고, 한쪽만 써도 됩니다.
   - 검색 조건 등은 하나하나 나열하지 않아도 됩니다 — 추상적으로 적어도 규칙대로 간결하게 생성됩니다.
4. **🚀 문서 생성** 클릭 (claude cli 요청)
5. 생성된 **xlsx 다운로드** → 사내 DRM 적용 후 공유

---

## 4. QA 작성 규칙 커스터마이즈

생성 규칙은 **`rules.json`**의 `rules`를 통해 생성할 때마다 자동 반영됩니다.

직접 파일을 열어 편집할 수 있으며, App에서 작성 규칙 편집을 통해 관리할 수 있습니다.

```json
{
  "rules": [
    "수행절차는 보통 3~5단계 내외로 간결하게 쓴다.",
    "..."
  ]
}
```

규칙을 바꾸면 **다음 생성부터** 바로 적용됩니다. (앱 재시작 불필요, CLI 별도 실행 불필요)

## 5. 제약 / 주의

- PPT는 **텍스트만 추출**합니다. (텍스트 상자·표·그룹 도형·발표자 노트, 토큰 사용 최소화) 
- **이미지/화면 캡처 속 글자는 추출되지 않으므로** 자유 입력칸에 직접 적어주세요.
- 생성 입력 텍스트는 Claude 호출 시 전송됩니다. 민감 정보 주의.
- 각 사용자 본인 계정 크레딧을 사용합니다.

---

## 파일 구조

```
qa-gen/
├── run_qa_gen.bat             # ★ 설치 겸 실행 런처 (이것만 더블클릭)
├── app.py                     # 웹 앱 (streamlit) — 생성/이력/규칙 편집/환경 점검 탭
├── setup_guide.py             # 환경 점검 & 설치 가이드 (앱에서 import)
├── template_builder.py        # xlsx 양식 빌더
├── ppt_extract.py             # PPT→텍스트 추출 CLI (보조)
├── build_xlsx.py              # JSON→xlsx 변환 CLI (보조)
├── qagen.py                   # API 키 직접 호출 CLI (보조, 선택)
├── rules.json                 # QA 작성 규칙 (앱 '작성 규칙 편집' 탭에서도 수정)
├── log_db_config.example.json # 로그/DB 접속정보 템플릿 → 복사해 log_db_config.json 생성
├── requirements.json          # 의존성 목록
├── CLAUDE.md                  # 프로젝트 컨텍스트 (Claude Code용)
├── inputs/                    # 입력 보관 (gitignore)
└── outputs/                   # 생성된 xlsx 출력 (gitignore)

# 로컬 전용(gitignore, 각자 생성): settings.json, log_db_config.json, .venv/
```

> `qagen.py`는 Claude Code CLI 대신 `ANTHROPIC_API_KEY`로 직접 API를 호출하는 보조 경로입니다(종량제 과금). 일반 사용은 위 웹 앱으로 충분합니다.
