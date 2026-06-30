@echo off
chcp 65001 > nul
setlocal
REM ============================================================
REM QA 문서 자동 생성 - 설치 겸 실행 런처
REM 이 파일을 더블클릭하면:
REM   - 최초 실행: .venv 가상환경 생성 + 의존성 자동 설치
REM   - 매 실행: Claude 계정 세션 확인 (없으면 로그인 절차로 이동)
REM   - 이후: 웹 앱 실행
REM ============================================================

cd /d "%~dp0"

REM --- 1) Python 확인 ---
set "PYCMD="
where py >nul 2>nul && set "PYCMD=py -3"
if not defined PYCMD (
    where python >nul 2>nul && set "PYCMD=python"
)
if not defined PYCMD (
    echo [오류] Python 이 설치돼 있지 않습니다.
    echo        https://www.python.org/downloads/ 에서 Python 3.12 를 설치한 뒤
    echo        설치 첫 화면에서 "Add Python to PATH" 를 반드시 체크하세요.
    pause
    exit /b 1
)

REM --- 2) .venv 가상환경 (없으면 생성 + 의존성 설치) ---
if not exist ".venv\Scripts\streamlit.exe" (
    echo ============================================================
    echo  최초 1회 환경을 구성합니다. 수 분 걸릴 수 있습니다...
    echo ============================================================
    if not exist ".venv\Scripts\python.exe" (
        %PYCMD% -m venv .venv
        if errorlevel 1 (
            echo [오류] 가상환경 생성 실패.
            pause
            exit /b 1
        )
    )
    ".venv\Scripts\python.exe" -m pip install --upgrade pip
    ".venv\Scripts\python.exe" -c "import json,subprocess,sys;r=subprocess.run([sys.executable,'-m','pip','install',*json.load(open('requirements.json',encoding='utf-8'))['packages']]);sys.exit(r.returncode)"
    if errorlevel 1 (
        echo [오류] 의존성 설치 실패. 사내 네트워크/프록시 설정을 확인하세요.
        pause
        exit /b 1
    )
    echo [완료] 설치가 끝났습니다.
)

REM --- 3) Claude Code CLI 설치 확인 ---
where claude >nul 2>nul
if errorlevel 1 (
    echo [오류] Claude Code CLI(claude) 가 설치돼 있지 않습니다.
    echo        문서 생성에 필요합니다. 아래로 설치 후 다시 실행하세요:
    echo            npm install -g @anthropic-ai/claude-code
    echo        (npm 이 없으면 https://nodejs.org 에서 Node.js 부터 설치)
    pause
    exit /b 1
)

REM --- 4) Claude 계정 세션 확인 (없으면 로그인) ---
call :check_login
if not errorlevel 1 goto :session_ok

echo.
echo [로그인 필요] Claude 계정 세션이 없습니다. 로그인 절차로 이동합니다...
echo  - 본인 Claude(Pro/Max) 계정으로 로그인하세요. 브라우저가 열립니다.
echo.
claude auth login
echo.
call :check_login
if not errorlevel 1 goto :session_ok

echo [오류] 로그인이 확인되지 않았습니다. 창을 닫고 다시 실행해 주세요.
pause
exit /b 1

:session_ok
echo [확인] Claude 계정 세션 정상.
echo.

echo ============================================================
echo  QA 문서 자동 생성 웹 앱을 시작합니다...
echo  브라우저가 자동으로 열립니다. (http://localhost:8501)
echo  종료하려면 이 창에서 Ctrl+C 를 누르거나 창을 닫으세요.
echo ============================================================
echo.

".venv\Scripts\streamlit.exe" run app.py --server.port 8501

pause
exit /b 0

REM ============================================================
REM 로그인 상태 확인 서브루틴
REM   세션 있으면 errorlevel 0, 없으면 1
REM ============================================================
:check_login
claude auth status --json 2>nul | findstr /R /C:"loggedIn.*true" >nul 2>nul
exit /b %errorlevel%
