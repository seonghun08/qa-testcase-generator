"""
환경 점검 & 설치 가이드 (Streamlit)

QA 문서 자동 생성 앱 실행에 필요한 구성요소를 점검하고,
빠진 항목은 '웹 다운로드 / 터미널 명령' 두 방법으로 단계별 안내한다.

  - 단독 실행(테스트):  streamlit run setup_guide.py
  - 앱에서 호출:        from setup_guide import render_setup_check
"""
import json
import shutil
import subprocess
import sys

import streamlit as st


# ============================================================
# 점검 로직
# ============================================================

def _run(args, timeout=15):
    """명령을 실행해 결과를 반환한다. (.cmd 호환을 위해 shell 경유)"""
    cmd = " ".join(f'"{a}"' if (" " in a or "\\" in a) else a for a in args)
    try:
        return subprocess.run(
            cmd, capture_output=True, text=True, timeout=timeout,
            encoding="utf-8", errors="replace", shell=True,
        )
    except Exception:
        return None


def check_python():
    v = sys.version_info
    ok = v >= (3, 10)
    detail = f"{v.major}.{v.minor}.{v.micro}" + ("" if ok else " (3.10 이상 권장)")
    return {"key": "python", "name": "Python", "ok": ok, "detail": detail}


def _check_tool(key, name, exe, version_args=("--version",), missing_hint="미설치"):
    path = shutil.which(exe)
    if not path:
        return {"key": key, "name": name, "ok": False, "detail": missing_hint}
    r = _run([path, *version_args])
    ver = (r.stdout.strip() if r and r.stdout and r.stdout.strip() else "설치됨")
    return {"key": key, "name": name, "ok": True, "detail": ver}


def check_node():
    return _check_tool("node", "Node.js", "node")


def check_npm():
    return _check_tool("npm", "npm", "npm",
                       missing_hint="미설치 (Node.js 설치 시 함께 설치됨)")


def check_claude():
    return _check_tool("claude", "Claude Code CLI", "claude")


def check_login():
    path = shutil.which("claude")
    if not path:
        return {"key": "login", "name": "Claude 로그인 세션", "ok": False,
                "detail": "Claude Code CLI 설치 후 가능"}
    r = _run([path, "auth", "status", "--json"])
    logged, email = False, ""
    if r and r.stdout and "{" in r.stdout:
        try:
            data = json.loads(r.stdout[r.stdout.index("{"):r.stdout.rindex("}") + 1])
            logged = bool(data.get("loggedIn"))
            email = data.get("email", "") or ""
        except Exception:
            logged = '"loggedIn": true' in r.stdout
    detail = f"로그인됨 ({email})" if logged else "로그인 안 됨"
    return {"key": "login", "name": "Claude 로그인 세션", "ok": logged, "detail": detail}


def run_all_checks():
    """설치 의존 순서대로 점검 결과를 반환한다."""
    return [check_python(), check_node(), check_npm(), check_claude(), check_login()]


# ============================================================
# 설치 가이드 콘텐츠
# ============================================================

GUIDES = {
    "python": {
        "title": "1) Python 설치",
        "web": {
            "label": "🌐 python.org 다운로드 페이지 열기",
            "url": "https://www.python.org/downloads/",
            "steps": [
                "위 버튼으로 사이트를 열고 노란 **Download Python 3.12.x** 버튼 클릭",
                "받은 설치 파일(.exe) 실행",
                "설치 첫 화면 맨 아래 **'Add Python to PATH'** 체크 — ★ 꼭 체크",
                "**Install Now** 클릭 → 완료 후 아래 '🔄 다시 점검'",
            ],
        },
        "terminal": {
            "intro": "Windows 터미널(PowerShell)에서 winget 사용 시:",
            "cmd": "winget install -e --id Python.Python.3.12",
            "note": "설치 후 터미널/런처를 새로 열어야 PATH가 반영됩니다.",
        },
    },
    "node": {
        "title": "2) Node.js 설치 (npm 포함)",
        "web": {
            "label": "🌐 nodejs.org 열기",
            "url": "https://nodejs.org",
            "steps": [
                "위 버튼으로 사이트를 열고 **LTS** 버튼으로 다운로드",
                "받은 설치 파일 실행 → 기본값 그대로 **Next** 로 진행",
                "완료 후 아래 '🔄 다시 점검'",
            ],
        },
        "terminal": {
            "intro": "winget 사용 시:",
            "cmd": "winget install -e --id OpenJS.NodeJS.LTS",
            "note": "npm은 Node.js에 포함되어 함께 설치됩니다.",
        },
    },
    "claude": {
        "title": "3) Claude Code CLI 설치",
        "web": None,
        "terminal": {
            "intro": "Node.js 설치 후, 터미널에서:",
            "cmd": "npm install -g @anthropic-ai/claude-code",
            "note": "설치 확인:  claude --version",
        },
    },
    "login": {
        "title": "4) Claude 계정 로그인",
        "web": None,
        "terminal": {
            "intro": "터미널에서 아래 실행 → 열리는 브라우저에서 본인 Claude(Pro/Max) 계정으로 로그인:",
            "cmd": "claude auth login",
            "note": "또는 run_qa_gen.bat 을 다시 실행하면 자동으로 로그인 절차로 이동합니다.",
        },
    },
}


def _render_guide(check):
    # npm은 Node.js 가이드를 공유한다.
    guide = GUIDES.get("node" if check["key"] == "npm" else check["key"])
    if not guide:
        return

    st.markdown(f"#### {guide['title']}")

    labels, renderers = [], []
    if guide.get("web"):
        labels.append("🌐 웹에서 다운로드")
        renderers.append(("web", guide["web"]))
    if guide.get("terminal"):
        labels.append("⌨️ 터미널 명령")
        renderers.append(("terminal", guide["terminal"]))

    tabs = st.tabs(labels)
    for tab, (kind, data) in zip(tabs, renderers):
        with tab:
            if kind == "web":
                st.link_button(data["label"], data["url"])
                for i, step in enumerate(data["steps"], 1):
                    st.markdown(f"{i}. {step}")
            else:
                st.markdown(data["intro"])
                st.code(data["cmd"], language="bash")
                if data.get("note"):
                    st.caption(data["note"])


# ============================================================
# 메인 렌더 (앱에서 import 해 호출)
# ============================================================

def render_setup_check():
    st.subheader("⚙️ 환경 점검 & 설치 가이드")
    st.caption("앱 실행·문서 생성에 필요한 구성요소를 점검합니다. "
               "❌ 항목은 아래 안내대로 설치한 뒤 '🔄 다시 점검'을 누르세요.")

    if st.button("🔄 다시 점검"):
        st.rerun()

    with st.spinner("환경을 점검하는 중입니다... (Node · npm · Claude CLI · 로그인 세션 확인)"):
        checks = run_all_checks()

    # --- 상태 요약 ---
    for c in checks:
        col1, col2 = st.columns([0.35, 0.65])
        col1.markdown(f"{'✅' if c['ok'] else '❌'} **{c['name']}**")
        col2.markdown(c["detail"])

    missing = [c for c in checks if not c["ok"]]
    if not missing:
        st.success("모든 항목이 준비됐습니다. 문서 생성을 사용할 수 있습니다. 🎉")
    else:
        st.divider()
        st.markdown("### 설치 안내 — 위에서부터 순서대로 진행하세요")
        st.info("앞 단계(Python → Node.js → Claude CLI → 로그인)를 순서대로 설치해야 합니다. "
                "한 단계 설치 후 '🔄 다시 점검'으로 확인하며 진행하세요.")
        for c in missing:
            _render_guide(c)
            st.divider()

    # 모든 항목이 설치돼 있어도 안내 UI 모습을 미리 볼 수 있도록 데모 제공
    with st.expander("📖 설치 안내 전체 미리보기 (데모 — 현재 설치 상태와 무관)"):
        st.caption("실제로는 위 점검에서 ❌ 인 항목만 안내가 나옵니다. "
                   "아래는 전체 안내 화면을 미리 보는 용도입니다.")
        for key in ["python", "node", "claude", "login"]:
            _render_guide({"key": key})
            st.divider()
    return checks


def main():
    st.set_page_config(page_title="환경 점검 & 설치 가이드", page_icon="⚙️", layout="centered")
    st.title("⚙️ QA 자동 생성 — 환경 점검 & 설치 가이드")
    render_setup_check()


if __name__ == "__main__":
    main()
