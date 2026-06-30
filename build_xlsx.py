#!/usr/bin/env python3
"""
JSON → xlsx 변환기

Claude Code가 생성한 테스트 케이스 JSON 파일을 읽어 xlsx 를 생성한다.
API 호출 없음.

사용법:
  python build_xlsx.py PGBO-1001 --json test_cases.json
  python build_xlsx.py PGBO-1001 --json test_cases.json --system-type SHOP
"""
import argparse
import json
import os
import sys

from template_builder import build_workbook, save_workbook


def main():
    parser = argparse.ArgumentParser(description="JSON → QA xlsx 변환기")
    parser.add_argument('jira_num', help='Jira 티켓 번호 (예: PGBO-1001)')
    parser.add_argument('--json', required=True, help='테스트 케이스 JSON 파일 경로')
    parser.add_argument('--system-type', choices=['BO', 'SHOP', 'OSS'], default='BO',
                        help='시스템 타입 (기본값: BO)')
    parser.add_argument('--output-dir', default=None,
                        help='출력 디렉토리 (기본값: outputs/)')
    args = parser.parse_args()

    if not os.path.isfile(args.json):
        print(f"[오류] JSON 파일을 찾을 수 없습니다: {args.json}", file=sys.stderr)
        sys.exit(1)

    with open(args.json, 'r', encoding='utf-8') as f:
        data = json.load(f)

    # JSON 최상위가 {"title": ..., "test_cases": [...]} 또는 [...] 둘 다 지원
    if isinstance(data, list):
        title = args.jira_num
        system_type = args.system_type
        test_cases = data
    elif isinstance(data, dict):
        title = data.get('title', args.jira_num)
        system_type = data.get('system_type', args.system_type)
        test_cases = data.get('test_cases', data.get('cases', []))
    else:
        print("[오류] JSON 형식이 올바르지 않습니다.", file=sys.stderr)
        sys.exit(1)

    if not test_cases:
        print("[오류] 테스트 케이스가 없습니다.", file=sys.stderr)
        sys.exit(1)

    print(f"[*] Jira: {args.jira_num}")
    print(f"[*] 제목: {title}")
    print(f"[*] 시스템: {system_type}")
    print(f"[*] 케이스: {len(test_cases)}건")

    wb, file_name = build_workbook(args.jira_num, title, test_cases, system_type=system_type)

    out_dir = args.output_dir or os.path.join(os.path.dirname(os.path.abspath(__file__)), 'outputs')
    out_path = save_workbook(wb, file_name, out_dir)

    print(f"\n[완료] {out_path}")
    print(f"[완료] {os.path.getsize(out_path)} bytes")


if __name__ == '__main__':
    main()
