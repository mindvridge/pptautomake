"""Office Add-in 백엔드 서버 실행

PowerPoint Office Add-in과 연동하기 위한 Flask 서버를 시작한다.
HTTPS로 실행되며, 기본 포트는 5000이다.

Usage:
    python run_server.py
    python run_server.py --port 8443
    python run_server.py --no-ssl
"""

from __future__ import annotations

import argparse
import sys


def main():
    parser = argparse.ArgumentParser(description='PPT AutoMake Add-in 서버')
    parser.add_argument('--port', type=int, default=5000, help='서버 포트 (기본: 5000)')
    parser.add_argument('--no-ssl', action='store_true', help='SSL 없이 HTTP로 실행')
    args = parser.parse_args()

    from src.api import create_app
    app = create_app()

    ssl_ctx = None if args.no_ssl else 'adhoc'

    protocol = 'http' if args.no_ssl else 'https'
    print(f'\nPPT AutoMake Add-in 서버')
    print(f'주소: {protocol}://localhost:{args.port}')
    print(f'Task Pane: {protocol}://localhost:{args.port}/')
    print(f'API Health: {protocol}://localhost:{args.port}/api/health')
    print()

    app.run(host='0.0.0.0', port=args.port, debug=True, ssl_context=ssl_ctx)


if __name__ == '__main__':
    main()
