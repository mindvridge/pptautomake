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
import os
from pathlib import Path


def _generate_self_signed_cert(cert_dir: Path) -> tuple[str, str]:
    """자체 서명 SSL 인증서를 생성한다.

    Office Add-in은 HTTPS가 필수이므로 개발용 인증서를 자동 생성한다.
    """
    cert_file = cert_dir / 'cert.pem'
    key_file = cert_dir / 'key.pem'

    if cert_file.exists() and key_file.exists():
        return str(cert_file), str(key_file)

    cert_dir.mkdir(parents=True, exist_ok=True)

    from OpenSSL import crypto

    # 키 생성
    key = crypto.PKey()
    key.generate_key(crypto.TYPE_RSA, 2048)

    # 인증서 생성
    cert = crypto.X509()
    subject = cert.get_subject()
    subject.C = 'KR'
    subject.O = 'PPT AutoMake Dev'
    subject.CN = 'localhost'

    cert.set_serial_number(1000)
    cert.gmtime_adj_notBefore(0)
    cert.gmtime_adj_notAfter(365 * 24 * 60 * 60)  # 1년
    cert.set_issuer(subject)
    cert.set_pubkey(key)
    cert.sign(key, 'sha256')

    with open(str(cert_file), 'wb') as f:
        f.write(crypto.dump_certificate(crypto.FILETYPE_PEM, cert))
    with open(str(key_file), 'wb') as f:
        f.write(crypto.dump_privatekey(crypto.FILETYPE_PEM, key))

    print(f'SSL 인증서 생성 완료: {cert_dir}')
    return str(cert_file), str(key_file)


def main():
    parser = argparse.ArgumentParser(description='PPT AutoMake Add-in 서버')
    parser.add_argument('--port', type=int, default=5000, help='서버 포트 (기본: 5000)')
    parser.add_argument('--no-ssl', action='store_true', help='SSL 없이 HTTP로 실행')
    args = parser.parse_args()

    from src.api import create_app
    app = create_app()

    ssl_ctx = None
    if not args.no_ssl:
        cert_dir = Path(__file__).parent / 'certs'
        cert_file, key_file = _generate_self_signed_cert(cert_dir)
        ssl_ctx = (cert_file, key_file)

    protocol = 'http' if args.no_ssl else 'https'
    print(f'\nPPT AutoMake Add-in 서버')
    print(f'주소: {protocol}://localhost:{args.port}')
    print(f'Task Pane: {protocol}://localhost:{args.port}/')
    print(f'API Health: {protocol}://localhost:{args.port}/api/health')
    if not args.no_ssl:
        print(f'\n[참고] 브라우저에서 자체 서명 인증서 경고가 나타나면')
        print(f'       "고급" → "localhost(안전하지 않음)으로 이동"을 클릭하세요.')
    print()

    app.run(host='0.0.0.0', port=args.port, debug=True, ssl_context=ssl_ctx)


if __name__ == '__main__':
    main()
