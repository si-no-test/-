#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
send_email.py — 透過 SendGrid 的 HTTPS API 把簡報當附件寄出。
(雲端沙箱封鎖 SMTP，所以改走 HTTPS；本程式只用 Python 內建套件，免安裝。)

需要的環境變數：
    SENDGRID_API_KEY   SendGrid 的 API 金鑰
    MAIL_FROM          已在 SendGrid 完成驗證的「寄件人」信箱
    MAIL_TO            收件人；多個用逗號分隔

用法:
    python send_email.py output.pptx "銀行新聞快覽 2026/06/10-2026/06/16"
"""
import os
import sys
import json
import base64
import urllib.request
import urllib.error
from datetime import date

API_URL = "https://api.sendgrid.com/v3/mail/send"
PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"


def main():
    if len(sys.argv) < 2:
        print("用法: python send_email.py <pptx檔> [主旨]"); sys.exit(1)
    path = sys.argv[1]
    subject = sys.argv[2] if len(sys.argv) > 2 else f"銀行新聞快覽 {date.today():%Y/%m/%d}"

    api_key = os.environ.get("SENDGRID_API_KEY")
    mail_from = os.environ.get("MAIL_FROM")
    mail_to = os.environ.get("MAIL_TO")
    if not (api_key and mail_from and mail_to):
        print("缺少環境變數 SENDGRID_API_KEY / MAIL_FROM / MAIL_TO"); sys.exit(1)
    recipients = [{"email": a.strip()} for a in mail_to.split(",") if a.strip()]

    with open(path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode("ascii")

    payload = {
        "personalizations": [{"to": recipients}],
        "from": {"email": mail_from, "name": "銀行新聞快覽"},
        "subject": subject,
        "content": [{"type": "text/plain",
                     "value": "附件為今日自動產生的銀行新聞快覽簡報，請查收。\n（本信由自動排程寄出）"}],
        "attachments": [{
            "content": b64,
            "type": PPTX_MIME,
            "filename": os.path.basename(path),
            "disposition": "attachment",
        }],
    }

    req = urllib.request.Request(
        API_URL,
        data=json.dumps(payload).encode("utf-8"),
        headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            print(f"已寄出（HTTP {resp.status}）-> {[r['email'] for r in recipients]}")
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", "ignore")
        print(f"寄信失敗 HTTP {e.code}: {body}")
        # 常見：403 多半是寄件人(MAIL_FROM)還沒在 SendGrid 完成驗證
        sys.exit(1)


if __name__ == "__main__":
    main()
