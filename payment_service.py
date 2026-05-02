import base64
import io
import json
import os
import secrets
import sqlite3
import time
import urllib.parse
import urllib.request
from datetime import datetime
from pathlib import Path
from typing import Dict, Any, Optional

import qrcode
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import padding
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from cryptography import x509


BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"
DB_PATH = DATA_DIR / "payments.sqlite3"
DATA_DIR.mkdir(exist_ok=True)


def init_payment_db():
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS payments (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                job_id TEXT NOT NULL,
                provider TEXT NOT NULL,
                out_trade_no TEXT NOT NULL UNIQUE,
                provider_trade_no TEXT,
                amount_cents INTEGER NOT NULL,
                status TEXT NOT NULL,
                created_at INTEGER NOT NULL,
                paid_at INTEGER
            )
            """
        )
        conn.execute("CREATE INDEX IF NOT EXISTS idx_payments_job_id ON payments(job_id)")
        conn.commit()


def paywall_enabled() -> bool:
    return os.getenv("PAYWALL_ENABLED", "false").lower() in {"1", "true", "yes", "on"}


def product_amount_cents() -> int:
    return int(os.getenv("PRICE_CENTS", "1990"))


def product_amount_yuan() -> str:
    return f"{product_amount_cents() / 100:.2f}"


def product_name() -> str:
    return os.getenv("PRODUCT_NAME", "论文格式终检修复版 Word")


def public_base_url() -> str:
    return os.getenv("PUBLIC_BASE_URL", "").rstrip("/")


def get_payment(job_id: str) -> Optional[Dict[str, Any]]:
    init_payment_db()
    with sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        row = conn.execute(
            "SELECT * FROM payments WHERE job_id=? ORDER BY id DESC LIMIT 1", (job_id,)
        ).fetchone()
        return dict(row) if row else None


def is_job_paid(job_id: str) -> bool:
    init_payment_db()
    with sqlite3.connect(DB_PATH) as conn:
        row = conn.execute(
            "SELECT 1 FROM payments WHERE job_id=? AND status='PAID' LIMIT 1", (job_id,)
        ).fetchone()
        return row is not None


def create_local_payment(job_id: str, provider: str) -> Dict[str, Any]:
    init_payment_db()
    out_trade_no = "LW" + datetime.now().strftime("%Y%m%d%H%M%S") + secrets.token_hex(5).upper()
    now = int(time.time())
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute(
            "INSERT INTO payments(job_id, provider, out_trade_no, amount_cents, status, created_at) VALUES(?,?,?,?,?,?)",
            (job_id, provider, out_trade_no, product_amount_cents(), "PENDING", now),
        )
        conn.commit()
    return {"job_id": job_id, "provider": provider, "out_trade_no": out_trade_no, "amount_cents": product_amount_cents(), "amount_yuan": product_amount_yuan(), "status": "PENDING"}


def mark_paid(out_trade_no: str, provider_trade_no: str = "") -> bool:
    init_payment_db()
    now = int(time.time())
    with sqlite3.connect(DB_PATH) as conn:
        cur = conn.execute(
            "UPDATE payments SET status='PAID', provider_trade_no=?, paid_at=? WHERE out_trade_no=?",
            (provider_trade_no, now, out_trade_no),
        )
        conn.commit()
        return cur.rowcount > 0


def require_env(name: str) -> str:
    value = os.getenv(name, "").strip()
    if not value:
        raise RuntimeError(f"缺少支付配置：{name}")
    return value


def _read_key_text(path_env: str, content_env: str = "") -> str:
    content = os.getenv(content_env, "").strip() if content_env else ""
    if content:
        return content.replace("\\n", "\n")
    path = require_env(path_env)
    return Path(path).read_text(encoding="utf-8")


def _normalize_private_pem(raw: str) -> bytes:
    raw = raw.strip().replace("\\n", "\n")
    if "BEGIN" not in raw:
        raw = "-----BEGIN PRIVATE KEY-----\n" + raw + "\n-----END PRIVATE KEY-----"
    return raw.encode("utf-8")


def _normalize_public_pem(raw: str) -> bytes:
    raw = raw.strip().replace("\\n", "\n")
    if "BEGIN" not in raw:
        raw = "-----BEGIN PUBLIC KEY-----\n" + raw + "\n-----END PUBLIC KEY-----"
    return raw.encode("utf-8")


def load_private_key(path_env: str, content_env: str = ""):
    raw = _read_key_text(path_env, content_env)
    return serialization.load_pem_private_key(_normalize_private_pem(raw), password=None)


def load_public_key_from_path_or_content(path_env: str, content_env: str = ""):
    raw = _read_key_text(path_env, content_env)
    return serialization.load_pem_public_key(_normalize_public_pem(raw))


def qr_data_uri(value: str) -> str:
    img = qrcode.make(value)
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode("utf-8")


# ---------------- Alipay ----------------

def alipay_gateway() -> str:
    return os.getenv("ALIPAY_GATEWAY", "https://openapi.alipay.com/gateway.do")


def alipay_sign(params: Dict[str, Any]) -> str:
    private_key = load_private_key("ALIPAY_APP_PRIVATE_KEY_PATH", "ALIPAY_APP_PRIVATE_KEY")
    sign_content = "&".join(f"{k}={params[k]}" for k in sorted(params) if params[k] not in {None, ""})
    signature = private_key.sign(sign_content.encode("utf-8"), padding.PKCS1v15(), hashes.SHA256())
    return base64.b64encode(signature).decode("utf-8")


def create_alipay_page_payment(job_id: str) -> Dict[str, Any]:
    if not public_base_url():
        raise RuntimeError("缺少支付配置：PUBLIC_BASE_URL")
    payment = create_local_payment(job_id, "alipay")
    biz_content = {
        "out_trade_no": payment["out_trade_no"],
        "total_amount": product_amount_yuan(),
        "subject": product_name(),
        "product_code": "FAST_INSTANT_TRADE_PAY",
    }
    params = {
        "app_id": require_env("ALIPAY_APP_ID"),
        "method": "alipay.trade.page.pay",
        "format": "JSON",
        "charset": "utf-8",
        "sign_type": "RSA2",
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "version": "1.0",
        "notify_url": public_base_url() + "/api/payment/alipay/notify",
        "return_url": public_base_url() + "/",
        "biz_content": json.dumps(biz_content, ensure_ascii=False, separators=(",", ":")),
    }
    params["sign"] = alipay_sign(params)
    pay_url = alipay_gateway() + "?" + urllib.parse.urlencode(params)
    return {**payment, "pay_url": pay_url, "method": "redirect"}


def verify_alipay_notify(params: Dict[str, str]) -> bool:
    sign = params.get("sign", "")
    if not sign:
        return False
    filtered = {k: v for k, v in params.items() if k not in {"sign", "sign_type"} and v != ""}
    sign_content = "&".join(f"{k}={filtered[k]}" for k in sorted(filtered))
    public_key = load_public_key_from_path_or_content("ALIPAY_PUBLIC_KEY_PATH", "ALIPAY_PUBLIC_KEY")
    public_key.verify(base64.b64decode(sign), sign_content.encode("utf-8"), padding.PKCS1v15(), hashes.SHA256())
    return True


def handle_alipay_notify(params: Dict[str, str]) -> bool:
    try:
        if not verify_alipay_notify(params):
            return False
        if params.get("trade_status") not in {"TRADE_SUCCESS", "TRADE_FINISHED"}:
            return False
        return mark_paid(params.get("out_trade_no", ""), params.get("trade_no", ""))
    except Exception:
        return False


# ---------------- WeChat Pay v3 Native ----------------

def wechat_sign(method: str, url_path: str, body: str, nonce: str, timestamp: str) -> str:
    private_key = load_private_key("WECHAT_PRIVATE_KEY_PATH", "WECHAT_PRIVATE_KEY")
    message = f"{method}\n{url_path}\n{timestamp}\n{nonce}\n{body}\n"
    signature = private_key.sign(message.encode("utf-8"), padding.PKCS1v15(), hashes.SHA256())
    return base64.b64encode(signature).decode("utf-8")


def wechat_authorization(method: str, url_path: str, body: str) -> str:
    mch_id = require_env("WECHAT_MCH_ID")
    serial_no = require_env("WECHAT_SERIAL_NO")
    nonce = secrets.token_urlsafe(16)
    timestamp = str(int(time.time()))
    signature = wechat_sign(method, url_path, body, nonce, timestamp)
    return (
        'WECHATPAY2-SHA256-RSA2048 '
        f'mchid="{mch_id}",nonce_str="{nonce}",signature="{signature}",'
        f'timestamp="{timestamp}",serial_no="{serial_no}"'
    )


def create_wechat_native_payment(job_id: str) -> Dict[str, Any]:
    if not public_base_url():
        raise RuntimeError("缺少支付配置：PUBLIC_BASE_URL")
    payment = create_local_payment(job_id, "wechat")
    payload = {
        "appid": require_env("WECHAT_APP_ID"),
        "mchid": require_env("WECHAT_MCH_ID"),
        "description": product_name(),
        "out_trade_no": payment["out_trade_no"],
        "notify_url": public_base_url() + "/api/payment/wechat/notify",
        "amount": {"total": product_amount_cents(), "currency": "CNY"},
    }
    body = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    url_path = "/v3/pay/transactions/native"
    request = urllib.request.Request(
        "https://api.mch.weixin.qq.com" + url_path,
        data=body.encode("utf-8"),
        method="POST",
        headers={"Content-Type": "application/json", "Accept": "application/json", "Authorization": wechat_authorization("POST", url_path, body)},
    )
    with urllib.request.urlopen(request, timeout=12) as response:
        data = json.loads(response.read().decode("utf-8"))
    code_url = data.get("code_url")
    if not code_url:
        raise RuntimeError("微信支付未返回二维码链接")
    return {**payment, "code_url": code_url, "qr_data_uri": qr_data_uri(code_url), "method": "native"}


def load_wechat_platform_public_key():
    cert_content = os.getenv("WECHAT_PLATFORM_CERT", "").strip().replace("\\n", "\n")
    if cert_content:
        cert = x509.load_pem_x509_certificate(cert_content.encode("utf-8"))
        return cert.public_key()
    cert_path = require_env("WECHAT_PLATFORM_CERT_PATH")
    cert = x509.load_pem_x509_certificate(Path(cert_path).read_bytes())
    return cert.public_key()


def verify_wechat_signature(body: str, headers: Dict[str, str]) -> bool:
    signature = headers.get("wechatpay-signature") or headers.get("Wechatpay-Signature") or ""
    timestamp = headers.get("wechatpay-timestamp") or headers.get("Wechatpay-Timestamp") or ""
    nonce = headers.get("wechatpay-nonce") or headers.get("Wechatpay-Nonce") or ""
    if not signature or not timestamp or not nonce:
        return False
    message = f"{timestamp}\n{nonce}\n{body}\n"
    public_key = load_wechat_platform_public_key()
    public_key.verify(base64.b64decode(signature), message.encode("utf-8"), padding.PKCS1v15(), hashes.SHA256())
    return True


def decrypt_wechat_resource(resource: Dict[str, Any]) -> Dict[str, Any]:
    api_v3_key = require_env("WECHAT_API_V3_KEY").encode("utf-8")
    aesgcm = AESGCM(api_v3_key)
    ciphertext = base64.b64decode(resource["ciphertext"])
    nonce = resource["nonce"].encode("utf-8")
    associated_data = resource.get("associated_data", "").encode("utf-8")
    plaintext = aesgcm.decrypt(nonce, ciphertext, associated_data)
    return json.loads(plaintext.decode("utf-8"))


def handle_wechat_notify(body: str, headers: Dict[str, str]) -> bool:
    try:
        if not verify_wechat_signature(body, headers):
            return False
        payload = json.loads(body)
        resource = decrypt_wechat_resource(payload.get("resource", {}))
        if resource.get("trade_state") != "SUCCESS":
            return False
        return mark_paid(resource.get("out_trade_no", ""), resource.get("transaction_id", ""))
    except Exception:
        return False


def create_payment_order(job_id: str, provider: str) -> Dict[str, Any]:
    if is_job_paid(job_id):
        return {"job_id": job_id, "paid": True, "amount_cents": product_amount_cents(), "amount_yuan": product_amount_yuan()}
    provider = provider.lower()
    if provider == "alipay":
        return create_alipay_page_payment(job_id)
    if provider == "wechat":
        return create_wechat_native_payment(job_id)
    raise RuntimeError("不支持的支付方式")
