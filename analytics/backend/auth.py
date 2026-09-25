import os
import base64
import hashlib
import secrets
from flask import request, Response

def _md5_apr1_check(password: str, full_hash: str) -> bool:
    """Verify Apache MD5 ($apr1$) password hash in pure Python without extra dependencies."""
    parts = full_hash.split("$")
    if len(parts) < 4 or parts[1] != "apr1":
        return False
    salt = parts[2]
    expected_hash = parts[3]

    magic = "$apr1$"
    ctx = hashlib.md5((password + magic + salt).encode("latin1"))
    ctx_alt = hashlib.md5((password + salt + password).encode("latin1"))
    alt_result = ctx_alt.digest()

    plen = len(password)
    i = plen
    while i > 16:
        ctx.update(alt_result)
        i -= 16
    ctx.update(alt_result[:i])

    i = plen
    while i > 0:
        if i & 1:
            ctx.update(b"\x00")
        else:
            ctx.update(password[:1].encode("latin1"))
        i >>= 1

    final = ctx.digest()

    for r in range(1000):
        ctx1 = hashlib.md5()
        if r & 1:
            ctx1.update(password.encode("latin1"))
        else:
            ctx1.update(final)
        if r % 3:
            ctx1.update(salt.encode("latin1"))
        if r % 7:
            ctx1.update(password.encode("latin1"))
        if r & 1:
            ctx1.update(final)
        else:
            ctx1.update(password.encode("latin1"))
        final = ctx1.digest()

    b64 = "./0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"

    def to64(v, n):
        res = ""
        while n > 0:
            res += b64[v & 0x3F]
            v >>= 6
            n -= 1
        return res

    reordered = (
        to64((final[0] << 16) | (final[6] << 8) | final[12], 4)
        + to64((final[1] << 16) | (final[7] << 8) | final[13], 4)
        + to64((final[2] << 16) | (final[8] << 8) | final[14], 4)
        + to64((final[3] << 16) | (final[9] << 8) | final[15], 4)
        + to64((final[4] << 16) | (final[10] << 8) | final[5], 4)
        + to64(final[11], 2)
    )
    return secrets.compare_digest(reordered, expected_hash)

def _sha1_check(password: str, hash_val: str) -> bool:
    """Verify Apache {SHA} hash format."""
    expected = "{SHA}" + base64.b64encode(hashlib.sha1(password.encode("utf-8")).digest()).decode("ascii")
    return secrets.compare_digest(hash_val, expected)

def _load_htpasswd_file() -> dict:
    """Find and parse .htpasswd file if it exists."""
    search_paths = [
        os.environ.get("HTPASSWD_PATH", ""),
        os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".htpasswd")),
        os.path.abspath(os.path.join(os.path.dirname(__file__), ".htpasswd")),
        os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "frontend", ".htpasswd")),
        "/app/.htpasswd",
    ]

    users = {}
    for p in search_paths:
        if p and os.path.isfile(p):
            try:
                with open(p, "r", encoding="utf-8") as f:
                    for line in f:
                        line = line.strip()
                        if line and not line.startswith("#") and ":" in line:
                            u, h = line.split(":", 1)
                            users[u.strip()] = h.strip()
                if users:
                    break
            except Exception:
                continue
    return users

def verify_password(stored_hash_or_pass: str, provided_password: str) -> bool:
    """Verify a provided password against a hash or plain text."""
    if not stored_hash_or_pass or not provided_password:
        return False

    # APR1 MD5 Apache format
    if stored_hash_or_pass.startswith("$apr1$"):
        return _md5_apr1_check(provided_password, stored_hash_or_pass)

    # SHA-1 Apache format
    if stored_hash_or_pass.startswith("{SHA}"):
        return _sha1_check(provided_password, stored_hash_or_pass)

    # Werkzeug hashes
    if stored_hash_or_pass.startswith(("scrypt:", "pbkdf2:")):
        try:
            from werkzeug.security import check_password_hash
            return check_password_hash(stored_hash_or_pass, provided_password)
        except Exception:
            pass

    # Plain text comparison
    return secrets.compare_digest(stored_hash_or_pass, provided_password)

def check_auth(username: str, password: str) -> bool:
    """Authenticate username and password against .htpasswd or environment variables."""
    # 1. Check against .htpasswd if available
    htpasswd_users = _load_htpasswd_file()
    if username in htpasswd_users:
        if verify_password(htpasswd_users[username], password):
            return True

    # 2. Check against environment variables
    env_user = (
        os.environ.get("ANALYTICS_USER")
        or os.environ.get("HTACCESS_USER")
        or os.environ.get("AUTH_USER")
        or "admin"
    )
    env_pass = (
        os.environ.get("ANALYTICS_PASSWORD")
        or os.environ.get("HTACCESS_PASSWORD")
        or os.environ.get("AUTH_PASSWORD")
        or os.environ.get("DASHBOARD_PASSWORD")
        or "admin123"
    )

    if secrets.compare_digest(username, env_user) and secrets.compare_digest(password, env_pass):
        return True

    return False

def get_auth_credentials():
    """Extract username and password from request."""
    # Standard Flask / Werkzeug authorization
    auth = request.authorization
    if auth and auth.username is not None and auth.password is not None:
        return auth.username, auth.password

    # Fallback: manual parsing of Authorization header
    auth_header = request.headers.get("Authorization", "")
    if auth_header.startswith("Basic "):
        try:
            encoded = auth_header.split(" ", 1)[1].strip()
            decoded = base64.b64decode(encoded).decode("utf-8")
            if ":" in decoded:
                u, p = decoded.split(":", 1)
                return u, p
        except Exception:
            pass

    return None, None

def init_auth(app):
    """Register HTTP Basic Auth middleware on the Flask application."""
    @app.before_request
    def require_basic_auth():
        # Allow CORS preflight requests
        if request.method == "OPTIONS":
            return None

        # Check if auth is explicitly disabled (e.g. for local offline dev)
        auth_enabled = os.environ.get("ANALYTICS_AUTH_ENABLED", "true").lower()
        if auth_enabled in ("false", "0", "no", "off"):
            return None

        username, password = get_auth_credentials()
        if not username or not password or not check_auth(username, password):
            return Response(
                "Access Denied: Authentication Required.\n",
                401,
                {
                    "WWW-Authenticate": 'Basic realm="Hello Park Analytics Exporter"',
                    "Content-Type": "text/plain; charset=utf-8",
                },
            )
        return None
