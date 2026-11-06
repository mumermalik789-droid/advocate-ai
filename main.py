"""Pakistani Legal AI - REST backend (FastAPI). Secrets sirf environment variables mein."""
import hashlib, hmac, json, os, re, secrets, smtplib, ssl, time
from datetime import date, datetime, timedelta
from email.message import EmailMessage
from io import BytesIO

import anthropic, docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
from fastapi import FastAPI, Header, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import Response
from pydantic import BaseModel

API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "")      # code mein kabhi na likhein
SMTP_HOST, SMTP_PORT = "smtp.gmail.com", 587
SMTP_USER, SMTP_PASS = os.environ.get("SMTP_USER", ""), os.environ.get("SMTP_PASS", "")
DEV_SHOW_OTP = os.environ.get("DEV_SHOW_OTP") == "1"
ADMIN_NAME = "Muhammad Umer"

DATA = os.environ.get("DATA_DIR", "data")
os.makedirs(DATA, exist_ok=True)
USERS_FILE, PENDING_FILE = f"{DATA}/users.json", f"{DATA}/pending.json"
OTP_VALID, RESEND, SESSION_DAYS = 600, 60, 30
DAILY_LIMITS = {"basic": 15, "standard": 25, "premium": None}
MAX_DEVICES = {"basic": 2, "standard": 2, "premium": 3}
EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

app = FastAPI(title="Advocate AI API")
# Chrome (Flutter web) mein test ke liye. Asli server par allow_origins mein apni website likhein.
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

# ---------------------------------------------------------------- helpers
def rj(path, default):
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default

def wj(path, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def _h(t): return hashlib.sha256(t.encode()).hexdigest()

def canonical_email(e):
    e = e.strip().lower()
    local, _, dom = e.partition("@")
    if dom in ("gmail.com", "googlemail.com"):
        local, dom = local.split("+")[0].replace(".", ""), "gmail.com"
    return f"{local}@{dom}"

def hash_password(pw, salt=None):
    salt = salt or secrets.token_hex(16)
    return f"{salt}${hashlib.pbkdf2_hmac('sha256', pw.encode(), bytes.fromhex(salt), 200_000).hex()}"

def check_password(pw, stored):
    if not stored: return False
    if "$" in stored:
        return hmac.compare_digest(hash_password(pw, stored.split('$', 1)[0]), stored)
    return hmac.compare_digest(hashlib.sha256(pw.encode()).hexdigest(), stored)

def load_users():
    d = rj(USERS_FILE, {})
    old = d.get("admin", {})
    d["admin"] = {"chamber_name": f"Administrator ({ADMIN_NAME})", "password": "", "package": "premium",
                  "is_admin": True, "expiry_date": "2099-12-31", "prompt_count": old.get("prompt_count", 0),
                  "last_reset_date": old.get("last_reset_date", str(date.today())),
                  "sessions": old.get("sessions", {})}
    return d

def find_account(users, ident):
    ident = ident.strip().lower()
    if not ident: return None
    canon = canonical_email(ident) if EMAIL_RE.match(ident) else None
    for uid, u in users.items():
        if uid.lower() == ident or u.get("email", "").lower() == ident or (canon and u.get("email_canon") == canon):
            return uid
    return None

def new_session(users, uid):
    u = users[uid]; s = u.setdefault("sessions", {}); now = time.time()
    for h, ts in list(s.items()):
        if now - ts > SESSION_DAYS * 86400: del s[h]
    if not u.get("is_admin"):
        while len(s) >= MAX_DEVICES.get(u.get("package"), 2):
            del s[min(s, key=s.get)]
    tok = secrets.token_urlsafe(24); s[_h(tok)] = now
    wj(USERS_FILE, users)
    return tok

def profile(u):
    lim = DAILY_LIMITS.get(u.get("package"))
    return {"chamber_name": u["chamber_name"], "email": u.get("email", ""), "package": u.get("package"),
            "expiry_date": u.get("expiry_date"), "used": u.get("prompt_count", 0), "limit": lim,
            "is_admin": bool(u.get("is_admin"))}

# ---------------------------------------------------------------- auth dependency
def current(authorization, allow_expired=False):
    if not authorization: raise HTTPException(401, "Login zaroori hai")
    h = _h(authorization.replace("Bearer ", "").strip())
    users = load_users()
    for uid, u in users.items():
        ts = u.get("sessions", {}).get(h)
        if ts and time.time() - ts <= SESSION_DAYS * 86400:
            if u.get("last_reset_date") != str(date.today()):
                u["prompt_count"], u["last_reset_date"] = 0, str(date.today()); wj(USERS_FILE, users)
            if not allow_expired and not u.get("is_admin") and str(date.today()) > u.get("expiry_date", "0"):
                raise HTTPException(402, "Subscription expire ho gayi hai")
            return users, uid
    raise HTTPException(401, "Session khatam, dobara login karein")

# ---------------------------------------------------------------- email OTP
def send_email(to, subject, body):
    if not (SMTP_USER and SMTP_PASS): return False, "Email service configure nahi hai."
    m = EmailMessage(); m["From"], m["To"], m["Subject"] = SMTP_USER, to, subject; m.set_content(body)
    try:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=20) as s:
            s.starttls(context=ssl.create_default_context()); s.login(SMTP_USER, SMTP_PASS); s.send_message(m)
        return True, ""
    except Exception as e:
        return False, f"Email send nahi ho saka: {e}"

def otp_send(purpose, email, payload=None):
    key = f"{purpose}:{canonical_email(email)}"; now = time.time()
    pend = {k: v for k, v in rj(PENDING_FILE, {}).items() if v.get("expires", 0) > now}
    old = pend.get(key)
    if old and now - old["sent_at"] < RESEND:
        return False, f"{int(RESEND - (now - old['sent_at']))} second baad dobara try karein."
    code = f"{secrets.randbelow(10**6):06d}"
    ok, err = send_email(email, "Your verification code - Advocate AI",
                         f"Aapka verification code: {code}\nYe {OTP_VALID // 60} minute tak valid hai.")
    note = ""
    if not ok:
        if DEV_SHOW_OTP and not (SMTP_USER and SMTP_PASS): note = f"TEST MODE - code: {code}"
        else: return False, err
    pend[key] = {"code": _h(code), "expires": now + OTP_VALID, "attempts": 0, "sent_at": now, "payload": payload or {}}
    wj(PENDING_FILE, pend)
    return True, note

def otp_check(purpose, email, code):
    key = f"{purpose}:{canonical_email(email)}"; pend = rj(PENDING_FILE, {}); rec = pend.get(key)
    if not rec or rec["expires"] < time.time(): return False, "Code expire ho gaya.", {}
    if rec["attempts"] >= 5: return False, "Bohat ghalat koshishein. Naya code mangwayein.", {}
    if not hmac.compare_digest(rec["code"], _h(code.strip())):
        rec["attempts"] += 1; wj(PENDING_FILE, pend); return False, "Code ghalat hai.", {}
    payload = rec["payload"]; del pend[key]; wj(PENDING_FILE, pend)
    return True, "", payload

# ---------------------------------------------------------------- auth routes
class Login(BaseModel): identifier: str; password: str
class Signup(BaseModel): chamber: str; email: str; password: str; package: str = "basic"
class Verify(BaseModel): email: str; code: str
class Forgot(BaseModel): email: str
class Reset(BaseModel): email: str; code: str; new_password: str

@app.post("/auth/login")
def login(r: Login):
    users = load_users(); uid = find_account(users, r.identifier)
    if not uid: raise HTTPException(401, "Email ya password ghalat hai.")
    u = users[uid]
    ok = (bool(ADMIN_PASSWORD) and hmac.compare_digest(r.password, ADMIN_PASSWORD)) if u.get("is_admin") \
        else check_password(r.password, u.get("password", ""))
    if not ok: raise HTTPException(401, "Email ya password ghalat hai.")
    if not u.get("is_admin") and "$" not in u["password"]: u["password"] = hash_password(r.password)
    return {"token": new_session(users, uid), "user": profile(u)}

@app.post("/auth/signup")
def signup(r: Signup):
    email = r.email.strip().lower(); users = load_users()
    if not (r.chamber.strip() and EMAIL_RE.match(email)): raise HTTPException(400, "Naam aur sahi email likhein.")
    if len(r.password) < 8: raise HTTPException(400, "Password kam az kam 8 characters ka ho.")
    if r.package not in DAILY_LIMITS: raise HTTPException(400, "Ghalat package.")
    if find_account(users, email): raise HTTPException(400, "Is email se account pehle se hai.")
    ok, info = otp_send("signup", email, {"chamber": r.chamber.strip(), "pw_hash": hash_password(r.password), "pkg": r.package})
    if not ok: raise HTTPException(400, info)
    return {"ok": True, "note": info}

@app.post("/auth/verify")
def verify(r: Verify):
    email = r.email.strip().lower(); users = load_users()
    ok, msg, d = otp_check("signup", email, r.code)
    if not ok: raise HTTPException(400, msg)
    if find_account(users, email): raise HTTPException(400, "Account pehle se ban chuka hai.")
    users[email] = {"chamber_name": d["chamber"], "email": email, "email_canon": canonical_email(email),
                    "email_verified": True, "password": d["pw_hash"], "package": d["pkg"], "prompt_count": 0,
                    "last_reset_date": str(date.today()), "expiry_date": str(date.today() + timedelta(days=3)),
                    "is_admin": False, "sessions": {}}
    return {"token": new_session(users, email), "user": profile(users[email])}

@app.post("/auth/forgot")
def forgot(r: Forgot):
    users = load_users(); uid = find_account(users, r.email)
    note = ""
    if uid and users[uid].get("email"):
        ok, info = otp_send("reset", users[uid]["email"])
        if not ok: raise HTTPException(400, info)
        note = info
    return {"ok": True, "note": note}      # account hai ya nahi, jawab same

@app.post("/auth/reset")
def reset(r: Reset):
    if len(r.new_password) < 8: raise HTTPException(400, "Password kam az kam 8 characters ka ho.")
    users = load_users(); uid = find_account(users, r.email)
    ok, msg, _ = otp_check("reset", users[uid]["email"] if uid else r.email, r.code)
    if not (ok and uid): raise HTTPException(400, msg or "Reset nahi ho saka.")
    users[uid]["password"], users[uid]["sessions"] = hash_password(r.new_password), {}
    wj(USERS_FILE, users)
    return {"ok": True}

@app.post("/auth/logout")
def logout(authorization: str = Header(None)):
    users, uid = current(authorization, True)
    users[uid]["sessions"].pop(_h(authorization.replace("Bearer ", "").strip()), None); wj(USERS_FILE, users)
    return {"ok": True}

@app.get("/me")
def me(authorization: str = Header(None)):
    users, uid = current(authorization, True)
    return {**profile(users[uid]), "expired": (not users[uid].get("is_admin") and str(date.today()) > users[uid]["expiry_date"])}

# ---------------------------------------------------------------- user data (clients / invoices / history)
COLS = {"clients", "invoices", "history"}

def dpath(uid, col): return f"{DATA}/{re.sub(r'[^a-z0-9]+', '_', uid.lower())}_{col}.json"

@app.get("/data/{col}")
def get_col(col: str, authorization: str = Header(None)):
    _, uid = current(authorization)
    if col not in COLS: raise HTTPException(404)
    return rj(dpath(uid, col), [])

@app.put("/data/{col}")
def put_col(col: str, rows: list[dict], authorization: str = Header(None)):
    _, uid = current(authorization)
    if col not in COLS: raise HTTPException(404)
    wj(dpath(uid, col), rows[:5000]); return {"ok": True}

# ---------------------------------------------------------------- AI
SYSTEM = ("You are an expert Senior Advocate and Legal Drafter practicing in the High Courts and Supreme Court "
          "of Pakistan. You excel at Pakistani law (Constitution, CPC, PPC, CrPC, QSO, Revenue Laws), precedents, "
          "and drafting formal legal documents. If you are not certain a citation is real, say so clearly instead of inventing one.")

def lang_instr(l):
    if l.startswith("Urdu"):
        return ("IMPORTANT: Write the ENTIRE output strictly in High Court & District Court standard Pakistani legal "
                "Urdu (بعدالت جناب, سائل/درخواست گزار, مسئول عنہ, عرض ہے کہ, استدعا ہے کہ, حلف نامہ).")
    if l.startswith("Roman"): return "IMPORTANT: Write the ENTIRE output strictly in Roman Urdu (Urdu in English script)."
    return "IMPORTANT: Write the ENTIRE output strictly in Standard Legal English as used in Pakistani courts."

class AiReq(BaseModel):
    kind: str                      # draft | research | citation | fir
    language: str = "English"
    fields: dict = {}
    image_b64: str | None = None   # JPEG

def build_prompt(r: AiReq):
    f, L = r.fields, lang_instr(r.language); blank = "(not provided - leave a clear blank placeholder)"
    g = lambda k: (f.get(k) or "").strip()
    if r.kind == "draft":
        if not g("facts"): raise HTTPException(400, "Pehle facts likhein.")
        return f"""{L}
Follow standard Pakistani court drafting style and terminology.
Prepare a complete, court-ready {g('category')} document: {g('doc_type')}
- Court/Forum: {g('court') or blank}
- Parties: {g('parties') or blank}
- Reference (FIR/Law): {g('ref') or blank}
- Other details: {g('extra') or blank}
- Facts & Grounds: {g('facts')}""", f"Draft - {g('doc_type')}"
    if r.kind == "research":
        if not (g("section") or g("question")): raise HTTPException(400, "Section ya sawal likhein.")
        return f"""{L}
Conduct legal research on Pakistani law: {g('act')}
- Provision: {g('section')}
- Precedents from: {g('court')}
- Question/Facts: {g('question')}
Give the bare-act position, key principles and leading case law. Mark any citation you are unsure of.""", f"Research - {g('act')}"
    if r.kind == "citation":
        if not g("citation"): raise HTTPException(400, "Citation likhein.")
        return f"""{L}
Review this Pakistani citation: {g('citation')} ({g('court')}). Topic: {g('topic')}
Cover: 1) what you can confirm about parties/year/reporter, 2) parallel citations,
3) whether it is still good law or overruled, 4) ratio decidendi.
Be explicit about anything you cannot confirm; do NOT invent details.""", "Citation Check"
    if r.kind == "fir":
        if not (g("text") or r.image_b64): raise HTTPException(400, "FIR text ya photo dein.")
        return f"""{L}
Summarize and analyze this FIR for a Pakistani advocate: allegations, sections applied,
weak points, bail grounds and next steps.
FIR TEXT:
{g('text')}""", "FIR Analysis"
    raise HTTPException(400, "Unknown kind")

@app.post("/ai")
def ai(r: AiReq, authorization: str = Header(None)):
    users, uid = current(authorization); u = users[uid]
    lim = DAILY_LIMITS.get(u.get("package"))
    if not u.get("is_admin") and lim is not None and u.get("prompt_count", 0) >= lim:
        raise HTTPException(429, f"Aaj ki limit ({lim} prompts) mukammal. Kal reset hogi ya package upgrade karein.")
    if not API_KEY: raise HTTPException(500, "AI service configure nahi hai.")
    prompt, label = build_prompt(r)
    content = prompt
    if r.image_b64:
        if len(r.image_b64) > 2_000_000: raise HTTPException(400, "Image bohat bari hai.")
        content = [{"type": "image", "source": {"type": "base64", "media_type": "image/jpeg", "data": r.image_b64}},
                   {"type": "text", "text": prompt}]
    try:
        res = anthropic.Anthropic(api_key=API_KEY).messages.create(
            model="claude-sonnet-4-6", max_tokens=3500, system=SYSTEM, messages=[{"role": "user", "content": content}])
    except Exception as e:
        raise HTTPException(502, f"AI se jawab nahi aa saka: {e}")
    text = res.content[0].text
    u["prompt_count"] = u.get("prompt_count", 0) + 1; wj(USERS_FILE, users)
    hist = rj(dpath(uid, "history"), [])
    hist.append({"time": datetime.now().strftime("%d-%m-%Y %H:%M"), "type": label, "result": text, "lang": r.language})
    wj(dpath(uid, "history"), hist[-300:])
    return {"result": text, "type": label, "used": u["prompt_count"], "limit": lim}

# ---------------------------------------------------------------- DOCX export
class Export(BaseModel): text: str; urdu: bool = False

@app.post("/export/docx")
def export_docx(r: Export, authorization: str = Header(None)):
    current(authorization)
    d = docx.Document()
    for line in r.text.split("\n"):
        line = line.strip()
        if not line: continue
        p = d.add_paragraph()
        if r.urdu: p._p.get_or_add_pPr().insert(0, OxmlElement("w:bidi"))
        for part in re.split(r"(\*\*.*?\*\*)", line):
            if not part: continue
            run = p.add_run()
            if part.startswith("**") and part.endswith("**"): run.text, run.bold = part[2:-2], True
            else: run.text = part.replace("###", "").replace("##", "").replace("#", "")
            if r.urdu:
                run.font.name, run.font.size = "Jameel Noori Nastaliq", Pt(14)
                rpr = run._r.get_or_add_rPr(); rpr.find(qn("w:rFonts")).set(qn("w:cs"), "Jameel Noori Nastaliq")
                rpr.append(OxmlElement("w:rtl"))
            else: run.font.name, run.font.size = "Calibri", Pt(12)
    bio = BytesIO(); d.save(bio)
    return Response(bio.getvalue(), media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

# ---------------------------------------------------------------- admin
@app.get("/admin/users")
def admin_users(authorization: str = Header(None)):
    users, uid = current(authorization)
    if not users[uid].get("is_admin"): raise HTTPException(403)
    return [{"id": k, **profile(v), "devices": len(v.get("sessions", {}))} for k, v in users.items() if k != "admin"]

class AdminSet(BaseModel): id: str; package: str | None = None; expiry_date: str | None = None; add_days: int = 0

@app.post("/admin/update")
def admin_update(r: AdminSet, authorization: str = Header(None)):
    users, uid = current(authorization)
    if not users[uid].get("is_admin"): raise HTTPException(403)
    v = users.get(r.id)
    if not v: raise HTTPException(404)
    if r.package in DAILY_LIMITS: v["package"] = r.package
    if r.expiry_date: v["expiry_date"] = r.expiry_date
    if r.add_days:
        base = max(datetime.strptime(v["expiry_date"], "%Y-%m-%d").date(), date.today())
        v["expiry_date"] = str(base + timedelta(days=r.add_days))
    wj(USERS_FILE, users); return {"ok": True}
