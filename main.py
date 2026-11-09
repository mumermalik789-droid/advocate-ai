"""Pakistani Legal AI - REST backend (FastAPI). Secrets sirf environment variables mein."""
import hashlib, hmac, json, os, re, secrets, smtplib, ssl, time, traceback, urllib.error, urllib.request
from datetime import date, datetime, timedelta
from email.message import EmailMessage
from io import BytesIO

import anthropic, docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
from fastapi import FastAPI, Header, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse, Response
from pydantic import BaseModel

API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "")      # code mein kabhi na likhein
SMTP_HOST, SMTP_PORT = "smtp.gmail.com", 587
SMTP_USER, SMTP_PASS = os.environ.get("SMTP_USER", ""), os.environ.get("SMTP_PASS", "")
DEV_SHOW_OTP = os.environ.get("DEV_SHOW_OTP") == "1"
BREVO_API_KEY = os.environ.get("BREVO_API_KEY", "")          # HTTP email (free plan par bhi chalta hai)
MAIL_FROM = os.environ.get("MAIL_FROM", SMTP_USER)            # Brevo mein verify kiya hua sender email
EMAIL_CONFIGURED = bool(BREVO_API_KEY or (SMTP_USER and SMTP_PASS))
ADMIN_NAME = "Muhammad Umer"
MODEL = os.environ.get("CLAUDE_MODEL", "claude-sonnet-5-5")

DATA = os.environ.get("DATA_DIR", "data")
os.makedirs(DATA, exist_ok=True)
USERS_FILE, PENDING_FILE = f"{DATA}/users.json", f"{DATA}/pending.json"
CAT_CACHE = f"{DATA}/category_cache.json"
OTP_VALID, RESEND, SESSION_DAYS = 600, 60, 30
DAILY_LIMITS = {"basic": 15, "standard": 25, "premium": None}
MAX_DEVICES = {"basic": 2, "standard": 2, "premium": 3}
EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

app = FastAPI(title="Advocate AI API")
# Chrome (Flutter web) mein test ke liye. Asli server par allow_origins mein apni website likhein.
@app.exception_handler(Exception)
async def _unhandled(request: Request, exc: Exception):
    traceback.print_exception(type(exc), exc, exc.__traceback__)      # Render ke Logs mein poori wajah
    return JSONResponse(status_code=500, content={"detail": f"Server error ({type(exc).__name__}): {str(exc)[:160]}"})

app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

@app.get("/")
def root():
    return {"status": "ok", "app": "Advocate AI API", "version": "v15-fix500", "model": MODEL}

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
    if BREVO_API_KEY:
        if not MAIL_FROM: return False, "MAIL_FROM set nahi hai."
        payload = {"sender": {"name": "Advocate AI", "email": MAIL_FROM}, "to": [{"email": to}],
                   "subject": subject, "textContent": body}
        req = urllib.request.Request("https://api.brevo.com/v3/smtp/email", data=json.dumps(payload).encode(),
                                     headers={"api-key": BREVO_API_KEY, "content-type": "application/json",
                                              "accept": "application/json"}, method="POST")
        try:
            with urllib.request.urlopen(req, timeout=20):
                return True, ""
        except urllib.error.HTTPError as e:
            return False, f"Email send nahi ho saka (Brevo {e.code}): {e.read().decode('utf-8', 'ignore')[:150]}"
        except Exception as e:
            return False, f"Email send nahi ho saka: {e}"
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
        if DEV_SHOW_OTP and not EMAIL_CONFIGURED: note = f"TEST MODE - code: {code}"
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
                    "last_reset_date": str(date.today()), "created": str(date.today()), "expiry_date": str(date.today() + timedelta(days=3)),
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
SYSTEM = ("You are a Senior Advocate of the High Court and Supreme Court of Pakistan with 30 years of drafting experience. "
          "You excel at Pakistani law (Constitution, CPC, PPC, CrPC, QSO, Family, Revenue and special laws) and precedents. "
          "When asked to draft, produce ONLY the finished court-ready document: proper heading and title of the case, parties, "
          "numbered paragraphs, grounds, prayer, verification/affidavit and signature block where applicable. Use a formal, "
          "precise, professional pleading register at all times, never a chatty or simplified tone. Do not add introductions, "
          "notes or explanations outside the document and do not use markdown symbols (#, *, tables). Put unknown facts in "
          "square-bracket blanks such as [____]. If you are not certain a citation is real, say so clearly instead of inventing one.")

URDU_STYLE = (
    "IMPORTANT: Write the ENTIRE document in formal, professional Pakistani court Urdu, exactly as a senior advocate drafts "
    "pleadings for the High Court and District Courts. Use Urdu script only (no Roman Urdu); English only for unavoidable names, "
    "section numbers and citations. Use the authentic pleading register and structure: court heading 'بعدالت جناب ...', case title "
    "block with 'درخواست گزار/سائل ... بنام ... مسئول علیہ/مدعا علیہ' (or 'مستغیث', 'ملزم', 'استغاثہ' as the case requires), a "
    "subject line 'عنوان: ...', an opening 'مؤدبانہ گزارش ہے کہ', numbered paragraphs each beginning 'یہ کہ', grounds as numbered "
    "'بنیادِ درخواست', the prayer introduced by 'لہٰذا استدعا ہے کہ ...', then 'مورخہ', 'العارض / درخواست گزار' with 'بذریعہ ایڈووکیٹ' "
    "signature block, and 'حلف نامہ' / 'تصدیق' in the standard form ('میں ... بحلف بیان کرتا/کرتی ہوں کہ ...') wherever applicable. "
    "Cite law in standard Urdu court form (for example 'زیر دفعہ 497 مجموعہ ضابطہ فوجداری 1898', 'مجموعہ تعزیراتِ پاکستان', 'ایف آئی آر', "
    "'تھانہ'). Avoid colloquial, Hindi-origin or casual words; use complete, dignified, grammatically correct legal sentences.")

def lang_instr(l):
    if l.startswith("Urdu"): return URDU_STYLE
    if l.startswith("Roman"): return "IMPORTANT: Write the ENTIRE output strictly in formal, professional Roman Urdu (Urdu in English script), keeping the structure of a Pakistani court pleading."
    return ("IMPORTANT: Write the ENTIRE output strictly in formal Standard Legal English as used in Pakistani courts: pleading style "
            "('It is respectfully submitted that...', numbered paragraphs beginning 'That ...', 'GROUNDS', 'PRAYER', 'VERIFICATION').")

CATEGORIES = {
    "qatal": "Murder and homicide (qatl-i-amd, qatl-i-shibh-i-amd, qatl-i-khata, qatl-bis-sabab, attempt to murder, abetment and related provisions)",
    "larai": "Fights, hurt and assault (every category of hurt, shajjah, itlaf-i-udw, grievous hurt, assault, hurt by weapons)",
    "fasad": "Rioting, unlawful assembly and affray",
    "chori": "Theft, house-trespass, house-breaking and receiving stolen property",
    "dakaiti": "Robbery, dacoity and extortion",
    "fraud": "Cheating, forgery, criminal breach of trust, misappropriation and fraud",
    "cheque": "Dishonour of cheques and related negotiable-instrument offences",
    "aghwa": "Kidnapping, abduction, wrongful confinement and trafficking",
    "zina": "Zina, rape and other sexual offences",
    "dhamki": "Criminal intimidation, threats, blackmail and harassment",
    "qabza": "Illegal possession, trespass, mischief and land-grabbing (criminal, civil and revenue provisions)",
    "narcotics": "Narcotics offences under the Control of Narcotic Substances Act 1997",
    "cyber": "Cyber crimes under the Prevention of Electronic Crimes Act 2016 and related provisions",
    "arms": "Arms and explosives offences",
    "tohmat": "Defamation, insult and offences against honour",
    "hadsa": "Rash and negligent driving, negligence causing death or hurt",
    "jhoot": "False evidence, false case or FIR, perjury and offences against public justice",
    "daheshat": "Terrorism offences under the Anti-Terrorism Act 1997",
    "ghareloo": "Domestic and family offences (dowry, domestic violence, cruelty, marriage-related offences)",
}

def lang_plain(l):
    if l.startswith("Urdu"): return "IMPORTANT: Write the answer in clear, professional legal Urdu (Urdu script); keep section numbers, statute names and case citations in English. Do not use pleading format."
    if l.startswith("Roman"): return "IMPORTANT: Write the answer in clear, professional Roman Urdu (Urdu in English script); keep statute names and citations in English."
    return "IMPORTANT: Write the answer in clear, professional Standard Legal English as used by Pakistani advocates."

TEMPLATE_INSTR = ("IMPORTANT: The advocate's own sample draft is attached as TEMPLATE images. Reproduce its layout exactly: the same "
                  "heading and title block, order of sections, numbering style, opening and closing formulas, prayer and verification wording, "
                  "and the same language, script and legal register as the template. Replace ONLY the case-specific details with the facts "
                  "given below and do not copy the template's own names, dates or numbers.")

class AiReq(BaseModel):
    kind: str                      # draft | research | citation | fir
    language: str = "English"
    fields: dict = {}
    image_b64: str | None = None   # purana (ek tasveer)
    images_b64: list[str] | None = None   # nayi: kai tasveerein (JPEG)
    template_id: str | None = None        # saved template (server par)

def build_prompt(r: AiReq, has_template=False):
    f, L = r.fields, (TEMPLATE_INSTR if has_template else lang_instr(r.language)); blank = "(not provided - leave a clear blank placeholder)"; LP = lang_plain(r.language)
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
- Facts & Grounds: {g('facts')}
Output only the finished document (no preface, notes or markdown symbols).""", f"Draft - {g('doc_type')}"
    if r.kind == "research":
        if not (g("section") or g("question")): raise HTTPException(400, "Section ya sawal likhein.")
        return f"""{LP}
Conduct legal research on Pakistani law: {g('act')}
- Provision: {g('section')}
- Precedents from: {g('court')}
- Question/Facts: {g('question')}
Give the bare-act position, key principles and leading case law. Mark any citation you are unsure of.""", f"Research - {g('act')}"
    if r.kind == "citation":
        if not g("citation"): raise HTTPException(400, "Citation likhein.")
        return f"""{LP}
Review this Pakistani citation: {g('citation')} ({g('court')}). Topic: {g('topic')}
Cover: 1) what you can confirm about parties/year/reporter, 2) parallel citations,
3) whether it is still good law or overruled, 4) ratio decidendi.
Be explicit about anything you cannot confirm; do NOT invent details.""", "Citation Check"
    if r.kind == "fir":
        if not (g("text") or r.image_b64 or r.images_b64): raise HTTPException(400, "FIR text ya photo dein.")
        return f"""{LP}
Summarize and analyze this FIR for a Pakistani advocate: allegations, sections applied,
weak points, bail grounds and next steps. Use a formal professional tone. End with one final line starting with 'Recommended draft:' naming the single most suitable petition or application to file now (for example pre-arrest bail, post-arrest bail, quashing of FIR).
FIR TEXT:
{g('text')}""", "FIR Analysis"
    if r.kind == "category":
        cid = g("category")
        if cid not in CATEGORIES: raise HTTPException(400, "Ghalat category.")
        return f"""{LP}
Topic: {CATEGORIES[cid]}
List the sections of Pakistani law that are commonly applied in cases of this kind, so that an advocate can see at a glance which sections apply. Group them as: A) Pakistan Penal Code sections, B) special-law sections (if any), C) Code of Criminal Procedure (or other procedural) sections commonly used in such cases (for example FIR, arrest, bail, quashing), D) sections that are commonly added together (such as 34, 109, 148/149, 120-B) when relevant.
For every section give one compact block in this form, without tables or markdown: 'Section <number> <Act> - <name of offence>: punishment; bailable or non-bailable; cognizable or non-cognizable; compoundable or not; court of trial'. Keep each block to one or two lines and order the sections logically (most serious first).
Be accurate. If you are not sure about a detail (for example a punishment after an amendment) write 'verify' instead of guessing, and never invent section numbers. Finish with one line reminding the advocate to verify the current text on the official Pakistan Code website.""", f"Dafaat - {cid}"
    if r.kind == "section":
        law = g("law") or "all relevant Pakistani laws"
        if g("mode") == "topic":
            if not g("topic"): raise HTTPException(400, "Jurm ya masla likhein.")
            return f"""{LP}
An advocate describes this offence or legal issue: {g('topic')}
Search within: {law}
Identify the most relevant sections or provisions of Pakistani law that apply: for each give the Act, the section number and a one-line explanation, ranked by relevance. For offences also state the punishment, whether it is bailable or non-bailable, cognizable or non-cognizable, compoundable, and the court of trial. Mention alternative or companion sections that are often invoked together, then add brief practical notes for the advocate.
Be explicit about anything you are not certain of and never invent section numbers.""", "Dafa search"
        if not g("section"): raise HTTPException(400, "Dafa / section number likhein.")
        q = ("Specific question: " + g("question")) if g("question") else ""
        return f"""{LP}
Explain this provision of Pakistani law for an advocate. Law: {law}. Provision: {g('section')}. {q}
Use these headings: 1) Name of the law and the section title; 2) Text or substance of the provision (if you cannot recall the exact wording, give the substance and say so); 3) Essential ingredients; 4) Punishment or relief, and for offences whether bailable or non-bailable, cognizable or non-cognizable, compoundable, and the court of trial; 5) Important amendments you know of; 6) Related and alternative sections; 7) Leading case law (mark any citation you are not sure of); 8) Practical points for bail, quashing or defence.
Statutes are amended often, so state clearly that the advocate must verify the current text on the official Pakistan Code website before relying on it. Never invent section numbers, text or citations.""", (f"Dafa - {g('section')} {g('law')}")[:80]
    raise HTTPException(400, "Unknown kind")

@app.post("/ai")
def ai(r: AiReq, authorization: str = Header(None)):
    users, uid = current(authorization); u = users[uid]
    lim = DAILY_LIMITS.get(u.get("package"))
    ckey = f"{(r.fields or {}).get('category', '')}|{r.language}"
    if r.kind == "category":      # ek hi list sab ke liye: cache mein ho to free aur foran
        hit = rj(CAT_CACHE, {}).get(ckey)
        if hit and time.time() - hit.get("ts", 0) < 30 * 86400:
            return {"result": hit["text"], "type": f"Dafaat - {ckey.split('|')[0]}", "used": u.get("prompt_count", 0), "limit": lim, "cached": True}
    if not u.get("is_admin") and lim is not None and u.get("prompt_count", 0) >= lim:
        raise HTTPException(429, f"Aaj ki limit ({lim} prompts) mukammal. Kal reset hogi ya package upgrade karein.")
    if not API_KEY: raise HTTPException(500, "AI service configure nahi hai.")
    tmpl = []
    if r.template_id:
        for t in rj(dpath(uid, "templates"), []):
            if t.get("id") == r.template_id: tmpl = t.get("images", [])
        if not tmpl: raise HTTPException(404, "Ye template nahi mila.")
    prompt, label = build_prompt(r, bool(tmpl))
    imgs = [i for i in (list(r.images_b64 or []) + ([r.image_b64] if r.image_b64 else [])) if i]
    if len(imgs) > 5: raise HTTPException(400, "Ek waqt mein ziyada se ziyada 5 tasveerein dein.")
    if any(len(i) > 2_000_000 for i in imgs): raise HTTPException(400, "Koi tasveer bohat bari hai.")
    img_block = lambda i: {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg", "data": i}}
    blocks = []
    if tmpl:
        blocks.append({"type": "text", "text": "TEMPLATE (advocate's own sample draft - copy its layout and language):"})
        blocks += [img_block(i) for i in tmpl]
    if imgs:
        blocks.append({"type": "text", "text": "CASE DOCUMENTS (FIR / order photos - use for facts only):"})
        blocks += [img_block(i) for i in imgs]
    content = (blocks + [{"type": "text", "text": prompt}]) if blocks else prompt
    client = anthropic.Anthropic(api_key=API_KEY)
    res, last_err = None, None
    for m in ([MODEL] + (["claude-sonnet-4-6"] if MODEL != "claude-sonnet-4-6" else [])):   # naya model na chale to purana
        try:
            res = client.messages.create(model=m, max_tokens=6000, system=SYSTEM, messages=[{"role": "user", "content": content}])
            break
        except Exception as e:
            last_err = e
    if res is None:
        raise HTTPException(502, f"AI se jawab nahi aa saka ({type(last_err).__name__}): {str(last_err)[:220]}")
    text = "".join(b.text for b in res.content if getattr(b, "type", "") == "text").strip()
    if not text:
        raise HTTPException(502, f"AI ne khali jawab diya (stop_reason: {getattr(res, 'stop_reason', '?')}). Dobara try karein.")
    if r.kind == "category":
        cache = rj(CAT_CACHE, {}); cache[ckey] = {"text": text, "ts": time.time()}; wj(CAT_CACHE, cache)
    u["prompt_count"] = u.get("prompt_count", 0) + 1; wj(USERS_FILE, users)
    hist = rj(dpath(uid, "history"), [])
    hist.append({"time": datetime.now().strftime("%d-%m-%Y %H:%M"), "type": label, "result": text, "lang": r.language})
    wj(dpath(uid, "history"), hist[-300:])
    return {"result": text, "type": label, "used": u["prompt_count"], "limit": lim}

# ---------------------------------------------------------------- templates (advocate ke apne draft ki photos)
class TemplateIn(BaseModel): name: str; images_b64: list[str]

@app.get("/templates")
def list_templates(authorization: str = Header(None)):
    _, uid = current(authorization)
    return [{"id": t["id"], "name": t["name"], "pages": len(t["images"])} for t in rj(dpath(uid, "templates"), [])]

@app.post("/templates")
def add_template(r: TemplateIn, authorization: str = Header(None)):
    _, uid = current(authorization)
    if not r.name.strip(): raise HTTPException(400, "Template ka naam likhein.")
    if not (1 <= len(r.images_b64) <= 3): raise HTTPException(400, "Template ki 1 se 3 tasveerein (pages) dein.")
    if any(len(i) > 2_000_000 for i in r.images_b64): raise HTTPException(400, "Koi tasveer bohat bari hai.")
    lst = rj(dpath(uid, "templates"), [])
    if len(lst) >= 10: raise HTTPException(400, "Ziyada se ziyada 10 templates save ho sakte hain. Purana delete karein.")
    t = {"id": secrets.token_hex(4), "name": r.name.strip()[:60], "images": r.images_b64}
    lst.append(t); wj(dpath(uid, "templates"), lst)
    return {"id": t["id"]}

@app.delete("/templates/{tid}")
def delete_template(tid: str, authorization: str = Header(None)):
    _, uid = current(authorization)
    wj(dpath(uid, "templates"), [t for t in rj(dpath(uid, "templates"), []) if t.get("id") != tid])
    return {"ok": True}

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
    return [{"id": k, **profile(v), "devices": len(v.get("sessions", {})), "email_verified": bool(v.get("email_verified")),
             "created": v.get("created", "")} for k, v in users.items() if k != "admin"]

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


class AdminId(BaseModel): id: str

@app.post("/admin/logout_all")
def admin_logout_all(r: AdminId, authorization: str = Header(None)):
    users, uid = current(authorization)
    if not users[uid].get("is_admin"): raise HTTPException(403)
    if r.id not in users: raise HTTPException(404)
    users[r.id]["sessions"] = {}; wj(USERS_FILE, users); return {"ok": True}

@app.post("/admin/delete")
def admin_delete(r: AdminId, authorization: str = Header(None)):
    users, uid = current(authorization)
    if not users[uid].get("is_admin"): raise HTTPException(403)
    if r.id == "admin" or r.id not in users: raise HTTPException(400, "Ye account delete nahi ho sakta.")
    del users[r.id]; wj(USERS_FILE, users)
    for col in list(COLS) + ["templates"]:
        try: os.remove(dpath(r.id, col))
        except OSError: pass
    return {"ok": True}
