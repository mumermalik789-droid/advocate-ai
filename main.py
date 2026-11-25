"""Pakistani Legal AI - REST backend (FastAPI). Secrets sirf environment variables mein."""
import hashlib, hmac, json, os, re, secrets, smtplib, ssl, threading, time, traceback, urllib.error, urllib.request
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

import types


def _embed(name, src):
    mod = types.ModuleType(name); exec(compile(src, name + '.py', 'exec'), mod.__dict__); return mod


cit = _embed('citations', r'''"""Verified-citation engine (stdlib only).

Rule: a case citation is shown as authentic ONLY if it exists in a verified index. No LLM is ever asked whether a
case is real. The index is filled from lawful sources: official court websites, notes you typed or exported from your
own Pakistan Law Site / East Law subscription, or entries an advocate marks as 'verified by me'."""
import csv, io, json, os, re, secrets, threading
from datetime import date

DATA = os.environ.get("DATA_DIR", "data")
GLOBAL_IDX = f"{DATA}/verified_citations.jsonl"
_LOCK, _CACHE = threading.Lock(), {"mtime": None, "rows": {}}
HOOK = {"rj": None, "wj": None}                    # main.py plugs its durable rj / wj in here
GLOBAL_JSON, REQ_JSON, _GV = f"{DATA}/verified_citations.json", f"{DATA}/citation_requests.json", {"ver": 0}
LIB_BYTES = 6_000_000                               # about 10,000 entries; Upstash free allows 10 MB per request


def _rjs(path, default): return HOOK["rj"](path, default) if HOOK["rj"] else _read_json(path, default)


def _wjs(path, data):
    if HOOK["wj"]: HOOK["wj"](path, data)
    else: os.makedirs(DATA, exist_ok=True); _write_json(path, data)

# First eight are the reporters Pakistan Law Site publishes. Other codes are recognised; name left blank if unsure.
REPORTERS = {"PLD": "Pakistan Legal Decisions", "SCMR": "Supreme Court Monthly Review", "CLC": "Civil Law Cases",
             "CLD": "Corporate Law Decisions", "YLR": "Yearly Law Reporter", "PCRLJ": "Pakistan Criminal Law Journal",
             "PTD": "Pakistan Tax Decisions", "PLC": "Pakistan Labour Cases",
             "MLD": "", "PLJ": "", "NLR": "", "GBLR": "", "SBLR": "", "PTCL": ""}
COURTS = {"SC": "Supreme Court of Pakistan", "FSC": "Federal Shariat Court", "LAHORE": "Lahore High Court",
          "LHR": "Lahore High Court", "LHC": "Lahore High Court", "KARACHI": "Sindh High Court", "KAR": "Sindh High Court",
          "SINDH": "Sindh High Court", "PESHAWAR": "Peshawar High Court", "PESH": "Peshawar High Court",
          "QUETTA": "Balochistan High Court", "BALOCHISTAN": "Balochistan High Court", "BALOCH": "Balochistan High Court",
          "ISLAMABAD": "Islamabad High Court", "ISL": "Islamabad High Court", "IHC": "Islamabad High Court",
          "AJ&K": "Azad Jammu & Kashmir", "AJK": "Azad Jammu & Kashmir", "GB": "Gilgit-Baltistan"}
_SHORT = {"Supreme Court of Pakistan": "SC", "Federal Shariat Court": "FSC", "Lahore High Court": "Lahore",
          "Sindh High Court": "Sindh", "Peshawar High Court": "Peshawar", "Balochistan High Court": "Quetta",
          "Islamabad High Court": "Islamabad", "Azad Jammu & Kashmir": "AJ&K", "Gilgit-Baltistan": "GB"}
_CODES = "|".join(sorted(REPORTERS, key=len, reverse=True))
_CT = "|".join(sorted((re.escape(k) for k in COURTS), key=len, reverse=True))
# PLD 2020 SC 1  |  2021 SCMR 1234  |  2020 YLR 123 (Lahore)
_A = re.compile(rf"\b(?P<code>{_CODES})\s*(?P<year>(?:19|20)\d\d)\s*(?P<court>{_CT})\s*(?P<page>\d{{1,4}})\b", re.I)
_B = re.compile(rf"\b(?P<year>(?:19|20)\d\d)\s*(?P<code>{_CODES})\s*(?:[(\[]\s*(?P<c1>{_CT})\s*[)\]]\s*)?"
                rf"(?P<page>\d{{1,4}})\b(?:\s*[(\[]\s*(?P<c2>{_CT})\s*[)\]])?", re.I)

MSG = {
    "English": {"verified": "VERIFIED ✔ found in the verified database",
                "unindexed": "NOT CONFIRMED - this citation is not in the verified library yet, so do not rely on it. It has been sent to the admin team for verification; once it is added, the full details will show here.",
                "implausible_year": "NOT VALID - this year is impossible for a Pakistani law report.",
                "invalid": "This is not a recognised Pakistani law-report citation. Expected forms: PLD 2020 SC 1, 2021 SCMR 1234, 2020 YLR 123 (Lahore).",
                "source": "Source", "open": "Check at"},
    "Roman": {"verified": "TASDEEQ SHUDA ✔ verified database mein maujood hai",
              "unindexed": "TASDEEQ NAHI HUI - ye citation verified library mein abhi nahi, is liye is par bharosa na karein. Tasdeeq ke liye admin team ko bhej diya gaya hai; add hote hi poori tafseel yahin nazar aayegi.",
              "implausible_year": "GHALAT - ye saal Pakistani law report ke liye mumkin nahi.",
              "invalid": "Ye Pakistani law-report ka durust citation nahi lagta. Durust misalein: PLD 2020 SC 1, 2021 SCMR 1234, 2020 YLR 123 (Lahore).",
              "source": "Zareya", "open": "Yahan check karein"},
    "Urdu": {"verified": "تصدیق شدہ ✔ تصدیق شدہ ڈیٹا بیس میں موجود ہے",
             "unindexed": "تصدیق نہیں ہوئی - یہ حوالہ تصدیق شدہ لائبریری میں ابھی موجود نہیں، اس لیے اس پر انحصار نہ کریں۔ تصدیق کے لیے ایڈمن ٹیم کو بھیج دیا گیا ہے؛ شامل ہوتے ہی مکمل تفصیل یہیں نظر آئے گی۔",
             "implausible_year": "غلط - یہ سال پاکستانی لاء رپورٹ کے لیے ممکن نہیں۔",
             "invalid": "یہ پاکستانی لاء رپورٹ کا درست حوالہ نہیں۔ درست مثالیں: PLD 2020 SC 1، 2021 SCMR 1234، 2020 YLR 123 (Lahore)",
             "source": "ذریعہ", "open": "یہاں دیکھیں"},
}


def _find(text):
    ms = sorted([*_A.finditer(text or ""), *_B.finditer(text or "")], key=lambda m: (m.start(), -m.end()))
    out, end = [], -1
    for m in ms:
        if m.start() >= end: out.append(m); end = m.end()
    return out


def _norm(m):
    g = m.groupdict(); code = g["code"].upper(); year, page = int(g["year"]), int(g["page"])
    ck = (g.get("court") or g.get("c1") or g.get("c2") or "").upper(); court = COURTS.get(ck, "")
    sc = "PCrLJ" if code == "PCRLJ" else code
    if code == "PLD": key, shown = f"PLD|{year}|{court}|{page}", f"PLD {year} {_SHORT.get(court, ck.title())} {page}"
    else: key, shown = f"{sc}|{year}||{page}", f"{year} {sc} {page}"
    return key, shown, {"code": sc, "reporter": REPORTERS.get(code, ""), "year": year, "court": court, "page": page}


def _read_json(path, default):
    try:
        with open(path, encoding="utf-8") as f: return json.load(f)
    except Exception: return default


def _write_json(path, data):
    tmp = f"{path}.{secrets.token_hex(3)}.tmp"
    with open(tmp, "w", encoding="utf-8") as f: json.dump(data, f, ensure_ascii=False, indent=1)
    os.replace(tmp, path)


def _read_jsonl(path):
    out = []
    try:
        with open(path, encoding="utf-8") as f:
            for ln in f:
                try: out.append(json.loads(ln))
                except ValueError: pass
    except OSError: pass
    return out


def _official_key(ref): return "OFF|" + re.sub(r"[^a-z0-9]+", "", str(ref).lower())


def _rows_to_idx(rows):
    idx = {}
    for r in rows:
        if not isinstance(r, dict): continue
        ms = _find(str(r.get("citation", "")))
        if len(ms) == 1: idx[_norm(ms[0])[0]] = r
        elif str(r.get("official", "")) == "1" and str(r.get("citation", "")).strip(): idx[_official_key(r["citation"])] = r      # official judgment, no law-report citation
    return idx


def _global_rows():
    rows = _rjs(GLOBAL_JSON, None)
    if rows is None:                                # first use: take over the old verified_citations.jsonl if it is still on the disk
        rows = _read_jsonl(GLOBAL_IDX)
        if rows: _wjs(GLOBAL_JSON, rows)
    return rows if isinstance(rows, list) else []


def index(user_path=None):
    """Global verified index + this advocate's own 'verified by me' library (his entries win)."""
    with _LOCK:
        if _CACHE["mtime"] != _GV["ver"]: _CACHE["rows"], _CACHE["mtime"] = _rows_to_idx(_global_rows()), _GV["ver"]
        g = _CACHE["rows"]
    return {**g, **(_rows_to_idx(_read_json(user_path, [])) if user_path else {})}


# ---- topic map: Urdu / Roman Urdu / English words -> the English words the library entries use (better search, and shows where the library is empty)
GAP_JSON = f"{DATA}/citation_gaps.json"


def _qn(s): return (s or "").lower().replace("ي", "ی").replace("ك", "ک").replace("ۃ", "ہ").replace("ه", "ہ")


_TOPICS = [(n, tuple(_qn(w) for w in tr), core, secs) for n, tr, core, secs in (
    ("Bail after arrest", ("bail", "zamanat", "ضمانت", "497"), ("bail",), "497 498 crpc"),
    ("Pre-arrest bail", ("pre-arrest", "prearrest", "qabl az giraftari", "قبل از گرفتاری", "interim bail"), ("bail", "pre-arrest", "prearrest"), "498 crpc"),
    ("FIR registration / 22-A, 22-B", ("22-a", "22a", "22-b", "justice of peace", "اندراج مقدمہ", "registration of fir"), ("fir", "registration"), "154 22-a 22-b crpc"),
    ("Quashment of FIR", ("quash", "quashment", "561-a", "561a"), ("quash", "quashment"), "561-a crpc"),
    ("Acquittal 265-K", ("265-k", "265k", "acquittal", "bari"), ("acquittal",), "265-k crpc"),
    ("Cheque dishonour 489-F", ("489-f", "489f", "cheque", "check bounce"), ("cheque", "489-f", "489f"), "489-f ppc"),
    ("Murder / qatl", ("302", "murder", "qatl", "قتل"), ("murder", "qatl"), "302 324 ppc"),
    ("Narcotics", ("narcotic", "narcotics", "منشیات", "manshiyat"), ("narcotic", "narcotics"), "9 cns act"),
    ("Specific performance / agreement to sell", ("specific performance", "agreement to sell", "بیعنامہ", "bayana", "muahida"), ("specific", "agreement"), "12 54 specific relief act"),
    ("Injunction / stay", ("injunction", "order 39", "حکم امتناعی", "stay order"), ("injunction",), "order xxxix cpc"),
    ("Rejection of plaint", ("rejection of plaint", "order 7 rule 11", "o7r11"), ("plaint", "rejection"), "order vii rule 11 cpc"),
    ("Limitation / condonation", ("limitation", "condonation", "میعاد", "mayad"), ("limitation", "condonation"), "limitation act 5"),
    ("Setting aside ex-parte", ("ex parte", "ex-parte", "order 9", "یکطرفہ"), ("ex-parte", "exparte"), "order ix rule 13 cpc"),
    ("Fraud on decree 12(2)", ("12(2)", "fraud", "misrepresentation", "فراڈ"), ("fraud", "misrepresentation"), "12(2) cpc"),
    ("Revision / appeal", ("revision", "appeal", "نگرانی", "اپیل"), ("revision", "appeal"), "115 96 cpc"),
    ("Summary suit Order 37", ("order 37", "summary suit", "leave to defend"), ("summary", "leave"), "order xxxvii cpc"),
    ("Dower / maintenance", ("dower", "haq mehr", "حق مہر", "maintenance", "nafaqa", "نان نفقہ"), ("dower", "maintenance"), "family courts act 1964"),
    ("Khula / dissolution of marriage", ("khula", "خلع", "dissolution of marriage"), ("khula", "dissolution"), "dissolution of muslim marriages act 1939"),
    ("Custody / guardianship", ("custody", "guardian", "hizanat", "حضانت"), ("custody", "guardian", "guardianship"), "guardians and wards act 1890"),
    ("Inheritance / partition", ("inheritance", "wirasat", "وراثت", "partition", "taqseem"), ("inheritance", "partition"), "legal heirs"),
    ("Pre-emption", ("pre-emption", "preemption", "shufa", "شفعہ"), ("pre-emption", "preemption"), "punjab pre-emption act"),
    ("Rent / ejectment", ("ejectment", "kiraya", "کرایہ", "tenant", "rented premises"), ("ejectment", "tenant", "rent"), "punjab rented premises act 2009"),
    ("Evidence / witnesses", ("qanun-e-shahadat", "شہادت", "article 17", "article 79"), ("evidence", "witness"), "qanun-e-shahadat 1984"),
    ("Cheating / fraud 420", ("420", "cheating", "dhoka"), ("cheating",), "420 406 ppc"),
    ("Mutation / revenue", ("mutation", "intiqal", "انتقال", "revenue"), ("mutation", "revenue"), "land revenue act 1967"),
    ("Rape / zina", ("376", "rape", "zina", "زنا"), ("rape", "zina"), "376 ppc"),
)]


def _topics(query):
    q = _qn(query)
    return [t for t in _TOPICS if any(w in q for w in t[1])]


def _gap(topics):
    """Only the topic NAME is counted - never the question text, which may contain a client's facts."""
    if not topics: return
    g = _rjs(GAP_JSON, {})
    for t in topics: g[t[0]] = g.get(t[0], 0) + 1
    _wjs(GAP_JSON, g)


def coverage():
    """Per topic: how many usable (not overruled) library entries mention its words. Fewest first."""
    rows = [r for r in _global_rows() if not str(r.get("status", "")).lower().startswith("overrul")]
    text = [" ".join(str(r.get(k, "")) for k in ("title", "keywords", "sections", "summary")).lower() for r in rows]
    out = [{"topic": n, "entries": sum(1 for t in text if any(w in t for w in core))} for n, _tr, core, _s in _TOPICS]
    return sorted(out, key=lambda x: (x["entries"], x["topic"])), len(rows)


def recheck_list():
    return [{"citation": r.get("citation", ""), "status": r.get("status", "")} for r in _global_rows()
            if str(r.get("status", "")).lower().startswith(("check", "disting"))]


def gaps_list(): return sorted(({"topic": k, "count": v} for k, v in _rjs(GAP_JSON, {}).items()), key=lambda x: -x["count"])


def _binding(court):
    k = str(court or "").lower()
    if "supreme court" in k: return "Supreme Court ka faisla Pakistan ki sab adalaton par paband hai (Article 189)."
    if "high court" in k: return "High Court ka faisla apne ilaqe ki matehat adalaton par paband hai (Article 201); kisi dusre High Court ka faisla sirf rehnuma hai."
    return ""


def _year(r):
    y = re.search(r"(?:19|20)\d{2}", str(r.get("citation", ""))); return int(y.group(0)) if y else 0


_STOPW = {"the", "and", "for", "with", "what", "when", "how", "does", "are", "was", "can", "about", "please", "kya", "hai", "hain", "kaun", "kab", "kaise",
          "kis", "kar", "mere", "meri", "koi", "liye", "wali", "wala", "tareeqa", "tarika", "baare", "client", "case"}


def authorities(query, user_path=None, limit=8):
    """Search groups: every query word is one group; a recognised topic replaces its own trigger words by ONE group of the English words the library uses
    (so 'zamanat' finds entries that say 'bail'). Overruled entries are never offered to the AI. Supreme Court first, then newer."""
    q, tp = _qn(query), _topics(query)
    base = {w for w in re.findall(r"[a-z0-9]{3,}", q) if w not in _STOPW}
    for t in tp: base -= {w for w in t[1] if w in base}
    tg = [set(t[2]) for t in tp]
    groups = [{w} for w in base] + tg
    sec = {w for t in tp for w in re.findall(r"[a-z0-9]{2,}", t[3])}
    if not groups: return []
    need, scored = (1 if tp else min(2, len(groups))), []         # a recognised topic is enough; chat-style questions carry many filler words
    for r in index(user_path).values():
        if str(r.get("status", "")).lower().startswith("overrul"): continue
        strong = " ".join(str(r.get(k, "")) for k in ("title", "keywords", "sections")).lower()
        weak = " ".join(str(r.get(k, "")) for k in ("summary", "court")).lower()
        hit = [g for g in groups if any(w in strong or w in weak for w in g)]
        if tp and not any(g is h for g in tg for h in hit): continue
        if len(hit) < need: continue
        s = sum(2 if any(w in strong for w in g) else 1 for g in hit) + sum(0.5 for w in sec if w in strong)
        s += 1.5 if "supreme" in weak else 0.5 if "high court" in weak else 0
        scored.append((s, _year(r), r))
    scored.sort(key=lambda x: (-x[0], -x[1]))
    return [r for _, _, r in scored[:limit]]


def authorities_block(query, user_path=None, auths=None):
    head = ("CITATION POLICY (mandatory): never write a case name, case citation or law-report citation (PLD, SCMR, CLC, YLR, PCrLJ ...) from your own "
            "memory or in your own words. The ONLY case law you may use is the numbered VERIFIED AUTHORITIES below, and you must refer to each one ONLY by "
            "its tag such as [A1] - the system prints the exact citation in its place. Never change a tag or invent one. After the legal position, finish "
            "with a short section titled 'Suggested next steps' (Kya karna chahiye): practical steps for the advocate - what to argue, which documents and "
            "dates to check, and which authority [A#] supports which point. ")
    a = authorities(query, user_path) if auths is None else auths
    if not a:
        return head + ("VERIFIED AUTHORITIES: none in the verified library for this topic. Do not cite any case. Say 'Is mauzoo par verified citation "
                       "library mein abhi nahi hai' once, then give the statutory position and the suggested next steps.")
    return head + "VERIFIED AUTHORITIES:\n" + "\n".join(
        f"[A{i}] {r['citation']} | {r.get('title', '')} | {r.get('court', '')} | {str(r.get('summary', ''))[:350]}" for i, r in enumerate(a, 1))


def _fmt(r):
    bits = [f"{r['citation']}  -  {r.get('title', '')}", "   " + " | ".join(str(r[k]) for k in ("court", "date", "bench") if r.get(k))]
    st = str(r.get("status", "")).strip().lower()
    if st.startswith("overrul"): bits.append("   WARNING: is faisle ko overrule kiya ja chuka hai - bharosa na karein.")
    elif st.startswith("disting") or st.startswith("check"): bits.append("   Ehtiyat: " + ("is faisle se farq (distinguish) kiya gaya hai" if st.startswith("disting") else "dobara check kiya jana hai") + " - istemal se pehle dekh lein.")
    bind = _binding(r.get("court"))
    if bind: bits.append("   " + bind)
    if r.get("summary"): bits.append(("   Khulasa (AI ne likha, admin ne dekha): " if r.get("ai_assisted") else "   Khulasa: ") + str(r["summary"])[:500])
    if r.get("extract"): bits.append("   Faisle ki asal ibarat" + (f" (para {r['para']})" if r.get("para") else "") + ': "' + str(r["extract"])[:700] + '"')
    src, vb, vo = str(r.get("source", "")), str(r.get("verified_by", "")), str(r.get("verified_on", ""))
    bits.append(f"   Zareya: {src}" + (f" | tasdeeq: {vb}" if vb else "") + (f" ({vo})" if vo else "") + (f" | {r['url']}" if r.get("url") else ""))
    return "\n".join(b for b in bits if b.strip())


def render(text, auths, user_path=None, query=""):
    """[A1] tags -> the exact library citation; any other citation-looking string is flagged; the verified details are appended by the SERVER."""
    tags, used = {f"A{i}": r for i, r in enumerate(auths, 1)}, []
    def sub(m):
        r = tags.get(m.group(1))
        if not r: return ""
        if r not in used: used.append(r)
        return r["citation"]
    text, rep = scrub(re.sub(r"\[(A\d{1,2})\]", sub, text or ""), user_path)
    head = "\n\nVERIFIED AUTHORITIES (library se - AI ne nahi likhe):\n"
    if not auths:
        _gap(_topics(query))
        return text + head + "Is mauzoo par verified library mein abhi koi citation nahi hai.", rep
    rest = [r for r in auths if r not in used]
    body = "\n\n".join(_fmt(r) for r in used) or "(AI ne inmein se kisi ka hawala nahi diya)"
    return text + head + body + (("\n\nMutaliq (ye bhi library mein hain):\n" + "\n\n".join(_fmt(r) for r in rest)) if rest else ""), rep


def scrub(text, user_path=None):
    """Mark every citation-looking string in model output that is not in the verified index. Returns (text, report)."""
    idx, out, last, rep = index(user_path), [], 0, []
    for m in _find(text):
        key, shown, _ = _norm(m); rec = idx.get(key)
        out.append(text[last:m.start()])
        out.append(m.group(0) if rec else f"[CITATION NOT VERIFIED: {m.group(0)}]")
        rep.append({"citation": shown, "status": "verified" if rec else "unverified", "title": (rec or {}).get("title", "")})
        last = m.end()
    out.append(text[last:])
    return "".join(out), rep


def _links(meta):
    L = [("Pakistan Law Site", "https://www.pakistanlawsite.com/"), ("EastLaw", "https://eastlaw.pk/")]
    if meta["code"] == "SCMR" or meta["court"] == "Supreme Court of Pakistan": L.append(("Supreme Court of Pakistan", "https://www.supremecourt.gov.pk/"))
    elif meta["court"] == "Lahore High Court": L.append(("Lahore High Court", "https://lhc.gov.pk/"))
    return L


def verify(text, user_path=None, language="English"):
    M = MSG["Urdu" if language.startswith("Urdu") else "Roman" if language.startswith("Roman") else "English"]
    idx, items = index(user_path), []
    for m in _find(text):
        key, shown, meta = _norm(m); rec = idx.get(key)
        st = "implausible_year" if not (1947 <= meta["year"] <= date.today().year) else ("verified" if rec else "unindexed")
        items.append({"input": m.group(0), "citation": shown, **meta, "status": st, "links": _links(meta),
                      "record": {k: str(rec.get(k, "")) for k in ("title", "court", "date", "bench", "summary", "url", "source", "verified_by")} if rec else None})
    if not items: return {"status": "invalid_format", "items": [], "report": M["invalid"]}
    lines = []
    for it in items:
        lines.append(f"{it['citation']}  -  {M[it['status']]}")
        r = it["record"]
        if r:
            lines.append(f"   {r['title']} | {r['court']} {r['date']}".rstrip())
            lines.append(f"   {M['source']}: {r['source']}" + (f" ({r['verified_by']})" if r["verified_by"] else ""))
            if r["summary"]: lines.append(f"   {r['summary']}")
        else: pass          # not in the library: the advocate stays in the app, the admin team verifies it
    return {"status": items[0]["status"], "items": items, "report": "\n".join(lines)}


def add_user(user_path, rec):
    """Advocate confirmed this citation himself (e.g. on Pakistan Law Site) and saves it to his own library."""
    ms = _find(rec.get("citation", ""))
    if len(ms) != 1: return False, "Citation ka format durust nahi (misal: PLD 2020 SC 1)."
    if not str(rec.get("title", "")).strip(): return False, "Case ka title (parties) likhein."
    rows, (key, shown, _) = _read_json(user_path, []), _norm(ms[0])
    if key in _rows_to_idx(rows): return False, "Ye citation aap ki library mein pehle se hai."
    if len(rows) >= 3000: return False, "Library full hai."
    rows.append({**{k: str(rec.get(k, ""))[:600] for k in ("title", "court", "summary", "source", "verified_by")},
                 "citation": shown, "added": date.today().isoformat()})
    _write_json(user_path, rows)
    return True, ""


def import_text(text, source=""):
    """Admin: bulk-load JSONL or CSV (columns: citation,title,court,date,bench,summary,keywords,sections,url,source)."""
    text, rows = (text or "").strip().lstrip("\ufeff"), []
    if text.startswith("{"):
        for ln in text.splitlines():
            try: rows.append(json.loads(ln))
            except ValueError: pass
    else: rows = list(csv.DictReader(io.StringIO(text)))
    allrows = _global_rows(); have, new, skipped, full = _rows_to_idx(allrows), [], 0, False
    size = len(json.dumps(allrows, ensure_ascii=False))
    for r in rows:
        ms = _find(str(r.get("citation", "")))
        if len(ms) != 1 or not str(r.get("title", "")).strip(): skipped += 1; continue
        key, shown, _ = _norm(ms[0])
        if key in have: skipped += 1; continue
        r = {str(k): str(v)[:2000] for k, v in r.items() if k and v}
        r["citation"], r["source"] = shown, r.get("source") or source or "import"
        sz = len(json.dumps(r, ensure_ascii=False)) + 2
        if size + sz > LIB_BYTES: full = True; break
        size += sz; have[key] = r; new.append(r)
    if new:
        _wjs(GLOBAL_JSON, allrows + new); _GV["ver"] += 1
        _resolve([_norm(_find(r["citation"])[0])[0] for r in new])
    return {"added": len(new), "skipped": skipped, **({"library_full": True} if full else {})}


def add_global(rec, by=""):
    """Admin team: one verified entry. 'source' (where it was verified) is mandatory."""
    if not str(rec.get("source", "")).strip(): return False, "Zareya (source) likhein - ye kahan se tasdeeq hua (jaise Pakistan Law Site, court ki website)."
    rec = {k: v for k, v in rec.items() if str(v or "").strip()}
    if by: rec["verified_by"] = by
    rec["verified_on"] = date.today().isoformat()
    ref = str(rec.get("citation", "")).strip()
    if len(_find(ref)) != 1 and str(rec.get("official", "")) == "1":      # official judgment without a PLD/SCMR citation: identified by its case reference
        if len(ref) < 8 or not str(rec.get("title", "")).strip(): return False, "Case reference aur title zaroori hain."
        allrows = _global_rows()
        if _official_key(ref) in _rows_to_idx(allrows): return False, "Ye faisla pehle se library mein hai."
        rec = {str(k): str(v)[:2000] for k, v in rec.items()}
        if len(json.dumps(allrows + [rec], ensure_ascii=False)) > LIB_BYTES: return False, "Library full hai."
        _wjs(GLOBAL_JSON, allrows + [rec]); _GV["ver"] += 1
        return True, ""
    r = import_text(json.dumps(rec, ensure_ascii=False))
    if r["added"]: return True, ""
    return False, "Library full hai." if r.get("library_full") else "Citation ka format durust nahi, title khali hai, ya ye pehle se library mein hai."


def search(query, user_path=None, limit=20):
    q = (query or "").strip(); ms = _find(q)
    if ms: rec = index(user_path).get(_norm(ms[0])[0]); return [rec] if rec else []
    return authorities(q, user_path, limit)


def request_cite(citation):
    """A citation somebody tried that is not in the library: queued (once, with a counter) for the admin team."""
    ms = _find(citation)
    if len(ms) != 1: return
    key, shown, meta = _norm(ms[0])
    if not (1947 <= meta["year"] <= date.today().year): return
    rows, today = _rjs(REQ_JSON, []), date.today().isoformat()
    for r in rows:
        if r.get("key") == key: r["count"], r["last"] = r.get("count", 1) + 1, today; break
    else:
        if len(rows) >= 500: return
        rows.append({"key": key, "citation": shown, "count": 1, "first": today, "last": today})
    _wjs(REQ_JSON, rows)


def requests_list(): return sorted(_rjs(REQ_JSON, []), key=lambda r: (-r.get("count", 1), r.get("citation", "")))


def _resolve(keys):
    rows, ks = _rjs(REQ_JSON, []), set(keys)
    keep = [r for r in rows if r.get("key") not in ks]
    if len(keep) != len(rows): _wjs(REQ_JSON, keep)
''')      # verified citations only
cs = _embed('court_sync', r'''"""Court sync for Punjab district courts (stdlib only).

dsj.punjab.gov.pk asks a human arithmetic question before every search, and this module never answers it.
Flow: the advocate searches on the portal himself, copies the result table and pastes it into the app -> parse ->
match to his clients -> update. If PITB / D&SJ ever grants an official feed, fetch_official() + sync_one() reuse the
same matching, so full automation is a configuration change (COURT_API_URL / COURT_API_KEY), not a rewrite."""
import copy, json, re, urllib.parse, urllib.request
from datetime import date, datetime
from html.parser import HTMLParser

STOP = {"the", "and", "versus", "state", "son", "wife", "daughter", "muhammad", "mohammad", "mst", "mussammat",
        "etc", "others", "ors", "govt", "government", "province", "punjab", "pakistan", "federation"}      # too common to prove a match
_MON = {m: i for i, m in enumerate("jan feb mar apr may jun jul aug sep oct nov dec".split(), 1)}
_PAIR = re.compile(r"(?<![\d/\-])(\d{1,9})\s*[/\-]\s*(\d{4}|\d{2})(?![\d/\-])")


def _year(y):
    y = int(y)
    return y if y >= 100 else (2000 + y if 2000 + y <= date.today().year else 1900 + y)


def _pairs_of(text):
    """Case/FIR numbers as (number, year): '101/2024', '101/24' and the portal's joined form '205411926' (number + 2-digit year)."""
    out = {(str(int(n)), _year(y)) for n, y in _PAIR.findall(text or "")}
    for tok in re.findall(r"(?<![\d/\-])(\d{5,10})(?![\d/\-])", text or ""):
        y = _year(tok[-2:])
        if 1990 <= y <= date.today().year and int(tok[:-2]): out.add((str(int(tok[:-2])), y))
    return out


# ---- Urdu helpers: matching only (the court's own spelling is what gets stored)
_UR_MAP = {0x64A: "ی", 0x649: "ی", 0x643: "ک", 0x647: "ہ", 0x6C3: "ہ", 0x6D3: "ے", 0x200C: None, 0x200D: None}
_UR_MAP.update({0x660 + i: 48 + i for i in range(10)}); _UR_MAP.update({0x6F0 + i: 48 + i for i in range(10)})
_DIG = {k: v for k, v in _UR_MAP.items() if isinstance(v, int)}


def _ur(s): return (s or "").translate(_UR_MAP)


def _dg(s): return (s or "").translate(_DIG)          # Arabic/Persian digits -> 0-9, text untouched


STOP_UR = {_ur(w) for w in ("بنام", "وغیرہ", "وغیر", "سرکار", "صوبہ", "پنجاب", "مسماۃ", "مسمات", "بیگم", "محمد", "ریاست")}


def _tok(s):
    s = s or ""
    return ({t for t in re.findall(r"[a-z]{3,}", s.lower()) if t not in STOP}
            | {t for t in re.findall(r"[\u0600-\u06FF]{3,}", _ur(s)) if t not in STOP_UR})


def _nt(s):
    return re.sub(r"[^a-z0-9]+", " ", (s or "").lower()).strip()


def _dates(text):
    raw = [(y, m, d) for d, m, y in re.findall(r"\b(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})\b", text)]
    raw += re.findall(r"\b(\d{4})-(\d{1,2})-(\d{1,2})\b", text)
    raw += [(y, str(_MON[mo.lower()]), d) for d, mo, y in re.findall(r"\b(\d{1,2})[\s-]([A-Za-z]{3})[a-z]*[\s,-]+(\d{4})\b", text) if mo.lower() in _MON]
    out = []
    for y, m, d in raw:
        try: out.append(date(int(y), int(m), int(d)).isoformat())
        except ValueError: pass
    return out


class _T(HTMLParser):
    def __init__(s):
        super().__init__(); s.rows, s.cur, s.cell, s.hdr = [], None, None, False

    def handle_starttag(s, t, a):
        if t == "tr": s.cur, s.hdr = [], False
        elif t in ("td", "th") and s.cur is not None:
            s.cell = []; s.hdr = s.hdr or t == "th"

    def handle_endtag(s, t):
        if t in ("td", "th") and s.cell is not None and s.cur is not None:
            s.cur.append(re.sub(r"\s+", " ", "".join(s.cell)).strip()); s.cell = None
        elif t == "tr" and s.cur is not None:
            if any(s.cur): s.rows.append((s.hdr, s.cur))
            s.cur = None

    def handle_data(s, d):
        if s.cell is not None: s.cell.append(d)


def _row(cells, header):
    text = " | ".join(cells)
    r = {"text": text[:400], "pairs": _pairs_of(text), "date": "", "stage": "", "court": "", "ps": "", "title": ""}
    for h, c in zip(header or [], cells):
        if not c: continue
        if "next" in h or "hearing" in h or ("date" in h and "institut" not in h and "decision" not in h):
            ds = _dates(c); r["date"] = r["date"] or (ds[0] if ds else "")
        elif "stage" in h or "status" in h or "purpose" in h: r["stage"] = c[:80]
        elif "court" in h or "judge" in h: r["court"] = c[:80]
        elif "police" in h or h.strip() in ("ps", "p.s", "p.s."): r["ps"] = c[:60]
        elif "title" in h or "part" in h: r["title"] = c[:120]
    if not r["date"]:
        ds = sorted(set(_dates(text))); fut = [d for d in ds if d >= date.today().isoformat()]
        r["date"] = fut[0] if fut else (ds[-1] if ds else "")
    if not r["ps"]:
        m = re.search(r"(?:P\.?\s?S\.?|Police Station|Thana)\s*[:\-]?\s*([^|,;\n]{2,40})", text, re.I)
        if m: r["ps"] = m.group(1).strip()
    r["tokens"] = _tok(text)
    return r


# ---- Urdu district-court cause list ("فہرست مقدمات مورخہ ...") copied as text from the court's page
_URD = {_ur(k).replace(" ", ""): v for k, v in {
    "سیالکوٹ": "Sialkot", "لاہور": "Lahore", "گوجرانوالہ": "Gujranwala", "فیصل آباد": "Faisalabad", "ملتان": "Multan",
    "راولپنڈی": "Rawalpindi", "سرگودھا": "Sargodha", "بہاولپور": "Bahawalpur", "شیخوپورہ": "Sheikhupura", "قصور": "Kasur",
    "ساہیوال": "Sahiwal", "اوکاڑہ": "Okara", "گجرات": "Gujrat", "جہلم": "Jhelum", "چکوال": "Chakwal", "اٹک": "Attock",
    "منڈی بہاؤالدین": "Mandi Bahauddin", "حافظ آباد": "Hafizabad", "نارووال": "Narowal", "ڈیرہ غازی خان": "Dera Ghazi Khan",
    "مظفرگڑھ": "Muzaffargarh", "لیہ": "Layyah", "راجن پور": "Rajanpur", "رحیم یار خان": "Rahim Yar Khan",
    "بہاولنگر": "Bahawalnagar", "لودھراں": "Lodhran", "خانیوال": "Khanewal", "وہاڑی": "Vehari", "پاکپتن": "Pakpattan",
    "جھنگ": "Jhang", "چنیوٹ": "Chiniot", "ٹوبہ ٹیک سنگھ": "Toba Tek Singh", "بھکر": "Bhakkar", "میانوالی": "Mianwali",
    "خوشاب": "Khushab", "ننکانہ صاحب": "Nankana Sahib"}.items()}
# ---- English labels (clients are kept in English): court's own Urdu text stays in extra["*_ur"]
def _key(s): return _ur(s).replace(" ", "")


def _rule(rules, text, default=None):
    k = _key(text)
    for words, en in rules:
        if all(_key(w) in k for w in words): return en
    return text if default is None else default


_STAGE_RULES = [(("ابتدائی", "بحث"), "Preliminary arguments"), (("بحث",), "For arguments"), (("طلبی",), "Summoning of witnesses"),
                (("شہادت",), "For evidence"), (("وکالت",), "Filing of vakalatnama"), (("مثل",), "Along with other file"),
                (("ضمانت",), "For filing of surety bond"), (("اشتہار",), "Ishtehar (publication of notice)"), (("نوٹس",), "Notice")]
_TYPE_RULES = [(("ضمانت", "گرفتاری"), "Pre-arrest bail application"), (("22",), "Application u/s 22-A/B Cr.P.C (registration of FIR)"),
               (("ریپ",), "Anti-rape cases"), (("اغوا",), "Other sessions cases (abduction, 365-B PPC)"), (("فیملی",), "Family appeal"),
               (("بحالی",), "Restoration application (appeal / suit / revision / application)"),
               (("دیوانی", "درمیانی"), "Civil appeal against interim order"), (("دیوانی", "آخیر"), "Civil appeal against final order"),
               (("نگرانی",), "Civil revision"), (("37",), "Suit under Order XXXVII CPC"), (("متفرق",), "Civil miscellaneous")]
_PS_EN = [(_key(k), v) for k, v in (
    ("نیکاپورہ", ("Nikapura", "Nekapura", "Neka Pura", "Nikka Pura")), ("کوتوالی", ("Kotwali", "Kotwali Sialkot")),
    ("رنگ پورہ", ("Rangpura", "Rang Pura")), ("مرادپور", ("Muradpur", "Murad Pur", "Moradpur")),
    ("کوٹلی سید امیر", ("Kotli Syed Amir", "Kotli Sayed Amir", "Kotli Syed Ameer", "Kotli Sayyed Amir")),
    ("سمبڑیال", ("Sambrial",)), ("ڈسکہ", ("Daska",)), ("پسرور", ("Pasrur",)), ("چونڈہ", ("Chawinda", "Chawnda")),
    ("بیگووالہ", ("Begowala", "Begowal")), ("اگوکی", ("Ugoke",)), ("بڈیانہ", ("Badiana",)), ("کینٹ", ("Cantt", "Cantonment")),
    ("سول لائن", ("Civil Lines",)))]


def _ps_en(s):
    k = _key(s)
    for key, names in _PS_EN:
        if key in k: return names[0], list(names[1:])
    return s, []


def _judge_en(judge_ln):
    """'Additional District & Sessions Judge <name>' when the designation is recognised, else the court's own Urdu line."""
    k = _key(judge_ln)
    nm = re.search(r"^جناب\s*(.*?)\s*صاحب", judge_ln)
    desig = ("Additional District & Sessions Judge" if "ایڈیشنل" in k and "سیشن" in k else
             "Senior Civil Judge" if "سینئرسول" in k else "Civil Judge" if "سولجج" in k else
             "District & Sessions Judge" if "ڈسٹرکٹ" in k and "سیشن" in k else "")
    return f"{desig} {nm.group(1)}".strip() if desig and nm else ""


_UR_C = re.compile(r"^(\d{1,2}[./-]\d{1,2}[./-]\d{4})\s+(\d{5,10})\s+(\d{1,4})$")
_UR_FIR = re.compile(r"^\d{1,6}\s*/\s*\d{4}$")


def _urdu_district(text):
    t = _ur(text).replace(" ", "")
    for k in sorted(_URD, key=len, reverse=True):
        if k in t: return _URD[k]
    return ""


def is_urdu_list(raw):
    t = _dg(raw or "")
    return bool(re.search(r"[\u0600-\u06FF]{3}", t) and re.search(r"^\s*\d{1,2}[./-]\d{1,2}[./-]\d{4}\s+\d{5,10}\s+\d{1,4}\s*$", t, re.M))


def _ur_row(cur, m, sec, ldate, court):
    cells, (inst, cno, serial) = cur["cells"], m.groups()
    ti = next((i for i, c in enumerate(cells) if "بنام" in _ur(c)), 1 if len(cells) > 1 else 0)
    title = cells[ti] if cells else ""
    pre, post = cells[:ti], cells[ti + 1:]
    nd = next((d for c in pre for d in _dates(c)[:1]), "")                 # next date, if the court already filled it in
    words = [c for c in pre if not _dates(c)]
    today_stage, next_stage = (words[-1] if words else ""), (words[-2] if len(words) > 1 else "")
    fir = next((c for c in post if _UR_FIR.match(c)), "")
    parts = re.split(r"\s*بنام\s*", title, maxsplit=1); first, opp = parts[0], (parts[1] if len(parts) > 1 else "")
    date_, stage = (nd, next_stage) if nd else (ldate, today_stage)           # no next date yet: the list's own day is the hearing
    stage_ur, sec_ur = stage, sec
    stage, sec = _rule(_STAGE_RULES, stage), _rule(_TYPE_RULES, sec)
    ps_en, ps_alias = _ps_en(cur["ps"])
    ex = {"type": sec, "case_no": cno, "first_party": first.strip(), "opposite_party": opp.strip(), "us": ", ".join(c for c in post if c != fir),
          "ps_company": ps_en, "board_date": ldate, "registered": (_dates(inst) or [""])[0], "serial_no": serial,
          "stage_ur": stage_ur if stage_ur != stage else "", "type_ur": sec_ur if sec_ur != sec else "", "ps_ur": cur["ps"] if cur["ps"] != ps_en else ""}
    return {"text": f"{title} | {fir or cno} | {stage}"[:400], "pairs": _pairs_of(cno) | (_pairs_of(fir) if fir else set()), "date": date_,
            "stage": stage[:80], "court": court[:80], "ps": ps_en[:60], "ps_alias": ps_alias, "title": title[:120], "tokens": _tok(f"{first} {opp}"), "listed": not nd,
            "extra": {k: v for k, v in ex.items() if v}}


def urdu_list_rows(raw):
    """One case = 1-3 lines: [stage, title, sections, FIR] / [police station] / [institution date, case number, serial].
    The list's own date is the hearing date; a next date is used instead only when the court has already filled it in."""
    lines = [ln.strip() for ln in _dg(raw).replace("\r", "").split("\n") if ln.strip()]
    skip, ldate, court_ln, judge_ln = set(), "", "", ""
    for i, ln in enumerate(lines[:10]):
        n = _ur(ln)
        if "نمبرشمار" in n.replace(" ", ""): skip.add(i)                      # column header
        elif "مورخ" in n and "فہرست" in n: ldate = (_dates(ln) or [""])[0]; skip.add(i)
        elif i < 3 and not court_ln and ("کورٹ" in n or "عدالت" in n): court_ln = ln; skip.add(i)
        elif i < 4 and not judge_ln and ("جج" in n or ln.startswith("جناب")): judge_ln = ln; skip.add(i)
    dist = _urdu_district(court_ln) or _urdu_district(judge_ln)                # Latin district name so the same-city check in _score works
    judge = re.sub(r"^جناب\s*", "", judge_ln).replace(" صاحب", "").strip()
    judge = _judge_en(judge_ln) or judge
    court = " - ".join(x for x in (dist, judge or court_ln) if x)
    out, sec, cur = [], "", None
    for i, ln in enumerate(lines):
        if i in skip: continue
        m = _UR_C.match(re.sub(r"\s+", " ", ln))
        if m:
            if cur: out.append(_ur_row(cur, m, sec, ldate, court))
            cur = None; continue
        cells = [c for c in (x.strip(" \u200c\u200d") for x in re.split(r"\t+|\s{3,}", ln)) if c]
        if "بنام" in _ur(ln) or len(cells) >= 2: cur = {"cells": cells, "ps": ""}
        elif cur is not None:
            if not cur["ps"]: cur["ps"] = ln
        else: sec = ln                                                             # category heading (family appeal, civil revision, ...)
    return out


def parse(raw):
    """HTML table, or text copied from a table (tab / 3+ spaces / pipe between cells, one row per line)."""
    raw = raw or ""
    if re.search(r"DAILY\s+DIARY\s+ON", raw, re.I): return board_text_rows(raw)
    if is_urdu_list(raw):
        rows = urdu_list_rows(raw)
        if rows: return rows
    if re.search(r"<tr[\s>]", raw, re.I):
        p = _T(); p.feed(raw); rows = p.rows
    else:
        rows = []
        for ln in raw.splitlines():
            cells = [x.strip() for x in re.split(r"\t|\s{3,}|\s\|\s", ln) if x.strip()]
            if cells: rows.append((False, cells))
    header, out = None, []
    for is_h, cells in rows:
        if is_h: header = [x.lower() for x in cells]; continue
        r = _row(cells, header)
        if r["pairs"] or r["title"] or (r["date"] and len(cells) > 2): out.append(r)
    return out


def clean_row(r):
    d = str(r.get("date", ""))[:10]
    try: date.fromisoformat(d)
    except ValueError: d = ""
    return {"text": str(r.get("text", ""))[:400], "date": d, "stage": str(r.get("stage", ""))[:80], "court": str(r.get("court", ""))[:80],
            "ps": str(r.get("ps", ""))[:60], "title": str(r.get("title", ""))[:120], "pairs": sorted([n, y] for n, y in r.get("pairs", []) or []),
            "extra": {str(k)[:40]: str(v)[:300] for k, v in (r.get("extra") or {}).items()}}


DISTRICTS = ["Attock", "Bahawalnagar", "Bahawalpur", "Bhakkar", "Chakwal", "Chiniot", "Dera Ghazi Khan", "Faisalabad", "Gujranwala", "Gujrat", "Hafizabad",
             "Jhang", "Jhelum", "Kasur", "Khanewal", "Khushab", "Lahore", "Layyah", "Lodhran", "Mandi Bahauddin", "Mianwali", "Multan", "Muzaffargarh",
             "Narowal", "Nankana Sahib", "Okara", "Pakpattan", "Rahim Yar Khan", "Rajanpur", "Rawalpindi", "Sahiwal", "Sargodha", "Sheikhupura", "Sialkot",
             "Toba Tek Singh", "Vehari"]


def district_in(text):
    t = f" {_nt(text)} "
    for d in sorted(DISTRICTS, key=len, reverse=True):
        if f" {_nt(d)} " in t: return d
    return ""


def district_of(client):
    """Client's district: the 'district' field, else a district name typed in 'court'. '' if unknown."""
    return district_in(client.get("district", "")) or district_in(client.get("court", ""))


def _score(client, row, hint_ps=""):
    cd, rd = district_of(client), district_in(row.get("court", ""))
    if cd and rd and cd != rd and "high court" not in row.get("court", "").lower(): return 0, ""       # same number, different city = different case
    ids = _pairs_of(f"{client.get('fir', '')} {client.get('case_no', '')}")
    nm = _tok(client.get("name", ""))
    if not (ids & set(map(tuple, row["pairs"]))):
        if len(nm) >= 2 and nm <= row.get("tokens", set()): return 50, "naam milta hai magar FIR/case number nahi milta"
        return 0, ""
    s, why = 80, "FIR/case number aur saal milte hain"
    if any(len(n) >= 5 for n, _ in ids & set(map(tuple, row["pairs"]))): s, why = s + 10, why + ", lamba case number"
    ctext = _nt(" ".join(str(client.get(k, "")) for k in ("fir", "court", "notes", "ps")))
    cands = [_nt(row.get("ps") or hint_ps)] + [_nt(a) for a in row.get("ps_alias", [])]
    if any(a and (a in ctext or a.replace(" ", "") in ctext.replace(" ", "")) for a in cands): s, why = s + 15, why + ", thana bhi"
    parties = _tok(" ".join(str(client.get(k, "")) for k in ("name", "first_party", "opposite_party")))
    if parties & row.get("tokens", set()): s, why = s + 10, why + ", naam/party bhi"
    ct, rt = _tok(client.get("case_type", "")), _tok((row.get("extra") or {}).get("type", ""))
    if ct and rt and ct & rt: s, why = s + 5, why + ", case ki type bhi"          # a hint only: never enough alone for a short number
    return s, why


def apply(c, row, today=None):
    """Update one client from one portal row. A board dated in the past never changes the current date or stage."""
    today, ch, nd, ex = today or date.today().isoformat(), [], row.get("date") or "", row.get("extra") or {}
    if row.get("listed") and nd == today and str(c.get("date") or "") > today: return ch      # today's list never pulls back a later date the advocate/court already set
    fresh = not nd or nd >= today
    if nd and nd >= today and nd != c.get("date"):
        ch.append(f"agli peshi {c.get('date') or '-'} -> {nd}"); c["date"] = nd
        c["notify"] = {"date": nd, "stage": row.get("stage", ""), "court": row.get("court", ""), "created": today}      # app offers to message the client
    if row.get("stage") and fresh and row["stage"] != c.get("stage"): ch.append(f"stage: {row['stage']}"); c["stage"] = row["stage"]
    if row.get("court"):
        c["court_portal"] = row["court"]
        if not c.get("court"): c["court"] = row["court"]
    if ex:
        c["portal"] = {**ex, "synced": today}
        for ck, ek in (("case_type", "type"), ("first_party", "first_party"), ("opposite_party", "opposite_party")):
            if ex.get(ek) and not c.get(ck): c[ck] = ex[ek]                    # fill what the advocate left empty, never overwrite
        if fresh:
            if ex.get("previous_date"): c["prev_date"] = ex["previous_date"]
            if ex.get("last_proceedings"): c["last_proceedings"] = ex["last_proceedings"]          # everything the board shows: parties, U/S, comments, last proceedings...
    if ch:
        bits = [f"Agli tareekh: {nd}" if nd and nd >= today else "", f"Stage: {row['stage']}" if row.get("stage") and fresh else "", row.get("court", "")]
        c["notes"] = f"[{today}] " + " | ".join(x for x in bits if x) + f" (court portal) | {c.get('notes', '')}".rstrip(" |")
    h = c.setdefault("history", [])
    ent = {"seen": today, "date": nd, "stage": row.get("stage", ""), "src": "court portal"}
    if not any(isinstance(x, dict) and x.get("date") == nd and x.get("stage") == ent["stage"] for x in h):
        h.append(ent); h[:] = h[-50:]
    c["last_synced"] = datetime.now().isoformat(timespec="minutes")
    return ch


def _prev(r):
    return " | ".join(x for x in (r.get("title"), r.get("date"), r.get("stage"), r.get("court")) if x) or r.get("text", "")[:120]


def _collapse(rows, today):
    """The same case listed on several dates (a range board): keep the earliest hearing still ahead, else the latest."""
    best, out = {}, []
    for r in rows:
        if not r["pairs"]: out.append(r); continue
        k, o = (frozenset(map(tuple, r["pairs"])), r.get("title", "")), best.get((frozenset(map(tuple, r["pairs"])), r.get("title", "")))
        a, b = (o or {}).get("date") or "", r.get("date") or ""
        af, bf = bool(a) and a >= today, bool(b) and b >= today
        if o is None or (bf and not af) or (bf and af and b < a) or (not af and not bf and b > a): best[k] = r
    return out + list(best.values())


def ingest(raw, clients, district="", ps="", today=None, rows=None):
    """Auto-update only when FIR/case number matches AND (thana or name) agrees and no other row ties; else ask the advocate."""
    rows = _collapse(rows if rows is not None else parse(raw), today or date.today().isoformat())
    for r in rows:
        if ps and not r["ps"]: r["ps"] = ps
    updated, sugg, hit, nf = [], [], set(), []
    td = today or date.today().isoformat()
    rds = {district_in(r.get("court", "")) for r in rows} - {""}
    for i, c in enumerate(clients):
        sc = sorted(((*_score(c, r, ps), k) for k, r in enumerate(rows)), key=lambda t: -t[0])
        if not sc or sc[0][0] == 0:
            cd, dd = district_of(c), str(c.get("date") or "")
            if _pairs_of(f"{c.get('fir', '')} {c.get('case_no', '')}") and (not dd or dd <= td) and not (cd and rds and cd not in rds):
                nf.append({"client_index": i, "client": c.get("name", ""), "case_no": str(c.get("case_no") or c.get("fir") or "")[:60]})
            continue
        s, why, k = sc[0]; tie = len(sc) > 1 and sc[1][0] >= s; hit.add(k)
        if s >= 90 and not tie:
            updated.append({"client": c.get("name", ""), "fir": c.get("fir", ""), "case_no": str(c.get("case_no") or c.get("fir") or ""), "changes": apply(c, rows[k], today) or ["koi naya badlao nahi"]})
        else:
            sugg.append({"client_index": i, "client": c.get("name", ""), "case_no": str(c.get("case_no") or c.get("fir") or ""), "preview": _prev(rows[k]), "row": clean_row(rows[k]),
                         "why": why + (" (aik se ziyada rows milti hain)" if tie else "")})
    return {"rows": len(rows), "dated": any(r.get("date") for r in rows), "updated": updated, "suggestions": sugg, "not_found": nf[:100], "not_found_total": len(nf), "unmatched_rows": max(0, len(rows) - len(hit))}


# ---- official feed (only if PITB / D&SJ grants one). Expected JSON: {"rows":[{"fir_no","fir_year","case_no","case_year",
# "ps","title","next_date","stage","court"}]}
def _n(v):
    d = re.sub(r"\D", "", str(v or ""))
    return int(d) if d else 0


def rows_from_json(items):
    out = []
    for it in items or []:
        pairs = {(str(_n(a)), _year(_n(b))) for a, b in ((it.get("fir_no"), it.get("fir_year")), (it.get("case_no"), it.get("case_year"))) if _n(a) and _n(b)}
        text = " ".join(str(v) for v in it.values())
        out.append({"text": text[:400], "pairs": pairs, "date": str(it.get("next_date") or it.get("date") or "")[:10], "stage": str(it.get("stage", ""))[:80],
                    "court": str(it.get("court", ""))[:80], "ps": str(it.get("ps", ""))[:60], "title": str(it.get("title", ""))[:120], "tokens": _tok(text)})
    return out


def fetch_official(url, key, client, timeout=20):
    ids = sorted(_pairs_of(f"{client.get('fir', '')} {client.get('case_no', '')}"))
    if not ids: return []
    q = urllib.parse.urlencode({"fir_no": ids[0][0], "fir_year": ids[0][1], "ps": client.get("ps", ""), "district": client.get("district", "")})
    req = urllib.request.Request(f"{url}?{q}", headers={"X-Api-Key": key, "Accept": "application/json"})
    with urllib.request.urlopen(req, timeout=timeout) as resp: data = json.loads(resp.read().decode("utf-8"))
    return rows_from_json(data.get("rows") if isinstance(data, dict) else data)


def sync_one(client, rows, today=None):
    """Feed path: the query was already specific to this client, so one FIR/case match is enough. None = no match."""
    best = None
    for r in rows:
        s, _ = _score(client, r)
        if s >= 80 and (best is None or s > best[0]): best = (s, r)
    return apply(client, best[1], today) if best else None


# ---- Daily Board / Daily Diary (the advocate's own board from the portal). Its date is the hearing date of every listed case.
def _board_date(txt):
    """Hearing date of the board: 'DAILY DIARY ON: dd-mm-yyyy', or 'Cause List / Board / Hearing Date: dd-mm-yyyy'. Never the 'Report Generated On' date."""
    for rx in (r"DIARY\s+ON:?", r"(?:CAUSE\s*LIST|BOARD|LISTING|HEARING)\s*(?:DATE|FOR|ON)?:?"):
        m = re.search(rx + r"\s*(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})", txt or "", re.I)
        if m:
            try: return date(int(m.group(3)), int(m.group(2)), int(m.group(1))).isoformat()
            except ValueError: pass
    return ""


# Other software names the same columns differently: map the common synonyms to the names the parser uses.
_H = [("case no.", r"case\s*(no|number|#|id)|suit\s*no|^cnr"), ("first party", r"first\s*party|petitioner|plaintiff|appellant|applicant|complainant"),
      ("opposite party", r"opposite\s*party|respondent|defendant|accused"), ("stage", r"^(stage|status|purpose|business)\b"),
      ("judge", r"^(judge|presiding)"), ("court", r"^court\b(?!\s*no)"), ("previous date", r"previous\s*date|last\s*date"),
      ("hearing date", r"next\s*(date|hearing)|hearing\s*date|date\s*of\s*hearing|fixed\s*for")]


def _canon(h):
    for k, rx in _H:
        if re.search(rx, h, re.I): return k
    return h


def _board_row(d, group, bdate):
    g = lambda k: d.get(k, "")
    cn, first, opp = g("case no.") or g("case no"), g("first party"), g("opposite party")
    court = g("judge") or g("court") or group
    pd = _dates(g("previous date")); ex = {"type": g("type"), "case_no": cn, "previous_date": pd[0] if pd else g("previous date"),
                                           "first_party": first, "opposite_party": opp, "us": g("u/s"), "ps_company": g("p.s/company"),
                                           "comments": g("comments"), "last_proceedings": g("last proceedings"), "board_date": bdate}
    return {"text": " | ".join(x for x in (g("type"), cn, court, first, opp, g("stage")) if x)[:400], "pairs": _pairs_of(cn), "date": (_dates(g("hearing date")) or [bdate])[0],
            "stage": g("stage")[:80], "court": court[:80], "ps": g("p.s/company")[:60], "title": f"{first} v. {opp}".strip(" v.")[:120],
            "tokens": _tok(f"{first} {opp}"), "extra": {k: v for k, v in ex.items() if v}}      # name matching uses parties only, never the judge


def _board_table(t, bdate):
    group, header, out = "", None, []
    for raw in t:
        cells = ["" if re.fullmatch(r"-+", x) else x for x in (re.sub(r"\s+", " ", c or "").strip() for c in raw)]       # '---' = empty
        if not any(cells): continue
        low = [_canon(c.lower()) for c in cells]
        if "case no." in low: header = low; continue
        if header is None: group = next((c for c in cells if c), group); continue          # court title row
        out.append(_board_row(dict(zip(header, cells)), group, bdate))
    return [r for r in out if r["pairs"]]


def board_rows(pdf_bytes):
    """Board PDF -> rows with real table cells (stage, judge, parties, comments stay separate). Needs pdfplumber.
    If no table with a case-number column is found, falls back to reading the text (case number + date only)."""
    try: import pdfplumber
    except ImportError: raise RuntimeError("Server par pdfplumber install nahi: requirements.txt mein 'pdfplumber' likh kar redeploy karein.")
    import io
    out, texts, bdate = [], [], ""
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for i, page in enumerate(pdf.pages[:60]):
            tx = page.extract_text() or ""; texts.append(tx)
            bdate = _board_date(tx) or bdate               # a range board has a new date heading on each day's pages
            for tb in page.extract_tables(): out += _board_table(tb, bdate)
    return out or board_text_rows("\n".join(texts))


_HDR = re.compile(r"Type\s+Case\s+No\.?\s+Court\s+No\.?\s+Judge\s+Court\s+Previous\s+Date\s+First\s+Party\s+Opposite\s+Party\s+Stage\s+U/S\s+P\.S/Company\s+Comments\s+Last\s*Proceedings", re.I)


def board_text_rows(raw):
    """Fallback when the board was copied as plain text: case number + hearing date only (stage cannot be separated reliably)."""
    t = _HDR.sub(" ", re.sub(r"\s+", " ", raw or "")); bdate = _board_date(t)
    hits = [(m.start(), m.end(), _pairs_of(m.group(0))) for m in re.finditer(r"(?<![\d/\-])\d{1,7}/\d{2,4}(?![\d/\-])|(?<![\d/\-])\d{5,10}(?![\d/\-])", t)]
    hits = [h for h in hits if h[2]]; out = []
    for i, (s, e, pairs) in enumerate(hits):
        seg = t[e:hits[i + 1][0] if i + 1 < len(hits) else len(t)]
        dm = re.search(r"\d{1,2}-\d{1,2}-\d{4}", seg)
        parties = seg[dm.end():] if dm else seg                       # drop the judge name that sits before the previous date
        pd = _dates(dm.group(0)) if dm else []
        out.append({"text": t[s:e] + " " + parties[:200], "pairs": pairs, "date": bdate, "stage": "", "court": "", "ps": "", "title": "",
                    "tokens": _tok(parties), "extra": {k: v for k, v in {"previous_date": pd[0] if pd else "", "board_date": bdate}.items() if v}})
    return out


def fetch_causelist(url, key, district, day, timeout=30):
    """Official bulk feed: every court of one district for one day in a single request. Same JSON rows as fetch_official."""
    q = urllib.parse.urlencode({"district": district, "date": day})
    req = urllib.request.Request(f"{url}?{q}", headers={"X-Api-Key": key, "Accept": "application/json"})
    with urllib.request.urlopen(req, timeout=timeout) as resp: data = json.loads(resp.read().decode("utf-8"))
    rows = rows_from_json(data.get("rows") if isinstance(data, dict) else data)
    for r in rows: r["date"] = r["date"] or day            # a cause list's date is the hearing date of its rows
    return rows


def sweep_causelist(data, fetch, days, today=None):
    """data: {file: [clients]}. Asks the feed once per district per day (all courts of that city), updates every advocate's
    matching clients, returns the set of files that changed. Clients without a district (field or city in 'court') are skipped."""
    districts = sorted({district_of(c) for cl in data.values() for c in cl if isinstance(c, dict)} - {""})
    changed = set()
    for d in districts:
        for day in sorted(days, reverse=True):            # earliest day is applied last, so the earliest hearing wins
            try: rows = fetch(d, day)
            except Exception as e:
                print(f"[court] cause list {d} {day}: {type(e).__name__}", flush=True); continue
            for name, cl in data.items():
                if ingest("", cl, rows=rows, today=today)["updated"]: changed.add(name)
    return changed


# ---- admin safety net: remember exactly what a bulk update changed, and put it back only where nobody edited since
def changed_entries(before, after):
    return [{"i": i, "before": b, "after": copy.deepcopy(a)} for i, (b, a) in enumerate(zip(before, after)) if b != a]


def undo_entries(clients, entries):
    done = skipped = 0
    for e in entries:
        i = e["i"]
        if 0 <= i < len(clients) and clients[i] == e["after"]: clients[i] = e["before"]; done += 1
        else: skipped += 1                       # edited by the advocate (or the list changed) after the sync: leave it alone
    return done, skipped
''')      # district court case sync

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
FAST_MODEL = os.environ.get("CLAUDE_FAST_MODEL", "claude-haiku-5-5")                      # cheap model for the light tasks
FAST_KINDS = {k.strip() for k in os.environ.get("FAST_KINDS", "category,section").split(",") if k.strip()}
# USD per million tokens (input, output) - only used for the admin cost report; change here if the prices change
PRICES = {"claude-haiku-5-5": (0.10, 0.50), "claude-sonnet-5-5": (2.0, 10.0), "claude-opus-5-5": (4.0, 20.0), "claude-sonnet-4-6": (3.0, 15.0)}


def model_for(kind):
    """MODEL_<KIND> (e.g. MODEL_DRAFT) wins; then the light kinds use the cheap model; everything else uses CLAUDE_MODEL."""
    return os.environ.get("MODEL_" + str(kind).upper()) or (FAST_MODEL if kind in FAST_KINDS else MODEL)

DATA = os.environ.get("DATA_DIR", "data")
os.makedirs(DATA, exist_ok=True)
USERS_FILE, PENDING_FILE = f"{DATA}/users.json", f"{DATA}/pending.json"
CAT_CACHE = f"{DATA}/category_cache.json"
OTP_VALID, RESEND, SESSION_DAYS = 600, 60, 90      # har istemal par session barhta rehta hai
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
    return {"status": "ok", "app": "Advocate AI API", "version": "v35-court-hub", "model": MODEL, "data_dir": DATA, "storage": "upstash" if REMOTE else "disk"}

# ---------------------------------------------------------------- helpers
# --- STORE BEGIN
import copy
UP_URL = os.environ.get("UPSTASH_REDIS_REST_URL", "").rstrip("/")
UP_TOKEN = os.environ.get("UPSTASH_REDIS_REST_TOKEN", "")
REMOTE = bool(UP_URL and UP_TOKEN)         # True: accounts/clients live in Upstash (permanent, free); False: local disk as before
_PFX, _CACHE = "advocate:", {}             # one server instance: reads are served from memory, every write goes to Upstash first
print("[storage]", "Upstash (permanent)" if REMOTE else f"local disk '{DATA}' (Render free par server sone/redeploy par mit jata hai)", flush=True)


def _up(cmd):
    """One Redis command over Upstash's REST API (plain HTTPS, so it also works on Render's free plan)."""
    req = urllib.request.Request(UP_URL, data=json.dumps(cmd).encode(), method="POST",
                                 headers={"Authorization": "Bearer " + UP_TOKEN, "Content-Type": "application/json"})
    last = None
    for _ in range(3):
        try:
            with urllib.request.urlopen(req, timeout=15) as r: d = json.loads(r.read().decode())
            if isinstance(d, dict) and d.get("error"): raise RuntimeError("Upstash: " + str(d["error"])[:120])
            return d.get("result")
        except urllib.error.HTTPError as e: raise RuntimeError(f"Upstash {e.code}: {e.read().decode('utf-8', 'ignore')[:120]}")
        except RuntimeError: raise
        except Exception as e: last = e; time.sleep(0.4)
    raise RuntimeError(f"Upstash se rabta nahi ho saka: {last}")


def _rk(path):
    if not REMOTE: return None
    p, d = os.path.abspath(path), os.path.abspath(DATA)
    return _PFX + p[len(d) + 1:].replace(os.sep, "/") if p.startswith(d + os.sep) else None


def _rj_file(path, default):
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default


def rj(path, default):
    k = _rk(path)
    if not k: return _rj_file(path, default)
    if k in _CACHE: return copy.deepcopy(_CACHE[k])
    raw = _up(["GET", k])                    # a failed request raises: never answer "empty" when the store is merely unreachable
    if raw is None:
        local = _rj_file(path, None)         # first start after switching to Upstash: take over what is still on the local disk
        if local is None: return default
        wj(path, local); return copy.deepcopy(local)
    obj = json.loads(raw); _CACHE[k] = obj
    return copy.deepcopy(obj)


_WLOCK = threading.Lock()
def wj(path, data):
    """Remote: written to Upstash. Local: temp file first, then rename, so a crash never leaves users.json empty."""
    k = _rk(path)
    with _WLOCK:
        if k:
            _up(["SET", k, json.dumps(data, ensure_ascii=False)]); _CACHE[k] = copy.deepcopy(data); return
        tmp = f"{path}.{secrets.token_hex(3)}.tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        os.replace(tmp, path)


def data_files():
    """Names of every stored file (Upstash keys plus anything still on the local disk)."""
    names = set(os.listdir(DATA)) if os.path.isdir(DATA) else set()
    if REMOTE: names |= {k[len(_PFX):] for k in (_up(["KEYS", _PFX + "*"]) or [])}
    return sorted(names)


def rm(path):
    k = _rk(path)
    if k:
        with _WLOCK: _up(["DEL", k]); _CACHE.pop(k, None)
    try: os.remove(path)
    except OSError: pass
# --- STORE END
cit.HOOK["rj"], cit.HOOK["wj"] = rj, wj          # the verified-citation library and its requests live in the same durable store

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
            "is_admin": bool(u.get("is_admin")), "disabled": bool(u.get("disabled"))}

# ---------------------------------------------------------------- auth dependency
def current(authorization, allow_expired=False):
    if not authorization: raise HTTPException(401, "Login zaroori hai")
    h = _h(authorization.replace("Bearer ", "").strip())
    users = load_users()
    for uid, u in users.items():
        ts = u.get("sessions", {}).get(h)
        if ts and time.time() - ts <= SESSION_DAYS * 86400:
            if u.get("disabled") and not u.get("is_admin"):
                raise HTTPException(403, "Aapka account admin ne band kar diya hai. Admin se rabta karein.")
            dirty = False
            if time.time() - ts > 86400:                  # roz ek dafa meaad barha do: active user kabhi logout nahi hoga
                u["sessions"][h] = time.time(); dirty = True
            if u.get("last_reset_date") != str(date.today()):
                u["prompt_count"], u["last_reset_date"] = 0, str(date.today()); dirty = True
            if dirty: wj(USERS_FILE, users)
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
        return False, f"Email send nahi ho saka: {e}. (Render free par Gmail SMTP band hota hai - Environment mein BREVO_API_KEY aur MAIL_FROM set karein.)"

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
class Signup(BaseModel): chamber: str; email: str; password: str; package: str = "basic"; name: str = ""; phone: str = ""; city: str = ""; bar_no: str = ""
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
    if u.get("disabled") and not u.get("is_admin"): raise HTTPException(403, "Aapka account admin ne band kar diya hai. Admin se rabta karein.")
    if not u.get("is_admin") and "$" not in u["password"]: u["password"] = hash_password(r.password)
    return {"token": new_session(users, uid), "user": profile(u)}

@app.post("/auth/signup")
def signup(r: Signup):
    email = r.email.strip().lower(); users = load_users()
    if not (r.chamber.strip() and EMAIL_RE.match(email)): raise HTTPException(400, "Naam aur sahi email likhein.")
    if len(r.password) < 8: raise HTTPException(400, "Password kam az kam 8 characters ka ho.")
    if r.package not in DAILY_LIMITS: raise HTTPException(400, "Ghalat package.")
    if not r.name.strip(): raise HTTPException(400, "Apna poora naam likhein.")
    phone = re.sub(r"[^0-9+]", "", r.phone)
    if len(re.sub(r"\D", "", phone)) < 10: raise HTTPException(400, "Sahi mobile / WhatsApp number likhein.")
    if not r.city.strip(): raise HTTPException(400, "Shehr ka naam likhein.")
    if find_account(users, email): raise HTTPException(400, "Is email se account pehle se hai.")
    ok, info = otp_send("signup", email, {"chamber": r.chamber.strip(), "pw_hash": hash_password(r.password), "pkg": r.package,
                                          "name": r.name.strip()[:80], "phone": phone[:20], "city": r.city.strip()[:40], "bar_no": r.bar_no.strip()[:40]})
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
                    "name": d.get("name", ""), "phone": d.get("phone", ""), "city": d.get("city", ""), "bar_no": d.get("bar_no", ""),
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

CMS_MIN_DIGITS = 5          # the long number of the court's case management system (e.g. 2348059/26); FIR / local numbers are shorter


def _cms(raw):
    """'NNNNNNN/YY' from NNNNNNN/YY, NNNNNNN/YYYY or joined digits; None when it does not look like a Case Management System number."""
    for n, y in sorted(cs._pairs_of(str(raw or ""))):
        if len(str(n)) >= CMS_MIN_DIGITS: return f"{n}/{str(y)[-2:]}"
    return None


def _adv_cases(k):
    rows = [c for c in rj(dpath(k, "clients"), []) if isinstance(c, dict)]
    cases = [{"case_no": _cms(c.get("case_no")) or str(c.get("case_no") or ""), "fir": str(c.get("fir") or ""), "court": str(c.get("court") or ""),
              "date": str(c.get("date") or ""), "stage": str(c.get("stage") or "")} for c in rows if c.get("case_no") or c.get("fir")]
    cases.sort(key=lambda x: (x["date"] or "9999", x["case_no"]))
    return cases, sum(1 for c in rows if not _cms(c.get("case_no"))), len(rows)


@app.get("/data/{col}")
def get_col(col: str, authorization: str = Header(None)):
    _, uid = current(authorization)
    if col not in COLS: raise HTTPException(404)
    return rj(dpath(uid, col), [])

@app.put("/data/{col}")
def put_col(col: str, rows: list[dict], authorization: str = Header(None)):
    _, uid = current(authorization)
    if col not in COLS: raise HTTPException(404)
    if col == "clients":
        key = lambda c: (str(c.get("name", "")), str(c.get("fir", "")), str(c.get("case_no", "")))
        known, seen = {key(c) for c in rj(dpath(uid, "clients"), []) if isinstance(c, dict)}, {}
        for c in rows:
            cn = _cms(c.get("case_no"))
            if cn: c["case_no"] = cn
            elif key(c) not in known:          # clients saved before this rule may stay as they are; every new one needs the number
                raise HTTPException(400, "Naye client ke liye Case Management System ka case number zaroori hai (jaise 2348059/26) - FIR ya court ka chhota number nahi.")
            if cn:
                if cn in seen: raise HTTPException(400, f"Case number {cn} pehle se maujood hai (client: {seen[cn]}).")
                seen[cn] = str(c.get("name", ""))
    wj(dpath(uid, col), rows[:5000]); return {"ok": True}

# ---------------------------------------------------------------- AI
SYSTEM = ("You are a Senior Advocate of the High Court and Supreme Court of Pakistan with 30 years of drafting experience. "
          "You excel at Pakistani law (Constitution, CPC, PPC, CrPC, QSO, Family, Revenue and special laws) and precedents. "
          "When asked to draft, produce ONLY the finished court-ready document: proper heading and title of the case, parties, "
          "numbered paragraphs, grounds, prayer, verification/affidavit and signature block where applicable. Use a formal, "
          "precise, professional pleading register at all times, never a chatty or simplified tone. Do not add introductions, "
          "notes or explanations outside the document and do not use markdown symbols (#, *, tables). Put unknown facts in "
          "square-bracket blanks such as [____]. Never cite a case, reporter citation or judgment from memory: cite case law only when it is supplied under VERIFIED AUTHORITIES, otherwise cite none.")

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

# --- FACTS BEGIN
CHECK_MARK = "\n\n=====JAANCH====="              # everything after this line is the server's checking report: it is not part of the document that is exported / copied
_DIG = {**{0x660 + i: 48 + i for i in range(10)}, **{0x6F0 + i: 48 + i for i in range(10)}}
_MONTHS = {m: i for i, m in enumerate(["january", "february", "march", "april", "may", "june", "july", "august", "september", "october", "november", "december"], 1)}
_ACTS = r"(?:PPC|P\.\s?P\.\s?C\.?|Pakistan\s+Penal\s+Code|Penal\s+Code|CrPC|Cr\.\s?P\.\s?C\.?|Code\s+of\s+Criminal\s+Procedure|CPC|C\.\s?P\.\s?C\.?|Code\s+of\s+Civil\s+Procedure|CNSA|CNS\s+Act|Control\s+of\s+Narcotic\s+Substances\s+Act|Qanun-e-Shahadat|QSO|Arms\s+Ordinance|Arms\s+Act|ATA|Anti-Terrorism\s+Act)"
_TOK = r"\d{1,4}(?:\s?-\s?[A-Za-z](?![A-Za-z])|[A-Za-z](?![A-Za-z]))?(?:\s?\(\s?[ivxlc\d]{1,4}\s?\))*"
_LIST = rf"{_TOK}(?:\s*(?:,|/|&|\band\b|\bor\b)\s*{_TOK})*"
_SEC_HEAD = re.compile(r"(?:\bu\s?/\s?ss?\b\.?|\bsections?\b|\bsec\.?|\bss\.|\bs\.(?=\s?\d)|دفعہ|دفعات|دفعه)\s*[:\-]?\s*", re.I)
_SEC_ACT = re.compile(rf"(?<![\w/.-])({_LIST})\s*(?:of\s+(?:the\s+)?)?({_ACTS})(?![A-Za-z])", re.I)
_ACT_AFTER = re.compile(rf"\s*(?:of\s+(?:the\s+)?)?({_ACTS})(?![A-Za-z])", re.I)
_LIST_RX = re.compile(_LIST)
_ART = re.compile(r"\b(?:article|art\.)\s*(\d{1,3}[A-Za-z]?)", re.I)
_ORDER = re.compile(r"\border\s+([ivxlc]+|\d{1,2})\b(?:\s*,?\s*rules?\s*(\d{1,3}(?:\s*(?:&|and|,)\s*\d{1,3})*))?", re.I)
_D_ISO = re.compile(r"\b(\d{4})\s?-\s?(\d{1,2})\s?-\s?(\d{1,2})\b")
_D_NUM = re.compile(r"\b(\d{1,2})\s?[-/.]\s?(\d{1,2})\s?[-/.]\s?(\d{4}|\d{2})\b")
_D_TXT = re.compile(r"\b(\d{1,2})(?:st|nd|rd|th)?\s+(?:of\s+)?([A-Za-z]{3,9})\.?,?\s+(\d{4})\b")
_D_TXT2 = re.compile(r"\b([A-Za-z]{3,9})\.?\s+(\d{1,2})(?:st|nd|rd|th)?,?\s+(\d{4})\b")


def _d(s): return str(s or "").translate(_DIG)


def _iso(y, m, d):
    try: return date(int(y), int(m), int(d)).isoformat()
    except ValueError: return None


def _mon(name):
    n = name.lower()
    return next((v for k, v in _MONTHS.items() if k == n or (len(n) >= 3 and k.startswith(n[:3]) and len(n) <= len(k))), None)


def _num_date(a, b, y):
    month, day = (b, a) if int(b) <= 12 else (a, b)          # dd-mm-yyyy; if the middle number cannot be a month it is mm-dd-yyyy
    return _iso(y if len(y) == 4 else "20" + y, month, day)


def find_dates(text):
    """(set of ISO dates, text with the dates blanked out). Reads 17-03-2021, 17/3/21, 2021-03-17, 17 March 2021, March 17, 2021."""
    found, t = set(), str(text)
    for rx, fn in ((_D_ISO, lambda m: _iso(m[1], m[2], m[3])), (_D_NUM, lambda m: _num_date(m[1], m[2], m[3])),
                   (_D_TXT, lambda m: _iso(m[3], _mon(m[2]) or 0, m[1])), (_D_TXT2, lambda m: _iso(m[3], _mon(m[1]) or 0, m[2]))):
        for m in rx.finditer(t):
            iso = fn(m)
            if iso: found.add(iso)
        t = rx.sub(" ", t)
    return found, t


def _skey(tok): return re.sub(r"[^0-9a-z]", "", tok.lower())


def _act(a):
    a = re.sub(r"[\s.]", "", a).lower()
    return ("PPC" if a in ("ppc", "pakistanpenalcode", "penalcode") else "CrPC" if a in ("crpc", "codeofcriminalprocedure") else "CPC" if a in ("cpc", "codeofcivilprocedure")
            else "CNSA" if a in ("cnsa", "cnsact", "controlofnarcoticsubstancesact") else "QSO" if a in ("qanun-e-shahadat", "qso") else "ARMS" if a.startswith("arms") else "ATA" if a in ("ata", "anti-terrorismact") else a.upper())


def find_sections(text):
    """([{'sec': '337-F(i)', 'act': 'PPC' or '', 'key': '337fi'}], text with them blanked out) from 'u/s 302/34 PPC', 'sections 324, 337-F(i)', '302 PPC', 'دفعہ 302' ..."""
    t, out, spans = str(text), [], []
    def add(listtxt, act, a, b):
        for tok in re.findall(_TOK, listtxt):
            tok = re.sub(r"\s+", "", tok)
            if tok and tok.strip("-"): out.append({"sec": tok, "act": _act(act) if act else "", "key": _skey(tok)})
        spans.append((a, b))
    for m in _SEC_ACT.finditer(t): add(m.group(1), m.group(2), m.start(1), m.end())
    for h in _SEC_HEAD.finditer(t):
        lm = _LIST_RX.match(t, h.end())
        if not lm or any(a <= lm.start() < b for a, b in spans): continue
        am = _ACT_AFTER.match(t, lm.end())
        add(lm.group(0), am.group(1) if am else "", h.start(), am.end() if am else lm.end())
    for a, b in sorted(spans, reverse=True): t = t[:a] + " " + t[b:]
    seen, uniq = set(), []
    for x in out:
        if (x["act"], x["key"]) not in seen: seen.add((x["act"], x["key"])); uniq.append(x)
    return uniq, t


def find_refs(text):
    """Articles of the Constitution and CPC orders/rules: legal references like sections. -> ({key: shown}, text without them)"""
    t, refs = str(text), {}
    for m in _ART.finditer(t): refs["art" + _skey(m.group(1))] = "Article " + m.group(1)
    for m in _ORDER.finditer(t):
        r = re.sub(r"\s+", "", m.group(2) or "")
        refs[_skey("o" + m.group(1) + "r" + r)] = "Order " + m.group(1).upper() + ((" Rule " + m.group(2).strip()) if m.group(2) else "")
    return refs, _ORDER.sub(" ", _ART.sub(" ", t))


def fact_items(text):
    t = _d(text)
    dates, t = find_dates(t)
    refs, t = find_refs(t)            # before sections, so "Order XXXIX Rule 1 CPC" is not read as section 1
    secs, t = find_sections(t)
    nums = set()
    for n in re.findall(r"\b\d[\d,]*\b", t):
        n = n.replace(",", "").lstrip("0") or "0"
        if len(n) >= 3 and not (len(n) == 4 and 1947 <= int(n) <= 2100): nums.add(n)
    return {"dates": dates, "numbers": nums, "sections": {x["key"]: ("S." + x["sec"] + (" " + x["act"] if x["act"] else "")) for x in secs}, "refs": refs}


def guard_report(output, source, facts=True):
    """What the AI wrote vs what the advocate supplied. Dates and numbers that were not supplied are red; sections / articles the AI added on its own are yellow."""
    o, s = fact_items(output), fact_items(source)
    new_prov = {**{k: v for k, v in o["sections"].items() if k not in s["sections"]}, **{k: v for k, v in o["refs"].items() if k not in s["refs"]}}
    rep = {"new_provisions": sorted(new_prov.values()), "facts_checked": bool(facts)}
    if facts:
        rep.update(new_dates=sorted(o["dates"] - s["dates"]), new_numbers=sorted(o["numbers"] - s["numbers"]),
                   missing_dates=sorted(s["dates"] - o["dates"]), missing_numbers=sorted(s["numbers"] - o["numbers"]),
                   missing_provisions=sorted(s["sections"][k] for k in s["sections"] if k not in o["sections"]))
    rep["red"] = len(rep.get("new_dates", [])) + len(rep.get("new_numbers", []))
    rep["yellow"] = len(rep["new_provisions"])
    rep["missing"] = len(rep.get("missing_dates", [])) + len(rep.get("missing_numbers", [])) + len(rep.get("missing_provisions", []))
    return rep


def guard_text(rep):
    L = [CHECK_MARK, "(Ye jaanch server ne ki hai, AI ne nahi. Ye hissa Word / copy / share mein shamil nahi hota.)"]
    if rep["facts_checked"]:
        if rep["new_dates"]: L.append("❌ Ye tareekhen jawab mein hain magar aap ke diye hue input mein nahi thin: " + ", ".join(rep["new_dates"]))
        if rep["new_numbers"]: L.append("❌ Ye raqam / number jawab mein hain magar aap ke input mein nahi thay: " + ", ".join(rep["new_numbers"]))
        if not rep["red"]: L.append("✅ Jawab mein koi nayi tareekh ya raqam nahi jo aap ke input mein na ho.")
    else: L.append("ℹ Tasveer wale kaghazat se liye gaye facts ki jaanch nahi ho sakti - tareekhein aur raqam khud mila lein.")
    if rep["new_provisions"]: L.append("⚠ Ye dafaat / qawanin AI ne khud likhe (input mein nahi thay) - apne bare act se verify karein: " + ", ".join(rep["new_provisions"]))
    miss = rep.get("missing_dates", []) + rep.get("missing_numbers", []) + rep.get("missing_provisions", [])
    if miss: L.append("ℹ Aap ke input ki ye cheezein jawab mein nazar nahi aayin: " + ", ".join(miss))
    return "\n".join(L)


# ---- Dafa table: punishment / bailable / cognizable ... come ONLY from entries the admin team verified, never from the AI's memory
DAFA_JSON, DAFA_REQ = f"{DATA}/dafa_table.json", f"{DATA}/dafa_requests.json"
DAFA_FIELDS = ("act", "section", "title", "punishment", "bailable", "cognizable", "compoundable", "court", "notes", "source")


def dafa_rows(): return rj(DAFA_JSON, [])


def dafa_add(rec, by=""):
    r = {k: str(rec.get(k, "")).strip()[:600] for k in DAFA_FIELDS}
    if not (r["act"] and r["section"] and r["title"]): return False, "Act, section aur dafa ka naam zaroori hain."
    if not r["source"]: return False, "Zareya (source) likhein - ye maloomat kahan se tasdeeq hui (jaise Pakistan Code website)."
    if not (r["punishment"] or r["bailable"] or r["cognizable"]): return False, "Saza, bailable ya cognizable mein se kam az kam ek likhein."
    key = (_act(r["act"]), _skey(r["section"])); rows = dafa_rows()
    if any((_act(x["act"]), _skey(x["section"])) == key for x in rows): return False, "Ye dafa pehle se table mein hai."
    if len(rows) >= 3000: return False, "Table full hai."
    r.update(act=_act(r["act"]), verified_by=by, verified_on=date.today().isoformat())
    wj(DAFA_JSON, rows + [r]); _resolve_dafa([key]); return True, ""


def dafa_import(text, by=""):
    import csv, io as _io
    added = skipped = 0
    for row in csv.DictReader(_io.StringIO(text or "")):
        ok_, _ = dafa_add({k: (row.get(k) or "") for k in DAFA_FIELDS}, by)
        added, skipped = added + ok_, skipped + (not ok_)
    return {"added": added, "skipped": skipped}


def dafa_lookup(items):
    rows, found, missing = dafa_rows(), [], []
    for it in items:
        hit = [r for r in rows if _skey(r["section"]) == it["key"] and (not it["act"] or _act(r["act"]) == it["act"])]
        if hit: found += [r for r in hit if r not in found]
        else: missing.append(it)
    return found, missing


def _resolve_dafa(keys):
    rq = rj(DAFA_REQ, {}); n = len(rq)
    for _a, k in keys:
        for x in [x for x in rq if x.split("|", 1)[1] == k]: rq.pop(x)
    if len(rq) != n: wj(DAFA_REQ, rq)


def dafa_requests():
    return sorted(({"section": v["section"], "act": v["act"], "count": v["count"]} for v in rj(DAFA_REQ, {}).values()), key=lambda x: -x["count"])


def _dafa_note(missing):
    if not missing: return
    rq = rj(DAFA_REQ, {})
    for it in missing:
        k = f"{it['act']}|{it['key']}"
        if k in rq: rq[k]["count"] += 1
        elif len(rq) < 300: rq[k] = {"section": it["sec"], "act": it["act"], "count": 1}
    wj(DAFA_REQ, rq)


def _dafa_line(r):
    return (f"{r['act']} {r['section']} - {r['title']}: saza: {r['punishment'] or '-'} | bailable: {r['bailable'] or '-'} | cognizable: {r['cognizable'] or '-'} | "
            f"compoundable: {r['compoundable'] or '-'} | court: {r['court'] or '-'} (zareya: {r['source']}; tasdeeq: {r.get('verified_by', '')} {r.get('verified_on', '')})")


def dafa_block(text, note=True):
    """Table details for every section mentioned in the text. Sections that are not in the table are listed and queued for the admin team."""
    items, _ = find_sections(_d(text))
    if not items: return ""
    found, missing = dafa_lookup(items)
    if note: _dafa_note(missing)
    out = "\n\nDAFA TABLE (admin ne tasdeeq kiye hue - AI ki yaadasht nahi):"
    out += ("\n" + "\n".join(_dafa_line(r) for r in found)) if found else "\n(In dafaat ki tafseel abhi table mein nahi.)"
    if missing: out += "\nTable mein nahi (admin ko bhej diya gaya, tab tak apne bare act se verify karein): " + ", ".join((x["act"] + " " if x["act"] else "") + "S." + x["sec"] for x in missing)
    return out


def fir_facts_prompt(text):
    """Facts read from the FIR by plain code (no AI) + the verified table rows, handed to the AI as the only allowed source for them."""
    t = _d(text); dates, _ = find_dates(t); items, _ = find_sections(t)
    no = re.search(r"F\.?\s?I\.?\s?R\.?\s*(?:No\.?|#|number)?\s*[:\-]?\s*(\d{1,5}\s?/\s?\d{2,4})", t, re.I)
    ps = re.search(r"(?:P\.?\s?S\.?|Police\s+Station|تھانہ)\s*[:\-]?\s*([A-Za-z][A-Za-z ]{2,28}|[\u0600-\u06FF ]{3,25})", t)
    found, missing = dafa_lookup(items)
    fno, fps, fd = (re.sub(r"\s", "", no.group(1)) if no else "not found"), (ps.group(1).strip() if ps else "not found"), (", ".join(sorted(dates)) or "none")
    s = "FIR FACTS (read by code, exact - use them as written; if something is not listed here and not in the FIR text, do not invent it):\n"
    s += f"- FIR no.: {fno}\n- Police station: {fps}\n- Dates found: {fd}\n"
    s += "- Sections found: " + (", ".join((x['act'] + " " if x['act'] else "") + "S." + x['sec'] for x in items) or "none") + "\n"
    s += "DAFA TABLE (verified by the admin team - the ONLY source you may use for punishment, bailable / non-bailable, cognizable, compoundable and court of trial):\n"
    s += ("\n".join(_dafa_line(r) for r in found) if found else "(no rows)") + "\n"
    s += ("Sections NOT in the table: " + ", ".join("S." + x['sec'] for x in missing) + " - for these write 'dafa table mein nahi - verify karein' and do NOT state punishment or bailability.\n") if missing else ""
    return s
# --- FACTS END


class AiReq(BaseModel):
    kind: str                      # draft | research | citation | fir
    language: str = "English"
    fields: dict = {}
    image_b64: str | None = None   # purana (ek tasveer)
    images_b64: list[str] | None = None   # nayi: kai tasveerein (JPEG)
    id_images_b64: list[str] | None = None   # parties ke shanakhti card (CNIC) ki tasveerein
    template_id: str | None = None        # saved template (server par)

def build_prompt(r: AiReq, has_template=False):
    f, L = r.fields, (TEMPLATE_INSTR if has_template else lang_instr(r.language)); blank = "(not provided - leave a clear blank placeholder)"; LP = lang_plain(r.language)
    g = lambda k: (f.get(k) or "").strip()
    if r.kind == "draft":
        if not (g("facts") or g("parties") or g("extra") or g("ref") or r.images_b64 or r.image_b64 or r.id_images_b64):
            raise HTTPException(400, "Kam az kam parties, facts ya document/ID ki photo dein.")
        nofacts = "" if g("facts") else ("\nNo facts were supplied: prepare the complete document with the standard recitals and grounds that are normal for this type of document, "
                                        "and leave [____] blanks for case-specific facts, dates and amounts. Never invent facts.")
        idnote = "\nParty details (full name, father's/husband's name, CNIC number, address) must be taken from the attached ID card images where provided; copy them exactly and do not guess unreadable characters." if r.id_images_b64 else ""
        return f"""{L}
Follow standard Pakistani court drafting style and terminology.
Prepare a complete, court-ready {g('category')} document: {g('doc_type')}
- Court/Forum: {g('court') or blank}
- Parties: {g('parties') or blank}
- Reference (case / suit no., FIR no., law provisions): {g('ref') or blank}
- Other details: {g('extra') or blank}
- Facts & Grounds: {g('facts') or blank}{nofacts}{idnote}
FACT RULE: every date, amount, number, name, address and place in the document must come from the details above. If one is not given, write a blank ________ instead of guessing. Do not state punishments or bailability from memory.
Output only the finished document (no preface, notes or markdown symbols).""", f"Draft - {g('doc_type')}"
    if r.kind == "research":
        if not (g("section") or g("question")): raise HTTPException(400, "Section ya sawal likhein.")
        return f"""{LP}
Conduct legal research on Pakistani law: {g('act')}
- Provision: {g('section')}
- Precedents from: {g('court')}
- Question/Facts: {g('question')}
Give the bare-act position and key principles. For case law use ONLY the VERIFIED AUTHORITIES supplied, referring to them by their [A#] tags; if none apply, say the verified library has nothing on this point.""", f"Research - {g('act')}"
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
weak points, bail grounds and next steps. Use a formal professional tone. Use the FIR FACTS below exactly as written; take punishment, bailable / non-bailable, cognizable, compoundable and court of trial ONLY from the DAFA TABLE below, never from memory. End with one final line starting with 'Recommended draft:' naming the single most suitable petition or application to file now (for example pre-arrest bail, post-arrest bail, quashing of FIR).
{fir_facts_prompt(g('text'))}
FIR TEXT:
{g('text')}""", "FIR Analysis"
    if r.kind == "category":
        cid = g("category")
        if cid not in CATEGORIES: raise HTTPException(400, "Ghalat category.")
        return f"""{LP}
Topic: {CATEGORIES[cid]}
List the sections of Pakistani law that are commonly applied in cases of this kind, so that an advocate can see at a glance which sections apply. Group them as: A) Pakistan Penal Code sections, B) special-law sections (if any), C) Code of Criminal Procedure (or other procedural) sections commonly used in such cases (for example FIR, arrest, bail, quashing), D) sections that are commonly added together (such as 34, 109, 148/149, 120-B) when relevant.
For every section give ONE line in this form, without tables or markdown: 'Section <number> <Act> - <name of offence>'. Do NOT state punishment, bailable / non-bailable, cognizable, compoundable or court of trial: the system adds those details from the admin-verified DAFA TABLE after your list. Order the sections logically (most serious first).
Be accurate. If you are not sure that a section applies write 'verify' instead of guessing, and never invent section numbers. Finish with one line reminding the advocate to verify the current text on the official Pakistan Code website.""", f"Dafaat - {cid}"
    if r.kind == "section":
        law = g("law") or "all relevant Pakistani laws"
        if g("mode") == "topic":
            if not g("topic"): raise HTTPException(400, "Jurm ya masla likhein.")
            return f"""{LP}
An advocate describes this offence or legal issue: {g('topic')}
Search within: {law}
Identify the most relevant sections or provisions of Pakistani law that apply: for each give the Act, the section number and a one-line explanation, ranked by relevance. Do NOT state punishment, bailable / non-bailable, cognizable, compoundable or court of trial: the system adds them from the admin-verified DAFA TABLE after your answer. Mention alternative or companion sections that are often invoked together, then add brief practical notes for the advocate.
Be explicit about anything you are not certain of and never invent section numbers.""", "Section search"
        if not g("section"): raise HTTPException(400, "Section number likhein.")
        q = ("Specific question: " + g("question")) if g("question") else ""
        return f"""{LP}
Explain this provision of Pakistani law for an advocate. Law: {law}. Provision: {g('section')}. {q}
Use these headings: 1) Name of the law and the section title; 2) Text or substance of the provision (if you cannot recall the exact wording, give the substance and say so); 3) Essential ingredients; 4) Relief or consequence in general terms - do NOT state punishment, bailable / non-bailable, cognizable, compoundable or court of trial: the system adds them from the admin-verified DAFA TABLE; 5) Important amendments you know of; 6) Related and alternative sections; 7) Leading case law (ONLY from the VERIFIED AUTHORITIES supplied; if none, write 'none in verified database'); 8) Practical points for bail, quashing or defence.
Statutes are amended often, so state clearly that the advocate must verify the current text on the official Pakistan Code website before relying on it. Never invent section numbers, text or citations.""", (f"Dafa - {g('section')} {g('law')}")[:80]
    raise HTTPException(400, "Unknown kind")

def call_claude(content, progress=None, model=None, usage=None, system=None):
    """Streaming: a long draft keeps flowing instead of dying at the 110 s read timeout. Overall limit 420 s.
    First thinking off (fast), then the other way, then the older model. progress(n) gets the number of characters written so far."""
    client = anthropic.Anthropic(api_key=API_KEY, timeout=110.0, max_retries=0)
    first = model or MODEL
    models = [first] + (["claude-sonnet-4-6"] if first != "claude-sonnet-4-6" else [])
    deadline = time.time() + 420
    last = None
    for m in models:
        for extra, max_tokens in (({"thinking": {"type": "disabled"}}, 8000), ({}, 16000)):
            if time.time() > deadline:
                raise last or RuntimeError("AI ko bohat der lag gayi")
            try:
                n = 0
                with client.messages.stream(model=m, max_tokens=max_tokens, system=system or SYSTEM, extra_body=extra,
                                            messages=[{"role": "user", "content": content}]) as stream:
                    for piece in stream.text_stream:
                        n += len(piece)
                        if progress: progress(n)
                    res = stream.get_final_message()
            except Exception as e:
                last = e
                print(f"[AI] {m} fail: {type(e).__name__}: {str(e)[:200]}", flush=True)
                continue
            text = "".join(b.text for b in res.content if getattr(b, "type", "") == "text").strip()
            if usage and getattr(res, "usage", None):
                try: usage(m, int(res.usage.input_tokens or 0), int(res.usage.output_tokens or 0))
                except Exception: pass
            if text:
                return text
            last = RuntimeError(f"AI ne khali jawab diya (stop_reason: {getattr(res, 'stop_reason', '?')})")
    raise last or RuntimeError("AI se koi jawab nahi mila")

# Lambay jawab ke liye job system: /ai foran job id deta hai, app /ai/job/{id} poochti rehti hai
JOBS, JOBS_LOCK = {}, threading.Lock()

_ULOCK = threading.Lock()


def _note_usage(kind, model, tin, tout):
    """Tokens only (no names, no text): month -> 'kind|model' -> calls / input / output."""
    try:
        with _ULOCK:
            u, mo = rj(f"{DATA}/ai_usage.json", {}), date.today().strftime("%Y-%m")
            e = u.setdefault(mo, {}).setdefault(f"{kind}|{model}", {"n": 0, "in": 0, "out": 0})
            e["n"] += 1; e["in"] += tin; e["out"] += tout
            wj(f"{DATA}/ai_usage.json", u)
    except Exception:
        traceback.print_exc()


def finalize_text(r, text, uid, note=True):
    """Everything the SERVER adds after the AI has answered: verified citations, Dafa-table details, the facts check.
    Returns (text, citations, raw AI text, stats). The raw text is what gets cached (the table may change later)."""
    raw, cites, stats, f = text, [], {"red": 0, "yellow": 0, "missing": 0}, (r.fields or {})
    path = dpath(uid, "mycitations")
    if r.kind in ("draft", "research", "section"):       # the AI only uses [A#] tags; the server prints the verified citations and their details
        q = " ".join(v for v in f.values() if isinstance(v, str))[:600]
        text, cites = cit.render(text, cit.authorities(q, path), path, q)
    elif r.kind == "fir":                                # flag any citation that is not in the verified index
        text, cites = cit.scrub(text, path)
    if r.kind in ("fir", "section", "category"):         # punishment / bailable / cognizable ... only from the admin-verified table
        extra = f" Section {f.get('section', '')} {f.get('law', '')}" if r.kind == "section" and f.get("section") else ""
        text += dafa_block(raw + extra, note)
    if r.kind in ("draft", "fir"):                       # dates, amounts and numbers must come from what the advocate supplied
        has_img = bool(r.images_b64 or r.image_b64 or r.id_images_b64)
        rep = guard_report(raw, " ".join(str(v) for v in f.values() if isinstance(v, str)), facts=not has_img)
        text += guard_text(rep); stats = {k: rep[k] for k in ("red", "yellow", "missing")}
    return text, cites, raw, stats


def _run_job(jid, uid, r, content, label, lim, ckey):
    def prog(n):
        with JOBS_LOCK:
            if jid in JOBS: JOBS[jid]["chars"] = n
    try:
        text = call_claude(content, prog, model_for(r.kind), lambda m, i, o: _note_usage(r.kind, m, i, o))
    except Exception as e:
        with JOBS_LOCK: JOBS[jid].update(status="error", code=502, detail=f"AI se jawab nahi aa saka ({type(e).__name__}): {str(e)[:220]}")
        return
    text, cites, raw, _stats = finalize_text(r, text, uid)
    used = 0
    try:
        if r.kind == "category":
            cache = rj(CAT_CACHE, {}); cache[ckey] = {"text": raw, "ts": time.time()}; wj(CAT_CACHE, cache)
        users = load_users(); u = users.get(uid)       # taaza load: dusri requests ke count overwrite na hon
        if u is not None:
            u["prompt_count"] = u.get("prompt_count", 0) + 1; used = u["prompt_count"]; wj(USERS_FILE, users)
        hist = rj(dpath(uid, "history"), [])
        hist.append({"time": datetime.now().strftime("%d-%m-%Y %H:%M"), "type": label, "result": text, "lang": r.language})
        keep, size = [], 0                                   # newest first, at most ~1.5 MB: every save rewrites the whole list
        for h in reversed(hist[-300:]):
            size += len(str(h.get("result", ""))) + 200
            if size > 1_500_000: break
            keep.append(h)
        wj(dpath(uid, "history"), keep[::-1])
    except Exception:
        traceback.print_exc()          # save fail ho bhi jaye to jawab user ko zaroor milay
    with JOBS_LOCK: JOBS[jid].update(status="done", data={"result": text, "type": label, "used": used, "limit": lim, "citations": cites})

@app.post("/ai")
def ai(r: AiReq, authorization: str = Header(None)):
    users, uid = current(authorization); u = users[uid]
    lim = DAILY_LIMITS.get(u.get("package"))
    ckey = f"{(r.fields or {}).get('category', '')}|{r.language}|v2"
    if r.kind == "citation":      # verified-index lookup only: the AI is never asked if a case is real
        txt = (r.fields or {}).get("citation", "")
        if not txt.strip(): raise HTTPException(400, "Citation likhein.")
        v = cit.verify(txt, dpath(uid, "mycitations"), r.language)
        for it in v["items"]:
            if it["status"] == "unindexed": cit.request_cite(it["citation"])
        return {"result": v["report"], "type": "Citation Check", "used": u.get("prompt_count", 0), "limit": lim, "citations": v["items"]}
    if r.kind == "category":      # ek hi list sab ke liye: cache mein ho to free aur foran
        hit = rj(CAT_CACHE, {}).get(ckey)
        if hit and time.time() - hit.get("ts", 0) < 30 * 86400:
            return {"result": hit["text"] + dafa_block(hit["text"]), "type": f"Sections - {ckey.split('|')[0]}", "used": u.get("prompt_count", 0), "limit": lim, "cached": True}
    if not u.get("is_admin") and lim is not None and u.get("prompt_count", 0) >= lim:
        raise HTTPException(429, f"Aaj ki limit ({lim} prompts) mukammal. Kal reset hogi ya package upgrade karein.")
    if not API_KEY: raise HTTPException(500, "AI service configure nahi hai.")
    tmpl = []
    if r.template_id:
        for t in rj(dpath(uid, "templates"), []):
            if t.get("id") == r.template_id: tmpl = t.get("images", [])
        if not tmpl: raise HTTPException(404, "Ye template nahi mila.")
    prompt, label = build_prompt(r, bool(tmpl))
    if r.kind in ("draft", "research", "section"):      # only verified authorities may be cited
        q = " ".join(v for v in (r.fields or {}).values() if isinstance(v, str))[:600]
        prompt = cit.authorities_block(q, dpath(uid, "mycitations")) + "\n\n" + prompt
    imgs = [i for i in (list(r.images_b64 or []) + ([r.image_b64] if r.image_b64 else [])) if i]
    if len(imgs) > 5: raise HTTPException(400, "Ek waqt mein ziyada se ziyada 5 tasveerein dein.")
    ids = [i for i in (r.id_images_b64 or []) if i]
    if len(ids) > 4: raise HTTPException(400, "ID card ki ziyada se ziyada 4 tasveerein dein.")
    if any(len(i) > 2_000_000 for i in imgs + ids): raise HTTPException(400, "Koi tasveer bohat bari hai.")
    img_block = lambda i: {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg", "data": i}}
    blocks = []
    if tmpl:
        blocks.append({"type": "text", "text": "TEMPLATE (advocate's own sample draft - copy its layout and language):"})
        blocks += [img_block(i) for i in tmpl]
    if imgs:
        blocks.append({"type": "text", "text": "CASE DOCUMENTS (FIR / order photos - use for facts only):"})
        blocks += [img_block(i) for i in imgs]
    if ids:
        blocks.append({"type": "text", "text": "ID CARDS (CNIC of the parties - read names, father/husband name, CNIC number and address for the party blocks):"})
        blocks += [img_block(i) for i in ids]
    content = (blocks + [{"type": "text", "text": prompt}]) if blocks else prompt
    jid = secrets.token_hex(8); now = time.time()
    with JOBS_LOCK:
        for k in [k for k, v in JOBS.items() if now - v["ts"] > 3600]: del JOBS[k]
        JOBS[jid] = {"uid": uid, "status": "pending", "ts": now}
    threading.Thread(target=_run_job, args=(jid, uid, r, content, label, lim, ckey), daemon=True).start()
    return {"job": jid}

@app.get("/ai/job/{jid}")
def ai_job(jid: str, authorization: str = Header(None)):
    _, uid = current(authorization, True)
    with JOBS_LOCK: j = JOBS.get(jid)
    if not j or j["uid"] != uid: raise HTTPException(404, "Ye request nahi mili (server restart ho gaya hoga). Dobara try karein.")
    if j["status"] == "pending": return {"status": "pending", "chars": j.get("chars", 0), "waited": int(time.time() - j["ts"])}
    with JOBS_LOCK: JOBS.pop(jid, None)
    if j["status"] == "error": raise HTTPException(j["code"], j["detail"])
    return {"status": "done", **j["data"]}

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
    for line in r.text.split(CHECK_MARK.strip())[0].split("\n"):
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

# ---------------------------------------------------------------- verified citations
class CitVerify(BaseModel): text: str; language: str = "English"
class CitIn(BaseModel): citation: str; title: str; court: str = ""; summary: str = ""; verified_from: str = ""
class CitImport(BaseModel): text: str; source: str = ""

@app.post("/citations/verify")
def citations_verify(r: CitVerify, authorization: str = Header(None)):
    _, uid = current(authorization)
    v = cit.verify(r.text, dpath(uid, "mycitations"), r.language)
    for it in v["items"]:
        if it["status"] == "unindexed": cit.request_cite(it["citation"])
    return v

@app.post("/citations/mine")
def citations_mine(r: CitIn, authorization: str = Header(None)):
    """Advocate confirmed it himself (e.g. on Pakistan Law Site): saved to his own library, labelled as his confirmation."""
    users, uid = current(authorization)
    ok, msg = cit.add_user(dpath(uid, "mycitations"), {"citation": r.citation, "title": r.title, "court": r.court, "summary": r.summary,
                                                      "source": f"user-verified ({r.verified_from.strip() or 'source not stated'})",
                                                      "verified_by": users[uid]["chamber_name"]})
    if not ok: raise HTTPException(400, msg)
    return {"ok": True}

# --- ASSISTANT BEGIN
ASSIST_SYSTEM = (
    "You are 'Law Assistant', a chat helper for Pakistani advocates inside the Advocate AI app. Answer questions about Pakistani law "
    "(criminal, civil, family, property, procedure, evidence, limitation, courts). If a question is not about law, say in one line that you only help with legal questions. "
    "STYLE: 'through the point' - short numbered points, no long introductions, no repeated disclaimers. When the user asks what the law is on something, or how to do something "
    "(procedure), answer in this order and skip any part that does not apply: 1) QANOON: the Act and section in one or two lines. 2) ZAROORI SHARAIT: what must exist or be proved. "
    "3) TAREEKA: numbered steps - which court or forum, which documents are filed, what happens next, and the time limit only if you are sure. 4) EHTIYAT: one caution line. "
    "Keep the whole answer to about 6 to 10 lines; give more detail only if the user asks. "
    "Reply in the language given below (Urdu script, Roman Urdu or English), matching the user. "
    "ACCURACY: name an Act and section only when you are sure; if unsure write 'section number apne bare act se verify karein'. Never invent a section, a time limit or a court. "
    "State when the answer depends on facts, the forum or a recent amendment. "
    "CASE LAW: never write a case name or law-report citation (PLD, SCMR, CLC, YLR, PCrLJ ...) yourself. Case law may be used ONLY through the numbered VERIFIED AUTHORITIES below, by their tag, e.g. [A1]. ")
ASSIST_HITS, _ALOCK = [], threading.Lock()
_PII = [(re.compile(r"\b\d{5}-?\d{7}-?\d\b"), "[CNIC]"), (re.compile(r"(?:\+92|0092|92|0)\s?3\d{2}[\s-]?\d{7}\b"), "[phone]"),
        (re.compile(r"\b0\d{2,3}[\s-]?\d{6,8}\b"), "[phone]"), (re.compile(r"[\w.+-]+@[\w-]+\.[\w.-]+"), "[email]"), (re.compile(r"\b\d{10,}\b"), "[number]")]


def _redact(t):
    """Free AI services may keep what they receive: CNIC, phone numbers, e-mails and long numbers never leave the server."""
    for rx, rp in _PII: t = rx.sub(rp, t)
    return t


def _assist_providers():
    out = []
    for p in ("ASSIST", "ASSIST2"):
        base, key, model = os.environ.get(f"{p}_BASE_URL", "").rstrip("/"), os.environ.get(f"{p}_API_KEY", ""), os.environ.get(f"{p}_MODEL", "")
        if p == "ASSIST" and key:
            base = base or "https://generativelanguage.googleapis.com/v1beta/openai"       # Google's OpenAI-compatible endpoint
            model = model or "gemini-3.5-flash-lite"
        if base and key and model: out.append((base, key, model))
    return out


def _clean_msgs(msgs):
    out = []
    for m in (msgs or [])[-12:]:
        if isinstance(m, dict) and m.get("role") in ("user", "assistant") and str(m.get("content", "")).strip():
            out.append({"role": m["role"], "content": _redact(str(m["content"])[:4000].strip())})
    while out and out[0]["role"] != "user": out.pop(0)
    return out


def _assist_call(base, key, model, system, msgs, max_tokens=900):
    body = json.dumps({"model": model, "messages": [{"role": "system", "content": system}] + msgs, "max_tokens": max_tokens, "temperature": 0.3}).encode()
    req = urllib.request.Request(base + "/chat/completions", data=body, method="POST", headers={"Authorization": "Bearer " + key, "Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=45) as resp: d = json.loads(resp.read().decode())
    return re.sub(r"<think>.*?</think>", "", str(d["choices"][0]["message"]["content"] or ""), flags=re.S).strip()


class AssistIn(BaseModel): messages: list; language: str = "Roman Urdu"

@app.post("/assistant/chat")
def assistant_chat(r: AssistIn, authorization: str = Header(None)):
    """Chat assistant on a FREE model (never Claude). Per-advocate daily limit, shared per-minute limit, PII removed, citations only from the verified library."""
    users, uid = current(authorization)
    prov = _assist_providers()
    if not prov: raise HTTPException(503, "Law Assistant abhi setup nahi hua (admin Render mein ASSIST_API_KEY lagaye)." if users[uid].get("is_admin") else "Law Assistant abhi chalu nahi hai. Admin se rabta karein.")
    msgs = _clean_msgs(r.messages)
    if not msgs or msgs[-1]["role"] != "user": raise HTTPException(400, "Pehle apna sawal likhein.")
    cap, today, adm = int(os.environ.get("ASSIST_DAILY", "30")), (datetime.utcnow() + timedelta(hours=5)).date().isoformat(), users[uid].get("is_admin")
    with _ALOCK:
        now = time.time(); ASSIST_HITS[:] = [t for t in ASSIST_HITS if now - t < 60]
        if len(ASSIST_HITS) >= int(os.environ.get("ASSIST_RPM", "10")): raise HTTPException(429, "Abhi bohat log istemal kar rahe hain. Ek minute baad dobara likhein.")
        use = rj(f"{DATA}/assistant_usage.json", {})
        if use.get("day") != today: use = {"day": today, "n": {}}
        n = int(use["n"].get(uid, 0))
        if n >= cap and not adm: raise HTTPException(429, f"Aaj ki {cap} messages ki had poori ho gayi. Kal dobara istemal karein.")
        use["n"][uid] = n + 1; wj(f"{DATA}/assistant_usage.json", use); ASSIST_HITS.append(now)
    q, path = msgs[-1]["content"][:600], dpath(uid, "mycitations")
    auths = cit.authorities(q, path)
    block = ("VERIFIED AUTHORITIES (cite ONLY by tag such as [A1]):\n" + "\n".join(f"[A{i}] {a['citation']} | {a.get('title', '')} | {str(a.get('summary', ''))[:250]}" for i, a in enumerate(auths, 1))
             if auths else "VERIFIED AUTHORITIES: none for this topic - do not cite any case law.")
    system = ASSIST_SYSTEM + f"\nReply language: {r.language}.\n" + block
    busy = False; text = ""
    for base, key, model in prov:
        try:
            text = _assist_call(base, key, model, system, msgs)
            if text: break
        except urllib.error.HTTPError as e:
            busy = busy or e.code == 429
            print(f"[assistant] {model} HTTP {e.code}: {e.read().decode('utf-8', 'ignore')[:200]}", flush=True)
        except Exception as e:
            print(f"[assistant] {model} fail: {type(e).__name__}: {str(e)[:200]}", flush=True)
    if not text:
        with _ALOCK:                                                     # a failed answer does not use up the advocate's daily limit
            use = rj(f"{DATA}/assistant_usage.json", {})
            if use.get("day") == today: use["n"][uid] = max(0, int(use["n"].get(uid, 1)) - 1); wj(f"{DATA}/assistant_usage.json", use)
        raise HTTPException(429 if busy else 503, "Free AI abhi busy hai (limit). Thori der baad dobara likhein." if busy else "Assistant abhi jawab nahi de saka. Thori der baad try karein.")
    if auths: text, _ = cit.render(text, auths, path, q)
    else:
        text, _ = cit.scrub(text, path); cit._gap(cit._topics(q))
    return {"reply": text, "left": None if adm else max(0, cap - n - 1)}
# --- ASSISTANT END

# --- JUDGMENT BEGIN
JUDG_DRAFTS, JUDG_JOBS, _JLOCK = {}, {}, threading.Lock()      # memory only (one server instance); drafts live one hour
J_PROMPT = (
    "You extract structured data from ONE Pakistani court judgment (its text follows). Reply with ONLY valid JSON - no markdown, no commentary.\n"
    "Never invent anything that is not in the text. Do NOT write any law-report citation (PLD, SCMR, CLC, YLR, PCrLJ ...) unless it is printed in the text itself; "
    "if it is, put it in printed_citation, otherwise leave it empty.\n"
    'Keys: "case_ref" (e.g. "Civil Appeal No. 123 of 2020"), "printed_citation", "court", "bench" (judges exactly as printed), "decision_date" (YYYY-MM-DD, "" if not printed), '
    '"title" (parties, e.g. "Ali v. The State"), "summary" (max 120 words, plain English: what was decided and why), "issues" (list, max 4), "holding" (1-3 sentences), '
    '"statutes" (list of Act + section exactly as written in the text), "keywords" (list of English words a lawyer would search with), '
    '"extracts" (list of up to 3 objects {"para": paragraph number as printed or "", "text": a VERBATIM quote copied character for character from the judgment, '
    "1 to 4 sentences, at most 700 characters, the passage that states the legal principle}).\n")


def _jnorm(s): return re.sub(r"[^a-z0-9]", "", str(s or "").lower())


def judgment_pages(pdf_bytes):
    import pdfplumber
    out = []
    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for i, pg in enumerate(pdf.pages, 1):
            t = pg.extract_text() or ""
            if t.strip(): out.append((i, t))
    return out


def _jparse(raw):
    s = str(raw or ""); a, b = s.find("{"), s.rfind("}")
    if a < 0 or b <= a: raise ValueError("AI ne JSON nahi diya")
    d = json.loads(s[a:b + 1])
    txt = lambda k: str(d.get(k) or "").strip()[:2000]
    lst = lambda k, n: [str(x).strip()[:300] for x in (d.get(k) or []) if str(x).strip()][:n] if isinstance(d.get(k), list) else []
    ex = [{"para": str(e.get("para", ""))[:20], "text": str(e.get("text", "")).strip()} for e in (d.get("extracts") or []) if isinstance(e, dict) and str(e.get("text", "")).strip()][:3]
    return {"case_ref": txt("case_ref")[:200], "printed_citation": txt("printed_citation")[:100], "court": txt("court")[:120], "bench": txt("bench")[:300],
            "decision_date": txt("decision_date")[:10], "title": txt("title")[:300], "summary": txt("summary"), "issues": lst("issues", 4), "holding": txt("holding"),
            "statutes": lst("statutes", 12), "keywords": lst("keywords", 15), "extracts": ex}


def judgment_verify(d, plain):
    """No AI here: what the model returned is compared with the judgment's own text. Quotes that are not in the text are removed."""
    pn, checks = _jnorm(plain), []
    chk = lambda field, ok, note: checks.append({"field": field, "ok": bool(ok), "note": note})
    good = []
    for e in d["extracts"]:
        n = _jnorm(e["text"]); hit = len(n) >= 40 and n in pn
        chk("ibarat", hit, ("faisle mein hu-ba-hu mili: " if hit else "AI ki likhi ibarat faisle mein nahi mili - hata di gayi: ") + e["text"][:70] + "...")
        if hit: good.append({"para": e["para"], "text": e["text"][:700]})
    d["extracts"] = good
    nums = re.findall(r"\d+", d["case_ref"])
    chk("case_ref", nums and all(n in plain for n in nums), "case number ke ankre faisle mein " + ("mile" if nums and all(n in plain for n in nums) else "nahi mile - check karein"))
    yr = re.search(r"(?:19|20)\d{2}", d["decision_date"])
    chk("decision_date", yr and yr.group(0) in plain, "tareekh ka saal faisle mein " + ("mila" if yr and yr.group(0) in plain else "nahi mila - check karein"))
    chk("court", d["court"] and _jnorm(d["court"])[:10] in pn, "court ka naam faisle mein " + ("mila" if d["court"] and _jnorm(d["court"])[:10] in pn else "nahi mila - check karein"))
    names = [x.strip() for x in re.split(r",|;| and ", d["bench"]) if len(x.strip()) > 3]
    bad = [x for x in names if _jnorm(re.sub(r"(?i)\b(mr|mrs|ms|dr|justice|hon|ble|chief|sardar|syed)\b\.?", "", x)) not in pn]
    chk("bench", names and not bad, "judges ke naam faisle mein " + ("mile" if names and not bad else "nahi mile: " + ", ".join(bad) if bad else "nahi likhe"))
    sides = [s for s in re.split(r"\s+(?:v\.?|vs\.?|versus)\s+", d["title"], flags=re.I) if s.strip()]
    chk("title", sides and all(_jnorm(s)[:12] in pn for s in sides), "parties ke naam faisle mein " + ("mile" if sides and all(_jnorm(s)[:12] in pn for s in sides) else "nahi mile - check karein"))
    if d["printed_citation"] and _jnorm(d["printed_citation"]) not in pn:        # a law-report citation is accepted only if it is printed in the judgment itself
        chk("citation", False, "AI ne '" + d["printed_citation"] + "' likha, lekin ye faisle mein chhapa hua nahi mila - hata diya")
        d["printed_citation"] = ""
    keep = [s for s in d["statutes"] if _jnorm(s) in pn]
    if len(keep) != len(d["statutes"]): chk("statutes", False, f"{len(d['statutes']) - len(keep)} Act/section faisle mein nahi mile - hata diye")
    d["statutes"] = keep
    return d, checks


def _judg_job(jid, text_ai, plain, uid):
    try:
        raw = call_claude(J_PROMPT + "\nJUDGMENT TEXT:\n" + text_ai, None, os.environ.get("MODEL_JUDGMENT") or "claude-sonnet-5-5",
                          lambda m, i, o: _note_usage("judgment", m, i, o), system="You are a careful legal data-extraction tool. You only report what the document says.")
        d, checks = judgment_verify(_jparse(raw), plain)
        with _JLOCK:
            JUDG_DRAFTS[jid] = {"text": plain, "ts": time.time()}
            JUDG_JOBS[jid] = {"status": "done", "draft": d, "checks": checks, "draft_id": jid, "ts": time.time()}
    except Exception as e:
        traceback.print_exc()
        with _JLOCK: JUDG_JOBS[jid] = {"status": "error", "error": str(e)[:200], "ts": time.time()}


class JudgIn(BaseModel): pdf_b64: str
class JobId(BaseModel): id: str

@app.post("/admin/judgment/extract")
def admin_judgment_extract(r: JudgIn, authorization: str = Header(None)):
    """Admin only: official judgment PDF -> Claude Sonnet reads it -> every quote / name / number is checked against the PDF's own text. Nothing is saved yet."""
    users, uid = current(authorization)
    if not users[uid].get("is_admin"): raise HTTPException(403, "Sirf admin.")
    if len(r.pdf_b64) > 28_000_000: raise HTTPException(400, "PDF bohat bara hai (20 MB se kam rakhein).")
    try:
        import base64
        pages = judgment_pages(base64.b64decode(r.pdf_b64))
    except Exception:
        raise HTTPException(400, "PDF nahi khul saka.")
    if not pages: raise HTTPException(400, "Is PDF mein text nahi (scan ki hui tasveer hai). Text wali PDF dein.")
    plain = "\n".join(t for _, t in pages)
    if len(plain) > 200_000: raise HTTPException(400, "Faisla bohat bara hai (2 lakh huroof se zyada). Pehle zaroori hissa alag PDF mein karein.")
    jid = secrets.token_hex(6)
    with _JLOCK:
        for k in [k for k, v in JUDG_DRAFTS.items() if time.time() - v["ts"] > 3600]: JUDG_DRAFTS.pop(k, None)
        for k in [k for k, v in JUDG_JOBS.items() if time.time() - v.get("ts", 0) > 3600]: JUDG_JOBS.pop(k, None)
        JUDG_JOBS[jid] = {"status": "pending", "ts": time.time()}
    threading.Thread(target=_judg_job, args=(jid, "\n\n".join(f"[p.{i}]\n{t}" for i, t in pages), plain, uid), daemon=True).start()
    return {"job": jid, "pages": len(pages), "chars": len(plain)}

@app.post("/admin/judgment/job")
def admin_judgment_job(r: JobId, authorization: str = Header(None)):
    users, uid = current(authorization)
    if not users[uid].get("is_admin"): raise HTTPException(403, "Sirf admin.")
    with _JLOCK: j = JUDG_JOBS.get(r.id)
    if not j: raise HTTPException(404, "Ye kaam nahi mila (server dobara chala ho sakta hai). PDF dobara upload karein.")
    return {k: v for k, v in j.items() if k != "ts"}
# --- JUDGMENT END

# --- HUB BEGIN
HUB_FILE = {k: f"{DATA}/hub_{k}.json" for k in ("links", "thana", "calendar", "news")}
HUB_COLS = {"links": ("section", "district", "title", "url", "note", "daily"), "thana": ("district", "thana", "court", "tehsil", "notes", "source"),
            "calendar": ("start", "end", "title", "kind", "source"), "news": ("date", "title", "url", "source")}
HUB_MAX = {"links": 400, "thana": 6000, "calendar": 400, "news": 300}
HUB_SECTIONS = ("district", "highcourt", "bareact", "bar")
HUB_SEED = [                                         # official sites that were checked; everything else is filled in by the admin
    {"section": "highcourt", "district": "", "title": "Lahore High Court", "url": "https://lhc.gov.pk/", "note": "Cause list, judgments, rules, judges", "daily": ""},
    {"section": "highcourt", "district": "", "title": "Supreme Court of Pakistan", "url": "https://www.supremecourt.gov.pk/", "note": "Cause lists and judgments", "daily": ""},
    {"section": "highcourt", "district": "", "title": "Sindh High Court", "url": "https://sindhhighcourt.gov.pk/", "note": "Cause list, case search, judgments", "daily": ""},
    {"section": "highcourt", "district": "", "title": "Balochistan High Court", "url": "https://bhc.gov.pk/", "note": "Cause list and judgments", "daily": ""},
    {"section": "district", "district": "*", "title": "Punjab District Courts - Case Management System", "url": "https://dsj.punjab.gov.pk/", "note": "Date-wise cause list and case details of the district courts", "daily": "1"},
    {"section": "bareact", "district": "", "title": "Pakistan Code (official laws database)", "url": "https://pakistancode.gov.pk/", "note": "Acts, Ordinances and Rules - Law & Justice Commission of Pakistan", "daily": ""}]
_URL_OK = re.compile(r"^https?://[^\s/]+\.[^\s/]+", re.I)


def hub_rows(kind):
    rows = rj(HUB_FILE[kind], None)
    if rows is None:
        rows = [dict(x) for x in HUB_SEED] if kind == "links" else []
        if rows: wj(HUB_FILE[kind], rows)
    return rows


def _hub_clean(kind, rec):
    r = {k: str(rec.get(k) or "").strip()[:500 if k == "url" else 300] for k in HUB_COLS[kind]}
    if kind == "links":
        if r["section"] not in HUB_SECTIONS: return None, "section district / highcourt / bareact / bar mein se ho"
        if not r["title"] or not _URL_OK.match(r["url"]): return None, "title aur sahi link (https://...) zaroori hain"
        if r["section"] != "district": r["district"] = ""
        elif r["district"] not in ("", "*"): r["district"] = r["district"].title()
    elif kind == "thana":
        if not (r["district"] and r["thana"] and r["court"]): return None, "district, thana aur court zaroori hain"
        r["district"] = r["district"].title()
    elif kind == "calendar":
        try:
            s = date.fromisoformat(r["start"]); e = date.fromisoformat(r["end"] or r["start"])
        except ValueError: return None, "tareekh YYYY-MM-DD mein likhein"
        if e < s or not r["title"]: return None, "end start se pehle nahi ho sakti aur title zaroori hai"
        r["start"], r["end"] = s.isoformat(), e.isoformat(); r["kind"] = r["kind"].lower() if r["kind"].lower() in ("holiday", "vacation", "other") else "holiday"
    else:
        if not r["title"] or not _URL_OK.match(r["url"]): return None, "title aur sahi link (https://...) zaroori hain"
        try: r["date"] = date.fromisoformat(r["date"]).isoformat() if r["date"] else date.today().isoformat()
        except ValueError: return None, "tareekh YYYY-MM-DD mein likhein"
    return r, ""


def hub_import(kind, text):
    import csv, io as _io
    rows, added, skipped, errs = hub_rows(kind), 0, 0, []
    for i, rec in enumerate(csv.DictReader(_io.StringIO(text or "")), 2):
        r, why = _hub_clean(kind, rec)
        if r is None: skipped += 1; errs.append(f"line {i}: {why}") if len(errs) < 5 else None; continue
        if r in rows: skipped += 1; continue
        if len(rows) >= HUB_MAX[kind]: skipped += 1; errs.append("hadd poori (fehrist full hai)") if len(errs) < 5 else None; continue
        rows.append(r); added += 1
    if added: wj(HUB_FILE[kind], rows)
    return {"added": added, "skipped": skipped, "errors": errs, "total": len(rows)}


_RSS = {}


def _news_date(s):
    s = str(s or "").strip()
    try:
        if re.match(r"\d{4}-\d{2}-\d{2}", s): return s[:10]
        from email.utils import parsedate_to_datetime
        return parsedate_to_datetime(s).date().isoformat()
    except Exception: return date.today().isoformat()


def _rss_items(url):
    """Headlines + links only, from feeds the ADMIN listed in NEWS_FEEDS; cached for 30 minutes."""
    c = _RSS.get(url)
    if c and time.time() - c[0] < 1800: return c[1]
    items = []
    try:
        import xml.etree.ElementTree as ET
        with urllib.request.urlopen(urllib.request.Request(url, headers={"User-Agent": "AdvocateAI/1.0"}), timeout=8) as r: root = ET.fromstring(r.read(600_000))
        atom = "{http://www.w3.org/2005/Atom}"
        host = re.sub(r"^https?://(www\.)?", "", url).split("/")[0]
        for it in (list(root.iter("item"))[:15] + list(root.iter(atom + "entry"))[:15]):
            t = (it.findtext("title") or it.findtext(atom + "title") or "").strip()
            link = (it.findtext("link") or "").strip()
            if not link:
                a = it.find(atom + "link"); link = (a.get("href") if a is not None else "") or ""
            if t and _URL_OK.match(link): items.append({"date": _news_date(it.findtext("pubDate") or it.findtext(atom + "updated") or ""), "title": t[:200], "url": link, "source": host})
    except Exception as e:
        print("[news] feed failed:", url, type(e).__name__, flush=True)
    _RSS[url] = (time.time(), items)
    return items


@app.get("/hub/links")
def hub_links(district: str = "", authorization: str = Header(None)):
    current(authorization)
    d, rows = district.strip().title(), hub_rows("links")
    pick = lambda sec: [{k: x.get(k, "") for k in ("title", "url", "note", "daily")} for x in rows if x["section"] == sec]
    return {"districts": sorted(set(cs._URD.values())), "district_name": d,
            "district": [{k: x.get(k, "") for k in ("title", "url", "note", "daily")} for x in rows if x["section"] == "district" and x.get("district", "") in ("", "*", d)],
            "highcourt": pick("highcourt"), "bareact": pick("bareact"), "bar": pick("bar")}

@app.get("/hub/news")
def hub_news(authorization: str = Header(None)):
    current(authorization)
    items = list(hub_rows("news"))
    for u in [x.strip() for x in os.environ.get("NEWS_FEEDS", "").split(",") if x.strip().startswith("http")][:6]: items += _rss_items(u)
    items.sort(key=lambda x: x.get("date", ""), reverse=True)
    return {"items": items[:60]}

@app.get("/hub/calendar")
def hub_calendar(authorization: str = Header(None)):
    current(authorization)
    since = (date.today() - timedelta(days=1)).isoformat()
    return {"items": sorted((x for x in hub_rows("calendar") if x["end"] >= since), key=lambda x: x["start"])[:200]}

@app.get("/hub/thana")
def hub_thana(district: str = "", q: str = "", authorization: str = Header(None)):
    current(authorization)
    d, k = district.strip().lower(), q.strip().lower()
    rows = [x for x in hub_rows("thana") if (not d or x["district"].lower() == d) and (not k or k in (x["thana"] + " " + x["court"] + " " + x["tehsil"]).lower())]
    return {"items": rows[:60], "total": len(rows)}


class HubIn(BaseModel): name: str; text: str = ""; index: int = -1

@app.post("/admin/hub/import")
def admin_hub_import(r: HubIn, authorization: str = Header(None)):
    _admin_only(authorization)
    if r.name not in HUB_COLS: raise HTTPException(400, "Qism sahi nahi.")
    return hub_import(r.name, r.text)

@app.post("/admin/hub/list")
def admin_hub_list(r: HubIn, authorization: str = Header(None)):
    _admin_only(authorization)
    if r.name not in HUB_COLS: raise HTTPException(400, "Qism sahi nahi.")
    return {"rows": hub_rows(r.name)[:400], "columns": list(HUB_COLS[r.name])}

@app.post("/admin/hub/delete")
def admin_hub_delete(r: HubIn, authorization: str = Header(None)):
    _admin_only(authorization)
    if r.name not in HUB_COLS: raise HTTPException(400, "Qism sahi nahi.")
    rows = hub_rows(r.name)
    if not 0 <= r.index < len(rows): raise HTTPException(404, "Ye entry nahi mili.")
    rows.pop(r.index); wj(HUB_FILE[r.name], rows); return {"ok": True, "total": len(rows)}
# --- HUB END

class CitSearch(BaseModel): q: str
class CitAdd(BaseModel): citation: str; title: str; source: str; court: str = ""; date: str = ""; bench: str = ""; summary: str = ""; keywords: str = ""; sections: str = ""; url: str = ""; extract: str = ""; para: str = ""; status: str = ""; official: str = ""; draft_id: str = ""

@app.post("/citations/search")
def citations_search(r: CitSearch, authorization: str = Header(None)):
    """Search the verified library (global + this advocate's own). Nothing here comes from the AI."""
    _, uid = current(authorization)
    return {"results": [{k: str(x.get(k, "")) for k in ("citation", "title", "court", "date", "bench", "summary", "extract", "para", "status", "source", "verified_by", "verified_on", "url")}
                        for x in cit.search(r.q, dpath(uid, "mycitations"))]}

@app.post("/admin/citations/add")
def admin_citations_add(r: CitAdd, authorization: str = Header(None)):
    users, uid = current(authorization)
    if not users[uid].get("is_admin"): raise HTTPException(403, "Sirf admin.")
    rec = dict(r); did = rec.pop("draft_id", "")
    if did:                                              # filled from an AI-read judgment: the quote must still be in that judgment's own text
        with _JLOCK: dr = JUDG_DRAFTS.get(did)
        if not dr: raise HTTPException(410, "Ye draft khatam ho gaya (server dobara chala ya 1 ghanta guzar gaya). PDF dobara upload karein.")
        ex = str(rec.get("extract", "")).strip()
        if ex and _jnorm(ex) not in _jnorm(dr["text"]): raise HTTPException(400, "Ye ibarat faisle ke text mein hu-ba-hu nahi mili. Sirf faisle se copy ki hui ibarat dein.")
        rec["official"], rec["ai_assisted"] = "1", "summary"
    ok, msg = cit.add_global(rec, users[uid].get("chamber_name", ""))
    if not ok: raise HTTPException(400, msg)
    return {"ok": True}

@app.post("/admin/citations/gaps")
def admin_citations_gaps(authorization: str = Header(None)):
    users, uid = current(authorization)
    if not users[uid].get("is_admin"): raise HTTPException(403, "Sirf admin.")
    return {"gaps": cit.gaps_list()}

@app.post("/admin/citations/requests")
def admin_citations_requests(authorization: str = Header(None)):
    users, uid = current(authorization)
    if not users[uid].get("is_admin"): raise HTTPException(403, "Sirf admin.")
    return {"requests": cit.requests_list()}

@app.post("/citations/import")
def citations_import(r: CitImport, authorization: str = Header(None)):
    users, uid = current(authorization)
    if not users[uid].get("is_admin"): raise HTTPException(403)
    return cit.import_text(r.text, r.source)

# ---------------------------------------------------------------- district court case sync
# --- DIARY BEGIN
_AR = "\u0600-\u06FF\u0750-\u077F\uFB50-\uFDFF\uFE70-\uFEFF"
_AR_RUN = re.compile(f"[{_AR}][{_AR}\u200c\u200d ]*")
_DIARY_FONTS = ("/usr/share/fonts/truetype/noto/NotoNaskhArabic-Regular.ttf", "/usr/share/fonts/opentype/noto/NotoNaskhArabic-Regular.ttf",
                "/usr/share/fonts/truetype/noto/NotoSansArabic-Regular.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf")


def _diary_setup():
    """Urdu in the PDF needs a font file AND arabic_reshaper + python-bidi on the server; without them Urdu words print as [Urdu]."""
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    ok = False
    for p in [os.environ.get("DIARY_FONT", "")] + list(_DIARY_FONTS):
        if p and os.path.exists(p):
            try: pdfmetrics.registerFont(TTFont("DiaryFont", p)); ok = True; break
            except Exception: pass
    shape = None
    if ok:
        try:
            import arabic_reshaper
            from bidi.algorithm import get_display
            shape = lambda seg: get_display(arabic_reshaper.reshape(seg))
        except Exception: shape = None
    return shape


def _dmy(d):
    m = re.fullmatch(r"(\d{4})-(\d{2})-(\d{2})", str(d or ""))
    return f"{m.group(3)}-{m.group(2)}-{m.group(1)}" if m else str(d or "-")


def build_diary_pdf(clients, chamber, day):
    """Daily diary of ALL the advocate's cases: today's hearings (with space to write the result), upcoming dates, then dates that passed / are missing."""
    from io import BytesIO
    from xml.sax.saxutils import escape
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
    shape, note = _diary_setup(), [False]

    def txt(s):
        s, out, last = str(s or ""), [], 0
        for m in _AR_RUN.finditer(s):
            out.append(escape(s[last:m.start()])); seg = m.group(0).strip()
            if shape and seg: out.append('<font name="DiaryFont">' + escape(shape(seg)) + "</font>" + (" " if m.group(0).endswith(" ") else ""))
            else: out.append("[Urdu] "); note[0] = True
            last = m.end()
        return "".join(out) + escape(s[last:])

    base = ParagraphStyle("b", fontName="Helvetica", fontSize=8.5, leading=10.5)
    hd = ParagraphStyle("hd", parent=base, fontName="Helvetica-Bold", textColor=colors.white)
    h1 = ParagraphStyle("h1", parent=base, fontName="Helvetica-Bold", fontSize=18, leading=22)
    h2 = ParagraphStyle("h2", parent=base, fontName="Helvetica-Bold", fontSize=11.5, leading=15, spaceBefore=12, spaceAfter=4)
    P = lambda *lines, st=base: Paragraph("<br/>".join(txt(x) for x in lines if str(x or "").strip()) or "-", st)

    def ids(c):
        fir, cn = str(c.get("fir") or ""), str(c.get("case_no") or "")
        return [fir, ("Case no: " + cn) if cn and cn not in fir else "", ("PS: " + str(c["ps"])) if c.get("ps") else ""]

    def table(head, rows, widths):
        t = Table([[Paragraph(h, hd) for h in head]] + rows, colWidths=[w * mm for w in widths], repeatRows=1)
        t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0b5d44")), ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#9aa7a1")),
                               ("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 4), ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                               ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                               ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f2f6f4")])]))
        return t

    nm = lambda c: str(c.get("name") or "")
    cl = [c for c in clients if isinstance(c, dict)]
    now = [c for c in cl if c.get("date") == day]
    later = [c for c in cl if str(c.get("date") or "") > day]
    gone = [c for c in cl if not c.get("date") or str(c.get("date")) < day]
    now.sort(key=lambda c: (str(c.get("court") or ""), nm(c))); later.sort(key=lambda c: (str(c["date"]), nm(c))); gone.sort(key=lambda c: (str(c.get("date") or "9999"), nm(c)))
    try:
        import datetime as _dt
        wd = _dt.date.fromisoformat(day).strftime("%A, %d %B %Y")
    except Exception: wd = _dmy(day)

    S = [Paragraph("DAILY DIARY", h1), P(chamber or "Advocate", st=ParagraphStyle("s", parent=base, fontSize=11, leading=14)), P(wd, st=ParagraphStyle("s2", parent=base, fontSize=11, leading=14)),
         Spacer(1, 3), P(f"Aaj ki peshi: {len(now)}   |   Aane wali peshian: {len(later)}   |   Tareekh guzri / darj nahi: {len(gone)}   |   Kul cases: {len(cl)}")]
    S.append(Paragraph(f"1) Aaj ki peshi ({_dmy(day)})", h2))
    if now:
        S.append(table(["#", "Client", "Phone", "FIR / Case no", "Court", "Stage", "Aaj ki karwai / agli tareekh"],
                       [[Paragraph(str(i), base), P(nm(c)), P(c.get("phone", "")), P(*ids(c)), P(c.get("court", "")), P(c.get("stage", "")), Paragraph("<br/><br/><br/>", base)]
                        for i, c in enumerate(now, 1)], [8, 45, 28, 45, 55, 35, 61]))
    else: S.append(P("Is din ki koi peshi darj nahi."))
    S.append(Paragraph("2) Aane wali peshian", h2))
    if later:
        S.append(table(["Tareekh", "Client", "FIR / Case no", "Court", "Stage", "Notes"],
                       [[Paragraph(_dmy(c["date"]), base), P(nm(c)), P(*ids(c)), P(c.get("court", "")), P(c.get("stage", "")), P(str(c.get("notes") or "")[:110])] for c in later],
                       [22, 50, 48, 55, 40, 62]))
    else: S.append(P("Aage ki koi tareekh darj nahi."))
    S.append(Paragraph("3) Tareekh guzar chuki ya darj nahi (update karein)", h2))
    if gone:
        S.append(table(["Aakhri tareekh", "Client", "FIR / Case no", "Court", "Stage", "Notes"],
                       [[Paragraph(_dmy(c.get("date")), base), P(nm(c)), P(*ids(c)), P(c.get("court", "")), P(c.get("stage", "")), P(str(c.get("notes") or "")[:110])] for c in gone],
                       [22, 50, 48, 55, 40, 62]))
    else: S.append(P("Koi nahi."))
    if note[0]: S.append(Spacer(1, 8)); S.append(P("Note: Urdu likhai is PDF mein nahi aa saki ([Urdu] likha hai) - app mein dekh lein.", st=ParagraphStyle("n", parent=base, textColor=colors.grey)))

    def foot(cv, d):
        cv.saveState(); cv.setFont("Helvetica", 7.5); cv.setFillColor(colors.grey)
        cv.drawString(12 * mm, 8 * mm, f"Advocate AI - Daily Diary {_dmy(day)}"); cv.drawRightString(landscape(A4)[0] - 12 * mm, 8 * mm, f"Page {d.page}"); cv.restoreState()
    bio = BytesIO()
    SimpleDocTemplate(bio, pagesize=landscape(A4), leftMargin=12 * mm, rightMargin=12 * mm, topMargin=12 * mm, bottomMargin=14 * mm, title="Daily Diary").build(S, onFirstPage=foot, onLaterPages=foot)
    return bio.getvalue()
# --- DIARY END


class CasesIn(BaseModel): raw: str = ""; pdf_b64: str = ""; district: str = ""; ps: str = ""
class CasesConfirm(BaseModel): client_index: int; row: dict; case_no: str = ""
class CasesFind(BaseModel): raw: str = ""; pdf_b64: str = ""; case_no: str

@app.post("/cases/ingest")
def cases_ingest(r: CasesIn, authorization: str = Header(None)):
    _, uid = current(authorization)
    if len(r.raw) > 600_000 or len(r.pdf_b64) > 12_000_000: raise HTTPException(400, "Data bohat bara hai.")
    path = dpath(uid, "clients"); clients = rj(path, [])
    res = cs.ingest(r.raw, clients, r.district, r.ps, rows=_pdf_rows(r.pdf_b64) if r.pdf_b64 else None)
    if res["updated"]: wj(path, clients)
    return res

@app.post("/cases/confirm")
def cases_confirm(r: CasesConfirm, authorization: str = Header(None)):
    _, uid = current(authorization)
    path = dpath(uid, "clients"); clients = rj(path, [])
    if not 0 <= r.client_index < len(clients): raise HTTPException(404, "Client nahi mila.")
    c = clients[r.client_index]; ch = cs.apply(c, cs.clean_row(r.row))
    cn = r.case_no.strip()[:40]
    if cn and cs._pairs_of(cn) and cn != str(c.get("case_no") or ""):        # number typed by the advocate: keep it so the next lists match by themselves
        if c.get("case_no"): c["old_case_no"] = c["case_no"]
        c["case_no"] = cn; ch = list(ch) + ["case no: " + cn]
    wj(path, clients)
    return {"changes": ch}

@app.post("/cases/find")
def cases_find(r: CasesFind, authorization: str = Header(None)):
    """Look one specific case number up in the list that was just uploaded (nothing is changed here)."""
    current(authorization)
    if len(r.raw) > 600_000 or len(r.pdf_b64) > 12_000_000: raise HTTPException(400, "Data bohat bara hai.")
    want = cs._pairs_of(r.case_no)
    if not want: raise HTTPException(400, "Case number sahi likhein (jaise 2633798/26 ya 776/2026).")
    rows = cs._collapse(_pdf_rows(r.pdf_b64) if r.pdf_b64 else cs.parse(r.raw), cs.date.today().isoformat())
    hits = [x for x in rows if want & set(map(tuple, x["pairs"]))]
    return {"matches": [{"preview": cs._prev(x), "row": cs.clean_row(x)} for x in hits[:5]]}

class DiaryIn(BaseModel): date: str = ""

@app.post("/diary/pdf")
def diary_pdf(r: DiaryIn, authorization: str = Header(None)):
    """Daily diary PDF of every client of the logged-in advocate. date = YYYY-MM-DD (empty = today, Pakistan time)."""
    users, uid = current(authorization)
    day = r.date.strip() or (datetime.utcnow() + timedelta(hours=5)).date().isoformat()
    if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", day): raise HTTPException(400, "Tareekh YYYY-MM-DD mein likhein (jaise 2026-11-14).")
    try: pdf = build_diary_pdf(rj(dpath(uid, "clients"), []), users[uid].get("chamber_name", ""), day)
    except ImportError: raise HTTPException(500, "Diary PDF ke liye server par reportlab install karein (requirements.txt mein 'reportlab' likh kar dobara deploy karein).")
    return Response(pdf, media_type="application/pdf", headers={"Content-Disposition": f'attachment; filename="Daily_Diary_{day}.pdf"'})

# Full automation switch: only if PITB / D&SJ grants an official feed (set COURT_API_URL and COURT_API_KEY on the server).
COURT_API_URL, COURT_API_KEY = os.environ.get("COURT_API_URL", ""), os.environ.get("COURT_API_KEY", "")

def _court_poll_causelist():
    """COURT_FEED_MODE=causelist: one feed request per district per day (all courts of that city), then update every advocate's clients."""
    import datetime as dt
    days = [(dt.date.today() + dt.timedelta(days=i)).isoformat() for i in range(max(1, int(os.environ.get("COURT_DAYS_AHEAD", "3"))))]
    data = {fn: rj(f"{DATA}/{fn}", []) for fn in data_files() if fn.endswith("_clients.json")}
    def fetch(d, day):
        time.sleep(1)
        return cs.fetch_causelist(COURT_API_URL, COURT_API_KEY, d, day)
    for fn in cs.sweep_causelist(data, fetch, days): wj(f"{DATA}/{fn}", data[fn])

def _court_poll_once():
    if os.environ.get("COURT_FEED_MODE", "case") == "causelist": return _court_poll_causelist()
    for fn in data_files():
        if not fn.endswith("_clients.json"): continue
        path, changed = f"{DATA}/{fn}", False
        clients = rj(path, [])
        for c in clients:
            try: res = cs.sync_one(c, cs.fetch_official(COURT_API_URL, COURT_API_KEY, c))
            except Exception as e:
                print(f"[court] feed error: {type(e).__name__}: {str(e)[:120]}", flush=True); continue
            changed = changed or bool(res)
            time.sleep(1)                                   # be gentle with the feed
        if changed: wj(path, clients)

def _court_poller():
    while True:
        try: _court_poll_once()
        except Exception: traceback.print_exc()
        time.sleep(max(1.0, float(os.environ.get("COURT_POLL_HOURS", "6"))) * 3600)

@app.on_event("startup")
def _start_court_poller():
    if COURT_API_URL: threading.Thread(target=_court_poller, daemon=True).start()

def _pdf_rows(b64):
    """Daily Board PDF (the advocate's own board) -> rows with real table cells."""
    import base64
    try: return cs.board_rows(base64.b64decode(b64))
    except RuntimeError as e: raise HTTPException(501, str(e))
    except Exception: raise HTTPException(400, "PDF parh nahi saka. Daily Board PDF dobara chunein.")

class CasesAllIn(BaseModel): raw: str = ""; pdf_b64: str = ""; dry: bool = False

@app.post("/admin/cases/ingest")
def admin_cases_ingest(r: CasesAllIn, authorization: str = Header(None)):
    """Admin solves the portal's check himself and pastes the result (or uploads a board PDF) once; every advocate's clients that
    match with high confidence are updated. The admin never sees other advocates' client names, only counts."""
    users, uid = current(authorization)
    if not users[uid].get("is_admin"): raise HTTPException(403, "Sirf admin.")
    if len(r.raw) > 3_000_000 or len(r.pdf_b64) > 12_000_000: raise HTTPException(400, "Data bohat bara hai.")
    import copy
    rows, checked, upd, review, batch = (_pdf_rows(r.pdf_b64) if r.pdf_b64 else cs.parse(r.raw)), 0, 0, 0, []
    for x in rows: (x.get("extra") or {}).pop("comments", None)     # the uploading office's own comments never go to other advocates
    for fn in data_files():
        if not fn.endswith("_clients.json"): continue
        path = f"{DATA}/{fn}"; clients = rj(path, []); before = copy.deepcopy(clients)
        res = cs.ingest("", clients, rows=rows); checked += 1
        upd += len(res["updated"]); review += len(res["suggestions"])
        if res["updated"] and not r.dry:                       # dry = preview only, nothing is written
            batch.append({"file": fn, "entries": cs.changed_entries(before, clients)}); wj(path, clients)
    if batch: wj(f"{DATA}/_last_bulk_update.json", {"at": time.time(), "batch": batch})      # so the admin can undo it
    return {"rows": len(rows), "dated": any(x.get("date") for x in rows), "advocates_checked": checked, "clients_updated": upd, "needs_review": review, "preview": r.dry}

@app.post("/admin/cases/numbers")
def admin_cases_numbers(authorization: str = Header(None)):
    """Admin only: case / FIR numbers (with court, date, stage) of every advocate's clients. Client names and phone numbers are not included."""
    users, uid = current(authorization)
    if not users[uid].get("is_admin"): raise HTTPException(403, "Sirf admin.")
    out = []
    for k, u in users.items():
        if u.get("is_admin"): continue
        cases, missing, total = _adv_cases(k)
        out.append({"id": k, "advocate": str(u.get("chamber_name", "")), "clients": total, "without_number": missing, "cases": cases})
    out.sort(key=lambda a: a["advocate"].lower())
    try:                                                              # who looked, and when - kept so the access can be shown to advocates
        lg = rj(f"{DATA}/_admin_views.json", []); lg.append({"at": datetime.utcnow().isoformat() + "Z", "admin": uid, "what": "cases/numbers"})
        wj(f"{DATA}/_admin_views.json", lg[-500:])
    except Exception: pass
    return {"advocates": out, "total_cases": sum(len(a["cases"]) for a in out)}

@app.post("/admin/usage")
def admin_usage(authorization: str = Header(None)):
    """Admin only: AI calls and the estimated cost (USD) per month, task kind and model."""
    users, uid = current(authorization)
    if not users[uid].get("is_admin"): raise HTTPException(403, "Sirf admin.")
    out = {}
    for mo, d in sorted(rj(f"{DATA}/ai_usage.json", {}).items())[-3:]:
        rows, total = [], 0.0
        for k, e in sorted(d.items()):
            kind, model = k.split("|", 1); p = PRICES.get(model)
            usd = round(e["in"] / 1e6 * p[0] + e["out"] / 1e6 * p[1], 4) if p else None
            total += usd or 0
            rows.append({"kind": kind, "model": model, "calls": e["n"], "input_tokens": e["in"], "output_tokens": e["out"], "usd": usd})
        out[mo] = {"rows": rows, "total_usd": round(total, 4)}
    return {"months": out, "models_in_use": {"default": MODEL, "light_tasks": FAST_MODEL, "light_kinds": sorted(FAST_KINDS)}}

class AdvId(BaseModel): id: str

@app.post("/admin/alerts")
def admin_alerts(authorization: str = Header(None)):
    """Admin only: cases that need the admin's update NOW - hearing today, or passed with no new date (red) and tomorrow (orange).
    Grouped by district / court so the admin knows which cause list to pull. Case numbers only: no client names, no phone numbers."""
    users, uid = current(authorization)
    if not users[uid].get("is_admin"): raise HTTPException(403, "Sirf admin.")
    today = (datetime.utcnow() + timedelta(hours=5)).date()
    t0, t1, oldest = today.isoformat(), (today + timedelta(days=1)).isoformat(), (today - timedelta(days=30)).isoformat()
    groups, old, nocms = {}, 0, 0
    for k, u in users.items():
        if u.get("is_admin"): continue
        for c in rj(dpath(k, "clients"), []):
            if not isinstance(c, dict): continue
            d = str(c.get("date") or "")
            if not d or d > t1: continue                         # no date, or later than tomorrow: nothing to do yet
            state = "aaj" if d == t0 else "kal" if d == t1 else "guzri"
            if state == "guzri" and d < oldest: old += 1; continue   # very old: probably finished, only counted
            cn = _cms(c.get("case_no"))
            nocms += 0 if cn else 1
            dist = (cs.district_of(c) or "").title()
            g = groups.setdefault(dist or str(c.get("court") or "")[:40] or "Court nahi likha", {"aaj": 0, "guzri": 0, "kal": 0, "cases": []})
            g[state] += 1
            g["cases"].append({"id": k, "advocate": str(u.get("chamber_name", "")), "case_no": cn or "", "fir": str(c.get("fir") or ""),
                               "court": str(c.get("court") or ""), "date": d, "state": state})
    order = {"aaj": 0, "guzri": 1, "kal": 2}
    out = []
    for name, g in groups.items():
        g["cases"].sort(key=lambda x: (order[x["state"]], x["case_no"] or "~"))
        out.append({"group": name, **g, "hint": f"{name} ki cause list /board se upload karein - ye cases khud update ho jayenge."})
    out.sort(key=lambda g: -(g["aaj"] + g["guzri"] + g["kal"] * 0.1))
    return {"today": t0, "red": sum(g["aaj"] + g["guzri"] for g in out), "orange": sum(g["kal"] for g in out), "no_case_no": nocms, "older_than_30_days": old, "groups": out}

@app.post("/admin/library/todo")
def admin_library_todo(authorization: str = Header(None)):
    """Admin only: what to do in the verified library today, most useful first."""
    users, uid = current(authorization)
    if not users[uid].get("is_admin"): raise HTTPException(403, "Sirf admin.")
    reqs, gaps, (cov, total), recheck = cit.requests_list(), cit.gaps_list(), cit.coverage(), cit.recheck_list()
    thin = [x for x in cov if x["entries"] < 5]
    tasks, dreq, drows = [], dafa_requests(), len(dafa_rows())
    if dreq: tasks.append({"priority": 0, "title": f"Dafa table mein {len(dreq)} dafaat ki tafseel nahi (lawyers ne FIR / dafa mein dekhe)",
                           "how": "Har dafa ki saza, bailable, cognizable, compoundable aur court Pakistan Code website se dekh kar 'Dafa table' card mein add karein (zareya zaroor likhein). FIR tajziya aur dafaat ke jawab inhi se aate hain.",
                           "items": [f"{x['act'] or '?'} {x['section']} - {x['count']} dafa" for x in dreq[:15]]})
    if drows < 20: tasks.append({"priority": 0, "title": f"Dafa table mein abhi sirf {drows} dafaat hain", "how": "Pehle sab se aam dafaat bharein:",
                                 "items": ["PPC: 302, 324, 337-A/F, 354, 365, 376, 379, 380, 392, 395, 406, 420, 489-F, 506", "CrPC: 22-A/B, 154, 161, 173, 265-K, 497, 498, 561-A", "CNSA 1997: 9"]})
    if reqs: tasks.append({"priority": 1, "title": f"Lawyers ne {len(reqs)} citation check kiye jo library mein nahi hain",
                           "how": "Har ek ko apni subscription ya court ki website par dekh kar 'Library mein add karein' se add karein (zareya zaroor likhein). Add hote hi darkhwast khud band ho jati hai.",
                           "items": [f"{r['citation']} - {r.get('count', 1)} dafa" for r in reqs[:15]]})
    if gaps: tasks.append({"priority": 2, "title": f"{len(gaps)} topics par lawyers ko library se jawab nahi mila",
                           "how": "Har topic ke liye 5 se 10 sab se mashhoor, khud tasdeeq kiye hue faisle add karein (Supreme Court pehle). Faisle ki asal ibarat aur para zaroor likhein.",
                           "items": [f"{g['topic']} - {g['count']} dafa" for g in gaps[:15]]})
    if thin: tasks.append({"priority": 3, "title": f"{len(thin)} topics mein 5 se kam entries hain",
                           "how": "Pehle wo topics bharein jin mein 0 hai. Fehrist: backend\\library_topics.txt.", "items": [f"{x['topic']} - {x['entries']} entries" for x in thin[:15]]})
    if recheck: tasks.append({"priority": 4, "title": f"{len(recheck)} entries 'dobara check' ya 'distinguished' hain",
                              "how": "Un faislon ko dobara dekhein: ab bhi theek hain, ya overruled hue? Status update karein.", "items": [f"{x['citation']} - {x['status']}" for x in recheck[:15]]})
    tasks.append({"priority": 9, "title": "Roz ka mamool (10 minute)", "how": "1) Darkhwastein dekhein aur add karein. 2) 'Library kahan khali hai' dekhein. 3) Ek topic ki 2 se 3 nayi entries add karein.", "items": []})
    for i, t in enumerate(tasks, 1): t["priority"] = i
    return {"library_entries": total, "tasks": tasks}

# --- DAFA+AB BEGIN
class DafaAdd(BaseModel): act: str; section: str; title: str; source: str; punishment: str = ""; bailable: str = ""; cognizable: str = ""; compoundable: str = ""; court: str = ""; notes: str = ""
class TextIn(BaseModel): text: str


def _admin_only(authorization):
    users, uid = current(authorization)
    if not users[uid].get("is_admin"): raise HTTPException(403, "Sirf admin.")
    return users, uid

@app.post("/admin/dafa/add")
def admin_dafa_add(r: DafaAdd, authorization: str = Header(None)):
    users, uid = _admin_only(authorization)
    ok, msg = dafa_add(dict(r), users[uid].get("chamber_name", ""))
    if not ok: raise HTTPException(400, msg)
    return {"ok": True, "rows": len(dafa_rows())}

@app.post("/admin/dafa/import")
def admin_dafa_import(r: TextIn, authorization: str = Header(None)):
    users, uid = _admin_only(authorization)
    return {**dafa_import(r.text, users[uid].get("chamber_name", "")), "rows": len(dafa_rows())}

@app.post("/admin/dafa/requests")
def admin_dafa_requests(authorization: str = Header(None)):
    _admin_only(authorization)
    return {"requests": dafa_requests(), "rows": len(dafa_rows())}


AB_JOBS = {}
class AbIn(BaseModel): kind: str = "draft"; language: str = "English"; text: str = ""; doc_type: str = "Petition"; category: str = "Criminal"; models: list = ["free", "claude-haiku-5-5", "claude-sonnet-5-5"]


def _ab_fields(a):
    return ({"category": a.category, "doc_type": a.doc_type, "facts": a.text} if a.kind == "draft" else {"text": a.text} if a.kind == "fir"
            else {"act": "Pakistani law", "question": a.text} if a.kind == "research" else {"mode": "topic", "topic": a.text})


def _ab_job(jid, uid, a):
    """The same question to the free AI, Haiku and Sonnet: same prompt, same server-side checks, so the numbers can be compared."""
    try:
        ar = AiReq(kind=a.kind, language=a.language, fields=_ab_fields(a))
        prompt, label = build_prompt(ar, False)
        if a.kind in ("draft", "research", "section"):
            q = " ".join(v for v in ar.fields.values() if isinstance(v, str))[:600]
            prompt = cit.authorities_block(q, dpath(uid, "mycitations")) + "\n\n" + prompt
        out = []
        for m in [str(x) for x in a.models][:4]:
            if not (m == "free" or m.startswith("claude-")): continue
            t0, use = time.time(), []
            try:
                if m == "free":
                    prov = _assist_providers()
                    if not prov: out.append({"model": "free", "error": "Free AI setup nahi hua (ASSIST_API_KEY)"}); continue
                    base, key, model = prov[0]
                    text, name = _assist_call(base, key, model, SYSTEM, [{"role": "user", "content": prompt}], 4000), "free: " + model
                else:
                    text, name = call_claude(prompt, None, m, lambda mm, i, o: (use.append((i, o)), _note_usage("abtest", mm, i, o))), m
                final, _c, _raw, stats = finalize_text(ar, text, uid, note=False)
                p = PRICES.get(m)
                usd = round(sum(i for i, _ in use) / 1e6 * p[0] + sum(o for _, o in use) / 1e6 * p[1], 4) if (p and use) else (0.0 if m == "free" else None)
                out.append({"model": name, "seconds": round(time.time() - t0, 1), "usd": usd, "red": stats["red"], "yellow": stats["yellow"], "missing": stats["missing"], "text": final})
            except Exception as e:
                out.append({"model": m, "error": f"{type(e).__name__}: {str(e)[:160]}"})
        with _JLOCK: AB_JOBS[jid] = {"status": "done", "results": out, "label": label, "ts": time.time()}
    except Exception as e:
        traceback.print_exc()
        with _JLOCK: AB_JOBS[jid] = {"status": "error", "error": str(getattr(e, "detail", e))[:200], "ts": time.time()}

@app.post("/admin/abtest/start")
def admin_abtest_start(r: AbIn, authorization: str = Header(None)):
    users, uid = _admin_only(authorization)
    if r.kind not in ("draft", "fir", "research", "section") or not r.text.strip(): raise HTTPException(400, "Qism chunein aur matan likhein.")
    jid = secrets.token_hex(6)
    with _JLOCK:
        for k in [k for k, v in AB_JOBS.items() if time.time() - v.get("ts", 0) > 3600]: AB_JOBS.pop(k, None)
        AB_JOBS[jid] = {"status": "pending", "ts": time.time()}
    threading.Thread(target=_ab_job, args=(jid, uid, r), daemon=True).start()
    return {"job": jid}

@app.post("/admin/abtest/job")
def admin_abtest_job(r: JobId, authorization: str = Header(None)):
    _admin_only(authorization)
    with _JLOCK: j = AB_JOBS.get(r.id)
    if not j: raise HTTPException(404, "Ye kaam nahi mila (server dobara chala ho sakta hai).")
    return {k: v for k, v in j.items() if k != "ts"}
# --- DAFA+AB END

@app.post("/admin/cases/of")
def admin_cases_of(r: AdvId, authorization: str = Header(None)):
    """Admin only: every case number of ONE advocate (no client names, no phone numbers)."""
    users, uid = current(authorization)
    if not users[uid].get("is_admin"): raise HTTPException(403, "Sirf admin.")
    u = users.get(r.id)
    if not u or u.get("is_admin"): raise HTTPException(404, "Advocate nahi mila.")
    cases, missing, total = _adv_cases(r.id)
    return {"advocate": str(u.get("chamber_name", "")), "clients": total, "without_number": missing, "cases": cases}

@app.post("/admin/cases/undo")
def admin_cases_undo(authorization: str = Header(None)):
    """Admin: revert the last bulk update, client by client, but never a client the advocate edited after it."""
    users, uid = current(authorization)
    if not users[uid].get("is_admin"): raise HTTPException(403, "Sirf admin.")
    last = rj(f"{DATA}/_last_bulk_update.json", None)
    if not last: raise HTTPException(404, "Wapas karne ke liye koi bulk update nahi mili.")
    restored = skipped = 0
    for b in last.get("batch", []):
        path = f"{DATA}/{b['file']}"; clients = rj(path, [])
        d, s = cs.undo_entries(clients, b["entries"]); restored += d; skipped += s
        if d: wj(path, clients)
    rm(f"{DATA}/_last_bulk_update.json")
    return {"restored": restored, "skipped_because_edited": skipped}

@app.post("/admin/cases/districts")
def admin_cases_districts(authorization: str = Header(None)):
    """Admin: how many clients (with a case number) are in each district, so he knows whose cause list is worth uploading.
    Counts only: no client or advocate names."""
    users, uid = current(authorization)
    if not users[uid].get("is_admin"): raise HTTPException(403, "Sirf admin.")
    counts = {}
    for fn in data_files():
        if not fn.endswith("_clients.json"): continue
        for c in rj(f"{DATA}/{fn}", []):
            if not isinstance(c, dict) or not cs._pairs_of(f"{c.get('fir', '')} {c.get('case_no', '')}"): continue
            d = cs.district_of(c) or (c.get("district") or "").strip().title() or "Ghair maloom"
            e = counts.setdefault(d, [0, set()]); e[0] += 1; e[1].add(fn)
    return {"districts": sorted(({"district": d, "clients": v[0], "advocates": len(v[1])} for d, v in counts.items()), key=lambda x: -x["clients"])}

@app.post("/admin/cases/status")
def admin_cases_status(authorization: str = Header(None)):
    """Admin: how many tracked cases already have a future date, how many are waiting for their next date (hearing passed), how many have none.
    Counts only."""
    import datetime as dt
    users, uid = current(authorization)
    if not users[uid].get("is_admin"): raise HTTPException(403, "Sirf admin.")
    today = dt.datetime.now(dt.timezone(dt.timedelta(hours=5))).date().isoformat()          # Pakistan time
    n = {"tracked": 0, "current": 0, "awaiting": 0, "no_date": 0}
    for fn in data_files():
        if not fn.endswith("_clients.json"): continue
        for c in rj(f"{DATA}/{fn}", []):
            if not isinstance(c, dict) or not cs._pairs_of(f"{c.get('fir', '')} {c.get('case_no', '')}"): continue
            d = str(c.get("date") or ""); n["tracked"] += 1
            n["no_date" if not d else "awaiting" if d < today else "current"] += 1
    return n

BOARD_HTML = r"""<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Advocate AI - Daily Board</title>
<style>
:root{--bg:#f5f7f6;--fg:#14211c;--card:#fff;--line:#c9d6cf;--acc:#0b5d44;--btn:#fff}
@media (prefers-color-scheme:dark){:root{--bg:#0f1a17;--fg:#e6efe9;--card:#16241f;--line:#2c4038;--acc:#3ecf9b;--btn:#06281d}}
[hidden]{display:none!important}
body{margin:0;background:var(--bg);color:var(--fg);font:15px/1.5 system-ui,sans-serif}
main{max-width:560px;margin:0 auto;padding:20px 16px}
.card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:16px;margin-bottom:14px}
input,button,textarea{font:inherit;padding:10px;border-radius:8px;border:1px solid var(--line);width:100%;box-sizing:border-box;margin:6px 0;background:var(--card);color:var(--fg)}
button{background:var(--acc);color:var(--btn);border:0;cursor:pointer;font-weight:600}
button:disabled{opacity:.5;cursor:default}
label.row{display:flex;gap:8px;align-items:center;margin:8px 0}
label.row input{width:auto;margin:0}
pre{white-space:pre-wrap;word-break:break-word;margin:0}
small{opacity:.75}
</style></head><body><main>
<h2>Daily Board upload</h2>
<div class="card" id="s1">
  <b>1) Login</b>
  <input id="id" placeholder="Email" autocomplete="username">
  <input id="pw" type="password" placeholder="Password" autocomplete="current-password">
  <button onclick="login()">Login</button>
  <small>Admin nahi hain to yeh login aap ke kisi purane device ko logout kar sakta hai.</small>
</div>
<div class="card" id="s2" hidden>
  <b>2) Daily Board PDF</b>
  <input id="f" type="file" accept="application/pdf,.pdf">
  <textarea id="txt" rows="5" placeholder="Ya cause list ka text yahan paste karein (PDF na ho to)" style="width:100%;box-sizing:border-box"></textarea>
  <label class="row" id="allw" hidden><input id="all" type="checkbox"> Sab advocates ke clients update karein (admin)</label>
  <label class="row" id="dryw" hidden><input id="dry" type="checkbox" checked> Pehle sirf preview (kuch update nahi hoga)</label>
  <button id="go" onclick="send()">Upload aur update</button>
</div>
<div class="card" id="s3" hidden>
  <b>Districts</b> <small>(admin)</small>
  <button onclick="districts()">Kis district ke kitne clients hain</button>
  <button onclick="status()">Kitni cases ki agli tareekh abhi baqi hai</button>
  <button onclick="undo()">Aakhri bulk update wapas lein</button>
  <pre id="dist"></pre>
</div>
<div class="card" id="s8" hidden style="border:2px solid #c62828">
  <b style="color:#c62828">🔴 Update hone wale cases (admin)</b> <small>- sirf case number, naam/phone nahi</small>
  <button onclick="alerts()">Dobara dekhein</button>
  <div id="al"></div>
</div>
<div class="card" id="s7" hidden>
  <b>AI ka kharch (admin)</b> <small>- sirf tokens ginti hai, kisi ka matan nahi</small>
  <button onclick="usage()">Is mahine ka kharch dikhayein</button>
  <div id="us"></div>
</div>
<div class="card" id="s12" hidden>
  <b>Court Hub data (admin)</b> <small>- links, thana-adalat, court calendar, legal news</small>
  <select id="h1" onchange="hubHint()"><option value="links">Links (District Courts, High Court, Bare Acts, Bar)</option><option value="thana">Thana - Adalat</option><option value="calendar">Court calendar (chhutti)</option><option value="news">Legal news</option></select>
  <textarea id="h2" rows="5" placeholder="" style="width:100%;box-sizing:border-box"></textarea>
  <button onclick="hubImport()">CSV import (pehli line header ki ho)</button>
  <button onclick="hubList()">Fehrist dekhein / hataein</button>
  <div id="hl"></div>
</div>
<div class="card" id="s11" hidden>
  <b>A/B test - teen model ek saath (admin)</b> <small>- Free AI, Haiku 5.5 aur Sonnet 5.5 par ek hi sawal; Claude walon ka kharch lagta hai</small>
  <select id="ab1"><option value="draft">Draft</option><option value="fir">FIR tajziya</option><option value="research">Research</option><option value="section">Section / dafa</option></select>
  <textarea id="ab2" rows="5" placeholder="Draft ke facts / FIR ka text / research ka sawal / jurm ka naam" style="width:100%;box-sizing:border-box"></textarea>
  <input id="ab3" placeholder="Draft ki qism (jaise Post-arrest bail application) - sirf draft ke liye">
  <button onclick="abStart()">Teeno par aazmayein</button>
  <div id="abr"></div>
</div>
<div class="card" id="s10" hidden>
  <b>Dafa Table (admin)</b> <small>- saza, bailable, cognizable wahi jo aap ne Pakistan Code se tasdeeq ki</small>
  <input id="d1" placeholder="Act (PPC, CrPC, CNSA ...)"><input id="d2" placeholder="Section (jaise 302 ya 337-F(i))">
  <input id="d3" placeholder="Dafa ka naam (jaise Qatl-e-amd)"><input id="d4" placeholder="Saza">
  <input id="d5" placeholder="Bailable (Haan / Nahi)"><input id="d6" placeholder="Cognizable (Haan / Nahi)">
  <input id="d7" placeholder="Compoundable"><input id="d8" placeholder="Court of trial">
  <input id="d9" placeholder="Zareya - kahan se tasdeeq hua (zaroori)">
  <button onclick="dafaAdd()">Table mein add karein</button>
  <textarea id="d10" rows="3" placeholder="Bulk CSV: act,section,title,punishment,bailable,cognizable,compoundable,court,notes,source" style="width:100%;box-sizing:border-box"></textarea>
  <button onclick="dafaImport()">Bulk import</button>
  <button onclick="dafaReq()">Kin dafaat ki tafseel chahiye</button>
  <div id="dr"></div>
</div>
<div class="card" id="s9" hidden>
  <b>Official faisla - AI se nikalein</b> <small>- Supreme Court / High Court ka official PDF (text wali)</small>
  <input id="jf" type="file" accept="application/pdf,.pdf">
  <button onclick="judgExtract()">AI se nikalein (Claude Sonnet)</button>
  <div id="jr"></div>
</div>
<div class="card" id="s6" hidden>
  <b>Verified Citation Library (admin)</b> <small>- sirf wahi citation jo aap ne khud tasdeeq kiya</small>
  <input id="c1" placeholder="Citation (jaise 2021 SCMR 1234)"><input id="c2" placeholder="Case ka title (parties)">
  <input id="c3" placeholder="Court"><input id="c4" placeholder="Faisle ki tareekh">
  <input id="c5" placeholder="Zareya - kahan se tasdeeq hua (zaroori)"><input id="c6" placeholder="Link (agar ho)">
  <input id="c12" placeholder="Keywords (English, comma se)"><input id="c13" placeholder="Dafaat / Acts (jaise 497 CrPC)">
  <textarea id="c7" rows="3" placeholder="Khulasa (apne alfaz mein)" style="width:100%;box-sizing:border-box"></textarea>
  <textarea id="c9" rows="3" placeholder="Faisle se asal ibarat (hu-ba-hu, jo usool batati ho)" style="width:100%;box-sizing:border-box"></textarea>
  <input id="c10" placeholder="Para / page number">
  <select id="c11"><option value="">Status: theek (good law)</option><option value="distinguished">Distinguished</option><option value="check">Dobara check karna hai</option><option value="overruled">Overruled</option></select>
  <button onclick="citAdd()">Library mein add karein</button>
  <textarea id="c8" rows="3" placeholder="Bulk: CSV (citation,title,court,date,summary,source,url) ya JSONL" style="width:100%;box-sizing:border-box"></textarea>
  <button onclick="citImport()">Bulk import</button>
  <button onclick="citReq()">Tasdeeq ki darkhwastein</button>
  <button onclick="citGaps()">Library kahan khali hai</button>
  <button onclick="libTodo()">Library ka kaam (aaj kya karna hai)</button>
  <div id="cr"></div>
</div>
<div class="card" id="s5" hidden>
  <b>Advocates ke case numbers (admin)</b> <small>- naam aur phone nahi dikhte</small>
  <input id="cq" placeholder="Case no / advocate / court se dhoondein" oninput="if (CN) drawCN()">
  <button onclick="caseNums()">Fehrist dikhayein</button>
  <div id="cn"></div>
</div>
<div class="card" id="s4" hidden>
  <b>Daily Diary (PDF)</b> <small>- aap ke sab cases</small>
  <input id="dd" type="date">
  <small>Tareekh khali chhorein to aaj ki diary banegi.</small>
  <button onclick="diary()">Diary PDF download karein</button>
</div>
<div class="card" id="out" hidden><pre id="msg"></pre><div id="sg"></div></div>
<script>
var T = null, LAST = null, CN = null, JDRAFT = '';
function $(i){ return document.getElementById(i); }
function show(t){ $('out').hidden = false; $('msg').textContent = t; }
async function api(path, body){
  var h = {'Content-Type': 'application/json'};
  if (T) h['Authorization'] = 'Bearer ' + T;
  var r = await fetch(path, {method: 'POST', headers: h, body: JSON.stringify(body)});
  var d = {};
  try { d = await r.json(); } catch (e) {}
  if (!r.ok) throw new Error(typeof d.detail === 'string' ? d.detail : ('Error ' + r.status));
  return d;
}
async function login(){
  try {
    var d = await api('/auth/login', {identifier: $('id').value.trim(), password: $('pw').value});
    T = d.token; $('s1').hidden = true; $('s2').hidden = false;
    $('allw').hidden = !(d.user && d.user.is_admin);
    $('s3').hidden = $('allw').hidden; $('dryw').hidden = $('allw').hidden; $('s4').hidden = false; $('s5').hidden = $('allw').hidden; $('s6').hidden = $('allw').hidden; $('s7').hidden = $('allw').hidden; $('s8').hidden = $('allw').hidden; $('s9').hidden = $('allw').hidden; $('s10').hidden = $('allw').hidden; $('s11').hidden = $('allw').hidden; $('s12').hidden = $('allw').hidden; hubHint(); if (!$('allw').hidden) alerts();
    show('Login ho gaya. Ab PDF chunein.');
  } catch (e) { show(e.message); }
}
var HUB_HINT = {
  links: 'section,district,title,url,note,daily\ndistrict,*,Naam,https://...,note,1     (section: district / highcourt / bareact / bar; district: * ya zile ka naam)',
  thana: 'district,thana,court,tehsil,notes,source\nSialkot,Thana ka naam,Court ka naam,Tehsil,,Zareya',
  calendar: 'start,end,title,kind,source\n2026-12-25,2026-12-25,Chhutti ka naam,holiday,Notification (kind: holiday / vacation / other)',
  news: 'date,title,url,source\n2026-11-24,Khabar ki sarkhi,https://...,Akhbar ka naam'
};
function hubHint(){ $('h2').placeholder = HUB_HINT[$('h1').value]; }
async function hubImport(){
  try { var d = await api('/admin/hub/import', {name: $('h1').value, text: $('h2').value}); show('Add: ' + d.added + ', chhode: ' + d.skipped + '. Kul ' + d.total + (d.errors.length ? ' | ' + d.errors.join('; ') : '')); }
  catch (e) { show(e.message); }
}
async function hubList(){
  try {
    var d = await api('/admin/hub/list', {name: $('h1').value}), box = $('hl'); box.textContent = '';
    if (!d.rows.length) { box.textContent = 'Abhi koi entry nahi.'; return; }
    d.rows.forEach(function(r, i){
      var row = document.createElement('div'); row.style.cssText = 'margin-top:8px;font-size:13px';
      row.appendChild(document.createTextNode(d.columns.map(function(c){ return r[c]; }).filter(Boolean).join(' | ') + '  '));
      var b = document.createElement('button'); b.textContent = 'Hataein'; b.style.cssText = 'width:auto;padding:4px 10px';
      b.onclick = async function(){ try { await api('/admin/hub/delete', {name: $('h1').value, index: i}); hubList(); } catch (e) { show(e.message); } };
      row.appendChild(b); box.appendChild(row);
    });
  } catch (e) { show(e.message); }
}
async function dafaAdd(){
  try {
    var d = await api('/admin/dafa/add', {act: $('d1').value, section: $('d2').value, title: $('d3').value, punishment: $('d4').value, bailable: $('d5').value, cognizable: $('d6').value, compoundable: $('d7').value, court: $('d8').value, source: $('d9').value});
    ['d2', 'd3', 'd4', 'd5', 'd6', 'd7', 'd8'].forEach(function(i){ $(i).value = ''; });
    show('Table mein add ho gaya. Ab ' + d.rows + ' dafaat hain.');
  } catch (e) { show(e.message); }
}
async function dafaImport(){
  try { var d = await api('/admin/dafa/import', {text: $('d10').value}); show('Add: ' + d.added + ', chhode: ' + d.skipped + '. Ab ' + d.rows + ' dafaat hain.'); }
  catch (e) { show(e.message); }
}
async function dafaReq(){
  try {
    var d = await api('/admin/dafa/requests', {}), box = $('dr'); box.textContent = 'Table mein abhi ' + d.rows + ' dafaat hain.';
    if (!d.requests.length) { box.appendChild(document.createElement('br')); box.appendChild(document.createTextNode('Koi darkhwast nahi.')); return; }
    d.requests.forEach(function(x){
      var b = document.createElement('button'); b.textContent = (x.act || '?') + ' ' + x.section + '  (' + x.count + ' dafa)';
      b.onclick = function(){ $('d1').value = x.act; $('d2').value = x.section; }; box.appendChild(b);
    });
  } catch (e) { show(e.message); }
}
async function abStart(){
  try {
    var s = await api('/admin/abtest/start', {kind: $('ab1').value, text: $('ab2').value, doc_type: $('ab3').value || 'Petition', language: 'English'}), box = $('abr');
    box.textContent = 'Teeno model chal rahe hain... 1 se 3 minute intezar karein.';
    for (var i = 0; i < 120; i++) {
      await new Promise(function(r){ setTimeout(r, 3000); });
      var j = await api('/admin/abtest/job', {id: s.job});
      if (j.status === 'pending') continue;
      if (j.status === 'error') throw new Error(j.error);
      box.textContent = '';
      j.results.forEach(function(x){
        var h = document.createElement('div'); h.style.cssText = 'margin-top:14px;font-weight:700';
        h.textContent = x.error ? (x.model + ' - ' + x.error) : (x.model + ' | ' + x.seconds + ' sec | ' + (x.usd === null ? 'qeemat maloom nahi' : '$' + x.usd) + ' | laal ' + x.red + ', peela ' + x.yellow + ', chhoota ' + x.missing);
        box.appendChild(h);
        if (x.text) { var p = document.createElement('pre'); p.style.cssText = 'white-space:pre-wrap;font-size:12px'; p.textContent = x.text; box.appendChild(p); }
      });
      return;
    }
    throw new Error('Bohat der lag gayi.');
  } catch (e) { show(e.message); }
}
async function judgExtract(){
  var f = $('jf').files[0]; if (!f) { show('Pehle faisle ka PDF chunein.'); return; }
  try {
    var s = await api('/admin/judgment/extract', {pdf_b64: b64(await f.arrayBuffer())}), box = $('jr');
    box.textContent = 'AI faisla parh raha hai (' + s.pages + ' page). 1 se 2 minute intezar karein...';
    for (var i = 0; i < 90; i++) {
      await new Promise(function(r){ setTimeout(r, 2500); });
      var j = await api('/admin/judgment/job', {id: s.job});
      if (j.status === 'pending') continue;
      if (j.status === 'error') throw new Error(j.error);
      return judgFill(j);
    }
    throw new Error('Bohat der lag gayi. Dobara koshish karein.');
  } catch (e) { show(e.message); }
}
function judgFill(j){
  var d = j.draft, box = $('jr'); box.textContent = ''; JDRAFT = j.draft_id;
  $('c1').value = d.printed_citation || (d.case_ref ? d.case_ref + (d.court ? ' (' + d.court + ')' : '') : '');
  $('c2').value = d.title || ''; $('c3').value = d.court || ''; $('c4').value = d.decision_date || '';
  $('c5').value = (d.court || 'Official') + ' - official faisla (AI se nikala, admin ne tasdeeq kiya)';
  $('c7').value = d.summary || ''; var e0 = (d.extracts || [])[0] || {}; $('c9').value = e0.text || ''; $('c10').value = e0.para || '';
  $('c12').value = (d.keywords || []).join(', '); $('c13').value = (d.statutes || []).join(', ');
  var h = document.createElement('div'); h.style.fontWeight = '700'; h.textContent = 'AI ne nikala - ab har cheez khud dekhein. Neeche server ki jaanch:'; box.appendChild(h);
  (j.checks || []).forEach(function(c){ var x = document.createElement('div'); x.style.color = c.ok ? '#2e7d32' : '#c62828'; x.textContent = (c.ok ? '✅ ' : '❌ ') + c.field + ': ' + c.note; box.appendChild(x); });
  (d.extracts || []).slice(1).forEach(function(e, i){
    var b = document.createElement('button'); b.textContent = 'Ibarat ' + (i + 2) + ' (para ' + (e.para || '?') + ') form mein lagayein';
    b.onclick = function(){ $('c9').value = e.text; $('c10').value = e.para || ''; }; box.appendChild(b);
  });
  var n = document.createElement('div'); n.style.fontSize = '12px';
  n.textContent = 'Ab neeche "Verified Citation Library" ka form dekhein (bhara hua hai), theek karein, status chunein, aur "Library mein add karein" dabayein. Citation mein agar PLD/SCMR ka hawala ho to wahan likh dein.';
  box.appendChild(n);
}
async function citAdd(){
  try {
    await api('/admin/citations/add', {citation: $('c1').value, title: $('c2').value, court: $('c3').value, date: $('c4').value, source: $('c5').value, url: $('c6').value, summary: $('c7').value, extract: $('c9').value, para: $('c10').value, status: $('c11').value, keywords: $('c12').value, sections: $('c13').value, draft_id: JDRAFT});
    ['c1', 'c2', 'c4', 'c6', 'c7', 'c9', 'c10', 'c12', 'c13'].forEach(function(i){ $(i).value = ''; }); $('c11').value = ''; JDRAFT = ''; $('jr').textContent = '';
    show('Library mein add ho gaya.');
  } catch (e) { show(e.message); }
}
async function citImport(){
  try { var d = await api('/citations/import', {text: $('c8').value, source: $('c5').value}); show('Add: ' + d.added + ', chhode: ' + d.skipped + (d.library_full ? ' (library full)' : '')); }
  catch (e) { show(e.message); }
}
async function alerts(){
  try {
    var d = await api('/admin/alerts', {}), box = $('al'); box.textContent = '';
    var h = document.createElement('div'); h.style.cssText = 'font-weight:700;color:' + (d.red ? '#c62828' : '#2e7d32');
    h.textContent = d.red ? ('🔴 ' + d.red + ' cases: peshi aaj hai ya guzar chuki (agli tareekh nahi mili)') : 'Koi red alert nahi';
    box.appendChild(h);
    var s = document.createElement('div'); s.textContent = '🟠 Kal ki peshi: ' + d.orange + ' | Bina case number: ' + d.no_case_no + ' | 30 din se purani (alert mein nahi): ' + d.older_than_30_days; box.appendChild(s);
    d.groups.forEach(function(g){
      var t = document.createElement('div'); t.style.cssText = 'margin-top:12px;font-weight:700';
      t.textContent = g.group + ' - aaj ' + g.aaj + ', guzri ' + g.guzri + ', kal ' + g.kal; box.appendChild(t);
      var hn = document.createElement('div'); hn.style.fontSize = '12px'; hn.textContent = g.hint; box.appendChild(hn);
      g.cases.forEach(function(c){
        var x = document.createElement('div'); x.style.color = c.state === 'kal' ? '#ef6c00' : '#c62828';
        x.textContent = (c.state === 'kal' ? '🟠 ' : '🔴 ') + (c.case_no || '(case no nahi)') + (c.fir ? ' | FIR ' + c.fir : '') + ' | ' + c.advocate + ' | ' + c.court + ' | ' + c.date + ' | ' + c.state;
        box.appendChild(x);
      });
    });
  } catch (e) { show(e.message); }
}
async function libTodo(){
  try {
    var d = await api('/admin/library/todo', {}), box = $('cr'); box.textContent = '';
    var h = document.createElement('div'); h.style.fontWeight = '700'; h.textContent = 'Library mein abhi ' + d.library_entries + ' entries hain'; box.appendChild(h);
    d.tasks.forEach(function(t){
      var a = document.createElement('div'); a.style.cssText = 'margin-top:12px;font-weight:700'; a.textContent = t.priority + '. ' + t.title; box.appendChild(a);
      var b = document.createElement('div'); b.style.fontSize = '13px'; b.textContent = t.how; box.appendChild(b);
      t.items.forEach(function(i){ var x = document.createElement('div'); x.textContent = '- ' + i; box.appendChild(x); });
    });
  } catch (e) { show(e.message); }
}
async function usage(){
  try {
    var d = await api('/admin/usage', {}), box = $('us'); box.textContent = '';
    var h = document.createElement('div'); h.textContent = 'Default model: ' + d.models_in_use.default + ' | halke kaam (' + d.models_in_use.light_kinds.join(', ') + '): ' + d.models_in_use.light_tasks; box.appendChild(h);
    var ms = Object.keys(d.months); if (!ms.length) { box.appendChild(document.createTextNode('Abhi koi AI istemal darj nahi.')); return; }
    ms.forEach(function(mo){
      var t = document.createElement('div'); t.style.cssText = 'margin-top:12px;font-weight:700'; t.textContent = mo + ' - kul takreeban $' + d.months[mo].total_usd; box.appendChild(t);
      d.months[mo].rows.forEach(function(r){ var x = document.createElement('div'); x.textContent = r.kind + ' | ' + r.model + ' | ' + r.calls + ' dafa | in ' + r.input_tokens + ' out ' + r.output_tokens + ' | ' + (r.usd === null ? 'qeemat maloom nahi' : '$' + r.usd); box.appendChild(x); });
    });
  } catch (e) { show(e.message); }
}
async function citGaps(){
  try {
    var d = await api('/admin/citations/gaps', {}), box = $('cr'); box.textContent = '';
    if (!d.gaps.length) { box.textContent = 'Abhi koi khala nahi mila.'; return; }
    d.gaps.forEach(function(g){ var x = document.createElement('div'); x.textContent = g.topic + ' - ' + g.count + ' dafa library mein kuch nahi mila'; box.appendChild(x); });
  } catch (e) { show(e.message); }
}
async function citReq(){
  try {
    var d = await api('/admin/citations/requests', {}), box = $('cr'); box.textContent = '';
    if (!d.requests.length) { box.textContent = 'Koi darkhwast nahi.'; return; }
    d.requests.forEach(function(x){
      var b = document.createElement('button'); b.textContent = x.citation + '  (' + x.count + ' dafa)';
      b.onclick = function(){ $('c1').value = x.citation; };
      box.appendChild(b);
    });
  } catch (e) { show(e.message); }
}
async function caseNums(){
  try { CN = await api('/admin/cases/numbers', {}); drawCN(); } catch (e) { show(e.message); }
}
function drawCN(){
  var q = $('cq').value.trim().toLowerCase(), box = $('cn'); box.textContent = '';
  CN.advocates.forEach(function(a){
    var rows = a.cases.filter(function(c){ return !q || (a.advocate + ' ' + c.case_no + ' ' + c.fir + ' ' + c.court).toLowerCase().indexOf(q) >= 0; });
    if (!rows.length) return;
    var h = document.createElement('div'); h.style.cssText = 'margin-top:14px;font-weight:700';
    h.textContent = a.advocate + ' - ' + rows.length + ' case(s), ' + a.clients + ' client(s)' + (a.without_number ? ', ' + a.without_number + ' ka number likha nahi' : '');
    box.appendChild(h);
    rows.forEach(function(c){ var d = document.createElement('div'); d.textContent = [c.case_no, c.fir, c.court, c.date, c.stage].filter(Boolean).join(' | '); box.appendChild(d); });
  });
  if (!box.firstChild) box.textContent = 'Kuch nahi mila.';
}
async function diary(){
  try {
    var r = await fetch('/diary/pdf', {method: 'POST', headers: {'Content-Type': 'application/json', 'Authorization': 'Bearer ' + T}, body: JSON.stringify({date: $('dd').value})});
    if (!r.ok) { var e = {}; try { e = await r.json(); } catch (x) {} throw new Error(e.detail || ('Error ' + r.status)); }
    var b = await r.blob(), a = document.createElement('a');
    a.href = URL.createObjectURL(b); a.download = 'Daily_Diary_' + ($('dd').value || 'today') + '.pdf';
    document.body.appendChild(a); a.click(); a.remove();
    show('Diary download ho gayi.');
  } catch (e) { show(e.message); }
}
async function districts(){
  try {
    var d = await api('/admin/cases/districts', {});
    $('dist').textContent = d.districts.length ? d.districts.map(function(x){ return x.district + ': ' + x.clients + ' clients (' + x.advocates + ' advocates)'; }).join('\n') : 'Abhi koi client nahi.';
  } catch (e) { show(e.message); }
}
async function status(){
  try {
    var d = await api('/admin/cases/status', {});
    $('dist').textContent = d.tracked + ' cases track ho rahi hain\n' + d.current + ' ki agli tareekh aage ki hai\n' + d.awaiting + ' ki peshi guzar gayi, agli tareekh abhi nahi mili\n' + d.no_date + ' par koi tareekh nahi';
  } catch (e) { show(e.message); }
}
async function undo(){
  if (!confirm('Aakhri bulk update wapas liya jaye?')) return;
  try {
    var d = await api('/admin/cases/undo', {});
    $('dist').textContent = d.restored + ' clients purani halat par wapas, ' + d.skipped_because_edited + ' chhor diye (advocate ne baad mein khud badal diye).';
  } catch (e) { show(e.message); }
}
function b64(buf){
  var s = '', a = new Uint8Array(buf);
  for (var i = 0; i < a.length; i += 32768) s += String.fromCharCode.apply(null, a.subarray(i, i + 32768));
  return btoa(s);
}
async function send(){
  var f = $('f').files[0], t = $('txt').value.trim();
  if (!f && !t) { show('Pehle PDF chunein ya cause list ka text paste karein.'); return; }
  $('go').disabled = true; $('sg').textContent = ''; show('PDF parh raha hoon...');
  try {
    var all = $('all').checked && !$('allw').hidden;
    var body = f ? {pdf_b64: b64(await f.arrayBuffer())} : {raw: t};
    LAST = body;
    if (all) body.dry = $('dry').checked;
    var d = await api(all ? '/admin/cases/ingest' : '/cases/ingest', body);
    var t = d.rows + ' cases parhi gayin.\n';
    if (d.rows && d.dated === false) t += 'Board par tareekh nahi mili, is liye sirf stage update ho sakti hai.\n';
    if (all) {
      if (d.preview) t = 'PREVIEW (kuch badla nahi): ' + t;
      t += d.advocates_checked + ' advocates check hue, ' + d.clients_updated + ' clients update hue, ' + d.needs_review + ' ki tasdeeq baqi (woh advocate apne account se karega).';
    } else {
      t += d.updated.length + ' client update hue, ' + d.suggestions.length + ' ki tasdeeq darkar, ' + d.not_found_total + ' ka case no is list mein nahi mila.';
      d.updated.forEach(function(u){ t += '\n- ' + u.client + (u.case_no ? ' (' + u.case_no + ')' : '') + ': ' + u.changes.join(', '); });
      d.suggestions.forEach(function(s){
        var div = document.createElement('div'); div.style.marginTop = '10px';
        var sp = document.createElement('span'); sp.textContent = s.client + (s.case_no ? ' (' + s.case_no + ')' : '') + ' ? ' + s.why + ' | ' + s.preview;
        var b = document.createElement('button'); b.textContent = 'Haan, update';
        b.onclick = async function(){
          try { await api('/cases/confirm', {client_index: s.client_index, row: s.row}); b.disabled = true; b.textContent = 'Ho gaya'; }
          catch (e) { show(e.message); }
        };
        div.appendChild(sp); div.appendChild(b); $('sg').appendChild(div);
      });
      (d.not_found || []).forEach(function(n){
        var div = document.createElement('div'); div.style.marginTop = '14px';
        var sp = document.createElement('div'); sp.textContent = n.client + ' - is list mein nahi mila';
        var inp = document.createElement('input'); inp.value = n.case_no || ''; inp.placeholder = 'Case no (jaise 2633798/26)';
        var b = document.createElement('button'); b.textContent = 'Is number se dhoondein';
        var res = document.createElement('div');
        b.onclick = async function(){
          res.textContent = '';
          try {
            var q = {case_no: inp.value.trim()};
            if (LAST.pdf_b64) q.pdf_b64 = LAST.pdf_b64; else q.raw = LAST.raw;
            var m = await api('/cases/find', q);
            if (!m.matches.length) { res.textContent = 'Is number ka case is list mein nahi hai.'; return; }
            m.matches.forEach(function(x){
              var r = document.createElement('div'); r.style.marginTop = '6px';
              var s2 = document.createElement('span'); s2.textContent = x.preview;
              var b2 = document.createElement('button'); b2.textContent = 'Haan, ' + n.client + ' ko update karein';
              b2.onclick = async function(){
                try { await api('/cases/confirm', {client_index: n.client_index, row: x.row, case_no: inp.value.trim()}); b2.disabled = true; b2.textContent = 'Ho gaya'; }
                catch (e) { show(e.message); }
              };
              r.appendChild(s2); r.appendChild(b2); res.appendChild(r);
            });
          } catch (e) { res.textContent = e.message; }
        };
        div.appendChild(sp); div.appendChild(inp); div.appendChild(b); div.appendChild(res); $('sg').appendChild(div);
      });
    }
    show(t);
  } catch (e) { show(e.message); }
  $('go').disabled = false;
}
</script>
</main></body></html>
"""

@app.get("/board")
def board_page():
    """Upload page for a Daily Board PDF from any PC browser (login, upload, confirm)."""
    return Response(BOARD_HTML, media_type="text/html; charset=utf-8", headers={"Cache-Control": "no-store"})

# ---------------------------------------------------------------- admin
@app.get("/admin/users")
def admin_users(authorization: str = Header(None)):
    users, uid = current(authorization)
    if not users[uid].get("is_admin"): raise HTTPException(403)
    out = []
    for k, v in users.items():
        if k == "admin": continue
        _, missing, total = _adv_cases(k)
        out.append({"id": k, **profile(v), "devices": len(v.get("sessions", {})), "email_verified": bool(v.get("email_verified")),
                    "name": v.get("name", ""), "phone": v.get("phone", ""), "city": v.get("city", ""), "bar_no": v.get("bar_no", ""),
                    "created": v.get("created", ""), "cases": total, "no_case_no": missing})
    return out

class AdminSet(BaseModel): id: str; package: str | None = None; expiry_date: str | None = None; add_days: int = 0; disabled: bool | None = None; email_verified: bool | None = None

@app.post("/admin/update")
def admin_update(r: AdminSet, authorization: str = Header(None)):
    users, uid = current(authorization)
    if not users[uid].get("is_admin"): raise HTTPException(403)
    v = users.get(r.id)
    if not v: raise HTTPException(404)
    if r.package in DAILY_LIMITS: v["package"] = r.package
    if r.disabled is not None:
        v["disabled"] = r.disabled
        if r.disabled: v["sessions"] = {}          # band karte hi sab devices logout
    if r.email_verified is not None: v["email_verified"] = r.email_verified
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
        rm(dpath(r.id, col))
    return {"ok": True}
