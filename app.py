import os, io, re, json, hashlib, logging, unicodedata, threading, time, math
import streamlit as st, psycopg, pymupdf as fitz, pandas as pd
from psycopg.rows import dict_row
from PIL import Image
st.set_page_config("Adaptive AI Study Engine", "🎓", layout="wide")
log = logging.getLogger("app")
def cfg(k, d=""):
    try: v = st.secrets.get(k)
    except Exception: v = None
    return str(v or os.environ.get(k) or d)
MAXS = int(cfg("MAX_SUBJECTS", 5)); LIM = float(cfg("DATABASE_STORAGE_LIMIT_MB", 450))
MAXUP = int(cfg("MAX_UPLOAD_MB", 50)); MINTXT = int(cfg("MIN_TEXT_THRESHOLD", 200))
GM = cfg("GEMINI_MODEL", "gemini-2.5-flash"); GQ = cfg("GROQ_REASONING_MODEL", "openai/gpt-oss-120b")
TYPES = ["Syllabus", "Previous Year Question Paper", "Notes", "Lecture Slides", "Assignment", "Reference Material", "Other"]

# ---------- database ----------
@st.cache_resource
def conn(): return psycopg.connect(cfg("DATABASE_URL"), autocommit=True, row_factory=dict_row)
def q(sql, a=None, one=False):
    for t in (0, 1):
        try:
            with conn().cursor() as cur:
                cur.execute(sql, a)
                if not cur.description: return None
                return cur.fetchone() if one else cur.fetchall()
        except psycopg.OperationalError:
            conn.clear()
            if t: raise
@st.cache_resource
def init_db():
    import glob
    base = os.path.dirname(os.path.abspath(__file__))
    fs = glob.glob(os.path.join(base, "**", "schema.sql"), recursive=True)
    if fs: q(open(fs[0]).read())
    q("ALTER TABLE questions ADD COLUMN IF NOT EXISTS figure TEXT")
    return True
def used_mb(): return q("SELECT pg_database_size(current_database()) s", one=True)["s"] / 1048576
if not cfg("DATABASE_URL"): st.error("DATABASE_URL is not configured. See .env.example."); st.stop()
try: init_db()
except Exception as e: st.error(f"Database connection failed: {e}"); st.stop()

# ---------- text utils ----------
def norm(s):
    s = str(s)
    for _ in range(3):
        s = re.sub(r"\\[dt]?frac\{([^{}]*)\}\{([^{}]*)\}", r"(\1)/(\2)", s)
        s = re.sub(r"\\sqrt\{([^{}]*)\}", r"sqrt(\1)", s)
    s = re.sub(r"\\int_\{?([^\s{}^]*)\}?\^\{?([^\s{}]*)\}?", r"Integral from \1 to \2 of ", s)
    s = re.sub(r"\\sum(_\{?[^\s{}^]*\}?)?(\^\{?[^\s{}]*\}?)?", "Sum of ", s)
    s = re.sub(r"\\(begin|end)\{[^}]*\}", "", s)
    s = re.sub(r"\\(cdot|times)", "*", s).replace("\\left", "").replace("\\right", "")
    s = s.replace("$", "").replace("^2", "²").replace("^3", "³").replace("{", "(").replace("}", ")")
    s = re.sub(r"\\([A-Za-z]+)", r"\1", s).replace("\\", "")
    return s
def pdf_safe(s):
    out = ""
    for ch in norm(s):
        if ord(ch) < 256: out += ch
        elif ch in "→⇒": out += "->"
        elif ch == "≤": out += "<="
        elif ch == "≥": out += ">="
        elif ch == "≈": out += "~"
        else:
            try: out += unicodedata.name(ch).split()[-1].lower() if "GREEK" in unicodedata.name(ch) else "?"
            except ValueError: out += "?"
    return out
def clean_pages(pages):
    cnt = {}
    for p in pages:
        for l in set(x.strip() for x in p.splitlines() if 0 < len(x.strip()) < 80): cnt[l] = cnt.get(l, 0) + 1
    rep = {l for l, c in cnt.items() if len(pages) >= 4 and c > len(pages) * 0.5}
    out = []
    for p in pages:
        t = "\n".join(l.rstrip() for l in p.replace("\t", " ").splitlines() if l.strip() not in rep)
        out.append(re.sub(r"\n{3,}", "\n\n", re.sub(r"[ ]{2,}", " ", t)).strip())
    return out
def chunk(text, n=1200):
    out, cur = [], ""
    for para in re.split(r"\n\s*\n|\n(?=Q\s?\d)", text):
        if len(cur) + len(para) > n and cur: out.append(cur); cur = ""
        cur += para + "\n"
    if cur.strip(): out.append(cur)
    return [c.strip() for c in out if c.strip()]

# ---------- extraction ----------
def extract(name, data):
    ext = name.rsplit(".", 1)[-1].lower(); pages, m = [], ""
    if ext == "pdf":
        d = fitz.open(stream=data, filetype="pdf"); pages = [p.get_text() for p in d]; m = "PyMuPDF"
        if sum(len(p.strip()) for p in pages) < MINTXT:
            import pytesseract; m = "Tesseract OCR"; pages = []
            for i, p in enumerate(d):
                try: pages.append(pytesseract.image_to_string(Image.open(io.BytesIO(p.get_pixmap(dpi=200).tobytes("png")))))
                except Exception as e: log.error(e); pages.append(""); st.warning(f"OCR failed on page {i+1}; continuing.")
    elif ext == "pptx":
        from pptx import Presentation; m = "python-pptx"
        for s in Presentation(io.BytesIO(data)).slides:
            t = []
            for sh in s.shapes:
                if sh.has_text_frame: t.append(sh.text_frame.text)
                if getattr(sh, "has_table", False) and sh.has_table:
                    t += [" | ".join(c.text for c in r.cells) for r in sh.table.rows]
            pages.append("\n".join(t))
    elif ext == "docx":
        import docx; m = "python-docx"; d = docx.Document(io.BytesIO(data))
        pages = ["\n".join([p.text for p in d.paragraphs] + [" | ".join(c.text for c in r.cells) for t in d.tables for r in t.rows])]
    else: m = "Text"; pages = [data.decode("utf-8", "ignore")]
    return clean_pages(pages), m
def ingest(sid, name, data, dtype):
    if not data: raise ValueError("File is empty.")
    if len(data) > MAXUP * 1048576: raise ValueError(f"File exceeds {MAXUP} MB limit.")
    if name.rsplit(".", 1)[-1].lower() not in ("pdf", "pptx", "docx", "txt"): raise ValueError("Unsupported file type.")
    h = hashlib.sha256(data).hexdigest()
    if q("SELECT 1 FROM documents WHERE subject_id=%s AND sha256=%s", (sid, h)): raise ValueError("This file already exists for this subject.")
    cur = used_mb()
    if cur >= LIM: raise ValueError(f"Database storage limit reached ({cur:.0f}/{LIM:.0f} MB). Delete material or raise the limit.")
    pages, m = extract(name, data); text_len = sum(len(p) for p in pages)
    if text_len < 20: raise ValueError("No text could be extracted from this file.")
    est = text_len * 3 / 1048576  # pages + chunks + tsvector index
    if cur + est > LIM: raise ValueError(f"Upload rejected: database storage limit would be exceeded. Current {cur:.1f} MB + estimated {est:.1f} MB > limit {LIM:.0f} MB.")
    did = q("INSERT INTO documents(subject_id,filename,doc_type,file_type,file_size,method,sha256) VALUES(%s,%s,%s,%s,%s,%s,%s) RETURNING id",
            (sid, name, dtype, name.rsplit(".", 1)[-1].lower(), len(data), m, h), one=True)["id"]
    for i, p in enumerate(pages, 1):
        q("INSERT INTO document_pages(document_id,subject_id,page_no,content) VALUES(%s,%s,%s,%s)", (did, sid, i, p))
        for c in chunk(p): q("INSERT INTO document_chunks(document_id,subject_id,page_no,content) VALUES(%s,%s,%s,%s)", (did, sid, i, c))
    return f"{name}: processed with {m} ({len(pages)} pages/slides)."

# ---------- retrieval (always filtered by subject_id) ----------
def search(sid, text, k=6):
    w = list(dict.fromkeys(x.lower() for x in re.findall(r"[A-Za-z0-9]+", text) if len(x) > 2))[:12]; r = []
    if w:
        tq = " | ".join(w)
        r = q("SELECT c.content,d.filename,c.page_no FROM document_chunks c JOIN documents d ON d.id=c.document_id WHERE c.subject_id=%s AND c.tsv @@ to_tsquery('english',%s) ORDER BY ts_rank(c.tsv,to_tsquery('english',%s)) DESC LIMIT %s", (sid, tq, tq, k))
    return r or q("SELECT c.content,d.filename,c.page_no FROM document_chunks c JOIN documents d ON d.id=c.document_id WHERE c.subject_id=%s ORDER BY c.id DESC LIMIT %s", (sid, k))
def ctx(sid, dtypes, limit):
    r = q("SELECT c.content FROM document_chunks c JOIN documents d ON d.id=c.document_id WHERE c.subject_id=%s AND d.doc_type = ANY(%s) ORDER BY c.id", (sid, dtypes))
    if not r: return ""
    tot = sum(len(x["content"]) for x in r)
    if tot > limit: r = r[::math.ceil(tot / limit)]
    return "\n".join(x["content"] for x in r)[:limit]

# ---------- AI ----------
def gem(prompt, parts=()):
    from google import genai; from google.genai import types
    c = genai.Client(api_key=cfg("GEMINI_API_KEY"))
    return c.models.generate_content(model=GM, contents=[types.Part.from_bytes(data=d, mime_type=m) for d, m in parts] + [prompt]).text
def reason(prompt, js=False, eff="low"):
    from groq import Groq
    kw = {"response_format": {"type": "json_object"}} if js else {}
    if "gpt-oss" in GQ: kw["reasoning_effort"] = eff
    last = None
    for i in range(5):  # Groq only; wait and retry when busy / rate-limited
        try: return Groq(api_key=cfg("GROQ_API_KEY"), timeout=90).chat.completions.create(model=GQ, messages=[{"role": "user", "content": prompt}], **kw).choices[0].message.content
        except Exception as e: last = e; log.warning(f"Groq retry {i}: {e}"); time.sleep(4 * 2 ** i)
    raise RuntimeError(f"Groq is busy or unavailable after several retries: {last}")
def jparse(s): s = re.sub(r"```json|```", "", s).strip(); return json.loads(s[s.find("{"):s.rfind("}") + 1])
def topic_id(sid, name):
    name = name.strip()[:100]
    r = q("SELECT id FROM topics WHERE subject_id=%s AND lower(name)=lower(%s)", (sid, name), one=True)
    return r["id"] if r else q("INSERT INTO topics(subject_id,name) VALUES(%s,%s) RETURNING id", (sid, name), one=True)["id"]
def extract_topics(sid):
    t = ctx(sid, ["Syllabus", "Notes", "Lecture Slides", "Previous Year Question Paper"], 9000)
    if not t: return
    known = [r["name"] for r in q("SELECT name FROM topics WHERE subject_id=%s", (sid,))]
    o = jparse(reason('List the main examinable topics in this study material as JSON {"topics":["..."]}, 10-30 short names. Existing topics (reuse names, no duplicates): ' + str(known) + "\n\n" + t, True))
    for n in o.get("topics", []):
        if isinstance(n, str) and n.strip(): topic_id(sid, n)
def mastery(sid): return q("SELECT t.id,t.name,COALESCE(m.attempts,0) attempts,COALESCE(m.score,0) score,COALESCE(m.max_score,0) mx,m.last_attempt FROM topics t LEFT JOIN topic_mastery m ON m.topic_id=t.id WHERE t.subject_id=%s ORDER BY t.name", (sid,))
def pct(r): return 100 * r["score"] / r["mx"] if r["mx"] else None
def grey(sid, n=5): return [r["name"] for r in sorted([r for r in mastery(sid) if pct(r) is not None and pct(r) < 60], key=pct)][:n]

# ---------- mock test ----------
FIGS = 'Optional "figure" (null unless a diagram truly helps): {"kind":"neural_net","layers":[3,4,2]} | {"kind":"matrix","rows":[[1,2],[3,4]],"label":"A"} | {"kind":"blocks","boxes":["Input","Conv","ReLU"]} | {"kind":"plot","x":[0,1,2],"y":[0,1,4],"xlabel":"","ylabel":""} | {"kind":"polar","values":[1,3,1,0.5,2,0.5],"label":"RCS pattern"}'
def fig_draw(sp):
    from reportlab.graphics.shapes import Drawing, Circle, Line, String, Rect, PolyLine, Polygon
    from reportlab.lib import colors
    k = sp["kind"]; W, H = 380, 170; d = Drawing(W, H); T = lambda x, y, t, **a: d.add(String(x, y, pdf_safe(str(t)), fontSize=a.pop("fs", 9), **a))
    if k == "neural_net":
        L = [min(int(x), 6) for x in sp["layers"]][:6]; xs = [40 + i * (W - 80) / max(len(L) - 1, 1) for i in range(len(L))]
        P = [[(xs[i], H / 2 + (j - (m - 1) / 2) * 28) for j in range(m)] for i, m in enumerate(L)]
        for a, b in zip(P, P[1:]):
            for p in a:
                for r in b: d.add(Line(p[0], p[1], r[0], r[1], strokeColor=colors.grey, strokeWidth=.5))
        for i, l in enumerate(P):
            for p in l: d.add(Circle(p[0], p[1], 8, fillColor=colors.white, strokeColor=colors.black))
            T(xs[i], 6, "Input" if i == 0 else "Output" if i == len(P) - 1 else "Hidden", textAnchor="middle", fs=8)
    elif k == "matrix":
        R = [[str(v) for v in r][:8] for r in sp["rows"][:8]]; cw, ch = 42, 22; w = cw * len(R[0]); x0 = (W - w) / 2; y0 = H / 2 + ch * len(R) / 2
        for i, r in enumerate(R):
            for j, v in enumerate(r): T(x0 + j * cw + cw / 2, y0 - (i + 1) * ch + 6, v, textAnchor="middle", fs=11)
        h = ch * len(R)
        for x, dx in ((x0 - 4, 5), (x0 + w + 4, -5)):
            d.add(Line(x, y0, x, y0 - h, strokeWidth=1.2)); d.add(Line(x, y0, x + dx, y0)); d.add(Line(x, y0 - h, x + dx, y0 - h))
        T(x0 - 20, y0 - h / 2 - 4, sp.get("label", "") + " =", textAnchor="end", fs=12)
    elif k == "blocks":
        B = [str(b) for b in sp["boxes"]][:6]; bw = (W - 20) / len(B) - 14
        for i, b in enumerate(B):
            x = 10 + i * (bw + 14); d.add(Rect(x, H / 2 - 18, bw, 36, fillColor=colors.whitesmoke, strokeColor=colors.black)); T(x + bw / 2, H / 2 - 3, b[:14], textAnchor="middle", fs=8)
            if i: d.add(Line(x - 14, H / 2, x, H / 2))
    elif k == "plot":
        xs_, ys_ = [float(v) for v in sp["x"]], [float(v) for v in sp["y"]]; x0, y0, w, h = 45, 30, W - 70, H - 55
        fx = lambda v: x0 + (v - min(xs_)) / ((max(xs_) - min(xs_)) or 1) * w; fy = lambda v: y0 + (v - min(ys_)) / ((max(ys_) - min(ys_)) or 1) * h
        d.add(Line(x0, y0, x0 + w, y0)); d.add(Line(x0, y0, x0, y0 + h))
        d.add(PolyLine([c for a, b in zip(xs_, ys_) for c in (fx(a), fy(b))], strokeColor=colors.blue))
        T(x0 + w / 2, 8, sp.get("xlabel", ""), textAnchor="middle"); T(4, y0 + h + 6, sp.get("ylabel", ""))
        T(x0, y0 - 11, f"{min(xs_):g}", textAnchor="middle", fs=7); T(x0 + w, y0 - 11, f"{max(xs_):g}", textAnchor="middle", fs=7)
    elif k == "polar":
        V = [abs(float(v)) for v in sp["values"]][:72]; m = max(V) or 1; cx, cy, R = W / 2, H / 2, 65
        for r in (R / 3, 2 * R / 3, R): d.add(Circle(cx, cy, r, fillColor=None, strokeColor=colors.lightgrey))
        d.add(Polygon([c for i, v in enumerate(V) for c in (cx + R * v / m * math.cos(2 * math.pi * i / len(V)), cy + R * v / m * math.sin(2 * math.pi * i / len(V)))], fillColor=None, strokeColor=colors.red))
        T(cx, 4, sp.get("label", ""), textAnchor="middle")
    else: raise ValueError("unknown figure")
    return d
QT = ["mcq", "theory", "numerical", "derivation", "case_study"]
def qtype(s):
    s = str(s).lower()
    return "mcq" if ("mcq" in s or "multiple" in s) else "case_study" if "case" in s else "derivation" if "deriv" in s else "numerical" if "numer" in s else "theory"
def gen_test(sid, n, diff, topics, pyq, use_grey, plan=None, instr="", prog=lambda m: None):
    ms = mastery(sid); names = topics or [r["name"] for r in ms]; weak = grey(sid) if use_grey else []
    prof = "; ".join(f"{r['name']} ({'untested' if pct(r) is None else str(round(pct(r)))+'% mastery'})" for r in ms if r["name"] in names)
    src = "SYLLABUS:\n" + ctx(sid, ["Syllabus"], 1500) + "\n\nSTUDENT NOTES (ask direct or closely related questions):\n" + ctx(sid, ["Notes", "Reference Material"], 3000) + "\n\nCLASS SLIDES (ask direct or closely related questions):\n" + ctx(sid, ["Lecture Slides", "Assignment"], 2500)
    p_ = ctx(sid, ["Previous Year Question Paper"], 3500) if pyq else ""
    need_t = {k: int(v) for k, v in plan.items() if int(v) > 0} if plan else None
    if need_t: n = sum(need_t.values())
    good, seen = [], set()
    for _ in range(8):
        need = n - len(good)
        if need <= 0: break
        prog(f"Generating questions… {len(good)}/{n} done")
        if need_t:
            ask, tot = {}, 0
            for k, v in need_t.items():
                t_ = min(v, 10 - tot)
                if t_ > 0: ask[k] = t_; tot += t_
            spec = f"Create exactly these NEW questions (question_type must be one of numerical, derivation, mcq, case_study, theory): {ask}. "
        else: spec = f"Create exactly {min(need, 10)} NEW questions of mixed types. "
        o = reason("Generate engineering examination questions using the supplied syllabus, notes, slides and PYQs as the primary source. Do not invent unsupported topics. Prefer the terminology of the material. Use plain engineering notation; NEVER output LaTeX (\\frac, \\sqrt, \\sum, \\begin). Use forms like 10 / [s(s+2)], sqrt(x² + y²).\n"
                   + spec + f"Difficulty: {diff}. Allowed topics: {names}. Topic profile: {prof}. " + (f"Weak topics needing MORE questions: {weak}. " if weak else "")
                   + ("Numericals should follow the topics/subtopics seen in the PYQs. Ask questions similar or related to the PYQs (same style, new numbers/wording), never copies. " if pyq else "")
                   + "MCQs: put the 4 options (A)-(D) on separate lines inside the question text. Case studies: a short scenario then sub-parts with marks. Numerical data must be self-consistent; solve each numerical yourself before writing expected_answer. "
                   + ("MANDATORY USER CONSTRAINTS (follow strictly): " + instr + " " if instr.strip() else "") + "Already created (avoid): " + str([g['text'][:50] for g in good]) + "\n"
                   + 'Return JSON only: {"questions":[{"question":"","marks":5,"difficulty":"easy|medium|hard","topic":"","question_type":"","expected_answer":"","figure":null}]}. ' + FIGS + "\n\n" + src + "\n\nPYQs:\n" + p_, True, "medium")
        try: items = jparse(o)["questions"]
        except Exception as e: log.error(e); continue
        left = dict(need_t) if need_t else None
        for it in items:
            try:
                t = norm(it["question"]).strip(); mk = int(it["marks"]); tp = str(it["topic"]).strip(); qt = qtype(it.get("question_type"))
                m_ = next((x for x in names if x.lower() == tp.lower()), None)
                if len(t) < 15 or mk <= 0 or (names and not m_) or (left is not None and left.get(qt, 0) <= 0): continue
                k = re.sub(r"\W", "", t.lower())[:80]
                if k in seen: continue
                fg = it.get("figure")
                try: fig_draw(fg); fg = json.dumps(fg)
                except Exception: fg = None
                seen.add(k); good.append(dict(text=t, marks=mk, difficulty=str(it.get("difficulty", diff)).lower(), topic=m_ or tp, qtype=qt, solution=norm(it.get("expected_answer", "")), figure=fg))
                if left is not None: left[qt] -= 1; need_t[qt] -= 1
                if len(good) >= n: break
            except Exception: continue
    if len(good) < n: raise RuntimeError(f"Only {len(good)} of {n} valid questions could be generated. Try again in a few minutes.")
    if plan: good.sort(key=lambda g: QT.index(g["qtype"]))
    prog("Saving test…"); sname = q("SELECT name FROM subjects WHERE id=%s", (sid,), one=True)["name"]
    tid = q("INSERT INTO tests(subject_id,title,difficulty,question_count,topics) VALUES(%s,%s,%s,%s,%s) RETURNING id", (sid, f"{sname} Mock Test", diff, n, ", ".join(sorted({g['topic'] for g in good}))), one=True)["id"]
    for i, g in enumerate(good, 1):
        qid = q("INSERT INTO questions(subject_id,test_id,num,text,marks,difficulty,topic,qtype,solution,figure) VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s) RETURNING id", (sid, tid, i, g["text"], g["marks"], g["difficulty"], g["topic"], g["qtype"], g["solution"], g["figure"]), one=True)["id"]
        q("INSERT INTO test_questions VALUES(%s,%s,%s)", (tid, qid, i))
    return tid
@st.cache_resource
def jobs(): return {}
def start_job(sid, *a):
    j = {"state": "running", "msg": "Starting…", "tid": None, "err": None}; jobs()[sid] = j; conn()
    def run():
        try: j["tid"] = gen_test(sid, *a, prog=lambda m: j.update(msg=m)); j["state"] = "done"
        except Exception as e: log.exception(e); j["err"] = str(e); j["state"] = "error"
    threading.Thread(target=run, daemon=True).start()
def apply_plan():
    p = st.session_state.plan; c = p["counts"]
    st.session_state.update(g_mode="Custom blueprint", g_num=c["numerical"], g_der=c["derivation"], g_mcq=c["mcq"], g_case=c["case_study"], g_theory=c["theory"], g_instr=p["instr"], g_diff=p["diff"] if p["diff"] in ("Mixed", "Easy", "Medium", "Hard") else "Mixed")

def make_pdf(tid, key=False):
    from reportlab.lib.pagesizes import A4; from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer; from xml.sax.saxutils import escape
    t = q("SELECT t.*,s.name sname FROM tests t JOIN subjects s ON s.id=t.subject_id WHERE t.id=%s", (tid,), one=True)
    qs = q("SELECT * FROM questions WHERE test_id=%s ORDER BY num", (tid,)); tot = sum(x["marks"] for x in qs)
    P = lambda s: escape(pdf_safe(s)).replace("\n", "<br/>")
    ss = getSampleStyleSheet(); b = io.BytesIO(); d = SimpleDocTemplate(b, pagesize=A4, leftMargin=50, rightMargin=50, topMargin=50, bottomMargin=50)
    el = [Paragraph("ANSWER KEY" if key else "ADAPTIVE MOCK TEST", ss["Title"]),
          Paragraph(P(f"Subject: {t['sname']}\nDuration: {round(tot * 1.2)} Minutes\nMaximum Marks: {tot}"), ss["Normal"]), Spacer(1, 8)]
    if not key: el += [Paragraph("<b>Instructions:</b><br/>1. Attempt all questions.<br/>2. Show necessary steps for numerical problems.<br/>3. Assume suitable values where required.", ss["Normal"])]
    el.append(Spacer(1, 12))
    for x in qs:
        el.append(Paragraph(P(f"Q{x['num']}. {x['text']} [{x['marks']}]"), ss["Normal"])); el.append(Spacer(1, 4))
        if x.get("figure"):
            try: el.append(fig_draw(json.loads(x["figure"])))
            except Exception as e: log.error(e)
        if key: el.append(Paragraph("<i>" + P("Answer: " + (x["solution"] or "-")) + "</i>", ss["Normal"]))
        el.append(Spacer(1, 10))
    d.build(el); data = b.getvalue()
    txt = "".join(p.get_text() for p in fitz.open(stream=data, filetype="pdf"))  # verification
    if not txt.strip() or re.search(r"\\(frac|sqrt|begin|end|sum|int)\b", txt): raise RuntimeError("PDF verification failed.")
    return data

# ---------- evaluation ----------
def evaluate(sid, aid, qrow, ans):
    o = jparse(reason(f"You are a strict but fair engineering examiner. Question ({qrow['marks']} marks): {qrow['text']}\nReference solution: {qrow['solution']}\nStudent answer: {ans or '(blank)'}\n"
        'Check concept, formula, procedure, final answer, units, missing steps, partial credit. Plain text, no LaTeX. Return JSON: {"score":0,"correct":"","missing":"","mistake":"","correct_approach":"","improvement":"","steps":["Step 1: Given values...","Step 2: Formula used...","Step 3: Substitution...","Step 4: Calculation...","Step 5: Final answer...","Step 6: Comparison with student\'s answer..."]}', True, "medium"))
    sc = max(0.0, min(float(o.get("score", 0)), qrow["marks"]))
    eid = q("INSERT INTO evaluations(attempt_answer_id,score,max_score,feedback) VALUES(%s,%s,%s,%s) RETURNING id", (aid, sc, qrow["marks"], json.dumps(o)), one=True)["id"]
    for i, s in enumerate(o.get("steps", []), 1): q("INSERT INTO evaluation_steps(evaluation_id,step_no,text) VALUES(%s,%s,%s)", (eid, i, norm(s)))
    tid = topic_id(sid, qrow["topic"])
    q("INSERT INTO topic_mastery(topic_id,subject_id,attempts,score,max_score,last_attempt) VALUES(%s,%s,1,%s,%s,now()) ON CONFLICT(topic_id) DO UPDATE SET attempts=topic_mastery.attempts+1,score=topic_mastery.score+EXCLUDED.score,max_score=topic_mastery.max_score+EXCLUDED.max_score,last_attempt=now()", (tid, sid, sc, qrow["marks"]))
    return sc, o

# ---------- sidebar ----------
subs = q("SELECT * FROM subjects ORDER BY id")
if not subs:
    st.title("Welcome to Adaptive AI Study Engine"); st.subheader("Create your first subject")
    n = st.text_input("Subject name")
    if st.button("Create Subject") and n.strip(): q("INSERT INTO subjects(name) VALUES(%s)", (n.strip(),)); st.rerun()
    st.stop()
names = {s["id"]: s["name"] for s in subs}
st.sidebar.title("Adaptive AI Study Engine")
sid = st.sidebar.selectbox("Subject", list(names), format_func=lambda i: names[i])
with st.sidebar.expander("Manage subjects"):
    st.caption(f"Subjects: {len(subs)} / {MAXS}")
    nn = st.text_input("New subject name", key="nn")
    if st.button("Add subject"):
        if len(subs) >= MAXS: st.error(f"Maximum of {MAXS} subjects reached.")
        elif nn.strip():
            try: q("INSERT INTO subjects(name) VALUES(%s)", (nn.strip(),)); st.rerun()
            except Exception: st.error("A subject with that name exists.")
    rn = st.text_input("Rename current subject to", key="rn")
    if st.button("Rename") and rn.strip(): q("UPDATE subjects SET name=%s WHERE id=%s", (rn.strip(), sid)); st.rerun()
    if st.checkbox("Confirm: delete current subject and ALL its data") and st.button("Delete subject"):
        q("DELETE FROM subjects WHERE id=%s", (sid,)); st.rerun()
used = used_mb(); st.sidebar.progress(min(used / LIM, 1.0)); st.sidebar.caption(f"Storage {used:.0f} / {LIM:.0f} MB ({used / LIM * 100:.0f}%)")
pg = st.sidebar.radio("Pages", ["🏠 Dashboard", "📚 Course Material", "📝 Generate Mock Test", "✍️ Take Test", "📊 Performance", "🤖 Test Planner", "💾 Storage", "⚙️ Settings"])
st.header(f"{names[sid]}")

_j = jobs().get(sid)
if _j: st.sidebar.info("⏳ Mock test generating…" if _j["state"] == "running" else "✅ Mock test ready — open 📝 Generate Mock Test" if _j["state"] == "done" else "⚠️ Generation failed — see 📝 page")
for _k in ["g_n", "g_diff", "g_sel", "g_pyq", "g_gr", "g_mode", "g_num", "g_der", "g_mcq", "g_case", "g_theory", "g_instr", "plan", "conv"]:  # keep form values when switching pages
    if _k in st.session_state: st.session_state[_k] = st.session_state[_k]
# ---------- pages ----------
if pg.startswith("🏠"):
    a = q("SELECT count(*) n, COALESCE(avg(score/NULLIF(max_score,0))*100,0) av FROM attempts WHERE subject_id=%s", (sid,), one=True)
    c = st.columns(5)
    for col, (l, v) in zip(c, [("Documents", q("SELECT count(*) n FROM documents WHERE subject_id=%s", (sid,), one=True)["n"]), ("Topics", q("SELECT count(*) n FROM topics WHERE subject_id=%s", (sid,), one=True)["n"]),
                               ("Tests Taken", a["n"]), ("Average Score", f"{float(a['av']):.0f}%"), ("Storage", f"{used:.0f}/{LIM:.0f} MB")]): col.metric(l, v)
    st.subheader("Grey Areas"); [st.write("• " + g) for g in grey(sid)] or None
    if not grey(sid): st.caption("No weak areas yet — take a test.")

elif pg.startswith("📚"):
    dt = st.radio("Document Type", TYPES, horizontal=True)
    fs = st.file_uploader("Upload Material (PDF, PPTX, TXT, DOCX)", type=["pdf", "pptx", "txt", "docx"], accept_multiple_files=True)
    if fs and st.button("Process files"):
        ok = False
        for f in fs:
            with st.spinner(f.name):
                try: st.success(ingest(sid, f.name, f.getvalue(), dt)); ok = True
                except ValueError as e: st.error(f"Cannot upload {f.name}: {e}")
                except Exception as e: log.exception(e); st.error(f"Upload failed for {f.name}: {e}")
        if ok:
            try: extract_topics(sid)
            except Exception as e: st.warning(f"Topic extraction failed (Gemini): {e}")
    if st.button("Extract topics now"):
        try: extract_topics(sid); st.success("Topics updated.")
        except Exception as e: st.error(f"Topic extraction failed: {e}")
    docs = q("SELECT * FROM documents WHERE subject_id=%s ORDER BY id DESC", (sid,))
    st.subheader("Your material, by type"); st.caption(" · ".join(f"{t}: {sum(1 for d in docs if d['doc_type'] == t)}" for t in TYPES))
    for t in TYPES:
        grp = [d for d in docs if d["doc_type"] == t]
        if not grp: continue
        st.markdown(f"### {t} ({len(grp)})")
        for d in grp:
            with st.expander(f"{d['filename']} — {d['file_size']/1048576:.1f} MB — {d['method']} — {d['created_at']:%Y-%m-%d} — {d['status']}"):
                txt = "\n".join(f"--- Page/Slide {p['page_no']} ---\n{p['content']}" for p in q("SELECT page_no,content FROM document_pages WHERE document_id=%s AND subject_id=%s ORDER BY page_no", (d["id"], sid)))
                st.text(txt[:1500]); c1, c2, c3 = st.columns(3)
                c1.download_button("Download TXT", txt, f"{d['filename'].rsplit('.',1)[0]}_extracted.txt", key=f"dl{d['id']}")
                nt = c2.selectbox("Change type", TYPES, index=TYPES.index(t), key=f"ty{d['id']}")
                if nt != t: q("UPDATE documents SET doc_type=%s WHERE id=%s AND subject_id=%s", (nt, d["id"], sid)); st.rerun()
                if c3.button("Delete", key=f"dd{d['id']}"): q("DELETE FROM documents WHERE id=%s AND subject_id=%s", (d["id"], sid)); st.rerun()
    with st.expander("Topics"):
        for r in q("SELECT id,name FROM topics WHERE subject_id=%s ORDER BY name", (sid,)):
            a, b = st.columns([6, 1]); a.write(r["name"])
            if b.button("✕", key=f"t{r['id']}"): q("DELETE FROM topics WHERE id=%s AND subject_id=%s", (r["id"], sid)); st.rerun()

elif pg.startswith("📝"):
    for k, v in {"g_n": 20, "g_diff": "Mixed", "g_pyq": True, "g_gr": True, "g_mode": "Random", "g_num": 8, "g_der": 4, "g_mcq": 6, "g_case": 2, "g_theory": 2, "g_instr": ""}.items(): st.session_state.setdefault(k, v)
    if st.session_state.get("g_sid") != sid: st.session_state.g_sel = []; st.session_state.g_sid = sid
    tps = [r["name"] for r in q("SELECT name FROM topics WHERE subject_id=%s ORDER BY name", (sid,))]
    mode = st.radio("Mode", ["Random", "Custom blueprint"], horizontal=True, key="g_mode")
    c = st.columns(2); diff = c[0].selectbox("Difficulty", ["Mixed", "Easy", "Medium", "Hard"], key="g_diff"); sel = c[1].multiselect("Topics (empty = all)", tps, key="g_sel")
    plan, instr = None, ""
    if mode == "Random": n = st.number_input("Number of questions", 1, 60, key="g_n")
    else:
        cc = st.columns(5); plan = {"numerical": cc[0].number_input("Numericals", 0, 30, key="g_num"), "derivation": cc[1].number_input("Derivations", 0, 30, key="g_der"), "mcq": cc[2].number_input("MCQs", 0, 30, key="g_mcq"), "case_study": cc[3].number_input("Case studies", 0, 30, key="g_case"), "theory": cc[4].number_input("Theory", 0, 30, key="g_theory")}
        n = sum(plan.values()); st.caption(f"Total questions: {n}")
        instr = st.text_area("Special instructions (or build them in 🤖 Test Planner)", key="g_instr", height=100, placeholder="e.g. CNN numerical must be on backpropagation; 3 numericals strictly on linear filtering")
    pyq = st.checkbox("Include PYQ style", key="g_pyq"); gr = st.checkbox("Focus on grey areas", key="g_gr")
    if st.button("Generate", disabled=bool(jobs().get(sid, {}).get("state") == "running")):
        if not q("SELECT 1 FROM documents WHERE subject_id=%s LIMIT 1", (sid,)): st.error("Upload course material first.")
        elif n < 1: st.error("Choose at least one question.")
        else: st.session_state.pop("gerr", None); start_job(sid, int(n), diff, sel, pyq, gr, plan, instr); st.rerun()
    @st.fragment(run_every=3)
    def poll():
        j = jobs().get(sid)
        if not j: return
        if j["state"] == "running": st.info("⏳ " + j["msg"] + " — you can switch pages; generation continues in the background and retries automatically if Groq is busy.")
        else:
            if j["state"] == "done": st.session_state.last = j["tid"]
            else: st.session_state.gerr = j["err"]
            jobs().pop(sid, None); st.rerun()
    poll()
    if st.session_state.get("gerr"): st.error(st.session_state.gerr)
    for t in q("SELECT * FROM tests WHERE subject_id=%s ORDER BY id DESC", (sid,)):
        with st.expander(f"#{t['id']} {t['title']} — {t['question_count']} Qs — {t['difficulty']} — {t['created_at']:%Y-%m-%d}", expanded=st.session_state.get("last") == t["id"]):
            st.caption("Topics: " + (t["topics"] or ""))
            try:
                c1, c2, c3 = st.columns(3)
                c1.download_button("Download Mock Test PDF", make_pdf(t["id"]), f"mock_test_{t['id']}.pdf", "application/pdf", key=f"p{t['id']}")
                c2.download_button("Download Answer Key", make_pdf(t["id"], True), f"answer_key_{t['id']}.pdf", "application/pdf", key=f"k{t['id']}")
            except Exception as e: st.error(f"PDF error: {e}")
            if c3.button("Delete test", key=f"x{t['id']}"): q("DELETE FROM tests WHERE id=%s AND subject_id=%s", (t["id"], sid)); st.rerun()

elif pg.startswith("✍️"):
    ts = q("SELECT * FROM tests WHERE subject_id=%s ORDER BY id DESC", (sid,))
    if not ts: st.info("Generate a test first."); st.stop()
    t = st.selectbox("Test", ts, format_func=lambda t: f"#{t['id']} {t['title']} ({t['question_count']} Qs)")
    qs = q("SELECT * FROM questions WHERE test_id=%s AND subject_id=%s ORDER BY num", (t["id"], sid)); ans = {}
    for tab, x in zip(st.tabs([f"Q{x['num']}" for x in qs]), qs):
        with tab:
            st.write(f"**Q{x['num']}.** {x['text']}  **[{x['marks']}]**")
            tx = st.text_area("Typed answer", key=f"a{t['id']}_{x['id']}", height=150)
            up = st.file_uploader("…or handwritten answer (PNG/JPG/PDF)", type=["png", "jpg", "jpeg", "pdf"], key=f"u{t['id']}_{x['id']}")
            ans[x["id"]] = (tx, up)
    if st.button("Submit test", type="primary"):
        with st.spinner("Transcribing and evaluating..."):
            aid = q("INSERT INTO attempts(subject_id,test_id,score,max_score) VALUES(%s,%s,0,%s) RETURNING id", (sid, t["id"], sum(x["marks"] for x in qs)), one=True)["id"]; tot = 0; res = []
            for x in qs:
                tx, up = ans[x["id"]]; text = tx or ""
                try:
                    if up is not None:
                        mime = "application/pdf" if up.name.lower().endswith("pdf") else ("image/png" if up.name.lower().endswith("png") else "image/jpeg")
                        text = gem("Transcribe this handwritten engineering answer faithfully in plain text (no LaTeX). Mark unclear parts as [unclear]. End with 'CONFIDENCE: NN%'.", [(up.getvalue(), mime)])
                    ansid = q("INSERT INTO attempt_answers(attempt_id,question_id,answer_text) VALUES(%s,%s,%s) RETURNING id", (aid, x["id"], text), one=True)["id"]
                    if up is not None: q("INSERT INTO transcriptions(attempt_answer_id,filename,text) VALUES(%s,%s,%s)", (ansid, up.name, text))
                    sc, o = evaluate(sid, ansid, x, text); tot += sc; res.append((x, text, sc, o))
                except Exception as e: log.exception(e); st.error(f"Q{x['num']} could not be evaluated: {e}")
            q("UPDATE attempts SET score=%s WHERE id=%s", (tot, aid)); st.session_state.res = (tot, sum(x["marks"] for x in qs), res)
    if "res" in st.session_state:
        tot, mx, res = st.session_state.res; st.success(f"Total: {tot:.1f} / {mx}")
        for x, text, sc, o in res:
            with st.expander(f"Q{x['num']} — Score {sc:g}/{x['marks']} — {x['topic']}"):
                for k, l in [("correct", "What you did correctly"), ("missing", "What is missing"), ("mistake", "Mistake"), ("correct_approach", "Correct approach"), ("improvement", "Suggested improvement")]: st.markdown(f"**{l}:** {norm(o.get(k, ''))}")
                for s in o.get("steps", []): st.write(norm(s))
                if text: st.caption("Your answer / transcription: " + text[:500])

elif pg.startswith("📊"):
    ms = mastery(sid); df = pd.DataFrame([{"Topic": r["name"], "Attempts": r["attempts"], "Mastery %": round(pct(r)) if pct(r) is not None else None} for r in ms])
    st.subheader("Topic mastery"); st.dataframe(df, use_container_width=True)
    if not df.empty and df["Mastery %"].notna().any(): st.bar_chart(df.dropna().set_index("Topic")["Mastery %"])
    st.subheader("Grey areas"); st.write(", ".join(grey(sid)) or "None yet")
    at = pd.DataFrame(q("SELECT a.id,t.title,a.score,a.max_score,a.created_at FROM attempts a JOIN tests t ON t.id=a.test_id WHERE a.subject_id=%s ORDER BY a.id DESC", (sid,)))
    st.subheader("Attempts"); st.dataframe(at, use_container_width=True)
    if not at.empty: st.download_button("Performance CSV", at.to_csv(index=False), "performance.csv")

elif pg.startswith("🤖"):
    st.subheader("Test Planner")
    st.caption("Describe the mock test you want (counts per type, topic rules). I turn it into a blueprint for the generator. I don't teach concepts here — use your notes for that.")
    if st.button("New plan"): st.session_state.conv = None; st.session_state.plan = None
    cid = st.session_state.get("conv"); hist = q("SELECT role,content FROM tutor_messages WHERE conversation_id=%s ORDER BY id", (cid,)) if cid else []
    for m in hist: st.chat_message(m["role"]).write(m["content"])
    if msg := st.chat_input("e.g. 20 questions: 8 numericals, 4 derivations, 6 MCQs, 2 case studies, 2 theory…"):
        st.chat_message("user").write(msg)
        if not cid: cid = st.session_state.conv = q("INSERT INTO tutor_conversations(subject_id,title) VALUES(%s,%s) RETURNING id", (sid, msg[:50]), one=True)["id"]
        q("INSERT INTO tutor_messages(conversation_id,role,content) VALUES(%s,'user',%s)", (cid, msg))
        tp = [r["name"] for r in q("SELECT name FROM topics WHERE subject_id=%s", (sid,))]
        try:
            o = jparse(reason(f"You are a mock-test planning assistant for '{names[sid]}'. You do NOT teach concepts; if asked to explain something, say to use the notes and steer back to planning the test. Turn the student's request into a blueprint, keeping and updating the current blueprint: {st.session_state.get('plan')}. Known topics: {tp}. Recent chat: " + " | ".join(f"{m['role']}: {m['content'][:200]}" for m in hist[-6:]) + f". New request: {msg}\n"
                              'Return JSON: {"reply":"short confirmation of counts and constraints; ask if something is missing","counts":{"numerical":0,"derivation":0,"mcq":0,"case_study":0,"theory":0},"instructions":"all topic-specific constraints as clear sentences, accumulated","difficulty":"Easy|Medium|Hard|Mixed"}', True))
            c = {k: int(o.get("counts", {}).get(k, 0) or 0) for k in ("numerical", "derivation", "mcq", "case_study", "theory")}
            st.session_state.plan = {"counts": c, "instr": str(o.get("instructions", "")), "diff": str(o.get("difficulty", "Mixed")).capitalize()}; r = norm(o.get("reply", "Plan updated."))
        except Exception as e: log.exception(e); r = f"Planner failed: {e}"
        st.chat_message("assistant").write(r); q("INSERT INTO tutor_messages(conversation_id,role,content) VALUES(%s,'assistant',%s)", (cid, r))
    pl = st.session_state.get("plan")
    if pl:
        st.info(f"**Blueprint:** {pl['counts']} (total {sum(pl['counts'].values())}), difficulty {pl['diff']}\n\n{pl['instr']}")
        st.button("Use this in Generate Mock Test", on_click=apply_plan)

elif pg.startswith("💾"):
    st.metric("Database storage", f"{used:.0f} / {LIM:.0f} MB"); st.progress(min(used / LIM, 1.0)); st.write(f"{used / LIM * 100:.0f}% used")
    st.dataframe(pd.DataFrame(q("SELECT s.name, round((COALESCE((SELECT sum(length(content)) FROM document_chunks WHERE subject_id=s.id),0)*2 + COALESCE((SELECT sum(length(content)) FROM document_pages WHERE subject_id=s.id),0))/1048576.0,2) AS approx_mb FROM subjects s ORDER BY 1")))
    st.caption("Per-subject figures are approximate (extracted text + chunks).")

else:
    st.write({"Gemini model": GM, "Groq model": GQ, "Max subjects": MAXS, "Storage limit (MB)": LIM, "Max upload (MB)": MAXUP, "OCR threshold (chars)": MINTXT})
    st.subheader("Diagnostics")
    st.write("Database connection: ✅"); st.write("Gemini key: " + ("✅" if cfg("GEMINI_API_KEY") else "❌")); st.write("Groq key: " + ("✅" if cfg("GROQ_API_KEY") else "❌"))
    try: import pytesseract; st.write(f"Tesseract: ✅ {pytesseract.get_tesseract_version()}")
    except Exception: st.write("Tesseract: ❌")
    try: import reportlab; st.write("PDF generation: ✅")
    except Exception: st.write("PDF generation: ❌")
    st.write(f"Storage: {used:.0f}/{LIM:.0f} MB — Subjects: {len(subs)}/{MAXS}")
    if st.button("Ping Gemini"):
        try: gem("Say OK"); st.success("Gemini OK")
        except Exception as e: st.error(f"Gemini request failed: {e}")
    if st.button("Ping Groq"):
        try: from groq import Groq; Groq(api_key=cfg("GROQ_API_KEY")).chat.completions.create(model=GQ, messages=[{"role": "user", "content": "OK"}]); st.success("Groq OK")
        except Exception as e: st.error(f"Groq reasoning service unavailable: {e}")
