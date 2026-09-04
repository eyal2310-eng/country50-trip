"""
בודק אוטומטי למסמכי Word בעברית.
מריץ שתי קבוצות בדיקות: טכניות (RTL ב-XML) ולשוניות (מלכודות bidi בטקסט).

שימוש:  python3 check_hebrew.py <file.docx> [file2.docx ...]
יציאה 0 = הכל תקין. יציאה 1 = נמצאו כשלים.
"""
import re, sys, zipfile

HEB = re.compile(r"[֐-׿]")
LATIN = re.compile(r"[A-Za-z]")
# משפט עברי שמסתיים בספרה / סוגר / תו לועזי -> הנקודה תקפוץ
BAD_END = re.compile(r"[֐-׿].*?[0-9)\]A-Za-z%]\s*[.,:;!?]+\s*$")
RLM = "\u200f"


def unprotected_end(x):
    """סיום בעייתי רק אם אין תו RLM שמגן על סימן הפיסוק."""
    return bool(BAD_END.search(x)) and RLM not in x[-8:]
PLACEHOLDER = re.compile(r"\[למילוי\]|\{\{|XXX|TODO|LOREM")
BANNED = {
    "קונטרקט": "הסכם/חוזה", "קנסלציה": "ביטול", "רמדי": "סעד",
    "נוטיס": "הודעה מוקדמת", "אנקס": "נספח", "אפנדיקס": "נספח",
    "קלאוזה": "תניה/סעיף",
}

def paragraphs(xml):
    return re.findall(r"<w:p[ >].*?</w:p>", xml, re.S)

def text_of(p):
    return "".join(re.findall(r"<w:t[^>]*>(.*?)</w:t>", p, re.S))

def check(path):
    z = zipfile.ZipFile(path)
    doc = z.read("word/document.xml").decode("utf-8")
    styles = z.read("word/styles.xml").decode("utf-8")
    fails, warns, passes = [], [], []

    def t(ok, label, detail=""):
        (passes if ok else fails).append(f"{label}{(' — ' + detail) if detail else ''}")

    # --- טכני ---
    t("<w:bidi" in re.search(r"<w:sectPr.*?</w:sectPr>", doc, re.S).group(0),
      "מקטע (section) מוגדר RTL")
    t("<w:bidi" in styles and "<w:rtl" in styles,
      "סגנון Normal מוגדר RTL", "בלעדיו כל פסקה חדשה תיווצר LTR")
    t('w:cs=' in styles, "גופן Complex Script מוגדר בסגנון",
      "בלעדיו העברית לא מקבלת את הגופן שהוגדר")

    heb_paras = [p for p in paragraphs(doc) if HEB.search(text_of(p))]
    no_bidi = [text_of(p)[:40] for p in heb_paras if "<w:bidi" not in p]
    t(not no_bidi, f"כל {len(heb_paras)} הפסקאות העבריות מוגדרות bidi",
      f"{len(no_bidi)} ללא: {no_bidi[:3]}")

    runs_no_rtl = 0
    for p in heb_paras:
        for r in re.findall(r"<w:r[ >].*?</w:r>", p, re.S):
            if HEB.search("".join(re.findall(r"<w:t[^>]*>(.*?)</w:t>", r, re.S))) \
               and "<w:rtl" not in r:
                runs_no_rtl += 1
    t(runs_no_rtl == 0, "כל הריצות (runs) העבריות מוגדרות rtl",
      f"{runs_no_rtl} ריצות ללא w:rtl")

    n_tbl = doc.count("<w:tbl>")
    if n_tbl:
        t(doc.count("<w:bidiVisual") >= n_tbl,
          f"כל {n_tbl} הטבלאות מוגדרות bidiVisual", "סדר העמודות יתהפך")

    # --- לשוני ---
    texts = [text_of(p).strip() for p in paragraphs(doc)]
    texts = [x for x in texts if x]

    bad_ends = [x for x in texts if unprotected_end(x)]
    t(not bad_ends, "כל סיומי המשפטים מוגנים מקפיצת פיסוק (RLM)",
      f"{len(bad_ends)}: {bad_ends[:2]}")

    ph = [x for x in texts if PLACEHOLDER.search(x)]
    t(not ph, "אין שדות שלא מולאו", f"{ph[:2]}")

    found = {w: r for w, r in BANNED.items() if any(w in x for x in texts)}
    t(not found, "אין מונחים לועזיים שיש להם מקבילה עברית", str(found))

    # מספור רציף
    nums = [int(m.group(1)) for x in texts
            if (m := re.match(r"^(\d+)\.\s", x))]
    t(nums == list(range(1, len(nums) + 1)) if nums else True,
      f"מספור הסעיפים רציף ({len(nums)} סעיפים)", f"נמצא: {nums}")

    # לועזית בלי בידוד -> אזהרה בלבד
    mixed = [x for x in texts if HEB.search(x) and LATIN.search(x)]
    if mixed:
        warns.append(f"{len(mixed)} פסקאות מערבבות עברית ולועזית — ודא run/span נפרד")

    return fails, warns, passes


rc = 0
for path in sys.argv[1:]:
    fails, warns, passes = check(path)
    name = path.split("/")[-1]
    print(f"\n{'='*60}\n{name}\n{'='*60}")
    for p in passes: print(f"  ✓ {p}")
    for w in warns: print(f"  ⚠ {w}")
    for f in fails: print(f"  ✗ {f}")
    print(f"  → {len(passes)} עברו, {len(fails)} נכשלו, {len(warns)} אזהרות")
    if fails: rc = 1
sys.exit(rc)
