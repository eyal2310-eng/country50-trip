# RTL — הגדרות מדויקות לכל פורמט

הפניה נכונה של המסמך אינה "יישור לימין". יישור לימין הוא עיצוב; RTL הוא כיוון
הזרימה של הטקסט. בלי הגדרת כיוון אמיתית, מספרים, סוגריים, נקודות וקטעים באנגלית
יקפצו למקום הלא נכון — וזה בדיוק מה שהופך מסמך משפטי ללא-מקצועי.

---

## DOCX (python-docx)

Word דורש **שלוש** הגדרות נפרדות. חסרה אחת — המסמך יראה שבור.

| רמה | תג XML | מה קורה בלעדיו |
|---|---|---|
| מקטע (section) | `w:bidi` ב-`sectPr` | סדר העמודות והשוליים הפוך |
| פסקה | `w:bidi` ב-`pPr` | הנקודה בסוף המשפט קופצת שמאלה |
| ריצה (run) | `w:rtl` ב-`rPr` | מספרים ומילים לועזיות מתהפכות |

בנוסף, לגופן עברי חייבים להגדיר את **Complex Script**: `w:rFonts/@w:cs` וגודל
`w:szCs`. הגדרת `run.font.name` לבדה לא משפיעה על עברית.

```python
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

FONT = "David"
SIZE = 12

def _el(tag, **attrs):
    e = OxmlElement(tag)
    for k, v in attrs.items():
        e.set(qn(k), v)
    return e

def rtl_document(doc, font=FONT, size=SIZE):
    """מגדיר את ברירת המחדל של המסמך כולו ל-RTL. להריץ פעם אחת, מיד אחרי Document()."""
    for section in doc.sections:
        section._sectPr.append(_el("w:bidi"))
    style = doc.styles["Normal"]
    style.font.name = font
    style.font.size = Pt(size)
    rPr = style.element.get_or_add_rPr()
    rPr.append(_el("w:rtl", **{"w:val": "1"}))
    rPr.get_or_add_rFonts().set(qn("w:cs"), font)
    rPr.append(_el("w:szCs", **{"w:val": str(size * 2)}))
    rPr.append(_el("w:lang", **{"w:bidi": "he-IL"}))
    pPr = style.element.get_or_add_pPr()
    pPr.append(_el("w:bidi", **{"w:val": "1"}))

def rtl_paragraph(p, align=WD_ALIGN_PARAGRAPH.JUSTIFY):
    p._p.get_or_add_pPr().append(_el("w:bidi", **{"w:val": "1"}))
    p.alignment = align
    return p

def rtl_run(run, font=FONT, size=SIZE, bold=False):
    run.font.name, run.font.size, run.bold = font, Pt(size), bold
    rPr = run._r.get_or_add_rPr()
    rPr.append(_el("w:rtl", **{"w:val": "1"}))
    rPr.get_or_add_rFonts().set(qn("w:cs"), font)
    rPr.append(_el("w:szCs", **{"w:val": str(size * 2)}))
    return run

def rtl_table(table):
    """בלי זה עמודות הטבלה יופיעו בסדר הפוך."""
    table._tbl.tblPr.append(_el("w:bidiVisual"))
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                rtl_paragraph(p, WD_ALIGN_PARAGRAPH.RIGHT)
    return table
```

שימוש:

```python
doc = Document()
rtl_document(doc)                      # תמיד ראשון
p = rtl_paragraph(doc.add_paragraph())
rtl_run(p.add_run("לכבוד"), bold=True)
```

**מלכודות ב-Word:**

- כותרות (`add_heading`) משתמשות בסגנון אחר מ-Normal — הרץ `rtl_paragraph` על כל כותרת בנפרד.
- רשימות ממוספרות: הפעל `rtl_paragraph` על כל פריט; אחרת המספר יופיע משמאל.
- כותרת עליונה/תחתונה (header/footer) הן פסקאות נפרדות — גם עליהן.
- עריכת מסמך קיים של המשרד: אל תבנה מחדש. פתח את הקובץ, ערוך במקום, ושמר —
  ההגדרות שלו כבר נכונות.

---

## HTML / PDF

```html
<div dir="rtl" lang="he">…</div>
```

וב-CSS:

```css
body { direction: rtl; text-align: right; font-family: "David", "Frank Ruehl CLM", "Times New Roman", serif; }
table { direction: rtl; }
ol, ul { padding-right: 1.5em; padding-left: 0; }
```

**באר­טיפקט (Artifact):** אין גישה לתג `<html>` — הוא נוצר אוטומטית. לכן חייבים
להגדיר את הכיוון ב-CSS על `body`, ולעטוף את התוכן ב-`<div dir="rtl" lang="he">`.
הסתמכות על `dir` ברמת ה-`<html>` פשוט לא תעבוד שם.

---

## Markdown

Markdown גולמי אינו תומך ב-RTL. אין להוציא מסמך משפטי כ-`.md` נקי.
אם נדרש Markdown (למשל להדבקה למערכת אחרת), עטוף כל בלוק ב-`<div dir="rtl">`
והבהר למשתמש שהתצוגה תלויה במערכת היעד.

---

## מייל (Gmail)

גוף המייל נשלח כ-HTML עם `<div dir="rtl" style="text-align:right">`.
מייל שנשלח כטקסט רגיל יוצג לפי הגדרות הנמען — אין לסמוך על זה במכתב רשמי.
מכתב רשמי נשלח כקובץ `.docx` או `.pdf` מצורף, עם גוף מייל קצר בלבד.

---

## מלכודות דו-כיווניות (bidi) בטקסט עצמו

אלה הבאגים שמופיעים גם כשההגדרות הטכניות נכונות:

| בעיה | דוגמה שגויה | פתרון |
|---|---|---|
| נקודה אחרי מספר בסוף משפט | `הסכום הוא 15,000.` הנקודה קופצת | נסח: `הסכום הוא 15,000 ש"ח.` — מילה עברית לפני הנקודה |
| סעיף עם אות בסוגריים | `סעיף 5(א)` נראה הפוך | כתוב `סעיף 5(א) לחוק` — מילה עברית אחרי |
| טקסט לועזי בתוך משפט | שם חברה באנגלית "בולע" את הפסיק | ב-HTML: `<span dir="ltr">Acme Ltd.</span>`; ב-DOCX: run נפרד בלי `w:rtl` |
| טווח תאריכים או מספרים | `2024-2026` מתהפך | הוסף תו RLM (U+200F) אחרי, או נסח `בין 2024 ל-2026` |
| כתובת אתר או אימייל | נשבר באמצע | שים בשורה נפרדת, או `<span dir="ltr">` |

כלל אצבע לניסוח: **אל תסיים משפט או פסקה במספר, בסוגר או בתו לועזי.**
הוסף מילה עברית אחריהם. זה פותר את רוב הבעיות בלי שום טריק טכני.
