"""Build sarvam_market_sizing.xlsx: bottom-up sizing for two Sarvam segments + a scored account map.

Every number is either a FACT (public source, cited on the Sources tab, black text on the Inputs tab)
or an ASSUMPTION (blue text, yellow fill, with the reasoning next to it). All results are live formulas.
Run: python build.py, then recalculate (LibreOffice or Excel).
"""
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

F = "Arial"
BOLD = Font(name=F, bold=True)
HEAD = Font(name=F, bold=True, color="FFFFFF")
HFILL = PatternFill("solid", fgColor="1F4E79")
BLUE = Font(name=F, color="0000FF")
BLACK = Font(name=F, color="000000")
GREEN = Font(name=F, color="008000")
YELLOW = PatternFill("solid", fgColor="FFFF00")
FACTFILL = PatternFill("solid", fgColor="E2EFDA")
THIN = Border(bottom=Side(style="thin", color="BFBFBF"))
WRAP = Alignment(wrap_text=True, vertical="top")

wb = Workbook()


def header(ws, row, cols, widths=None):
    for i, c in enumerate(cols, 1):
        cell = ws.cell(row=row, column=i, value=c)
        cell.font, cell.fill, cell.alignment = HEAD, HFILL, Alignment(wrap_text=True, vertical="center")
    if widths:
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w


def style_all(ws):
    for row in ws.iter_rows():
        for c in row:
            if c.font is None or c.font.name != F:
                c.font = Font(name=F, bold=c.font.bold if c.font else False, color=c.font.color if c.font else None)


# ------------------------------------------------------------------ Sources
src = wb.active
src.title = "Sources"
SOURCES = [
    ("S1", "Sarvam voice agents price per minute (PAYG 3.50, Pro 3.25, Growth 3.15, Business 3.00); Enterprise custom",
     "https://docs.sarvam.ai/voice-agents/pricing", "15 Nov 2026"),
    ("S2", "Sarvam dubbing price per minute of source media per target language (PAYG/Pro 40, Growth 38; Editor Flow 80)",
     "https://docs.sarvam.ai/api/pricing", "15 Nov 2026"),
    ("S3", "Indian banks: 12 public sector, 21 private sector, 12 small finance, 44 foreign, 28 regional rural",
     "https://en.wikipedia.org/wiki/List_of_banks_in_India (compiled from RBI lists)", "15 Nov 2026"),
    ("S4", "NBFCs: 15 in the upper layer, 421 in the middle layer (end-March 2025)",
     "RBI, Report on Trend and Progress of Banking in India 2024-25, ch. VI: https://rbi.org.in/Scripts/PublicationsView.aspx?id=23641",
     "Dec 2025"),
    ("S5", "Insurers: 26 life, 33 non-life (incl. 6 standalone health) recognised by IRDAI",
     "https://en.wikipedia.org/wiki/List_of_insurance_companies_in_India (compiled from IRDAI)", "15 Nov 2026"),
    ("S6", "908 private satellite TV channels permitted", "Ministry of Information & Broadcasting via All India Radio News: "
     "https://www.newsonair.gov.in/indias-media-landscape-expands-over-1-55-lakh-publications-908-private-tv-channels-ib-ministry",
     "2025"),
    ("S7", "Over 1,900 films released in India in 2025; M&E sector Rs 2.78 trillion; digital crossed Rs 1 trillion; "
     "216 million paid video subscriptions", "FICCI-EY M&E report, Mar 2026: https://www.ey.com/en_in/newsroom/2026/03/"
     "india-s-media-and-entertainment-sector-grew-9-percent-to-inr-2-point-78-trillion-in-2025-driven-by-digital-and-live-experiences-ficci-ey-report",
     "Mar 2026"),
]
header(src, 1, ["ID", "Fact", "Source", "Checked"], [6, 80, 90, 14])
for r, row in enumerate(SOURCES, 2):
    for c, v in enumerate(row, 1):
        src.cell(row=r, column=c, value=v).alignment = WRAP

# ------------------------------------------------------------------ Inputs
inp = wb.create_sheet("Inputs")
header(inp, 1, ["Input", "Value", "Unit", "Type", "Source / reasoning"], [52, 14, 22, 13, 90])
INPUTS = [  # name, value, unit, type, note
    ("Voice: Business-plan price", 3.00, "Rs per minute", "FACT", "S1"),
    ("Voice: enterprise discount vs Business price", 0.20, "%", "ASSUMPTION",
     "Large banks buy committed volume; assume 20% below the best public rate"),
    ("Voice: seat working days per year", 300, "days", "ASSUMPTION", "6-day weeks, holidays out"),
    ("Voice: talk minutes per seat per day", 300, "minutes", "ASSUMPTION", "5 hours of talk time in an 8-9 hour shift"),
    ("Voice: share of minutes a voice agent can handle (base)", 0.25, "%", "ASSUMPTION",
     "Reminders, collections nudges, KYC follow-ups, FAQs; complex or regulated calls stay human"),
    ("Voice: share of minutes (low case)", 0.15, "%", "ASSUMPTION", "Sensitivity"),
    ("Voice: share of minutes (high case)", 0.35, "%", "ASSUMPTION", "Sensitivity"),
    ("Voice: 3-year share Sarvam could win", 0.05, "%", "ASSUMPTION",
     "Crowded market (in-house builds, other voice-AI vendors); Indian-language quality and on-prem options help"),
    ("Dubbing: Growth-plan price", 38, "Rs per minute per language", "FACT", "S2"),
    ("Dubbing: enterprise discount", 0.20, "%", "ASSUMPTION", "Volume pricing for broadcasters and OTT libraries"),
    ("Dubbing: 3-year share Sarvam could win", 0.15, "%", "ASSUMPTION",
     "Newer category, fewer Indian-language competitors; still mostly human dubbing studios today"),
]
NAME = {}
for r, (n, v, u, t, note) in enumerate(INPUTS, 2):
    inp.cell(row=r, column=1, value=n)
    c = inp.cell(row=r, column=2, value=v)
    c.font = BLACK if t == "FACT" else BLUE
    c.fill = FACTFILL if t == "FACT" else YELLOW
    c.number_format = "0%" if u == "%" else "#,##0.00" if isinstance(v, float) and v < 10 else "#,##0"
    inp.cell(row=r, column=3, value=u)
    inp.cell(row=r, column=4, value=t)
    inp.cell(row=r, column=5, value=note).alignment = WRAP
    NAME[n] = f"Inputs!$B${r}"
r0 = len(INPUTS) + 3
inp.cell(row=r0, column=1, value="Legend").font = BOLD
inp.cell(row=r0 + 1, column=1, value="Green fill, black text = FACT from a cited public source (Sources tab)").fill = FACTFILL
c = inp.cell(row=r0 + 2, column=1, value="Yellow fill, blue text = ASSUMPTION: change it and every tab updates")
c.fill, c.font = YELLOW, BLUE

P_VOICE = f"{NAME['Voice: Business-plan price']}*(1-{NAME['Voice: enterprise discount vs Business price']})"
MIN_SEAT = f"{NAME['Voice: seat working days per year']}*{NAME['Voice: talk minutes per seat per day']}"

# ------------------------------------------------------------------ Voice agents in BFSI
v = wb.create_sheet("Voice_BFSI")
v["A1"] = "Samvaad voice agents in Indian BFSI: bottom-up annual market (Rs crore)"
v["A1"].font = Font(name=F, bold=True, size=13)
v["A2"] = ("accounts (FACT) x contact-centre seats per account (ASSUMPTION) x talk minutes per seat (Inputs) "
           "x share a voice agent can handle x price per minute")
cols = ["Account type", "Accounts", "Source", "Seats per account (avg)", "Seat reasoning", "Total seats",
        "Talk minutes per year (crore)", "Voice-agent minutes, base (crore)", "Market, base (Rs crore)",
        "Market, low (Rs crore)", "Market, high (Rs crore)", "Share of base market"]
header(v, 4, cols, [30, 10, 8, 12, 46, 12, 14, 14, 14, 13, 13, 11])
VROWS = [
    ("Public sector banks", 12, "S3", 1500, "Very large retail books; many inbound + collections seats"),
    ("Private sector banks", 21, "S3", 600, "A few very large (HDFC, ICICI, Axis), many mid-size; average"),
    ("Small finance banks", 12, "S3", 150, "Retail and microfinance borrowers, collections-heavy"),
    ("Regional rural banks", 28, "S3", 30, "Rural customers, mostly branch-led"),
    ("Foreign banks", 44, "S3", 40, "Small retail presence in India"),
    ("NBFC upper layer", 15, "S4", 800, "Largest lenders; heavy collections and renewals calling"),
    ("NBFC middle layer", 421, "S4", 60, "Mid-size lenders; mostly outsourced calling"),
    ("Life insurers", 26, "S5", 400, "Renewal-premium reminders, lapse revival"),
    ("Non-life insurers (incl. health)", 33, "S5", 250, "Renewals, claims status, health policy servicing"),
]
first = 5
for i, (name, n, s, seats, why) in enumerate(VROWS):
    r = first + i
    v.cell(row=r, column=1, value=name)
    c = v.cell(row=r, column=2, value=n); c.font, c.fill = BLACK, FACTFILL
    v.cell(row=r, column=3, value=s)
    c = v.cell(row=r, column=4, value=seats); c.font, c.fill = BLUE, YELLOW
    v.cell(row=r, column=5, value=why).alignment = WRAP
    v.cell(row=r, column=6, value=f"=B{r}*D{r}").number_format = "#,##0"
    v.cell(row=r, column=7, value=f"=F{r}*{MIN_SEAT}/1E7").number_format = "#,##0.0"
    v.cell(row=r, column=8, value=f"=G{r}*{NAME['Voice: share of minutes a voice agent can handle (base)']}").number_format = "#,##0.0"
    v.cell(row=r, column=9, value=f"=H{r}*{P_VOICE}").number_format = "#,##0"
    v.cell(row=r, column=10, value=f"=G{r}*{NAME['Voice: share of minutes (low case)']}*{P_VOICE}").number_format = "#,##0"
    v.cell(row=r, column=11, value=f"=G{r}*{NAME['Voice: share of minutes (high case)']}*{P_VOICE}").number_format = "#,##0"
last = first + len(VROWS) - 1
for i in range(len(VROWS)):
    r = first + i
    v.cell(row=r, column=12, value=f"=IFERROR(I{r}/$I${last + 1},0)").number_format = "0.0%"
tr = last + 1
v.cell(row=tr, column=1, value="Total").font = BOLD
for col in [2, 6, 7, 8, 9, 10, 11]:
    L = get_column_letter(col)
    c = v.cell(row=tr, column=col, value=f"=SUM({L}{first}:{L}{last})")
    c.font, c.number_format = BOLD, "#,##0" if col not in (7, 8) else "#,##0.0"
v.cell(row=tr + 2, column=1, value="Price used (Rs per minute)")
v.cell(row=tr + 2, column=2, value=f"={P_VOICE}").number_format = "0.00"
v.cell(row=tr + 3, column=1, value="Sarvam 3-year obtainable (Rs crore, base)").font = BOLD
c = v.cell(row=tr + 3, column=2, value=f"=I{tr}*{NAME['Voice: 3-year share Sarvam could win']}")
c.font, c.number_format = BOLD, "#,##0"
v.cell(row=tr + 4, column=1, value="Base market in US$ million (Rs 88 per $)")
v.cell(row=tr + 4, column=2, value=f"=I{tr}*1E7/88/1E6").number_format = "#,##0"
VOICE_TOTAL, VOICE_SOM = f"Voice_BFSI!$I${tr}", f"Voice_BFSI!$B${tr + 3}"
VOICE_LOW, VOICE_HIGH = f"Voice_BFSI!$J${tr}", f"Voice_BFSI!$K${tr}"

# ------------------------------------------------------------------ Content Studio dubbing
d = wb.create_sheet("Dubbing_Content_Studio")
d["A1"] = "Content Studio dubbing in India: bottom-up annual market (Rs crore)"
d["A1"].font = Font(name=F, bold=True, size=13)
d["A2"] = "buyers (FACT where available) x share that dub x source minutes dubbed per buyer per year x target languages x price"
cols = ["Buyer type", "Buyers", "Type", "Source / reasoning", "Share that dub", "Source minutes per buyer per year",
        "Target languages", "Dubbed minutes (lakh)", "Market (Rs crore)", "Share of market"]
header(d, 4, cols, [30, 10, 12, 56, 10, 14, 10, 13, 13, 11])
DROWS = [  # name, buyers, type, note, share, minutes, langs
    ("Private satellite TV channels", 908, "FACT", "S6. Share = news and general entertainment with original shows", 0.40, 3000, 2),
    ("Films released (per year)", 1900, "FACT", "S7. ~140 min a film; pan-India releases dub into 3+ languages", 0.30, 140, 3),
    ("OTT platforms (originals and library)", 40, "ASSUMPTION", "Count assumed; 216M paid video subscriptions (S7) shows demand", 0.60, 20000, 3),
    ("Education (universities, edtech, govt e-learning)", 500, "ASSUMPTION", "Course videos into regional languages", 0.30, 6000, 4),
    ("YouTube and social creators (large)", 5000, "ASSUMPTION", "Creators expanding into new language audiences", 0.20, 1200, 2),
    ("Enterprise training and marketing", 1000, "ASSUMPTION", "Product videos, training for frontline staff", 0.25, 600, 3),
]
first = 5
for i, (name, n, t, note, share, mins, langs) in enumerate(DROWS):
    r = first + i
    d.cell(row=r, column=1, value=name)
    c = d.cell(row=r, column=2, value=n)
    c.font, c.fill = (BLACK, FACTFILL) if t == "FACT" else (BLUE, YELLOW)
    d.cell(row=r, column=3, value=t)
    d.cell(row=r, column=4, value=note).alignment = WRAP
    for col, val, fmt in [(5, share, "0%"), (6, mins, "#,##0"), (7, langs, "0")]:
        c = d.cell(row=r, column=col, value=val); c.font, c.fill, c.number_format = BLUE, YELLOW, fmt
    d.cell(row=r, column=8, value=f"=B{r}*E{r}*F{r}*G{r}/1E5").number_format = "#,##0.0"
    d.cell(row=r, column=9, value=f"=H{r}*1E5*{NAME['Dubbing: Growth-plan price']}*(1-{NAME['Dubbing: enterprise discount']})/1E7").number_format = "#,##0.0"
last = first + len(DROWS) - 1
for i in range(len(DROWS)):
    r = first + i
    d.cell(row=r, column=10, value=f"=IFERROR(I{r}/$I${last + 1},0)").number_format = "0.0%"
tr = last + 1
d.cell(row=tr, column=1, value="Total").font = BOLD
for col in [8, 9]:
    L = get_column_letter(col)
    c = d.cell(row=tr, column=col, value=f"=SUM({L}{first}:{L}{last})"); c.font, c.number_format = BOLD, "#,##0.0"
d.cell(row=tr + 2, column=1, value="Sarvam 3-year obtainable (Rs crore)").font = BOLD
c = d.cell(row=tr + 2, column=2, value=f"=I{tr}*{NAME['Dubbing: 3-year share Sarvam could win']}")
c.font, c.number_format = BOLD, "#,##0.0"
DUB_TOTAL, DUB_SOM = f"Dubbing_Content_Studio!$I${tr}", f"Dubbing_Content_Studio!$B${tr + 2}"

# ------------------------------------------------------------------ Account map
a = wb.create_sheet("Account_Map")
a["A1"] = "Target account map: where sales should spend time first (scores are analyst judgment, not data)"
a["A1"].font = Font(name=F, bold=True, size=13)
a["A2"] = "Weights (edit):"
W = [("Pool size", 0.35), ("Regional-language need", 0.25), ("Data-sovereignty pull", 0.20), ("Reachability", 0.20)]
for i, (n, w) in enumerate(W):
    a.cell(row=2, column=2 + 2 * i, value=n)
    c = a.cell(row=2, column=3 + 2 * i, value=w); c.font, c.fill, c.number_format = BLUE, YELLOW, "0%"
a["J2"] = "=SUM(C2,E2,G2,I2)"
a["J2"].number_format = "0%"
a["K2"] = "must be 100%"
cols = ["Account", "Type", "Product", "Pool size (1-5)", "Regional-language need (1-5)", "Data-sovereignty pull (1-5)",
        "Reachability (1-5)", "Priority score", "Rank", "Why (one line)"]
header(a, 4, cols, [30, 22, 16, 10, 12, 12, 11, 10, 7, 70])
ACCTS = [
    ("State Bank of India", "Public sector bank", "Voice agents", 5, 5, 5, 3, "Largest customer base, deep rural reach, public-sector data rules"),
    ("Bank of Baroda", "Public sector bank", "Voice agents", 4, 4, 5, 3, "Large retail book across many states"),
    ("Punjab National Bank", "Public sector bank", "Voice agents", 4, 4, 5, 3, "Large north-India retail base"),
    ("Canara Bank", "Public sector bank", "Voice agents", 4, 5, 5, 3, "Strong south-India base, Kannada/Malayalam/Tamil need"),
    ("Union Bank of India", "Public sector bank", "Voice agents", 4, 4, 5, 3, "Large retail book"),
    ("HDFC Bank", "Private sector bank", "Voice agents", 5, 3, 3, 4, "Largest private bank; mature contact-centre tech"),
    ("ICICI Bank", "Private sector bank", "Voice agents", 5, 3, 3, 4, "Large retail; early digital adopter"),
    ("Axis Bank", "Private sector bank", "Voice agents", 4, 3, 3, 4, "Large retail and cards"),
    ("Kotak Mahindra Bank", "Private sector bank", "Voice agents", 3, 3, 3, 4, "Digital-first retail"),
    ("AU Small Finance Bank", "Small finance bank", "Voice agents", 3, 4, 3, 4, "Semi-urban borrowers, collections-heavy"),
    ("Bajaj Finance", "NBFC upper layer", "Voice agents", 5, 4, 3, 4, "Very large consumer-loan book, high call volumes"),
    ("Shriram Finance", "NBFC upper layer", "Voice agents", 4, 5, 3, 4, "Vehicle loans to rural and semi-urban borrowers"),
    ("Muthoot Finance", "NBFC", "Voice agents", 3, 5, 3, 4, "Gold loans; regional-language customers"),
    ("LIC", "Life insurer (public)", "Voice agents", 5, 5, 5, 2, "Huge policy base needing renewal reminders in many languages"),
    ("SBI Life", "Life insurer", "Voice agents", 4, 4, 4, 4, "Bancassurance through SBI's rural network"),
    ("HDFC Life", "Life insurer", "Voice agents", 4, 3, 3, 4, "Renewal and lapse-revival calling"),
    ("Star Health", "Standalone health insurer", "Voice agents", 3, 4, 3, 4, "Retail health; renewals and claim status"),
    ("Prasar Bharati (Doordarshan, AIR)", "Public broadcaster", "Content Studio", 4, 5, 5, 2, "Multi-language public broadcasting"),
    ("Large news networks", "TV / digital news", "Content Studio", 3, 5, 2, 4, "Daily news into several languages, speed matters"),
    ("OTT platforms", "Streaming", "Content Studio", 4, 4, 2, 3, "Dub libraries and originals for regional audiences"),
    ("Edtech and govt e-learning", "Education", "Content Studio", 3, 5, 4, 3, "Course videos into regional languages"),
]
first = 5
for i, (name, typ, prod, *scores, why) in enumerate(ACCTS):
    r = first + i
    a.cell(row=r, column=1, value=name)
    a.cell(row=r, column=2, value=typ)
    a.cell(row=r, column=3, value=prod)
    for j, s in enumerate(scores):
        c = a.cell(row=r, column=4 + j, value=s); c.font, c.fill = BLUE, YELLOW
    a.cell(row=r, column=8, value=f"=ROUND(D{r}*$C$2+E{r}*$E$2+F{r}*$G$2+G{r}*$I$2,2)").number_format = "0.00"
    a.cell(row=r, column=10, value=why).alignment = WRAP
last = first + len(ACCTS) - 1
for i in range(len(ACCTS)):
    r = first + i
    a.cell(row=r, column=9, value=f"=RANK(H{r},$H${first}:$H${last},0)")
a.cell(row=last + 2, column=1, value=("Scores are my judgment from public information about each organisation's size, "
                                      "customer base and sector; they are inputs to change, not data. "
                                      "None of these are claimed to be Sarvam customers.")).alignment = WRAP
a.freeze_panes = "B5"

# ------------------------------------------------------------------ Summary (first tab)
s = wb.create_sheet("Summary", 0)
s.column_dimensions["A"].width = 58
s.column_dimensions["B"].width = 18
s.column_dimensions["C"].width = 60
s["A1"] = "Sarvam AI: where the next revenue comes from (bottom-up sizing)"
s["A1"].font = Font(name=F, bold=True, size=14)
s["A2"] = ("Built exclusively for Sarvam AI's PS-II Strategy & Operations internship by Pavitra Seth (BITS Pilani). "
           "Public data only; not affiliated with Sarvam AI.")
s["A2"].alignment = WRAP
rows = [
    ("Voice agents in BFSI: annual market, base (Rs crore)", f"={VOICE_TOTAL}", "Voice_BFSI tab"),
    ("   low case (15% of minutes)", f"={VOICE_LOW}", ""),
    ("   high case (35% of minutes)", f"={VOICE_HIGH}", ""),
    ("   Sarvam 3-year obtainable (Rs crore)", f"={VOICE_SOM}", "Share on Inputs tab"),
    ("Content Studio dubbing: annual market (Rs crore)", f"={DUB_TOTAL}", "Dubbing_Content_Studio tab"),
    ("   Sarvam 3-year obtainable (Rs crore)", f"={DUB_SOM}", ""),
    ("Voice market is this many times the dubbing market", f"=IFERROR({VOICE_TOTAL}/{DUB_TOTAL},0)", ""),
]
header(s, 4, ["Result", "Value", "Where it comes from"])
for i, (lab, f_, where) in enumerate(rows, 5):
    s.cell(row=i, column=1, value=lab)
    c = s.cell(row=i, column=2, value=f_); c.font = GREEN
    c.number_format = "0.0x" if "times" in lab else "#,##0"
    s.cell(row=i, column=3, value=where)
s["A13"] = "Takeaways"
s["A13"].font = BOLD
TAKE = [
    "1. In rupees, voice agents in BFSI are the bigger pool by far, because call minutes are huge and every minute is billed. "
    "Public-sector banks and NBFCs make up most of it.",
    "2. Dubbing is a smaller pool at Rs 30-38 a minute, so Content Studio's case rests on fewer, larger accounts "
    "(OTT libraries, TV channels) and on adjacent products (voice cloning, TTS), not on volume alone.",
    "3. The two inputs that move the voice market most: seats per account and the share of minutes a voice agent can handle. "
    "Both should be checked in 5 customer calls before any target is set.",
    "4. Account map: start with the public-sector banks, the large NBFCs and LIC, where regional-language need "
    "and data-sovereignty pull are both high.",
]
for i, t in enumerate(TAKE, 14):
    s.cell(row=i, column=1, value=t).alignment = WRAP
    s.merge_cells(start_row=i, start_column=1, end_row=i, end_column=3)
    s.row_dimensions[i].height = 32
s["A19"] = "Colour key: blue/yellow = assumption you can change; black/green fill = fact with a source; green text = link to another tab."
s["A19"].alignment = WRAP

for ws in wb.worksheets:
    for row in ws.iter_rows():
        for c in row:
            if c.value is not None and (c.font is None or c.font.name != F):
                c.font = Font(name=F, bold=c.font.bold, color=c.font.color, size=c.font.size)
wb.save("sarvam_market_sizing.xlsx")
print("saved")
