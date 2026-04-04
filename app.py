import streamlit as st
import pdfplumber
import re
import io
from pypdf import PdfReader, PdfWriter
from pypdf.generic import ArrayObject, FloatObject, NameObject, DictionaryObject
from collections import defaultdict
from datetime import date

st.set_page_config(
    page_title="Arrival Report Highlighter",
    page_icon="🏝️",
    layout="wide"
)

# ── STYLING ──────────────────────────────────────────────────────────────
st.markdown("""
<style>
#MainMenu, footer, header {visibility: hidden;}
.main .block-container {padding-top: 1.5rem; padding-bottom: 1rem;}
.stAlert {border-radius: 8px;}
h1 {font-size: 1.6rem !important; margin-bottom: 0 !important;}
.property-badge {
    display: inline-block; padding: 3px 10px; border-radius: 12px;
    font-size: 12px; font-weight: 600; margin: 3px;
    background: #eef2ff; color: #1a56db; border: 1px solid #c7d7f8;
}
.result-box {
    background: #f0fdf4; border: 1px solid #bbf7d0; border-radius: 8px;
    padding: 12px 16px; margin: 6px 0;
}
.skip-box {
    background: #fafafa; border: 1px solid #e5e7eb; border-radius: 8px;
    padding: 10px 16px; margin: 6px 0; opacity: 0.7;
}
</style>
""", unsafe_allow_html=True)

# ── HIGHLIGHT COLOR ───────────────────────────────────────────────────────
HR, HG, HB = 1.0, 0.921, 0.231

# ── CATEGORIES ───────────────────────────────────────────────────────────
CATEGORIES = {
    'occasion':    {'label': 'Special Occasion',       'icon': '🎂',
                    'kw': ['birthday','anniversary','honeymoon','babymoon','celebrating','celebration','proposal','engagement','wedding']},
    'allergy':     {'label': 'Allergy / Dietary',      'icon': '🌿',
                    'kw': ['allerg','intoleranc','gluten','lactose','dietary restriction','vegan','vegetarian','halal','kosher','peanut']},
    'payment':     {'label': 'Outstanding Payment',    'icon': '💳',
                    'kw': ['collect','please collect','pls collect','balance due','outstanding balance','to be collected','collect upon arrival','pay upon arrival']},
    'flight':      {'label': 'Middle East Flights',    'icon': '✈️',  'kw': [], 'auto': 'flight'},
    'repeater':    {'label': 'Repeater Guest',         'icon': '⭐',  'kw': [], 'auto': 'repeater'},
    'stayhistory': {'label': 'Stay History',           'icon': '🏨',  'kw': [], 'auto': 'stayhistory'},
    'complaint':   {'label': 'Complaint / Glitch',     'icon': '⚠️',
                    'kw': ['glitch','recovery','complaint','inconvenien','apologi','disatisf','dissatisf','ttnglitch','ttncomp']},
    'membership':  {'label': 'GHA Membership',         'icon': '👑',  'kw': [], 'auto': 'membership'},
    'dbalance':    {'label': 'D$ Balance (> 0)',        'icon': '💰',  'kw': [], 'auto': 'dbalance'},
    'legs':        {'label': 'Multi-Villa Legs',       'icon': '🏝️',
                    'kw': ['1st leg','2nd leg','3rd leg','4th leg','5th leg','leg:','leg -','leg-','leg of the stay']},
    'vip':         {'label': 'VIP Guests',             'icon': '💎',  'kw': [], 'auto': 'vip'},
    'sharewith':   {'label': 'Travelling Together',    'icon': '👥',  'kw': [], 'auto': 'sharewith'},
    'welcomenote': {'label': 'Welcome Note',           'icon': '📝',
                    'kw': ['welcome note','welcome amenities','welcome fruit','welcome cake','welcome letter','welcome card','welcome drink','welcome set']},
    'earlyarr':    {'label': 'Early Arrival',          'icon': '🌅',  'kw': [], 'auto': 'earlyarr'},
    'roomno':      {'label': 'Room / Name / Conf No.', 'icon': '🚪',  'kw': [], 'auto': 'roomno'},
}

# ── CORE HIGHLIGHTER ─────────────────────────────────────────────────────
def highlight_pdf(pdf_bytes, enabled_cats):
    reader = PdfReader(io.BytesIO(pdf_bytes))
    writer = PdfWriter()
    writer.append(reader)
    counts = {k: 0 for k in CATEGORIES}
    total = 0

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for pg_idx, page in enumerate(pdf.pages):
            ph = float(page.height)
            footer_y = ph * 0.91
            words = [w for w in page.extract_words(x_tolerance=2, y_tolerance=2)
                     if w['top'] < footer_y]
            annots = []

            def hl(x0, top, x1, bottom, cat):
                nonlocal total
                if not enabled_cats.get(cat, True): return
                counts[cat] += 1
                total += 1
                px0, py0, px1, py1 = x0-1, ph-bottom-1, x1+1, ph-top+1
                a = DictionaryObject()
                a[NameObject('/Type')] = NameObject('/Annot')
                a[NameObject('/Subtype')] = NameObject('/Highlight')
                a[NameObject('/Rect')] = ArrayObject([FloatObject(px0), FloatObject(py0),
                                                       FloatObject(px1), FloatObject(py1)])
                a[NameObject('/C')] = ArrayObject([FloatObject(HR), FloatObject(HG), FloatObject(HB)])
                a[NameObject('/QuadPoints')] = ArrayObject([
                    FloatObject(px0), FloatObject(py1), FloatObject(px1), FloatObject(py1),
                    FloatObject(px0), FloatObject(py0), FloatObject(px1), FloatObject(py0)])
                a[NameObject('/F')] = FloatObject(4)
                annots.append(a)

            # Group words into lines
            lg = defaultdict(list)
            for w in words:
                lg[round(w['top'])].append(w)
            lines = []
            for k in sorted(lg):
                ws = sorted(lg[k], key=lambda x: x['x0'])
                lines.append({
                    'text': ' '.join(w['text'] for w in ws),
                    'x0': min(w['x0'] for w in ws), 'x1': max(w['x1'] for w in ws),
                    'top': min(w['top'] for w in ws), 'bottom': max(w['bottom'] for w in ws)
                })

            seen = set()
            def hl_line(ln, cat):
                k = f"{ln['x0']:.0f}_{ln['top']:.0f}"
                if k in seen: return
                seen.add(k)
                hl(ln['x0'], ln['top'], ln['x1'], ln['bottom'], cat)

            # ── Keywords ────────────────────────────────────────────────
            for cat_id, cat in CATEGORIES.items():
                if not enabled_cats.get(cat_id, True): continue
                for kw in cat.get('kw', []):
                    for ln in lines:
                        if kw.lower() in ln['text'].lower():
                            hl_line(ln, cat_id)

            # ── Auto patterns ────────────────────────────────────────────
            for ln in lines:
                t = ln['text']
                if enabled_cats.get('flight') and re.search(r'EK\s*\d{3,4}|EY\s*\d{3,4}|QR\s*\d{3,4}|G9\s*\d{3,4}|KU\s*\d{3,4}|GF\s*\d{3,4}', t, re.I):
                    hl_line(ln, 'flight')
                if enabled_cats.get('repeater') and re.search(r'\d+(st|nd|rd|th)\s+[Tt]ime\s+RPT|\bRPT\b', t):
                    hl_line(ln, 'repeater')
                if enabled_cats.get('stayhistory') and re.search(r'\d+(st|nd|rd|th)\s+Stay|Upcoming\s+Stay', t, re.I):
                    hl_line(ln, 'stayhistory')
                if enabled_cats.get('membership') and re.search(r'Membership\s+Level\s+(GOLD|PLATINUM|TITANIUM|RED)', t, re.I):
                    hl_line(ln, 'membership')
                if enabled_cats.get('dbalance') and re.search(r'[1-9]\d*D\$', t, re.I):
                    hl_line(ln, 'dbalance')
                if enabled_cats.get('vip') and re.search(r'VIP[A-Z]', t):
                    hl_line(ln, 'vip')
                if enabled_cats.get('sharewith') and re.search(r'Share\s+with:', t, re.I):
                    if len(re.sub(r'Share\s+with:', '', t, flags=re.I).strip()) > 1:
                        hl_line(ln, 'sharewith')
                if enabled_cats.get('earlyarr') and re.search(r'0[0-8]:\d{2}', t) and not re.search(r'Arrival\s+Time', t, re.I):
                    hl_line(ln, 'earlyarr')

            # ── Room / Name / Conf No ────────────────────────────────────
            if enabled_cats.get('roomno', True):
                processed = set()
                room_words = [w for w in words if w['x0'] < 40
                              and re.match(r'^[0-9]{1,4}$', w['text'])]
                for rw in room_words:
                    t = rw['top']
                    rk = round(t)
                    if rk in processed: continue
                    row = sorted([w for w in words if abs(w['top'] - t) <= 4],
                                 key=lambda x: x['x0'])
                    adl = next((w for w in row if 415 <= w['x0'] <= 435
                                and w['text'].isdigit()), None)
                    if not adl or int(adl['text']) == 0: continue
                    name = [w for w in row if 41 <= w['x0'] <= 280]
                    if not name: continue
                    processed.add(rk)
                    hl(rw['x0'], rw['top'], rw['x1'], rw['bottom'], 'roomno')
                    hl(min(w['x0'] for w in name), min(w['top'] for w in name),
                       max(w['x1'] for w in name), max(w['bottom'] for w in name), 'roomno')
                    conf = next((w for w in words if w['top'] > t + 3 and w['top'] < t + 25
                                 and 35 <= w['x0'] <= 180
                                 and re.match(r'^[0-9]{6,}$', w['text'])), None)
                    if conf:
                        hl(conf['x0'], conf['top'], conf['x1'], conf['bottom'], 'roomno')

            if annots:
                p = writer.pages[pg_idx]
                if '/Annots' in p:
                    for a in annots: p['/Annots'].append(a)
                else:
                    p[NameObject('/Annots')] = ArrayObject(annots)

    out = io.BytesIO()
    writer.write(out)
    return out.getvalue(), counts, total


# ── UI ───────────────────────────────────────────────────────────────────
st.markdown("## 🏝️ Arrival Report Highlighter")
st.markdown("**Anantara Veli · Anantara Dhigu · Naladhu Private Island**")
st.markdown("---")

col_left, col_right = st.columns([1, 2])

with col_left:
    st.markdown("### 📂 Upload Reports")
    uploaded_files = st.file_uploader(
        "Upload one or more arrival PDFs",
        type=['pdf'],
        accept_multiple_files=True,
        help="Upload AVEL, ADHI, MNLD PDFs — all at once or one by one"
    )

    st.markdown("### ⚙️ Categories")
    enabled = {}
    for cat_id, cat in CATEGORIES.items():
        enabled[cat_id] = st.checkbox(
            f"{cat['icon']} {cat['label']}",
            value=True,
            key=f"cat_{cat_id}"
        )

    process_btn = st.button(
        "✨ Apply Highlights",
        type="primary",
        use_container_width=True,
        disabled=not uploaded_files
    )

with col_right:
    if not uploaded_files:
        st.markdown("""
        <div style="background:#f8fafc;border:2px dashed #cbd5e1;border-radius:12px;
                    padding:48px;text-align:center;margin-top:20px">
            <div style="font-size:48px;margin-bottom:12px">📄</div>
            <div style="font-size:18px;font-weight:600;color:#374151;margin-bottom:8px">
                No reports uploaded yet
            </div>
            <div style="font-size:14px;color:#6b7280;line-height:1.6">
                Upload your daily arrival PDFs on the left<br>
                then click <strong>Apply Highlights</strong>
            </div>
        </div>
        """, unsafe_allow_html=True)

    elif process_btn or st.session_state.get('processed'):
        if process_btn:
            st.session_state['results'] = []
            today = date.today().strftime("%d%b").upper()

            progress_bar = st.progress(0)
            status = st.empty()

            for i, uploaded_file in enumerate(uploaded_files):
                fname = uploaded_file.name
                status.markdown(f"⏳ Processing **{fname}**...")
                progress_bar.progress((i) / len(uploaded_files))

                try:
                    pdf_bytes = uploaded_file.read()
                    result_bytes, counts, total = highlight_pdf(pdf_bytes, enabled)

                    # Smart output filename
                    base = fname.replace('.pdf', '').replace('.PDF', '')
                    out_name = f"{base}_{today}_highlighted.pdf"

                    st.session_state['results'].append({
                        'name': out_name,
                        'original': fname,
                        'bytes': result_bytes,
                        'counts': counts,
                        'total': total,
                        'success': True
                    })
                except Exception as e:
                    st.session_state['results'].append({
                        'name': fname,
                        'original': fname,
                        'error': str(e),
                        'success': False
                    })

            progress_bar.progress(1.0)
            status.markdown(f"✅ Done! {len(uploaded_files)} file(s) processed.")
            st.session_state['processed'] = True

        # Show results
        if st.session_state.get('results'):
            st.markdown("### 📥 Download Highlighted Reports")
            for r in st.session_state['results']:
                if r['success']:
                    cat_summary = " · ".join(
                        f"{CATEGORIES[k]['icon']} {v}"
                        for k, v in r['counts'].items() if v > 0
                    )
                    st.markdown(f"""
                    <div class="result-box">
                        <strong>✅ {r['original']}</strong><br>
                        <span style="font-size:12px;color:#166534">
                            {r['total']} highlights — {cat_summary}
                        </span>
                    </div>
                    """, unsafe_allow_html=True)
                    st.download_button(
                        label=f"⬇️ Download {r['name']}",
                        data=r['bytes'],
                        file_name=r['name'],
                        mime='application/pdf',
                        use_container_width=True,
                        key=f"dl_{r['name']}"
                    )
                else:
                    st.error(f"❌ {r['original']}: {r.get('error', 'Unknown error')}")

