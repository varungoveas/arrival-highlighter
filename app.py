import streamlit as st
import pdfplumber
import re
import io
from pypdf import PdfReader, PdfWriter
from pypdf.generic import (ArrayObject, FloatObject, NameObject,
                            DictionaryObject, TextStringObject)
from collections import defaultdict
from datetime import date

st.set_page_config(page_title="Arrival Report Highlighter", page_icon="🏝️", layout="wide")
st.markdown("""
<style>
#MainMenu, footer, header {visibility: hidden;}
.main .block-container {padding-top: 1.5rem; padding-bottom: 1rem;}
h1 {font-size: 1.6rem !important; margin-bottom: 0 !important;}
</style>
""", unsafe_allow_html=True)

YELLOW = (1.0, 0.921, 0.231)
ORANGE = (1.0, 0.600, 0.100)

CATEGORIES = {
    'occasion':    {'label': 'Special Occasion',       'icon': '🎂',
                    'kw': ['birthday','anniversary','honeymoon','babymoon','celebrating','celebration','proposal','engagement','wedding']},
    'allergy':     {'label': 'Allergy / Dietary',      'icon': '🌿',
                    'kw': ['allerg','intoleranc','gluten','lactose','dietary restriction','vegan','vegetarian','halal','kosher','peanut','shellfish']},
    'payment':     {'label': 'Outstanding Payment',    'icon': '💳',
                    'kw': ['collect upon arrival','please collect','pls collect','balance due','outstanding balance','to be collected','pay upon arrival']},
    'flight':      {'label': 'Flights (🟡 ETA present / 🟠 ETA missing)', 'icon': '✈️', 'kw': []},
    'repeater':    {'label': 'Repeater Guest',         'icon': '⭐', 'kw': []},
    'stayhistory': {'label': 'Stay History',           'icon': '🏨', 'kw': []},
    'complaint':   {'label': 'Complaint / Glitch',     'icon': '⚠️',
                    'kw': ['glitch','recovery','complaint','inconvenien','apologi','disatisf','dissatisf','ttnglitch','ttncomp','feedback','upset']},
    'membership':  {'label': 'GHA Membership',         'icon': '👑', 'kw': []},
    'dbalance':    {'label': 'D$ Balance (> 0)',        'icon': '💰', 'kw': []},
    'legs':        {'label': 'Multi-Villa Legs',       'icon': '🏝️',
                    'kw': ['1st leg','2nd leg','3rd leg','4th leg','5th leg','1ST Leg','2nd Leg','leg:','leg -','leg-','leg of the stay']},
    'vip':         {'label': 'VIP Guests',             'icon': '💎', 'kw': []},
    'sharewith':   {'label': 'Travelling Together (Share with tag)', 'icon': '👥', 'kw': []},
    'welcomenote': {'label': 'Welcome Note',           'icon': '📝',
                    'kw': ['welcome note','welcome amenities','welcome fruit','welcome cake','welcome letter','welcome card','welcome drink','welcome set']},
    'earlyarr':    {'label': 'Early Arrival',          'icon': '🌅', 'kw': []},
    'roomno':      {'label': 'Room / Name / Conf No.', 'icon': '🚪', 'kw': []},
}

def make_annot(x0, top, x1, bottom, ph, color, popup=None):
    px0, py0, px1, py1 = x0-1, ph-bottom-1, x1+1, ph-top+1
    a = DictionaryObject()
    a[NameObject('/Type')]    = NameObject('/Annot')
    a[NameObject('/Subtype')] = NameObject('/Highlight')
    a[NameObject('/Rect')]    = ArrayObject([FloatObject(px0), FloatObject(py0),
                                             FloatObject(px1), FloatObject(py1)])
    a[NameObject('/C')]       = ArrayObject([FloatObject(c) for c in color])
    a[NameObject('/QuadPoints')] = ArrayObject([
        FloatObject(px0), FloatObject(py1), FloatObject(px1), FloatObject(py1),
        FloatObject(px0), FloatObject(py0), FloatObject(px1), FloatObject(py0)])
    a[NameObject('/F')] = FloatObject(4)
    if popup:
        a[NameObject('/Contents')] = TextStringObject(popup)
    return a


def highlight_pdf(pdf_bytes, enabled_cats):
    reader = PdfReader(io.BytesIO(pdf_bytes))
    writer = PdfWriter()
    writer.append(reader)
    counts = {k: 0 for k in CATEGORIES}
    total  = 0

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:

        # ══ PASS 1: identify all bookings & sharers per page ═════════════
        # booking = {room, conf, pg_idx, row_top, row_bottom, is_sharer}
        all_bookings = []

        for pg_idx, page in enumerate(pdf.pages):
            ph = float(page.height)
            words = [w for w in page.extract_words(x_tolerance=2, y_tolerance=2)
                     if w['top'] < ph * 0.91]
            processed = set()
            for rw in [w for w in words if w['x0'] < 40
                       and re.match(r'^[0-9]{1,4}$', w['text'])]:
                t  = rw['top']
                rk = round(t)
                if rk in processed: continue
                processed.add(rk)
                row = sorted([w for w in words if abs(w['top'] - t) <= 4],
                             key=lambda x: x['x0'])
                adl  = next((w for w in row if 415 <= w['x0'] <= 435
                             and w['text'].isdigit()), None)
                name = next((w for w in row if 41 <= w['x0'] <= 280), None)
                conf = next((w for w in words if w['top'] > t+3 and w['top'] < t+25
                             and 35 <= w['x0'] <= 180
                             and re.match(r'^[0-9]{6,}$', w['text'])), None)
                adl_val   = int(adl['text']) if adl else -1
                has_star  = bool(name and name['text'].startswith('*'))
                is_sharer = has_star and adl_val == 0
                all_bookings.append({
                    'room':      rw['text'],
                    'conf':      conf['text'] if conf else '',
                    'pg_idx':    pg_idx,
                    'row_top':   t,
                    'is_sharer': is_sharer,
                })

        # ── Build sharer y-ranges per page ────────────────────────────────
        # For each page, collect (top_start, top_end) ranges belonging to sharers
        # A sharer's range = from its row_top to just before the next booking row
        sharer_ranges = defaultdict(list)  # pg_idx → list of (start, end)

        for pg_idx in set(b['pg_idx'] for b in all_bookings):
            page_bkgs = sorted([b for b in all_bookings if b['pg_idx'] == pg_idx],
                               key=lambda x: x['row_top'])
            for i, b in enumerate(page_bkgs):
                if not b['is_sharer']:
                    continue
                top_start = b['row_top'] - 2  # include the header row itself
                top_end   = (page_bkgs[i+1]['row_top'] - 2
                             if i+1 < len(page_bkgs) else 999)
                sharer_ranges[pg_idx].append((top_start, top_end))

        def is_sharer_line(pg_idx, line_top):
            """Return True if this line falls within a sharer booking's range."""
            for (s, e) in sharer_ranges.get(pg_idx, []):
                if s <= line_top < e:
                    return True
            return False

        # ══ PASS 2: build travelling-together map ════════════════════════
        all_rooms = {b['room'] for b in all_bookings}
        travelling_together = defaultdict(set)

        # Link by shared conf number across bookings
        conf_to_rooms = defaultdict(set)
        for b in all_bookings:
            if b['conf']:
                conf_to_rooms[b['conf']].add(b['room'])
        for conf, rooms in conf_to_rooms.items():
            if len(rooms) > 1:
                for r in rooms:
                    travelling_together[r].update(rooms - {r})

        # Link by conf/room numbers mentioned in comments
        for pg_idx, page in enumerate(pdf.pages):
            ph = float(page.height)
            words = [w for w in page.extract_words(x_tolerance=2, y_tolerance=2)
                     if w['top'] < ph * 0.91]
            page_bkgs = sorted([b for b in all_bookings if b['pg_idx'] == pg_idx],
                               key=lambda x: x['row_top'])
            for i, b in enumerate(page_bkgs):
                top_start    = b['row_top'] + 20
                top_end      = (page_bkgs[i+1]['row_top']
                                if i+1 < len(page_bkgs) else ph * 0.91)
                comment_text = ' '.join(w['text'] for w in words
                                        if top_start <= w['top'] <= top_end)
                for other_b in all_bookings:
                    if other_b['room'] == b['room'] and other_b['conf'] == b['conf']:
                        continue
                    if other_b['conf'] and re.search(
                            r'\b' + re.escape(other_b['conf']) + r'\b', comment_text):
                        travelling_together[b['room']].add(other_b['room'])
                        travelling_together[other_b['room']].add(b['room'])
                    if (other_b['room'] in all_rooms
                            and other_b['room'] != b['room']
                            and re.search(r'\b' + re.escape(other_b['room']) + r'\b',
                                          comment_text)):
                        travelling_together[b['room']].add(other_b['room'])
                        travelling_together[other_b['room']].add(b['room'])

        # ══ PASS 3: highlight each page ══════════════════════════════════
        for pg_idx, page in enumerate(pdf.pages):
            ph       = float(page.height)
            footer_y = ph * 0.91
            words    = [w for w in page.extract_words(x_tolerance=2, y_tolerance=2)
                        if w['top'] < footer_y]
            annots   = []

            def hl(x0, top, x1, bottom, cat, color=YELLOW, popup=None):
                nonlocal total
                counts[cat] += 1
                total += 1
                annots.append(make_annot(x0, top, x1, bottom, ph, color, popup))

            # Build lines
            lg = defaultdict(list)
            for w in words: lg[round(w['top'])].append(w)
            lines = []
            for k in sorted(lg):
                ws = sorted(lg[k], key=lambda x: x['x0'])
                lines.append({
                    'text':   ' '.join(w['text'] for w in ws),
                    'x0':     min(w['x0'] for w in ws),
                    'x1':     max(w['x1'] for w in ws),
                    'top':    min(w['top'] for w in ws),
                    'bottom': max(w['bottom'] for w in ws),
                })

            seen = set()
            def hl_line(ln, cat, color=YELLOW, popup=None):
                # Skip if this line belongs to a sharer booking
                if is_sharer_line(pg_idx, ln['top']):
                    return
                k = f"{ln['x0']:.0f}_{ln['top']:.0f}"
                if k in seen: return
                seen.add(k)
                hl(ln['x0'], ln['top'], ln['x1'], ln['bottom'], cat, color, popup)

            # Keywords
            kw_cats = {k: v['kw'] for k, v in CATEGORIES.items() if v.get('kw')}
            for cat_id, kws in kw_cats.items():
                if not enabled_cats.get(cat_id, True): continue
                for ln in lines:
                    if any(kw.lower() in ln['text'].lower() for kw in kws):
                        hl_line(ln, cat_id)

            # Auto patterns
            for ln in lines:
                t = ln['text']
                if enabled_cats.get('flight'):
                    fm = re.search(
                        r'(EK\s*\d{3,4}|EY\s*\d{3,4}|QR\s*\d{3,4}|'
                        r'G9\s*\d{3,4}|KU\s*\d{3,4}|GF\s*\d{3,4})', t, re.I)
                    if fm:
                        has_eta = bool(re.search(
                            r'\b([0-1]?[0-9]|2[0-3]):[0-5][0-9]\b', t))
                        hl_line(ln, 'flight',
                                YELLOW if has_eta else ORANGE,
                                None if has_eta else '⚠️ ETA MISSING — please check manually')
                if enabled_cats.get('repeater') and re.search(
                        r'\d+(st|nd|rd|th)\s+[Tt]ime\s+RPT|\bRPT\b', t):
                    hl_line(ln, 'repeater')
                if enabled_cats.get('stayhistory') and re.search(
                        r'\d+(st|nd|rd|th)\s+[Ss]tay|Upcoming\s+[Ss]tay', t, re.I):
                    hl_line(ln, 'stayhistory')
                if enabled_cats.get('membership') and re.search(
                        r'Membership\s+Level\s+(GOLD|PLATINUM|TITANIUM|RED)', t, re.I):
                    hl_line(ln, 'membership')
                if enabled_cats.get('dbalance') and re.search(r'[1-9]\d*D\$', t, re.I):
                    hl_line(ln, 'dbalance')
                if enabled_cats.get('vip') and re.search(r'VIP[A-Z0-9]', t):
                    hl_line(ln, 'vip')
                if enabled_cats.get('sharewith') and re.search(r'Share\s+with:', t, re.I):
                    if len(re.sub(r'Share\s+with:', '', t, flags=re.I).strip()) > 1:
                        hl_line(ln, 'sharewith')
                if enabled_cats.get('earlyarr') and re.search(r'0[0-8]:\d{2}', t) \
                        and not re.search(r'Arrival\s+Time', t, re.I):
                    hl_line(ln, 'earlyarr')

            # Room / Name / Conf No — skip sharers entirely
            if enabled_cats.get('roomno'):
                processed = set()
                for rw in [w for w in words if w['x0'] < 40
                           and re.match(r'^[0-9]{1,4}$', w['text'])]:
                    t  = rw['top']
                    rk = round(t)
                    if rk in processed: continue
                    # Skip sharer rows
                    if is_sharer_line(pg_idx, t):
                        processed.add(rk)
                        continue
                    row = sorted([w for w in words if abs(w['top'] - t) <= 4],
                                 key=lambda x: x['x0'])
                    adl = next((w for w in row if 415 <= w['x0'] <= 435
                                and w['text'].isdigit()), None)
                    if not adl or int(adl['text']) == 0:
                        continue
                    name = [w for w in row if 41 <= w['x0'] <= 280]
                    if not name: continue
                    processed.add(rk)

                    together = travelling_together.get(rw['text'], set())
                    popup = ('👥 Travelling Together with Room(s): ' +
                             ', '.join(sorted(together))) if together else None

                    hl(rw['x0'], rw['top'], rw['x1'], rw['bottom'], 'roomno', YELLOW, popup)
                    hl(min(w['x0'] for w in name), min(w['top'] for w in name),
                       max(w['x1'] for w in name), max(w['bottom'] for w in name),
                       'roomno', YELLOW, popup)
                    conf = next((w for w in words if w['top'] > t+3 and w['top'] < t+25
                                 and 35 <= w['x0'] <= 180
                                 and re.match(r'^[0-9]{6,}$', w['text'])), None)
                    if conf:
                        hl(conf['x0'], conf['top'], conf['x1'], conf['bottom'],
                           'roomno', YELLOW, popup)

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
        "Upload one or more arrival PDFs", type=['pdf'],
        accept_multiple_files=True, help="Upload AVEL, ADHI, MNLD PDFs")
    st.markdown("### ⚙️ Categories")
    enabled = {}
    for cat_id, cat in CATEGORIES.items():
        enabled[cat_id] = st.checkbox(
            f"{cat['icon']} {cat['label']}", value=True, key=f"cat_{cat_id}")
    process_btn = st.button("✨ Apply Highlights", type="primary",
                            use_container_width=True, disabled=not uploaded_files)

with col_right:
    if not uploaded_files:
        st.markdown("""
        <div style="background:#f8fafc;border:2px dashed #cbd5e1;border-radius:12px;
                    padding:48px;text-align:center;margin-top:20px">
            <div style="font-size:48px;margin-bottom:12px">📄</div>
            <div style="font-size:18px;font-weight:600;color:#374151;margin-bottom:8px">
                No reports uploaded yet</div>
            <div style="font-size:14px;color:#6b7280;line-height:1.6">
                Upload your daily arrival PDFs on the left<br>
                then click <strong>Apply Highlights</strong><br><br>
                🟡 Yellow = normal highlight<br>
                🟠 Orange = flight with missing ETA<br>
                💬 Hover on room highlight = travelling together rooms<br>
                ✖️ Sharer bookings = no highlights at all
            </div>
        </div>""", unsafe_allow_html=True)

    elif process_btn or st.session_state.get('processed'):
        if process_btn:
            st.session_state['results'] = []
            today = date.today().strftime("%d%b").upper()
            progress_bar = st.progress(0)
            status = st.empty()

            for i, uploaded_file in enumerate(uploaded_files):
                fname = uploaded_file.name
                status.markdown(f"⏳ Processing **{fname}**...")
                progress_bar.progress(i / len(uploaded_files))
                try:
                    pdf_bytes = uploaded_file.read()
                    result_bytes, counts, total = highlight_pdf(pdf_bytes, enabled)
                    base     = fname.replace('.pdf','').replace('.PDF','')
                    out_name = f"{base}_{today}_highlighted.pdf"
                    st.session_state['results'].append({
                        'name': out_name, 'original': fname,
                        'bytes': result_bytes, 'counts': counts,
                        'total': total, 'success': True
                    })
                except Exception as e:
                    st.session_state['results'].append({
                        'name': fname, 'original': fname,
                        'error': str(e), 'success': False
                    })

            progress_bar.progress(1.0)
            status.markdown(f"✅ Done! {len(uploaded_files)} file(s) processed.")
            st.session_state['processed'] = True

        if st.session_state.get('results'):
            st.markdown("### 📥 Download Highlighted Reports")
            for r in st.session_state['results']:
                if r['success']:
                    cat_summary = " · ".join(
                        f"{CATEGORIES[k]['icon']} {v}"
                        for k, v in r['counts'].items() if v > 0)
                    st.markdown(f"""
                    <div style="background:#f0fdf4;border:1px solid #bbf7d0;
                                border-radius:8px;padding:12px 16px;margin:6px 0">
                        <strong>✅ {r['original']}</strong><br>
                        <span style="font-size:12px;color:#166534">
                            {r['total']} highlights — {cat_summary}
                        </span>
                    </div>""", unsafe_allow_html=True)
                    st.download_button(
                        label=f"⬇️ Download {r['name']}",
                        data=r['bytes'], file_name=r['name'],
                        mime='application/pdf', use_container_width=True,
                        key=f"dl_{r['name']}")
                else:
                    st.error(f"❌ {r['original']}: {r.get('error','Unknown error')}")
