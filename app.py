import streamlit as st
import pdfplumber
import re
import io
from pypdf import PdfReader, PdfWriter
from pypdf.generic import (ArrayObject, FloatObject, NameObject,
                            DictionaryObject, TextStringObject)
from collections import defaultdict
from datetime import date
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER

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


def merge_groups(travelling_together):
    """Merge connected rooms into complete groups using union-find,
    so every room shows ALL rooms in its group, not just direct links."""
    all_rooms = set(travelling_together.keys())
    for rooms in travelling_together.values():
        all_rooms.update(rooms)
    parent = {r: r for r in all_rooms}

    def find(x):
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    def union(a, b):
        parent[find(a)] = find(b)

    for room, linked in travelling_together.items():
        for other in linked:
            union(room, other)

    groups = defaultdict(set)
    for room in all_rooms:
        groups[find(room)].add(room)

    room_to_group = {}
    for group in groups.values():
        if len(group) > 1:
            for room in group:
                room_to_group[room] = sorted(group)  # include self
    return room_to_group


def build_summary_page(summary_data):
    """Build a landscape A4 summary/briefing page as PDF bytes."""
    buf = io.BytesIO()
    page_size = landscape(A4)
    doc = SimpleDocTemplate(buf, pagesize=page_size,
                            leftMargin=15*mm, rightMargin=15*mm,
                            topMargin=12*mm, bottomMargin=12*mm)
    W = page_size[0] - 30*mm
    story = []

    C_HEADER  = colors.HexColor('#1a1a2e')
    C_MUT     = colors.HexColor('#6b7280')
    C_BG      = colors.HexColor('#f8fafc')
    C_BORDER  = colors.HexColor('#e2e8f0')
    C_ROW_ALT = colors.HexColor('#f1f5f9')

    def ps(size, color=C_HEADER, bold=False, align=TA_LEFT):
        return ParagraphStyle('x', fontSize=size, textColor=color,
                              fontName='Helvetica-Bold' if bold else 'Helvetica',
                              alignment=align, leading=size * 1.4)

    # Header — two rows: title line + subtitle line, well spaced
    ht = Table([[
        Paragraph(
            f'<font size="8" color="#6b7280">{summary_data["property"]}</font>',
            ps(8, C_MUT)),
        Paragraph(
            f'<font size="8" color="#6b7280">Generated: {summary_data["generated"]}  |  '
            f'{summary_data["total"]} arrivals  |  {summary_data["rooms"]} rooms</font>',
            ps(8, C_MUT, align=TA_CENTER)),
        Paragraph(
            f'<font size="8" color="#6b7280">Arrival Date: {summary_data["date"]}</font>',
            ps(8, C_MUT, align=TA_CENTER)),
    ]], colWidths=[W*0.4, W*0.35, W*0.25])
    ht.setStyle(TableStyle([
        ('BACKGROUND',(0,0),(-1,-1),C_BG),
        ('BOX',(0,0),(-1,-1),0.5,C_BORDER),
        ('TOPPADDING',(0,0),(-1,-1),6),('BOTTOMPADDING',(0,0),(-1,-1),6),
        ('LEFTPADDING',(0,0),(-1,-1),14),('RIGHTPADDING',(0,0),(-1,-1),14),
        ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
    ]))

    title_row = Table([[
        Paragraph(
            '<font size="18"><b>Daily Operations Meeting</b></font>',
            ps(18, C_HEADER, bold=True))
    ]], colWidths=[W])
    title_row.setStyle(TableStyle([
        ('BACKGROUND',(0,0),(-1,-1),C_BG),
        ('LINEBELOW',(0,0),(-1,-1),0.5,C_BORDER),
        ('LINEBEFORE',(0,0),(0,-1),0.5,C_BORDER),
        ('LINEAFTER',(-1,0),(-1,-1),0.5,C_BORDER),
        ('TOPPADDING',(0,0),(-1,-1),8),('BOTTOMPADDING',(0,0),(-1,-1),8),
        ('LEFTPADDING',(0,0),(-1,-1),14),('RIGHTPADDING',(0,0),(-1,-1),14),
        ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
    ]))

    story.append(title_row)
    story.append(ht)
    story.append(Spacer(1, 5))

    # Stats
    stat_items = [
        ('Rooms',      summary_data['rooms']),
        ('Repeaters',  summary_data['counts'].get('repeater', 0)),
        ('Flights',    summary_data['counts'].get('flight', 0)),
        ('GHA',        summary_data['counts'].get('membership', 0)),
        ('VIP',        summary_data['counts'].get('vip', 0)),
        ('Complaints', summary_data['counts'].get('complaint', 0)),
        ('Allergies',  summary_data['counts'].get('allergy', 0)),
        ('Occasions',  summary_data['counts'].get('occasion', 0)),
        ('Payment',    summary_data['counts'].get('payment', 0)),
    ]
    cw = W / len(stat_items)
    label_row  = [Paragraph(f'<font size="8" color="#6b7280">{s[0]}</font>',
                             ps(8, align=TA_CENTER)) for s in stat_items]
    number_row = [Paragraph(f'<b>{s[1]}</b>',
                             ps(22, C_HEADER, bold=True, align=TA_CENTER)) for s in stat_items]
    st = Table([label_row, number_row], colWidths=[cw]*len(stat_items))
    st.setStyle(TableStyle([
        ('BACKGROUND',(0,0),(-1,-1),C_BG),
        ('BOX',(0,0),(-1,-1),0.5,C_BORDER),
        ('LINEAFTER',(0,0),(-2,-1),0.5,C_BORDER),
        ('ALIGN',(0,0),(-1,-1),'CENTER'),
        ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
        ('TOPPADDING',(0,0),(-1,0),8),('BOTTOMPADDING',(0,0),(-1,0),1),
        ('TOPPADDING',(0,1),(-1,1),1),('BOTTOMPADDING',(0,1),(-1,1),9),
        ('LINEBELOW',(0,0),(-1,0),0,C_BG),
    ]))
    story.append(st)
    story.append(Spacer(1, 5))

    # Guest table
    def flag_para(flags):
        parts = []
        for f in flags:
            if any(x in f for x in ['VIP','Titanium','Platinum','Gold','Red']):
                bg, fg = '#EEEDFE','#26215C'
            elif any(x in f.lower() for x in ['allergy','shellfish','gluten','peanut','lactose','vegan','vegetarian','halal','kosher']):
                bg, fg = '#FCEBEB','#501313'
            elif 'Complaint' in f or 'Glitch' in f:
                bg, fg = '#FCEBEB','#501313'
            elif 'Collect' in f or 'Payment' in f:
                bg, fg = '#FAEEDA','#633806'
            elif 'Together' in f:
                bg, fg = '#E6F1FB','#042C53'
            elif any(x in f for x in ['RPT','stay','Repeat']):
                bg, fg = '#E1F5EE','#04342C'
            elif any(x in f for x in ['Honeymoon','Anniversary','Birthday','Wedding','Occasion']):
                bg, fg = '#EEEDFE','#26215C'
            elif 'Leg' in f or 'leg' in f:
                bg, fg = '#FAEEDA','#633806'
            else:
                bg, fg = '#F1EFE8','#2C2C2A'
            parts.append(f'<font size="8" color="{fg}" backColor="{bg}"> {f} </font>')
        return Paragraph('   '.join(parts), ps(8))

    def flight_para(f):
        if 'NO ETA' in f:
            return Paragraph(f'<font size="8" color="#ffffff" backColor="#E24B4A"> {f} </font>', ps(8))
        elif f and f not in ('--', '-', ''):
            return Paragraph(f'<font size="8" color="#633806" backColor="#FAEEDA"> {f} </font>', ps(8))
        return Paragraph('-', ps(8, C_MUT))

    cws = [12*mm, 44*mm, 27*mm, 26*mm, 70*mm, W-179*mm]
    thead = [
        Paragraph('Room',     ps(8, C_MUT, bold=True)),
        Paragraph('Guest',    ps(8, C_MUT, bold=True)),
        Paragraph('Conf No.', ps(8, C_MUT, bold=True)),
        Paragraph('Flight',   ps(8, C_MUT, bold=True)),
        Paragraph('Flags',    ps(8, C_MUT, bold=True)),
        Paragraph('Notes',    ps(8, C_MUT, bold=True)),
    ]
    rows = [thead]
    for g in summary_data['guests']:
        rows.append([
            Paragraph(f'<b>{g["room"]}</b>', ps(9, C_HEADER, bold=True)),
            Paragraph(g['name'], ps(9)),
            Paragraph(f'<font color="#6b7280">{g["conf"]}</font>', ps(8)),
            flight_para(g['flight']),
            flag_para(g['flags']),
            Paragraph(f'<font color="#6b7280">{g["note"]}</font>', ps(8)),
        ])

    gt = Table(rows, colWidths=cws, repeatRows=1)
    rstyles = [
        ('BACKGROUND',(0,0),(-1,0),C_BG),
        ('BOX',(0,0),(-1,-1),0.5,C_BORDER),
        ('INNERGRID',(0,0),(-1,-1),0.5,C_BORDER),
        ('TOPPADDING',(0,0),(-1,-1),7),('BOTTOMPADDING',(0,0),(-1,-1),7),
        ('LEFTPADDING',(0,0),(-1,-1),7),('RIGHTPADDING',(0,0),(-1,-1),7),
        ('VALIGN',(0,0),(-1,-1),'TOP'),
    ]
    for i in range(1, len(rows)):
        if i % 2 == 0:
            rstyles.append(('BACKGROUND',(0,i),(-1,i),C_ROW_ALT))
    gt.setStyle(TableStyle(rstyles))
    story.append(gt)
    story.append(Spacer(1, 5))

    # Legend
    legend_items = [
        ('#EEEDFE','#26215C','VIP / membership'),
        ('#FCEBEB','#501313','Complaint / allergy'),
        ('#FAEEDA','#633806','Payment to collect'),
        ('#E6F1FB','#042C53','Travelling together'),
        ('#E1F5EE','#04342C','Repeater / stay'),
        ('#E24B4A','#ffffff','Flight - no ETA'),
    ]
    lparts = ['<font size="8" color="#6b7280"><b>Legend:  </b></font>']
    for bg, fg, label in legend_items:
        lparts.append(f'<font size="8" color="{fg}" backColor="{bg}"> {label} </font>  ')
    lt = Table([[Paragraph(''.join(lparts), ps(8))]], colWidths=[W])
    lt.setStyle(TableStyle([
        ('BACKGROUND',(0,0),(-1,-1),C_BG),
        ('BOX',(0,0),(-1,-1),0.5,C_BORDER),
        ('TOPPADDING',(0,0),(-1,-1),7),('BOTTOMPADDING',(0,0),(-1,-1),7),
        ('LEFTPADDING',(0,0),(-1,-1),10),
    ]))
    story.append(lt)

    doc.build(story)
    buf.seek(0)
    return buf.getvalue()


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
                    # Only link via conf number mentions in comments — room number
                    # scanning causes false links from adjacent booking headers
                    if other_b['conf'] and re.search(
                            r'\b' + re.escape(other_b['conf']) + r'\b', comment_text):
                        travelling_together[b['room']].add(other_b['room'])
                        travelling_together[other_b['room']].add(b['room'])

            # Also link rooms sharing the same T# ticket number in comments
            t_refs = re.findall(r'T#\s*(\d{5,})', comment_text)
            if t_refs:
                for other_b in all_bookings:
                    if other_b['room'] == b['room'] and other_b['conf'] == b['conf']:
                        continue
                    other_top_start = other_b['row_top'] + 20
                    other_pg = pdf.pages[other_b['pg_idx']]
                    other_ph = float(other_pg.height)
                    other_words = [w for w in other_pg.extract_words(
                        x_tolerance=2, y_tolerance=2) if w['top'] < other_ph * 0.91]
                    other_page_bkgs = sorted(
                        [bk for bk in all_bookings if bk['pg_idx'] == other_b['pg_idx']],
                        key=lambda x: x['row_top'])
                    other_idx = next((j for j, bk in enumerate(other_page_bkgs)
                                      if bk['room'] == other_b['room']
                                      and bk['conf'] == other_b['conf']), None)
                    if other_idx is None: continue
                    other_top_end = (other_page_bkgs[other_idx+1]['row_top']
                                     if other_idx+1 < len(other_page_bkgs)
                                     else other_ph * 0.91)
                    other_comment = ' '.join(w['text'] for w in other_words
                                             if other_top_start <= w['top'] <= other_top_end)
                    other_t_refs = re.findall(r'T#\s*(\d{5,})', other_comment)
                    # If they share at least one T# reference, link them
                    if set(t_refs) & set(other_t_refs):
                        travelling_together[b['room']].add(other_b['room'])
                        travelling_together[other_b['room']].add(b['room'])

        # Merge all connected rooms into complete groups
        room_to_group = merge_groups(travelling_together)

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

                    together = room_to_group.get(rw['text'], [])
                    popup = ('👥 Travelling Together with Room(s): ' +
                             ', '.join(together)) if together else None

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

        # ══ PASS 4: collect summary guest data ═══════════════════════════
        # Get property name and date from first page
        first_page_words = pdf.pages[0].extract_words(x_tolerance=2, y_tolerance=2)
        property_name = 'Arrival Report'
        arrival_date  = date.today().strftime('%d %B %Y')
        generated     = ''
        for w in first_page_words:
            if 'Anantara' in w['text'] or 'Naladhu' in w['text']:
                # Get full line
                t = w['top']
                line_words = sorted([x for x in first_page_words if abs(x['top']-t)<=3], key=lambda x: x['x0'])
                property_name = ' '.join(x['text'] for x in line_words)
                break
        # Find arrival date from report
        for w in first_page_words:
            if re.match(r'\d{2}/\d{2}/\d{2}', w['text']):
                generated = w['text']
                break
        for w in first_page_words:
            if w['text'] == 'Arrival' and w['x0'] < 100:
                # next token should be 'Date' then the date
                t = w['top']
                line = sorted([x for x in first_page_words if abs(x['top']-t)<=3], key=lambda x: x['x0'])
                line_text = ' '.join(x['text'] for x in line)
                dm = re.search(r'(\d{2}/\d{2}/\d{2})', line_text)
                if dm:
                    parts = dm.group(1).split('/')
                    months = ['','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
                    try:
                        arrival_date = f"{int(parts[0])} {months[int(parts[1])]} 20{parts[2]}"
                    except: pass
                break

        # Collect one row per non-sharer booking with adults > 0
        summary_guests = []
        for pg_idx, page in enumerate(pdf.pages):
            ph = float(page.height)
            words = [w for w in page.extract_words(x_tolerance=2, y_tolerance=2)
                     if w['top'] < ph * 0.91]
            page_bkgs = sorted([b for b in all_bookings if b['pg_idx'] == pg_idx],
                               key=lambda x: x['row_top'])

            processed_s = set()
            for rw in [w for w in words if w['x0'] < 40
                       and re.match(r'^[0-9]{1,4}$', w['text'])]:
                t  = rw['top']
                rk = round(t)
                if rk in processed_s: continue
                processed_s.add(rk)

                row = sorted([w for w in words if abs(w['top'] - t) <= 4],
                             key=lambda x: x['x0'])
                adl  = next((w for w in row if 415 <= w['x0'] <= 435
                             and w['text'].isdigit()), None)
                name = next((w for w in row if 41 <= w['x0'] <= 280), None)
                if not adl or int(adl['text']) == 0: continue
                if not name: continue
                if is_sharer_line(pg_idx, t): continue

                # Conf
                conf = next((w for w in words if w['top'] > t+3 and w['top'] < t+25
                             and 35 <= w['x0'] <= 180
                             and re.match(r'^[0-9]{6,}$', w['text'])), None)

                # Flight on sub-row
                flight_w = next((w for w in words if w['top'] > t+3 and w['top'] < t+25
                                 and w['x0'] > 280 and w['x0'] < 420
                                 and re.match(r'^[A-Z0-9]{2,3}\s*\d{3,4}$', w['text'].replace(' ',''))), None)
                flight_str = ''
                if flight_w:
                    has_eta = any(re.match(r'^([0-1]?[0-9]|2[0-3]):[0-5][0-9]$', w['text'])
                                  for w in words if abs(w['top'] - flight_w['top']) <= 3)
                    flight_str = flight_w['text'] if has_eta else f"{flight_w['text']} NO ETA"

                # Collect flags for this booking
                bflags = []
                # Check categories from counts — scan comment text
                i_bkg = next((j for j, b in enumerate(page_bkgs)
                              if round(b['row_top']) == rk), None)
                top_start = t + 20
                top_end   = (page_bkgs[i_bkg+1]['row_top']
                             if i_bkg is not None and i_bkg+1 < len(page_bkgs)
                             else ph * 0.91)
                ct = ' '.join(w['text'] for w in words if top_start <= w['top'] <= top_end)

                if re.search(r'Membership\s+Level\s+(GOLD|PLATINUM|TITANIUM|RED)', ct, re.I):
                    m = re.search(r'Membership\s+Level\s+(\w+)', ct, re.I)
                    if m: bflags.append(m.group(1).capitalize())
                if re.search(r'VIP[A-Z0-9]', ' '.join(w['text'] for w in row)):
                    vip_w = next((w for w in row if re.match(r'VIP[A-Z0-9]', w['text'])), None)
                    if vip_w: bflags.append(vip_w['text'])
                if re.search(r'\d+(st|nd|rd|th)\s+[Tt]ime\s+RPT|\bRPT\b', ct):
                    m = re.search(r'(\d+(?:st|nd|rd|th)\s+[Tt]ime\s+RPT|\bRPT\b)', ct)
                    if m: bflags.append(m.group(1))
                for kw in ['birthday','anniversary','honeymoon','babymoon','wedding','proposal','engagement']:
                    if kw in ct.lower(): bflags.append(kw.capitalize()); break
                for kw in ['allerg','shellfish','gluten','peanut','lactose','vegan','vegetarian','halal']:
                    if kw in ct.lower():
                        m = re.search(r'(\w*' + kw + r'\w*)', ct, re.I)
                        bflags.append(m.group(1).capitalize() if m else 'Allergy'); break
                if re.search(r'complaint|glitch|feedback|upset', ct, re.I):
                    bflags.append('Complaint')
                if re.search(r'collect upon arrival|to be collected|pls collect|please collect', ct, re.I):
                    bflags.append('Collect')
                if re.search(r'\d+(st|nd|rd|th)\s+[Ll]eg|1ST\s+Leg|2nd\s+Leg', ct):
                    bflags.append('Multi-leg')
                together = room_to_group.get(rw['text'], [])
                if together:
                    others = [r for r in together if r != rw['text']]
                    if others:
                        bflags.append(f"Together: {', '.join(others)}")

                # Note: first meaningful line of comment
                comment_lines = [ln for ln in ct.split('  ') if len(ln.strip()) > 10]
                note = comment_lines[0].strip()[:80] if comment_lines else ''

                summary_guests.append({
                    'room':  rw['text'],
                    'name':  ' '.join(w['text'] for w in row if 41 <= w['x0'] <= 280)[:30],
                    'conf':  conf['text'] if conf else '',
                    'flight': flight_str,
                    'flags': bflags,
                    'note':  note,
                })

    # Build summary data dict
    summary_data = {
        'property':  property_name,
        'date':      arrival_date,
        'generated': generated,
        'total':     len([b for b in all_bookings]),
        'rooms':     len(set(b['room'] for b in all_bookings if not b['is_sharer'])),
        'counts':    counts,
        'guests':    summary_guests,
    }

    # Build summary PDF page and prepend to highlighted PDF
    summary_bytes = build_summary_page(summary_data)

    # Merge: summary page first, then highlighted report
    final_writer = PdfWriter()
    summary_reader = PdfReader(io.BytesIO(summary_bytes))
    for p in summary_reader.pages:
        final_writer.add_page(p)
    highlighted_reader = PdfReader(io.BytesIO((lambda b: (writer.write(b), b)[1])(io.BytesIO())[1] if False else _get_writer_bytes(writer)))
    for p in highlighted_reader.pages:
        final_writer.add_page(p)

    out = io.BytesIO()
    final_writer.write(out)
    return out.getvalue(), counts, total


def _get_writer_bytes(writer):
    buf = io.BytesIO()
    writer.write(buf)
    buf.seek(0)
    return buf.read()


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
