import io, math, re
from pathlib import Path
import streamlit as st
import pandas as pd
from PIL import Image, ImageOps, ImageDraw, ImageFont
from pptx import Presentation
from pptx.util import Inches

BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / 'output'
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

st.set_page_config(page_title='Soft Slides Generator', layout='wide')
st.title('Conference Soft Slides Generator')
st.caption('Upload template, speaker cutout PNG, Excel, photos, and optional font.')

TEMPLATE = st.file_uploader('Upload flat template image', type=['png', 'jpg', 'jpeg'])
CUTOUT = st.file_uploader('Upload single speaker cutout PNG (transparent)', type=['png'])
EXCEL = st.file_uploader('Upload speaker Excel', type=['xlsx', 'xls'])
PHOTOS = st.file_uploader('Upload speaker photos folder (multiple files)', type=['png', 'jpg', 'jpeg'], accept_multiple_files=True)
FONT_FILE = st.file_uploader('Upload font file (optional)', type=['ttf', 'otf'])
EVENT_NAME = st.text_input('Event / session title', value='')
HALL_NAME = st.text_input('Hall name', value='')
DATE_TEXT = st.text_input('Date text', value='')
SHOW_LEADERSHIP = st.checkbox('Show leadership roles above photos only', value=True)

if 'order' not in st.session_state:
    st.session_state.order = []
if 'version' not in st.session_state:
    st.session_state.version = 1

ROLE_PRIORITY = {'MODERATOR': 0, 'CO-MODERATOR': 1, 'CHAIR': 2, 'CO-CHAIR': 3, 'KEYNOTE': 4, 'KEYNOTE SPEAKER': 4, 'PANELIST': 5, 'SPEAKER': 6}
LEADERSHIP_ROLES = {'MODERATOR', 'CO-MODERATOR', 'CHAIR', 'CO-CHAIR', 'KEYNOTE', 'KEYNOTE SPEAKER'}
FIELD_HINT = {'role': ['role'], 'speaker name': ['speaker name', 'name'], 'title': ['title', 'designation'], 'company': ['company', 'org', 'organization']}


def safe(v):
    return '' if pd.isna(v) else str(v)


def clean(s):
    s = re.sub(r'[^a-z0-9]+', ' ', safe(s).lower())
    return re.sub(r'\s+', ' ', s).strip()


def tokens(s):
    return [t for t in clean(s).split() if t]


def pick_col(cols, candidates):
    lower = {str(c).lower().strip(): c for c in cols}
    for cand in candidates:
        if cand in lower:
            return lower[cand]
    return cols[0] if len(cols) else None


def role_rank(role):
    r = clean(role).upper()
    for k, v in ROLE_PRIORITY.items():
        kk = clean(k).upper()
        if r == kk or kk in r:
            return v
    return 99


def ordered_indices_by_role(df):
    role_col = pick_col(df.columns, FIELD_HINT['role'])
    if not role_col:
        return list(df.index)
    roles = df[role_col].astype(str).fillna('')
    return sorted(df.index.tolist(), key=lambda i: (role_rank(roles.loc[i]), i))


def text_size(draw, txt, font):
    box = draw.textbbox((0, 0), txt, font=font)
    return box[2] - box[0], box[3] - box[1]


def get_font(font_path, size):
    if font_path and font_path.exists():
        try:
            return ImageFont.truetype(str(font_path), int(round(size)))
        except Exception:
            pass
    for p in ['/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf', '/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf']:
        try:
            return ImageFont.truetype(p, int(round(size)))
        except Exception:
            pass
    return ImageFont.load_default()


def fit_font(draw, texts, font_path, start=12, min_size=7.5, max_width=300):
    size = start
    while size >= min_size:
        f = get_font(font_path, size)
        if all(text_size(draw, safe(t), f)[0] <= max_width for t in texts):
            return size
        size -= 0.5
    return min_size


def display_name(v):
    v = safe(v).strip()
    parts = v.split()
    if len(parts) >= 3 and len(v) > 20:
        return parts[0] + ' ' + parts[1][0] + '.'
    if len(v) > 28:
        return v[:25].rstrip() + '...'
    return v


def extract_role_from_row(row):
    vals = [safe(x).strip() for x in row.tolist() if safe(x).strip()]
    if not vals:
        return ''
    first = vals[0].upper()
    if first in ['MODERATOR', 'CO-MODERATOR', 'CHAIR', 'CO-CHAIR', 'KEYNOTE', 'KEYNOTE SPEAKER', 'SPEAKER', 'PANELIST']:
        return vals[0]
    return ''


def normalize_photo_key(name):
    s = clean(name)
    for prefix in ['h e ', 'he ', 'shri ', 'smt ', 'dr ', 'prof ', 'mr ', 'mrs ', 'ms ']:
        s = s.replace(prefix, '')
    return clean(s)


def score_match(name, filename):
    n = normalize_photo_key(name)
    f = normalize_photo_key(filename)
    nt = tokens(n)
    ft = tokens(f)
    if not n or not f:
        return 0
    if n == f:
        return 100
    if nt and ft and nt[-1] == ft[-1] and len(nt[-1]) > 2:
        return 90
    common = len(set(nt) & set(ft))
    if common:
        return 20 + common * 10
    if n in f or f in n:
        return 50
    return 0


def assign_photos(df, order, photo_map):
    assigned = {}
    used = set()
    rows = [df.iloc[i] for i in order]
    for idx, row in zip(order, rows):
        name = safe(row.iloc[0]) if len(row) else ''
        best = None
        best_score = 0
        for fname in photo_map:
            if fname in used:
                continue
            sc = score_match(name, fname)
            if sc > best_score:
                best = fname
                best_score = sc
        assigned[idx] = best if best_score >= 20 else None
        if best and best_score >= 20:
            used.add(best)
    return assigned


def role_label(role):
    r = clean(role).upper()
    if 'CO-MODERATOR' in r:
        return 'CO-MODERATOR'
    if 'MODERATOR' in r:
        return 'MODERATOR'
    if 'CO-CHAIR' in r:
        return 'CO-CHAIR'
    if 'CHAIR' in r:
        return 'CHAIR'
    if 'KEYNOTE' in r:
        return 'KEYNOTE SPEAKER'
    return ''


def layout(n):
    if n <= 2:
        return 2, 1
    if n <= 4:
        return 2, 2
    if n <= 6:
        return 3, 2
    if n <= 9:
        return 3, 3
    return 4, 4


if EXCEL is not None:
    df = pd.read_excel(EXCEL, header=None)
    st.subheader('Speaker Data')
    st.dataframe(df, use_container_width=True)

    if st.button('Auto order by role'):
        st.session_state.order = ordered_indices_by_role(df)
    if not st.session_state.order:
        st.session_state.order = list(df.index)

    st.subheader('Reorder speakers')
    st.caption('Leadership roles stay first when explicitly provided.')
    for i, idx in enumerate(list(st.session_state.order)):
        cols = st.columns([6, 1, 1])
        row = df.iloc[idx]
        row_role = extract_role_from_row(row)
        row_name = safe(row.iloc[0]) if len(row) else ''
        cols[0].write(f'{i+1}. {row_name} {"(" + row_role + ")" if row_role else ""}')
        if cols[1].button('▲', key=f'u{idx}', disabled=i == 0):
            st.session_state.order[i - 1], st.session_state.order[i] = st.session_state.order[i], st.session_state.order[i - 1]
            st.rerun()
        if cols[2].button('▼', key=f'd{idx}', disabled=i == len(st.session_state.order) - 1):
            st.session_state.order[i + 1], st.session_state.order[i] = st.session_state.order[i], st.session_state.order[i + 1]
            st.rerun()

    if PHOTOS:
        st.subheader('Photo matching review')
        photo_map = {Path(f.name).stem: f for f in PHOTOS}
        assigned = assign_photos(df, st.session_state.order, photo_map)
        for idx in st.session_state.order:
            row = df.iloc[idx]
            name = safe(row.iloc[0]) if len(row) else ''
            cols = st.columns([3, 2, 4])
            cols[0].write(name)
            cols[1].write(assigned.get(idx) if assigned.get(idx) else 'NO MATCH')
            cols[2].write('')

    if st.button('Generate Downloads'):
        if TEMPLATE is None or CUTOUT is None:
            st.error('Template and cutout PNG are required.')
            st.stop()
        if not PHOTOS:
            st.error('Please upload speaker photos.')
            st.stop()

        photo_map = {Path(f.name).stem: f for f in PHOTOS}
        assigned = assign_photos(df, st.session_state.order, photo_map)
        base_img = Image.open(TEMPLATE).convert('RGBA')
        cutout_img = Image.open(CUTOUT).convert('RGBA')
        cutout_size = cutout_img.size
        W, H = base_img.size
        base_draw = ImageDraw.Draw(base_img)

        font_path = None
        if FONT_FILE is not None:
            font_path = OUTPUT_DIR / FONT_FILE.name
            font_path.write_bytes(FONT_FILE.getbuffer())

        parsed = []
        for i in st.session_state.order:
            row = df.iloc[i]
            vals = [safe(x).strip() for x in row.tolist() if safe(x).strip()]
            role = extract_role_from_row(row)
            if role and len(vals) > 1:
                vals = vals[1:]
            elif not role and vals:
                vals = vals[1:]
            name = safe(row.iloc[0]) if len(row) else ''
            title = vals[0] if len(vals) > 0 else ''
            company = vals[1] if len(vals) > 1 else ''
            parsed.append((role, name, title, company))

        names = [p[1] for p in parsed]
        name_size = fit_font(base_draw, names, font_path, start=14, min_size=8, max_width=int(W * 0.16))
        title_size = max(8, name_size - 2)
        role_size = max(8, title_size)
        header_size = 26
        small_size = 16

        def render_slide(indices, parsed_rows):
            img = base_img.copy()
            d = ImageDraw.Draw(img)

            header_y = int(H * 0.045)
            hall_y = int(H * 0.085)
            date_y = int(H * 0.115)
            for txt, y, fs in [(EVENT_NAME, header_y, header_size), (HALL_NAME, hall_y, small_size), (DATE_TEXT, date_y, small_size)]:
                if txt.strip():
                    f = get_font(font_path, fs)
                    tw, th = text_size(d, txt, f)
                    d.text(((W - tw) / 2, y), txt, font=f, fill=(255, 255, 255, 255))

            n = len(indices)
            cols, rows = layout(n)
            x_margin = int(W * 0.06)
            y_start = int(H * 0.30)
            usable_w = W - 2 * x_margin
            usable_h = int(H * 0.54)
            cell_w = usable_w / cols
            cell_h = usable_h / rows
            photo_w = int(min(cell_w, cell_h) * 0.72)
            photo_h = int(photo_w * cutout_size[1] / cutout_size[0]) if cutout_size[0] else photo_w
            photo_h = max(photo_h, 1)

            d.text((int(W * 0.50), int(H * 0.24)), 'SPEAKERS', font=get_font(font_path, 20), fill=(255, 255, 255, 255), anchor='mm')

            for pos, (idx, rowdata) in enumerate(zip(indices, parsed_rows)):
                role, name, title, company = rowdata
                row, col = divmod(pos, cols)
                x = int(x_margin + col * cell_w + (cell_w - photo_w) / 2)
                y = int(y_start + row * cell_h)

                matched_key = assigned.get(idx)
                if matched_key:
                    ph = Image.open(photo_map[matched_key]).convert('RGBA')
                    ph = ImageOps.fit(ph, (photo_w, photo_h), method=Image.Resampling.LANCZOS, centering=(0.5, 0.33))
                    mask = cutout_img.resize((photo_w, photo_h)).split()[-1]
                    img.paste(ph, (x, y), mask)
                else:
                    outline = cutout_img.convert('RGBA').resize((photo_w, photo_h))
                    img.alpha_composite(outline, (x, y))
                    d.text((x, y + photo_h + 2), f'No photo for {name}', font=get_font(font_path, 8), fill=(255, 180, 180, 255))

                nf = get_font(font_path, name_size)
                tf = get_font(font_path, title_size)
                cf = get_font(font_path, title_size)
                rf = get_font(font_path, role_size)

                disp = display_name(name)
                nw, _ = text_size(d, disp, nf)
                d.text((x + (photo_w - nw) / 2, y + photo_h + 8), disp, font=nf, fill=(255, 
