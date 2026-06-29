import io
import re
from pathlib import Path

import pandas as pd
import streamlit as st
from PIL import Image, ImageDraw, ImageFont, ImageOps
from pptx import Presentation
from pptx.util import Inches

BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / 'output'
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

st.set_page_config(page_title='Soft Slides Generator', layout='wide')
st.title('Conference Soft Slides Generator')
st.caption('Upload template, placeholder PNG, Excel, photos, and optional font.')

TEMPLATE = st.file_uploader('Upload flat template image', type=['png', 'jpg', 'jpeg'])
CUTOUT = st.file_uploader('Upload speaker placeholder PNG', type=['png'])
EXCEL = st.file_uploader('Upload speaker Excel', type=['xlsx', 'xls'])
PHOTOS = st.file_uploader('Upload speaker photos', type=['png', 'jpg', 'jpeg'], accept_multiple_files=True)
FONT_FILE = st.file_uploader('Upload font file (optional)', type=['ttf', 'otf'])
EVENT_NAME = st.text_input('Event / session title', value='')
HALL_NAME = st.text_input('Hall name', value='')
DATE_TEXT = st.text_input('Date text', value='')

if 'order' not in st.session_state:
    st.session_state.order = []
if 'mapping' not in st.session_state:
    st.session_state.mapping = {}

ROLE_PRIORITY = ['MODERATOR', 'CO-MODERATOR', 'CHAIR', 'CO-CHAIR', 'KEYNOTE', 'KEYNOTE SPEAKER', 'PANELIST', 'SPEAKER']
LEADERSHIP = {'MODERATOR', 'CO-MODERATOR', 'CHAIR', 'CO-CHAIR', 'KEYNOTE', 'KEYNOTE SPEAKER'}


def safe(v):
    return '' if pd.isna(v) else str(v)


def clean(s):
    return re.sub(r'\s+', ' ', re.sub(r'[^a-z0-9]+', ' ', safe(s).lower())).strip()


def strip_prefixes(s):
    s = clean(s)
    for p in ['h e ', 'he ', 'shri ', 'smt ', 'dr ', 'prof ', 'mr ', 'mrs ', 'ms ']:
        s = s.replace(p, '')
    return clean(s)


def name_tokens(s):
    return [t for t in strip_prefixes(s).split() if t]


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


def text_size(draw, txt, font):
    box = draw.textbbox((0, 0), txt, font=font)
    return box[2] - box[0], box[3] - box[1]


def fit_font(draw, texts, font_path, start, min_size, max_width):
    size = start
    while size >= min_size:
        f = get_font(font_path, size)
        if all(text_size(draw, t, f)[0] <= max_width for t in texts):
            return size
        size -= 0.5
    return min_size


def role_from_row(row):
    vals = [safe(x).strip() for x in row.tolist() if safe(x).strip()]
    for v in vals:
        if clean(v).upper() in ROLE_PRIORITY:
            return v
    return ''


def pick_name(row):
    vals = [safe(x).strip() for x in row.tolist() if safe(x).strip()]
    if not vals:
        return ''
    r = role_from_row(row)
    if r and clean(vals[0]).upper() == clean(r).upper() and len(vals) > 1:
        return vals[1]
    return vals[0]


def pick_title_company(row):
    vals = [safe(x).strip() for x in row.tolist() if safe(x).strip()]
    r = role_from_row(row)
    if r and vals and clean(vals[0]).upper() == clean(r).upper():
        vals = vals[1:]
    elif vals:
        vals = vals[1:]
    title = vals[0] if len(vals) > 0 else ''
    company = vals[1] if len(vals) > 1 else ''
    return title, company


def display_name(v):
    v = safe(v).strip()
    parts = v.split()
    if len(parts) >= 3 and len(v) > 20:
        return parts[0] + ' ' + parts[1][0] + '.'
    if len(v) > 28:
        return v[:25].rstrip() + '...'
    return v


def match_score(name, filename):
    n = strip_prefixes(name)
    f = strip_prefixes(filename)
    nt = name_tokens(n)
    ft = name_tokens(f)
    if not n or not f:
        return 0
    if n == f:
        return 100
    if nt and ft and nt[-1] == ft[-1] and len(nt[-1]) > 2:
        return 92 if len(set(nt) & set(ft)) >= 1 else 80
    if n in f or f in n:
        return 70
    common = len(set(nt) & set(ft))
    return 20 + common * 10 if common else 0


def unique_match(name, photo_names, used):
    scored = [(match_score(name, p), p) for p in photo_names if p not in used]
    scored.sort(reverse=True)
    if not scored or scored[0][0] < 20:
        return None, 0, scored[:5]
    top = scored[0][0]
    tied = [p for s, p in scored if s == top]
    return (tied[0] if len(tied) == 1 else None), top, scored[:5]


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

    if st.button('Build speaker order from role priority'):
        rows = []
        for i in df.index:
            row = df.iloc[i]
            role = role_from_row(row)
            rows.append((i, 0 if clean(role).upper().startswith('MODERATOR') else 1 if 'CHAIR' in clean(role).upper() else 2, i))
        rows.sort()
        st.session_state.order = [r[0] for r in rows]

    if not st.session_state.order:
        st.session_state.order = list(df.index)

    st.subheader('Manual order')
    for i, idx in enumerate(list(st.session_state.order)):
        c1, c2, c3 = st.columns([6, 1, 1])
        row = df.iloc[idx]
        nm = pick_name(row)
        rl = role_from_row(row)
        c1.write(f'{i+1}. {nm} {"(" + rl + ")" if rl else ""}')
        if c2.button('▲', key=f'u{idx}', disabled=i == 0):
            st.session_state.order[i - 1], st.session_state.order[i] = st.session_state.order[i], st.session_state.order[i - 1]
            st.rerun()
        if c3.button('▼', key=f'd{idx}', disabled=i == len(st.session_state.order) - 1):
            st.session_state.order[i + 1], st.session_state.order[i] = st.session_state.order[i], st.session_state.order[i + 1]
            st.rerun()

    if PHOTOS:
        st.subheader('Preview matches')
        photo_map = {Path(f.name).stem: f for f in PHOTOS}
        photo_names = list(photo_map.keys())
        used = set()
        mapping = {}
        for idx in st.session_state.order:
            row = df.iloc[idx]
            nm = pick_name(row)
            best, score, top5 = unique_match(nm, photo_names, used)
            mapping[idx] = best
            if best:
                used.add(best)
            c1, c2, c3, c4 = st.columns([3, 2, 1, 4])
            c1.write(nm)
            c2.write(best if best else 'NO MATCH')
            c3.write(score)
            c4.write(', '.join([f'{p}:{s}' for s, p in top5]))
        st.session_state.mapping = mapping
        st.caption('If any match is wrong, rename the photo file or move the row order before generating.')

    if st.button('Generate Output'):
        if TEMPLATE is None or CUTOUT is None:
            st.error('Template and placeholder PNG are required.')
            st.stop()
        if not PHOTOS:
            st.error('Please upload speaker photos.')
            st.stop()

        photo_map = {Path(f.name).stem: f for f in PHOTOS}
        base = Image.open(TEMPLATE).convert('RGBA')
        cutout = Image.open(CUTOUT).convert('RGBA')
        W, H = base.size
        d = ImageDraw.Draw(base)
        fp = None
        if FONT_FILE is not None:
            fp = OUTPUT_DIR / FONT_FILE.name
            fp.write_bytes(FONT_FILE.getbuffer())

        def centered(txt, y, size, fill=(255, 255, 255, 255)):
            if not txt.strip():
                return
            f = get_font(fp, size)
            tw, _ = text_size(d, txt, f)
            d.text(((W - tw) / 2, y), txt, font=f, fill=fill)

        centered(EVENT_NAME, int(H * 0.04), 34)
        centered(HALL_NAME, int(H * 0.08), 20)
        centered(DATE_TEXT, int(H * 0.11), 20)
        centered('SPEAKERS', int(H * 0.24), 22)

        names = [pick_name(df.iloc[i]) for i in st.session_state.order]
        name_sz = fit_font(d, names, fp, 16, 8, int(W * 0.16))
        title_sz = max(8, name_sz - 2)
        role_sz = max(8, title_sz)

        cols, rows = layout(len(st.session_state.order))
        x_margin = int(W * 0.06)
        y_start = int(H * 0.30)
        usable_w = W - 2 * x_margin
        usable_h = int(H * 0.54)
        cell_w = usable_w / cols
        cell_h = usable_h / rows
        photo_w = int(min(cell_w, cell_h) * 0.72)
        photo_h = int(photo_w * cutout.size[1] / cutout.size[0]) if cutout.size[0] else photo_w
        mask = cutout.resize((photo_w, photo_h)).split()[-1]

        for pos, idx in enumerate(st.session_state.order):
            row = df.iloc[idx]
            nm = pick_name(row)
            rl = role_from_row(row)
            title, company = pick_title_company(row)
            r, c = divmod(pos, cols)
            x = int(x_margin + c * cell_w + (cell_w - photo_w) / 2)
            y = int(y_start + r * cell_h)

            mapped = st.session_state.mapping.get(idx)
            if mapped and mapped in photo_map:
                ph = Image.open(photo_map[mapped]).convert('RGBA')
                ph = ImageOps.fit(ph, (photo_w, photo_h), method=Image.Resampling.LANCZOS, centering=(0.5, 0.33))
                base.paste(ph, (x, y), mask)
            else:
                outline = cutout.resize((photo_w, photo_h))
                base.alpha_composite(outline, (x, y))

            name_f = get_font(fp, name_sz)
            title_f = get_font(fp, title_sz)
            comp_f = get_font(fp, title_sz)
            role_f = get_font(fp, role_sz)

            disp = display_name(nm)
            nw, _ = text_size(d, disp, name_f)
            d.text((x + (photo_w - nw) / 2, y + photo_h + 8), disp, font=name_f, fill=(255, 235, 80, 255))
            tw, _ = text_size(d, title, title_f)
            cw, _ = text_size(d, company, comp_f)
            d.text((x + (photo_w - tw) / 2, y + photo_h + 25), title, font=title_f, fill=(240, 240, 240, 255))
            d.text((x + (photo_w - cw) / 2, y + photo_h + 43), company, font=comp_f, fill=(240, 240, 240, 255))
            lab = role_from_row(row)
            if lab and clean(lab).upper() in LEADERSHIP:
                d.text((x + photo_w / 2, y - 20), lab.upper(), font=role_f, fill=(255, 255, 255, 255), anchor='mm')

        buf = io.BytesIO()
        base.save(buf, format='PNG')
        ppt = Presentation()
        ppt.slide_width = Inches(13.333)
        ppt.slide_height = Inches(7.5)
        slide = ppt.slides.add_slide(ppt.slide_layouts[6])
        tmp = OUTPUT_DIR / 'softslide_render.png'
        tmp.write_bytes(buf.getvalue())
        slide.shapes.add_picture(str(tmp), 0, 0, width=ppt.slide_width, height=ppt.slide_height)
        ppt_buf = io.BytesIO()
        ppt.save(ppt_buf)
        ppt_buf.seek(0)
        st.download_button('Download PNG', data=buf.getvalue(), file_name='soft_slide.png', mime='image/png')
        st.download_button('Download PPTX', data=ppt_buf.getvalue(), file_name='soft_slide.pptx', mime='application/vnd.openxmlformats-officedocument.presentationml.presentation')
