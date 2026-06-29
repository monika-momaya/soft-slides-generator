import io, math, zipfile, re
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
st.caption('Upload template, single speaker cutout PNG, Excel, photos, and optional font.')

TEMPLATE = st.file_uploader('Upload flat template image', type=['png', 'jpg', 'jpeg'])
CUTOUT = st.file_uploader('Upload single speaker cutout PNG (transparent)', type=['png'])
EXCEL = st.file_uploader('Upload speaker Excel', type=['xlsx', 'xls'])
PHOTOS = st.file_uploader('Upload speaker photos folder (multiple files)', type=['png', 'jpg', 'jpeg'], accept_multiple_files=True)
FONT_FILE = st.file_uploader('Upload font file (optional)', type=['ttf', 'otf'])
EVENT_NAME = st.text_input('Event / session title', value='')
HALL_NAME = st.text_input('Hall name', value='')
DATE_TEXT = st.text_input('Date text', value='')
SHOW_ROLES = st.checkbox('Show roles above speakers', value=True)

if 'order' not in st.session_state:
    st.session_state.order = []
if 'version' not in st.session_state:
    st.session_state.version = 1

ROLE_PRIORITY = {'MODERATOR': 0, 'CHAIR': 1, 'KEYNOTE': 2, 'SPEAKER': 3, 'PANELIST': 4}
FIELD_HINT = {'speaker name': ['speaker name', 'name'], 'title': ['title', 'designation'], 'company': ['company', 'org', 'organization'], 'role': ['role']}


def pick_col(cols, candidates):
    for cand in candidates:
        for c in cols:
            if cand == str(c).lower().strip():
                return c
    return cols[0] if len(cols) else None


def safe(v):
    return '' if pd.isna(v) else str(v)


def ordered_indices_by_role(df):
    role_col = pick_col(df.columns, FIELD_HINT['role'])
    if not role_col:
        return list(df.index)
    roles = df[role_col].astype(str).fillna('')
    return sorted(df.index.tolist(), key=lambda i: (ROLE_PRIORITY.get(roles.loc[i].split()[0].upper(), 99), i))


def text_size(draw, txt, font):
    box = draw.textbbox((0, 0), txt, font=font)
    return box[2] - box[0], box[3] - box[1]


def get_font(font_path, size):
    if font_path and font_path.exists():
        try:
            return ImageFont.truetype(str(font_path), int(round(size)))
        except Exception:
            pass
    try:
        return ImageFont.truetype('/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf', int(round(size)))
    except Exception:
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


def guess_row_fields(row):
    vals = [safe(x).strip() for x in row.tolist() if safe(x).strip()]
    role = ''
    name = ''
    extras = []
    for v in vals:
        up = v.upper()
        if any(k in up for k in ['MODERATOR', 'CHAIR', 'KEYNOTE', 'PANELIST', 'SPEAKER']):
            role = v
        elif not name:
            name = v
        else:
            extras.append(v)
    title, company = '', ''
    if len(extras) >= 2:
        title, company = extras[0], extras[1]
    elif len(extras) == 1:
        tc = extras[0]
        if '|' in tc:
            title, company = [x.strip() for x in tc.split('|', 1)]
        elif ',' in tc:
            title, company = [x.strip() for x in tc.split(',', 1)]
        else:
            title = tc
    return role, name, title, company


if EXCEL is not None:
    df = pd.read_excel(EXCEL)
    st.subheader('Speaker Data')
    st.dataframe(df, use_container_width=True)

    if st.button('Auto order by role'):
        st.session_state.order = ordered_indices_by_role(df)
    if not st.session_state.order:
        st.session_state.order = list(df.index)

    st.subheader('Reorder speakers')
    st.caption('Use the arrows to reorder. Moderator / Chair / Keynote are prioritized automatically when possible.')
    for i, idx in enumerate(list(st.session_state.order)):
        cols = st.columns([6, 1, 1])
        row = df.iloc[idx]
        row_role, row_name, _, _ = guess_row_fields(row)
        cols[0].write(f'{i+1}. {row_name or safe(row.iloc[0])} {"(" + row_role + ")" if row_role else ""}')
        if cols[1].button('▲', key=f'u{idx}', disabled=i == 0):
            st.session_state.order[i - 1], st.session_state.order[i] = st.session_state.order[i], st.session_state.order[i - 1]
            st.rerun()
        if cols[2].button('▼', key=f'd{idx}', disabled=i == len(st.session_state.order) - 1):
            st.session_state.order[i + 1], st.session_state.order[i] = st.session_state.order[i], st.session_state.order[i + 1]
            st.rerun()

    if st.button('Generate Downloads'):
        if TEMPLATE is None or CUTOUT is None:
            st.error('Template and cutout PNG are required.')
            st.stop()
        photo_map = {Path(f.name).stem.lower(): f for f in PHOTOS} if PHOTOS else {}
        base_img = Image.open(TEMPLATE).convert('RGBA')
        cutout_img = Image.open(CUTOUT).convert('RGBA')
        cutout_size = cutout_img.size
        W, H = base_img.size
        base_draw = ImageDraw.Draw(base_img)

        font_path = None
        if FONT_FILE is not None:
            font_path = OUTPUT_DIR / FONT_FILE.name
            font_path.write_bytes(FONT_FILE.getbuffer())

        parsed = [guess_row_fields(df.iloc[i]) for i in st.session_state.order]
        names = [p[1] for p in parsed]
        name_size = fit_font(base_draw, names, font_path, start=12, min_size=7.5, max_width=int(W * 0.16))
        title_size = max(7.5, name_size - 2)
        role_size = max(7.5, title_size)

        def render_slide(indices, parsed_rows):
            img = base_img.copy()
            d = ImageDraw.Draw(img)
            for txt, base_fs, y_ratio in [(EVENT_NAME, 24, 0.07), (HALL_NAME, 18, 0.13), (DATE_TEXT, 16, 0.17)]:
                if txt.strip():
                    fs = fit_font(d, [txt], font_path, start=base_fs, min_size=7.5, max_width=int(W * 0.75))
                    f = get_font(font_path, fs)
                    tw, _ = text_size(d, txt, f)
                    d.text(((W - tw) / 2, int(H * y_ratio)), txt, font=f, fill=(255, 255, 255, 255))

            n = len(indices)
            cols = 1 if n == 1 else 2 if n <= 6 else 3 if n <= 9 else 4
            rows = math.ceil(n / cols)
            x_margin = int(W * 0.06)
            y_start = int(H * 0.24)
            usable_w = W - 2 * x_margin
            usable_h = int(H * 0.62)
            cell_w = usable_w / cols
            cell_h = usable_h / rows
            photo_w = int(min(cell_w, cell_h) * 0.40)
            photo_h = int(photo_w * cutout_size[1] / cutout_size[0]) if cutout_size[0] else photo_w
            photo_h = max(photo_h, 1)

            for pos, (idx, rowdata) in enumerate(zip(indices, parsed_rows)):
                role, name, title, company = rowdata
                row, col = divmod(pos, cols)
                x = int(x_margin + col * cell_w + (cell_w - photo_w) / 2)
                y = int(y_start + row * cell_h)
                key = Path(name).stem.lower()
                if key in photo_map:
                    ph = Image.open(photo_map[key]).convert('RGBA')
                    ph = ImageOps.fit(ph, (photo_w, photo_h), method=Image.Resampling.LANCZOS, centering=(0.5, 0.33))
                    mask = Image.new('L', (photo_w, photo_h), 0)
                    ImageDraw.Draw(mask).ellipse((0, 0, photo_w - 1, photo_h - 1), fill=255)
                    img.paste(ph, (x, y), mask)
                else:
                    outline = Image.new('RGBA', (photo_w, photo_h), (0, 0, 0, 0))
                    ImageDraw.Draw(outline).ellipse((0, 0, photo_w - 1, photo_h - 1), outline=(255, 255, 255, 255), width=3)
                    img.alpha_composite(outline, (x, y))
                nf = get_font(font_path, name_size)
                tf = get_font(font_path, title_size)
                cf = get_font(font_path, title_size)
                rf = get_font(font_path, role_size)
                disp = display_name(name)
                nw, _ = text_size(d, disp, nf)
                d.text((x + (photo_w - nw) / 2, y + photo_h + 8), disp, font=nf, fill=(255, 235, 80, 255))
                if role and SHOW_ROLES:
                    d.text((x, y - 22), role.upper(), font=rf, fill=(255, 255, 255, 255))
                d.text((x, y + photo_h + 25), title, font=tf, fill=(240, 240, 240, 255))
                d.text((x, y + photo_h + 43), company, font=cf, fill=(240, 240, 240, 255))
            return img

        slide_img = render_slide(st.session_state.order, parsed)
        buf = io.BytesIO()
        slide_img.save(buf, format='PNG')

        ppt = Presentation()
        ppt.slide_width = Inches(13.333)
        ppt.slide_height = Inches(7.5)
        slide = ppt.slides.add_slide(ppt.slide_layouts[6])
        tmp = OUTPUT_DIR / f'softslide_{st.session_state.version}_1.png'
        tmp.write_bytes(buf.getvalue())
        slide.shapes.add_picture(str(tmp), 0, 0, width=ppt.slide_width, height=ppt.slide_height)

        ppt_buf = io.BytesIO()
        ppt.save(ppt_buf)
        ppt_buf.seek(0)
        st.download_button('Download PNG', data=buf.getvalue(), file_name=f'soft_slide_v{st.session_state.version}.png', mime='image/png')
        st.download_button('Download PPTX', data=ppt_buf.getvalue(), file_name=f'soft_slide_v{st.session_state.version}.pptx', mime='application/vnd.openxmlformats-officedocument.presentationml.presentation')
        st.session_state.version += 1
        st.success('Generated 1 slide.')
