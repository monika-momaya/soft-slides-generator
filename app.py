import io, math, zipfile
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
st.caption('Upload flat template, Excel, photos, and optional font. Generate PNG + PPTX downloads.')

TEMPLATE = st.file_uploader('Upload flat template image', type=['png', 'jpg', 'jpeg'])
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

ROLE_PRIORITY = {'MODERATOR': 0, 'CHAIR': 1, 'PANELIST': 2, 'SPEAKER': 3}
FIELD_HINT = {'speaker name': ['speaker name', 'name'], 'title': ['title', 'designation'], 'company': ['company', 'org', 'organization'], 'role': ['role']}


def pick_col(cols, candidates):
    for cand in candidates:
        for c in cols:
            if cand == c.lower().strip():
                return c
    return cols[0] if len(cols) else None


def safe(v):
    return '' if pd.isna(v) else str(v)


def ordered_indices_by_role(df):
    role_col = pick_col(df.columns, FIELD_HINT['role'])
    if not role_col:
        return list(df.index)
    roles = df[role_col].astype(str).str.upper().fillna('')
    return sorted(df.index.tolist(), key=lambda i: (ROLE_PRIORITY.get(roles.loc[i], 99), i))


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


def crop_portrait(im, size):
    im = ImageOps.fit(im, (size, size), method=Image.Resampling.LANCZOS, centering=(0.5, 0.33))
    mask = Image.new('L', (size, size), 0)
    ImageDraw.Draw(mask).ellipse((0, 0, size - 1, size - 1), fill=255)
    out = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    out.paste(im, (0, 0), mask)
    return out


def fit_box(draw, txt, font_path, target_w, start=12, min_size=7.5):
    size = start
    while size >= min_size:
        if text_size(draw, txt, get_font(font_path, size))[0] <= target_w:
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


if EXCEL is not None:
    df = pd.read_excel(EXCEL)
    name_col = pick_col(df.columns, FIELD_HINT['speaker name'])
    title_col = pick_col(df.columns, FIELD_HINT['title'])
    company_col = pick_col(df.columns, FIELD_HINT['company'])
    role_col = pick_col(df.columns, FIELD_HINT['role'])
    st.subheader('Speaker Data')
    st.dataframe(df, use_container_width=True)

    if st.button('Auto order by role'):
        st.session_state.order = ordered_indices_by_role(df)
    if not st.session_state.order:
        st.session_state.order = list(df.index)

    st.subheader('Reorder speakers')
    st.caption('Use the arrows to reorder. The generator will preserve the final sequence.')
    for i, idx in enumerate(list(st.session_state.order)):
        cols = st.columns([6, 1, 1])
        display = safe(df.iloc[idx].get(name_col, df.iloc[idx].iloc[0]))
        role = safe(df.iloc[idx].get(role_col, '')) if role_col else ''
        cols[0].write(f'{i+1}. {display} {"(" + role + ")" if role else ""}')
        if cols[1].button('▲', key=f'u{idx}', disabled=i == 0):
            st.session_state.order[i - 1], st.session_state.order[i] = st.session_state.order[i], st.session_state.order[i - 1]
            st.rerun()
        if cols[2].button('▼', key=f'd{idx}', disabled=i == len(st.session_state.order) - 1):
            st.session_state.order[i + 1], st.session_state.order[i] = st.session_state.order[i], st.session_state.order[i + 1]
            st.rerun()

    if st.button('Generate Downloads'):
        if TEMPLATE is None:
            st.error('Template is required.')
            st.stop()
        photo_map = {Path(f.name).stem.lower(): f for f in PHOTOS} if PHOTOS else {}
        base_img = Image.open(TEMPLATE).convert('RGBA')
        W, H = base_img.size
        base_draw = ImageDraw.Draw(base_img)

        font_path = None
        if FONT_FILE is not None:
            font_path = OUTPUT_DIR / FONT_FILE.name
            font_path.write_bytes(FONT_FILE.getbuffer())

        names = [safe(df.iloc[i].get(name_col, '')) for i in st.session_state.order]
        name_size = fit_font(base_draw, names, font_path, start=12, min_size=7.5, max_width=int(W * 0.16))
        title_size = max(7.5, name_size - 2)
        role_size = max(7.5, title_size)
        speaker_count = len(st.session_state.order)
        slide_count = 1 if speaker_count <= 9 else math.ceil(speaker_count / 9)
        per_slide = math.ceil(speaker_count / slide_count)

        def layout(n):
            if n <= 3:
                return 1, n
            if n <= 6:
                return 2, math.ceil(n / 2)
            return 3, math.ceil(n / 3)

        def render_slide(indices):
            img = base_img.copy()
            d = ImageDraw.Draw(img)
            head_specs = [(EVENT_NAME, 24, 0.07), (HALL_NAME, 18, 0.13), (DATE_TEXT, 16, 0.17)]
            for txt, base_fs, y_ratio in head_specs:
                if txt.strip():
                    fs = fit_box(d, txt, font_path, int(W * 0.75), start=base_fs, min_size=7.5)
                    f = get_font(font_path, fs)
                    tw, _ = text_size(d, txt, f)
                    d.text(((W - tw) / 2, int(H * y_ratio)), txt, font=f, fill=(255, 255, 255, 255))

            cols, rows = layout(len(indices))
            x_margin = int(W * 0.06)
            y_start = int(H * 0.23)
            usable_w = W - 2 * x_margin
            usable_h = int(H * 0.58)
            cell_w = usable_w / cols
            cell_h = usable_h / rows
            photo_size = int(min(cell_w, cell_h) * 0.5)

            for pos, idx in enumerate(indices):
                rec = df.iloc[idx]
                full_name = safe(rec.get(name_col, ''))
                short_name = display_name(full_name)
                title = safe(rec.get(title_col, ''))
                company = safe(rec.get(company_col, ''))
                role = safe(rec.get(role_col, '')) if role_col else ''
                row, col = divmod(pos, cols)
                x = int(x_margin + col * cell_w + (cell_w - photo_size) / 2)
                y = int(y_start + row * cell_h)
                key = Path(full_name).stem.lower()
                if key in photo_map:
                    ph = Image.open(photo_map[key]).convert('RGBA')
                    ph = crop_portrait(ph, photo_size)
                    img.alpha_composite(ph, (x, y))
                else:
                    d.ellipse((x, y, x + photo_size, y + photo_size), outline=(255, 255, 255, 255), width=3)

                nf = get_font(font_path, name_size)
                tf = get_font(font_path, title_size)
                cf = get_font(font_path, title_size)
                rf = get_font(font_path, role_size)

                name_w, _ = text_size(d, short_name, nf)
                name_x = x + (photo_size - name_w) / 2
                d.text((name_x, y + photo_size + 8), short_name, font=nf, fill=(255, 235, 80, 255))
                if role and SHOW_ROLES:
                    d.text((x, y - 22), role.upper(), font=rf, fill=(255, 255, 255, 255))
                d.text((x, y + photo_size + 25), title, font=tf, fill=(240, 240, 240, 255))
                d.text((x, y + photo_size + 43), company, font=cf, fill=(240, 240, 240, 255))
            return img

        png_files = []
        ppt = Presentation()
        ppt.slide_width = Inches(13.333)
        ppt.slide_height = Inches(7.5)
        for s in range(slide_count):
            start = s * per_slide
            end = min(start + per_slide, speaker_count)
            slide_img = render_slide(st.session_state.order[start:end])
            buf = io.BytesIO()
            slide_img.save(buf, format='PNG')
            png_files.append(buf.getvalue())
            slide = ppt.slides.add_slide(ppt.slide_layouts[6])
            tmp = OUTPUT_DIR / f'softslide_{st.session_state.version}_{s+1}.png'
            tmp.write_bytes(buf.getvalue())
            slide.shapes.add_picture(str(tmp), 0, 0, width=ppt.slide_width, height=ppt.slide_height)

        ppt_buf = io.BytesIO()
        ppt.save(ppt_buf)
        ppt_buf.seek(0)

        if len(png_files) == 1:
            st.download_button('Download PNG', data=png_files[0], file_name=f'soft_slide_v{st.session_state.version}.png', mime='image/png')
        else:
            zbuf = io.BytesIO()
            with zipfile.ZipFile(zbuf, 'w', zipfile.ZIP_DEFLATED) as z:
                for i, b in enumerate(png_files, 1):
                    z.writestr(f'soft_slide_{i}.png', b)
            zbuf.seek(0)
            st.download_button('Download PNGs (ZIP)', data=zbuf.getvalue(), file_name=f'soft_slides_v{st.session_state.version}.zip', mime='application/zip')
        st.download_button('Download PPTX', data=ppt_buf.getvalue(), file_name=f'soft_slide_v{st.session_state.version}.pptx', mime='application/vnd.openxmlformats-officedocument.presentationml.presentation')
        st.session_state.version += 1
        st.success(f'Generated {slide_count} slide(s).')
