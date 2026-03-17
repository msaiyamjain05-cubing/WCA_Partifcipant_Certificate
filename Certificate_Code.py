import io
import os
import sys
from copy import deepcopy

import pandas as pd

# PDF / drawing libs
from reportlab.pdfgen import canvas
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.colors import HexColor, Color
from reportlab.lib.utils import ImageReader
from reportlab.graphics import renderPDF
from reportlab.graphics.shapes import Path, String
from svglib.svglib import svg2rlg

# Prefer CairoSVG for accurate SVG rendering (evenodd/clip-rule, etc.)
try:
    import cairosvg  # type: ignore
    _CairoSVGAvailable = True
except Exception:
    cairosvg = None
    _CairoSVGAvailable = False

# Windows-friendly SVG renderer (no system cairo dll needed)
try:
    import resvg_py  # type: ignore
except Exception:
    resvg_py = None

try:
    from PIL import Image, ImageOps  # type: ignore
except Exception:
    Image = None
    ImageOps = None

# PDF readers/writers
try:
    from pypdf import PdfReader, PdfWriter, PdfMerger
except Exception:
    from PyPDF2 import PdfReader, PdfWriter, PdfMerger


# -----------------------------
# CONFIG / PATHS
# -----------------------------
TEMPLATE_PATH = r'./templates_and_suggested_fonts/Template2/template.pdf'
EXCEL_FILE = 'comp.xlsx'
OUTPUT_FILE = 'newcertis.pdf'
FONT_PATH = r'./Hammersmith_One16/HammersmithOne-Regular.ttf'

SVG_DIR = r'./events'
DEDUP_BY_NAME = False

# -----------------------------
# ICON LAYOUT CONFIG
# -----------------------------
# Dedicated top strip where icons should appear.
# Tune these values only.
ICON_AREA_LEFT = 12
ICON_AREA_RIGHT = 12
ICON_ROW_Y = 780

# Global icon sizing controls
MAX_ICON_SIZE_PT = 42.0
MIN_ICON_SIZE_PT = 24.0

# Gap rules
PREFERRED_MAX_GAP_PT = 22.0
PREFERRED_MIN_GAP_PT = 4.0

# Extra breathing room inside each icon slot
# Lower value -> more empty space around the icon
DENSE_INNER_PAD_RATIO = 0.76
NORMAL_INNER_PAD_RATIO = 0.78
LOOSE_INNER_PAD_RATIO = 0.82

# If True, only draw icons for participated events (flag == 1)
# If False, draw all icons and gray out the inactive ones
DRAW_ONLY_ACTIVE_EVENTS = False


# -----------------------------
# READ TEMPLATE + PAGE SIZE + FONT
# -----------------------------
if not os.path.exists(TEMPLATE_PATH):
    raise FileNotFoundError(f"Template PDF not found at: {TEMPLATE_PATH}")

with open(TEMPLATE_PATH, 'rb') as _tf:
    TEMPLATE_BYTES = _tf.read()

tmp_reader_for_size = PdfReader(io.BytesIO(TEMPLATE_BYTES))
base_page = tmp_reader_for_size.pages[0]
page_width = float(base_page.mediabox.width)
page_height = float(base_page.mediabox.height)

if not os.path.exists(FONT_PATH):
    raise FileNotFoundError(f"Font not found at: {FONT_PATH}")

pdfmetrics.registerFont(TTFont('HammersmithOne', FONT_PATH))


# -----------------------------
# HELPERS
# -----------------------------
def normalize_key(s: str) -> str:
    """
    Normalize Excel headers / filenames so matching is robust.
    Example:
    '333 OH' -> '333oh'
    '333_oh' -> '333oh'
    '333-oh' -> '333oh'
    """
    if s is None:
        return ""
    s = str(s).strip().lower()
    for ch in [" ", "_", "-", "."]:
        s = s.replace(ch, "")
    return s


def build_svg_lookup(svg_dir):
    """
    Scan SVG_DIR and build:
    {
        '333': '/full/path/333.svg',
        '222': '/full/path/222.svg',
        '333oh': '/full/path/333oh.svg',
        ...
    }
    """
    if not os.path.isdir(svg_dir):
        raise FileNotFoundError(f"SVG folder not found: {svg_dir}")

    svg_map = {}
    for root, _, files in os.walk(svg_dir):
        for file in files:
            if file.lower().endswith(".svg"):
                stem = os.path.splitext(file)[0]
                key = normalize_key(stem)
                full_path = os.path.join(root, file)
                svg_map[key] = full_path
    return svg_map


def find_event_columns_and_svgs(df, svg_dir):
    """
    Looks at all columns after 'Name' and keeps only those whose header
    matches an available SVG file.
    """
    cols = [str(c).strip() for c in df.columns]
    df.columns = cols

    if 'Name' not in df.columns:
        for alt in ['name', 'Full Name', 'Participant', 'Participant Name']:
            if alt in df.columns:
                df.rename(columns={alt: 'Name'}, inplace=True)
                break

    if 'Name' not in df.columns:
        raise KeyError("Could not find a 'Name' column in the Excel file.")

    name_idx = list(df.columns).index('Name')
    candidate_cols = df.columns[name_idx + 1:]

    svg_lookup = build_svg_lookup(svg_dir)

    event_cols = []
    svg_paths = []
    missing_headers = []

    for col in candidate_cols:
        col_clean = str(col).strip()
        if not col_clean:
            continue

        key = normalize_key(col_clean)

        if key in svg_lookup:
            event_cols.append(col)
            svg_paths.append(svg_lookup[key])
        else:
            missing_headers.append(col)

    if not event_cols:
        raise RuntimeError(
            "No event columns were matched with SVG files.\n"
            f"Check your Excel headers and SVG names in: {svg_dir}"
        )

    if missing_headers:
        print(
            "[INFO] Ignored columns after 'Name' because no matching SVG was found:",
            ", ".join(map(str, missing_headers))
        )

    return event_cols, svg_paths


def choose_inner_pad_ratio(n):
    if n >= 15:
        return DENSE_INNER_PAD_RATIO
    if n >= 10:
        return NORMAL_INNER_PAD_RATIO
    return LOOSE_INNER_PAD_RATIO


def get_svg_layout_positions(n):
    """
    Create one clean centered row across the certificate width.

    Returns:
        positions: [(cx, cy), ...]
        icon_size_pt
        inner_pad_ratio
        gap_x
    """
    if n <= 0:
        return [], MAX_ICON_SIZE_PT, NORMAL_INNER_PAD_RATIO, 0.0

    usable_width = page_width - ICON_AREA_LEFT - ICON_AREA_RIGHT
    inner_pad_ratio = choose_inner_pad_ratio(n)

    if n == 1:
        return [(round(page_width / 2.0, 2), round(ICON_ROW_Y, 2))], MAX_ICON_SIZE_PT, inner_pad_ratio, 0.0

    # Start with max icon size and compute natural gap
    icon_size = MAX_ICON_SIZE_PT
    gap_x = (usable_width - (n * icon_size)) / (n - 1)

    # Clamp gap to preferred range
    if gap_x > PREFERRED_MAX_GAP_PT:
        gap_x = PREFERRED_MAX_GAP_PT
    if gap_x < PREFERRED_MIN_GAP_PT:
        gap_x = PREFERRED_MIN_GAP_PT
        icon_size = (usable_width - ((n - 1) * gap_x)) / n

    # Final safety lower bound
    if icon_size < MIN_ICON_SIZE_PT:
        icon_size = MIN_ICON_SIZE_PT
        gap_x = (usable_width - (n * icon_size)) / (n - 1)
        if gap_x < 0:
            gap_x = 0.0
            icon_size = usable_width / n

    total_width = (n * icon_size) + ((n - 1) * gap_x)
    start_x = (page_width - total_width) / 2.0 + (icon_size / 2.0)

    positions = []
    for i in range(n):
        cx = start_x + i * (icon_size + gap_x)
        positions.append((round(cx, 2), round(ICON_ROW_Y, 2)))

    return positions, float(icon_size), float(inner_pad_ratio), float(gap_x)


# -----------------------------
# COLOR UTILITY
# -----------------------------
def change_color_to_gray(drawing):
    light_gray = Color(0.7294117647, 0.6862745098, 0.6274509804)

    def _recurse(el):
        if hasattr(el, 'fillColor') and getattr(el, 'fillColor'):
            el.fillColor = light_gray
        if hasattr(el, 'strokeColor') and getattr(el, 'strokeColor'):
            el.strokeColor = light_gray
        if isinstance(el, (Path, String)):
            el.fillColor = light_gray
            try:
                el.strokeColor = light_gray
            except Exception:
                pass
        if hasattr(el, 'contents') and el.contents:
            for child in el.contents:
                _recurse(child)

    _recurse(drawing)


# -----------------------------
# SVG DRAWER - FIXED VERSION
# -----------------------------
def _maybe_gray_png(png_bytes: bytes) -> bytes:
    """
    Convert a PNG (bytes) to a light-gray version, if Pillow is available.
    If Pillow isn't installed, return original bytes.
    """
    if Image is None:
        return png_bytes
    try:
        im = Image.open(io.BytesIO(png_bytes)).convert("RGBA")
        # Convert to grayscale then colorize to a light gray similar to change_color_to_gray()
        gray = ImageOps.grayscale(im)
        light = ImageOps.colorize(gray, black="#BAB0A0", white="#BAB0A0").convert("RGBA")
        # Preserve alpha from original
        if im.mode == "RGBA":
            light.putalpha(im.split()[-1])
        out = io.BytesIO()
        light.save(out, format="PNG")
        return out.getvalue()
    except Exception:
        return png_bytes


def _render_svg_to_png_bytes(svg_path: str, px: int, gray_out: bool) -> bytes | None:
    """
    Render SVG to PNG bytes at roughly px x px.
    Tries (1) CairoSVG if available, else (2) resvg_py, else returns None.
    """
    # 1) CairoSVG (can be unavailable on Windows if cairo dll missing)
    if cairosvg is not None:
        try:
            png = cairosvg.svg2png(url=svg_path, output_width=px, output_height=px)
            if gray_out:
                png = _maybe_gray_png(png)
            return png
        except Exception:
            pass

    # 2) resvg_py (Rust renderer; ships wheels for Windows)
    if resvg_py is not None:
        try:
            svg_text = open(svg_path, "r", encoding="utf-8").read()
            # resvg_py.svg_to_bytes returns PNG by default
            png = resvg_py.svg_to_bytes(svg_text, width=int(px), height=int(px))
            if gray_out:
                png = _maybe_gray_png(png)
            return png
        except Exception:
            pass

    # Warn once so it's obvious we're falling back to svglib
    if not getattr(_render_svg_to_png_bytes, "_warned_no_svg_renderer", False):
        print(
            "[WARN] No robust SVG renderer available (cairosvg/resvg_py). Falling back to svglib; some icons may distort.",
            file=sys.stderr,
        )
        setattr(_render_svg_to_png_bytes, "_warned_no_svg_renderer", True)
    return None


def draw_svg_centered(can, svg_path, center_x, center_y, box_w, box_h, inner_pad_ratio=0.78, gray_out=False):
    """
    Draw SVG centered in a box with safer normalization.
    This avoids the visual distortion caused by odd SVG bounds/viewBox offsets.
    """
    if not os.path.exists(svg_path):
        return

    # --- Preferred path: CairoSVG -> PNG -> drawImage
    # svglib doesn't fully support SVG fill/clip rules (evenodd, clip-rule) and can distort icons.
    target_w_pt = float(box_w) * float(inner_pad_ratio)
    target_h_pt = float(box_h) * float(inner_pad_ratio)

    # Render at high DPI to keep tiny icons crisp.
    # (points -> inches -> pixels)
    # At ~30pt icons, 300 DPI can lose small cutouts/details; 1200 DPI stays sharp.
    dpi = 1200
    px = int(max(1.0, max(target_w_pt, target_h_pt) * dpi / 72.0))
    png_bytes = _render_svg_to_png_bytes(svg_path, px=px, gray_out=gray_out)
    if png_bytes:
        try:
            img = ImageReader(io.BytesIO(png_bytes))
            img_w_px, img_h_px = img.getSize()
            # Convert image pixel size to points for our chosen DPI
            img_w_pt = float(img_w_px) * 72.0 / dpi
            img_h_pt = float(img_h_px) * 72.0 / dpi

            # Fit image into the target box while preserving aspect ratio
            scale = min(target_w_pt / img_w_pt, target_h_pt / img_h_pt)
            draw_w = img_w_pt * scale
            draw_h = img_h_pt * scale

            render_x = float(center_x) - (draw_w / 2.0)
            render_y = float(center_y) - (draw_h / 2.0)

            can.saveState()
            can.drawImage(img, render_x, render_y, width=draw_w, height=draw_h, mask="auto")
            can.restoreState()
            return
        except Exception:
            # If image path fails for any reason, fall back to svglib below
            pass

    drawing = svg2rlg(svg_path)
    if drawing is None:
        return

    drawing = deepcopy(drawing)

    if gray_out:
        change_color_to_gray(drawing)

    x0, y0, x1, y1 = drawing.getBounds()
    content_w = float(x1 - x0)
    content_h = float(y1 - y0)

    if content_w <= 0 or content_h <= 0:
        return

    scale = min(
        (box_w * inner_pad_ratio) / content_w,
        (box_h * inner_pad_ratio) / content_h
    )

    final_w = content_w * scale
    final_h = content_h * scale

    render_x = center_x - (final_w / 2.0)
    render_y = center_y - (final_h / 2.0)

    can.saveState()
    can.translate(render_x, render_y)
    can.scale(scale, scale)
    can.translate(-x0, -y0)
    renderPDF.draw(drawing, can, 0, 0)
    can.restoreState()


# -----------------------------
# OVERLAY BUILDER
# -----------------------------
def draw_overlay_pdf(name, transparency_values, svg_paths, svg_positions, icon_size_pt, inner_pad_ratio, page_w, page_h):
    name = ("" if pd.isna(name) else str(name)).strip().upper()

    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=(page_w, page_h))

    # Fit name width
    max_width = page_w - 200
    font_size = 45.3
    while can.stringWidth(name, "HammersmithOne", font_size) > max_width and font_size > 10:
        font_size -= 1.0

    can.setFont("HammersmithOne", font_size)

    fill_color = HexColor("#000000")
    can.setFillColor(fill_color)
    can.setStrokeColor(fill_color)

    # Draw participant name
    char_spacing = 6
    text_width = can.stringWidth(name, "HammersmithOne", font_size) + max(0, len(name) - 1) * char_spacing
    x = (page_w - text_width) / 2.0
    y = 350

    can.saveState()
    can.setLineWidth(1.2)

    txt = can.beginText()
    txt.setTextOrigin(x, y)
    txt.setFont("HammersmithOne", font_size)

    try:
        txt.setCharSpace(char_spacing)
    except Exception:
        pass

    try:
        txt.setTextRenderMode(2)
    except Exception:
        try:
            can._code.append("2 Tr")
        except Exception:
            pass

    txt.textLine(name)
    can.drawText(txt)
    can.restoreState()

    # Event SVGs
    for idx, (svg_path, (svg_cx, svg_cy), transparency) in enumerate(
        zip(svg_paths, svg_positions, transparency_values)
    ):
        try:
            try:
                tval = int(float(transparency)) if (transparency is not None and str(transparency) != "nan") else 0
            except Exception:
                tval = 0

            if DRAW_ONLY_ACTIVE_EVENTS and tval == 0:
                continue

            draw_svg_centered(
                can=can,
                svg_path=svg_path,
                center_x=svg_cx,
                center_y=svg_cy,
                box_w=icon_size_pt,
                box_h=icon_size_pt,
                inner_pad_ratio=inner_pad_ratio,
                gray_out=(False if DRAW_ONLY_ACTIVE_EVENTS else (tval == 0))
            )

        except Exception as e:
            print(
                f"[WARN] SVG render failed (row name={name}, idx={idx}, path={svg_path}): {e}",
                file=sys.stderr
            )
            continue

    can.save()
    packet.seek(0)
    return packet


# -----------------------------
# PAGE CREATOR
# -----------------------------
def create_certificate_page_bytes(name, transparency_values, svg_paths, svg_positions, icon_size_pt, inner_pad_ratio):
    overlay_stream = draw_overlay_pdf(
        name,
        transparency_values,
        svg_paths,
        svg_positions,
        icon_size_pt,
        inner_pad_ratio,
        page_width,
        page_height
    )
    overlay_reader = PdfReader(overlay_stream)

    base_reader = PdfReader(io.BytesIO(TEMPLATE_BYTES))
    page = base_reader.pages[0]
    page.merge_page(overlay_reader.pages[0])

    out_buf = io.BytesIO()
    one = PdfWriter()
    one.add_page(page)
    one.write(out_buf)
    out_buf.seek(0)
    return out_buf.getvalue()


# -----------------------------
# MAIN
# -----------------------------
def main():
    if not os.path.exists(EXCEL_FILE):
        raise FileNotFoundError(f"Excel file not found at: {EXCEL_FILE}")

    try:
        df = pd.read_excel(EXCEL_FILE)
    except Exception as e:
        raise RuntimeError(
            f"Failed to read '{EXCEL_FILE}'. If it's .xlsx, ensure 'openpyxl' is installed. Original error: {e}"
        )

    # Normalize headers
    df.columns = [str(c).strip() for c in df.columns]

    if 'Name' not in df.columns:
        for alt in ['name', 'Full Name', 'Participant', 'Participant Name']:
            if alt in df.columns:
                df.rename(columns={alt: 'Name'}, inplace=True)
                break

    if 'Name' not in df.columns:
        raise KeyError("Could not find a 'Name' column in the Excel file.")

    # Clean names
    df['Name'] = df['Name'].astype(object)
    df['Name'] = df['Name'].ffill()
    df['Name'] = df['Name'].apply(lambda x: "" if pd.isna(x) else str(x).strip())

    # Drop fully blank rows except Name
    if df.shape[1] > 1:
        non_name_cols = [c for c in df.columns if c != 'Name']

        def _row_all_blank(s):
            return all((pd.isna(v) or str(v).strip() == "") for v in s)

        mask_keep = ~df[non_name_cols].apply(_row_all_blank, axis=1)
        df = df[mask_keep].copy()

    if DEDUP_BY_NAME:
        before = len(df)
        df = df.drop_duplicates(subset=['Name']).reset_index(drop=True)
        after = len(df)
        print(f"[INFO] Dedup by name: {before} -> {after}")

    # Detect event columns + SVGs
    event_cols, svg_paths = find_event_columns_and_svgs(df, SVG_DIR)
    event_count = len(svg_paths)

    svg_positions, icon_size_pt, inner_pad_ratio, gap_x = get_svg_layout_positions(event_count)

    print("[INFO] Detected event columns:", event_cols)
    print("[INFO] Matching SVG files:")
    for col, path in zip(event_cols, svg_paths):
        print(f"   {col} -> {path}")

    print(f"[INFO] Total events: {event_count}")
    print(f"[INFO] Page width: {round(page_width, 2)}")
    print(f"[INFO] Icon area: left={ICON_AREA_LEFT}, right={ICON_AREA_RIGHT}, row_y={ICON_ROW_Y}")
    print(f"[INFO] Computed icon size: {round(icon_size_pt, 2)} pt")
    print(f"[INFO] Computed gap_x: {round(gap_x, 2)} pt")
    print(f"[INFO] Computed inner_pad_ratio: {inner_pad_ratio}")

    # Convert event flags to 0/1
    def _to_flag(x):
        if pd.isna(x):
            return 0
        s = str(x).strip().lower()
        if s in {"1", "true", "yes", "y", "✓", "check", "x"}:
            return 1
        try:
            return 1 if int(float(s)) != 0 else 0
        except Exception:
            return 0

    df_events = df[event_cols].copy()
    for col in event_cols:
        df_events[col] = df_events[col].map(_to_flag)

    total = len(df)
    print(f"[INFO] Preparing {total} certificate(s)...")

    merger = PdfMerger()
    rendered_count = 0

    for idx in range(total):
        row = df.iloc[idx]
        safe_name = "" if pd.isna(row['Name']) else str(row['Name']).strip()
        if not safe_name:
            safe_name = "PARTICIPANT"

        flags = df_events.iloc[idx].tolist()

        try:
            page_bytes = create_certificate_page_bytes(
                safe_name,
                flags,
                svg_paths,
                svg_positions,
                icon_size_pt,
                inner_pad_ratio
            )
            merger.append(io.BytesIO(page_bytes))
            rendered_count += 1
            print(f"  - [{idx + 1}/{total}] {safe_name}")
        except Exception as e:
            print(f"[ERROR] Failed for row {idx + 1} (name={safe_name}): {e}", file=sys.stderr)
            continue

    if rendered_count == 0:
        merger.close()
        raise RuntimeError("No pages were rendered. Check your Excel and SVG paths.")

    with open(OUTPUT_FILE, "wb") as out:
        merger.write(out)
    merger.close()

    print(f"Combined certificates created successfully! Pages written: {rendered_count}/{total} -> {OUTPUT_FILE}")


if __name__ == "__main__":
    main()