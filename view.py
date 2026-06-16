from __future__ import annotations

import base64
import json
import mimetypes
import re
import sqlite3
from datetime import datetime
from html import escape
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import Image as RLImage
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

APP_DIR = Path(__file__).resolve().parent
DB_PATH = APP_DIR / "tong_ket.db"
UPLOAD_DIR = APP_DIR / "uploads"

DEFAULT_PAYMENT_INFO = """Payment methods: Bank transfer
Bank name: Asia Commercial Joint Stock Bank (ACB - A Chau Bank)
Account number: 196653719
Full name: LUONG NHAT TU
Swift code: ASCBVNVX"""

st.set_page_config(
    page_title="Tong Ket Manager - View",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ============================================================
# GLOBAL CSS + MODAL (giống app.py)
# ============================================================
st.markdown("""
<style>
.period-table { width: 100%; border-collapse: collapse; font-size: 14px; }
.period-table th {
    background: #2D2D2D; color: white; padding: 10px;
    text-align: left; position: sticky; top: 0;
}
.period-table td {
    border-bottom: 1px solid #E8E8E4; padding: 10px;
    vertical-align: top; line-height: 1.4;
}
.period-table .num { text-align: right; white-space: nowrap; }
.period-table .amount { font-weight: 700; color: #B8760A; }
.period-table .owner { color: #6B7280; font-size: 12px; }
.period-table .desc { color: #6B7280; font-size: 12px; margin-top: 5px; line-height: 1.35; }
.period-table .total td { background: #F7F7F5; font-weight: 700; }
.image-grid { display: grid; grid-template-columns: repeat(3, 64px); gap: 5px; }
.image-grid img {
    width: 64px; height: 64px; object-fit: contain;
    border: 1px solid #E8E8E4; border-radius: 6px; background: #F7F7F5;
    cursor: zoom-in;
}
.image-grid img:hover {
    transform: scale(1.05);
    box-shadow: 0 0 0 2px #B8760A;
}
.gallery-card {
    border: 1px solid #E8E8E4;
    border-radius: 10px;
    overflow: hidden;
    background: white;
    margin-bottom: 12px;
}
.gallery-img-wrap {
    width: 100%;
    aspect-ratio: 1;
    background: #F7F7F5;
    overflow: hidden;
    display: flex;
    align-items: center;
    justify-content: center;
    cursor: zoom-in;
}
.gallery-img-wrap img {
    width: 100%;
    height: 100%;
    object-fit: contain;
}
.gallery-caption {
    padding: 8px 10px;
    font-size: 12px;
    line-height: 1.4;
    color: #1A1A1A;
}
.image-modal {
    position: fixed; inset: 0; z-index: 999999;
    display: none; align-items: center; justify-content: center;
    padding: 24px; background: rgba(0,0,0,0.92);
    backdrop-filter: blur(8px);
    cursor: zoom-out;
}
.image-modal.is-open { display: flex; }
.image-modal img {
    max-width: min(96vw, 1400px);
    max-height: 92vh;
    object-fit: contain;
    border-radius: 8px;
    box-shadow: 0 20px 80px rgba(0,0,0,0.6);
}
.image-modal-close {
    position: fixed; top: 20px; right: 20px;
    width: 48px; height: 48px;
    border: 0; border-radius: 50%;
    background: rgba(255,255,255,0.2);
    color: white; font-size: 30px;
    cursor: pointer;
    display: flex;
    align-items: center;
    justify-content: center;
    z-index: 1000000;
}
</style>

<div id="globalImageModal" class="image-modal" aria-hidden="true">
    <button type="button" class="image-modal-close" onclick="closeModal()">×</button>
    <img id="globalImageModalImg" src="" alt="full size" />
</div>

<script>
(function() {
    var modal = document.getElementById('globalImageModal');
    var modalImg = document.getElementById('globalImageModalImg');

    window.openModal = function(src) {
        if (!modal || !modalImg) return;
        modalImg.src = src;
        modal.classList.add('is-open');
        modal.setAttribute('aria-hidden', 'false');
    };

    window.closeModal = function() {
        if (!modal) return;
        modal.classList.remove('is-open');
        modal.setAttribute('aria-hidden', 'true');
        if (modalImg) setTimeout(function(){ modalImg.src = ''; }, 300);
    };

    // Click background to close
    if (modal) {
        modal.addEventListener('click', function(e) {
            if (e.target === modal) closeModal();
        });
    }

    // ESC
    document.addEventListener('keydown', function(e) {
        if (e.key === 'Escape') closeModal();
    });
})();
</script>
""", unsafe_allow_html=True)


# ============================================================
# HELPERS
# ============================================================
def normalize_text(value: object) -> str:
    if value is None:
        return ""
    return str(value).replace("\xa0", " ").strip()


def uppercase_label(value: object) -> str:
    return normalize_text(value).upper()


def number_or_zero(value: object) -> float:
    try:
        if value is None or value == "":
            return 0.0
        return float(value)
    except (TypeError, ValueError):
        return 0.0


MONTH_ORDER = {
    "JANUARY": 1, "FEBRUARY": 2, "MARCH": 3, "APRIL": 4,
    "MAY": 5, "JUNE": 6, "JULY": 7, "AUGUST": 8,
    "SEPTEMBER": 9, "OCTOBER": 10, "NOVEMBER": 11, "DECEMBER": 12,
}


def parse_year(text: str) -> int | None:
    match = re.search(r"(20\d{2})", text or "")
    return int(match.group(1)) if match else None


def period_sort_key(period_label: str) -> tuple:
    text = normalize_text(period_label).upper()
    year = parse_year(text) or 0
    months = [m for name, m in MONTH_ORDER.items() if name in text]
    month = max(months) if months else 0
    return (year, month, text)


def image_paths_from_value(value: object) -> list[str]:
    raw = normalize_text(value)
    if not raw:
        return []
    return [part.strip() for part in raw.split("|") if part.strip()]


# ============================================================
# DATABASE READ-ONLY
# ============================================================
@st.cache_data(show_spinner=False)
def load_entries() -> pd.DataFrame:
    if not DB_PATH.exists():
        st.error("Database not found. Please sync database to GitHub first.")
        return pd.DataFrame()
    try:
        conn = sqlite3.connect(f"file:{DB_PATH.as_posix()}?mode=ro", uri=True)
        data = pd.read_sql_query("SELECT * FROM entries WHERE deleted_at IS NULL", conn)
        conn.close()
    except Exception as exc:
        st.error(f"Error reading database: {exc}")
        return pd.DataFrame()

    for column in ["client_name", "project_name", "owner"]:
        if column in data.columns:
            data[column] = data[column].fillna("").map(uppercase_label)
    if "period_label" in data.columns:
        data["period_label"] = data["period_label"].fillna("").map(normalize_text)
    if "period_year" in data.columns:
        data["period_year"] = pd.to_numeric(data["period_year"], errors="coerce")
    for column in ["drawing_qty", "unit_price", "amount"]:
        if column in data.columns:
            data[column] = pd.to_numeric(data[column], errors="coerce").fillna(0)
    return data


@st.cache_data(show_spinner=False)
def load_last_sync() -> str:
    if not DB_PATH.exists():
        return ""
    try:
        conn = sqlite3.connect(f"file:{DB_PATH.as_posix()}?mode=ro", uri=True)
        row = conn.execute(
            "SELECT value FROM settings WHERE key = 'last_import_summary'"
        ).fetchone()
        conn.close()
        if row and row[0]:
            summary = json.loads(row[0])
            return str(summary.get("imported_at", ""))
    except Exception:
        pass
    return ""


@st.cache_data(show_spinner=False)
def load_payment_info() -> str:
    if not DB_PATH.exists():
        return DEFAULT_PAYMENT_INFO
    try:
        conn = sqlite3.connect(f"file:{DB_PATH.as_posix()}?mode=ro", uri=True)
        row = conn.execute(
            "SELECT value FROM settings WHERE key = 'payment_info'"
        ).fetchone()
        conn.close()
        if row and row[0]:
            return str(row[0])
    except Exception:
        pass
    return DEFAULT_PAYMENT_INFO


# ============================================================
# IMAGE RESOLUTION
# ============================================================
@st.cache_data(show_spinner=False)
def _build_upload_filename_index() -> dict[str, list[Path]]:
    index: dict[str, list[Path]] = {}
    if not UPLOAD_DIR.exists():
        return index
    for p in UPLOAD_DIR.rglob("*"):
        if p.is_file():
            index.setdefault(p.name, []).append(p)
    return index


def resolve_image_path(value: object) -> Path | None:
    path_text = normalize_text(value)
    if not path_text:
        return None
    normalized = path_text.replace("\\", "/")
    path = Path(normalized)
    filename = path.name
    candidates: list[Path] = []

    is_windows_style = (
        len(path.parts) > 0
        and len(path.parts[0]) >= 2
        and path.parts[0][-1] == ":"
    )
    if not path.is_absolute() and not is_windows_style:
        candidates.append(APP_DIR / path)
        candidates.append(UPLOAD_DIR / path)
        candidates.append(UPLOAD_DIR / filename)

    try:
        parts_lower = [p.lower() for p in path.parts]
        if "uploads" in parts_lower:
            idx = parts_lower.index("uploads")
            rel_parts = path.parts[idx + 1:]
            if rel_parts:
                candidates.append(UPLOAD_DIR.joinpath(*rel_parts))
    except (ValueError, IndexError):
        pass

    if filename:
        for match in _build_upload_filename_index().get(filename, []):
            candidates.append(Path(match))

    seen: set[str] = set()
    for candidate in candidates:
        try:
            key = str(candidate.resolve())
        except Exception:
            key = str(candidate)
        if key in seen:
            continue
        seen.add(key)
        if candidate.exists() and candidate.is_file():
            return candidate
    return None


def image_data_uri(path: Path) -> str:
    mime_type = mimetypes.guess_type(str(path))[0] or "image/jpeg"
    encoded = base64.b64encode(path.read_bytes()).decode("utf-8")
    return f"data:{mime_type};base64,{encoded}"


# ============================================================
# SEARCH AUTOCOMPLETE
# ============================================================
def search_autocomplete_options(data: pd.DataFrame, query: str, limit: int = 20) -> list[tuple[str, str]]:
    query_text = normalize_text(query).lower()
    if data.empty or not query_text:
        return []
    buckets = [
        ("Dự án", "project_name"),
        ("Công ty", "client_name"),
        ("Người phụ trách", "owner"),
        ("Kỳ", "period_label"),
    ]
    matches: list[tuple[str, str]] = []
    seen: set[tuple[str, str]] = set()
    for label, column in buckets:
        if column not in data.columns:
            continue
        values = data[column].dropna().unique().tolist()
        for val in values:
            v = str(val).strip()
            if not v or query_text not in v.lower():
                continue
            key = (label, v.lower())
            if key in seen:
                continue
            seen.add(key)
            matches.append((f"{label} · {v}", v))
            if len(matches) >= limit:
                return matches
    return matches


# ============================================================
# PDF GENERATION (sửa lỗi escape)
# ============================================================
def calc_pdf_row_height(image_value: object, text_lines: list[str]) -> float:
    image_count = min(len(image_paths_from_value(image_value)), 6)
    image_rows = (image_count + 2) // 3 if image_count else 0
    height_from_images = (image_rows * 22 * mm) + (6 * mm) if image_rows else 0

    wrapped = 0
    for line in text_lines:
        text = normalize_text(line)
        wrapped += max(1, (len(text) // 32) + 1) if text else 1
    height_from_text = max(1, wrapped) * 4.5 * mm + (6 * mm)
    return max(20 * mm, height_from_images, height_from_text)


def pdf_image_grid(value: object, max_images: int = 6):
    images = []
    for path_text in image_paths_from_value(value)[:max_images]:
        path = resolve_image_path(path_text)
        if not path:
            continue
        try:
            img = RLImage(str(path))
            img._restrictSize(20 * mm, 20 * mm)
            images.append(img)
        except Exception:
            continue
    if not images:
        return ""
    rows = []
    for i in range(0, len(images), 3):
        row = images[i:i+3]
        while len(row) < 3:
            row.append("")
        rows.append(row)
    table = Table(rows, colWidths=[22 * mm, 22 * mm, 22 * mm], rowHeights=[22 * mm] * len(rows))
    table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("LEFTPADDING", (0, 0), (-1, -1), 1),
        ("RIGHTPADDING", (0, 0), (-1, -1), 1),
        ("TOPPADDING", (0, 0), (-1, -1), 1),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
    ]))
    return table


def make_pdf_bytes(data: pd.DataFrame, title: str, include_payment: bool = True) -> bytes:
    if data.empty:
        buffer = BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4)
        styles = getSampleStyleSheet()
        story = [
            Paragraph("No data to display", styles["Title"]),
            Spacer(1, 12 * mm),
            Paragraph(f"Requested title: {escape(title)}", styles["Normal"]),
        ]
        doc.build(story)
        return buffer.getvalue()

    try:
        buffer = BytesIO()
        doc = SimpleDocTemplate(
            buffer,
            pagesize=A4,
            rightMargin=15 * mm,
            leftMargin=15 * mm,
            topMargin=15 * mm,
            bottomMargin=15 * mm,
        )
        styles = getSampleStyleSheet()
        title_style = ParagraphStyle(
            "Title", parent=styles["Heading1"],
            fontSize=18, leading=22, alignment=1, spaceAfter=4
        )
        subtitle_style = ParagraphStyle(
            "Subtitle", parent=styles["Normal"],
            fontSize=9, alignment=1, textColor=colors.HexColor("#6B7280"), spaceAfter=12
        )
        body_style = ParagraphStyle("Body", parent=styles["BodyText"], fontSize=9, leading=12)
        project_style = ParagraphStyle("Project", parent=body_style, fontSize=9, leading=12)
        period_label_style = ParagraphStyle(
            "PeriodLabel", parent=body_style,
            fontSize=9, leading=11, textColor=colors.HexColor("#374151")
        )
        payment_style = ParagraphStyle("Payment", parent=body_style, fontSize=9, leading=13)

        story = []
        story.append(Paragraph(escape(title), title_style))
        total_qty = data["drawing_qty"].sum() if not data.empty else 0
        total_amount = data["amount"].sum() if not data.empty else 0
        story.append(Paragraph(
            f"Total drawings: <b>{total_qty:g}</b> &nbsp;|&nbsp; Total amount: <b>SGD {total_amount:,.2f}</b>",
            subtitle_style,
        ))

        sorted_data = data.sort_values(
            ["period_year", "period_label", "source_file", "source_row", "id"],
            ascending=[False, False, True, True, True],
            na_position="last",
        ).reset_index(drop=True)

        table_rows: list = []
        row_heights: list = []
        last_period = None

        for idx, row in sorted_data.iterrows():
            project_name = escape(str(row.get("project_name") or ""))
            owner = escape(normalize_text(row.get("owner")))
            desc = escape(normalize_text(row.get("description")))
            period_label = escape(str(row.get("period_label") or ""))

            if period_label != last_period:
                table_rows.append([
                    Paragraph(f"<b>📅 {period_label or '-'}</b>", period_label_style),
                    "", "", "", "", "",
                ])
                row_heights.append(8 * mm)
                last_period = period_label

            project_parts = [f"<b>{project_name}</b>"]
            if owner:
                project_parts.append(f"<i>{owner}</i>")
            if desc:
                project_parts.append(f"<font size=8 color='#6B7280'>{desc.replace(chr(10), '<br/>')}</font>")
            project_html = "<br/>".join(project_parts)
            text_for_height = f"{project_name} {owner} {desc}"

            image_table = pdf_image_grid(row.get("image_path"))
            table_rows.append([
                str(idx + 1),
                Paragraph(project_html, project_style),
                f"{number_or_zero(row.get('drawing_qty')):g}",
                f"SGD {number_or_zero(row.get('unit_price')):,.0f}",
                f"SGD {number_or_zero(row.get('amount')):,.0f}",
                image_table if image_table != "" else "",
            ])
            row_heights.append(calc_pdf_row_height(row.get("image_path"), [text_for_height]))

        table_rows.append([
            "",
            Paragraph("<b>TOTAL</b>", project_style),
            f"<b>{total_qty:g}</b>",
            "",
            f"<b>SGD {total_amount:,.0f}</b>",
            "",
        ])
        row_heights.append(9 * mm)

        col_widths = [10 * mm, 60 * mm, 14 * mm, 20 * mm, 24 * mm, 52 * mm]
        table = Table(table_rows, colWidths=col_widths, rowHeights=row_heights, repeatRows=0)

        table_style = TableStyle([
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#D9D9D9")),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("ALIGN", (0, 0), (0, -1), "CENTER"),
            ("ALIGN", (2, 0), (4, -1), "RIGHT"),
            ("BACKGROUND", (4, 0), (4, -1), colors.HexColor("#FFF7ED")),
            ("TEXTCOLOR", (4, 0), (4, -1), colors.HexColor("#B45309")),
            ("FONTNAME", (4, 0), (4, -1), "Helvetica-Bold"),
            ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#E5E7EB")),
            ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
        ])

        for i, row in enumerate(table_rows):
            if row and hasattr(row[0], 'text') and '📅' in str(row[0].text):
                table_style.add("BACKGROUND", (0, i), (-1, i), colors.HexColor("#F3F4F6"))
                table_style.add("SPAN", (0, i), (-1, i))
                table_style.add("LINEABOVE", (0, i), (-1, i), 0.5, colors.HexColor("#9CA3AF"))
                table_style.add("LINEBELOW", (0, i), (-1, i), 0.25, colors.HexColor("#D9D9D9"))

        table.setStyle(table_style)
        story.append(table)

        if include_payment:
            story.append(Spacer(1, 10 * mm))
            story.append(Paragraph("<b>Payment Information</b>", styles["Heading3"]))
            story.append(Spacer(1, 3 * mm))
            payment = load_payment_info()
            for line in payment.splitlines():
                if line.strip():
                    story.append(Paragraph(escape(line), payment_style))

        doc.build(story)
        buffer.seek(0)
        return buffer.getvalue()

    except Exception as e:
        import traceback
        error_html = escape(traceback.format_exc()).replace("\n", "<br/>")
        buffer = BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4)
        styles = getSampleStyleSheet()
        story = [
            Paragraph("Error generating PDF", styles["Title"]),
            Spacer(1, 6 * mm),
            Paragraph(f"<font color='red'>{escape(str(e))}</font>", styles["Normal"]),
            Spacer(1, 6 * mm),
            Paragraph(error_html, styles["Code"]),
        ]
        doc.build(story)
        return buffer.getvalue()


# ============================================================
# RENDER PERIOD TABLE (có modal ảnh)
# ============================================================
def render_period_table(data: pd.DataFrame) -> None:
    if data.empty:
        st.info("No data to display.")
        return

    sorted_data = data.sort_values(
        ["period_label", "source_file", "source_row", "id"],
        ascending=[False, True, True, True],
        na_position="last",
    ).reset_index(drop=True)

    rows = []
    for idx, row in sorted_data.iterrows():
        owner = normalize_text(row.get("owner"))
        desc = normalize_text(row.get("description"))
        owner_html = f'<span class="owner"> · {escape(owner)}</span>' if owner else ""
        desc_html = f'<div class="desc">{escape(desc).replace(chr(10), "<br/>")}</div>' if desc else ""

        img_html_parts = []
        for path_text in image_paths_from_value(row.get("image_path"))[:8]:
            path = resolve_image_path(path_text)
            if not path:
                continue
            uri = image_data_uri(path)
            img_html_parts.append(
                f'<img class="zoomable" src="{uri}" data-full="{escape(uri, quote=True)}" onclick="window.parent.openModal(\'{escape(uri, quote=True)}\')" />'
            )
        img_html = '<div class="image-grid">' + "".join(img_html_parts) + "</div>" if img_html_parts else ""

        rows.append(
            "<tr>"
            f"<td>{idx + 1}</td>"
            f"<td><strong>{escape(str(row.get('project_name') or ''))}</strong>{owner_html}{desc_html}</td>"
            f"<td class='num'>{number_or_zero(row.get('drawing_qty')):g}</td>"
            f"<td class='num'>SGD {number_or_zero(row.get('unit_price')):,.0f}</td>"
            f"<td class='num amount'>SGD {number_or_zero(row.get('amount')):,.0f}</td>"
            f"<td>{img_html}</td>"
            "</tr>"
        )

    rows.append(
        "<tr class='total'>"
        "<td></td><td>TOTAL</td>"
        f"<td class='num'>{data['drawing_qty'].sum():g}</td><td></td>"
        f"<td class='num amount'>SGD {data['amount'].sum():,.0f}</td><td></td>"
        "</tr>"
    )

    table_html = f"""
    <table class="period-table">
        <thead><tr><th>No</th><th>Project</th><th>Qty</th><th>Unit</th><th>Amount</th><th>Image</th></tr></thead>
        <tbody>{"".join(rows)}</tbody>
    </table>
    """
    st.markdown(table_html, unsafe_allow_html=True)


# ============================================================
# GALLERY RENDER
# ============================================================
def render_gallery(data: pd.DataFrame) -> None:
    image_rows = data[data["image_path"].fillna("").str.strip() != ""].copy()
    if image_rows.empty:
        st.info("No images in the current filter.")
        return

    image_rows["_sort"] = image_rows["period_label"].fillna("").map(period_sort_key)
    image_rows = image_rows.sort_values(
        ["_sort", "client_name", "source_file", "source_row", "id"],
        ascending=[False, True, True, True, True],
        na_position="last",
    ).drop(columns=["_sort"], errors="ignore")

    cols = st.columns(4)
    item_index = 0
    for _, row in image_rows.iterrows():
        for path_text in image_paths_from_value(row.get("image_path")):
            path = resolve_image_path(path_text)
            if not path:
                continue
            uri = image_data_uri(path)
            project = str(row.get("project_name") or "")
            period = str(row.get("period_label") or "")
            short_project = project if len(project) <= 32 else project[:30] + "…"
            with cols[item_index % 4]:
                st.markdown(
                    f"""
                    <div class="gallery-card">
                        <div class="gallery-img-wrap" onclick="window.parent.openModal('{escape(uri, quote=True)}')">
                            <img src="{uri}" />
                        </div>
                        <div class="gallery-caption">
                            <strong>{escape(short_project)}</strong><br/>
                            <span style="color:#888;font-size:11px;">{escape(period)}</span>
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
            item_index += 1


# ============================================================
# FILTER
# ============================================================
def filter_entries(data: pd.DataFrame, keyword: str, client: str, year: str, period: str) -> pd.DataFrame:
    filtered = data.copy()
    if client and "client_name" in filtered.columns:
        filtered = filtered[filtered["client_name"] == client]
    if year != "All" and "period_year" in filtered.columns:
        filtered = filtered[filtered["period_year"].fillna(0).astype(int).astype(str) == year]
    if period != "All" and "period_label" in filtered.columns:
        filtered = filtered[filtered["period_label"] == period]
    if keyword and keyword.strip():
        query = keyword.strip().lower()
        cols = [col for col in ["client_name", "period_label", "project_name", "owner", "description", "notes"] if col in filtered.columns]
        mask = pd.Series(False, index=filtered.index)
        for col in cols:
            mask = mask | filtered[col].fillna("").astype(str).str.lower().str.contains(query, regex=False)
        filtered = filtered[mask]
    return filtered


# ============================================================
# MAIN
# ============================================================
def main() -> None:
    st.title("📊 Tong Ket Manager - View")

    last_sync = load_last_sync()
    if last_sync:
        try:
            sync_dt = datetime.fromisoformat(last_sync)
            sync_str = sync_dt.strftime("%Y-%m-%d %H:%M:%S")
            st.caption(f"📅 Last sync: **{sync_str}** (from local app → GitHub)")
        except Exception:
            st.caption(f"📅 Last sync: **{last_sync}**")
    else:
        st.caption("📅 No sync information available.")

    data = load_entries()
    if data.empty:
        st.warning("No data to display.")
        return

    # ============================================================
    # SIDEBAR: FILTERS (cascading + search autocomplete)
    # ============================================================
    st.sidebar.header("Filters")

    all_clients = sorted([v for v in data["client_name"].dropna().unique().tolist() if v])
    if not all_clients:
        st.warning("No clients found.")
        return

    # Tìm client gần nhất
    data_sorted = data.copy()
    data_sorted["_sort"] = data_sorted["period_label"].fillna("").map(period_sort_key)
    data_sorted = data_sorted.sort_values(
        ["_sort", "client_name", "source_file", "source_row", "id"],
        ascending=[False, True, True, True, True],
        na_position="last",
    ).drop(columns=["_sort"], errors="ignore")
    latest_client = data_sorted.iloc[0]["client_name"]

    # Khởi tạo session state
    if "filter_client" not in st.session_state or st.session_state.filter_client not in all_clients:
        st.session_state.filter_client = latest_client
    if "filter_year" not in st.session_state:
        st.session_state.filter_year = "All"
    if "filter_period" not in st.session_state:
        st.session_state.filter_period = "All"
    if "search_keyword" not in st.session_state:
        st.session_state.search_keyword = ""

    # Nút reset
    if st.sidebar.button("🔄 Xóa bộ lọc", use_container_width=True):
        st.session_state.filter_client = latest_client
        st.session_state.filter_year = "All"
        st.session_state.filter_period = "All"
        st.session_state.search_keyword = ""
        st.rerun()

    st.sidebar.markdown("---")

    # Client
    selected_client = st.sidebar.selectbox(
        "Client *",
        all_clients,
        index=all_clients.index(st.session_state.filter_client),
        key="filter_client_widget",
    )
    if selected_client != st.session_state.filter_client:
        st.session_state.filter_client = selected_client
        st.session_state.filter_year = "All"
        st.session_state.filter_period = "All"
    st.session_state.filter_client = selected_client

    scoped_by_client = data[data["client_name"] == selected_client]

    # Year
    years_avail = ["All"] + sorted(
        [str(int(v)) for v in scoped_by_client["period_year"].dropna().unique().tolist()],
        reverse=True,
    )
    if st.session_state.filter_year not in years_avail:
        st.session_state.filter_year = "All"
    selected_year = st.sidebar.selectbox(
        "Year",
        years_avail,
        index=years_avail.index(st.session_state.filter_year),
        key="filter_year_widget",
    )
    if selected_year != st.session_state.filter_year:
        st.session_state.filter_year = selected_year
        st.session_state.filter_period = "All"
    st.session_state.filter_year = selected_year

    scoped_by_year = scoped_by_client if selected_year == "All" else scoped_by_client[
        scoped_by_client["period_year"].fillna(0).astype(int).astype(str) == selected_year
    ]

    # Period
    period_list = sorted(
        [v for v in scoped_by_year["period_label"].dropna().unique().tolist() if v],
        key=period_sort_key,
        reverse=True,
    )
    periods_avail = ["All"] + period_list
    if st.session_state.filter_period not in periods_avail:
        st.session_state.filter_period = "All"
    selected_period = st.sidebar.selectbox(
        "Period",
        periods_avail,
        index=periods_avail.index(st.session_state.filter_period),
        key="filter_period_widget",
    )
    st.session_state.filter_period = selected_period

    # Search + Autocomplete
    st.sidebar.markdown("---")
    st.sidebar.subheader("🔍 Tìm kiếm")
    search_input = st.sidebar.text_input(
        "Project / description",
        value=st.session_state.search_keyword,
        key="search_input_widget",
        on_change=lambda: st.session_state.update(search_keyword=st.session_state.search_input_widget)
    )
    if search_input:
        suggestions = search_autocomplete_options(data, search_input, limit=20)
        if suggestions:
            # Tạo danh sách các giá trị gợi ý (label, value)
            options = [s[0] for s in suggestions]
            selected_label = st.sidebar.selectbox(
                "Gợi ý (chọn để áp dụng)",
                [""] + options,
                key="autocomplete_select",
                label_visibility="collapsed",
            )
            if selected_label:
                # Tìm value tương ứng
                for label, value in suggestions:
                    if label == selected_label:
                        st.session_state.search_keyword = value
                        st.session_state.search_input_widget = value
                        st.rerun()
                        break

    keyword = st.session_state.search_keyword

    # ============================================================
    # FILTER & DISPLAY
    # ============================================================
    # Nếu chưa chọn period, lấy period gần nhất
    if selected_period == "All" and selected_year == "All":
        scoped_for_latest = scoped_by_client.copy()
        scoped_for_latest["_sort"] = scoped_for_latest["period_label"].fillna("").map(period_sort_key)
        latest_period = None
        if not scoped_for_latest.empty:
            first_row = scoped_for_latest.sort_values(
                ["_sort", "source_file", "source_row", "id"],
                ascending=[False, True, True, True],
                na_position="last",
            ).iloc[0]
            latest_period = str(first_row.get("period_label") or "")
        if latest_period:
            filtered = scoped_by_client[scoped_by_client["period_label"] == latest_period].copy()
            st.info(f"📅 Showing latest period: **{latest_period}** (use sidebar to see others)")
        else:
            filtered = scoped_by_client.copy()
    else:
        filtered = filter_entries(data, keyword, selected_client, selected_year, selected_period)

    # Sắp xếp hiển thị
    filtered["_sort"] = filtered["period_label"].fillna("").map(period_sort_key)
    display_data = filtered.sort_values(
        ["_sort", "source_file", "source_row", "id"],
        ascending=[False, True, True, True],
        na_position="last",
    ).drop(columns=["_sort"], errors="ignore")

    # Metrics
    metric_cols = st.columns(4)
    metric_cols[0].metric("Rows", f"{len(filtered):,}")
    metric_cols[1].metric("Projects", f"{filtered['project_name'].nunique() if not filtered.empty else 0:,}")
    metric_cols[2].metric("Drawings", f"{filtered['drawing_qty'].sum() if not filtered.empty else 0:g}")
    metric_cols[3].metric("Total SGD", f"{filtered['amount'].sum() if not filtered.empty else 0:,.0f}")

    # ============================================================
    # TABS
    # ============================================================
    tab_table, tab_gallery, tab_pdf = st.tabs(["📋 Period / Data Table", "🖼️ Gallery", "📄 PDF"])

    with tab_table:
        if filtered.empty:
            st.info("No data matches the current filter.")
        else:
            st.markdown("---")
            render_period_table(display_data)

            # Download PDF button at bottom of table
            default_title = f"{selected_client} - {display_data.iloc[0].get('period_label', 'Project Summary')}"
            if st.button("📥 Tải PDF", type="primary", use_container_width=True, key="download_pdf_table"):
                with st.spinner("Đang tạo PDF..."):
                    pdf_bytes = make_pdf_bytes(display_data, default_title)
                    st.session_state["pdf_download_bytes"] = pdf_bytes
                    st.session_state["pdf_download_filename"] = f"{default_title}.pdf"
                st.success("PDF đã sẵn sàng!")

            if "pdf_download_bytes" in st.session_state:
                st.download_button(
                    "⬇️ Tải xuống",
                    data=st.session_state["pdf_download_bytes"],
                    file_name=st.session_state.get("pdf_download_filename", "summary.pdf"),
                    mime="application/pdf",
                    use_container_width=True,
                )
                # Xóa sau khi tải? Có thể giữ lại để tải lại nếu cần, hoặc xóa để tránh rối.
                # Tôi sẽ giữ lại để người dùng có thể tải lại.

    with tab_gallery:
        render_gallery(display_data)

    with tab_pdf:
        st.subheader("Generate & Download PDF")
        default_title = f"{selected_client} - {display_data.iloc[0].get('period_label', 'Project Summary')}" if not display_data.empty else "Project Summary"
        title = st.text_input("PDF Title", value=default_title, key="pdf_title")

        if st.button("🔄 Generate PDF", type="primary", use_container_width=True, disabled=display_data.empty):
            with st.spinner("Đang tạo PDF..."):
                pdf_bytes = make_pdf_bytes(display_data, title)
                st.session_state["pdf_download_bytes"] = pdf_bytes
                st.session_state["pdf_download_filename"] = f"{title}.pdf"
            st.success("PDF generated successfully!")

        if "pdf_download_bytes" in st.session_state:
            st.download_button(
                "⬇️ Download PDF",
                data=st.session_state["pdf_download_bytes"],
                file_name=st.session_state.get("pdf_download_filename", f"{title}.pdf"),
                mime="application/pdf",
                use_container_width=True,
            )


if __name__ == "__main__":
    main()