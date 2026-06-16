from __future__ import annotations

import base64
import mimetypes
import sqlite3
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


st.set_page_config(
    page_title="Tong Ket Manager - View",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)


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


@st.cache_data(show_spinner=False)
def load_entries() -> pd.DataFrame:
    if not DB_PATH.exists():
        st.error("Không tìm thấy `tong_ket.db`. Hãy đồng bộ database lên GitHub trước.")
        return pd.DataFrame()

    try:
        conn = sqlite3.connect(f"file:{DB_PATH.as_posix()}?mode=ro", uri=True)
        data = pd.read_sql_query("SELECT * FROM entries WHERE deleted_at IS NULL", conn)
        conn.close()
    except Exception as exc:
        st.error(f"Lỗi khi đọc database: {exc}")
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


def image_paths_from_value(value: object) -> list[str]:
    raw = normalize_text(value)
    if not raw:
        return []
    return [part.strip() for part in raw.split("|") if part.strip()]


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

    # relative path
    is_windows_style = (
        len(path.parts) > 0
        and len(path.parts[0]) >= 2
        and path.parts[0][-1] == ":"
    )
    if not path.is_absolute() and not is_windows_style:
        candidates.append(APP_DIR / path)
        candidates.append(UPLOAD_DIR / path)
        candidates.append(UPLOAD_DIR / filename)

    # absolute path (Windows hoặc Linux) -> lấy phần sau "uploads"
    try:
        parts_lower = [p.lower() for p in path.parts]
        if "uploads" in parts_lower:
            idx = parts_lower.index("uploads")
            rel_parts = path.parts[idx + 1:]
            if rel_parts:
                candidates.append(UPLOAD_DIR.joinpath(*rel_parts))
    except (ValueError, IndexError):
        pass

    # fallback theo tên file
    if filename:
        for match in _build_upload_filename_index().get(filename, []):
            candidates.append(match)

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


def render_images(value: object, max_images: int = 8) -> str:
    cells = []
    for path_text in image_paths_from_value(value)[:max_images]:
        path = resolve_image_path(path_text)
        if not path:
            continue
        uri = image_data_uri(path)
        cells.append(f'<img class="thumb" src="{uri}" alt="image" />')
    if not cells:
        return ""
    return '<div class="image-grid">' + "".join(cells) + "</div>"


def filter_entries(data: pd.DataFrame, keyword: str, client: str, year: str, period: str) -> pd.DataFrame:
    filtered = data.copy()
    if client != "Tất cả" and "client_name" in filtered.columns:
        filtered = filtered[filtered["client_name"] == client]
    if year != "Tất cả" and "period_year" in filtered.columns:
        filtered = filtered[filtered["period_year"].fillna(0).astype(int).astype(str) == year]
    if period != "Tất cả" and "period_label" in filtered.columns:
        filtered = filtered[filtered["period_label"] == period]
    if keyword.strip():
        query = keyword.strip().lower()
        cols = [col for col in ["client_name", "period_label", "project_name", "owner", "description", "notes"] if col in filtered.columns]
        mask = pd.Series(False, index=filtered.index)
        for col in cols:
            mask = mask | filtered[col].fillna("").astype(str).str.lower().str.contains(query, regex=False)
        filtered = filtered[mask]
    return filtered


def pdf_image_grid(value: object, max_images: int = 6) -> Table | str:
    images = []
    for path_text in image_paths_from_value(value)[:max_images]:
        path = resolve_image_path(path_text)
        if not path:
            continue
        try:
            image = RLImage(str(path))
            image._restrictSize(14 * mm, 14 * mm)
            images.append(image)
        except Exception:
            continue
    if not images:
        return ""
    rows = [images[index : index + 3] for index in range(0, len(images), 3)]
    for row in rows:
        while len(row) < 3:
            row.append("")
    table = Table(rows, colWidths=[15 * mm, 15 * mm, 15 * mm])
    table.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE")]))
    return table


def make_pdf_bytes(data: pd.DataFrame, title: str) -> bytes:
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=12 * mm, leftMargin=12 * mm, topMargin=12 * mm, bottomMargin=12 * mm)
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("Title", parent=styles["Heading1"], fontSize=15, leading=18, alignment=1)
    body_style = ParagraphStyle("Body", parent=styles["BodyText"], fontSize=7.5, leading=9)

    story = [Paragraph(escape(title.upper()), title_style), Spacer(1, 8)]
    rows = [["No", "Period", "Project", "Qty", "Unit", "Amount", "Image"]]
    sorted_data = data.sort_values(["period_year", "period_label", "source_file", "source_row", "id"], ascending=[False, False, True, True, True], na_position="last")
    for index, row in sorted_data.reset_index(drop=True).iterrows():
        project_parts = [f"<b>{escape(str(row.get('project_name') or ''))}</b>"]
        owner = normalize_text(row.get("owner"))
        desc = normalize_text(row.get("description"))
        if owner:
            project_parts.append(escape(owner))
        if desc:
            project_parts.append(escape(desc).replace("\n", "<br/>"))
        rows.append(
            [
                str(index + 1),
                Paragraph(escape(str(row.get("period_label") or "")), body_style),
                Paragraph("<br/>".join(project_parts), body_style),
                f"{number_or_zero(row.get('drawing_qty')):g}",
                f"SGD {number_or_zero(row.get('unit_price')):,.0f}",
                f"SGD {number_or_zero(row.get('amount')):,.0f}",
                pdf_image_grid(row.get("image_path")),
            ]
        )

    total_qty = data["drawing_qty"].sum() if "drawing_qty" in data.columns else 0
    total_amount = data["amount"].sum() if "amount" in data.columns else 0
    rows.append(["", "", Paragraph("<b>TOTAL</b>", body_style), f"{total_qty:g}", "", f"SGD {total_amount:,.0f}", ""])

    table = Table(rows, colWidths=[10 * mm, 30 * mm, 65 * mm, 14 * mm, 23 * mm, 25 * mm, 48 * mm], repeatRows=1)
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2D2D2D")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 7.5),
                ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#D9D9D9")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("ALIGN", (3, 1), (5, -1), "RIGHT"),
                ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#F3F4F6")),
            ]
        )
    )
    story.append(table)
    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()


def render_period_table(data: pd.DataFrame) -> None:
    rows = []
    sorted_data = data.sort_values(["source_file", "source_row", "id"], na_position="last").reset_index(drop=True)
    for index, row in sorted_data.iterrows():
        owner = normalize_text(row.get("owner"))
        desc = normalize_text(row.get("description"))
        owner_html = f'<span class="owner"> · {escape(owner)}</span>' if owner else ""
        desc_html = f'<div class="desc">{escape(desc).replace(chr(10), "<br/>")}</div>' if desc else ""
        rows.append(
            "<tr>"
            f"<td>{index + 1}</td>"
            f"<td><strong>{escape(str(row.get('project_name') or ''))}</strong>{owner_html}{desc_html}</td>"
            f"<td class='num'>{number_or_zero(row.get('drawing_qty')):g}</td>"
            f"<td class='num'>SGD {number_or_zero(row.get('unit_price')):,.0f}</td>"
            f"<td class='num amount'>SGD {number_or_zero(row.get('amount')):,.0f}</td>"
            f"<td>{render_images(row.get('image_path'))}</td>"
            "</tr>"
        )
    rows.append(
        "<tr class='total'>"
        "<td></td><td>TOTAL</td>"
        f"<td class='num'>{data['drawing_qty'].sum():g}</td><td></td>"
        f"<td class='num amount'>SGD {data['amount'].sum():,.0f}</td><td></td>"
        "</tr>"
    )
    html = """
    <style>
    .period-table{width:100%;border-collapse:collapse;font-size:14px;}
    .period-table th{background:#2D2D2D;color:white;padding:10px;text-align:left;position:sticky;top:0;}
    .period-table td{border-bottom:1px solid #E8E8E4;padding:10px;vertical-align:top;}
    .period-table .num{text-align:right;white-space:nowrap;}
    .period-table .amount{font-weight:700;color:#B8760A;}
    .period-table .owner,.period-table .desc{color:#6B7280;font-size:12px;}
    .period-table .desc{margin-top:5px;line-height:1.35;}
    .period-table .total td{background:#F7F7F5;font-weight:700;}
    .image-grid{display:grid;grid-template-columns:repeat(3,58px);gap:5px;}
    .thumb{width:58px;height:58px;object-fit:contain;border:1px solid #E8E8E4;border-radius:6px;background:#F7F7F5;}
    </style>
    <table class="period-table">
      <thead><tr><th>No</th><th>Project</th><th>Qty</th><th>Unit</th><th>Amount</th><th>Image</th></tr></thead>
      <tbody>
    """ + "".join(rows) + "</tbody></table>"
    st.markdown(html, unsafe_allow_html=True)


def main() -> None:
    st.title("📊 Tong Ket Manager")
    st.caption("Chế độ Cloud read-only. Dữ liệu cập nhật sau khi Local app đồng bộ lên GitHub.")

    data = load_entries()
    if data.empty:
        st.warning("Chưa có dữ liệu để hiển thị.")
        return

    st.sidebar.header("Bộ lọc")
    clients = ["Tất cả"] + sorted([value for value in data["client_name"].dropna().unique().tolist() if value])
    years = ["Tất cả"] + sorted([str(int(value)) for value in data["period_year"].dropna().unique().tolist()], reverse=True)
    periods_source = data.copy()

    selected_client = st.sidebar.selectbox("Công ty", clients)
    selected_year = st.sidebar.selectbox("Năm", years)
    if selected_client != "Tất cả":
        periods_source = periods_source[periods_source["client_name"] == selected_client]
    if selected_year != "Tất cả":
        periods_source = periods_source[periods_source["period_year"].fillna(0).astype(int).astype(str) == selected_year]
    periods = ["Tất cả"] + sorted([value for value in periods_source["period_label"].dropna().unique().tolist() if value], reverse=True)
    selected_period = st.sidebar.selectbox("Kỳ", periods)
    keyword = st.sidebar.text_input("Tìm dự án / mô tả")

    filtered = filter_entries(data, keyword, selected_client, selected_year, selected_period)

    metric_cols = st.columns(4)
    metric_cols[0].metric("Dòng", f"{len(filtered):,}")
    metric_cols[1].metric("Dự án", f"{filtered['project_name'].nunique() if not filtered.empty else 0:,}")
    metric_cols[2].metric("Drawings", f"{filtered['drawing_qty'].sum() if not filtered.empty else 0:g}")
    metric_cols[3].metric("Tổng SGD", f"{filtered['amount'].sum() if not filtered.empty else 0:,.0f}")

    tab_table, tab_gallery, tab_pdf = st.tabs(["Kỳ / Bảng dữ liệu", "Hình ảnh", "PDF"])

    with tab_table:
        if filtered.empty:
            st.info("Không có dữ liệu phù hợp bộ lọc.")
        else:
            render_period_table(filtered)

    with tab_gallery:
        image_rows = filtered[filtered["image_path"].fillna("").str.strip() != ""] if "image_path" in filtered.columns else pd.DataFrame()
        if image_rows.empty:
            st.info("Không có hình ảnh trong dữ liệu đang lọc.")
        else:
            cols = st.columns(4)
            item_index = 0
            for _, row in image_rows.iterrows():
                for path_text in image_paths_from_value(row.get("image_path")):
                    path = resolve_image_path(path_text)
                    if not path:
                        continue
                    with cols[item_index % 4]:
                        st.image(str(path), caption=f"{row.get('project_name')} · {row.get('period_label')}", use_container_width=True)
                    item_index += 1

    with tab_pdf:
        title = st.text_input("Tiêu đề PDF", value=f"{selected_client} - {selected_period}".replace("Tất cả", "Project Summary"))
        if st.button("Tạo preview PDF", type="primary", disabled=filtered.empty, use_container_width=True):
            st.session_state["cloud_pdf_bytes"] = make_pdf_bytes(filtered, title)
        pdf_bytes = st.session_state.get("cloud_pdf_bytes")
        if pdf_bytes:
            st.download_button(
                "Tải PDF",
                data=pdf_bytes,
                file_name=f"{title}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
            b64_pdf = base64.b64encode(pdf_bytes).decode("utf-8")
            st.markdown(
                f'<iframe src="data:application/pdf;base64,{b64_pdf}" width="100%" height="800" '
                'style="border:1px solid #E5E7EB;border-radius:8px;"></iframe>',
                unsafe_allow_html=True,
            )


if __name__ == "__main__":
    main()
