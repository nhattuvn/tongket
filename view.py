import sqlite3
import base64
from pathlib import Path
from datetime import datetime
import pandas as pd
import streamlit as st
from PIL import Image

# Cấu hình trang hiển thị của Streamlit
st.set_page_config(
    page_title="Tong Ket Manager - Chế độ Xem",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Đường dẫn thư mục dữ liệu (Đảm bảo đồng bộ với GitHub)
APP_DIR = Path(__file__).resolve().parent
DB_PATH = APP_DIR / "tong_ket.db"
UPLOAD_DIR = APP_DIR / "uploads"

# -------------------------------------------------------------------------
# HÀM KẾT NỐI VÀ LẤY DỮ LIỆU (CHỈ ĐỌC)
# -------------------------------------------------------------------------
def load_data_from_db():
    """Kết nối cơ sở dữ liệu SQLite và lấy toàn bộ dữ liệu bảng entries"""
    if not DB_PATH.exists():
        st.error(f"❌ Không tìm thấy file cơ sở dữ liệu `tong_ket.db` tại {DB_PATH}. Hãy đảm bảo bạn đã push file này lên GitHub.")
        return pd.DataFrame()
    
    try:
        conn = sqlite3.connect(str(DB_PATH))
        # Sử dụng thuộc tính chỉ đọc để tăng tính an toàn trên Cloud
        query = "SELECT * FROM entries"
        df = pd.read_sql_query(query, conn)
        conn.close()
        return df
    except Exception as e:
        st.error(f"Lỗi khi đọc cơ sở dữ liệu: {e}")
        return pd.DataFrame()

# -------------------------------------------------------------------------
# HÀM TẠO BÁO CÁO PDF TRỰC TUYẾN (Dựa trên cấu trúc ReportLab sẵn có)
# -------------------------------------------------------------------------
def make_pdf_bytes(df_data, title_name):
    """Tạo file PDF và trả về dữ liệu dạng Bytes để preview thay vì lưu ra ổ cứng"""
    from io import BytesIO
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors

    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=15, leftMargin=15, topMargin=15, bottomMargin=15)
    story = []
    
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'TitleStyle',
        parent=styles['Heading1'],
        fontSize=18,
        leading=22,
        textColor=colors.HexColor('#1E3A8A'),
        alignment=1 # Center
    )
    
    # Tiêu đề báo cáo
    story.append(Paragraph(f"BÁO CÁO TỔNG KẾT: {title_name.upper()}", title_style))
    story.append(Spacer(1, 15))
    
    # Tạo bảng dữ liệu (Đơn giản hóa để xuất PDF trên Cloud)
    # Bạn có thể giữ nguyên hàm make_pdf gốc của bạn nếu nó trả về bytes.
    table_data = [[ "Dự án", "Người phụ trách", "Mô tả", "Thành tiền (SGD)" ]]
    for _, row in df_data.iterrows():
        table_data.append([
            str(row.get('project_name', '')),
            str(row.get('owner', '')),
            str(row.get('description', '')),
            f"{row.get('amount_sgd', 0):,.2f}" if 'amount_sgd' in row else "0.00"
        ])
        
    t = Table(table_data, colWidths=[120, 100, 220, 100])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#2563EB')),
        ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('ALIGN', (-1,0), (-1,-1), 'RIGHT'),
        ('BOTTOMPADDING', (0,0), (-1,0), 8),
        ('BACKGROUND', (0,1), (-1,-1), colors.HexColor('#F3F4F6')),
        ('GRID', (0,0), (-1,-1), 1, colors.HexColor('#E5E7EB')),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,0), 10),
    ]))
    
    story.append(t)
    doc.build(story)
    
    buffer.seek(0)
    return buffer.getvalue()

# -------------------------------------------------------------------------
# GIAO DIỆN CHÍNH (MAIN APP)
# -------------------------------------------------------------------------
st.title("📊 Hệ Thống Quản Lý Tổng Kết (Chế độ xem Internet)")
st.caption("Ứng dụng chạy trên Streamlit Cloud - Dữ liệu cập nhật từ máy tính cá nhân")

# Tải dữ liệu
df_raw = load_data_from_db()

if not df_raw.empty:
    # Xử lý chuẩn hóa chữ hoa như bản local của bạn
    if "project_name" in df_raw.columns:
        df_raw["project_name"] = df_raw["project_name"].astype(str).str.upper()
    if "client_name" in df_raw.columns:
        df_raw["client_name"] = df_raw["client_name"].astype(str).str.upper()
    if "owner" in df_raw.columns:
        df_raw["owner"] = df_raw["owner"].astype(str).str.upper()

    # -------------------------------------------------------------------------
    # BỘ LỌC TÌM KIẾM (SIDEBAR)
    # -------------------------------------------------------------------------
    st.sidebar.header("🔍 Bộ lọc tìm kiếm")
    
    # Lọc theo Công ty / Khách hàng
    all_clients = ["TẤT CẢ"] + sorted(df_raw["client_name"].unique().tolist()) if "client_name" in df_raw.columns else ["TẤT CẢ"]
    selected_client = st.sidebar.selectbox("Công ty / Khách hàng", all_clients)
    
    # Lọc theo Dự án
    all_projects = ["TẤT CẢ"] + sorted(df_raw["project_name"].unique().tolist()) if "project_name" in df_raw.columns else ["TẤT CẢ"]
    selected_project = st.sidebar.selectbox("Tên dự án", all_projects)
    
    # Lọc theo Người phụ trách
    all_owners = ["TẤT CẢ"] + sorted(df_raw["owner"].unique().tolist()) if "owner" in df_raw.columns else ["TẤT CẢ"]
    selected_owner = st.sidebar.selectbox("Người phụ trách (Owner)", all_owners)
    
    # Tìm kiếm theo mô tả
    search_desc = st.sidebar.text_input("Tìm trong mô tả công việc")

    # Thực thi lọc dữ liệu
    df_filtered = df_raw.copy()
    if selected_client != "TẤT CẢ":
        df_filtered = df_filtered[df_filtered["client_name"] == selected_client]
    if selected_project != "TẤT CẢ":
        df_filtered = df_filtered[df_filtered["project_name"] == selected_project]
    if selected_owner != "TẤT CẢ":
        df_filtered = df_filtered[df_filtered["owner"] == selected_owner]
    if search_desc:
        df_filtered = df_filtered[df_filtered["description"].astype(str).str.contains(search_desc, case=False, na=False)]

    # -------------------------------------------------------------------------
    # HIỂN THỊ METRICS THỐNG KÊ (BENTO STYLE)
    # -------------------------------------------------------------------------
    st.subheader("📈 Thống kê nhanh")
    col1, col2, col3 = st.columns(3)
    
    total_records = len(df_filtered)
    total_amount = df_filtered["amount_sgd"].sum() if "amount_sgd" in df_filtered.columns else 0.0
    distinct_projects = df_filtered["project_name"].nunique() if "project_name" in df_filtered.columns else 0
    
    col1.metric("Tổng số đầu mục", f"{total_records} mục")
    col2.metric("Tổng doanh thu (SGD)", f"${total_amount:,.2f}")
    col3.metric("Số lượng dự án", f"{distinct_projects} dự án")

    # -------------------------------------------------------------------------
    # BẢNG DỮ LIỆU CHI TIẾT
    # -------------------------------------------------------------------------
    st.subheader("📋 Danh sách chi tiết công việc")
    st.dataframe(
        df_filtered,
        use_container_width=True,
        hide_index=True
    )

    # -------------------------------------------------------------------------
    # XỬ LÝ XEM ẢNH ĐÍNH KÈM (Nếu có đường dẫn ảnh)
    # -------------------------------------------------------------------------
    if "image_path" in df_filtered.columns:
        st.subheader("🖼️ Hình ảnh minh chứng đính kèm")
        df_with_images = df_filtered[df_filtered["image_path"].notna() & (df_filtered["image_path"] != "")]
        
        if not df_with_images.empty:
            # Tạo lưới hiển thị ảnh (3 cột)
            img_cols = st.columns(4)
            for idx, (_, row) in enumerate(df_with_images.iterrows()):
                local_img_path = UPLOAD_DIR / Path(row["image_path"]).name
                col_to_use = img_cols[idx % 4]
                
                if local_img_path.exists():
                    try:
                        img = Image.open(local_img_path)
                        col_to_use.image(img, caption=f"Dự án: {row.get('project_name','')}", use_container_width=True)
                    except:
                        col_to_use.warning("Không thể đọc định dạng ảnh")
                else:
                    col_to_use.info(f"Ảnh chưa đồng bộ: {local_img_path.name}")
        else:
            st.info("Không có hình ảnh nào đính kèm trong danh sách đang lọc.")

    # -------------------------------------------------------------------------
    # XUẤT VÀ PREVIEW BÁO CÁO PDF TRỰC TUYẾN
    # -------------------------------------------------------------------------
    st.markdown("---")
    st.subheader("📄 Xuất báo cáo PDF trực tuyến")
    
    pdf_title = st.text_input("Tiêu đề báo cáo xuất ra:", value=f"Bao_Cao_{selected_client}")
    
    if st.button("👁️ Xem trước (Preview) và Tải PDF", type="primary", use_container_width=True, disabled=df_filtered.empty):
        with st.spinner("Đang khởi tạo cấu trúc PDF..."):
            try:
                pdf_bytes = make_pdf_bytes(df_filtered, pdf_title)
                st.session_state["cloud_pdf_preview"] = pdf_bytes
                st.toast("Tạo bản xem trước thành công!", icon="✅")
            except Exception as pdf_err:
                st.error(f"Lỗi khi dựng PDF bằng ReportLab: {pdf_err}")

    # Nhúng iFrame hiển thị PDF trực tiếp nếu nút bấm đã được kích hoạt
    pdf_data = st.session_state.get("cloud_pdf_preview")
    if pdf_data:
        b64_pdf = base64.b64encode(pdf_data).decode('utf-8')
        pdf_display = f'<iframe src="data:application/pdf;base64,{b64_pdf}" width="100%" height="800px" style="border:1px solid #E5E7EB; border-radius:8px;"></iframe>'
        st.markdown(pdf_display, unsafe_allow_html=True)
else:
    st.warning("Cơ sở dữ liệu đang trống hoặc chưa được đồng bộ chính xác.")