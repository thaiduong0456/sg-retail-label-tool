from __future__ import annotations

from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

from label_tool import InputError, create_label_pdf, get_font_status, match_labels, render_pdf_pages


st.set_page_config(page_title="SG Retail Label", page_icon="🏷️", layout="wide")
st.markdown(
    """
    <style>
    .block-container {max-width: 1180px; padding-top: 2rem;}
    [data-testid="stMetricValue"] {font-size: 1.65rem;}
    .small-note {color:#64748b; font-size:.9rem;}
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("SG Retail Label Generator")
st.caption("Lọc EAN từ Master data và tạo PDF SG gồm 3 tem 4 × 3 cm trên mỗi trang.")
font_name, is_calibri = get_font_status()
if is_calibri:
    st.success("Sẵn sàng: Calibri 5pt sẽ được nhúng vào PDF.", icon="✅")
elif font_name == "Carlito":
    st.success("Sẵn sàng: Carlito 5pt (tương thích Calibri) sẽ được nhúng vào PDF.", icon="✅")
else:
    st.warning(
        f"Chưa tìm thấy Calibri; hệ thống đang dùng {font_name}. "
        "Để đúng tiêu chuẩn, hãy cài Calibri hoặc đặt Calibri.ttf và Calibrib.ttf trong thư mục assets/.",
        icon="⚠️",
    )

with st.expander("Cấu trúc file đầu vào", expanded=False):
    st.markdown(
        """
        - **Master data:** cần các cột `Item Code`, `Item Description`, `RRP`, `EAN`; `Currency` là tùy chọn.
        - **File EAN:** cần cột `EAN`; nên có `CO`/`COO` để in Country of Origin. Mỗi dòng tạo một tem và giữ nguyên thứ tự.
        - EAN/SKU được xử lý như mã định danh để tránh dạng scientific notation. CO được lấy từ file EAN và giữ dạng mã như `CN`, `VN`.
        """
    )
    template_left, template_right = st.columns(2)
    master_template = Path("examples/Master_data_template.xlsx")
    request_template = Path("examples/EAN_request_template.xlsx")
    if master_template.exists():
        template_left.download_button("Tải Master data mẫu", master_template.read_bytes(), master_template.name, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    if request_template.exists():
        template_right.download_button("Tải file EAN mẫu", request_template.read_bytes(), request_template.name, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

left, right = st.columns(2)
with left:
    master_file = st.file_uploader("1. Upload Master data", type=["xlsx", "xlsm"], key="master")
with right:
    request_file = st.file_uploader("2. Upload danh sách EAN cần tạo tem", type=["xlsx", "xlsm"], key="request")

generate = st.button("Tạo tem PDF", type="primary", use_container_width=True, disabled=not (master_file and request_file))

if generate:
    try:
        result = match_labels(master_file.getvalue(), request_file.getvalue())
        pdf_bytes = create_label_pdf(result.labels) if result.labels else b""
        st.session_state["result"] = result
        st.session_state["pdf"] = pdf_bytes
    except InputError as exc:
        st.error(str(exc))
        st.session_state.pop("result", None)
        st.session_state.pop("pdf", None)
    except Exception as exc:
        st.exception(exc)

if "result" in st.session_state:
    result = st.session_state["result"]
    pdf_bytes = st.session_state.get("pdf", b"")
    total_requested = len(result.labels) + len(result.missing)
    m1, m2, m3 = st.columns(3)
    m1.metric("EAN yêu cầu", total_requested)
    m2.metric("Tem đã tạo", len(result.labels))
    m3.metric("Không tìm thấy", len(result.missing))

    if result.duplicate_master_eans:
        st.warning(f"Master có {len(result.duplicate_master_eans)} EAN trùng. Tool dùng dòng xuất hiện đầu tiên.")
    if result.missing:
        st.warning("Một số EAN không có trong Master data. Các dòng này không được đưa vào PDF.")
        st.dataframe(pd.DataFrame(result.missing).rename(columns={"ean": "EAN", "co": "CO", "sku": "SKU", "row": "Dòng Excel"}), use_container_width=True, hide_index=True)
    blank_country_count = sum(not label.country for label in result.labels)
    if blank_country_count:
        st.warning(f"Có {blank_country_count} tem chưa có CO/Country of Origin. Hãy kiểm tra lại cột CO trong file EAN.")

    if pdf_bytes:
        filename = f"SG_retail_labels_{datetime.now():%Y%m%d_%H%M}.pdf"
        st.download_button("Tải PDF tem", pdf_bytes, filename, "application/pdf", type="primary", use_container_width=True)
        st.caption("Khi in: chọn Actual size / 100%, không chọn Fit to page.")
        st.subheader("Xem trước")
        previews = render_pdf_pages(pdf_bytes, max_pages=12)
        cols = st.columns(4)
        for index, png in enumerate(previews):
            cols[index % 4].image(png, caption=f"Tem {index + 1}", use_container_width=True)
        if len(result.labels) > len(previews):
            st.info(f"Đang hiển thị 12/{len(result.labels)} tem. PDF tải xuống chứa đầy đủ.")
