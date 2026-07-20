# SG Retail Label Generator

Streamlit tool lọc danh sách EAN từ Master Data và tạo PDF tem SG Retail theo mẫu cung cấp.

## Quy cách đầu ra

- Mỗi tem: 4 x 3 cm, chữ 5pt.
- Mỗi trang PDF: 9 x 4 cm, gồm 3 tem giống nhau để in/cắt.
- Nội dung: EAN, Item Code, Item Desc., RRP (SGD), CO.
- CO lấy từ file EAN và giữ dạng mã như `CN`, `VN`.
- Calibri được ưu tiên khi có sẵn; Streamlit Cloud dùng Carlito tương thích Calibri.

## File đầu vào

1. Master Data: `EAN`, `Item Code`, `Item Description`, `Currency`, `RRP (SGD)`.
2. File yêu cầu: `EAN`, `CO` (CO/COO/Country đều được nhận diện).

## Chạy local

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

Khi in, chọn **Actual size / 100%** và không chọn **Fit to page**.

