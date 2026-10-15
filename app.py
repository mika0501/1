import io
import re
import pandas as pd
import pdfplumber
import streamlit as st

st.set_page_config(
    page_title="GP Cargo 發票資料擷取工具",
    page_icon="🧾",
    layout="wide"
)

st.title("🧾 GP Cargo 發票資料自動轉 Excel 工具（保留標籤前綴版）")
st.caption("自動完整保留「BOL#」與「SKU#」標籤前綴，支援 1 至 N 筆動態抓取並換行合併於同一儲存格！")

# ==================== 1. 操作介面 ====================
st.markdown("---")
col_upload1, col_upload2 = st.columns(2)

with col_upload1:
    st.markdown("### 📥 步驟 1：上傳 PDF 發票檔案（必填）")
    uploaded_files = st.file_uploader(
        "請選擇或拖曳 PDF 發票至此（支援多檔上傳）",
        type=["pdf"],
        accept_multiple_files=True,
        key="gp_pdf_uploader"
    )

with col_upload2:
    st.markdown("### 📑 步驟 2：上傳現有 Excel 檔案（選填，累加用）")
    existing_file = st.file_uploader(
        "若要將新發票資料接續累加至既有 Excel，請在此上傳",
        type=["xlsx", "xls"],
        key="gp_excel_uploader"
    )
st.markdown("---")


# ==================== 2. 解析核心 ====================
def parse_gp_multi_invoice(file_bytes):
    """解析 GP Cargo 發票：完整保留 BOL# 與 SKU# 前綴並動態彙總"""
    full_text = ""
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                full_text += t + "\n"

    # 1. INVOICE NUMBER
    inv_match = re.search(
        r"INVOICE\s*#?\s*\n?\s*([0-9]{6}\s*-\s*[0-9]{3}|[0-9]{2}-[0-9]+|[A-Za-z0-9\-]+)",
        full_text,
        re.IGNORECASE
    )
    invoice_number = ""
    if inv_match:
        invoice_number = inv_match.group(1).strip()
    else:
        alt_match = re.search(r"\b(\d{6}\s*-\s*\d{3}|\d{2}-\d{5,})\b", full_text)
        invoice_number = alt_match.group(1).strip() if alt_match else ""

    # 2. INVOICE DATE
    date_match = re.search(
        r"DATE\s*\n?\s*([0-9]{1,2}/[0-9]{1,2}/[0-9]{4})",
        full_text,
        re.IGNORECASE
    )
    if not date_match:
        date_match = re.search(
            r"Invoice\s*date[:：]?\s*([0-9]{1,2}/[0-9]{1,2}/[0-9]{4})",
            full_text,
            re.IGNORECASE
        )
    invoice_date = date_match.group(1).strip() if date_match else ""

    # 3. TOTAL AMOUNT
    total_match = re.search(
        r"TOTAL\s*\$?\s*([0-9,]+(?:\.[0-9]{2})?)",
        full_text,
        re.IGNORECASE
    )
    total_amount = ""
    if total_match:
        val = total_match.group(1).strip()
        total_amount = f"${val}"

    # ================= 4. 提取全部 BOL# (保留前綴 "BOL#") =================
    # 抓取包含 BOL# 前綴的代碼，遇到空格、換行或 SKU 標籤立即停住
    bol_matches = re.findall(r"\b(BOL#[A-Za-z0-9\-]+)", full_text, re.IGNORECASE)
    clean_bols = [b.strip() for b in bol_matches if b.strip()]
    bol_combined = "\n".join(clean_bols)

    # ================= 5. 提取全部 SKU# (保留前綴 "SKU#") =================
    # 抓取包含 SKU# 前綴直到行末的完整規格與說明文字
    sku_matches = re.findall(r"\b(SKU#[^\n\r]+)", full_text, re.IGNORECASE)
    clean_skus = []
    for s in sku_matches:
        item = re.sub(r"\s+", " ", s).strip()
        if item:
            clean_skus.append(item)
    sku_combined = "\n".join(clean_skus)

    return {
        "INVOICE NUMBER": invoice_number,
        "INVOICE DATE": invoice_date,
        "TOTAL AMOUNT": total_amount,
        "BOL": bol_combined,
        "SKU": sku_combined,
    }


# ==================== 3. 處理與輸出 ====================
if uploaded_files:
    records = []
    progress_bar = st.progress(0)

    for i, file in enumerate(uploaded_files):
        data = parse_gp_multi_invoice(file.read())
        records.append(data)
        progress_bar.progress((i + 1) / len(uploaded_files))

    new_df = pd.DataFrame(records)

    if existing_file:
        try:
            old_df = pd.read_excel(existing_file)
            final_df = pd.concat([old_df, new_df], ignore_index=True)
            if "INVOICE NUMBER" in final_df.columns:
                final_df = final_df.drop_duplicates(subset=["INVOICE NUMBER"], keep="last")
            st.success(f"🎉 成功解析 {len(records)} 筆發票，並已累加至原有的 Excel！目前累積共 {len(final_df)} 筆。")
        except Exception as e:
            st.error(f"讀取既有 Excel 失敗: {e}")
            final_df = new_df
    else:
        final_df = new_df
        st.success(f"🎉 成功解析 {len(records)} 筆發票資料！")

    st.subheader("📊 擷取結果預覽")
    st.dataframe(final_df, use_container_width=True)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        final_df.to_excel(writer, index=False, sheet_name="Invoices")
    excel_data = output.getvalue()

    st.download_button(
        label="📥 下載整理好的 Excel 檔案 (.xlsx)",
        data=excel_data,
        file_name="GP_Cargo_BOL_SKU整理.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )
