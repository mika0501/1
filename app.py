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

st.title("🧾 GP Cargo 發票資料自動轉 Excel 工具（明細多列展開版）")
st.caption("自動將同一張發票的多筆 BOL# 與 SKU# 依序往下展開到獨立列，其餘欄位留白，下一張發票接續在全新列！")

# ==================== 1. 操作介面 ====================
st.markdown("---")
col_upload1, col_upload2 = st.columns(2)

with col_upload1:
    st.markdown("### 📥 步驟 1：上傳 PDF 發票檔案（必填）")
    uploaded_files = st.file_uploader(
        "請選擇或拖曳 PDF 發票至此（支援多檔批次上傳）",
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
def parse_gp_invoice_rows(file_bytes):
    """
    解析 GP Cargo 發票：
    將發票資訊與多筆 BOL/SKU 拆解為多列清單 (List of dicts)
    第一列帶有完整表頭與金額，第 2 列起僅填入對應的 BOL 與 SKU，其餘留空
    """
    full_text = ""
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                full_text += t + "\n"

    # 1. INVOICE NUMBER
    invoice_number = ""
    dash_3_match = re.search(r"\b(\d{6}\s*-\s*\d{3})\b", full_text)
    if dash_3_match:
        invoice_number = dash_3_match.group(1).strip()
    else:
        dash_old_match = re.search(r"\b([0-9]{2}-[0-9]{5,})\b", full_text)
        if dash_old_match:
            invoice_number = dash_old_match.group(1).strip()
        else:
            inv_lbl = re.search(
                r"INVOICE\s*#?\s*[:：]?\s*\n?\s*([0-9A-Za-z\s\-]+?)(?=\s+DATE|\s+\d{1,2}/\d{1,2}|$)",
                full_text,
                re.IGNORECASE
            )
            if inv_lbl:
                invoice_number = inv_lbl.group(1).strip()

    # 2. INVOICE DATE
    date_match = re.search(r"DATE\s*\n?\s*([0-9]{1,2}/[0-9]{1,2}/[0-9]{4})", full_text, re.IGNORECASE)
    if not date_match:
        date_match = re.search(r"Invoice\s*date[:：]?\s*([0-9]{1,2}/[0-9]{1,2}/[0-9]{4})", full_text, re.IGNORECASE)
    invoice_date = date_match.group(1).strip() if date_match else ""

    # 3. TOTAL AMOUNT
    total_match = re.search(r"TOTAL\s*\$?\s*([0-9,]+(?:\.[0-9]{2})?)", full_text, re.IGNORECASE)
    total_amount = ""
    if total_match:
        val = total_match.group(1).strip()
        total_amount = f"${val}"

    # 4. 提取全部 BOL# (保留前綴 "BOL#")
    bol_matches = re.findall(r"\b(BOL#[A-Za-z0-9\-]+)", full_text, re.IGNORECASE)
    clean_bols = [b.strip() for b in bol_matches if b.strip()]

    # 5. 提取全部 SKU# (保留前綴 "SKU#")
    sku_matches = re.findall(r"\b(SKU#[^\n\r]+)", full_text, re.IGNORECASE)
    clean_skus = []
    for s in sku_matches:
        item = re.sub(r"\s+", " ", s).strip()
        if item:
            clean_skus.append(item)

    # 6. 計算這張發票需要展開的列數 (取 BOL 與 SKU 數量的最大值，至少 1 列)
    row_count = max(len(clean_bols), len(clean_skus), 1)

    rows = []
    for i in range(row_count):
        cur_bol = clean_bols[i] if i < len(clean_bols) else ""
        cur_sku = clean_skus[i] if i < len(clean_skus) else ""

        # 只有第一列填入發票主要資訊，第二列以後全部留空白字串
        if i == 0:
            rows.append({
                "INVOICE NUMBER": invoice_number,
                "INVOICE DATE": invoice_date,
                "TOTAL AMOUNT": total_amount,
                "BOL": cur_bol,
                "SKU": cur_sku,
            })
        else:
            rows.append({
                "INVOICE NUMBER": "",
                "INVOICE DATE": "",
                "TOTAL AMOUNT": "",
                "BOL": cur_bol,
                "SKU": cur_sku,
            })

    return rows


# ==================== 3. 處理與輸出 ====================
if uploaded_files:
    all_rows = []
    progress_bar = st.progress(0)

    for i, file in enumerate(uploaded_files):
        file_rows = parse_gp_invoice_rows(file.read())
        all_rows.extend(file_rows)
        progress_bar.progress((i + 1) / len(uploaded_files))

    new_df = pd.DataFrame(all_rows)

    # 既有 Excel 合併累加
    if existing_file:
        try:
            old_df = pd.read_excel(existing_file)
            final_df = pd.concat([old_df, new_df], ignore_index=True)
            st.success(f"🎉 成功解析 {len(uploaded_files)} 份發票（共展開 {len(new_df)} 列明細），並已合併至原有的 Excel！目前累積共 {len(final_df)} 列。")
        except Exception as e:
            st.error(f"讀取既有 Excel 失敗: {e}")
            final_df = new_df
    else:
        final_df = new_df
        st.success(f"🎉 成功解析 {len(uploaded_files)} 份發票（共展開 {len(new_df)} 列明細資料）！")

    # 預覽結果
    st.subheader("📊 擷取結果預覽")
    st.dataframe(final_df, use_container_width=True)

    # 匯出 Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        final_df.to_excel(writer, index=False, sheet_name="GP_Invoices")
    excel_data = output.getvalue()

    st.download_button(
        label="📥 下載多列展開的 Excel 檔案 (.xlsx)",
        data=excel_data,
        file_name="GP_Cargo_發票展開表.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )
