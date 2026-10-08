import io
import re
import pandas as pd
import pdfplumber
import streamlit as st

st.set_page_config(
    page_title="GP Cargo 發票資料擷取工具", page_icon="🧾", layout="wide"
)
st.title("🧾 GP Cargo 發票資料自動轉 Excel 工具")
st.caption(
    "已修正：排除左上角地址數字干擾，精確鎖定 INVOICE # 單號（如 26-100737）"
)


def parse_gp_invoice(file_bytes):
    """解析 GP Cargo 格式發票"""
    full_text = ""
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                full_text += t + "\n"

    # ================= 1. 擷取 INVOICE # (修正重點) =================
    # 策略 A: 精準匹配 "26-100737" 這種年份開頭帶連字號的單號格式 (2位數-5~6位數)
    inv_dash_match = re.search(r"\b([0-9]{2}-[0-9]{5,})\b", full_text)

    # 策略 B: 鎖定 INVOICE # 關鍵字下方/右側的號碼，排除純門牌號碼
    inv_label_match = re.search(
        r"INVOICE\s*#?\s*\n?\s*([A-Za-z0-9]+-[A-Za-z0-9]+)",
        full_text,
        re.IGNORECASE,
    )

    if inv_dash_match:
        invoice_number = inv_dash_match.group(1).strip()
    elif inv_label_match:
        invoice_number = inv_label_match.group(1).strip()
    else:
        # 備援：若格式不同，抓取緊接在 INVOICE # 後面的非空白字串
        fallback = re.search(
            r"INVOICE\s*#?\s*[:\s\n]+([A-Za-z0-9\-]+)", full_text, re.IGNORECASE
        )
        invoice_number = fallback.group(1).strip() if fallback else ""

    # ================= 2. 擷取 DATE =================
    # 優先抓表頭 DATE 下方的日期 (例如 10/07/2026 或 10/05/2026)
    date_match = re.search(
        r"DATE\s*\n?\s*([0-9]{1,2}/[0-9]{1,2}/[0-9]{4})",
        full_text,
        re.IGNORECASE,
    )
    if not date_match:
        date_match = re.search(
            r"Invoice\s*date[:：]?\s*([0-9]{1,2}/[0-9]{1,2}/[0-9]{4})",
            full_text,
            re.IGNORECASE,
        )
    invoice_date = date_match.group(1).strip() if date_match else ""

    # ================= 3. 擷取 TOTAL =================
    # 抓取底部 TOTAL 後面的金額 (例如 $4,520 或 4,520.00)
    total_match = re.search(
        r"TOTAL\s*\$?\s*([0-9,]+(?:\.[0-9]{2})?)", full_text, re.IGNORECASE
    )
    total_amount = ""
    if total_match:
        val = total_match.group(1).strip()
        total_amount = f"${val}"

    return {
        "INVOICE NUMBER": invoice_number,
        "INVOICE DATE": invoice_date,
        "TOTAL AMOUNT": total_amount,
    }


# ==================== 操作介面 ====================
col1, col2 = st.columns(2)

with col1:
    st.markdown("### 步驟 1：(選填) 上傳先前整理好的 Excel")
    existing_file = st.file_uploader(
        "若要將新發票資料接續累加至既有 Excel，請在此上傳",
        type=["xlsx", "xls"],
        key="existing_excel",
    )

with col2:
    st.markdown("### 步驟 2：上傳 PDF 發票檔案")
    uploaded_files = st.file_uploader(
        "請選擇或拖曳 PDF 發票（支援多個檔案同時拖入）",
        type=["pdf"],
        accept_multiple_files=True,
        key="pdf_files",
    )

if uploaded_files:
    records = []
    progress_bar = st.progress(0)

    for i, file in enumerate(uploaded_files):
        data = parse_gp_invoice(file.read())
        records.append(data)
        progress_bar.progress((i + 1) / len(uploaded_files))

    new_df = pd.DataFrame(records)

    # 合併既有 Excel 邏輯
    if existing_file:
        try:
            old_df = pd.read_excel(existing_file)
            final_df = pd.concat([old_df, new_df], ignore_index=True)
            # 自動去重（依據 INVOICE NUMBER 保留最新）
            if "INVOICE NUMBER" in final_df.columns:
                final_df = final_df.drop_duplicates(
                    subset=["INVOICE NUMBER"], keep="last"
                )
            st.success(
                f"🎉 成功解析 {len(records)} 筆發票，並累加至現有 Excel！目前共 {len(final_df)} 筆資料。"
            )
        except Exception as e:
            st.error(f"讀取既有 Excel 失敗: {e}")
            final_df = new_df
    else:
        final_df = new_df
        st.success(f"🎉 成功解析 {len(records)} 筆發票資料！")

    # 預覽表格
    st.subheader("📊 擷取結果預覽")
    st.dataframe(final_df, use_container_width=True)

    # 匯出 Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        final_df.to_excel(writer, index=False, sheet_name="GP_Invoices")
    excel_data = output.getvalue()

    st.download_button(
        label="📥 下載整理好的 Excel 檔案 (.xlsx)",
        data=excel_data,
        file_name="GP_Cargo_發票整理.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
