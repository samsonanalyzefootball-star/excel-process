# =========================================================
# Excel Data Cleaner - Streamlit App
# =========================================================

import pandas as pd
import numpy as np
import re
import io
from datetime import datetime

import streamlit as st

# =========================================================
# Streamlit 頁面設定
# =========================================================
st.set_page_config(
    page_title="Excel Data Cleaner",
    page_icon="🧹",
    layout="wide"
)

st.title("🧹 Excel 資料清洗與分類工具")
st.write(
    "上傳 Excel 檔案，系統會自動將資料分為 "
    "**完整資料 Clean_Data** 與 **不完整資料 Incomplete_Data**，"
    "並產生可下載的清洗結果 Excel。"
)

# =========================================================
# 2. 定義無效值
# =========================================================
INVALID_TEXTS = {
    "", "null", "none", "nan", "n/a", "na", "-", "--",
    "unknown", "nil", "missing", "not available",
    "invalid_time", "test row", "bad"
}


def normalize_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def is_missing(value):
    return normalize_text(value).lower() in INVALID_TEXTS


# =========================================================
# 3. 欄位設定
# =========================================================
DATE_COLUMNS = [
    "DOB", "Hire_Date", "Contract_End_Date",
    "Registration_Date", "Last_Purchase_Date"
]

DATETIME_COLUMNS = [
    "Last_Login_Timestamp"
]

PHONE_COLUMNS = [
    "Phone_Number", "Emergency_Contact_Phone"
]

EMAIL_COLUMNS = [
    "Email_Address"
]

NUMERIC_RULES = {
    "Age": (16, 80),
    "Salary_Monthly": (0, 1000000),
    "Annual_Bonus": (0, 1000000),
    "Years_of_Experience": (0, 50),
    "Credit_Score": (300, 850),
    "Total_Spend_USD": (0, 1000000),
    "Customer_Satisfaction_Score": (0, 10),
    "Discount_Rate": (0, 1)
}


# =========================================================
# 4. 格式檢查函數
# =========================================================
def is_invalid_date(value):
    """檢查純日期"""
    if is_missing(value):
        return True

    text = normalize_text(value)

    if isinstance(value, (pd.Timestamp, pd.Timestamp)):
        return False

    text = re.sub(r"[/.]", "-", text)

    if re.match(r"^\d{4}-\d{2}-\d{2}$", text):
        try:
            pd.to_datetime(text, format="%Y-%m-%d", errors="raise")
            return False
        except Exception:
            return True

    if re.match(r"^\d{2}-\d{2}-\d{4}$", text):
        try:
            pd.to_datetime(text, format="%d-%m-%Y", errors="raise")
            return False
        except Exception:
            return True

    return True


def is_invalid_datetime(value):
    """檢查日期時間，例如 2026-10-12 09:30:00"""
    if is_missing(value):
        return True

    text = normalize_text(value)

    if isinstance(value, pd.Timestamp):
        return False

    # 允許 010:30:00 這類格式
    text = re.sub(r"\s0(\d{2}):", r" \1:", text)

    try:
        pd.to_datetime(text, errors="raise")
        return False
    except Exception:
        return True


def is_invalid_phone(value):
    """檢查電話：8 位澳門號碼，或 853 + 8 位"""
    if is_missing(value):
        return True

    text = normalize_text(value)

    if not re.match(r"^[+()\d\s\-]+$", text):
        return True

    digits = re.sub(r"\D", "", text)

    if len(digits) == 8:
        return False

    if len(digits) == 11 and digits.startswith("853"):
        return False

    return True


def is_invalid_email(value):
    """檢查 Email 格式"""
    if is_missing(value):
        return True

    text = normalize_text(value)
    pattern = r"^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$"

    return not re.match(pattern, text)


def clean_numeric(value):
    """將 $、,、%、yrs、stars 等符號移除後轉為數值"""
    if is_missing(value):
        return np.nan

    text = normalize_text(value)
    text = text.replace(",", "").replace("$", "")
    text = re.sub(r"%$", "", text)
    text = re.sub(r"\s*(yrs|years?|stars)$", "", text, flags=re.I)

    try:
        return float(text)
    except Exception:
        return np.nan


# =========================================================
# 5. 資料清洗主函數
# =========================================================
def clean_dataframe(df):
    issues = []

    def add_issue(row_index, column, issue, original_value=""):
        record_id = ""

        if "Record_ID" in df.columns:
            record_id = df.loc[row_index, "Record_ID"]

        issues.append({
            "Row_Index": row_index,
            "Record_ID": record_id,
            "Column": column,
            "Original_Value": original_value,
            "Issue": issue
        })

    for idx, row in df.iterrows():
        for col in df.columns:
            value = row[col]

            # NULL / 空值 / N/A / unknown
            if is_missing(value):
                add_issue(
                    idx,
                    col,
                    "缺失值 / NULL / 空值 / N/A / unknown",
                    value
                )
                continue

            # 純日期欄位
            if col in DATE_COLUMNS and is_invalid_date(value):
                add_issue(
                    idx,
                    col,
                    "日期格式不一致、無效或無法解析",
                    value
                )
                continue

            # 日期時間欄位
            if col in DATETIME_COLUMNS and is_invalid_datetime(value):
                add_issue(
                    idx,
                    col,
                    "日期時間格式無效或無法解析",
                    value
                )
                continue

            # 電話格式
            if col in PHONE_COLUMNS and is_invalid_phone(value):
                add_issue(
                    idx,
                    col,
                    "電話格式不一致、長度錯誤或無效",
                    value
                )
                continue

            # Email 格式
            if col in EMAIL_COLUMNS and is_invalid_email(value):
                add_issue(
                    idx,
                    col,
                    "Email 格式無效",
                    value
                )
                continue

            # 數值格式與異常值
            if col in NUMERIC_RULES:
                num = clean_numeric(value)

                if pd.isna(num):
                    add_issue(
                        idx,
                        col,
                        "數值格式錯誤或包含文字",
                        value
                    )
                else:
                    low, high = NUMERIC_RULES[col]

                    if num < low or num > high:
                        issue_text = (
                            "數值異常：" + str(num) +
                            "，合理範圍 " + str(low) + " 至 " + str(high)
                        )

                        add_issue(
                            idx,
                            col,
                            issue_text,
                            value
                        )

    # Age 與 DOB 交叉檢查
    if "Age" in df.columns and "DOB" in df.columns:
        age = df["Age"].map(clean_numeric)

        for idx in df.index:
            dob_value = df.loc[idx, "DOB"]

            if is_missing(dob_value):
                continue

            dob_text = normalize_text(dob_value)
            dob_text = re.sub(r"[/.]", "-", dob_text)

            dob = None

            try:
                if re.match(r"^\d{4}-\d{2}-\d{2}$", dob_text):
                    dob = pd.to_datetime(
                        dob_text,
                        format="%Y-%m-%d",
                        errors="raise"
                    )

                elif re.match(r"^\d{2}-\d{2}-\d{4}$", dob_text):
                    dob = pd.to_datetime(
                        dob_text,
                        format="%d-%m-%Y",
                        errors="raise"
                    )

            except Exception:
                dob = None

            if dob is not None and pd.notna(age[idx]):
                expected_age = datetime.now().year - dob.year

                if age[idx] != expected_age:
                    issue_text = (
                        "年齡與出生日期不一致：Age=" + str(age[idx]) +
                        "，DOB=" + str(dob.date())
                    )

                    add_issue(
                        idx,
                        "Age",
                        issue_text,
                        df.loc[idx, "Age"]
                    )

                    add_issue(
                        idx,
                        "DOB",
                        issue_text,
                        df.loc[idx, "DOB"]
                    )

    issue_df = pd.DataFrame(issues)

    if issue_df.empty:
        issue_df = pd.DataFrame(
            columns=[
                "Row_Index",
                "Record_ID",
                "Column",
                "Original_Value",
                "Issue"
            ]
        )

    problem_rows = set(issue_df["Row_Index"].tolist())

    df["Data_Quality_Status"] = np.where(
        df.index.isin(problem_rows),
        "Incomplete",
        "Complete"
    )

    clean_df = df[df["Data_Quality_Status"] == "Complete"].copy()
    incomplete_df = df[df["Data_Quality_Status"] == "Incomplete"].copy()

    return clean_df, incomplete_df, issue_df


# =========================================================
# 6. 產生下載用 Excel
# =========================================================
def create_excel_download(clean_df, incomplete_df, issue_df):
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        clean_df.to_excel(
            writer,
            sheet_name="Clean_Data",
            index=False
        )

        incomplete_df.to_excel(
            writer,
            sheet_name="Incomplete_Data",
            index=False
        )

        issue_df.to_excel(
            writer,
            sheet_name="Data_Issues",
            index=False
        )

        summary_df = pd.DataFrame({
            "Item": [
                "原始資料筆數",
                "完整資料筆數",
                "不完整資料筆數",
                "問題明細數量",
                "處理時間"
            ],
            "Value": [
                len(clean_df) + len(incomplete_df),
                len(clean_df),
                len(incomplete_df),
                len(issue_df),
                datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            ]
        })

        summary_df.to_excel(
            writer,
            sheet_name="Summary",
            index=False
        )

    output.seek(0)
    return output


# =========================================================
# 7. Streamlit 上傳與處理
# =========================================================
uploaded_file = st.file_uploader(
    "請上傳 Excel 檔案",
    type=["xlsx", "xls"]
)

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file, sheet_name=0)

        st.success("檔案上傳成功")

        st.subheader("📄 原始資料預覽")
        st.dataframe(df.head(10), use_container_width=True)

        if st.button("開始清洗資料"):
            with st.spinner("正在清洗資料..."):
                clean_df, incomplete_df, issue_df = clean_dataframe(df)

            st.subheader("📊 清洗結果")

            col1, col2, col3 = st.columns(3)

            col1.metric("完整資料", len(clean_df))
            col2.metric("不完整資料", len(incomplete_df))
            col3.metric("問題明細", len(issue_df))

            tab1, tab2, tab3 = st.tabs(
                [
                    "✅ Clean_Data",
                    "⚠️ Incomplete_Data",
                    "🔍 Data_Issues"
                ]
            )

            with tab1:
                st.dataframe(clean_df, use_container_width=True)

            with tab2:
                st.dataframe(incomplete_df, use_container_width=True)

            with tab3:
                st.dataframe(issue_df, use_container_width=True)

            excel_file = create_excel_download(
                clean_df,
                incomplete_df,
                issue_df
            )

            st.download_button(
                label="⬇️ 下載清洗後 Excel",
                data=excel_file,
                file_name="Cleaned_Org_Data.xlsx",
                mime=(
                    "application/vnd.openxmlformats-"
                    "officedocument.spreadsheetml.sheet"
                )
            )

    except Exception as e:
        st.error("處理檔案時發生錯誤：")
        st.error(str(e))

else:
    st.info("請先上傳 .xlsx 或 .xls 檔案。")
