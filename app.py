import io

import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Font

st.set_page_config(page_title="รายงานสรุปการใช้งาน", layout="centered")
st.title("รายงานสรุปการใช้งานแยกตามสถานที่")

THAI_MONTHS = ["", "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
               "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]

REQUIRED_COLUMNS = ["building", "policy", "session time", "download", "upload", "start"]


def find_col(df, keyword):
    """ค้นหาคอลัมน์ที่มีคำ keyword ในชื่อ (ไม่สนตัวพิมพ์เล็ก-ใหญ่)"""
    for col in df.columns:
        if keyword in str(col).lower():
            return col
    return None


def to_hours(text):
    if pd.isna(text):
        return 0.0
    h, m, s = str(text).split(":")
    return int(h) + int(m) / 60 + int(s) / 3600


def display_name(building):
    """แปลงชื่อสถานที่ดิบเป็นชื่อเต็มสำหรับแสดงผล"""
    if building.startswith("สพ."):
        return "สถานพินิจฯ " + building[len("สพ."):]
    if building == "กรมพินิจและคุ้มครองเด็กและเยาวชน":
        return "สถานพินิจฯ กรมพินิจและคุ้มครองเด็กฯ"
    return building


def build_summary(session_df):
    building_col = find_col(session_df, "building")
    policy_col = find_col(session_df, "policy")
    time_col = find_col(session_df, "session time")
    download_col = find_col(session_df, "download")
    upload_col = find_col(session_df, "upload")
    start_col = find_col(session_df, "start")

    df = session_df.copy()
    df["_hours"] = df[time_col].apply(to_hours)
    df["_volume"] = df[download_col].fillna(0) + df[upload_col].fillna(0)

    rows = []
    for building, group in df.groupby(building_col, sort=True):
        rows.append({
            "สถานที่": display_name(str(building)),
            "ประชาชนผู้มารับบริการ (ครั้ง)": int((group[policy_col] == "Guest").sum()),
            "เจ้าหน้าที่กรมพินิจฯ (ครั้ง)": int((group[policy_col] == "Staff").sum()),
            "ชั่วโมงการใช้งาน": group["_hours"].sum(),
            "ปริมาณการใช้งาน": group["_volume"].sum()
        })
    summary = pd.DataFrame(rows)

    total_row = pd.DataFrame([{
        "สถานที่": "รวมทั้งหมด",
        "ประชาชนผู้มารับบริการ (ครั้ง)": summary["ประชาชนผู้มารับบริการ (ครั้ง)"].sum(),
        "เจ้าหน้าที่กรมพินิจฯ (ครั้ง)": summary["เจ้าหน้าที่กรมพินิจฯ (ครั้ง)"].sum(),
        "ชั่วโมงการใช้งาน": summary["ชั่วโมงการใช้งาน"].sum(),
        "ปริมาณการใช้งาน": summary["ปริมาณการใช้งาน"].sum()
    }])
    summary = pd.concat([summary, total_row], ignore_index=True)

    month_num = pd.to_datetime(df[start_col]).dt.month.mode()[0]
    title = f"จำนวน ผู้ใช้งาน แยกตาม หน่วยงานกรมพินิจฯ ประจำเดือน{THAI_MONTHS[month_num]}"

    return summary, title


def build_excel(internet_df, session_df, summary, title):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        internet_df.to_excel(writer, index=False, sheet_name="internet user")
        session_df.to_excel(writer, index=False, sheet_name="session logs")

        ws = writer.book.create_sheet("สรุป")

        ws.append([title, None, None, None, None])
        ws.merge_cells("A1:E1")
        ws["A1"].font = Font(bold=True)
        ws["A1"].alignment = Alignment(horizontal="center")

        headers = list(summary.columns)
        ws.append(headers)
        for cell in ws[2]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center")

        last_row_idx = len(summary) - 1
        for i, row in summary.iterrows():
            ws.append(list(row))
            r = ws.max_row
            is_total = (i == last_row_idx)
            for col_idx in range(1, 6):
                cell = ws.cell(row=r, column=col_idx)
                if is_total:
                    cell.font = Font(bold=True)
                if col_idx == 1:
                    cell.alignment = Alignment(horizontal="left")
                elif col_idx in (2, 3):
                    cell.alignment = Alignment(horizontal="center")
                    cell.number_format = "#,##0"
                elif col_idx == 4:
                    cell.alignment = Alignment(horizontal="center")
                    cell.number_format = "#,##0.00"
                elif col_idx == 5:
                    cell.alignment = Alignment(horizontal="right")
                    cell.number_format = "#,##0.00"

        ws.column_dimensions["A"].width = 50
        ws.column_dimensions["B"].width = 30
        ws.column_dimensions["C"].width = 30
        ws.column_dimensions["D"].width = 20
        ws.column_dimensions["E"].width = 22

    buffer.seek(0)
    return buffer


st.write("อัปโหลดไฟล์ Excel ทั้ง 2 ไฟล์ตามลำดับ")
file1 = st.file_uploader("ไฟล์ที่ 1: Internet Users", type=["xlsx", "xls"])
file2 = st.file_uploader("ไฟล์ที่ 2: Session Logs", type=["xlsx", "xls"])

if st.button("สร้างรายงาน", type="primary"):
    if not (file1 and file2):
        st.warning("กรุณาอัปโหลดไฟล์ให้ครบทั้ง 2 ไฟล์")
    else:
        internet_df = pd.read_excel(file1, engine="openpyxl")
        session_df = pd.read_excel(file2, engine="openpyxl")

        missing = [k for k in REQUIRED_COLUMNS if find_col(session_df, k) is None]
        if missing:
            st.error(f"ไม่พบคอลัมน์ที่ต้องใช้ในไฟล์ Session Logs: {', '.join(missing)}")
        else:
            summary, title = build_summary(session_df)

            st.success(title)
            st.dataframe(summary, use_container_width=True, hide_index=True)

            excel_data = build_excel(internet_df, session_df, summary, title)
            st.download_button(
                label="ดาวน์โหลดรายงาน Excel",
                data=excel_data,
                file_name="location_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
