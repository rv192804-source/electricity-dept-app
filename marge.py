import datetime
import io
import re
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Electricity Dept - Master Utility Portal", layout="wide")

# ==========================================
# SESSION STATE MANAGEMENT FOR MODE
# ==========================================
if "active_mode" not in st.session_state:
    st.session_state["active_mode"] = "URJAS"

# ==========================================
# SIDEBAR - SEPARATE BUTTONS DISPLAY
# ==========================================
st.sidebar.title("🎛️ PORTAL NAVIGATION")
st.sidebar.markdown("---")

btn_urjas = st.sidebar.button(
    "⚡ URJAS PENDENCY PORTAL",
    use_container_width=True,
    type="primary" if st.session_state["active_mode"] == "URJAS" else "secondary"
)

st.sidebar.markdown("")

btn_merge = st.sidebar.button(
    "📊 MERGER & ZONE SPLITTER",
    use_container_width=True,
    type="primary" if st.session_state["active_mode"] == "MERGE" else "secondary"
)

if btn_urjas:
    st.session_state["active_mode"] = "URJAS"
    st.rerun()

if btn_merge:
    st.session_state["active_mode"] = "MERGE"
    st.rerun()

st.sidebar.markdown("---")
if st.session_state["active_mode"] == "URJAS":
    st.sidebar.success("📌 Active: **URJAS Master Portal**")
    st.sidebar.caption("👉 Create a Master Pendency & Time-Wise Dashboard by uploading a single Excel file.")
else:
    st.sidebar.success("📌 Active: **Universal Merger Portal**")
    st.sidebar.caption("👉 Merge multiple Excel files and make separate sheets for each zone.")


# ==============================================================================
# MODE 1: URJAS MASTER PENDENCY PORTAL
# ==============================================================================
if st.session_state["active_mode"] == "URJAS":
    st.title("⚡ Electricity Department - Master Pendency Report Generator")
    selected_date = st.date_input("📅 Target Pendency Date:", datetime.date.today())
    uploaded_file = st.file_uploader("Upload Raw Excel File (.xlsx)", type=["xlsx"])

    if uploaded_file:
        try:
            target_date = pd.to_datetime(selected_date).normalize()
            formatted_date_str = target_date.strftime("%d/%m/%Y")

            zones = [
                "ANNAPURNA",
                "GUMASTA NAGAR",
                "Hawa Bangla",
                "RAJ MOHALLA",
                "RAJENDRA NAGAR",
                "RAU",
                "SILICON CITY",
                "Sirpur",
            ]
            slabs = ["0 - 3 days", "4 - 6 days", "7 - 15 days", "16 - 30 days", "MORE THAN 30 DAYS"]

            sheet_configs = [
                {"title": "NSC LT Application", "keywords": ["nsc lt", "nsc"], "dc_col": "DC", "date_col": "DATEOFAPPLICATION"},
                {"title": "LT Load Change", "keywords": ["lt load change", "load change"], "dc_col": "DCNAME", "date_col": "DATEOFAPP"},
                {"title": "Meter Replacement App", "keywords": ["meter replacement"], "dc_col": "DC", "date_col": "DATEOFAPPLICATION"},
                {"title": "Bill Correction App", "keywords": ["bill correction"], "dc_col": "DC", "date_col": "DATEOFAPP"},
                {"title": "Permanent Disconnection App", "keywords": ["permanent disconnection"], "dc_col": "DESCRIPTION", "date_col": "DATEOFAPP"},
                {"title": "LT Name Transfer App", "keywords": ["lt name transfer", "name transfer"], "dc_col": "DC", "date_col": "AADHARNO"},
                {"title": "LT Change Of Category", "keywords": ["lt change of category", "change of category"], "dc_col": "DC", "date_col": "DATEOFAPP"},
                {"title": "Cabel Replacement APP", "keywords": ["cabel", "cable"], "dc_col": "DC", "date_col": "DATEOFAPP"},
                {"title": "Transformer Fail App", "keywords": ["transformer fail", "transformer", "fail app"], "dc_col": "DC", "date_col": "DATEOFAPPLICATION"},
                {"title": "LT Line/Meter Shifting App", "keywords": ["lt line", "meter shifting", "line shifting", "shifting app"], "dc_col": "DC", "date_col": "DATEOFAPPLICATION"},
            ]

            xls = pd.ExcelFile(uploaded_file)
            available_sheets = xls.sheet_names

            wb = openpyxl.Workbook()

            font_title = Font(name="Arial", size=10, bold=True)
            font_header = Font(name="Arial", size=9, bold=True)
            font_body = Font(name="Arial", size=9)
            font_total = Font(name="Arial", size=9, bold=True)
            fill_yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
            thin_border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))

            def clean_zone(z):
                if pd.isna(z): return "OTHER"
                z_str = str(z).strip().upper()
                if "ANNAPURNA" in z_str: return "ANNAPURNA"
                if "GUMASTA" in z_str: return "GUMASTA NAGAR"
                if "HAWA" in z_str or "BANGLA" in z_str: return "Hawa Bangla"
                if "MOHALLA" in z_str: return "RAJ MOHALLA"
                if "RAJENDRA" in z_str: return "RAJENDRA NAGAR"
                if "RAU" in z_str: return "RAU"
                if "SILICON" in z_str: return "SILICON CITY"
                if "SIRPUR" in z_str: return "Sirpur"
                return "OTHER"

            def get_slab(days):
                if pd.isna(days) or days < 0: return None
                try: days = int(days)
                except: return None
                if days <= 3: return "0 - 3 days"
                elif days <= 6: return "4 - 6 days"
                elif days <= 15: return "7 - 15 days"
                elif days <= 30: return "16 - 30 days"
                else: return "MORE THAN 30 DAYS"

            def parse_flexible_date(val):
                if pd.isna(val): return pd.NaT
                if isinstance(val, (datetime.datetime, datetime.date, pd.Timestamp)): return pd.to_datetime(val)
                val_str = str(val).replace("\xa0", "").strip()
                if val_str == "" or val_str.lower() in ["nan", "nat", "none"]: return pd.NaT
                try:
                    if val_str.replace(".", "", 1).isdigit():
                        num = float(val_str)
                        if num > 30000: return pd.to_datetime("1899-12-30") + pd.to_timedelta(num, unit="D")
                except: pass
                dt = pd.to_datetime(val_str, errors="coerce", dayfirst=True)
                if pd.isna(dt): dt = pd.to_datetime(val_str, errors="coerce", dayfirst=False)
                return dt

            ws_overall = wb.active
            ws_overall.title = "URJAS Pendency (Over all)"
            ws_overall.views.sheetView[0].showGridLines = True

            overall_summary = pd.DataFrame(0, index=zones, columns=[s["title"] for s in sheet_configs])
            processed_data_store = {}

            for config in sheet_configs:
                matched = next((s for s in available_sheets if any(k in s.lower() for k in config["keywords"])), None)
                if matched:
                    df = pd.read_excel(uploaded_file, sheet_name=matched)
                    zone_col = config["dc_col"] if config["dc_col"] in df.columns else next((c for c in df.columns if any(x in c.upper() for x in ["DC", "ZONE", "DESC"])), df.columns[0])
                    date_col = config["date_col"] if config["date_col"] in df.columns else next((c for c in df.columns if "DATE" in c.upper()), df.columns[1])

                    df["CleanZone"] = df[zone_col].apply(clean_zone)
                    app_dates = df[date_col].apply(parse_flexible_date)
                    df["Days"] = (target_date - app_dates).dt.days
                    df["Slab"] = df["Days"].apply(get_slab)
                    df = df[df["Days"] >= 0]

                    counts = df["CleanZone"].value_counts()
                    for z in zones:
                        overall_summary.at[z, config["title"]] = counts.get(z, 0)

                    processed_data_store[config["title"]] = df
                else:
                    processed_data_store[config["title"]] = None

            num_cols = len(sheet_configs) + 2
            ws_overall.merge_cells(start_row=1, start_column=1, end_row=1, end_column=num_cols)
            t_cell = ws_overall.cell(row=1, column=1, value=f"URJAS Pendency (Over all) Dtd.({formatted_date_str})")
            t_cell.fill = fill_yellow
            t_cell.font = font_title
            t_cell.alignment = Alignment(horizontal="center", vertical="center")

            ws_overall.cell(row=2, column=1, value="Zone").fill = fill_yellow
            ws_overall.cell(row=2, column=1).font = font_header
            ws_overall.cell(row=2, column=1).alignment = Alignment(horizontal="center", vertical="center")
            ws_overall.cell(row=2, column=1).border = thin_border

            for i, col_name in enumerate(overall_summary.columns):
                c = ws_overall.cell(row=2, column=i + 2, value=col_name)
                c.fill = fill_yellow; c.font = font_header
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.border = thin_border

            gt_head = ws_overall.cell(row=2, column=num_cols, value="Grand Total")
            gt_head.fill = fill_yellow; gt_head.font = font_header
            gt_head.alignment = Alignment(horizontal="center", vertical="center")
            gt_head.border = thin_border

            for r_idx, zone in enumerate(zones):
                row_num = r_idx + 3
                zn_cell = ws_overall.cell(row=row_num, column=1, value=zone)
                zn_cell.font = font_body; zn_cell.alignment = Alignment(horizontal="left", vertical="center"); zn_cell.border = thin_border

                row_tot = 0
                for c_idx, col_name in enumerate(overall_summary.columns):
                    val = int(overall_summary.at[zone, col_name])
                    val_cell = ws_overall.cell(row=row_num, column=c_idx + 2, value=val if val > 0 else 0)
                    val_cell.font = font_body; val_cell.alignment = Alignment(horizontal="center", vertical="center"); val_cell.border = thin_border
                    row_tot += val

                tot_cell = ws_overall.cell(row=row_num, column=num_cols, value=row_tot if row_tot > 0 else 0)
                tot_cell.font = font_body; tot_cell.alignment = Alignment(horizontal="center", vertical="center"); tot_cell.border = thin_border

            tot_row_num = len(zones) + 3
            tot_label = ws_overall.cell(row=tot_row_num, column=1, value="Total")
            tot_label.fill = fill_yellow; tot_label.font = font_total; tot_label.alignment = Alignment(horizontal="center", vertical="center"); tot_label.border = thin_border

            for c_idx, col_name in enumerate(overall_summary.columns):
                col_sum = int(overall_summary[col_name].sum())
                col_sum_cell = ws_overall.cell(row=tot_row_num, column=c_idx + 2, value=col_sum if col_sum > 0 else 0)
                col_sum_cell.fill = fill_yellow; col_sum_cell.font = font_total; col_sum_cell.alignment = Alignment(horizontal="center", vertical="center"); col_sum_cell.border = thin_border

            grand_sum_val = int(overall_summary.values.sum())
            grand_sum_cell = ws_overall.cell(row=tot_row_num, column=num_cols, value=grand_sum_val if grand_sum_val > 0 else 0)
            grand_sum_cell.fill = fill_yellow; grand_sum_cell.font = font_total; grand_sum_cell.alignment = Alignment(horizontal="center", vertical="center"); grand_sum_cell.border = thin_border

            # Time Wise Sheet
            ws_time = wb.create_sheet(title="Time Wise Pendency")
            ws_time.views.sheetView[0].showGridLines = True

            for idx, config in enumerate(sheet_configs):
                df = processed_data_store.get(config["title"])
                pvt = pd.DataFrame(0, index=zones, columns=slabs)

                if df is not None and not df.empty:
                    pvt = pd.pivot_table(df, index="CleanZone", columns="Slab", aggfunc="size", fill_value=0)
                    pvt = pvt.reindex(index=zones, columns=slabs, fill_value=0)

                col_side = idx % 2
                row_group = idx // 2
                start_col = 1 if col_side == 0 else 10
                start_row = 1 + (row_group * 13)

                ws_time.merge_cells(start_row=start_row, start_column=start_col, end_row=start_row, end_column=start_col + 7)
                t_cell = ws_time.cell(row=start_row, column=start_col, value=f"{config['title'].upper()} TIME WISE PENDENCY TILL ({formatted_date_str})")
                t_cell.font = font_title; t_cell.fill = fill_yellow; t_cell.alignment = Alignment(horizontal="center", vertical="center")

                headers = ["SR/ NO.", "Zone", "0 - 3 days", "4 - 6 days", "7 - 15 days", "16 - 30 days", "MORE THAN 30 DAYS", "Grand Total"]
                h_row = start_row + 1
                for c_idx, h_text in enumerate(headers):
                    cell = ws_time.cell(row=h_row, column=start_col + c_idx, value=h_text)
                    cell.font = font_header; cell.fill = fill_yellow; cell.alignment = Alignment(horizontal="center", vertical="center"); cell.border = thin_border

                for r_idx, z_name in enumerate(zones):
                    curr_row = h_row + 1 + r_idx
                    ws_time.cell(row=curr_row, column=start_col, value=r_idx + 1).border = thin_border
                    ws_time.cell(row=curr_row, column=start_col + 1, value=z_name).border = thin_border

                    row_sum = 0
                    for s_idx, slab in enumerate(slabs):
                        val = int(pvt.loc[z_name, slab]) if z_name in pvt.index and slab in pvt.columns else 0
                        ws_time.cell(row=curr_row, column=start_col + 2 + s_idx, value=val if val > 0 else "").border = thin_border
                        row_sum += val

                    ws_time.cell(row=curr_row, column=start_col + 7, value=row_sum if row_sum > 0 else "").border = thin_border

            output = io.BytesIO()
            wb.save(output)
            output.seek(0)

            st.success("🎉 URJAS Master File successfully generated!")
            st.download_button(
                label="📥 Download URJAS Master Complete Report (.xlsx)",
                data=output,
                file_name=f"URJAS_Master_Pendency_Report_{target_date.strftime('%d_%m_%Y')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

        except Exception as e:
            st.error(f"Error: {e}")


# ==============================================================================
# MODE 2: UNIVERSAL EXCEL MERGER & ZONE SPLITTER
# ==============================================================================
else:
    st.title("📊 Universal Excel Merger & Zone-Wise Data Generator")

    uploaded_files = st.file_uploader(
        "Upload Excel Files (.xlsx, .xls)",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
    )

    def process_and_clean_file(file):
        xls = pd.ExcelFile(file)
        
        target_sheet = xls.sheet_names[0]
        for s in xls.sheet_names:
            if "data" in str(s).lower():
                target_sheet = s
                break

        # Raw read bina header ke
        df_raw = pd.read_excel(file, sheet_name=target_sheet, header=None)

        # 1. Top 3-4 extra/blank rows hata kar actual Header Row dhundho
        header_row_idx = 0
        for idx, row in df_raw.iterrows():
            row_values = row.dropna().astype(str).str.strip().tolist()
            if len(row_values) > 2:  # Actual table header row
                header_row_idx = idx
                break

        # 2. Skipping top unwanted title rows
        df = pd.read_excel(file, sheet_name=target_sheet, skiprows=header_row_idx)

        # 3. Clean columns (Remove Unnamed & Standardize names to Uppercase)
        df = df.loc[:, ~df.columns.astype(str).str.contains("^Unnamed", na=False)]
        df.columns = [str(c).strip().upper() for c in df.columns]

        # 4. Remove Bottom Total / Blank rows
        if not df.empty:
            first_col = df.columns[0]
            df = df[~df[first_col].astype(str).str.contains(r"TOTAL|GRAND TOTAL|RECORD", case=False, na=False)]
            df = df.dropna(how="all")

        df["SOURCE_FILE"] = file.name
        return df

    if uploaded_files:
        combined_list = []
        for file in uploaded_files:
            try:
                cleaned_df = process_and_clean_file(file)
                combined_list.append(cleaned_df)
            except Exception as e:
                st.error(f"⚠️ Error reading file {file.name}: {e}")

        if combined_list:
            # Vertical Concat (Ek ke niche ek data stacking)
            merged_df = pd.concat(combined_list, ignore_index=True, axis=0)

            # Re-index S.No column
            sno_col = next((c for c in merged_df.columns if "S.NO" in c or "SL.NO" in c or "S. NO" in c), None)
            if sno_col:
                merged_df[sno_col] = range(1, len(merged_df) + 1)

            st.success(f"✅ **Merged Successfully!** Total Records: **{len(merged_df):,}**")

            # Zone Column detection for splitting
            zone_col = next((c for c in merged_df.columns if "ZONE" in c or "DC" in c), None)

            col1, col2 = st.columns(2)

            with col1:
                st.markdown("### 📥 1. Download Master Merged File")
                output_single = io.BytesIO()
                with pd.ExcelWriter(output_single, engine="openpyxl") as writer:
                    merged_df.to_excel(writer, sheet_name="All_Data", index=False)

                st.download_button(
                    label="📄 Download All Data (.xlsx)",
                    data=output_single.getvalue(),
                    file_name="Master_Merged_All_Data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True,
                )

            with col2:
                st.markdown("### 📁 2. Download Zone-Wise Split File")
                if zone_col:
                    output_zone = io.BytesIO()
                    with pd.ExcelWriter(output_zone, engine="openpyxl") as writer:
                        merged_df.to_excel(writer, sheet_name="All_Data", index=False)
                        merged_df[zone_col] = merged_df[zone_col].astype(str).str.strip()
                        unique_zones = [z for z in merged_df[zone_col].unique() if z and z.lower() != "nan"]

                        for z in sorted(unique_zones):
                            z_df = merged_df[merged_df[zone_col] == z].copy()
                            if not z_df.empty:
                                if sno_col in z_df.columns:
                                    z_df[sno_col] = range(1, len(z_df) + 1)
                                sheet_name = re.sub(r"[:*?/\\[\]]", "_", str(z))[:30]
                                z_df.to_excel(writer, sheet_name=sheet_name, index=False)

                    st.download_button(
                        label="📊 Download Zone-Wise Excel (.xlsx)",
                        data=output_zone.getvalue(),
                        file_name="Zone_Wise_Separated.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        use_container_width=True,
                    )
                else:
                    st.warning("⚠️ Zone Column nahi mila.")

            st.markdown("---")
            st.markdown("### 📋 Data Preview")
            st.dataframe(merged_df, use_container_width=True)
