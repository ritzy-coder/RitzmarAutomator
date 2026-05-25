import streamlit as st
import pandas as pd
import io
import msoffcrypto
import tempfile
import os
from datetime import datetime, date

# Page configuration
st.set_page_config(
    page_title="Data Processing & Volare Export",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================================
# SIDEBAR NAVIGATION
# ============================================================================

with st.sidebar:
    st.markdown("""
    <div style='text-align:center; padding: 10px 0 20px 0;'>
        <span style='font-size:2rem;'>📊</span>
        <div style='font-size:1.1rem; font-weight:700; color:#1a4f8a; margin-top:4px;'>Data Processing Suite</div>
        <div style='font-size:0.75rem; color:#888; margin-top:2px;'>v1.7</div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("### 🗂️ Modules")

    if "active_module" not in st.session_state:
        st.session_state.active_module = "Volare Export"

    modules = {
        "Volare Export": "📤 Volare Export System",
        "PTP Report": "📋 PTP Report Consolidate",
        "Worklist Consolidation": "📂 Worklist Consolidation",
        "Universal Consolidator": "🗂️ Universal Consolidator",
    }

    for key, label in modules.items():
        is_active = st.session_state.active_module == key
        btn_style = "primary" if is_active else "secondary"
        if st.button(label, key=f"nav_{key}", use_container_width=True, type=btn_style):
            st.session_state.active_module = key

    st.markdown("---")
    st.markdown("""
    <div style='font-size:0.75rem; color:#aaa; text-align:center; padding-top:8px;'>
        Built with Streamlit & Pandas
    </div>
    """, unsafe_allow_html=True)

# ============================================================================
# PTP REPORT HEADERS
# ============================================================================

PTP_HEADERS = [
    "LAN",
    "NAME (2)",
    "CATEGORY (58)",
    "TAGGING",
    "PAY-OFF (6)",
    "ACTION (59)",
    "PTP DATE (60)",
    "RFNP (62)",
    "SUBCATEGORIES (63)",
    "CALL OUT DATE (DATE TODAY)",
    "COLL (36)",
    "REMARK BY (64)",
    "NOTES (59)",
    "CALL OUT PERIOD (AM, PM)",
]

PTP_SHEET_NAME = "PTP REPORT TEMPLATE"

# ============================================================================
# HELPER: decrypt and read a sheet
# ============================================================================

def decrypt_and_read(file_buffer, password, sheet_name, skip_rows=0):
    """
    Attempt to decrypt file_buffer with password, then read sheet_name.
    If sheet_name is '__first__', load the first available sheet.
    Returns (df, sheet_used, error_msg).
    """
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp.write(file_buffer)
        tmp_path = tmp.name

    try:
        decrypted = io.BytesIO()
        with open(tmp_path, "rb") as f:
            office_file = msoffcrypto.OfficeFile(f)
            office_file.load_key(password=password)
            office_file.decrypt(decrypted)
        decrypted.seek(0)

        xls = pd.ExcelFile(decrypted)
        if sheet_name == "__first__":
            decrypted.seek(0)
            df = pd.read_excel(decrypted, sheet_name=0, skiprows=skip_rows)
            return df, xls.sheet_names[0], None
        elif sheet_name in xls.sheet_names:
            decrypted.seek(0)
            df = pd.read_excel(decrypted, sheet_name=sheet_name, skiprows=skip_rows)
            return df, sheet_name, None
        else:
            decrypted.seek(0)
            df = pd.read_excel(decrypted, sheet_name=0, skiprows=skip_rows)
            return df, xls.sheet_names[0], f"Sheet '{sheet_name}' not found; loaded '{xls.sheet_names[0]}' instead."
    except Exception as e:
        return None, None, str(e)
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)


def read_file_no_password(file_buffer, file_name, sheet_name, skip_rows=0):
    """
    Read a non-encrypted file.
    If sheet_name is '__first__', load the first available sheet.
    Returns (df, sheet_used, error_msg).
    """
    buf = io.BytesIO(bytes(file_buffer))
    try:
        if file_name.endswith(".csv"):
            df = pd.read_csv(buf, skiprows=skip_rows)
            return df, "CSV", None

        xls = pd.ExcelFile(buf)
        if sheet_name == "__first__":
            buf.seek(0)
            df = pd.read_excel(buf, sheet_name=0, skiprows=skip_rows)
            return df, xls.sheet_names[0], None
        elif sheet_name in xls.sheet_names:
            buf.seek(0)
            df = pd.read_excel(buf, sheet_name=sheet_name, skiprows=skip_rows)
            return df, sheet_name, None
        else:
            buf.seek(0)
            df = pd.read_excel(buf, sheet_name=0, skiprows=skip_rows)
            return df, xls.sheet_names[0], f"Sheet '{sheet_name}' not found; loaded '{xls.sheet_names[0]}' instead."
    except Exception as e:
        return None, None, str(e)


# ============================================================================
# MODULE: VOLARE EXPORT (original)
# ============================================================================

def render_volare_module():
    MERGE_KEY_PLACEHOLDER = "-- Select merge key --"

    VOLARE_HEADERS = [
        "S.No", "Date", "Time", "Debtor", "Account No.", "Card No.", "Service No.", "DPD",
        "Call Status", "Status", "Remark", "Remark By", "Remark Type", "Field Visit Date",
        "Collector", "Client", "Product Description", "Product Type", "Batch No",
        "Account Type", "Relation", "PTP Amount", "PTP Date", "Next Call",
        "Claim Paid Amount", "Claim Paid Date", "Dialed Number", "Days Past Write Off",
        "Balance", "Contact Type", "Cycle", "Old IC", "I.C Issue Date", "Bank Code",
        "Over Limit Amount", "Min Payment", "Due Date", "Monthly Installment", "30 Days",
        "MIA", "Area", "Call Duration", "Talk Time Duration", "Debtor ID", "Black Case No.",
        "Red Case No.", "Court Name", "Lawyer", "Legal Stage", "Legal Status", "Next Legal Follow up"
    ]

    # Initialize session state
    for key, default in [
        ("texxen_df", None), ("worklist_df", None), ("combined_df", None),
        ("volare_df", None), ("column_mapping", {}),
        ("worklist_needs_manual_password", False), ("worklist_manual_password", ""),
        ("worklist_default_attempted_token", None), ("worklist_decrypt_error", ""),
        ("worklist_loaded_sheet", "RAW"), ("custom_volare_headers", []),
        ("custom_header_static_values", {}), ("merge_stats", None),
    ]:
        if key not in st.session_state:
            st.session_state[key] = default

    def find_merge_key(df):
        for col in df.columns:
            if col.lower() == "debtor id": return col
        for col in df.columns:
            if col.lower() == "chcode": return col
        return df.columns[0] if len(df.columns) > 0 else None

    def reset_downstream():
        st.session_state.combined_df = None
        st.session_state.volare_df = None
        st.session_state.column_mapping = {}
        st.session_state.volare_row_mapping = {h: "(leave blank)" for h in VOLARE_HEADERS}
        st.session_state.pop("volare_mapping_editor", None)
        st.session_state.texxen_merge_key = MERGE_KEY_PLACEHOLDER
        st.session_state.worklist_merge_key = MERGE_KEY_PLACEHOLDER

    def all_volare_headers():
        return VOLARE_HEADERS + [h for h in st.session_state.custom_volare_headers if h not in VOLARE_HEADERS]

    st.title("📤 Data Processing & Volare Export System")
    st.markdown("---")

    # ── STEP 1 ──────────────────────────────────────────────────────────────
    st.header("Step 1: Upload Files")

    with st.expander("ℹ️ **Merge Strategy & File Info**"):
        st.markdown("""
        **Recommended Merge Configuration:**
        - **DRR File:** Already formatted with Volare headers (51 columns)
        - **WORKLIST File:** Contains 70 columns with additional details
        - **Merge Key:** Debtor ID (100% match rate confirmed)
        - **Expected Result:** Only rows matched in BOTH files are kept
        """)

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("📁 TEXXEN DRR File")
        texxen_file = st.file_uploader("Upload TEXXEN DRR file", type=["xlsx","xls","csv","xlsm","xlsb"], key="texxen_uploader")
        texxen_skip_rows = st.number_input("Skip rows before headers (TEXXEN)", min_value=0, max_value=10, value=0, key="texxen_skip_rows")
        if texxen_file:
            try:
                token = f"{texxen_file.name}:{texxen_file.size}:{texxen_skip_rows}"
                if st.session_state.get("texxen_file_token") != token or st.session_state.texxen_df is None:
                    st.session_state.texxen_file_token = token
                    reset_downstream()
                    if texxen_file.name.endswith(".csv"):
                        st.session_state.texxen_df = pd.read_csv(texxen_file, skiprows=texxen_skip_rows)
                    else:
                        st.session_state.texxen_df = pd.read_excel(texxen_file, skiprows=texxen_skip_rows)
                st.success(f"✅ TEXXEN DRR loaded: {st.session_state.texxen_df.shape[0]} rows, {st.session_state.texxen_df.shape[1]} columns")
            except Exception as e:
                st.error(f"❌ Error reading TEXXEN DRR: {e}")
                st.session_state.texxen_df = None
        elif st.session_state.texxen_df is not None:
            st.session_state.texxen_df = None
            st.session_state.texxen_file_token = None
            reset_downstream()

    with col2:
        st.subheader("🔐 WORKLIST File (Password Protected)")
        worklist_file = st.file_uploader("Upload WORKLIST file", type=["xlsx","xls","csv","xlsm","xlsb"], key="worklist_uploader")
        if worklist_file:
            default_password = "30PL2026"
            skip_rows = st.number_input("Skip rows before headers", min_value=0, max_value=10, value=8, key="worklist_skip_rows")
            worklist_file_token = f"{worklist_file.name}:{worklist_file.size}:{skip_rows}"

            if st.session_state.get("worklist_file_token") != worklist_file_token:
                st.session_state.worklist_file_token = worklist_file_token
                st.session_state.worklist_df = None
                st.session_state.worklist_needs_manual_password = False
                st.session_state.worklist_default_attempted_token = None
                st.session_state.worklist_decrypt_error = ""
                st.session_state.worklist_loaded_sheet = "RAW"
                st.session_state.worklist_manual_password = ""
                reset_downstream()

            if st.session_state.worklist_needs_manual_password:
                st.warning("Default password failed. Enter the file password below.")
            else:
                st.info("🔑 Trying standard password first (auto).")

            password_to_try = None
            should_attempt_decrypt = False

            if st.session_state.worklist_df is None:
                if st.session_state.worklist_needs_manual_password:
                    if st.session_state.worklist_decrypt_error:
                        st.caption(f"Last decrypt error: {st.session_state.worklist_decrypt_error}")
                    st.text_input("Enter WORKLIST password", type="password", key="worklist_manual_password")
                    if st.button("🔓 Retry Decrypt", key="worklist_password_retry_btn"):
                        manual_pw = st.session_state.worklist_manual_password.strip()
                        if manual_pw:
                            password_to_try = manual_pw
                            should_attempt_decrypt = True
                        else:
                            st.warning("Please enter a password before retrying.")
                else:
                    if st.session_state.worklist_default_attempted_token != worklist_file_token:
                        st.session_state.worklist_default_attempted_token = worklist_file_token
                        password_to_try = default_password
                        should_attempt_decrypt = True

            try:
                if should_attempt_decrypt and password_to_try:
                    df, sheet, err = decrypt_and_read(worklist_file.getbuffer(), password_to_try, "RAW", skip_rows)
                    if df is not None:
                        st.session_state.worklist_df = df
                        st.session_state.worklist_loaded_sheet = sheet
                        st.session_state.worklist_needs_manual_password = False
                        st.session_state.worklist_decrypt_error = ""
                        if err:
                            st.warning(err)
                    else:
                        st.session_state.worklist_df = None
                        st.session_state.worklist_decrypt_error = err
                        st.session_state.worklist_needs_manual_password = True
                        st.error(f"❌ Decryption failed: {err}")

                if st.session_state.worklist_df is not None:
                    st.success(f"✅ WORKLIST loaded: {st.session_state.worklist_df.shape[0]} rows, {st.session_state.worklist_df.shape[1]} columns ({st.session_state.worklist_loaded_sheet} sheet)")
            except Exception as e:
                st.error(f"❌ Unexpected error: {e}")
                st.session_state.worklist_df = None
        else:
            if st.session_state.worklist_df is not None or st.session_state.get("worklist_file_token"):
                st.session_state.worklist_df = None
                st.session_state.worklist_file_token = None
                st.session_state.worklist_needs_manual_password = False
                st.session_state.worklist_manual_password = ""
                st.session_state.worklist_default_attempted_token = None
                st.session_state.worklist_decrypt_error = ""
                st.session_state.worklist_loaded_sheet = "RAW"
                reset_downstream()

    # ── STEP 2 ──────────────────────────────────────────────────────────────
    if st.session_state.texxen_df is not None or st.session_state.worklist_df is not None:
        st.markdown("---")
        st.header("Step 2: Review Uploaded Data")
        tab1, tab2 = st.tabs(["TEXXEN DRR Preview", "WORKLIST Preview"])
        with tab1:
            if st.session_state.texxen_df is not None:
                col_filter = st.multiselect("Select columns to display:", options=st.session_state.texxen_df.columns.tolist(), default=[], key="texxen_col_filter")
                display_df = st.session_state.texxen_df[col_filter] if col_filter else st.session_state.texxen_df
                st.dataframe(display_df, use_container_width=True, height=400)
                with st.expander("📋 Column Names & Data Info"):
                    col_info = pd.DataFrame({'Column': st.session_state.texxen_df.columns, 'Data Type': st.session_state.texxen_df.dtypes, 'Non-Null Count': st.session_state.texxen_df.count(), 'Null Count': st.session_state.texxen_df.isnull().sum()})
                    st.dataframe(col_info, use_container_width=True)
            else:
                st.info("No TEXXEN DRR file uploaded yet.")
        with tab2:
            if st.session_state.worklist_df is not None:
                col_filter = st.multiselect("Select columns to display:", options=st.session_state.worklist_df.columns.tolist(), default=[], key="worklist_col_filter")
                display_df = st.session_state.worklist_df[col_filter] if col_filter else st.session_state.worklist_df
                st.dataframe(display_df, use_container_width=True, height=400)
                with st.expander("📋 Column Names & Data Info"):
                    col_info = pd.DataFrame({'Column': st.session_state.worklist_df.columns, 'Data Type': st.session_state.worklist_df.dtypes, 'Non-Null Count': st.session_state.worklist_df.count(), 'Null Count': st.session_state.worklist_df.isnull().sum()})
                    st.dataframe(col_info, use_container_width=True)
            else:
                st.info("No WORKLIST file uploaded yet.")

    # ── STEP 3 MERGE ────────────────────────────────────────────────────────
    if st.session_state.texxen_df is not None and st.session_state.worklist_df is not None:
        st.markdown("---")
        st.header("Step 3: Merge & Combine Data")
        st.info("ℹ️ Only rows with matching keys in **both** files will be kept.")

        texxen_key_options = [MERGE_KEY_PLACEHOLDER] + st.session_state.texxen_df.columns.tolist()
        worklist_key_options = [MERGE_KEY_PLACEHOLDER] + st.session_state.worklist_df.columns.tolist()

        if st.session_state.get("texxen_merge_key") not in texxen_key_options:
            st.session_state.texxen_merge_key = MERGE_KEY_PLACEHOLDER
        if st.session_state.get("worklist_merge_key") not in worklist_key_options:
            st.session_state.worklist_merge_key = MERGE_KEY_PLACEHOLDER

        col1, col2 = st.columns(2)
        with col1:
            texxen_selected_key = st.selectbox("🔑 Merge key from TEXXEN DRR:", options=texxen_key_options, index=0, key="texxen_merge_key")
        with col2:
            worklist_selected_key = st.selectbox("🔑 Merge key from WORKLIST:", options=worklist_key_options, index=0, key="worklist_merge_key")

        can_merge = texxen_selected_key != MERGE_KEY_PLACEHOLDER and worklist_selected_key != MERGE_KEY_PLACEHOLDER

        if st.button("🔀 Merge Data (Matched Rows Only)", key="merge_button", disabled=not can_merge):
            try:
                texxen_temp = st.session_state.texxen_df.copy()
                worklist_temp = st.session_state.worklist_df.copy()
                texxen_temp[texxen_selected_key] = texxen_temp[texxen_selected_key].astype(str).str.lower().str.strip()
                worklist_temp[worklist_selected_key] = worklist_temp[worklist_selected_key].astype(str).str.lower().str.strip()

                texxen_unique = set(texxen_temp[texxen_selected_key].dropna())
                worklist_unique = set(worklist_temp[worklist_selected_key].dropna())

                combined = pd.merge(texxen_temp, worklist_temp, left_on=texxen_selected_key, right_on=worklist_selected_key, how="inner", suffixes=("_texxen", "_worklist"))
                if texxen_selected_key != worklist_selected_key and worklist_selected_key in combined.columns:
                    combined = combined.drop(columns=[worklist_selected_key])

                for col in list(combined.columns):
                    if col.endswith("_texxen"):
                        base = col.replace("_texxen", "")
                        wl_col = base + "_worklist"
                        if wl_col in combined.columns:
                            combined[base] = combined[col].fillna(combined[wl_col])
                            combined = combined.drop(columns=[col, wl_col])

                combined.columns = [c.replace("_worklist", "").replace("_texxen", "") for c in combined.columns]
                st.session_state.combined_df = combined

                matched = len(texxen_unique & worklist_unique)
                st.session_state.merge_stats = {
                    "texxen_keys": len(texxen_unique), "worklist_keys": len(worklist_unique),
                    "matched": matched, "texxen_only": len(texxen_unique - worklist_unique),
                    "worklist_only": len(worklist_unique - texxen_unique),
                    "total_rows": combined.shape[0], "total_cols": combined.shape[1],
                }
            except Exception as e:
                st.error(f"❌ Error during merge: {e}")

        if st.session_state.combined_df is not None and st.session_state.merge_stats:
            stats = st.session_state.merge_stats
            st.success(f"✅ Merge complete: **{stats['total_rows']}** matched rows × **{stats['total_cols']}** columns")
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("TEXXEN Keys", stats["texxen_keys"])
            c2.metric("WORKLIST Keys", stats["worklist_keys"])
            c3.metric("Matched (kept)", stats["matched"])
            c4.metric("Match %", f"{(stats['matched']/max(stats['texxen_keys'],1)*100):.1f}%")
            with st.expander("📋 View Full Combined Dataset", expanded=True):
                st.dataframe(st.session_state.combined_df, use_container_width=True, height=400)

    # ── STEP 4 VOLARE GENERATOR ─────────────────────────────────────────────
    if st.session_state.combined_df is not None:
        st.markdown("---")
        st.header("Step 4: Volare Generator")

        texxen_cols_set = set(st.session_state.texxen_df.columns) if st.session_state.texxen_df is not None else set()
        worklist_cols_set = set(st.session_state.worklist_df.columns) if st.session_state.worklist_df is not None else set()

        def get_source(col):
            in_t = col in texxen_cols_set
            in_w = col in worklist_cols_set
            if in_t and in_w: return "both"
            elif in_t: return "drr"
            elif in_w: return "worklist"
            return "unknown"

        source_cols = [col for col in st.session_state.combined_df.columns if col != "_merge_source"]

        def label_for(col):
            src = get_source(col)
            prefix = {"drr": "[DRR] ", "worklist": "[WL] ", "both": "[BOTH] ", "unknown": ""}[src]
            return f"{prefix}{col}"

        labeled_options = ["(leave blank)"] + [label_for(c) for c in source_cols]
        label_to_col = {"(leave blank)": None}
        label_to_col.update({label_for(c): c for c in source_cols})

        if "volare_row_mapping" not in st.session_state:
            st.session_state.volare_row_mapping = {h: "(leave blank)" for h in all_volare_headers()}
        if "saved_volare_mapping" not in st.session_state:
            st.session_state.saved_volare_mapping = {}
        if "visible_volare_headers" not in st.session_state:
            st.session_state.visible_volare_headers = all_volare_headers().copy()

        current_all_headers = all_volare_headers()
        st.session_state.visible_volare_headers = [h for h in st.session_state.visible_volare_headers if h in current_all_headers]
        for h in current_all_headers:
            if h not in st.session_state.volare_row_mapping:
                st.session_state.volare_row_mapping[h] = "(leave blank)"

        with st.expander("➕ Add / Manage Custom Headers"):
            ch_col1, ch_col2 = st.columns([3, 1])
            with ch_col1:
                new_custom_header = st.text_input("New custom header name:", key="new_custom_header_input")
            with ch_col2:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("➕ Add Header", key="add_custom_header_btn", use_container_width=True):
                    name = new_custom_header.strip()
                    if name and name not in VOLARE_HEADERS and name not in st.session_state.custom_volare_headers:
                        st.session_state.custom_volare_headers.append(name)
                        st.session_state.volare_row_mapping[name] = "(leave blank)"
                        st.session_state.custom_header_static_values[name] = ""
                        if name not in st.session_state.visible_volare_headers:
                            st.session_state.visible_volare_headers.append(name)

        controls_col1, _ = st.columns([2.2, 1.2])
        with controls_col1:
            visible_headers = st.multiselect("Headers to show/export:", options=current_all_headers, key="visible_volare_headers")

        active_headers = [h for h in current_all_headers if h in visible_headers]
        active_headers_for_table = [h for h in active_headers if not (h in st.session_state.custom_volare_headers and st.session_state.custom_header_static_values.get(h, "").strip())]

        with st.form("volare_table_form"):
            st.markdown("#### 🗂️ Column Mapping Table")
            clear_btn = generate_btn = False

            if active_headers_for_table:
                mapping_row = {}
                for header in active_headers_for_table:
                    current = st.session_state.volare_row_mapping.get(header, "(leave blank)")
                    if current not in labeled_options:
                        current = "(leave blank)"
                    mapping_row[header] = current

                mapping_df = pd.DataFrame([mapping_row], index=["Source Column"])
                column_config = {h: st.column_config.SelectboxColumn(h, options=labeled_options, required=False, width="medium") for h in active_headers_for_table}
                edited_df = st.data_editor(mapping_df, column_config=column_config, num_rows="fixed", hide_index=False, use_container_width=True, height=210, key="volare_mapping_editor")
                for header in active_headers_for_table:
                    sel = edited_df.iloc[0].get(header, "(leave blank)")
                    if pd.isna(sel) or sel not in labeled_options:
                        sel = "(leave blank)"
                    st.session_state.volare_row_mapping[header] = sel

            st.divider()
            c1, c2, c3 = st.columns(3)
            with c2:
                clear_btn = st.form_submit_button("🔄 Clear All Mappings", use_container_width=True)
            with c3:
                generate_btn = st.form_submit_button("✅ Generate Volare Export", use_container_width=True, type="primary", disabled=len(active_headers) == 0)

        if clear_btn:
            st.session_state.volare_row_mapping = {h: "(leave blank)" for h in current_all_headers}

        if generate_btn:
            try:
                volare_data = {}
                row_count = len(st.session_state.combined_df)
                for header in active_headers:
                    if header == "S.No":
                        volare_data[header] = list(range(1, row_count + 1))
                        continue
                    if header in st.session_state.custom_volare_headers:
                        static_val = st.session_state.custom_header_static_values.get(header, "").strip()
                        if static_val:
                            volare_data[header] = [static_val] * row_count
                            continue
                    selected_label = st.session_state.volare_row_mapping.get(header, "(leave blank)")
                    actual_col = label_to_col.get(selected_label)
                    if actual_col and actual_col in st.session_state.combined_df.columns:
                        volare_data[header] = st.session_state.combined_df[actual_col].values
                    else:
                        volare_data[header] = [""] * row_count
                st.session_state.volare_df = pd.DataFrame(volare_data, columns=active_headers)
            except Exception as e:
                st.error(f"❌ Error generating export: {e}")

        if st.session_state.volare_df is not None:
            st.success(f"✅ Volare export ready: {st.session_state.volare_df.shape[0]} rows × {st.session_state.volare_df.shape[1]} columns")
            st.dataframe(st.session_state.volare_df, use_container_width=True, height=400)

    # ── EXPORT ───────────────────────────────────────────────────────────────
    if st.session_state.volare_df is not None:
        st.markdown("---")
        st.header("📥 Export Results")
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            st.session_state.volare_df.to_excel(writer, sheet_name="Volare Export", index=False)
        output.seek(0)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        st.download_button("⬇️ Download Volare Export as Excel", data=output.getvalue(), file_name=f"Volare_Export_{timestamp}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ============================================================================
# MODULE: PTP REPORT TEMPLATE CONSOLIDATE
# ============================================================================

def render_ptp_module():
    st.title("📋 PTP Report Template Consolidate")
    st.markdown("Consolidate multiple PTP Report Template sheets into a single master file.")
    st.markdown("---")

    # ── Init session state ───────────────────────────────────────────────────
    for key, default in [
        ("ptp_master_df", None),
        ("ptp_append_files", []),       # list of dicts: {name, order, df, sheet, warning}
        ("ptp_step1_done", False),
        ("ptp_consolidated_df", None),
    ]:
        if key not in st.session_state:
            st.session_state[key] = default

    # ============================================================
    # STEP 1: Initiation and Data Input
    # ============================================================
    st.header("Step 1: Initiation & Data Input")
    st.markdown("Choose how to start your PTP report:")

    option = st.radio(
        "Select starting option:",
        options=["Option A: Upload Existing Report", "Option B: Create New Template"],
        key="ptp_step1_option",
        horizontal=True,
    )

    # ── OPTION A ─────────────────────────────────────────────────
    if option == "Option A: Upload Existing Report":
        st.markdown(f"Upload a file containing the **`{PTP_SHEET_NAME}`** sheet. If password-protected, enter the password below.")

        ptp_a_file = st.file_uploader(
            "Upload existing PTP Report file",
            type=["xlsx", "xls", "xlsm", "xlsb", "csv"],
            key="ptp_a_uploader",
        )

        if ptp_a_file:
            ptp_a_token = f"{ptp_a_file.name}:{ptp_a_file.size}"
            if st.session_state.get("ptp_a_file_token") != ptp_a_token:
                st.session_state.ptp_a_file_token = ptp_a_token
                st.session_state.ptp_master_df = None
                st.session_state.ptp_step1_done = False
                st.session_state.ptp_consolidated_df = None
                st.session_state.pop("ptp_a_needs_pw", None)
                st.session_state.pop("ptp_a_pw_error", None)

            needs_pw = st.session_state.get("ptp_a_needs_pw", False)
            pw_error = st.session_state.get("ptp_a_pw_error", "")

            if needs_pw:
                st.warning(f"🔐 File appears to be password-protected.{' Error: ' + pw_error if pw_error else ''}")
                ptp_a_pw = st.text_input("Enter file password:", type="password", key="ptp_a_password_input")
                load_btn = st.button("🔓 Decrypt & Load", key="ptp_a_decrypt_btn")

                if load_btn:
                    pw = ptp_a_pw.strip()
                    if not pw:
                        st.warning("Please enter a password.")
                    else:
                        df, sheet, err = decrypt_and_read(ptp_a_file.getbuffer(), pw, PTP_SHEET_NAME)
                        if df is not None:
                            st.session_state.ptp_master_df = df
                            st.session_state.ptp_step1_done = True
                            st.session_state.ptp_a_needs_pw = False
                            st.session_state.ptp_a_pw_error = ""
                            st.session_state.ptp_consolidated_df = None
                            if err:
                                st.warning(err)
                        else:
                            st.session_state.ptp_a_pw_error = err
                            st.error(f"❌ Decryption failed: {err}")
            else:
                # Try without password first
                if st.session_state.ptp_master_df is None and not st.session_state.ptp_step1_done:
                    buf = ptp_a_file.getbuffer()
                    df, sheet, err = read_file_no_password(buf, ptp_a_file.name, PTP_SHEET_NAME)
                    if df is not None:
                        st.session_state.ptp_master_df = df
                        st.session_state.ptp_step1_done = True
                        st.session_state.ptp_consolidated_df = None
                        if err:
                            st.warning(err)
                    else:
                        # Could be encrypted
                        st.session_state.ptp_a_needs_pw = True
                        st.session_state.ptp_a_pw_error = err or ""
                        st.rerun()

            if st.session_state.ptp_master_df is not None:
                df = st.session_state.ptp_master_df
                st.success(f"✅ Base report loaded: **{df.shape[0]}** rows × **{df.shape[1]}** columns (sheet: `{PTP_SHEET_NAME}`)")
                with st.expander("👁️ Preview base report"):
                    st.dataframe(df.head(20), use_container_width=True, height=300)

    # ── OPTION B ─────────────────────────────────────────────────
    else:
        st.markdown("Generate a **fresh blank template** with the standard PTP Report headers.")
        st.markdown("**Headers that will be generated:**")
        st.code(", ".join(PTP_HEADERS))

        if st.button("🆕 Generate Blank Template", key="ptp_generate_template_btn", type="primary"):
            st.session_state.ptp_master_df = pd.DataFrame(columns=PTP_HEADERS)
            st.session_state.ptp_step1_done = True
            st.session_state.ptp_consolidated_df = None
            st.session_state.ptp_append_files = []

        if st.session_state.ptp_master_df is not None and st.session_state.ptp_step1_done:
            st.success(f"✅ Blank template ready with **{len(PTP_HEADERS)}** headers.")
            st.dataframe(st.session_state.ptp_master_df, use_container_width=True, height=150)

    # ============================================================
    # STEP 2: Multi-File Data Append
    # ============================================================
    if st.session_state.ptp_step1_done and st.session_state.ptp_master_df is not None:
        st.markdown("---")
        st.header("Step 2: Multi-File Data Append")
        st.markdown(
            f"Upload additional status report files. Data from the **`{PTP_SHEET_NAME}`** sheet of each file "
            "will be appended to the master list. Use the **order numbers** to control the append sequence."
        )

        # ── Upload area ──────────────────────────────────────────
        new_files = st.file_uploader(
            "Upload one or more additional status report files:",
            type=["xlsx", "xls", "xlsm", "xlsb", "csv"],
            accept_multiple_files=True,
            key="ptp_append_uploader",
        )

        # Detect newly added files and queue them
        existing_tokens = {f["token"] for f in st.session_state.ptp_append_files}
        for nf in (new_files or []):
            token = f"{nf.name}:{nf.size}"
            if token not in existing_tokens:
                # Auto-assign next order number
                next_order = len(st.session_state.ptp_append_files) + 1
                st.session_state.ptp_append_files.append({
                    "token": token,
                    "name": nf.name,
                    "order": next_order,
                    "raw": nf.getbuffer(),
                    "df": None,
                    "sheet": None,
                    "warning": None,
                    "needs_pw": False,
                    "pw_error": "",
                    "loaded": False,
                })
                existing_tokens.add(token)

        # Remove files that were deselected from the uploader
        if new_files is not None:
            active_tokens = {f"{nf.name}:{nf.size}" for nf in new_files}
            st.session_state.ptp_append_files = [
                f for f in st.session_state.ptp_append_files if f["token"] in active_tokens
            ]

        if not st.session_state.ptp_append_files:
            st.info("No additional files uploaded yet. Upload files above to append data.")
        else:
            # ── Auto-probe unprobed files (try no-password first) ─────────
            needs_rerun = False
            for fdata in st.session_state.ptp_append_files:
                if not fdata["loaded"] and not fdata["needs_pw"] and not fdata.get("probed"):
                    fdata["probed"] = True
                    df, sheet, err = read_file_no_password(fdata["raw"], fdata["name"], PTP_SHEET_NAME)
                    if df is not None:
                        fdata["df"] = df
                        fdata["sheet"] = sheet
                        fdata["loaded"] = True
                        fdata["warning"] = err
                    else:
                        fdata["needs_pw"] = True
                        fdata["pw_error"] = ""
                    needs_rerun = True
            if needs_rerun:
                st.rerun()

            # ── Order editor ──────────────────────────────────────
            st.markdown("### 📋 File Queue & Order Management")
            st.caption(
                "Edit the **Order** column to set the exact append sequence. "
                "Lower numbers are appended first."
            )

            queue_data = pd.DataFrame([
                {
                    "#": i + 1,
                    "Order": f["order"],
                    "File Name": f["name"],
                    "Status": (
                        "✅ Loaded" if f["loaded"]
                        else ("🔐 Needs Password" if f["needs_pw"] else "⏳ Pending")
                    ),
                }
                for i, f in enumerate(st.session_state.ptp_append_files)
            ])

            edited_queue = st.data_editor(
                queue_data,
                column_config={
                    "#": st.column_config.NumberColumn("#", disabled=True, width="small"),
                    "Order": st.column_config.NumberColumn("Order", min_value=1, max_value=99, step=1, width="small"),
                    "File Name": st.column_config.TextColumn("File Name", disabled=True),
                    "Status": st.column_config.TextColumn("Status", disabled=True),
                },
                num_rows="fixed",
                hide_index=True,
                use_container_width=True,
                key="ptp_queue_editor",
            )

            for i, row in edited_queue.iterrows():
                if i < len(st.session_state.ptp_append_files):
                    try:
                        st.session_state.ptp_append_files[i]["order"] = int(row["Order"])
                    except (ValueError, TypeError):
                        pass

            # ── Check how many files need a password ──────────────
            pw_needed_files = [f for f in st.session_state.ptp_append_files if f["needs_pw"] and not f["loaded"]]
            n_loaded = sum(1 for f in st.session_state.ptp_append_files if f["loaded"])
            n_total = len(st.session_state.ptp_append_files)

            # ── Decrypt section ───────────────────────────────────
            if pw_needed_files:
                st.markdown("### 🔐 Decrypt Password-Protected Files")

                # Status summary
                col_s1, col_s2, col_s3 = st.columns(3)
                col_s1.metric("Total Files", n_total)
                col_s2.metric("✅ Already Loaded", n_loaded)
                col_s3.metric("🔐 Need Password", len(pw_needed_files))

                st.markdown("**Choose your decryption method:**")

                decrypt_mode = st.radio(
                    "Decryption mode:",
                    options=[
                        "🔑 Option 1: Same Password for All — Decrypt all locked files at once",
                        "🗝️ Option 2: Different Passwords — Decrypt each file individually",
                    ],
                    key="ptp_decrypt_mode",
                    label_visibility="collapsed",
                )

                st.markdown("---")

                # ════════════════════════════════════════════════════
                # OPTION 1 — BULK DECRYPT (same password for all)
                # ════════════════════════════════════════════════════
                if "Option 1" in decrypt_mode:
                    st.markdown(
                        f"Enter **one shared password** to decrypt all **{len(pw_needed_files)}** locked file(s) at once."
                    )

                    # Show the locked files list
                    locked_names = [f["name"] for f in pw_needed_files]
                    for idx, name in enumerate(locked_names, 1):
                        st.markdown(f"&nbsp;&nbsp;&nbsp;{idx}. 📄 `{name}`", unsafe_allow_html=True)

                    bulk_pw = st.text_input(
                        "Shared password for all locked files:",
                        type="password",
                        key="ptp_bulk_password",
                        placeholder="Enter password…",
                    )

                    bulk_btn = st.button(
                        f"🔓 Decrypt All {len(pw_needed_files)} File(s) Now",
                        key="ptp_bulk_decrypt_btn",
                        type="primary",
                        use_container_width=True,
                    )

                    if bulk_btn:
                        pw = bulk_pw.strip()
                        if not pw:
                            st.warning("Please enter a password before decrypting.")
                        else:
                            success_count = 0
                            fail_count = 0
                            progress = st.progress(0, text="Decrypting files…")
                            for idx, fdata in enumerate(pw_needed_files):
                                progress.progress(
                                    (idx) / len(pw_needed_files),
                                    text=f"Decrypting {fdata['name']}…"
                                )
                                df, sheet, err = decrypt_and_read(fdata["raw"], pw, PTP_SHEET_NAME)
                                if df is not None:
                                    fdata["df"] = df
                                    fdata["sheet"] = sheet
                                    fdata["loaded"] = True
                                    fdata["needs_pw"] = False
                                    fdata["pw_error"] = ""
                                    if err:
                                        fdata["warning"] = err
                                    success_count += 1
                                else:
                                    fdata["pw_error"] = err
                                    fail_count += 1
                            progress.progress(1.0, text="Done!")

                            if success_count:
                                st.success(f"✅ Successfully decrypted **{success_count}** file(s).")
                            if fail_count:
                                st.error(
                                    f"❌ **{fail_count}** file(s) failed — wrong password or unsupported format. "
                                    "Switch to **Option 2** to enter individual passwords."
                                )
                            st.rerun()

                # ════════════════════════════════════════════════════
                # OPTION 2 — INDIVIDUAL DECRYPT (different passwords)
                # ════════════════════════════════════════════════════
                else:
                    st.markdown(
                        "Enter a **separate password** for each locked file below, then decrypt them individually "
                        "or use **Decrypt All** after filling in all passwords."
                    )

                    # Individual password inputs
                    for fdata in pw_needed_files:
                        tok = fdata["token"]
                        c1, c2, c3 = st.columns([3, 2, 1])
                        with c1:
                            st.markdown(
                                f"<div style='padding:8px 0 2px 0; font-weight:600;'>📄 {fdata['name']}</div>"
                                + (f"<div style='color:#c0392b; font-size:0.8rem;'>❌ {fdata['pw_error']}</div>" if fdata["pw_error"] else ""),
                                unsafe_allow_html=True,
                            )
                        with c2:
                            st.text_input(
                                f"Password:",
                                type="password",
                                key=f"ptp_indiv_pw_{tok}",
                                placeholder="Enter password…",
                                label_visibility="collapsed",
                            )
                        with c3:
                            st.markdown("<div style='padding-top:4px;'></div>", unsafe_allow_html=True)
                            if st.button("🔓", key=f"ptp_indiv_btn_{tok}", use_container_width=True, help=f"Decrypt {fdata['name']}"):
                                pw = st.session_state.get(f"ptp_indiv_pw_{tok}", "").strip()
                                if not pw:
                                    st.warning(f"Enter a password for '{fdata['name']}' first.")
                                else:
                                    df, sheet, err = decrypt_and_read(fdata["raw"], pw, PTP_SHEET_NAME)
                                    if df is not None:
                                        fdata["df"] = df
                                        fdata["sheet"] = sheet
                                        fdata["loaded"] = True
                                        fdata["needs_pw"] = False
                                        fdata["pw_error"] = ""
                                        if err:
                                            fdata["warning"] = err
                                        st.rerun()
                                    else:
                                        fdata["pw_error"] = err
                                        st.error(f"❌ Wrong password for '{fdata['name']}': {err}")

                    st.markdown("---")
                    # Decrypt-all-at-once using individual fields
                    if st.button(
                        "🔓 Decrypt All (using passwords entered above)",
                        key="ptp_indiv_decrypt_all_btn",
                        use_container_width=True,
                    ):
                        success_count = fail_count = skipped = 0
                        for fdata in pw_needed_files:
                            tok = fdata["token"]
                            pw = st.session_state.get(f"ptp_indiv_pw_{tok}", "").strip()
                            if not pw:
                                skipped += 1
                                continue
                            df, sheet, err = decrypt_and_read(fdata["raw"], pw, PTP_SHEET_NAME)
                            if df is not None:
                                fdata["df"] = df
                                fdata["sheet"] = sheet
                                fdata["loaded"] = True
                                fdata["needs_pw"] = False
                                fdata["pw_error"] = ""
                                if err:
                                    fdata["warning"] = err
                                success_count += 1
                            else:
                                fdata["pw_error"] = err
                                fail_count += 1

                        if success_count:
                            st.success(f"✅ Decrypted **{success_count}** file(s).")
                        if fail_count:
                            st.error(f"❌ **{fail_count}** file(s) failed — check passwords above.")
                        if skipped:
                            st.warning(f"⚠️ **{skipped}** file(s) skipped — no password entered.")
                        st.rerun()

            # ── Show per-file load results ─────────────────────────
            if n_loaded > 0:
                st.markdown("### ✅ Loaded Files")
                for i, fdata in enumerate(st.session_state.ptp_append_files):
                    if fdata["loaded"]:
                        with st.expander(f"📄 File {i+1}: {fdata['name']} — {fdata['df'].shape[0]} rows", expanded=False):
                            st.success(f"Sheet: `{fdata['sheet']}` · {fdata['df'].shape[0]} rows × {fdata['df'].shape[1]} columns")
                            if fdata.get("warning"):
                                st.warning(fdata["warning"])
                            st.dataframe(fdata["df"].head(5), use_container_width=True)

            # ── Append / Consolidate button ──────────────────────
            st.markdown("---")
            all_loaded = all(f["loaded"] for f in st.session_state.ptp_append_files)
            n_loaded = sum(1 for f in st.session_state.ptp_append_files if f["loaded"])
            n_total = len(st.session_state.ptp_append_files)

            if not all_loaded and n_total > 0:
                st.warning(f"⚠️ {n_total - n_loaded} file(s) still pending. Resolve password issues above, or proceed with {n_loaded} loaded file(s).")

            if st.button("🔀 Consolidate All Files", key="ptp_consolidate_btn", type="primary", disabled=n_loaded == 0):
                # Sort by user-defined order, then by original index (stable)
                sorted_files = sorted(
                    [(i, f) for i, f in enumerate(st.session_state.ptp_append_files) if f["loaded"]],
                    key=lambda x: (x[1]["order"], x[0])
                )

                frames = []
                base_df = st.session_state.ptp_master_df.copy()

                # Align base df to PTP_HEADERS
                for h in PTP_HEADERS:
                    if h not in base_df.columns:
                        base_df[h] = ""

                if len(base_df) > 0:
                    frames.append(base_df[PTP_HEADERS])

                for _, fdata in sorted_files:
                    df = fdata["df"].copy()
                    # Align columns: keep only PTP_HEADERS, add missing as blank
                    for h in PTP_HEADERS:
                        if h not in df.columns:
                            # Try case-insensitive match
                            matched = [c for c in df.columns if c.strip().upper() == h.strip().upper()]
                            if matched:
                                df = df.rename(columns={matched[0]: h})
                            else:
                                df[h] = ""
                    frames.append(df[PTP_HEADERS])

                consolidated = pd.concat(frames, ignore_index=True)

                # Auto-fill CALL OUT DATE with today's date where blank
                today_str = date.today().strftime("%d/%m/%Y")
                call_out_col = "CALL OUT DATE (DATE TODAY)"
                if call_out_col in consolidated.columns:
                    mask = consolidated[call_out_col].isna() | (consolidated[call_out_col].astype(str).str.strip() == "")
                    consolidated.loc[mask, call_out_col] = today_str

                st.session_state.ptp_consolidated_df = consolidated
                st.success(f"✅ Consolidation complete: **{consolidated.shape[0]}** total rows from {1 + n_loaded} source(s).")

    # ============================================================
    # STEP 3: Data Review and Export
    # ============================================================
    if st.session_state.ptp_consolidated_df is not None:
        st.markdown("---")
        st.header("Step 3: Data Review & Export")

        df = st.session_state.ptp_consolidated_df
        st.subheader("📊 Generated Data Preview")

        # Summary metrics
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total Rows", df.shape[0])
        m2.metric("Total Columns", df.shape[1])
        m3.metric("Files Appended", len([f for f in st.session_state.ptp_append_files if f["loaded"]]))
        filled = df.notna().sum().sum()
        total_cells = df.shape[0] * df.shape[1]
        m4.metric("Data Fill Rate", f"{filled/max(total_cells,1)*100:.1f}%")

        # Column filter
        col_filter = st.multiselect(
            "Filter columns to display:",
            options=df.columns.tolist(),
            default=df.columns.tolist(),
            key="ptp_preview_col_filter",
        )
        display_df = df[col_filter] if col_filter else df
        st.dataframe(display_df, use_container_width=True, height=450)

        # Column info
        with st.expander("📋 Column Summary"):
            col_info = pd.DataFrame({
                "Column": df.columns,
                "Non-Null": df.count().values,
                "Null": df.isnull().sum().values,
                "Fill %": (df.count() / max(len(df), 1) * 100).round(1).values,
                "Sample": [str(df[c].dropna().iloc[0])[:40] if df[c].notna().any() else "—" for c in df.columns],
            })
            st.dataframe(col_info, use_container_width=True)

        # Export
        st.markdown("---")
        st.subheader("📥 Download Consolidated Report")

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name=PTP_SHEET_NAME, index=False)
            wb = writer.book
            ws = writer.sheets[PTP_SHEET_NAME]
            # Auto-size columns
            for col_cells in ws.columns:
                max_len = max((len(str(cell.value or "")) for cell in col_cells), default=10)
                ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 3, 50)
        output.seek(0)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"PTP_Consolidated_{timestamp}.xlsx"

        st.download_button(
            label="⬇️ Download Consolidated Report (.xlsx)",
            data=output.getvalue(),
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
        st.success(f"✅ Ready for download: **{filename}** — {df.shape[0]} rows, sheet name: `{PTP_SHEET_NAME}`")

        # Reset button
        st.markdown("---")
        if st.button("🔄 Start Over (Reset PTP Module)", key="ptp_reset_btn"):
            for key in ["ptp_master_df", "ptp_append_files", "ptp_step1_done", "ptp_consolidated_df",
                        "ptp_a_file_token", "ptp_a_needs_pw", "ptp_a_pw_error"]:
                st.session_state.pop(key, None)
            st.rerun()


# ============================================================================
# MODULE: WORKLIST CONSOLIDATION
# ============================================================================

WORKLIST_CAMPAIGNS = ["30 DPD", "60 DPD", "XDAYS"]

def render_worklist_module():
    st.title("📂 Worklist Consolidation")
    st.markdown("Consolidate multiple worklist files by campaign into a single master sheet.")
    st.markdown("---")

    # ── Init session state ───────────────────────────────────────────────────
    wl_defaults = [
        ("wl_campaign", None),
        ("wl_header_row", 0),           # 0-based skiprows for pd.read_excel
        ("wl_files", []),               # list of file dicts
        ("wl_all_columns", []),         # union of all loaded columns
        ("wl_selected_columns", []),    # columns user wants to keep/export
        ("wl_consolidated_df", None),
        ("wl_step2_ready", False),
    ]
    for key, default in wl_defaults:
        if key not in st.session_state:
            st.session_state[key] = default

    # ════════════════════════════════════════════════════════════════════════
    # STEP 1 — Campaign Selection & File Upload
    # ════════════════════════════════════════════════════════════════════════
    st.header("Step 1: Campaign Selection & File Upload")

    # ── Campaign picker ──────────────────────────────────────────────────────
    st.markdown("#### 🎯 Select Campaign")
    campaign_cols = st.columns(len(WORKLIST_CAMPAIGNS))
    for idx, camp in enumerate(WORKLIST_CAMPAIGNS):
        with campaign_cols[idx]:
            is_sel = st.session_state.wl_campaign == camp
            btn_type = "primary" if is_sel else "secondary"
            badge_color = "#1a4f8a" if is_sel else "#888"
            st.markdown(
                f"<div style='text-align:center; margin-bottom:4px;'>"
                f"<span style='background:{badge_color};color:white;padding:3px 14px;"
                f"border-radius:12px;font-size:0.85rem;font-weight:700;'>{camp}</span></div>",
                unsafe_allow_html=True,
            )
            if st.button(
                f"{'✅ ' if is_sel else ''}Select {camp}",
                key=f"wl_camp_{camp}",
                use_container_width=True,
                type=btn_type,
            ):
                if st.session_state.wl_campaign != camp:
                    st.session_state.wl_campaign = camp
                    # Default header row per campaign: 30 DPD → row 1 (skip 0), others → row 2 (skip 1)
                    default_skip = 0 if camp == "30 DPD" else 1
                    st.session_state.wl_header_row = default_skip
                    st.session_state.wl_header_row_widget = default_skip + 1
                    # Reset downstream when campaign changes
                    st.session_state.wl_files = []
                    st.session_state.wl_all_columns = []
                    st.session_state.wl_selected_columns = []
                    st.session_state.wl_consolidated_df = None
                    st.session_state.wl_step2_ready = False
                    st.rerun()

    if not st.session_state.wl_campaign:
        st.info("☝️ Select a campaign above to continue.")
        return

    st.success(f"✅ Campaign selected: **{st.session_state.wl_campaign}**")

    # ── Header row selector ──────────────────────────────────────────────────
    st.markdown("#### 📍 Header Row Configuration")
    hr_col1, hr_col2 = st.columns([2, 3])

    # Seed the widget key from session state only once (on first render or campaign change)
    if "wl_header_row_widget" not in st.session_state:
        st.session_state.wl_header_row_widget = st.session_state.wl_header_row + 1

    with hr_col1:
        header_row_display = st.number_input(
            "Header starts at row (Excel row number):",
            min_value=1,
            max_value=20,
            step=1,
            key="wl_header_row_widget",
            help="Row 1 = first row of the file. 30 DPD default: Row 1 · 60 DPD & XDAYS default: Row 2",
        )
        new_skip = int(st.session_state.wl_header_row_widget) - 1
        if new_skip != st.session_state.wl_header_row:
            st.session_state.wl_header_row = new_skip
            # Re-probe all files with new skip value
            for f in st.session_state.wl_files:
                f["loaded"] = False
                f["probed"] = False
                f["df"] = None
                f["sheet"] = None
                f["warning"] = None
                f["needs_pw"] = False
                f["pw_error"] = ""
            st.session_state.wl_step2_ready = False
            st.session_state.wl_consolidated_df = None
            st.session_state.wl_all_columns = []
            st.session_state.wl_selected_columns = []
            st.rerun()
    with hr_col2:
        campaign_defaults = {"30 DPD": "Row 1 (A1)", "60 DPD": "Row 2 (A2)", "XDAYS": "Row 2 (A2)"}
        st.info(
            f"**Default for {st.session_state.wl_campaign}:** {campaign_defaults.get(st.session_state.wl_campaign, 'Row 1')}  \n"
            "Change this if your file's column headers start on a different row."
        )

    st.markdown("---")

    # ── File upload ──────────────────────────────────────────────────────────
    st.markdown("#### 📁 Upload Worklist Files")
    st.caption(
        "Upload one or more Excel files. "
        "Files without a password load automatically. "
        "Encrypted files will prompt for a password."
    )

    new_files = st.file_uploader(
        f"Upload {st.session_state.wl_campaign} worklist files:",
        type=["xlsx", "xls", "xlsm", "xlsb"],
        accept_multiple_files=True,
        key="wl_uploader",
    )

    # Sync file queue
    existing_tokens = {f["token"] for f in st.session_state.wl_files}
    for nf in (new_files or []):
        token = f"{nf.name}:{nf.size}"
        if token not in existing_tokens:
            next_order = len(st.session_state.wl_files) + 1
            st.session_state.wl_files.append({
                "token": token,
                "name": nf.name,
                "order": next_order,
                "raw": bytes(nf.getbuffer()),
                "df": None,
                "sheet": None,
                "warning": None,
                "needs_pw": False,
                "pw_error": "",
                "loaded": False,
                "probed": False,
            })
            existing_tokens.add(token)

    if new_files is not None:
        active_tokens = {f"{nf.name}:{nf.size}" for nf in new_files}
        removed = [f for f in st.session_state.wl_files if f["token"] not in active_tokens]
        st.session_state.wl_files = [f for f in st.session_state.wl_files if f["token"] in active_tokens]
        if removed:
            # Rebuild columns if files removed
            st.session_state.wl_step2_ready = False
            st.session_state.wl_consolidated_df = None

    if not st.session_state.wl_files:
        st.info("No files uploaded yet.")
        return

    # ── Auto-probe unprobed files ────────────────────────────────────────────
    needs_rerun = False
    skip_rows = st.session_state.wl_header_row
    for fdata in st.session_state.wl_files:
        if not fdata["loaded"] and not fdata["needs_pw"] and not fdata["probed"]:
            fdata["probed"] = True
            df, sheet, err = read_file_no_password(fdata["raw"], fdata["name"], "__first__", skip_rows=skip_rows)
            if df is not None:
                fdata["df"] = df
                fdata["sheet"] = sheet
                fdata["loaded"] = True
                fdata["warning"] = err
            else:
                fdata["needs_pw"] = True
                fdata["pw_error"] = ""
            needs_rerun = True
    if needs_rerun:
        st.rerun()

    # ── File status table ────────────────────────────────────────────────────
    st.markdown("#### 📋 File Queue")
    queue_df = pd.DataFrame([
        {
            "#": i + 1,
            "Order": f["order"],
            "File Name": f["name"],
            "Rows": f["df"].shape[0] if f["loaded"] else "—",
            "Cols": f["df"].shape[1] if f["loaded"] else "—",
            "Status": "✅ Loaded" if f["loaded"] else ("🔐 Needs Password" if f["needs_pw"] else "⏳ Pending"),
        }
        for i, f in enumerate(st.session_state.wl_files)
    ])

    edited_queue = st.data_editor(
        queue_df,
        column_config={
            "#":         st.column_config.NumberColumn("#", disabled=True, width="small"),
            "Order":     st.column_config.NumberColumn("Order", min_value=1, max_value=99, step=1, width="small"),
            "File Name": st.column_config.TextColumn("File Name", disabled=True),
            "Rows":      st.column_config.TextColumn("Rows", disabled=True, width="small"),
            "Cols":      st.column_config.TextColumn("Cols", disabled=True, width="small"),
            "Status":    st.column_config.TextColumn("Status", disabled=True),
        },
        num_rows="fixed",
        hide_index=True,
        use_container_width=True,
        key="wl_queue_editor",
    )
    for i, row in edited_queue.iterrows():
        if i < len(st.session_state.wl_files):
            try:
                st.session_state.wl_files[i]["order"] = int(row["Order"])
            except (ValueError, TypeError):
                pass

    # ── Decrypt section ──────────────────────────────────────────────────────
    pw_needed = [f for f in st.session_state.wl_files if f["needs_pw"] and not f["loaded"]]
    n_loaded  = sum(1 for f in st.session_state.wl_files if f["loaded"])
    n_total   = len(st.session_state.wl_files)

    if pw_needed:
        st.markdown("---")
        st.markdown("### 🔐 Decrypt Password-Protected Files")

        ms1, ms2, ms3 = st.columns(3)
        ms1.metric("Total Files",      n_total)
        ms2.metric("✅ Loaded",         n_loaded)
        ms3.metric("🔐 Need Password", len(pw_needed))

        st.markdown("**Choose your decryption method:**")
        decrypt_mode = st.radio(
            "Decryption mode:",
            options=[
                "🔑 Option 1: Same Password for All — Decrypt all locked files at once",
                "🗝️ Option 2: Different Passwords — Decrypt each file individually",
            ],
            key="wl_decrypt_mode",
            label_visibility="collapsed",
        )
        st.markdown("---")

        # ── OPTION 1: Bulk decrypt ────────────────────────────────────────
        if "Option 1" in decrypt_mode:
            st.markdown(f"Enter **one shared password** to decrypt all **{len(pw_needed)}** locked file(s).")
            for idx, f in enumerate(pw_needed, 1):
                st.markdown(f"&nbsp;&nbsp;&nbsp;{idx}. 📄 `{f['name']}`", unsafe_allow_html=True)

            bulk_pw  = st.text_input("Shared password:", type="password", key="wl_bulk_pw", placeholder="Enter password…")
            bulk_btn = st.button(
                f"🔓 Decrypt All {len(pw_needed)} File(s)",
                key="wl_bulk_btn",
                type="primary",
                use_container_width=True,
            )

            if bulk_btn:
                pw = bulk_pw.strip()
                if not pw:
                    st.warning("Please enter a password first.")
                else:
                    ok = fail = 0
                    prog = st.progress(0, text="Decrypting…")
                    for idx, fdata in enumerate(pw_needed):
                        prog.progress(idx / len(pw_needed), text=f"Decrypting {fdata['name']}…")
                        df, sheet, err = decrypt_and_read(fdata["raw"], pw, "__first__", skip_rows=skip_rows)
                        if df is not None:
                            fdata["df"], fdata["sheet"], fdata["loaded"] = df, sheet, True
                            fdata["needs_pw"] = False
                            fdata["pw_error"] = ""
                            if err: fdata["warning"] = err
                            ok += 1
                        else:
                            fdata["pw_error"] = err
                            fail += 1
                    prog.progress(1.0, text="Done!")
                    if ok:   st.success(f"✅ Decrypted {ok} file(s).")
                    if fail: st.error(f"❌ {fail} file(s) failed. Try **Option 2** for individual passwords.")
                    st.session_state.wl_step2_ready = False
                    st.session_state.wl_consolidated_df = None
                    st.rerun()

        # ── OPTION 2: Individual decrypt ──────────────────────────────────
        else:
            st.markdown("Enter a **separate password** for each locked file, then decrypt individually or all at once.")

            for fdata in pw_needed:
                tok = fdata["token"]
                c1, c2, c3 = st.columns([3, 2, 1])
                with c1:
                    st.markdown(
                        f"<div style='padding:8px 0 2px 0;font-weight:600;'>📄 {fdata['name']}</div>"
                        + (f"<div style='color:#c0392b;font-size:0.8rem;'>❌ {fdata['pw_error']}</div>" if fdata["pw_error"] else ""),
                        unsafe_allow_html=True,
                    )
                with c2:
                    st.text_input("Password:", type="password", key=f"wl_indiv_pw_{tok}",
                                  placeholder="Enter password…", label_visibility="collapsed")
                with c3:
                    st.markdown("<div style='padding-top:4px;'></div>", unsafe_allow_html=True)
                    if st.button("🔓", key=f"wl_indiv_btn_{tok}", use_container_width=True, help=f"Decrypt {fdata['name']}"):
                        pw = st.session_state.get(f"wl_indiv_pw_{tok}", "").strip()
                        if not pw:
                            st.warning(f"Enter a password for '{fdata['name']}' first.")
                        else:
                            df, sheet, err = decrypt_and_read(fdata["raw"], pw, "__first__", skip_rows=skip_rows)
                            if df is not None:
                                fdata["df"], fdata["sheet"], fdata["loaded"] = df, sheet, True
                                fdata["needs_pw"] = False
                                fdata["pw_error"] = ""
                                if err: fdata["warning"] = err
                                st.session_state.wl_step2_ready = False
                                st.session_state.wl_consolidated_df = None
                                st.rerun()
                            else:
                                fdata["pw_error"] = err
                                st.error(f"❌ Wrong password for '{fdata['name']}': {err}")

            st.markdown("---")
            if st.button("🔓 Decrypt All (using passwords entered above)", key="wl_indiv_all_btn", use_container_width=True):
                ok = fail = skipped = 0
                for fdata in pw_needed:
                    tok = fdata["token"]
                    pw  = st.session_state.get(f"wl_indiv_pw_{tok}", "").strip()
                    if not pw:
                        skipped += 1
                        continue
                    df, sheet, err = decrypt_and_read(fdata["raw"], pw, "__first__", skip_rows=skip_rows)
                    if df is not None:
                        fdata["df"], fdata["sheet"], fdata["loaded"] = df, sheet, True
                        fdata["needs_pw"] = False
                        fdata["pw_error"] = ""
                        if err: fdata["warning"] = err
                        ok += 1
                    else:
                        fdata["pw_error"] = err
                        fail += 1
                if ok:      st.success(f"✅ Decrypted {ok} file(s).")
                if fail:    st.error(f"❌ {fail} file(s) failed — check passwords above.")
                if skipped: st.warning(f"⚠️ {skipped} file(s) skipped — no password entered.")
                st.session_state.wl_step2_ready = False
                st.session_state.wl_consolidated_df = None
                st.rerun()

    # ── Loaded file previews ─────────────────────────────────────────────────
    if n_loaded > 0:
        st.markdown("---")
        st.markdown("### ✅ Loaded Files")
        for i, fdata in enumerate(st.session_state.wl_files):
            if fdata["loaded"]:
                with st.expander(f"📄 File {i+1}: {fdata['name']} — {fdata['df'].shape[0]} rows × {fdata['df'].shape[1]} cols", expanded=False):
                    if fdata.get("warning"):
                        st.warning(fdata["warning"])
                    st.dataframe(fdata["df"].head(5), use_container_width=True)

        # ── Load into Step 2 button ──────────────────────────────────────────
        pending_count = n_total - n_loaded
        if pending_count > 0:
            st.warning(f"⚠️ {pending_count} file(s) still pending. You can still proceed with {n_loaded} loaded file(s).")

        if st.button("➡️ Proceed to Column Manager", key="wl_proceed_btn", type="primary", use_container_width=True):
            # Build union of all columns in load-order
            sorted_files = sorted(
                [(i, f) for i, f in enumerate(st.session_state.wl_files) if f["loaded"]],
                key=lambda x: (x[1]["order"], x[0]),
            )
            all_cols = []
            seen = set()
            for _, fdata in sorted_files:
                for col in fdata["df"].columns:
                    if col not in seen:
                        all_cols.append(col)
                        seen.add(col)
            st.session_state.wl_all_columns     = all_cols
            st.session_state.wl_selected_columns = all_cols.copy()
            st.session_state.wl_step2_ready      = True
            st.session_state.wl_consolidated_df  = None
            st.rerun()

    # ════════════════════════════════════════════════════════════════════════
    # STEP 2 — Column Manager
    # ════════════════════════════════════════════════════════════════════════
    if not st.session_state.wl_step2_ready:
        return

    st.markdown("---")
    st.header("Step 2: Column Manager")
    st.markdown(
        f"All **{len(st.session_state.wl_all_columns)}** columns discovered across "
        f"**{n_loaded}** loaded file(s) are listed below. "
        "Remove columns you don't need, add custom ones, or reorder as required."
    )

    # ── Metrics ─────────────────────────────────────────────────────────────
    cm1, cm2, cm3 = st.columns(3)
    cm1.metric("All Discovered Columns", len(st.session_state.wl_all_columns))
    cm2.metric("Selected for Export",    len(st.session_state.wl_selected_columns))
    cm3.metric("Removed",
               len(st.session_state.wl_all_columns) - len(st.session_state.wl_selected_columns))

    # ── Quick-action buttons ─────────────────────────────────────────────────
    qa1, qa2, qa3 = st.columns(3)
    with qa1:
        if st.button("✅ Select All Columns", key="wl_sel_all", use_container_width=True):
            st.session_state.wl_selected_columns = st.session_state.wl_all_columns.copy()
            st.session_state.wl_consolidated_df  = None
            st.rerun()
    with qa2:
        if st.button("🔲 Deselect All Columns", key="wl_desel_all", use_container_width=True):
            st.session_state.wl_selected_columns = []
            st.session_state.wl_consolidated_df  = None
            st.rerun()
    with qa3:
        if st.button("🔄 Reset to Discovered Columns", key="wl_reset_cols", use_container_width=True):
            st.session_state.wl_selected_columns = st.session_state.wl_all_columns.copy()
            st.session_state.wl_consolidated_df  = None
            st.rerun()

    # ── Column selection multiselect ─────────────────────────────────────────
    # Ensure selected_columns only contains valid entries from wl_all_columns + any custom ones
    valid_existing = [c for c in st.session_state.wl_selected_columns if c in st.session_state.wl_all_columns]
    custom_cols    = [c for c in st.session_state.wl_selected_columns if c not in st.session_state.wl_all_columns]
    all_options    = st.session_state.wl_all_columns + custom_cols

    selected = st.multiselect(
        "Columns to include in export (drag to reorder in the list):",
        options=all_options,
        default=st.session_state.wl_selected_columns,
        key="wl_col_multiselect",
    )
    # Sync back
    if selected != st.session_state.wl_selected_columns:
        st.session_state.wl_selected_columns = selected
        st.session_state.wl_consolidated_df  = None

    st.markdown("---")

    # ── Column editor table (reorder + toggle) ───────────────────────────────
    st.markdown("#### 📝 Column Editor — Add, Remove & Rename")
    st.caption(
        "Toggle **Include** to remove a column from the export. "
        "Edit **Export Name** to rename the column header in the output file. "
        "Use **Add Custom Column** below to insert a brand-new blank column."
    )

    if "wl_col_editor_data" not in st.session_state or st.session_state.get("wl_col_editor_stale"):
        # Build editor rows from current selected columns
        st.session_state.wl_col_editor_data = [
            {"Include": True, "Source Column": c, "Export Name": c}
            for c in st.session_state.wl_selected_columns
        ]
        # Add discovered-but-not-selected as excluded rows
        for c in st.session_state.wl_all_columns:
            if c not in st.session_state.wl_selected_columns:
                st.session_state.wl_col_editor_data.append(
                    {"Include": False, "Source Column": c, "Export Name": c}
                )
        st.session_state.wl_col_editor_stale = False

    editor_df = pd.DataFrame(st.session_state.wl_col_editor_data)

    edited_col_df = st.data_editor(
        editor_df,
        column_config={
            "Include":       st.column_config.CheckboxColumn("Include", width="small"),
            "Source Column": st.column_config.TextColumn("Source Column", disabled=True),
            "Export Name":   st.column_config.TextColumn("Export Name"),
        },
        num_rows="fixed",
        hide_index=True,
        use_container_width=True,
        height=min(400, 50 + 35 * len(editor_df)),
        key="wl_col_editor",
    )

    # Sync editor state back
    st.session_state.wl_col_editor_data = edited_col_df.to_dict("records")
    selected_from_editor = [
        row for row in st.session_state.wl_col_editor_data if row["Include"]
    ]

    st.markdown("---")

    # ── Add custom column ────────────────────────────────────────────────────
    st.markdown("#### ➕ Add Custom Column")
    st.caption("Custom columns will be added as blank columns (empty values) in the consolidated output.")
    cc1, cc2, cc3 = st.columns([3, 1, 1])
    with cc1:
        new_col_name = st.text_input("New column name:", key="wl_new_col_input", placeholder="e.g. Verified By, Region, Remarks…")
    with cc2:
        st.markdown("<br>", unsafe_allow_html=True)
        add_col_btn = st.button("➕ Add Column", key="wl_add_col_btn", use_container_width=True)
    with cc3:
        st.markdown("<br>", unsafe_allow_html=True)
        static_fill = st.text_input("Static value (optional):", key="wl_static_val_input", placeholder="Leave blank = empty")

    if add_col_btn:
        name = new_col_name.strip()
        if not name:
            st.warning("Enter a column name first.")
        elif name in [r["Source Column"] for r in st.session_state.wl_col_editor_data]:
            st.warning(f"Column **'{name}'** already exists in the editor.")
        else:
            st.session_state.wl_col_editor_data.append({
                "Include": True,
                "Source Column": name,
                "Export Name": name,
                "_custom": True,
                "_static": static_fill.strip(),
            })
            if name not in st.session_state.wl_all_columns:
                st.session_state.wl_all_columns.append(name)
            if name not in st.session_state.wl_selected_columns:
                st.session_state.wl_selected_columns.append(name)
            st.session_state.wl_consolidated_df = None
            st.success(f"✅ Custom column **'{name}'** added.")
            st.rerun()

    st.markdown("---")

    # ── Consolidate button ───────────────────────────────────────────────────
    final_cols = [r for r in st.session_state.wl_col_editor_data if r.get("Include", False)]

    if not final_cols:
        st.warning("Select at least one column to consolidate.")
    else:
        col_summary_c1, col_summary_c2 = st.columns([3, 1])
        with col_summary_c1:
            st.info(f"**{len(final_cols)}** column(s) selected for export across **{n_loaded}** file(s).")
        with col_summary_c2:
            consolidate_btn = st.button(
                "🔀 Consolidate Now",
                key="wl_consolidate_btn",
                type="primary",
                use_container_width=True,
            )

        if consolidate_btn:
            try:
                sorted_files = sorted(
                    [(i, f) for i, f in enumerate(st.session_state.wl_files) if f["loaded"]],
                    key=lambda x: (x[1]["order"], x[0]),
                )
                frames = []
                for _, fdata in sorted_files:
                    src_df = fdata["df"].copy()
                    frame_data = {}
                    for row in final_cols:
                        src  = row["Source Column"]
                        exp  = row.get("Export Name", src) or src
                        is_custom = row.get("_custom", False)
                        static_v  = row.get("_static", "")
                        if is_custom:
                            frame_data[exp] = [static_v] * len(src_df) if static_v else [""] * len(src_df)
                        elif src in src_df.columns:
                            frame_data[exp] = src_df[src].values
                        else:
                            # Try case-insensitive
                            matched = [c for c in src_df.columns if c.strip().upper() == src.strip().upper()]
                            frame_data[exp] = src_df[matched[0]].values if matched else [""] * len(src_df)
                    frames.append(pd.DataFrame(frame_data))

                consolidated = pd.concat(frames, ignore_index=True)
                st.session_state.wl_consolidated_df = consolidated
                st.success(f"✅ Consolidation complete — **{consolidated.shape[0]}** rows × **{consolidated.shape[1]}** columns.")
            except Exception as e:
                st.error(f"❌ Consolidation error: {e}")

    # ════════════════════════════════════════════════════════════════════════
    # STEP 3 — Preview & Download
    # ════════════════════════════════════════════════════════════════════════
    if st.session_state.wl_consolidated_df is not None:
        st.markdown("---")
        st.header("Step 3: Preview & Download")

        df = st.session_state.wl_consolidated_df

        # Metrics
        pm1, pm2, pm3, pm4 = st.columns(4)
        pm1.metric("Campaign",     st.session_state.wl_campaign)
        pm2.metric("Total Rows",   df.shape[0])
        pm3.metric("Total Cols",   df.shape[1])
        filled = df.notna().sum().sum()
        pm4.metric("Fill Rate",    f"{filled / max(df.shape[0]*df.shape[1], 1)*100:.1f}%")

        # Preview with column filter
        preview_cols = st.multiselect(
            "Filter columns to display:",
            options=df.columns.tolist(),
            default=df.columns.tolist(),
            key="wl_preview_filter",
        )
        st.dataframe(df[preview_cols] if preview_cols else df, use_container_width=True, height=450)

        # Column summary
        with st.expander("📋 Column Summary"):
            csum = pd.DataFrame({
                "Column":    df.columns,
                "Non-Null":  df.count().values,
                "Null":      df.isnull().sum().values,
                "Fill %":    (df.count() / max(len(df), 1) * 100).round(1).values,
                "Sample":    [str(df[c].dropna().iloc[0])[:50] if df[c].notna().any() else "—" for c in df.columns],
            })
            st.dataframe(csum, use_container_width=True)

        # Download
        st.markdown("---")
        output = io.BytesIO()
        sheet_name = f"{st.session_state.wl_campaign} Worklist"[:31]  # Excel sheet name limit
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            ws = writer.sheets[sheet_name]
            for col_cells in ws.columns:
                max_len = max((len(str(cell.value or "")) for cell in col_cells), default=10)
                ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 3, 50)
        output.seek(0)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        campaign_slug = st.session_state.wl_campaign.replace(" ", "_")
        filename = f"Worklist_{campaign_slug}_{timestamp}.xlsx"

        dl1, dl2 = st.columns([2, 1])
        with dl1:
            st.download_button(
                label=f"⬇️ Download {st.session_state.wl_campaign} Consolidated Worklist (.xlsx)",
                data=output.getvalue(),
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )
        with dl2:
            if st.button("🔄 Start Over", key="wl_reset_btn", use_container_width=True):
                for k in list(st.session_state.keys()):
                    if k.startswith("wl_"):
                        st.session_state.pop(k, None)
                st.rerun()

        st.success(f"✅ **{filename}** ready — {df.shape[0]} rows, sheet: `{sheet_name}`")


# ============================================================================
# MODULE: UNIVERSAL CONSOLIDATOR
# ============================================================================

def render_universal_module():
    st.title("🗂️ Universal Consolidator")
    st.markdown("Upload **any** Excel or CSV files, shape the columns exactly how you want, and download a clean consolidated file.")
    st.markdown("---")

    # ── Session state init ───────────────────────────────────────────────────
    uc_defaults = [
        ("uc_files", []),
        ("uc_header_row_widget", 1),
        ("uc_header_row", 0),
        ("uc_step2_ready", False),
        ("uc_all_columns", []),
        ("uc_col_editor_data", []),
        ("uc_col_editor_stale", False),
        ("uc_consolidated_df", None),
    ]
    for k, v in uc_defaults:
        if k not in st.session_state:
            st.session_state[k] = v

    # ════════════════════════════════════════════════════════════════════════
    # STEP 1 — Upload Files
    # ════════════════════════════════════════════════════════════════════════
    st.header("Step 1: Upload Files")

    # ── Header row picker ────────────────────────────────────────────────────
    hr_c1, hr_c2 = st.columns([2, 3])
    with hr_c1:
        st.number_input(
            "Header starts at row (Excel row number):",
            min_value=1, max_value=30, step=1,
            key="uc_header_row_widget",
            help="Row 1 = very first row of the file. Change if your column headers aren't on the first row.",
        )
        new_skip = int(st.session_state.uc_header_row_widget) - 1
        if new_skip != st.session_state.uc_header_row:
            st.session_state.uc_header_row = new_skip
            for f in st.session_state.uc_files:
                f.update({"loaded": False, "probed": False, "df": None,
                           "sheet": None, "warning": None,
                           "needs_pw": False, "pw_error": ""})
            st.session_state.uc_step2_ready = False
            st.session_state.uc_consolidated_df = None
            st.session_state.uc_all_columns = []
            st.session_state.uc_col_editor_data = []
            st.rerun()
    with hr_c2:
        st.info("Default is **Row 1**. If your file has a title row above the headers, set this to **Row 2** (or higher).")

    st.markdown("---")

    # ── File uploader ────────────────────────────────────────────────────────
    st.markdown("#### 📁 Upload Files")
    st.caption("Accepts Excel (.xlsx .xls .xlsm .xlsb) and CSV files. Password-protected files will be detected automatically.")

    new_files = st.file_uploader(
        "Upload one or more files to consolidate:",
        type=["xlsx", "xls", "xlsm", "xlsb", "csv"],
        accept_multiple_files=True,
        key="uc_uploader",
    )

    # Sync queue
    existing_tokens = {f["token"] for f in st.session_state.uc_files}
    for nf in (new_files or []):
        token = f"{nf.name}:{nf.size}"
        if token not in existing_tokens:
            st.session_state.uc_files.append({
                "token": token,
                "name": nf.name,
                "order": len(st.session_state.uc_files) + 1,
                "raw": bytes(nf.getbuffer()),
                "df": None, "sheet": None, "warning": None,
                "needs_pw": False, "pw_error": "",
                "loaded": False, "probed": False,
            })
            existing_tokens.add(token)

    if new_files is not None:
        active_tokens = {f"{nf.name}:{nf.size}" for nf in new_files}
        removed = [f for f in st.session_state.uc_files if f["token"] not in active_tokens]
        st.session_state.uc_files = [f for f in st.session_state.uc_files if f["token"] in active_tokens]
        if removed:
            st.session_state.uc_step2_ready = False
            st.session_state.uc_consolidated_df = None
            st.session_state.uc_col_editor_data = []

    if not st.session_state.uc_files:
        st.info("No files uploaded yet.")
        return

    # ── Auto-probe ───────────────────────────────────────────────────────────
    skip_rows = st.session_state.uc_header_row
    needs_rerun = False
    for fdata in st.session_state.uc_files:
        if not fdata["loaded"] and not fdata["needs_pw"] and not fdata["probed"]:
            fdata["probed"] = True
            df, sheet, err = read_file_no_password(fdata["raw"], fdata["name"], "__first__", skip_rows=skip_rows)
            if df is not None:
                fdata["df"], fdata["sheet"], fdata["loaded"], fdata["warning"] = df, sheet, True, err
            else:
                fdata["needs_pw"], fdata["pw_error"] = True, ""
            needs_rerun = True
    if needs_rerun:
        st.rerun()

    # ── File queue table ─────────────────────────────────────────────────────
    st.markdown("#### 📋 File Queue")
    st.caption("Edit **Order** to control the consolidation sequence. Lower numbers go first.")

    queue_df = pd.DataFrame([
        {
            "#": i + 1,
            "Order": f["order"],
            "File Name": f["name"],
            "Sheet Loaded": f["sheet"] if f["loaded"] else "—",
            "Rows": f["df"].shape[0] if f["loaded"] else "—",
            "Cols": f["df"].shape[1] if f["loaded"] else "—",
            "Status": "✅ Loaded" if f["loaded"] else ("🔐 Needs Password" if f["needs_pw"] else "⏳ Pending"),
        }
        for i, f in enumerate(st.session_state.uc_files)
    ])

    edited_queue = st.data_editor(
        queue_df,
        column_config={
            "#":            st.column_config.NumberColumn("#", disabled=True, width="small"),
            "Order":        st.column_config.NumberColumn("Order", min_value=1, max_value=999, step=1, width="small"),
            "File Name":    st.column_config.TextColumn("File Name", disabled=True),
            "Sheet Loaded": st.column_config.TextColumn("Sheet", disabled=True, width="medium"),
            "Rows":         st.column_config.TextColumn("Rows", disabled=True, width="small"),
            "Cols":         st.column_config.TextColumn("Cols", disabled=True, width="small"),
            "Status":       st.column_config.TextColumn("Status", disabled=True),
        },
        num_rows="fixed", hide_index=True, use_container_width=True,
        key="uc_queue_editor",
    )
    for i, row in edited_queue.iterrows():
        if i < len(st.session_state.uc_files):
            try:
                st.session_state.uc_files[i]["order"] = int(row["Order"])
            except (ValueError, TypeError):
                pass

    # ── Decrypt section ──────────────────────────────────────────────────────
    pw_needed = [f for f in st.session_state.uc_files if f["needs_pw"] and not f["loaded"]]
    n_loaded  = sum(1 for f in st.session_state.uc_files if f["loaded"])
    n_total   = len(st.session_state.uc_files)

    if pw_needed:
        st.markdown("---")
        st.markdown("### 🔐 Decrypt Password-Protected Files")

        ds1, ds2, ds3 = st.columns(3)
        ds1.metric("Total Files", n_total)
        ds2.metric("✅ Loaded", n_loaded)
        ds3.metric("🔐 Need Password", len(pw_needed))

        st.markdown("**Choose decryption method:**")
        decrypt_mode = st.radio(
            "Decrypt mode",
            options=[
                "🔑 Option 1: Same Password for All",
                "🗝️ Option 2: Different Passwords — Decrypt Each File Individually",
            ],
            key="uc_decrypt_mode",
            label_visibility="collapsed",
        )
        st.markdown("---")

        # Option 1 — Bulk
        if "Option 1" in decrypt_mode:
            st.markdown(f"One shared password to unlock all **{len(pw_needed)}** locked file(s):")
            for idx, f in enumerate(pw_needed, 1):
                st.markdown(f"&nbsp;&nbsp;&nbsp;{idx}. 📄 `{f['name']}`", unsafe_allow_html=True)

            bulk_pw = st.text_input("Shared password:", type="password", key="uc_bulk_pw", placeholder="Enter password…")
            if st.button(f"🔓 Decrypt All {len(pw_needed)} File(s)", key="uc_bulk_btn", type="primary", use_container_width=True):
                pw = bulk_pw.strip()
                if not pw:
                    st.warning("Please enter a password first.")
                else:
                    ok = fail = 0
                    prog = st.progress(0, text="Decrypting…")
                    for idx, fdata in enumerate(pw_needed):
                        prog.progress(idx / len(pw_needed), text=f"Decrypting {fdata['name']}…")
                        df, sheet, err = decrypt_and_read(fdata["raw"], pw, "__first__", skip_rows=skip_rows)
                        if df is not None:
                            fdata.update({"df": df, "sheet": sheet, "loaded": True,
                                          "needs_pw": False, "pw_error": ""})
                            if err: fdata["warning"] = err
                            ok += 1
                        else:
                            fdata["pw_error"] = err
                            fail += 1
                    prog.progress(1.0, text="Done!")
                    if ok:   st.success(f"✅ Decrypted {ok} file(s).")
                    if fail: st.error(f"❌ {fail} file(s) failed — try Option 2 for individual passwords.")
                    st.session_state.uc_step2_ready = False
                    st.session_state.uc_consolidated_df = None
                    st.rerun()

        # Option 2 — Individual
        else:
            st.markdown("Enter a separate password for each locked file:")
            for fdata in pw_needed:
                tok = fdata["token"]
                c1, c2, c3 = st.columns([3, 2, 1])
                with c1:
                    st.markdown(
                        f"<div style='padding:8px 0 2px 0;font-weight:600;'>📄 {fdata['name']}</div>"
                        + (f"<div style='color:#c0392b;font-size:0.8rem;'>❌ {fdata['pw_error']}</div>" if fdata["pw_error"] else ""),
                        unsafe_allow_html=True,
                    )
                with c2:
                    st.text_input("pw", type="password", key=f"uc_indiv_pw_{tok}",
                                  placeholder="Enter password…", label_visibility="collapsed")
                with c3:
                    st.markdown("<div style='padding-top:4px;'></div>", unsafe_allow_html=True)
                    if st.button("🔓", key=f"uc_indiv_btn_{tok}", use_container_width=True, help=f"Decrypt {fdata['name']}"):
                        pw = st.session_state.get(f"uc_indiv_pw_{tok}", "").strip()
                        if not pw:
                            st.warning(f"Enter a password for '{fdata['name']}' first.")
                        else:
                            df, sheet, err = decrypt_and_read(fdata["raw"], pw, "__first__", skip_rows=skip_rows)
                            if df is not None:
                                fdata.update({"df": df, "sheet": sheet, "loaded": True,
                                              "needs_pw": False, "pw_error": ""})
                                if err: fdata["warning"] = err
                                st.session_state.uc_step2_ready = False
                                st.session_state.uc_consolidated_df = None
                                st.rerun()
                            else:
                                fdata["pw_error"] = err
                                st.error(f"❌ Wrong password for '{fdata['name']}': {err}")

            st.markdown("---")
            if st.button("🔓 Decrypt All (using passwords above)", key="uc_indiv_all_btn", use_container_width=True):
                ok = fail = skipped = 0
                for fdata in pw_needed:
                    tok = fdata["token"]
                    pw  = st.session_state.get(f"uc_indiv_pw_{tok}", "").strip()
                    if not pw: skipped += 1; continue
                    df, sheet, err = decrypt_and_read(fdata["raw"], pw, "__first__", skip_rows=skip_rows)
                    if df is not None:
                        fdata.update({"df": df, "sheet": sheet, "loaded": True,
                                      "needs_pw": False, "pw_error": ""})
                        if err: fdata["warning"] = err
                        ok += 1
                    else:
                        fdata["pw_error"] = err; fail += 1
                if ok:      st.success(f"✅ Decrypted {ok} file(s).")
                if fail:    st.error(f"❌ {fail} file(s) failed.")
                if skipped: st.warning(f"⚠️ {skipped} file(s) skipped — no password entered.")
                st.session_state.uc_step2_ready = False
                st.session_state.uc_consolidated_df = None
                st.rerun()

    # ── Loaded file previews ─────────────────────────────────────────────────
    if n_loaded > 0:
        st.markdown("---")
        st.markdown("### ✅ Loaded Files")
        for i, fdata in enumerate(st.session_state.uc_files):
            if fdata["loaded"]:
                with st.expander(
                    f"📄 {fdata['name']}  ·  sheet: `{fdata['sheet']}`  ·  "
                    f"{fdata['df'].shape[0]} rows × {fdata['df'].shape[1]} cols",
                    expanded=False,
                ):
                    if fdata.get("warning"):
                        st.warning(fdata["warning"])
                    st.dataframe(fdata["df"].head(10), use_container_width=True)

        # pending warning
        pending = n_total - n_loaded
        if pending > 0:
            st.warning(f"⚠️ {pending} file(s) still pending. You can proceed with {n_loaded} loaded file(s).")

        if st.button("➡️ Proceed to Column Editor", key="uc_proceed_btn", type="primary", use_container_width=True):
            sorted_files = sorted(
                [(i, f) for i, f in enumerate(st.session_state.uc_files) if f["loaded"]],
                key=lambda x: (x[1]["order"], x[0]),
            )
            all_cols, seen = [], set()
            for _, fdata in sorted_files:
                for col in fdata["df"].columns:
                    if col not in seen:
                        all_cols.append(col); seen.add(col)

            st.session_state.uc_all_columns  = all_cols
            st.session_state.uc_col_editor_data = [
                {"Include": True, "Order": i + 1, "Source Column": c,
                 "Export Name": c, "Static Value": ""}
                for i, c in enumerate(all_cols)
            ]
            st.session_state.uc_col_editor_stale = False
            st.session_state.uc_step2_ready      = True
            st.session_state.uc_consolidated_df  = None
            st.rerun()

    # ════════════════════════════════════════════════════════════════════════
    # STEP 2 — Column Editor
    # ════════════════════════════════════════════════════════════════════════
    if not st.session_state.uc_step2_ready:
        return

    st.markdown("---")
    st.header("Step 2: Column Editor")

    st.markdown(
        f"**{len(st.session_state.uc_all_columns)}** columns discovered across "
        f"**{n_loaded}** file(s). Configure exactly what goes into the final output below."
    )

    # ── Legend ───────────────────────────────────────────────────────────────
    st.markdown("""
    <div style='display:flex;gap:10px;flex-wrap:wrap;margin-bottom:10px;'>
      <span style='background:#1a4f8a;color:white;padding:3px 10px;border-radius:10px;font-size:0.8rem;font-weight:600;'>✅ Include — keep column</span>
      <span style='background:#555;color:white;padding:3px 10px;border-radius:10px;font-size:0.8rem;font-weight:600;'>Order — drag/type to reorder</span>
      <span style='background:#2d6a4f;color:white;padding:3px 10px;border-radius:10px;font-size:0.8rem;font-weight:600;'>Export Name — rename header in output</span>
      <span style='background:#7a4f00;color:white;padding:3px 10px;border-radius:10px;font-size:0.8rem;font-weight:600;'>Static Value — fill every row with this text</span>
    </div>
    """, unsafe_allow_html=True)

    # ── Quick actions ────────────────────────────────────────────────────────
    qa1, qa2, qa3, qa4 = st.columns(4)
    with qa1:
        if st.button("✅ Include All", key="uc_incl_all", use_container_width=True):
            for r in st.session_state.uc_col_editor_data: r["Include"] = True
            st.session_state.uc_consolidated_df = None
            st.rerun()
    with qa2:
        if st.button("🔲 Exclude All", key="uc_excl_all", use_container_width=True):
            for r in st.session_state.uc_col_editor_data: r["Include"] = False
            st.session_state.uc_consolidated_df = None
            st.rerun()
    with qa3:
        if st.button("🔄 Reset Editor", key="uc_reset_editor", use_container_width=True):
            st.session_state.uc_col_editor_data = [
                {"Include": True, "Order": i + 1, "Source Column": c,
                 "Export Name": c, "Static Value": ""}
                for i, c in enumerate(st.session_state.uc_all_columns)
            ]
            st.session_state.uc_consolidated_df = None
            st.rerun()
    with qa4:
        included_count = sum(1 for r in st.session_state.uc_col_editor_data if r.get("Include"))
        st.metric("Included Columns", f"{included_count} / {len(st.session_state.uc_col_editor_data)}")

    # ── Main column editor ───────────────────────────────────────────────────
    st.caption(
        "• **Include** — uncheck to exclude a column from the output  \n"
        "• **Order** — set the column position in the output (sorted ascending)  \n"
        "• **Export Name** — rename the header in the downloaded file  \n"
        "• **Static Value** — if filled, every row gets this fixed text (ignores source data)"
    )

    editor_df = pd.DataFrame(st.session_state.uc_col_editor_data)
    # Ensure all required columns exist (handles custom-added rows)
    for col in ["Include", "Order", "Source Column", "Export Name", "Static Value"]:
        if col not in editor_df.columns:
            editor_df[col] = "" if col in ["Export Name", "Static Value"] else (True if col == "Include" else 0)

    edited_col_df = st.data_editor(
        editor_df[["Include", "Order", "Source Column", "Export Name", "Static Value"]],
        column_config={
            "Include":       st.column_config.CheckboxColumn("Include", width="small"),
            "Order":         st.column_config.NumberColumn("Order", min_value=1, max_value=9999, step=1, width="small"),
            "Source Column": st.column_config.TextColumn("Source Column", disabled=True, width="medium"),
            "Export Name":   st.column_config.TextColumn("Export Name", width="medium"),
            "Static Value":  st.column_config.TextColumn("Static Value (fills every row)", width="large"),
        },
        num_rows="fixed",
        hide_index=True,
        use_container_width=True,
        height=min(500, 55 + 35 * len(editor_df)),
        key="uc_col_editor",
    )

    # Sync back — preserve _custom flag from existing rows
    existing_customs = {r["Source Column"]: r.get("_custom", False) for r in st.session_state.uc_col_editor_data}
    st.session_state.uc_col_editor_data = edited_col_df.to_dict("records")
    for r in st.session_state.uc_col_editor_data:
        r["_custom"] = existing_customs.get(r["Source Column"], False)

    st.markdown("---")

    # ── Add custom column ────────────────────────────────────────────────────
    st.markdown("#### ➕ Add Custom Column")
    st.caption("Adds a brand-new column to the output. Combine with **Static Value** to fill every row with fixed text.")

    acc1, acc2, acc3, acc4 = st.columns([2, 2, 1, 1])
    with acc1:
        new_col_name   = st.text_input("Column name:", key="uc_new_col_name", placeholder="e.g. Region, Agent, Source…")
    with acc2:
        new_col_static = st.text_input("Static value (optional):", key="uc_new_col_static", placeholder="Leave blank = empty cells")
    with acc3:
        st.markdown("<br>", unsafe_allow_html=True)
        add_col_btn = st.button("➕ Add", key="uc_add_col_btn", use_container_width=True)
    with acc4:
        st.markdown("<br>", unsafe_allow_html=True)
        # Position: top or bottom
        add_position = st.selectbox("Position:", options=["End", "Start"], key="uc_add_col_pos", label_visibility="collapsed")

    if add_col_btn:
        name = new_col_name.strip()
        if not name:
            st.warning("Enter a column name first.")
        elif name in [r["Source Column"] for r in st.session_state.uc_col_editor_data]:
            st.warning(f"Column **'{name}'** already exists.")
        else:
            new_order = 0 if add_position == "Start" else (
                max((r.get("Order", 0) or 0) for r in st.session_state.uc_col_editor_data) + 1
                if st.session_state.uc_col_editor_data else 1
            )
            new_row = {
                "Include": True,
                "Order": new_order,
                "Source Column": name,
                "Export Name": name,
                "Static Value": new_col_static.strip(),
                "_custom": True,
            }
            if add_position == "Start":
                st.session_state.uc_col_editor_data.insert(0, new_row)
                # Re-number orders
                for i, r in enumerate(st.session_state.uc_col_editor_data):
                    r["Order"] = i + 1
            else:
                st.session_state.uc_col_editor_data.append(new_row)

            if name not in st.session_state.uc_all_columns:
                st.session_state.uc_all_columns.append(name)
            st.session_state.uc_consolidated_df = None
            st.success(f"✅ Custom column **'{name}'** added at the {add_position.lower()}.")
            st.rerun()

    st.markdown("---")

    # ── Consolidate ──────────────────────────────────────────────────────────
    final_rows = sorted(
        [r for r in st.session_state.uc_col_editor_data if r.get("Include")],
        key=lambda r: (r.get("Order") or 0),
    )

    if not final_rows:
        st.warning("Include at least one column to consolidate.")
        return

    fc1, fc2 = st.columns([3, 1])
    with fc1:
        included_names = [r.get("Export Name") or r["Source Column"] for r in final_rows]
        st.info(
            f"**{len(final_rows)}** column(s) will be included in order: "
            + " → ".join(f"`{n}`" for n in included_names[:8])
            + (" …" if len(included_names) > 8 else "")
        )
    with fc2:
        consolidate_btn = st.button("🔀 Consolidate Now", key="uc_consolidate_btn", type="primary", use_container_width=True)

    if consolidate_btn:
        try:
            sorted_files = sorted(
                [(i, f) for i, f in enumerate(st.session_state.uc_files) if f["loaded"]],
                key=lambda x: (x[1]["order"], x[0]),
            )
            frames = []
            for _, fdata in sorted_files:
                src_df = fdata["df"].copy()
                frame_data = {}
                for row in final_rows:
                    src    = row["Source Column"]
                    exp    = (row.get("Export Name") or src).strip() or src
                    static = (row.get("Static Value") or "").strip()
                    is_custom = row.get("_custom", False)

                    if static:
                        # Static value overrides everything
                        frame_data[exp] = [static] * len(src_df)
                    elif is_custom:
                        # Custom column with no static value → blank
                        frame_data[exp] = [""] * len(src_df)
                    elif src in src_df.columns:
                        frame_data[exp] = src_df[src].values
                    else:
                        # Case-insensitive fallback
                        matched = [c for c in src_df.columns if c.strip().upper() == src.strip().upper()]
                        frame_data[exp] = src_df[matched[0]].values if matched else [""] * len(src_df)

                frames.append(pd.DataFrame(frame_data))

            consolidated = pd.concat(frames, ignore_index=True)
            st.session_state.uc_consolidated_df = consolidated
            st.success(f"✅ Done — **{consolidated.shape[0]}** rows × **{consolidated.shape[1]}** columns.")
        except Exception as e:
            st.error(f"❌ Consolidation error: {e}")

    # ════════════════════════════════════════════════════════════════════════
    # STEP 3 — Preview & Download
    # ════════════════════════════════════════════════════════════════════════
    if st.session_state.uc_consolidated_df is None:
        return

    st.markdown("---")
    st.header("Step 3: Preview & Download")

    df = st.session_state.uc_consolidated_df

    # Metrics
    pm1, pm2, pm3, pm4 = st.columns(4)
    pm1.metric("Files Merged",  n_loaded)
    pm2.metric("Total Rows",    df.shape[0])
    pm3.metric("Total Columns", df.shape[1])
    filled = df.notna().sum().sum()
    pm4.metric("Fill Rate",     f"{filled / max(df.shape[0] * df.shape[1], 1) * 100:.1f}%")

    # Preview
    st.markdown("#### 📊 Data Preview")
    prev_cols = st.multiselect(
        "Filter columns to display:",
        options=df.columns.tolist(),
        default=df.columns.tolist(),
        key="uc_preview_filter",
    )
    st.dataframe(df[prev_cols] if prev_cols else df, use_container_width=True, height=450)

    # Column summary
    with st.expander("📋 Column Summary"):
        csum = pd.DataFrame({
            "Column":   df.columns,
            "Non-Null": df.count().values,
            "Null":     df.isnull().sum().values,
            "Fill %":   (df.count() / max(len(df), 1) * 100).round(1).values,
            "Sample":   [str(df[c].dropna().iloc[0])[:60] if df[c].notna().any() else "—" for c in df.columns],
        })
        st.dataframe(csum, use_container_width=True)

    # Download
    st.markdown("---")
    st.markdown("#### 📥 Download")

    dl_fmt_col, dl_sheet_col = st.columns([1, 2])
    with dl_fmt_col:
        dl_format = st.selectbox("Output format:", options=["Excel (.xlsx)", "CSV (.csv)"], key="uc_dl_format")
    with dl_sheet_col:
        dl_sheet = st.text_input("Excel sheet name:", value="Consolidated", max_chars=31, key="uc_dl_sheet",
                                  help="Max 31 characters — Excel limit.")

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    if dl_format == "Excel (.xlsx)":
        output = io.BytesIO()
        sheet_name = (dl_sheet.strip() or "Consolidated")[:31]
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            ws = writer.sheets[sheet_name]
            for col_cells in ws.columns:
                max_len = max((len(str(cell.value or "")) for cell in col_cells), default=10)
                ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 3, 60)
        output.seek(0)
        filename = f"Consolidated_{timestamp}.xlsx"
        mime     = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        dl_data  = output.getvalue()
    else:
        csv_buf = io.StringIO()
        df.to_csv(csv_buf, index=False)
        dl_data  = csv_buf.getvalue().encode("utf-8")
        filename = f"Consolidated_{timestamp}.csv"
        mime     = "text/csv"

    dl_btn_c, reset_c = st.columns([3, 1])
    with dl_btn_c:
        st.download_button(
            label=f"⬇️ Download Consolidated File ({dl_format})",
            data=dl_data,
            file_name=filename,
            mime=mime,
            type="primary",
            use_container_width=True,
        )
    with reset_c:
        if st.button("🔄 Start Over", key="uc_reset_btn", use_container_width=True):
            for k in list(st.session_state.keys()):
                if k.startswith("uc_"):
                    st.session_state.pop(k, None)
            st.rerun()

    st.success(f"✅ **{filename}** ready — {df.shape[0]} rows, {df.shape[1]} columns.")


# ============================================================================
# ROUTER
# ============================================================================

if st.session_state.active_module == "Volare Export":
    render_volare_module()
elif st.session_state.active_module == "PTP Report":
    render_ptp_module()
elif st.session_state.active_module == "Worklist Consolidation":
    render_worklist_module()
elif st.session_state.active_module == "Universal Consolidator":
    render_universal_module()

# Footer
st.markdown("---")
st.markdown("""
<div style="text-align: center; color: gray; font-size: 0.9em;">
    <p>Data Processing Suite v1.7 — Volare Export · PTP Report · Worklist Consolidation · Universal Consolidator</p>
    <p>Built with Streamlit, Pandas, and msoffcrypto</p>
</div>
""", unsafe_allow_html=True)