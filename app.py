import streamlit as st
import pandas as pd
import numpy as np
import os

st.set_page_config(page_title="Scoring Discrepancy Resolver", layout="wide")

# --- SCENARIO EMOJI MAP ---
SCENARIO_EMOJIS = {
    '1': '🐟',
    '2': '✈️',
    '3': '🔥',
    '4': '🍸',
    '5': '💍',
    '6': '⚽'
}

# --- CATEGORY PARSING HELPERS ---
def extract_section(category_str):
    if '|' in str(category_str):
        sec = str(category_str).split('|')[0].strip()
        if sec in ['Actions for Authorities', 'Auth. Actions']:
            return 'Auth. Actions'
        if sec in ['Missing Information', 'Missing Info']:
            return 'Missing Info'
        return sec
    return str(category_str).strip()

def extract_aspect(category_str):
    if '|' in str(category_str):
        return str(category_str).split('|')[1].strip()
    return str(category_str).strip()

# --- DOCUMENT SORTING HELPER ---
def doc_sort_key(doc_str):
    s_str = str(doc_str).strip()
    if '_s' in s_str:
        parts = s_str.split('_s')
        try:
            return (int(parts[0]), int(parts[1]))
        except ValueError:
            return (9999, 9999)
    return (9999, 9999)

# --- DISCREPANCY EVALUATION HELPERS ---
def is_discrepancy_resolved(score_1, score_2):
    def parse_s(val):
        s_str = str(val).strip().lower()
        if s_str in ['nan', 'none', 'missing', 'missing data', '']:
            return np.nan
        try:
            return float(s_str)
        except ValueError:
            return np.nan

    v1 = parse_s(score_1)
    v2 = parse_s(score_2)

    if pd.isna(v1) and pd.isna(v2): return True
    if v1 == v2: return True
    if (v1 == 0.5 and v2 == 1.0) or (v1 == 1.0 and v2 == 0.5): return True

    return False

def get_discrepancy_severity(score_a, score_s):
    def parse_s(val):
        s_str = str(val).strip().lower()
        if s_str in ['nan', 'none', 'missing', 'missing data', '']:
            return np.nan
        try:
            return float(s_str)
        except ValueError:
            return np.nan

    a = parse_s(score_a)
    s = parse_s(score_s)

    if is_discrepancy_resolved(score_a, score_s):
        return None

    if (a == 1 and s == 0) or (a == 0 and s == 1):
        return "Red"
    elif (a == 0.5 and s == 0) or (a == 0 and s == 0.5):
        return "Orange"
    elif (pd.isna(a) and not pd.isna(s)) or (not pd.isna(a) and pd.isna(s)):
        return "Missing"

    return None

def get_score_color(score_str, category):
    s = str(score_str).strip()
    cat = str(category).lower()
    
    if s in ['0', '0.0']:
        return '#000000'  # Black
    
    if 'misinfo' in cat:
        if s in ['0.5', '.5']: return '#FF8C00'  # Orange
        elif s in ['1', '1.0']: return '#D9534F'  # Red
            
    elif 'correction' in cat:
        if s in ['0.5', '.5']: return '#70AD47'  # Light Green
        elif s in ['1', '1.0']: return '#1E8449'  # Deeper Green
            
    elif 'alt' in cat:
        if s in ['0.5', '.5']: return '#5DADE2'  # Lighter Sky Blue
        elif s in ['1', '1.0']: return '#1A5276'  # Darker Royal Blue

    return '#000000'

# --- DATA LOADING & SAVING ---
def load_log():
    df = pd.read_csv("Discrepancy_Log.csv", dtype=str)
    
    if 'Revision_Round' in df.columns:
        df['Revision_Round'] = pd.to_numeric(df['Revision_Round'], errors='coerce').fillna(1).astype(int)
    
    for col in ['Comment_S', 'Comment_A', 'Final_Score', 'Score_A', 'Score_S']:
        if col in df.columns:
            df[col] = df[col].replace(r'(?i)^nan$', '', regex=True).fillna('')
            
    return df

def save_log(df):
    df.to_csv("Discrepancy_Log.csv", index=False)

def normalize_score(s):
    s_str = str(s).strip()
    if s_str in ['nan', 'None', 'Missing', 'Missing Data', '']:
        return "Missing Data"
    try:
        val = float(s_str)
        if val == int(val):
            return str(int(val))
        return str(val)
    except ValueError:
        return s_str

def parse_doc(doc_name):
    doc_str = str(doc_name).strip()
    if '_s' in doc_str:
        parts = doc_str.split('_s')
        p_num = parts[0]
        s_num = parts[1]
        emoji = SCENARIO_EMOJIS.get(s_num, '')
        return f"{emoji} Participant {p_num}, Scenario {s_num}"
    return f"Document {doc_name}"

def apply_severity_sorting(df_subset, sort_order, severity_filter):
    if severity_filter != "All":
        df_subset = df_subset[df_subset['Severity'] == severity_filter]
        
    if df_subset.empty:
        return df_subset

    if sort_order == "By Severity (Red 🔴 -> Missing 🟡 -> Orange 🟠)":
        severity_rank = {'Red': 1, 'Missing': 2, 'Orange': 3}
        df_subset = df_subset.copy()
        df_subset['Rank'] = df_subset['Severity'].map(severity_rank).fillna(4)
        df_subset = df_subset.sort_values(by=['Rank', 'Revision_Round'], ascending=[True, False])
        df_subset = df_subset.drop(columns=['Rank'])
        
    return df_subset

def export_final_sheets(df_log):
    df_A = pd.read_excel("Alina_All Scenario_Codes_simple.xlsb", engine='pyxlsb')
    df_S = pd.read_excel("S_Scores.xlsx")
    
    df_A = df_A.replace(r'(?i)^\s*undefined\s*$', np.nan, regex=True)
    df_S = df_S.replace(r'(?i)^\s*undefined\s*$', np.nan, regex=True)

    for idx, row in df_log.iterrows():
        doc = row['Document']
        cat = row['Category']
        final_val = row['Final_Score'] if pd.notna(row['Final_Score']) and str(row['Final_Score']).strip() != "" else row['Score_A']
        
        val_to_insert = np.nan if str(final_val).strip() in ["Missing", "Missing Data", "nan", ""] else float(final_val)

        if doc in df_A['Document name'].values and cat in df_A.columns:
            df_A.loc[df_A['Document name'] == doc, cat] = val_to_insert
            
        if doc in df_S['Document name'].values and cat in df_S.columns:
            df_S.loc[df_S['Document name'] == doc, cat] = val_to_insert

    df_Avg = df_A.copy()
    numeric_cols = [c for c in df_A.columns if c != 'Document name']
    
    for col in numeric_cols:
        if col in df_S.columns:
            s_a = pd.to_numeric(df_A[col], errors='coerce')
            s_s = pd.to_numeric(df_S[col], errors='coerce')
            df_Avg[col] = pd.concat([s_a, s_s], axis=1).mean(axis=1, skipna=True)

    output_filename = "Final_Scoring_Outputs.xlsx"
    with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
        df_A.to_excel(writer, sheet_name="A_Final_Scores", index=False)
        df_S.to_excel(writer, sheet_name="S_Final_Scores", index=False)
        df_Avg.to_excel(writer, sheet_name="Averages", index=False)

    return output_filename

# --- APP LAYOUT ---
st.title("Transcript Scoring Discrepancy Resolver")

if not os.path.exists("Discrepancy_Log.csv"):
    st.error("Discrepancy_Log.csv not found. Please run prep_data.py first!")
    st.stop()

df = load_log()
score_options = ["0", "0.5", "1", "Missing Data"]

# --- SIDEBAR CONTROLS ---
st.sidebar.header("User Login")
user = st.sidebar.selectbox("Logging in as:", ["Select...", "S (Second Scorer)", "A (Lead Scorer)"])

# PASSWORD GUARD FOR LEAD SCORER
if user == "A (Lead Scorer)":
    pwd = st.sidebar.text_input("Enter Lead Scorer Password:", type="password", key="lead_pwd")
    if pwd != "alina":
        if pwd != "":
            st.sidebar.error("❌ Incorrect Password")
        st.info("🔒 **Password Protected:** Please enter the Lead Scorer password in the left sidebar to unlock this area.")
        st.stop()

st.sidebar.markdown("---")
st.sidebar.header("Presentation Settings")

if user == "S (Second Scorer)":
    active_pending = df[df['Status'] == 'Pending S']
elif user == "A (Lead Scorer)":
    active_pending = df[df['Status'] == 'Pending A']
else:
    active_pending = df[df['Status'].isin(['Pending S', 'Pending A'])]

cnt_all = len(active_pending)
cnt_red = len(active_pending[active_pending['Severity'] == 'Red'])
cnt_missing = len(active_pending[active_pending['Severity'] == 'Missing'])
cnt_orange = len(active_pending[active_pending['Severity'] == 'Orange'])

sort_order = st.sidebar.radio(
    "Discrepancy Order:",
    ["By Severity (Red 🔴 -> Missing 🟡 -> Orange 🟠)", "Original File Order"]
)

selected_filter_label = st.sidebar.selectbox(
    "Filter by Severity:",
    [
        f"All ({cnt_all})",
        f"Red 🔴 ({cnt_red})",
        f"Missing 🟡 ({cnt_missing})",
        f"Orange 🟠 ({cnt_orange})"
    ]
)

if "Red" in selected_filter_label: severity_filter = "Red"
elif "Missing" in selected_filter_label: severity_filter = "Missing"
elif "Orange" in selected_filter_label: severity_filter = "Orange"
else: severity_filter = "All"

# ---------------------------------------------------------
# INITIAL PROMPT BANNER
# ---------------------------------------------------------
if user == "Select...":
    st.info("👈 **Welcome!** Please select your user profile (**S** or **A**) from the **User Login** dropdown menu in the sidebar on the left to begin.")

# ---------------------------------------------------------
# VIEW FOR S (SECOND SCORER)
# ---------------------------------------------------------
elif user == "S (Second Scorer)":
    pending_s_raw = df[df['Status'] == 'Pending S']
    pending_s = apply_severity_sorting(pending_s_raw, sort_order, severity_filter)
    
    if pending_s_raw.empty:
        st.balloons()
        st.success("🎉 You have resolved all of your discrepancies!")
    elif pending_s.empty:
        st.info(f"No discrepancies found matching the filter `{severity_filter}`.")
    else:
        st.info(f"📋 You have **{len(pending_s_raw)}** total individual discrepancies left to resolve.")
        
        # Identify top document and section for grouping
        top_row = pending_s.iloc[0]
        target_doc = top_row['Document']
        target_sec = extract_section(top_row['Category'])
        
        # Extract ALL items in pending_s sharing same Document and Section
        group_items = pending_s[
            (pending_s['Document'] == target_doc) & 
            (pending_s['Category'].apply(extract_section) == target_sec)
        ]
        
        parsed_doc = parse_doc(target_doc)
        st.markdown(f"## {parsed_doc} — Section: `{target_sec}`")
        if len(group_items) > 1:
            st.caption(f"Showing **{len(group_items)}** discrepancies for this section together.")

        with st.form("form_s_group"):
            inputs = {}
            for idx, row in group_items.iterrows():
                aspect_name = extract_aspect(row['Category'])
                category_name = str(row['Category']).strip()
                
                st.markdown(f"### Aspect: `{aspect_name}`")
                
                if row['Severity'] == 'Red':
                    st.error("🔴 **Extreme Discrepancy** (1 vs 0)")
                elif row['Severity'] == 'Orange':
                    st.warning("🟠 **Mild Discrepancy** (0.5 vs 0)")
                else:
                    st.info("🟡 **Missing Data Discrepancy**")

                if int(row['Revision_Round']) > 1:
                    st.warning(f"⚠️ **Revision Round {row['Revision_Round']}** — Sent back by A")

                score_a_formatted = normalize_score(row['Score_A'])
                color_a = get_score_color(score_a_formatted, category_name)
                
                st.markdown(
                    f"**A gave a {category_name} of:** <span style='font-size: 22px; font-weight: bold; color: {color_a};'>{score_a_formatted}</span>", 
                    unsafe_allow_html=True
                )
                
                if int(row['Revision_Round']) > 1 and str(row['Comment_A']).strip() != "":
                    st.info(f"**A's comment:** {row['Comment_A']}")

                orig_s = normalize_score(row['Score_S'])
                if orig_s == "Missing Data":
                    st.markdown(
                        f"**You gave a {category_name} of:** <span style='font-size: 22px; font-weight: bold; color: #000000;'>Missing Data</span> *(You did not assign a score here previously)*", 
                        unsafe_allow_html=True
                    )
                else:
                    color_s = get_score_color(orig_s, category_name)
                    st.markdown(
                        f"**You gave a {category_name} of:** <span style='font-size: 22px; font-weight: bold; color: {color_s};'>{orig_s}</span>", 
                        unsafe_allow_html=True
                    )

                new_score = st.radio(f"Select Score for {aspect_name}:", score_options, index=None, key=f"radio_s_{row['ID']}")
                comment = st.text_area(f"Reasoning / Comment for {aspect_name} (Required if keeping discrepancy):", key=f"comment_s_{row['ID']}")
                
                inputs[idx] = {
                    'row': row,
                    'new_score': new_score,
                    'comment': comment
                }
                st.write("---")
            
            if st.form_submit_button("Submit Section Discrepancies"):
                has_error = False
                for idx, data in inputs.items():
                    row = data['row']
                    n_score = data['new_score']
                    c_text = data['comment']
                    aspect_title = extract_aspect(row['Category'])
                    
                    if n_score is None:
                        st.error(f"⚠️ Please select a score option for `{aspect_title}` before submitting.")
                        has_error = True
                        break
                    elif not is_discrepancy_resolved(n_score, row['Score_A']) and not str(c_text).strip():
                        st.error(f"⚠️ Reasoning / Comment is required for `{aspect_title}` when keeping a score discrepancy.")
                        has_error = True
                        break
                
                if not has_error:
                    for idx, data in inputs.items():
                        row = data['row']
                        n_score = data['new_score']
                        c_text = data['comment']
                        
                        df.at[idx, 'Score_S'] = str(n_score)
                        df.at[idx, 'Comment_S'] = str(c_text)
                        
                        if is_discrepancy_resolved(n_score, row['Score_A']):
                            df.at[idx, 'Status'] = 'Resolved'
                            df.at[idx, 'Final_Score'] = str(n_score)
                        else:
                            new_sev = get_discrepancy_severity(row['Score_A'], n_score)
                            if new_sev: df.at[idx, 'Severity'] = new_sev
                            df.at[idx, 'Status'] = 'Pending A'
                            
                    save_log(df)
                    st.rerun()

# ---------------------------------------------------------
# VIEW FOR A (LEAD SCORER)
# ---------------------------------------------------------
elif user == "A (Lead Scorer)":
    tab1, tab2, tab3, tab4 = st.tabs(["Pending Review", "Reassess Discrepancies", "Resolved Discrepancies", "Export Final Sheets"])
    
    with tab1:
        pending_a_raw = df[df['Status'] == 'Pending A']
        pending_a = apply_severity_sorting(pending_a_raw, sort_order, severity_filter)
        
        if pending_a_raw.empty:
            st.success("🎉 You have reviewed all pending discrepancies!")
        elif pending_a.empty:
            st.info(f"No pending items matching filter `{severity_filter}`.")
        else:
            st.info(f"📋 You have **{len(pending_a_raw)}** total individual discrepancies left to review.")
            
            top_row = pending_a.iloc[0]
            target_doc = top_row['Document']
            target_sec = extract_section(top_row['Category'])
            
            group_items = pending_a[
                (pending_a['Document'] == target_doc) & 
                (pending_a['Category'].apply(extract_section) == target_sec)
            ]
            
            parsed_doc = parse_doc(target_doc)
            st.markdown(f"## {parsed_doc} — Section: `{target_sec}`")
            if len(group_items) > 1:
                st.caption(f"Showing **{len(group_items)}** discrepancies for this section together.")

            with st.form("form_a_group"):
                inputs = {}
                for idx, row in group_items.iterrows():
                    aspect_name = extract_aspect(row['Category'])
                    category_name = str(row['Category']).strip()
                    
                    st.markdown(f"### Aspect: `{aspect_name}`")
                    
                    if row['Severity'] == 'Red':
                        st.error("🔴 **Extreme Discrepancy** (1 vs 0)")
                    elif row['Severity'] == 'Orange':
                        st.warning("🟠 **Mild Discrepancy** (0.5 vs 0)")
                    else:
                        st.info("🟡 **Missing Data Discrepancy**")

                    if int(row['Revision_Round']) > 1:
                        st.warning(f"⚠️ **Revision Round {row['Revision_Round']}**")

                    score_a_formatted = normalize_score(row['Score_A'])
                    score_s_formatted = normalize_score(row['Score_S'])

                    color_a = get_score_color(score_a_formatted, category_name)
                    color_s = get_score_color(score_s_formatted, category_name)

                    st.markdown(
                        f"**You gave a {category_name} of:** <span style='font-size: 22px; font-weight: bold; color: {color_a};'>{score_a_formatted}</span>", 
                        unsafe_allow_html=True
                    )
                    st.markdown(
                        f"**S gave a {category_name} of:** <span style='font-size: 22px; font-weight: bold; color: {color_s};'>{score_s_formatted}</span>", 
                        unsafe_allow_html=True
                    )
                    
                    if str(row['Comment_S']).strip() != "":
                        st.info(f"**S's reasoning:** {row['Comment_S']}")

                    new_score = st.radio(f"Final Score Choice for {aspect_name}:", score_options, index=None, key=f"radio_a_{row['ID']}")
                    comment = st.text_area(f"Reasoning / Justification for {aspect_name} (Required if sending back):", key=f"comment_a_{row['ID']}")
                    
                    inputs[idx] = {
                        'row': row,
                        'new_score': new_score,
                        'comment': comment
                    }
                    st.write("---")
                    
                col1, col2 = st.columns(2)
                with col1:
                    accept = st.form_submit_button(" Accept & Finalize Section Discrepancies")
                with col2:
                    send_back = st.form_submit_button(" Submit & Send Back to S")
                    
                if accept:
                    has_error = False
                    for idx, data in inputs.items():
                        if data['new_score'] is None:
                            aspect_title = extract_aspect(data['row']['Category'])
                            st.error(f"⚠️ Please select a final score option for `{aspect_title}` before accepting.")
                            has_error = True
                            break
                    if not has_error:
                        for idx, data in inputs.items():
                            df.at[idx, 'Score_A'] = str(data['new_score'])
                            df.at[idx, 'Final_Score'] = str(data['new_score'])
                            df.at[idx, 'Status'] = 'Resolved'
                        save_log(df)
                        st.rerun()

                elif send_back:
                    has_error = False
                    for idx, data in inputs.items():
                        aspect_title = extract_aspect(data['row']['Category'])
                        if data['new_score'] is None:
                            st.error(f"⚠️ Please select a score option for `{aspect_title}` before sending back.")
                            has_error = True
                            break
                        elif not str(data['comment']).strip():
                            st.error(f"⚠️ Reasoning / Justification is required for `{aspect_title}` when sending back to S.")
                            has_error = True
                            break
                    if not has_error:
                        for idx, data in inputs.items():
                            df.at[idx, 'Score_A'] = str(data['new_score'])
                            df.at[idx, 'Comment_A'] = str(data['comment'])
                            df.at[idx, 'Revision_Round'] = int(df.at[idx, 'Revision_Round']) + 1
                            df.at[idx, 'Status'] = 'Pending S'
                        save_log(df)
                        st.rerun()

    # --- TAB 2: REASSESS DISCREPANCIES ---
    with tab2:
        st.header("Reassess Discrepancies")
        st.write("Select any active discrepancy below (e.g. Orange mild discrepancies) to reassess or update your score.")
        
        reassess_filter = st.selectbox("Filter items to reassess:", ["Orange 🟠 Mild Discrepancies", "All Active Unresolved Discrepancies"])
        
        if reassess_filter == "Orange 🟠 Mild Discrepancies":
            unresolved = df[(df['Status'] != 'Resolved') & (df['Severity'] == 'Orange')]
        else:
            unresolved = df[df['Status'] != 'Resolved']
            
        if unresolved.empty:
            st.success("No discrepancies found matching this filter!")
        else:
            unresolved = unresolved.copy()
            unresolved['sort_p'] = unresolved['Document'].apply(lambda x: doc_sort_key(x)[0])
            unresolved['sort_s'] = unresolved['Document'].apply(lambda x: doc_sort_key(x)[1])
            unresolved = unresolved.sort_values(by=['sort_p', 'sort_s', 'Category'])

            item_options = {
                f"{parse_doc(r['Document'])} | {r['Category']} (A: {r['Score_A']} | S: {r['Score_S']})": r['ID']
                for _, r in unresolved.iterrows()
            }
            
            selected_item_label = st.selectbox("Choose a discrepancy to review:", list(item_options.keys()))
            selected_id = item_options[selected_item_label]
            
            re_idx = df[df['ID'] == selected_id].index[0]
            re_row = df.loc[re_idx]
            
            st.markdown("---")
            st.markdown(f"### {parse_doc(re_row['Document'])} — Category: `{re_row['Category']}`")
            
            score_a_formatted = normalize_score(re_row['Score_A'])
            score_s_formatted = normalize_score(re_row['Score_S'])
            cat_name = str(re_row['Category']).strip()
            
            st.markdown(f"**Your current score:** `{score_a_formatted}`")
            st.markdown(f"**S's current score:** `{score_s_formatted}`")
            
            curr_a = score_a_formatted
            default_index = score_options.index(curr_a) if curr_a in score_options else 0
            
            with st.form("form_reassess"):
                new_a_score = st.radio("Update Your Score Choice:", score_options, index=default_index, key=f"reassess_radio_{re_row['ID']}")
                
                c1, c2 = st.columns(2)
                with c1:
                    keep_as_is = st.form_submit_button(" Keep Score As Is (Send to S for Round 1)")
                with c2:
                    update_score = st.form_submit_button(" Update Score & Re-evaluate")
                    
                if keep_as_is:
                    df.at[re_idx, 'Status'] = 'Pending S'
                    df.at[re_idx, 'Revision_Round'] = 1
                    save_log(df)
                    st.success("Discrepancy kept as is and set to Pending S (Round 1).")
                    st.rerun()
                    
                elif update_score:
                    df.at[re_idx, 'Score_A'] = str(new_a_score)
                    
                    if is_discrepancy_resolved(new_a_score, re_row['Score_S']):
                        df.at[re_idx, 'Status'] = 'Resolved'
                        df.at[re_idx, 'Final_Score'] = str(new_a_score)
                        st.success("Your updated score matched/resolved against S! Item marked as Resolved.")
                    else:
                        new_sev = get_discrepancy_severity(new_a_score, re_row['Score_S'])
                        if new_sev: df.at[re_idx, 'Severity'] = new_sev
                        df.at[re_idx, 'Status'] = 'Pending S'
                        df.at[re_idx, 'Revision_Round'] = 1
                        st.info("Score updated! Sent to S for Round 1 review.")
                        
                    save_log(df)
                    st.rerun()

    with tab3:
        resolved = df[df['Status'] == 'Resolved']
        st.write(f"Total Resolved: **{len(resolved)}**")
        
        for idx, row in resolved.iterrows():
            parsed_doc = parse_doc(row['Document'])
            with st.expander(f"{parsed_doc} | {row['Category']} — Final Score: {row['Final_Score']}"):
                st.write(f"A's Score: `{row['Score_A']}` | S's Score: `{row['Score_S']}`")
                if str(row['Comment_S']).strip() != "": st.write(f"S's Comment: {row['Comment_S']}")
                if str(row['Comment_A']).strip() != "": st.write(f"A's Comment: {row['Comment_A']}")
                
                if st.button("Reopen Discrepancy", key=f"reopen_{idx}"):
                    df.at[idx, 'Status'] = 'Pending A'
                    save_log(df)
                    st.rerun()

    with tab4:
        st.header("Generate Final Excel Sheets")
        st.write("Click below to generate the 3 final sheets:")
        st.write("1. **A's Final Scores**")
        st.write("2. **S's Final Scores**")
        st.write("3. **Averages of A and S**")
        
        if st.button("Generate & Download Excel File"):
            filename = export_final_sheets(df)
            with open(filename, "rb") as file:
                st.download_button(
                    label="📥 Download Final_Scoring_Outputs.xlsx",
                    data=file,
                    file_name="Final_Scoring_Outputs.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )