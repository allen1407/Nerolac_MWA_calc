# ── Metadata ──────────────────────────────────────────────────────────────────
# Author      : Allen Peter
# Developed   : 16-06-2026
# Description : MWA Calculation Report — processes FA input file and generates
#               a 4-sheet Excel output: MWA Base, MWA Summary, FA Upload Summary,
#               FA Upload Base.
# ─────────────────────────────────────────────────────────────────────────────

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.set_page_config(page_title="MWA Calculator", layout="centered")
# Run on port 8503: streamlit run mwa_calculator.py --server.port 8503
st.title("📊 MWA Calculation Report")
st.markdown("Upload the **FA File** to generate the MWA Calculation Report.")
st.divider()

fa_file = st.file_uploader("Upload FA File", type=["xlsx"], key="fa")
st.divider()

# ── Static Maps ───────────────────────────────────────────────────────────────

MWA_BASE_COLS = [
    'Business','Zone','Division','Depot Code','Sales Group','KNPL_ID','DOJ',
    'VISIT_ID','MONTH','PLANNED','CREATION_DATE','ACTUAL_DATE','SYSTEM_DAY',
    'VISIT_CATEGORY_DESCRIPTION','ACTIVITY_TYPE','ACTIVITY_TYPE_DESCRIPTION',
    'EMPLOYEE_RESPONSIBLE','EMPLOYEE_LAST_NAME','EMPLOYEE_LAST_NAME_2',
    'BUSINESS_PARTNER_ROLE','BP_ROLE_DESCRIPTION','PROSPECT_CUST_ID',
    'PROSPECT_CUSTOMER_NAME','Business Category','Business Category Description',
    'Category','ZSTATUS','DATE_OF_ENTRY_IN_TABLE','SGRP','KNPL_ID','KNPL_NAME',
    'USER_NAME','GPS_IN_TIME','GPS_OUT_TIME','ACTUAL_IN_TIME','ACTUAL_OUT_TIME',
    'LAT','LONGT','UPDATEDDATE','User Group'
]

MIN_TARGET_CAT1 = {
    'Adhesive': 30, 'CC': 30, 'CCCR': 90, 'CCTA': 90,
    'Distribution': 85, 'Project': 80, 'Retail': 40,
    'TMO': 45, 'TMO - AID': 35, 'Wood Coating': 30,
    'Soldier': 40, 'Sales Trainee': 40, 'NDO': 50,
}

MIN_TARGET_CAT2 = {
    'Adhesive': 10, 'CC': 10, 'CCCR': 0, 'CCTA': 0,
    'Distribution': 5, 'Project': 30, 'Retail': 10,
    'TMO': 0, 'TMO - AID': 0, 'Wood Coating': 10,
    'Soldier': 10, 'Sales Trainee': 10, 'NDO': 0,
}

# Slab thresholds: {business: [(cat1_thresh, cat2_thresh, amount), ...]} sorted slab1..slab_n
SLAB_MAP = {
    'Adhesive':     [(30,10,2000),(40,10,2500),(40,20,3000),(45,25,3500),(50,30,4000)],
    'CC':           [(30,10,2000),(50,20,4500)],
    'Distribution': [(85,5,2000),(110,10,4500)],
    'Project':      [(80,35,2000),(110,50,4500)],
    'Retail':       [(40,10,2000),(50,20,4500)],
    'Sales Trainee':[(40,10,2000),(50,20,4500)],
    'Soldier':      [(40,10,2000),(50,20,4500)],
    'Wood Coating': [(30,10,2000),(40,20,4500)],
    'TMO':          [(45,None,2000),(65,None,4500)],
    'TMO - AID':    [(35,None,2000),(55,None,4500)],
    'CCCR':         [(90,None,2000),(120,None,4500)],
    'CCTA':         [(90,None,2000),(120,None,4500)],
    'NDO':          [(50,None,2000),(70,None,4500)],
}

# Category rule table: Business -> (cat1_descriptions, cat2_descriptions, mode)
# mode='explicit' -> Category 1 if Description in cat1_set, Category 2 if in cat2_set, else Not Assigned
# mode='cat1_only' -> Category 1 if Description in cat1_set, else Not Assigned (no Category 2 exists)
# mode='invert' -> Category 2 if Description in cat2_set, else Category 1 (used for Project)
CATEGORY_RULES = {
    'Adhesive':      ({'SAP Dealer','Dealer'}, {'Contractors','Architect','Customer','Site','Painter','OEM','Consumer'}, 'explicit'),
    'Retail':        ({'SAP Dealer','Dealer'}, {'Contractors','Architect','Customer','Site','Painter','OEM','Consumer'}, 'explicit'),
    'Sales Trainee': ({'SAP Dealer','Dealer'}, {'Contractors','Architect','Customer','Site','Painter','OEM','Consumer'}, 'explicit'),
    'Soldier':       ({'SAP Dealer','Dealer'}, {'Contractors','Architect','Customer','Site','Painter','OEM','Consumer'}, 'explicit'),
    'Wood Coating':  ({'SAP Dealer','Dealer'}, {'Painter','Contractors','Architect','Customer','OEM','Site','Consumer'}, 'explicit'),
    'CC':            ({'SAP Dealer','Dealer'}, {'Painter','Contractors','Architect','Customer','OEM','Site','Consumer'}, 'explicit'),
    'NDO':           ({'SAP Dealer','Dealer','Architect'}, set(), 'cat1_only'),
    'CCCR':          ({'SAP Dealer','Dealer','Painter','Contractors','Site','Customer','Consumer'}, set(), 'cat1_only'),
    'CCTA':          ({'SAP Dealer','Dealer','Painter','Contractors','Site','Customer','Consumer'}, set(), 'cat1_only'),
    'Distribution':  ({'SAP Dealer','Dealer'}, {'Painter','Contractors','Architect'}, 'explicit'),
    'TMO':           ({'SAP Dealer','Dealer','Painter','Contractors','Architect','Customer','OEM','Site','Others','Consumer'}, set(), 'cat1_only'),
    'TMO - AID':     ({'Painter','Contractors','Architect','OEM','Site'}, set(), 'cat1_only'),
    'Project':       (set(), {'Consumer','Site','CHS/RWA'}, 'invert'),
}

def lookup_category(business, business_category_description):
    biz = str(business).strip()
    desc = str(business_category_description).strip()
    rule = CATEGORY_RULES.get(biz)
    if rule is None:
        return 'Not Assigned'
    cat1_set, cat2_set, mode = rule
    if mode == 'explicit':
        if desc in cat1_set:
            return 'Category 1'
        if desc in cat2_set:
            return 'Category 2'
        return 'Not Assigned'
    if mode == 'cat1_only':
        if desc in cat1_set:
            return 'Category 1'
        return 'Not Assigned'
    if mode == 'invert':
        if desc in cat2_set:
            return 'Category 2'
        return 'Category 1'
    return 'Not Assigned'

# Special PROSPECT_CUST_IDs that always get "Common Code" as description
COMMON_CODE_IDS = {
    '55865','55866','55840','59249','57911','55868','55830',
    '55869','6033621','6033622','6033697','55846','55867','55832'
}

# EMPLOYEE_LAST_NAME_2 code -> Business Category Description (BUP002 branch)
BCD_CODE_MAP = {
    'Z001': 'Contractors', 'Z002': 'Architect',   'Z003': 'Architect',
    'Z004': 'Customer',    'Z005': 'Dealer',       'Z006': 'Contractors',
    'Z007': 'Contractors', 'Z008': 'Contractors',  'Z009': 'Contractors',
    'Z010': 'Dealer',      'Z011': 'Dealer',       'Z012': 'Dealer',
    'Z013': 'Dealer',      'Z019': 'OEM',          'Z020': 'Consumer',
    'Z021': 'Painter',     'Z022': 'Others',       'Z023': 'Customer',
    'Z026': 'Dealer',      'Z027': 'Site',         'Z032': 'Dealer',
    'Z033': 'Dealer',      'Z041': 'Dealer',       'Z042': 'Dealer',
    'Z043': 'Dealer',      'Z044': 'OEM',          'Z046': 'Dealer',
    'Z047': 'Site',
}

# ── Helpers ───────────────────────────────────────────────────────────────────

def derive_business_category(bp_role):
    role = str(bp_role).strip()
    if role == 'CRM000':
        return 'SAP Dealer'
    if role == 'BUP002':
        return None  # caller fills with EMPLOYEE_LAST_NAME_2
    return 'N.A'

def derive_business_category_and_description(prospect_cust_id, bp_role, emp_last_name_2):
    pid = str(prospect_cust_id).strip()
    # Strip leading zeros / trailing .0 for numeric-looking IDs to match the common-code list robustly
    pid_norm = pid.lstrip('0') or '0'
    pid_norm = pid_norm.replace('.0', '') if pid_norm.endswith('.0') else pid_norm
    if pid in COMMON_CODE_IDS or pid_norm in COMMON_CODE_IDS:
        # Business Category itself still follows the role-based rule; only description is overridden
        role = str(bp_role).strip()
        if role == 'CRM000':
            biz_cat = 'SAP Dealer'
        elif role == 'BUP002':
            biz_cat = str(emp_last_name_2).strip()
        else:
            biz_cat = 'N.A'
        return biz_cat, 'Common Code'

    role = str(bp_role).strip()
    if role == 'CRM000':
        return 'SAP Dealer', 'SAP Dealer'
    if role == 'BUP002':
        code = str(emp_last_name_2).strip()
        desc = BCD_CODE_MAP.get(code, 'Not Assigned')
        return code, desc
    return 'N.A', 'Not Assigned'

def get_slab_and_amount(business, cat1_valid, cat2_valid):
    slabs = SLAB_MAP.get(business, [])
    if not slabs:
        return 'Not Eligible', 0
    achieved = 'Not Eligible'
    achieved_amount = 0
    for i, slab in enumerate(slabs):
        c1, c2, amt = slab
        if c2 is None:
            if cat1_valid >= c1:
                achieved = f'Slab {i+1}'
                achieved_amount = amt
        else:
            if cat1_valid >= c1 and cat2_valid >= c2:
                achieved = f'Slab {i+1}'
                achieved_amount = amt
    return achieved, achieved_amount

def get_next_slab_cat1(business, cat1_valid):
    slabs = SLAB_MAP.get(business, [])
    for slab in slabs:
        c1 = slab[0]
        if cat1_valid < c1:
            return c1
    return 0

def get_next_slab_cat2(business, cat2_valid):
    slabs = SLAB_MAP.get(business, [])
    for slab in slabs:
        c2 = slab[1]
        if c2 is None:
            continue
        if cat2_valid < c2:
            return c2
    return 0

def _to_amount(value):
    """Safely coerce a value (including numpy int64/float64 from DataFrame columns)
    into a plain Python number. Returns 0 if it isn't a real number."""
    if isinstance(value, bool):
        return 0
    if isinstance(value, (int, float, np.integer, np.floating)):
        if pd.isna(value):
            return 0
        return float(value) if isinstance(value, (float, np.floating)) else int(value)
    return 0

def get_next_slab_amount(business, current_amount):
    slabs = SLAB_MAP.get(business, [])
    amounts = [s[2] for s in slabs]
    amt_val = _to_amount(current_amount)
    for amt in amounts:
        if amt_val < amt:
            return amt
    return 0

def format_visit_month(month_val, creation_date_val):
    try:
        m = str(int(float(str(month_val)))).zfill(2)
    except:
        m = '00'
    try:
        parts = str(creation_date_val).split('.')
        year = parts[2] if len(parts) >= 3 else ''
    except:
        year = ''
    return f"{m}.{year}"

# ── Core Processing ───────────────────────────────────────────────────────────

def process_mwa(fa: pd.DataFrame) -> dict:

    fa = fa.copy()

    # Rule 1: Eliminate rows where Sales Group is "Not Assigned"
    fa = fa[fa['Sales Group'].astype(str).str.strip() != 'Not Assigned'].copy()

    # Rule 2: Eliminate rows where PROSPECT_CUST_ID is in the special Common Code list
    def _pid_norm(x):
        s = str(x).strip()
        s = s.lstrip('0') or '0'
        if s.endswith('.0'):
            s = s[:-2]
        return s
    fa = fa[~fa['PROSPECT_CUST_ID'].apply(_pid_norm).isin(COMMON_CODE_IDS)].copy()

    # Rule 3: Eliminate rows where EMPLOYEE_LAST_NAME_2 is not in format "Z" + 3 digits (e.g. Z023)
    import re as _re
    Z_CODE_RE = _re.compile(r'^Z\d{3}$')
    fa = fa[fa['EMPLOYEE_LAST_NAME_2'].astype(str).str.strip().str.match(Z_CODE_RE)].copy()

    # Zone: blank/null -> "N.A"
    if 'Zone' in fa.columns:
        fa['Zone'] = fa['Zone'].apply(
            lambda x: 'N.A' if pd.isna(x) or str(x).strip() == '' else x)

    # Derive Business Category and Business Category Description
    bc_results = fa.apply(
        lambda r: derive_business_category_and_description(
            r.get('PROSPECT_CUST_ID',''), r.get('BUSINESS_PARTNER_ROLE',''), r.get('EMPLOYEE_LAST_NAME_2','')
        ), axis=1)
    fa['Business Category'] = bc_results.apply(lambda t: t[0])
    fa['Business Category Description'] = bc_results.apply(lambda t: t[1])

    # Derive Category via rule table (Business + Business Category Description)
    fa['Category'] = fa.apply(
        lambda r: lookup_category(r.get('Business',''), r.get('Business Category Description','')), axis=1)

    # ── Sheet 1: MWA Base ─────────────────────────────────────────────────────
    available = [c for c in MWA_BASE_COLS if c in fa.columns]
    # KNPL_ID appears twice — handle duplicates
    # Build list allowing duplicates by position

    base_cols = []
    seen = {}
    fa_cols_list = list(fa.columns)
    for col in MWA_BASE_COLS:
        if col in fa.columns:
            if col not in seen:
                seen[col] = 0
            else:
                seen[col] += 1
            base_cols.append(col)
        # if col not present just skip
    mwa_base = fa[base_cols].copy() if base_cols else fa.copy()
    # Drop any extra cols not in MWA_BASE_COLS
    # Already handled by selecting only base_cols

    # Format ACTUAL_DATE as DD-MM-YYYY
    if 'ACTUAL_DATE' in mwa_base.columns:
        mwa_base['ACTUAL_DATE'] = pd.to_datetime(
            mwa_base['ACTUAL_DATE'], errors='coerce', dayfirst=True
        ).dt.strftime('%d-%m-%Y')

    # Fix #5: GPS_IN_TIME / GPS_OUT_TIME / ACTUAL_IN_TIME / ACTUAL_OUT_TIME — if filled,
    # strip the Excel placeholder date "12/31/1899" so only the time portion remains.
    # If blank, leave blank.
    def clean_time_field(x):
        if pd.isna(x):
            return x
        s = str(x)
        if s.strip() == '':
            return x
        return s.replace('12/31/1899', '').strip()

    for _time_col in ['GPS_IN_TIME', 'GPS_OUT_TIME', 'ACTUAL_IN_TIME', 'ACTUAL_OUT_TIME']:
        if _time_col in mwa_base.columns:
            mwa_base[_time_col] = mwa_base[_time_col].apply(clean_time_field)

    # Sort by Business A-Z only (stable sort — preserves original FA row order for
    # everything else, matching the pattern observed in the manual reference file).
    mwa_base = mwa_base.sort_values('Business', kind='stable').reset_index(drop=True)

    # ── Sheet 2: MWA Summary ─────────────────────────────────────────────────
    # One row per unique KNPL_ID
    id_cols = ['Business','Zone','Division','Depot Code','Sales Group','KNPL_NAME','KNPL_ID','DOJ']
    summary_base = fa.drop_duplicates('KNPL_ID')[id_cols].copy().reset_index(drop=True)

    # Division: blank/null -> "NA" (Sheet 2 only)
    summary_base['Division'] = summary_base['Division'].apply(
        lambda x: 'NA' if pd.isna(x) or str(x).strip() == '' else x)

    # DOJ Month in mm.yyyy format
    summary_base['DOJ Month'] = pd.to_datetime(summary_base['DOJ'], errors='coerce').dt.strftime('%m.%Y')

    # Col 10 & 13: Max Visit allowed per Category
    summary_base['Max Visit allowed per Category_1'] = summary_base['Business'].apply(
        lambda b: 'All' if b == 'Project' else '1')
    summary_base['Minimum Target Visit_1'] = summary_base['Business'].map(MIN_TARGET_CAT1).fillna('Not Assigned')

    # ── MWA Valid Cat1 / Cat2 ───────────────────────────────────────────────
    # Special businesses (Project, CCCR, CCTA): count ALL visit rows (Category 1,
    # Category 2, Not Assigned combined — i.e. every row, since Category never
    # takes any other value), deduplicated by (PROSPECT_CUST_ID, ACTUAL_DATE) pair.
    SPECIAL_CAT1_BUSINESSES = {'Project', 'CCCR', 'CCTA'}

    special_mask = fa['Business'].isin(SPECIAL_CAT1_BUSINESSES)
    special_df = fa[special_mask]
    special_dedup = special_df.drop_duplicates(subset=['KNPL_ID', 'PROSPECT_CUST_ID', 'ACTUAL_DATE'])
    cat1_counts_special = special_dedup.groupby('KNPL_ID').size().reset_index()
    cat1_counts_special.columns = ['KNPL_ID', 'MWA_Valid_Cat1']

    # All other businesses: unique PROSPECT_CUST_ID where Category == 'Category 1' only
    # (Not Assigned is excluded here).
    normal_df = fa[~special_mask]
    cat1_filter_normal = normal_df[normal_df['Category'] == 'Category 1']
    cat1_counts_normal = cat1_filter_normal.groupby('KNPL_ID')['PROSPECT_CUST_ID'].nunique().reset_index()
    cat1_counts_normal.columns = ['KNPL_ID', 'MWA_Valid_Cat1']

    cat1_counts = pd.concat([cat1_counts_special, cat1_counts_normal], ignore_index=True)

    # MWA Valid Cat2:
    # - Project: Category == 'Category 2', deduplicated by (PROSPECT_CUST_ID, ACTUAL_DATE) pair
    # - All other businesses (including CCCR/CCTA, which never produce Category 2 rows
    #   so this naturally evaluates to 0 for them): Category == 'Category 2',
    #   deduplicated by PROSPECT_CUST_ID only. No exclusion against the Cat1 pool.
    project_cat2 = fa[(fa['Business'] == 'Project') & (fa['Category'] == 'Category 2')]
    project_cat2_dedup = project_cat2.drop_duplicates(subset=['KNPL_ID', 'PROSPECT_CUST_ID', 'ACTUAL_DATE'])
    cat2_counts_project = project_cat2_dedup.groupby('KNPL_ID').size().reset_index()
    cat2_counts_project.columns = ['KNPL_ID', 'MWA_Valid_Cat2']

    other_cat2 = fa[(fa['Business'] != 'Project') & (fa['Category'] == 'Category 2')]
    cat2_counts_other = other_cat2.groupby('KNPL_ID')['PROSPECT_CUST_ID'].nunique().reset_index()
    cat2_counts_other.columns = ['KNPL_ID', 'MWA_Valid_Cat2']

    cat2_counts = pd.concat([cat2_counts_project, cat2_counts_other], ignore_index=True)

    summary_base = summary_base.merge(cat1_counts, on='KNPL_ID', how='left')
    summary_base['MWA_Valid_Cat1'] = summary_base['MWA_Valid_Cat1'].fillna(0).astype(int)
    summary_base = summary_base.merge(cat2_counts, on='KNPL_ID', how='left')
    summary_base['MWA_Valid_Cat2'] = summary_base['MWA_Valid_Cat2'].fillna(0).astype(int)

    summary_base['Max Visit allowed per Category_2'] = summary_base['Max Visit allowed per Category_1']
    summary_base['Minimum Target Visit_2'] = summary_base['Business'].map(MIN_TARGET_CAT2).fillna(0)

    # Slab & Amount
    def get_slab(row):
        return get_slab_and_amount(row['Business'], row['MWA_Valid_Cat1'], row['MWA_Valid_Cat2'])

    summary_base[['Slab','Amount']] = summary_base.apply(
        lambda r: pd.Series(get_slab_and_amount(r['Business'], r['MWA_Valid_Cat1'], r['MWA_Valid_Cat2'])), axis=1)

    # Total Days: unique ACTUAL_DATE per KNPL_ID
    total_days = fa.groupby('KNPL_ID')['ACTUAL_DATE'].nunique().reset_index()
    total_days.columns = ['KNPL_ID','Total Days']
    summary_base = summary_base.merge(total_days, on='KNPL_ID', how='left')
    summary_base['Total Days'] = summary_base['Total Days'].fillna(0).astype(int)

    # Total Visit
    summary_base['Total Visit'] = summary_base['MWA_Valid_Cat1'] + summary_base['MWA_Valid_Cat2']

    # Balance Target Visit — blank
    summary_base['Balance Target Visit\n**Cat1 - Category 1\n**Cat2 - Category 2'] = ''

    mwa_summary = summary_base[[
        'Business','Zone','Division','Depot Code','Sales Group','KNPL_NAME','KNPL_ID','DOJ','DOJ Month',
        'Max Visit allowed per Category_1','Minimum Target Visit_1','MWA_Valid_Cat1',
        'Max Visit allowed per Category_2','Minimum Target Visit_2','MWA_Valid_Cat2',
        'Slab','Amount','Total Days','Total Visit',
        'Balance Target Visit\n**Cat1 - Category 1\n**Cat2 - Category 2'
    ]].copy()

    mwa_summary.columns = [
        'Business','Zone','Division','Depot Code','Sales Group','KNPL_NAME','KNPL_ID','DOJ','DOJ Month',
        'Max Visit allowed per Category','Minimum Target Visit','MWA Valid',
        'Max Visit allowed per Category','Minimum Target Visit','MWA Valid',
        'Slab','Amount','Total Days','Total Visit',
        'Balance Target Visit\n**Cat1 - Category 1\n**Cat2 - Category 2'
    ]

    # Sort by Business A-Z only (stable sort — preserves original row order otherwise,
    # matching the pattern observed in the manual reference file).
    mwa_summary = mwa_summary.sort_values('Business', kind='stable').reset_index(drop=True)


    # ── Sheet 3: FA Upload Summary ────────────────────────────────────────────
    rows = []
    # Sort numerically by Employee ID (KNPL_ID); fall back to string sort if non-numeric
    def _id_sort_key(x):
        try:
            return (0, int(x))
        except (ValueError, TypeError):
            return (1, str(x))
    sorted_ids = sorted(summary_base['KNPL_ID'].dropna().unique(), key=_id_sort_key)

    for kid in sorted_ids:
        row_data = summary_base[summary_base['KNPL_ID'] == kid].iloc[0]
        biz = row_data['Business']
        cat1_v = int(row_data['MWA_Valid_Cat1'])
        cat2_v = int(row_data['MWA_Valid_Cat2'])
        current_amount = row_data['Amount']

        next_c1 = get_next_slab_cat1(biz, cat1_v)
        next_c2 = get_next_slab_cat2(biz, cat2_v)
        next_amt = get_next_slab_amount(biz, current_amount)
        curr_amt_val = _to_amount(current_amount)

        rows.append({
            'Id': f'KNPL - {kid} - 1',
            'Business': biz,
            'Employee ID': kid,
            'Category': 'Category 1',
            'Next Slab Category': next_c1,
            'MTD Category': cat1_v,
            'Balance Category': abs(next_c1 - cat1_v),
        })
        rows.append({
            'Id': f'KNPL - {kid} - 2',
            'Business': biz,
            'Employee ID': kid,
            'Category': 'Category 2',
            'Next Slab Category': next_c2,
            'MTD Category': cat2_v,
            'Balance Category': abs(next_c2 - cat2_v),
        })
        rows.append({
            'Id': f'KNPL - {kid} - Amount',
            'Business': biz,
            'Employee ID': kid,
            'Category': 'Amount',
            'Next Slab Category': next_amt,
            'MTD Category': curr_amt_val,
            'Balance Category': abs(next_amt - curr_amt_val),
        })

    fa_upload_summary = pd.DataFrame(rows, columns=[
        'Id','Business','Employee ID','Category','Next Slab Category','MTD Category','Balance Category'
    ])

    # ── Sheet 4: FA Upload Base ───────────────────────────────────────────────
    # Expand: for each KNPL_ID row in fa_upload_summary (only cat1 & cat2 rows),
    # join with sheet1 records
    base_rows = []
    for kid in sorted_ids:
        kid_fa = fa[fa['KNPL_ID'] == kid].copy()
        biz = summary_base[summary_base['KNPL_ID'] == kid].iloc[0]['Business']

        for _, fa_row in kid_fa.iterrows():
            partner_id = fa_row.get('PROSPECT_CUST_ID', '')
            partner_name = fa_row.get('PROSPECT_CUSTOMER_NAME', '')
            is_blank = (pd.isna(partner_id) or str(partner_id).strip() == '') and \
                       (pd.isna(partner_name) or str(partner_name).strip() == '')
            considered = 0 if is_blank else 1
            excess = 1 if is_blank else 0

            month_val = fa_row.get('MONTH', '')
            creation_date_val = fa_row.get('CREATION_DATE', '')
            visit_month = format_visit_month(month_val, creation_date_val)

            base_rows.append({
                'Business': biz,
                'Employee ID': kid,
                'Category': fa_row.get('Category', ''),
                'Next Slab Category': '',
                'Visit Month': visit_month,
                'Actual Date': fa_row.get('ACTUAL_DATE', ''),
                'Partner ID': partner_id,
                'Partner Name': partner_name,
                'Category Description': fa_row.get('Business Category Description', ''),
                'Considered': considered,
                'Excess Visits': excess,
            })

    fa_upload_base = pd.DataFrame(base_rows, columns=[
        'Business','Employee ID','Category','Next Slab Category','Visit Month',
        'Actual Date','Partner ID','Partner Name','Category Description',
        'Considered','Excess Visits'
    ])

    # Format Actual Date as DD.MM.YYYY
    fa_upload_base['Actual Date'] = pd.to_datetime(
        fa_upload_base['Actual Date'], errors='coerce', dayfirst=True
    ).dt.strftime('%d.%m.%Y')

    # Sort by Business A-Z only (stable sort — preserves original row order otherwise)
    fa_upload_base = fa_upload_base.sort_values('Business', kind='stable').reset_index(drop=True)

    return {
        'MWA Base': mwa_base,
        'MWA Summary': mwa_summary,
        'FA Upload Summary': fa_upload_summary,
        'FA Upload Base': fa_upload_base,
    }


def to_excel(sheets: dict) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        for sheet_name, df in sheets.items():
            df.to_excel(writer, index=False, sheet_name=sheet_name)
    return buf.getvalue()


# ── UI ────────────────────────────────────────────────────────────────────────

if fa_file:
    if st.button("⚙️ Generate MWA Report", use_container_width=True, type="primary"):
        with st.spinner("Processing..."):
            try:
                fa_df = pd.read_excel(fa_file, dtype={'KNPL_ID': str, 'PROSPECT_CUST_ID': str, 'EMPLOYEE_RESPONSIBLE': str})
                sheets = process_mwa(fa_df)

                st.success(f"✅ Done! MWA Report generated successfully.")

                m1, m2, m3, m4 = st.columns(4)
                m1.metric("MWA Base Records",        len(sheets['MWA Base']))
                m2.metric("Unique Employees",         len(sheets['MWA Summary']))
                m3.metric("FA Upload Summary Rows",  len(sheets['FA Upload Summary']))
                m4.metric("FA Upload Base Rows",     len(sheets['FA Upload Base']))

                st.subheader("MWA Summary Preview")
                preview_df = sheets['MWA Summary'].copy()
                preview_df.columns = [
                    'Business','Zone','Division','Depot Code','Sales Group','KNPL_NAME','KNPL_ID','DOJ','DOJ Month',
                    'Max Visit allowed (Cat1)','Min Target Visit (Cat1)','MWA Valid (Cat1)',
                    'Max Visit allowed (Cat2)','Min Target Visit (Cat2)','MWA Valid (Cat2)',
                    'Slab','Amount','Total Days','Total Visit','Balance Target Visit'
                ]
                st.dataframe(preview_df.head(20), use_container_width=True)

                st.download_button(
                    "📥 Download MWA Report (.xlsx)",
                    data=to_excel(sheets),
                    file_name="MWA_Calculation_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True, type="primary")

            except Exception as e:
                st.error(f"❌ Error: {e}")
                st.exception(e)
else:
    st.info("👆 Please upload the FA file to proceed.")