import streamlit as st
import pandas as pd
import gspread
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
import io
import os
import socket
import httplib2
from datetime import datetime, date
import re
import math

# --- 1. SETTING LAYOUT & CUSTOM ELEGANT CSS ---
st.set_page_config(layout="wide", page_title="Legal Dash - PT ASUKA", page_icon="⚖️")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap');
    
    html, body, [class*="css"] { 
        font-family: 'Plus Jakarta Sans', sans-serif; 
        background-color: #f4f7f9;
    }

    .main-title {
        text-align: center; color: #1e3799; font-weight: 800;
        padding-bottom: 20px;
    }

    .sidebar-company-title {
        color: #1e3799; font-size: 18px; font-weight: 800;
        text-align: center; margin-bottom: 20px; padding: 10px;
        border-bottom: 2px solid #f1f2f6;
    }

    [data-testid="stMetric"] {
        background: white; border-radius: 15px; padding: 20px;
        box-shadow: 0 8px 16px rgba(0,0,0,0.05); border: 1px solid #e1e8ed;
    }
    
    [data-testid="stMetric"]:nth-child(1) { border-top: 5px solid #4a69bd; } 
    [data-testid="stMetric"]:nth-child(2) { border-top: 5px solid #78e08f; } 
    [data-testid="stMetric"]:nth-child(3) { border-top: 5px solid #f6b93b; } 
    [data-testid="stMetric"]:nth-child(4) { border-top: 5px solid #eb2f06; }

    .styled-table {
        width: 100%; border-collapse: collapse; margin: 10px 0;
        font-size: 13px; border-radius: 12px; overflow: hidden;
        background-color: white;
    }
    
    .styled-table thead tr {
        background: linear-gradient(135deg, #1e3799 0%, #0c2461 100%);
        color: #ffffff; text-align: center; font-weight: bold;
    }
    
    .styled-table th { padding: 15px; text-align: center !important; }
    .styled-table td { padding: 12px 15px; text-align: left; border-bottom: 1px solid #f1f2f6; }
    .styled-table tbody tr:hover { background-color: #f8faff; }

    .link-pill {
        background: #f1f2f6; color: #1e3799; padding: 5px 10px;
        border-radius: 8px; text-decoration: none; font-size: 11px;
        font-weight: 600; display: inline-block; margin: 2px;
        border: 1px solid #dcdde1; transition: 0.3s;
    }
    .link-pill:hover {
        background: #1e3799; color: white; transform: translateY(-2px);
    }

    section[data-testid="stSidebar"] { background-color: #ffffff; border-right: 1px solid #e1e8ed; }

    .stButton>button {
        background: linear-gradient(135deg, #1e3799 0%, #0c2461 100%);
        color: white; border-radius: 10px; border: none; font-weight: 600; transition: 0.3s;
    }
    </style>
    """, unsafe_allow_html=True)

# --- 2. SETUP KONEKSI ---
@st.cache_resource
def get_services():
    creds_info = dict(st.secrets["gcp_service_account"])
    if "private_key" in creds_info:
        creds_info["private_key"] = creds_info["private_key"].strip().replace("\\n", "\n")
    
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive",
        "https://www.googleapis.com/auth/spreadsheets"
    ]
    
    creds = service_account.Credentials.from_service_account_info(creds_info, scopes=scope)
    gc = gspread.authorize(creds)
    drive = build('drive', 'v3', credentials=creds, cache_discovery=False)
    return gc, drive

# --- 3. KONFIGURASI ---
SPREADSHEET_ID = "1y0rbCetf7-995OWA4LuuBeGbQH2EOCD__q1Uh2iip-M"
ADDENDUM_SPREADSHEET_ID = "1grGo4RLbXa1u-eKaOTAogUqYkx5cIqiX2FbuXIgcRJ8"

CONFIG = {
    'PKWT': {
        'SID': SPREADSHEET_ID,
        'SHEET_NAME': "PKWT_DATA",
        'TEMPLATE_KONTRAK': "18c_YrWda1Wm0SwtcFg8nudBdOXnRmnwo",
        'TEMPLATE_TEMP': "1-vb4oftPxul3YyP-T4Z_tE8waMysMsf6",
        'FOLDERS': {
            'PHOTO': "1-SxkAS8iA-PKIbMztjGmWDxjNxoQxDKj", 
            'SIGNED': "1dtX_EQsyZFeZKJuMABshUiAPBBSx5-rP", 
            'PAKTA': "1xUfwApec3dxz2RpnCFbi0GwDog5K-G4X",
            'PAKTA_T': "1dpdSRe73yTmZFShT0RpzpqwzPZNxcV8T",
            'PAKTA_S': "1GbeSAvGOWcFIAlPGQ2xgjcOTMp8Cte_V"
        },
        'COLS': { 'ID': 5, 'NAMA': 6, 'AWAL': 18, 'AKHIR': 19, 'DEPT': 23, 'AREA': 24, 'DRAFT': 35, 'DRAFT2': 40, 'SIGNED': 43, 'PHOTO': 44, 'PAKTA': 45, 'PAKTA_T': 46, 'PAKTA_S': 47 }
    },
    'PKHL': {
        'SID': SPREADSHEET_ID,
        'SHEET_NAME': "PKHL_DATA",
        'TEMPLATE_KONTRAK': "14M2XJEI684M4GdlWhtVUOYkHiZv27rKY",
        'TEMPLATE_TEMP': "19aheUSFt2UKeb0yryWoidAYSckZuHh9S",
        'FOLDERS': {
            'PHOTO': "17RiCJZtINm250cR7mpcUKTmQYxuY2IXE", 
            'SIGNED': "1dBCDOcsfXksKV-N28ikaKE4hBY9s6abc", 
            'PAKTA': "1PY002I8lzN5FuJuZ8JCPR3-TR1j6-1hH",
            'PAKTA_T': "1Mz84DJtRE1Jlo3fbaAQCXQrNh0E12hi0", 
            'PAKTA_S': "1_t6tzVBpk93CzXf9YXmlIHni8OvoQR6U"
        },
        'COLS': { 'ID': 3, 'NAMA': 4, 'DEPT': 15, 'AREA': 16, 'AWAL': 18, 'AKHIR': 19, 'DRAFT': 24, 'SIGNED': 29, 'PHOTO': 30, 'PAKTA': 31, 'PAKTA_T': 32, 'PAKTA_S': 33 }
    },
    'ADDENDUM': {
        'SID': ADDENDUM_SPREADSHEET_ID,
        'SHEET_NAME': "Merge ADD",
        'TEMPLATE_KONTRAK': "1n1-NhzMobuG3_vCloeSF6N5q341KMQUc",
        'TEMPLATE_TEMP': "1JvBP6QlW86obpzbWBE_XqT43AkumoE4J",
        'FOLDERS': {
            'PHOTO': "1BVs1sKsK9w54-8M7oYKanX9W2xxpNOEb",
            'SIGNED': "1Ob_eLCbciXzDuhpxf0eVhUX_nqQjCyfC",
            'PAKTA': "1MeLj5VA5VkIRTMdR-NvQk98x6wvgx-_Q",
            'PAKTA_T': "1tQ2NN3Aj1RT7lHB6Zm4DhKZYpMBP5cwt",
            'PAKTA_S': "1CfdiXKeFmrHuX34MNODGVKQkesfnudnu"
        },
        'COLS': { 'ID': 10, 'NAMA': 11, 'AWAL': 15, 'AKHIR': 16, 'DRAFT': 26, 'SIGNED': 29, 'PHOTO': 30, 'PAKTA': 31, 'PAKTA_T': 32, 'PAKTA_S': 33 }
    }
}

# --- 4. DATA PROCESSING ---
def parse_indo_date(date_str):
    if not date_str or str(date_str).strip() in ["", "-"]: return None
    indo_months = {'Januari': 'January', 'Februari': 'February', 'Maret': 'March', 'April': 'April', 'Mei': 'May', 'Juni': 'June', 'Juli': 'July', 'Agustus': 'August', 'September': 'September', 'Oktober': 'October', 'November': 'November', 'Desember': 'December'}
    temp_str = str(date_str)
    for indo, eng in indo_months.items(): temp_str = temp_str.replace(indo, eng)
    return pd.to_datetime(temp_str, errors='coerce')

def get_status_logic(akhir_val, mode_type):
    dt = parse_indo_date(akhir_val)
    if dt is None or pd.isnull(dt): 
        return ["Aktif", "🟢 Aktif"] if mode_type in ["PKHL", "ADDENDUM"] else ["Unknown", "⚪ Unknown"]
    today = date.today()
    diff = (dt.date() - today).days
    if diff < 0: return ["Habis", "🔴 Habis"]
    return ["Akan Habis", "🟡 Akan Habis"] if diff <= 30 else ["Aktif", "🟢 Aktif"]

@st.cache_data(ttl=600)
def load_data_optimized(mode):
    gc, _ = get_services()
    cfg = CONFIG[mode]
    sheet = gc.open_by_key(cfg['SID']).worksheet(cfg['SHEET_NAME'])
    raw_data = sheet.get_all_values()
    df = pd.DataFrame(raw_data[1:], columns=raw_data[0])
    df = df.applymap(lambda x: str(x).strip() if x else "")
    
    col = cfg['COLS']
    df['AWAL_DT'] = df.iloc[:, col['AWAL']].apply(parse_indo_date)
    df['AKHIR_DT'] = df.iloc[:, col['AKHIR']].apply(parse_indo_date)
    
    if mode == "PKWT":
        df = df[df['AWAL_DT'].dt.year >= 2024].copy()
    
    status_res = [get_status_logic(val, mode) for val in df.iloc[:, col['AKHIR']]]
    status_df = pd.DataFrame(status_res, columns=['Status_T', 'Status_I'], index=df.index)
    return pd.concat([df, status_df], axis=1)

# --- 5. ACTION HANDLER ---
def execute_action(file_obj, emp_id, emp_name, tgl_awal, col_key, mode, action_type="upload"):
    gc, drive = get_services()
    cfg = CONFIG[mode]
    try:
        with st.spinner("⏳ Sedang memproses..."):
            sheet = gc.open_by_key(cfg['SID']).worksheet(cfg['SHEET_NAME'])
            all_matches = sheet.findall(str(emp_id))
            target_row = None
            
            for match in all_matches:
                if str(sheet.cell(match.row, cfg['COLS']['AWAL'] + 1).value).strip() == str(tgl_awal).strip():
                    target_row = match.row
                    break
            
            if not target_row:
                st.error("❌ Baris data tidak ditemukan.")
                return

            old_link = sheet.cell(target_row, cfg['COLS'][col_key] + 1).value
            if old_link and "drive.google.com" in str(old_link):
                try:
                    fid = re.search(r'(?:id=|\/d\/)([a-zA-Z0-9_-]{25,})', str(old_link)).group(1)
                    drive.files().delete(fileId=fid, supportsAllDrives=True).execute()
                except: pass

            clean_nama = re.sub(r'[^a-zA-Z0-9\s]', '', emp_name).strip()
            new_name = f"{emp_id}_{clean_nama}_{col_key}"
            
            if action_type == "upload":
                fld_id = cfg['FOLDERS'].get(col_key, cfg['FOLDERS']['SIGNED'])
                mime = "image/jpeg" if col_key == 'PHOTO' else "application/pdf"
                media = MediaIoBaseUpload(io.BytesIO(file_obj.getvalue()), mimetype=mime, resumable=True)
                f = drive.files().create(body={'name': new_name, 'parents': [fld_id]}, media_body=media, fields='id, webViewLink', supportsAllDrives=True).execute()
            else:
                tpl_id = cfg['TEMPLATE_KONTRAK'] if col_key == 'PAKTA' else cfg['TEMPLATE_TEMP']
                fld_id = cfg['FOLDERS']['PAKTA'] if col_key == 'PAKTA' else cfg['FOLDERS']['PAKTA_T']
                f = drive.files().copy(fileId=tpl_id, body={'name': new_name, 'parents': [fld_id]}, supportsAllDrives=True, fields='id, webViewLink').execute()

            drive.permissions().create(fileId=f.get('id'), body={'type': 'anyone', 'role': 'reader'}, supportsAllDrives=True).execute()
            sheet.update_cell(target_row, cfg['COLS'][col_key] + 1, f.get('webViewLink'))
            
            st.cache_data.clear()
            st.success("✅ Berhasil Diperbarui!")
            st.rerun()
    except Exception as e: st.error(f"⚠️ Error: {str(e)}")

# --- 6. UI RENDER ---
with st.sidebar:
    st.markdown('<div class="sidebar-company-title">PT ASUKA ENGINEERING INDONESIA</div>', unsafe_allow_html=True)
    st.markdown("### 📋 Menu Utama")
    mode = st.radio("DASHBOARD MODE:", ["PKWT", "PKHL", "ADDENDUM"])
    if st.button("🔄 Refresh Seluruh Data"): st.cache_data.clear(); st.rerun()

df_master = load_data_optimized(mode)
cfg, col_idx = CONFIG[mode], CONFIG[mode]['COLS']

st.markdown(f"<h1 class='main-title'>DASHBOARD LEGAL - {mode}</h1>", unsafe_allow_html=True)

# Metrics
m1, m2, m3, m4 = st.columns(4)
m1.metric("Total Personel", len(df_master))
m2.metric("Aktif", len(df_master[df_master['Status_T'] == "Aktif"]))
m3.metric("Akan Habis", len(df_master[df_master['Status_T'] == "Akan Habis"]))
m4.metric("Habis", len(df_master[df_master['Status_T'] == "Habis"]), delta_color="inverse")

# Filtering Section
st.markdown("---")
with st.container():
    st.markdown("### 🔍 Filter & Pencarian")
    c1, c2, c3, c4 = st.columns([2, 1, 1, 1])
    search_query = c1.text_input("Cari Nama atau ID...")
    
    if mode != "ADDENDUM":
        dept_f = c2.selectbox("Departemen", ["Semua"] + sorted(df_master.iloc[:, col_idx['DEPT']].unique().tolist()))
        area_f = c3.selectbox("Area Kerja", ["Semua"] + sorted(df_master.iloc[:, col_idx['AREA']].unique().tolist()))
        stat_f = c4.selectbox("Status", ["Semua", "Aktif", "Akan Habis", "Habis"])
    else:
        stat_f = c2.selectbox("Status", ["Semua", "Aktif", "Akan Habis", "Habis"])
        dept_f, area_f = "Semua", "Semua"

    start_date, end_date = None, None
    if mode in ["PKWT", "ADDENDUM"]:
        st.markdown("#### 📅 Rentang Tanggal")
        dc1, dc2 = st.columns(2)
        start_date = dc1.date_input("Mulai Dari", value=None)
        end_date = dc2.date_input("Sampai Dengan", value=None)

# Apply Filter
dff = df_master.copy()
if search_query:
    mask = (dff.iloc[:, col_idx['NAMA']].str.contains(search_query, case=False, na=False)) | \
           (dff.iloc[:, col_idx['ID']].astype(str).str.contains(search_query, case=False, na=False))
    dff = dff[mask]

if dept_f != "Semua": dff = dff[dff.iloc[:, col_idx['DEPT']] == dept_f]
if area_f != "Semua": dff = dff[dff.iloc[:, col_idx['AREA']] == area_f]
if stat_f != "Semua": dff = dff[dff['Status_T'] == stat_f]
if start_date: dff = dff[dff['AWAL_DT'].dt.date >= start_date]
if end_date: dff = dff[dff['AWAL_DT'].dt.date <= end_date]

# --- 7. PAGINATION SYSTEM ---
st.markdown("---")
st.subheader("📋 Tabel Data Personel")

items_per_page = 30
total_data = len(dff)
total_pages = math.ceil(total_data / items_per_page) if total_data > 0 else 1
col_page1, col_page2 = st.columns([1, 4])
current_page = col_page1.number_input(f"Halaman (1 - {total_pages})", min_value=1, max_value=total_pages, step=1)

start_idx = (current_page - 1) * items_per_page
end_idx = start_idx + items_per_page
disp_df = dff.iloc[start_idx:end_idx].copy()

# LOGIKA PREVIEW BERKAS - DIPERBAIKI (Tahan Error Non-String)
def make_pills(r):
    p = []
    
    # Fungsi pembantu untuk cek link dengan aman
    def check_link(val, label):
        s_val = str(val).strip() if val else ""
        if s_val.lower().startswith("http"):
            return f"<a class='link-pill' href='{s_val}' target='_blank'>{label}</a>"
        return ""

    # Draft
    p.append(check_link(r.iloc[col_idx['DRAFT']], "Draft 1"))
    if 'DRAFT2' in col_idx:
        p.append(check_link(r.iloc[col_idx['DRAFT2']], "Draft 2"))
    
    # Berkas Utama
    check_keys = [('SIGNED','Kontrak'), ('PAKTA','Pakta'), ('PAKTA_T','P-Temp'), ('PAKTA_S','P-Signed'), ('PHOTO','Foto')]
    for k, label in check_keys:
        p.append(check_link(r.iloc[col_idx[k]], label))
            
    # Gabung semua pill yang tidak kosong
    p_final = [x for x in p if x != ""]
    return " ".join(p_final) if p_final else "-"

disp_df['BERKAS'] = disp_df.apply(make_pills, axis=1)

# Pilih kolom tampil
if mode != "ADDENDUM":
    vcols = [dff.columns[col_idx['ID']], dff.columns[col_idx['NAMA']], dff.columns[col_idx['DEPT']], dff.columns[col_idx['AREA']], dff.columns[col_idx['AWAL']], dff.columns[col_idx['AKHIR']], 'Status_I', 'BERKAS']
else:
    vcols = [dff.columns[col_idx['ID']], dff.columns[col_idx['NAMA']], dff.columns[col_idx['AWAL']], dff.columns[col_idx['AKHIR']], 'Status_I', 'BERKAS']

st.markdown(f"**Menampilkan data {start_idx + 1} sampai {min(end_idx, total_data)} dari total {total_data} record**")
st.markdown(f"<table class='styled-table'><thead><tr>{''.join(f'<th>{c}</th>' for c in vcols)}</tr></thead><tbody>" + 
            "".join(f"<tr>{''.join(f'<td>{row[c]}</td>' for c in vcols)}</tr>" for _, row in disp_df.iterrows()) + 
            "</tbody></table>", unsafe_allow_html=True)

# --- 8. CONTROL CENTER ---
st.markdown("---")
st.subheader("⚙️ Control Center")
cc1, cc2 = st.columns([1, 2])
emp_search = cc1.text_input("Cari Nama/ID untuk Update Data")
u_list = df_master.copy()
if emp_search:
    u_list = u_list[u_list.iloc[:, col_idx['NAMA']].str.contains(emp_search, case=False, na=False) | \
                   u_list.iloc[:, col_idx['ID']].astype(str).str.contains(emp_search, case=False, na=False)]

u_list['label'] = u_list.apply(lambda r: f"{r.iloc[col_idx['NAMA']]} ({r.iloc[col_idx['ID']]}) | {r.iloc[col_idx['AWAL']]}", axis=1)
target_emp = cc2.selectbox("Pilih Personel", ["-- Pilih --"] + u_list['label'].tolist())

if target_emp != "-- Pilih --":
    row_data = u_list[u_list['label'] == target_emp].iloc[0]
    id_e, nm_e, aw_e = row_data.iloc[col_idx['ID']], row_data.iloc[col_idx['NAMA']], row_data.iloc[col_idx['AWAL']]
    st.info(f"📍 Sedang memproses: **{nm_e}** ({id_e})")
    a1, a2, a3 = st.columns(3)
    with a1:
        st.markdown("#### 📂 Upload")
        f1 = st.file_uploader("Upload Kontrak Signed", type="pdf", key="up1")
        if f1 and st.button("Simpan Kontrak"): execute_action(f1, id_e, nm_e, aw_e, 'SIGNED', mode)
        f2 = st.file_uploader("Upload Pakta Signed", type="pdf", key="up2")
        if f2 and st.button("Simpan Pakta Signed"): execute_action(f2, id_e, nm_e, aw_e, 'PAKTA_S', mode)
    with a2:
        st.markdown("#### 🤖 Generate")
        if st.button("Gen Pakta Kontrak"): execute_action(None, id_e, nm_e, aw_e, 'PAKTA', mode, "gen")
        if st.button("Gen Pakta Temporary"): execute_action(None, id_e, nm_e, aw_e, 'PAKTA_T', mode, "gen")
    with a3:
        st.markdown("#### 📸 Foto")
        shot = st.camera_input("Ambil Foto Baru")
        if shot and st.button("Upload Foto"): execute_action(shot, id_e, nm_e, aw_e, 'PHOTO', mode)

st.markdown("<br><p style='text-align: center; color: #7f8c8d;'>© 2026 PT ASUKA ENGINEERING INDONESIA - Legal Management System v2.7</p>", unsafe_allow_html=True)