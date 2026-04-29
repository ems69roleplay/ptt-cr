import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import time
import requests
from streamlit_option_menu import option_menu
from streamlit_paste_button import paste_image_button
from fpdf import FPDF
import io

# Fungsi Export Excel
def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Laporan')
    return output.getvalue()

# Fungsi Export PDF (Sederhana & Rapi)
def create_pdf(df, title_text):
    pdf = FPDF(orientation='L', unit='mm', format='A4') # Landscape agar muat banyak kolom
    pdf.add_page()
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(0, 10, title_text, ln=True, align='C')
    pdf.set_font("Arial", '', 10)
    pdf.cell(0, 10, f"Dicetak pada: {datetime.now().strftime('%d/%m/%Y %H:%M')}", ln=True, align='C')
    pdf.ln(5)
    
    # Header Tabel
    pdf.set_fill_color(200, 220, 255)
    pdf.set_font("Arial", 'B', 8)
    
    # Tentukan lebar kolom (total ~270mm untuk Landscape A4)
    widths = [30, 25, 80, 30, 30, 30, 45] 
    cols = ["ID", "TANGGAL", "KEPERLUAN", "MASUK", "KELUAR", "PIC", "BUKTI"]
    
    # Sesuaikan kolom yang ada di dataframe kamu
    for i, col in enumerate(cols):
        pdf.cell(widths[i], 10, col, border=1, fill=True, align='C')
    pdf.ln()
    
    # Isi Tabel
    pdf.set_font("Arial", '', 7)
    for _, row in df.iterrows():
        # Mengambil data yang sesuai (sesuaikan nama kolom GSheets kamu)
        data = [
            str(row.get('ID_TRANSAKSI', '')),
            str(row.get('TANGGAL', ''))[:10],
            str(row.get('KEPERLUAN', ''))[:50],
            f"{row.get('UANG_MASUK', 0):,.0f}",
            f"{row.get('UANG_KELUAR', 0):,.0f}",
            str(row.get('PIC', '')),
            str(row.get('BUKTI_LINK', ''))[:30]
        ]
        for i, item in enumerate(data):
            pdf.cell(widths[i], 8, item, border=1)
        pdf.ln()
        
    return pdf.output(dest='S').encode('latin-1')

# --- CONFIGURATION & STYLES ---
st.set_page_config(page_title="PTT Management CR Roleplay", layout="wide")

# Gunakan Path atau URL jika nanti sudah online, untuk sekarang pastikan background.png satu folder
import base64

def get_base64_of_bin_file(bin_file):
    with open(bin_file, 'rb') as f:
        data = f.read()
    return base64.b64encode(data).decode()

def set_png_as_page_bg(bin_file):
    bin_str = get_base64_of_bin_file(bin_file)
    page_bg_img = f'''
    <style>
    .stApp {{
        background-image: url("data:image/png;base64,{bin_str}");
        background-size: cover;
        background-attachment: fixed;
    }}
    #MainMenu {{visibility: hidden;}}
    footer {{visibility: hidden;}}
    header {{visibility: hidden;}}
    
    /* Agar box konten tetap terbaca dengan latar belakang gambar */
    [data-testid="stSidebar"] {{
        background-color: rgba(0, 8, 20, 0.8) !important;
        backdrop-filter: blur(10px);
    }}
    
    .main-card, div[data-testid="stExpander"] {{
        background-color: rgba(1, 42, 74, 0.6) !important;
        backdrop-filter: blur(5px);
        border: 1px solid rgba(0, 255, 153, 0.2);
    }}
    </style>
    '''
    st.markdown(page_bg_img, unsafe_allow_html=True)

# Panggil fungsi background
try:
    set_png_as_page_bg('background.png')
except:
    st.warning("File background.png tidak ditemukan, menggunakan warna polos.")

PRIMARY_COLOR = "#001d3d"
ACCENT_COLOR = "#00ff99"

# --- DATABASE CONNECTION ---
conn = st.connection("gsheets", type=GSheetsConnection)
@st.cache_data(show_spinner=False)
def get_data(sheet):
    return conn.read(worksheet=sheet, ttl=0)

def write_data(sheet, df):
    conn.update(worksheet=sheet, data=df)
    st.cache_data.clear()

# --- HELPER FUNCTIONS ---
def log_act(user, aksi, detail):
    try:
        df_log = get_data("LOG_AKTIVITAS")
        new_log = pd.DataFrame([{
            "WAKTU": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "USER": user, "AKSI": aksi, "DETAIL": detail
        }])
        write_data("LOG_AKTIVITAS", pd.concat([df_log, new_log], ignore_index=True))
    except:
        pass

def send_discord(msg, file_data=None):
    try:
        webhook_url = st.secrets["discord"]["webhook_url"]
        payload = {"content": msg}
        
        if file_data:
            # Jika ada file, kirim sebagai attachment
            files = {"file": ("bukti.png", file_data, "image/png")}
            r = requests.post(webhook_url, data=payload, files=files)
        else:
            # Jika hanya teks
            r = requests.post(webhook_url, json=payload)
        
        # Mengambil link gambar dari respon Discord untuk disimpan ke Sheets
        if r.status_code in [200, 204] and file_data:
            res_json = r.json()
            if "attachments" in res_json and len(res_json["attachments"]) > 0:
                return res_json["attachments"][0]["url"]
        return None
    except:
        return None

def generate_id(prefix, df):
    if df.empty or "ID_TRANSAKSI" not in df.columns: 
        return f"{prefix}-001"
    
    # Filter hanya ID yang punya prefix yang sama (contoh: hanya ambil yang ada 'EO')
    df_filtered = df[df["ID_TRANSAKSI"].astype(str).str.startswith(prefix)]
    
    if df_filtered.empty:
        return f"{prefix}-001"
    
    # Ambil angka dari semua ID yang terfilter, lalu cari yang paling besar
    try:
        # Mengambil angka setelah tanda '-' dan mengubahnya jadi integer
        numbers = df_filtered["ID_TRANSAKSI"].str.split('-').str[1].astype(int)
        next_num = numbers.max() + 1
        return f"{prefix}-{next_num:03d}"
    except:
        return f"{prefix}-001"

# --- AUTHENTICATION SYSTEM ---
if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.user = None

def login_page():
    # CSS Minimalis hanya untuk meratakan teks dan memperkecil logo
    st.markdown("""
        <style>
        .block-container {
            padding-top: 2rem !important;
        }
        /* Memastikan judul berada di tengah teks */
        .centered-text {
            text-align: center;
            width: 100%;
            margin-bottom: 20px;
        }
        /* Menghilangkan margin bawah default pada image */
        [data-testid="stImage"] {
            margin-bottom: 0px;
        }
        </style>
    """, unsafe_allow_html=True)

    # Menggunakan 3 kolom: kolom tengah yang akan kita isi (Rasio 1:2:1)
    col_left, col_mid, col_right = st.columns([1, 2, 1])
    
    with col_mid:
        # 1. Logo Pasukan Tutul (Menggunakan kolom di dalam kolom agar logo benar-benar tengah)
        sub_left, sub_mid, sub_right = st.columns([1, 2, 1])
        with sub_mid:
            st.image("ptt-logo.png", use_container_width=True)
            
        # 2. Judul Sistem
        st.markdown("<h3 class='centered-text'>SISTEM MANAJEMEN KEUANGAN</h3>", unsafe_allow_html=True)
        
        # 3. Form Login/Daftar
        tab1, tab2 = st.tabs(["LOGIN", "DAFTAR"])
        
        with tab1:
            u = st.text_input("Username", key="login_user").lower()
            p = st.text_input("Password", type="password", key="login_pass")
            st.write("") 
            if st.button("Masuk", use_container_width=True):
                df_u = get_data("KEANGGOTAAN")
                user_data = df_u[(df_u['USERNAME'].str.lower() == u) & (df_u['PASSWORD'].astype(str) == p)]
                if not user_data.empty:
                    if user_data.iloc[0]['STATUS'] == 'Aktif':
                        st.session_state.logged_in = True
                        st.session_state.user = user_data.iloc[0].to_dict()
                        log_act(u, "LOGIN", "Berhasil masuk ke sistem")
                        st.rerun()
                    else: 
                        st.error("Akun Anda dinonaktifkan.")
                else: 
                    st.error("Username/Password salah.")

        with tab2:
            new_u = st.text_input("Username Baru", key="reg_user").lower()
            new_n = st.text_input("Nama IC (Sesuai KTP)", key="reg_name")
            new_p = st.text_input("Password Baru", type="password", key="reg_pass")
            if st.button("Daftar Sekarang", use_container_width=True):
                df_u = get_data("KEANGGOTAAN")
                if new_u in df_u['USERNAME'].str.lower().values:
                    st.warning("Username sudah terdaftar.")
                else:
                    new_reg = pd.DataFrame([{"USERNAME": new_u, "PASSWORD": new_p, "NAMA_IC": new_n, "ROLE": "Anggota", "STATUS": "Aktif"}])
                    write_data("KEANGGOTAAN", pd.concat([df_u, new_reg], ignore_index=True))
                    st.success("Berhasil daftar! Silakan login.")

# --- MAIN APP ---
if not st.session_state.logged_in:
    login_page()
else:
    user = st.session_state.user
    with st.sidebar:
        st.image("ptt-logo.png", width=250)
        st.write("---")

        st.title(f"👤 {user['ROLE']}")
        st.write(f"User: {user['NAMA_IC']}")
        
        menu_options = ["Dashboard", "Form Pengajuan", "Daftar Transaksi"]
        menu_icons = ["speedometer2", "pencil-square", "list-task"]
        
        if user['ROLE'] in ['Admin', 'Sekretaris', 'Bendahara']:
            menu_options += ["Persetujuan", "Kelola Event", "Administrasi", "Log"]
            menu_icons += ["check2-circle", "calendar-event", "shield-lock", "journal-text"]
        
        menu = option_menu(None, menu_options, icons=menu_icons, menu_icon="cast", default_index=0)
        if st.button("🚪 Logout"):
            st.session_state.logged_in = False
            st.rerun()

    # --- DASHBOARD ---
    if menu == "Dashboard":
        st.title("📈 Dashboard Keuangan")
        df_bal = get_data("RINGKASAN_SALDO")
        df_trx = get_data("KAS_TRANSAKSI")
        df_pen = get_data("PENGAJUAN_DANA")
    
        df_bal.columns = df_bal.columns.str.strip()
        kolom_saldo = 'Sisa Saldo' 
        df_bal[kolom_saldo] = pd.to_numeric(df_bal[kolom_saldo], errors='coerce').fillna(0)
        df_bal = df_bal[~df_bal.iloc[:, 0].astype(str).str.contains("#N/A", na=True)]
        df_bal = df_bal.dropna(subset=[df_bal.columns[0]])
        df_bal['Sisa Saldo'] = pd.to_numeric(df_bal['Sisa Saldo'], errors='coerce').fillna(0)
    
        total_dana = df_bal[kolom_saldo].sum()
        pending_count = len(df_pen[df_pen['STATUS'].astype(str).str.strip().str.upper() == 'PENDING'])

        c1, c2, c3 = st.columns(3)
        c1.metric("Total Seluruh Saldo", f"Rp {total_dana:,.0f}")
        c2.metric("Jumlah Event", f"{len(df_bal)} Item")
        c3.metric("Pengajuan Pending", f"{pending_count} Item")

        col_l, col_r = st.columns(2)
        with col_l:
            st.subheader("Arus Kas (Cash Flow)")
            if not df_trx.empty:
                # 1. Paksa kolom TANGGAL jadi tipe datetime, lalu ambil tanggalnya saja
                df_trx['TANGGAL_DT'] = pd.to_datetime(df_trx['TANGGAL'], errors='coerce')
                df_trx['TANGGAL_STR'] = df_trx['TANGGAL_DT'].dt.strftime('%d %b %Y') # Format: 26 Apr 2026
                
                # 2. Kelompokkan berdasarkan tanggal (agar tidak ada jam sama sekali)
                df_daily = df_trx.groupby(['TANGGAL_STR', 'TANGGAL_DT'])[['UANG_MASUK', 'UANG_KELUAR']].sum().reset_index()
                
                # Urutkan berdasarkan waktu asli agar urutan tanggal di grafik benar
                df_daily = df_daily.sort_values('TANGGAL_DT')
                
                # 3. Buat Grafik dengan sumbu X sebagai teks (Kategorikal) agar tidak muncul milidetik
                fig1 = px.line(df_daily, x="TANGGAL_STR", y=["UANG_MASUK", "UANG_KELUAR"], 
                            template="plotly_dark", 
                            labels={"TANGGAL_STR": "Tanggal", "value": "Rupiah", "variable": "Jenis"},
                            color_discrete_sequence=[ACCENT_COLOR, "#ff4b4b"])
                
                # Tambahkan titik agar data harian terlihat jelas
                fig1.update_traces(mode='lines+markers')
                
                # Paksa sumbu X agar tidak mencoba menebak skala waktu
                fig1.update_xaxes(type='category', tickangle=0)
                
                st.plotly_chart(fig1, use_container_width=True, key="dash_flow")
            
        with col_r:
            st.subheader("Distribusi Saldo per Event")
            if not df_bal.empty:
                fig2 = px.pie(df_bal, values=kolom_saldo, names=df_bal.columns[0], 
                            hole=.5, template="plotly_dark", color_discrete_sequence=px.colors.sequential.Greens_r)
                st.plotly_chart(fig2, use_container_width=True, key="dash_pie")

    # --- FORM PENGAJUAN ---
    elif menu == "Form Pengajuan":
        st.title("📝 Form Transaksi")
        
        df_ev = get_data("MASTER_EVENT")
        df_ev.columns = df_ev.columns.str.strip()
        list_event_aktif = df_ev[df_ev['STATUS'].str.strip().str.lower() == 'aktif']['NAMA_EVENT'].tolist()
        pilihan_kategori = ["INTERNAL"] + [f"Event ({ev})" for ev in list_event_aktif]

        # --- FITUR RESET KEY ---
        # --- FITUR BUKTI (VERSI RAPI) ---
        st.write("📸 **BUKTI TRANSAKSI (UPLOAD/PASTE)**")
        
        if "file_key" not in st.session_state:
            st.session_state.file_key = 0

        col_u1, col_u2 = st.columns(2)
        
        with col_u1:
            uploaded_file = st.file_uploader("Upload File Manual", type=["png", "jpg", "jpeg"], key=f"u_{st.session_state.file_key}")
        
        with col_u2:
            # Trik CSS untuk memberi jarak atas agar sejajar dengan tombol upload di sampingnya
            st.markdown('<div style="margin-top: 28px;"></div>', unsafe_allow_html=True)
            pasted_image = paste_image_button(
                label="📋 Klik & Paste Screenshot (Ctrl+V)", 
                key=f"p_{st.session_state.file_key}"
            )

        img_bin = None 
        if uploaded_file:
            img_bin = uploaded_file.getvalue()
            st.image(img_bin, caption="Preview Upload", width=300)
        elif pasted_image.image_data is not None:
            import io
            buf = io.BytesIO()
            pasted_image.image_data.save(buf, format="PNG")
            img_bin = buf.getvalue()
            st.image(img_bin, caption="Preview Paste", width=300)

        with st.form("trx_form", clear_on_submit=True):
            tipe = st.selectbox("JENIS TRANSAKSI", ["Uang Keluar (Butuh Persetujuan)", "Uang Masuk (Langsung Simpan)"])
            kat_pilihan = st.selectbox("KATEGORI / EVENT", pilihan_kategori)
            kat_final = "INTERNAL" if kat_pilihan == "INTERNAL" else "EVENT"
            ev_name_final = "KAS INTERNAL" if kat_pilihan == "INTERNAL" else kat_pilihan.replace("Event (", "").replace(")", "")
            desc = st.text_area("DESKRIPSI / KEPERLUAN")
            nominal = st.number_input("NOMINAL", min_value=0)
            bukti_manual = st.text_input("LINK BUKTI (Isi jika ingin kirim link foto luar)")
            catatan = st.text_input("CATATAN TAMBAHAN")
            
            if st.form_submit_button("EKSEKUSI TRANSAKSI"):
                now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                # --- LOGIKA DISCORD CERDAS ---
                msg_discord = f"💰 **[UANG MASUK] DANA MASUK BARU**\n━━━━━━━━━━━━━━━━━━\n🆔 ID: GENERATING...\n📅 Event: {ev_name_final}\n💵 Nominal: Rp {nominal:,.0f}\n👤 Oleh: {user['NAMA_IC']}\n📝 Detail: {desc}" if "Masuk" in tipe else f"🔔 **[UANG KELUAR] PENGAJUAN BARU**\n━━━━━━━━━━━━━━━━━━\n🆔 ID: GENERATING...\n📅 Event: {ev_name_final}\n💵 Nominal: Rp {nominal:,.0f}\n👤 Pengaju: {user['NAMA_IC']}\n📝 Status: PENDING"
                
                # Jika ada link manual tapi tidak ada upload gambar, tambahkan linknya ke pesan agar preview muncul di Discord
                if not img_bin and bukti_manual:
                    msg_discord += f"\n🖼️ Bukti: {bukti_manual}"

                link_dari_discord = send_discord(msg_discord, file_data=img_bin)
                final_bukti_link = link_dari_discord if link_dari_discord else bukti_manual

                if "Masuk" in tipe:
                    df_kas = get_data("KAS_TRANSAKSI")
                    # FIX ID: Paksa ambil ID unik tanpa cache
                    new_id = generate_id("EO", df_kas)
                    new_row = pd.DataFrame([{"ID_TRANSAKSI": new_id, "TANGGAL": now, "KATEGORI": kat_final, "NAMA_EVENT": ev_name_final, "KEPERLUAN": desc, "UANG_MASUK": nominal, "UANG_KELUAR": 0, "PIC": user['NAMA_IC'], "BUKTI_LINK": final_bukti_link}])
                    write_data("KAS_TRANSAKSI", pd.concat([df_kas, new_row], ignore_index=True))
                    
                    time.sleep(2)
                    df_bal_new = get_data("RINGKASAN_SALDO")
                    current_saldo = df_bal_new[df_bal_new['Kantong/Event'].astype(str).str.strip() == ev_name_final]['Sisa Saldo'].values[0]
                    send_discord(f"🆔 Update ID: {new_id}\n💳 Sisa Saldo {ev_name_final}: Rp {current_saldo:,.0f}")
                    log_act(user['NAMA_IC'], "INPUT KAS MASUK", f"Input {new_id} ke {ev_name_final} sebesar Rp {nominal:,.0f}")
                    
                    st.success("✅ Dana Masuk Tercatat!")
                else:
                    df_pen = get_data("PENGAJUAN_DANA")
                    new_id = generate_id("REQ", df_pen)
                    new_req = pd.DataFrame([{"ID_TRANSAKSI": new_id, "TANGGAL_PENGAJUAN": now, "NAMA_ANGGOTA": user['NAMA_IC'], "NAMA_EVENT": ev_name_final, "KEPERLUAN": desc, "NOMINAL": nominal, "STATUS": "PENDING", "BUKTI_LINK": final_bukti_link, "CATATAN": catatan}])
                    write_data("PENGAJUAN_DANA", pd.concat([df_pen, new_req], ignore_index=True))
                    send_discord(f"🆔 Update ID: {new_id}")
                    log_act(user['NAMA_IC'], "INPUT PENGAJUAN", f"Mengajukan {new_id} untuk {ev_name_final}")
                    
                    st.success(f"📩 Pengajuan {new_id} Berhasil!")

                st.session_state.file_key += 1
                st.balloons()
                time.sleep(1)
                st.rerun()

    elif menu == "Daftar Transaksi":
        st.title("📜 Riwayat Transaksi Pribadi")
        st.write(f"Menampilkan riwayat aktivitas untuk: **{user['NAMA_IC']}**")
        
        # 1. Ambil data dari KAS_TRANSAKSI
        df_trx = get_data("KAS_TRANSAKSI")
        
        # 2. Filter data agar HANYA menampilkan transaksi milik user yang sedang login
        # Kita filter berdasarkan kolom PIC (yang berisi NAMA_IC)
        df_pribadi = df_trx[df_trx['PIC'] == user['NAMA_IC']]
        
        if df_pribadi.empty:
            st.info("Kamu belum memiliki riwayat transaksi.")
        else:
            # 3. Percantik tampilan tabel
            # Urutkan dari yang terbaru (asumsi baris bawah adalah terbaru)
            df_display = df_pribadi.iloc[::-1] 
            
            # Tampilkan metrik ringkasan pribadi
            c1, c2 = st.columns(2)
            total_masuk = df_display['UANG_MASUK'].sum()
            total_keluar = df_display['UANG_KELUAR'].sum()
            
            c1.metric("Total Input Masuk", f"Rp {total_masuk:,.0f}")
            c2.metric("Total Input Keluar", f"Rp {total_keluar:,.0f}")
            
            st.write("---")
            
            # Tampilkan dataframe dengan desain yang rapi
            st.dataframe(
                df_display[['TANGGAL', 'KATEGORI', 'NAMA_EVENT', 'KEPERLUAN', 'UANG_MASUK', 'UANG_KELUAR', 'BUKTI_LINK']],
                use_container_width=True,
                hide_index=True
            )


    # --- PERSETUJUAN ---
    elif menu == "Persetujuan":
        st.title("⚔️ Panel Persetujuan")
        df_pen = get_data("PENGAJUAN_DANA")
        pending = df_pen[df_pen['STATUS'].astype(str).str.upper() == 'PENDING']
        
        if pending.empty:
            st.info("Tidak ada pengajuan pending.")
        else:
            for i, r in pending.iterrows():
                with st.expander(f"ID: {r['ID_TRANSAKSI']} | {r['NAMA_ANGGOTA']} - Rp {r['NOMINAL']:,.0f}"):
                    st.write(f"Keperluan: {r['KEPERLUAN']}")
                    col_b1, col_b2 = st.columns(2)
                    
                    if col_b1.button("SETUJUI", key=f"acc_{r['ID_TRANSAKSI']}"):
                        df_kas = get_data("KAS_TRANSAKSI") # <-- Harus ada ini
                        df_pen.at[i, 'STATUS'] = 'DITERIMA'
                        new_kas = pd.DataFrame([{ # <-- Dan ini
                            "ID_TRANSAKSI": r['ID_TRANSAKSI'], 
                            "TANGGAL": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "KATEGORI": "EVENT" if r['NAMA_EVENT'] != "KAS INTERNAL" else "INTERNAL",
                            "NAMA_EVENT": r['NAMA_EVENT'], 
                            "KEPERLUAN": r['KEPERLUAN'], 
                            "UANG_MASUK": 0,
                            "UANG_KELUAR": r['NOMINAL'], 
                            "PIC": r['NAMA_ANGGOTA'], 
                            "BUKTI_LINK": r['BUKTI_LINK']
                        }])
                        write_data("KAS_TRANSAKSI", pd.concat([df_kas, new_kas], ignore_index=True))
                        write_data("PENGAJUAN_DANA", df_pen)
                        
                        # Ambil saldo terbaru setelah disetujui (Sudah berkurang)
                        time.sleep(2)
                        df_bal_new = get_data("RINGKASAN_SALDO")
                        current_saldo = df_bal_new[df_bal_new['Kantong/Event'].astype(str).str.strip() == r['NAMA_EVENT']]['Sisa Saldo'].values[0]
                        
                        send_discord(f"✅ **[UANG KELUAR] PENGAJUAN DISETUJUI**\n"
                                     f"━━━━━━━━━━━━━━━━━━\n"
                                     f"🆔 ID: {r['ID_TRANSAKSI']}\n"
                                     f"📅 Event: {r['NAMA_EVENT']}\n"
                                     f"💵 Nominal: Rp {r['NOMINAL']:,.0f}\n"
                                     f"💳 Sisa Saldo {r['NAMA_EVENT']}: Rp {current_saldo:,.0f}")
                        log_act(user['NAMA_IC'], "PERSETUJUAN", f"Menyetujui dana {r['ID_TRANSAKSI']} sebesar Rp {r['NOMINAL']:,.0f}")
                        
                        st.rerun()

                    if col_b2.button("TOLAK", key=f"rej_{r['ID_TRANSAKSI']}"):
                        df_pen.at[i, 'STATUS'] = 'DITOLAK'
                        write_data("PENGAJUAN_DANA", df_pen)
                        send_discord(f"❌ **DITOLAK**\nID: {r['ID_TRANSAKSI']}")
                        log_act(user['NAMA_IC'], "PERSETUJUAN", f"Menolak pengajuan dana {r['ID_TRANSAKSI']}")
                        
                        st.rerun()

    # --- KELOLA EVENT ---
    elif menu == "Kelola Event":
        st.title("📅 Kelola Master Event")
        df_ev = get_data("MASTER_EVENT")
        
        with st.form("new_ev"):
            n_ev = st.text_input("Nama Event Baru")
            if st.form_submit_button("Tambah Event"):
                new_row = pd.DataFrame([{"NAMA_EVENT": n_ev, "STATUS": "Aktif"}])
                write_data("MASTER_EVENT", pd.concat([df_ev, new_row], ignore_index=True))
                log_act(user['NAMA_IC'], "KELOLA EVENT", f"Menambah event baru: {n_ev}")
                st.rerun()
        
        for i, r in df_ev.iterrows():
            with st.container():
                c1, c2 = st.columns([4, 1])
                color = ACCENT_COLOR if r['STATUS'] == 'Aktif' else "#ff4b4b"
                c1.markdown(f"<div style='padding:10px; border-left:5px solid {color}; background:#012a4a; border-radius:10px;'><b>{r['NAMA_EVENT']}</b> ({r['STATUS']})</div>", unsafe_allow_html=True)
                if c2.button("Hapus", key=f"del_ev_{i}"):
                    write_data("MASTER_EVENT", df_ev.drop(i))
                    st.rerun()

    # --- ADMINISTRASI ---
    elif menu == "Administrasi":
        st.title("🛡️ Administrasi")
        tab_trx, tab_user = st.tabs(["📊 Transaksi", "👥 Anggota"])
        with tab_trx:
            st.subheader("📊 Monitoring Seluruh Transaksi")
            df_admin_trx = get_data("KAS_TRANSAKSI")
            
            # --- BAGIAN FILTER (Samping-sampingan) ---
            with st.container():
                c1, c2, c3, c4 = st.columns(4)
                
                # Filter berdasarkan PIC
                list_pic = ["Semua"] + sorted(df_admin_trx['PIC'].unique().tolist())
                f_user = c1.selectbox("👤 Pilih PIC", list_pic)
                
                # Filter berdasarkan Nama Event
                list_ev = ["Semua"] + sorted(df_admin_trx['NAMA_EVENT'].unique().tolist())
                f_ev = c2.selectbox("📅 Pilih Event", list_ev)
                
                # Filter berdasarkan Jenis Transaksi
                f_tipe = c3.selectbox("💰 Jenis Dana", ["Semua", "Uang Masuk", "Uang Keluar"])
                
                # Input pencarian teks
                search_query = c4.text_input("🔍 Cari Keperluan...", placeholder="Ketik sesuatu...")

            # --- LOGIKA FILTERING ---
            df_filtered = df_admin_trx.copy()
            
            if f_user != "Semua":
                df_filtered = df_filtered[df_filtered['PIC'] == f_user]
            
            if f_ev != "Semua":
                df_filtered = df_filtered[df_filtered['NAMA_EVENT'] == f_ev]
            
            if f_tipe == "Uang Masuk":
                df_filtered = df_filtered[df_filtered['UANG_MASUK'] > 0]
            elif f_tipe == "Uang Keluar":
                df_filtered = df_filtered[df_filtered['UANG_KELUAR'] > 0]
                
            if search_query:
                df_filtered = df_filtered[df_filtered['KEPERLUAN'].str.contains(search_query, case=False, na=False)]

            # --- TAMPILAN RINGKASAN FILTER ---
            st.write("---")
            total_m = df_filtered['UANG_MASUK'].sum()
            total_k = df_filtered['UANG_KELUAR'].sum()
            saldo_f = total_m - total_k
            
            m1, m2, m3, m4 = st.columns(4)
            m1.markdown(f"<div style='padding:10px; background:#012a4a; border-radius:10px; border-left:5px solid #00ff99'><b>Total Masuk:</b><br>Rp {total_m:,.0f}</div>", unsafe_allow_html=True)
            m2.markdown(f"<div style='padding:10px; background:#012a4a; border-radius:10px; border-left:5px solid #ff4b4b'><b>Total Keluar:</b><br>Rp {total_k:,.0f}</div>", unsafe_allow_html=True)
            m3.markdown(f"<div style='padding:10px; background:#012a4a; border-radius:10px; border-left:5px solid #00d4ff'><b>Saldo Akhir:</b><br>Rp {saldo_f:,.0f}</div>", unsafe_allow_html=True)
            m4.markdown(f"<div style='padding:10px; background:#012a4a; border-radius:10px; border-left:5px solid #ffc107'><b>Jumlah Data:</b><br>{len(df_filtered)} Transaksi</div>", unsafe_allow_html=True)
            
            st.write("")

            # --- TABEL UTAMA (RAHESIA TAMPILAN RAPI) ---
            # Mengatur urutan kolom agar lebih enak dibaca dan menyembunyikan index
            st.dataframe(
                df_filtered[['TANGGAL', 'PIC', 'NAMA_EVENT', 'KEPERLUAN', 'UANG_MASUK', 'UANG_KELUAR', 'ID_TRANSAKSI', 'BUKTI_LINK']],
                use_container_width=True,
                hide_index=True,
                column_config={
                    "UANG_MASUK": st.column_config.NumberColumn("Masuk (Rp)", format="Rp %d"),
                    "UANG_KELUAR": st.column_config.NumberColumn("Keluar (Rp)", format="Rp %d"),
                    "BUKTI_LINK": st.column_config.LinkColumn("Bukti"),
                    "TANGGAL": st.column_config.DatetimeColumn("Waktu", format="DD/MM/YY HH:mm")
                }
            )
            # --- BAGIAN EKSPOR LAPORAN (DIPINDAH KE SINI) ---
            st.markdown("### 🖨️ Cetak Laporan")
            st.info("Gunakan data di atas (berdasarkan filter) untuk mencetak laporan resmi.")
        
            c_exp1, c_exp2 = st.columns(2)
            nama_file_laporan = f"Laporan_PTT_{datetime.now().strftime('%Y%m%d')}"
        
            with c_exp1:
                excel_bin = to_excel(df_filtered)
                st.download_button(
                    label="📥 Download Excel (.xlsx)",
                    data=excel_bin,
                    file_name=f"{nama_file_laporan}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

        with c_exp2:
            try:
                # Kita pakai judul dinamis berdasarkan filter yang sedang aktif
                judul_pdf = f"LAPORAN KEUANGAN: {f_ev if 'f_ev' in locals() else 'SEMUA'}"
                pdf_bin = create_pdf(df_filtered, judul_pdf)
                st.download_button(
                    label="📄 Download PDF (.pdf)",
                    data=pdf_bin,
                    file_name=f"{nama_file_laporan}.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"Gagal generate PDF: {e}")
            
            
        with tab_user:
            st.subheader("👥 Manajemen Hak Akses Anggota")
            df_u = get_data("KEANGGOTAAN")
            
            # Tambahkan kolom pencarian anggota agar lebih cepat
            search_user = st.text_input("🔍 Cari Nama atau Username Anggota...")
            if search_user:
                df_u = df_u[df_u.astype(str).apply(lambda x: x.str.contains(search_user, case=False)).any(axis=1)]

            st.write("")

            # Loop untuk menampilkan setiap anggota dalam kotak (Expander)
            for i, row in df_u.iterrows():
                # Warna label status
                stat_emoji = "🟢" if row['STATUS'] == 'Aktif' else "🔴"
                
                with st.expander(f"{stat_emoji} {row['NAMA_IC']} (@{row['USERNAME']}) — Role: {row['ROLE']}"):
                    col_edit1, col_edit2, col_edit3 = st.columns([2, 2, 1])
                    
                    # 1. Menu Ganti Role
                    list_role = ["Anggota", "Bendahara", "Sekretaris", "Admin"]
                    current_role_idx = list_role.index(row['ROLE']) if row['ROLE'] in list_role else 0
                    new_role = col_edit1.selectbox(
                        f"Update Role untuk {row['USERNAME']}", 
                        list_role, 
                        index=current_role_idx,
                        key=f"role_sel_{i}"
                    )
                    
                    # 2. Menu Ganti Status
                    list_status = ["Aktif", "Nonaktif"]
                    current_stat_idx = list_status.index(row['STATUS']) if row['STATUS'] in list_status else 0
                    new_stat = col_edit2.selectbox(
                        f"Update Status Akun", 
                        list_status, 
                        index=current_stat_idx,
                        key=f"stat_sel_{i}"
                    )
                    
                    # 3. Tombol Eksekusi
                    st.write("") # Memberi jarak
                    btn_save, btn_del = st.columns(2)
                    
                    if btn_save.button(f"💾 Simpan Perubahan", key=f"save_u_{i}"):
                        df_u.at[i, 'ROLE'] = new_role
                        df_u.at[i, 'STATUS'] = new_stat
                        write_data("KEANGGOTAAN", df_u)
                        
                        log_act(user['NAMA_IC'], "MANAJEMEN USER", f"Update {row['USERNAME']} -> {new_role} ({new_stat})")
                        st.success(f"Berhasil memperbarui data {row['NAMA_IC']}!")
                        time.sleep(1)
                        st.rerun()
                    
                    # Tombol Hapus dengan warna merah (type="secondary" di streamlit sering muncul abu-abu, kita beri konfirmasi)
                    if btn_del.button(f"🗑️ Hapus Anggota", key=f"del_u_{i}"):
                        if row['USERNAME'] == user['USERNAME']:
                            st.error("Kamu tidak bisa menghapus akunmu sendiri!")
                        else:
                            df_u = df_u.drop(i)
                            write_data("KEANGGOTAAN", df_u)
                            log_act(user['NAMA_IC'], "HAPUS USER", f"Menghapus user {row['USERNAME']}")
                            st.warning(f"User {row['NAMA_IC']} telah dihapus dari sistem.")
                            time.sleep(1)
                            st.rerun()

    # --- LOG ---
    elif menu == "Log":
        st.title("📜 Log Aktivitas Sistem")
        df_log = get_data("LOG_AKTIVITAS")
        
        # ILOC[::-1] buat balik urutan: Yang baru di GSheets (paling bawah) jadi paling atas di App
        st.dataframe(df_log.iloc[::-1], use_container_width=True, hide_index=True)
        
