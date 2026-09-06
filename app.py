import streamlit as st
import requests
import datetime

# Gunakan Token API asli Anda
BOT_TOKEN = "8893067990:AAFbbn0xxxXGyCYq5MpV760481spUMONqIg"

st.set_page_config(page_title="Dashboard Report Telegram", layout="centered")
st.title("📡 Live Report PM Lapangan")

# 1. Konfigurasi Filter Laporan (Bisa ditambah dengan kode area seperti "tarakan", "pangkalan bun")
KATA_KUNCI_REPORT = ["done", "pm", "plk", "bkeminting"] 

def get_telegram_updates():
    url = f"https://api.telegram.org/bot{BOT_TOKEN}/getUpdates"
    try:
        response = requests.get(url)
        data = response.json()
        if data.get("ok"):
            return data.get("result", [])
    except Exception as e:
        st.error(f"Gagal terhubung ke Telegram: {e}")
    return []

def get_image_url(file_id):
    url = f"https://api.telegram.org/bot{BOT_TOKEN}/getFile?file_id={file_id}"
    try:
        response = requests.get(url).json()
        if response.get("ok"):
            file_path = response["result"]["file_path"]
            return f"https://api.telegram.org/file/bot{BOT_TOKEN}/{file_path}"
    except Exception:
        return None
    return None

updates = get_telegram_updates()
laporan_valid = []

# 2. Proses Filtering Obrolan vs Laporan
if updates:
    for item in updates:
        msg = item.get("message") or item.get("channel_post")
        if not msg:
            continue
            
        # Ubah teks menjadi huruf kecil semua untuk mempermudah pengecekan
        text = str(msg.get("text") or msg.get("caption") or "").lower()
        
        # Cek apakah pesan mengandung salah satu kata kunci di atas
        if any(kata in text for kata in KATA_KUNCI_REPORT):
            laporan_valid.append(msg)

# 3. Menampilkan Akumulasi Laporan Tim
st.metric(label="Total Report PM Masuk", value=len(laporan_valid), delta="Data ditarik hari ini")
st.divider()

if not laporan_valid:
    st.info("Belum ada laporan PM yang sesuai kriteria hari ini.")
else:
    # Membalik urutan agar PM paling akhir (Last PM) berada di atas
    for msg in reversed(laporan_valid):
        date_unix = msg.get("date")
        date_str = datetime.datetime.fromtimestamp(date_unix).strftime('%d/%m/%Y %H:%M:%S')
        sender = msg.get("from", {}).get("first_name", "Tim")
        
        # Menampilkan teks asli (huruf besar/kecil tetap sesuai aslinya)
        text_asli = msg.get("text") or msg.get("caption") or "*(Tanpa keterangan)*"
        
        with st.container():
            st.markdown(f"**Pelapor:** {sender} | 🕒 {date_str}")
            st.text(text_asli)
            
            if "photo" in msg:
                # Ambil foto dengan resolusi tertinggi
                file_id = msg["photo"][-1]["file_id"]
                img_url = get_image_url(file_id)
                
                if img_url:
                    try:
                        # 4. Perbaikan TypeError Streamlit Cloud
                        st.image(img_url, use_container_width=True)
                    except Exception:
                        st.warning("⚠️ Gagal merender gambar ini dari server Telegram.")
            
            st.markdown("---")

if st.button("🔄 Refresh Data"):
    st.rerun()
