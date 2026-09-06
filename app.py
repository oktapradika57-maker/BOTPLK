import streamlit as st
import requests
import datetime

# Ganti dengan Token Bot Anda yang asli
BOT_TOKEN = "8893067990:AAFbbn0xxxXGyCYq5MpV760481spUMONqIg"

st.set_page_config(page_title="Dashboard Report Telegram", layout="centered")
st.title("📡 Live Report Tim Lapangan")
st.markdown("Menampilkan foto dan laporan PM terbaru dari grup Telegram.")

# Fungsi untuk mengambil pesan terbaru dari Telegram
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

# Fungsi untuk mengubah File ID foto menjadi Link URL yang bisa dirender Streamlit
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

# Menarik data
updates = get_telegram_updates()

if not updates:
    st.info("Belum ada laporan baru di Telegram.")
else:
    # Membalik urutan agar pesan paling baru (last PM) muncul di paling atas
    for item in reversed(updates):
        msg = item.get("message") or item.get("channel_post")
        if not msg:
            continue
            
        # Mengambil informasi dasar
        date_unix = msg.get("date")
        date_str = datetime.datetime.fromtimestamp(date_unix).strftime('%d/%m/%Y %H:%M:%S')
        sender = msg.get("from", {}).get("first_name", "Tim")
        
        # Mengambil teks laporan atau caption foto
        text = msg.get("text") or msg.get("caption") or "*(Hanya mengirim file tanpa keterangan)*"
        
        # Membuat UI Card di Streamlit
        with st.container():
            st.markdown(f"**Pelapor:** {sender} | 🕒 {date_str}")
            
            # Jika laporan berupa teks/checklist
            st.text(text)
            
            # Jika laporan mengandung foto, cari foto dengan resolusi tertinggi (array terakhir)
            if "photo" in msg:
                file_id = msg["photo"][-1]["file_id"]
                img_url = get_image_url(file_id)
                if img_url:
                    st.image(img_url, use_column_width=True)
            
            st.divider()

# Tombol untuk memuat ulang data terbaru
if st.button("🔄 Refresh Data"):
    st.rerun()
