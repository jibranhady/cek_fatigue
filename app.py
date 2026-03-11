from flask import Flask, render_template, request
import pandas as pd
import os

app = Flask(__name__)

# =========================
# PATH FILE
# =========================
UPLOAD_PATH = "upload.xlsx"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
df_raw = pd.read_excel(os.path.join(BASE_DIR, "Rawdata.xlsx"))


@app.route("/", methods=["GET", "POST"])
def index():

    hasil = None

    # =========================
    # UPLOAD FILE LAPORAN
    # =========================
    if request.method == "POST" and "file" in request.files:

        file = request.files["file"]

        if file.filename != "":
            file.save(UPLOAD_PATH)
            hasil = "✅ File laporan berhasil diupload"

    # =========================
    # BULK CEK RAW DATA
    # =========================
    if request.method == "POST" and "raw" in request.form:

        if not os.path.exists(UPLOAD_PATH):
            hasil = "❌ Upload laporan dulu"
        else:

            raw_text = request.form["raw"]

            if not raw_text.strip():
                hasil = "❌ Tidak ada input"
            else:

                # load file event tiap request (anti crash Railway)
                df_event = pd.read_excel(
                    UPLOAD_PATH,
                    usecols=["KODE KENDARAAN", "WAKTU KEJADIAN", "WAKTU KE SERVER GABUNGAN"]
                )

                df_event["WAKTU KEJADIAN"] = pd.to_datetime(df_event["WAKTU KEJADIAN"])
                df_event["WAKTU KE SERVER GABUNGAN"] = pd.to_datetime(df_event["WAKTU KE SERVER GABUNGAN"])

                df_event["ANGKA_UNIT"] = df_event["KODE KENDARAAN"].astype(str).str.extract(r"(\d+)")
                df_event["JAM"] = df_event["WAKTU KEJADIAN"].dt.strftime("%H%M%S")
                df_event["TANGGAL"] = df_event["WAKTU KEJADIAN"].dt.strftime("%Y%m%d")

                rows = []

                for raw in raw_text.splitlines():

                    raw = raw.strip().upper()
                    if not raw:
                        continue

                    try:
                        bagian1, tanggal, jam = raw.split("_")
                        unit_raw, pelanggaran = bagian1.split("-")

                        jam_dt = pd.to_datetime(jam, format="%H%M%S") - pd.Timedelta(hours=1)
                        jam_final = jam_dt.strftime("%H%M%S")

                        cek_unit = df_raw[df_raw["deviceid"] == unit_raw]

                        if cek_unit.empty:
                            rows.append([raw, "❌ Unit tidak ditemukan", "", "", ""])
                            continue

                        nama_unit = cek_unit.iloc[0]["unitno"]
                        angka_unit = ''.join(filter(str.isdigit, nama_unit))

                        cari = df_event[
                            (df_event["ANGKA_UNIT"] == angka_unit) &
                            (df_event["JAM"] == jam_final) &
                            (df_event["TANGGAL"] == tanggal)
                        ]

                        if cari.empty:
                            rows.append([raw, nama_unit, pelanggaran, "❌ Tidak ditemukan", ""])
                        else:
                            row = cari.iloc[0]
                            rows.append([
                                raw,
                                nama_unit,
                                pelanggaran,
                                row["WAKTU KEJADIAN"],
                                row["WAKTU KE SERVER GABUNGAN"]
                            ])

                    except:
                        rows.append([raw, "❌ Format salah", "", "", ""])

                hasil = rows

    return render_template("index.html", hasil=hasil)


if __name__ == "__main__":
    app.run()
