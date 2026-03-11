from flask import Flask, render_template, request, send_file
import pandas as pd
import os
from io import BytesIO

app = Flask(__name__)

# =========================
# PATH
# =========================
UPLOAD_PATH = "upload.xlsx"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
df_raw = pd.read_excel(os.path.join(BASE_DIR, "Rawdata.xlsx"))

# simpan hasil terakhir (buat export)
last_rows = []


@app.route("/", methods=["GET", "POST"])
def index():
    global last_rows

    hasil = None

    # =========================
    # UPLOAD FILE
    # =========================
    if request.method == "POST" and "file" in request.files:

        file = request.files["file"]

        if file.filename != "":
            file.save(UPLOAD_PATH)
            hasil = "✅ File laporan berhasil diupload"

    # =========================
    # BULK CEK RAW
    # =========================
    if request.method == "POST" and "raw" in request.form:

        if not os.path.exists(UPLOAD_PATH):
            hasil = "❌ Upload laporan dulu"

        else:

            raw_text = request.form["raw"]

            if not raw_text.strip():
                hasil = "❌ Tidak ada input"

            else:

                # =========================
                # LOAD EVENT DATA
                # =========================
                df_event = pd.read_excel(
                    UPLOAD_PATH,
                    usecols=[
                        "KODE KENDARAAN",
                        "WAKTU KEJADIAN",
                        "WAKTU KE SERVER GABUNGAN",
                        "INTERVENSI - STATUS CONTEXT"
                    ]
                )

                df_event["WAKTU KEJADIAN"] = pd.to_datetime(df_event["WAKTU KEJADIAN"])
                df_event["WAKTU KE SERVER GABUNGAN"] = pd.to_datetime(df_event["WAKTU KE SERVER GABUNGAN"])

                df_event["ANGKA_UNIT"] = df_event["KODE KENDARAAN"].astype(str).str.extract(r"(\d+)")
                df_event["JAM"] = df_event["WAKTU KEJADIAN"].dt.strftime("%H%M%S")
                df_event["TANGGAL"] = df_event["WAKTU KEJADIAN"].dt.strftime("%Y%m%d")

                rows = []

                # =========================
                # LOOP RAW
                # =========================
                for raw in raw_text.splitlines():

                    raw = raw.strip().upper()
                    if not raw:
                        continue

                    try:
                        bagian1, tanggal, jam = raw.split("_")
                        unit_raw, pelanggaran = bagian1.split("-")

                        # kurangi 1 jam
                        jam_dt = pd.to_datetime(jam, format="%H%M%S") - pd.Timedelta(hours=1)
                        jam_final = jam_dt.strftime("%H%M%S")

                        # =========================
                        # CEK UNIT
                        # =========================
                        cek_unit = df_raw[df_raw["deviceid"] == unit_raw]

                        if cek_unit.empty:
                            rows.append([raw, "❌ Unit tidak ditemukan", "", "", "", ""])
                            continue

                        nama_unit = cek_unit.iloc[0]["unitno"]
                        angka_unit = ''.join(filter(str.isdigit, nama_unit))

                        # =========================
                        # CARI EVENT
                        # =========================
                        cari = df_event[
                            (df_event["ANGKA_UNIT"] == angka_unit) &
                            (df_event["JAM"] == jam_final) &
                            (df_event["TANGGAL"] == tanggal)
                        ]

                        if cari.empty:
                            rows.append([raw, nama_unit, pelanggaran, "❌ Tidak ditemukan", "", ""])
                        else:
                            row = cari.iloc[0]
                            rows.append([
                                raw,
                                nama_unit,
                                pelanggaran,
                                row["WAKTU KEJADIAN"],
                                row["WAKTU KE SERVER GABUNGAN"],
                                row["INTERVENSI - STATUS CONTEXT"]
                            ])

                    except Exception:
                        rows.append([raw, "❌ Format salah", "", "", "", ""])

                hasil = rows
                last_rows = rows  # simpan buat export

    return render_template("index.html", hasil=hasil)


# =========================
# EXPORT EXCEL
# =========================
@app.route("/export")
def export_excel():

    global last_rows

    if not last_rows:
        return "Tidak ada data untuk diexport"

    df_export = pd.DataFrame(last_rows, columns=[
        "RAW",
        "Nama Unit",
        "Pelanggaran",
        "Waktu Kejadian",
        "Masuk Gabungan",
        "Status Context"
    ])

    output = BytesIO()
    df_export.to_excel(output, index=False)
    output.seek(0)

    return send_file(
        output,
        download_name="hasil_cek_raw.xlsx",
        as_attachment=True
    )


if __name__ == "__main__":
    app.run()
