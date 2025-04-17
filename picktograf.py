import pandas as pd  # Untuk membaca dan mengolah data
import plotly.graph_objects as go  # Untuk membuat visualisasi
import numpy as np  # Untuk operasi numerik

# Baca file Excel, lewati 3 baris pertama
file_path = 'dataset.xlsx'
df = pd.read_excel(file_path, skiprows=3)

# Ambil kolom ke-1 sebagai NIM dan kolom ke-13 sebagai Total Nilai Akhir
ids = df.iloc[:, 0].astype(str).tolist()     # Kolom ke-1 (Nim)
nilai_akhir = df.iloc[:, 13].astype(float).tolist()  # Kolom ke-13 (Nilai Akhir)

# Buat kanvas kosong
fig = go.Figure()

# Buat data untuk visualisasi
for i, (id_mahasiswa, nilai) in enumerate(zip(ids, nilai_akhir)):
    # Buat skala nilai dari 0 sampai 100 dengan interval 5
    y_values = list(range(0, 105, 5))  # 0, 5, 10, ..., 100

    # Tentukan warna untuk setiap titik
    # Jika nilai mahasiswa ≥ nilai titik, warnai orange, jika tidak abu-abu
    colors = ['orange' if nilai >= y else 'lightgray' for y in y_values]
    
    # Tambahkan kumpulan titik untuk setiap mahasiswa
    fig.add_trace(go.Scatter(
        x=[i] * len(y_values),  # Posisi X tetap (satu kolom per mahasiswa)
        y=y_values,             # Posisi Y sesuai skala nilai
        mode='markers',         # Tampilkan sebagai titik
        marker=dict(
            color=colors,       # Warna sesuai kondisi
            size=15,            # Ukuran titik
            line=dict(color='gray', width=0.5)  # Garis tepi titik
        ),
        showlegend=False,
        hoverinfo='text',
        text=[f'Nim: {id_mahasiswa}, Nilai: {nilai}, Level: {y}' for y in y_values]  # Info saat hover
    ))

# Atur tampilan keseluruhan
fig.update_layout(
    title="Grafik Piktograf Total Nilai Akhir per Nim",
    xaxis=dict(
        title="NIM",
        tickmode='array',
        tickvals=list(range(len(ids))),  # Posisi label
        ticktext=ids,                    # Teks label (ID/NIM)
        tickangle=45                     # Putar label 45 derajat
    ),
    yaxis=dict(
        title="Total Nilai Akhir",
        range=[-5, 105],                 # Rentang nilai
        tickmode='array',
        tickvals=list(range(0, 110, 10)), # Posisi label nilai
        ticktext=[str(i) for i in range(0, 110, 10)]  # Teks label nilai
    ),
    height=600,
    width=1000
)

# Simpan sebagai file HTML interaktif
fig.write_html("pictograf.html")

# Tampilkan grafik
fig.show()