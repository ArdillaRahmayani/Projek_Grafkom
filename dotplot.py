import pandas as pd  # Untuk membaca dan mengolah data
import plotly.graph_objects as go  # Untuk membuat visualisasi
import numpy as np  # Untuk operasi numerik

# Baca file Excel, lewati 3 baris pertama
file_path = 'dataset.xlsx'
df = pd.read_excel(file_path, skiprows=3)  # Skip 3 baris pertama

# Ambil NIM mahasiswa dari kolom pertama
nim = df.iloc[:, 0].astype(str).tolist()  # Kolom ke-1 (indeks 0)

# Ambil kolom untuk Tugas, UTS, UAS (kolom ke-5, 7, 10)
tugas = df.iloc[:, 4].values  # Kolom ke-5 (indeks 4)
uts = df.iloc[:, 6].values    # Kolom ke-7 (indeks 6)
uas = df.iloc[:, 9].values    # Kolom ke-10 (indeks 9)

# Buat figure/ kanvas kosong untuk visualisasi
fig = go.Figure()

# Fungsi untuk menentukan warna berdasarkan nilai
def get_color(nilai):
    if nilai < 60:  # Nilai rendah
        return 'red'
    elif 60 <= nilai < 80:  # Nilai sedang
        return 'orange'
    else:  # Nilai tinggi
        return 'green'

# Tambahkan titik-titik untuk nilai Tugas
tugas_colors = [get_color(val) for val in tugas] # Tentukan warna untuk setiap nilai
fig.add_trace(go.Scatter(
    x=nim, # NIM di sumbu X
    y=tugas, # NIM di sumbu Y
    mode='markers', # Tampilkan sebagai titik
    marker=dict(
        size=12, # Ukuran titik
        color=tugas_colors,  # Warna sesuai nilai
        symbol='circle', # Bentuk lingkaran
        line=dict(color='black', width=1) # Garis tepi titik
    ),
    name='Nilai Tugas', # Label
    text=[f"Tugas: {val:.1f}" for val in tugas], # Teks saat kursor diarahkan ke titik
    hoverinfo='text+x' # Teks saat kursor diarahkan ke titik
))

# Tambahkan titik-titik untuk nilai UTS
uts_colors = [get_color(val) for val in uts]
fig.add_trace(go.Scatter(
    x=nim,
    y=uts,
    mode='markers',
    marker=dict(
        size=12,
        color=uts_colors,
        symbol='square',
        line=dict(
            color='black',
            width=1
        )
    ),
    name='Nilai UTS',
    text=[f"UTS: {val:.1f}" for val in uts],
    hoverinfo='text+x'
))

# Tambahkan titik-titik untuk nilai UAS
uas_colors = [get_color(val) for val in uas]
fig.add_trace(go.Scatter(
    x=nim,
    y=uas,
    mode='markers',
    marker=dict(
        size=12,
        color=uas_colors,
        symbol='triangle-up',
        line=dict(
            color='black',
            width=1
        )
    ),
    name='Nilai UAS',
    text=[f"UAS: {val:.1f}" for val in uas],
    hoverinfo='text+x'
))

# Tambahkan garis horizontal untuk batas grade
# Tambahkan garis horizontal pada nilai 60 (batas nilai rendah-sedang)
fig.add_shape(
    type="line",
    x0=-0.5,  # Mulai dari kiri area plot
    y0=60,    # Nilai Y = 60
    x1=len(nim)-0.5,  # Sampai kanan area plot
    y1=60,    # Tetap pada Y = 60
    line=dict(color="gray", width=1, dash="dash")  # Garis abu-abu putus-putus
)

# Tambahkan garis horizontal pada nilai 80 (batas nilai sedang-tinggi)
fig.add_shape(
    type="line",
    x0=-0.5,
    y0=80,
    x1=len(nim)-0.5,
    y1=80,
    line=dict(
        color="gray",
        width=1,
        dash="dash",
    )
)

# Tambahkan anotasi/teks untuk kategori nilai rendah
fig.add_annotation(
    x=len(nim)-0.5, # Posisi X di ujung kanan
    y=50, # Posisi Y pada nilai 50
    xref="x",
    yref="y",
    text="Low: <60", # Teks yang ditampilkan
    showarrow=False, # Tanpa panah
    font=dict(size=12, color="red"), # Format teks
    align="right",  # Rata kanan
)

# Tambahkan anotasi/teks untuk kategori nilai sedang
fig.add_annotation(
    x=len(nim)-0.5,
    y=70,
    xref="x",
    yref="y",
    text="Medium: 60-79",
    showarrow=False,
    font=dict(
        size=12,
        color="orange"
    ),
    align="right",
)

# Tambahkan anotasi/teks untuk kategori nilai tinggi
fig.add_annotation(
    x=len(nim)-0.5,
    y=90,
    xref="x",
    yref="y",
    text="High: ≥80",
    showarrow=False,
    font=dict(
        size=12,
        color="green"
    ),
    align="right",
)

# Atur tampilan keseluruhan grafik
fig.update_layout(
    title="Dot Plot Nilai Mahasiswa per Nim dengan Kategori Grade Low, Medium, High (Berdasar Data Nilai Tugas, UTS, UAS)",
    xaxis_title="NIM",  # Label sumbu X
    yaxis_title="Nilai",  # Label sumbu Y
    xaxis=dict(tickangle=90,), # Putar label NIM 90 derajat
    yaxis=dict(
        range=[0, 105],  # Rentang nilai 0-105
        gridcolor='lightgray',  # Warna grid
        gridwidth=0.5,  # Ketebalan grid
        griddash='dot'  # Pola grid (titik-titik)
    ),
    height=600,  # Tinggi grafik
    width=1000,  # Lebar grafik
    legend=dict(  # Pengaturan legenda
        orientation="h",  # Horizontal
        yanchor="bottom",
        y=1.02,
        xanchor="center",
        x=0.5
    )
)

# Simpan sebagai file HTML interaktif
fig.write_html("dotplot.html")

# Tampilkan grafik
fig.show()