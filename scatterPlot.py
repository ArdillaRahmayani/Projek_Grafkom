import pandas as pd  # Untuk membaca dan mengolah data
import plotly.graph_objects as go  # Untuk membuat visualisasi
import numpy as np  # Untuk operasi numerik

# Baca file Excel, lewati 3 baris pertama
file_path = 'dataset.xlsx'
df = pd.read_excel(file_path, skiprows=3)  # Skip 3 baris pertama

# Pisahkan data menjadi kategori (kolom kiri) dan nilai capaian (kolom kanan)
data_kategori = df.iloc[:, [14, 16, 18, 20]].values  # 4 kolom sebagai kategori
data_nilai = df.iloc[:, [15, 17, 19, 21]].values     # 4 kolom sebagai nilai capaian

# Ubah ke dalam array numpy
data_kategori = np.array(data_kategori)
data_nilai = np.array(data_nilai)

# Buat kanvas kosong
fig = go.Figure()

# Warna untuk setiap CPMK
colors = ['rgba(31, 119, 180, 0.8)', 'rgba(255, 127, 14, 0.8)', 
          'rgba(44, 160, 44, 0.8)', 'rgba(214, 39, 40, 0.8)']

# Daftar nama CPMK baru
cpmk_labels = ['CPMK 012', 'CPMK 031', 'CPMK 071', 'CPMK 072']

# Loop untuk memetakan setiap pasangan data
# Untuk setiap jenis CPMK
for i in range(data_kategori.shape[1]):
    # Tambahkan titik-titik scatter
    fig.add_trace(go.Scatter(
        x=data_kategori[:, i],  # Nilai X dari data kategori
        y=data_nilai[:, i],     # Nilai Y dari data nilai
        mode='markers',         # Tampilkan sebagai titik
        marker=dict(
            size=12,            # Ukuran titik
            color=colors[i],    # Warna sesuai CPMK
            line=dict(color='white', width=0.5)  # Garis tepi titik
        ),
        name=cpmk_labels[i],    # Label untuk legenda dengan nama baru
        text=[f'{cpmk_labels[i]}: ({x:.1f}, {y:.1f})' for x, y in zip(data_kategori[:, i], data_nilai[:, i])],
        hoverinfo='text'        # Info saat hover
    ))

# Atur tampilan keseluruhan
fig.update_layout(
    title="Scatter Plot Nilai Capaian CPMK",
    xaxis=dict(
        title="Capaian CPMK (%)",  # Label sumbu X
        gridcolor='lightgray',     # Warna grid
        gridwidth=0.5,             # Ketebalan grid
        griddash='dot'             # Pola grid (titik-titik)
    ),
    yaxis=dict(
        title="Capaian CPMK (Nilai)",  # Label sumbu Y
        gridcolor='lightgray',
        gridwidth=0.5,
        griddash='dot'
    ),
    height=600,
    width=800,
    legend=dict(
        orientation="h",            # Legenda horizontal
        yanchor="bottom",
        y=1.02,
        xanchor="center",
        x=0.5
    )
)

# Simpan sebagai file HTML interaktif
fig.write_html("scatter_plot.html")

# Tampilkan grafik
fig.show()