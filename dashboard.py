import io
import os
import textwrap
import re
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import numpy as np
import openpyxl
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import (
    HRFlowable,
    Image as RLImage,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)
import streamlit as st

# ==========================================
# 1. KONFIGURASI HALAMAN & TEMA PLN
# ==========================================
st.set_page_config(
    page_title="Dashboard Evaluasi Ulasan PLN Mobile",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
<style>
    .main-header {
        background: linear-gradient(90deg, #0A3960 0%, #00A4E4 100%);
        padding: 22px 26px;
        border-radius: 10px;
        color: white;
        margin-bottom: 20px;
    }
    .main-header h1 {
        color: white;
        margin: 0;
        font-size: 26px;
        font-weight: 700;
    }
    .main-header p {
        color: #E0F2FE;
        margin: 5px 0 0 0;
        font-size: 14px;
    }
    div[data-testid="stMetricValue"] {
        font-size: 26px !important;
        font-weight: 700 !important;
        color: #0A3960 !important;
    }
</style>
""",
    unsafe_allow_html=True,
)


# ==========================================
# 2. FUNGSI LOAD DATA & PREPROCESSING
# ==========================================
@st.cache_data(show_spinner=False)
def load_data(file_source):
  try:
    df = pd.read_excel(file_source, sheet_name="Detail Ulasan Negatif")
  except Exception:
    xls = pd.ExcelFile(file_source)
    sheet_name = (
        "Detail Ulasan Negatif"
        if "Detail Ulasan Negatif" in xls.sheet_names
        else xls.sheet_names[0]
    )
    df = pd.read_excel(file_source, sheet_name=sheet_name)

  df.columns = df.columns.str.strip()

  # Datetime & Periode
  if "at" in df.columns:
    df["at"] = pd.to_datetime(df["at"], errors="coerce")
    df["Tahun"] = df["at"].dt.year
    df["Bulan_Tahun"] = df["at"].dt.to_period("M").astype(str)
  else:
    df["Tahun"] = 2026
    df["Bulan_Tahun"] = "2026-01"

  if "repliedAt" in df.columns:
    df["repliedAt"] = pd.to_datetime(df["repliedAt"], errors="coerce")

  # Perhitungan SLA Respon
  if "at" in df.columns and "repliedAt" in df.columns:
    df["Response_Hours"] = (
        df["repliedAt"] - df["at"]
    ).dt.total_seconds() / 3600.0

    def assign_sla(h):
      if pd.isna(h):
        return "Belum Direspon"
      elif h <= 1:
        return "< 1 Jam"
      elif h <= 6:
        return "1 - 6 Jam"
      elif h <= 24:
        return "6 - 24 Jam"
      else:
        return "> 24 Jam"

    df["SLA_Bucket"] = df["Response_Hours"].apply(assign_sla)
  else:
    df["SLA_Bucket"] = "Tidak Tersedia"

  # Standardisasi Klaster Versi Aplikasi
  def group_version(v):
    if pd.isna(v) or str(v).strip() in ["", "nan", "-", "None"]:
      return "Versi 5.x & Tanpa Versi"
    v_str = str(v).strip()
    if "8.1.3" in v_str:
      return "Versi 8.1.3"
    elif "8.1.2" in v_str:
      return "Versi 8.1.2"
    elif "8.1.1" in v_str or "08.01.01" in v_str:
      return "Versi 8.1.1"
    elif "8.1.0" in v_str or "08.01.00" in v_str:
      return "Versi 8.1.0"
    elif v_str.startswith("5."):
      return "Versi 5.x & Tanpa Versi"
    return "Versi Lainnya"

  if "reviewCreatedVersion" in df.columns:
    df["Versi_Group"] = df["reviewCreatedVersion"].apply(group_version)
  else:
    df["Versi_Group"] = "Versi 5.x & Tanpa Versi"

  # Ekstraksi Kata Pemicu (Murni Kata Kunci, Menghapus Non-Keyword)
  def extract_trigger(text):
    if pd.isna(text):
      return None
    m = re.search(
        r"Kata/frasa pemicu terdeteksi:\s*(.+)", str(text), re.IGNORECASE
    )
    if m:
      kw = m.group(1).lower().strip()
      if kw not in [
          "non-keyword",
          "non keyword",
          "lainnya",
          "tidak ditemukan kata kunci yang cukup kuat untuk menentukan domain",
          "ulasan tidak menjelaskan kendala secara spesifik",
      ]:
        return kw
    return None

  if "Dasar Klasifikasi" in df.columns:
    df["Kata_Pemicu_Clean"] = df["Dasar Klasifikasi"].apply(extract_trigger)
  else:
    df["Kata_Pemicu_Clean"] = None

  return df


# ==========================================
# 3. GENERATOR LAPORAN RESMI (PDF LENGKAP)
# ==========================================
def generate_full_pdf_report(df_source):
  pdf_buffer = io.BytesIO()
  doc = SimpleDocTemplate(
      pdf_buffer,
      pagesize=A4,
      rightMargin=32,
      leftMargin=32,
      topMargin=32,
      bottomMargin=32,
  )

  styles = getSampleStyleSheet()

  # Tipografi Laporan Eksekutif
  title_style = ParagraphStyle(
      "DocTitle",
      parent=styles["Heading1"],
      fontName="Helvetica-Bold",
      fontSize=18,
      leading=22,
      textColor=colors.HexColor("#0A3960"),
  )
  subtitle_style = ParagraphStyle(
      "DocSubTitle",
      parent=styles["Normal"],
      fontName="Helvetica",
      fontSize=9.5,
      leading=13,
      textColor=colors.HexColor("#555555"),
  )
  h2_style = ParagraphStyle(
      "Heading2Custom",
      parent=styles["Heading2"],
      fontName="Helvetica-Bold",
      fontSize=12,
      leading=15,
      textColor=colors.HexColor("#0A3960"),
      spaceBefore=8,
      spaceAfter=4,
  )
  insight_style = ParagraphStyle(
      "InsightText",
      parent=styles["Normal"],
      fontName="Helvetica",
      fontSize=8.5,
      leading=12.5,
      textColor=colors.HexColor("#2C3E50"),
  )
  body_style = ParagraphStyle(
      "BodyCustom",
      parent=styles["BodyText"],
      fontName="Helvetica",
      fontSize=8.5,
      leading=12.5,
      textColor=colors.HexColor("#2C3E50"),
  )

  # Helper untuk konversi chart matplotlib ke Flowable ReportLab (Bebas Cut-off)
  def make_clean_image(fig, w=525, h=215):
    buf = io.BytesIO()
    plt.savefig(
        buf, format="png", dpi=200, bbox_inches="tight", facecolor="white"
    )
    plt.close(fig)
    buf.seek(0)
    return RLImage(buf, width=w, height=h)

  story = []
  tot = len(df_source)

  # ========================================================
  # HALAMAN 1: COVER RESMI, 5 METRICS KPI & DEFINISI INDIKATOR
  # ========================================================
  story.append(
      Paragraph(
          "PT PLN (PERSERO) - DIVISI MANAJEMEN DIGITAL (DIV MDG)",
          subtitle_style,
      )
  )
  story.append(Spacer(1, 2))
  story.append(
      Paragraph(
          "Laporan Eksekutif Evaluasi & Analisis Klasifikasi Ulasan Negatif PLN"
          " Mobile",
          title_style,
      )
  )
  story.append(Spacer(1, 2))
  story.append(
      Paragraph(
          "Evaluasi Komprehensif: 5 Metrik Kinerja, 11 Visualisasi Data, dan"
          " Rekomendasi Strategis | Tahun 2026",
          subtitle_style,
      )
  )
  story.append(Spacer(1, 6))
  story.append(
      HRFlowable(
          width="100%",
          thickness=2.5,
          color=colors.HexColor("#00A4E4"),
          spaceBefore=1,
          spaceAfter=10,
      )
  )

  app_cnt = (df_source["Klasifikasi Utama"] == "Kendala Aplikasi").sum()
  biz_cnt = (df_source["Klasifikasi Utama"] == "Kendala Proses Bisnis").sum()
  has_reply = (
      df_source["replyContent"].notna()
      & (df_source["replyContent"].str.strip() != "")
  ).sum()
  reply_rate = (has_reply / tot * 100) if tot > 0 else 0
  median_sla = (
      df_source["Response_Hours"].dropna().median()
      if "Response_Hours" in df_source.columns
      else 0.7
  )

  kpi_table_data = [[
      Paragraph(
          f"<b>Total Ulasan Negatif</b><br/><font size=14 color='#0A3960'><b>{tot:,}</b></font><br/><font size=7 color='#7F8C8D'>Bintang 1 s.d. 3</font>",
          body_style,
      ),
      Paragraph(
          f"<b>Kendala Aplikasi</b><br/><font size=14 color='#C0392B'><b>{app_cnt:,}</b></font><br/><font size=7 color='#C0392B'>Porsi: {app_cnt/tot*100:.1f}%</font>",
          body_style,
      ),
      Paragraph(
          f"<b>Kendala Bisnis</b><br/><font size=14 color='#C0392B'><b>{biz_cnt:,}</b></font><br/><font size=7 color='#C0392B'>Porsi: {biz_cnt/tot*100:.1f}%</font>",
          body_style,
      ),
      Paragraph(
          f"<b>Tingkat Respon CS</b><br/><font size=14 color='#27AE60'><b>{reply_rate:.1f}%</b></font><br/><font size=7 color='#27AE60'>{has_reply:,} Terjawab</font>",
          body_style,
      ),
      Paragraph(
          f"<b>Median SLA Respon</b><br/><font size=14 color='#27AE60'><b>{median_sla:.1f} Jam</b></font><br/><font size=7 color='#27AE60'>Target: < 2 Jam</font>",
          body_style,
      ),
  ]]
  t_kpi = Table(kpi_table_data, colWidths=[106, 106, 106, 106, 106])
  t_kpi.setStyle(
      TableStyle([
          ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F8FAFC")),
          ("BOX", (0, 0), (-1, -1), 1, colors.HexColor("#CBD5E1")),
          ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#E2E8F0")),
          ("TOPPADDING", (0, 0), (-1, -1), 8),
          ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
          ("ALIGN", (0, 0), (-1, -1), "CENTER"),
      ])
  )
  story.append(t_kpi)
  story.append(Spacer(1, 10))

  story.append(
      Paragraph(
          "Penjelasan Metrik KPI & Definisi Indikator Merah/Hijau", h2_style
      )
  )
  story.append(
      Paragraph(
          "Untuk mempermudah pembacaan laporan eksekutif manajemen, arti"
          " indikator pada 5 metrik KPI di atas adalah sebagai berikut:<br/>•"
          f" <b>Total Ulasan Negatif ({tot:,} ulasan)</b>: Akumulasi seluruh"
          " keluhan pengguna bintang 1 s.d. 3 yang terekam pada Google Play"
          " Store.<br/>• <b>Badge Merah pada Kendala Aplikasi"
          f" ({app_cnt/tot*100:.1f}%) dan Kendala Bisnis"
          f" ({biz_cnt/tot*100:.1f}%)</b>: Merupakan <b>persentase porsi"
          " kontribusi dari total ulasan</b> (bukan nilai lonjakan bulanan/MoM)."
          " Diberi highlight merah karena parameter kendala dikonfigurasi"
          " sebagai metrik negatif (beban keluhan operasional yang membutuhkan"
          " penanganan).<br/>• <b>Badge Hijau pada Tingkat Respon CS"
          f" ({reply_rate:.1f}%) dan SLA ({median_sla:.1f} Jam)</b>:"
          " Menunjukkan pencapaian positif tim pelayanan digital PLN 123 yang"
          " berhasil menanggapi mayoritas keluhan pelanggan dengan kecepatan"
          " rata-rata di bawah ambang batas toleransi 2 jam.",
          body_style,
      )
  )

  # ========================================================
  # HALAMAN 2: GRAFIK 1 & GRAFIK 2 (DONUT DENGAN LEGENDA SAMPING)
  # ========================================================
  story.append(PageBreak())
  story.append(
      Paragraph(
          "Bagian I: Proporsi Klasifikasi Utama & Sebaran Versi Aplikasi",
          h2_style,
      )
  )

  # Grafik 1: Donut Klasifikasi Utama (Bebas tumpang tindih)
  fig1, ax1 = plt.subplots(figsize=(6.6, 2.7))
  main_counts = df_source["Klasifikasi Utama"].value_counts()
  colors_main = ["#F39C12", "#95A5A6", "#00A4E4"]
  wedges1, _, autotexts1 = ax1.pie(
      main_counts.values,
      autopct="%1.1f%%",
      startangle=140,
      colors=colors_main[: len(main_counts)],
      wedgeprops=dict(width=0.45, edgecolor="white", linewidth=2),
      pctdistance=0.75,
  )
  plt.setp(autotexts1, size=8.5, weight="bold", color="white")
  ax1.text(
      0,
      0,
      f"{tot:,}\nTotal",
      ha="center",
      va="center",
      fontsize=8.5,
      weight="bold",
      color="#0A3960",
  )
  legend_labels1 = [
      f"{k}: {v:,} ({v/tot*100:.1f}%)"
      for k, v in zip(main_counts.index, main_counts.values)
  ]
  ax1.legend(
      wedges1,
      legend_labels1,
      title="Kategori Utama",
      loc="center left",
      bbox_to_anchor=(1, 0.5),
      fontsize=8,
      title_fontsize=8.5,
      frameon=False,
  )
  ax1.set_title(
      "Grafik 1: Persentase Klasifikasi Utama Ulasan Negatif",
      fontsize=10.5,
      weight="bold",
      color="#0A3960",
      pad=14,
  )
  story.append(make_clean_image(fig1, 520, 180))
  story.append(
      Paragraph(
          "<b>Analisis & Insight Grafik 1:</b> Keluhan Proses Bisnis (38,6%)"
          " lebih dominan dibanding kendala murni teknis aplikasi (29,3%). Hal"
          " ini membuktikan bahwa pengguna memperlakukan PLN Mobile sebagai"
          " gerbang layanan listrik secara riil. Ketidakpuasan pelanggan atas"
          " padam listrik, lambatnya regu lapangan, dan lonjakan tagihan"
          " diluapkan langsung ke rating toko aplikasi.",
          insight_style,
      )
  )
  story.append(Spacer(1, 8))

  # Grafik 2: Donut Sebaran Versi (Slices < 5% disembunyikan labelnya agar tidak overlap)
  fig2, ax2 = plt.subplots(figsize=(6.6, 2.7))
  ver_counts = df_source["Versi_Group"].value_counts()
  colors_ver = [
      "#0A3960",
      "#00A4E4",
      "#F39C12",
      "#2ECC71",
      "#95A5A6",
      "#BDC3C7",
  ]
  wedges2, _, autotexts2 = ax2.pie(
      ver_counts.values,
      autopct=lambda p: f"{p:.1f}%" if p >= 5 else "",
      startangle=140,
      colors=colors_ver[: len(ver_counts)],
      wedgeprops=dict(width=0.45, edgecolor="white", linewidth=2),
      pctdistance=0.75,
  )
  plt.setp(autotexts2, size=8.5, weight="bold", color="white")
  ax2.text(
      0,
      0,
      f"{len(ver_counts)}\nKlaster",
      ha="center",
      va="center",
      fontsize=8.5,
      weight="bold",
      color="#0A3960",
  )
  legend_labels2 = [
      f"{k}: {v:,} ({v/tot*100:.1f}%)"
      for k, v in zip(ver_counts.index, ver_counts.values)
  ]
  ax2.legend(
      wedges2,
      legend_labels2,
      title="Klaster Versi",
      loc="center left",
      bbox_to_anchor=(1, 0.5),
      fontsize=8,
      title_fontsize=8.5,
      frameon=False,
  )
  ax2.set_title(
      "Grafik 2: Sebaran Keluhan Menurut Klaster Versi Aplikasi",
      fontsize=10.5,
      weight="bold",
      color="#0A3960",
      pad=14,
  )
  story.append(make_clean_image(fig2, 520, 180))
  story.append(
      Paragraph(
          "<b>Analisis & Insight Grafik 2:</b> Sebagian besar keluhan berasal"
          " dari basis instalasi lawas (Versi 5.x & Tanpa Versi). Namun pada"
          " rilis modern 8.x, Versi 8.1.1 (842 ulasan) dan Versi 8.1.2 (471"
          " ulasan) mencatatkan kendala teknis yang signifikan. Kabar baiknya,"
          " rilis termutakhir Versi 8.1.3 (80 ulasan) memiliki volume keluhan"
          " terendah, mencerminkan efektivitas patch stabilisasi yang"
          " digulirkan.",
          insight_style,
      )
  )

  # ========================================================
  # HALAMAN 3: GRAFIK 3 & GRAFIK 4 (BAR DENGAN TEXTWRAP LENGKAP)
  # ========================================================
  story.append(PageBreak())
  story.append(
      Paragraph(
          "Bagian II: Rincian Subklasifikasi Kendala Aplikasi & Proses Bisnis",
          h2_style,
      )
  )

  # Grafik 3: Subklasifikasi Kendala Aplikasi
  fig3, ax3 = plt.subplots(figsize=(7.2, 3.0))
  sub_app = (
      df_source[df_source["Klasifikasi Utama"] == "Kendala Aplikasi"][
          "Subklasifikasi"
      ]
      .value_counts()
      .head(8)
      .sort_values(ascending=True)
  )
  wrapped_labels3 = [textwrap.fill(s, width=28) for s in sub_app.index]
  bars3 = ax3.barh(
      range(len(sub_app)),
      sub_app.values,
      color="#00A4E4",
      edgecolor="#0A3960",
      height=0.65,
  )
  ax3.set_yticks(range(len(sub_app)))
  ax3.set_yticklabels(wrapped_labels3, fontsize=7.5, color="#2C3E50")
  max_app = max(sub_app.values)
  ax3.set_xlim(0, max_app * 1.18)
  for bar in bars3:
    w = bar.get_width()
    ax3.text(
        w + max_app * 0.015,
        bar.get_y() + bar.get_height() / 2,
        f"{w:,} ({w/tot*100:.1f}%)",
        va="center",
        ha="left",
        fontsize=7.5,
        weight="bold",
        color="#0A3960",
    )
  ax3.set_title(
      "Grafik 3: Subklasifikasi Kendala Aplikasi Terbesar",
      fontsize=10.5,
      weight="bold",
      color="#0A3960",
      pad=10,
  )
  ax3.grid(axis="x", linestyle="--", alpha=0.3)
  ax3.spines["top"].set_visible(False)
  ax3.spines["right"].set_visible(False)
  story.append(make_clean_image(fig3, 520, 185))
  story.append(
      Paragraph(
          "<b>Analisis & Insight Grafik 3:</b> Dua pemicu utama di sisi teknis"
          " adalah kegagalan modul <i>Login, Registrasi, dan OTP</i> (721"
          " ulasan) serta <i>Performa dan konektivitas aplikasi</i> (698"
          " ulasan). Antrean pengiriman OTP SMS/WA dan latensi server saat"
          " beban puncak menjadi titik kritis yang mendesak untuk diperbaiki.",
          insight_style,
      )
  )
  story.append(Spacer(1, 8))

  # Grafik 4: Subklasifikasi Kendala Proses Bisnis
  fig4, ax4 = plt.subplots(figsize=(7.2, 3.0))
  sub_biz = (
      df_source[df_source["Klasifikasi Utama"] == "Kendala Proses Bisnis"][
          "Subklasifikasi"
      ]
      .value_counts()
      .head(8)
      .sort_values(ascending=True)
  )
  wrapped_labels4 = [textwrap.fill(s, width=28) for s in sub_biz.index]
  bars4 = ax4.barh(
      range(len(sub_biz)),
      sub_biz.values,
      color="#F39C12",
      edgecolor="#B9770E",
      height=0.65,
  )
  ax4.set_yticks(range(len(sub_biz)))
  ax4.set_yticklabels(wrapped_labels4, fontsize=7.5, color="#2C3E50")
  max_biz = max(sub_biz.values)
  ax4.set_xlim(0, max_biz * 1.18)
  for bar in bars4:
    w = bar.get_width()
    ax4.text(
        w + max_biz * 0.015,
        bar.get_y() + bar.get_height() / 2,
        f"{w:,} ({w/tot*100:.1f}%)",
        va="center",
        ha="left",
        fontsize=7.5,
        weight="bold",
        color="#B9770E",
    )
  ax4.set_title(
      "Grafik 4: Subklasifikasi Kendala Proses Bisnis Terbesar",
      fontsize=10.5,
      weight="bold",
      color="#0A3960",
      pad=10,
  )
  ax4.grid(axis="x", linestyle="--", alpha=0.3)
  ax4.spines["top"].set_visible(False)
  ax4.spines["right"].set_visible(False)
  story.append(make_clean_image(fig4, 520, 185))
  story.append(
      Paragraph(
          "<b>Analisis & Insight Grafik 4:</b> Keluhan operasional terbesar"
          " terpusat pada <i>Gangguan pasokan listrik / pemadaman</i> (967"
          " ulasan) dan <i>Tindak lanjut aduan lambat</i> (462 ulasan)."
          " Sementara itu, isu <i>Pengaduan dibatalkan / ditutup sepihak</i>"
          " (221 ulasan) menjadi alasan utama pelanggan memberikan bintang 1"
          " akibat status selesai sebelum perbaikan fisik selesai.",
          insight_style,
      )
  )

  # ========================================================
  # HALAMAN 4: GRAFIK 5 & GRAFIK 6 (TREN WAKTU BULANAN & TAHUNAN)
  # ========================================================
  story.append(PageBreak())
  story.append(
      Paragraph(
          "Bagian III: Dinamika Waktu (Tren Bulanan 2026 & Tren Tahunan)",
          h2_style,
      )
  )

  # Grafik 5: Tren Bulanan 2026 Saja (Legenda di atas, semua bulan tampil)
  fig5, ax5 = plt.subplots(figsize=(7.2, 2.9))
  df_2026 = df_source[df_source["Tahun"] == 2026]
  monthly_2026 = (
      df_2026.groupby(["Bulan_Tahun", "Klasifikasi Utama"])
      .size()
      .unstack(fill_value=0)
  )
  color_dict = {
      "Kendala Aplikasi": "#00A4E4",
      "Kendala Proses Bisnis": "#F39C12",
      "Belum dapat diklasifikasikan": "#95A5A6",
  }
  for col in monthly_2026.columns:
    ax5.plot(
        range(len(monthly_2026)),
        monthly_2026[col],
        marker="o",
        markersize=4.5,
        linewidth=2,
        label=col,
        color=color_dict.get(col, "#333333"),
    )
  bulan_short_pdf = {
      "01": "Jan", "02": "Feb", "03": "Mar", "04": "Apr",
      "05": "Mei", "06": "Jun", "07": "Jul", "08": "Agu",
      "09": "Sep", "10": "Okt", "11": "Nov", "12": "Des"
  }
  xticklabels_pdf = []
  for bt in monthly_2026.index:
    parts = str(bt).split("-")
    if len(parts) == 2 and parts[1] in bulan_short_pdf:
      xticklabels_pdf.append(f"{bulan_short_pdf[parts[1]]} {parts[0]}")
    else:
      xticklabels_pdf.append(str(bt))

  ax5.set_xticks(range(len(monthly_2026)))
  ax5.set_xticklabels(
      xticklabels_pdf, rotation=35, ha="right", fontsize=7.5, color="#2C3E50"
  )
  ax5.set_title(
      "Grafik 5: Tren Bulanan Ulasan Negatif Tahun 2026",
      fontsize=10.5,
      weight="bold",
      color="#0A3960",
      pad=24,
  )
  ax5.set_ylabel("Jumlah Ulasan", fontsize=8, color="#2C3E50")
  ax5.grid(True, linestyle="--", alpha=0.3)
  ax5.legend(
      loc="lower center",
      bbox_to_anchor=(0.5, 1.01),
      ncol=3,
      fontsize=7.5,
      frameon=False,
  )
  ax5.spines["top"].set_visible(False)
  ax5.spines["right"].set_visible(False)

  # Anotasi garis vertikal Update Sistem (September 2026) jika ada
  if "2026-09" in list(monthly_2026.index):
    sep_idx = list(monthly_2026.index).index("2026-09")
    ax5.axvline(x=sep_idx, color="#2C3E50", linestyle="--", linewidth=1.2, alpha=0.8)
    ax5.text(
        sep_idx,
        ax5.get_ylim()[1] * 0.92,
        " Update Sistem",
        fontsize=7,
        color="#0A3960",
        weight="bold",
    )

  story.append(make_clean_image(fig5, 520, 180))
  story.append(
      Paragraph(
          "<b>Analisis & Insight Grafik 5:</b> Volume keluhan stabil di kisaran"
          " 500-600 ulasan pada Q1 2026, lalu melonjak pada Mei (830) dan Juni"
          " (922). Puncak anomali terjadi pada Oktober 2026 (1.094 ulasan) di"
          " mana kendala aplikasi melesat tajam akibat peluncuran rilis v8.1.2"
          " yang mengalami kendala otentikasi login.",
          insight_style,
      )
  )
  story.append(Spacer(1, 8))

  # Grafik 6: Tren Tahunan (Stacked Bar Bersih - Khusus Tahun Ini & Tahun Lalu)
  fig6, ax6 = plt.subplots(figsize=(7.2, 2.9))
  max_year_pdf = (
      int(df_source["Tahun"].max())
      if "Tahun" in df_source.columns and df_source["Tahun"].notna().any()
      else 2026
  )
  target_years_pdf = [max_year_pdf - 1, max_year_pdf]
  df_source_yoy = df_source[df_source["Tahun"].isin(target_years_pdf)]
  yearly_df = (
      df_source_yoy.groupby(["Tahun", "Klasifikasi Utama"])
      .size()
      .unstack(fill_value=0)
  )
  years = [str(y) for y in yearly_df.index]
  bottom = np.zeros(len(years))
  for col in yearly_df.columns:
    vals = yearly_df[col].values
    ax6.bar(
        years,
        vals,
        bottom=bottom,
        label=col,
        color=color_dict.get(col, "#333333"),
        width=0.45,
    )
    bottom += vals
  for i, tot_yr in enumerate(bottom):
    ax6.text(
        i,
        tot_yr + (max(bottom) * 0.02 if len(bottom) > 0 and max(bottom) > 0 else 100),
        f"{int(tot_yr):,}",
        ha="center",
        va="bottom",
        fontsize=8,
        weight="bold",
        color="#0A3960",
    )
  ax6.set_title(
      f"Grafik 6: Tren Tahunan Ulasan Negatif (Year-over-Year {target_years_pdf[0]} - {target_years_pdf[1]})",
      fontsize=10.5,
      weight="bold",
      color="#0A3960",
      pad=24,
  )
  ax6.set_ylabel("Total Ulasan", fontsize=8, color="#2C3E50")
  ax6.set_ylim(0, (max(bottom) if len(bottom) > 0 else 1000) * 1.15)
  ax6.yaxis.set_major_formatter(ticker.StrMethodFormatter("{x:,.0f}"))
  ax6.grid(axis="y", linestyle="--", alpha=0.3)
  ax6.legend(
      loc="lower center",
      bbox_to_anchor=(0.5, 1.01),
      ncol=3,
      fontsize=7.5,
      frameon=False,
  )
  ax6.spines["top"].set_visible(False)
  ax6.spines["right"].set_visible(False)
  story.append(make_clean_image(fig6, 520, 180))
  story.append(
      Paragraph(
          f"<b>Analisis & Insight Grafik 6:</b> Perbandingan YoY menunjukkan total ulasan pada tahun {target_years_pdf[0]} tercatat {int(bottom[0]) if len(bottom) > 0 else 0:,} ulasan dan pada tahun {target_years_pdf[1]} tercatat {int(bottom[1]) if len(bottom) > 1 else (int(bottom[0]) if len(bottom) > 0 else 0):,} ulasan. Peningkatan ini sejalan dengan percepatan adopsi pengguna aktif PLN Mobile pasca program digitalisasi nasional.",
          insight_style,
      )
  )

  # ========================================================
  # HALAMAN 5: GRAFIK 7 & GRAFIK 8 (KATA KUNCI & SEVERITY)
  # ========================================================
  story.append(PageBreak())
  story.append(
      Paragraph(
          "Bagian IV: Analisis Kata Kunci & Tingkat Keparahan Rating", h2_style
      )
  )

  # Grafik 7: Top 10 Kata Pemicu (Murni kata kunci)
  fig7, ax7 = plt.subplots(figsize=(7.2, 3.0))
  top_kw = (
      df_source["Kata_Pemicu_Clean"]
      .dropna()
      .value_counts()
      .head(10)
      .sort_values(ascending=True)
  )
  bars7 = ax7.barh(
      range(len(top_kw)),
      top_kw.values,
      color="#C0392B",
      edgecolor="#922B21",
      height=0.65,
  )
  ax7.set_yticks(range(len(top_kw)))
  ax7.set_yticklabels(top_kw.index, fontsize=8, color="#2C3E50")
  max_kw = max(top_kw.values)
  ax7.set_xlim(0, max_kw * 1.16)
  for bar in bars7:
    w = bar.get_width()
    ax7.text(
        w + max_kw * 0.015,
        bar.get_y() + bar.get_height() / 2,
        f"{w:,}",
        va="center",
        ha="left",
        fontsize=7.5,
        weight="bold",
        color="#922B21",
    )
  ax7.set_title(
      "Grafik 7: Top 10 Kata / Frasa Pemicu Keluhan (Bersih Non-Keyword)",
      fontsize=10.5,
      weight="bold",
      color="#0A3960",
      pad=10,
  )
  ax7.grid(axis="x", linestyle="--", alpha=0.3)
  ax7.spines["top"].set_visible(False)
  ax7.spines["right"].set_visible(False)
  story.append(make_clean_image(fig7, 520, 185))
  story.append(
      Paragraph(
          "<b>Analisis & Insight Grafik 7:</b> Kata kunci pemicu komplain"
          " terbesar adalah <i>'login'</i> (264x), <i>'mati lampu'</i> (247x),"
          " <i>'pembayaran'</i> (172x), <i>'pemadaman'</i> (162x), dan"
          " <i>'lemot'</i> (160x). Kata-kata ini merefleksikan dua pilar fokus"
          " manajemen: keandalan sistem auth/payment di sisi aplikasi dan"
          " keandalan suplai listrik fisik.",
          insight_style,
      )
  )
  story.append(Spacer(1, 8))

  # Grafik 8: Distribusi Tingkat Keparahan Rating (Labels rapi 2 baris)
  fig8, ax8 = plt.subplots(figsize=(7.2, 3.0))
  rating_df = (
      df_source.groupby(["Klasifikasi Utama", "Rating"])
      .size()
      .unstack(fill_value=0)
  )
  wrapped_cats = [textwrap.fill(c, 16) for c in rating_df.index]
  b1 = (
      rating_df[1].values
      if 1 in rating_df
      else np.zeros(len(rating_df))
  )
  b2 = (
      rating_df[2].values
      if 2 in rating_df
      else np.zeros(len(rating_df))
  )
  b3 = (
      rating_df[3].values
      if 3 in rating_df
      else np.zeros(len(rating_df))
  )
  w_bar = 0.5
  ax8.bar(
      range(len(rating_df)),
      b1,
      width=w_bar,
      label="Rating 1 (Kritis)",
      color="#C0392B",
  )
  ax8.bar(
      range(len(rating_df)),
      b2,
      width=w_bar,
      bottom=b1,
      label="Rating 2 (Sedang)",
      color="#E67E22",
  )
  ax8.bar(
      range(len(rating_df)),
      b3,
      width=w_bar,
      bottom=b1 + b2,
      label="Rating 3 (Ringan)",
      color="#F1C40F",
  )
  totals = b1 + b2 + b3
  for i, tot_r in enumerate(totals):
    ax8.text(
        i,
        tot_r + 60,
        f"{tot_r:,}",
        ha="center",
        va="bottom",
        fontsize=7.5,
        weight="bold",
        color="#0A3960",
    )
  ax8.set_xticks(range(len(rating_df)))
  ax8.set_xticklabels(wrapped_cats, fontsize=8, color="#2C3E50")
  ax8.set_ylim(0, max(totals) * 1.15)
  ax8.yaxis.set_major_formatter(ticker.StrMethodFormatter("{x:,.0f}"))
  ax8.set_title(
      "Grafik 8: Distribusi Tingkat Keparahan Rating (1, 2, 3) per Kategori",
      fontsize=10.5,
      weight="bold",
      color="#0A3960",
      pad=24,
  )
  ax8.legend(
      loc="lower center",
      bbox_to_anchor=(0.5, 1.01),
      ncol=3,
      fontsize=7.5,
      frameon=False,
  )
  ax8.grid(axis="y", linestyle="--", alpha=0.3)
  ax8.spines["top"].set_visible(False)
  ax8.spines["right"].set_visible(False)
  story.append(make_clean_image(fig8, 520, 185))
  story.append(
      Paragraph(
          "<b>Analisis & Insight Grafik 8:</b> Lebih dari 85% ulasan negatif"
          " terpusat pada Rating 1 (Kritis). Sentimen negatif paling keras"
          " dialami pada kelompok Kendala Proses Bisnis (2.661 ulasan bintang"
          " 1), menegaskan tingginya sensitivitas pelanggan terhadap pemadaman"
          " listrik dan ketidaksesuaian nominal tagihan.",
          insight_style,
      )
  )

  # ========================================================
  # HALAMAN 6: GRAFIK 9, 10 & 11 (HEATMAP & EVALUASI CS)
  # ========================================================
  story.append(PageBreak())
  story.append(
      Paragraph(
          "Bagian V: Korelasi Versi Rilis & Evaluasi Kinerja Pelayanan CS",
          h2_style,
      )
  )

  # Grafik 9: Heatmap Korelasi Versi vs Kendala
  fig9, ax9 = plt.subplots(figsize=(7.2, 2.5))
  top_versions = ["08.01.01", "5.2.77", "5.2.79", "5.2.80", "5.2.81"]
  top_subcats = [
      "Gangguan pasokan listrik / pemadaman",
      "Login, Registrasi, dan OTP",
      "Performa dan konektivitas aplikasi",
      "Tagihan, pencatatan meter, dan tarif",
      "Transaksi digital gagal / status tidak sinkron",
  ]
  df_heat = df_source[
      df_source["reviewCreatedVersion"].isin(top_versions)
      & df_source["Subklasifikasi"].isin(top_subcats)
  ]
  heat_matrix = pd.crosstab(
      df_heat["Subklasifikasi"], df_heat["reviewCreatedVersion"]
  )
  im = ax9.imshow(heat_matrix.values, cmap="Blues", aspect="auto")
  ax9.set_xticks(range(len(heat_matrix.columns)))
  ax9.set_xticklabels(heat_matrix.columns, fontsize=7.5)
  ax9.set_yticks(range(len(heat_matrix.index)))
  wrapped_sub_heat = [textwrap.fill(s, 26) for s in heat_matrix.index]
  ax9.set_yticklabels(wrapped_sub_heat, fontsize=7)
  for i in range(len(heat_matrix.index)):
    for j in range(len(heat_matrix.columns)):
      val_h = heat_matrix.values[i, j]
      ax9.text(
          j,
          i,
          str(val_h),
          ha="center",
          va="center",
          color="white" if val_h > 120 else "black",
          fontsize=7.5,
          weight="bold",
      )
  ax9.set_title(
      "Grafik 9: Heatmap Korelasi Versi Aplikasi vs Subklasifikasi Kendala",
      fontsize=10.5,
      weight="bold",
      color="#0A3960",
      pad=10,
  )
  plt.colorbar(im, ax=ax9, fraction=0.03, pad=0.04)
  story.append(make_clean_image(fig9, 520, 160))
  story.append(
      Paragraph(
          "<b>Analisis & Insight Grafik 9:</b> Versi 5.2.81 mencatat konsentrasi"
          " keluhan terbesar pada gangguan pasokan listrik (246 kasus) dan"
          " login/OTP (146 kasus). Data korelasi ini menegaskan pentingnya"
          " pemantauan pasca rilis agar regu pengembang dapat segera"
          " meluncurkan patch perbaikan modul terkait.",
          insight_style,
      )
  )
  story.append(Spacer(1, 6))

  # Grafik 10 & 11: SLA Donut & Response Rate Bar (Side-by-side)
  fig10_11, (ax10, ax11) = plt.subplots(1, 2, figsize=(9.5, 2.5))
  sla_order = [
      "< 1 Jam",
      "1 - 6 Jam",
      "6 - 24 Jam",
      "> 24 Jam",
      "Belum Direspon",
  ]
  sla_data = (
      df_source["SLA_Bucket"]
      .value_counts()
      .reindex(sla_order)
      .fillna(0)
  )
  colors_sla = ["#27AE60", "#2980B9", "#F39C12", "#E74C3C", "#7F8C8D"]
  wedges10, _, autotexts10 = ax10.pie(
      sla_data.values,
      autopct=lambda p: f"{p:.1f}%" if p >= 5 else "",
      colors=colors_sla,
      wedgeprops=dict(width=0.45, edgecolor="white", linewidth=1.5),
      pctdistance=0.75,
  )
  plt.setp(autotexts10, size=7.5, weight="bold", color="white")
  ax10.legend(
      wedges10,
      [f"{k} ({int(v):,})" for k, v in zip(sla_data.index, sla_data.values)],
      loc="center left",
      bbox_to_anchor=(1, 0.5),
      fontsize=7,
      frameon=False,
  )
  ax10.set_title(
      "Grafik 10: Distribusi SLA Waktu Tanggap CS",
      fontsize=9.5,
      weight="bold",
      color="#0A3960",
      pad=10,
  )

  resp_rate = df_source.groupby("Klasifikasi Utama").apply(
      lambda g: (
          g["replyContent"].notna() & (g["replyContent"].str.strip() != "")
      ).mean()
      * 100
  )
  wrapped_cat_resp = [textwrap.fill(c, 14) for c in resp_rate.index]
  bars11 = ax11.bar(
      range(len(resp_rate)),
      resp_rate.values,
      color=["#95A5A6", "#00A4E4", "#F39C12"],
      edgecolor="#0A3960",
      width=0.55,
  )
  ax11.set_xticks(range(len(resp_rate)))
  ax11.set_xticklabels(wrapped_cat_resp, fontsize=7)
  ax11.set_ylim(0, 115)
  ax11.set_title(
      "Grafik 11: Tingkat Respon CS per Kategori (%)",
      fontsize=9.5,
      weight="bold",
      color="#0A3960",
      pad=10,
  )
  for bar in bars11:
    yval = bar.get_height()
    ax11.text(
        bar.get_x() + bar.get_width() / 2.0,
        yval + 2,
        f"{yval:.1f}%",
        ha="center",
        va="bottom",
        fontsize=7.5,
        weight="bold",
    )
  story.append(make_clean_image(fig10_11, 520, 155))
  story.append(
      Paragraph(
          "<b>Analisis & Insight Grafik 10 & 11:</b> Kinerja agen penanganan"
          " keluhan sangat baik pada ranah kendala teknis dan proses bisnis"
          " (mendekati 100% terjawab, dan 65%+ selesai dalam rentang waktu < 6"
          " jam). Namun pada ulasan tanpa detail spesifik, respon agen baru"
          " mencapai 48,3%, sehingga diperlukan auto-reply edukatif untuk"
          " mempertahankan rating toko aplikasi.",
          insight_style,
      )
  )

  # ========================================================
  # HALAMAN 7: KESIMPULAN & ROADMAP REKOMENDASI MANAJEMEN
  # ========================================================
  story.append(PageBreak())
  story.append(
      Paragraph(
          "Bagian VI: Kesimpulan & Roadmap Rekomendasi Strategis", h2_style
      )
  )
  story.append(
      Paragraph(
          "Berdasarkan analisis menyeluruh terhadap 5 metrik KPI dan 11"
          " visualisasi data di atas, berikut roadmap rekomendasi strategis bagi"
          " Divisi Manajemen Digital (DIV MDG):<br/><br/><b>1. Penguatan"
          " Arsitektur Otentikasi & Redundansi OTP (Jangka"
          " Pendek):</b><br/>Kendala nomor satu aplikasi adalah kegagalan login"
          " dan pengiriman OTP (721 ulasan). Diperlukan implementasi multi-SMS"
          " vendor gateway dan fallback otomatis ke WhatsApp OTP untuk menjamin"
          " verifikasi akun tetap andal di jam sibuk.<br/><br/><b>2. Sistem"
          " Validasi Dua Arah Status Pengaduan (Jangka Menengah):</b><br/>Untuk"
          " meniadakan keluhan 'pengaduan ditutup sepihak' (221 ulasan), sistem"
          " aplikasi wajib menerapkan <i>customer closure validation</i>,"
          " yaitu tiket aduan hanya dapat ditutup jika pelanggan telah"
          " mengonfirmasi penyelesaian lewat aplikasi.<br/><br/><b>3."
          " Personalisasi Respon CS & ETA Penanganan Gangguan (Jangka"
          " Panjang):</b><br/>Meningkatkan kepuasan pelanggan dengan mengurangi"
          " template jawaban statis dan menyajikan estimasi waktu pemulihan"
          " gangguan (ETA) serta nomor tiket tindak lanjut lapangan secara"
          " transparan.",
          body_style,
      )
  )

  doc.build(story)
  pdf_buffer.seek(0)
  return pdf_buffer.getvalue()


# ==========================================
# 4. SIDEBAR & KONTROL FILTER
# ==========================================
st.sidebar.image(
    "https://upload.wikimedia.org/wikipedia/commons/9/97/Logo_PLN.png",
    width=110,
)
st.sidebar.title("Kontrol Filter")

def _find_default_excel():
    fname = "Klasifikasi_Ulasan_Negatif_PLN_Mobile.xlsx"
    if os.path.exists(fname):
        return fname
    base_dir = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.path.dirname(os.path.abspath(__file__))
    cand = os.path.join(base_dir, fname)
    if os.path.exists(cand):
        return cand
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        bundle_cand = os.path.join(sys._MEIPASS, fname)
        if os.path.exists(bundle_cand):
            return bundle_cand
    return None

default_file = _find_default_excel()
uploaded_file = st.sidebar.file_uploader(
    "Unggah File Excel Lain (Opsional)", type=["xlsx", "xls"]
)

active_source = (
    uploaded_file
    if uploaded_file is not None
    else default_file
)
if active_source is None:
  st.info("👋 Silakan unggah file Excel ulasan untuk menampilkan dashboard.")
  st.stop()

df_raw = load_data(active_source)

min_date = (
    df_raw["at"].min().date()
    if df_raw["at"].notna().any()
    else pd.to_datetime("2026-01-01").date()
)
max_date = (
    df_raw["at"].max().date()
    if df_raw["at"].notna().any()
    else pd.to_datetime("2026-11-05").date()
)

date_range = st.sidebar.date_input(
    "Rentang Tanggal",
    value=(min_date, max_date),
    min_value=min_date,
    max_value=max_date,
)
available_ratings = sorted(df_raw["Rating"].dropna().unique().astype(int))
selected_ratings = st.sidebar.multiselect(
    "Rating Bintang", options=available_ratings, default=available_ratings
)
available_main = sorted(df_raw["Klasifikasi Utama"].dropna().unique().tolist())
selected_main = st.sidebar.multiselect(
    "Klasifikasi Utama", options=available_main, default=available_main
)
available_versions = sorted(df_raw["Versi_Group"].dropna().unique().tolist())
selected_versions = st.sidebar.multiselect(
    "Versi Aplikasi", options=available_versions, default=available_versions
)

df_filtered = df_raw.copy()
if len(date_range) == 2:
  start_dt, end_dt = pd.to_datetime(date_range[0]), pd.to_datetime(
      date_range[1]
  ) + pd.Timedelta(days=1)
  df_filtered = df_filtered[
      (df_filtered["at"] >= start_dt) & (df_filtered["at"] < end_dt)
  ]
if selected_ratings:
  df_filtered = df_filtered[df_filtered["Rating"].isin(selected_ratings)]
if selected_main:
  df_filtered = df_filtered[
      df_filtered["Klasifikasi Utama"].isin(selected_main)
  ]
if selected_versions:
  df_filtered = df_filtered[df_filtered["Versi_Group"].isin(selected_versions)]

# ==========================================
# 5. HEADER DASHBOARD
# ==========================================
st.markdown(
    """
<div class="main-header">
    <h1>Dashboard Analisis Klasifikasi Ulasan Negatif PLN Mobile</h1>
    <p>Monitoring Kualitas Layanan Digital, Stabilitas Versi Aplikasi, dan Kinerja SLA | Divisi Manajemen Digital (DIV MDG)</p>
</div>
""",
    unsafe_allow_html=True,
)

if df_filtered.empty:
  st.warning("Tidak ada data yang memenuhi kriteria filter.")
  st.stop()

# Perhitungan Metrik & Pertumbuhan MoM Aktual
total_reviews = len(df_filtered)
app_issues = (df_filtered["Klasifikasi Utama"] == "Kendala Aplikasi").sum()
biz_issues = (df_filtered["Klasifikasi Utama"] == "Kendala Proses Bisnis").sum()
has_reply = (
    df_filtered["replyContent"].notna()
    & (df_filtered["replyContent"].str.strip() != "")
).sum()
reply_rate = (has_reply / total_reviews * 100) if total_reviews > 0 else 0
median_sla = df_filtered["Response_Hours"].dropna().median()

# Hitung perbandingan MoM riil
monthly_counts_all = (
    df_filtered.groupby("Bulan_Tahun").size().sort_index()
)
if len(monthly_counts_all) >= 2:
  mom_growth = (
      (monthly_counts_all.iloc[-1] - monthly_counts_all.iloc[-2])
      / monthly_counts_all.iloc[-2]
      * 100
  )
  mom_text = f"{mom_growth:+.1f}% MoM"
else:
  mom_text = "N/A"

k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("Total Ulasan Negatif", f"{total_reviews:,}", mom_text)
k2.metric(
    "Kendala Aplikasi",
    f"{app_issues:,}",
    f"{app_issues/total_reviews*100:.1f}% dari total",
    delta_color="off",
)
k3.metric(
    "Kendala Bisnis",
    f"{biz_issues:,}",
    f"{biz_issues/total_reviews*100:.1f}% dari total",
    delta_color="off",
)
k4.metric("Tingkat Respon CS", f"{reply_rate:.1f}%", f"{has_reply:,} terjawab")
k5.metric(
    "Median SLA Respon",
    f"{median_sla:.1f} Jam" if not np.isnan(median_sla) else "N/A",
    "Target: < 2 Jam",
)

st.write("")
color_palette = {
    "Kendala Aplikasi": "#00A4E4",
    "Kendala Proses Bisnis": "#F39C12",
    "Belum dapat diklasifikasikan": "#95A5A6",
}

# ==========================================
# 6. TABS DASHBOARD MANAJEMEN
# ==========================================
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📊 Ringkasan Eksekutif (Format Laporan)",
    "📈 Tren Waktu (Bulanan 2026 & Tahunan)",
    "🔍 Korelasi Versi & Kata Kunci",
    "⏱️ Kinerja Waktu Tanggap CS",
    "📥 Unduh Laporan PDF & Data",
])

with tab1:
  col_t1_l, col_t1_r = st.columns(2)
  with col_t1_l:
    st.subheader("Grafik 1: Persentase Klasifikasi Utama")
    df_pie_m = df_filtered["Klasifikasi Utama"].value_counts().reset_index()
    df_pie_m.columns = ["Kategori", "Jumlah"]
    fig1 = go.Figure(
        go.Pie(
            labels=df_pie_m["Kategori"],
            values=df_pie_m["Jumlah"],
            hole=0.6,
            marker=dict(
                colors=[
                    color_palette.get(k, "#0A3960") for k in df_pie_m["Kategori"]
                ]
            ),
        )
    )
    fig1.update_layout(
        annotations=[
            dict(
                text=f"<b>{total_reviews:,}</b><br>Ulasan",
                x=0.5,
                y=0.5,
                font_size=16,
                showarrow=False,
            )
        ],
        legend=dict(
            orientation="v", yanchor="middle", y=0.5, xanchor="left", x=1.02
        ),
        margin=dict(t=20, b=20, l=10, r=10),
        height=380,
    )
    st.plotly_chart(fig1, use_container_width=True)

  with col_t1_r:
    st.subheader("Grafik 2: Sebaran Versi Aplikasi")
    df_ver = df_filtered["Versi_Group"].value_counts().reset_index()
    df_ver.columns = ["Versi", "Jumlah"]
    fig2 = go.Figure(
        go.Pie(
            labels=df_ver["Versi"],
            values=df_ver["Jumlah"],
            hole=0.6,
            marker=dict(
                colors=[
                    "#00A4E4",
                    "#0A3960",
                    "#F39C12",
                    "#2ECC71",
                    "#95A5A6",
                    "#BDC3C7",
                ]
            ),
        )
    )
    fig2.update_layout(
        annotations=[
            dict(
                text=f"<b>{df_ver['Versi'].nunique()}</b><br>Klaster",
                x=0.5,
                y=0.5,
                font_size=16,
                showarrow=False,
            )
        ],
        legend=dict(
            orientation="v", yanchor="middle", y=0.5, xanchor="left", x=1.02
        ),
        margin=dict(t=20, b=20, l=10, r=10),
        height=380,
    )
    st.plotly_chart(fig2, use_container_width=True)

  st.markdown("---")
  cb1, cb2 = st.columns(2)
  with cb1:
    st.subheader(f"Grafik 3: Kendala Aplikasi ({app_issues:,})")
    df_sa = (
        df_filtered[df_filtered["Klasifikasi Utama"] == "Kendala Aplikasi"][
            "Subklasifikasi"
        ]
        .value_counts()
        .reset_index()
    )
    df_sa.columns = ["Subklasifikasi", "Jumlah"]
    fig3 = go.Figure(
        go.Bar(
            x=df_sa["Jumlah"],
            y=df_sa["Subklasifikasi"],
            orientation="h",
            text=[f"{j:,}" for j in df_sa["Jumlah"]],
            textposition="outside",
            marker=dict(color=df_sa["Jumlah"], colorscale="Blues"),
        )
    )
    fig3.update_layout(
        height=420,
        margin=dict(t=10, b=30, l=10, r=40),
        yaxis=dict(autorange="reversed"),
    )
    st.plotly_chart(fig3, use_container_width=True)

  with cb2:
    st.subheader(f"Grafik 4: Kendala Proses Bisnis ({biz_issues:,})")
    df_sb = (
        df_filtered[
            df_filtered["Klasifikasi Utama"] == "Kendala Proses Bisnis"
        ]["Subklasifikasi"]
        .value_counts()
        .reset_index()
    )
    df_sb.columns = ["Subklasifikasi", "Jumlah"]
    fig4 = go.Figure(
        go.Bar(
            x=df_sb["Jumlah"],
            y=df_sb["Subklasifikasi"],
            orientation="h",
            text=[f"{j:,}" for j in df_sb["Jumlah"]],
            textposition="outside",
            marker=dict(color=df_sb["Jumlah"], colorscale="YlOrRd"),
        )
    )
    fig4.update_layout(
        height=420,
        margin=dict(t=10, b=30, l=10, r=40),
        yaxis=dict(autorange="reversed"),
    )
    st.plotly_chart(fig4, use_container_width=True)

with tab2:
  st.subheader("Grafik 5: Tren Bulanan Ulasan Negatif Tahun 2026")
  df_2026 = df_filtered[df_filtered["Tahun"] == 2026].copy()
  if not df_2026.empty:
    bulan_map_id = {
        1: "Januari",
        2: "Februari",
        3: "Maret",
        4: "April",
        5: "Mei",
        6: "Juni",
        7: "Juli",
        8: "Agustus",
        9: "September",
        10: "Oktober",
        11: "November",
        12: "Desember",
    }
    bulan_short_id = {
        1: "Jan",
        2: "Feb",
        3: "Mar",
        4: "Apr",
        5: "Mei",
        6: "Jun",
        7: "Jul",
        8: "Agu",
        9: "Sep",
        10: "Okt",
        11: "Nov",
        12: "Des",
    }
    df_2026["Bulan_Num"] = df_2026["at"].dt.month
    m_2026 = (
        df_2026.groupby(["Bulan_Num", "Klasifikasi Utama"])
        .size()
        .reset_index(name="Jumlah")
    )
    m_2026 = m_2026.sort_values(by="Bulan_Num")
    m_2026["Periode"] = m_2026["Bulan_Num"].apply(
        lambda m: f"{bulan_map_id.get(m, '')} 2026"
    )
    m_2026["Periode_Short"] = m_2026["Bulan_Num"].apply(
        lambda m: f"{bulan_short_id.get(m, '')} 2026"
    )

    kat_order = [
        "Kendala Proses Bisnis",
        "Belum dapat diklasifikasikan",
        "Kendala Aplikasi",
    ]
    kat_order = [
        k for k in kat_order if k in m_2026["Klasifikasi Utama"].unique()
    ] or sorted(m_2026["Klasifikasi Utama"].unique())

    fig5 = px.line(
        m_2026,
        x="Periode",
        y="Jumlah",
        color="Klasifikasi Utama",
        markers=True,
        color_discrete_map=color_palette,
        category_orders={"Klasifikasi Utama": kat_order},
        labels={"Periode": "Periode", "Jumlah": "Jumlah Ulasan"},
    )

    # Hovertemplate ringkas & rapi:
    # September 2026
    # • Kendala Proses Bisnis: 337 ulasan
    # • Belum dapat diklasifikasikan: 322 ulasan
    # • Kendala Aplikasi: 305 ulasan
    for trace in fig5.data:
      trace.hovertemplate = f"• {trace.name}: %{{y:,}} ulasan<extra></extra>"

    unique_periods = (
        m_2026[["Bulan_Num", "Periode", "Periode_Short"]]
        .drop_duplicates()
        .sort_values("Bulan_Num")
    )

    fig5.update_layout(
        height=380,
        hovermode="x unified",
        legend=dict(
            orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1
        ),
        xaxis=dict(
            title="Periode",
            tickmode="array",
            tickvals=unique_periods["Periode"].tolist(),
            ticktext=unique_periods["Periode_Short"].tolist(),
            showgrid=True,
            gridcolor="rgba(0,0,0,0.06)",
        ),
        yaxis=dict(
            title="Jumlah Ulasan",
            showgrid=True,
            gridcolor="rgba(0,0,0,0.06)",
        ),
        margin=dict(t=30, b=40, l=40, r=20),
    )

    # Anotasi garis vertikal pada bulan September 2026 (Update Sistem)
    sep_periode = f"{bulan_map_id[9]} 2026"
    if sep_periode in unique_periods["Periode"].values:
      fig5.add_shape(
          type="line",
          x0=sep_periode,
          x1=sep_periode,
          y0=0,
          y1=1,
          yref="paper",
          line=dict(color="#2C3E50", width=1.5, dash="dash"),
      )
      fig5.add_annotation(
          x=sep_periode,
          y=0.98,
          yref="paper",
          text="<b>Update Sistem</b>",
          showarrow=False,
          xanchor="right",
          yanchor="top",
          font=dict(size=10.5, color="#0A3960"),
          bgcolor="rgba(255, 255, 255, 0.9)",
          bordercolor="#00A4E4",
          borderwidth=1,
          borderpad=3,
      )

    st.plotly_chart(
        fig5, use_container_width=True, config={"displayModeBar": False}
    )

  st.markdown("---")
  st.subheader("Grafik 6: Tren Tahunan Ulasan Negatif (Historical YoY)")
  # Hanya data tahun ini dan tahun lalu saja yang ditampilkan
  current_year = (
      int(df_filtered["Tahun"].max())
      if not df_filtered.empty and df_filtered["Tahun"].notna().any()
      else 2026
  )
  target_years = [current_year - 1, current_year]
  df_yoy = df_filtered[df_filtered["Tahun"].isin(target_years)]

  if not df_yoy.empty:
    y_df = (
        df_yoy.groupby(["Tahun", "Klasifikasi Utama"])
        .size()
        .reset_index(name="Jumlah")
    )
    y_df["Tahun"] = y_df["Tahun"].astype(str)
    fig6 = px.bar(
        y_df,
        x="Tahun",
        y="Jumlah",
        color="Klasifikasi Utama",
        barmode="stack",
        color_discrete_map=color_palette,
        text_auto=True,
        labels={"Tahun": "Tahun", "Jumlah": "Jumlah Ulasan"},
    )
    fig6.update_layout(
        height=380,
        legend=dict(
            orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1
        ),
        xaxis=dict(title="Tahun"),
        yaxis=dict(title="Jumlah Ulasan"),
    )
    st.plotly_chart(
        fig6, use_container_width=True, config={"displayModeBar": False}
    )
  else:
    st.info("Tidak ada data ulasan untuk perbandingan tahun ini dan tahun lalu pada filter saat ini.")

with tab3:
  ct3_l, ct3_r = st.columns(2)
  with ct3_l:
    st.subheader("Grafik 7: Top 10 Kata Pemicu (Bersih Non-Keyword)")
    df_kw = (
        df_filtered["Kata_Pemicu_Clean"]
        .dropna()
        .value_counts()
        .head(10)
        .reset_index()
    )
    df_kw.columns = ["Kata Kunci", "Frekuensi"]
    fig7 = px.bar(
        df_kw.sort_values(by="Frekuensi", ascending=True),
        x="Frekuensi",
        y="Kata Kunci",
        orientation="h",
        color="Frekuensi",
        color_continuous_scale="Reds",
        text="Frekuensi",
    )
    fig7.update_layout(height=380, coloraxis_showscale=False)
    st.plotly_chart(fig7, use_container_width=True)

  with ct3_r:
    st.subheader("Grafik 8: Tingkat Keparahan Rating")
    df_rd = (
        df_filtered.groupby(["Klasifikasi Utama", "Rating"])
        .size()
        .reset_index(name="Jumlah")
    )
    fig8 = px.bar(
        df_rd,
        x="Klasifikasi Utama",
        y="Jumlah",
        color="Rating",
        barmode="stack",
        color_discrete_sequence=["#C0392B", "#E67E22", "#F1C40F"],
    )
    fig8.update_layout(
        height=380,
        legend=dict(
            orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1
        ),
    )
    st.plotly_chart(fig8, use_container_width=True)

  st.subheader("Grafik 9: Heatmap Korelasi Versi vs Subklasifikasi")
  top_subs = (
      df_filtered[
          df_filtered["Klasifikasi Utama"].isin(
              ["Kendala Aplikasi", "Kendala Proses Bisnis"]
          )
      ]["Subklasifikasi"]
      .value_counts()
      .head(7)
      .index.tolist()
  )
  df_hm = df_filtered[df_filtered["Subklasifikasi"].isin(top_subs)]
  fig9 = px.imshow(
      pd.crosstab(df_hm["Subklasifikasi"], df_hm["Versi_Group"]),
      text_auto=True,
      aspect="auto",
      color_continuous_scale="Blues",
  )
  fig9.update_layout(height=360)
  st.plotly_chart(fig9, use_container_width=True)

with tab4:
  ct4_l, ct4_r = st.columns(2)
  with ct4_l:
    st.subheader("Grafik 10: Distribusi SLA Waktu Tanggap CS")
    sla_cnt = (
        df_filtered["SLA_Bucket"]
        .value_counts()
        .reindex(
            ["< 1 Jam", "1 - 6 Jam", "6 - 24 Jam", "> 24 Jam", "Belum Direspon"]
        )
        .fillna(0)
        .reset_index()
    )
    sla_cnt.columns = ["SLA", "Jumlah"]
    fig10 = px.pie(
        sla_cnt,
        names="SLA",
        values="Jumlah",
        hole=0.55,
        color="SLA",
        color_discrete_map={
            "< 1 Jam": "#27AE60",
            "1 - 6 Jam": "#2980B9",
            "6 - 24 Jam": "#F39C12",
            "> 24 Jam": "#E74C3C",
            "Belum Direspon": "#7F8C8D",
        },
    )
    fig10.update_layout(
        height=360,
        legend=dict(
            orientation="v", yanchor="middle", y=0.5, xanchor="left", x=1.02
        ),
    )
    st.plotly_chart(fig10, use_container_width=True)

  with ct4_r:
    st.subheader("Grafik 11: Tingkat Respon CS per Kategori (%)")
    cat_resp = (
        df_filtered.groupby("Klasifikasi Utama")
        .apply(
            lambda g: (
                g["replyContent"].notna()
                & (g["replyContent"].str.strip() != "")
            ).mean()
            * 100
        )
        .reset_index(name="Respon_Rate")
    )
    fig11 = px.bar(
        cat_resp,
        x="Klasifikasi Utama",
        y="Respon_Rate",
        color="Klasifikasi Utama",
        color_discrete_map=color_palette,
        text=cat_resp["Respon_Rate"].apply(lambda v: f"{v:.1f}%"),
    )
    fig11.update_layout(height=360, showlegend=False, yaxis_range=[0, 115])
    st.plotly_chart(fig11, use_container_width=True)

with tab5:
  st.subheader("Unduh Dokumen Laporan Eksekutif Resmi (PDF)")
  st.write(
      "Dapatkan dokumen laporan PDF eksekutif lengkap yang mencakup **seluruh 5"
      " Metrik KPI**, **11 Grafik Analisis**, tabel kata kunci, dan rekomendasi"
      " strategis bagi manajemen."
  )

  with st.spinner("Mengompilasi dokumen PDF..."):
    pdf_output = generate_full_pdf_report(df_filtered)

  col_d1, col_d2 = st.columns([2, 3])
  with col_d1:
    st.download_button(
        label="📥 Unduh Laporan Eksekutif Lengkap (PDF)",
        data=pdf_output,
        file_name="Laporan_Eksekutif_Evaluasi_PLN_Mobile_2026.pdf",
        mime="application/pdf",
        use_container_width=True,
    )

  st.markdown("---")
  st.subheader("Eksplorasi Data Mentah Ulasan")
  search_kw = st.text_input("🔍 Cari kata kunci ulasan atau balasan CS:", "")
  df_v = df_filtered.copy()
  if search_kw.strip():
    kw_str = search_kw.strip().lower()
    df_v = df_v[
        df_v["Ulasan"].astype(str).str.lower().str.contains(kw_str)
        | df_v["replyContent"].astype(str).str.lower().str.contains(kw_str)
    ]
  st.dataframe(
      df_v[[
          "at",
          "Rating",
          "Klasifikasi Utama",
          "Subklasifikasi",
          "Versi_Group",
          "Ulasan",
          "Dasar Klasifikasi",
          "SLA_Bucket",
          "replyContent",
      ]].sort_values(by="at", ascending=False),
      use_container_width=True,
      height=380,
  )