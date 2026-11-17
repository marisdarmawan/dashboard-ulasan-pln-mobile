import io
import sys
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
# 3. GENERATOR LAPORAN RESMI (PDF DINAMIS)
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

  def make_clean_image(fig, w=525, h=215):
    buf = io.BytesIO()
    plt.savefig(
        buf, format="png", dpi=200, bbox_inches="tight", facecolor="white"
    )
    plt.close(fig)
    buf.seek(0)
    return RLImage(buf, width=w, height=h)

  bulan_full_id = {
      "01": "Januari",
      "02": "Februari",
      "03": "Maret",
      "04": "April",
      "05": "Mei",
      "06": "Juni",
      "07": "Juli",
      "08": "Agustus",
      "09": "September",
      "10": "Oktober",
      "11": "November",
      "12": "Desember",
  }
  bulan_short_pdf = {
      "01": "Jan",
      "02": "Feb",
      "03": "Mar",
      "04": "Apr",
      "05": "Mei",
      "06": "Jun",
      "07": "Jul",
      "08": "Agu",
      "09": "Sep",
      "10": "Okt",
      "11": "Nov",
      "12": "Des",
  }

  def format_periode_id(bt, short=False):
    parts = str(bt).split("-")
    mapping = bulan_short_pdf if short else bulan_full_id
    if len(parts) == 2 and parts[1] in mapping:
      return f"{mapping[parts[1]]} {parts[0]}"
    return str(bt)

  story = []
  tot = len(df_source)

  valid_years = (
      sorted(df_source["Tahun"].dropna().astype(int).unique())
      if "Tahun" in df_source.columns
      else []
  )
  if len(valid_years) > 1:
    periode_header = f"Tahun {valid_years[0]} - {valid_years[-1]}"
  elif len(valid_years) == 1:
    periode_header = f"Tahun {valid_years[0]}"
  else:
    periode_header = "Periode Aktif"

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
          f" Rekomendasi Strategis | {periode_header}",
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

  app_cnt = int((df_source["Klasifikasi Utama"] == "Kendala Aplikasi").sum())
  biz_cnt = int(
      (df_source["Klasifikasi Utama"] == "Kendala Proses Bisnis").sum()
  )
  app_pct = (app_cnt / tot * 100) if tot > 0 else 0.0
  biz_pct = (biz_cnt / tot * 100) if tot > 0 else 0.0

  has_reply = int(
      (
          df_source["replyContent"].notna()
          & (df_source["replyContent"].astype(str).str.strip() != "")
      ).sum()
  )
  reply_rate = (has_reply / tot * 100) if tot > 0 else 0.0
  median_sla_raw = (
      df_source["Response_Hours"].dropna().median()
      if "Response_Hours" in df_source.columns
      else np.nan
  )
  median_sla = 0.0 if pd.isna(median_sla_raw) else float(median_sla_raw)

  kpi_table_data = [[
      Paragraph(
          f"<b>Total Ulasan Negatif</b><br/><font size=14 color='#0A3960'><b>{tot:,}</b></font><br/><font size=7 color='#7F8C8D'>Bintang 1 s.d. 3</font>",
          body_style,
      ),
      Paragraph(
          f"<b>Kendala Aplikasi</b><br/><font size=14 color='#C0392B'><b>{app_cnt:,}</b></font><br/><font size=7 color='#C0392B'>Porsi: {app_pct:.1f}%</font>",
          body_style,
      ),
      Paragraph(
          f"<b>Kendala Bisnis</b><br/><font size=14 color='#C0392B'><b>{biz_cnt:,}</b></font><br/><font size=7 color='#C0392B'>Porsi: {biz_pct:.1f}%</font>",
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
          " keluhan pengguna bintang 1 s.d. 3 pada dataset/filter aktif.<br/>•"
          f" <b>Badge Merah pada Kendala Aplikasi ({app_pct:.1f}%) dan Kendala"
          f" Bisnis ({biz_pct:.1f}%)</b>: Merupakan <b>persentase porsi"
          " kontribusi dari total ulasan</b> (bukan nilai lonjakan"
          " bulanan/MoM), diberi highlight merah sebagai beban keluhan"
          " operasional yang membutuhkan penanganan.<br/>• <b>Badge Hijau pada"
          f" Tingkat Respon CS ({reply_rate:.1f}%) dan SLA ({median_sla:.1f}"
          " Jam)</b>: Menunjukkan capaian responsivitas tim pelayanan digital"
          " PLN 123 dalam menanggapi keluhan pelanggan.",
          body_style,
      )
  )

  # ========================================================
  # HALAMAN 2: GRAFIK 1 & GRAFIK 2
  # ========================================================
  story.append(PageBreak())
  story.append(
      Paragraph(
          "Bagian I: Proporsi Klasifikasi Utama & Sebaran Versi Aplikasi",
          h2_style,
      )
  )

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

  if biz_cnt > app_cnt:
    g1_lead = (
        f"Keluhan Proses Bisnis ({biz_pct:.1f}%) lebih dominan dibanding"
        f" kendala murni teknis aplikasi ({app_pct:.1f}%)."
    )
  elif app_cnt > biz_cnt:
    g1_lead = (
        f"Kendala teknis aplikasi ({app_pct:.1f}%) lebih dominan dibanding"
        f" keluhan proses bisnis ({biz_pct:.1f}%)."
    )
  else:
    g1_lead = (
        f"Keluhan Proses Bisnis ({biz_pct:.1f}%) dan Kendala Aplikasi"
        f" ({app_pct:.1f}%) memiliki proporsi yang berimbang."
    )

  story.append(
      Paragraph(
          f"<b>Analisis & Insight Grafik 1:</b> {g1_lead} Hal ini membuktikan"
          " bahwa pengguna memperlakukan PLN Mobile sebagai gerbang layanan"
          " kelistrikan secara menyeluruh, sehingga kendala lapangan maupun"
          " performa aplikasi berdampak langsung terhadap penilaian di toko"
          " aplikasi.",
          insight_style,
      )
  )
  story.append(Spacer(1, 8))

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

  if len(ver_counts) >= 2:
    g2_text = (
        f"<b>Analisis & Insight Grafik 2:</b> Konsentrasi keluhan terbesar"
        f" berasal dari <b>{ver_counts.index[0]}</b> ({int(ver_counts.iloc[0]):,}"
        f" ulasan / {ver_counts.iloc[0]/tot*100:.1f}%), diikuti oleh"
        f" <b>{ver_counts.index[1]}</b> ({int(ver_counts.iloc[1]):,} ulasan)."
        f" Sementara itu, klaster <b>{ver_counts.index[-1]}</b> mencatatkan"
        f" volume terendah ({int(ver_counts.iloc[-1]):,} ulasan) pada periode"
        " ini."
    )
  else:
    g2_text = (
        "<b>Analisis & Insight Grafik 2:</b> Seluruh ulasan pada filter aktif"
        f" berasal dari klaster <b>{ver_counts.index[0]}</b>"
        f" ({int(ver_counts.iloc[0]):,} ulasan)."
    )
  story.append(Paragraph(g2_text, insight_style))

  # ========================================================
  # HALAMAN 3: GRAFIK 3 & GRAFIK 4
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
  sub_app_desc = (
      df_source[df_source["Klasifikasi Utama"] == "Kendala Aplikasi"][
          "Subklasifikasi"
      ]
      .dropna()
      .value_counts()
  )
  sub_app = sub_app_desc.head(8).sort_values(ascending=True)
  if not sub_app.empty:
    wrapped_labels3 = [textwrap.fill(str(s), width=28) for s in sub_app.index]
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
          f"{int(w):,} ({w/tot*100:.1f}%)",
          va="center",
          ha="left",
          fontsize=7.5,
          weight="bold",
          color="#0A3960",
      )
  else:
    ax3.text(
        0.5,
        0.5,
        "Tidak ada data Kendala Aplikasi pada filter aktif",
        ha="center",
        va="center",
        fontsize=9,
        color="#7F8C8D",
    )
    ax3.set_xticks([])
    ax3.set_yticks([])

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

  if len(sub_app_desc) >= 2:
    g3_text = (
        "<b>Analisis & Insight Grafik 3:</b> Dua pemicu utama di sisi teknis"
        f" adalah modul <i>{sub_app_desc.index[0]}</i>"
        f" ({int(sub_app_desc.iloc[0]):,} ulasan) serta"
        f" <i>{sub_app_desc.index[1]}</i> ({int(sub_app_desc.iloc[1]):,}"
        " ulasan). Kedua komponen ini menjadi titik kritis utama yang mendesak"
        " untuk dioptimalkan."
    )
  elif len(sub_app_desc) == 1:
    g3_text = (
        "<b>Analisis & Insight Grafik 3:</b> Pemicu utama kendala teknis"
        f" terpusat pada <i>{sub_app_desc.index[0]}</i>"
        f" ({int(sub_app_desc.iloc[0]):,} ulasan)."
    )
  else:
    g3_text = (
        "<b>Analisis & Insight Grafik 3:</b> Tidak terdapat data subklasifikasi"
        " Kendala Aplikasi pada filter yang dipilih."
    )
  story.append(Paragraph(g3_text, insight_style))
  story.append(Spacer(1, 8))

  # Grafik 4: Subklasifikasi Kendala Proses Bisnis
  fig4, ax4 = plt.subplots(figsize=(7.2, 3.0))
  sub_biz_desc = (
      df_source[df_source["Klasifikasi Utama"] == "Kendala Proses Bisnis"][
          "Subklasifikasi"
      ]
      .dropna()
      .value_counts()
  )
  sub_biz = sub_biz_desc.head(8).sort_values(ascending=True)
  if not sub_biz.empty:
    wrapped_labels4 = [textwrap.fill(str(s), width=28) for s in sub_biz.index]
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
          f"{int(w):,} ({w/tot*100:.1f}%)",
          va="center",
          ha="left",
          fontsize=7.5,
          weight="bold",
          color="#B9770E",
      )
  else:
    ax4.text(
        0.5,
        0.5,
        "Tidak ada data Kendala Proses Bisnis pada filter aktif",
        ha="center",
        va="center",
        fontsize=9,
        color="#7F8C8D",
    )
    ax4.set_xticks([])
    ax4.set_yticks([])

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

  if len(sub_biz_desc) >= 3:
    g4_text = (
        "<b>Analisis & Insight Grafik 4:</b> Keluhan operasional terbesar"
        f" terpusat pada <i>{sub_biz_desc.index[0]}</i>"
        f" ({int(sub_biz_desc.iloc[0]):,} ulasan) dan"
        f" <i>{sub_biz_desc.index[1]}</i> ({int(sub_biz_desc.iloc[1]):,}"
        f" ulasan), diikuti oleh <i>{sub_biz_desc.index[2]}</i>"
        f" ({int(sub_biz_desc.iloc[2]):,} ulasan) yang sangat memengaruhi"
        " persepsi kepuasan pelanggan."
    )
  elif len(sub_biz_desc) >= 1:
    g4_text = (
        "<b>Analisis & Insight Grafik 4:</b> Keluhan operasional terbesar"
        f" terpusat pada <i>{sub_biz_desc.index[0]}</i>"
        f" ({int(sub_biz_desc.iloc[0]):,} ulasan)."
    )
  else:
    g4_text = (
        "<b>Analisis & Insight Grafik 4:</b> Tidak terdapat data subklasifikasi"
        " Kendala Proses Bisnis pada filter yang dipilih."
    )
  story.append(Paragraph(g4_text, insight_style))

  # ========================================================
  # HALAMAN 4: GRAFIK 5 & GRAFIK 6
  # ========================================================
  story.append(PageBreak())
  max_year_pdf = (
      int(df_source["Tahun"].max())
      if "Tahun" in df_source.columns and df_source["Tahun"].notna().any()
      else 2026
  )
  story.append(
      Paragraph(
          f"Bagian III: Dinamika Waktu (Tren Bulanan {max_year_pdf} & Tren"
          " Tahunan)",
          h2_style,
      )
  )

  fig5, ax5 = plt.subplots(figsize=(7.2, 2.9))
  df_latest_yr = df_source[df_source["Tahun"] == max_year_pdf]
  monthly_yr = (
      df_latest_yr.groupby(["Bulan_Tahun", "Klasifikasi Utama"])
      .size()
      .unstack(fill_value=0)
  )
  color_dict = {
      "Kendala Aplikasi": "#00A4E4",
      "Kendala Proses Bisnis": "#F39C12",
      "Belum dapat diklasifikasikan": "#95A5A6",
  }
  for col in monthly_yr.columns:
    ax5.plot(
        range(len(monthly_yr)),
        monthly_yr[col],
        marker="o",
        markersize=4.5,
        linewidth=2,
        label=col,
        color=color_dict.get(col, "#333333"),
    )

  xticklabels_pdf = [format_periode_id(bt, short=True) for bt in monthly_yr.index]
  ax5.set_xticks(range(len(monthly_yr)))
  ax5.set_xticklabels(
      xticklabels_pdf, rotation=35, ha="right", fontsize=7.5, color="#2C3E50"
  )
  ax5.set_title(
      f"Grafik 5: Tren Bulanan Ulasan Negatif Tahun {max_year_pdf}",
      fontsize=10.5,
      weight="bold",
      color="#0A3960",
      pad=24,
  )
  ax5.set_ylabel("Jumlah Ulasan", fontsize=8, color="#2C3E50")
  ax5.grid(True, linestyle="--", alpha=0.3)
  if len(monthly_yr.columns) > 0:
    ax5.legend(
        loc="lower center",
        bbox_to_anchor=(0.5, 1.01),
        ncol=3,
        fontsize=7.5,
        frameon=False,
    )
  ax5.spines["top"].set_visible(False)
  ax5.spines["right"].set_visible(False)

  if "2026-09" in list(monthly_yr.index):
    sep_idx = list(monthly_yr.index).index("2026-09")
    ax5.axvline(
        x=sep_idx, color="#2C3E50", linestyle="--", linewidth=1.2, alpha=0.8
    )
    ax5.text(
        sep_idx,
        ax5.get_ylim()[1] * 0.92,
        " Update Sistem",
        fontsize=7,
        color="#0A3960",
        weight="bold",
    )

  story.append(make_clean_image(fig5, 520, 180))

  monthly_tot = df_latest_yr.groupby("Bulan_Tahun").size()
  if not monthly_tot.empty:
    peak_bt = monthly_tot.idxmax()
    peak_val = int(monthly_tot.max())
    min_bt = monthly_tot.idxmin()
    min_val = int(monthly_tot.min())
    peak_cat_s = df_latest_yr[df_latest_yr["Bulan_Tahun"] == peak_bt][
        "Klasifikasi Utama"
    ].value_counts()
    peak_dom_cat = peak_cat_s.index[0] if not peak_cat_s.empty else "-"
    peak_dom_cnt = int(peak_cat_s.iloc[0]) if not peak_cat_s.empty else 0
    g5_text = (
        f"<b>Analisis & Insight Grafik 5:</b> Sepanjang tahun {max_year_pdf},"
        f" volume keluhan bulanan bergerak dari titik terendah {min_val:,}"
        f" ulasan ({format_periode_id(min_bt)}) hingga titik puncak"
        f" <b>{peak_val:,} ulasan pada {format_periode_id(peak_bt)}</b>."
        f" Lonjakan puncak tersebut didominasi oleh kategori <b>{peak_dom_cat}</b>"
        f" ({peak_dom_cnt:,} ulasan)."
    )
  else:
    g5_text = (
        "<b>Analisis & Insight Grafik 5:</b> Tidak terdapat data bulanan pada"
        " filter yang dipilih."
    )
  story.append(Paragraph(g5_text, insight_style))
  story.append(Spacer(1, 8))

  # Grafik 6: Tren Tahunan (YoY)
  fig6, ax6 = plt.subplots(figsize=(7.2, 2.9))
  target_years_pdf = [max_year_pdf - 1, max_year_pdf]
  df_source_yoy = df_source[df_source["Tahun"].isin(target_years_pdf)]
  yearly_df = (
      df_source_yoy.groupby(["Tahun", "Klasifikasi Utama"])
      .size()
      .unstack(fill_value=0)
  )
  years = [str(int(y)) for y in yearly_df.index]
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
        tot_yr
        + (max(bottom) * 0.02 if len(bottom) > 0 and max(bottom) > 0 else 10),
        f"{int(tot_yr):,}",
        ha="center",
        va="bottom",
        fontsize=8,
        weight="bold",
        color="#0A3960",
    )
  ax6.set_title(
      "Grafik 6: Tren Tahunan Ulasan Negatif (Year-over-Year"
      f" {target_years_pdf[0]} - {target_years_pdf[1]})",
      fontsize=10.5,
      weight="bold",
      color="#0A3960",
      pad=24,
  )
  ax6.set_ylabel("Total Ulasan", fontsize=8, color="#2C3E50")
  ax6.set_ylim(0, (max(bottom) if len(bottom) > 0 else 100) * 1.15)
  ax6.yaxis.set_major_formatter(ticker.StrMethodFormatter("{x:,.0f}"))
  ax6.grid(axis="y", linestyle="--", alpha=0.3)
  if len(yearly_df.columns) > 0:
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

  if len(bottom) >= 2 and bottom[0] > 0:
    yoy_pct = (bottom[1] - bottom[0]) / bottom[0] * 100
    arah = (
        "peningkatan"
        if yoy_pct > 0
        else ("penurunan" if yoy_pct < 0 else "stabilitas")
    )
    g6_text = (
        "<b>Analisis & Insight Grafik 6:</b> Perbandingan YoY menunjukkan total"
        f" ulasan pada tahun {years[0]} tercatat {int(bottom[0]):,} ulasan dan"
        f" pada tahun {years[1]} tercatat {int(bottom[1]):,} ulasan (mengalami"
        f" {arah} sebesar {abs(yoy_pct):.1f}%)."
    )
  elif len(bottom) == 1:
    g6_text = (
        "<b>Analisis & Insight Grafik 6:</b> Data historis pada filter saat ini"
        f" mencakup tahun {years[0]} dengan total {int(bottom[0]):,} ulasan."
    )
  else:
    g6_text = (
        "<b>Analisis & Insight Grafik 6:</b> Tidak terdapat data perbandingan"
        " tahunan pada filter aktif."
    )
  story.append(Paragraph(g6_text, insight_style))

  # ========================================================
  # HALAMAN 5: GRAFIK 7 & GRAFIK 8
  # ========================================================
  story.append(PageBreak())
  story.append(
      Paragraph(
          "Bagian IV: Analisis Kata Kunci & Tingkat Keparahan Rating", h2_style
      )
  )

  fig7, ax7 = plt.subplots(figsize=(7.2, 3.0))
  top_kw_desc = df_source["Kata_Pemicu_Clean"].dropna().value_counts()
  top_kw = top_kw_desc.head(10).sort_values(ascending=True)
  if not top_kw.empty:
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
          f"{int(w):,}",
          va="center",
          ha="left",
          fontsize=7.5,
          weight="bold",
          color="#922B21",
      )
  else:
    ax7.text(
        0.5,
        0.5,
        "Tidak ada kata kunci terdeteksi",
        ha="center",
        va="center",
        fontsize=9,
        color="#7F8C8D",
    )
    ax7.set_xticks([])
    ax7.set_yticks([])

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

  top5_kw = top_kw_desc.head(5)
  if not top5_kw.empty:
    kw_list_str = ", ".join(
        [f"<i>'{k}'</i> ({int(v):,}x)" for k, v in top5_kw.items()]
    )
    g7_text = (
        "<b>Analisis & Insight Grafik 7:</b> Kata kunci pemicu komplain"
        f" terbesar adalah {kw_list_str}. Frasa dominan ini merefleksikan"
        " fokus utama keluhan pelanggan pada layanan aplikasi maupun suplai"
        " listrik."
    )
  else:
    g7_text = (
        "<b>Analisis & Insight Grafik 7:</b> Tidak ditemukan kata kunci pemicu"
        " spesifik pada data yang difilter."
    )
  story.append(Paragraph(g7_text, insight_style))
  story.append(Spacer(1, 8))

  # Grafik 8: Distribusi Tingkat Keparahan Rating
  fig8, ax8 = plt.subplots(figsize=(7.2, 3.0))
  rating_df = (
      df_source.groupby(["Klasifikasi Utama", "Rating"])
      .size()
      .unstack(fill_value=0)
  )
  wrapped_cats = [textwrap.fill(str(c), 16) for c in rating_df.index]
  b1 = rating_df[1].values if 1 in rating_df else np.zeros(len(rating_df))
  b2 = rating_df[2].values if 2 in rating_df else np.zeros(len(rating_df))
  b3 = rating_df[3].values if 3 in rating_df else np.zeros(len(rating_df))
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
  max_tot_r = max(totals) if len(totals) > 0 and max(totals) > 0 else 100
  for i, tot_r in enumerate(totals):
    ax8.text(
        i,
        tot_r + max_tot_r * 0.02,
        f"{int(tot_r):,}",
        ha="center",
        va="bottom",
        fontsize=7.5,
        weight="bold",
        color="#0A3960",
    )
  ax8.set_xticks(range(len(rating_df)))
  ax8.set_xticklabels(wrapped_cats, fontsize=8, color="#2C3E50")
  ax8.set_ylim(0, max_tot_r * 1.15)
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

  r1_cnt = int((df_source["Rating"] == 1).sum())
  r1_pct = (r1_cnt / tot * 100) if tot > 0 else 0.0
  r1_by_cat = df_source[df_source["Rating"] == 1][
      "Klasifikasi Utama"
  ].value_counts()
  if not r1_by_cat.empty:
    g8_text = (
        f"<b>Analisis & Insight Grafik 8:</b> Sebanyak {r1_pct:.1f}% ({r1_cnt:,}"
        " ulasan) dari total ulasan negatif terpusat pada Rating 1 (Kritis)."
        " Sentimen negatif paling keras dialami pada kelompok"
        f" <b>{r1_by_cat.index[0]}</b> ({int(r1_by_cat.iloc[0]):,} ulasan"
        " bintang 1), menegaskan tingginya sensitivitas pelanggan terhadap"
        " kendala pada kategori tersebut."
    )
  else:
    g8_text = (
        f"<b>Analisis & Insight Grafik 8:</b> Proporsi ulasan Rating 1 (Kritis)"
        f" pada filter aktif tercatat sebesar {r1_pct:.1f}% ({r1_cnt:,} ulasan)."
    )
  story.append(Paragraph(g8_text, insight_style))

  # ========================================================
  # HALAMAN 6: GRAFIK 9, 10 & 11
  # ========================================================
  story.append(PageBreak())
  story.append(
      Paragraph(
          "Bagian V: Korelasi Versi Rilis & Evaluasi Kinerja Pelayanan CS",
          h2_style,
      )
  )

  # Grafik 9: Heatmap Korelasi Versi vs Kendala (100% Dinamis Top 5)
  fig9, ax9 = plt.subplots(figsize=(7.2, 2.5))
  top_versions = (
      df_source["reviewCreatedVersion"]
      .dropna()
      .astype(str)
      .loc[lambda s: ~s.str.strip().isin(["", "nan", "-", "None"])]
      .value_counts()
      .head(5)
      .index.tolist()
      if "reviewCreatedVersion" in df_source.columns
      else []
  )
  top_subcats = (
      df_source[
          df_source["Klasifikasi Utama"].isin(
              ["Kendala Aplikasi", "Kendala Proses Bisnis"]
          )
      ]["Subklasifikasi"]
      .dropna()
      .value_counts()
      .head(5)
      .index.tolist()
  )
  df_heat = df_source[
      df_source["reviewCreatedVersion"].astype(str).isin(top_versions)
      & df_source["Subklasifikasi"].isin(top_subcats)
  ]
  heat_matrix = pd.crosstab(
      df_heat["Subklasifikasi"], df_heat["reviewCreatedVersion"].astype(str)
  )

  if not heat_matrix.empty:
    im = ax9.imshow(heat_matrix.values, cmap="Blues", aspect="auto")
    ax9.set_xticks(range(len(heat_matrix.columns)))
    ax9.set_xticklabels(heat_matrix.columns, fontsize=7.5)
    ax9.set_yticks(range(len(heat_matrix.index)))
    wrapped_sub_heat = [textwrap.fill(str(s), 26) for s in heat_matrix.index]
    ax9.set_yticklabels(wrapped_sub_heat, fontsize=7)
    thresh_h = heat_matrix.values.max() * 0.55
    for i in range(len(heat_matrix.index)):
      for j in range(len(heat_matrix.columns)):
        val_h = heat_matrix.values[i, j]
        ax9.text(
            j,
            i,
            str(val_h),
            ha="center",
            va="center",
            color="white" if val_h > thresh_h else "black",
            fontsize=7.5,
            weight="bold",
        )
    plt.colorbar(im, ax=ax9, fraction=0.03, pad=0.04)

    max_pos = np.unravel_index(
        np.argmax(heat_matrix.values), heat_matrix.shape
    )
    max_sub_h = heat_matrix.index[max_pos[0]]
    max_ver_h = heat_matrix.columns[max_pos[1]]
    max_val_h = int(heat_matrix.values[max_pos])
    g9_text = (
        f"<b>Analisis & Insight Grafik 9:</b> Konsentrasi korelasi tertinggi"
        f" tercatat pada <b>Versi {max_ver_h}</b> untuk kendala"
        f" <i>{max_sub_h}</i> ({max_val_h:,} kasus). Pemetaan silang ini"
        " membantu tim pengembang memprioritaskan perbaikan modul pasca rilis."
    )
  else:
    ax9.text(
        0.5,
        0.5,
        "Data versi spesifik tidak tersedia pada filter aktif",
        ha="center",
        va="center",
        fontsize=9,
        color="#7F8C8D",
    )
    ax9.set_xticks([])
    ax9.set_yticks([])
    g9_text = (
        "<b>Analisis & Insight Grafik 9:</b> Informasi versi rilis spesifik"
        " tidak mencukupi untuk membentuk matriks korelasi pada filter saat ini."
    )

  ax9.set_title(
      "Grafik 9: Heatmap Korelasi Versi Aplikasi vs Subklasifikasi Kendala",
      fontsize=10.5,
      weight="bold",
      color="#0A3960",
      pad=10,
  )
  story.append(make_clean_image(fig9, 520, 160))
  story.append(Paragraph(g9_text, insight_style))
  story.append(Spacer(1, 6))

  # Grafik 10 & 11: SLA Donut & Response Rate Bar
  fig10_11, (ax10, ax11) = plt.subplots(1, 2, figsize=(9.5, 2.5))
  sla_order = [
      "< 1 Jam",
      "1 - 6 Jam",
      "6 - 24 Jam",
      "> 24 Jam",
      "Belum Direspon",
  ]
  sla_data = (
      df_source["SLA_Bucket"].value_counts().reindex(sla_order).fillna(0)
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
          g["replyContent"].notna()
          & (g["replyContent"].astype(str).str.strip() != "")
      ).mean()
      * 100
  )
  wrapped_cat_resp = [textwrap.fill(str(c), 14) for c in resp_rate.index]
  bar_colors_11 = [color_dict.get(c, "#00A4E4") for c in resp_rate.index]
  bars11 = ax11.bar(
      range(len(resp_rate)),
      resp_rate.values,
      color=bar_colors_11,
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

  fast_sla_pct = (
      (sla_data.get("< 1 Jam", 0) + sla_data.get("1 - 6 Jam", 0)) / tot * 100
      if tot > 0
      else 0.0
  )
  if not resp_rate.empty:
    low_resp_cat = resp_rate.idxmin()
    low_resp_val = float(resp_rate.min())
    g10_11_text = (
        f"<b>Analisis & Insight Grafik 10 & 11:</b> Sebanyak {fast_sla_pct:.1f}%"
        " ulasan berhasil ditanggapi dalam rentang waktu di bawah 6 jam."
        " Sementara itu, kategori dengan tingkat respon terendah saat ini"
        f" adalah <b>{low_resp_cat}</b> ({low_resp_val:.1f}%), sehingga"
        " optimalisasi auto-reply edukatif diperlukan untuk menjaga konsistensi"
        " pelayanan di toko aplikasi."
    )
  else:
    g10_11_text = (
        "<b>Analisis & Insight Grafik 10 & 11:</b> Distribusi waktu tanggap CS"
        " ditampilkan sesuai filter aktif."
    )
  story.append(Paragraph(g10_11_text, insight_style))

  # ========================================================
  # HALAMAN 7: KESIMPULAN & ROADMAP REKOMENDASI MANAJEMEN
  # ========================================================
  story.append(PageBreak())
  story.append(
      Paragraph(
          "Bagian VI: Kesimpulan & Roadmap Rekomendasi Strategis", h2_style
      )
  )

  top_app_str = (
      f"<i>{sub_app_desc.index[0]}</i> ({int(sub_app_desc.iloc[0]):,} ulasan)"
      if not sub_app_desc.empty
      else "modul otentikasi dan performa aplikasi"
  )
  top_biz_str = (
      f"<i>{sub_biz_desc.index[0]}</i> ({int(sub_biz_desc.iloc[0]):,} ulasan)"
      if not sub_biz_desc.empty
      else "penanganan gangguan dan tindak lanjut aduan"
  )

  story.append(
      Paragraph(
          "Berdasarkan analisis menyeluruh terhadap 5 metrik KPI dan 11"
          " visualisasi data di atas, berikut roadmap rekomendasi strategis bagi"
          " Divisi Manajemen Digital (DIV MDG):<br/><br/><b>1. Penguatan"
          " Keandalan Modul Aplikasi Utama (Jangka Pendek):</b><br/>Prioritas"
          f" perbaikan teknis difokuskan pada {top_app_str} sebagai penyumbang"
          " keluhan aplikasi tertinggi, melalui peningkatan kapasitas layanan"
          " pada jam beban puncak serta redundansi kanal verifikasi.<br/><br/>"
          "<b>2. Validasi Transparansi Status Pengaduan (Jangka"
          " Menengah):</b><br/>Untuk menekan keluhan pada kelompok"
          f" {top_biz_str} dan pencegahan penutupan tiket sepihak, sistem perlu"
          " menerapkan <i>customer closure validation</i> agar tiket aduan"
          " berstatus selesai setelah mendapat konfirmasi"
          " pelanggan.<br/><br/><b>3. Personalisasi Respon CS & Informasi ETA"
          " Gangguan (Jangka Panjang):</b><br/>Meningkatkan kepuasan pelanggan"
          " dengan mengurangi template jawaban statis serta menyajikan estimasi"
          " waktu pemulihan (ETA) dan progres tindak lanjut lapangan secara"
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
    "Unggah File Excel Review PLN Mobile", type=["xlsx", "xls"]
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
    <h1>Dashboard Klasifikasi Ulasan Negatif PLN Mobile</h1>
    <p>Monitoring Kualitas Layanan Digital, Stabilitas Versi Aplikasi, dan Kinerja SLA tim Contact Center 123. Dibuat oleh Tim Beyond Kwh Divisi Manajemen Digital (DIV MDG)</p>
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
    "📊 Klasifikasi Ulasan Negatif",
    "📈 Tren Waktu (Bulanan & Tahunan)",
    "🔍 Versi Aplikasi & Kata Kunci (Topik)",
    "⏱️ Kinerja Waktu Tanggap CS",
    "📥 Unduh Laporan PDF & Data",
])

with tab1:
  col_t1_l, col_t1_r = st.columns(2)
  with col_t1_l:
    st.subheader("Grafik 1: Klasifikasi Ulasan Negatif")
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
    st.subheader("Grafik 7: Top 10 Topik Ulasan Negatif")
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

  st.subheader("Grafik 9: Heatmap Versi Aplikasi dan Klasifikasi")
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
