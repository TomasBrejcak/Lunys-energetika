import io
import os
import re
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import pdfplumber
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# --- NASTAVENIE STRÁNKY A RESPONSÍVNEHO CSS ---
st.set_page_config(
    page_title="Energetický Dashboard LUNYS 2026",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
        padding-left: 2rem;
        padding-right: 2rem;
    }
    @media (max-width: 768px) {
        .block-container {
            padding-left: 1rem;
            padding-right: 1rem;
        }
    }
    </style>
""",
    unsafe_allow_html=True,
)

# --- BEZPEČNÉ VYTVORENIE / OPRAVA PRIEČINKA PRE FAKTÚRY ---
INVOICE_DIR = "faktury_zse"
if os.path.exists(INVOICE_DIR) and not os.path.isdir(INVOICE_DIR):
  try:
    os.remove(INVOICE_DIR)
  except:
    pass
if not os.path.exists(INVOICE_DIR):
  os.makedirs(INVOICE_DIR)

# --- LOGO V SIDEBARE ---
if os.path.exists("logo.svg"):
  st.sidebar.image("logo.svg", width=220)
elif os.path.exists("logo.png"):
  st.sidebar.image("logo.png", width=220)
elif os.path.exists("logo.jpg"):
  st.sidebar.image("logo.jpg", width=220)
else:
  st.sidebar.markdown("### LUNYS, s.r.o.")

st.sidebar.divider()

st.title("⚡ Energetický Dashboard & Analýza FVE 2026")
st.markdown(
    "Interaktívny nástroj pre vyhodnotenie reálnych spotrieb, nákladov (PDF"
    " faktúry + spravodlivo prepočítané podružné merania) a úspor z FVE."
)

# --- MAPOVANIE EIC KÓDOV NA ODBERNÉ MIESTA LUNYS ---
EIC_MAP = {
    "24ZVS00007394912": {"nazov": "Trafostanica PEBEK", "lokalita": "Poprad"},
    "24ZZS4000182286S": {
        "nazov": "Bratislava – Trafo Feld",
        "lokalita": "Bratislava",
    },
    "24ZZS40001822865": {
        "nazov": "Bratislava – Trafo Feld",
        "lokalita": "Bratislava",
    },
    "24ZVS0000016878H": {"nazov": "Beňovský Apartments", "lokalita": "Poprad"},
    "24ZVS0000059036Q": {
        "nazov": "Hlavný elektromer / Nájomcovia",
        "lokalita": "Poprad",
    },
    "24ZVS00000590360": {
        "nazov": "Hlavný elektromer / Nájomcovia",
        "lokalita": "Poprad",
    },
    "24ZVS0000787041D": {"nazov": "Trafostanica HARAMIA", "lokalita": "Poprad"},
}


# --- FUNKCIA NA GENEROVANIE PREZENTAČNÉHO EXCEL REPORTU ---
def build_presentation_excel(
    df_detail,
    df_kpi_clean,
    df_invoices,
    df_fve,
    selected_months_list,
    selected_oms_list,
    report_title="ENERGETICKÝ REPORT 2026",
):
  output = io.BytesIO()
  wb = openpyxl.Workbook()
  wb.remove(wb.active)

  navy_header = PatternFill("solid", fgColor="17365D")
  sub_header = PatternFill("solid", fgColor="1F4E78")
  zebra_fill = PatternFill("solid", fgColor="F5F8FC")
  total_fill = PatternFill("solid", fgColor="D9E1F2")
  alert_fill = PatternFill("solid", fgColor="FCE4D6")

  font_title = Font(name="Calibri", size=15, bold=True, color="FFFFFF")
  font_meta = Font(name="Calibri", size=10, italic=True, color="444444")
  font_note = Font(name="Calibri", size=9, color="595959")
  font_tbl_hdr = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
  font_data = Font(name="Calibri", size=10, color="000000")
  font_total = Font(name="Calibri", size=10, bold=True, color="17365D")
  font_alert = Font(name="Calibri", size=10, bold=True, color="C00000")

  thin_border = Border(
      left=Side(style="thin", color="D9D9D9"),
      right=Side(style="thin", color="D9D9D9"),
      top=Side(style="thin", color="D9D9D9"),
      bottom=Side(style="thin", color="D9D9D9"),
  )
  total_border = Border(
      top=Side(style="thin", color="17365D"),
      bottom=Side(style="double", color="17365D"),
  )

  def populate_formatted_sheet(
      sheet_name,
      df_table,
      heading,
      sub_note,
      add_sum_row=True,
      highlight_penalties=False,
  ):
    ws = wb.create_sheet(title=sheet_name)
    ws.views.sheetView[0].showGridLines = True

    max_cols = max(len(df_table.columns), 5) if not df_table.empty else 5

    ws["A1"] = f"LUNYS, s.r.o. | {heading}"
    ws["A1"].font = font_title
    ws["A1"].fill = navy_header
    ws["A1"].alignment = Alignment(vertical="center", indent=1)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_cols)
    ws.row_dimensions[1].height = 30

    obdobie_txt = (
        ", ".join(selected_months_list)
        if selected_months_list
        else "Bez vybraného obdobia"
    )
    ws["A2"] = (
        f"Obdobie: {obdobie_txt}  |  Počet vybraných OM:"
        f" {len(selected_oms_list)}"
    )
    ws["A2"].font = font_meta
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=max_cols)

    ws["A3"] = sub_note
    ws["A3"].font = font_note
    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=max_cols)

    if df_table.empty:
      ws["A5"] = "Pre zvolený filter nie sú k dispozícii žiadne dáta."
      ws["A5"].font = font_meta
      return

    hdr_row = 5
    cols = list(df_table.columns)

    for c_idx, col_name in enumerate(cols, start=1):
      c = ws.cell(row=hdr_row, column=c_idx, value=col_name)
      c.font = font_tbl_hdr
      c.fill = sub_header
      c.alignment = Alignment(
          horizontal="center", vertical="center", wrap_text=True
      )
    ws.row_dimensions[hdr_row].height = 28

    for r_offset, row_vals in enumerate(df_table.values):
      r_num = hdr_row + 1 + r_offset
      row_bg = zebra_fill if r_offset % 2 == 1 else None
      for c_idx, val in enumerate(row_vals, start=1):
        if pd.isna(val):
          val = ""
        cell = ws.cell(row=r_num, column=c_idx, value=val)
        cell.font = font_data
        cell.border = thin_border
        if row_bg:
          cell.fill = row_bg

        col_title = cols[c_idx - 1]
        if any(k in col_title for k in ["(kWh)", "Vyroba"]):
          cell.number_format = "#,##0"
          cell.alignment = Alignment(horizontal="right")
        elif any(k in col_title for k in ["(€)", "Uspora", "POKUTY"]):
          cell.number_format = "#,##0.00 €"
          cell.alignment = Alignment(horizontal="right")
          if (
              highlight_penalties
              and any(
                  k in col_title
                  for k in ["Prekročenie", "Jalová", "účinník", "POKUTY"]
              )
              and isinstance(val, (int, float))
              and val > 0
          ):
            cell.fill = alert_fill
            cell.font = font_alert
        elif "Cena" in col_title:
          cell.number_format = "0.0000 €"
          cell.alignment = Alignment(horizontal="right")
        elif any(k in col_title for k in ["Podiel", "%"]):
          cell.number_format = "0.0%"
          cell.alignment = Alignment(horizontal="right")

    if add_sum_row and len(df_table) > 0:
      sum_row = hdr_row + len(df_table) + 1
      ws.cell(row=sum_row, column=1, value="SPOLU").font = font_total
      for c_idx, col_title in enumerate(cols, start=1):
        cell = ws.cell(row=sum_row, column=c_idx)
        cell.fill = total_fill
        cell.border = total_border
        cell.font = font_total
        col_let = get_column_letter(c_idx)

        if any(
            k in col_title
            for k in ["(kWh)", "(€)", "Vyroba", "Uspora", "POKUTY"]
        ):
          cell.value = f"=SUM({col_let}{hdr_row+1}:{col_let}{sum_row-1})"
          cell.number_format = (
              "#,##0"
              if any(k in col_title for k in ["(kWh)", "Vyroba"])
              else "#,##0.00 €"
          )
          cell.alignment = Alignment(horizontal="right")
        elif "Cena" in col_title:
          kwh_cols = [i + 1 for i, c in enumerate(cols) if "(kWh)" in c]
          eur_cols = [
              i + 1
              for i, c in enumerate(cols)
              if "Náklady" in c or "Spolu bez DPH" in c
          ]
          if kwh_cols and eur_cols:
            k_let = get_column_letter(kwh_cols[0])
            e_let = get_column_letter(eur_cols[0])
            cell.value = f"=IFERROR({e_let}{sum_row}/{k_let}{sum_row}, 0)"
          else:
            cell.value = f"=AVERAGE({col_let}{hdr_row+1}:{col_let}{sum_row-1})"
          cell.number_format = "0.0000 €"
          cell.alignment = Alignment(horizontal="right")
        elif any(k in col_title for k in ["Podiel", "%"]):
          cell.value = f"=AVERAGE({col_let}{hdr_row+1}:{col_let}{sum_row-1})"
          cell.number_format = "0.0%"
          cell.alignment = Alignment(horizontal="right")

    ws.freeze_panes = f"A{hdr_row + 1}"
    ws.auto_filter.ref = (
        f"A{hdr_row}:{get_column_letter(len(cols))}{hdr_row + len(df_table)}"
    )

    for c_idx in range(1, len(cols) + 1):
      col_let = get_column_letter(c_idx)
      max_len = len(str(cols[c_idx - 1]))
      for r_idx in range(hdr_row + 1, hdr_row + len(df_table) + 1):
        val_s = str(ws.cell(row=r_idx, column=c_idx).value or "")
        max_len = max(max_len, len(val_s))
      ws.column_dimensions[col_let].width = min(max(max_len + 4, 14), 36)

  # 1. Manažérsky súhrn
  if not df_kpi_clean.empty:
    df_sum_clean = (
        df_kpi_clean.groupby(["Lokalita", "Odberné miesto"], as_index=False)
        .agg(**{
            "Spotreba spolu (kWh)": ("Spotreba (kWh)", "sum"),
            "Náklady spolu (€)": ("Náklady (€)", "sum"),
        })
        .sort_values(
            by=["Lokalita", "Náklady spolu (€)"], ascending=[True, False]
        )
    )
    df_sum_clean["Priemerná cena (€/kWh)"] = df_sum_clean.apply(
        lambda r: (r["Náklady spolu (€)"] / r["Spotreba spolu (kWh)"])
        if r["Spotreba spolu (kWh)"] > 0
        else 0.0,
        axis=1,
    )
  else:
    df_sum_clean = pd.DataFrame()

  populate_formatted_sheet(
      sheet_name="1. Manažérsky súhrn",
      df_table=df_sum_clean,
      heading=report_title,
      sub_note=(
          "Konsolidovaný sumár s vylúčením podružných meraní"
          " (Skladová hala I., Správa 99, Bananáreň, Zemiakareň) pre zamedzenie"
          " duplicity."
      ),
      add_sum_row=True,
  )

  # 2. Audit faktúr a pokút
  if not df_invoices.empty:
    inv_clean = df_invoices.drop_duplicates(
        subset=["EIC", "Číslo faktúry (VS)"], keep="first"
    ).copy()
    if "_sort_date" in inv_clean.columns:
      inv_clean = inv_clean.sort_values("_sort_date")
    inv_cols = [
        c
        for c in [
            "Mesiac",
            "Lokalita",
            "Odberné miesto",
            "Číslo faktúry (VS)",
            "Fakturačné obdobie",
            "Spotreba (kWh)",
            "Silová zložka (€)",
            "Cena silovej zložky (€/kWh)",
            "Distribúcia celkom (€)",
            "Prekročenie RK (€)",
            "Prekročenie MRK (€)",
            "Jalová elektrina (€)",
            "Prirážka za účinník (€)",
            "SPOLU POKUTY (€)",
            "Spolu bez DPH (€)",
        ]
        if c in inv_clean.columns
    ]
    inv_clean = inv_clean[inv_cols]
  else:
    inv_clean = pd.DataFrame()

  populate_formatted_sheet(
      sheet_name="2. Audit faktúr a pokút",
      df_table=inv_clean,
      heading="AUDIT ZSE FAKTÚR A SANKCIÍ 2026",
      sub_note=(
          "Rozpad fakturovaných položiek vrátane čistej ceny silovej zložky."
          " Červeným sú zvýraznené sankčné poplatky."
      ),
      add_sum_row=True,
      highlight_penalties=True,
  )

  # 3. Mesačný detail
  if not df_detail.empty:
    det_cols = [
        c
        for c in [
            "Mesiac",
            "Lokalita",
            "Odberné miesto",
            "Spotreba (kWh)",
            "Náklady (€)",
            "Cena (€/kWh)",
            "Cena silovej zložky (€/kWh)",
            "Zdroj",
            "Číslo faktúry (VS)",
            "Fakturačné obdobie",
        ]
        if c in df_detail.columns
    ]
    det_clean = df_detail[det_cols].copy()
  else:
    det_clean = pd.DataFrame()

  populate_formatted_sheet(
      sheet_name="3. Mesačný detail",
      df_table=det_clean,
      heading="DETAILNÝ MESAČNÝ PREHĽAD VŠETKÝCH MERANÍ",
      sub_note=(
          "Kompletný mesačný rozpis všetkých vybraných miest vrátane"
          " spravodlivo prepočítaných podružných okruhov."
      ),
      add_sum_row=False,
  )

  # 4. Výkon FVE
  if not df_fve.empty:
    fve_clean = df_fve.rename(
        columns={
            "Lokalita_raw": "Lokalita",
            "Vyroba_kWh": "Výroba FVE (kWh)",
            "Uspora_EUR": "Finančná úspora (€)",
            "Podiel_Pokrytia_Pct": "Podiel pokrytia (%)",
        }
    ).copy()
    populate_formatted_sheet(
        sheet_name="4. Výkon FVE",
        df_table=fve_clean,
        heading="VYHODNOTENIE FOTOVOLTIKY (FVE) 2026",
        sub_note=(
            "Mesačná výroba elektriny z FVE, finančná úspora a miera pokrytia"
            " spotreby."
        ),
        add_sum_row=True,
    )

  wb.save(output)
  return output.getvalue()


# --- PARSER ZSE FAKTÚR ---
@st.cache_data
def parse_zse_pdf(pdf_path, file_mtime=None):
  if not os.path.exists(pdf_path):
    return None

  with pdfplumber.open(pdf_path) as pdf:
    text = "\n".join([page.extract_text() or "" for page in pdf.pages])

  eic_match = re.search(r"(24Z[A-Z0-9]{13})", text)
  eic_code = eic_match.group(1) if eic_match else None

  if not eic_code or eic_code not in EIC_MAP:
    return None

  om_info = EIC_MAP[eic_code]

  vs_match = re.search(
      r"(?:Vyúčtovacia faktúra za elektrinu č\.|Variabilný symbol:)\s*(\d+)",
      text,
  )
  vs_num = vs_match.group(1) if vs_match else "Neuvedené"

  obdobie_match = re.search(
      r"Fakturačné obdobie\s*(\d{2}\.\d{2}\.\d{4})\s*-\s*(\d{2}\.\d{2}\.\d{4})",
      text,
  )
  if not obdobie_match:
    obdobie_match = re.search(
        r"(\d{2}\.\d{2}\.\d{4})\s*-\s*(\d{2}\.\d{2}\.\d{4})", text
    )

  obdobie = (
      f"{obdobie_match.group(1)} - {obdobie_match.group(2)}"
      if obdobie_match
      else "Neuvedené"
  )

  sort_date = obdobie_match.group(1).split(".") if obdobie_match else []
  sort_date_str = (
      f"{sort_date[2]}-{sort_date[1]}-{sort_date[0]}"
      if len(sort_date) == 3
      else "0000-00-00"
  )

  month_map = {
      "01": "Január",
      "02": "Február",
      "03": "Marec",
      "04": "Apríl",
      "05": "Máj",
      "06": "Jún",
      "07": "Júl",
      "08": "August",
      "09": "September",
      "10": "Október",
      "11": "November",
      "12": "December",
  }
  invoice_month = (
      month_map.get(sort_date[1], "Neuvedené")
      if len(sort_date) == 3
      else "Neuvedené"
  )

  def get_exact_line_cost(keywords):
    if isinstance(keywords, str):
      keywords = [keywords]
    for line in text.split("\n"):
      if any(kw.lower() in line.lower() for kw in keywords):
        if any(x in line.lower() for x in ["technické údaje", "stav"]):
          continue
        amounts = re.findall(r"(\d{1,3}(?:\s?\d{3})*,\d{2})", line)
        if amounts:
          val_str = amounts[-1].replace(" ", "").replace(",", ".")
          try:
            val = float(val_str)
            if 0 <= val < 1000000:
              return val
          except:
            pass
    return 0.0

  def get_penalty_cost(keywords):
    if isinstance(keywords, str):
      keywords = [keywords]
    for kw in keywords:
      for line in text.split("\n"):
        if kw.lower() in line.lower():
          if any(x in line.lower() for x in ["technické údaje", "stav"]):
            continue
          amounts = re.findall(r"(\d{1,3}(?:\s?\d{3})*,\d{2})", line)
          if amounts:
            val_str = amounts[-1].replace(" ", "").replace(",", ".")
            try:
              val = float(val_str)
              if 0 <= val < 1000000:
                return val
            except:
              pass
      matches = [
          m.start() for m in re.finditer(re.escape(kw), text, re.IGNORECASE)
      ]
      for pos in matches:
        window = text[pos : pos + 150]
        if any(x in window.lower() for x in ["technické údaje", "stav"]):
          continue
        amounts = re.findall(r"(\d{1,3}(?:\s?\d{3})*,\d{2})", window)
        if amounts:
          val_str = amounts[-1].replace(" ", "").replace(",", ".")
          try:
            val = float(val_str)
            if 0 <= val < 1000000:
              return val
          except:
            pass
    return 0.0

  def get_line_kwh(keyword):
    if isinstance(keyword, str):
      keyword = [keyword]
    for line in text.split("\n"):
      if any(kw.lower() in line.lower() for kw in keyword):
        if "technické údaje" in line.lower():
          continue
        m = re.search(r"([\d\s,]+)\s*kWh", line, re.IGNORECASE)
        if m:
          val_str = (
              m.group(1).split()[0].replace(" ", "").replace(",", ".")
          )
          try:
            val = float(val_str)
            if val < 10000000:
              return val
          except:
            pass
    return 0.0

  silova_kwh = get_line_kwh(["1T za elektrinu", "ČE 1T", "Dodávka elektriny"])
  silova_eur = get_exact_line_cost(["Dodávka elektriny"])
  if silova_eur == 0.0:
    silova_eur = get_exact_line_cost(
        ["1T za elektrinu", "Register-Special", "ČE 1T"]
    )

  cena_silovej_za_kwh = (
      (silova_eur / silova_kwh) if silova_kwh > 0 else 0.0
  )

  distribucia_eur = 0.0
  sumar_section = re.search(
      r"(?:Sumár za miesto spotreby|Sumár za faktúru).*?(?=Suma spolu|$)",
      text,
      re.IGNORECASE | re.DOTALL,
  )
  if sumar_section:
    distrib_match = re.search(
        r"Distribúcia elektriny[^\d]*([\d\s]+,\d{2})\s*€",
        sumar_section.group(0),
        re.IGNORECASE,
    )
    if distrib_match:
      try:
        distribucia_eur = float(
            distrib_match.group(1).replace(" ", "").replace(",", ".")
        )
      except:
        pass

  if distribucia_eur == 0.0:
    distrib_section = re.search(
        r"Služba:\s*Distribúcia.*?(Spolu bez DPH.*?€)",
        text,
        re.IGNORECASE | re.DOTALL,
    )
    if distrib_section:
      match_sum = re.search(
          r"Spolu bez DPH\s*([\d\s]+,\d{2})\s*€",
          distrib_section.group(1),
          re.IGNORECASE,
      )
      if match_sum:
        try:
          distribucia_eur = float(
              match_sum.group(1).replace(" ", "").replace(",", ".")
          )
        except:
          pass

  rk_prekrocenie = get_penalty_cost([
      "Prekročenie rezervovanej kapacity",
      "Prekročenie RK",
      "Tarifa za prekročenie rezervovanej kapacity",
  ])

  mrk_prekrocenie = get_penalty_cost([
      "Prekročenie maximálnej rezervovanej kapacity",
      "Tarifa za prekročenie MRK",
      "Prekročenie MRK",
  ])

  jalovina = get_penalty_cost([
      "Dodávka jalovej elektriny",
      "Nevyžiad.dod.kapacitnej zložky jal.en.",
      "Jalová elektrina",
      "Dodávka jalovej induktívnej",
      "Dodávka jalovej kapacitnej",
  ])

  ucinnik = get_penalty_cost([
      "Prirážka za nedodržanie účinníka",
      "Nedodržanie účinníka",
      "Prirážka za účinník",
      "Účinník",
  ])

  celkove_pokuty = rk_prekrocenie + mrk_prekrocenie + jalovina + ucinnik

  spolu_bez_dph = 0.0
  rekap_match = re.search(
      r"(?:Základ dane celkom|Spolu bez DPH celkom|Celkom bez DPH).*?(\d{1,3}(?:\s?\d{3})*,\d{2})",
      text,
      re.IGNORECASE,
  )
  if rekap_match:
    try:
      val = float(rekap_match.group(1).replace(" ", "").replace(",", "."))
      if 0 < val < 1000000:
        spolu_bez_dph = val
    except:
      pass

  if spolu_bez_dph == 0.0:
    spolu_bez_dph = silova_eur + distribucia_eur

  return {
      "EIC": eic_code,
      "Odberné miesto": om_info["nazov"],
      "Lokalita": om_info["lokalita"],
      "Číslo faktúry (VS)": vs_num,
      "Fakturačné obdobie": obdobie,
      "Mesiac": invoice_month,
      "_sort_date": sort_date_str,
      "Spotreba (kWh)": silova_kwh,
      "Silová zložka (€)": silova_eur,
      "Cena silovej zložky (€/kWh)": cena_silovej_za_kwh,
      "Distribúcia celkom (€)": distribucia_eur,
      "Prekročenie RK (€)": rk_prekrocenie,
      "Prekročenie MRK (€)": mrk_prekrocenie,
      "Jalová elektrina (€)": jalovina,
      "Prirážka za účinník (€)": ucinnik,
      "SPOLU POKUTY (€)": celkove_pokuty,
      "Spolu bez DPH (€)": spolu_bez_dph,
      "Raw_Text": text,
  }


# --- NAČÍTANIE FVE A EXCEL DÁT ---
@st.cache_data
def load_excel_data():
  file_path = "Spotreby energií.xlsx"
  if not os.path.exists(file_path):
    return pd.DataFrame(), pd.DataFrame()

  xls = pd.ExcelFile(file_path)
  months = [
      "Január",
      "Február",
      "Marec",
      "Apríl",
      "Máj",
      "Jún",
      "Júl",
      "August",
      "September",
      "Október",
      "November",
      "December",
  ]

  df_fve_raw = pd.read_excel(xls, sheet_name="Prehľad po FVE.2026")
  poprad_kwh = df_fve_raw.iloc[1, 2:11].values.astype(float)
  poprad_eur = df_fve_raw.iloc[2, 2:11].values.astype(float)
  poprad_pct = df_fve_raw.iloc[3, 2:11].values.astype(float)
  poprad_pct = [p if p <= 1.0 else p / 100.0 for p in poprad_pct]

  ba_kwh = df_fve_raw.iloc[4, 2:11].values.astype(float)
  ba_eur = df_fve_raw.iloc[5, 2:11].values.astype(float)
  ba_pct = df_fve_raw.iloc[6, 2:11].values.astype(float)
  ba_pct = [p if p <= 1.0 else p / 100.0 for p in ba_pct]

  fve_df = pd.DataFrame({
      "Mesiac": months[:9] * 2,
      "Lokalita_raw": ["Poprad"] * 9 + ["Bratislava"] * 9,
      "Vyroba_kWh": list(poprad_kwh) + list(ba_kwh),
      "Uspora_EUR": list(poprad_eur) + list(ba_eur),
      "Podiel_Pokrytia_Pct": list(poprad_pct) + list(ba_pct),
  })

  df_elec_raw = pd.read_excel(
      xls, sheet_name="Prehľad elektro.2026", header=None
  )
  records = []

  current_lokalita = "Poprad"
  for idx in range(len(df_elec_raw)):
    row = df_elec_raw.iloc[idx]
    val0 = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
    val1 = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
    combined_text = f"{val0} {val1}".strip()

    if (
        "nové trafo" in combined_text.lower()
        or "nove trafo" in combined_text.lower()
    ):
      continue

    if "bratislava" in combined_text.lower():
      current_lokalita = "Bratislava"
    elif "poprad" in combined_text.lower():
      current_lokalita = "Poprad"

    if "spotreba" in combined_text.lower():
      raw_lower = combined_text.lower()
      if "feld" in raw_lower or "bratislava" in raw_lower:
        clean_name = "Bratislava – Trafo Feld"
        lok = "Bratislava"
      else:
        lok = "Poprad"
        if "bananáreň" in raw_lower or "bananaren" in raw_lower:
          clean_name = "Bananáreň"
        elif "zemiakareň" in raw_lower or "zemiakaren" in raw_lower:
          clean_name = "Zemiakareň"
        elif (
            "správa 99" in raw_lower
            or "sprava 99" in raw_lower
            or "ab správa" in raw_lower
        ):
          clean_name = "Správa 99"
        elif "skladová hala" in raw_lower or "skladova hala" in raw_lower:
          clean_name = "Skladová hala I."
        elif "haramia" in raw_lower:
          clean_name = "Trafostanica HARAMIA"
        elif "pebek" in raw_lower:
          clean_name = "Trafostanica PEBEK"
        elif "beňovský" in raw_lower:
          clean_name = "Beňovský Apartments"
        elif "nájomcovia" in raw_lower:
          clean_name = "Hlavný elektromer / Nájomcovia"
        else:
          clean_name = (
              combined_text.replace("spotreba", "")
              .replace("kWh", "")
              .strip()
              .capitalize()
          )

      kwh_vals = pd.to_numeric(row.values[2 : len(months) + 2], errors="coerce")
      eur_vals = None
      for next_idx in range(idx + 1, min(idx + 4, len(df_elec_raw))):
        next_row = df_elec_raw.iloc[next_idx]
        next_text = f"{next_row.iloc[0]} {next_row.iloc[1]}".lower()
        if "náklad" in next_text or "naklad" in next_text:
          eur_vals = pd.to_numeric(
              next_row.values[2 : len(months) + 2], errors="coerce"
          )
          break

      if eur_vals is None:
        eur_vals = [0.0] * len(months)

      for m, k, e in zip(months, kwh_vals, eur_vals):
        k_val = float(k) if pd.notna(k) else 0.0
        e_val = float(e) if pd.notna(e) else 0.0
        if k_val > 0 or e_val > 0:
          records.append({
              "Mesiac": m,
              "Lokalita": lok,
              "Odberné miesto": clean_name,
              "Spotreba (kWh)": k_val,
              "Náklady (€)": e_val,
              "Cena (€/kWh)": (e_val / k_val) if k_val > 0 else 0.0,
              "Cena silovej zložky (€/kWh)": 0.0,
              "Zdroj": "Excel (Podružné meranie / Plán)",
          })

  excel_elec_df = pd.DataFrame(records)
  return fve_df, excel_elec_df


fve_df, excel_elec_df = load_excel_data()

# Načítanie PDF faktúr pre hlavné miesta
all_pdf_files = (
    [f for f in os.listdir(INVOICE_DIR) if f.lower().endswith(".pdf")]
    if (os.path.exists(INVOICE_DIR) and os.path.isdir(INVOICE_DIR))
    else []
)
parsed_invoice_records = []
for f in all_pdf_files:
  f_path = os.path.join(INVOICE_DIR, f)
  res = parse_zse_pdf(f_path, os.path.getmtime(f_path))
  if (
      res
      and res["Odberné miesto"]
      and "neznáme" not in res["Odberné miesto"].lower()
  ):
    parsed_invoice_records.append({
        "Mesiac": res["Mesiac"],
        "Lokalita": res["Lokalita"],
        "Odberné miesto": res["Odberné miesto"],
        "Spotreba (kWh)": res["Spotreba (kWh)"],
        "Náklady (€)": res["Spolu bez DPH (€)"],
        "Cena (€/kWh)": (res["Spolu bez DPH (€)"] / res["Spotreba (kWh)"])
        if res["Spotreba (kWh)"] > 0
        else 0.0,
        "Cena silovej zložky (€/kWh)": res["Cena silovej zložky (€/kWh)"],
        "Číslo faktúry (VS)": res["Číslo faktúry (VS)"],
        "Fakturačné obdobie": res["Fakturačné obdobie"],
        "_sort_date": res["_sort_date"],
        "Raw_Text": res["Raw_Text"],
        "Zdroj": "PDF Faktúra (ZSE)",
        **res,
    })

pdf_elec_df = pd.DataFrame(parsed_invoice_records)

# --- SPOJENIE DÁT S PRÍSNYM VYNECHANÍM EXCELU PRE HLAVNÉ OM, AK EXISTUJE PDF ---
if not pdf_elec_df.empty:
  oms_with_pdf = set(pdf_elec_df["Odberné miesto"].unique())
  if not excel_elec_df.empty:
    excel_elec_df = excel_elec_df[~excel_elec_df["Odberné miesto"].isin(oms_with_pdf)]
  elec_df = pd.concat([pdf_elec_df, excel_elec_df], ignore_index=True)
else:
  elec_df = excel_elec_df

# --- DYNAMICKÝ A SPRAVODLIVÝ PREPOČET PODRUŽNÝCH MERANÍ PODĽA NADRADENEJ TRAFOSTANICE ---
month_order = [
    "Január",
    "Február",
    "Marec",
    "Apríl",
    "Máj",
    "Jún",
    "Júl",
    "August",
    "September",
    "Október",
    "November",
    "December",
]

parent_mapping = {
    "Skladová hala I.": "Trafostanica PEBEK",
    "Správa 99": "Trafostanica PEBEK",
    "Bananáreň": "Trafostanica HARAMIA",
    "Zemiakareň": "Trafostanica HARAMIA",
}

for sub_om, parent_om in parent_mapping.items():
  for m in month_order:
    parent_row = elec_df[
        (elec_df["Odberné miesto"] == parent_om)
        & (elec_df["Mesiac"] == m)
        & (elec_df["Zdroj"].str.contains("PDF"))
    ]
    if not parent_row.empty:
      p_kwh = parent_row["Spotreba (kWh)"].values[0]
      p_eur = parent_row["Náklady (€)"].values[0]
      if p_kwh > 0:
        parent_unit_price = p_eur / p_kwh
        mask = (elec_df["Odberné miesto"] == sub_om) & (elec_df["Mesiac"] == m)
        if mask.any():
          sub_kwh = elec_df.loc[mask, "Spotreba (kWh)"].values[0]
          new_cost = sub_kwh * parent_unit_price
          elec_df.loc[mask, "Náklady (€)"] = new_cost
          elec_df.loc[mask, "Cena (€/kWh)"] = parent_unit_price
          elec_df.loc[mask, "Zdroj"] = (
              f"Prepočítané podľa {parent_om} (PDF cena)"
          )

# --- STAV A HIERARCHICKÝ VÝBER ODBERNÝCH MIEST ---
TAB_NAMES = [
    "📉 Spotreby & Náklady Odberných Miest",
    "☀️ Výkon Fotovoltiky (FVE)",
    "💶 Cena za 1 kWh",
    "📄 Analýza PDF Faktúr & Audit",
]

if "active_tab" not in st.session_state:
  st.session_state.active_tab = TAB_NAMES[0]


def set_tab(tab_name):
  st.session_state.active_tab = tab_name


all_possible_oms = (
    sorted(elec_df["Odberné miesto"].dropna().astype(str).unique())
    if not elec_df.empty
    else []
)

OM_GROUPS = {
    "Trafostanica PEBEK": ["Skladová hala I.", "Správa 99"],
    "Trafostanica HARAMIA": ["Bananáreň", "Zemiakareň"],
}

if "selected_oms" not in st.session_state:
  st.session_state.selected_oms = list(all_possible_oms)
else:
  st.session_state.selected_oms = [
      x for x in st.session_state.selected_oms if x in all_possible_oms
  ]


def _om_key(prefix, name):
  return prefix + "_" + re.sub(r"[^A-Za-z0-9]+", "_", name).strip("_")


def _set_group_selection(parent, children):
  selected = set(st.session_state.get("selected_oms", []))
  group_checked = st.session_state.get(_om_key("grp", parent), False)
  if group_checked:
    selected.add(parent)
    selected.update(children)
  else:
    selected.discard(parent)
    selected.difference_update(children)
  for child in children:
    if child in all_possible_oms:
      st.session_state[_om_key("om", child)] = group_checked
  st.session_state.selected_oms = [x for x in all_possible_oms if x in selected]


def _set_child_selection(child, parent, children):
  selected = set(st.session_state.get("selected_oms", []))
  if st.session_state.get(_om_key("om", child), False):
    selected.add(child)
  else:
    selected.discard(child)
  all_children_selected = all(x in selected for x in children)
  parent_key = _om_key("grp", parent)
  st.session_state[parent_key] = all_children_selected
  if all_children_selected:
    selected.add(parent)
  else:
    selected.discard(parent)
  st.session_state.selected_oms = [x for x in all_possible_oms if x in selected]


def _sync_single_om(name):
  selected = set(st.session_state.get("selected_oms", []))
  if st.session_state.get(_om_key("om", name), False):
    selected.add(name)
  else:
    selected.discard(name)
  st.session_state.selected_oms = [x for x in all_possible_oms if x in selected]


for parent, children in OM_GROUPS.items():
  existing_children = [x for x in children if x in all_possible_oms]
  if parent in all_possible_oms:
    parent_key = _om_key("grp", parent)
    if parent_key not in st.session_state:
      st.session_state[parent_key] = parent in st.session_state.selected_oms or (
          bool(existing_children)
          and all(x in st.session_state.selected_oms for x in existing_children)
      )
    for child in existing_children:
      child_key = _om_key("om", child)
      if child_key not in st.session_state:
        st.session_state[child_key] = child in st.session_state.selected_oms

for om in all_possible_oms:
  if not any(
      om == parent or om in children for parent, children in OM_GROUPS.items()
  ):
    key = _om_key("om", om)
    if key not in st.session_state:
      st.session_state[key] = om in st.session_state.selected_oms

st.sidebar.subheader("🏢 Výber odberných miest")
if all_possible_oms:
  col_s1, col_s2 = st.sidebar.columns(2)
  if col_s1.button("Vybrať všetko", key="btn_all_oms", use_container_width=True):
    st.session_state.selected_oms = list(all_possible_oms)
    for parent, children in OM_GROUPS.items():
      st.session_state[_om_key("grp", parent)] = True
      for child in children:
        if child in all_possible_oms:
          st.session_state[_om_key("om", child)] = True
    for om in all_possible_oms:
      st.session_state[_om_key("om", om)] = True
    st.rerun()
  if col_s2.button("Zrušiť všetko", key="btn_none_oms", use_container_width=True):
    st.session_state.selected_oms = []
    for parent, children in OM_GROUPS.items():
      st.session_state[_om_key("grp", parent)] = False
    for om in all_possible_oms:
      st.session_state[_om_key("om", om)] = False
    st.rerun()

  st.sidebar.markdown("**📍 Poprad**")
  poprad_oms = (
      set(
          elec_df.loc[elec_df["Lokalita"] == "Poprad", "Odberné miesto"]
          .dropna()
          .astype(str)
      )
      if not elec_df.empty
      else set()
  )
  for parent, children in OM_GROUPS.items():
    if parent not in all_possible_oms:
      continue
    visible_children = [x for x in children if x in all_possible_oms]
    if parent in poprad_oms:
      st.sidebar.checkbox(
          f"⚡ {parent}",
          key=_om_key("grp", parent),
          on_change=_set_group_selection,
          args=(parent, visible_children),
      )
      if visible_children:
        for child in visible_children:
          st.sidebar.checkbox(
              f" ↳ {child}",
              key=_om_key("om", child),
              on_change=_set_child_selection,
              args=(child, parent, visible_children),
          )

  grouped_names = {parent for parent in OM_GROUPS}
  grouped_children = {
      child for children in OM_GROUPS.values() for child in children
  }
  standalone_poprad = sorted(
      x
      for x in poprad_oms
      if x not in grouped_names and x not in grouped_children
  )
  for om in standalone_poprad:
    st.sidebar.checkbox(
        f"• {om}",
        key=_om_key("om", om),
        on_change=lambda name=om: _sync_single_om(name),
    )

  bratislava_oms = (
      sorted(
          set(
              elec_df.loc[elec_df["Lokalita"] == "Bratislava", "Odberné miesto"]
              .dropna()
              .astype(str)
          )
      )
      if not elec_df.empty
      else []
  )
  if bratislava_oms:
    st.sidebar.markdown("**📍 Bratislava**")
    for om in bratislava_oms:
      st.sidebar.checkbox(
          f"• {om}",
          key=_om_key("om", om),
          on_change=lambda name=om: _sync_single_om(name),
      )
else:
  st.sidebar.info("Žiadne dáta nie sú k dispozícii.")

st.sidebar.divider()
st.sidebar.subheader("🔍 Časové filtre")
month_to_num = {name: i + 1 for i, name in enumerate(month_order)}
num_to_month = {v: k for k, v in month_to_num.items()}
months_in_data = (
    set(elec_df["Mesiac"].dropna().astype(str)) if not elec_df.empty else set()
)
if not fve_df.empty:
  months_in_data.update(fve_df["Mesiac"].dropna().astype(str))
available_month_nums = [
    month_to_num[m] for m in month_order if m in months_in_data
]
if not available_month_nums:
  available_month_nums = list(range(1, 13))

if "global_month_filter" not in st.session_state:
  st.session_state.global_month_filter = list(available_month_nums)
else:
  st.session_state.global_month_filter = [
      m
      for m in st.session_state.global_month_filter
      if m in available_month_nums
  ]
selected_month_nums = st.sidebar.multiselect(
    "📅 Vyber mesiace (2026)",
    options=available_month_nums,
    format_func=lambda m: num_to_month[m],
    key="global_month_filter",
)
selected_months = [num_to_month[m] for m in selected_month_nums]

st.sidebar.divider()

# --- MANAŽÉRSKE REPORTY V SIDEBARE ---
st.sidebar.subheader("📥 Manažérske Reporty & Exporty")

rep_elec = (
    elec_df[elec_df["Mesiac"].isin(selected_months)].copy()
    if not elec_df.empty and selected_months
    else pd.DataFrame()
)
if not rep_elec.empty and st.session_state.selected_oms:
  rep_elec = rep_elec[
      rep_elec["Odberné miesto"].isin(st.session_state.selected_oms)
  ]
else:
  rep_elec = pd.DataFrame()

filtered_elec = rep_elec.copy()

selected_lokality = (
    filtered_elec["Lokalita"].unique().tolist()
    if not filtered_elec.empty
    else []
)
filtered_fve = (
    fve_df[
        fve_df["Mesiac"].isin(selected_months)
        & fve_df["Lokalita_raw"].isin(selected_lokality)
    ].copy()
    if not fve_df.empty and selected_months and selected_lokality
    else pd.DataFrame(columns=fve_df.columns)
)

if not filtered_fve.empty:
  filtered_fve["Mesiac"] = pd.Categorical(
      filtered_fve["Mesiac"], categories=month_order, ordered=True
  )
  filtered_fve = filtered_fve.sort_values("Mesiac")

if not filtered_elec.empty:
  filtered_elec["Mesiac"] = pd.Categorical(
      filtered_elec["Mesiac"], categories=month_order, ordered=True
  )
  filtered_elec = filtered_elec.sort_values("Mesiac")

kpi_df = filtered_elec.copy()
if not kpi_df.empty:
  if "Trafostanica PEBEK" in kpi_df["Odberné miesto"].values:
    kpi_df = kpi_df[
        ~kpi_df["Odberné miesto"].isin(["Skladová hala I.", "Správa 99"])
    ]
  if "Trafostanica HARAMIA" in kpi_df["Odberné miesto"].values:
    kpi_df = kpi_df[
        ~kpi_df["Odberné miesto"].isin(["Bananáreň", "Zemiakareň"])
    ]

pdf_filtered_for_export = pd.DataFrame()
if not pdf_elec_df.empty and selected_months and st.session_state.selected_oms:
  pdf_filtered_for_export = pdf_elec_df[
      pdf_elec_df["Mesiac"].isin(selected_months)
      & pdf_elec_df["Odberné miesto"].isin(st.session_state.selected_oms)
  ].copy()

if not filtered_elec.empty:
  excel_data = build_presentation_excel(
      df_detail=filtered_elec,
      df_kpi_clean=kpi_df,
      df_invoices=pdf_filtered_for_export,
      df_fve=filtered_fve,
      selected_months_list=selected_months,
      selected_oms_list=st.session_state.selected_oms,
      report_title="ENERGETICKÝ REPORT 2026 (MANAŽÉRSKY PREHĽAD)",
  )

  st.sidebar.download_button(
      label="📊 Stiahnuť manažérsky Excel report",
      data=excel_data,
      file_name="LUNYS_Energeticky_Report_2026.xlsx",
      mime=(
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      ),
      use_container_width=True,
  )

st.sidebar.divider()
st.sidebar.subheader("⚡ Rýchla navigácia")
if st.sidebar.button(
    "📉 Spotreby & Náklady",
    use_container_width=True,
    on_click=set_tab,
    args=(TAB_NAMES[0],),
):
  pass
if st.sidebar.button(
    "☀️ Výkon FVE",
    use_container_width=True,
    on_click=set_tab,
    args=(TAB_NAMES[1],),
):
  pass
if st.sidebar.button(
    "💶 Cena za 1 kWh",
    use_container_width=True,
    on_click=set_tab,
    args=(TAB_NAMES[2],),
):
  pass
if st.sidebar.button(
    "📄 Analýza PDF Faktúr",
    use_container_width=True,
    on_click=set_tab,
    args=(TAB_NAMES[3],),
):
  pass

# --- HLAVNÉ KPI KARTY ---
total_elec_kwh = (
    kpi_df["Spotreba (kWh)"].sum() if not kpi_df.empty else 0.0
)
total_elec_eur = (
    kpi_df["Náklady (€)"].sum() if not kpi_df.empty else 0.0
)
avg_price_per_kwh = (
    (total_elec_eur / total_elec_kwh) if total_elec_kwh > 0 else 0.0
)
total_fve_kwh = (
    filtered_fve["Vyroba_kWh"].sum() if not filtered_fve.empty else 0.0
)
total_fve_eur = (
    filtered_fve["Uspora_EUR"].sum() if not filtered_fve.empty else 0.0
)

st.subheader("📊 Kľúčové ukazovatele 2026")
c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("🔌 Odber zo siete (Spolu)", f"{total_elec_kwh:,.0f} kWh")
c2.metric("💶 Náklady na sieť spolu", f"{total_elec_eur:,.2f} €")
c3.metric("☀️ Výroba FVE", f"{total_fve_kwh:,.0f} kWh")
c4.metric("💰 Úspora FVE", f"{total_fve_eur:,.2f} €")
c5.metric("⚡ Priemerná cena 1 kWh", f"{avg_price_per_kwh:.4f} €/kWh")

st.divider()

selected_tab = st.radio(
    "Zvoľ sekciu dashboardu:",
    TAB_NAMES,
    key="active_tab",
    horizontal=True,
    label_visibility="collapsed",
)
st.divider()

if selected_tab == TAB_NAMES[0]:
  st.subheader("Mesačný vývoj spotreby a nákladov")
  if not filtered_elec.empty:
    col_left, col_right = st.columns(2)
    with col_left:
      fig_elec_kwh = px.bar(
          filtered_elec,
          x="Mesiac",
          y="Spotreba (kWh)",
          color="Odberné miesto",
          barmode="group",
          title="Spotreba elektriny (kWh)",
          template="plotly_dark",
      )
      fig_elec_kwh.update_layout(
          legend=dict(
              orientation="h", yanchor="top", y=-0.3, xanchor="center", x=0.5
          )
      )
      st.plotly_chart(fig_elec_kwh, use_container_width=True)
    with col_right:
      fig_elec_eur = px.bar(
          filtered_elec,
          x="Mesiac",
          y="Náklady (€)",
          color="Odberné miesto",
          barmode="group",
          title="Mesačné náklady bez DPH (€)",
          template="plotly_dark",
      )
      fig_elec_eur.update_layout(
          legend=dict(
              orientation="h", yanchor="top", y=-0.3, xanchor="center", x=0.5
          )
      )
      st.plotly_chart(fig_elec_eur, use_container_width=True)

    st.markdown("### 📌 Súhrn pre vybraný filter grafov")
    sc1, sc2, sc3 = st.columns(3)
    sum_chart_kwh = filtered_elec["Spotreba (kWh)"].sum()
    sum_chart_eur = filtered_elec["Náklady (€)"].sum()
    avg_chart_price = (
        sum_chart_eur / sum_chart_kwh if sum_chart_kwh > 0 else 0.0
    )

    sc1.metric("Celková spotreba (Spolu)", f"{sum_chart_kwh:,.0f} kWh")
    sc2.metric("Celkové náklady (Spolu)", f"{sum_chart_eur:,.2f} €")
    sc3.metric("Priemerná cena", f"{avg_chart_price:.4f} €/kWh")

    st.divider()
    st.dataframe(
        filtered_elec[[
            "Mesiac",
            "Odberné miesto",
            "Lokalita",
            "Spotreba (kWh)",
            "Náklady (€)",
            "Cena (€/kWh)",
            "Zdroj",
        ]],
        use_container_width=True,
        hide_index=True,
    )
  else:
    st.info("Pre zvolené filtre nie sú k dispozícii žiadne dáta.")

elif selected_tab == TAB_NAMES[1]:
  st.subheader("☀️ Výkon a úspory FVE")
  fve_lokalita_filter = st.radio(
      "📍 Vyber lokalitu pre FVE:",
      options=["Všetky lokality", "Poprad", "Bratislava"],
      horizontal=True,
      key="fve_lok_filter",
  )

  tab_fve_df = filtered_fve.copy()
  if fve_lokalita_filter != "Všetky lokality":
    tab_fve_df = tab_fve_df[tab_fve_df["Lokalita_raw"] == fve_lokalita_filter]

  if not tab_fve_df.empty:
    summary_fve = (
        tab_fve_df.groupby("Lokalita_raw")
        .agg({
            "Vyroba_kWh": "sum",
            "Uspora_EUR": "sum",
            "Podiel_Pokrytia_Pct": "mean",
        })
        .reset_index()
    )

    pie_labels_kwh = [
        f"{row['Lokalita_raw']}: {row['Vyroba_kWh']:,.0f} kWh"
        for _, row in summary_fve.iterrows()
    ]
    pie_labels_eur = [
        f"{row['Lokalita_raw']}: {row['Uspora_EUR']:,.0f} €"
        for _, row in summary_fve.iterrows()
    ]
    summary_fve["Label_KWh"] = pie_labels_kwh
    summary_fve["Label_EUR"] = pie_labels_eur

    col_sum1, col_sum2 = st.columns(2)
    with col_sum1:
      fig_pie_kwh = px.pie(
          summary_fve,
          names="Label_KWh",
          values="Vyroba_kWh",
          title="Podiel lokalít na celkovej výrobe FVE",
          hole=0.4,
          template="plotly_dark",
      )
      st.plotly_chart(fig_pie_kwh, use_container_width=True)
    with col_sum2:
      fig_pie_eur = px.pie(
          summary_fve,
          names="Label_EUR",
          values="Uspora_EUR",
          title="Podiel lokalít na finančných úsporách",
          hole=0.4,
          template="plotly_dark",
      )
      st.plotly_chart(fig_pie_eur, use_container_width=True)

    st.dataframe(tab_fve_df, use_container_width=True, hide_index=True)
  else:
    st.info("FVE dáta pre zvolenú lokalitu nie sú k dispozícii.")

elif selected_tab == TAB_NAMES[2]:
  st.subheader("💶 Porovnanie jednotkových cien za 1 kWh")
  if not filtered_elec.empty:
    allowed_main_oms = [
        "Trafostanica PEBEK",
        "Trafostanica HARAMIA",
        "Hlavný elektromer / Nájomcovia",
        "Beňovský Apartments",
        "Bratislava – Trafo Feld",
    ]
    df_price_plots = filtered_elec[
        filtered_elec["Odberné miesto"].isin(allowed_main_oms)
    ].copy()

    col_p1, col_p2 = st.columns(2)
    with col_p1:
      if not df_price_plots.empty:
        fig_price = px.line(
            df_price_plots,
            x="Mesiac",
            y="Cena (€/kWh)",
            color="Odberné miesto",
            markers=True,
            title=(
                "Celková jednotková cena (€/kWh) – Všetkých 5 hlavných miest"
            ),
            template="plotly_dark",
        )
        fig_price.update_layout(
            xaxis={"categoryorder": "array", "categoryarray": month_order},
            legend=dict(
                orientation="h", yanchor="top", y=-0.3, xanchor="center", x=0.5
            ),
        )
        st.plotly_chart(fig_price, use_container_width=True)

    with col_p2:
      df_silova_plot = df_price_plots.copy()
      if "Cena silovej zložky (€/kWh)" in df_silova_plot.columns:
        df_silova_plot["Cena silovej zložky (€/kWh)"] = df_silova_plot[
            "Cena silovej zložky (€/kWh)"
        ].fillna(0.0)

      if not df_silova_plot.empty:
        fig_silova = px.line(
            df_silova_plot,
            x="Mesiac",
            y="Cena silovej zložky (€/kWh)",
            color="Odberné miesto",
            markers=True,
            title=(
                "Čistá cena silovej zložky (€/kWh) – Všetkých 5 hlavných miest"
            ),
            template="plotly_dark",
        )
        fig_silova.update_layout(
            xaxis={"categoryorder": "array", "categoryarray": month_order},
            legend=dict(
                orientation="h", yanchor="top", y=-0.3, xanchor="center", x=0.5
            ),
        )
        st.plotly_chart(fig_silova, use_container_width=True)

    st.divider()
    st.markdown(
        "### 📋 Prehľad hlavných odberných miest (bez podružných meraní)"
    )

    sub_to_exclude = ["Skladová hala I.", "Správa 99", "Bananáreň", "Zemiakareň"]
    df_table_clean = filtered_elec[
        ~filtered_elec["Odberné miesto"].isin(sub_to_exclude)
    ].copy()

    st.dataframe(
        df_table_clean[[
            "Mesiac",
            "Odberné miesto",
            "Spotreba (kWh)",
            "Náklady (€)",
            "Cena (€/kWh)",
            "Cena silovej zložky (€/kWh)",
            "Zdroj",
        ]],
        use_container_width=True,
        hide_index=True,
    )
  else:
    st.info("Pre zvolené filtre nie sú k dispozícii žiadne dáta.")

elif selected_tab == TAB_NAMES[3]:
  st.subheader("📄 Analýza PDF faktúr zo ZSE & Audit")

  pdf_only_df = pdf_elec_df.copy()
  if not pdf_only_df.empty:
    if selected_months:
      pdf_only_df = pdf_only_df[pdf_only_df["Mesiac"].isin(selected_months)]
    if st.session_state.selected_oms:
      pdf_only_df = pdf_only_df[
          pdf_only_df["Odberné miesto"].isin(st.session_state.selected_oms)
      ]

  if not pdf_only_df.empty:
    df_filtered_inv = (
        pdf_only_df.drop_duplicates(
            subset=["EIC", "Číslo faktúry (VS)"], keep="first"
        )
        .sort_values(by="_sort_date", ascending=True)
        .reset_index(drop=True)
    )

    st.success(f"Zobrazených jedinečných faktúr: {len(df_filtered_inv)}")
    st.divider()

    st.markdown("### 📈 Súhrnný prehľad nákladov a sankcií z faktúr")
    col_g1, col_g2 = st.columns(2)

    with col_g1:
      tot_silova = df_filtered_inv["Silová zložka (€)"].sum()
      tot_distrib = df_filtered_inv["Distribúcia celkom (€)"].sum()
      sum_cost_total = df_filtered_inv["Spolu bez DPH (€)"].sum()

      df_pie_cost = pd.DataFrame({
          "Zložka": [
              f"Silová zložka ({tot_silova:,.2f} €)",
              f"Distribúcia celkom ({tot_distrib:,.2f} €)",
          ],
          "Suma": [tot_silova, tot_distrib],
      })
      fig_pie1 = px.pie(
          df_pie_cost,
          names="Zložka",
          values="Suma",
          hole=0.4,
          template="plotly_dark",
      )
      st.plotly_chart(fig_pie1, use_container_width=True)
      st.markdown(f"**Spolu fakturované náklady bez DPH: {sum_cost_total:,.2f} €**")

    with col_g2:
      tot_rk = df_filtered_inv["Prekročenie RK (€)"].sum()
      tot_mrk = df_filtered_inv["Prekročenie MRK (€)"].sum()
      tot_jalova = df_filtered_inv["Jalová elektrina (€)"].sum()
      tot_ucinnik = df_filtered_inv["Prirážka za účinník (€)"].sum()
      sum_penalties_total = tot_rk + tot_mrk + tot_jalova + tot_ucinnik

      df_pie_penalties = pd.DataFrame({
          "Typ sankcie": [
              f"Prekročenie RK ({tot_rk:,.2f} €)",
              f"Prekročenie MRK ({tot_mrk:,.2f} €)",
              f"Jalová elektrina ({tot_jalova:,.2f} €)",
              f"Prirážka za účinník ({tot_ucinnik:,.2f} €)",
          ],
          "Suma": [tot_rk, tot_mrk, tot_jalova, tot_ucinnik],
      })
      fig_pie2 = px.pie(
          df_pie_penalties,
          names="Typ sankcie",
          values="Suma",
          hole=0.4,
          template="plotly_dark",
      )
      st.plotly_chart(fig_pie2, use_container_width=True)
      st.markdown(f"**Spolu pokuty a sankcií: {sum_penalties_total:,.2f} €**")

    st.divider()

    st.markdown("### 📋 Prehľad jednotlivých faktúr")
    display_cols = [
        c
        for c in df_filtered_inv.columns
        if c not in ["Raw_Text", "_sort_date", "Mesiac", "Raw_Data"]
    ]
    st.dataframe(
        df_filtered_inv[display_cols], use_container_width=True, hide_index=True
    )

    st.divider()

    with st.expander("📥 Nahrať nové PDF faktúry a správa súborov", expanded=False):
      uploaded_files = st.file_uploader(
          "Vyber PDF faktúry zo ZSE na nahranie",
          type=["pdf"],
          accept_multiple_files=True,
          key="pdf_uploader_tab",
      )

      if uploaded_files:
        for uploaded_file in uploaded_files:
          target_path = os.path.join(INVOICE_DIR, uploaded_file.name)
          if not os.path.exists(target_path):
            with open(target_path, "wb") as f:
              f.write(uploaded_file.getbuffer())
        st.success("Súbory boli úspešne nahrané.")
        st.rerun()

      existing_pdfs = (
          [f for f in os.listdir(INVOICE_DIR) if f.lower().endswith(".pdf")]
          if os.path.exists(INVOICE_DIR)
          else []
      )
      if existing_pdfs:
        st.markdown(f"**Uložené PDF súbory v priečinku ({len(existing_pdfs)}):**")
        for pdf_file in existing_pdfs:
          col_f1, col_f2 = st.columns([5, 1])
          col_f1.text(pdf_file)
          if col_f2.button("🗑️ Zmazať", key=f"del_{pdf_file}"):
            try:
              os.remove(os.path.join(INVOICE_DIR, pdf_file))
              st.rerun()
            except Exception as e:
              st.error(f"Chyba pri mazaní: {e}")
      else:
        st.info(
            "Zatiaľ nie sú v priečinku `faktury_zse` nahrané žiadne PDF súbory."
        )

    st.divider()

    st.markdown("### 🔎 Detail faktúr podľa mesiacov")
    for idx, row in df_filtered_inv.iterrows():
      st.markdown(
          f"#### Faktúra č. {row['Číslo faktúry (VS)']} – {row['Odberné miesto']}"
          f" ({row['Fakturačné obdobie']})"
      )

      c_f1, c_f2, c_f3, c_f4 = st.columns(4)
      c_f1.metric(
          "Cena silovej zložky",
          f"{row['Cena silovej zložky (€/kWh)']:,.4f} €/kWh",
      )
      c_f2.metric(
          "Distribúcia celkom", f"{row['Distribúcia celkom (€)']:,.2f} €"
      )
      c_f3.metric(
          "Prekročenie RK / MRK",
          f"{(row['Prekročenie RK (€)'] + row['Prekročenie MRK (€)']):,.2f} €",
      )
      c_f4.metric(
          "Jalová + Účinník",
          f"{(row['Jalová elektrina (€)'] + row['Prirážka za účinník (€)']):,.2f} €",
      )

      with st.expander(
          "🔍 Diagnostika surového textu z PDF pre faktúru"
          f" {row['Číslo faktúry (VS)']}"
      ):
        st.text(row["Raw_Text"])
  else:
    st.info("Pre zvolené filtre neboli nájdené žiadne spracované PDF faktúry.")

    st.divider()
    with st.expander("📥 Nahrať nové PDF faktúry a správa súborov", expanded=True):
      uploaded_files = st.file_uploader(
          "Vyber PDF faktúry zo ZSE na nahranie",
          type=["pdf"],
          accept_multiple_files=True,
          key="pdf_uploader_tab_empty",
      )

      if uploaded_files:
        for uploaded_file in uploaded_files:
          target_path = os.path.join(INVOICE_DIR, uploaded_file.name)
          if not os.path.exists(target_path):
            with open(target_path, "wb") as f:
              f.write(uploaded_file.getbuffer())
        st.success("Súbory boli úspešne nahrané.")
        st.rerun()

      existing_pdfs = (
          [f for f in os.listdir(INVOICE_DIR) if f.lower().endswith(".pdf")]
          if os.path.exists(INVOICE_DIR)
          else []
      )
      if existing_pdfs:
        st.markdown(f"**Uložené PDF súbory v priečinku ({len(existing_pdfs)}):**")
        for pdf_file in existing_pdfs:
          col_f1, col_f2 = st.columns([5, 1])
          col_f1.text(pdf_file)
          if col_f2.button("🗑️ Zmazať", key=f"del_empty_{pdf_file}"):
            try:
              os.remove(os.path.join(INVOICE_DIR, pdf_file))
              st.rerun()
            except Exception as e:
              st.error(f"Chyba pri mazaní: {e}")
