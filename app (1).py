import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os
import re
import pdfplumber

# Nastavenie stránky
st.set_page_config(
    page_title="Energetický Dashboard LUNYS 2026",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- VYTVORENIE PRIEČINKA PRE TRVALÉ UKLADANIE FAKTÚR ---
INVOICE_DIR = "faktury_zse"
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
st.markdown("Interaktívny nástroj pre vyhodnotenie spotreby elektriny, nákladov, úspor z FVE (Poprad & Bratislava) a priemernej ceny za 1 kWh.")

# --- MAPOVANIE EIC KÓDOV NA ODBERNÉ MIESTA LUNYS ---
EIC_MAP = {
    "24ZVS00007394912": {"nazov": "Trafostanica PEBEK", "lokalita": "Poprad"},
    "24ZZS4000182286S": {"nazov": "Trafo Bratislava", "lokalita": "Bratislava"},
    "24ZZS40001822865": {"nazov": "Trafo Bratislava", "lokalita": "Bratislava"},
    "24ZVS0000016878H": {"nazov": "Beňovský Apartments", "lokalita": "Poprad"},
    "24ZVS0000059036Q": {"nazov": "Areál / Nájomcovia (Poprad)", "lokalita": "Poprad"},
    "24ZVS00000590360": {"nazov": "Areál / Nájomcovia (Poprad)", "lokalita": "Poprad"},
    "24ZVS0000787041D": {"nazov": "Trafostanica HARAMIA", "lokalita": "Poprad"}
}

# --- PRESNÝ A BEZPEČNÝ PARSER ZSE FAKTÚR ---
def parse_zse_pdf(pdf_path_or_file):
    with pdfplumber.open(pdf_path_or_file) as pdf:
        text = "\n".join([page.extract_text() or "" for page in pdf.pages])

    # 1. Detekcia EIC kódu
    eic_match = re.search(r'(24Z[A-Z0-9]{13})', text)
    eic_code = eic_match.group(1) if eic_match else "Neznáme EIC"
    om_info = EIC_MAP.get(eic_code, {"nazov": f"Neznáme odberné miesto ({eic_code})", "lokalita": "Neznáma"})

    # 2. Detekcia čísla faktúry a obdobia
    vs_match = re.search(r'(?:Vyúčtovacia faktúra za elektrinu č\.|Variabilný symbol:)\s*(\d+)', text)
    vs_num = vs_match.group(1) if vs_match else "Neuvedené"

    obdobie_match = re.search(r'Fakturačné obdobie\s*(\d{2}\.\d{2}\.\d{4})\s*-\s*(\d{2}\.\d{2}\.\d{4})', text)
    obdobie = f"{obdobie_match.group(1)} - {obdobie_match.group(2)}" if obdobie_match else "Neuvedené"

    def extract_euro_amount(line_start_text):
        """Bezpečne vyťahuje presnú sumu v EUR na konci riadku daného textu"""
        # Hľadá text položky a na konci toho istého riadku zoberie sumu vo formáte X XXX,XX
        pattern = re.escape(line_start_text) + r'.*?([\d\s]+,\d{2})\s*(?:€)?\s*$'
        m = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
        if m:
            val_str = m.group(1).replace(" ", "").replace(",", ".")
            try:
                return float(val_str)
            except:
                pass
        return 0.0

    # Spotreba v kWh (hľadáme množstvo pred kWh v riadku 1T)
    kwh_match = re.search(r'1T za elektrinu\s+([\d\s,]+)\s*kWh', text, re.IGNORECASE)
    if not kwh_match:
        kwh_match = re.search(r'ČE 1T vrát\.str reg\.mer-účt\..*?([\d\s,]+)\s*kWh', text, re.IGNORECASE)
    
    silova_kwh = 0.0
    if kwh_match:
        try:
            silova_kwh = float(kwh_match.group(1).split()[0].replace(" ", "").replace(",", "."))
        except:
            pass

    # Silová zložka v EUR (suma z riadku 1T za elektrinu)
    silova_eur = extract_euro_amount('1T za elektrinu')

    # Celková suma bez DPH pre celú faktúru
    spolu_match = re.search(r'Suma spolu\s*([\d\s,]+)\s*€', text, re.IGNORECASE)
    spolu_bez_dph = 0.0
    if spolu_match:
        try:
            spolu_bez_dph = float(spolu_match.group(1).replace(" ", "").replace(",", "."))
        except:
            pass

    # Distribúcia celkom: hľadáme súčet pod sekciou Distribúcia ("Spolu bez DPH")
    distribucia_eur = 0.0
    distrib_section = re.search(r'Služba:\s*Distribúcia.*?(Spolu bez DPH.*?€)', text, re.IGNORECASE | re.DOTALL)
    if distrib_section:
        sec_text = distrib_section.group(1)
        nums = re.findall(r'([\d\s]+,\d{2})\s*€', sec_text)
        if nums:
            try:
                distribucia_eur = float(nums[-1].replace(" ", "").replace(",", "."))
            except:
                pass

    if distribucia_eur == 0.0 and spolu_bez_dph > 0 and silova_eur > 0:
        distribucia_eur = max(0.0, spolu_bez_dph - silova_eur)

    # --- POKUTY A SANKČNÉ POLOŽKY ---
    mrk_prekrocenie = extract_euro_amount('Tarifa za prekročenie MRK')
    if mrk_prekrocenie == 0.0:
        mrk_prekrocenie = extract_euro_amount('Prekročenie rezervovanej kapacity')
    if mrk_prekrocenie == 0.0:
        mrk_prekrocenie = extract_euro_amount('Prekročenie maximálnej rezervovanej kapacity')

    jalovina = extract_euro_amount('Dodávka jalovej elektriny')
    if jalovina == 0.0:
        jalovina = extract_euro_amount('Nevyžiad.dod.kapacitnej zložky jal.en.')
    
    ucinnik = extract_euro_amount('Prirážka za nedodržanie účinníka')
    celkove_pokuty = mrk_prekrocenie + jalovina + ucinnik

    return {
        "EIC": eic_code,
        "Odberné miesto": om_info["nazov"],
        "Lokalita": om_info["lokalita"],
        "Číslo faktúry (VS)": vs_num,
        "Fakturačné obdobie": obdobie,
        "Spotreba (kWh)": silova_kwh,
        "Silová zložka (€)": silova_eur,
        "Distribúcia celkom (€)": distribucia_eur,
        "Prekročenie MRK (€)": mrk_prekrocenie,
        "Jalová elektrina (€)": jalovina,
        "Prirážka za účinník (€)": ucinnik,
        "SPOLU POKUTY (€)": celkove_pokuty,
        "Spolu bez DPH (€)": spolu_bez_dph,
        "Raw_Text": text
    }

@st.cache_data
def load_data():
    file_path = "Spotreby energií.xlsx"
    xls = pd.ExcelFile(file_path)
    months = ['Január', 'Február', 'Marec', 'Apríl', 'Máj', 'Jún', 'Júl', 'August', 'September']
    
    df_fve_raw = pd.read_excel(xls, sheet_name='Prehľad po FVE.2026')
    poprad_kwh = df_fve_raw.iloc[1, 2:11].values.astype(float)
    poprad_eur = df_fve_raw.iloc[2, 2:11].values.astype(float)
    poprad_pct = df_fve_raw.iloc[3, 2:11].values.astype(float)
    
    ba_kwh = df_fve_raw.iloc[4, 2:11].values.astype(float)
    ba_eur = df_fve_raw.iloc[5, 2:11].values.astype(float)
    ba_pct = df_fve_raw.iloc[6, 2:11].values.astype(float)
    
    fve_df = pd.DataFrame({
        'Mesiac': months * 2,
        'Lokalita_raw': ['Poprad'] * 9 + ['Bratislava'] * 9,
        'Vyroba_kWh': list(poprad_kwh) + list(ba_kwh),
        'Uspora_EUR': list(poprad_eur) + list(ba_eur),
        'Podiel_Pokrytia_Pct': list(poprad_pct) + list(ba_pct)
    })
    
    df_elec_raw = pd.read_excel(xls, sheet_name='Prehľad elektro.2026', header=None)
    records = []
    
    for idx in range(len(df_elec_raw)):
        row = df_elec_raw.iloc[idx]
        val0 = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        val1 = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
        combined_text = f"{val0} {val1}".strip()
        
        if any(keyword in combined_text.lower() for keyword in ["trafostanica", "vodomer", "odpis meračov"]):
            continue
            
        if "spotreba" in combined_text.lower():
            raw_lower = combined_text.lower()
            if "feld" in raw_lower or "bratislava" in raw_lower:
                clean_name = "Bratislava – Trafo Feld"
                lokalita_parent = "Bratislava"
            else:
                lokalita_parent = "Poprad"
                if "bananáreň" in raw_lower:
                    clean_name = "Bananáreň"
                elif "zemiakareň" in raw_lower:
                    clean_name = "Zemiakareň"
                elif "haramia" in raw_lower or "nové trafo" in raw_lower:
                    clean_name = "Trafostanica Haramia (Celkom)"
                elif "pebek" in raw_lower or "staré trafo" in raw_lower:
                    clean_name = "Trafostanica Pebek (Celkom)"
                elif "skladová hala" in raw_lower:
                    clean_name = "Skladová hala I."
                elif "ab správa" in raw_lower:
                    clean_name = "AB Správa"
                elif "beňovský" in raw_lower:
                    clean_name = "Beňovský Apartments"
                elif "nájomcovia" in raw_lower or "arál" in raw_lower:
                    clean_name = "Hlavný elektromer / Nájomcovia"
                else:
                    clean_name = combined_text.replace("spotreba", "").replace("kWh", "").replace("k kWh", "").strip().capitalize()
            
            kwh_vals = pd.to_numeric(row.values[2:11], errors='coerce')
            eur_vals = None
            for next_idx in range(idx + 1, min(idx + 4, len(df_elec_raw))):
                next_row = df_elec_raw.iloc[next_idx]
                next_text = f"{next_row.iloc[0]} {next_row.iloc[1]}".lower()
                if "náklad" in next_text or "naklad" in next_text:
                    eur_vals = pd.to_numeric(next_row.values[2:11], errors='coerce')
                    break
                    
            if eur_vals is None:
                eur_vals = [0.0] * len(months)
                
            for m, k, e in zip(months, kwh_vals, eur_vals):
                k_val = float(k) if pd.notna(k) else 0.0
                e_val = float(e) if pd.notna(e) else 0.0
                if k_val > 0 or e_val > 0:
                    records.append({
                        'Mesiac': m,
                        'Lokalita': lokalita_parent,
                        'Odberné miesto': clean_name,
                        'Spotreba (kWh)': k_val,
                        'Náklady (€)': e_val,
                        'Cena (€/kWh)': (e_val / k_val) if k_val > 0 else 0.0
                    })
                    
    elec_df = pd.DataFrame(records)
    return fve_df, elec_df

try:
    fve_df, elec_df = load_data()
except Exception as e:
    st.error(f"Chyba pri načítaní Excel súboru: {e}")
    st.stop()

# --- FILTRE V SIDEBARE ---
st.sidebar.subheader("🔍 Filtre")

selected_lokalita = st.sidebar.selectbox(
    "📍 Vyber lokalitu",
    options=["Všetky lokality", "Poprad", "Bratislava"]
)

month_order = list(fve_df['Mesiac'].unique())
selected_months = st.sidebar.multiselect(
    "📅 Vyber mesiace", 
    options=month_order, 
    default=month_order
)

filtered_fve = fve_df[fve_df['Mesiac'].isin(selected_months)].copy()
filtered_elec = elec_df[elec_df['Mesiac'].isin(selected_months)].copy()

if selected_lokalita != "Všetky lokality":
    filtered_fve = filtered_fve[filtered_fve['Lokalita_raw'] == selected_lokalita]
    filtered_elec = filtered_elec[filtered_elec['Lokalita'] == selected_lokalita]

filtered_fve['Mesiac'] = pd.Categorical(filtered_fve['Mesiac'], categories=month_order, ordered=True)
filtered_elec['Mesiac'] = pd.Categorical(filtered_elec['Mesiac'], categories=month_order, ordered=True)
filtered_fve = filtered_fve.sort_values('Mesiac')
filtered_elec = filtered_elec.sort_values('Mesiac')

total_elec_kwh = filtered_elec['Spotreba (kWh)'].sum()
total_elec_eur = filtered_elec['Náklady (€)'].sum()
avg_price_per_kwh = (total_elec_eur / total_elec_kwh) if total_elec_kwh > 0 else 0.0

total_fve_kwh = filtered_fve['Vyroba_kWh'].sum()
total_fve_eur = filtered_fve['Uspora_EUR'].sum()

st.subheader(f"📊 Kľúčové ukazovatele 2026 ({selected_lokalita})")
c1, c2, c3, c4, c5 = st.columns(5)

c1.metric("🔌 Odber zo siete (Spolu)", f"{total_elec_kwh:,.0f} kWh")
c2.metric("💶 Náklady na sieť spolu", f"{total_elec_eur:,.2f} €")
c3.metric("☀️ Výroba FVE", f"{total_fve_kwh:,.0f} kWh")
c4.metric("💰 Úspora FVE", f"{total_fve_eur:,.2f} €")
c5.metric("⚡ Priemerná cena 1 kWh", f"{avg_price_per_kwh:.4f} €/kWh")

st.divider()

tab1, tab2, tab3, tab4 = st.tabs([
    "📉 Spotreby & Náklady Odberných Miest", 
    "☀️ Výkon Fotovoltiky (FVE)", 
    "💶 Jednotková cena za 1 kWh",
    "📄 Analýza PDF Faktúr (ZSE)"
])

with tab1:
    st.subheader(f"Mesačný vývoj spotreby a nákladov ({selected_lokalita})")
    all_om = list(filtered_elec['Odberné miesto'].unique())
    selected_om = st.multiselect("Vyber odberné miesta pre zobrazenie v grafoch:", options=all_om, default=all_om, key="om_filter_tab1")
    sub_elec = filtered_elec[filtered_elec['Odberné miesto'].isin(selected_om)]
    
    if not sub_elec.empty:
        col_left, col_right = st.columns(2)
        with col_left:
            fig_elec_kwh = px.bar(sub_elec, x='Mesiac', y='Spotreba (kWh)', color='Odberné miesto', barmode='group', title="Spotreba elektriny (kWh)", template="plotly_dark")
            st.plotly_chart(fig_elec_kwh, use_container_width=True)
        with col_right:
            fig_elec_eur = px.bar(sub_elec, x='Mesiac', y='Náklady (€)', color='Odberné miesto', barmode='group', title="Mesačné náklady (€)", template="plotly_dark")
            st.plotly_chart(fig_elec_eur, use_container_width=True)
    st.dataframe(filtered_elec, use_container_width=True, hide_index=True)

with tab2:
    st.subheader(f"☀️ Výkon a úspory FVE ({selected_lokalita})")
    summary_fve = filtered_fve.groupby('Lokalita_raw').agg({'Vyroba_kWh': 'sum', 'Uspora_EUR': 'sum', 'Podiel_Pokrytia_Pct': 'mean'}).reset_index()

    pie_labels_kwh = []
    pie_labels_eur = []
    for idx, row in summary_fve.iterrows():
        loc = row['Lokalita_raw']
        pie_labels_kwh.append(f"{loc}: {row['Vyroba_kWh']:,.0f} kWh")
        pie_labels_eur.append(f"{loc}: {row['Uspora_EUR']:,.0f} €")
        
    summary_fve['Label_KWh'] = pie_labels_kwh
    summary_fve['Label_EUR'] = pie_labels_eur

    col_sum1, col_sum2 = st.columns(2)
    with col_sum1:
        fig_pie_kwh = px.pie(summary_fve, names='Label_KWh', values='Vyroba_kWh', title="Podiel lokalít na celkovej výrobe FVE", hole=0.4, template="plotly_dark")
        st.plotly_chart(fig_pie_kwh, use_container_width=True)
    with col_sum2:
        fig_pie_eur = px.pie(summary_fve, names='Label_EUR', values='Uspora_EUR', title="Podiel lokalít na finančných úsporách", hole=0.4, template="plotly_dark")
        st.plotly_chart(fig_pie_eur, use_container_width=True)

    st.dataframe(filtered_fve, use_container_width=True, hide_index=True)

with tab3:
    st.subheader(f"Porovnanie a vývoj jednotkovej ceny elektriny (€ / kWh) – {selected_lokalita}")
    if not filtered_elec.empty:
        fig_price_trend = px.line(filtered_elec, x='Mesiac', y='Cena (€/kWh)', color='Odberné miesto', markers=True, title="Vývoj ceny 1 kWh podľa odberných miest", template="plotly_dark")
        st.plotly_chart(fig_price_trend, use_container_width=True)

    records_price = []
    for om in filtered_elec['Odberné miesto'].unique():
        sub = filtered_elec[filtered_elec['Odberné miesto'] == om]
        tot_kwh = sub['Spotreba (kWh)'].sum()
        tot_eur = sub['Náklady (€)'].sum()
        avg_p = (tot_eur / tot_kwh) if tot_kwh > 0 else 0.0
        records_price.append({'Odberné miesto': om, 'Celková spotreba (kWh)': tot_kwh, 'Celkové náklady (€)': tot_eur, 'Priemerná cena (€/kWh)': avg_p})
    st.dataframe(pd.DataFrame(records_price), use_container_width=True, hide_index=True)

# ZÁLOŽKA 4: ANALÝZA PDF FAKTÚR S GRAFOVÝM SÚMAROM A SANKCIAMI
with tab4:
    st.subheader("📄 Analýza komodity, distribúcie a sankčných poplatkov zo ZSE faktúr")
    st.markdown(f"Faktúry sú načítavané z priečinka `/{INVOICE_DIR}`.")

    uploaded_files = st.file_uploader("Nahraj nové PDF faktúry:", type=["pdf"], accept_multiple_files=True)
    if uploaded_files:
        for uploaded_file in uploaded_files:
            file_path = os.path.join(INVOICE_DIR, uploaded_file.name)
            with open(file_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
        st.success(f"Úspešne uložených {len(uploaded_files)} súborov do priečinka `{INVOICE_DIR}`.")

    all_pdf_files = [os.path.join(INVOICE_DIR, f) for f in os.listdir(INVOICE_DIR) if f.lower().endswith('.pdf')]

    if all_pdf_files:
        parsed_results = []
        for pdf_path in all_pdf_files:
            try:
                data = parse_zse_pdf(pdf_path)
                parsed_results.append(data)
            except Exception as e:
                st.error(f"Chyba pri spracovaní súboru {os.path.basename(pdf_path)}: {e}")

        if parsed_results:
            df_invoices = pd.DataFrame(parsed_results)

            st.divider()
            
            # --- FILTER NA KONKRÉTNE ODBERNÉ MIESTO ---
            st.markdown("### 🔍 Filter odberných miest")
            unique_oms = ["Všetky odberné miesta"] + list(df_invoices['Odberné miesto'].unique())
            selected_single_om = st.selectbox("Vyber konkrétne odberné miesto pre detailnú analýzu:", options=unique_oms)

            if selected_single_om != "Všetky odberné miesta":
                df_filtered_inv = df_invoices[df_invoices['Odberné miesto'] == selected_single_om]
            else:
                df_filtered_inv = df_invoices

            st.success(f"Zobrazených faktúr: {len(df_filtered_inv)}")

            # ==========================================
            # SÚHRNNÉ GRAFY NA ZAČIATKU ZÁLOŽKY
            # ==========================================
            st.markdown("### 📈 Súhrnný prehľad načítaných faktúr")
            
            col_g1, col_g2 = st.columns(2)
            
            with col_g1:
                # 1. Graf: Silová zložka vs Distribúcia
                df_melt_cost = df_filtered_inv.melt(
                    id_vars=['Odberné miesto', 'Fakturačné obdobie'],
                    value_vars=['Silová zložka (€)', 'Distribúcia celkom (€)'],
                    var_name='Zložka', value_name='Suma (€)'
                )
                fig_g1 = px.bar(
                    df_melt_cost, x='Fakturačné obdobie', y='Suma (€)', color='Zložka', barmode='group',
                    title="Porovnané náklady: Silová zložka vs. Distribúcia",
                    template="plotly_dark"
                )
                st.plotly_chart(fig_g1, use_container_width=True)

            with col_g2:
                # 2. Graf: Pokuty a sankcie (MRK, Jalovina, Účinník)
                df_melt_penalties = df_filtered_inv.melt(
                    id_vars=['Odberné miesto', 'Fakturačné obdobie'],
                    value_vars=['Prekročenie MRK (€)', 'Jalová elektrina (€)', 'Prirážka za účinník (€)'],
                    var_name='Typ sankcie', value_name='Suma (€)'
                )
                fig_g2 = px.bar(
                    df_melt_penalties, x='Fakturačné obdobie', y='Suma (€)', color='Typ sankcie', barmode='group',
                    title="Súhrn pokút a sankcií (MRK, Jalovina, Účinník)",
                    template="plotly_dark"
                )
                st.plotly_chart(fig_g2, use_container_width=True)

            st.divider()
            
            display_cols = [c for c in df_filtered_inv.columns if c != "Raw_Text"]
            st.dataframe(df_filtered_inv[display_cols], use_container_width=True, hide_index=True)

            st.divider()
            
            for idx, row in df_filtered_inv.iterrows():
                st.markdown(f"#### Faktúra č. {row['Číslo faktúry (VS)']} – {row['Odberné miesto']} ({row['Fakturačné obdobie']})")
                
                c_f1, c_f2, c_f3, c_f4 = st.columns(4)
                c_f1.metric("Silová zložka (Komodita)", f"{row['Silová zložka (€)']:,.2f} €")
                c_f2.metric("Distribúcia celkom", f"{row['Distribúcia celkom (€)']:,.2f} €")
                c_f3.metric("Prekročenie MRK", f"{row['Prekročenie MRK (€)']:,.2f} €", delta_color="inverse" if row['Prekročenie MRK (€)'] > 0 else "normal")
                c_f4.metric("Jalová + Účinník", f"{(row['Jalová elektrina (€)'] + row['Prirážka za účinník (€)']):,.2f} €", delta_color="inverse" if (row['Jalová elektrina (€)'] + row['Prirážka za účinník (€)']) > 0 else "normal")

                # Koláčový graf pre komoditu, distribúciu a pokuty
                cost_breakdown = {
                    "Silová zložka (Komodita)": row['Silová zložka (€)'],
                    "Distribúcia celkom": row['Distribúcia celkom (€)'],
                    "Prekročenie MRK": row['Prekročenie MRK (€)'],
                    "Jalová elektrina": row['Jalová elektrina (€)'],
                    "Prirážka za účinník": row['Prirážka za účinník (€)']
                }
                
                df_break = pd.DataFrame([{"Položka": k, "Suma (€)": v} for k, v in cost_breakdown.items() if v > 0])

                if not df_break.empty:
                    fig_break = px.pie(
                        df_break, names='Položka', values='Suma (€)',
                        title=f"Podiel nákladov a sankcií ({row['Odberné miesto']})",
                        hole=0.4, template="plotly_dark"
                    )
                    fig_break.update_traces(textinfo='percent+label', textfont=dict(size=14, color='white'))
                    st.plotly_chart(fig_break, use_container_width=True)

                with st.expander(f"🔍 Diagnostika surového textu z PDF pre faktúru {row['Číslo faktúry (VS)']}"):
                    st.text(row['Raw_Text'])

                st.divider()
    else:
        st.info(f"V priečinku `{INVOICE_DIR}/` zatiaľ nie sú žiadne PDF faktúry.")
