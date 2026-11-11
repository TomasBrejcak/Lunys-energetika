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
    "24ZZS4000182286S": {"nazov": "Bratislava – Trafo Feld", "lokalita": "Bratislava"},
    "24ZZS40001822865": {"nazov": "Bratislava – Trafo Feld", "lokalita": "Bratislava"},
    "24ZVS0000016878H": {"nazov": "Beňovský Apartments", "lokalita": "Poprad"},
    "24ZVS0000059036Q": {"nazov": "Hlavný elektromer / Nájomcovia", "lokalita": "Poprad"},
    "24ZVS00000590360": {"nazov": "Hlavný elektromer / Nájomcovia", "lokalita": "Poprad"},
    "24ZVS0000787041D": {"nazov": "Trafostanica HARAMIA", "lokalita": "Poprad"}
}

# --- OPTIMALIZOVANÝ A CACHOVANÝ PARSER ZSE FAKTÚR ---
@st.cache_data
def parse_zse_pdf(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        text = "\n".join([page.extract_text() or "" for page in pdf.pages])

    # 1. Detekcia EIC kódu
    eic_match = re.search(r'(24Z[A-Z0-9]{13})', text)
    eic_code = eic_match.group(1) if eic_match else "Neznáme EIC"
    om_info = EIC_MAP.get(eic_code, {"nazov": f"Neznáme odberné miesto ({eic_code})", "lokalita": "Neznáma"})

    # 2. Detekcia čísla faktúry a obdobia
    vs_match = re.search(r'(?:Vyúčtovacia faktúra za elektrinu č\.|Variabilný symbol:)\s*(\d+)', text)
    vs_num = vs_match.group(1) if vs_match else "Neuvedené"

    obdobie_match = re.search(r'Fakturačné obdobie\s*(\d{2}\.\d{2}\.\d{4})\s*-\s*(\d{2}\.\d{2}\.\d{4})', text)
    if not obdobie_match:
        obdobie_match = re.search(r'(\d{2}\.\d{2}\.\d{4})\s*-\s*(\d{2}\.\d{2}\.\d{4})', text)

    obdobie = f"{obdobie_match.group(1)} - {obdobie_match.group(2)}" if obdobie_match else "Neuvedené"
    
    sort_date = obdobie_match.group(1).split('.') if obdobie_match else []
    sort_date_str = f"{sort_date[2]}-{sort_date[1]}-{sort_date[0]}" if len(sort_date) == 3 else "0000-00-00"
    
    month_map = {"01": "Január", "02": "Február", "03": "Marec", "04": "Apríl", "05": "Máj", "06": "Jún", "07": "Júl", "08": "August", "09": "September", "10": "Október", "11": "November", "12": "December"}
    invoice_month = month_map.get(sort_date[1], "Neuvedené") if len(sort_date) == 3 else "Neuvedené"

    def get_exact_line_cost(keywords):
        if isinstance(keywords, str):
            keywords = [keywords]
        for line in text.split('\n'):
            if any(kw.lower() in line.lower() for kw in keywords):
                if any(x in line.lower() for x in ['technické údaje', 'stav']):
                    continue
                amounts = re.findall(r'(\d{1,3}(?:\s?\d{3})*,\d{2})', line)
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
            for line in text.split('\n'):
                if kw.lower() in line.lower():
                    if any(x in line.lower() for x in ['technické údaje', 'stav']):
                        continue
                    amounts = re.findall(r'(\d{1,3}(?:\s?\d{3})*,\d{2})', line)
                    if amounts:
                        val_str = amounts[-1].replace(" ", "").replace(",", ".")
                        try:
                            val = float(val_str)
                            if 0 <= val < 1000000:
                                return val
                        except:
                            pass
            matches = [m.start() for m in re.finditer(re.escape(kw), text, re.IGNORECASE)]
            for pos in matches:
                window = text[pos:pos+150]
                if any(x in window.lower() for x in ['technické údaje', 'stav']):
                    continue
                amounts = re.findall(r'(\d{1,3}(?:\s?\d{3})*,\d{2})', window)
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
        for line in text.split('\n'):
            if any(kw.lower() in line.lower() for kw in keyword):
                if 'technické údaje' in line.lower():
                    continue
                m = re.search(r'([\d\s,]+)\s*kWh', line, re.IGNORECASE)
                if m:
                    val_str = m.group(1).split()[0].replace(" ", "").replace(",", ".")
                    try:
                        val = float(val_str)
                        if val < 10000000:
                            return val
                    except:
                        pass
        return 0.0

    silova_kwh = get_line_kwh(['1T za elektrinu', 'ČE 1T', 'Dodávka elektriny'])
    silova_eur = get_exact_line_cost(['Dodávka elektriny'])
    if silova_eur == 0.0:
        silova_eur = get_exact_line_cost(['1T za elektrinu', 'Register-Special', 'ČE 1T'])

    distribucia_eur = 0.0
    sumar_section = re.search(r'(?:Sumár za miesto spotreby|Sumár za faktúru).*?(?=Suma spolu|$)', text, re.IGNORECASE | re.DOTALL)
    if sumar_section:
        distrib_match = re.search(r'Distribúcia elektriny[^\d]*([\d\s]+,\d{2})\s*€', sumar_section.group(0), re.IGNORECASE)
        if distrib_match:
            try:
                distribucia_eur = float(distrib_match.group(1).replace(" ", "").replace(",", "."))
            except:
                pass

    if distribucia_eur == 0.0:
        distrib_section = re.search(r'Služba:\s*Distribúcia.*?(Spolu bez DPH.*?€)', text, re.IGNORECASE | re.DOTALL)
        if distrib_section:
            match_sum = re.search(r'Spolu bez DPH\s*([\d\s]+,\d{2})\s*€', distrib_section.group(1), re.IGNORECASE)
            if match_sum:
                try:
                    distribucia_eur = float(match_sum.group(1).replace(" ", "").replace(",", "."))
                except:
                    pass

    rk_prekrocenie = get_penalty_cost([
        'Prekročenie rezervovanej kapacity', 
        'Prekročenie RK',
        'Tarifa za prekročenie rezervovanej kapacity'
    ])
    
    mrk_prekrocenie = get_penalty_cost([
        'Prekročenie maximálnej rezervovanej kapacity', 
        'Tarifa za prekročenie MRK',
        'Prekročenie MRK'
    ])
    
    jalovina = get_penalty_cost([
        'Dodávka jalovej elektriny', 
        'Nevyžiad.dod.kapacitnej zložky jal.en.', 
        'Jalová elektrina',
        'Dodávka jalovej induktívnej',
        'Dodávka jalovej kapacitnej'
    ])
    
    ucinnik = get_penalty_cost([
        'Prirážka za nedodržanie účinníka', 
        'Nedodržanie účinníka',
        'Prirážka za účinník',
        'Účinník'
    ])
    
    celkove_pokuty = rk_prekrocenie + mrk_prekrocenie + jalovina + ucinnik

    spolu_bez_dph = 0.0
    rekap_match = re.search(r'(?:Základ dane celkom|Spolu bez DPH celkom|Celkom bez DPH).*?(\d{1,3}(?:\s?\d{3})*,\d{2})', text, re.IGNORECASE)
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
        "Distribúcia celkom (€)": distribucia_eur,
        "Prekročenie RK (€)": rk_prekrocenie,
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
                    clean_name = "Trafostanica HARAMIA"
                elif "pebek" in raw_lower or "staré trafo" in raw_lower:
                    clean_name = "Trafostanica PEBEK"
                elif "skladová hala" in raw_lower:
                    clean_name = "Skladová hala I."
                elif "ab správa" in raw_lower:
                    clean_name = "AB Správa"
                elif "beňovský" in raw_lower:
                    clean_name = "Beňovský Apartments"
                elif "nájomcovia" in raw_lower or "arál" in raw_lower:
                    clean_name = "Areál / Nájomcovia (Poprad)"
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

# --- INITIALIZÁCIA STAVU ZÁLOŽIEK ---
TAB_NAMES = [
    "📉 Spotreby & Náklady Odberných Miest", 
    "☀️ Výkon Fotovoltiky (FVE)", 
    "💶 Jednotková cena za 1 kWh",
    "📄 Analýza PDF Faktúr & Audit"
]

if 'active_tab' not in st.session_state:
    st.session_state.active_tab = TAB_NAMES[0]

# --- FUNKCIA PRE ZMENU STAVU CEZ SIDEBAR ---
def set_tab(tab_name):
    st.session_state.active_tab = tab_name

# --- FILTRE & OVLÁDANIE V SIDEBARE ---
st.sidebar.subheader("🔍 Filtre & Ovládanie")

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

st.sidebar.divider()

st.sidebar.subheader("⚡ Rýchla navigácia")
if st.sidebar.button("📉 Spotreby & Náklady", use_container_width=True, on_click=set_tab, args=(TAB_NAMES[0],)):
    pass
if st.sidebar.button("☀️ Výkon FVE", use_container_width=True, on_click=set_tab, args=(TAB_NAMES[1],)):
    pass
if st.sidebar.button("💶 Jednotková cena kWh", use_container_width=True, on_click=set_tab, args=(TAB_NAMES[2],)):
    pass
if st.sidebar.button("📄 Analýza PDF Faktúr", use_container_width=True, on_click=set_tab, args=(TAB_NAMES[3],)):
    pass

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

# SYNCHRONIZOVANÉ PREPÍNANIE SEKCÍ
selected_tab = st.radio("Zvoľ sekciu dashboardu:", TAB_NAMES, key='active_tab', horizontal=True, label_visibility="collapsed")
st.divider()

if selected_tab == TAB_NAMES[0]:
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

elif selected_tab == TAB_NAMES[1]:
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

elif selected_tab == TAB_NAMES[2]:
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

elif selected_tab == TAB_NAMES[3]:
    st.subheader("📄 Analýza PDF faktúr zo ZSE")
    
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
            df_invoices = df_invoices.drop_duplicates(subset=['EIC', 'Číslo faktúry (VS)'], keep='first')

            # 1. FILTER ODBERNÝCH MIEST NA ÚPLNOM VRCHU
            st.markdown("### 🔍 Filter odberných miest pre analýzu")
            unique_oms = ["Všetky odberné miesta"] + list(df_invoices['Odberné miesto'].unique())
            selected_single_om = st.selectbox("Vyber konkrétne odberné miesto:", options=unique_oms)

            if selected_single_om != "Všetky odberné miesta":
                df_filtered_inv = df_invoices[df_invoices['Odberné miesto'] == selected_single_om].copy()
            else:
                df_filtered_inv = df_invoices.copy()

            df_filtered_inv = df_filtered_inv.sort_values(by='_sort_date', ascending=True).reset_index(drop=True)
            st.success(f"Zobrazených jedinečných faktúr: {len(df_filtered_inv)}")

            st.divider()

            # 2. KOLÁČOVÉ GRAFY POD FILTROM (ČISTÉ POPISKY + TUČNÉ SUMY)
            st.markdown("### 📈 Súhrnný prehľad nákladov a sankcií")
            col_g1, col_g2 = st.columns(2)
            
            with col_g1:
                tot_silova = df_filtered_inv['Silová zložka (€)'].sum()
                tot_distrib = df_filtered_inv['Distribúcia celkom (€)'].sum()
                sum_cost_total = tot_silova + tot_distrib
                
                df_pie_cost = pd.DataFrame({
                    'Zložka': [f"Silová zložka ({tot_silova:,.2f} €)", f"Distribúcia celkom ({tot_distrib:,.2f} €)"],
                    'Suma': [tot_silova, tot_distrib]
                })
                
                fig_pie1 = px.pie(
                    df_pie_cost, names='Zložka', values='Suma', hole=0.4,
                    title="Podiel nákladov: Silová zložka vs. Distribúcia",
                    template="plotly_dark"
                )
                fig_pie1.update_traces(textinfo='percent+label', textfont=dict(size=13, color='white'))
                st.plotly_chart(fig_pie1, use_container_width=True)
                st.markdown(f"**Spolu náklady: {sum_cost_total:,.2f} €**")

            with col_g2:
                tot_rk = df_filtered_inv['Prekročenie RK (€)'].sum()
                tot_mrk = df_filtered_inv['Prekročenie MRK (€)'].sum()
                tot_jalova = df_filtered_inv['Jalová elektrina (€)'].sum()
                tot_ucinnik = df_filtered_inv['Prirážka za účinník (€)'].sum()
                sum_penalties_total = tot_rk + tot_mrk + tot_jalova + tot_ucinnik
                
                df_pie_penalties = pd.DataFrame({
                    'Typ sankcie': [
                        f"Prekročenie RK ({tot_rk:,.2f} €)",
                        f"Prekročenie MRK ({tot_mrk:,.2f} €)", 
                        f"Jalová elektrina ({tot_jalova:,.2f} €)", 
                        f"Prirážka za účinník ({tot_ucinnik:,.2f} €)"
                    ],
                    'Suma': [tot_rk, tot_mrk, tot_jalova, tot_ucinnik]
                })
                
                fig_pie2 = px.pie(
                    df_pie_penalties, names='Typ sankcie', values='Suma', hole=0.4,
                    title="Súhrn pokút a sankcií",
                    template="plotly_dark"
                )
                fig_pie2.update_traces(textinfo='percent', textfont=dict(size=13, color='white'))
                st.plotly_chart(fig_pie2, use_container_width=True)
                st.markdown(f"**Spolu pokuty a sankcie: {sum_penalties_total:,.2f} €**")

            st.divider()

            # 3. TABUĽKA S DÁTAMI Z FAKTÚR
            st.markdown("### 📋 Prehľadná tabuľka načítaných faktúr")
            display_cols = [c for c in df_filtered_inv.columns if c not in ["Raw_Text", "_sort_date", "Mesiac"]]
            st.dataframe(df_filtered_inv[display_cols], use_container_width=True, hide_index=True)

            st.divider()

            # 4. MOŽNOSŤ NAHRAŤ ĎALŠIE FAKTÚRY
            st.markdown("### 📤 Nahranie ďalších PDF faktúr")
            uploaded_files = st.file_uploader("Pridaj ďalšie PDF faktúry:", type=["pdf"], accept_multiple_files=True, key="uploader_bottom")
            if uploaded_files:
                for uploaded_file in uploaded_files:
                    file_path = os.path.join(INVOICE_DIR, uploaded_file.name)
                    with open(file_path, "wb") as f:
                        f.write(uploaded_file.getbuffer())
                st.success(f"Úspešne uložených {len(uploaded_files)} súborov do priečinka `{INVOICE_DIR}`. Obnov stránku pre aktualizáciu.")

            st.divider()

            # 5. DETAIL PO MESIACOCH (KARTY FAKTÚR)
            st.markdown("### 🔎 Detail faktúr podľa mesiacov")
            for idx, row in df_filtered_inv.iterrows():
                st.markdown(f"#### Faktúra č. {row['Číslo faktúry (VS)']} – {row['Odberné miesto']} ({row['Fakturačné obdobie']})")
                
                c_f1, c_f2, c_f3, c_f4 = st.columns(4)
                c_f1.metric("Silová zložka (Komodita)", f"{row['Silová zložka (€)']:,.2f} €")
                c_f2.metric("Distribúcia celkom", f"{row['Distribúcia celkom (€)']:,.2f} €")
                c_f3.metric("Prekročenie RK / MRK", f"{(row['Prekročenie RK (€)'] + row['Prekročenie MRK (€)']):,.2f} €", delta_color="inverse" if (row['Prekročenie RK (€)'] + row['Prekročenie MRK (€)']) > 0 else "normal")
                c_f4.metric("Jalová + Účinník", f"{(row['Jalová elektrina (€)'] + row['Prirážka za účinník (€)']):,.2f} €", delta_color="inverse" if (row['Jalová elektrina (€)'] + row['Prirážka za účinník (€)']) > 0 else "normal")

                cost_breakdown = {
                    "Silová zložka (Komodita)": row['Silová zložka (€)'],
                    "Distribúcia celkom": row['Distribúcia celkom (€)'],
                    "Prekročenie RK": row['Prekročenie RK (€)'],
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

            # 6. TABUĽKA KRÍŽOVEJ KONTROLY (PDF vs EXCEL) NA ÚPLNOM KONCI
            st.markdown("### 🔍 Krížová kontrola (PDF vs. Excel)")
            st.markdown("Porovnanie nákladov bez DPH vyťažených z nahratých PDF faktúr oproti hodnotám zapísaným v Exceli (`Spotreby energií.xlsx`).")

            audit_records = []
            for _, inv in df_invoices.iterrows():
                om = inv['Odberné miesto']
                m = inv['Mesiac']
                pdf_val = inv['Spolu bez DPH (€)']

                match_row = elec_df[(elec_df['Odberné miesto'] == om) & (elec_df['Mesiac'] == m)]
                
                if not match_row.empty:
                    excel_val = match_row['Náklady (€)'].values[0]
                    diff = pdf_val - excel_val
                    status = "✅ V poriadku" if abs(diff) < 1.0 else "⚠️ Rozdiel"
                else:
                    excel_val = 0.0
                    diff = pdf_val
                    status = "❓ Chýba v Exceli"

                audit_records.append({
                    "Odberné miesto": om,
                    "Mesiac": m,
                    "Číslo faktúry (VS)": inv['Číslo faktúry (VS)'],
                    "PDF Spolu bez DPH (€)": pdf_val,
                    "Excel Náklady (€)": excel_val,
                    "Rozdiel (€)": diff,
                    "Stav kontroly": status
                })

            df_audit = pd.DataFrame(audit_records)
            st.dataframe(df_audit, use_container_width=True, hide_index=True)
    else:
        st.info(f"V priečinku `{INVOICE_DIR}/` zatiaľ nie sú žiadne PDF faktúry.")
