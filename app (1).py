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
st.markdown("Interaktívny nástroj pre vyhodnotenie reálnych spotrieb, nákladov a úspor z FVE priamo zo ZSE faktúr.")

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
    if not os.path.exists(pdf_path):
        return None
        
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
def load_fve_data():
    file_path = "Spotreby energií.xlsx"
    if not os.path.exists(file_path):
        return pd.DataFrame()
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
    return fve_df

fve_df = load_fve_data()

# Načítanie reálnych dát z nahratých PDF faktúr pre elektrinu
all_pdf_files = [f for f in os.listdir(INVOICE_DIR) if f.lower().endswith('.pdf')]
parsed_invoice_records = []
for f in all_pdf_files:
    f_path = os.path.join(INVOICE_DIR, f)
    res = parse_zse_pdf(f_path)
    if res:
        parsed_invoice_records.append({
            'Mesiac': res['Mesiac'],
            'Lokalita': res['Lokalita'],
            'Odberné miesto': res['Odberné miesto'],
            'Spotreba (kWh)': res['Spotreba (kWh)'],
            'Náklady (€)': res['Spolu bez DPH (€)'],
            'Cena (€/kWh)': (res['Spolu bez DPH (€)'] / res['Spotreba (kWh)']) if res['Spotreba (kWh)'] > 0 else 0.0,
            'Číslo faktúry (VS)': res['Číslo faktúry (VS)'],
            'Fakturačné obdobie': res['Fakturačné obdobie'],
            '_sort_date': res['_sort_date'],
            'Raw_Text': res['Raw_Text'],
            **res
        })

elec_df = pd.DataFrame(parsed_invoice_records)

# --- INITIALIZÁCIA STAVU ZÁLOŽIEK A FILTER STAVU ---
TAB_NAMES = [
    "📉 Spotreby & Náklady Odberných Miest", 
    "☀️ Výkon Fotovoltiky (FVE)", 
    "💶 Spotová cena za 1 kWh (Faktúry)",
    "📄 Analýza PDF Faktúr & Audit"
]

if 'active_tab' not in st.session_state:
    st.session_state.active_tab = TAB_NAMES[0]

def set_tab(tab_name):
    st.session_state.active_tab = tab_name

# Zoznam všetkých dostupných odberných miest pre Excel-like výber
all_possible_oms = sorted(elec_df['Odberné miesto'].unique()) if not elec_df.empty else []

# Inicializácia session_state pre výber odberných miest, aby si pamätal stav
if 'selected_oms' not in st.session_state:
    st.session_state.selected_oms = all_possible_oms

# --- FILTRE & OVLÁDANIE V SIDEBARE ---
st.sidebar.subheader("🔍 Globálne filtre")

selected_lokalita = st.sidebar.selectbox(
    "📍 Vyber lokalitu",
    options=["Všetky lokality", "Poprad", "Bratislava"]
)

month_order = ['Január', 'Február', 'Marec', 'Apríl', 'Máj', 'Jún', 'Júl', 'August', 'September', 'Október', 'November', 'December']
available_months = [m for m in month_order if not elec_df.empty and m in elec_df['Mesiac'].unique()]
if not available_months:
    available_months = month_order[:9]

selected_months = st.sidebar.multiselect(
    "📅 Vyber mesiace", 
    options=month_order, 
    default=available_months
)

st.sidebar.divider()

# --- EXCEL-LIKE VÝBER ODBERNÝCH MIEST V SIDEBARE ---
st.sidebar.subheader("🏢 Odberné miesta (Filter)")
if all_possible_oms:
    # Tlačidlá na rýchle označenie / odznakečenie všetkých
    col_s1, col_s2 = st.sidebar.columns(2)
    if col_s1.button("Vybrať všetko"):
        st.session_state.selected_oms = all_possible_oms
    if col_s2.button("Zrušiť všetko"):
        st.session_state.selected_oms = []

    # Dynamické checkboxy pre každé odberné miesto (štýl Excel filtra)
    updated_oms = []
    for om in all_possible_oms:
        is_checked = st.sidebar.checkbox(om, value=om in st.session_state.selected_oms, key=f"cb_om_{om}")
        if is_checked:
            updated_oms.append(om)
    st.session_state.selected_oms = updated_oms
else:
    st.sidebar.info("Žiadne dáta zatiaľ nie sú dostupné.")

st.sidebar.divider()

st.sidebar.subheader("⚡ Rýchla navigácia")
if st.sidebar.button("📉 Spotreby & Náklady", use_container_width=True, on_click=set_tab, args=(TAB_NAMES[0],)):
    pass
if st.sidebar.button("☀️ Výkon FVE", use_container_width=True, on_click=set_tab, args=(TAB_NAMES[1],)):
    pass
if st.sidebar.button("💶 Spotová cena kWh", use_container_width=True, on_click=set_tab, args=(TAB_NAMES[2],)):
    pass
if st.sidebar.button("📄 Analýza PDF Faktúr", use_container_width=True, on_click=set_tab, args=(TAB_NAMES[3],)):
    pass

# Filtrovanie dát podľa zvoleného stavu
filtered_fve = fve_df[fve_df['Mesiac'].isin(selected_months)].copy() if not fve_df.empty else pd.DataFrame()
filtered_elec = elec_df[elec_df['Mesiac'].isin(selected_months)].copy() if not elec_df.empty else pd.DataFrame()

if selected_lokalita != "Všetky lokality":
    if not filtered_fve.empty:
        filtered_fve = filtered_fve[filtered_fve['Lokalita_raw'] == selected_lokalita]
    if not filtered_elec.empty:
        filtered_elec = filtered_elec[filtered_elec['Lokalita'] == selected_lokalita]

if not filtered_elec.empty and st.session_state.selected_oms:
    filtered_elec = filtered_elec[filtered_elec['Odberné miesto'].isin(st.session_state.selected_oms)]

if not filtered_fve.empty:
    filtered_fve['Mesiac'] = pd.Categorical(filtered_fve['Mesiac'], categories=month_order, ordered=True)
    filtered_fve = filtered_fve.sort_values('Mesiac')

if not filtered_elec.empty:
    filtered_elec['Mesiac'] = pd.Categorical(filtered_elec['Mesiac'], categories=month_order, ordered=True)
    filtered_elec = filtered_elec.sort_values('Mesiac')

total_elec_kwh = filtered_elec['Spotreba (kWh)'].sum() if not filtered_elec.empty else 0.0
total_elec_eur = filtered_elec['Náklady (€)'].sum() if not filtered_elec.empty else 0.0
avg_price_per_kwh = (total_elec_eur / total_elec_kwh) if total_elec_kwh > 0 else 0.0

total_fve_kwh = filtered_fve['Vyroba_kWh'].sum() if not filtered_fve.empty else 0.0
total_fve_eur = filtered_fve['Uspora_EUR'].sum() if not filtered_fve.empty else 0.0

st.subheader(f"📊 Kľúčové ukazovatele 2026 ({selected_lokalita}) – Z faktúr")
c1, c2, c3, c4, c5 = st.columns(5)

c1.metric("🔌 Odber zo siete (Spolu)", f"{total_elec_kwh:,.0f} kWh")
c2.metric("💶 Náklady na sieť spolu", f"{total_elec_eur:,.2f} €")
c3.metric("☀️ Výroba FVE", f"{total_fve_kwh:,.0f} kWh")
c4.metric("💰 Úspora FVE", f"{total_fve_eur:,.2f} €")
c5.metric("⚡ Priemerná cena 1 kWh", f"{avg_price_per_kwh:.4f} €/kWh")

st.divider()

selected_tab = st.radio("Zvoľ sekciu dashboardu:", TAB_NAMES, key='active_tab', horizontal=True, label_visibility="collapsed")
st.divider()

if selected_tab == TAB_NAMES[0]:
    st.subheader(f"Mesačný vývoj spotreby a nákladov z faktúr ({selected_lokalita})")
    if not filtered_elec.empty:
        col_left, col_right = st.columns(2)
        with col_left:
            fig_elec_kwh = px.bar(filtered_elec, x='Mesiac', y='Spotreba (kWh)', color='Odberné miesto', barmode='group', title="Spotreba elektriny (kWh)", template="plotly_dark")
            st.plotly_chart(fig_elec_kwh, use_container_width=True)
        with col_right:
            fig_elec_eur = px.bar(filtered_elec, x='Mesiac', y='Náklady (€)', color='Odberné miesto', barmode='group', title="Mesačné náklady bez DPH (€)", template="plotly_dark")
            st.plotly_chart(fig_elec_eur, use_container_width=True)
        st.dataframe(filtered_elec[['Mesiac', 'Odberné miesto', 'Lokalita', 'Spotreba (kWh)', 'Náklady (€)', 'Cena (€/kWh)', 'Fakturačné obdobie']], use_container_width=True, hide_index=True)
    else:
        st.info("Zatiaľ nie sú nahraté žiadne PDF faktúry alebo pre zvolené filtre neexistujú dáta.")

elif selected_tab == TAB_NAMES[1]:
    st.subheader(f"☀️ Výkon a úspory FVE ({selected_lokalita})")
    if not filtered_fve.empty:
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
    else:
        st.info("FVE dáta nie sú k dispozícii.")

elif selected_tab == TAB_NAMES[2]:
    st.subheader(f"💶 Reálna spotová cena za 1 kWh z PDF faktúr ({selected_lokalita})")
    st.markdown("Porovnanie čistej **silovej zložky (spotová cena komodity)** vs. **celkovej reálnej ceny vrátane distribúcie, poplatkov a pokút** vypočítanej priamo z nahratých faktúr.")

    if not filtered_elec.empty:
        df_inv_prices = filtered_elec.copy()
        df_inv_prices = df_inv_prices.drop_duplicates(subset=['EIC', 'Číslo faktúry (VS)'], keep='first')

        if not df_inv_prices.empty:
            df_inv_prices['Cena sila (€/kWh)'] = df_inv_prices.apply(lambda r: (r['Silová zložka (€)'] / r['Spotreba (kWh)']) if r['Spotreba (kWh)'] > 0 else 0, axis=1)
            df_inv_prices['Cena celkom s distribúciou (€/kWh)'] = df_inv_prices.apply(lambda r: (r['Spolu bez DPH (€)'] / r['Spotreba (kWh)']) if r['Spotreba (kWh)'] > 0 else 0, axis=1)

            fig_spot_trend = go.Figure()
            if '_sort_date' in df_inv_prices.columns:
                df_inv_prices = df_inv_prices.sort_values('_sort_date')

            for om in df_inv_prices['Odberné miesto'].unique():
                sub_om = df_inv_prices[df_inv_prices['Odberné miesto'] == om]
                
                fig_spot_trend.add_trace(go.Scatter(
                    x=sub_om['Fakturačné obdobie'],
                    y=sub_om['Cena sila (€/kWh)'],
                    mode='lines+markers',
                    name=f"{om} – Sila (€/kWh)"
                ))
                
                fig_spot_trend.add_trace(go.Scatter(
                    x=sub_om['Fakturačné obdobie'],
                    y=sub_om['Cena celkom s distribúciou (€/kWh)'],
                    mode='lines+markers',
                    name=f"{om} – Celkom s distribúciou",
                    line=dict(dash='dash')
                ))

            fig_spot_trend.update_layout(
                title="Mesačný vývoj spotovej ceny komodity vs. reálnej ceny s distribúciou (€/kWh)",
                xaxis_title="Fakturačné obdobie",
                yaxis_title="Cena (€/kWh)",
                template="plotly_dark",
                hovermode="x unified"
            )
            st.plotly_chart(fig_spot_trend, use_container_width=True)

            st.markdown("### 📋 Súhrnné reálne ceny za 1 kWh podľa odberných miest")
            summary_spot = []
            for om in df_inv_prices['Odberné miesto'].unique():
                sub = df_inv_prices[df_inv_prices['Odberné miesto'] == om]
                tot_kwh = sub['Spotreba (kWh)'].sum()
                tot_silova = sub['Silová zložka (€)'].sum()
                tot_spolu = sub['Spolu bez DPH (€)'].sum()
                
                avg_sila = (tot_silova / tot_kwh) if tot_kwh > 0 else 0
                avg_celkom = (tot_spolu / tot_kwh) if tot_kwh > 0 else 0
                
                summary_spot.append({
                    'Odberné miesto': om,
                    'Celková spotreba (kWh)': tot_kwh,
                    'Priemerná cena sila (spot) (€/kWh)': round(avg_sila, 4),
                    'Reálna celková cena s distribúciou (€/kWh)': round(avg_celkom, 4)
                })
            
            st.dataframe(pd.DataFrame(summary_spot), use_container_width=True, hide_index=True)
        else:
            st.info("Po aplikovaní filtrov nezostali žiadne dáta na zobrazenie.")
    else:
        st.info("V priečinku `faktury_zse/` nie sú nahraté žiadne PDF faktúry alebo pre zvolené filtre neexistujú dáta.")

elif selected_tab == TAB_NAMES[3]:
    st.subheader("📄 Analýza PDF faktúr zo ZSE")
    
    if not filtered_elec.empty:
        df_filtered_inv = filtered_elec.copy()
        df_filtered_inv = df_filtered_inv.drop_duplicates(subset=['EIC', 'Číslo faktúry (VS)'], keep='first')
        df_filtered_inv = df_filtered_inv.sort_values(by='_sort_date', ascending=True).reset_index(drop=True)
        
        st.success(f"Zobrazených jedinečných faktúr: {len(df_filtered_inv)}")
        st.divider()

        st.markdown("### 📈 Súhrnný prehľad nákladov a sankcií z faktúr")
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

        st.markdown("### 📋 Prehľadná tabuľka načítaných faktúr")
        display_cols = [c for c in df_filtered_inv.columns if c not in ["Raw_Text", "_sort_date", "Mesiac", "Raw_Data"]]
        st.dataframe(df_filtered_inv[display_cols], use_container_width=True, hide_index=True)

        st.divider()

        st.markdown("### 📤 Nahranie ďalších PDF faktúr")
        uploaded_files = st.file_uploader("Pridaj ďalšie PDF faktúry:", type=["pdf"], accept_multiple_files=True, key="uploader_bottom")
        if uploaded_files:
            for uploaded_file in uploaded_files:
                file_path = os.path.join(INVOICE_DIR, uploaded_file.name)
                with open(file_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
            st.success(f"Úspešne uložených {len(uploaded_files)} súborov do priečinka `{INVOICE_DIR}`. Obnov stránku pre aktualizáciu.")

        st.divider()

        st.markdown("### 🔎 Detail faktúr podľa mesiacov")
        for idx, row in df_filtered_inv.iterrows():
            st.markdown(f"#### Faktúra č. {row['Číslo faktúry (VS)']} – {row['Odberné miesto']} ({row['Fakturačné obdobie']})")
            
            c_f1, c_f2, c_f3, c_f4 = st.columns(4)
            c_f1.metric("Silová zložka (Komodita)", f"{row['Silová zložka (€)']:,.2f} €")
            c_f2.metric("Distribúcia celkom", f"{row['Distribúcia celkom (€)']:,.2f} €")
            c_f3.metric("Prekročenie RK / MRK", f"{(row['Prekročenie RK (€)'] + row['Prekročenie MRK (€)']):,.2f} €", delta_color="inverse" if (row['Prekročenie RK (€)'] + row['Prekročenie MRK (€)']) > 0 else "normal")
            c_f4.metric("Jalová + Účinník", f"{(row['Jalová elektrina (€)'] + row['Prirážka za účinník (€)']):,.2f} €", delta_color="inverse" if (row['Jalová elektrina (€)'] + row['Prirážka za účinník (€)']) > 0 else "normal")

            with st.expander(f"🔍 Diagnostika surového textu z PDF pre faktúry {row['Číslo faktúry (VS)']}"):
                st.text(row['Raw_Text'])

            st.divider()
    else:
        st.info(f"V priečinku `{INVOICE_DIR}/` zatiaľ nie sú žiadne PDF faktúry alebo pre zvolené filtre neexistujú dáta.")
