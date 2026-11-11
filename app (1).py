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
    "24ZVS0000016878H": {"nazov": "Beňovský Apartments", "lokalita": "Poprad"},
    "24ZVS0000059036Q": {"nazov": "Areál / Nájomcovia (Poprad)", "lokalita": "Poprad"},
    "24ZVS0000787041D": {"nazov": "Trafostanica HARAMIA", "lokalita": "Poprad"}
}

# --- FUNKCIA PRE AUTOMATICKÉ PARSOVANIE ZSE PDF FAKTÚRY ---
def parse_zse_pdf(pdf_file):
    with pdfplumber.open(pdf_file) as pdf:
        text = "\n".join([page.extract_text() or "" for page in pdf.pages])

    # 1. Detekcia EIC kódu
    eic_match = re.search(r'24Z[A-Z0-9]{13}', text)
    eic_code = eic_match.group(0) if eic_match else "Neznáme EIC"
    om_info = EIC_MAP.get(eic_code, {"nazov": f"Neznáme odberné miesto ({eic_code})", "lokalita": "Neznáma"})

    # 2. Detekcia čísla faktúry a obdobia
    vs_match = re.search(r'Vyúčtovacia faktúra za elektrinu č\.\s*(\d+)', text)
    vs_num = vs_match.group(1) if vs_match else "Neuvedené"

    obdobie_match = re.search(r'Fakturačné obdobie\s*(\d{2}\.\d{2}\.\d{4})\s*-\s*(\d{2}\.\d{2}\.\d{4})', text)
    obdobie = f"{obdobie_match.group(1)} - {obdobie_match.group(2)}" if obdobie_match else "Neuvedené"

    # 3. Silová zložka (Komodita)
    silova_kwh_m = re.search(r'1T za elektrinu\s*([\d\s,]+)\s*kWh', text)
    silova_eur_m = re.search(r'1T za elektrinu\s*[\d\s,]+\s*kWh\s*[\d\s,]+\s*([\d\s,]+)', text)
    
    silova_kwh = float(silova_kwh_m.group(1).replace(" ", "").replace(",", ".")) if silova_kwh_m else 0.0
    silova_eur = float(silova_eur_m.group(1).replace(" ", "").replace(",", ".")) if silova_eur_m else 0.0

    # 4. Distribúcia celkom
    distribucia_m = re.search(r'Služba:\s*Distribúcia.*?Spolu bez DPH\s*([\d\s,]+)\s*€', text, re.DOTALL)
    distribucia_eur = float(distribucia_m.group(1).replace(" ", "").replace(",", ".")) if distribucia_m else 0.0

    # 5. Detailné distribučné a regulované položky
    def extract_val(pattern):
        m = re.search(pattern, text)
        return float(m.group(1).replace(" ", "").replace(",", ".")) if m else 0.0

    pevna_distr = extract_val(r'Pevná zložka tarify za distribúciu\s*[\d\s,]+\s*A\s*[\d\s,]+\s*[\d\s,]+\s*([\d\s,]+)')
    var_distr = extract_val(r'Var\.zložka tarify za distr\.,vrát\.pren\.el\s*[\d\s,]+\s*kWh\s*[\d\s,]+\s*[\d\s,]+\s*([\d\s,]+)')
    straty_distr = extract_val(r'Tarifa za distribučné straty\s*[\d\s,]+\s*kWh\s*[\d\s,]+\s*[\d\s,]+\s*([\d\s,]+)')
    jadro_fon = extract_val(r'Tarifa za odvod do jadrového fondu\s*[\d\s,]+\s*kWh\s*[\d\s,]+\s*[\d\s,]+\s*([\d\s,]+)')
    sistem_sluzby = extract_val(r'Platba za systémové služby\s*-\s*pásmo 2\s*[\d\s,]+\s*kWh\s*[\d\s,]+\s*[\d\s,]+\s*([\d\s,]+)')
    prevadzka_sist = extract_val(r'Platba za prevádzkovanie systému\s*-\s*pásmo 2\s*[\d\s,]+\s*kWh\s*[\d\s,]+\s*[\d\s,]+\s*([\d\s,]+)')

    # 6. Pokuty a jalová elektrina
    jalovina = extract_val(r'Nevyžiad\.dod\.kapacitnej zložky jal\.en\.\s*[\d\s,]+\s*kVArh\s*[\d\s,]+\s*[\d\s,]+\s*([\d\s,]+)')

    # 7. Celkové sumy
    spolu_bez_dph = extract_val(r'Spolu za združ\.dod\.elektriny\s*([\d\s,]+)\s*19,00')
    dph = extract_val(r'Spolu za združ\.dod\.elektriny\s*[\d\s,]+\s*19,00\s*([\d\s,]+)')
    spolu_s_dph = extract_val(r'Spolu za združ\.dod\.elektriny\s*[\d\s,]+\s*19,00\s*[\d\s,]+\s*([\d\s,]+)')

    return {
        "EIC": eic_code,
        "Odberné miesto": om_info["nazov"],
        "Lokalita": om_info["lokalita"],
        "Číslo faktúry (VS)": vs_num,
        "Fakturačné obdobie": obdobie,
        "Spotreba (kWh)": silova_kwh,
        "Silová zložka (€)": silova_eur,
        "Distribúcia celkom (€)": distribucia_eur,
        "Pevná zložka (ističe) (€)": pevna_distr,
        "Variabilná distribúcia (€)": var_distr,
        "Distribučné straty (€)": straty_distr,
        "Jadrový fond (€)": jadro_fon,
        "Systémové služby (€)": sistem_sluzby,
        "Prevádzkovanie systému (€)": prevadzka_sist,
        "Jalová elektrina (€)": jalovina,
        "Spolu bez DPH (€)": spolu_bez_dph if spolu_bez_dph > 0 else (silova_eur + distribucia_eur),
        "DPH (€)": dph,
        "Spolu s DPH (€)": spolu_s_dph
    }

@st.cache_data
def load_data():
    file_path = "Spotreby energií.xlsx"
    xls = pd.ExcelFile(file_path)
    months = ['Január', 'Február', 'Marec', 'Apríl', 'Máj', 'Jún', 'Júl', 'August', 'September']
    
    # 1. Načítanie FVE dáta 2026
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
    
    # 2. DYNAMICKÉ NAČÍTANIE VŠETKÝCH ODBERNÝCH MIEST Z PREHĽAD ELEKTRO 2026
    df_elec_raw = pd.read_excel(xls, sheet_name='Prehľad elektro.2026', header=None)
    
    records = []
    current_section = "Všeobecné"
    
    for idx in range(len(df_elec_raw)):
        row = df_elec_raw.iloc[idx]
        val0 = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        val1 = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
        combined_text = f"{val0} {val1}".strip()
        
        if any(keyword in combined_text.lower() for keyword in ["trafostanica", "vodomer", "odpis meračov"]):
            current_section = combined_text
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

# --- APLIKÁCIA FILTROV ---
filtered_fve = fve_df[fve_df['Mesiac'].isin(selected_months)].copy()
filtered_elec = elec_df[elec_df['Mesiac'].isin(selected_months)].copy()

if selected_lokalita != "Všetky lokality":
    filtered_fve = filtered_fve[filtered_fve['Lokalita_raw'] == selected_lokalita]
    filtered_elec = filtered_elec[filtered_elec['Lokalita'] == selected_lokalita]

filtered_fve['Mesiac'] = pd.Categorical(filtered_fve['Mesiac'], categories=month_order, ordered=True)
filtered_elec['Mesiac'] = pd.Categorical(filtered_elec['Mesiac'], categories=month_order, ordered=True)
filtered_fve = filtered_fve.sort_values('Mesiac')
filtered_elec = filtered_elec.sort_values('Mesiac')

# Kľúčové ukazovatele
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

# DEFINÍCIA 4 ZÁLOŽIEK (PRIDANÁ ZÁLOŽKA PRE FAKTÚRY)
tab1, tab2, tab3, tab4 = st.tabs([
    "📉 Spotreby & Náklady Odberných Miest", 
    "☀️ Výkon Fotovoltiky (FVE)", 
    "💶 Jednotková cena za 1 kWh",
    "📄 Analýza PDF Faktúr (ZSE)"
])

# ZÁLOŽKA 1: SPOTREBY A NÁKLADY
with tab1:
    st.subheader(f"Mesačný vývoj spotreby a nákladov ({selected_lokalita})")
    
    all_om = list(filtered_elec['Odberné miesto'].unique())
    selected_om = st.multiselect(
        "Vyber odberné miesta pre zobrazenie v grafoch:", 
        options=all_om, 
        default=all_om, 
        key="om_filter_tab1"
    )
    
    sub_elec = filtered_elec[filtered_elec['Odberné miesto'].isin(selected_om)]
    
    if not sub_elec.empty:
        col_left, col_right = st.columns(2)
        with col_left:
            fig_elec_kwh = px.bar(
                sub_elec, x='Mesiac', y='Spotreba (kWh)', color='Odberné miesto', barmode='group',
                title="Spotreba elektriny (kWh)", labels={'Spotreba (kWh)': 'Spotreba (kWh)', 'Mesiac': 'Mesiac'},
                template="plotly_dark"
            )
            fig_elec_kwh.update_xaxes(categoryorder='array', categoryarray=month_order)
            fig_elec_kwh.update_layout(font=dict(size=14))
            st.plotly_chart(fig_elec_kwh, use_container_width=True)
            
        with col_right:
            fig_elec_eur = px.bar(
                sub_elec, x='Mesiac', y='Náklady (€)', color='Odberné miesto', barmode='group',
                title="Mesačné náklady (€)", labels={'Náklady (€)': 'Náklady (€)', 'Mesiac': 'Mesiac'},
                template="plotly_dark"
            )
            fig_elec_eur.update_xaxes(categoryorder='array', categoryarray=month_order)
            fig_elec_eur.update_layout(font=dict(size=14))
            st.plotly_chart(fig_elec_eur, use_container_width=True)
    else:
        st.info("Prosím, vyberte aspoň jedno odberné miesto.")

    st.markdown("#### Tabuľkový prehľad spotreby a nákladov")
    st.dataframe(filtered_elec, use_container_width=True, hide_index=True)

# ZÁLOŽKA 2: FOTOVOLTIKA
with tab2:
    st.subheader(f"☀️ Výkon a úspory FVE ({selected_lokalita})")
    
    summary_fve = filtered_fve.groupby('Lokalita_raw').agg({
        'Vyroba_kWh': 'sum',
        'Uspora_EUR': 'sum',
        'Podiel_Pokrytia_Pct': 'mean'
    }).reset_index()

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
        fig_pie_kwh = px.pie(
            summary_fve, names='Label_KWh', values='Vyroba_kWh',
            title="Podiel lokalít na celkovej výrobe FVE",
            hole=0.4, template="plotly_dark"
        )
        fig_pie_kwh.update_traces(textinfo='percent', textfont=dict(size=16, color='white'))
        fig_pie_kwh.update_layout(font=dict(size=14))
        st.plotly_chart(fig_pie_kwh, use_container_width=True)
        
    with col_sum2:
        fig_pie_eur = px.pie(
            summary_fve, names='Label_EUR', values='Uspora_EUR',
            title="Podiel lokalít na finančných úsporách",
            hole=0.4, template="plotly_dark"
        )
        fig_pie_eur.update_traces(textinfo='percent', textfont=dict(size=16, color='white'))
        fig_pie_eur.update_layout(font=dict(size=14))
        st.plotly_chart(fig_pie_eur, use_container_width=True)

    st.divider()
    st.subheader("Mesačný vývoj a detailné porovnanie")

    df_kwh = filtered_fve.copy()
    df_eur = filtered_fve.copy()
    df_pct = filtered_fve.copy()
    
    lbl_kwh, lbl_eur, lbl_pct = [], [], []
    for idx, row in filtered_fve.iterrows():
        loc = row['Lokalita_raw']
        sub = filtered_fve[filtered_fve['Lokalita_raw'] == loc]
        tot_kwh = sub['Vyroba_kWh'].sum()
        tot_eur = sub['Uspora_EUR'].sum()
        
        lbl_kwh.append(f"{loc} (Spolu: {tot_kwh:,.0f} kWh)")
        lbl_eur.append(f"{loc} (Spolu: {tot_eur:,.0f} €)")
        lbl_pct.append(f"{loc} (Spolu: {tot_kwh:,.0f} kWh / {tot_eur:,.0f} €)")
        
    df_kwh['Lokalita_FVE'] = lbl_kwh
    df_eur['Lokalita_FVE'] = lbl_eur
    df_pct['Lokalita_FVE'] = lbl_pct

    fig_fve_kwh = px.bar(
        df_kwh, x='Mesiac', y='Vyroba_kWh', color='Lokalita_FVE', barmode='group',
        text='Vyroba_kWh',
        title="Mesačná výroba FVE (kWh)", labels={'Vyroba_kWh': 'Výroba (kWh)', 'Mesiac': 'Mesiac'},
        template="plotly_dark", height=500
    )
    fig_fve_kwh.update_traces(texttemplate='%{text:,.0f} kWh', textposition='outside', textfont=dict(size=20, color='white', family="Arial Black"))
    fig_fve_kwh.update_xaxes(categoryorder='array', categoryarray=month_order)
    fig_fve_kwh.update_layout(font=dict(size=14), margin=dict(t=100, b=20, l=20, r=20), bargap=0.15, bargroupgap=0.05)
    st.plotly_chart(fig_fve_kwh, use_container_width=True)

    st.divider()

    fig_fve_eur = px.bar(
        df_eur, x='Mesiac', y='Uspora_EUR', color='Lokalita_FVE', barmode='group',
        text='Uspora_EUR',
        title="Finančná úspora FVE v €", labels={'Uspora_EUR': 'Úspora (€)', 'Mesiac': 'Mesiac'},
        template="plotly_dark", height=500
    )
    fig_fve_eur.update_traces(texttemplate='%{text:,.0f} €', textposition='outside', textfont=dict(size=20, color='white', family="Arial Black"))
    fig_fve_eur.update_xaxes(categoryorder='array', categoryarray=month_order)
    fig_fve_eur.update_layout(font=dict(size=14), margin=dict(t=100, b=20, l=20, r=20), bargap=0.15, bargroupgap=0.05)
    st.plotly_chart(fig_fve_eur, use_container_width=True)

    st.divider()

    fig_pct = px.bar(
        df_pct, x='Mesiac', y='Podiel_Pokrytia_Pct', color='Lokalita_FVE', barmode='group',
        text='Podiel_Pokrytia_Pct',
        title="Pokrytie spotreby objektu z FVE (%)", labels={'Podiel_Pokrytia_Pct': 'Pokrytie (%)', 'Mesiac': 'Mesiac'},
        template="plotly_dark", height=500
    )
    fig_pct.update_traces(texttemplate='%{text:.1f} %', textposition='outside', textfont=dict(size=20, color='white', family="Arial Black"))
    fig_pct.update_xaxes(categoryorder='array', categoryarray=month_order)
    fig_pct.update_layout(font=dict(size=14), margin=dict(t=100, b=20, l=20, r=20), bargap=0.15, bargroupgap=0.05)
    st.plotly_chart(fig_pct, use_container_width=True)

    st.markdown("#### Detailné dáta FVE 2026")
    st.dataframe(filtered_fve, use_container_width=True, hide_index=True)

# ZÁLOŽKA 3: JEDNOTKOVÁ CENA ZA 1 KWH
with tab3:
    st.subheader(f"Porovnanie a vývoj jednotkovej ceny elektriny (€ / kWh) – {selected_lokalita}")
    
    if not filtered_elec.empty:
        fig_price_trend = px.line(
            filtered_elec, 
            x='Mesiac', 
            y='Cena (€/kWh)', 
            color='Odberné miesto', 
            markers=True,
            title="Vývoj ceny 1 kWh podľa odberných miest (€/kWh)",
            labels={'Cena (€/kWh)': 'Cena (€/kWh)', 'Mesiac': 'Mesiac'},
            template="plotly_dark"
        )
        fig_price_trend.update_xaxes(categoryorder='array', categoryarray=month_order)
        fig_price_trend.update_yaxes(tickformat=".4f")
        fig_price_trend.update_layout(font=dict(size=14))
        st.plotly_chart(fig_price_trend, use_container_width=True)

    st.markdown("### Priemerná cena 1 kWh za vybrané obdobie")
    
    records_price = []
    for om in filtered_elec['Odberné miesto'].unique():
        sub = filtered_elec[filtered_elec['Odberné miesto'] == om]
        tot_kwh = sub['Spotreba (kWh)'].sum()
        tot_eur = sub['Náklady (€)'].sum()
        avg_p = (tot_eur / tot_kwh) if tot_kwh > 0 else 0.0
        records_price.append({
            'Odberné miesto': om,
            'Celková spotreba (kWh)': tot_kwh,
            'Celkové náklady (€)': tot_eur,
            'Priemerná cena (€/kWh)': avg_p
        })
    df_price_table = pd.DataFrame(records_price)
    
    if not df_price_table.empty:
        st.dataframe(
            df_price_table.style.format({
                'Celková spotreba (kWh)': '{:,.0f}',
                'Celkové náklady (€)': '{:,.2f} €',
                'Priemerná cena (€/kWh)': '{:.4f} €'
            }),
            use_container_width=True,
            hide_index=True
        )

# ZÁLOŽKA 4: AUTOMATICKÉ PARSOVANIE PDF FAKTÚR (NOVINKA)
with tab4:
    st.subheader("📄 Automatická analýza a parsovanie ZSE faktúr z PDF")
    st.markdown("Nahrajte jednu alebo viaceré vyúčtovacie PDF faktúry od ZSE. Systém automaticky zistí EIC kód, priradí odberné miesto a rozbije sumu na silovú zložku, distribúciu a poplatky.")

    uploaded_files = st.file_uploader(
        "Sem pretiahnite PDF faktúry (zse_faktura.pdf):", 
        type=["pdf"], 
        accept_multiple_files=True
    )

    if uploaded_files:
        parsed_results = []
        for pdf_file in uploaded_files:
            try:
                data = parse_zse_pdf(pdf_file)
                parsed_results.append(data)
            except Exception as e:
                st.error(f"Chyba pri spracovaní súboru {pdf_file.name}: {e}")

        if parsed_results:
            df_invoices = pd.DataFrame(parsed_results)
            st.success(f"Úspešne spracovaných faktúr: {len(df_invoices)}")
            
            # Zobrazenie výslednej prehľadnej tabuľky
            st.markdown("### 📊 Extrahované údaje z faktúr")
            st.dataframe(df_invoices, use_container_width=True, hide_index=True)

            st.divider()
            
            # Vizuálny rozpad nákladov pre prvú nahrávanú faktúru (alebo agregát)
            st.markdown("### 🔍 Štruktúra nákladov faktúry (Silová zložka vs. Distribúcia a poplatky)")
            
            for idx, row in df_invoices.iterrows():
                st.markdown(f"#### Faktúra č. {row['Číslo faktúry (VS)']} – {row['Odberné miesto']} ({row['Fakturačné obdobie']})")
                
                c_f1, c_f2, c_f3, c_f4 = st.columns(4)
                c_f1.metric("Spotreba elektriny", f"{row['Spotreba (kWh)']:,.0f} kWh")
                c_f2.metric("Silová zložka (Komodita)", f"{row['Silová zložka (€)']:,.2f} €")
                c_f3.metric("Distribúcia & Poplatky", f"{row['Distribúcia celkom (€)']:,.2f} €")
                
                jal_val = row['Jalová elektrina (€)']
                c_f4.metric("Jalová elektrina / Pokuty", f"{jal_val:,.2f} €", delta_color="inverse" if jal_val > 0 else "normal")

                # Koláčový graf rozpadu položiek danej faktúry
                cost_breakdown = {
                    "Silová zložka": row['Silová zložka (€)'],
                    "Pevná zložka (ističe/kapacita)": row['Pevná zložka (ističe) (€)'],
                    "Variabilná distribúcia": row['Variabilná distribúcia (€)'],
                    "Distribučné straty": row['Distribučné straty (€)'],
                    "Jadrový fond": row['Jadrový fond (€)'],
                    "Systémové služby": row['Systémové služby (€)'],
                    "Prevádzkovanie systému": row['Prevádzkovanie systému (€)'],
                    "Jalová elektrina": row['Jalová elektrina (€)']
                }
                
                df_break = pd.DataFrame([
                    {"Položka": k, "Suma (€)": v} for k, v in cost_breakdown.items() if v > 0
                ])

                fig_break = px.pie(
                    df_break, names='Položka', values='Suma (€)',
                    title=f"Podiel jednotlivých zložiek na cene bez DPH ({row['Odberné miesto']})",
                    hole=0.4, template="plotly_dark"
                )
                fig_break.update_traces(textinfo='percent+label', textfont=dict(size=14, color='white'))
                fig_break.update_layout(font=dict(size=14))
                st.plotly_chart(fig_break, use_container_width=True)
                st.divider()
