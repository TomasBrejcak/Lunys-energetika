import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os

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

st.title("⚡ Energetický Dashboard Lunys ")
st.markdown("Interaktívny nástroj pre vyhodnotenie spotreby elektriny, nákladov, úspor z FVE (Poprad & Bratislava) a priemernej ceny za 1 kWh.")

@st.cache_data
def load_data():
    file_path = "Spotreby energií.xlsx"
    xls = pd.ExcelFile(file_path)
    
    # 1. Načítanie FVE dáta 2026
    df_fve_raw = pd.read_excel(xls, sheet_name='Prehľad po FVE.2026')
    months = ['Január', 'Február', 'Marec', 'Apríl', 'Máj', 'Jún', 'Júl', 'August', 'September']
    
    poprad_kwh = df_fve_raw.iloc[1, 2:11].values.astype(float)
    poprad_eur = df_fve_raw.iloc[2, 2:11].values.astype(float)
    poprad_pct = df_fve_raw.iloc[3, 2:11].values.astype(float)
    
    ba_kwh = df_fve_raw.iloc[4, 2:11].values.astype(float)
    ba_eur = df_fve_raw.iloc[5, 2:11].values.astype(float)
    ba_pct = df_fve_raw.iloc[6, 2:11].values.astype(float)
    
    fve_df = pd.DataFrame({
        'Mesiac': months * 2,
        'Lokalita': ['Poprad'] * 9 + ['Bratislava'] * 9,
        'Vyroba_kWh': list(poprad_kwh) + list(ba_kwh),
        'Uspora_EUR': list(poprad_eur) + list(ba_eur),
        'Podiel_Pokrytia_Pct': list(poprad_pct) + list(ba_pct)
    })
    
    # 2. Načítanie Spotreby Elektro 2026
    df_elec_raw = pd.read_excel(xls, sheet_name='Prehľad elektro.2026')
    
    items = [
        ('Trafo Haramia (Celkom)', 1, 2),
        ('Bananáreň', 3, 4),
        ('Zemiakareň', 5, 6),
        ('Trafo Pebek (Celkom)', 11, 12),
        ('Skladová hala I.', 13, 14),
        ('AB Správa', 16, 17),
        ('Beňovský Apartments', 27, 28)
    ]
    
    records = []
    for name, r_kwh, r_eur in items:
        kwh_vals = df_elec_raw.iloc[r_kwh, 2:11].values.astype(float)
        eur_vals = df_elec_raw.iloc[r_eur, 2:11].values.astype(float)
        for m, k, e in zip(months, kwh_vals, eur_vals):
            records.append({
                'Mesiac': m,
                'Odberne_Miesto': name,
                'Spotreba_kWh': k,
                'Naklad_EUR': e,
                'Jednotkova_Cena_EUR_kWh': e / k if k > 0 else 0
            })
            
    elec_df = pd.DataFrame(records)
    return fve_df, elec_df

try:
    fve_df, elec_df = load_data()
except Exception as e:
    st.error(f"Chyba pri načítaní Excel súboru: {e}")
    st.stop()

# Filtre v sidebare
st.sidebar.subheader("🔍 Filtre")
selected_months = st.sidebar.multiselect("Vyber mesiace", options=fve_df['Mesiac'].unique(), default=fve_df['Mesiac'].unique())

filtered_fve = fve_df[fve_df['Mesiac'].isin(selected_months)]
filtered_elec = elec_df[elec_df['Mesiac'].isin(selected_months)]

# Výpočet celkových dát pre hlavné odberné miesta
main_om = filtered_elec[filtered_elec['Odberne_Miesto'].isin(['Trafo Haramia (Celkom)', 'Trafo Pebek (Celkom)', 'Beňovský Apartments'])]
total_elec_kwh = main_om['Spotreba_kWh'].sum()
total_elec_eur = main_om['Naklad_EUR'].sum()

# Výpočet priemernej ceny 1 kWh zo siete
avg_price_per_kwh = (total_elec_eur / total_elec_kwh) if total_elec_kwh > 0 else 0

total_fve_kwh = filtered_fve['Vyroba_kWh'].sum()
total_fve_eur = filtered_fve['Uspora_EUR'].sum()

# KPI Sumár (5 stĺpcov)
st.subheader("📊 Kľúčové ukazovatele 2026 (Vybrané mesiace)")
c1, c2, c3, c4, c5 = st.columns(5)

c1.metric("🔌 Odber zo siete", f"{total_elec_kwh:,.0f} kWh")
c2.metric("💶 Náklady na sieť", f"{total_elec_eur:,.2f} €")
c3.metric("☀️ Výroba FVE", f"{total_fve_kwh:,.0f} kWh")
c4.metric("💰 Úspora FVE", f"{total_fve_eur:,.2f} €")
c5.metric("⚡ Priemerná cena 1 kWh", f"{avg_price_per_kwh:.4f} €/kWh")

st.divider()

# Usporiadanie záložiek: 1. Spotreba, 2. Fotovoltika, 3. Jednotková cena
tab1, tab2, tab3 = st.tabs([
    "📉 Spotreby & Náklady Odberných Miest", 
    "☀️ Výkon Fotovoltiky (FVE)", 
    "💶 Jednotková cena za 1 kWh"
])

# ZÁLOŽKA 1: SPOTREBY A NÁKLADY
with tab1:
    st.subheader("Mesačný vývoj spotreby a nákladov podľa prevádzok")
    
    om_options = filtered_elec['Odberne_Miesto'].unique()
    selected_om = st.multiselect("Vyber odberné miesta pre zobrazenie:", options=om_options, default=['Bananáreň', 'Zemiakareň', 'Skladová hala I.', 'Beňovský Apartments'])
    
    sub_elec = filtered_elec[filtered_elec['Odberne_Miesto'].isin(selected_om)]
    
    col_left, col_right = st.columns(2)
    with col_left:
        fig_elec_kwh = px.bar(
            sub_elec, x='Mesiac', y='Spotreba_kWh', color='Odberne_Miesto', barmode='group',
            title="Spotreba elektriny (kWh)", labels={'Spotreba_kWh': 'Spotreba (kWh)', 'Mesiac': 'Mesiac'}
        )
        st.plotly_chart(fig_elec_kwh, use_container_width=True)
        
    with col_right:
        fig_elec_eur = px.bar(
            sub_elec, x='Mesiac', y='Naklad_EUR', color='Odberne_Miesto', barmode='group',
            title="Mesačné náklady (€)", labels={'Naklad_EUR': 'Náklady (€)', 'Mesiac': 'Mesiac'}
        )
        st.plotly_chart(fig_elec_eur, use_container_width=True)

    st
