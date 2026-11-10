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
st.markdown("Interaktívny nástroj pre vyhodnotenie spotreby elektriny, nákladov a úspor z fotovoltických elektrární (Poprad & Bratislava).")

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

# KPI Sumár
st.subheader("📊 Kľúčové ukazovatele 2026 (Jan - Sep)")
c1, c2, c3, c4 = st.columns(4)

total_fve_kwh = filtered_fve['Vyroba_kWh'].sum()
total_fve_eur = filtered_fve['Uspora_EUR'].sum()
total_elec_kwh = filtered_elec[filtered_elec['Odberne_Miesto'].isin(['Trafo Haramia (Celkom)', 'Trafo Pebek (Celkom)', 'Beňovský Apartments'])]['Spotreba_kWh'].sum()
total_elec_eur = filtered_elec[filtered_elec['Odberne_Miesto'].isin(['Trafo Haramia (Celkom)', 'Trafo Pebek (Celkom)', 'Beňovský Apartments'])]['Naklad_EUR'].sum()

c1.metric("☀️ Celková výroba FVE", f"{total_fve_kwh:,.0f} kWh")
c2.metric("💰 Celková úspora FVE", f"{total_fve_eur:,.2f} €")
c3.metric("🔌 Odber zo siete (Hlavné OM)", f"{total_elec_kwh:,.0f} kWh")
c4.metric("💶 Náklady na sieťovú energiu", f"{total_elec_eur:,.2f} €")

st.divider()

# Záložky
tab1, tab2, tab3 = st.tabs(["☀️ Výkon Fotovoltiky (FVE)", "📉 Spotreby & Náklady Odberných Miest", "📋 Surové Dáta"])

with tab1:
    st.subheader("Porovnanie výroby a úspor FVE: Poprad vs. Bratislava")
    
    col_left, col_right = st.columns(2)
    with col_left:
        fig_fve_kwh = px.bar(
            filtered_fve, x='Mesiac', y='Vyroba_kWh', color='Lokalita', barmode='group',
            title="Mesačná výroba FVE (kWh)", labels={'Vyroba_kWh': 'Výroba (kWh)'}
        )
        st.plotly_chart(fig_fve_kwh, use_container_width=True)
        
    with col_right:
        fig_fve_eur = px.line(
            filtered_fve, x='Mesiac', y='Uspora_EUR', color='Lokalita', markers=True,
            title="Finančná úspora FVE v €", labels={'Uspora_EUR': 'Úspora (€)'}
        )
        st.plotly_chart(fig_fve_eur, use_container_width=True)

    st.markdown("### % Podiel FVE na celkovej spotrebe prevádzky")
    fig_pct = px.bar(
        filtered_fve, x='Mesiac', y='Podiel_Pokrytia_Pct', color='Lokalita', barmode='group',
        title="Pokrytie spotreby objektu z FVE (%)", labels={'Podiel_Pokrytia_Pct': 'Pokrytie (%)'}
    )
    st.plotly_chart(fig_pct, use_container_width=True)

with tab2:
    st.subheader("Mesačný vývoj spotreby a nákladov podľa prevádzok")
    
    om_options = filtered_elec['Odberne_Miesto'].unique()
    selected_om = st.multiselect("Vyber odberné miesta pre zobrazenie:", options=om_options, default=['Bananáreň', 'Zemiakareň', 'Skladová hala I.', 'Beňovský Apartments'])
    
    sub_elec = filtered_elec[filtered_elec['Odberne_Miesto'].isin(selected_om)]
    
    fig_elec_kwh = px.bar(
        sub_elec, x='Mesiac', y='Spotreba_kWh', color='Odberne_Miesto', barmode='group',
        title="Spotreba elektriny (kWh)"
    )
    st.plotly_chart(fig_elec_kwh, use_container_width=True)
    
    fig_unit_price = px.line(
        sub_elec, x='Mesiac', y='Jednotkova_Cena_EUR_kWh', color='Odberne_Miesto', markers=True,
        title="Vývoj efektívnej jednotkovej ceny (€ / kWh - SPOT + Marža)"
    )
    st.plotly_chart(fig_unit_price, use_container_width=True)

with tab3:
    st.subheader("Detailné tabuľkové prehľady")
    st.markdown("#### FVE 2026")
    st.dataframe(filtered_fve, use_container_width=True)
    st.markdown("#### Spotreby Elektro 2026")
    st.dataframe(filtered_elec, use_container_width=True)
