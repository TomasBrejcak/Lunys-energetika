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

st.title("⚡ Energetický Dashboard Lunys")
st.markdown("Interaktívny nástroj pre vyhodnotenie spotreby elektriny, nákladov, úspor z FVE (Poprad & Bratislava) a priemernej ceny za 1 kWh.")

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
        'Lokalita': ['Poprad'] * 9 + ['Bratislava'] * 9,
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
        
        # Sledovanie nadpisov sekcií / trafostaníc
        if any(keyword in combined_text.lower() for keyword in ["trafostanica", "vodomer", "odpis meračov"]):
            current_section = combined_text
            continue
            
        # Ak riadok obsahuje "spotreba"
       # Ak riadok obsahuje "spotreba"
      # Ak riadok obsahuje "spotreba"
        if "spotreba" in combined_text.lower():
            raw_lower = combined_text.lower()
            
            if "bananáreň" in raw_lower:
                clean_name = "Bananáreň"
            elif "zemiakareň" in raw_lower:
                clean_name = "Zemiakareň"
            elif "feld" in raw_lower or "bratislava" in raw_lower:
                clean_name = "Bratislava – Trafo Feld"
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
            # Úprava formátu
            if not clean_name:
                clean_name = current_section
            
            kwh_vals = pd.to_numeric(row.values[2:11], errors='coerce')
            
            # Hľadanie zodpovedajúcich nákladov v nasledujúcich riadkoch
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
                        'Odberné miesto': clean_name,
                        'Spotreba (kWh)': k_val,
                        'Náklady (€)': e_val,
                        'Cena (€/kWh)': (e_val / k_val) if k_val > 0 else 0.0
                    })
                    
    elec_df = pd.DataFrame(records)
    
    # Ak by dynamický parser nenašiel nič, použijeme zoznam všetkých prevádzok
    if elec_df.empty:
        tracked_items = [
            ("Trafostanica Haramia (Celkom)", "Nové trafo pre SH II a SH III. spotreba", "Nové trafo pre SH II a SH III. - náklad"),
            ("Bananáreň", "Bananáreň spotreba", "Bananáreň - náklad"),
            ("Zemiakareň", "Zemiakareň spotreba", "Zemiakareň - náklad"),
            ("Trafostanica Pebek (Celkom)", "Staré trafo Pebek spotreba", "Staré trafo Pebek náklad"),
            ("Skladová hala I.", "Skladová hala I. - spotreba", "Skladová hala I. - náklad"),
            ("AB Správa", "AB Správa spotreba", "AB Správa- náklad"),
            ("Hlavný elektromer / Náomcovia", "Hlavný elektromer /arál nájomcovia/ spotreba", "Hlavný elektromer /arál nájomcovia/ - náklad"),
            ("Beňovský Apartments", "Beňovský Apartments spotreba", "Beňovský Apartments náklad")
        ]
        fallback_records = []
        def extract_row_data(keyword):
            for idx, row in df_elec_raw.iterrows():
                row_text = str(row.values)
                if keyword.lower() in row_text.lower():
                    return pd.to_numeric(row.values[2:11], errors='coerce')
            return None

        for name, kwh_key, eur_key in tracked_items:
            kwh_vals = extract_row_data(kwh_key)
            eur_vals = extract_row_data(eur_key)
            if kwh_vals is not None and eur_vals is not None:
                for m, k, e in zip(months, kwh_vals, eur_vals):
                    k_val = float(k) if pd.notna(k) else 0.0
                    e_val = float(e) if pd.notna(e) else 0.0
                    fallback_records.append({
                        'Mesiac': m,
                        'Odberné miesto': name,
                        'Spotreba (kWh)': k_val,
                        'Náklady (€)': e_val,
                        'Cena (€/kWh)': (e_val / k_val) if k_val > 0 else 0.0
                    })
        elec_df = pd.DataFrame(fallback_records)

    return fve_df, elec_df

try:
    fve_df, elec_df = load_data()
except Exception as e:
    st.error(f"Chyba pri načítaní Excel súboru: {e}")
    st.stop()

# Filtre v sidebare
st.sidebar.subheader("🔍 Filtre")
month_order = list(fve_df['Mesiac'].unique())
selected_months = st.sidebar.multiselect(
    "Vyber mesiace", 
    options=month_order, 
    default=month_order
)

# Filtrovanie dát
filtered_fve = fve_df[fve_df['Mesiac'].isin(selected_months)].copy()
filtered_elec = elec_df[elec_df['Mesiac'].isin(selected_months)].copy()

filtered_fve['Mesiac'] = pd.Categorical(filtered_fve['Mesiac'], categories=month_order, ordered=True)
filtered_elec['Mesiac'] = pd.Categorical(filtered_elec['Mesiac'], categories=month_order, ordered=True)
filtered_fve = filtered_fve.sort_values('Mesiac')
filtered_elec = filtered_elec.sort_values('Mesiac')

# Kľúčové ukazovatele za VŠETKY prevádzky
total_elec_kwh = filtered_elec['Spotreba (kWh)'].sum()
total_elec_eur = filtered_elec['Náklady (€)'].sum()
avg_price_per_kwh = (total_elec_eur / total_elec_kwh) if total_elec_kwh > 0 else 0.0

total_fve_kwh = filtered_fve['Vyroba_kWh'].sum()
total_fve_eur = filtered_fve['Uspora_EUR'].sum()

st.subheader("📊 Kľúčové ukazovatele 2026 (Vybrané mesiace)")
c1, c2, c3, c4, c5 = st.columns(5)

c1.metric("🔌 Odber zo siete (Spolu)", f"{total_elec_kwh:,.0f} kWh")
c2.metric("💶 Náklady na sieť spolu", f"{total_elec_eur:,.2f} €")
c3.metric("☀️ Výroba FVE", f"{total_fve_kwh:,.0f} kWh")
c4.metric("💰 Úspora FVE", f"{total_fve_eur:,.2f} €")
c5.metric("⚡ Priemerná cena 1 kWh", f"{avg_price_per_kwh:.4f} €/kWh")

st.divider()

# Definícia záložiek
tab1, tab2, tab3 = st.tabs([
    "📉 Spotreby & Náklady Odberných Miest", 
    "☀️ Výkon Fotovoltiky (FVE)", 
    "💶 Jednotková cena za 1 kWh"
])

# ZÁLOŽKA 1: SPOTREBY A NÁKLADY
with tab1:
    st.subheader("Mesačný vývoj spotreby a nákladov podľa prevádzok")
    
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
                title="Spotreba elektriny (kWh)", labels={'Spotreba (kWh)': 'Spotreba (kWh)', 'Mesiac': 'Mesiac'}
            )
            fig_elec_kwh.update_xaxes(categoryorder='array', categoryarray=month_order)
            st.plotly_chart(fig_elec_kwh, use_container_width=True)
            
        with col_right:
            fig_elec_eur = px.bar(
                sub_elec, x='Mesiac', y='Náklady (€)', color='Odberné miesto', barmode='group',
                title="Mesačné náklady (€)", labels={'Náklady (€)': 'Náklady (€)', 'Mesiac': 'Mesiac'}
            )
            fig_elec_eur.update_xaxes(categoryorder='array', categoryarray=month_order)
            st.plotly_chart(fig_elec_eur, use_container_width=True)
    else:
        st.info("Prosím, vyberte aspoň jedno odberné miesto.")

    st.markdown("#### Tabuľkový prehľad spotreby a nákladov")
    st.dataframe(filtered_elec, use_container_width=True, hide_index=True)

# ZÁLOŽKA 2: FOTOVOLTIKA
with tab2:
    st.subheader("Porovnanie výroby a úspor FVE: Poprad vs. Bratislava")
    
    col_left, col_right = st.columns(2)
    with col_left:
        fig_fve_kwh = px.bar(
            filtered_fve, x='Mesiac', y='Vyroba_kWh', color='Lokalita', barmode='group',
            title="Mesačná výroba FVE (kWh)", labels={'Vyroba_kWh': 'Výroba (kWh)', 'Mesiac': 'Mesiac'}
        )
        fig_fve_kwh.update_xaxes(categoryorder='array', categoryarray=month_order)
        st.plotly_chart(fig_fve_kwh, use_container_width=True)
        
    with col_right:
        fig_fve_eur = px.line(
            filtered_fve, x='Mesiac', y='Uspora_EUR', color='Lokalita', markers=True,
            title="Finančná úspora FVE v €", labels={'Uspora_EUR': 'Úspora (€)', 'Mesiac': 'Mesiac'}
        )
        fig_fve_eur.update_xaxes(categoryorder='array', categoryarray=month_order)
        st.plotly_chart(fig_fve_eur, use_container_width=True)

    st.markdown("### % Podiel FVE na celkovej spotrebe prevádzky")
    fig_pct = px.bar(
        filtered_fve, x='Mesiac', y='Podiel_Pokrytia_Pct', color='Lokalita', barmode='group',
        title="Pokrytie spotreby objektu z FVE (%)", labels={'Podiel_Pokrytia_Pct': 'Pokrytie (%)', 'Mesiac': 'Mesiac'}
    )
    fig_pct.update_xaxes(categoryorder='array', categoryarray=month_order)
    st.plotly_chart(fig_pct, use_container_width=True)

    st.markdown("#### Detailné dáta FVE 2026")
    st.dataframe(filtered_fve, use_container_width=True, hide_index=True)

# ZÁLOŽKA 3: JEDNOTKOVÁ CENA ZA 1 KWH
with tab3:
    st.subheader("Porovnanie a vývoj jednotkovej ceny elektriny (€ / kWh)")
    st.markdown("Tento prehľad zobrazuje, koľko nás stála 1 kWh elektriny v jednotlivých mesiacoch po prevádzkach.")
    
    if not filtered_elec.empty:
        fig_price_trend = px.line(
            filtered_elec, 
            x='Mesiac', 
            y='Cena (€/kWh)', 
            color='Odberné miesto', 
            markers=True,
            title="Vývoj ceny 1 kWh podľa odberných miest (€/kWh)",
            labels={'Cena (€/kWh)': 'Cena (€/kWh)', 'Mesiac': 'Mesiac'}
        )
        fig_price_trend.update_xaxes(categoryorder='array', categoryarray=month_order)
        fig_price_trend.update_yaxes(tickformat=".4f")
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
    else:
        st.info("Nie sú k dispozícii žiadne dáta pre vybrané filtre.")
