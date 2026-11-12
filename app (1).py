import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os
import re
import pdfplumber
import io

# Nastavenie stránky
st.set_page_config(
    page_title="Energetický Dashboard LUNYS 2026",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
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
st.markdown("Interaktívny nástroj pre vyhodnotenie reálnych spotrieb, nákladov (PDF faktúry + Excel pre podružné merania) a úspor z FVE.")

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

# --- PARSER ZSE FAKTÚR ---
@st.cache_data
def parse_zse_pdf(pdf_path, file_mtime=None):
    if not os.path.exists(pdf_path):
        return None
        
    with pdfplumber.open(pdf_path) as pdf:
        text = "\n".join([page.extract_text() or "" for page in pdf.pages])

    eic_match = re.search(r'(24Z[A-Z0-9]{13})', text)
    eic_code = eic_match.group(1) if eic_match else None
    
    if not eic_code or eic_code not in EIC_MAP:
        return None

    om_info = EIC_MAP[eic_code]

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

# --- NAČÍTANIE FVE A EXCEL DÁT ---
@st.cache_data
def load_excel_data():
    file_path = "Spotreby energií.xlsx"
    if not os.path.exists(file_path):
        return pd.DataFrame(), pd.DataFrame()
        
    xls = pd.ExcelFile(file_path)
    months = ['Január', 'Február', 'Marec', 'Apríl', 'Máj', 'Jún', 'Júl', 'August', 'September', 'Október', 'November', 'December']
    
    df_fve_raw = pd.read_excel(xls, sheet_name='Prehľad po FVE.2026')
    poprad_kwh = df_fve_raw.iloc[1, 2:11].values.astype(float)
    poprad_eur = df_fve_raw.iloc[2, 2:11].values.astype(float)
    poprad_pct = df_fve_raw.iloc[3, 2:11].values.astype(float)
    
    ba_kwh = df_fve_raw.iloc[4, 2:11].values.astype(float)
    ba_eur = df_fve_raw.iloc[5, 2:11].values.astype(float)
    ba_pct = df_fve_raw.iloc[6, 2:11].values.astype(float)
    
    fve_df = pd.DataFrame({
        'Mesiac': months[:9] * 2,
        'Lokalita_raw': ['Poprad'] * 9 + ['Bratislava'] * 9,
        'Vyroba_kWh': list(poprad_kwh) + list(ba_kwh),
        'Uspora_EUR': list(poprad_eur) + list(ba_eur),
        'Podiel_Pokrytia_Pct': list(poprad_pct) + list(ba_pct)
    })
    
    df_elec_raw = pd.read_excel(xls, sheet_name='Prehľad elektro.2026', header=None)
    records = []
    
    current_lokalita = "Poprad"
    for idx in range(len(df_elec_raw)):
        row = df_elec_raw.iloc[idx]
        val0 = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        val1 = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
        combined_text = f"{val0} {val1}".strip()
        
        # Ignorovať nechcené riadky s "nové trafo"
        if "nové trafo" in combined_text.lower() or "nove trafo" in combined_text.lower():
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
                elif "správa 99" in raw_lower or "sprava 99" in raw_lower or "ab správa" in raw_lower:
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
                    clean_name = combined_text.replace("spotreba", "").replace("kWh", "").strip().capitalize()
            
            kwh_vals = pd.to_numeric(row.values[2:len(months)+2], errors='coerce')
            eur_vals = None
            for next_idx in range(idx + 1, min(idx + 4, len(df_elec_raw))):
                next_row = df_elec_raw.iloc[next_idx]
                next_text = f"{next_row.iloc[0]} {next_row.iloc[1]}".lower()
                if "náklad" in next_text or "naklad" in next_text:
                    eur_vals = pd.to_numeric(next_row.values[2:len(months)+2], errors='coerce')
                    break
                    
            if eur_vals is None:
                eur_vals = [0.0] * len(months)
                
            for m, k, e in zip(months, kwh_vals, eur_vals):
                k_val = float(k) if pd.notna(k) else 0.0
                e_val = float(e) if pd.notna(e) else 0.0
                if k_val > 0 or e_val > 0:
                    records.append({
                        'Mesiac': m,
                        'Lokalita': lok,
                        'Odberné miesto': clean_name,
                        'Spotreba (kWh)': k_val,
                        'Náklady (€)': e_val,
                        'Cena (€/kWh)': (e_val / k_val) if k_val > 0 else 0.0,
                        'Zdroj': 'Excel (Podružné meranie / Plán)'
                    })
                    
    excel_elec_df = pd.DataFrame(records)
    return fve_df, excel_elec_df

fve_df, excel_elec_df = load_excel_data()

# Načítanie PDF faktúr pre hlavné miesta
all_pdf_files = [f for f in os.listdir(INVOICE_DIR) if f.lower().endswith('.pdf')] if (os.path.exists(INVOICE_DIR) and os.path.isdir(INVOICE_DIR)) else []
parsed_invoice_records = []
for f in all_pdf_files:
    f_path = os.path.join(INVOICE_DIR, f)
    res = parse_zse_pdf(f_path, os.path.getmtime(f_path))
    if res and res["Odberné miesto"] and "neznáme" not in res["Odberné miesto"].lower():
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
            'Zdroj': 'PDF Faktúra (ZSE)',
            **res
        })

pdf_elec_df = pd.DataFrame(parsed_invoice_records)

# Spojenie dát
if not pdf_elec_df.empty:
    elec_df = pd.concat([pdf_elec_df, excel_elec_df], ignore_index=True)
    elec_df = elec_df.drop_duplicates(subset=['Odberné miesto', 'Mesiac'], keep='first')
else:
    elec_df = excel_elec_df

# --- STAV A HIERARCHICKÝ VÝBER ODBERNÝCH MIEST ---
TAB_NAMES = [
    "📉 Spotreby & Náklady Odberných Miest",
    "☀️ Výkon Fotovoltiky (FVE)",
    "💶 Cena za 1 kWh",
    "📄 Analýza PDF Faktúr & Audit"
]

if 'active_tab' not in st.session_state:
    st.session_state.active_tab = TAB_NAMES[0]

def set_tab(tab_name):
    st.session_state.active_tab = tab_name

all_possible_oms = sorted(elec_df['Odberné miesto'].dropna().astype(str).unique()) if not elec_df.empty else []

# Trafostanica je nadradený elektromer. Podružné merania sú jej detaily.
OM_GROUPS = {
    "Trafostanica PEBEK": ["Skladová hala I.", "Správa 99"],
    "Trafostanica HARAMIA": ["Bananáreň", "Zemiakareň"],
}

# Zoznam všetkých reálne dostupných meraní. Pri prvom spustení vyberieme všetko.
if 'selected_oms' not in st.session_state:
    st.session_state.selected_oms = list(all_possible_oms)
else:
    st.session_state.selected_oms = [x for x in st.session_state.selected_oms if x in all_possible_oms]

# Kľúče widgetov sú stabilné a nezávislé od poradia dát.
def _om_key(prefix, name):
    return prefix + "_" + re.sub(r"[^A-Za-z0-9]+", "_", name).strip("_")

def _set_group_selection(parent, children):
    selected = set(st.session_state.get('selected_oms', []))
    group_checked = st.session_state.get(_om_key('grp', parent), False)
    if group_checked:
        selected.add(parent)
        selected.update(children)
    else:
        selected.discard(parent)
        selected.difference_update(children)
    # Synchronizácia vizuálneho stavu podružných checkboxov.
    for child in children:
        if child in all_possible_oms:
            st.session_state[_om_key('om', child)] = group_checked
    st.session_state.selected_oms = [x for x in all_possible_oms if x in selected]

def _set_child_selection(child, parent, children):
    selected = set(st.session_state.get('selected_oms', []))
    if st.session_state.get(_om_key('om', child), False):
        selected.add(child)
    else:
        selected.discard(child)
    # Rodič je vybraný iba vtedy, keď sú vybrané všetky jeho podružné merania.
    # Samotný hlavný elektromer sa do výberu pridá/odoberie spolu so skupinou.
    all_children_selected = all(x in selected for x in children)
    parent_key = _om_key('grp', parent)
    st.session_state[parent_key] = all_children_selected
    if all_children_selected:
        selected.add(parent)
    else:
        selected.discard(parent)
    st.session_state.selected_oms = [x for x in all_possible_oms if x in selected]

def _sync_single_om(name):
    selected = set(st.session_state.get('selected_oms', []))
    if st.session_state.get(_om_key('om', name), False):
        selected.add(name)
    else:
        selected.discard(name)
    st.session_state.selected_oms = [x for x in all_possible_oms if x in selected]

# Inicializácia stavu widgetov pred ich vykreslením.
for parent, children in OM_GROUPS.items():
    existing_children = [x for x in children if x in all_possible_oms]
    if parent in all_possible_oms:
        parent_key = _om_key('grp', parent)
        if parent_key not in st.session_state:
            st.session_state[parent_key] = (parent in st.session_state.selected_oms or
                                            (bool(existing_children) and all(x in st.session_state.selected_oms for x in existing_children)))
        for child in existing_children:
            child_key = _om_key('om', child)
            if child_key not in st.session_state:
                st.session_state[child_key] = child in st.session_state.selected_oms

for om in all_possible_oms:
    if not any(om == parent or om in children for parent, children in OM_GROUPS.items()):
        key = _om_key('om', om)
        if key not in st.session_state:
            st.session_state[key] = om in st.session_state.selected_oms

st.sidebar.subheader("🏢 Výber odberných miest")
if all_possible_oms:
    col_s1, col_s2 = st.sidebar.columns(2)
    if col_s1.button("Vybrať všetko", key="btn_all_oms", use_container_width=True):
        st.session_state.selected_oms = list(all_possible_oms)
        for parent, children in OM_GROUPS.items():
            st.session_state[_om_key('grp', parent)] = True
            for child in children:
                if child in all_possible_oms:
                    st.session_state[_om_key('om', child)] = True
        for om in all_possible_oms:
            st.session_state[_om_key('om', om)] = True
        st.rerun()
    if col_s2.button("Zrušiť všetko", key="btn_none_oms", use_container_width=True):
        st.session_state.selected_oms = []
        for parent, children in OM_GROUPS.items():
            st.session_state[_om_key('grp', parent)] = False
        for om in all_possible_oms:
            st.session_state[_om_key('om', om)] = False
        st.rerun()

    st.sidebar.markdown("**📍 Poprad**")
    poprad_oms = set(elec_df.loc[elec_df['Lokalita'] == 'Poprad', 'Odberné miesto'].dropna().astype(str)) if not elec_df.empty else set()
    for parent, children in OM_GROUPS.items():
        if parent not in all_possible_oms:
            continue
        visible_children = [x for x in children if x in all_possible_oms]
        # Ak podružné merania v dátach nie sú, rodič funguje ako samostatné OM.
        if parent in poprad_oms:
            st.sidebar.checkbox(
                f"⚡ {parent}", key=_om_key('grp', parent),
                on_change=_set_group_selection, args=(parent, visible_children)
            )
            if visible_children:
                for child in visible_children:
                    st.sidebar.checkbox(
                        f"　↳ {child}", key=_om_key('om', child),
                        on_change=_set_child_selection, args=(child, parent, visible_children)
                    )
                chosen_children = sum(x in st.session_state.selected_oms for x in visible_children)
                if 0 < chosen_children < len(visible_children):
                    st.sidebar.caption(f"Čiastočný výber: {chosen_children}/{len(visible_children)} podružných meraní")

    grouped_names = {parent for parent in OM_GROUPS}
    grouped_children = {child for children in OM_GROUPS.values() for child in children}
    standalone_poprad = sorted(x for x in poprad_oms if x not in grouped_names and x not in grouped_children)
    for om in standalone_poprad:
        st.sidebar.checkbox(f"• {om}", key=_om_key('om', om), on_change=lambda name=om: _sync_single_om(name))

    bratislava_oms = sorted(set(elec_df.loc[elec_df['Lokalita'] == 'Bratislava', 'Odberné miesto'].dropna().astype(str))) if not elec_df.empty else []
    if bratislava_oms:
        st.sidebar.markdown("**📍 Bratislava**")
        for om in bratislava_oms:
            st.sidebar.checkbox(f"• {om}", key=_om_key('om', om), on_change=lambda name=om: _sync_single_om(name))
else:
    st.sidebar.info("Žiadne dáta nie sú k dispozícii.")

st.sidebar.divider()
st.sidebar.subheader("🔍 Časové filtre")
month_order = ['Január', 'Február', 'Marec', 'Apríl', 'Máj', 'Jún', 'Júl', 'August', 'September', 'Október', 'November', 'December']
month_to_num = {name: i + 1 for i, name in enumerate(month_order)}
num_to_month = {v: k for k, v in month_to_num.items()}
months_in_data = set(elec_df['Mesiac'].dropna().astype(str)) if not elec_df.empty else set()
if not fve_df.empty:
    months_in_data.update(fve_df['Mesiac'].dropna().astype(str))
available_month_nums = [month_to_num[m] for m in month_order if m in months_in_data]
if not available_month_nums:
    available_month_nums = list(range(1, 13))

# Číselné hodnoty 1..12 sú stabilné; používateľovi sa zobrazujú slovenské názvy.
if 'global_month_filter' not in st.session_state:
    st.session_state.global_month_filter = list(available_month_nums)
else:
    st.session_state.global_month_filter = [m for m in st.session_state.global_month_filter if m in available_month_nums]
selected_month_nums = st.sidebar.multiselect(
    "📅 Vyber mesiace (2026)",
    options=available_month_nums,
    format_func=lambda m: num_to_month[m],
    key="global_month_filter"
)
selected_months = [num_to_month[m] for m in selected_month_nums]

st.sidebar.divider()

# --- MANAŽÉRSKE REPORTY V ĽAVOM PANELI (SIDEBAR) ---
st.sidebar.subheader("📥 Manažérske Reporty & Exporty")

# Filtrovanie dát pre reporty vopred
rep_elec = elec_df[elec_df['Mesiac'].isin(selected_months)].copy() if not elec_df.empty else pd.DataFrame()
if not rep_elec.empty:
    rep_elec = rep_elec[rep_elec['Odberné miesto'].isin(st.session_state.selected_oms)]

# 1. Excel export celého portfólia
if not rep_elec.empty:
    output = io.BytesIO()
    export_cols = [c for c in [
        'Mesiac', 'Lokalita', 'Odberné miesto', 'Spotreba (kWh)',
        'Náklady (€)', 'Cena (€/kWh)', 'Zdroj', 'Číslo faktúry (VS)',
        'Fakturačné obdobie'
    ] if c in rep_elec.columns]
    export_detail = rep_elec[export_cols].copy()
    export_summary = (export_detail.groupby(['Lokalita', 'Odberné miesto'], dropna=False)
                      .agg(**{'Spotreba spolu (kWh)': ('Spotreba (kWh)', 'sum'),
                              'Náklady spolu (€)': ('Náklady (€)', 'sum')})
                      .reset_index())
    export_summary['Priemerná cena (€/kWh)'] = export_summary.apply(
        lambda r: r['Náklady spolu (€)'] / r['Spotreba spolu (kWh)'] if r['Spotreba spolu (kWh)'] else 0,
        axis=1
    )
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        export_summary.to_excel(writer, index=False, sheet_name='Manažérsky súhrn', startrow=4)
        export_detail.to_excel(writer, index=False, sheet_name='Mesačné dáta')
        fve_report_export = fve_df[fve_df['Mesiac'].isin(selected_months)].copy() if not fve_df.empty else pd.DataFrame()
        if not fve_report_export.empty:
            fve_report_export.to_excel(writer, index=False, sheet_name='FVE')
        if not pdf_elec_df.empty:
            invoice_export = pdf_elec_df.copy()
            invoice_export = invoice_export[
                invoice_export['Mesiac'].isin(selected_months) &
                invoice_export['Odberné miesto'].isin(st.session_state.selected_oms)
            ]
            invoice_cols = [c for c in [
                'Mesiac', 'Odberné miesto', 'EIC', 'Číslo faktúry (VS)',
                'Fakturačné obdobie', 'Spotreba (kWh)', 'Silová zložka (€)',
                'Distribúcia celkom (€)', 'Prekročenie RK (€)',
                'Prekročenie MRK (€)', 'Jalová elektrina (€)',
                'Prirážka za účinník (€)', 'Spolu bez DPH (€)'
            ] if c in invoice_export.columns]
            invoice_export[invoice_cols].to_excel(writer, index=False, sheet_name='Audit faktúr')

        # Formátovanie pre prezentáciu a ďalšiu prácu v Exceli.
        from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
        from openpyxl.utils import get_column_letter
        wb = writer.book
        ws = wb['Manažérsky súhrn']
        ws['A1'] = 'LUNYS, s.r.o. | ENERGETICKÝ REPORT 2026'
        ws['A1'].font = Font(name='Aptos Display', size=16, bold=True, color='FFFFFF')
        ws['A1'].fill = PatternFill('solid', fgColor='17365D')
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(3, len(export_summary.columns)))
        ws['A2'] = f'Obdobie: {", ".join(selected_months) if selected_months else "bez vybraného mesiaca"} | Vybrané OM: {len(st.session_state.selected_oms)}'
        ws['A2'].font = Font(size=10, italic=True, color='666666')
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=max(3, len(export_summary.columns)))
        ws['A3'] = 'Súhrn spotreby a nákladov podľa odberného miesta. Hlavné trafostanice a podružné merania sa v KPI nesčítavajú duplicitne.'
        ws['A3'].font = Font(size=9, color='555555')
        ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=max(3, len(export_summary.columns)))
        for sheet in wb.worksheets:
            sheet.freeze_panes = 'A2' if sheet.title != 'Manažérsky súhrn' else 'A6'
            sheet.sheet_view.showGridLines = False
            header_row = 5 if sheet.title == 'Manažérsky súhrn' else 1
            for cell in sheet[header_row]:
                cell.font = Font(bold=True, color='FFFFFF')
                cell.fill = PatternFill('solid', fgColor='1F4E78')
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            sheet.auto_filter.ref = f"A{header_row}:{get_column_letter(sheet.max_column)}{sheet.max_row}"
            for col_idx in range(1, sheet.max_column + 1):
                values = [str(sheet.cell(row=r, column=col_idx).value or '') for r in range(1, min(sheet.max_row, 200) + 1)]
                width = min(max(max((len(v) for v in values), default=10) + 2, 12), 34)
                sheet.column_dimensions[get_column_letter(col_idx)].width = width
            for row in sheet.iter_rows(min_row=header_row + 1):
                for cell in row:
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = '#,##0.00'
        ws.row_dimensions[1].height = 28
    excel_data = output.getvalue()

    st.sidebar.download_button(
        label="📊 Stiahnuť manažérsky Excel report",
        data=excel_data,
        file_name="LUNYS_Energeticky_Report_2026.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

# 2. Výber miesta pre špecifický report
if st.session_state.selected_oms and not rep_elec.empty:
    selected_rep_om = st.sidebar.selectbox("Vyber miesto pre detailný audit:", options=st.session_state.selected_oms, key="rep_om_select")
    
    om_rep_df = rep_elec[rep_elec['Odberné miesto'] == selected_rep_om]
    if not om_rep_df.empty:
        om_txt = om_rep_df.to_csv(index=False).encode('utf-8')
        st.sidebar.download_button(
            label=f"📄 Stiahnuť report pre: {selected_rep_om}",
            data=om_txt,
            file_name=f"Audit_{selected_rep_om.replace(' ', '_')}.csv",
            mime="text/csv",
            use_container_width=True
        )

st.sidebar.divider()

st.sidebar.subheader("⚡ Rýchla navigácia")
if st.sidebar.button("📉 Spotreby & Náklady", use_container_width=True, on_click=set_tab, args=(TAB_NAMES[0],)):
    pass
if st.sidebar.button("☀️ Výkon FVE", use_container_width=True, on_click=set_tab, args=(TAB_NAMES[1],)):
    pass
if st.sidebar.button("💶 Cena za 1 kWh", use_container_width=True, on_click=set_tab, args=(TAB_NAMES[2],)):
    pass
if st.sidebar.button("📄 Analýza PDF Faktúr", use_container_width=True, on_click=set_tab, args=(TAB_NAMES[3],)):
    pass

# Filtrovanie dát pre grafy a tabuľky
filtered_fve = fve_df[fve_df['Mesiac'].isin(selected_months)].copy() if not fve_df.empty else pd.DataFrame()
filtered_elec = rep_elec.copy()

if not filtered_fve.empty:
    filtered_fve['Mesiac'] = pd.Categorical(filtered_fve['Mesiac'], categories=month_order, ordered=True)
    filtered_fve = filtered_fve.sort_values('Mesiac')

if not filtered_elec.empty:
    filtered_elec['Mesiac'] = pd.Categorical(filtered_elec['Mesiac'], categories=month_order, ordered=True)
    filtered_elec = filtered_elec.sort_values('Mesiac')

# --- BEZPEČNÝ VÝPOČET KPI SÚHRNOV (ELIMINÁCIA DUPLICITY TRAFA A PODRUŽNÝCH MERANÍ) ---
kpi_df = filtered_elec.copy()
if not kpi_df.empty:
    if "Trafostanica PEBEK" in kpi_df['Odberné miesto'].values:
        kpi_df = kpi_df[~kpi_df['Odberné miesto'].isin(["Skladová hala I.", "Správa 99"])]
    if "Trafostanica HARAMIA" in kpi_df['Odberné miesto'].values:
        kpi_df = kpi_df[~kpi_df['Odberné miesto'].isin(["Bananáreň", "Zemiakareň"])]

total_elec_kwh = kpi_df['Spotreba (kWh)'].sum() if not kpi_df.empty else 0.0
total_elec_eur = kpi_df['Náklady (€)'].sum() if not kpi_df.empty else 0.0
avg_price_per_kwh = (total_elec_eur / total_elec_kwh) if total_elec_kwh > 0 else 0.0

total_fve_kwh = filtered_fve['Vyroba_kWh'].sum() if not filtered_fve.empty else 0.0
total_fve_eur = filtered_fve['Uspora_EUR'].sum() if not filtered_fve.empty else 0.0

st.subheader("📊 Kľúčové ukazovatele 2026 (Z vybraných miest a faktúr)")
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
    st.subheader("Mesačný vývoj spotreby a nákladov")
    if not filtered_elec.empty:
        col_left, col_right = st.columns(2)
        with col_left:
            fig_elec_kwh = px.bar(filtered_elec, x='Mesiac', y='Spotreba (kWh)', color='Odberné miesto', barmode='group', title="Spotreba elektriny (kWh)", template="plotly_dark")
            st.plotly_chart(fig_elec_kwh, use_container_width=True)
        with col_right:
            fig_elec_eur = px.bar(filtered_elec, x='Mesiac', y='Náklady (€)', color='Odberné miesto', barmode='group', title="Mesačné náklady bez DPH (€)", template="plotly_dark")
            st.plotly_chart(fig_elec_eur, use_container_width=True)
            
        st.divider()
        col_sub1, col_sub2 = st.columns(2)
        col_sub1.metric("Celková spotreba (čistý súhrn bez duplicity traf)", f"{kpi_df['Spotreba (kWh)'].sum():,.0f} kWh")
        col_sub2.metric("Celkové náklady (čistý súhrn bez duplicity traf)", f"{kpi_df['Náklady (€)'].sum():,.2f} €")
        st.divider()

        st.dataframe(filtered_elec[['Mesiac', 'Odberné miesto', 'Lokalita', 'Spotreba (kWh)', 'Náklady (€)', 'Cena (€/kWh)', 'Zdroj']], use_container_width=True, hide_index=True)
    else:
        st.info("Pre zvolené filtre nie sú k dispozícii žiadne dáta.")

elif selected_tab == TAB_NAMES[1]:
    st.subheader("☀️ Výkon a úspory FVE")
    
    fve_lokalita_filter = st.radio(
        "📍 Vyber lokalitu pre FVE:",
        options=["Všetky lokality", "Poprad", "Bratislava"],
        horizontal=True,
        key="fve_lok_filter"
    )
    
    tab_fve_df = filtered_fve.copy()
    if fve_lokalita_filter != "Všetky lokality":
        tab_fve_df = tab_fve_df[tab_fve_df['Lokalita_raw'] == fve_lokalita_filter]

    if not tab_fve_df.empty:
        summary_fve = tab_fve_df.groupby('Lokalita_raw').agg({'Vyroba_kWh': 'sum', 'Uspora_EUR': 'sum', 'Podiel_Pokrytia_Pct': 'mean'}).reset_index()

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

        st.dataframe(tab_fve_df, use_container_width=True, hide_index=True)
    else:
        st.info("FVE dáta pre zvolenú lokalitu nie sú k dispozícii.")

elif selected_tab == TAB_NAMES[2]:
    st.subheader("💶 Priemerné náklady na 1 kWh")
    st.markdown("Efektívne náklady bez DPH na odobratú 1 kWh podľa dostupných faktúr a Excel dát. Nejde o burzovú spotovú cenu elektriny.")

    if not filtered_elec.empty:
        fig_spot_trend = go.Figure()
        
        for om in filtered_elec['Odberné miesto'].unique():
            sub_om = filtered_elec[filtered_elec['Odberné miesto'] == om]
            fig_spot_trend.add_trace(go.Scatter(
                x=sub_om['Mesiac'],
                y=sub_om['Cena (€/kWh)'],
                mode='lines+markers',
                name=f"{om} (€/kWh)"
            ))

        fig_spot_trend.update_layout(
            title="Mesačný vývoj ceny za 1 kWh (€/kWh)",
            xaxis_title="Mesiac",
            yaxis_title="Cena (€/kWh)",
            template="plotly_dark",
            hovermode="x unified"
        )
        st.plotly_chart(fig_spot_trend, use_container_width=True)

        st.dataframe(filtered_elec[['Mesiac', 'Odberné miesto', 'Spotreba (kWh)', 'Náklady (€)', 'Cena (€/kWh)', 'Zdroj']], use_container_width=True, hide_index=True)
    else:
        st.info("Pre zvolené filtre nie sú k dispozícii žiadne dáta.")

elif selected_tab == TAB_NAMES[3]:
    st.subheader("📄 Analýza PDF faktúr zo ZSE & Audit")
    
    pdf_only_df = pdf_elec_df.copy()
    if not pdf_only_df.empty:
        if selected_months:
            pdf_only_df = pdf_only_df[pdf_only_df['Mesiac'].isin(selected_months)]
        pdf_only_df = pdf_only_df[pdf_only_df['Odberné miesto'].isin(st.session_state.selected_oms)]

    if not pdf_only_df.empty:
        df_filtered_inv = pdf_only_df.copy()
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

        st.markdown("### 🔎 Detail faktúr podľa mesiacov")
        for idx, row in df_filtered_inv.iterrows():
            st.markdown(f"#### Faktúra č. {row['Číslo faktúry (VS)']} – {row['Odberné miesto']} ({row['Fakturačné obdobie']})")
            
            c_f1, c_f2, c_f3, c_f4 = st.columns(4)
            c_f1.metric("Silová zložka (Komodita)", f"{row['Silová zložka (€)']:,.2f} €")
            c_f2.metric("Distribúcia celkom", f"{row['Distribúcia celkom (€)']:,.2f} €")
            c_f3.metric("Prekročenie RK / MRK", f"{(row['Prekročenie RK (€)'] + row['Prekročenie MRK (€)']):,.2f} €", delta_color="inverse" if (row['Prekročenie RK (€)'] + row['Prekročenie MRK (€)']) > 0 else "normal")
            c_f4.metric("Jalová + Účinník", f"{(row['Jalová elektrina (€)'] + row['Prirážka za účinník (€)']):,.2f} €", delta_color="inverse" if (row['Jalová elektrina (€)'] + row['Prirážka za účinník (€)']) > 0 else "normal")

            with st.expander(f"🔍 Diagnostika surového textu z PDF pre faktúru {row['Číslo faktúry (VS)']}"):
                st.text(row['Raw_Text'])

            st.divider()
    else:
        st.info("Pre vybrané odberné miesta / mesiace nie sú k dispozícii žiadne načítané PDF faktúry.")

    st.divider()
    st.markdown("### 📤 Nahranie ďalších PDF faktúr")
    uploaded_files = st.file_uploader("Pridaj ďalšie PDF faktúry:", type=["pdf"], accept_multiple_files=True, key="uploader_bottom")
    if uploaded_files:
        for uploaded_file in uploaded_files:
            file_path = os.path.join(INVOICE_DIR, uploaded_file.name)
            with open(file_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
        st.success(f"Úspešne uložených {len(uploaded_files)} súborov do `{INVOICE_DIR}`. Obnov stránku.")
