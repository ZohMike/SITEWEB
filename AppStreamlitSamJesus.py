import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO
import os
import tempfile
from fpdf import FPDF
from datetime import datetime
import unicodedata
import re
import numpy as np

# Dictionnaire pour traduire les mois en français
mois_fr = {
    "January": "Janvier", "February": "Février", "March": "Mars", "April": "Avril",
    "May": "Mai", "June": "Juin", "July": "Juillet", "August": "Août",
    "September": "Septembre", "October": "Octobre", "November": "Novembre", "December": "Décembre"
}

def format_date_fr(date):
    return date.strftime("%B %Y").replace(date.strftime("%B"), mois_fr.get(date.strftime("%B"), date.strftime("%B")))

# Fonction pour extraire intelligemment un nombre adapté de mots d'un nom de client
def extract_client_words(client, max_words=4, max_chars=30):
    """
    Extrait un nombre adapté de mots d'un nom de client en fonction de la longueur des mots.
    Les mots liés par une apostrophe (ex. "d'ALMEIDA") sont considérés comme un seul mot.
    
    Args:
        client (str): Nom du client.
        max_words (int): Nombre maximum initial de mots à extraire (par défaut 4).
        max_chars (int): Nombre approximatif maximum de caractères souhaité (par défaut 30).
    
    Returns:
        str: Chaîne contenant un nombre adapté de mots, ou "(vide)" si l'entrée est vide.
    
    Exemple:
        >>> extract_client_words("Société d'ALMEIDA Jean Baptiste KOUADIO SARL")
        'Société d'ALMEIDA Jean'
        
        >>> extract_client_words("ETABLISSEMENTS COMMERCIAUX PHARMACEUTIQUES INTERNATIONAUX")
        'ETABLISSEMENTS COMMERCIAUX'
    """
    if not client or not client.strip():
        return "(vide)"
    
    # Normaliser les caractères Unicode (ex. ' → ')
    normalized_client = unicodedata.normalize('NFKD', client).strip()
    # Splitter uniquement sur les espaces, préserver les apostrophes
    client_words = re.split(r'\s+', normalized_client)
    # Filtrer les mots non vides
    client_words = [word for word in client_words if word]
    
    # Logique adaptative pour déterminer le nombre optimal de mots
    actual_max_words = max_words
    
    # Si les 2 premiers mots sont déjà très longs (plus de max_chars/2 caractères)
    if len(" ".join(client_words[:2])) > max_chars/2:
        actual_max_words = 2
    # Si les 3 premiers mots sont déjà longs (plus de 2*max_chars/3 caractères)
    elif len(" ".join(client_words[:3])) > 2*max_chars/3:
        actual_max_words = 3
    # Pour les cas extrêmes où même un seul mot est très long
    elif client_words and len(client_words[0]) > max_chars - 5:
        # Tronquer le premier mot s'il est extrêmement long
        return client_words[0][:max_chars-3] + "..."
        
    # Prendre les actual_max_words premiers mots et joindre avec des espaces
    result = " ".join(client_words[:actual_max_words])
    
    # Si le résultat reste trop long, ajouter des points de suspension
    if len(result) > max_chars:
        result = result[:max_chars-3] + "..."
        
    return result

# Fonction pour traiter les noms d'assureurs trop longs
def extract_assureur_words(assureur, max_words=3, max_chars=25):
    """
    Extrait un nombre adapté de mots d'un nom d'assureur en fonction de la longueur des mots.
    
    Args:
        assureur (str): Nom de l'assureur.
        max_words (int): Nombre maximum initial de mots à extraire (par défaut 3).
        max_chars (int): Nombre approximatif maximum de caractères souhaité (par défaut 25).
    
    Returns:
        str: Chaîne contenant un nombre adapté de mots, ou "(vide)" si l'entrée est vide.
    
    Exemple:
        >>> extract_assureur_words("COMPAGNIE NATIONALE D'ASSURANCE AFRICAINE")
        'COMPAGNIE NATIONALE D'ASSURANCE'
    """
    if not assureur or not assureur.strip():
        return "(vide)"
    
    # Normaliser les caractères Unicode
    normalized_assureur = unicodedata.normalize('NFKD', assureur).strip()
    # Splitter uniquement sur les espaces, préserver les apostrophes
    assureur_words = re.split(r'\s+', normalized_assureur)
    # Filtrer les mots non vides
    assureur_words = [word for word in assureur_words if word]
    
    # Logique adaptative pour déterminer le nombre optimal de mots
    actual_max_words = max_words
    
    # Si les 2 premiers mots sont déjà très longs
    if len(" ".join(assureur_words[:2])) > max_chars/2:
        actual_max_words = 2
    # Pour les cas extrêmes où même un seul mot est très long
    elif assureur_words and len(assureur_words[0]) > max_chars - 5:
        # Tronquer le premier mot s'il est extrêmement long
        return assureur_words[0][:max_chars-3] + "..."
        
    # Prendre les actual_max_words premiers mots et joindre avec des espaces
    result = " ".join(assureur_words[:actual_max_words])
    
    # Si le résultat reste trop long, ajouter des points de suspension
    if len(result) > max_chars:
        result = result[:max_chars-3] + "..."
        
    return result

# Chemins temporaires pour graphiques et logos
graph_path = None
evol_effectif_path = None
evol_mensuel_path = None
conso_benef_path = None
logo_ankara_path = None
logo_assureur_path = None

# Configuration de la page
st.set_page_config(page_title="Générateur de Rapport Santé", layout="wide", initial_sidebar_state="collapsed")

# CSS personnalisé moderne
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700&display=swap');
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@500;600;700&display=swap');
    
    * {
        font-family: 'Poppins', sans-serif;
        transition: all 0.3s ease;
    }
    
    :root {
        --primary-color: #06A77D;
        --primary-light: #3DDCAE;
        --primary-dark: #057156;
        --secondary-color: #F58634;
        --secondary-light: #FFAB72;
        --secondary-dark: #D36A1B;
        --bg-light: #F8FBFF;
        --bg-white: #FFFFFF;
        --text-dark: #2C3E50;
        --text-muted: #7F8C8D;
        --border-color: #E6EEF8;
        --success: #2ECC71;
        --warning: #F1C40F;
        --error: #E74C3C;
        --info: #3498DB;
    }
    
    body {
        background: var(--bg-light);
        color: var(--text-dark);
    }
    
    /* Main container styling */
    .stApp {
        max-width: 1300px;
        margin: 0 auto;
        background: linear-gradient(135deg, var(--bg-light) 0%, #F0F8FF 100%);
    }
    
    /* Header styling with modern gradient */
    header {
        background: linear-gradient(120deg, var(--primary-color) 0%, var(--primary-dark) 100%);
        padding: 2rem 0;
        margin-bottom: 2rem;
        border-radius: 0 0 20px 20px;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.1);
    }
    
    /* Improve container styling */
    section.main, div[data-testid="stVerticalBlock"] {
        background-color: var(--bg-white);
        border-radius: 16px;
        box-shadow: 0 8px 30px rgba(0, 0, 0, 0.08);
        padding: 25px;
        margin-bottom: 2rem;
        border: 1px solid var(--border-color);
        transition: transform 0.3s ease, box-shadow 0.3s ease;
    }
    
    div[data-testid="stVerticalBlock"]:hover {
        transform: translateY(-3px);
        box-shadow: 0 12px 40px rgba(0, 0, 0, 0.12);
    }
    
    /* Modern title */
    h1 {
        font-family: 'Montserrat', sans-serif;
        font-size: 2.6rem;
        font-weight: 700;
        color: var(--primary-color);
        text-align: center;
        margin-bottom: 2.5rem;
        letter-spacing: -0.5px;
        position: relative;
        padding-bottom: 1rem;
    }
    
    h1:after {
        content: '';
        position: absolute;
        bottom: 0;
        left: 50%;
        transform: translateX(-50%);
        width: 150px;
        height: 4px;
        background: linear-gradient(90deg, var(--primary-color), var(--secondary-color));
        border-radius: 10px;
    }
    
    /* Section headers */
    h2 {
        font-family: 'Montserrat', sans-serif;
        font-size: 1.6rem;
        font-weight: 600;
        color: var(--primary-color);
        margin-bottom: 1.5rem;
        padding-bottom: 8px;
        position: relative;
        display: inline-block;
    }
    
    h2:after {
        content: '';
        position: absolute;
        bottom: 0;
        left: 0;
        width: 100%;
        height: 3px;
        background: linear-gradient(90deg, var(--secondary-color), var(--secondary-light));
        border-radius: 10px;
    }
    
    /* File uploader styling */
    .stFileUploader {
        background-color: var(--bg-white);
        border: 2px dashed var(--primary-light);
        border-radius: 12px;
        padding: 15px;
        margin-bottom: 1.5rem;
        color: var(--text-dark);
        text-align: center;
        transition: all 0.3s ease;
    }
    
    .stFileUploader:hover {
        border-color: var(--primary-color);
        background-color: rgba(6, 167, 125, 0.05);
    }
    
    /* Form controls styling */
    .stSelectbox, .stTextInput, .stNumberInput {
        background-color: var(--bg-white);
        border: 2px solid var(--border-color);
        border-radius: 12px;
        padding: 14px;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.03);
        margin-bottom: 1.5rem;
        color: var(--text-dark);
        transition: all 0.3s ease;
    }
    
    .stSelectbox:hover, .stTextInput:hover, .stNumberInput:hover {
        border-color: var(--primary-light);
        box-shadow: 0 6px 20px rgba(0, 0, 0, 0.08);
    }
    
    .stSelectbox:focus, .stTextInput:focus, .stNumberInput:focus {
        border-color: var(--primary-color);
        box-shadow: 0 6px 25px rgba(6, 167, 125, 0.15);
    }
    
    /* Main action button */
    .stButton>button {
        background: linear-gradient(140deg, var(--primary-color) 0%, var(--primary-dark) 100%);
        color: white;
        border: none;
        border-radius: 12px;
        padding: 0.9rem 2.2rem;
        font-weight: 600;
        font-size: 1.05rem;
        letter-spacing: 0.3px;
        box-shadow: 0 5px 15px rgba(6, 167, 125, 0.25);
        transition: all 0.3s ease;
        text-transform: uppercase;
    }
    
    .stButton>button:hover {
        background: linear-gradient(140deg, var(--primary-color) 20%, var(--primary-dark) 100%);
        transform: translateY(-3px);
        box-shadow: 0 8px 25px rgba(6, 167, 125, 0.35);
    }
    
    .stButton>button:active {
        transform: translateY(0);
    }
    
    /* Download button styling */
    .stDownloadButton>button {
        background: linear-gradient(140deg, var(--secondary-color) 0%, var(--secondary-dark) 100%);
        color: white;
        border: none;
        border-radius: 12px;
        padding: 0.9rem 2.2rem;
        font-weight: 600;
        font-size: 1.05rem;
        letter-spacing: 0.3px;
        box-shadow: 0 5px 15px rgba(245, 134, 52, 0.25);
        transition: all 0.3s ease;
    }
    
    .stDownloadButton>button:hover {
        background: linear-gradient(140deg, var(--secondary-color) 20%, var(--secondary-dark) 100%);
        transform: translateY(-3px);
        box-shadow: 0 8px 25px rgba(245, 134, 52, 0.35);
    }
    
    /* DataFrames styling */
    .stDataFrame {
        border: 1px solid var(--border-color);
        border-radius: 12px;
        overflow: hidden;
        box-shadow: 0 5px 20px rgba(0, 0, 0, 0.05);
        background-color: var(--bg-white);
        color: var(--text-dark);
    }
    
    .stDataFrame table {
        width: 100%;
        border-collapse: separate;
        border-spacing: 0;
    }
    
    .stDataFrame th {
        background: linear-gradient(140deg, var(--primary-color) 0%, var(--primary-dark) 100%);
        color: white;
        font-weight: 600;
        padding: 12px;
        text-transform: uppercase;
        font-size: 0.85rem;
        letter-spacing: 0.5px;
    }
    
    .stDataFrame td {
        padding: 10px;
        border-bottom: 1px solid var(--border-color);
    }
    
    .stDataFrame tr:last-child td {
        border-bottom: none;
    }
    
    .stDataFrame tr:nth-child(even) {
        background-color: rgba(6, 167, 125, 0.03);
    }
    
    /* Notification styling */
    .stError, .stWarning, .stInfo, .stSuccess {
        border-radius: 12px;
        padding: 16px;
        margin-bottom: 1.5rem;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.05);
        display: flex;
        align-items: center;
        border-left: 5px solid;
    }
    
    .stError {
        background-color: rgba(231, 76, 60, 0.08);
        color: var(--error);
        border-left-color: var(--error);
    }
    
    .stWarning {
        background-color: rgba(241, 196, 15, 0.08);
        color: var(--warning);
        border-left-color: var(--warning);
    }
    
    .stInfo {
        background-color: rgba(52, 152, 219, 0.08);
        color: var(--info);
        border-left-color: var(--info);
    }
    
    .stSuccess {
        background-color: rgba(46, 204, 113, 0.08);
        color: var(--success);
        border-left-color: var(--success);
    }
    
    /* Responsive adjustments */
    @media (max-width: 768px) {
        .stApp {
            padding: 15px;
        }
        
        h1 {
            font-size: 2.2rem;
        }
        
        h2 {
            font-size: 1.4rem;
        }
        
        .stButton>button, .stDownloadButton>button {
            width: 100%;
            padding: 0.85rem;
        }
        
        section.main, div[data-testid="stVerticalBlock"] {
            padding: 15px;
        }
    }
    </style>
    
    <!-- Custom Header for a more polished look -->
    <header>
        <div style="text-align: center; padding: 0 20px;">
            <h1 style="color: white; margin-bottom: 10px; font-size: 2.8rem;">Générateur de Rapport Santé</h1>
            <p style="color: rgba(255,255,255,0.85); font-size: 1.1rem; max-width: 800px; margin: 0 auto;">
                Créez des rapports statistiques détaillés pour analyser les performances des contrats d'assurance santé
            </p>
        </div>
    </header>
""", unsafe_allow_html=True)

# Titre principal
st.title("Ankara Services")

# Section 1 : Informations du contrat
st.subheader("Informations du contrat")
with st.container():
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        fichier_detail = st.file_uploader("DETAIL.xlsx", type="xlsx")
    with col2:
        fichier_production = st.file_uploader("PRODUCTION.xlsx", type="xlsx")
    with col3:
        fichier_effectif = st.file_uploader("EFFECTIF.xlsx", type="xlsx")
    with col4:
        fichier_clause = st.file_uploader("Clause Ajustement Santé.xlsx", type="xlsx")

df_detail = None
nom_assureur = ""
client = ""
police_assureur = ""
police_ankara = ""
periode = ""
prime_nette = 0.0
prime_acquise = 0.0
df_clause = None
df_production = None

# Placeholder global pour la période qui sera mise à jour dynamiquement
periode_placeholder = st.empty()

if fichier_detail:
    try:
        xls = pd.ExcelFile(fichier_detail)
        if "DETAIL" in xls.sheet_names:
            df_detail = xls.parse("DETAIL")
            df_detail.iloc[:, 5] = df_detail.iloc[:, 5].astype(str).str.strip().str.upper().str.replace(r'\s+', ' ', regex=True)
            df_detail.iloc[:, 27] = df_detail.iloc[:, 27].astype(str).str.strip().str.upper().str.replace(r'\s+', ' ', regex=True)
            
            # Créer une clé unique combinant client et numéro de police
            df_detail['client_police_key'] = df_detail[df_detail.columns[5]] + " | " + df_detail[df_detail.columns[6]]
            clients = df_detail[df_detail.columns[5]].dropna().unique().tolist()
            polices_dict = df_detail.groupby(df_detail.columns[5])[df_detail.columns[6]].unique().apply(list).to_dict()
            assureurs = df_detail[df_detail.columns[27]].dropna().unique().tolist()

            with st.container():
                col1, col2 = st.columns(2)
                with col1:
                    nom_assureur = st.selectbox("Nom de l'assureur", options=assureurs)
                    client = st.selectbox("Client", options=clients)
                    polices = polices_dict.get(client, [])
                    police_ankara = st.selectbox("N° Police Ankara", options=polices)
                with col2:
                    police_assureur = st.text_input("N° Police Assureur")
                    # Placeholder pour la période qui sera mise à jour plus tard
                    periode_placeholder.text_input("Période concernée", value="", disabled=True)
                    periode = ""
        else:
            st.error("❌ Le fichier DETAIL.xlsx ne contient pas de feuille 'DETAIL'.")
    except Exception as e:
        st.error(f"❌ Erreur lors du chargement du fichier DETAIL : {e}")

# Charger et filtrer le fichier PRODUCTION.xlsx
if fichier_production:
    try:
        df_production = pd.read_excel(fichier_production)
        expected_columns = ["Id Police Ankara", "N° Police Assureur", "Assureur", "Client", 
                           "Primes Émises Nettes", "Primes Acquises", "Sinistres", "S/P"]
        if not all(col in df_production.columns for col in expected_columns):
            st.error("❌ Le fichier PRODUCTION.xlsx ne contient pas toutes les colonnes attendues : " + ", ".join(expected_columns))
        else:
            df_production["Assureur"] = df_production["Assureur"].astype(str).str.strip().str.upper().str.replace(r'\s+', ' ', regex=True)
            df_production["Client"] = df_production["Client"].astype(str).str.strip().str.upper().str.replace(r'\s+', ' ', regex=True)
            df_production['client_police_key'] = df_production["Client"] + " | " + df_production["Id Police Ankara"]

            df_production_filtered = df_production[
                (df_production["Assureur"] == nom_assureur) & 
                (df_production['client_police_key'] == client + " | " + police_ankara)
            ]
            if df_production_filtered.empty:
                st.warning("⚠️ Aucune donnée dans PRODUCTION.xlsx pour l'assureur, le client et la police sélectionnés.")
                prime_nette = 0.0
                prime_acquise = 0.0
            else:
                prime_nette = df_production_filtered["Primes Émises Nettes"].iloc[0]
                prime_acquise = df_production_filtered["Primes Acquises"].iloc[0]
                try:
                    prime_nette = float(prime_nette)
                    prime_acquise = float(prime_acquise)
                except (ValueError, TypeError):
                    st.warning("⚠️ Impossible de convertir les primes en nombres. Valeurs définies à 0.")
                    prime_nette = 0.0
                    prime_acquise = 0.0
    except Exception as e:
        st.error(f"❌ Erreur lors du chargement du fichier PRODUCTION : {e}")
        prime_nette = 0.0
        prime_acquise = 0.0

# Charger le fichier Clause Ajustement Santé
if fichier_clause:
    try:
        df_clause = pd.read_excel(fichier_clause)
        df_clause.columns = [c.strip().lower() for c in df_clause.columns]
        possible_min_cols = ['tranche min', 'minimum', 'min', 'tranche_min', 'rapport s/p min']
        possible_max_cols = ['tranche max', 'maximum', 'max', 'tranche_max', 'rapport s/p max']
        tranche_min_col = next((col for col in df_clause.columns if col in possible_min_cols), None)
        tranche_max_col = next((col for col in df_clause.columns if col in possible_max_cols), None)
        if not tranche_min_col or not tranche_max_col:
            st.warning("⚠️ Les colonnes 'Rapport S/P min' ou 'Rapport S/P max' (ou équivalentes) sont introuvables dans le fichier Clause Ajustement Santé.")
        else:
            if tranche_min_col:
                df_clause[tranche_min_col] = df_clause[tranche_min_col].apply(lambda x: f"{float(x)*100:.0f}%" if pd.notna(x) else x)
            if tranche_max_col:
                df_clause[tranche_max_col] = df_clause[tranche_max_col].apply(lambda x: f"{float(x)*100:.0f}%" if pd.notna(x) else x)
            st.success("✅ Colonnes 'Rapport S/P min' et 'Rapport S/P max' détectées et converties en pourcentages.")
    except Exception as e:
        st.error(f"❌ Erreur lors du chargement du fichier Clause Ajustement Santé : {e}")
        df_clause = None

# Section 2 : Filtrage et sinistralité
df_filtre = None
df_effectif = None
sinistralite_ok = False

if df_detail is not None:
    try:
        colonne_client = df_detail.columns[5]
        colonne_police = df_detail.columns[6]
        df_filtre = df_detail[
            (df_detail[colonne_client] == client) & 
            (df_detail[colonne_police] == police_ankara)
        ]

        if df_filtre is not None and not df_filtre.empty:
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                df_filtre.to_excel(writer, index=False, sheet_name="DETAIL_FILTRÉ")
            buffer.seek(0)
            st.download_button("Télécharger DETAIL filtré", buffer.getvalue(), file_name="DETAIL_filtre.xlsx")

            montant_sinistres = df_filtre.iloc[:, 22].sum()
            ratio_sp = montant_sinistres / prime_acquise if prime_acquise > 0 else 0
            # Extraire intelligemment les noms du client et de l'assureur
            client_short = extract_client_words(client, max_words=4)
            assureur_short = extract_assureur_words(nom_assureur, max_words=3)
            df_sin = pd.DataFrame([{
                "Id Police Ankara": police_ankara,
                "N° Police Assureur": police_assureur or "(vide)",
                "Assureur": assureur_short,
                "Client": client_short,
                "Primes Émises Nettes": f"{prime_nette:,.0f}".replace(",", " "),
                "Primes Acquises": f"{prime_acquise:,.0f}".replace(",", " "),
                "Sinistres": f"{montant_sinistres:,.0f}".replace(",", " "),
                "S/P": f"{ratio_sp:.0%}"
            }])
            df_sin = df_sin[["Id Police Ankara", "N° Police Assureur", "Assureur", "Client", "Primes Émises Nettes", "Primes Acquises", "Sinistres", "S/P"]]
            st.markdown("## I - Sinistralité")
            
            st.markdown("""
                <style>
                .stDataFrame table td, .stDataFrame table th {
                    white-space: normal !important;
                    word-wrap: break-word !important;
                    text-align: center !important;
                    overflow-wrap: break-word !important;
                    max-width: 100% !important;
                    font-size: 12px !important;
                }
                </style>
            """, unsafe_allow_html=True)

            column_config = {
                "Id Police Ankara": st.column_config.TextColumn(width=120),
                "N° Police Assureur": st.column_config.TextColumn(width=120),
                "Assureur": st.column_config.TextColumn(width=200),
                "Client": st.column_config.TextColumn(width=450),
                "Primes Émises Nettes": st.column_config.TextColumn(width=90),
                "Primes Acquises": st.column_config.TextColumn(width=90),
                "Sinistres": st.column_config.TextColumn(width=90),
                "S/P": st.column_config.TextColumn(width=50)
            }
            st.dataframe(df_sin, column_config=column_config, use_container_width=True)

            if df_clause is not None:
                st.markdown("### Clause Ajustement Santé")
                ratio_sp_value = float(df_sin["S/P"].iloc[0].replace("%", "")) / 100
                df_clause_styled = df_clause.copy()
                highlight_row = None
                if tranche_min_col and tranche_max_col:
                    ratio_sp_rounded = round(ratio_sp_value * 100) / 100
                    for idx, row in df_clause_styled.iterrows():
                        try:
                            tranche_min = float(str(row[tranche_min_col]).replace('%', '')) / 100
                            tranche_max = float(str(row[tranche_max_col]).replace('%', '')) / 100
                            if tranche_min <= ratio_sp_rounded <= tranche_max:
                                highlight_row = idx
                                break
                        except (ValueError, TypeError):
                            continue
                if highlight_row is not None:
                    def highlight_row_func(s):
                        return ['background-color: #f77f00' if s.name == highlight_row else '' for _ in s]
                    st.dataframe(df_clause_styled.style.apply(highlight_row_func, axis=1))
                else:
                    st.dataframe(df_clause_styled)

            sinistralite_ok = True
        else:
            st.warning("⚠️ Aucun résultat après filtrage.")
    except Exception as e:
        st.error(f"❌ Erreur lors du filtrage : {e}")
else:
    st.warning("⚠️ Chargez un fichier DETAIL.xlsx valide.")

# Section 3 : Évolution des effectifs
if sinistralite_ok and fichier_effectif:
    st.markdown("## II - Évolution des effectifs")
    try:
        df_effectif = pd.read_excel(fichier_effectif)
        df_effectif.columns = [c.strip().upper() for c in df_effectif.columns]
        required_columns = ['MOIS', 'ASSUREUR', 'CLIENT', 'ADHERENT', 'CONJOINT', 'ENFANT', 'TOTAL']
        missing_columns = [col for col in required_columns if col not in df_effectif.columns]
        if missing_columns:
            st.error(f"❌ Les colonnes suivantes sont manquantes dans le fichier EFFECTIF.xlsx : {', '.join(missing_columns)}")
        else:
            # Renommer les colonnes après vérification
            df_effectif = df_effectif.rename(columns={
                'CONJOINT': 'CONJOINTS',
                'ENFANT': 'ENFANTS'
            })
            df_effectif['ASSUREUR'] = df_effectif['ASSUREUR'].astype(str).str.strip().str.upper().str.replace(r'\s+', ' ', regex=True)
            df_effectif['CLIENT'] = df_effectif['CLIENT'].astype(str).str.strip().str.upper().str.replace(r'\s+', ' ', regex=True)
            df_effectif_filtered = df_effectif[
                (df_effectif['ASSUREUR'] == nom_assureur) &
                (df_effectif['CLIENT'] == client)
            ]
            if df_effectif_filtered.empty:
                st.warning("⚠️ Aucune donnée dans EFFECTIF.xlsx pour l'assureur et le client sélectionnés.")
                # Mettre à jour le placeholder avec une période vide si pas de données d'effectifs
                periode_placeholder.text_input("Période concernée", value="", disabled=True)
                periode = ""
            else:
                df_effectif_filtered["MOIS"] = pd.to_datetime(df_effectif_filtered["MOIS"], format="%d/%m/%Y", errors="coerce")
                df_effectif_filtered = df_effectif_filtered.sort_values(by="MOIS", ascending=True)
                
                # Calculer la période à partir des données d'effectifs pour l'affichage UI
                try:
                    raw_dates = df_effectif_filtered["MOIS"].copy()
                    if not raw_dates.isna().all():
                        date_min = raw_dates.min()
                        date_max = raw_dates.max()
                        mois_min = mois_fr.get(date_min.strftime("%B"), date_min.strftime("%B"))
                        mois_max = mois_fr.get(date_max.strftime("%B"), date_max.strftime("%B"))
                        annee_min = date_min.strftime("%Y")
                        annee_max = date_max.strftime("%Y")
                        
                        if annee_min == annee_max:
                            periode = f"de {mois_min} à {mois_max} {annee_max}"
                        else:
                            periode = f"de {mois_min} {annee_min} à {mois_max} {annee_max}"
                        
                        # Mettre à jour le placeholder de la période dans l'interface
                        periode_placeholder.text_input("Période concernée", value=periode, disabled=True)
                except:
                    periode = ""
                    periode_placeholder.text_input("Période concernée", value="", disabled=True)
                
                df_effectif_filtered["MOIS"] = df_effectif_filtered["MOIS"].apply(format_date_fr)
                display_columns = ["MOIS", "ADHERENT", "CONJOINTS", "ENFANTS", "TOTAL"]
                df_effectif_display = df_effectif_filtered[display_columns]
                st.dataframe(df_effectif_display)
                fig, ax = plt.subplots(figsize=(10, 5))
                colors = ['#279244', '#f77f00', '#ff6f61']
                for i, col in enumerate(["ADHERENT", "CONJOINTS", "ENFANTS", "TOTAL"]):
                    if col in df_effectif_filtered.columns:
                        ax.plot(df_effectif_filtered["MOIS"], df_effectif_filtered[col], marker='o', label=col.title(), color=colors[i % len(colors)])
                ax.set_title("Évolution des effectifs")
                ax.set_xlabel("Mois")
                ax.set_ylabel("Effectifs")
                ax.legend()
                plt.xticks(rotation=45)
                st.pyplot(fig)
                evol_effectif_path = os.path.join(tempfile.gettempdir(), "graph_effectif.png")
                fig.savefig(evol_effectif_path, bbox_inches='tight')
                plt.close(fig)
                df_effectif = df_effectif_filtered
    except Exception as e:
        st.error(f"❌ Erreur lors du chargement des effectifs : {e}")
        # Mettre à jour le placeholder avec une période vide en cas d'erreur
        periode_placeholder.text_input("Période concernée", value="", disabled=True)
        periode = ""

# Section 4 : Consommation par type de bénéficiaire
if sinistralite_ok and df_effectif is not None and df_filtre is not None and not df_filtre.empty:
    st.markdown("## III - Consommation par type de bénéficiaire")
    try:
        col_carte = df_filtre.columns[9]
        col_filiation = df_filtre.columns[10]
        col_montant = df_filtre.columns[22]
        mapping = {"ADHERENT": "ASSURÉ PRINCIPAL", "ASSURE PRINCIPAL": "ASSURÉ PRINCIPAL", "ASSURÉ PRINCIPAL": "ASSURÉ PRINCIPAL",
                   "assure principal": "ASSURÉ PRINCIPAL", "CONJOINT": "CONJOINT", "conjoint": "CONJOINT", "ENFANT": "ENFANT", "enfant": "ENFANT"}
        df_filtre["FILIATION"] = df_filtre[col_filiation].map(mapping)
        patients_uniques = df_filtre.drop_duplicates(subset=[col_carte])[[col_carte, "FILIATION"]]
        patients_counts = patients_uniques["FILIATION"].value_counts().rename("Nombre de patients")
        required_columns = ["ADHERENT", "CONJOINTS", "ENFANTS"]
        missing_columns = [col for col in required_columns if col not in df_effectif.columns]
        if missing_columns:
            st.error(f"❌ Les colonnes suivantes sont manquantes dans le fichier EFFECTIF.xlsx : {', '.join(missing_columns)}")
        else:
            effectifs = pd.Series({
                "ASSURÉ PRINCIPAL": df_effectif["ADHERENT"].max(),
                "CONJOINT": df_effectif["CONJOINTS"].max(),
                "ENFANT": df_effectif["ENFANTS"].max()
            }).rename("Effectif Total")
            
            # Remplacer les valeurs NaN par 0 dans les effectifs
            effectifs = effectifs.fillna(0)
            
            montants = df_filtre.groupby("FILIATION")[col_montant].sum().rename("Montant couvert")
            tableau = pd.concat([patients_counts, effectifs, montants], axis=1)
            
            # Gérer la division par zéro pour le taux d'utilisation
            tableau["Taux d'utilisation"] = tableau.apply(
                lambda row: row["Nombre de patients"] / row["Effectif Total"] if row["Effectif Total"] > 0 else 0, 
                axis=1
            )
            
            total_montant = tableau["Montant couvert"].sum()
            tableau["Part de consommation"] = tableau["Montant couvert"] / total_montant if total_montant > 0 else 0
            
            total = pd.DataFrame({
                "Nombre de patients": [tableau["Nombre de patients"].sum()],
                "Effectif Total": [tableau["Effectif Total"].sum()],
                "Taux d'utilisation": [tableau["Nombre de patients"].sum() / tableau["Effectif Total"].sum() if tableau["Effectif Total"].sum() > 0 else 0],
                "Montant couvert": [total_montant],
                "Part de consommation": [1.0]
            }, index=["Total général"])
            tableau_final = pd.concat([tableau, total])
            cols = ["Nombre de patients", "Effectif Total", "Taux d'utilisation", "Montant couvert", "Part de consommation"]
            ordre_filiation = ["ASSURÉ PRINCIPAL", "CONJOINT", "ENFANT", "Total général"]
            tableau_final = tableau_final.reindex(ordre_filiation)[cols]
            
            # Remplacer les valeurs non-finies avant conversion
            tableau_final = tableau_final.replace([np.inf, -np.inf], 0)
            tableau_final = tableau_final.fillna(0)
            
            # Conversions sécurisées
            tableau_final["Taux d'utilisation"] = (tableau_final["Taux d'utilisation"] * 100).round(0).astype(int).astype(str) + "%"
            tableau_final["Nombre de patients"] = tableau_final["Nombre de patients"].round(0).astype(int)
            tableau_final["Part de consommation"] = (tableau_final["Part de consommation"] * 100).round(0).astype(int).astype(str) + "%"
            tableau_final["Montant couvert"] = tableau_final["Montant couvert"].apply(lambda x: f"{int(x):,}".replace(",", " "))
            
            st.dataframe(tableau_final)
            df_graph_benef = tableau_final[tableau_final.index != "Total général"].copy()
            df_graph_benef["Montant couvert"] = df_graph_benef["Montant couvert"].str.replace(" ", "").astype(float)
            fig, ax = plt.subplots(figsize=(10, 5))
            colors = ['#279244', '#f77f00', '#2a9d8f']
            ax.bar(df_graph_benef.index, df_graph_benef["Montant couvert"], color=colors)
            ax.set_title("Montants couverts par bénéficiaire")
            ax.set_xlabel("Type de bénéficiaire")
            ax.set_ylabel("Montant (FCFA)")
            ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f"{int(x):,}".replace(",", " ")))
            plt.xticks(rotation=45)
            st.pyplot(fig)
            conso_benef_path = os.path.join(tempfile.gettempdir(), "graph_benef.png")
            fig.savefig(conso_benef_path, bbox_inches='tight')
            plt.close(fig)
    except Exception as e:
        st.error(f"❌ Erreur lors du traitement de la consommation : {e}")

# Section 5 : Consommations mensuelles
if sinistralite_ok and df_filtre is not None and not df_filtre.empty:
    st.markdown("## IV - Consommations mensuelles")
    try:
        col_date = df_filtre.columns[1]
        col_police = df_filtre.columns[6]
        col_frais = df_filtre.columns[20]
        col_couvert = df_filtre.columns[22]
        col_rejet = df_filtre.columns[24] if 24 < len(df_filtre.columns) else None
        df_mensuel = df_filtre.copy()
        df_mensuel["DATE"] = pd.to_datetime(df_mensuel[col_date], errors="coerce")
        if df_mensuel["DATE"].isna().all():
            st.error("❌ La colonne des dates (colonne 1) contient des valeurs invalides ou est vide.")
            raise ValueError("Dates invalides pour les consommations mensuelles.")
        if col_rejet is not None and col_rejet in df_mensuel.columns and not df_mensuel[col_rejet].isna().all():
            df_mensuel[col_rejet] = pd.to_numeric(df_mensuel[col_rejet], errors='coerce')
            df_mensuel = df_mensuel[df_mensuel[col_rejet].notna()]
        else:
            st.warning("⚠️ La colonne des rejets (colonne 24) est absente ou vide. Traitement sans filtrage des rejets.")
        if df_mensuel.empty:
            st.warning("⚠️ Aucune donnée disponible pour les consommations mensuelles après filtrage.")
        else:
            df_mensuel = df_mensuel.sort_values(by="DATE", ascending=True)
            df_mensuel["MOIS"] = df_mensuel["DATE"].apply(format_date_fr)
            if df_mensuel["MOIS"].isna().any():
                st.warning("⚠️ Certaines dates n'ont pas pu être converties en mois. Vérifiez le format des dates.")
                df_mensuel = df_mensuel.dropna(subset=["MOIS"])
            if df_mensuel.empty:
                st.error("❌ Aucune donnée valide pour les mois après conversion des dates.")
            else:
                agg_dict = {
                    col_police: "count",
                    col_frais: "sum",
                    col_couvert: "sum"
                }
                if col_rejet is not None and col_rejet in df_mensuel.columns and not df_mensuel[col_rejet].isna().all():
                    agg_dict[col_rejet] = "sum"
                df_mensuel_grouped = df_mensuel.groupby("MOIS").agg(agg_dict).rename(columns={
                    col_police: "Nombre de Sinistres",
                    col_frais: "Frais réels",
                    col_couvert: "Montant Couvert",
                    col_rejet: "Rejets" if col_rejet in agg_dict else None
                }).reset_index()
                df_mensuel_grouped = df_mensuel_grouped.loc[:, ~df_mensuel_grouped.columns.isin([None])]
                mois_fr_to_en = {
                    "Janvier": "January", "Février": "February", "Mars": "March", "Avril": "April",
                    "Mai": "May", "Juin": "June", "Juillet": "July", "Août": "August",
                    "Septembre": "September", "Octobre": "October", "Novembre": "November", "Décembre": "December"
                }
                df_mensuel_grouped["MOIS_EN"] = df_mensuel_grouped["MOIS"].apply(
                    lambda x: " ".join([mois_fr_to_en.get(x.split()[0], x.split()[0]), x.split()[1]])
                )
                df_mensuel_grouped["MOIS_DATE"] = pd.to_datetime(df_mensuel_grouped["MOIS_EN"], format="%B %Y", errors='coerce')
                df_mensuel_grouped = df_mensuel_grouped.sort_values(by="MOIS_DATE", ascending=True).drop(columns=["MOIS_DATE", "MOIS_EN"])
                total_row = pd.DataFrame({
                    "MOIS": ["Total général"],
                    "Nombre de Sinistres": [df_mensuel_grouped["Nombre de Sinistres"].sum()],
                    "Frais réels": [df_mensuel_grouped["Frais réels"].sum()],
                    "Montant Couvert": [df_mensuel_grouped["Montant Couvert"].sum()],
                })
                if "Rejets" in df_mensuel_grouped.columns:
                    total_row["Rejets"] = [df_mensuel_grouped["Rejets"].sum()]
                df_mensuel_grouped = pd.concat([df_mensuel_grouped, total_row], ignore_index=True)
                for col in df_mensuel_grouped.columns:
                    if col != "MOIS":
                        df_mensuel_grouped[col] = df_mensuel_grouped[col].apply(lambda x: f"{int(x):,}".replace(",", " "))
                st.dataframe(df_mensuel_grouped)
                df_graph_mensuel = df_mensuel_grouped[df_mensuel_grouped["MOIS"] != "Total général"].copy()
                df_graph_mensuel["Montant Couvert"] = df_graph_mensuel["Montant Couvert"].str.replace(" ", "").astype(float)
                if "Rejets" in df_graph_mensuel.columns:
                    df_graph_mensuel["Rejets"] = df_graph_mensuel["Rejets"].str.replace(" ", "").astype(float)
                fig, ax = plt.subplots(figsize=(10, 5))
                bar_width = 0.35
                index = range(len(df_graph_mensuel["MOIS"]))
                ax.bar([i - bar_width/2 for i in index], df_graph_mensuel["Montant Couvert"], bar_width, label="Montant Couvert", color='#279244')
                if "Rejets" in df_graph_mensuel.columns:
                    ax.bar([i + bar_width/2 for i in index], df_graph_mensuel["Rejets"], bar_width, label="Rejets", color='#f77f00')
                ax.set_title("Montants Couverts et Rejets par Mois")
                ax.set_xlabel("Mois")
                ax.set_ylabel("Montant (FCFA)")
                ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f"{int(x):,}".replace(",", " ")))
                ax.set_xticks(index)
                ax.set_xticklabels(df_graph_mensuel["MOIS"], rotation=45)
                ax.legend()
                st.pyplot(fig)
                evol_mensuel_path = os.path.join(tempfile.gettempdir(), "graph_mensuel.png")
                fig.savefig(evol_mensuel_path, bbox_inches='tight')
                plt.close(fig)
    except Exception as e:
        st.error(f"❌ Erreur lors du traitement des consommations mensuelles : {e}")
else:
    st.warning("⚠️ Impossible de traiter les consommations mensuelles : données filtrées manquantes ou invalides.")

# Section 6 : Consommations par spécialité
if sinistralite_ok and df_filtre is not None and not df_filtre.empty:
    st.markdown("## V - Consommations par spécialité")
    try:
        col_specialite = df_filtre.columns[17]
        col_couvert = df_filtre.columns[22]
        col_rejets = df_filtre.columns[24] if 24 < len(df_filtre.columns) else None
        df_specialite = df_filtre.copy()
        if col_rejets is not None and col_rejets in df_specialite.columns:
            df_specialite = df_specialite[df_specialite[col_rejets].apply(lambda x: pd.notna(x) and isinstance(x, (int, float)))]
        tableau_spec = df_specialite.groupby(col_specialite).agg({
            col_specialite: "count",
            col_couvert: "sum",
            col_rejets: "sum" if col_rejets in df_specialite.columns else lambda x: 0
        }).rename(columns={
            col_specialite: "Nombre",
            col_couvert: "Couvert",
            col_rejets: "Rejets" if col_rejets in df_specialite.columns else None
        })
        total_row = pd.DataFrame({
            "Nombre": [tableau_spec["Nombre"].sum()],
            "Couvert": [tableau_spec["Couvert"].sum()],
            "Rejets": [tableau_spec["Rejets"].sum()] if "Rejets" in tableau_spec.columns else [0]
        }, index=["Total général"])
        tableau_spec = pd.concat([tableau_spec, total_row]).reset_index().rename(columns={'index': 'Spécialité'})
        for col in ["Nombre", "Couvert", "Rejets"]:
            tableau_spec[col] = tableau_spec[col].apply(lambda x: f"{int(x):,}".replace(",", " "))
        st.dataframe(tableau_spec)
        df_graph = tableau_spec[tableau_spec["Spécialité"] != "Total général"].copy()
        df_graph["Couvert"] = df_graph["Couvert"].str.replace(" ", "").astype(int)
        fig, ax = plt.subplots(figsize=(10, 6))
        total_couvert = df_graph["Couvert"].sum()
        
        # Préparer les données pour le pie chart
        colors = ['#06A77D', '#F58634', '#2a9d8f', '#ff6f61', '#264653', '#3498db', '#9b59b6', '#e74c3c', '#f1c40f', '#1abc9c']
        # S'assurer qu'il y a assez de couleurs pour toutes les spécialités
        while len(colors) < len(df_graph):
            colors.extend(colors)
        colors = colors[:len(df_graph)]
        
        # Ajuster les paramètres du pie chart pour ajouter les connexions
        wedges, texts, autotexts = ax.pie(
            df_graph["Couvert"], 
            autopct='%1.1f%%', 
            startangle=90, 
            explode=[0.05] * len(df_graph), 
            colors=colors, 
            textprops={'fontsize': 0},  # Masquer les étiquettes par défaut
            labeldistance=None,  # Supprimer les étiquettes par défaut
            pctdistance=0.75,    # Positionner les pourcentages plus près du centre
            wedgeprops={'edgecolor': 'white', 'linewidth': 1.5},  # Bordures blanches pour plus de clarté
            shadow=False  # Désactiver l'ombre pour éviter l'effet de double cercle
        )
        
        # Personnaliser le style des pourcentages
        for autotext in autotexts:
            autotext.set_color('white')
            autotext.set_fontsize(9)
            autotext.set_fontweight('bold')
            
        # Ajouter des annotations avec des lignes de connexion améliorées
        bbox_props = dict(boxstyle="round,pad=0.3", fc="white", ec="gray", lw=1, alpha=0.9)
        
        # Créer une liste des positions et des étiquettes
        labels_data = []
        for i, (wedge, text_label) in enumerate(zip(wedges, [f"{spec}" for spec in df_graph["Spécialité"]])):
            ang = (wedge.theta2 - wedge.theta1) / 2. + wedge.theta1
            x = np.cos(np.deg2rad(ang))
            y = np.sin(np.deg2rad(ang))
            
            # Ajuster la distance selon la longueur du texte
            text_length = len(text_label)
            if text_length > 15:  # Étiquettes longues comme "TRANSPORT PAR AMBULANCE"
                distance = 1.5
            elif text_length > 10:
                distance = 1.4
            else:
                distance = 1.35
            
            # Calculer la position finale
            final_x = distance * x
            final_y = distance * y
            
            # Déterminer l'alignement horizontal
            if x > 0.1:
                ha = "left"
            elif x < -0.1:
                ha = "right"
            else:
                ha = "center"
            
            labels_data.append({
                'text': text_label,
                'xy': (x, y),
                'xytext': (final_x, final_y),
                'ha': ha,
                'angle': ang
            })
        
        # Appliquer les annotations avec les paramètres optimisés
        for label_info in labels_data:
            # Style de connexion adaptatif
            kw_adaptive = dict(
                arrowprops=dict(
                    arrowstyle="-", 
                    color="gray", 
                    linewidth=1.2,
                    connectionstyle="arc3,rad=0.1"
                ),
                bbox=bbox_props, 
                zorder=10, 
                va="center",
                ha=label_info['ha'],
                fontsize=9,
                fontweight='normal'
            )
            
            ax.annotate(
                label_info['text'], 
                xy=label_info['xy'], 
                xytext=label_info['xytext'],
                **kw_adaptive
            )
        
        # Titre et style - augmenter l'espace avec le graphique
        fig.subplots_adjust(top=0.85, bottom=0.1, left=0.1, right=0.9)  # Améliorer les marges
        ax.set_title("Répartition par spécialité", fontsize=14, fontweight='bold', pad=40, y=1.1)
        
        # Ajuster les limites pour accueillir les étiquettes longues
        ax.set_xlim(-2, 2)
        ax.set_ylim(-2, 2)
        
        # Supprimer les cadres et les axes
        ax.set_frame_on(False)
        ax.axis('equal')  # Assurer un cercle parfait
        
        st.pyplot(fig)
        graph_path = os.path.join(tempfile.gettempdir(), "graph_specialite.png")
        fig.savefig(graph_path, bbox_inches='tight', dpi=300)
        plt.close(fig)
    except Exception as e:
        st.error(f"❌ Erreur lors du traitement des spécialités : {e}")

# Section 7 : Top des prestataires
if sinistralite_ok and df_filtre is not None and not df_filtre.empty:
    st.markdown("## VI - Top des prestataires")
    try:
        col_prestataire = df_filtre.columns[13]
        col_ville = df_filtre.columns[14]
        col_commune = df_filtre.columns[15]
        col_couvert = df_filtre.columns[22]
        df_prestataires = df_filtre.groupby([col_prestataire, col_ville, col_commune]).agg({
            col_prestataire: "count",
            col_couvert: "sum"
        }).rename(columns={col_prestataire: "NOMBRE", col_couvert: "Couvert"}).reset_index()
        df_prestataires.columns = ["PRESTATAIRE", "VILLE", "COMMUNE", "NOMBRE", "Couvert"]
        df_prestataires = df_prestataires.sort_values(by="Couvert", ascending=False)
        total_covered = df_prestataires["Couvert"].sum()
        total_nombre = df_prestataires["NOMBRE"].sum()
        df_prestataires["Proportion"] = (df_prestataires["Couvert"] / total_covered * 100).round(0).astype(int).astype(str) + "%"
        df_prestataires["Ordre"] = range(1, len(df_prestataires) + 1)
        df_prestataires = df_prestataires[["Ordre", "PRESTATAIRE", "VILLE", "COMMUNE", "NOMBRE", "Couvert", "Proportion"]]
        df_prestataires["Couvert"] = df_prestataires["Couvert"].apply(lambda x: f"{int(x):,}".replace(",", " "))
        total_row = pd.DataFrame({
            "Ordre": [""],
            "PRESTATAIRE": ["Total"],
            "VILLE": [""],
            "COMMUNE": [""],
            "NOMBRE": [f"{int(total_nombre):,}".replace(",", " ")],
            "Couvert": [f"{int(total_covered):,}".replace(",", " ")],
            "Proportion": ["100%"]
        })
        df_prestataires = pd.concat([df_prestataires, total_row], ignore_index=True)
        st.dataframe(df_prestataires)
    except Exception as e:
        st.error(f"❌ Erreur lors du traitement des prestataires : {e}")

# Section 8 : Top des Familles de Consommateurs
if sinistralite_ok and df_filtre is not None and not df_filtre.empty:
    st.markdown("## VII - Top des Familles de Consommateurs")
    try:
        # Rechercher les colonnes spécifiques par nom plutôt que par position
        col_couvert = df_filtre.columns[22]  # Colonne montant couvert
        
        # Rechercher la colonne N°CARTE ASSURÉ PRINCIPAL par nom approximatif
        col_carte_assure_principal = None
        col_nom_assure_principal = None
        
        for col in df_filtre.columns:
            col_str = str(col).upper().strip()
            # Recherche de la colonne N°CARTE ASSURÉ PRINCIPAL
            if "CARTE" in col_str and ("ASSURE" in col_str or "ASSURÉ" in col_str) and "PRINCIPAL" in col_str:
                col_carte_assure_principal = col
            # Recherche de la colonne ASSURÉ PRINCIPAL
            elif ("ASSURE" in col_str or "ASSURÉ" in col_str) and "PRINCIPAL" in col_str and "CARTE" not in col_str:
                col_nom_assure_principal = col
                
        # Si les colonnes spécifiques n'ont pas été trouvées, utiliser des positions par défaut
        if col_carte_assure_principal is None:
            col_carte_assure_principal = df_filtre.columns[9]  # Position par défaut
            st.warning("⚠️ Colonne 'N°CARTE ASSURÉ PRINCIPAL' non trouvée. Utilisation de la colonne par défaut.")
            
        if col_nom_assure_principal is None:
            col_nom_assure_principal = df_filtre.columns[11]  # Position par défaut
            st.warning("⚠️ Colonne 'ASSURÉ PRINCIPAL' non trouvée. Utilisation de la colonne par défaut.")
            
        # Afficher les colonnes utilisées (pour debug et information)
        st.info(f"Colonnes utilisées : Carte Assuré Principal = '{df_filtre.columns[df_filtre.columns.get_loc(col_carte_assure_principal)]}', Nom Assuré Principal = '{df_filtre.columns[df_filtre.columns.get_loc(col_nom_assure_principal)]}'")
            
        # Grouper par numéro de carte de l'assuré principal pour obtenir le cumul des dépenses par famille
        df_familles = df_filtre.groupby([col_carte_assure_principal, col_nom_assure_principal]).agg({
            col_carte_assure_principal: "count",
            col_couvert: "sum"
        }).rename(columns={col_carte_assure_principal: "Nombre d'actes", col_couvert: "Couvert"}).reset_index()
        
        df_familles.columns = ["N° de Famille", "Assuré Principal", "Nombre d'actes", "Couvert"]
        df_familles = df_familles.sort_values(by="Couvert", ascending=False)
        total_covered = df_familles["Couvert"].sum()
        total_actes = df_familles["Nombre d'actes"].sum()
        df_familles["Proportion"] = (df_familles["Couvert"] / total_covered * 100).round(0).astype(int).astype(str) + "%"
        df_familles["Ordre"] = range(1, len(df_familles) + 1)
        df_familles["Couvert"] = df_familles["Couvert"].apply(lambda x: f"{int(x):,}".replace(",", " "))
        df_familles = df_familles[["Ordre", "N° de Famille", "Assuré Principal", "Nombre d'actes", "Couvert", "Proportion"]]
        total_row = pd.DataFrame({
            "Ordre": [""],
            "N° de Famille": ["Total"],
            "Assuré Principal": [""],
            "Nombre d'actes": [f"{int(total_actes):,}".replace(",", " ")],
            "Couvert": [f"{int(total_covered):,}".replace(",", " ")],
            "Proportion": ["100%"]
        })
        df_familles = pd.concat([df_familles, total_row], ignore_index=True)
        st.dataframe(df_familles)
    except Exception as e:
        st.error(f"❌ Erreur lors du traitement des familles de consommateurs : {e}")

# Section 9 : Logos et génération PDF
st.subheader("Ajouter des logos (optionnel)")
st.markdown("### Logo Ankara")
fichier_logo_ankara = st.file_uploader("Joindre le logo Ankara (PNG, JPG)", type=["png", "jpg", "jpeg"], key="logo_ankara")
if fichier_logo_ankara:
    logo_bytes = fichier_logo_ankara.read()
    logo_ankara_path = os.path.join(tempfile.gettempdir(), "logo_ankara_temp.png")
    with open(logo_ankara_path, "wb") as f:
        f.write(logo_bytes)
    st.success("✅ Logo Ankara chargé avec succès !")

st.markdown("### Logo Assureur")
fichier_logo_assureur = st.file_uploader("Joindre le logo de l'assureur (PNG, JPG)", type=["png", "jpg", "jpeg"], key="logo_assureur")
if fichier_logo_assureur:
    logo_bytes = fichier_logo_assureur.read()
    logo_assureur_path = os.path.join(tempfile.gettempdir(), "logo_assureur_temp.png")
    with open(logo_assureur_path, "wb") as f:
        f.write(logo_bytes)
    st.success("✅ Logo Assureur chargé avec succès !")

# Fonction pour nettoyer les fichiers temporaires
def cleanup_temp_files():
    for path in [graph_path, evol_effectif_path, evol_mensuel_path, conso_benef_path, logo_ankara_path, logo_assureur_path]:
        if path and os.path.exists(path):
            try:
                os.remove(path)
            except:
                pass

# Validation des fichiers avant génération
if not all([fichier_detail, fichier_production, fichier_effectif]):
    st.warning("⚠️ Veuillez charger tous les fichiers requis (DETAIL, PRODUCTION, EFFECTIF) avant de générer le PDF.")
elif st.button("Générer le PDF"):
    try:
        with st.spinner("Génération du PDF en cours..."):
            class PDFWithPageNumbers(FPDF):
                def __init__(self, total_pages=0):
                    super().__init__()
                    self.total_pages = total_pages

                def header(self):
                    if self.page_no() > 2:
                        if logo_ankara_path and os.path.exists(logo_ankara_path):
                            self.image(logo_ankara_path, x=10, y=10, w=30)
                        self.ln(10)

                def footer(self):
                    if self.page_no() > 2:
                        self.set_y(-15)
                        self.set_fill_color(39, 146, 68)
                        self.rect(0, self.h - 15, self.w, 10, 'F')
                        self.set_font("Arial", "I", 8)  # Augmenté de 7 à 8
                        self.set_text_color(255, 255, 255)
                        page_text = f"Statistiques {clean_text(nom_assureur)}_{clean_text(client_short)} - Page {self.page_no() - 2} / {self.total_pages}"
                        self.cell(0, 10, page_text, align="C")

            def clean_text(text):
                """
                Nettoie le texte pour gérer correctement les caractères accentués et spéciaux.
                Convertit les caractères Unicode en leur équivalent ASCII compatible avec Latin-1.
                Préserve certains caractères accentués importants comme 'à'.
                """
                if isinstance(text, pd.DataFrame):
                    # Traitement pour DataFrame
                    for col in text.columns:
                        text[col] = text[col].apply(lambda x: clean_text(x) if isinstance(x, str) else str(x))
                    return text
                elif isinstance(text, str):
                    # Préserver certains caractères accentués spécifiques
                    preserved_chars = {'à': 'à', 'À': 'À', 'é': 'é', 'É': 'É', 'è': 'è', 'È': 'È'}
                    
                    # Sauvegarder les caractères à préserver
                    for char, replacement in preserved_chars.items():
                        text = text.replace(char, f"__PRESERVED_{ord(char)}__")
                    
                    # Normalisation Unicode pour décomposer les caractères accentués
                    text = unicodedata.normalize('NFKD', text)
                    # Convertir en ASCII, en ignorant les caractères non-ASCII
                    text = text.encode('ascii', 'ignore').decode('ascii')
                    
                    # Restaurer les caractères préservés
                    for char, replacement in preserved_chars.items():
                        text = text.replace(f"__PRESERVED_{ord(char)}__", replacement)
                    
                    # Remplacements supplémentaires pour caractères spécifiques
                    replacements = {
                        'Œ': 'OE', 'œ': 'oe', '…': '...', '–': '-', '—': '-', '\u2013': '-', '\u2014': '-',
                        '\u2018': "'", '\u2019': "'", '\u2022': '*', '"': '"', '"': '"'
                    }
                    for k, v in replacements.items():
                        text = text.replace(k, v)
                    return text
                return str(text)

            def add_table_section(pdf, title, df, is_prestataires=False, is_familles=False, highlight_row=None, new_page=True):
                if new_page:
                    pdf.add_page()
                section_page = pdf.page_no() - 2
                pdf.set_font("Arial", 'B', 13)
                pdf.set_text_color(0, 0, 0)
                pdf.cell(0, 10, clean_text(title), ln=True)
                pdf.ln(5)
                pdf.set_font("Arial", '', 8)  # Augmenté de 7 à 8
                pdf.set_text_color(0, 0, 0)
                page_width = float(pdf.w - 2 * pdf.l_margin)
                line_height = 5.0
                
                # Définir les largeurs des colonnes selon la section
                if title == "Section I - Sinistralité":
                    # Largeur augmentée pour la colonne "Client" (index 3)
                    col_widths = [25.0, 25.0, 30.0, 60.0, 20.0, 20.0, 20.0, 20.0]
                elif is_prestataires:
                    col_widths = [15.0, 50.0, 20.0, 25.0, 20.0, 30.0, 30.0]
                elif is_familles:
                    col_widths = [15.0, 30.0, 50.0, 30.0, 30.0, 30.0]
                else:
                    num_cols = len(df.columns)
                    col_widths = [page_width / num_cols] * num_cols
                
                total_width = sum(col_widths)
                if total_width > page_width:
                    scale_factor = page_width / total_width
                    col_widths = [w * scale_factor for w in col_widths]
                
                df_display = clean_text(df.copy())
                df_calc = df_display.copy()
                for col in df_calc.columns:
                    temp_col = df_calc[col].astype(str)
                    df_calc[col] = pd.to_numeric(temp_col.str.replace(" ", "").str.replace("%", ""), errors='coerce').fillna(df_calc[col])
                
                max_header_lines = 1
                for col in df_display.columns:
                    col_width = col_widths[df_display.columns.get_loc(col)]
                    num_lines = max(1, len(str(col).split('\n')) + int(pdf.get_string_width(str(col).upper()) / (col_width - 2)))
                    max_header_lines = max(max_header_lines, num_lines)
                
                header_height = line_height * float(max_header_lines) + 2
                header_text_y_offset = (header_height - line_height * float(max_header_lines)) / 2
                
                max_lines_per_row = []
                for _, row in df_display.iterrows():
                    max_lines = 1
                    for j, item in enumerate(row):
                        num_lines = max(1, len(str(item).split('\n')) + int(pdf.get_string_width(str(item)) / (col_widths[j] - 2)))
                        max_lines = max(max_lines, num_lines)
                    max_lines_per_row.append(max_lines)
                
                pdf.set_fill_color(39, 146, 68)
                pdf.set_text_color(255, 255, 255)
                pdf.set_font("Arial", 'B', 7)
                x_start = float(pdf.l_margin)
                y_start = float(pdf.get_y())
                
                for i, col in enumerate(df_display.columns):
                    pdf.set_xy(x_start, y_start)
                    pdf.cell(col_widths[i], header_height, '', border=0, fill=True)
                    pdf.set_xy(x_start, y_start + header_text_y_offset)
                    pdf.multi_cell(col_widths[i], line_height, clean_text(str(col).upper()), border=0, align='C')
                    x_start += col_widths[i]
                
                pdf.set_y(y_start + header_height)
                table_width = float(sum(col_widths))
                pdf.set_draw_color(0, 0, 0)
                pdf.set_line_width(0.2)
                pdf.line(pdf.l_margin, y_start, pdf.l_margin + table_width, y_start)
                pdf.set_text_color(0, 0, 0)
                pdf.set_font("Arial", '', 7)
                table_y_start = float(pdf.get_y())
                page_height = float(pdf.h)
                bottom_margin = float(pdf.b_margin)
                
                for i, (_, row) in enumerate(df_display.iterrows()):
                    row_height = line_height * float(max_lines_per_row[i])
                    current_y = float(pdf.get_y())
                    if current_y + row_height > page_height - bottom_margin - 15:
                        y_end = page_height - bottom_margin - 15
                        pdf.line(pdf.l_margin, table_y_start, pdf.l_margin, y_end)
                        pdf.line(pdf.l_margin + table_width, table_y_start, pdf.l_margin + table_width, y_end)
                        pdf.add_page()
                        table_y_start = float(pdf.get_y())
                        pdf.set_fill_color(39, 146, 68)
                        pdf.set_text_color(255, 255, 255)
                        pdf.set_font("Arial", 'B', 7)
                        x_start = float(pdf.l_margin)
                        for j, col in enumerate(df_display.columns):
                            pdf.set_xy(x_start, table_y_start)
                            pdf.cell(col_widths[j], header_height, '', border=0, fill=True)
                            pdf.set_xy(x_start, table_y_start + header_text_y_offset)
                            pdf.multi_cell(col_widths[j], line_height, clean_text(str(col).upper()), border=0, align='C')
                            x_start += col_widths[j]
                        pdf.set_y(table_y_start + header_height)
                        pdf.set_draw_color(0, 0, 0)
                        pdf.set_line_width(0.2)
                        pdf.line(pdf.l_margin, table_y_start, pdf.l_margin + table_width, table_y_start)
                        pdf.set_text_color(0, 0, 0)
                        pdf.set_font("Arial", '', 6.5)
                    
                    if highlight_row is not None and i == highlight_row:
                        pdf.set_fill_color(247, 127, 0)
                    else:
                        if i % 2 == 0:
                            pdf.set_fill_color(255, 255, 255)
                        else:
                            pdf.set_fill_color(220, 220, 220)
                    
                    pdf.set_draw_color(0, 0, 0)
                    pdf.set_line_width(0.1)
                    x_start = float(pdf.l_margin)
                    y_start = float(pdf.get_y())
                    for j, item in enumerate(row):
                        pdf.set_xy(x_start, y_start)
                        pdf.cell(col_widths[j], row_height, '', border='T' if i == 0 else 'TB', fill=True)
                        pdf.set_xy(x_start, y_start)
                        pdf.multi_cell(col_widths[j], line_height, clean_text(str(item)), border=0, align='C')
                        x_start += col_widths[j]
                    pdf.set_y(y_start + row_height)
                
                table_y_end = float(pdf.get_y())
                pdf.set_draw_color(0, 0, 0)
                pdf.set_line_width(0.2)
                pdf.line(pdf.l_margin, table_y_start, pdf.l_margin, table_y_end)
                pdf.line(pdf.l_margin + table_width, table_y_start, pdf.l_margin + table_width, table_y_end)
                pdf.line(pdf.l_margin, table_y_end, pdf.l_margin + table_width, table_y_end)
                pdf.ln(5)
                return section_page

            temp_pdf = PDFWithPageNumbers()
            temp_pdf.set_auto_page_break(auto=True, margin=15)
            temp_pdf.alias_nb_pages()
            temp_pdf.add_page()
            temp_pdf.add_page()
            page_numbers = []
            if df_sin is not None:
                page_numbers.append(add_table_section(temp_pdf, "Section I - Sinistralité", df_sin))
                if df_clause is not None:
                    page_numbers.append(add_table_section(temp_pdf, "Clause Ajustement Santé", df_clause, highlight_row=highlight_row, new_page=False))
                else:
                    page_numbers.append(0)
            if df_effectif is not None:
                page_numbers.append(add_table_section(temp_pdf, "Section II - Évolution des effectifs", df_effectif_display))
                if evol_effectif_path and os.path.exists(evol_effectif_path):
                    temp_pdf.image(evol_effectif_path, x=10, w=180)
                    temp_pdf.ln(5)
            if 'tableau_final' in locals():
                page_numbers.append(add_table_section(temp_pdf, "Section III - Consommation par type de bénéficiaire", tableau_final.reset_index().rename(columns={"index": "Type de bénéficiaire"})))
                if conso_benef_path and os.path.exists(conso_benef_path):
                    temp_pdf.image(conso_benef_path, x=10, w=180)
                    temp_pdf.ln(5)
            if 'df_mensuel_grouped' in locals():
                page_numbers.append(add_table_section(temp_pdf, "Section IV - Consommation mensuelle", df_mensuel_grouped))
                if evol_mensuel_path and os.path.exists(evol_mensuel_path):
                    temp_pdf.image(evol_mensuel_path, x=10, w=180)
                    temp_pdf.ln(5)
            if 'tableau_spec' in locals():
                page_numbers.append(add_table_section(temp_pdf, "Section V - Consommation par spécialité", tableau_spec))
                if graph_path and os.path.exists(graph_path):
                    temp_pdf.image(graph_path, x=10, w=180)
                    temp_pdf.ln(5)
            if 'df_prestataires' in locals():
                page_numbers.append(add_table_section(temp_pdf, "Section VI - Top des prestataires", df_prestataires, is_prestataires=True))
            if 'df_familles' in locals():
                page_numbers.append(add_table_section(temp_pdf, "Section VII - Top des Familles de Consommateurs", df_familles, is_familles=True))
            total_pages = temp_pdf.page_no() - 2
            
            pdf = PDFWithPageNumbers(total_pages=total_pages)
            pdf.set_auto_page_break(auto=True, margin=15)
            pdf.alias_nb_pages()
            pdf.add_page()
            if logo_ankara_path and os.path.exists(logo_ankara_path):
                pdf.image(logo_ankara_path, x=(pdf.w - 80) / 2, y=20, w=80)
                pdf.ln(90)
            pdf.set_font("Arial", 'B', 24)
            pdf.set_text_color(39, 146, 68)
            pdf.cell(0, 20, clean_text("STATISTIQUES DE GESTION SANTE"), ln=True, align="C")
            pdf.ln(20)
            info_box_width, info_box_height = 140, 50
            info_box_x, info_box_y = (pdf.w - info_box_width) / 2, float(pdf.get_y())
            pdf.set_fill_color(245, 245, 245)
            pdf.rect(info_box_x, info_box_y, info_box_width, info_box_height, 'F')
            pdf.set_draw_color(54, 69, 79)
            pdf.set_line_width(0.3)
            pdf.rect(info_box_x, info_box_y, info_box_width, info_box_height)
            line_height = 10
            total_text_height = 4 * line_height
            start_y = info_box_y + (info_box_height - total_text_height) / 2
            label_width = 40
            value_width = info_box_width - label_width - 15
            pdf.set_font("Arial", '', 12)
            pdf.set_text_color(54, 69, 79)

            # Calculer la période à partir des données d'effectifs
            periode_filtre = clean_text(periode)  # Valeur par défaut
            if 'df_effectif_filtered' in locals() and df_effectif_filtered is not None and not df_effectif_filtered.empty:
                try:
                    # Récupérer les dates de la première colonne de la table d'effectifs
                    dates_effectif = df_effectif_filtered["MOIS"]
                    
                    # Convertir les dates selon leur format actuel
                    raw_dates = []
                    mois_inverse = {v.lower(): k for k, v in mois_fr.items()}
                    
                    for date_val in dates_effectif:
                        try:
                            # Si c'est déjà un objet datetime, l'utiliser directement
                            if isinstance(date_val, (pd.Timestamp, datetime)):
                                raw_dates.append(date_val)
                            else:
                                # Convertir depuis format français "Mois Année"
                                parts = str(date_val).split()
                                if len(parts) >= 2:
                                    mois_fr_name = parts[0].lower()
                                    annee = parts[1]
                                    mois_en = mois_inverse.get(mois_fr_name)
                                    if mois_en:
                                        raw_dates.append(pd.to_datetime(f"{mois_en} 15, {annee}"))
                                    else:
                                        raw_dates.append(pd.NaT)
                                else:
                                    raw_dates.append(pd.NaT)
                        except:
                            raw_dates.append(pd.NaT)
                    
                    raw_dates = pd.Series(raw_dates)
                    
                    # Calculer la période si on a des dates valides
                    if not raw_dates.isna().all():
                        date_min = raw_dates.min()
                        date_max = raw_dates.max()
                        mois_min = mois_fr.get(date_min.strftime("%B"), date_min.strftime("%B"))
                        mois_max = mois_fr.get(date_max.strftime("%B"), date_max.strftime("%B"))
                        annee_min = date_min.strftime("%Y")
                        annee_max = date_max.strftime("%Y")
                        
                        if annee_min == annee_max:
                            periode_filtre = clean_text(f"de {mois_min} à {mois_max} {annee_max}")
                        else:
                            periode_filtre = clean_text(f"de {mois_min} {annee_min} à {mois_max} {annee_max}")
                except:
                    pass

            pdf.set_xy(info_box_x + 10, start_y)
            pdf.set_font("Arial", 'B', 12)
            pdf.cell(label_width, line_height, "Assureur : ", align='L')
            pdf.set_font("Arial", '', 12)
            pdf.set_x(info_box_x + 10 + label_width)
            # Utiliser la version courte du nom de l'assureur
            assureur_short = extract_assureur_words(nom_assureur, max_words=3)
            pdf.multi_cell(value_width, line_height, clean_text(assureur_short), align='L')

            pdf.set_xy(info_box_x + 10, start_y + line_height)
            pdf.set_font("Arial", 'B', 12)
            pdf.cell(label_width, line_height, "Client : ", align='L')
            pdf.set_font("Arial", '', 12)
            pdf.set_x(info_box_x + 10 + label_width)
            pdf.multi_cell(value_width, line_height, clean_text(client_short), align='L')

            pdf.set_xy(info_box_x + 10, start_y + 2 * line_height)
            pdf.set_font("Arial", 'B', 12)
            pdf.cell(label_width, line_height, "Période : ", align='L')
            pdf.set_font("Arial", '', 12)
            pdf.set_x(info_box_x + 10 + label_width)
            pdf.multi_cell(value_width, line_height, periode_filtre, align='L')

            pdf.set_xy(info_box_x + 10, start_y + 3 * line_height)
            pdf.set_font("Arial", 'B', 12)
            pdf.cell(label_width, line_height, "Date d'édition : ", align='L')
            pdf.set_font("Arial", '', 12)
            pdf.set_x(info_box_x + 10 + label_width)
            pdf.multi_cell(value_width, line_height, clean_text(datetime.now().strftime("%d/%m/%Y")), align='L')

            pdf.set_y(info_box_y + info_box_height + 10)
            if logo_assureur_path and os.path.exists(logo_assureur_path):
                pdf.image(logo_assureur_path, x=(pdf.w - 50) / 2, y=float(pdf.get_y()), w=50)
                pdf.ln(60)
            pdf.set_font("Arial", 'I', 10)
            pdf.set_text_color(54, 69, 79)
            pdf.multi_cell(0, 5, clean_text("Ankara Services, Abidjan – Plateau, Avenue Noguès Immeuble Borija, Tel :+225 25 20 01 31 05/06\n"
                                "Société Anonyme avec Conseil d'Administration au Capital de 10.000.000 FCFA - 01 BP 1194 ABJ 01\n"
                                "RCCM CI- ABJ-03-2021-B14-00020-NCC 2110076 V-Banque : STANBIC CI198 01001 918000005921 84\n"
                                "www.ankaraservives.com"), align="C")
            pdf.add_page()
            pdf.set_font("Arial", 'B', 16)
            pdf.set_text_color(39, 146, 68)
            pdf.cell(0, 10, clean_text("SOMMAIRE"), ln=True, align="C")
            pdf.ln(10)
            sections_updated = [
                ("Section I - Sinistralité", page_numbers[0] if len(page_numbers) > 0 else 1),
                ("Clause Ajustement Santé", page_numbers[1] if len(page_numbers) > 1 and page_numbers[1] != 0 else 0),
                ("Section II - Évolution des effectifs", page_numbers[2] if len(page_numbers) > 2 else page_numbers[1] if len(page_numbers) > 1 else 2),
                ("Section III - Consommation par type de bénéficiaire", page_numbers[3] if len(page_numbers) > 3 else page_numbers[2] if len(page_numbers) > 2 else 3),
                ("Section IV - Consommation mensuelle", page_numbers[4] if len(page_numbers) > 4 else page_numbers[3] if len(page_numbers) > 3 else 4),
                ("Section V - Consommation par spécialité", page_numbers[5] if len(page_numbers) > 5 else page_numbers[4] if len(page_numbers) > 4 else 5),
                ("Section VI - Top des prestataires", page_numbers[6] if len(page_numbers) > 6 else page_numbers[5] if len(page_numbers) > 5 else 6),
                ("Section VII - Top des Familles de Consommateurs", page_numbers[7] if len(page_numbers) > 7 else page_numbers[6] if len(page_numbers) > 6 else 7)
            ]
            pdf.set_font("Arial", '', 12)
            pdf.set_text_color(54, 69, 79)
            page_width = pdf.w - pdf.l_margin - pdf.r_margin
            right_align_x = pdf.w - pdf.r_margin - 20
            left_shift = pdf.l_margin + 20
            for title, page in sections_updated:
                if page == 0:
                    continue
                dots = "." * int((right_align_x - pdf.get_string_width(clean_text(title)) - pdf.get_string_width(f"Page {page}") - left_shift - 10) / pdf.get_string_width("."))
                pdf.set_x(left_shift)
                pdf.cell(0, 8, f"{clean_text(title)}{dots}Page {page}", ln=True, align="L")
                pdf.ln(8)
            if df_sin is not None:
                add_table_section(pdf, "Section I - Sinistralité", df_sin)
                if df_clause is not None:
                    ratio_sp_value = float(df_sin["S/P"].iloc[0].replace("%", "")) / 100
                    ratio_sp_rounded = round(ratio_sp_value * 100) / 100
                    highlight_row = None
                    if tranche_min_col and tranche_max_col:
                        for idx, row in df_clause.iterrows():
                            try:
                                tranche_min = float(str(row[tranche_min_col]).replace('%', '')) / 100
                                tranche_max = float(str(row[tranche_max_col]).replace('%', '')) / 100
                                if tranche_min <= ratio_sp_rounded <= tranche_max:
                                    highlight_row = idx
                                    break
                            except (ValueError, TypeError):
                                continue
                    add_table_section(pdf, "Clause Ajustement Santé", df_clause, highlight_row=highlight_row, new_page=False)
            if df_effectif is not None:
                add_table_section(pdf, "Section II - Évolution des effectifs", df_effectif_display)
                if evol_effectif_path and os.path.exists(evol_effectif_path):
                    pdf.image(evol_effectif_path, x=10, w=180)
                    pdf.ln(5)
            if 'tableau_final' in locals():
                add_table_section(pdf, "Section III - Consommation par type de bénéficiaire", tableau_final.reset_index().rename(columns={"index": "Type de bénéficiaire"}))
                if conso_benef_path and os.path.exists(conso_benef_path):
                    pdf.image(conso_benef_path, x=10, w=180)
                    pdf.ln(5)
            if 'df_mensuel_grouped' in locals():
                add_table_section(pdf, "Section IV - Consommation mensuelle", df_mensuel_grouped)
                if evol_mensuel_path and os.path.exists(evol_mensuel_path):
                    pdf.image(evol_mensuel_path, x=10, w=180)
                    pdf.ln(5)
            if 'tableau_spec' in locals():
                add_table_section(pdf, "Section V - Consommation par spécialité", tableau_spec)
                if graph_path and os.path.exists(graph_path):
                    pdf.image(graph_path, x=10, w=180)
                    pdf.ln(5)
            if 'df_prestataires' in locals():
                add_table_section(pdf, "Section VI - Top des prestataires", df_prestataires, is_prestataires=True)
            if 'df_familles' in locals():
                add_table_section(pdf, "Section VII - Top des Familles de Consommateurs", df_familles, is_familles=True)
            
            # Génération du PDF
            pdf_output = BytesIO()
            pdf_bytes = pdf.output(dest='S').encode('latin1')
            pdf_output.write(pdf_bytes)
            pdf_output.seek(0)
            filename = f"{clean_text(nom_assureur)}_{clean_text(client_short)}_rapport_sante.pdf"
            
            # Téléchargement du PDF
            st.download_button("Télécharger le PDF", pdf_output, file_name=filename)
            st.success("✅ PDF généré avec succès !")
    except Exception as e:
        st.error(f"❌ Erreur lors de la génération du PDF : {e}")
        import traceback
        st.error(traceback.format_exc())
    finally:
        cleanup_temp_files()
