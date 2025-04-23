import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO
import os
import tempfile
from fpdf import FPDF
from datetime import datetime
import locale

# Définir la locale française pour gérer les mois en français
locale.setlocale(locale.LC_TIME, 'fr_FR.UTF-8')  # Pour Linux/Mac
try:
    locale.setlocale(locale.LC_TIME, 'french')  # Pour Windows
except locale.Error:
    locale.setlocale(locale.LC_TIME, '')  # Retour à la locale par défaut si 'french' échoue

# Chemins temporaires pour graphiques et logos
graph_path = None
evol_effectif_path = None
evol_mensuel_path = None
conso_benef_path = None
logo_ankara_path = None
logo_assureur_path = None

# Configuration de la page
st.set_page_config(page_title="Générateur de Rapport Santé", layout="wide", initial_sidebar_state="collapsed")

# CSS personnalisé
st.markdown("""
    <style>
    * {
        font-family: 'Arial', sans-serif;
        transition: all 0.3s ease;
    }
    body {
        background: #f5f7fa;
        color: #333333;
    }
    .stApp {
        max-width: 1200px;
        margin: 0 auto;
        padding: 30px;
        background-color: #ffffff;
        border-radius: 16px;
        box-shadow: 0 8px 30px rgba(0, 0, 0, 0.05);
    }
    h1 {
        font-size: 2.5rem;
        font-weight: 700;
        color: #279244;
        text-align: center;
        margin-bottom: 2.5rem;
        letter-spacing: -0.5px;
    }
    h2 {
        font-size: 1.5rem;
        font-weight: 600;
        color: #279244;
        margin-bottom: 1.5rem;
        border-bottom: 3px solid #f77f00;
        padding-bottom: 8px;
        display: inline-block;
    }
    .stFileUploader, .stSelectbox, .stTextInput, .stNumberInput {
        background-color: #ffffff;
        border: 2px solid #d1d5db;
        border-radius: 8px;
        padding: 12px;
        box-shadow: 0 2px 10px rgba(0, 0, 0, 0.05);
        margin-bottom: 1.5rem;
        color: #333333;
    }
    .stFileUploader:hover, .stSelectbox:hover, .stTextInput:hover, .stNumberInput:hover {
        border-color: #9ca3af;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.1);
    }
    .stButton>button {
        background-color: #279244;
        color: #ffffff;
        border: none;
        border-radius: 8px;
        padding: 0.75rem 2rem;
        font-weight: 600;
        box-shadow: 0 3px 10px rgba(39, 146, 68, 0.3);
        transition: background-color 0.3s ease, transform 0.2s ease, box-shadow 0.3s ease;
    }
    .stButton>button:hover {
        background-color: #1e6f33;
        transform: translateY(-2px);
        box-shadow: 0 5px 15px rgba(39, 146, 68, 0.4);
    }
    .stDownloadButton>button {
        background-color: #f77f00;
        color: #ffffff;
        border: none;
        border-radius: 8px;
        padding: 0.75rem 2rem;
        font-weight: 600;
        box-shadow: 0 3px 10px rgba(247, 127, 0, 0.3);
        transition: background-color 0.3s ease, transform 0.2s ease, box-shadow 0.3s ease;
    }
    .stDownloadButton>button:hover {
        background-color: #d66c00;
        transform: translateY(-2px);
        box-shadow: 0 5px 15px rgba(247, 127, 0, 0.4);
    }
    .stDataFrame {
        border: 2px solid #d1d5db;
        border-radius: 10px;
        overflow: hidden;
        box-shadow: 0 3px 10px rgba(0, 0, 0, 0.05);
        background-color: #ffffff;
        color: #333333;
    }
    .stDataFrame table {
        width: 100%;
    }
    .stDataFrame th {
        background-color: #ff6f61;
        color: #ffffff;
        font-weight: 600;
    }
    .stDataFrame tr:nth-child(even) {
        background-color: #f8f9fa;
    }
    .stError, .stWarning {
        background-color: #ff6f6110;
        color: #ff6f61;
        border: 1px solid #ff6f61;
        border-radius: 8px;
        padding: 12px;
        box-shadow: 0 2px 8px rgba(255, 111, 97, 0.1);
    }
    div[data-testid="stVerticalBlock"] {
        padding: 20px;
        border-radius: 12px;
        background-color: #ffffff;
        margin-bottom: 1.5rem;
        box-shadow: 0 2px 10px rgba(0, 0, 0, 0.05);
    }
    @media (max-width: 768px) {
        .stApp {
            padding: 15px;
        }
        h1 {
            font-size: 2rem;
        }
        h2 {
            font-size: 1.25rem;
        }
        .stButton>button, .stDownloadButton>button {
            width: 100%;
            padding: 0.75rem;
        }
        .stFileUploader, .stSelectbox, .stTextInput, .stNumberInput {
            padding: 10px;
        }
    }
    </style>
""", unsafe_allow_html=True)

# Titre principal
st.title("Générateur de Rapport Santé")

# Section 1 : Informations du contrat
st.subheader("Informations du contrat")
with st.container():
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        fichier_detail = st.file_uploader("Joindre DETAIL.xlsx", type="xlsx", help="Fichier Excel contenant les détails des sinistres")
    with col2:
        fichier_production = st.file_uploader("Joindre PRODUCTION.xlsx", type="xlsx", help="Fichier Excel avec les données de production")
    with col3:
        fichier_effectif = st.file_uploader("Joindre EFFECTIF.xlsx", type="xlsx", help="Fichier Excel des effectifs")
    with col4:
        fichier_clause = st.file_uploader("Joindre Clause Ajustement Santé.xlsx", type="xlsx", help="Fichier Excel des clauses d'ajustement santé")

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

if fichier_detail:
    try:
        xls = pd.ExcelFile(fichier_detail)
        if "DETAIL" in xls.sheet_names:
            df_detail = xls.parse("DETAIL")
            # Nettoyer les données pour éviter les problèmes d'espaces ou de casse
            df_detail.iloc[:, 5] = df_detail.iloc[:, 5].astype(str).str.strip().str.upper().str.replace(r'\s+', ' ', regex=True)
            df_detail.iloc[:, 27] = df_detail.iloc[:, 27].astype(str).str.strip().str.upper().str.replace(r'\s+', ' ', regex=True)
            
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
                    periode_value = ""
                    if 'df_filtre' in locals() and df_filtre is not None and not df_filtre.empty:
                        date_col = df_filtre.iloc[:, 0]
                        date_col = pd.to_datetime(date_col, errors='coerce')
                        if not date_col.isna().all():
                            date_min = date_col.min().strftime("%d/%m/%Y")
                            date_max = date_col.max().strftime("%d/%m/%Y")
                            periode_value = f"Du {date_min} au {date_max}"
                    st.text_input("Période concernée", value=periode_value, disabled=True)
                    periode = periode_value
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
            # Nettoyer les données dans PRODUCTION.xlsx
            df_production["Assureur"] = df_production["Assureur"].astype(str).str.strip().str.upper().str.replace(r'\s+', ' ', regex=True)
            df_production["Client"] = df_production["Client"].astype(str).str.strip().str.upper().str.replace(r'\s+', ' ', regex=True)

            # Filtrer selon l'assureur et le client sélectionnés
            df_production_filtered = df_production[
                (df_production["Assureur"] == nom_assureur) & 
                (df_production["Client"] == client)
            ]
            if df_production_filtered.empty:
                st.warning("⚠️ Aucune donnée dans PRODUCTION.xlsx pour l'assureur et le client sélectionnés.")
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
        df_filtre = df_detail[df_detail[colonne_client] == client]

        if df_filtre is not None and not df_filtre.empty:
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                df_filtre.to_excel(writer, index=False, sheet_name="DETAIL_FILTRÉ")
            buffer.seek(0)
            st.download_button("Télécharger DETAIL filtré", buffer.getvalue(), file_name="DETAIL_filtre.xlsx")

            montant_sinistres = df_filtre.iloc[:, 22].sum()
            ratio_sp = montant_sinistres / prime_acquise if prime_acquise > 0 else 0
            df_sin = pd.DataFrame([{
                "Id Police Ankara": police_ankara,
                "N° Police Assureur": police_assureur or "(vide)",
                "Assureur": nom_assureur,
                "Client": client,
                "Primes Émises Nettes": f"{prime_nette:,.0f}".replace(",", " "),
                "Primes Acquises": f"{prime_acquise:,.0f}".replace(",", " "),
                "Sinistres": f"{montant_sinistres:,.0f}".replace(",", " "),
                "S/P": f"{ratio_sp:.0%}"
            }])
            df_sin = df_sin[["Id Police Ankara", "N° Police Assureur", "Assureur", "Client", "Primes Émises Nettes", "Primes Acquises", "Sinistres", "S/P"]]
            st.markdown("## I - Sinistralité")
            st.dataframe(df_sin)

            if df_clause is not None:
                st.markdown("### Clause Ajustement Santé")
                ratio_sp_value = float(df_sin["S/P"].iloc[0].replace("%", "")) / 100
                df_clause_styled = df_clause.copy()
                highlight_row = None
                if tranche_min_col and tranche_max_col:
                    # Arrondi arithmétique du ratio S/P
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
        # Charger le fichier EFFECTIF.xlsx
        df_effectif = pd.read_excel(fichier_effectif)
        
        # Nettoyer les noms des colonnes (supprimer espaces, convertir en majuscules)
        df_effectif.columns = [c.strip().upper() for c in df_effectif.columns]
        
        # Vérifier la présence des colonnes nécessaires
        required_columns = ['MOIS', 'ASSUREUR', 'CLIENT', 'ADHERENT', 'CONJOINT', 'ENFANT', 'TOTAL']
        missing_columns = [col for col in required_columns if col not in df_effectif.columns]
        if missing_columns:
            st.error(f"❌ Les colonnes suivantes sont manquantes dans le fichier EFFECTIF.xlsx : {', '.join(missing_columns)}")
        else:
            # Renommer les colonnes pour correspondre au format attendu
            df_effectif = df_effectif.rename(columns={
                'CONJOINT': 'CONJOINTS',
                'ENFANT': 'ENFANTS'
            })
            
            # Normaliser les colonnes ASSUREUR et CLIENT
            df_effectif['ASSUREUR'] = df_effectif['ASSUREUR'].astype(str).str.strip().str.upper().str.replace(r'\s+', ' ', regex=True)
            df_effectif['CLIENT'] = df_effectif['CLIENT'].astype(str).str.strip().str.upper().str.replace(r'\s+', ' ', regex=True)
            
            # Appliquer les filtres sur ASSUREUR et CLIENT
            df_effectif_filtered = df_effectif[
                (df_effectif['ASSUREUR'] == nom_assureur) &
                (df_effectif['CLIENT'] == client)
            ]
            
            if df_effectif_filtered.empty:
                st.warning("⚠️ Aucune donnée dans EFFECTIF.xlsx pour l'assureur et le client sélectionnés.")
            else:
                # Convertir la colonne MOIS en datetime
                df_effectif_filtered["MOIS"] = pd.to_datetime(df_effectif_filtered["MOIS"], format="%d/%m/%Y", errors="coerce")
                
                # Trier par ordre chronologique
                df_effectif_filtered = df_effectif_filtered.sort_values(by="MOIS", ascending=True)
                
                # Formater les mois en français
                mois_fr = {
                    "January": "Janvier", "February": "Février", "March": "Mars", "April": "Avril", "May": "Mai",
                    "June": "Juin", "July": "Juillet", "August": "Août", "September": "Septembre", "October": "Octobre",
                    "November": "Novembre", "December": "Décembre"
                }
                df_effectif_filtered["MOIS"] = df_effectif_filtered["MOIS"].dt.strftime("%B %Y").replace(mois_fr, regex=True)
                
                # Afficher uniquement les colonnes demandées
                display_columns = ["MOIS", "ADHERENT", "CONJOINTS", "ENFANTS", "TOTAL"]
                df_effectif_display = df_effectif_filtered[display_columns]
                st.dataframe(df_effectif_display)

                # Créer le graphique d'évolution des effectifs
                fig, ax = plt.subplots(figsize=(10, 5))
                colors = ['#279244', '#f77f00', '#ff6f61', '#2a9d8f']
                for i, col in enumerate(["ADHERENT", "CONJOINTS", "ENFANTS", "TOTAL"]):
                    if col in df_effectif_filtered.columns:
                        ax.plot(df_effectif_filtered["MOIS"], df_effectif_filtered[col], marker='o', label=col.title(), color=colors[i % len(colors)])
                ax.set_title("Évolution des effectifs")
                ax.set_xlabel("Mois")
                ax.set_ylabel("Effectifs")
                ax.legend()
                plt.xticks(rotation=45)
                st.pyplot(fig)
                
                # Sauvegarder le graphique pour le PDF
                evol_effectif_path = os.path.join(tempfile.gettempdir(), "graph_effectif.png")
                fig.savefig(evol_effectif_path, bbox_inches='tight')
                plt.close(fig)
                
                # Mettre à jour df_effectif
                df_effectif = df_effectif_filtered
    except Exception as e:
        st.error(f"❌ Erreur lors du chargement des effectifs : {e}")

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

            montants = df_filtre.groupby("FILIATION")[col_montant].sum().rename("Montant couvert")
            tableau = pd.concat([patients_counts, effectifs, montants], axis=1)
            tableau["Taux d'utilisation"] = tableau["Nombre de patients"] / tableau["Effectif Total"]
            total_montant = tableau["Montant couvert"].sum()
            tableau["Part de consommation"] = tableau["Montant couvert"] / total_montant

            total = pd.DataFrame({
                "Nombre de patients": [tableau["Nombre de patients"].sum()],
                "Effectif Total": [tableau["Effectif Total"].sum()],
                "Taux d'utilisation": [tableau["Nombre de patients"].sum() / tableau["Effectif Total"].sum()],
                "Montant couvert": [total_montant],
                "Part de consommation": [1.0]
            }, index=["Total général"])

            tableau_final = pd.concat([tableau, total])
            cols = ["Nombre de patients", "Effectif Total", "Taux d'utilisation", "Montant couvert", "Part de consommation"]
            ordre_filiation = ["ASSURÉ PRINCIPAL", "CONJOINT", "ENFANT", "Total général"]
            tableau_final = tableau_final.reindex(ordre_filiation)[cols]
            tableau_final["Taux d'utilisation"] = (tableau_final["Taux d'utilisation"].fillna(0) * 100).round(0).astype(int).astype(str) + "%"
            tableau_final["Nombre de patients"] = tableau_final["Nombre de patients"].fillna(0).round(0).astype(int)
            tableau_final["Part de consommation"] = (tableau_final["Part de consommation"].fillna(0) * 100).round(0).astype(int).astype(str) + "%"
            tableau_final["Montant couvert"] = tableau_final["Montant couvert"].fillna(0).apply(lambda x: f"{int(x):,}".replace(",", " "))
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

        # Convertir la colonne des dates
        df_mensuel = df_filtre.copy()
        df_mensuel["DATE"] = pd.to_datetime(df_mensuel[col_date], errors="coerce")
        if df_mensuel["DATE"].isna().all():
            st.error("❌ La colonne des dates (colonne 1) contient des valeurs invalides ou est vide.")
            raise ValueError("Dates invalides pour les consommations mensuelles.")

        # Filtrer sur les rejets si la colonne existe et contient des données valides
        if col_rejet is not None and col_rejet in df_mensuel.columns and not df_mensuel[col_rejet].isna().all():
            df_mensuel[col_rejet] = pd.to_numeric(df_mensuel[col_rejet], errors='coerce')
            df_mensuel = df_mensuel[df_mensuel[col_rejet].notna()]
        else:
            st.warning("⚠️ La colonne des rejets (colonne 24) est absente ou vide. Traitement sans filtrage des rejets.")

        if df_mensuel.empty:
            st.warning("⚠️ Aucune donnée disponible pour les consommations mensuelles après filtrage.")
        else:
            df_mensuel = df_mensuel.sort_values(by="DATE", ascending=True)
            mois_fr = {
                "January": "Janvier", "February": "Février", "March": "Mars", "April": "Avril", "May": "Mai",
                "June": "Juin", "July": "Juillet", "August": "Août", "September": "Septembre", "October": "Octobre",
                "November": "Novembre", "December": "Décembre"
            }
            df_mensuel["MOIS"] = df_mensuel["DATE"].dt.strftime("%B %Y").replace(mois_fr, regex=True)

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
                
                df_mensuel_grouped = df_mensuel_grouped.sort_values(by="MOIS", key=lambda x: pd.to_datetime(x, format="%B %Y", errors='coerce'))

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
        fig, ax = plt.subplots(figsize=(8, 6))
        total_couvert = df_graph["Couvert"].sum()
        labels = [f"{spec}" for spec, montant in zip(df_graph["Spécialité"], df_graph["Couvert"])]
        colors = ['#279244', '#f77f00', '#2a9d8f', '#ff6f61', '#264653']
        ax.pie(df_graph["Couvert"], labels=labels, autopct='%1.1f%%', startangle=90, explode=[0.05] * len(df_graph), colors=colors, textprops={'fontsize': 8}, labeldistance=1.1, pctdistance=0.85)
        ax.set_title("Répartition par spécialité")
        st.pyplot(fig)
        graph_path = os.path.join(tempfile.gettempdir(), "graph_specialite.png")
        fig.savefig(graph_path, bbox_inches='tight')
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
        }).rename(columns={col_prestataire: "Nombre de Sinistres", col_couvert: "Couvert"}).reset_index()
        df_prestataires.columns = ["PRESTATAIRE", "VILLE", "COMMUNE", "Nombre de Sinistres", "Couvert"]
        df_prestataires = df_prestataires.sort_values(by="Couvert", ascending=False)
        total_covered = df_prestataires["Couvert"].sum()
        df_prestataires["Proportion"] = (df_prestataires["Couvert"] / total_covered * 100).round(0).astype(int).astype(str) + "%"
        df_prestataires["Ordre"] = range(1, len(df_prestataires) + 1)
        df_prestataires = df_prestataires[["Ordre", "PRESTATAIRE", "VILLE", "COMMUNE", "Nombre de Sinistres", "Couvert", "Proportion"]]
        df_prestataires["Couvert"] = df_prestataires["Couvert"].apply(lambda x: f"{int(x):,}".replace(",", " "))
        st.dataframe(df_prestataires.head(10))
    except Exception as e:
        st.error(f"❌ Erreur lors du traitement des prestataires : {e}")

# Section 8 : Top des Familles de Consommateurs
if sinistralite_ok and df_filtre is not None and not df_filtre.empty:
    st.markdown("## VII - Top des Familles de Consommateurs")
    try:
        col_carte = df_filtre.columns[9]
        col_nom = df_filtre.columns[11]
        col_couvert = df_filtre.columns[22]

        df_familles = df_filtre.groupby([col_carte, col_nom]).agg({
            col_carte: "count",
            col_couvert: "sum"
        }).rename(columns={col_carte: "Nombre d’actes", col_couvert: "Couvert"}).reset_index()
        df_familles.columns = ["N° de Famille", "Assuré Principal", "Nombre d’actes", "Couvert"]
        df_familles = df_familles.sort_values(by="Couvert", ascending=False)
        total_covered = df_familles["Couvert"].sum()
        df_familles["Proportion"] = (df_familles["Couvert"] / total_covered * 100).round(0).astype(int).astype(str) + "%"
        df_familles["Ordre"] = range(1, len(df_familles) + 1)
        df_familles = df_familles[["Ordre", "N° de Famille", "Assuré Principal", "Nombre d’actes", "Couvert", "Proportion"]]
        df_familles["Couvert"] = df_familles["Couvert"].apply(lambda x: f"{int(x):,}".replace(",", " "))
        st.dataframe(df_familles.head(10))
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
            os.remove(path)

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
                        self.set_font("Arial", "I", 8)
                        self.set_text_color(255, 255, 255)
                        page_text = f"Statistiques {nom_assureur}_{client} - Page {self.page_no() - 2} / {self.total_pages}"
                        self.cell(0, 10, page_text, align="C")

            def clean_text(text):
                if isinstance(text, pd.DataFrame):
                    replacements = {
                        '’': "'", 'Œ': 'OE', 'œ': 'oe', '…': '...', 'é': 'e', 'è': 'e', 'ê': 'e', 'ë': 'e',
                        'à': 'a', 'â': 'a', 'ä': 'a', 'î': 'i', 'ï': 'i', 'ô': 'o', 'ö': 'o', 'ù': 'u',
                        'û': 'u', 'ü': 'u', 'ç': 'c', '\u2013': '-', '\u2014': '-', '\u2019': "'"
                    }
                    return text.replace(replacements, regex=True).fillna('')
                elif isinstance(text, str):
                    replacements = {
                        '’': "'", 'Œ': 'OE', 'œ': 'oe', '…': '...', 'é': 'e', 'è': 'e', 'ê': 'e', 'ë': 'e',
                        'à': 'a', 'â': 'a', 'ä': 'a', 'î': 'i', 'ï': 'i', 'ô': 'o', 'ö': 'o', 'ù': 'u',
                        'û': 'u', 'ü': 'u', 'ç': 'c', '\u2013': '-', '\u2014': '-', '\u2019': "'"
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

                pdf.set_font("Arial", '', 9)
                pdf.set_text_color(0, 0, 0)
                page_width = float(pdf.w - 2 * pdf.l_margin)
                line_height = 5.0

                # Définir les largeurs de colonnes en fonction du nombre de colonnes du DataFrame
                if is_prestataires:
                    col_widths = [15.0, 50.0, 20.0, 25.0, 20.0, 30.0, 30.0]
                elif is_familles:
                    col_widths = [15.0, 30.0, 50.0, 30.0, 30.0, 30.0]
                else:
                    # Largeur par défaut : diviser la largeur de la page par le nombre de colonnes
                    num_cols = len(df.columns)
                    col_widths = [page_width / num_cols] * num_cols

                # Ajuster les largeurs si elles dépassent la largeur de la page
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

                pdf.set_fill_color(39, 146, 68)  # Couleur #279244
                pdf.set_text_color(255, 255, 255)
                pdf.set_font("Arial", 'B', 9)
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
                pdf.set_font("Arial", '', 9)
                table_y_start = float(pdf.get_y())

                for i, (_, row) in enumerate(df_display.iterrows()):
                    row_height = line_height * float(max_lines_per_row[i])
                    current_y = float(pdf.get_y())
                    page_height = float(pdf.h)
                    bottom_margin = float(pdf.b_margin)

                    if current_y + row_height > page_height - bottom_margin - 15:
                        pdf.line(pdf.l_margin, table_y_start, pdf.l_margin, current_y)
                        pdf.line(pdf.l_margin + table_width, table_y_start, pdf.l_margin + table_width, current_y)
                        pdf.add_page()
                        table_y_start = float(pdf.get_y())
                        pdf.set_fill_color(39, 146, 68)  # Couleur #279244
                        pdf.set_text_color(255, 255, 255)
                        pdf.set_font("Arial", 'B', 9)
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
                        pdf.set_font("Arial", '', 9)

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
                    ratio_sp_value = float(df_sin["S/P"].iloc[0].replace("%", "")) / 100
                    # Arrondi arithmétique du ratio S/P
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
                    page_numbers.append(add_table_section(temp_pdf, "Clause Ajustement Santé", df_clause, highlight_row=highlight_row, new_page=False))
                else:
                    page_numbers.append(0)
            if df_effectif is not None:
                page_numbers.append(add_table_section(temp_pdf, "Section II - Évolution des effectifs", df_effectif_display))  # Utiliser df_effectif_display
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
            pdf.cell(0, 20, "STATISTIQUES DE GESTION SANTE", ln=True, align="C")
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

            pdf.set_xy(info_box_x + 10, start_y)
            pdf.set_font("Arial", 'B', 12)
            pdf.cell(label_width, line_height, "Assureur : ", align='L')
            pdf.set_font("Arial", '', 12)
            pdf.set_x(info_box_x + 10 + label_width)
            pdf.multi_cell(value_width, line_height, clean_text(nom_assureur), align='L')

            pdf.set_xy(info_box_x + 10, start_y + line_height)
            pdf.set_font("Arial", 'B', 12)
            pdf.cell(label_width, line_height, "Client : ", align='L')
            pdf.set_font("Arial", '', 12)
            pdf.set_x(info_box_x + 10 + label_width)
            pdf.multi_cell(value_width, line_height, clean_text(client), align='L')

            pdf.set_xy(info_box_x + 10, start_y + 2 * line_height)
            pdf.set_font("Arial", 'B', 12)
            pdf.cell(label_width, line_height, "Période : ", align='L')
            pdf.set_font("Arial", '', 12)
            pdf.set_x(info_box_x + 10 + label_width)
            pdf.multi_cell(value_width, line_height, clean_text(periode), align='L')

            pdf.set_xy(info_box_x + 10, start_y + 3 * line_height)
            pdf.set_font("Arial", 'B', 12)
            pdf.cell(label_width, line_height, "Date d'édition : ", align='L')
            pdf.set_font("Arial", '', 12)
            pdf.set_x(info_box_x + 10 + label_width)
            pdf.multi_cell(value_width, line_height, datetime.now().strftime("%d/%m/%Y"), align='L')

            pdf.set_y(info_box_y + info_box_height + 10)
            if logo_assureur_path and os.path.exists(logo_assureur_path):
                pdf.image(logo_assureur_path, x=(pdf.w - 50) / 2, y=float(pdf.get_y()), w=50)
                pdf.ln(60)

            pdf.set_font("Arial", 'I', 10)
            pdf.set_text_color(54, 69, 79)
            pdf.multi_cell(0, 5, clean_text("Ankara Services, Abidjan – Plateau, Avenue Noguès Immeuble Borija, Tel :+225 25 20 01 31 05/06\n"
                                "Société Anonyme avec Conseil d’Administration au Capital de 10.000.000 FCFA - 01 BP 1194 ABJ 01\n"
                                "RCCM CI- ABJ-03-2021-B14-00020-NCC 2110076 V-Banque : STANBIC CI198 01001 918000005921 84\n"
                                "www.ankaraservives.com"), align="C")

            pdf.add_page()
            pdf.set_font("Arial", 'B', 16)
            pdf.set_text_color(39, 146, 68)
            pdf.cell(0, 10, "SOMMAIRE", ln=True, align="C")
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
                dots = "." * int((right_align_x - pdf.get_string_width(title) - pdf.get_string_width(f"Page {page}") - left_shift - 10) / pdf.get_string_width("."))
                pdf.set_x(left_shift)
                pdf.cell(0, 8, f"{title}{dots}Page {page}", ln=True, align="L")
                pdf.ln(8)

            if df_sin is not None:
                add_table_section(pdf, "Section I - Sinistralité", df_sin)
                if df_clause is not None:
                    ratio_sp_value = float(df_sin["S/P"].iloc[0].replace("%", "")) / 100
                    # Arrondi arithmétique du ratio S/P
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
                add_table_section(pdf, "Section II - Évolution des effectifs", df_effectif_display)  # Utiliser df_effectif_display
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

            pdf_output = BytesIO()
            pdf_bytes = pdf.output(dest='S').encode('latin1')
            pdf_output.write(pdf_bytes)
            pdf_output.seek(0)
            st.download_button("Télécharger le PDF", pdf_output, f"{nom_assureur}_{client}_rapport_sante.pdf", "application/pdf")

        st.success("✅ PDF généré avec succès !")
        cleanup_temp_files()

    except Exception as e:
        st.error(f"❌ Erreur lors de la génération du PDF : {e}")
        import traceback
        st.write("Traceback complet de l'erreur :", traceback.format_exc())
