import streamlit as st
import pandas as pd
import docx
from datetime import datetime
import base64

# Fonction pour extraire les données du fichier Word
def extract_consultant_data(doc):
    data = []
    entries = doc.paragraphs
    current = {}

    for para in entries:
        text = para.text.strip()
        if not text:
            continue

        if '—' in text:
            if current:
                data.append(current)
                current = {}
            parts = text.split('—')
            current["Consultant"] = parts[0].strip()
        elif "Disponibilité" in text:
            current["Disponibilité (%)"] = text.split(":")[-1].strip()
        elif "Actuellement" in text or "Envisagé" in text or "Staffing" in text:
            current["Projet en cours / envisagé"] = text.strip()
        elif "https://" in text:
            current["Lien CV"] = text.strip()
        elif any(date in text for date in ["24/03", "07/04", "11/04"]):
            current["Date début interco"] = text.split(":")[0].strip()

    if current:
        data.append(current)

    return pd.DataFrame(data)

# Fonction pour calculer la durée d'intercontrat
def calculate_duration(date_str):
    try:
        date_ref = datetime.strptime(date_str, "%d/%m")
        date_ref = date_ref.replace(year=datetime.now().year)
        delta = datetime.now() - date_ref
        return delta.days
    except:
        return None

# Style coloré selon durée
def highlight_duration(val):
    if pd.isna(val):
        return ''
    if val >= 21:
        return 'background-color: #ff4d4d; color: white'  # rouge vif
    elif val >= 14:
        return 'background-color: #ffa64d; color: black'  # orange
    elif val >= 7:
        return 'background-color: #ffff66; color: black'  # jaune
    else:
        return 'background-color: #b3ffb3; color: black'  # vert clair

# Interface Streamlit
st.set_page_config(page_title="Suivi Intercontrat", layout="wide")
st.title("🧮 Générateur de Tableau Intercontrat")
st.markdown("""
<style>
    .main { background-color: #f0f2f6; }
    .css-1v0mbdj, .css-1d391kg { color: #333333; font-size: 16px; }
</style>
""", unsafe_allow_html=True)

uploaded_file = st.file_uploader("📄 Charge ton fichier Word (.docx)", type="docx")

if uploaded_file:
    doc = docx.Document(uploaded_file)
    df = extract_consultant_data(doc)

    df["Durée interco (jours)"] = df["Date début interco"].apply(calculate_duration)

    # Affichage visuel
    st.subheader("🔍 Résultat du tableau de suivi")
    styled_df = df.style.applymap(highlight_duration, subset=["Durée interco (jours)"])
    st.dataframe(styled_df, use_container_width=True)

    # Filtres dynamiques
    with st.expander("🔎 Filtres avancés"):
        dispo_filter = st.selectbox("Filtrer par disponibilité", options=["Tous"] + list(df["Disponibilité (%)"].dropna().unique()))
        if dispo_filter != "Tous":
            df = df[df["Disponibilité (%)"] == dispo_filter]

    # Export Excel
    if st.button("📥 Exporter en Excel"):
        export_df = df.copy()
        export_df.to_excel("interco_tableau.xlsx", index=False)
        with open("interco_tableau.xlsx", "rb") as f:
            st.download_button("Télécharger le fichier Excel", f, file_name="Interco_Tableau.xlsx")
