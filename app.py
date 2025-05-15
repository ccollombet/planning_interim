import streamlit as st
import pandas as pd
from datetime import datetime
import re
import io



# 🔒 Mot de passe défini par toi
PASSWORD = "Celine01$"

def check_password():
    """Affiche une zone de saisie pour le mot de passe et bloque l’accès si incorrect."""
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False

    if not st.session_state.authenticated:
        with st.form("Login"):
            pwd = st.text_input("🔐 Entrez le mot de passe", type="password")
            submitted = st.form_submit_button("Se connecter")
            if submitted and pwd == PASSWORD:
                st.session_state.authenticated = True
            elif submitted:
                st.error("Mot de passe incorrect")

    return st.session_state.authenticated

# Blocage tant que mot de passe non bon
if not check_password():
    st.stop()


# === TABLE DE CORRESPONDANCE HORAIRE ===
horaire_mapping = {
    "700-1430": (7.5, 0),
    "1400-2130": (7.5, 0),
    "800-2000": (12, 0),
    "700-1700": (10, 0),
    "730-1200/1700-2030": (8, 5),
    "1200-2130": (9.5, 0)
}

jours_fr = {
    "Monday": "Lundi", "Tuesday": "Mardi", "Wednesday": "Mercredi",
    "Thursday": "Jeudi", "Friday": "Vendredi",
    "Saturday": "Samedi", "Sunday": "Dimanche"
}

# === FONCTIONS ===
def normalize_horaire(horaire):
    if pd.isna(horaire): return ""
    horaire = str(horaire).strip().upper()
    horaire = horaire.replace(',', '.').replace(';', '.').replace(' ', '')
    horaire = horaire.replace('H', '')
    horaire = re.sub(r'[^0-9./\-]', '', horaire)
    horaire = horaire.replace('.', '')
    return horaire

def extract_hours(horaire):
    horaire = normalize_horaire(horaire)
    segments = re.split(r"/", horaire)
    start = segments[0].split("-")[0]
    end = segments[-1].split("-")[-1]
    return start, end

def get_travail_coupure(horaire_normalise):
    horaire_normalise = horaire_normalise.replace(" ", "")
    for key in horaire_mapping:
        key_normalise = normalize_horaire(key).replace(" ", "")
        if horaire_normalise == key_normalise:
            return horaire_mapping[key]
    return 0, 0

def horaire_to_hhmm(heure_str):
    if not heure_str or pd.isna(heure_str): return ""
    match = re.match(r"(\d{1,2})(\d{2})", heure_str)
    if match:
        h = int(match.group(1))
        m = int(match.group(2))
        return f"{h:02d}:{m:02d}"
    return ""

def decimal_to_hhmm(decimal_val):
    try:
        total_minutes = round(float(decimal_val) * 60)
        h = total_minutes // 60
        m = total_minutes % 60
        return f"{h:02d}:{m:02d}"
    except:
        return "00:00"

# === INTERFACE STREAMLIT ===
st.title("🧾 Convertisseur de planning intérimaires en CSV")

uploaded_file = st.file_uploader("Déposez un fichier Excel du planning", type="xlsx")

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file, sheet_name=0, header=None)
    results = []

    for col in range(2, df.shape[1]):
        jour = df.iloc[0, col]
        try:
            day = int(jour)
            date_obj = datetime(2025, 6, day)
            jour_fr = jours_fr[date_obj.strftime("%A")]
            date_str = f"{jour_fr} {date_obj.strftime('%d/%m/%Y')}"
        except:
            continue

        for base in [2, 5, 8, 11]:
            nom = str(df.iloc[base, col]).strip()
            agence = str(df.iloc[base + 1, col]).strip()
            horaire = str(df.iloc[base - 1, col]).strip()

            if nom.upper() == "INTERIMAIRE":
                heure_debut_raw, heure_fin_raw = extract_hours(horaire)
                horaire_normalise = normalize_horaire(horaire)
                t_travail, t_coupure = get_travail_coupure(horaire_normalise)

                results.append({
                    "Nom": "Interimaire",
                    "Prénom": "Interimaire",
                    "Poste": "Accompagnant éducatif et soc",
                    "Stucture(s)": "Mas Montaines",
                    "Date": date_str,
                    "Heure de début de travail": horaire_to_hhmm(heure_debut_raw),
                    "Temps de coupure": decimal_to_hhmm(t_coupure),
                    "Heure de fin de travail": horaire_to_hhmm(heure_fin_raw),
                    "Temps travaillé": decimal_to_hhmm(t_travail),
                    "Personne remplacée": "",
                    "Motif": "",
                    "Info complémentaire sur le motif": "",
                    "Unite(s)": "",
                    "Précisez si coefficient EXTERNAT": "",
                    "Commentaires": ""
                })

    df_result = pd.DataFrame(results)
    csv_buffer = io.StringIO()
    df_result.to_csv(csv_buffer, sep=';', index=False, encoding='utf-8')
    csv_bytes = csv_buffer.getvalue().encode('utf-8')

    st.success("✅ Conversion terminée avec succès !")
    st.download_button(
        label="📥 Télécharger le fichier CSV",
        data=csv_bytes,
        file_name="planning_interimaires_converti.csv",
        mime="text/csv"
    )
