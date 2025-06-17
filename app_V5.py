import streamlit as st
import pandas as pd
from datetime import datetime
import re
import io

# === CONFIGURATION ===
PASSWORD = "Celine01$"

horaire_mapping = {
    "700-1430": (7.5, 0),
    "1400-2130": (7.5, 0),
    "800-2000": (12, 0),
    "700-1700": (10, 0),
    "730-1200/1700-2030": (8, 5),
    "1200-2130": (9.5, 0),
    "1000-2000": (10, 0),
    "1400-2030": (6.5, 0),
    "1400-2130": (7.5, 0),
}

jours_fr = {
    "Monday": "Lundi", "Tuesday": "Mardi", "Wednesday": "Mercredi",
    "Thursday": "Jeudi", "Friday": "Vendredi",
    "Saturday": "Samedi", "Sunday": "Dimanche"
}

# === UTILITAIRES ===
def normalize_horaire(horaire):
    if pd.isna(horaire): return ""
    horaire = str(horaire).strip().upper()
    horaire = horaire.replace(',', '.').replace(';', '.').replace(' ', '').replace('H', '')
    horaire = re.sub(r'[^0-9./\-]', '', horaire)
    return horaire.replace('.', '')

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

# === AUTHENTIFICATION ===
def check_password():
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

# === LOGIQUE DE TRAITEMENT ===
def traiter_fichier(df, structure_name, groupe_prefix):
    results = []
    groupe_courant = None

    for row in range(1, df.shape[0]):
        valeur_colA = df.iloc[row, 0]
        valeur_colB = str(df.iloc[row, 1]).strip().upper()

        #if pd.notna(valeur_colA) and valeur_colA != '':
        #    #groupe_courant = str(valeur_colA)
        #    if float(valeur_colA).is_integer():
        #        groupe_courant = int(valeur_colA)
        #    else:
        #        groupe_courant = str(valeur_colA)
        if pd.notna(valeur_colA) and valeur_colA != '':
            try:
                if float(valeur_colA).is_integer():
                    groupe_courant = str(int(float(valeur_colA)))  # Ex: 1.0 → "1"
                else:
                    groupe_courant = str(valeur_colA)
            except ValueError:
                # Si ce n'est pas un nombre, on garde tel quel (ex: "A", "B")
                groupe_courant = str(valeur_colA).strip()



        if valeur_colB == "NOM" and groupe_courant:
            for col in range(2, df.shape[1]):
                nom = str(df.iloc[row, col]).strip()
                agence = str(df.iloc[row + 1, col]).strip()
                horaire = str(df.iloc[row - 1, col]).strip()

                try:
                    day = int(df.iloc[0, col])
                    date_obj = datetime(2025, 6, day)
                    jour_fr = jours_fr[date_obj.strftime("%A")]
                    date_str = f"{jour_fr} {date_obj.strftime('%d/%m/%Y')}"
                except:
                    continue

                if nom.upper() == "INTERIMAIRE":
                    heure_debut_raw, heure_fin_raw = extract_hours(horaire)
                    horaire_normalise = normalize_horaire(horaire)
                    t_travail, t_coupure = get_travail_coupure(horaire_normalise)

                    results.append({
                        "Nom": "Interimaire",
                        "Prénom": "Interimaire",
                        "Poste": "Accompagnant éducatif et soc",
                        "Stucture(s)": structure_name,
                        "Date": date_str,
                        "Heure de début de travail": horaire_to_hhmm(heure_debut_raw),
                        "Temps de coupure": decimal_to_hhmm(t_coupure),
                        "Heure de fin de travail": horaire_to_hhmm(heure_fin_raw),
                        "Temps travaillé": decimal_to_hhmm(t_travail),
                        "Personne remplacée": "",
                        "Motif": "",
                        #"Info complémentaire sur le motif": f"Groupe {groupe_prefix}{groupe_courant}",
                        "Info complémentaire sur le motif": f"Groupe {groupe_courant}",
                        "Unite(s)": "",
                        "Précisez si coefficient EXTERNAT": "",
                        "Commentaires": ""
                    })
    return pd.DataFrame(results)

# === INTERFACE ===
if not check_password():
    st.stop()

st.title("🧾 Convertisseur de planning en CSV avec choix de la MAS")
choix_mas = st.radio("Choisissez la MAS :", ["Montaines", "Montplaisant"])

uploaded_file = st.file_uploader("Déposez un fichier Excel du planning", type="xlsx")

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file, sheet_name=0, header=None)

    if choix_mas == "Montaines":
        df_result = traiter_fichier(df, "Mas Montaines", groupe_prefix="")
    else:
        df_result = traiter_fichier(df, "Mas Montplaisant", groupe_prefix="")

    if not df_result.empty:
        csv_buffer = io.StringIO()
        df_result.to_csv(csv_buffer, sep=';', index=False, encoding='utf-8-sig')
        csv_bytes = csv_buffer.getvalue().encode('utf-8-sig')

        st.success("✅ Conversion terminée avec succès !")
        st.download_button(
            label="📥 Télécharger le fichier CSV",
            data=csv_bytes,
            file_name="planning_interimaires_converti.csv",
            mime="text/csv"
        )
    else:
        st.warning("Le fichier n'a pas produit de lignes exploitables.")
