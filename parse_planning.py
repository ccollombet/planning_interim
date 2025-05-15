import pandas as pd
from datetime import datetime, timedelta

# Charger le fichier Excel
fichier = "PLANNING_Montplaisant.xlsx"
df = pd.read_excel(fichier, header=None)

# Dates des colonnes (à partir de la 4e colonne : index 3)
dates = df.iloc[0, 3:].tolist()
dates_nettoyees = []
for d in dates:
    try:
        jour = d.split('\n')[0].strip(" 0123456789")
        numero = ''.join(filter(str.isdigit, d))
        date = f"{numero}/04/2024"  # Mois d'avril
        dates_nettoyees.append(date)
    except:
        dates_nettoyees.append("")

# Extraire les lignes à partir de la ligne 9
resultats = []

for i in range(9, df.shape[0]):
    ligne = df.iloc[i]
    nom_prenom = ligne[0]
    structure = ligne[1]

    # Nettoyage nom / prénom
    if isinstance(nom_prenom, str) and "\n" in nom_prenom:
        nom, prenom = nom_prenom.split("\n")
        nom = nom.strip().upper()
        prenom = prenom.strip().capitalize()
    else:
        continue  # sauter les lignes vides

    for j, valeur in enumerate(ligne[3:], start=3):
        if pd.notna(valeur) and isinstance(valeur, (int, float)):
            heure_float = float(valeur)
            heure_debut = int(heure_float)
            minutes_debut = int((heure_float - heure_debut) * 60)
            heure_debut_str = f"{heure_debut:02d}:{minutes_debut:02d}"

            # Supposons une durée de 7h pour test (on pourrait améliorer avec les paires)
            heure_fin_dt = datetime.strptime(heure_debut_str, "%H:%M") + timedelta(hours=7)
            heure_fin_str = heure_fin_dt.strftime("%H:%M")

            date_jour = dates_nettoyees[j - 3] if j - 3 < len(dates_nettoyees) else "??/04/2024"

            resultats.append({
                "Nom": nom,
                "Prénom": prenom,
                "Poste": "Accompagnant éducatif et soc",
                "Structure(s)": "MAS LES ACACIAS",
                "Date": date_jour,
                "Heure de début de travail": heure_debut_str,
                "Temps de coupure": "00:20",
                "Heure de fin de travail": heure_fin_str,
                "Temps travaillé": "7,00",  # À ajuster si tu veux le calculer
                "Personne remplacée": "",
                "Motif": "",
                "Info complémentaire sur le motif": "",
                "Unite(s)": "JARDIN",
                "Précisez si coefficient EXTERNAT": "EXTERNAT",
                "Commentaires": ""
            })

# Export CSV
df_export = pd.DataFrame(resultats)
df_export.to_csv("export_interimaires.csv", index=False, sep=';')
print("✅ Fichier exporté : export_interimaires.csv")
