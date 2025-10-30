# Automatiser-la-g-n-ration-de-rapports-Excel-avec-Python
Automatiser la génération de rapports Excel avec Python 
import pandas as pd
from datetime import datetime

# --- 1. Génération de données simulées ---
data = {
    "Produit": ["PC Portable", "Clavier", "Souris", "Écran", "Imprimante"],
    "Quantité vendue": [25, 45, 67, 23, 12],
    "Prix unitaire ($)": [899, 49, 29, 199, 149]
}

df = pd.DataFrame(data)
df["Total ($)"] = df["Quantité vendue"] * df["Prix unitaire ($)"]

# --- 2. Calcul de statistiques ---
total_global = df["Total ($)"].sum()
moyenne_prix = df["Prix unitaire ($)"].mean()

# --- 3. Ajout d’un résumé ---
resume = pd.DataFrame({
    "Description": ["Total global des ventes", "Prix unitaire moyen"],
    "Valeur": [total_global, round(moyenne_prix, 2)]
})

# --- 4. Enregistrement dans un fichier Excel ---
date_rapport = datetime.now().strftime("%Y-%m-%d")
nom_fichier = f"rapport_ventes_{date_rapport}.xlsx"

with pd.ExcelWriter(nom_fichier, engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name="Détails des ventes", index=False)
    resume.to_excel(writer, sheet_name="Résumé", index=False)

print(f"✅ Rapport généré avec succès : {nom_fichier}")
