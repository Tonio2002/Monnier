from fastapi import FastAPI

app = FastAPI()

@app.get("/")
def home():
    return {"message": "API fonctionne !"}

@app.post("/format-excel/")
async def format_excel():
    return {"message": "Traitement du fichier Excel en cours"}


    # Charger le fichier Excel
    xls = pd.ExcelFile(file_location)
    df = pd.read_excel(xls, sheet_name="Report1")

    # Étape 1: Ajouter la colonne "Site Name"
    site_names = []
    current_site = None
    for index, row in df.iterrows():
        if isinstance(row.iloc[0], str) and " - " in row.iloc[0]:
            current_site = row.iloc[0]  
        site_names.append(current_site)

    df.insert(0, "Site Name", site_names)

    # Étape 2: Supprimer la ligne 4 (index 3)
    df = df.drop(index=3).reset_index(drop=True)

    # Étape 3: Ajouter une colonne vide en B
    df.insert(1, "New Column", "")

    # Étape 4: Déplacer le contenu des parenthèses en colonne B
    df["New Column"] = df["Retail Sales"].apply(lambda x: re.search(r"\((.*?)\)", x).group(1) if isinstance(x, str) and "(" in x else "")

    # Étape 5: Supprimer les parenthèses et leur contenu en colonne originale
    df["Retail Sales"] = df["Retail Sales"].str.replace(r"\s*\(.*?\)", "", regex=True)

    # Étape 6: Sauvegarder le fichier transformé
    output_file_path = "/tmp/RetailSales_Final.xlsx"
    df.to_excel(output_file_path, index=False)

    return {"download_url": "http://localhost:8000/download/RetailSales_Final.xlsx"}
