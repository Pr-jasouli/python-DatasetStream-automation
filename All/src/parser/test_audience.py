import tkinter as tk

import pandas as pd

from ui.audience_tab import AudienceTab

def setup_ui(audience_tab):
    """Simulates setting up user input and application paths."""
    fields = {
        'references_month': "12",
        'references_year': "2024",
        'target_start_year': "2025",
        'target_end_year': "2026",
    }
    for field, value in fields.items():
        getattr(audience_tab, field).delete(0, tk.END)
        getattr(audience_tab, field).insert(0, value)
    audience_tab.df = "C:\\Users\\dbotton\\OneDrive - Orange Business (ex-B&D)\\Desktop\\python-DatasetStream-automation\\All\\inputs\\data audience creation .xlsx"
    audience_tab.output_path.insert(0, "C:\\Users\\dbotton\\OneDrive - Orange Business (ex-B&D)\\Desktop\\python-DatasetStream-automation\\All\\outputs")
# logging.basicConfig(level=logging.DEBUG)

def load_excel_data(filepath):
    """Loads Excel data into a pandas DataFrame."""
    return pd.read_excel(filepath)

def get_column_names(df):
    """Dynamically fetches column names based on content criteria."""
    identifiers = ['DATA_TYPE', 'PROD_NUM', 'BUS_CHANL_NUM']
    viewing_columns = [col for col in df.columns if 'VIEWING' in col]
    return identifiers, viewing_columns

def duplicate_data(df, ref_year, ref_month, target_years, identifiers):
    """Duplicates data for specified target years resetting non-key columns."""
    months = range(1, int(ref_month) + 1)
    mask = (df['PERIOD_YEAR'] == int(ref_year)) & (df['PERIOD_MONTH'].isin(months))
    data_to_duplicate = df.loc[mask]

    zero_columns = [col for col in df.columns if col not in ['PERIOD_YEAR', 'PERIOD_MONTH'] + identifiers]
    new_rows = []
    for year in target_years:
        for _, row in data_to_duplicate.iterrows():
            new_row = row.copy()
            new_row['PERIOD_YEAR'] = year
            new_row.update({col: 0 for col in zero_columns})
            new_rows.append(new_row)

    if new_rows:
        new_rows_df = pd.DataFrame(new_rows)
        return pd.concat([df, new_rows_df], ignore_index=True)
    else:
        return df


def adjust_viewing_columns(df, ref_year, identifiers, viewing_columns):
    """
    Définition des colonnes somme : Les variables sum_eop_2024 et sum_eop_2025 représentent les noms des colonnes dans le DataFrame qui contiennent les valeurs à utiliser pour le calcul des ratios.
    """
    sum_eop_2024 = 'sum_eop_vol_2024'
    sum_eop_2025 = 'sum_eop_vol_2025'

    """
    Itération sur les lignes ciblées : La boucle itère sur toutes les lignes du DataFrame df où l'année (PERIOD_YEAR) est supérieure à l'année de référence (ref_year). Cela permet de cibler uniquement les lignes des années futures par rapport à l'année de référence.
    """
    for index, row in df[df['PERIOD_YEAR'] > int(ref_year)].iterrows():
        """
        Un masque est créé pour identifier la ligne de référence qui correspond au même mois (PERIOD_MONTH) et aux mêmes identifiants que la ligne actuelle traitée dans la boucle. Le masque est également conditionné à ce que l'année soit l'année de référence (ref_year).
        """
        ref_mask = (df['PERIOD_YEAR'] == int(ref_year)) & (df['PERIOD_MONTH'] == row['PERIOD_MONTH'])
        for id_col in identifiers:
            """
            Ce masque est affiné pour s'assurer qu'il correspond exactement aux identifiants de la ligne actuellement traitée.
            """
            ref_mask &= (df[id_col] == row[id_col])
        """
        Vérification de la présence de la ligne de référence : Avant de procéder à tout calcul, on vérifie si le DataFrame filtré par le masque n'est pas vide. Cela évite des erreurs lors de l'accès à un DataFrame vide.
        """
        if not df.loc[ref_mask].empty:
            ref_row = df.loc[ref_mask].iloc[0]
            """
            Si la ligne de référence existe et que la valeur dans sum_eop_2024 est supérieure à zéro, un ratio est calculé comme étant la valeur de sum_eop_2025 divisée par celle de sum_eop_2024.
            """
            if ref_row[sum_eop_2024] > 0:
                ratio = ref_row[sum_eop_2025] / ref_row[sum_eop_2024]
                """
                Ce ratio est ensuite utilisé pour ajuster les valeurs des colonnes spécifiées dans viewing_columns. Pour chaque colonne de viewing_columns, la nouvelle valeur est calculée en multipliant la valeur originale de la colonne par le ratio calculé.
                """
                for col in viewing_columns:
                    df.at[index, col] = ref_row[col] * ratio
                """
                Si la valeur dans sum_eop_2024 est zéro ou si la ligne de référence n'existe pas, les valeurs dans les colonnes de viewing_columns sont mises à zéro pour la ligne courante traitée.
                """
            else:
                df.loc[index, viewing_columns] = 0
        else:
            df.loc[index, viewing_columns] = 0


def save_dataframe(df, output_path):
    """Saves the modified DataFrame to the specified Excel file."""
    output_filepath = f"{output_path}\\forecast_audience.xlsx"
    df.to_excel(output_filepath, index=False)
    print(f"Data saved to {output_filepath}")

def main():
    root = tk.Tk()
    tab = AudienceTab(root, {})
    setup_ui(tab)

    df = load_excel_data(tab.df)
    identifiers, viewing_columns = get_column_names(df)
    target_years = range(int(tab.target_start_year.get()), int(tab.target_end_year.get()) + 1)

    modified_df = duplicate_data(df, tab.references_year.get(), tab.references_month.get(), target_years, identifiers)
    adjust_viewing_columns(modified_df, tab.references_year.get(), identifiers, viewing_columns)
    save_dataframe(modified_df, tab.output_path.get())

    # root.mainloop()

if __name__ == "__main__":
    main()