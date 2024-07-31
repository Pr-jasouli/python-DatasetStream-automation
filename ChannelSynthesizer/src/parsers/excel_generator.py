from pathlib import Path
import pandas as pd
from ChannelSynthesizer.src.utils import get_provider_and_year, read_section_names, parse_tsv, adjust_region_columns

def filter_columns_with_values(df):
    # Colonnes non vides
    non_empty_columns = df.loc[:, (df != 0).any(axis=0)].columns
    return df[non_empty_columns]

"""
Fonction pour fusionner les colonnes Telenet
"""

def merge_telenet_columns(df):
    # Colonnes à fusionner en 'Offre de base'
    base_columns_to_merge = ['OFFRE DE BASE', 'BASISAANBOD', 'BASISAANBOD / OFFRE DE BASE']
    base_merged_column_name = 'Offre de base'

    # Colonnes à fusionner en 'Offre radio'
    radio_columns_to_merge = ['CHAÎNES DE RADIO', 'RADIOZENDERS / CHAÎNES DE RADIO']
    radio_merged_column_name = 'Offre radio'

    # Fusion des colonnes de base
    existing_base_columns = [col for col in base_columns_to_merge if col in df.columns]
    if existing_base_columns:
        df[base_merged_column_name] = df[existing_base_columns].sum(axis=1)
        df[base_merged_column_name] = df[base_merged_column_name].apply(lambda x: 1 if x > 0 else 0)
        df.drop(columns=existing_base_columns, inplace=True)

    # Fusion des colonnes radio
    existing_radio_columns = [col for col in radio_columns_to_merge if col in df.columns]
    if existing_radio_columns:
        df[radio_merged_column_name] = df[existing_radio_columns].sum(axis=1)
        df[radio_merged_column_name] = df[radio_merged_column_name].apply(lambda x: 1 if x > 0 else 0)
        df.drop(columns=existing_radio_columns, inplace=True)

    # Fusion 'MUZIEKZENDERS/CHAÎNES DE MUSIQUE' et 'MUSIC' en 'Chaînes Musique'
    music_columns_to_merge = ['MUZIEKZENDERS/CHAÎNES DE MUSIQUE', 'MUSIC']
    music_merged_column_name = 'Chaînes Musique'

    existing_music_columns = [col for col in music_columns_to_merge if col in df.columns]
    if existing_music_columns:
        df[music_merged_column_name] = df[existing_music_columns].sum(axis=1)
        df[music_merged_column_name] = df[music_merged_column_name].apply(lambda x: 1 if x > 0 else 0)
        df.drop(columns=existing_music_columns, inplace=True)

    # Renommer 'DOCU' en 'Chaînes Documentaires'
    if 'DOCU' in df.columns:
        df.rename(columns={'DOCU': 'Chaînes Documentaires'}, inplace=True)

    # Renommer 'PASSION XL' en '+18'
    if 'PASSION XL' in df.columns:
        df.rename(columns={'PASSION XL': '+18'}, inplace=True)

    # Régler les colonnes pour 'OPTION FR'
    if 'OPTION FR' in df.columns:
        df.loc[df['OPTION FR'] == 1, ['Region Wallonia', 'Brussels', 'Communauté Germanophone', 'Region Flanders']] = [1, 0, 0, 0]

    # Fusionner 'OPTION FR' en 'Region Wallonia'
    if 'OPTION FR' in df.columns and 'Region Wallonia' in df.columns:
        df['Region Wallonia'] += df['OPTION FR']
        df['Region Wallonia'] = df['Region Wallonia'].apply(lambda x: 1 if x > 0 else 0)
        df.drop(columns=['OPTION FR'], inplace=True)

    return df

"""
Fonction pour fusionner les colonnes VOO
"""

def merge_voo_columns(df):
    # Régler les colonnes pour 'Bruxelles'
    if 'Bruxelles' in df.columns:
        df.loc[df['Bruxelles'] == 1, ['Brussels', 'Region Wallonia', 'Communauté Germanophone', 'Region Flanders']] = [1, 0, 0, 0]

    # Régler les colonnes pour 'Comm. German'
    if 'Comm. German' in df.columns:
        df.loc[df['Comm. German'] == 1, ['Communauté Germanophone', 'Region Wallonia', 'Brussels', 'Region Flanders']] = [1, 0, 0, 0]

    # Régler les colonnes pour 'Wallonie'
    if 'Wallonie' in df.columns:
        df.loc[df['Wallonie'] == 1, ['Region Wallonia', 'Communauté Germanophone', 'Brussels', 'Region Flanders']] = [1, 0, 0, 0]

    # Régler les colonnes pour 'Chaînes Néerlandophones'
    if 'Chaînes Néerlandophones' in df.columns:
        df.loc[df['Chaînes Néerlandophones'] == 1, ['Region Flanders', 'Communauté Germanophone', 'Region Wallonia', 'Brussels']] = [1, 0, 0, 0]

    # Fusionner 'Bruxelles' en 'Brussels'
    if 'Bruxelles' in df.columns and 'Brussels' in df.columns:
        df['Brussels'] += df['Bruxelles']
        df['Brussels'] = df['Brussels'].apply(lambda x: 1 if x > 0 else 0)
        df.drop(columns=['Bruxelles'], inplace=True)

    # Fusionner 'Comm. German' en 'Communauté Germanophone'
    if 'Comm. German' in df.columns and 'Communauté Germanophone' in df.columns:
        df['Communauté Germanophone'] += df['Comm. German']
        df['Communauté Germanophone'] = df['Communauté Germanophone'].apply(lambda x: 1 if x > 0 else 0)
        df.drop(columns=['Comm. German'], inplace=True)

    # Fusionner 'Wallonie' en 'Region Wallonia'
    if 'Wallonie' in df.columns and 'Region Wallonia' in df.columns:
        df['Region Wallonia'] += df['Wallonie']
        df['Region Wallonia'] = df['Region Wallonia'].apply(lambda x: 1 if x > 0 else 0)
        df.drop(columns=['Wallonie'], inplace=True)

    # Fusionner 'Chaînes Néerlandophones' en 'Region Flanders'
    if 'Chaînes Néerlandophones' in df.columns and 'Region Flanders' in df.columns:
        df['Region Flanders'] += df['Chaînes Néerlandophones']
        df['Region Flanders'] = df['Region Flanders'].apply(lambda x: 1 if x > 0 else 0)
        df.drop(columns=['Chaînes Néerlandophones'], inplace=True)

    # Renommer 'Chaînes Radios' en 'Offre radio'
    if 'Chaînes Radios' in df.columns:
        df.rename(columns={'Chaînes Radios': 'Offre radio'}, inplace=True)

    # Renommer 'Classé X' en '+18'
    if 'Classé X' in df.columns:
        df.rename(columns={'Classé X': '+18'}, inplace=True)

    return df

"""
Fonction pour fusionner les colonnes Orange
"""

def merge_orange_columns(df):
    # Régler les colonnes pour 'Duitstalig'
    if 'Duitstalig' in df.columns:
        df.loc[df['Duitstalig'] == 1, ['Communauté Germanophone', 'Brussels', 'Region Flanders', 'Region Wallonia']] = [1, 0, 0, 0]

    # Régler les colonnes pour 'Franstalig'
    if 'Franstalig' in df.columns:
        df.loc[df['Franstalig'] == 1, ['Region Wallonia', 'Communauté Germanophone', 'Region Flanders', 'Brussels']] = [1, 0, 0, 0]

    # Régler les colonnes pour 'Nederlandstalig'
    if 'Nederlandstalig' in df.columns:
        df.loc[df['Nederlandstalig'] == 1, ['Region Flanders', 'Communauté Germanophone', 'Region Wallonia', 'Brussels']] = [1, 0, 0, 0]

    # Fusionner 'Duitstalig' en 'Communauté Germanophone'
    if 'Duitstalig' in df.columns and 'Communauté Germanophone' in df.columns:
        df['Communauté Germanophone'] += df['Duitstalig']
        df['Communauté Germanophone'] = df['Communauté Germanophone'].apply(lambda x: 1 if x > 0 else 0)
        df.drop(columns=['Duitstalig'], inplace=True)

    # Fusionner 'Franstalig' en 'Region Wallonia'
    if 'Franstalig' in df.columns and 'Region Wallonia' in df.columns:
        df['Region Wallonia'] += df['Franstalig']
        df['Region Wallonia'] = df['Region Wallonia'].apply(lambda x: 1 if x > 0 else 0)
        df.drop(columns=['Franstalig'], inplace=True)

    # Fusionner 'Nederlandstalig' en 'Region Flanders'
    if 'Nederlandstalig' in df.columns and 'Region Flanders' in df.columns:
        df['Region Flanders'] += df['Nederlandstalig']
        df['Region Flanders'] = df['Region Flanders'].apply(lambda x: 1 if x > 0 else 0)
        df.drop(columns=['Nederlandstalig'], inplace=True)

    # Fusionner 'OPTION FR' en 'Region Wallonia'
    if 'OPTION FR' in df.columns and 'Region Wallonia' in df.columns:
        df['Region Wallonia'] += df['OPTION FR']
        df['Region Wallonia'] = df['Region Wallonia'].apply(lambda x: 1 if x > 0 else 0)
        df.drop(columns=['OPTION FR'], inplace=True)

    # Renommer 'Regionale zenders' en 'Chaînes locales'
    if 'Regionale zenders' in df.columns:
        df.rename(columns={'Regionale zenders': 'Chaînes locales'}, inplace=True)

    # Renommer 'Radio' en 'Offre radio'
    if 'Radio' in df.columns:
        df.rename(columns={'Radio': 'Offre radio'}, inplace=True)

    # Renommer 'Muziek' en 'Chaînes Musique'
    if 'Muziek' in df.columns:
        df.rename(columns={'Muziek': 'Chaînes Musique'}, inplace=True)

    return df

"""
Fonction pour réorganiser les colonnes
"""

def reorder_columns(df, static_columns, merged_renamed_columns):
    # Colonnes filtrées existantes
    ordered_columns = ['Channel', 'Provider_Period'] + static_columns + merged_renamed_columns
    ordered_columns = [col for col in ordered_columns if col in df.columns] + [col for col in df.columns if col not in ordered_columns]
    return df[ordered_columns]

"""
Fonction pour créer un Excel consolidé
"""

def create_consolidated_excel(all_data, output_path):
    print("Creating consolidated Excel...")

    combined_data = []
    combined_columns = ['Provider_Period', 'Channel']

    combined_data_by_provider_year = {}
    for provider, year, data, section_names in all_data:
        key = (provider, year)
        if key not in combined_data_by_provider_year:
            combined_data_by_provider_year[key] = {'data': [], 'section_names': set()}
        combined_data_by_provider_year[key]['data'].extend(data)
        combined_data_by_provider_year[key]['section_names'].update(section_names)

    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        for (provider, year), value in combined_data_by_provider_year.items():
            period = f"{provider} {year}"
            data = value['data']
            section_names = list(value['section_names'])

            text_rows = list(set(entry[1] for entry in data))
            static_columns = ['Region Flanders', 'Brussels', 'Region Wallonia', 'Communauté Germanophone']
            merged_renamed_columns = ['Offre de base', 'Offre radio', 'Chaînes locales', 'Chaînes Documentaires', 'Chaînes Musique']
            columns = static_columns + section_names

            additional_columns = set()
            for entry in data:
                additional_columns.update(entry[2:])  # Colonnes supplémentaires VOO

            columns.extend(additional_columns)
            df = pd.DataFrame(0, index=text_rows, columns=columns)

            for entry in data:
                section = entry[0]
                text = entry[1]
                if section in df.columns:
                    df.at[text, section] = 1
                for additional_column in entry[2:]:
                    if additional_column in df.columns:
                        df.at[text, additional_column] = 1

            df = adjust_region_columns(df)

            # Fusion des colonnes pour Telenet
            if provider == 'Telenet':
                df = merge_telenet_columns(df)

            # Fusion des colonnes pour VOO
            if provider == 'Voo':
                df = merge_voo_columns(df)

            # Fusion des colonnes pour Orange
            if provider == 'Orange':
                df = merge_orange_columns(df)

            df.insert(0, 'Provider_Period', period)
            df = df.reset_index().rename(columns={'index': 'Channel'})

            # Réorganiser les colonnes
            df = reorder_columns(df, static_columns, merged_renamed_columns)

            filtered_df = filter_columns_with_values(df)

            sheet_name = f"{provider}_{year}"
            filtered_df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=0)

            workbook = writer.book
            worksheet = writer.sheets[sheet_name]

            header_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter'})

            for col_num, value in enumerate(filtered_df.columns.values):
                worksheet.write(0, col_num, value, header_format)

            for i, col in enumerate(filtered_df.columns, 0):
                max_length = max((filtered_df[col].astype(str).map(len).max(), len(col))) + 2
                worksheet.set_column(i, i, max_length)

            worksheet.autofilter(0, 0, 0, len(filtered_df.columns) - 1)

        combined_data = []
        for (provider, year), value in combined_data_by_provider_year.items():
            period = f"{provider} {year}"
            data = value['data']
            section_names = list(value['section_names'])

            text_rows = list(set(entry[1] for entry in data))
            static_columns = ['Region Flanders', 'Brussels', 'Region Wallonia', 'Communauté Germanophone']
            merged_renamed_columns = ['Offre de base', 'Offre radio', 'Chaînes locales', 'Chaînes Documentaires', 'Chaînes Musique']
            columns = static_columns + section_names

            additional_columns = set()
            for entry in data:
                additional_columns.update(entry[2:])  # Colonnes supplémentaires VOO

            columns.extend(additional_columns)
            df = pd.DataFrame(0, index=text_rows, columns=columns)

            for entry in data:
                section = entry[0]
                text = entry[1]
                if section in df.columns:
                    df.at[text, section] = 1
                for additional_column in entry[2:]:
                    if additional_column in df.columns:
                        df.at[text, additional_column] = 1

            df = adjust_region_columns(df)

            # Fusion des colonnes pour Telenet
            if provider == 'Telenet':
                df = merge_telenet_columns(df)

            # Fusion des colonnes pour VOO
            if provider == 'Voo':
                df = merge_voo_columns(df)

            # Fusion des colonnes pour Orange
            if provider == 'Orange':
                df = merge_orange_columns(df)

            df.insert(0, 'Provider_Period', period)
            df = df.reset_index().rename(columns={'index': 'Channel'})

            # Réorganiser les colonnes
            df = reorder_columns(df, static_columns, merged_renamed_columns)

            combined_data.append(df)

        if combined_data:
            final_df = pd.concat(combined_data)
        else:
            final_df = pd.DataFrame(columns=combined_columns + columns)

        final_df = final_df.fillna(0)

        final_df.to_excel(writer, sheet_name='Consolidated', index=False, startrow=0)

        workbook = writer.book
        worksheet = writer.sheets['Consolidated']


        header_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter'})
        cell_format = workbook.add_format({'align': 'center'})
        left_align_format = workbook.add_format({'align': 'left'})
        right_align_format = workbook.add_format({'align': 'right'})

        for col_num, value in enumerate(final_df.columns.values):
            if value == 'Channel':
                worksheet.write(0, col_num, value, header_format)
                worksheet.set_column(col_num, col_num, None, right_align_format)
            elif value == 'Provider_Period':
                worksheet.write(0, col_num, value, header_format)
                worksheet.set_column(col_num, col_num, None, left_align_format)
            else:
                worksheet.write(0, col_num, value, header_format)
                worksheet.set_column(col_num, col_num, None, cell_format)

        for i, col in enumerate(final_df.columns, 0):
            max_length = max((final_df[col].astype(str).map(len).max(), len(col))) + 2
            worksheet.set_column(i, i, max_length, cell_format if col not in ['Channel', 'Provider_Period'] else None)

        worksheet.autofilter(0, 0, 0, len(final_df.columns) - 1)
        worksheet.freeze_panes(1, 0)

    print(f"Consolidated Excel created at: {output_path}")

"""
Fonction pour trouver des paires de fichiers
"""

def find_file_pairs(section_dir, text_dir):
    print(f"Finding file pairs in {section_dir} and {text_dir}")
    section_files = list(Path(section_dir).glob('*_sections.tsv'))
    text_files = list(Path(text_dir).glob('*.tsv'))

    print(f"Section files found: {[f.name for f in section_files]}")
    print(f"Text files found: {[f.name for f in text_files]}")

    file_pairs = []
    for section_file in section_files:
        base_name = section_file.stem.replace('_sections', '')
        text_file = next(
            (text_file for text_file in text_files if base_name in text_file.stem and '_text' in text_file.stem), None)
        if text_file and text_file.exists():
            file_pairs.append((section_file, text_file))

    print(f"Found {len(file_pairs)} file pairs.")
    return file_pairs

"""
Fonction pour traiter les fournisseurs
"""

def process_providers(section_dir, text_dir, output_path):
    file_pairs = find_file_pairs(section_dir, text_dir)
    all_data = []

    for section_file, text_file in file_pairs:
        provider, year = get_provider_and_year(text_file.stem)
        section_names = read_section_names(section_file)
        data = parse_tsv(text_file, section_names, provider)
        all_data.append((provider, year, data, section_names))

    create_consolidated_excel(all_data, output_path)
