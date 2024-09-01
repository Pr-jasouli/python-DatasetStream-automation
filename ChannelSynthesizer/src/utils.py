import re

import pandas as pd

def get_provider_and_year(filename):
    """
    cette fonction determine le nom du fournisseur et l'année du fichier
    le nom du fournisseur est chercher dans la liste et l'année est trouvé par une expression regulière
    """
    provider_names = ['voo', 'orange', 'telenet']
    provider = None
    year = None

    for name in provider_names:
        if name.lower() in filename.lower():
            provider = name.capitalize()
            break

    year_match = re.search(r'\d{4}', filename)
    if year_match:
        year = year_match.group(0)

    return provider, year

def read_section_names(file_path):
    """
    lis les noms des sections à partir d'un fichier, retournes une liste des noms
    chaque ligne du fichier est traité comme un nom de section
    """
    with open(file_path, 'r', encoding='utf-8') as f:
        section_names = [line.strip() for line in f.readlines()]
    return section_names


def parse_tsv(tsv_path, section_names, provider):
    """
    analyse un fichier TSV, ignors les codes d'info VOO et retourne les données importantes
    :param tsv_path: chemin vers le fichier TSV
    :param section_names: liste des noms de sections
    :param provider: le nom du fournisseur (e.g., "Voo", "Telenet", "Orange")
    :return: liste des données analysées
    """
    with open(tsv_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    data = []
    current_section = None

    for line in lines:
        stripped_line = line.strip()
        if stripped_line in section_names:
            current_section = stripped_line
        elif not stripped_line.isdigit() and not re.match(r'^\d{1,3}$', stripped_line):
            if current_section:
                data.append([current_section, stripped_line])

    return data


def ensure_region_columns_exist(df):
    """verifier que toutes les colonnes des régions existent dans le DataFrame."""
    region_columns = ['Region Flanders', 'Brussels', 'Region Wallonia', 'Communauté Germanophone']
    for col in region_columns:
        if col not in df.columns:
            df[col] = 0  # Initialiser la colonne avec des zéros
    return df

def process_providers(section_dir, text_dir, output_path):
    """
    traite les fichiers des differents fournisseurs et creer un rapport consolidé
    :param section_dir: répertoire contenant les fichiers de section
    :param text_dir: répertoire contenant les fichiers texte
    :param output_path: chemin vers le fichier de sortie
    """
    file_pairs = find_file_pairs(section_dir, text_dir)
    all_data = []

    for section_file, text_file in file_pairs:
        provider, year = get_provider_and_year(text_file.stem)
        section_names = read_section_names(section_file)
        data = parse_tsv(text_file, section_names, provider)
        all_data.append((provider, year, data, section_names))

    create_consolidated_excel(all_data, output_path)

from pathlib import Path

def find_file_pairs(section_dir, text_dir):
    """
    Trouver les paires de fichiers de section et texte basé sur leurs noms de fichier
    :param section_dir: répertoire contenant les fichiers de section
    :param text_dir: répertoire contenant les fichiers texte
    :return: liste de tuples, chaque tuple contient une paire de (section_file, text_file)
    """
    section_files = list(Path(section_dir).glob('*_sections.tsv'))
    text_files = list(Path(text_dir).glob('*.tsv'))

    file_pairs = []
    for section_file in section_files:
        base_name = section_file.stem.replace('_sections', '')
        text_file = next(
            (text_file for text_file in text_files if base_name in text_file.stem and '_text' in text_file.stem), None)
        if text_file:
            file_pairs.append((section_file, text_file))

    return file_pairs


def post_process_orange_regions(final_df):
    """
    verifier que tous les lignes pour Orange ont des codes de région cohérents
    :param final_df: Le DataFrame contenant toutes les données consolidées
    :return: Le DataFrame avec les codes de région Orange post-traités
    """
    # Filtrer seulement les lignes Orange
    orange_df = final_df[final_df['Provider_Period'].str.contains("Orange")]

    # Boucler à travers les lignes Orange pour appliquer la cohérence des régions
    for index, row in orange_df.iterrows():
        # Trouver quelle colonne de région a la valeur '1'
        regions = row[['Region Flanders', 'Brussels', 'Region Wallonia', 'Communauté Germanophone']]
        if regions.sum() > 1:  # Si plus d'une région est marqué comme disponible
            # Déterminer la bonne région en vérifiant les autres lignes avec le même nom de chaîne
            matching_rows = orange_df[orange_df['Channel'] == row['Channel']]
            correct_region = matching_rows[
                ['Region Flanders', 'Brussels', 'Region Wallonia', 'Communauté Germanophone']].sum(axis=0).idxmax()

            # Mettre seulement la région correcte à 1, et les autres à 0
            final_df.loc[index, ['Region Flanders', 'Brussels', 'Region Wallonia', 'Communauté Germanophone']] = 0
            final_df.loc[index, correct_region] = 1

    return final_df


def is_basic_section(section_name):
    """
    determine si un nom de section correspond à une offre de base
    :param section_name: Le nom de la section à vérifier
    :return: True si la section est considérée comme basique, sinon False
    """
    basic_sections = [
        r'BASISAANBOD',
        r'RADIOZENDERS',
        r'STINGRAY MUSIC',
        r'OFFRE DE BASE',
        r'CHAÎNES DE RADIO',
        r'CHAÎNES DE MUSIQUE',
        r'BASISAANBOD / OFFRE DE BASE',
        r'RADIOZENDERS / CHAÎNES DE RADIO',
        r'MUZIEKZENDERS/CHAÎNES DE MUSIQUE'
    ]

    # Compiler regex pour être insensible à la casse
    basic_sections_regex = re.compile('|'.join(basic_sections), re.IGNORECASE)

    return bool(basic_sections_regex.search(section_name))

def synchronize_channel_group_case(final_df):
    """
    synchroniser la casse pour les valeurs de 'Channel Group Level' qui diffèrent uniquement par la casse
    :param final_df: Le DataFrame contenant les données consolidées
    :return: DataFrame avec la casse 'Channel Group Level' synchronisé
    """
    # Grouper par minuscule 'Channel Group Level' pour trouver des doublons en ignorant la casse
    group_map = final_df.groupby(final_df['Channel Group Level'].str.lower())['Channel Group Level'].first().to_dict()

    # Appliquer la casse synchronisé
    final_df['Channel Group Level'] = final_df['Channel Group Level'].str.lower().map(group_map)

    return final_df
def create_summary_table(final_df):
    """
    cree un tableau récapitulatif basé sur les niveaux de groupe de chaînes uniques
    exclure les lignes où TV/Radio est 'Radio'
    :param final_df: Le DataFrame contenant les données consolidées
    :return: un DataFrame avec des statistiques récapitulatives
    """
    # Exclure les lignes où TV/Radio est 'Radio'
    filtered_df = final_df[final_df['TV/Radio'] != 'Radio']

    # Grouper par Provider_Period et Channel Group Level, puis supprimer les doublons
    unique_channels = filtered_df.drop_duplicates(subset=['Provider_Period', 'Channel Group Level'])

    # Tableau croisé pour compter les chaînes Basic et Option
    summary_df = unique_channels.pivot_table(index='Provider_Period',
                                             columns='Basic/Option',
                                             aggfunc='size',
                                             fill_value=0).reset_index()

    # Ajouter le total général
    summary_df['Grand Total'] = summary_df['Basic'] + summary_df['Option']

    # Renommer les colonnes
    summary_df.columns.name = None  # Retirer le nom de la colonne croisée
    summary_df = summary_df.rename(columns={'Provider_Period': 'Row Labels'})
    overall_totals = {
        'Row Labels': 'Grand Total',
        'Basic': summary_df['Basic'].sum(),
        'Option': summary_df['Option'].sum(),
        'Grand Total': summary_df['Grand Total'].sum()
    }

    # Ajouter la ligne des totaux généraux au DataFrame récapitulatif
    summary_df = summary_df._append(overall_totals, ignore_index=True)
    return summary_df


def create_consolidated_excel(all_data, output_path):
    """
    crée un rapport Excel consolidé à partir des données analysées
    :param all_data: liste de tuples contenant le fournisseur, l'année, les données et les noms de section
    :param output_path: chemin vers le fichier Excel à enregistrer
    """
    combined_data = []

    VOO_info_codes = {
        "VS": "VOOsport",
        "w VS": "VOOsport World",
        "Pa": "Bouquet Panorama",
        "Ci": "Option Ciné Pass",
        "Doc": "Be Bouquet Documentaires",
        "Div": "Be Bouquet Divertissement",
        "Co": "Be Cool",
        "Enf": "Be Bouquet Enfant",
        "Sp": "Be Bouquet Sport",
        "Sel": "Be Bouquet Selection",
        "Inf": "Option Infos",
        "Sen": "Option Sensation",
        "Ch": "Option Charme",
        "FF": "Family Fun",
        "DM": "Discover More",
        "CX": "Classé X",
        "MX": "Man-X",
    }

    orange_option_keywords = re.compile(
        r'Be 1|Be Séries|Be seri|Be Cin|Ciné\+|Cine\+|ELEVEN Pro|VOOSPORT WORLD',
        re.IGNORECASE
    )

    for provider, year, data, section_names in all_data:
        period = f"{provider} {year}"
        static_columns = ['Region Flanders', 'Brussels', 'Region Wallonia', 'Communauté Germanophone']
        df_data = []

        for entry in data:
            section = entry[0]
            channel = entry[1]



            # Initialiser les régions comme non disponibles
            regions = [0, 0, 0, 0]  # [Flanders, Brussels, Wallonia, Germanophone]

            # Déterminer les régions en fonction des codes de région dans le nom de la chaîne
            if provider in ["Orange", "Voo"]:
                if 'W' in channel.split():
                    regions = [0, 0, 1, 0]  # Seulement Wallonie
                elif 'B' in channel.split():
                    regions = [0, 1, 0, 0]  # Seulement Bruxelles
                elif 'G' in channel.split():
                    regions = [0, 0, 0, 1]  # Seulement Germanophone
                elif 'F' in channel.split():
                    regions = [1, 0, 0, 0]  # Seulement Flandre
                else:
                    regions = [1, 1, 1, 1]  # Par défaut toutes les régions

                # Supprimer le code de région du nom de la chaîne
                channel = re.sub(r'\b(W|B|G|F| w)\b', '', channel).strip()

            elif provider == "Telenet":
                regions = [1, 1, 0, 0]  # Disponible en Flandre et Bruxelles

            # Déterminer si la section est Basic ou Option pour VOO
            if provider == "Voo":
                if any(code in channel.split() for code in VOO_info_codes.keys()):
                    option = 'Option'
                else:
                    option = 'Basic'
                # Supprimer les codes d'info VOO du nom de la chaîne
                channel = ' '.join([word for word in channel.split() if word not in VOO_info_codes.keys()])
            elif provider == "Orange":
                # Pour Orange, par défaut Basic sauf si le nom de la chaîne correspond à l'un des mots-clés d'option
                if orange_option_keywords.search(channel):
                    option = 'Option'
                else:
                    option = 'Basic'
            else:
                # Pour d'autres fournisseurs, utiliser le nom de la section pour déterminer Basic ou Option
                if is_basic_section(section):
                    option = 'Basic'
                else:
                    option = 'Option'

            # Déterminer si la chaîne est TV ou Radio en fonction du nom de la chaîne
            if channel.endswith('TV'):
                tv_radio = 'TV'
                channel = channel[:-2].strip()  # Supprimer le suffixe 'TV' du nom de la chaîne
            elif channel.endswith('R'):
                tv_radio = 'Radio'
                channel = channel[:-1].strip()  # Supprimer le suffixe 'R' du nom de la chaîne
            else:
                tv_radio = 'TV'

            # Déterminer si la chaîne est HD, SD, ou ni l'un ni l'autre
            if 'HD' in channel:
                hd_sd = 'HD'
            elif 'SD' in channel:
                hd_sd = 'SD'
            else:
                hd_sd = ''

            # Ajouter les données traitées pour cette chaîne
            df_data.append([channel, period] + regions + [option, tv_radio, hd_sd])

        # Créer un DataFrame à partir des données traitées
        df = pd.DataFrame(df_data, columns=['Channel', 'Provider_Period'] + static_columns + ['Basic/Option', 'TV/Radio', 'HD/SD'])
        combined_data.append(df)

    # Combiner tous les DataFrames en un seul DataFrame final
    final_df = pd.concat(combined_data, ignore_index=True)

    # Appliquer un post-traitement pour la cohérence des régions Orange
    final_df = post_process_orange_regions(final_df)

    # Supprimer les lignes où la valeur de Channel est simplement 'w'
    final_df = final_df[final_df['Channel'] != 'w']

    # Supprimer les lignes en double avec le même 'Channel' et 'Provider_Period'
    final_df = final_df.drop_duplicates(subset=['Channel', 'Provider_Period'])

    # Ajouter une colonne 'Channel Group Level'
    final_df['Channel Group Level'] = final_df['Channel'].apply(lambda x: re.sub(
        r"\b(HD|SD|FR|NL|van DAZN|de DAZN|Gent|Sint-Niklaas|Dendermonde|Oudenaarde|Eeklo|Herenthout|Turnhout|Vl\.|Disco & Funk|Komen|Limburg|Mechelen|Geel|Antwerpen|West-Vl\.|60\'s-70\'s|80\'s & 90\'s|VL|Internationale|Vlaams Brabant|Oost Vlaanderen|West Vlaanderen|Eng|NWS|CANVAS|Canvas|Nl|Ukraine|Crime|Dagpas|Info|Open|Premier League|-Music Radio|-music radio|-music|-Maximum Hits|-Foute Radio|-Allstar| France|Germany|Italy|Nordic|Spain|Turk|Müzigi|non stop dokters|non-stop dokters|Maroc|Monde|Europe|.|English|Int.|Belgique| Classic|Frisson|Premier|R'n'B & Soul|Rock|21|Calm|Greats|Orchestral|International|Vl.|JR|Jr|Wild|WILD|- TVi|-tvi| tvi|Television|Plug|Club|club|- UK|Breizh|Oost| Plus| Polonia|Classic| •)\b",
        '',
        x).strip()
    )

    final_df['Channel Group Level'] = final_df['Channel Group Level'].str.replace(r'[()]', '', regex=True)


    final_df = synchronize_channel_group_case(final_df)

    # Ajouter une colonne HD/SD
    final_df['HD/SD'] = final_df['Channel'].apply(lambda x: 'HD' if 'HD' in x else ('SD' if 'SD' in x else ''))

    final_df = final_df[final_df['Channel'].str.strip() != '']



    # Écrire le DataFrame final dans un fichier Excel
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        final_df.to_excel(writer, sheet_name='Consolidated', index=False)

    print(f"Rapport Excel consolidé créé à : {output_path}")
