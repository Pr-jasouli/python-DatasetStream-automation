from pathlib import Path
from ChannelSynthesizer.src.utils import get_provider_and_year, read_section_names, parse_tsv, find_file_pairs, \
    create_consolidated_excel


def generate_excel_report(output_directory):
    """
    genere un rapport Excel consolide a partir des fichiers de section et de texte.

    :param output_directory: repertoire contenant les fichiers de sortie
    :return: Le chemin vers le fichier Excel genere
    """
    section_dir = Path(output_directory) / 'section'
    text_dir = Path(output_directory) / 'text'
    output_path = Path(output_directory) / 'xlsx/consolidated_report.xlsx'

    all_data = []

    # lire et traiter les donnes des fournisseurs
    for section_file, text_file in find_file_pairs(section_dir, text_dir):
        provider, year = get_provider_and_year(text_file.stem)
        section_names = read_section_names(section_file)
        data = parse_tsv(text_file, section_names, provider)
        all_data.append((provider, year, data, section_names))

    # creer le rapport Excel consolide
    create_consolidated_excel(all_data, output_path)

    # renvoyer le chemin du fichier excel généré
    return output_path


if __name__ == "__main__":
    output_directory = Path('../../outputs')
    output_path = generate_excel_report(output_directory)  # capturer le chemin de sortie
    print(f"Generated Excel report at: {output_path}")  # afficher le chemin du fichier genere
