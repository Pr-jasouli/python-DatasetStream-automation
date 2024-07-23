import os

from utils.utils import extract_text, parse, write_section_tsv, get_provider_colors, detect_provider_and_year, get_pages_to_process, remove_redundant_sections

"""
Ce script va venir détecter tous les "menus" ou "sections" présents dans les listes de chaines (PDF) de n'importe quel provider.
Une fois déterminée, cette liste est envoyée au script 'B' qui va nettoyer l'entierete des pdf de cette façon:
                -> row: section name
                -> row: channel number
                -> row: channel text (+additional tags for VOO)
Quand c'est OK, script 'C' formate le resultat de 'B' en une feuille excel centralisée
Cette séparation va permettre de localiser à quel moment les problèmes surviennent, chaque résultat intermédiaire ('A' et 'B') étant auditable
le script utils contient l'implementation précis de chaque step.

ctrl + left_click sur un appel de fonction pour inspecter son implémentation
"""

PAGE_SELECTION_FILE = "page_selection.json"


def process(folder_path):
    for file in os.listdir(folder_path):
        if file.endswith(".pdf"):
            path = os.path.join(folder_path, file)
            try:
                provider, year = detect_provider_and_year(path)
                colors = get_provider_colors(provider)
                pages = get_pages_to_process(path)

                all_sections = []

                for page_number in pages:
                    text, max_size = extract_text(path, colors, provider, page_number)
                    sections = parse(text, provider, max_size)
                    all_sections.extend(sections)
                all_sections = remove_redundant_sections(all_sections)

                tsv_filename = os.path.join(folder_path, os.path.splitext(file)[0] + "a.tsv")
                write_section_tsv(tsv_filename, all_sections)
                print(f"Saved {tsv_filename} for provider {provider} and year {year}")

            except ValueError as e:
                print(f"Error processing {file}: {e}")


if __name__ == "__main__":
    folder = '.'
    process(folder)
