import os
import re
import json
from utils.utils import extract_text, parse, write_section_tsv, get_provider_colors, detect_provider_and_year, get_pages_to_process


"""
Ce script détecte chaque section sur base de la couleur de police, écris dans un fichier {filename}a.tsv
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
                for page_number in pages:
                    text, max_size = extract_text(path, colors, provider, page_number)
                    sections = parse(text, provider, max_size)
                    section_tsv_file = os.path.join(folder_path, os.path.splitext(file)[0] + f"_page{page_number}a.tsv")
                    write_section_tsv(section_tsv_file, sections)
                    print(f"Saved {section_tsv_file} for provider {provider} and year {year}, page {page_number}")
            except ValueError as e:
                print(f"Error processing {file}: {e}")

if __name__ == "__main__":
    folder = '.'
    process(folder)
