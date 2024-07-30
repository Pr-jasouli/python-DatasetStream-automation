# excel.py

from pathlib import Path
import pandas as pd
from ChannelSynthesizer.src.parsers.excel_generator import process_providers

def generate_excel_report(output_directory):
    """
    Génére le rapport Excel consolidé à partir des fichiers de sections et de texte.

    :param output_directory: Répertoire contenant les fichiers de sortie
    """
    section_dir = Path(output_directory) / 'section'
    text_dir = Path(output_directory) / 'text'
    output_path = Path(output_directory) / 'xlsx/consolidated_report.xlsx'

    process_providers(section_dir, text_dir, output_path)

if __name__ == "__main__":
    output_directory = Path('../../outputs')
    generate_excel_report(output_directory)
