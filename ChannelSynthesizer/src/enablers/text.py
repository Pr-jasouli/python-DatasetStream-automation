import os
import json
import fitz
from ChannelSynthesizer.src.parsers.Orange_text_parser import parse_orange_pdf
from ChannelSynthesizer.src.parsers.VOO_text_parser import parse_voo_pdf
from ChannelSynthesizer.src.parsers.Telenet_text_parser import parse_telenet_pdf
from ChannelSynthesizer.src.parsers.all_sections_parser import detect_provider_and_year

CONFIG_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), '../../.config/page_selection.json'))

def load_page_selection() -> dict:
    """
    Charge la configuration de sélection de pages à partir d'un fichier JSON.
    Retourne:
    Un dictionnaire de la sélection de pages.
    """
    if os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}

def get_pages_to_process(pdf_path: str, total_pages: int) -> list:
    """
    Obtient les pages à traiter selon la configuration.
    Arguments:
    pdf_path -- le chemin du fichier PDF
    total_pages -- le nombre total de pages dans le PDF
    Retourne:
    Une liste de numéros de pages à traiter.
    """
    page_selection = load_page_selection()
    filename = os.path.basename(pdf_path)
    return page_selection.get(filename, list(range(1, total_pages + 1)))

def process_pdfs(directory):
    """
    Traite tous les fichiers PDF dans le répertoire donné.
    Arguments:
    directory -- le répertoire contenant les fichiers PDF
    """
    for filename in os.listdir(directory):
        if filename.endswith(".pdf"):
            pdf_path = os.path.join(directory, filename)
            try:
                provider, year = detect_provider_and_year(pdf_path)
                document = fitz.open(pdf_path)
                total_pages = document.page_count
                document.close()
                pages_to_process = get_pages_to_process(pdf_path, total_pages)

                if provider == "VOO":
                    parse_voo_pdf(pdf_path)
                elif provider == "Telenet":
                    parse_telenet_pdf(pdf_path, pages_to_process)
                elif provider == "Orange":
                    parse_orange_pdf(pdf_path)
                else:
                    print(f"Unsupported provider {provider} for file {filename}")

            except ValueError as e:
                print(f"Error processing {filename}: {e}")

if __name__ == "__main__":
    input_directory = os.path.abspath(os.path.join(os.path.dirname(__file__), '../../inputs/pdf'))
    process_pdfs(input_directory)
