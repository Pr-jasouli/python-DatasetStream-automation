import os
from enablers.sections import process as process_sections
from enablers.text import process_pdfs

def main():
    """
    Script principal pour traiter les fichiers PDF en deux étapes:
    1. Extraction des sections.
    2. Traitement du texte des PDF.
    """
    input_directory = os.path.abspath(os.path.join(os.path.dirname(__file__), '../inputs/pdf'))

    print("Processing sections...")
    process_sections(input_directory)

    print("Processing text...")
    process_pdfs(input_directory)

if __name__ == "__main__":
    main()
