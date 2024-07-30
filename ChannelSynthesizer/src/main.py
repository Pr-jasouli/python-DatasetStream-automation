import os
from enablers.sections import process as process_sections
from enablers.text import process_pdfs
from enablers.excel import generate_excel_report

def main():
    """
    Script principal pour traiter les fichiers PDF en trois étapes:
    1. Extraction des sections.
    2. Traitement du texte des PDF.
    3. Génération du rapport Excel consolidé.
    """
    input_directory = os.path.abspath(os.path.join(os.path.dirname(__file__), '../inputs/pdf'))
    output_directory = os.path.abspath(os.path.join(os.path.dirname(__file__), '../outputs'))

    print("Processing sections...")
    process_sections(input_directory)

    print("Processing text...")
    process_pdfs(input_directory)

    print("Generating consolidated Excel report...")
    generate_excel_report(output_directory)


if __name__ == "__main__":
    main()
