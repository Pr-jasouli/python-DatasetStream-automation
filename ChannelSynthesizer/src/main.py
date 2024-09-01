import os

import utilities
from ChannelSynthesizer.src.parsers.BASE_scraper import scrape_base_offer
from ChannelSynthesizer.src.utils import create_summary_table
from enablers.sections import process as process_sections
from enablers.text import process_pdfs
from enablers.excel import generate_excel_report
import pandas as pd


def main():
    """
    script principale pour traiter les fichiers pdf en quatre étapes:
    1. extraire les sections.
    2. traiter le texte des pdfs
    3. générer le rapport excel consolidé
    4. ajouter les données de l'offre BASE (y compris les chaînes radio)
    """

    # définir le répertoire d'entrée et de sortie
    input_directory = os.path.abspath(os.path.join(os.path.dirname(__file__), '../inputs/pdf'))
    output_directory = os.path.abspath(os.path.join(os.path.dirname(__file__), '../outputs'))

    print("Traitement des sections...")
    process_sections(input_directory)

    print("Traitement du texte...")
    process_pdfs(input_directory)

    print("Génération du rapport Excel consolidé...")
    output_path = generate_excel_report(output_directory)

    if output_path:
        print("Scraping des offres BASE et ajout au rapport...")
        base_url = "https://www.prd.base.be/en/support/tv/your-base-tv-box-and-remote/what-channels-does-base-offer/"
        base_offer_df = scrape_base_offer(base_url)

        # ajouter les données de l'offre BASE au fichier excel existant
        with pd.ExcelWriter(output_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            base_offer_df.to_excel(writer, sheet_name='Consolidated', index=False, header=False, startrow=writer.sheets['Consolidated'].max_row)

        # génération du rapport résumé après ajout des données BASE
        print("Génération du rapport de synthèse...")
        consolidated_df = pd.read_excel(output_path, sheet_name='Consolidated')
        summary_df = create_summary_table(consolidated_df)

        # ajouter le tableau résumé au fichier excel
        with pd.ExcelWriter(output_path, engine='openpyxl', mode='a') as writer:
            summary_df.to_excel(writer, sheet_name='Summary', index=False)
    else:
        print("Erreur : le rapport Excel consolidé n'a pas été généré.")

if __name__ == "__main__":
    main()
