import os
from parsers.Orange_text_parser import parse_orange_pdf
from parsers.VOO_text_parser import parse_voo_pdf
from parsers.all_sections_parser import detect_provider_and_year


def process_pdfs(directory):
    for filename in os.listdir(directory):
        if filename.endswith(".pdf"):
            pdf_path = os.path.join(directory, filename)
            try:
                provider, year = detect_provider_and_year(pdf_path)

                if provider == "VOO":
                    parse_voo_pdf(pdf_path)
                # elif provider == "Telenet":
                #     parse_telenet_pdf(pdf_path)
                elif provider == "Orange":
                    parse_orange_pdf(pdf_path)
                else:
                    print(f"Unsupported provider {provider} for file {filename}")

            except ValueError as e:
                print(f"Error processing {filename}: {e}")

if __name__ == "__main__":
    input_directory = os.path.abspath(os.path.join(os.path.dirname(__file__), '../../inputs/pdf'))
    process_pdfs(input_directory)