import os

from ChannelSynthesizer.src.parsers.all_sections_parser import extract_text, parse, get_provider_colors, detect_provider_and_year, get_pages_to_process, remove_redundant_sections, save_sections

"""
This script detects all "menus" or "sections" present in the channel lists (PDFs) of any provider.
Once determined, this list is sent to script 'B' which cleans the entire PDF as follows:
                -> row: section name
                -> row: channel number
                -> row: channel text (+additional tags for VOO)
When it's OK, script 'C' formats the result of 'B' into a centralized Excel sheet.
This separation allows locating when problems occur, with each intermediate result ('A' and 'B') being auditable.
The utils script contains the precise implementation of each step.
"""

def process(folder_path: str) -> None:
    """Processes all PDF files in the given folder."""
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
                save_sections(path, all_sections)
                print(f"Saved sections for {file} for provider {provider} and year {year}")

            except ValueError as e:
                print(f"Error processing {file}: {e}")

if __name__ == "__main__":
    folder = os.path.abspath(os.path.join(os.path.dirname(__file__), '../../inputs/pdf'))
    process(folder)
