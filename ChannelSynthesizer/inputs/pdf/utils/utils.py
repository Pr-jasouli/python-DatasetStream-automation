import re
import fitz
import os
import json

PAGE_SELECTION_FILE = "page_selection.json"
TELENET_WHITE_COLOR = 16777215
TELENET_BLACK_COLOR = 1113103  # (hex 11110f)


def extract_text(pdf_path, colors, provider, page_number):
    document = fitz.open(pdf_path)
    extracted_text = []
    sizes = set()

    page = document.load_page(page_number - 1)
    blocks = page.get_text("dict")["blocks"]

    for block in blocks:
        if 'lines' in block:
            for line in block["lines"]:
                for span in line["spans"]:
                    if provider == "Telenet":
                        sizes.add(span["size"])
                        is_bold = "bold" in span["font"].lower()
                        extracted_text.append((span["text"], span["color"], span["size"], is_bold, line["bbox"]))
                    elif provider == "Orange" and span["color"] == TELENET_WHITE_COLOR and (
                            span["text"][0].isupper() or span["text"].startswith('+')):
                        extracted_text.append(span["text"])
                    elif span["color"] in colors:
                        sizes.add(span["size"])
                        extracted_text.append(span["text"])

    # Determine the maximum font size to exclude for VOO
    max_size = max(sizes) if provider == "VOO" else None

    filtered_text = []
    for block in blocks:
        if 'lines' in block:
            for line in block["lines"]:
                for span in line["spans"]:
                    if provider == "Telenet":
                        is_bold = "bold" in span["font"].lower()
                        filtered_text.append((span["text"], span["color"], span["size"], is_bold, line["bbox"]))
                    elif provider == "Orange" and span["color"] == TELENET_WHITE_COLOR and (
                            span["text"][0].isupper() or span["text"].startswith('+')):
                        filtered_text.append((span["text"], span["color"]))
                    elif span["color"] in colors and (provider != "VOO" or span["size"] != max_size):
                        filtered_text.append((span["text"], span["color"]))

    document.close()
    return filtered_text, max_size


def parse(text, provider, max_size=None):
    lines = text
    sections = []
    current_section = None
    prev_line_info = None

    for line_info in lines:
        if provider == "Telenet":
            line, color, size, is_bold, bbox = line_info
            parsable = False
            if color == TELENET_WHITE_COLOR or (
                    color == TELENET_BLACK_COLOR and line.isupper() and len(line) >= 4 and not any(
                    char.isdigit() for char in line)) or is_bold:
                parsable = True

            if prev_line_info:
                prev_line, prev_color, prev_size, prev_is_bold, prev_bbox = prev_line_info
                prev_parsable = prev_color == TELENET_WHITE_COLOR or (
                            prev_color == TELENET_BLACK_COLOR and prev_line.isupper() and len(
                        prev_line) >= 4 and not any(char.isdigit() for char in prev_line)) or prev_is_bold
                if parsable and prev_parsable and abs(bbox[1] - prev_bbox[3]) < 10:
                    sections[-1] = [line.strip()]
                    prev_line_info = line_info
                    continue

            if parsable:
                sections.append([line.strip()])
                prev_line_info = line_info
            else:
                prev_line_info = None
            continue

        else:
            line, color = line_info

        words = line.split()
        if len(words) == 0:
            continue

        if provider == "Orange" and (len(words) > 0 and (line[0].isupper() or line.startswith('+'))):
            if current_section:
                sections.append(current_section)
            current_section = [line.strip()]
        elif 1 < len(words) <= 3 and not any(char.isdigit() for char in line):
            if current_section:
                sections.append(current_section)
            current_section = [line.strip()]
        else:
            if current_section is not None:
                current_section.append(line.strip())

    if current_section and current_section not in sections:
        sections.append(current_section)

    return remove_redundant_sections(sections)
def remove_redundant_sections(sections):
    seen_sections = set()
    unique_sections = []
    for section in sections:
        section_name = section[0]
        if section_name not in seen_sections:
            unique_sections.append(section)
            seen_sections.add(section_name)
    return unique_sections


def write_section_tsv(file, sections):
    with open(file, 'w', encoding='utf-8') as f:
        for section in sections:
            f.write(section[0] + '\n')


def get_provider_colors(provider):
    if provider == "VOO":
        return [16777215, 14092940]
    elif provider == "Telenet":
        return [TELENET_WHITE_COLOR, TELENET_BLACK_COLOR]
    elif provider == "Orange":
        return [16777215]
    else:
        raise ValueError("Unknown provider")


def detect_provider_and_year(pdf_path):
    filename = os.path.basename(pdf_path).lower()
    if "voo" in filename:
        provider = "VOO"
    elif "telenet" in filename:
        provider = "Telenet"
    elif "orange" in filename:
        provider = "Orange"
    else:
        raise ValueError("Provider could not be determined from filename")

    year_match = re.search(r'\d{4}', filename)
    if year_match:
        year = year_match.group(0)
    else:
        raise ValueError("Year could not be determined from filename")

    return provider, year


def load_page_selection():
    if os.path.exists(PAGE_SELECTION_FILE):
        with open(PAGE_SELECTION_FILE, "r") as file:
            return json.load(file)
    return {}


def save_page_selection(page_selection):
    with open(PAGE_SELECTION_FILE, "w") as file:
        json.dump(page_selection, file)


def get_pages_to_process(pdf_path):
    page_selection = load_page_selection()
    filename = os.path.basename(pdf_path)
    if filename in page_selection:
        return page_selection[filename]

    document = fitz.open(pdf_path)
    page_count = document.page_count
    document.close()

    if page_count == 1:
        return [1]

    while True:
        pages_input = input(
            f"The document '{filename}' has {page_count} pages. Which pages would you like to process (e.g., 1,3,5)? ")
        pages = [int(p.strip()) for p in pages_input.split(",") if
                 p.strip().isdigit() and 1 <= int(p.strip()) <= page_count]
        if pages:
            page_selection[filename] = pages
            save_page_selection(page_selection)
            return pages
        else:
            print(f"Invalid input. Please enter page numbers between 1 and {page_count}.")
