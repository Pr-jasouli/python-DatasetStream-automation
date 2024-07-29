import os
import fitz

VOO_info_codes = {
    "VS": "VOOsport",
    "w VS": "VOOsport World",
    "Pa": "Bouquet Panorama",
    "Ci": "Option Ciné Pass",
    "Doc": "Be Bouquet Documentaires",
    "Div": "Be Bouquet Divertissement",
    "Co": "Be Cool",
    "Enf": "Be Bouquet Enfant",
    "Sp": "Be Bouquet Sport",
    "Sel": "Be Bouquet Selection",
    "Inf": "Option Infos",
    "Sen": "Option Sensation",
    "Ch": "Option Charme",
    "FF": "Family Fun",
    "DM": "Discover More",
    "CX": "Classé X",
    "MX": "Man-X",
    "B": "Bruxelles",
    "G": "Comm. German",
    "W": "Wallonie",
}

def read_section_names(file_path):
    with open(file_path, 'r', encoding='utf-8') as f:
        section_names = [line.strip() for line in f.readlines()]
    return section_names

def is_section_name_in_row(words, section_names):
    section_indices = []
    for i in range(len(words)):
        for section in section_names:
            section_words = section.split()
            if words[i:i + len(section_words)] == section_words:
                section_indices.append((i, i + len(section_words) - 1))
    return section_indices

def modify_row(row, section_names):
    words = row.split()
    last_valid_index = -1
    section_indices = is_section_name_in_row(words, section_names)

    for i, word in enumerate(words):
        if word in VOO_info_codes:
            last_valid_index = i

    if last_valid_index != -1 and section_indices:
        first_section_index = min(section_indices, key=lambda x: x[0])[0]
        return " ".join(words[:last_valid_index + 1]), " ".join(words[first_section_index:])
    else:
        return row, ""

def extract_text(pdf_path):
    document = fitz.open(pdf_path)
    text = []

    for i in range(document.page_count):
        page = document.load_page(i)
        blocks = page.get_text("dict")["blocks"]

        for block in blocks:
            if 'lines' in block:
                for line in block["lines"]:
                    for span in line["spans"]:
                        text.append(span["text"])

    return "\n".join(text)

def save_as_tsv(text, filename: str) -> None:
    output_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '../../outputs/text/'))
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    base_name = os.path.basename(filename)
    base_name_no_ext = os.path.splitext(base_name)[0]
    new_filename = base_name_no_ext + '_text.tsv'
    output_path = os.path.join(output_dir, new_filename)

    with open(output_path, 'w', encoding='utf-8') as f:
        for line in text.splitlines():
            f.write(line + '\n')

def clean_tsv(tsv_path):
    with open(tsv_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    cleaned_lines = []
    temp_line = ""
    last_number = None

    for line in lines:
        stripped_line = line.strip()
        if stripped_line.isdigit():
            if temp_line:
                cleaned_lines.append(temp_line.strip())
                temp_line = ""
            if last_number == stripped_line:
                temp_line += " " + stripped_line
            else:
                cleaned_lines.append(stripped_line)
            last_number = stripped_line
        else:
            temp_line += " " + stripped_line

    if temp_line:
        cleaned_lines.append(temp_line.strip())

    with open(tsv_path, 'w', encoding='utf-8') as f:
        for line in cleaned_lines:
            f.write(line + '\n')

def process_single_tsv(tsv_path, section_names):
    with open(tsv_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    if not lines:
        print(f"Aucun contenu à traiter dans {tsv_path}")
        return

    modified_lines = []
    for line in lines:
        modified_row, section_row = modify_row(line.strip(), section_names)
        modified_lines.append(modified_row + "\n")
        if section_row:
            modified_lines.append(section_row + "\n")

    with open(tsv_path, 'w', encoding='utf-8') as f:
        f.writelines(modified_lines)

    print(f"Traitement et enregistrement de {tsv_path} terminé")

def insert_section_name_rows(tsv_path, section_names):
    with open(tsv_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    new_lines = []
    for line in lines:
        words = line.strip().split()
        section_indices = is_section_name_in_row(words, section_names)
        if section_indices:
            for start, end in section_indices:
                section_name = " ".join(words[start:end + 1])
                if section_name.strip() != line.strip():
                    # Remove the section name from its original position
                    new_line = " ".join(words[:start] + words[end + 1:]).strip()
                    if new_line:
                        new_lines.append(new_line + "\n")
                    new_lines.append(section_name + "\n")
                else:
                    new_lines.append(line)
        else:
            new_lines.append(line)

    with open(tsv_path, 'w', encoding='utf-8') as f:
        f.writelines(new_lines)

    print(f"Insertion des noms de section terminée pour {tsv_path}")

def remove_specific_string(tsv_path, target_string):
    with open(tsv_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    cleaned_lines = []
    i = 0
    while i < len(lines):
        if target_string in lines[i]:
            if i > 0:
                cleaned_lines.pop()
            i += 1
            if i < len(lines):
                i += 0
        else:
            cleaned_lines.append(lines[i])
        i += 1

    with open(tsv_path, 'w', encoding='utf-8') as f:
        f.writelines(cleaned_lines)

def remove_everything_after_word(tsv_path, target_word):
    with open(tsv_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    cleaned_lines = []
    for line in lines:
        if target_word in line:
            index = line.find(target_word)
            cleaned_lines.append(line[:index].strip() + "\n")
            break
        cleaned_lines.append(line)

    with open(tsv_path, 'w', encoding='utf-8') as f:
        f.writelines(cleaned_lines)

def parse_voo_pdf(pdf_path):
    text = extract_text(pdf_path)
    save_as_tsv(text, pdf_path)
    tsv_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '../../outputs/text/', os.path.splitext(os.path.basename(pdf_path))[0] + '_text.tsv'))
    clean_tsv(tsv_path)
    print(f"Sauvegardé et nettoyé {tsv_path}")

    section_tsv_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '../../outputs/section/', os.path.splitext(os.path.basename(pdf_path))[0] + '_sections.tsv'))
    if os.path.exists(section_tsv_path):
        section_names = read_section_names(section_tsv_path)
        process_single_tsv(tsv_path, section_names)
        insert_section_name_rows(tsv_path, section_names)

    remove_specific_string(tsv_path, "Retrouvez votre chaîne locale ici")
    remove_everything_after_word(tsv_path, "Retrouvez")

    parse_long_lines(tsv_path)

def parse_long_lines(tsv_path):
    with open(tsv_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    processed_lines = []
    for line in lines:
        if len(line.strip()) > 15:
            processed_lines.extend(split_long_line(line))
        else:
            processed_lines.append(line)

    with open(tsv_path, 'w', encoding='utf-8') as f:
        f.writelines(processed_lines)

def split_long_line(line):
    words = line.split()
    new_lines = []
    current_line = []

    for i, word in enumerate(words):
        current_line.append(word)
        if word in VOO_info_codes and (i + 1 < len(words) and words[i + 1] not in VOO_info_codes):
            new_lines.append(" ".join(current_line) + "\n")
            current_line = []

    if current_line:
        new_lines.append(" ".join(current_line) + "\n")

    return new_lines

def remove_following_lines(lines, start_string):
    for i, line in enumerate(lines):
        if line.startswith(start_string):
            return lines[:i]
    return lines
