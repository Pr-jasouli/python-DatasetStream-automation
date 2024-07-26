import os
import fitz

info_codes = {
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
    section_lengths = {name: len(name.split()) for name in section_names}

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
        if word in info_codes:
            last_valid_index = i
            if i + 1 < len(words):
                next_word = words[i + 1]
                if next_word not in info_codes and next_word not in info_codes.values():
                    break

    if last_valid_index != -1:
        if section_indices:
            # Inclure les noms de section et supprimer tout ce qui se trouve entre le dernier code et le premier nom de section
            first_section_index = min(section_indices, key=lambda x: x[0])[0]
            return " ".join(words[:last_valid_index + 1] + words[first_section_index:])
        else:
            return " ".join(words[:last_valid_index + 1])
    else:
        return row

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

def save_as_tsv(text, tsv_path):
    with open(tsv_path, 'w', encoding='utf-8') as f:
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
                # Traiter la deuxième occurrence comme partie de la ligne de texte suivante
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

    modified_lines = [modify_row(line.strip(), section_names) + "\n" for line in lines]

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

def process_pdfs(directory):
    for filename in os.listdir(directory):
        if filename.endswith(".pdf"):
            pdf_path = os.path.join(directory, filename)
            text = extract_text(pdf_path)
            tsv_path = os.path.join(directory, os.path.splitext(filename)[0] + "b.tsv")
            save_as_tsv(text, tsv_path)
            clean_tsv(tsv_path)
            print(f"Sauvegardé et nettoyé {tsv_path}")
            # Lire les noms de section à partir du fichier .tsv associé
            a_tsv_path = os.path.join(directory, os.path.splitext(filename)[0] + "a.tsv")
            if os.path.exists(a_tsv_path):
                section_names = read_section_names(a_tsv_path)
                process_single_tsv(tsv_path, section_names)
                insert_section_name_rows(tsv_path, section_names)


if __name__ == "__main__":
    current_directory = '.'
    process_pdfs(current_directory)
