import os

# Dictionnaire des codes d'informations supplémentaires
info_codes = {
    "VS": "VOOsport",
    "w VS": "VOOsport World",
    "Pa": "Bouquet Panorama",
    "Ci": "Option Ciné Pass",
    "Doc": "Be Bouquet Documentaires",
    "Div": "Be Bouquet Divertissement",
    "Co Be": "Cool",
    "Enf Be": "Bouquet Enfant",
    "Sp Be": "Bouquet Sport",
    "Sel Be": "Bouquet Selection",
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

# Fonction pour lire les noms de section depuis un fichier
def read_section_names(file_path):
    with open(file_path, 'r', encoding='utf-8') as f:
        section_names = [line.strip() for line in f.readlines()]
    return section_names

# Fonction pour vérifier si des noms de section se trouvent dans une ligne donnée
def is_section_name_in_row(words, section_names):
    section_indices = []
    section_lengths = {name: len(name.split()) for name in section_names}

    for i in range(len(words)):
        for section in section_names:
            section_words = section.split()
            if words[i:i + len(section_words)] == section_words:
                section_indices.append((i, i + len(section_words) - 1))
    return section_indices

# Fonction pour modifier une ligne en fonction des codes d'information et des noms de section
def modify_row(row, section_names):
    words = row.split()
    last_valid_index = -1

    section_indices = is_section_name_in_row(words, section_names)

    # Identifier le dernier index valide d'un code reconnu
    for i, word in enumerate(words):
        if word in info_codes:
            last_valid_index = i
            # Vérifier si le mot suivant n'est pas un code ou la valeur d'un code
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

# Fonction pour traiter un seul fichier TSV
def process_single_tsv(tsv_path, section_names):
    output_path = tsv_path.replace("b.tsv", "c.tsv")

    with open(tsv_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    if not lines:
        print(f"Aucun contenu à traiter dans {tsv_path}")
        return

    modified_lines = [modify_row(line.strip(), section_names) + "\n" for line in lines]

    with open(output_path, 'w', encoding='utf-8') as f:
        f.writelines(modified_lines)

    print(f"Traitement et enregistrement de {output_path} terminé")

# Fonction pour traiter tous les fichiers TSV dans un répertoire donné
def process_tsv_files(directory):
    for file in os.listdir(directory):
        if file.endswith(".pdf"):
            base_name = os.path.splitext(file)[0]
            a_tsv_path = os.path.join(directory, f"{base_name}a.tsv")
            b_tsv_path = os.path.join(directory, f"{base_name}b.tsv")
            if os.path.exists(a_tsv_path) and os.path.exists(b_tsv_path):
                section_names = read_section_names(a_tsv_path)
                process_single_tsv(b_tsv_path, section_names)

# Point d'entrée principal
if __name__ == "__main__":
    current_directory = '.'
    process_tsv_files(current_directory)
