import re

VOO_INFO_CODES = {
    "w VS": "VOOsport World",
    "VS": "VOOsport",
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

def get_provider_and_year(filename):
    provider_names = ['voo', 'orange', 'telenet']
    provider = None
    year = None

    for name in provider_names:
        if name.lower() in filename.lower():
            provider = name.capitalize()
            break

    year_match = re.search(r'\d{4}', filename)
    if year_match:
        year = year_match.group(0)

    return provider, year

def read_section_names(file_path):
    with open(file_path, 'r', encoding='utf-8') as f:
        section_names = [line.strip() for line in f.readlines()]
    return section_names

def parse_tsv(tsv_path, section_names, provider):
    with open(tsv_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    data = []
    current_section = None

    for line in lines:
        stripped_line = line.strip()
        if stripped_line in section_names:
            current_section = stripped_line
        elif not stripped_line.isdigit() and not re.match(r'^\d{1,3}$', stripped_line):
            if current_section:
                if provider == "Voo":
                    stripped_line, additional_columns = handle_voo_info_codes(stripped_line)
                    data.append([current_section, stripped_line] + additional_columns)
                else:
                    data.append([current_section, stripped_line])

    return data

def handle_voo_info_codes(line):
    additional_columns = []
    words = line.split()
    filtered_words = []
    i = 0

    while i < len(words):
        matched = False
        for j in range(len(words), i, -1):
            potential_key = ' '.join(words[i:j])
            if potential_key in VOO_INFO_CODES:
                additional_columns.append(VOO_INFO_CODES[potential_key])
                i = j  # Skip the matched key
                matched = True
                break
        if not matched:
            filtered_words.append(words[i])
            i += 1

    filtered_line = ' '.join(filtered_words)
    return filtered_line, additional_columns

def adjust_region_columns(df):
    updated_index = []
    for text in df.index:
        if re.search(r'\sW(\s|$)', text):
            df.at[text, 'Region Flanders'] = 0
            df.at[text, 'Brussels'] = 0
            df.at[text, 'Region Wallonia'] = 1
            df.at[text, 'Communauté Germanophone'] = 0
            text = re.sub(r'\sW(\s|$)', ' ', text).strip()
        elif re.search(r'\sB(\s|$)', text):
            df.at[text, 'Region Flanders'] = 0
            df.at[text, 'Brussels'] = 1
            df.at[text, 'Region Wallonia'] = 0
            df.at[text, 'Communauté Germanophone'] = 0
            text = re.sub(r'\sB(\s|$)', ' ', text).strip()
        elif re.search(r'\sG(\s|$)', text):
            df.at[text, 'Region Flanders'] = 0
            df.at[text, 'Brussels'] = 0
            df.at[text, 'Region Wallonia'] = 0
            df.at[text, 'Communauté Germanophone'] = 1
            text = re.sub(r'\sG(\s|$)', ' ', text).strip()
        else:
            df.at[text, 'Region Flanders'] = 1
            df.at[text, 'Brussels'] = 1
            df.at[text, 'Region Wallonia'] = 1
            df.at[text, 'Communauté Germanophone'] = 1

        updated_index.append(text)

    df.index = updated_index
    return df
