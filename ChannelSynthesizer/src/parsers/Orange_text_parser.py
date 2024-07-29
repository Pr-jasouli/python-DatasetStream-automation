import os
import fitz

def extract_text(pdf_path, min_font_size=8.0):
    """
    Extrait le texte d'un fichier PDF en filtrant les textes selon une taille de police minimale.

    Arguments:
    pdf_path -- le chemin du fichier PDF
    min_font_size -- la taille de police minimale à inclure (par défaut 8.0)

    Retourne:
    Le texte extrait du PDF sous forme de chaîne.
    """
    document = fitz.open(pdf_path)
    text = []

    for i in range(document.page_count):
        page = document.load_page(i)
        blocks = page.get_text("dict")["blocks"]

        for block in blocks:
            if 'lines' in block:
                for line in block["lines"]:
                    for span in line["spans"]:
                        if span['size'] >= min_font_size:
                            text.append(span['text'])

    return "\n".join(text)

def clean_text(text):
    """
    Nettoie le texte extrait en supprimant les lignes vides, les lignes contenant uniquement 'app',
    les lignes de plus de 35 caractères, et les lignes commençant par 'Optie' ou '(1)'.

    Arguments:
    text -- le texte extrait du PDF

    Retourne:
    Le texte nettoyé sous forme de chaîne.
    """
    cleaned_lines = []
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        if line.lower() == 'app':
            continue
        if len(line) > 35:
            continue
        if line.startswith("Optie") or line.startswith("(1)"):
            continue
        cleaned_lines.append(line)
    return "\n".join(cleaned_lines)

def save_as_tsv(text, filename: str) -> None:
    """
    Enregistre le texte nettoyé dans un fichier TSV.

    Arguments:
    text -- le texte nettoyé à enregistrer
    filename -- le nom du fichier PDF original pour générer le nom du fichier TSV
    """
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

    print(f"Saved TSV to {output_path}")

def parse_orange_pdf(pdf_path, min_font_size=8.0):
    """
    Extrait le texte d'un PDF, nettoie le texte et l'enregistre dans un fichier TSV.

    Arguments:
    pdf_path -- le chemin du fichier PDF
    min_font_size -- la taille de police minimale à inclure (par défaut 8.0)
    """
    print(f"Extracting text from {pdf_path} with minimum font size {min_font_size}")
    text = extract_text(pdf_path, min_font_size)
    cleaned_text = clean_text(text)
    save_as_tsv(cleaned_text, pdf_path)