import os
import fitz

def extract_text(pdf_path):
    document = fitz.open(pdf_path)
    colored_text = []
    sizes = set()

    colors = [16777215, 14092940]

    for i in range(document.page_count):
        page = document.load_page(i)
        blocks = page.get_text("dict")["blocks"]

        for block in blocks:
            if 'lines' in block:
                for line in block["lines"]:
                    for span in line["spans"]:
                        if span["color"] == 16777215:
                            sizes.add(span["size"])

    for i in range(document.page_count):
        page = document.load_page(i)
        blocks = page.get_text("dict")["blocks"]

        for block in blocks:
            if 'lines' in block:
                for line in block["lines"]:
                    for span in line["spans"]:
                        if span["color"] in colors and span["size"] in sizes:
                            colored_text.append(span["text"])

    return "\n".join(colored_text)

def parse(text):
    lines = text.splitlines()
    sections = []
    current_section = None
    for line in lines:
        words = line.split()
        if len(words) == 0:  # Skip empty lines
            continue
        if 1 < len(words) <= 3 and not any(char.isdigit() for char in line):
            if current_section:
                sections.append(current_section)
            current_section = [line.strip()]  # Start a new section
        else:
            if current_section is not None:
                current_section.append(line.strip())  # Add channels to the current section
    if current_section:
        sections.append(current_section)
    return sections

def write_section_tsv(file, sections):
    with open(file, 'w', encoding='utf-8') as f:
        for section in sections:
            f.write('\t'.join([section[0]]) + '\n')  # Write only section names to SECTION.tsv

def process(folder_path):
    for file in os.listdir(folder_path):
        if file.endswith(".pdf"):
            path = os.path.join(folder_path, file)
            text = extract_text(path)
            sections = parse(text)
            section_tsv_file = os.path.join(folder_path, os.path.splitext(file)[0] + "a.tsv")
            write_section_tsv(section_tsv_file, sections)
            print(f"Saved {section_tsv_file}")

if __name__ == "__main__":
    folder = '.'
    process(folder)
