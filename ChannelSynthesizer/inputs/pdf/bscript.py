import os
import fitz

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
                temp_line += " " + stripped_line  # Treat the second occurrence as part of the following text row
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

def process_pdfs(directory):
    for filename in os.listdir(directory):
        if filename.endswith(".pdf"):
            pdf_path = os.path.join(directory, filename)
            text = extract_text(pdf_path)
            tsv_path = os.path.join(directory, os.path.splitext(filename)[0] + "b.tsv")
            save_as_tsv(text, tsv_path)
            clean_tsv(tsv_path)
            print(f"Saved and cleaned {tsv_path}")

if __name__ == "__main__":
    current_directory = '.'
    process_pdfs(current_directory)
