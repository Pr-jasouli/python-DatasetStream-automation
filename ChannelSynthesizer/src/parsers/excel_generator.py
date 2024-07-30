from pathlib import Path
import pandas as pd
from ChannelSynthesizer.src.utils import get_provider_and_year, read_section_names, parse_tsv, adjust_region_columns

def filter_columns_with_values(df):
    non_empty_columns = df.loc[:, (df != 0).any(axis=0)].columns
    return df[non_empty_columns]

def create_consolidated_excel(all_data, output_path):
    print("Creating consolidated Excel...")

    combined_data = []
    combined_columns = ['Provider_Period', 'Channel']

    combined_data_by_provider_year = {}
    for provider, year, data, section_names in all_data:
        key = (provider, year)
        if key not in combined_data_by_provider_year:
            combined_data_by_provider_year[key] = {'data': [], 'section_names': set()}
        combined_data_by_provider_year[key]['data'].extend(data)
        combined_data_by_provider_year[key]['section_names'].update(section_names)

    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        for (provider, year), value in combined_data_by_provider_year.items():
            period = f"{provider} {year}"
            data = value['data']
            section_names = list(value['section_names'])

            text_rows = list(set(text for section, text in data))
            columns = ['Region Flanders', 'Brussels', 'Region Wallonia', 'Communauté Germanophone'] + section_names

            df = pd.DataFrame(0, index=text_rows, columns=columns)

            for section, text in data:
                if section in df.columns:
                    df.at[text, section] = 1

            df = adjust_region_columns(df)

            df.insert(0, 'Provider_Period', period)
            df = df.reset_index().rename(columns={'index': 'Channel'})

            filtered_df = filter_columns_with_values(df)

            sheet_name = f"{provider}_{year}"
            filtered_df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=0)

            workbook = writer.book
            worksheet = writer.sheets[sheet_name]

            header_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter'})

            for col_num, value in enumerate(filtered_df.columns.values):
                worksheet.write(0, col_num, value, header_format)

            for i, col in enumerate(filtered_df.columns, 0):
                max_length = max((filtered_df[col].astype(str).map(len).max(), len(col))) + 2
                worksheet.set_column(i, i, max_length)

            worksheet.autofilter(0, 0, 0, len(filtered_df.columns) - 1)

        combined_data = []
        for (provider, year), value in combined_data_by_provider_year.items():
            period = f"{provider} {year}"
            data = value['data']
            section_names = list(value['section_names'])

            text_rows = list(set(text for section, text in data))
            columns = ['Region Flanders', 'Brussels', 'Region Wallonia', 'Communauté Germanophone'] + section_names

            df = pd.DataFrame(0, index=text_rows, columns=columns)

            for section, text in data:
                if section in df.columns:
                    df.at[text, section] = 1

            df = adjust_region_columns(df)

            df.insert(0, 'Provider_Period', period)
            df = df.reset_index().rename(columns={'index': 'Channel'})
            combined_data.append(df)

        if combined_data:
            final_df = pd.concat(combined_data)
        else:
            final_df = pd.DataFrame(columns=combined_columns + columns)

        final_df.to_excel(writer, sheet_name='Consolidated', index=False, startrow=0)

        workbook = writer.book
        worksheet = writer.sheets['Consolidated']

        bold_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'font_size': 14})
        header_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter'})

        worksheet.autofilter(0, 0, 0, len(final_df.columns) - 1)

    print(f"Consolidated Excel created at: {output_path}")


def find_file_pairs(section_dir, text_dir):
    print(f"Finding file pairs in {section_dir} and {text_dir}")
    section_files = list(Path(section_dir).glob('*_sections.tsv'))
    text_files = list(Path(text_dir).glob('*.tsv'))

    print(f"Section files found: {[f.name for f in section_files]}")
    print(f"Text files found: {[f.name for f in text_files]}")

    file_pairs = []
    for section_file in section_files:
        base_name = section_file.stem.replace('_sections', '')
        text_file = next((text_file for text_file in text_files if base_name in text_file.stem and '_text' in text_file.stem), None)
        if text_file and text_file.exists():
            file_pairs.append((section_file, text_file))

    print(f"Found {len(file_pairs)} file pairs.")
    return file_pairs

def process_providers(section_dir, text_dir, output_path):
    file_pairs = find_file_pairs(section_dir, text_dir)
    all_data = []

    for section_file, text_file in file_pairs:
        provider, year = get_provider_and_year(text_file.stem)
        section_names = read_section_names(section_file)
        data = parse_tsv(text_file, section_names)
        all_data.append((provider, year, data, section_names))

    create_consolidated_excel(all_data, output_path)
