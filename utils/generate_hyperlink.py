import pandas as pd
from datetime import datetime
import os
import configparser

config = configparser.ConfigParser()
config.read('config.ini')
country_code = config.get('whatsapp', 'COUNTRY_CODE').strip()
phone_number_len = config.get('whatsapp', 'PHONE_NUMBER_LEN').strip()
anchor_text = config.get('whatsapp', 'ANCHOR_TEXT').strip()

def excel_column_letter_to_index(col_letter):
    col_letter = col_letter.upper()
    num = 0
    for i in range(len(col_letter)):
        num = num * 26 + (ord(col_letter[i]) - ord("A") + 1)
    return num - 1

def load_data(file_path):
    return pd.read_excel(file_path)

def filter_data(data_frame, branch_name, branch_index):
    if branch_index == None or excel_column_letter_to_index(branch_index) <= -1:
        return data_frame.copy(), True
    else:
        return data_frame[data_frame.iloc[:, excel_column_letter_to_index(branch_index)].astype(str).str.strip() == branch_name].copy(), False

def write_to_excel(branch_data, branch_file_path):
    writer = pd.ExcelWriter(branch_file_path, engine='xlsxwriter')
    branch_data.to_excel(writer, sheet_name='Sheet1', index=False)
    return writer

def add_hyperlink_formula(worksheet, branch_data, mob_index, hyperlink_index, msg_index):
    for row_index in range(2, len(branch_data) + 2):
        msg_index__ = msg_index.replace('#', str(row_index))
        hyperlink_formula = f'=HYPERLINK("https://wa.me/"& IF(LEN({mob_index}{row_index})={phone_number_len},"{country_code}"&{mob_index}{row_index},{mob_index}{row_index})&"?text="&TRIM(CONCATENATE({msg_index__})),"{anchor_text}")'
        worksheet.write_formula(f'{hyperlink_index}{row_index}', hyperlink_formula)

def adjust_column_width(worksheet, branch_data):
    for col_index, col_name in enumerate(branch_data.columns):
        max_length = max(branch_data[col_name].astype(str).apply(len).max(), len(str(col_name)))  # Calculate max length of column
        max_length = min(max_length, 255)  # Max width is 255 characters
        worksheet.set_column(col_index, col_index, max_length)

def save_to_excel(data, file_path, mob_index, hyperlink_index, msg_index):
    writer = write_to_excel(data, file_path)
    worksheet = writer.sheets['Sheet1']
    add_hyperlink_formula(worksheet, data, mob_index, hyperlink_index, msg_index)
    adjust_column_width(worksheet, data)
    writer._save()

def generate_file_path(branch_name, file_path, output_path, i, no_branch_col):
    base_name = os.path.splitext(os.path.basename(file_path))[0]
    file_name = f"{'_ALL_' if no_branch_col else branch_name}_{i+1 if no_branch_col else ''}__[{base_name}]_{datetime.now().strftime("%d-%m-%Y")}.xlsx"
    return os.path.join(output_path, file_name)

def process_branch_data(data_frame, branch_name, file_path, output_path, branch_index, mob_index, hyperlink_index, msg_index, chunk_size=200):
    branch_data, no_branch_col = filter_data(data_frame, branch_name, branch_index)
    if no_branch_col:
        print(f"Chunking data for branches by {chunk_size} rows per file...")
        chunks = [branch_data[i:i+chunk_size] for i in range(0, branch_data.shape[0], chunk_size)]
        for i, chunk in enumerate(chunks):
            branch_file_path = generate_file_path(branch_name, file_path, output_path, i, no_branch_col)
            save_to_excel(chunk, branch_file_path, mob_index, hyperlink_index, msg_index)
        return True
    else:
        branch_file_path = generate_file_path(branch_name, file_path, output_path, 0, no_branch_col)
        save_to_excel(branch_data, branch_file_path, mob_index, hyperlink_index, msg_index)
        return False

if __name__ == "__main__":
    pass

    # file_path = r"Sample.xlsx"
    # output_path = r"Output/"
    # data_frame = load_data(file_path)
    # branch_index="A"
    # mob_index = "B"
    # hyperlink_index = "X"

    # # Branch Names (e.g. 512, 523, 524, 529, 537, 538, 543, 5001, 5011, 51E, 51F, 52B, 52D, 53B)
    # branch_names = [
    #     "512", "523", "524", "529", "537", "538", "543", "5001", "5011",
    #     "51E", "51F", "52B", "52D", "53B"
    # ]

    # msg_index = 'C# ,D# ,E#, F#, G#,' # Alphabate must be at first followed by Number
    # for branch_name in branch_names:
    #     no_branch_col = process_branch_data(data_frame, branch_name, file_path,  output_path, branch_index, mob_index, hyperlink_index, msg_index[:-1], chunk_size=1000)
    #     if no_branch_col:
    #         break
    