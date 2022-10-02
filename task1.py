import os
import pandas
import shutil
from pathlib import Path
from datetime import datetime



def create_folders(folder_path, folder_names):
    for folder_name in folder_names:
        os.mkdir(Path() / folder_path / folder_name)


def create_files(excel_data, path_to_docs):
    for document in excel_data:
        code_ka = document['Код_КА']
        document_types = document['Документы для формирования'].split(',')
        document_date = document['Дата документа']
        for document_type in document_types:
            filename = f'КА_{code_ka}_{document_type}_{document_date}.txt'
            with open (Path() / path_to_docs / filename, 'w', encoding='utf-8') as file:
                pass


def get_files_names(folder_path):
    file_paths = [file.path for file in os.scandir(folder_path) if not file.is_dir()]
    files_names = []
    for file_path in file_paths:
        path, file = os.path.split(file_path)
        files_names.append(file)
    return files_names


def get_subfolders_paths(folder_path):
    subfolder_paths = [folder.path for folder in os.scandir(folder_path) if folder.is_dir()]
    return subfolder_paths


def sorted_by_date(filename, path_to_folder, path_to_folder_for_copy):
    file_path = os.path.join(path_to_folder, filename)
    document_date = filename.split('_')[-1][:-4]
    date_start = datetime(2020, 6, 20)
    date_end = datetime(2020, 7, 10)
    try:
        if date_start <= datetime.strptime(document_date, '%d.%m.%Y') <= date_end:
            shutil.copy(file_path, path_to_folder_for_copy)
    except ValueError:
        print(f'Проверьте даты в документе {filename}')


def sorted_by_code_ka(filename, path_to_folder, path_to_folder_for_copy):
    file_path = os.path.join(path_to_folder, filename)
    code_ka = filename.split('_')[2]
    last_digits_in_code_ka = code_ka[-2:]
    if last_digits_in_code_ka == last_digits_in_code_ka[::-1]:
        shutil.copy(file_path, path_to_folder_for_copy)


def sorted_by_column_d(excel_data, filename, path_to_folder, path_to_folder_for_copy):
    file_path = os.path.join(path_to_folder, filename)
    code_ka_in_filename = filename.split('_')[2]
    for document in excel_data:
        code_ka_in_excel = document['Код_КА'].split('_')[1]
        chek_number = document['Отправить']
        if code_ka_in_filename == code_ka_in_excel and chek_number == 1:
            shutil.copy(file_path, path_to_folder_for_copy)


def is_file_in_folder(path_to_folder, filename):
    file_path = os.path.join(path_to_folder, filename)
    return os.path.exists(file_path)


if __name__ == '__main__':
    os.mkdir('documents')
    path_to_docs = Path.cwd() / 'documents'
    excel_data = pandas.read_excel('Задание.xlsx', sheet_name='Лист1').to_dict(
        orient='record',
     )
    folder_names = ['folder_1', 'folder_2', 'folder_3']
    create_folders(path_to_docs, folder_names)
    create_files(excel_data, path_to_docs)
    filenames = get_files_names(path_to_docs)
    folder_3, folder_1, folder_2 = get_subfolders_paths(path_to_docs)
    for filename in filenames:
        sorted_by_column_d(excel_data, filename, path_to_docs, folder_1)
        sorted_by_code_ka(filename, path_to_docs, folder_2)
        sorted_by_date(filename, path_to_docs, folder_3)
        file_path = os.path.join(path_to_docs, filename)
        if (is_file_in_folder(folder_1, filename) or
            is_file_in_folder(folder_2, filename) or 
            is_file_in_folder(folder_3, filename)):
            os.remove(file_path)
