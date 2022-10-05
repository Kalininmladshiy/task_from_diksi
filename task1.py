import os
import pandas
import shutil
import smtplib
import random
from pathlib import Path
from datetime import datetime
from email.message import EmailMessage
from dotenv import load_dotenv
from PIL import Image, ImageDraw



def create_folders(folder_path, folder_names):
    for folder_name in folder_names:
        os.mkdir(Path() / folder_path / folder_name)


def create_files(excel_data, path_to_docs):
    for document in excel_data:
        code_ka = document['Код_КА']
        document_types = document['Документы для формирования'].split(',')
        document_date = document['Дата документа']
        for document_type in document_types:
            img = Image.new('L', (1000, 1000), 'white')
            filename = f'КА_{code_ka}_{document_type}_{document_date}.png'
            img.save(Path() / path_to_docs / filename)
            img.close()


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


def send_mail_with_random_file(
    sender_email,
    recipient_email,
    subject,
    email_body,
    path_to_folder,
    file_type,
    email_password,
     ):
    file_name = random.choice(get_files_names(path_to_folder))
    past_stamp(path_to_folder, file_name)
    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg.set_content(email_body)
    with open(Path() / path_to_folder / file_name, 'rb') as f:
        file_data = f.read()
    msg.add_attachment(file_data, maintype="application", subtype=file_type, filename=file_name)
    with smtplib.SMTP_SSL('smtp.mail.ru', 465) as smtp:
        smtp.login(sender_email, email_password)
        smtp.send_message(msg)


def past_stamp(path_to_folder, img_filename):
    img2 = Image.open('stamp.png')
    img = Image.open(Path() / path_to_folder / img_filename)
    img.paste(img2, (30, 700))
    img.save(Path() / path_to_folder / img_filename)
    img.close()
    img2.close()


if __name__ == '__main__':
    load_dotenv()
    sender_email = os.environ['SENDER_EMAIL']
    email_password = os.environ['EMAIL_PASSWORD']
    recipient_email = 'kalinin.maksim85@yandex.ru'
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
    subject = 'Документы от Dixi'
    email_body = 'Уважаемый партнер, высылаем вам документы во вложении!'
    file_type = 'png'
    send_mail_with_random_file(
        sender_email,
        recipient_email,
        subject,
        email_body,
        path_to_docs,
        file_type,
        email_password,
     )
