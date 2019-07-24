import pyodbc
import os
import argparse
import random


def parse():
    """У программы есть дополнительная возможность поменять таблицу "Титульный_картаплан",
    Для этого необходимо добавить аргумент '-c' или '--change' при ее инициализации."""
    global changing_needed

    parser = argparse.ArgumentParser()
    parser.add_argument('-c', '--change',
                        help='also changes main table',
                        action='store_true',
                        default=False)
    args = parser.parse_args()

    changing_needed = args.change


def get_source_data():
    """
    Собирает необходимые данные для дальнейшей работы. Пользователю необходимо
    ввести данные, уникальные для заменяемой строки таблицы,
    а затем выбрать файлы, в которых содержаться данные на замену устаревшим.
    """
    os.system('cls')
    print("\n\nПриветствую тебя, дорогой %USER_NAME%. \nСейчас мы будем менять твои базы. Для начала давай введем значения,",
          "которые уникальны для изменяемой строки.\n")
    if changing_needed:
        print('Напоминаю, что программа запущена с дополнительным аргументом.')
        print('Это значит, помимо таблицы "Перечень_документов_отп_картаплан", поменяется еще и "Титульный_картаплан"\n')
    source_date = input("Дата_выдачи (ДД.ММ.ГГ):")
    if source_date == "":
        exit()
    source_author = input("Автор:")
    if source_author == "":
        exit()

    print("\nДля удобства давай раз и навсегда присвоем этой строке уникальный Nn")
    new_nn = input("Nn:")
    if new_nn == "":
        new_nn = random.randint(1, 100)
        print(f"Ты ничего не указал, поэтому Nn будет {new_nn}")

    print("\nТеперь давай напишем название файла в корневой папке, содержащей необходимые изменения.")
    print(list(filter(lambda x: x.endswith(".txt"), os.listdir())))
    source_file = input("Название .txt файла:")

    print("\nИ под конец введи, пожалуйста, название .pdf документа под замену.")
    print(list(filter(lambda x: x.endswith(".pdf"), os.listdir())))
    source_pdf = input("Название .pdf файла:")

    print("\nПоехали...\n")

    return {'source_date': source_date,
            'source_author': source_author,
            'source_file': source_file,
            'source_pdf': source_pdf,
            'new_nn': new_nn}


def get_target_data(source_file, source_pdf):
    "Собирает данные на замену."
    with open(source_file, 'r') as info:
        order_name = info.readline().rstrip("\n")
        order_number = info.readline().rstrip("\n")
        order_date = info.readline().rstrip("\n")
        order_author = info.readline().rstrip("\n")
    with open(source_pdf, 'rb') as pdf:
        target_pdf = pdf.read()
    return {'order_name': order_name,
            'order_number': order_number,
            'order_date': order_date,
            'order_author': order_author,
            'target_pdf': target_pdf}


def path_list_mdb():
    """Создает лист всех баз данных, содержащихся в рабочем каталоге,
    включая подкаталоги"""
    cur_dir_path = os.getcwd()
    list_of_dbs = []
    for root, dirs, files in os.walk(cur_dir_path):
        for file in files:
            if file.endswith(".mdb"):
                list_of_dbs.append(os.path.join(root, file))
    return list_of_dbs


def change_mdb_data(crsr, source_data, target_data):
    "Заменяет исходные данные в базе данных на данные, предоставленные пользователем"
    crsr.execute(f"""
            UPDATE Перечень_документов_отп_картаплан
            SET Наименование = ?, Номер = ?, Дата_выдачи = ?, Автор = ?, Nn = ?, Изображение = ?
            WHERE Дата_выдачи = CDate('{source_data['source_date]'}')
            AND Автор LIKE '{source_data['source_author']}'
            """, target_data['order_name'],
                 target_data['order_number'],
                 target_data['order_date'],
                 target_data['order_author'],
                 target_data['target_pdf'])
    if changing_needed:
        crsr.execute("""
                update Титульный_картаплан
                set Наименование_документа = ?, Номер_документа = ?, Дата_документа = ?, Автор_документа = ?
                where ID = 1
                """, [target_data['order_name'],
                      target_data['order_number'],
                      target_data['order_date'],
                      target_data['order_author']])
    crsr.commit()


def changes_checker(crsr, new_nn):
    "Проверяет успешность выполнения программы. На всякий случай."
    crsr.execute(f"""
                  SELECT COUNT(1)
                  FROM Перечень_документов_отп_картаплан
                  WHERE Nn = {new_nn}
                  """)
    checker = crsr.fetchall()
    if checker[0] == 1:
        print("Таблица обработана\n")
    elif checker[0] > 1:
        print("В ТАБЛИЦЕ ИЗМЕНЕНЫ НЕСКОЛЬКО СТРОК!\n")
    else:
        print("ВНИМАНИЕ ТАБЛИЦА НЕ ОБРАБОТАНА!\n")


def main():

    source_data = get_source_data()
    target_data = get_target_data(source_data['source_file'],
                                  source_data['source_pdf'])
    list_of_dbs = path_list_mdb()

    for mdb_file in list_of_dbs:
        print("Обрабатывается", mdb_file)
        pth = ("DBQ=" + mdb_file)
        conn_str = (r'DRIVER={Microsoft Access Driver (*.mdb)};' + (pth))
        cnxn = pyodbc.connect(conn_str)
        crsr = cnxn.cursor()
        change_mdb_data(crsr, source_data, target_data)
        changes_checker(crsr, target_data['new_nn'])


while True:
    # Программа выполняется до тех пор, пока не заменятся все требуемые строки
    parse()

    main()
    answer = input('\n\nПовторить для нового файла? (y, yes)')
    if answer is ('y' or 'yes'):
        changing_needed = False
        continue
    else:
        print('\nСпасибо за работу :3')
        break
