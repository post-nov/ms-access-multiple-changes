import pyodbc
import os


def write_file_from_db(crsr, mdb_file, ID):
    directory = os.path.dirname(mdb_file)
    crsr.execute(f"""
                select Изображение
                from Перечень_документов_отп_картаплан
                where Nn = {ID}
                 """)

    file = (crsr.fetchall()[0][0])

    with open(directory + '/test.pdf', 'wb') as test:
        test.write(file)


def path_list_mdb():
    cur_dir_path = os.getcwd()
    list_of_files = []
    for root, dirs, files in os.walk(cur_dir_path):
        for file in files:
            if file.endswith(".mdb"):
                list_of_files.append(os.path.join(root, file))
    return list_of_files


def main():
    ID = input("Nn таблицы ?:")
    for mdb_file in path_list_mdb():
        print("Обрабатывается", mdb_file)
        pth = ("DBQ=" + mdb_file)
        conn_str = (r'DRIVER={Microsoft Access Driver (*.mdb)};' + (pth))
        cnxn = pyodbc.connect(conn_str)
        crsr = cnxn.cursor()
        write_file_from_db(crsr, mdb_file, ID)
        print("Таблица обработана\n")


main()
