По умолчанию программа берет значения из "info.txt" и меняет пдэфы на "target.pdf". Строка в таблице "Перечень_документов_отп_картаплан" выбирается по Nn=6.
В "info.txt" должны содержаться ровно четыре строки:
(Наименование документа)
(Номер документа)
(Дата документа)
(Автор документа)

У скрипта есть несколько опциональных аргументов, которые инициализируются через командную строку.
-i {info.txt} - меняет наименования документа с настройками
-t {target.pdf} - меняет наименование pdf файла, который надо добавить во все БД
-c - "вытаскивает" pdf файл из БД в соответствующую папку
d - отключает основной функционал программы. Создан для комбинирования с предыдущим аргументом
Область прикрепленных файлов