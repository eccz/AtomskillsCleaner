# Atomskills Cleaner
Утилита решает вечно повторяющуюся проблему с пробелами и знаками переноса строки в конце значений атрибутов
при выполнении конкурсного задания на Atomskills в компетенции Инженерное проектирование.

## Технологии
- Интерфейс - [tkinter](https://docs.python.org/3/library/tkinter.html/)
- Работа с IFC - [IfcOpenShell](https://ifcopenshell.org/)
- Работа с Excel - [openpyxl](https://pypi.org/project/openpyxl/)

## Использование
Можно клонировать репозиторий, установить библиотеки из requirements.txt и пользоваться.\
Можно скачать .exe из releases и пользоваться.\
Собирается в один файл с помощью pyinstaller:

`pip install pyinstaller`\
`pyinstaller --onefile --windowed --icon=logo2.ico app.py`

## Описание работы 
* Excel - проходит по всем непустым ячейкам и заменяет их на очищенные .strip() значения при необходимости.
* IFC - проходит по всем IfcPropertySingleValue и заменяет их на очищенные .strip() значения при необходимости.
* Cadlib - проходит по всем значениям в таблице dbo.Parameters_STR и заменяет их на очищенные .strip() значения при необходимости.
С Cadlib использовать на свой страх и риск.
 