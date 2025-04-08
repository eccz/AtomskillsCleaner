import ifcopenshell
import os
import csv
from datetime import datetime


def build_report_path(f_path, suffix="_clean_report", ext=".csv", custom_name=None):
    # Получаем только директорию исходного файла
    dir_name = os.path.dirname(f_path)

    file_name = os.path.basename(f_path)
    name_without_ext = os.path.splitext(file_name)[0]  # Убираем расширение файла

    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    if custom_name:
        return os.path.join(dir_name, f"{custom_name}{suffix}_{timestamp}{ext}")
    else:
        return os.path.join(dir_name, f"{name_without_ext}{suffix}_{timestamp}{ext}")


def ifc_cleaner(inp_filepath, out_filepath):
    f = ifcopenshell.open(inp_filepath)
    single_value_props = f.by_type('IfcPropertySingleValue')
    counter = 0
    report_rows = []

    for prop in single_value_props:
        val = prop.NominalValue
        if val:
            try:
                if isinstance(val.wrappedValue, str):
                    original = val.wrappedValue
                    cleaned = val.wrappedValue.rstrip()  # Удаляет пробелы, табы, \n, \r с конца строки
                    if cleaned != val.wrappedValue:
                        val.wrappedValue = cleaned
                        counter += 1
                        # Добавляем в CSV отчет
                        name = prop.Name if prop.Name else "(Без имени)"
                        report_rows.append([name, original, cleaned])
            except Exception as e:
                print(f"Ошибка при обработке {prop}: {e}")

    f.write(out_filepath)

    file_name = os.path.basename(inp_filepath)
    name_without_ext = os.path.splitext(file_name)[0]
    report_filepath = build_report_path(out_filepath, custom_name=name_without_ext)

    with open(report_filepath, 'w', encoding='utf-8-sig', newline='') as csvfile:
        writer = csv.writer(csvfile, delimiter=';')
        writer.writerow(['Property Name', 'Old Value', 'New Value'])  # Заголовки
        writer.writerows(report_rows)

    return f'Удалены лишние символы в {counter} значениях атрибутов.', report_filepath


if __name__ == '__main__':
    ifc_cleaner(r'C:\Python\Projects\AtomskillsCleaner\МодельАС.ifc',
                r'C:\Python\Projects\AtomskillsCleaner\МодельАС-fixed.ifc')
