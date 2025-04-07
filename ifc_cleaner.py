import ifcopenshell


def ifc_cleaner(inp_filepath, out_filepath):
    f = ifcopenshell.open(inp_filepath)
    single_value_props = f.by_type('IfcPropertySingleValue')
    counter = 0
    for prop in single_value_props:
        val = prop.NominalValue
        if val:
            try:
                if isinstance(val.wrappedValue, str):
                    cleaned = val.wrappedValue.rstrip()  # Удаляет пробелы, табы, \n, \r с конца строки
                    if cleaned != val.wrappedValue:
                        val.wrappedValue = cleaned
                        counter += 1
            except Exception as e:
                print(f"Ошибка при обработке {prop}: {e}")

    f.write(out_filepath)
    return f'Удалены лишние символы в {counter} значениях атрибутов'


if __name__ == '__main__':
    ifc_cleaner(r'C:\Python\Projects\AtomskillsCleaner\МодельТГВ.ifc',
                r'C:\Python\Projects\AtomskillsCleaner\МодельТГВ-fixed.ifc')
