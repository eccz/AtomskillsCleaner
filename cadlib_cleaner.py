import pyodbc


def cadlib_cleaner(crsr):
    counter = 0
    crsr.execute("SELECT idParamDef, idObject, Value, Comment FROM dbo.Parameters_STR")
    rows = crsr.fetchall()

    for id1, id2, value, comment in rows:
        if value and isinstance(value, str):
            try:
                cleaned = value.rstrip()
                if cleaned != value:
                    crsr.execute(
                        """
                        UPDATE dbo.Parameters_STR
                        SET Value = ?
                        WHERE idParamDef = ? AND idObject = ? AND Value = ?
                        """,
                        (cleaned, id1, id2, value)
                    )
                    counter += 1
            except Exception as e:
                print(f"Ошибка при обработке значения '{value}': {e}")

    return f'Очищено {counter} значений в базе данных'


if __name__ == '__main__':
    conn = pyodbc.connect(f'DRIVER={{SQL Server}};SERVER=(local)\\SQLEXPRESS;DATABASE=AS25_7;UID=sa;PWD=123')
    cursor = conn.cursor()
    cadlib_cleaner(cursor)
    cursor.commit()
    cursor.close()
    conn.close()
