import pyodbc


def cadlib_cleaner(crsr):
    counter = 0
    crsr.execute("SELECT idParamDef, idObject, Value, Comment FROM dbo.Parameters_STR")
    rows = crsr.fetchall()
    report_rows = []

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
                    param = crsr.execute(f"SELECT Name FROM dbo.ParamDefs WHERE idParamDef = {id1}").fetchone()[0]
                    obj = crsr.execute(f"SELECT Name FROM dbo.ObjectsShadow WHERE idObject = {id2}").fetchone()[0]
                    report_rows.append([param, obj, value, cleaned])
                    counter += 1

            except Exception as e:
                print(f"Ошибка при обработке значения '{value}': {e}")

    return f'Очищено {counter} значений в базе данных', report_rows


if __name__ == '__main__':
    conn = pyodbc.connect(f'DRIVER={{SQL Server}};SERVER=(local)\\SQLEXPRESS;DATABASE=AS25_7;UID=sa;PWD=123')
    cursor = conn.cursor()
    cadlib_cleaner(cursor)
    cursor.commit()
    cursor.close()
    conn.close()
