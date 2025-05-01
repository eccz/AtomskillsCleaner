import tkinter as tk
from tkinter import messagebox, filedialog, ttk
from cadlib_cleaner import cadlib_cleaner
from excel_cleaner import excel_cleaner
from ifc_cleaner import ifc_cleaner
from datetime import datetime
import os
import tempfile
import shutil
import pyodbc
import threading
import csv


def handle_file_clean(filetype, file_extension):
    input_path = filedialog.askopenfilename(
        title=f"Выберите {filetype} файл для очистки",
        filetypes=[(f"{filetype} files", f"*.{file_extension}")]
    )
    if not input_path:
        return

    # Создаём временную папку
    temp_dir = tempfile.mkdtemp()
    temp_output_path = os.path.join(temp_dir, f"temp_cleaned.{file_extension}")

    # Создаём окно ожидания с прогресс-баром
    wait_win = tk.Toplevel()
    wait_win.title("Обработка файла...")
    wait_win.geometry("300x100")
    wait_win.resizable(False, False)
    wait_win.grab_set()  # делаем модальным
    tk.Label(wait_win, text=f"Идёт обработка {filetype} файла...\nПожалуйста, подождите.").pack(pady=10)

    progress = ttk.Progressbar(wait_win, mode="indeterminate", length=250)
    progress.pack(pady=5)
    progress.start(10)  # скорость анимации

    # Функция для выполнения в отдельном потоке
    def process_file():
        try:
            if filetype == 'IFC':
                res, report_path = ifc_cleaner(input_path, temp_output_path)
            elif filetype == 'Excel':
                res = excel_cleaner(input_path, temp_output_path)
                report_path = ''

            # Сохраняем файл
            def after_success():
                wait_win.destroy()
                output_path = filedialog.asksaveasfilename(
                    title="Сохранить очищенный файл как...",
                    defaultextension=f".{file_extension}",
                    initialfile=os.path.basename(input_path).replace(f".{file_extension}",
                                                                     f"_cleaned.{file_extension}"),
                    filetypes=[(f"{filetype} files", f"*.{file_extension}")]
                )
                if output_path:
                    try:
                        shutil.copy(temp_output_path, output_path)
                        messagebox.showinfo("Успех", f"{res}\n\nФайл сохранён:\n{output_path}")
                    except PermissionError:
                        messagebox.showerror("Ошибка", "Недостаточно прав для записи файла.\n")
                    except Exception as ex:
                        messagebox.showerror("Ошибка", f"Не удалось сохранить файл:\n{ex}")
                else:
                    messagebox.showinfo("Отмена", "Сохранение отменено.")

                if report_path and os.path.exists(report_path):
                    save_report = messagebox.askyesno("Сохранить отчет?", "Хотите сохранить CSV-отчет об изменениях?")
                    if save_report:
                        report_output = filedialog.asksaveasfilename(
                            title="Сохранить отчет как...",
                            defaultextension=".csv",
                            initialfile=os.path.basename(report_path),
                            filetypes=[("CSV файлы", "*.csv")]
                        )
                        if report_output:
                            try:
                                shutil.copy(report_path, report_output)
                            except Exception as ex:
                                messagebox.showerror("Ошибка", f"Не удалось сохранить отчет:\n{ex}")

                shutil.rmtree(temp_dir)

            wait_win.after(0, after_success)

        except Exception as e:
            wait_win.after(0, lambda: (
                wait_win.destroy(),
                messagebox.showerror("Ошибка", f"Ошибка при очистке:\n{e}"),
                shutil.rmtree(temp_dir)
            ))

    # Запуск в отдельном потоке
    threading.Thread(target=process_file).start()


def handle_cadlib_clean():
    cadlib_window = tk.Toplevel(root)

    cadlib_window.title("Подключение к базе Cadlib")
    cadlib_window.geometry("250x450")

    tk.Label(cadlib_window, text="Сервер:").pack(pady=5)
    server_entry = tk.Entry(cadlib_window)
    server_entry.pack(pady=5)
    server_entry.insert(0, "(local)\\SQLEXPRESS")

    tk.Label(cadlib_window, text="Аутентификация:").pack(pady=5)
    auth_var = tk.StringVar(value="SQL")
    radio_sql = tk.Radiobutton(cadlib_window, text="SQL Server", variable=auth_var, value="SQL")
    radio_sql.pack()
    radio_windows = tk.Radiobutton(cadlib_window, text="Windows", variable=auth_var, value="Windows")
    radio_windows.pack()

    tk.Label(cadlib_window, text="Пользователь:").pack(pady=5)
    user_entry = tk.Entry(cadlib_window)
    user_entry.pack(pady=5)
    user_entry.insert(0, "sa")

    tk.Label(cadlib_window, text="Пароль:").pack(pady=5)
    password_entry = tk.Entry(cadlib_window, show="*")
    password_entry.pack(pady=5)
    password_entry.insert(0, "123")

    def connect_to_db():
        server = server_entry.get()
        database = db_combobox.get()
        username = user_entry.get()
        password = password_entry.get()
        auth_mode = auth_var.get()

        try:
            if auth_mode == "Windows":
                return pyodbc.connect(
                    f'DRIVER={{SQL Server}};SERVER={server};DATABASE={database};Trusted_Connection=yes')
            else:
                return pyodbc.connect(
                    f'DRIVER={{SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}')

        except Exception as err:
            messagebox.showerror("Ошибка", f"Ошибка подключения к sql-server: {err}")

    def fetch_databases():
        conn = connect_to_db()

        try:
            cursor = conn.cursor()
            cursor.execute("SELECT name FROM sys.databases")
            databases = [row[0] for row in cursor.fetchall()]
            db_combobox['values'] = databases
            # print(db_combobox['values'])
            if databases:
                db_combobox.current(0)

            conn.close()
        except Exception as err:
            messagebox.showerror("Ошибка", f"Ошибка получения баз данных: {err}")

    def clean_cadlib_db():
        conn = connect_to_db()
        report_rows = []
        try:
            cursor = conn.cursor()
            res, report_rows = cadlib_cleaner(cursor)
            cursor.commit()
            cursor.close()
            conn.close()
            cadlib_window.destroy()
            messagebox.showinfo("Успех", f"{res}\n\nЛишние символы в атрибутах базы данных Cadlib очищены!")

        except Exception as err:
            messagebox.showerror("Ошибка", f"Ошибка получения баз данных: {err}")

        try:
            timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            report_path = filedialog.asksaveasfilename(
                title="Сохранить отчет",
                defaultextension=".csv",
                initialfile=f"cadlib_clean_report_{timestamp}.csv",
                filetypes=[("CSV файлы", "*.csv")]
            )

            if report_path and report_rows:
                with open(report_path, 'w', encoding='utf-8-sig', newline='') as csvfile:
                    writer = csv.writer(csvfile, delimiter=';')
                    writer.writerow(['Parameter', 'Object', 'Old Value', 'New Value'])  # Заголовки
                    writer.writerows(report_rows)

        except Exception as err:
            messagebox.showerror("Ошибка", f"Ошибка создания отчета: {err}")

    # Кнопка для загрузки списка баз данных
    fetch_db_button = tk.Button(cadlib_window, text="Получить базы", command=fetch_databases)
    fetch_db_button.pack(pady=5)

    tk.Label(cadlib_window, text="База данных:").pack(pady=5)
    db_combobox = ttk.Combobox(cadlib_window, state="readonly")
    db_combobox.pack(pady=5)

    # Кнопка очистки параметров CadLib
    clear_button = tk.Button(cadlib_window, text="Очистить значения атрибутов", command=clean_cadlib_db)
    clear_button.pack(pady=20)


# Создание главного окна
root = tk.Tk()
root.title("Очистка атрибутов AtomSkills ЯОК")
root.geometry("360x150")  # Размер окна

# Кнопка для очистки файла Excel
btn_excel = tk.Button(root, text="Очистка файла Excel", command=lambda: handle_file_clean('Excel', 'xlsx'))
btn_excel.pack(fill=tk.BOTH, expand=True, pady=5)

# Кнопка для очистки файла IFC
btn_ifc = tk.Button(root, text="Очистка файла IFC", command=lambda: handle_file_clean('IFC', 'ifc'))
btn_ifc.pack(fill=tk.BOTH, expand=True, pady=5)

# Кнопка для очистки базы Cadlib
btn_cadlib = tk.Button(root, text="Очистка базы Cadlib", command=handle_cadlib_clean)
btn_cadlib.pack(fill=tk.BOTH, expand=True, pady=5)

# Запуск главного цикла приложения
root.mainloop()
