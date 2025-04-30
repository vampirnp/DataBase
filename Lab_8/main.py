import tkinter as tk
from tkinter import ttk, messagebox
import pyodbc


class EditRecordDialog:
    def __init__(self, parent, table_name, db_path, record_values, columns):
        self.dialog = tk.Toplevel(parent)
        self.table_name = table_name
        self.db_path = db_path
        self.record_values = record_values
        self.columns = columns
        self.result = False

        self.dialog.title("Редактирование записи")
        self.dialog.geometry("400x300")

        self.create_widgets()

    def create_widgets(self):
        try:
            # Создаем фрейм с прокруткой
            canvas = tk.Canvas(self.dialog)
            scrollbar = ttk.Scrollbar(self.dialog, orient="vertical", command=canvas.yview)
            scrollable_frame = ttk.Frame(canvas)

            scrollable_frame.bind(
                "<Configure>",
                lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
            )

            canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)

            self.fields = []
            for i, (column, value) in enumerate(zip(self.columns, self.record_values)):
                ttk.Label(scrollable_frame, text=column).grid(row=i, column=0, padx=5, pady=5)
                entry = ttk.Entry(scrollable_frame)
                entry.insert(0, str(value) if value is not None else '')
                entry.grid(row=i, column=1, padx=5, pady=5)
                self.fields.append((column, entry))

            # Кнопки
            btn_frame = ttk.Frame(self.dialog)
            ttk.Button(btn_frame, text="Сохранить", command=self.save).pack(side='left', padx=5)
            ttk.Button(btn_frame, text="Отмена", command=self.dialog.destroy).pack(side='left', padx=5)

            # Размещаем элементы
            canvas.pack(side="left", fill="both", expand=True, padx=5, pady=5)
            scrollbar.pack(side="right", fill="y")
            btn_frame.pack(side='bottom', pady=10)

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при создании формы:\n{str(e)}")
            self.dialog.destroy()

    def save(self):
        try:
            values = []
            set_clauses = []
            primary_key_name = self.columns[0]
            primary_key_value = self.record_values[0]

            for (field_name, entry), old_value in zip(self.fields[1:], self.record_values[1:]):
                new_value = entry.get().strip()
                if new_value != str(old_value):
                    values.append(new_value)
                    set_clauses.append(f"[{field_name}] = ?")

            if not set_clauses:
                messagebox.showinfo("Информация", "Данные не были изменены")
                self.dialog.destroy()
                return

            conn_str = f'DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={self.db_path};'
            conn = pyodbc.connect(conn_str)
            cursor = conn.cursor()

            # Формируем SQL запрос для обновления
            query = f"UPDATE [{self.table_name}] SET {', '.join(set_clauses)} WHERE [{primary_key_name}] = ?"
            values.append(primary_key_value)

            cursor.execute(query, values)
            conn.commit()
            conn.close()

            self.result = True
            self.dialog.destroy()
            messagebox.showinfo("Успех", "Запись обновлена")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при сохранении:\n{str(e)}")

class MedicalDBApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Медицинская База Данных")
        self.root.geometry("1000x600")

        self.db_path = r'C:\Users\Dmitriy\Desktop\database.accdb'
        self.current_table = "Лечебное учреждение"  # Обратите внимание на пробел

        try:
            self.create_widgets()
            self.refresh_data()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка инициализации:\n{str(e)}")

    def get_connection(self):
        conn_str = (
            r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
            f'DBQ={self.db_path};'
        )
        return pyodbc.connect(conn_str)

    def create_widgets(self):
        # Верхняя панель
        frame_top = ttk.Frame(self.root)
        frame_top.pack(fill='x', padx=5, pady=5)

        ttk.Label(frame_top, text="Таблица:").pack(side='left')
        self.table_var = tk.StringVar(value=self.current_table)
        tables = ["Лечебное учреждение", "Врач", "Пациент", "Прием", "Процедура", "Диагноз"]
        self.table_combo = ttk.Combobox(frame_top, textvariable=self.table_var, values=tables)
        self.table_combo.pack(side='left', padx=5)
        self.table_combo.bind('<<ComboboxSelected>>', lambda e: self.on_table_change())

        # Таблица с прокруткой
        frame_table = ttk.Frame(self.root)
        frame_table.pack(fill='both', expand=True, padx=5)

        self.tree = ttk.Treeview(frame_table)
        scrollbar_y = ttk.Scrollbar(frame_table, orient='vertical', command=self.tree.yview)
        scrollbar_x = ttk.Scrollbar(frame_table, orient='horizontal', command=self.tree.xview)

        self.tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

        scrollbar_y.pack(side='right', fill='y')
        scrollbar_x.pack(side='bottom', fill='x')
        self.tree.pack(side='left', fill='both', expand=True)

        # Кнопки управления
        frame_bottom = ttk.Frame(self.root)
        frame_bottom.pack(fill='x', padx=5, pady=5)

        ttk.Button(frame_bottom, text="Добавить", command=self.add_record).pack(side='left', padx=2)
        ttk.Button(frame_bottom, text="Редактировать", command=self.edit_record).pack(side='left', padx=2)
        ttk.Button(frame_bottom, text="Удалить", command=self.delete_record).pack(side='left', padx=2)
        ttk.Button(frame_bottom, text="Обновить", command=self.refresh_data).pack(side='left', padx=2)

    def on_table_change(self):
        self.current_table = self.table_var.get()
        self.refresh_data()

    def refresh_data(self):
        try:
            for item in self.tree.get_children():
                self.tree.delete(item)

            conn = self.get_connection()
            cursor = conn.cursor()

            cursor.execute(f'SELECT * FROM [{self.current_table}]')

            # Настройка колонок
            columns = [column[0] for column in cursor.description]
            self.tree['columns'] = columns
            self.tree['show'] = 'headings'

            for col in columns:
                self.tree.heading(col, text=col)
                self.tree.column(col, width=100)

            # Заполнение данными
            for row in cursor.fetchall():
                values = [str(value) if value is not None else '' for value in row]
                self.tree.insert('', 'end', values=values)

            conn.close()

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при загрузке данных:\n{str(e)}")

    def add_record(self):
        dialog = AddRecordDialog(self.root, self.current_table, self.db_path)
        if dialog.result:
            self.refresh_data()

    def edit_record(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("Предупреждение", "Выберите запись для редактирования")
            return

        values = self.tree.item(selected_item)['values']
        columns = self.tree['columns']

        dialog = EditRecordDialog(self.root, self.current_table, self.db_path, values, columns)
        if dialog.result:
            self.refresh_data()

    def delete_record(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("Предупреждение", "Выберите запись для удаления")
            return

        if messagebox.askyesno("Подтверждение", "Удалить выбранную запись?"):
            try:
                conn = self.get_connection()
                cursor = conn.cursor()

                values = self.tree.item(selected_item)['values']
                primary_key = values[0]  # Предполагаем, что первый столбец - первичный ключ

                cursor.execute(f"DELETE FROM [{self.current_table}] WHERE [{self.tree['columns'][0]}] = ?",
                               (primary_key,))
                conn.commit()
                conn.close()

                self.refresh_data()
                messagebox.showinfo("Успех", "Запись удалена")

            except Exception as e:
                messagebox.showerror("Ошибка", f"Ошибка при удалении:\n{str(e)}")

class AddRecordDialog:
    def __init__(self, parent, table_name, db_path):
        self.dialog = tk.Toplevel(parent)
        self.table_name = table_name
        self.db_path = db_path
        self.result = False

        self.dialog.title("Добавление записи")
        self.dialog.geometry("400x300")

        self.create_widgets()

    def create_widgets(self):
        try:
            conn_str = f'DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={self.db_path};'
            conn = pyodbc.connect(conn_str)
            cursor = conn.cursor()
            cursor.execute(f"SELECT * FROM [{self.table_name}]")

            # Создаем фрейм с прокруткой
            canvas = tk.Canvas(self.dialog)
            scrollbar = ttk.Scrollbar(self.dialog, orient="vertical", command=canvas.yview)
            scrollable_frame = ttk.Frame(canvas)

            scrollable_frame.bind(
                "<Configure>",
                lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
            )

            canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)

            self.fields = []
            for i, column in enumerate(cursor.description):
                ttk.Label(scrollable_frame, text=column[0]).grid(row=i, column=0, padx=5, pady=5)
                entry = ttk.Entry(scrollable_frame)
                entry.grid(row=i, column=1, padx=5, pady=5)
                self.fields.append((column[0], entry))

            # Кнопки
            btn_frame = ttk.Frame(self.dialog)
            ttk.Button(btn_frame, text="Сохранить", command=self.save).pack(side='left', padx=5)
            ttk.Button(btn_frame, text="Отмена", command=self.dialog.destroy).pack(side='left', padx=5)

            # Размещаем элементы
            canvas.pack(side="left", fill="both", expand=True, padx=5, pady=5)
            scrollbar.pack(side="right", fill="y")
            btn_frame.pack(side='bottom', pady=10)

            conn.close()

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при создании формы:\n{str(e)}")
            self.dialog.destroy()

    def save(self):
        try:
            values = []
            columns = []
            for field_name, entry in self.fields:
                value = entry.get().strip()
                if value:  # Добавляем только непустые значения
                    values.append(value)
                    columns.append(field_name)

            if not values:
                messagebox.showwarning("Предупреждение", "Заполните хотя бы одно поле")
                return

            conn_str = f'DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={self.db_path};'
            conn = pyodbc.connect(conn_str)
            cursor = conn.cursor()

            # Формируем SQL запрос
            columns_str = ','.join([f'[{col}]' for col in columns])
            placeholders = ','.join(['?' for _ in values])
            query = f"INSERT INTO [{self.table_name}] ({columns_str}) VALUES ({placeholders})"

            cursor.execute(query, values)
            conn.commit()
            conn.close()

            self.result = True
            self.dialog.destroy()
            messagebox.showinfo("Успех", "Запись добавлена")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при сохранении:\n{str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = MedicalDBApp(root)
    root.mainloop()